use std::collections::HashMap;
use std::fs;
use std::path::PathBuf;
use std::sync::{Arc, Mutex, OnceLock};
use std::thread;
use std::time::Instant;

use serde::{Deserialize, Serialize};
use windows::Win32::Foundation::{
    CloseHandle, HGLOBAL, HANDLE, HWND, LPARAM, LRESULT, POINT, RECT, WPARAM,
};
use windows::Win32::Graphics::Gdi::InvalidateRect;
use windows::Win32::Storage::FileSystem::ReadFile;
use windows::Win32::System::DataExchange::{CloseClipboard, GetClipboardData, OpenClipboard};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Memory::{GlobalLock, GlobalSize, GlobalUnlock};
use windows::Win32::System::Ole::CF_UNICODETEXT;
use windows::Win32::System::Threading::{WaitForSingleObject, INFINITE};
use windows::Win32::UI::Input::Ime::{
    CFS_POINT, CANDIDATEFORM, COMPOSITIONFORM,
    ImmGetContext, ImmReleaseContext, ImmSetCandidateWindow, ImmSetCompositionWindow,
};
use windows::Win32::UI::Input::KeyboardAndMouse::*;
use windows::Win32::UI::WindowsAndMessaging::*;

use crate::app::{App, Selection, SelectionMode};
use crate::input::{build_key_sequence, get_modifiers};
use crate::pty::{ShellType, get_wsl_distros};
use crate::render::TabBarHitResult;
use crate::tab::TabId;
use crate::tsf::TsfContext;

const CLASS_NAME: &str = "TaaminaluWindow";
const WINDOW_TITLE: &str = "taaminalu";
const DEFAULT_WIDTH: i32 = 800;
const DEFAULT_HEIGHT: i32 = 600;

#[derive(Serialize, Deserialize)]
struct WindowGeometry {
    x: i32,
    y: i32,
    width: i32,
    height: i32,
}

fn config_path() -> PathBuf {
    let appdata = std::env::var("APPDATA").unwrap_or_else(|_| ".".to_string());
    PathBuf::from(appdata).join("taaminalu").join("window.json")
}

fn load_geometry() -> Option<WindowGeometry> {
    let path = config_path();
    let data = fs::read_to_string(path).ok()?;
    serde_json::from_str(&data).ok()
}

fn save_geometry(hwnd: HWND) {
    let mut rect = RECT::default();
    if unsafe { GetWindowRect(hwnd, &mut rect) }.is_err() {
        return;
    }
    let geo = WindowGeometry {
        x: rect.left,
        y: rect.top,
        width: rect.right - rect.left,
        height: rect.bottom - rect.top,
    };
    let path = config_path();
    if let Some(parent) = path.parent() {
        let _ = fs::create_dir_all(parent);
    }
    if let Ok(json) = serde_json::to_string(&geo) {
        let _ = fs::write(path, json);
    }
}

/// カスタムメッセージ: PTY からデータ受信で再描画要求 (WPARAM = TabId)
pub const WM_PTY_OUTPUT: u32 = WM_USER + 1;
/// カスタムメッセージ: タブの PTY プロセスが終了 (WPARAM = TabId)
pub const WM_TAB_CLOSED: u32 = WM_USER + 2;

// シェル選択メニュー ID
const MENU_ID_WSL_DEFAULT: u32 = 1;
const MENU_ID_CMD: u32 = 2;
const MENU_ID_POWERSHELL: u32 = 3;
// WSL ディストリ用 ID: 100 + index
const MENU_ID_WSL_DISTRO_BASE: u32 = 100;

/// HWND ごとの TSF コンテキスト
static TSF_CONTEXTS: OnceLock<Mutex<HashMap<isize, TsfContext>>> = OnceLock::new();

pub fn set_tsf_context(hwnd: HWND, ctx: Option<TsfContext>) {
    let map = TSF_CONTEXTS.get_or_init(|| Mutex::new(HashMap::new()));
    let mut map = map.lock().unwrap();
    if let Some(ctx) = ctx {
        map.insert(hwnd.0 as isize, ctx);
    }
}

/// TSF コンテキストにアクセスしてクロージャを実行
fn with_tsf<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&TsfContext) -> R,
{
    let map = TSF_CONTEXTS.get()?;
    let map = map.lock().unwrap();
    map.get(&(hwnd.0 as isize)).map(f)
}

fn notify_tsf_change(hwnd: HWND) {
    with_tsf(hwnd, |ctx| ctx.notify_change());
}

fn is_composing(hwnd: HWND) -> bool {
    with_tsf(hwnd, |ctx| ctx.is_composing()).unwrap_or(false)
}

fn get_preedit(hwnd: HWND) -> String {
    with_tsf(hwnd, |ctx| ctx.preedit()).unwrap_or_default()
}

/// 再描画要求 + TSF テキスト変更通知
fn repaint(hwnd: HWND) {
    unsafe { let _ = InvalidateRect(Some(hwnd), None, false); }
    notify_tsf_change(hwnd);
}

/// IMM32 API で候補ウィンドウ・変換ウィンドウの位置をカーソル位置に設定
fn update_ime_position(hwnd: HWND) {
    let app = get_app(hwnd);
    let Some(app) = app else { return };
    let app = app.lock().unwrap();
    let (cell_w, cell_h) = app.cell_size();
    let (_, grid_y) = app.grid_origin();
    let (cursor_row, cursor_col) = app.active().term.cursor_pos();
    let x = (cursor_col as f32 * cell_w) as i32;
    let y = (cursor_row as f32 * cell_h) as i32 + grid_y as i32;
    drop(app);

    let pt = POINT { x, y };
    unsafe {
        let himc = ImmGetContext(hwnd);
        if !himc.0.is_null() {
            let _ = ImmSetCompositionWindow(
                himc,
                &COMPOSITIONFORM {
                    dwStyle: CFS_POINT,
                    ptCurrentPos: pt,
                    ..Default::default()
                },
            );
            let _ = ImmSetCandidateWindow(
                himc,
                &CANDIDATEFORM {
                    dwIndex: 0,
                    dwStyle: CFS_POINT,
                    ptCurrentPos: POINT { x: pt.x, y: pt.y + cell_h as i32 },
                    ..Default::default()
                },
            );
            let _ = ImmReleaseContext(hwnd, himc);
        }
    }
}

pub fn create_window(app: Arc<Mutex<App>>) -> windows::core::Result<HWND> {
    let class_name_wide: Vec<u16> = CLASS_NAME.encode_utf16().chain(std::iter::once(0)).collect();
    let title_wide: Vec<u16> = WINDOW_TITLE.encode_utf16().chain(std::iter::once(0)).collect();

    unsafe {
        let hinstance = GetModuleHandleW(None)?;

        let hicon = LoadIconW(
            Some(windows::Win32::Foundation::HINSTANCE(hinstance.0)),
            windows::core::PCWSTR(1 as *const u16), // MAKEINTRESOURCE(1)
        )
        .unwrap_or_default();

        let wc = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(wnd_proc),
            hInstance: hinstance.into(),
            lpszClassName: windows::core::PCWSTR(class_name_wide.as_ptr()),
            hCursor: LoadCursorW(None, IDC_ARROW)?,
            hIcon: hicon,
            hIconSm: hicon,
            ..Default::default()
        };

        RegisterClassExW(&wc);

        let app_ptr = Box::into_raw(Box::new(app));

        let (x, y, w, h) = match load_geometry() {
            Some(geo) => (geo.x, geo.y, geo.width, geo.height),
            None => (CW_USEDEFAULT, CW_USEDEFAULT, DEFAULT_WIDTH, DEFAULT_HEIGHT),
        };

        let hwnd = CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            windows::core::PCWSTR(class_name_wide.as_ptr()),
            windows::core::PCWSTR(title_wide.as_ptr()),
            WS_OVERLAPPEDWINDOW | WS_VISIBLE,
            x,
            y,
            w,
            h,
            None,
            None,
            Some(hinstance.into()),
            Some(app_ptr as *const _),
        )?;

        Ok(hwnd)
    }
}

pub fn run_message_loop() {
    let mut msg = MSG::default();
    unsafe {
        while GetMessageW(&mut msg, None, 0, 0).as_bool() {
            let _ = TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
}

/// PTY 読み取りスレッドを起動
pub(crate) fn start_pty_reader(
    app: Arc<Mutex<App>>,
    hwnd: HWND,
    read_handle: HANDLE,
    tab_id: TabId,
) {
    let hwnd_val = hwnd.0 as usize;
    let handle_val = read_handle.0 as usize;

    thread::spawn(move || {
        let hwnd = HWND(hwnd_val as *mut _);
        let read_handle = HANDLE(handle_val as *mut _);
        let mut buf = [0u8; 4096];
        loop {
            let mut bytes_read = 0u32;
            let ok = unsafe {
                ReadFile(read_handle, Some(&mut buf), Some(&mut bytes_read), None)
            };
            match ok {
                Ok(()) if bytes_read == 0 => break,
                Ok(()) => {
                    let n = bytes_read as usize;
                    let mut app_lock = app.lock().unwrap();
                    app_lock.process_pty_output_for_tab(tab_id, &buf[..n]);
                    drop(app_lock);
                    unsafe {
                        let _ = PostMessageW(
                            Some(hwnd),
                            WM_PTY_OUTPUT,
                            WPARAM(tab_id as usize),
                            LPARAM(0),
                        );
                    }
                }
                Err(_) => break,
            }
        }
        unsafe {
            let _ = CloseHandle(read_handle);
            // PTY が終了したことを通知
            let _ = PostMessageW(
                Some(hwnd),
                WM_TAB_CLOSED,
                WPARAM(tab_id as usize),
                LPARAM(0),
            );
        }
    });
}

/// プロセス終了を監視するスレッドを起動
/// プロセスが終了したら WM_TAB_CLOSED を送信し、Tab の Drop で ConPTY が閉じられ
/// ReadFile がエラーになってリーダースレッドも終了する
pub(crate) fn start_process_watcher(
    hwnd: HWND,
    process_handle: HANDLE,
    tab_id: TabId,
) {
    let hwnd_val = hwnd.0 as usize;
    let handle_val = process_handle.0 as usize;

    thread::spawn(move || {
        let hwnd = HWND(hwnd_val as *mut _);
        let process_handle = HANDLE(handle_val as *mut _);
        unsafe {
            WaitForSingleObject(process_handle, INFINITE);
            let _ = CloseHandle(process_handle);
            let _ = PostMessageW(
                Some(hwnd),
                WM_TAB_CLOSED,
                WPARAM(tab_id as usize),
                LPARAM(0),
            );
        }
    });
}

unsafe extern "system" fn wnd_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let cs = unsafe { &*(lparam.0 as *const CREATESTRUCTW) };
            let app_ptr = cs.lpCreateParams as *mut Arc<Mutex<App>>;
            unsafe {
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, app_ptr as isize);
            }
            LRESULT(0)
        }
        WM_SETFOCUS => {
            // TSF ドキュメントマネージャにフォーカスを復元
            // AssociateFocus で自動的に処理されるが、明示的にも呼び出す
            with_tsf(hwnd, |ctx| {
                unsafe { let _ = ctx.thread_mgr.SetFocus(&ctx.doc_mgr); }
            });
            LRESULT(0)
        }
        WM_IME_STARTCOMPOSITION => {
            update_ime_position(hwnd);
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
        WM_IME_COMPOSITION => {
            update_ime_position(hwnd);
            unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
        }
        WM_PAINT => {
            let preedit = get_preedit(hwnd);
            let app = get_app(hwnd);
            if let Some(app) = app {
                let app = app.lock().unwrap();
                app.paint(hwnd, &preedit);
            }
            LRESULT(0)
        }
        WM_SIZE => {
            let width = (lparam.0 & 0xFFFF) as u32;
            let height = ((lparam.0 >> 16) & 0xFFFF) as u32;
            if width > 0 && height > 0 {
                let app = get_app(hwnd);
                if let Some(app) = app {
                    let mut app = app.lock().unwrap();
                    app.on_resize(width, height);
                }
            }
            LRESULT(0)
        }
        WM_CHAR | WM_SYSCHAR => {
            // IME composition 中は WM_CHAR を無視（TSF 経由で処理される）
            if is_composing(hwnd) {
                return LRESULT(0);
            }
            let ch = wparam.0 as u32;
            if let Some(c) = char::from_u32(ch) {
                // VK_BACK は WM_KEYDOWN で処理済み（0x08=BS, 0x7F=DEL 両方をスキップ）
                if c == '\x08' || c == '\x7f' {
                    return LRESULT(0);
                }
                let app = get_app(hwnd);
                if let Some(app) = app {
                    let mut app = app.lock().unwrap();
                    app.scroll_to_bottom();
                    // Alt が押されていたら ESC プレフィックス付き
                    let alt = (lparam.0 >> 29) & 1 != 0; // bit 29 = context code (Alt)
                    if alt && c.is_ascii() {
                        let mut buf = vec![0x1bu8]; // ESC
                        let mut char_buf = [0u8; 4];
                        let s = c.encode_utf8(&mut char_buf);
                        buf.extend_from_slice(s.as_bytes());
                        let _ = app.write_pty(&buf);
                    } else {
                        let mut buf = [0u8; 4];
                        let s = c.encode_utf8(&mut buf);
                        let _ = app.write_pty(s.as_bytes());
                    }
                }
            }
            LRESULT(0)
        }
        WM_KEYDOWN | WM_SYSKEYDOWN => {
            let vk = VIRTUAL_KEY(wparam.0 as u16);
            let mods = get_modifiers();

            // IME composition 中はキー入力を IME に委ねる（タブ操作以外）
            if is_composing(hwnd) && !mods.ctrl {
                return unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) };
            }

            // タブ操作ショートカット
            if mods.ctrl && mods.shift {
                match vk {
                    VK_T => {
                        show_new_tab_menu(hwnd);
                        return LRESULT(0);
                    }
                    VK_W => {
                        close_active_tab(hwnd);
                        return LRESULT(0);
                    }
                    VK_TAB => {
                        // Ctrl+Shift+Tab: 前のタブ
                        switch_tab_relative(hwnd, -1);
                        return LRESULT(0);
                    }
                    _ => {}
                }
            }

            // Ctrl+Tab: 次のタブ
            if mods.ctrl && !mods.shift && vk == VK_TAB {
                switch_tab_relative(hwnd, 1);
                return LRESULT(0);
            }

            // 特殊キーのエスケープシーケンス
            if let Some(seq) = build_key_sequence(vk, &mods) {
                let app = get_app(hwnd);
                if let Some(app) = app {
                    let mut app = app.lock().unwrap();
                    app.scroll_to_bottom();
                    let _ = app.write_pty(&seq);
                }
                return LRESULT(0);
            }

            // Alt+key は WM_SYSCHAR に任せるため DefWindowProc に渡す
            if mods.alt && !mods.ctrl {
                return unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) };
            }

            LRESULT(0)
        }
        WM_LBUTTONDOWN => {
            let px = (lparam.0 & 0xFFFF) as i16;
            let py = ((lparam.0 >> 16) & 0xFFFF) as i16;
            if let Some(app) = get_app(hwnd) {
                let mut app = app.lock().unwrap();
                let (_, grid_y) = app.grid_origin();
                if (py as f32) < grid_y {
                    drop(app);
                    handle_tab_bar_click(hwnd, px as f32, py as f32);
                } else {
                    let had_selection = app.selection.is_some();
                    let grid_pos = app.screen_to_grid(px as f32, py as f32);
                    let current_offset = app.active().term.display_offset();

                    // ダブルクリック判定（500ms 以内・同一セル）
                    let now = Instant::now();
                    let is_double = app.last_click.take().is_some_and(|(time, r, c)| {
                        now.duration_since(time).as_millis() < 500
                            && r == grid_pos.0
                            && c == grid_pos.1
                    });

                    if is_double {
                        // 単語選択
                        let history = app.active().term.history_size();
                        let (sc, ec) = app.active().term.word_boundary(grid_pos.0, grid_pos.1);
                        let stable_row = grid_pos.0 + history - current_offset;
                        let start = (stable_row, sc);
                        let end = (stable_row, ec);
                        app.selection = Some(Selection {
                            start,
                            end,
                            active: true,
                            mode: SelectionMode::Word,
                            origin_word: Some((start, end)),
                        });
                        app.drag_origin = Some((px, py, grid_pos.0, grid_pos.1));
                        // ダブルクリック後の last_click をクリアして
                        // 次のクリックが新規になるようにする
                        app.last_click = None;
                    } else {
                        // 通常クリック
                        app.selection = None;
                        app.drag_origin = Some((px, py, grid_pos.0, grid_pos.1));
                        app.last_click = Some((now, grid_pos.0, grid_pos.1));
                    }

                    drop(app);
                    unsafe { SetCapture(hwnd); }
                    if had_selection || is_double {
                        unsafe { let _ = InvalidateRect(Some(hwnd), None, false); }
                    }
                }
            }
            LRESULT(0)
        }
        WM_MOUSEMOVE => {
            if let Some(app) = get_app(hwnd) {
                let mut app = app.lock().unwrap();
                let px = (lparam.0 & 0xFFFF) as i16;
                let py = ((lparam.0 >> 16) & 0xFFFF) as i16;
                let pos = app.screen_to_grid(px as f32, py as f32);
                let current_offset = app.active().term.display_offset();
                let history = app.active().term.history_size();
                // Word モードのドラッグ用: 借用の競合を避けるため必要時のみ先に計算
                let is_word_drag = app.selection.as_ref()
                    .is_some_and(|s| s.active && s.mode == SelectionMode::Word);
                let word_at_pos = if is_word_drag {
                    app.active().term.word_boundary(pos.0, pos.1)
                } else {
                    (0, 0)
                };
                if app.selection.as_ref().is_some_and(|s| s.active) {
                    let stable_pos_row = pos.0 + history - current_offset;
                    if let Some(ref mut sel) = app.selection {
                        if sel.mode == SelectionMode::Word {
                            // Word モード: 単語単位で選択範囲を拡張
                            let (wsc, wec) = word_at_pos;
                            if let Some((origin_start, origin_end)) = sel.origin_word {
                                if stable_pos_row < origin_start.0
                                    || (stable_pos_row == origin_start.0 && wsc < origin_start.1)
                                {
                                    // 起点より前方にドラッグ
                                    sel.start = (stable_pos_row, wsc);
                                    sel.end = origin_end;
                                } else if stable_pos_row > origin_end.0
                                    || (stable_pos_row == origin_end.0 && wec > origin_end.1)
                                {
                                    // 起点より後方にドラッグ
                                    sel.start = origin_start;
                                    sel.end = (stable_pos_row, wec);
                                } else {
                                    // 起点単語内
                                    sel.start = origin_start;
                                    sel.end = origin_end;
                                }
                            }
                        } else {
                            // 通常モード: セル単位で更新
                            sel.end = (stable_pos_row, pos.1);
                        }
                    }
                    drop(app);
                    unsafe { let _ = InvalidateRect(Some(hwnd), None, false); }
                } else if let Some((ox, oy, gr, gc)) = app.drag_origin {
                    // ピクセル単位で少しでも動いたらドラッグ開始
                    if px != ox || py != oy {
                        let history = app.active().term.history_size();
                        let stable_start = (gr + history - current_offset, gc);
                        let stable_end = (pos.0 + history - current_offset, pos.1);
                        app.selection = Some(Selection {
                            start: stable_start,
                            end: stable_end,
                            active: true,
                            mode: SelectionMode::Normal,
                            origin_word: None,
                        });
                        drop(app);
                        unsafe { let _ = InvalidateRect(Some(hwnd), None, false); }
                    }
                }
            }
            LRESULT(0)
        }
        WM_LBUTTONUP => {
            unsafe { let _ = ReleaseCapture(); }
            if let Some(app) = get_app(hwnd) {
                let mut app = app.lock().unwrap();
                app.drag_origin = None;
                // 選択範囲があればコピーしてハイライトを残す
                let copy_range = app.selection.as_ref().map(|sel| (sel.start, sel.end));
                if let Some((start, end)) = copy_range {
                    let text = app.active().term.selected_text(start, end);
                    if !text.is_empty() {
                        crate::term::set_clipboard_text(&text);
                    }
                    // ハイライトを残す（次のクリックで消える）
                    if let Some(ref mut sel) = app.selection {
                        sel.active = false;
                    }
                } else {
                    app.selection = None;
                }
            }
            LRESULT(0)
        }
        WM_MBUTTONDOWN => {
            paste_from_clipboard(hwnd);
            LRESULT(0)
        }
        WM_MOUSEWHEEL => {
            let delta = (wparam.0 >> 16) as i16;
            // システム設定からノッチあたりのスクロール行数を取得
            let mut lines_per_notch: u32 = 3;
            unsafe {
                let _ = SystemParametersInfoW(
                    SPI_GETWHEELSCROLLLINES,
                    0,
                    Some(&mut lines_per_notch as *mut u32 as *mut _),
                    SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS(0),
                );
            }
            let lines = calc_scroll_lines(delta, lines_per_notch);
            if lines != 0
                && let Some(app) = get_app(hwnd) {
                    let mut app = app.lock().unwrap();
                    let idx = app.active_tab;
                    let tab = &mut app.tabs[idx];
                    if tab.term.is_alt_screen() {
                        let keys = alt_screen_arrow_keys(lines);
                        let _ = tab.write_pty(&keys);
                    } else {
                        tab.term.scroll_display(alacritty_terminal::grid::Scroll::Delta(lines));
                        drop(app);
                        repaint(hwnd);
                    }
            }
            LRESULT(0)
        }
        WM_PTY_OUTPUT => {
            let tab_id = wparam.0 as TabId;
            // アクティブタブの出力なら再描画 + TSF通知
            let is_active = {
                let app = get_app(hwnd);
                app.map(|a| {
                    let app = a.lock().unwrap();
                    app.active_tab_id() == tab_id
                }).unwrap_or(false)
            };
            if is_active {
                repaint(hwnd);
            }
            LRESULT(0)
        }
        WM_TAB_CLOSED => {
            let tab_id = wparam.0 as TabId;
            let app = get_app(hwnd);
            if let Some(app) = app {
                let mut app_lock = app.lock().unwrap();
                if let Some(index) = app_lock.find_tab_index(tab_id) {
                    let should_close = app_lock.close_tab(index);
                    drop(app_lock);
                    if should_close {
                        unsafe { let _ = DestroyWindow(hwnd); }
                    } else {
                        repaint(hwnd);
                    }
                }
            }
            LRESULT(0)
        }
        WM_DESTROY => {
            save_geometry(hwnd);
            unsafe {
                let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut Arc<Mutex<App>>;
                if !ptr.is_null() {
                    drop(Box::from_raw(ptr));
                    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                }
                PostQuitMessage(0);
            }
            LRESULT(0)
        }
        _ => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
    }
}

fn get_app(hwnd: HWND) -> Option<Arc<Mutex<App>>> {
    let ptr = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut Arc<Mutex<App>>;
    if ptr.is_null() {
        return None;
    }
    let app = unsafe { &*ptr };
    Some(Arc::clone(app))
}

// --- タブ操作 ---

fn handle_tab_bar_click(hwnd: HWND, x: f32, y: f32) {
    let app = get_app(hwnd);
    let app = match app {
        Some(a) => a,
        None => return,
    };

    let hit = {
        let app_lock = app.lock().unwrap();
        if let Some(ref renderer) = app_lock.renderer {
            renderer.hit_test_tab_bar(x, y, app_lock.tab_count())
        } else {
            return;
        }
    };

    match hit {
        TabBarHitResult::Tab(index) => {
            let mut app_lock = app.lock().unwrap();
            app_lock.switch_tab(index);
            drop(app_lock);
            repaint(hwnd);
        }
        TabBarHitResult::CloseTab(index) => {
            close_tab_at(hwnd, index);
        }
        TabBarHitResult::NewTab => {
            show_new_tab_menu(hwnd);
        }
        TabBarHitResult::None => {}
    }
}

fn show_new_tab_menu(hwnd: HWND) {
    unsafe {
        let menu = match CreatePopupMenu() {
            Ok(m) => m,
            Err(_) => return,
        };

        // WSL ディストリ一覧を取得してサブメニュー化
        let distros = get_wsl_distros();
        if distros.is_empty() {
            // ディストリ取得できなかった場合はデフォルト WSL 1項目
            let wsl_text: Vec<u16> = "WSL\0".encode_utf16().collect();
            let _ = AppendMenuW(menu, MF_STRING, MENU_ID_WSL_DEFAULT as usize, windows::core::PCWSTR(wsl_text.as_ptr()));
        } else {
            let wsl_sub = match CreatePopupMenu() {
                Ok(m) => m,
                Err(_) => return,
            };
            // デフォルト WSL（ディストリ指定なし）
            let default_text: Vec<u16> = "デフォルト\0".encode_utf16().collect();
            let _ = AppendMenuW(wsl_sub, MF_STRING, MENU_ID_WSL_DEFAULT as usize, windows::core::PCWSTR(default_text.as_ptr()));
            let _ = AppendMenuW(wsl_sub, MF_SEPARATOR, 0, None);
            // 各ディストリ
            for (i, distro) in distros.iter().enumerate() {
                let text: Vec<u16> = format!("{distro}\0").encode_utf16().collect();
                let id = MENU_ID_WSL_DISTRO_BASE.saturating_add(i as u32);
                let _ = AppendMenuW(wsl_sub, MF_STRING, id as usize, windows::core::PCWSTR(text.as_ptr()));
            }
            let wsl_text: Vec<u16> = "WSL\0".encode_utf16().collect();
            let _ = AppendMenuW(menu, MF_POPUP, wsl_sub.0 as usize, windows::core::PCWSTR(wsl_text.as_ptr()));
        }

        let cmd_text: Vec<u16> = "CMD\0".encode_utf16().collect();
        let ps_text: Vec<u16> = "PowerShell\0".encode_utf16().collect();
        let _ = AppendMenuW(menu, MF_STRING, MENU_ID_CMD as usize, windows::core::PCWSTR(cmd_text.as_ptr()));
        let _ = AppendMenuW(menu, MF_STRING, MENU_ID_POWERSHELL as usize, windows::core::PCWSTR(ps_text.as_ptr()));

        let mut pt = POINT { x: 0, y: 0 };
        let _ = GetCursorPos(&mut pt);

        let result = TrackPopupMenu(
            menu,
            TPM_RETURNCMD | TPM_LEFTBUTTON,
            pt.x,
            pt.y,
            Some(0),
            hwnd,
            None,
        );

        let _ = DestroyMenu(menu);

        if result.as_bool() {
            let selected = result.0 as u32;
            let shell = match selected {
                MENU_ID_WSL_DEFAULT => ShellType::Wsl { distro: None },
                MENU_ID_CMD => ShellType::Cmd,
                MENU_ID_POWERSHELL => ShellType::PowerShell,
                id if id >= MENU_ID_WSL_DISTRO_BASE => {
                    let idx = (id - MENU_ID_WSL_DISTRO_BASE) as usize;
                    match distros.get(idx) {
                        Some(name) => ShellType::Wsl { distro: Some(name.clone()) },
                        None => return,
                    }
                }
                _ => return,
            };
            create_new_tab(hwnd, shell);
        }
    }
}

fn create_new_tab(hwnd: HWND, shell: ShellType) {
    let app = match get_app(hwnd) {
        Some(a) => a,
        None => return,
    };

    let (tab_id, read_handle, process_handle) = {
        let mut app_lock = app.lock().unwrap();
        let (cols, rows) = app_lock.grid_size();
        let tab_id = match app_lock.add_tab(shell, cols, rows) {
            Ok(id) => id,
            Err(e) => {
                eprintln!("[taaminalu] Failed to create tab: {}", e);
                return;
            }
        };
        let tab = app_lock.tabs.iter().find(|t| t.id == tab_id).unwrap();
        let read_handle = match tab.dup_output_read() {
            Ok(h) => h,
            Err(e) => {
                eprintln!("[taaminalu] Failed to dup read handle: {}", e);
                return;
            }
        };
        let process_handle = match tab.dup_process_handle() {
            Ok(h) => h,
            Err(e) => {
                eprintln!("[taaminalu] Failed to dup process handle: {}", e);
                return;
            }
        };
        (tab_id, read_handle, process_handle)
    };

    start_pty_reader(Arc::clone(&app), hwnd, read_handle, tab_id);
    start_process_watcher(hwnd, process_handle, tab_id);
    repaint(hwnd);
}

fn close_active_tab(hwnd: HWND) {
    let app = match get_app(hwnd) {
        Some(a) => a,
        None => return,
    };
    let index = {
        let app_lock = app.lock().unwrap();
        app_lock.active_tab
    };
    close_tab_at(hwnd, index);
}

fn close_tab_at(hwnd: HWND, index: usize) {
    let app = match get_app(hwnd) {
        Some(a) => a,
        None => return,
    };
    let should_close = {
        let mut app_lock = app.lock().unwrap();
        app_lock.close_tab(index)
    };
    if should_close {
        unsafe { let _ = DestroyWindow(hwnd); }
    } else {
        repaint(hwnd);
    }
}

fn switch_tab_relative(hwnd: HWND, delta: i32) {
    let app = match get_app(hwnd) {
        Some(a) => a,
        None => return,
    };
    let mut app_lock = app.lock().unwrap();
    let count = app_lock.tab_count();
    if count <= 1 {
        return;
    }
    let current = app_lock.active_tab as i32;
    let next = ((current + delta) % count as i32 + count as i32) as usize % count;
    app_lock.switch_tab(next);
    drop(app_lock);
    repaint(hwnd);
}

// --- クリップボード ---

fn paste_from_clipboard(hwnd: HWND) {
    unsafe {
        if OpenClipboard(Some(hwnd)).is_err() {
            return;
        }
        let handle = GetClipboardData(CF_UNICODETEXT.0 as u32);
        if let Ok(handle) = handle {
            let hglobal = HGLOBAL(handle.0);
            let ptr = GlobalLock(hglobal) as *const u16;
            if !ptr.is_null() {
                // GlobalSize でバッファサイズを取得し、上限を設定して null 終端スキャン
                let buf_bytes = GlobalSize(hglobal);
                let max_u16 = buf_bytes / 2;
                let mut len = 0;
                while len < max_u16 && *ptr.add(len) != 0 {
                    len += 1;
                }
                let slice = std::slice::from_raw_parts(ptr, len);
                if let Ok(text) = String::from_utf16(slice) {
                    let app = get_app(hwnd);
                    if let Some(app) = app {
                        let mut app = app.lock().unwrap();
                        app.scroll_to_bottom();
                        let _ = app.write_pty(text.as_bytes());
                    }
                }
                let _ = GlobalUnlock(hglobal);
            }
        }
        let _ = CloseClipboard();
    }
}

/// ホイールデルタからスクロール行数を計算
/// delta: WM_MOUSEWHEEL の上位ワード（正=上、負=下）
/// lines_per_notch: 1ノッチ (WHEEL_DELTA=120) あたりのスクロール行数
/// 戻り値: スクロール行数（正=上、負=下）。半ノッチ以下でも最低1行。
pub fn calc_scroll_lines(delta: i16, lines_per_notch: u32) -> i32 {
    const WHEEL_DELTA: i32 = 120;
    let raw = (delta as i32) * (lines_per_notch as i32) / WHEEL_DELTA;
    if raw == 0 {
        // 半ノッチ等の小さなデルタでも最低1行は返す
        if delta > 0 { 1 } else if delta < 0 { -1 } else { 0 }
    } else {
        raw
    }
}

/// ALT_SCREEN 時に送信する矢印キーシーケンスを生成
/// lines > 0: 上矢印 (ESC[A) を lines 回
/// lines < 0: 下矢印 (ESC[B) を |lines| 回
pub fn alt_screen_arrow_keys(lines: i32) -> Vec<u8> {
    let (seq, count) = if lines > 0 {
        (b"\x1b[A", lines as usize)
    } else if lines < 0 {
        (b"\x1b[B", lines.unsigned_abs() as usize)
    } else {
        return Vec::new();
    };
    let mut result = Vec::with_capacity(seq.len() * count);
    for _ in 0..count {
        result.extend_from_slice(seq);
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    // --- calc_scroll_lines ---

    #[test]
    fn test_1ノッチ上スクロールで3行() {
        assert_eq!(calc_scroll_lines(120, 3), 3);
    }

    #[test]
    fn test_1ノッチ下スクロールでマイナス3行() {
        assert_eq!(calc_scroll_lines(-120, 3), -3);
    }

    #[test]
    fn test_2ノッチ上スクロールで6行() {
        assert_eq!(calc_scroll_lines(240, 3), 6);
    }

    #[test]
    fn test_半ノッチでも最低1行() {
        // 高精度ホイール: delta=60 (WHEEL_DELTA の半分)
        let result = calc_scroll_lines(60, 3);
        assert!(result >= 1, "半ノッチでも最低1行スクロールすべき: got {result}");
    }

    #[test]
    fn test_半ノッチ下スクロールでも最低マイナス1行() {
        let result = calc_scroll_lines(-60, 3);
        assert!(result <= -1, "半ノッチ下スクロールでも最低-1行すべき: got {result}");
    }

    #[test]
    fn test_lines_per_notchが5のとき() {
        assert_eq!(calc_scroll_lines(120, 5), 5);
    }

    // --- alt_screen_arrow_keys ---

    #[test]
    fn test_上矢印キー3回() {
        let keys = alt_screen_arrow_keys(3);
        // ESC[A = [0x1b, 0x5b, 0x41] × 3
        assert_eq!(keys, b"\x1b[A\x1b[A\x1b[A");
    }

    #[test]
    fn test_下矢印キー2回() {
        let keys = alt_screen_arrow_keys(-2);
        // ESC[B = [0x1b, 0x5b, 0x42] × 2
        assert_eq!(keys, b"\x1b[B\x1b[B");
    }

    #[test]
    fn test_0行なら空() {
        let keys = alt_screen_arrow_keys(0);
        assert!(keys.is_empty());
    }
}
