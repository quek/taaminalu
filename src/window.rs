use std::collections::HashMap;
use std::fs;
use std::path::PathBuf;
use std::sync::{Arc, Mutex, OnceLock};
use std::thread;

use serde::{Deserialize, Serialize};
use windows::Win32::Foundation::{
    CloseHandle, HGLOBAL, HANDLE, HWND, LPARAM, LRESULT, POINT, RECT, WPARAM,
};
use windows::Win32::Graphics::Gdi::InvalidateRect;
use windows::Win32::Storage::FileSystem::ReadFile;
use windows::Win32::System::DataExchange::{CloseClipboard, GetClipboardData, OpenClipboard};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Memory::{GlobalLock, GlobalUnlock};
use windows::Win32::System::Ole::CF_UNICODETEXT;
use windows::Win32::System::Threading::{WaitForSingleObject, INFINITE};
use windows::Win32::UI::Input::KeyboardAndMouse::*;
use windows::Win32::UI::WindowsAndMessaging::*;

use crate::app::App;
use crate::pty::ShellType;
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
const MENU_ID_WSL: u32 = 1;
const MENU_ID_CMD: u32 = 2;
const MENU_ID_POWERSHELL: u32 = 3;

/// HWND ごとの TSF コンテキスト
static TSF_CONTEXTS: OnceLock<Mutex<HashMap<isize, TsfContext>>> = OnceLock::new();

pub fn set_tsf_context(hwnd: HWND, ctx: Option<TsfContext>) {
    let map = TSF_CONTEXTS.get_or_init(|| Mutex::new(HashMap::new()));
    let mut map = map.lock().unwrap();
    if let Some(ctx) = ctx {
        map.insert(hwnd.0 as isize, ctx);
    }
}

fn notify_tsf_change(hwnd: HWND) {
    if let Some(map) = TSF_CONTEXTS.get() {
        let map = map.lock().unwrap();
        if let Some(ctx) = map.get(&(hwnd.0 as isize)) {
            ctx.notify_change();
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
            windows::core::PCWSTR(1 as *const u16),
        )?;

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
        WM_PAINT => {
            let app = get_app(hwnd);
            if let Some(app) = app {
                let app = app.lock().unwrap();
                app.paint(hwnd);
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
            let ch = wparam.0 as u32;
            if let Some(c) = char::from_u32(ch) {
                // VK_BACK は WM_KEYDOWN で 0x7F として送信済み
                if c == '\x7f' {
                    return LRESULT(0);
                }
                let app = get_app(hwnd);
                if let Some(app) = app {
                    let app = app.lock().unwrap();
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

            // Ctrl+Shift+V: ペースト
            if vk == VK_V && mods.ctrl && mods.shift {
                paste_from_clipboard(hwnd);
                return LRESULT(0);
            }

            // 特殊キーのエスケープシーケンス
            if let Some(seq) = build_key_sequence(vk, &mods) {
                let app = get_app(hwnd);
                if let Some(app) = app {
                    let app = app.lock().unwrap();
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
            let x = (lparam.0 & 0xFFFF) as i16 as f32;
            let y = ((lparam.0 >> 16) & 0xFFFF) as i16 as f32;
            handle_tab_bar_click(hwnd, x, y);
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
                unsafe {
                    let _ = InvalidateRect(Some(hwnd), None, false);
                }
                notify_tsf_change(hwnd);
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
                        unsafe { let _ = InvalidateRect(Some(hwnd), None, false); }
                        notify_tsf_change(hwnd);
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
            unsafe { let _ = InvalidateRect(Some(hwnd), None, false); }
            notify_tsf_change(hwnd);
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
        let menu = CreatePopupMenu();
        let menu = match menu {
            Ok(m) => m,
            Err(_) => return,
        };

        let wsl_text: Vec<u16> = "WSL\0".encode_utf16().collect();
        let cmd_text: Vec<u16> = "CMD\0".encode_utf16().collect();
        let ps_text: Vec<u16> = "PowerShell\0".encode_utf16().collect();

        let _ = AppendMenuW(menu, MF_STRING, MENU_ID_WSL as usize, windows::core::PCWSTR(wsl_text.as_ptr()));
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
                MENU_ID_WSL => ShellType::Wsl,
                MENU_ID_CMD => ShellType::Cmd,
                MENU_ID_POWERSHELL => ShellType::PowerShell,
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
    unsafe { let _ = InvalidateRect(Some(hwnd), None, false); }
    notify_tsf_change(hwnd);
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
        unsafe { let _ = InvalidateRect(Some(hwnd), None, false); }
        notify_tsf_change(hwnd);
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
    unsafe { let _ = InvalidateRect(Some(hwnd), None, false); }
    notify_tsf_change(hwnd);
}

// --- 修飾キー状態 ---

struct Modifiers {
    shift: bool,
    alt: bool,
    ctrl: bool,
}

fn get_modifiers() -> Modifiers {
    unsafe {
        Modifiers {
            shift: GetKeyState(VK_SHIFT.0 as i32) < 0,
            alt: GetKeyState(VK_MENU.0 as i32) < 0,
            ctrl: GetKeyState(VK_CONTROL.0 as i32) < 0,
        }
    }
}

/// xterm 修飾キーパラメータ: 1 + (Shift=1 | Alt=2 | Ctrl=4)
fn modifier_param(mods: &Modifiers) -> u8 {
    let mut p = 0u8;
    if mods.shift { p |= 1; }
    if mods.alt { p |= 2; }
    if mods.ctrl { p |= 4; }
    1 + p
}

fn has_modifiers(mods: &Modifiers) -> bool {
    mods.shift || mods.alt || mods.ctrl
}

// --- キーシーケンス生成 ---

/// 特殊キー → VT エスケープシーケンス (修飾キー対応)
fn build_key_sequence(vk: VIRTUAL_KEY, mods: &Modifiers) -> Option<Vec<u8>> {
    // Backspace: 修飾キー対応
    if vk == VK_BACK {
        let mut seq = Vec::new();
        if mods.alt { seq.push(0x1b); }
        if mods.ctrl {
            seq.push(0x08); // Ctrl+Backspace = BS
        } else {
            seq.push(0x7f); // Backspace = DEL
        }
        return Some(seq);
    }

    // CSI キー: 矢印、Home/End、Insert/Delete、PageUp/Down
    if let Some((code, suffix)) = csi_key_params(vk) {
        let mp = modifier_param(mods);
        let seq = if mp > 1 {
            // 修飾キーあり: \x1b[1;{mod}{suffix} or \x1b[{code};{mod}~
            if suffix == b'~' {
                format!("\x1b[{};{}~", code, mp).into_bytes()
            } else {
                format!("\x1b[1;{}{}", mp, suffix as char).into_bytes()
            }
        } else {
            // 修飾キーなし
            if suffix == b'~' {
                format!("\x1b[{}~", code).into_bytes()
            } else {
                vec![0x1b, b'[', suffix]
            }
        };
        return Some(seq);
    }

    // ファンクションキー F1-F12
    if let Some(seq) = function_key_sequence(vk, mods) {
        return Some(seq);
    }

    None
}

/// CSI キーのパラメータ: (数値コード, サフィックス文字)
/// サフィックスが '~' の場合は \x1b[{code}~ 形式
/// それ以外は \x1b[{suffix} 形式
fn csi_key_params(vk: VIRTUAL_KEY) -> Option<(u8, u8)> {
    match vk {
        VK_UP => Some((1, b'A')),
        VK_DOWN => Some((1, b'B')),
        VK_RIGHT => Some((1, b'C')),
        VK_LEFT => Some((1, b'D')),
        VK_HOME => Some((1, b'H')),
        VK_END => Some((1, b'F')),
        VK_INSERT => Some((2, b'~')),
        VK_DELETE => Some((3, b'~')),
        VK_PRIOR => Some((5, b'~')), // Page Up
        VK_NEXT => Some((6, b'~')),  // Page Down
        _ => None,
    }
}

/// ファンクションキー F1-F12 → エスケープシーケンス
fn function_key_sequence(vk: VIRTUAL_KEY, mods: &Modifiers) -> Option<Vec<u8>> {
    // F1-F4: SS3 形式 (修飾キーなし), CSI 形式 (修飾キーあり)
    // F5-F12: CSI {code}~ 形式
    let mp = modifier_param(mods);
    let has_mods = has_modifiers(mods);

    match vk {
        VK_F1 => Some(if has_mods {
            format!("\x1b[1;{}P", mp).into_bytes()
        } else {
            b"\x1bOP".to_vec()
        }),
        VK_F2 => Some(if has_mods {
            format!("\x1b[1;{}Q", mp).into_bytes()
        } else {
            b"\x1bOQ".to_vec()
        }),
        VK_F3 => Some(if has_mods {
            format!("\x1b[1;{}R", mp).into_bytes()
        } else {
            b"\x1bOR".to_vec()
        }),
        VK_F4 => Some(if has_mods {
            format!("\x1b[1;{}S", mp).into_bytes()
        } else {
            b"\x1bOS".to_vec()
        }),
        VK_F5 => Some(fkey_csi(15, mp, has_mods)),
        VK_F6 => Some(fkey_csi(17, mp, has_mods)),
        VK_F7 => Some(fkey_csi(18, mp, has_mods)),
        VK_F8 => Some(fkey_csi(19, mp, has_mods)),
        VK_F9 => Some(fkey_csi(20, mp, has_mods)),
        VK_F10 => Some(fkey_csi(21, mp, has_mods)),
        VK_F11 => Some(fkey_csi(23, mp, has_mods)),
        VK_F12 => Some(fkey_csi(24, mp, has_mods)),
        _ => None,
    }
}

/// F5-F12 の CSI シーケンス: \x1b[{code}~ or \x1b[{code};{mod}~
fn fkey_csi(code: u8, mp: u8, has_mods: bool) -> Vec<u8> {
    if has_mods {
        format!("\x1b[{};{}~", code, mp).into_bytes()
    } else {
        format!("\x1b[{}~", code).into_bytes()
    }
}

// --- クリップボード ---

fn paste_from_clipboard(hwnd: HWND) {
    unsafe {
        if OpenClipboard(Some(hwnd)).is_err() {
            return;
        }
        let handle = GetClipboardData(CF_UNICODETEXT.0 as u32);
        if let Ok(handle) = handle {
            let ptr = GlobalLock(HGLOBAL(handle.0)) as *const u16;
            if !ptr.is_null() {
                // null 終端の UTF-16 文字列を読み取り
                let mut len = 0;
                while *ptr.add(len) != 0 {
                    len += 1;
                }
                let slice = std::slice::from_raw_parts(ptr, len);
                if let Ok(text) = String::from_utf16(slice) {
                    let app = get_app(hwnd);
                    if let Some(app) = app {
                        let app = app.lock().unwrap();
                        let _ = app.write_pty(text.as_bytes());
                    }
                }
                let _ = GlobalUnlock(HGLOBAL(handle.0));
            }
        }
        let _ = CloseClipboard();
    }
}
