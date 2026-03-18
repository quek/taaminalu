use std::collections::HashMap;
use std::sync::{Arc, Mutex, OnceLock};

use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::InvalidateRect;
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::*;

use crate::app::App;
use crate::tsf::TsfContext;

const CLASS_NAME: &str = "TaaminaluWindow";
const WINDOW_TITLE: &str = "taaminalu";
const DEFAULT_WIDTH: i32 = 800;
const DEFAULT_HEIGHT: i32 = 600;

/// カスタムメッセージ: PTY からデータ受信で再描画要求
pub const WM_PTY_OUTPUT: u32 = WM_USER + 1;

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

        let wc = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(wnd_proc),
            hInstance: hinstance.into(),
            lpszClassName: windows::core::PCWSTR(class_name_wide.as_ptr()),
            hCursor: LoadCursorW(None, IDC_ARROW)?,
            ..Default::default()
        };

        RegisterClassExW(&wc);

        // App を Box でヒープに確保して LPARAM で渡す
        let app_ptr = Box::into_raw(Box::new(app));

        let hwnd = CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            windows::core::PCWSTR(class_name_wide.as_ptr()),
            windows::core::PCWSTR(title_wide.as_ptr()),
            WS_OVERLAPPEDWINDOW | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            DEFAULT_WIDTH,
            DEFAULT_HEIGHT,
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
        WM_CHAR => {
            let ch = wparam.0 as u32;
            if let Some(c) = char::from_u32(ch) {
                // WM_KEYDOWN で処理済みのキーは WM_CHAR では送らない
                if !is_handled_by_keydown(c) {
                    let app = get_app(hwnd);
                    if let Some(app) = app {
                        let app = app.lock().unwrap();
                        let mut buf = [0u8; 4];
                        let s = c.encode_utf8(&mut buf);
                        let _ = app.write_pty(s.as_bytes());
                    }
                }
            }
            LRESULT(0)
        }
        WM_KEYDOWN => {
            let vk = wparam.0 as u16;
            let seq = vk_to_escape_seq(vk);
            if let Some(seq) = seq {
                let app = get_app(hwnd);
                if let Some(app) = app {
                    let app = app.lock().unwrap();
                    let _ = app.write_pty(seq);
                }
            }
            LRESULT(0)
        }
        WM_PTY_OUTPUT => {
            unsafe {
                let _ = InvalidateRect(Some(hwnd), None, false);
            }
            // TSF シンクにテキスト/カーソル変更を通知
            notify_tsf_change(hwnd);
            LRESULT(0)
        }
        WM_DESTROY => {
            // App ポインタをクリーンアップ
            let ptr = unsafe { GetWindowLongPtrW(hwnd, GWLP_USERDATA) } as *mut Arc<Mutex<App>>;
            if !ptr.is_null() {
                unsafe {
                    drop(Box::from_raw(ptr));
                }
                unsafe {
                    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                }
            }
            unsafe {
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

/// WM_KEYDOWN で処理済みで WM_CHAR では送らないキー
fn is_handled_by_keydown(c: char) -> bool {
    matches!(c, '\r' | '\x7f' | '\x08' | '\t' | '\x1b')
}

/// 仮想キーコード → VT エスケープシーケンス
fn vk_to_escape_seq(vk: u16) -> Option<&'static [u8]> {
    use windows::Win32::UI::Input::KeyboardAndMouse::*;
    match VIRTUAL_KEY(vk) {
        VK_UP => Some(b"\x1b[A"),
        VK_DOWN => Some(b"\x1b[B"),
        VK_RIGHT => Some(b"\x1b[C"),
        VK_LEFT => Some(b"\x1b[D"),
        VK_HOME => Some(b"\x1b[H"),
        VK_END => Some(b"\x1b[F"),
        VK_DELETE => Some(b"\x1b[3~"),
        VK_PRIOR => Some(b"\x1b[5~"), // Page Up
        VK_NEXT => Some(b"\x1b[6~"),  // Page Down
        VK_INSERT => Some(b"\x1b[2~"),
        VK_RETURN => Some(b"\r"),
        VK_BACK => Some(b"\x7f"),
        VK_TAB => Some(b"\t"),
        VK_ESCAPE => Some(b"\x1b"),
        _ => None,
    }
}
