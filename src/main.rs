mod app;
mod pty;
mod render;
mod tab;
mod term;
mod tsf;
mod window;

use std::sync::{Arc, Mutex};

use windows::Win32::Foundation::HWND;
use windows::Win32::System::Com::{CoInitializeEx, COINIT_APARTMENTTHREADED};
use windows::Win32::UI::WindowsAndMessaging::GetClientRect;

use app::App;

fn main() {
    eprintln!("[taaminalu] starting...");

    // COM 初期化（STA — TSF に必要）
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }

    let initial_cols = 80usize;
    let initial_rows = 24usize;

    let app = match App::new(initial_cols, initial_rows) {
        Ok(app) => app,
        Err(e) => {
            eprintln!("[taaminalu] Failed to create app: {}", e);
            return;
        }
    };
    let app = Arc::new(Mutex::new(app));

    // ウィンドウ作成
    let hwnd = match window::create_window(Arc::clone(&app)) {
        Ok(hwnd) => hwnd,
        Err(e) => {
            eprintln!("[taaminalu] Failed to create window: {}", e);
            return;
        }
    };

    // Renderer 初期化 + 初期タブのグリッドリサイズ
    init_renderer_and_resize(&app, hwnd);

    // TSF セットアップ
    let tsf_ctx = tsf::setup_tsf(Arc::clone(&app), hwnd).ok();
    // TSF コンテキストを window に保存（WM_PTY_OUTPUT で通知に使う）
    window::set_tsf_context(hwnd, tsf_ctx);

    // 初期タブの PTY 読み取りスレッド
    let (tab_id, pty_read_handle) = {
        let app_lock = app.lock().unwrap();
        let tab = &app_lock.tabs[0];
        let id = tab.id;
        let handle = tab.dup_output_read().expect("Failed to duplicate PTY read handle");
        (id, handle)
    };
    window::start_pty_reader(Arc::clone(&app), hwnd, pty_read_handle, tab_id);

    // メッセージループ
    window::run_message_loop();
}

fn init_renderer_and_resize(app: &Arc<Mutex<App>>, hwnd: HWND) {
    let mut rect = windows::Win32::Foundation::RECT::default();
    unsafe {
        let _ = GetClientRect(hwnd, &mut rect);
    }
    let width = (rect.right - rect.left) as u32;
    let height = (rect.bottom - rect.top) as u32;

    let mut app_lock = app.lock().unwrap();
    app_lock.init_renderer(hwnd, width.max(1), height.max(1));

    // Renderer のセルサイズでグリッドサイズ再計算（タブバー高さ考慮済み）
    if let Some(ref renderer) = app_lock.renderer {
        let (cols, rows) = renderer.calc_grid_size(width, height);
        if cols > 0 && rows > 0 {
            for tab in &mut app_lock.tabs {
                tab.resize(cols, rows);
            }
        }
    }
}
