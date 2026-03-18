mod app;
mod pty;
mod render;
mod tab;
mod term;
mod tsf;
mod window;

use std::sync::{Arc, Mutex};

use windows::Win32::System::Com::{CoInitializeEx, COINIT_APARTMENTTHREADED};

use app::App;

fn main() {
    eprintln!("[taaminalu] starting...");

    // COM 初期化（STA — TSF に必要）
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }

    let app = match App::new(80, 24) {
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

    // Renderer 初期化 + グリッドリサイズ
    app.lock().unwrap().init_renderer(hwnd);

    // TSF セットアップ
    let tsf_ctx = tsf::setup_tsf(Arc::clone(&app), hwnd).ok();
    window::set_tsf_context(hwnd, tsf_ctx);

    // 初期タブの PTY リーダー + プロセス監視スレッドを起動
    {
        let app_lock = app.lock().unwrap();
        let tab = &app_lock.tabs[0];
        let read_h = tab.dup_output_read().expect("Failed to duplicate PTY read handle");
        let proc_h = tab.dup_process_handle().expect("Failed to duplicate process handle");
        window::start_pty_reader(Arc::clone(&app), hwnd, read_h, tab.id);
        window::start_process_watcher(hwnd, proc_h, tab.id);
    }

    // メッセージループ
    window::run_message_loop();
}
