mod app;
mod pty;
mod render;
mod term;
mod tsf;
mod window;

use std::sync::{Arc, Mutex};
use std::thread;

use windows::Win32::Foundation::{CloseHandle, HANDLE, HWND, LPARAM, WPARAM};
use windows::Win32::Storage::FileSystem::ReadFile;
use windows::Win32::System::Com::{CoInitializeEx, COINIT_APARTMENTTHREADED};
use windows::Win32::UI::WindowsAndMessaging::{GetClientRect, PostMessageW};

use app::App;
use window::WM_PTY_OUTPUT;

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

    // Renderer 初期化
    {
        let mut rect = windows::Win32::Foundation::RECT::default();
        unsafe {
            let _ = GetClientRect(hwnd, &mut rect);
        }
        let width = (rect.right - rect.left) as u32;
        let height = (rect.bottom - rect.top) as u32;

        let mut app_lock = app.lock().unwrap();
        app_lock.init_renderer(hwnd, width.max(1), height.max(1));

        // Renderer のセルサイズでグリッドサイズ再計算
        if let Some(ref renderer) = app_lock.renderer {
            let (cols, rows) = renderer.calc_grid_size(width, height);
            if cols > 0 && rows > 0 {
                app_lock.term.resize(cols, rows);
                let _ = app_lock.pty.resize(cols as u16, rows as u16);
            }
        }
    }

    // TSF セットアップ
    let _tsf = tsf::setup_tsf(Arc::clone(&app), hwnd);

    // PTY 読み取りスレッド（ハンドルを複製してロック不要にする）
    let pty_read_handle = {
        let app_lock = app.lock().unwrap();
        app_lock.pty.dup_output_read().expect("Failed to duplicate PTY read handle")
    };
    start_pty_reader(Arc::clone(&app), hwnd, pty_read_handle);

    // メッセージループ
    window::run_message_loop();
}

fn start_pty_reader(app: Arc<Mutex<App>>, hwnd: HWND, read_handle: HANDLE) {
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
                    app_lock.process_pty_output(&buf[..n]);
                    drop(app_lock);
                    unsafe {
                        let _ = PostMessageW(Some(hwnd), WM_PTY_OUTPUT, WPARAM(0), LPARAM(0));
                    }
                }
                Err(_) => break,
            }
        }
        unsafe {
            let _ = CloseHandle(read_handle);
        }
    });
}
