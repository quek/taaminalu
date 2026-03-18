use std::io;

use windows::Win32::Foundation::HWND;

use crate::pty::Pty;
use crate::render::Renderer;
use crate::term::TermWrapper;

/// アプリケーション全体の状態
pub struct App {
    pub pty: Pty,
    pub term: TermWrapper,
    pub renderer: Option<Renderer>,
}

impl App {
    pub fn new(cols: usize, rows: usize) -> io::Result<Self> {
        let pty = Pty::new(cols as u16, rows as u16)?;
        let term = TermWrapper::new(cols, rows);

        Ok(Self {
            pty,
            term,
            renderer: None,
        })
    }

    pub fn init_renderer(&mut self, hwnd: HWND, width: u32, height: u32) {
        match Renderer::new(hwnd, width, height) {
            Ok(renderer) => self.renderer = Some(renderer),
            Err(e) => eprintln!("Renderer init failed: {}", e),
        }
    }

    pub fn paint(&self, hwnd: HWND) {
        if let Some(ref renderer) = self.renderer {
            renderer.paint(hwnd, &self.term);
        }
    }

    pub fn on_resize(&mut self, width: u32, height: u32) {
        if let Some(ref renderer) = self.renderer {
            renderer.resize(width, height);
            let (cols, rows) = renderer.calc_grid_size(width, height);
            if cols > 0 && rows > 0 {
                self.term.resize(cols, rows);
                let _ = self.pty.resize(cols as u16, rows as u16);
            }
        }
    }

    pub fn process_pty_output(&mut self, data: &[u8]) {
        self.term.process(data);
    }

    pub fn write_pty(&self, data: &[u8]) -> io::Result<usize> {
        self.pty.write(data)
    }

    pub fn screen_text(&self) -> String {
        self.term.screen_text()
    }

    pub fn cursor_pos(&self) -> (usize, usize) {
        self.term.cursor_pos()
    }

    pub fn columns(&self) -> usize {
        self.term.columns()
    }

    pub fn cell_size(&self) -> (f32, f32) {
        if let Some(ref renderer) = self.renderer {
            (renderer.cell_width, renderer.cell_height)
        } else {
            (8.0, 16.0)
        }
    }
}
