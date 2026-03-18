use std::io;

use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::GetClientRect;

use crate::pty::ShellType;
use crate::render::Renderer;
use crate::tab::{Tab, TabId};

/// アプリケーション全体の状態
pub struct App {
    pub tabs: Vec<Tab>,
    pub active_tab: usize,
    pub renderer: Option<Renderer>,
}

impl App {
    pub fn new(cols: usize, rows: usize) -> io::Result<Self> {
        let tab = Tab::new(cols, rows, ShellType::Cmd)?;
        Ok(Self {
            tabs: vec![tab],
            active_tab: 0,
            renderer: None,
        })
    }

    /// Renderer を初期化し、グリッドサイズを再計算して全タブをリサイズ
    pub fn init_renderer(&mut self, hwnd: HWND) {
        let mut rect = windows::Win32::Foundation::RECT::default();
        unsafe { let _ = GetClientRect(hwnd, &mut rect); }
        let width = (rect.right - rect.left).max(1) as u32;
        let height = (rect.bottom - rect.top).max(1) as u32;

        match Renderer::new(hwnd, width, height) {
            Ok(renderer) => {
                let (cols, rows) = renderer.calc_grid_size(width, height);
                self.renderer = Some(renderer);
                if cols > 0 && rows > 0 {
                    for tab in &mut self.tabs {
                        tab.resize(cols, rows);
                    }
                }
            }
            Err(e) => eprintln!("Renderer init failed: {}", e),
        }
    }

    pub fn active(&self) -> &Tab {
        &self.tabs[self.active_tab]
    }

    pub fn paint(&self, hwnd: HWND, preedit: &str) {
        if let Some(ref renderer) = self.renderer {
            let tab_infos: Vec<(&str, TabId)> = self.tabs.iter()
                .map(|t| (t.title.as_str(), t.id))
                .collect();
            renderer.paint_with_tabs(hwnd, &self.active().term, &tab_infos, self.active_tab, preedit);
        }
    }

    pub fn on_resize(&mut self, width: u32, height: u32) {
        if let Some(ref renderer) = self.renderer {
            renderer.resize(width, height);
            let (cols, rows) = renderer.calc_grid_size(width, height);
            if cols > 0 && rows > 0 {
                for tab in &mut self.tabs {
                    tab.resize(cols, rows);
                }
            }
        }
    }

    pub fn process_pty_output_for_tab(&mut self, tab_id: TabId, data: &[u8]) {
        if let Some(tab) = self.tabs.iter_mut().find(|t| t.id == tab_id) {
            tab.process_pty_output(data);
        }
    }

    // --- アクティブタブへの委譲 ---

    pub fn write_pty(&self, data: &[u8]) -> io::Result<usize> {
        self.active().write_pty(data)
    }

    pub fn screen_text(&self) -> String {
        self.active().term.screen_text()
    }

    pub fn screen_text_utf16_len(&self) -> usize {
        self.active().term.screen_text_utf16_len()
    }

    pub fn cursor_acp(&self) -> usize {
        self.active().term.cursor_acp()
    }

    pub fn acp_to_grid(&self, acp: usize) -> (usize, usize) {
        self.active().term.acp_to_grid(acp)
    }

    pub fn cell_size(&self) -> (f32, f32) {
        self.renderer.as_ref()
            .map(|r| (r.cell_width, r.cell_height))
            .unwrap_or((8.0, 16.0))
    }

    pub fn grid_origin(&self) -> (f32, f32) {
        self.renderer.as_ref()
            .map(|r| (0.0, r.tab_bar_height()))
            .unwrap_or((0.0, 30.0))
    }

    // --- タブ管理 ---

    pub fn add_tab(&mut self, shell: ShellType, cols: usize, rows: usize) -> io::Result<TabId> {
        let tab = Tab::new(cols, rows, shell)?;
        let id = tab.id;
        self.tabs.push(tab);
        self.active_tab = self.tabs.len() - 1;
        Ok(id)
    }

    /// タブを閉じる。最後のタブだった場合は true を返す
    pub fn close_tab(&mut self, index: usize) -> bool {
        if index >= self.tabs.len() {
            return false;
        }
        self.tabs.remove(index);
        if self.tabs.is_empty() {
            return true;
        }
        if self.active_tab >= self.tabs.len() {
            self.active_tab = self.tabs.len() - 1;
        }
        false
    }

    pub fn switch_tab(&mut self, index: usize) {
        if index < self.tabs.len() {
            self.active_tab = index;
        }
    }

    pub fn find_tab_index(&self, id: TabId) -> Option<usize> {
        self.tabs.iter().position(|t| t.id == id)
    }

    pub fn tab_count(&self) -> usize {
        self.tabs.len()
    }

    pub fn active_tab_id(&self) -> TabId {
        self.active().id
    }

    pub fn grid_size(&self) -> (usize, usize) {
        self.renderer.as_ref()
            .map(|r| r.current_grid_size())
            .unwrap_or((80, 24))
    }
}
