use std::io;
use std::time::Instant;

use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::GetClientRect;

use crate::pty::ShellType;
use crate::render::Renderer;
use crate::tab::{Tab, TabId};

/// 選択モード
#[derive(Clone, Copy, PartialEq)]
pub enum SelectionMode {
    /// 通常のドラッグ選択
    Normal,
    /// ダブルクリックによる単語単位選択
    Word,
}

/// マウスドラッグによるテキスト選択状態
pub struct Selection {
    /// 選択開始位置 (row, col) 0-indexed
    pub start: (usize, usize),
    /// 選択終了位置 (row, col) 0-indexed
    pub end: (usize, usize),
    /// ドラッグ中か
    pub active: bool,
    /// 選択モード
    pub mode: SelectionMode,
    /// Word モードの起点単語範囲 ((row, start_col), (row, end_col))
    pub origin_word: Option<((usize, usize), (usize, usize))>,
}

impl Selection {
    /// 正規化: start が end より前になるよう並べ替え
    pub fn ordered(&self) -> ((usize, usize), (usize, usize)) {
        if self.start.0 < self.end.0
            || (self.start.0 == self.end.0 && self.start.1 <= self.end.1)
        {
            (self.start, self.end)
        } else {
            (self.end, self.start)
        }
    }

    /// スクロール時に選択座標をビューポートに追従させる
    /// delta > 0: 上スクロール（コンテンツが画面下方向に移動 → row 増加）
    /// delta < 0: 下スクロール（コンテンツが画面上方向に移動 → row 減少）
    /// 選択範囲が完全に画面外に出た場合は false を返す
    pub fn adjust_for_scroll(&mut self, delta: i32, screen_lines: usize) -> bool {
        let new_start = self.start.0 as i32 + delta;
        let new_end = self.end.0 as i32 + delta;

        // 両方とも画面外なら不可視
        let max_line = screen_lines as i32 - 1;
        if (new_start > max_line && new_end > max_line) || (new_start < 0 && new_end < 0) {
            return false;
        }

        self.start.0 = new_start.max(0) as usize;
        self.end.0 = new_end.max(0) as usize;
        true
    }

    /// セル (row, col) が選択範囲内か
    pub fn contains(&self, row: usize, col: usize) -> bool {
        let ((sr, sc), (er, ec)) = self.ordered();
        if row < sr || row > er {
            return false;
        }
        if sr == er {
            col >= sc && col <= ec
        } else if row == sr {
            col >= sc
        } else if row == er {
            col <= ec
        } else {
            true
        }
    }
}

/// スクロール後に選択範囲を調整。画面外に出たら None にする。
pub fn adjust_selection_after_scroll(selection: &mut Option<Selection>, delta: i32, screen_lines: usize) {
    if let Some(sel) = selection {
        if !sel.adjust_for_scroll(delta, screen_lines) {
            *selection = None;
        }
    }
}

/// アプリケーション全体の状態
pub struct App {
    pub tabs: Vec<Tab>,
    pub active_tab: usize,
    pub renderer: Option<Renderer>,
    pub selection: Option<Selection>,
    /// マウスボタン押下位置（ピクセル座標 + グリッド座標、ドラッグ開始判定用）
    pub drag_origin: Option<(i16, i16, usize, usize)>,
    /// 直前のクリック情報（ダブルクリック検出用: 時刻, row, col）
    pub last_click: Option<(Instant, usize, usize)>,
}

impl App {
    pub fn new(cols: usize, rows: usize) -> io::Result<Self> {
        let tab = Tab::new(cols, rows, ShellType::Cmd)?;
        Ok(Self {
            tabs: vec![tab],
            active_tab: 0,
            renderer: None,
            selection: None,
            drag_origin: None,
            last_click: None,
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
            renderer.paint_with_tabs(hwnd, &self.active().term, &tab_infos, self.active_tab, preedit, self.selection.as_ref());
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

    /// 画面座標からグリッド座標に変換
    pub fn screen_to_grid(&self, x: f32, y: f32) -> (usize, usize) {
        let (cell_w, cell_h) = self.cell_size();
        let (_, grid_y) = self.grid_origin();
        let (cols, rows) = self.grid_size();
        let row = ((y - grid_y) / cell_h).max(0.0) as usize;
        let col = (x / cell_w).max(0.0) as usize;
        (row.min(rows.saturating_sub(1)), col.min(cols.saturating_sub(1)))
    }

    pub fn grid_size(&self) -> (usize, usize) {
        self.renderer.as_ref()
            .map(|r| r.current_grid_size())
            .unwrap_or((80, 24))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_selection(start: (usize, usize), end: (usize, usize)) -> Selection {
        Selection {
            start,
            end,
            active: false,
            mode: SelectionMode::Normal,
            origin_word: None,
        }
    }

    #[test]
    fn test_上スクロールで選択行が増加() {
        let mut sel = make_selection((5, 0), (7, 10));
        let visible = sel.adjust_for_scroll(3, 24);
        assert!(visible);
        assert_eq!(sel.start.0, 8, "start.row が 3 増加すべき");
        assert_eq!(sel.end.0, 10, "end.row が 3 増加すべき");
        // 列は変わらない
        assert_eq!(sel.start.1, 0);
        assert_eq!(sel.end.1, 10);
    }

    #[test]
    fn test_下スクロールで選択行が減少() {
        let mut sel = make_selection((8, 0), (10, 5));
        let visible = sel.adjust_for_scroll(-3, 24);
        assert!(visible);
        assert_eq!(sel.start.0, 5);
        assert_eq!(sel.end.0, 7);
    }

    #[test]
    fn test_上スクロールで選択が画面外に出たらfalse() {
        // 選択が row 20-22、画面24行で delta=5 → row 25-27 → 画面外
        let mut sel = make_selection((20, 0), (22, 5));
        let visible = sel.adjust_for_scroll(5, 24);
        assert!(!visible, "選択が完全に画面外に出たら false");
    }

    #[test]
    fn test_下スクロールで選択が画面外に出たらfalse() {
        // 選択が row 1-2、delta=-3 → 負の行 → 画面外
        let mut sel = make_selection((1, 0), (2, 5));
        let visible = sel.adjust_for_scroll(-3, 24);
        assert!(!visible, "選択が完全に画面外に出たら false");
    }

    // --- adjust_selection_after_scroll ---

    #[test]
    fn test_選択ありスクロールで座標が調整される() {
        let mut sel = Some(make_selection((5, 0), (7, 10)));
        adjust_selection_after_scroll(&mut sel, 3, 24);
        let s = sel.as_ref().expect("選択が残っているべき");
        assert_eq!(s.start.0, 8);
        assert_eq!(s.end.0, 10);
    }

    #[test]
    fn test_選択ありスクロールで画面外になったら消える() {
        let mut sel = Some(make_selection((20, 0), (22, 5)));
        adjust_selection_after_scroll(&mut sel, 5, 24);
        assert!(sel.is_none(), "画面外に出たら None になるべき");
    }

    #[test]
    fn test_選択なしスクロールでも変化なし() {
        let mut sel: Option<Selection> = None;
        adjust_selection_after_scroll(&mut sel, 3, 24);
        assert!(sel.is_none());
    }
}
