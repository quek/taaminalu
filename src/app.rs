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
///
/// 行座標は stable_row（スクロールバック先頭を 0 とする絶対行番号）で保持する。
/// stable_row = viewport_row + history_size - display_offset
/// これにより、ホイールスクロール（display_offset 変化）と出力スクロール
/// （history_size 増加）の両方で選択がコンテンツに追従する。
pub struct Selection {
    /// 選択開始位置 (stable_row, col) 0-indexed
    pub start: (usize, usize),
    /// 選択終了位置 (stable_row, col) 0-indexed
    pub end: (usize, usize),
    /// ドラッグ中か
    pub active: bool,
    /// 選択モード
    pub mode: SelectionMode,
    /// Word モードの起点単語範囲 ((stable_row, start_col), (stable_row, end_col))
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

    /// ビューポート座標 (viewport_row, col) が選択範囲内か判定する。
    /// history_size と display_offset からビューポート行を stable_row に変換して比較。
    pub fn viewport_contains(&self, viewport_row: usize, col: usize, history_size: usize, display_offset: usize) -> bool {
        let stable_row = viewport_row + history_size - display_offset;
        self.contains(stable_row, col)
    }

    /// セル (row, col) が選択範囲内か（stable_row 座標）
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

    pub fn scroll_to_bottom(&mut self) {
        self.tabs[self.active_tab].term.scroll_to_bottom();
    }

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

    // --- viewport_contains ---

    #[test]
    fn test_履歴なしでビューポート行がそのままstable_rowになる() {
        // history_size=0, display_offset=0 → stable_row = viewport_row
        let sel = make_selection((5, 0), (7, 10));
        assert!(sel.viewport_contains(6, 5, 0, 0), "選択範囲内");
        assert!(!sel.viewport_contains(4, 5, 0, 0), "選択範囲外");
    }

    #[test]
    fn test_ホイールスクロールで選択がコンテンツに追従する() {
        // history_size=50 で、display_offset=0 のときに画面 row 5-7 を選択
        // stable_row = 5 + 50 - 0 = 55, 57
        let sel = make_selection((55, 0), (57, 10));
        // display_offset=3 で上スクロール → viewport row 8 の stable_row = 8 + 50 - 3 = 55
        assert!(sel.viewport_contains(8, 5, 50, 3), "row 8 は stable_row 55 の内容");
        assert!(!sel.viewport_contains(5, 5, 50, 3), "row 5 は stable_row 52 で範囲外");
    }

    #[test]
    fn test_出力スクロールで選択がコンテンツに追従する() {
        // history_size=50 で画面 row 5 を選択 → stable_row = 55
        let sel = make_selection((55, 0), (57, 10));
        // 新しい出力で history_size が 53 に増加（3行スクロール）、display_offset=0
        // 元の stable_row 55 のコンテンツはビューポート row 2 に移動: 2 + 53 - 0 = 55
        assert!(sel.viewport_contains(2, 5, 53, 0), "3行出力後、row 2 が stable_row 55");
        assert!(!sel.viewport_contains(5, 5, 53, 0), "row 5 は stable_row 58 で範囲外");
    }

    #[test]
    fn test_選択が画面外ならfalse() {
        // stable_row 55-57 の選択、history_size=100, display_offset=0
        // viewport_row 55 → stable_row = 55 + 100 = 155 ≠ 55
        let sel = make_selection((55, 0), (57, 10));
        assert!(!sel.viewport_contains(55, 5, 100, 0), "stable_row が一致しないので範囲外");
    }
}
