use alacritty_terminal::event::{Event, EventListener};
use alacritty_terminal::grid::{Dimensions, Scroll};
use alacritty_terminal::index::{Column, Line};
use alacritty_terminal::term::cell::Flags;
use alacritty_terminal::term::Config;
use alacritty_terminal::term::Term;
use alacritty_terminal::term::TermMode;
use alacritty_terminal::vte::ansi::Processor;

use windows::Win32::System::DataExchange::{
    CloseClipboard, EmptyClipboard, OpenClipboard, SetClipboardData,
};
use windows::Win32::System::Memory::{
    GlobalAlloc, GlobalLock, GlobalUnlock, GMEM_MOVEABLE,
};
use windows::Win32::System::Ole::CF_UNICODETEXT;

/// alacritty_terminal のイベントリスナー
/// OSC 52 (ClipboardStore) で Windows クリップボードに書き込む
pub struct TermEventListener;

impl EventListener for TermEventListener {
    fn send_event(&self, event: Event) {
        if let Event::ClipboardStore(_, text) = event {
            set_clipboard_text(&text);
        }
    }
}

/// Windows クリップボードにテキストを書き込む
pub fn set_clipboard_text(text: &str) {
    unsafe {
        if OpenClipboard(None).is_err() {
            return;
        }
        let _ = EmptyClipboard();

        let wide: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();
        let byte_len = wide.len() * 2;
        let hmem = GlobalAlloc(GMEM_MOVEABLE, byte_len);
        if let Ok(hmem) = hmem {
            let ptr = GlobalLock(hmem) as *mut u16;
            if !ptr.is_null() {
                std::ptr::copy_nonoverlapping(wide.as_ptr(), ptr, wide.len());
                let _ = GlobalUnlock(hmem);
                let _ = SetClipboardData(CF_UNICODETEXT.0 as u32, Some(windows::Win32::Foundation::HANDLE(hmem.0)));
            }
        }
        let _ = CloseClipboard();
    }
}

/// ターミナルサイズ
pub struct TermSize {
    pub cols: usize,
    pub rows: usize,
}

impl Dimensions for TermSize {
    fn total_lines(&self) -> usize {
        self.rows
    }
    fn screen_lines(&self) -> usize {
        self.rows
    }
    fn columns(&self) -> usize {
        self.cols
    }
}

/// alacritty_terminal のラッパー
pub struct TermWrapper {
    term: Term<TermEventListener>,
    parser: Processor,
}

impl TermWrapper {
    pub fn new(cols: usize, rows: usize) -> Self {
        let size = TermSize { cols, rows };
        let config = Config::default();
        let term = Term::new(config, &size, TermEventListener);
        Self {
            term,
            parser: Processor::new(),
        }
    }

    /// PTY から受信したバイト列を処理
    pub fn process(&mut self, bytes: &[u8]) {
        self.parser.advance(&mut self.term, bytes);
    }


    /// 選択範囲のテキストを抽出
    /// start/end は (row, col) 0-indexed。行末の空白はトリム、行間は \n で連結。
    pub fn selected_text(&self, start: (usize, usize), end: (usize, usize)) -> String {
        let grid = self.term.grid();
        let cols = grid.columns();
        let lines = grid.screen_lines();
        let display_offset = grid.display_offset() as i32;

        // start/end を正規化
        let (start, end) = if start.0 < end.0 || (start.0 == end.0 && start.1 <= end.1) {
            (start, end)
        } else {
            (end, start)
        };

        let mut result = String::new();
        for row_idx in start.0..=end.0.min(lines.saturating_sub(1)) {
            let row = &grid[Line(row_idx as i32 - display_offset)];
            let col_start = if row_idx == start.0 { start.1 } else { 0 };
            let col_end = if row_idx == end.0 { end.1 } else { cols.saturating_sub(1) };

            let mut line = String::new();
            for col in col_start..=col_end.min(cols.saturating_sub(1)) {
                let cell = &row[Column(col)];
                if cell.flags.contains(Flags::WIDE_CHAR_SPACER) {
                    continue;
                }
                line.push(cell.c);
            }
            let trimmed = line.trim_end();
            result.push_str(trimmed);
            if row_idx < end.0 {
                result.push('\n');
            }
        }
        result
    }

    /// 全画面テキスト取得（wide char spacer をスキップ、行末トリムなし）
    /// 各行は固定長ではなく、spacer スキップ後の実文字 + '\n' で構成
    pub fn screen_text(&self) -> String {
        let grid = self.term.grid();
        let cols = grid.columns();
        let lines = grid.screen_lines();
        let mut text = String::new();

        for line_idx in 0..lines {
            let row = &grid[Line(line_idx as i32)];
            for col in 0..cols {
                let cell = &row[Column(col)];
                // wide char spacer (全角文字の2セル目) はスキップ
                if cell.flags.contains(Flags::WIDE_CHAR_SPACER) {
                    continue;
                }
                text.push(cell.c);
            }
            if line_idx + 1 < lines {
                text.push('\n');
            }
        }
        text
    }

    /// screen_text の UTF-16 長を String を作らずに計算
    pub fn screen_text_utf16_len(&self) -> usize {
        let grid = self.term.grid();
        let cols = grid.columns();
        let lines = grid.screen_lines();
        let mut len = 0usize;
        for line_idx in 0..lines {
            let row = &grid[Line(line_idx as i32)];
            for col in 0..cols {
                let cell = &row[Column(col)];
                if !cell.flags.contains(Flags::WIDE_CHAR_SPACER) {
                    len += cell.c.len_utf16();
                }
            }
            if line_idx + 1 < lines {
                len += 1;
            }
        }
        len
    }

    /// グリッド座標 (row, col) を ACP (screen_text 内のオフセット) に変換
    pub fn grid_to_acp(&self, target_row: usize, target_col: usize) -> usize {
        let grid = self.term.grid();
        let cols = grid.columns();
        let mut acp = 0usize;

        for line_idx in 0..target_row {
            let row = &grid[Line(line_idx as i32)];
            for col in 0..cols {
                let cell = &row[Column(col)];
                if !cell.flags.contains(Flags::WIDE_CHAR_SPACER) {
                    acp += 1;
                }
            }
            acp += 1; // '\n'
        }

        // target_row の中で target_col まで
        let row = &grid[Line(target_row as i32)];
        for col in 0..target_col {
            let cell = &row[Column(col)];
            if !cell.flags.contains(Flags::WIDE_CHAR_SPACER) {
                acp += 1;
            }
        }

        acp
    }

    /// カーソル位置 (row, col) 0-indexed
    pub fn cursor_pos(&self) -> (usize, usize) {
        let point = self.term.grid().cursor.point;
        (point.line.0 as usize, point.column.0)
    }

    /// カーソルが表示状態か（DECTCEM: ESC[?25h/l）
    pub fn is_cursor_visible(&self) -> bool {
        self.term.mode().contains(TermMode::SHOW_CURSOR)
    }

    /// カーソル位置の ACP
    /// ACP は文字間の位置（カーソルが N 番目の文字の前にいる = ACP N）
    /// wide char spacer セルの場合のみ本体セルにスナップ
    pub fn cursor_acp(&self) -> usize {
        let grid = self.term.grid();
        let (row, mut col) = self.cursor_pos();

        // wide char spacer 上にいる場合、本体セルにスナップ
        if col > 0 {
            let cell = &grid[Line(row as i32)][Column(col)];
            if cell.flags.contains(Flags::WIDE_CHAR_SPACER) {
                col -= 1;
            }
        }

        self.grid_to_acp(row, col)
    }

    /// ACP をグリッド座標 (row, col) に逆変換
    pub fn acp_to_grid(&self, target_acp: usize) -> (usize, usize) {
        let grid = self.term.grid();
        let cols = grid.columns();
        let lines = grid.screen_lines();
        let mut acp = 0usize;

        for line_idx in 0..lines {
            let row = &grid[Line(line_idx as i32)];
            for col in 0..cols {
                if acp == target_acp {
                    return (line_idx, col);
                }
                let cell = &row[Column(col)];
                if !cell.flags.contains(Flags::WIDE_CHAR_SPACER) {
                    acp += 1;
                }
            }
            // '\n' のカウント
            if acp == target_acp {
                return (line_idx, cols);
            }
            acp += 1; // '\n'
        }

        // 末尾を超えた場合
        let last_line = if lines > 0 { lines - 1 } else { 0 };
        (last_line, cols)
    }

    /// 指定位置の単語の列範囲を返す (start_col, end_col)
    /// 単語境界文字: WezTerm 準拠 (スペース、タブ、括弧、引用符など)
    pub fn word_boundary(&self, row: usize, col: usize) -> (usize, usize) {
        const BOUNDARY: &str = " \t\n{[}]()'\"`,;:|<>";

        let grid = self.term.grid();
        let cols_count = grid.columns();
        let lines = grid.screen_lines();
        if row >= lines || col >= cols_count {
            return (col, col);
        }

        let display_offset = grid.display_offset() as i32;
        let row_data = &grid[Line(row as i32 - display_offset)];
        let click_cell = &row_data[Column(col)];

        // WIDE_CHAR_SPACER 上をクリックした場合、本体セルにスナップ
        let col = if click_cell.flags.contains(Flags::WIDE_CHAR_SPACER) && col > 0 {
            col - 1
        } else {
            col
        };
        let click_char = row_data[Column(col)].c;

        // クリック位置が境界文字なら、同じ境界文字の連続を選択
        let is_boundary = BOUNDARY.contains(click_char);

        // 左方向に走査
        let mut start_col = col;
        while start_col > 0 {
            let prev = start_col - 1;
            let cell = &row_data[Column(prev)];
            if cell.flags.contains(Flags::WIDE_CHAR_SPACER) {
                start_col = prev;
                continue;
            }
            if is_boundary != BOUNDARY.contains(cell.c) {
                break;
            }
            start_col = prev;
        }

        // 右方向に走査
        let mut end_col = col;
        while end_col + 1 < cols_count {
            let next = end_col + 1;
            let cell = &row_data[Column(next)];
            if cell.flags.contains(Flags::WIDE_CHAR_SPACER) {
                end_col = next;
                continue;
            }
            if is_boundary != BOUNDARY.contains(cell.c) {
                break;
            }
            end_col = next;
        }

        (start_col, end_col)
    }

    /// リサイズ
    pub fn resize(&mut self, cols: usize, rows: usize) {
        self.term.resize(TermSize { cols, rows });
    }

    pub fn screen_lines(&self) -> usize {
        self.term.grid().screen_lines()
    }

    pub fn columns(&self) -> usize {
        self.term.grid().columns()
    }

    /// 内部 Term への参照（render で Grid にアクセスするため）
    pub fn inner(&self) -> &Term<TermEventListener> {
        &self.term
    }

    /// スクロールバック表示を移動
    pub fn scroll_display(&mut self, scroll: Scroll) {
        self.term.scroll_display(scroll);
    }

    /// 現在の表示オフセット（0 = 最下部、値が大きいほど履歴方向）
    pub fn display_offset(&self) -> usize {
        self.term.grid().display_offset()
    }

    /// ALT_SCREEN モードかどうか
    pub fn is_alt_screen(&self) -> bool {
        self.term.mode().contains(TermMode::ALT_SCREEN)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// 画面行数を超える行を送り込んでスクロールバック履歴を作る
    fn fill_history(term: &mut TermWrapper, extra_lines: usize) {
        let total = term.screen_lines() + extra_lines;
        for i in 0..total {
            let line = format!("line {i}\r\n");
            term.process(line.as_bytes());
        }
    }

    #[test]
    fn test_初期状態のdisplay_offsetは0() {
        let term = TermWrapper::new(80, 24);
        assert_eq!(term.display_offset(), 0);
    }

    #[test]
    fn test_履歴がある状態で上スクロールするとdisplay_offsetが増加() {
        let mut term = TermWrapper::new(80, 24);
        fill_history(&mut term, 50);

        term.scroll_display(Scroll::Delta(3));
        assert!(term.display_offset() > 0, "上スクロール後に display_offset > 0 であるべき");
    }

    #[test]
    fn test_スクロール後に下スクロールするとdisplay_offsetが減少() {
        let mut term = TermWrapper::new(80, 24);
        fill_history(&mut term, 50);

        // まず上にスクロール
        term.scroll_display(Scroll::Delta(10));
        let offset_after_up = term.display_offset();

        // 下にスクロール
        term.scroll_display(Scroll::Delta(-3));
        assert!(term.display_offset() < offset_after_up, "下スクロール後に display_offset が減少するべき");
    }

    #[test]
    fn test_最下部でさらに下スクロールしても0のまま() {
        let mut term = TermWrapper::new(80, 24);
        fill_history(&mut term, 50);

        // 最下部（初期状態）でさらに下へ
        term.scroll_display(Scroll::Delta(-10));
        assert_eq!(term.display_offset(), 0, "最下部でさらに下スクロールしても 0 のまま");
    }

    #[test]
    fn test_is_alt_screenのデフォルトはfalse() {
        let term = TermWrapper::new(80, 24);
        assert!(!term.is_alt_screen());
    }

    #[test]
    fn test_スクロールバック時にselected_textがビューポート相対で正しいテキストを返す() {
        let mut term = TermWrapper::new(80, 24);
        // "line 0" ~ "line 73" を書き込む（24行画面 + 50行の履歴）
        fill_history(&mut term, 50);
        // 最下部の表示は "line 50" ~ "line 73"

        // 10行上にスクロール → ビューポート先頭は "line 40" になるはず
        term.scroll_display(Scroll::Delta(10));
        assert_eq!(term.display_offset(), 10);

        // ビューポート row=0 の先頭数文字を選択
        let text = term.selected_text((0, 0), (0, 6));
        assert!(
            text.starts_with("line 4"),
            "スクロールバック時 row=0 は 'line 4x' であるべきだが実際は: '{text}'"
        );
    }

    #[test]
    fn test_スクロールバック時にword_boundaryがビューポート相対で正しい境界を返す() {
        let mut term = TermWrapper::new(80, 24);
        // 履歴行: "hello world N" (50行)
        for i in 0..50 {
            term.process(format!("hello world {i}\r\n").as_bytes());
        }
        // 画面行: "hi there N" (24行)
        for i in 0..24 {
            term.process(format!("hi there {i}\r\n").as_bytes());
        }

        // display_offset=0: row=0 は "hi there X" → word_boundary(0,0) = (0, 1) "hi"
        // 10行上にスクロール → row=0 は "hello world X" → word_boundary(0,0) = (0, 4) "hello"
        term.scroll_display(Scroll::Delta(10));

        let (start, end) = term.word_boundary(0, 0);
        assert_eq!(start, 0, "単語の開始列");
        assert_eq!(end, 4, "'hello' の終了列 (0-indexed) であるべき");
    }
}
