use alacritty_terminal::event::{Event, EventListener};
use alacritty_terminal::grid::Dimensions;
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

        // start/end を正規化
        let (start, end) = if start.0 < end.0 || (start.0 == end.0 && start.1 <= end.1) {
            (start, end)
        } else {
            (end, start)
        };

        let mut result = String::new();
        for row_idx in start.0..=end.0.min(lines.saturating_sub(1)) {
            let row = &grid[Line(row_idx as i32)];
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

}
