use alacritty_terminal::event::EventListener;
use alacritty_terminal::grid::Dimensions;
use alacritty_terminal::index::{Column, Line};
use alacritty_terminal::term::cell::Flags;
use alacritty_terminal::term::Config;
use alacritty_terminal::term::Term;
use alacritty_terminal::vte::ansi::Processor;

/// alacritty_terminal にイベントを通知不要なので空実装
pub struct VoidListener;

impl EventListener for VoidListener {
    fn send_event(&self, _event: alacritty_terminal::event::Event) {}
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
    term: Term<VoidListener>,
    parser: Processor,
}

impl TermWrapper {
    pub fn new(cols: usize, rows: usize) -> Self {
        let size = TermSize { cols, rows };
        let config = Config::default();
        let term = Term::new(config, &size, VoidListener);
        Self {
            term,
            parser: Processor::new(),
        }
    }

    /// PTY から受信したバイト列を処理
    pub fn process(&mut self, bytes: &[u8]) {
        self.parser.advance(&mut self.term, bytes);
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
    pub fn inner(&self) -> &Term<VoidListener> {
        &self.term
    }
}
