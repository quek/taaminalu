use alacritty_terminal::event::EventListener;
use alacritty_terminal::grid::Dimensions;
use alacritty_terminal::index::{Column, Line};
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

    /// 指定行範囲のテキストを抽出（start_line..end_line, 0-indexed from screen top）
    pub fn grid_text(&self, start_line: usize, end_line: usize) -> String {
        let grid = self.term.grid();
        let cols = grid.columns();
        let mut text = String::new();

        for line_idx in start_line..end_line {
            let row = &grid[Line(line_idx as i32)];
            let mut line_text = String::with_capacity(cols);
            for col in 0..cols {
                let cell = &row[Column(col)];
                line_text.push(cell.c);
            }
            // 末尾の空白をトリム
            let trimmed = line_text.trim_end();
            text.push_str(trimmed);
            if line_idx + 1 < end_line {
                text.push('\n');
            }
        }
        text
    }

    /// 全画面テキスト取得
    pub fn screen_text(&self) -> String {
        self.grid_text(0, self.screen_lines())
    }

    /// カーソル位置 (row, col) 0-indexed
    pub fn cursor_pos(&self) -> (usize, usize) {
        let point = self.term.grid().cursor.point;
        (point.line.0 as usize, point.column.0)
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
