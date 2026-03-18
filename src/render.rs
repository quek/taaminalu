use windows::Win32::Foundation::HWND;
use windows::Win32::Graphics::Direct2D::Common::{
    D2D1_COLOR_F, D2D_RECT_F, D2D_SIZE_U,
};
use windows::Win32::Graphics::Direct2D::{
    D2D1CreateFactory, D2D1_DRAW_TEXT_OPTIONS_NONE, D2D1_FACTORY_TYPE_SINGLE_THREADED,
    D2D1_HWND_RENDER_TARGET_PROPERTIES, D2D1_RENDER_TARGET_PROPERTIES, ID2D1Factory,
    ID2D1HwndRenderTarget,
};
use windows::Win32::Graphics::DirectWrite::{
    DWriteCreateFactory, DWRITE_FACTORY_TYPE_SHARED, DWRITE_FONT_STRETCH_NORMAL,
    DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_WEIGHT_REGULAR, DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
    DWRITE_TEXT_ALIGNMENT_CENTER, IDWriteFactory, IDWriteTextFormat,
};
use windows::Win32::Graphics::Gdi::{BeginPaint, EndPaint, PAINTSTRUCT};

use alacritty_terminal::grid::Dimensions;
use alacritty_terminal::index::{Column, Line};
use alacritty_terminal::vte::ansi::{Color, NamedColor};

use crate::tab::TabId;
use crate::term::TermWrapper;

// --- 定数 ---

const FONT_NAME: &str = "HackGen Console NF";
const FONT_SIZE: f32 = 16.0;
const TAB_FONT_SIZE: f32 = 12.0;

const TAB_BAR_HEIGHT: f32 = 30.0;
const TAB_MIN_WIDTH: f32 = 80.0;
const TAB_MAX_WIDTH: f32 = 200.0;
const TAB_PADDING: f32 = 6.0;
const TAB_CLOSE_SIZE: f32 = 16.0;
const NEW_TAB_BUTTON_WIDTH: f32 = 30.0;

/// ANSI 16色パレット (Windows Terminal "Campbell" 配色)
const ANSI_COLORS: [(u8, u8, u8); 16] = [
    (0x0C, 0x0C, 0x0C), // Black
    (0xC5, 0x0F, 0x1F), // Red
    (0x13, 0xA1, 0x0E), // Green
    (0xC1, 0x9C, 0x00), // Yellow
    (0x00, 0x37, 0xDA), // Blue
    (0x88, 0x17, 0x98), // Magenta
    (0x3A, 0x96, 0xDD), // Cyan
    (0xCC, 0xCC, 0xCC), // White
    (0x76, 0x76, 0x76), // Bright Black
    (0xE7, 0x48, 0x56), // Bright Red
    (0x16, 0xC6, 0x0C), // Bright Green
    (0xF9, 0xF1, 0xA5), // Bright Yellow
    (0x3B, 0x78, 0xFF), // Bright Blue
    (0xB4, 0x00, 0x9E), // Bright Magenta
    (0x61, 0xD6, 0xD6), // Bright Cyan
    (0xF2, 0xF2, 0xF2), // Bright White
];

const DEFAULT_FG: (u8, u8, u8) = (0xCC, 0xCC, 0xCC);
const DEFAULT_BG: (u8, u8, u8) = (0x0C, 0x0C, 0x0C);

// --- D2D カラー定数 ---

const fn rgb(r: u8, g: u8, b: u8) -> D2D1_COLOR_F {
    D2D1_COLOR_F {
        r: r as f32 / 255.0,
        g: g as f32 / 255.0,
        b: b as f32 / 255.0,
        a: 1.0,
    }
}

const BG_COLOR: D2D1_COLOR_F = rgb(0x0C, 0x0C, 0x0C);
const CURSOR_COLOR: D2D1_COLOR_F = rgb(0xCC, 0xCC, 0xCC);
const TAB_BAR_BG: D2D1_COLOR_F = rgb(0x1E, 0x1E, 0x1E);
const TAB_ACTIVE_BG: D2D1_COLOR_F = rgb(0x0C, 0x0C, 0x0C);
const TAB_INACTIVE_BG: D2D1_COLOR_F = rgb(0x2D, 0x2D, 0x2D);
const TAB_TEXT_COLOR: D2D1_COLOR_F = rgb(0xCC, 0xCC, 0xCC);
const TAB_CLOSE_COLOR: D2D1_COLOR_F = rgb(0x80, 0x80, 0x80);

// --- カラー変換 ---

fn color_to_d2d(color: &Color) -> D2D1_COLOR_F {
    let (r, g, b) = match color {
        Color::Spec(c) => (c.r, c.g, c.b),
        Color::Named(named) => named_color_rgb(named),
        Color::Indexed(idx) => indexed_color_rgb(*idx),
    };
    rgb(r, g, b)
}

fn named_color_rgb(named: &NamedColor) -> (u8, u8, u8) {
    match named {
        NamedColor::Black => ANSI_COLORS[0],
        NamedColor::Red => ANSI_COLORS[1],
        NamedColor::Green => ANSI_COLORS[2],
        NamedColor::Yellow => ANSI_COLORS[3],
        NamedColor::Blue => ANSI_COLORS[4],
        NamedColor::Magenta => ANSI_COLORS[5],
        NamedColor::Cyan => ANSI_COLORS[6],
        NamedColor::White => ANSI_COLORS[7],
        NamedColor::BrightBlack => ANSI_COLORS[8],
        NamedColor::BrightRed => ANSI_COLORS[9],
        NamedColor::BrightGreen => ANSI_COLORS[10],
        NamedColor::BrightYellow => ANSI_COLORS[11],
        NamedColor::BrightBlue => ANSI_COLORS[12],
        NamedColor::BrightMagenta => ANSI_COLORS[13],
        NamedColor::BrightCyan => ANSI_COLORS[14],
        NamedColor::BrightWhite => ANSI_COLORS[15],
        NamedColor::Foreground => DEFAULT_FG,
        NamedColor::Background => DEFAULT_BG,
        NamedColor::Cursor => (0xCC, 0xCC, 0x33),
        _ => DEFAULT_FG,
    }
}

fn indexed_color_rgb(idx: u8) -> (u8, u8, u8) {
    match idx {
        0..=15 => ANSI_COLORS[idx as usize],
        16..=231 => {
            let i = (idx - 16) as usize;
            let to_val = |c: u8| if c == 0 { 0 } else { 55 + c * 40 };
            (to_val((i / 36) as u8), to_val(((i % 36) / 6) as u8), to_val((i % 6) as u8))
        }
        232..=255 => {
            let v = 8 + (idx - 232) * 10;
            (v, v, v)
        }
    }
}

// --- タブバーヒットテスト ---

pub enum TabBarHitResult {
    Tab(usize),
    CloseTab(usize),
    NewTab,
    None,
}

// --- Renderer ---

pub struct Renderer {
    rt: ID2D1HwndRenderTarget,
    dwrite_factory: IDWriteFactory,
    text_format: IDWriteTextFormat,
    tab_text_format: IDWriteTextFormat,
    pub cell_width: f32,
    pub cell_height: f32,
}

impl Renderer {
    pub fn new(hwnd: HWND, width: u32, height: u32) -> windows::core::Result<Self> {
        let d2d_factory: ID2D1Factory =
            unsafe { D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, None)? };

        let rt = unsafe {
            d2d_factory.CreateHwndRenderTarget(
                &D2D1_RENDER_TARGET_PROPERTIES::default(),
                &D2D1_HWND_RENDER_TARGET_PROPERTIES {
                    hwnd,
                    pixelSize: D2D_SIZE_U { width, height },
                    ..Default::default()
                },
            )?
        };

        let dwrite_factory: IDWriteFactory =
            unsafe { DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED)? };

        let font_name_wide: Vec<u16> = FONT_NAME.encode_utf16().chain(std::iter::once(0)).collect();
        let locale_wide: Vec<u16> = "en-us\0".encode_utf16().collect();

        let text_format = unsafe {
            dwrite_factory.CreateTextFormat(
                windows::core::PCWSTR(font_name_wide.as_ptr()),
                None,
                DWRITE_FONT_WEIGHT_REGULAR,
                DWRITE_FONT_STYLE_NORMAL,
                DWRITE_FONT_STRETCH_NORMAL,
                FONT_SIZE,
                windows::core::PCWSTR(locale_wide.as_ptr()),
            )?
        };

        let tab_text_format = unsafe {
            let fmt = dwrite_factory.CreateTextFormat(
                windows::core::PCWSTR(font_name_wide.as_ptr()),
                None,
                DWRITE_FONT_WEIGHT_REGULAR,
                DWRITE_FONT_STYLE_NORMAL,
                DWRITE_FONT_STRETCH_NORMAL,
                TAB_FONT_SIZE,
                windows::core::PCWSTR(locale_wide.as_ptr()),
            )?;
            fmt.SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER)?;
            fmt.SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER)?;
            fmt
        };

        let (cell_width, cell_height) = Self::measure_cell(&dwrite_factory, &text_format)?;

        Ok(Self { rt, dwrite_factory, text_format, tab_text_format, cell_width, cell_height })
    }

    fn measure_cell(
        factory: &IDWriteFactory,
        format: &IDWriteTextFormat,
    ) -> windows::core::Result<(f32, f32)> {
        let text: Vec<u16> = "M".encode_utf16().collect();
        let layout = unsafe { factory.CreateTextLayout(&text, format, 1000.0, 1000.0)? };
        let mut metrics = windows::Win32::Graphics::DirectWrite::DWRITE_TEXT_METRICS::default();
        unsafe { layout.GetMetrics(&mut metrics)? };
        Ok((metrics.width, metrics.height))
    }

    pub fn resize(&self, width: u32, height: u32) {
        let _ = unsafe { self.rt.Resize(&D2D_SIZE_U { width, height }) };
    }

    pub fn tab_bar_height(&self) -> f32 {
        TAB_BAR_HEIGHT
    }

    fn pixel_size(&self) -> (u32, u32) {
        let size = unsafe { self.rt.GetPixelSize() };
        (size.width, size.height)
    }

    pub fn calc_grid_size(&self, width: u32, height: u32) -> (usize, usize) {
        let usable_height = (height as f32 - TAB_BAR_HEIGHT).max(0.0);
        let cols = (width as f32 / self.cell_width).floor() as usize;
        let rows = (usable_height / self.cell_height).floor() as usize;
        (cols.max(1), rows.max(1))
    }

    /// 現在のレンダーターゲットサイズからグリッドサイズを計算
    pub fn current_grid_size(&self) -> (usize, usize) {
        let (w, h) = self.pixel_size();
        self.calc_grid_size(w, h)
    }

    fn calc_tab_width(&self, tab_count: usize) -> f32 {
        let (width, _) = self.pixel_size();
        let available = width as f32 - NEW_TAB_BUTTON_WIDTH;
        let per_tab = available / tab_count.max(1) as f32;
        per_tab.clamp(TAB_MIN_WIDTH, TAB_MAX_WIDTH)
    }

    // --- ヒットテスト ---

    pub fn hit_test_tab_bar(&self, x: f32, y: f32, tab_count: usize) -> TabBarHitResult {
        if y >= TAB_BAR_HEIGHT {
            return TabBarHitResult::None;
        }

        let tab_width = self.calc_tab_width(tab_count);
        let tabs_total_width = tab_width * tab_count as f32;

        // ＋ボタン
        if x >= tabs_total_width && x < tabs_total_width + NEW_TAB_BUTTON_WIDTH {
            return TabBarHitResult::NewTab;
        }

        // タブエリア
        if x < tabs_total_width {
            let tab_index = (x / tab_width) as usize;
            if tab_index < tab_count {
                // ×ボタン
                let tab_right = (tab_index + 1) as f32 * tab_width;
                let close_right = tab_right - TAB_PADDING;
                let close_left = close_right - TAB_CLOSE_SIZE;
                let close_top = (TAB_BAR_HEIGHT - TAB_CLOSE_SIZE) / 2.0;
                let close_bottom = close_top + TAB_CLOSE_SIZE;

                if x >= close_left && x <= close_right && y >= close_top && y <= close_bottom {
                    return TabBarHitResult::CloseTab(tab_index);
                }
                return TabBarHitResult::Tab(tab_index);
            }
        }

        TabBarHitResult::None
    }

    // --- 描画 ---

    pub fn paint_with_tabs(
        &self,
        hwnd: HWND,
        term: &TermWrapper,
        tabs: &[(&str, TabId)],
        active_index: usize,
        preedit: &str,
    ) {
        let mut ps = PAINTSTRUCT::default();
        unsafe { let _ = BeginPaint(hwnd, &mut ps); }

        unsafe {
            self.rt.BeginDraw();
            self.rt.Clear(Some(&BG_COLOR));
            self.draw_tab_bar(tabs, active_index);
            self.draw_grid(term);
            if !preedit.is_empty() {
                let (cursor_row, cursor_col) = term.cursor_pos();
                self.draw_preedit(preedit, cursor_row, cursor_col);
            }
            let _ = self.rt.EndDraw(None, None);
        }

        unsafe { let _ = EndPaint(hwnd, &ps); }
    }

    /// IME preedit（変換中テキスト）をカーソル位置にインライン描画
    fn draw_preedit(&self, preedit: &str, cursor_row: usize, cursor_col: usize) {
        let x = cursor_col as f32 * self.cell_width;
        let y = TAB_BAR_HEIGHT + cursor_row as f32 * self.cell_height;

        // preedit の表示幅を計算（全角文字は2セル）
        let display_width: usize = preedit.chars().map(|c| {
            if c.is_ascii() { 1 } else { 2 }
        }).sum();
        let preedit_pixel_width = display_width as f32 * self.cell_width;

        unsafe {
            // 背景
            let bg = D2D1_COLOR_F { r: 0.2, g: 0.2, b: 0.3, a: 1.0 };
            self.fill_rect(x, y, x + preedit_pixel_width, y + self.cell_height, &bg);

            // テキスト
            let fg = D2D1_COLOR_F { r: 1.0, g: 1.0, b: 1.0, a: 1.0 };
            let wide: Vec<u16> = preedit.encode_utf16().collect();
            if let Ok(layout) = self.dwrite_factory.CreateTextLayout(
                &wide, &self.text_format, preedit_pixel_width, self.cell_height,
            ) {
                if let Ok(brush) = self.rt.CreateSolidColorBrush(&fg, None) {
                    self.rt.DrawTextLayout(
                        windows_numerics::Vector2 { X: x, Y: y },
                        &layout, &brush, D2D1_DRAW_TEXT_OPTIONS_NONE,
                    );
                }
            }

            // 下線
            let underline_y = y + self.cell_height - 1.0;
            let underline_color = D2D1_COLOR_F { r: 0.8, g: 0.8, b: 0.2, a: 1.0 };
            self.fill_rect(x, underline_y, x + preedit_pixel_width, underline_y + 1.0, &underline_color);
        }
    }

    fn draw_tab_bar(&self, tabs: &[(&str, TabId)], active_index: usize) {
        unsafe {
            let (width, _) = self.pixel_size();
            let tab_width = self.calc_tab_width(tabs.len());

            // タブバー背景
            self.fill_rect(0.0, 0.0, width as f32, TAB_BAR_HEIGHT, &TAB_BAR_BG);

            // 各タブ
            for (i, (title, _)) in tabs.iter().enumerate() {
                let x = i as f32 * tab_width;
                let bg = if i == active_index { TAB_ACTIVE_BG } else { TAB_INACTIVE_BG };

                self.fill_rect(x + 1.0, 2.0, x + tab_width - 1.0, TAB_BAR_HEIGHT, &bg);

                // タブタイトル
                let title_right = x + tab_width - TAB_PADDING - TAB_CLOSE_SIZE - 4.0;
                self.draw_text(title, &self.tab_text_format, x + TAB_PADDING, 0.0, title_right - x - TAB_PADDING, TAB_BAR_HEIGHT, &TAB_TEXT_COLOR);

                // ×ボタン
                let close_left = x + tab_width - TAB_PADDING - TAB_CLOSE_SIZE;
                let close_top = (TAB_BAR_HEIGHT - TAB_CLOSE_SIZE) / 2.0;
                self.draw_text("×", &self.tab_text_format, close_left, close_top, TAB_CLOSE_SIZE, TAB_CLOSE_SIZE, &TAB_CLOSE_COLOR);
            }

            // ＋ボタン
            let plus_x = tabs.len() as f32 * tab_width;
            self.draw_text("+", &self.tab_text_format, plus_x, 0.0, NEW_TAB_BUTTON_WIDTH, TAB_BAR_HEIGHT, &TAB_TEXT_COLOR);
        }
    }

    fn draw_grid(&self, term: &TermWrapper) {
        unsafe {
            let grid = term.inner().grid();
            let cols = grid.columns();
            let lines = grid.screen_lines();
            let (cursor_row, cursor_col) = term.cursor_pos();

            for line_idx in 0..lines {
                let row = &grid[Line(line_idx as i32)];
                let y = TAB_BAR_HEIGHT + line_idx as f32 * self.cell_height;

                for col_idx in 0..cols {
                    let cell = &row[Column(col_idx)];
                    let x = col_idx as f32 * self.cell_width;
                    let c = cell.c;
                    let is_cursor = line_idx == cursor_row && col_idx == cursor_col;

                    // セル背景色
                    let cell_bg = color_to_d2d(&cell.bg);
                    let has_bg = cell_bg.r != BG_COLOR.r || cell_bg.g != BG_COLOR.g || cell_bg.b != BG_COLOR.b;

                    if has_bg || is_cursor {
                        let bg = if is_cursor { &CURSOR_COLOR } else { &cell_bg };
                        self.fill_rect(x, y, x + self.cell_width, y + self.cell_height, bg);
                    }

                    // テキスト
                    if c != ' ' && c != '\0' {
                        let fg = if is_cursor { BG_COLOR } else { color_to_d2d(&cell.fg) };
                        let text: Vec<u16> = [c as u16].to_vec();
                        if let Ok(layout) = self.dwrite_factory.CreateTextLayout(
                            &text, &self.text_format, self.cell_width, self.cell_height,
                        ) {
                            if let Ok(brush) = self.rt.CreateSolidColorBrush(&fg, None) {
                                self.rt.DrawTextLayout(
                                    windows_numerics::Vector2 { X: x, Y: y },
                                    &layout, &brush, D2D1_DRAW_TEXT_OPTIONS_NONE,
                                );
                            }
                        }
                    }
                }
            }
        }
    }

    // --- 描画ヘルパー ---

    /// 矩形を塗りつぶし
    unsafe fn fill_rect(&self, left: f32, top: f32, right: f32, bottom: f32, color: &D2D1_COLOR_F) {
        unsafe {
            if let Ok(brush) = self.rt.CreateSolidColorBrush(color, None) {
                self.rt.FillRectangle(&D2D_RECT_F { left, top, right, bottom }, &brush);
            }
        }
    }

    /// テキストを描画
    unsafe fn draw_text(&self, text: &str, format: &IDWriteTextFormat, x: f32, y: f32, width: f32, height: f32, color: &D2D1_COLOR_F) {
        unsafe {
            let wide: Vec<u16> = text.encode_utf16().collect();
            if let Ok(layout) = self.dwrite_factory.CreateTextLayout(&wide, format, width, height) {
                if let Ok(brush) = self.rt.CreateSolidColorBrush(color, None) {
                    self.rt.DrawTextLayout(
                        windows_numerics::Vector2 { X: x, Y: y },
                        &layout, &brush, D2D1_DRAW_TEXT_OPTIONS_NONE,
                    );
                }
            }
        }
    }
}
