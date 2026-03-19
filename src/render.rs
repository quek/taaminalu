use std::cell::RefCell;
use std::collections::HashMap;

use windows::Win32::Foundation::HWND;
use windows::Win32::Graphics::Direct2D::Common::{
    D2D1_COLOR_F, D2D1_FIGURE_BEGIN_HOLLOW, D2D1_FIGURE_END_OPEN, D2D_RECT_F, D2D_SIZE_U,
};
use windows::Win32::Graphics::Direct2D::{
    D2D1CreateFactory, D2D1_DASH_STYLE_DASH, D2D1_DASH_STYLE_DOT, D2D1_DRAW_TEXT_OPTIONS_NONE,
    D2D1_CAP_STYLE_ROUND, D2D1_FACTORY_TYPE_SINGLE_THREADED,
    D2D1_HWND_RENDER_TARGET_PROPERTIES, D2D1_QUADRATIC_BEZIER_SEGMENT,
    D2D1_RENDER_TARGET_PROPERTIES, D2D1_STROKE_STYLE_PROPERTIES, ID2D1Factory,
    ID2D1HwndRenderTarget, ID2D1SolidColorBrush, ID2D1StrokeStyle,
};
use windows::Win32::Graphics::DirectWrite::{
    DWriteCreateFactory, DWRITE_FACTORY_TYPE_SHARED, DWRITE_FONT_STRETCH_NORMAL,
    DWRITE_FONT_STYLE_ITALIC, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_WEIGHT_BOLD,
    DWRITE_FONT_WEIGHT_REGULAR, DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
    DWRITE_TEXT_ALIGNMENT_CENTER, IDWriteFactory, IDWriteTextFormat,
};
use windows::Win32::Graphics::Gdi::{BeginPaint, EndPaint, PAINTSTRUCT};

use alacritty_terminal::grid::Dimensions;
use alacritty_terminal::index::{Column, Line};
use alacritty_terminal::term::cell::Flags;
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
const SELECTION_COLOR: D2D1_COLOR_F = rgb(0x26, 0x4F, 0x78);
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

/// BOLD 時に Named color (0-7) を Bright variant (8-15) に変換
fn bold_color(color: &Color) -> D2D1_COLOR_F {
    match color {
        Color::Named(named) => {
            let bright = match named {
                NamedColor::Black => NamedColor::BrightBlack,
                NamedColor::Red => NamedColor::BrightRed,
                NamedColor::Green => NamedColor::BrightGreen,
                NamedColor::Yellow => NamedColor::BrightYellow,
                NamedColor::Blue => NamedColor::BrightBlue,
                NamedColor::Magenta => NamedColor::BrightMagenta,
                NamedColor::Cyan => NamedColor::BrightCyan,
                NamedColor::White => NamedColor::BrightWhite,
                other => *other,
            };
            let (r, g, b) = named_color_rgb(&bright);
            rgb(r, g, b)
        }
        _ => color_to_d2d(color),
    }
}

/// DIM: 前景色の輝度を半分にする
fn dim_color(color: &D2D1_COLOR_F) -> D2D1_COLOR_F {
    D2D1_COLOR_F {
        r: color.r * 0.5,
        g: color.g * 0.5,
        b: color.b * 0.5,
        a: color.a,
    }
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
    d2d_factory: ID2D1Factory,
    dwrite_factory: IDWriteFactory,
    text_format: IDWriteTextFormat,
    bold_text_format: IDWriteTextFormat,
    italic_text_format: IDWriteTextFormat,
    bold_italic_text_format: IDWriteTextFormat,
    tab_text_format: IDWriteTextFormat,
    dotted_stroke: ID2D1StrokeStyle,
    dashed_stroke: ID2D1StrokeStyle,
    pub cell_width: f32,
    pub cell_height: f32,
    brush_cache: RefCell<HashMap<u32, ID2D1SolidColorBrush>>,
}

/// DirectWrite TextFormat を作成するヘルパー
fn create_text_format(
    factory: &IDWriteFactory,
    font: &[u16],
    locale: &[u16],
    size: f32,
    weight: windows::Win32::Graphics::DirectWrite::DWRITE_FONT_WEIGHT,
    style: windows::Win32::Graphics::DirectWrite::DWRITE_FONT_STYLE,
) -> windows::core::Result<IDWriteTextFormat> {
    unsafe {
        factory.CreateTextFormat(
            windows::core::PCWSTR(font.as_ptr()),
            None,
            weight,
            style,
            DWRITE_FONT_STRETCH_NORMAL,
            size,
            windows::core::PCWSTR(locale.as_ptr()),
        )
    }
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

        let font_wide: Vec<u16> = FONT_NAME.encode_utf16().chain(std::iter::once(0)).collect();
        let locale_wide: Vec<u16> = "en-us\0".encode_utf16().collect();

        let text_format = create_text_format(
            &dwrite_factory, &font_wide, &locale_wide,
            FONT_SIZE, DWRITE_FONT_WEIGHT_REGULAR, DWRITE_FONT_STYLE_NORMAL,
        )?;
        let bold_text_format = create_text_format(
            &dwrite_factory, &font_wide, &locale_wide,
            FONT_SIZE, DWRITE_FONT_WEIGHT_BOLD, DWRITE_FONT_STYLE_NORMAL,
        )?;
        let italic_text_format = create_text_format(
            &dwrite_factory, &font_wide, &locale_wide,
            FONT_SIZE, DWRITE_FONT_WEIGHT_REGULAR, DWRITE_FONT_STYLE_ITALIC,
        )?;
        let bold_italic_text_format = create_text_format(
            &dwrite_factory, &font_wide, &locale_wide,
            FONT_SIZE, DWRITE_FONT_WEIGHT_BOLD, DWRITE_FONT_STYLE_ITALIC,
        )?;
        let tab_text_format = {
            let fmt = create_text_format(
                &dwrite_factory, &font_wide, &locale_wide,
                TAB_FONT_SIZE, DWRITE_FONT_WEIGHT_REGULAR, DWRITE_FONT_STYLE_NORMAL,
            )?;
            unsafe {
                fmt.SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER)?;
                fmt.SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER)?;
            }
            fmt
        };

        let dotted_stroke = unsafe {
            d2d_factory.CreateStrokeStyle(
                &D2D1_STROKE_STYLE_PROPERTIES {
                    dashStyle: D2D1_DASH_STYLE_DOT,
                    dashCap: D2D1_CAP_STYLE_ROUND,
                    ..Default::default()
                },
                None,
            )?
        };
        let dashed_stroke = unsafe {
            d2d_factory.CreateStrokeStyle(
                &D2D1_STROKE_STYLE_PROPERTIES {
                    dashStyle: D2D1_DASH_STYLE_DASH,
                    ..Default::default()
                },
                None,
            )?
        };

        let (cell_width, cell_height) = Self::measure_cell(&dwrite_factory, &text_format)?;

        Ok(Self {
            rt, d2d_factory, dwrite_factory,
            text_format, bold_text_format, italic_text_format, bold_italic_text_format,
            tab_text_format, dotted_stroke, dashed_stroke, cell_width, cell_height,
            brush_cache: RefCell::new(HashMap::new()),
        })
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
        self.brush_cache.borrow_mut().clear();
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
        selection: Option<&crate::app::Selection>,
    ) {
        let mut ps = PAINTSTRUCT::default();
        unsafe { let _ = BeginPaint(hwnd, &mut ps); }

        unsafe {
            self.rt.BeginDraw();
            self.rt.Clear(Some(&BG_COLOR));
            self.draw_tab_bar(tabs, active_index);
            self.draw_grid(term, selection);
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
                if let Some(brush) = self.get_brush(&fg) {
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

    fn draw_grid(&self, term: &TermWrapper, selection: Option<&crate::app::Selection>) {
        unsafe {
            let grid = term.inner().grid();
            let cols = grid.columns();
            let lines = grid.screen_lines();
            let cursor_visible = term.is_cursor_visible();
            let (cursor_row, cursor_col) = term.cursor_pos();

            for line_idx in 0..lines {
                let row = &grid[Line(line_idx as i32)];
                let y = (TAB_BAR_HEIGHT + line_idx as f32 * self.cell_height).floor();

                for col_idx in 0..cols {
                    let cell = &row[Column(col_idx)];
                    let x = (col_idx as f32 * self.cell_width).floor();
                    let c = cell.c;
                    let is_cursor =
                        cursor_visible && line_idx == cursor_row && col_idx == cursor_col;

                    // wide char spacer はスキップ（本体セルで描画済み）
                    if cell.flags.contains(Flags::WIDE_CHAR_SPACER) {
                        continue;
                    }

                    let is_wide = cell.flags.contains(Flags::WIDE_CHAR);
                    // 隣接セルとのサブピクセル隙間を防ぐため、右端もスナップ
                    let cell_w = if is_wide {
                        ((col_idx + 2) as f32 * self.cell_width).floor() - x
                    } else {
                        ((col_idx + 1) as f32 * self.cell_width).floor() - x
                    };

                    // セル属性フラグ
                    let flags = cell.flags;
                    let is_inverse = flags.contains(Flags::INVERSE);
                    let is_bold = flags.contains(Flags::BOLD);
                    let is_italic = flags.contains(Flags::ITALIC);
                    let is_dim = flags.contains(Flags::DIM);
                    let is_hidden = flags.contains(Flags::HIDDEN);
                    let is_strikeout = flags.contains(Flags::STRIKEOUT);

                    // INVERSE: 前景色と背景色を入れ替え
                    // BOLD: Named color (0-7) を Bright variant (8-15) に変換
                    let (mut cell_fg, cell_bg) = if is_inverse {
                        let fg = if is_bold { bold_color(&cell.bg) } else { color_to_d2d(&cell.bg) };
                        (fg, color_to_d2d(&cell.fg))
                    } else {
                        let fg = if is_bold { bold_color(&cell.fg) } else { color_to_d2d(&cell.fg) };
                        (fg, color_to_d2d(&cell.bg))
                    };

                    // DIM: 前景色の輝度を半分にする
                    if is_dim {
                        cell_fg = dim_color(&cell_fg);
                    }

                    // 選択範囲内かチェック
                    let is_selected = selection.is_some_and(|s| s.contains(line_idx, col_idx));

                    // セル背景色
                    let has_bg = cell_bg.r != BG_COLOR.r || cell_bg.g != BG_COLOR.g || cell_bg.b != BG_COLOR.b;

                    if has_bg || is_cursor || is_selected {
                        let bg = if is_cursor {
                            &CURSOR_COLOR
                        } else if is_selected {
                            &SELECTION_COLOR
                        } else {
                            &cell_bg
                        };
                        self.fill_rect(x, y, x + cell_w, y + self.cell_height, bg);
                    }

                    // HIDDEN: テキストを描画しない
                    // テキスト描画
                    if !is_hidden && c != ' ' && c != '\0' {
                        let fg = if is_cursor { BG_COLOR } else { cell_fg };

                        // BOLD/ITALIC に応じた TextFormat を選択
                        let format = match (is_bold, is_italic) {
                            (true, true) => &self.bold_italic_text_format,
                            (true, false) => &self.bold_text_format,
                            (false, true) => &self.italic_text_format,
                            (false, false) => &self.text_format,
                        };

                        let text = [c as u16];
                        if let Ok(layout) = self.dwrite_factory.CreateTextLayout(
                            &text, format, cell_w, self.cell_height,
                        ) {
                            if let Some(brush) = self.get_brush(&fg) {
                                self.rt.DrawTextLayout(
                                    windows_numerics::Vector2 { X: x, Y: y },
                                    &layout, &brush, D2D1_DRAW_TEXT_OPTIONS_NONE,
                                );
                            }
                        }
                    }

                    // 下線色: cell.underline_color() があれば使用、なければ前景色
                    let ul_color = if is_cursor {
                        BG_COLOR
                    } else if let Some(uc) = cell.underline_color() {
                        color_to_d2d(&uc)
                    } else {
                        cell_fg
                    };

                    // UNDERLINE 各種: セル下部に線を描画
                    if flags.contains(Flags::UNDERCURL) {
                        self.draw_undercurl(x, y + self.cell_height - 2.0, cell_w, &ul_color);
                    } else if flags.contains(Flags::DOTTED_UNDERLINE) {
                        self.draw_styled_line(x, y + self.cell_height - 0.5, cell_w, &ul_color, &self.dotted_stroke);
                    } else if flags.contains(Flags::DASHED_UNDERLINE) {
                        self.draw_styled_line(x, y + self.cell_height - 0.5, cell_w, &ul_color, &self.dashed_stroke);
                    } else if flags.contains(Flags::DOUBLE_UNDERLINE) {
                        let line_y = y + self.cell_height - 1.0;
                        self.fill_rect(x, line_y, x + cell_w, line_y + 1.0, &ul_color);
                        self.fill_rect(x, line_y - 2.0, x + cell_w, line_y - 1.0, &ul_color);
                    } else if flags.contains(Flags::UNDERLINE) {
                        let line_y = y + self.cell_height - 1.0;
                        self.fill_rect(x, line_y, x + cell_w, line_y + 1.0, &ul_color);
                    }

                    // STRIKEOUT: セル中央に取り消し線
                    if is_strikeout {
                        let strike_color = if is_cursor { BG_COLOR } else { cell_fg };
                        self.fill_rect(x, y + self.cell_height * 0.5, x + cell_w, y + self.cell_height * 0.5 + 1.0, &strike_color);
                    }
                }
            }
        }
    }

    // --- 描画ヘルパー ---

    /// 色からキャッシュ済み Brush を取得（なければ生成してキャッシュ）
    fn get_brush(&self, color: &D2D1_COLOR_F) -> Option<ID2D1SolidColorBrush> {
        let key = ((color.r * 255.0) as u32) << 24
                | ((color.g * 255.0) as u32) << 16
                | ((color.b * 255.0) as u32) << 8
                | ((color.a * 255.0) as u32);
        let mut cache = self.brush_cache.borrow_mut();
        if let Some(brush) = cache.get(&key) {
            return Some(brush.clone());
        }
        let brush = unsafe { self.rt.CreateSolidColorBrush(color, None).ok()? };
        cache.insert(key, brush.clone());
        Some(brush)
    }

    /// 波線（undercurl）を描画
    unsafe fn draw_undercurl(&self, x: f32, baseline_y: f32, width: f32, color: &D2D1_COLOR_F) {
        unsafe {
            let Some(brush) = self.get_brush(color) else { return };
            let Ok(geometry) = self.d2d_factory.CreatePathGeometry() else { return };
            let Ok(sink) = geometry.Open() else { return };

            let amplitude = 2.0f32;
            let half_period = self.cell_width / 2.0;

            sink.BeginFigure(
                windows_numerics::Vector2 { X: x, Y: baseline_y },
                D2D1_FIGURE_BEGIN_HOLLOW,
            );

            let mut cx = x;
            let mut going_down = true;
            while cx < x + width {
                let next_x = (cx + half_period).min(x + width);
                let ctrl_y = if going_down {
                    baseline_y + amplitude
                } else {
                    baseline_y - amplitude
                };
                sink.AddQuadraticBezier(&D2D1_QUADRATIC_BEZIER_SEGMENT {
                    point1: windows_numerics::Vector2 {
                        X: (cx + next_x) / 2.0,
                        Y: ctrl_y,
                    },
                    point2: windows_numerics::Vector2 {
                        X: next_x,
                        Y: baseline_y,
                    },
                });
                cx = next_x;
                going_down = !going_down;
            }

            sink.EndFigure(D2D1_FIGURE_END_OPEN);
            let _ = sink.Close();

            self.rt
                .DrawGeometry(&geometry, &brush, 1.0, None::<&ID2D1StrokeStyle>);
        }
    }

    /// スタイル付き線（点線/破線）を描画
    unsafe fn draw_styled_line(
        &self,
        x: f32,
        y: f32,
        width: f32,
        color: &D2D1_COLOR_F,
        style: &ID2D1StrokeStyle,
    ) {
        unsafe {
            if let Some(brush) = self.get_brush(color) {
                self.rt.DrawLine(
                    windows_numerics::Vector2 { X: x, Y: y },
                    windows_numerics::Vector2 {
                        X: x + width,
                        Y: y,
                    },
                    &brush,
                    1.0,
                    style,
                );
            }
        }
    }

    /// 矩形を塗りつぶし
    unsafe fn fill_rect(&self, left: f32, top: f32, right: f32, bottom: f32, color: &D2D1_COLOR_F) {
        unsafe {
            if let Some(brush) = self.get_brush(color) {
                self.rt.FillRectangle(&D2D_RECT_F { left, top, right, bottom }, &brush);
            }
        }
    }

    /// テキストを描画
    unsafe fn draw_text(&self, text: &str, format: &IDWriteTextFormat, x: f32, y: f32, width: f32, height: f32, color: &D2D1_COLOR_F) {
        unsafe {
            let wide: Vec<u16> = text.encode_utf16().collect();
            if let Ok(layout) = self.dwrite_factory.CreateTextLayout(&wide, format, width, height) {
                if let Some(brush) = self.get_brush(color) {
                    self.rt.DrawTextLayout(
                        windows_numerics::Vector2 { X: x, Y: y },
                        &layout, &brush, D2D1_DRAW_TEXT_OPTIONS_NONE,
                    );
                }
            }
        }
    }
}
