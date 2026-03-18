use windows::Win32::Foundation::HWND;
use windows::Win32::Graphics::Direct2D::Common::{
    D2D1_COLOR_F, D2D_RECT_F, D2D_SIZE_U,
};
use windows::Win32::Graphics::Direct2D::{
    D2D1CreateFactory, D2D1_FACTORY_TYPE_SINGLE_THREADED, D2D1_HWND_RENDER_TARGET_PROPERTIES,
    D2D1_RENDER_TARGET_PROPERTIES, ID2D1Factory, ID2D1HwndRenderTarget, ID2D1SolidColorBrush,
};
use windows::Win32::Graphics::DirectWrite::{
    DWriteCreateFactory, DWRITE_FACTORY_TYPE_SHARED, DWRITE_FONT_STRETCH_NORMAL,
    DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_WEIGHT_REGULAR, IDWriteFactory, IDWriteTextFormat,
};
use windows::Win32::Graphics::Gdi::{BeginPaint, EndPaint, PAINTSTRUCT};

use alacritty_terminal::grid::Dimensions;
use alacritty_terminal::index::{Column, Line};

use crate::term::TermWrapper;

const FONT_NAME: &str = "Consolas";
const FONT_SIZE: f32 = 16.0;
const BG_COLOR: D2D1_COLOR_F = D2D1_COLOR_F {
    r: 0.1,
    g: 0.1,
    b: 0.12,
    a: 1.0,
};
const FG_COLOR: D2D1_COLOR_F = D2D1_COLOR_F {
    r: 0.9,
    g: 0.9,
    b: 0.9,
    a: 1.0,
};
const CURSOR_COLOR: D2D1_COLOR_F = D2D1_COLOR_F {
    r: 0.8,
    g: 0.8,
    b: 0.2,
    a: 0.8,
};

pub struct Renderer {
    rt: ID2D1HwndRenderTarget,
    dwrite_factory: IDWriteFactory,
    text_format: IDWriteTextFormat,
    fg_brush: ID2D1SolidColorBrush,
    _bg_brush: ID2D1SolidColorBrush,
    cursor_brush: ID2D1SolidColorBrush,
    pub cell_width: f32,
    pub cell_height: f32,
}

impl Renderer {
    pub fn new(hwnd: HWND, width: u32, height: u32) -> windows::core::Result<Self> {
        let d2d_factory: ID2D1Factory =
            unsafe { D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, None)? };

        let rt_props = D2D1_RENDER_TARGET_PROPERTIES::default();
        let hwnd_props = D2D1_HWND_RENDER_TARGET_PROPERTIES {
            hwnd,
            pixelSize: D2D_SIZE_U { width, height },
            ..Default::default()
        };

        let rt = unsafe { d2d_factory.CreateHwndRenderTarget(&rt_props, &hwnd_props)? };

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

        let fg_brush = unsafe { rt.CreateSolidColorBrush(&FG_COLOR, None)? };
        let _bg_brush = unsafe { rt.CreateSolidColorBrush(&BG_COLOR, None)? };
        let cursor_brush = unsafe { rt.CreateSolidColorBrush(&CURSOR_COLOR, None)? };

        // セルサイズ計算: "M" のレイアウトでメトリクス取得
        let (cell_width, cell_height) = Self::measure_cell(&dwrite_factory, &text_format)?;

        Ok(Self {
            rt,
            dwrite_factory,
            text_format,
            fg_brush,
            _bg_brush,
            cursor_brush,
            cell_width,
            cell_height,
        })
    }

    fn measure_cell(
        factory: &IDWriteFactory,
        format: &IDWriteTextFormat,
    ) -> windows::core::Result<(f32, f32)> {
        let text: Vec<u16> = "M".encode_utf16().collect();
        let layout = unsafe {
            factory.CreateTextLayout(
                &text,
                format,
                1000.0,
                1000.0,
            )?
        };
        let mut metrics = windows::Win32::Graphics::DirectWrite::DWRITE_TEXT_METRICS::default();
        unsafe { layout.GetMetrics(&mut metrics)? };
        Ok((metrics.width, metrics.height))
    }

    pub fn resize(&self, width: u32, height: u32) {
        let size = D2D_SIZE_U { width, height };
        let _ = unsafe { self.rt.Resize(&size) };
    }

    /// ウィンドウサイズからターミナルのカラム数・行数を計算
    pub fn calc_grid_size(&self, width: u32, height: u32) -> (usize, usize) {
        let cols = (width as f32 / self.cell_width).floor() as usize;
        let rows = (height as f32 / self.cell_height).floor() as usize;
        (cols.max(1), rows.max(1))
    }

    pub fn paint(&self, hwnd: HWND, term: &TermWrapper) {
        let mut ps = PAINTSTRUCT::default();
        unsafe {
            let _ = BeginPaint(hwnd, &mut ps);
        }

        unsafe {
            self.rt.BeginDraw();
            self.rt.Clear(Some(&BG_COLOR));

            let grid = term.inner().grid();
            let cols = grid.columns();
            let lines = grid.screen_lines();
            let (cursor_row, cursor_col) = term.cursor_pos();

            for line_idx in 0..lines {
                let row = &grid[Line(line_idx as i32)];
                let y = line_idx as f32 * self.cell_height;

                for col_idx in 0..cols {
                    let cell = &row[Column(col_idx)];
                    let x = col_idx as f32 * self.cell_width;
                    let c = cell.c;

                    if c == ' ' && !(line_idx == cursor_row && col_idx == cursor_col) {
                        continue;
                    }

                    // カーソル位置に背景描画
                    if line_idx == cursor_row && col_idx == cursor_col {
                        let rect = D2D_RECT_F {
                            left: x,
                            top: y,
                            right: x + self.cell_width,
                            bottom: y + self.cell_height,
                        };
                        self.rt.FillRectangle(&rect, &self.cursor_brush);
                    }

                    if c != ' ' && c != '\0' {
                        let text: Vec<u16> = [c as u16].to_vec();
                        let layout = self.dwrite_factory.CreateTextLayout(
                            &text,
                            &self.text_format,
                            self.cell_width,
                            self.cell_height,
                        );
                        if let Ok(layout) = layout {
                            let origin = windows_numerics::Vector2 { X: x, Y: y };
                            self.rt.DrawTextLayout(
                                origin,
                                &layout,
                                &self.fg_brush,
                                windows::Win32::Graphics::Direct2D::D2D1_DRAW_TEXT_OPTIONS_NONE,
                            );
                        }
                    }
                }
            }

            let _ = self.rt.EndDraw(None, None);
        }

        unsafe {
            let _ = EndPaint(hwnd, &ps);
        }
    }
}
