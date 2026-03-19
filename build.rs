use std::env;
use std::fs;
use std::path::Path;

fn main() {
    let out_dir = env::var("OUT_DIR").unwrap();
    let icon_path = Path::new(&out_dir).join("taaminalu.ico");

    // 32x32 32-bit RGBA アイコンを生成
    let pixels = generate_icon_32x32();
    let ico_data = build_ico(&pixels, 32, 32);
    fs::write(&icon_path, &ico_data).unwrap();

    // リソーススクリプト生成
    let rc_path = Path::new(&out_dir).join("taaminalu.rc");
    fs::write(&rc_path, format!("1 ICON \"{}\"", icon_path.to_str().unwrap().replace('\\', "/"))).unwrap();

    let _ = embed_resource::compile(&rc_path, embed_resource::NONE);
}

fn generate_icon_32x32() -> Vec<[u8; 4]> {
    // 32x32 ピクセル (RGBA)
    // ターミナルアイコン: 暗い背景に緑のプロンプト ">_"
    let bg: [u8; 4] = [0x1a, 0x1a, 0x2e, 0xFF];
    let border: [u8; 4] = [0x3B, 0x78, 0xFF, 0xFF]; // Bright Blue
    let titlebar: [u8; 4] = [0x2a, 0x2a, 0x4e, 0xFF];
    let prompt: [u8; 4] = [0x16, 0xC6, 0x0C, 0xFF]; // Green
    let cursor: [u8; 4] = [0xCC, 0xCC, 0xCC, 0xFF]; // White
    let transparent: [u8; 4] = [0, 0, 0, 0];

    let mut pixels = vec![transparent; 32 * 32];

    let set = |pixels: &mut Vec<[u8; 4]>, x: usize, y: usize, c: [u8; 4]| {
        if x < 32 && y < 32 {
            pixels[y * 32 + x] = c;
        }
    };

    // 角丸の背景 (2px 丸め)
    for y in 0..32 {
        for x in 0..32 {
            // 角丸判定
            let in_corner = (x < 2 && y < 2 && (x + y < 2))
                || (x >= 30 && y < 2 && ((31 - x) + y < 2))
                || (x < 2 && y >= 30 && (x + (31 - y) < 2))
                || (x >= 30 && y >= 30 && ((31 - x) + (31 - y) < 2));

            if !in_corner {
                let color = if y < 6 { titlebar } else { bg };
                set(&mut pixels, x, y, color);
            }
        }
    }

    // 上辺ボーダー
    for x in 2..30 {
        set(&mut pixels, x, 0, border);
    }
    for x in 1..31 {
        set(&mut pixels, x, 1, if x < 2 || x >= 30 { border } else { titlebar });
    }

    // タイトルバーのボタン (Windows 風: 右寄せ ─ □ ×)
    let btn: [u8; 4] = [0xCC, 0xCC, 0xCC, 0xFF];
    let btn_close: [u8; 4] = [0xE7, 0x48, 0x56, 0xFF];
    // 最小化 ─ (x=18..20, y=3)
    for dx in 0..3 {
        set(&mut pixels, 18 + dx, 3, btn);
    }
    // 最大化 □ (x=22..24, y=2..4)
    for dx in 0..3 {
        set(&mut pixels, 22 + dx, 2, btn);
        set(&mut pixels, 22 + dx, 4, btn);
    }
    set(&mut pixels, 22, 3, btn);
    set(&mut pixels, 24, 3, btn);
    // 閉じる × (x=26..28, y=2..4)
    set(&mut pixels, 26, 2, btn_close);
    set(&mut pixels, 28, 2, btn_close);
    set(&mut pixels, 27, 3, btn_close);
    set(&mut pixels, 26, 4, btn_close);
    set(&mut pixels, 28, 4, btn_close);

    // ">" プロンプト (row 12-19, col 4-10)
    // >  shape:
    //  ##
    //  ####
    //    ####
    //      ##
    //    ####
    //  ####
    //  ##
    let gt_pixels: [(usize, usize); 14] = [
        (5, 12), (6, 12),
        (5, 13), (6, 13), (7, 13), (8, 13),
        (7, 14), (8, 14), (9, 14), (10, 14),
        (9, 15), (10, 15),
        (7, 16), (8, 16),
    ];
    for &(px, py) in &gt_pixels {
        set(&mut pixels, px, py, prompt);
        // ミラーで下半分
        if py > 14 || py == 14 {
            continue;
        }
        let mirror_y = 14 + (14 - py);
        if mirror_y < 32 {
            set(&mut pixels, px, mirror_y, prompt);
        }
    }

    // "_" カーソル (row 18-19, col 13-18)
    for x in 14..20 {
        set(&mut pixels, x, 20, cursor);
        set(&mut pixels, x, 21, cursor);
    }

    // ブリンクカーソルブロック (col 20-21, row 12-20)
    for y in 13..21 {
        set(&mut pixels, 22, y, cursor);
        set(&mut pixels, 23, y, cursor);
    }

    // テキスト行のヒント (薄いグレーの横線)
    let hint: [u8; 4] = [0x44, 0x44, 0x55, 0xFF];
    for x in 5..18 {
        set(&mut pixels, x, 25, hint);
    }
    for x in 5..14 {
        set(&mut pixels, x, 27, hint);
    }

    pixels
}

fn build_ico(pixels: &[[u8; 4]], width: u32, height: u32) -> Vec<u8> {
    let pixel_count = (width * height) as usize;
    let and_mask_row = ((width + 31) / 32 * 4) as usize;
    let and_mask_size = and_mask_row * height as usize;
    let bmp_size = 40 + pixel_count * 4 + and_mask_size;

    let mut data = Vec::new();

    // ICO Header
    data.extend_from_slice(&0u16.to_le_bytes()); // reserved
    data.extend_from_slice(&1u16.to_le_bytes()); // type = icon
    data.extend_from_slice(&1u16.to_le_bytes()); // count = 1

    // Directory Entry
    data.push(width as u8);
    data.push(height as u8);
    data.push(0); // color count
    data.push(0); // reserved
    data.extend_from_slice(&1u16.to_le_bytes()); // planes
    data.extend_from_slice(&32u16.to_le_bytes()); // bpp
    data.extend_from_slice(&(bmp_size as u32).to_le_bytes());
    data.extend_from_slice(&22u32.to_le_bytes()); // offset

    // BITMAPINFOHEADER
    data.extend_from_slice(&40u32.to_le_bytes()); // biSize
    data.extend_from_slice(&(width as i32).to_le_bytes());
    data.extend_from_slice(&((height * 2) as i32).to_le_bytes()); // doubled for ICO
    data.extend_from_slice(&1u16.to_le_bytes()); // biPlanes
    data.extend_from_slice(&32u16.to_le_bytes()); // biBitCount
    data.extend_from_slice(&0u32.to_le_bytes()); // biCompression
    data.extend_from_slice(&0u32.to_le_bytes()); // biSizeImage
    data.extend_from_slice(&0i32.to_le_bytes()); // biXPelsPerMeter
    data.extend_from_slice(&0i32.to_le_bytes()); // biYPelsPerMeter
    data.extend_from_slice(&0u32.to_le_bytes()); // biClrUsed
    data.extend_from_slice(&0u32.to_le_bytes()); // biClrImportant

    // Pixel data (bottom-to-top, BGRA)
    for y in (0..height as usize).rev() {
        for x in 0..width as usize {
            let p = pixels[y * width as usize + x];
            data.push(p[2]); // B
            data.push(p[1]); // G
            data.push(p[0]); // R
            data.push(p[3]); // A
        }
    }

    // AND mask (all 0 for 32-bit alpha)
    for _ in 0..and_mask_size {
        data.push(0);
    }

    data
}
