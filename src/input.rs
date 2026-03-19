use windows::Win32::UI::Input::KeyboardAndMouse::*;

// --- 修飾キー状態 ---

pub struct Modifiers {
    pub shift: bool,
    pub alt: bool,
    pub ctrl: bool,
}

pub fn get_modifiers() -> Modifiers {
    unsafe {
        Modifiers {
            shift: GetKeyState(VK_SHIFT.0 as i32) < 0,
            alt: GetKeyState(VK_MENU.0 as i32) < 0,
            ctrl: GetKeyState(VK_CONTROL.0 as i32) < 0,
        }
    }
}

/// xterm 修飾キーパラメータ: 1 + (Shift=1 | Alt=2 | Ctrl=4)
fn modifier_param(mods: &Modifiers) -> u8 {
    let mut p = 0u8;
    if mods.shift { p |= 1; }
    if mods.alt { p |= 2; }
    if mods.ctrl { p |= 4; }
    1 + p
}

fn has_modifiers(mods: &Modifiers) -> bool {
    mods.shift || mods.alt || mods.ctrl
}

// --- キーシーケンス生成 ---

/// 特殊キー → VT エスケープシーケンス (修飾キー対応)
pub fn build_key_sequence(vk: VIRTUAL_KEY, mods: &Modifiers) -> Option<Vec<u8>> {
    // Backspace: 修飾キー対応
    if vk == VK_BACK {
        let mut seq = Vec::new();
        if mods.alt { seq.push(0x1b); }
        if mods.ctrl {
            seq.push(0x08); // Ctrl+Backspace = BS
        } else {
            seq.push(0x7f); // Backspace = DEL
        }
        return Some(seq);
    }

    // CSI キー: 矢印、Home/End、Insert/Delete、PageUp/Down
    if let Some((code, suffix)) = csi_key_params(vk) {
        let mp = modifier_param(mods);
        let seq = if mp > 1 {
            // 修飾キーあり: \x1b[1;{mod}{suffix} or \x1b[{code};{mod}~
            if suffix == b'~' {
                format!("\x1b[{};{}~", code, mp).into_bytes()
            } else {
                format!("\x1b[1;{}{}", mp, suffix as char).into_bytes()
            }
        } else {
            // 修飾キーなし
            if suffix == b'~' {
                format!("\x1b[{}~", code).into_bytes()
            } else {
                vec![0x1b, b'[', suffix]
            }
        };
        return Some(seq);
    }

    // ファンクションキー F1-F12
    if let Some(seq) = function_key_sequence(vk, mods) {
        return Some(seq);
    }

    None
}

/// CSI キーのパラメータ: (数値コード, サフィックス文字)
/// サフィックスが '~' の場合は \x1b[{code}~ 形式
/// それ以外は \x1b[{suffix} 形式
fn csi_key_params(vk: VIRTUAL_KEY) -> Option<(u8, u8)> {
    match vk {
        VK_UP => Some((1, b'A')),
        VK_DOWN => Some((1, b'B')),
        VK_RIGHT => Some((1, b'C')),
        VK_LEFT => Some((1, b'D')),
        VK_HOME => Some((1, b'H')),
        VK_END => Some((1, b'F')),
        VK_INSERT => Some((2, b'~')),
        VK_DELETE => Some((3, b'~')),
        VK_PRIOR => Some((5, b'~')), // Page Up
        VK_NEXT => Some((6, b'~')),  // Page Down
        _ => None,
    }
}

/// ファンクションキー F1-F12 → エスケープシーケンス
fn function_key_sequence(vk: VIRTUAL_KEY, mods: &Modifiers) -> Option<Vec<u8>> {
    // F1-F4: SS3 形式 (修飾キーなし), CSI 形式 (修飾キーあり)
    // F5-F12: CSI {code}~ 形式
    let mp = modifier_param(mods);
    let has_mods = has_modifiers(mods);

    match vk {
        VK_F1 => Some(if has_mods {
            format!("\x1b[1;{}P", mp).into_bytes()
        } else {
            b"\x1bOP".to_vec()
        }),
        VK_F2 => Some(if has_mods {
            format!("\x1b[1;{}Q", mp).into_bytes()
        } else {
            b"\x1bOQ".to_vec()
        }),
        VK_F3 => Some(if has_mods {
            format!("\x1b[1;{}R", mp).into_bytes()
        } else {
            b"\x1bOR".to_vec()
        }),
        VK_F4 => Some(if has_mods {
            format!("\x1b[1;{}S", mp).into_bytes()
        } else {
            b"\x1bOS".to_vec()
        }),
        VK_F5 => Some(fkey_csi(15, mp, has_mods)),
        VK_F6 => Some(fkey_csi(17, mp, has_mods)),
        VK_F7 => Some(fkey_csi(18, mp, has_mods)),
        VK_F8 => Some(fkey_csi(19, mp, has_mods)),
        VK_F9 => Some(fkey_csi(20, mp, has_mods)),
        VK_F10 => Some(fkey_csi(21, mp, has_mods)),
        VK_F11 => Some(fkey_csi(23, mp, has_mods)),
        VK_F12 => Some(fkey_csi(24, mp, has_mods)),
        _ => None,
    }
}

/// F5-F12 の CSI シーケンス: \x1b[{code}~ or \x1b[{code};{mod}~
fn fkey_csi(code: u8, mp: u8, has_mods: bool) -> Vec<u8> {
    if has_mods {
        format!("\x1b[{};{}~", code, mp).into_bytes()
    } else {
        format!("\x1b[{}~", code).into_bytes()
    }
}
