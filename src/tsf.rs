use std::sync::{Arc, Mutex};

use windows::core::*;
use windows::Win32::Foundation::*;
use windows::Win32::Graphics::Gdi::{ClientToScreen, InvalidateRect, ScreenToClient};
use windows::Win32::UI::WindowsAndMessaging::GetClientRect;
use windows::Win32::UI::TextServices::*;

use crate::app::App;

/// 共有 TSF シンク（TextStore と外部通知コードの両方からアクセス）
pub struct SharedSink {
    pub sink: Option<ITextStoreACPSink>,
    pub mask: u32,
}

/// 共有 composition 状態（TextStore, TsfContext, window から参照）
pub struct CompositionState {
    pub composing: bool,
    pub preedit: String,
    /// 変換 composition 開始時に検出した、消去すべき既存文字数。
    /// SetText(acpstart, acpend) で acpend > acpstart なら置換操作。
    /// OnEndComposition でこの数だけバックスペースを送ってから確定テキストを送信する。
    pub chars_to_erase: usize,
    /// 確定したがまだ PTY エコーで grid に反映されていない確定文字（エコー待ちオーバーレイ）。
    /// grid のカーソル前 `pending_erase` 文字を `pending_text` で置換して IME に見せ、
    /// grid にエコーされたら（reconcile_pending で検出）解除する。snapshot 凍結の代替。
    pub pending_erase: usize,
    pub pending_text: String,
}

impl CompositionState {
    fn new() -> Self {
        Self {
            composing: false,
            preedit: String::new(),
            chars_to_erase: 0,
            pending_erase: 0,
            pending_text: String::new(),
        }
    }

    /// OnStartComposition: composition 開始。
    /// pending_text は維持する（直前の確定がエコー前なら、その上に新しい preedit を
    /// 重ねて「ねこ」のような連続入力を正しく見せるため）。
    pub fn start(&mut self) {
        self.composing = true;
        self.chars_to_erase = 0;
    }

    /// SetText (composing 中): preedit の ACP 範囲を編集 + 置換検出
    ///
    /// `base_cursor_acp`: preedit 開始位置（ターミナルバッファ上のカーソル ACP）。
    /// MS IME は各文字を個別の ACP 範囲で挿入するため、preedit 全体の上書きではなく
    /// 範囲編集が必要。
    pub fn set_text(&mut self, text: String, acpstart: i32, acpend: i32, base_cursor_acp: i32) {
        // acpstart が preedit 領域の前（ベーステキスト内）にある場合は、
        // PTY に送信済みのテキストを含む置換操作。
        if acpstart < base_cursor_acp {
            if self.chars_to_erase == 0 {
                self.chars_to_erase = (base_cursor_acp - acpstart) as usize;
            }
            self.preedit = text;
            return;
        }

        // preedit 内のオフセット（UTF-16 コードユニット単位）
        let offset_start = (acpstart - base_cursor_acp) as usize;
        let offset_end = (acpend - base_cursor_acp) as usize;

        // UTF-16 オフセット → preedit 文字列のバイトオフセットに変換
        let preedit_utf16: Vec<u16> = self.preedit.encode_utf16().collect();
        let byte_start = String::from_utf16_lossy(
            &preedit_utf16[..offset_start.min(preedit_utf16.len())]
        ).len();
        let byte_end = String::from_utf16_lossy(
            &preedit_utf16[..offset_end.min(preedit_utf16.len())]
        ).len();

        // 範囲 [byte_start..byte_end] を text で置換
        let mut new_preedit = String::with_capacity(self.preedit.len() + text.len());
        new_preedit.push_str(&self.preedit[..byte_start]);
        new_preedit.push_str(&text);
        if byte_end <= self.preedit.len() {
            new_preedit.push_str(&self.preedit[byte_end..]);
        }
        self.preedit = new_preedit;
    }

    /// OnEndComposition: composition 終了。PTY に送るバイト列を返す。
    /// 確定文字は「エコー待ちオーバーレイ」(pending_erase/pending_text) として保持し、
    /// grid にエコーされるまで GetText に重ねて見せる（snapshot 凍結の代替）。
    pub fn end(&mut self) -> Vec<u8> {
        self.composing = false;
        let preedit = std::mem::take(&mut self.preedit);
        let erase = self.chars_to_erase;
        self.chars_to_erase = 0;

        if !preedit.is_empty() {
            self.pending_erase = erase;
            self.pending_text = preedit.clone();
        }

        let mut output = Vec::new();
        if !preedit.is_empty() {
            // 置換対象の既存文字をバックスペースで消去
            output.extend(std::iter::repeat_n(0x7fu8, erase));
            output.extend(preedit.as_bytes());
        }
        output
    }

    /// grid のカーソル前テキストに確定文字（pending_text）が現れていたら、
    /// エコー完了とみなしてオーバーレイを解除する。
    pub fn reconcile_pending(&mut self, grid_before_cursor: &str) {
        if self.pending_text.is_empty() {
            return;
        }
        if grid_before_cursor.ends_with(&self.pending_text) {
            self.pending_erase = 0;
            self.pending_text.clear();
        }
    }
}

/// grid（または snapshot）のテキスト・カーソルに、composition の pending（確定エコー
/// 待ち）と preedit（変換中）を重ねた「仮想ドキュメント」を計算する。
/// TextStore（ロック中スナップショット基準）と TsfContext（最新 grid 基準）の両方が使う。
fn apply_overlays(base: &str, base_cursor: usize, comp: &CompositionState) -> (String, usize) {
    let (t1, c1) = overlay_one(base, base_cursor, comp.pending_erase, &comp.pending_text);
    if comp.composing {
        overlay_one(&t1, c1, comp.chars_to_erase, &comp.preedit)
    } else {
        (t1, c1)
    }
}

/// text のカーソル位置から erase 文字（UTF-16 単位）を削除し ins を挿入する。
/// 戻り値は (新テキスト, 新カーソル UTF-16 ACP)。
fn overlay_one(text: &str, cursor: usize, erase: usize, ins: &str) -> (String, usize) {
    if erase == 0 && ins.is_empty() {
        return (text.to_string(), cursor);
    }
    let utf16: Vec<u16> = text.encode_utf16().collect();
    let pos = cursor.min(utf16.len());
    let erase_start = pos.saturating_sub(erase);
    let before = String::from_utf16_lossy(&utf16[..erase_start]);
    let after = String::from_utf16_lossy(&utf16[pos..]);
    let new_cursor = erase_start + ins.encode_utf16().count();
    let mut s = String::with_capacity(before.len() + ins.len() + after.len());
    s.push_str(&before);
    s.push_str(ins);
    s.push_str(&after);
    (s, new_cursor)
}

/// 旧テキストと新テキストの差分範囲を UTF-16 ACP で算出（前方/後方スキャン、Chromium 方式）。
/// 戻り値は (acpStart, acpOldEnd, acpNewEnd)。
fn text_diff(old: &str, new: &str) -> (i32, i32, i32) {
    let o: Vec<u16> = old.encode_utf16().collect();
    let n: Vec<u16> = new.encode_utf16().collect();
    let mut start = 0usize;
    let min_len = o.len().min(n.len());
    while start < min_len && o[start] == n[start] {
        start += 1;
    }
    let (mut old_end, mut new_end) = (o.len(), n.len());
    while old_end > start && new_end > start && o[old_end - 1] == n[new_end - 1] {
        old_end -= 1;
        new_end -= 1;
    }
    (start as i32, old_end as i32, new_end as i32)
}

/// ロック中にバッファの一貫性を保つためのスナップショット（1 ロック単位）。
/// RequestLock で実画面から作成し OnLockGranted 中だけ有効。次のロックで作り直す。
/// PTY リーダーがロック中にバッファを更新しても GetText/GetSelection は同じ状態を返す。
struct LockSnapshot {
    base_text: Arc<String>, // screen_text（pending/preedit を含まない素の grid）
    base_cursor_acp: usize, // grid のカーソル ACP（pending/preedit を含まない）
}

/// TSF ITextStoreACP + ITfContextOwnerCompositionSink 実装
#[implement(ITextStoreACP, ITfContextOwnerCompositionSink)]
pub struct TextStore {
    app: Arc<Mutex<App>>,
    hwnd: HWND,
    shared_sink: Arc<Mutex<SharedSink>>,
    lock_flags: Mutex<u32>,
    composition: Arc<Mutex<CompositionState>>,
    /// OnLockGranted 中のみ有効な 1 ロック単位のスナップショット
    snapshot: Mutex<Option<LockSnapshot>>,
}

impl TextStore {
    pub fn new(
        app: Arc<Mutex<App>>,
        hwnd: HWND,
        shared_sink: Arc<Mutex<SharedSink>>,
        composition: Arc<Mutex<CompositionState>>,
    ) -> Self {
        Self {
            app,
            hwnd,
            shared_sink,
            lock_flags: Mutex::new(0),
            composition,
            snapshot: Mutex::new(None),
        }
    }

    /// ベーステキストとカーソル ACP を取得（スナップショットがあればそこから）
    fn base_text_and_cursor(&self) -> (Arc<String>, usize) {
        if let Some(ref snap) = *self.snapshot.lock().unwrap() {
            return (Arc::clone(&snap.base_text), snap.base_cursor_acp);
        }
        let app = self.app.lock().unwrap();
        (Arc::new(app.screen_text()), app.cursor_acp())
    }

    /// IME に見せる仮想ドキュメント（grid に pending と preedit を重ねたもの）
    fn get_text_content(&self) -> Arc<String> {
        let (base, cursor_acp) = self.base_text_and_cursor();
        let comp = self.composition.lock().unwrap();
        // pending も preedit も無ければ素の grid をそのまま返す（clone 回避）
        if !comp.composing && comp.pending_text.is_empty() {
            return base;
        }
        let (text, _) = apply_overlays(&base, cursor_acp, &comp);
        Arc::new(text)
    }

    /// IME に見せるカーソル ACP（pending と preedit を重ねた後の末尾位置）
    fn cursor_to_acp(&self) -> i32 {
        let (base, base_acp) = self.base_text_and_cursor();
        let comp = self.composition.lock().unwrap();
        // オーバーレイが無ければ grid のカーソルそのもの（clone を避ける）
        if !comp.composing && comp.pending_text.is_empty() {
            return base_acp as i32;
        }
        let (_, cursor) = apply_overlays(&base, base_acp, &comp);
        cursor as i32
    }

    /// preedit の開始位置 ACP（pending は適用済み・preedit は含まない）。
    /// SetText/InsertTextAtSelection が「変換中テキストの基準点」として使う。
    fn base_cursor_acp(&self) -> i32 {
        let (base, base_acp) = self.base_text_and_cursor();
        let comp = self.composition.lock().unwrap();
        // 確定エコー待ちが無ければ grid のカーソルそのもの（clone を避ける）
        if comp.pending_text.is_empty() {
            return base_acp as i32;
        }
        let (_, cursor) = overlay_one(&base, base_acp, comp.pending_erase, &comp.pending_text);
        cursor as i32
    }

    fn invalidate(&self) {
        unsafe { let _ = InvalidateRect(Some(self.hwnd), None, false); }
    }

    /// preedit 文字列内の UTF-16 オフセットを表示カラム数に変換
    fn preedit_utf16_offset_to_cols(&self, utf16_offset: i32) -> usize {
        let comp = self.composition.lock().unwrap();
        let mut utf16_count = 0i32;
        let mut cols = 0usize;
        for c in comp.preedit.chars() {
            if utf16_count >= utf16_offset {
                break;
            }
            utf16_count += c.len_utf16() as i32;
            cols += if c.is_ascii() { 1 } else { 2 };
        }
        cols
    }
}

// --- ITfContextOwnerCompositionSink ---

impl ITfContextOwnerCompositionSink_Impl for TextStore_Impl {
    fn OnStartComposition(
        &self,
        _pcomposition: Ref<ITfCompositionView>,
    ) -> Result<BOOL> {
        self.composition.lock().unwrap().start();
        Ok(TRUE)
    }

    fn OnUpdateComposition(
        &self,
        _pcomposition: Ref<ITfCompositionView>,
        _prangenew: Ref<ITfRange>,
    ) -> Result<()> {
        self.invalidate();
        Ok(())
    }

    fn OnEndComposition(
        &self,
        _pcomposition: Ref<ITfCompositionView>,
    ) -> Result<()> {
        // end() が確定文字を pending（エコー待ちオーバーレイ）として保持するので、
        // rtry が直後に GetText しても確定済み文字が見える（snapshot 凍結は不要）。
        let output = self.composition.lock().unwrap().end();
        if !output.is_empty() {
            // TSF ロック中の直接 write_pty は ConPTY エコーのタイミング問題を起こすため、
            // PostMessage で遅延送信する。
            crate::window::post_deferred_pty_write(self.hwnd, output);
        }
        self.invalidate();
        Ok(())
    }
}

// --- ITextStoreACP ---

impl ITextStoreACP_Impl for TextStore_Impl {
    fn AdviseSink(
        &self,
        _riid: *const GUID,
        punk: Ref<IUnknown>,
        dwmask: u32,
    ) -> Result<()> {
        let sink: ITextStoreACPSink = punk.ok()?.cast()?;
        let mut shared = self.shared_sink.lock().unwrap();
        shared.sink = Some(sink);
        shared.mask = dwmask;
        Ok(())
    }

    fn UnadviseSink(&self, _punk: Ref<IUnknown>) -> Result<()> {
        let mut shared = self.shared_sink.lock().unwrap();
        shared.sink = None;
        shared.mask = 0;
        Ok(())
    }

    fn RequestLock(&self, dwlockflags: u32) -> Result<HRESULT> {
        let sink = self.shared_sink.lock().unwrap().sink.clone();
        if let Some(sink) = sink {
            // 実画面からこのロック専用のスナップショットを作る（凍結しない）。
            // 同時に、確定文字が grid にエコー済みなら pending を解除する。
            let (grid_text, grid_cursor) = {
                let app = self.app.lock().unwrap();
                (app.screen_text(), app.cursor_acp())
            };
            {
                let mut comp = self.composition.lock().unwrap();
                let utf16: Vec<u16> = grid_text.encode_utf16().collect();
                let pos = grid_cursor.min(utf16.len());
                let before = String::from_utf16_lossy(&utf16[..pos]);
                comp.reconcile_pending(&before);
            }
            *self.snapshot.lock().unwrap() = Some(LockSnapshot {
                base_text: Arc::new(grid_text),
                base_cursor_acp: grid_cursor,
            });

            *self.lock_flags.lock().unwrap() = dwlockflags;
            let hr = unsafe {
                sink.OnLockGranted(TEXT_STORE_LOCK_FLAGS(dwlockflags))
            };
            *self.lock_flags.lock().unwrap() = 0;

            // ロックを抜けたらスナップショットは無効化（次のロックで作り直す）
            *self.snapshot.lock().unwrap() = None;

            match hr {
                Ok(()) => Ok(S_OK),
                Err(e) => Ok(e.code()),
            }
        } else {
            Ok(E_FAIL)
        }
    }

    fn GetStatus(&self) -> Result<TS_STATUS> {
        Ok(TS_STATUS {
            dwDynamicFlags: 0,
            dwStaticFlags: TS_SS_NOHIDDENTEXT,
        })
    }

    fn QueryInsert(
        &self,
        acpteststart: i32,
        acptestend: i32,
        _cch: u32,
        pacpresultstart: *mut i32,
        pacpresultend: *mut i32,
    ) -> Result<()> {
        unsafe {
            *pacpresultstart = acpteststart;
            *pacpresultend = acptestend;
        }
        Ok(())
    }

    fn GetSelection(
        &self,
        _ulindex: u32,
        ulcount: u32,
        pselection: *mut TS_SELECTION_ACP,
        pcfetched: *mut u32,
    ) -> Result<()> {
        if ulcount == 0 {
            return Ok(());
        }
        let acp = self.cursor_to_acp();
        unsafe {
            (*pselection).acpStart = acp;
            (*pselection).acpEnd = acp;
            (*pselection).style.ase = TS_AE_END;
            (*pselection).style.fInterimChar = false.into();
            *pcfetched = 1;
        }
        Ok(())
    }

    fn SetSelection(&self, _ulcount: u32, _pselection: *const TS_SELECTION_ACP) -> Result<()> {
        Ok(())
    }

    fn GetText(
        &self,
        acpstart: i32,
        acpend: i32,
        pchplain: PWSTR,
        cchplainreq: u32,
        pcchplainret: *mut u32,
        prgruninfo: *mut TS_RUNINFO,
        cruninforeq: u32,
        pcruninforet: *mut u32,
        pacpnext: *mut i32,
    ) -> Result<()> {
        let text = self.get_text_content();
        let utf16: Vec<u16> = text.encode_utf16().collect();

        let start = (acpstart as usize).min(utf16.len());
        let end = if acpend == -1 {
            utf16.len()
        } else {
            (acpend as usize).min(utf16.len())
        };
        let end = end.max(start); // start > end でのパニックを防止

        let slice = &utf16[start..end];
        let copy_len = slice.len().min(cchplainreq as usize);

        unsafe {
            if !pchplain.is_null() && copy_len > 0 {
                std::ptr::copy_nonoverlapping(slice.as_ptr(), pchplain.0, copy_len);
            }
            *pcchplainret = copy_len as u32;

            if cruninforeq > 0 && !prgruninfo.is_null() {
                (*prgruninfo).uCount = copy_len as u32;
                (*prgruninfo).r#type = TS_RT_PLAIN;
                *pcruninforet = 1;
            }

            *pacpnext = (start + copy_len) as i32;
        }
        Ok(())
    }

    fn SetText(
        &self,
        _dwflags: u32,
        acpstart: i32,
        acpend: i32,
        pchtext: &PCWSTR,
        cch: u32,
    ) -> Result<TS_TEXTCHANGE> {
        let slice = unsafe { std::slice::from_raw_parts(pchtext.0, cch as usize) };
        let text = String::from_utf16_lossy(slice);

        let composing = self.composition.lock().unwrap().composing;
        if composing {
            // base_cursor_acp は内部で composition をロックするため、comp を保持したまま
            // 呼ぶと同一スレッドで二重ロックしてデッドロックする。先にロック外で計算する。
            let base_acp = self.base_cursor_acp();
            self.composition.lock().unwrap().set_text(text, acpstart, acpend, base_acp);
            self.invalidate();
        } else {
            let app = self.app.lock().unwrap();
            let _ = app.write_pty(text.as_bytes());
        }

        Ok(TS_TEXTCHANGE {
            acpStart: acpstart,
            acpOldEnd: acpend,
            acpNewEnd: acpstart.saturating_add(cch as i32),
        })
    }

    fn GetFormattedText(
        &self,
        _acpstart: i32,
        _acpend: i32,
    ) -> Result<windows::Win32::System::Com::IDataObject> {
        Err(E_NOTIMPL.into())
    }

    fn GetEmbedded(
        &self,
        _acppos: i32,
        _rguidservice: *const GUID,
        _riid: *const GUID,
    ) -> Result<IUnknown> {
        Err(E_NOTIMPL.into())
    }

    fn QueryInsertEmbedded(
        &self,
        _pguidservice: *const GUID,
        _pformatetc: *const windows::Win32::System::Com::FORMATETC,
    ) -> Result<BOOL> {
        Ok(FALSE)
    }

    fn InsertEmbedded(
        &self,
        _dwflags: u32,
        _acpstart: i32,
        _acpend: i32,
        _pdataobject: Ref<windows::Win32::System::Com::IDataObject>,
    ) -> Result<TS_TEXTCHANGE> {
        Err(E_NOTIMPL.into())
    }

    fn InsertTextAtSelection(
        &self,
        dwflags: u32,
        pchtext: &PCWSTR,
        cch: u32,
        pacpstart: *mut i32,
        pacpend: *mut i32,
        pchange: *mut TS_TEXTCHANGE,
    ) -> Result<()> {
        let (composing, preedit_len) = {
            let comp = self.composition.lock().unwrap();
            (comp.composing, comp.preedit.encode_utf16().count() as i32)
        };
        let base_acp = self.base_cursor_acp();
        // composition 中は preedit 末尾が挿入点
        let insert_acp = base_acp.saturating_add(preedit_len);

        if dwflags & TF_IAS_QUERYONLY.0 != 0 {
            unsafe {
                if !pacpstart.is_null() { *pacpstart = insert_acp; }
                if !pacpend.is_null() { *pacpend = insert_acp; }
            }
            return Ok(());
        }

        let slice = unsafe { std::slice::from_raw_parts(pchtext.0, cch as usize) };
        let text = String::from_utf16_lossy(slice);

        if composing {
            // Composition 中: preedit に追加（PTY には送信しない）
            self.composition.lock().unwrap().preedit.push_str(&text);
            self.invalidate();
        } else {
            // Composition 外: PTY に直接送信
            let app = self.app.lock().unwrap();
            let _ = app.write_pty(text.as_bytes());
        }

        let new_end = insert_acp.saturating_add(cch as i32);
        unsafe {
            if !pacpstart.is_null() { *pacpstart = insert_acp; }
            if !pacpend.is_null() { *pacpend = new_end; }
            if !pchange.is_null() {
                (*pchange).acpStart = insert_acp;
                (*pchange).acpOldEnd = insert_acp;
                (*pchange).acpNewEnd = new_end;
            }
        }
        Ok(())
    }

    fn InsertEmbeddedAtSelection(
        &self,
        _dwflags: u32,
        _pdataobject: Ref<windows::Win32::System::Com::IDataObject>,
        _pacpstart: *mut i32,
        _pacpend: *mut i32,
        _pchange: *mut TS_TEXTCHANGE,
    ) -> Result<()> {
        Err(E_NOTIMPL.into())
    }

    fn RequestSupportedAttrs(&self, _dwflags: u32, _cfilterattrs: u32, _pafilterattrs: *const GUID) -> Result<()> {
        Ok(())
    }

    fn RequestAttrsAtPosition(&self, _acppos: i32, _cfilterattrs: u32, _pafilterattrs: *const GUID, _dwflags: u32) -> Result<()> {
        Ok(())
    }

    fn RequestAttrsTransitioningAtPosition(&self, _acppos: i32, _cfilterattrs: u32, _pafilterattrs: *const GUID, _dwflags: u32) -> Result<()> {
        Ok(())
    }

    fn FindNextAttrTransition(
        &self, _acpstart: i32, _acphalt: i32, _cfilterattrs: u32,
        _pafilterattrs: *const GUID, _dwflags: u32,
        pacpnext: *mut i32, pffound: *mut BOOL, _plfoundoffset: *mut i32,
    ) -> Result<()> {
        unsafe { *pacpnext = 0; *pffound = FALSE; }
        Ok(())
    }

    fn RetrieveRequestedAttrs(&self, _ulcount: u32, _paattrvals: *mut TS_ATTRVAL, pcfetched: *mut u32) -> Result<()> {
        unsafe { *pcfetched = 0; }
        Ok(())
    }

    fn GetEndACP(&self) -> Result<i32> {
        let text = self.get_text_content();
        let len: usize = text.encode_utf16().count();
        Ok(len as i32)
    }

    fn GetActiveView(&self) -> Result<u32> { Ok(1) }

    fn GetACPFromPoint(&self, _vcview: u32, ptscreen: *const POINT, _dwflags: u32) -> Result<i32> {
        let pt = unsafe { *ptscreen };
        let app = self.app.lock().unwrap();
        let (cell_w, cell_h) = app.cell_size();
        let (_, grid_y) = app.grid_origin();
        let mut client_pt = pt;
        unsafe { let _ = ScreenToClient(self.hwnd, &mut client_pt); }
        let col = (client_pt.x as f32 / cell_w) as usize;
        let row = ((client_pt.y as f32 - grid_y) / cell_h).max(0.0) as usize;
        let acp = app.active().term.grid_to_acp(row, col);
        Ok(acp as i32)
    }

    fn GetTextExt(&self, _vcview: u32, acpstart: i32, acpend: i32, prc: *mut RECT, pfclipped: *mut BOOL) -> Result<()> {
        let composing = self.composition.lock().unwrap().composing;
        let app = self.app.lock().unwrap();
        let (cell_w, cell_h) = app.cell_size();
        let (_, grid_y) = app.grid_origin();
        let grid_y_i = grid_y as i32;

        let mut rect = if composing {
            // Composition 中: カーソル位置を直接使う（ACP→グリッド変換の誤差を回避）
            let (cursor_row, cursor_col) = app.active().term.cursor_pos();
            let preedit_cols = self.preedit_utf16_offset_to_cols(acpend - acpstart);
            let x = (cursor_col as f32 * cell_w) as i32;
            let y = (cursor_row as f32 * cell_h) as i32 + grid_y_i;
            let right = x + (preedit_cols as f32 * cell_w) as i32;
            RECT { left: x, top: y, right: right.max(x), bottom: y + cell_h as i32 }
        } else {
            let (start_row, start_col) = app.acp_to_grid(acpstart as usize);
            if acpstart == acpend {
                let x = (start_col as f32 * cell_w) as i32;
                let y = (start_row as f32 * cell_h) as i32 + grid_y_i;
                RECT { left: x, top: y, right: x, bottom: y + cell_h as i32 }
            } else {
                let (end_row, end_col) = app.acp_to_grid(acpend as usize);
                RECT {
                    left: (start_col as f32 * cell_w) as i32,
                    top: (start_row as f32 * cell_h) as i32 + grid_y_i,
                    right: (end_col as f32 * cell_w) as i32,
                    bottom: ((end_row + 1) as f32 * cell_h) as i32 + grid_y_i,
                }
            }
        };

        // クライアント座標 → スクリーン座標
        let mut top_left = POINT { x: rect.left, y: rect.top };
        let mut bottom_right = POINT { x: rect.right, y: rect.bottom };
        unsafe {
            let _ = ClientToScreen(self.hwnd, &mut top_left);
            let _ = ClientToScreen(self.hwnd, &mut bottom_right);
            rect.left = top_left.x;
            rect.top = top_left.y;
            rect.right = bottom_right.x;
            rect.bottom = bottom_right.y;
            *prc = rect;
            *pfclipped = FALSE;
        }
        Ok(())
    }

    fn GetScreenExt(&self, _vcview: u32) -> Result<RECT> {
        let mut rect = RECT::default();
        unsafe {
            let _ = GetClientRect(self.hwnd, &mut rect);
            let mut top_left = POINT { x: rect.left, y: rect.top };
            let mut bottom_right = POINT { x: rect.right, y: rect.bottom };
            let _ = ClientToScreen(self.hwnd, &mut top_left);
            let _ = ClientToScreen(self.hwnd, &mut bottom_right);
            rect.left = top_left.x;
            rect.top = top_left.y;
            rect.right = bottom_right.x;
            rect.bottom = bottom_right.y;
        }
        Ok(rect)
    }

    fn GetWnd(&self, _vcview: u32) -> Result<HWND> { Ok(self.hwnd) }
}

/// TSF セットアップの戻り値
pub struct TsfContext {
    pub thread_mgr: ITfThreadMgr,
    pub doc_mgr: ITfDocumentMgr,
    pub keystroke_mgr: ITfKeystrokeMgr,
    pub shared_sink: Arc<Mutex<SharedSink>>,
    pub composition: Arc<Mutex<CompositionState>>,
    app: Arc<Mutex<App>>,
    /// 前回 IME に通知した (仮想ドキュメント, カーソル ACP)。差分通知の基準。
    last_notified: Mutex<(String, i32)>,
    _composition_sink_cookie: u32,
}

// COM オブジェクトはメインスレッド (STA) でのみ使用するため安全
unsafe impl Send for TsfContext {}
unsafe impl Sync for TsfContext {}

impl TsfContext {
    /// Composition 中かどうか
    pub fn is_composing(&self) -> bool {
        self.composition.lock().unwrap().composing
    }

    /// 現在の preedit テキスト
    pub fn preedit(&self) -> String {
        self.composition.lock().unwrap().preedit.clone()
    }

    /// 画面更新後、テキスト/カーソルが実際に変化したぶんだけ TSF シンクに通知する。
    /// ロック外（フレームタイマー等）から呼ばれる。Chromium tsf_text_store と同じ差分方式：
    /// 前回通知値と比較し、変化が無ければ何も送らない。
    pub fn notify_change(&self) {
        // composition 中は通知しない（preedit の最中に ACP が動くと変換が壊れる）
        if self.is_composing() {
            return;
        }

        let (sink, mask) = {
            let shared = self.shared_sink.lock().unwrap();
            (shared.sink.clone(), shared.mask)
        };
        let Some(sink) = sink else { return };

        // 現在の仮想ドキュメント（grid + pending）とカーソルを計算
        let (cur_text, cur_cursor) = {
            let app = self.app.lock().unwrap();
            let comp = self.composition.lock().unwrap();
            let base = app.screen_text();
            let cursor = app.cursor_acp();
            // オーバーレイが無ければ素の grid をそのまま使う（clone を避ける）
            if !comp.composing && comp.pending_text.is_empty() {
                (base, cursor)
            } else {
                apply_overlays(&base, cursor, &comp)
            }
        };
        let cur_cursor = cur_cursor as i32;

        let mut last = self.last_notified.lock().unwrap();
        // テキストが変わったぶんだけ最小レンジで通知
        if mask & TS_AS_TEXT_CHANGE != 0 && last.0 != cur_text {
            let (start, old_end, new_end) = text_diff(&last.0, &cur_text);
            if start != old_end || start != new_end {
                let change = TS_TEXTCHANGE { acpStart: start, acpOldEnd: old_end, acpNewEnd: new_end };
                unsafe { let _ = sink.OnTextChange(TEXT_STORE_TEXT_CHANGE_FLAGS(0), &change); }
            }
        }
        // カーソルが動いたときだけ選択変更を通知
        if mask & TS_AS_SEL_CHANGE != 0 && last.1 != cur_cursor {
            unsafe { let _ = sink.OnSelectionChange(); }
        }
        *last = (cur_text, cur_cursor);
    }

    /// テキスト/カーソルの画面座標が変わったとき（スクロール・リサイズ・移動）に
    /// レイアウト変更を通知し、IME に GetTextExt を再取得させ候補位置を追従させる。
    pub fn notify_layout_change(&self) {
        let (sink, mask) = {
            let shared = self.shared_sink.lock().unwrap();
            (shared.sink.clone(), shared.mask)
        };
        if let Some(sink) = sink
            && mask & TS_AS_LAYOUT_CHANGE != 0
        {
            unsafe { let _ = sink.OnLayoutChange(TS_LC_CHANGE, 1); }
        }
    }
}

/// TSF をセットアップして TextStore を登録
pub fn setup_tsf(
    app: Arc<Mutex<App>>,
    hwnd: HWND,
) -> Result<TsfContext> {
    let thread_mgr: ITfThreadMgr = unsafe {
        windows::Win32::System::Com::CoCreateInstance(
            &CLSID_TF_ThreadMgr,
            None,
            windows::Win32::System::Com::CLSCTX_INPROC_SERVER,
        )?
    };

    let client_id = unsafe { thread_mgr.Activate()? };

    let doc_mgr = unsafe { thread_mgr.CreateDocumentMgr()? };

    #[allow(clippy::arc_with_non_send_sync)] // COM オブジェクトは Send/Sync ではないが Arc<Mutex> で保護
    let shared_sink = Arc::new(Mutex::new(SharedSink { sink: None, mask: 0 }));
    let composition = Arc::new(Mutex::new(CompositionState::new()));
    let text_store = TextStore::new(
        Arc::clone(&app),
        hwnd,
        Arc::clone(&shared_sink),
        Arc::clone(&composition),
    );
    let text_store_unk: IUnknown = text_store.into();

    let mut context: Option<ITfContext> = None;
    let mut edit_cookie = 0u32;
    let mut composition_sink_cookie = 0u32;
    unsafe {
        doc_mgr.CreateContext(client_id, 0, &text_store_unk, &mut context, &mut edit_cookie)?;

        if let Some(ref ctx) = context {
            let source: ITfSource = ctx.cast()?;
            // CONNECT_E_ADVISELIMIT: CreateContext 時に TSF が TextStore の
            // ITfContextOwnerCompositionSink を自動登録済みの場合があるため、
            // エラーは無視する。
            if let Ok(cookie) = source.AdviseSink(
                &ITfContextOwnerCompositionSink::IID,
                &text_store_unk,
            ) {
                composition_sink_cookie = cookie;
            }

            doc_mgr.Push(ctx)?;
        }
        thread_mgr.SetFocus(&doc_mgr)?;
    }

    // ウィンドウにドキュメントマネージャを関連付け。
    // WM_SETFOCUS 時に TSF が自動的に SetFocus を呼び出し、
    // IME (rtry 等) が ITextStoreACP にアクセスできるようにする。
    let _ = unsafe { thread_mgr.AssociateFocus(hwnd, &doc_mgr) };

    // ITfKeystrokeMgr を取得。メッセージループで TSF にキーをルーティングするために必要。
    // これがないと CUAS が IMM32 互換パスにフォールバックし、MS IME の composition が
    // ITextStoreACP 経由ではなく WM_IME_COMPOSITION 経由になる。
    let keystroke_mgr: ITfKeystrokeMgr = thread_mgr.cast()?;

    Ok(TsfContext {
        thread_mgr,
        doc_mgr,
        keystroke_mgr,
        shared_sink,
        composition,
        app,
        last_notified: Mutex::new((String::new(), -1)),
        _composition_sink_cookie: composition_sink_cookie,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::term::TermWrapper;

    const BS: u8 = 0x7f;

    // --- Composition: rtry パターン（1文字ずつ独立した composition）---
    // rtry は各文字を独立した composition で確定する。変換時は PTY 送信済みテキストを置換。
    // composition 間でカーソル（base_cursor_acp）が進む。

    #[test]
    fn test_rtry風変換でバックスペースと漢字が送られる() {
        let mut comp = CompositionState::new();

        // "ね" 入力（base=0, 挿入: acp=0..0）
        comp.start();
        comp.set_text("ね".into(), 0, 0, 0);
        let out1 = comp.end();
        assert_eq!(out1, "ね".as_bytes());
        // カーソル → 1

        // "こ" 入力（base=1, 挿入: acp=1..1）
        comp.start();
        comp.set_text("こ".into(), 1, 1, 1);
        let out2 = comp.end();
        assert_eq!(out2, "こ".as_bytes());
        // カーソル → 2

        // "ねこ" → "猫" 変換（base=2, 置換: acp=0..2 → base より前 → BS×2）
        comp.start();
        comp.set_text("猫".into(), 0, 2, 2);
        let out3 = comp.end();
        let mut expected = vec![BS, BS];
        expected.extend("猫".as_bytes());
        assert_eq!(out3, expected);
    }

    #[test]
    fn test_変換なしの連続入力でバックスペースが送られない() {
        let mut comp = CompositionState::new();

        comp.start();
        comp.set_text("ね".into(), 0, 0, 0);
        let out1 = comp.end();
        assert_eq!(out1, "ね".as_bytes());

        comp.start();
        comp.set_text("こ".into(), 1, 1, 1);
        let out2 = comp.end();
        assert_eq!(out2, "こ".as_bytes());

        assert!(!out1.contains(&BS));
        assert!(!out2.contains(&BS));
    }

    #[test]
    fn test_単一文字の変換() {
        let mut comp = CompositionState::new();

        // "あ" 入力（base=0）
        comp.start();
        comp.set_text("あ".into(), 0, 0, 0);
        let out1 = comp.end();
        assert_eq!(out1, "あ".as_bytes());
        // カーソル → 1

        // "あ" → "亜" 変換（base=1, acp=0..1 → base より前 → BS×1）
        comp.start();
        comp.set_text("亜".into(), 0, 1, 1);
        let out2 = comp.end();
        let mut expected = vec![BS];
        expected.extend("亜".as_bytes());
        assert_eq!(out2, expected);
    }

    #[test]
    fn test_preeditが空なら何も送信しない() {
        let mut comp = CompositionState::new();
        comp.start();
        let output = comp.end();
        assert!(output.is_empty());
    }

    // --- Composition: MS IME パターン（1 composition 内で逐次挿入）---
    // MS IME は1つの composition 内で全文字を逐次挿入し、変換も preedit 内で行う。
    // base_cursor_acp は composition 中変わらない。

    #[test]
    fn test_ms_ime風の逐次挿入でpreeditが蓄積される() {
        let mut comp = CompositionState::new();
        comp.start();
        comp.set_text("あ".into(), 0, 0, 0);
        assert_eq!(comp.preedit, "あ");
        comp.set_text("い".into(), 1, 1, 0);
        assert_eq!(comp.preedit, "あい");
        comp.set_text("う".into(), 2, 2, 0);
        assert_eq!(comp.preedit, "あいう");
        let output = comp.end();
        assert_eq!(output, "あいう".as_bytes());
    }

    #[test]
    fn test_ms_ime風の変換はpreedit内置換なのでバックスペースなし() {
        let mut comp = CompositionState::new();
        comp.start();
        comp.set_text("あ".into(), 0, 0, 0);
        comp.set_text("い".into(), 1, 1, 0);
        comp.set_text("う".into(), 2, 2, 0);
        assert_eq!(comp.preedit, "あいう");
        // 変換: preedit 内で "あいう" → "合い"（acp=0..3, base=0 → preedit 内置換）
        comp.set_text("合い".into(), 0, 3, 0);
        assert_eq!(comp.preedit, "合い");
        let output = comp.end();
        // preedit 内の置換なので BS なし
        assert_eq!(output, "合い".as_bytes());
    }

    // --- GetText 基盤: screen_text ---

    #[test]
    fn test_screen_textが日本語テキストを含む() {
        let mut term = TermWrapper::new(80, 24);
        term.process("echo 漢字\r\n".as_bytes());
        let text = term.screen_text();
        assert!(text.contains("漢字"), "screen_text に '漢字' が含まれるべき: {:?}", &text[..80.min(text.len())]);
    }

    #[test]
    fn test_screen_textがwide_char_spacerをスキップする() {
        let mut term = TermWrapper::new(80, 24);
        term.process("日本\r\n".as_bytes());
        let text = term.screen_text();
        // "日本" の後に spacer 文字が混じっていないこと
        let first_line: String = text.lines().next().unwrap_or("").chars().take(4).collect();
        assert!(first_line.starts_with("日本"), "先頭が '日本' であるべき: {:?}", first_line);
    }

    // --- オーバーレイ（pending / preedit を grid に重ねる）---

    #[test]
    fn test_overlay_oneは末尾に挿入する() {
        // "ab" のカーソル末尾(2) に erase=0 で "X" 挿入 → "abX"
        assert_eq!(overlay_one("ab", 2, 0, "X"), ("abX".into(), 3));
    }

    #[test]
    fn test_overlay_oneは置換する() {
        // "ねこ" のカーソル末尾(2) で erase=2, "猫" → "猫"（交ぜ書き変換の確定）
        assert_eq!(overlay_one("ねこ", 2, 2, "猫"), ("猫".into(), 1));
    }

    #[test]
    fn test_確定pendingは挿入でなく置換で重なる() {
        // 「ねこ猫」バグ再発防止: grid "ねこ" に pending(erase=2,"猫") → "猫"
        let mut comp = CompositionState::new();
        comp.pending_erase = 2;
        comp.pending_text = "猫".into();
        let (text, cursor) = apply_overlays("ねこ", 2, &comp);
        assert_eq!(text, "猫", "確定 overlay は挿入でなく置換（'ねこ猫' にならない）");
        assert_eq!(cursor, 1);
    }

    #[test]
    fn test_preeditは確定pendingの上に乗る() {
        // 連続入力: grid 空 + pending "ね"(エコー前) + preedit "こ" → "ねこ"
        let mut comp = CompositionState::new();
        comp.pending_text = "ね".into();
        comp.composing = true;
        comp.preedit = "こ".into();
        assert_eq!(apply_overlays("", 0, &comp), ("ねこ".into(), 2));
    }

    #[test]
    fn test_reconcile_pendingはエコー検出で解除する() {
        let mut comp = CompositionState::new();
        comp.pending_text = "ね".into();
        // grid のカーソル前にまだ確定文字が無い → 維持
        comp.reconcile_pending("> ");
        assert_eq!(comp.pending_text, "ね", "エコー前は維持");
        // grid のカーソル前に "ね" が現れた → 解除
        comp.reconcile_pending("> ね");
        assert_eq!(comp.pending_text, "", "エコー後は解除");
        assert_eq!(comp.pending_erase, 0);
    }

    #[test]
    fn test_text_diffは最小レンジを返す() {
        assert_eq!(text_diff("text", "Text"), (0, 1, 1)); // 先頭 t→T 置換
        assert_eq!(text_diff("text", "tex"), (3, 4, 3));  // 末尾 t 削除
        assert_eq!(text_diff("text", "teaxt"), (2, 2, 3)); // e,x 間に a 挿入
        assert_eq!(text_diff("abc", "abc"), (3, 3, 3));   // 変化なし（呼び出し側で送らない）
    }
}
