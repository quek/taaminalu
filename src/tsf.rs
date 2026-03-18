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
}

impl CompositionState {
    fn new() -> Self {
        Self { composing: false, preedit: String::new() }
    }
}

/// TSF ITextStoreACP + ITfContextOwnerCompositionSink 実装
#[implement(ITextStoreACP, ITfContextOwnerCompositionSink)]
pub struct TextStore {
    app: Arc<Mutex<App>>,
    hwnd: HWND,
    shared_sink: Arc<Mutex<SharedSink>>,
    lock_flags: Mutex<u32>,
    composition: Arc<Mutex<CompositionState>>,
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
        }
    }

    /// ターミナルバッファのテキスト（composition 中は preedit を含む仮想ドキュメント）
    fn get_text_content(&self) -> String {
        let preedit = self.composition.lock().unwrap().preedit.clone();
        let app = self.app.lock().unwrap();
        let base = app.screen_text();
        if preedit.is_empty() {
            return base;
        }
        // カーソル位置に preedit を挿入した仮想ドキュメントを返す
        let cursor_acp = app.cursor_acp();
        let utf16: Vec<u16> = base.encode_utf16().collect();
        let pos = cursor_acp.min(utf16.len());
        let before = String::from_utf16_lossy(&utf16[..pos]);
        let after = String::from_utf16_lossy(&utf16[pos..]);
        format!("{}{}{}", before, preedit, after)
    }

    /// カーソル ACP（composition 中は preedit 末尾を返す）
    fn cursor_to_acp(&self) -> i32 {
        let preedit_len = self.composition.lock().unwrap().preedit.encode_utf16().count() as i32;
        let app = self.app.lock().unwrap();
        app.cursor_acp() as i32 + preedit_len
    }

    /// ターミナルバッファ上のカーソル ACP（preedit を含まない）
    fn base_cursor_acp(&self) -> i32 {
        self.app.lock().unwrap().cursor_acp() as i32
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
        self.composition.lock().unwrap().composing = true;
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
        let preedit = {
            let mut comp = self.composition.lock().unwrap();
            comp.composing = false;
            std::mem::take(&mut comp.preedit)
        };
        if !preedit.is_empty() {
            let app = self.app.lock().unwrap();
            let _ = app.write_pty(preedit.as_bytes());
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
            *self.lock_flags.lock().unwrap() = dwlockflags;
            let hr = unsafe {
                sink.OnLockGranted(TEXT_STORE_LOCK_FLAGS(dwlockflags))
            };
            *self.lock_flags.lock().unwrap() = 0;
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
            dwStaticFlags: 0,
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
        _acpstart: i32,
        _acpend: i32,
        pchtext: &PCWSTR,
        cch: u32,
    ) -> Result<TS_TEXTCHANGE> {
        let slice = unsafe { std::slice::from_raw_parts(pchtext.0, cch as usize) };
        let text = String::from_utf16_lossy(slice);

        {
            let mut comp = self.composition.lock().unwrap();
            if comp.composing {
                comp.preedit = text;
                drop(comp);
                self.invalidate();
            } else {
                drop(comp);
                let app = self.app.lock().unwrap();
                let _ = app.write_pty(text.as_bytes());
            }
        }

        Ok(TS_TEXTCHANGE {
            acpStart: _acpstart,
            acpOldEnd: _acpend,
            acpNewEnd: _acpstart + cch as i32,
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
        let insert_acp = base_acp + preedit_len;

        if dwflags & TF_IAS_QUERYONLY.0 as u32 != 0 {
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

        let new_end = insert_acp + cch as i32;
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
    pub _thread_mgr: ITfThreadMgr,
    pub _doc_mgr: ITfDocumentMgr,
    pub shared_sink: Arc<Mutex<SharedSink>>,
    pub composition: Arc<Mutex<CompositionState>>,
    app: Arc<Mutex<App>>,
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

    /// PTY出力後にテキスト変更を TSF シンクに通知
    pub fn notify_change(&self) {
        // composition 中は通知しない（preedit が消えてしまうため）
        if self.is_composing() {
            return;
        }

        // ロックを先にドロップしてからコールバック（デッドロック防止）
        let (sink, mask) = {
            let shared = self.shared_sink.lock().unwrap();
            (shared.sink.clone(), shared.mask)
        };
        if let Some(sink) = sink {
            if mask & TS_AS_TEXT_CHANGE != 0 {
                let end_acp = {
                    let app = self.app.lock().unwrap();
                    app.screen_text_utf16_len() as i32
                };
                let change = TS_TEXTCHANGE {
                    acpStart: 0,
                    acpOldEnd: end_acp,
                    acpNewEnd: end_acp,
                };
                unsafe { let _ = sink.OnTextChange(TEXT_STORE_TEXT_CHANGE_FLAGS(0), &change); }
            }
            if mask & TS_AS_SEL_CHANGE != 0 {
                unsafe { let _ = sink.OnSelectionChange(); }
            }
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
            // ITfContextOwnerCompositionSink を登録して composition の開始/終了を検知
            let source: ITfSource = ctx.cast()?;
            composition_sink_cookie = source.AdviseSink(
                &ITfContextOwnerCompositionSink::IID,
                &text_store_unk,
            )?;

            doc_mgr.Push(ctx)?;
        }
        thread_mgr.SetFocus(&doc_mgr)?;
    }

    Ok(TsfContext {
        _thread_mgr: thread_mgr,
        _doc_mgr: doc_mgr,
        shared_sink,
        composition,
        app,
        _composition_sink_cookie: composition_sink_cookie,
    })
}
