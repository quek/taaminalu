# 調査対象プロジェクト

| プロジェクト | 言語 | 特徴 | クローン先 | URL |
|---|---|---|---|---|
| Alacritty | Rust | alacritty_terminal の使い方、ConPTY 接続。**最優先** | /tmp/alacritty | https://github.com/alacritty/alacritty |
| WezTerm | Rust | ターミナルバッファ設計、Windows 描画 | /tmp/wezterm | https://github.com/wezterm/wezterm |
| Windows Terminal | C++ | ConPTY 設計元、TSF/IME 統合、DirectWrite 描画 | /tmp/terminal | https://github.com/microsoft/terminal |
| Chromium | C++ | ITextStoreACP 実装のリファレンス | (Web参照) | https://chromium.googlesource.com/chromium/src/+/lkgr/ui/base/ime/win/tsf_text_store.cc |
| Firefox | C++ | ITextStoreACP + composition。レガシー IME 回避策が豊富 | (Web参照) | https://searchfox.org/mozilla-central/source/widget/windows/TSFTextStore.cpp |

全プロジェクトを調査する必要はない。機能に最も関連するものを優先する。

# API リファレンス・ガイド

| ドキュメント | URL |
|---|---|
| TSF (Text Services Framework) | https://learn.microsoft.com/en-us/windows/win32/tsf/text-services-framework |
| ITextStoreACP | https://learn.microsoft.com/en-us/windows/win32/api/textstor/nn-textstor-itextstoreacp |
| ITfContextOwnerCompositionSink | https://learn.microsoft.com/en-us/windows/win32/api/msctf/nn-msctf-itfcontextownercompositionsink |
| IMM32 (Input Method Manager) | https://learn.microsoft.com/en-us/windows/win32/intl/input-method-manager |
| ConPTY (Creating a Pseudoconsole) | https://learn.microsoft.com/en-us/windows/console/creating-a-pseudoconsole-session |
| Direct2D | https://learn.microsoft.com/en-us/windows/win32/direct2d/direct2d-portal |
| DirectWrite | https://learn.microsoft.com/en-us/windows/win32/directwrite/direct-write-portal |
| COM プログラミング | https://learn.microsoft.com/en-us/windows/win32/com/component-object-model--com--portal |
| windows crate (Rust) | https://microsoft.github.io/windows-docs-rs/ |
| Win32 API | https://learn.microsoft.com/en-us/windows/win32/api/ |
| alacritty_terminal docs | https://docs.rs/alacritty_terminal/latest/alacritty_terminal/ |

# 機能と API の対応例

| 機能 | 主な API / クレート |
|---|---|
| ConPTY 接続 | `CreatePseudoConsole`, `ReadFile`, `WriteFile` |
| VT パース | `alacritty_terminal::Term`, `alacritty_terminal::event::EventListener` |
| グリッドバッファ | `alacritty_terminal::grid::Grid`, `alacritty_terminal::term::cell::Cell` |
| テキストストア | `ITextStoreACP`, `ITextStoreACPSink` |
| Direct2D 描画 | `ID2D1Factory`, `ID2D1HwndRenderTarget`, `ID2D1SolidColorBrush` |
| DirectWrite テキスト | `IDWriteFactory`, `IDWriteTextFormat`, `IDWriteTextLayout` |
| ウィンドウ管理 | `CreateWindowExW`, `DefWindowProcW`, `GetMessageW` |
| リサイズ | `ResizePseudoConsole`, `WM_SIZE` |
| キーボード入力 | `WM_KEYDOWN`, `WM_CHAR` → ConPTY WriteFile |
| IME composition | `ITfContextOwnerCompositionSink`, `ITextStoreACP::InsertTextAtSelection/SetText` |
| IME 候補位置 | `ImmSetCandidateWindow`, `ImmSetCompositionWindow`, `ITextStoreACP::GetTextExt` |
