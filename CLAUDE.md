# taaminalu - TSF対応ターミナルエミュレータ

TSF `ITextStoreACP` を実装し、ターミナルバッファの内容を IME に公開するターミナルエミュレータ。
Rust + Win32 (Direct2D/DirectWrite) で実装。WSL2 に ConPTY 経由で接続。

## 目的

既存のターミナルアプリは TSF テキストストアの読み取りに対応していない（`GetText` が空を返す）。
rtry (Try-Code IME) のストロークヘルプ等、カーソル位置の文字を読み取る IME 機能が動作しない。
このターミナルは IME が `GetText` でバッファ内容を読み取れるようにする。

## アーキテクチャ

```
WSL2 shell/tmux
    ↕ stdin/stdout
wsl.exe (ConPTY 経由で起動)
    ↕ ReadFile/WriteFile
alacritty_terminal (VT パース + グリッドバッファ)
    ↓                ↓
ITextStoreACP     Direct2D/DirectWrite
(TSF テキストストア)   (レンダリング)
```

## 主要コンポーネント

| モジュール | 責務 |
|---|---|
| pty | ConPTY で wsl.exe を起動、読み書き |
| term | `alacritty_terminal` のラッパー。グリッドバッファの管理 |
| tsf | `ITextStoreACP` 実装。グリッドバッファの内容を IME に公開 |
| render | Direct2D + DirectWrite でターミナル描画 |
| window | Win32 ウィンドウ作成・メッセージループ |

## ビルド・実行

```
cargo build --release
cargo run --release
```

## 開発ルール

### コミット
- コミットメッセージは日本語で書く

### コーディング
- KISS・DRY
- edition 2024、最新の crate バージョン
- コードを書いたら都度レビュー・リファクタリング

### 調査・デバッグ
- **推測でコードを書くな**。まずコードを読んで原因を追ってから修正する
- **安易な解決策を採用するな**。類似プロダクトの実装やベストプラクティスを調べてから実装する
- **デバッグは上流から下流へ**: 関数が呼ばれているか → 引数は正しいか → ロジックは正しいか
- **前提を検証してから実装する**: 「呼ばれるはず」等の前提はログで検証してから次に進む
- 「可能性がある」ではなく、確実に原因を特定してから修正する
- windows crate の API は `~/.cargo/registry/src/` 内のソースを grep して確認する

## 技術スタック

| 用途 | クレート/API |
|---|---|
| ConPTY | `windows` crate (`CreatePseudoConsole`, `ReadFile`, `WriteFile`) |
| VT パース + バッファ | `alacritty_terminal` (0.25+) |
| レンダリング | `windows` crate (Direct2D 1.1, DirectWrite) |
| TSF テキストストア | `windows` crate (`ITextStoreACP`, `#[implement]` マクロ) |
| ウィンドウ | `windows` crate (Win32 `CreateWindowExW`, メッセージループ) |

## ITextStoreACP 実装の要点

26 メソッドの実装が必要（大半はスタブで可）。重要なメソッド:

| メソッド | 役割 | 実装方針 |
|---|---|---|
| `GetText` | **核心**。指定範囲のテキストを返す | alacritty_terminal の Grid からカーソル行周辺のテキストを返す |
| `GetSelection` | カーソル位置を返す | ターミナルのカーソル位置を ACP に変換 |
| `GetTextExt` | テキスト範囲の画面座標を返す | IME 候補ウィンドウの位置決めに必要 |
| `GetScreenExt` | ウィンドウの画面座標を返す | ウィンドウ RECT を返す |
| `GetEndACP` | テキスト末尾位置を返す | バッファ全体の文字数 |
| `RequestLock` | ロック管理 | 正しく実装しないと TSF マネージャーがクラッシュ |
| `AdviseSink` | TSF マネージャーのシンク登録 | シンクを保持して変更通知に使う |
| `SetText` | IME がテキストを挿入 | ConPTY への入力として転送 |
| `GetStatus` | ドキュメント属性を返す | `TS_SD_READONLY` は返さない（入力可能） |

## ConPTY の基本フロー

1. `CreatePipe()` で入出力パイプ作成
2. `CreatePseudoConsole(size, input, output, 0)` で ConPTY 作成
3. `STARTUPINFOEXW` + `PROC_THREAD_ATTRIBUTE_PSEUDOCONSOLE` で `wsl.exe` 起動
4. 読み取りスレッドで `ReadFile` → alacritty_terminal でパース → バッファ更新
5. キー入力を `WriteFile` で ConPTY に送信
6. `ResizePseudoConsole` でリサイズ対応

## 重要な罠

### Edition 2024
- `ManuallyDrop` union フィールドへの書き込みに `(*field)` が必要
- `#[unsafe(no_mangle)]` が必要
- `unsafe fn` 本体内でも `unsafe {}` ブロックが必要

### windows crate 0.62
- COM メソッドの引数型は `Ref<'_, T>`（`Option<&T>` ではない）

### ConPTY
- ReadFile/WriteFile は OVERLAPPED 非対応（同期 I/O のみ）
- 読み取りと書き込みは別スレッドで処理（デッドロック防止）

## 類似プロジェクト（参考）

| プロジェクト | 参考ポイント |
|---|---|
| [Alacritty](https://github.com/alacritty/alacritty) | ConPTY 接続、alacritty_terminal の使い方、ターミナルバッファ設計 |
| [WezTerm](https://github.com/wezterm/wezterm) | ターミナルバッファ、Windows 描画 |
| [Windows Terminal](https://github.com/microsoft/terminal) | ConPTY、TSF/IME 統合、Direct3D+DirectWrite 描画 |
| [Chromium TSF](https://chromium.googlesource.com/chromium/src/+/lkgr/ui/base/ime/win/tsf_text_store.h) | `ITextStoreACP` 実装のリファレンス |
