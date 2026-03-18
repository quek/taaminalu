# taaminalu

TSF 対応ターミナルエミュレータ for Windows

## 概要

**taaminalu** は、TSF (Text Services Framework) の `ITextStoreACP` を実装し、ターミナルバッファの内容を IME に公開するターミナルエミュレータです。

既存のターミナルアプリ（Windows Terminal, Alacritty 等）は TSF テキストストアの読み取りに対応しておらず、`GetText` が空を返します。そのため、カーソル位置の文字を読み取る IME 機能（[rtry](https://github.com/jmkgit/rtry) のストロークヘルプ等）が動作しません。

taaminalu は IME が `GetText` でターミナルバッファの内容を読み取れるようにします。

## 特徴

- **TSF テキストストア対応** — IME がターミナル画面のテキストを読み取り可能
- **タブ機能** — WSL / CMD / PowerShell を選んで複数セッションを同時利用
- **Direct2D / DirectWrite 描画** — GPU アクセラレーションによる高品質レンダリング
- **ConPTY 接続** — Windows の Pseudo Console API で WSL2 やコマンドプロンプトに接続
- **256 色 + ANSI カラー対応** — Campbell パレット準拠

## スクリーンショット

(TODO)

## 動作環境

- Windows 10 (1809+) / Windows 11
- WSL2（WSL タブを使う場合）
- [HackGen Console NF](https://github.com/yuru7/HackGen) フォント（推奨）

## ビルド

```
cargo build --release
```

## 実行

```
cargo run --release
```

## キーボードショートカット

### タブ操作

| ショートカット | 動作 |
|---|---|
| `Ctrl+Shift+T` | 新規タブ（シェル選択メニュー） |
| `Ctrl+Shift+W` | アクティブタブを閉じる |
| `Ctrl+Tab` | 次のタブに切替 |
| `Ctrl+Shift+Tab` | 前のタブに切替 |

### 編集

| ショートカット | 動作 |
|---|---|
| `Ctrl+Shift+V` | クリップボードからペースト |

## アーキテクチャ

```
WSL2 / CMD / PowerShell
    ↕ stdin/stdout
wsl.exe / cmd.exe (ConPTY 経由で起動)
    ↕ ReadFile/WriteFile
alacritty_terminal (VT パース + グリッドバッファ)
    ↓                ↓
ITextStoreACP     Direct2D/DirectWrite
(TSF テキストストア)   (レンダリング)
```

### モジュール構成

| モジュール | 責務 |
|---|---|
| `pty` | ConPTY でシェルプロセスを起動、読み書き |
| `term` | `alacritty_terminal` のラッパー。グリッドバッファの管理 |
| `tab` | タブ単位の状態管理（PTY + Term） |
| `app` | アプリケーション全体の状態、複数タブの管理 |
| `tsf` | `ITextStoreACP` 実装。グリッドバッファの内容を IME に公開 |
| `render` | Direct2D + DirectWrite でターミナル描画、タブバー UI |
| `window` | Win32 ウィンドウ作成・メッセージループ・キー入力処理 |

## 技術スタック

| 用途 | 技術 |
|---|---|
| 言語 | Rust (Edition 2024) |
| ConPTY | `windows` crate (`CreatePseudoConsole`) |
| VT パース + バッファ | `alacritty_terminal` |
| レンダリング | Direct2D 1.1 + DirectWrite |
| TSF テキストストア | `ITextStoreACP` (`windows` crate `#[implement]`) |
| ウィンドウ | Win32 API (`CreateWindowExW`) |

## ライセンス

[MIT](LICENSE)
