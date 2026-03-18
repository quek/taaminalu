---
name: research-similar-impl
description: |
  類似ターミナルエミュレータ（Alacritty, WezTerm, Windows Terminal）のソースコードと
  公式 API リファレンスを調査し、実装方針レポートを出力する。
  「実装して」「追加して」「修正して」「対応して」「機能を作って」「バグを直して」等、
  コード変更を伴う指示があったとき、または Windows API / COM の使い方が不明なときに発動。
  調査のみ行い、コードの編集は行わない。
argument-hint: "[調査対象の機能名]"
allowed-tools: Bash(git clone *), Bash(git pull *), Read, Grep, Glob, WebSearch, WebFetch, Agent
---

# 類似プロダクト & API リファレンス調査

$ARGUMENTS に関する調査を行い、taaminaru での実装方針を立てるためのレポートを出力する。

## 手順

### 1. 調査対象の特定

ユーザーの要求から実装対象の機能・Windows API を特定する。
[references.md](references.md) の「機能と API の対応例」を参照。

### 2. リポジトリのクローン

[references.md](references.md) の調査対象プロジェクトから、機能に関連するものを `/tmp` にクローンする。

```bash
[ -d /tmp/alacritty ] || git clone --depth 1 https://github.com/alacritty/alacritty.git /tmp/alacritty
[ -d /tmp/wezterm ]   || git clone --depth 1 https://github.com/wezterm/wezterm.git /tmp/wezterm
```

- クローン先は `/tmp` 配下（作業ディレクトリを汚さない）
- `--depth 1` で軽量クローン
- 既存ならスキップ

### 3. 並列調査（Agent を並列起動）

以下の A) と B) を **Agent を並列起動して同時に** 実行する。

**A) 類似プロダクトのソースコード調査**

クローン済みリポジトリを Grep / Read で横断検索。調査ポイント:
- API の呼び出しパターン（シグネチャ、引数、戻り値）
- 設計パターン（状態管理、スレッドモデル）
- ConPTY の使い方（パイプ管理、リサイズ、終了処理）
- alacritty_terminal の使い方（Term, Grid, EventListener）
- Direct2D/DirectWrite の初期化・描画パターン
- TSF/ITextStoreACP の実装パターン

**B) 公式 API リファレンス・ガイド調査**

[references.md](references.md) の API ドキュメント URL を WebFetch / WebSearch で調査:
- 該当 API の公式仕様・制約
- 使用上の注意点やベストプラクティス
- サンプルコード

### 4. windows crate の API 確認

C++ の API と Rust バインディングでシグネチャが異なるため、
`~/.cargo/registry/src/` 内のソースを Grep して実際の Rust シグネチャを確認する。

確認ポイント:
- COM メソッドの引数型（`Ref<'_, T>` vs `Option<&T>` vs 生ポインタ）
- `Result<T>` のラッピング
- `ManuallyDrop` / `VARIANT` の扱い
- 定数名の違い

### 5. レポート出力

[report-template.md](report-template.md) の形式で日本語でまとめる。

## 制約

- **調査のみ**。ファイルの編集・作成・ビルド・インストール等は一切行わない
- Alacritty は alacritty_terminal の使い方を知る上で最優先で参照する
- windows crate の API 確認は省略しない
