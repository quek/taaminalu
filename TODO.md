# TODO

## 描画属性

- [x] UNDERCURL を波線で描画（PathGeometry + QuadraticBezier）
- [x] DOTTED_UNDERLINE を点線で描画（StrokeStyle DOT + DrawLine）
- [x] DASHED_UNDERLINE を破線で描画（StrokeStyle DASH + DrawLine）
- [x] 下線色（ESC[58;2;R;G;Bm）の対応（※ConPTY が SGR 58 を除去するため現状未動作）

## ターミナルモード

- [ ] APP_CURSOR (DECCKM): 矢印キーで ESC O A 形式を送信（ESC[?1h/l）
- [ ] APP_KEYPAD (DECKPAM): テンキーのアプリケーションモード
- [ ] Bracketed Paste: 貼り付け時に ESC[200~ / ESC[201~ で囲む（ESC[?2004h/l）
- [ ] Focus Events: フォーカス取得/喪失時に ESC[I / ESC[O を送信（ESC[?1004h/l）

## マウス

- [ ] マウスレポート（ESC[?1000h, SGR マウス等）
- [x] テキスト選択（マウスドラッグ → クリップボードにコピー）— 9850ed6
- [x] ミドルクリックでペースト — 9850ed6
- [x] マウスホイールスクロール（スクロールバック）— 47ce55f
- [x] ダブルクリックで単語選択 — f67a4a0

## OSC シーケンス

- [x] OSC 52: クリップボード書き込み（tmux copy-pipe 対応）— f95737c
- [ ] ウィンドウタイトル変更（OSC 0/1/2）
- [ ] カラーパレット変更（OSC 4）
- [ ] ハイパーリンク（OSC 8）

## その他

- [ ] BEL 文字（0x07）でビープ音 or 視覚ベル
- [ ] URL 検出・クリックで開く
