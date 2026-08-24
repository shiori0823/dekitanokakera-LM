# できたのかけら診断

LINEプレゼント用リードマグネット。「今日も何もできなかった」と感じる人に向けた、8問のセルフ診断PDF（全8ページ）。

## 完成物

- `dist/dekitanokakera-shindan.pdf` — 配布用PDF本体（1080×1350px／全8ページ）

## 構成

```
src/index.html    診断本文（全8ページ分のマークアップ）
src/style.css     デザイン（配色・タイポグラフィ・レイアウト）
assets/fonts/     Zen Maru Gothic（本文で使用する文字のみを含むサブセット）
scripts/render.js Playwrightでsrc/index.htmlをPDF化するビルドスクリプト
dist/             生成済みPDF（ビルド成果物）
```

## PDFの再生成

```
npm install
npm run build
```

Chromiumの実行ファイルパスは環境変数 `CHROMIUM_PATH` で上書きできます（未指定時は `/opt/pw-browsers/chromium`）。

## 内容を修正するとき

`src/index.html` の本文・`src/style.css` の見た目を編集し、`npm run build` で `dist/dekitanokakera-shindan.pdf` を再生成してください。
