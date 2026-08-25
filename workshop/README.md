# できたのかけらワークショップ

「私、ちゃんと進んでた」に気づく90分 ― できたのかけらワークショップ（9月9日開催）の告知・登壇用スライド（全29枚）。

## 完成物

- `dist/dekitanokakera-workshop.pptx` — 本体（PowerPoint、16:9）
- `dist/dekitanokakera-workshop.pdf` — 閲覧用PDF

## 構成

```
scripts/build.js  pptxgenjsでスライドを生成するスクリプト
dist/             生成済みファイル（ビルド成果物）
```

## 再生成

```
npm install
npm run build
```

`dist/dekitanokakera-workshop.pptx` が再生成されます。内容を修正する場合は `scripts/build.js` 内の各スライドのテキスト・レイアウトを編集してください。
