const path = require("path");
const pptxgen = require("pptxgenjs");

// ============================= palette =============================
const C = {
  bgLight: "FFF8F0",
  bgDark: "2A211B",
  card: "FFFFFF",
  cardWarm: "FFFDFA",
  line: "EFE2D3",
  ink: "4A3B30",
  inkSoft: "8A7A6C",
  inkFaint: "B7A99A",
  creamText: "FBEFE4",
  creamSoft: "D8C4B3",
  coral: "F0876A",
  coralDark: "DD6E4E",
  coralBg: "FDEEE7",
  gold: "CE9A2E",
  goldBg: "FFF3DA",
  green: "5E9E5B",
  greenBg: "E9F5E7",
  teal: "2F9188",
  tealBg: "E1F4F1",
  indigo: "6B74C4",
  indigoBg: "ECEDFB",
  white: "FFFFFF",
  mute: "C9BBAC",
  muteBg: "F3ECE3",
};

const FONT = "Meiryo";
const FONT_TITLE = "Meiryo";

const W = 13.333;
const H = 7.5;
const MX = 0.7;
const CW = W - MX * 2;

// ============================= helpers =============================

function newSlide(pres, bg) {
  const slide = pres.addSlide();
  slide.background = { color: bg || C.bgLight };
  return slide;
}

function shard(slide, x, y, size, color, rotate, opacity) {
  slide.addShape("hexagon", {
    x, y, w: size, h: size,
    rotate: rotate || 0,
    fill: { color, transparency: opacity === undefined ? 88 : opacity },
    line: { type: "none" },
  });
}

function shardsCover(slide) {
  shard(slide, -1.1, -1.0, 3.2, C.coral, 12, 78);
  shard(slide, 10.6, -1.3, 2.6, C.gold, -18, 80);
  shard(slide, 11.6, 5.6, 3.0, C.indigo, 20, 82);
  shard(slide, -1.4, 5.4, 2.4, C.teal, -10, 82);
  shard(slide, 5.6, 6.6, 1.4, C.green, 30, 85);
}
function shardsLast(slide) {
  shard(slide, -1.2, -1.1, 3.0, C.gold, 10, 80);
  shard(slide, 11.0, 5.3, 3.2, C.coral, -14, 78);
  shard(slide, 11.2, -1.2, 2.0, C.teal, 22, 85);
}
function shardCorner(slide, corner, color) {
  const pos = {
    tr: [12.15, -0.75, 2.0, 15],
    br: [12.2, 6.15, 2.0, -12],
    bl: [-0.95, 6.15, 1.9, 12],
    tl: [-0.9, -0.8, 1.9, -10],
  }[corner];
  shard(slide, pos[0], pos[1], pos[2], color, pos[3], 87);
}

function folio(slide, num, dark) {
  slide.addText("できたのかけらワークショップ", {
    x: MX, y: H - 0.48, w: 6, h: 0.32,
    fontFace: FONT, fontSize: 10, color: dark ? C.creamSoft : C.inkFaint,
    align: "left", margin: 0,
  });
  slide.addText(`${num} / 29`, {
    x: W - MX - 1.0, y: H - 0.48, w: 1.0, h: 0.32,
    fontFace: FONT, fontSize: 10, color: dark ? C.creamSoft : C.inkFaint,
    align: "right", margin: 0,
  });
}

function eyebrow(slide, text, opts) {
  const o = Object.assign({
    x: MX, y: 0.5, w: 10, h: 0.35,
    fontFace: FONT, fontSize: 13, bold: true, color: C.coralDark,
    charSpacing: 3, margin: 0,
  }, opts || {});
  slide.addText(text, o);
}

function title(slide, text, opts) {
  const o = Object.assign({
    x: MX, y: 0.88, w: CW, h: 1.0,
    fontFace: FONT_TITLE, fontSize: 28, bold: true, color: C.ink,
    align: "left", margin: 0, lineSpacing: 34,
  }, opts || {});
  slide.addText(text, o);
}

function iconCircle(slide, emoji, x, y, d, bg, fontSize) {
  slide.addShape("ellipse", { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" } });
  slide.addText(emoji, {
    x, y, w: d, h: d, align: "center", valign: "middle",
    fontSize: fontSize || Math.round(d * 22), margin: 0,
  });
}

function card(slide, x, y, w, h, opts) {
  const o = Object.assign({
    x, y, w, h, rectRadius: 0.12,
    fill: { color: C.cardWarm },
    line: { color: C.line, width: 1 },
  }, opts || {});
  slide.addShape("roundRect", o);
}

function pill(slide, text, x, y, w, h, color, bgColor, fontSize) {
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: h / 2,
    fill: { color: bgColor }, line: { type: "none" },
  });
  slide.addText(text, {
    x, y, w, h, align: "center", valign: "middle",
    fontFace: FONT, fontSize: fontSize || 12, bold: true, color, margin: 0,
  });
}

function tagline(slide, text, x, y, w, opts) {
  const o = Object.assign({
    x, y, w, h: 0.6, align: "left", valign: "middle",
    fontFace: FONT, fontSize: 16, bold: true, color: C.coralDark, margin: 0, lineSpacing: 21,
  }, opts || {});
  slide.addText(text, o);
}

function closingBanner(slide, text, y) {
  const yy = y === undefined ? H - 1.22 : y;
  card(slide, MX, yy, CW, 0.72, { fill: { color: C.coralBg }, line: { type: "none" }, rectRadius: 0.14 });
  slide.addText(text, {
    x: MX + 0.35, y: yy, w: CW - 0.7, h: 0.72, valign: "middle",
    fontFace: FONT, fontSize: 16, bold: true, color: C.coralDark, margin: 0, lineSpacing: 21,
  });
}

function bodyText(slide, text, x, y, w, h, opts) {
  const o = Object.assign({
    x, y, w, h,
    fontFace: FONT, fontSize: 14, color: C.ink, margin: 0, lineSpacing: 22, valign: "top",
  }, opts || {});
  slide.addText(text, o);
}

// ============================= build =============================

const pres = new pptxgen();
pres.defineLayout({ name: "WIDE", width: W, height: H });
pres.layout = "WIDE";

let PAGE = 2; // slide 1 is the cover (no folio shown there)
function P() { return PAGE++; }

// ---------------------------------------------------------------
// 1. COVER
// ---------------------------------------------------------------
{
  const s = newSlide(pres, C.bgDark);
  shardsCover(s);
  iconCircle(s, "🌱", 5.92, 0.62, 0.85, "3A2E26", 34);
  s.addText("9月9日（水）10:00〜11:30　｜　オンライン（Zoom）　｜　参加費：無料", {
    x: 0.8, y: 1.68, w: 11.73, h: 0.4, align: "center",
    fontFace: FONT, fontSize: 14, bold: true, color: C.gold, charSpacing: 1, margin: 0,
  });
  s.addText("「私、ちゃんと進んでた」に気づく90分", {
    x: 0.8, y: 2.5, w: 11.73, h: 0.85, align: "center",
    fontFace: FONT_TITLE, fontSize: 32, bold: true, color: C.white, margin: 0,
  });
  s.addText("できたのかけらワークショップ", {
    x: 0.8, y: 3.42, w: 11.73, h: 0.95, align: "center",
    fontFace: FONT_TITLE, fontSize: 44, bold: true, color: C.coral, margin: 0,
  });
  s.addText(
    "「頑張ってるはずなのに、なんだか進んでない気がする…」\nそんな方と一緒に、自分の中にもうある小さな「できた」を見つけて、\n「あれ？私、ちゃんと進んでた！」に気づく90分にしたいと思っています",
    {
      x: 1.4, y: 4.78, w: 10.53, h: 1.4, align: "center",
      fontFace: FONT, fontSize: 15, color: C.creamText, lineSpacing: 25, margin: 0,
    }
  );
  s.addText("できたのかけら診断 ── ワークショップ編", {
    x: 0.8, y: H - 0.62, w: 11.73, h: 0.35, align: "center",
    fontFace: FONT, fontSize: 11, color: C.creamSoft, charSpacing: 2, margin: 0,
  });
}

// ---------------------------------------------------------------
// 2. S01｜今日も何もできなかったと思っていませんか
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "tr", C.coral);
  eyebrow(s, "OPENING");
  title(s, "「今日も何もできなかった」と\n思っていませんか？", { fontSize: 27, h: 1.3 });
  bodyText(s, "頑張っているのに、なぜか前に進んでいる気がしない。", MX, 2.15, CW, 0.4, { fontSize: 16, color: C.inkSoft, bold: true });

  const items = [
    ["🎯", "目標を立てても\n途中で止まる"],
    ["😔", "「またできなかった」と\n自分を責める"],
    ["👀", "周りの「できた」ばかり\n目に入る"],
  ];
  const cw = (CW - 0.5) / 3;
  items.forEach((it, i) => {
    const x = MX + i * (cw + 0.25);
    const y = 2.75;
    card(s, x, y, cw, 1.5);
    iconCircle(s, it[0], x + cw / 2 - 0.35, y + 0.22, 0.7, C.coralBg, 26);
    s.addText(it[1], {
      x: x + 0.15, y: y + 1.0, w: cw - 0.3, h: 0.45, align: "center",
      fontFace: FONT, fontSize: 12.5, color: C.ink, margin: 0, lineSpacing: 16,
    });
  });

  card(s, MX, 4.55, CW, 1.95, { fill: { color: C.coralBg }, line: { type: "none" } });
  s.addText([
    { text: "でも本当に、あなたは「何もできていない」のでしょうか？", options: { fontSize: 17, bold: true, color: C.coralDark, breakLine: true } },
    { text: "今日の90分で見つけるのは、あなたの中に、もうある", options: { fontSize: 15, color: C.ink, breakLine: false } },
    { text: "「できたのかけら」", options: { fontSize: 18, bold: true, color: C.coralDark, breakLine: true } },
    { text: "小さな「できた」が、未来への道しるべになる。", options: { fontSize: 12.5, color: C.inkSoft, italic: true } },
  ], {
    x: MX + 0.4, y: 4.55, w: CW - 0.8, h: 1.95, valign: "middle", margin: 0,
    paraSpaceAfter: 10,
  });
  folio(s, P());
}
// ---------------------------------------------------------------
// 3. S02a｜私も「何もしていない」と思っていた（日記）
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "tr", C.gold);
  eyebrow(s, "MY STORY");
  title(s, "私も「何もしていない」と思っていた");

  card(s, MX, 2.0, CW, 1.95, { fill: { color: C.card } });
  pill(s, "2025年11月5日の日記", MX + 0.4, 2.28, 2.6, 0.42, C.white, C.gold, 13);
  s.addText(
    [
      { text: "「欲しいのは、お金？多分違う。」", options: { breakLine: true } },
      { text: "「私の口癖、何もしていない。」", options: {} },
    ],
    {
      x: MX + 0.4, y: 2.85, w: CW - 0.8, h: 0.95, valign: "top",
      fontFace: FONT, fontSize: 19, bold: true, color: C.ink, margin: 0, paraSpaceAfter: 8,
    }
  );

  const tags = ["子育て", "ジム経営", "パワーリフティング"];
  let tx = MX;
  tags.forEach((t) => {
    const tw = 0.5 + t.length * 0.28;
    pill(s, t, tx, 4.2, tw, 0.44, C.teal, C.tealBg, 13);
    tx += tw + 0.2;
  });
  bodyText(s, "毎日やることはたくさんある。", MX, 4.85, CW, 0.4, { fontSize: 15, color: C.inkSoft });

  closingBanner(s, "それでも一日の終わりには、\n「今日も何もできなかった。」", 5.55);
  folio(s, P());
}

// ---------------------------------------------------------------
// 4. S02b｜気づき
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "bl", C.coral);
  iconCircle(s, "✨", W / 2 - 0.45, 0.55, 0.9, C.coralBg, 34);
  s.addText("でも、本当に欲しかったのは、\nお金でも、完璧な自分でもありませんでした。", {
    x: 1.2, y: 1.7, w: W - 2.4, h: 1.0, align: "center",
    fontFace: FONT, fontSize: 17, color: C.ink, margin: 0, lineSpacing: 26,
  });

  card(s, 2.4, 2.85, W - 4.8, 1.15, { fill: { color: C.coralBg }, line: { type: "none" } });
  s.addText([
    { text: "「私にもできた。」", options: { fontSize: 26, bold: true, color: C.coralDark, breakLine: true } },
    { text: "そう、自分自身で思えること。", options: { fontSize: 14, color: C.ink } },
  ], { x: 2.4, y: 2.85, w: W - 4.8, h: 1.15, align: "center", valign: "middle", margin: 0, paraSpaceAfter: 6 });

  s.addText([
    { text: "そして気づきました。", options: { fontSize: 15, color: C.inkSoft, breakLine: true } },
    { text: "何もしていなかったんじゃない。", options: { fontSize: 19, bold: true, color: C.ink, breakLine: true } },
    { text: "自分の「できた」に気づいていなかっただけ。", options: { fontSize: 19, bold: true, color: C.coralDark } },
  ], { x: 1.2, y: 4.4, w: W - 2.4, h: 1.6, align: "center", margin: 0, paraSpaceAfter: 10 });

  folio(s, P());
}

// ---------------------------------------------------------------
// 5. S03｜続かないのは「意志」のせいじゃない
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "tr", C.indigo);
  eyebrow(s, "なぜ続かないのか");
  title(s, "続かないのは「意志」のせいじゃない");
  bodyText(s, "こんな経験、ありませんか？", MX, 1.85, CW, 0.4, { fontSize: 15, color: C.inkSoft, bold: true });

  const reasons = [
    ["①", "目標が大きすぎる", "「毎日30分やる！」\n↓\nできない日が来る\n↓\n0になる"],
    ["②", "できなかったことばかり見る", "9個できても、\n1個できなかったら、\n「今日もダメだった」"],
    ["③", "止まる＝失敗になっている", "1日できない\n↓\n「また続かなかった」\n↓\n全部やめる"],
  ];
  const cw = (CW - 0.5) / 3;
  reasons.forEach((r, i) => {
    const x = MX + i * (cw + 0.25);
    const y = 2.4;
    card(s, x, y, cw, 2.65);
    pill(s, r[0], x + 0.25, y + 0.25, 0.5, 0.5, C.white, C.indigo, 18);
    s.addText(r[1], {
      x: x + 0.25, y: y + 0.9, w: cw - 0.5, h: 0.65,
      fontFace: FONT, fontSize: 15, bold: true, color: C.ink, margin: 0, lineSpacing: 19,
    });
    s.addText(r[2], {
      x: x + 0.25, y: y + 1.55, w: cw - 0.5, h: 1.0,
      fontFace: FONT, fontSize: 12.5, color: C.inkSoft, margin: 0, lineSpacing: 17, align: "center",
    });
  });

  closingBanner(s, "だから必要なのは、もっと頑張ることではありません。「戻り方」を知ること。");
  folio(s, P());
}

// ---------------------------------------------------------------
// 6. S03A｜目標達成を遠ざける3つの思い込み
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "tr", C.gold);
  title(s, "目標達成を遠ざける3つの思い込み");

  const rows = [
    ["「続けられる人は、意志が強い」", "意志ではなく、続けやすい仕組みをつくる。"],
    ["「私は続けられない人」", "続かなかった方法が、あなたに合っていなかっただけかもしれない。"],
    ["「忙しいから今は無理」", "時間ができるのを待つのではなく、今の生活に入る大きさにする。"],
  ];
  const rowH = 0.92, gap = 0.14, top = 1.95;
  const leftW = 4.5, arrowW = 0.6, rightW = CW - leftW - arrowW;
  rows.forEach((r, i) => {
    const y = top + i * (rowH + gap);
    card(s, MX, y, leftW, rowH, { fill: { color: C.muteBg }, line: { type: "none" } });
    s.addText([
      { text: "❌ 思い込み\n", options: { fontSize: 11, bold: true, color: C.inkFaint, breakLine: true } },
      { text: r[0], options: { fontSize: 14, bold: true, color: C.inkSoft } },
    ], { x: MX + 0.3, y, w: leftW - 0.6, h: rowH, valign: "middle", margin: 0 });

    s.addText("→", {
      x: MX + leftW, y, w: arrowW, h: rowH, align: "center", valign: "middle",
      fontFace: FONT, fontSize: 20, color: C.coral, margin: 0,
    });

    const rx = MX + leftW + arrowW;
    card(s, rx, y, rightW, rowH, { fill: { color: C.coralBg }, line: { type: "none" } });
    s.addText([
      { text: "✅ 真実\n", options: { fontSize: 11, bold: true, color: C.coral, breakLine: true } },
      { text: r[1], options: { fontSize: 14, bold: true, color: C.coralDark } },
    ], { x: rx + 0.3, y, w: rightW - 0.6, h: rowH, valign: "middle", margin: 0, lineSpacing: 18 });
  });

  s.addText("必要なのは「もっと頑張れる私」ではない。今の私でも進める方法をつくること。", {
    x: MX, y: top + 3 * (rowH + gap) + 0.05, w: CW, h: 0.5, align: "center",
    fontFace: FONT, fontSize: 15, bold: true, color: C.ink, margin: 0,
  });
  folio(s, P());
}

// ---------------------------------------------------------------
// 7. S04A｜目標にも「今持てる重さ」がある
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "bl", C.teal);
  title(s, "目標にも「今持てる重さ」がある");
  bodyText(s, "いきなり100kgのバーベルを渡されて、「頑張って持って！」と言われても、持てません。", MX, 1.85, CW, 0.55, { fontSize: 15, color: C.ink, lineSpacing: 21 });

  const weights = [
    ["100kg", "そのままでは持てない", C.mute, C.muteBg],
    ["20kg", "なら？", C.teal, C.tealBg],
    ["22.5kg", "なら？", C.teal, C.tealBg],
    ["＋2.5kg", "昨日より、なら？", C.coral, C.coralBg],
  ];
  const cw = (CW - 0.6) / 4;
  weights.forEach((w, i) => {
    const x = MX + i * (cw + 0.2);
    const y = 2.65;
    card(s, x, y, cw, 1.5, { fill: { color: w[3] }, line: { type: "none" } });
    s.addText(w[0], {
      x, y: y + 0.22, w: cw, h: 0.6, align: "center",
      fontFace: FONT, fontSize: 24, bold: true, color: w[2], margin: 0,
    });
    s.addText(w[1], {
      x, y: y + 0.95, w: cw, h: 0.45, align: "center",
      fontFace: FONT, fontSize: 12, color: C.inkSoft, margin: 0,
    });
  });

  s.addText("目標も同じです。", {
    x: MX, y: 4.35, w: CW, h: 0.4, fontFace: FONT, fontSize: 15, bold: true, color: C.ink, margin: 0,
  });
  s.addText([
    { text: "❌「できない私はダメ」", options: { fontSize: 15, color: C.inkFaint, breakLine: true } },
    { text: "ではなく、", options: { fontSize: 13, color: C.inkSoft, breakLine: true } },
    { text: "✅「今の私には重すぎない？」と考える。", options: { fontSize: 16, bold: true, color: C.teal } },
  ], { x: MX, y: 4.8, w: CW, h: 1.0, margin: 0, paraSpaceAfter: 4 });

  closingBanner(s, "大きな目標はそのままでいい。今日持つ重さだけ、小さくする。");
  folio(s, P());
}

// ---------------------------------------------------------------
// 8. S04B｜「いつか」より、今日ひとつ
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "tr", C.coral);
  title(s, "「いつか」より、今日ひとつ");
  s.addText([
    { text: "大きな目標を見ると、「時間ができたら」「もう少し余裕ができたら」「ちゃんと準備できたら」と思ってしまう。", options: { fontSize: 14, color: C.ink, breakLine: true } },
    { text: "でも、“いつか”は、なかなか始まらない。", options: { fontSize: 14, bold: true, color: C.coralDark } },
  ], { x: MX, y: 1.85, w: CW, h: 1.0, margin: 0, lineSpacing: 21, paraSpaceAfter: 8 });

  s.addText("だから考えるのは、「今日の私なら、何ができる？」", {
    x: MX, y: 2.95, w: CW, h: 0.45, fontFace: FONT, fontSize: 17, bold: true, color: C.ink, margin: 0,
  });

  const steps = ["30分", "10分", "3分", "準備するだけ"];
  const sizes = [1.4, 1.15, 0.9, 1.5];
  let sx = MX;
  const baseY = 3.75, rowH = 1.5;
  steps.forEach((label, i) => {
    const sw = i === 3 ? 2.5 : 1.7;
    const sh = sizes[i];
    const y = baseY + (rowH - sh);
    card(s, sx, y, sw, sh, { fill: { color: i === 3 ? C.coralBg : C.tealBg }, line: { type: "none" } });
    s.addText(label, {
      x: sx, y, w: sw, h: sh, align: "center", valign: "middle",
      fontFace: FONT, fontSize: i === 3 ? 16 : 18, bold: true, color: i === 3 ? C.coralDark : C.teal, margin: 0,
    });
    sx += sw + 0.25;
    if (i < steps.length - 1) {
      s.addText("→", {
        x: sx - 0.22, y: baseY + rowH - 0.55, w: 0.3, h: 0.4, align: "center",
        fontFace: FONT, fontSize: 16, color: C.inkFaint, margin: 0,
      });
    }
  });

  closingBanner(s, "小さくするのは、夢ではない。今日の一歩。", 5.75);
  folio(s, P());
}

// helper for STEP cards used across slides 9-12
function stepCard(s, x, y, w, h, num, stepTitle, body, accent, accentBg) {
  card(s, x, y, w, h);
  pill(s, `STEP ${num}`, x + 0.3, y + 0.28, 1.3, 0.42, C.white, accent, 13);
  s.addText(stepTitle, {
    x: x + 0.3, y: y + 0.85, w: w - 0.6, h: 0.55,
    fontFace: FONT, fontSize: 18, bold: true, color: C.ink, margin: 0, lineSpacing: 22,
  });
  s.addText(body, {
    x: x + 0.3, y: y + 1.42, w: w - 0.6, h: h - 1.65,
    fontFace: FONT, fontSize: 13.5, color: C.inkSoft, margin: 0, lineSpacing: 19,
  });
}

// ---------------------------------------------------------------
// 9. S04｜7STEP intro + STEP1-2
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "tr", C.coral);
  eyebrow(s, "「できた」を結果につなげる7STEP　｜　STEP 1 - 2");
  title(s, "「できた」を結果につなげる7STEP", { fontSize: 26 });
  bodyText(s, "たとえば「3ヶ月で−3kg」を目指すなら？", MX, 1.72, CW, 0.4, { fontSize: 15, bold: true, color: C.coralDark });

  const cw = (CW - 0.3) / 2;
  stepCard(s, MX, 2.35, cw, 3.9, 1, "今の自分を知る",
    "仕事・生活・身体・食事。\n「この生活の中で何ならできる？」", C.coral);
  stepCard(s, MX + cw + 0.3, 2.35, cw, 3.9, 2, "すでにある「できた」を見つける",
    "朝食は食べられている。水分は取れている。\n週末は歩いている。\n0からのスタートではない。", C.coral);
  folio(s, P());
}

// ---------------------------------------------------------------
// 10. S04｜STEP3-4
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "tr", C.coral);
  eyebrow(s, "「できた」を結果につなげる7STEP　｜　STEP 3 - 4");
  title(s, "「できた」を結果につなげる7STEP", { fontSize: 26 });

  const cw = (CW - 0.3) / 2;
  stepCard(s, MX, 2.0, cw, 4.25, 3, "本当に叶えたい未来を決める",
    "−3kgの先にある、\n「お気に入りの服を着て出かけたい」\nまで考える。", C.coral);
  stepCard(s, MX + cw + 0.3, 2.0, cw, 4.25, 4, "今日できる大きさまで分解",
    "3ヶ月 → 1ヶ月 → 1週間 → 今日\n\n大きな目標はそのまま。\n今日の一歩だけ小さくする。", C.coral);
  folio(s, P());
}

// ---------------------------------------------------------------
// 11. S04｜STEP5
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "bl", C.coral);
  eyebrow(s, "「できた」を結果につなげる7STEP　｜　STEP 5");
  title(s, "「できた」を結果につなげる7STEP", { fontSize: 26 });

  card(s, MX, 2.0, CW, 1.05, { fill: { color: C.card } });
  pill(s, "STEP 5", MX + 0.3, 2.28, 1.3, 0.42, C.white, C.coral, 13);
  s.addText("身体・思考・習慣を見る", {
    x: MX + 1.8, y: 2.0, w: CW - 2.1, h: 1.05, valign: "middle",
    fontFace: FONT, fontSize: 20, bold: true, color: C.ink, margin: 0,
  });

  s.addText([
    { text: "夜、お菓子を食べた。", options: { fontSize: 15, color: C.ink, breakLine: true } },
    { text: "❌「私って意志が弱い」ではなく、", options: { fontSize: 15, bold: true, color: C.inkFaint } },
  ], { x: MX, y: 3.3, w: CW, h: 0.9, margin: 0, paraSpaceAfter: 4 });

  const boxes = [
    ["身体", "疲れていた？", C.coral, C.coralBg],
    ["思考", "我慢しすぎた？", C.gold, C.goldBg],
    ["習慣", "昼食が少なかった？", C.teal, C.tealBg],
  ];
  const bw = (CW - 0.6) / 3;
  boxes.forEach((b, i) => {
    const x = MX + i * (bw + 0.3);
    const y = 4.35;
    card(s, x, y, bw, 1.5, { fill: { color: b[3] }, line: { type: "none" } });
    s.addText(b[0], { x, y: y + 0.22, w: bw, h: 0.45, align: "center", fontFace: FONT, fontSize: 16, bold: true, color: b[2], margin: 0 });
    s.addText(b[1], { x, y: y + 0.8, w: bw, h: 0.5, align: "center", fontFace: FONT, fontSize: 14, color: C.ink, margin: 0 });
  });
  folio(s, P());
}

// ---------------------------------------------------------------
// 12. S04｜STEP6-7
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "tr", C.coral);
  eyebrow(s, "「できた」を結果につなげる7STEP　｜　STEP 6 - 7");
  title(s, "「できた」を結果につなげる7STEP", { fontSize: 26 });

  const cw = (CW - 0.3) / 2;
  stepCard(s, MX, 2.0, cw, 4.25, 6, "「できた」を記録",
    "10分歩けた。食事を調整できた。\n野菜を増やせた。\n食べすぎても翌日戻れた。\n\n結果が出る前の前進を、\nなかったことにしない。", C.coral);
  stepCard(s, MX + cw + 0.3, 2.0, cw, 4.25, 7, "止まったら戻る",
    "外食した。体重が増えた。\n↓\n❌「全部台無し！」ではなく、\n「楽しかった。じゃあ今日から戻ろう。」\n\n戻れたことも「できた」。", C.coral);
  folio(s, P());
}

// ---------------------------------------------------------------
// 13. S05｜CASE1 −8.6kg
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "tr", C.teal);
  eyebrow(s, "大きな結果は、小さな「できた」の集合体　｜　CASE 1 / 3");
  title(s, "CASE1｜−8.6kg", { fontSize: 27 });
  bodyText(s, "ダイエットサポートのお客様", MX, 1.68, CW, 0.4, { fontSize: 14, color: C.inkSoft, bold: true });

  card(s, MX, 2.2, 3.6, 4.05, { fill: { color: C.tealBg }, line: { type: "none" } });
  s.addText("−8.6kg", {
    x: MX, y: 2.2, w: 3.6, h: 1.6, align: "center", valign: "middle",
    fontFace: FONT, fontSize: 46, bold: true, color: C.teal, margin: 0,
  });
  s.addText("達成", {
    x: MX, y: 3.65, w: 3.6, h: 0.5, align: "center",
    fontFace: FONT, fontSize: 16, bold: true, color: C.teal, margin: 0,
  });
  s.addText("でも、変わったのは\n体重だけではありません。", {
    x: MX + 0.3, y: 4.55, w: 3.0, h: 1.5, align: "center",
    fontFace: FONT, fontSize: 14, color: C.ink, margin: 0, lineSpacing: 21,
  });

  const checks = [
    "階段を軽々登れるようになった",
    "足が攣らなくなった",
    "躊躇していた水泳を始めた",
    "自分が満足できる食事量が分かった",
    "「朝、身体が軽い！」",
  ];
  const lx = MX + 3.9;
  const lw = CW - 3.9;
  checks.forEach((t, i) => {
    const y = 2.25 + i * 0.72;
    card(s, lx, y, lw, 0.58, { fill: { color: C.card } });
    s.addText("✓", { x: lx + 0.2, y, w: 0.5, h: 0.58, valign: "middle", fontFace: FONT, fontSize: 18, bold: true, color: C.teal, margin: 0 });
    s.addText(t, { x: lx + 0.7, y, w: lw - 0.9, h: 0.58, valign: "middle", fontFace: FONT, fontSize: 14.5, bold: true, color: C.ink, margin: 0 });
  });

  closingBanner(s, "−8.6kgの中には、たくさんの「できた」がある。");
  folio(s, P());
}

// ---------------------------------------------------------------
// 14. S05｜CASE2 パソコンを開くだけ
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "bl", C.indigo);
  eyebrow(s, "大きな結果は、小さな「できた」の集合体　｜　CASE 2 / 3");
  title(s, "CASE2｜「パソコンを開くだけ」から課題達成", { fontSize: 23 });
  bodyText(s, "最初の一歩は、「パソコンを開くだけ」", MX, 1.78, CW, 0.4, { fontSize: 14, color: C.inkSoft, bold: true });

  const flow = ["開けた", "少し\n取り組めた", "また\n取り組めた", "1週間後までに\n自分で決めた\n課題をクリア"];
  const fw = (CW - 0.9) / 4;
  flow.forEach((t, i) => {
    const x = MX + i * (fw + 0.3);
    const y = 2.35;
    const isLast = i === flow.length - 1;
    card(s, x, y, fw, 1.5, { fill: { color: isLast ? C.indigoBg : C.card } });
    s.addText(t, {
      x: x + 0.1, y, w: fw - 0.2, h: 1.5, align: "center", valign: "middle",
      fontFace: FONT, fontSize: isLast ? 13 : 15, bold: true, color: isLast ? C.indigo : C.ink, margin: 0, lineSpacing: 17,
    });
    if (!isLast) {
      s.addText("→", { x: x + fw, y, w: 0.3, h: 1.5, align: "center", valign: "middle", fontFace: FONT, fontSize: 16, color: C.inkFaint, margin: 0 });
    }
  });

  s.addText([
    { text: "さらに、", options: { fontSize: 14, color: C.inkSoft, breakLine: true } },
    { text: "正しく立つ → 身体が安定 → 落ち着いて考えやすい", options: { fontSize: 16, bold: true, color: C.indigo, breakLine: true } },
    { text: "という感覚を体験。", options: { fontSize: 14, color: C.ink } },
  ], { x: MX, y: 4.35, w: CW, h: 1.1, margin: 0, paraSpaceAfter: 4 });

  closingBanner(s, "身体・思考・習慣を一緒に整えることで、「動けない」が「できた」に変わった。");
  folio(s, P());
}

// ---------------------------------------------------------------
// 15. S05｜CASE3 日本記録
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "tr", C.gold);
  eyebrow(s, "大きな結果は、小さな「できた」の集合体　｜　CASE 3 / 3");
  title(s, "CASE3｜2.5kgの積み重ねが、日本記録へ", { fontSize: 24 });

  card(s, MX, 2.0, 3.3, 1.5, { fill: { color: C.goldBg }, line: { type: "none" } });
  s.addText([
    { text: "前回より、", options: { fontSize: 13, color: C.ink, breakLine: true } },
    { text: "＋2.5kg", options: { fontSize: 30, bold: true, color: C.gold } },
  ], { x: MX, y: 2.0, w: 3.3, h: 1.5, align: "center", valign: "middle", margin: 0 });

  bodyText(s, "今日も練習できた。フォームが少し良くなった。失敗の原因が分かった。\nその積み重ねを11年以上。", MX + 3.6, 2.15, CW - 3.6, 1.2, { fontSize: 15, lineSpacing: 24, valign: "middle" });

  const stats = [
    ["全日本57kg級・52kg級", "日本一"],
    ["スクワット142.5kg", "日本記録"],
  ];
  const sw = (CW - 0.3) / 2;
  stats.forEach((st, i) => {
    const x = MX + i * (sw + 0.3);
    const y = 3.75;
    card(s, x, y, sw, 1.15, { fill: { color: C.card } });
    s.addText(st[0], { x: x + 0.3, y, w: sw - 0.6, h: 1.15, valign: "middle", fontFace: FONT, fontSize: 15, color: C.inkSoft, margin: 0 });
    s.addText(st[1], { x: x + 0.3, y, w: sw - 0.6, h: 1.15, align: "right", valign: "middle", fontFace: FONT, fontSize: 20, bold: true, color: C.gold, margin: 0 });
  });

  closingBanner(s, "大きな結果は、数えきれない「できた」の集合体。");
  folio(s, P());
}

// ---------------------------------------------------------------
// 16. S05B｜今度は、あなたの番です
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "tr", C.coral);
  eyebrow(s, "WORK");
  title(s, "今度は、あなたの番です", { fontSize: 28 });
  bodyText(s, "5分で最初の「できたのかけら」をつくろう", MX, 1.72, CW, 0.4, { fontSize: 15, bold: true, color: C.coralDark });

  const qs = [
    "今、叶えたいことは？",
    "それを叶えたら、何が嬉しい？",
    "すでにできていることは？",
    "今日なら何ができる？",
  ];
  const qw = (CW - 0.3) / 2;
  const qh = 0.95;
  qs.forEach((q, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = MX + col * (qw + 0.3);
    const y = 2.35 + row * (qh + 0.22);
    card(s, x, y, qw, qh);
    pill(s, `${i + 1}`, x + 0.25, y + (qh - 0.44) / 2, 0.44, 0.44, C.white, C.coral, 15);
    s.addText(q, { x: x + 0.9, y, w: qw - 1.5, h: qh, valign: "middle", fontFace: FONT, fontSize: 15.5, bold: true, color: C.ink, margin: 0 });
    s.addShape("line", { x: x + qw - 2.2, y: y + qh - 0.28, w: 1.8, h: 0, line: { color: C.line, width: 1.5 } });
  });

  card(s, MX, 4.65, CW, 1.85, { fill: { color: C.coralBg }, line: { type: "none" } });
  s.addText([
    { text: "ポイントは、", options: { fontSize: 13, color: C.ink, breakLine: true } },
    { text: "❌「頑張ればできる」ではなく、「これならできる」まで小さくすること。", options: { fontSize: 15, bold: true, color: C.coralDark, breakLine: true } },
    { text: "そして今できるなら、やってみる。「できた！」", options: { fontSize: 14, color: C.ink, breakLine: true } },
    { text: "それが今日の、最初のできたのかけら。", options: { fontSize: 16, bold: true, color: C.coralDark } },
  ], { x: MX + 0.4, y: 4.65, w: CW - 0.8, h: 1.85, valign: "middle", margin: 0, paraSpaceAfter: 6 });

  folio(s, P());
}

// ---------------------------------------------------------------
// 17. S06A｜「できた」を生む3つの土台
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "tr", C.indigo);
  title(s, "「できた」を生む3つの土台");
  bodyText(s, "身体 × 思考 × 習慣", MX, 1.75, CW, 0.4, { fontSize: 15, bold: true, color: C.inkSoft });

  const found = [
    ["①", "身体", "「今の私」に戻る", "正しく立つ。呼吸する。\nゆるめる。\n自分の状態を感じる。", C.coral, C.coralBg],
    ["②", "思考", "「進む方向」を決める", "何を叶えたい？\nなぜ叶えたい？\n今の私がやることは？", C.gold, C.goldBg],
    ["③", "習慣", "「また戻れる」をつくる", "いつやる？どれくらいやる？\nできない日は？\n止まったらどう戻る？", C.teal, C.tealBg],
  ];
  const cw = (CW - 0.5) / 3;
  found.forEach((f, i) => {
    const x = MX + i * (cw + 0.25);
    const y = 2.35;
    card(s, x, y, cw, 3.0, { fill: { color: f[5] }, line: { type: "none" } });
    pill(s, f[0], x + 0.28, y + 0.28, 0.48, 0.48, "FFFFFF", f[4], 16);
    s.addText(f[1], { x: x + 0.28, y: y + 0.9, w: cw - 0.56, h: 0.5, fontFace: FONT, fontSize: 22, bold: true, color: f[4], margin: 0 });
    s.addText(f[2], { x: x + 0.28, y: y + 1.4, w: cw - 0.56, h: 0.5, fontFace: FONT, fontSize: 14, bold: true, color: C.ink, margin: 0 });
    s.addText(f[3], { x: x + 0.28, y: y + 1.95, w: cw - 0.56, h: 0.95, fontFace: FONT, fontSize: 12.5, color: C.ink, margin: 0, lineSpacing: 18 });
  });

  closingBanner(s, "そして真ん中に、「できたのかけら」。頑張って動くのではなく、動ける私を整える。");
  folio(s, P());
}

// ---------------------------------------------------------------
// 18. S06B｜90日後、手に入るのは「戻れる私」（リフレーム＋統計）
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "bl", C.coral);
  title(s, "90日後、手に入るのは「戻れる私」", { fontSize: 27 });
  bodyText(s, "90日間で目指すのは、", MX, 1.85, CW, 0.4, { fontSize: 15, color: C.inkSoft });

  card(s, MX, 2.35, 5.7, 1.1, { fill: { color: C.muteBg }, line: { type: "none" } });
  s.addText([
    { text: "❌ ", options: { fontSize: 16, color: C.inkFaint, breakLine: false } },
    { text: "一度も止まらない私", options: { fontSize: 16, bold: true, color: C.inkSoft, breakLine: true } },
    { text: "ではありません。", options: { fontSize: 13, color: C.inkFaint } },
  ], { x: MX + 0.35, y: 2.35, w: 5.0, h: 1.1, valign: "middle", margin: 0 });

  s.addText("→", { x: MX + 5.75, y: 2.35, w: 0.5, h: 1.1, align: "center", valign: "middle", fontFace: FONT, fontSize: 22, color: C.coral, margin: 0 });

  card(s, MX + 6.3, 2.35, CW - 6.3, 1.1, { fill: { color: C.coralBg }, line: { type: "none" } });
  s.addText([
    { text: "✅ 止まっても、", options: { fontSize: 16, bold: true, color: C.coralDark, breakLine: true } },
    { text: "自分で戻れる私。", options: { fontSize: 16, bold: true, color: C.coralDark } },
  ], { x: MX + 6.65, y: 2.35, w: CW - 7.0, h: 1.1, valign: "middle", margin: 0 });

  card(s, MX, 3.85, CW, 2.0, { fill: { color: C.card } });
  s.addText([
    { text: "もし1日1つ、「できた」を見つけたら、", options: { fontSize: 15, color: C.ink, breakLine: true } },
    { text: "1日 × 90日 ＝ 90個", options: { fontSize: 40, bold: true, color: C.coral, breakLine: true } },
    { text: "の「できた」の証拠。", options: { fontSize: 15, color: C.ink } },
  ], { x: MX, y: 3.85, w: CW, h: 2.0, align: "center", valign: "middle", margin: 0, paraSpaceAfter: 4 });

  folio(s, P());
}

// ---------------------------------------------------------------
// 19. S06B｜BEFORE / AFTER
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "tr", C.coral);
  title(s, "「戻れる私」になると、口ぐせが変わる", { fontSize: 25 });

  const beforeItems = ["「またできなかった」", "「私は続かない」", "「何からやればいい？」", "「また失敗したらどうしよう」"];
  const afterItems = ["「今日はここまでできた」", "「今回は何があった？」", "「じゃあ次はどうする？」", "「今の私はこれをやろう」", "「大丈夫。私はまた戻れる。」"];

  const colW = (CW - 0.6) / 2;
  card(s, MX, 1.95, colW, 4.6, { fill: { color: C.muteBg }, line: { type: "none" } });
  pill(s, "BEFORE", MX + 0.35, 2.25, 1.6, 0.44, "FFFFFF", C.inkFaint, 13);
  beforeItems.forEach((t, i) => {
    s.addText(t, {
      x: MX + 0.35, y: 2.95 + i * 0.72, w: colW - 0.7, h: 0.6, valign: "middle",
      fontFace: FONT, fontSize: 15, color: C.inkSoft, margin: 0,
    });
  });

  const rx = MX + colW + 0.6;
  card(s, rx, 1.95, colW, 4.6, { fill: { color: C.coralBg }, line: { type: "none" } });
  pill(s, "AFTER", rx + 0.35, 2.25, 1.5, 0.44, "FFFFFF", C.coral, 13);
  afterItems.forEach((t, i) => {
    s.addText(t, {
      x: rx + 0.35, y: 2.95 + i * 0.58, w: colW - 0.7, h: 0.5, valign: "middle",
      fontFace: FONT, fontSize: 15, bold: true, color: C.coralDark, margin: 0,
    });
  });

  s.addText("→", { x: MX + colW, y: 3.7, w: 0.6, h: 0.8, align: "center", valign: "middle", fontFace: FONT, fontSize: 24, color: C.coral, margin: 0 });
  folio(s, P());
}

// month-pillar template used by slides 20-22
function monthSlide(num, tagColor, tagBg, kanji, subtitle, bodyContent, closing) {
  const s = newSlide(pres);
  shardCorner(s, num % 2 === 0 ? "bl" : "tr", tagColor);
  eyebrow(s, "「戻れる私」をつくる3つの柱");
  pill(s, `${num}ヶ月目`, MX, 1.3, 1.55, 0.5, "FFFFFF", tagColor, 15);
  s.addText(kanji, {
    x: MX + 1.75, y: 1.22, w: 5, h: 0.65,
    fontFace: FONT, fontSize: 30, bold: true, color: C.ink, margin: 0,
  });
  bodyText(s, subtitle, MX, 2.0, CW, 0.4, { fontSize: 16, bold: true, color: tagColor });
  bodyContent(s, tagColor, tagBg);
  closingBanner(s, closing);
  folio(s, P());
}

// ---------------------------------------------------------------
// 20. 1ヶ月目｜身体
// ---------------------------------------------------------------
monthSlide(1, C.coral, C.coralBg, "身体", "動ける土台をつくる", (s, tagColor, tagBg) => {
  const items = ["正しく立つ", "呼吸", "ゆるめる", "スクワット", "自分の状態を感じる"];
  const cw = (CW - 0.8) / 5;
  items.forEach((t, i) => {
    const x = MX + i * (cw + 0.2);
    const y = 2.65;
    card(s, x, y, cw, 2.2, { fill: { color: tagBg }, line: { type: "none" } });
    s.addText(t, { x: x + 0.12, y, w: cw - 0.24, h: 2.2, align: "center", valign: "middle", fontFace: FONT, fontSize: 14.5, bold: true, color: C.coralDark, margin: 0, lineSpacing: 18 });
  });
}, "身体から最初の「できた」をつくる");

// ---------------------------------------------------------------
// 21. 2ヶ月目｜思考
// ---------------------------------------------------------------
monthSlide(2, C.gold, C.goldBg, "思考", "「今日はこれ」を決める", (s, tagColor, tagBg) => {
  const flow = ["目的", "3ヶ月", "1ヶ月", "1週間", "今日"];
  const cw = (CW - 0.8) / 5;
  flow.forEach((t, i) => {
    const x = MX + i * (cw + 0.2);
    const y = 2.65;
    const isLast = i === flow.length - 1;
    card(s, x, y, cw, 1.1, { fill: { color: isLast ? tagBg : C.card } });
    s.addText(t, { x, y, w: cw, h: 1.1, align: "center", valign: "middle", fontFace: FONT, fontSize: 15, bold: true, color: isLast ? C.gold : C.ink, margin: 0 });
  });
  s.addText([
    { text: "自分で決める。やってみる。できた。", options: { fontSize: 16, bold: true, color: C.ink, breakLine: true } },
    { text: "「私は自分の判断を信用していい」へ。", options: { fontSize: 15, color: C.inkSoft } },
  ], { x: MX, y: 4.15, w: CW, h: 1.0, margin: 0, paraSpaceAfter: 8 });
}, "自分で決めて、やってみる経験を積み重ねる");

// ---------------------------------------------------------------
// 22. 3ヶ月目｜習慣
// ---------------------------------------------------------------
monthSlide(3, C.teal, C.tealBg, "習慣", "戻る仕組みをつくる", (s, tagColor, tagBg) => {
  const flow = ["なぜ止まった？", "何を変える？", "次はどうする？", "また戻る。"];
  const cw = (CW - 0.6) / 4;
  flow.forEach((t, i) => {
    const x = MX + i * (cw + 0.2);
    const y = 2.65;
    const isLast = i === flow.length - 1;
    card(s, x, y, cw, 1.1, { fill: { color: isLast ? tagBg : C.card } });
    s.addText(t, { x: x + 0.1, y, w: cw - 0.2, h: 1.1, align: "center", valign: "middle", fontFace: FONT, fontSize: 14, bold: true, color: isLast ? C.teal : C.ink, margin: 0 });
  });
  s.addText("身体で戻る。思考で決める。習慣でまた戻る。", {
    x: MX, y: 4.15, w: CW, h: 0.5, fontFace: FONT, fontSize: 16, bold: true, color: C.ink, margin: 0,
  });
}, "そのすべてをつなぐのが、「できたのかけら」");

// ---------------------------------------------------------------
// 23. S07｜できたのかけらメソッド全体像
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "tr", C.coral);
  title(s, "90日間、一人で頑張らなくていい", { fontSize: 26 });
  bodyText(s, "できたのかけらメソッド ― 小さな「できた」から、自分を信じて目標を叶える90日間 ―", MX, 1.62, CW, 0.4, { fontSize: 13.5, bold: true, color: C.coralDark });

  const months = [
    ["1ヶ月目｜身体", ["身体の現在地を知る", "身体から「できた」をつくる"], C.coral, C.coralBg],
    ["2ヶ月目｜思考", ["本当に叶えたい目標を決める", "目標を「今日の一歩」まで小さくする"], C.gold, C.goldBg],
    ["3ヶ月目｜習慣", ["私の「止まるパターン」を知る", "「またできる」私の取扱説明書をつくる"], C.teal, C.tealBg],
  ];
  const cw = (CW - 0.5) / 3;
  months.forEach((m, i) => {
    const x = MX + i * (cw + 0.25);
    const y = 2.2;
    card(s, x, y, cw, 2.15, { fill: { color: m[3] }, line: { type: "none" } });
    s.addText(m[0], { x: x + 0.25, y: y + 0.2, w: cw - 0.5, h: 0.4, fontFace: FONT, fontSize: 15, bold: true, color: m[2], margin: 0 });
    m[1].forEach((lesson, li) => {
      const ly = y + 0.75 + li * 0.7;
      s.addText(`Lesson ${i * 2 + li + 1}`, { x: x + 0.25, y: ly, w: cw - 0.5, h: 0.25, fontFace: FONT, fontSize: 10.5, bold: true, color: m[2], margin: 0 });
      s.addText(lesson, { x: x + 0.25, y: ly + 0.24, w: cw - 0.5, h: 0.45, fontFace: FONT, fontSize: 12.5, bold: true, color: C.ink, margin: 0, lineSpacing: 15 });
    });
  });

  s.addText("グループ講義　月2回 × 3ヶ月 ＝ 全6回", {
    x: MX, y: 4.55, w: CW, h: 0.4, align: "center",
    fontFace: FONT, fontSize: 14, bold: true, color: C.inkSoft, margin: 0,
  });

  card(s, MX, 5.05, CW, 1.35, { fill: { color: C.coralBg }, line: { type: "none" } });
  s.addText([
    { text: "でも主役は、講義と講義の間にある日常。", options: { fontSize: 15, bold: true, color: C.coralDark, breakLine: true } },
    { text: "学ぶ → やる → 止まる → 相談する → 変える → 戻る", options: { fontSize: 14, bold: true, color: C.ink, breakLine: true } },
    { text: "この経験を90日間積み重ねます。", options: { fontSize: 12.5, color: C.inkSoft } },
  ], { x: MX + 0.4, y: 5.05, w: CW - 0.8, h: 1.35, valign: "middle", margin: 0, paraSpaceAfter: 4 });

  folio(s, P());
}

// ---------------------------------------------------------------
// 24. S08｜止まった日にこそ、ひとりにしない
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "bl", C.indigo);
  title(s, "止まった日にこそ、ひとりにしない", { fontSize: 27 });
  bodyText(s, "「戻るため」の5つのサポート", MX, 1.75, CW, 0.4, { fontSize: 15, bold: true, color: C.inkSoft });

  const supports = [
    ["👤", "個別セッション3回", "現在地→軌道修正→90日間の振り返り。「私の場合は？」を一緒に整理"],
    ["🌅", "朝活", "身体を整え、その日最初の「できた」をつくる"],
    ["📝", "できたのかけら報告", "自分では見えない「できた」を、仲間と見つける"],
    ["💬", "オープンチャット", "成功報告だけじゃない。「止まってます」と言っていい場所。"],
    ["🎥", "講義アーカイブ", "参加できなかった日も、あとから戻れる。"],
  ];
  const cw = (CW - 0.8) / 5;
  supports.forEach((sp, i) => {
    const x = MX + i * (cw + 0.2);
    const y = 2.35;
    card(s, x, y, cw, 2.7);
    iconCircle(s, sp[0], x + cw / 2 - 0.33, y + 0.25, 0.66, C.indigoBg, 24);
    s.addText(sp[1], {
      x: x + 0.15, y: y + 1.0, w: cw - 0.3, h: 0.55, align: "center",
      fontFace: FONT, fontSize: 13, bold: true, color: C.ink, margin: 0, lineSpacing: 16,
    });
    s.addText(sp[2], {
      x: x + 0.15, y: y + 1.55, w: cw - 0.3, h: 1.05, align: "center",
      fontFace: FONT, fontSize: 10.5, color: C.inkSoft, margin: 0, lineSpacing: 14,
    });
  });

  closingBanner(s, "すべてのサポートの目的は一つ。「止まっても、また戻れる」ため。");
  folio(s, P());
}

// ---------------------------------------------------------------
// 25. S09a｜プログラム内容
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "tr", C.coral);
  title(s, "90日後の「私なら、またできる」へ", { fontSize: 26 });
  bodyText(s, "できたのかけらメソッド　小さな「できた」から、自分を信じて目標を叶える90日間", MX, 1.7, CW, 0.4, { fontSize: 13, color: C.inkSoft, bold: true });

  const items = [
    "グループ講義　全6回／各90分",
    "個別セッション　3回",
    "朝活",
    "できたのかけら報告コミュニティ",
    "オープンチャット伴走",
    "講義アーカイブ",
  ];
  const cw2 = (CW - 0.3) / 2;
  items.forEach((t, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = MX + col * (cw2 + 0.3);
    const y = 2.35 + row * 0.72;
    card(s, x, y, cw2, 0.58, { fill: { color: C.card } });
    s.addText("＋", { x: x + 0.2, y, w: 0.4, h: 0.58, valign: "middle", fontFace: FONT, fontSize: 15, bold: true, color: C.coral, margin: 0 });
    s.addText(t, { x: x + 0.65, y, w: cw2 - 0.85, h: 0.58, valign: "middle", fontFace: FONT, fontSize: 14.5, bold: true, color: C.ink, margin: 0 });
  });

  card(s, MX, 4.75, CW, 1.65, { fill: { color: C.coralBg }, line: { type: "none" } });
  pill(s, "第1期募集　限定5名", MX + 0.4, 5.0, 3.0, 0.5, "FFFFFF", C.coralDark, 15);
  s.addText("一人ひとりの目標と、「できたのかけら」をちゃんと見つけながら伴走したいから。", {
    x: MX + 0.4, y: 5.65, w: CW - 0.8, h: 0.6, valign: "top",
    fontFace: FONT, fontSize: 13.5, color: C.ink, margin: 0, lineSpacing: 19,
  });

  folio(s, P());
}

// ---------------------------------------------------------------
// 26. S09b｜参加費
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "bl", C.gold);
  title(s, "参加費", { fontSize: 26 });

  card(s, MX, 1.9, CW, 1.7, { fill: { color: C.goldBg }, line: { type: "none" } });
  s.addText([
    { text: "298,000", options: { fontSize: 54, bold: true, color: C.gold, breakLine: false } },
    { text: "円", options: { fontSize: 26, bold: true, color: C.gold } },
  ], { x: MX, y: 1.9, w: CW, h: 1.7, align: "center", valign: "middle", margin: 0 });

  s.addText([
    { text: "これは、6回の講義を聞くためだけの90日間ではありません。", options: { fontSize: 16, bold: true, color: C.ink, breakLine: true } },
    { text: "できる。止まる。相談する。変える。戻る。", options: { fontSize: 20, bold: true, color: C.coralDark, breakLine: true } },
    { text: "その経験を積み重ね、「私は私を信用していい」を育てる90日間です。", options: { fontSize: 16, bold: true, color: C.ink } },
  ], { x: 1.3, y: 4.05, w: W - 2.6, h: 2.3, align: "center", margin: 0, paraSpaceAfter: 14, lineSpacing: 24 });

  folio(s, P());
}

// ---------------------------------------------------------------
// 27. S10a｜Q&A
// ---------------------------------------------------------------
{
  const s = newSlide(pres);
  shardCorner(s, "tr", C.teal);
  title(s, "「私にもできる？」と思っているあなたへ", { fontSize: 24 });

  const qas = [
    ["私にもできますか？", "「これならできる」から始めます。"],
    ["忙しくても大丈夫？", "今の生活に入る大きさを一緒に探します。"],
    ["また止まったら？", "止まってください。そこから戻る練習をします。"],
    ["結果が出なかったら？", "結果だけではなく、この先も使える「自分の進み方」をつくります。"],
    ["一人でやってみてもいい？", "もちろんです。でも、一人で何度も止まってきたなら、「環境を変える」ことも一つの選択です。"],
  ];
  const rowH = 0.83;
  qas.forEach((qa, i) => {
    const y = 1.95 + i * (rowH + 0.1);
    card(s, MX, y, CW, rowH);
    s.addText("Q", { x: MX + 0.25, y, w: 0.5, h: rowH, valign: "middle", fontFace: FONT, fontSize: 18, bold: true, color: C.teal, margin: 0 });
    s.addText(qa[0], { x: MX + 0.75, y, w: 3.4, h: rowH, valign: "middle", fontFace: FONT, fontSize: 14, bold: true, color: C.ink, margin: 0, lineSpacing: 16 });
    s.addText("→", { x: MX + 4.2, y, w: 0.45, h: rowH, valign: "middle", align: "center", fontFace: FONT, fontSize: 15, color: C.inkFaint, margin: 0 });
    s.addText(qa[1], { x: MX + 4.75, y, w: CW - 5.05, h: rowH, valign: "middle", fontFace: FONT, fontSize: 13.5, color: C.inkSoft, margin: 0, lineSpacing: 16 });
  });

  folio(s, P());
}

// ---------------------------------------------------------------
// 28. S10b｜最後に
// ---------------------------------------------------------------
{
  const s = newSlide(pres, C.bgDark);
  shardCorner(s, "tr", C.coral);
  shardCorner(s, "bl", C.teal);
  eyebrow(s, "最後に", { color: C.gold });
  s.addText([
    { text: "この90日間、一度も止まらなくていい。完璧じゃなくていい。\nできない日があっていい。迷っていい。誰かに頼っていい。", options: { fontSize: 18, color: C.creamText, breakLine: true } },
    { text: "ただ、一つだけ。", options: { fontSize: 16, color: C.creamSoft, breakLine: true } },
    { text: "止まっても、また戻ってくる。", options: { fontSize: 22, bold: true, color: C.coral, breakLine: true } },
    { text: "そして戻ってきたとき、一緒に言いましょう。", options: { fontSize: 16, color: C.creamSoft, breakLine: true } },
    { text: "「戻ってこられた。今日も一つ、できた。」", options: { fontSize: 26, bold: true, color: C.white } },
  ], { x: 1.2, y: 1.6, w: W - 2.4, h: 5.2, align: "center", valign: "middle", margin: 0, lineSpacing: 30, paraSpaceAfter: 16 });

  folio(s, P(), true);
}

// ---------------------------------------------------------------
// 29. LAST｜あなたの未来へ
// ---------------------------------------------------------------
{
  const s = newSlide(pres, C.bgDark);
  shardsLast(s);
  eyebrow(s, "できたのかけらワークショップ", { x: MX, y: 0.7, color: C.gold, align: "center", w: W - 1.4 });
  s.addText("できたのかけらメソッド", {
    x: 0.8, y: 1.35, w: 11.73, h: 0.75, align: "center",
    fontFace: FONT_TITLE, fontSize: 32, bold: true, color: C.white, margin: 0,
  });
  s.addText("― 小さな「できた」から、自分を信じて目標を叶える90日間 ―", {
    x: 0.8, y: 2.1, w: 11.73, h: 0.5, align: "center",
    fontFace: FONT, fontSize: 14, color: C.creamSoft, margin: 0,
  });

  s.addText("第1期｜限定5名　90日間　298,000円", {
    x: 0.8, y: 2.85, w: 11.73, h: 0.5, align: "center",
    fontFace: FONT, fontSize: 16, bold: true, color: C.gold, margin: 0,
  });

  s.addText("そして最後は、この言葉で。", {
    x: 0.8, y: 3.6, w: 11.73, h: 0.4, align: "center",
    fontFace: FONT, fontSize: 14, color: C.creamSoft, margin: 0,
  });
  s.addText("「私なら、またできる。」", {
    x: 0.8, y: 4.05, w: 11.73, h: 1.0, align: "center",
    fontFace: FONT_TITLE, fontSize: 40, bold: true, color: C.coral, margin: 0,
  });
  s.addText("小さな「できた」が、あなたの未来への道しるべになる。", {
    x: 0.8, y: 5.15, w: 11.73, h: 0.5, align: "center",
    fontFace: FONT, fontSize: 15, color: C.creamText, margin: 0,
  });

  folio(s, P(), true);
}

console.log(`build complete: ${PAGE - 1} content slides + cover = ${PAGE} total`);
pres.writeFile({ fileName: path.join(__dirname, "..", "dist", "dekitanokakera-workshop.pptx") }).then(() => {
  console.log("wrote final pptx");
});
