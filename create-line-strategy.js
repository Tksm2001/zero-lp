const pptxgen = require("pptxgenjs");
let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "ZERO PROJECT";
pres.title = "LINE流入後フロー戦略・戦術・戦闘";

// ── Color Palette: Cherry Bold inspired ──
const C = {
  bg: "1A1208",       // dark brown
  bgLight: "2C1F0E",  // lighter brown
  gold: "B8860B",      // dark gold accent
  goldLight: "D4A843", // light gold
  text: "F5EDD8",      // cream text
  textMuted: "A89B7A", // muted text
  red: "A83232",       // cherry red
  blue: "2E6B8A",      // trust blue
  orange: "C47D1A",    // warm orange
  green: "4A7A4A",     // success green
  white: "FFFFFF",
  black: "0D0904",
};

function addBg(slide) {
  slide.background = { color: C.bg };
}
function addFooter(slide, num, total) {
  slide.addText(`${num} / ${total}`, { x: 8.5, y: 5.2, w: 1.2, h: 0.3, fontSize: 9, color: C.textMuted, align: "right" });
  slide.addText("ZERO PROJECT — CONFIDENTIAL", { x: 0.5, y: 5.2, w: 3, h: 0.3, fontSize: 8, color: C.textMuted });
}

const TOTAL = 20;
let slideNum = 0;

// ═══════════════════════════════════════
// SLIDE 1: Title
// ═══════════════════════════════════════
slideNum++;
let s1 = pres.addSlide();
addBg(s1);
s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.black, transparency: 30 } });
s1.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.5, w: 0.06, h: 2.5, fill: { color: C.gold } });
s1.addText("LINE流入後\nフロー戦略・戦術・戦闘", { x: 0.9, y: 1.5, w: 8, h: 1.8, fontSize: 38, fontFace: "Arial Black", color: C.text, bold: true, lineSpacingMultiple: 1.2, margin: 0 });
s1.addText("— 高額商材（50〜200万円）のLINE教育ファネル完全設計書 —", { x: 0.9, y: 3.3, w: 8, h: 0.5, fontSize: 14, color: C.goldLight, margin: 0 });
s1.addText("ZERO PROJECT  |  2026.04", { x: 0.9, y: 4.0, w: 4, h: 0.4, fontSize: 12, color: C.textMuted, margin: 0 });
s1.addText("CONFIDENTIAL", { x: 7, y: 4.8, w: 2.5, h: 0.3, fontSize: 10, color: C.red, align: "right", bold: true });

// ═══════════════════════════════════════
// SLIDE 2: 目次
// ═══════════════════════════════════════
slideNum++;
let s2 = pres.addSlide();
addBg(s2);
addFooter(s2, slideNum, TOTAL);
s2.addText("目次", { x: 0.5, y: 0.3, w: 4, h: 0.7, fontSize: 32, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });
s2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.95, w: 1.5, h: 0.04, fill: { color: C.gold } });

const tocItems = [
  ["01", "戦略概要", "全体ファネルと3フェーズ設計"],
  ["02", "市場リサーチ", "被害者データから学ぶ最前線の心理技術"],
  ["03", "72時間の戦場マップ", "4つの戦闘区間と制圧戦略"],
  ["04", "Day0：制圧戦", "登録直後〜1時間の設計"],
  ["05", "Day0：追撃戦", "1〜6時間のフォロー設計"],
  ["06", "Day1-3：定着戦", "加速/標準2ルート並走"],
  ["07", "Day4-9：価値観書き換え", "痛み→解放→希望の感情設計"],
  ["08", "Day10-13：トスアップ", "公式LINE→個人LINE移行"],
  ["09", "配信頻度と重さ", "マルチフォーマット戦略"],
  ["10", "5つの心理原則", "合法×最大効果の境界線"],
  ["11", "オペレーション設計", "担当者マニュアルとAI化ロードマップ"],
  ["12", "数値予測とKPI", "月100人登録→成約のファネル"],
];
tocItems.forEach((item, i) => {
  const y = 1.3 + i * 0.34;
  s2.addText(item[0], { x: 0.5, y, w: 0.5, h: 0.3, fontSize: 11, fontFace: "Arial Black", color: C.gold, margin: 0 });
  s2.addText(item[1], { x: 1.1, y, w: 2.8, h: 0.3, fontSize: 12, color: C.text, bold: true, margin: 0 });
  s2.addText(item[2], { x: 4.2, y, w: 5.3, h: 0.3, fontSize: 10, color: C.textMuted, margin: 0 });
});

// ═══════════════════════════════════════
// SLIDE 3: 戦略概要 — ファネル全体フロー
// ═══════════════════════════════════════
slideNum++;
let s3 = pres.addSlide();
addBg(s3);
addFooter(s3, slideNum, TOTAL);
s3.addText("01  戦略概要：ファネル全体フロー", { x: 0.5, y: 0.3, w: 8, h: 0.6, fontSize: 24, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });

// 3 phase boxes
const phases = [
  { label: "Phase 1", title: "関係構築", days: "Day0〜3", color: C.blue, desc: "孤独の解消\n最初の理解者になる\n2往復以上のやり取り" },
  { label: "Phase 2", title: "価値観書き換え", days: "Day4〜9", color: C.orange, desc: "痛み→解放→希望\n投資観インストール\n師匠のセリフ引用" },
  { label: "Phase 3", title: "トスアップ", days: "Day10〜13", color: C.red, desc: "個人LINE移行\n師匠の紹介\n逆説的クロージング" },
];
phases.forEach((p, i) => {
  const x = 0.5 + i * 3.1;
  s3.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 2.9, h: 2.5, fill: { color: C.bgLight } });
  s3.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 2.9, h: 0.05, fill: { color: p.color } });
  s3.addText(p.label, { x, y: 1.35, w: 2.9, h: 0.3, fontSize: 10, color: p.color, bold: true, align: "center", margin: 0 });
  s3.addText(p.title, { x, y: 1.65, w: 2.9, h: 0.4, fontSize: 18, fontFace: "Arial Black", color: C.text, bold: true, align: "center", margin: 0 });
  s3.addText(p.days, { x, y: 2.05, w: 2.9, h: 0.3, fontSize: 11, color: C.goldLight, align: "center", margin: 0 });
  s3.addText(p.desc, { x: x + 0.3, y: 2.4, w: 2.3, h: 1.2, fontSize: 10, color: C.textMuted, lineSpacingMultiple: 1.4, margin: 0 });
});

// Flow arrows
s3.addText("→", { x: 3.3, y: 2.0, w: 0.4, h: 0.4, fontSize: 20, color: C.gold, align: "center" });
s3.addText("→", { x: 6.4, y: 2.0, w: 0.4, h: 0.4, fontSize: 20, color: C.gold, align: "center" });

// Bottom key insight
s3.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.0, w: 9, h: 0.9, fill: { color: C.bgLight } });
s3.addText("核心設計：共感 → 痛み → 解放 → 希望 → 投資観 → 信頼移転 → 決断", { x: 0.8, y: 4.05, w: 8.4, h: 0.35, fontSize: 13, color: C.gold, bold: true, margin: 0 });
s3.addText("高額商材は「恐怖×希望の振り子」で感情を動かす。ノウハウより先に考え方を変える。", { x: 0.8, y: 4.4, w: 8.4, h: 0.35, fontSize: 11, color: C.textMuted, margin: 0 });

// ═══════════════════════════════════════
// SLIDE 4: 市場リサーチ — 被害者データからの学び
// ═══════════════════════════════════════
slideNum++;
let s4 = pres.addSlide();
addBg(s4);
addFooter(s4, slideNum, TOTAL);
s4.addText("02  被害者データから学ぶ最前線の心理技術", { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 22, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });

// Left: Source data
s4.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 4.3, h: 3.8, fill: { color: C.bgLight } });
s4.addText("情報源", { x: 0.7, y: 1.2, w: 2, h: 0.35, fontSize: 14, color: C.gold, bold: true, margin: 0 });
const sources = [
  "国民生活センター報告書：年間4,000〜7,000件",
  "被害金額帯：30万〜200万円が主流",
  "50万円契約体験談（ブログ）",
  "500万円注ぎ込んだ体験談",
  "勧誘トップ経験者の内部告発（note）",
  "個別面談クロージング率「ほぼ100%」",
];
sources.forEach((s, i) => {
  s4.addText(s, { x: 0.9, y: 1.65 + i * 0.5, w: 3.7, h: 0.45, fontSize: 10, color: C.text, valign: "top", margin: 0 });
  s4.addShape(pres.shapes.OVAL, { x: 0.7, y: 1.78 + i * 0.5, w: 0.1, h: 0.1, fill: { color: C.gold } });
});

// Right: Key findings
s4.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.1, w: 4.4, h: 3.8, fill: { color: C.bgLight } });
s4.addText("最重要発見", { x: 5.3, y: 1.2, w: 3, h: 0.35, fontSize: 14, color: C.red, bold: true, margin: 0 });
const findings = [
  ["面談では「悩み」ではなく「欲」を刺激", "→ 自由・収入・時間の理想を描かせる"],
  ["断っても1年間LINEで欲を刺激し続ける", "→ 長期育成シナリオの根拠"],
  ["コミュニティの一体感が最強のロックイン", "→ 契約後のグループ参加が必須"],
];
findings.forEach((f, i) => {
  s4.addText(f[0], { x: 5.3, y: 1.7 + i * 1.1, w: 4, h: 0.35, fontSize: 11, color: C.text, bold: true, margin: 0 });
  s4.addText(f[1], { x: 5.3, y: 2.05 + i * 1.1, w: 4, h: 0.3, fontSize: 10, color: C.goldLight, margin: 0 });
});

// ═══════════════════════════════════════
// SLIDE 5: 72時間の戦場マップ
// ═══════════════════════════════════════
slideNum++;
let s5 = pres.addSlide();
addBg(s5);
addFooter(s5, slideNum, TOTAL);
s5.addText("03  72時間の戦場マップ", { x: 0.5, y: 0.3, w: 8, h: 0.6, fontSize: 24, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });

// Timeline bar
s5.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.2, w: 9, h: 0.5, fill: { color: C.bgLight } });
const timePoints = [
  { label: "0h", x: 0.5, color: C.red },
  { label: "1h", x: 2.3, color: C.red },
  { label: "6h", x: 4.1, color: C.orange },
  { label: "24h", x: 5.9, color: C.blue },
  { label: "48h", x: 7.7, color: C.green },
  { label: "72h", x: 9.0, color: C.green },
];
timePoints.forEach(t => {
  s5.addShape(pres.shapes.OVAL, { x: t.x, y: 1.3, w: 0.3, h: 0.3, fill: { color: t.color } });
  s5.addText(t.label, { x: t.x - 0.1, y: 1.65, w: 0.5, h: 0.25, fontSize: 9, color: C.textMuted, align: "center", margin: 0 });
});

// 4 battle zones
const battles = [
  { title: "制圧戦", time: "0〜1h", desc: "あいさつ+PDF\n欲求4択\n共感+思い出す問い", color: C.red, x: 0.5 },
  { title: "追撃戦", time: "1〜6h", desc: "2往復目の回収\nリマインド配信\n加速ルート判定", color: C.orange, x: 2.9 },
  { title: "定着戦", time: "6〜24h", desc: "音声メッセージ\nあいりの体験談\n「明日続き話すね」", color: C.blue, x: 5.3 },
  { title: "確定戦", time: "24〜72h", desc: "動画リッチメッセージ\n恐怖の2択\nAランク確定", color: C.green, x: 7.7 },
];
battles.forEach(b => {
  s5.addShape(pres.shapes.RECTANGLE, { x: b.x, y: 2.1, w: 2.2, h: 2.8, fill: { color: C.bgLight } });
  s5.addShape(pres.shapes.RECTANGLE, { x: b.x, y: 2.1, w: 2.2, h: 0.05, fill: { color: b.color } });
  s5.addText(b.title, { x: b.x, y: 2.2, w: 2.2, h: 0.4, fontSize: 16, fontFace: "Arial Black", color: b.color, align: "center", bold: true, margin: 0 });
  s5.addText(b.time, { x: b.x, y: 2.6, w: 2.2, h: 0.25, fontSize: 10, color: C.goldLight, align: "center", margin: 0 });
  s5.addText(b.desc, { x: b.x + 0.2, y: 3.0, w: 1.8, h: 1.6, fontSize: 10, color: C.text, lineSpacingMultiple: 1.5, margin: 0 });
});

// ═══════════════════════════════════════
// SLIDE 6: Day0 制圧戦 — 登録直後
// ═══════════════════════════════════════
slideNum++;
let s6 = pres.addSlide();
addBg(s6);
addFooter(s6, slideNum, TOTAL);
s6.addText("04  Day0：制圧戦（0〜1時間）", { x: 0.5, y: 0.3, w: 8, h: 0.6, fontSize: 22, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });
s6.addText("目的：「この人は私のことをわかってる」を刻む", { x: 0.5, y: 0.85, w: 8, h: 0.3, fontSize: 12, color: C.goldLight, italic: true, margin: 0 });

// 4 steps
const day0Steps = [
  { time: "0分", action: "あいさつ+PDF配布", detail: "約束の即時履行\n冒頭1行がプッシュ通知に\n「〇ページだけ見て」で\nハードル極限まで下げる", type: "自動" },
  { time: "30秒後", action: "欲求4択リプライ", detail: "時間/収入/脱出/家族\n4択クイックリプライ\nセグメント取得が目的", type: "自動" },
  { time: "返信直後", action: "共感+思い出す問い", detail: "相手の選択に即応答\n「いつまで続くんだろう\nって思った瞬間ある？」\n3つの具体例を列挙", type: "自動" },
  { time: "3時間後", action: "PDFリマインド", detail: "返信なし組のみ\n「保存できた？」\n具体的ページ番号を指定\n行動のハードルを下げる", type: "自動" },
];
day0Steps.forEach((step, i) => {
  const x = 0.5 + i * 2.35;
  s6.addShape(pres.shapes.RECTANGLE, { x, y: 1.3, w: 2.15, h: 3.5, fill: { color: C.bgLight } });
  s6.addText(step.time, { x, y: 1.4, w: 2.15, h: 0.3, fontSize: 10, color: C.gold, bold: true, align: "center", margin: 0 });
  s6.addText(step.action, { x, y: 1.7, w: 2.15, h: 0.4, fontSize: 13, color: C.text, bold: true, align: "center", margin: 0 });
  s6.addShape(pres.shapes.RECTANGLE, { x: x + 0.3, y: 2.15, w: 1.55, h: 0.03, fill: { color: C.gold, transparency: 50 } });
  s6.addText(step.detail, { x: x + 0.15, y: 2.3, w: 1.85, h: 1.8, fontSize: 9.5, color: C.textMuted, lineSpacingMultiple: 1.4, margin: 0 });
  // Badge
  s6.addShape(pres.shapes.RECTANGLE, { x: x + 0.6, y: 4.4, w: 0.95, h: 0.25, fill: { color: C.blue, transparency: 30 } });
  s6.addText(step.type, { x: x + 0.6, y: 4.4, w: 0.95, h: 0.25, fontSize: 8, color: C.blue, align: "center", margin: 0 });
});

// ═══════════════════════════════════════
// SLIDE 7: 分岐ロジック — 加速/標準ルート
// ═══════════════════════════════════════
slideNum++;
let s7 = pres.addSlide();
addBg(s7);
addFooter(s7, slideNum, TOTAL);
s7.addText("05  Day0：加速/標準ルート分岐", { x: 0.5, y: 0.3, w: 8, h: 0.6, fontSize: 22, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });

// Decision box
s7.addShape(pres.shapes.RECTANGLE, { x: 2.5, y: 1.1, w: 5, h: 0.8, fill: { color: C.bgLight } });
s7.addText("加速ルート判定条件（3つ中2つ該当）", { x: 2.7, y: 1.15, w: 4.6, h: 0.3, fontSize: 12, color: C.gold, bold: true, margin: 0 });
s7.addText("10分以内に返信  /  3行以上の自由返信  /  感情を語っている", { x: 2.7, y: 1.5, w: 4.6, h: 0.3, fontSize: 10, color: C.text, margin: 0 });

// Left: 加速ルート
s7.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.2, w: 4.3, h: 3.0, fill: { color: C.bgLight } });
s7.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.2, w: 4.3, h: 0.05, fill: { color: C.red } });
s7.addText("加速ルート（10〜15人/100人）", { x: 0.7, y: 2.35, w: 3.9, h: 0.35, fontSize: 14, color: C.red, bold: true, margin: 0 });
s7.addText([
  { text: "自動配信を一時停止", options: { bullet: true, breakLine: true, fontSize: 10, color: C.text } },
  { text: "手動チャットに完全移行", options: { bullet: true, breakLine: true, fontSize: 10, color: C.text } },
  { text: "チェックポイント①〜⑤を会話内で通過", options: { bullet: true, breakLine: true, fontSize: 10, color: C.text } },
  { text: "最短3日でアポ到達", options: { bullet: true, breakLine: true, fontSize: 10, color: C.gold, bold: true } },
  { text: "Day0で音声メッセージ投入", options: { bullet: true, breakLine: true, fontSize: 10, color: C.text } },
  { text: "アポ到達率：50〜60%", options: { bullet: true, fontSize: 10, color: C.goldLight, bold: true } },
], { x: 0.7, y: 2.8, w: 3.9, h: 2.2, lineSpacingMultiple: 1.5, margin: 0 });

// Right: 標準ルート
s7.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.2, w: 4.3, h: 3.0, fill: { color: C.bgLight } });
s7.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 2.2, w: 4.3, h: 0.05, fill: { color: C.blue } });
s7.addText("標準ルート（85〜90人/100人）", { x: 5.4, y: 2.35, w: 3.9, h: 0.35, fontSize: 14, color: C.blue, bold: true, margin: 0 });
s7.addText([
  { text: "Day0〜13自動配信を継続", options: { bullet: true, breakLine: true, fontSize: 10, color: C.text } },
  { text: "テンプレ配信+軽い手動フォロー", options: { bullet: true, breakLine: true, fontSize: 10, color: C.text } },
  { text: "スケジュール通り進行", options: { bullet: true, breakLine: true, fontSize: 10, color: C.text } },
  { text: "10〜13日でアポ打診", options: { bullet: true, breakLine: true, fontSize: 10, color: C.text } },
  { text: "Day2に動画リッチメッセージ投入", options: { bullet: true, breakLine: true, fontSize: 10, color: C.text } },
  { text: "アポ到達率：5〜10%", options: { bullet: true, fontSize: 10, color: C.goldLight, bold: true } },
], { x: 5.4, y: 2.8, w: 3.9, h: 2.2, lineSpacingMultiple: 1.5, margin: 0 });

// ═══════════════════════════════════════
// SLIDE 8: セグメント別メッセージ
// ═══════════════════════════════════════
slideNum++;
let s8 = pres.addSlide();
addBg(s8);
addFooter(s8, slideNum, TOTAL);
s8.addText("04b  セグメント別「思い出させる問い」", { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 22, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });
s8.addText("原則：考えさせるな、思い出させろ。人は「考える」より「思い出す」方が10倍簡単。", { x: 0.5, y: 0.85, w: 9, h: 0.3, fontSize: 11, color: C.goldLight, italic: true, margin: 0 });

const segments = [
  { emoji: "時間", q: "「いつまで続くんだろう」\nって思った瞬間は？", examples: "朝の満員電車？\n寝かしつけの後？\n日曜の夜？" },
  { emoji: "収入", q: "お金のことで\n眠れなかった夜ある？", examples: "給料日の通帳？\nクレカの引落日？\n習い事の月謝？" },
  { emoji: "脱出", q: "「もう辞めたい」って\n一番強く思った瞬間は？", examples: "上司に言われた時？\n日曜の夜？\n朝のアラーム？" },
  { emoji: "家族", q: "子どもに「ごめんね」\nって思った瞬間は？", examples: "お迎えに間に合わない？\n「遊ぼう」を断った？\n寝顔しか見れない？" },
];
segments.forEach((seg, i) => {
  const x = 0.5 + i * 2.35;
  s8.addShape(pres.shapes.RECTANGLE, { x, y: 1.3, w: 2.15, h: 3.8, fill: { color: C.bgLight } });
  s8.addText(seg.emoji, { x, y: 1.4, w: 2.15, h: 0.35, fontSize: 14, color: C.gold, bold: true, align: "center", margin: 0 });
  s8.addShape(pres.shapes.RECTANGLE, { x: x + 0.3, y: 1.8, w: 1.55, h: 0.03, fill: { color: C.gold, transparency: 50 } });
  s8.addText(seg.q, { x: x + 0.15, y: 1.95, w: 1.85, h: 1.0, fontSize: 10, color: C.text, lineSpacingMultiple: 1.3, margin: 0 });
  s8.addText("3つの具体例：", { x: x + 0.15, y: 3.0, w: 1.85, h: 0.25, fontSize: 9, color: C.goldLight, bold: true, margin: 0 });
  s8.addText(seg.examples, { x: x + 0.15, y: 3.3, w: 1.85, h: 1.0, fontSize: 9, color: C.textMuted, lineSpacingMultiple: 1.4, margin: 0 });
  s8.addText("「なんでもいいよ」", { x: x + 0.15, y: 4.5, w: 1.85, h: 0.3, fontSize: 9, color: C.gold, italic: true, margin: 0 });
});

// ═══════════════════════════════════════
// SLIDE 9: Day1-3 関係構築
// ═══════════════════════════════════════
slideNum++;
let s9 = pres.addSlide();
addBg(s9);
addFooter(s9, slideNum, TOTAL);
s9.addText("06  Day1〜3：関係構築フェーズ", { x: 0.5, y: 0.3, w: 8, h: 0.6, fontSize: 22, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });

const days13 = [
  { day: "Day1 朝8:04", title: "音声メッセージ（30秒）", content: "あいりの肉声で信頼確定\n「続けられた理由を話したい」\n「続きは夜に送るね」", icon: "NEW", color: C.red },
  { day: "Day1 夜21:02", title: "体験談+問い", content: "「戻りたくない朝」のエピソード\n「初めて検索した日覚えてる？」\nor「誰かに相談した？言えなかった？」", icon: "", color: C.orange },
  { day: "Day2 夜21:03", title: "動画リッチメッセージ", content: "あいりが画面で語る60秒\n「3年後も変わってない自分」\n自動再生→アクションボタン", icon: "NEW", color: C.red },
  { day: "Day3 朝8:04", title: "恐怖の2択", content: "①変わらないこと vs ②失敗\nどちらでもDay4への布石\n+ Aランク返信者に手動対応", icon: "", color: C.blue },
];
days13.forEach((d, i) => {
  const y = 1.1 + i * 1.1;
  s9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 9, h: 0.95, fill: { color: C.bgLight } });
  s9.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 0.06, h: 0.95, fill: { color: d.color } });
  s9.addText(d.day, { x: 0.8, y: y + 0.05, w: 1.8, h: 0.3, fontSize: 11, color: d.color, bold: true, margin: 0 });
  s9.addText(d.title, { x: 0.8, y: y + 0.35, w: 1.8, h: 0.3, fontSize: 12, color: C.text, bold: true, margin: 0 });
  s9.addText(d.content, { x: 3.0, y: y + 0.1, w: 5, h: 0.75, fontSize: 10, color: C.textMuted, lineSpacingMultiple: 1.3, margin: 0 });
  if (d.icon === "NEW") {
    s9.addShape(pres.shapes.RECTANGLE, { x: 8.5, y: y + 0.1, w: 0.8, h: 0.25, fill: { color: C.red } });
    s9.addText("新武器", { x: 8.5, y: y + 0.1, w: 0.8, h: 0.25, fontSize: 8, color: C.white, align: "center", bold: true, margin: 0 });
  }
});

// ═══════════════════════════════════════
// SLIDE 10: Day4-9 価値観書き換え
// ═══════════════════════════════════════
slideNum++;
let s10 = pres.addSlide();
addBg(s10);
addFooter(s10, slideNum, TOTAL);
s10.addText("07  Day4〜9：価値観書き換えフェーズ", { x: 0.5, y: 0.3, w: 8, h: 0.6, fontSize: 22, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });

const days49 = [
  { day: "Day4", title: "「あなたのせいじゃない」宣言", desc: "挫折した人が最も聞きたい言葉。自己否定からの解放で強烈な信頼。" },
  { day: "Day5", title: "「考える前に動く訓練」", desc: "オープンクエスチョンで自由返信を促す。返信者=熱量の高い見込み客。" },
  { day: "Day6", title: "「今の仕事あと何年？」", desc: "欲求セグメント別に4パターン出し分け。「逃げ道が欲しかった」等の本音を代弁。" },
  { day: "Day7", title: "感情のピーク「最初の1円」", desc: "500円の案件が振り込まれた瞬間の感動。+ 師匠のセリフを初めて引用。" },
  { day: "Day8", title: "「Aちゃんの話」", desc: "独学のAちゃん vs サポートを受けた私。差は「正しいサポートの有無」だけ。" },
  { day: "Day9", title: "温度確認クイックリプライ", desc: "「わかる気がする」→加速 / 「まだ怖い」→丁寧。投資=未来の前倒し。" },
];
days49.forEach((d, i) => {
  const y = 1.1 + i * 0.72;
  s10.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 9, h: 0.62, fill: { color: i % 2 === 0 ? C.bgLight : C.bg } });
  s10.addText(d.day, { x: 0.7, y: y + 0.05, w: 0.8, h: 0.3, fontSize: 11, color: C.gold, bold: true, margin: 0 });
  s10.addText(d.title, { x: 1.6, y: y + 0.05, w: 2.8, h: 0.3, fontSize: 12, color: C.text, bold: true, margin: 0 });
  s10.addText(d.desc, { x: 4.5, y: y + 0.05, w: 4.8, h: 0.5, fontSize: 10, color: C.textMuted, margin: 0 });
});

// Day7 highlight
s10.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.7, w: 9, h: 0.6, fill: { color: C.bgLight } });
s10.addText("Day7で師匠のセリフ「才能の問題じゃない。環境と地図の問題だ」を引用 → Day11で人物を明かす = 4日間の伏線", { x: 0.7, y: 4.75, w: 8.6, h: 0.45, fontSize: 11, color: C.gold, margin: 0 });

// ═══════════════════════════════════════
// SLIDE 11: Day10-13 トスアップ + 個人LINE
// ═══════════════════════════════════════
slideNum++;
let s11 = pres.addSlide();
addBg(s11);
addFooter(s11, slideNum, TOTAL);
s11.addText("08  Day10〜13：トスアップ＋個人LINE移行", { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 22, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });

// Platform switch
s11.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 4.2, h: 1.2, fill: { color: C.bgLight } });
s11.addText("公式LINE（Day0〜9）", { x: 0.7, y: 1.15, w: 3.8, h: 0.3, fontSize: 13, color: C.blue, bold: true, margin: 0 });
s11.addText("自動配信+セグメント管理\nLステップで効率重視", { x: 0.7, y: 1.5, w: 3.8, h: 0.6, fontSize: 10, color: C.textMuted, margin: 0 });

s11.addText("→", { x: 4.6, y: 1.4, w: 0.8, h: 0.5, fontSize: 28, color: C.gold, align: "center" });

s11.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.1, w: 4.2, h: 1.2, fill: { color: C.bgLight } });
s11.addText("個人LINE（Day10〜）", { x: 5.5, y: 1.15, w: 3.8, h: 0.3, fontSize: 13, color: C.red, bold: true, margin: 0 });
s11.addText("1対1チャット+音声通話\n成約率重視・全て手動", { x: 5.5, y: 1.5, w: 3.8, h: 0.6, fontSize: 10, color: C.textMuted, margin: 0 });

// Day10-13 flow
const tossup = [
  { day: "Day10", msg: "「ここから先は公式LINEではできない話」\n→ 個人LINE誘導（移行率70-80%）" },
  { day: "Day11", msg: "師匠の人物を明かす\n「才能じゃない。環境と地図の問題だ」の伏線回収" },
  { day: "Day12", msg: "意図的な空白（配信なし）\n個人LINEだからこそ不在感が強烈に効く" },
  { day: "Day13", msg: "逆説的クロージング\n「無理にとは言わない。焦る必要はない」" },
];
tossup.forEach((t, i) => {
  const y = 2.6 + i * 0.72;
  s11.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 9, h: 0.62, fill: { color: i % 2 === 0 ? C.bgLight : C.bg } });
  s11.addText(t.day, { x: 0.7, y: y + 0.05, w: 1, h: 0.3, fontSize: 12, color: C.gold, bold: true, margin: 0 });
  s11.addText(t.msg, { x: 1.8, y: y + 0.05, w: 7.5, h: 0.5, fontSize: 10, color: C.text, margin: 0 });
});

// ═══════════════════════════════════════
// SLIDE 12: 配信頻度とマルチフォーマット
// ═══════════════════════════════════════
slideNum++;
let s12 = pres.addSlide();
addBg(s12);
addFooter(s12, slideNum, TOTAL);
s12.addText("09  配信頻度と重さ：マルチフォーマット戦略", { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 22, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });

// Format cards
const formats = [
  { title: "テキスト", use: "基本の教育配信", strong: "読み返せる安心感", weak: "長文はスルーされる", pct: "60%" },
  { title: "音声", use: "Day1朝＋加速ルート", strong: "「実在する人」の信頼", weak: "利用率5%以下で差別化", pct: "25%" },
  { title: "動画", use: "Day2痛みの日", strong: "感情が直接届く", weak: "制作コスト高", pct: "15%" },
];
formats.forEach((f, i) => {
  const x = 0.5 + i * 3.15;
  s12.addShape(pres.shapes.RECTANGLE, { x, y: 1.1, w: 2.95, h: 2.3, fill: { color: C.bgLight } });
  s12.addText(f.title, { x, y: 1.2, w: 2.95, h: 0.4, fontSize: 18, fontFace: "Arial Black", color: C.gold, align: "center", bold: true, margin: 0 });
  s12.addText(f.pct, { x: x + 1.8, y: 1.2, w: 1, h: 0.35, fontSize: 12, color: C.textMuted, align: "right", margin: 0 });
  s12.addText(f.use, { x: x + 0.2, y: 1.7, w: 2.55, h: 0.3, fontSize: 10, color: C.text, margin: 0 });
  s12.addText(f.strong, { x: x + 0.2, y: 2.1, w: 2.55, h: 0.3, fontSize: 10, color: C.green, margin: 0 });
  s12.addText(f.weak, { x: x + 0.2, y: 2.5, w: 2.55, h: 0.3, fontSize: 10, color: C.textMuted, margin: 0 });
});

// Frequency table
s12.addText("配信頻度の設計", { x: 0.5, y: 3.6, w: 3, h: 0.35, fontSize: 14, color: C.text, bold: true, margin: 0 });
const freqData = [
  ["Day0", "4通", "重1+中1+軽1+双方向2"],
  ["Day1-3", "1-2通/日", "重2+中2+軽1"],
  ["Day4-7", "1-2通/日", "重4+軽1+手動1"],
  ["Day8-9", "1通/日", "重2"],
  ["Day10-12", "0-1通/日", "中2+空白1日"],
  ["Day13", "1通", "重1（決断）"],
];
freqData.forEach((row, i) => {
  const y = 4.05 + i * 0.24;
  s12.addText(row[0], { x: 0.7, y, w: 1.2, h: 0.22, fontSize: 9, color: C.gold, bold: true, margin: 0 });
  s12.addText(row[1], { x: 2.0, y, w: 1.2, h: 0.22, fontSize: 9, color: C.text, margin: 0 });
  s12.addText(row[2], { x: 3.3, y, w: 3, h: 0.22, fontSize: 9, color: C.textMuted, margin: 0 });
});

// Time shift
s12.addShape(pres.shapes.RECTANGLE, { x: 6.8, y: 3.6, w: 2.7, h: 1.6, fill: { color: C.bgLight } });
s12.addText("配信時間の微ズラし", { x: 7.0, y: 3.7, w: 2.3, h: 0.3, fontSize: 11, color: C.red, bold: true, margin: 0 });
s12.addText("8:00 → 8:04\n12:00 → 12:03\n21:00 → 21:02\n\n他社の通知と被らず\n通知欄の最上位に表示", { x: 7.0, y: 4.05, w: 2.3, h: 1.0, fontSize: 9, color: C.textMuted, lineSpacingMultiple: 1.3, margin: 0 });

// ═══════════════════════════════════════
// SLIDE 13: 5つの心理原則
// ═══════════════════════════════════════
slideNum++;
let s13 = pres.addSlide();
addBg(s13);
addFooter(s13, slideNum, TOTAL);
s13.addText("10  5つの心理原則：合法×最大効果", { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 22, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });

const principles = [
  { name: "孤独の解消", power: "★★★★★", evil: "依存させて搾取", good: "本当に価値ある環境に繋げる" },
  { name: "コミュニティ", power: "★★★★★", evil: "相互監視で抜けられない", good: "相互成長の場を作る" },
  { name: "逆説的CL", power: "★★★★☆", evil: "操作的に「断る側」を演じる", good: "本当に無理強いしない" },
  { name: "段階的上昇", power: "★★★★☆", evil: "最終金額を隠す", good: "各段階で価値提供+金額明示" },
  { name: "権威の移転", power: "★★★☆☆", evil: "架空の実績で捏造", good: "実績と人格で信頼を得る" },
];
principles.forEach((p, i) => {
  const y = 1.1 + i * 0.88;
  s13.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 9, h: 0.78, fill: { color: i % 2 === 0 ? C.bgLight : C.bg } });
  s13.addText(p.name, { x: 0.7, y: y + 0.05, w: 1.6, h: 0.3, fontSize: 13, color: C.gold, bold: true, margin: 0 });
  s13.addText(p.power, { x: 2.3, y: y + 0.05, w: 1.2, h: 0.3, fontSize: 11, color: C.goldLight, margin: 0 });
  s13.addText("悪用：" + p.evil, { x: 3.6, y: y + 0.05, w: 3, h: 0.3, fontSize: 10, color: C.red, margin: 0 });
  s13.addText("正用：" + p.good, { x: 3.6, y: y + 0.38, w: 3, h: 0.3, fontSize: 10, color: C.green, margin: 0 });
});

// Legal boundary
s13.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.7, w: 9, h: 0.6, fill: { color: C.red, transparency: 80 } });
s13.addText("法的境界線：①契約書明記 ②クーリングオフ案内 ③借金勧誘禁止 ④返金規定遵守 ⑤特商法表記", { x: 0.7, y: 4.75, w: 8.6, h: 0.45, fontSize: 11, color: C.red, bold: true, margin: 0 });

// ═══════════════════════════════════════
// SLIDE 14: クロージングの心理構造
// ═══════════════════════════════════════
slideNum++;
let s14 = pres.addSlide();
addBg(s14);
addFooter(s14, slideNum, TOTAL);
s14.addText("10b  被害者データが証明するクロージング手法", { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 22, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });

// The killer closing sequence
const closingSteps = [
  { line: "「稼げるようになるには一人では無理」", effect: "恐怖喚起：孤独の継続への恐怖", timing: "電話アポ1回目" },
  { line: "「学べる環境にないとダメ」", effect: "唯一の解決策：他に選択肢がない", timing: "電話アポ2回目" },
  { line: "「嫌なら別に良いんですよ」", effect: "逆説的CL：追いかけない=追いかける", timing: "電話アポ3回目" },
  { line: "「本気の人だけしか相手にしたくない」", effect: "資格制限：「本気だ」と証明したくなる", timing: "電話アポ3回目" },
];
closingSteps.forEach((c, i) => {
  const y = 1.1 + i * 1.05;
  s14.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 9, h: 0.9, fill: { color: C.bgLight } });
  s14.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 0.06, h: 0.9, fill: { color: C.gold } });
  s14.addText(c.line, { x: 0.8, y: y + 0.05, w: 5, h: 0.35, fontSize: 13, color: C.text, bold: true, margin: 0 });
  s14.addText(c.effect, { x: 0.8, y: y + 0.45, w: 5, h: 0.3, fontSize: 10, color: C.textMuted, margin: 0 });
  s14.addText(c.timing, { x: 7.5, y: y + 0.2, w: 1.8, h: 0.3, fontSize: 10, color: C.goldLight, align: "right", margin: 0 });
});

// Our version
s14.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.5, w: 9, h: 0.8, fill: { color: C.green, transparency: 80 } });
s14.addText("我々の正当版：「焦る必要はない」＋「本気なら責任を持って紹介する」", { x: 0.7, y: 4.55, w: 8.6, h: 0.3, fontSize: 12, color: C.green, bold: true, margin: 0 });
s14.addText("詐欺は即断を迫る。我々は時間を与える。この1点が決定的な違い。", { x: 0.7, y: 4.85, w: 8.6, h: 0.3, fontSize: 10, color: C.textMuted, margin: 0 });

// ═══════════════════════════════════════
// SLIDE 15: チェックポイント①〜⑤
// ═══════════════════════════════════════
slideNum++;
let s15 = pres.addSlide();
addBg(s15);
addFooter(s15, slideNum, TOTAL);
s15.addText("06b  5つのチェックポイント（日数より感情段階）", { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 22, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });
s15.addText("日程は圧縮していい。感情の段階は飛ばせない。", { x: 0.5, y: 0.85, w: 9, h: 0.3, fontSize: 12, color: C.goldLight, italic: true, margin: 0 });

const checkpoints = [
  { num: "①", name: "共感成立", sign: "相手が自分の話をし始めた", skip: "表面的な関係のまま" },
  { num: "②", name: "痛みに触れた", sign: "現状への不満を言語化した", skip: "「変わりたい」が弱い" },
  { num: "③", name: "解放を感じた", sign: "「私のせいじゃないんだ」的な発言", skip: "自己否定が残り投資に踏み切れない" },
  { num: "④", name: "投資観が変わった", sign: "「お金かけてでも変わりたい」", skip: "価格提示で逃げる" },
  { num: "⑤", name: "信頼が個人に移った", sign: "「あいりさんがいうなら」", skip: "トスアップが効かない" },
];
checkpoints.forEach((cp, i) => {
  const y = 1.3 + i * 0.82;
  s15.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 9, h: 0.72, fill: { color: i % 2 === 0 ? C.bgLight : C.bg } });
  s15.addShape(pres.shapes.OVAL, { x: 0.7, y: y + 0.15, w: 0.4, h: 0.4, fill: { color: C.gold } });
  s15.addText(cp.num, { x: 0.7, y: y + 0.15, w: 0.4, h: 0.4, fontSize: 12, color: C.bg, bold: true, align: "center", valign: "middle", margin: 0 });
  s15.addText(cp.name, { x: 1.3, y: y + 0.05, w: 2, h: 0.3, fontSize: 13, color: C.text, bold: true, margin: 0 });
  s15.addText("確認：" + cp.sign, { x: 3.5, y: y + 0.05, w: 3.5, h: 0.3, fontSize: 10, color: C.goldLight, margin: 0 });
  s15.addText("飛ばすと：" + cp.skip, { x: 3.5, y: y + 0.38, w: 3.5, h: 0.3, fontSize: 10, color: C.red, margin: 0 });
});

// ═══════════════════════════════════════
// SLIDE 16: Aランク集中戦略
// ═══════════════════════════════════════
slideNum++;
let s16 = pres.addSlide();
addBg(s16);
addFooter(s16, slideNum, TOTAL);
s16.addText("06c  Aランク集中戦略：15%に80%の時間を使う", { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 22, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });

// Funnel
const ranks = [
  { rank: "A", count: "15人", pct: "15%", action: "手動チャット集中投下", time: "80%", color: C.red },
  { rank: "B", count: "15人", pct: "15%", action: "自動配信＋軽い手動フォロー", time: "15%", color: C.orange },
  { rank: "C", count: "40人", pct: "40%", action: "自動配信のみ（既読あり）", time: "5%", color: C.blue },
  { rank: "D", count: "30人", pct: "30%", action: "放置→長期育成", time: "0%", color: C.textMuted },
];
ranks.forEach((r, i) => {
  const y = 1.1 + i * 1.05;
  const w = 9 - i * 1.5;
  s16.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w, h: 0.9, fill: { color: C.bgLight } });
  s16.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 0.06, h: 0.9, fill: { color: r.color } });
  s16.addText(r.rank + "ランク", { x: 0.8, y: y + 0.05, w: 1.2, h: 0.3, fontSize: 14, color: r.color, bold: true, margin: 0 });
  s16.addText(r.count + "（" + r.pct + "）", { x: 2.0, y: y + 0.05, w: 1.5, h: 0.3, fontSize: 11, color: C.text, margin: 0 });
  s16.addText(r.action, { x: 3.5, y: y + 0.05, w: 3.5, h: 0.3, fontSize: 11, color: C.textMuted, margin: 0 });
  s16.addText("時間配分：" + r.time, { x: 0.8, y: y + 0.45, w: 3, h: 0.3, fontSize: 10, color: C.goldLight, margin: 0 });
});

// ═══════════════════════════════════════
// SLIDE 17: オペレーション設計
// ═══════════════════════════════════════
slideNum++;
let s17 = pres.addSlide();
addBg(s17);
addFooter(s17, slideNum, TOTAL);
s17.addText("11  担当者オペレーション設計", { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 22, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });

// Daily routine
s17.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 4.3, h: 2.8, fill: { color: C.bgLight } });
s17.addText("1日のルーティン（合計90分）", { x: 0.7, y: 1.2, w: 3.9, h: 0.35, fontSize: 13, color: C.gold, bold: true, margin: 0 });
const routine = [
  ["朝 8:00-8:30", "30分", "前夜の返信チェック\nAランク個別返信作成"],
  ["昼 12:00-12:15", "15分", "加速ルートのみ対応"],
  ["夜 21:00-21:45", "45分", "当日の返信チェック\n加速ルート集中チャット"],
];
routine.forEach((r, i) => {
  const y = 1.65 + i * 0.7;
  s17.addText(r[0], { x: 0.7, y, w: 1.8, h: 0.25, fontSize: 10, color: C.text, bold: true, margin: 0 });
  s17.addText(r[1], { x: 2.5, y, w: 0.6, h: 0.25, fontSize: 10, color: C.goldLight, margin: 0 });
  s17.addText(r[2], { x: 0.7, y: y + 0.3, w: 3.9, h: 0.35, fontSize: 9, color: C.textMuted, margin: 0 });
});

// Response template
s17.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.1, w: 4.3, h: 2.8, fill: { color: C.bgLight } });
s17.addText("返信の構造化（70%テンプレ+30%カスタム）", { x: 5.4, y: 1.2, w: 3.9, h: 0.35, fontSize: 12, color: C.gold, bold: true, margin: 0 });
s17.addText([
  { text: "① 相手の言葉を引用（1行）", options: { breakLine: true, fontSize: 10, color: C.text, bold: true } },
  { text: "  → 「ディズニー、いいね」= 30%カスタム", options: { breakLine: true, fontSize: 9, color: C.goldLight } },
  { text: "", options: { breakLine: true, fontSize: 6 } },
  { text: "② あいりの共感 or 体験（2-3行）", options: { breakLine: true, fontSize: 10, color: C.text, bold: true } },
  { text: "  → テンプレ体験談 = 70%テンプレ", options: { breakLine: true, fontSize: 9, color: C.goldLight } },
  { text: "", options: { breakLine: true, fontSize: 6 } },
  { text: "③ 次の問い or 予告（1行）", options: { breakLine: true, fontSize: 10, color: C.text, bold: true } },
  { text: "  → 「明日その話するね」= テンプレ", options: { fontSize: 9, color: C.goldLight } },
], { x: 5.4, y: 1.65, w: 3.9, h: 2.0, margin: 0, lineSpacingMultiple: 1.2 });

// AI roadmap
s17.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.1, w: 9, h: 1.2, fill: { color: C.bgLight } });
s17.addText("AI化ロードマップ", { x: 0.7, y: 4.2, w: 3, h: 0.3, fontSize: 13, color: C.gold, bold: true, margin: 0 });
s17.addText("現在：人が対応しデータを蓄積 → 中期：AIが下書き、人が確認 → 将来：AIが80%の質で24時間対応", { x: 0.7, y: 4.55, w: 8.5, h: 0.3, fontSize: 10, color: C.text, margin: 0 });
s17.addText("記録すべきデータ：感情段階(①-⑤) / 相手の反応（生の言葉）/ 次のアクション判断 / 個人メモ", { x: 0.7, y: 4.9, w: 8.5, h: 0.25, fontSize: 10, color: C.textMuted, margin: 0 });

// ═══════════════════════════════════════
// SLIDE 18: 手動介入タイミング
// ═══════════════════════════════════════
slideNum++;
let s18 = pres.addSlide();
addBg(s18);
addFooter(s18, slideNum, TOTAL);
s18.addText("11b  手動介入の3タイミング", { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 22, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });
s18.addText("全自動では高額商材は絶対に売れない。人が入るポイントを明確に設計する。", { x: 0.5, y: 0.85, w: 9, h: 0.3, fontSize: 11, color: C.goldLight, italic: true, margin: 0 });

const manualPoints = [
  { num: "1", day: "Day5", title: "返信者へのフォロー", detail: "自由返信が来た人にあいりが翌日中に個別返信\n1人3分 × 想定20人 = 月60分", effect: "「ちゃんと見てくれてる」信頼が爆上がり" },
  { num: "2", day: "Day9", title: "前向き者への一言", detail: "「実はあなたに話したいことがある」を手動で送信\n1人2分 × 想定30人 = 月60分", effect: "Day10以降の自動配信への期待値を上げる" },
  { num: "3", day: "Day13", title: "アポ後の一言", detail: "「楽しみにしてるね。少し緊張してる」を手動送信\n1人3分 × 想定15人 = 月45分", effect: "この一言がキャンセル率を劇的に下げる" },
];
manualPoints.forEach((mp, i) => {
  const y = 1.3 + i * 1.35;
  s18.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 9, h: 1.2, fill: { color: C.bgLight } });
  s18.addShape(pres.shapes.OVAL, { x: 0.7, y: y + 0.3, w: 0.5, h: 0.5, fill: { color: C.gold } });
  s18.addText(mp.num, { x: 0.7, y: y + 0.3, w: 0.5, h: 0.5, fontSize: 18, color: C.bg, bold: true, align: "center", valign: "middle", margin: 0 });
  s18.addText(mp.day, { x: 1.4, y: y + 0.05, w: 1, h: 0.3, fontSize: 11, color: C.gold, bold: true, margin: 0 });
  s18.addText(mp.title, { x: 2.5, y: y + 0.05, w: 3, h: 0.3, fontSize: 13, color: C.text, bold: true, margin: 0 });
  s18.addText(mp.detail, { x: 1.4, y: y + 0.4, w: 5, h: 0.5, fontSize: 10, color: C.textMuted, margin: 0 });
  s18.addText(mp.effect, { x: 6.5, y: y + 0.3, w: 2.8, h: 0.5, fontSize: 10, color: C.goldLight, italic: true, margin: 0 });
});

s18.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.7, w: 9, h: 0.6, fill: { color: C.bgLight } });
s18.addText("月100人登録 → 手動工数は月3時間 → 15人が電話アポまで進む", { x: 0.7, y: 4.75, w: 8.6, h: 0.45, fontSize: 12, color: C.gold, bold: true, margin: 0 });

// ═══════════════════════════════════════
// SLIDE 19: 数値予測とKPI
// ═══════════════════════════════════════
slideNum++;
let s19 = pres.addSlide();
addBg(s19);
addFooter(s19, slideNum, TOTAL);
s19.addText("12  数値予測とKPI", { x: 0.5, y: 0.3, w: 8, h: 0.6, fontSize: 24, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });

const funnel = [
  { step: "LP流入", count: "300人/月", rate: "—", w: 9 },
  { step: "公式LINE登録", count: "100人", rate: "33%", w: 8 },
  { step: "Day0-3 返信あり（Aランク）", count: "15人", rate: "15%", w: 6.5 },
  { step: "Day9 加速ルート選択", count: "10人", rate: "67%", w: 5.5 },
  { step: "個人LINE追加", count: "7-8人", rate: "70-80%", w: 4.5 },
  { step: "Day13 電話アポ希望", count: "5人", rate: "65%", w: 3.5 },
  { step: "3回アポ完走", count: "3-4人", rate: "70%", w: 2.5 },
  { step: "成約", count: "1-2人", rate: "30-50%", w: 1.8 },
];
funnel.forEach((f, i) => {
  const y = 1.1 + i * 0.52;
  const x = (10 - f.w) / 2;
  s19.addShape(pres.shapes.RECTANGLE, { x, y, w: f.w, h: 0.42, fill: { color: i === funnel.length - 1 ? C.gold : C.bgLight } });
  s19.addText(f.step, { x: x + 0.2, y, w: 3.5, h: 0.42, fontSize: 10, color: i === funnel.length - 1 ? C.bg : C.text, bold: true, valign: "middle", margin: 0 });
  s19.addText(f.count, { x: x + f.w - 2.5, y, w: 1.2, h: 0.42, fontSize: 10, color: i === funnel.length - 1 ? C.bg : C.goldLight, align: "right", valign: "middle", margin: 0 });
  s19.addText(f.rate, { x: x + f.w - 1.1, y, w: 0.9, h: 0.42, fontSize: 9, color: i === funnel.length - 1 ? C.bg : C.textMuted, align: "right", valign: "middle", margin: 0 });
});

// Revenue projection
s19.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.5, w: 9, h: 0.8, fill: { color: C.bgLight } });
s19.addText("月間売上予測：1〜2件成約 × 50〜200万円 = 50〜400万円/月", { x: 0.7, y: 4.55, w: 8.6, h: 0.35, fontSize: 14, color: C.gold, bold: true, margin: 0 });
s19.addText("年間：600〜4,800万円（LTV含まず）", { x: 0.7, y: 4.9, w: 8.6, h: 0.25, fontSize: 11, color: C.textMuted, margin: 0 });

// ═══════════════════════════════════════
// SLIDE 20: Next Actions
// ═══════════════════════════════════════
slideNum++;
let s20 = pres.addSlide();
addBg(s20);
s20.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.8, w: 0.06, h: 3.5, fill: { color: C.gold } });
s20.addText("Next Actions", { x: 0.9, y: 0.8, w: 5, h: 0.7, fontSize: 32, fontFace: "Arial Black", color: C.text, bold: true, margin: 0 });

const actions = [
  { priority: "即時", task: "セグメント別メッセージ全文確定", note: "12パターンの送信文を作成" },
  { priority: "即時", task: "あいりの音声メッセージ録音（30秒）", note: "Day1朝配信用" },
  { priority: "即時", task: "痛みの動画撮影（60秒）", note: "Day2夜配信用" },
  { priority: "今週", task: "担当者用日次記録シート作成", note: "Googleスプレッドシート" },
  { priority: "今週", task: "Lステップシナリオ実装", note: "16タグ・6シナリオ" },
  { priority: "今月", task: "「すごい人」人物設定+プロフィール確定", note: "電話アポ用" },
  { priority: "今月", task: "3回アポのクローザー台本設計", note: "ヒアリング→課題→CL" },
  { priority: "今月", task: "コミュニティ設計（契約後の受け皿）", note: "LTV最大化の鍵" },
];
actions.forEach((a, i) => {
  const y = 1.6 + i * 0.48;
  const priorityColor = a.priority === "即時" ? C.red : a.priority === "今週" ? C.orange : C.blue;
  s20.addShape(pres.shapes.RECTANGLE, { x: 0.9, y, w: 0.8, h: 0.35, fill: { color: priorityColor, transparency: 50 } });
  s20.addText(a.priority, { x: 0.9, y, w: 0.8, h: 0.35, fontSize: 9, color: priorityColor, bold: true, align: "center", valign: "middle", margin: 0 });
  s20.addText(a.task, { x: 1.9, y, w: 4, h: 0.35, fontSize: 11, color: C.text, bold: true, valign: "middle", margin: 0 });
  s20.addText(a.note, { x: 6.0, y, w: 3.5, h: 0.35, fontSize: 10, color: C.textMuted, valign: "middle", margin: 0 });
});

s20.addText("ZERO PROJECT  |  CONFIDENTIAL", { x: 0.9, y: 5.0, w: 4, h: 0.3, fontSize: 10, color: C.textMuted, margin: 0 });

// ═══ SAVE ═══
const outPath = "/Users/kt/Documents/zero-lp/LINE-flow-strategy-v2.pptx";
pres.writeFile({ fileName: outPath }).then(() => {
  console.log("DONE: " + outPath);
});
