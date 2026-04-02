const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "ZERO PROJECT";
pres.title = "LINE教育ファネル戦略";

// Color palette
const C = {
  bg: "1A1208",
  bg2: "2C1F0E",
  gold: "B8860B",
  goldL: "C9A84C",
  text: "F5EDD8",
  text2: "B8A88A",
  text3: "7A6E5E",
  white: "FFFFFF",
  phase1: "2E6B8A",
  phase2: "C47D1A",
  phase3: "A83232",
  green: "06C755",
};

// ===== SLIDE 1: TITLE =====
let s1 = pres.addSlide();
s1.background = { color: C.bg };
s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.gold } });
s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.545, w: 10, h: 0.08, fill: { color: C.gold } });
s1.addText("LINE教育ファネル戦略", {
  x: 0.8, y: 1.2, w: 8.4, h: 1.2,
  fontSize: 40, fontFace: "Georgia", bold: true, color: C.text, align: "left", margin: 0
});
s1.addText("— 顧客やり取りの全設計 —", {
  x: 0.8, y: 2.3, w: 8.4, h: 0.6,
  fontSize: 20, fontFace: "Georgia", color: C.goldL, align: "left", margin: 0
});
s1.addShape(pres.shapes.LINE, { x: 0.8, y: 3.15, w: 2, h: 0, line: { color: C.gold, width: 2 } });
s1.addText([
  { text: "ZERO PROJECT", options: { fontSize: 14, bold: true, color: C.gold, breakLine: true } },
  { text: "高額商材（50〜200万円）/ ターゲット：副業初心者 20〜40代女性", options: { fontSize: 11, color: C.text2, breakLine: true } },
  { text: "キャラクター：あいり（30代ママ副業コーチ）", options: { fontSize: 11, color: C.text2 } }
], { x: 0.8, y: 3.4, w: 8.4, h: 1.4 });

// ===== SLIDE 2: FUNNEL FLOW =====
let s2 = pres.addSlide();
s2.background = { color: C.bg };
s2.addText("ファネル全体フロー", {
  x: 0.8, y: 0.3, w: 8.4, h: 0.7,
  fontSize: 28, fontFace: "Georgia", bold: true, color: C.text, margin: 0
});
s2.addText("Day0〜Day13の13日間で電話アポまで進める設計", {
  x: 0.8, y: 0.9, w: 8.4, h: 0.4,
  fontSize: 12, color: C.text2, margin: 0
});

// Phase labels
const phases = [
  { label: "Phase1 関係構築", sub: "Day0〜3", color: C.phase1, x: 0.5, w: 2.8 },
  { label: "Phase2 価値観書き換え", sub: "Day4〜9", color: C.phase2, x: 3.5, w: 2.8 },
  { label: "Phase3 トスアップ", sub: "Day10〜13", color: C.phase3, x: 6.5, w: 3.0 },
];
phases.forEach((p) => {
  s2.addShape(pres.shapes.RECTANGLE, { x: p.x, y: 1.5, w: p.w, h: 0.5, fill: { color: p.color } });
  s2.addText(p.label, { x: p.x, y: 1.5, w: p.w, h: 0.3, fontSize: 11, bold: true, color: C.white, align: "center", margin: 0 });
  s2.addText(p.sub, { x: p.x, y: 1.75, w: p.w, h: 0.25, fontSize: 9, color: C.white, align: "center", margin: 0 });
});

// Flow steps
const steps = [
  { label: "LP", y: 2.3 },
  { label: "LINE登録", y: 2.3 },
  { label: "あいさつ\n+PDF", y: 2.3 },
  { label: "欲求\nヒアリング", y: 2.3 },
  { label: "関係構築", y: 2.3 },
  { label: "価値観\n書き換え", y: 2.3 },
  { label: "投資観\nインストール", y: 2.3 },
  { label: "トスアップ", y: 2.3 },
  { label: "意思確認", y: 2.3 },
  { label: "電話アポ", y: 2.3 },
];
steps.forEach((st, i) => {
  const xPos = 0.3 + i * 0.96;
  const fillColor = i <= 4 ? C.phase1 : i <= 7 ? C.phase2 : C.phase3;
  s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: xPos, y: st.y, w: 0.88, h: 0.7,
    fill: { color: C.bg2 }, line: { color: fillColor, width: 1.5 }, rectRadius: 0.05,
  });
  s2.addText(st.label, {
    x: xPos, y: st.y, w: 0.88, h: 0.7,
    fontSize: 8, color: C.text, align: "center", valign: "middle", margin: 0
  });
  if (i < steps.length - 1) {
    s2.addText("→", { x: xPos + 0.85, y: st.y + 0.15, w: 0.15, h: 0.4, fontSize: 10, color: C.gold, align: "center", margin: 0 });
  }
});

// Key insight
s2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.4, w: 9.0, h: 1.8, fill: { color: C.bg2 }, line: { color: C.gold, width: 0.5 } });
s2.addText("核心設計", { x: 0.7, y: 3.5, w: 2, h: 0.4, fontSize: 12, bold: true, color: C.gold, margin: 0 });
s2.addText([
  { text: "共感 → 痛み → 解放 → 希望 → 納得 → 信頼 → 決断", options: { fontSize: 14, bold: true, color: C.goldL, breakLine: true } },
  { text: "\n高額商材は「恐怖×希望の振り子」で感情を動かす。ノウハウを教える前に考え方を変える。", options: { fontSize: 11, color: C.text2 } }
], { x: 0.7, y: 3.9, w: 8.6, h: 1.1 });

// ===== SLIDE 3: DAY0 Registration =====
let s3 = pres.addSlide();
s3.background = { color: C.bg };
s3.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.phase1 } });
s3.addText("Day0", { x: 0.4, y: 0.2, w: 1.5, h: 0.5, fontSize: 12, color: C.phase1, bold: true, margin: 0 });
s3.addText("登録直後のやり取り", {
  x: 0.4, y: 0.5, w: 9, h: 0.7,
  fontSize: 26, fontFace: "Georgia", bold: true, color: C.text, margin: 0
});
s3.addText("72時間が全てを決める", {
  x: 0.4, y: 1.1, w: 9, h: 0.35,
  fontSize: 13, color: C.goldL, margin: 0
});

// Two columns
// Left: Immediate
s3.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.7, w: 4.3, h: 3.5, fill: { color: C.bg2 }, line: { color: C.gold, width: 0.5 } });
s3.addText("即時送信", { x: 0.6, y: 1.8, w: 3, h: 0.35, fontSize: 12, bold: true, color: C.gold, margin: 0 });
s3.addText([
  { text: "あいさつメッセージ + PDF配布", options: { fontSize: 11, bold: true, color: C.text, breakLine: true } },
  { text: "\n約束を即履行して信頼を得る", options: { fontSize: 10, color: C.text2, breakLine: true } },
  { text: "\n30秒後：欲求型クイックリプライ", options: { fontSize: 11, bold: true, color: C.text, breakLine: true } },
  { text: "「副業で一番手に入れたいものは何？」", options: { fontSize: 10, italic: true, color: C.goldL } }
], { x: 0.6, y: 2.2, w: 3.9, h: 2.8 });

// Right: 4 choices
s3.addShape(pres.shapes.RECTANGLE, { x: 5.0, y: 1.7, w: 4.6, h: 3.5, fill: { color: C.bg2 }, line: { color: C.gold, width: 0.5 } });
s3.addText("4択クイックリプライ", { x: 5.2, y: 1.8, w: 3, h: 0.35, fontSize: 12, bold: true, color: C.gold, margin: 0 });

const choices = [
  { icon: "⏰", label: "自由な時間", tag: "欲求_時間" },
  { icon: "💰", label: "安定した収入", tag: "欲求_収入" },
  { icon: "🔓", label: "会社からの脱出", tag: "欲求_脱出" },
  { icon: "👨‍👩‍👧", label: "家族との時間", tag: "欲求_家族" },
];
choices.forEach((c, i) => {
  const yPos = 2.3 + i * 0.65;
  s3.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.2, y: yPos, w: 4.2, h: 0.5,
    fill: { color: "2A1F0E" }, line: { color: C.goldL, width: 0.5 }, rectRadius: 0.05
  });
  s3.addText(`${c.icon}  ${c.label}`, {
    x: 5.4, y: yPos, w: 2.5, h: 0.5,
    fontSize: 11, color: C.text, valign: "middle", margin: 0
  });
  s3.addText(`→ ${c.tag}`, {
    x: 7.8, y: yPos, w: 1.5, h: 0.5,
    fontSize: 9, color: C.text3, valign: "middle", align: "right", margin: 0
  });
});

// ===== SLIDE 4: Segment Response =====
let s4 = pres.addSlide();
s4.background = { color: C.bg };
s4.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.phase1 } });
s4.addText("Day0", { x: 0.4, y: 0.2, w: 1.5, h: 0.5, fontSize: 12, color: C.phase1, bold: true, margin: 0 });
s4.addText("セグメント別の初回応答", {
  x: 0.4, y: 0.5, w: 9, h: 0.7,
  fontSize: 26, fontFace: "Georgia", bold: true, color: C.text, margin: 0
});
s4.addText("最初の返信で「この人は私のことをわかってる」を作る", {
  x: 0.4, y: 1.1, w: 9, h: 0.35,
  fontSize: 13, color: C.goldL, margin: 0
});

const segments = [
  { tag: "時間", msg: "毎日満員電車に乗りながら\n「この時間、私には戻らない」\nって思ってたんだよね。", color: C.phase1 },
  { tag: "収入", msg: "お金の不安って、\n他のことが全部\n霞んで見えるよね。", color: C.phase2 },
  { tag: "脱出", msg: "「辞めたい」と「でも怖い」の\n間でぐるぐるしてた。\n3年間ずっと。", color: C.phase3 },
  { tag: "家族", msg: "子どもの寝顔しか\n見られない時期があって。\nそれが一番しんどかった。", color: "6D2E46" },
];
segments.forEach((seg, i) => {
  const xPos = 0.4 + i * 2.35;
  s4.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.7, w: 2.2, h: 3.3, fill: { color: C.bg2 }, line: { color: seg.color, width: 1 } });
  s4.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.7, w: 2.2, h: 0.45, fill: { color: seg.color } });
  s4.addText(seg.tag, { x: xPos, y: 1.7, w: 2.2, h: 0.45, fontSize: 12, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
  s4.addText(seg.msg, { x: xPos + 0.15, y: 2.4, w: 1.9, h: 2.0, fontSize: 10, color: C.text, italic: true, margin: 0 });
  s4.addText("→ 共感を即座に返す", { x: xPos + 0.15, y: 4.4, w: 1.9, h: 0.4, fontSize: 8, color: C.text3, margin: 0 });
});

// ===== SLIDE 5: Day1-3 =====
let s5 = pres.addSlide();
s5.background = { color: C.bg };
s5.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.phase1 } });
s5.addText("Day1〜3", { x: 0.4, y: 0.2, w: 2, h: 0.5, fontSize: 12, color: C.phase1, bold: true, margin: 0 });
s5.addText("関係構築フェーズ", {
  x: 0.4, y: 0.5, w: 9, h: 0.7,
  fontSize: 26, fontFace: "Georgia", bold: true, color: C.text, margin: 0
});

// Rule box
s5.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.3, w: 9.2, h: 0.5, fill: { color: C.phase1 } });
s5.addText("鉄則：売り込みゼロ。あいりのエピソードだけで構成する", {
  x: 0.6, y: 1.3, w: 8.8, h: 0.5,
  fontSize: 13, bold: true, color: C.white, valign: "middle", margin: 0
});

// Day cards
const days1to3 = [
  { day: "Day1", time: "朝 8:00", title: "続けられた理由", desc: "やり方じゃなかった。\n明日話すね。\n\n→ 引きを作る", emotion: "期待" },
  { day: "Day2", time: "夜 21:00", title: "3年後も変わっていない自分", desc: "痛みの底に\n連れて行く。\n\n→ 感情を揺さぶる\n→ 状況クイックリプライ\n  （未着手/挫折/行動中）", emotion: "痛み ▼" },
  { day: "Day3", time: "昼 12:00", title: "3回挫折した話", desc: "ブログ2週間で挫折。\nせどり1日で断念。\n教材買って放置。\n\n→ 等身大の失敗で\n   親近感を作る", emotion: "共感" },
];
days1to3.forEach((d, i) => {
  const xPos = 0.4 + i * 3.15;
  s5.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 2.1, w: 2.95, h: 3.2, fill: { color: C.bg2 }, line: { color: C.goldL, width: 0.5 } });
  s5.addText(d.day, { x: xPos + 0.15, y: 2.2, w: 1.2, h: 0.35, fontSize: 14, bold: true, color: C.gold, margin: 0 });
  s5.addText(d.time, { x: xPos + 1.5, y: 2.2, w: 1.3, h: 0.35, fontSize: 9, color: C.text3, align: "right", margin: 0 });
  s5.addText(d.title, { x: xPos + 0.15, y: 2.55, w: 2.65, h: 0.4, fontSize: 12, bold: true, color: C.text, margin: 0 });
  s5.addText(d.desc, { x: xPos + 0.15, y: 3.0, w: 2.65, h: 1.8, fontSize: 9, color: C.text2, margin: 0 });
  s5.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: xPos + 0.15, y: 4.85, w: 1.0, h: 0.3, fill: { color: "2A1F0E" }, line: { color: C.goldL, width: 0.5 }, rectRadius: 0.05 });
  s5.addText(d.emotion, { x: xPos + 0.15, y: 4.85, w: 1.0, h: 0.3, fontSize: 8, color: C.goldL, align: "center", valign: "middle", margin: 0 });
});

// ===== SLIDE 6: Day4-6 =====
let s6 = pres.addSlide();
s6.background = { color: C.bg };
s6.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.phase2 } });
s6.addText("Day4〜6", { x: 0.4, y: 0.2, w: 2, h: 0.5, fontSize: 12, color: C.phase2, bold: true, margin: 0 });
s6.addText("価値観を揺さぶる", {
  x: 0.4, y: 0.5, w: 9, h: 0.7,
  fontSize: 26, fontFace: "Georgia", bold: true, color: C.text, margin: 0
});

const days4to6 = [
  { day: "Day4", title: "「あなたのせいじゃない」宣言", desc: "挫折した人が最も聞きたい言葉。\n自己否定からの解放で強烈な信頼。\n\n状況セグメント別に3パターン出し分け", icon: "解放", color: C.phase2 },
  { day: "Day5", title: "「考える前に動く訓練」", desc: "オープンクエスチョンで自由返信を促す。\n返信者＝熱量の高い見込み客。\n\nタグ付与してあいりが手動フォロー", icon: "行動", color: C.phase2 },
  { day: "Day6", title: "「今の仕事あと何年？」", desc: "欲求セグメント別に4パターン出し分け。\n「逃げ道が欲しかった」等の\n本音を代弁する。", icon: "恐怖", color: C.phase2 },
];
days4to6.forEach((d, i) => {
  const xPos = 0.4 + i * 3.15;
  s6.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.4, w: 2.95, h: 3.8, fill: { color: C.bg2 }, line: { color: d.color, width: 0.5 } });
  s6.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.4, w: 2.95, h: 0.45, fill: { color: d.color } });
  s6.addText(d.day, { x: xPos, y: 1.4, w: 2.95, h: 0.45, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
  s6.addText(d.title, { x: xPos + 0.15, y: 2.0, w: 2.65, h: 0.5, fontSize: 12, bold: true, color: C.goldL, margin: 0 });
  s6.addText(d.desc, { x: xPos + 0.15, y: 2.55, w: 2.65, h: 2.0, fontSize: 10, color: C.text2, margin: 0 });
  s6.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: xPos + 0.15, y: 4.75, w: 1.0, h: 0.3, fill: { color: "2A1F0E" }, line: { color: C.phase2, width: 0.5 }, rectRadius: 0.05 });
  s6.addText(d.icon, { x: xPos + 0.15, y: 4.75, w: 1.0, h: 0.3, fontSize: 8, color: C.phase2, align: "center", valign: "middle", margin: 0 });
});

// ===== SLIDE 7: Day7 =====
let s7 = pres.addSlide();
s7.background = { color: C.bg };
s7.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.bg2, transparency: 30 } });
s7.addText("Day7", { x: 0.8, y: 0.5, w: 2, h: 0.5, fontSize: 14, color: C.phase2, bold: true, margin: 0 });
s7.addText("感情のピーク", {
  x: 0.8, y: 0.9, w: 8.4, h: 0.8,
  fontSize: 32, fontFace: "Georgia", bold: true, color: C.text, margin: 0
});
s7.addShape(pres.shapes.LINE, { x: 0.8, y: 1.8, w: 2, h: 0, line: { color: C.gold, width: 2 } });
s7.addText("「最初の1円を稼いだ日」", {
  x: 0.8, y: 2.0, w: 8.4, h: 0.6,
  fontSize: 22, fontFace: "Georgia", color: C.goldL, margin: 0
});
s7.addText("全セグメント共通の配信", {
  x: 0.8, y: 2.6, w: 8.4, h: 0.4,
  fontSize: 11, color: C.text3, margin: 0
});
// Quote
s7.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 3.2, w: 0.08, h: 1.5, fill: { color: C.gold } });
s7.addText([
  { text: "500円のライティング案件が振り込まれた瞬間、泣いた。", options: { fontSize: 14, italic: true, color: C.text, breakLine: true } },
  { text: "\n「私でも稼げる」じゃなく「私、やれてる」という感覚。", options: { fontSize: 13, italic: true, color: C.goldL, breakLine: true } },
  { text: "\nこの感動の余韻が Day8〜9 の投資観インストールへの受容性を高める", options: { fontSize: 10, color: C.text3 } }
], { x: 1.1, y: 3.2, w: 8.1, h: 1.8 });

// ===== SLIDE 8: Day8-9 =====
let s8 = pres.addSlide();
s8.background = { color: C.bg };
s8.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.phase2 } });
s8.addText("Day8〜9", { x: 0.4, y: 0.2, w: 2, h: 0.5, fontSize: 12, color: C.phase2, bold: true, margin: 0 });
s8.addText("投資への価値観インストール", {
  x: 0.4, y: 0.5, w: 9, h: 0.7,
  fontSize: 26, fontFace: "Georgia", bold: true, color: C.text, margin: 0
});
s8.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.2, w: 9.2, h: 0.45, fill: { color: C.phase3 } });
s8.addText("★ 高額商材クロージングへの最重要布石", {
  x: 0.6, y: 1.2, w: 8.8, h: 0.45,
  fontSize: 12, bold: true, color: C.white, valign: "middle", margin: 0
});

// Day8 card
s8.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.9, w: 4.5, h: 3.3, fill: { color: C.bg2 }, line: { color: C.goldL, width: 0.5 } });
s8.addText("Day8「Aちゃんの話」", { x: 0.6, y: 2.0, w: 4.1, h: 0.4, fontSize: 13, bold: true, color: C.gold, margin: 0 });
s8.addText([
  { text: "同時期に始めた友人の物語", options: { fontSize: 11, bold: true, color: C.text, breakLine: true } },
  { text: "\n独学のAちゃん vs サポートを受けた私", options: { fontSize: 10, color: C.text2, breakLine: true } },
  { text: "\n「なんで最初から教えてくれなかったの」\nと泣かれた", options: { fontSize: 10, italic: true, color: C.goldL, breakLine: true } },
  { text: "\n差は才能でも努力でもなく\n「正しいサポートの有無」だけだった", options: { fontSize: 10, bold: true, color: C.text } }
], { x: 0.6, y: 2.5, w: 4.1, h: 2.5 });

// Day9 card
s8.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 1.9, w: 4.5, h: 3.3, fill: { color: C.bg2 }, line: { color: C.goldL, width: 0.5 } });
s8.addText("Day9「時間をお金で買う」", { x: 5.3, y: 2.0, w: 4.1, h: 0.4, fontSize: 13, bold: true, color: C.gold, margin: 0 });
s8.addText([
  { text: "クイックリプライで温度確認", options: { fontSize: 11, bold: true, color: C.text, breakLine: true } },
  { text: "\n価格に直接触れず\n「投資＝未来の前倒し」をインストール", options: { fontSize: 10, color: C.text2 } }
], { x: 5.3, y: 2.5, w: 4.1, h: 1.2 });

// Route split
const routes = [
  { label: "わかる気がする", dest: "→ 加速ルート", color: C.green },
  { label: "まだちょっと怖い", dest: "→ 丁寧ルート", color: C.phase2 },
];
routes.forEach((r, i) => {
  s8.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.3, y: 3.8 + i * 0.6, w: 4.1, h: 0.45,
    fill: { color: "2A1F0E" }, line: { color: r.color, width: 1 }, rectRadius: 0.05
  });
  s8.addText(`${r.label}  ${r.dest}`, {
    x: 5.5, y: 3.8 + i * 0.6, w: 3.7, h: 0.45,
    fontSize: 10, color: C.text, valign: "middle", margin: 0
  });
});

// ===== SLIDE 9: Day10-12 =====
let s9 = pres.addSlide();
s9.background = { color: C.bg };
s9.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.phase3 } });
s9.addText("Day10〜12", { x: 0.4, y: 0.2, w: 2, h: 0.5, fontSize: 12, color: C.phase3, bold: true, margin: 0 });
s9.addText("トスアップ演出", {
  x: 0.4, y: 0.5, w: 9, h: 0.7,
  fontSize: 26, fontFace: "Georgia", bold: true, color: C.text, margin: 0
});
s9.addText("2ルート並走：加速ルート（前向き者）＋丁寧ルート（不安者）", {
  x: 0.4, y: 1.1, w: 9, h: 0.35,
  fontSize: 12, color: C.text2, margin: 0
});

const tossup = [
  { day: "Day10", title: "「私だけでは\n届けられない」告白", desc: "師匠の存在を匂わせる\n\nあいりの限界を正直に\n認めることで信頼が深まる" },
  { day: "Day11", title: "師匠のセリフ\n直接引用", desc: "「才能の問題じゃない。\n環境と地図の問題だ」\n\nDay4「あなたのせいじゃない」\nとの伏線回収" },
  { day: "Day12", title: "「今月2名限定\n話が通った」", desc: "限定感の演出\n\n「本当に変わる覚悟が\nある人だけ」という条件\n\n売り込みではなく\n「紹介の機会」として演出" },
];
tossup.forEach((d, i) => {
  const xPos = 0.4 + i * 3.15;
  s9.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.7, w: 2.95, h: 3.6, fill: { color: C.bg2 }, line: { color: C.phase3, width: 0.5 } });
  s9.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.7, w: 2.95, h: 0.45, fill: { color: C.phase3 } });
  s9.addText(d.day, { x: xPos, y: 1.7, w: 2.95, h: 0.45, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
  s9.addText(d.title, { x: xPos + 0.15, y: 2.3, w: 2.65, h: 0.7, fontSize: 12, bold: true, color: C.goldL, margin: 0 });
  s9.addText(d.desc, { x: xPos + 0.15, y: 3.1, w: 2.65, h: 2.0, fontSize: 9, color: C.text2, margin: 0 });
});

// ===== SLIDE 10: Day13 =====
let s10 = pres.addSlide();
s10.background = { color: C.bg };
s10.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 5.625, fill: { color: C.phase3 } });
s10.addText("Day13", { x: 0.4, y: 0.2, w: 2, h: 0.5, fontSize: 12, color: C.phase3, bold: true, margin: 0 });
s10.addText("意思確認 → 電話アポへ", {
  x: 0.4, y: 0.5, w: 9, h: 0.7,
  fontSize: 26, fontFace: "Georgia", bold: true, color: C.text, margin: 0
});

// 3 branches
const branches = [
  { label: "「話を聞いてみたい」", action: "日程調整URL自動送信\n+3時間後にあいりが手動で\n「楽しみにしてるね」\n\n→ この一言がキャンセル率を\n   劇的に下げる", color: C.green, dest: "→ 電話アポへ" },
  { label: "「もう少し考えたい」", action: "「焦らなくていいよ」\n\n温め継続シナリオへ移行\n\n→ 後日再度チャンスあり", color: C.phase2, dest: "→ 温め継続" },
  { label: "「今は遠慮しておく」", action: "「全然OK」\n\n長期育成シナリオへ移行\n\n→ 切り捨てず資産として残す", color: C.text3, dest: "→ 長期育成" },
];
branches.forEach((b, i) => {
  const xPos = 0.4 + i * 3.15;
  s10.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.5, w: 2.95, h: 3.8, fill: { color: C.bg2 }, line: { color: b.color, width: 1 } });
  s10.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.5, w: 2.95, h: 0.5, fill: { color: b.color } });
  s10.addText(b.label, { x: xPos, y: 1.5, w: 2.95, h: 0.5, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
  s10.addText(b.action, { x: xPos + 0.15, y: 2.2, w: 2.65, h: 2.5, fontSize: 10, color: C.text2, margin: 0 });
  s10.addText(b.dest, { x: xPos + 0.15, y: 4.8, w: 2.65, h: 0.35, fontSize: 10, bold: true, color: b.color, margin: 0 });
});

// ===== SLIDE 11: Emotion Graph =====
let s11 = pres.addSlide();
s11.background = { color: C.bg };
s11.addText("13日間の感情グラフ", {
  x: 0.8, y: 0.3, w: 8.4, h: 0.7,
  fontSize: 28, fontFace: "Georgia", bold: true, color: C.text, margin: 0
});
s11.addText("「恐怖×希望の振り子」で感情を動かす — 一度痛みを経験させないと解放の喜びが生まれない", {
  x: 0.8, y: 0.95, w: 8.4, h: 0.35,
  fontSize: 10, color: C.text2, margin: 0
});

// Graph area
s11.addShape(pres.shapes.LINE, { x: 1.0, y: 1.5, w: 0, h: 3.5, line: { color: C.text3, width: 1 } }); // Y axis
s11.addShape(pres.shapes.LINE, { x: 1.0, y: 5.0, w: 8.5, h: 0, line: { color: C.text3, width: 1 } }); // X axis

// Data points with labels
const emotionPoints = [
  { day: "Day0", emotion: "期待", y: 2.0, color: C.phase1 },
  { day: "Day2", emotion: "痛み", y: 4.2, color: C.phase3 },
  { day: "Day4", emotion: "解放", y: 2.5, color: C.phase2 },
  { day: "Day7", emotion: "希望", y: 1.6, color: C.gold },
  { day: "Day9", emotion: "納得", y: 2.2, color: C.phase2 },
  { day: "Day11", emotion: "信頼", y: 1.8, color: C.phase3 },
  { day: "Day13", emotion: "決断", y: 1.5, color: C.green },
];
emotionPoints.forEach((p, i) => {
  const xPos = 1.5 + i * 1.2;
  // Dot
  s11.addShape(pres.shapes.OVAL, { x: xPos - 0.08, y: p.y - 0.08, w: 0.16, h: 0.16, fill: { color: p.color } });
  // Label above/below
  const labelY = p.y < 3 ? p.y - 0.5 : p.y + 0.2;
  s11.addText(p.emotion, { x: xPos - 0.5, y: labelY, w: 1.0, h: 0.3, fontSize: 10, bold: true, color: p.color, align: "center", margin: 0 });
  // Day below axis
  s11.addText(p.day, { x: xPos - 0.4, y: 5.05, w: 0.8, h: 0.3, fontSize: 8, color: C.text3, align: "center", margin: 0 });
  // Connect lines
  if (i > 0) {
    const prevX = 1.5 + (i - 1) * 1.2;
    const prevY = emotionPoints[i - 1].y;
    s11.addShape(pres.shapes.LINE, {
      x: prevX, y: prevY,
      w: xPos - prevX, h: p.y - prevY,
      line: { color: C.goldL, width: 1.5 }
    });
  }
});

// Annotations
s11.addText("▲ 高い", { x: 0.3, y: 1.5, w: 0.7, h: 0.3, fontSize: 8, color: C.text3, margin: 0 });
s11.addText("▼ 低い", { x: 0.3, y: 4.5, w: 0.7, h: 0.3, fontSize: 8, color: C.text3, margin: 0 });

// ===== SLIDE 12: Manual Intervention =====
let s12 = pres.addSlide();
s12.background = { color: C.bg };
s12.addText("手動介入の3タイミング", {
  x: 0.8, y: 0.3, w: 8.4, h: 0.7,
  fontSize: 28, fontFace: "Georgia", bold: true, color: C.text, margin: 0
});
s12.addText("全自動では高額商材は絶対に売れない。人が入るポイントを明確に設計する", {
  x: 0.8, y: 0.95, w: 8.4, h: 0.35,
  fontSize: 11, color: C.text2, margin: 0
});

const manual = [
  { num: "1", timing: "Day5", title: "返信者へのフォロー", desc: "自由返信が来た人に\nあいりが翌日中に個別返信\n\n1人3分 × 想定20人\n= 月60分", why: "「ちゃんと見てくれてる人だ」\nという信頼が爆上がり" },
  { num: "2", timing: "Day9", title: "前向き者への一言", desc: "「実はあなたに話したい\nことがある」を手動で送信\n\n1人2分 × 想定30人\n= 月60分", why: "Day10〜の自動配信への\n期待値を上げておく" },
  { num: "3", timing: "Day13", title: "アポ後の一言", desc: "「楽しみにしてるね。\n少し緊張してる」を手動送信\n\n1人3分 × 想定15人\n= 月45分", why: "この一言がキャンセル率を\n劇的に下げる" },
];
manual.forEach((m, i) => {
  const xPos = 0.4 + i * 3.15;
  s12.addShape(pres.shapes.RECTANGLE, { x: xPos, y: 1.5, w: 2.95, h: 3.7, fill: { color: C.bg2 }, line: { color: C.gold, width: 0.5 } });
  // Number circle
  s12.addShape(pres.shapes.OVAL, { x: xPos + 1.15, y: 1.65, w: 0.5, h: 0.5, fill: { color: C.gold } });
  s12.addText(m.num, { x: xPos + 1.15, y: 1.65, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
  s12.addText(m.timing, { x: xPos + 0.15, y: 2.25, w: 2.65, h: 0.3, fontSize: 10, color: C.phase2, align: "center", margin: 0 });
  s12.addText(m.title, { x: xPos + 0.15, y: 2.5, w: 2.65, h: 0.35, fontSize: 13, bold: true, color: C.text, align: "center", margin: 0 });
  s12.addText(m.desc, { x: xPos + 0.15, y: 2.95, w: 2.65, h: 1.5, fontSize: 9, color: C.text2, margin: 0 });
  s12.addShape(pres.shapes.LINE, { x: xPos + 0.3, y: 4.4, w: 2.35, h: 0, line: { color: C.gold, width: 0.5, dashType: "dash" } });
  s12.addText(m.why, { x: xPos + 0.15, y: 4.5, w: 2.65, h: 0.6, fontSize: 8, italic: true, color: C.goldL, margin: 0 });
});

// Summary bar
s12.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 5.3, w: 9.2, h: 0.25, fill: { color: C.gold } });
s12.addText("月100人登録 → 手動工数は月3時間 → 15人が電話アポまで進む", {
  x: 0.6, y: 5.3, w: 8.8, h: 0.25,
  fontSize: 10, bold: true, color: C.bg, valign: "middle", align: "center", margin: 0
});

// Write file
pres.writeFile({ fileName: "/Users/kt/Documents/zero-lp/LINE-funnel-strategy.pptx" })
  .then(() => console.log("DONE: /Users/kt/Documents/zero-lp/LINE-funnel-strategy.pptx"))
  .catch(err => console.error(err));
