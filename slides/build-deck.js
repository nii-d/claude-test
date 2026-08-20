const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";           // 13.333 x 7.5
pres.author = "人事組織";
pres.title = "案件管理と会議運営の標準化";

// ---------- palette ----------
const INK   = "16233F";  // 主色（濃紺）
const INK2  = "24365C";
const SOFT  = "F1F4F9";  // 淡い面
const CARD  = "FFFFFF";
const A_COL = "2C5F8D";  // 課題A: 案件管理
const B_COL = "1F7A6C";  // 課題B: 会議と議事録
const ACC   = "E8A33D";  // アクセント（琥珀）
const TXT   = "1A1F2B";
const MUT   = "5A6472";
const LINE  = "D5DCE6";
const NG    = "B3423A";
const OK    = "2E7D5B";
const F     = "Meiryo";

const W = 13.333, H = 7.5, ML = 0.62, CW = W - ML * 2;

// ---------- helpers ----------
const sh = () => ({ type: "outer", color: "8A97AC", blur: 10, offset: 2, angle: 90, opacity: 0.22 });

function light(title, kicker) {
  const s = pres.addSlide();
  s.background = { color: CARD };
  if (kicker) {
    s.addText(kicker, { x: ML, y: 0.36, w: CW, h: 0.3, fontFace: F, fontSize: 11.5,
      bold: true, color: ACC, charSpacing: 1.4, margin: 0 });
  }
  s.addText(title, { x: ML, y: kicker ? 0.66 : 0.5, w: CW, h: 0.62, fontFace: F,
    fontSize: 27, bold: true, color: INK, margin: 0, valign: "middle" });
  return s;
}

function dark(title, sub) {
  const s = pres.addSlide();
  s.background = { color: INK };
  return s;
}

// 丸バッジ（モチーフ）
function badge(s, x, y, d, label, fill, fsz) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill } });
  s.addText(label, { x, y, w: d, h: d, fontFace: F, fontSize: fsz || 13, bold: true,
    color: "FFFFFF", align: "center", valign: "middle", margin: 0 });
}

function card(s, x, y, w, h, fill) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.06,
    fill: { color: fill || SOFT }, shadow: sh() });
}

// ============================================================
// S1 表紙
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addShape(pres.ShapeType.ellipse, { x: 10.5, y: -1.5, w: 4.6, h: 4.6,
    fill: { color: INK2, transparency: 35 } });
  s.addShape(pres.ShapeType.ellipse, { x: 11.9, y: 5.1, w: 2.8, h: 2.8,
    fill: { color: A_COL, transparency: 62 } });

  s.addText("ご提案", { x: ML, y: 1.62, w: 6, h: 0.32, fontFace: F, fontSize: 12.5,
    bold: true, color: ACC, charSpacing: 2.2, margin: 0 });
  s.addText("案件管理と会議運営の\n標準化", { x: ML, y: 2.06, w: 9.4, h: 1.86, fontFace: F,
    fontSize: 40, bold: true, color: "FFFFFF", lineSpacing: 50, margin: 0 });
  s.addText("PMO立ち上げに合わせた全社基盤の整備", { x: ML, y: 4.06, w: 9.4, h: 0.4,
    fontFace: F, fontSize: 16, color: "C3CEE0", margin: 0 });
  s.addText("― 実行基盤としての Notion 全社導入について ―", { x: ML, y: 4.52, w: 9.4, h: 0.36,
    fontFace: F, fontSize: 13, color: "8FA0BC", margin: 0 });
  s.addText("2026年◯月◯日　人事組織", { x: ML, y: 6.3, w: 6, h: 0.34, fontFace: F,
    fontSize: 12, color: "8FA0BC", margin: 0 });
  s.addNotes("タイトルに「Notion導入」と書かない。「標準化」と書くことで、聞き手は『ツールの良し悪し』ではなく『組織の話』として聞く体勢に入る。");
}

// ============================================================
// S2 本日のお願い
// ============================================================
{
  const s = light("本日、決めていただきたいこと", "本日のお願い");
  const items = [
    ["1", "全社（140名）へのNotion導入", "案件管理と会議運営の標準化を、PMO発足と同時に開始する"],
    ["2", "年間費用 529.2万円 のご承認", "Businessプラン・年払い（プラン比較は後半でご説明）"],
    ["3", "ホールディングスへの申請着手", "セキュリティ審査の事前相談を開始する許可"],
  ];
  let y = 1.72;
  items.forEach(([n, t, d]) => {
    card(s, ML, y, CW, 1.32, SOFT);
    badge(s, ML + 0.42, y + 0.36, 0.6, n, INK, 17);
    s.addText(t, { x: ML + 1.28, y: y + 0.26, w: 8.6, h: 0.42, fontFace: F, fontSize: 18,
      bold: true, color: INK, margin: 0, valign: "middle" });
    s.addText(d, { x: ML + 1.28, y: y + 0.72, w: 10.2, h: 0.36, fontFace: F, fontSize: 12.5,
      color: MUT, margin: 0, valign: "middle" });
    y += 1.5;
  });
  s.addText("稼働目標：PMO発足と同時", { x: ML, y: 6.44, w: CW, h: 0.34, fontFace: F,
    fontSize: 13, bold: true, color: ACC, margin: 0 });
  s.addNotes("3番目を冒頭に出しておくと、「で、親会社はOKなの？」という質問を後半まで先送りできる。");
}

// ============================================================
// S3 本日の流れ
// ============================================================
{
  const s = light("本日の流れ", "AGENDA");
  const cols = [
    ["1", "なぜ必要か", "弊社が抱える2つの課題と、\nその解決像", A_COL],
    ["2", "なぜNotionか", "既存ツールで代替できない\n理由", B_COL],
    ["3", "導入判断の材料", "プラン・費用・効果・リスク・\n進め方", INK],
  ];
  const cw = 3.86, gap = 0.375;
  cols.forEach(([n, t, d, c], i) => {
    const x = ML + i * (cw + gap);
    card(s, x, 2.0, cw, 3.66, SOFT);
    badge(s, x + cw / 2 - 0.36, 2.62, 0.72, n, c, 21);
    s.addText(t, { x: x + 0.2, y: 3.66, w: cw - 0.4, h: 0.44, fontFace: F, fontSize: 17,
      bold: true, color: INK, align: "center", margin: 0 });
    s.addText(d, { x: x + 0.24, y: 4.26, w: cw - 0.48, h: 0.9, fontFace: F, fontSize: 12.5,
      color: MUT, align: "center", lineSpacing: 20, margin: 0 });
  });
  s.addText("本日は 1 と 3 を中心に、20分でご説明します。", { x: ML, y: 6.14, w: CW, h: 0.36,
    fontFace: F, fontSize: 13, color: MUT, margin: 0 });
}

// ============================================================
// セクション扉
// ============================================================
function divider(num, title, sub, col) {
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addShape(pres.ShapeType.ellipse, { x: 11.0, y: 4.4, w: 3.6, h: 3.6,
    fill: { color: col, transparency: 68 } });
  badge(s, ML, 2.66, 0.98, num, col, 30);
  s.addText(title, { x: ML + 1.34, y: 2.62, w: 9.2, h: 0.72, fontFace: F, fontSize: 33,
    bold: true, color: "FFFFFF", margin: 0, valign: "middle" });
  s.addText(sub, { x: ML + 1.34, y: 3.42, w: 9.2, h: 0.4, fontFace: F, fontSize: 14,
    color: "9FB0CC", margin: 0 });
  return s;
}
divider("1", "なぜ必要か", "弊社が抱える2つの課題と、その解決像", A_COL);

// ============================================================
// S5 なぜ今なのか
// ============================================================
{
  const s = light("PMOが立ち上がる今が、仕組みを整えるタイミングです", "なぜ今なのか");
  card(s, ML, 1.78, 6.0, 2.16, SOFT);
  badge(s, ML + 0.4, 2.12, 0.54, "✓", OK, 16);
  s.addText("決まっていること", { x: ML + 1.1, y: 2.1, w: 4.6, h: 0.38, fontFace: F,
    fontSize: 15, bold: true, color: INK, margin: 0, valign: "middle" });
  s.addText("PMO部署と人事組織が新設される。\n案件管理と会議運営を、組織として\n標準化する転機を迎えている。",
    { x: ML + 0.4, y: 2.78, w: 5.2, h: 0.98, fontFace: F, fontSize: 12.5, color: MUT,
      lineSpacing: 20, margin: 0 });

  card(s, ML + 6.35, 1.78, 6.0, 2.16, SOFT);
  badge(s, ML + 6.75, 2.12, 0.54, "!", NG, 16);
  s.addText("足りていないこと", { x: ML + 7.45, y: 2.1, w: 4.6, h: 0.38, fontFace: F,
    fontSize: 15, bold: true, color: INK, margin: 0, valign: "middle" });
  s.addText("PMOが使う「管制塔」が存在しない。\n案件の状況を一望できる場所が、\n現在どこにもない。",
    { x: ML + 6.75, y: 2.78, w: 5.2, h: 0.98, fontFace: F, fontSize: 12.5, color: MUT,
      lineSpacing: 20, margin: 0 });

  card(s, ML, 4.42, CW, 2.0, INK);
  s.addText("器のないままPMOを作れば、PMOは各部署に\nチャットと電話で状況を聞いて回る係になります。",
    { x: ML + 0.6, y: 4.42, w: CW - 1.2, h: 2.0, fontFace: F, fontSize: 21, bold: true,
      color: "FFFFFF", lineSpacing: 34, margin: 0, valign: "middle" });
  s.addNotes("本編で最も効く一行。反論の対象がNotionではなく『PMOという意思決定そのもの』になるため、誰も反対できない構図をつくる。");
}

// ============================================================
// S6 課題は2つ ★
// ============================================================
{
  const s = light("弊社の課題は、2つあります", "課題の全体像");
  const cw = 6.0, gap = 0.35;
  const data = [
    { c: A_COL, k: "A", t: "案件管理",
      rows: [["何が起きているか", "年間700件の研修案件が部門をまたいで進むのに、進捗を一望する場所がない"],
             ["現在の管理方法", "Teamsのグループチャット ＋ 各自オリジナルのExcel"],
             ["放置するとどうなるか", "納期遅延と手戻りが再発し、顧客の信用を損なう"]] },
    { c: B_COL, k: "B", t: "会議と議事録",
      rows: [["何が起きているか", "会議が多く、目的が曖昧なまま設定され、決定事項が後から追えない"],
             ["現在の管理方法", "AI議事録をWord・OneNoteに保存し、リンクをチャットで共有"],
             ["放置するとどうなるか", "意思決定が遅れ、同じ議論が繰り返される"]] },
  ];
  data.forEach((d, i) => {
    const x = ML + i * (cw + gap);
    card(s, x, 1.72, cw, 4.62, SOFT);
    badge(s, x + 0.34, 1.98, 0.62, d.k, d.c, 19);
    s.addText("課題" + d.k, { x: x + 1.08, y: 1.96, w: 2.2, h: 0.3, fontFace: F,
      fontSize: 11, bold: true, color: d.c, charSpacing: 1, margin: 0 });
    s.addText(d.t, { x: x + 1.08, y: 2.24, w: 4.6, h: 0.42, fontFace: F, fontSize: 20,
      bold: true, color: INK, margin: 0 });
    let y = 3.02;
    d.rows.forEach(([h, b]) => {
      s.addText(h, { x: x + 0.34, y, w: cw - 0.68, h: 0.28, fontFace: F, fontSize: 10.5,
        bold: true, color: d.c, margin: 0 });
      s.addText(b, { x: x + 0.34, y: y + 0.28, w: cw - 0.68, h: 0.62, fontFace: F,
        fontSize: 12, color: TXT, lineSpacing: 18, margin: 0 });
      y += 1.08;
    });
  });
  s.addText("この2つは、1つの仕組みで解決できます（後半でご説明します）",
    { x: ML, y: 6.56, w: CW, h: 0.4, fontFace: F, fontSize: 13.5, bold: true,
      color: ACC, margin: 0 });
  s.addNotes("資料全体の骨格。2つの課題を最初に並べて宣言することで、以降のスライドがどちらの話かが明確になる。色分け（A=青／B=緑）は最後まで一貫させる。");
}

// ============================================================
// S7 課題A 規模と実例
// ============================================================
{
  const s = light("案件に関わる人は、これからさらに増えます", "課題A：案件管理");
  badge(s, 12.1, 0.52, 0.56, "A", A_COL, 17);

  card(s, ML, 1.72, 5.5, 4.5, SOFT);
  s.addText("1案件に関わる人", { x: ML + 0.3, y: 1.9, w: 4.9, h: 0.32, fontFace: F,
    fontSize: 13.5, bold: true, color: INK, margin: 0 });

  const before = [["営業", "1〜2名", 0], ["運営", "1〜2名", 0], ["講師管理", "1名", 0],
    ["講師", "1〜2名", 0], ["＋ 教材担当部門と連携", "", 2], ["＋ コンテンツ作成部門と連携", "", 2]];
  const after = [["営業", "1〜2名", 0], ["運営", "1〜2名", 0], ["講師管理", "1名", 0],
    ["講師", "1〜2名", 0], ["PMO", "1名", 1], ["コンテンツ制作担当", "1〜2名", 1],
    ["システム担当", "1名", 1]];

  [["これまで", before, "4〜7名", MUT, ML + 0.3],
   ["今後", after, "7〜11名", A_COL, ML + 2.95]].forEach(([lab, list, total, col, x]) => {
    s.addText(lab, { x, y: 2.32, w: 2.4, h: 0.3, fontFace: F, fontSize: 11.5,
      bold: true, color: col, margin: 0 });
    let y = 2.7;
    list.forEach(([nm, cnt, kind]) => {
      if (kind !== 2) {
        s.addShape(pres.ShapeType.ellipse, { x: x + 0.04, y: y + 0.09, w: 0.12, h: 0.12,
          fill: { color: kind === 1 ? ACC : col } });
      }
      s.addText(nm + (cnt ? "  " + cnt : ""), { x: x + (kind === 2 ? 0 : 0.26), y,
        w: 2.4 - (kind === 2 ? 0 : 0.26), h: 0.3, fontFace: F,
        fontSize: kind === 2 ? 9.5 : 10.5, bold: kind === 1,
        color: kind === 2 ? "9AA5B4" : (kind === 1 ? INK : MUT), valign: "middle", margin: 0 });
      y += 0.34;
    });
    s.addText("のべ " + total, { x, y: 5.36, w: 2.4, h: 0.42, fontFace: F, fontSize: 16,
      bold: true, color: col, margin: 0 });
  });
  s.addShape(pres.ShapeType.rightArrow, { x: ML + 2.6, y: 2.36, w: 0.24, h: 0.22,
    fill: { color: LINE } });
  s.addText("PMO・コンテンツ制作・システム担当が加わり、関係者はさらに増える",
    { x: ML + 0.3, y: 5.82, w: 4.9, h: 0.32, fontFace: F, fontSize: 10.5, bold: true,
      color: ACC, margin: 0 });

  s.addText("700", { x: ML + 5.9, y: 1.72, w: 2.1, h: 0.92, fontFace: F, fontSize: 54,
    bold: true, color: A_COL, margin: 0 });
  s.addText("件 ／ 年", { x: ML + 7.75, y: 2.14, w: 1.6, h: 0.4, fontFace: F, fontSize: 15,
    bold: true, color: MUT, margin: 0, valign: "middle" });

  s.addText("現に起きたこと", { x: ML + 5.9, y: 2.86, w: 6.2, h: 0.32, fontFace: F,
    fontSize: 14, bold: true, color: INK, margin: 0 });
  const cases = [
    ["チャットでの依頼を見逃し、期日が守られなかった", "納期"],
    ["教材開発の成果物が、期日になっても出てこない", "納期・部門間の信頼"],
    ["講師選定が遅れ、機材準備が間に合わなかった", "実施品質・顧客信用"],
  ];
  let cy = 3.34;
  cases.forEach(([t, loss]) => {
    s.addShape(pres.ShapeType.roundRect, { x: ML + 5.9, y: cy, w: 6.2, h: 0.8,
      rectRadius: 0.05, fill: { color: "FFFFFF" }, line: { color: LINE, width: 1 } });
    s.addText(t, { x: ML + 6.14, y: cy + 0.08, w: 5.7, h: 0.34, fontFace: F, fontSize: 12,
      color: TXT, margin: 0, valign: "middle" });
    s.addText("失われたもの： " + loss, { x: ML + 6.14, y: cy + 0.42, w: 5.7, h: 0.28,
      fontFace: F, fontSize: 10.5, color: NG, margin: 0, valign: "middle" });
    cy += 0.96;
  });
  s.addNotes("部署名・個人名は絶対に出さない。「担当者が見逃した」ではなく「見逃せてしまう仕組みだった」と、主語を必ず仕組みに置くこと。この資料は社内に回覧される。関与者が増えるという事実は、チャットとExcelでの管理が今後さらに苦しくなることの裏づけになる。");
}

// ============================================================
// S8 課題A 構造欠陥
// ============================================================
{
  const s = light("原因は能力ではなく、3つの構造にあります", "課題A：なぜ起きるのか");
  badge(s, 12.1, 0.52, 0.56, "A", A_COL, 17);
  const items = [
    ["A-1", "可視性がない", "グループチャットは招待された人しか見られない。招待漏れは、そのまま情報からの遮断を意味する"],
    ["A-2", "横断して見えない", "Excelが各自バラバラの形式。案件間の比較も、全体の集計も物理的に不可能"],
    ["A-3", "管理手法がない", "WBS・ガントが標準装備されておらず、進捗と作業の依存関係を追えない"],
  ];
  const cw = 3.86, gap = 0.375;
  items.forEach(([n, t, d], i) => {
    const x = ML + i * (cw + gap);
    card(s, x, 1.86, cw, 3.94, SOFT);
    s.addText(n, { x: x + 0.32, y: 2.3, w: 2.0, h: 0.32, fontFace: F, fontSize: 12,
      bold: true, color: A_COL, charSpacing: 1, margin: 0 });
    s.addText(t, { x: x + 0.32, y: 2.76, w: cw - 0.64, h: 0.8, fontFace: F, fontSize: 17,
      bold: true, color: INK, lineSpacing: 24, margin: 0 });
    s.addText(d, { x: x + 0.32, y: 3.74, w: cw - 0.64, h: 1.5, fontFace: F, fontSize: 12,
      color: MUT, lineSpacing: 20, margin: 0 });
  });
  s.addText("いずれも「誰が、いつまでに、何をするのか」が一箇所にまとまっていないことから生じています。",
    { x: ML, y: 6.14, w: CW, h: 0.38, fontFace: F, fontSize: 13.5, bold: true, color: INK, margin: 0 });
}

// ============================================================
// S9 課題B 会議コスト
// ============================================================
{
  const s = light("目的が曖昧な会議に、年間約6,600万円を投じています", "課題B：会議のコスト");
  badge(s, 12.1, 0.52, 0.56, "B", B_COL, 17);

  card(s, ML, 1.74, 5.5, 2.34, INK);
  s.addText("年間の会議コスト（全社140名）", { x: ML + 0.42, y: 1.96, w: 4.7, h: 0.3,
    fontFace: F, fontSize: 12, color: "9FB0CC", margin: 0 });
  s.addText("約1億8,900万円", { x: ML + 0.42, y: 2.3, w: 4.9, h: 0.7, fontFace: F,
    fontSize: 31, bold: true, color: "FFFFFF", margin: 0 });
  s.addText("うち目的が曖昧なもの（35%）", { x: ML + 0.42, y: 3.06, w: 4.7, h: 0.3,
    fontFace: F, fontSize: 12, color: "9FB0CC", margin: 0 });
  s.addText("約6,600万円", { x: ML + 0.42, y: 3.36, w: 4.9, h: 0.56, fontFace: F,
    fontSize: 27, bold: true, color: ACC, margin: 0 });

  const rows = [
    ["MG以上", "20名", "週20時間", "5,400万円"],
    ["リーダー", "40名", "週15時間", "8,100万円"],
    ["一般社員", "80名", "週5時間", "5,400万円"],
  ];
  s.addText("階層別の内訳", { x: ML + 5.9, y: 1.78, w: 6.2, h: 0.3, fontFace: F,
    fontSize: 13, bold: true, color: INK, margin: 0 });
  let y = 2.18;
  rows.forEach(([a, b, c, d], i) => {
    s.addShape(pres.ShapeType.roundRect, { x: ML + 5.9, y, w: 6.2, h: 0.58,
      rectRadius: 0.05, fill: { color: i === 1 ? "E4EFEC" : SOFT } });
    s.addText(a, { x: ML + 6.1, y, w: 1.5, h: 0.58, fontFace: F, fontSize: 12.5,
      bold: true, color: INK, margin: 0, valign: "middle" });
    s.addText(b, { x: ML + 7.5, y, w: 1.0, h: 0.58, fontFace: F, fontSize: 12,
      color: MUT, margin: 0, valign: "middle" });
    s.addText(c, { x: ML + 8.5, y, w: 1.7, h: 0.58, fontFace: F, fontSize: 12,
      color: TXT, margin: 0, valign: "middle" });
    s.addText(d, { x: ML + 10.1, y, w: 1.85, h: 0.58, fontFace: F, fontSize: 12.5,
      bold: true, color: B_COL, align: "right", margin: 0, valign: "middle" });
    y += 0.66;
  });
  s.addText("計算式：人数 × 週の会議時間 × 45週 × 時間単価3,000円\n時間単価＝平均年収450万円 × 1.25 ÷ 1,900時間",
    { x: ML + 5.9, y: 4.24, w: 6.2, h: 0.66, fontFace: F, fontSize: 10.5, color: MUT,
      lineSpacing: 16, margin: 0 });

  s.addText("※ リーダー層の週15時間は実測にもとづく値。MG以上・一般社員は仮置きのため、実測値で更新が必要",
    { x: ML, y: 4.32, w: 5.5, h: 0.6, fontFace: F, fontSize: 10, color: MUT,
      lineSpacing: 15, margin: 0 });

  card(s, ML, 5.14, CW, 0.92, "FBF3E4");
  s.addText("削減した時間は、案件対応と提案活動に充てます。人件費の削減を目的とするものではありません。",
    { x: ML + 0.42, y: 5.14, w: CW - 0.84, h: 0.92, fontFace: F, fontSize: 13,
      bold: true, color: "7A5410", margin: 0, valign: "middle" });
  s.addNotes("「では人を減らせるのか」と必ず聞かれる。人件費が浮くとは絶対に言わないこと。言えば人員削減の話にすり替わり、現場の協力を失う。");
}

// ============================================================
// S10 課題B 議事録の現状
// ============================================================
{
  const s = light("議事録は残っているのに、決定事項が追えません", "課題B：議事録の現状");
  badge(s, 12.1, 0.52, 0.56, "B", B_COL, 17);

  s.addText("いまの流れ ― 情報が5か所を経由している", { x: ML, y: 1.72, w: CW, h: 0.32,
    fontFace: F, fontSize: 13, bold: true, color: INK, margin: 0 });
  const flow = ["Teams会議", "文字起こし", "AIで議事録生成", "Teams上のWord\nまたは OneNote", "リンクをチャットで共有"];
  const fw = 2.24, fgap = 0.2;
  flow.forEach((t, i) => {
    const x = ML + i * (fw + fgap);
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.14, w: fw, h: 0.82, rectRadius: 0.06,
      fill: { color: i === 4 ? "F7E9E7" : SOFT } });
    s.addText(t, { x: x + 0.08, y: 2.14, w: fw - 0.16, h: 0.82, fontFace: F, fontSize: 11.5,
      bold: true, color: i === 4 ? NG : INK, align: "center", valign: "middle", margin: 0 });
    if (i < 4) {
      s.addShape(pres.ShapeType.rightArrow, { x: x + fw + 0.03, y: 2.44, w: 0.15, h: 0.22,
        fill: { color: LINE } });
    }
  });
  s.addText("議事録は Teams上のWord と OneNote に分散。研修資料は Teams と SharePoint に分散している。",
    { x: ML, y: 3.06, w: CW, h: 0.32, fontFace: F, fontSize: 12, color: MUT, margin: 0 });

  const items = [
    ["B-1", "目的が事前に定義されない", "アジェンダを作る型がないため、「とりあえず集まる」定例が生き残る"],
    ["B-2", "決定が構造化されない", "決定事項・担当・期日が文章の中に埋もれ、検索も集計もできない"],
    ["B-3", "置き場所が分散する", "議事録はTeams上のWordとOneNoteに、研修資料はTeamsとSharePointに分かれている"],
  ];
  const cw = 3.86, gap = 0.375;
  items.forEach(([n, t, d], i) => {
    const x = ML + i * (cw + gap);
    card(s, x, 3.62, cw, 2.86, SOFT);
    s.addText(n, { x: x + 0.3, y: 3.9, w: 2.0, h: 0.3, fontFace: F, fontSize: 12,
      bold: true, color: B_COL, charSpacing: 1, margin: 0 });
    s.addText(t, { x: x + 0.3, y: 4.24, w: cw - 0.6, h: 0.44, fontFace: F, fontSize: 15,
      bold: true, color: INK, margin: 0 });
    s.addText(d, { x: x + 0.3, y: 4.86, w: cw - 0.6, h: 1.2, fontFace: F, fontSize: 11.5,
      color: MUT, lineSpacing: 18, margin: 0 });
  });
}

// ============================================================
// S11 Notionによる解決 ★
// ============================================================
{
  const s = light("Notionによる解決", "解決像");
  const cw = 6.0, gap = 0.35;
  const data = [
    { c: A_COL, k: "A", t: "案件管理",
      rows: [["A-1 可視性", "招待された人しか見えない", "全案件が誰からも見える"],
             ["A-2 横断性", "Excelがバラバラで集計不能", "PMOが1画面で700件を俯瞰"],
             ["A-3 管理手法", "進捗と依存関係が追えない", "遅延案件が自動で浮かび上がる"]] },
    { c: B_COL, k: "B", t: "会議と議事録",
      rows: [["B-1 目的", "とりあえず集まる定例が残る", "アジェンダを事前に書くことが型になる"],
             ["B-2 決定", "文章に埋もれて追えない", "決定・担当・期日が項目として残る"],
             ["B-3 置き場所", "4つのアプリに分散", "すべてNotionの1か所に集まる"]] },
  ];
  data.forEach((d, i) => {
    const x = ML + i * (cw + gap);
    badge(s, x, 1.62, 0.5, d.k, d.c, 15);
    s.addText("課題" + d.k + "：" + d.t, { x: x + 0.66, y: 1.6, w: 5.2, h: 0.44,
      fontFace: F, fontSize: 17, bold: true, color: INK, margin: 0, valign: "middle" });
    let y = 2.28;
    d.rows.forEach(([lab, bef, aft]) => {
      s.addText(lab, { x: x + 0.02, y, w: cw, h: 0.28, fontFace: F, fontSize: 10.5,
        bold: true, color: d.c, margin: 0 });
      s.addShape(pres.ShapeType.roundRect, { x, y: y + 0.3, w: 2.62, h: 0.72,
        rectRadius: 0.05, fill: { color: "F4F5F7" } });
      s.addText(bef, { x: x + 0.14, y: y + 0.3, w: 2.34, h: 0.72, fontFace: F,
        fontSize: 10.5, color: MUT, valign: "middle", margin: 0 });
      s.addShape(pres.ShapeType.rightArrow, { x: x + 2.72, y: y + 0.58, w: 0.18, h: 0.18,
        fill: { color: d.c } });
      s.addShape(pres.ShapeType.roundRect, { x: x + 3.0, y: y + 0.3, w: 3.0, h: 0.72,
        rectRadius: 0.05, fill: { color: d.c } });
      s.addText(aft, { x: x + 3.14, y: y + 0.3, w: 2.72, h: 0.72, fontFace: F,
        fontSize: 10.5, bold: true, color: "FFFFFF", valign: "middle", margin: 0 });
      y += 1.16;
    });
  });
  s.addText("「アジェンダが書けない会議は、開かれなくなる」 ― 目的のない会議は自然に淘汰されます",
    { x: ML, y: 5.92, w: CW, h: 0.44, fontFace: F, fontSize: 14, bold: true,
      color: B_COL, italic: true, margin: 0 });
  s.addNotes("B-1の「書けない会議は開かれなくなる」がこの提案の思想を最もよく表す。ツールが会議を減らすのではなく、アジェンダを書く行為が目的のない会議を淘汰する。口頭で強調する。");
}

// ============================================================
// S12 2つが繋がる ★
// ============================================================
{
  const s = light("議事録を書くことが、そのまま案件管理になります", "2つの課題が繋がる");
  const boxes = [
    { x: ML, t: "会議前", h: "アジェンダを\n事前に作る", n: "課題B の解決", c: B_COL },
    { x: ML + 4.0, t: "会議中・会議後", h: "議事録を書き\n決定事項・ToDoを\n項目として登録", n: "課題B の解決", c: B_COL },
  ];
  boxes.forEach(b => {
    s.addText(b.t, { x: b.x, y: 1.78, w: 3.4, h: 0.3, fontFace: F, fontSize: 11,
      bold: true, color: MUT, margin: 0 });
    s.addShape(pres.ShapeType.roundRect, { x: b.x, y: 2.12, w: 3.4, h: 1.9,
      rectRadius: 0.06, fill: { color: SOFT }, shadow: sh() });
    s.addText(b.h, { x: b.x + 0.2, y: 2.28, w: 3.0, h: 1.2, fontFace: F, fontSize: 15,
      bold: true, color: INK, align: "center", valign: "middle", lineSpacing: 24, margin: 0 });
    s.addText(b.n, { x: b.x + 0.2, y: 3.56, w: 3.0, h: 0.3, fontFace: F, fontSize: 10.5,
      bold: true, color: b.c, align: "center", margin: 0 });
  });
  s.addShape(pres.ShapeType.rightArrow, { x: ML + 3.52, y: 2.92, w: 0.36, h: 0.3,
    fill: { color: LINE } });
  s.addShape(pres.ShapeType.rightArrow, { x: ML + 7.52, y: 2.92, w: 0.36, h: 0.3,
    fill: { color: LINE } });

  s.addText("自動的に現れる場所", { x: ML + 8.0, y: 1.78, w: 4.1, h: 0.3, fontFace: F,
    fontSize: 11, bold: true, color: MUT, margin: 0 });
  const outs = ["① 案件ページ", "② 担当者のToDo一覧", "③ PMOの全案件ビュー"];
  let oy = 2.12;
  outs.forEach(t => {
    s.addShape(pres.ShapeType.roundRect, { x: ML + 8.0, y: oy, w: 4.1, h: 0.56,
      rectRadius: 0.05, fill: { color: A_COL } });
    s.addText(t, { x: ML + 8.2, y: oy, w: 3.7, h: 0.56, fontFace: F, fontSize: 12.5,
      bold: true, color: "FFFFFF", valign: "middle", margin: 0 });
    oy += 0.67;
  });
  s.addText("課題A の解決", { x: ML + 8.0, y: 4.14, w: 4.1, h: 0.3, fontFace: F,
    fontSize: 10.5, bold: true, color: A_COL, align: "center", margin: 0 });

  card(s, ML, 4.92, CW, 1.6, INK);
  s.addText("会議で決まったToDoは、転記せずにそのまま案件のタスクになります。二重入力は発生しません。",
    { x: ML + 0.5, y: 4.92, w: CW - 1.0, h: 1.6, fontFace: F, fontSize: 17, bold: true,
      color: "FFFFFF", valign: "middle", margin: 0 });
  s.addNotes("順番が重要。先に2つの課題を並べ、それぞれの解決を示した上で、ここで初めて『実は1つの仕組みだった』と繋げる。最初から統合して語ると、聞き手はどちらの課題の話かを見失う。");
}

divider("2", "なぜNotionなのか", "既存ツールで代替できない理由", B_COL);

// ============================================================
// S14 M365では足りないのか
// ============================================================
{
  const s = light("一部はMicrosoft 365で対応します。残りが本日のご相談です", "検討の経緯");
  const rows = [
    ["チャットの可視化", "Teamsのパブリックチーム／チャネル", "Notionは不要。この方針で実施します", OK],
    ["タスク・WBS・ガント（課題A）", "Planner（上位機能は追加ライセンス）", "案件情報・議事録と別の場所に持つことになる", MUT],
    ["ストック情報・議事録（課題B）", "Loop / SharePoint / OneNote", "文書であってデータではない。集計・横断が困難", MUT],
    ["案件の一覧・台帳（課題A）", "Lists", "議事録・ドキュメントと連動しない", MUT],
  ];
  s.addTable([
    [{ text: "やりたいこと", options: { bold: true, color: "FFFFFF", fill: { color: INK } } },
     { text: "Microsoft 365 での解", options: { bold: true, color: "FFFFFF", fill: { color: INK } } },
     { text: "判断", options: { bold: true, color: "FFFFFF", fill: { color: INK } } }],
    ...rows.map(([a, b, c, col]) => ([
      { text: a, options: { bold: true, color: TXT } },
      { text: b, options: { color: MUT } },
      { text: c, options: { color: col, bold: col === OK } },
    ]))
  ], {
    x: ML, y: 1.74, w: CW, colW: [3.5, 3.9, 4.71], rowH: 0.5,
    fontFace: F, fontSize: 11.5, valign: "middle", border: { pt: 1, color: LINE },
    fill: { color: "FFFFFF" }, margin: 0.08,
  });

  card(s, ML, 4.74, CW, 1.24, "F7E9E7");
  s.addText("すべてを満たそうとすると Planner ＋ Lists ＋ Loop ＋ SharePoint の併用になります。",
    { x: ML + 0.44, y: 4.86, w: CW - 0.88, h: 0.4, fontFace: F, fontSize: 14, bold: true,
      color: NG, margin: 0, valign: "middle" });
  s.addText("同じ情報を複数箇所に入力する二重管理が生まれ、どれが最新か分からなくなる。いま解決したい「分散」を、別の形で作り直すことになります。",
    { x: ML + 0.44, y: 5.28, w: CW - 0.88, h: 0.56, fontFace: F, fontSize: 12,
      color: "7A3A34", lineSpacing: 18, margin: 0 });

  card(s, ML, 6.16, CW, 0.9, INK);
  s.addText("Microsoft 365 ＝ 保管・伝達　／　Notion ＝ 構造化された実行管理　― M365は継続利用し、置き換えません",
    { x: ML + 0.44, y: 6.16, w: CW - 0.88, h: 0.9, fontFace: F, fontSize: 14, bold: true,
      color: "FFFFFF", valign: "middle", margin: 0 });
  s.addNotes("「M365で済む部分は済ませます」と正直に認めている点がこのスライドの価値。全否定すると売り込みに見える。一部を譲ることで残りの主張の信頼性が上がる。");
}

// ============================================================
// S15 Notionでなければならない一点
// ============================================================
{
  const s = light("違いは一点。議事録が「文書」ではなく「データ」であること", "選定理由");
  const cw = 6.0, gap = 0.35;
  const cols = [
    { t: "Word / OneNote / Loop", sub: "議事録の正体：文書",
      body: "決定事項もToDoも、文中に埋まった「文字」でしかない。",
      res: "転記しなければ管理に使えない。\n転記されなければ、消える。", c: MUT, fill: "F4F5F7", tc: TXT },
    { t: "Notion", sub: "議事録の正体：データベースのレコード",
      body: "ToDoは、案件に紐づいたタスクそのもの。",
      res: "議事録を書いた時点で、\n管理の入力が完了している。", c: ACC, fill: INK, tc: "FFFFFF" },
  ];
  cols.forEach((c, i) => {
    const x = ML + i * (cw + gap);
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.7, w: cw, h: 3.86, rectRadius: 0.06,
      fill: { color: c.fill }, shadow: sh() });
    s.addText(c.t, { x: x + 0.36, y: 1.94, w: cw - 0.72, h: 0.44, fontFace: F,
      fontSize: 19, bold: true, color: c.tc, margin: 0 });
    s.addText(c.sub, { x: x + 0.36, y: 2.42, w: cw - 0.72, h: 0.3, fontFace: F,
      fontSize: 11.5, bold: true, color: c.c, margin: 0 });
    s.addText(c.body, { x: x + 0.36, y: 2.86, w: cw - 0.72, h: 0.6, fontFace: F,
      fontSize: 12.5, color: i === 1 ? "C3CEE0" : MUT, lineSpacing: 20, margin: 0 });
    s.addText(c.res, { x: x + 0.36, y: 3.94, w: cw - 0.72, h: 1.2, fontFace: F,
      fontSize: 15, bold: true, color: c.tc, lineSpacing: 26, margin: 0 });
  });
  s.addText("「ドキュメントを書く機能」と「データベース機能」が同じ製品の中で繋がっていること。",
    { x: ML, y: 5.82, w: CW, h: 0.36, fontFace: F, fontSize: 14, bold: true, color: INK, margin: 0 });
  s.addText("案件 ⇔ 議事録 ⇔ タスクのリレーションと、立場別のビュー（営業／運営／講師管理／PMO）が、1つの実体の上に成立します。",
    { x: ML, y: 6.22, w: CW, h: 0.36, fontFace: F, fontSize: 12.5, color: MUT, margin: 0 });
  s.addNotes("「一元化したい」だけでは理由にならない（OneNoteに一元化すればタダ、と返される）。『文書ではなくデータになる』の一点に絞ることで、初めて他ツールとの差が説明できる。");
}

// ============================================================
// S16 他ツール比較
// ============================================================
{
  const s = light("専用ツールも検討しましたが、今回の要件には合いません", "他ツールとの比較");
  const head = ["評価軸", "Microsoft 365", "専用PM管理ツール\n（Asana / monday.com 等）", "Notion"];
  const rows = [
    ["案件管理（課題A）", "△　Plannerで一部", "◎　得意領域", "○　実務水準を満たす"],
    ["議事録との統合（課題B）", "△　別アプリになる", "△　ドキュメント機能が弱い", "◎　同一製品内で連結"],
    ["ナレッジ蓄積", "○　SharePoint", "△　不得意", "◎　得意領域"],
    ["全社員が使えるか", "◎　既に全員が保有", "△　管理者向けの設計が多い", "○　学習容易"],
    ["課題A・Bを1つで満たすか", "✕", "✕", "◎"],
  ];
  s.addTable([
    head.map(t => ({ text: t, options: { bold: true, color: "FFFFFF",
      fill: { color: t === "Notion" ? A_COL : INK } } })),
    ...rows.map(r => r.map((t, i) => ({
      text: t,
      options: i === 0 ? { bold: true, color: TXT }
        : i === 3 ? { color: A_COL, bold: true, fill: { color: "EDF3F8" } }
        : { color: MUT },
    })))
  ], {
    x: ML, y: 1.74, w: CW, colW: [3.3, 2.94, 3.3, 2.57], rowH: 0.5,
    fontFace: F, fontSize: 11.5, valign: "middle", border: { pt: 1, color: LINE },
    fill: { color: "FFFFFF" }, margin: 0.08,
  });
  card(s, ML, 5.24, CW, 1.5, SOFT);
  s.addText("専用PMツールは案件管理単体では優れていますが、議事録・ナレッジと分離するため課題Bが解決しません。",
    { x: ML + 0.44, y: 5.4, w: CW - 0.88, h: 0.4, fontFace: F, fontSize: 13,
      color: TXT, margin: 0, valign: "middle" });
  s.addText("課題AとBを1つの基盤で解くという要件を満たすのは、Notionのみです。",
    { x: ML + 0.44, y: 5.9, w: CW - 0.88, h: 0.5, fontFace: F, fontSize: 15, bold: true,
      color: INK, margin: 0, valign: "middle" });
  s.addNotes("「Notionが全部で一番」とは言わない。個別項目では他ツールが勝ることを正直に認めた上で、『2つの課題を1つで解く』という要件で選んでいると説明する。この誠実さが説得力になる。");
}

// ============================================================
// S17 実際の画面
// ============================================================
{
  const s = light("実際の画面", "デモンストレーション");
  const caps = [
    ["1", "会議ページ", "アジェンダ → 議事録 → 決定事項"],
    ["2", "生成されたタスク", "決定事項がそのままタスクに"],
    ["3", "案件ページ", "タスクが案件に紐づいた状態"],
    ["4", "PMOの横断ビュー", "全案件の進捗とガント"],
  ];
  const cw = 2.94, gap = 0.29;
  caps.forEach(([n, t, d], i) => {
    const x = ML + i * (cw + gap);
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.76, w: cw, h: 2.5, rectRadius: 0.06,
      fill: { color: SOFT }, line: { color: LINE, width: 1, dashType: "dash" } });
    s.addText("スクリーンショット\n（差し込み）", { x: x + 0.14, y: 1.76, w: cw - 0.28, h: 2.5,
      fontFace: F, fontSize: 11, color: "9AA5B4", align: "center", valign: "middle",
      lineSpacing: 18, margin: 0 });
    badge(s, x, 4.4, 0.44, n, INK, 13);
    s.addText(t, { x: x + 0.58, y: 4.38, w: cw - 0.58, h: 0.36, fontFace: F, fontSize: 13.5,
      bold: true, color: INK, margin: 0, valign: "middle" });
    s.addText(d, { x, y: 4.92, w: cw, h: 0.6, fontFace: F, fontSize: 11, color: MUT,
      lineSpacing: 17, margin: 0 });
  });
  card(s, ML, 5.72, CW, 0.82, "FBF3E4");
  s.addText("要準備：架空の研修案件で、アジェンダ → 議事録 → タスク → ガントが繋がった状態を作っておくこと",
    { x: ML + 0.42, y: 5.72, w: CW - 0.84, h: 0.82, fontFace: F, fontSize: 12.5, bold: true,
      color: "7A5410", valign: "middle", margin: 0 });
  s.addNotes("このスライドの準備の有無が提案の成否を分ける。文字10枚より画面4枚のほうが伝わる。PoC環境で必ず作っておくこと。");
}

divider("3", "導入判断の材料", "プラン・費用・効果・リスク・進め方", ACC);

// ============================================================
// プラン別 機能比較（Enterprise含む）
// ============================================================
{
  const s = light("Businessは最上位ではありません", "プラン別 機能比較");
  const head = ["", "Plus", "Business（推奨）", "Enterprise"];
  const rows = [
    ["月額（1人・年払い）", "1,650円", "3,150円", "要見積（約5,425円）", "m"],
    ["年間費用（140名）", "277.2万円", "529.2万円", "約911.4万円", "m"],
    ["SAML SSO（Microsoft 365連携）", "✕", "○", "○", "b"],
    ["プライベートチームスペース", "✕", "○", "○", "b"],
    ["ページ履歴", "30日", "90日", "要確認", "b"],
    ["SCIM（アカウント自動削除）", "✕", "✕", "○", "e"],
    ["監査ログ", "✕", "✕", "○", "e"],
    ["データの国内保管", "✕", "✕", "○（追加費用なし）", "e"],
    ["ゲスト上限", "100名", "250名", "250名〜", "b"],
  ];
  s.addTable([
    head.map((t, i) => ({ text: t, options: { bold: true, color: "FFFFFF",
      fill: { color: i === 2 ? A_COL : INK } } })),
    ...rows.map(r => {
      const kind = r[4];
      return r.slice(0, 4).map((t, i) => {
        if (i === 0) return { text: t, options: { bold: true, color: TXT } };
        const isRec = i === 2;
        const eOnly = kind === "e" && i === 3;
        return { text: t, options: {
          color: eOnly ? ACC : (isRec ? A_COL : MUT),
          bold: isRec || eOnly,
          fill: { color: isRec ? "EDF3F8" : "FFFFFF" } } };
      });
    })
  ], {
    x: ML, y: 1.62, w: CW, colW: [4.0, 2.5, 3.0, 2.61], rowH: 0.4,
    fontFace: F, fontSize: 11, valign: "middle", border: { pt: 1, color: LINE },
    fill: { color: "FFFFFF" }, margin: 0.07,
  });
  card(s, ML, 6.02, CW, 1.0, SOFT);
  s.addText("監査ログ・SCIM・国内データ保管はEnterpriseの機能です（差額は年間約382万円）。今回のスコープでは機微情報を置かない運用で代替し、Businessを選択します。",
    { x: ML + 0.44, y: 6.02, w: CW - 0.88, h: 1.0, fontFace: F, fontSize: 12.5,
      bold: true, color: INK, valign: "middle", margin: 0 });
  s.addNotes("狙いは「一番高いプランを買おうとしているわけではない」と示すこと。上位プランの機能を自分から並べ、それを使わない理由まで説明することで、価格交渉ではなく要件で選んでいる姿勢が伝わる。情シスが監査ログや国内保管を必須とした場合は、この表がそのままEnterprise切替の根拠資料になる。");
}

// ============================================================
// S19 プラン比較 ★
// ============================================================
{
  const s = light("3案を比較し、Businessプランの全社導入を推奨します", "プラン比較");
  const head = ["", "案1：Plus 全社", "案2：Business 全社", "案3：Business 上層部のみ\n＋一般社員はゲスト"];
  const rows = [
    ["年間費用（140名・年払い）", "277.2万円", "529.2万円", "見かけ上は最小"],
    ["SAML SSO（M365連携）", "✕", "○", "○（対象者のみ）"],
    ["全社員の横断ビュー", "○", "○", "✕"],
    ["監査ログ / SCIM", "✕", "✕", "✕"],
    ["判定", "却下", "推奨", "実現不可"],
  ];
  s.addTable([
    head.map((t, i) => ({ text: t, options: { bold: true, color: "FFFFFF",
      fill: { color: i === 2 ? A_COL : INK } } })),
    ...rows.map(r => r.map((t, i) => {
      const last = r[0] === "判定";
      if (i === 0) return { text: t, options: { bold: true, color: TXT,
        fill: { color: last ? "EDEFF3" : "FFFFFF" } } };
      const col = last ? (i === 2 ? OK : NG) : (i === 2 ? A_COL : MUT);
      return { text: t, options: { color: col, bold: last || i === 2,
        fill: { color: i === 2 ? "EDF3F8" : (last ? "EDEFF3" : "FFFFFF") } } };
    }))
  ], {
    x: ML, y: 1.7, w: CW, colW: [3.5, 2.7, 3.0, 2.91], rowH: 0.46,
    fontFace: F, fontSize: 12, valign: "middle", border: { pt: 1, color: LINE },
    fill: { color: "FFFFFF" }, margin: 0.08,
  });
  const notes = [
    ["案1を却下する理由", "SAML SSOが使えず、Microsoft 365アカウントと連携できない。退職者のアカウント停止が手作業となり、親会社のセキュリティ審査を通せない可能性が高い"],
    ["案3が成立しない理由", "ゲストは招待されたページしか見られず、横断ビューを持てない。課題A-1（可視性）とA-2（横断性）が解決せず、導入目的そのものが失われる"],
  ];
  const ncw = 6.0, ngap = 0.35;
  notes.forEach(([t, d], i) => {
    const x = ML + i * (ncw + ngap);
    card(s, x, 4.94, ncw, 1.6, SOFT);
    badge(s, x + 0.3, 5.14, 0.4, "!", NG, 12);
    s.addText(t, { x: x + 0.82, y: 5.12, w: ncw - 1.1, h: 0.42, fontFace: F, fontSize: 13,
      bold: true, color: NG, valign: "middle", margin: 0 });
    s.addText(d, { x: x + 0.3, y: 5.62, w: ncw - 0.6, h: 0.8, fontFace: F, fontSize: 11,
      color: MUT, lineSpacing: 16, margin: 0 });
  });
  s.addText("※ 2026年8月時点の公表価格。監査ログまたは国内データ保管が必須要件となる場合はEnterprise（約911万円／年）への切替が必要",
    { x: ML, y: 6.7, w: CW, h: 0.3, fontFace: F, fontSize: 10, color: MUT, margin: 0 });
  s.addNotes("案3は費用面だけ見ると魅力的に映るため、上層部から必ず提案される。「金額ではなく、目的が達成できないため不可」と即答すること。");
}

// ============================================================
// S20 費用
// ============================================================
{
  const s = light("費用", "投資額");
  card(s, ML, 1.74, 6.4, 2.6, INK);
  s.addText("年間ライセンス費用", { x: ML + 0.44, y: 1.98, w: 5.5, h: 0.3, fontFace: F,
    fontSize: 12, color: "9FB0CC", margin: 0 });
  s.addText("529.2万円", { x: ML + 0.44, y: 2.4, w: 5.5, h: 0.82, fontFace: F, fontSize: 40,
    bold: true, color: "FFFFFF", margin: 0 });
  s.addText("Business 年払い　3,150円 × 140名 × 12か月", { x: ML + 0.44, y: 3.42,
    w: 5.5, h: 0.32, fontFace: F, fontSize: 12, color: "C3CEE0", margin: 0 });

  card(s, ML + 6.75, 1.74, 5.6, 2.6, SOFT);
  s.addText("社内工数（現金支出なし）", { x: ML + 7.09, y: 1.96, w: 4.9, h: 0.32,
    fontFace: F, fontSize: 12.5, bold: true, color: INK, margin: 0 });
  const eff = [["初期構築（テンプレート設計・DB構築）", "◯人月"], ["教育・定着支援（説明会・3か月伴走）", "◯人月"]];
  let ey = 2.42;
  eff.forEach(([t, v]) => {
    s.addText(t, { x: ML + 7.09, y: ey, w: 3.7, h: 0.56, fontFace: F, fontSize: 11,
      color: MUT, valign: "middle", lineSpacing: 16, margin: 0 });
    s.addText(v, { x: ML + 10.8, y: ey, w: 1.2, h: 0.56, fontFace: F, fontSize: 13,
      bold: true, color: INK, align: "right", valign: "middle", margin: 0 });
    ey += 0.62;
  });
  s.addText("いずれも内製。PMO・人事組織の立ち上げ業務の一部として実施し、新規採用は行いません。",
    { x: ML + 7.09, y: 3.76, w: 4.9, h: 0.5, fontFace: F, fontSize: 10.5, color: MUT,
      lineSpacing: 15, margin: 0 });

  card(s, ML, 4.6, CW, 1.0, SOFT);
  s.addText("削減できる既存コスト", { x: ML + 0.44, y: 4.6, w: 3.0, h: 1.0, fontFace: F,
    fontSize: 12.5, bold: true, color: INK, valign: "middle", margin: 0 });
  s.addText("現在はMicrosoft 365に含まれるツールを使用しているため、直接的なライセンス削減は発生しません。",
    { x: ML + 3.5, y: 4.6, w: 8.4, h: 1.0, fontFace: F, fontSize: 12, color: MUT,
      valign: "middle", margin: 0 });

  card(s, ML, 5.82, CW, 1.0, "FBF3E4");
  s.addText("参考：監査ログまたは国内データ保管が必須要件となった場合は Enterprise 約911万円／年（要見積）",
    { x: ML + 0.44, y: 5.82, w: CW - 0.88, h: 1.0, fontFace: F, fontSize: 12.5, bold: true,
      color: "7A5410", valign: "middle", margin: 0 });
  s.addNotes("工数は金額換算して合計に足さない。内製なら現金支出は増えず、足せば投資額が膨らんで求められる効果も大きくなる。一方、書かないと「タダでできる」と誤解されるため人月では必ず明示する。");
}

// ============================================================
// S21 想定効果
// ============================================================
{
  const s = light("会議時間の削減だけで、投資額の約3.7倍の効果を見込みます", "想定効果");
  const stats = [
    ["投資額", "529.2万円", "年間ライセンス費用", INK],
    ["効果（会議のみ）", "約1,980万円", "目的が曖昧な会議の3割を削減", B_COL],
    ["倍率", "約3.7倍", "会議削減効果 ÷ 投資額", ACC],
  ];
  const cw = 3.86, gap = 0.375;
  stats.forEach(([lab, val, sub, c], i) => {
    const x = ML + i * (cw + gap);
    card(s, x, 1.74, cw, 1.9, i === 2 ? INK : SOFT);
    s.addText(lab, { x: x + 0.3, y: 1.92, w: cw - 0.6, h: 0.3, fontFace: F, fontSize: 11.5,
      bold: true, color: i === 2 ? "9FB0CC" : MUT, margin: 0 });
    s.addText(val, { x: x + 0.3, y: 2.26, w: cw - 0.6, h: 0.72, fontFace: F,
      fontSize: i === 1 ? 28 : 30, bold: true, color: i === 2 ? ACC : c, margin: 0 });
    s.addText(sub, { x: x + 0.3, y: 3.04, w: cw - 0.6, h: 0.44, fontFace: F, fontSize: 11,
      color: i === 2 ? "C3CEE0" : MUT, lineSpacing: 16, margin: 0 });
  });

  s.addText("その他の効果（今回は金額に算入していません）", { x: ML, y: 3.94, w: CW, h: 0.32,
    fontFace: F, fontSize: 13, bold: true, color: INK, margin: 0 });
  const others = [
    ["A", "手戻り・遅延の削減", "実測後に算定"],
    ["A", "PMOが全案件のリスクを事前に検知できる", "定性"],
    ["B", "属人化の解消。異動・退職時の引き継ぎコスト低減", "定性"],
    ["A", "顧客への納期遵守率向上による信用の維持", "定性"],
  ];
  let y = 4.36;
  others.forEach(([k, t, v]) => {
    s.addShape(pres.ShapeType.roundRect, { x: ML, y, w: CW, h: 0.48, rectRadius: 0.04,
      fill: { color: "FFFFFF" }, line: { color: LINE, width: 1 } });
    badge(s, ML + 0.16, y + 0.09, 0.3, k, k === "A" ? A_COL : B_COL, 10);
    s.addText(t, { x: ML + 0.62, y, w: 9.0, h: 0.48, fontFace: F, fontSize: 12,
      color: TXT, valign: "middle", margin: 0 });
    s.addText(v, { x: ML + 9.7, y, w: 2.3, h: 0.48, fontFace: F, fontSize: 11,
      color: MUT, align: "right", valign: "middle", margin: 0 });
    y += 0.56;
  });
  s.addText("※ 削減率3割は意図的に控えめな設定です", { x: ML, y: 6.66, w: CW, h: 0.3,
    fontFace: F, fontSize: 10, color: MUT, margin: 0 });
}

// ============================================================
// S22 リスクと対策
// ============================================================
{
  const s = light("想定されるリスクは3つ。すべて対策があります", "リスクと対策");
  const risks = [
    ["定着しない", "一部の人しか使わず、結局チャットに戻る",
     ["経営層で3か月試行済み", "テンプレート・ミニマニュアルを整備", "PMOが運用ルールの管理責任を持つ", "導入後3か月は人事組織が伴走"]],
    ["情報の置き場所が分散する", "Teams・Notion・SharePointのどこに何があるか分からなくなる",
     ["役割分担を1枚の図で定義し全社に周知", "「会話はTeams、記録はNotion」を徹底", "SharePointは当面そのまま併存"]],
    ["セキュリティ", "親会社の基準を満たせない",
     ["SOC 2 Type II・ISO 27001 等を取得済み", "SAML SSOでMicrosoft 365と統合", "機微情報は置かない運用とする", "詳細は別添にて説明"]],
  ];
  const cw = 3.86, gap = 0.375;
  risks.forEach(([t, d, ms], i) => {
    const x = ML + i * (cw + gap);
    card(s, x, 1.74, cw, 4.5, SOFT);
    badge(s, x + 0.3, 1.98, 0.44, "!", NG, 13);
    s.addText(t, { x: x + 0.86, y: 1.96, w: cw - 1.1, h: 0.48, fontFace: F, fontSize: 15,
      bold: true, color: INK, valign: "middle", margin: 0 });
    s.addText(d, { x: x + 0.3, y: 2.54, w: cw - 0.6, h: 0.66, fontFace: F, fontSize: 11,
      color: MUT, lineSpacing: 17, margin: 0 });
    s.addText("対策", { x: x + 0.3, y: 3.26, w: 1.5, h: 0.28, fontFace: F, fontSize: 10.5,
      bold: true, color: OK, margin: 0 });
    let my = 3.6;
    ms.forEach(m => {
      s.addShape(pres.ShapeType.ellipse, { x: x + 0.32, y: my + 0.09, w: 0.12, h: 0.12,
        fill: { color: OK } });
      s.addText(m, { x: x + 0.58, y: my, w: cw - 0.86, h: 0.56, fontFace: F, fontSize: 11,
        color: TXT, lineSpacing: 16, margin: 0 });
      my += 0.6;
    });
  });
}

// ============================================================
// S23 進め方
// ============================================================
{
  const s = light("ライセンスは全社、業務の移行は段階的に行います", "進め方");
  const ph = [
    ["Phase 0", "承認後すぐ", "情シスへの申請\nセキュリティ審査", "―", MUT],
    ["Phase 1", "審査通過後 1〜2か月", "会議運用から開始\nアジェンダ・議事録の標準化", "課題B", B_COL],
    ["Phase 2", "3〜4か月目", "案件管理を1部門で先行運用\nテンプレートを実案件で検証", "課題A", A_COL],
    ["Phase 3", "5か月目以降", "全案件へ展開\nPMOの横断ビュー稼働", "課題A", A_COL],
  ];
  const cw = 2.94, gap = 0.29;
  ph.forEach(([p, t, d, k, c], i) => {
    const x = ML + i * (cw + gap);
    card(s, x, 1.86, cw, 2.86, SOFT);
    s.addText(p, { x: x + 0.26, y: 2.08, w: cw - 0.52, h: 0.36, fontFace: F, fontSize: 15,
      bold: true, color: c === MUT ? INK : c, margin: 0 });
    s.addText(t, { x: x + 0.26, y: 2.46, w: cw - 0.52, h: 0.3, fontFace: F, fontSize: 10.5,
      color: MUT, margin: 0 });
    s.addText(d, { x: x + 0.26, y: 2.88, w: cw - 0.52, h: 1.0, fontFace: F, fontSize: 12,
      bold: true, color: INK, lineSpacing: 19, margin: 0 });
    if (k !== "―") {
      s.addShape(pres.ShapeType.roundRect, { x: x + 0.26, y: 4.06, w: 1.16, h: 0.34,
        rectRadius: 0.05, fill: { color: c } });
      s.addText(k, { x: x + 0.26, y: 4.06, w: 1.16, h: 0.34, fontFace: F, fontSize: 10.5,
        bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0 });
    }
    if (i < 3) {
      s.addShape(pres.ShapeType.rightArrow, { x: x + cw + 0.06, y: 3.14, w: 0.17, h: 0.22,
        fill: { color: LINE } });
    }
  });
  card(s, ML, 5.0, 6.05, 1.36, SOFT);
  s.addText("なぜ会議運用から始めるのか", { x: ML + 0.34, y: 5.14, w: 5.4, h: 0.3,
    fontFace: F, fontSize: 12, bold: true, color: B_COL, margin: 0 });
  s.addText("習得が容易で全社員が即日使え、効果が早く見える。操作に慣れた状態で案件管理に進むほうが定着率が高い。",
    { x: ML + 0.34, y: 5.46, w: 5.4, h: 0.76, fontFace: F, fontSize: 11.5, color: MUT,
      lineSpacing: 17, margin: 0 });

  card(s, ML + 6.4, 5.0, 5.95, 1.36, INK);
  s.addText("なぜライセンスは全社なのか", { x: ML + 6.74, y: 5.14, w: 5.3, h: 0.3,
    fontFace: F, fontSize: 12, bold: true, color: ACC, margin: 0 });
  s.addText("案件は部門をまたいで進むため、一部の部署だけでは外側とチャット連絡が残り、二重管理が生まれる。",
    { x: ML + 6.74, y: 5.46, w: 5.3, h: 0.76, fontFace: F, fontSize: 11.5, color: "C3CEE0",
      lineSpacing: 17, margin: 0 });
}

// ============================================================
// S24 今後の展望 ★
// ============================================================
{
  const s = light("今後の展望 ― ただし、今回は範囲を絞ります", "中長期の構想");
  s.addText("今回のスコープ", { x: ML, y: 1.76, w: 5.4, h: 0.3, fontFace: F, fontSize: 11.5,
    bold: true, color: ACC, margin: 0 });
  const now = [["①", "案件管理", A_COL], ["②", "議事録・会議", B_COL]];
  let ny = 2.12;
  now.forEach(([n, t, c]) => {
    s.addShape(pres.ShapeType.roundRect, { x: ML, y: ny, w: 4.5, h: 0.86, rectRadius: 0.06,
      fill: { color: c }, shadow: sh() });
    s.addText(n + "　" + t, { x: ML + 0.34, y: ny, w: 4.0, h: 0.86, fontFace: F,
      fontSize: 17, bold: true, color: "FFFFFF", valign: "middle", margin: 0 });
    ny += 0.98;
  });
  s.addText("まずここだけ", { x: ML, y: 4.1, w: 4.5, h: 0.3, fontFace: F, fontSize: 11,
    bold: true, color: MUT, align: "center", margin: 0 });

  s.addShape(pres.ShapeType.rightArrow, { x: ML + 4.78, y: 2.78, w: 0.5, h: 0.34,
    fill: { color: LINE } });

  s.addText("将来の構想", { x: ML + 5.6, y: 1.76, w: 5.4, h: 0.3, fontFace: F, fontSize: 11.5,
    bold: true, color: MUT, margin: 0 });
  s.addShape(pres.ShapeType.roundRect, { x: ML + 5.6, y: 2.12, w: 6.75, h: 1.84,
    rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { color: LINE, width: 1.5, dashType: "dash" } });
  s.addText("③　社内ナレッジの移行", { x: ML + 5.94, y: 2.36, w: 6.1, h: 0.42, fontFace: F,
    fontSize: 17, bold: true, color: MUT, margin: 0 });
  s.addText("現在SharePointに集約しているナレッジをNotionに統合し、\n案件・議事録・ナレッジが1か所で繋がる状態を目指したい。",
    { x: ML + 5.94, y: 2.86, w: 6.1, h: 0.84, fontFace: F, fontSize: 12, color: MUT,
      lineSpacing: 20, margin: 0 });
  s.addText("今回は実施しません", { x: ML + 5.6, y: 4.1, w: 6.75, h: 0.3, fontFace: F,
    fontSize: 11, bold: true, color: MUT, align: "center", margin: 0 });

  s.addText("なぜ絞るのか", { x: ML, y: 4.62, w: CW, h: 0.3, fontFace: F, fontSize: 12.5,
    bold: true, color: INK, margin: 0 });
  const why = [
    ["撤退可能性の確保", "思うような成果が得られなかった場合に、やめられる状態を保つ"],
    ["移行コストの回避", "SharePointの既存資産を移行すると、後戻りが困難になる"],
    ["効果検証の明確化", "範囲が狭いほど、効果が出たかどうかを判定しやすい"],
  ];
  const cw = 3.86, gap = 0.375;
  why.forEach(([t, d], i) => {
    const x = ML + i * (cw + gap);
    card(s, x, 4.98, cw, 1.38, SOFT);
    s.addText(t, { x: x + 0.28, y: 5.14, w: cw - 0.56, h: 0.34, fontFace: F, fontSize: 13,
      bold: true, color: INK, margin: 0 });
    s.addText(d, { x: x + 0.28, y: 5.5, w: cw - 0.56, h: 0.7, fontFace: F, fontSize: 11,
      color: MUT, lineSpacing: 17, margin: 0 });
  });
  s.addNotes("このスライドは守りではなく攻め。展望を語ることで『構想がある』と伝わり、同時に『今回はここまで』と自ら線を引くことで、なし崩しに全社システムを乗り換えるのではという警戒を解除できる。上層部が最も嫌う『気づいたら引き返せない』を先回りして潰す1枚。");
}

// ============================================================
// S25 成功指標と撤退基準
// ============================================================
{
  const s = light("効果が出なければ、1年で見直します", "成功指標と撤退基準");
  s.addTable([
    [{ text: "", options: { fill: { color: INK } } },
     { text: "指標", options: { bold: true, color: "FFFFFF", fill: { color: INK } } },
     { text: "現状", options: { bold: true, color: "FFFFFF", fill: { color: INK } } },
     { text: "目標", options: { bold: true, color: "FFFFFF", fill: { color: INK } } }],
    ...[["A", "案件の期日遵守率", "未計測", "計測開始 → ◯%"],
        ["A", "案件情報がNotionに登録されている割合", "0%", "95% 以上"],
        ["B", "1人あたり週の会議時間", "実測中", "▲◯%"],
        ["B", "議事録が残っている会議の割合", "実測中", "90% 以上"]]
      .map(([k, a, b, c]) => ([
        { text: k, options: { bold: true, color: "FFFFFF", align: "center",
          fill: { color: k === "A" ? A_COL : B_COL } } },
        { text: a, options: { bold: true, color: TXT } },
        { text: b, options: { color: MUT } },
        { text: c, options: { bold: true, color: INK } },
      ]))
  ], {
    x: ML, y: 1.74, w: CW, colW: [0.55, 6.0, 2.6, 2.96], rowH: 0.52,
    fontFace: F, fontSize: 12, valign: "middle", border: { pt: 1, color: LINE },
    fill: { color: "FFFFFF" }, margin: 0.08,
  });
  card(s, ML, 4.56, CW, 1.5, INK);
  s.addText("撤退基準", { x: ML + 0.5, y: 4.76, w: 3.0, h: 0.34, fontFace: F, fontSize: 13,
    bold: true, color: ACC, margin: 0 });
  s.addText("導入12か月後、上記KPIが未達の場合は翌年度の契約更新を見送ります。",
    { x: ML + 0.5, y: 5.14, w: CW - 1.0, h: 0.4, fontFace: F, fontSize: 17, bold: true,
      color: "FFFFFF", margin: 0 });
  s.addText("対象範囲を①案件管理・②議事録に限定しているため、撤退時の影響は限定的です。",
    { x: ML + 0.5, y: 5.58, w: CW - 1.0, h: 0.34, fontFace: F, fontSize: 12,
      color: "C3CEE0", margin: 0 });
  s.addText("測定時期：導入6か月後（中間）／12か月後（判定）", { x: ML, y: 6.3, w: CW, h: 0.32,
    fontFace: F, fontSize: 12, color: MUT, margin: 0 });
  s.addNotes("撤退基準を自分から示すのがこの資料で最も効く技術。「効果がなければやめます」と言える提案者は信頼される。前スライドの展望とセットで、『構想はあるが暴走はしない』という印象をつくる。");
}

// ============================================================
// S26 お願い（再掲）
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addShape(pres.ShapeType.ellipse, { x: 11.4, y: 5.0, w: 3.4, h: 3.4,
    fill: { color: A_COL, transparency: 66 } });
  s.addText("改めて、お願いする3点", { x: ML, y: 0.72, w: CW, h: 0.6, fontFace: F,
    fontSize: 28, bold: true, color: "FFFFFF", margin: 0, valign: "middle" });
  const items = [
    ["1", "全社（140名）へのNotion導入のご承認"],
    ["2", "年間費用 529.2万円 のご承認（Business・年払い）"],
    ["3", "ホールディングスへの導入申請着手のご許可"],
  ];
  let y = 1.72;
  items.forEach(([n, t]) => {
    s.addShape(pres.ShapeType.roundRect, { x: ML, y, w: CW, h: 0.98, rectRadius: 0.06,
      fill: { color: INK2 } });
    badge(s, ML + 0.34, y + 0.23, 0.52, n, ACC, 15);
    s.addText(t, { x: ML + 1.12, y, w: CW - 1.5, h: 0.98, fontFace: F, fontSize: 17,
      bold: true, color: "FFFFFF", valign: "middle", margin: 0 });
    y += 1.14;
  });
  s.addText("承認後ただちに着手すること", { x: ML, y: 5.3, w: CW, h: 0.32, fontFace: F,
    fontSize: 11.5, bold: true, color: ACC, margin: 0 });
  const acts = ["情シスへのセキュリティ要件のヒアリング", "ホールディングスへの申請書類の作成", "テンプレート設計とPMO運用ルールの策定"];
  let ay = 5.68;
  acts.forEach(t => {
    s.addShape(pres.ShapeType.ellipse, { x: ML + 0.04, y: ay + 0.1, w: 0.13, h: 0.13,
      fill: { color: "8FA0BC" } });
    s.addText(t, { x: ML + 0.34, y: ay, w: 11.0, h: 0.32, fontFace: F, fontSize: 12.5,
      color: "C3CEE0", valign: "middle", margin: 0 });
    ay += 0.4;
  });
}

// ---------- slide numbers ----------
pres.writeFile({ fileName: "/home/user/claude-test/slides/notion-proposal.pptx" })
  .then(f => console.log("written:", f));
