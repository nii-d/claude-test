const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";           // 13.333 x 7.5
pres.author = "人事組織";
pres.title = "案件管理と会議運営の標準化";

// ---------- palette ----------
const INK   = "16233F";
const INK2  = "24365C";
const SOFT  = "F1F4F9";
const A_COL = "2C5F8D";  // 課題① 案件管理
const B_COL = "1F7A6C";  // 課題② 会議と議事録
const ACC   = "E8A33D";
const TXT   = "1A1F2B";
const MUT   = "5A6472";
const LINE  = "D5DCE6";
const NG    = "B3423A";
const OK    = "2E7D5B";
const F     = "Meiryo";

const W = 13.333, H = 7.5, ML = 0.62, CW = W - ML * 2;

// ---------- helpers ----------
const sh = () => ({ type: "outer", color: "8A97AC", blur: 10, offset: 2, angle: 90, opacity: 0.22 });

function light(title, kicker, size) {
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  if (kicker) {
    s.addText(kicker, { x: ML, y: 0.34, w: CW, h: 0.3, fontFace: F, fontSize: 11.5,
      bold: true, color: ACC, charSpacing: 1.2, margin: 0 });
  }
  s.addText(title, { x: ML, y: 0.64, w: CW, h: 0.6, fontFace: F, fontSize: size || 26,
    bold: true, color: INK, margin: 0, valign: "middle" });
  return s;
}

function badge(s, x, y, d, label, fill, fsz) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill } });
  s.addText(label, { x, y, w: d, h: d, fontFace: F, fontSize: fsz || 13, bold: true,
    color: "FFFFFF", align: "center", valign: "middle", margin: 0 });
}

function card(s, x, y, w, h, fill) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.06,
    fill: { color: fill || SOFT }, shadow: sh() });
}

// 小見出し（カードの上に置くラベル）
function label(s, x, y, w, t, col) {
  s.addText(t, { x, y, w, h: 0.3, fontFace: F, fontSize: 11.5, bold: true,
    color: col || MUT, margin: 0 });
}

// 箇条書き1行（丸ポチ＋テキスト）
function bullet(s, x, y, w, t, col, fsz, h) {
  s.addShape(pres.ShapeType.ellipse, { x, y: y + 0.11, w: 0.12, h: 0.12,
    fill: { color: col || MUT } });
  s.addText(t, { x: x + 0.26, y, w: w - 0.26, h: h || 0.34, fontFace: F,
    fontSize: fsz || 11.5, color: TXT, valign: "middle", lineSpacing: 16, margin: 0 });
}

// ============================================================
// 1. 表紙
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
}

// ============================================================
// 2. アジェンダ
// ============================================================
{
  const s = light("本日の流れ", "AGENDA");
  const cols = [
    ["1", "弊社が抱える課題", "140名・フルリモートという特性と、\nそこから生じている2つの課題", A_COL],
    ["2", "Notionを選ぶ理由", "3つの選択肢を比較した結果と、\nNotionでなければならない理由", B_COL],
  ];
  const cw = 5.85, gap = 0.39;
  cols.forEach(([n, t, d, c], i) => {
    const x = ML + i * (cw + gap);
    card(s, x, 1.9, cw, 3.3, SOFT);
    badge(s, x + 0.44, 2.24, 0.76, n, c, 23);
    s.addText(t, { x: x + 1.42, y: 2.3, w: cw - 1.8, h: 0.64, fontFace: F, fontSize: 22,
      bold: true, color: INK, valign: "middle", margin: 0 });
    s.addText(d, { x: x + 0.44, y: 3.3, w: cw - 0.88, h: 1.0, fontFace: F, fontSize: 13,
      color: MUT, lineSpacing: 21, margin: 0 });
  });
  card(s, ML, 5.5, CW, 0.94, SOFT);
  s.addText("費用・プラン比較・進め方・セキュリティは Appendix に収録しています。",
    { x: ML + 0.44, y: 5.5, w: CW - 0.88, h: 0.94, fontFace: F, fontSize: 13,
      color: MUT, valign: "middle", margin: 0 });
}

// ============================================================
// 3. 前提条件
// ============================================================
{
  const s = light("PCAはTPR各社と異なり、140名・フルリモートで案件ごとにチームを組み替える会社",
    "第1章｜弊社が抱える課題", 21);

  label(s, ML, 1.56, 5.85, "PCAの特性", A_COL);
  card(s, ML, 1.88, 5.85, 1.98, SOFT);
  const traits = [
    "140名の中小規模（TPR各社は大企業）",
    "フルリモート。対面での補完が効かない",
    "案件ごとに部門横断でチームを組み替える\n年約700件の研修案件と、社内プロジェクト",
  ];
  let ty = 2.06;
  traits.forEach((t, i) => {
    bullet(s, ML + 0.34, ty, 5.2, t, A_COL, 11.5, i === 2 ? 0.6 : 0.34);
    ty += i === 1 ? 0.44 : 0.5;
  });

  label(s, ML + 6.25, 1.56, 5.85, "だから、業務にはこれが必要", INK);
  card(s, ML + 6.25, 1.88, 5.85, 1.98, SOFT);
  const needs = [
    "少人数のため、手戻りがそのまま効率に響く",
    "決定事項を口頭に頼れず、文書に残す必要がある",
    "プロジェクト管理と、離れた場所からでも\n進捗が見える状態が要る",
  ];
  let ny = 2.06;
  needs.forEach((t, i) => {
    bullet(s, ML + 6.59, ny, 5.2, t, INK, 11.5, i === 2 ? 0.6 : 0.34);
    ny += i === 1 ? 0.44 : 0.5;
  });

  label(s, ML, 4.1, CW, "現状の管理方法と、起きていること", MUT);
  s.addTable([
    [{ text: "", options: { fill: { color: INK } } },
     { text: "現状の管理方法", options: { bold: true, color: "FFFFFF", fill: { color: INK } } },
     { text: "起きていること", options: { bold: true, color: "FFFFFF", fill: { color: INK } } }],
    [{ text: "案件管理", options: { bold: true, color: "FFFFFF", fill: { color: A_COL }, align: "center" } },
     { text: "Teamsのグループチャット ＋ 各自のExcel", options: { color: TXT } },
     { text: "依頼の見落としと期日超過。進捗が横断的に見えない", options: { color: NG } }],
    [{ text: "会議・議事録", options: { bold: true, color: "FFFFFF", fill: { color: B_COL }, align: "center" } },
     { text: "会議はTeams。議事録はTeams上のWordとOneNoteに分散", options: { color: TXT } },
     { text: "目的が曖昧なまま設定される。決定が後から追えない", options: { color: NG } }],
  ], {
    x: ML, y: 4.42, w: CW, colW: [1.9, 4.9, 5.29], rowH: 0.52,
    fontFace: F, fontSize: 11.5, valign: "middle", border: { pt: 1, color: LINE },
    fill: { color: "FFFFFF" }, margin: 0.09,
  });
  s.addNotes("TPR各社との違いを最初に置くのは、聞き手（および将来ホールディングス）が大企業の感覚で判断しないようにするため。フルリモートであることが、文書に残す必要性の根拠になる。");
}

// ============================================================
// 4. 課題の全体像
// ============================================================
{
  const s = light("課題は2つ ― 案件の進捗が見えない／会議に時間が溶ける",
    "第1章｜弊社が抱える課題", 24);
  const cw = 5.85, gap = 0.39;
  const data = [
    { c: A_COL, k: "①", t: "案件管理",
      rows: [["何が起きているか", "年約700件の案件が部門をまたいで進むのに、進捗を一望する場所がない"],
             ["誰が困っているか", "営業・運営・講師管理・コンテンツ制作・システム担当、そして新設のPMO"],
             ["放置するとどうなるか", "納期遅延と手戻りが再発し、顧客の信用を損なう"]] },
    { c: B_COL, k: "②", t: "会議と議事録",
      rows: [["何が起きているか", "会議が多く、目的が曖昧なまま設定され、決定事項が後から追えない"],
             ["誰が困っているか", "全社員。とくに会議時間の長いリーダー・MG層"],
             ["放置するとどうなるか", "意思決定が遅れ、同じ議論が繰り返される"]] },
  ];
  data.forEach((d, i) => {
    const x = ML + i * (cw + gap);
    card(s, x, 1.7, cw, 4.34, SOFT);
    badge(s, x + 0.34, 1.96, 0.6, d.k, d.c, 18);
    s.addText("課題" + d.k, { x: x + 1.06, y: 1.94, w: 2.2, h: 0.28, fontFace: F,
      fontSize: 10.5, bold: true, color: d.c, margin: 0 });
    s.addText(d.t, { x: x + 1.06, y: 2.2, w: 4.4, h: 0.42, fontFace: F, fontSize: 20,
      bold: true, color: INK, margin: 0 });
    let y = 2.9;
    d.rows.forEach(([h, b]) => {
      s.addText(h, { x: x + 0.34, y, w: cw - 0.68, h: 0.28, fontFace: F, fontSize: 10.5,
        bold: true, color: d.c, margin: 0 });
      s.addText(b, { x: x + 0.34, y: y + 0.28, w: cw - 0.68, h: 0.66, fontFace: F,
        fontSize: 12, color: TXT, lineSpacing: 18, margin: 0 });
      y += 1.06;
    });
  });
  s.addText("次の2枚で、それぞれどれだけの時間と金額が失われているかをご説明します。",
    { x: ML, y: 6.24, w: CW, h: 0.38, fontFace: F, fontSize: 13, bold: true,
      color: ACC, margin: 0 });
}

// ============================================================
// 5. 課題① 案件管理
// ============================================================
{
  const s = light("課題① 案件管理 ― 1案件のべ3〜5時間が、確認の往復と緊急会議に消えている",
    "第1章｜課題① 案件管理", 21);
  badge(s, 12.1, 0.5, 0.56, "①", A_COL, 17);

  card(s, ML, 1.66, 5.5, 4.62, SOFT);
  [["700", "件 ／ 年", "研修案件（提案段階から）", ML + 0.34],
   ["7〜11", "名 ／ 案件", "PMO・コンテンツ制作・システム担当を含む", ML + 2.94]]
   .forEach(([n, u, d, x]) => {
    s.addText(n, { x, y: 1.86, w: 2.4, h: 0.6, fontFace: F, fontSize: 30, bold: true,
      color: A_COL, margin: 0 });
    s.addText(u, { x, y: 2.46, w: 2.4, h: 0.28, fontFace: F, fontSize: 11.5, bold: true,
      color: MUT, margin: 0 });
    s.addText(d, { x, y: 2.74, w: 2.4, h: 0.5, fontFace: F, fontSize: 9.5, color: MUT,
      lineSpacing: 14, margin: 0 });
  });

  s.addText("現に起きたこと", { x: ML + 0.34, y: 3.36, w: 4.9, h: 0.3, fontFace: F,
    fontSize: 13, bold: true, color: INK, margin: 0 });
  const cases = [
    ["チャットでの依頼を見逃し、期日が守られなかった", "納期"],
    ["教材開発の成果物が、期日になっても出てこない", "納期・部門間の信頼"],
    ["講師選定が遅れ、機材準備が間に合わなかった", "実施品質・顧客信用"],
  ];
  let cy = 3.74;
  cases.forEach(([t, loss]) => {
    s.addShape(pres.ShapeType.roundRect, { x: ML + 0.34, y: cy, w: 4.82, h: 0.76,
      rectRadius: 0.05, fill: { color: "FFFFFF" }, line: { color: LINE, width: 1 } });
    s.addText(t, { x: ML + 0.54, y: cy + 0.06, w: 4.4, h: 0.34, fontFace: F, fontSize: 11,
      color: TXT, valign: "middle", margin: 0 });
    s.addText("失われたもの： " + loss, { x: ML + 0.54, y: cy + 0.4, w: 4.4, h: 0.28,
      fontFace: F, fontSize: 10, color: NG, valign: "middle", margin: 0 });
    cy += 0.84;
  });

  card(s, ML + 5.9, 1.66, 6.2, 2.16, INK);
  s.addText("確認の往復・手戻り・緊急会議で失われている時間", { x: ML + 6.24, y: 1.88,
    w: 5.5, h: 0.3, fontFace: F, fontSize: 11.5, color: "9FB0CC", margin: 0 });
  s.addText("年間 約500万円", { x: ML + 6.24, y: 2.2, w: 5.5, h: 0.66, fontFace: F,
    fontSize: 30, bold: true, color: "FFFFFF", margin: 0 });
  s.addText("実施に至った420件 × のべ4時間 × 時間単価3,000円", { x: ML + 6.24, y: 2.92,
    w: 5.5, h: 0.3, fontFace: F, fontSize: 11.5, color: "C3CEE0", margin: 0 });
  s.addText("※ 提案段階の案件と社内プロジェクトは含めていない、控えめな試算です",
    { x: ML + 6.24, y: 3.26, w: 5.5, h: 0.32, fontFace: F, fontSize: 9.5,
      color: "8FA0BC", margin: 0 });

  card(s, ML + 5.9, 3.98, 6.2, 1.62, SOFT);
  s.addText("解決すると", { x: ML + 6.24, y: 4.12, w: 5.5, h: 0.3, fontFace: F,
    fontSize: 12, bold: true, color: OK, margin: 0 });
  ["全案件が同じ型で登録され、誰からも見える",
   "遅延している案件が自動で浮かび上がる",
   "PMOが1画面で全案件を把握できる"].forEach((t, i) => {
    bullet(s, ML + 6.24, 4.44 + i * 0.36, 5.5, t, OK, 11);
  });

  card(s, ML + 5.9, 5.76, 6.2, 0.52, "F7E9E7");
  s.addText("失われるのは時間だけではない ― 納期遅延は、顧客の信用に直結する",
    { x: ML + 6.24, y: 5.76, w: 5.6, h: 0.52, fontFace: F, fontSize: 11.5, bold: true,
      color: NG, valign: "middle", margin: 0 });
  s.addNotes("部署名・個人名は絶対に出さない。「担当者が見逃した」ではなく「見逃せてしまう仕組みだった」と、主語を必ず仕組みに置く。\n\n試算の分母を実施420件に絞っているのは、控えめに見せるため。「提案段階の案件を入れればもっと大きい」と口頭で補足できる。");
}

// ============================================================
// 6. 課題② 会議
// ============================================================
{
  const s = light("課題② 会議 ― 目的が曖昧な会議に、年間約4,700万円が費やされている",
    "第1章｜課題② 会議と議事録", 21);
  badge(s, 12.1, 0.5, 0.56, "②", B_COL, 17);

  card(s, ML, 1.66, 5.9, 2.76, INK);
  s.addText("年間の会議コスト（全社140名）", { x: ML + 0.4, y: 1.9, w: 5.1, h: 0.3,
    fontFace: F, fontSize: 11.5, color: "9FB0CC", margin: 0 });
  s.addText("約1億8,900万円", { x: ML + 0.4, y: 2.22, w: 5.2, h: 0.66, fontFace: F,
    fontSize: 30, bold: true, color: "FFFFFF", margin: 0 });
  s.addText("うち目的が曖昧なもの（25%）", { x: ML + 0.4, y: 3.0, w: 5.1, h: 0.3,
    fontFace: F, fontSize: 11.5, color: "9FB0CC", margin: 0 });
  s.addText("約4,700万円", { x: ML + 0.4, y: 3.32, w: 5.2, h: 0.62, fontFace: F,
    fontSize: 28, bold: true, color: ACC, margin: 0 });
  s.addText("140名 × 週の会議時間 × 45週 × 時間単価3,000円", { x: ML + 0.4, y: 4.0,
    w: 5.1, h: 0.3, fontFace: F, fontSize: 10.5, color: "8FA0BC", margin: 0 });

  card(s, ML + 6.3, 1.66, 5.8, 2.76, SOFT);
  s.addText("議事録は、こう扱われている", { x: ML + 6.64, y: 1.9, w: 5.1, h: 0.3,
    fontFace: F, fontSize: 12.5, bold: true, color: INK, margin: 0 });
  ["Teams上のWord と OneNote に分かれて保存される",
   "共有はチャットのリンク。時間が経てば遡れない",
   "決定事項・担当・期日が文章の中に埋もれる",
   "結果として、議事録は残っているのに決定が追えない"].forEach((t, i) => {
    bullet(s, ML + 6.64, 2.3 + i * 0.5, 5.1, t, B_COL, 11.5);
  });

  s.addText("※ 会議時間は各階層へのヒアリングにもとづく概算です。厳密な集計は行っていないため、規模感としてご覧ください。",
    { x: ML, y: 4.6, w: CW, h: 0.32, fontFace: F, fontSize: 10, color: MUT, margin: 0 });

  card(s, ML, 5.1, CW, 1.1, "FBF3E4");
  s.addText("削減した時間は、案件対応と提案活動に充てます。人件費の削減を目的とするものではありません。",
    { x: ML + 0.44, y: 5.1, w: CW - 0.88, h: 1.1, fontFace: F, fontSize: 13.5, bold: true,
      color: "7A5410", valign: "middle", margin: 0 });
  s.addNotes("「では人を減らせるのか」と必ず聞かれる。人件費が浮くとは絶対に言わないこと。\n\n【数字の根拠を問われたら】Outlookの予定表から抽出することは可能だが、140名分の集計に相応の工数がかかるため、今回は各階層へのヒアリングによる概算としている。精緻な数字が必要であれば別途集計する、と答える。");
}

// ============================================================
// 7. なぜ起き続けるのか
// ============================================================
{
  const s = light("なぜ2つの課題が起き続けるのか ― 情報がチャットに流れ、どこにも残らないから",
    "第2章｜Notionを選ぶ理由", 21);

  const cw = 5.85, gap = 0.39;
  const data = [
    { c: A_COL, k: "①", t: "案件管理",
      rows: [["招待された人しか見えない", "グループチャットは招待制。漏れた人はそのまま情報から遮断される"],
             ["Excelが各自バラバラ", "案件間の比較も、全体の集計も物理的にできない"],
             ["WBS・ガントがない", "進捗と作業の依存関係を追う手段がない"]] },
    { c: B_COL, k: "②", t: "会議と議事録",
      rows: [["アジェンダを作る型がない", "目的が曖昧なまま「とりあえず集まる」定例が生き残る"],
             ["決定が文章に埋もれる", "決定事項・担当・期日を検索も集計もできない"],
             ["議事録の置き場所が分かれる", "Teams上のWordとOneNoteに分散し、どこにあるか分からない"]] },
  ];
  data.forEach((d, i) => {
    const x = ML + i * (cw + gap);
    badge(s, x, 1.6, 0.5, d.k, d.c, 15);
    s.addText("課題" + d.k + "　" + d.t, { x: x + 0.66, y: 1.58, w: 5.1, h: 0.44,
      fontFace: F, fontSize: 16, bold: true, color: INK, valign: "middle", margin: 0 });
    let y = 2.22;
    d.rows.forEach(([t, b]) => {
      card(s, x, y, cw, 1.16, SOFT);
      s.addText(t, { x: x + 0.3, y: y + 0.14, w: cw - 0.6, h: 0.34, fontFace: F,
        fontSize: 13.5, bold: true, color: d.c, margin: 0 });
      s.addText(b, { x: x + 0.3, y: y + 0.5, w: cw - 0.6, h: 0.56, fontFace: F,
        fontSize: 11, color: MUT, lineSpacing: 16, margin: 0 });
      y += 1.3;
    });
  });

  card(s, ML, 6.14, CW, 0.86, INK);
  s.addText("どれも担当者の問題ではなく、情報が1か所に集まらないことから起きています。",
    { x: ML + 0.44, y: 6.14, w: CW - 0.88, h: 0.86, fontFace: F, fontSize: 14, bold: true,
      color: "FFFFFF", valign: "middle", margin: 0 });
}

// ============================================================
// 8. 3つの選択肢を比較
// ============================================================
{
  const s = light("3つの選択肢を比較した ― 2つの課題を1つで解けるのはNotionだけ",
    "第2章｜Notionを選ぶ理由", 23);

  const head = ["", "Microsoft 365", "専用のPM管理ツール\n（Asana / monday.com 等）", "Notion"];
  const rows = [
    ["課題① 案件管理", "△　Plannerで一部", "◎　得意領域", "○　実務水準を満たす"],
    ["課題② 議事録との連結", "✕　別アプリになる", "✕　ドキュメント機能が弱い", "◎　同一製品内で連結"],
    ["ナレッジの蓄積", "○　SharePoint", "△　不得意", "◎　得意領域"],
    ["全社員が使えるか", "◎　既に全員が保有", "△　管理者向けの設計が多い", "○　学習が容易"],
    ["①と②を1つで解けるか", "✕", "✕", "◎"],
  ];
  s.addTable([
    head.map((t, i) => ({ text: t, options: { bold: true, color: "FFFFFF",
      fill: { color: i === 3 ? A_COL : INK } } })),
    ...rows.map(r => r.map((t, i) => {
      const last = r[0] === "①と②を1つで解けるか";
      if (i === 0) return { text: t, options: { bold: true, color: TXT,
        fill: { color: last ? "EDEFF3" : "FFFFFF" } } };
      return { text: t, options: {
        color: i === 3 ? A_COL : MUT, bold: i === 3 || last,
        fill: { color: i === 3 ? "EDF3F8" : (last ? "EDEFF3" : "FFFFFF") } } };
    }))
  ], {
    x: ML, y: 1.62, w: CW, colW: [3.2, 2.9, 3.5, 2.49], rowH: 0.46,
    fontFace: F, fontSize: 11.5, valign: "middle", border: { pt: 1, color: LINE },
    fill: { color: "FFFFFF" }, margin: 0.08,
  });

  const notes = [
    ["Microsoft 365 で満たすと", "Planner＋Lists＋Loop＋SharePoint の併用になる。同じ情報を複数箇所に入力することになり、いま解決したい分散を作り直すことになる"],
    ["専用のPM管理ツールは", "案件管理単体では優れているが、議事録とナレッジが別の場所に分かれる。課題②がそのまま残る"],
  ];
  const ncw = 5.85, ngap = 0.39;
  notes.forEach(([t, d], i) => {
    const x = ML + i * (ncw + ngap);
    card(s, x, 4.8, ncw, 1.4, SOFT);
    s.addText(t, { x: x + 0.3, y: 4.94, w: ncw - 0.6, h: 0.32, fontFace: F, fontSize: 12.5,
      bold: true, color: NG, margin: 0 });
    s.addText(d, { x: x + 0.3, y: 5.3, w: ncw - 0.6, h: 0.78, fontFace: F, fontSize: 11,
      color: MUT, lineSpacing: 16, margin: 0 });
  });
  s.addText("※ チャットの可視化は、Teamsのパブリックチーム／チャネルで対応します（Notionは不要）。Microsoft 365 は継続利用し、置き換えません。",
    { x: ML, y: 6.36, w: CW, h: 0.32, fontFace: F, fontSize: 10.5, color: MUT, margin: 0 });
  s.addNotes("「Notionが全部で一番」とは言わない。個別項目では他ツールが勝ることを正直に認めた上で、『①と②を1つで解く』という要件で選んでいると説明する。この誠実さが説得力になる。");
}

// ============================================================
// 9. Notionでは議事録を書くことが案件管理になる
// ============================================================
{
  const s = light("Notionでは、議事録を書くことがそのまま案件管理になる",
    "第2章｜Notionを選ぶ理由", 25);

  label(s, ML, 1.6, 5.5, "議事録の正体が違う", MUT);
  const cols = [
    { t: "Word / OneNote / Loop", sub: "議事録＝文書",
      body: "決定事項もToDoも、文中に埋まった「文字」でしかない",
      res: "転記しなければ管理に使えない", fill: "F4F5F7", tc: TXT, sc: MUT, bc: MUT },
    { t: "Notion", sub: "議事録＝データベースのレコード",
      body: "ToDoは、案件に紐づいたタスクそのもの",
      res: "書いた時点で、管理の入力が完了している", fill: INK, tc: "FFFFFF", sc: ACC, bc: "C3CEE0" },
  ];
  cols.forEach((c, i) => {
    const y = 1.92 + i * 2.2;
    s.addShape(pres.ShapeType.roundRect, { x: ML, y, w: 5.5, h: 2.02, rectRadius: 0.06,
      fill: { color: c.fill }, shadow: sh() });
    s.addText(c.t, { x: ML + 0.34, y: y + 0.16, w: 4.9, h: 0.38, fontFace: F,
      fontSize: 17, bold: true, color: c.tc, margin: 0 });
    s.addText(c.sub, { x: ML + 0.34, y: y + 0.56, w: 4.9, h: 0.28, fontFace: F,
      fontSize: 11, bold: true, color: c.sc, margin: 0 });
    s.addText(c.body, { x: ML + 0.34, y: y + 0.9, w: 4.9, h: 0.5, fontFace: F,
      fontSize: 11.5, color: c.bc, lineSpacing: 17, margin: 0 });
    s.addText(c.res, { x: ML + 0.34, y: y + 1.44, w: 4.9, h: 0.42, fontFace: F,
      fontSize: 13.5, bold: true, color: c.tc, margin: 0 });
  });

  label(s, ML + 5.9, 1.6, 6.2, "会議から案件管理までの流れ", MUT);
  const flow = [
    ["会議前", "アジェンダを事前に作る", B_COL],
    ["会議中・会議後", "議事録に、決定事項とToDoを項目として登録する", B_COL],
  ];
  let fy = 1.92;
  flow.forEach(([lab, t, c]) => {
    card(s, ML + 5.9, fy, 6.2, 0.86, SOFT);
    s.addText(lab, { x: ML + 6.2, y: fy + 0.08, w: 5.6, h: 0.26, fontFace: F,
      fontSize: 10, bold: true, color: c, margin: 0 });
    s.addText(t, { x: ML + 6.2, y: fy + 0.34, w: 5.6, h: 0.42, fontFace: F, fontSize: 12.5,
      bold: true, color: INK, valign: "middle", margin: 0 });
    fy += 1.0;
  });
  s.addText("↓　転記せずに、同じ実体が3か所に現れる", { x: ML + 5.9, y: 3.94, w: 6.2, h: 0.3,
    fontFace: F, fontSize: 11, bold: true, color: ACC, align: "center", margin: 0 });
  ["① 案件ページ", "② 担当者のToDo一覧", "③ PMOの全案件ビュー"].forEach((t, i) => {
    s.addShape(pres.ShapeType.roundRect, { x: ML + 5.9, y: 4.3 + i * 0.6, w: 6.2, h: 0.5,
      rectRadius: 0.05, fill: { color: A_COL } });
    s.addText(t, { x: ML + 6.2, y: 4.3 + i * 0.6, w: 5.6, h: 0.5, fontFace: F,
      fontSize: 12, bold: true, color: "FFFFFF", valign: "middle", margin: 0 });
  });

  card(s, ML, 6.36, CW, 0.72, INK);
  s.addText("会議で決まったToDoは、転記せずにそのまま案件のタスクになります。二重入力は発生しません。",
    { x: ML + 0.44, y: 6.36, w: CW - 0.88, h: 0.72, fontFace: F, fontSize: 13.5, bold: true,
      color: "FFFFFF", valign: "middle", margin: 0 });
}

// ============================================================
// 10. 実際の画面
// ============================================================
{
  const s = light("実際の画面 ― 会議から案件管理までが1か所で繋がる",
    "第2章｜Notionを選ぶ理由", 25);
  const caps = [
    ["1", "会議ページ", "アジェンダ → 議事録 → 決定事項"],
    ["2", "生成されたタスク", "決定事項がそのままタスクに"],
    ["3", "案件ページ", "タスクが案件に紐づいた状態"],
    ["4", "PMOの横断ビュー", "全案件の進捗とガント"],
  ];
  const cw = 2.94, gap = 0.29;
  caps.forEach(([n, t, d], i) => {
    const x = ML + i * (cw + gap);
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.7, w: cw, h: 2.7, rectRadius: 0.06,
      fill: { color: SOFT }, line: { color: LINE, width: 1, dashType: "dash" } });
    s.addText("スクリーンショット\n（差し込み）", { x: x + 0.14, y: 1.7, w: cw - 0.28, h: 2.7,
      fontFace: F, fontSize: 11, color: "9AA5B4", align: "center", valign: "middle",
      lineSpacing: 18, margin: 0 });
    badge(s, x, 4.56, 0.44, n, INK, 13);
    s.addText(t, { x: x + 0.58, y: 4.54, w: cw - 0.58, h: 0.36, fontFace: F, fontSize: 13.5,
      bold: true, color: INK, valign: "middle", margin: 0 });
    s.addText(d, { x, y: 5.08, w: cw, h: 0.6, fontFace: F, fontSize: 11, color: MUT,
      lineSpacing: 17, margin: 0 });
  });
  card(s, ML, 5.86, CW, 0.86, "FBF3E4");
  s.addText("要準備：架空の研修案件で、アジェンダ → 議事録 → タスク → ガントが繋がった状態を作っておくこと",
    { x: ML + 0.44, y: 5.86, w: CW - 0.88, h: 0.86, fontFace: F, fontSize: 12.5, bold: true,
      color: "7A5410", valign: "middle", margin: 0 });
  s.addNotes("このスライドの準備の有無が提案の成否を分ける。文字10枚より画面4枚のほうが伝わる。");
}

// ============================================================
// 11. 結び
// ============================================================
{
  const s = light("だから、Notionを全社に導入したい", "まとめ");

  const sums = [
    ["①", "案件管理", "1案件のべ3〜5時間が確認と手戻りに消え、年約500万円。加えて納期遅延は顧客の信用に響く", A_COL],
    ["②", "会議と議事録", "目的が曖昧な会議に、年間約4,700万円が費やされている", B_COL],
  ];
  let y = 1.66;
  sums.forEach(([k, t, d, c]) => {
    card(s, ML, y, CW, 1.06, SOFT);
    badge(s, ML + 0.34, y + 0.28, 0.5, k, c, 15);
    s.addText(t, { x: ML + 1.0, y: y + 0.26, w: 2.4, h: 0.5, fontFace: F, fontSize: 15,
      bold: true, color: INK, valign: "middle", margin: 0 });
    s.addText(d, { x: ML + 3.5, y: y + 0.26, w: 8.4, h: 0.5, fontFace: F, fontSize: 12,
      color: MUT, valign: "middle", margin: 0 });
    y += 1.18;
  });

  card(s, ML, 4.06, CW, 1.06, INK);
  s.addText("Notionなら、議事録を書くことがそのまま案件管理になり、2つの課題を1つの仕組みで解決できます。",
    { x: ML + 0.44, y: 4.06, w: CW - 0.88, h: 1.06, fontFace: F, fontSize: 15, bold: true,
      color: "FFFFFF", valign: "middle", margin: 0 });

  s.addText("承認いただけましたら、ただちに着手すること", { x: ML, y: 5.36, w: CW, h: 0.3,
    fontFace: F, fontSize: 12, bold: true, color: ACC, margin: 0 });
  ["情報システム部門へのセキュリティ要件のヒアリング",
   "ホールディングスへの導入申請書類の作成",
   "テンプレート設計とPMO運用ルールの策定"].forEach((t, i) => {
    bullet(s, ML, 5.72 + i * 0.4, 11.5, t, MUT, 12);
  });
}

// ============================================================
// Appendix 扉
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: INK };
  s.addShape(pres.ShapeType.ellipse, { x: 11.0, y: 4.4, w: 3.6, h: 3.6,
    fill: { color: ACC, transparency: 68 } });
  s.addText("Appendix", { x: ML, y: 2.62, w: 9.2, h: 0.72, fontFace: F, fontSize: 33,
    bold: true, color: "FFFFFF", margin: 0, valign: "middle" });
  s.addText("費用・プラン比較・進め方・セキュリティ", { x: ML, y: 3.42, w: 9.2, h: 0.4,
    fontFace: F, fontSize: 14, color: "9FB0CC", margin: 0 });
}

// ============================================================
// A1. 費用
// ============================================================
{
  const s = light("年間529.2万円（税別）で、全社140名に導入できる", "Appendix｜費用");
  card(s, ML, 1.74, 6.4, 2.6, INK);
  s.addText("年間ライセンス費用（税別）", { x: ML + 0.44, y: 1.96, w: 5.5, h: 0.3,
    fontFace: F, fontSize: 12, color: "9FB0CC", margin: 0 });
  s.addText("529.2万円", { x: ML + 0.44, y: 2.34, w: 5.5, h: 0.82, fontFace: F,
    fontSize: 40, bold: true, color: "FFFFFF", margin: 0 });
  s.addText("Business 年払い　3,150円 × 140名 × 12か月", { x: ML + 0.44, y: 3.24,
    w: 5.5, h: 0.32, fontFace: F, fontSize: 12, color: "C3CEE0", margin: 0 });
  s.addText("消費税10%を加えた場合　582.1万円", { x: ML + 0.44, y: 3.62,
    w: 5.5, h: 0.32, fontFace: F, fontSize: 12, bold: true, color: ACC, margin: 0 });

  card(s, ML + 6.75, 1.74, 5.6, 2.6, SOFT);
  s.addText("実施体制（現金支出なし）", { x: ML + 7.09, y: 1.96, w: 4.9, h: 0.32,
    fontFace: F, fontSize: 12.5, bold: true, color: INK, margin: 0 });
  ["初期構築（テンプレート設計・データベース構築）",
   "教育・定着支援（説明会・マニュアル・3か月の伴走）"].forEach((t, i) => {
    bullet(s, ML + 7.09, 2.44 + i * 0.44, 4.9, t, A_COL, 11.5);
  });
  s.addText("いずれも人事組織とPMOが内製で担当します。\n新規採用は行いません。",
    { x: ML + 7.09, y: 3.42, w: 4.9, h: 0.62, fontFace: F, fontSize: 12, bold: true,
      color: INK, lineSpacing: 19, margin: 0 });

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
  s.addNotes("税区分は公式ページに明記がないため税別として扱っている。契約時に必ず確認すること。承認額が約53万円変わる。");
}

// ============================================================
// A2. プラン別 機能比較
// ============================================================
{
  const s = light("要件を満たす最小構成は、Businessプランです", "Appendix｜プラン選定");
  const head = ["", "Plus", "Business（推奨）", "Enterprise"];
  const rows = [
    ["月額（1人・年払い／税別）", "1,650円", "3,150円", "要見積（約5,425円）", "m"],
    ["年間費用（140名／税別）", "277.2万円", "529.2万円", "約911.4万円", "m"],
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
        const isRec = i === 2, eOnly = kind === "e" && i === 3;
        return { text: t, options: {
          color: eOnly ? ACC : (isRec ? A_COL : MUT), bold: isRec || eOnly,
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
  s.addNotes("狙いは「一番高いプランを買おうとしているわけではない」と示すこと。情シスが監査ログや国内保管を必須とした場合は、この表がそのままEnterprise切替の根拠資料になる。");
}

// ============================================================
// A3. 導入案の比較
// ============================================================
{
  const s = light("全社導入以外の2案は、目的を達成できない", "Appendix｜導入案の比較");
  const head = ["", "案1：Plus 全社", "案2：Business 全社", "案3：Business 上層部のみ\n＋一般社員はゲスト"];
  const rows = [
    ["年間費用（140名・年払い／税別）", "277.2万円", "529.2万円", "見かけ上は最小"],
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
    ["案3が成立しない理由", "ゲストは招待されたページしか見られず、横断ビューを持てない。課題①の可視性と横断性が解決せず、導入目的そのものが失われる"],
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
}

// ============================================================
// A4. 想定効果
// ============================================================
{
  const s = light("会議時間の削減だけで、投資額の約2.7倍", "Appendix｜想定効果");
  const stats = [
    ["投資額", "529.2万円", "年間ライセンス費用（税別）", INK],
    ["効果（会議のみ）", "約1,420万円", "目的が曖昧な会議の3割を削減", B_COL],
    ["倍率", "約2.7倍", "会議削減効果 ÷ 投資額", ACC],
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
    ["①", "案件管理の手戻り・遅延の削減（年約500万円）", A_COL],
    ["①", "PMOが全案件のリスクを事前に検知できる", A_COL],
    ["②", "属人化の解消。異動・退職時の引き継ぎコスト低減", B_COL],
    ["①", "顧客への納期遵守率向上による信用の維持", A_COL],
  ];
  let y = 4.36;
  others.forEach(([k, t, c]) => {
    s.addShape(pres.ShapeType.roundRect, { x: ML, y, w: CW, h: 0.48, rectRadius: 0.04,
      fill: { color: "FFFFFF" }, line: { color: LINE, width: 1 } });
    badge(s, ML + 0.16, y + 0.09, 0.3, k, c, 10);
    s.addText(t, { x: ML + 0.62, y, w: 11.2, h: 0.48, fontFace: F, fontSize: 12,
      color: TXT, valign: "middle", margin: 0 });
    y += 0.56;
  });
  s.addText("※ 削減率3割は意図的に控えめな設定です", { x: ML, y: 6.66, w: CW, h: 0.3,
    fontFace: F, fontSize: 10, color: MUT, margin: 0 });
}

// ============================================================
// A5. 進め方
// ============================================================
{
  const s = light("ライセンスは全社、業務の移行は段階的に行う", "Appendix｜進め方");
  const ph = [
    ["Phase 0", "承認後すぐ", "情報システム部門への申請\nセキュリティ審査", "―", MUT],
    ["Phase 1", "審査通過後 1〜2か月", "会議運用から開始\nアジェンダ・議事録の標準化", "課題②", B_COL],
    ["Phase 2", "3〜4か月目", "案件管理を1部門で先行運用\nテンプレートを実案件で検証", "課題①", A_COL],
    ["Phase 3", "5か月目以降", "全案件へ展開\nPMOの横断ビュー稼働", "課題①", A_COL],
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
// A6. 今回のスコープ
// ============================================================
{
  const s = light("今回のスコープ ― やること／やらないこと", "Appendix｜対象範囲");
  s.addText("今回ご承認いただきたい範囲", { x: ML, y: 1.76, w: 5.4, h: 0.3, fontFace: F,
    fontSize: 11.5, bold: true, color: ACC, margin: 0 });
  [["①", "案件管理", A_COL], ["②", "議事録・会議運営", B_COL]].forEach(([n, t, c], i) => {
    const y = 2.12 + i * 0.98;
    s.addShape(pres.ShapeType.roundRect, { x: ML, y, w: 4.9, h: 0.86, rectRadius: 0.06,
      fill: { color: c }, shadow: sh() });
    s.addText(n + "　" + t, { x: ML + 0.34, y, w: 4.4, h: 0.86, fontFace: F,
      fontSize: 17, bold: true, color: "FFFFFF", valign: "middle", margin: 0 });
  });
  s.addText("やること", { x: ML, y: 4.1, w: 4.9, h: 0.3, fontFace: F, fontSize: 11,
    bold: true, color: INK, align: "center", margin: 0 });

  s.addText("今回は対象に含めないもの", { x: ML + 5.6, y: 1.76, w: 6.75, h: 0.3, fontFace: F,
    fontSize: 11.5, bold: true, color: MUT, margin: 0 });
  s.addShape(pres.ShapeType.roundRect, { x: ML + 5.6, y: 2.12, w: 6.75, h: 1.84,
    rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { color: LINE, width: 1.5, dashType: "dash" } });
  s.addText("③　研修資料・ナレッジの保管", { x: ML + 5.94, y: 2.36, w: 6.1, h: 0.42, fontFace: F,
    fontSize: 17, bold: true, color: MUT, margin: 0 });
  s.addText("SharePointでの管理方法（ドキュメントセット化・タグ運用）を\n別途検討中です。本提案とは切り離して進めます。",
    { x: ML + 5.94, y: 2.86, w: 6.1, h: 0.84, fontFace: F, fontSize: 12, color: MUT,
      lineSpacing: 20, margin: 0 });
  s.addText("やらないこと", { x: ML + 5.6, y: 4.1, w: 6.75, h: 0.3, fontFace: F,
    fontSize: 11, bold: true, color: MUT, align: "center", margin: 0 });

  s.addText("なぜ範囲を絞るのか", { x: ML, y: 4.62, w: CW, h: 0.3, fontFace: F, fontSize: 12.5,
    bold: true, color: INK, margin: 0 });
  const why = [
    ["撤退可能性の確保", "思うような成果が得られなかった場合に、やめられる状態を保つ"],
    ["進行中の検討と干渉しない", "資料保管のルールは別軸で検討中。今回の判断がそちらを縛らないようにする"],
    ["効果検証の明確化", "範囲が狭いほど、効果が出たかどうかを判定しやすい"],
  ];
  const cw = 3.86, gap = 0.375;
  why.forEach(([t, d], i) => {
    const x = ML + i * (cw + gap);
    card(s, x, 4.98, cw, 1.38, SOFT);
    s.addText(t, { x: x + 0.28, y: 5.12, w: cw - 0.56, h: 0.36, fontFace: F, fontSize: 13,
      bold: true, color: INK, margin: 0 });
    s.addText(d, { x: x + 0.28, y: 5.5, w: cw - 0.56, h: 0.74, fontFace: F, fontSize: 11,
      color: MUT, lineSpacing: 17, margin: 0 });
  });
  s.addNotes("資料保管のルールはSharePointのドキュメントセット化・タグ運用として別途検討が進行中であり、Notionへ移すかどうかは決まっていない。この提案がそちらの選択肢を狭めないことを明示するため、ここではNotionという語を使わない。");
}

// ============================================================
// A7. 成功指標と撤退基準
// ============================================================
{
  const s = light("効果が出なければ、1年で見直す", "Appendix｜成功指標と撤退基準");
  s.addTable([
    [{ text: "", options: { fill: { color: INK } } },
     { text: "指標", options: { bold: true, color: "FFFFFF", fill: { color: INK } } },
     { text: "現状", options: { bold: true, color: "FFFFFF", fill: { color: INK } } },
     { text: "目標", options: { bold: true, color: "FFFFFF", fill: { color: INK } } }],
    ...[["①", "案件の期日遵守率", "未計測", "計測開始 → ◯%"],
        ["①", "案件情報がNotionに登録されている割合", "0%", "95% 以上"],
        ["②", "1人あたり週の会議時間", "概算のみ", "▲◯%"],
        ["②", "議事録が残っている会議の割合", "未計測", "90% 以上"]]
      .map(([k, a, b, c]) => ([
        { text: k, options: { bold: true, color: "FFFFFF", align: "center",
          fill: { color: k === "①" ? A_COL : B_COL } } },
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
}

// ============================================================
// A8. リスクと対策
// ============================================================
{
  const s = light("想定されるリスクは3つ。すべて対策がある", "Appendix｜リスクと対策");
  const risks = [
    ["定着しない", "一部の人しか使わず、結局チャットに戻る",
     ["経営層で3か月試行済み", "テンプレート・ミニマニュアルを整備", "PMOが運用ルールの管理責任を持つ", "導入後3か月は人事組織が伴走"]],
    ["情報の置き場所が分散する", "Teams・Notion・SharePointのどこに何があるか分からなくなる",
     ["役割分担を1枚の図で定義し全社に周知", "「会話はTeams、記録はNotion」を徹底", "SharePointは当面そのまま併存"]],
    ["セキュリティ", "親会社の基準を満たせない",
     ["SOC 2 Type II・ISO 27001 等を取得済み", "SAML SSOでMicrosoft 365と統合", "機微情報は置かない運用とする", "詳細は次ページ"]],
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
    ms.forEach((m, j) => bullet(s, x + 0.3, 3.6 + j * 0.6, cw - 0.6, m, OK, 11, 0.56));
  });
}

// ============================================================
// A9. セキュリティ
// ============================================================
{
  const s = light("SOC 2 Type II・ISO 27001を取得済み", "Appendix｜セキュリティ");

  card(s, ML, 1.7, 5.85, 2.5, SOFT);
  s.addText("第三者認証（取得済み）", { x: ML + 0.34, y: 1.9, w: 5.2, h: 0.32, fontFace: F,
    fontSize: 12.5, bold: true, color: OK, margin: 0 });
  ["SOC 2 Type II（年次の第三者監査）",
   "ISO/IEC 27001・27017・27018・27701",
   "年次ペネトレーションテスト＋バグバウンティ",
   "監査報告書はトラストセンターで入手可能"].forEach((t, i) => {
    bullet(s, ML + 0.34, 2.3 + i * 0.44, 5.2, t, OK, 11.5);
  });

  card(s, ML + 6.25, 1.7, 5.85, 2.5, SOFT);
  s.addText("Businessプランでできないこと", { x: ML + 6.59, y: 1.9, w: 5.2, h: 0.32,
    fontFace: F, fontSize: 12.5, bold: true, color: NG, margin: 0 });
  ["監査ログ（情報持ち出しの追跡はできない）",
   "SCIM（退職者のアカウント自動削除）",
   "データの国内保管（Enterpriseのみ）"].forEach((t, i) => {
    bullet(s, ML + 6.59, 2.3 + i * 0.44, 5.2, t, NG, 11.5);
  });
  s.addText("いずれかが必須要件とされた場合は Enterprise への切替となります。",
    { x: ML + 6.59, y: 3.68, w: 5.2, h: 0.32, fontFace: F, fontSize: 10.5, color: MUT, margin: 0 });

  s.addText("運用面での安全設計（こちらから制約を課す）", { x: ML, y: 4.4, w: CW, h: 0.32,
    fontFace: F, fontSize: 13, bold: true, color: INK, margin: 0 });
  const ops = [
    ["外部共有の原則禁止", "公開リンクの生成を管理者権限で禁止。必要時のみ申請制"],
    ["扱わない情報の明示", "個人情報・機微情報はNotionに保存しないルールとする"],
    ["退職・異動時の手順", "SAML SSOでM365のアカウント停止と同時にログイン不能。メンバー削除は月次で突合"],
  ];
  let y = 4.8;
  ops.forEach(([t, d]) => {
    s.addShape(pres.ShapeType.roundRect, { x: ML, y, w: CW, h: 0.56, rectRadius: 0.04,
      fill: { color: "FFFFFF" }, line: { color: LINE, width: 1 } });
    s.addText(t, { x: ML + 0.3, y, w: 3.4, h: 0.56, fontFace: F, fontSize: 12, bold: true,
      color: INK, valign: "middle", margin: 0 });
    s.addText(d, { x: ML + 3.8, y, w: 8.0, h: 0.56, fontFace: F, fontSize: 11,
      color: MUT, valign: "middle", margin: 0 });
    y += 0.64;
  });
  s.addNotes("情シスへの事前相談が最優先。監査ログと国内データ保管の要否を先に確認すること。どちらかが必須ならEnterprise（約911万円）になり、承認額が変わる。");
}

// ============================================================
// A10. 承認いただきたい3点
// ============================================================
{
  const s = light("承認いただきたい3点", "Appendix｜お願い");
  const items = [
    ["1", "全社（140名）へのNotion導入", "案件管理と会議運営の標準化を、PMO発足と同時に開始する"],
    ["2", "年間費用 529.2万円（税別）", "Businessプラン・年払い"],
    ["3", "ホールディングスへの申請着手", "セキュリティ審査の事前相談を開始する許可"],
  ];
  let y = 1.8;
  items.forEach(([n, t, d]) => {
    card(s, ML, y, CW, 1.32, SOFT);
    badge(s, ML + 0.42, y + 0.36, 0.6, n, INK, 17);
    s.addText(t, { x: ML + 1.28, y: y + 0.26, w: 8.6, h: 0.42, fontFace: F, fontSize: 18,
      bold: true, color: INK, valign: "middle", margin: 0 });
    s.addText(d, { x: ML + 1.28, y: y + 0.72, w: 10.2, h: 0.36, fontFace: F, fontSize: 12.5,
      color: MUT, valign: "middle", margin: 0 });
    y += 1.5;
  });
  s.addText("稼働目標：PMO発足と同時", { x: ML, y: 6.5, w: CW, h: 0.34, fontFace: F,
    fontSize: 13, bold: true, color: ACC, margin: 0 });
}

pres.writeFile({ fileName: "/home/user/claude-test/slides/notion-proposal.pptx" })
  .then(f => console.log("written:", f));
