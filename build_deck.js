/**
 * Notion導入 THD説明資料 — スライド生成スクリプト
 * 骨組み: 本編15枚 + Appendix10枚 (+ 表紙・目次・Appendix扉) = 計28枚
 */
const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.author = "PCA";
pres.title = "Notion導入のご説明";

/* ── パレット（Midnight Executive 基調） ───────────────── */
const NAVY = "1E2761";
const NAVY_DEEP = "141B47";
const NAVY_MID = "3A4A8F";
const ICE = "CADCFC";
const ICE_PALE = "EDF3FD";
const WHITE = "FFFFFF";
const INK = "1A1A2E";
const MUTED = "6B7280";
const LINE = "D6DEEC";
const RED = "B3261E";
const RED_PALE = "FBEDEC";
const AMBER = "9A6700";
const GREEN = "1E7A5A";
const GREEN_PALE = "E9F5F0";

const F = "Meiryo";
const W = 13.333, H = 7.5, M = 0.6, CW = W - M * 2;

let pageNo = 0;
const sh = () => ({ type: "outer", color: "1E2761", blur: 8, offset: 1, angle: 90, opacity: 0.10 });

/* ── 共通パーツ ───────────────────────────────────────── */
function footer(slide) {
  pageNo++;
  slide.addText(String(pageNo), {
    x: W - 1.0, y: H - 0.48, w: 0.5, h: 0.3, align: "right",
    fontFace: F, fontSize: 10, color: MUTED,
  });
}

/** 標準コンテンツスライド: 章チップ + タイトル + キーメッセージカード */
function content(chip, title, keyMsg, chipColor) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 0.40, w: 1.35, h: 0.34, rectRadius: 0.06,
    fill: { color: chipColor || NAVY }, line: { color: chipColor || NAVY },
  });
  s.addText(chip, {
    x: M, y: 0.40, w: 1.35, h: 0.34, align: "center", valign: "middle",
    fontFace: F, fontSize: 11, bold: true, color: WHITE, margin: 0,
  });
  s.addText(title, {
    x: M, y: 0.84, w: CW, h: 0.62, valign: "middle",
    fontFace: F, fontSize: 27, bold: true, color: NAVY, margin: 0,
  });
  if (keyMsg) {
    s.addShape(pres.ShapeType.roundRect, {
      x: M, y: 1.56, w: CW, h: 0.72, rectRadius: 0.08,
      fill: { color: ICE_PALE }, line: { color: ICE },
    });
    s.addText(keyMsg, {
      x: M + 0.22, y: 1.56, w: CW - 0.44, h: 0.72, valign: "middle",
      fontFace: F, fontSize: 14, bold: true, color: NAVY, margin: 0,
    });
  }
  footer(s);
  return s;
}

/** カード（角丸 + 淡色塗り） */
function card(s, x, y, w, h, fill, lineC) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: fill }, line: { color: lineC || LINE }, shadow: sh(),
  });
}

/** 判定マーク付きセル */
function mk(mark, note) {
  const c = mark === "○" || mark === "◎" ? GREEN : mark === "△" ? AMBER : mark === "×" ? RED : INK;
  const rt = [{ text: mark, options: { color: c, bold: true, fontSize: 12 } }];
  if (note) rt.push({ text: "\n" + note, options: { color: INK, fontSize: 9 } });
  return { text: rt };
}
const tx = (t, o) => ({ text: t, options: Object.assign({ color: INK, fontSize: 10 }, o || {}) });

/** 表 */
function table(s, rows, y, colW, opts) {
  const o = Object.assign({
    x: M, y, w: CW, colW,
    border: { type: "solid", pt: 0.5, color: LINE },
    fontFace: F, fontSize: 10, color: INK, valign: "middle",
    align: "left", autoPage: false,
  }, opts || {});
  s.addTable(rows, o);
}
function hdr(t, o) {
  return Object.assign({
    text: t,
    options: Object.assign({ fill: { color: NAVY }, color: WHITE, bold: true, fontSize: 10.5, align: "center", valign: "middle" }, o || {}),
  });
}

/* ═══════════════ 1. 表紙 ═══════════════ */
{
  const s = pres.addSlide();
  s.background = { color: NAVY_DEEP };
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.9, y: 1.55, w: 2.7, h: 0.42, rectRadius: 0.07,
    fill: { color: NAVY_MID }, line: { color: NAVY_MID },
  });
  s.addText("THD ご説明資料", {
    x: 0.9, y: 1.55, w: 2.7, h: 0.42, align: "center", valign: "middle",
    fontFace: F, fontSize: 12, bold: true, color: ICE, margin: 0,
  });
  s.addText("実行管理の仕組みの新規立ち上げ", {
    x: 0.9, y: 2.25, w: 11.5, h: 0.85,
    fontFace: F, fontSize: 38, bold: true, color: WHITE, margin: 0,
  });
  s.addText("案件管理と議事録を1か所に — Notion の正式導入について", {
    x: 0.9, y: 3.15, w: 11.5, h: 0.5,
    fontFace: F, fontSize: 17, color: ICE, margin: 0,
  });
  s.addShape(pres.ShapeType.rect, { x: 0.92, y: 3.95, w: 3.4, h: 0.02, fill: { color: NAVY_MID }, line: { color: NAVY_MID } });
  s.addText(
    [
      { text: "対象範囲　①案件管理　②議事録", options: { breakLine: true, fontSize: 13, color: WHITE, bold: true } },
      { text: "対象人数　140名（Phase2 着地見込）", options: { breakLine: true, fontSize: 13, color: WHITE, bold: true } },
      { text: "契約プラン　Notion Business（SAML SSO 対応）", options: { fontSize: 13, color: WHITE, bold: true } },
    ],
    { x: 0.9, y: 4.25, w: 7.5, h: 1.3, fontFace: F, margin: 0, lineSpacing: 24 }
  );
  s.addText("株式会社PCA　2026年8月", {
    x: 0.9, y: 6.5, w: 6, h: 0.35, fontFace: F, fontSize: 12, color: MUTED, margin: 0,
  });
  s.addNotes("THDへの説明は仲川。実物デモを併用する。本日ご承認いただきたいのは、①②に限定した正式導入・Businessプラン契約・コスト上限の3点。");
}

/* ═══════════════ 2. 目次 ═══════════════ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText("本日のご説明", {
    x: M, y: 0.55, w: CW, h: 0.7, valign: "middle",
    fontFace: F, fontSize: 30, bold: true, color: NAVY, margin: 0,
  });
  s.addText("背景 → 課題 → 要件 → 選定 → 実証 → セキュリティ → 計画 → お願い の順にご説明します", {
    x: M, y: 1.25, w: CW, h: 0.4, fontFace: F, fontSize: 13, color: MUTED, margin: 0,
  });

  const items = [
    ["序", "30秒サマリー", "本日のご依頼事項"],
    ["1", "背景", "PCAは再建の実行フェーズにある"],
    ["2", "課題", "実行管理の仕組みが存在しない"],
    ["3", "要件", "必要なのは「つながる」管理基盤"],
    ["4", "選定", "4案を比較し、Notionのみが要件を充足"],
    ["5", "実証", "経営層で3か月の実運用、効果を確認済み"],
    ["6", "セキュリティ", "認証・SSO・プラン要件（THD審査の要点）"],
    ["7", "導入計画", "目指す運用の姿と、コスト・撤退基準"],
    ["8", "お願い事項", "ご承認いただきたいこと／確認事項"],
  ];
  const cw = 5.95, ch = 0.52, gap = 0.13;
  items.forEach((it, i) => {
    const col = i < 5 ? 0 : 1;
    const row = i < 5 ? i : i - 5;
    const x = M + col * (cw + 0.24);
    const y = 1.85 + row * (ch + gap);
    card(s, x, y, cw, ch, i === 5 ? ICE_PALE : WHITE, i === 5 ? ICE : LINE);
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.13, y: y + 0.1, w: 0.42, h: 0.32, rectRadius: 0.05,
      fill: { color: NAVY }, line: { color: NAVY },
    });
    s.addText(it[0], {
      x: x + 0.13, y: y + 0.1, w: 0.42, h: 0.32, align: "center", valign: "middle",
      fontFace: F, fontSize: 10, bold: true, color: WHITE, margin: 0,
    });
    s.addText(
      [
        { text: it[1] + "　", options: { bold: true, color: NAVY, fontSize: 12 } },
        { text: it[2], options: { color: MUTED, fontSize: 10 } },
      ],
      { x: x + 0.68, y, w: cw - 0.8, h: ch, valign: "middle", fontFace: F, margin: 0 }
    );
  });

  card(s, M, 5.45, CW, 1.35, ICE_PALE, ICE);
  s.addText("本資料は2つの話を分けてご説明します", {
    x: M + 0.25, y: 5.58, w: CW - 0.5, h: 0.32, fontFace: F, fontSize: 12.5, bold: true, color: NAVY, margin: 0,
  });
  s.addText(
    [
      { text: "① 仕組みを作る（序〜第6章）", options: { bold: true, color: NAVY, fontSize: 11, breakLine: true } },
      { text: "案件管理と議事録が紐づく器を作る。ご承認をお願いする対象はこちらです。", options: { color: INK, fontSize: 10 } },
    ],
    { x: M + 0.25, y: 5.95, w: 5.7, h: 0.72, fontFace: F, margin: 0 }
  );
  s.addText(
    [
      { text: "② 会議の在り方を変える（第7章 7-1）", options: { bold: true, color: NAVY, fontSize: 11, breakLine: true } },
      { text: "その仕組みができて初めて可能になる運用の話。承認対象ではありません。", options: { color: INK, fontSize: 10 } },
    ],
    { x: M + 6.35, y: 5.95, w: 5.7, h: 0.72, fontFace: F, margin: 0 }
  );
  footer(s);
}

/* ═══════════════ 3. 序 30秒サマリー ═══════════════ */
{
  const s = content("序", "30秒サマリー", "決めたことをやり切る「実行管理の仕組み」を、PCAに新しく立ち上げます");
  const items = [
    ["1", "課題", "PCAには案件を管理する仕組みが無い", "議事録も OneNote・Word・個人メモに散在し、探すところから仕事が始まっている"],
    ["2", "対応", "①案件管理 ②議事録 に限定して新規立ち上げ", "③ファイル・④社内規定は SharePoint 継続。M365 は置き換えません"],
    ["3", "根拠", "経営層で3か月の実運用、効果は確認済み", "机上の選定ではなく、週次PMOの管理台帳として既に定着しています"],
    ["4", "安全性", "Business プランの SAML SSO で M365 と統合", "当社の既存セキュリティポリシーがそのまま Notion にも適用されます"],
  ];
  const rh = 0.72, gap = 0.16;
  items.forEach((it, i) => {
    const y = 2.5 + i * (rh + gap);
    card(s, M, y, CW - 3.5, rh, WHITE, LINE);
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 0.16, y: y + 0.16, w: 0.4, h: 0.4, rectRadius: 0.2,
      fill: { color: NAVY }, line: { color: NAVY },
    });
    s.addText(it[0], {
      x: M + 0.16, y: y + 0.16, w: 0.4, h: 0.4, align: "center", valign: "middle",
      fontFace: F, fontSize: 11, bold: true, color: WHITE, margin: 0,
    });
    s.addText(it[1], {
      x: M + 0.68, y: y + 0.08, w: 0.85, h: 0.28, valign: "middle",
      fontFace: F, fontSize: 10, bold: true, color: MUTED, margin: 0,
    });
    s.addText(it[2], {
      x: M + 1.55, y: y + 0.06, w: 6.7, h: 0.3, valign: "middle",
      fontFace: F, fontSize: 12, bold: true, color: NAVY, margin: 0,
    });
    s.addText(it[3], {
      x: M + 0.68, y: y + 0.37, w: 7.6, h: 0.3, valign: "middle",
      fontFace: F, fontSize: 9.5, color: INK, margin: 0,
    });
  });

  card(s, M + CW - 3.3, 2.5, 3.3, 3.4, NAVY, NAVY);
  s.addText("本日のお願い", {
    x: M + CW - 3.1, y: 2.68, w: 2.9, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: ICE, margin: 0,
  });
  s.addText("140名", {
    x: M + CW - 3.1, y: 3.05, w: 2.9, h: 0.72, fontFace: F, fontSize: 40, bold: true, color: WHITE, margin: 0,
  });
  s.addText("Phase2 対象人数（着地見込）", {
    x: M + CW - 3.1, y: 3.75, w: 2.9, h: 0.26, fontFace: F, fontSize: 9.5, color: ICE, margin: 0,
  });
  s.addText("年額 約520.8万円", {
    x: M + CW - 3.1, y: 4.2, w: 2.9, h: 0.42, fontFace: F, fontSize: 20, bold: true, color: WHITE, margin: 0,
  });
  s.addText("Business プラン・年払\nこの金額を上限としてご承認ください", {
    x: M + CW - 3.1, y: 4.65, w: 2.9, h: 0.6, fontFace: F, fontSize: 9.5, color: ICE, margin: 0,
  });
  s.addText("1人あたり 1営業日 約169円", {
    x: M + CW - 3.1, y: 5.4, w: 2.9, h: 0.3, valign: "middle",
    fontFace: F, fontSize: 10, bold: true, color: ICE, italic: true, margin: 0,
  });
  s.addNotes("冒頭30秒で全体像を示す。THDが最初に気にするセキュリティを4番目に置き、この後の第6章で詳述することを予告する。金額は上限としての提示であり、実際の調達は実人数でTrue-upする旨を口頭で補足。");
}

/* ═══════════════ 4. 第1章 背景 ═══════════════ */
{
  const s = content("第1章", "背景 — PCAは再建の実行フェーズにある", "再建の成否は、戦略の中身以上に「やり切る」実行管理で決まります");
  const blocks = [
    ["状況", NAVY, [
      "中期経営計画 FY27-29 を実行中",
      "10月1日の新体制で、部門をまたぐ変革案件が同時多発",
    ]],
    ["帰結", RED, [
      "決めたことの実行漏れが最大のリスク",
      "実行管理の巧拙が、そのまま業績に直結する",
    ]],
  ];
  blocks.forEach((b, i) => {
    const x = M + i * (CW / 2 + 0.15);
    const w = CW / 2 - 0.15;
    card(s, x, 2.55, w, 2.05, i === 1 ? RED_PALE : ICE_PALE, i === 1 ? "F0CFCC" : ICE);
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.25, y: 2.75, w: 0.95, h: 0.34, rectRadius: 0.05,
      fill: { color: b[1] }, line: { color: b[1] },
    });
    s.addText(b[0], {
      x: x + 0.25, y: 2.75, w: 0.95, h: 0.34, align: "center", valign: "middle",
      fontFace: F, fontSize: 11, bold: true, color: WHITE, margin: 0,
    });
    s.addText(
      b[2].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j !== b[2].length - 1 } })),
      { x: x + 0.25, y: 3.25, w: w - 0.5, h: 1.15, fontFace: F, fontSize: 12.5, color: INK, margin: 0, paraSpaceAfter: 8 }
    );
  });

  card(s, M, 4.85, CW, 1.75, WHITE, LINE);
  s.addText("いま起きていること", {
    x: M + 0.28, y: 5.0, w: 4, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: NAVY, margin: 0,
  });
  const facts = [
    ["進捗の確認", "会議での口頭確認に頼っている"],
    ["決定事項の追跡", "担当者の記憶と個人メモに依存している"],
    ["過去の経緯", "「どこにあるか」を探すところから始まる"],
  ];
  facts.forEach((f, i) => {
    const x = M + 0.28 + i * 3.9;
    s.addText(f[0], { x, y: 5.42, w: 3.6, h: 0.28, fontFace: F, fontSize: 10, bold: true, color: MUTED, margin: 0 });
    s.addText(f[1], { x, y: 5.7, w: 3.6, h: 0.6, fontFace: F, fontSize: 11.5, color: INK, margin: 0 });
  });
  s.addNotes("この章では課題に踏み込まず、再建フェーズという前提だけを共有する。実行管理の巧拙が業績に直結するという一点に絞る。");
}

/* ═══════════════ 5. 第2章 課題 ═══════════════ */
{
  const s = content("第2章", "課題 — 実行管理の仕組みが存在しない", "これはツールの移行ではなく、空白領域の新規立ち上げです");
  card(s, M, 2.5, 5.9, 1.9, WHITE, LINE);
  s.addText("現状", { x: M + 0.25, y: 2.65, w: 2, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: NAVY, margin: 0 });
  s.addText(
    [
      { text: "案件管理の仕組みが、そもそも無い", options: { bullet: true, breakLine: true } },
      { text: "議事録は OneNote・Word・個人メモに散在", options: { bullet: true } },
    ],
    { x: M + 0.25, y: 3.05, w: 5.4, h: 1.1, fontFace: F, fontSize: 12.5, color: INK, margin: 0, paraSpaceAfter: 10 }
  );

  card(s, M + 6.23, 2.5, 5.9, 1.9, WHITE, LINE);
  s.addText("影響", { x: M + 6.48, y: 2.65, w: 2, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: NAVY, margin: 0 });
  s.addText(
    [
      { text: "進捗は会議の口頭確認頼み", options: { bullet: true, breakLine: true } },
      { text: "決定事項の追跡が属人的になっている", options: { bullet: true } },
    ],
    { x: M + 6.48, y: 3.05, w: 5.4, h: 1.1, fontFace: F, fontSize: 12.5, color: INK, margin: 0, paraSpaceAfter: 10 }
  );

  card(s, M, 4.65, CW, 2.0, NAVY, NAVY);
  s.addText("探したいときに、探せない", {
    x: M + 0.4, y: 4.9, w: 6.5, h: 0.55, valign: "middle",
    fontFace: F, fontSize: 26, bold: true, color: WHITE, margin: 0,
  });
  s.addText("「あの件どうなった？」「あの資料どこ？」を探すところから仕事が始まっている。\n情報が無いのではなく、たどり着けない状態です。", {
    x: M + 0.4, y: 5.5, w: 7.2, h: 0.85, fontFace: F, fontSize: 12.5, color: ICE, margin: 0, lineSpacing: 20,
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M + 8.2, y: 4.95, w: 3.6, h: 1.4, rectRadius: 0.08,
    fill: { color: NAVY_MID }, line: { color: NAVY_MID },
  });
  s.addText("この「探す時間」が、\n再建の実行スピードを\n直接削っています", {
    x: M + 8.35, y: 4.95, w: 3.3, h: 1.4, valign: "middle", align: "center",
    fontFace: F, fontSize: 12.5, bold: true, color: WHITE, margin: 0, lineSpacing: 22,
  });
  s.addNotes("会議の質や進め方の話はここでは触れない（第7章で別軸として扱う）。この章は「器が無い」という一点に絞る。");
}

/* ═══════════════ 6. 第3章 要件 ═══════════════ */
{
  const s = content("第3章", "要件 — 必要なのは「つながる」管理基盤", "案件 ⇔ タスク ⇔ 議事録 が1か所で紐づく、データベース型の基盤が必要です");
  s.addText("機能要件", { x: M, y: 2.45, w: 3, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: MUTED, margin: 0 });
  const fn = [
    ["相互リンク", "案件を開けば、関連する議事録・決定事項・担当タスクにそのまま辿り着ける"],
    ["立場別ビュー", "同じデータを、経営はガント／PMOは週次／部門はボード・カレンダーで見る"],
    ["探したいときに探せる", "横断検索で、過去の経緯に数クリックで到達できる"],
    ["決定事項の一斉展開", "全社へ展開でき、原本を直せば展開先すべてに自動反映される"],
  ];
  fn.forEach((f, i) => {
    const x = M + (i % 2) * (CW / 2 + 0.12);
    const y = 2.8 + Math.floor(i / 2) * 0.95;
    const w = CW / 2 - 0.12;
    card(s, x, y, w, 0.82, WHITE, LINE);
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.18, y: y + 0.24, w: 0.34, h: 0.34, rectRadius: 0.17,
      fill: { color: ICE }, line: { color: ICE },
    });
    s.addText(String(i + 1), {
      x: x + 0.18, y: y + 0.24, w: 0.34, h: 0.34, align: "center", valign: "middle",
      fontFace: F, fontSize: 10, bold: true, color: NAVY, margin: 0,
    });
    s.addText(f[0], { x: x + 0.62, y: y + 0.1, w: w - 0.8, h: 0.3, valign: "middle", fontFace: F, fontSize: 12.5, bold: true, color: NAVY, margin: 0 });
    s.addText(f[1], { x: x + 0.62, y: y + 0.4, w: w - 0.8, h: 0.36, fontFace: F, fontSize: 9.5, color: INK, margin: 0 });
  });

  s.addText("運用要件", { x: M, y: 4.75, w: 3, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: MUTED, margin: 0 });
  const op = [
    ["非エンジニアの全員が使えること", "特別な教育を前提にしない"],
    ["SSO・権限による統制", "→ 第6章で詳しくご説明します"],
    ["1製品で完結すること", "教育コストと現場の混乱を最小化する"],
  ];
  op.forEach((o, i) => {
    const x = M + i * (CW / 3 + 0.03);
    const w = CW / 3 - 0.06;
    const hl = i === 1;
    card(s, x, 5.1, w, 1.0, hl ? ICE_PALE : WHITE, hl ? ICE : LINE);
    s.addText(o[0], { x: x + 0.2, y: 5.24, w: w - 0.4, h: 0.44, fontFace: F, fontSize: 12, bold: true, color: NAVY, margin: 0 });
    s.addText(o[1], { x: x + 0.2, y: 5.68, w: w - 0.4, h: 0.3, fontFace: F, fontSize: 9.5, color: hl ? NAVY : MUTED, bold: hl, margin: 0 });
  });
  s.addNotes("要件を先に固定することが重要。ツールありきに見せないため、この章では製品名を出さない。SSO・権限は深追いせず第6章に送る。");
}

/* ═══════════════ 7. 第4章 4-1 比較表 ═══════════════ */
{
  const s = content("第4章", "選定 — 要件を固定したうえで、4案を比較する", "先に要件を決めてから比較すると、1製品で満たせるのは Notion だけでした");
  const rows = [
    [hdr("要件", { align: "left" }), hdr("A. M365の組合せ\n(Planner/Lists/Loop)"), hdr("B. 無料ツール"), hdr("C. 専用ツール\n(Asana/Backlog等)"), hdr("D. Notion")],
    [tx("1. 案件⇔タスク⇔議事録の相互リンク（DB型）", { bold: true }), mk("△", "ツールごとに分散しリンク不可"), mk("△", "機能不足"), mk("△", "タスク特化・議事録は別"), mk("○", "リレーションで実現")],
    [tx("2. 同一データの立場別ビュー", { bold: true }), mk("×", "不可"), mk("△", "限定的"), mk("○", "対応"), mk("○", "対応")],
    [tx("3. 横断検索・過去経緯への即到達", { bold: true }), mk("△", "ツールをまたぐ"), mk("△", "限定的"), mk("△", "議事録は範囲外"), mk("○", "1か所で完結")],
    [tx("4. 非エンジニアの全員が使える", { bold: true }), mk("△", "複数ツールの習熟が必要"), mk("○", "簡易"), mk("△", "専門的"), mk("○", "ミニマニュアル前提")],
    [tx("5. ガバナンス（SSO・権限）", { bold: true }), mk("○", "M365で統制済"), mk("×", "有料機能・統制不可"), mk("○", "対応"), mk("○", "SSO=Business以上（第6章）")],
    [tx("6. 1製品で完結（教育コスト最小化）", { bold: true }), mk("×", "複数併用"), mk("×", "複数併用"), mk("×", "複数併用"), mk("○", "1製品")],
    [tx("（参考）社内での実証", { bold: true }), tx("—", { align: "center" }), tx("—", { align: "center" }), tx("—", { align: "center" }), mk("◎", "経営層で3か月実運用済み")],
  ];
  table(s, rows, 2.42, [3.5, 2.16, 2.16, 2.16, 2.15], { rowH: [0.42, 0.4, 0.32, 0.4, 0.4, 0.4, 0.32, 0.4], fontSize: 8.5 });

  card(s, M, 6.2, CW, 0.72, NAVY, NAVY);
  s.addText("「部門ごとの無料ツール乱立こそシャドーITリスク。有料1本に統制する方が、ガバナンスは強くなる」", {
    x: M + 0.35, y: 6.2, w: CW - 3.4, h: 0.72, valign: "middle",
    fontFace: F, fontSize: 12.5, bold: true, color: WHITE, margin: 0,
  });
  s.addText("B案（無料ツール）を選ばない理由", {
    x: M + CW - 3.15, y: 6.2, w: 2.9, h: 0.72, valign: "middle", align: "right",
    fontFace: F, fontSize: 9.5, color: ICE, margin: 0,
  });
  s.addNotes("コストをかけない＝安全ではなく、統制されていない無料利用こそが危険、という逆転の論法。コストを払う理由をガバナンス側から正当化し、第6章への橋渡しにする。");
}

/* ═══════════════ 8. 第4章 4-2 対M365 ═══════════════ */
{
  const s = content("第4章", "選定 — M365とは役割が異なり、置き換えではありません", "M365 は継続して併用します。Notion は M365 が担っていない領域を埋めます");
  card(s, M, 2.45, CW, 1.15, NAVY, NAVY);
  s.addText("「M365で解決できるなら、この数年ですでに解決していたはず。\n複数アプリの組み合わせは情報散在の再生産であり、いま解決したい課題そのものである」", {
    x: M + 0.35, y: 2.45, w: CW - 0.7, h: 1.15, valign: "middle",
    fontFace: F, fontSize: 14.5, bold: true, color: WHITE, margin: 0, lineSpacing: 26,
  });

  s.addText("役割の住み分け", { x: M, y: 3.78, w: 4, h: 0.28, fontFace: F, fontSize: 11.5, bold: true, color: MUTED, margin: 0 });
  // [名称, 役割, 箇条書き, 背景色, 見出し色, 副題色]
  const dom = [
    ["Notion", "案件と議事録", ["案件・タスク・議事録の管理", "相互リンクと横断検索", "立場別のビュー"], NAVY, WHITE, ICE],
    ["SharePoint", "ファイルの正本", ["案件関連ファイルの保管", "社内規定・マニュアル", "既存運用をそのまま継続"], WHITE, NAVY, MUTED],
    ["Teams", "伝達・コミュニケーション", ["チャット・会議", "日々のやりとり", "既存運用をそのまま継続"], WHITE, NAVY, MUTED],
  ];
  dom.forEach((d, i) => {
    const w = CW / 3 - 0.14;
    const x = M + i * (w + 0.21);
    card(s, x, 4.12, w, 1.85, d[3], i === 0 ? NAVY : LINE);
    s.addText(d[0], { x: x + 0.25, y: 4.28, w: w - 0.5, h: 0.34, fontFace: F, fontSize: 16, bold: true, color: d[4], margin: 0 });
    s.addText(d[1], { x: x + 0.25, y: 4.62, w: w - 0.5, h: 0.28, fontFace: F, fontSize: 11, bold: true, color: d[5], margin: 0 });
    s.addText(
      d[2].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j !== d[2].length - 1 } })),
      { x: x + 0.25, y: 4.98, w: w - 0.5, h: 0.85, fontFace: F, fontSize: 9.5, color: i === 0 ? ICE : INK, margin: 0, paraSpaceAfter: 4 }
    );
  });

  card(s, M, 6.15, CW, 0.72, ICE_PALE, ICE);
  s.addText("PCA側のお約束　③案件関連ファイル と ④社内規定・マニュアル は SharePoint を継続します。M365 の置き換え・二重投資にはあたりません。", {
    x: M + 0.3, y: 6.15, w: CW - 0.6, h: 0.72, valign: "middle",
    fontFace: F, fontSize: 12, bold: true, color: NAVY, margin: 0,
  });
  s.addNotes("THDが必ず問う最大の論点。「機能が足りない」ではなく「組み合わせること自体が課題の再生産」という論理で消す。Copilotについて聞かれたらAppendix Bを参照。");
}

/* ═══════════════ 9. 第5章 5-1 実運用中の姿 ═══════════════ */
{
  const s = content("第5章", "実証 — 経営層ですでに3か月、実運用しています", "机上の選定ではありません。パイロット済みのため、導入失敗のリスクは最小です");
  const feats = [
    ["経営イシューマップ", "7領域・40件超", "経営課題を7つの領域に整理し、案件単位で進捗を管理"],
    ["タスクDB", "イシューと双方向連携", "イシューを開けば担当タスクが、タスクからはイシューが辿れる"],
    ["議事録DB", "会議体別・カレンダービュー", "会議体ごとに整理し、カレンダーから議事録に直接到達できる"],
  ];
  feats.forEach((f, i) => {
    const w = 3.72;
    const x = M + i * (w + 0.24);
    card(s, x, 2.5, w, 1.72, WHITE, LINE);
    s.addText(f[0], { x: x + 0.22, y: 2.66, w: w - 0.44, h: 0.32, fontFace: F, fontSize: 13.5, bold: true, color: NAVY, margin: 0 });
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.22, y: 3.03, w: w - 0.44, h: 0.3, rectRadius: 0.05,
      fill: { color: ICE_PALE }, line: { color: ICE },
    });
    s.addText(f[1], {
      x: x + 0.22, y: 3.03, w: w - 0.44, h: 0.3, align: "center", valign: "middle",
      fontFace: F, fontSize: 10, bold: true, color: NAVY, margin: 0,
    });
    s.addText(f[2], { x: x + 0.22, y: 3.42, w: w - 0.44, h: 0.68, fontFace: F, fontSize: 10, color: INK, margin: 0 });
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 4.45, w: CW, h: 2.05, rectRadius: 0.08,
    fill: { color: "F7F9FC" }, line: { color: LINE, dashType: "dash" },
  });
  s.addText("［ 実画面キャプチャ 差し込み枠 ］", {
    x: M, y: 4.85, w: CW, h: 0.4, align: "center",
    fontFace: F, fontSize: 15, bold: true, color: MUTED, margin: 0,
  });
  s.addText("経営イシューマップ／タスクDB／議事録DB の実画面をこの枠に貼り付けてください（機微情報はマスキングのうえ）。", {
    x: M, y: 5.28, w: CW, h: 0.3, align: "center",
    fontFace: F, fontSize: 10.5, color: MUTED, margin: 0,
  });
  s.addText("THDご説明の場では、この資料に加えて実物のデモをご覧いただきます。", {
    x: M, y: 5.72, w: CW, h: 0.3, align: "center",
    fontFace: F, fontSize: 11.5, bold: true, color: NAVY, margin: 0,
  });
  s.addNotes("ここで実物デモに切り替える（仲川）。イシューマップから案件を1件開き、紐づく議事録とタスクに数クリックで到達するところを見せるのが最も効果的。");
}

/* ═══════════════ 10. 第5章 5-2 効果 ═══════════════ */
{
  const s = content("第5章", "実証 — 3か月で確認できた効果", "情報の置き場所ではなく、たどり着くまでの速さが変わりました");
  const eff = [
    ["週次PMOの管理台帳として定着", "毎週のPMO会議が、この基盤を見ながら進む状態になった。特別な運用努力なしに続いている。"],
    ["カレンダーから議事録に即到達", "「いつの会議か」さえ分かれば、そこから議事録・決定事項・担当タスクへ一直線に辿れる。"],
    ["決定事項が案件に紐づいて残る", "決めたことが会議単位ではなく案件単位で蓄積され、後から経緯を追える。"],
  ];
  eff.forEach((e, i) => {
    const y = 2.5 + i * 0.78;
    card(s, M, y, 7.5, 0.68, WHITE, LINE);
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 0.18, y: y + 0.17, w: 0.34, h: 0.34, rectRadius: 0.17,
      fill: { color: GREEN }, line: { color: GREEN },
    });
    s.addText("✓", {
      x: M + 0.18, y: y + 0.17, w: 0.34, h: 0.34, align: "center", valign: "middle",
      fontFace: F, fontSize: 12, bold: true, color: WHITE, margin: 0,
    });
    s.addText(e[0], { x: M + 0.64, y: y + 0.06, w: 6.7, h: 0.28, valign: "middle", fontFace: F, fontSize: 12.5, bold: true, color: NAVY, margin: 0 });
    s.addText(e[1], { x: M + 0.64, y: y + 0.34, w: 6.7, h: 0.3, fontFace: F, fontSize: 9.5, color: INK, margin: 0 });
  });

  card(s, M + 7.75, 2.5, 4.38, 2.24, NAVY, NAVY);
  s.addText("探したいときに、探せる", {
    x: M + 7.98, y: 2.7, w: 3.95, h: 0.62, valign: "middle",
    fontFace: F, fontSize: 20, bold: true, color: WHITE, margin: 0,
  });
  s.addText("価値は「どこに置くか」ではなく「どれだけ速くたどり着けるか」。\n横断検索と案件との紐づけで、過去の経緯に数クリックで届きます。", {
    x: M + 7.98, y: 3.35, w: 3.95, h: 1.1, fontFace: F, fontSize: 10.5, color: ICE, margin: 0, lineSpacing: 18,
  });

  card(s, M, 4.95, CW, 1.55, ICE_PALE, ICE);
  s.addText("従来（SharePoint・OneNote 分散）と何が違うか", {
    x: M + 0.3, y: 5.1, w: 6, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: NAVY, margin: 0,
  });
  const cmp = [
    ["従来", "「どのフォルダだったか」を思い出すところから始まる。担当者に聞かないと分からないことがある。", MUTED],
    ["Notion", "案件から辿る／検索から辿るの2経路があり、担当者に聞かなくても本人が到達できる。", NAVY],
  ];
  cmp.forEach((c, i) => {
    const x = M + 0.3 + i * 5.9;
    s.addText(c[0], { x, y: 5.48, w: 5.5, h: 0.26, fontFace: F, fontSize: 10, bold: true, color: c[2], margin: 0 });
    s.addText(c[1], { x, y: 5.74, w: 5.5, h: 0.62, fontFace: F, fontSize: 10.5, color: INK, margin: 0 });
  });
  s.addNotes("効果KPIとして「案件進捗の可視化率」「議事録への到達時間」を第7章で設定する旨をここで予告してもよい。");
}

/* ═══════════════ 11. 第6章 6-1 第三者認証 ═══════════════ */
{
  const s = content("第6章", "セキュリティ — 第三者認証の取得状況", "国際的な第三者監査を毎年受けており、認証の面でのご懸念はありません", NAVY_MID);
  const certs = [
    ["SOC 2 Type II", "セキュリティ・可用性・機密性に関する第三者監査報告書"],
    ["SOC 3", "一般公開向けの監査報告書"],
    ["ISO/IEC 27001", "情報セキュリティマネジメントシステム（ISMS）"],
    ["ISO/IEC 27017", "クラウドサービスのセキュリティ"],
    ["ISO/IEC 27018", "パブリッククラウドにおける個人データ保護"],
    ["ISO/IEC 27701", "プライバシー情報マネジメント（PIMS）"],
    ["BSI C5", "ドイツ連邦情報セキュリティ庁のクラウド基準"],
    ["GDPR / HIPAA", "GDPR準拠。HIPAA BAA は Enterprise 契約時に対応"],
  ];
  certs.forEach((c, i) => {
    const w = CW / 4 - 0.16;
    const x = M + (i % 4) * (w + 0.213);
    const y = 2.5 + Math.floor(i / 4) * 1.28;
    card(s, x, y, w, 1.12, WHITE, LINE);
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.2, y: y + 0.18, w: 0.34, h: 0.34, rectRadius: 0.17,
      fill: { color: ICE }, line: { color: ICE },
    });
    s.addText("✓", {
      x: x + 0.2, y: y + 0.18, w: 0.34, h: 0.34, align: "center", valign: "middle",
      fontFace: F, fontSize: 11, bold: true, color: NAVY, margin: 0,
    });
    s.addText(c[0], { x: x + 0.63, y: y + 0.18, w: w - 0.83, h: 0.34, valign: "middle", fontFace: F, fontSize: 12, bold: true, color: NAVY, margin: 0 });
    s.addText(c[1], { x: x + 0.2, y: y + 0.58, w: w - 0.4, h: 0.44, fontFace: F, fontSize: 9, color: INK, margin: 0 });
  });

  card(s, M, 5.2, CW, 1.3, ICE_PALE, ICE);
  s.addText("一次資料のご確認先", {
    x: M + 0.3, y: 5.35, w: 5, h: 0.28, fontFace: F, fontSize: 11.5, bold: true, color: NAVY, margin: 0,
  });
  s.addText("監査報告書・認証書は Notion Trust Center（trust.notion.com）にて公開されています。\n情報システム部門にて直接ご確認いただけます。必要であれば、当社より取得のうえご提出します。", {
    x: M + 0.3, y: 5.66, w: CW - 0.6, h: 0.7, fontFace: F, fontSize: 11.5, color: INK, margin: 0, lineSpacing: 20,
  });
  s.addNotes("この枚の目的は、情シスが自分で一次資料を引ける状態にしておくこと。認証名を読み上げるのではなく、Trust Centerで確認できる旨を伝えるだけでよい。");
}

/* ═══════════════ 12. 第6章 6-2 機能×プラン ═══════════════ */
{
  const s = content("第6章", "セキュリティ — 統制機能とプランの対応", "THD側で何を必須とされるかによって、選ぶべきプランが一意に決まります", NAVY_MID);
  const L = { bold: true, align: "left" };
  const rows = [
    [hdr("機能要件", { align: "left" }), hdr("Plus"), hdr("Business", { fill: { color: GREEN } }), hdr("Enterprise")],
    [tx("SAML SSO（シングルサインオン）", L), mk("×"), mk("○"), mk("○")],
    [tx("監査ログ（Audit Log）", L), mk("×"), mk("×"), mk("○")],
    [tx("SCIM（自動IDプロビジョニング）", L), mk("×"), mk("×"), mk("○")],
    [tx("外部共有・公開の組織一括禁止", L), mk("×"), mk("×"), mk("○")],
    [tx("SIEM / DLP 連携", L), mk("×"), mk("×"), mk("○")],
    [tx("プライベートチームスペース", L), mk("×"), mk("○"), mk("○")],
    [tx("データ履歴保持（バージョン復元）", L), tx("30日", { align: "center" }), tx("90日", { align: "center", bold: true }), tx("無制限", { align: "center" })],
  ];
  table(s, rows, 2.44, [5.53, 2.2, 2.2, 2.2], { rowH: 0.36, fontSize: 10, align: "center" });

  const boxes = [
    ["SSO が必須要件の場合", "Business プラン以上", GREEN, GREEN_PALE],
    ["監査ログ・SCIM も必須の場合", "Enterprise プラン一択", AMBER, "FDF6E7"],
  ];
  boxes.forEach((b, i) => {
    const x = M + i * (CW / 2 + 0.12);
    const w = CW / 2 - 0.12;
    card(s, x, 5.62, w, 0.95, b[3], b[2]);
    s.addText(b[0], { x: x + 0.25, y: 5.74, w: w - 0.5, h: 0.28, fontFace: F, fontSize: 10.5, color: INK, margin: 0 });
    s.addText(b[1], { x: x + 0.25, y: 6.02, w: w - 0.5, h: 0.4, fontFace: F, fontSize: 16, bold: true, color: b[2], margin: 0 });
  });

  s.addText("PCAからのご提案：Business プラン（SSO要件を満たす最小構成）。監査ログを必須とされるかは第8章にてご確認させてください。", {
    x: M, y: 6.7, w: CW, h: 0.3, fontFace: F, fontSize: 11, bold: true, color: NAVY, margin: 0,
  });
  s.addNotes("要件を先に確定させてから金額の話に進む。監査ログが必須かどうかでプランが変わるため、この場で確認を取りたい旨を伝える。");
}

/* ═══════════════ 13. 第6章 6-3 なぜSSOが要るのか ═══════════════ */
{
  const s = content("第6章", "セキュリティ — なぜ SAML SSO が必要なのか", "Plusの「Microsoftでサインイン」は個人の利便性であり、会社の統制ではありません", NAVY_MID);
  const rows = [
    [hdr("論点", { align: "left" }), hdr("① 通常ログイン / OAuth（Plus）"), hdr("② SAML SSO（Business以上）", { fill: { color: GREEN } })],
    [tx("ログイン方法", { bold: true }), tx("各自がID/PWを入力、またはMicrosoftボタンを押す"), tx("社内M365（Entra ID）の認証画面へ自動転送")],
    [tx("SSOの強制", { bold: true }), mk("×", "社員がNotion専用のパスワードを作れてしまう"), mk("○", "SSO以外のログインを遮断できる")],
    [tx("M365ポリシーの適用\n（MFA・端末制限・IP制限）", { bold: true }), mk("△", "パスワードで抜け穴ができ、適用が不完全"), mk("○", "条件付きアクセスが100%適用される")],
    [tx("退職時のアクセス遮断", { bold: true }), mk("×", "Notion側の手動削除を忘れると入れ続ける"), mk("○", "M365停止と同時に即時遮断")],
    [tx("情シスによる一元管理", { bold: true }), mk("×", "社員個人のアカウント管理に依存"), mk("○", "M365の管理画面から一元統制")],
  ];
  table(s, rows, 2.44, [2.9, 4.62, 4.61], { rowH: 0.42, fontSize: 9.5 });

  const paths = [
    ["SSO なし（Plus プラン）", RED, RED_PALE,
      "退職  →  情シスがM365を停止  →  Notion側の削除を失念\n→  元社員が自宅PCから社内Notionを閲覧・コピー可能",
      "機密情報・ノウハウの流出"],
    ["SSO あり（Business プラン以上）", GREEN, GREEN_PALE,
      "退職  →  情シスがM365を停止  →  Notionの鍵も同時に自動ロック\n→  即座にアクセス不能",
      "人為的な削除ミスによる流出を防止"],
  ];
  paths.forEach((p, i) => {
    const w = CW / 2 - 0.12;
    const x = M + i * (w + 0.24);
    card(s, x, 5.42, w, 1.5, p[2], p[1]);
    s.addText(p[0], { x: x + 0.22, y: 5.54, w: w - 0.44, h: 0.28, fontFace: F, fontSize: 11.5, bold: true, color: p[1], margin: 0 });
    s.addText(p[3], { x: x + 0.22, y: 5.84, w: w - 0.44, h: 0.6, fontFace: F, fontSize: 10, color: INK, margin: 0, lineSpacing: 17 });
    s.addText(p[4], { x: x + 0.22, y: 6.5, w: w - 0.44, h: 0.32, valign: "middle", fontFace: F, fontSize: 12, bold: true, color: p[1], margin: 0 });
  });
  s.addNotes("この資料で最も重要な1枚。Plusを選ばない理由の唯一の論拠。SSOの目的は利便性ではなく、退職者による情報持ち出しを会社として確実に防ぐこと、と言い切る。");
}

/* ═══════════════ 14. 第6章 6-4 M365環境での挙動 ═══════════════ */
{
  const s = content("第6章", "セキュリティ — 当社M365環境での各プランの挙動", "当社はすでに Microsoft 365（Entra ID）をID基盤として運用しています", NAVY_MID);
  const rows = [
    [hdr("プラン", { align: "left" }), hdr("M365（Entra ID）との連携"), hdr("THD審査の想定結果")],
    [tx("Plus", { bold: true, fontSize: 12 }),
     tx("SAML SSO 連携は不可。各自が M365 のメールアドレスで個別登録する形となり、退職者の手動削除漏れリスクが残る。"),
     mk("×", "セキュリティポリシー違反として否決されるリスクが大きい")],
    [{ text: "Business", options: { color: WHITE, bold: true, fontSize: 12, fill: { color: GREEN }, align: "left", valign: "middle" } },
     tx("SAML SSO 連携が可能（プロビジョニングは手動）。M365 の認証でログインを統一でき、M365 を止めればアクセスを遮断できる。", { bold: true }),
     mk("○", "「SSO必須」の要件をクリア。承認いただける可能性が高い")],
    [tx("Enterprise", { bold: true, fontSize: 12 }),
     tx("SAML SSO ＋ SCIM（自動同期）連携が可能。M365側で社員を追加・異動・削除すると、Notion側も自動で作成・削除される。"),
     mk("○", "大手ホールディングス基準を完全にクリア（最も安全）")],
  ];
  table(s, rows, 2.42, [1.75, 6.2, 4.18], { rowH: [0.34, 0.72, 0.72, 0.72], fontSize: 9.5 });

  card(s, M, 5.72, 7.55, 1.28, "FDF6E7", AMBER);
  s.addText("Business プランでできないこと（あらかじめ申し上げます）", {
    x: M + 0.25, y: 5.84, w: 7.05, h: 0.26, fontFace: F, fontSize: 11, bold: true, color: AMBER, margin: 0,
  });
  s.addText("監査ログと SCIM は Business では利用できません。①②のスコープでは「機微情報を Notion に置かない」運用で担保します。監査ログを必須要件とされる場合は Enterprise への切替となり、年額は約911.4万円（140名）に変動します。", {
    x: M + 0.25, y: 6.12, w: 7.05, h: 0.76, fontFace: F, fontSize: 9.5, color: INK, margin: 0, lineSpacing: 15,
  });

  card(s, M + 7.78, 5.72, 4.35, 1.28, ICE_PALE, ICE);
  s.addText("退職者のアカウント運用", {
    x: M + 8.0, y: 5.84, w: 3.9, h: 0.26, fontFace: F, fontSize: 11, bold: true, color: NAVY, margin: 0,
  });
  s.addText("SSOによりM365停止と同時にログイン不能。課金停止のためのメンバー削除は手動運用とし、対象者リストを月次で突合します。", {
    x: M + 8.0, y: 6.12, w: 3.9, h: 0.76, fontFace: F, fontSize: 9.5, color: INK, margin: 0, lineSpacing: 15,
  });
  s.addNotes("稟議トーク：「当社ではすでに全社のID認証基盤としてMicrosoft 365（Entra ID）を導入・運用しております。NotionはBusinessプラン以上のSAML SSO機能でEntra IDと認証を統合します。これにより社員のパスワード管理リスクを排除し、M365の多要素認証や端末制限ポリシーをNotionにも強制適用できます。退職時にもM365アカウントの無効化と同時にNotionへのアクセスが即座に遮断されます。Plusではこの組織統制ができないため、Businessプランでの導入を提案します。」弱点を先に自己申告することで、質疑での信頼を確保する。");
}

/* ═══════════════ 15. 第7章 7-1 別軸 ═══════════════ */
{
  const s = content("第7章", "【別軸】基盤の上で目指す運用改革", "ここからは別軸のお話です。仕組みができて初めて可能になる「運用」の話であり、ご承認の対象ではありません", NAVY_MID);
  s.addText("会議の目的を先に決め、決まったことが確実に現場まで届く状態にする", {
    x: M, y: 2.42, w: CW, h: 0.34, fontFace: F, fontSize: 14, bold: true, color: NAVY, margin: 0,
  });

  const steps = [
    ["1", "会議前", "ゴールとアジェンダを作って共有", "議事録ページを会議の前に起票し、この会議で決めることを明示。参加者が準備して臨める。"],
    ["2", "会議中", "その場で議事録化", "決定事項と宿題（担当者・期限）を、会議の場で確定させる。"],
    ["3", "会議後", "タスクと案件に接続", "決定事項がタスクDBに紐づき、案件・イシューにつながる。"],
    ["4", "全社展開", "各MGを通じて現場まで届ける", "全社共有用ページを用意し、同期ブロック（シンクロ）で各部門ページへ配信する。"],
  ];
  const sw = 2.845, sg = 0.25;
  steps.forEach((st, i) => {
    const x = M + i * (sw + sg);
    const hl = i === 3;
    card(s, x, 2.85, sw, 2.0, hl ? ICE_PALE : WHITE, hl ? ICE : LINE);
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.2, y: 3.02, w: 0.36, h: 0.36, rectRadius: 0.18,
      fill: { color: NAVY }, line: { color: NAVY },
    });
    s.addText(st[0], {
      x: x + 0.2, y: 3.02, w: 0.36, h: 0.36, align: "center", valign: "middle",
      fontFace: F, fontSize: 11, bold: true, color: WHITE, margin: 0,
    });
    s.addText(st[1], { x: x + 0.65, y: 3.02, w: sw - 0.85, h: 0.36, valign: "middle", fontFace: F, fontSize: 13, bold: true, color: NAVY, margin: 0 });
    s.addText(st[2], { x: x + 0.2, y: 3.48, w: sw - 0.4, h: 0.56, fontFace: F, fontSize: 11, bold: true, color: INK, margin: 0 });
    s.addText(st[3], { x: x + 0.2, y: 4.06, w: sw - 0.4, h: 0.7, fontFace: F, fontSize: 9, color: MUTED, margin: 0 });
    if (i < 3) {
      s.addText("▶", {
        x: x + sw + 0.01, y: 3.6, w: sg - 0.02, h: 0.3, align: "center", valign: "middle",
        fontFace: F, fontSize: 11, color: NAVY_MID, margin: 0,
      });
    }
  });

  card(s, M, 5.05, 7.55, 1.55, WHITE, LINE);
  s.addText("全社展開のイメージ", { x: M + 0.25, y: 5.18, w: 4, h: 0.28, fontFace: F, fontSize: 11.5, bold: true, color: NAVY, margin: 0 });
  s.addText("例：経営会議で決定した全社方針を、全部門へ展開する必要が生じたとき。全社共有用ページの内容を各部門ページへ同期配信し、原本を1回直せば展開先すべてに即時反映されます。各MGが個別に転記・再作成する必要がなくなり、伝達漏れ・版ズレ・「言った/聞いてない」がなくなります。", {
    x: M + 0.25, y: 5.5, w: 7.05, h: 0.95, fontFace: F, fontSize: 10, color: INK, margin: 0, lineSpacing: 16,
  });

  card(s, M + 7.78, 5.05, 4.35, 1.55, RED_PALE, RED);
  s.addText("設計上の注意（重要）", { x: M + 8.0, y: 5.18, w: 3.9, h: 0.28, fontFace: F, fontSize: 11.5, bold: true, color: RED, margin: 0 });
  s.addText("同期ブロックは「同期元ページの閲覧権限」が適用されます。経営会議の議事録（非公開）に原本を置くと、一般社員には中身が表示されません。原本は必ず全社共有ページ側に置きます（Appendix I）。", {
    x: M + 8.0, y: 5.5, w: 3.9, h: 0.95, fontFace: F, fontSize: 9.5, color: INK, margin: 0, lineSpacing: 15,
  });
  s.addNotes("口頭補足：「これは単なるツールの使い方ではなく、経営が全社の働き方としてこうなってほしいと考えている姿です。会議の目的を先に決めて共有し、決まったことが確実に現場まで届く状態を作りたい。その受け皿としてこの基盤が必要になります。」承認対象ではないことを明示し、次の7-2で軸を戻す。");
}

/* ═══════════════ 16. 第7章 7-2 計画とコスト ═══════════════ */
{
  const s = content("第7章", "導入計画 — 小さく始めて、効果で拡張を判断します", "話を戻します。スコープ・コスト・撤退基準を先に限定し、失敗リスクを封じます", NAVY_MID);
  const phases = [
    ["Phase 1", "経営・PMO", "実施済", "経営層で3か月の実運用。効果を確認済み。", GREEN],
    ["Phase 2", "部門展開", "今回のご承認対象", "約100名から開始し、着地140名。テンプレート＋ミニマニュアル整備後に展開。", NAVY],
    ["Phase 3", "拡張判断", "効果測定後", "効果KPIを測定し、達しなければ拡張しない。", MUTED],
  ];
  phases.forEach((p, i) => {
    const w = 3.72;
    const x = M + i * (w + 0.24);
    card(s, x, 2.5, w, 1.5, i === 1 ? ICE_PALE : WHITE, i === 1 ? ICE : LINE);
    s.addText(p[0], { x: x + 0.22, y: 2.62, w: w - 0.44, h: 0.28, fontFace: F, fontSize: 10, bold: true, color: p[4], margin: 0 });
    s.addText(p[1], { x: x + 0.22, y: 2.9, w: w - 0.44, h: 0.32, fontFace: F, fontSize: 14, bold: true, color: NAVY, margin: 0 });
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.22, y: 3.26, w: 1.85, h: 0.26, rectRadius: 0.04,
      fill: { color: p[4] }, line: { color: p[4] },
    });
    s.addText(p[2], {
      x: x + 0.22, y: 3.26, w: 1.85, h: 0.26, align: "center", valign: "middle",
      fontFace: F, fontSize: 8.5, bold: true, color: WHITE, margin: 0,
    });
    s.addText(p[3], { x: x + 0.22, y: 3.58, w: w - 0.44, h: 0.34, fontFace: F, fontSize: 9, color: INK, margin: 0 });
    if (i < 2) {
      s.addText("▶", { x: x + w + 0.01, y: 3.1, w: 0.22, h: 0.3, align: "center", valign: "middle", fontFace: F, fontSize: 11, color: NAVY_MID, margin: 0 });
    }
  });

  card(s, M, 4.2, 4.7, 2.4, NAVY, NAVY);
  s.addText("コスト上限", { x: M + 0.3, y: 4.36, w: 4.1, h: 0.28, fontFace: F, fontSize: 11.5, bold: true, color: ICE, margin: 0 });
  s.addText("約520.8万円", { x: M + 0.3, y: 4.66, w: 4.1, h: 0.75, fontFace: F, fontSize: 36, bold: true, color: WHITE, margin: 0 });
  s.addText("/ 年（140名・Business・年払）", { x: M + 0.3, y: 5.4, w: 4.1, h: 0.28, fontFace: F, fontSize: 11, color: ICE, margin: 0 });
  s.addText("月額換算 約43.4万円　／　1人あたり 年37,200円\n＝ 1営業日あたり 約169円", {
    x: M + 0.3, y: 5.75, w: 4.1, h: 0.6, fontFace: F, fontSize: 10, color: ICE, margin: 0, lineSpacing: 17,
  });

  card(s, M + 4.93, 4.2, 3.5, 2.4, WHITE, LINE);
  s.addText("統制（3点）", { x: M + 5.15, y: 4.36, w: 3.1, h: 0.28, fontFace: F, fontSize: 11.5, bold: true, color: NAVY, margin: 0 });
  s.addText(
    [
      { text: "コストは対象人数×単価で上限を明示する", options: { bullet: true, breakLine: true } },
      { text: "SSO・権限で統制し、機微情報はNotionに置かない", options: { bullet: true, breakLine: true } },
      { text: "効果KPI（案件進捗の可視化率・議事録への到達時間）が未達なら拡張しない", options: { bullet: true } },
    ],
    { x: M + 5.15, y: 4.7, w: 3.1, h: 1.75, fontFace: F, fontSize: 9.5, color: INK, margin: 0, paraSpaceAfter: 7 }
  );

  card(s, M + 8.66, 4.2, 3.47, 2.4, WHITE, LINE);
  s.addText("コスト最適化", { x: M + 8.88, y: 4.36, w: 3.05, h: 0.28, fontFace: F, fontSize: 11.5, bold: true, color: NAVY, margin: 0 });
  s.addText("段階的な調達（True-up）", { x: M + 8.88, y: 4.7, w: 3.05, h: 0.24, fontFace: F, fontSize: 10, bold: true, color: NAVY, margin: 0 });
  s.addText("初期は約100名から開始し、展開に合わせて追加購入（日割り精算）。上限140名の範囲内で実人数のみ調達します。", {
    x: M + 8.88, y: 4.96, w: 3.05, h: 0.7, fontFace: F, fontSize: 9, color: INK, margin: 0, lineSpacing: 14,
  });
  s.addText("外部委託者の切り分け", { x: M + 8.88, y: 5.66, w: 3.05, h: 0.24, fontFace: F, fontSize: 10, bold: true, color: NAVY, margin: 0 });
  s.addText("業務委託・社外パートナーはゲスト（無料）として必要ページのみに招待し、課金対象を正規社員に限定します。", {
    x: M + 8.88, y: 5.92, w: 3.05, h: 0.6, fontFace: F, fontSize: 9, color: INK, margin: 0, lineSpacing: 14,
  });

  s.addText("※ $20/月は年払時の想定単価（$1=155円換算）。正規代理店の見積により変動しますが、上限額は上振れしない前提でご提示しています。", {
    x: M, y: 6.72, w: CW, h: 0.28, fontFace: F, fontSize: 9, color: MUTED, margin: 0,
  });
  s.addNotes("撤退基準を先に定義していることが、承認する側にとって最も安心材料になる。①②に限定しているからこそ、やめられるという点を強調する。");
}

/* ═══════════════ 17. 第8章 お願い事項 ═══════════════ */
{
  const s = content("第8章", "お願い事項", "①案件管理 ＋ ②議事録 に限定した正式導入と、コスト上限のご承認をいただきたく存じます");
  s.addText("ご承認いただきたいこと", { x: M, y: 2.42, w: 5, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: MUTED, margin: 0 });
  const asks = [
    ["1", "①案件管理 ＋ ②議事録 に限定した正式導入"],
    ["2", "Business プランでの契約（SAML SSO 要件を満たすため）"],
    ["3", "Phase2 のコスト上限：140名 ／ 年額 約520.8万円"],
  ];
  asks.forEach((a, i) => {
    const y = 2.78 + i * 0.62;
    card(s, M, y, 7.55, 0.54, ICE_PALE, ICE);
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 0.16, y: y + 0.11, w: 0.32, h: 0.32, rectRadius: 0.16,
      fill: { color: NAVY }, line: { color: NAVY },
    });
    s.addText(a[0], {
      x: M + 0.16, y: y + 0.11, w: 0.32, h: 0.32, align: "center", valign: "middle",
      fontFace: F, fontSize: 10, bold: true, color: WHITE, margin: 0,
    });
    s.addText(a[1], { x: M + 0.6, y, w: 6.85, h: 0.54, valign: "middle", fontFace: F, fontSize: 12.5, bold: true, color: NAVY, margin: 0 });
  });

  s.addText("PCA側のお約束", { x: M, y: 4.72, w: 5, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: MUTED, margin: 0 });
  card(s, M, 5.06, 7.55, 1.5, WHITE, LINE);
  s.addText(
    [
      { text: "③案件関連ファイル・④社内規定は SharePoint を継続（M365 は置き換えません）", options: { bullet: true, breakLine: true } },
      { text: "機微情報は Notion に置かない運用とします", options: { bullet: true, breakLine: true } },
      { text: "効果測定の結果と、拡張の要否をご報告します", options: { bullet: true } },
    ],
    { x: M + 0.28, y: 5.24, w: 7.0, h: 1.15, fontFace: F, fontSize: 11, color: INK, margin: 0, paraSpaceAfter: 8 }
  );

  card(s, M + 7.78, 2.78, 4.35, 3.78, NAVY, NAVY);
  s.addText("THDへのご確認事項", { x: M + 8.0, y: 2.95, w: 3.9, h: 0.32, fontFace: F, fontSize: 13, bold: true, color: WHITE, margin: 0 });
  const qs = [
    ["1", "SaaS導入手続きの様式", "グループ所定のセキュリティチェックシート等がございましたら、受領のうえ別途ご回答を作成します。"],
    ["2", "監査ログを必須とされるか", "必須の場合は Enterprise が対象となり、年額は約911.4万円（140名）に変動します。"],
    ["3", "認証資料のご提出要否", "必要な場合、提出先と形式をご指定ください。"],
  ];
  qs.forEach((q, i) => {
    const y = 3.4 + i * 1.05;
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 8.0, y, w: 0.3, h: 0.3, rectRadius: 0.15,
      fill: { color: ICE }, line: { color: ICE },
    });
    s.addText(q[0], { x: M + 8.0, y, w: 0.3, h: 0.3, align: "center", valign: "middle", fontFace: F, fontSize: 9, bold: true, color: NAVY, margin: 0 });
    s.addText(q[1], { x: M + 8.4, y: y - 0.02, w: 3.5, h: 0.3, valign: "middle", fontFace: F, fontSize: 11, bold: true, color: WHITE, margin: 0 });
    s.addText(q[2], { x: M + 8.4, y: y + 0.3, w: 3.5, h: 0.62, fontFace: F, fontSize: 9, color: ICE, margin: 0, lineSpacing: 14 });
  });
  s.addNotes("この場で②監査ログの要否について回答をいただけると、プランとコストが確定する。①の様式については持ち帰り確認でも構わない旨を伝える。");
}

/* ═══════════════ 18. Appendix 扉 ═══════════════ */
{
  const s = pres.addSlide();
  s.background = { color: NAVY_DEEP };
  s.addText("Appendix", {
    x: 0.9, y: 2.5, w: 6, h: 0.9, fontFace: F, fontSize: 44, bold: true, color: WHITE, margin: 0,
  });
  s.addText("補強材料 — ご質問に応じてご参照ください", {
    x: 0.9, y: 3.4, w: 8, h: 0.4, fontFace: F, fontSize: 15, color: ICE, margin: 0,
  });
  const ap = [
    ["A", "4案比較表（詳細版）"], ["B", "想定問答"], ["C", "プラン別導入パターンの比較"],
    ["D", "SSOのメリット・デメリット"], ["E", "コスト感度分析"], ["F", "ファイル仕様の確認結果"],
    ["G", "役割分担と残タスク"], ["H", "出典一覧"], ["I", "会議設計と全社展開の設計"],
  ];
  ap.forEach((a, i) => {
    const x = 0.9 + (i % 3) * 3.95;
    const y = 4.3 + Math.floor(i / 3) * 0.62;
    s.addShape(pres.ShapeType.roundRect, {
      x, y, w: 0.36, h: 0.36, rectRadius: 0.06,
      fill: { color: NAVY_MID }, line: { color: NAVY_MID },
    });
    s.addText(a[0], { x, y, w: 0.36, h: 0.36, align: "center", valign: "middle", fontFace: F, fontSize: 11, bold: true, color: ICE, margin: 0 });
    s.addText(a[1], { x: x + 0.48, y, w: 3.3, h: 0.36, valign: "middle", fontFace: F, fontSize: 12, color: WHITE, margin: 0 });
  });
  footer(s);
}

/* ═══════════════ 19. Appendix A ═══════════════ */
{
  const s = content("Appendix A", "4案比較表（詳細版）", "判定の理由を各欄に注記しています。消去の論拠は本編 第4章をご参照ください", NAVY_MID);
  const rows = [
    [hdr("要件", { align: "left" }), hdr("A. M365の組合せ"), hdr("B. 無料ツール"), hdr("C. 専用ツール"), hdr("D. Notion", { fill: { color: GREEN } })],
    [tx("1. 相互リンク（DB型）", { bold: true }), mk("△", "Planner/Lists/Loopが個別に存在し、案件と議事録を関連付ける仕組みが無い"), mk("△", "リレーション機能を持たないものが大半"), mk("△", "タスク管理に特化し、議事録・wikiは別ツールになる"), mk("○", "データベース間のリレーションで案件⇔タスク⇔議事録を相互参照")],
    [tx("2. 立場別ビュー", { bold: true }), mk("×", "同一データを別形式で見る仕組みが無い"), mk("△", "ビューの種類が限られる"), mk("○", "ガント・ボード等に対応"), mk("○", "同一DBをガント/ボード/カレンダー/テーブルで切替")],
    [tx("3. 横断検索・即到達", { bold: true }), mk("△", "SharePoint/OneNote/Plannerをまたぐ検索が実用水準にない"), mk("△", "ツール単位の検索に留まる"), mk("△", "議事録が対象外のため片手落ち"), mk("○", "1か所に集約されているため検索が到達手段として機能する")],
    [tx("4. 非エンジニアが使える", { bold: true }), mk("△", "複数ツールそれぞれの習熟が必要"), mk("○", "操作は簡易"), mk("△", "プロジェクト管理の専門知識を前提とする"), mk("○", "UIは平易。テンプレート＋ミニマニュアルで補う")],
    [tx("5. ガバナンス（SSO・権限）", { bold: true }), mk("○", "M365の統制下にある"), mk("×", "SSOが有料機能。無料利用はシャドーIT化する"), mk("○", "多くが対応"), mk("○", "SAML SSO は Business 以上で対応（第6章）")],
    [tx("6. 1製品で完結", { bold: true }), mk("×", "3製品以上の併用が前提"), mk("×", "用途ごとに別ツール"), mk("×", "議事録・wikiで別ツール併用"), mk("○", "案件・タスク・議事録・wikiを1製品で完結")],
    [tx("（参考）社内での実証", { bold: true }), tx("—", { align: "center" }), tx("—", { align: "center" }), tx("—", { align: "center" }), mk("◎", "経営層で3か月の実運用。週次PMOの管理台帳として定着")],
  ];
  table(s, rows, 2.42, [2.35, 2.5, 2.35, 2.35, 2.58], { rowH: [0.3, 0.62, 0.62, 0.62, 0.55, 0.55, 0.5, 0.55], fontSize: 7.5 });
  s.addNotes("この表は質疑があった場合にのみ開く。通常は本編4-1で足りる。");
}

/* ═══════════════ 20. Appendix B ═══════════════ */
{
  const s = content("Appendix B", "想定問答", "THDおよび情報システム部門からのご質問を想定した回答の骨子です", NAVY_MID);
  const rows = [
    [hdr("想定されるご質問", { align: "left" }), hdr("回答の骨子", { align: "left" })],
    [tx("M365があるのに二重投資では？", { bold: true }), tx("役割が異なります（M365＝保管・伝達／Notion＝構造化された実行管理）。M365は継続し、置き換えません。")],
    [tx("M365 Copilot で良いのでは？", { bold: true }), tx("Copilotは検索・要約の道具であり、案件管理の仕組み自体は提供しません。AIは基盤の代替ではなく基盤の上に乗るもので、構造化DBがある方がAI活用の精度は上がります。")],
    [tx("Planner や Lists で良いのでは？", { bold: true }), tx("案件⇔議事録⇔タスクのリレーションと立場別ビューが実務水準に達しません。複数アプリの組合せは、いま解決したい「分散」の再生産です。")],
    [tx("セキュリティは大丈夫か？", { bold: true }), tx("SOC2 Type II・ISO27001他を取得済み。SAML SSOでM365と統合し、機微情報は置かない運用とします（第6章）。")],
    [tx("定着するのか？", { bold: true }), tx("経営層での3か月の実運用実績に加え、テンプレート＋ミニマニュアル＋段階展開で担保します。")],
    [tx("失敗したらどうするのか？", { bold: true }), tx("撤退基準を事前に定義しています。①②に限定しているからこそ、やめられます。")],
    [tx("監査ログが無いBusinessで、情報持ち出しをどう追跡するのか？", { bold: true, color: NAVY }), tx("Businessでは追跡できません。①②のスコープでは機微情報を置かない運用で担保します。監査ログを必須要件とされる場合はEnterpriseへの切替となり、年額約911.4万円（140名）に変動します。")],
    [tx("退職者のアカウントはどう消すのか？", { bold: true, color: NAVY }), tx("SAML SSOによりM365アカウント停止と同時にログイン不能となります。BusinessはSCIM非対応のため、課金停止のためのメンバー削除は手動運用とし、対象者リストを月次で突合します。")],
    [tx("業務委託・社外パートナーはどうするのか？", { bold: true, color: NAVY }), tx("ゲスト（無料）として必要ページのみに招待します。ライセンス課金対象は正規社員に限定し、ゲスト運用ルールは導入前に策定します。")],
  ];
  table(s, rows, 2.44, [4.0, 8.13], { rowH: 0.44, fontSize: 8.5 });
  s.addNotes("下3問は情報システム部門が同席された場合を想定した追加分。");
}

/* ═══════════════ 21. Appendix C-1 ═══════════════ */
{
  const s = content("Appendix C-1", "プラン別 導入パターンの比較", "検討した3案と、参考としてのEnterpriseを同じ軸で比較しています", NAVY_MID);
  const rows = [
    [hdr("比較項目", { align: "left" }), hdr("案A：Business 全員", { fill: { color: GREEN } }), hdr("案B：幹部Business＋他Free"), hdr("案C：Plus 全員"), hdr("（参考）Enterprise 全員")],
    [tx("月額単価（年払・1人）", { bold: true }), tx("$20", { align: "center", bold: true }), tx("幹部のみ $20 ／ 他 $0", { align: "center" }), tx("$10", { align: "center" }), tx("個別見積（$30〜40目安）", { align: "center" })],
    [tx("年額（140名）", { bold: true }), tx("約520.8万円", { align: "center", bold: true, color: NAVY, fontSize: 11 }), tx("幹部人数による", { align: "center" }), tx("約260.4万円", { align: "center" }), tx("約911.4万円", { align: "center" })],
    [tx("SAML SSO", { bold: true }), mk("○", "全員がIdP認証可能"), mk("×", "一般社員はSSO対象外"), mk("×", "利用不可"), mk("○", "全員がIdP認証可能")],
    [tx("監査ログ・SCIM", { bold: true }), mk("×", "利用不可"), mk("×", "利用不可"), mk("×", "利用不可"), mk("○", "完全対応")],
    [tx("アクセス・権限統制", { bold: true }), mk("○", "部署別チームスペースが可能"), mk("×", "ページごとの手動共有"), mk("△", "閲覧制限が限定的"), mk("○", "組織単位で完全制御")],
    [tx("全社ナレッジ共有", { bold: true }), mk("○", "検索・共同編集が快適"), mk("×", "部分的にしか見られない"), mk("○", "基本的な共有は可能"), mk("○", "検索・共同編集が快適")],
    [tx("THD承認の見通し", { bold: true }), tx("中（監査ログ不要なら可）", { align: "center", bold: true, color: GREEN }), tx("極めて低い（否決リスク大）", { align: "center", color: RED }), tx("低い（SSO非対応がネック）", { align: "center", color: AMBER }), tx("高い（要件を全クリア）", { align: "center", color: GREEN })],
  ];
  table(s, rows, 2.42, [2.35, 2.5, 2.45, 2.35, 2.48], { rowH: [0.32, 0.36, 0.4, 0.44, 0.44, 0.44, 0.44, 0.36], fontSize: 8.5 });

  card(s, M, 6.22, CW, 0.7, NAVY, NAVY);
  s.addText("SSO を必須要件とする限り、実務的に成立するのは案A（Business 全員）のみです", {
    x: M + 0.35, y: 6.22, w: CW - 0.7, h: 0.7, valign: "middle",
    fontFace: F, fontSize: 13, bold: true, color: WHITE, margin: 0,
  });
}

/* ═══════════════ 22. Appendix C-2 ═══════════════ */
{
  const s = content("Appendix C-2", "各案の詳細分析", "C-1の判定に至った理由の裏付けです", NAVY_MID);
  const cases = [
    ["案A：Business 全員", "推奨度 中〜高／本資料のご提案", GREEN, GREEN_PALE,
      "できること：全員がIdP経由でSAML SSOログイン可能。部署ごとのプライベートチームスペースを作成し、経営情報・人事情報の閲覧を限定できる。データベースの高度な権限設定。",
      "できないこと：監査ログの閲覧・出力（事後追跡が不可）。SCIMによる入退社時の自動同期（手動でのメンバー削除が必要）。",
      "判定：「SSO必須・監査ログは不要」というポリシーであれば、コストと機能のバランスが最も良い。"],
    ["案B：幹部のみBusiness＋他社員Free", "推奨度 不可・非推奨", RED, RED_PALE,
      "できること：費用を最小限に抑えられる。",
      "重大な欠陥：①一般社員はゲスト扱いとなりSAML SSOが適用されず、個人のフリーアカウントでログインする。②ゲストはチームスペース全体を閲覧できず、ページごとに個別招待が必要で全社Wikiとして成立しない。③退職時にM365を停止しても個人アカウントに権限が残る。",
      "判定：セキュリティガバナンスが完全に崩壊するため、THD審査で否決される可能性が極めて高く、運用も不可能。"],
    ["案C：Plus 全員", "推奨度 低〜中", AMBER, "FDF6E7",
      "できること：ページ作成・ブロック数無制限、共同編集、外部ゲスト招待。Businessの半額（140名で約260.4万円/年）。",
      "できないこと：SAML SSOが利用不可（個別のパスワードや二要素認証に依存）。プライベートチームスペースが作成できず、部署限定の非公開空間を作りにくい。",
      "判定：THD側が「SAML SSOによる一括認証」を必須条件とされる場合、要件未達で却下となる。"],
  ];
  cases.forEach((c, i) => {
    const y = 2.40 + i * 1.36;
    card(s, M, y, CW, 1.28, c[3], c[2]);
    s.addText(c[0], { x: M + 0.25, y: y + 0.1, w: 4.2, h: 0.3, valign: "middle", fontFace: F, fontSize: 13, bold: true, color: c[2], margin: 0 });
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 4.55, y: y + 0.14, w: 2.5, h: 0.24, rectRadius: 0.04,
      fill: { color: c[2] }, line: { color: c[2] },
    });
    s.addText(c[1], { x: M + 4.55, y: y + 0.14, w: 2.5, h: 0.24, align: "center", valign: "middle", fontFace: F, fontSize: 8.5, bold: true, color: WHITE, margin: 0 });
    s.addText(c[4], { x: M + 0.25, y: y + 0.44, w: CW - 0.5, h: 0.26, fontFace: F, fontSize: 9, color: INK, margin: 0 });
    s.addText(c[5], { x: M + 0.25, y: y + 0.7, w: CW - 0.5, h: 0.34, fontFace: F, fontSize: 9, color: INK, margin: 0 });
    s.addText(c[6], { x: M + 0.25, y: y + 1.02, w: CW - 0.5, h: 0.24, fontFace: F, fontSize: 9.5, bold: true, color: c[2], margin: 0 });
  });

  card(s, M, 6.52, CW, 0.46, WHITE, LINE);
  s.addText("（参考）Enterprise 全員：セキュリティ要件を全てクリア（SSO＋監査ログ＋SCIM＋組織統制）。承認難易度は最も低いが、140名で約911.4万円/年。個別見積が必要。", {
    x: M + 0.25, y: 6.52, w: CW - 0.5, h: 0.46, valign: "middle", fontFace: F, fontSize: 9.5, color: INK, margin: 0,
  });
}

/* ═══════════════ 23. Appendix D ═══════════════ */
{
  const s = content("Appendix D", "SSO のメリット・デメリット（4つの視点）", "導入する側・使う側・管理する側それぞれの観点を整理しています", NAVY_MID);
  const rows = [
    [hdr("視点", { align: "left" }), hdr("メリット（良い点）", { align: "left" }), hdr("デメリット・注意点", { align: "left" })],
    [tx("① 社員\n（使う側）", { bold: true, color: NAVY, fontSize: 11 }),
     tx("・パスワードを覚える必要がゼロ。Notion専用のID/PWが不要で、忘れ・再設定の手間がなくなる。\n・毎朝PC（M365）にログインしていれば、ボタン1つで即座にNotionが開く。"),
     tx("・Microsoft側で大規模なシステム障害が起きた場合、Notion単体にも一時的にログインできなくなる。")],
    [tx("② 会社・HD\n（セキュリティ）", { bold: true, color: NAVY, fontSize: 11 }),
     tx("・退職時にM365アカウントを停止すれば、Notionのアクセス権も同時に自動遮断される。\n・「会社の指定PC以外はアクセス禁止」「多要素認証必須」などのM365の制限がそのままNotionにも効く。\n・簡単なパスワードや私生活のパスワードの使い回しによる乗っ取りを防げる。"),
     tx("・自社のM365アカウントを持たない外部パートナーを招待する場合、「ゲスト」としての運用ルールを事前に決める必要がある。")],
    [tx("③ 情シス\n（管理・運用）", { bold: true, color: NAVY, fontSize: 11 }),
     tx("・「PWを忘れたのでリセットしてほしい」という社内問い合わせがゼロになる。\n・入退社時のアカウント削除漏れ（消し忘れリスク）を防げる。"),
     tx("・導入時に、情シスがM365（Entra ID）とNotionを紐付ける初期設定が必要（約1〜2時間程度の作業）。")],
    [tx("④ コスト\n（費用面）", { bold: true, color: NAVY, fontSize: 11 }),
     tx("・アカウント管理ミスによる情報漏洩の損害と、事後対応コストを未然に防止できる。"),
     tx("・SSOを使うにはBusinessプラン以上が必要（月額$20/人〜）。最安のPlus（$10）では利用できず、140名で年間約260.4万円の差額が生じる。")],
  ];
  table(s, rows, 2.42, [1.75, 6.2, 4.18], { rowH: [0.32, 0.85, 1.05, 0.8, 0.85], fontSize: 8.5 });
  s.addNotes("まとめの一言：「SSOを導入する最大の理由は、社員のログインを便利にするためだけでなく、退職者による情報持ち出しや不正アクセスを会社として確実に防ぐためです。全社員が社内マニュアルや業務データを扱うツールだからこそ、安全に統制できるBusinessプランでの導入が必要となります。」");
}

/* ═══════════════ 24. Appendix E ═══════════════ */
{
  const s = content("Appendix E", "コスト感度分析", "対象人数とプランの組み合わせによる年間コストの試算です（$1=155円・年払）", NAVY_MID);
  const rows = [
    [hdr("想定対象人数", { align: "left" }), hdr("Plus（$10/月）"), hdr("Business（$20/月）", { fill: { color: GREEN } }), hdr("Enterprise（仮 $35/月）")],
    [tx("50名", { bold: true }), tx("約 93.0 万円 / 年", { align: "center" }), tx("約 186.0 万円 / 年", { align: "center" }), tx("約 325.5 万円 / 年", { align: "center" })],
    [tx("100名", { bold: true }), tx("約 186.0 万円 / 年", { align: "center" }), tx("約 372.0 万円 / 年", { align: "center" }), tx("約 651.0 万円 / 年", { align: "center" })],
    [{ text: "140名（本計画）", options: { bold: true, color: WHITE, fontSize: 11, fill: { color: NAVY }, valign: "middle" } },
     tx("約 260.4 万円 / 年", { align: "center", fill: { color: ICE_PALE } }),
     { text: "約 520.8 万円 / 年", options: { align: "center", bold: true, fontSize: 13, color: NAVY, fill: { color: ICE_PALE }, valign: "middle" } },
     tx("約 911.4 万円 / 年", { align: "center", fill: { color: ICE_PALE } })],
    [tx("300名", { bold: true }), tx("約 558.0 万円 / 年", { align: "center" }), tx("約 1,116.0 万円 / 年", { align: "center" }), tx("約 1,953.0 万円 / 年", { align: "center" })],
    [tx("500名", { bold: true }), tx("約 930.0 万円 / 年", { align: "center" }), tx("約 1,860.0 万円 / 年", { align: "center" }), tx("約 3,255.0 万円 / 年", { align: "center" })],
  ];
  table(s, rows, 2.44, [3.0, 3.04, 3.04, 3.05], { rowH: 0.44, fontSize: 10.5 });

  const notes = [
    ["前提条件", "年払契約時の金額です。月払契約の場合は約20%割高になります。為替は $1=155円 で換算しています。"],
    ["Enterprise単価について", "ボリュームディスカウントにより変動するため、Notion社または正規代理店への相見積もりが必要です。表中の $35 は仮置きの値です。"],
    ["Business単価について", "$20/月（年払）で試算していますが、$18/月とする情報もあります。見積により下振れする可能性があり、上限額は上振れしない前提です。"],
  ];
  notes.forEach((n, i) => {
    const y = 5.4 + i * 0.5;
    s.addText(n[0], { x: M, y, w: 2.3, h: 0.42, valign: "middle", fontFace: F, fontSize: 10, bold: true, color: NAVY, margin: 0 });
    s.addText(n[1], { x: M + 2.4, y, w: CW - 2.4, h: 0.42, valign: "middle", fontFace: F, fontSize: 9.5, color: INK, margin: 0 });
  });
}

/* ═══════════════ 25. Appendix F ═══════════════ */
{
  const s = content("Appendix F", "ファイルアップロード仕様の確認結果", "検討段階で挙がっていた「1ファイル5MB制限」は、無料プラン限定の仕様でした", NAVY_MID);
  const rows = [
    [hdr("対象", { align: "left" }), hdr("1ファイルあたりの上限"), hdr("補足")],
    [tx("Free プラン", { bold: true }), tx("5 MB", { align: "center", bold: true, color: RED, fontSize: 12 }), tx("当初「5MB制限」として認識していたのは、この無料プランの仕様です。")],
    [tx("有料プラン（Plus / Business / Enterprise）", { bold: true }), tx("5 GB", { align: "center", bold: true, color: GREEN, fontSize: 12 }), tx("一般的なファイル（動画・zip等）の上限。無料プランの1,000倍にあたります。")],
    [tx("　└ PDF ファイル", {}), tx("20 MB 未満", { align: "center" }), tx("有料プランでも、PDFにはこの個別上限が適用されます。")],
    [tx("　└ 画像（PNG / JPG）", {}), tx("5 MB 未満", { align: "center" }), tx("有料プランでも、画像にはこの個別上限が適用されます。")],
    [tx("ワークスペース全体の総容量", { bold: true }), tx("制限なし", { align: "center", bold: true, color: GREEN }), tx("個々のファイルが上限内であれば、保存総量の制限はありません。")],
  ];
  table(s, rows, 2.44, [4.0, 2.6, 5.53], { rowH: 0.5, fontSize: 10 });

  card(s, M, 5.6, 7.55, 1.35, ICE_PALE, ICE);
  s.addText("この確認が必要だった理由", { x: M + 0.25, y: 5.73, w: 7.05, h: 0.28, fontFace: F, fontSize: 11.5, bold: true, color: NAVY, margin: 0 });
  s.addText("検討段階では「1ファイル5MB制限があり動画は不可」という前提で議論していました。誤った数値のままTHDにご説明すると資料全体の信頼を損なうため、Notion公式ヘルプにて確認し、本表の値に訂正しています。", {
    x: M + 0.25, y: 6.04, w: 7.05, h: 0.8, fontFace: F, fontSize: 10, color: INK, margin: 0, lineSpacing: 16,
  });

  card(s, M + 7.78, 5.6, 4.35, 1.35, WHITE, LINE);
  s.addText("なお、スコープには含めません", { x: M + 8.0, y: 5.73, w: 3.9, h: 0.28, fontFace: F, fontSize: 11.5, bold: true, color: NAVY, margin: 0 });
  s.addText("この結果により③案件関連ファイルの前提は変わりますが、今回のスコープは①案件管理・②議事録に限定しており、ファイル保管は SharePoint を継続します。", {
    x: M + 8.0, y: 6.04, w: 3.9, h: 0.8, fontFace: F, fontSize: 9.5, color: INK, margin: 0, lineSpacing: 15,
  });
}

/* ═══════════════ 26. Appendix G ═══════════════ */
{
  const s = content("Appendix G", "役割分担と残タスク", "本件の推進体制と、THDご説明後に残る作業です", NAVY_MID);
  const roles = [
    ["紙谷", "資料作成・設計", ["本ストーリーに基づく説明資料の作成", "Phase2 対象範囲の設計", "会議設計テンプレートの作成"]],
    ["仲川", "説明・裁定", ["実物デモの準備と実施", "THDへの説明", "プラン・コストの裁定"]],
  ];
  roles.forEach((r, i) => {
    const w = CW / 2 - 0.12;
    const x = M + i * (w + 0.24);
    card(s, x, 2.5, w, 1.85, WHITE, LINE);
    s.addText(r[0], { x: x + 0.25, y: 2.65, w: 2.5, h: 0.36, valign: "middle", fontFace: F, fontSize: 17, bold: true, color: NAVY, margin: 0 });
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 2.0, y: 2.72, w: 2.0, h: 0.26, rectRadius: 0.04,
      fill: { color: ICE }, line: { color: ICE },
    });
    s.addText(r[1], { x: x + 2.0, y: 2.72, w: 2.0, h: 0.26, align: "center", valign: "middle", fontFace: F, fontSize: 9, bold: true, color: NAVY, margin: 0 });
    s.addText(
      r[2].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j !== r[2].length - 1 } })),
      { x: x + 0.25, y: 3.12, w: w - 0.5, h: 1.1, fontFace: F, fontSize: 11, color: INK, margin: 0, paraSpaceAfter: 7 }
    );
  });

  s.addText("THDご説明後の残タスク", { x: M, y: 4.55, w: 5, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: MUTED, margin: 0 });
  const tasks = [
    ["1", "監査ログ要否の確定", "THDのご回答を受けて、Business / Enterprise を最終決定する"],
    ["2", "正規代理店からの見積取得", "Business単価（$18〜$20）を確定し、コスト上限を確定値に更新する"],
    ["3", "セキュリティチェックシートへの回答", "THD所定の様式がある場合、受領のうえ回答を作成する"],
    ["4", "テンプレート・ミニマニュアルの整備", "Phase2展開の前提となる。会議設計テンプレートを含む（Appendix I）"],
    ["5", "ゲスト運用ルールの策定", "業務委託・社外パートナーの招待範囲と手続きを定める"],
  ];
  tasks.forEach((t, i) => {
    const y = 4.9 + i * 0.42;
    s.addShape(pres.ShapeType.roundRect, {
      x: M, y: y + 0.05, w: 0.3, h: 0.3, rectRadius: 0.05,
      fill: { color: NAVY }, line: { color: NAVY },
    });
    s.addText(t[0], { x: M, y: y + 0.05, w: 0.3, h: 0.3, align: "center", valign: "middle", fontFace: F, fontSize: 9, bold: true, color: WHITE, margin: 0 });
    s.addText(t[1], { x: M + 0.42, y, w: 4.1, h: 0.4, valign: "middle", fontFace: F, fontSize: 11, bold: true, color: NAVY, margin: 0 });
    s.addText(t[2], { x: M + 4.6, y, w: CW - 4.6, h: 0.4, valign: "middle", fontFace: F, fontSize: 10, color: INK, margin: 0 });
  });
}

/* ═══════════════ 27. Appendix H ═══════════════ */
{
  const s = content("Appendix H", "出典一覧", "本資料に記載した仕様・プラン要件は、以下の公式情報で確認しています", NAVY_MID);
  const srcs = [
    ["ファイルアップロード上限", "Notion Help Center — Images, files & media", "www.notion.com/help/images-files-and-media", "Free 5MB / 有料 5GB、PDF 20MB・画像 5MB の個別上限"],
    ["SAML SSO のプラン要件", "Notion Help Center — SAML SSO configuration", "www.notion.com/help/saml-sso-configuration", "Business プラン以上で利用可能。IdP は SAML 2.0 対応が必要"],
    ["監査ログのプラン要件", "Notion Help Center — Workspace audit log", "www.notion.com/help/audit-log", "Enterprise プランの組織オーナーのみ利用可能"],
    ["SCIM のプラン要件", "Notion Help Center — Provision users & groups with SCIM", "www.notion.com/help/provision-users-and-groups-with-scim", "Enterprise プランでのみ利用可能"],
    ["セキュリティ認証の取得状況", "Notion Trust Center", "trust.notion.com", "SOC 2 Type II、ISO 27001/27017/27018/27701、BSI C5 等"],
    ["プラン別の機能・価格", "Notion Pricing", "www.notion.com/pricing", "単価は改定される場合があるため、契約時に見積で確定する"],
  ];
  srcs.forEach((sc, i) => {
    const y = 2.5 + i * 0.72;
    card(s, M, y, CW, 0.62, i % 2 === 0 ? WHITE : "FAFCFF", LINE);
    s.addText(sc[0], { x: M + 0.25, y, w: 3.3, h: 0.62, valign: "middle", fontFace: F, fontSize: 11, bold: true, color: NAVY, margin: 0 });
    s.addText(sc[1], { x: M + 3.7, y: y + 0.04, w: 5.05, h: 0.28, valign: "middle", fontFace: F, fontSize: 10.5, color: INK, margin: 0 });
    s.addText(sc[2], { x: M + 3.7, y: y + 0.32, w: 5.05, h: 0.26, valign: "middle", fontFace: F, fontSize: 9, color: NAVY_MID, margin: 0 });
    s.addText(sc[3], { x: M + 8.95, y, w: CW - 9.2, h: 0.62, valign: "middle", fontFace: F, fontSize: 9, color: INK, margin: 0 });
  });
  s.addText("※ 記載内容は2026年8月時点で確認したものです。プラン仕様・価格は改定される場合があるため、契約前に再確認します。", {
    x: M, y: 6.9, w: CW, h: 0.3, fontFace: F, fontSize: 9, color: MUTED, margin: 0,
  });
}

/* ═══════════════ 28. Appendix I ═══════════════ */
{
  const s = content("Appendix I", "会議設計テンプレートと全社展開の権限設計", "第7章 7-1 でご説明した運用を、実際に成立させるための設計です", NAVY_MID);

  s.addText("① 会議設計テンプレート（議事録DBのページテンプレート）", {
    x: M, y: 2.42, w: 6.2, h: 0.3, fontFace: F, fontSize: 11.5, bold: true, color: NAVY, margin: 0,
  });
  const rows = [
    [hdr("項目", { align: "left" }), hdr("記入時点"), hdr("目的", { align: "left" })],
    [tx("会議のゴール（この会議で決めること）", { bold: true }), tx("会議前", { align: "center", bold: true, color: NAVY }), tx("目的が曖昧な会議をなくす")],
    [tx("アジェンダ（時間配分つき）", { bold: true }), tx("会議前", { align: "center", bold: true, color: NAVY }), tx("参加者が準備して臨める")],
    [tx("事前共有資料へのリンク", { bold: true }), tx("会議前", { align: "center", bold: true, color: NAVY }), tx("「資料どこ？」をなくす")],
    [tx("決定事項", { bold: true }), tx("会議中", { align: "center" }), tx("後から検索・追跡できる形にする")],
    [tx("宿題・ToDo（担当者・期限）", { bold: true }), tx("会議中", { align: "center" }), tx("タスクDBへ連携する")],
    [tx("全社展開の要否", { bold: true }), tx("会議後", { align: "center" }), tx("展開漏れを防ぐフラグ")],
  ];
  s.addTable(rows, {
    x: M, y: 2.76, w: 6.2, colW: [2.9, 0.9, 2.4],
    border: { type: "solid", pt: 0.5, color: LINE },
    fontFace: F, fontSize: 8.5, color: INK, valign: "middle", rowH: 0.36, autoPage: false,
  });

  s.addText("② 全社展開のページ構成（同期ブロックの権限制約を踏まえた設計）", {
    x: M + 6.53, y: 2.42, w: 5.6, h: 0.3, fontFace: F, fontSize: 11.5, bold: true, color: NAVY, margin: 0,
  });
  card(s, M + 6.53, 2.76, 5.6, 1.95, GREEN_PALE, GREEN);
  s.addText("正しい設計", { x: M + 6.75, y: 2.86, w: 3, h: 0.26, fontFace: F, fontSize: 10.5, bold: true, color: GREEN, margin: 0 });
  s.addText(
    "全社共有ページ（全社員=閲覧可）　← 原本をここに置く\n" +
    "　├─ 同期先 → 営業部ページ\n" +
    "　├─ 同期先 → 管理部ページ\n" +
    "　└─ 同期先 → 各部門ページ\n" +
    "経営会議 議事録（役員=限定公開）\n" +
    "　└─ 全社共有ページへの「リンク」を置く",
    { x: M + 6.75, y: 3.14, w: 5.2, h: 1.5, fontFace: F, fontSize: 9, color: INK, margin: 0, lineSpacing: 14 }
  );

  card(s, M + 6.53, 4.85, 5.6, 1.6, RED_PALE, RED);
  s.addText("誤った設計 — 一般社員に表示されません", { x: M + 6.75, y: 4.95, w: 5.2, h: 0.26, fontFace: F, fontSize: 10.5, bold: true, color: RED, margin: 0 });
  s.addText(
    "経営会議 議事録（役員限定）に原本を置く\n" +
    "　└─ 全社ページに同期先を貼る\n\n" +
    "同期元ページの閲覧権限が適用されるため、\n閲覧権限のない一般社員には中身が見えません。",
    { x: M + 6.75, y: 5.23, w: 5.2, h: 1.12, fontFace: F, fontSize: 9, color: INK, margin: 0, lineSpacing: 14 }
  );

  card(s, M, 5.4, 6.2, 1.5, ICE_PALE, ICE);
  s.addText("運用ルールとして明文化する2点", { x: M + 0.22, y: 5.5, w: 5.7, h: 0.26, fontFace: F, fontSize: 10.5, bold: true, color: NAVY, margin: 0 });
  s.addText(
    [
      { text: "経営会議の生議事録と、全社に出す展開用の要約は、必ず別ページに分ける", options: { bullet: { type: "number" }, breakLine: true } },
      { text: "同期ブロックの原本は、常に閲覧範囲が最も広いページ側に置く", options: { bullet: { type: "number" } } },
    ],
    { x: M + 0.22, y: 5.8, w: 5.7, h: 1.0, fontFace: F, fontSize: 9.5, color: INK, margin: 0, paraSpaceAfter: 6 }
  );
  s.addNotes("この権限制約は実際に運用を始めてから発覚すると手戻りが大きい。テンプレート整備の段階で構成を固定しておく。");
}

pres.writeFile({ fileName: "notion-thd-proposal.pptx" }).then(() => {
  console.log("done: notion-thd-proposal.pptx / slides =", pageNo + 1);
});
