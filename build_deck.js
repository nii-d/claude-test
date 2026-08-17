/**
 * Notion導入 THD説明資料 — スライド生成スクリプト（改訂版）
 * 表紙・目次 + 本編18枚 + Appendix扉 + Appendix 9枚 = 計30枚
 * ※Appendix ごと THD に配布する前提で記述する（相手の判断を予測する表現は置かない）
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
const NAV_OFF = "E8ECF3";
const RED = "B3261E";
const RED_PALE = "FBEDEC";
const AMBER = "9A6700";
const AMBER_PALE = "FDF6E7";
const GREEN = "1E7A5A";
const GREEN_PALE = "E9F5F0";

const F = "Meiryo";
const W = 13.333, H = 7.5, M = 0.6, CW = W - M * 2;

const MAIN_NAV = ["序", "1", "2", "3", "4", "5", "6", "7", "8"];
const APX_NAV = ["A", "B", "C", "D", "E", "F", "G", "H"];

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

/** 現在地インジケータ（右上・小さなチップの列。装飾の帯にはしない） */
function nav(s, set, current) {
  const items = set === "apx" ? APX_NAV : MAIN_NAV;
  const cw = 0.26, gap = 0.055;
  const total = items.length * cw + (items.length - 1) * gap;
  const x0 = W - M - total;
  items.forEach((it, i) => {
    const on = it === current;
    const x = x0 + i * (cw + gap);
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 0.44, w: cw, h: 0.26, rectRadius: 0.04,
      fill: { color: on ? NAVY : NAV_OFF }, line: { color: on ? NAVY : NAV_OFF },
    });
    s.addText(it, {
      x, y: 0.44, w: cw, h: 0.26, align: "center", valign: "middle",
      fontFace: F, fontSize: 7.5, bold: on, color: on ? WHITE : MUTED, margin: 0,
    });
  });
}

/** 標準コンテンツスライド: 章チップ + 現在地 + タイトル + キーメッセージカード */
function content(chip, title, keyMsg, navSet, navCur, chipColor) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 0.40, w: 1.5, h: 0.34, rectRadius: 0.06,
    fill: { color: chipColor || NAVY }, line: { color: chipColor || NAVY },
  });
  s.addText(chip, {
    x: M, y: 0.40, w: 1.5, h: 0.34, align: "center", valign: "middle",
    fontFace: F, fontSize: 10.5, bold: true, color: WHITE, margin: 0,
  });
  nav(s, navSet, navCur);
  s.addText(title, {
    x: M, y: 0.84, w: CW, h: 0.62, valign: "middle",
    fontFace: F, fontSize: 26, bold: true, color: NAVY, margin: 0,
  });
  if (keyMsg) {
    s.addShape(pres.ShapeType.roundRect, {
      x: M, y: 1.56, w: CW, h: 0.72, rectRadius: 0.08,
      fill: { color: ICE_PALE }, line: { color: ICE },
    });
    s.addText(keyMsg, {
      x: M + 0.22, y: 1.56, w: CW - 0.44, h: 0.72, valign: "middle",
      fontFace: F, fontSize: 13.5, bold: true, color: NAVY, margin: 0,
    });
  }
  footer(s);
  return s;
}

function card(s, x, y, w, h, fill, lineC) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: fill }, line: { color: lineC || LINE }, shadow: sh(),
  });
}

function mk(mark, note) {
  const c = mark === "○" || mark === "◎" ? GREEN : mark === "△" ? AMBER : mark === "×" ? RED : INK;
  const rt = [{ text: mark, options: { color: c, bold: true, fontSize: 12 } }];
  if (note) rt.push({ text: "\n" + note, options: { color: INK, fontSize: 9 } });
  return { text: rt };
}
const tx = (t, o) => ({ text: t, options: Object.assign({ color: INK, fontSize: 10 }, o || {}) });

function table(s, rows, y, colW, opts) {
  s.addTable(rows, Object.assign({
    x: M, y, w: CW, colW,
    border: { type: "solid", pt: 0.5, color: LINE },
    fontFace: F, fontSize: 10, color: INK, valign: "middle",
    align: "left", autoPage: false,
  }, opts || {}));
}
function hdr(t, o) {
  return {
    text: t,
    options: Object.assign({ fill: { color: NAVY }, color: WHITE, bold: true, fontSize: 10.5, align: "center", valign: "middle" }, o || {}),
  };
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
      { text: "対象人数　140名", options: { breakLine: true, fontSize: 13, color: WHITE, bold: true } },
      { text: "契約プラン　Notion Business（SAML SSO 対応）", options: { fontSize: 13, color: WHITE, bold: true } },
    ],
    { x: 0.9, y: 4.25, w: 7.5, h: 1.3, fontFace: F, margin: 0, lineSpacing: 24 }
  );
  s.addText("株式会社PCA　2026年8月", {
    x: 0.9, y: 6.5, w: 6, h: 0.35, fontFace: F, fontSize: 12, color: MUTED, margin: 0,
  });
  s.addNotes("ご説明には実物のデモを併用する。ご承認いただきたいのは、①②に限定した正式導入・Businessプラン契約・コスト上限の3点。");
}

/* ═══════════════ 2. 目次 ═══════════════ */
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText("本日のご説明", {
    x: M, y: 0.55, w: CW, h: 0.7, valign: "middle",
    fontFace: F, fontSize: 30, bold: true, color: NAVY, margin: 0,
  });
  s.addText("背景 → 課題 → 要件 → 選定 → 試行 → セキュリティ → 計画 → お願い の順にご説明します", {
    x: M, y: 1.25, w: CW, h: 0.4, fontFace: F, fontSize: 13, color: MUTED, margin: 0,
  });

  const items = [
    ["序", "30秒サマリー", "本日のご依頼事項"],
    ["1", "背景", "PCAは再建の実行フェーズにある"],
    ["2", "課題", "2つの意味で、案件を管理する仕組みがない"],
    ["3", "要件", "誰が・いつまでに・何のためのタスクかが分かること"],
    ["4", "選定", "4案を比較し、Notionのみが要件を充足"],
    ["5", "試行と運用イメージ", "経営層で3か月試した結果と、実際の使い方"],
    ["6", "セキュリティとプラン選定", "認証・SSO・Businessを選ぶ理由"],
    ["7", "導入計画とコスト", "対象範囲・コスト上限・続ける判断"],
    ["8", "お願い事項", "ご承認いただきたいこと／ご確認事項"],
  ];
  const cw = 5.95, ch = 0.52, gap = 0.13;
  items.forEach((it, i) => {
    const col = i < 5 ? 0 : 1;
    const row = i < 5 ? i : i - 5;
    const x = M + col * (cw + 0.24);
    const y = 1.85 + row * (ch + gap);
    card(s, x, y, cw, ch, WHITE, LINE);
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
        { text: it[1] + "　", options: { bold: true, color: NAVY, fontSize: 11.5 } },
        { text: it[2], options: { color: MUTED, fontSize: 9.5 } },
      ],
      { x: x + 0.68, y, w: cw - 0.8, h: ch, valign: "middle", fontFace: F, margin: 0 }
    );
  });

  card(s, M, 5.45, CW, 1.35, ICE_PALE, ICE);
  s.addText("「案件を管理する仕組みがない」には、2つの意味があります", {
    x: M + 0.25, y: 5.58, w: CW - 0.5, h: 0.32, fontFace: F, fontSize: 12.5, bold: true, color: NAVY, margin: 0,
  });
  s.addText(
    [
      { text: "① プロジェクト全体を統括する仕組みがない", options: { bold: true, color: NAVY, fontSize: 11, breakLine: true } },
      { text: "11月にPMOを新設しますが、部署を越えて進捗と決定を束ねる道具がありません。", options: { color: INK, fontSize: 10 } },
    ],
    { x: M + 0.25, y: 5.95, w: 5.7, h: 0.72, fontFace: F, margin: 0 }
  );
  s.addText(
    [
      { text: "② 案件と議事録が紐づかず、探せない", options: { bold: true, color: NAVY, fontSize: 11, breakLine: true } },
      { text: "決めたことがどこに残ったか分からず、探すところから仕事が始まっています。", options: { color: INK, fontSize: 10 } },
    ],
    { x: M + 6.35, y: 5.95, w: 5.7, h: 0.72, fontFace: F, margin: 0 }
  );
  footer(s);
}

/* ═══════════════ 3. 序 30秒サマリー ═══════════════ */
{
  const s = content("序", "30秒サマリー", "決めたことをやり切る「実行管理の仕組み」を、PCAに新しく立ち上げます", "main", "序");
  const items = [
    ["1", "課題①", "プロジェクト全体を統括する仕組みが無い", "11月にPMOを新設しますが、部署を越えて束ねる道具がありません", RED],
    ["2", "課題②", "案件と議事録が紐づかず、探したいときに探せない", "議事録は OneNote・Word・個人メモに散在しています", RED],
    ["3", "対応", "①案件管理 ②議事録 に限定して新規立ち上げ", "③ファイル・④社内規定は SharePoint 継続。Microsoft 365 は置き換えません", NAVY],
    ["4", "根拠", "経営層で3か月試し、有効性を確認しました", "週次会議の管理台帳として、そのまま使えることを確認しています", NAVY],
    ["5", "安全性", "Business プランの SAML SSO で Microsoft 365 と統合", "当社の既存セキュリティポリシーがそのまま Notion にも適用されます", NAVY],
  ];
  const rh = 0.6, gap = 0.13;
  items.forEach((it, i) => {
    const y = 2.46 + i * (rh + gap);
    card(s, M, y, CW - 3.5, rh, i < 2 ? RED_PALE : WHITE, i < 2 ? "F0CFCC" : LINE);
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 0.15, y: y + 0.14, w: 0.32, h: 0.32, rectRadius: 0.16,
      fill: { color: it[4] }, line: { color: it[4] },
    });
    s.addText(it[0], {
      x: M + 0.15, y: y + 0.14, w: 0.32, h: 0.32, align: "center", valign: "middle",
      fontFace: F, fontSize: 10, bold: true, color: WHITE, margin: 0,
    });
    s.addText(it[1], {
      x: M + 0.6, y: y + 0.04, w: 0.85, h: 0.26, valign: "middle",
      fontFace: F, fontSize: 9.5, bold: true, color: MUTED, margin: 0,
    });
    s.addText(it[2], {
      x: M + 1.48, y: y + 0.02, w: 6.8, h: 0.28, valign: "middle",
      fontFace: F, fontSize: 11.5, bold: true, color: NAVY, margin: 0,
    });
    s.addText(it[3], {
      x: M + 0.6, y: y + 0.32, w: 7.7, h: 0.26, valign: "middle",
      fontFace: F, fontSize: 9, color: INK, margin: 0,
    });
  });

  card(s, M + CW - 3.3, 2.46, 3.3, 3.4, NAVY, NAVY);
  s.addText("本日のお願い", {
    x: M + CW - 3.1, y: 2.64, w: 2.9, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: ICE, margin: 0,
  });
  s.addText("140名", {
    x: M + CW - 3.1, y: 3.0, w: 2.9, h: 0.72, fontFace: F, fontSize: 40, bold: true, color: WHITE, margin: 0,
  });
  s.addText("対象人数", {
    x: M + CW - 3.1, y: 3.7, w: 2.9, h: 0.26, fontFace: F, fontSize: 9.5, color: ICE, margin: 0,
  });
  s.addText("年額 約520.8万円", {
    x: M + CW - 3.1, y: 4.14, w: 2.9, h: 0.42, fontFace: F, fontSize: 20, bold: true, color: WHITE, margin: 0,
  });
  s.addText("Business プラン・年払\nこの金額を上限としてご承認ください", {
    x: M + CW - 3.1, y: 4.6, w: 2.9, h: 0.6, fontFace: F, fontSize: 9.5, color: ICE, margin: 0,
  });
  s.addText("1人あたり 1営業日 約169円", {
    x: M + CW - 3.1, y: 5.35, w: 2.9, h: 0.3, valign: "middle",
    fontFace: F, fontSize: 10, bold: true, color: ICE, italic: true, margin: 0,
  });
  s.addNotes("冒頭30秒で全体像を示す。課題は2つ（PMO統括の不在／情報が探せない）ある点を先に共有する。金額は上限としての提示。");
}

/* ═══════════════ 4. 第1章 背景 ═══════════════ */
{
  const s = content("第1章", "背景 — PCAは再建の実行フェーズにある", "再建の成否は、戦略の中身以上に「やり切る」実行管理で決まります", "main", "1");
  const blocks = [
    ["状況", NAVY, ICE_PALE, ICE, [
      "中期経営計画 FY27-29 を実行中",
      "10月に新体制を発表し、11月に組織変更",
      "部門をまたぐ変革案件が同時多発",
    ]],
    ["帰結", RED, RED_PALE, "F0CFCC", [
      "決めたことの実行漏れが最大のリスク",
      "実行管理の巧拙が、そのまま業績に直結する",
    ]],
  ];
  blocks.forEach((b, i) => {
    const x = M + i * (CW / 2 + 0.15);
    const w = CW / 2 - 0.15;
    card(s, x, 2.55, w, 2.05, b[2], b[3]);
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.25, y: 2.75, w: 0.95, h: 0.34, rectRadius: 0.05,
      fill: { color: b[1] }, line: { color: b[1] },
    });
    s.addText(b[0], {
      x: x + 0.25, y: 2.75, w: 0.95, h: 0.34, align: "center", valign: "middle",
      fontFace: F, fontSize: 11, bold: true, color: WHITE, margin: 0,
    });
    s.addText(
      b[4].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j !== b[4].length - 1 } })),
      { x: x + 0.25, y: 3.22, w: w - 0.5, h: 1.2, fontFace: F, fontSize: 12, color: INK, margin: 0, paraSpaceAfter: 7 }
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
  s.addNotes("この章では課題に踏み込まず、再建フェーズという前提を共有する。組織変更の中身は次章で扱う。");
}

/* ═══════════════ 5. 第2章 2-1 課題①（新規） ═══════════════ */
{
  const s = content("第2章", "課題① — プロジェクト全体を統括する仕組みがない", "11月にPMOを新設しますが、PMOが部署を越えて働くための仕組みがありません", "main", "2");

  const before = ["営業", "運営", "講師管理", "講師"];
  const after = [
    ["営業", false], ["PMO", true], ["コンテンツ制作", true], ["講師管理", false],
    ["講師", false], ["運営", false], ["システム", true],
  ];

  // 左：これまで
  card(s, M, 2.44, 5.3, 2.85, WHITE, LINE);
  s.addText("これまで（4部門）", {
    x: M + 0.25, y: 2.56, w: 4.8, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: MUTED, margin: 0,
  });
  before.forEach((b, i) => {
    const y = 2.92 + i * 0.36;
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 0.25, y, w: 4.8, h: 0.3, rectRadius: 0.04,
      fill: { color: "F4F6FA" }, line: { color: LINE },
    });
    s.addText(b, {
      x: M + 0.45, y, w: 4.4, h: 0.3, valign: "middle",
      fontFace: F, fontSize: 11, color: INK, margin: 0,
    });
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M + 0.25, y: 4.45, w: 4.8, h: 0.68, rectRadius: 0.05,
    fill: { color: RED_PALE }, line: { color: RED },
  });
  s.addText("研修案件（プロジェクト）ごとに\n全体を見る立場が不在", {
    x: M + 0.4, y: 4.45, w: 4.5, h: 0.68, valign: "middle",
    fontFace: F, fontSize: 10.5, bold: true, color: RED, margin: 0, lineSpacing: 15,
  });

  // 矢印
  s.addText("▶", {
    x: M + 5.42, y: 3.5, w: 0.7, h: 0.4, align: "center", valign: "middle",
    fontFace: F, fontSize: 18, color: NAVY_MID, margin: 0,
  });

  // 右：今後
  card(s, M + 6.23, 2.44, 5.9, 2.85, ICE_PALE, ICE);
  s.addText("今後（7部門・11月〜）", {
    x: M + 6.48, y: 2.56, w: 5.4, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: NAVY, margin: 0,
  });
  after.forEach((a, i) => {
    const y = 2.92 + i * 0.33;
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 6.48, y, w: 5.4, h: 0.28, rectRadius: 0.04,
      fill: { color: a[1] ? NAVY : WHITE }, line: { color: a[1] ? NAVY : LINE },
    });
    s.addText(a[0], {
      x: M + 6.68, y, w: 3.4, h: 0.28, valign: "middle",
      fontFace: F, fontSize: 10.5, bold: a[1], color: a[1] ? WHITE : INK, margin: 0,
    });
    if (a[1]) {
      s.addText("新設", {
        x: M + 10.5, y, w: 1.2, h: 0.28, align: "right", valign: "middle",
        fontFace: F, fontSize: 8.5, bold: true, color: ICE, margin: 0,
      });
    }
  });

  card(s, M, 5.42, CW, 0.9, NAVY, NAVY);
  s.addText("これまで PCA には、研修案件ごとに全体を見る立場がありませんでした。11月にその役割として PMO を置きますが、\n部署を越えて進捗・決定・課題を束ねるための仕組みは、現在どこにもありません。", {
    x: M + 0.35, y: 5.42, w: CW - 0.7, h: 0.9, valign: "middle",
    fontFace: F, fontSize: 12.5, bold: true, color: WHITE, margin: 0, lineSpacing: 22,
  });

  s.addText("PMO は営業・コンテンツ制作・講師管理・運営・システムの5部門にまたがって1つの案件を見る立場になります。新設される役割ほど、拠り所となる仕組みが必要です。", {
    x: M, y: 6.44, w: CW, h: 0.3, fontFace: F, fontSize: 10, color: INK, margin: 0,
  });
  s.addNotes("組織変更の正確な実施日は確定次第差し替える。ここでは11月という月のみ記載している。");
}

/* ═══════════════ 6. 第2章 2-2 課題② ═══════════════ */
{
  const s = content("第2章", "課題② — 案件と議事録が紐づかず、探せない", "決めたことが、どこに残ったか分からなくなります", "main", "2");
  card(s, M, 2.5, 5.9, 1.9, WHITE, LINE);
  s.addText("現状", { x: M + 0.25, y: 2.65, w: 2, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: NAVY, margin: 0 });
  s.addText(
    [
      { text: "案件と議事録を結びつける仕組みが無い", options: { bullet: true, breakLine: true } },
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
  s.addText("「あの件どうなった？」「あの資料どこ？」を探すところから仕事が始まっています。\n情報が無いのではなく、たどり着けない状態です。", {
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
}

/* ═══════════════ 7. 第3章 要件 ═══════════════ */
{
  const s = content("第3章", "要件 — 必要なのは「つながる」管理基盤", "誰が・いつまでに・何のためのタスクなのかが、常に分かる状態が必要です", "main", "3");
  s.addText("機能要件", { x: M, y: 2.45, w: 3, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: MUTED, margin: 0 });
  const fn = [
    ["タスクに「誰が・いつまでに」が必ず付く", "担当者と期限が空欄のまま放置されない。誰の手元で止まっているかが常に分かる"],
    ["遅れが自動で浮かび上がる", "期限を過ぎたタスクを部署をまたいで一覧できる。PMOは「遅れているものだけ」を見て動ける"],
    ["そのタスクが何のためのものか辿れる", "タスクから案件へ、案件から決定事項・議事録へ遡れる"],
    ["立場ごとに見え方を変えられる", "PMOは全案件を横断、部門は自分の担当分、経営は全体を、同じデータから見る"],
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
    s.addText(f[0], { x: x + 0.62, y: y + 0.08, w: w - 0.8, h: 0.3, valign: "middle", fontFace: F, fontSize: 12, bold: true, color: NAVY, margin: 0 });
    s.addText(f[1], { x: x + 0.62, y: y + 0.38, w: w - 0.8, h: 0.38, fontFace: F, fontSize: 9, color: INK, margin: 0 });
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
  s.addNotes("要件を先に固定する章のため、ここでは製品名を出さない。");
}

/* ═══════════════ 8. 第4章 4-1 比較表 ═══════════════ */
{
  const s = content("第4章", "選定 — 要件を固定したうえで、4案を比較する", "要件を先に決めてから比較すると、1製品で満たせるのは Notion だけでした", "main", "4");
  const rows = [
    [hdr("要件", { align: "left" }), hdr("A. Microsoft 365 の組合せ\n(Planner/Lists/Loop)"), hdr("B. 無料ツール"), hdr("C. 専用ツール\n(Asana/Backlog等)"), hdr("D. Notion")],
    [tx("1. タスクに担当者・期限が必ず付く", { bold: true }), mk("△", "ツールごとに運用がばらつく"), mk("△", "必須化できない"), mk("○", "タスク管理は得意"), mk("○", "必須プロパティで担保")],
    [tx("2. 遅れを部署横断で一覧できる", { bold: true }), mk("×", "ツールをまたぐと不可"), mk("△", "限定的"), mk("△", "ツール内のみ"), mk("○", "横断ビューで抽出")],
    [tx("3. 何のためのタスクか辿れる", { bold: true }), mk("△", "案件と議事録が別管理"), mk("△", "機能不足"), mk("△", "親子課題までが限界"), mk("○", "案件・議事録まで遡れる")],
    [tx("4. 立場別のビューを切り替えられる", { bold: true }), mk("×", "不可"), mk("△", "限定的"), mk("○", "対応"), mk("○", "同一データで切替")],
    [tx("5. ガバナンス（SSO・権限）", { bold: true }), mk("○", "Microsoft 365 の統制下"), mk("×", "統制不可"), mk("○", "対応"), mk("○", "SSO=Business以上（第6章）")],
    [tx("6. 1製品で完結（教育コスト最小化）", { bold: true }), mk("×", "3製品以上の併用"), mk("×", "用途ごとに別ツール"), mk("×", "議事録・wikiは別"), mk("○", "1製品")],
    [tx("（参考）当社での試用", { bold: true }), tx("—", { align: "center" }), tx("—", { align: "center" }), mk("△", "Backlogを試用（4-2）"), mk("○", "経営層で3か月試用")],
  ];
  table(s, rows, 2.42, [3.5, 2.16, 2.16, 2.16, 2.15], { rowH: [0.42, 0.4, 0.32, 0.4, 0.32, 0.4, 0.32, 0.36], fontSize: 8.5 });

  s.addText("次ページで、A・B・C の3案が外れる理由をそれぞれご説明します。", {
    x: M, y: 6.2, w: CW, h: 0.3, fontFace: F, fontSize: 11.5, bold: true, color: NAVY, margin: 0,
  });
  s.addNotes("要件1〜3は第3章の要件に対応している。表は読み上げず、最下段の「当社での試用」に触れて次ページへつなぐ。");
}

/* ═══════════════ 9. 第4章 4-2 外れる理由（新規） ═══════════════ */
{
  const s = content("第4章", "選定 — 他の3案が外れる理由", "3案が外れる理由は、それぞれ性質が異なります", "main", "4");
  const cases = [
    ["A", "Microsoft 365 の組合せ", NAVY,
      "Microsoft 365 で解決できるのであれば、この数年のあいだにすでに解決できていたはずです。複数のアプリを組み合わせる形は、情報が散らばる状態を作り直すことになり、いま解決したい課題そのものになってしまいます。"],
    ["B", "無料ツール", NAVY,
      "部門ごとに無料ツールが乱立する状態こそ、会社として把握できないシャドーITのリスクになります。有料の1製品に統制するほうが、ガバナンスはむしろ強くなります。"],
    ["C", "専用ツール（Asana / Backlog 等）", RED,
      "当社での試用結果：システム部門で Backlog を試用しました。ガントチャートでのタスク管理は有効でしたが、親子課題までしか紐づけられず、「そのタスクが何の目的で存在するのか」を管理できませんでした。第3章の要件3を満たしません。"],
  ];
  cases.forEach((c, i) => {
    const y = 2.44 + i * 1.35;
    card(s, M, y, CW, 1.22, i === 2 ? RED_PALE : WHITE, i === 2 ? RED : LINE);
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 0.22, y: y + 0.2, w: 0.42, h: 0.42, rectRadius: 0.06,
      fill: { color: c[2] }, line: { color: c[2] },
    });
    s.addText(c[0], {
      x: M + 0.22, y: y + 0.2, w: 0.42, h: 0.42, align: "center", valign: "middle",
      fontFace: F, fontSize: 13, bold: true, color: WHITE, margin: 0,
    });
    s.addText(c[1], {
      x: M + 0.78, y: y + 0.16, w: CW - 1.0, h: 0.32, valign: "middle",
      fontFace: F, fontSize: 13, bold: true, color: c[2], margin: 0,
    });
    s.addText(c[3], {
      x: M + 0.78, y: y + 0.5, w: CW - 1.05, h: 0.62,
      fontFace: F, fontSize: 10.5, color: INK, margin: 0, lineSpacing: 17,
    });
  });

  card(s, M, 6.5, CW, 0.55, ICE_PALE, ICE);
  s.addText("では、すでに全社導入している Microsoft 365 はどうなるのか。次ページでご説明します。", {
    x: M + 0.3, y: 6.5, w: CW - 0.6, h: 0.55, valign: "middle",
    fontFace: F, fontSize: 12, bold: true, color: NAVY, margin: 0,
  });
  s.addNotes("C案は机上の比較ではなく当社の試用結果である。ガントチャートは有効だったが、目的とタスクの結びつきを管理できなかった点を具体的に補足する。");
}

/* ═══════════════ 10. 第4章 4-3 Microsoft 365 と併用 ═══════════════ */
{
  const s = content("第4章", "選定 — Microsoft 365 とは併用します", "Microsoft 365 は置き換えません。役割が異なるため、それぞれの得意分野で使い分けます", "main", "4");
  s.addText("前ページのとおり Microsoft 365 の組合せでは要件を満たせませんが、これは Microsoft 365 が不要という意味ではありません。", {
    x: M, y: 2.42, w: CW, h: 0.3, fontFace: F, fontSize: 11.5, color: INK, margin: 0,
  });

  s.addText("役割の住み分け", { x: M, y: 2.8, w: 4, h: 0.28, fontFace: F, fontSize: 11.5, bold: true, color: MUTED, margin: 0 });
  const dom = [
    ["Notion", "案件と議事録", ["案件・タスク・議事録の管理", "相互リンクと横断検索", "立場別のビュー"], NAVY, WHITE, ICE],
    ["SharePoint", "ファイルの正本", ["案件関連ファイルの保管", "社内規定・マニュアル", "既存運用をそのまま継続"], WHITE, NAVY, MUTED],
    ["Teams", "伝達・コミュニケーション", ["チャット・会議", "日々のやりとり", "既存運用をそのまま継続"], WHITE, NAVY, MUTED],
  ];
  dom.forEach((d, i) => {
    const w = CW / 3 - 0.14;
    const x = M + i * (w + 0.21);
    card(s, x, 3.14, w, 2.1, d[3], i === 0 ? NAVY : LINE);
    s.addText(d[0], { x: x + 0.25, y: 3.32, w: w - 0.5, h: 0.38, fontFace: F, fontSize: 17, bold: true, color: d[4], margin: 0 });
    s.addText(d[1], { x: x + 0.25, y: 3.72, w: w - 0.5, h: 0.28, fontFace: F, fontSize: 11, bold: true, color: d[5], margin: 0 });
    s.addText(
      d[2].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j !== d[2].length - 1 } })),
      { x: x + 0.25, y: 4.1, w: w - 0.5, h: 1.0, fontFace: F, fontSize: 10, color: i === 0 ? ICE : INK, margin: 0, paraSpaceAfter: 5 }
    );
  });

  card(s, M, 5.45, CW, 0.85, ICE_PALE, ICE);
  s.addText("PCA側のお約束　③案件関連ファイル と ④社内規定・マニュアル は SharePoint を継続します。Microsoft 365 の置き換え・二重投資にはあたりません。", {
    x: M + 0.3, y: 5.45, w: CW - 0.6, h: 0.85, valign: "middle",
    fontFace: F, fontSize: 12, bold: true, color: NAVY, margin: 0,
  });
  s.addText("Microsoft 365 Copilot についてもご質問をいただくことがありますが、Copilot は検索・要約の道具であり、案件管理の仕組み自体は提供しません（Appendix B）。", {
    x: M, y: 6.45, w: CW, h: 0.3, fontFace: F, fontSize: 10, color: MUTED, margin: 0,
  });
}

/* ═══════════════ 11. 第5章 5-1 試行結果 ═══════════════ */
{
  const s = content("第5章", "試行 — 経営層で3か月試した結果", "経営層で3か月試し、週次の管理台帳として使えることを確認しました", "main", "5");
  const feats = [
    ["経営イシューマップ", "7領域・40件超", "経営課題を7つの領域に整理し、案件単位で進捗を管理した"],
    ["タスクDB", "イシューと双方向連携", "イシューを開けば担当タスクが、タスクからはイシューが辿れる"],
    ["議事録DB", "会議体別・カレンダー", "会議体ごとに整理し、カレンダーから議事録に直接到達できる"],
  ];
  feats.forEach((f, i) => {
    const w = 3.72;
    const x = M + i * (w + 0.24);
    card(s, x, 2.44, w, 1.6, WHITE, LINE);
    s.addText(f[0], { x: x + 0.22, y: 2.58, w: w - 0.44, h: 0.32, fontFace: F, fontSize: 13, bold: true, color: NAVY, margin: 0 });
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.22, y: 2.93, w: w - 0.44, h: 0.28, rectRadius: 0.05,
      fill: { color: ICE_PALE }, line: { color: ICE },
    });
    s.addText(f[1], {
      x: x + 0.22, y: 2.93, w: w - 0.44, h: 0.28, align: "center", valign: "middle",
      fontFace: F, fontSize: 9.5, bold: true, color: NAVY, margin: 0,
    });
    s.addText(f[2], { x: x + 0.22, y: 3.28, w: w - 0.44, h: 0.66, fontFace: F, fontSize: 9.5, color: INK, margin: 0 });
  });

  s.addText("3か月試して分かったこと", { x: M, y: 4.2, w: 5, h: 0.3, fontFace: F, fontSize: 11.5, bold: true, color: MUTED, margin: 0 });
  const found = [
    "週次会議の場で、そのまま管理台帳として使えた",
    "カレンダーから議事録へ直接たどり着けるようになった",
    "決定事項が会議単位ではなく案件単位で蓄積された",
  ];
  found.forEach((t, i) => {
    const y = 4.54 + i * 0.5;
    card(s, M, y, 7.5, 0.42, WHITE, LINE);
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 0.18, y: y + 0.06, w: 0.3, h: 0.3, rectRadius: 0.15,
      fill: { color: GREEN }, line: { color: GREEN },
    });
    s.addText("✓", {
      x: M + 0.18, y: y + 0.06, w: 0.3, h: 0.3, align: "center", valign: "middle",
      fontFace: F, fontSize: 10, bold: true, color: WHITE, margin: 0,
    });
    s.addText(t, { x: M + 0.6, y, w: 6.7, h: 0.42, valign: "middle", fontFace: F, fontSize: 11, color: INK, margin: 0 });
  });

  card(s, M + 7.75, 4.54, 4.38, 1.9, AMBER_PALE, AMBER);
  s.addText("試した範囲について", {
    x: M + 7.97, y: 4.66, w: 3.95, h: 0.28, fontFace: F, fontSize: 11, bold: true, color: AMBER, margin: 0,
  });
  s.addText("試したのは経営層です。PMO・各部門での運用はこれからになります。\n\n本資料でご説明する運用イメージ（次ページ）は、この試行を踏まえた設計であり、実績ではありません。", {
    x: M + 7.97, y: 4.96, w: 3.95, h: 1.35, fontFace: F, fontSize: 9.5, color: INK, margin: 0, lineSpacing: 15,
  });

  s.addText("THDご説明の場では、実物のデモをご覧いただきます。", {
    x: M, y: 6.5, w: 7.5, h: 0.3, fontFace: F, fontSize: 11, bold: true, color: NAVY, margin: 0,
  });
  s.addNotes("ここで実物のデモに切り替える。試した範囲は経営層であり、PMO・各部門での運用はこれからである点を明確に伝える。");
}

/* ═══════════════ 12. 第5章 5-2 3つのDB（新規・章の主役） ═══════════════ */
{
  const s = content("第5章", "運用イメージ — 3つのデータベースをつないで使う", "1つの案件を開けば、誰が・いつまでに・何をしていて、どこで何が決まったかが1画面に集まります", "main", "5");

  const dbs = [
    ["案件DB", "1研修案件 ＝ 1レコード", NAVY, WHITE, ICE,
      ["クライアント／案件名", "PMO担当・営業担当", "ステータス", "開始日・納品日", "関連部門"],
      "ガント（全案件の時系列）\nボード（ステータス別）"],
    ["タスクDB", "案件にぶら下がる作業", NAVY, WHITE, ICE,
      ["担当者", "期限", "ステータス・遅延", "所属案件（リレーション）", "担当部門"],
      "遅延中のみ／担当者別\n今週やること"],
    ["議事録DB", "会議の記録", NAVY, WHITE, ICE,
      ["会議体・日付", "参加者", "関連案件（リレーション）", "決定事項", "宿題（タスクへ連携）"],
      "カレンダー／案件別"],
  ];
  dbs.forEach((d, i) => {
    const w = 3.85;
    const x = M + i * (w + 0.29);
    card(s, x, 2.44, w, 1.78, d[2], d[2]);
    s.addText(d[0], { x: x + 0.22, y: 2.52, w: w - 0.44, h: 0.3, fontFace: F, fontSize: 14.5, bold: true, color: d[3], margin: 0 });
    s.addText(d[1], { x: x + 0.22, y: 2.82, w: w - 0.44, h: 0.22, fontFace: F, fontSize: 8.5, color: d[4], margin: 0 });
    s.addText(
      d[5].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j !== d[5].length - 1 } })),
      { x: x + 0.22, y: 3.08, w: w - 0.44, h: 1.06, fontFace: F, fontSize: 8.5, color: d[4], margin: 0, paraSpaceAfter: 1 }
    );
    if (i < 2) {
      s.addText("⇔", {
        x: x + w + 0.01, y: 3.18, w: 0.27, h: 0.3, align: "center", valign: "middle",
        fontFace: F, fontSize: 14, bold: true, color: NAVY_MID, margin: 0,
      });
    }
    // PMOが使うビュー
    card(s, x, 4.58, w, 0.6, ICE_PALE, ICE);
    s.addText(d[6], {
      x: x + 0.18, y: 4.58, w: w - 0.36, h: 0.6, valign: "middle",
      fontFace: F, fontSize: 8.5, color: NAVY, margin: 0, lineSpacing: 12,
    });
  });
  s.addText("PMOが使う主なビュー", {
    x: M, y: 4.28, w: 4, h: 0.24, fontFace: F, fontSize: 9.5, bold: true, color: MUTED, margin: 0,
  });

  card(s, M, 5.36, 7.5, 1.5, WHITE, LINE);
  s.addText("3つがつながると、こうなります", {
    x: M + 0.25, y: 5.46, w: 7.0, h: 0.26, fontFace: F, fontSize: 11, bold: true, color: NAVY, margin: 0,
  });
  s.addText(
    [
      { text: "案件を開く → その案件のタスク一覧・議事録・決定事項が同じ画面に並ぶ", options: { bullet: true, breakLine: true } },
      { text: "タスクを開く → どの案件のどの目的の作業か分かる（Backlogでできなかった点）", options: { bullet: true, breakLine: true } },
      { text: "議事録を開く → どの案件の会議か、宿題がどのタスクになったかを辿れる", options: { bullet: true } },
    ],
    { x: M + 0.25, y: 5.76, w: 7.0, h: 1.0, fontFace: F, fontSize: 9.5, color: INK, margin: 0, paraSpaceAfter: 5 }
  );

  card(s, M + 7.75, 5.36, 4.38, 1.5, NAVY, NAVY);
  s.addText("PMOの一週間", {
    x: M + 7.97, y: 5.46, w: 3.95, h: 0.26, fontFace: F, fontSize: 11, bold: true, color: ICE, margin: 0,
  });
  s.addText(
    [
      { text: "週初：「遅延中」ビューで、部署をまたいで遅れているタスクだけを確認する", options: { bullet: true, breakLine: true } },
      { text: "週次会議：議事録をその場で作り、決定事項をタスクに変える", options: { bullet: true, breakLine: true } },
      { text: "随時：案件ガントで全案件の山谷を見て、要員の偏りを把握する", options: { bullet: true } },
    ],
    { x: M + 7.97, y: 5.76, w: 3.95, h: 1.0, fontFace: F, fontSize: 8.5, color: WHITE, margin: 0, paraSpaceAfter: 4 }
  );
  s.addNotes("デモでは、この3つのデータベースのつながりを実際に辿って見せる。案件→タスク→議事録の順で開くと分かりやすい。");
}

/* ═══════════════ 13. 第5章 5-3 会議設計と全社展開 ═══════════════ */
{
  const s = content("第5章", "運用イメージ — 会議の設計と、決定事項の全社展開", "会議の目的を先に決め、決まったことが確実に現場まで届く状態にします", "main", "5");

  const steps = [
    ["1", "会議前", "ゴールとアジェンダを作って共有", "議事録ページを会議の前に起票し、この会議で決めることを明示する。参加者が準備して臨める。"],
    ["2", "会議中", "その場で議事録化", "決定事項と宿題（担当者・期限）を、会議の場で確定させる。"],
    ["3", "会議後", "タスクと案件に接続", "決定事項がタスクDBに紐づき、案件・イシューにつながる。"],
    ["4", "全社展開", "各MGを通じて現場まで届ける", "全社共有用ページを用意し、同期ブロック（シンクロ）で各部門ページへ配信する。"],
  ];
  const sw = 2.845, sg = 0.25;
  steps.forEach((st, i) => {
    const x = M + i * (sw + sg);
    const hl = i === 3;
    card(s, x, 2.44, sw, 2.0, hl ? ICE_PALE : WHITE, hl ? ICE : LINE);
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.2, y: 2.61, w: 0.36, h: 0.36, rectRadius: 0.18,
      fill: { color: NAVY }, line: { color: NAVY },
    });
    s.addText(st[0], {
      x: x + 0.2, y: 2.61, w: 0.36, h: 0.36, align: "center", valign: "middle",
      fontFace: F, fontSize: 11, bold: true, color: WHITE, margin: 0,
    });
    s.addText(st[1], { x: x + 0.65, y: 2.61, w: sw - 0.85, h: 0.36, valign: "middle", fontFace: F, fontSize: 13, bold: true, color: NAVY, margin: 0 });
    s.addText(st[2], { x: x + 0.2, y: 3.07, w: sw - 0.4, h: 0.56, fontFace: F, fontSize: 11, bold: true, color: INK, margin: 0 });
    s.addText(st[3], { x: x + 0.2, y: 3.65, w: sw - 0.4, h: 0.7, fontFace: F, fontSize: 9, color: MUTED, margin: 0 });
    if (i < 3) {
      s.addText("▶", {
        x: x + sw + 0.01, y: 3.19, w: sg - 0.02, h: 0.3, align: "center", valign: "middle",
        fontFace: F, fontSize: 11, color: NAVY_MID, margin: 0,
      });
    }
  });

  card(s, M, 4.64, 7.55, 1.75, WHITE, LINE);
  s.addText("全社展開のイメージ", { x: M + 0.25, y: 4.77, w: 4, h: 0.28, fontFace: F, fontSize: 11.5, bold: true, color: NAVY, margin: 0 });
  s.addText("例：経営会議で決定した全社方針を、全部門へ展開する必要が生じたとき。全社共有用ページの内容を各部門ページへ同期配信し、原本を1回直せば展開先すべてに即時反映されます。各MGが個別に転記・再作成する必要がなくなり、伝達漏れ・版ズレ・「言った／聞いてない」がなくなります。", {
    x: M + 0.25, y: 5.09, w: 7.05, h: 1.15, fontFace: F, fontSize: 10, color: INK, margin: 0, lineSpacing: 16,
  });

  card(s, M + 7.78, 4.64, 4.35, 1.75, RED_PALE, RED);
  s.addText("設計上の注意（重要）", { x: M + 8.0, y: 4.77, w: 3.9, h: 0.28, fontFace: F, fontSize: 11.5, bold: true, color: RED, margin: 0 });
  s.addText("同期ブロックは「同期元ページの閲覧権限」が適用されます。経営会議の議事録（非公開）に原本を置くと、一般社員には中身が表示されません。原本は必ず全社共有ページ側に置きます（Appendix H）。", {
    x: M + 8.0, y: 5.09, w: 3.9, h: 1.15, fontFace: F, fontSize: 9.5, color: INK, margin: 0, lineSpacing: 15,
  });
  s.addNotes("会議の目的を先に決めて共有し、決まったことが確実に現場まで届く状態を作りたい、という経営の意向を口頭で補足する。");
}

/* ═══════════════ 14. 第6章 6-1 第三者認証 ═══════════════ */
{
  const s = content("第6章", "セキュリティ — 第三者認証の取得状況", "国際的な第三者監査を毎年受けており、認証の面でのご懸念はありません", "main", "6", NAVY_MID);
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
  s.addText("監査報告書・認証書は Notion Trust Center（trust.notion.com）にて公開されています。\n直接ご確認いただけます。必要であれば、当社より取得のうえご提出します。", {
    x: M + 0.3, y: 5.66, w: CW - 0.6, h: 0.7, fontFace: F, fontSize: 11.5, color: INK, margin: 0, lineSpacing: 20,
  });
}

/* ═══════════════ 15. 第6章 6-2 プラン別の機能対応 ═══════════════ */
{
  const s = content("第6章", "セキュリティ — プラン別の機能対応", "統制機能に加え、実務で効く機能もプランによって差があります", "main", "6", NAVY_MID);
  const L = { bold: true, align: "left" };
  const rows = [
    [hdr("機能", { align: "left" }), hdr("Plus"), hdr("Business", { fill: { color: GREEN } }), hdr("Enterprise")],
    [tx("SAML SSO（シングルサインオン）", L), mk("×"), mk("○"), mk("○")],
    [tx("監査ログ（Audit Log）", L), mk("×"), mk("×"), mk("○")],
    [tx("SCIM（自動IDプロビジョニング）", L), mk("×"), mk("×"), mk("○")],
    [tx("外部共有・公開の組織一括禁止", L), mk("×"), mk("×"), mk("○")],
    [tx("プライベートチームスペース", L), mk("×"), mk("○"), mk("○")],
    [tx("Notion AI（追加費用なしで利用）", Object.assign({}, L, { color: NAVY })), tx("×  試用のみ", { align: "center", color: RED }), tx("○  全員が利用可", { align: "center", color: GREEN, bold: true }), tx("○", { align: "center", color: GREEN })],
    [tx("ゲスト上限（外部講師・委託先）", Object.assign({}, L, { color: NAVY })), tx("100名", { align: "center" }), tx("250名", { align: "center", bold: true, color: GREEN }), tx("拡張可", { align: "center" })],
    [tx("高度なページ分析（誰が編集したか）", Object.assign({}, L, { color: NAVY })), mk("×"), mk("○"), mk("○")],
    [tx("データ履歴保持（バージョン復元）", L), tx("30日", { align: "center" }), tx("90日", { align: "center", bold: true }), tx("無制限", { align: "center" })],
  ];
  table(s, rows, 2.42, [5.53, 2.2, 2.2, 2.2], { rowH: 0.33, fontSize: 9.5, align: "center" });

  const boxes = [
    ["SSO が必須要件の場合", "Business プラン以上", GREEN, GREEN_PALE],
    ["監査ログ・SCIM も必須の場合", "Enterprise プラン一択", AMBER, AMBER_PALE],
  ];
  boxes.forEach((b, i) => {
    const x = M + i * (CW / 2 + 0.12);
    const w = CW / 2 - 0.12;
    card(s, x, 6.0, w, 0.8, b[3], b[2]);
    s.addText(b[0], { x: x + 0.25, y: 6.08, w: w - 0.5, h: 0.26, fontFace: F, fontSize: 10, color: INK, margin: 0 });
    s.addText(b[1], { x: x + 0.25, y: 6.32, w: w - 0.5, h: 0.38, fontFace: F, fontSize: 15, bold: true, color: b[2], margin: 0 });
  });

  s.addText("※ Notion AI のバンドル有無とゲスト上限は、契約時に正規代理店へ確認のうえ確定します。", {
    x: M, y: 6.9, w: CW, h: 0.26, fontFace: F, fontSize: 9, color: MUTED, margin: 0,
  });
  s.addNotes("下段3行（Notion AI・ゲスト上限・ページ分析）がセキュリティ以外の差にあたる。詳細は次の 6-5 で説明する。");
}

/* ═══════════════ 16. 第6章 6-3 SSOの違い ═══════════════ */
{
  const s = content("第6章", "セキュリティ — SAML SSO とは何が違うのか", "Plus の「Microsoftでサインイン」は個人の利便性であり、会社の統制ではありません", "main", "6", NAVY_MID);
  const rows = [
    [hdr("論点", { align: "left" }), hdr("① 通常ログイン / OAuth（Plus）"), hdr("② SAML SSO（Business以上）", { fill: { color: GREEN } })],
    [tx("ログイン方法", { bold: true }), tx("各自がID/PWを入力、またはMicrosoftボタンを押す"), tx("社内 Microsoft 365（Entra ID）の認証画面へ自動転送")],
    [tx("SSOの強制", { bold: true }), mk("×", "社員がNotion専用のパスワードを作れてしまう"), mk("○", "SSO以外のログインを遮断できる")],
    [tx("Microsoft 365 のポリシー適用\n（MFA・端末制限・IP制限）", { bold: true }), mk("△", "パスワードで抜け穴ができ、適用が不完全"), mk("○", "条件付きアクセスが100%適用される")],
    [tx("退職時のアクセス遮断", { bold: true }), mk("×", "Notion側の手動削除を忘れると入れ続ける"), mk("○", "Microsoft 365 の停止と同時に即時遮断")],
    [tx("管理部による一元管理", { bold: true }), mk("×", "社員個人のアカウント管理に依存"), mk("○", "Microsoft 365 の管理画面から一元統制")],
  ];
  table(s, rows, 2.42, [2.9, 4.62, 4.61], { rowH: [0.32, 0.36, 0.42, 0.46, 0.42, 0.42], fontSize: 9.5 });

  const paths = [
    ["SSO なし（Plus プラン）", RED, RED_PALE,
      "退職  →  管理部が Microsoft 365 を停止  →  Notion側の削除を失念\n→  元社員が自宅PCから社内Notionを閲覧・コピーできてしまう",
      "機密情報・ノウハウの流出"],
    ["SSO あり（Business プラン以上）", GREEN, GREEN_PALE,
      "退職  →  管理部が Microsoft 365 を停止  →  Notionの鍵も同時に自動ロック\n→  即座にアクセス不能",
      "人為的な削除ミスによる流出を防止"],
  ];
  paths.forEach((p, i) => {
    const w = CW / 2 - 0.12;
    const x = M + i * (w + 0.24);
    card(s, x, 5.35, w, 1.55, p[2], p[1]);
    s.addText(p[0], { x: x + 0.22, y: 5.47, w: w - 0.44, h: 0.28, fontFace: F, fontSize: 11.5, bold: true, color: p[1], margin: 0 });
    s.addText(p[3], { x: x + 0.22, y: 5.77, w: w - 0.44, h: 0.65, fontFace: F, fontSize: 9.5, color: INK, margin: 0, lineSpacing: 16 });
    s.addText(p[4], { x: x + 0.22, y: 6.46, w: w - 0.44, h: 0.32, valign: "middle", fontFace: F, fontSize: 12, bold: true, color: p[1], margin: 0 });
  });
  s.addNotes("SSO の目的は利便性ではなく、退職者による情報持ち出しを会社として確実に防ぐことにある。");
}

/* ═══════════════ 17. 第6章 6-4 Microsoft 365 環境での挙動 ═══════════════ */
{
  const s = content("第6章", "セキュリティ — 当社 Microsoft 365 環境での挙動", "当社はすでに Microsoft 365（Entra ID）をID基盤として運用しています", "main", "6", NAVY_MID);
  const rows = [
    [hdr("プラン", { align: "left" }), hdr("Microsoft 365（Entra ID）との連携"), hdr("セキュリティ要件の充足状況")],
    [tx("Plus", { bold: true, fontSize: 12 }),
     tx("SAML SSO 連携は不可。各自が Microsoft 365 のメールアドレスで個別登録する形となり、退職者の手動削除漏れリスクが残る。"),
     mk("×", "SSOによる統制ができない")],
    [{ text: "Business", options: { color: WHITE, bold: true, fontSize: 12, fill: { color: GREEN }, align: "left", valign: "middle" } },
     tx("SAML SSO 連携が可能（プロビジョニングは手動）。認証を Entra ID に統合でき、Microsoft 365 を止めればアクセスを遮断できる。", { bold: true }),
     mk("○", "SSOによる統制が可能。監査ログ・SCIM は対象外")],
    [tx("Enterprise", { bold: true, fontSize: 12 }),
     tx("SAML SSO ＋ SCIM（自動同期）連携が可能。Microsoft 365 側で社員を追加・異動・削除すると、Notion側も自動で作成・削除される。"),
     mk("○", "統制機能をすべて充足")],
  ];
  table(s, rows, 2.42, [1.75, 6.2, 4.18], { rowH: [0.34, 0.72, 0.72, 0.72], fontSize: 9.5 });

  card(s, M, 5.72, 7.55, 1.28, AMBER_PALE, AMBER);
  s.addText("Business プランでできないこと（あらかじめ申し上げます）", {
    x: M + 0.25, y: 5.84, w: 7.05, h: 0.26, fontFace: F, fontSize: 11, bold: true, color: AMBER, margin: 0,
  });
  s.addText("監査ログと SCIM は Business では利用できません。①②のスコープでは「機微情報を Notion に置かない」運用で担保します。監査ログを必須要件とされる場合は Enterprise への切替となり、年額は約911.4万円（140名）に変動します。", {
    x: M + 0.25, y: 6.12, w: 7.05, h: 0.76, fontFace: F, fontSize: 9.5, color: INK, margin: 0, lineSpacing: 15,
  });

  card(s, M + 7.78, 5.72, 4.35, 1.28, ICE_PALE, ICE);
  s.addText("運用体制", {
    x: M + 8.0, y: 5.84, w: 3.9, h: 0.26, fontFace: F, fontSize: 11, bold: true, color: NAVY, margin: 0,
  });
  s.addText("当社の Microsoft 365 は管理部が管理しています。Notion との SAML SSO 連携設定、および退職時のアカウント運用も管理部が担います。", {
    x: M + 8.0, y: 6.12, w: 3.9, h: 0.76, fontFace: F, fontSize: 9.5, color: INK, margin: 0, lineSpacing: 15,
  });
  s.addNotes("説明の言い回し例：「当社ではすでに全社のID認証基盤として Microsoft 365（Entra ID）を導入・運用しております。Notion は Business プラン以上の SAML SSO 機能で Entra ID と認証を統合します。これにより社員のパスワード管理リスクを排除し、多要素認証や端末制限ポリシーを Notion にも強制適用できます。退職時にも Microsoft 365 アカウントの無効化と同時に Notion へのアクセスが即座に遮断されます。」");
}

/* ═══════════════ 18. 第6章 6-5 Businessを選ぶ理由（新規） ═══════════════ */
{
  const s = content("第6章", "プラン選定 — Business を選ぶ理由", "Business を選ぶ理由は、セキュリティだけではありません", "main", "6", NAVY_MID);
  const reasons = [
    ["1", "退職時のアクセスを確実に断つ", GREEN,
      "SAML SSO により、Microsoft 365 のアカウントを停止すれば Notion にも同時に入れなくなります。",
      "Plus では、社員が個別に設定したパスワードで退職後もログインできてしまいます。"],
    ["2", "Notion AI を追加費用なしで使える", NAVY,
      "議事録の要約、横断検索、文章作成を全社員が使えます。Business には追加費用なしで含まれます。",
      "Plus は試用のみです。別途契約する場合、1人あたりの月額がさらに上乗せになります。"],
    ["3", "外部の講師・委託先を250名まで招ける", NAVY,
      "ゲストは無料で招待でき、必要なページだけを共有できます。研修事業では社外の講師・パートナーとの協働が多く、この枠が実務上の制約になります。",
      "Plus の上限は100名です。"],
  ];
  reasons.forEach((r, i) => {
    const w = 3.85;
    const x = M + i * (w + 0.29);
    card(s, x, 2.44, w, 2.5, WHITE, LINE);
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.22, y: 2.62, w: 0.36, h: 0.36, rectRadius: 0.18,
      fill: { color: r[2] }, line: { color: r[2] },
    });
    s.addText(r[0], {
      x: x + 0.22, y: 2.62, w: 0.36, h: 0.36, align: "center", valign: "middle",
      fontFace: F, fontSize: 11, bold: true, color: WHITE, margin: 0,
    });
    s.addText(r[1], { x: x + 0.22, y: 3.08, w: w - 0.44, h: 0.6, fontFace: F, fontSize: 12.5, bold: true, color: NAVY, margin: 0 });
    s.addText(r[3], { x: x + 0.22, y: 3.7, w: w - 0.44, h: 0.72, fontFace: F, fontSize: 9.5, color: INK, margin: 0, lineSpacing: 15 });
    card(s, x + 0.18, 4.42, w - 0.36, 0.42, "F7F9FC", LINE);
    s.addText(r[4], {
      x: x + 0.3, y: 4.42, w: w - 0.6, h: 0.42, valign: "middle",
      fontFace: F, fontSize: 8.5, color: MUTED, margin: 0,
    });
  });

  card(s, M, 5.1, 7.55, 1.05, WHITE, LINE);
  s.addText("このほかの差", { x: M + 0.25, y: 5.2, w: 4, h: 0.26, fontFace: F, fontSize: 10.5, bold: true, color: MUTED, margin: 0 });
  s.addText(
    [
      { text: "部署別のプライベートチームスペースを作れる（人事・役員情報の分離）", options: { bullet: true, breakLine: true } },
      { text: "誤削除・改ざん時に遡れる履歴が 30日 → 90日 に延びる", options: { bullet: true } },
    ],
    { x: M + 0.25, y: 5.48, w: 7.05, h: 0.6, fontFace: F, fontSize: 9.5, color: INK, margin: 0, paraSpaceAfter: 4 }
  );

  card(s, M + 7.78, 5.1, 4.35, 1.05, NAVY, NAVY);
  s.addText("SSO を除いても、\nPlus では実務に足りません", {
    x: M + 8.0, y: 5.1, w: 3.95, h: 1.05, valign: "middle",
    fontFace: F, fontSize: 14, bold: true, color: WHITE, margin: 0, lineSpacing: 22,
  });

  s.addText("※ Notion AI のバンドル有無とゲスト上限は、契約時に正規代理店へ確認のうえ確定します。", {
    x: M, y: 6.3, w: CW, h: 0.26, fontFace: F, fontSize: 9, color: MUTED, margin: 0,
  });
  s.addNotes("理由2と3はセキュリティとは独立した論点。研修事業では社外講師・パートナーとの協働が多く、ゲスト枠が実務上の制約になる点を具体的に補足する。");
}

/* ═══════════════ 19. 第7章 導入計画とコスト ═══════════════ */
{
  const s = content("第7章", "導入計画とコスト", "スコープとコストの上限を先に決め、失敗したときにやめられる状態を保ちます", "main", "7", NAVY_MID);

  card(s, M, 2.44, CW, 0.78, NAVY, NAVY);
  s.addText("経営層で3か月試した結果を踏まえ、対象140名・Businessプランで導入します。対象範囲は ①案件管理 ②議事録 に限定します。", {
    x: M + 0.35, y: 2.44, w: CW - 0.7, h: 0.78, valign: "middle",
    fontFace: F, fontSize: 13.5, bold: true, color: WHITE, margin: 0,
  });

  // 左：範囲を絞る
  card(s, M, 3.4, 4.3, 2.5, WHITE, LINE);
  s.addText("絞るのは「人数」ではなく「範囲」", {
    x: M + 0.25, y: 3.54, w: 3.8, h: 0.3, fontFace: F, fontSize: 11.5, bold: true, color: NAVY, margin: 0,
  });
  s.addText(
    [
      { text: "③案件関連ファイル・④社内規定は SharePoint を継続し、Notion に持ち込まない", options: { bullet: true, breakLine: true } },
      { text: "機微情報は Notion に置かない", options: { bullet: true, breakLine: true } },
      { text: "対象業務を①②に絞ることで、合わなければやめられる状態を保つ", options: { bullet: true } },
    ],
    { x: M + 0.25, y: 3.9, w: 3.8, h: 1.85, fontFace: F, fontSize: 9.5, color: INK, margin: 0, paraSpaceAfter: 8 }
  );

  // 中：コスト
  card(s, M + 4.5, 3.4, 4.3, 2.5, NAVY, NAVY);
  s.addText("コスト上限", { x: M + 4.75, y: 3.54, w: 3.8, h: 0.28, fontFace: F, fontSize: 11.5, bold: true, color: ICE, margin: 0 });
  s.addText("約520.8万円", { x: M + 4.75, y: 3.84, w: 3.8, h: 0.72, fontFace: F, fontSize: 34, bold: true, color: WHITE, margin: 0 });
  s.addText("/ 年（140名・Business・年払）", { x: M + 4.75, y: 4.56, w: 3.8, h: 0.28, fontFace: F, fontSize: 11, color: ICE, margin: 0 });
  s.addText("月額換算　約43.4万円\n1人あたり　年 37,200円\n＝ 1営業日あたり 約169円", {
    x: M + 4.75, y: 4.92, w: 3.8, h: 0.85, fontFace: F, fontSize: 10, color: ICE, margin: 0, lineSpacing: 17,
  });

  // 右：続ける判断
  card(s, M + 9.0, 3.4, 3.13, 2.5, WHITE, LINE);
  s.addText("続けるかどうかの判断", {
    x: M + 9.22, y: 3.54, w: 2.7, h: 0.3, fontFace: F, fontSize: 11.5, bold: true, color: NAVY, margin: 0,
  });
  s.addText(
    [
      { text: "効果KPI（案件進捗の可視化率・議事録への到達時間）を測る", options: { bullet: true, breakLine: true } },
      { text: "未達であれば、翌年度の更新を見送る", options: { bullet: true, breakLine: true } },
      { text: "外部委託・社外パートナーはゲスト（無料）とし、課金対象は正規社員に限定する", options: { bullet: true } },
    ],
    { x: M + 9.22, y: 3.9, w: 2.7, h: 1.85, fontFace: F, fontSize: 9, color: INK, margin: 0, paraSpaceAfter: 7 }
  );

  s.addText("※ $20/月は年払時の想定単価（$1=155円換算）。正規代理店の見積により変動しますが、上限額は上振れしない前提でご提示しています。", {
    x: M, y: 6.1, w: CW, h: 0.28, fontFace: F, fontSize: 9, color: MUTED, margin: 0,
  });
  s.addNotes("対象は最初から140名。やめられる根拠は、対象範囲を①②に絞っている点にある。");
}

/* ═══════════════ 20. 第8章 お願い事項 ═══════════════ */
{
  const s = content("第8章", "お願い事項", "①案件管理 ＋ ②議事録 に限定した正式導入と、コスト上限のご承認をいただきたく存じます", "main", "8");
  s.addText("ご承認いただきたいこと", { x: M, y: 2.42, w: 5, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: MUTED, margin: 0 });
  const asks = [
    ["1", "①案件管理 ＋ ②議事録 に限定した正式導入"],
    ["2", "Business プランでの契約（SAML SSO 要件を満たすため）"],
    ["3", "コスト上限：140名 ／ 年額 約520.8万円"],
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

  card(s, M, 4.68, 7.55, 0.62, WHITE, LINE);
  s.addText("11月の組織変更で PMO が発足します。PMO が部署を越えて案件を束ねるための基盤として、本件をお願いするものです。", {
    x: M + 0.28, y: 4.68, w: 7.0, h: 0.62, valign: "middle",
    fontFace: F, fontSize: 10.5, color: INK, margin: 0,
  });

  s.addText("PCA側のお約束", { x: M, y: 5.44, w: 5, h: 0.3, fontFace: F, fontSize: 12, bold: true, color: MUTED, margin: 0 });
  card(s, M, 5.78, 7.55, 1.1, WHITE, LINE);
  s.addText(
    [
      { text: "③案件関連ファイル・④社内規定は SharePoint を継続（Microsoft 365 は置き換えません）", options: { bullet: true, breakLine: true } },
      { text: "機微情報は Notion に置かない運用とします", options: { bullet: true, breakLine: true } },
      { text: "効果測定の結果と、継続の要否をご報告します", options: { bullet: true } },
    ],
    { x: M + 0.28, y: 5.92, w: 7.0, h: 0.85, fontFace: F, fontSize: 10, color: INK, margin: 0, paraSpaceAfter: 5 }
  );

  card(s, M + 7.78, 2.78, 4.35, 4.1, NAVY, NAVY);
  s.addText("ご確認事項", { x: M + 8.0, y: 2.95, w: 3.9, h: 0.32, fontFace: F, fontSize: 13, bold: true, color: WHITE, margin: 0 });
  const qs = [
    ["1", "SaaS導入手続きの様式", "グループ所定のセキュリティチェックシート等がございましたら、受領のうえ別途ご回答を作成します。"],
    ["2", "監査ログを必須とされるか", "必須の場合は Enterprise が対象となり、年額は約911.4万円（140名）に変動します。"],
    ["3", "認証資料のご提出要否", "必要な場合、提出先と形式をご指定ください。"],
  ];
  qs.forEach((q, i) => {
    const y = 3.45 + i * 1.15;
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 8.0, y, w: 0.3, h: 0.3, rectRadius: 0.15,
      fill: { color: ICE }, line: { color: ICE },
    });
    s.addText(q[0], { x: M + 8.0, y, w: 0.3, h: 0.3, align: "center", valign: "middle", fontFace: F, fontSize: 9, bold: true, color: NAVY, margin: 0 });
    s.addText(q[1], { x: M + 8.4, y: y - 0.02, w: 3.5, h: 0.3, valign: "middle", fontFace: F, fontSize: 11, bold: true, color: WHITE, margin: 0 });
    s.addText(q[2], { x: M + 8.4, y: y + 0.3, w: 3.5, h: 0.7, fontFace: F, fontSize: 9, color: ICE, margin: 0, lineSpacing: 14 });
  });
  s.addNotes("ご確認事項②（監査ログの要否）へのご回答をもって、プランとコストが確定する。");
}

/* ═══════════════ 21. Appendix 扉 ═══════════════ */
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
    ["D", "SSOのメリット・デメリット"], ["E", "コスト感度分析"], ["F", "ファイルアップロード仕様"],
    ["G", "出典一覧"], ["H", "会議設計と全社展開の設計"],
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

/* ═══════════════ 22. Appendix A ═══════════════ */
{
  const s = content("Appendix A", "4案比較表（詳細版）", "判定の理由を各欄に注記しています。消去の論拠は本編 第4章をご参照ください", "apx", "A", NAVY_MID);
  const rows = [
    [hdr("要件", { align: "left" }), hdr("A. Microsoft 365 の組合せ"), hdr("B. 無料ツール"), hdr("C. 専用ツール"), hdr("D. Notion", { fill: { color: GREEN } })],
    [tx("1. タスクに担当者・期限が必ず付く", { bold: true }), mk("△", "Planner等で設定は可能だが、ツールごとに運用がばらつく"), mk("△", "入力必須にできず空欄が残る"), mk("○", "タスク管理は本来の得意分野"), mk("○", "必須プロパティとして担保できる")],
    [tx("2. 遅れを部署横断で一覧できる", { bold: true }), mk("×", "Planner/Lists/Loopをまたいだ抽出ができない"), mk("△", "ツール単位に留まる"), mk("△", "当該ツールに登録された範囲のみ"), mk("○", "全案件のタスクを横断ビューで抽出できる")],
    [tx("3. 何のためのタスクか辿れる", { bold: true }), mk("△", "案件・議事録が別管理で結びつかない"), mk("△", "リレーション機能を持たない"), mk("△", "Backlogは親子課題まで。目的まで遡れない（当社試用）"), mk("○", "タスク→案件→決定事項→議事録まで遡れる")],
    [tx("4. 立場別のビューを切り替えられる", { bold: true }), mk("×", "同一データを別形式で見る仕組みが無い"), mk("△", "ビューの種類が限られる"), mk("○", "ガント・ボード等に対応"), mk("○", "同一DBをガント/ボード/カレンダー/テーブルで切替")],
    [tx("5. ガバナンス（SSO・権限）", { bold: true }), mk("○", "Microsoft 365 の統制下にある"), mk("×", "SSOが有料機能。無料利用はシャドーIT化する"), mk("○", "多くが対応"), mk("○", "SAML SSO は Business 以上で対応（第6章）")],
    [tx("6. 1製品で完結", { bold: true }), mk("×", "3製品以上の併用が前提"), mk("×", "用途ごとに別ツール"), mk("×", "議事録・wikiで別ツール併用"), mk("○", "案件・タスク・議事録・wikiを1製品で完結")],
    [tx("（参考）当社での試用", { bold: true }), tx("—", { align: "center" }), tx("—", { align: "center" }), mk("△", "Backlogを試用。ガントは有効だが、目的との結びつきは管理できず"), mk("○", "経営層で3か月試用。管理台帳として使えることを確認")],
  ];
  table(s, rows, 2.38, [2.35, 2.5, 2.35, 2.35, 2.58], { rowH: [0.28, 0.5, 0.5, 0.5, 0.5, 0.5, 0.4, 0.42], fontSize: 7 });
}

/* ═══════════════ 23. Appendix B ═══════════════ */
{
  const s = content("Appendix B", "想定問答", "ご質問が想定される点について、あらかじめ回答を整理しています", "apx", "B", NAVY_MID);
  const rows = [
    [hdr("ご質問", { align: "left" }), hdr("回答", { align: "left" })],
    [tx("Microsoft 365 があるのに二重投資では？", { bold: true }), tx("役割が異なります（Microsoft 365 ＝保管・伝達／Notion ＝構造化された実行管理）。Microsoft 365 は継続し、置き換えません。")],
    [tx("Microsoft 365 Copilot で良いのでは？", { bold: true }), tx("Copilot は検索・要約の道具であり、案件管理の仕組み自体は提供しません。AI は基盤の代替ではなく基盤の上に乗るもので、構造化されたデータベースがある方が AI 活用の精度は上がります。")],
    [tx("Planner や Lists で良いのでは？", { bold: true }), tx("案件⇔議事録⇔タスクのリレーションと立場別ビューが実務水準に達しません。複数アプリの組合せは、いま解決したい「分散」を作り直すことになります。")],
    [tx("セキュリティは大丈夫か？", { bold: true }), tx("SOC 2 Type II・ISO 27001 他を取得済みです。SAML SSO で Microsoft 365 と統合し、機微情報は置かない運用とします（第6章）。")],
    [tx("定着するのか？", { bold: true }), tx("経営層で3か月試した結果に加え、テンプレート・ミニマニュアルを整備したうえで展開します。")],
    [tx("うまくいかなかった場合は？", { bold: true }), tx("対象範囲を①②に限定しているため、合わなければやめられます。効果KPIが未達であれば翌年度の更新を見送ります。")],
    [tx("監査ログが無い Business で、情報持ち出しをどう追跡するのか？", { bold: true, color: NAVY }), tx("Business では追跡できません。①②のスコープでは機微情報を置かない運用で担保します。監査ログを必須要件とされる場合は Enterprise への切替となり、年額約911.4万円（140名）に変動します。")],
    [tx("退職者のアカウントはどう消すのか？", { bold: true, color: NAVY }), tx("SAML SSO により Microsoft 365 のアカウント停止と同時にログイン不能となります。Business は SCIM 非対応のため、課金停止のためのメンバー削除は管理部による手動運用とし、対象者リストを月次で突合します。")],
    [tx("業務委託・社外パートナーはどうするのか？", { bold: true, color: NAVY }), tx("ゲスト（無料）として必要ページのみに招待します。ライセンス課金対象は正規社員に限定し、ゲスト運用ルールは導入前に策定します。")],
  ];
  table(s, rows, 2.42, [4.0, 8.13], { rowH: 0.44, fontSize: 8.5 });
}

/* ═══════════════ 24. Appendix C-1 ═══════════════ */
{
  const s = content("Appendix C-1", "プラン別 導入パターンの比較", "検討した3案と、参考としての Enterprise を、費用の低い順に並べています", "apx", "C", NAVY_MID);
  const rows = [
    [hdr("比較項目", { align: "left" }), hdr("案B：幹部のみBusiness\n＋他はFree"), hdr("案C：Plus 全員"), hdr("案A：Business 全員", { fill: { color: GREEN } }), hdr("（参考）Enterprise 全員")],
    [tx("月額単価（年払・1人）", { bold: true }), tx("幹部30名のみ $20\n他110名 $0", { align: "center" }), tx("$10", { align: "center" }), tx("$20", { align: "center", bold: true }), tx("個別見積（$30〜40目安）", { align: "center" })],
    [tx("年額（140名）", { bold: true }), tx("約 111.6 万円", { align: "center" }), tx("約 260.4 万円", { align: "center" }), tx("約 520.8 万円", { align: "center", bold: true, color: NAVY, fontSize: 11 }), tx("約 911.4 万円", { align: "center" })],
    [tx("SAML SSO", { bold: true }), mk("×", "一般社員は対象外"), mk("×", "利用不可"), mk("○", "全員が対象"), mk("○", "全員が対象")],
    [tx("監査ログ・SCIM", { bold: true }), mk("×", "利用不可"), mk("×", "利用不可"), mk("×", "利用不可"), mk("○", "完全対応")],
    [tx("アクセス・権限統制", { bold: true }), mk("×", "ページごとの手動共有"), mk("△", "閲覧制限が限定的"), mk("○", "部署別チームスペース"), mk("○", "組織単位で制御")],
    [tx("全社ナレッジ共有", { bold: true }), mk("×", "部分的にしか見られない"), mk("○", "基本的な共有は可能"), mk("○", "検索・共同編集が快適"), mk("○", "検索・共同編集が快適")],
    [tx("Notion AI", { bold: true }), tx("幹部30名のみ", { align: "center" }), tx("×  試用のみ", { align: "center", color: RED }), tx("○  全員", { align: "center", color: GREEN, bold: true }), tx("○  全員", { align: "center", color: GREEN })],
    [tx("セキュリティ要件の充足", { bold: true }), tx("満たさない", { align: "center", color: RED }), tx("満たさない（SSO非対応）", { align: "center", color: RED }), tx("満たす（監査ログを除く）", { align: "center", bold: true, color: GREEN }), tx("すべて満たす", { align: "center", color: GREEN })],
  ];
  table(s, rows, 2.42, [2.35, 2.5, 2.35, 2.45, 2.48], { rowH: [0.42, 0.4, 0.32, 0.4, 0.34, 0.4, 0.4, 0.3, 0.3], fontSize: 8 });

  card(s, M, 6.28, CW, 0.62, NAVY, NAVY);
  s.addText("SSO を必須要件とする限り、実務的に成立するのは案A（Business 全員）のみです", {
    x: M + 0.35, y: 6.28, w: CW - 0.7, h: 0.62, valign: "middle",
    fontFace: F, fontSize: 13, bold: true, color: WHITE, margin: 0,
  });
}

/* ═══════════════ 25. Appendix C-2 ═══════════════ */
{
  const s = content("Appendix C-2", "各案の詳細分析", "C-1 の判定に至った理由の裏付けです", "apx", "C", NAVY_MID);
  const cases = [
    ["案B：幹部のみBusiness＋他社員Free", "推奨度 不可・非推奨", RED, RED_PALE,
      "できること：費用を最小限に抑えられる（幹部30名分・約111.6万円/年）。",
      "重大な欠陥：①一般社員はゲスト扱いとなり SAML SSO が適用されず、個人のフリーアカウントでログインする。②ゲストはチームスペース全体を閲覧できず、ページごとに個別招待が必要で全社Wikiとして成立しない。③退職時に Microsoft 365 を停止しても個人アカウントに権限が残る。",
      "判定：セキュリティガバナンスが成立せず、全社の情報基盤としての運用が困難です。"],
    ["案C：Plus 全員", "推奨度 低〜中", AMBER, AMBER_PALE,
      "できること：ページ作成・ブロック数無制限、共同編集、外部ゲスト招待。案Aの半額（約260.4万円/年）。",
      "できないこと：SAML SSO が利用不可（個別のパスワードや二要素認証に依存）。プライベートチームスペースが作成できない。Notion AI は試用のみ。ゲスト上限は100名。",
      "判定：SAML SSO による一括認証を必須条件とされる場合、要件を満たしません。"],
    ["案A：Business 全員", "本資料のご提案", GREEN, GREEN_PALE,
      "できること：全員が Entra ID 経由で SAML SSO ログイン。部署ごとのプライベートチームスペース。Notion AI を全員が利用可能。ゲストは250名まで。",
      "できないこと：監査ログの閲覧・出力（事後追跡が不可）。SCIM による入退社時の自動同期（管理部による手動削除が必要）。",
      "判定：SSO を必須とし監査ログを必須としない場合、コストと機能のバランスが最も良い構成です。"],
  ];
  cases.forEach((c, i) => {
    const y = 2.40 + i * 1.36;
    card(s, M, y, CW, 1.28, c[3], c[2]);
    s.addText(c[0], { x: M + 0.25, y: y + 0.1, w: 4.5, h: 0.3, valign: "middle", fontFace: F, fontSize: 13, bold: true, color: c[2], margin: 0 });
    s.addShape(pres.ShapeType.roundRect, {
      x: M + 4.85, y: y + 0.14, w: 2.3, h: 0.24, rectRadius: 0.04,
      fill: { color: c[2] }, line: { color: c[2] },
    });
    s.addText(c[1], { x: M + 4.85, y: y + 0.14, w: 2.3, h: 0.24, align: "center", valign: "middle", fontFace: F, fontSize: 8.5, bold: true, color: WHITE, margin: 0 });
    s.addText(c[4], { x: M + 0.25, y: y + 0.44, w: CW - 0.5, h: 0.26, fontFace: F, fontSize: 9, color: INK, margin: 0 });
    s.addText(c[5], { x: M + 0.25, y: y + 0.7, w: CW - 0.5, h: 0.34, fontFace: F, fontSize: 9, color: INK, margin: 0 });
    s.addText(c[6], { x: M + 0.25, y: y + 1.02, w: CW - 0.5, h: 0.24, fontFace: F, fontSize: 9.5, bold: true, color: c[2], margin: 0 });
  });

  card(s, M, 6.52, CW, 0.46, WHITE, LINE);
  s.addText("（参考）Enterprise 全員：セキュリティ要件をすべて充足（SSO＋監査ログ＋SCIM＋組織統制）。140名で約911.4万円/年。個別見積が必要です。", {
    x: M + 0.25, y: 6.52, w: CW - 0.5, h: 0.46, valign: "middle", fontFace: F, fontSize: 9.5, color: INK, margin: 0,
  });
}

/* ═══════════════ 26. Appendix D ═══════════════ */
{
  const s = content("Appendix D", "SSO のメリット・デメリット（4つの視点）", "導入する側・使う側・管理する側それぞれの観点を整理しています", "apx", "D", NAVY_MID);
  const rows = [
    [hdr("視点", { align: "left" }), hdr("メリット", { align: "left" }), hdr("デメリット・注意点", { align: "left" })],
    [tx("① 社員\n（使う側）", { bold: true, color: NAVY, fontSize: 11 }),
     tx("・パスワードを覚える必要がありません。Notion 専用の ID/PW が不要で、忘れ・再設定の手間がなくなります。\n・毎朝 PC（Microsoft 365）にログインしていれば、ボタン1つで Notion が開きます。"),
     tx("・Microsoft 側で大規模なシステム障害が起きた場合、Notion にも一時的にログインできなくなります。")],
    [tx("② 会社\n（セキュリティ）", { bold: true, color: NAVY, fontSize: 11 }),
     tx("・退職時に Microsoft 365 のアカウントを停止すれば、Notion のアクセス権も同時に自動遮断されます。\n・「会社の指定PC以外はアクセス禁止」「多要素認証必須」などの制限がそのまま Notion にも効きます。\n・簡単なパスワードや私生活のパスワードの使い回しによる乗っ取りを防げます。"),
     tx("・自社の Microsoft 365 アカウントを持たない外部パートナーを招待する場合、「ゲスト」としての運用ルールを事前に定める必要があります。")],
    [tx("③ 管理部\n（管理・運用）", { bold: true, color: NAVY, fontSize: 11 }),
     tx("・「パスワードを忘れたのでリセットしてほしい」という社内問い合わせがなくなります。\n・入退社時のアカウント削除漏れを防げます。"),
     tx("・導入時に、管理部が Microsoft 365（Entra ID）と Notion を紐付ける初期設定が必要です（約1〜2時間程度の作業）。")],
    [tx("④ コスト\n（費用面）", { bold: true, color: NAVY, fontSize: 11 }),
     tx("・アカウント管理ミスによる情報漏洩の損害と、事後対応コストを未然に防止できます。"),
     tx("・SSO を使うには Business プラン以上が必要です（月額 $20/人〜）。Plus（$10）では利用できず、140名で年間約260.4万円の差額が生じます。")],
  ];
  table(s, rows, 2.42, [1.75, 6.2, 4.18], { rowH: [0.32, 0.85, 1.05, 0.8, 0.85], fontSize: 8.5 });
}

/* ═══════════════ 27. Appendix E ═══════════════ */
{
  const s = content("Appendix E", "コスト感度分析", "対象人数とプランの組み合わせによる年間コストの試算です（$1=155円・年払）", "apx", "E", NAVY_MID);
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
  table(s, rows, 2.42, [3.0, 3.04, 3.04, 3.05], { rowH: 0.44, fontSize: 10.5 });

  const notes = [
    ["前提条件", "年払契約時の金額です。月払契約の場合は約20%割高になります。為替は $1=155円 で換算しています。"],
    ["Enterprise 単価について", "ボリュームディスカウントにより変動するため、Notion 社または正規代理店への相見積もりが必要です。表中の $35 は仮置きの値です。"],
    ["Business 単価について", "$20/月（年払）で試算しています。$18/月とする情報もあり、見積により下振れする可能性があります。上限額は上振れしない前提です。"],
  ];
  notes.forEach((n, i) => {
    const y = 5.4 + i * 0.5;
    s.addText(n[0], { x: M, y, w: 2.4, h: 0.42, valign: "middle", fontFace: F, fontSize: 10, bold: true, color: NAVY, margin: 0 });
    s.addText(n[1], { x: M + 2.5, y, w: CW - 2.5, h: 0.42, valign: "middle", fontFace: F, fontSize: 9.5, color: INK, margin: 0 });
  });
}

/* ═══════════════ 28. Appendix F ═══════════════ */
{
  const s = content("Appendix F", "ファイルアップロード仕様", "Notion 公式ヘルプで確認した、1ファイルあたりの上限です", "apx", "F", NAVY_MID);
  const rows = [
    [hdr("対象", { align: "left" }), hdr("1ファイルあたりの上限"), hdr("補足", { align: "left" })],
    [tx("Free プラン", { bold: true }), tx("5 MB", { align: "center", bold: true, color: RED, fontSize: 12 }), tx("無料プランのみに適用される上限です。")],
    [tx("有料プラン（Plus / Business / Enterprise）", { bold: true }), tx("5 GB", { align: "center", bold: true, color: GREEN, fontSize: 12 }), tx("一般的なファイル（動画・zip 等）の上限です。")],
    [tx("　└ PDF ファイル", {}), tx("20 MB 未満", { align: "center" }), tx("有料プランでも、PDF にはこの個別上限が適用されます。")],
    [tx("　└ 画像（PNG / JPG）", {}), tx("5 MB 未満", { align: "center" }), tx("有料プランでも、画像にはこの個別上限が適用されます。")],
    [tx("ワークスペース全体の総容量", { bold: true }), tx("制限なし", { align: "center", bold: true, color: GREEN }), tx("個々のファイルが上限内であれば、保存総量に制限はありません。")],
  ];
  table(s, rows, 2.42, [4.0, 2.6, 5.53], { rowH: 0.5, fontSize: 10 });

  card(s, M, 5.6, CW, 1.0, ICE_PALE, ICE);
  s.addText("本件のスコープについて", { x: M + 0.25, y: 5.72, w: 6, h: 0.28, fontFace: F, fontSize: 11.5, bold: true, color: NAVY, margin: 0 });
  s.addText("今回のスコープは ①案件管理 ②議事録 に限定しており、③案件関連ファイルの保管は SharePoint を継続します。本表は、議事録や案件ページに添付する資料の上限をご確認いただくためのものです。", {
    x: M + 0.25, y: 6.02, w: CW - 0.5, h: 0.5, fontFace: F, fontSize: 10, color: INK, margin: 0, lineSpacing: 16,
  });

  s.addText("出典：Notion Help Center — Images, files & media（www.notion.com/help/images-files-and-media）", {
    x: M, y: 6.75, w: CW, h: 0.26, fontFace: F, fontSize: 9, color: MUTED, margin: 0,
  });
}

/* ═══════════════ 29. Appendix G 出典一覧 ═══════════════ */
{
  const s = content("Appendix G", "出典一覧", "本資料に記載した仕様・プラン要件は、以下の公式情報で確認しています", "apx", "G", NAVY_MID);
  const srcs = [
    ["ファイルアップロード上限", "Notion Help Center — Images, files & media", "www.notion.com/help/images-files-and-media", "Free 5MB ／ 有料 5GB、PDF 20MB・画像 5MB の個別上限"],
    ["SAML SSO のプラン要件", "Notion Help Center — SAML SSO configuration", "www.notion.com/help/saml-sso-configuration", "Business プラン以上で利用可能。IdP は SAML 2.0 対応が必要"],
    ["監査ログのプラン要件", "Notion Help Center — Workspace audit log", "www.notion.com/help/audit-log", "Enterprise プランの組織オーナーのみ利用可能"],
    ["SCIM のプラン要件", "Notion Help Center — Provision users & groups with SCIM", "www.notion.com/help/provision-users-and-groups-with-scim", "Enterprise プランでのみ利用可能"],
    ["セキュリティ認証の取得状況", "Notion Trust Center", "trust.notion.com", "SOC 2 Type II、ISO 27001/27017/27018/27701、BSI C5 等"],
    ["プラン別の機能・価格", "Notion Pricing", "www.notion.com/pricing", "単価・上限は改定される場合があるため、契約時に見積で確定する"],
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

/* ═══════════════ 30. Appendix H 会議設計と権限設計 ═══════════════ */
{
  const s = content("Appendix H", "会議設計テンプレートと全社展開の権限設計", "第5章 5-3 でご説明した運用を、実際に成立させるための設計です", "apx", "H", NAVY_MID);

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
    "全社共有ページ（全社員＝閲覧可）　← 原本をここに置く\n" +
    "　├─ 同期先 → 営業部ページ\n" +
    "　├─ 同期先 → 管理部ページ\n" +
    "　└─ 同期先 → 各部門ページ\n" +
    "経営会議 議事録（役員＝限定公開）\n" +
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
}

pres.writeFile({ fileName: "notion-thd-proposal.pptx" }).then(() => {
  console.log("done: notion-thd-proposal.pptx / slides =", pageNo + 1);
});
