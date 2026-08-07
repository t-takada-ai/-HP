const pptxgen = require("pptxgenjs");
const L = require("./lib");
const { renderAll } = require("./icons");

const { W, H, M, CW, C, F } = L;

// ---------- local composites ----------

function chartTakeaway(s, { type, data, opts = {}, take, axisNote, y = 1.95, h = 4.35 }) {
  const cardW = 3.55, gap = 0.4;
  const chartW = CW - cardW - gap;
  const inner = L.chartPanel(s, { x: M, y, w: chartW, h });
  if (axisNote) {
    s.addText(axisNote, {
      x: inner.x + 0.05, y: inner.y, w: inner.w - 0.1, h: 0.3, fontSize: 11.5, bold: true,
      color: C.muted, fontFace: F, margin: 0, valign: "middle",
    });
    inner.y += 0.34; inner.h -= 0.34;
  }
  s.addChart(type, data, Object.assign(L.chartBase(), opts, inner));
  const x = M + chartW + gap;
  s.addShape("roundRect", {
    x, y, w: cardW, h, rectRadius: 0.1, fill: { color: C.navy }, line: { color: C.navy },
    shadow: L.shadow({ opacity: 0.18 }),
  });
  L.iconBadge(s, { x: x + 0.34, y: y + 0.34, d: 0.6, name: take.icon, tone: "dark" });
  s.addText(take.title, {
    x: x + 0.34, y: y + 1.1, w: cardW - 0.68, h: 1.0, fontSize: 19, bold: true,
    color: C.white, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.2,
  });
  s.addText(take.body, {
    x: x + 0.34, y: y + 2.15, w: cardW - 0.68, h: h - 2.5, fontSize: 13,
    color: C.ice, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
  });
}

function timeline(s, items, { y = 2.35 } = {}) {
  const n = items.length;
  const lineY = y + 0.62;
  const cardW = (CW - 0.3 * (n - 1)) / n;
  const cxOf = (i) => M + i * (cardW + 0.3) + cardW / 2;
  s.addShape("line", {
    x: cxOf(0), y: lineY, w: cxOf(n - 1) - cxOf(0), h: 0,
    line: { color: C.tintDeep, width: 3 },
  });
  items.forEach((it, i) => {
    const cx = cxOf(i);
    s.addText(it.year, {
      x: cx - 0.9, y, w: 1.8, h: 0.4, align: "center", fontSize: 16, bold: true,
      color: C.navy, fontFace: F, margin: 0, valign: "middle",
    });
    s.addShape("ellipse", {
      x: cx - 0.13, y: lineY - 0.13, w: 0.26, h: 0.26,
      fill: { color: C.navy }, line: { color: C.white, width: 2 },
    });
    const cardX = M + i * (cardW + 0.3);
    s.addShape("roundRect", {
      x: cardX, y: lineY + 0.42, w: cardW, h: 2.25, rectRadius: 0.08,
      fill: { color: C.tint }, line: { color: C.line, width: 1 }, shadow: L.shadow(),
    });
    s.addText(it.title, {
      x: cardX + 0.26, y: lineY + 0.66, w: cardW - 0.52, h: 0.72, fontSize: 15, bold: true,
      color: C.navy, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.15,
    });
    s.addText(it.body, {
      x: cardX + 0.26, y: lineY + 1.42, w: cardW - 0.52, h: 1.15, fontSize: 12.5,
      color: C.steel, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
    });
  });
}

function caseSlide(pres, cfg) {
  const s = L.light(pres);
  L.head(s, cfg.title, cfg.lead);
  const leftW = 3.5;
  s.addShape("roundRect", {
    x: M, y: 1.95, w: leftW, h: 4.3, rectRadius: 0.1,
    fill: { color: cfg.fail ? C.accentSoft : C.navy },
    line: { color: cfg.fail ? "EBD9C2" : C.navy }, shadow: L.shadow({ opacity: 0.15 }),
  });
  L.iconBadge(s, { x: M + 0.34, y: 2.28, d: 0.62, name: cfg.icon, tone: cfg.fail ? "light" : "dark" });
  s.addText(cfg.company, {
    x: M + 0.34, y: 3.06, w: leftW - 0.68, h: 0.9, fontSize: 19, bold: true,
    color: cfg.fail ? C.navy : C.white, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.15,
  });
  s.addText(cfg.profile, {
    x: M + 0.34, y: 4.0, w: leftW - 0.68, h: 2.0, fontSize: 13,
    color: cfg.fail ? C.steel : C.ice, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.4,
  });
  L.rows(s, cfg.items, { y: 1.95, x: M + leftW + 0.4, w: CW - leftW - 0.4, rowH: 1.3, gap: 0.2 });
  L.note(s, cfg.note || "※想定事例(実際の導入相談をもとにした構成例)。自社の実績が入る場合は差し替えてください。");
  return s;
}

// ---------- deck ----------

async function main() {
  L.state.icons = await renderAll({ navy: C.navy, white: C.white, ice: C.ice });

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "セミナー資料";
  pres.title = "人を増やさずに、事務は回る。";

  let s;

  // ===== 第0部 オープニング =====

  // 1 — title
  s = L.dark(pres, { page: false });
  s.addImage({ data: L.icon("FiCpu", "ice"), x: W - 4.2, y: 1.5, w: 3.6, h: 3.6, transparency: 82 });
  s.addText("中小企業の経営者向けセミナー", {
    x: M + 0.3, y: 1.85, w: 7.5, h: 0.4, fontSize: 15, bold: true, color: C.ice,
    fontFace: F, margin: 0, charSpacing: 2,
  });
  s.addText("人を増やさずに、\n事務は回る。", {
    x: M + 0.3, y: 2.45, w: 8.6, h: 2.3, fontSize: 46, bold: true, color: C.white,
    fontFace: F, margin: 0, valign: "middle", lineSpacingMultiple: 1.2,
  });
  s.addText("AIを活用したバックオフィス業務効率化 ─ 経営者が決める、はじめの一歩", {
    x: M + 0.3, y: 4.95, w: 9.0, h: 0.5, fontSize: 16, color: C.ice, fontFace: F, margin: 0,
  });
  s.addShape("line", { x: M + 0.3, y: 5.75, w: 3.2, h: 0, line: { color: "3A4585", width: 1 } });
  s.addText("所要 90分(質疑10分を含む)　|　登壇者名／所属・肩書き", {
    x: M + 0.3, y: 5.95, w: 9.0, h: 0.4, fontSize: 13, color: "9AA5CC", fontFace: F, margin: 0,
  });
  s.addNotes("開始の挨拶は30秒。すぐに次の問いへ入る。");

  // 2 — Q
  s = L.question(pres,
    "御社のバックオフィスは、\n3年後も同じ人数で回せますか?",
    "── 答えは、今日の最後にご自身の中で出してください");
  s.addNotes("ここで数秒止める。挙手も指名もしない。考える時間だけを渡す。");

  // 3 — 今日の3つ
  s = L.light(pres);
  L.head(s, "今日、持ち帰っていただく3つ", "この90分は「AIの勉強」ではなく、「自社の意思決定」の時間です");
  L.cards(s, [
    { num: "1", title: "なぜ「今」なのか", body: "人手不足は景気の波ではなく人口構造の問題。待っても状況は好転しません。その前提を数字で共有します。" },
    { num: "2", title: "自社のどこから\n手をつけるか", body: "経理・総務・人事・労務。どの業務が、どれだけ軽くなるのか。判断できる粒度まで具体化します。" },
    { num: "3", title: "経営者は\n何を決めるのか", body: "現場に丸投げした会社は、ほぼ失敗します。社長が決めるべき3つのことを持ち帰っていただきます。" },
  ], { y: 2.45, h: 3.6 });
  s.addNotes("ゴールを先に示す。特に③が本日の肝であることを予告しておく。");

  // 4 — 講師紹介
  s = L.light(pres);
  L.head(s, "本日の講師", null);
  s.addShape("roundRect", {
    x: M, y: 1.85, w: 4.2, h: 4.4, rectRadius: 0.12, fill: { color: C.navy }, line: { color: C.navy },
    shadow: L.shadow({ opacity: 0.18 }),
  });
  s.addShape("ellipse", { x: M + 1.05, y: 2.35, w: 2.1, h: 2.1, fill: { color: C.ice }, line: { color: C.ice } });
  s.addImage({ data: L.icon("FiUserCheck", "navy"), x: M + 1.6, y: 2.9, w: 1.0, h: 1.0 });
  s.addText("登壇者名", {
    x: M + 0.3, y: 4.65, w: 3.6, h: 0.5, align: "center", fontSize: 22, bold: true,
    color: C.white, fontFace: F, margin: 0,
  });
  s.addText("所属・肩書き", {
    x: M + 0.3, y: 5.2, w: 3.6, h: 0.4, align: "center", fontSize: 13, color: C.ice, fontFace: F, margin: 0,
  });
  L.rows(s, [
    { icon: "FiBriefcase", title: "中小企業の現場で支援", body: "製造・建設・サービス業を中心に、バックオフィス業務の見直しを支援。" },
    { icon: "FiCpu", title: "「使える範囲」に絞って伝える", body: "難しい技術の話はしません。今日から動かせることだけをお話しします。" },
    { icon: "FiUsers", title: "対象は「人を増やせない」会社", body: "大企業のDX事例ではなく、10〜100名規模の現実解を扱います。" },
  ], { y: 1.95, x: M + 4.6, w: CW - 4.6, rowH: 1.35, gap: 0.2 });
  s.addNotes("30秒で切り上げる。実績の羅列より「誰のための話か」を伝えるほうが効く。");

  // ===== 第1部 =====

  // 5 — divider
  s = L.divider(pres, {
    part: "第 1 部", title: "事務が回らなくなる日",
    sub: "まず、これから5年で自社に何が起きるのかを共有します。",
    iconName: "FiTrendingDown",
  });
  s.addNotes("第1部は危機感のパート。事実と数字で押し、感情に訴えすぎない。");

  // 6 — 人手不足の現在地
  s = L.light(pres);
  L.head(s, "人手不足は「感覚」ではなく、構造", "景気が回復しても、人は戻ってきません");
  chartTakeaway(s, {
    type: "line",
    data: [{
      name: "人員が不足していると答えた中小企業の割合(%)",
      labels: ["2016", "2018", "2020", "2022", "2024", "2026"],
      values: [46, 54, 38, 51, 58, 61],
    }],
    opts: { lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 8, valAxisMaxVal: 80, valAxisMinVal: 0 },
    axisNote: "人員が不足していると答えた中小企業の割合(%)",
    take: {
      icon: "FiUsers", title: "コロナ期の\n一時的な緩みは終了",
      body: "採用難は「一時的に厳しい状態」ではなく、平常運転になりました。求人を出せば採れた時代を前提にした人員計画は、すでに成立していません。",
    },
  });
  L.note(s, "※推移イメージ。登壇時は日本商工会議所LOBO調査・中小企業白書の最新値に差し替えてください。");
  s.addNotes("会場に「採用に苦労している方」と軽く問いかけてもよいが、挙手を強要しない。");

  // 7 — Q
  s = L.question(pres, "あと5年、今と同じ人数で\n採用と育成を続けられますか?");
  s.addNotes("採用できるか、ではなく「続けられるか」を問う。継続コストに意識を向けさせる。");

  // 8 — 労働人口
  s = L.light(pres);
  L.head(s, "働き手そのものが減っていく", "採用競争が激しいのではなく、母集団が縮んでいます");
  chartTakeaway(s, {
    type: "area",
    data: [{
      name: "生産年齢人口(15〜64歳・百万人)",
      labels: ["2000", "2010", "2020", "2030", "2040", "2050"],
      values: [86, 81, 75, 69, 60, 52],
    }],
    opts: { chartColors: [C.navy], valAxisMaxVal: 100, valAxisMinVal: 0, showValue: true, dataLabelPosition: "t", dataLabelColor: C.navy },
    axisNote: "生産年齢人口(15〜64歳)/ 百万人",
    take: {
      icon: "FiTrendingDown", title: "2040年には\n今の8割を切る",
      body: "この線は景気対策で変わりません。すでに生まれている人数の話だからです。「人が採れないなら、人がいなくても回る形にする」以外に打ち手がありません。",
    },
  });
  L.note(s, "※推移イメージ。登壇時は総務省・国立社会保障人口問題研究所の最新推計に差し替えてください。");
  s.addNotes("ここは断定してよい。将来推計の中でもっとも確度が高い数字であることを添える。");

  // 9 — 1人依存
  s = L.light(pres);
  L.head(s, "中小企業のバックオフィスは「1人依存」", "業務が人につき、その人しか手順を知らない状態になっています");
  ["請求書の発行と入金消込", "給与計算・社会保険手続き", "経費精算・支払処理", "契約書・規程の管理", "問い合わせの一次対応"].forEach((t, i) => {
    const y = 2.05 + i * 0.78;
    s.addShape("roundRect", { x: M, y, w: 3.5, h: 0.62, rectRadius: 0.08, fill: { color: C.tint }, line: { color: C.line, width: 1 } });
    s.addText(t, { x: M + 0.22, y, w: 3.06, h: 0.62, fontSize: 13, color: C.steel, fontFace: F, margin: 0, valign: "middle" });
    s.addImage({ data: L.icon("FiArrowRight", "navy"), x: M + 3.66, y: y + 0.2, w: 0.22, h: 0.22 });
  });
  s.addShape("roundRect", {
    x: 5.0, y: 2.55, w: 3.3, h: 3.0, rectRadius: 0.12, fill: { color: C.navy }, line: { color: C.navy },
    shadow: L.shadow({ opacity: 0.18 }),
  });
  s.addShape("ellipse", { x: 6.15, y: 3.0, w: 1.0, h: 1.0, fill: { color: C.ice }, line: { color: C.ice } });
  s.addImage({ data: L.icon("FiUserCheck", "navy"), x: 6.4, y: 3.25, w: 0.5, h: 0.5 });
  s.addText("経理・総務担当\n1名", {
    x: 5.25, y: 4.15, w: 2.8, h: 0.9, align: "center", fontSize: 19, bold: true,
    color: C.white, fontFace: F, margin: 0, lineSpacingMultiple: 1.15,
  });
  s.addText("勤続12年・手順は本人の頭の中", {
    x: 5.25, y: 5.05, w: 2.8, h: 0.4, align: "center", fontSize: 12, color: C.ice, fontFace: F, margin: 0,
  });
  s.addImage({ data: L.icon("FiArrowRight", "navy"), x: 8.5, y: 3.94, w: 0.28, h: 0.28 });
  s.addShape("roundRect", {
    x: 8.95, y: 2.95, w: 3.63, h: 2.2, rectRadius: 0.1, fill: { color: C.accentSoft }, line: { color: "EBD9C2", width: 1 },
  });
  L.iconBadge(s, { x: 9.25, y: 3.25, d: 0.6, name: "FiAlertTriangle", tone: "light" });
  s.addText("この1名が抜けた瞬間、\n全社の支払と給与が止まる", {
    x: 9.25, y: 4.0, w: 3.03, h: 1.0, fontSize: 15, bold: true, color: C.navy,
    fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
  });
  s.addNotes("「引き継ぎ書はありますか」と一言添えると、会場の顔が変わる。");

  // 10 — Q
  s = L.question(pres, "その方が明日辞めたら、\n今月の給与計算は誰がやりますか?");
  s.addNotes("第1部の危機感のピーク。具体的な人の顔が浮かぶ問い方をする。");

  // 11 — 間接コスト
  s = L.light(pres);
  L.head(s, "見えていない間接コスト", "事務職1名にかかっている本当の金額を、時間単価まで分解します");
  L.stats(s, [
    { label: "事務職1名の企業負担", value: "520", unit: "万円/年", sub: "年収400万円＋社会保険料・法定福利費・採用教育費を含む総額" },
    { label: "1時間あたりの人件費", value: "2,900", unit: "円", sub: "年間労働1,800時間で割り戻した実質単価(管理コストを除く)" },
    { label: "月20時間の手作業", value: "70", unit: "万円/年", sub: "「たった月20時間」の転記作業が、年間でこの金額になります" },
  ], { y: 2.15, h: 2.6 });
  s.addShape("roundRect", { x: M, y: 5.05, w: CW, h: 1.05, rectRadius: 0.08, fill: { color: C.tint }, line: { color: C.line, width: 1 } });
  L.iconBadge(s, { x: M + 0.3, y: 5.27, d: 0.6, name: "FiDollarSign", tone: "light" });
  s.addText("問題は人件費の高さではなく、その時間が売上にも改善にもつながらない作業に消えていることです。", {
    x: M + 1.15, y: 5.05, w: CW - 1.45, h: 1.05, fontSize: 14, bold: true, color: C.navy,
    fontFace: F, margin: 0, valign: "middle",
  });
  L.note(s, "※試算例です。自社の実際の人件費・労働時間に置き換えてご覧ください。");
  s.addNotes("自社の数字に置き換えてもらうため、ここは少しゆっくり話す。");

  // 12 — 法改正
  s = L.light(pres);
  L.head(s, "そして、事務量は増え続ける", "人が減る一方で、制度対応の事務は毎年積み上がっています");
  timeline(s, [
    { year: "2023", title: "インボイス制度", body: "登録番号の確認、区分記載、経過措置の管理。請求1件あたりの確認工数が増加。" },
    { year: "2024", title: "電子帳簿保存法\n本格運用", body: "電子取引データの保存要件対応。ファイル名規則と検索性の担保が必要に。" },
    { year: "2025", title: "社会保険の\n適用拡大", body: "対象者判定と資格取得手続きが増加。パート比率の高い会社ほど負担が重い。" },
    { year: "2026〜", title: "対応は続く", body: "制度改正は今後も継続。「今年は落ち着く」年は、実務上ほぼ訪れません。" },
  ]);
  s.addNotes("会場の多くが実際に苦労したインボイス・電帳法から入ると共感が取れる。");

  // 13 — ハサミ
  s = L.light(pres);
  L.head(s, "増える仕事 × 減る人", "この2本の線が開き続ける ── それが今起きていることです");
  chartTakeaway(s, {
    type: "line",
    data: [
      { name: "処理すべき事務量", labels: ["2020", "2022", "2024", "2026", "2028", "2030"], values: [100, 108, 118, 128, 138, 150] },
      { name: "対応できる人員", labels: ["2020", "2022", "2024", "2026", "2028", "2030"], values: [100, 97, 93, 88, 82, 75] },
    ],
    opts: {
      lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 7,
      chartColors: [C.accent, C.navy], showLegend: true, legendPos: "b",
      legendColor: C.muted, legendFontFace: F, legendFontSize: 11,
      valAxisMaxVal: 160, valAxisMinVal: 60,
    },
    axisNote: "2020年 = 100 とした指数",
    take: {
      icon: "FiActivity", title: "開いた差は\n誰が埋めるのか",
      body: "この差を、これまでは残業と気合いで埋めてきました。差が広がり続ける以上、同じやり方では必ずどこかで折れます。",
    },
  });
  L.note(s, "※2020年を100とした指数のイメージ図です。");
  s.addNotes("この1枚が第1部の主張。ここだけは図を指しながら丁寧に。");

  // 14 — 打ち手の限界
  s = L.light(pres);
  L.head(s, "これまでの打ち手は、もう効かない", "3つの定番の対処法が、いずれも限界を迎えています");
  L.cards(s, [
    { icon: "FiClock", title: "残業でしのぐ", body: "上限規制で物理的に不可能。人件費も増える一方で、担当者の離職リスクが上がります。" },
    { icon: "FiUsers", title: "パートを採用する", body: "そもそも応募が来ません。採用できても、教える時間を既存担当者から奪います。" },
    { icon: "FiBriefcase", title: "外部に委託する", body: "コストが読める代わりに、社内にノウハウが残りません。委託先も人手不足です。" },
  ], { y: 2.05, h: 3.2 });
  s.addShape("roundRect", { x: M, y: 5.5, w: CW, h: 0.85, rectRadius: 0.08, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText("いずれも「人を前提にした打ち手」── 人が減る局面では、どれも効きません。", {
    x: M + 0.4, y: 5.5, w: CW - 0.8, h: 0.85, fontSize: 15, bold: true, color: C.white,
    fontFace: F, margin: 0, valign: "middle",
  });
  s.addNotes("否定して終わらない。次のスライドで唯一の方向性を示す。");

  // 15 — I statement
  s = L.statement(pres, {
    kicker: "残された選択肢",
    main: "人を増やさずに、\n処理量を増やす。",
    sub: "採用でも、根性でも、外注でもない。同じ人数のまま、こなせる量そのものを引き上げる。\nそれが、これから10年の中小企業の前提条件になります。",
    iconName: "FiZap",
  });
  s.addNotes("ここで一拍置く。第2部の「では、どうやって」につなぐ。");

  // 16 — まとめ
  s = L.light(pres);
  L.head(s, "第1部のまとめ", null);
  L.rows(s, [
    { icon: "FiTrendingDown", title: "人手不足は景気ではなく人口の問題", body: "待っていても解決しません。生産年齢人口は2040年に向けて一直線に減っていきます。" },
    { icon: "FiUserX", title: "バックオフィスは1人依存で止まる", body: "経理・総務が1名体制の会社では、その1名の不在がそのまま経営リスクになります。" },
    { icon: "FiCalendar", title: "制度対応で事務量は増え続ける", body: "インボイス、電帳法、社会保険。事務が減る年は当面来ません。" },
    { icon: "FiZap", title: "打ち手は「人を増やさず処理量を増やす」だけ", body: "残業・採用・外注はすべて人が前提。人が減る局面では機能しません。" },
  ], { y: 1.72, rowH: 1.14, gap: 0.16 });
  s.addNotes("4点を30秒で復唱。長く語らずに第2部へ。");

  // ===== 第2部 =====

  // 17 — divider
  s = L.divider(pres, {
    part: "第 2 部", title: "では、AIは\n何をしてくれるのか",
    sub: "技術の話はしません。「自社の何が変わるのか」だけを扱います。",
    iconName: "FiCpu",
  });
  s.addNotes("第2部は理解のパート。専門用語を使わないことを冒頭で宣言する。");

  // 18 — Q
  s = L.question(pres, "「AIって、結局うちの\n何に使えるんですか?」",
    "── 多くの経営者が、ここで止まっています");
  s.addNotes("参加者の本音を代弁する位置づけ。ここは共感を取りにいく。");

  // 19 — 正体は3つ
  s = L.light(pres);
  L.head(s, "生成AIの正体は、この3つの代行", "難しく考える必要はありません。できることは、突き詰めるとこれだけです");
  L.cards(s, [
    { icon: "FiSearch", title: "読む", body: "長い資料、規程、契約書、メールの束を読み取って、必要な部分だけを抜き出す。\n\n例:100ページの就業規則から該当条文を探す" },
    { icon: "FiEdit3", title: "書く", body: "形式の決まった文章を、条件を伝えるだけで下書きまで仕上げる。\n\n例:議事録、案内文、求人票、報告書の初稿" },
    { icon: "FiGrid", title: "整理する", body: "バラバラな情報を、決めた形式に並べ替える。表にする。要点にまとめる。\n\n例:問い合わせ内容の分類、データの転記" },
  ], { y: 2.0, h: 4.15 });
  s.addNotes("「考える」「判断する」は入っていない点を強調しておくと、後半のリスクの話に効く。");

  // 20 — 相性
  s = L.light(pres);
  L.head(s, "だから、バックオフィスと相性が最も良い", "業務時間のうち「文字と数字を扱う時間」の割合が、圧倒的に高いからです");
  chartTakeaway(s, {
    type: "bar",
    data: [{
      name: "業務時間に占める文字・数字を扱う時間の割合(%)",
      labels: ["経理・総務・人事", "営業", "製造・現場"],
      values: [78, 45, 18],
    }],
    opts: { barDir: "bar", showValue: true, dataLabelPosition: "outEnd", valAxisMaxVal: 100, valAxisMinVal: 0, barGapWidthPct: 60, chartColors: [C.navy, "AEC3E8", "AEC3E8"] },
    axisNote: "業務時間に占める文字・数字を扱う時間の割合(%)",
    take: {
      icon: "FiFileText", title: "現場より先に、\n事務から効く",
      body: "AIが得意なのは、まさにバックオフィスが1日中やっている仕事です。工場のライン改善より、はるかに早く成果が出ます。",
    },
  });
  L.note(s, "※業務実態のイメージ図です。");
  s.addNotes("「うちは製造業だからAIは関係ない」と思っている社長にこそ刺さる1枚。");

  // 21 — できる/苦手
  s = L.light(pres);
  L.head(s, "できること / まだ苦手なこと", "期待しすぎると必ず失敗します。先に線を引いておきます");
  L.compare(s,
    {
      label: "得意", tone: "after", head: "任せてよい仕事",
      items: [
        "決まった形式の文章を下書きする",
        "長い資料から必要な情報を探す",
        "バラバラな情報を表や要点に整理する",
        "同じ作業を、疲れずに何度でも繰り返す",
        "たたき台を、数十秒で出す",
      ],
    },
    {
      label: "苦手", tone: "before", head: "人が引き受けるべき仕事",
      items: [
        "最終的な判断と責任を負うこと",
        "社内の暗黙のルールを察すること",
        "事実確認(もっともらしい誤りを出す)",
        "最新の法改正を正確に把握すること",
        "誰にどう伝えるかの機微を読むこと",
      ],
    },
    { y: 1.95 }
  );
  s.addNotes("「たたき台までは任せる、決めるのは人」という線引きを何度も繰り返す。");

  // 22 — 3レベル
  s = L.light(pres);
  L.head(s, "活用には3つのレベルがある", "自社がどこを目指すのかを、先に決めておく必要があります");
  L.steps(s, [
    { title: "個人の作業補助", body: "担当者が自分の作業でAIを使う。文書の下書き、要約、調べもの。\n\n導入期間:即日\n費用:月数千円/人" },
    { title: "業務フローへの組込", body: "手順を決めて、部署の業務としてAIを使う。誰がやっても同じ品質になる。\n\n導入期間:1〜3か月\n費用:月数万円" },
    { title: "システム連携", body: "会計・販売管理システムと接続し、処理を自動化する。\n\n導入期間:半年〜\n費用:数十万円〜" },
  ], { y: 2.05, h: 3.5 });
  s.addShape("roundRect", { x: M, y: 5.75, w: CW, h: 0.75, rectRadius: 0.08, fill: { color: C.tint }, line: { color: C.line, width: 1 } });
  s.addText("多くの会社が、いきなりレベル3の話を聞いて「うちには無理だ」と止まってしまいます。", {
    x: M + 0.35, y: 5.75, w: CW - 0.7, h: 0.75, fontSize: 14, color: C.navy, fontFace: F, margin: 0, valign: "middle",
  });
  s.addNotes("レベル3の見積を持ってこられて止まった会社の話をすると伝わりやすい。");

  // 23 — 中小はレベル1-2
  s = L.light(pres);
  L.head(s, "中小企業は、レベル1と2で十分に効く", "大企業のDX事例と、自社の打ち手を切り離してください");
  L.compare(s,
    {
      label: "大企業", tone: "before", head: "よく報道されるDX",
      items: [
        "全社システムの刷新から入る",
        "専任のDX推進部を立ち上げる",
        "投資額は数千万円〜数億円",
        "成果が出るまで2〜3年",
        "専門人材を中途採用して進める",
      ],
    },
    {
      label: "中小企業", tone: "after", head: "今日から動かせる現実解",
      items: [
        "1つの業務、1人から始める",
        "今いる担当者がそのまま使う",
        "初期費用ゼロ、月数千円から",
        "早ければ1か月で効果が見える",
        "新しい人材の採用は不要",
      ],
    },
    { y: 1.95 }
  );
  s.addNotes("「うちには専門の人がいない」という反論を、ここで先回りして潰しておく。");

  // 24 — Q
  s = L.question(pres, "いま、社員が「勝手に」使っているのを\n把握できていますか?",
    "── 禁止していない限り、すでに誰かが使っています");
  s.addNotes("シャドーAIの問題提起。第5部のルール整備への伏線として置く。");

  // 25 — 費用感
  s = L.light(pres);
  L.head(s, "費用感の現実", "投資判断の話ではなく、月々の消耗品に近い金額です");
  chartTakeaway(s, {
    type: "bar",
    data: [{
      name: "月額コスト(円)",
      labels: ["AI利用料(1名分)", "コピー機リース", "パート1名(月80時間)", "正社員事務1名"],
      values: [3000, 25000, 110000, 430000],
    }],
    opts: { barDir: "bar", showValue: true, dataLabelPosition: "outEnd", barGapWidthPct: 55, chartColors: [C.navy] },
    axisNote: "月額コスト(円)",
    take: {
      icon: "FiDollarSign", title: "決裁が要る額\nですらない",
      body: "月3,000円は、社長決裁の対象にもならない金額です。にもかかわらず、多くの会社がこの一歩を踏み出せずにいます。障害はコストではありません。",
    },
  });
  L.note(s, "※一般的な水準からの試算例です。ツール・契約形態により変動します。");
  s.addNotes("「高いから導入できない」という言い訳を、ここで完全に外す。");

  // 26 — まとめ
  s = L.light(pres);
  L.head(s, "第2部のまとめ", null);
  L.rows(s, [
    { icon: "FiZap", title: "AIができるのは「読む・書く・整理する」の代行", body: "判断はしません。だからこそ、任せる範囲を人が決める必要があります。" },
    { icon: "FiFileText", title: "文字を扱う仕事が多いバックオフィスに最も効く", body: "現場改善より早く、確実に成果が見えます。" },
    { icon: "FiLayers", title: "中小企業はレベル1・2で十分", body: "全社システム刷新は不要。1業務・1人から始められます。" },
    { icon: "FiDollarSign", title: "止めているのはコストではなく、判断", body: "月数千円。決裁が要らない金額で、最初の検証はできます。" },
  ], { y: 1.72, rowH: 1.14, gap: 0.16 });
  s.addNotes("第3部で「では自社のどこか」に落とすことを予告する。");

  // ===== 第3部 =====

  // 27 — divider
  s = L.divider(pres, {
    part: "第 3 部", title: "うちの、\nどの業務が変わるのか",
    sub: "経理・総務・人事・労務・営業事務。部門ごとに、具体的に見ていきます。",
    iconName: "FiMap",
  });
  s.addNotes("ここからが本題。自社の業務を思い浮かべながら聞いてもらうよう促す。");

  // 28 — 業務マップ
  s = L.light(pres);
  L.head(s, "バックオフィス業務マップ", "この後、この5領域を順番に見ていきます");
  L.cards(s, [
    { icon: "FiDollarSign", title: "経理", body: "請求書処理\n仕訳入力\n入金消込\n月次締め" },
    { icon: "FiMessageCircle", title: "総務", body: "社内問い合わせ\n議事録作成\n文書・規程管理\n備品・契約管理" },
    { icon: "FiUserCheck", title: "人事", body: "求人票作成\n応募者対応\n日程調整\n入社手続き" },
    { icon: "FiShield", title: "労務", body: "給与計算の確認\n社保手続き\n就業規則の改定\n法改正対応" },
    { icon: "FiClipboard", title: "営業事務", body: "見積書作成\n提案書の下書き\n受注データ整理\n顧客リスト管理" },
  ], { y: 1.95, h: 3.6, gap: 0.28 });
  s.addShape("roundRect", { x: M, y: 5.75, w: CW, h: 0.85, rectRadius: 0.08, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText("どの部門も、仕事の中身は「探す・書く・転記する・確認する」の組み合わせでできています。", {
    x: M + 0.4, y: 5.75, w: CW - 0.8, h: 0.85, fontSize: 15, bold: true, color: C.white,
    fontFace: F, margin: 0, valign: "middle",
  });
  s.addNotes("最後のカードが第3部の結論の伏線。ここでは軽く触れる程度に。");

  // 29 — 経理の内訳
  s = L.light(pres);
  L.head(s, "経理:時間が溶けている工程", "月次業務の内訳を分解すると、多くが定型作業に消えています");
  chartTakeaway(s, {
    type: "bar",
    data: [{
      name: "月間工数(時間)",
      labels: ["請求書の受領・確認", "仕訳入力・転記", "入金消込", "経費精算チェック", "月次資料の作成", "分析・経営への報告"],
      values: [22, 28, 14, 16, 12, 4],
    }],
    opts: { barDir: "bar", showValue: true, dataLabelPosition: "outEnd", barGapWidthPct: 45, chartColors: [C.navy] },
    axisNote: "月間工数(時間)",
    take: {
      icon: "FiPieChart", title: "本来やるべき\n仕事が4時間",
      body: "経営判断に効く「分析と報告」に使えているのは、全体のわずか4%。残りは、やらなければ止まるが、付加価値は生まない作業です。",
    },
  });
  L.note(s, "※従業員30名規模・経理2名体制での試算例です。");
  s.addNotes("「うちも似たようなものだ」と思わせられれば成功。");

  // 30 — 経理 B/A
  s = L.light(pres);
  L.head(s, "経理 ─ Before / After", "同じ人が、同じ業務を、短い時間で終える形に変わります");
  L.compare(s,
    {
      label: "Before", tone: "before", head: "いまのやり方",
      items: [
        "届いた請求書を1件ずつ目視で確認する",
        "金額・登録番号を会計ソフトに手入力する",
        "摘要を過去の仕訳から探して合わせる",
        "月次資料は前月ファイルを開いて作り直す",
        "月末3日間は他の仕事が止まる",
      ],
    },
    {
      label: "After", tone: "after", head: "AIを挟んだ形",
      items: [
        "請求書を読み取り、必要項目を一覧に自動整理",
        "仕訳の候補まで出させ、人は確認と承認だけ",
        "摘要のルールを覚えさせ、表記ゆれを吸収",
        "月次コメントの下書きは自動生成、加筆で完成",
        "月末の山が平準化し、締めが1日早まる",
      ],
    },
    { y: 1.95 }
  );
  s.addNotes("重要なのは「人が消える」のではなく「確認と承認に集中する」形になること。");

  // 31 — 経理の試算
  s = L.light(pres);
  L.head(s, "経理:削減イメージの試算", "自社の数字に置き換えてご覧ください");
  L.stats(s, [
    { label: "月間の削減時間", value: "32", unit: "時間", sub: "請求書処理・仕訳・月次資料作成の合計から算出した削減幅" },
    { label: "年間に換算すると", value: "384", unit: "時間", sub: "ひと月あたり約2人日。年間では正社員0.2人分に相当します" },
    { label: "人件費換算", value: "111", unit: "万円/年", sub: "時間単価2,900円で換算。AI利用料は年間4万円程度です" },
  ], { y: 2.15, h: 2.6 });
  s.addShape("roundRect", { x: M, y: 5.05, w: CW, h: 1.05, rectRadius: 0.08, fill: { color: C.tint }, line: { color: C.line, width: 1 } });
  L.iconBadge(s, { x: M + 0.3, y: 5.27, d: 0.6, name: "FiTarget", tone: "light" });
  s.addText("狙いは人件費の削減ではありません。「1人辞めても回る状態」を、採用せずに作ることです。", {
    x: M + 1.15, y: 5.05, w: CW - 1.45, h: 1.05, fontSize: 14, bold: true, color: C.navy,
    fontFace: F, margin: 0, valign: "middle",
  });
  L.note(s, "※試算例です。業務量・体制により結果は変動します。");
  s.addNotes("削減額より「属人化の解消」を前面に出したほうが、経営者には響く。");

  // 32 — 総務の問い合わせ
  s = L.light(pres);
  L.head(s, "総務:同じ質問に、何度も答えている", "1件は3分でも、年間では無視できない時間になります");
  L.rows(s, [
    { icon: "FiHelpCircle", title: "「有給の残日数って何日でしたっけ」", body: "社員50名 × 年6回 × 5分 = 年間25時間。台帳を開いて確認し、口頭で答える作業の繰り返し。" },
    { icon: "FiBookOpen", title: "「この経費、精算できますか」", body: "規程のどこに書いてあるかを都度探す。担当者によって回答が違うこともある。" },
    { icon: "FiFileText", title: "「あの申請書、どこにありますか」", body: "共有フォルダの階層を辿って探す。見つからず作り直すことすらある。" },
  ], { y: 2.0, rowH: 1.3, gap: 0.24 });
  s.addShape("roundRect", { x: M, y: 6.0, w: CW, h: 0.72, rectRadius: 0.08, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText("社内規程・様式・過去の回答をAIに読み込ませておけば、この一次対応はほぼ不要になります。", {
    x: M + 0.35, y: 6.0, w: CW - 0.7, h: 0.72, fontSize: 14, bold: true, color: C.white, fontFace: F, margin: 0, valign: "middle",
  });
  s.addNotes("「総務が本来やるべき仕事は何ですか」と問いかけてから次へ。");

  // 33 — 総務 B/A
  s = L.light(pres);
  L.head(s, "総務 ─ Before / After", "問い合わせ対応と文書作成が、まとめて軽くなります");
  L.compare(s,
    {
      label: "Before", tone: "before", head: "いまのやり方",
      items: [
        "社内からの質問に、その都度手を止めて答える",
        "規程の該当箇所を毎回探し直す",
        "会議のたびに議事録を1時間かけて起こす",
        "案内文・通知文をゼロから書き起こす",
        "担当者が休むと、社内が止まる",
      ],
    },
    {
      label: "After", tone: "after", head: "AIを挟んだ形",
      items: [
        "よくある質問は、規程を読み込ませたAIが一次回答",
        "根拠となる条文の場所まで示させる",
        "録音から要点・決定事項・TODOを自動抽出",
        "案内文は下書きを出させ、確認して発信",
        "担当者不在でも、一次対応は止まらない",
      ],
    },
    { y: 1.95 }
  );
  s.addNotes("最後の行(担当者不在でも止まらない)が、経営者にとって最大の価値。");

  // 34 — 議事録の実例
  s = L.light(pres);
  L.head(s, "実例:1時間の会議が、3分で議事録になる", "音声を渡すだけで、この形まで自動で整理されます");
  const gw = (CW - 0.7) / 2;
  s.addShape("roundRect", { x: M, y: 1.95, w: gw, h: 4.3, rectRadius: 0.1, fill: { color: C.tint }, line: { color: C.line, width: 1 } });
  s.addShape("roundRect", { x: M + 0.32, y: 2.25, w: 1.8, h: 0.4, rectRadius: 0.2, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText("会議の録音", { x: M + 0.32, y: 2.25, w: 1.8, h: 0.4, align: "center", valign: "middle", fontSize: 13, bold: true, color: C.white, fontFace: F, margin: 0 });
  s.addText("62分・出席6名・発言は散在", { x: M + 2.28, y: 2.25, w: gw - 2.6, h: 0.4, fontSize: 12, color: C.muted, fontFace: F, margin: 0, valign: "middle" });
  s.addText(
    "…で、その件なんですけど、前回の話だと来月からという話でしたよね。いや、それは仮の話で、実際には見積がまだ出ていなくて。そうすると発注は再来月ですか。うーん、価格次第ですね。あと、田中さんのところで人が足りていないという話もあって…",
    { x: M + 0.32, y: 2.85, w: gw - 0.64, h: 3.1, fontSize: 12, color: C.muted, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.4 }
  );
  const rx = M + gw + 0.7;
  s.addShape("roundRect", { x: rx, y: 1.95, w: gw, h: 4.3, rectRadius: 0.1, fill: { color: C.navy }, line: { color: C.navy }, shadow: L.shadow({ opacity: 0.18 }) });
  s.addShape("roundRect", { x: rx + 0.32, y: 2.25, w: 1.8, h: 0.4, rectRadius: 0.2, fill: { color: C.ice }, line: { color: C.ice } });
  s.addText("出力された議事録", { x: rx + 0.32, y: 2.25, w: 1.8, h: 0.4, align: "center", valign: "middle", fontSize: 12, bold: true, color: C.navy, fontFace: F, margin: 0 });
  s.addText("作成時間:約3分", { x: rx + 2.28, y: 2.25, w: gw - 2.6, h: 0.4, fontSize: 12, color: C.ice, fontFace: F, margin: 0, valign: "middle" });
  s.addText([
    { text: "決定事項", options: { bold: true, color: C.ice, fontSize: 13, breakLine: true } },
    { text: "・発注は見積受領後に判断。目標は再来月初旬", options: { color: C.white, fontSize: 13, breakLine: true } },
    { text: "・価格が想定を超える場合は再検討", options: { color: C.white, fontSize: 13, breakLine: true, paraSpaceAfter: 10 } },
    { text: "TODO", options: { bold: true, color: C.ice, fontSize: 13, breakLine: true } },
    { text: "・見積取得(担当:購買 / 期限:今月20日)", options: { color: C.white, fontSize: 13, breakLine: true } },
    { text: "・人員状況の確認(担当:製造部 / 期限:今週中)", options: { color: C.white, fontSize: 13, breakLine: true, paraSpaceAfter: 10 } },
    { text: "継続協議", options: { bold: true, color: C.ice, fontSize: 13, breakLine: true } },
    { text: "・製造部の要員不足への対応方針", options: { color: C.white, fontSize: 13 } },
  ], { x: rx + 0.32, y: 2.85, w: gw - 0.64, h: 3.1, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.35 });
  s.addNotes("実物を見せるのがいちばん早い。可能なら当日その場で実演してもよい。");

  // 35 — 人事
  s = L.light(pres);
  L.head(s, "人事:採用まわりは、文書仕事の塊", "「人が採れない」の前に、採用業務そのものが重すぎるケースが多くあります");
  L.rows(s, [
    { icon: "FiFileText", title: "求人票が、何年も使い回しのまま", body: "書き直す時間がないため、他社と横並びの内容に。応募が集まらない原因になっている。" },
    { icon: "FiMessageCircle", title: "応募者への連絡が後回しになる", body: "文面を考える手間で返信が遅れ、その間に他社に決まってしまう。" },
    { icon: "FiClipboard", title: "入社手続きの書類が毎回ゼロから", body: "雇用契約書、社会保険、労働条件通知書。定型なのに、毎回作り直している。" },
  ], { y: 2.0, rowH: 1.3, gap: 0.24 });
  s.addShape("roundRect", { x: M, y: 6.0, w: CW, h: 0.72, rectRadius: 0.08, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText("求人票の書き直しは、AIがもっとも早く効果を出す業務のひとつです。", {
    x: M + 0.35, y: 6.0, w: CW - 0.7, h: 0.72, fontSize: 14, bold: true, color: C.white, fontFace: F, margin: 0, valign: "middle",
  });
  s.addNotes("採用難の話は第1部と接続する。「採れないから効率化」ではなく両輪であること。");

  // 36 — 人事 B/A
  s = L.light(pres);
  L.head(s, "人事 ─ Before / After", "採用のスピードが上がると、それ自体が採用力になります");
  L.compare(s,
    {
      label: "Before", tone: "before", head: "いまのやり方",
      items: [
        "求人票は数年前の原稿をそのまま使う",
        "応募者への返信文を毎回考える",
        "書類選考は担当者が全件を読み込む",
        "面接の日程調整メールを手作業で往復",
        "内定通知・入社書類を都度作成",
      ],
    },
    {
      label: "After", tone: "after", head: "AIを挟んだ形",
      items: [
        "自社の強みを伝えれば、求人票の案を複数出せる",
        "返信文はテンプレート化し、その日のうちに返す",
        "職務要件との一致点・懸念点を先に整理させる",
        "調整文面を自動生成し、往復回数を減らす",
        "書類一式の下書きを自動作成、確認して発行",
      ],
    },
    { y: 1.95 }
  );
  s.addNotes("選考の最終判断は必ず人が行う点を、この場でも一度添えておく。");

  // 37 — 労務
  s = L.light(pres);
  L.head(s, "労務:法改正のキャッチアップが重い", "調べる時間が取れず、対応が後手に回っている会社が少なくありません");
  L.cards(s, [
    { icon: "FiBookOpen", title: "就業規則の改定", body: "改正内容を読み解き、自社の規程のどこを直すかを特定する作業。専門書を読む時間が取れない。" },
    { icon: "FiMessageCircle", title: "社内への周知文", body: "制度が変わるたびに、社員向けの説明文を作る必要がある。難しい言葉のままでは伝わらない。" },
    { icon: "FiClipboard", title: "手続き書類の整備", body: "様式の変更、記載事項の追加。毎年どこかが変わり、確認だけで時間が消える。" },
  ], { y: 2.05, h: 3.3 });
  s.addShape("roundRect", { x: M, y: 5.6, w: CW, h: 0.8, rectRadius: 0.08, fill: { color: C.tint }, line: { color: C.line, width: 1 } });
  s.addText("改正内容の要約、自社規程との突き合わせ、社員向けの平易な説明文。この3つは下書きまで任せられます。", {
    x: M + 0.35, y: 5.6, w: CW - 0.7, h: 0.8, fontSize: 14, color: C.navy, fontFace: F, margin: 0, valign: "middle",
  });
  s.addNotes("次のスライドで必ず注意点に触れる。ここだけで終わらせない。");

  // 38 — 労務の注意点
  s = L.light(pres);
  L.head(s, "ただし、労務にはブレーキが要る", "この領域だけは、確認を省略してはいけません");
  L.cards(s, [
    { icon: "FiAlertTriangle", title: "最新の法改正に\n追いついていない", body: "AIは学習時点の情報で答えます。施行日・経過措置は、必ず一次情報で確認してください。" },
    { icon: "FiSlash", title: "もっともらしい\n誤りを出す", body: "存在しない条文番号を自信たっぷりに答えることがあります。条文は原典に当たること。" },
    { icon: "FiShield", title: "最終判断は\n人と専門家が行う", body: "社労士・顧問への確認は今まで通り必要です。AIは相談の準備を速くするだけです。" },
  ], { y: 2.05, h: 3.3, accent: true });
  s.addShape("roundRect", { x: M, y: 5.6, w: CW, h: 0.8, rectRadius: 0.08, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText("AIは「士業の代わり」ではありません。士業に相談する前の、整理と下準備の道具です。", {
    x: M + 0.35, y: 5.6, w: CW - 0.7, h: 0.8, fontSize: 15, bold: true, color: C.white, fontFace: F, margin: 0, valign: "middle",
  });
  s.addNotes("会場に士業の方がいる可能性を踏まえ、対立ではなく補完であることを明言する。");

  // 39 — 営業事務
  s = L.light(pres);
  L.head(s, "営業事務:営業が事務で止まっている", "売る時間を、事務作業が削っているケースです");
  L.rows(s, [
    { icon: "FiClipboard", title: "見積書・提案書の作成", body: "過去案件を探して流用し、体裁を整える。1件あたり1〜2時間。営業が自分でやっていることも多い。" },
    { icon: "FiRepeat", title: "受注データの転記", body: "メールやFAXで来た内容を、販売管理システムに手入力。転記ミスが後工程のトラブルになる。" },
    { icon: "FiGrid", title: "顧客リストの整理", body: "表記ゆれ、重複、更新漏れ。整理する時間がなく、使えないリストのまま放置されている。" },
  ], { y: 2.0, rowH: 1.3, gap: 0.24 });
  s.addShape("roundRect", { x: M, y: 6.0, w: CW, h: 0.72, rectRadius: 0.08, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText("営業1人あたり月10時間が戻れば、それは訪問件数と見積提出数に直結します。", {
    x: M + 0.35, y: 6.0, w: CW - 0.7, h: 0.72, fontSize: 14, bold: true, color: C.white, fontFace: F, margin: 0, valign: "middle",
  });
  s.addNotes("バックオフィスの改善が、売上側にも効くという接続を示す。");

  // 40 — Q
  s = L.question(pres, "この中で、御社がいちばん\n「人の時間」を使っているのはどれですか?",
    "── 経理 / 総務 / 人事 / 労務 / 営業事務");
  s.addNotes("直前の業務マップに戻す。ここで1つ思い浮かべてもらうことが第6部の伏線になる。");

  // 41 — 4つの型
  s = L.light(pres);
  L.head(s, "削減の型は、4つしかない", "部門が違っても、軽くなる仕事の構造は同じです");
  L.cards(s, [
    { icon: "FiSearch", title: "探す", body: "規程・過去資料を探す時間。\n読み込ませておけば、探す作業そのものがなくなります。" },
    { icon: "FiEdit3", title: "書く", body: "定型文書をゼロから書く時間。\n下書きを任せ、人は確認と加筆に集中します。" },
    { icon: "FiRepeat", title: "転記する", body: "同じ情報を打ち直す時間。\n読み取りと整形を任せ、ミスも同時に減ります。" },
    { icon: "FiCheckSquare", title: "確認する", body: "抜け漏れ・表記ゆれの確認。\n一次チェックを任せ、人は最終確認だけ。" },
  ], { y: 2.0, h: 3.75 });
  s.addShape("roundRect", { x: M, y: 5.95, w: CW, h: 0.75, rectRadius: 0.08, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText("自社の業務を、この4つの言葉で棚卸ししてください。それが第5部でお話しする最初の一歩です。", {
    x: M + 0.35, y: 5.95, w: CW - 0.7, h: 0.75, fontSize: 14, bold: true, color: C.white, fontFace: F, margin: 0, valign: "middle",
  });
  s.addNotes("この4語は、この後のワークシート(配布資料)の見出しと揃えておくとよい。");

  // 42 — 積み上げ
  s = L.light(pres);
  L.head(s, "全社で積み上げると、何人分になるか", "部門ごとの削減時間を合算した試算です");
  chartTakeaway(s, {
    type: "bar",
    data: [{
      name: "月間削減時間(時間)",
      labels: ["経理", "総務", "人事", "労務", "営業事務"],
      values: [32, 26, 14, 10, 30],
    }],
    opts: { barDir: "col", showValue: true, dataLabelPosition: "outEnd", barGapWidthPct: 60, chartColors: [C.navy], valAxisMaxVal: 45, valAxisMinVal: 0 },
    axisNote: "月間削減時間(時間)",
    take: {
      icon: "FiBarChart2", title: "月112時間\n= 0.7人分",
      body: "年間1,344時間。正社員0.7人分の時間が、採用せずに生まれる計算になります。この時間をどこに振り向けるかが、経営判断です。",
    },
  });
  L.note(s, "※従業員50名規模での試算例です。実際の効果は業務量と定着度により変動します。");
  s.addNotes("「0.7人分」という言い方が、経営者にはいちばん伝わる。");

  // 43 — I statement
  s = L.statement(pres, {
    kicker: "誤解のないように",
    main: "これは、\n「人減らし」の話ではない。",
    sub: "空いた時間を、値上げ交渉に。新規開拓に。後継者の育成に。現場の改善に。\n中小企業にいちばん足りていないのは、人手ではなく「考える時間」です。",
    iconName: "FiTarget",
  });
  s.addShape("roundRect", { x: M + 0.3, y: 6.0, w: 3.3, h: 0.6, rectRadius: 0.3, fill: { color: "2A3470" }, line: { color: "3A4585", width: 1 } });
  s.addText("値上げ交渉に使う", { x: M + 0.3, y: 6.0, w: 3.3, h: 0.6, align: "center", valign: "middle", fontSize: 13, bold: true, color: C.ice, fontFace: F, margin: 0 });
  s.addShape("roundRect", { x: M + 3.85, y: 6.0, w: 3.3, h: 0.6, rectRadius: 0.3, fill: { color: "2A3470" }, line: { color: "3A4585", width: 1 } });
  s.addText("新規開拓に使う", { x: M + 3.85, y: 6.0, w: 3.3, h: 0.6, align: "center", valign: "middle", fontSize: 13, bold: true, color: C.ice, fontFace: F, margin: 0 });
  s.addShape("roundRect", { x: M + 7.4, y: 6.0, w: 3.3, h: 0.6, rectRadius: 0.3, fill: { color: "2A3470" }, line: { color: "3A4585", width: 1 } });
  s.addText("後継者の育成に使う", { x: M + 7.4, y: 6.0, w: 3.3, h: 0.6, align: "center", valign: "middle", fontSize: 13, bold: true, color: C.ice, fontFace: F, margin: 0 });
  s.addNotes("社員の不安を煽らないため、社内展開時もこの説明から入るよう助言する。");

  // 44 — まとめ
  s = L.light(pres);
  L.head(s, "第3部のまとめ", null);
  L.rows(s, [
    { icon: "FiMap", title: "5領域すべてに、削減できる仕事がある", body: "経理・総務・人事・労務・営業事務。どこか1つだけの話ではありません。" },
    { icon: "FiFilter", title: "軽くなる仕事の型は「探す・書く・転記・確認」", body: "この4語で自社の業務を棚卸しすれば、対象はすぐに見つかります。" },
    { icon: "FiBarChart2", title: "全社で積み上げると人1人分に届く", body: "50名規模で月100時間超。採用せずに生まれる時間です。" },
    { icon: "FiTarget", title: "空いた時間の使い道は、経営が決める", body: "削減して終わりでは意味がありません。何に振り向けるかが成果を分けます。" },
  ], { y: 1.72, rowH: 1.14, gap: 0.16 });
  s.addNotes("第4部で「では、やった会社はどうなったか」に接続する。");

  // ===== 第4部 =====

  // 45 — divider
  s = L.divider(pres, {
    part: "第 4 部", title: "やった会社と、\nやらなかった会社",
    sub: "同じ業種、同じ規模。1年で何が変わったのかを見ていきます。",
    iconName: "FiTrendingUp",
  });
  s.addNotes("事例パート。数字より「何を決めたか」に焦点を当てて話す。");

  // 46 — 1年後の差
  s = L.light(pres);
  L.head(s, "1年で、ここまで差がつく", "同業・同規模2社の、バックオフィス業務時間の推移");
  chartTakeaway(s, {
    type: "line",
    data: [
      { name: "取り組んだA社", labels: ["開始時", "3か月", "6か月", "9か月", "12か月"], values: [100, 94, 82, 74, 68] },
      { name: "様子を見たB社", labels: ["開始時", "3か月", "6か月", "9か月", "12か月"], values: [100, 102, 105, 108, 112] },
    ],
    opts: {
      lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 8,
      chartColors: [C.navy, C.accent], showLegend: true, legendPos: "b",
      legendColor: C.muted, legendFontFace: F, legendFontSize: 11,
      valAxisMaxVal: 130, valAxisMinVal: 50,
    },
    axisNote: "開始時 = 100 とした事務業務時間の指数",
    take: {
      icon: "FiActivity", title: "差は44ポイント\n1年で開いた",
      body: "B社が悪化しているのは、何もしなかったからではありません。制度対応で事務量が自然に増えたためです。現状維持は、後退を意味します。",
    },
  });
  L.note(s, "※開始時を100とした指数のイメージ図です。");
  s.addNotes("「何もしない」の中身が現状維持ではなく後退である、という点が最大のメッセージ。");

  // 47 — 事例A
  s = caseSlide(pres, {
    title: "事例A:製造業・従業員30名", lead: "経理2名体制のうち、1名分の時間を新規業務へ振り向けた",
    company: "金属加工業\n従業員30名", icon: "FiSettings",
    profile: "経理2名・総務1名。月末は全員が締め作業に追われ、原価分析まで手が回っていなかった。",
    items: [
      { icon: "FiAlertTriangle", title: "課題", body: "請求書処理と仕訳入力で月50時間。ベテラン経理の退職が2年後に迫っていた。" },
      { icon: "FiPlay", title: "やったこと", body: "請求書の読み取りと仕訳候補の自動作成から着手。1名・1業務・3か月で検証した。" },
      { icon: "FiTrendingUp", title: "結果", body: "月32時間を削減。空いた時間を原価分析に充て、赤字受注2件の是正につながった。" },
    ],
  });
  s.addNotes("「削減」より「原価分析ができるようになった」という成果を強調する。");

  // 48 — 事例B
  s = caseSlide(pres, {
    title: "事例B:サービス業・従業員50名", lead: "問い合わせ対応を内製化し、外注費と待ち時間を同時に削減した",
    company: "設備メンテナンス業\n従業員50名", icon: "FiMessageCircle",
    profile: "総務1名が全社の問い合わせを一手に引き受け、休むと社内が止まる状態だった。",
    items: [
      { icon: "FiAlertTriangle", title: "課題", body: "社内問い合わせが1日15件。総務担当が本来業務に着手できるのは夕方以降だった。" },
      { icon: "FiPlay", title: "やったこと", body: "就業規則・経費規程・様式集を読み込ませ、一次回答の仕組みを社内に用意した。" },
      { icon: "FiTrendingUp", title: "結果", body: "問い合わせの6割が自己解決。担当者不在でも社内が止まらない体制になった。" },
    ],
  });
  s.addNotes("属人化の解消がそのままBCP(事業継続)になる、という話につなげてもよい。");

  // 49 — 事例C
  s = caseSlide(pres, {
    title: "事例C:導入したのに、使われなかった会社", lead: "同じツールを入れても、結果が正反対になることがあります",
    company: "卸売業\n従業員40名", icon: "FiSlash", fail: true,
    profile: "経営会議で導入を決定し、全社員にアカウントを配布。3か月後、利用者は2名だけだった。",
    items: [
      { icon: "FiUserX", title: "何が起きたか", body: "「各自で使ってみてほしい」と全社展開。誰も自分の業務との結びつけ方が分からなかった。" },
      { icon: "FiClock", title: "現場の声", body: "「今の仕事で手一杯。新しいことを覚える時間がない」── 通常業務は一切減っていなかった。" },
      { icon: "FiAlertTriangle", title: "結末", body: "半年で契約を解除。社内には「うちには合わなかった」という記憶だけが残った。" },
    ],
    note: "※想定事例。実際にもっとも多い失敗のパターンです。",
  });
  s.addNotes("失敗事例を必ず入れる。成功事例だけだと「うまくいった会社の話」で終わる。");

  // 50 — 失敗の共通点
  s = L.light(pres);
  L.head(s, "失敗する会社の、3つの共通点", "いずれも技術の問題ではなく、進め方の問題です");
  L.cards(s, [
    { icon: "FiUserX", title: "現場に丸投げした", body: "「各自で使ってみて」で終わらせた。通常業務が減らないまま、新しい仕事だけが増える形になっていた。" },
    { icon: "FiHelpCircle", title: "目的が曖昧だった", body: "「AIを導入する」が目的化した。どの業務を、どれだけ軽くするのかが決まっていなかった。" },
    { icon: "FiBarChart2", title: "効果を測らなかった", body: "before/afterの時間を記録していないため、成果が見えず、続ける理由も説明できなかった。" },
  ], { y: 2.05, h: 3.3 });
  s.addShape("roundRect", { x: M, y: 5.6, w: CW, h: 0.8, rectRadius: 0.08, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText("3つとも、現場では決められないことです。つまり、失敗の原因は経営側にあります。", {
    x: M + 0.35, y: 5.6, w: CW - 0.7, h: 0.8, fontSize: 15, bold: true, color: C.white, fontFace: F, margin: 0, valign: "middle",
  });
  s.addNotes("ここで経営者を当事者にする。次のQで自問させる。");

  // 51 — Q
  s = L.question(pres, "導入したのに使われない。\nその原因は、現場ですか。経営ですか。");
  s.addNotes("答えは前のスライドで既に出ている。沈黙の時間を長めに取る。");

  // 52 — I statement
  s = L.statement(pres, {
    kicker: "差を生んでいるもの",
    main: "違いは技術力ではなく、\n「決めたかどうか」。",
    sub: "A社もB社も、使ったツールは同じでした。分かれ目は、どの業務を、誰が、いつまでに、\nという範囲を経営が決めたかどうか。それだけです。",
    iconName: "FiFlag",
  });
  s.addNotes("第5部への橋渡し。「では何を決めるのか」を予告して切り替える。");

  // ===== 第5部 =====

  // 53 — divider
  s = L.divider(pres, {
    part: "第 5 部", title: "では、社長は\n何を決めるのか",
    sub: "進め方の4ステップと、決めておくべきルールを整理します。",
    iconName: "FiFlag",
  });
  s.addNotes("ここからは実践編。持ち帰って使える粒度で話す。");

  // 54 — 4STEP
  s = L.light(pres);
  L.head(s, "進め方は、この4ステップ", "順番を飛ばさないことが、いちばんの近道です");
  L.steps(s, [
    { title: "棚卸し", body: "時間を食っている業務を洗い出す。\n\n所要:2週間\n経営の関与:指示を出す" },
    { title: "選定", body: "効果と難易度で、最初の1つを決める。\n\n所要:1週間\n経営の関与:自ら決める" },
    { title: "小さく試す", body: "1業務・1か月・1人で検証する。\n\n所要:1か月\n経営の関与:報告を受ける" },
    { title: "型化して展開", body: "手順書にして、他部署へ広げる。\n\n所要:3か月〜\n経営の関与:評価に組み込む" },
  ], { y: 2.05, h: 3.4 });
  s.addShape("roundRect", { x: M, y: 5.7, w: CW, h: 0.8, rectRadius: 0.08, fill: { color: C.tint }, line: { color: C.line, width: 1 } });
  s.addText("失敗する会社は、この1と2を飛ばして、いきなり3から始めています。", {
    x: M + 0.35, y: 5.7, w: CW - 0.7, h: 0.8, fontSize: 14, bold: true, color: C.navy, fontFace: F, margin: 0, valign: "middle",
  });
  s.addNotes("所要期間を示すことで「半年後には形になる」という見通しを持たせる。");

  // 55 — STEP1
  s = L.light(pres);
  L.head(s, "STEP 1:棚卸し ─ 時間を食う業務を出す", "「感覚で多そうな業務」ではなく、実際に測った時間で並べます");
  L.rows(s, [
    { icon: "FiClipboard", title: "各担当者に、2週間だけ作業時間を記録してもらう", body: "細かい分類は不要。「何に、何分」だけ。エクセルでも紙でも構いません。" },
    { icon: "FiFilter", title: "「探す・書く・転記・確認」の4分類で仕分ける", body: "第3部の4つの型を使います。この4つに当てはまる作業が、そのまま候補になります。" },
    { icon: "FiBarChart2", title: "時間の多い順に並べ、上位5つを書き出す", body: "この5つが、次のステップで判断する対象です。全部やろうとしないこと。" },
  ], { y: 2.0, rowH: 1.3, gap: 0.24 });
  s.addShape("roundRect", { x: M, y: 6.0, w: CW, h: 0.72, rectRadius: 0.08, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText("経営者がやることは1つ ── 「2週間、時間を記録してほしい」と指示を出すことです。", {
    x: M + 0.35, y: 6.0, w: CW - 0.7, h: 0.72, fontSize: 14, bold: true, color: C.white, fontFace: F, margin: 0, valign: "middle",
  });
  s.addNotes("配布資料に記録シートを付けておくと、実行率が大きく変わる。");

  // 56 — STEP2 マトリクス
  s = L.light(pres);
  L.head(s, "STEP 2:選定 ─ 最初の1つを決める", "効果の大きさと、着手のしやすさで並べます");
  const ox = 1.5, oy = 2.15, cw2 = 2.55, ch2 = 1.95;
  const quads = [
    { i: 0, j: 0, label: "最初に取る", body: "効果 大 × 着手 易", hot: true },
    { i: 1, j: 0, label: "次に取る", body: "効果 大 × 着手 難", hot: false },
    { i: 0, j: 1, label: "余力で取る", body: "効果 小 × 着手 易", hot: false },
    { i: 1, j: 1, label: "やらない", body: "効果 小 × 着手 難", hot: false },
  ];
  quads.forEach((q) => {
    const x = ox + q.i * cw2, y = oy + q.j * ch2;
    s.addShape("roundRect", {
      x, y, w: cw2 - 0.12, h: ch2 - 0.12, rectRadius: 0.08,
      fill: { color: q.hot ? C.navy : C.tint }, line: { color: q.hot ? C.navy : C.line, width: 1 },
      shadow: q.hot ? L.shadow({ opacity: 0.2 }) : undefined,
    });
    s.addText(q.label, {
      x: x + 0.2, y: y + 0.45, w: cw2 - 0.52, h: 0.5, align: "center", fontSize: 17, bold: true,
      color: q.hot ? C.white : C.navy, fontFace: F, margin: 0, valign: "middle",
    });
    s.addText(q.body, {
      x: x + 0.2, y: y + 0.98, w: cw2 - 0.52, h: 0.4, align: "center", fontSize: 12,
      color: q.hot ? C.ice : C.muted, fontFace: F, margin: 0, valign: "middle",
    });
  });
  s.addText("効果 大 ↑", { x: 0.6, y: oy - 0.05, w: 0.85, h: 0.35, fontSize: 11, bold: true, color: C.muted, fontFace: F, margin: 0, align: "right" });
  s.addText("効果 小 ↓", { x: 0.6, y: oy + 2 * ch2 - 0.42, w: 0.85, h: 0.35, fontSize: 11, bold: true, color: C.muted, fontFace: F, margin: 0, align: "right" });
  s.addText("着手しやすい", { x: ox, y: oy + 2 * ch2 - 0.04, w: cw2 - 0.12, h: 0.35, align: "center", fontSize: 11, bold: true, color: C.muted, fontFace: F, margin: 0 });
  s.addText("着手しにくい", { x: ox + cw2, y: oy + 2 * ch2 - 0.04, w: cw2 - 0.12, h: 0.35, align: "center", fontSize: 11, bold: true, color: C.muted, fontFace: F, margin: 0 });
  s.addShape("roundRect", { x: 7.15, y: 2.15, w: 5.43, h: 4.15, rectRadius: 0.1, fill: { color: C.tint }, line: { color: C.line, width: 1 }, shadow: L.shadow() });
  L.iconBadge(s, { x: 7.5, y: 2.5, d: 0.6, name: "FiTarget", tone: "light" });
  s.addText("最初の1つは、必ず社長が決める", {
    x: 7.5, y: 3.25, w: 4.75, h: 0.55, fontSize: 19, bold: true, color: C.navy, fontFace: F, margin: 0, valign: "middle",
  });
  s.addText("現場に選ばせると、いちばん無難な業務が選ばれます。効果が小さいため成果が見えず、そのまま立ち消えになります。\n\n「請求書処理から始める」「議事録から始める」── この一言を経営者が言えるかどうかで、結果が変わります。", {
    x: 7.5, y: 3.9, w: 4.75, h: 2.15, fontSize: L.fit("現場に選ばせると、いちばん無難な業務が選ばれます。効果が小さいため成果が見えず、そのまま立ち消えになります。\n\n「請求書処理から始める」「議事録から始める」── この一言を経営者が言えるかどうかで、結果が変わります。", 4.75, 2.15, 13.5, 1.35, 11),
    color: C.steel, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
  });
  s.addNotes("マトリクスの左上だけを指して「ここから」と言い切る。");

  // 57 — STEP3
  s = L.light(pres);
  L.head(s, "STEP 3:小さく試す ─ 1業務・1か月・1人", "全社導入から始めた会社は、ほぼ例外なく失敗しています");
  L.stats(s, [
    { label: "対象業務", value: "1", unit: "つだけ", sub: "STEP2で選んだ業務に限定。増やしたくなっても、この期間は広げない" },
    { label: "検証期間", value: "1", unit: "か月", sub: "短すぎると慣れず、長すぎると熱が冷める。1か月がちょうどよい長さ" },
    { label: "担当者", value: "1", unit: "名", sub: "いちばん困っている人を選ぶ。前向きな人ではなく、負荷が重い人" },
  ], { y: 2.15, h: 2.6 });
  s.addShape("roundRect", { x: M, y: 5.05, w: CW, h: 1.05, rectRadius: 0.08, fill: { color: C.tint }, line: { color: C.line, width: 1 } });
  L.iconBadge(s, { x: M + 0.3, y: 5.27, d: 0.6, name: "FiClock", tone: "light" });
  s.addText("必ず、開始前と終了後の作業時間を記録してください。数字がないと、次の判断ができません。", {
    x: M + 1.15, y: 5.05, w: CW - 1.45, h: 1.05, fontSize: 14, bold: true, color: C.navy, fontFace: F, margin: 0, valign: "middle",
  });
  s.addNotes("「担当者は前向きな人ではなく、いちばん困っている人」は現場で効く助言。");

  // 58 — STEP4
  s = L.light(pres);
  L.head(s, "STEP 4:型化して、横に広げる", "1人の工夫で終わらせず、会社の仕組みに変えます");
  L.rows(s, [
    { icon: "FiClipboard", title: "手順を1枚の紙にまとめる", body: "「何を入力すると、何が出てくるか」だけで十分。担当者本人に書いてもらいます。" },
    { icon: "FiUsers", title: "同じ業務を持つ人に渡す", body: "説明会は不要。手順書を渡して、隣で1回一緒にやれば伝わります。" },
    { icon: "FiSettings", title: "次の業務に着手する", body: "STEP2で選んだ2番目の業務へ。1つずつ、確実に潰していきます。" },
    { icon: "FiAward", title: "成果を評価に反映する", body: "改善した人が報われる形にしないと、2周目は起きません。ここは経営の仕事です。" },
  ], { y: 1.85, rowH: 1.12, gap: 0.18 });
  s.addNotes("最後の1行が、継続するかどうかの分岐点であることを強調する。");

  // 59 — 経営者の役割
  s = L.light(pres);
  L.head(s, "経営者の役割は「やっていい」と決めること", "現場は、許可がないと動けません");
  s.addShape("roundRect", { x: M, y: 2.05, w: 5.4, h: 4.1, rectRadius: 0.1, fill: { color: C.tint }, line: { color: C.line, width: 1 } });
  s.addText("現場の頭の中", { x: M + 0.35, y: 2.3, w: 4.7, h: 0.45, fontSize: 16, bold: true, color: C.navy, fontFace: F, margin: 0, valign: "middle" });
  s.addText([
    { text: "「勝手に使って、情報漏れたら責任問題では」", options: { bullet: true, breakLine: true, paraSpaceAfter: 12 } },
    { text: "「今の仕事を止めてまで、やっていいのか」", options: { bullet: true, breakLine: true, paraSpaceAfter: 12 } },
    { text: "「失敗したら、自分の評価が下がる」", options: { bullet: true, breakLine: true, paraSpaceAfter: 12 } },
    { text: "「そもそも、会社として認めているのか」", options: { bullet: true } },
  ], { x: M + 0.35, y: 2.9, w: 4.7, h: 3.0, fontSize: 14, color: C.steel, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.3 });
  s.addImage({ data: L.icon("FiArrowRight", "navy"), x: 6.42, y: 3.95, w: 0.35, h: 0.35 });
  s.addShape("roundRect", { x: 7.18, y: 2.05, w: 5.4, h: 4.1, rectRadius: 0.1, fill: { color: C.navy }, line: { color: C.navy }, shadow: L.shadow({ opacity: 0.18 }) });
  s.addText("社長の一言で外れる", { x: 7.53, y: 2.3, w: 4.7, h: 0.45, fontSize: 16, bold: true, color: C.ice, fontFace: F, margin: 0, valign: "middle" });
  s.addText("「この業務で、\n1か月試してほしい。\n失敗しても構わない。」", {
    x: 7.53, y: 3.0, w: 4.7, h: 1.9, fontSize: 24, bold: true, color: C.white, fontFace: F, margin: 0, valign: "middle", lineSpacingMultiple: 1.3,
  });
  s.addText("投資判断でも、ツール選定でもありません。決めるのは、やっていいという範囲だけです。", {
    x: 7.53, y: 5.05, w: 4.7, h: 0.9, fontSize: 13, color: C.ice, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
  });
  s.addNotes("このスライドが第5部の核。ここは時間をかけてよい。");

  // 60 — 動かない理由
  s = L.light(pres);
  L.head(s, "現場が動かない3つの理由と、その外し方", "精神論では動きません。構造を外してください");
  L.cards(s, [
    { icon: "FiClock", title: "時間がない", body: "【外し方】\n通常業務を減らす許可とセットで出す。「その分、月次資料は簡略化してよい」まで言う。" },
    { icon: "FiAlertTriangle", title: "怒られたくない", body: "【外し方】\n「失敗しても構わない」を明言する。検証期間中の失敗は責任を問わないと宣言する。" },
    { icon: "FiAward", title: "評価されない", body: "【外し方】\n改善した時間を成果として評価する。朝礼で名前を出すだけでも効きます。" },
  ], { y: 2.05, h: 3.4 });
  s.addShape("roundRect", { x: M, y: 5.7, w: CW, h: 0.8, rectRadius: 0.08, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText("3つとも、社長にしか外せません。だから「現場に丸投げ」は必ず失敗するのです。", {
    x: M + 0.35, y: 5.7, w: CW - 0.7, h: 0.8, fontSize: 15, bold: true, color: C.white, fontFace: F, margin: 0, valign: "middle",
  });
  s.addNotes("第4部の失敗事例と対応させながら話すと、腹落ちが早い。");

  // 61 — リスク
  s = L.light(pres);
  L.head(s, "リスク管理:3つの落とし穴", "使わないリスクと同じくらい、無防備に使うリスクも現実的です");
  L.cards(s, [
    { icon: "FiLock", title: "情報漏洩", body: "顧客名簿・個人情報・見積単価を、無防備に入力してしまう。入力してよい情報の線引きが必要です。" },
    { icon: "FiAlertTriangle", title: "誤った情報", body: "事実と異なる内容を、もっともらしく出力します。そのまま社外に出すと信用問題になります。" },
    { icon: "FiFileText", title: "著作権・権利", body: "生成物が他者の権利を侵害する可能性。特に画像・ロゴ・キャッチコピーは要注意です。" },
  ], { y: 2.05, h: 3.4, accent: true });
  s.addShape("roundRect", { x: M, y: 5.7, w: CW, h: 0.8, rectRadius: 0.08, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText("いずれも、ルールを1枚決めておけば大半は防げます。難しい規程は要りません。", {
    x: M + 0.35, y: 5.7, w: CW - 0.7, h: 0.8, fontSize: 15, bold: true, color: C.white, fontFace: F, margin: 0, valign: "middle",
  });
  s.addNotes("怖がらせて終わらない。次のスライドで対処法を必ずセットで見せる。");

  // 62 — 社内ルール
  s = L.light(pres);
  L.head(s, "最低限、決めておく社内ルール3項目", "A4で1枚。これ以上は要りません");
  L.checklist(s, [
    { title: "入力してよい情報の線引き", body: "顧客名・個人情報・単価・未公開情報は入力しない。判断に迷うものは上長に確認する。" },
    { title: "使ってよいツールの限定", body: "会社が認めたサービスのみ使用する。無料版・個人アカウントの業務利用は禁止する。" },
    { title: "社外に出す前の確認者を決める", body: "AIが作った文章を社外へ出す場合は、必ず担当者以外が1回確認する。" },
  ], { y: 2.1, itemH: 1.0, gap: 0.24 });
  s.addShape("roundRect", { x: M, y: 5.85, w: CW, h: 0.85, rectRadius: 0.08, fill: { color: C.tint }, line: { color: C.line, width: 1 } });
  L.iconBadge(s, { x: M + 0.3, y: 6.0, d: 0.56, name: "FiShield", tone: "light" });
  s.addText("この3行を朝礼で共有するだけでも、無法状態からは抜け出せます。", {
    x: M + 1.1, y: 5.85, w: CW - 1.4, h: 0.85, fontSize: 14, bold: true, color: C.navy, fontFace: F, margin: 0, valign: "middle",
  });
  s.addNotes("完璧な規程を作ろうとして止まる会社が多い。3行で十分だと言い切る。");

  // 63 — Q
  s = L.question(pres, "ルールなしで使われるのと、\nルールを決めて使わせるの。\nどちらが危険ですか?",
    "── 禁止しても、使用はなくなりません");
  s.addNotes("禁止は選択肢にならないことを、この問いで理解してもらう。");

  // 64 — 公的支援
  s = L.light(pres);
  L.head(s, "使える公的支援もあります", "自社だけで抱え込む必要はありません");
  L.cards(s, [
    { icon: "FiDollarSign", title: "IT導入補助金", body: "対象ツールの導入費用の一部が補助されます。対象・要件は年度ごとに変わるため、公募要領の確認を。" },
    { icon: "FiMessageCircle", title: "商工会議所の\n経営相談窓口", body: "無料で相談できます。何から手をつけるかの整理だけでも、外部の目を入れる価値があります。" },
    { icon: "FiUserCheck", title: "専門家派遣制度", body: "専門家を一定回数まで低額または無料で招けます。棚卸しの段階から入ってもらうのが有効です。" },
  ], { y: 2.05, h: 3.4 });
  s.addShape("roundRect", { x: M, y: 5.7, w: CW, h: 0.8, rectRadius: 0.08, fill: { color: C.tint }, line: { color: C.line, width: 1 } });
  s.addText("補助金ありきで進めないこと。まず1業務で試し、効果が見えてから制度を使うのが順番です。", {
    x: M + 0.35, y: 5.7, w: CW - 0.7, h: 0.8, fontSize: 14, bold: true, color: C.navy, fontFace: F, margin: 0, valign: "middle",
  });
  L.note(s, "※制度内容・要件は年度により変更されます。登壇時点の最新の公募要領をご確認ください。");
  s.addNotes("主催の商工会議所の窓口名を入れておくと、その場で相談につながりやすい。");

  // ===== 第6部 =====

  // 65 — divider
  s = L.divider(pres, {
    part: "第 6 部", title: "持ち帰っていただくこと",
    sub: "最後に、明日から動かすための3つのアクションを整理します。",
    iconName: "FiPlay",
  });
  s.addNotes("残り時間を確認。押していれば要点3つとアクションだけに絞る。");

  // 66 — 要点3つ
  s = L.light(pres);
  L.head(s, "今日の要点", "この3つだけ覚えて帰ってください");
  L.cards(s, [
    { num: "1", title: "人手不足は、\n待っても終わらない", body: "景気ではなく人口構造の問題です。事務量は制度対応で増え続け、対応できる人は減り続けます。この差は自然には埋まりません。" },
    { num: "2", title: "バックオフィスから\n始めるのが最短", body: "探す・書く・転記・確認。この4つの作業が多い部門ほど早く効きます。月数千円、1業務、1人から検証できます。" },
    { num: "3", title: "決めるのは、\n経営者です", body: "現場は許可がないと動けません。どの業務を、誰が、いつまでに試すのか。ここを決めるだけで結果が変わります。" },
  ], { y: 2.05, h: 3.9 });
  s.addNotes("3つを言い切って、次の問いに入る。補足説明は加えない。");

  // 67 — Q
  s = L.question(pres, "1年後、同じ業務に\n同じ時間をかけていていいですか?",
    "── 冒頭の問いを、もう一度");
  s.addNotes("冒頭の問いと対になる。ここが感情のピーク。間を長く取る。");

  // 68 — 3アクション
  s = L.light(pres);
  L.head(s, "明日からの3アクション", "この順番で、そのまま実行できます");
  L.checklist(s, [
    { title: "今週:担当者に「2週間、作業時間を記録してほしい」と伝える", body: "細かい分類は不要。何に何分かかったかだけ。" },
    { title: "今月:記録を見て、最初の1業務を社長が決める", body: "効果が大きく、着手しやすいものから。現場に選ばせない。" },
    { title: "来月:1業務・1か月・1人で試す。失敗しても構わないと伝える", body: "開始前と終了後の時間を必ず記録する。" },
  ], { y: 2.1, itemH: 1.05, gap: 0.26 });
  s.addShape("roundRect", { x: M, y: 5.95, w: CW, h: 0.85, rectRadius: 0.08, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText("3か月後には、自社に合うかどうかの答えが、数字で出ています。", {
    x: M + 0.35, y: 5.95, w: CW - 0.7, h: 0.85, fontSize: 15, bold: true, color: C.white, fontFace: F, margin: 0, valign: "middle",
  });
  s.addNotes("配布資料の記録シートと連動させる。ここで具体的な次アクションを渡し切る。");

  // 69 — I statement
  s = L.statement(pres, {
    kicker: "最後にひとつだけ",
    main: "最初の1業務を、\n今ここで決めてください。",
    sub: "会場を出てからでは、日常業務に戻って忘れてしまいます。\n経理の請求書処理でも、総務の議事録でも構いません。ひとつだけ選んでください。",
    iconName: "FiPlay",
  });
  s.addNotes("実際に10秒黙る。この沈黙が、行動につながるかどうかを分ける。");

  // 70 — 窓口
  s = L.light(pres);
  L.head(s, "ご相談窓口", "自社だけで抱え込まず、まずは相談から始めてください");
  L.cards(s, [
    { icon: "FiPhone", title: "主催者の経営相談窓口", body: "〇〇商工会議所 経営支援課\nTEL:000-0000-0000\n受付:平日 9:00〜17:00" },
    { icon: "FiUserCheck", title: "本日の講師への相談", body: "所属・部署名\nメール:info@example.jp\n※本日の資料もこちらからお送りします" },
    { icon: "FiClipboard", title: "配布資料について", body: "業務棚卸しシート\n社内ルールのひな形(A4・1枚)\n本日のスライド(抜粋版)" },
  ], { y: 2.0, h: 3.55 });
  s.addShape("roundRect", { x: M, y: 5.75, w: CW, h: 0.85, rectRadius: 0.08, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText("本日はありがとうございました。まずは「2週間の時間記録」から始めてみてください。", {
    x: M + 0.4, y: 5.75, w: CW - 0.8, h: 0.85, fontSize: 16, bold: true, color: C.white, fontFace: F, margin: 0, valign: "middle",
  });
  L.note(s, "※窓口・連絡先は、開催前に主催者および登壇者の情報へ差し替えてください。");
  s.addNotes("質疑応答へ。回答は短く、個別具体は窓口へ誘導する。");

  const out = process.argv[2] || "seminar.pptx";
  await pres.writeFile({ fileName: out });
  console.log("slides:", L.state.page, "->", out);
}

main().catch((e) => { console.error(e); process.exit(1); });
