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

// two-panel before/after showcase (left = 現状, right = AI後)
function twoPanel(s, left, right, { y = 1.95, h = 4.3 } = {}) {
  const gw = (CW - 0.7) / 2;
  [[left, M, false], [right, M + gw + 0.7, true]].forEach(([col, x, dark]) => {
    s.addShape("roundRect", {
      x, y, w: gw, h, rectRadius: 0.1,
      fill: { color: dark ? C.navy : C.tint }, line: { color: dark ? C.navy : C.line, width: 1 },
      shadow: L.shadow({ opacity: dark ? 0.18 : 0.1 }),
    });
    s.addShape("roundRect", {
      x: x + 0.32, y: y + 0.3, w: 1.95, h: 0.4, rectRadius: 0.2,
      fill: { color: dark ? C.ice : C.navy }, line: { color: dark ? C.ice : C.navy },
    });
    s.addText(col.label, {
      x: x + 0.32, y: y + 0.3, w: 1.95, h: 0.4, align: "center", valign: "middle",
      fontSize: 12.5, bold: true, color: dark ? C.navy : C.white, fontFace: F, margin: 0,
    });
    s.addText(col.meta, {
      x: x + 2.42, y: y + 0.3, w: gw - 2.74, h: 0.4, fontSize: 12,
      color: dark ? C.ice : C.muted, fontFace: F, margin: 0, valign: "middle",
    });
    s.addText(col.body, {
      x: x + 0.32, y: y + 0.9, w: gw - 0.64, h: h - 1.2,
      fontSize: L.fit(typeof col.body === "string" ? col.body : col.body.map(r => r.text).join(""), gw - 0.64, h - 1.2, col.size || 13, 1.4, 10.5),
      color: dark ? C.white : C.steel, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.4,
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

// banner strip used under card rows
function banner(s, text, { y, h = 0.8, dark = true, icon } = {}) {
  s.addShape("roundRect", {
    x: M, y, w: CW, h, rectRadius: 0.08,
    fill: { color: dark ? C.navy : C.tint }, line: { color: dark ? C.navy : C.line, width: 1 },
  });
  const tx = icon ? M + 1.15 : M + 0.4;
  if (icon) L.iconBadge(s, { x: M + 0.32, y: y + (h - 0.6) / 2, d: 0.6, name: icon, tone: dark ? "dark" : "light" });
  s.addText(text, {
    x: tx, y, w: CW - (tx - M) - 0.4, h, fontSize: 14.5, bold: true,
    color: dark ? C.white : C.navy, fontFace: F, margin: 0, valign: "middle",
  });
}

// ---------- deck ----------

async function main() {
  L.state.icons = await renderAll({ navy: C.navy, white: C.white, ice: C.ice });

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "セミナー資料";
  pres.title = "売る時間を、取り戻す。";

  let s;

  // ===== 第0部 オープニング =====

  // 1 — title
  s = L.dark(pres, { page: false });
  s.addImage({ data: L.icon("FiTrendingUp", "ice"), x: W - 4.2, y: 1.5, w: 3.6, h: 3.6, transparency: 82 });
  s.addText("中小企業の経営者向けセミナー", {
    x: M + 0.3, y: 1.85, w: 7.5, h: 0.4, fontSize: 15, bold: true, color: C.ice,
    fontFace: F, margin: 0, charSpacing: 2,
  });
  s.addText("売る時間を、\n取り戻す。", {
    x: M + 0.3, y: 2.45, w: 8.6, h: 2.3, fontSize: 46, bold: true, color: C.white,
    fontFace: F, margin: 0, valign: "middle", lineSpacingMultiple: 1.2,
  });
  s.addText("Claude Code で営業の事務作業を自動化する ─ 経営者が決める、はじめの一歩", {
    x: M + 0.3, y: 4.95, w: 9.4, h: 0.5, fontSize: 16, color: C.ice, fontFace: F, margin: 0,
  });
  s.addShape("line", { x: M + 0.3, y: 5.75, w: 3.2, h: 0, line: { color: "3A4585", width: 1 } });
  s.addText("所要 90分(質疑10分を含む)　|　登壇者名／所属・肩書き", {
    x: M + 0.3, y: 5.95, w: 9.0, h: 0.4, fontSize: 13, color: "9AA5CC", fontFace: F, margin: 0,
  });
  s.addNotes("開始の挨拶は30秒。ツールの説明から入らず、まず営業の時間の話から始める。");

  // 2 — Q
  s = L.question(pres,
    "御社の営業は、1日のうち\n何時間「売って」いますか?",
    "── 訪問と商談に使えている時間だけを数えてください");
  s.addNotes("いきなりツール名を出さない。営業の時間という自社の話から入る。");

  // 3 — 今日の3つ
  s = L.light(pres);
  L.head(s, "今日、持ち帰っていただく3つ", "この90分は「AIの勉強」ではなく、「営業の時間をどこから取り戻すか」の時間です");
  L.cards(s, [
    { num: "1", title: "営業の時間が\nどこへ消えているか", body: "売る時間より、資料作成と報告に使われている時間のほうが長い。その内訳を数字で共有します。" },
    { num: "2", title: "Claude Codeで\n何が自動化できるか", body: "リスト作成、提案書、日報集計。営業のどの作業が、どこまで手を離れるのかを具体的に見ます。" },
    { num: "3", title: "経営者は\n何を決めるのか", body: "現場に配って終わりにした会社は、ほぼ失敗します。社長が決めるべき3つを持ち帰っていただきます。" },
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
    { icon: "FiBriefcase", title: "中小企業の営業現場を支援", body: "卸売・商社・建設資材・製造業を中心に、営業まわりの業務の見直しを支援。" },
    { icon: "FiMonitor", title: "専門知識は前提にしません", body: "プログラミングの話はしません。マウスとキーボードで動かせる範囲だけを扱います。" },
    { icon: "FiUsers", title: "対象は「営業を増やせない」会社", body: "大企業のSFA導入事例ではなく、営業3〜20名規模の現実解を扱います。" },
  ], { y: 1.95, x: M + 4.6, w: CW - 4.6, rowH: 1.35, gap: 0.2 });
  s.addNotes("30秒で切り上げる。技術者向けではないことをここで明言しておく。");

  // ===== 第1部 =====

  // 5 — divider
  s = L.divider(pres, {
    part: "第 1 部", title: "営業の時間は、\nどこへ消えているのか",
    sub: "まず、自社の営業が何に時間を使っているのかを共有します。",
    iconName: "FiClock",
  });
  s.addNotes("第1部は危機感のパート。ツールの話はまだ一切しない。");

  // 6 — 1日の使われ方
  s = L.light(pres);
  L.head(s, "営業1人の1日は、こう使われている", "8時間のうち、実際に「売っている」のは3時間ほどです");
  chartTakeaway(s, {
    type: "bar",
    data: [{
      name: "1日あたりの時間(時間)",
      labels: ["訪問・商談", "移動", "資料・見積の作成", "報告・システム入力", "情報を探す", "社内調整・その他"],
      values: [3.0, 1.2, 1.5, 1.0, 0.8, 0.5],
    }],
    axisNote: "営業1人の1日(8時間)の内訳 / 時間",
    opts: { barDir: "bar", showValue: true, dataLabelPosition: "outEnd", barGapWidthPct: 45, chartColors: [C.navy] },
    take: {
      icon: "FiClock", title: "売る時間は、\n1日の4割以下",
      body: "残りの5時間は、やらなければ回らないが、それ自体では1円も生まない作業です。ここを削らない限り、訪問件数は増えません。",
    },
  });
  L.note(s, "※業務実態のイメージ図です。自社の営業日報の実績値に置き換えてご覧ください。");
  s.addNotes("会場に「うちの営業の訪問件数、月何件ですか」と軽く投げかけてもよい。");

  // 7 — Q
  s = L.question(pres, "その残りの5時間は、\n何に使われていますか?");
  s.addNotes("答えさせない。自社の営業の顔を思い浮かべてもらう時間にする。");

  // 8 — 年間コスト
  s = L.light(pres);
  L.head(s, "事務作業に消えている、本当の金額", "営業の時間単価で計算すると、無視できない額になります");
  L.stats(s, [
    { label: "営業1人あたりの事務時間", value: "1,100", unit: "時間/年", sub: "1日5時間 × 年間220日。営業活動以外に使われている時間の合計" },
    { label: "人件費換算", value: "390", unit: "万円/年", sub: "時間単価3,500円で換算。営業1名あたりの「売っていない時間」のコスト" },
    { label: "営業5名の会社なら", value: "1,950", unit: "万円/年", sub: "この金額が、受注につながらない作業に毎年消えている計算です" },
  ], { y: 2.15, h: 2.6 });
  banner(s, "問題は営業の働きが悪いことではありません。売る以外の仕事が、営業に積み上がっていることです。", { y: 5.05, h: 1.05, icon: "FiDollarSign" });
  L.note(s, "※試算例です。自社の営業人数・人件費に置き換えてご覧ください。");
  s.addNotes("自社の数字に置き換えてもらうため、ここは少しゆっくり話す。");

  // 9 — 属人化
  s = L.light(pres);
  L.head(s, "しかも、そのやり方は共有されない", "できる営業のノウハウが、本人の中にしか残らない構造です");
  L.rows(s, [
    { icon: "FiFileText", title: "提案書は、各自のパソコンの中", body: "うまくいった提案書がどこにあるか誰も知らない。毎回ゼロから作り直している。" },
    { icon: "FiMessageCircle", title: "刺さる言い回しは、本人の頭の中", body: "トップ営業がなぜ受注できるのかを、若手が言語化して学べる形になっていない。" },
    { icon: "FiUserX", title: "辞めた瞬間、すべて消える", body: "顧客との経緯も、断られた理由も、引き継ぎ資料には残らない。ゼロからやり直しになる。" },
  ], { y: 2.0, rowH: 1.3, gap: 0.24 });
  banner(s, "属人化は「人の問題」ではなく、記録と共有の仕組みがないという構造の問題です。", { y: 6.0, h: 0.72 });
  s.addNotes("「引き継ぎ書はありますか」と一言添えると、会場の顔が変わる。");

  // 10 — Q
  s = L.question(pres, "トップ営業が辞めたら、\nそのノウハウはどこに残りますか?");
  s.addNotes("第1部の危機感のピーク。具体的な人の顔が浮かぶ問い方をする。");

  // 11 — 採用も難しい
  s = L.light(pres);
  L.head(s, "そして、営業も採れなくなっている", "人を増やして解決する道は、年々細くなっています");
  chartTakeaway(s, {
    type: "line",
    data: [{
      name: "営業職の採用が難しいと答えた中小企業の割合(%)",
      labels: ["2016", "2018", "2020", "2022", "2024", "2026"],
      values: [44, 52, 39, 50, 57, 62],
    }],
    axisNote: "営業職の採用が難しいと答えた中小企業の割合(%)",
    opts: { lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 8, valAxisMaxVal: 80, valAxisMinVal: 0 },
    take: {
      icon: "FiUsers", title: "増やす前提の計画は、\nもう立たない",
      body: "採用できても、育つまでに2〜3年かかります。今いる営業が売る時間を取り戻すほうが、確実で早い打ち手です。",
    },
  });
  L.note(s, "※推移イメージ。登壇時は日本商工会議所LOBO調査・中小企業白書の最新値に差し替えてください。");
  s.addNotes("採用の話は共感が取りやすい。ただし長引かせない。");

  // 12 — AIは試した
  s = L.light(pres);
  L.head(s, "「AIはもう試した」── それでも時間は戻らなかった", "多くの会社が、ここで止まっています");
  L.cards(s, [
    { icon: "FiMessageCircle", title: "聞くだけで終わった", body: "文章は作ってくれる。でも、その文章を貼り付けて、体裁を整えて、送るのは結局人。作業は減らなかった。" },
    { icon: "FiRepeat", title: "毎回、同じ説明が必要", body: "自社の商品も、顧客の事情も知らない。使うたびに背景から説明していたら、自分で書いたほうが早い。" },
    { icon: "FiFolder", title: "資料もデータも渡せない", body: "顧客リストも過去の提案書も社内にある。AIはそれを見られないので、机上の一般論しか返ってこない。" },
  ], { y: 2.05, h: 3.4 });
  banner(s, "問題はAIの性能ではなく、AIが「作業の外」にいたことでした。", { y: 5.7, h: 0.8 });
  s.addNotes("すでに試して失望した経営者が会場にいる前提で話す。ここで共感を取る。");

  // 13 — チャットAIの限界(図)
  s = L.light(pres);
  L.head(s, "チャット型AIに残る、4つの手作業", "聞いて答えをもらうまでの往復が、そのまま時間になっています");
  ["資料を開いて内容をコピーする", "AIの画面に貼り付けて質問する", "返ってきた答えをコピーする", "資料に貼り直して体裁を整える"].forEach((t, i) => {
    const x = M + i * ((CW - 0.9) / 4 + 0.3);
    const w = (CW - 0.9) / 4;
    s.addShape("roundRect", { x, y: 2.2, w, h: 2.5, rectRadius: 0.1, fill: { color: C.tint }, line: { color: C.line, width: 1 }, shadow: L.shadow() });
    s.addShape("ellipse", { x: x + (w - 0.6) / 2, y: 2.5, w: 0.6, h: 0.6, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText(String(i + 1), { x: x + (w - 0.6) / 2, y: 2.5, w: 0.6, h: 0.6, align: "center", valign: "middle", fontSize: 20, bold: true, color: C.white, fontFace: F, margin: 0 });
    s.addText(t, { x: x + 0.25, y: 3.3, w: w - 0.5, h: 1.2, align: "center", fontSize: 13.5, color: C.navy, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.3 });
    if (i < 3) s.addImage({ data: L.icon("FiArrowRight", "navy"), x: x + w + 0.04, y: 3.34, w: 0.22, h: 0.22 });
  });
  banner(s, "AIが賢くなっても、この1・2・4は人の手作業のまま残ります。ここが削れないと、時間は戻りません。", { y: 5.15, h: 0.9 });
  L.note(s, "※チャット型AIを業務で使うときの一般的な流れです。");
  s.addNotes("会場の多くが体験している往復。実際に手で動かす仕草を入れると伝わる。");

  // 14 — つなぎ目
  s = L.light(pres);
  L.head(s, "残っているのは「つなぎ目」の作業", "考える仕事ではなく、運ぶ仕事が残っています");
  L.compare(s,
    {
      label: "AIが担当", tone: "after", head: "考える・書く部分",
      items: [
        "文章を考える",
        "要点にまとめる",
        "表現を整える",
        "案を複数出す",
        "専門用語なしで、日本語のまま指示できる",
      ],
    },
    {
      label: "人が担当", tone: "before", head: "運ぶ・つなぐ部分",
      items: [
        "資料を探して開く",
        "内容をコピーして貼り付ける",
        "結果をファイルに書き戻す",
        "体裁を整えて保存する",
        "毎回、同じ手順を繰り返す",
      ],
    },
    { y: 1.95 }
  );
  s.addNotes("「運ぶ仕事」という言い方が、経営者にはいちばん伝わる。");

  // 15 — I statement
  s = L.statement(pres, {
    kicker: "必要だったもの",
    main: "答えるAIではなく、\n作業をやり切るAI。",
    sub: "資料を自分で開き、書き出し、保存し、次の作業まで進める。\n人が「運ぶ」必要のないAIが出てきたことが、これまでとの決定的な違いです。",
    iconName: "FiZap",
  });
  s.addNotes("ここで一拍置く。第2部でその正体を明かす。");

  // 16 — まとめ
  s = L.light(pres);
  L.head(s, "第1部のまとめ", null);
  L.rows(s, [
    { icon: "FiClock", title: "営業が売っている時間は1日の4割以下", body: "残りは資料作成・報告・情報探し。やらなければ回らないが、売上には直結しない作業です。" },
    { icon: "FiUserX", title: "そのやり方は本人の中にしか残らない", body: "提案書もノウハウも共有されず、辞めた瞬間にゼロに戻ります。" },
    { icon: "FiUsers", title: "営業を増やして解決する道は細くなっている", body: "採用も育成も年々難しくなっています。今いる営業の時間を取り戻すほうが早い。" },
    { icon: "FiZap", title: "チャット型AIでは「運ぶ作業」が残った", body: "必要なのは、資料を自分で開いて最後まで進めるAIです。" },
  ], { y: 1.72, rowH: 1.14, gap: 0.16 });
  s.addNotes("4点を30秒で復唱。長く語らずに第2部へ。");

  // ===== 第2部 =====

  // 17 — divider
  s = L.divider(pres, {
    part: "第 2 部", title: "Claude Code とは\n何なのか",
    sub: "技術の話はしません。「自社の営業の何が変わるのか」だけを扱います。",
    iconName: "FiCpu",
  });
  s.addNotes("第2部は理解のパート。専門用語を使わないことを冒頭で宣言する。");

  // 18 — Q
  s = L.question(pres, "AIが、自分のパソコンの中で\n手を動かすとしたら?",
    "── 聞くのではなく、やってもらう");
  s.addNotes("チャットAIとの違いを、この問いで先に印象づける。");

  // 19 — 正体
  s = L.light(pres);
  L.head(s, "Claude Code の正体は、この3つ", "難しく考える必要はありません。できることは、突き詰めるとこれだけです");
  L.cards(s, [
    { icon: "FiFolder", title: "読む", body: "パソコンの中のファイルを自分で開いて読む。\n\n例:過去の提案書50件、顧客リスト、商談メモ" },
    { icon: "FiEdit3", title: "作る", body: "文書やデータを作り、ファイルとして保存する。\n\n例:提案書、見積、メール文面、集計表" },
    { icon: "FiLink", title: "つなぐ", body: "社内で使っているサービスと接続して情報をやり取りする。\n\n例:表計算、共有フォルダ、チャット、カレンダー" },
  ], { y: 2.0, h: 4.15 });
  s.addNotes("「読む・作る・つなぐ」の3語で覚えてもらう。以降この3語で説明する。");

  // 20 — 比較
  s = L.light(pres);
  L.head(s, "チャット型AIとの、決定的な違い", "同じAIでも、できる仕事の範囲がまったく違います");
  L.compare(s,
    {
      label: "チャット型", tone: "before", head: "これまでのAI",
      items: [
        "画面に貼り付けたものだけを見る",
        "答えを返すところまでが仕事",
        "ファイルの保存は人がやる",
        "毎回、背景から説明が必要",
        "社内のデータは渡せない",
      ],
    },
    {
      label: "Claude Code", tone: "after", head: "作業をやり切るAI",
      items: [
        "フォルダの中を自分で探して読む",
        "ファイルを作って保存するまで進める",
        "決めた手順をそのまま実行する",
        "自社のルールを一度覚えさせれば再説明は不要",
        "社内サービスと接続できる",
      ],
    },
    { y: 1.95 }
  );
  s.addNotes("左右を指差しながら比較する。ここが第2部でいちばん重要な1枚。");

  // 21 — なぜ営業に効くか
  s = L.light(pres);
  L.head(s, "「Code」という名前だが、営業事務にこそ効く", "もともとは技術者向けですが、効く仕事の構造は営業事務と同じです");
  L.rows(s, [
    { icon: "FiFileText", title: "決まった形式の書類を、大量に作る仕事", body: "提案書も見積も報告書も、形が決まっていて中身だけが変わる。もっとも自動化しやすい形です。" },
    { icon: "FiSearch", title: "過去の資料から探して、引っ張ってくる仕事", body: "似た案件の提案書、以前の見積条件、前回の議事録。探す作業がそのまま消えます。" },
    { icon: "FiRepeat", title: "毎週・毎月、同じ手順で繰り返す仕事", body: "日報の集計、週次の売上まとめ、月末の実績報告。一度手順を決めれば、以後は自動で回せます。" },
  ], { y: 2.0, rowH: 1.3, gap: 0.24 });
  banner(s, "「文字と数字を、決まった手順で扱う仕事」── 営業事務はまさにこれです。", { y: 6.0, h: 0.72 });
  s.addNotes("名前で敬遠されないよう、ここは丁寧に。技術者専用ではないと明言する。");

  // 22 — 使える場所
  s = L.light(pres);
  L.head(s, "使える場所は、5つ", "自分のパソコンでも、外出先のスマホでも動きます");
  L.cards(s, [
    { icon: "FiMonitor", title: "デスクトップ", body: "Windows / Mac\n\n通常のソフトと同じ。画面を見ながら作業を確認できる" },
    { icon: "FiCloud", title: "ブラウザ", body: "claude.ai/code\n\nインストール不要。長い作業を任せて後で結果を見る" },
    { icon: "FiSmartphone", title: "スマホ", body: "iOS / Android\n\n移動中に指示を出し、戻ったら結果を確認する" },
    { icon: "FiMessageCircle", title: "チャット連携", body: "Slack など\n\nチームのチャットから直接、作業を依頼できる" },
    { icon: "FiTerminal", title: "ターミナル", body: "技術者向け(開発用)\n\n経営者・営業が使う必要はありません" },
  ], { y: 1.88, h: 3.8, gap: 0.28 });
  banner(s, "営業と経営者が使うのは、左の3つだけです。黒い画面を触る必要はありません。", { y: 5.88, h: 0.82 });
  L.note(s, "※提供形態は変更される場合があります。登壇前に公式情報をご確認ください。");
  s.addNotes("「黒い画面」への抵抗をここで完全に外す。いちばん質問が出る箇所。");

  // 23 — 非技術者の入口
  s = L.light(pres);
  L.head(s, "営業が使うのは、この2つだけ", "特別な知識は要りません。日本語で指示するだけです");
  twoPanel(s,
    {
      label: "デスクトップアプリ", meta: "自席のパソコンで、結果を見ながら",
      body: "普通のソフトと同じようにインストールして起動します。\n\n「このフォルダの提案書を読んで、A社向けの提案書を作って」\n\n── 日本語でこう頼むだけです。作られたファイルはそのまま自分のパソコンに保存されます。作業の途中経過も画面で確認できるので、任せきりにはなりません。",
    },
    {
      label: "ブラウザ", meta: "外出先から、時間のかかる作業を",
      body: "インストール不要。ブラウザを開いてログインするだけです。\n\n「先週の営業日報30件を読んで、要注意の案件を3つ挙げて」\n\n── 時間のかかる作業を出先から投げておき、戻ったときに結果を確認する使い方に向いています。スマートフォンからも同じことができます。",
    },
    { y: 1.95, h: 4.3 }
  );
  s.addNotes("実際の画面を見せられるなら、ここで実演すると効果が高い。");

  // 24 — MCPで社内とつながる
  s = L.light(pres);
  L.head(s, "社内で使っているサービスと、つながる", "AIを社内の情報につなぐ共通の仕組み(MCP)が用意されています");
  const hub = { x: 5.6, y: 3.15, w: 2.2, h: 2.2 };
  s.addShape("roundRect", { x: hub.x, y: hub.y, w: hub.w, h: hub.h, rectRadius: 0.15, fill: { color: C.navy }, line: { color: C.navy }, shadow: L.shadow({ opacity: 0.2 }) });
  s.addShape("ellipse", { x: hub.x + 0.75, y: hub.y + 0.42, w: 0.7, h: 0.7, fill: { color: C.ice }, line: { color: C.ice } });
  s.addImage({ data: L.icon("FiCpu", "navy"), x: hub.x + 0.93, y: hub.y + 0.6, w: 0.34, h: 0.34 });
  s.addText("Claude Code", { x: hub.x + 0.15, y: hub.y + 1.25, w: hub.w - 0.3, h: 0.5, align: "center", fontSize: 16, bold: true, color: C.white, fontFace: F, margin: 0 });
  s.addText("指示は日本語で", { x: hub.x + 0.15, y: hub.y + 1.72, w: hub.w - 0.3, h: 0.35, align: "center", fontSize: 11.5, color: C.ice, fontFace: F, margin: 0 });
  const spokes = [
    { icon: "FiFolder", t: "共有フォルダ", d: "提案書・見積・契約書", x: M, y: 2.0 },
    { icon: "FiGrid", t: "表計算・顧客リスト", d: "案件管理・売上データ", x: M, y: 3.35 },
    { icon: "FiMail", t: "メール・チャット", d: "顧客とのやり取り", x: M, y: 4.7 },
    { icon: "FiCalendar", t: "カレンダー", d: "訪問予定・商談履歴", x: 8.15, y: 2.0 },
    { icon: "FiDatabase", t: "販売管理システム", d: "受注・在庫・出荷", x: 8.15, y: 3.35 },
    { icon: "FiMessageCircle", t: "社内チャット", d: "報告・情報共有", x: 8.15, y: 4.7 },
  ];
  spokes.forEach((sp) => {
    s.addShape("roundRect", { x: sp.x, y: sp.y, w: 4.35, h: 1.15, rectRadius: 0.08, fill: { color: C.tint }, line: { color: C.line, width: 1 } });
    L.iconBadge(s, { x: sp.x + 0.28, y: sp.y + 0.27, d: 0.6, name: sp.icon, tone: "light" });
    s.addText(sp.t, { x: sp.x + 1.05, y: sp.y + 0.18, w: 3.05, h: 0.4, fontSize: 14.5, bold: true, color: C.navy, fontFace: F, margin: 0, valign: "middle" });
    s.addText(sp.d, { x: sp.x + 1.05, y: sp.y + 0.58, w: 3.05, h: 0.38, fontSize: 12, color: C.steel, fontFace: F, margin: 0, valign: "middle" });
  });
  s.addImage({ data: L.icon("FiArrowRight", "navy"), x: 4.95, y: 4.14, w: 0.28, h: 0.28 });
  s.addImage({ data: L.icon("FiArrowRight", "navy"), x: 8.1, y: 4.14, w: 0.28, h: 0.28 });
  L.note(s, "※接続できる範囲は利用中のサービス・社内の許可設定によります。すべてが自動でつながるわけではありません。");
  s.addNotes("何でもつながると誤解されないよう、注記は口頭でも触れる。");

  // 25 — 費用感
  s = L.light(pres);
  L.head(s, "費用感の現実", "投資判断の話ではなく、月々の消耗品に近い金額です");
  chartTakeaway(s, {
    type: "bar",
    data: [{
      name: "月額コスト(円)",
      labels: ["Claude 個人プラン(1名)", "携帯電話の通信費(1名)", "営業車のリース(1台)", "営業社員1名(月額人件費)"],
      values: [3000, 6000, 45000, 550000],
    }],
    axisNote: "月額コストの比較(円)",
    opts: { barDir: "bar", showValue: true, dataLabelPosition: "outEnd", barGapWidthPct: 55, chartColors: [C.navy] },
    take: {
      icon: "FiDollarSign", title: "決裁が要る額\nですらない",
      body: "個人プランは月額20ドル(約3,000円)から。チームで使う場合も1席あたりの月額課金です。障害はコストではありません。",
    },
  });
  L.note(s, "※2026年8月時点の一般的な水準からの試算例。プラン内容・価格は変更されるため、登壇前に公式情報をご確認ください。");
  s.addNotes("「高いから無理」という言い訳をここで完全に外す。価格は必ず最新を確認。");

  // 26 — まとめ
  s = L.light(pres);
  L.head(s, "第2部のまとめ", null);
  L.rows(s, [
    { icon: "FiZap", title: "Claude Codeは「読む・作る・つなぐ」を代行する", body: "答えを返すだけでなく、ファイルを作って保存するところまで進めます。" },
    { icon: "FiFileText", title: "決まった形式の書類を大量に作る仕事に効く", body: "提案書・見積・報告書。営業事務はまさにこの形です。" },
    { icon: "FiMonitor", title: "使うのはデスクトップアプリとブラウザ", body: "黒い画面は不要。指示は日本語です。" },
    { icon: "FiDollarSign", title: "止めているのはコストではなく、判断", body: "月数千円から。決裁が要らない金額で最初の検証はできます。" },
  ], { y: 1.72, rowH: 1.14, gap: 0.16 });
  s.addNotes("第3部で「では営業のどの仕事か」に落とすことを予告する。");

  // ===== 第3部 =====

  // 27 — divider
  s = L.divider(pres, {
    part: "第 3 部", title: "営業の、\nどの仕事が変わるのか",
    sub: "リスト作成から提案・報告まで、営業プロセスごとに具体的に見ていきます。",
    iconName: "FiMap",
  });
  s.addNotes("ここからが本題。自社の営業を思い浮かべながら聞いてもらうよう促す。");

  // 28 — プロセスマップ
  s = L.light(pres);
  L.head(s, "営業プロセスマップ", "この5工程を順番に見ていきます");
  L.cards(s, [
    { icon: "FiSearch", title: "①リスト作成", body: "見込み客の抽出\n企業情報の調査\n担当者の特定\n優先順位づけ" },
    { icon: "FiMail", title: "②アプローチ", body: "メール文面作成\n電話トークの準備\n案内資料の作成\n返信対応" },
    { icon: "FiMessageCircle", title: "③商談", body: "商談メモ\nお礼メール\n議事録の作成\n次アクション整理" },
    { icon: "FiFileText", title: "④提案・見積", body: "提案書の作成\n見積書の作成\n過去案件の参照\n条件の確認" },
    { icon: "FiBarChart2", title: "⑤報告・分析", body: "営業日報\n週次・月次の集計\n案件の進捗管理\n失注理由の分析" },
  ], { y: 1.95, h: 3.6, gap: 0.28 });
  banner(s, "どの工程も、中身は「集める・書く・まとめる・つなぐ」の組み合わせでできています。", { y: 5.75, h: 0.85 });
  s.addNotes("最後のバナーが第3部の結論の伏線。ここでは軽く触れる程度に。");

  // 29 — ①リスト作成の現状
  s = L.light(pres);
  L.head(s, "①リスト作成:探すだけで半日が終わる", "売る前の準備に、想像以上の時間が使われています");
  L.rows(s, [
    { icon: "FiSearch", title: "1社ずつ調べて、1行ずつ入力する", body: "業種・規模・所在地・担当者名。100社分を手で調べれば、それだけで2日仕事になります。" },
    { icon: "FiRepeat", title: "調べた内容が、共有されない", body: "同じ会社を別の営業がまた調べている。個人のパソコンに散らばったまま蓄積されません。" },
    { icon: "FiFilter", title: "優先順位がつけられない", body: "リストはあるが、どこから当たるべきかの判断材料がない。結局、上から順に電話している。" },
  ], { y: 2.0, rowH: 1.3, gap: 0.24 });
  banner(s, "調べる・整える・並べ替える ── すべてClaude Codeが得意とする作業です。", { y: 6.0, h: 0.72 });
  s.addNotes("リスト作成は成果が見えやすいので、最初の1業務の候補として有力。");

  // 30 — ① B/A
  s = L.light(pres);
  L.head(s, "①リスト作成 ─ Before / After", "同じ人が、同じ準備を、短い時間で終える形に変わります");
  L.compare(s,
    {
      label: "Before", tone: "before", head: "いまのやり方",
      items: [
        "1社ずつ検索して情報を書き写す",
        "業種や規模の条件で手作業で絞り込む",
        "表記ゆれや重複をあとから直す",
        "優先順位は担当者の感覚で決める",
        "100社分で丸2日かかる",
      ],
    },
    {
      label: "After", tone: "after", head: "Claude Codeを使った形",
      items: [
        "条件を伝えれば、候補一覧を作って表に整理",
        "既存顧客・過去の失注先を自動で突き合わせ",
        "表記ゆれと重複をまとめて整える",
        "受注実績に近い特徴の会社から順に並べる",
        "人がやるのは、最終確認と当たる順の判断",
      ],
    },
    { y: 1.95 }
  );
  s.addNotes("「人がやるのは確認と判断だけ」という形になることを強調する。");

  // 31 — ②アプローチの現状
  s = L.light(pres);
  L.head(s, "②アプローチ:文面を考えるだけで手が止まる", "1通10分でも、積み上がると大きな時間になります");
  L.rows(s, [
    { icon: "FiEdit3", title: "1通ずつ、文面を考え直している", body: "相手の業種や課題に合わせて書き直す。100社に送るなら、それだけで2日分の作業です。" },
    { icon: "FiClock", title: "後回しにされ、送られない", body: "文面を考える手間があるため着手が遅れ、結局送らないまま案件が流れていく。" },
    { icon: "FiFileText", title: "うまくいった文面が残らない", body: "反応が良かったメールも、個人の送信済みフォルダに埋もれて共有されない。" },
  ], { y: 2.0, rowH: 1.3, gap: 0.24 });
  banner(s, "文面作成は、Claude Codeがもっとも早く効果を出す業務のひとつです。", { y: 6.0, h: 0.72 });
  s.addNotes("「送れていない」ことによる機会損失を意識させる。");

  // 32 — ② B/A
  s = L.light(pres);
  L.head(s, "②アプローチ ─ Before / After", "文面を考える時間が、送る件数に変わります");
  L.compare(s,
    {
      label: "Before", tone: "before", head: "いまのやり方",
      items: [
        "1通ずつ、相手に合わせて書き直す",
        "過去のうまくいった文面を探せない",
        "返信対応の文面も毎回考える",
        "案内資料は都度、体裁から作る",
        "手が回らず、送信自体が後回しになる",
      ],
    },
    {
      label: "After", tone: "after", head: "Claude Codeを使った形",
      items: [
        "業種と課題を伝えれば、案を複数出させる",
        "反応の良かった文面を読み込ませて型にする",
        "返信のパターン別に下書きを用意させる",
        "案内資料は雛形から自動で差し込み",
        "人がやるのは、選んで直して送ることだけ",
      ],
    },
    { y: 1.95 }
  );
  s.addNotes("「人がやるのは選ぶこと」。判断は人に残ることを毎回添える。");

  // 33 — 実例:商談メモ
  s = L.light(pres);
  L.head(s, "実例:商談メモが、3点セットになる", "手書きのメモを渡すだけで、この形まで自動で整理されます");
  twoPanel(s,
    {
      label: "商談メモ", meta: "訪問後に走り書きした10行", size: 12,
      body: "8/5 A建設 田中部長\n・現場3件で足場材が不足\n・今は他社から購入、納期に不満\n・価格は既存より高くても納期優先\n・9月着工分から切替を検討\n・見積は8月中旬までにほしい\n・決裁は社長、月末の会議で\n・競合はB社(単価は安いが在庫薄)",
    },
    {
      label: "自動で出てくるもの", meta: "作成時間:約3分",
      body: [
        { text: "① お礼メールの下書き\n", options: { bold: true, color: C.ice } },
        { text: "納期の課題に触れ、8月中旬の見積提出を明記した文面\n\n", options: { color: C.white } },
        { text: "② 商談議事録\n", options: { bold: true, color: C.ice } },
        { text: "課題・要望・決裁ルート・競合状況を項目別に整理\n\n", options: { color: C.white } },
        { text: "③ 次アクション一覧\n", options: { bold: true, color: C.ice } },
        { text: "・見積作成(担当:自分 / 期限:8月15日)\n・在庫と納期の社内確認(担当:業務課 / 今週中)\n・月末会議前のフォロー連絡(8月25日)", options: { color: C.white } },
      ],
    },
    { y: 1.95, h: 4.3 }
  );
  s.addNotes("実物を見せるのがいちばん早い。可能なら当日その場で実演してもよい。");

  // 34 — ③提案・見積の現状
  s = L.light(pres);
  L.head(s, "③提案・見積:1件2時間の常態化", "営業がもっとも時間を取られている工程です");
  L.rows(s, [
    { icon: "FiSearch", title: "似た案件の提案書を探すところから始まる", body: "どこに保存したか分からず、探すだけで30分。見つからなければゼロから作り直す。" },
    { icon: "FiEdit3", title: "流用して、体裁を整え直す", body: "前の顧客名が残っていないか確認しながら、1ページずつ直していく作業。" },
    { icon: "FiAlertTriangle", title: "条件の転記ミスが後で問題になる", body: "見積の単価や納期を写し間違え、受注後にトラブルになるケースが後を絶たない。" },
  ], { y: 2.0, rowH: 1.3, gap: 0.24 });
  banner(s, "探す・流用する・転記する。この3つが、そのまま削減対象になります。", { y: 6.0, h: 0.72 });
  s.addNotes("転記ミスの話は共感が強い。実例があれば1つ挟むとよい。");

  // 35 — ③ B/A
  s = L.light(pres);
  L.head(s, "③提案・見積 ─ Before / After", "提案の質ではなく、提案にかかる時間が変わります");
  L.compare(s,
    {
      label: "Before", tone: "before", head: "いまのやり方",
      items: [
        "似た案件の提案書を探すところから始める",
        "前回ファイルを開いて1ページずつ直す",
        "見積条件を手で転記する",
        "作成に1件あたり約2時間かかる",
        "忙しい週は、提案そのものが出せない",
      ],
    },
    {
      label: "After", tone: "after", head: "Claude Codeを使った形",
      items: [
        "過去の提案書を読み込ませ、近い案件を提示",
        "顧客名・条件を差し替えた案を自動で作成",
        "見積条件は元データから転記させ、ミスを防ぐ",
        "たたき台まで15分。人は中身の判断に集中",
        "提案件数そのものを増やせる",
      ],
    },
    { y: 1.95 }
  );
  s.addNotes("「提案件数が増える」ことが受注に直結する点を強調する。");

  // 36 — ④報告の現状
  s = L.light(pres);
  L.head(s, "④報告・集計:書く側も、読む側も報われない", "毎日書かせているが、経営判断には使えていません");
  L.rows(s, [
    { icon: "FiEdit3", title: "営業は、毎日30分かけて日報を書く", body: "5名なら月に50時間。それだけの時間をかけて書かせています。" },
    { icon: "FiFolder", title: "溜まるだけで、読まれない", body: "件数が多すぎて全部は読めない。結局、気になった案件だけを口頭で聞いている。" },
    { icon: "FiBarChart2", title: "集計は月末に手作業でまとめる", body: "表計算に打ち直して合計する。出てきた頃には、もう手を打てない時期になっている。" },
  ], { y: 2.0, rowH: 1.3, gap: 0.24 });
  banner(s, "書かせるのをやめるのではなく、読んでまとめる側を自動化するのが正解です。", { y: 6.0, h: 0.72 });
  s.addNotes("「日報をやめよう」という結論に飛ばないよう注意して話す。");

  // 37 — ④ B/A
  s = L.light(pres);
  L.head(s, "④報告・集計 ─ Before / After", "溜めるだけの日報が、手を打つための材料に変わります");
  L.compare(s,
    {
      label: "Before", tone: "before", head: "いまのやり方",
      items: [
        "日報は書かせているが読み切れない",
        "危ない案件は、報告されて初めて気づく",
        "集計は月末に手作業で打ち直す",
        "数字が出る頃には手遅れになっている",
        "失注理由は記録されず、次に活かせない",
      ],
    },
    {
      label: "After", tone: "after", head: "Claude Codeを使った形",
      items: [
        "全件を読み込ませ、要点だけを一覧化",
        "停滞している案件・競合が出た案件を抽出",
        "集計は自動。毎週決まった時間に出力",
        "週の途中で手を打てる状態になる",
        "失注理由を分類し、傾向として蓄積できる",
      ],
    },
    { y: 1.95 }
  );
  s.addNotes("「週の途中で手を打てる」が、経営者にとって最大の価値。");

  // 38 — 実例:週次レポート
  s = L.light(pres);
  L.head(s, "実例:月曜の朝、要注意案件が机の上にある", "毎週決まった時間に自動で動かす設定もできます");
  twoPanel(s,
    {
      label: "入力", meta: "先週1週間の営業日報 30件", size: 12.5,
      body: "営業5名 × 週6件 = 30件の日報。\n\nそれぞれ数百文字の自由記述で、案件名・訪問先・所感が混在した状態のまま。\n\nこれまでは、全部読む時間がないため、気になった数件だけを口頭で確認していました。",
    },
    {
      label: "月曜8時に自動出力", meta: "作成時間:自動(所要3分)",
      body: [
        { text: "要注意案件 3件\n", options: { bold: true, color: C.ice } },
        { text: "・A建設:2週間動きなし。競合B社が接触\n・C工業:見積提出後7日、反応なし\n・D商店:担当者交代、決裁ルート不明\n\n", options: { color: C.white } },
        { text: "今週の数字\n", options: { bold: true, color: C.ice } },
        { text: "訪問28件(前週比 +3) / 見積提出9件 / 受注3件\n\n", options: { color: C.white } },
        { text: "気づき\n", options: { bold: true, color: C.ice } },
        { text: "納期を理由にした失注が今月3件。在庫方針の見直しを検討", options: { color: C.white } },
      ],
    },
    { y: 1.95, h: 4.3 }
  );
  s.addNotes("「毎週勝手に出てくる」ことのインパクトを伝える。ここは反応が大きい。");

  // 39 — ⑤フォロー
  s = L.light(pres);
  L.head(s, "⑤フォロー・失注分析:いちばん放置されている工程", "受注しなかった案件から学べていない会社がほとんどです");
  L.rows(s, [
    { icon: "FiClock", title: "フォローのタイミングを逃す", body: "見積を出したまま2週間。思い出したときには、すでに他社で決まっている。" },
    { icon: "FiSlash", title: "失注理由が「価格」で片づけられる", body: "本当は納期や対応の遅さが理由でも、記録が残らないため改善につながらない。" },
    { icon: "FiRepeat", title: "過去の顧客に、二度と当たらない", body: "1年前に断られた会社の事情が変わっていても、掘り起こす仕組みがない。" },
  ], { y: 2.0, rowH: 1.3, gap: 0.24 });
  banner(s, "記録を読んで、拾って、知らせる ── これも「読む・まとめる・つなぐ」の作業です。", { y: 6.0, h: 0.72 });
  s.addNotes("失注分析は経営者の関心が高い。ここで身を乗り出す会場も多い。");

  // 40 — Q
  s = L.question(pres, "この5工程のうち、御社の営業が\nいちばん時間を使っているのはどれですか?",
    "── リスト作成 / アプローチ / 商談 / 提案・見積 / 報告");
  s.addNotes("直前のプロセスマップに戻す。ここで1つ選んでもらうことが第6部の伏線になる。");

  // 41 — 4つの型
  s = L.light(pres);
  L.head(s, "効く仕事の型は、4つしかない", "工程が違っても、軽くなる作業の構造は同じです");
  L.cards(s, [
    { icon: "FiSearch", title: "集める", body: "顧客情報・過去案件・条件を探す時間。\n読み込ませておけば、探す作業自体がなくなります。" },
    { icon: "FiEdit3", title: "書く", body: "メール・提案書・報告書を書く時間。\n下書きを任せ、人は確認と判断に集中します。" },
    { icon: "FiFilter", title: "まとめる", body: "日報や数字を集計する時間。\n読んで要点を出す作業をそのまま任せられます。" },
    { icon: "FiLink", title: "つなぐ", body: "資料と資料の間を運ぶ時間。\nファイルの読み書きまで進むので、往復が消えます。" },
  ], { y: 2.0, h: 3.75 });
  banner(s, "自社の営業の1日を、この4つの言葉で棚卸ししてください。それが第5部でお話しする最初の一歩です。", { y: 5.95, h: 0.75 });
  s.addNotes("この4語は、配布資料の見出しと揃えておくとよい。");

  // 42 — 積み上げ
  s = L.light(pres);
  L.head(s, "営業5名で積み上げると、何人分になるか", "工程ごとの削減時間を合算した試算です");
  chartTakeaway(s, {
    type: "bar",
    data: [{
      name: "月間削減時間(時間)",
      labels: ["リスト作成", "アプローチ", "商談まわり", "提案・見積", "報告・集計"],
      values: [18, 22, 15, 40, 25],
    }],
    axisNote: "営業5名合計での月間削減時間(時間)",
    opts: { barDir: "col", showValue: true, dataLabelPosition: "outEnd", barGapWidthPct: 60, chartColors: [C.navy], valAxisMaxVal: 55, valAxisMinVal: 0 },
    take: {
      icon: "FiBarChart2", title: "月120時間\n= 訪問80件分",
      body: "年間1,440時間。営業1人を採用せずに、訪問と提案に使える時間が生まれる計算です。この時間をどこに向けるかが経営判断です。",
    },
  });
  L.note(s, "※営業5名規模での試算例です。実際の効果は業務量と定着度により変動します。");
  s.addNotes("「訪問◯件分」に換算すると、経営者にはいちばん伝わる。");

  // 43 — I statement
  s = L.statement(pres, {
    kicker: "空いた時間の使い道",
    main: "戻ってくるのは、\n訪問と提案の時間。",
    sub: "人を減らす話ではありません。売る以外に使われていた時間を、売る時間に戻すだけです。\n訪問件数が増え、提案件数が増えれば、受注は後からついてきます。",
    iconName: "FiTarget",
  });
  s.addShape("roundRect", { x: M + 0.3, y: 6.0, w: 3.3, h: 0.6, rectRadius: 0.3, fill: { color: "2A3470" }, line: { color: "3A4585", width: 1 } });
  s.addText("訪問件数を増やす", { x: M + 0.3, y: 6.0, w: 3.3, h: 0.6, align: "center", valign: "middle", fontSize: 13, bold: true, color: C.ice, fontFace: F, margin: 0 });
  s.addShape("roundRect", { x: M + 3.85, y: 6.0, w: 3.3, h: 0.6, rectRadius: 0.3, fill: { color: "2A3470" }, line: { color: "3A4585", width: 1 } });
  s.addText("提案の質を上げる", { x: M + 3.85, y: 6.0, w: 3.3, h: 0.6, align: "center", valign: "middle", fontSize: 13, bold: true, color: C.ice, fontFace: F, margin: 0 });
  s.addShape("roundRect", { x: M + 7.4, y: 6.0, w: 3.3, h: 0.6, rectRadius: 0.3, fill: { color: "2A3470" }, line: { color: "3A4585", width: 1 } });
  s.addText("既存顧客を掘り起こす", { x: M + 7.4, y: 6.0, w: 3.3, h: 0.6, align: "center", valign: "middle", fontSize: 13, bold: true, color: C.ice, fontFace: F, margin: 0 });
  s.addNotes("営業の不安を煽らないため、社内展開時もこの説明から入るよう助言する。");

  // 44 — まとめ
  s = L.light(pres);
  L.head(s, "第3部のまとめ", null);
  L.rows(s, [
    { icon: "FiMap", title: "5工程すべてに、削減できる仕事がある", body: "リスト作成・アプローチ・商談・提案・報告。どこか1つだけの話ではありません。" },
    { icon: "FiFilter", title: "軽くなる仕事の型は「集める・書く・まとめる・つなぐ」", body: "この4語で自社の営業を棚卸しすれば、対象はすぐに見つかります。" },
    { icon: "FiBarChart2", title: "営業5名で月120時間、訪問80件分", body: "採用せずに生まれる時間です。しかも売る側の時間として戻ってきます。" },
    { icon: "FiTarget", title: "戻った時間の使い道は、経営が決める", body: "削減して終わりでは意味がありません。何に振り向けるかが成果を分けます。" },
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
  L.head(s, "1年で、訪問件数にここまで差がつく", "同業・同規模2社の、営業1人あたり月間訪問件数の推移");
  chartTakeaway(s, {
    type: "line",
    data: [
      { name: "取り組んだA社", labels: ["開始時", "3か月", "6か月", "9か月", "12か月"], values: [100, 106, 118, 129, 138] },
      { name: "様子を見たB社", labels: ["開始時", "3か月", "6か月", "9か月", "12か月"], values: [100, 99, 97, 95, 93] },
    ],
    axisNote: "開始時 = 100 とした営業1人あたり訪問件数の指数",
    opts: {
      lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 8,
      chartColors: [C.navy, C.accent], showLegend: true, legendPos: "b",
      legendColor: C.muted, legendFontFace: F, legendFontSize: 11,
      valAxisMaxVal: 150, valAxisMinVal: 80,
    },
    take: {
      icon: "FiActivity", title: "差は45ポイント\n1年で開いた",
      body: "B社が下がっているのは、何もしなかったからではありません。事務作業が自然に増えたためです。現状維持は、後退を意味します。",
    },
  });
  L.note(s, "※開始時を100とした指数のイメージ図です。");
  s.addNotes("「何もしない」の中身が現状維持ではなく後退である、という点が最大のメッセージ。");

  // 47 — 事例A
  s = caseSlide(pres, {
    title: "事例A:建設資材卸・営業5名", lead: "提案書作成の時間を、そのまま訪問件数に振り替えた",
    company: "建設資材卸\n従業員22名", icon: "FiShoppingCart",
    profile: "営業5名。提案・見積作成が1件2時間かかり、月末は誰も外に出られない状態だった。",
    items: [
      { icon: "FiAlertTriangle", title: "課題", body: "提案・見積作成で営業5名合計 月80時間。提案が出せず失注する案件が続いていた。" },
      { icon: "FiPlay", title: "やったこと", body: "過去3年分の提案書を読み込ませ、たたき台を自動作成。1名・1業務・3か月で検証した。" },
      { icon: "FiTrendingUp", title: "結果", body: "月40時間を削減。訪問件数が2割増え、提案件数の増加が受注2件につながった。" },
    ],
  });
  s.addNotes("「削減」より「訪問件数と受注が増えた」という成果を強調する。");

  // 48 — 事例B
  s = caseSlide(pres, {
    title: "事例B:機械商社・営業12名", lead: "日報を読む側を自動化し、失注する前に手を打てるようにした",
    company: "産業機械商社\n従業員45名", icon: "FiBarChart2",
    profile: "営業12名の日報が毎日届くが、社長も部長も全件は読めず、報告を待つ形になっていた。",
    items: [
      { icon: "FiAlertTriangle", title: "課題", body: "週60件の日報が読まれないまま蓄積。停滞案件に気づくのが、いつも失注の後だった。" },
      { icon: "FiPlay", title: "やったこと", body: "毎週月曜朝に日報を自動で読ませ、停滞案件と競合接触を一覧で出力する形にした。" },
      { icon: "FiTrendingUp", title: "結果", body: "週の途中でフォロー指示が出せるようになり、見積提出後の失注率が改善した。" },
    ],
  });
  s.addNotes("経営者自身が受け取る形になるのがポイント。社長が使う数少ない例。");

  // 49 — 事例C
  s = caseSlide(pres, {
    title: "事例C:導入したのに、使われなかった会社", lead: "同じツールを入れても、結果が正反対になることがあります",
    company: "住宅設備販売\n従業員30名", icon: "FiSlash", fail: true,
    profile: "「営業のDXを進める」と全営業にアカウントを配布。3か月後、使っていたのは1名だけだった。",
    items: [
      { icon: "FiUserX", title: "何が起きたか", body: "「各自で使ってみてほしい」と全員に配布。誰も自分の営業活動との結びつけ方が分からなかった。" },
      { icon: "FiClock", title: "現場の声", body: "「数字を追うので手一杯。新しいことを覚える時間がない」── 訪問目標は一切変わっていなかった。" },
      { icon: "FiAlertTriangle", title: "結末", body: "半年で契約を解除。社内には「うちの営業には合わなかった」という記憶だけが残った。" },
    ],
    note: "※想定事例。実際にもっとも多い失敗のパターンです。",
  });
  s.addNotes("失敗事例を必ず入れる。成功事例だけだと「うまくいった会社の話」で終わる。");

  // 50 — 失敗の共通点
  s = L.light(pres);
  L.head(s, "失敗する会社の、3つの共通点", "いずれもツールの問題ではなく、進め方の問題です");
  L.cards(s, [
    { icon: "FiUserX", title: "全員に配って、あとは任せた", body: "「各自で使って」で終わらせた。営業目標が変わらないまま、新しい作業だけが増える形になっていた。" },
    { icon: "FiHelpCircle", title: "目的が曖昧だった", body: "「営業をDXする」が目的化した。どの工程を、どれだけ軽くするのかが決まっていなかった。" },
    { icon: "FiBarChart2", title: "効果を測らなかった", body: "訪問件数も作成時間も記録していないため、成果が見えず、続ける理由も説明できなかった。" },
  ], { y: 2.05, h: 3.4 });
  banner(s, "3つとも、営業現場では決められないことです。つまり、失敗の原因は経営側にあります。", { y: 5.7, h: 0.8 });
  s.addNotes("ここで経営者を当事者にする。次のQで自問させる。");

  // 51 — Q
  s = L.question(pres, "導入したのに使われない。\nその原因は、現場ですか。経営ですか。");
  s.addNotes("答えは前のスライドで既に出ている。沈黙の時間を長めに取る。");

  // 52 — I statement
  s = L.statement(pres, {
    kicker: "差を生んでいるもの",
    main: "違いは、\n誰か1人に時間を渡したか。",
    sub: "A社もB社も、使ったものは同じでした。分かれ目は、どの工程を、誰が、いつまでに試すのか──\nその範囲を経営が決め、その人の通常業務を減らしたかどうか。それだけです。",
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
    { title: "時間を測る", body: "営業が何に時間を使っているかを記録する。\n\n所要:1週間\n経営の関与:指示を出す" },
    { title: "1工程を選ぶ", body: "効果と難易度で、最初の1つを決める。\n\n所要:数日\n経営の関与:自ら決める" },
    { title: "1人で試す", body: "1工程・1か月・1人で検証する。\n\n所要:1か月\n経営の関与:報告を受ける" },
    { title: "営業全員へ", body: "手順書にして、他の営業へ広げる。\n\n所要:3か月〜\n経営の関与:評価に組み込む" },
  ], { y: 2.05, h: 3.4 });
  banner(s, "失敗する会社は、この1と2を飛ばして、いきなり3から始めています。", { y: 5.7, h: 0.8, dark: false });
  s.addNotes("所要期間を示すことで「半年後には形になる」という見通しを持たせる。");

  // 55 — STEP1
  s = L.light(pres);
  L.head(s, "STEP 1:営業の時間を、1週間だけ測る", "「忙しい」ではなく、何に何分かかっているかを数字にします");
  L.rows(s, [
    { icon: "FiClipboard", title: "営業全員に、1週間だけ作業時間を記録してもらう", body: "細かい分類は不要。「何に、何分」だけ。日報の末尾に1行足すだけでも構いません。" },
    { icon: "FiFilter", title: "「集める・書く・まとめる・つなぐ」の4分類で仕分ける", body: "第3部の4つの型を使います。この4つに当てはまる作業が、そのまま候補になります。" },
    { icon: "FiBarChart2", title: "時間の多い順に並べ、上位3つを書き出す", body: "この3つが、次のステップで判断する対象です。全部やろうとしないこと。" },
  ], { y: 2.0, rowH: 1.3, gap: 0.24 });
  banner(s, "経営者がやることは1つ ── 「1週間、時間を記録してほしい」と指示を出すことです。", { y: 6.0, h: 0.72 });
  s.addNotes("配布資料に記録シートを付けておくと、実行率が大きく変わる。");

  // 56 — STEP2 マトリクス
  s = L.light(pres);
  L.head(s, "STEP 2:最初の1工程を決める", "効果の大きさと、着手のしやすさで並べます");
  const ox = 1.5, oy = 2.15, cw2 = 2.55, ch2 = 1.95;
  [
    { i: 0, j: 0, label: "最初に取る", body: "効果 大 × 着手 易", hot: true },
    { i: 1, j: 0, label: "次に取る", body: "効果 大 × 着手 難", hot: false },
    { i: 0, j: 1, label: "余力で取る", body: "効果 小 × 着手 易", hot: false },
    { i: 1, j: 1, label: "やらない", body: "効果 小 × 着手 難", hot: false },
  ].forEach((q) => {
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
  s.addShape("roundRect", { x: 7.15, y: 2.15, w: 5.43, h: 4.25, rectRadius: 0.1, fill: { color: C.tint }, line: { color: C.line, width: 1 }, shadow: L.shadow() });
  L.iconBadge(s, { x: 7.5, y: 2.5, d: 0.6, name: "FiTarget", tone: "light" });
  s.addText("最初の1つは、必ず社長が決める", {
    x: 7.5, y: 3.25, w: 4.75, h: 0.55, fontSize: 19, bold: true, color: C.navy, fontFace: F, margin: 0, valign: "middle",
  });
  const s56body = "営業に選ばせると、いちばん無難な作業が選ばれます。効果が小さいため成果が見えず、そのまま立ち消えになります。\n\n「提案書のたたき台から始める」「日報の集計から始める」── この一言を経営者が言えるかどうかで、結果が変わります。";
  s.addText(s56body, {
    x: 7.5, y: 3.9, w: 4.75, h: 2.25, fontSize: L.fit(s56body, 4.75, 2.25, 13.5, 1.35, 11),
    color: C.steel, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
  });
  s.addNotes("マトリクスの左上だけを指して「ここから」と言い切る。");

  // 57 — STEP3
  s = L.light(pres);
  L.head(s, "STEP 3:1工程・1か月・1人で試す", "全営業に配って始めた会社は、ほぼ例外なく失敗しています");
  L.stats(s, [
    { label: "対象工程", value: "1", unit: "つだけ", sub: "STEP2で選んだ工程に限定。増やしたくなっても、この期間は広げない" },
    { label: "検証期間", value: "1", unit: "か月", sub: "短すぎると慣れず、長すぎると熱が冷める。1か月がちょうどよい長さ" },
    { label: "担当者", value: "1", unit: "名", sub: "いちばん困っている営業を選ぶ。前向きな人ではなく、負荷が重い人" },
  ], { y: 2.15, h: 2.6 });
  banner(s, "その1名の訪問目標を、検証期間だけ下げてください。時間を渡さない限り、誰も新しいことは始められません。", { y: 5.05, h: 1.05, dark: false, icon: "FiClock" });
  L.note(s, "※開始前と終了後の作業時間・訪問件数を必ず記録してください。数字がないと次の判断ができません。");
  s.addNotes("「訪問目標を下げる」がこのスライドの核心。ここを飛ばすと必ず失敗する。");

  // 58 — STEP4
  s = L.light(pres);
  L.head(s, "STEP 4:型にして、営業全員へ広げる", "1人の工夫で終わらせず、会社の仕組みに変えます");
  L.rows(s, [
    { icon: "FiClipboard", title: "指示の出し方を1枚の紙にまとめる", body: "「何をどう頼むと、何が出てくるか」だけで十分。担当者本人に書いてもらいます。" },
    { icon: "FiUsers", title: "同じ工程を持つ営業に渡す", body: "説明会は不要。手順書を渡して、隣で1回一緒にやれば伝わります。" },
    { icon: "FiSettings", title: "次の工程に着手する", body: "STEP2で選んだ2番目の工程へ。1つずつ、確実に潰していきます。" },
    { icon: "FiAward", title: "成果を評価に反映する", body: "時間を生んだ人が報われる形にしないと、2周目は起きません。ここは経営の仕事です。" },
  ], { y: 1.85, rowH: 1.12, gap: 0.18 });
  s.addNotes("最後の1行が、継続するかどうかの分岐点であることを強調する。");

  // 59 — 経営者の役割
  s = L.light(pres);
  L.head(s, "経営者の役割は「やっていい」と決めること", "営業は、数字を追いながら新しいことを始められません");
  s.addShape("roundRect", { x: M, y: 2.05, w: 5.4, h: 4.1, rectRadius: 0.1, fill: { color: C.tint }, line: { color: C.line, width: 1 } });
  s.addText("営業の頭の中", { x: M + 0.35, y: 2.3, w: 4.7, h: 0.45, fontSize: 16, bold: true, color: C.navy, fontFace: F, margin: 0, valign: "middle" });
  s.addText([
    { text: "「今月の数字を追うので手一杯だ」", options: { bullet: true, breakLine: true, paraSpaceAfter: 12 } },
    { text: "「顧客情報を入れて、問題にならないか」", options: { bullet: true, breakLine: true, paraSpaceAfter: 12 } },
    { text: "「試している間に、訪問件数が落ちる」", options: { bullet: true, breakLine: true, paraSpaceAfter: 12 } },
    { text: "「そもそも、会社として認めているのか」", options: { bullet: true } },
  ], { x: M + 0.35, y: 2.9, w: 4.7, h: 3.0, fontSize: 14, color: C.steel, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.3 });
  s.addImage({ data: L.icon("FiArrowRight", "navy"), x: 6.42, y: 3.95, w: 0.35, h: 0.35 });
  s.addShape("roundRect", { x: 7.18, y: 2.05, w: 5.4, h: 4.1, rectRadius: 0.1, fill: { color: C.navy }, line: { color: C.navy }, shadow: L.shadow({ opacity: 0.18 }) });
  s.addText("社長の一言で外れる", { x: 7.53, y: 2.3, w: 4.7, h: 0.45, fontSize: 16, bold: true, color: C.ice, fontFace: F, margin: 0, valign: "middle" });
  s.addText("「提案書の作成で、\n1か月試してほしい。\nその分、今月の訪問目標は\n下げていい。」", {
    x: 7.53, y: 2.95, w: 4.7, h: 2.0, fontSize: 20, bold: true, color: C.white, fontFace: F, margin: 0, valign: "middle", lineSpacingMultiple: 1.28,
  });
  s.addText("投資判断でも、ツール選定でもありません。決めるのは、やっていい範囲と、渡す時間だけです。", {
    x: 7.53, y: 5.1, w: 4.7, h: 0.9, fontSize: 13, color: C.ice, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
  });
  s.addNotes("このスライドが第5部の核。ここは時間をかけてよい。");

  // 60 — 動かない理由
  s = L.light(pres);
  L.head(s, "営業が動かない3つの理由と、その外し方", "精神論では動きません。構造を外してください");
  L.cards(s, [
    { icon: "FiClock", title: "数字を追うので手一杯", body: "【外し方】\n検証期間だけ訪問目標を下げる。「その分の数字は問わない」まで言い切る。" },
    { icon: "FiAlertTriangle", title: "問題になったら困る", body: "【外し方】\n入力してよい情報の線引きを先に示す。「このルールの範囲なら責任は問わない」と宣言する。" },
    { icon: "FiAward", title: "やっても評価されない", body: "【外し方】\n生まれた時間と訪問件数の増加を成果として評価する。朝礼で名前を出すだけでも効きます。" },
  ], { y: 2.05, h: 3.4 });
  banner(s, "3つとも、社長にしか外せません。だから「現場に丸投げ」は必ず失敗するのです。", { y: 5.7, h: 0.8 });
  s.addNotes("第4部の失敗事例と対応させながら話すと、腹落ちが早い。");

  // 61 — リスク
  s = L.light(pres);
  L.head(s, "リスク管理:営業ならではの3つの落とし穴", "顧客情報を扱う以上、ここは避けて通れません");
  L.cards(s, [
    { icon: "FiLock", title: "顧客情報の取り扱い", body: "顧客名簿・取引条件・見積単価をそのまま入力してしまう。何を入れてよいかの線引きが必要です。" },
    { icon: "FiAlertTriangle", title: "事実と異なる内容", body: "納期や仕様をもっともらしく書いてしまうことがあります。そのまま顧客に送ると信用問題になります。" },
    { icon: "FiSend", title: "確認せずに社外へ送る", body: "作られた文面をそのままメール送信してしまう事故。社外に出る前の確認者を決める必要があります。" },
  ], { y: 2.05, h: 3.4, accent: true });
  banner(s, "いずれも、ルールを1枚決めておけば大半は防げます。難しい規程は要りません。", { y: 5.7, h: 0.8 });
  s.addNotes("怖がらせて終わらない。次のスライドで対処法を必ずセットで見せる。");

  // 62 — 社内ルール
  s = L.light(pres);
  L.head(s, "最低限、決めておく社内ルール3項目", "A4で1枚。これ以上は要りません");
  L.checklist(s, [
    { title: "入力してよい情報の線引き", body: "顧客名・取引単価・未公開の商談情報は入力しない。判断に迷うものは上長に確認する。" },
    { title: "使ってよいサービスの限定", body: "会社が認めたものだけを使用する。無料版・個人アカウントでの業務利用は禁止する。" },
    { title: "社外に出す前の確認者を決める", body: "AIが作った文面を顧客に送る場合は、必ず担当者以外が1回確認する。" },
  ], { y: 2.1, itemH: 1.0, gap: 0.24 });
  banner(s, "この3行を朝礼で共有するだけでも、無法状態からは抜け出せます。", { y: 5.85, h: 0.85, dark: false, icon: "FiShield" });
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
    { icon: "FiUserCheck", title: "専門家派遣制度", body: "専門家を一定回数まで低額または無料で招けます。時間の棚卸しの段階から入ってもらうのが有効です。" },
  ], { y: 2.05, h: 3.4 });
  banner(s, "補助金ありきで進めないこと。まず1工程で試し、効果が見えてから制度を使うのが順番です。", { y: 5.7, h: 0.8, dark: false });
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
    { num: "1", title: "営業が売っている\n時間は4割以下", body: "残りは資料作成・報告・情報探しです。営業を増やせない以上、この時間を取り戻すしかありません。" },
    { num: "2", title: "作業をやり切るAIが\n出てきた", body: "答えを返すだけでなく、資料を読み、作り、保存するところまで進みます。黒い画面は不要、指示は日本語です。" },
    { num: "3", title: "決めるのは、\n経営者です", body: "どの工程を、誰が、いつまでに試すのか。そして、その人の訪問目標を下げること。ここが結果を分けます。" },
  ], { y: 2.05, h: 3.9 });
  s.addNotes("3つを言い切って、次の問いに入る。補足説明は加えない。");

  // 67 — Q
  s = L.question(pres, "1年後も、御社の営業は\n同じ時間を事務に使っていていいですか?",
    "── 冒頭の問いを、もう一度");
  s.addNotes("冒頭の問いと対になる。ここが感情のピーク。間を長く取る。");

  // 68 — 3アクション
  s = L.light(pres);
  L.head(s, "明日からの3アクション", "この順番で、そのまま実行できます");
  L.checklist(s, [
    { title: "今週:営業に「1週間、作業時間を記録してほしい」と伝える", body: "細かい分類は不要。何に何分かかったかだけ。" },
    { title: "今月:記録を見て、最初の1工程を社長が決める", body: "効果が大きく、着手しやすいものから。営業に選ばせない。" },
    { title: "来月:1工程・1か月・1人で試す。その人の訪問目標を下げる", body: "開始前と終了後の時間・訪問件数を必ず記録する。" },
  ], { y: 2.1, itemH: 1.05, gap: 0.26 });
  banner(s, "3か月後には、自社に合うかどうかの答えが、数字で出ています。", { y: 5.95, h: 0.85 });
  s.addNotes("配布資料の記録シートと連動させる。ここで具体的な次アクションを渡し切る。");

  // 69 — I statement
  s = L.statement(pres, {
    kicker: "最後にひとつだけ",
    main: "最初の1工程を、\n今ここで決めてください。",
    sub: "会場を出てからでは、日常業務に戻って忘れてしまいます。\nリスト作成でも、提案書でも、日報の集計でも構いません。ひとつだけ選んでください。",
    iconName: "FiPlay",
  });
  s.addNotes("実際に10秒黙る。この沈黙が、行動につながるかどうかを分ける。");

  // 70 — 窓口
  s = L.light(pres);
  L.head(s, "ご相談窓口", "自社だけで抱え込まず、まずは相談から始めてください");
  L.cards(s, [
    { icon: "FiPhone", title: "主催者の経営相談窓口", body: "〇〇商工会議所 経営支援課\nTEL:000-0000-0000\n受付:平日 9:00〜17:00" },
    { icon: "FiUserCheck", title: "本日の講師への相談", body: "所属・部署名\nメール:info@example.jp\n※本日の資料もこちらからお送りします" },
    { icon: "FiClipboard", title: "配布資料について", body: "営業時間の記録シート\n社内ルールのひな形(A4・1枚)\n本日のスライド(抜粋版)" },
  ], { y: 2.0, h: 3.55 });
  banner(s, "本日はありがとうございました。まずは「1週間の時間記録」から始めてみてください。", { y: 5.75, h: 0.85 });
  L.note(s, "※窓口・連絡先は、開催前に主催者および登壇者の情報へ差し替えてください。");
  s.addNotes("質疑応答へ。回答は短く、個別具体は窓口へ誘導する。");

  const out = process.argv[2] || "seminar2.pptx";
  await pres.writeFile({ fileName: out });
  console.log("slides:", L.state.page, "->", out);
}

main().catch((e) => { console.error(e); process.exit(1); });
