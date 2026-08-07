// Shared design system for the seminar deck — "Midnight Executive"
const W = 13.333, H = 7.5;
const M = 0.75;            // side margin
const CW = W - 2 * M;      // content width

const C = {
  navy: "1E2761",
  deep: "141C4A",
  steel: "1A3A5C",
  ice: "CADCFC",
  white: "FFFFFF",
  tint: "EFF3FC",
  tintDeep: "DEE7F8",
  line: "D6E0F3",
  muted: "667090",
  accent: "D98A34",
  accentSoft: "F7ECDD",
  darkCard: "2A3470",
};

const F = "Noto Sans JP";

// fresh option objects every call — pptxgenjs mutates them in place
const shadow = (o = {}) => ({
  type: "outer", angle: 90, blur: 12, offset: 2,
  color: "1E2761", opacity: 0.10, ...o,
});

const state = { icons: {}, page: 0 };

// ---------- CJK-aware text fitting ----------

function charW(ch) {
  const o = ch.codePointAt(0);
  if ((o >= 0x3000 && o <= 0x9FFF) || (o >= 0xFF00 && o <= 0xFF60)) return 1.0;
  if (ch >= "0" && ch <= "9") return 0.56;
  if (ch >= "A" && ch <= "Z") return 0.66;
  if (ch === " ") return 0.28;
  return 0.52;
}

function estLines(text, widthIn, sizePt, indent = 0) {
  const em = sizePt / 72;
  const cap = Math.max(1, (widthIn - indent) / em);
  let lines = 0;
  for (const seg of String(text).split("\n")) {
    let cur = 0, n = 1;
    for (const ch of seg) {
      const cw = charW(ch);
      if (cur + cw > cap) { n++; cur = cw; } else { cur += cw; }
    }
    lines += n;
  }
  return lines;
}

// largest font size (<= base) whose wrapped text fits the box
function fit(text, w, h, base, mult = 1.3, min = 10.5, indent = 0) {
  for (let sz = base; sz >= min; sz -= 0.5) {
    if (estLines(text, w, sz, indent) * (sz / 72) * mult * 1.2 <= h) return sz;
  }
  return min;
}

function icon(name, tone = "navy") {
  const d = state.icons[name + ":" + tone];
  if (!d) throw new Error("icon not rendered: " + name + ":" + tone);
  return d;
}

// ---------- slide scaffolds ----------

function light(pres, { page = true } = {}) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  state.page++;
  if (page) pageNum(s, C.muted);
  return s;
}

function dark(pres, { page = true, bg = C.navy } = {}) {
  const s = pres.addSlide();
  s.background = { color: bg };
  state.page++;
  if (page) pageNum(s, "8B96BC");
  return s;
}

function pageNum(s, color) {
  s.addText(String(state.page), {
    x: W - 1.05, y: H - 0.55, w: 0.55, h: 0.3, align: "right",
    fontSize: 10, color, fontFace: F, margin: 0,
  });
}

// unified title block for every content slide
function head(s, title, lead) {
  s.addText(title, {
    x: M, y: 0.46, w: CW, h: 0.72, fontSize: 30, bold: true,
    color: C.navy, fontFace: F, margin: 0, valign: "middle",
  });
  if (lead) {
    s.addText(lead, {
      x: M, y: 1.22, w: CW, h: 0.42, fontSize: 14, color: C.muted,
      fontFace: F, margin: 0, valign: "middle",
    });
  }
  return lead ? 1.9 : 1.62;
}

function note(s, text, y = H - 0.62) {
  s.addText(text, {
    x: M, y, w: CW - 1.2, h: 0.3, fontSize: 10, color: C.muted,
    fontFace: F, margin: 0, italic: true,
  });
}

// ---------- motif: icon in a circle ----------

function iconBadge(s, { x, y, d = 0.66, name, tone = "light" }) {
  const fill = tone === "light" ? C.ice : C.white;
  const ic = tone === "light" ? "navy" : "navy";
  s.addShape("ellipse", { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill } });
  const p = d * 0.28;
  s.addImage({ data: icon(name, ic), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}

// ---------- components ----------

// cards in a row: items = [{icon, title, body}]
function cards(s, items, opts = {}) {
  const { y = 2.0, h = 2.6, gap = 0.34, dark: isDark = false, accent = false } = opts;
  const n = items.length;
  const w = (CW - gap * (n - 1)) / n;
  items.forEach((it, i) => {
    const x = M + i * (w + gap);
    s.addShape("roundRect", {
      x, y, w, h, rectRadius: 0.1,
      fill: { color: isDark ? C.darkCard : (accent ? C.accentSoft : C.tint) },
      line: { color: isDark ? C.darkCard : (accent ? "EBD9C2" : C.line), width: 1 },
      shadow: isDark ? undefined : shadow(),
    });
    let cy = y + 0.34;
    if (it.icon) {
      iconBadge(s, { x: x + 0.34, y: cy, d: 0.62, name: it.icon, tone: isDark ? "dark" : "light" });
      cy += 0.86;
    }
    if (it.num) {
      s.addShape("ellipse", { x: x + 0.34, y: cy, w: 0.62, h: 0.62, fill: { color: isDark ? C.ice : C.navy }, line: { color: isDark ? C.ice : C.navy } });
      s.addText(it.num, { x: x + 0.34, y: cy, w: 0.62, h: 0.62, align: "center", valign: "middle", fontSize: 20, bold: true, color: isDark ? C.navy : C.white, fontFace: F, margin: 0 });
      cy += 0.86;
    }
    const tw = w - 0.68;
    s.addText(it.title, {
      x: x + 0.34, y: cy, w: tw, h: 0.74, fontSize: fit(it.title, tw, 0.74, 17, 1.1, 13.5),
      bold: true, color: isDark ? C.white : C.navy, fontFace: F, margin: 0,
      valign: "top", lineSpacingMultiple: 1.1,
    });
    if (it.body) {
      const bh = y + h - (cy + 0.8) - 0.28;
      s.addText(it.body, {
        x: x + 0.34, y: cy + 0.8, w: tw, h: bh,
        fontSize: opts.bodySize || fit(it.body, tw, bh, 13, 1.32, 10.5),
        color: isDark ? C.ice : C.steel, fontFace: F, margin: 0,
        valign: "top", lineSpacingMultiple: 1.32,
      });
    }
  });
}

// stacked rows: items = [{icon, title, body}]
function rows(s, items, opts = {}) {
  const { y = 2.0, rowH = 1.15, gap = 0.22, x = M, w = CW, dark: isDark = false } = opts;
  items.forEach((it, i) => {
    const ry = y + i * (rowH + gap);
    s.addShape("roundRect", {
      x, y: ry, w, h: rowH, rectRadius: 0.08,
      fill: { color: isDark ? C.darkCard : C.tint },
      line: { color: isDark ? C.darkCard : C.line, width: 1 },
    });
    const bd = Math.min(0.62, rowH - 0.4);
    iconBadge(s, { x: x + 0.3, y: ry + (rowH - bd) / 2, d: bd, name: it.icon, tone: isDark ? "dark" : "light" });
    const tx = x + 0.3 + bd + 0.28;
    if (it.body) {
      s.addText(it.title, {
        x: tx, y: ry + 0.16, w: w - (tx - x) - 0.3, h: 0.36, fontSize: 16, bold: true,
        color: isDark ? C.white : C.navy, fontFace: F, margin: 0, valign: "middle",
      });
      const bw = w - (tx - x) - 0.3, bh = rowH - 0.66;
      s.addText(it.body, {
        x: tx, y: ry + 0.52, w: bw, h: bh, fontSize: fit(it.body, bw, bh, 13, 1.25, 10.5),
        color: isDark ? C.ice : C.steel, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
      });
    } else {
      s.addText(it.title, {
        x: tx, y: ry, w: w - (tx - x) - 0.3, h: rowH, fontSize: 16, bold: true,
        color: isDark ? C.white : C.navy, fontFace: F, margin: 0, valign: "middle",
      });
    }
  });
}

// big numbers: items = [{value, unit, label, sub}]
function stats(s, items, opts = {}) {
  const { y = 2.2, h = 2.1, gap = 0.34 } = opts;
  const n = items.length;
  const w = (CW - gap * (n - 1)) / n;
  items.forEach((it, i) => {
    const x = M + i * (w + gap);
    s.addShape("roundRect", {
      x, y, w, h, rectRadius: 0.1, fill: { color: C.navy }, line: { color: C.navy },
      shadow: shadow({ opacity: 0.18 }),
    });
    s.addText(it.label, {
      x: x + 0.3, y: y + 0.26, w: w - 0.6, h: 0.34, fontSize: 13, bold: true,
      color: C.ice, fontFace: F, margin: 0, valign: "middle",
    });
    s.addText(
      [
        { text: it.value, options: { fontSize: 44, bold: true, color: C.white } },
        { text: it.unit ? " " + it.unit : "", options: { fontSize: 16, bold: true, color: C.ice } },
      ],
      { x: x + 0.3, y: y + 0.66, w: w - 0.6, h: 0.86, fontFace: F, margin: 0, valign: "middle" }
    );
    if (it.sub) {
      s.addText(it.sub, {
        x: x + 0.3, y: y + 1.5, w: w - 0.6, h: h - 1.74, fontSize: 12,
        color: C.ice, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.25,
      });
    }
  });
}

// two-column compare: {label, items:[], tone:'before'|'after'}
function compare(s, left, right, opts = {}) {
  const { y = 2.0, gap = 0.7 } = opts;
  const w = (CW - gap) / 2;
  let h = opts.h;
  if (h === undefined) {
    const need = (col) => col.items.reduce(
      (a, t) => a + estLines(t, w - 0.72 - 0.22, 14) * (14 / 72) * 1.3 * 1.2 + 0.125, 0);
    h = Math.min(4.35, Math.max(3.0, 0.92 + Math.max(need(left), need(right)) + 0.3));
  }
  [[left, M], [right, M + w + gap]].forEach(([col, x]) => {
    const isAfter = col.tone === "after";
    s.addShape("roundRect", {
      x, y, w, h, rectRadius: 0.1,
      fill: { color: isAfter ? C.navy : C.tint },
      line: { color: isAfter ? C.navy : C.line, width: 1 },
      shadow: shadow({ opacity: isAfter ? 0.18 : 0.1 }),
    });
    s.addShape("roundRect", {
      x: x + 0.32, y: y + 0.3, w: 1.62, h: 0.4, rectRadius: 0.2,
      fill: { color: isAfter ? C.ice : C.navy }, line: { color: isAfter ? C.ice : C.navy },
    });
    s.addText(col.label, {
      x: x + 0.32, y: y + 0.3, w: 1.62, h: 0.4, align: "center", valign: "middle",
      fontSize: 13, bold: true, color: isAfter ? C.navy : C.white, fontFace: F, margin: 0,
    });
    if (col.head) {
      s.addText(col.head, {
        x: x + 2.08, y: y + 0.3, w: w - 2.4, h: 0.4, fontSize: 14, bold: true,
        color: isAfter ? C.white : C.navy, fontFace: F, margin: 0, valign: "middle",
      });
    }
    s.addText(
      col.items.map((t, i) => ({
        text: t,
        options: { bullet: true, breakLine: i !== col.items.length - 1, paraSpaceAfter: 9 },
      })),
      {
        x: x + 0.36, y: y + 0.92, w: w - 0.72, h: h - 1.2,
        fontSize: fit(col.items.join("\n"), w - 0.72, h - 1.2 - 0.09 * col.items.length, 14, 1.3, 11),
        color: isAfter ? C.white : C.steel, fontFace: F, margin: 0, valign: "top",
        lineSpacingMultiple: 1.3,
      }
    );
  });
}

// horizontal numbered flow: items = [{num, title, body}]
function steps(s, items, opts = {}) {
  const { y = 2.2, h = 2.5, gap = 0.3 } = opts;
  const n = items.length;
  const w = (CW - gap * (n - 1)) / n;
  items.forEach((it, i) => {
    const x = M + i * (w + gap);
    s.addShape("roundRect", {
      x, y, w, h, rectRadius: 0.1, fill: { color: C.tint }, line: { color: C.line, width: 1 },
      shadow: shadow(),
    });
    s.addShape("ellipse", { x: x + 0.3, y: y + 0.3, w: 0.6, h: 0.6, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText(String(i + 1), {
      x: x + 0.3, y: y + 0.3, w: 0.6, h: 0.6, align: "center", valign: "middle",
      fontSize: 20, bold: true, color: C.white, fontFace: F, margin: 0,
    });
    s.addText(it.title, {
      x: x + 1.02, y: y + 0.3, w: w - 1.32, h: 0.6, fontSize: 16, bold: true,
      color: C.navy, fontFace: F, margin: 0, valign: "middle",
    });
    const sbw = w - 0.6, sbh = h - 1.34;
    s.addText(it.body, {
      x: x + 0.3, y: y + 1.06, w: sbw, h: sbh, fontSize: fit(it.body, sbw, sbh, 13, 1.3, 10.5),
      color: C.steel, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.3,
    });
    if (i < n - 1) {
      s.addImage({ data: icon("FiArrowRight", "navy"), x: x + w + (gap - 0.22) / 2, y: y + h / 2 - 0.11, w: 0.22, h: 0.22 });
    }
  });
}

// checklist card
function checklist(s, items, opts = {}) {
  const { y = 2.0, x = M, w = CW, itemH = 0.82, gap = 0.18 } = opts;
  items.forEach((it, i) => {
    const ry = y + i * (itemH + gap);
    s.addShape("roundRect", { x, y: ry, w, h: itemH, rectRadius: 0.08, fill: { color: C.tint }, line: { color: C.line, width: 1 } });
    s.addShape("ellipse", { x: x + 0.28, y: ry + (itemH - 0.44) / 2, w: 0.44, h: 0.44, fill: { color: C.navy }, line: { color: C.navy } });
    s.addImage({ data: icon("FiCheckSquare", "white"), x: x + 0.28 + 0.11, y: ry + (itemH - 0.44) / 2 + 0.11, w: 0.22, h: 0.22 });
    s.addText(
      [
        { text: it.title, options: { bold: true, color: C.navy, fontSize: 15 } },
        { text: it.body ? "　" + it.body : "", options: { color: C.steel, fontSize: 13 } },
      ],
      { x: x + 0.92, y: ry, w: w - 1.2, h: itemH, fontFace: F, margin: 0, valign: "middle" }
    );
  });
}

// full-bleed statement slide (I pages)
function statement(pres, { kicker, main, sub, iconName }) {
  const s = dark(pres);
  if (iconName) {
    s.addImage({ data: icon(iconName, "ice"), x: W - 3.5, y: 2.35, w: 3.0, h: 3.0, transparency: 78 });
  }
  if (kicker) {
    s.addText(kicker, {
      x: M + 0.3, y: 2.15, w: CW - 0.6, h: 0.4, fontSize: 15, bold: true,
      color: C.ice, fontFace: F, margin: 0, charSpacing: 1,
    });
  }
  s.addText(main, {
    x: M + 0.3, y: 2.68, w: CW - 3.0, h: 1.9, fontSize: 40, bold: true,
    color: C.white, fontFace: F, margin: 0, valign: "middle", lineSpacingMultiple: 1.24,
  });
  if (sub) {
    s.addText(sub, {
      x: M + 0.3, y: 4.72, w: CW - 3.0, h: 1.12, fontSize: fit(sub, CW - 3.0, 1.12, 15, 1.35, 12),
      color: C.ice, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
    });
  }
  return s;
}

// question slide (Q pages)
function question(pres, text, sub) {
  const s = dark(pres, { bg: C.deep });
  s.addShape("ellipse", { x: W / 2 - 0.33, y: 1.95, w: 0.66, h: 0.66, fill: { color: C.ice }, line: { color: C.ice } });
  s.addImage({ data: icon("FiHelpCircle", "navy"), x: W / 2 - 0.19, y: 2.09, w: 0.38, h: 0.38 });
  s.addText(text, {
    x: 1.1, y: 2.85, w: W - 2.2, h: 2.3, align: "center", valign: "middle",
    fontSize: fit(text, W - 2.2, 2.3, 36, 1.28, 26), bold: true, color: C.white,
    fontFace: F, margin: 0, lineSpacingMultiple: 1.28,
  });
  if (sub) {
    s.addText(sub, {
      x: 1.1, y: 5.35, w: W - 2.2, h: 0.5, align: "center", valign: "middle",
      fontSize: 14, color: C.ice, fontFace: F, margin: 0,
    });
  }
  return s;
}

// chapter divider — unified type sizes across all chapters
function divider(pres, { part, title, sub, iconName }) {
  const s = dark(pres);
  if (iconName) {
    s.addImage({ data: icon(iconName, "ice"), x: W - 3.9, y: 1.9, w: 3.2, h: 3.2, transparency: 80 });
  }
  s.addText(part, {
    x: M + 0.3, y: 2.35, w: 5.0, h: 0.44, fontSize: 16, bold: true,
    color: C.ice, fontFace: F, margin: 0, charSpacing: 2,
  });
  s.addText(title, {
    x: M + 0.3, y: 2.78, w: CW - 3.6, h: 1.8, fontSize: 40, bold: true,
    color: C.white, fontFace: F, margin: 0, valign: "middle",
  });
  if (sub) {
    s.addText(sub, {
      x: M + 0.3, y: 4.68, w: CW - 3.6, h: 0.9, fontSize: 15,
      color: C.ice, fontFace: F, margin: 0, valign: "top", lineSpacingMultiple: 1.35,
    });
  }
  return s;
}

// chart frame + common styling
const chartBase = (extra = {}) => ({
  showLegend: false,
  chartColors: [C.navy, C.ice, C.accent],
  catAxisLabelColor: C.muted, catAxisLabelFontSize: 11, catAxisLabelFontFace: F,
  valAxisLabelColor: C.muted, valAxisLabelFontSize: 11, valAxisLabelFontFace: F,
  valGridLine: { color: C.line, size: 1 },
  catGridLine: { style: "none" },
  dataLabelFontFace: F, dataLabelFontSize: 11, dataLabelColor: C.navy,
  border: { pt: 0, color: "FFFFFF" },
  ...extra,
});

function chartPanel(s, { x = M, y = 2.0, w = CW, h = 4.0 } = {}) {
  s.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.08, fill: { color: C.white },
    line: { color: C.line, width: 1 }, shadow: shadow(),
  });
  return { x: x + 0.25, y: y + 0.25, w: w - 0.5, h: h - 0.5 };
}

module.exports = {
  W, H, M, CW, C, F, shadow, state, icon, fit, estLines,
  light, dark, head, note, iconBadge,
  cards, rows, stats, compare, steps, checklist,
  statement, question, divider, chartBase, chartPanel,
};
