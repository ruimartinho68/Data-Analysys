// HTML → DOCX converter for the Screens Down Family Up ebook.
// Maps each of the book's components (cover, chapter heroes, callouts,
// script boxes, stat cards/bars, schedules, checklists, templates, TOC)
// to native Word constructs via docx-js.
const fs = require('fs');
const { parse } = require('node-html-parser');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign,
  TabStopType, LeaderType, HeightRule, Bookmark, InternalHyperlink, PageReference,
} = require('docx');

// ---------- palette ----------
const NAVY = '101B33', BLUE = '2A4C8C', GOLD = 'F0B429', CREAM_TXT = 'FDFBF6';
const BODY = '2C3448', MUTED = '5D6E96', TAN = '8A8570', LINE = 'E4DCC8';
const FILL_LINE = 'C7BFA0';
const OSWALD = 'Oswald', LORA = 'Lora', SS3 = 'Source Sans 3';

const CW = 10080; // content width in DXA (7in)

const hex = (s) => s.replace('#', '').toUpperCase();
const blend = (fg, bg, a) => {
  const p = (h, i) => parseInt(h.substr(i, 2), 16);
  fg = hex(fg); bg = hex(bg);
  return [0, 2, 4].map((i) =>
    Math.round(p(fg, i) * a + p(bg, i) * (1 - a)).toString(16).padStart(2, '0')
  ).join('').toUpperCase();
};

const NONE = { style: BorderStyle.NONE, size: 0, color: 'auto' };
const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const allNone = { ...noBorders, insideHorizontal: NONE, insideVertical: NONE };
const bd = (color, size = 8) => ({ style: BorderStyle.SINGLE, size, color: hex(color) });
const shade = (fill) => ({ type: ShadingType.CLEAR, fill: hex(fill), color: 'auto' });

// ---------- inline content → TextRuns ----------
const decode = (s) => s
  .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<')
  .replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'");

// fmt: {font,size,color,bold,italics,caps,spacing,strongColor}
function runsOf(node, fmt, out = [], state = { br: false }) {
  for (const c of node.childNodes) {
    if (c.nodeType === 3) {
      const text = decode(c.rawText).replace(/\s+/g, ' ');
      if (!text) continue;
      out.push(new TextRun({
        text, font: fmt.font, size: fmt.size, color: fmt.color,
        bold: fmt.bold, italics: fmt.italics, allCaps: fmt.caps,
        characterSpacing: fmt.spacing, shading: fmt.shading,
        break: state.br ? 1 : undefined,
      }));
      state.br = false;
    } else if (c.nodeType === 1) {
      const tag = c.tagName;
      if (tag === 'BR') { state.br = true; continue; }
      const cls = c.getAttribute('class') || '';
      const style = c.getAttribute('style') || '';
      if (cls.includes('ci-dot')) continue; // rendered as ☐ by caller
      if (cls.includes('fill-line') || /border-bottom/.test(style)) {
        // blank line to write on → underlined no-break spaces
        const long = cls.includes('fill-line');
        out.push(new TextRun({
          text: '_'.repeat(long ? 30 : 6),
          font: fmt.font, size: fmt.size, color: FILL_LINE,
          break: state.br ? 1 : undefined,
        }));
        state.br = !long ? state.br : true; // block fill-line ends the line
        if (long) state.br = true;
        continue;
      }
      const sub = { ...fmt };
      if (tag === 'STRONG' || tag === 'B') { sub.bold = true; if (fmt.strongColor) sub.color = fmt.strongColor; if (fmt.strongShading) sub.shading = fmt.strongShading; }
      if (tag === 'EM' || tag === 'I') sub.italics = true;
      runsOf(c, sub, out, state);
    }
  }
  return out;
}

const els = (node) => node.childNodes.filter((c) => c.nodeType === 1);
const cls1 = (node) => (node.getAttribute('class') || '').split(/\s+/)[0] || '';
const styleOf = (node) => {
  const s = {}; (node.getAttribute('style') || '').split(';').forEach((d) => {
    const i = d.indexOf(':'); if (i > 0) s[d.slice(0, i).trim()] = d.slice(i + 1).trim();
  }); return s;
};

// one-cell colored panel
function panel(children, { fill, borders, margins }) {
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: { ...(borders || allNone) },
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        shading: fill ? shade(fill) : undefined,
        margins: margins || { top: 200, bottom: 200, left: 260, right: 260 },
        borders: borders ? undefined : noBorders,
        children,
      })],
    })],
  });
}

const spacer = (after = 120) => new Paragraph({ spacing: { after, before: 0, line: 40, lineRule: 'exact' }, children: [new TextRun({ text: '', size: 2 })] });
const pageBreak = () => new Paragraph({ pageBreakBefore: true, spacing: { after: 0, line: 40, lineRule: 'exact' }, children: [new TextRun({ text: '', size: 2 })] });

// ---------- block builders ----------
const out = [];

function cover(node) {
  const inner = node.querySelector('.cover-inner');
  const get = (sel) => inner.querySelector(sel);
  const kids = [];
  const C = AlignmentType.CENTER;
  kids.push(new Paragraph({
    alignment: C, spacing: { before: 600, after: 160 },
    children: [new TextRun({ text: decode(get('.cover-eyebrow').text).trim(), font: OSWALD, size: 18, bold: true, allCaps: true, characterSpacing: 60, color: 'C6D2EE' })],
  }));
  kids.push(new Paragraph({ // gold rule
    alignment: C, spacing: { after: 300 },
    children: [new TextRun({ text: '        ', size: 8, underline: { type: 'thick', color: GOLD } })],
  }));
  kids.push(new Paragraph({
    alignment: C, spacing: { after: 0 },
    children: [new TextRun({ text: 'Screens Down,', font: OSWALD, size: 60, bold: true, allCaps: true, color: 'C6D2EE' })],
  }));
  kids.push(new Paragraph({
    alignment: C, spacing: { after: 240 },
    children: [new TextRun({ text: 'Family Up', font: OSWALD, size: 105, bold: true, allCaps: true, color: GOLD })],
  }));
  kids.push(new Paragraph({
    alignment: C, spacing: { after: 260 },
    children: [new TextRun({ text: decode(get('.cover-subtitle').text).trim(), font: SS3, size: 24, bold: true, color: 'EEF1F8' })],
  }));
  for (const prop of inner.querySelectorAll('.cover-prop')) {
    const strong = prop.querySelector('strong');
    const rest = decode(prop.text).trim().replace(decode(strong.text).trim(), '').trim();
    kids.push(new Paragraph({
      alignment: C, spacing: { after: 60 },
      children: [
        new TextRun({ text: decode(strong.text).trim() + ' ', font: SS3, size: 18, bold: true, color: GOLD }),
        new TextRun({ text: rest, font: SS3, size: 18, color: 'DCE3F4' }),
      ],
    }));
  }
  kids.push(new Paragraph({
    alignment: C, spacing: { before: 260, after: 200 },
    children: [new ImageRun({ type: 'png', data: fs.readFileSync(__dirname + '/cover-badge.png'), transformation: { width: 96, height: 96 } })],
  }));
  kids.push(new Paragraph({
    alignment: C, spacing: { after: 40 },
    children: [new TextRun({ text: 'Lawrence Martin', font: OSWALD, size: 22, bold: true, allCaps: true, characterSpacing: 45, color: CREAM_TXT })],
  }));
  kids.push(new Paragraph({
    alignment: C, spacing: { after: 600 },
    children: [new TextRun({ text: 'Father of Three', font: OSWALD, size: 15, allCaps: true, characterSpacing: 30, color: GOLD })],
  }));
  out.push(new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [CW],
    borders: allNone,
    rows: [new TableRow({
      height: { value: 13400, rule: HeightRule.ATLEAST },
      children: [new TableCell({
        width: { size: CW, type: WidthType.DXA },
        shading: shade(NAVY), verticalAlign: VerticalAlign.CENTER,
        margins: { top: 300, bottom: 300, left: 500, right: 500 },
        borders: noBorders, children: kids,
      })],
    })],
  }));
}

// Bookmark anchors for the clickable TOC: hero label text → bookmark id
// ("Introduction" → "toc_introduction", "Chapter Three" → "toc_chapter_three").
const heroAnchor = (labelText) => 'toc_' + labelText.toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');

function chHero(node) {
  const st = styleOf(node);
  const bg = hex(st.background || NAVY);
  const fg = hex(st.color || CREAM_TXT);
  const label = node.querySelector('.ch-hero-label');
  const title = node.querySelector('.ch-hero-title');
  const desc = node.querySelector('.ch-hero-desc');
  const kids = [];
  if (label) kids.push(new Paragraph({
    spacing: { after: 100 },
    children: [new TextRun({ text: decode(label.text).trim(), font: OSWALD, size: 16, bold: true, allCaps: true, characterSpacing: 45, color: blend(fg, bg, 0.65) })],
  }));
  if (title) {
    const titleRun = new TextRun({ text: decode(title.text).trim(), font: OSWALD, size: 44, bold: true, allCaps: true, color: fg });
    kids.push(new Paragraph({
      spacing: { after: 120 },
      children: label
        ? [new Bookmark({ id: heroAnchor(decode(label.text).trim()), children: [titleRun] })]
        : [titleRun],
    }));
  }
  if (desc) kids.push(new Paragraph({
    spacing: { after: 0 },
    children: runsOf(desc, { font: SS3, size: 20, color: blend(fg, bg, 0.85) }),
  }));
  if ((node.getAttribute('class') || '').includes('new-page')) out.push(pageBreak());
  out.push(panel(kids, { fill: bg, margins: { top: 420, bottom: 380, left: 400, right: 400 } }));
  out.push(spacer(200));
}

function heading2(node) {
  out.push(new Paragraph({
    spacing: { before: 260, after: 120 }, keepNext: true,
    children: [new TextRun({ text: decode(node.text).trim(), font: OSWALD, size: 28, bold: true, allCaps: true, color: NAVY })],
  }));
}
function heading4(node, opts = {}) {
  out.push(new Paragraph({
    spacing: { before: 160, after: 80 }, keepNext: true,
    children: runsOf(node, { font: SS3, size: 19, bold: true, color: NAVY, ...opts }),
  }));
}

function para(node) {
  const lede = (node.getAttribute('class') || '').includes('lede');
  out.push(new Paragraph({
    spacing: { after: 130 },
    children: runsOf(node, lede
      ? { font: LORA, size: 22, color: '1C2333', strongColor: NAVY, strongShading: shade('FBE9BE') }
      : { font: LORA, size: 19, color: BODY, strongColor: NAVY }),
  }));
}

function pullquote(node) {
  for (const p of node.querySelectorAll('p')) {
    out.push(new Paragraph({
      spacing: { before: 140, after: 140 },
      border: { left: { style: BorderStyle.SINGLE, size: 28, color: GOLD } },
      shading: shade('FEF9EC'),
      indent: { left: 260 },
      children: runsOf(p, { font: LORA, size: 20, italics: true, color: NAVY }),
    }));
  }
}

const CALLOUT_STYLES = {
  cg: { fill: 'EEF3FB', border: 'C4D3EE', label: BLUE, text: BODY, strong: NAVY },
  ca: { fill: 'FBF4E0', border: 'EAD394', label: '8B6914', text: BODY, strong: NAVY },
  cr: { fill: 'FBEBEA', border: 'E8B8B0', label: 'B23A2C', text: BODY, strong: NAVY },
  cd: { fill: NAVY, border: NAVY, label: GOLD, text: 'C6D2EE', strong: CREAM_TXT },
  cwork: { fill: NAVY, border: '223058', label: GOLD, text: 'DCE3F4', strong: CREAM_TXT, goldLeft: true },
};

function callout(node) {
  const classes = (node.getAttribute('class') || '').split(/\s+/);
  const variant = classes.find((c) => CALLOUT_STYLES[c]) || 'cg';
  const v = CALLOUT_STYLES[variant];
  const st = styleOf(node);
  const center = (st['text-align'] || '') === 'center';
  const kids = [];
  for (const child of els(node)) {
    const cc = cls1(child);
    const align = center ? AlignmentType.CENTER : undefined;
    if (cc === 'callout-label') {
      kids.push(new Paragraph({
        alignment: align, spacing: { after: 80 },
        children: [new TextRun({ text: decode(child.text).trim(), font: OSWALD, size: 14, bold: true, allCaps: true, characterSpacing: 22, color: v.label })],
      }));
    } else if (child.tagName === 'P') {
      const cst = styleOf(child);
      const italic = (cst['font-style'] || '') === 'italic';
      kids.push(new Paragraph({
        alignment: align, spacing: { after: 60 },
        children: runsOf(child, {
          font: italic ? LORA : SS3,
          size: cst['font-size'] ? Math.round(parseFloat(cst['font-size']) * 1.5) : 17,
          italics: italic || undefined,
          color: cst.color ? hex(cst.color) : v.text, strongColor: v.strong,
        }),
      }));
    } else if (child.tagName === 'UL') {
      for (const li of child.querySelectorAll('li')) {
        kids.push(new Paragraph({
          spacing: { after: 40 }, indent: { left: 220, hanging: 160 },
          children: [new TextRun({ text: '•  ', font: SS3, size: 17, color: v.label }),
            ...runsOf(li, { font: SS3, size: 17, color: v.text, strongColor: v.strong })],
        }));
      }
    }
  }
  const borders = {
    top: bd(v.border), bottom: bd(v.border), right: bd(v.border),
    left: v.goldLeft ? bd(GOLD, 28) : bd(v.border),
  };
  out.push(spacer(60));
  out.push(panel(kids, { fill: v.fill, borders }));
  out.push(spacer(100));
}

function scriptBox(node) {
  const head = node.querySelector('.script-head');
  const rows = [new TableRow({
    children: [new TableCell({
      width: { size: CW, type: WidthType.DXA }, shading: shade(NAVY),
      margins: { top: 110, bottom: 110, left: 220, right: 220 },
      children: [new Paragraph({
        spacing: { after: 0 },
        children: [new TextRun({ text: decode(head.text).trim(), font: OSWALD, size: 14, bold: true, allCaps: true, characterSpacing: 30, color: GOLD })],
      })],
    })],
  })];
  const bodyKids = [];
  const body = node.querySelector('.script-body');
  const SAY = { 'script-say': ['EEF3FB', NAVY], 'script-dont': ['FBEBEA', '7A2E22'], 'script-quote': ['FBF4E0', '3A2F0E'] };
  for (const line of els(body)) {
    if (cls1(line) === 'script-line') {
      for (const part of els(line)) {
        const pc = cls1(part);
        if (pc === 'script-who') bodyKids.push(new Paragraph({
          spacing: { before: 90, after: 40 }, keepNext: true,
          children: [new TextRun({ text: decode(part.text).trim(), font: OSWALD, size: 14, bold: true, allCaps: true, characterSpacing: 22, color: MUTED })],
        }));
        else if (SAY[pc]) bodyKids.push(new Paragraph({
          spacing: { after: 60 }, shading: shade(SAY[pc][0]), indent: { left: 120, right: 120 },
          children: runsOf(part, { font: SS3, size: pc === 'script-quote' ? 19 : 17, italics: pc === 'script-quote' || undefined, color: SAY[pc][1] }),
        }));
      }
    } else if (cls1(line) === 'script-why') {
      bodyKids.push(new Paragraph({
        spacing: { before: 60, after: 0 },
        children: runsOf(line, { font: SS3, size: 16, color: MUTED, strongColor: NAVY }),
      }));
    }
  }
  rows.push(new TableRow({
    children: [new TableCell({
      width: { size: CW, type: WidthType.DXA },
      margins: { top: 140, bottom: 140, left: 220, right: 220 },
      children: bodyKids,
    })],
  }));
  out.push(spacer(60));
  out.push(new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: [CW],
    borders: { top: bd(NAVY, 16), bottom: bd(NAVY, 16), left: bd(NAVY, 16), right: bd(NAVY, 16), insideHorizontal: NONE, insideVertical: NONE },
    rows,
  }));
  out.push(spacer(100));
}

function statsRow(node) {
  const cards = node.querySelectorAll('.stat-card');
  const w = Math.floor(CW / cards.length);
  const widths = cards.map((_, i) => (i === cards.length - 1 ? CW - w * (cards.length - 1) : w));
  out.push(spacer(60));
  out.push(new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: widths,
    borders: { ...allNone, insideVertical: { style: BorderStyle.SINGLE, size: 24, color: 'FFFFFF' } },
    rows: [new TableRow({
      children: cards.map((card, i) => new TableCell({
        width: { size: widths[i], type: WidthType.DXA },
        shading: shade(NAVY), verticalAlign: VerticalAlign.CENTER,
        margins: { top: 160, bottom: 140, left: 120, right: 120 },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER, spacing: { after: 50 },
            children: [new TextRun({ text: decode(card.querySelector('.stat-num').text).trim(), font: OSWALD, size: 36, bold: true, color: GOLD })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER, spacing: { after: 0 },
            children: [new TextRun({ text: decode(card.querySelector('.stat-label').text).trim(), font: SS3, size: 14, color: 'C6D2EE' })],
          }),
          ...(card.querySelector('.stat-source') ? [new Paragraph({
            alignment: AlignmentType.CENTER, spacing: { before: 40, after: 0 },
            children: [new TextRun({ text: decode(card.querySelector('.stat-source').text).trim(), font: SS3, size: 12, color: MUTED })],
          })] : []),
        ],
      })),
    })],
  }));
  out.push(spacer(60));
}

function statSource(node) {
  out.push(new Paragraph({
    spacing: { after: 140 },
    children: [new TextRun({ text: decode(node.text).trim(), font: SS3, size: 13, color: MUTED })],
  }));
}

function statBarRow(node) {
  for (const bar of node.querySelectorAll('.stat-bar')) {
    const label = bar.querySelector('.stat-bar-label span');
    const value = bar.querySelector('.stat-bar-label strong');
    const fill = bar.querySelector('.stat-bar-fill');
    const pct = parseFloat((styleOf(fill).width || '50%')) / 100;
    const navy = (fill.getAttribute('class') || '').includes('navy');
    const w1 = Math.max(200, Math.round(CW * pct));
    out.push(new Paragraph({
      spacing: { before: 90, after: 50 }, keepNext: true,
      tabStops: [{ type: TabStopType.RIGHT, position: CW }],
      children: [
        new TextRun({ text: decode(label.text).trim(), font: SS3, size: 16, color: BODY }),
        new TextRun({ text: '\t' + decode(value.text).trim(), font: OSWALD, size: 18, bold: true, color: NAVY }),
      ],
    }));
    out.push(new Table({
      width: { size: CW, type: WidthType.DXA }, columnWidths: [w1, CW - w1],
      borders: allNone,
      rows: [new TableRow({
        height: { value: 150, rule: HeightRule.EXACT },
        children: [
          new TableCell({ width: { size: w1, type: WidthType.DXA }, shading: shade(navy ? NAVY : GOLD), borders: noBorders, margins: { top: 0, bottom: 0, left: 0, right: 0 }, children: [spacer(0)] }),
          new TableCell({ width: { size: CW - w1, type: WidthType.DXA }, shading: shade('E9E3D3'), borders: noBorders, margins: { top: 0, bottom: 0, left: 0, right: 0 }, children: [spacer(0)] }),
        ],
      })],
    }));
  }
  out.push(spacer(140));
}

function caseStudy(node) {
  const tag = node.querySelector('.case-tag');
  const title = node.querySelector('h4');
  const before = node.querySelector('.case-before');
  const after = node.querySelector('.case-after');
  const half = CW / 2;
  const cell = (el, fill, labelColor) => {
    const strong = el.querySelector('strong');
    const rest = decode(el.text).replace(decode(strong.text), '').trim();
    return new TableCell({
      width: { size: half, type: WidthType.DXA }, shading: shade(fill),
      margins: { top: 140, bottom: 140, left: 170, right: 170 },
      borders: noBorders,
      children: [
        new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: decode(strong.text).trim(), font: SS3, size: 13, bold: true, allCaps: true, characterSpacing: 22, color: labelColor })] }),
        new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: rest, font: SS3, size: 16, color: BODY })] }),
      ],
    });
  };
  out.push(spacer(60));
  out.push(new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: [half, half],
    borders: { top: bd('D6DCEB', 12), bottom: bd('D6DCEB', 12), left: bd('D6DCEB', 12), right: bd('D6DCEB', 12), insideHorizontal: NONE, insideVertical: { style: BorderStyle.SINGLE, size: 16, color: 'FFFFFF' } },
    rows: [
      new TableRow({
        children: [new TableCell({
          columnSpan: 2, width: { size: CW, type: WidthType.DXA },
          margins: { top: 140, bottom: 40, left: 170, right: 170 }, borders: noBorders,
          children: [
            new Paragraph({
              spacing: { after: 70 },
              children: [new TextRun({ text: ' ' + decode(tag.text).trim() + ' ', font: OSWALD, size: 13, bold: true, allCaps: true, characterSpacing: 22, color: BLUE, shading: shade('EEF3FB') })],
            }),
            new Paragraph({
              spacing: { after: 60 },
              children: [new TextRun({ text: decode(title.text).trim(), font: OSWALD, size: 21, bold: true, allCaps: true, color: NAVY })],
            }),
          ],
        })],
      }),
      new TableRow({ children: [cell(before, 'FBEBEA', 'B23A2C'), cell(after, 'EEF3FB', BLUE)] }),
    ],
  }));
  out.push(spacer(100));
}

function daySchedule(node) {
  const head = node.querySelector('.sched-head');
  const st = styleOf(head);
  const bg = hex(st.background || NAVY);
  const T = 1500, C = CW - T;
  const rows = [new TableRow({
    children: [new TableCell({
      columnSpan: 2, width: { size: CW, type: WidthType.DXA }, shading: shade(bg),
      margins: { top: 130, bottom: 130, left: 220, right: 220 }, borders: noBorders,
      children: [
        new Paragraph({ spacing: { after: 20 }, children: [new TextRun({ text: decode(head.querySelector('.sched-head-day').text).trim(), font: OSWALD, size: 13, bold: true, allCaps: true, characterSpacing: 45, color: blend(CREAM_TXT, bg, 0.7) })] }),
        new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: decode(head.querySelector('.sched-head-name').text).trim(), font: OSWALD, size: 24, bold: true, allCaps: true, color: CREAM_TXT })] }),
      ],
    })],
  })];
  for (const item of node.querySelectorAll('.sched-item')) {
    const time = item.querySelector('.sched-time');
    const content = item.querySelector('.sched-content');
    const strong = content.querySelector('strong');
    const contentKids = [];
    if (strong) contentKids.push(new Paragraph({
      spacing: { after: 20 }, keepNext: true,
      children: [new TextRun({ text: decode(strong.text).trim(), font: SS3, size: 16, bold: true, color: NAVY })],
    }));
    for (const p of content.querySelectorAll('p')) contentKids.push(new Paragraph({
      spacing: { after: 0 },
      children: runsOf(p, { font: SS3, size: 16, color: BODY, strongColor: NAVY }),
    }));
    rows.push(new TableRow({
      cantSplit: true,
      children: [
        new TableCell({
          width: { size: T, type: WidthType.DXA }, shading: shade('F8F4E8'),
          margins: { top: 100, bottom: 100, left: 140, right: 100 },
          borders: { ...noBorders, bottom: bd('EFE9DA', 6) },
          children: [new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: decode(time.text).trim(), font: SS3, size: 14, bold: true, color: TAN })] })],
        }),
        new TableCell({
          width: { size: C, type: WidthType.DXA },
          margins: { top: 100, bottom: 100, left: 170, right: 170 },
          borders: { ...noBorders, bottom: bd('EFE9DA', 6) },
          children: contentKids,
        }),
      ],
    }));
  }
  out.push(spacer(60));
  out.push(new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: [T, C],
    borders: { top: bd('E4DCC8', 10), bottom: bd('E4DCC8', 10), left: bd('E4DCC8', 10), right: bd('E4DCC8', 10), insideHorizontal: NONE, insideVertical: NONE },
    rows,
  }));
  out.push(spacer(100));
}

function twoGrid(node) {
  const cards = els(node);
  const half = CW / 2;
  const rows = [];
  for (let i = 0; i < cards.length; i += 2) {
    const pair = [cards[i], cards[i + 1]].filter(Boolean);
    rows.push(new TableRow({
      cantSplit: true,
      children: pair.map((card) => {
        const isA = cls1(card) === 'a-card';
        const st = styleOf(card);
        const fill = isA ? hex(st.background || 'F0EEE4') : 'F0EEE4';
        const border = isA ? hex(st['border-color'] || 'D9D3BF') : 'F0EEE4';
        const kids = [];
        const h4 = card.querySelector('h4');
        if (h4) kids.push(new Paragraph({
          spacing: { after: 50 }, keepNext: true,
          children: [new TextRun({ text: decode(h4.text).trim(), font: SS3, size: 18, bold: true, color: NAVY })],
        }));
        for (const p of card.querySelectorAll('p')) kids.push(new Paragraph({
          spacing: { after: 0 },
          children: runsOf(p, { font: SS3, size: 16, color: '3C4458', strongColor: NAVY }),
        }));
        return new TableCell({
          width: { size: half, type: WidthType.DXA }, shading: shade(fill),
          margins: { top: 140, bottom: 140, left: 190, right: 190 },
          borders: { top: bd(border, 10), bottom: bd(border, 10), left: bd(border, 10), right: bd(border, 10) },
          children: kids,
        });
      }).concat(pair.length === 1 ? [new TableCell({ width: { size: half, type: WidthType.DXA }, borders: noBorders, children: [spacer(0)] })] : []),
    }));
  }
  out.push(spacer(60));
  out.push(new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: [half, half],
    borders: { ...allNone, insideHorizontal: { style: BorderStyle.SINGLE, size: 16, color: 'FFFFFF' }, insideVertical: { style: BorderStyle.SINGLE, size: 16, color: 'FFFFFF' } },
    rows,
  }));
  out.push(spacer(100));
}

function ciParagraph(el, size, last) {
  return new Paragraph({
    spacing: { after: 40, before: 40 },
    border: last ? undefined : { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'EFE9DA' } },
    indent: { left: 280, hanging: 280 },
    children: [
      new TextRun({ text: '☐', font: 'Segoe UI Symbol', size: size + 2, color: FILL_LINE }),
      new TextRun({ text: '  ', size }),
      ...runsOf(el, { font: SS3, size, color: BODY, strongColor: NAVY }),
    ],
  });
}

function cklistBox(node) {
  const kids = [];
  const children = els(node);
  const cis = children.filter((c) => cls1(c) === 'ci');
  let i = 0;
  for (const child of children) {
    if (child.tagName === 'H4') kids.push(new Paragraph({
      spacing: { after: 90 },
      children: [new TextRun({ text: decode(child.text).trim(), font: OSWALD, size: 14, bold: true, allCaps: true, characterSpacing: 22, color: MUTED })],
    }));
    else if (cls1(child) === 'ci') { i++; kids.push(ciParagraph(child, 17, i === cis.length)); }
    else if (cls1(child) === 'fill-text' || child.tagName === 'P') kids.push(new Paragraph({
      spacing: { after: 40 },
      children: runsOf(child, { font: SS3, size: 16, color: BODY, strongColor: NAVY }),
    }));
  }
  out.push(spacer(60));
  out.push(panel(kids, { fill: 'F8F4E8' }));
  out.push(spacer(100));
}

function fillLinePara(node) {
  const st = node ? styleOf(node) : {};
  const dark = (st['border-color'] || '').includes('101B33');
  return new Paragraph({
    spacing: { before: 60, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: dark ? 16 : 10, color: dark ? NAVY : FILL_LINE } },
    children: [new TextRun({ text: '', size: 16 })],
  });
}

function tmplSection(section, width) {
  const kids = [];
  const children = els(section);
  const cis = children.filter((c) => cls1(c) === 'ci');
  let ci = 0;
  for (const child of children) {
    const cc = cls1(child);
    if (child.tagName === 'H5') kids.push(new Paragraph({
      spacing: { before: kids.length ? 100 : 0, after: 60 }, keepNext: true,
      children: [new TextRun({ text: decode(child.text).trim(), font: OSWALD, size: 13, bold: true, allCaps: true, characterSpacing: 22, color: MUTED })],
    }));
    else if (cc === 'ci') { ci++; kids.push(ciParagraph(child, 16, ci === cis.length)); }
    else if (cc === 'fill-line') kids.push(fillLinePara(child));
    else if (cc === 'fill-text' || child.tagName === 'P') {
      const st = styleOf(child);
      kids.push(new Paragraph({
        spacing: { after: 40, line: 320, lineRule: 'auto' },
        children: runsOf(child, {
          font: SS3, size: st['font-size'] ? Math.round(parseFloat(st['font-size']) * 1.5) : 16,
          color: st.color ? hex(st.color) : '3C4458', strongColor: NAVY,
        }),
      }));
    }
  }
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    margins: { top: 120, bottom: 120, left: 150, right: 150 },
    borders: { top: bd('D9D3BF', 6), bottom: bd('D9D3BF', 6), left: bd('D9D3BF', 6), right: bd('D9D3BF', 6) },
    children: kids,
  });
}

function tmpl(node) {
  const kids = [];
  for (const child of els(node)) {
    const cc = cls1(child);
    if (cc === 'tmpl-title') kids.push(new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { after: 40 },
      children: [new TextRun({ text: decode(child.text).trim(), font: OSWALD, size: 21, bold: true, allCaps: true, characterSpacing: 15, color: NAVY })],
    }));
    else if (cc === 'tmpl-sub') kids.push(new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { after: 160 },
      children: runsOf(child, { font: SS3, size: 15, color: MUTED }),
    }));
    else if (child.tagName === 'P') {
      const st = styleOf(child);
      kids.push(new Paragraph({
        alignment: st['text-align'] === 'center' ? AlignmentType.CENTER : undefined,
        spacing: { after: 100 },
        children: runsOf(child, { font: SS3, size: st['font-size'] ? Math.round(parseFloat(st['font-size']) * 1.5) : 16, color: st.color ? hex(st.color) : BODY, strongColor: NAVY }),
      }));
    } else if (cc === 'fill-line') kids.push(fillLinePara(child));
    else if (cc === 'tmpl-grid') {
      const sections = els(child);
      const half = (CW - 400) / 2;
      const rows = [];
      for (let i = 0; i < sections.length; i += 2) {
        const pair = [sections[i], sections[i + 1]].filter(Boolean);
        rows.push(new TableRow({
          children: pair.map((s) => tmplSection(s, half)).concat(
            pair.length === 1 ? [new TableCell({ width: { size: half, type: WidthType.DXA }, borders: noBorders, children: [spacer(0)] })] : []),
        }));
      }
      kids.push(new Table({
        width: { size: CW - 400, type: WidthType.DXA }, columnWidths: [half, half],
        borders: { ...allNone, insideHorizontal: { style: BorderStyle.SINGLE, size: 16, color: 'FFFFFF' }, insideVertical: { style: BorderStyle.SINGLE, size: 16, color: 'FFFFFF' } },
        rows,
      }));
      kids.push(spacer(80));
    } else if (cc === 'sig-row') {
      const blocks = els(child);
      const w = Math.floor((CW - 400) / blocks.length);
      kids.push(new Table({
        width: { size: CW - 400, type: WidthType.DXA },
        columnWidths: blocks.map((_, i) => (i === blocks.length - 1 ? CW - 400 - w * (blocks.length - 1) : w)),
        borders: { ...allNone, insideVertical: { style: BorderStyle.SINGLE, size: 24, color: 'FFFFFF' } },
        rows: [new TableRow({
          children: blocks.map((b) => new TableCell({
            width: { size: w, type: WidthType.DXA }, borders: noBorders,
            margins: { top: 80, bottom: 20, left: 60, right: 60 },
            children: [
              new Paragraph({ spacing: { after: 20 }, children: [new TextRun({ text: decode(b.querySelector('label').text).trim(), font: SS3, size: 12, color: TAN })] }),
              fillLinePara(null),
            ],
          })),
        })],
      }));
    }
  }
  out.push(spacer(80));
  out.push(panel(kids, {
    fill: 'FFFFFF',
    borders: { top: bd(NAVY, 16), bottom: bd(NAVY, 16), left: bd(NAVY, 16), right: bd(NAVY, 16), insideHorizontal: NONE, insideVertical: NONE },
    margins: { top: 220, bottom: 200, left: 280, right: 280 },
  }));
  out.push(spacer(100));
}

function convoGrid(node) {
  const cards = els(node);
  const half = CW / 2;
  const rows = [];
  for (let i = 0; i < cards.length; i += 2) {
    const pair = [cards[i], cards[i + 1]].filter(Boolean);
    rows.push(new TableRow({
      cantSplit: true,
      children: pair.map((card) => new TableCell({
        width: { size: half, type: WidthType.DXA }, shading: shade('EEF3FB'),
        margins: { top: 120, bottom: 120, left: 160, right: 160 },
        borders: noBorders,
        children: [new Paragraph({
          spacing: { after: 0 },
          children: [
            new TextRun({ text: '“ ', font: LORA, size: 24, color: blend(BLUE, 'EEF3FB', 0.5) }),
            ...runsOf(card, { font: SS3, size: 16, italics: true, color: NAVY }),
          ],
        })],
      })),
    }));
  }
  out.push(spacer(60));
  out.push(new Table({
    width: { size: CW, type: WidthType.DXA }, columnWidths: [half, half],
    borders: { ...allNone, insideHorizontal: { style: BorderStyle.SINGLE, size: 12, color: 'FFFFFF' }, insideVertical: { style: BorderStyle.SINGLE, size: 12, color: 'FFFFFF' } },
    rows,
  }));
  out.push(spacer(100));
}

// "Ch. 3" → the matching hero's bookmark ("toc_chapter_three"); "—" → the
// Introduction hero. Page numbers are PAGEREF fields, so they always show
// the chapter's real page and click through to it.
const NUM_WORDS = ['one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine'];
function tocAnchor(chText) {
  const m = chText.match(/(\d+)/);
  return m ? 'toc_chapter_' + NUM_WORDS[parseInt(m[1], 10) - 1] : 'toc_introduction';
}

function tocList(node) {
  for (const li of node.querySelectorAll('.toc-item')) {
    const ch = li.querySelector('.toc-ch');
    const name = li.querySelector('.toc-name');
    const anchor = tocAnchor(decode(ch.text).trim());
    out.push(new Paragraph({
      spacing: { before: 80, after: 80 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE } },
      tabStops: [
        { type: TabStopType.LEFT, position: 1250 },
        { type: TabStopType.RIGHT, position: CW, leader: LeaderType.DOT },
      ],
      children: [
        new TextRun({ text: decode(ch.text).trim(), font: OSWALD, size: 16, bold: true, allCaps: true, characterSpacing: 15, color: NAVY }),
        new TextRun({ text: '\t' }),
        new InternalHyperlink({
          anchor,
          children: [new TextRun({ text: decode(name.text).trim(), font: SS3, size: 18, color: BODY, underline: false })],
        }),
        new TextRun({ text: '\t' }),
        new PageReference(anchor, { hyperlink: true }),
      ],
    }));
  }
}

function footerDiv(node) {
  out.push(new Paragraph({
    spacing: { before: 300, after: 60 }, alignment: AlignmentType.CENTER,
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: LINE } },
    children: [new TextRun({ text: 'Lawrence Martin', font: OSWALD, size: 20, bold: true, allCaps: true, color: NAVY })],
  }));
  const tail = node.childNodes.filter((c) => c.nodeType === 3).map((c) => decode(c.rawText).trim()).join(' ').replace(/\s+/g, ' ').trim();
  out.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 0 },
    children: [new TextRun({ text: tail, font: SS3, size: 16, color: TAN })],
  }));
}

// ---------- walk the document ----------
const page = fs.readFileSync(__dirname + '/screens-down-family-up-ebook-v3.html', 'utf8');
const root = parse(page).querySelector('doc-page');

for (const node of els(root)) {
  const cc = cls1(node);
  const tag = node.tagName;
  if (cc === 'cover') cover(node);
  else if (cc === 'ch-hero') chHero(node);
  else if (tag === 'H2') heading2(node);
  else if (tag === 'H3' || tag === 'H4') heading4(node);
  else if (tag === 'P') para(node);
  else if (cc === 'pullquote') pullquote(node);
  else if (cc === 'callout') callout(node);
  else if (cc === 'script-box') scriptBox(node);
  else if (cc === 'stats-row') statsRow(node);
  else if (cc === 'stat-source') statSource(node);
  else if (cc === 'stat-bar-row') statBarRow(node);
  else if (cc === 'case-study') caseStudy(node);
  else if (cc === 'day-schedule') daySchedule(node);
  else if (cc === 'two-grid') twoGrid(node);
  else if (cc === 'cklist-box') cklistBox(node);
  else if (cc === 'tmpl') tmpl(node);
  else if (cc === 'convo-grid') convoGrid(node);
  else if (tag === 'UL' && cc === 'toc-list') tocList(node);
  else if (tag === 'UL') { // generic list
    for (const li of node.querySelectorAll('li')) out.push(new Paragraph({
      spacing: { after: 50 }, indent: { left: 320, hanging: 200 },
      children: [new TextRun({ text: '•  ', font: SS3, size: 18, color: GOLD }),
        ...runsOf(li, { font: SS3, size: 18, color: BODY, strongColor: NAVY })],
    }));
  }
  else if (tag === 'DIV' && (node.getAttribute('style') || '').includes('border-top')) footerDiv(node);
  else if (tag === 'DIV') { // unknown styled div — render paragraphs inside
    for (const p of node.querySelectorAll('p')) out.push(new Paragraph({
      spacing: { after: 100 },
      children: runsOf(p, { font: LORA, size: 19, color: BODY, strongColor: NAVY }),
    }));
  }
}

const doc = new Document({
  creator: 'Lawrence Martin',
  title: 'Screens Down, Family Up — The 7-Day Reset',
  features: { updateFields: true }, // Word refreshes the PAGEREF fields on open
  styles: {
    default: {
      document: { run: { font: SS3, size: 19, color: BODY } },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 },
      },
    },
    children: out,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(__dirname + '/screens-down-family-up-ebook-v3.docx', buf);
  console.log('wrote ebook.docx', buf.length, 'bytes,', out.length, 'blocks');
});
