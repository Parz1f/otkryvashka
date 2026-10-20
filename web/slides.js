/*
 * Открывашка: отрисовка презентаций PowerPoint (.pptx).
 *
 * Слайд собирается из HTML и SVG по разметке файла: фигуры, текст, картинки, таблицы,
 * группы, SmartArt и графики. Графики рисуются по данным, сохранённым в самом файле.
 * Из сторонних библиотек нужен только JSZip, чтобы распаковать .pptx.
 *
 * Использование: const view = await PptxView.render(arrayBuffer, hostElement, { width });
 *                view.count — число слайдов; view.setWidth(px); view.dispose().
 */
(function () {
'use strict';

const EMU_PX = 1 / 9525;           // 914400 EMU в дюйме, 96 px в дюйме
const PT_PX = 4 / 3;
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/* ---------- XML ---------- */
function kids(el, name) {
  const out = [];
  if (el) for (const c of el.children) if (!name || c.localName === name) out.push(c);
  return out;
}
function kid(el, name) {
  if (el) for (const c of el.children) if (c.localName === name) return c;
  return null;
}
function q(el, path) {
  for (const p of path.split('/')) { el = kid(el, p); if (!el) return null; }
  return el;
}
function attr(el, name, def) {
  if (!el) return def;
  const v = el.getAttribute(name);
  return v == null ? def : v;
}
function num(el, name, def) {
  const v = attr(el, name, null);
  return v == null || v === '' || isNaN(+v) ? def : +v;
}
function bool(el, name, def) {
  const v = attr(el, name, null);
  return v == null ? def : v === '1' || v === 'true' || v === 'on';
}
function rid(el, name) {
  return el ? (el.getAttributeNS(R_NS, name) || el.getAttribute('r:' + name)) : null;
}
function desc(el, name) { return el ? [...el.getElementsByTagNameNS('*', name)] : []; }
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const r2 = (n) => Math.round(n * 100) / 100;

/* ---------- Package: parts, relationships, media ---------- */
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', jpe: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff', ico: 'image/x-icon' };
const BROWSER_IMAGES = new Set(['png', 'jpg', 'jpeg', 'jpe', 'gif', 'bmp', 'svg', 'webp', 'ico']);

function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/');
  parts.pop();
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.' && seg !== '') parts.push(seg);
  }
  return parts.join('/');
}

class Pkg {
  constructor(zip) { this.zip = zip; this.xmls = new Map(); this.relMap = new Map(); this.urls = new Map(); this.made = []; }
  file(path) { return this.zip.file(path) || this.zip.file(decodeURIComponent(path)); }
  xml(path) {
    if (!path) return Promise.resolve(null);
    if (!this.xmls.has(path)) {
      this.xmls.set(path, (async () => {
        const f = this.file(path);
        if (!f) return null;
        const doc = new DOMParser().parseFromString(await f.async('string'), 'application/xml');
        return doc.getElementsByTagName('parsererror').length ? null : doc.documentElement;
      })());
    }
    return this.xmls.get(path);
  }
  rels(path) {
    if (!path) return Promise.resolve({});
    if (!this.relMap.has(path)) {
      this.relMap.set(path, (async () => {
        const i = path.lastIndexOf('/');
        const root = await this.xml(path.slice(0, i) + '/_rels/' + path.slice(i + 1) + '.rels');
        const out = {};
        for (const r of kids(root, 'Relationship')) {
          const external = attr(r, 'TargetMode', '') === 'External';
          out[attr(r, 'Id')] = { type: attr(r, 'Type', ''), target: external ? attr(r, 'Target', '') : resolvePath(path, attr(r, 'Target', '')), external };
        }
        return out;
      })());
    }
    return this.relMap.get(path);
  }
  url(path) {
    if (!this.urls.has(path)) {
      this.urls.set(path, (async () => {
        const f = this.file(path);
        if (!f) return null;
        const ext = path.split('.').pop().toLowerCase();
        const u = URL.createObjectURL(new Blob([await f.async('uint8array')], { type: MIME[ext] || 'application/octet-stream' }));
        this.made.push(u);
        return u;
      })());
    }
    return this.urls.get(path);
  }
  dispose() { this.made.forEach((u) => URL.revokeObjectURL(u)); }
}
function relOfType(rels, suffix) {
  for (const id in rels) if (rels[id].type.endsWith(suffix)) return rels[id];
  return null;
}

/* ---------- Colors ---------- */
const COLOR_TAGS = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr']);
const PRESET = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', cyan: '00FFFF', magenta: 'FF00FF',
  gray: '808080', grey: '808080', darkGray: 'A9A9A9', lightGray: 'D3D3D3', dkGray: 'A9A9A9', ltGray: 'D3D3D3', orange: 'FFA500', navy: '000080',
  purple: '800080', maroon: '800000', olive: '808000', teal: '008080', silver: 'C0C0C0', lime: '00FF00', darkBlue: '00008B', darkRed: '8B0000',
  darkGreen: '006400', gold: 'FFD700', pink: 'FFC0CB', brown: 'A52A2A', violet: 'EE82EE', indigo: '4B0082',
};
const DEFAULT_THEME_COLORS = {
  dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6', accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5',
  accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47', hlink: '0563C1', folHlink: '954F72',
};
const hexRgb = (h) => [parseInt(h.slice(0, 2), 16) || 0, parseInt(h.slice(2, 4), 16) || 0, parseInt(h.slice(4, 6), 16) || 0];
function rgbHsl([r, g, b]) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min, s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}
function hslRgb(h, s, l) {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const f = (p, q2, t) => {
    if (t < 0) t += 1; if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q2 - p) * 6 * t;
    if (t < 1 / 2) return q2;
    if (t < 2 / 3) return p + (q2 - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q2 = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q2;
  return [f(p, q2, h + 1 / 3) * 255, f(p, q2, h) * 255, f(p, q2, h - 1 / 3) * 255];
}
// Оттенки (tint/shade) по спецификации считаются в линейном RGB
const toLinear = (x) => { x /= 255; return x <= 0.04045 ? x / 12.92 : Math.pow((x + 0.055) / 1.055, 2.4); };
const fromLinear = (v) => 255 * (v <= 0.0031308 ? 12.92 * v : 1.055 * Math.pow(Math.max(0, v), 1 / 2.4) - 0.055);
function schemeHex(name, ctx) {
  const m = (ctx.clrMap && ctx.clrMap[name]) || name;
  return ctx.theme.colors[m] || ctx.theme.colors[name] || DEFAULT_THEME_COLORS[m] || '000000';
}
function colorOf(el, ctx, phClr) {
  if (!el) return null;
  let c = COLOR_TAGS.has(el.localName) ? el : null;
  if (!c) for (const ch of el.children) if (COLOR_TAGS.has(ch.localName)) { c = ch; break; }
  if (!c) return null;
  let rgb = [0, 0, 0], a = 1;
  switch (c.localName) {
    case 'srgbClr': rgb = hexRgb(attr(c, 'val', '000000')); break;
    case 'schemeClr': {
      const v = attr(c, 'val', 'tx1');
      if (v === 'phClr') { if (phClr) { rgb = phClr.rgb.slice(); a = phClr.a; } }
      else rgb = hexRgb(schemeHex(v, ctx));
      break;
    }
    case 'sysClr': rgb = hexRgb(attr(c, 'lastClr', attr(c, 'val') === 'window' ? 'FFFFFF' : '000000')); break;
    case 'prstClr': rgb = hexRgb(PRESET[attr(c, 'val')] || '000000'); break;
    case 'scrgbClr': rgb = ['r', 'g', 'b'].map((k) => { const v = num(c, k, 0) / 100000; return 255 * (v <= 0.0031308 ? 12.92 * v : 1.055 * Math.pow(v, 1 / 2.4) - 0.055); }); break;
    case 'hslClr': rgb = hslRgb(num(c, 'hue', 0) / 21600000, num(c, 'sat', 0) / 100000, num(c, 'lum', 0) / 100000); break;
  }
  for (const m of c.children) {
    const v = num(m, 'val', 0) / 100000;
    switch (m.localName) {
      case 'alpha': a = v; break;
      case 'alphaMod': a *= v; break;
      case 'lumMod': { const h = rgbHsl(rgb); rgb = hslRgb(h[0], h[1], Math.min(1, h[2] * v)); break; }
      case 'lumOff': { const h = rgbHsl(rgb); rgb = hslRgb(h[0], h[1], Math.max(0, Math.min(1, h[2] + v))); break; }
      case 'satMod': { const h = rgbHsl(rgb); rgb = hslRgb(h[0], Math.min(1, h[1] * v), h[2]); break; }
      case 'tint': rgb = rgb.map((x) => fromLinear(toLinear(x) + (1 - toLinear(x)) * (1 - v))); break;
      case 'shade': rgb = rgb.map((x) => fromLinear(toLinear(x) * v)); break;
      case 'inv': rgb = rgb.map((x) => 255 - x); break;
      case 'gray': { const g = rgb[0] * 0.3 + rgb[1] * 0.59 + rgb[2] * 0.11; rgb = [g, g, g]; break; }
    }
  }
  return { rgb: rgb.map((x) => Math.max(0, Math.min(255, Math.round(x)))), a };
}
function css(c) {
  if (!c) return 'none';
  const [r, g, b] = c.rgb;
  return c.a >= 1 ? `rgb(${r},${g},${b})` : `rgba(${r},${g},${b},${+c.a.toFixed(3)})`;
}
const BLACK = { rgb: [0, 0, 0], a: 1 };
function mix(c, white) { return { rgb: c.rgb.map((x) => x + (255 - x) * white), a: c.a }; }
function darken(c, k) { return { rgb: c.rgb.map((x) => x * k), a: c.a }; }

/* ---------- Theme ---------- */
async function loadTheme(pkg, path) {
  const t = { colors: Object.assign({}, DEFAULT_THEME_COLORS), major: 'Calibri Light', minor: 'Calibri', fills: [], lines: [], bgFills: [] };
  const root = await pkg.xml(path);
  const el = kid(root, 'themeElements');
  if (!el) return t;
  for (const c of kids(kid(el, 'clrScheme'))) {
    const v = c.firstElementChild;
    if (!v) continue;
    t.colors[c.localName] = v.localName === 'sysClr' ? attr(v, 'lastClr', attr(v, 'val') === 'window' ? 'FFFFFF' : '000000') : attr(v, 'val', '000000');
  }
  const fs = kid(el, 'fontScheme');
  t.major = attr(q(fs, 'majorFont/latin'), 'typeface', t.major) || t.major;
  t.minor = attr(q(fs, 'minorFont/latin'), 'typeface', t.minor) || t.minor;
  const fmt = kid(el, 'fmtScheme');
  t.fills = kids(kid(fmt, 'fillStyleLst'));
  t.lines = kids(kid(fmt, 'lnStyleLst'));
  t.bgFills = kids(kid(fmt, 'bgFillStyleLst'));
  return t;
}

/* ---------- Fills and lines ---------- */
const FILL_TAGS = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);
function findFill(el) {
  if (el) for (const c of el.children) if (FILL_TAGS.has(c.localName)) return c;
  return null;
}
function parseFill(f, ctx, phClr) {
  if (!f) return null;
  switch (f.localName) {
    case 'noFill': return { type: 'none' };
    case 'solidFill': return { type: 'solid', color: colorOf(f, ctx, phClr) || BLACK };
    case 'gradFill': {
      const stops = kids(kid(f, 'gsLst'), 'gs').map((g) => ({ pos: num(g, 'pos', 0) / 100000, color: colorOf(g, ctx, phClr) || BLACK })).sort((a, b) => a.pos - b.pos);
      if (!stops.length) return null;
      const lin = kid(f, 'lin'), path = kid(f, 'path');
      return { type: 'grad', stops, angle: lin ? num(lin, 'ang', 0) / 60000 : 90, radial: !!path && !lin };
    }
    case 'blipFill': {
      const rel = ctx.rels[rid(kid(f, 'blip'), 'embed')];
      return rel ? { type: 'img', path: rel.target } : null;
    }
    case 'pattFill': {
      const fg = colorOf(kid(f, 'fgClr'), ctx, phClr), bg = colorOf(kid(f, 'bgClr'), ctx, phClr);
      if (!fg) return bg ? { type: 'solid', color: bg } : null;
      // Узор заменяем средним цветом: на маленьком экране он всё равно сливается
      return { type: 'solid', color: bg ? { rgb: fg.rgb.map((x, i) => (x + bg.rgb[i]) / 2), a: 1 } : fg };
    }
    case 'grpFill': return ctx.groupFill || null;
  }
  return null;
}
function styleRef(ref, list, ctx) {
  if (!ref) return null;
  const idx = num(ref, 'idx', 0);
  if (!idx || !list.length) return null;
  return { el: list[Math.min(idx, list.length) - 1], phClr: colorOf(ref, ctx) };
}
function solidOf(fill) {
  if (!fill) return null;
  if (fill.type === 'solid') return fill.color;
  if (fill.type === 'grad') return fill.stops[Math.floor(fill.stops.length / 2)].color;
  return null;
}
function shapeFill(spPrs, style, ctx) {
  for (const spPr of spPrs) {
    const f = findFill(spPr);
    if (f) return parseFill(f, ctx);
  }
  const r = styleRef(kid(style, 'fillRef'), ctx.theme.fills, ctx);
  if (r) return parseFill(r.el, ctx, r.phClr) || (r.phClr ? { type: 'solid', color: r.phClr } : null);
  return null;
}
function shapeLine(spPrs, style, ctx) {
  let fill = null, w = null, dash = null, head = null, tail = null;
  for (const spPr of spPrs) {
    const ln = kid(spPr, 'ln');
    if (!ln) continue;
    if (!fill) { const f = findFill(ln); if (f) fill = parseFill(f, ctx); }
    if (w == null) w = num(ln, 'w', null);
    if (!dash) dash = attr(kid(ln, 'prstDash'), 'val', null);
    if (!head) head = attr(kid(ln, 'headEnd'), 'type', null);
    if (!tail) tail = attr(kid(ln, 'tailEnd'), 'type', null);
  }
  const r = styleRef(kid(style, 'lnRef'), ctx.theme.lines, ctx);
  if (r && r.el) {
    if (!fill) { const f = findFill(r.el); fill = f ? parseFill(f, ctx, r.phClr) : (r.phClr ? { type: 'solid', color: r.phClr } : null); }
    if (w == null) w = num(r.el, 'w', null);
    if (!dash) dash = attr(kid(r.el, 'prstDash'), 'val', null);
  }
  const color = solidOf(fill);
  if (!color) return null;
  return { color, w: Math.max(0.6, (w == null ? 9525 : w) * EMU_PX), dash, head, tail };
}
async function fillCss(fill, pkg) {
  if (!fill || fill.type === 'none') return 'transparent';
  if (fill.type === 'solid') return css(fill.color);
  if (fill.type === 'grad') {
    const stops = fill.stops.map((s) => `${css(s.color)} ${r2(s.pos * 100)}%`).join(',');
    return fill.radial ? `radial-gradient(circle,${stops})` : `linear-gradient(${r2(fill.angle + 90)}deg,${stops})`;
  }
  if (fill.type === 'img') {
    const u = await pkg.url(fill.path);
    return u ? `url("${u}") center/100% 100% no-repeat` : 'transparent';
  }
  return 'transparent';
}

/* ---------- Shape geometry ---------- */
function adjValues(geom) {
  const adj = [];
  for (const gd of kids(kid(geom, 'avLst'), 'gd')) {
    const m = /^val\s+(-?\d+)/.exec(attr(gd, 'fmla', ''));
    if (!m) continue;
    const name = attr(gd, 'name', 'adj');
    const i = name === 'adj' ? 0 : (parseInt(name.replace(/\D/g, ''), 10) || 1) - 1;
    adj[i] = +m[1];
  }
  return adj;
}
function polygon(n, w, h, inner) {
  const pts = [];
  for (let i = 0; i < n * (inner ? 2 : 1); i++) {
    const ang = -Math.PI / 2 + (i * Math.PI * 2) / (n * (inner ? 2 : 1));
    const k = inner && i % 2 ? inner : 1;
    pts.push(`${r2(w / 2 + (Math.cos(ang) * w / 2) * k)} ${r2(h / 2 + (Math.sin(ang) * h / 2) * k)}`);
  }
  return 'M' + pts.join('L') + 'Z';
}
function roundRectPath(w, h, r, corners) {
  r = Math.max(0, Math.min(r, w / 2, h / 2));
  const [tl, tr, br, bl] = (corners || [1, 1, 1, 1]).map((c) => (c ? r : 0));
  return `M${tl} 0H${w - tr}${tr ? `A${tr} ${tr} 0 0 1 ${w} ${tr}` : ''}V${h - br}${br ? `A${br} ${br} 0 0 1 ${w - br} ${h}` : ''}H${bl}${bl ? `A${bl} ${bl} 0 0 1 0 ${h - bl}` : ''}V${tl}${tl ? `A${tl} ${tl} 0 0 1 ${tl} 0` : ''}Z`;
}
const ELLIPSE = (w, h) => `M0 ${h / 2}A${w / 2} ${h / 2} 0 1 1 ${w} ${h / 2}A${w / 2} ${h / 2} 0 1 1 0 ${h / 2}Z`;
// Возвращает { d, open }: open — линия без заливки (соединители, дуги, скобки)
function presetPath(prst, w, h, adj) {
  const a = (i, d) => (adj[i] != null ? adj[i] : d) / 100000;
  const ss = Math.min(w, h);
  switch (prst) {
    case 'roundRect': case 'flowChartAlternateProcess': return { d: roundRectPath(w, h, ss * a(0, 16667)) };
    case 'round1Rect': return { d: roundRectPath(w, h, ss * a(0, 16667), [0, 1, 0, 0]) };
    case 'round2SameRect': return { d: roundRectPath(w, h, ss * a(0, 16667), [1, 1, 0, 0]) };
    case 'round2DiagRect': return { d: roundRectPath(w, h, ss * a(0, 16667), [1, 0, 1, 0]) };
    case 'flowChartTerminator': return { d: roundRectPath(w, h, Math.min(w, h) / 2) };
    case 'ellipse': case 'flowChartConnector': case 'cloud': case 'cloudCallout': case 'wedgeEllipseCallout': case 'sun': case 'smileyFace': case 'flowChartOnlineStorage': return { d: ELLIPSE(w, h) };
    case 'donut': { const d = ss * a(0, 25000); return { d: ELLIPSE(w, h) + `M${d} ${h / 2}A${w / 2 - d} ${h / 2 - d} 0 1 0 ${w - d} ${h / 2}A${w / 2 - d} ${h / 2 - d} 0 1 0 ${d} ${h / 2}Z`, evenodd: true }; }
    case 'triangle': case 'flowChartExtract': { const x = w * a(0, 50000); return { d: `M${x} 0L${w} ${h}L0 ${h}Z` }; }
    case 'flowChartMerge': return { d: `M0 0L${w} 0L${w / 2} ${h}Z` };
    case 'rtTriangle': return { d: `M0 0L${w} ${h}L0 ${h}Z` };
    case 'diamond': case 'flowChartDecision': return { d: `M${w / 2} 0L${w} ${h / 2}L${w / 2} ${h}L0 ${h / 2}Z` };
    case 'parallelogram': case 'flowChartInputOutput': { const x = prst === 'parallelogram' ? ss * a(0, 25000) : w / 5; return { d: `M${x} 0L${w} 0L${w - x} ${h}L0 ${h}Z` }; }
    case 'trapezoid': { const x = ss * a(0, 25000); return { d: `M0 ${h}L${x} 0L${w - x} 0L${w} ${h}Z` }; }
    case 'flowChartManualOperation': return { d: `M0 0L${w} 0L${w * 0.8} ${h}L${w * 0.2} ${h}Z` };
    case 'flowChartManualInput': return { d: `M0 ${h * 0.2}L${w} 0L${w} ${h}L0 ${h}Z` };
    case 'pentagon': return { d: polygon(5, w, h) };
    case 'hexagon': case 'flowChartPreparation': { const x = prst === 'hexagon' ? ss * a(0, 25000) : w / 5; return { d: `M0 ${h / 2}L${x} 0L${w - x} 0L${w} ${h / 2}L${w - x} ${h}L${x} ${h}Z` }; }
    case 'heptagon': return { d: polygon(7, w, h) };
    case 'octagon': { const x = ss * a(0, 29289); return { d: `M${x} 0L${w - x} 0L${w} ${x}L${w} ${h - x}L${w - x} ${h}L${x} ${h}L0 ${h - x}L0 ${x}Z` }; }
    case 'decagon': return { d: polygon(10, w, h) };
    case 'dodecagon': return { d: polygon(12, w, h) };
    case 'homePlate': { const x = w - ss * a(0, 50000); return { d: `M0 0L${x} 0L${w} ${h / 2}L${x} ${h}L0 ${h}Z` }; }
    case 'chevron': { const d = ss * a(0, 50000); return { d: `M0 0L${w - d} 0L${w} ${h / 2}L${w - d} ${h}L0 ${h}L${d} ${h / 2}Z` }; }
    case 'rightArrow': case 'notchedRightArrow': case 'stripedRightArrow': case 'curvedRightArrow': case 'rightArrowCallout': {
      const t = h * a(0, 50000), hd = ss * a(1, 50000), y1 = (h - t) / 2, y2 = (h + t) / 2, x = w - hd;
      return { d: `M0 ${y1}L${x} ${y1}L${x} 0L${w} ${h / 2}L${x} ${h}L${x} ${y2}L0 ${y2}Z` };
    }
    case 'leftArrow': case 'curvedLeftArrow': case 'leftArrowCallout': {
      const t = h * a(0, 50000), hd = ss * a(1, 50000), y1 = (h - t) / 2, y2 = (h + t) / 2;
      return { d: `M${w} ${y1}L${hd} ${y1}L${hd} 0L0 ${h / 2}L${hd} ${h}L${hd} ${y2}L${w} ${y2}Z` };
    }
    case 'upArrow': case 'curvedUpArrow': case 'upArrowCallout': {
      const t = w * a(0, 50000), hd = ss * a(1, 50000), x1 = (w - t) / 2, x2 = (w + t) / 2;
      return { d: `M${x1} ${h}L${x1} ${hd}L0 ${hd}L${w / 2} 0L${w} ${hd}L${x2} ${hd}L${x2} ${h}Z` };
    }
    case 'downArrow': case 'curvedDownArrow': case 'downArrowCallout': {
      const t = w * a(0, 50000), hd = ss * a(1, 50000), x1 = (w - t) / 2, x2 = (w + t) / 2;
      return { d: `M${x1} 0L${x2} 0L${x2} ${h - hd}L${w} ${h - hd}L${w / 2} ${h}L0 ${h - hd}L${x1} ${h - hd}Z` };
    }
    case 'leftRightArrow': {
      const t = h * a(0, 50000), hd = ss * a(1, 50000), y1 = (h - t) / 2, y2 = (h + t) / 2;
      return { d: `M0 ${h / 2}L${hd} 0L${hd} ${y1}L${w - hd} ${y1}L${w - hd} 0L${w} ${h / 2}L${w - hd} ${h}L${w - hd} ${y2}L${hd} ${y2}L${hd} ${h}Z` };
    }
    case 'upDownArrow': {
      const t = w * a(0, 50000), hd = ss * a(1, 50000), x1 = (w - t) / 2, x2 = (w + t) / 2;
      return { d: `M${w / 2} 0L${w} ${hd}L${x2} ${hd}L${x2} ${h - hd}L${w} ${h - hd}L${w / 2} ${h}L0 ${h - hd}L${x1} ${h - hd}L${x1} ${hd}L0 ${hd}Z` };
    }
    case 'plus': case 'flowChartOr': { const x = prst === 'plus' ? ss * a(0, 25000) : 0; if (!x) return { d: ELLIPSE(w, h) }; return { d: `M${x} 0L${w - x} 0L${w - x} ${x}L${w} ${x}L${w} ${h - x}L${w - x} ${h - x}L${w - x} ${h}L${x} ${h}L${x} ${h - x}L0 ${h - x}L0 ${x}L${x} ${x}Z` }; }
    case 'star4': return { d: polygon(4, w, h, 0.25) };
    case 'star5': return { d: polygon(5, w, h, 0.382) };
    case 'star6': return { d: polygon(6, w, h, 0.5) };
    case 'star7': return { d: polygon(7, w, h, 0.5) };
    case 'star8': return { d: polygon(8, w, h, 0.6) };
    case 'star10': return { d: polygon(10, w, h, 0.65) };
    case 'star12': return { d: polygon(12, w, h, 0.7) };
    case 'star16': case 'star24': case 'star32': case 'irregularSeal1': case 'irregularSeal2': return { d: polygon(16, w, h, 0.8) };
    case 'snip1Rect': { const d = ss * a(0, 16667); return { d: `M0 0L${w - d} 0L${w} ${d}L${w} ${h}L0 ${h}Z` }; }
    case 'snip2SameRect': { const d = ss * a(0, 16667); return { d: `M${d} 0L${w - d} 0L${w} ${d}L${w} ${h}L0 ${h}L0 ${d}Z` }; }
    case 'snip2DiagRect': { const d = ss * a(1, 16667); return { d: `M0 0L${w - d} 0L${w} ${d}L${w} ${h}L${d} ${h}L0 ${h - d}Z` }; }
    case 'plaque': { const d = ss * a(0, 16667); return { d: `M${d} 0L${w - d} 0A${d} ${d} 0 0 0 ${w} ${d}L${w} ${h - d}A${d} ${d} 0 0 0 ${w - d} ${h}L${d} ${h}A${d} ${d} 0 0 0 0 ${h - d}L0 ${d}A${d} ${d} 0 0 0 ${d} 0Z` }; }
    case 'frame': { const d = ss * a(0, 12500); return { d: `M0 0H${w}V${h}H0Z M${d} ${d}V${h - d}H${w - d}V${d}Z`, evenodd: true }; }
    case 'flowChartDocument': return { d: `M0 0H${w}V${h * 0.83}C${w * 0.7} ${h * 0.6} ${w * 0.3} ${h * 1.1} 0 ${h * 0.9}Z` };
    case 'wave': case 'doubleWave': return { d: `M0 ${h * 0.1}C${w * 0.3} ${-h * 0.1} ${w * 0.7} ${h * 0.3} ${w} ${h * 0.1}V${h * 0.9}C${w * 0.7} ${h * 1.1} ${w * 0.3} ${h * 0.7} 0 ${h * 0.9}Z` };
    case 'can': case 'flowChartMagneticDisk': { const e = Math.min(h * a(0, 25000), h / 2) / 2; return { d: `M0 ${e}A${w / 2} ${e} 0 0 1 ${w} ${e}V${h - e}A${w / 2} ${e} 0 0 1 0 ${h - e}Z M0 ${e}A${w / 2} ${e} 0 0 0 ${w} ${e}` }; }
    case 'cube': { const d = ss * a(0, 25000); return { d: `M0 ${d}L${d} 0H${w}V${h - d}L${w - d} ${h}H0Z M0 ${d}H${w - d}V${h} M${w - d} ${d}L${w} 0` }; }
    case 'heart': return { d: `M${w / 2} ${h * 0.25}C${w * 0.5} 0 0 0 0 ${h * 0.3}C0 ${h * 0.6} ${w * 0.4} ${h * 0.8} ${w / 2} ${h}C${w * 0.6} ${h * 0.8} ${w} ${h * 0.6} ${w} ${h * 0.3}C${w} 0 ${w * 0.5} 0 ${w / 2} ${h * 0.25}Z` };
    case 'teardrop': return { d: `M0 ${h / 2}A${w / 2} ${h / 2} 0 1 1 ${w / 2} ${h}A${w / 2} ${h / 2} 0 0 1 0 ${h / 2}Z` };
    case 'line': case 'straightConnector1': return { d: `M0 0L${w} ${h}`, open: true };
    case 'bentConnector2': return { d: `M0 0L${w} 0L${w} ${h}`, open: true };
    case 'bentConnector3': { const x = w * a(0, 50000); return { d: `M0 0L${x} 0L${x} ${h}L${w} ${h}`, open: true }; }
    case 'bentConnector4': case 'bentConnector5': { const x = w * a(0, 50000), y = h * a(1, 50000); return { d: `M0 0L${x} 0L${x} ${y}L${w} ${y}L${w} ${h}`, open: true }; }
    case 'curvedConnector2': case 'curvedConnector3': case 'curvedConnector4': case 'curvedConnector5': return { d: `M0 0C${w / 2} 0 ${w / 2} ${h} ${w} ${h}`, open: true };
    case 'arc': {
      const st = (adj[0] != null ? adj[0] : 16200000) / 60000 * Math.PI / 180, en = (adj[1] != null ? adj[1] : 0) / 60000 * Math.PI / 180;
      let sw = en - st; if (sw <= 0) sw += Math.PI * 2;
      const p = (t) => `${r2(w / 2 + Math.cos(t) * w / 2)} ${r2(h / 2 + Math.sin(t) * h / 2)}`;
      return { d: `M${p(st)}A${w / 2} ${h / 2} 0 ${sw > Math.PI ? 1 : 0} 1 ${p(st + sw)}`, open: true };
    }
    case 'leftBracket': return { d: `M${w} 0Q0 0 0 ${h * 0.1}V${h * 0.9}Q0 ${h} ${w} ${h}`, open: true };
    case 'rightBracket': return { d: `M0 0Q${w} 0 ${w} ${h * 0.1}V${h * 0.9}Q${w} ${h} 0 ${h}`, open: true };
    case 'leftBrace': return { d: `M${w} 0Q${w / 2} 0 ${w / 2} ${h * 0.1}V${h * 0.4}Q${w / 2} ${h / 2} 0 ${h / 2}Q${w / 2} ${h / 2} ${w / 2} ${h * 0.6}V${h * 0.9}Q${w / 2} ${h} ${w} ${h}`, open: true };
    case 'rightBrace': return { d: `M0 0Q${w / 2} 0 ${w / 2} ${h * 0.1}V${h * 0.4}Q${w / 2} ${h / 2} ${w} ${h / 2}Q${w / 2} ${h / 2} ${w / 2} ${h * 0.6}V${h * 0.9}Q${w / 2} ${h} 0 ${h}`, open: true };
    case 'bracketPair': return { d: `M${w * 0.15} 0Q0 0 0 ${h * 0.15}V${h * 0.85}Q0 ${h} ${w * 0.15} ${h}M${w * 0.85} 0Q${w} 0 ${w} ${h * 0.15}V${h * 0.85}Q${w} ${h} ${w * 0.85} ${h}`, open: true };
    case 'wedgeRoundRectCallout': return { d: roundRectPath(w, h, ss * 0.16667) };
    default: return { d: `M0 0H${w}V${h}H0Z` };
  }
}
// Свободная фигура (custGeom): координаты путей заданы в собственной системе w×h
function customPath(geom, W, H) {
  let d = '', open = false, nofill = false;
  for (const path of kids(kid(geom, 'pathLst'), 'path')) {
    const pw = num(path, 'w', 0) || W / EMU_PX, ph = num(path, 'h', 0) || H / EMU_PX;
    const sx = W / pw, sy = H / ph;
    if (attr(path, 'fill', '') === 'none') nofill = true;
    let cx = 0, cy = 0, closed = false;
    const pt = (p) => [num(p, 'x', 0) * sx, num(p, 'y', 0) * sy];
    for (const c of path.children) {
      const pts = kids(c, 'pt').map(pt);
      switch (c.localName) {
        case 'moveTo': [cx, cy] = pts[0]; d += `M${r2(cx)} ${r2(cy)}`; break;
        case 'lnTo': [cx, cy] = pts[0]; d += `L${r2(cx)} ${r2(cy)}`; break;
        case 'cubicBezTo': d += 'C' + pts.map((p) => `${r2(p[0])} ${r2(p[1])}`).join(' '); [cx, cy] = pts[2]; break;
        case 'quadBezTo': d += 'Q' + pts.map((p) => `${r2(p[0])} ${r2(p[1])}`).join(' '); [cx, cy] = pts[1]; break;
        case 'arcTo': {
          const wR = num(c, 'wR', 0) * sx, hR = num(c, 'hR', 0) * sy;
          const st = num(c, 'stAng', 0) / 60000 * Math.PI / 180, sw = num(c, 'swAng', 0) / 60000 * Math.PI / 180;
          const ox = cx - wR * Math.cos(st), oy = cy - hR * Math.sin(st);
          const ex = ox + wR * Math.cos(st + sw), ey = oy + hR * Math.sin(st + sw);
          d += `A${r2(wR)} ${r2(hR)} 0 ${Math.abs(sw) > Math.PI ? 1 : 0} ${sw > 0 ? 1 : 0} ${r2(ex)} ${r2(ey)}`;
          cx = ex; cy = ey;
          break;
        }
        case 'close': d += 'Z'; closed = true; break;
      }
    }
    if (!closed) open = true;
  }
  return { d: d || `M0 0H${W}V${H}H0Z`, open: open && nofill, nofill };
}

/* ---------- Text ---------- */
const WINGDINGS = { '§': '■', 'Ø': '➢', 'ü': '✓', 'q': '❑', 'v': '❖', 'l': '●', 'n': '■', 'Ÿ': '•', 'è': '➔', 'à': '➔', 'Þ': '➔', 'ð': '➔', 'o': '□', 'p': '◆', 'u': '◆' };
const SERIF = /times|georgia|cambria|garamond|book|serif|roman|palatino|constantia|baskerville/i;
const MONO = /consolas|courier|mono|lucida console/i;
function fontCss(name) {
  if (!name) return 'sans-serif';
  const gen = MONO.test(name) ? 'monospace' : SERIF.test(name) && !/sans/i.test(name) ? 'serif' : 'sans-serif';
  return `"${name.replace(/"/g, '')}",${gen}`;
}
function themeFont(name, ctx) {
  if (!name) return null;
  if (name.startsWith('+mj')) return ctx.theme.major;
  if (name.startsWith('+mn')) return ctx.theme.minor;
  return name;
}
function spacing(el, fontPt) {
  if (!el) return null;
  const pct = kid(el, 'spcPct'), pts = kid(el, 'spcPts');
  if (pts) return { px: num(pts, 'val', 0) / 100 * PT_PX };
  if (pct) return { pct: num(pct, 'val', 100000) / 100000 };
  return null;
}
function resolvePara(chain) {
  const p = { algn: 'l', marL: 0, indent: 0, lnSpc: null, spcBef: null, spcAft: null, bu: { kind: 'none' }, buFont: null, buClr: null, buSz: 1, defs: [], rtl: false };
  for (const e of chain) {
    if (!e) continue;
    p.algn = attr(e, 'algn', p.algn);
    p.marL = num(e, 'marL', p.marL);
    p.indent = num(e, 'indent', p.indent);
    p.rtl = bool(e, 'rtl', p.rtl);
    p.lnSpc = spacing(kid(e, 'lnSpc')) || p.lnSpc;
    p.spcBef = spacing(kid(e, 'spcBef')) || p.spcBef;
    p.spcAft = spacing(kid(e, 'spcAft')) || p.spcAft;
    for (const c of e.children) {
      switch (c.localName) {
        case 'buNone': p.bu = { kind: 'none' }; break;
        case 'buChar': p.bu = { kind: 'char', char: attr(c, 'char', '•') }; break;
        case 'buAutoNum': p.bu = { kind: 'num', type: attr(c, 'type', 'arabicPeriod'), start: num(c, 'startAt', 1) }; break;
        case 'buBlip': p.bu = { kind: 'char', char: '•' }; break;
        case 'buFont': p.buFont = attr(c, 'typeface', null); break;
        case 'buFontTx': p.buFont = null; break;
        case 'buClr': p.buClr = c; break;
        case 'buClrTx': p.buClr = null; break;
        case 'buSzPct': p.buSz = num(c, 'val', 100000) / 100000; break;
        case 'defRPr': p.defs.push(c); break;
      }
    }
  }
  return p;
}
function runProps(list, ctx) {
  const r = { sz: 1800, b: false, i: false, u: false, strike: false, cap: 'none', baseline: 0, spc: 0, color: null, font: null, hl: null };
  for (const e of list) {
    if (!e) continue;
    r.sz = num(e, 'sz', r.sz);
    r.b = bool(e, 'b', r.b);
    r.i = bool(e, 'i', r.i);
    if (e.hasAttribute('u')) r.u = attr(e, 'u') !== 'none';
    if (e.hasAttribute('strike')) r.strike = attr(e, 'strike') !== 'noStrike';
    r.cap = attr(e, 'cap', r.cap);
    r.baseline = num(e, 'baseline', r.baseline);
    r.spc = num(e, 'spc', r.spc);
    const f = findFill(e);
    if (f) { const pf = parseFill(f, ctx); r.color = pf && pf.type === 'none' ? { rgb: [0, 0, 0], a: 0 } : solidOf(pf) || r.color; }
    const lat = kid(e, 'latin');
    if (lat && attr(lat, 'typeface', '')) r.font = attr(lat, 'typeface');
    else if (!r.font) { const cs = kid(e, 'cs') || kid(e, 'ea'); if (cs && attr(cs, 'typeface', '') && !attr(cs, 'typeface').startsWith('+')) r.font = attr(cs, 'typeface'); }
    const hl = kid(e, 'highlight');
    if (hl) r.hl = colorOf(hl, ctx);
  }
  return r;
}
function autoNumber(type, n) {
  const roman = (x) => { const v = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']]; let s = ''; for (const [k, r] of v) while (x >= k) { s += r; x -= k; } return s; };
  const alpha = (x) => { let s = ''; while (x > 0) { s = String.fromCharCode(97 + ((x - 1) % 26)) + s; x = Math.floor((x - 1) / 26); } return s; };
  let body = String(n);
  if (/^romanUc/.test(type)) body = roman(n);
  else if (/^romanLc/.test(type)) body = roman(n).toLowerCase();
  else if (/^alphaUc/.test(type)) body = alpha(n).toUpperCase();
  else if (/^alphaLc/.test(type)) body = alpha(n);
  if (/ParenBoth$/.test(type)) return `(${body})`;
  if (/ParenR$/.test(type)) return body + ')';
  if (/Plain$/.test(type)) return body;
  return body + '.';
}
/**
 * HTML блока текста. o.lst — стили списка от общих к частным, o.bodyPrs — свойства
 * блока (bodyPr) от общих к частным, o.fontColor — цвет из стиля фигуры.
 */
function textHtml(txBody, ctx, o) {
  const paras = kids(txBody, 'p');
  const hasText = paras.some((p) => desc(p, 't').some((t) => t.textContent.trim()));
  if (!hasText) return null;
  const bodyPrs = (o.bodyPrs || []).concat([kid(txBody, 'bodyPr')]).filter(Boolean);
  const bp = (name, def) => { let v = def; for (const b of bodyPrs) if (b.hasAttribute(name)) v = b.getAttribute(name); return v; };
  let fontScale = 1, lnRed = 0;
  for (const b of bodyPrs) {
    const n = kid(b, 'normAutofit');
    if (n) { fontScale = num(n, 'fontScale', 100000) / 100000; lnRed = num(n, 'lnSpcReduction', 0) / 100000; }
    else if (kid(b, 'noAutofit') || kid(b, 'spAutoFit')) { fontScale = 1; lnRed = 0; }
  }
  const lst = (o.lst || []).concat([kid(txBody, 'lstStyle')]).filter(Boolean);
  const tx1 = colorOf(null, ctx) || { rgb: hexRgb(schemeHex('tx1', ctx)), a: 1 };
  const counters = [];
  let html = '';
  let firstSize = null;
  for (const p of paras) {
    const pPr = kid(p, 'pPr');
    const lvl = Math.min(9, num(pPr, 'lvl', 0) + 1);
    const pp = resolvePara(lst.map((ls) => kid(ls, `lvl${lvl}pPr`)).concat([pPr]));
    const items = [];
    for (const c of p.children) {
      if (c.localName === 'r' || c.localName === 'fld') {
        let t = (kid(c, 't') || { textContent: '' }).textContent;
        if (c.localName === 'fld' && attr(c, 'type', '') === 'slidenum') t = String(ctx.slideNum || t);
        items.push({ t, rPr: kid(c, 'rPr') });
      } else if (c.localName === 'br') items.push({ br: true, rPr: kid(c, 'rPr') });
    }
    const endRPr = kid(p, 'endParaRPr');
    const textless = !items.some((it) => it.t);
    const baseRun = runProps(pp.defs, ctx);
    let runs = '';
    let maxSz = 0;
    for (const it of items) {
      const rp = runProps(pp.defs.concat(o.runDefaults ? [o.runDefaults] : [], [it.rPr]), ctx);
      const own = it.rPr ? runProps([it.rPr], ctx).color : null;
      const pt = rp.sz / 100 * fontScale;
      maxSz = Math.max(maxSz, pt);
      if (it.br) { runs += '<br>'; continue; }
      let style = `font-size:${r2(pt * PT_PX)}px;font-family:${esc(fontCss(themeFont(rp.font, ctx) || o.font || ctx.theme.minor))};color:${css(own || o.fontColor || rp.color || tx1)};`;
      if (rp.b) style += 'font-weight:700;';
      if (rp.i) style += 'font-style:italic;';
      if (rp.u || rp.strike) style += `text-decoration:${rp.u ? 'underline ' : ''}${rp.strike ? 'line-through' : ''};`;
      if (rp.cap === 'all') style += 'text-transform:uppercase;';
      else if (rp.cap === 'small') style += 'font-variant:small-caps;';
      if (rp.spc) style += `letter-spacing:${r2(rp.spc / 100 * PT_PX)}px;`;
      if (rp.baseline) style += `vertical-align:${rp.baseline > 0 ? 'super' : 'sub'};font-size:${r2(pt * PT_PX * 0.7)}px;`;
      if (rp.hl) style += `background:${css(rp.hl)};`;
      runs += `<span style="${style}">${esc(it.t)}</span>`;
    }
    if (textless) {
      const rp = runProps(pp.defs.concat([endRPr]), ctx);
      maxSz = rp.sz / 100 * fontScale;
      runs = `<span style="font-size:${r2(maxSz * PT_PX)}px">&#8203;</span>`;
    }
    if (firstSize == null) firstSize = maxSz;
    // Нумерация продолжается, пока идут абзацы того же уровня с номерами
    if (pp.bu.kind === 'num' && !textless) {
      counters[lvl] = counters[lvl] == null ? pp.bu.start : counters[lvl] + 1;
      for (let k = lvl + 1; k < counters.length; k++) counters[k] = null;
    } else if (!textless) counters[lvl] = null;
    let bullet = '';
    if (pp.bu.kind !== 'none' && !textless) {
      const bcol = pp.buClr ? colorOf(pp.buClr, ctx) : (o.fontColor || baseRun.color || tx1);
      const firstRun = items.find((it) => it.t);
      const bsz = runProps(pp.defs.concat([firstRun && firstRun.rPr]), ctx).sz / 100 * fontScale * pp.buSz;
      let ch = pp.bu.kind === 'num' ? autoNumber(pp.bu.type, counters[lvl]) : pp.bu.char;
      let bfont = pp.buFont;
      if (pp.bu.kind === 'char' && bfont && /wingdings|symbol/i.test(bfont)) { ch = WINGDINGS[ch] || (ch === '·' || ch === '' ? '•' : ch); bfont = null; }
      const width = Math.max(0, -pp.indent * EMU_PX);
      bullet = `<span class="sv-bu" style="min-width:${r2(width)}px;font-size:${r2(bsz * PT_PX)}px;color:${css(bcol)};${bfont ? `font-family:${esc(fontCss(bfont))};` : ''}">${esc(ch)}</span>`;
    }
    const lh = pp.lnSpc ? (pp.lnSpc.px ? `${r2(pp.lnSpc.px)}px` : r2(pp.lnSpc.pct * (1 - lnRed) * 1.2)) : r2(1.2 * (1 - lnRed));
    const sp = (s) => (!s ? 0 : s.px != null ? s.px : s.pct * maxSz * PT_PX * 1.2);
    const align = { ctr: 'center', r: 'right', just: 'justify', dist: 'justify', l: 'left' }[pp.algn] || 'left';
    html += `<div class="sv-p" style="text-align:${align};padding-left:${r2(pp.marL * EMU_PX)}px;text-indent:${r2(pp.indent * EMU_PX)}px;line-height:${lh};margin-top:${r2(sp(pp.spcBef))}px;margin-bottom:${r2(sp(pp.spcAft))}px;${pp.rtl ? 'direction:rtl;' : ''}">${bullet}${runs}</div>`;
  }
  const ins = (n, d) => num({ getAttribute: (k) => bp(k, null), hasAttribute: () => true }, n, d) * EMU_PX;
  const anchor = bp('anchor', 't');
  const vert = bp('vert', 'horz');
  return {
    html,
    pad: [ins('tIns', 45720), ins('rIns', 91440), ins('bIns', 45720), ins('lIns', 91440)],
    justify: { ctr: 'center', b: 'flex-end', just: 'center', dist: 'center' }[anchor] || 'flex-start',
    nowrap: bp('wrap', 'square') === 'none',
    vert: vert === 'vert' || vert === 'eaVert' ? 90 : vert === 'vert270' ? 270 : 0,
    rot: num({ getAttribute: (k) => bp(k, null), hasAttribute: () => true }, 'rot', 0) / 60000,
  };
}
function textEl(t, box) {
  const div = document.createElement('div');
  div.className = 'sv-tx';
  let w = box.w, h = box.h, extra = '';
  if (t.vert) { w = box.h; h = box.w; extra = `transform:translate(${r2((box.w - w) / 2)}px,${r2((box.h - h) / 2)}px) rotate(${t.vert}deg);`; }
  else if (t.rot) extra = `transform:rotate(${t.rot}deg);`;
  div.style.cssText = `left:${r2(box.x)}px;top:${r2(box.y)}px;width:${r2(w)}px;height:${r2(h)}px;padding:${t.pad.map((p) => r2(p) + 'px').join(' ')};justify-content:${t.justify};${t.nowrap ? 'white-space:pre;' : ''}${extra}`;
  div.innerHTML = t.html;
  return div;
}

/* ---------- Placeholders and inheritance ---------- */
function phOf(el) {
  const nv = kid(el, 'nvSpPr') || kid(el, 'nvPicPr') || kid(el, 'nvGraphicFramePr') || kid(el, 'nvCxnSpPr') || kid(el, 'nvGrpSpPr');
  const ph = q(nv, 'nvPr/ph');
  return ph ? { type: attr(ph, 'type', null), idx: attr(ph, 'idx', null) } : null;
}
const normType = (t) => ({ ctrTitle: 'title', subTitle: 'body', obj: 'body', null: 'body', undefined: 'body' }[t] || t);
function phIndex(root) {
  const list = [];
  const walk = (tree) => {
    for (const el of kids(tree)) {
      if (el.localName === 'grpSp') { walk(el); continue; }
      const ph = phOf(el);
      if (ph) list.push({ el, type: ph.type, idx: ph.idx });
    }
  };
  walk(q(root, 'cSld/spTree'));
  return list;
}
function findPh(list, ph, byIdx) {
  if (!list || !ph) return null;
  if (byIdx && ph.idx != null) {
    const m = list.find((p) => p.idx === ph.idx);
    if (m) return m;
  }
  const t = ph.type;
  return list.find((p) => p.type === t) || list.find((p) => normType(p.type) === normType(t)) || null;
}
function textStyleFor(type, ctx) {
  const tx = kid(ctx.masterRoot, 'txStyles');
  if (type === 'title' || type === 'ctrTitle') return kid(tx, 'titleStyle');
  if (['body', 'subTitle', 'obj', null, undefined].includes(type)) return kid(tx, 'bodyStyle');
  return kid(tx, 'otherStyle');
}

/* ---------- Transform helpers ---------- */
function readXfrm(x) {
  if (!x) return null;
  const off = kid(x, 'off'), ext = kid(x, 'ext');
  if (!off || !ext) return null;
  return { x: num(off, 'x', 0), y: num(off, 'y', 0), w: num(ext, 'cx', 0), h: num(ext, 'cy', 0), rot: num(x, 'rot', 0) / 60000, flipH: bool(x, 'flipH', false), flipV: bool(x, 'flipV', false) };
}
const identity = (x, y, w, h) => ({ x, y, w, h });
function pxBox(b) { return { x: b.x * EMU_PX, y: b.y * EMU_PX, w: b.w * EMU_PX, h: b.h * EMU_PX }; }

/* ---------- Renderer ---------- */
let uidCounter = 0;
const uid = (p) => `${p}${++uidCounter}`;

async function svgFor(geomD, w, h, fill, line, pkg, flip) {
  const defs = [];
  let fillAttr = 'none';
  if (fill && !geomD.open && !geomD.nofill) {
    if (fill.type === 'solid') fillAttr = css(fill.color);
    else if (fill.type === 'grad') {
      const id = uid('g');
      const stops = fill.stops.map((s) => `<stop offset="${r2(s.pos * 100)}%" stop-color="${css({ rgb: s.color.rgb, a: 1 })}" stop-opacity="${s.color.a}"/>`).join('');
      if (fill.radial) defs.push(`<radialGradient id="${id}">${stops}</radialGradient>`);
      else {
        const ang = fill.angle * Math.PI / 180, dx = Math.cos(ang) / 2, dy = Math.sin(ang) / 2;
        defs.push(`<linearGradient id="${id}" x1="${r2(0.5 - dx)}" y1="${r2(0.5 - dy)}" x2="${r2(0.5 + dx)}" y2="${r2(0.5 + dy)}">${stops}</linearGradient>`);
      }
      fillAttr = `url(#${id})`;
    } else if (fill.type === 'img') {
      const u = await pkg.url(fill.path);
      if (u) {
        const id = uid('p');
        defs.push(`<pattern id="${id}" patternUnits="userSpaceOnUse" width="${w}" height="${h}"><image href="${u}" width="${w}" height="${h}" preserveAspectRatio="none"/></pattern>`);
        fillAttr = `url(#${id})`;
      }
    }
  }
  let stroke = '';
  if (line) {
    const lw = line.w;
    const dashes = { dash: [4, 3], sysDash: [3, 1], dot: [1, 1], sysDot: [1, 1], dashDot: [4, 3, 1, 3], lgDash: [8, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1] }[line.dash];
    stroke = ` stroke="${css(line.color)}" stroke-width="${r2(lw)}"${dashes ? ` stroke-dasharray="${dashes.map((d) => r2(d * lw)).join(' ')}"` : ''} stroke-linejoin="round"`;
    const marker = (type, end) => {
      if (!type || type === 'none') return '';
      const id = uid('m');
      const shape = type === 'oval' ? '<circle cx="5" cy="5" r="4"/>' : type === 'diamond' ? '<path d="M0 5L5 0L10 5L5 10Z"/>' : '<path d="M0 0L10 5L0 10Z"/>';
      defs.push(`<marker id="${id}" viewBox="0 0 10 10" refX="${type === 'oval' || type === 'diamond' ? 5 : 8}" refY="5" markerWidth="${3 + 4 / Math.max(1, lw / 2)}" markerHeight="${3 + 4 / Math.max(1, lw / 2)}" orient="${end ? 'auto' : 'auto-start-reverse'}" fill="${css(line.color)}">${shape}</marker>`);
      return ` marker-${end ? 'end' : 'start'}="url(#${id})"`;
    };
    stroke += marker(line.head, false) + marker(line.tail, true);
  }
  if (fillAttr === 'none' && !stroke) return '';
  const tr = flip.h || flip.v ? ` transform="translate(${flip.h ? w : 0} ${flip.v ? h : 0}) scale(${flip.h ? -1 : 1} ${flip.v ? -1 : 1})"` : '';
  return `<svg class="sv-g" width="${r2(w)}" height="${r2(h)}" viewBox="0 0 ${r2(w)} ${r2(h)}">${defs.length ? `<defs>${defs.join('')}</defs>` : ''}<path d="${geomD.d}" fill="${fillAttr}"${geomD.evenodd ? ' fill-rule="evenodd"' : ''}${stroke}${tr}/></svg>`;
}

class SlideRenderer {
  constructor(pkg, pres) { this.pkg = pkg; this.pres = pres; }

  async render(slidePath, slideNum, canvas) {
    const pkg = this.pkg;
    const slideRoot = await pkg.xml(slidePath);
    if (!slideRoot) throw new Error('нет слайда ' + slidePath);
    const slideRels = await pkg.rels(slidePath);
    const layoutPath = (relOfType(slideRels, '/slideLayout') || {}).target;
    const layoutRoot = await pkg.xml(layoutPath);
    const layoutRels = await pkg.rels(layoutPath);
    const masterPath = (relOfType(layoutRels, '/slideMaster') || {}).target;
    const masterRoot = await pkg.xml(masterPath);
    const masterRels = await pkg.rels(masterPath);
    const themePath = (relOfType(masterRels, '/theme') || {}).target;
    const theme = await this.pres.theme(themePath);

    const clrMap = {};
    const mapEl = kid(masterRoot, 'clrMap');
    for (const a of (mapEl ? mapEl.attributes : [])) clrMap[a.name] = a.value;
    for (const root of [layoutRoot, slideRoot]) {
      const ovr = q(root, 'clrMapOvr/overrideClrMapping');
      if (ovr) for (const a of ovr.attributes) clrMap[a.name] = a.value;
    }
    const base = { pkg, theme, clrMap, slideNum, masterRoot, pres: this.pres, groupFill: null };
    const mCtx = Object.assign({}, base, { rels: masterRels, part: masterPath });
    const lCtx = Object.assign({}, base, { rels: layoutRels, part: layoutPath });
    const sCtx = Object.assign({}, base, { rels: slideRels, part: slidePath });
    const masterPh = phIndex(masterRoot), layoutPh = phIndex(layoutRoot);
    sCtx.inherit = (ph) => {
      const l = findPh(layoutPh, ph, true);
      const m = findPh(masterPh, l ? { type: l.type, idx: l.idx } : ph, false);
      return { layout: l && l.el, master: m && m.el, type: ph.type != null ? ph.type : l ? l.type : m ? m.type : null };
    };
    lCtx.inherit = (ph) => { const m = findPh(masterPh, ph, false); return { layout: null, master: m && m.el, type: ph.type }; };
    mCtx.inherit = (ph) => ({ layout: null, master: null, type: ph.type });

    // Фон: слайд → макет → образец
    let bgCss = '#fff';
    for (const [root, ctx] of [[slideRoot, sCtx], [layoutRoot, lCtx], [masterRoot, mCtx]]) {
      const bg = q(root, 'cSld/bg');
      if (!bg) continue;
      const bgPr = kid(bg, 'bgPr'), ref = kid(bg, 'bgRef');
      let fill = null;
      if (bgPr) fill = parseFill(findFill(bgPr), ctx);
      else if (ref) {
        const idx = num(ref, 'idx', 0), ph = colorOf(ref, ctx);
        const list = idx >= 1001 ? theme.bgFills : theme.fills;
        fill = parseFill(list[(idx >= 1001 ? idx - 1001 : idx - 1)], ctx, ph) || (ph ? { type: 'solid', color: ph } : null);
      }
      if (fill) { bgCss = await fillCss(fill, pkg); break; }
    }
    canvas.style.background = bgCss;

    const showMaster = attr(slideRoot, 'showMasterSp', '1') !== '0' && attr(layoutRoot, 'showMasterSp', '1') !== '0';
    const showLayout = attr(slideRoot, 'showMasterSp', '1') !== '0';
    if (showMaster && masterRoot) await this.tree(q(masterRoot, 'cSld/spTree'), mCtx, canvas, identity, true);
    if (showLayout && layoutRoot) await this.tree(q(layoutRoot, 'cSld/spTree'), lCtx, canvas, identity, true);
    await this.tree(q(slideRoot, 'cSld/spTree'), sCtx, canvas, identity, false);
  }

  async tree(tree, ctx, parent, xf, skipPh) {
    for (const el of kids(tree)) {
      try {
        switch (el.localName) {
          case 'AlternateContent': {
            const alt = kid(el, 'Fallback') || kid(el, 'Choice');
            if (alt) await this.tree(alt, ctx, parent, xf, skipPh);
            break;
          }
          case 'sp': case 'cxnSp': await this.shape(el, ctx, parent, xf, skipPh); break;
          case 'pic': await this.picture(el, ctx, parent, xf, skipPh); break;
          case 'grpSp': await this.group(el, ctx, parent, xf, skipPh); break;
          case 'graphicFrame': await this.frame(el, ctx, parent, xf, skipPh); break;
        }
      } catch (e) {
        console.warn('Не удалось нарисовать элемент слайда', el.localName, e);
      }
    }
  }

  placeBox(el, inh, xf) {
    const own = readXfrm(kid(kid(el, 'spPr') || kid(el, 'grpSpPr'), 'xfrm') || kid(el, 'xfrm'));
    const x = own || readXfrm(q(inh.layout, 'spPr/xfrm')) || readXfrm(q(inh.master, 'spPr/xfrm')) || readXfrm(kid(inh.layout, 'xfrm')) || readXfrm(kid(inh.master, 'xfrm'));
    if (!x) return null;
    const b = xf(x.x, x.y, x.w, x.h);
    return Object.assign(pxBox(b), { rot: x.rot, flipH: x.flipH, flipV: x.flipV });
  }

  wrapper(box, parent) {
    const div = document.createElement('div');
    div.className = 'sv-sp';
    div.style.cssText = `left:${r2(box.x)}px;top:${r2(box.y)}px;width:${r2(box.w)}px;height:${r2(box.h)}px;${box.rot ? `transform:rotate(${r2(box.rot)}deg);` : ''}`;
    parent.appendChild(div);
    return div;
  }

  async shape(sp, ctx, parent, xf, skipPh) {
    const ph = phOf(sp);
    if (ph && skipPh) return;
    const inh = ph && ctx.inherit ? ctx.inherit(ph) : { layout: null, master: null, type: null };
    const box = this.placeBox(sp, inh, xf);
    if (!box) return;
    const spPr = kid(sp, 'spPr');
    const spPrs = [spPr, q(inh.layout, 'spPr'), q(inh.master, 'spPr')].filter(Boolean);
    const style = kid(sp, 'style');
    const fill = shapeFill(spPrs, style, ctx);
    const line = shapeLine(spPrs, style, ctx);
    let geomEl = null;
    for (const s of spPrs) { geomEl = kid(s, 'custGeom') || kid(s, 'prstGeom'); if (geomEl) break; }
    const geom = geomEl && geomEl.localName === 'custGeom' ? customPath(geomEl, box.w, box.h) : presetPath(attr(geomEl, 'prst', 'rect'), box.w, box.h, adjValues(geomEl));
    const div = this.wrapper(box, parent);
    const svg = await svgFor(geom, box.w, box.h, fill && fill.type !== 'none' ? fill : null, line, ctx.pkg, { h: box.flipH, v: box.flipV });
    if (svg) div.insertAdjacentHTML('beforeend', svg);

    const txBody = kid(sp, 'txBody');
    if (!txBody) return;
    const type = inh.type;
    const lst = ph
      ? [textStyleFor(type, ctx), q(inh.master, 'txBody/lstStyle'), q(inh.layout, 'txBody/lstStyle')]
      : [kid(ctx.pres.root, 'defaultTextStyle')];
    const fontRef = kid(style, 'fontRef');
    const t = textHtml(txBody, ctx, {
      lst,
      bodyPrs: ph ? [q(inh.master, 'txBody/bodyPr'), q(inh.layout, 'txBody/bodyPr')] : [],
      fontColor: colorOf(fontRef, ctx),
      font: fontRef ? themeFont(attr(fontRef, 'idx', 'minor') === 'major' ? '+mj' : '+mn', ctx) : null,
    });
    if (!t) return;
    // В SmartArt у текста бывает своя рамка (txXfrm)
    const tx = readXfrm(kid(sp, 'txXfrm'));
    let tbox = { x: 0, y: 0, w: box.w, h: box.h };
    if (tx) { const b = pxBox(xf(tx.x, tx.y, tx.w, tx.h)); tbox = { x: b.x - box.x, y: b.y - box.y, w: b.w, h: b.h }; }
    div.appendChild(textEl(t, tbox));
  }

  async picture(pic, ctx, parent, xf, skipPh) {
    const ph = phOf(pic);
    if (ph && skipPh) return;
    const inh = ph && ctx.inherit ? ctx.inherit(ph) : { layout: null, master: null };
    const box = this.placeBox(pic, inh, xf);
    if (!box) return;
    const blipFill = kid(pic, 'blipFill');
    const blip = kid(blipFill, 'blip');
    let rel = ctx.rels[rid(blip, 'embed')];
    // Векторная картинка с PNG-заменой: берём SVG, если он есть
    const svgBlip = desc(blip, 'svgBlip')[0];
    if (svgBlip && ctx.rels[rid(svgBlip, 'embed')]) rel = ctx.rels[rid(svgBlip, 'embed')];
    const div = this.wrapper(box, parent);
    div.classList.add('sv-pic');
    const spPr = kid(pic, 'spPr');
    const prst = attr(kid(spPr, 'prstGeom'), 'prst', 'rect');
    if (prst === 'ellipse') div.style.borderRadius = '50%';
    else if (prst === 'roundRect') div.style.borderRadius = r2(Math.min(box.w, box.h) * 0.16667) + 'px';
    const line = shapeLine([spPr], kid(pic, 'style'), ctx);
    if (line) div.style.boxShadow = `inset 0 0 0 ${r2(line.w)}px ${css(line.color)}`;
    if (box.flipH || box.flipV) div.style.transform = `${box.rot ? `rotate(${r2(box.rot)}deg) ` : ''}scale(${box.flipH ? -1 : 1},${box.flipV ? -1 : 1})`;
    const ext = rel ? rel.target.split('.').pop().toLowerCase() : '';
    if (!rel || rel.external || !BROWSER_IMAGES.has(ext)) {
      div.classList.add('sv-missing');
      div.innerHTML = `<span>${rel && !rel.external ? 'Картинка ' + esc(ext.toUpperCase()) : 'Картинка'}</span>`;
      return;
    }
    const url = await ctx.pkg.url(rel.target);
    const sr = kid(blipFill, 'srcRect');
    const l = num(sr, 'l', 0) / 100000, t = num(sr, 't', 0) / 100000, r = num(sr, 'r', 0) / 100000, b = num(sr, 'b', 0) / 100000;
    const iw = box.w / Math.max(0.01, 1 - l - r), ih = box.h / Math.max(0.01, 1 - t - b);
    const img = document.createElement('img');
    img.alt = '';
    img.decoding = 'async';
    img.src = url;
    img.style.cssText = `width:${r2(iw)}px;height:${r2(ih)}px;left:${r2(-l * iw)}px;top:${r2(-t * ih)}px;`;
    div.insertBefore(img, div.firstChild);
  }

  async group(g, ctx, parent, xf, skipPh) {
    const gp = kid(g, 'grpSpPr');
    const x = kid(gp, 'xfrm');
    const off = kid(x, 'off'), ext = kid(x, 'ext'), chOff = kid(x, 'chOff'), chExt = kid(x, 'chExt');
    let childXf = xf;
    if (off && ext) {
      const ox = num(off, 'x', 0), oy = num(off, 'y', 0), cx = num(ext, 'cx', 0), cy = num(ext, 'cy', 0);
      const chx = num(chOff, 'x', ox), chy = num(chOff, 'y', oy), chcx = num(chExt, 'cx', cx) || 1, chcy = num(chExt, 'cy', cy) || 1;
      const sx = cx / chcx, sy = cy / chcy;
      childXf = (x2, y2, w2, h2) => xf(ox + (x2 - chx) * sx, oy + (y2 - chy) * sy, w2 * sx, h2 * sy);
    }
    const gf = findFill(gp);
    const gctx = Object.assign({}, ctx, { groupFill: gf ? parseFill(gf, ctx) : ctx.groupFill });
    await this.tree(g, gctx, parent, childXf, skipPh);
  }

  async frame(fr, ctx, parent, xf, skipPh) {
    const ph = phOf(fr);
    if (ph && skipPh) return;
    const inh = ph && ctx.inherit ? ctx.inherit(ph) : { layout: null, master: null };
    const box = this.placeBox(fr, inh, xf);
    if (!box) return;
    const gd = q(fr, 'graphic/graphicData');
    const uri = attr(gd, 'uri', '');
    if (uri.endsWith('/table')) {
      const div = this.wrapper(box, parent);
      div.appendChild(await this.table(kid(gd, 'tbl'), ctx, box));
      return;
    }
    if (uri.endsWith('/chart')) {
      const rel = ctx.rels[rid(kid(gd, 'chart'), 'id')];
      const root = rel && await ctx.pkg.xml(rel.target);
      if (!root) return;
      const div = this.wrapper(box, parent);
      const cctx = Object.assign({}, ctx, { rels: await ctx.pkg.rels(rel.target) });
      div.innerHTML = renderChart(root, box.w, box.h, cctx);
      return;
    }
    if (uri.endsWith('/diagram')) {
      await this.smartArt(gd, ctx, parent, xf, box);
      return;
    }
    // OLE-объекты и прочее: показываем их картинку-заменитель
    const pic = desc(gd, 'pic')[0];
    if (pic) {
      const ownX = readXfrm(kid(fr, 'xfrm'));
      const pin = (x2, y2, w2, h2) => (ownX ? xf(ownX.x, ownX.y, ownX.w, ownX.h) : { x: x2, y: y2, w: w2, h: h2 });
      await this.picture(pic, ctx, parent, pin, false);
    }
  }

  async smartArt(gd, ctx, parent, xf, box) {
    const relIds = kid(gd, 'relIds');
    const dm = ctx.rels[rid(relIds, 'dm')];
    let drawingPath = null;
    if (dm) {
      const data = await ctx.pkg.xml(dm.target);
      const ext = desc(data, 'dataModelExt')[0];
      const id = attr(ext, 'relId', null);
      if (id && ctx.rels[id]) drawingPath = ctx.rels[id].target;
      if (!drawingPath) drawingPath = dm.target.replace(/data(\d*)\.xml$/, 'drawing$1.xml');
    }
    const root = await ctx.pkg.xml(drawingPath);
    const tree = kid(root, 'spTree');
    if (!tree) return;
    // Координаты фигур SmartArt отсчитываются от левого верхнего угла рамки
    const ox = box.x / EMU_PX, oy = box.y / EMU_PX;
    const dctx = Object.assign({}, ctx, { rels: await ctx.pkg.rels(drawingPath), inherit: null });
    await this.tree(tree, dctx, parent, (x, y, w, h) => ({ x: ox + x, y: oy + y, w, h }), false);
  }

  async table(tbl, ctx, box) {
    const cols = kids(kid(tbl, 'tblGrid'), 'gridCol').map((c) => num(c, 'w', 0) * EMU_PX);
    const rows = kids(tbl, 'tr');
    const tblPr = kid(tbl, 'tblPr');
    const flags = { firstRow: bool(tblPr, 'firstRow', false), lastRow: bool(tblPr, 'lastRow', false), firstCol: bool(tblPr, 'firstCol', false), lastCol: bool(tblPr, 'lastCol', false), bandRow: bool(tblPr, 'bandRow', false), bandCol: bool(tblPr, 'bandCol', false) };
    const styleId = (kid(tblPr, 'tableStyleId') || { textContent: '' }).textContent.trim();
    const style = await this.pres.tableStyle(styleId, ctx);
    const table = document.createElement('table');
    table.className = 'sv-tbl';
    table.style.width = r2(cols.reduce((a, b) => a + b, 0)) + 'px';
    let colgroup = '<colgroup>' + cols.map((w) => `<col style="width:${r2(w)}px">`).join('') + '</colgroup>';
    let body = '';
    const nr = rows.length, nc = cols.length;
    rows.forEach((tr, ri) => {
      body += `<tr style="height:${r2(num(tr, 'h', 0) * EMU_PX)}px">`;
      kids(tr, 'tc').forEach((tc, ci) => {
        if (bool(tc, 'hMerge', false) || bool(tc, 'vMerge', false)) return;
        const regions = ['wholeTbl'];
        const dataRow = ri - (flags.firstRow ? 1 : 0);
        if (flags.bandRow && dataRow >= 0 && !(flags.lastRow && ri === nr - 1)) regions.push(dataRow % 2 ? 'band2H' : 'band1H');
        if (flags.bandCol) regions.push((ci - (flags.firstCol ? 1 : 0)) % 2 ? 'band2V' : 'band1V');
        if (flags.firstCol && ci === 0) regions.push('firstCol');
        if (flags.lastCol && ci === nc - 1) regions.push('lastCol');
        if (flags.lastRow && ri === nr - 1) regions.push('lastRow');
        if (flags.firstRow && ri === 0) regions.push('firstRow');
        const tcPr = kid(tc, 'tcPr');
        let fill = null, txt = { b: false, color: null };
        const borders = {};
        for (const reg of regions) {
          const s = style[reg];
          if (!s) continue;
          if (s.fill) fill = s.fill;
          if (s.b != null) txt.b = s.b;
          if (s.color) txt.color = s.color;
          for (const side of ['left', 'right', 'top', 'bottom']) if (s.borders[side]) borders[side] = s.borders[side];
          if (s.borders.insideH) { if (ri > 0) borders.top = borders.top || s.borders.insideH; if (ri < nr - 1) borders.bottom = borders.bottom || s.borders.insideH; }
          if (s.borders.insideV) { if (ci > 0) borders.left = borders.left || s.borders.insideV; if (ci < nc - 1) borders.right = borders.right || s.borders.insideV; }
        }
        const own = findFill(tcPr);
        if (own) { const f = parseFill(own, ctx); fill = f && f.type === 'none' ? null : solidOf(f) || fill; }
        for (const [tag, side] of [['lnL', 'left'], ['lnR', 'right'], ['lnT', 'top'], ['lnB', 'bottom']]) {
          const ln = kid(tcPr, tag);
          if (!ln) continue;
          const f = parseFill(findFill(ln), ctx);
          borders[side] = f && f.type === 'none' ? null : { color: solidOf(f) || BLACK, w: Math.max(0.6, num(ln, 'w', 12700) * EMU_PX) };
        }
        const pad = [num(tcPr, 'marT', 45720), num(tcPr, 'marR', 91440), num(tcPr, 'marB', 45720), num(tcPr, 'marL', 91440)].map((v) => r2(v * EMU_PX) + 'px').join(' ');
        const va = { ctr: 'middle', b: 'bottom' }[attr(tcPr, 'anchor', 't')] || 'top';
        let st = `padding:${pad};vertical-align:${va};background:${fill ? css(fill) : 'transparent'};`;
        for (const side of ['left', 'right', 'top', 'bottom']) {
          const bd = borders[side];
          st += `border-${side}:${bd ? `${r2(bd.w)}px solid ${css(bd.color)}` : 'none'};`;
        }
        const runDefaults = document.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:defRPr');
        if (txt.b) runDefaults.setAttribute('b', '1');
        const t = kid(tc, 'txBody') ? textHtml(kid(tc, 'txBody'), ctx, { lst: [kid(ctx.pres.root, 'defaultTextStyle')], fontColor: txt.color, runDefaults }) : null;
        const span = `${num(tc, 'gridSpan', 1) > 1 ? ` colspan="${num(tc, 'gridSpan', 1)}"` : ''}${num(tc, 'rowSpan', 1) > 1 ? ` rowspan="${num(tc, 'rowSpan', 1)}"` : ''}`;
        body += `<td${span} style="${st}">${t ? t.html : ''}</td>`;
      });
      body += '</tr>';
    });
    table.innerHTML = colgroup + body;
    return table;
  }
}

/* ---------- Table styles ---------- */
function parseTableStyle(el, ctx) {
  const out = {};
  for (const reg of kids(el)) {
    const tcTx = kid(reg, 'tcTxStyle'), tcStyle = kid(reg, 'tcStyle');
    const r = { fill: null, b: null, color: null, borders: {} };
    if (tcTx) {
      if (tcTx.hasAttribute('b')) r.b = attr(tcTx, 'b') === 'on';
      r.color = colorOf(tcTx, ctx) || colorOf(kid(tcTx, 'fontRef'), ctx);
    }
    const fillEl = kid(tcStyle, 'fill');
    if (fillEl) {
      const f = findFill(fillEl);
      if (f) { const pf = parseFill(f, ctx); r.fill = pf && pf.type === 'none' ? null : solidOf(pf); }
      else if (kid(fillEl, 'fillRef')) r.fill = colorOf(kid(fillEl, 'fillRef'), ctx);
    }
    for (const b of kids(kid(tcStyle, 'tcBdr'))) {
      const ln = kid(b, 'ln');
      if (ln) {
        const f = parseFill(findFill(ln), ctx);
        r.borders[b.localName] = f && f.type === 'none' ? null : { color: solidOf(f) || BLACK, w: Math.max(0.6, num(ln, 'w', 12700) * EMU_PX) };
      } else if (kid(b, 'lnRef')) r.borders[b.localName] = { color: colorOf(kid(b, 'lnRef'), ctx) || BLACK, w: 1 };
    }
    out[reg.localName] = r;
  }
  return out;
}
// Встроенный стиль PowerPoint «Средний стиль 2 — акцент 1»: им оформлены таблицы по умолчанию
function defaultTableStyle(ctx) {
  const acc = { rgb: hexRgb(schemeHex('accent1', ctx)), a: 1 };
  const white = { rgb: [255, 255, 255], a: 1 };
  const line = { color: white, w: 1.33 };
  const borders = { left: line, right: line, top: line, bottom: line, insideH: line, insideV: line };
  return {
    wholeTbl: { fill: mix(acc, 0.8), b: null, color: { rgb: hexRgb(schemeHex('dk1', ctx)), a: 1 }, borders },
    band1H: { fill: mix(acc, 0.6), b: null, color: null, borders: {} },
    band1V: { fill: mix(acc, 0.6), b: null, color: null, borders: {} },
    firstRow: { fill: acc, b: true, color: white, borders: { bottom: { color: white, w: 4 } } },
    lastRow: { fill: acc, b: true, color: white, borders: { top: { color: white, w: 4 } } },
    firstCol: { fill: acc, b: true, color: white, borders: {} },
    lastCol: { fill: acc, b: true, color: white, borders: {} },
  };
}

/* ---------- Charts ---------- */
function cachePts(cache) {
  if (!cache) return [];
  const n = num(kid(cache, 'ptCount'), 'val', 0);
  const out = new Array(n).fill(null);
  for (const pt of kids(cache, 'pt')) {
    const i = num(pt, 'idx', 0);
    const v = kid(pt, 'v');
    out[i] = v ? v.textContent : null;
  }
  return out;
}
function refValues(el) {
  if (!el) return [];
  const ml = desc(el, 'multiLvlStrCache')[0];
  if (ml) {
    const lvls = kids(ml, 'lvl');
    const n = num(kid(ml, 'ptCount'), 'val', 0);
    const lvl = lvls[0];
    if (lvl) { const out = new Array(n).fill(null); for (const pt of kids(lvl, 'pt')) out[num(pt, 'idx', 0)] = (kid(pt, 'v') || {}).textContent || null; return out; }
  }
  return cachePts(desc(el, 'strCache')[0] || desc(el, 'numCache')[0] || kid(el, 'strLit') || kid(el, 'numLit'));
}
function formatCode(el) {
  const fc = desc(el, 'formatCode')[0];
  return fc ? fc.textContent : 'General';
}
const NF = new Map();
function fmtNum(v, code) {
  if (v == null || isNaN(v)) return '';
  code = code || 'General';
  const plain = code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
  const pct = plain.includes('%');
  const dec = (/\.(0+)/.exec(plain) || [, ''])[1].length;
  const group = /#,##|0,0/.test(plain);
  let x = pct ? v * 100 : v;
  let digits = dec;
  if (code === 'General') digits = Math.abs(x) >= 100 ? 0 : Math.abs(x) >= 10 ? 1 : 2;
  const key = digits + (group || code === 'General' ? 'g' : '');
  if (!NF.has(key)) NF.set(key, new Intl.NumberFormat('ru-RU', { maximumFractionDigits: digits, minimumFractionDigits: code === 'General' ? 0 : digits, useGrouping: group || code === 'General' }));
  let s = NF.get(key).format(x);
  const lit = /"([^"]*)"/.exec(code), cur = /\[\$([^\]-]*)/.exec(code);
  const sym = (lit && lit[1]) || (cur && cur[1]) || (/[$€£₽]/.exec(plain) || [''])[0];
  if (sym) s = code.indexOf(sym) < code.search(/[0#]/) ? sym + s : s + (sym.startsWith(' ') ? '' : ' ') + sym.trim();
  return s + (pct ? '%' : '');
}
// Категории-даты хранятся числами Excel (дни с 30.12.1899)
function fmtCategory(c, code) {
  if (!/[dmy]/i.test(code.replace(/"[^"]*"|\[[^\]]*\]/g, '')) || code === 'General' || isNaN(+c)) return c;
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(+c * 86400000));
  const o = { timeZone: 'UTC' };
  if (/d/i.test(code)) o.day = 'numeric';
  if (/mmmm/i.test(code)) o.month = 'long'; else if (/mmm/i.test(code)) o.month = 'short'; else if (/m/i.test(code)) o.month = '2-digit';
  if (/y/i.test(code)) o.year = /yyyy/i.test(code) ? 'numeric' : '2-digit';
  return d.toLocaleDateString('ru-RU', o);
}
function niceScale(min, max, ticks) {
  if (min === max) { if (max === 0) max = 1; else if (max > 0) min = 0; else max = 0; }
  const range = max - min;
  const raw = range / ticks;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const step = [1, 2, 5, 10].map((k) => k * mag).find((s) => range / s <= ticks) || 10 * mag;
  return { min: Math.floor(min / step + 1e-9) * step, max: Math.ceil(max / step - 1e-9) * step, step };
}
function textWidth(s, px) { let w = 0; for (const ch of String(s)) w += /[0-9.,%\s]/.test(ch) ? 0.55 : /[A-ZА-ЯЁШЩЖМЮ]/.test(ch) ? 0.68 : 0.56; return w * px; }
function chartTxt(el) {
  const rich = q(el, 'tx/rich');
  if (rich) return kids(rich, 'p').map((p) => desc(p, 't').map((t) => t.textContent).join('')).join('\n');
  const s = desc(q(el, 'tx'), 'v')[0];
  return s ? s.textContent : '';
}
function txSize(el) {
  const d = desc(el, 'defRPr')[0] || desc(el, 'rPr')[0];
  return num(d, 'sz', null);
}
function txColor(el, ctx) {
  const d = desc(el, 'defRPr')[0];
  return d ? colorOf(findFill(d), ctx) : null;
}
function renderChart(cs, W, H, ctx) {
  const chart = kid(cs, 'chart');
  const plot = kid(chart, 'plotArea');
  if (!plot) return '';
  const baseSz = txSize(kid(cs, 'txPr')) || 1000;
  const fpx = (sz) => (sz || baseSz) / 100 * PT_PX;
  const baseColor = css(txColor(kid(cs, 'txPr'), ctx) || { rgb: [89, 89, 89], a: 1 });
  const font = fontCss(ctx.theme.minor);
  const out = [];
  const T = (x, y, s, o = {}) => out.push(`<text x="${r2(x)}" y="${r2(y)}" font-size="${r2(o.px || fpx())}" fill="${o.color || baseColor}"${o.anchor ? ` text-anchor="${o.anchor}"` : ''}${o.bold ? ' font-weight="700"' : ''} dominant-baseline="${o.base || 'central'}">${esc(s)}</text>`);

  const csFill = parseFill(findFill(kid(cs, 'spPr')), ctx);
  if (csFill && csFill.type === 'solid') out.push(`<rect width="${r2(W)}" height="${r2(H)}" fill="${css(csFill.color)}"/>`);

  const accent = (i) => {
    const c = { rgb: hexRgb(schemeHex('accent' + ((i % 6) + 1), ctx)), a: 1 };
    const round = Math.floor(i / 6) % 3;
    return round === 1 ? darken(c, 0.6) : round === 2 ? mix(c, 0.4) : c;
  };
  const groups = [];
  for (const g of kids(plot)) {
    const m = /^(bar|line|pie|doughnut|area|scatter|radar|bubble|ofPie|stock)(3D)?Chart$/.exec(g.localName);
    if (!m) continue;
    const type = { ofPie: 'pie', stock: 'line' }[m[1]] || m[1];
    const series = kids(g, 'ser').map((s) => {
      const spPr = kid(s, 'spPr');
      const ln = kid(spPr, 'ln');
      const dPts = {};
      for (const d of kids(s, 'dPt')) dPts[num(kid(d, 'idx'), 'val', 0)] = solidOf(parseFill(findFill(kid(d, 'spPr')), ctx));
      const marker = kid(s, 'marker');
      const tx = kid(s, 'tx');
      const nameEl = tx && (desc(tx, 'v')[0]);
      const val = kid(s, 'val') || kid(s, 'yVal');
      return {
        idx: num(kid(s, 'idx'), 'val', 0),
        name: nameEl ? nameEl.textContent : '',
        cats: refValues(kid(s, 'cat') || kid(s, 'xVal')),
        catFmt: formatCode(kid(s, 'cat')),
        vals: refValues(val).map((v) => (v == null || v === '' ? null : +v)),
        fmt: formatCode(val),
        fill: parseFill(findFill(spPr), ctx),
        lnFill: ln ? parseFill(findFill(ln), ctx) : null,
        lnW: Math.max(1, num(ln, 'w', 28575) * EMU_PX),
        dPts,
        symbol: attr(kid(marker, 'symbol'), 'val', null),
        mSize: num(kid(marker, 'size'), 'val', 5) * PT_PX,
        mFill: solidOf(parseFill(findFill(kid(marker, 'spPr')), ctx)),
        dLbls: kid(s, 'dLbls'),
        smooth: num(kid(s, 'smooth'), 'val', 0) === 1,
        sizes: refValues(kid(s, 'bubbleSize')).map((v) => +v || 0),
        explosion: num(kid(s, 'explosion'), 'val', 0) / 100,
        dExpl: Object.fromEntries(kids(s, 'dPt').map((d) => [num(kid(d, 'idx'), 'val', 0), num(kid(d, 'explosion'), 'val', 0) / 100])),
      };
    });
    groups.push({
      el: g, type, series,
      dir: attr(kid(g, 'barDir'), 'val', 'col'),
      grouping: attr(kid(g, 'grouping'), 'val', type === 'bar' ? 'clustered' : 'standard'),
      vary: num(kid(g, 'varyColors'), 'val', 0) === 1,
      gap: num(kid(g, 'gapWidth'), 'val', 150) / 100,
      overlap: num(kid(g, 'overlap'), 'val', 0) / 100,
      hole: num(kid(g, 'holeSize'), 'val', 50) / 100,
      firstAng: num(kid(g, 'firstSliceAng'), 'val', 0),
      dLbls: kid(g, 'dLbls'),
      markers: num(kid(g, 'marker'), 'val', 1) === 1,
      scatterStyle: attr(kid(g, 'scatterStyle'), 'val', m[1] === 'bubble' ? 'bubble' : 'marker'),
      bubbleScale: num(kid(g, 'bubbleScale'), 'val', 100) / 100,
    });
  }
  if (!groups.length) return '';
  const allSeries = groups.flatMap((g) => g.series);
  const serColor = (s, g) => {
    if (g.type === 'line' || g.type === 'scatter' || g.type === 'radar') return solidOf(s.lnFill) || solidOf(s.fill) || accent(s.idx);
    return solidOf(s.fill) || accent(s.idx);
  };
  const pieLike = groups[0].type === 'pie' || groups[0].type === 'doughnut';

  // Заголовок
  let top = 6, bottom = H - 6, left = 6, right = W - 6;
  const titleEl = kid(chart, 'title');
  const autoDeleted = num(kid(chart, 'autoTitleDeleted'), 'val', 0) === 1;
  let title = titleEl ? chartTxt(titleEl) : '';
  if (!title && !autoDeleted && allSeries.length === 1) title = allSeries[0].name;
  if (title) {
    const tpx = fpx(txSize(titleEl) || Math.round(baseSz * 1.4));
    const lines = title.split('\n');
    lines.forEach((ln, i) => T(W / 2, top + tpx * (0.7 + i * 1.2), ln, { px: tpx, anchor: 'middle', color: css(txColor(titleEl, ctx) || { rgb: [89, 89, 89], a: 1 }) }));
    top += tpx * (lines.length * 1.2 + 0.5);
  }

  // Легенда
  const legEl = kid(chart, 'legend');
  const entries = [];
  if (pieLike || (groups[0].vary && allSeries.length === 1 && groups[0].type !== 'line')) {
    const s = allSeries[0];
    (s.cats.length ? s.cats : s.vals.map((_, i) => String(i + 1))).forEach((c, i) => entries.push({ label: c == null ? '' : c, color: s.dPts[i] || accent(i), kind: 'rect' }));
  } else {
    for (const g of groups) for (const s of g.series) entries.push({ label: s.name || `Ряд ${s.idx + 1}`, color: serColor(s, g), kind: g.type === 'line' || g.type === 'scatter' || g.type === 'radar' ? 'line' : 'rect' });
  }
  const lpx = fpx(txSize(legEl));
  if (legEl && entries.length) {
    const pos = attr(kid(legEl, 'legendPos'), 'val', 'r');
    const sw = lpx * 0.8, gapX = lpx * 1.2;
    const lc = css(txColor(legEl, ctx) || { rgb: [89, 89, 89], a: 1 });
    const mark = (x, y, e) => out.push(e.kind === 'line'
      ? `<line x1="${r2(x)}" y1="${r2(y)}" x2="${r2(x + sw * 1.6)}" y2="${r2(y)}" stroke="${css(e.color)}" stroke-width="${r2(Math.max(2, lpx / 5))}"/>`
      : `<rect x="${r2(x)}" y="${r2(y - sw / 2)}" width="${r2(sw)}" height="${r2(sw)}" fill="${css(e.color)}"/>`);
    const mw = (e) => (e.kind === 'line' ? sw * 1.6 : sw) + lpx * 0.4;
    if (pos === 'r' || pos === 'l' || pos === 'tr') {
      const maxW = Math.min(W * 0.35, Math.max(...entries.map((e) => mw(e) + textWidth(e.label, lpx))));
      const lh = lpx * 1.5;
      const total = entries.length * lh;
      let y = pos === 'tr' ? top + lh / 2 : (top + bottom) / 2 - total / 2 + lh / 2;
      const x = pos === 'l' ? left + 4 : right - maxW - 4;
      for (const e of entries) { mark(x, y, e); T(x + mw(e), y, e.label, { px: lpx, color: lc }); y += lh; }
      if (pos === 'l') left += maxW + 10; else right -= maxW + 10;
    } else {
      const rowsArr = [[]];
      let rowW = 0;
      for (const e of entries) {
        const ew = mw(e) + textWidth(e.label, lpx) + gapX;
        if (rowW + ew > W - 12 && rowsArr[rowsArr.length - 1].length) { rowsArr.push([]); rowW = 0; }
        rowsArr[rowsArr.length - 1].push(e); rowW += ew;
      }
      const lh = lpx * 1.5;
      let y = pos === 't' ? top + lh / 2 : bottom - rowsArr.length * lh + lh / 2;
      for (const row of rowsArr) {
        const rw = row.reduce((a, e) => a + mw(e) + textWidth(e.label, lpx) + gapX, -gapX);
        let x = W / 2 - rw / 2;
        for (const e of row) { mark(x, y, e); T(x + mw(e), y, e.label, { px: lpx, color: lc }); x += mw(e) + textWidth(e.label, lpx) + gapX; }
        y += lh;
      }
      if (pos === 't') top += rowsArr.length * lh + 6; else bottom -= rowsArr.length * lh + 6;
    }
  }

  const lblOn = (s, g, key) => {
    for (const d of [s.dLbls, g.dLbls]) {
      if (!d) continue;
      if (num(kid(d, 'delete'), 'val', 0) === 1) return false;
      const k = kid(d, key);
      if (k) return num(k, 'val', 0) === 1;
    }
    return false;
  };
  const lblFmt = (s, g) => { for (const d of [s.dLbls, g.dLbls]) { const nf = kid(d, 'numFmt'); if (nf && attr(nf, 'sourceLinked', '0') !== '1') return attr(nf, 'formatCode', s.fmt); } return s.fmt; };
  const lblPx = (s, g) => fpx(txSize(s.dLbls) || txSize(g.dLbls));

  // Круговые
  if (pieLike) {
    const g = groups[0], s = g.series[0];
    const vals = s.vals.map((v) => Math.max(0, v || 0));
    const sum = vals.reduce((a, b) => a + b, 0) || 1;
    const showAny = lblOn(s, g, 'showVal') || lblOn(s, g, 'showPercent') || lblOn(s, g, 'showCatName');
    const cx = (left + right) / 2, cy = (top + bottom) / 2;
    const maxExpl = Math.max(s.explosion, ...Object.values(s.dExpl), 0);
    const r = Math.max(4, Math.min(right - left, bottom - top) / 2 * (showAny ? 0.8 : 0.92) / (1 + Math.min(0.5, maxExpl)));
    const hole = g.type === 'doughnut' ? r * g.hole : 0;
    let ang = (g.firstAng - 90) * Math.PI / 180;
    const lp = lblPx(s, g);
    vals.forEach((v, i) => {
      if (!v) return;
      const a2 = ang + (v / sum) * Math.PI * 2;
      const col = css(s.dPts[i] || accent(i));
      const ex = Math.min(0.5, s.dExpl[i] != null ? s.dExpl[i] : s.explosion) * r, mA = (ang + a2) / 2;
      const ecx = cx + Math.cos(mA) * ex, ecy = cy + Math.sin(mA) * ex;
      const p = (rad, t) => `${r2(ecx + Math.cos(t) * rad)} ${r2(ecy + Math.sin(t) * rad)}`;
      const large = a2 - ang > Math.PI ? 1 : 0;
      if (vals.filter(Boolean).length === 1) out.push(`<circle cx="${r2(cx)}" cy="${r2(cy)}" r="${r2(r)}" fill="${col}"/>${hole ? `<circle cx="${r2(cx)}" cy="${r2(cy)}" r="${r2(hole)}" fill="${csFill && csFill.type === 'solid' ? css(csFill.color) : '#fff'}"/>` : ''}`);
      else out.push(`<path d="M${p(hole, ang)}L${p(r, ang)}A${r2(r)} ${r2(r)} 0 ${large} 1 ${p(r, a2)}L${p(hole, a2)}${hole ? `A${r2(hole)} ${r2(hole)} 0 ${large} 0 ${p(hole, ang)}` : ''}Z" fill="${col}" stroke="#fff" stroke-width="1"/>`);
      if (showAny) {
        const parts = [];
        if (lblOn(s, g, 'showCatName')) parts.push(s.cats[i] || '');
        if (lblOn(s, g, 'showVal')) parts.push(fmtNum(v, lblFmt(s, g)));
        if (lblOn(s, g, 'showPercent')) parts.push(fmtNum(v / sum, '0%'));
        const mid = (ang + a2) / 2;
        const inside = g.type === 'doughnut' || (v / sum > 0.07);
        const rr = inside ? (hole ? (r + hole) / 2 : r * 0.62) : r * 1.15;
        const fill = inside ? '#fff' : baseColor;
        T(cx + Math.cos(mid) * rr, cy + Math.sin(mid) * rr, parts.join('; '), { px: lp, anchor: inside ? 'middle' : Math.cos(mid) >= 0 ? 'start' : 'end', color: fill });
      }
      ang = a2;
    });
    return wrapSvg(W, H, out, font);
  }

  // Радар
  if (groups[0].type === 'radar') {
    const g = groups[0];
    const cats = g.series[0] ? g.series[0].cats : [];
    const n = Math.max(3, cats.length || (g.series[0] ? g.series[0].vals.length : 3));
    const all = g.series.flatMap((s) => s.vals).filter((v) => v != null);
    const sc = niceScale(Math.min(0, ...all), Math.max(...all, 0), 4);
    const cx = (left + right) / 2, cy = (top + bottom) / 2, r = Math.min(right - left, bottom - top) / 2 * 0.78;
    const pt = (i, v) => { const t = -Math.PI / 2 + (i / n) * Math.PI * 2; const k = (v - sc.min) / (sc.max - sc.min || 1); return [cx + Math.cos(t) * r * k, cy + Math.sin(t) * r * k]; };
    for (let v = sc.min + sc.step; v <= sc.max + 1e-9; v += sc.step) out.push(`<polygon points="${[...Array(n).keys()].map((i) => pt(i, v).map(r2).join(',')).join(' ')}" fill="none" stroke="#D9D9D9"/>`);
    for (let i = 0; i < n; i++) { const [x, y] = pt(i, sc.max); out.push(`<line x1="${r2(cx)}" y1="${r2(cy)}" x2="${r2(x)}" y2="${r2(y)}" stroke="#D9D9D9"/>`); const [lx, ly] = pt(i, sc.max + (sc.max - sc.min) * 0.12); T(lx, ly, fmtCategory(cats[i] == null ? '' : cats[i], g.series[0].catFmt), { anchor: Math.abs(lx - cx) < 4 ? 'middle' : lx > cx ? 'start' : 'end' }); }
    for (const s of g.series) out.push(`<polygon points="${s.vals.map((v, i) => pt(i, v || 0).map(r2).join(',')).join(' ')}" fill="${attr(kid(g.el, 'radarStyle'), 'val', '') === 'filled' ? css(serColor(s, g)) : 'none'}" fill-opacity="0.5" stroke="${css(serColor(s, g))}" stroke-width="${r2(s.lnW)}"/>`);
    return wrapSvg(W, H, out, font);
  }

  // Декартовы графики: столбцы, линии, области, точечные
  const valAx = kids(plot, 'valAx');
  const catAx = kid(plot, 'catAx') || kid(plot, 'dateAx');
  const scatter = groups[0].type === 'scatter' || groups[0].type === 'bubble';
  const yAx = scatter ? valAx.find((a) => ['l', 'r'].includes(attr(kid(a, 'axPos'), 'val', 'l'))) || valAx[1] || valAx[0] : valAx[0];
  const xAx = scatter ? valAx.find((a) => a !== yAx) : catAx;
  const horizontal = groups.some((g) => g.type === 'bar' && g.dir === 'bar');
  const nCat = Math.max(1, ...allSeries.map((s) => Math.max(s.vals.length, s.cats.length)));
  const cats = (allSeries.find((s) => s.cats.length) || { cats: [] }).cats;

  // Диапазон значений с учётом накопления
  let vmin = Infinity, vmax = -Infinity;
  for (const g of groups) {
    if (g.type === 'scatter' || g.type === 'bubble') { for (const s of g.series) for (const v of s.vals) if (v != null) { vmin = Math.min(vmin, v); vmax = Math.max(vmax, v); } continue; }
    if (g.grouping === 'percentStacked') { vmin = Math.min(vmin, 0); vmax = Math.max(vmax, 1); continue; }
    if (g.grouping === 'stacked') {
      for (let i = 0; i < nCat; i++) {
        let pos = 0, neg = 0;
        for (const s of g.series) { const v = s.vals[i] || 0; if (v >= 0) pos += v; else neg += v; }
        vmin = Math.min(vmin, neg); vmax = Math.max(vmax, pos);
      }
    } else for (const s of g.series) for (const v of s.vals) if (v != null) { vmin = Math.min(vmin, v); vmax = Math.max(vmax, v); }
  }
  if (!isFinite(vmin)) { vmin = 0; vmax = 1; }
  if (!scatter || vmin > 0) vmin = Math.min(0, vmin);
  if (vmax < 0) vmax = 0;
  const scl = kid(yAx, 'scaling');
  let sc = niceScale(vmin, vmax, horizontal ? 5 : 6);
  const fixedMin = num(kid(scl, 'min'), 'val', null), fixedMax = num(kid(scl, 'max'), 'val', null);
  if (fixedMin != null) sc.min = fixedMin;
  if (fixedMax != null) sc.max = fixedMax;
  const fixedStep = num(kid(yAx, 'majorUnit'), 'val', null);
  if (fixedStep) sc.step = fixedStep;
  const percentAxis = groups.some((g) => g.grouping === 'percentStacked');
  const yFmt = attr(kid(yAx, 'numFmt'), 'sourceLinked', '1') === '0' ? attr(kid(yAx, 'numFmt'), 'formatCode', 'General') : percentAxis ? '0%' : (allSeries[0] && allSeries[0].fmt) || 'General';
  const yDeleted = num(kid(yAx, 'delete'), 'val', 0) === 1;
  const xDeleted = num(kid(xAx, 'delete'), 'val', 0) === 1;
  const ypx = fpx(txSize(kid(yAx, 'txPr'))), xpx = fpx(txSize(kid(xAx, 'txPr')));
  const ticks = [];
  for (let v = sc.min; v <= sc.max + sc.step * 1e-6; v += sc.step) ticks.push(+v.toFixed(10));

  // Скатерный график: свой диапазон по X
  let xsc = null;
  if (scatter) {
    const xs = groups[0].series.flatMap((s) => s.cats.map(Number)).filter((v) => !isNaN(v));
    xsc = niceScale(Math.min(...xs, Infinity) === Infinity ? 0 : Math.min(0, ...xs), xs.length ? Math.max(...xs) : 1, 5);
  }

  // Место под подписи осей
  const catFmt = (allSeries.find((s) => s.cats.length) || { catFmt: 'General' }).catFmt;
  const catLabels = cats.map((c) => (c == null ? '' : scatter ? c : fmtCategory(c, catFmt)));
  const maxCatW = Math.max(10, ...catLabels.map((c) => textWidth(c, xpx)));
  let rotateCats = false;
  if (horizontal) {
    if (!xDeleted) left += Math.min(W * 0.4, maxCatW) + 8;
    if (!yDeleted) bottom -= ypx * 1.8;
  } else {
    if (!yDeleted) left += Math.max(...ticks.map((t) => textWidth(fmtNum(t, yFmt), ypx))) + 8;
    if (!xDeleted) {
      rotateCats = !scatter && maxCatW > ((right - left) / nCat) * 0.95 && nCat <= 40;
      bottom -= rotateCats ? Math.min(H * 0.35, maxCatW * 0.72 + xpx) : xpx * 1.8;
    }
  }
  right -= 6;
  const ml = q(plot, 'layout/manualLayout');
  if (ml && attr(kid(ml, 'layoutTarget'), 'val', 'outer') === 'inner' && kid(ml, 'x')) {
    left = num(kid(ml, 'x'), 'val', 0) * W; top = num(kid(ml, 'y'), 'val', 0) * H;
    right = left + num(kid(ml, 'w'), 'val', 1) * W; bottom = top + num(kid(ml, 'h'), 'val', 1) * H;
  }
  const pw = Math.max(10, right - left), ph = Math.max(10, bottom - top);
  const vPos = (v) => (horizontal ? left + ((v - sc.min) / (sc.max - sc.min || 1)) * pw : bottom - ((v - sc.min) / (sc.max - sc.min || 1)) * ph);
  const slot = (horizontal ? ph : pw) / nCat;
  const catPos = (i) => (horizontal ? bottom - slot * (i + 0.5) : left + slot * (i + 0.5));
  const xPos = (x) => left + ((x - xsc.min) / (xsc.max - xsc.min || 1)) * pw;

  // Сетка и подписи значений
  const grid = !!kid(yAx, 'majorGridlines');
  // Подписи значений вдоль горизонтальной оси прореживаем, если не помещаются
  const tickW = Math.max(...ticks.map((t) => textWidth(fmtNum(t, yFmt), ypx))) + 6;
  const tickEvery = horizontal ? Math.max(1, Math.ceil(tickW / Math.max(1, pw / Math.max(1, ticks.length - 1)))) : 1;
  ticks.forEach((t, ti) => {
    const p = vPos(t);
    const lbl = ti % tickEvery === 0;
    if (grid) out.push(horizontal ? `<line x1="${r2(p)}" y1="${r2(top)}" x2="${r2(p)}" y2="${r2(bottom)}" stroke="#D9D9D9" stroke-width="0.75"/>` : `<line x1="${r2(left)}" y1="${r2(p)}" x2="${r2(right)}" y2="${r2(p)}" stroke="#D9D9D9" stroke-width="0.75"/>`);
    if (!yDeleted && lbl) {
      if (horizontal) T(p, bottom + ypx * 0.9, fmtNum(t, yFmt), { px: ypx, anchor: 'middle' });
      else T(left - 6, p, fmtNum(t, yFmt), { px: ypx, anchor: 'end' });
    }
  });
  if (scatter && xsc) {
    for (let v = xsc.min; v <= xsc.max + 1e-9; v += xsc.step) if (!xDeleted) T(xPos(v), bottom + xpx * 0.9, fmtNum(+v.toFixed(10), 'General'), { px: xpx, anchor: 'middle' });
  } else if (!xDeleted) {
    const fitEach = horizontal ? xpx * 1.3 : rotateCats ? xpx * 1.2 : maxCatW + 6;
    const step = Math.max(1, Math.ceil(nCat / Math.max(1, Math.floor((horizontal ? ph : pw) / fitEach))));
    for (let i = 0; i < nCat; i += step) {
      const c = catLabels[i] != null ? catLabels[i] : String(i + 1);
      if (horizontal) T(left - 6, catPos(i), c, { px: xpx, anchor: 'end' });
      else if (rotateCats) out.push(`<text transform="translate(${r2(catPos(i))} ${r2(bottom + xpx * 0.6)}) rotate(-45)" font-size="${r2(xpx)}" fill="${baseColor}" text-anchor="end" dominant-baseline="central">${esc(c)}</text>`);
      else T(catPos(i), bottom + xpx * 0.9, c, { px: xpx, anchor: 'middle' });
    }
  }
  const zero = vPos(Math.max(sc.min, Math.min(sc.max, 0)));
  out.push(horizontal ? `<line x1="${r2(zero)}" y1="${r2(top)}" x2="${r2(zero)}" y2="${r2(bottom)}" stroke="#BFBFBF"/>` : `<line x1="${r2(left)}" y1="${r2(zero)}" x2="${r2(right)}" y2="${r2(zero)}" stroke="#BFBFBF"/>`);

  // Ряды
  const labels = [];
  for (const g of groups) {
    if (g.type === 'bar') {
      const ns = g.series.length;
      const stacked = g.grouping === 'stacked' || g.grouping === 'percentStacked';
      const overlap = stacked ? 1 : Math.max(-1, Math.min(1, g.overlap));
      const bw = slot / (stacked ? 1 + g.gap : ns - (ns - 1) * overlap + g.gap);
      const totals = [];
      for (let i = 0; i < nCat; i++) totals[i] = g.series.reduce((a, s) => a + Math.abs(s.vals[i] || 0), 0) || 1;
      const posAcc = new Array(nCat).fill(0), negAcc = new Array(nCat).fill(0);
      g.series.forEach((s, si) => {
        for (let i = 0; i < nCat; i++) {
          let v = s.vals[i];
          if (v == null) continue;
          if (g.grouping === 'percentStacked') v = v / totals[i];
          let a0 = 0, a1 = v;
          if (stacked) { if (v >= 0) { a0 = posAcc[i]; a1 = posAcc[i] += v; } else { a0 = negAcc[i]; a1 = negAcc[i] += v; } }
          const off = stacked ? -bw / 2 : -((ns - (ns - 1) * overlap) * bw) / 2 + si * bw * (1 - overlap);
          const c0 = catPos(i) + (horizontal ? -off - bw : off);
          const p0 = vPos(a0), p1 = vPos(a1);
          const col = css(s.dPts[i] || (g.vary && ns === 1 ? accent(i) : serColor(s, g)));
          if (horizontal) out.push(`<rect x="${r2(Math.min(p0, p1))}" y="${r2(c0)}" width="${r2(Math.abs(p1 - p0))}" height="${r2(bw)}" fill="${col}"/>`);
          else out.push(`<rect x="${r2(c0)}" y="${r2(Math.min(p0, p1))}" width="${r2(bw)}" height="${r2(Math.abs(p1 - p0))}" fill="${col}"/>`);
          if (lblOn(s, g, 'showVal')) {
            const txt = fmtNum(s.vals[i], lblFmt(s, g));
            const lp = lblPx(s, g);
            if (horizontal) labels.push([stacked ? (p0 + p1) / 2 : p1 + 4, c0 + bw / 2, txt, { px: lp, anchor: stacked ? 'middle' : 'start', color: stacked ? '#fff' : baseColor }]);
            else labels.push([c0 + bw / 2, stacked ? (p0 + p1) / 2 : Math.min(p0, p1) - lp * 0.7, txt, { px: lp, anchor: 'middle', color: stacked ? '#fff' : baseColor }]);
          }
        }
      });
    } else if (g.type === 'line' || g.type === 'area') {
      const acc = new Array(nCat).fill(0);
      const totals = [];
      for (let i = 0; i < nCat; i++) totals[i] = g.series.reduce((a, s) => a + Math.abs(s.vals[i] || 0), 0) || 1;
      const prev = [];
      g.series.forEach((s) => {
        const pts = [];
        const base = [];
        for (let i = 0; i < nCat; i++) {
          let v = s.vals[i];
          if (v == null && g.type === 'line') { pts.push(null); continue; }
          v = v || 0;
          if (g.grouping === 'percentStacked') v /= totals[i];
          const b = g.grouping === 'standard' ? 0 : acc[i];
          if (g.grouping !== 'standard') acc[i] += v;
          base.push([catPos(i), vPos(b)]);
          pts.push([catPos(i), vPos(b + v)]);
        }
        const col = css(serColor(s, g));
        if (g.type === 'area') {
          const top2 = pts.filter(Boolean);
          const poly = top2.concat(base.slice().reverse());
          out.push(`<polygon points="${poly.map((p) => p.map(r2).join(',')).join(' ')}" fill="${col}" fill-opacity="${solidOf(s.fill) ? 1 : 0.85}"/>`);
        } else {
          const lineOff = s.lnFill && s.lnFill.type === 'none';
          const d = linePath(pts, s.smooth);
          if (!lineOff && d) out.push(`<path d="${d}" fill="none" stroke="${col}" stroke-width="${r2(s.lnW)}" stroke-linejoin="round" stroke-linecap="round"/>`);
          const showM = s.symbol ? s.symbol !== 'none' : g.markers;
          const sym = !s.symbol || s.symbol === 'auto' ? AUTO_MARKERS[s.idx % AUTO_MARKERS.length] : s.symbol;
          if (showM) for (const p of pts) if (p) out.push(markerSvg(sym, p[0], p[1], s.mSize, css(s.mFill || serColor(s, g))));
        }
        pts.forEach((p, i) => {
          if (p && lblOn(s, g, 'showVal')) labels.push([p[0], p[1] - lblPx(s, g) * 0.9, fmtNum(s.vals[i], lblFmt(s, g)), { px: lblPx(s, g), anchor: 'middle' }]);
        });
        prev.push(pts);
      });
    } else if (g.type === 'bubble' && xsc) {
      const maxSize = Math.max(...g.series.flatMap((s) => s.sizes), 1e-9);
      const maxR = Math.min(pw, ph) * 0.125 * g.bubbleScale;
      g.series.forEach((s) => {
        const col = css(serColor(s, g));
        s.vals.forEach((v, i) => {
          if (v == null || isNaN(+s.cats[i])) return;
          const rr = Math.sqrt(Math.max(0, s.sizes[i] || 0) / maxSize) * maxR;
          out.push(`<circle cx="${r2(xPos(+s.cats[i]))}" cy="${r2(vPos(v))}" r="${r2(rr)}" fill="${css(s.dPts[i] || serColor(s, g))}" fill-opacity="0.85"/>`);
        });
        void col;
      });
    } else if (g.type === 'scatter' && xsc) {
      g.series.forEach((s) => {
        const col = css(serColor(s, g));
        const pts = s.vals.map((v, i) => (v == null || s.cats[i] == null || isNaN(+s.cats[i]) ? null : [xPos(+s.cats[i]), vPos(v)]));
        const lineOn = !(s.lnFill && s.lnFill.type === 'none') && g.scatterStyle !== 'marker' && g.scatterStyle !== 'none';
        const smooth = s.smooth || /smooth/i.test(g.scatterStyle);
        if (lineOn) out.push(`<path d="${linePath(pts, smooth)}" fill="none" stroke="${col}" stroke-width="${r2(s.lnW)}"/>`);
        if (s.symbol !== 'none' && (s.symbol || !/^(line|smooth)$/.test(g.scatterStyle))) for (const p of pts) if (p) out.push(markerSvg(s.symbol || 'circle', p[0], p[1], s.mSize, css(s.mFill || serColor(s, g))));
      });
    }
  }
  for (const l of labels) T(...l);
  return wrapSvg(W, H, out, font);
}
// Ломаная или сглаженная линия (кривые Катмулла — Рома), разрывы на пустых точках
function linePath(pts, smooth) {
  let d = '';
  let run = [];
  const flush = () => {
    if (!run.length) return;
    d += `M${r2(run[0][0])} ${r2(run[0][1])}`;
    for (let i = 1; i < run.length; i++) {
      if (!smooth || run.length < 3) { d += `L${r2(run[i][0])} ${r2(run[i][1])}`; continue; }
      const p0 = run[i - 2] || run[i - 1], p1 = run[i - 1], p2 = run[i], p3 = run[i + 1] || p2;
      const c1 = [p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6], c2 = [p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6];
      d += `C${r2(c1[0])} ${r2(c1[1])} ${r2(c2[0])} ${r2(c2[1])} ${r2(p2[0])} ${r2(p2[1])}`;
    }
    run = [];
  };
  for (const p of pts) { if (p) run.push(p); else flush(); }
  flush();
  return d;
}
const AUTO_MARKERS = ['diamond', 'square', 'triangle', 'x', 'star', 'circle', 'plus'];
function markerSvg(sym, x, y, size, color) {
  const s = Math.max(3, size) / 2;
  switch (sym) {
    case 'square': return `<rect x="${r2(x - s)}" y="${r2(y - s)}" width="${r2(s * 2)}" height="${r2(s * 2)}" fill="${color}"/>`;
    case 'diamond': return `<path d="M${r2(x)} ${r2(y - s)}L${r2(x + s)} ${r2(y)}L${r2(x)} ${r2(y + s)}L${r2(x - s)} ${r2(y)}Z" fill="${color}"/>`;
    case 'triangle': return `<path d="M${r2(x)} ${r2(y - s)}L${r2(x + s)} ${r2(y + s)}L${r2(x - s)} ${r2(y + s)}Z" fill="${color}"/>`;
    case 'x': case 'star': case 'plus': case 'dash': case 'dot': return `<path d="M${r2(x - s)} ${r2(y - s)}L${r2(x + s)} ${r2(y + s)}M${r2(x + s)} ${r2(y - s)}L${r2(x - s)} ${r2(y + s)}" stroke="${color}" stroke-width="1.5"/>`;
    default: return `<circle cx="${r2(x)}" cy="${r2(y)}" r="${r2(s)}" fill="${color}"/>`;
  }
}
function wrapSvg(W, H, parts, font) {
  return `<svg class="sv-chart" width="${r2(W)}" height="${r2(H)}" viewBox="0 0 ${r2(W)} ${r2(H)}" font-family="${esc(font)}">${parts.join('')}</svg>`;
}

/* ---------- Presentation ---------- */
class Presentation {
  constructor(pkg) { this.pkg = pkg; this.themes = new Map(); this.styles = null; }
  async load() {
    this.root = await this.pkg.xml('ppt/presentation.xml');
    if (!this.root) throw new Error('В файле нет презентации (ppt/presentation.xml).');
    const sz = kid(this.root, 'sldSz');
    this.W = num(sz, 'cx', 12192000) * EMU_PX;
    this.H = num(sz, 'cy', 6858000) * EMU_PX;
    const rels = await this.pkg.rels('ppt/presentation.xml');
    this.slides = kids(kid(this.root, 'sldIdLst'), 'sldId').map((s) => rels[rid(s, 'id')]).filter(Boolean).map((r) => r.target);
  }
  theme(path) {
    if (!this.themes.has(path)) this.themes.set(path, loadTheme(this.pkg, path));
    return this.themes.get(path);
  }
  async tableStyle(id, ctx) {
    if (!this.styles) {
      const root = await this.pkg.xml('ppt/tableStyles.xml');
      this.styles = new Map(kids(root, 'tblStyle').map((s) => [attr(s, 'styleId', ''), s]));
      this.defaultStyleId = attr(root, 'def', '');
    }
    const el = this.styles.get(id || this.defaultStyleId);
    if (el) return parseTableStyle(el, ctx);
    if (!id && !this.defaultStyleId) return {};
    return defaultTableStyle(ctx);
  }
}

const STYLE = `
.sv-list{display:flex;flex-direction:column;align-items:center;gap:18px}
.sv-slide{position:relative;flex:none;box-shadow:0 1px 3px rgb(0 0 0/.12),0 6px 20px rgb(0 0 0/.06);background:#fff;overflow:hidden}
.sv-slide>.sv-num{position:absolute;right:6px;bottom:4px;font:11px/1 ui-monospace,monospace;color:rgb(0 0 0/.35);z-index:2;pointer-events:none}
.sv-canvas{position:absolute;left:0;top:0;transform-origin:0 0;overflow:hidden;color:#000;font-family:sans-serif;line-height:1.2;-webkit-text-size-adjust:none;text-size-adjust:none}
.sv-sp{position:absolute;box-sizing:border-box}
.sv-g{position:absolute;left:0;top:0;overflow:visible}
.sv-tx{position:absolute;box-sizing:border-box;display:flex;flex-direction:column;white-space:pre-wrap;word-wrap:break-word;overflow:visible}
.sv-p{margin:0}
.sv-bu{display:inline-block;text-indent:0}
.sv-pic{overflow:hidden}
.sv-pic img{position:absolute;max-width:none}
.sv-missing{background:repeating-linear-gradient(45deg,#eee 0 8px,#f6f6f6 8px 16px);display:flex;align-items:center;justify-content:center;font:12px sans-serif;color:#888}
.sv-tbl{position:absolute;left:0;top:0;border-collapse:collapse;table-layout:fixed}
.sv-tbl td{box-sizing:border-box;overflow:hidden}
.sv-chart{position:absolute;left:0;top:0;overflow:visible}
.sv-error{position:absolute;inset:0;display:flex;flex-direction:column;justify-content:center;padding:24px;font:14px sans-serif;color:#555;background:#fafafa}
`;

async function render(buf, host, opts) {
  opts = opts || {};
  if (!document.getElementById('sv-style')) {
    const st = document.createElement('style');
    st.id = 'sv-style';
    st.textContent = STYLE;
    document.head.appendChild(st);
  }
  const zip = await JSZip.loadAsync(buf);
  const pkg = new Pkg(zip);
  const pres = new Presentation(pkg);
  await pres.load();
  const renderer = new SlideRenderer(pkg, pres);
  const list = document.createElement('div');
  list.className = 'sv-list';
  host.innerHTML = '';
  host.appendChild(list);
  const slides = pres.slides.map((path, i) => {
    const wrap = document.createElement('div');
    wrap.className = 'sv-slide';
    const canvas = document.createElement('div');
    canvas.className = 'sv-canvas';
    canvas.style.width = r2(pres.W) + 'px';
    canvas.style.height = r2(pres.H) + 'px';
    wrap.appendChild(canvas);
    wrap.insertAdjacentHTML('beforeend', `<span class="sv-num">${i + 1}</span>`);
    list.appendChild(wrap);
    return { path, wrap, canvas, done: false };
  });
  let width = opts.width || 800;
  const setWidth = (w) => {
    width = w;
    const s = w / pres.W;
    for (const sl of slides) {
      sl.wrap.style.width = r2(w) + 'px';
      sl.wrap.style.height = r2(pres.H * s) + 'px';
      sl.canvas.style.transform = `scale(${s})`;
    }
  };
  setWidth(width);
  let disposed = false;
  const drawOne = async (sl, i) => {
    if (sl.done) return;
    sl.done = true;
    try {
      await renderer.render(sl.path, i + 1, sl.canvas);
    } catch (e) {
      console.warn('Слайд не нарисован', i + 1, e);
      sl.canvas.innerHTML = `<div class="sv-error">Этот слайд не получилось нарисовать.</div>`;
    }
  };
  // Первые слайды сразу, остальные по очереди, чтобы страница не подвисала
  for (let i = 0; i < Math.min(3, slides.length); i++) await drawOne(slides[i], i);
  const rest = (async () => {
    for (let i = 3; i < slides.length && !disposed; i++) {
      await new Promise((r) => setTimeout(r, 0));
      await drawOne(slides[i], i);
    }
  })();
  return {
    count: slides.length,
    finished: rest,
    setWidth,
    dispose() { disposed = true; pkg.dispose(); },
  };
}

window.PptxView = { render };
})();
