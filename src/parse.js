import { unzipSync, strFromU8 } from 'fflate';
import { eotToTtf } from 'mtx-decompressor';
import { fontKey, normalizeTypeface } from './fonts.js';

// Deck coordinates are CSS pixels at 96 dpi; OOXML uses 914400 EMU/inch.
const px = (v, fallback = 0) => v == null || v === '' ? fallback : Number(v) / 9525;
const children = (el, name) => Array.from(el?.childNodes ?? []).filter(n => n.nodeType === 1 && (!name || n.localName === name));
const child = (el, name) => children(el, name)[0];
const find = (el, name) => el?.getElementsByTagNameNS('*', name)[0];
const attr = (el, key, fallback) => el?.hasAttribute(key) ? el.getAttribute(key) : fallback;
const relId = el => el?.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
const points = value => Number(value) / 75;

const base64 = bytes => {
  let binary = '';
  for (let i = 0; i < bytes.length; i += 8192) binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
  return btoa(binary);
};

const fontFormat = bytes => {
  const signature = String.fromCharCode(...bytes.subarray(0, 4));
  if (bytes[0] === 0 && bytes[1] === 1 && bytes[2] === 0 && bytes[3] === 0) return { format: 'truetype', mimeType: 'font/ttf' };
  if (signature === 'OTTO' || signature === 'true' || signature === 'ttcf') return { format: 'opentype', mimeType: 'font/otf' };
  if (signature === 'wOFF') return { format: 'woff', mimeType: 'font/woff' };
  if (signature === 'wOF2') return { format: 'woff2', mimeType: 'font/woff2' };
  return null;
};

function deobfuscateOdttf(bytes, path) {
  const match = path.match(/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})(?:\.[^/]*)?$/i);
  if (!match || bytes.length < 32) return null;
  const guid = match[1].replaceAll('-', '').match(/../g).map(value => parseInt(value, 16)).reverse();
  const result = bytes.slice();
  for (let i = 0; i < 32; i++) result[i] ^= guid[i % 16];
  return fontFormat(result) ? result : null;
}

function extractEmbeddedFont(bytes, path) {
  let data = bytes;
  let detected = fontFormat(data);
  if (detected) return { bytes: data, ...detected };
  const decoded = deobfuscateOdttf(data, path);
  if (decoded) {
    data = decoded;
    detected = fontFormat(data);
    return { bytes: data, ...detected };
  }
  // PowerPoint font parts commonly use an EOT wrapper, often with MTX compression.
  if (data.length >= 82) {
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const fontDataSize = view.getUint32(4, true);
    const magic = view.getUint16(34, true);
    if (magic === 0x504c && fontDataSize > 0 && fontDataSize <= data.length) {
      try {
        const sfnt = eotToTtf(data);
        detected = fontFormat(sfnt);
        if (detected) return { bytes: sfnt, ...detected };
      } catch (error) {
        return { error: `The EOT font could not be decoded: ${error instanceof Error ? error.message : String(error)}` };
      }
      return { error: 'The EOT font decoder did not produce a supported OpenType payload.' };
    }
  }
  return { error: 'The embedded font format is not supported by this browser renderer.' };
}

function resolve(part, target) {
  const segments = (target.startsWith('/') ? target.slice(1) : part.slice(0, part.lastIndexOf('/') + 1) + target).split('/');
  const out = [];
  for (const segment of segments) {
    if (segment === '..') out.pop();
    else if (segment && segment !== '.') out.push(segment);
  }
  return out.join('/');
}

/** Import an OOXML presentation into serializable JSON. DOMParser is required. */
export async function parsePptx(input, options = {}) {
  const bytes = input instanceof Blob ? new Uint8Array(await input.arrayBuffer()) : new Uint8Array(input);
  if (bytes.byteLength > (options.maxBytes ?? 50 * 1024 * 1024)) throw new Error('Presentation exceeds the 50 MB import limit.');
  const Parser = options.DOMParser ?? globalThis.DOMParser;
  if (!Parser) throw new Error('PPTX import requires DOMParser (a browser or an injected implementation).');
  let expanded = 0;
  let entries = 0;
  const files = unzipSync(bytes, { filter: file => {
    expanded += file.originalSize;
    if (++entries > 10000 || expanded > 200 * 1024 * 1024) throw new Error('Presentation expands beyond the import limit.');
    return true;
  }});
  const docs = new Map();
  function xml(path) {
    if (!path || !files[path]) return null;
    if (!docs.has(path)) {
      const source = strFromU8(files[path]);
      if (/<!DOCTYPE|<!ENTITY/i.test(source)) throw new Error('Unsupported XML document declarations.');
      const doc = new Parser().parseFromString(source, 'application/xml');
      if (find(doc, 'parsererror')) throw new Error(`Invalid XML in ${path}`);
      docs.set(path, doc.documentElement);
    }
    return docs.get(path);
  }
  function relationships(part) {
    const slash = part.lastIndexOf('/');
    const root = xml(`${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`);
    return Object.fromEntries(children(root).map(r => [attr(r, 'Id'), {
      type: attr(r, 'Type', '').split('/').pop(),
      external: attr(r, 'TargetMode') === 'External',
      target: resolve(part, attr(r, 'Target', '')),
    }]));
  }
  const presentation = xml('ppt/presentation.xml');
  if (!presentation) throw new Error('This file is not an OOXML presentation.');
  const size = child(presentation, 'sldSz');
  const deck = { version: 2, width: px(attr(size, 'cx')), height: px(attr(size, 'cy')), slides: [], assets: {}, fonts: [], warnings: [] };
  if (!deck.width || !deck.height) throw new Error('Presentation has no valid slide size.');
  const fonts = new Map();
  const warn = (slideId, message, elementId) => deck.warnings.push({ slideId, elementId, message });
  const presRels = relationships('ppt/presentation.xml');
  const presStyle = child(presentation, 'defaultTextStyle');
  const embeddedFonts = new Map();
  for (const entry of children(child(presentation, 'embeddedFontLst'), 'embeddedFont')) {
    const family = attr(child(entry, 'font'), 'typeface');
    if (!family) continue;
    for (const [tag, bold, italic] of [['regular', false, false], ['bold', true, false], ['italic', false, true], ['boldItalic', true, true]]) {
      const face = child(entry, tag);
      const relation = presRels[relId(face)];
      if (!face || !relation || relation.external || !files[relation.target]) continue;
      const request = normalizeTypeface(family, bold, italic);
      const extracted = extractEmbeddedFont(files[relation.target], relation.target);
      if (extracted.error) {
        warn(undefined, `${family} ${tag}: ${extracted.error}`);
        continue;
      }
      embeddedFonts.set(fontKey(request), {
        part: relation.target,
        format: extracted.format,
        mimeType: extracted.mimeType,
        dataUrl: `data:${extracted.mimeType};base64,${base64(extracted.bytes)}`,
      });
    }
  }

  for (const [index, sldId] of children(child(presentation, 'sldIdLst'), 'sldId').entries()) {
    const part = presRels[relId(sldId)]?.target;
    const root = xml(part);
    if (!root) throw new Error(`Slide ${index + 1} is missing.`);
    const slideId = `slide-${attr(sldId, 'id', index + 1)}`;
    const rels = relationships(part);
    const related = (rs, type) => Object.values(rs).find(r => r.type === type && !r.external)?.target;
    const layoutPart = related(rels, 'slideLayout');
    const layout = xml(layoutPart);
    const masterPart = layoutPart && related(relationships(layoutPart), 'slideMaster');
    const master = xml(masterPart);
    const themePart = masterPart && related(relationships(masterPart), 'theme');
    const theme = xml(themePart ?? related(presRels, 'theme'));
    const colors = {};
    for (const c of children(find(theme, 'clrScheme'))) colors[c.localName] = attr(children(c)[0], 'val', attr(children(c)[0], 'lastClr', '000000'));
    const colorMap = find(root, 'overrideClrMapping') ?? find(layout, 'overrideClrMapping') ?? child(master, 'clrMap');
    function color(node, fallback = '#000000') {
      if (!node) return fallback;
      const c = children(node).find(c => ['srgbClr', 'schemeClr', 'sysClr'].includes(c.localName));
      if (!c) return fallback;
      let hex = c.localName === 'schemeClr' ? colors[attr(colorMap, attr(c, 'val'), attr(c, 'val'))] : attr(c, 'lastClr', attr(c, 'val'));
      hex = /^[0-9a-f]{6}$/i.test(hex ?? '') ? hex : '000000';
      let rgb = [0, 2, 4].map(offset => parseInt(hex.slice(offset, offset + 2), 16));
      for (const transform of children(c)) {
        const v = Number(attr(transform, 'val', 100000)) / 100000;
        if (transform.localName === 'tint') rgb = rgb.map(n => n + (255 - n) * v);
        if (transform.localName === 'shade' || transform.localName === 'lumMod') rgb = rgb.map(n => n * v);
        if (transform.localName === 'lumOff') rgb = rgb.map(n => n + 255 * v);
      }
      const alpha = Number(attr(child(c, 'alpha'), 'val', 100000)) / 100000;
      return `rgba(${rgb.map(n => Math.round(Math.max(0, Math.min(255, n)))).join(',')},${alpha})`;
    }
    function fill(el, fallback = 'none') {
      if (child(el, 'noFill')) return 'none';
      return color(child(el, 'solidFill'), fallback);
    }
    function fontName(name) {
      if (name?.startsWith('+')) name = attr(child(find(theme, name.startsWith('+mj') ? 'majorFont' : 'minorFont'), 'latin'), 'typeface', 'Arial');
      return name || 'Arial';
    }
    function runStyle(nodes) {
      const style = { fontFamily: 'Arial', fontSize: 24, color: '#000000', bold: false, italic: false, underline: false, letterSpacing: 0 };
      for (const n of nodes.filter(Boolean)) {
        if (n.hasAttribute('sz')) style.fontSize = points(attr(n, 'sz'));
        if (n.hasAttribute('b')) style.bold = attr(n, 'b') === '1';
        if (n.hasAttribute('i')) style.italic = attr(n, 'i') === '1';
        if (n.hasAttribute('u')) style.underline = attr(n, 'u') !== 'none';
        if (n.hasAttribute('spc')) style.letterSpacing = points(attr(n, 'spc'));
        if (child(n, 'latin')) style.fontFamily = fontName(attr(child(n, 'latin'), 'typeface'));
        if (child(n, 'solidFill')) style.color = color(child(n, 'solidFill'));
      }
      const normalized = normalizeTypeface(style.fontFamily, style.bold, style.italic);
      style.sourceTypeface = normalized.sourceTypeface;
      style.fontFamily = normalized.family;
      style.fontWeight = normalized.weight;
      style.italic = normalized.style === 'italic';
      style.bold = normalized.weight >= 700;
      style.fontId = fontKey(normalized);
      if (!fonts.has(style.fontId)) fonts.set(style.fontId, { id: style.fontId, ...normalized });
      return style;
    }
    function textBody(body, defaults = [], cellProps) {
      if (!body) return undefined;
      const bodyProps = child(body, 'bodyPr');
      const fit = child(bodyProps, 'normAutofit');
      const result = {
        anchor: attr(cellProps, 'anchor', attr(bodyProps, 'anchor', 't')),
        wrap: attr(bodyProps, 'wrap', 'square') !== 'none',
        autofit: fit ? 'shrink' : child(bodyProps, 'spAutoFit') ? 'grow' : 'none',
        fontScale: Number(attr(fit, 'fontScale', 100000)) / 100000,
        inset: {
          left: px(attr(cellProps, 'marL', attr(bodyProps, 'lIns')), 9.6),
          right: px(attr(cellProps, 'marR', attr(bodyProps, 'rIns')), 9.6),
          top: px(attr(cellProps, 'marT', attr(bodyProps, 'tIns')), 4.8),
          bottom: px(attr(cellProps, 'marB', attr(bodyProps, 'bIns')), 4.8),
        }, paragraphs: [],
      };
      for (const p of children(body, 'p')) {
        const ppr = child(p, 'pPr');
        const level = Number(attr(ppr, 'lvl', 0)) + 1;
        const nodes = [...defaults.map(d => child(d, `lvl${level}pPr`)), child(child(body, 'lstStyle'), `lvl${level}pPr`), ppr].filter(Boolean);
        const props = { align: 'l', marginLeft: 0, marginRight: 0, indent: 0, lineHeight: 1, before: 0, after: 0, bullet: null };
        for (const n of nodes) {
          props.align = attr(n, 'algn', props.align);
          if (n.hasAttribute('marL')) props.marginLeft = px(attr(n, 'marL'));
          if (n.hasAttribute('marR')) props.marginRight = px(attr(n, 'marR'));
          if (n.hasAttribute('indent')) props.indent = px(attr(n, 'indent'));
          for (const [tag, key] of [['spcBef', 'before'], ['spcAft', 'after']]) {
            const spacing = child(child(n, tag), 'spcPts');
            if (spacing) props[key] = points(attr(spacing, 'val'));
          }
          const line = child(n, 'lnSpc');
          if (child(line, 'spcPct')) { props.lineHeight = Number(attr(child(line, 'spcPct'), 'val')) / 100000; delete props.lineHeightPx; }
          if (child(line, 'spcPts')) props.lineHeightPx = points(attr(child(line, 'spcPts'), 'val'));
          if (child(n, 'buNone')) props.bullet = null;
          if (child(n, 'buChar')) props.bullet = attr(child(n, 'buChar'), 'char', '•');
          if (child(n, 'buAutoNum')) props.bullet = `${result.paragraphs.length + 1}.`;
        }
        const base = nodes.map(n => child(n, 'defRPr'));
        const runs = children(p).filter(n => ['r', 'fld', 'br'].includes(n.localName)).map(r => ({
          text: r.localName === 'br' ? '\n' : child(r, 't')?.textContent ?? '',
          style: runStyle([...base, child(r, 'rPr')]),
        }));
        if (!runs.length) runs.push({ text: '', style: runStyle([...base, child(p, 'endParaRPr')]) });
        result.paragraphs.push({ ...props, runs });
      }
      return result;
    }
    function frame(el) {
      const x = child(child(el, 'spPr'), 'xfrm') ?? child(el, 'xfrm');
      const off = child(x, 'off'), ext = child(x, 'ext');
      return { x: px(attr(off, 'x')), y: px(attr(off, 'y')), width: px(attr(ext, 'cx')), height: px(attr(ext, 'cy')),
        rotation: Number(attr(x, 'rot', 0)) / 60000, flipH: attr(x, 'flipH') === '1', flipV: attr(x, 'flipV') === '1' };
    }
    function asset(path) {
      if (!files[path]) return null;
      if (!deck.assets[path]) {
        const ext = path.split('.').pop().toLowerCase();
        const mimeType = ({ png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp' })[ext];
        if (!mimeType) return null;
        deck.assets[path] = { id: path, mimeType, dataUrl: `data:${mimeType};base64,${base64(files[path])}` };
      }
      return path;
    }
    const bg = child(child(root, 'cSld'), 'bg') ?? child(child(layout, 'cSld'), 'bg') ?? child(child(master, 'cSld'), 'bg');
    const slide = { id: slideId, sourcePart: part, background: fill(child(bg, 'bgPr'), '#ffffff'), elements: [] };
    for (const el of children(child(child(root, 'cSld'), 'spTree'))) {
      if (!['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp'].includes(el.localName)) continue;
      const nv = find(el, 'cNvPr');
      const id = `${slideId}/object-${attr(nv, 'id', slide.elements.length)}`;
      const base = { id, name: attr(nv, 'name', 'Object'), sourceId: attr(nv, 'id'), ...frame(el), locked: false };
      const sp = child(el, 'spPr');
      let element;
      if (el.localName === 'pic') {
        const blipFill = child(el, 'blipFill');
        const blip = child(blipFill, 'blip');
        const rid = blip?.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');
        const rel = rels[rid];
        const assetId = rel && !rel.external && asset(rel.target);
        const crop = child(blipFill, 'srcRect');
        if (assetId) element = { ...base, type: 'image', assetId,
          crop: Object.fromEntries(['l', 't', 'r', 'b'].map(k => [k, Number(attr(crop, k, 0)) / 100000])) };
      } else if (el.localName === 'sp' || el.localName === 'cxnSp') {
        const geometry = attr(child(sp, 'prstGeom'), 'prst', el.localName === 'cxnSp' ? 'line' : 'rect');
        if (['rect', 'roundRect', 'ellipse', 'line'].includes(geometry) && !child(sp, 'custGeom')) {
          const ph = find(el, 'ph');
          const isTextBox = attr(find(el, 'cNvSpPr'), 'txBox') === '1';
          // Free text boxes inherit presentation defaults. Placeholder master styles
          // would incorrectly add the master's body bullets or shape alignment.
          const textStyle = ['title', 'ctrTitle'].includes(attr(ph, 'type')) ? 'titleStyle' : ph ? 'bodyStyle' : isTextBox ? null : 'otherStyle';
          const text = textBody(child(el, 'txBody'), [presStyle, textStyle ? child(child(master, 'txStyles'), textStyle) : null]);
          const hasText = text?.paragraphs.some(p => p.runs.some(r => r.text));
          const line = child(sp, 'ln');
          const adjust = child(child(child(sp, 'prstGeom'), 'avLst'), 'gd');
          element = { ...base, type: attr(find(el, 'cNvSpPr'), 'txBox') === '1' ? 'text' : 'shape', geometry,
            fill: fill(sp), stroke: fill(line), strokeWidth: px(attr(line, 'w'), 1.333),
            radius: Math.min(base.width, base.height) * Number(attr(adjust, 'fmla', 'val 16667').split(' ').pop()) / 100000,
            text: hasText ? text : undefined };
        }
      } else if (el.localName === 'graphicFrame' && find(el, 'tbl')) {
        const tbl = find(el, 'tbl');
        const columns = children(child(tbl, 'tblGrid'), 'gridCol').map(c => px(attr(c, 'w')));
        const rows = children(tbl, 'tr').map(r => px(attr(r, 'h')));
        let y = 0;
        const cells = [];
        for (const [ri, row] of children(tbl, 'tr').entries()) {
          let x = 0;
          for (const [ci, cell] of children(row, 'tc').entries()) {
            const cp = child(cell, 'tcPr');
            if (['gridSpan', 'rowSpan', 'hMerge', 'vMerge'].some(k => cell.hasAttribute(k))) warn(slideId, 'Merged table cells are approximated.', id);
            cells.push({ id: `${id}/cell-${ri}-${ci}`, row: ri, column: ci, x, y, width: columns[ci], height: rows[ri], fill: fill(cp, '#ffffff'),
              borders: Object.fromEntries(['L', 'R', 'T', 'B'].map(side => { const line = child(cp, `ln${side}`); return [side, { color: fill(line), width: px(attr(line, 'w'), 1.333) }]; })),
              text: textBody(child(cell, 'txBody'), [presStyle], cp) });
            x += columns[ci];
          }
          y += rows[ri];
        }
        element = { ...base, type: 'table', columns, rows, cells };
      }
      if (!element) {
        element = { ...base, type: 'unsupported', reason: `Unsupported ${el.localName} object` };
        warn(slideId, element.reason, id);
      }
      if (find(el, 'gradFill') || find(el, 'outerShdw') || find(el, 'scene3d')) warn(slideId, 'Some visual effects are not rendered.', id);
      if (element.type === 'image' && slide.elements.length === 0 && Math.abs(base.x) < 1 && Math.abs(base.y) < 1 && Math.abs(base.width - deck.width) < 2 && Math.abs(base.height - deck.height) < 2) {
        element.role = 'background'; element.locked = true;
      }
      slide.elements.push(element);
    }
    for (const [source, label] of [[master, 'master'], [layout, 'layout']]) {
      const inherited = children(child(child(source, 'cSld'), 'spTree')).filter(e => ['sp', 'pic', 'grpSp', 'graphicFrame'].includes(e.localName) && !find(e, 'ph'));
      if (inherited.length && attr(root, 'showMasterSp') !== '0') warn(slideId, `Visible ${label} objects are not yet rendered.`);
    }
    if (find(root, 'timing')) warn(slideId, 'Animations are not rendered.');
    deck.slides.push(slide);
  }
  deck.fonts = [...fonts.values()].map(request => {
    const exact = embeddedFonts.get(request.id);
    if (exact) return { ...request, embedded: exact };
    const compatible = [...embeddedFonts.entries()]
      .filter(([key]) => key.startsWith(`${request.family.toLowerCase()}|`))
      .sort(([left], [right]) => {
        const [, leftWeight, leftStyle] = left.split('|');
        const [, rightWeight, rightStyle] = right.split('|');
        const score = (weight, style) => Math.abs(Number(weight) - request.weight) + (style === request.style ? 0 : 1000);
        return score(leftWeight, leftStyle) - score(rightWeight, rightStyle);
      })[0]?.[1];
    return compatible ? { ...request, embedded: compatible } : request;
  });
  return deck;
}
