import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { parsePptx } from './parse.js';

const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const EMU_PER_PIXEL = 9525;

const clone = value => typeof structuredClone === 'function' ? structuredClone(value) : JSON.parse(JSON.stringify(value));
const elements = (node, name) => Array.from(node?.childNodes ?? []).filter(child => child.nodeType === 1 && (!name || child.localName === name));
const first = (node, name) => node?.getElementsByTagNameNS('*', name)[0];
const all = (node, name) => Array.from(node?.getElementsByTagNameNS('*', name) ?? []);
const xmlPathForRelationships = part => {
  const slash = part.lastIndexOf('/');
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
};
const toBytes = async input => input instanceof Blob ? new Uint8Array(await input.arrayBuffer()) : new Uint8Array(input);
const textOf = body => body?.paragraphs?.map(paragraph => paragraph.runs.map(run => run.text).join('')).join('\n') ?? '';

function sourceHash(bytes) {
  let left = 2166136261;
  let right = 0x9e3779b9;
  for (const byte of bytes) {
    left = Math.imul(left ^ byte, 16777619);
    right = Math.imul(right ^ (byte + 1), 2246822519);
  }
  return `${bytes.byteLength.toString(36)}-${(left >>> 0).toString(36)}${(right >>> 0).toString(36)}`;
}

function bodyWithText(body, value) {
  if (!body?.paragraphs?.length) throw new Error('The selected object has no editable text body.');
  const next = clone(body);
  let wrote = false;
  for (const paragraph of next.paragraphs) {
    for (const run of paragraph.runs) {
      run.text = wrote ? '' : String(value);
      wrote = true;
    }
  }
  if (!wrote) throw new Error('The selected object has no editable text run.');
  return next;
}

function target(deck, operation) {
  const slide = deck.slides.find(item => item.id === operation.slideId);
  if (!slide) throw new Error(`Slide ${operation.slideId} does not exist.`);
  const element = slide.elements.find(item => item.id === operation.elementId);
  if (!element) throw new Error(`Element ${operation.elementId} does not exist on ${operation.slideId}.`);
  let cell;
  if (operation.cellId) {
    if (element.type !== 'table') throw new Error('cellId can only target a table.');
    cell = element.cells.find(item => item.id === operation.cellId);
    if (!cell) throw new Error(`Cell ${operation.cellId} does not exist.`);
  }
  return { slide, element, cell };
}

function frameOf(element) {
  return Object.fromEntries(['x', 'y', 'width', 'height', 'rotation', 'flipH', 'flipV'].map(key => [key, element[key]]));
}

function applyOne(deck, operation) {
  if (!operation || typeof operation !== 'object') throw new Error('Every edit operation must be an object.');
  if (operation.type === 'alignElements') {
    const slide = deck.slides.find(item => item.id === operation.slideId);
    if (!slide) throw new Error(`Slide ${operation.slideId} does not exist.`);
    const ids = [...new Set(operation.elementIds ?? [])];
    if (ids.length < 2) throw new Error('alignElements requires at least two elements.');
    const targets = ids.map(id => slide.elements.find(item => item.id === id));
    if (targets.some(item => !item)) throw new Error('One or more alignment targets do not exist.');
    const inverse = targets.map(element => ({ type: 'setFrame', slideId: slide.id, elementId: element.id, frame: frameOf(element) }));
    const alignment = operation.alignment;
    if (alignment) {
      const left = Math.min(...targets.map(item => item.x));
      const right = Math.max(...targets.map(item => item.x + item.width));
      const top = Math.min(...targets.map(item => item.y));
      const bottom = Math.max(...targets.map(item => item.y + item.height));
      for (const item of targets) {
        if (alignment === 'left') item.x = left;
        else if (alignment === 'center') item.x = (left + right - item.width) / 2;
        else if (alignment === 'right') item.x = right - item.width;
        else if (alignment === 'top') item.y = top;
        else if (alignment === 'middle') item.y = (top + bottom - item.height) / 2;
        else if (alignment === 'bottom') item.y = bottom - item.height;
        else throw new Error(`Unsupported alignment: ${alignment}`);
      }
    }
    if (operation.distribution === 'horizontal') {
      const ordered = [...targets].sort((a, b) => a.x - b.x);
      const start = ordered[0].x;
      const end = ordered.at(-1).x + ordered.at(-1).width;
      const gap = (end - start - ordered.reduce((sum, item) => sum + item.width, 0)) / (ordered.length - 1);
      let cursor = start;
      for (const item of ordered) { item.x = cursor; cursor += item.width + gap; }
    } else if (operation.distribution === 'vertical') {
      const ordered = [...targets].sort((a, b) => a.y - b.y);
      const start = ordered[0].y;
      const end = ordered.at(-1).y + ordered.at(-1).height;
      const gap = (end - start - ordered.reduce((sum, item) => sum + item.height, 0)) / (ordered.length - 1);
      let cursor = start;
      for (const item of ordered) { item.y = cursor; cursor += item.height + gap; }
    } else if (operation.distribution && operation.distribution !== 'none') throw new Error(`Unsupported distribution: ${operation.distribution}`);
    return inverse;
  }

  const { element, cell } = target(deck, operation);
  if (operation.type === 'setText') {
    const owner = cell ?? element;
    const previous = textOf(owner.text);
    owner.text = bodyWithText(owner.text, operation.text);
    return [{ type: 'setText', slideId: operation.slideId, elementId: operation.elementId, ...(operation.cellId ? { cellId: operation.cellId } : {}), text: previous }];
  }
  if (operation.type === 'setTextStyle') {
    const owner = cell ?? element;
    if (!owner.text) throw new Error('The selected object has no editable text style.');
    const previous = {};
    for (const paragraph of owner.text.paragraphs) for (const run of paragraph.runs) {
      for (const [key, value] of Object.entries(operation.style ?? {})) {
        if (!(key in previous)) previous[key] = run.style[key];
        run.style[key] = value;
      }
    }
    return [{ type: 'setTextStyle', slideId: operation.slideId, elementId: operation.elementId, ...(operation.cellId ? { cellId: operation.cellId } : {}), style: previous }];
  }
  if (operation.type === 'setFrame') {
    if (cell) throw new Error('Table cells cannot be moved independently.');
    const previous = frameOf(element);
    for (const key of ['x', 'y', 'width', 'height', 'rotation', 'flipH', 'flipV']) {
      if (operation.frame?.[key] !== undefined) element[key] = operation.frame[key];
    }
    if (![element.x, element.y, element.width, element.height, element.rotation].every(Number.isFinite) || element.width < 0 || element.height < 0) throw new Error('setFrame contains invalid geometry.');
    return [{ type: 'setFrame', slideId: operation.slideId, elementId: operation.elementId, frame: previous }];
  }
  if (operation.type === 'setShapeStyle') {
    if (!['shape', 'text'].includes(element.type)) throw new Error('setShapeStyle requires a shape or text element.');
    const previous = {};
    for (const key of ['fill', 'stroke', 'strokeWidth']) if (operation.style?.[key] !== undefined) {
      previous[key] = element[key];
      element[key] = operation.style[key];
    }
    return [{ type: 'setShapeStyle', slideId: operation.slideId, elementId: operation.elementId, style: previous }];
  }
  if (operation.type === 'replaceImage') {
    if (element.type !== 'image') throw new Error('replaceImage requires an image element.');
    const asset = operation.asset;
    if (!asset?.id || !asset?.mimeType || !asset?.dataUrl) throw new Error('replaceImage requires an asset with id, mimeType, and dataUrl.');
    const previous = deck.assets[element.assetId];
    if (!previous) throw new Error('The selected image has no source asset.');
    deck.assets[asset.id] = clone(asset);
    element.assetId = asset.id;
    return [{ type: 'replaceImage', slideId: operation.slideId, elementId: operation.elementId, asset: clone(previous) }];
  }
  throw new Error(`Unsupported edit operation: ${operation.type}`);
}

function parseXml(files, path, Parser) {
  const source = files[path];
  if (!source) throw new Error(`Missing OOXML part: ${path}`);
  const document = new Parser().parseFromString(strFromU8(source), 'application/xml');
  if (first(document, 'parsererror')) throw new Error(`Invalid XML in ${path}`);
  return document;
}

function sourceElement(document, sourceId) {
  return all(document, 'cNvPr').find(node => node.getAttribute('id') === String(sourceId))?.parentNode?.parentNode;
}

function patchFrame(node, frame) {
  const xfrm = first(node, 'xfrm');
  const off = first(xfrm, 'off');
  const ext = first(xfrm, 'ext');
  if (!xfrm || !off || !ext) throw new Error('The target has no editable transform.');
  off.setAttribute('x', String(Math.round(frame.x * EMU_PER_PIXEL)));
  off.setAttribute('y', String(Math.round(frame.y * EMU_PER_PIXEL)));
  ext.setAttribute('cx', String(Math.round(frame.width * EMU_PER_PIXEL)));
  ext.setAttribute('cy', String(Math.round(frame.height * EMU_PER_PIXEL)));
  if (frame.rotation) xfrm.setAttribute('rot', String(Math.round(frame.rotation * 60000))); else xfrm.removeAttribute('rot');
  if (frame.flipH) xfrm.setAttribute('flipH', '1'); else xfrm.removeAttribute('flipH');
  if (frame.flipV) xfrm.setAttribute('flipV', '1'); else xfrm.removeAttribute('flipV');
}

function textTarget(node, operation) {
  if (!operation.cellId) return first(node, 'txBody');
  const match = String(operation.cellId).match(/\/cell-(\d+)-(\d+)$/);
  if (!match) throw new Error('Invalid table cell identifier.');
  const row = elements(first(node, 'tbl'), 'tr')[Number(match[1])];
  const cell = elements(row, 'tc')[Number(match[2])];
  return first(cell, 'txBody');
}

function patchText(body, value) {
  const nodes = all(body, 't');
  if (!nodes.length) throw new Error('The target has no editable text nodes.');
  nodes.forEach((node, index) => {
    node.textContent = index ? '' : String(value);
    if (!index && (/^\s|\s$/.test(String(value)))) node.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  });
}

function normalizeHex(color) {
  const value = String(color ?? '').trim();
  const short = value.match(/^#([0-9a-f]{3})$/i);
  if (short) return short[1].split('').map(char => char + char).join('').toUpperCase();
  const full = value.match(/^#([0-9a-f]{6})$/i);
  if (full) return full[1].toUpperCase();
  const rgb = value.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
  return rgb ? rgb.slice(1, 4).map(part => Number(part).toString(16).padStart(2, '0')).join('').toUpperCase() : null;
}

function ensureChild(document, parent, localName) {
  let result = elements(parent, localName)[0];
  if (!result) {
    result = document.createElementNS(DRAWING_NS, `a:${localName}`);
    parent.appendChild(result);
  }
  return result;
}

function setSolidFill(document, parent, color) {
  for (const node of elements(parent).filter(item => ['solidFill', 'noFill', 'gradFill', 'pattFill'].includes(item.localName))) parent.removeChild(node);
  if (color === 'none') {
    parent.appendChild(document.createElementNS(DRAWING_NS, 'a:noFill'));
    return;
  }
  const hex = normalizeHex(color);
  if (!hex) return;
  const fill = document.createElementNS(DRAWING_NS, 'a:solidFill');
  const rgb = document.createElementNS(DRAWING_NS, 'a:srgbClr');
  rgb.setAttribute('val', hex);
  fill.appendChild(rgb);
  parent.appendChild(fill);
}

function patchTextStyle(document, body, style) {
  const runs = all(body, 'r');
  for (const run of runs) {
    let props = elements(run, 'rPr')[0];
    if (!props) {
      props = document.createElementNS(DRAWING_NS, 'a:rPr');
      run.insertBefore(props, run.firstChild);
    }
    if (style.bold !== undefined) props.setAttribute('b', style.bold ? '1' : '0');
    if (style.italic !== undefined) props.setAttribute('i', style.italic ? '1' : '0');
    if (style.underline !== undefined) props.setAttribute('u', style.underline ? 'sng' : 'none');
    if (style.fontSize !== undefined) props.setAttribute('sz', String(Math.round(Number(style.fontSize) * 75)));
    if (style.fontFamily) ensureChild(document, props, 'latin').setAttribute('typeface', String(style.fontFamily));
    if (style.color) setSolidFill(document, props, style.color);
  }
}

function dataUrlBytes(dataUrl) {
  const match = String(dataUrl).match(/^data:([^;,]+);base64,(.+)$/s);
  if (!match) throw new Error('Only base64 image data URLs can be exported.');
  const binary = atob(match[2]);
  return Uint8Array.from(binary, char => char.charCodeAt(0));
}

function uniqueMediaPath(files, mimeType) {
  const extension = ({ 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp' })[mimeType];
  if (!extension) throw new Error(`Unsupported replacement image type: ${mimeType}`);
  let index = 1;
  while (files[`ppt/media/butterstudio-image-${index}.${extension}`]) index += 1;
  return `ppt/media/butterstudio-image-${index}.${extension}`;
}

function ensureImageContentType(files, mediaPath, mimeType, Parser, Serializer) {
  const path = '[Content_Types].xml';
  const document = parseXml(files, path, Parser);
  const extension = mediaPath.split('.').pop();
  const exists = all(document, 'Default').some(node => node.getAttribute('Extension')?.toLowerCase() === extension.toLowerCase());
  if (!exists) {
    const entry = document.createElementNS(CONTENT_TYPES_NS, 'Default');
    entry.setAttribute('Extension', extension);
    entry.setAttribute('ContentType', mimeType);
    document.documentElement.appendChild(entry);
    files[path] = strToU8(new Serializer().serializeToString(document));
  }
}

function patchImage(files, slideDocument, slidePart, node, asset, Parser, Serializer) {
  const mediaPath = uniqueMediaPath(files, asset.mimeType);
  files[mediaPath] = dataUrlBytes(asset.dataUrl);
  ensureImageContentType(files, mediaPath, asset.mimeType, Parser, Serializer);
  const relPath = xmlPathForRelationships(slidePart);
  const relDocument = parseXml(files, relPath, Parser);
  const ids = all(relDocument, 'Relationship').map(item => Number((item.getAttribute('Id') ?? '').replace(/^rId/, ''))).filter(Number.isFinite);
  const relationId = `rId${Math.max(0, ...ids) + 1}`;
  const relation = relDocument.createElementNS(PACKAGE_REL_NS, 'Relationship');
  relation.setAttribute('Id', relationId);
  relation.setAttribute('Type', IMAGE_REL_TYPE);
  relation.setAttribute('Target', `../media/${mediaPath.split('/').pop()}`);
  relDocument.documentElement.appendChild(relation);
  files[relPath] = strToU8(new Serializer().serializeToString(relDocument));
  const blip = first(node, 'blip');
  if (!blip) throw new Error('The selected image has no OOXML image reference.');
  blip.setAttributeNS(REL_NS, 'r:embed', relationId);
  return [mediaPath, relPath, '[Content_Types].xml'];
}

/** Open a PPTX as a mutable editing session while retaining its OOXML package. */
export async function openPptx(input, options = {}) {
  const sourceBytes = await toBytes(input);
  const Parser = options.DOMParser ?? globalThis.DOMParser;
  const Serializer = options.XMLSerializer ?? globalThis.XMLSerializer;
  if (!Parser || !Serializer) throw new Error('PPTX editing requires DOMParser and XMLSerializer implementations.');
  let deck = options.deck ? clone(options.deck) : await parsePptx(sourceBytes, { ...options, DOMParser: Parser });
  if (deck?.version !== 2 || !Array.isArray(deck.slides) || !deck.assets || typeof deck.assets !== 'object') {
    throw new Error('The supplied Deck JSON is not a supported version 2 document.');
  }
  const originalDeck = clone(deck);
  const operations = [];
  const inverseOperations = [];
  const revision = sourceHash(sourceBytes);

  return {
    get deck() { return deck; },
    revision,
    applyOperations(nextOperations) {
      if (!Array.isArray(nextOperations) || !nextOperations.length) throw new Error('At least one edit operation is required.');
      if (nextOperations.length > 500) throw new Error('A single edit batch cannot exceed 500 operations.');
      const candidate = clone(deck);
      const inverses = [];
      for (const operation of nextOperations) inverses.unshift(...applyOne(candidate, clone(operation)));
      deck = candidate;
      operations.push(...clone(nextOperations));
      inverseOperations.unshift(...inverses);
      return { deck: clone(deck), inverseOperations: clone(inverses) };
    },
    async exportPptx() {
      const files = unzipSync(sourceBytes);
      const changedParts = new Set();
      const bySlide = new Map();
      for (const operation of operations) {
        const list = bySlide.get(operation.slideId) ?? [];
        list.push(operation);
        bySlide.set(operation.slideId, list);
      }
      for (const [slideId, slideOperations] of bySlide) {
        const originalSlide = originalDeck.slides.find(item => item.id === slideId);
        const currentSlide = deck.slides.find(item => item.id === slideId);
        if (!originalSlide || !currentSlide) throw new Error(`Cannot export missing slide ${slideId}.`);
        const document = parseXml(files, originalSlide.sourcePart, Parser);
        for (const operation of slideOperations) {
          if (operation.type === 'alignElements') continue;
          const original = originalSlide.elements.find(item => item.id === operation.elementId);
          const current = currentSlide.elements.find(item => item.id === operation.elementId);
          const node = original && sourceElement(document, original.sourceId);
          if (!original || !current || !node) throw new Error(`Cannot locate ${operation.elementId} in source OOXML.`);
          if (operation.type === 'setText') patchText(textTarget(node, operation), operation.text);
          else if (operation.type === 'setTextStyle') patchTextStyle(document, textTarget(node, operation), operation.style ?? {});
          else if (operation.type === 'setFrame') patchFrame(node, current);
          else if (operation.type === 'setShapeStyle') {
            const spPr = first(node, 'spPr');
            if (operation.style?.fill !== undefined) setSolidFill(document, spPr, current.fill);
            if (operation.style?.stroke !== undefined || operation.style?.strokeWidth !== undefined) {
              const line = ensureChild(document, spPr, 'ln');
              if (operation.style?.stroke !== undefined) setSolidFill(document, line, current.stroke);
              if (operation.style?.strokeWidth !== undefined) line.setAttribute('w', String(Math.round(current.strokeWidth * EMU_PER_PIXEL)));
            }
          } else if (operation.type === 'replaceImage') {
            for (const part of patchImage(files, document, originalSlide.sourcePart, node, deck.assets[current.assetId], Parser, Serializer)) changedParts.add(part);
          }
        }
        // Alignment operations change several frames and are exported from final Deck JSON.
        const alignedIds = slideOperations.filter(item => item.type === 'alignElements').flatMap(item => item.elementIds ?? []);
        for (const id of new Set(alignedIds)) {
          const original = originalSlide.elements.find(item => item.id === id);
          const current = currentSlide.elements.find(item => item.id === id);
          const node = original && sourceElement(document, original.sourceId);
          if (!original || !current || !node) throw new Error(`Cannot locate ${id} in source OOXML.`);
          patchFrame(node, current);
        }
        files[originalSlide.sourcePart] = strToU8(new Serializer().serializeToString(document));
        changedParts.add(originalSlide.sourcePart);
      }
      return { data: zipSync(files, { level: 6 }), deck: clone(deck), inverseOperations: clone(inverseOperations), changedParts: [...changedParts] };
    },
  };
}
