import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { DOMParser } from '@xmldom/xmldom';
import { zipSync, strToU8 } from 'fflate';
import { getRequiredFonts, googleFonts, normalizeTypeface, parsePptx, resolveFonts } from '../src/index.js';

const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const relationships = body => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
const rel = (id, target, type) => `<Relationship Id="${id}" Target="${target}" Type="${r}/${type}"/>`;
function fixture(overrides = {}) {
  const shape = `<p:sp><p:nvSpPr><p:cNvPr id="7" name="Title"/><p:cNvSpPr txBox="1"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="95250" y="190500"/><a:ext cx="1905000" cy="476250"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr anchor="ctr"><a:spAutoFit/></a:bodyPr><a:lstStyle><a:lvl1pPr><a:defRPr sz="3000"><a:latin typeface="Poppins Bold"/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:t>Hello &amp; world</a:t></a:r><a:r><a:rPr i="1"/><a:t>!</a:t></a:r></a:p></p:txBody></p:sp>`;
  const parts = {
    'ppt/presentation.xml': `<p:presentation xmlns:p="${p}" xmlns:a="${a}" xmlns:r="${r}"><p:sldIdLst><p:sldId id="99" r:id="second"/><p:sldId id="42" r:id="first"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:defaultTextStyle><a:lvl1pPr algn="l"><a:defRPr sz="1800"/></a:lvl1pPr></p:defaultTextStyle></p:presentation>`,
    'ppt/_rels/presentation.xml.rels': relationships(rel('first', 'slides/slide1.xml', 'slide') + rel('second', 'slides/slide2.xml', 'slide')),
    'ppt/slides/slide1.xml': `<p:sld xmlns:p="${p}" xmlns:a="${a}"><p:cSld><p:spTree>${shape}</p:spTree></p:cSld></p:sld>`,
    'ppt/slides/slide2.xml': `<p:sld xmlns:p="${p}" xmlns:a="${a}"><p:cSld><p:spTree>${shape}</p:spTree></p:cSld></p:sld>`,
    'ppt/slides/_rels/slide2.xml.rels': relationships(rel('layout', '../slideLayouts/slideLayout1.xml', 'slideLayout')),
    'ppt/slideLayouts/slideLayout1.xml': `<p:sldLayout xmlns:p="${p}"/>`,
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': relationships(rel('master', '../slideMasters/slideMaster1.xml', 'slideMaster')),
    'ppt/slideMasters/slideMaster1.xml': `<p:sldMaster xmlns:p="${p}" xmlns:a="${a}"><p:txStyles><p:otherStyle><a:lvl1pPr algn="r"/></p:otherStyle><p:bodyStyle><a:lvl1pPr><a:buChar char="•"/></a:lvl1pPr></p:bodyStyle></p:txStyles></p:sldMaster>`,
    ...overrides,
  };
  return zipSync(Object.fromEntries(Object.entries(parts).map(([key, value]) => [key, typeof value === 'string' ? strToU8(value) : value])));
}

test('uses relationship slide order and scopes object IDs to each slide', async () => {
  const d = await parsePptx(fixture(), { DOMParser });
  assert.equal(d.width, 1280);
  assert.equal(d.height, 720);
  assert.deepEqual(d.slides.map(s => s.sourcePart), ['ppt/slides/slide2.xml', 'ppt/slides/slide1.xml']);
  assert.notEqual(d.slides[0].elements[0].id, d.slides[1].elements[0].id);
  assert.deepEqual(JSON.parse(JSON.stringify(d)), d);
});

test('free text uses local list/run styles without inheriting placeholder bullets', async () => {
  const d = await parsePptx(fixture(), { DOMParser });
  const e = d.slides[0].elements[0];
  assert.equal(e.x, 10);
  assert.equal(e.y, 20);
  const p = e.text.paragraphs[0];
  assert.equal(p.align, 'l');
  assert.equal(p.bullet, null);
  assert.equal(p.runs[0].text, 'Hello & world');
  assert.equal(p.runs[0].style.fontSize, 40);
  assert.equal(p.runs[1].style.fontFamily, 'Poppins');
  assert.equal(p.runs[1].style.sourceTypeface, 'Poppins Bold');
  assert.equal(p.runs[1].style.fontWeight, 700);
  assert.equal(p.runs[1].style.italic, true);
});

test('normalizes PowerPoint face suffixes and retains a serializable font manifest', async () => {
  assert.deepEqual(normalizeTypeface('Poppins SemiBold Italic'), {
    family: 'Poppins', weight: 600, style: 'italic', sourceTypeface: 'Poppins SemiBold Italic',
  });
  const d = await parsePptx(fixture(), { DOMParser });
  assert.deepEqual(d.fonts, [{
    id: 'poppins|700|normal', family: 'Poppins', weight: 700, style: 'normal', sourceTypeface: 'Poppins Bold',
  }, {
    id: 'poppins|700|italic', family: 'Poppins', weight: 700, style: 'italic', sourceTypeface: 'Poppins Bold',
  }]);
  assert.deepEqual(JSON.parse(JSON.stringify(d.fonts)), d.fonts);
  assert.deepEqual(getRequiredFonts(d).map(font => ({ id: font.id, text: font.text })), [
    { id: 'poppins|700|normal', text: 'Helo &wrd' },
    { id: 'poppins|700|italic', text: '!' },
  ]);
});

test('extracts directly embedded OpenType font data from the presentation manifest', async () => {
  const presentation = `<p:presentation xmlns:p="${p}" xmlns:a="${a}" xmlns:r="${r}"><p:sldIdLst><p:sldId id="99" r:id="second"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:embeddedFontLst><p:embeddedFont><p:font typeface="Poppins"/><p:bold r:id="font1"/></p:embeddedFont></p:embeddedFontLst></p:presentation>`;
  const font = new Uint8Array([0, 1, 0, 0, 0, 0, 0, 0]);
  const d = await parsePptx(fixture({
    'ppt/presentation.xml': presentation,
    'ppt/_rels/presentation.xml.rels': relationships(rel('second', 'slides/slide2.xml', 'slide') + rel('font1', 'fonts/font1.fntdata', 'font')),
    'ppt/fonts/font1.fntdata': font,
  }), { DOMParser });
  const bold = d.fonts.find(item => item.id === 'poppins|700|normal');
  assert.equal(bold.embedded.format, 'truetype');
  assert.equal(bold.embedded.part, 'ppt/fonts/font1.fntdata');
  assert.match(bold.embedded.dataUrl, /^data:font\/ttf;base64,/);
});

test('resolves fonts in embedded, consumer, system, default order', async () => {
  class FakeFontFace {
    constructor(family, source) { this.family = family; this.source = source; }
    async load() {
      if (typeof this.source === 'string' && this.source.startsWith('local(') && !this.source.includes('System Face')) throw new Error('not installed');
      return this;
    }
  }
  const fontSet = new Set();
  const requests = [
    { id: 'embedded|400|normal', family: 'Embedded', sourceTypeface: 'Embedded', weight: 400, style: 'normal', embedded: { dataUrl: 'data:font/ttf;base64,AAEAAA==', format: 'truetype', part: 'ppt/fonts/font1.fntdata', mimeType: 'font/ttf' } },
    { id: 'consumer|400|normal', family: 'Consumer', sourceTypeface: 'Consumer', weight: 400, style: 'normal' },
    { id: 'system face|400|normal', family: 'System Face', sourceTypeface: 'System Face', weight: 400, style: 'normal' },
    { id: 'missing|400|normal', family: 'Missing', sourceTypeface: 'Missing', weight: 400, style: 'normal' },
  ];
  const asked = [];
  const session = await resolveFonts({ fonts: requests }, {
    FontFace: FakeFontFace,
    fontSet,
    resolveFont(request) {
      asked.push(request.family);
      return request.family === 'Consumer' ? { data: new Uint8Array([0, 1, 0, 0]) } : null;
    },
  });
  assert.equal(session.resolutions['embedded|400|normal'].source, 'embedded');
  assert.equal(session.resolutions['consumer|400|normal'].source, 'resolver');
  assert.equal(session.resolutions['system face|400|normal'].source, 'system');
  assert.equal(session.resolutions['missing|400|normal'].source, 'default');
  assert.deepEqual(asked, ['Consumer', 'System Face', 'Missing']);
  assert.equal(session.warnings.length, 1);
  session.destroy();
  assert.equal(fontSet.size, 0);
});

test('Google Fonts resolver requests an exact subset and caches the downloaded face', async () => {
  const calls = [];
  const bytes = new Uint8Array([119, 79, 70, 50]);
  const fetcher = async input => {
    calls.push(String(input));
    if (String(input).startsWith('https://fonts.googleapis.com/css2')) {
      return {
        ok: true,
        async text() {
          return `@font-face { src: url(https://fonts.gstatic.com/l/font.woff2) format('woff2'); }`;
        },
      };
    }
    return { ok: true, async arrayBuffer() { return bytes.buffer; } };
  };
  const resolver = googleFonts({ fetch: fetcher });
  const request = { family: 'Poppins', weight: 600, style: 'italic', text: 'baab' };
  const [first, second] = await Promise.all([resolver(request), resolver(request)]);

  assert.equal(calls.length, 2);
  const cssUrl = new URL(calls[0]);
  assert.equal(cssUrl.searchParams.get('family'), 'Poppins:ital,wght@1,600');
  assert.equal(cssUrl.searchParams.get('text'), 'ab');
  assert.equal(first.format, 'woff2');
  assert.deepEqual(new Uint8Array(first.data), bytes);
  assert.equal(second, first);
});

test('Google Fonts resolver returns null when a requested face is unavailable', async () => {
  const resolver = googleFonts({ fetch: async () => ({ ok: false }) });
  assert.equal(await resolver({ family: 'Missing Face', weight: 400, style: 'normal', text: 'Hi' }), null);
});

test('rejects excessive inputs, non-presentations and XML entity declarations', async () => {
  await assert.rejects(parsePptx(fixture(), { DOMParser, maxBytes: 10 }), /limit/);
  await assert.rejects(parsePptx(zipSync({}), { DOMParser }), /not an OOXML/);
  await assert.rejects(parsePptx(fixture({ 'ppt/slides/slide2.xml': '<!DOCTYPE x [<!ENTITY x "bad">]><x/>' }), { DOMParser }), /declarations/);
});

test('unsupported objects are surfaced instead of silently disappearing', async () => {
  const d = await parsePptx(fixture({ 'ppt/slides/slide2.xml': `<p:sld xmlns:p="${p}"><p:cSld><p:spTree><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Chart"/></p:nvGraphicFramePr></p:graphicFrame></p:spTree></p:cSld></p:sld>` }), { DOMParser });
  assert.equal(d.slides[0].elements[0].type, 'unsupported');
  assert.equal(d.warnings.length, 1);
});

test('Kroger integration: complete object inventory, media, table and stable IDs', { skip: !process.env.PPTX_TEST_FILE }, async () => {
  const bytes = await readFile(process.env.PPTX_TEST_FILE);
  const d = await parsePptx(bytes, { DOMParser });
  assert.equal(d.slides.length, 10);
  assert.deepEqual(d.slides.map(s => s.elements.length), [5, 18, 21, 19, 28, 28, 11, 20, 24, 9]);
  assert.equal(Object.keys(d.assets).length, 6);
  assert.equal(d.warnings.length, 0);
  assert.ok(d.slides.every(s => s.elements[0].locked));
  const table = d.slides[6].elements.find(e => e.type === 'table');
  assert.equal(table.cells.length, 18);
  assert.equal(table.cells[0].text.paragraphs[0].runs[0].text, 'Game');
  assert.ok(d.slides[9].elements.some(e => e.rotation === 90));
  const second = await parsePptx(bytes, { DOMParser });
  assert.deepEqual(d.slides.map(s => s.elements.map(e => e.id)), second.slides.map(s => s.elements.map(e => e.id)));
});
