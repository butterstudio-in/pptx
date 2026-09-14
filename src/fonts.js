import { DEFAULT_FONT_DATA_URL } from './default-font.js';

const WEIGHTS = new Map([
  ['thin', 100], ['extralight', 200], ['ultralight', 200], ['light', 300],
  ['regular', 400], ['normal', 400], ['book', 400], ['medium', 500],
  ['semibold', 600], ['demibold', 600], ['bold', 700],
  ['extrabold', 800], ['ultrabold', 800], ['black', 900], ['heavy', 900],
]);

const cleanName = value => String(value || 'Arial').replace(/["'\\]/g, '').trim() || 'Arial';

export function fontKey(value) {
  return `${value.family.toLowerCase()}|${value.weight}|${value.style}`;
}

/** Normalize names such as "Poppins SemiBold" into CSS family/face fields. */
export function normalizeTypeface(typeface, bold = false, italic = false) {
  const sourceTypeface = cleanName(typeface);
  let family = sourceTypeface;
  let weight = bold ? 700 : 400;
  let style = italic ? 'italic' : 'normal';
  const tokens = family.split(/\s+/);
  let changed = true;
  while (tokens.length > 1 && changed) {
    changed = false;
    const token = tokens.at(-1).toLowerCase().replace(/[-_]/g, '');
    if (token === 'italic' || token === 'oblique') {
      style = 'italic';
      tokens.pop();
      changed = true;
    } else if (WEIGHTS.has(token)) {
      weight = WEIGHTS.get(token);
      tokens.pop();
      changed = true;
    }
  }
  family = tokens.join(' ') || sourceTypeface;
  return { family, weight, style, sourceTypeface };
}

/** Return the exact font faces and unique characters used by a Deck document. */
export function getRequiredFonts(deck) {
  const required = new Map((deck.fonts ?? []).map(font => [font.id ?? fontKey(font), { ...font, characters: new Set() }]));
  const collectBody = body => {
    for (const paragraph of body?.paragraphs ?? []) {
      for (const run of paragraph.runs ?? []) {
        const key = run.style.fontId;
        let font = required.get(key);
        if (!font) {
          const normalized = normalizeTypeface(run.style.sourceTypeface ?? run.style.fontFamily, run.style.bold, run.style.italic);
          font = { id: key ?? fontKey(normalized), ...normalized, characters: new Set() };
          required.set(font.id, font);
        }
        for (const character of run.text ?? '') font.characters.add(character);
      }
    }
  };
  for (const slide of deck.slides ?? []) {
    for (const element of slide.elements ?? []) {
      collectBody(element.text);
      for (const cell of element.cells ?? []) collectBody(cell.text);
    }
  }
  return [...required.values()].map(({ characters, ...font }) => ({ ...font, text: [...characters].join('') }));
}

const hash = value => {
  let result = 2166136261;
  for (const char of value) result = Math.imul(result ^ char.charCodeAt(0), 16777619);
  return (result >>> 0).toString(36);
};

const cssString = value => JSON.stringify(String(value));

async function faceSource(source) {
  if (source.data != null) {
    if (typeof source.data.arrayBuffer === 'function') return source.data.arrayBuffer();
    if (ArrayBuffer.isView(source.data)) return source.data.buffer.slice(source.data.byteOffset, source.data.byteOffset + source.data.byteLength);
    return source.data;
  }
  const url = source.dataUrl ?? source.url;
  if (!url) throw new Error('A font source must contain data, dataUrl, or url.');
  const format = source.format ? ` format(${cssString(source.format)})` : '';
  return `url(${cssString(url)})${format}`;
}

async function loadFace(FontFaceClass, fontSet, alias, request, source) {
  const face = new FontFaceClass(alias, await faceSource(source), {
    style: request.style,
    weight: String(request.weight),
  });
  await face.load();
  fontSet.add(face);
  return face;
}

function embeddedSource(request) {
  return request.embedded?.dataUrl ? request.embedded : null;
}

function resolverSource(value) {
  if (!value) return null;
  return typeof value === 'string' ? { url: value } : value;
}

/**
 * Load all deck fonts in priority order: PPTX, consumer resolver, system, default.
 * The returned session is intentionally separate from serializable Deck JSON.
 */
export async function resolveFonts(deck, options = {}) {
  const FontFaceClass = options.FontFace ?? globalThis.FontFace;
  const fontSet = options.fontSet ?? globalThis.document?.fonts;
  if (!FontFaceClass || !fontSet) throw new Error('Font resolution requires the browser FontFace API and a FontFaceSet.');
  const resolutions = {};
  const warnings = [];
  const loadedFaces = [];
  let defaultFace;
  let defaultPromise;
  const defaultAlias = options.defaultFontFamily ?? 'PPTX Fallback';

  async function loadDefault() {
    if (!defaultPromise) {
      const source = options.defaultFont ?? { dataUrl: DEFAULT_FONT_DATA_URL, format: 'woff2' };
      defaultPromise = loadFace(FontFaceClass, fontSet, defaultAlias, { weight: 400, style: 'normal' }, source).then(face => {
        defaultFace = face;
        loadedFaces.push(face);
        return face;
      });
    }
    return defaultPromise;
  }

  await Promise.all(getRequiredFonts(deck).map(async request => {
    const key = request.id ?? fontKey(request);
    const alias = `pptx-${hash(key)}`;
    let face;
    let source;
    let detail;
    const failures = [];

    const embedded = embeddedSource(request);
    if (embedded) {
      try {
        face = await loadFace(FontFaceClass, fontSet, alias, request, embedded);
        source = 'embedded';
        detail = embedded.part;
      } catch (error) {
        failures.push(`embedded font: ${error instanceof Error ? error.message : String(error)}`);
      }
    }

    if (!face && options.resolveFont) {
      try {
        const supplied = resolverSource(await options.resolveFont({ ...request }));
        if (supplied) {
          face = await loadFace(FontFaceClass, fontSet, alias, request, supplied);
          source = 'resolver';
          detail = supplied.url;
        }
      } catch (error) {
        failures.push(`font resolver: ${error instanceof Error ? error.message : String(error)}`);
      }
    }

    if (!face && options.useSystemFonts !== false) {
      const candidates = [...new Set([request.postscriptName, request.sourceTypeface, request.family].filter(Boolean))];
      for (const candidate of candidates) {
        try {
          const localFace = new FontFaceClass(alias, `local(${cssString(candidate)})`, { style: request.style, weight: String(request.weight) });
          await localFace.load();
          fontSet.add(localFace);
          face = localFace;
        } catch { /* Try the next local name. */ }
        if (face) {
          loadedFaces.push(face);
          source = 'system';
          detail = candidate;
          break;
        }
      }
    }

    if (!face) {
      await loadDefault();
      source = 'default';
      detail = defaultAlias;
      warnings.push({
        fontId: key,
        family: request.family,
        message: `${request.sourceTypeface || request.family} was unavailable; ${defaultAlias} is being used.`,
        failures,
      });
    } else if (source !== 'system') {
      loadedFaces.push(face);
    }

    resolutions[key] = {
      fontId: key,
      requestedFamily: request.family,
      resolvedFamily: source === 'default' ? defaultAlias : alias,
      weight: request.weight,
      style: request.style,
      source,
      detail,
    };
  }));

  return {
    resolutions,
    warnings,
    familyFor(style) {
      return resolutions[style.fontId]?.resolvedFamily ?? style.fontFamily;
    },
    destroy() {
      for (const face of loadedFaces) fontSet.delete?.(face);
      loadedFaces.length = 0;
    },
  };
}
