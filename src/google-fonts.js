const DEFAULT_CSS_ENDPOINT = 'https://fonts.googleapis.com/css2';
const GOOGLE_FONT_URL = /^https:\/\/fonts\.gstatic\.com\//;

const abortError = error => error?.name === 'AbortError';

function uniqueCharacters(text, limit) {
  return [...new Set(Array.from(text ?? ''))].sort().slice(0, limit).join('');
}

function usableFontSource(css) {
  return [...css.matchAll(/src:\s*url\(([^)]+)\)\s*format\(["']?([^"')]+)["']?\)/g)]
    .map(match => ({ url: match[1].replace(/["']/g, ''), format: match[2] }))
    .find(candidate => GOOGLE_FONT_URL.test(candidate.url));
}

/**
 * Create an opt-in resolver for the public Google Fonts CSS API.
 * Pass the returned function to resolveFonts as `resolveFont`.
 */
export function googleFonts(options = {}) {
  const fetcher = options.fetch ?? globalThis.fetch;
  if (typeof fetcher !== 'function') throw new Error('googleFonts requires a Fetch API implementation.');
  const cssEndpoint = options.cssEndpoint ?? DEFAULT_CSS_ENDPOINT;
  const maxCharacters = Math.max(1, options.maxCharacters ?? 512);
  const cacheSize = Math.max(1, options.cacheSize ?? 64);
  const cache = new Map();

  return function resolveGoogleFont(font) {
    const family = String(font.family ?? '').trim();
    if (!family || family.length > 100) return Promise.resolve(null);
    const weight = Number.isInteger(font.weight) && font.weight >= 100 && font.weight <= 900 ? font.weight : 400;
    const style = font.style === 'italic' ? 'italic' : 'normal';
    const characters = uniqueCharacters(font.text, maxCharacters);
    const key = `${family}|${weight}|${style}|${characters}`;
    const cached = cache.get(key);
    if (cached) return cached;

    const resolution = (async () => {
      const familySpec = style === 'italic' ? `${family}:ital,wght@1,${weight}` : `${family}:wght@${weight}`;
      const cssUrl = new URL(cssEndpoint);
      cssUrl.searchParams.set('family', familySpec);
      cssUrl.searchParams.set('display', 'block');
      if (characters) cssUrl.searchParams.set('text', characters);

      const cssResponse = await fetcher(cssUrl, { signal: options.signal });
      if (!cssResponse.ok) return null;
      const source = usableFontSource(await cssResponse.text());
      if (!source) return null;

      const fontResponse = await fetcher(source.url, { signal: options.signal });
      if (!fontResponse.ok) return null;
      return { data: await fontResponse.arrayBuffer(), format: source.format || 'woff2' };
    })().catch(error => {
      cache.delete(key);
      if (abortError(error)) throw error;
      return null;
    });

    if (cache.size >= cacheSize) cache.delete(cache.keys().next().value);
    cache.set(key, resolution);
    return resolution;
  };
}
