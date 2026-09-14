# @butterstudio/pptx

A headless PPTX engine for JavaScript applications. Its public API is designed to parse, render, inspect, edit, and export PowerPoint files without imposing a UI framework or editor interface. React is not required.

> **Public preview:** the current release imports PPTX files into serializable Deck JSON and renders individually selectable SVG/DOM objects. Editing and PPTX export are planned capabilities; see [Current scope](#current-scope) before adopting it.

## Installation

Install the current preview channel:

```sh
npm install @butterstudio/pptx@next
```

The parser works in modern browsers and Node.js 20+. Rendering and font resolution require browser DOM APIs, including `FontFace`, `document.fonts`, SVG and `foreignObject`.

## Quick start

```js
import { getRequiredFonts, googleFonts, parsePptx, renderSlide, resolveFonts } from '@butterstudio/pptx';

const deck = await parsePptx(await file.arrayBuffer());
console.log(getRequiredFonts(deck));
// [{ family, weight, style, text, embedded, ... }]
const fonts = await resolveFonts(deck, {
  // Optional and API-key-free. Embedded fonts still take priority.
  resolveFont: googleFonts(),
});
const view = renderSlide(container, deck, 0, {
  fonts,
  onSelectionChange(selection) {
    // slideId, elementId, optional cellId, name, type, text
    console.log(selection);
  },
});
// When navigating away:
view.destroy();
fonts.destroy();
```

`parsePptx` does not flatten a slide into an image. Each supported PowerPoint object becomes a separate Deck JSON element and a separately selectable DOM/SVG object.

## API

### `parsePptx(input, options?)`

Parses a PPTX `ArrayBuffer`, `Uint8Array`, or `Blob` and returns a `Promise<DeckDocument>`. The returned document is JSON-serializable and contains slide geometry, stable source-derived object IDs, supported elements, embedded media, font requests and import warnings.

Options:

- `maxBytes`: maximum compressed input size. The default is 50 MB.
- `DOMParser`: optional DOM parser implementation for non-browser runtimes.

### `getRequiredFonts(deck)`

Returns each exact font face used by the deck, including `family`, `weight`, `style`, source name, optional embedded data and a `text` field containing each required character once.

### `googleFonts(options?)`

Creates an opt-in resolver for the public Google Fonts CSS API. Pass it to `resolveFonts` as `resolveFont`:

```js
const fonts = await resolveFonts(deck, {
  resolveFont: googleFonts({
    signal: abortController.signal,
    maxCharacters: 512,
    cacheSize: 64,
  }),
});
```

It requests the required family, weight, style and character subset, then returns font bytes to the package. It requires no Google API key. The application must allow browser connections to `fonts.googleapis.com` and `fonts.gstatic.com`.

Options:

- `signal`: aborts CSS and font downloads.
- `fetch`: custom Fetch API implementation, useful for testing or controlled networking.
- `maxCharacters`: maximum unique characters sent for one face; default `512`.
- `cacheSize`: maximum face requests cached by this resolver instance; default `64`.
- `cssEndpoint`: alternate Google Fonts-compatible CSS endpoint.

### `resolveFonts(deck, options?)`

Loads the faces needed to render a deck. Resolution happens in this order:

1. Recoverable font embedded in the PPTX.
2. Consumer `resolveFont` callback, such as `googleFonts()`.
3. Locally installed browser/system font.
4. Bundled default font.

The returned `ResolvedFontSession` contains a resolution record and warning for each fallback. Pass it to `renderSlide`, then call `destroy()` when the deck is closed to unregister loaded faces.

Useful options include `useSystemFonts: false`, `defaultFont`, `defaultFontFamily`, and a custom `resolveFont` callback that returns bytes, a `Blob`, data URL, URL, or `null`.

### `renderSlide(container, deck, slideIndex, options?)`

Renders one slide into an HTML container and returns a view handle with `clearSelection()` and `destroy()` methods.

```js
const view = renderSlide(container, deck, 0, {
  selectable: true,
  fonts,
  onSelectionChange(selection) {
    sendToChat({
      deckId,
      selectedObjects: selection,
    });
  },
});
```

Each selection record contains `slideId`, `elementId`, `type`, `name`, and extracted `text`. Table-cell selections additionally contain `cellId`, `row`, and `column`. The package reports selection context but does not send data to an AI service or network endpoint.

## Using Deck JSON in an editor

Deck JSON is intended to be the application-facing model instead of raw OOXML. Keep the original PPTX alongside it until round-trip export is implemented. A typical AI editor flow is:

```text
PPTX → parsePptx → Deck JSON → renderSlide → selection
     → application chat context → structured edit operation → Deck JSON
```

Mutation operations, history and PPTX writing are not part of the current preview, so consumers should not promise round-trip editing yet.

## Current scope

- OOXML slide relationships and order; deterministic IDs based on slide/object IDs.
- Deck coordinates in CSS pixels at 96 dpi. Layers follow source object order.
- Text paragraphs and runs, presentation defaults, basic master text styles for placeholders, local list styles, normalized font faces, alignment, margins, bullets, spacing and explicit shrink scale.
- Font resolution in this order: recoverable fonts embedded in the PPTX, a consumer callback, locally installed fonts through CSS `local()`, then the bundled AI Deck Default font.
- Rectangles, rounded rectangles, ellipses and lines with solid fills/strokes.
- Embedded PNG/JPEG/GIF/WebP media, shared asset references, cropping, rotation and flipping.
- Tables with individually selectable cells, explicit dimensions, padding, fills and borders.
- Single selection, Shift/Ctrl/Command multi-selection, Enter/Space selection and Escape clearing. Full-slide bottommost pictures are locked as backgrounds.
- Unsupported objects become labelled placeholders and import warnings.

This is a rendering and selection preview. There are **no edit operations, undo/redo, JSON persistence, or PPTX writer yet**. Rendering uses Deck JSON as its input; the application still stores PPTX as its persisted source and reloads JSON when the file revision changes. Keep the original PPTX for export. IDs are stable for the same source objects, not guaranteed across a skill regenerating slides.

Text uses SVG `foreignObject` with DOM runs. Face suffixes such as Regular, SemiBold, Bold and Italic are normalized into CSS family, weight and style fields. Deck JSON contains serializable font requests and recoverable embedded font payloads; the browser-only resolved font session stays separate and is passed to `renderSlide`. The package does not contact a font service unless the consumer explicitly passes a network resolver such as `googleFonts()`.

`getRequiredFonts(deck)` returns one entry for every face used by the presentation. Its `text` field contains each required character once, which lets a consumer request a safe subset without sending the deck's sentences to a font provider. `resolveFonts` uses this same inventory when it calls the consumer resolver.

`googleFonts()` is an optional built-in resolver for the public Google Fonts CSS API. It needs no API key, requests the exact family/weight/style, limits the request to unique characters used by that face, downloads the returned font bytes, and caches results for the resolver's lifetime. Configure `fonts.googleapis.com` and `fonts.gstatic.com` in the application's Content Security Policy if necessary. Applications that cannot send character subsets to Google should instead supply their own `resolveFont` callback.

System font loading is best effort because browsers can restrict local font access. Package consumers can disable it with `useSystemFonts: false`, replace the bundled fallback with `defaultFont`, or supply licensed fonts through `resolveFont`. The bundled fallback is the Latin subset of Poppins Regular under the SIL Open Font License. Missing faces use browser-synthesized weight/style and emit a resolution warning.

PowerPoint often stores embedded fonts as EOT. Bare OpenType, WOFF/WOFF2, obfuscated OpenType and EOT payloads are supported, including MicroType Express compression. Invalid or browser-incompatible embedded data produces an import warning and continues through the resolver/system/default chain. Natural font metrics, wrapping and automatic box growth remain approximate; `spAutoFit` is retained but this release keeps source geometry rather than implementing PowerPoint's layout engine.

Not implemented: charts, SmartArt, groups, custom geometry, OLE, video/audio, complex effects, master/layout graphics, complete placeholder inheritance, merged tables, theme table styles, and export/round-tripping. Notes and unsupported XML are not retained in JSON in this version. Warnings cover unsupported objects and selected visual features, not every possible OOXML extension.

Inputs are limited to 50 MB compressed / 200 MB expanded / 10,000 ZIP entries. External relationships are not fetched, XML entities are rejected, and document text is assigned through `textContent`. Parsing currently runs on the main thread; a worker is a future performance improvement for larger decks.

## Security and privacy

- The parser does not fetch external OOXML relationships.
- `googleFonts()` is opt-in. When enabled, unique characters used by a font face are sent to Google as the CSS API `text` parameter; complete sentences are not sent in their original order.
- A custom font resolver can keep all font traffic on infrastructure controlled by the application.
- Inspect `npm pack --dry-run` before publishing or vendoring a modified build.

## Development

```sh
npm install
npm test
# Optional local integration fixture; customer decks are not packaged:
PPTX_TEST_FILE=/absolute/path/to/kroger.pptx npm test
npm pack --dry-run
```

The package is prepared for local development, **not published**. It will be published as `@butterstudio/pptx`; choose a license before public publication. No customer deck, media or application secrets are included in the package allowlist.
