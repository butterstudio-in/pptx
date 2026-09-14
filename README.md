# @butterstudio/pptx

A headless PPTX engine for JavaScript applications. Its public API is designed to parse, render, inspect, edit, and export PowerPoint files without imposing a UI framework or editor interface. React is not required.

Today, the package imports PPTX files into serializable Deck JSON and renders individually selectable SVG/DOM objects. Editing and PPTX export are planned capabilities; see [Current scope](#current-scope) for what is available in this release.

```js
import { getRequiredFonts, parsePptx, renderSlide, resolveFonts } from '@butterstudio/pptx';

const deck = await parsePptx(await file.arrayBuffer());
console.log(getRequiredFonts(deck));
// [{ family, weight, style, text, embedded, ... }]
const fonts = await resolveFonts(deck, {
  // Optional. Return bytes, a Blob, data URL, or URL for fonts your app owns.
  async resolveFont({ family, weight, style }) {
    return myFontStore.find({ family, weight, style });
  },
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

Text uses SVG `foreignObject` with DOM runs. Face suffixes such as Regular, SemiBold, Bold and Italic are normalized into CSS family, weight and style fields. Deck JSON contains serializable font requests and recoverable embedded font payloads; the browser-only resolved font session stays separate and is passed to `renderSlide`. The package never contacts a font service.

`getRequiredFonts(deck)` returns one entry for every face used by the presentation. Its `text` field contains each required character once, which lets a consumer request a safe subset without sending the deck's sentences to a font provider. `resolveFonts` uses this same inventory when it calls the consumer resolver.

System font loading is best effort because browsers can restrict local font access. Package consumers can disable it with `useSystemFonts: false`, replace the bundled fallback with `defaultFont`, or supply licensed fonts through `resolveFont`. The bundled fallback is the Latin subset of Poppins Regular under the SIL Open Font License. Missing faces use browser-synthesized weight/style and emit a resolution warning.

PowerPoint often stores embedded fonts as EOT. Bare OpenType, WOFF/WOFF2, obfuscated OpenType and EOT payloads are supported, including MicroType Express compression. Invalid or browser-incompatible embedded data produces an import warning and continues through the resolver/system/default chain. Natural font metrics, wrapping and automatic box growth remain approximate; `spAutoFit` is retained but this release keeps source geometry rather than implementing PowerPoint's layout engine.

Not implemented: charts, SmartArt, groups, custom geometry, OLE, video/audio, complex effects, master/layout graphics, complete placeholder inheritance, merged tables, theme table styles, and export/round-tripping. Notes and unsupported XML are not retained in JSON in this version. Warnings cover unsupported objects and selected visual features, not every possible OOXML extension.

Inputs are limited to 50 MB compressed / 200 MB expanded / 10,000 ZIP entries. External relationships are not fetched, XML entities are rejected, and document text is assigned through `textContent`. Parsing currently runs on the main thread; a worker is a future performance improvement for larger decks.

## Development

```sh
npm install
npm test
# Optional local integration fixture; customer decks are not packaged:
PPTX_TEST_FILE=/absolute/path/to/kroger.pptx npm test
npm pack --dry-run
```

The package is prepared for local development, **not published**. It will be published as `@butterstudio/pptx`; choose a license before public publication. No customer deck, media or application secrets are included in the package allowlist.
