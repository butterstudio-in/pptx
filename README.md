# @butterstudio/pptx

A headless PPTX import, rendering, and targeted editing engine for JavaScript applications. It parses PowerPoint files into serializable Deck JSON, renders selectable slide objects, applies deterministic edit operations, and writes the changes back to PPTX without imposing a UI framework. React is not required.

> **Public preview:** targeted text, style, geometry, alignment, and image replacement operations are available. This is not yet a complete PowerPoint implementation; see [Current scope](#current-scope) before adopting it.

## Installation

Install the current preview channel:

```sh
npm install @butterstudio/pptx@next
```

The parser works in modern browsers and Node.js 20+. Rendering and font resolution require browser DOM APIs, including `FontFace`, `document.fonts`, SVG and `foreignObject`.

For parsing or editing in Node.js, install a DOM implementation in your application:

```sh
npm install @xmldom/xmldom
```

```js
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { openPptx } from '@butterstudio/pptx';

const session = await openPptx(pptxBytes, { DOMParser, XMLSerializer });
const exported = await session.exportPptx();
// Write exported.data with your server runtime's file or object-storage API.
```

`renderSlide` and `resolveFonts` are browser-only. A server can parse, edit, and export without either one.

## Quick start

```js
import { getRequiredFonts, googleFonts, openPptx, renderSlide, resolveFonts } from '@butterstudio/pptx';

const session = await openPptx(await file.arrayBuffer());
const deck = session.deck;
console.log(getRequiredFonts(deck));
// [{ family, weight, style, text, embedded, ... }]
const fonts = await resolveFonts(deck, {
  // Optional and API-key-free. Embedded fonts still take priority.
  resolveFont: googleFonts(),
});
const view = renderSlide(container, deck, 0, {
  fonts,
  onSelectionChange(selection) {
    // Save these IDs to construct an edit operation.
    // A table-cell selection also includes cellId.
    console.log(selection);
  },
});
// When navigating away:
view.destroy();
fonts.destroy();
```

`parsePptx` does not flatten a slide into an image. Each supported PowerPoint object becomes a separate Deck JSON element and a separately selectable DOM/SVG object.

For an editing workflow, keep the `session` open, render `session.deck`, and use the IDs from the deck or a selection callback. Do not construct IDs from OOXML paths: the exact IDs are source-derived and look like `slide-99` and `slide-99/object-7`.

## API

### `parsePptx(input, options?)`

Parses a PPTX `ArrayBuffer`, `Uint8Array`, or `Blob` and returns a `Promise<DeckDocument>`. The returned document is JSON-serializable and contains slide geometry, stable source-derived object IDs, supported elements, embedded media, font requests and import warnings.

Options:

- `maxBytes`: maximum compressed input size. The default is 50 MB.
- `DOMParser`: optional DOM parser implementation for non-browser runtimes.

### `openPptx(input, options?)`

Opens the PPTX as a mutable editing session while retaining the original OOXML package. In Node.js, pass compatible `DOMParser` and `XMLSerializer` implementations such as those from `@xmldom/xmldom`. A backend can also pass the last saved `DeckDocument` as `options.deck`; this preserves application metadata such as asset provenance while applying changes to its matching PPTX revision.

```js
const session = await openPptx(pptxBytes, { DOMParser, XMLSerializer });
const slide = session.deck.slides[0];
const title = slide.elements.find(element => element.type === 'text');
if (!title) throw new Error('This slide has no text element.');

session.applyOperations([
  {
    type: 'setText',
    slideId: slide.id,
    elementId: title.id,
    text: 'A shorter title',
  },
]);

const result = await session.exportPptx();
// result.data              Uint8Array containing the updated .pptx
// result.deck              updated serializable Deck JSON
// result.inverseOperations operations suitable for application-level undo
// result.changedParts      OOXML parts changed during export
```

Supported operations are `setText`, `setTextStyle`, `setFrame`, `setShapeStyle`, `alignElements`, and `replaceImage`. Operations target the stable `slideId`, `elementId`, and optional table `cellId` returned by the parser/selection API. `applyOperations` updates Deck JSON atomically: if any operation is invalid, the session keeps its previous deck.

For `replaceImage`, provide an asset containing `id`, `mimeType`, and a base64 `dataUrl`. Extra serializable metadata such as a public `url`, content hash, dimensions, or AI-generation provenance is retained in Deck JSON.

### Edit operations

Pass one to 500 operations to `session.applyOperations(operations)`. An operation batch is atomic: if one operation is invalid, none of that batch is applied. `applyOperations()` returns inverses for that batch; `exportPptx()` returns inverses for every batch applied in the session, in undo order.

All lengths and positions use CSS pixels at 96 dpi. Obtain `slideId`, `elementId`, and, for a table cell, `cellId` from `session.deck` or `onSelectionChange`.

```js
const slide = session.deck.slides[0];
const textElement = slide.elements.find(element =>
  !element.locked && (element.type === 'text' || (element.type === 'shape' && element.text))
);
const shapeElement = slide.elements.find(element =>
  !element.locked && (element.type === 'text' || element.type === 'shape')
);
const [first, second] = slide.elements.filter(element => !element.locked);
if (!textElement || !shapeElement || !first || !second) {
  throw new Error('Choose compatible editable elements from session.deck.');
}

session.applyOperations([
  // Replaces all text in a text element with one string. For a table, add cellId.
  { type: 'setText', slideId: slide.id, elementId: textElement.id, text: 'New title' },

  // Applies these properties to every text run in the selected object or cell.
  {
    type: 'setTextStyle', slideId: slide.id, elementId: textElement.id,
    style: { fontFamily: 'Poppins', fontSize: 32, color: '#1d4ed8', bold: true },
  },

  // x, y, width, height, and rotation are numbers; flipH and flipV are booleans.
  { type: 'setFrame', slideId: slide.id, elementId: textElement.id, frame: { x: 64, y: 48, width: 560 } },

  // Applies to text and basic shape elements.
  { type: 'setShapeStyle', slideId: slide.id, elementId: shapeElement.id, style: { fill: '#eff6ff', stroke: '#2563eb', strokeWidth: 2 } },

  // Supply alignment, distribution, or both. At least two element IDs are required.
  { type: 'alignElements', slideId: slide.id, elementIds: [first.id, second.id], alignment: 'left', distribution: 'vertical' },
]);
```

For a table cell, target the containing table and add its `cellId`:

```js
const table = slide.elements.find(element => element.type === 'table');
if (!table) throw new Error('This slide has no table.');
const cell = table.cells[0];
session.applyOperations([{
  type: 'setText', slideId: slide.id, elementId: table.id, cellId: cell.id, text: 'Updated cell',
}]);
```

`setText` intentionally replaces an object's text with one string; it does not preserve individual runs or paragraphs. `setTextStyle` updates every existing run. Table cells cannot be moved independently. Locked elements, including detected full-slide background images, should not be treated as editable.

To replace an image, use one of the four supported raster MIME types and a base64 data URL:

```js
const imageElement = slide.elements.find(element => element.type === 'image');
if (!imageElement) throw new Error('This slide has no image.');

session.applyOperations([{
  type: 'replaceImage',
  slideId: slide.id,
  elementId: imageElement.id,
  asset: {
    id: 'hero-v2',
    mimeType: 'image/png', // image/png, image/jpeg, image/gif, or image/webp
    dataUrl: 'data:image/png;base64,...',
    provenance: { kind: 'ai-generated', provider: 'Example provider' },
  },
}]);
```

### Exporting changes

`session.revision` identifies the original PPTX bytes for the lifetime of the session. Store it with your application revision if you need to ensure a saved Deck JSON is applied to its matching source PPTX.

```js
const exported = await session.exportPptx();

// Browser: offer the updated file for download.
const blob = new Blob([exported.data], {
  type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
});
const url = URL.createObjectURL(blob);
// Attach `url` to a temporary <a download="edited.pptx">, click it, then revoke it.

// `exported.deck`, `exported.inverseOperations`, and `exported.changedParts`
// are serializable metadata for your application's revision/undo store.
```

`exportPptx()` writes the supported edits back to a new `Uint8Array`; it does not modify the uploaded file. Keep the original PPTX and its matching Deck JSON as a recoverable application revision, especially while using this preview release.

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

Each selection record contains `slideId`, `elementId`, source ID, type, name, extracted text, and frame geometry. Table-cell selections additionally contain `cellId`, `row`, and `column`. Image selections contain asset metadata with the large `dataUrl` omitted, so they can be attached safely to an application chat request. The package reports selection context but does not send it to an AI service or network endpoint.

## Using Deck JSON in an editor

Deck JSON is intended to be the application-facing model instead of raw OOXML. Keep the original PPTX alongside its matching Deck JSON for revision recovery and export. A typical AI editor flow is:

```text
PPTX → openPptx → Deck JSON → renderSlide → selection
     → application chat context → structured edit operations
     → applyOperations → updated Deck JSON → exportPptx
```

The package returns inverse operations but deliberately does not own storage, revision history, chat, or network calls. Applications should persist the original PPTX, Deck JSON, operations, and exported PPTX together as one revision.

## Current scope

- OOXML slide relationships and order; deterministic IDs based on slide/object IDs.
- Deck coordinates in CSS pixels at 96 dpi. Layers follow source object order.
- Text paragraphs and runs, presentation defaults, basic master text styles for placeholders, local list styles, normalized font faces, alignment, margins, bullets, spacing and explicit shrink scale.
- Font resolution in this order: recoverable fonts embedded in the PPTX, a consumer callback, locally installed fonts through CSS `local()`, then the bundled fallback font.
- Rectangles, rounded rectangles, ellipses and lines with solid fills/strokes.
- Embedded PNG/JPEG/GIF/WebP media, shared asset references, cropping, rotation and flipping.
- Tables with individually selectable cells, explicit dimensions, padding, fills and borders.
- Single selection, Shift/Ctrl/Command multi-selection, Enter/Space selection and Escape clearing. Full-slide bottommost pictures are locked as backgrounds.
- Unsupported objects become labelled placeholders and import warnings.

The editing preview currently round-trips targeted text content/style, basic shape styling, element frames, alignment/distribution, and raster image replacement. It does not provide a history store or user interface; applications own revisions and may use the returned inverse operations for undo. IDs are stable for the same source objects but are not guaranteed to survive regeneration of the source presentation.

Text uses SVG `foreignObject` with DOM runs. Face suffixes such as Regular, SemiBold, Bold and Italic are normalized into CSS family, weight and style fields. Deck JSON contains serializable font requests and recoverable embedded font payloads; the browser-only resolved font session stays separate and is passed to `renderSlide`. The package does not contact a font service unless the consumer explicitly passes a network resolver such as `googleFonts()`.

`getRequiredFonts(deck)` returns one entry for every face used by the presentation. Its `text` field contains each required character once, which lets a consumer request a safe subset without sending the deck's sentences to a font provider. `resolveFonts` uses this same inventory when it calls the consumer resolver.

`googleFonts()` is an optional built-in resolver for the public Google Fonts CSS API. It needs no API key, requests the exact family/weight/style, limits the request to unique characters used by that face, downloads the returned font bytes, and caches results for the resolver's lifetime. Configure `fonts.googleapis.com` and `fonts.gstatic.com` in the application's Content Security Policy if necessary. Applications that cannot send character subsets to Google should instead supply their own `resolveFont` callback.

System font loading is best effort because browsers can restrict local font access. Package consumers can disable it with `useSystemFonts: false`, replace the bundled fallback with `defaultFont`, or supply licensed fonts through `resolveFont`. The bundled fallback is the Latin subset of Poppins Regular under the SIL Open Font License. Missing faces use browser-synthesized weight/style and emit a resolution warning.

PowerPoint often stores embedded fonts as EOT. Bare OpenType, WOFF/WOFF2, obfuscated OpenType and EOT payloads are supported, including MicroType Express compression. Invalid or browser-incompatible embedded data produces an import warning and continues through the resolver/system/default chain. Natural font metrics, wrapping and automatic box growth remain approximate; `spAutoFit` is retained but this release keeps source geometry rather than implementing PowerPoint's layout engine.

Not implemented for editing: adding/removing/reordering slides or objects, charts, SmartArt, groups, custom geometry, OLE, video/audio, animations, complex effects, master/layout graphics, merged tables, and theme table styles. Unsupported OOXML is retained in the source package when other objects are edited, but it is not represented completely in Deck JSON. Warnings cover unsupported objects and selected visual features, not every possible OOXML extension.

Inputs are limited to 50 MB compressed / 200 MB expanded / 10,000 ZIP entries. External relationships are not fetched, XML entities are rejected, and document text is assigned through `textContent`. Parsing currently runs on the main thread; a worker is a future performance improvement for larger decks.

## Security and privacy

- The parser does not fetch external OOXML relationships.
- `googleFonts()` is opt-in. When enabled, unique characters used by a font face are sent to Google as the CSS API `text` parameter; complete sentences are not sent in their original order.
- A custom font resolver can keep all font traffic on infrastructure controlled by the application.

## Contributing

```sh
npm install
npm test
```
