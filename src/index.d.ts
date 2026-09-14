export type FontStyle = 'normal' | 'italic';
export type EmbeddedFont = { part: string; format: string; mimeType: string; dataUrl: string };
export type DeckFont = {
  id: string;
  family: string;
  weight: number;
  style: FontStyle;
  sourceTypeface: string;
  postscriptName?: string;
  embedded?: EmbeddedFont;
};
export type RequiredFont = DeckFont & { text: string };
export type TextStyle = {
  fontFamily: string;
  sourceTypeface: string;
  fontId: string;
  fontWeight: number;
  fontSize: number;
  color: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  letterSpacing: number;
};
export type Paragraph = { align: string; marginLeft: number; marginRight: number; indent: number; lineHeight: number; lineHeightPx?: number; before: number; after: number; bullet: string | null; runs: { text: string; style: TextStyle }[] };
export type TextBody = { anchor: string; wrap: boolean; autofit: 'shrink' | 'grow' | 'none'; fontScale: number; inset: { left: number; right: number; top: number; bottom: number }; paragraphs: Paragraph[] };
export type Frame = { x: number; y: number; width: number; height: number; rotation: number; flipH: boolean; flipV: boolean };
export type BaseElement = Frame & { id: string; sourceId: string; name: string; locked: boolean; role?: 'background' };
export type ShapeElement = BaseElement & { type: 'text' | 'shape'; geometry: 'rect' | 'roundRect' | 'ellipse' | 'line'; fill: string; stroke: string; strokeWidth: number; radius: number; text?: TextBody };
export type ImageElement = BaseElement & { type: 'image'; assetId: string; crop: { l: number; t: number; r: number; b: number } };
export type TableCell = { id: string; row: number; column: number; x: number; y: number; width: number; height: number; fill: string; text: TextBody; borders: Record<'L' | 'R' | 'T' | 'B', { color: string; width: number }> };
export type TableElement = BaseElement & { type: 'table'; columns: number[]; rows: number[]; cells: TableCell[] };
export type DeckElement = ShapeElement | ImageElement | TableElement | (BaseElement & { type: 'unsupported'; reason: string });
export type DeckSlide = { id: string; sourcePart: string; background: string; elements: DeckElement[] };
export type DeckWarning = { slideId?: string; elementId?: string; message: string };
export type DeckDocument = { version: 2; width: number; height: number; slides: DeckSlide[]; assets: Record<string, { id: string; mimeType: string; dataUrl: string }>; fonts: DeckFont[]; warnings: DeckWarning[] };
export type Selection = { slideId: string; elementId: string; cellId?: string; row?: number; column?: number; type: string; name: string; text: string };

export type FontSource = {
  data?: ArrayBuffer | ArrayBufferView | Blob;
  dataUrl?: string;
  url?: string;
  format?: 'truetype' | 'opentype' | 'woff' | 'woff2' | string;
};
export type FontResolution = {
  fontId: string;
  requestedFamily: string;
  resolvedFamily: string;
  weight: number;
  style: FontStyle;
  source: 'embedded' | 'resolver' | 'system' | 'default';
  detail?: string;
};
export type FontWarning = { fontId: string; family: string; message: string; failures: string[] };
export type ResolvedFontSession = {
  resolutions: Record<string, FontResolution>;
  warnings: FontWarning[];
  familyFor(style: Pick<TextStyle, 'fontId' | 'fontFamily'>): string;
  destroy(): void;
};

export function fontKey(font: Pick<DeckFont, 'family' | 'weight' | 'style'>): string;
export function getRequiredFonts(deck: DeckDocument): RequiredFont[];
export function normalizeTypeface(typeface: string, bold?: boolean, italic?: boolean): Omit<DeckFont, 'id' | 'embedded'>;
export function parsePptx(input: ArrayBuffer | Uint8Array | Blob, options?: { maxBytes?: number; DOMParser?: typeof DOMParser }): Promise<DeckDocument>;
export function resolveFonts(deck: DeckDocument, options?: {
  resolveFont?: (font: RequiredFont) => FontSource | string | null | undefined | Promise<FontSource | string | null | undefined>;
  useSystemFonts?: boolean;
  defaultFontFamily?: string;
  defaultFont?: FontSource;
  FontFace?: typeof FontFace;
  fontSet?: FontFaceSet;
}): Promise<ResolvedFontSession>;
export function renderSlide(container: HTMLElement, deck: DeckDocument, slideIndex: number, options?: { selectable?: boolean; fonts?: ResolvedFontSession; onSelectionChange?: (selection: Selection[]) => void }): { warnings: { elementId?: string; message: string }[]; clearSelection(): void; destroy(): void };
