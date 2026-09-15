const SVG = 'http://www.w3.org/2000/svg';
const HTML = 'http://www.w3.org/1999/xhtml';
const svg = (name, attrs = {}) => {
  const el = document.createElementNS(SVG, name);
  for (const [key, value] of Object.entries(attrs)) el.setAttribute(key, String(value));
  return el;
};
const html = name => document.createElementNS(HTML, name);
const css = (node, rules) => Object.assign(node.style, rules);
const textContent = element => element.text?.paragraphs.map(p => p.runs.map(r => r.text).join('')).join('\n') ?? '';

function drawText(parent, body, width, height, fonts) {
  if (!body) return;
  const fo = svg('foreignObject', { width: Math.max(width, 1), height: Math.max(height, 1), overflow: 'visible' });
  const box = html('div');
  const inset = body.inset;
  css(box, {
    width: '100%', height: '100%', boxSizing: 'border-box', display: 'flex', flexDirection: 'column',
    justifyContent: body.anchor === 'ctr' ? 'center' : body.anchor === 'b' ? 'flex-end' : 'flex-start',
    padding: `${inset.top}px ${inset.right}px ${inset.bottom}px ${inset.left}px`,
    overflow: 'visible', pointerEvents: 'none', color: '#000', textAlign: 'left',
  });
  const content = html('div');
  css(content, { flexShrink: '0', minWidth: '0', width: '100%' });
  for (const paragraph of body.paragraphs) {
    const p = html('div');
    const firstStyle = paragraph.runs.find(r => r.text)?.style ?? paragraph.runs[0]?.style;
    const fontSize = (firstStyle?.fontSize ?? 24) * body.fontScale;
    css(p, {
      margin: `${paragraph.before}px ${paragraph.marginRight}px ${paragraph.after}px ${paragraph.marginLeft}px`,
      padding: '0', fontSize: `${fontSize}px`,
      // OOXML percentage spacing scales the natural line height, not the em size.
      lineHeight: paragraph.lineHeightPx ? `${paragraph.lineHeightPx}px` : String(paragraph.lineHeight * 1.2),
      textAlign: ({ l: 'left', r: 'right', ctr: 'center', just: 'justify', dist: 'justify' })[paragraph.align] ?? 'left',
      whiteSpace: body.wrap ? 'pre-wrap' : 'pre', overflowWrap: 'break-word',
      textIndent: `${paragraph.indent}px`, minHeight: `${fontSize * 1.2}px`,
    });
    if (paragraph.bullet) {
      const bullet = html('span');
      bullet.textContent = `${paragraph.bullet}\u00a0 `;
      const bulletFamily = firstStyle ? fonts?.familyFor(firstStyle) ?? firstStyle.fontFamily : 'sans-serif';
      css(bullet, { color: firstStyle?.color, fontFamily: `"${bulletFamily}", sans-serif` });
      p.append(bullet);
    }
    for (const run of paragraph.runs) {
      const span = html('span');
      const style = run.style;
      const family = (fonts?.familyFor(style) ?? style.fontFamily).replaceAll('"', '');
      css(span, { fontFamily: `"${family}", sans-serif`,
        fontSize: `${style.fontSize * body.fontScale}px`, color: style.color,
        fontWeight: String(style.fontWeight ?? (style.bold ? 700 : 400)), fontStyle: style.italic ? 'italic' : 'normal',
        textDecoration: style.underline ? 'underline' : 'none', letterSpacing: `${style.letterSpacing}px` });
      span.textContent = run.text;
      p.append(span);
    }
    content.append(p);
  }
  box.append(content);
  fo.append(box);
  parent.append(fo);
}

/** Render one Deck JSON slide. No XML, scripts, or HTML from the file are injected. */
export function renderSlide(container, deck, slideIndex, options = {}) {
  const slide = deck.slides[slideIndex];
  if (!slide) throw new Error(`Slide ${slideIndex + 1} does not exist.`);
  const root = svg('svg', { viewBox: `0 0 ${deck.width} ${deck.height}`, width: '100%', height: '100%', role: 'group', 'aria-label': `Slide ${slideIndex + 1}`, 'data-pptx-slide': slide.id });
  css(root, { display: 'block', overflow: 'hidden', background: slide.background, isolation: 'isolate' });
  const layer = svg('g');
  const overlay = svg('g', { 'pointer-events': 'none', 'data-selection-overlay': '' });
  const targets = new Map();
  let selected = new Set();
  const notify = () => {
    overlay.replaceChildren();
    for (const id of selected) {
      const target = targets.get(id);
      if (!target) continue;
      const outline = svg('g', { transform: target.transform });
      outline.append(svg('rect', { x: target.cell?.x ?? 0, y: target.cell?.y ?? 0,
        width: target.cell?.width ?? target.element.width, height: target.cell?.height ?? target.element.height,
        fill: 'none', stroke: '#2563eb', 'stroke-width': 2, 'vector-effect': 'non-scaling-stroke' }));
      overlay.append(outline);
    }
    options.onSelectionChange?.([...selected].map(id => {
      const { element, cell } = targets.get(id);
      return { slideId: slide.id, elementId: element.id, ...(cell ? { cellId: cell.id, row: cell.row, column: cell.column } : {}),
        sourceId: element.sourceId,
        type: cell ? 'cell' : element.type, name: cell ? `${element.name}, row ${cell.row + 1}, column ${cell.column + 1}` : element.name,
        text: textContent(cell ?? element),
        frame: { x: element.x, y: element.y, width: element.width, height: element.height, rotation: element.rotation, flipH: element.flipH, flipV: element.flipV },
        ...(element.type === 'image' && deck.assets[element.assetId]
          ? { asset: Object.fromEntries(Object.entries(deck.assets[element.assetId]).filter(([key]) => key !== 'dataUrl')) }
          : {}) };
    }));
  };
  function target(node, element, transform, cell) {
    if (element.locked || options.selectable === false) return;
    const id = cell?.id ?? element.id;
    node.setAttribute('data-element-id', element.id);
    if (cell) node.setAttribute('data-cell-id', cell.id);
    node.setAttribute('tabindex', '0');
    node.setAttribute('role', 'button');
    node.setAttribute('aria-label', cell ? `Cell ${cell.row + 1}, ${cell.column + 1}: ${textContent(cell)}` : `${element.type}: ${textContent(element) || element.name}`);
    css(node, { cursor: 'pointer', outline: 'none' });
    targets.set(id, { element, cell, transform });
    const select = event => {
      event.stopPropagation();
      if (event.shiftKey || event.metaKey || event.ctrlKey) {
        if (selected.has(id)) selected.delete(id); else selected.add(id);
      } else selected = new Set([id]);
      notify();
    };
    node.addEventListener('click', select);
    node.addEventListener('keydown', event => {
      if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); select(event); }
    });
  }
  for (const element of slide.elements) {
    const { width: w, height: h } = element;
    const transform = `translate(${element.x} ${element.y}) translate(${w / 2} ${h / 2}) rotate(${element.rotation}) scale(${element.flipH ? -1 : 1} ${element.flipV ? -1 : 1}) translate(${-w / 2} ${-h / 2})`;
    const group = svg('g', { transform, 'data-object-id': element.id });
    if (element.type === 'image') {
      const image = deck.assets[element.assetId];
      if (image) {
        const crop = element.crop;
        const inner = svg('svg', { width: w, height: h, overflow: 'hidden' });
        const cw = Math.max(0.001, 1 - crop.l - crop.r), ch = Math.max(0.001, 1 - crop.t - crop.b);
        inner.append(svg('image', { href: image.dataUrl, x: -crop.l * w / cw, y: -crop.t * h / ch, width: w / cw, height: h / ch, preserveAspectRatio: 'none' }));
        group.append(inner);
      }
    } else if (element.type === 'shape' || element.type === 'text') {
      const attrs = { fill: element.fill, stroke: element.stroke, 'stroke-width': element.strokeWidth };
      if (element.geometry === 'ellipse') group.append(svg('ellipse', { cx: w / 2, cy: h / 2, rx: w / 2, ry: h / 2, ...attrs }));
      else if (element.geometry === 'line') group.append(svg('line', { x1: 0, y1: 0, x2: w, y2: h, ...attrs }));
      else group.append(svg('rect', { width: w, height: h, rx: element.geometry === 'roundRect' ? element.radius : 0, ...attrs }));
      drawText(group, element.text, w, h, options.fonts);
    } else if (element.type === 'table') {
      target(group, element, transform);
      for (const cell of element.cells) {
        const cellGroup = svg('g', { transform: `translate(${cell.x} ${cell.y})` });
        cellGroup.append(svg('rect', { width: cell.width, height: cell.height, fill: cell.fill }));
        for (const [side, border] of Object.entries(cell.borders)) {
          const vertical = side === 'L' || side === 'R';
          const x = side === 'R' ? cell.width : 0, y = side === 'B' ? cell.height : 0;
          cellGroup.append(svg('line', { x1: x, y1: y, x2: vertical ? x : cell.width, y2: vertical ? cell.height : y, stroke: border.color, 'stroke-width': border.width }));
        }
        drawText(cellGroup, cell.text, cell.width, cell.height, options.fonts);
        cellGroup.append(svg('rect', { width: cell.width, height: cell.height, fill: 'transparent' }));
        target(cellGroup, element, transform, cell);
        group.append(cellGroup);
      }
    } else {
      group.append(svg('rect', { width: w || 180, height: h || 50, fill: '#fff7ed', stroke: '#ea580c', 'stroke-dasharray': '5 3' }));
      const label = svg('text', { x: 8, y: 20, 'font-size': 14, fill: '#9a3412' });
      label.textContent = element.reason;
      group.append(label);
    }
    if (element.type !== 'table') {
      if (!element.locked) group.append(svg('rect', { width: Math.max(w, 1), height: Math.max(h, 1), fill: 'transparent' }));
      target(group, element, transform);
    }
    layer.append(group);
  }
  root.append(layer, overlay);
  root.addEventListener('click', () => { selected.clear(); notify(); });
  root.addEventListener('keydown', e => { if (e.key === 'Escape') { selected.clear(); notify(); } });
  container.replaceChildren(root);
  const warnings = [];
  for (const fo of root.querySelectorAll('foreignObject')) {
    const box = fo.firstElementChild, content = box?.firstElementChild;
    if (!box || !content) continue;
    const available = box.clientHeight - parseFloat(box.style.paddingTop) - parseFloat(box.style.paddingBottom);
    if (content.scrollHeight > available + 2) warnings.push({
      elementId: fo.closest('[data-object-id]')?.getAttribute('data-object-id'),
      message: 'Text extends beyond its original box; font metrics or autofit need adjustment.',
    });
  }
  return {
    warnings,
    clearSelection() { selected.clear(); notify(); },
    destroy() { root.remove(); },
  };
}
