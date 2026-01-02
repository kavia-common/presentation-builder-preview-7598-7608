import PptxGenJS from 'pptxgenjs';
import { buildTemplateIndex, getAssetPublicUrl, getLayout, getTemplatePlaceholder } from './schemaLoader';

/**
 * Generator rules (template-driven):
 * - Prefer coordinates and placeholder mappings from extracted normalized template (templateModel.layouts[*].placeholders[*].box).
 * - Slide order is authoritative: GlobalFirst + N SkillFactory groups (4 slides each) + GlobalLast.
 * - If extracted template is minimal, fall back to safe stacking layout (non-breaking).
 */

function isProbablyDataUrl(v) {
  return typeof v === 'string' && v.startsWith('data:');
}

async function fileToDataUrl(file) {
  if (!file) return null;
  if (typeof file === 'string' && isProbablyDataUrl(file)) return file;
  if (typeof File !== 'undefined' && file instanceof File) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new Error('Failed to read image file'));
      reader.onload = () => resolve(String(reader.result));
      reader.readAsDataURL(file);
    });
  }
  return null;
}

async function urlToDataUrl(url) {
  if (!url || typeof url !== 'string') return null;
  if (isProbablyDataUrl(url)) return url;
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    const blob = await res.blob();
    return await new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onerror = () => reject(new Error('Failed to read asset'));
      r.onload = () => resolve(String(r.result));
      r.readAsDataURL(blob);
    });
  } catch {
    return null;
  }
}

function applyTransform(value, transformName) {
  if (!transformName) return value;
  if (value == null) return value;

  switch (transformName) {
    case 'bulletsToLines': {
      if (Array.isArray(value)) return value.map(String).join('\n');
      return String(value);
    }
    case 'formatPercent': {
      const n = typeof value === 'number' ? value : Number(value);
      if (Number.isNaN(n)) return String(value);
      return `${(n * 100).toFixed(0)}%`;
    }
    case 'formatCurrency': {
      const n = typeof value === 'number' ? value : Number(value);
      if (Number.isNaN(n)) return String(value);
      return `$${n.toLocaleString()}`;
    }
    default:
      return value;
  }
}

function resolveDeckTitle(wizardData) {
  const maybeTitle = wizardData?.globalFirst?.title;
  if (typeof maybeTitle === 'string' && maybeTitle.trim()) return maybeTitle.trim();
  return 'presentation';
}

function safeFileName(name) {
  return (
    String(name)
      .toLowerCase()
      .replace(/[^a-z0-9\-_ ]/g, '')
      .trim()
      .replace(/\s+/g, '-')
      .slice(0, 80) || 'presentation'
  );
}

function resolveValueForStep(wizardData, slideStep, fieldId) {
  const p = slideStep?.dataPath;
  if (!p) return undefined;

  if (p.scope === 'globalFirst') return wizardData?.globalFirst?.[fieldId];
  if (p.scope === 'globalLast') return wizardData?.globalLast?.[fieldId];

  if (p.scope === 'skillFactories') {
    const group = (wizardData?.skillFactories || []).find((g) => g.id === p.groupId);
    return group?.slides?.[p.slideKey]?.[fieldId];
  }

  return undefined;
}

function ptToIn(pt) {
  // PptxGen uses inches. 72 points per inch.
  return pt / 72;
}

function addFallbackText(slide, i, placeholderId, value) {
  const x = 0.6;
  const y = 0.6 + i * 0.65;
  const w = 12.3;
  const h = 1.0;
  const hint = `[${placeholderId}]`;

  const text = value == null || value === '' ? `${hint} (empty)` : `${hint}\n${String(value)}`;
  slide.addText(text, {
    x,
    y,
    w,
    h,
    fontSize: 14,
    color: '111827',
  });
}

async function addFallbackImage(slide, i, placeholderId, fileOrUrl) {
  const x = 0.6;
  const y = 0.6 + i * 0.65;
  const w = 3.0;
  const h = 2.0;
  const hint = `[${placeholderId}]`;

  const dataUrl = await fileToDataUrl(fileOrUrl);
  if (dataUrl) {
    slide.addImage({ data: dataUrl, x, y, w, h });
    slide.addText(hint, { x: x + 3.2, y, w: 9.1, h: 0.5, fontSize: 10, color: '666666' });
  } else {
    slide.addText(`${hint} (image missing)`, { x, y, w: 12.3, h: 0.6, fontSize: 12, color: '999999' });
  }
}

async function renderFixedShapes(slide, layout, templateIndex) {
  const fixed = Array.isArray(layout?.fixedShapes) ? layout.fixedShapes : [];
  for (const sh of fixed) {
    if (!sh?.box) continue;
    const x = ptToIn(sh.box.xPt);
    const y = ptToIn(sh.box.yPt);
    const w = ptToIn(sh.box.wPt);
    const h = ptToIn(sh.box.hPt);

    if (sh.shapeType === 'picture' && sh.assetId) {
      const ref = templateIndex?.assetsById?.get(sh.assetId) || null;
      const url = getAssetPublicUrl(ref);
      const data = await urlToDataUrl(url);
      if (data) {
        slide.addImage({ data, x, y, w, h });
      }
      continue;
    }

    // Minimal support for filled rectangles (common for master bands); others can be expanded later.
    if (sh.shapeType === 'rect' || sh.shapeType === 'roundRect' || sh.shapeType === 'unknown') {
      slide.addShape(PptxGenJS.ShapeType.rect, {
        x,
        y,
        w,
        h,
        fill: sh.fill && sh.fill !== 'none' ? { color: String(sh.fill).replace('#', '') } : undefined,
        line: sh.stroke && sh.stroke !== 'none' ? { color: String(sh.stroke).replace('#', ''), width: sh.strokeWidthPt || 0.5 } : undefined,
      });
    }
  }
}

// PUBLIC_INTERFACE
export async function generatePptx({ templateModel, extractedTemplate, orderedSlides, wizardData }) {
  /**
   * Generate a PPTX in-browser and return a Blob plus a suggested filename.
   *
   * - templateModel: extracted normalized template model (preferred) or fallback schema contract
   * - extractedTemplate: {masters, relationships, assetsManifest} (optional)
   * - orderedSlides: ordered steps for the current wizard instance
   * - wizardData: grouped data model {globalFirst, skillFactories[], globalLast}
   */
  const pptx = new PptxGenJS();

  pptx.author = 'Presentation Builder';
  pptx.company = 'Kavia';
  pptx.subject = 'Generated deck';

  // Layout: keep wide. If extracted provides exact size, we still set to wide; coordinates are in inches anyway.
  pptx.layout = 'LAYOUT_WIDE';

  const slides = Array.isArray(orderedSlides) ? orderedSlides : [];
  const templateIndex = buildTemplateIndex(templateModel, extractedTemplate);

  for (const s of slides) {
    const slide = pptx.addSlide();
    slide.addNotes(`layoutId=${s.layoutId || ''} slideType=${s.slideType || ''}`);

    const layout = getLayout(templateIndex, s.layoutId);

    // Add fixed shapes (layout/master) if present in normalized template.
    if (layout) {
      // eslint-disable-next-line no-await-in-loop
      await renderFixedShapes(slide, layout, templateIndex);
    }

    const fields = Array.isArray(s.fields) ? s.fields : [];

    for (let i = 0; i < fields.length; i += 1) {
      const field = fields[i];
      const valueRaw = resolveValueForStep(wizardData, s, field.id);
      const value = applyTransform(valueRaw, field?.mapping?.transform);

      const placeholderId =
        field?.mapping?.placeholderId ||
        `${String(s.slideType || 'slide').toUpperCase()}:${field.id}`;

      // Template-driven placement:
      // 1) Find placeholder by id in extracted template
      // 2) Use its box (points) -> inches
      // 3) Apply style hints (font size, color) when available
      const ph = getTemplatePlaceholder(templateIndex, placeholderId);
      const box = ph?.box;

      if (field.type === 'image') {
        let dataUrl = await fileToDataUrl(valueRaw);
        if (!dataUrl && ph?.assetId) {
          const ref = templateIndex?.assetsById?.get(ph.assetId) || null;
          const url = getAssetPublicUrl(ref);
          dataUrl = await urlToDataUrl(url);
        }

        if (box && typeof box.xPt === 'number') {
          const x = ptToIn(box.xPt);
          const y = ptToIn(box.yPt);
          const w = ptToIn(box.wPt);
          const h = ptToIn(box.hPt);

          if (dataUrl) {
            slide.addImage({ data: dataUrl, x, y, w, h });
          } else {
            // Keep placeholder id visible for later fidelity
            slide.addText(`[${placeholderId}] (image missing)`, { x, y, w, h: Math.min(h, 0.4), fontSize: 10, color: '999999' });
          }
        } else {
          // eslint-disable-next-line no-await-in-loop
          await addFallbackImage(slide, i, placeholderId, valueRaw);
        }
      } else {
        if (box && typeof box.xPt === 'number') {
          const x = ptToIn(box.xPt);
          const y = ptToIn(box.yPt);
          const w = ptToIn(box.wPt);
          const h = ptToIn(box.hPt);

          const text = value == null || value === '' ? '' : String(value);
          const style = ph?.style || null;

          const fontSize = style?.fontSizePt ? Math.max(8, style.fontSizePt) : 14;
          const color = style?.color ? String(style.color).replace('#', '') : '111827';
          const bold = typeof style?.fontWeight === 'number' ? style.fontWeight >= 700 : false;
          const align = style?.align || 'left';

          slide.addText(text || `[${placeholderId}]`, {
            x,
            y,
            w,
            h,
            fontSize,
            bold,
            color,
            align,
            fontFace: style?.fontFamily || undefined,
          });
        } else {
          addFallbackText(slide, i, placeholderId, value);
        }
      }
    }
  }

  const title = resolveDeckTitle(wizardData);
  const fileName = `${safeFileName(title)}.pptx`;

  const blob = await pptx.write('blob');
  return { blob, fileName };
}
