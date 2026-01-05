import PptxGenJS from 'pptxgenjs';
import {
  buildTemplateIndex,
  getLayout,
  getTemplatePlaceholder,
  resolveTemplateAssetUrl,
} from './schemaLoader';
import { formatDdMmmYyyy, formatDateRangeDdMmmYyyy } from '../utils/dateFormat';

/**
 * Generator rules (template-driven):
 * - Prefer coordinates and placeholder mappings from extracted normalized template (templateModel.layouts[*].placeholders[*].box).
 * - Slide order is authoritative: GlobalFirst + N SkillFactory groups (4 slides each) + GlobalLast.
 * - If extracted template is minimal, fall back to safe stacking layout (non-breaking).
 *
 * IMPORTANT for this task:
 * - Do not insert any warning/diagnostic/placeholder-notice text into the generated PPTX.
 * - Keep any diagnostics only in developer console logs (non-visual).
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
      if (Array.isArray(value)) return value.map((s) => String(s ?? '').trim()).filter(Boolean).map((s) => `• ${s}`).join('\n');
      return String(value);
    }
    case 'teamMembersToLines': {
      if (!Array.isArray(value)) return '';
      return value
        .map((r) => {
          const name = String(r?.name ?? '').trim();
          const role = String(r?.role ?? '').trim();
          if (!name && !role) return '';
          return name && role ? `${name}    ${role}` : name || role;
        })
        .filter(Boolean)
        .join('\n');
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
  // Global First no longer collects a "title" input; use Name as the best available human-friendly filename base.
  const maybeName = wizardData?.globalFirst?.name;
  if (typeof maybeName === 'string' && maybeName.trim()) return maybeName.trim();
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

/**
 * Fallback rendering helpers:
 * These are kept non-diagnostic (no placeholder IDs, no "(missing)" messages).
 * They exist only so that content can still be written somewhere if template geometry is unavailable.
 */
function addFallbackText(slide, i, value) {
  const x = 0.6;
  const y = 0.6 + i * 0.65;
  const w = 12.3;
  const h = 1.0;

  const text = value == null ? '' : String(value);
  if (!text) return;

  slide.addText(text, {
    x,
    y,
    w,
    h,
    fontSize: 14,
    color: '111827',
  });
}

async function addFallbackImage(slide, i, fileOrUrl) {
  const x = 0.6;
  const y = 0.6 + i * 0.65;
  const w = 3.0;
  const h = 2.0;

  const dataUrl = await fileToDataUrl(fileOrUrl);
  if (!dataUrl) return;

  slide.addImage({ data: dataUrl, x, y, w, h });
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
      const url = resolveTemplateAssetUrl(templateIndex, sh.assetId);
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
    const isGlobalFirst = s.slideType === 'global_first';
    const isGlobalLast = s.slideType === 'global_last';

    // Global First and Global Last must match the provided reference images as backgrounds.
    // Global Last is locked to background-only.
    // Global First has ONLY one editable overlay: Date (GF_DATE).
    if (isGlobalFirst || isGlobalLast) {
      // eslint-disable-next-line no-await-in-loop
      const bgData = await urlToDataUrl(isGlobalFirst ? '/assets/global_first_background.png' : '/assets/global_last_background.png');
      if (bgData) {
        // Match preview: preserve aspect ratio with NO cropping (letterbox if needed).
        // NOTE: PptxGenJS sizing:'contain' keeps full image visible without distortion.
        slide.addImage({ data: bgData, x: 0, y: 0, w: 13.333, h: 7.5, sizing: { type: 'contain' } });
      }
    }

    // Global Last: after inserting the fixed background, render nothing else.
    if (isGlobalLast) {
      // eslint-disable-next-line no-continue
      continue;
    }

    // Global First: write ONLY the date text at GF_DATE using extracted box/style.
    // No fallback positioning is allowed for this locked slide.
    if (isGlobalFirst) {
      const ph = getTemplatePlaceholder(templateIndex, 'GF_DATE');
      const box = ph?.box;

      const rawDate = wizardData?.globalFirst?.date;
      const safeText = rawDate ? formatDdMmmYyyy(rawDate) : '';

      if (!box || typeof box.xPt !== 'number') {
        // eslint-disable-next-line no-console
        console.warn('[generatePptx] GF_DATE placeholder missing; date suppressed (no fallback placement).');
        // eslint-disable-next-line no-continue
        continue;
      }

      if (safeText) {
        const x = ptToIn(box.xPt);
        const y = ptToIn(box.yPt);
        const w = ptToIn(box.wPt);
        const h = ptToIn(box.hPt);

        const style = ph?.style || null;
        const fontSize = typeof style?.fontSizePt === 'number' ? Math.max(1, style.fontSizePt) : 14;
        const color = style?.color ? String(style.color).replace('#', '') : 'FFFFFF';
        const align = style?.align || 'left';
        const bold = typeof style?.fontWeight === 'number' ? style.fontWeight >= 700 : false;

        slide.addText(safeText, {
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
      }

      // eslint-disable-next-line no-continue
      continue;
    }

    // For non-locked slides (everything except global_first/global_last), we also render
    // template background assets + fixed shapes.
    if (!isGlobalFirst && !isGlobalLast) {
      // Layout background for non-locked slides.
      if (layout?.background?.assetId) {
        // eslint-disable-next-line no-await-in-loop
        const bgUrl = resolveTemplateAssetUrl(templateIndex, layout.background.assetId);
        // eslint-disable-next-line no-await-in-loop
        const bgData = await urlToDataUrl(bgUrl);
        if (bgData) {
          slide.addImage({ data: bgData, x: 0, y: 0, w: 13.333, h: 7.5 });
        }
      }

      // Add fixed shapes (layout/master) if present in normalized template.
      if (layout) {
        // eslint-disable-next-line no-await-in-loop
        await renderFixedShapes(slide, layout, templateIndex);
      }
    }

    // Other slides: render all fields.
    const fieldsAll = Array.isArray(s.fields) ? s.fields : [];
    const fields = fieldsAll;

    for (let i = 0; i < fields.length; i += 1) {
      const field = fields[i];
      let valueRaw = resolveValueForStep(wizardData, s, field.id);

      // Skill Factory Slide 1: Start/End date fields are rendered together as a single date-range string.
      if ((field.id === 'dateRangeStart' || field.id === 'dateRangeEnd') && s.slideType === 'sf1') {
        const start = resolveValueForStep(wizardData, s, 'dateRangeStart');
        const end = resolveValueForStep(wizardData, s, 'dateRangeEnd');
        valueRaw = formatDateRangeDdMmmYyyy(start, end);
      }

      // Support rich types as simple text for PPT (template-driven box/style still applies).
      const transform =
        field?.mapping?.transform ||
        (field.type === 'bullets' ? 'bulletsToLines' : null) ||
        (field.type === 'table' ? 'teamMembersToLines' : null);

      const value = applyTransform(valueRaw, transform);

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
          const url = resolveTemplateAssetUrl(templateIndex, ph.assetId);
          dataUrl = await urlToDataUrl(url);
        }

        if (box && typeof box.xPt === 'number') {
          const x = ptToIn(box.xPt);
          const y = ptToIn(box.yPt);
          const w = ptToIn(box.wPt);
          const h = ptToIn(box.hPt);

          // IMPORTANT: No placeholder notices. If image missing, render nothing.
          if (dataUrl) {
            slide.addImage({ data: dataUrl, x, y, w, h });
          } else {
            // eslint-disable-next-line no-console
            console.warn('[generatePptx] image missing (suppressed in PPT output):', { slide: s?.key, placeholderId });
          }
        } else if (!isGlobalFirst) {
          // eslint-disable-next-line no-await-in-loop
          await addFallbackImage(slide, i, valueRaw);
        }
      } else {
        if (box && typeof box.xPt === 'number') {
          const x = ptToIn(box.xPt);
          const y = ptToIn(box.yPt);
          const w = ptToIn(box.wPt);
          const h = ptToIn(box.hPt);

          const templateDefault = typeof ph?.text === 'string' && ph.text.trim() ? ph.text.trim() : '';
          const text = value == null || value === '' ? templateDefault : String(value);
          const style = ph?.style || null;

          // Exact template typography (no overrides).
          const fontSize = typeof style?.fontSizePt === 'number' ? Math.max(1, style.fontSizePt) : 14;
          const color = style?.color ? String(style.color).replace('#', '') : '111827';
          const align = style?.align || 'left';
          const bold = typeof style?.fontWeight === 'number' ? style.fontWeight >= 700 : false;

          // Global First is handled earlier (background + date-only), so we never enter here for it.
          // Non-GlobalFirst: no diagnostic text; use template default or empty.
          const safeText = text || '';

          if (safeText) {
            slide.addText(safeText, {
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
          }
        } else if (!isGlobalFirst) {
          addFallbackText(slide, i, value);
        }
      }
    }
  }

  const title = resolveDeckTitle(wizardData);
  const fileName = `${safeFileName(title)}.pptx`;

  const blob = await pptx.write('blob');
  return { blob, fileName };
}
