import PptxGenJS from 'pptxgenjs';

/**
 * This generator is intentionally conservative: because extraction may be partial,
 * it focuses on stable placeholder IDs + modular mapping to refine later without changing UI contracts.
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

// PUBLIC_INTERFACE
export async function generatePptx({ templateModel, flowSchema, orderedSlides, wizardData }) {
  /**
   * Generate a PPTX in-browser and return a Blob plus a suggested filename.
   * - templateModel: normalized template model (eventually from extractor; currently partial/contract)
   * - flowSchema: slide type definitions (fields + placeholder IDs)
   * - orderedSlides: ordered steps for the current wizard instance
   * - wizardData: grouped data model {globalFirst, skillFactories[], globalLast}
   */
  const pptx = new PptxGenJS();

  // Basic theme alignment; PptxGenJS supports theme, but we keep it minimal for now.
  // Colors are from Ocean Professional style guide.
  pptx.author = 'Presentation Builder';
  pptx.company = 'Kavia';
  pptx.subject = 'Generated deck';

  // If templateModel has pageSize, try to apply; else default to wide.
  const widthPt = templateModel?.meta?.pageSize?.widthPt;
  const heightPt = templateModel?.meta?.pageSize?.heightPt;
  if (typeof widthPt === 'number' && typeof heightPt === 'number') {
    pptx.layout = 'LAYOUT_WIDE';
  } else {
    pptx.layout = 'LAYOUT_WIDE';
  }

  const slides = Array.isArray(orderedSlides) ? orderedSlides : [];

  // Each ordered slide becomes one PPT slide.
  // Content is written with placeholder IDs in text to preserve stable mapping even without exact coordinates.
  for (const s of slides) {
    const slide = pptx.addSlide();
    slide.addNotes(`layoutId=${s.layoutId || ''} slideType=${s.slideType || ''}`);

    const fields = Array.isArray(s.fields) ? s.fields : [];
    for (let i = 0; i < fields.length; i += 1) {
      const field = fields[i];
      const valueRaw = resolveValueForStep(wizardData, s, field.id);
      const value = applyTransform(valueRaw, field?.mapping?.transform);

      const placeholderId =
        field?.mapping?.placeholderId ||
        `${String(s.slideType || 'slide').toUpperCase()}:${field.id}`;

      const hint = `[${placeholderId}]`;

      // Fallback layout: flowing vertical stack.
      // NOTE: We keep placeholderId visible so later coordinate-accurate placement can be implemented
      // without changing the schema or wizard.
      const x = 0.6;
      const y = 0.6 + i * 0.65;
      const w = 12.3;
      const h = 0.5;

      if (field.type === 'image') {
        const dataUrl = await fileToDataUrl(valueRaw);
        if (dataUrl) {
          slide.addImage({ data: dataUrl, x, y, w: 3.0, h: 2.0 });
          slide.addText(hint, { x: x + 3.2, y, w: w - 3.2, h, fontSize: 10, color: '666666' });
        } else {
          slide.addText(`${hint} (image missing)`, { x, y, w, h, fontSize: 12, color: '999999' });
        }
      } else {
        const text = value == null || value === '' ? `${hint} (empty)` : `${hint}\n${String(value)}`;
        slide.addText(text, {
          x,
          y,
          w,
          h: 1.0,
          fontSize: 14,
          color: '111827',
        });
      }
    }
  }

  const title = resolveDeckTitle(wizardData);
  const fileName = `${safeFileName(title)}.pptx`;

  const blob = await pptx.write('blob');
  return { blob, fileName };
}
