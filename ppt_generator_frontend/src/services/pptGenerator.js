import PptxGenJS from 'pptxgenjs';

/**
 * This generator is intentionally conservative: because the current extraction is partial,
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

function resolveDeckTitle(wizardSchema, formData) {
  // Prefer a common "title" field if present anywhere; otherwise fallback to generic.
  const maybeTitle =
    formData?.title ||
    formData?.deckTitle ||
    formData?.presentationTitle ||
    formData?.coverTitle;

  if (typeof maybeTitle === 'string' && maybeTitle.trim()) return maybeTitle.trim();

  const firstStepTitle = wizardSchema?.slides?.[0]?.title;
  if (typeof firstStepTitle === 'string' && firstStepTitle.trim()) return firstStepTitle.trim();

  return 'presentation';
}

function safeFileName(name) {
  return String(name)
    .toLowerCase()
    .replace(/[^a-z0-9\-_ ]/g, '')
    .trim()
    .replace(/\s+/g, '-')
    .slice(0, 80) || 'presentation';
}

// PUBLIC_INTERFACE
export async function generatePptx({ templateModel, wizardSchema, formData }) {
  /**
   * Generate a PPTX in-browser and return a Blob plus a suggested filename.
   * - templateModel: normalized template model (eventually from extractor; currently partial/contract)
   * - wizardSchema: drives slide order and mapping of fields to placeholder IDs
   * - formData: values keyed by wizard field id
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
    // PptxGenJS supports "layout" strings; custom sizes are more involved.
    // For now, use LAYOUT_WIDE (16:9) which matches common templates.
    pptx.layout = 'LAYOUT_WIDE';
  } else {
    pptx.layout = 'LAYOUT_WIDE';
  }

  const slides = Array.isArray(wizardSchema?.slides) ? wizardSchema.slides : [];

  // Minimal layout abstraction: each wizard slide becomes one PPT slide.
  // Mapping is by placeholderId (preferred) or stable fallback: slide-{index}:{fieldId}
  for (const s of slides) {
    const slide = pptx.addSlide();
    // Store layout name for later debugging (not all viewers show it)
    slide.addNotes(`layoutId=${s.layoutId || ''}`);

    const sections = Array.isArray(s.sections) ? s.sections : [];
    for (const section of sections) {
      const fields = Array.isArray(section.fields) ? section.fields : [];
      for (const field of fields) {
        const valueRaw = formData?.[field.id];
        const value = applyTransform(valueRaw, field?.mapping?.transform);

        const placeholderId = field?.mapping?.placeholderId || `slide-${s.slideIndex}:${field.id}`;
        const hint = `[${placeholderId}]`;

        // When coordinates are unknown, we place content in a simple flowing grid.
        // This is a scaffold; later, we will use templateModel layouts/placeholders boxes.
        const idx = fields.indexOf(field);
        const x = 0.6;
        const y = 0.6 + idx * 0.55;
        const w = 12.3;
        const h = 0.45;

        if (field.type === 'image') {
          const dataUrl = await fileToDataUrl(valueRaw);
          if (dataUrl) {
            slide.addImage({ data: dataUrl, x, y, w: 3.0, h: 2.0 });
            slide.addText(hint, { x: x + 3.2, y, w: w - 3.2, h, fontSize: 10, color: '666666' });
          } else {
            slide.addText(`${hint} (image missing)`, { x, y, w, h, fontSize: 12, color: '999999' });
          }
        } else {
          const text =
            value == null || value === ''
              ? `${hint} (empty)`
              : `${hint}\n${String(value)}`;

          slide.addText(text, {
            x,
            y,
            w,
            h: 0.9,
            fontSize: 14,
            color: '111827',
          });
        }
      }
    }
  }

  const title = resolveDeckTitle(wizardSchema, formData);
  const fileName = `${safeFileName(title)}.pptx`;

  const blob = await pptx.write('blob');
  return { blob, fileName };
}
