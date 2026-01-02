/**
 * Runtime schema loader: loads wizard schema + template model JSON from public/assets.
 * This keeps the UI fully client-side and avoids bundling large templates in JS.
 */

// PUBLIC_INTERFACE
export async function loadWizardSchema() {
  /** Load the form wizard schema JSON (drives step rendering). */
  const res = await fetch('/assets/form_wizard_schema.template.json', { cache: 'no-store' });
  if (!res.ok) {
    throw new Error(`Failed to load wizard schema: ${res.status}`);
  }
  return res.json();
}

// PUBLIC_INTERFACE
export async function loadTemplateModel() {
  /** Load the PPT normalized template model JSON (drives preview + generation mapping). */
  const res = await fetch('/assets/pptx_template_schema.json', { cache: 'no-store' });
  if (!res.ok) {
    throw new Error(`Failed to load template schema: ${res.status}`);
  }
  return res.json();
}

// PUBLIC_INTERFACE
export function buildTemplateIndex(templateModel) {
  /**
   * Build quick lookup maps for layouts and placeholders.
   * The "templateModel" might be a full extracted model in the future; today it's a schema contract,
   * so we keep this defensive and allow "layouts/slides" to be empty.
   */
  const layouts = Array.isArray(templateModel?.layouts) ? templateModel.layouts : [];
  const slides = Array.isArray(templateModel?.slides) ? templateModel.slides : [];
  const layoutById = new Map(layouts.map(l => [l.id, l]));
  const slideByIndex = new Map(slides.map(s => [s.index, s]));

  const placeholderById = new Map();
  for (const l of layouts) {
    if (Array.isArray(l?.placeholders)) {
      for (const ph of l.placeholders) {
        if (ph?.id) placeholderById.set(ph.id, { ...ph, layoutId: l.id });
      }
    }
  }

  return { layoutById, slideByIndex, placeholderById, layouts, slides };
}
