/**
 * Runtime schema loader: loads wizard schema + flow schema + template model JSON from public/assets.
 * This keeps the UI fully client-side and avoids bundling large templates in JS.
 */

// PUBLIC_INTERFACE
export async function loadWizardSchema() {
  /** Load the legacy form wizard schema JSON (drives step rendering in older flat mode). */
  const res = await fetch('/assets/form_wizard_schema.template.json', { cache: 'no-store' });
  if (!res.ok) {
    throw new Error(`Failed to load wizard schema: ${res.status}`);
  }
  return res.json();
}

// PUBLIC_INTERFACE
export async function loadWizardFlowSchema() {
  /** Load the flow schema JSON (slide types + fields + placeholder mappings). */
  const res = await fetch('/assets/wizard_flow_schema.json', { cache: 'no-store' });
  if (!res.ok) {
    throw new Error(`Failed to load flow schema: ${res.status}`);
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
export function getSlideTypeDefinition(flowSchema, slideType) {
  /** Return the flow schema definition for a slide type (global_first, sf1..sf4, global_last). */
  return flowSchema?.slideTypes?.[slideType] || null;
}

// PUBLIC_INTERFACE
export function getFieldsForSlideType(flowSchema, slideType) {
  /** Return fields array for a slide type; always returns an array. */
  const def = getSlideTypeDefinition(flowSchema, slideType);
  return Array.isArray(def?.fields) ? def.fields : [];
}

// PUBLIC_INTERFACE
export function getPlaceholderMappingForSlideType(flowSchema, slideType) {
  /**
   * Return mapping { fieldId -> placeholderId } for a slide type.
   * This is used by PPT generation and preview. Coordinates may be absent; IDs stay stable.
   */
  const fields = getFieldsForSlideType(flowSchema, slideType);
  const out = {};
  for (const f of fields) {
    const ph = f?.mapping?.placeholderId;
    if (f?.id && typeof ph === 'string' && ph) out[f.id] = ph;
  }
  return out;
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
  const layoutById = new Map(layouts.map((l) => [l.id, l]));
  const slideByIndex = new Map(slides.map((s) => [s.index, s]));

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
