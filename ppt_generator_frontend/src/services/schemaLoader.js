/**
 * Runtime schema loader: loads wizard schema + flow schema + template model JSON from public/assets.
 * This keeps the UI fully client-side and avoids bundling large templates in JS.
 *
 * IMPORTANT:
 * - The extracted template JSONs under public/assets/template_extracted/* are the source of truth
 *   for placeholder IDs and (when present) exact geometry (x,y,cx,cy -> pts).
 * - Extraction may be partial/minimal; all helper APIs must be defensive and fall back gracefully.
 */

const TEMPLATE_EXTRACTED_BASE = '/assets/template_extracted';

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

async function loadJsonOrNull(path) {
  try {
    const res = await fetch(path, { cache: 'no-store' });
    if (!res.ok) return null;
    return await res.json();
  } catch {
    return null;
  }
}

// PUBLIC_INTERFACE
export async function loadExtractedTemplateArtifacts() {
  /**
   * Load extracted template artifacts:
   * - normalized (layouts/slides/theme; positioning when available)
   * - masters (more verbose master details; styles/fixed shapes)
   * - relationships (debug maps)
   * - assets manifest (images inventory)
   *
   * Returns { normalized, masters, relationships, assetsManifest } with nulls when missing.
   */
  const [normalized, masters, relationships, assetsManifest] = await Promise.all([
    loadJsonOrNull(`${TEMPLATE_EXTRACTED_BASE}/pptx_template.normalized.json`),
    loadJsonOrNull(`${TEMPLATE_EXTRACTED_BASE}/pptx_template.masters.json`),
    loadJsonOrNull(`${TEMPLATE_EXTRACTED_BASE}/pptx_template.relationships.json`),
    loadJsonOrNull(`${TEMPLATE_EXTRACTED_BASE}/assets_manifest.json`),
  ]);
  return { normalized, masters, relationships, assetsManifest };
}

// PUBLIC_INTERFACE
export async function loadTemplateModel() {
  /**
   * Load the PPT normalized template model JSON (drives preview + generation mapping).
   *
   * Preference order:
   * 1) extracted normalized template: /assets/template_extracted/pptx_template.normalized.json
   * 2) fallback contract/template: /assets/pptx_template_schema.json
   */
  const extracted = await loadJsonOrNull(`${TEMPLATE_EXTRACTED_BASE}/pptx_template.normalized.json`);
  if (extracted && typeof extracted === 'object') return extracted;

  const res = await fetch('/assets/pptx_template_schema.json', { cache: 'no-store' });
  if (!res.ok) {
    throw new Error(`Failed to load template schema: ${res.status}`);
  }
  return res.json();
}

// PUBLIC_INTERFACE
export async function loadTemplateBundle() {
  /**
   * Load a "bundle" used throughout the app: templateModel + extracted artifacts.
   * This preserves existing App contract (templateModel), but also provides masters/assets for
   * template-driven positioning and asset placement (when available).
   */
  const [templateModel, extracted] = await Promise.all([loadTemplateModel(), loadExtractedTemplateArtifacts()]);
  return {
    templateModel,
    extractedTemplate: extracted,
  };
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

function isValidRect(box) {
  return (
    box &&
    typeof box.xPt === 'number' &&
    typeof box.yPt === 'number' &&
    typeof box.wPt === 'number' &&
    typeof box.hPt === 'number' &&
    box.wPt > 0 &&
    box.hPt > 0
  );
}

// PUBLIC_INTERFACE
export function buildTemplateIndex(templateModel, extractedTemplate = null) {
  /**
   * Build quick lookup maps for layouts, placeholders, and extracted assets.
   *
   * - templateModel is expected to follow pptx_template_schema.json, but may be minimal.
   * - extractedTemplate (masters/assets/relationships) may help for future fidelity.
   *
   * Returns:
   * {
   *   layoutById, placeholderById, layouts,
   *   pageSizePt, theme,
   *   assetsById, assetsManifest,
   *   isTemplatePositioningActive
   * }
   */
  const layouts = Array.isArray(templateModel?.layouts) ? templateModel.layouts : [];
  const layoutById = new Map(layouts.map((l) => [l.id, l]));

  const placeholderById = new Map();
  let hasAnyExactBox = false;

  for (const l of layouts) {
    if (Array.isArray(l?.placeholders)) {
      for (const ph of l.placeholders) {
        if (ph?.id) {
          placeholderById.set(ph.id, { ...ph, layoutId: l.id });
          if (isValidRect(ph.box)) hasAnyExactBox = true;
        }
      }
    }
  }

  const pageSizePt = {
    widthPt: templateModel?.meta?.pageSize?.widthPt,
    heightPt: templateModel?.meta?.pageSize?.heightPt,
  };

  const theme = templateModel?.theme || {};

  // Assets: from schema model (templateModel.assets.*) and from extracted assets_manifest.json
  const assetsById = new Map();
  const addAssetRef = (a) => {
    if (!a?.id) return;
    assetsById.set(a.id, a);
  };

  const assets = templateModel?.assets || {};
  (assets.images || []).forEach(addAssetRef);
  (assets.icons || []).forEach(addAssetRef);
  (assets.fonts || []).forEach(addAssetRef);

  const assetsManifest = extractedTemplate?.assetsManifest || extractedTemplate?.assets_manifest || null;
  const manifestAssets = Array.isArray(assetsManifest?.assets) ? assetsManifest.assets : [];
  for (const a of manifestAssets) addAssetRef(a);

  return {
    layoutById,
    placeholderById,
    layouts,
    pageSizePt,
    theme,
    assetsById,
    assetsManifest,
    isTemplatePositioningActive: hasAnyExactBox,
  };
}

// PUBLIC_INTERFACE
export function getTemplatePlaceholder(templateIndex, placeholderId) {
  /** Lookup a placeholder definition by its extracted placeholder id/name. */
  if (!templateIndex || !placeholderId) return null;
  return templateIndex.placeholderById?.get(placeholderId) || null;
}

// PUBLIC_INTERFACE
export function getLayout(templateIndex, layoutId) {
  /** Lookup a layout by id. */
  if (!templateIndex || !layoutId) return null;
  return templateIndex.layoutById?.get(layoutId) || null;
}

// PUBLIC_INTERFACE
export function getAssetPublicUrl(assetRef) {
  /**
   * Resolve an asset reference to a public URL under /assets/template_extracted.
   * This assumes extracted binaries (e.g. images) are stored within that folder.
   *
   * If we only have an id/originalTarget but no suggestedName, we attempt to use suggestedName if present.
   */
  if (!assetRef) return null;

  // Prefer suggestedName if provided by extractor.
  if (assetRef.suggestedName) return `${TEMPLATE_EXTRACTED_BASE}/${assetRef.suggestedName}`;

  // If originalTarget references ppt/media/image.png, we cannot serve that directly unless extracted wrote it.
  // Keep a best-effort path: if originalTarget ends with a filename, use that.
  const t = assetRef.originalTarget;
  if (typeof t === 'string' && t.includes('/')) {
    const fileName = t.split('/').pop();
    if (fileName) return `${TEMPLATE_EXTRACTED_BASE}/${fileName}`;
  }

  return null;
}
