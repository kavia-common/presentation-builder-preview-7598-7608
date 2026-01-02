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

/**
 * Canonical extracted template artifact paths (must not be invented elsewhere).
 * These are treated as the fixed template source for this app.
 */
const TEMPLATE_EXTRACTED_FILES = Object.freeze({
  normalized: `${TEMPLATE_EXTRACTED_BASE}/pptx_template.normalized.json`,
  masters: `${TEMPLATE_EXTRACTED_BASE}/pptx_template.masters.json`,
  relationships: `${TEMPLATE_EXTRACTED_BASE}/pptx_template.relationships.json`,
  assetsManifest: `${TEMPLATE_EXTRACTED_BASE}/assets_manifest.json`,
});

/**
 * Canonical slide type sequence for the product workflow.
 * The template bundle may optionally map these to deck slide indices/layoutIds.
 */
const CANONICAL_FLOW = Object.freeze({
  globalFirst: 'global_first',
  skillFactory: ['sf1', 'sf2', 'sf3', 'sf4'],
  globalLast: 'global_last',
});

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
    loadJsonOrNull(TEMPLATE_EXTRACTED_FILES.normalized),
    loadJsonOrNull(TEMPLATE_EXTRACTED_FILES.masters),
    loadJsonOrNull(TEMPLATE_EXTRACTED_FILES.relationships),
    loadJsonOrNull(TEMPLATE_EXTRACTED_FILES.assetsManifest),
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
   *
   * IMPORTANT:
   * If the extracted template JSON exists but is clearly a placeholder (empty layouts or 0x0 page),
   * we throw an error to avoid silently showing an incorrect "fallback" preview while claiming the
   * template bundle is present.
   */
  const extracted = await loadJsonOrNull(`${TEMPLATE_EXTRACTED_BASE}/pptx_template.normalized.json`);
  if (extracted && typeof extracted === 'object') {
    const widthPt = extracted?.meta?.pageSize?.widthPt;
    const heightPt = extracted?.meta?.pageSize?.heightPt;
    const layoutsCount = Array.isArray(extracted?.layouts) ? extracted.layouts.length : 0;

    const looksPlaceholder =
      layoutsCount === 0 ||
      !Number.isFinite(widthPt) ||
      !Number.isFinite(heightPt) ||
      widthPt <= 0 ||
      heightPt <= 0;

    if (looksPlaceholder) {
      throw new Error(
        'Extracted template bundle found but appears incomplete (missing layouts and/or slide size). ' +
          'Re-extract the PPTX or populate public/assets/template_extracted/*.json with real data.'
      );
    }

    return extracted;
  }

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
  // IMPORTANT: assets_manifest.json is canonical for extracted binaries available under public/assets/template_extracted/.
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

/**
 * Attempt to resolve a template "asset reference" into a public URL.
 * We prefer assets from assets_manifest.json when possible since that reflects extracted binaries.
 */
function resolveAssetRef(templateIndex, assetOrId) {
  if (!assetOrId) return null;
  if (typeof assetOrId === 'object') return assetOrId;
  if (typeof assetOrId === 'string') return templateIndex?.assetsById?.get(assetOrId) || null;
  return null;
}

// PUBLIC_INTERFACE
export function resolveTemplateAssetUrl(templateIndex, assetOrId) {
  /**
   * Resolve an asset (id or ref object) to a public URL under /assets/template_extracted.
   * Returns null if asset not resolvable.
   */
  const ref = resolveAssetRef(templateIndex, assetOrId);
  return getAssetPublicUrl(ref);
}

/**
 * Best-effort: find slide indices for GlobalFirst / GlobalLast and the 4-slide SkillFactory group in the extracted template.
 * Since extracted JSON may be minimal, we fall back to canonical ordering purely by flow schema + wizardData.
 */
function deriveTemplateSlideIndexMap(templateModel) {
  // The extractor may (in future) include mapping in meta.slideTypeToSlideIndex.
  const map = templateModel?.meta?.slideTypeToSlideIndex;
  if (map && typeof map === 'object') return map;

  // Fallback: if template slides exist and have notes containing our known slideType markers.
  // This is non-breaking and optional; if not found, return {}.
  const out = {};
  const slides = Array.isArray(templateModel?.slides) ? templateModel.slides : [];
  for (const s of slides) {
    const notes = typeof s?.notes === 'string' ? s.notes : '';
    if (notes.includes('slideType=global_first')) out[CANONICAL_FLOW.globalFirst] = s.index;
    if (notes.includes('slideType=global_last')) out[CANONICAL_FLOW.globalLast] = s.index;
    // SkillFactory slide keys can't be uniquely inferred without additional metadata.
  }
  return out;
}

// PUBLIC_INTERFACE
export function buildCanonicalOrderedSlidesFromTemplate({ flowSchema, wizardData, templateModel }) {
  /**
   * Build ordered slide list strictly following canonical flow:
   * [Global First] + k * [SF1..SF4] + [Global Last]
   *
   * IMPORTANT:
   * - Slide *types* and group structure are locked; only the number/order of Skill Factory groups can change.
   * - LayoutIds are sourced from flowSchema.slideTypes[*].layoutId, with optional override from template slide index mapping.
   *
   * Returns an array of step objects.
   */
  if (!flowSchema || !wizardData) return [];
  const map = deriveTemplateSlideIndexMap(templateModel);

  const resolveLayoutId = (slideType, fallback) => {
    // If the template explicitly maps slideType -> slideIndex, and templateModel.slides has layoutId, use it.
    const idx = map?.[slideType];
    if (typeof idx === 'number') {
      const found = (Array.isArray(templateModel?.slides) ? templateModel.slides : []).find((s) => s?.index === idx);
      if (found?.layoutId) return found.layoutId;
    }
    return fallback;
  };

  const steps = [];

  steps.push({
    key: 'global_first',
    kind: 'global',
    slideType: CANONICAL_FLOW.globalFirst,
    title: flowSchema?.slideTypes?.global_first?.label || 'Global First',
    layoutId: resolveLayoutId(CANONICAL_FLOW.globalFirst, flowSchema?.slideTypes?.global_first?.layoutId || 'global_first'),
    fields: getFieldsForSlideType(flowSchema, CANONICAL_FLOW.globalFirst),
    dataPath: { scope: 'globalFirst' },
  });

  const factories = Array.isArray(wizardData?.skillFactories) ? wizardData.skillFactories : [];
  factories.forEach((sf, idx) => {
    const groupLabel = `Skill Factory ${idx + 1}`;
    CANONICAL_FLOW.skillFactory.forEach((sft) => {
      steps.push({
        key: `${sf.id}:${sft}`,
        kind: 'skillFactory',
        groupId: sf.id,
        groupIndex: idx,
        groupLabel,
        slideType: sft,
        title: `${groupLabel} — ${flowSchema?.slideTypes?.[sft]?.label || sft.toUpperCase()}`,
        layoutId: resolveLayoutId(sft, flowSchema?.slideTypes?.[sft]?.layoutId || sft),
        fields: getFieldsForSlideType(flowSchema, sft),
        dataPath: { scope: 'skillFactories', groupId: sf.id, slideKey: sft },
      });
    });
  });

  steps.push({
    key: 'global_last',
    kind: 'global',
    slideType: CANONICAL_FLOW.globalLast,
    title: flowSchema?.slideTypes?.global_last?.label || 'Global Last',
    layoutId: resolveLayoutId(CANONICAL_FLOW.globalLast, flowSchema?.slideTypes?.global_last?.layoutId || 'global_last'),
    fields: getFieldsForSlideType(flowSchema, CANONICAL_FLOW.globalLast),
    dataPath: { scope: 'globalLast' },
  });

  return steps;
}

// PUBLIC_INTERFACE
export function validateSlideRequiredFields({ slideStep, wizardData, templateIndex }) {
  /**
   * Validate required placeholders for a slide step without breaking preview.
   *
   * Sources of "required":
   * 1) Field schema validation.required (always authoritative)
   * 2) Template placeholder constraints.required when placeholder is referenced by a field mapping
   *
   * Returns: Array<{ fieldId, placeholderId, kind, message }>
   */
  const warnings = [];
  if (!slideStep || !wizardData) return warnings;

  const fields = Array.isArray(slideStep?.fields) ? slideStep.fields : [];
  for (const f of fields) {
    const placeholderId = f?.mapping?.placeholderId;
    const isRequiredBySchema = Boolean(f?.validation?.required);
    const tpl = placeholderId ? getTemplatePlaceholder(templateIndex, placeholderId) : null;
    const isRequiredByTemplate = Boolean(tpl?.constraints?.required);

    if (!isRequiredBySchema && !isRequiredByTemplate) continue;

    // Resolve value (duplicated logic kept tiny to avoid circular imports).
    const p = slideStep?.dataPath;
    let v;
    if (p?.scope === 'globalFirst') v = wizardData?.globalFirst?.[f.id];
    else if (p?.scope === 'globalLast') v = wizardData?.globalLast?.[f.id];
    else if (p?.scope === 'skillFactories') {
      const group = (wizardData?.skillFactories || []).find((g) => g.id === p.groupId);
      v = group?.slides?.[p.slideKey]?.[f.id];
    }

    const empty = v == null || v === '' || (Array.isArray(v) && v.length === 0);
    if (!empty) continue;

    warnings.push({
      fieldId: f.id,
      placeholderId: placeholderId || null,
      kind: f.type === 'image' ? 'image' : 'text',
      message: f.type === 'image' ? 'Missing required image.' : 'Missing required text.',
    });
  }

  return warnings;
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
