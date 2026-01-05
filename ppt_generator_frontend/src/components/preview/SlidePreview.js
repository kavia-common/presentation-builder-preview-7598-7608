import React, { useEffect, useMemo, useState } from 'react';
import {
  buildTemplateIndex,
  getLayout,
  getTemplatePlaceholder,
  resolveTemplateAssetUrl,
  validateSlideRequiredFields,
} from '../../services/schemaLoader';
import { formatDdMmmYyyy, formatDateRangeDdMmmYyyy } from '../../utils/dateFormat';

function getPageSize(templateModel) {
  const wPt = templateModel?.meta?.pageSize?.widthPt;
  const hPt = templateModel?.meta?.pageSize?.heightPt;
  if (typeof wPt === 'number' && typeof hPt === 'number' && wPt > 0 && hPt > 0) return { wPt, hPt };
  // Default wide slide points (approx 13.333in x 7.5in at 72pt/in)
  return { wPt: 960, hPt: 540 };
}

function scaleRect(rect, page) {
  // Convert from points to % of slide canvas to keep responsive.
  const left = (rect.xPt / page.wPt) * 100;
  const top = (rect.yPt / page.hPt) * 100;
  const width = (rect.wPt / page.wPt) * 100;
  const height = (rect.hPt / page.hPt) * 100;
  return { left: `${left}%`, top: `${top}%`, width: `${width}%`, height: `${height}%` };
}

function fallbackBoxesForFields(fields) {
  // Simple grid fallback: list placeholders vertically.
  return fields.map((f, idx) => ({
    id: f?.mapping?.placeholderId || `fallback:${f.id}`,
    kind: f.type === 'image' ? 'image' : 'body',
    box: { xPt: 60, yPt: 60 + idx * 50, wPt: 840, hPt: f.type === 'image' ? 120 : 40 },
  }));
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

function bulletsToLines(arr) {
  if (!Array.isArray(arr)) return '';
  return arr.map((s) => String(s ?? '').trim()).filter(Boolean).map((s) => `• ${s}`).join('\n');
}

function teamMembersToLines(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return '';
  // Simple fixed-width-ish columns using spacing. Exact table fidelity in PPT is handled in generator.
  // Preview uses template box/style; we keep it legible and deterministic.
  const lines = [];
  for (const r of rows) {
    const name = String(r?.name ?? '').trim();
    const role = String(r?.role ?? '').trim();
    if (!name && !role) continue;
    if (name && role) lines.push(`${name}    ${role}`);
    else lines.push(name || role);
  }
  return lines.join('\n');
}

function resolveDisplayTextForField({ slideStep, field, wizardData }) {
  // Special cases for Skill Factory Slide 1 requirements.
  if (field?.id === 'dateRangeStart' || field?.id === 'dateRangeEnd') {
    const start = resolveValueForStep(wizardData, slideStep, 'dateRangeStart');
    const end = resolveValueForStep(wizardData, slideStep, 'dateRangeEnd');
    return formatDateRangeDdMmmYyyy(start, end);
  }

  const value = resolveValueForStep(wizardData, slideStep, field?.id);

  if (field?.type === 'bullets') return bulletsToLines(value);
  if (field?.type === 'table') return teamMembersToLines(value);

  return value;
}

function templateFontStack(themeFonts) {
  const major = themeFonts?.major;
  const minor = themeFonts?.minor;
  const fallback = themeFonts?.fallbackStack || 'Helvetica Neue, Arial, sans-serif';
  const chosen = (major && major.trim()) || (minor && minor.trim());
  return chosen ? `"${chosen}", ${fallback}` : fallback;
}

function templateTextColor(themeColors) {
  // Prefer dk1 if available; otherwise fallback to dark gray.
  const c = themeColors?.dk1 || themeColors?.text || '#111827';
  return typeof c === 'string' && c ? c : '#111827';
}

function normalizeHexColor(c) {
  if (!c || typeof c !== 'string') return null;
  if (c === 'none') return null;
  if (c.startsWith('#')) return c;
  // if extractor emits hex without '#'
  if (/^[0-9a-fA-F]{6}$/.test(c)) return `#${c}`;
  return c;
}

function ptToPx(pt) {
  // PowerPoint points map well to CSS px for on-screen preview; keep 1pt ~= 1px for fidelity.
  // We intentionally do NOT scale (previous behavior used *0.9) because the request requires exact template typography.
  if (typeof pt !== 'number' || Number.isNaN(pt)) return undefined;
  return `${Math.max(1, pt)}px`;
}

function normalizeLineHeight(lineHeight) {
  // Template extractor may emit:
  // - unitless number (e.g., 1.1)
  // - string (e.g., "1.1" or "14px")
  // - numeric point value (rare) that should be treated as px-equivalent for preview
  if (lineHeight == null) return undefined;

  if (typeof lineHeight === 'number') {
    if (lineHeight > 0 && lineHeight <= 3) return lineHeight; // treat as multiplier
    return ptToPx(lineHeight); // treat as pt
  }

  if (typeof lineHeight === 'string') {
    const s = lineHeight.trim();
    if (!s) return undefined;
    if (/^[0-9]*\.?[0-9]+$/.test(s)) {
      const n = Number(s);
      if (!Number.isNaN(n)) {
        if (n > 0 && n <= 3) return n;
        return ptToPx(n);
      }
    }
    return s;
  }

  return undefined;
}

function normalizeFontWeight(fontWeight) {
  if (fontWeight == null) return undefined;
  if (typeof fontWeight === 'number') return fontWeight;
  if (typeof fontWeight === 'string') return fontWeight;
  return undefined;
}

function renderFixedShape(shape, rectCss, templateIndex) {
  const fill = normalizeHexColor(shape?.fill);
  const stroke = normalizeHexColor(shape?.stroke);
  const strokeWidthPt = typeof shape?.strokeWidthPt === 'number' ? shape.strokeWidthPt : 0;

  if (shape?.shapeType === 'picture' && shape?.assetId) {
    const url = resolveTemplateAssetUrl(templateIndex, shape.assetId);
    if (!url) return null;
    return (
      <img
        key={shape.id}
        src={url}
        alt={shape.id}
        style={{
          ...rectCss,
          position: 'absolute',
          objectFit: 'contain',
          border: 'none',
        }}
      />
    );
  }

  return (
    <div
      key={shape.id}
      className="ocean-slide-shape fixed"
      style={{
        ...rectCss,
        background: fill || 'rgba(17, 24, 39, 0.02)',
        border: stroke ? `${Math.max(1, strokeWidthPt)}px solid ${stroke}` : 'none',
        zIndex: shape?.zIndex || 0,
      }}
      title={shape.id}
    />
  );
}

// PUBLIC_INTERFACE
export default function SlidePreview({ slideStep, templateModel, extractedTemplate, wizardData }) {
  /** Render one slide preview using exact template coordinates when available; otherwise falls back. */
  const page = useMemo(() => getPageSize(templateModel), [templateModel]);

  const templateIndex = useMemo(() => buildTemplateIndex(templateModel, extractedTemplate), [templateModel, extractedTemplate]);

  const isGlobalFirst = slideStep?.slideType === 'global_first';
  const isGlobalLast = slideStep?.slideType === 'global_last';

  const fields = useMemo(() => {
    // New flow uses slideStep.fields (already flattened), but keep compatibility fallback.
    if (Array.isArray(slideStep?.fields)) return slideStep.fields;
    const out = [];
    for (const sec of slideStep?.sections || []) {
      for (const f of sec?.fields || []) out.push(f);
    }
    return out;
  }, [slideStep]);

  const layoutId = slideStep?.layoutId;
  const layout = useMemo(() => getLayout(templateIndex, layoutId), [templateIndex, layoutId]);

  const placeholders = useMemo(() => {
    /**
     * Locked slides policy:
     * - Global First: render ONLY GF_DATE using extracted geometry (no other overlays).
     * - Global Last: locked to background-only (no overlays).
     */
    let list = [];
    if (layout?.placeholders?.length) list = layout.placeholders;

    if (isGlobalFirst) {
      // Global First must show ONLY the date overlay at the exact template GF_DATE box.
      // If GF_DATE cannot be found in the extracted template, render no overlay at all
      // (prevents any fallback/misplaced date overlay).
      const precise = getTemplatePlaceholder(templateIndex, 'GF_DATE');
      if (!precise?.box) return [];
      return [{ id: 'GF_DATE' }];
    }

    if (isGlobalLast) {
      return [];
    }

    // Other slides: fallback is allowed.
    if (!list.length) return fallbackBoxesForFields(fields);
    return list;
  }, [layout, fields, isGlobalFirst, isGlobalLast, templateIndex]);

  // Global First: do not render fixedShapes to avoid duplicating template background elements.
  // Global Last: background-only.
  const fixedShapes = useMemo(() => {
    if (isGlobalFirst) return [];
    if (isGlobalLast) return [];
    const list = [];
    if (Array.isArray(layout?.fixedShapes)) list.push(...layout.fixedShapes);
    return list;
  }, [layout, isGlobalFirst, isGlobalLast]);

  /**
   * Suppress all warning/notice rendering in UI.
   * We still compute warnings to help developers debug (console only), but we never render them
   * and we do not add any visible "placeholder notice" overlays in the preview.
   */
  const requiredWarnings = useMemo(
    () => validateSlideRequiredFields({ slideStep, wizardData, templateIndex }),
    [slideStep, wizardData, templateIndex]
  );

  const missingPlaceholderWarnings = useMemo(() => {
    const out = [];
    const fs = Array.isArray(slideStep?.fields) ? slideStep.fields : [];
    for (const f of fs) {
      const placeholderId = f?.mapping?.placeholderId;
      if (!placeholderId) continue;
      const exists = Boolean(getTemplatePlaceholder(templateIndex, placeholderId));
      if (!exists) {
        out.push({
          fieldId: f.id,
          placeholderId,
          kind: f.type === 'image' ? 'image' : 'text',
          message: `Mapped placeholder not found in extracted template: ${placeholderId}`,
        });
      }
    }
    return out;
  }, [slideStep, templateIndex]);

  const warnings = useMemo(() => [...requiredWarnings, ...missingPlaceholderWarnings], [requiredWarnings, missingPlaceholderWarnings]);

  useEffect(() => {
    if (warnings.length > 0) {
      // Keep warnings only in non-visual logs (developer console).
      // This satisfies: "Keep warnings (if any) only in non-visual logs or developer console, not in UI".
      // eslint-disable-next-line no-console
      console.warn('[SlidePreview] non-blocking warnings (suppressed in UI):', warnings);
    }
  }, [warnings]);

  const fontFamily = templateFontStack(templateModel?.theme?.fonts);
  const textColor = templateTextColor(templateModel?.theme?.colors);

  // Render extracted layout background as an image when available.
  // Global Last is locked to a fixed background image and must render identically in preview and PPT generation.
  const backgroundUrl = useMemo(() => {
    // Locked backgrounds must use the known reference images to match PPT export.
    if (isGlobalFirst) return '/assets/global_first_background.png';
    if (isGlobalLast) return '/assets/global_last_background.png';

    if (!layout?.background?.assetId) return null;
    return resolveTemplateAssetUrl(templateIndex, layout.background.assetId);
  }, [layout, templateIndex, isGlobalLast, isGlobalFirst]);

  // Locked slides should not crop/distort backgrounds.
  const backgroundFit = isGlobalFirst || isGlobalLast ? 'contain' : 'cover';

  // Avoid leaking object URLs when user provides images (File inputs).
  const [objectUrlByPlaceholderId, setObjectUrlByPlaceholderId] = useState({});

  // Create object URLs in an effect (never during render).
  useEffect(() => {
    const nextUrls = {};
    for (const ph of placeholders) {
      const mappedField = fields.find((f) => (f?.mapping?.placeholderId || '') === ph.id);
      const value = mappedField ? resolveValueForStep(wizardData, slideStep, mappedField.id) : null;
      const hasUserFile = value && typeof File !== 'undefined' && value instanceof File;

      if (!hasUserFile) continue;
      if (objectUrlByPlaceholderId[ph.id]) continue;

      nextUrls[ph.id] = URL.createObjectURL(value);
    }

    if (Object.keys(nextUrls).length > 0) {
      setObjectUrlByPlaceholderId((prev) => ({ ...prev, ...nextUrls }));
    }

    return undefined;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [placeholders, fields, wizardData, slideStep?.key]);

  useEffect(() => {
    // Cleanup all object URLs when slide changes/unmounts.
    return () => {
      for (const url of Object.values(objectUrlByPlaceholderId)) {
        try {
          URL.revokeObjectURL(url);
        } catch {
          // ignore
        }
      }
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [slideStep?.key]);

  // Global Last must be purely the provided fixed background image with no overlays.
  if (isGlobalLast) {
    return (
      <div className="ocean-slide-canvas" aria-label={`Preview for ${slideStep?.title || 'slide'}`}>
        <div className="ocean-slide-layer" style={{ fontFamily, color: textColor }}>
          <img
            src="/assets/global_last_background.png"
            alt="slide background"
            style={{
              position: 'absolute',
              inset: 0,
              width: '100%',
              height: '100%',
              objectFit: 'contain', // no crop, no distortion (letterbox if needed)
              objectPosition: 'center center',
              zIndex: 0,
              transform: 'translateZ(0)',
            }}
          />
        </div>
      </div>
    );
  }

  return (
    <div className="ocean-slide-canvas" aria-label={`Preview for ${slideStep?.title || 'slide'}`}>
      <div className="ocean-slide-layer" style={{ fontFamily, color: textColor }}>
        {backgroundUrl && (
          <img
            src={backgroundUrl}
            alt="slide background"
            style={{
              position: 'absolute',
              inset: 0,
              width: '100%',
              height: '100%',
              objectFit: backgroundFit,
              objectPosition: 'center center',
              zIndex: 0,
            }}
          />
        )}

        {fixedShapes.map((sh) => {
          if (!sh?.box) return null;
          const rectCss = scaleRect(sh.box, page);
          return renderFixedShape(sh, rectCss, templateIndex);
        })}

        {placeholders.map((ph) => {
          // Prefer exact coordinates from extracted placeholders by id.
          const precise = getTemplatePlaceholder(templateIndex, ph.id);

          // Pixel-perfect rule for Global First: use ONLY the extracted box.
          const box = isGlobalFirst ? precise?.box : precise?.box || ph.box;

          const rect = box ? scaleRect(box, page) : { left: '5%', top: '5%', width: '90%', height: '12%' };
          const mappedField = fields.find((f) => (f?.mapping?.placeholderId || '') === ph.id);

          const value = mappedField ? resolveDisplayTextForField({ slideStep, field: mappedField, wizardData }) : null;

          const isImage = (precise?.kind || ph.kind) === 'image' || mappedField?.type === 'image';

          if (isImage) {
            // Priority order for image rendering:
            // 1) user-provided file (object URL created in effect)
            // 2) template default asset referenced by placeholder (if any)
            const hasUserFile = value && typeof File !== 'undefined' && value instanceof File;

            let src = null;
            if (hasUserFile) {
              src = objectUrlByPlaceholderId[ph.id] || null;
            } else if (precise?.assetId) {
              src = resolveTemplateAssetUrl(templateIndex, precise.assetId);
            }

            return (
              <div
                key={ph.id}
                className="ocean-slide-shape placeholder"
                style={{
                  ...rect,
                  zIndex: ph.zIndex || 1,
                  transform: ph.rotationDeg ? `rotate(${ph.rotationDeg}deg)` : undefined,
                  opacity: typeof ph.opacity === 'number' ? ph.opacity : 1,
                }}
                title={ph.id}
              >
                {src ? (
                  <img src={src} alt={ph.id} style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
                ) : (
                  <div style={{ width: '100%', height: '100%' }} aria-hidden="true" />
                )}
              </div>
            );
          }

          const templateDefault = typeof precise?.text === 'string' && precise.text.trim() ? precise.text : null;

          const hasValue = value !== undefined && value !== null && !(typeof value === 'string' && value.trim() === '');

          let labelText;
          // Global First: GF_DATE is the only editable overlay. Use DD MMM YYYY format and template style/geometry.
          // Important: do not depend on the wizard field id here; always apply to the GF_DATE placeholder itself.
          if (isGlobalFirst && ph.id === 'GF_DATE') {
            const formatted = hasValue ? formatDdMmmYyyy(value) : '';
            labelText = formatted;
          } else {
            labelText = hasValue ? String(value) : templateDefault || '';
          }

          const style = precise?.style || ph?.style || null;
          const align = style?.align || 'left';

          const resolvedFontFamily = style?.fontFamily ? `"${style.fontFamily}", ${fontFamily}` : fontFamily;

          return (
            <div
              key={ph.id}
              className="ocean-slide-shape placeholder"
              style={{
                ...rect,
                zIndex: ph.zIndex || 1,
                transform: ph.rotationDeg ? `rotate(${ph.rotationDeg}deg)` : undefined,
                opacity: typeof ph.opacity === 'number' ? ph.opacity : 1,
                display: 'flex',
                alignItems: 'flex-start',
                justifyContent: align === 'center' ? 'center' : align === 'right' ? 'flex-end' : 'flex-start',
                padding: 0,
                boxSizing: 'border-box',
              }}
              title={ph.id}
            >
              <div
                className="shape-label"
                style={{
                  width: '100%',
                  whiteSpace: 'pre-wrap',
                  fontFamily: resolvedFontFamily,
                  fontSize: style?.fontSizePt ? ptToPx(style.fontSizePt) : undefined,
                  fontWeight: normalizeFontWeight(style?.fontWeight),
                  color: normalizeHexColor(style?.color) || (isGlobalFirst ? undefined : textColor),
                  textAlign: align,
                  lineHeight: normalizeLineHeight(style?.lineHeight),
                }}
              >
                {labelText}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}
