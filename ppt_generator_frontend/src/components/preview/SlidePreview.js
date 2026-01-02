import React, { useEffect, useMemo, useState } from 'react';
import {
  buildTemplateIndex,
  getLayout,
  getTemplatePlaceholder,
  resolveTemplateAssetUrl,
  validateSlideRequiredFields,
} from '../../services/schemaLoader';

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
    if (layout?.placeholders?.length) return layout.placeholders;
    // fallback: create boxes from current step fields (still stable by placeholderId when provided)
    return fallbackBoxesForFields(fields);
  }, [layout, fields]);

  const fixedShapes = useMemo(() => {
    // Template supports fixed shapes on layouts (and slide overrides); for now, layout-only.
    const list = [];
    if (Array.isArray(layout?.fixedShapes)) list.push(...layout.fixedShapes);
    // slide overrides could be in templateModel.slides[index], but we don't have slide index mapping here.
    return list;
  }, [layout]);

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

  const fontFamily = templateFontStack(templateModel?.theme?.fonts);
  const textColor = templateTextColor(templateModel?.theme?.colors);

  // Render extracted layout background as an image when available.
  const backgroundUrl = useMemo(() => {
    if (!layout?.background?.assetId) return null;
    return resolveTemplateAssetUrl(templateIndex, layout.background.assetId);
  }, [layout, templateIndex]);

  // Avoid leaking object URLs when user provides images (File inputs).
  const [objectUrlByPlaceholderId, setObjectUrlByPlaceholderId] = useState({});

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

  return (
    <div className="ocean-slide-canvas" aria-label={`Preview for ${slideStep?.title || 'slide'}`}>
      <div className="ocean-slide-layer" style={{ fontFamily, color: textColor }}>
        {backgroundUrl && (
          <img
            src={backgroundUrl}
            alt="slide background"
            style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', objectFit: 'cover', zIndex: 0 }}
          />
        )}

        {warnings.length > 0 && (
          <div
            style={{
              position: 'absolute',
              left: 10,
              top: 10,
              right: 10,
              zIndex: 999,
              background: 'rgba(245, 158, 11, 0.12)',
              border: '1px solid rgba(245, 158, 11, 0.25)',
              borderRadius: 10,
              padding: '8px 10px',
              color: '#92400e',
              fontSize: 12,
              fontWeight: 700,
            }}
            role="status"
            aria-live="polite"
          >
            Preview warnings (non-blocking):
            <ul style={{ margin: '6px 0 0', paddingLeft: 18, fontWeight: 600 }}>
              {warnings.slice(0, 6).map((w) => (
                <li key={`${w.fieldId}:${w.placeholderId || ''}`}>{w.message}</li>
              ))}
            </ul>
          </div>
        )}

        {fixedShapes.map((sh) => {
          if (!sh?.box) return null;
          const rectCss = scaleRect(sh.box, page);
          return renderFixedShape(sh, rectCss, templateIndex);
        })}

        {placeholders.map((ph) => {
          // Prefer exact coordinates from extracted placeholders by id if present.
          const precise = getTemplatePlaceholder(templateIndex, ph.id);
          const box = precise?.box || ph.box;

          const rect = box ? scaleRect(box, page) : { left: '5%', top: '5%', width: '90%', height: '12%' };
          const mappedField = fields.find((f) => (f?.mapping?.placeholderId || '') === ph.id);
          const value = mappedField ? resolveValueForStep(wizardData, slideStep, mappedField.id) : null;

          const isImage = (precise?.kind || ph.kind) === 'image' || mappedField?.type === 'image';

          if (isImage) {
            // Priority order for image rendering:
            // 1) user-provided file
            // 2) template default asset referenced by placeholder (if any)
            const hasUserFile = value && typeof File !== 'undefined' && value instanceof File;

            let src = null;
            if (hasUserFile) {
              src = objectUrlByPlaceholderId[ph.id] || null;
              if (!src) {
                const url = URL.createObjectURL(value);
                setObjectUrlByPlaceholderId((prev) => ({ ...prev, [ph.id]: url }));
                src = url;
              }
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
                  <div className="shape-label">[{ph.id}] image</div>
                )}
              </div>
            );
          }

          // If no user data exists, show template's default text (if provided); else show placeholder id.
          const templateDefault = typeof precise?.text === 'string' && precise.text.trim() ? precise.text : null;
          const labelText = value ? String(value) : templateDefault || `[${ph.id}]`;

          // Apply placeholder style when present (font size, weight, color, align).
          const style = precise?.style || ph?.style || null;
          const align = style?.align || 'left';

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
                padding: 6,
                boxSizing: 'border-box',
              }}
              title={ph.id}
            >
              <div
                className="shape-label"
                style={{
                  width: '100%',
                  whiteSpace: 'pre-wrap',
                  fontFamily: style?.fontFamily ? `"${style.fontFamily}", ${fontFamily}` : fontFamily,
                  fontSize: style?.fontSizePt ? `${Math.max(8, style.fontSizePt * 0.9)}px` : undefined,
                  fontWeight: style?.fontWeight || undefined,
                  color: normalizeHexColor(style?.color) || textColor,
                  textAlign: align,
                  lineHeight: style?.lineHeight || undefined,
                  opacity: value ? 1 : 0.82,
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
