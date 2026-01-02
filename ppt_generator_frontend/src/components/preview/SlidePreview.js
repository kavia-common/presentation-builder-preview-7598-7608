import React, { useMemo } from 'react';

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

// PUBLIC_INTERFACE
export default function SlidePreview({ slideStep, templateModel, wizardData }) {
  /** Render one slide preview as absolutely positioned boxes approximating PPT coordinates. */
  const page = useMemo(() => getPageSize(templateModel), [templateModel]);

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
  const layout = useMemo(() => {
    const layouts = Array.isArray(templateModel?.layouts) ? templateModel.layouts : [];
    return layouts.find((l) => l.id === layoutId) || null;
  }, [templateModel, layoutId]);

  const placeholders = useMemo(() => {
    if (layout?.placeholders?.length) return layout.placeholders;
    // fallback: create boxes from current step fields (still stable by placeholderId when provided)
    return fallbackBoxesForFields(fields);
  }, [layout, fields]);

  return (
    <div className="ocean-slide-canvas" aria-label={`Preview for ${slideStep?.title || 'slide'}`}>
      <div className="ocean-slide-layer">
        {placeholders.map((ph) => {
          const rect = ph.box ? scaleRect(ph.box, page) : { left: '5%', top: '5%', width: '90%', height: '12%' };
          const mappedField = fields.find((f) => (f?.mapping?.placeholderId || '') === ph.id);
          const value = mappedField ? resolveValueForStep(wizardData, slideStep, mappedField.id) : null;

          const labelText =
            ph.kind === 'image'
              ? (value ? `[${ph.id}] image selected` : `[${ph.id}] image`)
              : (value ? String(value) : `[${ph.id}]`);

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
              <div className="shape-label">{labelText}</div>
            </div>
          );
        })}
      </div>
    </div>
  );
}
