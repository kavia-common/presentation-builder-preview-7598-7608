import React, { useMemo } from 'react';
import Stepper from './Stepper';
import FieldRenderer from './FieldRenderer';
import PreviewCarousel from '../preview/PreviewCarousel';
import DownloadActions from '../actions/DownloadActions';
import { useWizard } from '../../state/wizardContext';
import { buildTemplateIndex } from '../../services/schemaLoader';

function groupSteps(orderedSlides) {
  const globalFirst = orderedSlides.find((s) => s.slideType === 'global_first');
  const globalLast = orderedSlides.find((s) => s.slideType === 'global_last');
  const sfSteps = orderedSlides.filter((s) => s.kind === 'skillFactory');
  const groups = [];
  const byId = new Map();
  for (const s of sfSteps) {
    if (!byId.has(s.groupId)) {
      byId.set(s.groupId, { groupId: s.groupId, groupIndex: s.groupIndex, label: s.groupLabel, steps: [] });
      groups.push(byId.get(s.groupId));
    }
    byId.get(s.groupId).steps.push(s);
  }
  groups.sort((a, b) => a.groupIndex - b.groupIndex);
  for (const g of groups) {
    const order = ['sf1', 'sf2', 'sf3', 'sf4'];
    g.steps.sort((a, b) => order.indexOf(a.slideType) - order.indexOf(b.slideType));
  }
  return { globalFirst, groups, globalLast };
}

function flowLabelForSubstep(step) {
  // Ensures sidebar substep label matches required SF-1..SF-4 naming.
  const type = step.slideType;
  if (type === 'sf1') return 'SF-1';
  if (type === 'sf2') return 'SF-2';
  if (type === 'sf3') return 'SF-3';
  if (type === 'sf4') return 'SF-4';
  return step.title;
}

function Sidebar({
  globalFirst,
  groups,
  globalLast,
  currentStep,
  previewStepIndex,
  onSelect,
  onAdd,
  onRemove,
  onMove,
  canRemove,
  canReorder,
}) {
  return (
    <div className="ocean-card" style={{ padding: 16 }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
        <div style={{ fontWeight: 800, fontSize: 13 }}>Wizard Steps</div>
        <button type="button" className="ocean-btn ocean-btn-secondary" onClick={onAdd}>
          + Add Skill Factory
        </button>
      </div>

      <div className="ocean-divider" />

      <div style={{ display: 'grid', gap: 10 }}>
        {globalFirst && (
          <button
            type="button"
            className={`ocean-btn ${currentStep === 0 ? 'ocean-btn-primary' : 'ocean-btn-ghost'}`}
            onClick={() => onSelect(0)}
          >
            {globalFirst.title}
          </button>
        )}

        {groups.map((g, idx) => (
          <div key={g.groupId} className="ocean-card" style={{ padding: 12, borderRadius: 12, borderStyle: 'dashed' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 10 }}>
              <div style={{ fontWeight: 800, fontSize: 13 }}>{g.label}</div>
              <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                <button
                  type="button"
                  className="ocean-btn ocean-btn-ghost"
                  onClick={() => onMove(g.groupId, -1)}
                  disabled={!canReorder || idx === 0}
                  title="Move up"
                >
                  ↑
                </button>
                <button
                  type="button"
                  className="ocean-btn ocean-btn-ghost"
                  onClick={() => onMove(g.groupId, 1)}
                  disabled={!canReorder || idx === groups.length - 1}
                  title="Move down"
                >
                  ↓
                </button>
                <button
                  type="button"
                  className="ocean-btn ocean-btn-danger"
                  onClick={() => onRemove(g.groupId)}
                  disabled={!canRemove}
                  title={canRemove ? 'Remove this group' : 'At least one Skill Factory is recommended'}
                >
                  Remove
                </button>
              </div>
            </div>

            <div style={{ display: 'grid', gap: 8, marginTop: 10 }}>
              {g.steps.map((s) => (
                <button
                  key={s.key}
                  type="button"
                  className={`ocean-btn ${currentStep === s.stepIndex ? 'ocean-btn-primary' : 'ocean-btn-ghost'}`}
                  onClick={() => onSelect(s.stepIndex)}
                  style={{ justifyContent: 'space-between' }}
                >
                  <span style={{ fontSize: 13 }}>{flowLabelForSubstep(s)}</span>
                  <span
                    className="ocean-badge"
                    style={{ background: 'rgba(37,99,235,0.08)', borderColor: 'rgba(37,99,235,0.15)' }}
                  >
                    {s.slideType.toUpperCase()}
                  </span>
                </button>
              ))}
            </div>
          </div>
        ))}

        {globalLast && (
          <button
            type="button"
            className={`ocean-btn ${currentStep === globalLast.stepIndex ? 'ocean-btn-primary' : 'ocean-btn-ghost'}`}
            onClick={() => onSelect(globalLast.stepIndex)}
          >
            {globalLast.title}
          </button>
        )}

        <button
          type="button"
          className={`ocean-btn ${currentStep === previewStepIndex ? 'ocean-btn-primary' : 'ocean-btn-ghost'}`}
          onClick={() => onSelect(previewStepIndex)}
        >
          Preview
        </button>
      </div>
    </div>
  );
}

// PUBLIC_INTERFACE
export default function WizardLayout() {
  /** Main wizard page: template-driven grouped wizard + preview + download actions. */
  const { state, actions } = useWizard();

  /**
   * Hook-order safety:
   * - All hooks must be called every render in the same order.
   * - Therefore: compute safe defaults and declare all hooks BEFORE any early-return-like branching.
   * - Render is selected via `body` and returned once.
   */

  const templateIndex = useMemo(
    () => buildTemplateIndex(state.templateModel, state.extractedTemplate),
    [state.templateModel, state.extractedTemplate]
  );
  const templateDrivenPositioningActive = Boolean(templateIndex?.isTemplatePositioningActive);

  const orderedSlides = Array.isArray(state.orderedSlides) ? state.orderedSlides : [];
  const previewStepIndex = state.previewStepIndex ?? orderedSlides.length;

  const indexed = orderedSlides.map((s, idx) => ({ ...s, stepIndex: idx }));
  const { globalFirst, groups, globalLast } = groupSteps(indexed);

  const currentStep = Number.isFinite(state.currentStep) ? state.currentStep : 0;
  const isPreview = currentStep === previewStepIndex;
  const stepData = !isPreview ? orderedSlides[currentStep] : null;

  const showPartialBanner = Boolean(state.templateModel?.meta?.partial || state.templateModel?.meta?.isPartial);

  const stepErrors = stepData ? state.validation?.errors?.[stepData.key] || {} : {};
  const stepTouched = stepData ? Boolean(state.validation?.touched?.[stepData.key]) : false;

  const steps = useMemo(() => {
    const list = [];
    list.push({ key: 'global_first', label: 'Global First' });

    const factories = groups || [];
    for (const g of factories) {
      for (const s of g.steps) list.push({ key: s.key, label: `${g.label} ${s.slideType.toUpperCase()}` });
    }

    list.push({ key: 'global_last', label: 'Global Last' });
    list.push({ key: 'preview', label: 'Preview' });

    return list;
  }, [groups]);

  const stepperIndexToWizardStep = useMemo(() => {
    const map = [];
    map.push(0);

    const factories = groups || [];
    for (const g of factories) {
      for (const s of g.steps) map.push(s.stepIndex);
    }

    if (globalLast) map.push(globalLast.stepIndex);
    map.push(previewStepIndex);

    return map;
  }, [groups, globalLast, previewStepIndex]);

  const canRemove = (state.wizardData?.skillFactories?.length || 0) > 1;
  const canReorder = (state.wizardData?.skillFactories?.length || 0) > 1;

  let body = null;

  if (state.status === 'loading' || state.status === 'idle') {
    body = (
      <div className="ocean-container">
        <div className="ocean-card ocean-card-body">Loading schemas…</div>
      </div>
    );
  } else if (state.status === 'error') {
    body = (
      <div className="ocean-container">
        <div className="ocean-card ocean-card-body">
          <div style={{ color: 'var(--ocean-error)', fontWeight: 700 }}>Failed to load.</div>
          <div className="ocean-help">{state.error}</div>
          <div className="ocean-help">
            Ensure <span className="ocean-kbd">public/assets/*.json</span> exists.
          </div>
        </div>
      </div>
    );
  } else {
    body = (
      <div className="ocean-app">
        <div className="ocean-container">
          <div className="ocean-header">
            <div>
              <h1 className="ocean-title">Presentation Builder</h1>
              <p className="ocean-subtitle">Template-driven wizard → live preview → PPT download (client-side).</p>
              <div className="ocean-help">
                Template-driven positioning:{' '}
                <span className="ocean-kbd">{templateDrivenPositioningActive ? 'active' : 'fallback'}</span>
              </div>
            </div>
            <div className="ocean-toolbar" aria-label="Global actions">
              <button type="button" className="ocean-btn ocean-btn-ghost" onClick={actions.restoreDefaults}>
                Restore Defaults
              </button>
              <button type="button" className="ocean-btn ocean-btn-danger" onClick={actions.clearAll}>
                Clear All
              </button>
            </div>
          </div>

          <Stepper
            steps={steps}
            current={Math.max(0, stepperIndexToWizardStep.indexOf(currentStep))}
            onSelect={(idx) => actions.setStep(stepperIndexToWizardStep[idx] ?? 0)}
          />

          <div className="ocean-grid" style={{ marginTop: 16 }}>
            <div style={{ display: 'grid', gap: 16 }}>
              <Sidebar
                globalFirst={globalFirst}
                groups={groups}
                globalLast={globalLast}
                currentStep={currentStep}
                previewStepIndex={previewStepIndex}
                onSelect={actions.setStep}
                onAdd={actions.addSkillFactory}
                onRemove={actions.removeSkillFactory}
                onMove={actions.moveSkillFactory}
                canRemove={canRemove}
                canReorder={canReorder}
              />

              {!isPreview && (
                <div className="ocean-help">
                  Required flow enforced: <span className="ocean-kbd">Global First</span> +{' '}
                  <span className="ocean-kbd">Skill Factory (SF-1..SF-4)</span> × N +{' '}
                  <span className="ocean-kbd">Global Last</span>. Preview is blocked until required fields are complete.
                </div>
              )}
            </div>

            <div className="ocean-card">
              <div className="ocean-card-header">
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
                  <h2 className="ocean-step-title" style={{ margin: 0 }}>
                    {isPreview ? 'Preview & Download' : stepData?.title || 'Step'}
                  </h2>
                  {!isPreview && <span className="ocean-badge">Layout: {stepData?.layoutId || 'unknown'}</span>}
                </div>
                <div className="ocean-divider" />
              </div>

              <div className="ocean-card-body">
                {isPreview ? (
                  <div>
                    {showPartialBanner && (
                      <div className="ocean-banner" role="status" aria-live="polite" style={{ marginBottom: 12 }}>
                        <div>
                          <strong>Template Not Fully Extracted</strong>
                          <div style={{ fontSize: 13, marginTop: 4 }}>
                            Preview and PPT generation use a safe fallback layout. Placeholder IDs stay stable for future
                            refinement.
                          </div>
                          {state.refinedLaterNote && (
                            <div style={{ fontSize: 13, marginTop: 6 }}>Note saved: “Refine Template Later”.</div>
                          )}
                        </div>
                        <button type="button" className="ocean-btn ocean-btn-secondary" onClick={actions.markRefineLater}>
                          Refine Template Later
                        </button>
                      </div>
                    )}

                    {!showPartialBanner && !templateDrivenPositioningActive && (
                      <div className="ocean-banner" role="status" aria-live="polite" style={{ marginBottom: 12 }}>
                        <div>
                          <strong>Template-driven positioning not available</strong>
                          <div style={{ fontSize: 13, marginTop: 4 }}>
                            Extracted template JSON is minimal. Preview and PPT generation are using fallback coordinates.
                          </div>
                        </div>
                      </div>
                    )}

                    {(state.previewGateErrors || []).length > 0 && (
                      <div className="ocean-error" role="alert" style={{ marginBottom: 12 }}>
                        <div style={{ fontWeight: 800, marginBottom: 6 }}>Preview is blocked:</div>
                        <ul style={{ margin: 0, paddingLeft: 18 }}>
                          {state.previewGateErrors.map((e) => (
                            <li key={e}>{e}</li>
                          ))}
                        </ul>
                        <div className="ocean-help" style={{ marginTop: 8 }}>
                          Fill required fields, then click Preview again.
                        </div>
                      </div>
                    )}

                    <DownloadActions />
                  </div>
                ) : (
                  <div>
                    <div style={{ display: 'grid', gap: 14 }}>
                      {(stepData?.fields || []).map((f) => {
                        const localError = stepTouched ? stepErrors?.[f.id]?.[0] : null;

                        return (
                          <div key={f.id}>
                            <FieldRenderer
                              field={f}
                              value={resolveFieldValue(state.wizardData, stepData, f.id)}
                              onChange={(val) => actions.setFieldForStep(stepData, f.id, val)}
                            />
                            {localError && (
                              <div className="ocean-error" role="alert">
                                {localError}
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>

                    <div style={{ display: 'flex', gap: 10, justifyContent: 'space-between', marginTop: 14 }}>
                      <button
                        type="button"
                        className="ocean-btn ocean-btn-ghost"
                        onClick={actions.back}
                        disabled={currentStep === 0}
                      >
                        Back
                      </button>
                      <button type="button" className="ocean-btn ocean-btn-primary" onClick={actions.next}>
                        Next
                      </button>
                    </div>

                    {stepTouched && Object.keys(stepErrors || {}).length > 0 && (
                      <div className="ocean-help" style={{ marginTop: 10 }}>
                        Fix highlighted fields to continue.
                      </div>
                    )}
                  </div>
                )}
              </div>
            </div>

            <div className="ocean-card">
              <div className="ocean-card-header">
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
                  <h2 className="ocean-step-title" style={{ margin: 0 }}>
                    Live Preview
                  </h2>
                  <span className="ocean-badge">Slide-by-slide</span>
                </div>
                <div className="ocean-divider" />
              </div>
              <div className="ocean-card-body">
                <PreviewCarousel
                  orderedSlides={orderedSlides}
                  templateModel={state.templateModel}
                  extractedTemplate={state.extractedTemplate}
                  wizardData={state.wizardData}
                  currentWizardStep={state.currentStep}
                />
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return body;
}

function resolveFieldValue(wizardData, step, fieldId) {
  if (step?.dataPath?.scope === 'globalFirst') return wizardData?.globalFirst?.[fieldId];
  if (step?.dataPath?.scope === 'globalLast') return wizardData?.globalLast?.[fieldId];
  if (step?.dataPath?.scope === 'skillFactories') {
    const group = (wizardData?.skillFactories || []).find((g) => g.id === step.dataPath.groupId);
    return group?.slides?.[step.dataPath.slideKey]?.[fieldId];
  }
  return undefined;
}
