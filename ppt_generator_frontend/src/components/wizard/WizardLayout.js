import React, { useMemo } from 'react';
import Stepper from './Stepper';
import FieldRenderer from './FieldRenderer';
import PreviewCarousel from '../preview/PreviewCarousel';
import DownloadActions from '../actions/DownloadActions';
import { useWizard } from '../../state/wizardContext';

function flattenFields(step) {
  const sections = Array.isArray(step?.sections) ? step.sections : [];
  const list = [];
  for (const sec of sections) {
    for (const f of sec.fields || []) list.push({ section: sec, field: f });
  }
  return list;
}

// PUBLIC_INTERFACE
export default function WizardLayout() {
  /** Main wizard page: stepper + form fields per slide + preview + download actions. */
  const { state, actions } = useWizard();

  // Hooks must run unconditionally; compute safe defaults even while loading/error.
  const slides = Array.isArray(state.wizardSchema?.slides) ? state.wizardSchema.slides : [];
  const previewStepIndex = slides.length;

  const steps = useMemo(() => {
    const s = slides.map((sl) => ({
      key: `slide-${sl.slideIndex}`,
      label: sl.title || `Slide ${sl.slideIndex + 1}`,
    }));
    s.push({ key: 'preview', label: 'Preview' });
    return s;
  }, [slides]);

  if (state.status === 'loading' || state.status === 'idle') {
    return (
      <div className="ocean-container">
        <div className="ocean-card ocean-card-body">Loading schemas…</div>
      </div>
    );
  }

  if (state.status === 'error') {
    return (
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
  }

  const currentStep = state.currentStep;
  const isPreview = currentStep === previewStepIndex;
  const stepData = !isPreview ? slides[currentStep] : null;

  const showPartialBanner = Boolean(state.templateModel?.meta?.partial || state.templateModel?.meta?.isPartial);

  return (
    <div className="ocean-app">
      <div className="ocean-container">
        <div className="ocean-header">
          <div>
            <h1 className="ocean-title">Presentation Builder</h1>
            <p className="ocean-subtitle">
              Multi-step wizard → live preview → PPT download (client-side).
            </p>
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

        <Stepper steps={steps} current={currentStep} onSelect={actions.setStep} />

        <div className="ocean-grid" style={{ marginTop: 16 }}>
          <div className="ocean-card">
            <div className="ocean-card-header">
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
                <h2 className="ocean-step-title" style={{ margin: 0 }}>
                  {isPreview ? 'Preview & Download' : (stepData?.title || 'Step')}
                </h2>
                {!isPreview && (
                  <span className="ocean-badge">
                    Layout: {stepData?.layoutId || 'unknown'}
                  </span>
                )}
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
                          Preview and PPT generation use a safe fallback layout. Placeholder IDs stay stable for future refinement.
                        </div>
                        {state.refinedLaterNote && (
                          <div style={{ fontSize: 13, marginTop: 6 }}>
                            Note saved: “Refine Template Later”.
                          </div>
                        )}
                      </div>
                      <button type="button" className="ocean-btn ocean-btn-secondary" onClick={actions.markRefineLater}>
                        Refine Template Later
                      </button>
                    </div>
                  )}

                  <DownloadActions />
                </div>
              ) : (
                <div>
                  {(stepData?.sections || []).map((sec) => {
                    const fields = (sec.fields || []);
                    return (
                      <div key={sec.id} style={{ marginBottom: 14 }}>
                        <div style={{ fontWeight: 800, fontSize: 13 }}>{sec.label}</div>
                        {sec.description && <div className="ocean-help">{sec.description}</div>}
                        <div style={{ marginTop: 10 }}>
                          {fields.map((f) => (
                            <FieldRenderer
                              key={f.id}
                              field={f}
                              value={state.formData[f.id]}
                              onChange={(val) => actions.setField(f.id, val)}
                            />
                          ))}
                        </div>
                        <div className="ocean-divider" />
                      </div>
                    );
                  })}

                  <div style={{ display: 'flex', gap: 10, justifyContent: 'space-between' }}>
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
                </div>
              )}
            </div>
          </div>

          <div className="ocean-card">
            <div className="ocean-card-header">
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
                <h2 className="ocean-step-title" style={{ margin: 0 }}>Live Preview</h2>
                <span className="ocean-badge">Slide-by-slide</span>
              </div>
              <div className="ocean-divider" />
            </div>
            <div className="ocean-card-body">
              <PreviewCarousel
                wizardSchema={state.wizardSchema}
                templateModel={state.templateModel}
                formData={state.formData}
                currentWizardStep={state.currentStep}
              />
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
