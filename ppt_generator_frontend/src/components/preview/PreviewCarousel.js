import React, { useMemo, useState } from 'react';
import SlidePreview from './SlidePreview';

// PUBLIC_INTERFACE
export default function PreviewCarousel({ orderedSlides, templateModel, wizardData, currentWizardStep }) {
  /** Slide-by-slide preview with navigation, driven by the actual ordered flow steps. */
  const slides = useMemo(() => (Array.isArray(orderedSlides) ? orderedSlides : []), [orderedSlides]);
  const [idx, setIdx] = useState(0);

  // If user is on a slide step, sync preview to that slide index (ignore Preview step)
  React.useEffect(() => {
    if (currentWizardStep >= 0 && currentWizardStep < slides.length) setIdx(currentWizardStep);
  }, [currentWizardStep, slides.length]);

  if (!slides.length) {
    return <div className="ocean-help">No slides in current flow.</div>;
  }

  const current = slides[idx];

  return (
    <div>
      <div style={{ display: 'flex', gap: 10, justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
        <button
          type="button"
          className="ocean-btn ocean-btn-ghost"
          onClick={() => setIdx((p) => Math.max(p - 1, 0))}
          disabled={idx === 0}
        >
          Previous
        </button>

        <div style={{ textAlign: 'center' }}>
          <div style={{ fontWeight: 800, fontSize: 13 }}>{current.title || `Slide ${idx + 1}`}</div>
          <div className="ocean-help">
            {idx + 1} / {slides.length} • layout <span className="ocean-kbd">{current.layoutId || 'unknown'}</span>
          </div>
        </div>

        <button
          type="button"
          className="ocean-btn ocean-btn-ghost"
          onClick={() => setIdx((p) => Math.min(p + 1, slides.length - 1))}
          disabled={idx === slides.length - 1}
        >
          Next
        </button>
      </div>

      <SlidePreview slideStep={current} templateModel={templateModel} wizardData={wizardData} />
    </div>
  );
}
