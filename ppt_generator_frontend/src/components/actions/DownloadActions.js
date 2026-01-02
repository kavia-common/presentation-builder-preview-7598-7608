import React, { useMemo, useState } from 'react';
import { useWizard } from '../../state/wizardContext';
import { generatePptx } from '../../services/pptGenerator';

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// PUBLIC_INTERFACE
export default function DownloadActions() {
  /** UI actions for generating and downloading the PPTX client-side. */
  const { state } = useWizard();
  const [status, setStatus] = useState('idle'); // idle | generating | done | error
  const [error, setError] = useState(null);

  const title = useMemo(() => {
    const t =
      state.wizardData?.globalFirst?.title ||
      state.formData?.title ||
      state.formData?.deckTitle ||
      state.formData?.presentationTitle;
    return typeof t === 'string' ? t : '';
  }, [state.wizardData, state.formData]);

  return (
    <div>
      <div className="ocean-help" style={{ marginBottom: 10 }}>
        PPT generation is fully client-side. If template coordinates are incomplete, content is still mapped using stable placeholder IDs.
      </div>

      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
        <button
          type="button"
          className="ocean-btn ocean-btn-primary"
          onClick={async () => {
            setStatus('generating');
            setError(null);
            try {
              const { blob, fileName } = await generatePptx({
                templateModel: state.templateModel,
                flowSchema: state.flowSchema,
                orderedSlides: state.orderedSlides,
                wizardData: state.wizardData,
              });
              downloadBlob(blob, fileName);
              setStatus('done');
            } catch (e) {
              setStatus('error');
              setError(String(e?.message || e));
            }
          }}
          disabled={status === 'generating'}
        >
          {status === 'generating' ? 'Generating…' : 'Download PPT'}
        </button>

        <div className="ocean-help">
          Filename from title: <span className="ocean-kbd">{title || '(empty)'}</span>
        </div>
      </div>

      {status === 'error' && (
        <div className="ocean-error" role="alert" style={{ marginTop: 10 }}>
          {error || 'Failed to generate PPTX.'}
        </div>
      )}

      {status === 'done' && (
        <div className="ocean-help" style={{ marginTop: 10 }}>
          Download started. You can iterate fields and download again.
        </div>
      )}
    </div>
  );
}
