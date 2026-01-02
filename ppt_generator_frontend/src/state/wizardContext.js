import React, { createContext, useContext, useEffect, useMemo, useReducer } from 'react';

const STORAGE_KEY = 'ppt_wizard_draft_v1';

function safeJsonParse(str) {
  try {
    return JSON.parse(str);
  } catch {
    return null;
  }
}

function buildDefaultFormData(wizardSchema) {
  const slides = Array.isArray(wizardSchema?.slides) ? wizardSchema.slides : [];
  const data = {};
  for (const s of slides) {
    const sections = Array.isArray(s.sections) ? s.sections : [];
    for (const sec of sections) {
      const fields = Array.isArray(sec.fields) ? sec.fields : [];
      for (const f of fields) {
        if (f?.id) data[f.id] = f.defaultValue ?? (f.type === 'image' ? null : '');
      }
    }
  }
  // Common title fallback used for file naming
  if (data.title == null) data.title = '';
  return data;
}

const initialState = {
  status: 'idle', // idle | loading | ready | error
  error: null,
  wizardSchema: null,
  templateModel: null,
  currentStep: 0, // 0..slides.length (last is Preview)
  formData: {},
  refinedLaterNote: false,
};

function reducer(state, action) {
  switch (action.type) {
    case 'LOAD_START':
      return { ...state, status: 'loading', error: null };
    case 'LOAD_SUCCESS':
      return {
        ...state,
        status: 'ready',
        error: null,
        wizardSchema: action.payload.wizardSchema,
        templateModel: action.payload.templateModel,
        formData: action.payload.formData,
      };
    case 'LOAD_ERROR':
      return { ...state, status: 'error', error: action.payload };
    case 'SET_STEP':
      return { ...state, currentStep: action.payload };
    case 'SET_FIELD':
      return { ...state, formData: { ...state.formData, [action.payload.id]: action.payload.value } };
    case 'RESTORE_DEFAULTS':
      return { ...state, formData: action.payload.formData };
    case 'CLEAR_ALL':
      return { ...state, formData: action.payload.formData };
    case 'SET_REFINE_LATER':
      return { ...state, refinedLaterNote: true };
    default:
      return state;
  }
}

const WizardContext = createContext(null);

// PUBLIC_INTERFACE
export function WizardProvider({ children, loadSchemas }) {
  /** Provider for wizard state + persistence. Requires a loadSchemas() function returning {wizardSchema, templateModel}. */
  const [state, dispatch] = useReducer(reducer, initialState);

  // Load draft from localStorage and merge with defaults when schemas load
  useEffect(() => {
    let cancelled = false;

    async function run() {
      dispatch({ type: 'LOAD_START' });
      try {
        const { wizardSchema, templateModel } = await loadSchemas();
        const defaults = buildDefaultFormData(wizardSchema);

        const stored = safeJsonParse(localStorage.getItem(STORAGE_KEY) || '');
        const storedData = stored?.formData && typeof stored.formData === 'object' ? stored.formData : {};
        const storedStep = Number.isFinite(stored?.currentStep) ? stored.currentStep : 0;

        const merged = { ...defaults, ...storedData };

        if (!cancelled) {
          dispatch({
            type: 'LOAD_SUCCESS',
            payload: {
              wizardSchema,
              templateModel,
              formData: merged,
              currentStep: storedStep,
            },
          });
          if (Number.isFinite(storedStep)) dispatch({ type: 'SET_STEP', payload: storedStep });
        }
      } catch (e) {
        if (!cancelled) dispatch({ type: 'LOAD_ERROR', payload: String(e?.message || e) });
      }
    }

    run();
    return () => {
      cancelled = true;
    };
  }, [loadSchemas]);

  // Persist draft
  useEffect(() => {
    if (state.status !== 'ready') return;
    const payload = { formData: state.formData, currentStep: state.currentStep };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
  }, [state.status, state.formData, state.currentStep]);

  const api = useMemo(() => {
    return {
      state,
      dispatch,
      actions: {
        // PUBLIC_INTERFACE
        setStep(step) {
          /** Set current wizard step (0..N). */
          dispatch({ type: 'SET_STEP', payload: step });
        },
        // PUBLIC_INTERFACE
        next() {
          /** Go to next step if possible. */
          const max = (state.wizardSchema?.slides?.length || 0); // preview is max index
          dispatch({ type: 'SET_STEP', payload: Math.min(state.currentStep + 1, max) });
        },
        // PUBLIC_INTERFACE
        back() {
          /** Go to previous step if possible. */
          dispatch({ type: 'SET_STEP', payload: Math.max(state.currentStep - 1, 0) });
        },
        // PUBLIC_INTERFACE
        setField(id, value) {
          /** Update one field in the draft. */
          dispatch({ type: 'SET_FIELD', payload: { id, value } });
        },
        // PUBLIC_INTERFACE
        restoreDefaults() {
          /** Restore defaults for all fields (schema-defined defaultValue). */
          const defaults = buildDefaultFormData(state.wizardSchema);
          dispatch({ type: 'RESTORE_DEFAULTS', payload: { formData: defaults } });
        },
        // PUBLIC_INTERFACE
        clearAll() {
          /** Clear all user-entered values (sets strings to '', images to null). */
          const slides = Array.isArray(state.wizardSchema?.slides) ? state.wizardSchema.slides : [];
          const cleared = {};
          for (const s of slides) {
            for (const sec of s.sections || []) {
              for (const f of sec.fields || []) {
                cleared[f.id] = f.type === 'image' ? null : '';
              }
            }
          }
          if (cleared.title == null) cleared.title = '';
          dispatch({ type: 'CLEAR_ALL', payload: { formData: cleared } });
        },
        // PUBLIC_INTERFACE
        markRefineLater() {
          /** Mark that user wants to refine template later (UI-only placeholder for future workflow). */
          dispatch({ type: 'SET_REFINE_LATER' });
        },
      },
    };
  }, [state]);

  return <WizardContext.Provider value={api}>{children}</WizardContext.Provider>;
}

// PUBLIC_INTERFACE
export function useWizard() {
  /** Access wizard state/actions. */
  const ctx = useContext(WizardContext);
  if (!ctx) throw new Error('useWizard must be used inside WizardProvider');
  return ctx;
}
