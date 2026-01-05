import React, { createContext, useContext, useEffect, useMemo, useReducer } from 'react';
import { buildCanonicalOrderedSlidesFromTemplate, loadWizardFlowSchema, getFieldsForSlideType } from '../services/schemaLoader';

function pickFirstMatchingPlaceholderId(templateModel, candidates) {
  const layouts = Array.isArray(templateModel?.layouts) ? templateModel.layouts : [];
  for (const l of layouts) {
    const phs = Array.isArray(l?.placeholders) ? l.placeholders : [];
    for (const c of candidates) {
      const found = phs.find((p) => p?.id === c);
      if (found?.id) return found.id;
    }
  }
  return null;
}

function normalizeGlobalFirstFieldsFromTemplate(flowSchema, templateModel) {
  /**
   * Global First slide customization (per user requirements):
   * - Fixed background MUST remain identical to template.
   * - Expose exactly ONE input:
   *   - Date (date picker) -> GF_DATE
   * - Render the date value in "DD MMM YYYY" on-slide using the extracted template placeholder geometry + style.
   * - Do NOT expose or render any other editable fields (e.g., Name).
   *
   * Global Last slide customization:
   * - MUST be fully locked / non-editable.
   * - No inputs are exposed by the wizard.
   *
   * IMPORTANT:
   * - Do NOT alter layout/coordinates. Only map fields to already-existing placeholders.
   */
  if (!flowSchema || typeof flowSchema !== 'object') return flowSchema;

  const next = { ...flowSchema, slideTypes: { ...(flowSchema.slideTypes || {}) } };

  // -------------------------
  // Global First: enforce ONLY Date input
  // -------------------------
  const gf = next.slideTypes.global_first
    ? { ...next.slideTypes.global_first }
    : { label: 'Global First', layoutId: 'global_first', fields: [] };

  // Detect canonical placeholders by stable IDs (template-extracted bundle is authoritative)
  const gfDatePh = pickFirstMatchingPlaceholderId(templateModel, ['GF_DATE', 'DATE', 'DATE_PLACEHOLDER', 'SLIDEDATE']);

  const enforcedFields = [];
  if (gfDatePh) {
    enforcedFields.push({
      id: 'date',
      label: 'Date',
      type: 'date',
      defaultValue: '',
      validation: { required: true },
      mapping: { placeholderId: gfDatePh },
    });
  }

  gf.fields = enforcedFields;
  next.slideTypes.global_first = gf;

  // -------------------------
  // Global Last: enforce no inputs
  // -------------------------
  const gl = next.slideTypes.global_last
    ? { ...next.slideTypes.global_last }
    : { label: 'Global Last', layoutId: 'global_last', fields: [] };
  gl.fields = [];
  next.slideTypes.global_last = gl;

  // Keep fixed-text default machinery intact for backward compatibility (but Global First fixed text is already in template).
  next.__templateFixedTextDefaults = {
    globalFirst: {
      tagline: null,
      subtitle: null,
    },
    globalLast: {
      thankYou: { placeholderId: 'GL_THANK_YOU', text: 'THANK YOU' },
      brand: { placeholderId: 'GL_BRAND', text: 'TATA ELXSI' },
      cta: { placeholderId: 'GL_CTA', text: 'FIND OUT MORE' },
      website: { placeholderId: 'GL_WEBSITE', text: 'www.tataelxsi.com' },
    },
  };

  return next;
}

const STORAGE_KEY = 'ppt_wizard_draft_v2_grouped';

function safeJsonParse(str) {
  try {
    return JSON.parse(str);
  } catch {
    return null;
  }
}

function makeId() {
  return `sf_${Math.random().toString(16).slice(2)}_${Date.now().toString(16)}`;
}

function buildEmptySlideData(flowSchema, slideType) {
  const data = {};
  const fields = getFieldsForSlideType(flowSchema, slideType);
  for (const f of fields) {
    if (!f?.id) continue;
    data[f.id] = f.defaultValue ?? (f.type === 'image' ? null : '');
  }
  return data;
}

function buildDefaultWizardData(flowSchema) {
  return {
    globalFirst: buildEmptySlideData(flowSchema, 'global_first'),
    skillFactories: [
      {
        id: makeId(),
        slides: {
          sf1: buildEmptySlideData(flowSchema, 'sf1'),
          sf2: buildEmptySlideData(flowSchema, 'sf2'),
          sf3: buildEmptySlideData(flowSchema, 'sf3'),
          sf4: buildEmptySlideData(flowSchema, 'sf4'),
        },
      },
    ],
    globalLast: buildEmptySlideData(flowSchema, 'global_last'),
  };
}

function validateValue(field, value) {
  const v = field?.validation || {};
  const errors = [];

  if (v.required) {
    const empty = value == null || value === '' || (Array.isArray(value) && value.length === 0);
    if (empty) errors.push('This field is required.');
  }

  if (typeof value === 'string') {
    if (typeof v.minLength === 'number' && value.length < v.minLength) errors.push(`Must be at least ${v.minLength} characters.`);
    if (typeof v.maxLength === 'number' && value.length > v.maxLength) errors.push(`Must be at most ${v.maxLength} characters.`);
    if (typeof v.maxLines === 'number') {
      const lines = value.split('\n').length;
      if (lines > v.maxLines) errors.push(`Must be at most ${v.maxLines} lines.`);
    }
  }

  if (typeof value === 'number') {
    if (typeof v.min === 'number' && value < v.min) errors.push(`Must be ≥ ${v.min}.`);
    if (typeof v.max === 'number' && value > v.max) errors.push(`Must be ≤ ${v.max}.`);
  }

  return errors;
}

// PUBLIC_INTERFACE
export function buildOrderedSlides(flowSchema, wizardData) {
  /**
   * Build the ordered slide list required by the template-driven flow:
   * [global_first] + for each skillFactory [sf1..sf4] + [global_last].
   *
   * Returns an array of "step objects" used by WizardLayout + preview + generation.
   */
  const steps = [];

  steps.push({
    key: 'global_first',
    kind: 'global',
    slideType: 'global_first',
    title: flowSchema?.slideTypes?.global_first?.label || 'Global First',
    layoutId: flowSchema?.slideTypes?.global_first?.layoutId || 'global_first',
    fields: getFieldsForSlideType(flowSchema, 'global_first'),
    dataPath: { scope: 'globalFirst' },
  });

  const factories = Array.isArray(wizardData?.skillFactories) ? wizardData.skillFactories : [];
  factories.forEach((sf, idx) => {
    const groupLabel = `Skill Factory ${idx + 1}`;
    (['sf1', 'sf2', 'sf3', 'sf4'] || []).forEach((sft) => {
      steps.push({
        key: `${sf.id}:${sft}`,
        kind: 'skillFactory',
        groupId: sf.id,
        groupIndex: idx,
        groupLabel,
        slideType: sft,
        title: `${groupLabel} — ${flowSchema?.slideTypes?.[sft]?.label || sft.toUpperCase()}`,
        layoutId: flowSchema?.slideTypes?.[sft]?.layoutId || sft,
        fields: getFieldsForSlideType(flowSchema, sft),
        dataPath: { scope: 'skillFactories', groupId: sf.id, slideKey: sft },
      });
    });
  });

  steps.push({
    key: 'global_last',
    kind: 'global',
    slideType: 'global_last',
    title: flowSchema?.slideTypes?.global_last?.label || 'Global Last',
    layoutId: flowSchema?.slideTypes?.global_last?.layoutId || 'global_last',
    fields: getFieldsForSlideType(flowSchema, 'global_last'),
    dataPath: { scope: 'globalLast' },
  });

  return steps;
}

function getStepValue(wizardData, step, fieldId) {
  if (!step?.dataPath) return undefined;

  if (step.dataPath.scope === 'globalFirst') return wizardData?.globalFirst?.[fieldId];
  if (step.dataPath.scope === 'globalLast') return wizardData?.globalLast?.[fieldId];

  if (step.dataPath.scope === 'skillFactories') {
    const group = (wizardData?.skillFactories || []).find((g) => g.id === step.dataPath.groupId);
    return group?.slides?.[step.dataPath.slideKey]?.[fieldId];
  }

  return undefined;
}

function setStepValue(wizardData, step, fieldId, value) {
  if (step.dataPath.scope === 'globalFirst') {
    return { ...wizardData, globalFirst: { ...(wizardData.globalFirst || {}), [fieldId]: value } };
  }
  if (step.dataPath.scope === 'globalLast') {
    return { ...wizardData, globalLast: { ...(wizardData.globalLast || {}), [fieldId]: value } };
  }
  if (step.dataPath.scope === 'skillFactories') {
    const nextFactories = (wizardData.skillFactories || []).map((g) => {
      if (g.id !== step.dataPath.groupId) return g;
      const prevSlides = g.slides || {};
      const prevSlideData = prevSlides[step.dataPath.slideKey] || {};
      return {
        ...g,
        slides: { ...prevSlides, [step.dataPath.slideKey]: { ...prevSlideData, [fieldId]: value } },
      };
    });
    return { ...wizardData, skillFactories: nextFactories };
  }
  return wizardData;
}

function validateStep(flowSchema, wizardData, step) {
  const fields = Array.isArray(step?.fields) ? step.fields : [];
  const errorsByField = {};
  for (const f of fields) {
    const v = getStepValue(wizardData, step, f.id);
    const errs = validateValue(f, v);
    if (errs.length) errorsByField[f.id] = errs;
  }
  return errorsByField;
}

function isNonEmptyString(v) {
  return typeof v === 'string' && v.trim().length > 0;
}

function requiredFieldKeyFromStep(step, preferredKeys) {
  const fields = Array.isArray(step?.fields) ? step.fields : [];
  for (const k of preferredKeys) {
    if (fields.some((f) => f?.id === k)) return k;
  }
  // Fallback: first string-like field
  const first = fields.find((f) => f?.type === 'string' || f?.type === 'richText' || f?.type === 'textarea');
  return first?.id || null;
}

function minimalPreviewValidation(flowSchema, wizardData, orderedSlides) {
  /**
   * Requested gating:
   * - GlobalFirst title present
   * - Each SkillFactory has minimally required field stubs
   * - GlobalLast closing/title present
   *
   * We implement this conservatively:
   * - Use known common keys (title/closing/cta) when present; otherwise fall back to first text field.
   * - Also enforce each SF group has all 4 slides present in data structure.
   */
  const errors = [];

  const globalFirstStep = orderedSlides.find((s) => s.slideType === 'global_first');
  if (globalFirstStep) {
    // Per requirements: Global First must have exactly one required input: Date.
    if (!isNonEmptyString(wizardData?.globalFirst?.date)) errors.push('Global First: Date is required.');
  }

  const factories = Array.isArray(wizardData?.skillFactories) ? wizardData.skillFactories : [];
  if (!factories.length) errors.push('At least one Skill Factory group is required.');

  factories.forEach((sf, idx) => {
    const slides = sf?.slides || {};
    const hasAll = ['sf1', 'sf2', 'sf3', 'sf4'].every((k) => slides && Object.prototype.hasOwnProperty.call(slides, k));
    if (!hasAll) errors.push(`Skill Factory ${idx + 1}: must contain exactly 4 slides (SF-1..SF-4).`);

    // Minimal required content: at least one text field filled on SF-1 (or fallback).
    const sf1Step = orderedSlides.find((s) => s.groupId === sf.id && s.slideType === 'sf1');
    if (sf1Step) {
      const key = requiredFieldKeyFromStep(sf1Step, ['title', 'name', 'skillFactoryTitle']);
      if (key && !isNonEmptyString(slides?.sf1?.[key])) errors.push(`Skill Factory ${idx + 1} (SF-1): required field is missing.`);
    }
  });

  const globalLastStep = orderedSlides.find((s) => s.slideType === 'global_last');
  if (globalLastStep) {
    const key = requiredFieldKeyFromStep(globalLastStep, ['title', 'closing', 'cta']);
    if (key && !isNonEmptyString(wizardData?.globalLast?.[key])) errors.push('Global Last: closing/title is required.');
  }

  // Also run schema-required validations as a superset (if schema marks required fields).
  if (flowSchema) {
    for (const step of orderedSlides) {
      const errs = validateStep(flowSchema, wizardData, step);
      if (Object.keys(errs).length) {
        // Keep this generic: the UI will show inline errors; for gating, one error is enough.
        errors.push('Some required fields are missing.');
        break;
      }
    }
  }

  return errors;
}

const initialState = {
  status: 'idle', // idle | loading | ready | error
  error: null,

  // Legacy schemas retained for non-breaking change
  wizardSchema: null,
  templateModel: null,

  // Extracted artifacts bundle (masters/relationships/assetsManifest)
  extractedTemplate: null,

  // New flow schema
  flowSchema: null,

  // New grouped data model
  wizardData: null,

  currentStep: 0, // 0..orderedSlides.length (preview is last)
  refinedLaterNote: false,

  // Validation cache for inline display
  validation: {
    touched: {}, // stepKey -> true
    errors: {}, // stepKey -> {fieldId: [errs]}
  },

  // Preview gating errors (high-level)
  previewGateErrors: [],
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
        extractedTemplate: action.payload.extractedTemplate,
        flowSchema: action.payload.flowSchema,
        wizardData: action.payload.wizardData,
        currentStep: action.payload.currentStep,
      };
    case 'LOAD_ERROR':
      return { ...state, status: 'error', error: action.payload };

    case 'SET_STEP':
      return { ...state, currentStep: action.payload };

    case 'SET_WIZARD_DATA':
      return { ...state, wizardData: action.payload };

    case 'ADD_SKILL_FACTORY': {
      const sf = {
        id: makeId(),
        slides: {
          sf1: buildEmptySlideData(state.flowSchema, 'sf1'),
          sf2: buildEmptySlideData(state.flowSchema, 'sf2'),
          sf3: buildEmptySlideData(state.flowSchema, 'sf3'),
          sf4: buildEmptySlideData(state.flowSchema, 'sf4'),
        },
      };
      const next = { ...state.wizardData, skillFactories: [...(state.wizardData?.skillFactories || []), sf] };
      return { ...state, wizardData: next };
    }

    case 'REMOVE_SKILL_FACTORY': {
      const groupId = action.payload;
      const nextFactories = (state.wizardData?.skillFactories || []).filter((g) => g.id !== groupId);
      const next = { ...state.wizardData, skillFactories: nextFactories.length ? nextFactories : [] };
      return { ...state, wizardData: next };
    }

    case 'MOVE_SKILL_FACTORY': {
      const { groupId, direction } = action.payload; // -1 up, +1 down
      const arr = [...(state.wizardData?.skillFactories || [])];
      const idx = arr.findIndex((g) => g.id === groupId);
      if (idx < 0) return state;
      const nextIdx = idx + direction;
      if (nextIdx < 0 || nextIdx >= arr.length) return state;
      const tmp = arr[idx];
      arr[idx] = arr[nextIdx];
      arr[nextIdx] = tmp;
      return { ...state, wizardData: { ...state.wizardData, skillFactories: arr } };
    }

    case 'SET_FIELD_FOR_STEP': {
      const { step, fieldId, value } = action.payload;
      const nextData = setStepValue(state.wizardData, step, fieldId, value);
      return { ...state, wizardData: nextData };
    }

    case 'TOUCH_STEP': {
      const stepKey = action.payload;
      return {
        ...state,
        validation: {
          ...state.validation,
          touched: { ...state.validation.touched, [stepKey]: true },
        },
      };
    }

    case 'SET_STEP_ERRORS': {
      const { stepKey, errors } = action.payload;
      return {
        ...state,
        validation: {
          ...state.validation,
          errors: { ...state.validation.errors, [stepKey]: errors },
        },
      };
    }

    case 'SET_PREVIEW_GATE_ERRORS':
      return { ...state, previewGateErrors: action.payload || [] };

    case 'RESTORE_DEFAULTS':
      return { ...state, wizardData: action.payload.wizardData, validation: initialState.validation, previewGateErrors: [] };

    case 'CLEAR_ALL': {
      // Clear to empty values but keep structure.
      const cleared = buildDefaultWizardData(state.flowSchema);
      // Ensure title exists for filename, even if schema changes.
      if (cleared?.globalFirst?.title == null) cleared.globalFirst.title = '';
      return { ...state, wizardData: cleared, validation: initialState.validation, previewGateErrors: [] };
    }

    case 'SET_REFINE_LATER':
      return { ...state, refinedLaterNote: true };

    default:
      return state;
  }
}

const WizardContext = createContext(null);

// PUBLIC_INTERFACE
export function WizardProvider({ children, loadSchemas }) {
  /** Provider for wizard state + persistence. Requires a loadSchemas() function returning {wizardSchema, templateModel, extractedTemplate?}. */
  const [state, dispatch] = useReducer(reducer, initialState);

  // Load draft from localStorage and merge with defaults when schemas load
  useEffect(() => {
    let cancelled = false;

    async function run() {
      dispatch({ type: 'LOAD_START' });
      try {
        const [{ wizardSchema, templateModel, extractedTemplate }, flowSchemaRaw] = await Promise.all([loadSchemas(), loadWizardFlowSchema()]);

        // Enrich flow schema using extracted template placeholders.
        // Global First is enforced to exactly: Name (required) + Date (required).
        // Global Last is enforced to have NO inputs (fully locked).
        const flowSchema = normalizeGlobalFirstFieldsFromTemplate(flowSchemaRaw, templateModel);

        // Apply fixed-text defaults directly to templateModel so preview/PPT show them without inputs.
        const fixed = flowSchema?.__templateFixedTextDefaults || null;
        if (fixed && templateModel && Array.isArray(templateModel.layouts)) {
          const applyFixedText = (item) => {
            if (!item?.placeholderId) return;
            for (const l of templateModel.layouts) {
              const phs = Array.isArray(l?.placeholders) ? l.placeholders : [];
              const idx = phs.findIndex((p) => p?.id === item.placeholderId);
              if (idx >= 0) {
                phs[idx] = { ...phs[idx], text: item.text };
              }
            }
          };

          // Global First fixed texts
          applyFixedText(fixed?.globalFirst?.tagline);
          applyFixedText(fixed?.globalFirst?.subtitle);

          // Global Last fixed texts (locked slide)
          applyFixedText(fixed?.globalLast?.thankYou);
          applyFixedText(fixed?.globalLast?.brand);
          applyFixedText(fixed?.globalLast?.cta);
          applyFixedText(fixed?.globalLast?.website);
        }

        const defaults = buildDefaultWizardData(flowSchema);

        const stored = safeJsonParse(localStorage.getItem(STORAGE_KEY) || '');
        const storedData = stored?.wizardData && typeof stored.wizardData === 'object' ? stored.wizardData : null;
        const storedStep = Number.isFinite(stored?.currentStep) ? stored.currentStep : 0;

        const merged = storedData ? { ...defaults, ...storedData } : defaults;

        if (!cancelled) {
          dispatch({
            type: 'LOAD_SUCCESS',
            payload: {
              wizardSchema,
              templateModel,
              extractedTemplate,
              flowSchema,
              wizardData: merged,
              currentStep: storedStep,
            },
          });
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
    const payload = { wizardData: state.wizardData, currentStep: state.currentStep };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
  }, [state.status, state.wizardData, state.currentStep]);

  const api = useMemo(() => {
    const orderedSlides =
      state.flowSchema && state.wizardData
        ? buildCanonicalOrderedSlidesFromTemplate({
            flowSchema: state.flowSchema,
            wizardData: state.wizardData,
            templateModel: state.templateModel,
          })
        : [];
    const previewStepIndex = orderedSlides.length;

    function canGoToStep(targetStep) {
      // Allow moving among form steps freely; Preview is gated by requested "basic validation".
      if (targetStep < previewStepIndex) return true;
      if (targetStep !== previewStepIndex) return false;

      const gateErrors = minimalPreviewValidation(state.flowSchema, state.wizardData, orderedSlides);
      dispatch({ type: 'SET_PREVIEW_GATE_ERRORS', payload: gateErrors });
      return gateErrors.length === 0;
    }

    function markAndValidateStep(step) {
      dispatch({ type: 'TOUCH_STEP', payload: step.key });
      const errs = validateStep(state.flowSchema, state.wizardData, step);
      dispatch({ type: 'SET_STEP_ERRORS', payload: { stepKey: step.key, errors: errs } });
      return errs;
    }

    return {
      state: {
        ...state,
        orderedSlides,
        previewStepIndex,
      },
      dispatch,
      actions: {
        // PUBLIC_INTERFACE
        setStep(step) {
          /** Set current wizard step (0..N). Preview is gated by validation. */
          if (canGoToStep(step)) {
            dispatch({ type: 'SET_STEP', payload: step });
            return;
          }

          // If trying to go to preview but blocked, touch+validate all steps to show inline errors.
          for (const s of orderedSlides) markAndValidateStep(s);
        },

        // PUBLIC_INTERFACE
        next() {
          /** Go to next step if possible. Validates current step before advancing (and blocks preview). */
          const current = state.currentStep;
          const max = previewStepIndex; // preview is max index
          const currentIsPreview = current === previewStepIndex;

          if (currentIsPreview) return;

          const stepObj = orderedSlides[current];
          const errs = stepObj ? markAndValidateStep(stepObj) : {};
          if (Object.keys(errs).length) return;

          const nextStep = Math.min(current + 1, max);
          if (canGoToStep(nextStep)) dispatch({ type: 'SET_STEP', payload: nextStep });
        },

        // PUBLIC_INTERFACE
        back() {
          /** Go to previous step if possible. */
          dispatch({ type: 'SET_STEP', payload: Math.max(state.currentStep - 1, 0) });
        },

        // PUBLIC_INTERFACE
        setFieldForStep(step, fieldId, value) {
          /** Update one field in nested wizard data for a given ordered slide step. */
          dispatch({ type: 'SET_FIELD_FOR_STEP', payload: { step, fieldId, value } });

          // If the step was touched before, revalidate live for inline errors.
          if (state.validation.touched?.[step.key]) {
            const nextWizardData = setStepValue(state.wizardData, step, fieldId, value);
            const errs = validateStep(state.flowSchema, nextWizardData, step);
            dispatch({ type: 'SET_STEP_ERRORS', payload: { stepKey: step.key, errors: errs } });
          }
        },

        // PUBLIC_INTERFACE
        addSkillFactory() {
          /** Append a new Skill Factory group (4 slides). */
          dispatch({ type: 'ADD_SKILL_FACTORY' });
        },

        // PUBLIC_INTERFACE
        removeSkillFactory(groupId) {
          /** Remove an existing Skill Factory group by id. */
          dispatch({ type: 'REMOVE_SKILL_FACTORY', payload: groupId });
        },

        // PUBLIC_INTERFACE
        moveSkillFactory(groupId, direction) {
          /** Reorder Skill Factory groups. direction: -1 up, +1 down. */
          dispatch({ type: 'MOVE_SKILL_FACTORY', payload: { groupId, direction } });
        },

        // PUBLIC_INTERFACE
        restoreDefaults() {
          /** Restore defaults for all slide fields from flow schema. */
          const defaults = buildDefaultWizardData(state.flowSchema);
          dispatch({ type: 'RESTORE_DEFAULTS', payload: { wizardData: defaults } });
        },

        // PUBLIC_INTERFACE
        clearAll() {
          /** Clear all user-entered values (sets strings to '', images to null) but preserves structure. */
          dispatch({ type: 'CLEAR_ALL' });
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
