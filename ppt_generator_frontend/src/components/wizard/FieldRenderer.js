import React, { useMemo } from 'react';

function validateField(field, value) {
  const v = field?.validation || {};
  const errors = [];

  if (v.required) {
    const empty =
      value == null ||
      value === '' ||
      (Array.isArray(value) && value.length === 0);

    if (empty) errors.push('This field is required.');
  }

  if (typeof value === 'string') {
    if (typeof v.minLength === 'number' && value.length < v.minLength) {
      errors.push(`Must be at least ${v.minLength} characters.`);
    }
    if (typeof v.maxLength === 'number' && value.length > v.maxLength) {
      errors.push(`Must be at most ${v.maxLength} characters.`);
    }
    if (v.pattern) {
      try {
        const re = new RegExp(v.pattern);
        if (!re.test(value)) errors.push('Invalid format.');
      } catch {
        // ignore invalid regex patterns in schema
      }
    }
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

function humanizeType(type) {
  switch (type) {
    case 'string': return 'Text';
    case 'richText': return 'Rich text';
    case 'number': return 'Number';
    case 'date': return 'Date';
    case 'select': return 'Select';
    case 'image': return 'Image';
    default: return type || 'Field';
  }
}

// PUBLIC_INTERFACE
export default function FieldRenderer({ field, value, onChange }) {
  /** Render one field driven by wizard schema; includes basic schema validation. */
  const id = field.id;
  const label = field.label || id;
  const helpText = field.helpText;

  const errors = useMemo(() => validateField(field, value), [field, value]);

  const common = {
    id,
    name: id,
    'aria-describedby': helpText ? `${id}-help` : undefined,
    'aria-invalid': errors.length ? 'true' : 'false',
  };

  return (
    <div style={{ marginBottom: 14 }}>
      <label className="ocean-label" htmlFor={id}>
        {label}
        <span style={{ marginLeft: 8, color: 'var(--ocean-muted)', fontWeight: 600, fontSize: 12 }}>
          ({humanizeType(field.type)})
        </span>
      </label>

      {field.type === 'string' && (
        <input
          {...common}
          className="ocean-input"
          type="text"
          value={value ?? ''}
          onChange={(e) => onChange(e.target.value)}
          placeholder={field?.validation?.required ? 'Required' : ''}
        />
      )}

      {field.type === 'number' && (
        <input
          {...common}
          className="ocean-input"
          type="number"
          value={value ?? ''}
          onChange={(e) => {
            const raw = e.target.value;
            onChange(raw === '' ? '' : Number(raw));
          }}
        />
      )}

      {(field.type === 'richText' || field.type === 'textarea') && (
        <textarea
          {...common}
          className="ocean-textarea"
          value={value ?? ''}
          onChange={(e) => onChange(e.target.value)}
          placeholder={field?.validation?.maxLines ? `Max ${field.validation.maxLines} lines` : ''}
        />
      )}

      {field.type === 'date' && (
        <input
          {...common}
          className="ocean-input"
          type="date"
          value={value ?? ''}
          onChange={(e) => onChange(e.target.value)}
        />
      )}

      {field.type === 'select' && (
        <select
          {...common}
          className="ocean-select"
          value={value ?? ''}
          onChange={(e) => onChange(e.target.value)}
        >
          <option value="" disabled>
            Select…
          </option>
          {(field.options || []).map((opt) => (
            <option key={String(opt.value)} value={opt.value}>
              {opt.label}
            </option>
          ))}
        </select>
      )}

      {field.type === 'image' && (
        <div>
          <input
            {...common}
            className="ocean-input"
            type="file"
            accept="image/*"
            onChange={(e) => {
              const file = e.target.files && e.target.files[0];
              onChange(file || null);
            }}
          />
          {value && (
            <div className="ocean-help">
              Selected: <span className="ocean-kbd">{value?.name || 'image'}</span>
            </div>
          )}
        </div>
      )}

      {helpText && (
        <div id={`${id}-help`} className="ocean-help">
          {helpText}
        </div>
      )}

      {field?.mapping?.placeholderId && (
        <div className="ocean-help">
          Maps to placeholder: <span className="ocean-kbd">{field.mapping.placeholderId}</span>
        </div>
      )}

      {errors.length > 0 && (
        <div className="ocean-error" role="alert">
          {errors[0]}
        </div>
      )}
    </div>
  );
}
