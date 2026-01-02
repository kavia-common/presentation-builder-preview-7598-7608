import React from 'react';

// PUBLIC_INTERFACE
export default function Stepper({ steps, current, onSelect }) {
  /** Horizontal stepper with accessible buttons for each step. */
  return (
    <div className="ocean-card" style={{ padding: 12 }}>
      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
        {steps.map((s, idx) => {
          const active = idx === current;
          const done = idx < current;
          return (
            <button
              key={s.key}
              type="button"
              className={`ocean-btn ${active ? 'ocean-btn-primary' : 'ocean-btn-ghost'}`}
              onClick={() => onSelect(idx)}
              aria-current={active ? 'step' : undefined}
              aria-label={`Go to step ${idx + 1}: ${s.label}`}
              style={{
                borderColor: done ? 'rgba(37, 99, 235, 0.25)' : undefined,
                background: active ? undefined : done ? 'rgba(37, 99, 235, 0.08)' : undefined,
              }}
            >
              <span
                style={{
                  width: 22,
                  height: 22,
                  borderRadius: 999,
                  display: 'inline-flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  fontSize: 12,
                  fontWeight: 800,
                  background: active ? 'rgba(255,255,255,0.18)' : 'rgba(17,24,39,0.04)',
                }}
                aria-hidden="true"
              >
                {idx + 1}
              </span>
              <span style={{ fontSize: 13, whiteSpace: 'nowrap' }}>{s.label}</span>
            </button>
          );
        })}
      </div>
    </div>
  );
}
