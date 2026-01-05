/**
 * Canonical template style utilities used by BOTH preview and PPT generation.
 *
 * IMPORTANT (per requirements):
 * - Do NOT introduce non-template fallbacks for fonts/colors/sizes/line-heights/alignments.
 * - If a style property is missing from the extracted template, we return undefined for that property.
 *   (Callers may decide to omit drawing rather than inventing a fallback.)
 */

// PUBLIC_INTERFACE
export function normalizeHexColor(templateColor) {
  /** Normalize a template color into a CSS hex string (e.g. "#RRGGBB") or null/undefined. */
  if (templateColor == null) return undefined;
  if (typeof templateColor !== 'string') return undefined;
  const c = templateColor.trim();
  if (!c || c === 'none') return undefined;
  if (c.startsWith('#')) return c;
  if (/^[0-9a-fA-F]{6}$/.test(c)) return `#${c}`;
  return c;
}

// PUBLIC_INTERFACE
export function cssFontFamilyFromTemplate(fontFamily) {
  /** Convert a template font family name to a CSS font-family value, verbatim. */
  if (fontFamily == null) return undefined;
  if (typeof fontFamily !== 'string') return undefined;
  const f = fontFamily.trim();
  if (!f) return undefined;
  // Use the template-provided font as-is (quoted). Do not add fallback stacks.
  return `"${f}"`;
}

// PUBLIC_INTERFACE
export function cssFontSizeFromPt(fontSizePt) {
  /** Convert template point size to CSS px for preview: 1pt ~= 1px for on-screen fidelity. */
  if (typeof fontSizePt !== 'number' || Number.isNaN(fontSizePt)) return undefined;
  return `${Math.max(1, fontSizePt)}px`;
}

// PUBLIC_INTERFACE
export function cssLineHeightFromTemplate(lineHeight) {
  /**
   * Normalize lineHeight for CSS.
   * Extractor may emit:
   * - multiplier (e.g. 1.1)
   * - string "1.1"
   * - pt-like numeric values (rare), treated as px-equivalent for preview
   *
   * IMPORTANT: no fallbacks.
   */
  if (lineHeight == null) return undefined;

  if (typeof lineHeight === 'number') {
    if (lineHeight > 0 && lineHeight <= 3) return lineHeight;
    return `${Math.max(1, lineHeight)}px`;
  }

  if (typeof lineHeight === 'string') {
    const s = lineHeight.trim();
    if (!s) return undefined;
    if (/^[0-9]*\.?[0-9]+$/.test(s)) {
      const n = Number(s);
      if (!Number.isNaN(n)) {
        if (n > 0 && n <= 3) return n;
        return `${Math.max(1, n)}px`;
      }
    }
    return s;
  }

  return undefined;
}

// PUBLIC_INTERFACE
export function cssTextAlignFromTemplate(align) {
  /** Normalize template align into CSS textAlign, verbatim where possible. */
  if (align == null) return undefined;
  if (typeof align !== 'string') return undefined;
  const a = align.trim().toLowerCase();
  if (!a) return undefined;
  if (a === 'left' || a === 'center' || a === 'right' || a === 'justify') return a;
  return a;
}

// PUBLIC_INTERFACE
export function cssFontWeightFromTemplate(fontWeight) {
  /** Return template-provided fontWeight verbatim (number or string), without fallbacks. */
  if (fontWeight == null) return undefined;
  if (typeof fontWeight === 'number') return fontWeight;
  if (typeof fontWeight === 'string') return fontWeight;
  return undefined;
}

// PUBLIC_INTERFACE
export function toPptColorRgb(templateColor) {
  /** Convert template color string into PptxGenJS RGB value without '#', or undefined. */
  const c = normalizeHexColor(templateColor);
  if (!c || typeof c !== 'string') return undefined;
  if (c.startsWith('#')) return c.slice(1);
  // If extractor emitted a non-hex token, pass through only if it looks like RGB hex.
  if (/^[0-9a-fA-F]{6}$/.test(c)) return c;
  return undefined;
}

// PUBLIC_INTERFACE
export function toPptBoldFromFontWeight(fontWeight) {
  /** Convert template fontWeight into a boolean bold flag for PptxGenJS. */
  if (fontWeight == null) return undefined;
  if (typeof fontWeight === 'number') return fontWeight >= 700;
  if (typeof fontWeight === 'string') {
    const s = fontWeight.trim().toLowerCase();
    if (!s) return undefined;
    if (s === 'bold' || s === 'bolder') return true;
    if (s === 'normal' || s === 'regular' || s === 'lighter') return false;
    const n = Number(s);
    if (!Number.isNaN(n)) return n >= 700;
  }
  return undefined;
}

// PUBLIC_INTERFACE
export function toPptAlignFromTemplate(align) {
  /**
   * PptxGenJS uses: 'left' | 'center' | 'right' | 'justify'
   * Return undefined if not provided by the template.
   */
  if (align == null) return undefined;
  if (typeof align !== 'string') return undefined;
  const a = align.trim().toLowerCase();
  if (!a) return undefined;
  if (a === 'left' || a === 'center' || a === 'right' || a === 'justify') return a;
  return undefined;
}

/**
 * Note on paragraph spacing:
 * - The current normalized extractor JSON in this repo does not expose paragraph spacing fields.
 * - This module intentionally does not fabricate such values.
 * - If/when the extractor provides e.g. { spaceBeforePt, spaceAfterPt }, callers can pass them through.
 */

// PUBLIC_INTERFACE
export function buildCssTextStyleFromTemplateStyle(style) {
  /**
   * Convert a template placeholder .style into CSS for preview.
   * Returns an object with only template-derived properties (no defaults).
   */
  if (!style || typeof style !== 'object') return {};

  return {
    fontFamily: cssFontFamilyFromTemplate(style.fontFamily),
    fontSize: cssFontSizeFromPt(style.fontSizePt),
    fontWeight: cssFontWeightFromTemplate(style.fontWeight),
    color: normalizeHexColor(style.color),
    textAlign: cssTextAlignFromTemplate(style.align),
    lineHeight: cssLineHeightFromTemplate(style.lineHeight),
  };
}

// PUBLIC_INTERFACE
export function buildPptTextOptionsFromTemplateStyle(style) {
  /**
   * Convert a template placeholder .style into PptxGenJS addText options.
   * Returns an object with only template-derived properties (no defaults).
   */
  if (!style || typeof style !== 'object') return {};

  const opts = {};

  if (typeof style.fontSizePt === 'number' && !Number.isNaN(style.fontSizePt)) {
    opts.fontSize = Math.max(1, style.fontSizePt);
  }

  const bold = toPptBoldFromFontWeight(style.fontWeight);
  if (typeof bold === 'boolean') opts.bold = bold;

  const color = toPptColorRgb(style.color);
  if (color) opts.color = color;

  const align = toPptAlignFromTemplate(style.align);
  if (align) opts.align = align;

  if (typeof style.fontFamily === 'string' && style.fontFamily.trim()) {
    opts.fontFace = style.fontFamily.trim();
  }

  // Line height and paragraph spacing are intentionally not fabricated.
  // If the extractor later adds them, add pass-through fields here.

  return opts;
}
