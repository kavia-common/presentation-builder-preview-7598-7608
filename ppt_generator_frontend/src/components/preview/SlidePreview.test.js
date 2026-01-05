import React from 'react';
import { render, screen } from '@testing-library/react';
import SlidePreview from './SlidePreview';

function minimalTemplateWithImagePlaceholder() {
  return {
    meta: { pageSize: { widthPt: 960, heightPt: 540 } },
    theme: { colors: { dk1: '#111827' }, fonts: { major: 'Arial', minor: 'Arial', fallbackStack: 'Arial, sans-serif' } },
    layouts: [
      {
        id: 'global_first',
        name: 'Global First',
        placeholders: [
          {
            id: 'GF_LOGO',
            kind: 'image',
            box: { xPt: 10, yPt: 10, wPt: 100, hPt: 50 },
            assetId: 'logo1',
          },
        ],
      },
    ],
    slides: [],
    assets: {
      images: [{ id: 'logo1', type: 'image', suggestedName: 'logo.png' }],
      icons: [],
      fonts: [],
    },
  };
}

function minimalTemplateWithSf1DateRange() {
  return {
    meta: { pageSize: { widthPt: 960, heightPt: 540 } },
    theme: { colors: { dk1: '#111827' }, fonts: { major: 'Arial', minor: 'Arial', fallbackStack: 'Arial, sans-serif' } },
    layouts: [
      {
        id: 'sf1',
        name: 'Skill Factory Slide 1',
        placeholders: [
          {
            id: 'SF1_DATE_RANGE',
            kind: 'body',
            box: { xPt: 680, yPt: 46, wPt: 250, hPt: 16 },
            zIndex: 12,
            style: { fontFamily: 'Arial', fontSizePt: 10, fontWeight: 600, color: '#111827', align: 'right', lineHeight: 1 },
          },
        ],
      },
    ],
    slides: [],
    assets: { images: [], icons: [], fonts: [] },
  };
}

test('uses template default image asset when user image is absent', () => {
  render(
    <SlidePreview
      slideStep={{
        key: 'global_first',
        title: 'Global First',
        layoutId: 'global_first',
        fields: [],
        dataPath: { scope: 'globalFirst' },
      }}
      templateModel={minimalTemplateWithImagePlaceholder()}
      extractedTemplate={{ assetsManifest: { assets: [{ id: 'logo1', type: 'image', suggestedName: 'logo.png' }] } }}
      wizardData={{ globalFirst: {}, skillFactories: [], globalLast: {} }}
    />
  );

  // Should render an <img> for the placeholder with template asset URL.
  const img = screen.getByAltText('GF_LOGO');
  expect(img).toBeInTheDocument();
  expect(img.getAttribute('src')).toBe('/assets/template_extracted/logo.png');
});

test('SF1_DATE_RANGE is rendered once using template typography (deduped like PPT export)', () => {
  render(
    <SlidePreview
      slideStep={{
        key: 'sf_1:sf1',
        title: 'Skill Factory 1 — SF-1',
        slideType: 'sf1',
        layoutId: 'sf1',
        fields: [
          { id: 'dateRangeStart', type: 'date', mapping: { placeholderId: 'SF1_DATE_RANGE' } },
          { id: 'dateRangeEnd', type: 'date', mapping: { placeholderId: 'SF1_DATE_RANGE' } },
        ],
        dataPath: { scope: 'skillFactories', groupId: 'sf_1', slideKey: 'sf1' },
      }}
      templateModel={minimalTemplateWithSf1DateRange()}
      extractedTemplate={{}}
      wizardData={{
        globalFirst: {},
        skillFactories: [
          {
            id: 'sf_1',
            slides: {
              sf1: { dateRangeStart: '2026-01-01', dateRangeEnd: '2026-01-07' },
            },
          },
        ],
        globalLast: {},
      }}
    />
  );

  // The placeholder container uses title={ph.id}.
  const placeholders = screen.getAllByTitle('SF1_DATE_RANGE');
  expect(placeholders).toHaveLength(1);

  // Typography is applied to inner .shape-label. Ensure fontSize and right alignment are present.
  const shapeLabel = placeholders[0].querySelector('.shape-label');
  expect(shapeLabel).toBeTruthy();
  expect(shapeLabel.style.fontSize).toBe('10px');
  expect(shapeLabel.style.textAlign).toBe('right');

  // Content should be some non-empty combined range (exact formatting is handled by dateFormat.js).
  expect(shapeLabel.textContent.trim().length).toBeGreaterThan(0);
});
