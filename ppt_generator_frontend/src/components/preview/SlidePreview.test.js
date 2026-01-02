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
