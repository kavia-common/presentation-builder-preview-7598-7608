# Extracted Template Artifacts

This folder contains **machine-extracted PPTX template artifacts** used by the preview/generator for pixel-perfect matching.

Expected files:
- `pptx_template.normalized.json` — normalized template (layouts/slides/theme) following `public/assets/pptx_template_schema.json`
- `pptx_template.masters.json` — master-specific detail (more verbose than the normalized schema)
- `pptx_template.relationships.json` — relationship maps (rId -> target) for debugging
- `assets_manifest.json` — media assets inventory (hashes, original paths)

If these files are missing, run the template extraction step (backend/utility) to regenerate them from the source PPTX.

