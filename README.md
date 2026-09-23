# CSL.json → Word Sources.xml

A static, browser-based source-library manager and converter for turning structured **CSL-JSON** references into Microsoft Word bibliography **Sources.xml**.

## Stack

- React + TypeScript + Vite
- Tailwind CSS v4
- shadcn/ui-style components
- Lucide icons
- GitHub Pages deployment via GitHub Actions

## Source-library workflow

The site now acts as a small portable bibliography database without a backend:

- Keep a searchable library of sources in the current workspace
- Select only the sources you want in a Word export
- Edit workspace-only metadata such as category, purpose, and tags
- Optionally timestamp newly added sources with an `addedAt` value
- Export the complete application state as a `.cslword-workspace.json` file
- Re-open that workspace file later to restore the library and conversion settings
- Import raw CSL-JSON into the current library
- Export raw CSL-JSON as well as Word `Sources.xml`

No server-side database is required; the workspace file is the database.

## Included FDM project preset

The default workspace contains the nine JIS standards selected for the FDM strength project:

- JIS K 7139:2026 — coupon/test-specimen geometry
- JIS K 7161-1:2024 — tensile general principles
- JIS K 7161-2:2014 — tensile test conditions
- JIS K 7171:2022 — three-point bending
- JIS K 7111-1:2012 — Charpy impact
- JIS K 7100:1999 — conditioning/test atmosphere
- JIS B 7721:2018 — force calibration
- JIS B 7741:2019 — extensometer calibration
- JIS B 7739:2020 — impact-machine verification

Each preset entry also includes project metadata such as purpose, category, and tags.

## CSL → Word conversion

The converter:

- Accepts a single CSL item, an array, or `{ "items": [...] }`
- Preserves structured personal and corporate authors
- Maps common CSL types to Word bibliography source types
- Handles issued/accessed dates, DOI, URL, ISBN/ISSN, publisher, volume, issue, pages, standards, patents, theses, and more
- Supports conservative duplicate removal
- Supports the Office 2006+ and legacy Word 2004 bibliography namespaces

## Local development

```bash
npm install
npm test
npm run build
npm run dev
```

## Word import

1. Select the sources to export.
2. Download `Sources.xml`.
3. In Word, open **References → Manage Sources**.
4. Choose **Browse…** and select the generated XML file.
5. Copy the sources you want into Word's current/master source list.

## GitHub Pages

Every push to `main` runs tests, builds the Vite app, and deploys it with GitHub Actions.

https://marcusfunt.github.io/CSL.json-to-WordSource/
