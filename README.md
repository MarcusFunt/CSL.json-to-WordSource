# CSL.json → Word Sources.xml

A browser-only converter for turning structured **CSL-JSON** references into Microsoft Word bibliography **Sources.xml**.

## Stack

- React + TypeScript + Vite
- Tailwind CSS v4
- shadcn/ui-style components
- Lucide icons
- GitHub Pages deployment via GitHub Actions

## Features

- Paste CSL-JSON or upload a `.json` file
- Accepts a single CSL item, an array of items, or `{ "items": [...] }`
- Converts structured personal and corporate authors
- Maps common CSL types to Word bibliography source types
- Handles issued/accessed dates, DOI, URL, ISBN/ISSN, publisher, volume, issue, pages, standards, patents, theses, and more
- Optional conservative duplicate removal
- Choose modern Office 2006 or legacy Word 2004 bibliography namespace
- Download the result as `Sources.xml`
- Entirely client-side: references never leave the browser

## Local development

```bash
npm install
npm run dev
```

Build and test:

```bash
npm test
npm run build
```

## Word import

1. Convert your CSL-JSON and download `Sources.xml`.
2. In Word, open **References → Manage Sources**.
3. Choose **Browse…** and select the generated XML file.
4. Copy the sources you want into Word's current/master source list.
5. Insert citations normally through **References → Insert Citation**.

## GitHub Pages

The workflow in `.github/workflows/deploy-pages.yml` builds and deploys the site on every push to `main`.

Expected URL:

https://marcusfunt.github.io/CSL.json-to-WordSource/

## Notes

CSL and Word do not have identical source-type models. The converter uses deterministic mappings and reports when an unknown CSL type falls back to Word's `Misc` type.
