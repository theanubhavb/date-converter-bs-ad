# BS/AD Date Converter for Excel

An Excel custom-functions add-in scaffolded in the same structure as the **Yeoman Office generator** `excel-functions` project. It exposes the following names under the `NPDATE` namespace:

| Excel formula | Result |
| --- | --- |
| `=NPDATE.ADTOBS(A1)` | BS date as `YYYY.MM.DD` |
| `=NPDATE.ADTOBSLONG(A1)` | BS date such as `01 Baisakh 2081` |
| `=NPDATE.BSTOAD("2081.01.01")` | AD date as `YYYY-MM-DD` |
| `=NPDATE.BSTOADLONG("2081.01.01")` | AD date such as `April 13, 2024` |
| `=NPDATE.BSDATEDIF("2081.01.02","2081.01.01")` | Signed whole-day difference |

The supported BS data range is 2000.01.01–2090.12.30. Invalid or out-of-range inputs return a readable Excel cell message.

## GitHub Pages deployment

Install Node.js 22 LTS, then run:

```sh
npm install
npm run build
```

The app is designed to run from GitHub Pages using the published manifest at `https://theanubhavb.github.io/date-converter-bs-ad/manifest.xml`.

## Organization deployment

1. Push the repository to GitHub and enable **Settings → Pages → GitHub Actions**.
2. The included workflow builds `dist/` and publishes it whenever `main` changes.
3. The workflow publishes the production manifest at `https://theanubhavb.github.io/date-converter-bs-ad/manifest.xml`.
4. Run `npm run validate:github-pages` before deployment.
5. Download the published `manifest.xml` and submit it in the Microsoft 365 admin center: **Settings → Integrated apps → Upload custom apps → Office Add-in**.

The manifest is intentionally separate from the hosted files: Microsoft 365 administrators upload the XML, while GitHub Pages hosts the HTML, JavaScript, and generated custom-function metadata.

## Verification

```sh
npm test
npm run lint
npm run validate
```

`npm run build` generates `dist/functions.json` from the JSDoc custom-function metadata and produces the static files GitHub Pages serves.
