# Screens Down, Family Up — Ebook (v3, print edition)

Standalone, print-ready implementation of the **"Screens Down Family Up Ebook v3-print"** design
(Claude Design project *Ebook upsell printables*).

## Files

| File | Purpose |
|------|---------|
| `screens-down-family-up-ebook-v3.html` | The ebook — open in any browser. Self-contained (no CDN dependencies). |
| `doc-page.js` | `<doc-page>` web component: renders the on-screen sheet and owns the print geometry (`@page` letter, 0.75in margins). |
| `fonts/` | Locally hosted Oswald, Lora, and Source Sans 3 (woff2, SIL Open Font License), so the book renders identically offline. |
| `screens-down-family-up-ebook-v3.pdf` | Pre-exported letter-size PDF (25 pages). |

## Exporting the PDF

Open the HTML file in Chrome and print (**⌘/Ctrl+P**) with:

- Destination: *Save as PDF*
- Paper size: *Letter* — margins are handled by the document, leave them at default
- **Background graphics: on**

Or headlessly:

```sh
chromium --headless --no-sandbox --print-to-pdf=ebook.pdf --no-pdf-header-footer \
  screens-down-family-up-ebook-v3.html
```

## Structure

The whole book is normal HTML flow inside `<doc-page size="letter" margin="0.75in">`;
the browser's print engine paginates it. Chapter heroes (`.ch-hero`) and the cover use
negative margins against the sheet padding for full-bleed color bands, and
`break-before/after: page` marks chapter boundaries.
