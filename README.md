# Timeline Widget (Iteration 1)

This project provides a self-contained timeline widget in one embeddable file: `timeline-widget.html`.

## What it does

- Fetches CSV data from a public Google Sheet.
- Parses `Timeline` first, derives Trend sheet names from `Type = Trend` rows.
- Fetches Trend sheets in parallel using `Promise.all`.
- Builds vis-timeline groups/items with required stacking behavior.
- Applies date parsing defaults exactly as specified.
- Shows loading, ready, and error states.
- Supports optional local CORS proxy usage for browser testing.

## Files

- `timeline-widget.html` - main embeddable widget (single-file deliverable).
- `smoke-test.mjs` - tiny local static validation script.
- `proxy-server.mjs` - local CORS proxy for Google Sheets CSV requests.

## Quick local smoke test

```powershell
node .\smoke-test.mjs
```

## Local preview (with CORS proxy)

Google Sheets CSV endpoints commonly block direct browser CORS requests. For local testing, run both servers:

 ```powershell
node .\proxy-server.mjs
```

In another terminal:

```powershell
python -m http.server 8080
```

Then open `http://localhost:8080/timeline-widget.html` in your browser.

Set this constant in `timeline-widget.html` before loading the page:

```js
const CORS_PROXY = "http://127.0.0.1:8787";
```

## Squarespace usage

1. Open `timeline-widget.html`.
2. Copy all file contents.
3. Paste into a Squarespace Code Block.
4. If needed, adjust `SHEET_ID` and `SHEET_NAMES` constants near the top of the script.
5. Leave `CORS_PROXY` empty in production unless you intentionally run your own proxy.
