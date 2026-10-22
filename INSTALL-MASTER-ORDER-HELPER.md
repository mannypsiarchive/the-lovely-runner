# Master Order Helper — install package

This is an additive package for the Lovely Runner static site. It adds a home-page card named **Master Order Helper**, which opens the **Master Order Delivery** workflow.

## Install

Copy these files into the root of the existing repository, beside `index.html` and `style.css`:

- `master-order-delivery.html`
- `master-order-delivery.js`

Replace the repository's `index.html` with the included updated `index.html` (or copy only the Master Order Helper `<a class="tool-card...">` block into the existing `.tool-grid`). No existing CSS or tool files need to be replaced.

## Google Sheets OAuth

The page reuses the Google OAuth web client ID already present in the supplied Lovely Runner source. The OAuth client must allow the published GitHub Pages origin in its authorized JavaScript origins. The page requests the Sheets read/write scope so it can create order tabs in a tracker the user chooses. No client secret is used in browser code.

## What the first version does

1. Connect to a tracker URL and choose its cost/legal tab and `VENDOR INFO AND RATES` tab.
2. Review guessed column mappings; select the source, archival flag/type, filename, still/footage, source and timeline timecode, description, and ARC columns.
3. Keep rows marked Archival or with an ARC-pattern filename; choose a vendor and accept or change the suggested order type. Suggestions read footage rate type from column O and still rate type from column Q.
4. Create a new Per Asset tab containing unique Original File Names for one vendor, or all per-asset vendors in alphabetical two-column blocks separated by a blank column.
5. For Per Second, choose the cut and a local folder of raw screeners. Files are matched to rows by ARC number. The page displays cut in/out and screener source in/out frames; OCR can suggest burnt-in timecodes, which remain editable. The reviewed values can be written to a new tracker tab.

## Notes / current limits

- This is a first installable iteration; the actual Cost Report/Legal Flag workbook was not included in the source ZIP, so column names are guessed from common headers and then confirmed by the user in the UI.
- Rate type is read from VENDOR INFO AND RATES columns O and Q. Vendor names are detected from a vendor/source header, or column A if no header is found.
- The per-second tab stores tracker rows, source/cut timecodes, and editable OCR suggestions. The four image frames are shown in the browser for review and are not embedded in Google Sheets.
- The browser must be able to decode the selected video formats. Some MXF codecs may not play in Chrome; proxy copies in a browser-supported format may be needed.
- OCR uses Tesseract.js from a CDN and reads the raw screener's sampled source-in/source-out frames. Results need human review. There is no automatic mismatch detection.
- If column labels are unfamiliar, map them manually before loading archival rows. The row filter uses Type containing “Archival” or an ARC filename pattern (`4–5` prefix characters + `ARC` + digits + underscore).
