# Master Order Helper — install package

This is an additive package for the Lovely Runner static site. It adds a home-page card named **Master Order Helper**, which opens the **Master Order Helper** workflow.

## Install

Copy these files into the root of the existing repository, beside `index.html` and `style.css`:

- `master-order-delivery.html`
- `master-order-delivery.js`

Replace the repository's `index.html` with the included updated `index.html` (or copy only the Master Order Helper `<a class="tool-card...">` block into the existing `.tool-grid`). No existing CSS or tool files need to be replaced.

## Google Sheets OAuth

The page reuses the Google OAuth web client ID already present in the supplied Lovely Runner source. The OAuth client must allow the published GitHub Pages origin in its authorized JavaScript origins. The page requests the Sheets read/write scope so it can create order tabs in a tracker the user chooses. No client secret is used in browser code.

## What this version does

1. Connect to a tracker URL and choose its cost/legal tab and `VENDOR INFO AND RATES` tab.
2. Review guessed column mappings; select the source, type, filename, still/footage, source and timeline timecode, description, and ARC columns.
3. Load the show rows, then assign each source Per Second, Per Asset, Fair Use, Public Domain, or N/A. Per Asset and Per Second suggestions read rate types from the vendor rate tab.
4. Create one combined Per Asset tab with one column per source group in A, C, E, and subsequent odd columns, with a bold light-blue header row. Fair Use gets a tab with element/file, source, show in/out, description, and tracker row. Public Domain gets the matching EDL rows.
5. Choose a reference cut and raw screener folder once. The helper reads each screener’s embedded start timecode and frame rate with MediaInfo, with FFmpeg as a fallback, and shows those values for review. Build a separate Per Second shot review for each source; files match to rows by ARC number. Cut and screener offsets use their own detected metadata. PaddleOCR checks enlarged, unthresholded top-center and bottom-third frame regions, with Tesseract as a fallback; suggestions remain editable and can be written to a separate tracker tab for each source.

## Notes / current limits

- This is a first installable iteration; the actual Cost Report/Legal Flag workbook was not included in the source ZIP, so column names are guessed from common headers and then confirmed by the user in the UI.
- Rate type is read from VENDOR INFO AND RATES columns O and Q. Vendor names are detected from a vendor/source header, or column A if no header is found.
- N/A sources are intentionally skipped and do not appear on generated tabs. Sources left unassigned are also skipped.
- The per-second review tab stores tracker rows, source/cut timecodes, and editable OCR suggestions. OCR checks enlarged crops from the top-center and bottom third of each screener frame, where burnt-in timecodes commonly appear. N/A sources are excluded from all order tabs. Per Asset tabs use a bold, light-blue header row. The four image frames are shown in the browser for review and are not embedded in Google Sheets. Frames are captured with a reused cut player and a reused screener player, then the full shot list appears together; each frame is captured independently so one failed seek does not block the remaining frames. Each frame capture uses the browser video player and remains independent so a problem on one frame does not block the remaining frames. The player waits for a newly presented frame after each seek and cancels expired frame callbacks to prevent a Cut In image from being reused as Cut Out. BITC OCR runs for every captured source frame. Click a frame to open it full-size. The Google Sheet URL field is wide and white for readability.
- The page follows Fair Use Log metadata detection: FFmpeg reads the embedded file start timecode and frame rate, with MediaInfo as a fallback. EDL timecodes are converted to frame counts and the detected file start is subtracted before seeking. For example, a cut that starts at `00:59:50:00` maps EDL time `01:00:06:20` to `00:00:16:20` at 24 fps. If metadata cannot provide both values, frame capture stops instead of guessing. Metadata probing runs in the browser; the selected media is not uploaded. The browser must be able to decode the selected video formats. Some MXF codecs may not play in Chrome; proxy copies in a browser-supported format may be needed.
- OCR runs PaddleOCR.js in the browser on enlarged crops of the raw screener’s sampled source-in/source-out frames; Tesseract.js tries multiple crop and contrast variants if PaddleOCR cannot find a timecode. The OCR model is downloaded by the browser on first use; video frames are not uploaded. Results need human review, and there is no automatic mismatch detection.
- If column labels are unfamiliar, map them manually before loading archival rows. Rows are included when a source and at least one useful file, description, or timecode field are present; sources marked N/A or left unassigned are skipped when output tabs are created.
