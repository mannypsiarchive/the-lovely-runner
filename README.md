# Master Order Helper 1.13 test

This build fixes slow/black rows during long shot reviews by attaching the shot list first and capturing images serially as each card nears the viewport. Cards show an explicit loading state. Preview canvases are capped at 960 pixels wide; full-resolution source frames are kept only temporarily while OCR runs. Previews far outside the viewport are released and reloaded when you scroll back. Cut screenshots remain three frames inside the listed boundaries, and all shots remain sorted by show Cut In / Cut Out. Timecode column correction and first-frame BITC fallback from 1.12 are retained.

## Install for a test

Copy `master-order-delivery.html` and `master-order-delivery.js` into the Lovely Runner site root. The page visibly identifies itself as **1.13 Test**. The reference cut and screeners are read locally by the browser.

## Roll back

Copy the two files from `rollback-1.9/` back into the site root to restore 1.9.

This package has not been pushed to GitHub or deployed.
