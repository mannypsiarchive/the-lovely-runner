# Master Order Helper 1.11 test

This test restores Cut In and Cut Out reference screenshots. Each cut screenshot is captured three frames inside its listed show timecode: Cut In at +3 frames and Cut Out at -3 frames, to move past transition frames. Per Second visual review is ordered by show Cut In timecode, then Cut Out timecode, regardless of source file or source timecode order. Source In and Source Out images and their editable OCR suggestions remain available.

## Install for a test

Copy `master-order-delivery.html` and `master-order-delivery.js` into the Lovely Runner site root. The page visibly identifies itself as **1.11 Test**. The reference cut and screeners are read locally by the browser.

## Roll back

Copy the two files from `rollback-1.9/` back into the site root to restore the 1.9 behavior.

This package has not been pushed to GitHub or deployed.
