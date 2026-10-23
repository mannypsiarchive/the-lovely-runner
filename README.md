# Master Order Helper 1.12 test

This build retains the 1.11 Cut In and Cut Out screenshots, captured three frames inside their listed show timecodes, and show-time ordering. It adds timecode-aware tracker-column correction using actual cell values, shows sample values beside mapping controls, and checks the first visible reference-cut frame for BITC when file metadata reports a zero start timecode. For source screeners, it checks visible BITC only when the metadata-based seek would fall outside the file. Invalid cut times stop capture before black frames are produced; invalid source seeks and reversed source ranges are labeled for correction.

## Install for a test

Copy `master-order-delivery.html` and `master-order-delivery.js` into the Lovely Runner site root. The page visibly identifies itself as **1.12 Test**. The reference cut and screeners are read locally by the browser.

## Roll back

Copy the two files from `rollback-1.9/` back into the site root to restore 1.9.

This package has not been pushed to GitHub or deployed.
