# Master Order Helper 1.17 Test

Version 1.17 updates the A+E Media Resources one-by-one order review.

## A+E review changes

- The Previous Shot control is now below the current shot, alongside the review count and tape-title link. Reviewers can move backward through confirmed shots; editing a confirmed shot requires confirming it again.
- The Download A+E Order Form button appears beside the final shot's confirmation button after all shots are confirmed.
- After the final confirmation, the workbook is prepared in the background. When ready, a direct click on the download button saves the prepared Excel file. If preparation fails, the button offers a retry and shows the error beside it.
- Screenshot borders and review cards now use the smaller layout as well as the images. The 80% image size is preserved while the frame borders shrink to fit.
- Exact EDL Source Duration text from the selected Duration column is preserved in the order form.

Other Per Second vendors retain the existing scroll review and Google Sheet output flow.

## Install for a test

Copy `master-order-delivery.html` and `master-order-delivery.js` into the Lovely Runner site root, replacing the existing helper files. The page displays **1.17 Test** and loads the matching `1.17-test` script.

## Roll back

Copy both files from `rollback-1.9/` into the site root to restore version 1.9.

## Verification

The JavaScript syntax check and a mocked order-form build/download check passed. This package was not run in a browser against the live tracker or source video files.
