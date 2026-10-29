# Master Order Helper 1.16 Test

Version 1.16 keeps the A+E Media Resources workflow from 1.15 and fixes the EDL duration mapping.

## A+E duration mapping

When the referenced tracker contains a header named **Duration**, the A+E helper selects that column automatically. For the current sheet, this is column J. Other trackers can still use a `Source Duration` or similar header, and the A+E-only dropdown shows the selected column letter and a sample value so it can be checked. Duration is displayed and written to the Excel form as the exact formatted cell text.

## A+E review

- One shot is shown at a time; screenshots are 20% smaller on screen.
- Each shot must be confirmed before it is included in the final order form. The download becomes available after all shots are confirmed.
- Tape IDs are detected from uploaded filenames containing a numeric segment immediately before `_CRS_`, such as `..._28804_CRS_...`. The tape title page lists detected IDs and matching ARC numbers. A title entered once fills all matching shots.

Other Per Second vendors retain the existing scroll review and Google Sheet output flow.

## Install for a test

Copy `master-order-delivery.html` and `master-order-delivery.js` into the Lovely Runner site root. The page visibly identifies itself as **1.16 Test**. ExcelJS is loaded from jsDelivr only when the A+E form download is requested.

## Roll back

Copy the two files from `rollback-1.9/` back into the site root to restore 1.9.

The mapping logic and ZIP package were checked locally. The browser flow has not been tested against the live tracker or actual video files.
