# Master Order Helper 1.15 Test

Version 1.15 changes the Per Second flow only when the selected vendor is exactly **A+E Media Resources**.

## A+E review changes

- The review is click-through, with one shot shown at a time. Each shot's four screenshots are displayed at 80% of the prior on-screen width.
- Confirming a shot adds it to the final A+E order form. The form download becomes available after every shot is confirmed. Previous Shot and Edit Tape Titles allow corrections.
- The uploaded elements folder is scanned for A+E Tape IDs in names containing a numeric segment immediately before `_CRS_`, such as `..._28804_CRS_...`. A title page lists detected IDs, file counts, and matching ARC numbers. Enter a Program/Tape Title once to apply it to every matching shot.
- The A+E Source Duration column is selected using both header and displayed values. The selected column's exact formatted text is shown in review and copied as text to the Excel order form, preserving timecode values such as `00:00:05:12`.

All other Per Second vendors retain the existing scroll review, screenshot loading, and Google Sheet output behavior.

## Install for a test

Copy `master-order-delivery.html` and `master-order-delivery.js` into the Lovely Runner site root. The page visibly identifies itself as **1.15 Test**. ExcelJS is loaded from jsDelivr only when the A+E form download is requested.

## Roll back

Copy the two files from `rollback-1.9/` back into the site root to restore 1.9.

This package has not been pushed to GitHub or deployed. The browser flow has not been tested against the live tracker or actual video files.
