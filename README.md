# Master Order Helper 1.14 test

For a Per Second source named exactly **A+E Media Resources**, the shot review adds A+E Tape #, Program/Tape Title, and Clip/Shot Description fields. After timecodes are confirmed and the required fields are complete, **Download A+E Order Form** creates a downloadable Excel form matching the attached Hi-Res Clip Order Form layout. Source TC In/Out come from the editable on-screen timecode fields, duration comes from the A+E-only EDL Source Duration column selector, and the Image column uses the screener Source In frame. The Project/Episode line defaults to the connected tracker title and can be edited.

All other Per Second vendors retain the existing review and Google Sheet output. Cut screenshots remain three frames inside the listed boundaries, and the progressive screenshot loading from 1.13 remains.

## Install for a test

Copy `master-order-delivery.html` and `master-order-delivery.js` into the Lovely Runner site root. The page visibly identifies itself as **1.14 Test**. ExcelJS is loaded from jsDelivr only when the A+E form download is requested.

## Roll back

Copy the two files from `rollback-1.9/` back into the site root to restore 1.9.

This package has not been pushed to GitHub or deployed.
