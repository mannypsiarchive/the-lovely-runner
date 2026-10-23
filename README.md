# Master Order Helper 1.10 test

This test replaces the Cut In and Cut Out still images with an inline player for the reference cut. Each player seeks to the EDL Cut In, plays picture and audio, and pauses at Cut Out. Screener Source In and Source Out stills remain for burnt-in timecode review and OCR.

## Install for a test

Copy `master-order-delivery.html` and `master-order-delivery.js` into the Lovely Runner site root, replacing the existing 1.9 pair. The existing home-page link should already point to `master-order-delivery.html`. The selected reference cut is read locally by the browser and is not uploaded.

The player is not muted. Actual playback depends on the browser supporting the video and audio codecs in the selected cut. If the player reports an unsupported format, test with an H.264 MP4 proxy.

## Roll back

Copy the two files from `rollback-1.9/` back into the site root to restore the 1.9 behavior.

The package has not been pushed to GitHub or deployed. Test it on the site before deciding whether to keep 1.10.
