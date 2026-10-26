# Version history

## Unreleased

- Added optional sampling protocol on separate habitat labels, retaining the three-line limit.
- Added compact pin-barcode labels with blank pin space and the main label’s barcode size.
- Added a seven-layout printable sampler, filled workbook, and matching blank all-options template.

- Determination labels prefix the determiner with “Det.” and start the authority on a new line when it cannot fit with the taxon.

- Adopted the display name Pittacia Insectorum; existing script and repository paths are unchanged.
- Added compact generic barcode labels and determination labels with Data Matrix and 4-point catalog numbers, with input templates and fictional examples.
- New modes omit collection-event codes and use a 1.05 barcode scale; locality-mode defaults are unchanged.

## 2026-10-26

- Added a minimal Excel input template, README label previews, and documentation of the lab printer and paper setup.
- Changed default barcode rendering from vector squares to 1200-dpi raster images while retaining vector text.
- Increased the default barcode scale to 1.02 relative to the earlier rendering, enlarging the symbol and its quiet margin together by 2%. Physical label dimensions remain unchanged.
- Adjusted the default Arial coordinate horizontal scale from .62 to .61 to accommodate the barcode change while retaining the preferred font size where the data fit.
- Added options for barcode rendering and scale, coordinate compression, and whole-label raster rendering.

These entries describe the current update; earlier changes remain available in the Git commit history.
