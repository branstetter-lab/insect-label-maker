# Text-length guidance for the normal six-line layout

Use character counts as an editing guide, not an Excel entry restriction. Arial and Verdana are proportional fonts: `W` consumes much more width than `i`, and punctuation, spaces, italic glyphs, and horizontal compression all affect fitting. The generator measures rendered text width rather than counting characters.

## Practical starting points

For the default 0.68 × 0.33-inch locality label, 1.02 barcode scale, and the standard text boundary, the ordinary body-text column is approximately **31.44 points (11.09 mm)** wide. With the current font files and default horizontal scales, representative locality/collector wording gives these approximate capacities:

| Font / default body scale | At 4 pt | At 3.5 pt |
| --- | --- | --- |
| Arial / 75% width | About 20–23 characters | About 23–25 characters |
| Verdana / 70% width | About 22–24 characters | About 24–27 characters |

These are measured examples, not minimum or maximum allowed lengths. Counts include spaces and punctuation. They were calculated using repeated locality wording (`Tony Grove, 0.5km NNE `) and collector wording (`M.Branstetter, C.Meidt `), conservatively allowing for regular/italic character widths. Other wording will differ. For comparison, an Arial 4-point line holds only about 11 `W` characters but about 47 `i` characters. Font versions may vary.

A useful drafting target is **around 20 characters per ordinary line at 4 pt**, followed by checking the generated PDF/report. Do not discard useful collection information merely to satisfy that target. Store full data separately and abbreviate only display fields.

## Count the entire printed line

The budget is shared by all fields and separators on a line:

- Line 1: Country + `: ` + StateProv + `, ` + County.
- Line 2: LocalityLabel alone.
- Line 3: Latitude + space + Longitude + optional uncertainty. **Do not apply the body-text character estimates to coordinates:** these have their own compression and uncertainty font size. Keep meaningful coordinate precision and inspect the result.
- Line 4: Verbatim_Elevation + `, ` + Verbatim_Date.
- Line 5: Collector_Label. Spaces after initials are removed during rendering.
- Line 6: MacroHabitat_Label + `; ` + Microhabitat_Label. This line always uses 3.5 pt in the normal profile and may have less room depending on the Field Number. The full habitat may fall back to microhabitat alone, as recorded in the report.

Collection-only labels without barcodes, separate habitat labels, wrapped layouts, generic barcode labels, determination labels, and vertical Catalog Numbers have different width/height constraints. This table does not establish limits for them. Larger barcode scales, smaller label widths, alternate fonts, and custom horizontal scales change capacities.

## Entry guidance versus validation

The Excel templates include advisory header comments; they do not reject text based on length. That avoids blocking narrow text that fits or accepting wide text that does not. In the standard 4-point normal profile, lines 1–5 fall back individually to 3.5 pt when necessary. If no permitted layout fits, generation stops with the source row and Catalog Number. Review JSON `line_font_points`, `uncertainty_omitted`, and `macrohabitat_omitted` to see whether the desired size/data were retained.

A hard `LEN()` restriction is not recommended. If entry-time warnings are added later, they should be advisory and ideally based on the same font-width measurements as the generator. Existing career-master data are not modified by these suggestions.
