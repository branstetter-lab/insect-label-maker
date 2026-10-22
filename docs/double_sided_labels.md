# Double-sided insect specimen labels: research and printing considerations

Research checked 22 October 2026. Double-sided output is not implemented in the tool.

## Evidence

- **Practitioner discussion:** A 2019 [r/Entomology thread, “Labeling advice for a collection”](https://www.reddit.com/r/Entomology/comments/ajsrim/) describes placing coordinates on one side and date/locality on the other to save space. It establishes that the idea has been discussed; it is not a professional recommendation or a validation study.
- **Collection guidance:** The [SAFRINET manual, *Collecting and Preserving Insects and Arachnids*](https://assets.ippc.int/static/media/uploads/resources/collecting_and_preserving_insects_and_arachnids.pdf), printed page 82, advises against putting information on a label’s underside because it can be overlooked.
- **Institutional guidance:** The [IPPC *Guide to Delivering Phytosanitary Diagnostic Services*](https://www.ippc.int/static/media/files/publication/en/2018/06/Delivering_Phytosanitary_Diagnostic_Services_Guide_Pr2Final_WEB.pdf), section 8.2.4, printed page 106, advises against writing on both sides and allows host information on a second label. This is guidance for reference collections, not a universal legal requirement.
- **Published digitization experience:** [Paterson et al. (2016), *iCollections – Digitising the British and Irish Butterflies in the Natural History Museum, London*](https://doi.org/10.3897/BDJ.4.e9559), describes accommodating label data on both sides through multiple images and an explicit visual cue for the digitiser. It documents handling existing two-sided information; it does not establish that new duplex labels are preferable.

The search did not identify a controlled comparison showing that identical front/back specimen labels improve labeling speed or preservation. That is a limit of this search, not proof that no such study exists.

## Interpretation for this tool

Splitting essential data across the two sides risks hiding information from normal viewing and imaging workflows. Keeping all essential locality, date, collector, and identifier information on the front is consistent with the concerns above.

Printing an identical complete label on the back is a distinct proposal: the visible side would still contain the full data. A small duplicate-on-back trial could test the proposed handling advantage, but it should be treated as an experiment rather than established best practice. The sources above do not specifically validate duplicate-on-back printing.

## Practical design for a possible trial

This section is engineering judgment, not a conclusion from the cited literature.

- Front and back need matched positions and specimen IDs. Duplex back-page **positions** must be reflected for the chosen sheet-flip direction, while text and barcode artwork remain readable, not mirror-printed.
- Include calibration offsets and a test sheet. Front/back registration errors at this label size can put text or a barcode across a cut line.
- Decide the desired orientation after the label is flipped; test both that orientation and the printer’s long-edge/short-edge setting.
- Use opaque archival stock suitable for the printer. Check show-through, toner adhesion after a second pass, and barcode scanning from both sides of the actual paper.
- Test complete paired-label sheets before cutting. Separate habitat labels also need to map to their own matching reverse positions.
- A primary-label front should remain complete even if additional nonessential information is tried on the reverse. Distinct reverse information should be explicitly signposted so it is not silently missed.
