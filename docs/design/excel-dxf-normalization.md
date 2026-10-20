# Excel Differential Format Index Normalization

## Context

Excel stores differential formats (DXFs) as an ordered list in `xl/styles.xml` under `<dxfs>`. Tables and conditional-formatting rules refer to entries by zero-based index. Excel can reorder or rebuild this list when a workbook is saved, then update the references to preserve appearance. The resulting XML diff can be noisy even when each reference still resolves to an equivalent format.

An Excel save-and-extract round-trip reproduced this behavior. Excel reordered the DXFs, changed table references such as `dataDxfId` and `headerRowDxfId`, and rewrote `styles.xml`. Comparing each changed reference against the before-and-after DXF showed that the referenced definitions were semantically equivalent. The existing worksheet normalizer already receives both the freshly extracted package and the previous target directory, which provides the baseline needed for a similar DXF normalization pass.

## Goal

Make stable DXF definitions retain their previous indexes after extraction, so Excel's incidental reordering does not create tracked changes. Rewrite references to match the normalized list, preserving the formatting represented by the freshly saved workbook.

## Non-Goals

- Do not normalize cell styles (`cellXfs`), fonts, fills, borders, or number formats as part of this change.
- Do not suppress actual formatting changes. A changed DXF must remain visible as a new or changed definition with references updated to it.
- Do not normalize workbook timestamps, revision identifiers, absolute paths, or `vbaProject.bin` in this work.
- Do not remove or reorder worksheet normalization.

## Proposed Behavior

Implemented in `src/targets/transforms/normalize-dxf-indexes.ts` and integrated into target extraction.

Add a focused transform, tentatively `normalizeDxfIndexes(extractedDir, previousTargetDir)`, and call it during target extraction alongside `normalizeWorksheetNames`. It should run on the extracted OOXML package before that package replaces the project target. Like sheet normalization, it should be a no-op when there is no previous target or when either package lacks the required styles data.

1. Read the current and previous `xl/styles.xml` files and locate their `<dxfs>` lists. Treat list positions as the DXF indexes used by OOXML references.
2. Compute a deterministic semantic signature for each `<dxf>`. The signature should ignore irrelevant XML serialization differences, such as namespace-prefix and attribute-order changes, while preserving child-element order and all formatting values.
3. Match current DXFs against previous DXFs by signature. For duplicate signatures, match occurrences deterministically and one-to-one; identical duplicate definitions are interchangeable.
4. Construct the normalized list using the previous list as the index baseline:
   - Put each matched current definition back at its previous index.
   - Retain a previous definition at its old index if it is absent from the current list. It may still be referenced by another part, and retaining it avoids shifting every later index. An unused retained DXF is harmless but means the list may grow over time.
   - Append genuinely new current definitions after all previous indexes, in their current relative order.
   - Set the `<dxfs count>` to the resulting list length.
5. Build a mapping from each current DXF index to its normalized index. Rewrite all supported DXF-index references in the extracted package using that mapping.
6. Leave all unrelated XML and styles content unchanged.

Keeping absent previous entries is the proposed initial policy because the primary requirement is stable indexes, including when a format is removed. Garbage-collecting obsolete entries would shift indexes and should be considered separately, with an explicit migration strategy if desired.

## Reference Rewriting

The transform must inventory and cover every OOXML attribute whose value is a DXF index, not just the references seen in the demo. At minimum, verify handling for:

- `dxfId` on conditional-formatting rules in worksheet XML.
- `dataDxfId`, `headerRowDxfId`, and `totalsRowDxfId` on table-column/table XML elements.

Search worksheet and table XML structurally and update only recognized attributes with integer index values. Preserve missing attributes and non-DXF attributes. If a recognized reference is malformed or outside the current `<dxfs>` range, log the attribute, value, and XML path through the transform's debug logger, then leave it unchanged. Do not guess a replacement.

## Integration and Options

- Implement in `src/targets/transforms/normalize-dxf-indexes.ts`, keeping DXF logic separate from worksheet filename normalization.
- Integrate in `src/targets/export-target.ts` after extraction and before the extracted package is copied over the target.
- Keep the operation active for XML-only extraction as well as full extraction; it changes target XML only.
- Add `skipDxfIndexNormalization` to `ExportOptions` and a matching `--skip-dxf-index-normalization` option to extract/export commands, analogous to the worksheet-normalization opt-out. This provides a diagnostic escape hatch and allows isolated comparisons.
- Emit debug output summarizing matched, appended, and retained DXF counts. Avoid noisy default output for successful normalization.

## Tests

### Unit Tests

- Excel-reordered DXFs return to previous indexes and all references still resolve to equivalent definitions.
- References are remapped in conditional-formatting rules and table DXF attributes.
- New DXFs append without changing prior indexes; their references point to the appended entries.
- Removed DXFs remain in their previous slots, preventing shifts to surviving definitions.
- Duplicate equivalent DXFs are matched deterministically and one-to-one.
- Changes to a DXF definition are not mistaken for serialization noise; the changed definition is appended and references point to it.
- No previous target, missing styles, empty `<dxfs>`, and malformed/out-of-range references follow the documented no-op or diagnostic behavior.
- A second normalization pass is idempotent.

### End-to-End Test

The unit tests cover table references and conditional-formatting references. The Excel-backed XML-only extraction test simulates reordered DXFs and verifies that the target returns to its baseline order and remaps a conditional-formatting reference. Unit tests also verify that a real format change remains represented in the output.

## Implementation Sequence

1. Confirm the supported references and choose diagnostic logging with no rewrite for malformed or out-of-range indexes. **Complete.**
2. Add the transform and unit fixtures for reordered, duplicate, added, removed, and changed DXFs. **Complete.**
3. Integrate the transform into extraction and add the `--skip-dxf-index-normalization` option to extract/export. **Complete.**
4. Add extraction-level regression coverage and verify unit-test idempotence. **Complete.**
5. Re-run the original workbook save/extract scenario and confirm that DXF-only diffs disappear while real format edits remain visible. **Pending.**

## Risks and Open Questions

- Retaining removed DXFs preserves indexes but can accumulate unused definitions. This is intentional for the first version; cleanup is a separate policy decision.
- DXFs without a semantic match in the baseline must be treated as real additions/changes, even if Excel merely rewrote their XML representation in a way the signature does not yet normalize. Tests should expose which serialization differences need canonicalization.
- DXF equality must be semantic enough to ignore harmless serialization changes but strict enough not to collapse distinct formatting. Canonicalization should be conservative and covered by representative fixtures.
- Package extensions or future Excel features may introduce additional DXF-index attributes. The transform should be structured so supported reference locations are explicit, testable, and easy to extend.
