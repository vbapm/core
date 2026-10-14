## Context

When exporting Excel worksheets in XML, the filenames are unstable and dependent on the order in which the worksheets appear in the workbook. This can lead to inconsistencies and difficulties in tracking worksheet identities across different exports. To address this, a normalization workflow is introduced to assign stable and predictable filenames based on the worksheet's VBA metadata.


The normalization workflow is:

1. **Export VBA and metadata**
   - `Build.ExportTo` exports VBA components.
   - For each worksheet component, it records:
     - VBA codename, e.g. `Sheet8`
     - visible tab name, e.g. `Historical Candidates`
   - This metadata is written to `project.json`.

2. **Extract worksheet XML**
   - Excel’s XML may use unstable names such as `sheet8.xml`.
   - `export-target.ts` passes the exported metadata and previous target path to the normalizer.

3. **Resolve each worksheet’s stable name**
   The normalizer uses this priority:

   1. Match the current workbook tab name to the VBA metadata:
      `Historical Candidates -> Sheet8 -> shtSheet8.xml`
   2. Fall back to `<sheetPr codeName="...">` in the worksheet XML.
   3. Fall back to the previous target’s `sheetId` mapping.
   4. If no identity is available, preserve the current filename.

4. **Rename related files together**
   When `sheet8.xml` becomes `shtSheet8.xml`, its sidecar is also renamed:

   `sheet8.xml.rels -> shtSheet8.xml.rels`

5. **Rewrite package references**
   The normalizer updates:

   - `xl/_rels/workbook.xml.rels`
   - `[Content_Types].xml`

6. **Prevent collisions**
   A destination filename can only be assigned once. This prevents two worksheet relationships from accidentally pointing to the same file.
