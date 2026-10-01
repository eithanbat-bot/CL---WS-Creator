# SigmaNEST X1.4 SP3 Integration

The supplied installation reports SigmaNEST X1.4 SP3 (487) (X64) and includes Advanced Batch Processing.

## Geometry sources

The creator uses a local Windows bridge so the Excel task pane can access two separate server trees:

- PRS root: S:\\SNDataX1\\PARTS
- DXF root: Y:\\

Both roots are searched recursively through all subfolders. PRS and DXF files are indexed separately, while part-name matching can consider both sources when selecting the geometry for a CL line.

## Implemented in the current build

- CL sheet discovery and BAT-first selection.
- Flexible header discovery.
- Consolidation of duplicate Part + Material + Thickness lines.
- Recursive .PRS indexing under the PRS root.
- Recursive .DXF indexing under the DXF root.
- Exact, embedded-name, and unambiguous variation matching across both file types.
- When the same exact part exists in both trees and there is exactly one PRS candidate, the PRS source is preferred.
- Lightweight .PRS metadata extraction; DXF sources use CL material/thickness as the authoritative values.
- Material/thickness review checks and explicit review reasons.
- Conflicting CL material/thickness variants are not silently combined.
- Job staging with WS_PARTS.csv, PART_REVIEW.csv and SIGMANEST_JOB.json.
- Excel report output with a CL WS Summary worksheet and a Part Review worksheet.

## Next adapter step

SigmaNEST publicly documents native DXF import plus Part Create API, batch commands and SDK capabilities, but the public pages do not expose the precise X1.4 SP3 COM/import signature needed for every DXF import path.

The final adapter should be based on the batch/API sample or SDK installed with the customer's X1.4 SP3 deployment rather than undocumented editing of binary WS/PRS records.

Target: create/load SigmaNEST work order or task -> load validated .PRS parts and import validated .DXF geometry -> apply quantity/material/thickness through the supported interface -> hand control to the operator for physical nesting.

Official reference: https://www.sigmanest.com/en/sigmanest
