# CL - WS Creator

Excel add-in + local Windows bridge for turning a cutting list into a SigmaNEST job package from separate server geometry libraries.

Default PRS geometry path: S:\\SNDataX1\\PARTS.
Default DXF geometry path: Y:\\.

The add-in lets the user choose a separate CL workbook (.xlsx/.xls), preselects BAT sheets, consolidates CL quantities, recursively searches the PRS server tree for legacy .PRS parts and the DXF server tree for new .DXF geometry, checks .PRS material/thickness metadata when available, and stages the matched sources. A Summary and Part Review worksheet are written into the active Excel workbook, and the bridge also creates WS_PARTS.csv, PART_REVIEW.csv and SIGMANEST_JOB.json in the staging folder.

The two geometry roots are intentionally separate: SigmaNEST .PRS files remain under the existing PARTS server tree, while new/unestablished geometry is sourced from Y:\\.

SigmaNEST X1.4 SP3 (487) (X64) and Advanced Batch Processing are the target environment. SigmaNEST also documents a Part Create API, Basic Batch Commands, Advanced Batch Module and SDK; the final adapter should use the supported interface available in the customer's licensed installation.

The supplied production .PRS files are not included in this public repository.

See docs/SIGMANEST-INTEGRATION.md for the remaining SigmaNEST hand-off work.

Official reference: https://www.sigmanest.com/en/sigmanest
