# SigmaNEST X1.4 SP3 Integration

The supplied installation reports SigmaNEST X1.4 SP3 (487) (X64) and includes Advanced Batch Processing.

SigmaNEST's public product information lists a Part Create API, Basic Batch Commands, an Advanced Batch Module, and an SDK. The creator uses a local Windows bridge so the Excel task pane can access S:\\SNDataX1\\PARTS.

## Implemented in the current build

- CL sheet discovery and BAT-first selection.
- Flexible header discovery.
- Consolidation of duplicate Part + Material + Thickness lines.
- Default Part List path S:\\SNDataX1\\PARTS.
- Recursive .PRS indexing.
- Exact, embedded-name, and unambiguous variation matching.
- Lightweight .PRS metadata extraction.
- Material/thickness review checks.
- Job staging with WS_PARTS.csv and SIGMANEST_JOB.json.

## Next adapter step

The public pages retrieved during development confirm that SigmaNEST has batch/API/SDK automation capabilities, but they do not expose the precise X1.4 SP3 command syntax needed to create/load the work order from a job package.

The final adapter should be based on the batch/API sample or SDK installed with the customer's X1.4 SP3 deployment rather than undocumented editing of binary WS/PRS records.

Target: create/load SigmaNEST work order or task -> load the validated .PRS parts -> apply quantity/material/thickness through the supported interface -> hand control to the operator for physical nesting.

Official reference: https://www.sigmanest.com/en/sigmanest
