# CL - WS Creator

Excel add-in + local Windows bridge for turning a cutting list into a SigmaNEST job package from the local .PRS Part List.

Default SigmaNEST Part List path: S:\\SNDataX1\\PARTS.

The add-in preselects BAT sheets, consolidates CL quantities, matches .PRS geometry, checks material/thickness metadata when available, stages matched parts, and creates a machine-readable SIGMANEST_JOB.json plus WS_PARTS.csv.

SigmaNEST X1.4 SP3 (487) (X64) and Advanced Batch Processing are the target environment. SigmaNEST also documents a Part Create API, Basic Batch Commands, Advanced Batch Module and SDK; the final adapter should use the supported interface available in the customer's licensed installation.

The supplied production .PRS files are not included in this public repository.

See docs/SIGMANEST-INTEGRATION.md for the remaining SigmaNEST hand-off work.

Official reference: https://www.sigmanest.com/en/sigmanest
