# CL - WS Creator release tests

Run these tests from the checked-out repository on the SigmaNEST workstation. No extra software installation is required beyond existing Windows PowerShell and SigmaNEST.

**Before running `-RealCom` or `-HttpE2E`, save and close any working SigmaNEST job.** The COM harness uses `ResetSigmaNEST`/`FileNew` in the running SigmaNEST session, but it writes its test workspaces under the Windows TEMP folder and does not save over production `.ws` files.

## Local gates (no live SigmaNEST operations)

Run from the repository root:

    powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\tests\run-sigmanest-release-gate.ps1"

This runs static architecture/release checks and the deterministic COM mock. GitHub Actions runs the same two gates on pushes.

## Real SigmaNEST COM test

    powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\tests\run-sigmanest-release-gate.ps1" -RealCom

This runs a real single-DXF import and a mixed DXF+PRS import through SigmaNEST, applies the CL values, creates tasks, applies batch/order labels and reload-verifies the saved results. The mixed-source harness retains its diagnostic workspace in TEMP.

By default the test DXF is `Y:\AutoCAD\Hino old\Hino 300\Rear Bodies\SBV\2023 (HSW) SBV Hino 300 816 Body\Dxf Files\PC-2A.DXF`. Override it with `-DxfPath "Y:\path\to\another-part.DXF"` and optionally pass `-TestQty`, `-TestMaterial`, `-TestThicknessMm` and `-BatchMultiplier`.

## Live bridge HTTP end-to-end test

First start the normal `Start Bridge Fixed.bat` launcher and ensure SigmaNEST is available. Then run:

    powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\tests\run-sigmanest-release-gate.ps1" -HttpE2E

This submits a disposable PC-2A job through the actual bridge APIs, waits for Import Geometry and AutoTask to finish, reloads the saved workspace, and verifies DXF sourcing, material, thickness, Number To Nest, task label, batch multiplier and task quantity. It removes the disposable test workspace, request staging folder and test job status files.

## Full workstation release gate

    powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\tests\run-sigmanest-release-gate.ps1" -RealCom -HttpE2E

A full pass is the release check before sending a runtime update. GitHub Actions cannot execute COM-backed tests because SigmaNEST is installed only on the workstation; the real COM and HTTP tests run there.