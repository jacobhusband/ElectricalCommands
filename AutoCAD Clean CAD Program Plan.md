AutoCAD Clean CAD Program

## Goals
- Deliver a WPF desktop utility (no admin privileges) that guides users through cleaning project CAD assets for Mechanical, Electrical, and Plumbing disciplines.
- Provide granular control over the cleanup workflow including pause/cancel support and user confirmation when needed.
- Persist a detailed run report both in-app and as a .txt file saved under `%USERPROFILE%\Documents\CleanCAD Program`.

## User Workflow
1. Launch app (single-session settings).
2. User clicks "Select Project" to browse for the root directory. Show folder picker dialog.
3. Validate that `Mechanical`, `Electrical`, `Plumbing`, and `Xref` subdirectories exist. If any are missing, prompt user to manually locate each missing folder before proceeding.
4. Display checkboxes for M/E/P so user can choose which disciplines to process.
5. After selections, enable "Clean CAD Files" button to begin workflow.

## Processing Pipeline
- Copy selected discipline directories (and required `Xref`) into a new workspace under the system temp path (e.g. `%TEMP%\CleanCAD\<timestamp>`). Maintain directory structure. Plan to clean up temp folders automatically after a successful run; preserve or rename them if an error aborts execution for easier inspection.
- Track state so users can Pause, Resume, or Cancel during long operations. Cancelling stops new work immediately; pausing waits for the current AutoCAD command to complete before halting.
- For each drawing in the temp workspace, execute the cleaning sequence:
  1. Run the TBLK cleaning AutoCAD command set (exact commands TBD).
  2. Display a modal message box instructing the user to review the title block inside AutoCAD and click Continue or Cancel. Cancel aborts the run; Continue resumes.
  3. Run the remaining sheet-cleaning command set (exact commands TBD) for layouts beyond the title block.
  4. On any AutoCAD error, abort the entire pipeline, record details, and surface options to retry later.
- After commands succeed, rename each DWG according to its paper space layout names concatenated with spaces. Remove invalid filename characters and trim whitespace. If the resulting name is empty, skip renaming. On duplicate names, append an incrementing counter suffix (`_2`, `_3`, …).

## Reporting & UX
- Show live status/progress updates in the UI; include indicators for current drawing, command phase, and pause/cancel state.
- When the run completes or aborts, display the full report in a read-only multiline text box. Include:
  - Timestamp and selected disciplines.
  - Each processed file with outcome.
  - Any errors (including AutoCAD responses) and whether the run was aborted or canceled.
- Save the same report to `%USERPROFILE%\Documents\CleanCAD Program\RunReport-<timestamp>.txt`, creating the folder if needed.
- Offer buttons to Retry (re-run from the start) or Exit after viewing the report.

## Outstanding Details / Questions
- Provide the exact AutoCAD command sequences or scripts for both the TBLK cleaning and the remaining sheet cleanup. Are these LISP routines, command macros, or Document.SendCommand strings?
- Confirm whether Xref handling requires detaching, reloading, or copying additional references.
- Should we capture any AutoCAD version-specific considerations or guardrails for unsupported versions?
- Do we need localization (strings, date formats) or is English-only acceptable?
