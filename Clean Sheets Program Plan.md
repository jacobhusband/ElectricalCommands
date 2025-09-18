# Clean Sheets Program Plan

## Scope
- Clean one discipline at a time (Electrical, Mechanical, Plumbing) using the single known root folder for that discipline.
- Process only DWG files located directly in the discipline root directory (no recursion into subfolders).
- Discover and process DWG sheets in alphanumeric order; the first sheet is typically the cover sheet (e.g., `E00.00 Cover Sheet`, `E-0 Cover Sheet`).

## High-Level Workflow
1. **Gather context**: Confirm discipline selection, resolve the associated root directory, and retrieve the sorted list of DWG files to clean.
2. **Prepare AutoCAD session**: Open the first DWG, ensure required command methods are available, and capture any external title block metadata supplied by the upstream Title Block cleaner (if available).
3. **Paper space cleanup**
   - For each layout in the drawing:
     - Detect XREF insertions in paper space.
     - If the external metadata identifies the correct title block instance, use it; otherwise, list the found XREFs and prompt the user (via WPF dialog) to identify the title block—only one insertion should exist per layout; if multiples exist, require the user to click on the correct one.
     - Zoom to the chosen title block (mimicking `ZoomToTitleBlock`) and obtain its extents.
     - Select the title block and everything within its bounds.
     - Run an `EraseOther`-style routine to remove all paper space geometry outside the title block extents. Retain geometry that intersects or touches the boundary.
4. **Viewport-driven model space cleanup**
   - Return to the first paper space layout and iterate through viewports.
   - For each viewport:
     - Ensure the viewport is unfrozen/unlocked if necessary to support `CHSPACE`.
     - Create a temporary polyline matching the viewport boundary (polygon viewports handled natively).
     - Use `CHSPACE` to push the polyline into model space through the viewport.
   - Switch to model space, turn on and thaw all layers to expose all objects.
   - For each temporary polyline region now in model space:
     - Select geometry inside the region and mark it to keep.
   - Run `EraseOther` in model space to delete anything not captured by the viewport regions.
   - Delete the temporary polylines after cleanup.
5. **XREF handling**
   - Bind all DWG XREFs using `BIND` (never `INSERT`).
   - For PNG XREFs, embed them directly.
   - For PDF XREFs:
     - Ship the PdfiumViewer NuGet package along with its native Pdfium binaries so the WPF converter runs on client machines without extra installs.
     - From each XREF definition, extract the referenced page index and render only that page to PNG (no full-document conversion).
     - Target a configurable DPI (default 300) to balance clarity and file size, and store the output in a discipline-specific temporary folder.
     - Attach the generated PNG in place of the PDF and embed it, deleting the temporary files after a successful embed. Skip caching; re-render if the same PDF page appears again.
   - Run existing routines (e.g., `CLEANTBLK`) to purge ghost XREF data.
   - If embedding/conversion fails for a sheet, log the issue and continue with the remaining cleanup steps so the user can review manually.
6. **Verification**
   - Confirm no XREFs remain attached to the DWG before proceeding.
   - Optionally report any lingering issues to the user for manual follow-up.
7. **Output**
   - Save the cleaned DWG to `Documents\CleanCAD\<Project Name>\MEP\<Discipline>\`.
   - Name the file by concatenating the layout names with spaces (e.g., `E01.00 E01.01.dwg`).
   - Overwrite existing files with the same name.
8. **Iteration**
   - Advance to the next DWG in the sorted list and repeat steps 3–7 until all sheets are processed.

## Logging & User Feedback
- Maintain a per-sheet plain-text log in `Documents\CleanCAD\<Project Name>\Logs` summarizing actions taken, XREF decisions, and any failures (e.g., embedding issues) for user review.
- Use WPF dialogs for required user input (title block confirmation, failure notices) to keep interactions consistent.

## Error Handling
- If an AutoCAD command fails, attempt a retry when safe; otherwise record the failure and continue to the next actionable step.
- Preserve the original DWG if a catastrophic failure occurs during processing.



