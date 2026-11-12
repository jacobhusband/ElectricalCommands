# ACIES AutoCAD plugins

Download compiled plugins from [GitHub Releases](https://github.com/jacobhusband/ElectricalCommands/releases/latest). No source build is needed to install them.

## Install or update

1. Close AutoCAD.
2. In the ACIES desktop application, open **Tools > AutoCAD Plugins** and click **Install** or **Update** for each desired plugin.
3. Restart AutoCAD.

For manual installation, download the desired `ElectricalCommands.<Plugin>-v<version>.zip` release asset. Create or open `%APPDATA%\Autodesk\ApplicationPlugins\ElectricalCommands.<Plugin>.bundle` and extract the ZIP contents into that folder, replacing existing files when updating. `PackageContents.xml`, `version.txt`, and `Contents` must sit directly inside the `.bundle` folder. Keep a backup of any personally edited plugin files before replacing them. Restart AutoCAD after installation.

The .NET plugin bundles contain separate builds for AutoCAD 2021–2024 (.NET Framework 4.8) and AutoCAD 2025+ (.NET 8). Runtime testing in the intended AutoCAD version is still recommended.

## Scaled text with a leader

Run `TXTL` from GeneralCommands. Set the drawing scale with `SS` first, then
enter the text, pick the leader tip, and drag the text into place. The live
preview accepts `North`, `East`, `South`, and `West` for text direction, plus
`Arrow` and `Circle` for the leader end. TXTL uses TXT's text height and style
and places editable MText and a separate Leader on the current layer.
Escape cancels before either object is placed; one Undo removes both objects.

## Note tables

Run `NOTETABLE` from GeneralCommands to type keyed notes and general notes in a
window and draw them as tables. Each tab is one list of notes, numbered with
hexagon keyed-note symbols or plain "1." numbers. Add a tab for each plan as
needed. The tabs are saved in `NoteTables.json` in the drawing's folder, so
every sheet in that folder shares them. The first run in a drawing whose name
contains power, lighting (light, ltg), site, roof, demo, systems, fire, or sld
adds a keyed-note tab for that plan, such as "Power Plan".

`KNT` and `GNT` open the same window on a keyed-note or general-note tab.
**Import From Text** adds notes from MText or DBText, and **Import From
Table** adds the notes of an existing note table. **Create Table** saves the
tabs and then places the current tab's table, either in a drawn boundary or at
a picked corner. It can also replace the imported table in place or erase the
imported text. **Update Table** saves the tabs, then asks you to pick any
existing note table and redraws it with the current tab's notes and numbering.
The redrawn table keeps the old one's position, width, overflow and text size,
and one Undo restores the old table. `-KNT` and `-GNT` keep the original command-line versions,
including renumbering a selected table in place.

Note tables use the LFS and EQS text: ARIALNARROW-1-8 in color 4, with a
1/4" "KEYED NOTES" or "GENERAL NOTES" title over 3/32" notes (scaled for
model space). Each note's first line is centered on its keyed-note hexagon, and
notes are 1/8" apart. Renumbering with `-KNT` or `-GNT` reapplies this styling.

## Build and publish

Set the next version in `version.props`, then build into an empty staging directory:

```powershell
.\build-all-bundles.ps1 -SourceRoot build\release-bundles -OutputRoot AutoCADCommands\dist
```

Use `-DotnetPath <path-to-dotnet.exe>` when the SDK is not on PATH. Packaging requires all ten bundles, matching manifest versions, and existing entry-point modules. ZIPs have the flat layout expected by the desktop installer and contain BOM-free version markers.

Commit and push the release source, set `GITHUB_TOKEN` with repository release write access, then publish:

```powershell
.\publish-release.ps1 -NotesFile RELEASE_NOTES.md
```

The publisher creates a draft tied to the current commit, uploads and checks asset sizes, then publishes it as the latest release. An interrupted draft can be retried. Published versions cannot be overwritten; increment the version for subsequent updates.
