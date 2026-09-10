# ACIES AutoCAD plugins

Download compiled plugins from [GitHub Releases](https://github.com/jacobhusband/ElectricalCommands/releases/latest). No source build is needed to install them.

## Install or update

1. Close AutoCAD.
2. In the ACIES desktop application, open **Tools > AutoCAD Plugins** and click **Install** or **Update** for each desired plugin.
3. Restart AutoCAD.

For manual installation, download the desired `ElectricalCommands.<Plugin>-v<version>.zip` release asset. Create or open `%APPDATA%\Autodesk\ApplicationPlugins\ElectricalCommands.<Plugin>.bundle` and extract the ZIP contents into that folder, replacing existing files when updating. `PackageContents.xml`, `version.txt`, and `Contents` must sit directly inside the `.bundle` folder. Keep a backup of any personally edited plugin files before replacing them. Restart AutoCAD after installation.

The .NET plugin bundles contain separate builds for AutoCAD 2021–2024 (.NET Framework 4.8) and AutoCAD 2025+ (.NET 8). Runtime testing in the intended AutoCAD version is still recommended.

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
