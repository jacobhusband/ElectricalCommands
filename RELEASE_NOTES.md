ACIES AutoCAD plugins v0.2.1

Fresh builds of all ten plugins, including AuditCommands, which was missing from the previous release assets. Every package includes a consistent installed-version marker for desktop update detection.

To install or update: close AutoCAD, open Tools > AutoCAD Plugins in the ACIES desktop application, and click Install or Update. Restart AutoCAD afterward. No desktop application rebuild is required to discover this release.

For manual installation, download the desired ElectricalCommands.<Plugin>-v0.2.1.zip asset and extract its contents into %APPDATA%\Autodesk\ApplicationPlugins\ElectricalCommands.<Plugin>.bundle. Replace existing files when updating; back up personally edited plugin files first. PackageContents.xml, version.txt, and Contents must be directly inside the .bundle folder. Restart AutoCAD.

Packages include .NET Framework 4.8 and .NET 8 plugin builds. Build and archive validation were performed; interactive AutoCAD command behavior has not been smoke-tested for this release.
