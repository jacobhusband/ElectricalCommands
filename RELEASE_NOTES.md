ACIES AutoCAD plugins v0.3.0

New commands and improvements across the plugins since v0.2.1:

- Note tables: NOTETABLE opens a tabbed editor for keyed notes and general notes, saved in NoteTables.json in the drawing's folder. KNT and GNT open it on a keyed-note or general-note tab; -KNT and -GNT keep the command-line versions.
- Equipment tags: ETAG places an oval equipment tag in a paper-space layout; ETAGOCR copies the text from an existing tag using bundled OCR.
- Scaled text: TXT places text sized to the drawing scale, and TXTL adds a leader with live direction and arrow/circle end options.
- Room templates: RTC captures, RTP places, and RTD lists room templates.
- Symbols: SWITCH, DATA, JBOX, and DISCONNECT symbol commands, plus KNL keyed-note leaders.
- Panel sheets: SPS can create a missing panel sheet from a template.
- CIRCUITREMAP remaps a run of panel circuits in every text, attribute, and multileader label.
- NEWLAYER (NL) creates a layer from a name and color and makes it current.
- BINDALLXREFS (BAX) binds all xrefs in the drawing.
- Electrical Drafting palette (EDP) opens when AutoCAD starts.

To install or update: close AutoCAD, open Tools > AutoCAD Plugins in the ACIES desktop application, and click Install or Update. Restart AutoCAD afterward.

For manual installation, download the desired ElectricalCommands.<Plugin>-v0.3.0.zip asset and extract its contents into %APPDATA%\Autodesk\ApplicationPlugins\ElectricalCommands.<Plugin>.bundle. Replace existing files when updating; back up personally edited plugin files first. PackageContents.xml, version.txt, and Contents must be directly inside the .bundle folder. Restart AutoCAD.

Packages include .NET Framework 4.8 and .NET 8 plugin builds.
