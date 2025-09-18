using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace AutoCADCleanupTool
{
    public partial class CleanupCommands
    {
        [CommandMethod("LISTPSXREFS", CommandFlags.Modal)]
        public static void ListPaperSpaceXrefs()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    LayoutManager lm = LayoutManager.Current;
                    if (lm == null)
                    {
                        ed.WriteMessage("\nLayout manager is unavailable; cannot inspect paper space.");
                        tr.Commit();
                        return;
                    }

                    DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    if (layoutDict == null)
                    {
                        ed.WriteMessage("\nLayout dictionary is unavailable; cannot inspect paper space.");
                        tr.Commit();
                        return;
                    }
                    string drawingDirectory = GetDrawingDirectory(doc, db);
                    var entries = new List<XrefReportEntry>();

                    foreach (DBDictionaryEntry entry in layoutDict)
                    {
                        Layout layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                        if (layout == null || layout.ModelType) continue;

                        BlockTableRecord btr = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                        if (btr == null) continue;

                        foreach (ObjectId entId in btr)
                        {
                            if (!entId.IsValid) continue;

                            BlockReference br = tr.GetObject(entId, OpenMode.ForRead, false) as BlockReference;
                            if (br == null) continue;

                            BlockTableRecord def = null;
                            try { def = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord; }
                            catch { }
                            if (def == null) continue;

                            if (!def.IsFromExternalReference && !def.IsFromOverlayReference) continue;

                            string blockName = def.Name;
                            string layer = br.Layer;
                            string status = def.XrefStatus.ToString();
                            string rawPath = def.PathName ?? string.Empty;
                            string absolutePath = ResolveXrefPath(db, rawPath, drawingDirectory);

                            Point3d[] footprint = null;
                            var extents = TryGetExtents(br);
                            if (extents != null)
                            {
                                footprint = BuildRectangleFromExtents(extents.Value);
                            }

                            entries.Add(new XrefReportEntry(layout.LayoutName, blockName, layer, status, absolutePath, footprint));
                        }
                    }

                    if (entries.Count == 0)
                    {
                        ed.WriteMessage("\nNo external references were detected in paper space layouts.");
                    }
                    else
                    {
                        ed.WriteMessage($"\nDetected {entries.Count} external reference insertion(s) in paper space.");
                        var indexedEntries = entries
                            .Select((entry, idx) => new XrefReportItem(entry, idx + 1))
                            .ToList();

                        foreach (var group in indexedEntries
                            .GroupBy(i => i.Entry.LayoutName, StringComparer.OrdinalIgnoreCase)
                            .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase))
                        {
                            ed.WriteMessage($"\n  Layout: {group.Key} ({group.Count()} XREF(s))");
                            foreach (var item in group)
                            {
                                string pathInfo = string.IsNullOrWhiteSpace(item.Entry.Path) ? "<no saved path>" : item.Entry.Path;
                                string zoomNote = item.Entry.HasZoomTarget ? string.Empty : " (no extents available)";
                                ed.WriteMessage($"\n    [{item.Index}] Block \"{item.Entry.BlockName}\" on layer \"{item.Entry.Layer}\" | Status: {item.Entry.Status} | Path: {pathInfo}{zoomNote}");
                            }
                        }

                        var prompt = new PromptIntegerOptions("\nEnter XREF number to zoom (Enter to finish): ")
                        {
                            AllowNone = true,
                            LowerLimit = 1,
                            UpperLimit = indexedEntries.Count
                        };

                        while (true)
                        {
                            PromptIntegerResult selection = ed.GetInteger(prompt);
                            if (selection.Status != PromptStatus.OK)
                            {
                                break;
                            }

                            XrefReportItem target = indexedEntries.FirstOrDefault(i => i.Index == selection.Value);
                            if (target == null)
                            {
                                ed.WriteMessage("\nSelection is out of range.");
                                continue;
                            }

                            if (!target.Entry.HasZoomTarget)
                            {
                                ed.WriteMessage("\nUnable to determine extents for the selected XREF.");
                                continue;
                            }

                            ZoomToTitleBlock(ed, target.Entry.Footprint);
                            ed.WriteMessage($"\nZoomed to XREF [{target.Index}] \"{target.Entry.BlockName}\".");
                        }
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to enumerate paper space XREFs: {ex.Message}");
            }
        }

        private static string ResolveXrefPath(Database db, string rawPath, string drawingDirectory)
        {
            if (string.IsNullOrWhiteSpace(rawPath))
            {
                return string.Empty;
            }

            try
            {
                string candidate = rawPath;

                if (!Path.IsPathRooted(candidate))
                {
                    if (!string.IsNullOrWhiteSpace(drawingDirectory))
                    {
                        candidate = Path.Combine(drawingDirectory, candidate);
                    }

                    if (!Path.IsPathRooted(candidate) && db != null)
                    {
                        try
                        {
                            string found = HostApplicationServices.Current?.FindFile(rawPath, db, FindFileHint.Default);
                            if (!string.IsNullOrWhiteSpace(found))
                            {
                                candidate = found;
                            }
                        }
                        catch
                        {
                        }
                    }

                    if (!Path.IsPathRooted(candidate))
                    {
                        return rawPath;
                    }
                }

                return Path.GetFullPath(candidate);
            }
            catch
            {
                return rawPath;
            }
        }

        private static string GetDrawingDirectory(Document doc, Database db)
        {
            try
            {
                if (doc != null && !string.IsNullOrWhiteSpace(doc.Name))
                {
                    string docDir = Path.GetDirectoryName(doc.Name);
                    if (!string.IsNullOrWhiteSpace(docDir))
                    {
                        return docDir;
                    }
                }

                if (db != null && !string.IsNullOrWhiteSpace(db.Filename))
                {
                    string dbDir = Path.GetDirectoryName(db.Filename);
                    if (!string.IsNullOrWhiteSpace(dbDir))
                    {
                        return dbDir;
                    }
                }
            }
            catch
            {
            }

            return string.Empty;
        }

        private static Point3d[] BuildRectangleFromExtents(Extents3d ext)
        {
            double width = Math.Abs(ext.MaxPoint.X - ext.MinPoint.X);
            double height = Math.Abs(ext.MaxPoint.Y - ext.MinPoint.Y);
            const double minSize = 1e-6;
            if (width < minSize || height < minSize)
            {
                return null;
            }

            return new[]
            {
                new Point3d(ext.MinPoint.X, ext.MinPoint.Y, 0),
                new Point3d(ext.MaxPoint.X, ext.MinPoint.Y, 0),
                new Point3d(ext.MaxPoint.X, ext.MaxPoint.Y, 0),
                new Point3d(ext.MinPoint.X, ext.MaxPoint.Y, 0)
            };
        }

        private sealed class XrefReportEntry
        {
            public XrefReportEntry(string layoutName, string blockName, string layer, string status, string path, Point3d[] footprint)
            {
                LayoutName = layoutName;
                BlockName = blockName;
                Layer = layer;
                Status = status;
                Path = path;
                Footprint = footprint;
            }

            public string LayoutName { get; }
            public string BlockName { get; }
            public string Layer { get; }
            public string Status { get; }
            public string Path { get; }
            public Point3d[] Footprint { get; }
            public bool HasZoomTarget => Footprint != null && Footprint.Length == 4;
        }

        private sealed class XrefReportItem
        {
            public XrefReportItem(XrefReportEntry entry, int index)
            {
                Entry = entry;
                Index = index;
            }

            public XrefReportEntry Entry { get; }
            public int Index { get; }
        }
    }
}
