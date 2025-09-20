using Autodesk.AutoCAD.Runtime;

using Autodesk.AutoCAD.ApplicationServices;

using Autodesk.AutoCAD.DatabaseServices;

using Autodesk.AutoCAD.EditorInput;

using Autodesk.AutoCAD.Geometry;

using System;

using System.Collections.Generic;

using System.Linq;

using System.IO;

using System.Globalization;

namespace AutoCADCleanupTool

{

    public partial class CleanupCommands

    {

        [CommandMethod("CLEANPS", CommandFlags.Modal)]

        public static void ListPaperSpaceXrefs()

        {

            Document doc = Application.DocumentManager.MdiActiveDocument;

            if (doc == null) return;

            Database db = doc.Database;

            Editor ed = doc.Editor;

            try

            {

                var entries = new List<XrefReportEntry>();

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

                            Extents3d? extents = TryGetExtents(br);

                            if (extents != null)

                            {

                                footprint = BuildRectangleFromExtents(extents.Value);

                            }

                            var orientedFootprint = BuildBlockFootprint(br, def, tr);

                            if (orientedFootprint != null)

                            {

                                footprint = orientedFootprint;

                            }

                            entries.Add(new XrefReportEntry(br.ObjectId, layout.LayoutName, blockName, layer, status, absolutePath, footprint, extents));

                        }

                    }

                    tr.Commit();

                }

                if (entries.Count == 0)

                {

                    ed.WriteMessage("\nNo external references were detected in paper space layouts.");

                    return;

                }

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

                        if (item.Entry.Extents != null)

                        {

                            ed.WriteMessage($"\n      Extents: {DescribeExtents2D(item.Entry.Extents.Value)}");

                        }

                        else

                        {

                            ed.WriteMessage("\n      Extents: <not available>");

                        }

                    }

                }

                if (indexedEntries.Count == 1)

                {

                    var only = indexedEntries[0];

                    if (!only.Entry.HasZoomTarget)

                    {

                        ed.WriteMessage("\nNo extents available to zoom for the lone XREF.");

                        return;

                    }

                    if (TryZoomToEntry(doc, ed, only))

                    {

                        ed.WriteMessage($"\nZoomed to XREF [{only.Index}] \"{only.Entry.BlockName}\".");

                        PromptEraseOutside(doc, ed, db, only);

                    }

                    return;

                }

                var zoomableIndices = new HashSet<int>(indexedEntries

                    .Where(i => i.Entry.HasZoomTarget)

                    .Select(i => i.Index));

                if (zoomableIndices.Count == 0)

                {

                    ed.WriteMessage("\nNo zoomable XREFs were found.");

                    return;

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

                    if (!zoomableIndices.Contains(selection.Value))

                    {

                        ed.WriteMessage("\nSelected XREF does not have extents available for zoom.");

                        continue;

                    }

                    XrefReportItem target = indexedEntries.First(i => i.Index == selection.Value);

                    if (TryZoomToEntry(doc, ed, target))

                    {

                        ed.WriteMessage($"\nZoomed to XREF [{target.Index}] \"{target.Entry.BlockName}\".");

                        PromptEraseOutside(doc, ed, db, target);

                    }

                }

            }

            catch (System.Exception ex)

            {

                ed.WriteMessage($"\nFailed to enumerate paper space XREFs: {ex.Message}");

            }

        }

        private static bool TryZoomToEntry(Document doc, Editor ed, XrefReportItem target)

        {

            if (doc == null || ed == null || target == null)

            {

                return false;

            }



            if (!target.Entry.HasZoomTarget)

            {

                return false;

            }



            try

            {

                using (doc.LockDocument())

                {

                    if (!EnsurePaperSpaceLayoutActive(doc, ed, target.Entry.LayoutName))

                    {

                        return false;

                    }



                    ZoomToTitleBlock(ed, target.Entry.Footprint);

                }



                return true;

            }

            catch (System.Exception ex)

            {

                ed.WriteMessage($"\nFailed to zoom to XREF [{target.Index}]: {ex.Message}");

                return false;

            }

        }



        private static bool EnsurePaperSpaceLayoutActive(Document doc, Editor ed, string layoutName)

        {

            try

            {

                LayoutManager lm = LayoutManager.Current;

                if (lm == null)

                {

                    ed.WriteMessage("\nLayout manager is unavailable; cannot activate paper space.");

                    return false;

                }

                if (doc.Database.TileMode)

                {

                    Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("TILEMODE", 0);

                }

                if (!string.IsNullOrWhiteSpace(layoutName) &&

                    !string.Equals(lm.CurrentLayout, layoutName, StringComparison.OrdinalIgnoreCase))

                {

                    lm.CurrentLayout = layoutName;

                }

                try

                {

                    ed.SwitchToPaperSpace();

                }

                catch

                {

                }

                try

                {

                    Application.SetSystemVariable("CVPORT", 1);

                }

                catch

                {

                }

                return true;

            }

            catch (System.Exception ex)

            {

                ed.WriteMessage($"\nUnable to activate layout \"{layoutName}\": {ex.Message}");

                return false;

            }

        }

        private static void PromptEraseOutside(Document doc, Editor ed, Database db, XrefReportItem target)

        {

            if (doc == null || ed == null || db == null || target == null)

            {

                return;

            }



            if (target.Entry.Extents == null)

            {

                return;

            }



            var prompt = new PromptKeywordOptions("\nErase all paper space geometry outside this title block? [Yes/No] ")

            {

                AllowNone = true

            };

            prompt.Keywords.Add("Yes");

            prompt.Keywords.Add("No");

            prompt.Keywords.Default = "No";



            PromptResult result = ed.GetKeywords(prompt);

            if (result.Status == PromptStatus.Keyword &&

                string.Equals(result.StringResult, "Yes", StringComparison.OrdinalIgnoreCase))

            {

                int erased = TryEraseOutsideEntry(doc, ed, db, target);

                if (erased >= 0)

                {

                    if (erased > 0)

                    {

                        ed.WriteMessage("\nRemoved paper space geometry outside the title block extents.");

                    }

                    else

                    {

                        ed.WriteMessage("\nNo paper space geometry outside the title block extents was detected.");

                    }

                }

            }

        }



        private static int TryEraseOutsideEntry(Document doc, Editor ed, Database db, XrefReportItem target)

        {

            if (doc == null || ed == null || db == null || target == null || target.Entry.Extents == null)

            {

                return -1;

            }



            try

            {

                using (doc.LockDocument())

                {

                    if (!EnsurePaperSpaceLayoutActive(doc, ed, target.Entry.LayoutName))

                    {

                        return -1;

                    }



                    ObjectId layoutBtrId = GetLayoutBlockTableRecordId(db, target.Entry.LayoutName);

                    if (layoutBtrId.IsNull)

                    {

                        ed.WriteMessage($"\nUnable to locate layout \"{target.Entry.LayoutName}\" for cleanup.");

                        return -1;

                    }



                    Extents3d expanded = ExpandExtents(target.Entry.Extents.Value, 0.01);

                    var keepIds = new HashSet<ObjectId>();

                    if (target.Entry.BlockReferenceId != ObjectId.Null)

                    {

                        keepIds.Add(target.Entry.BlockReferenceId);

                    }



                    using (Transaction tr = db.TransactionManager.StartTransaction())

                    {

                        var btr = (BlockTableRecord)tr.GetObject(layoutBtrId, OpenMode.ForRead);

                        foreach (ObjectId entId in btr)

                        {

                            if (!entId.IsValid || keepIds.Contains(entId))

                            {

                                continue;

                            }



                            var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;

                            if (ent == null)

                            {

                                continue;

                            }



                            if (ent is Viewport vp && vp.Number == 1)

                            {

                                keepIds.Add(entId);

                                continue;

                            }



                            Extents3d? entExt = TryGetExtents(ent);

                            if (entExt != null && ExtentsContainsXY(expanded, entExt.Value))

                            {

                                keepIds.Add(entId);

                            }

                        }



                        tr.Commit();

                    }



                    if (keepIds.Count == 0)

                    {

                        ed.WriteMessage("\nErase skipped; no geometry detected within the title block bounds.");

                        return -1;

                    }



                    return EraseEntitiesExcept(db, ed, layoutBtrId, keepIds);

                }

            }

            catch (System.Exception ex)

            {

                ed.WriteMessage($"\nFailed to erase geometry outside the title block: {ex.Message}");

                return -1;

            }

        }



        private static ObjectId GetLayoutBlockTableRecordId(Database db, string layoutName)

        {

            if (db == null || string.IsNullOrWhiteSpace(layoutName))

            {

                return ObjectId.Null;

            }



            try

            {

                using (Transaction tr = db.TransactionManager.StartTransaction())

                {

                    var layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                    foreach (DBDictionaryEntry entry in layoutDict)

                    {

                        var layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;

                        if (layout != null && string.Equals(layout.LayoutName, layoutName, StringComparison.OrdinalIgnoreCase))

                        {

                            ObjectId result = layout.BlockTableRecordId;

                            tr.Commit();

                            return result;

                        }

                    }



                    tr.Commit();

                }

            }

            catch

            {

            }



            return ObjectId.Null;

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

        private static Point3d[] BuildBlockFootprint(BlockReference br, BlockTableRecord def, Transaction tr)

        {

            if (br == null || def == null || tr == null)

            {

                return null;

            }

            try

            {

                Extents3d? localExtents = null;

                foreach (ObjectId id in def)

                {

                    if (!id.IsValid) continue;

                    Entity ent = null;

                    try

                    {

                        ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity;

                    }

                    catch

                    {

                    }

                    if (ent == null) continue;

                    var entExt = TryGetExtents(ent);

                    if (entExt == null) continue;

                    if (localExtents == null)

                    {

                        localExtents = entExt.Value;

                    }

                    else

                    {

                        var agg = localExtents.Value;

                        agg.AddExtents(entExt.Value);

                        localExtents = agg;

                    }

                }

                if (localExtents == null)

                {

                    return null;

                }

                var ext = localExtents.Value;

                double width = Math.Abs(ext.MaxPoint.X - ext.MinPoint.X);

                double height = Math.Abs(ext.MaxPoint.Y - ext.MinPoint.Y);

                if (width < 1e-6 || height < 1e-6)

                {

                    return null;

                }

                Matrix3d transform = br.BlockTransform;

                var corners = new Point3d[4];

                corners[0] = new Point3d(ext.MinPoint.X, ext.MinPoint.Y, ext.MinPoint.Z).TransformBy(transform);

                corners[1] = new Point3d(ext.MaxPoint.X, ext.MinPoint.Y, ext.MinPoint.Z).TransformBy(transform);

                corners[2] = new Point3d(ext.MaxPoint.X, ext.MaxPoint.Y, ext.MaxPoint.Z).TransformBy(transform);

                corners[3] = new Point3d(ext.MinPoint.X, ext.MaxPoint.Y, ext.MaxPoint.Z).TransformBy(transform);

                for (int i = 0; i < corners.Length; i++)

                {

                    corners[i] = new Point3d(corners[i].X, corners[i].Y, 0.0);

                }

                return corners;

            }

            catch

            {

                return null;

            }

        }

        private static string FormatPoint2D(Point3d pt)

        {

            return string.Format(CultureInfo.InvariantCulture, "{0:F4}, {1:F4}", pt.X, pt.Y);

        }

        private static string DescribeExtents2D(Extents3d ext)

        {

            double width = Math.Abs(ext.MaxPoint.X - ext.MinPoint.X);

            double height = Math.Abs(ext.MaxPoint.Y - ext.MinPoint.Y);

            return string.Format(CultureInfo.InvariantCulture,

                "min ({0}) max ({1}) | width={2:F4}, height={3:F4}",

                FormatPoint2D(ext.MinPoint),

                FormatPoint2D(ext.MaxPoint),

                width,

                height);

        }

        private sealed class XrefReportEntry

        {

            public XrefReportEntry(ObjectId blockId, string layoutName, string blockName, string layer, string status, string path, Point3d[] footprint, Extents3d? extents)

            {

                BlockReferenceId = blockId;
                LayoutName = layoutName;

                BlockName = blockName;

                Layer = layer;

                Status = status;

                Path = path;

                Footprint = footprint;

                Extents = extents;

            }

            public ObjectId BlockReferenceId { get; }

            public string LayoutName { get; }

            public string BlockName { get; }

            public string Layer { get; }

            public string Status { get; }

            public string Path { get; }

            public Point3d[] Footprint { get; }

            public Extents3d? Extents { get; }

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




