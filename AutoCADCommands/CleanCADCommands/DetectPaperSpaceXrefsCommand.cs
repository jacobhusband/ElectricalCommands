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

            ed.Regen();

            try
            {
                var entries = new List<XrefReportEntry>();

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    LayoutManager lm = LayoutManager.Current;
                    if (lm == null)
                    {
                        ed.WriteMessage("\r\nLayout manager is unavailable; cannot inspect paper space.");
                        tr.Commit();
                        return;
                    }

                    DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                    if (layoutDict == null)
                    {
                        ed.WriteMessage("\r\nLayout dictionary is unavailable; cannot inspect paper space.");
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

                            // Do NOT open ForWrite unless actually modifying — avoids dirtying extents after UNDO
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

                            // Robust extents/footprint derivation — tolerant of UNDO wiping BR.GeometricExtents cache
                            Point3d[] footprint = null;
                            Extents3d? extents = null;

                            // 1) Try BR's own extents first
                            extents = TryGetExtents(br);

                            // 2) Try an orientation-aware footprint from the block contents (already world-space)
                            var orientedFootprint = BuildBlockFootprint(br, def, tr);
                            if (orientedFootprint != null)
                            {
                                footprint = orientedFootprint;
                                if (extents == null)
                                    extents = ExtentsFromPoints(orientedFootprint);
                            }

                            // 3) Still nothing? Transform definition extents by BR transform
                            if (extents == null)
                            {
                                var defExt = TryGetBlockDefExtents(def, tr);
                                if (defExt != null)
                                {
                                    var worldExt = defExt.Value;
                                    worldExt.TransformBy(br.BlockTransform);

                                    // Flatten Z for paper space usage
                                    extents = new Extents3d(
                                        new Point3d(worldExt.MinPoint.X, worldExt.MinPoint.Y, 0),
                                        new Point3d(worldExt.MaxPoint.X, worldExt.MaxPoint.Y, 0)
                                    );

                                    if (footprint == null)
                                        footprint = BuildRectangleFromExtents(extents.Value);
                                }
                            }

                            // 4) Last resort — rectangle from whatever extents we have
                            if (footprint == null && extents != null)
                            {
                                footprint = BuildRectangleFromExtents(extents.Value);
                            }

                            entries.Add(new XrefReportEntry(br.ObjectId, layout.LayoutName, blockName, layer, status, absolutePath, footprint, extents));
                        }
                    }

                    tr.Commit();
                }

                if (entries.Count == 0)
                {
                    ed.WriteMessage("\r\nNo external references were detected in paper space layouts.");
                    return;
                }

                ed.WriteMessage($"\r\nDetected {entries.Count} external reference insertion(s) in paper space.");

                var indexedEntries = entries
                    .Select((entry, idx) => new XrefReportItem(entry, idx + 1))
                    .ToList();

                var groupedByLayout = indexedEntries
                    .GroupBy(i => i.Entry.LayoutName, StringComparer.OrdinalIgnoreCase)
                    .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                int autoCleanedLayouts = 0;

                foreach (var group in groupedByLayout)
                {
                    if (group.Count() == 1)
                    {
                        var target = group.First();
                        if (target.Entry.Extents == null)
                        {
                            ed.WriteMessage($"\r\n  Layout: {group.Key}: Skipping automatic cleanup; extents not available for this XREF.");
                            continue;
                        }

                        int erased = TryEraseOutsideEntry(doc, ed, db, target);
                        if (erased >= 0)
                        {
                            autoCleanedLayouts++;
                            if (erased > 0)
                            {
                                ed.WriteMessage($"\r\n  Layout: {group.Key}: Removed paper space geometry outside the title block extents.");
                            }
                            else
                            {
                                ed.WriteMessage($"\r\n  Layout: {group.Key}: No paper space geometry outside the title block extents was detected.");
                            }
                        }
                    }
                    else
                    {
                        ed.WriteMessage($"\r\n  Layout: {group.Key} ({group.Count()} XREF(s)): Skipping automatic cleanup; multiple XREF block references detected on this layout.");
                    }
                }

                if (autoCleanedLayouts > 0)
                {
                    ed.WriteMessage($"\r\nAutomatically cleaned {autoCleanedLayouts} layout(s) containing a single XREF.");
                }
                else
                {
                    ed.WriteMessage("\r\nNo layouts were automatically cleaned.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\r\nFailed to enumerate paper space XREFs: {ex.Message}");
            }
        }

        private static bool EnsurePaperSpaceLayoutActive(Document doc, Editor ed, string layoutName)
        {
            try
            {
                LayoutManager lm = LayoutManager.Current;
                if (lm == null)
                {
                    ed.WriteMessage("\r\nLayout manager is unavailable; cannot activate paper space.");
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

                try { ed.SwitchToPaperSpace(); } catch { }
                try { Application.SetSystemVariable("CVPORT", 1); } catch { }

                return true;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\r\nUnable to activate layout \"{layoutName}\": {ex.Message}");
                return false;
            }
        }

        private static int TryEraseOutsideEntry(Document doc, Editor ed, Database db, XrefReportItem target)
        {
            if (doc == null || ed == null || db == null || target == null || target.Entry.Extents == null)
                return -1;

            try
            {
                using (doc.LockDocument())
                {
                    if (!EnsurePaperSpaceLayoutActive(doc, ed, target.Entry.LayoutName))
                        return -1;

                    ObjectId layoutBtrId = GetLayoutBlockTableRecordId(db, target.Entry.LayoutName);
                    if (layoutBtrId.IsNull)
                    {
                        ed.WriteMessage($"\r\nUnable to locate layout \"{target.Entry.LayoutName}\" for cleanup.");
                        return -1;
                    }

                    Extents3d expanded = ExpandExtents(target.Entry.Extents.Value, 0.01);

                    var keepIds = new HashSet<ObjectId>();
                    if (target.Entry.BlockReferenceId != ObjectId.Null)
                        keepIds.Add(target.Entry.BlockReferenceId);

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        var btr = (BlockTableRecord)tr.GetObject(layoutBtrId, OpenMode.ForRead);
                        foreach (ObjectId entId in btr)
                        {
                            if (!entId.IsValid || keepIds.Contains(entId))
                                continue;

                            var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                            if (ent == null)
                                continue;

                            if (ent is Viewport vp && vp.Number == 1)
                            {
                                keepIds.Add(entId);
                                continue;
                            }

                            Extents3d? entExt = TryGetExtents(ent);
                            if (entExt != null && ExtentsIntersectsXY(expanded, entExt.Value, includeTouch: true))
                                keepIds.Add(entId);
                        }

                        tr.Commit();
                    }

                    if (keepIds.Count == 0)
                    {
                        ed.WriteMessage("\r\nErase skipped; no geometry detected within the title block bounds.");
                        return -1;
                    }

                    return EraseEntitiesExcept(db, ed, layoutBtrId, keepIds);
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\r\nFailed to erase geometry outside the title block: {ex.Message}");
                return -1;
            }
        }

        private static ObjectId GetLayoutBlockTableRecordId(Database db, string layoutName)
        {
            if (db == null || string.IsNullOrWhiteSpace(layoutName))
                return ObjectId.Null;

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
            catch { }

            return ObjectId.Null;
        }

        private static string ResolveXrefPath(Database db, string rawPath, string drawingDirectory)
        {
            if (string.IsNullOrWhiteSpace(rawPath))
                return string.Empty;

            try
            {
                string candidate = rawPath;
                if (!Path.IsPathRooted(candidate))
                {
                    if (!string.IsNullOrWhiteSpace(drawingDirectory))
                        candidate = Path.Combine(drawingDirectory, candidate);

                    if (!Path.IsPathRooted(candidate) && db != null)
                    {
                        try
                        {
                            string found = HostApplicationServices.Current?.FindFile(rawPath, db, FindFileHint.Default);
                            if (!string.IsNullOrWhiteSpace(found))
                                candidate = found;
                        }
                        catch { }
                    }

                    if (!Path.IsPathRooted(candidate))
                        return rawPath;
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
                        return docDir;
                }

                if (db != null && !string.IsNullOrWhiteSpace(db.Filename))
                {
                    string dbDir = Path.GetDirectoryName(db.Filename);
                    if (!string.IsNullOrWhiteSpace(dbDir))
                        return dbDir;
                }
            }
            catch { }

            return string.Empty;
        }

        private static Point3d[] BuildRectangleFromExtents(Extents3d ext)
        {
            double width = Math.Abs(ext.MaxPoint.X - ext.MinPoint.X);
            double height = Math.Abs(ext.MaxPoint.Y - ext.MinPoint.Y);
            const double minSize = 1e-6;
            if (width < minSize || height < minSize)
                return null;

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
                return null;

            try
            {
                Extents3d? localExtents = null;

                foreach (ObjectId id in def)
                {
                    if (!id.IsValid) continue;

                    Entity ent = null;
                    try { ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity; }
                    catch { }
                    if (ent == null) continue;

                    var entExt = TryGetExtents(ent);
                    if (entExt == null) continue;

                    if (localExtents == null)
                        localExtents = entExt.Value;
                    else
                    {
                        var agg = localExtents.Value;
                        agg.AddExtents(entExt.Value);
                        localExtents = agg;
                    }
                }

                if (localExtents == null)
                    return null;

                var ext = localExtents.Value;
                double width = Math.Abs(ext.MaxPoint.X - ext.MinPoint.X);
                double height = Math.Abs(ext.MaxPoint.Y - ext.MinPoint.Y);
                if (width < 1e-6 || height < 1e-6)
                    return null;

                Matrix3d transform = br.BlockTransform;
                var corners = new Point3d[4];
                corners[0] = new Point3d(ext.MinPoint.X, ext.MinPoint.Y, ext.MinPoint.Z).TransformBy(transform);
                corners[1] = new Point3d(ext.MaxPoint.X, ext.MinPoint.Y, ext.MinPoint.Z).TransformBy(transform);
                corners[2] = new Point3d(ext.MaxPoint.X, ext.MaxPoint.Y, ext.MaxPoint.Z).TransformBy(transform);
                corners[3] = new Point3d(ext.MinPoint.X, ext.MaxPoint.Y, ext.MaxPoint.Z).TransformBy(transform);

                // flatten Z for paper space
                for (int i = 0; i < corners.Length; i++)
                    corners[i] = new Point3d(corners[i].X, corners[i].Y, 0.0);

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

        // ---------- NEW HELPERS (robust extents fallbacks) ----------
        private static Extents3d? ExtentsFromPoints(Point3d[] pts)
        {
            if (pts == null || pts.Length == 0) return null;
            double minX = pts[0].X, minY = pts[0].Y, maxX = pts[0].X, maxY = pts[0].Y;
            for (int i = 1; i < pts.Length; i++)
            {
                if (pts[i].X < minX) minX = pts[i].X;
                if (pts[i].Y < minY) minY = pts[i].Y;
                if (pts[i].X > maxX) maxX = pts[i].X;
                if (pts[i].Y > maxY) maxY = pts[i].Y;
            }
            return new Extents3d(new Point3d(minX, minY, 0), new Point3d(maxX, maxY, 0));
        }

        // Keep if the two AABBs overlap in XY (touching counts as overlap)
        private static bool ExtentsIntersectsXY(Extents3d a, Extents3d b, bool includeTouch = true, double tol = 1e-6)
        {
            // Flatten Z (defensive; your extents are already flattened elsewhere)
            var aMinX = Math.Min(a.MinPoint.X, a.MaxPoint.X);
            var aMaxX = Math.Max(a.MinPoint.X, a.MaxPoint.X);
            var aMinY = Math.Min(a.MinPoint.Y, a.MaxPoint.Y);
            var aMaxY = Math.Max(a.MinPoint.Y, a.MaxPoint.Y);

            var bMinX = Math.Min(b.MinPoint.X, b.MaxPoint.X);
            var bMaxX = Math.Max(b.MinPoint.X, b.MaxPoint.X);
            var bMinY = Math.Min(b.MinPoint.Y, b.MaxPoint.Y);
            var bMaxY = Math.Max(b.MinPoint.Y, b.MaxPoint.Y);

            if (includeTouch)
            {
                return (aMinX <= bMaxX + tol) && (aMaxX + tol >= bMinX) &&
                       (aMinY <= bMaxY + tol) && (aMaxY + tol >= bMinY);
            }
            else
            {
                return (aMinX < bMaxX - tol) && (aMaxX - tol > bMinX) &&
                       (aMinY < bMaxY - tol) && (aMaxY - tol > bMinY);
            }
        }


        private static Extents3d? TryGetBlockDefExtents(BlockTableRecord def, Transaction tr)
        {
            if (def == null || tr == null) return null;

            Extents3d? localExtents = null;
            foreach (ObjectId id in def)
            {
                if (!id.IsValid) continue;

                Entity ent = null;
                try { ent = tr.GetObject(id, OpenMode.ForRead, false) as Entity; }
                catch { }
                if (ent == null) continue;

                var eext = TryGetExtents(ent);
                if (eext == null) continue;

                if (localExtents == null)
                    localExtents = eext.Value;
                else
                {
                    var agg = localExtents.Value;
                    agg.AddExtents(eext.Value);
                    localExtents = agg;
                }
            }
            return localExtents;
        }
    }
}
