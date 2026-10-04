using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AutoCADCleanupTool
{
    // BINDALLXREFS: strips every data link, PDF underlay and raster image from the drawing,
    // then binds every DWG XREF whose file is found. Media nested inside an XREF only becomes
    // removable once that XREF is bound, so the media removal runs again after the bind.
    public static class BindAllXrefsCommand
    {
        private const string ImageDictionaryName = "ACAD_IMAGE_DICT";
        private const string PdfDictionaryName = "ACAD_PDFDEFINITIONS";

        // Binding an XREF that contains nested XREFs can leave the nested ones attached to the
        // host, so keep binding until a pass makes no progress.
        private const int MaxBindPasses = 5;

        [CommandMethod("BINDALLXREFS", CommandFlags.Modal)]
        [CommandMethod("BAX", CommandFlags.Modal)]
        public static void BindAllXrefs()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            ed.WriteMessage(
                "\nBind keeps each XREF's layers separate (XREF$0$Layer); " +
                "Insert merges them into matching layers in this drawing.");
            var bindTypeOptions = new PromptKeywordOptions(
                "\nBind type [Bind/Insert] <Bind>: ",
                "Bind Insert")
            {
                AllowNone = true,
            };
            bindTypeOptions.Keywords.Default = "Bind";
            PromptResult bindTypeResult = ed.GetKeywords(bindTypeOptions);
            if (bindTypeResult.Status != PromptStatus.OK &&
                bindTypeResult.Status != PromptStatus.None)
            {
                ed.WriteMessage("\nBINDALLXREFS canceled.");
                return;
            }
            bool insertBind = bindTypeResult.Status == PromptStatus.OK &&
                string.Equals(bindTypeResult.StringResult, "Insert", StringComparison.OrdinalIgnoreCase);

            var totals = new MediaTotals();
            var bindFailures = new List<string>();
            int bound = 0;

            try
            {
                ed.WriteMessage("\n--- BINDALLXREFS: removing data links, PDFs and images ---");
                RemoveExternalMedia(db, ed, totals);

                ed.WriteMessage("\n--- BINDALLXREFS: binding DWG XREFs ---");
                bound = BindResolvedDwgXrefs(db, ed, insertBind, bindFailures);

                if (bound > 0)
                {
                    ed.WriteMessage("\n--- BINDALLXREFS: removing media brought in by the bound XREFs ---");
                    RemoveExternalMedia(db, ed, totals);
                }

            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nBINDALLXREFS stopped: {ex.Message}");
            }

            try { ed.Regen(); } catch { }

            WriteSummary(db, ed, totals, bound, insertBind, bindFailures);
        }

        // ---------------------------------------------------------------------------------
        // Media removal
        // ---------------------------------------------------------------------------------

        private sealed class MediaTotals
        {
            internal int DataLinks;
            internal int TablesUnlinked;
            internal int ImageReferences;
            internal int ImageDefinitions;
            internal int PdfReferences;
            internal int PdfDefinitions;
            internal int Failures;
        }

        private static void RemoveExternalMedia(Database db, Editor ed, MediaTotals totals)
        {
            RemoveDataLinks(db, ed, totals);
            RemoveMediaReferences(db, ed, totals);
            totals.ImageDefinitions += RemoveDefinitions(db, ed, ImageDictionaryName, "image", totals);
            totals.PdfDefinitions += RemoveDefinitions(db, ed, PdfDictionaryName, "PDF", totals);
        }

        // XREF definitions hold read-only copies of another drawing's objects, and their
        // content can't be edited from here. Skip them; bound XREFs become ordinary blocks.
        private static bool IsXrefBlock(BlockTableRecord btr)
        {
            return btr.IsFromExternalReference || btr.IsFromOverlayReference || btr.IsDependent;
        }

        private static void RemoveDataLinks(Database db, Editor ed, MediaTotals totals)
        {
            DataLinkManager manager = db.DataLinkManager;
            if (manager.DataLinkCount == 0) return;

            // A table that still uses a data link would be left pointing at a deleted link,
            // so unlink the table cells first. The cell text is kept.
            var linkedTables = new List<ObjectId>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId btrId in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    if (IsXrefBlock(btr)) continue;

                    foreach (ObjectId entId in btr)
                    {
                        if (!string.Equals(entId.ObjectClass.DxfName, "ACAD_TABLE", StringComparison.OrdinalIgnoreCase)) continue;
                        try
                        {
                            var table = tr.GetObject(entId, OpenMode.ForRead) as Table;
                            ObjectIdCollection links = table?.GetDataLink();
                            if (links != null && links.Count > 0) linkedTables.Add(entId);
                        }
                        catch (System.Exception ex)
                        {
                            totals.Failures++;
                            ed.WriteMessage($"\nCould not inspect table {entId} for data links: {ex.Message}");
                        }
                    }
                }
                tr.Commit();
            }

            if (linkedTables.Count > 0)
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId tableId in linkedTables)
                    {
                        try
                        {
                            // forceOpenOnLockedLayer avoids changing the layer's lock state.
                            var table = (Table)tr.GetObject(tableId, OpenMode.ForWrite, false, true);
                            table.RemoveDataLink();
                            totals.TablesUnlinked++;
                        }
                        catch (System.Exception ex)
                        {
                            totals.Failures++;
                            ed.WriteMessage($"\nCould not unlink table {tableId} from its data link: {ex.Message}");
                        }
                    }
                    tr.Commit();
                }
            }

            ObjectIdCollection linkIds = manager.GetDataLink();
            foreach (ObjectId linkId in linkIds)
            {
                try
                {
                    manager.RemoveDataLink(linkId);
                    totals.DataLinks++;
                }
                catch (System.Exception ex)
                {
                    totals.Failures++;
                    ed.WriteMessage($"\nCould not remove data link {linkId}: {ex.Message}");
                }
            }
        }

        private static void RemoveMediaReferences(Database db, Editor ed, MediaTotals totals)
        {
            // Match on the DXF name, not the .NET type: Wipeout derives from RasterImage and
            // must be left alone.
            var imageRefs = new List<ObjectId>();
            var pdfRefs = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId btrId in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    if (IsXrefBlock(btr)) continue;

                    foreach (ObjectId entId in btr)
                    {
                        string dxf = entId.ObjectClass.DxfName;
                        if (string.Equals(dxf, "IMAGE", StringComparison.OrdinalIgnoreCase))
                        {
                            imageRefs.Add(entId);
                        }
                        else if (string.Equals(dxf, "PDFUNDERLAY", StringComparison.OrdinalIgnoreCase))
                        {
                            pdfRefs.Add(entId);
                        }
                    }
                }
                tr.Commit();
            }

            if (imageRefs.Count == 0 && pdfRefs.Count == 0) return;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                totals.ImageReferences += EraseEntities(tr, ed, imageRefs, "image", totals);
                totals.PdfReferences += EraseEntities(tr, ed, pdfRefs, "PDF underlay", totals);
                tr.Commit();
            }
        }

        private static int EraseEntities(Transaction tr, Editor ed, List<ObjectId> ids, string label, MediaTotals totals)
        {
            int erased = 0;
            foreach (ObjectId id in ids)
            {
                try
                {
                    // forceOpenOnLockedLayer avoids changing the layer's lock state.
                    var entity = tr.GetObject(id, OpenMode.ForWrite, false, true) as Entity;
                    if (entity == null || entity.IsErased) continue;
                    entity.Erase();
                    erased++;
                }
                catch (System.Exception ex)
                {
                    totals.Failures++;
                    ed.WriteMessage($"\nCould not erase {label} reference {id}: {ex.Message}");
                }
            }
            return erased;
        }

        // Erases every definition in the image or PDF dictionary. Entries named "xref|name"
        // belong to an XREF and can't be removed until that XREF is bound.
        private static int RemoveDefinitions(Database db, Editor ed, string dictionaryName, string label, MediaTotals totals)
        {
            int removed = 0;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                if (!nod.Contains(dictionaryName))
                {
                    tr.Commit();
                    return 0;
                }

                var dictionary = (DBDictionary)tr.GetObject(nod.GetAt(dictionaryName), OpenMode.ForRead);
                var entries = new List<KeyValuePair<string, ObjectId>>();
                foreach (DBDictionaryEntry entry in dictionary)
                {
                    entries.Add(new KeyValuePair<string, ObjectId>(entry.Key, entry.Value));
                }

                foreach (var entry in entries)
                {
                    if (entry.Key.IndexOf('|') >= 0) continue;

                    try
                    {
                        DBObject definition = tr.GetObject(entry.Value, OpenMode.ForWrite);
                        if (definition.IsErased) continue;
                        definition.Erase();
                        removed++;
                    }
                    catch (System.Exception ex)
                    {
                        totals.Failures++;
                        ed.WriteMessage($"\nCould not remove {label} definition '{entry.Key}': {ex.Message}");
                    }
                }
                tr.Commit();
            }
            return removed;
        }

        // Counts the image/PDF definitions still named "xref|name", which means the XREF
        // that owns them was not bound.
        private static int CountXrefDependentDefinitions(Database db)
        {
            int count = 0;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                foreach (string dictionaryName in new[] { ImageDictionaryName, PdfDictionaryName })
                {
                    if (!nod.Contains(dictionaryName)) continue;
                    var dictionary = (DBDictionary)tr.GetObject(nod.GetAt(dictionaryName), OpenMode.ForRead);
                    foreach (DBDictionaryEntry entry in dictionary)
                    {
                        if (entry.Key.IndexOf('|') >= 0) count++;
                    }
                }
                tr.Commit();
            }
            return count;
        }

        // ---------------------------------------------------------------------------------
        // XREF binding
        // ---------------------------------------------------------------------------------

        private struct XrefInfo
        {
            internal ObjectId Id;
            internal string Name;
            internal string Path;
            internal XrefStatus Status;
        }

        // Top-level DWG XREFs only. Nested XREFs ("parent|child") can't be bound on their own;
        // they are handled when their parent is bound.
        private static List<XrefInfo> GetTopLevelDwgXrefs(Database db)
        {
            var result = new List<XrefInfo>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId btrId in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    if (!btr.IsFromExternalReference && !btr.IsFromOverlayReference) continue;
                    if (btr.IsDependent || (btr.Name ?? string.Empty).IndexOf('|') >= 0) continue;

                    result.Add(new XrefInfo
                    {
                        Id = btrId,
                        Name = btr.Name ?? string.Empty,
                        Path = btr.PathName ?? string.Empty,
                        Status = btr.XrefStatus,
                    });
                }
                tr.Commit();
            }
            return result;
        }

        private static bool IsStillExternal(Database db, ObjectId xrefId)
        {
            if (!xrefId.IsValid || xrefId.IsErased) return false;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var btr = tr.GetObject(xrefId, OpenMode.ForRead, true) as BlockTableRecord;
                bool external = btr != null && !btr.IsErased &&
                    (btr.IsFromExternalReference || btr.IsFromOverlayReference);
                tr.Commit();
                return external;
            }
        }

        private static int BindResolvedDwgXrefs(Database db, Editor ed, bool insertBind, List<string> failures)
        {
            int totalBound = 0;
            var reloadAttempted = new HashSet<ObjectId>();

            for (int pass = 1; pass <= MaxBindPasses; pass++)
            {
                List<XrefInfo> xrefs = GetTopLevelDwgXrefs(db);

                // An unloaded XREF whose file can be found only needs a reload to become bindable.
                var toReload = new ObjectIdCollection();
                foreach (XrefInfo xref in xrefs)
                {
                    if (xref.Status == XrefStatus.Unloaded && reloadAttempted.Add(xref.Id))
                    {
                        toReload.Add(xref.Id);
                    }
                }
                if (toReload.Count > 0)
                {
                    ed.WriteMessage($"\nReloading {toReload.Count} unloaded XREF(s)...");
                    try
                    {
                        db.ReloadXrefs(toReload);
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nCould not reload unloaded XREFs: {ex.Message}");
                    }
                    xrefs = GetTopLevelDwgXrefs(db);
                }

                List<XrefInfo> bindable = xrefs.Where(x => x.Status == XrefStatus.Resolved).ToList();
                if (bindable.Count == 0) break;

                ed.WriteMessage(
                    $"\nBinding {bindable.Count} DWG XREF(s) ({(insertBind ? "Insert" : "Bind")})" +
                    (pass > 1 ? $" [pass {pass}: nested XREFs]" : string.Empty) + "...");

                int boundThisPass = BindXrefGroup(db, ed, bindable, insertBind, failures);
                totalBound += boundThisPass;
                if (boundThisPass == 0) break;
            }

            return totalBound;
        }

        private static int BindXrefGroup(Database db, Editor ed, List<XrefInfo> xrefs, bool insertBind, List<string> failures)
        {
            try
            {
                db.BindXrefs(new ObjectIdCollection(xrefs.Select(x => x.Id).ToArray()), insertBind);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nBulk bind failed ({ex.Message}); retrying one XREF at a time...");
                foreach (XrefInfo xref in xrefs)
                {
                    if (!IsStillExternal(db, xref.Id)) continue;
                    try
                    {
                        db.BindXrefs(new ObjectIdCollection(new[] { xref.Id }), insertBind);
                    }
                    catch (System.Exception exSingle)
                    {
                        failures.Add($"{xref.Name}: {exSingle.Message}");
                    }
                }
            }

            return xrefs.Count(x => !IsStillExternal(db, x.Id));
        }

        // ---------------------------------------------------------------------------------
        // Summary
        // ---------------------------------------------------------------------------------

        private static void WriteSummary(
            Database db, Editor ed, MediaTotals totals, int bound, bool insertBind, List<string> bindFailures)
        {
            ed.WriteMessage("\n--- BINDALLXREFS summary ---");
            ed.WriteMessage($"\nData links removed: {totals.DataLinks} (table(s) unlinked: {totals.TablesUnlinked}).");
            ed.WriteMessage($"\nPDF underlays: erased {totals.PdfReferences} reference(s), removed {totals.PdfDefinitions} definition(s).");
            ed.WriteMessage($"\nImages: erased {totals.ImageReferences} reference(s), removed {totals.ImageDefinitions} definition(s).");
            ed.WriteMessage($"\nDWG XREFs bound ({(insertBind ? "Insert" : "Bind")}): {bound}.");

            foreach (string failure in bindFailures)
            {
                ed.WriteMessage($"\n  Bind failed - {failure}");
            }

            // Anything still attached was not bindable: missing file, unreferenced, or failed.
            foreach (XrefInfo xref in GetTopLevelDwgXrefs(db))
            {
                ed.WriteMessage($"\n  Left as XREF - {xref.Name} [{DescribeStatus(xref.Status)}] {xref.Path}");
            }

            int xrefDependentMedia = CountXrefDependentDefinitions(db);
            if (xrefDependentMedia > 0)
            {
                ed.WriteMessage(
                    $"\n  {xrefDependentMedia} image/PDF definition(s) remain inside XREFs that were not bound.");
            }

            if (totals.Failures > 0)
            {
                ed.WriteMessage($"\n{totals.Failures} item(s) could not be removed; see messages above.");
            }
        }

        private static string DescribeStatus(XrefStatus status)
        {
            switch (status)
            {
                case XrefStatus.FileNotFound: return "file not found";
                case XrefStatus.Unresolved: return "unresolved";
                case XrefStatus.Unreferenced: return "not referenced";
                case XrefStatus.Unloaded: return "unloaded";
                case XrefStatus.Resolved: return "bind failed";
                default: return status.ToString();
            }
        }
    }
}
