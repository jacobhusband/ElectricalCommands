using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace AutoCADCleanupTool
{
    // One JSON request/result per isolated Core Console process. Never saves a source DWG.
    public partial class SimplerCommands
    {
        internal sealed class CleanJob
        {
            public string Operation { get; set; }
            public string ResultPath { get; set; }
            public string[] Files { get; set; }
            public string[] Catalog { get; set; }
            public Dictionary<string, string> Copies { get; set; }
            public List<CleanReference> References { get; set; }
            public List<CleanPdfReference> Pdfs { get; set; }
            public Dictionary<string, string> PdfCopies { get; set; }
            public List<CleanPdfReference> Images { get; set; }
            public Dictionary<string, string> ImageCopies { get; set; }
            public string Titleblock { get; set; }
            public string Output { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
            public string[] Layouts { get; set; }
        }

        internal sealed class CleanReference
        {
            public string Owner;
            public string Name;
            public string Path;
            public string Resolved;
        }

        internal sealed class CleanPdfReference
        {
            public string Owner, Handle, Path, Resolved, Page, Layout;
            public bool Clipped;
            public int Fade, Contrast;
        }

        [CommandMethod("ACIESCLEANJOB", CommandFlags.Modal)]
        public static void RunCleanDrawingJob()
        {
            CleanJob job = null;
            try
            {
                job = JsonConvert.DeserializeObject<CleanJob>(File.ReadAllText(
                    Environment.GetEnvironmentVariable("ACIES_CLEAN_JOB")));
                if (job == null || string.IsNullOrWhiteSpace(job.ResultPath))
                    throw new InvalidOperationException("Missing job/result path.");
                object details;
                if (job.Operation == "scan") details = ScanCleanInputs(job);
                else if (job.Operation == "plot-review") details = PlotCleanReview(job);
                else if (job.Operation == "remove-titleblock-marks")
                {
                    var db = Application.DocumentManager.MdiActiveDocument.Database;
                    if (string.IsNullOrWhiteSpace(job.Output) || File.Exists(job.Output))
                        throw new InvalidOperationException("Mark removal requires a new output DWG.");
                    details = RemoveTitleblockMarks(db);
                    db.SaveAs(job.Output, DwgVersion.Current);
                }
                else if (job.Operation == "validate-titleblock")
                    details = ValidateTitleblockBoundary(Application.DocumentManager.MdiActiveDocument.Database, job);
                else if (job.Operation == "prepare") { PrepareCleanCopies(job); details = new { }; }
                else if (job.Operation == "embed-media" || job.Operation == "verify-media")
                {
                    var db = Application.DocumentManager.MdiActiveDocument.Database;
                    if (job.Operation == "embed-media")
                    {
                        if (string.IsNullOrWhiteSpace(job.Output) || File.Exists(job.Output))
                            throw new InvalidOperationException("Media output must be a new DWG path.");
                        EmbedHeadlessMedia(db, Path.GetDirectoryName(job.ResultPath));
                    }
                    details = VerifyLocalMedia(db);
                    if (job.Operation == "embed-media") db.SaveAs(job.Output, DwgVersion.Current);
                }
                else
                {
                    var db = Application.DocumentManager.MdiActiveDocument.Database;
                    object embedding = null;
                    if (job.Operation == "embed-prototype")
                    {
                        if (string.IsNullOrWhiteSpace(job.Output) || File.Exists(job.Output))
                            throw new InvalidOperationException("Prototype output must be a new DWG path.");
                        embedding = EmbedNativeImageGeometry(db);
                    }
                    else if (job.Operation == "titleblock") CleanTitleblockBatch(db, job);
                    else if (job.Operation == "sheet")
                    {
                        TitleBlockXrefResolver.ConfirmedBatchPath = job.Titleblock;
                        TitleBlockXrefResolver.ConfirmedBatchBoundary = new[] {
                            Point3d.Origin, new Point3d(job.Width, 0, 0),
                            new Point3d(job.Width, job.Height, 0), new Point3d(0, job.Height, 0) };
                        RunCleanSheetHeadless();
                        if (!HeadlessSheetSucceeded) throw new InvalidOperationException("CLEANCAD2 did not complete. See worker log.");
                    }
                    else if (job.Operation != "verify") throw new InvalidOperationException("Unknown operation.");
                    var validation = VerifyCleanDatabase(db, job.Operation == "titleblock");
                    details = embedding == null ? validation : new { embedding, validation };
                    if (job.Operation != "verify")
                    {
                        if (string.Equals(Path.GetFullPath(job.Output), Path.GetFullPath(db.Filename), StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("Output must differ from worker input.");
                        db.SaveAs(job.Output, DwgVersion.Current);
                    }
                }
                File.WriteAllText(job.ResultPath, JsonConvert.SerializeObject(new { success = true, details }, Formatting.Indented));
            }
            catch (System.Exception ex)
            {
                if (job != null && !string.IsNullOrWhiteSpace(job.ResultPath))
                    File.WriteAllText(job.ResultPath, JsonConvert.SerializeObject(new { success = false, error = ex.ToString() }));
                Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage("\nACIESCLEAN_ERROR: " + ex.Message);
            }
            finally
            {
                TitleBlockXrefResolver.ConfirmedBatchPath = null;
                TitleBlockXrefResolver.ConfirmedBatchBoundary = null;
            }
        }

        private static object ValidateTitleblockBoundary(Database db, CleanJob job)
        {
            if (!(job.Width > 0 && job.Width <= 1000 && job.Height > 0 && job.Height <= 1000))
                throw new InvalidOperationException("Enter finite positive titleblock dimensions.");
            int inside = 0;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var active = new HashSet<ObjectId>();
                Action<ObjectId, Matrix3d> inspect = null;
                inspect = (blockId, transform) =>
                {
                    var block = (BlockTableRecord)tr.GetObject(blockId, OpenMode.ForRead);
                    if (block.IsFromExternalReference || block.IsDependent) return;
                    if (!active.Add(blockId)) throw new InvalidOperationException("Cyclic titleblock geometry.");
                    foreach (ObjectId id in block)
                    {
                        if (!(tr.GetObject(id, OpenMode.ForRead) is Entity entity)) continue;
                        if (entity is BlockReference reference)
                        { inspect(reference.BlockTableRecord, transform * reference.BlockTransform); continue; }
                        try
                        {
                            var bounds = entity.GeometricExtents;
                            bounds.TransformBy(transform);
                            if (bounds.MaxPoint.X >= 0 && bounds.MaxPoint.Y >= 0 && bounds.MinPoint.X <= job.Width && bounds.MinPoint.Y <= job.Height)
                                inside++;
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception) { /* Non-geometric metadata is not evidence of a border. */ }
                    }
                    active.Remove(blockId);
                };
                inspect(SymbolUtilityServices.GetBlockModelSpaceId(db), Matrix3d.Identity);
                tr.Commit();
            }
            if (inside == 0)
                throw new InvalidOperationException($"The selected XREF has no local titleblock geometry inside (0,0) to ({job.Width},{job.Height}). Select the border/titleblock DWG rather than a site or plan background, and confirm the sheet dimensions. No cleanup was started.");
            return new { localEntitiesInsideBoundary = inside };
        }

        private static object ScanCleanInputs(CleanJob job)
        {
            var pending = new Queue<string>(job.Files);
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var references = new List<CleanReference>();
            var media = new List<string>();
            var pdfs = new List<CleanPdfReference>();
            var images = new List<CleanPdfReference>();
            var paperReferences = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            while (pending.Count > 0)
            {
                string file = Path.GetFullPath(pending.Dequeue());
                if (!visited.Add(file)) continue;
                using (var db = new Database(false, true))
                {
                    db.ReadDwgFile(file, FileOpenMode.OpenForReadAndAllShare, false, null);
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        foreach (ObjectId id in bt)
                        {
                            var block = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                            if (block.IsFromExternalReference || block.IsFromOverlayReference)
                            {
                                if (block.IsDependent) continue;
                                string saved = block.PathName;
                                string resolved = Path.GetFullPath(Path.IsPathRooted(saved) ? saved : Path.Combine(Path.GetDirectoryName(file), saved));
                                if (!File.Exists(resolved))
                                {
                                    var matches = (job.Catalog ?? Array.Empty<string>()).Where(p =>
                                        string.Equals(Path.GetFileName(p), Path.GetFileName(saved), StringComparison.OrdinalIgnoreCase)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                                    if (matches.Count != 1) throw new IOException($"Missing or ambiguous XREF '{saved}' in '{file}'. Repair its path first.");
                                    resolved = Path.GetFullPath(matches[0]);
                                }
                                references.Add(new CleanReference { Owner = file, Name = block.Name, Path = saved, Resolved = resolved });
                                pending.Enqueue(resolved);
                                continue;
                            }
                            foreach (ObjectId entId in block)
                            {
                                var entity = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                                if (block.IsLayout && block.Name != BlockTableRecord.ModelSpace && entity is BlockReference br)
                                {
                                    var definition = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                                    if (definition.IsFromExternalReference || definition.IsFromOverlayReference)
                                        paperReferences.Add(file + "\n" + definition.Name);
                                }
                                if (entity is PdfReference pdf)
                                {
                                    var definition = (PdfDefinition)tr.GetObject(pdf.DefinitionId, OpenMode.ForRead);
                                    string path = definition.SourceFileName;
                                    pdfs.Add(new CleanPdfReference { Owner = file, Handle = entId.Handle.ToString(), Path = path,
                                        Resolved = Path.GetFullPath(Path.IsPathRooted(path) ? path : Path.Combine(Path.GetDirectoryName(file), path)),
                                        Page = definition.ItemName, Layout = block.IsLayout ? ((Layout)tr.GetObject(block.LayoutId, OpenMode.ForRead)).LayoutName : null,
                                        Clipped = pdf.IsClipped, Fade = pdf.Fade, Contrast = pdf.Contrast });
                                }
                                else if (IsExternalRasterImage(entity))
                                {
                                    var image = (RasterImage)entity;
                                    string path = ((RasterImageDef)tr.GetObject(image.ImageDefId, OpenMode.ForRead)).SourceFileName;
                                    images.Add(new CleanPdfReference { Owner = file, Handle = entId.Handle.ToString(), Path = path,
                                        Resolved = Path.GetFullPath(Path.IsPathRooted(path) ? path : Path.Combine(Path.GetDirectoryName(file), path)) });
                                }
                                else if (entity is UnderlayReference || (entity is Ole2Frame ole && ole.IsLinked))
                                    media.Add(file + ": " + entity.GetType().Name + " " + entId.Handle);
                            }
                        }
                        tr.Commit();
                    }
                }
            }
            return new { files = visited.ToArray(), references, media, pdfs, images,
                titleblocks = references.Where(r => paperReferences.Contains(r.Owner + "\n" + r.Name))
                    .Select(r => r.Resolved).Distinct(StringComparer.OrdinalIgnoreCase).ToArray() };
        }

        private static void PrepareCleanCopies(CleanJob job)
        {
            var copies = new Dictionary<string, string>(job.Copies, StringComparer.OrdinalIgnoreCase);
            foreach (var pair in copies)
            {
                using (var db = new Database(false, true))
                {
                    db.ReadDwgFile(pair.Value, FileOpenMode.OpenForReadAndAllShare, false, null);
                    db.CloseInput(true);
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var expected = job.References.Where(r => string.Equals(r.Owner, pair.Key, StringComparison.OrdinalIgnoreCase))
                            .ToDictionary(r => r.Name, StringComparer.OrdinalIgnoreCase);
                        int found = 0;
                        foreach (ObjectId id in bt)
                        {
                            var block = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                            if ((!block.IsFromExternalReference && !block.IsFromOverlayReference) || block.IsDependent) continue;
                            if (!expected.TryGetValue(block.Name, out var reference) ||
                                !string.Equals(block.PathName, reference.Path, StringComparison.OrdinalIgnoreCase))
                                throw new IOException("Reference graph changed since inspection: " + pair.Key);
                            found++;
                            block.UpgradeOpen();
                            block.PathName = copies[reference.Resolved];
                        }
                        if (found != expected.Count) throw new IOException("Reference graph changed since inspection: " + pair.Key);
                        RedirectCleanMedia(db, tr, job, pair.Key);
                        tr.Commit();
                    }
                    db.SaveAs(pair.Value, DwgVersion.Current);
                }
            }
        }

        [CommandMethod("CLEANTBLK2", CommandFlags.Modal)]
        public static void RunCleanTitleBlockHeadless()
        {
            // Batch supplies the confirmed dimensions; no interactive fallback.
            RunCleanDrawingJob();
        }

        private static void CleanTitleblockBatch(Database db, CleanJob job)
        {
            if (double.IsNaN(job.Width) || double.IsInfinity(job.Width) || job.Width <= 0 ||
                double.IsNaN(job.Height) || double.IsInfinity(job.Height) || job.Height <= 0)
                throw new InvalidOperationException("A finite positive titleblock size is required.");
            var preflight = InspectForHeadlessBlockers(db);
            if (!preflight.Succeeded || preflight.RasterImageCount > 0 || preflight.UnderlayCount > 0 || preflight.LinkedOleCount > 0)
                throw new InvalidOperationException("Unsupported media or unreadable titleblock. No cleanup performed.");
            ResetHeadlessWorkflowState();
            EnsureAllLayersVisibleAndUnlocked(db, Application.DocumentManager.MdiActiveDocument.Editor);
            ExplodeAllBlockReferences();
            CleanupCommands.PruneFixedTitleblock(db, job.Width, job.Height);
            // CLEANTBLK intentionally removes DWG references after retaining local border geometry.
            ObjectId[] xrefsToDetach;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                xrefsToDetach = bt.Cast<ObjectId>().Where(id => {
                    var b = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    return (b.IsFromExternalReference || b.IsFromOverlayReference) && !b.IsDependent;
                }).ToArray();
                tr.Commit();
            }
            foreach (var id in xrefsToDetach) db.DetachXref(id);
        }

        private static object VerifyCleanDatabase(Database db, bool titleblock)
        {
            var check = InspectForHeadlessBlockers(db);
            if (!check.Succeeded || check.RasterImageCount > 0 || check.UnderlayCount > 0 || check.LinkedOleCount > 0 || FindRemainingDwgXrefs(db).Count > 0)
                throw new InvalidOperationException("Output contains external references/media or could not be inspected.");
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                int modelEntities = ((BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead)).Cast<ObjectId>().Count(id => !id.IsErased);
                int paperEntities = 0;
                int wipeouts = 0;
                foreach (ObjectId id in bt)
                {
                    var block = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    foreach (ObjectId entityId in block)
                        if (!entityId.IsErased && tr.GetObject(entityId, OpenMode.ForRead) is Wipeout) wipeouts++;
                    if (block.IsLayout && id != bt[BlockTableRecord.ModelSpace])
                        paperEntities += block.Cast<ObjectId>().Count(e => !e.IsErased && e.ObjectClass.DxfName != "VIEWPORT");
                }
                if (titleblock ? modelEntities == 0 : modelEntities + paperEntities == 0)
                    throw new InvalidOperationException("Cleanup produced an empty drawing.");
                tr.Commit();
                return new { modelEntities, paperEntities, externalReferences = 0, wipeouts };
            }
        }
    }

    public partial class CleanupCommands
    {
        internal static void PruneFixedTitleblock(Database db, double width, double height)
        {
            var keep = new HashSet<ObjectId>();
            var polygon = new Point3dCollection(new[] { Point3d.Origin, new Point3d(width, 0, 0), new Point3d(width, height, 0), new Point3d(0, height, 0) });
            ObjectId spaceId = SymbolUtilityServices.GetBlockModelSpaceId(db);
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var space = (BlockTableRecord)tr.GetObject(spaceId, OpenMode.ForRead);
                foreach (ObjectId id in space)
                {
                    var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (entity == null) continue;
                    if (entity is BlockReference br)
                    {
                        var definition = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                        if (definition.IsFromExternalReference || definition.IsFromOverlayReference) continue;
                        throw new InvalidOperationException("A local titleblock block could not be exploded: " + id.Handle);
                    }
                    var bounds = TryGetExtents(entity);
                    if (!bounds.HasValue) throw new InvalidOperationException("Cannot measure titleblock entity " + id.Handle);
                    // Conservative crossing semantics: never discard an entity whose bounds touch the sheet.
                    if (ExtentsIntersectsPolygonXY(bounds.Value, polygon)) keep.Add(id);
                }
                tr.Commit();
            }
            if (keep.Count == 0) throw new InvalidOperationException("No local titleblock geometry inside the confirmed boundary.");
            EraseEntitiesExcept(db, spaceId, keep);
        }
    }
}
