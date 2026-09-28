using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        private static object VerifyLocalMedia(Database db)
        {
            int entities = 0, solids = 0;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead))
                {
                    var block = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    if (block.IsFromExternalReference || block.IsDependent) continue;
                    foreach (ObjectId entityId in block)
                    {
                        var entity = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                        if (IsExternalRasterImage(entity) || entity is UnderlayReference || (entity is Ole2Frame ole && ole.IsLinked))
                            throw new InvalidOperationException("External media remains after conversion: " + entityId.Handle);
                        entities++;
                        if (entity is Solid) solids++;
                    }
                }
                tr.Commit();
            }
            return new { entities, solids, externalMedia = 0 };
        }

        private static void EmbedHeadlessMedia(Database db, string workspace)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var ids = new List<ObjectId>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead))
                {
                    var block = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    if (block.IsFromExternalReference || block.IsDependent) continue;
                    foreach (ObjectId entityId in block)
                    {
                        if (!(tr.GetObject(entityId, OpenMode.ForRead) is PdfReference pdf)) continue;
                        // A fresh attachment has clipping enabled but no explicit boundary.
                        if (!block.IsLayout || (pdf.IsClipped && pdf.GetClipBoundary().Length > 0) || pdf.Fade != 0 || pdf.Contrast != 100 || pdf.Monochrome)
                            throw new InvalidOperationException("PDF conversion requires layout-level, unadjusted underlays without a custom clip. Underlay " + entityId.Handle);
                        var definition = (PdfDefinition)tr.GetObject(pdf.DefinitionId, OpenMode.ForRead);
                        if (!File.Exists(definition.SourceFileName)) throw new FileNotFoundException("Staged PDF is missing.", definition.SourceFileName);
                        ids.Add(entityId);
                    }
                }
                tr.Commit();
            }
            string layout = LayoutManager.Current.CurrentLayout;
            var ucs = ed.CurrentUserCoordinateSystem;
            var settings = new Dictionary<string, object>();
            try
            {
                foreach (string name in new[] { "FILEDIA", "PDFIMPORTFILTER", "PDFIMPORTMODE", "PDFIMPORTLAYERS", "PDFIMPORTIMAGEPATH" })
                    settings[name] = Application.GetSystemVariable(name);
                Application.SetSystemVariable("FILEDIA", 0);
                Application.SetSystemVariable("PDFIMPORTFILTER", 0); // Never silently omit image content.
                Application.SetSystemVariable("PDFIMPORTMODE", 7); // One block, lineweights and joined paths; opaque solids.
                Application.SetSystemVariable("PDFIMPORTLAYERS", 2);
                Application.SetSystemVariable("PDFIMPORTIMAGEPATH", workspace);
                ed.CurrentUserCoordinateSystem = Matrix3d.Identity;
                EnsureAllLayersVisibleAndUnlocked(db, ed);
                foreach (ObjectId id in ids)
                {
                    ObjectId ownerId;
                    HashSet<ObjectId> before;
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var pdf = (PdfReference)tr.GetObject(id, OpenMode.ForRead);
                        ownerId = pdf.OwnerId;
                        var owner = (BlockTableRecord)tr.GetObject(ownerId, OpenMode.ForRead);
                        LayoutManager.Current.CurrentLayout = ((Layout)tr.GetObject(owner.LayoutId, OpenMode.ForRead)).LayoutName;
                        before = new HashSet<ObjectId>(owner.Cast<ObjectId>());
                        db.Clayer = pdf.LayerId;
                        tr.Commit();
                    }
                    ed.Command("_.-PDFIMPORT", id, "_All", "_Keep");
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var owner = (BlockTableRecord)tr.GetObject(ownerId, OpenMode.ForRead);
                        var imported = owner.Cast<ObjectId>().Where(x => !before.Contains(x) && !x.IsErased).ToArray();
                        if (imported.Length != 1 || !(tr.GetObject(imported[0], OpenMode.ForWrite) is BlockReference replacement))
                            throw new InvalidOperationException("PDF import did not produce one native content block for " + id.Handle);
                        var content = (BlockTableRecord)tr.GetObject(replacement.BlockTableRecord, OpenMode.ForRead);
                        if (!content.Cast<ObjectId>().Any()) throw new InvalidOperationException("PDF import returned empty content.");
                        var original = (PdfReference)tr.GetObject(id, OpenMode.ForWrite);
                        replacement.SetPropertiesFrom(original);
                        replacement.Visible = original.Visible && original.IsOn;
                        var order = (DrawOrderTable)tr.GetObject(owner.DrawOrderTableId, OpenMode.ForWrite);
                        order.MoveAbove(new ObjectIdCollection(imported), id);
                        original.Erase();
                        tr.Commit();
                    }
                    ed.WriteMessage("\nConverted PDF underlay " + id.Handle);
                }
                EmbedNativeImageGeometry(db);
                VerifyLocalMedia(db);
            }
            finally
            {
                foreach (var setting in settings) Application.SetSystemVariable(setting.Key, setting.Value);
                LayoutManager.Current.CurrentLayout = layout;
                ed.CurrentUserCoordinateSystem = ucs;
            }
        }

        private static void RedirectCleanMedia(Database db, Transaction tr, CleanJob job, string originalOwner)
        {
            foreach (bool pdf in new[] { true, false })
            {
                var records = pdf ? job.Pdfs : job.Images;
                var copies = pdf ? job.PdfCopies : job.ImageCopies;
                if (records == null) continue;
                foreach (var record in records.Where(x => string.Equals(x.Owner, originalOwner, StringComparison.OrdinalIgnoreCase)))
                {
                    var id = db.GetObjectId(false, new Handle(Convert.ToInt64(record.Handle, 16)), 0);
                    if (pdf)
                    {
                        var entity = (PdfReference)tr.GetObject(id, OpenMode.ForRead);
                        var definition = (PdfDefinition)tr.GetObject(entity.DefinitionId, OpenMode.ForWrite);
                        string target = copies[record.Resolved];
                        if (definition.SourceFileName != record.Path && definition.SourceFileName != target)
                            throw new IOException("PDF path changed since inspection.");
                        if (definition.ItemName != record.Page) throw new IOException("PDF page changed since inspection.");
                        definition.SourceFileName = target;
                    }
                    else
                    {
                        var entity = (RasterImage)tr.GetObject(id, OpenMode.ForRead);
                        var definition = (RasterImageDef)tr.GetObject(entity.ImageDefId, OpenMode.ForWrite);
                        string target = copies[record.Resolved];
                        if (definition.SourceFileName != record.Path && definition.SourceFileName != target)
                            throw new IOException("Image path changed since inspection.");
                        definition.SourceFileName = target;
                    }
                }
            }
        }
    }
}
