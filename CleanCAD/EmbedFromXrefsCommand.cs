using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        [CommandMethod("EMBEDFROMXREFS", CommandFlags.Modal)]
        public static void EmbedFromXrefs()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            _pending.Clear();
            _lastPastedOle = ObjectId.Null;

            try
            {
                // Ensure layer "0" is thawed and make it current for embedding
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (lt.Has("0"))
                    {
                        var zeroId = lt["0"];
                        var zeroLtr = (LayerTableRecord)tr.GetObject(zeroId, OpenMode.ForWrite);
                        if (zeroLtr.IsFrozen) zeroLtr.IsFrozen = false;
                        if (_savedClayer.IsNull) _savedClayer = db.Clayer;
                        db.Clayer = zeroId;
                    }
                    tr.Commit();
                }

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                    foreach (ObjectId id in space)
                    {
                        var img = tr.GetObject(id, OpenMode.ForRead) as RasterImage;
                        if (img == null) continue;
                        if (img.ImageDefId.IsNull) continue;

                        var def = tr.GetObject(img.ImageDefId, OpenMode.ForRead) as RasterImageDef;
                        if (def == null) continue;

                        string resolved = ResolveImagePath(db, def.SourceFileName);
                        if (string.IsNullOrWhiteSpace(resolved))
                        {
                            ed.WriteMessage($"\nSkipping missing image: {def.SourceFileName}");
                            continue;
                        }

                        var cs = img.Orientation;
                        var placement = new ImagePlacement
                        {
                            Path = resolved,
                            Pos = cs.Origin,
                            U = cs.Xaxis,
                            V = cs.Yaxis,
                            ImageId = img.ObjectId
                        };
                        _pending.Enqueue(placement);
                    }
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to collect raster images: {ex.Message}");
                // Restore original current layer if we changed it
                try { if (!_savedClayer.IsNull && db.Clayer != _savedClayer) db.Clayer = _savedClayer; } catch { }
                _savedClayer = ObjectId.Null;
                return;
            }

            if (_pending.Count == 0)
            {
                ed.WriteMessage("\nNo raster images found in current space.");
                // Close PowerPoint if it was open from a prior run (no save)
                ClosePowerPoint(ed);
                // Restore original current layer if we changed it
                try { if (!_savedClayer.IsNull && db.Clayer != _savedClayer) db.Clayer = _savedClayer; } catch { }
                _savedClayer = ObjectId.Null;
                // If running as part of CLEANCAD chain, move on to FINALIZE immediately
                if (_chainFinalizeAfterEmbed)
                {
                    _chainFinalizeAfterEmbed = false;
                    doc.SendStringToExecute("_.FINALIZE ", true, false, false);
                }
                return;
            }

            if (!EnsurePowerPoint(ed))
            {
                // Restore original current layer if we changed it
                try { if (!_savedClayer.IsNull && db.Clayer != _savedClayer) db.Clayer = _savedClayer; } catch { }
                _savedClayer = ObjectId.Null;
                return;
            }
            AttachHandlers(db, doc);
            ed.WriteMessage($"\nEmbedding over {_pending.Count} raster image(s)...");
            ProcessNextPaste(doc, ed);
        }
    }
}

