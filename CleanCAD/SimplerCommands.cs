using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.Collections.Generic;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.IO;

namespace AutoCADCleanupTool
{
    public class SimplerCommands
    {
        // Win32 helpers to auto-dismiss the "OLE Text Size" dialog
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint BM_CLICK = 0x00F5;
        private const uint WM_COMMAND = 0x0111;
        private const int IDOK = 1;

        private static void StartOleTextSizeDialogCloser(double seconds = 120)
        {
            Task.Run(() =>
            {
                var until = DateTime.UtcNow.AddSeconds(seconds);
                while (DateTime.UtcNow < until)
                {
                    try
                    {
                        // Common dialog class is #32770; exact title is "OLE Text Size"
                        IntPtr dlg = FindWindow("#32770", "OLE Text Size");
                        if (dlg != IntPtr.Zero)
                        {
                            // Prefer clicking OK button directly
                            IntPtr ok = FindWindowEx(dlg, IntPtr.Zero, "Button", "OK");
                            if (ok != IntPtr.Zero)
                            {
                                SendMessage(ok, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                            }
                            else
                            {
                                // Fallbacks: try IDOK, then Enter
                                SendMessage(dlg, WM_COMMAND, (IntPtr)IDOK, IntPtr.Zero);
                                SetForegroundWindow(dlg);
                                System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                            }
                            return;
                        }
                    }
                    catch { /* ignore and retry */ }
                    Thread.Sleep(150);
                }
            });
        }
        // The CommandFlags.UsePickSet flag tells AutoCAD that this command
        // is aware of and can use the pre-selected "PickFirst" set.
        [CommandMethod("EraseOther", CommandFlags.UsePickSet)]
        public static void EraseOtherCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptSelectionResult psr;

            // Step 1: Use SelectImplied() to get the pre-selected ("PickFirst") set.
            psr = ed.SelectImplied();

            // If the status is not OK, it means nothing was pre-selected.
            // In this case, we fall back to prompting the user for a new selection.
            if (psr.Status != PromptStatus.OK)
            {
                PromptSelectionOptions pso = new PromptSelectionOptions();
                pso.MessageForAdding = "\nSelect objects to keep: ";
                psr = ed.GetSelection(pso);
            }

            // If the user cancelled or made an empty selection at any point, exit.
            if (psr.Status != PromptStatus.OK)
            {
                return;
            }

            // Create a HashSet for fast lookups of the objects we want to keep.
            var idsToKeep = new HashSet<ObjectId>(psr.Value.GetObjectIds());
            int erasedCount = 0;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockTableRecord currentSpace = trans.GetObject(db.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                    var layersToUnlock = new HashSet<ObjectId>();
                    var idsToErase = new List<ObjectId>();

                    // Iterate through the current space to find objects to erase.
                    foreach (ObjectId id in currentSpace)
                    {
                        if (idsToKeep.Contains(id))
                        {
                            continue; // Skip objects in our "keep" list.
                        }

                        Entity ent = trans.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent != null)
                        {
                            idsToErase.Add(id);
                            LayerTableRecord layer = trans.GetObject(ent.LayerId, OpenMode.ForRead) as LayerTableRecord;
                            if (layer != null && layer.IsLocked)
                            {
                                layersToUnlock.Add(ent.LayerId);
                            }
                        }
                    }

                    // Temporarily unlock any locked layers.
                    foreach (ObjectId layerId in layersToUnlock)
                    {
                        LayerTableRecord layer = trans.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
                        layer.IsLocked = false;
                    }

                    // Erase the objects.
                    foreach (ObjectId idToErase in idsToErase)
                    {
                        Entity entToErase = trans.GetObject(idToErase, OpenMode.ForWrite) as Entity;
                        entToErase.Erase();
                        erasedCount++;
                    }

                    // Re-lock the layers that we unlocked.
                    foreach (ObjectId layerId in layersToUnlock)
                    {
                        LayerTableRecord layer = trans.GetObject(layerId, OpenMode.ForWrite) as LayerTableRecord;
                        layer.IsLocked = true;
                    }

                    trans.Commit();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nAn error occurred: {ex.Message}");
                    trans.Abort();
                }
            }

            if (erasedCount > 0)
            {
                ed.WriteMessage($"\nErased {erasedCount} object(s).");
                ed.Regen();
            }
        }


        // Data and helpers to embed OLE images over existing raster image references
        private class ImagePlacement
        {
            public string Path;
            public Autodesk.AutoCAD.Geometry.Point3d Pos;
            public Autodesk.AutoCAD.Geometry.Vector3d U;
            public Autodesk.AutoCAD.Geometry.Vector3d V;
            public ObjectId ImageId;
        }

        private static readonly Queue<ImagePlacement> _pending = new Queue<ImagePlacement>();
        private static ObjectId _lastPastedOle = ObjectId.Null;
        private static bool _handlersAttached = false;
        private static dynamic _pptAppShared = null;
        private static dynamic _pptPresentationShared = null;
        // Save/restore current layer while embedding
        private static ObjectId _savedClayer = ObjectId.Null;
        // Track candidate image definitions for purge after embedding
        private static readonly HashSet<ObjectId> _imageDefsToPurge = new HashSet<ObjectId>();
        // Chain flag: when true, run FINALIZE after EMBEDFROMXREFS completes
        private static bool _chainFinalizeAfterEmbed = false;

        private static bool EnsurePowerPoint(Editor ed)
        {
            if (_pptAppShared != null && _pptPresentationShared != null)
                return true;
            try
            {
                Type pptType = Type.GetTypeFromProgID("PowerPoint.Application");
                if (pptType == null)
                {
                    ed.WriteMessage("\nPowerPoint is not installed (COM ProgID not found).");
                    return false;
                }
                _pptAppShared = Activator.CreateInstance(pptType);
                try { _pptAppShared.Visible = true; } catch { }
                var presentations = _pptAppShared.Presentations;
                _pptPresentationShared = presentations.Add();
                // Ensure a blank slide exists (ppLayoutBlank = 12)
                _pptPresentationShared.Slides.Add(1, 12);
                return true;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to start PowerPoint: {ex.Message}");
                return false;
            }
        }

        private static bool PrepareClipboardWithImageShared(string path, Editor ed)
        {
            try
            {
                dynamic slide = _pptPresentationShared.Slides[1]; // 1-based
                var shapes = slide.Shapes;
                // Clear previous shapes
                for (int i = shapes.Count; i >= 1; i--)
                {
                    try { shapes[i].Delete(); } catch { }
                }
                dynamic pic = shapes.AddPicture(path, false, true, 10, 10);
                pic.Copy();
                return true;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to prepare clipboard for '{path}': {ex.Message}");
                return false;
            }
        }

        // Resolve a possibly relative image path against the DWG folder and AutoCAD search paths
        private static string ResolveImagePath(Database db, string rawPath)
        {
            if (string.IsNullOrWhiteSpace(rawPath)) return null;

            string p = rawPath.Trim().Trim('"').Replace('/', '\\');

            // If it already exists as-is
            if (Path.IsPathRooted(p) && File.Exists(p))
                return p;

            // Try base DWG folder if available
            try
            {
                string dbFile = db?.Filename;
                if (!string.IsNullOrWhiteSpace(dbFile))
                {
                    string baseDir = Path.GetDirectoryName(dbFile);
                    if (!string.IsNullOrWhiteSpace(baseDir))
                    {
                        // Handle drive-less rooted path like "\folder\file.png" by injecting drive from DWG
                        if (Path.IsPathRooted(p) && p.StartsWith("\\") && !p.StartsWith("\\\\"))
                        {
                            string drive = Path.GetPathRoot(baseDir); // e.g., C:\
                            string combined = Path.Combine(drive ?? string.Empty, p.TrimStart('\\'));
                            string full = Path.GetFullPath(combined);
                            if (File.Exists(full)) return full;
                        }

                        // Plain relative
                        string candidate = Path.GetFullPath(Path.Combine(baseDir, p));
                        if (File.Exists(candidate))
                            return candidate;
                    }
                }
            }
            catch { }

            // Fall back to AutoCAD search paths (support paths, etc.) using only the file name
            try
            {
                string nameOnly = Path.GetFileName(p);
                if (!string.IsNullOrEmpty(nameOnly))
                {
                    string found = HostApplicationServices.Current.FindFile(nameOnly, db, FindFileHint.Default);
                    if (!string.IsNullOrWhiteSpace(found) && File.Exists(found))
                        return found;
                }
            }
            catch { }

            return null;
        }

        private static void AttachHandlers(Database db, Document doc)
        {
            if (_handlersAttached) return;
            db.ObjectAppended += Db_ObjectAppended;
            doc.CommandEnded += Doc_CommandEnded;
            _handlersAttached = true;
        }

        private static void DetachHandlers(Database db, Document doc)
        {
            if (!_handlersAttached) return;
            try { db.ObjectAppended -= Db_ObjectAppended; } catch { }
            try { doc.CommandEnded -= Doc_CommandEnded; } catch { }
            _handlersAttached = false;
        }

        // Remove unused image definitions whose RasterImage references were replaced by OLEs
        private static void PurgeEmbeddedImageDefs(Database db, Editor ed)
        {
            if (_imageDefsToPurge.Count == 0) return;
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var refCounts = new Dictionary<ObjectId, int>();
                    foreach (var defId in _imageDefsToPurge)
                        refCounts[defId] = 0;

                    // Count remaining references to each image def across all block table records
                    foreach (ObjectId btrId in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        foreach (ObjectId entId in btr)
                        {
                            if (entId.ObjectClass.DxfName != "IMAGE") continue;
                            var img = tr.GetObject(entId, OpenMode.ForRead) as RasterImage;
                            if (img == null || img.ImageDefId.IsNull) continue;
                            if (refCounts.ContainsKey(img.ImageDefId))
                                refCounts[img.ImageDefId] = refCounts[img.ImageDefId] + 1;
                        }
                    }

                    // Open the image dictionary to remove entries
                    var named = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                    if (!named.Contains("ACAD_IMAGE_DICT"))
                    {
                        tr.Commit();
                        _imageDefsToPurge.Clear();
                        return;
                    }
                    var imageDict = (DBDictionary)tr.GetObject(named.GetAt("ACAD_IMAGE_DICT"), OpenMode.ForWrite);

                    int purged = 0;
                    foreach (var kvp in refCounts)
                    {
                        if (kvp.Value > 0) continue; // still referenced somewhere

                        // Find the dictionary key for this def id
                        string keyToRemove = null;
                        foreach (DBDictionaryEntry entry in imageDict)
                        {
                            if (entry.Value == kvp.Key)
                            {
                                keyToRemove = entry.Key;
                                break;
                            }
                        }

                        // Remove dict entry and erase the def
                        if (!string.IsNullOrEmpty(keyToRemove))
                        {
                            try { imageDict.Remove(keyToRemove); } catch { }
                        }
                        try
                        {
                            var defObj = tr.GetObject(kvp.Key, OpenMode.ForWrite, false);
                            defObj?.Erase();
                            purged++;
                        }
                        catch { }
                    }

                    if (purged > 0)
                        ed.WriteMessage($"\nPurged {purged} unused image definition(s).");

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to purge image definitions: {ex.Message}");
            }
            finally
            {
                _imageDefsToPurge.Clear();
            }
        }

        private static void Db_ObjectAppended(object sender, ObjectEventArgs e)
        {
            try
            {
                if (e.DBObject is Ole2Frame)
                {
                    _lastPastedOle = e.DBObject.ObjectId;
                }
            }
            catch { }
        }

        private static void Doc_CommandEnded(object sender, Autodesk.AutoCAD.ApplicationServices.CommandEventArgs e)
        {
            try
            {
                if (!string.Equals(e.GlobalCommandName, "PASTECLIP", StringComparison.OrdinalIgnoreCase)) return;

                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;
                var db = doc.Database;
                var ed = doc.Editor;

                if (_lastPastedOle.IsNull || _pending.Count == 0)
                {
                    return;
                }

                var target = _pending.Dequeue();

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var ole = tr.GetObject(_lastPastedOle, OpenMode.ForWrite) as Ole2Frame;
                    if (ole != null)
                    {
                        try
                        {
                            var ext = ole.GeometricExtents;
                            double oleW = Math.Max(1e-8, ext.MaxPoint.X - ext.MinPoint.X);
                            double oleH = Math.Max(1e-8, ext.MaxPoint.Y - ext.MinPoint.Y);

                            var srcOrigin = new Autodesk.AutoCAD.Geometry.Point3d(ext.MinPoint.X, ext.MinPoint.Y, ext.MinPoint.Z);
                            var srcX = new Autodesk.AutoCAD.Geometry.Vector3d(oleW, 0, 0);
                            var srcY = new Autodesk.AutoCAD.Geometry.Vector3d(0, oleH, 0);
                            var srcZ = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis;

                            var destOrigin = target.Pos;
                            var destX = target.U;
                            var destY = target.V;
                            var destZ = destX.CrossProduct(destY);

                            var m = Autodesk.AutoCAD.Geometry.Matrix3d.AlignCoordinateSystem(
                                srcOrigin, srcX, srcY, srcZ,
                                destOrigin, destX, destY, destZ);

                            ole.TransformBy(m);

                            // Ensure pasted OLE is on layer "0"
                            try { ole.Layer = "0"; } catch { }
                        }
                        catch (System.Exception ex1)
                        {
                            ed.WriteMessage($"\nFailed to transform pasted OLE: {ex1.Message}");
                        }
                    }
                    // After placing the OLE where we want it, erase the original raster image entity
                    try
                    {
                        var imgEnt = tr.GetObject(target.ImageId, OpenMode.ForWrite, false) as RasterImage;
                        if (imgEnt != null)
                        {
                            try
                            {
                                if (!imgEnt.ImageDefId.IsNull)
                                    _imageDefsToPurge.Add(imgEnt.ImageDefId);
                            }
                            catch { }
                            // Unlock layer if needed
                            var layer = (LayerTableRecord)tr.GetObject(imgEnt.LayerId, OpenMode.ForRead);
                            bool relock = false;
                            if (layer.IsLocked)
                            {
                                layer.UpgradeOpen();
                                layer.IsLocked = false;
                                relock = true;
                            }
                            imgEnt.Erase();
                            if (relock)
                            {
                                layer.IsLocked = true;
                            }
                        }
                    }
                    catch (System.Exception exErase)
                    {
                        ed.WriteMessage($"\nWarning: failed to erase source raster image: {exErase.Message}");
                    }
                    tr.Commit();
                }

                _lastPastedOle = ObjectId.Null;

                // Continue with next pending image
                if (_pending.Count > 0)
                {
                    ProcessNextPaste(doc, ed);
                }
                else
                {
                    ed.WriteMessage("\nCompleted embedding over all raster images.");
                    // Clean up handlers first, then purge unused image defs
                    DetachHandlers(db, doc);
                    PurgeEmbeddedImageDefs(db, ed);
                    // Restore original current layer if we changed it
                    try { if (!_savedClayer.IsNull && db.Clayer != _savedClayer) db.Clayer = _savedClayer; } catch { }
                    _savedClayer = ObjectId.Null;
                    // If requested, continue the chain by running FINALIZE
                    if (_chainFinalizeAfterEmbed)
                    {
                        _chainFinalizeAfterEmbed = false;
                        doc.SendStringToExecute("_.FINALIZE ", true, false, false);
                    }
                }
            }
            catch { }
        }

        private static void ProcessNextPaste(Document doc, Editor ed)
        {
            if (_pending.Count == 0) return;
            var target = _pending.Peek();

            if (!PrepareClipboardWithImageShared(target.Path, ed))
            {
                _pending.Dequeue();
                if (_pending.Count > 0) ProcessNextPaste(doc, ed);
                return;
            }

            StartOleTextSizeDialogCloser(120);

            string pt = $"{target.Pos.X},{target.Pos.Y}";
            string cmd = $"_.PASTECLIP {pt} 1 0 ";
            doc.SendStringToExecute(cmd, true, false, false);
        }

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

        // Orchestrator: run DETACHSPECIALXREFS, then EMBEDFROMXREFS, then FINALIZE
        [CommandMethod("CLEANCAD", CommandFlags.Modal)]
        public static void RunCleanCad()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            try
            {
                ed.WriteMessage("\n--- CleanCAD: Detaching special refs, embedding, and finalizing ---");
                // Step 1: detach targeted image/DWG xrefs, delete matching blocks, freeze layers
                DetachSpecialXrefs();

                // Step 2: embed images over raster xrefs; when done, automatically run FINALIZE
                _chainFinalizeAfterEmbed = true;
                EmbedFromXrefs();
                // Note: FINALIZE is triggered automatically when embedding completes or immediately if none found.
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nCleanCAD failed: {ex.Message}");
                _chainFinalizeAfterEmbed = false;
            }
        }

        // Detach DWG XREFs and Image references whose names match given substrings
        [CommandMethod("DETACHSPECIALXREFS", CommandFlags.Modal)]
        public static void DetachSpecialXrefs()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            // Case-insensitive tokens
            var imageTokens = new[] { "wl-sig", "christian" };
            var dwgTokens = new[] { "acieslogo", "cdstamp" };

            int dwgDetached = 0;
            int imagesErased = 0;
            int imageDefsDetached = 0;
            int blockRefsErased = 0;
            int blockDefsErased = 0;
            int layersFrozen = 0;

            try
            {
                // 1) Detach DWG XREFs that match tokens
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var toDetach = new List<ObjectId>();

                    foreach (ObjectId btrId in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        if (btr == null || !btr.IsFromExternalReference) continue;

                        string name = (btr.Name ?? string.Empty).ToLowerInvariant();
                        string path = (btr.PathName ?? string.Empty);
                        string fileNoExt = string.Empty;
                        try { fileNoExt = System.IO.Path.GetFileNameWithoutExtension(path)?.ToLowerInvariant() ?? string.Empty; } catch { }

                        bool match = false;
                        foreach (var t in dwgTokens)
                        {
                            var token = t.ToLowerInvariant();
                            if (name.Contains(token) || (!string.IsNullOrEmpty(fileNoExt) && fileNoExt.Contains(token)))
                            {
                                match = true;
                                break;
                            }
                        }

                        if (match)
                        {
                            toDetach.Add(btrId);
                        }
                    }

                    foreach (var xrefId in toDetach)
                    {
                        try { db.DetachXref(xrefId); dwgDetached++; }
                        catch (System.Exception ex) { ed.WriteMessage($"\nFailed to detach DWG XREF {xrefId}: {ex.Message}"); }
                    }

                    tr.Commit();
                }

                // 2) Detach image references/defs that match tokens
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                    if (!nod.Contains("ACAD_IMAGE_DICT"))
                    {
                        tr.Commit();
                    }
                    else
                    {
                        var imageDict = (DBDictionary)tr.GetObject(nod.GetAt("ACAD_IMAGE_DICT"), OpenMode.ForWrite);

                        // Gather candidate defs by key or source file name
                        var candidates = new List<(string Key, ObjectId DefId)>();
                        foreach (DBDictionaryEntry entry in imageDict)
                        {
                            string key = entry.Key ?? string.Empty;
                            var def = tr.GetObject(entry.Value, OpenMode.ForRead) as RasterImageDef;
                            string src = def?.SourceFileName ?? string.Empty;
                            string keyLower = key.ToLowerInvariant();
                            string nameLower = string.Empty;
                            string nameNoExtLower = string.Empty;
                            try { nameLower = System.IO.Path.GetFileName(src)?.ToLowerInvariant() ?? string.Empty; } catch { }
                            try { nameNoExtLower = System.IO.Path.GetFileNameWithoutExtension(src)?.ToLowerInvariant() ?? string.Empty; } catch { }

                            bool match = false;
                            foreach (var t in imageTokens)
                            {
                                var token = t.ToLowerInvariant();
                                if (keyLower.Contains(token) || nameLower.Contains(token) || nameNoExtLower.Contains(token))
                                {
                                    match = true;
                                    break;
                                }
                            }
                            if (match)
                            {
                                candidates.Add((key, entry.Value));
                            }
                        }

                        if (candidates.Count > 0)
                        {
                            // Build an index of raster images by def id
                            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                            var imagesByDef = new Dictionary<ObjectId, List<ObjectId>>();
                            foreach (var c in candidates)
                                imagesByDef[c.DefId] = new List<ObjectId>();

                            foreach (ObjectId btrId in bt)
                            {
                                var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                                foreach (ObjectId entId in btr)
                                {
                                    if (entId.ObjectClass.DxfName != "IMAGE") continue;
                                    var img = tr.GetObject(entId, OpenMode.ForRead) as RasterImage;
                                    if (img == null || img.ImageDefId.IsNull) continue;
                                    if (imagesByDef.ContainsKey(img.ImageDefId))
                                    {
                                        imagesByDef[img.ImageDefId].Add(entId);
                                    }
                                }
                            }

                            // Erase image references and then remove defs
                            foreach (var c in candidates)
                            {
                                // Erase all image entities referencing this def
                                if (imagesByDef.TryGetValue(c.DefId, out var entIds))
                                {
                                    foreach (var entId in entIds)
                                    {
                                        try
                                        {
                                            var img = tr.GetObject(entId, OpenMode.ForWrite, false) as RasterImage;
                                            if (img != null)
                                            {
                                                // Unlock layer if needed
                                                var layer = (LayerTableRecord)tr.GetObject(img.LayerId, OpenMode.ForRead);
                                                bool relock = false;
                                                if (layer.IsLocked)
                                                {
                                                    layer.UpgradeOpen();
                                                    layer.IsLocked = false;
                                                    relock = true;
                                                }
                                                img.Erase();
                                                imagesErased++;
                                                if (relock) layer.IsLocked = true;
                                            }
                                        }
                                        catch (System.Exception ex)
                                        {
                                            ed.WriteMessage($"\nFailed erasing image ref {entId}: {ex.Message}");
                                        }
                                    }
                                }

                                // Remove dictionary key and erase def
                                try { imageDict.Remove(c.Key); } catch { }
                                try
                                {
                                    var def = tr.GetObject(c.DefId, OpenMode.ForWrite, false);
                                    def?.Erase();
                                    imageDefsDetached++;
                                }
                                catch (System.Exception ex)
                                {
                                    ed.WriteMessage($"\nFailed to detach image def '{c.Key}': {ex.Message}");
                                }
                            }
                        }

                        tr.Commit();
                    }
                }

                // 3) Delete block references and block definitions matching patterns
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    // Identify matching block definitions (non-layout, non-xref)
                    var matchedDefs = new HashSet<ObjectId>();
                    foreach (ObjectId btrId in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        if (btr == null) continue;
                        if (btr.IsLayout || btr.IsDependent || btr.IsFromExternalReference) continue;

                        string bname = (btr.Name ?? string.Empty).ToLowerInvariant();
                        bool matchChristian = bname.Contains("christian");
                        bool matchWlSig = bname.Contains("wl") && bname.Contains("sig");
                        if (matchChristian || matchWlSig)
                        {
                            matchedDefs.Add(btrId);
                        }
                    }

                    if (matchedDefs.Count > 0)
                    {
                        // Collect all block references to these definitions (including dynamic base)
                        var refsToErase = new List<ObjectId>();
                        foreach (ObjectId spaceId in bt)
                        {
                            var btr = (BlockTableRecord)tr.GetObject(spaceId, OpenMode.ForRead);
                            foreach (ObjectId entId in btr)
                            {
                                if (entId.ObjectClass.DxfName != "INSERT") continue;
                                var br = tr.GetObject(entId, OpenMode.ForRead) as BlockReference;
                                if (br == null) continue;

                                bool refMatch = matchedDefs.Contains(br.BlockTableRecord);
                                if (!refMatch && br.IsDynamicBlock)
                                {
                                    try { refMatch = matchedDefs.Contains(br.DynamicBlockTableRecord); } catch { }
                                }
                                if (refMatch)
                                {
                                    refsToErase.Add(entId);
                                }
                            }
                        }

                        // Erase all matching references (unlock layer temporarily if needed)
                        foreach (var entId in refsToErase)
                        {
                            try
                            {
                                var br = tr.GetObject(entId, OpenMode.ForWrite, false) as BlockReference;
                                if (br != null)
                                {
                                    var layer = (LayerTableRecord)tr.GetObject(br.LayerId, OpenMode.ForRead);
                                    bool relock = false;
                                    if (layer.IsLocked)
                                    {
                                        layer.UpgradeOpen();
                                        layer.IsLocked = false;
                                        relock = true;
                                    }
                                    br.Erase();
                                    blockRefsErased++;
                                    if (relock) layer.IsLocked = true;
                                }
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nFailed erasing block reference {entId}: {ex.Message}");
                            }
                        }

                        // Try to erase the block definitions now that refs are gone
                        foreach (var defId in matchedDefs)
                        {
                            try
                            {
                                var btr = (BlockTableRecord)tr.GetObject(defId, OpenMode.ForWrite, false);
                                if (btr != null && !btr.IsErased && !btr.IsLayout && !btr.IsDependent && !btr.IsFromExternalReference)
                                {
                                    btr.Erase();
                                    blockDefsErased++;
                                }
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nFailed erasing block definition {defId}: {ex.Message}");
                            }
                        }
                    }

                    tr.Commit();
                }

                // 4) Freeze layers matching requested patterns
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    ObjectId zeroId = ObjectId.Null;
                    LayerTableRecord zeroLtr = null;
                    if (lt.Has("0"))
                    {
                        zeroId = lt["0"];
                        zeroLtr = (LayerTableRecord)tr.GetObject(zeroId, OpenMode.ForWrite);
                        if (zeroLtr.IsOff) zeroLtr.IsOff = false;
                        if (zeroLtr.IsFrozen) zeroLtr.IsFrozen = false;
                    }

                    foreach (ObjectId layerId in lt)
                    {
                        LayerTableRecord ltr = null;
                        try { ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite); }
                        catch { continue; }
                        if (ltr == null) continue;

                        string lname = (ltr.Name ?? string.Empty).ToLowerInvariant();
                        bool matchChristian = lname.Contains("christian");
                        bool matchWlSig = lname.Contains("wl") && lname.Contains("sig");
                        if ((matchChristian || matchWlSig) && !ltr.IsFrozen)
                        {
                            try
                            {
                                if (db.Clayer == layerId)
                                {
                                    if (!zeroId.IsNull && zeroId != layerId)
                                    {
                                        db.Clayer = zeroId;
                                    }
                                    else
                                    {
                                        // If we cannot change, skip freezing current layer
                                        continue;
                                    }
                                }
                                ltr.IsFrozen = true;
                                layersFrozen++;
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nFailed to freeze layer '{ltr.Name}': {ex.Message}");
                            }
                        }
                    }

                    tr.Commit();
                }

                ed.WriteMessage($"\nDetached {dwgDetached} DWG XREF(s). Erased {imagesErased} image ref(s) and detached {imageDefsDetached} image def(s). Erased {blockRefsErased} block ref(s) and deleted {blockDefsErased} block def(s). Froze {layersFrozen} layer(s).");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nDetach operation failed: {ex.Message}");
            }
        }
    }
}
