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
using System.Globalization;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
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
        // Track state for queuing paste insertion points
        private static ImagePlacement _activePlacement = null;
        private static Document _activePasteDocument = null;
        private static bool _waitingForPasteStart = false;
        private static bool _pastePointHandlerAttached = false;
        // State management to prevent re-entrancy and zoom to the final image
        private static bool _isEmbeddingProcessActive = false;
        private static ObjectId _finalPastedOleForZoom = ObjectId.Null;

        private static bool EnsurePowerPoint(Editor ed)
        {
            try
            {
                // Validate existing COM objects if present
                if (_pptAppShared != null)
                {
                    try
                    {
                        var presentations = _pptAppShared.Presentations;
                        try { _pptAppShared.Visible = true; } catch { }

                        if (_pptPresentationShared != null)
                        {
                            try
                            {
                                var slides = _pptPresentationShared.Slides;
                                int count = 0;
                                try { count = slides.Count; } catch { count = 0; }
                                if (count < 1)
                                {
                                    // Ensure a blank slide exists (ppLayoutBlank = 12)
                                    _pptPresentationShared.Slides.Add(1, 12);
                                }
                                return true;
                            }
                            catch
                            {
                                // Presentation reference is stale; create a new presentation
                                _pptPresentationShared = presentations.Add();
                                _pptPresentationShared.Slides.Add(1, 12);
                                return true;
                            }
                        }
                        else
                        {
                            _pptPresentationShared = presentations.Add();
                            _pptPresentationShared.Slides.Add(1, 12);
                            return true;
                        }
                    }
                    catch
                    {
                        // App reference is stale; release and recreate below
                        try { Marshal.FinalReleaseComObject(_pptAppShared); } catch { }
                        _pptAppShared = null;
                        _pptPresentationShared = null;
                    }
                }

                // Create new instance if we get here
                Type pptType = Type.GetTypeFromProgID("PowerPoint.Application");
                if (pptType == null)
                {
                    ed.WriteMessage("\nPowerPoint is not installed (COM ProgID not found).");
                    return false;
                }
                _pptAppShared = Activator.CreateInstance(pptType);
                try { _pptAppShared.Visible = true; } catch { }
                var pres = _pptAppShared.Presentations;
                _pptPresentationShared = pres.Add();
                _pptPresentationShared.Slides.Add(1, 12); // ppLayoutBlank
                return true;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to start or access PowerPoint: {ex.Message}");
                try { if (_pptAppShared != null) Marshal.FinalReleaseComObject(_pptAppShared); } catch { }
                _pptAppShared = null;
                _pptPresentationShared = null;
                return false;
            }
        }

        // Close and release the shared PowerPoint COM objects without saving
        private static void ClosePowerPoint(Editor ed = null)
        {
            try
            {
                if (_pptPresentationShared != null)
                {
                    try { _pptPresentationShared.Saved = -1; } catch { }
                    try { _pptPresentationShared.Close(); } catch { }
                    try { Marshal.FinalReleaseComObject(_pptPresentationShared); } catch { }
                    _pptPresentationShared = null;
                }
                if (_pptAppShared != null)
                {
                    try { _pptAppShared.Quit(); } catch { }
                    try { Marshal.FinalReleaseComObject(_pptAppShared); } catch { }
                    _pptAppShared = null;
                }
            }
            catch (System.Exception ex)
            {
                try { ed?.WriteMessage($"\nWarning: failed to close PowerPoint: {ex.Message}"); } catch { }
                _pptPresentationShared = null;
                _pptAppShared = null;
            }
        }

        private static bool PrepareClipboardWithImageShared(string path, Editor ed)
        {
            try
            {
                if (!EnsurePowerPoint(ed)) return false;
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
            doc.CommandWillStart += Doc_CommandWillStart;
            doc.CommandEnded += Doc_CommandEnded;
            _handlersAttached = true;
        }

        private static void DetachHandlers(Database db, Document doc)
        {
            if (!_handlersAttached) return;
            try { db.ObjectAppended -= Db_ObjectAppended; } catch { }
            try { doc.CommandWillStart -= Doc_CommandWillStart; } catch { }
            try { doc.CommandEnded -= Doc_CommandEnded; } catch { }
            if (_pastePointHandlerAttached)
            {
                try { Autodesk.AutoCAD.ApplicationServices.Application.Idle -= Application_OnIdleSendPastePoint; } catch { }
                _pastePointHandlerAttached = false;
            }
            _activePlacement = null;
            _activePasteDocument = null;
            _waitingForPasteStart = false;
            _handlersAttached = false;
        }

        // Remove unused image definitions whose RasterImage references were replaced by OLEs
        private static void PurgeEmbeddedImageDefs(Database db, Editor ed)
        {
            if (_imageDefsToPurge.Count == 0) return;
            try
            {
                var candidates = new HashSet<ObjectId>();
                foreach (var defId in _imageDefsToPurge)
                {
                    try
                    {
                        if (!defId.IsNull && defId.Database == db)
                        {
                            candidates.Add(defId);
                        }
                    }
                    catch
                    {
                        // Ignore ids from prior drawings or disposed databases
                    }
                }

                if (candidates.Count == 0)
                {
                    return;
                }

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var named = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                    if (!named.Contains("ACAD_IMAGE_DICT"))
                    {
                        tr.Commit();
                        return;
                    }

                    var imageDict = (DBDictionary)tr.GetObject(named.GetAt("ACAD_IMAGE_DICT"), OpenMode.ForWrite);
                    var keysById = new Dictionary<ObjectId, string>();
                    foreach (DBDictionaryEntry entry in imageDict)
                    {
                        if (candidates.Contains(entry.Value))
                        {
                            keysById[entry.Value] = entry.Key;
                        }
                    }

                    if (keysById.Count == 0)
                    {
                        tr.Commit();
                        return;
                    }

                    var imagesByDef = new Dictionary<ObjectId, List<ObjectId>>();
                    foreach (var id in keysById.Keys)
                    {
                        imagesByDef[id] = new List<ObjectId>();
                    }

                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    foreach (ObjectId btrId in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        foreach (ObjectId entId in btr)
                        {
                            if (entId.ObjectClass.DxfName != "IMAGE") continue;
                            var img = tr.GetObject(entId, OpenMode.ForRead) as RasterImage;
                            if (img == null || img.ImageDefId.IsNull) continue;
                            if (imagesByDef.TryGetValue(img.ImageDefId, out var list))
                            {
                                list.Add(entId);
                            }
                        }
                    }

                    int erasedImages = 0;
                    foreach (var kvp in imagesByDef)
                    {
                        foreach (var entId in kvp.Value)
                        {
                            try
                            {
                                var img = tr.GetObject(entId, OpenMode.ForWrite, false) as RasterImage;
                                if (img == null) continue;
                                var layer = (LayerTableRecord)tr.GetObject(img.LayerId, OpenMode.ForRead);
                                bool relock = false;
                                if (layer.IsLocked)
                                {
                                    layer.UpgradeOpen();
                                    layer.IsLocked = false;
                                    relock = true;
                                }
                                img.Erase();
                                if (relock)
                                {
                                    layer.IsLocked = true;
                                }
                                erasedImages++;
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nFailed to erase leftover raster image {entId}: {ex.Message}");
                            }
                        }
                    }

                    int purged = 0;
                    foreach (var defId in keysById.Keys)
                    {
                        if (keysById.TryGetValue(defId, out var key) && !string.IsNullOrEmpty(key))
                        {
                            try
                            {
                                imageDict.Remove(key);
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nFailed to remove image dictionary entry '{key}': {ex.Message}");
                            }
                        }

                        try
                        {
                            var defObj = tr.GetObject(defId, OpenMode.ForWrite, false);
                            if (defObj != null && !defObj.IsErased)
                            {
                                defObj.Erase();
                                purged++;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nFailed to erase image definition {defId}: {ex.Message}");
                        }
                    }

                    if (erasedImages > 0 || purged > 0)
                    {
                        ed.WriteMessage($"\nRemoved {erasedImages} leftover raster image(s) and purged {purged} image definition(s).");
                    }

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


        private static void Doc_CommandWillStart(object sender, Autodesk.AutoCAD.ApplicationServices.CommandEventArgs e)
        {
            try
            {
                if (!string.Equals(e.GlobalCommandName, "PASTECLIP", StringComparison.OrdinalIgnoreCase)) return;
                if (!_waitingForPasteStart || _activePlacement == null) return;
                if (_pastePointHandlerAttached) return;

                if (_activePasteDocument == null)
                {
                    _activePasteDocument = Application.DocumentManager.MdiActiveDocument;
                }

                Autodesk.AutoCAD.ApplicationServices.Application.Idle += Application_OnIdleSendPastePoint;
                _pastePointHandlerAttached = true;
            }
            catch { }
        }

        private static void Application_OnIdleSendPastePoint(object sender, EventArgs e)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Application.Idle -= Application_OnIdleSendPastePoint;
                _pastePointHandlerAttached = false;

                var doc = _activePasteDocument ?? Application.DocumentManager.MdiActiveDocument;
                var placement = _activePlacement;
                if (doc == null || placement == null)
                {
                    return;
                }

                string active = null;
                try { active = doc.CommandInProgress; } catch { }
                if (!string.Equals(active, "PASTECLIP", StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                string x = placement.Pos.X.ToString("G17", CultureInfo.InvariantCulture);
                string y = placement.Pos.Y.ToString("G17", CultureInfo.InvariantCulture);

                doc.SendStringToExecute($"{x},{y}\n1\n0\n", true, false, false);
            }
            catch
            {
            }
            finally
            {
                _waitingForPasteStart = false;
            }
        }

        private static void Doc_CommandEnded(object sender, Autodesk.AutoCAD.ApplicationServices.CommandEventArgs e)
        {
            try
            {
                if (!_isEmbeddingProcessActive || !string.Equals(e.GlobalCommandName, "PASTECLIP", StringComparison.OrdinalIgnoreCase)) return;

                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;
                var db = doc.Database;
                var ed = doc.Editor;

                _waitingForPasteStart = false;
                if (_pastePointHandlerAttached)
                {
                    try { Autodesk.AutoCAD.ApplicationServices.Application.Idle -= Application_OnIdleSendPastePoint; } catch { }
                    _pastePointHandlerAttached = false;
                }

                if (_pending.Count == 0)
                {
                    _activePlacement = null;
                    _activePasteDocument = null;
                    _lastPastedOle = ObjectId.Null;
                    return;
                }

                if (_lastPastedOle.IsNull)
                {
                    _pending.Dequeue();
                    _activePlacement = null;
                    _activePasteDocument = null;
                    ed.WriteMessage("\nSkipping a raster image because no OLE object was created.");

                    if (_pending.Count > 0)
                    {
                        ProcessNextPaste(doc, ed);
                    }
                    else
                    {
                        FinishEmbeddingRun(doc, ed, db);
                    }
                    return;
                }

                var target = _pending.Dequeue();
                _activePlacement = null;
                _activePasteDocument = doc;

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

                if (_pending.Count > 0)
                {
                    ProcessNextPaste(doc, ed);
                }
                else
                {
                    _finalPastedOleForZoom = _lastPastedOle; // This was the last image, save for zoom
                    FinishEmbeddingRun(doc, ed, db);
                }
                _lastPastedOle = ObjectId.Null;
            }
            catch { }
        }

        private static void FinishEmbeddingRun(Document doc, Editor ed, Database db)
        {
            if (!_isEmbeddingProcessActive) return; // Prevent re-entry
            _isEmbeddingProcessActive = false;      // Deactivate process immediately

            
            _finalPastedOleForZoom = ObjectId.Null; // Clean up

            ed.WriteMessage("\nCompleted embedding over all raster images.");
            DetachHandlers(db, doc);
            PurgeEmbeddedImageDefs(db, ed);
            ClosePowerPoint(ed);
            try { if (!_savedClayer.IsNull && db.Clayer != _savedClayer) db.Clayer = _savedClayer; } catch { }
            _savedClayer = ObjectId.Null;
            if (_chainFinalizeAfterEmbed)
            {
                _chainFinalizeAfterEmbed = false;
                doc.SendStringToExecute("_.FINALIZE ", true, false, false);
            }
        }

        private static void ProcessNextPaste(Document doc, Editor ed)
        {
            if (!_isEmbeddingProcessActive || _pending.Count == 0) return;
            var target = _pending.Peek();

            if (!PrepareClipboardWithImageShared(target.Path, ed))
            {
                _pending.Dequeue();
                _activePlacement = null;

                if (_pending.Count > 0)
                {
                    ProcessNextPaste(doc, ed);
                }
                else
                {
                    FinishEmbeddingRun(doc, ed, doc.Database);
                }
                return;
            }

            StartOleTextSizeDialogCloser(120);

            _activePlacement = target;
            _activePasteDocument = doc;
            _waitingForPasteStart = true;

            if (_pastePointHandlerAttached)
            {
                try { Autodesk.AutoCAD.ApplicationServices.Application.Idle -= Application_OnIdleSendPastePoint; } catch { }
                _pastePointHandlerAttached = false;
            }
            doc.SendStringToExecute("_.PASTECLIP\n", true, false, false);
        }

    }
}