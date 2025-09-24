using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry; // This using directive is required for Vector3d
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

        /// <summary>
        /// Unlocks, thaws, and turns on all layers in the drawing to ensure visibility and editability.
        /// </summary>
        internal static void EnsureAllLayersVisibleAndUnlocked(Database db, Editor ed)
        {
            ed.WriteMessage("\nEnsuring all layers are visible and unlocked...");
            int unlocked = 0;
            int thawed = 0;
            int turnedOn = 0;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    ObjectId originalClayerId = db.Clayer;

                    // Temporarily switch to a safe layer (like "0") to avoid issues with modifying the current layer.
                    if (lt.Has("0") && db.Clayer != lt["0"])
                    {
                        db.Clayer = lt["0"];
                    }

                    foreach (ObjectId layerId in lt)
                    {
                        if (layerId.IsErased) continue;
                        try
                        {
                            var ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);

                            if (ltr.IsLocked)
                            {
                                ltr.IsLocked = false;
                                unlocked++;
                            }
                            if (ltr.IsFrozen)
                            {
                                // This is safe because we are not on a frozen layer.
                                ltr.IsFrozen = false;
                                thawed++;
                            }
                            if (ltr.IsOff)
                            {
                                ltr.IsOff = false;
                                turnedOn++;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nCould not modify layer {layerId}: {ex.Message}");
                        }
                    }

                    // Attempt to restore the original current layer.
                    try
                    {
                        var originalLtr = (LayerTableRecord)tr.GetObject(originalClayerId, OpenMode.ForRead);
                        if (!originalLtr.IsErased && !originalLtr.IsFrozen)
                        {
                            db.Clayer = originalClayerId;
                        }
                    }
                    catch { }

                    tr.Commit();
                }
                ed.WriteMessage($"\nUnlocked: {unlocked}, Thawed: {thawed}, Turned On: {turnedOn} layer(s).");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred while modifying layers: {ex.Message}");
            }
        }

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
            public Point3d Pos;
            public Vector3d U;
            public Vector3d V;
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
        // Store original PowerPoint image dimensions for scaling correction
        private static readonly Dictionary<ObjectId, Point2d> _originalPptImageDims = new Dictionary<ObjectId, Point2d>();
        // Track XREFs that contained images that were embedded, so they can be detached
        private static readonly HashSet<ObjectId> _xrefsToDetach = new HashSet<ObjectId>();

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

        private static bool PrepareClipboardWithImageShared(ImagePlacement placement, Editor ed)
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
                string path = placement.Path;
                dynamic pic = shapes.AddPicture(path, false, true, 10, 10);

                // Store original dimensions before rotation
                _originalPptImageDims[placement.ImageId] = new Point2d(pic.Width, pic.Height);

                // START of new rotation logic
                float picW = pic.Width;
                float picH = pic.Height;
                double destWidth = placement.U.Length;
                double destHeight = placement.V.Length;

                bool picIsLandscape = picW > picH;
                bool destIsLandscape = destWidth > destHeight;

                double rotationAngleRad = Math.Atan2(placement.U.Y, placement.U.X);

                if (picIsLandscape != destIsLandscape)
                {
                    rotationAngleRad -= Math.PI / 2.0; // Adjust by -90 degrees
                }

                // PowerPoint's Rotation is in degrees and clockwise.
                float rotationAngleDeg = (float)(-rotationAngleRad * 180.0 / Math.PI);
                pic.Rotation = rotationAngleDeg;
                // END of new rotation logic

                pic.Copy();
                return true;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to prepare clipboard for '{placement.Path}': {ex.Message}");
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

                if (_pending.Count == 0 && _activePlacement == null)
                {
                    return;
                }

                if (_lastPastedOle.IsNull)
                {
                    ed.WriteMessage("\nSkipping a raster image because no OLE object was created.");
                    if (_pending.Count > 0) _pending.Dequeue();
                    _activePlacement = null;
                    _activePasteDocument = null;

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

                var target = _pending.Count > 0 ? _pending.Dequeue() : _activePlacement;
                if (target == null)
                {
                    ed.WriteMessage("\nError: Could not retrieve active image placement. Aborting.");
                    FinishEmbeddingRun(doc, ed, db);
                    return;
                }

                _activePlacement = null;
                _activePasteDocument = doc;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var ole = tr.GetObject(_lastPastedOle, OpenMode.ForWrite) as Ole2Frame;
                    if (ole != null)
                    {
                        try
                        {
                            // *** MODIFICATION START: New transformation logic with rotation in PowerPoint ***
                        var ext = ole.GeometricExtents;
                        double oleW_bbox = Math.Max(1e-8, ext.MaxPoint.X - ext.MinPoint.X);
                        double oleH_bbox = Math.Max(1e-8, ext.MaxPoint.Y - ext.MinPoint.Y);
                        Point3d oleCenter = ext.MinPoint + (ext.MaxPoint - ext.MinPoint) * 0.5;

                        var destU = target.U;
                        var destV = target.V;
                        double destWidth = destU.Length;
                        double destHeight = destV.Length;

                        if (destWidth < 1e-8 || destHeight < 1e-8)
                        {
                            ed.WriteMessage($"\nSkipping degenerate raster image {target.ImageId} with zero area.");
                            ole.Erase();
                            return;
                        }

                        // Retrieve original PowerPoint image dimensions
                        Point2d originalPptDims = Point2d.Origin;
                        if (_originalPptImageDims.ContainsKey(target.ImageId))
                        {
                            originalPptDims = _originalPptImageDims[target.ImageId];
                            _originalPptImageDims.Remove(target.ImageId); // Clean up
                        }
                        else
                        {
                            ed.WriteMessage($"\nWarning: Original PowerPoint image dimensions not found for {target.ImageId}. Scaling might be incorrect.");
                            // Fallback: assume oleW_bbox and oleH_bbox are the effective dimensions
                            originalPptDims = new Point2d(oleW_bbox, oleH_bbox);
                        }

                        double imgW_orig = originalPptDims.X;
                        double imgH_orig = originalPptDims.Y;

                        // Recalculate the rotation angle that was applied in PowerPoint
                        double rotationAngleRad_ppt = Math.Atan2(destU.Y, destU.X);
                        bool picIsLandscape = imgW_orig > imgH_orig;
                        bool destIsLandscape = destWidth > destHeight;

                        if (picIsLandscape != destIsLandscape)
                        {
                            rotationAngleRad_ppt -= Math.PI / 2.0; // Adjust by -90 degrees
                        }

                        // Calculate the unrotated dimensions of the OLE content from its bounding box
                        // This is the inverse of the bounding box calculation
                        double c = Math.Abs(Math.Cos(rotationAngleRad_ppt));
                        double s = Math.Abs(Math.Sin(rotationAngleRad_ppt));

                        double effectiveOleWidth, effectiveOleHeight;

                        // Handle cases where c^2 - s^2 is close to zero (e.g., rotation near 45 degrees)
                        // This can lead to division by zero or very large numbers.
                        // A more robust approach might be needed, but for now, let's use the direct solution.
                        double determinant = (c * c - s * s);
                        if (Math.Abs(determinant) < 1e-8) // Near 45 degrees or 135 degrees
                        {
                            // If rotation is near 45 degrees, the bounding box dimensions are roughly equal.
                            // In this case, it's hard to uniquely determine original width/height from bbox.
                            // Fallback: assume uniform scaling based on the larger dimension.
                            effectiveOleWidth = Math.Max(imgW_orig, imgH_orig);
                            effectiveOleHeight = Math.Min(imgW_orig, imgH_orig);
                            if (rotationAngleRad_ppt % (Math.PI / 2) != 0) // If not 0, 90, 180, 270
                            {
                                // This is a simplification. A more accurate solution would involve
                                // considering the actual bounding box of the rotated image.
                                // For now, we'll use the original dimensions as a best guess.
                                effectiveOleWidth = imgW_orig;
                                effectiveOleHeight = imgH_orig;
                            }
                            else // 0, 90, 180, 270 degrees
                            {
                                if (Math.Abs(rotationAngleRad_ppt % Math.PI) < 1e-8) // 0 or 180
                                {
                                    effectiveOleWidth = imgW_orig;
                                    effectiveOleHeight = imgH_orig;
                                }
                                else // 90 or 270
                                {
                                    effectiveOleWidth = imgH_orig;
                                    effectiveOleHeight = imgW_orig;
                                }
                            }
                        }
                        else
                        {
                            effectiveOleWidth = (oleW_bbox * c - oleH_bbox * s) / determinant;
                            effectiveOleHeight = (oleH_bbox * c - oleW_bbox * s) / determinant;
                        }

                        // 1. Initial uniform scale to match the target width.
                        // Use effectiveOleWidth for scaling
                        double initialScaleFactor = destWidth / effectiveOleWidth;
                        ole.TransformBy(Matrix3d.Scaling(initialScaleFactor, oleCenter));

                        // 2. Move to origin for clean aspect scaling.
                        ole.TransformBy(Matrix3d.Displacement(Point3d.Origin - oleCenter));

                        // 3. Rotation is done in PowerPoint. No rotation here.

                        // 4. Apply aspect ratio correction.
                        // Use effectiveOleHeight for scaling
                        double finalOleHeight = effectiveOleHeight * initialScaleFactor;
                        double aspectScaleFactor = destHeight / finalOleHeight;

                        var localYAxis = destU.GetPerpendicularVector().GetNormal();
                        var v = localYAxis;
                        double s_aspect = aspectScaleFactor; // Renamed to avoid conflict with 's' from sin
                        double sm1_aspect = s_aspect - 1.0;

                        var data = new double[] {
                            1.0 + sm1_aspect * v.X * v.X, sm1_aspect * v.X * v.Y,         sm1_aspect * v.X * v.Z,         0.0,
                            sm1_aspect * v.Y * v.X,         1.0 + sm1_aspect * v.Y * v.Y, sm1_aspect * v.Y * v.Z,         0.0,
                            sm1_aspect * v.Z * v.X,         sm1_aspect * v.Z * v.Y,         1.0 + sm1_aspect * v.Z * v.Z, 0.0,
                            0.0,                     0.0,                     0.0,                     1.0
                        };
                        var scalingMatrix = new Matrix3d(data);
                        ole.TransformBy(scalingMatrix);

                        // 5. Move the fully transformed object to its final destination.
                        var destCenter = target.Pos + (destU * 0.5) + (destV * 0.5);
                        ole.TransformBy(Matrix3d.Displacement(destCenter - Point3d.Origin));
                        // *** MODIFICATION END ***

                            try { ole.Layer = "0"; } catch { }
                        }
                        catch (System.Exception ex1)
                        {
                            ed.WriteMessage($"\nFailed to transform pasted OLE: {ex1.Message}");
                        }
                    }
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
                    _finalPastedOleForZoom = _lastPastedOle;
                    FinishEmbeddingRun(doc, ed, db);
                }
                _lastPastedOle = ObjectId.Null;
            }
            catch (System.Exception ex)
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    doc.Editor.WriteMessage($"\nAn unhandled error occurred in CommandEnded: {ex.Message}");
                    FinishEmbeddingRun(doc, doc.Editor, doc.Database); // Attempt cleanup on error
                }
            }
        }

        private static void FinalCleanupOnIdle(object sender, EventArgs e)
        {
            Application.Idle -= FinalCleanupOnIdle;

            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            if (!_isEmbeddingProcessActive) return;

            try
            {
                ed.WriteMessage("\nPerforming final cleanup for embedding process...");

                _isEmbeddingProcessActive = false;
                _finalPastedOleForZoom = ObjectId.Null;

                DetachHandlers(db, doc);
                PurgeEmbeddedImageDefs(db, ed);
                DetachXrefs(db, ed);
                ClosePowerPoint(ed);

                try
                {
                    if (!_savedClayer.IsNull && db.Clayer != _savedClayer)
                    {
                        db.Clayer = _savedClayer;
                    }
                }
                catch { /* Ignore layer restoration errors */ }

                _savedClayer = ObjectId.Null;

                if (_chainFinalizeAfterEmbed)
                {
                    _chainFinalizeAfterEmbed = false;
                    doc.SendStringToExecute("_.FINALIZE ", true, false, false);
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred during final cleanup: {ex.Message}");
            }
        }

                private static void DetachXrefs(Database db, Editor ed)
        {
            if (_xrefsToDetach.Count == 0) return;

            ed.WriteMessage($"\nAttempting to detach {_xrefsToDetach.Count} XREF(s) ...");

            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId xrefId in _xrefsToDetach)
                {
                    try
                    {
                        // Check if the XREF is still valid and loaded before attempting to detach
                        var btr = tr.GetObject(xrefId, OpenMode.ForRead) as BlockTableRecord;
                        if (btr != null && btr.IsFromExternalReference && btr.IsResolved)
                        {
                            db.DetachXref(xrefId); // Use Database.DetachXref
                            ed.WriteMessage($"\nSuccessfully detached XREF: {btr.Name}");
                        }
                        else if (btr != null && btr.IsFromExternalReference && !btr.IsResolved)
                        {
                            ed.WriteMessage($"\nSkipping detachment of unresolved XREF: {btr.Name}");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nFailed to detach XREF {xrefId}: {ex.Message}");
                    }
                }
                tr.Commit();
            }
            _xrefsToDetach.Clear(); // Clear the list after attempting detachment
        }

        private static void FinishEmbeddingRun(Document doc, Editor ed, Database db)
        {
            Application.Idle += FinalCleanupOnIdle;
        }

        private static void ProcessNextPaste(Document doc, Editor ed)
        {
            if (!_isEmbeddingProcessActive || _pending.Count == 0)
            {
                if (_isEmbeddingProcessActive) FinishEmbeddingRun(doc, ed, doc.Database);
                return;
            }
            var target = _pending.Peek();

            if (!PrepareClipboardWithImageShared(target, ed))
            {
                _pending.Dequeue();
                _activePlacement = null;
                ProcessNextPaste(doc, ed); // Try next one
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