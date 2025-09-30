// Updated SimplerCommands.cs

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
using System.Linq;
using Spire.Pdf;
using Spire.Pdf.Graphics;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

// This is needed for WindowOrchestrator
using System.Diagnostics;

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
            public ObjectId OriginalEntityId; // Changed from ImageId to be generic
            public Point2d[] ClipBoundary; // For PDF cropping
            public double Scale = 1.0;
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


        // --- NEW: Helper for PDF Cropping in PowerPoint ---
        private static bool PrepareClipboardWithCroppedImage(ImagePlacement placement, Editor ed)
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
                Shape pic = shapes.AddPicture(path, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10);


                // --- Cropping Logic ---
                if (placement.ClipBoundary != null && placement.ClipBoundary.Length >= 2)
                {
                    pic.ScaleHeight(1.00f, MsoTriState.msoTrue);
                    pic.ScaleWidth(1.00f, MsoTriState.msoTrue);

                    ed.WriteMessage("\nApplying clipping boundary in PowerPoint...");

                    var clip = placement.ClipBoundary;
                    double minX = clip.Min(p => p.X);
                    double minY = clip.Min(p => p.Y);
                    double maxX = clip.Max(p => p.X);
                    double maxY = clip.Max(p => p.Y);

                    float picWidthInPoints = pic.Width;
                    float picHeightInPoints = pic.Height;

                    bool clipIsPercentage = maxX <= 1.0 + 1e-6 && maxY <= 1.0 + 1e-6 && minX >= -1e-6 && minY >= -1e-6;
                    float cropLeft = 0f, cropTop = 0f, cropRight = 0f, cropBottom = 0f;
                    bool cropComputed = false;

                    if (clipIsPercentage)
                    {
                        cropLeft = (float)(minX * picWidthInPoints);
                        cropTop = (float)((1.0 - maxY) * picHeightInPoints);
                        cropRight = (float)((1.0 - maxX) * picWidthInPoints);
                        cropBottom = (float)(minY * picHeightInPoints);
                        cropComputed = true;

                        ed.WriteMessage($"\n[DEBUG] Clip Percentages: Left={minX:P2}, Bottom={minY:P2}, Right={maxX:P2}, Top={maxY:P2}");
                    }
                    else
                    {
                        double pdfWidthInUnits = placement.U.Length;
                        double pdfHeightInUnits = placement.V.Length;

                        if (pdfWidthInUnits < 1e-6 || pdfHeightInUnits < 1e-6)
                        {
                            ed.WriteMessage("\nWarning: PDF has zero size in AutoCAD. Cannot calculate crop ratios.");
                        }
                        else
                        {
                            double pointsPerUnitX = picWidthInPoints / pdfWidthInUnits;
                            double pointsPerUnitY = picHeightInPoints / pdfHeightInUnits;

                            cropLeft = (float)(minX * pointsPerUnitX);
                            cropTop = (float)((pdfHeightInUnits - maxY) * pointsPerUnitY);
                            cropRight = (float)((pdfWidthInUnits - maxX) * pointsPerUnitX);
                            cropBottom = (float)(minY * pointsPerUnitY);
                            cropComputed = true;
                        }
                    }

                    if (cropComputed)
                    {
                        var picFormat = pic.PictureFormat;
                        picFormat.CropLeft = cropLeft;
                        picFormat.CropTop = cropTop;
                        picFormat.CropRight = cropRight;
                        picFormat.CropBottom = cropBottom;

                        ed.WriteMessage($"\nCrop values (pts): L={cropLeft:F2}, T={cropTop:F2}, R={cropRight:F2}, B={cropBottom:F2}");
                    }
                }

                // Store original (pre-crop) dimensions for later scaling
                _originalPptImageDims[placement.OriginalEntityId] = new Point2d(pic.Width, pic.Height);

                // Apply rotation
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

                float rotationAngleDeg = (float)(-rotationAngleRad * 180.0 / Math.PI);
                pic.Rotation = rotationAngleDeg;

                pic.Copy();
                return true;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to prepare clipboard for '{placement.Path}': {ex.Message}");
                return false;
            }
        }

        // --- OLD Helper for XREF Images (uncropped) ---
        private static bool PrepareClipboardWithImageShared(ImagePlacement placement, Editor ed)
        {
            // This is the same function as before, now repurposed for the XREF workflow specifically.
            // It will be called by EMBEDFROMXREFS.
            // The new EMBEDFROMPDF will call PrepareClipboardWithCroppedImage.
            placement.ClipBoundary = null; // Ensure no cropping happens for this path
            return PrepareClipboardWithCroppedImage(placement, ed);
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
                try { AutoCADApp.Idle -= Application_OnIdleSendPastePoint; } catch { }
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
                    _activePasteDocument = AutoCADApp.DocumentManager.MdiActiveDocument;
                }

                AutoCADApp.Idle += Application_OnIdleSendPastePoint;
                _pastePointHandlerAttached = true;
            }
            catch { }
        }

        private static void Application_OnIdleSendPastePoint(object sender, EventArgs e)
        {
            try
            {
                AutoCADApp.Idle -= Application_OnIdleSendPastePoint;
                _pastePointHandlerAttached = false;

                var doc = _activePasteDocument ?? AutoCADApp.DocumentManager.MdiActiveDocument;
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

                var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
                if (doc == null) return;
                var db = doc.Database;
                var ed = doc.Editor;

                _waitingForPasteStart = false;
                if (_pastePointHandlerAttached)
                {
                    try { AutoCADApp.Idle -= Application_OnIdleSendPastePoint; } catch { }
                    _pastePointHandlerAttached = false;
                }

                if (_pending.Count == 0 && _activePlacement == null)
                {
                    return;
                }

                if (_lastPastedOle.IsNull)
                {
                    ed.WriteMessage("\nSkipping an item because no OLE object was created.");
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
                                ed.WriteMessage($"\nSkipping degenerate original entity {target.OriginalEntityId} with zero area.");
                                ole.Erase();
                                tr.Commit(); // Commit the erase
                                return;
                            }

                            Point2d originalPptDims = Point2d.Origin;
                            if (_originalPptImageDims.ContainsKey(target.OriginalEntityId))
                            {
                                originalPptDims = _originalPptImageDims[target.OriginalEntityId];
                                _originalPptImageDims.Remove(target.OriginalEntityId);
                            }
                            else
                            {
                                ed.WriteMessage($"\nWarning: Original PowerPoint image dimensions not found for {target.OriginalEntityId}. Scaling might be incorrect.");
                                originalPptDims = new Point2d(oleW_bbox, oleH_bbox);
                            }

                            double imgW_orig = originalPptDims.X;
                            double imgH_orig = originalPptDims.Y;

                            double rotationAngleRad_ppt = Math.Atan2(destU.Y, destU.X);
                            bool picIsLandscape = imgW_orig > imgH_orig;
                            bool destIsLandscape = destWidth > destHeight;

                            if (picIsLandscape != destIsLandscape)
                            {
                                rotationAngleRad_ppt -= Math.PI / 2.0;
                            }

                            double c = Math.Abs(Math.Cos(rotationAngleRad_ppt));
                            double s = Math.Abs(Math.Sin(rotationAngleRad_ppt));

                            double effectiveOleWidth, effectiveOleHeight;

                            double determinant = (c * c - s * s);
                            if (Math.Abs(determinant) < 1e-8)
                            {
                                effectiveOleWidth = Math.Max(imgW_orig, imgH_orig);
                                effectiveOleHeight = Math.Min(imgW_orig, imgH_orig);
                                if (rotationAngleRad_ppt % (Math.PI / 2) != 0)
                                {
                                    effectiveOleWidth = imgW_orig;
                                    effectiveOleHeight = imgH_orig;
                                }
                                else
                                {
                                    if (Math.Abs(rotationAngleRad_ppt % Math.PI) < 1e-8)
                                    {
                                        effectiveOleWidth = imgW_orig;
                                        effectiveOleHeight = imgH_orig;
                                    }
                                    else
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

                            double initialScaleFactor = destWidth / effectiveOleWidth;
                            ole.TransformBy(Matrix3d.Scaling(initialScaleFactor, oleCenter));

                            ole.TransformBy(Matrix3d.Displacement(Point3d.Origin - oleCenter));

                            double finalOleHeight = effectiveOleHeight * initialScaleFactor;
                            double aspectScaleFactor = destHeight / finalOleHeight;

                            var localYAxis = destU.GetPerpendicularVector().GetNormal();
                            var v = localYAxis;
                            double s_aspect = aspectScaleFactor;
                            double sm1_aspect = s_aspect - 1.0;

                            var data = new double[] {
                                1.0 + sm1_aspect * v.X * v.X, sm1_aspect * v.X * v.Y,         sm1_aspect * v.X * v.Z,         0.0,
                                sm1_aspect * v.Y * v.X,         1.0 + sm1_aspect * v.Y * v.Y, sm1_aspect * v.Y * v.Z,         0.0,
                                sm1_aspect * v.Z * v.X,         sm1_aspect * v.Z * v.Y,         1.0 + sm1_aspect * v.Z * v.Z, 0.0,
                                0.0,                     0.0,                     0.0,                     1.0
                            };
                            var scalingMatrix = new Matrix3d(data);
                            ole.TransformBy(scalingMatrix);

                            var destCenter = target.Pos + (destU * 0.5) + (destV * 0.5);
                            ole.TransformBy(Matrix3d.Displacement(destCenter - Point3d.Origin));

                            try { ole.Layer = "0"; } catch { }
                        }
                        catch (System.Exception ex1)
                        {
                            ed.WriteMessage($"\nFailed to transform pasted OLE: {ex1.Message}");
                        }
                    }
                    try
                    {
                        var originalEnt = tr.GetObject(target.OriginalEntityId, OpenMode.ForWrite, false) as Entity;
                        if (originalEnt != null)
                        {
                            if (originalEnt is RasterImage imgEnt)
                            {
                                try
                                {
                                    if (!imgEnt.ImageDefId.IsNull)
                                        _imageDefsToPurge.Add(imgEnt.ImageDefId);
                                }
                                catch { }
                            }

                            var layer = (LayerTableRecord)tr.GetObject(originalEnt.LayerId, OpenMode.ForRead);
                            bool relock = false;
                            if (layer.IsLocked)
                            {
                                layer.UpgradeOpen();
                                layer.IsLocked = false;
                                relock = true;
                            }
                            originalEnt.Erase();
                            if (relock)
                            {
                                layer.IsLocked = true;
                            }
                        }
                    }
                    catch (System.Exception exErase)
                    {
                        ed.WriteMessage($"\nWarning: failed to erase source entity: {exErase.Message}");
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
                var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    doc.Editor.WriteMessage($"\nAn unhandled error occurred in CommandEnded: {ex.Message}");
                    FinishEmbeddingRun(doc, doc.Editor, doc.Database); // Attempt cleanup on error
                }
            }
        }

        private static void FinalCleanupOnIdle(object sender, EventArgs e)
        {
            AutoCADApp.Idle -= FinalCleanupOnIdle;

            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
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
            AutoCADApp.Idle += FinalCleanupOnIdle;
        }

        private static void ProcessNextPaste(Document doc, Editor ed)
        {
            if (!_isEmbeddingProcessActive || _pending.Count == 0)
            {
                if (_isEmbeddingProcessActive) FinishEmbeddingRun(doc, ed, doc.Database);
                return;
            }
            var target = _pending.Peek();

            // --- MODIFIED: Use the appropriate clipboard prep function ---
            bool clipboardReady;
            if (target.ClipBoundary != null)
            {
                clipboardReady = PrepareClipboardWithCroppedImage(target, ed);
            }
            else
            {
                clipboardReady = PrepareClipboardWithImageShared(target, ed);
            }
            // --- END MODIFICATION ---

            if (!clipboardReady)
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
                try { AutoCADApp.Idle -= Application_OnIdleSendPastePoint; } catch { }
                _pastePointHandlerAttached = false;
            }
            doc.SendStringToExecute("_.PASTECLIP\n", true, false, false);
        }

        // --- NEW COMMAND: EMBEDFROMPDF with Percentage-Based Clipping ---
        [CommandMethod("EMBEDFROMPDF", CommandFlags.Modal)]
        public static void EmbedFromPdf()
        {
            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            ed.WriteMessage("\n--- Starting EMBEDFROMPDF (Debug Mode) ---");

            _pending.Clear();
            _isEmbeddingProcessActive = false;
            ObjectId originalClayer = db.Clayer;
            Matrix3d originalUcs = ed.CurrentUserCoordinateSystem;

            try
            {
                ed.CurrentUserCoordinateSystem = Matrix3d.Identity; // Set UCS to World

                ImagePlacement placementToDebug = null;

                // --- Phase 1: Find the FIRST PDF and prepare it ---
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    // Find first PDF
                    (ObjectId refId, ObjectId ownerId) firstPdf = default;
                    bool found = false;
                    foreach (ObjectId btrId in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        if (btr == null) continue;

                        foreach (ObjectId entId in btr)
                        {
                            if (entId.ObjectClass.DxfName.Equals("PDFUNDERLAY", StringComparison.OrdinalIgnoreCase))
                            {
                                firstPdf = (entId, btr.ObjectId);
                                found = true;
                                break;
                            }
                        }
                        if (found) break;
                    }

                    if (!found)
                    {
                        ed.WriteMessage("\nNo valid PDF underlays found to embed.");
                        tr.Commit();
                        return;
                    }

                    ed.WriteMessage($"\nFound a PDF underlay to process for debugging.");

                    var pdfRef = tr.GetObject(firstPdf.refId, OpenMode.ForRead) as UnderlayReference;
                    if (pdfRef == null) { tr.Commit(); return; }

                    var pdfDef = tr.GetObject(pdfRef.DefinitionId, OpenMode.ForRead) as UnderlayDefinition;
                    if (pdfDef == null) { tr.Commit(); return; }

                    // Get PDF page size in drawing units from the reference transform
                    Matrix3d pdfTransform = pdfRef.Transform;
                    Vector3d u_vec = pdfTransform.CoordinateSystem3d.Xaxis;
                    Vector3d v_vec = pdfTransform.CoordinateSystem3d.Yaxis;

                    double pdfWidthInUnits = u_vec.Length;
                    double pdfHeightInUnits = v_vec.Length;

                    ed.WriteMessage($"\n[DEBUG] PDF Size (drawing units): W={pdfWidthInUnits:F4}, H={pdfHeightInUnits:F4}");

                    string resolvedPdfPath = ResolveImagePath(db, pdfDef.SourceFileName);
                    if (string.IsNullOrEmpty(resolvedPdfPath) || !File.Exists(resolvedPdfPath))
                    {
                        ed.WriteMessage($"\nSkipping missing PDF file: {pdfDef.SourceFileName}");
                        tr.Commit();
                        return;
                    }

                    ed.WriteMessage($"\nProcessing {Path.GetFileName(resolvedPdfPath)}...");
                    string pngPath = ConvertPdfToPng(resolvedPdfPath, ed);
                    if (string.IsNullOrEmpty(pngPath))
                    {
                        ed.WriteMessage("\nFailed to convert PDF to PNG. Skipping.");
                        tr.Commit();
                        return;
                    }

                    // --- NEW: Get Clip Boundary and Convert to Percentages ---
                    Point2d[] percentageClipBoundary = null;

                    if (pdfRef.IsClipped)
                    {
                        ed.WriteMessage("\nPDF is clipped. Calculating percentage-based clip boundary...");

                        object rawClipBoundaryObject = pdfRef.GetClipBoundary();
                        Point2d[] wcsClipBoundary = null;

                        if (rawClipBoundaryObject is Point2dCollection rawClipBoundaryCollection)
                        {
                            wcsClipBoundary = new Point2d[rawClipBoundaryCollection.Count];
                            rawClipBoundaryCollection.CopyTo(wcsClipBoundary, 0);
                        }
                        else if (rawClipBoundaryObject is Point2d[] rawClipBoundaryArray)
                        {
                            wcsClipBoundary = rawClipBoundaryArray;
                        }

                        if (wcsClipBoundary != null && wcsClipBoundary.Length >= 2)
                        {
                            // Get min/max of WCS clip boundary (these are in drawing units)
                            double minX_wcs = wcsClipBoundary.Min(p => p.X);
                            double minY_wcs = wcsClipBoundary.Min(p => p.Y);
                            double maxX_wcs = wcsClipBoundary.Max(p => p.X);
                            double maxY_wcs = wcsClipBoundary.Max(p => p.Y);

                            ed.WriteMessage($"\n[DEBUG] WCS Clip (units): MinX={minX_wcs:F4}, MinY={minY_wcs:F4}, MaxX={maxX_wcs:F4}, MaxY={maxY_wcs:F4}");

                            // Get the actual PDF sheet size in inches using Spire.PDF
                            int pageNumber = 1;
                            if (int.TryParse(pdfDef.ItemName, out int pg) && pg > 0)
                            {
                                pageNumber = pg;
                            }

                            Point2d? pdfSheetSizeInches = GetPdfPageSizeInches(resolvedPdfPath, pageNumber);

                            if (!pdfSheetSizeInches.HasValue)
                            {
                                ed.WriteMessage("\nWarning: Could not determine PDF sheet size. Cannot calculate crop percentages.");
                            }
                            else
                            {
                                double sheetWidthInches = pdfSheetSizeInches.Value.X;
                                double sheetHeightInches = pdfSheetSizeInches.Value.Y;

                                ed.WriteMessage($"\n[DEBUG] PDF Sheet Size (inches): W={sheetWidthInches:F4}\", H={sheetHeightInches:F4}\"");

                                // Calculate the scale factor: drawing units per inch
                                double unitsPerInchX = pdfWidthInUnits / sheetWidthInches;
                                double unitsPerInchY = pdfHeightInUnits / sheetHeightInches;

                                ed.WriteMessage($"\n[DEBUG] Scale Factor: {unitsPerInchX:F4} units/inch (X), {unitsPerInchY:F4} units/inch (Y)");

                                // Get the PDF origin in WCS
                                Point3d pdfOrigin = pdfTransform.CoordinateSystem3d.Origin;
                                ed.WriteMessage($"\n[DEBUG] PDF Origin (WCS): ({pdfOrigin.X:F4}, {pdfOrigin.Y:F4})");

                                // Transform the clip boundary from WCS into the PDF's local coordinate system
                                Matrix3d inverseTransform = pdfTransform.Inverse();

                                double minLocalX = double.MaxValue;
                                double minLocalY = double.MaxValue;
                                double maxLocalX = double.MinValue;
                                double maxLocalY = double.MinValue;

                                foreach (var clipPt in wcsClipBoundary)
                                {
                                    Point3d localPt = new Point3d(clipPt.X, clipPt.Y, 0).TransformBy(inverseTransform);
                                    if (localPt.X < minLocalX) minLocalX = localPt.X;
                                    if (localPt.Y < minLocalY) minLocalY = localPt.Y;
                                    if (localPt.X > maxLocalX) maxLocalX = localPt.X;
                                    if (localPt.Y > maxLocalY) maxLocalY = localPt.Y;
                                }

                                ed.WriteMessage($"\n[DEBUG] Clip in PDF Local Space (units): MinU={minLocalX:F4}, MinV={minLocalY:F4}, MaxU={maxLocalX:F4}, MaxV={maxLocalY:F4}");

                                double clipMinX_inches = minLocalX / unitsPerInchX;
                                double clipMinY_inches = minLocalY / unitsPerInchY;
                                double clipMaxX_inches = maxLocalX / unitsPerInchX;
                                double clipMaxY_inches = maxLocalY / unitsPerInchY;

                                clipMinX_inches = Math.Max(0.0, Math.Min(sheetWidthInches, clipMinX_inches));
                                clipMaxX_inches = Math.Max(0.0, Math.Min(sheetWidthInches, clipMaxX_inches));
                                clipMinY_inches = Math.Max(0.0, Math.Min(sheetHeightInches, clipMinY_inches));
                                clipMaxY_inches = Math.Max(0.0, Math.Min(sheetHeightInches, clipMaxY_inches));

                                if (clipMaxX_inches < clipMinX_inches)
                                {
                                    double swap = clipMinX_inches;
                                    clipMinX_inches = clipMaxX_inches;
                                    clipMaxX_inches = swap;
                                }

                                if (clipMaxY_inches < clipMinY_inches)
                                {
                                    double swap = clipMinY_inches;
                                    clipMinY_inches = clipMaxY_inches;
                                    clipMaxY_inches = swap;
                                }

                                ed.WriteMessage($"\n[DEBUG] Clip in Inches (from origin): MinX={clipMinX_inches:F4}\", MinY={clipMinY_inches:F4}\", MaxX={clipMaxX_inches:F4}\", MaxY={clipMaxY_inches:F4}\"");

                                // Calculate percentages based on the sheet size
                                double leftPercent = minX_wcs / sheetWidthInches;
                                double bottomPercent = minY_wcs / sheetHeightInches;
                                double rightPercent = maxX_wcs / sheetWidthInches;
                                double topPercent = maxY_wcs / sheetHeightInches;

                                leftPercent = Math.Max(0.0, Math.Min(1.0, leftPercent));
                                bottomPercent = Math.Max(0.0, Math.Min(1.0, bottomPercent));
                                rightPercent = Math.Max(0.0, Math.Min(1.0, rightPercent));
                                topPercent = Math.Max(0.0, Math.Min(1.0, topPercent));

                                ed.WriteMessage($"\n[DEBUG] Clip Percentages: Left={leftPercent:P2}, Bottom={bottomPercent:P2}, Right={rightPercent:P2}, Top={topPercent:P2}");

                                // Store as Point2d array: [0] = min (left, bottom), [1] = max (right, top)
                                percentageClipBoundary = new Point2d[]
                                {
                            new Point2d(leftPercent, bottomPercent),
                            new Point2d(rightPercent, topPercent)
                                };
                            }
                        }
                        else
                        {
                            ed.WriteMessage("\nWarning: Could not retrieve valid clipping boundary.");
                        }
                    }
                    else
                    {
                        ed.WriteMessage("\nPDF is not clipped. No cropping will be applied.");
                    }

                    placementToDebug = new ImagePlacement
                    {
                        Path = pngPath,
                        Pos = pdfTransform.CoordinateSystem3d.Origin,
                        U = u_vec,
                        V = v_vec,
                        OriginalEntityId = pdfRef.ObjectId,
                        ClipBoundary = percentageClipBoundary // Store percentage-based boundary
                    };

                    tr.Commit();
                }

                // --- Phase 2: Show in PowerPoint and stop ---
                if (placementToDebug == null)
                {
                    ed.WriteMessage("\nCould not prepare PDF for debugging.");
                    return;
                }

                ed.WriteMessage($"\nOpening PowerPoint to show cropped image. The command will stop here.");
                ShowImageInPowerPointForDebugPercentage(placementToDebug, ed);
                ed.WriteMessage("\nPowerPoint should be open with the cropped image. Please inspect it.");
                ed.WriteMessage("\nCommand finished.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[!] An error occurred during the PDF embedding debug process: {ex.Message}");
            }
            finally
            {
                try
                {
                    if (db.Clayer != originalClayer) db.Clayer = originalClayer;
                }
                catch { }
                ed.CurrentUserCoordinateSystem = originalUcs;
            }
        }

        // --- NEW: Debug helper using percentage-based cropping ---
        private static void ShowImageInPowerPointForDebugPercentage(ImagePlacement placement, Editor ed)
        {
            try
            {
                if (!EnsurePowerPoint(ed)) return;
                dynamic slide = _pptPresentationShared.Slides[1];
                PowerPoint.Shapes shapes = slide.Shapes;

                // Clear previous shapes
                for (int i = shapes.Count; i >= 1; i--)
                {
                    try { shapes[i].Delete(); } catch { }
                }

                string path = placement.Path;
                Shape pic = shapes.AddPicture(path, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10);

                // --- Percentage-Based Cropping Logic ---
                if (placement.ClipBoundary != null && placement.ClipBoundary.Length >= 2)
                {
                    ed.WriteMessage("\nApplying percentage-based clipping boundary in PowerPoint...");

                    pic.ScaleWidth(1.00f, MsoTriState.msoTrue);

                    // Get the size of the inserted picture in PowerPoint points
                    float picWidthInPoints = pic.Width;
                    float picHeightInPoints = pic.Height;

                    ed.WriteMessage($"\n[DEBUG] PPT Picture Dimensions (points): W={picWidthInPoints:F4}, H={picHeightInPoints:F4}");

                    // Extract percentages from ClipBoundary
                    // [0] = (leftPercent, bottomPercent), [1] = (rightPercent, topPercent)
                    double leftPercent = placement.ClipBoundary[0].X;
                    double bottomPercent = placement.ClipBoundary[0].Y;
                    double rightPercent = placement.ClipBoundary[1].X;
                    double topPercent = placement.ClipBoundary[1].Y;

                    ed.WriteMessage($"\n[DEBUG] Clip Percentages: Left={leftPercent:P2}, Bottom={bottomPercent:P2}, Right={rightPercent:P2}, Top={topPercent:P2}");

                    // Calculate crop amounts in points
                    // Crop from left = leftPercent * width
                    float cropLeft = (float)(leftPercent * picWidthInPoints);

                    // Crop from top = (1 - topPercent) * height
                    float cropTop = (float)((1.0 - topPercent) * picHeightInPoints);

                    // Crop from right = (1 - rightPercent) * width
                    float cropRight = (float)((1.0 - rightPercent) * picWidthInPoints);

                    // Crop from bottom = bottomPercent * height
                    float cropBottom = (float)(bottomPercent * picHeightInPoints);

                    // Apply the crop
                    var picFormat = pic.PictureFormat;
                    picFormat.CropLeft = cropLeft;
                    picFormat.CropTop = cropTop;
                    picFormat.CropRight = cropRight;
                    picFormat.CropBottom = cropBottom;

                    ed.WriteMessage($"\n[DEBUG] Crop values (pts): L={cropLeft:F2}, T={cropTop:F2}, R={cropRight:F2}, B={cropBottom:F2}");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to show image in PowerPoint for debug: {ex.Message}");
            }
        }

        // --- EXISTING COMMAND: Get PDF page number ---
        [CommandMethod("GetPdfPageNumber", CommandFlags.Modal)]
        public static void GetPdfPageNumber()
        {
            Document doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Set up selection options to filter for only PDF underlays
            var selOpts = new PromptSelectionOptions();
            selOpts.MessageForAdding = "\nSelect a PDF underlay reference:";
            selOpts.SingleOnly = true;

            var filter = new SelectionFilter(new TypedValue[] {
                new TypedValue((int)DxfCode.Start, "PDFUNDERLAY")
            });

            PromptSelectionResult selRes = ed.GetSelection(selOpts, filter);

            if (selRes.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n*Cancel*");
                return;
            }

            SelectionSet selSet = selRes.Value;
            ObjectId selectedId = selSet.GetObjectIds()[0];

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Get the selected PDF underlay reference
                    var pdfRef = tr.GetObject(selectedId, OpenMode.ForRead) as UnderlayReference;
                    if (pdfRef == null)
                    {
                        ed.WriteMessage("\nSelected object is not a valid PDF underlay.");
                        return;
                    }

                    // Get the definition associated with the reference
                    var pdfDef = tr.GetObject(pdfRef.DefinitionId, OpenMode.ForRead) as UnderlayDefinition;
                    if (pdfDef == null)
                    {
                        ed.WriteMessage("\nCould not retrieve the PDF definition for the selected reference.");
                        return;
                    }

                    // The 'ItemName' property of the definition holds the page number for PDFs.
                    string pageNumber = pdfDef.ItemName;
                    string pdfFileName = Path.GetFileName(pdfDef.SourceFileName);

                    if (string.IsNullOrEmpty(pageNumber))
                    {
                        ed.WriteMessage($"\nThe selected PDF reference ('{pdfFileName}') does not have a specific page number associated with it.");
                    }
                    else
                    {
                        ed.WriteMessage($"\nThe selected PDF reference is on page: {pageNumber} of '{pdfFileName}'.");
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred: {ex.Message}");
            }
        }

        // ===================== NEW: GETPDFSHEETSIZE =====================
        /// <summary>
        /// Lets the user select a PDF underlay reference and reports the sheet size in inches.
        /// Uses Spire.PDF to determine the page size.
        /// </summary>
        [CommandMethod("GETPDFSHEETSIZE", CommandFlags.Modal)]
        public static void GetPdfSheetSize()
        {
            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            var opts = new PromptSelectionOptions
            {
                MessageForAdding = "\nSelect a PDF underlay reference:",
                SingleOnly = true
            };
            var filter = new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "PDFUNDERLAY") });

            var sel = ed.GetSelection(opts, filter);
            if (sel.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n*Cancel*");
                return;
            }

            var id = sel.Value.GetObjectIds().FirstOrDefault();
            if (id.IsNull) return;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var ur = tr.GetObject(id, OpenMode.ForRead) as UnderlayReference;
                    if (ur == null)
                    {
                        ed.WriteMessage("\nSelected object is not a PDF underlay.");
                        return;
                    }

                    var def = tr.GetObject(ur.DefinitionId, OpenMode.ForRead) as UnderlayDefinition;
                    if (def == null)
                    {
                        ed.WriteMessage("\nCould not retrieve PDF definition.");
                        return;
                    }

                    string resolvedPath = ResolveImagePath(db, def.SourceFileName);
                    if (string.IsNullOrEmpty(resolvedPath) || !File.Exists(resolvedPath))
                    {
                        ed.WriteMessage($"\nCould not find PDF file: {def.SourceFileName}");
                        return;
                    }

                    int pageNumber = 1;
                    if (int.TryParse(def.ItemName, out int pg) && pg > 0)
                    {
                        pageNumber = pg;
                    }

                    try
                    {
                        Point2d? pageSize = GetPdfPageSizeInches(resolvedPath, pageNumber);
                        if (pageSize.HasValue)
                        {
                            double wIn = pageSize.Value.X;
                            double hIn = pageSize.Value.Y;
                            string orient = wIn >= hIn ? "Landscape" : "Portrait";
                            string file = Path.GetFileName(def.SourceFileName);

                            ed.WriteMessage($"\nPDF sheet size (inches): {wIn:F2}\" × {hIn:F2}\"  [{orient}] (Page {pageNumber})");
                            ed.WriteMessage($"\nSource: {file}");
                        }
                        else
                        {
                            ed.WriteMessage("\nCould not determine PDF page size using Spire.PDF.");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nError while getting PDF size: {ex.Message}");
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
            }
        }

        private static Point2d? GetPdfPageSizeInches(string pdfPath, int pageNumber)
        {
            try
            {
                // Load a PDF file
                PdfDocument pdf = new PdfDocument();
                pdf.LoadFromFile(pdfPath);

                // Page numbers in Spire.PDF are 0-indexed
                int pageIndex = pageNumber - 1;

                if (pageIndex < 0 || pageIndex >= pdf.Pages.Count)
                {
                    var ed = AutoCADApp.DocumentManager.MdiActiveDocument.Editor;
                    ed.WriteMessage($"\nError: Page number {pageNumber} is out of range.");
                    return null;
                }

                // Get the specified page
                PdfPageBase page = pdf.Pages[pageIndex];

                // Get the width and height of the page in "point"
                float pointWidth = page.Size.Width;
                float pointHeight = page.Size.Height;

                // Convert to inches (1 inch = 72 points)
                float inchWidth = pointWidth / 72f;
                float inchHeight = pointHeight / 72f;

                pdf.Close();

                return new Point2d(inchWidth, inchHeight);
            }
            catch (System.Exception ex)
            {
                var ed = AutoCADApp.DocumentManager.MdiActiveDocument.Editor;
                ed.WriteMessage($"\nAn exception occurred while getting PDF page size: {ex.Message}");
                return null;
            }
        }

        [CommandMethod("GETCLIPBOUNDARY", CommandFlags.Modal)]
        public static void GetClippingBoundaryInches()
        {
            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            var opts = new PromptSelectionOptions
            {
                MessageForAdding = "\nSelect a clipped underlay (PDF, DWF, DGN, or Image):",
                SingleOnly = true
            };
            var filter = new SelectionFilter(new[] {
                new TypedValue((int)DxfCode.Start, "PDFUNDERLAY,DWFUNDERLAY,DGNUNDERLAY,IMAGE")
            });

            var sel = ed.GetSelection(opts, filter);
            if (sel.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n*Cancel*");
                return;
            }

            var id = sel.Value.GetObjectIds().FirstOrDefault();
            if (id.IsNull) return;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent == null)
                    {
                        ed.WriteMessage("\nSelected object is not a valid entity.");
                        return;
                    }

                    Point2d[] boundary = null;
                    string underlayType = "";
                    object boundaryObject = null;

                    if (ent is UnderlayReference ur)
                    {
                        if (!ur.IsClipped)
                        {
                            ed.WriteMessage("\nThe selected underlay is not clipped.");
                            return;
                        }
                        boundaryObject = ur.GetClipBoundary();
                        underlayType = ur.GetRXClass().DxfName.Replace("UNDERLAY", "");
                    }
                    else if (ent is RasterImage ri)
                    {
                        if (!ri.IsClipped)
                        {
                            ed.WriteMessage("\nThe selected image is not clipped.");
                            return;
                        }
                        boundaryObject = ri.GetClipBoundary();
                        underlayType = "Image";
                    }

                    if (boundaryObject is Point2dCollection boundaryCollection)
                    {
                        boundary = new Point2d[boundaryCollection.Count];
                        boundaryCollection.CopyTo(boundary, 0);
                    }
                    else if (boundaryObject is Point2d[] boundaryArray)
                    {
                        boundary = boundaryArray;
                    }

                    if (boundary == null || boundary.Length == 0)
                    {
                        ed.WriteMessage("\nCould not retrieve a valid clipping boundary.");
                        return;
                    }

                    ed.WriteMessage($"\nClipping boundary for {underlayType} in WCS coordinates (inches):");

                    // Get the current drawing units
                    var insUnits = db.Insunits;
                    double conversionFactor = 1.0; // Default to inches

                    // Find the conversion factor to inches
                    switch (insUnits)
                    {
                        case UnitsValue.Millimeters:
                            conversionFactor = 1.0 / 25.4;
                            break;
                        case UnitsValue.Centimeters:
                            conversionFactor = 1.0 / 2.54;
                            break;
                        case UnitsValue.Meters:
                            conversionFactor = 1.0 / 0.0254;
                            break;
                        case UnitsValue.Feet:
                            conversionFactor = 12.0;
                            break;
                            // Add other cases as needed
                    }


                    for (int i = 0; i < boundary.Length; i++)
                    {
                        Point2d pt = boundary[i];
                        ed.WriteMessage($"  Vertex {i + 1}: (X={pt.X * conversionFactor:F4}, Y={pt.Y * conversionFactor:F4})");
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred: {ex.Message}");
            }
        }

        [CommandMethod("GETPDFCLIPPING", CommandFlags.Modal)]
        public static void GetPdfClippingBoundary()
        {
            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            var opts = new PromptSelectionOptions
            {
                MessageForAdding = "\nSelect a clipped PDF underlay:",
                SingleOnly = true
            };
            var filter = new SelectionFilter(new[] {
                new TypedValue((int)DxfCode.Start, "PDFUNDERLAY")
            });

            var sel = ed.GetSelection(opts, filter);
            if (sel.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n*Cancel*");
                return;
            }

            var id = sel.Value.GetObjectIds().FirstOrDefault();
            if (id.IsNull) return;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent == null)
                    {
                        ed.WriteMessage("\nSelected object is not a valid entity.");
                        return;
                    }

                    if (ent is UnderlayReference ur)
                    {
                        if (!ur.IsClipped)
                        {
                            ed.WriteMessage("\nThe selected PDF is not clipped.");
                            return;
                        }

                        object boundaryObject = ur.GetClipBoundary();
                        Point2d[] boundary = null;

                        if (boundaryObject is Point2dCollection boundaryCollection)
                        {
                            boundary = new Point2d[boundaryCollection.Count];
                            boundaryCollection.CopyTo(boundary, 0);
                        }
                        else if (boundaryObject is Point2d[] boundaryArray)
                        {
                            boundary = boundaryArray;
                        }

                        if (boundary != null && boundary.Length >= 2)
                        {
                            Point2d minPoint = boundary[0];
                            Point2d maxPoint = boundary[1];
                            ed.WriteMessage($"\n[DEBUG] Raw WCS Clip Boundary (drawing units):  -> ({minPoint.X:F4}, {minPoint.Y:F4})  -> ({maxPoint.X:F4}, {maxPoint.Y:F4})");
                        }
                        else
                        {
                            ed.WriteMessage("\nCould not retrieve a valid clipping boundary (requires at least 2 points).");
                        }
                    }
                    else
                    {
                        ed.WriteMessage("\nSelected object is not a PDF underlay.");
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred: {ex.Message}");
            }
        }
    }
}
