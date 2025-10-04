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
            public ObjectId TargetBtrId = ObjectId.Null;
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
        // *** MODIFICATION: Track PDF definitions for detachment ***
        internal static readonly HashSet<ObjectId> _pdfDefinitionsToDetach = new HashSet<ObjectId>();
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


        // --- Helper for PDF Cropping in PowerPoint ---
        private static bool PrepareClipboardWithCroppedImage(ImagePlacement placement, Editor ed)
        {
            dynamic slide = null;
            dynamic shapes = null;
            Shape pic = null;
            dynamic picFormat = null;
            dynamic pngRange = null;
            Shape pngShape = null;

            try
            {
                if (!EnsurePowerPoint(ed)) return false;
                slide = _pptPresentationShared.Slides[1]; // 1-based
                shapes = slide.Shapes;
                // Clear previous shapes
                for (int i = shapes.Count; i >= 1; i--)
                {
                    try { shapes[i].Delete(); } catch { }
                }

                string path = placement.Path;
                pic = shapes.AddPicture(path, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10);


                // --- Cropping Logic ---
                if (placement.ClipBoundary != null && placement.ClipBoundary.Length >= 2)
                {
                    // Reset scale to ensure accurate crop calculations
                    pic.ScaleHeight(1, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromTopLeft);
                    pic.ScaleWidth(1, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromTopLeft);

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

                        // ed.WriteMessage($"\n[DEBUG] Clip Percentages: Left={minX:P2}, Bottom={minY:P2}, Right={maxX:P2}, Top={maxY:P2}");
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
                        picFormat = pic.PictureFormat;
                        picFormat.CropLeft = cropLeft;
                        picFormat.CropTop = cropTop;
                        picFormat.CropRight = cropRight;
                        picFormat.CropBottom = cropBottom;

                        ed.WriteMessage($"\nCrop values (pts): L={cropLeft:F2}, T={cropTop:F2}, R={cropRight:F2}, B={cropBottom:F2}");

                        // *** MODIFICATION: Removed debug pause ***
                    }
                }

                try { System.Windows.Forms.Clipboard.Clear(); } catch { }

                pic.Copy();

                try
                {
                    pngRange = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG);
                    try { pngShape = pngRange[1]; } catch { }
                    if (pngShape != null)
                    {
                        pngShape.Copy();
                        try { pngShape.Delete(); } catch { }
                    }
                    else
                    {
                        try { pngRange.Delete(); } catch { }
                    }
                }
                catch { }


                try { pic.Delete(); } catch { }

                return true;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to prepare clipboard for '{placement.Path}': {ex.Message}");
                return false;
            }
            finally
            {
                // This is the fix: Proactively release COM objects to prevent resource buildup
                // that can cause mid-process freezes.
                if (pngShape != null) Marshal.ReleaseComObject(pngShape);
                if (pngRange != null) Marshal.ReleaseComObject(pngRange);
                if (picFormat != null) Marshal.ReleaseComObject(picFormat);
                if (pic != null) Marshal.ReleaseComObject(pic);
                if (shapes != null) Marshal.ReleaseComObject(shapes);
                if (slide != null) Marshal.ReleaseComObject(slide);
            }
        }

        // --- Helper for XREF Images (uncropped) ---
        private static bool PrepareClipboardWithImageShared(ImagePlacement placement, Editor ed)
        {
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

        // *** MODIFICATION: Detach PDF definitions ***
        private static void DetachPdfDefinitions(Database db, Editor ed)
        {
            if (_pdfDefinitionsToDetach.Count == 0) return;
            ed.WriteMessage($"\nAttempting to detach {_pdfDefinitionsToDetach.Count} PDF definition(s)...");

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    // PDF definitions are stored in the ACAD_PDFDEFINITIONS dictionary
                    var named = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                    if (!named.Contains("ACAD_PDFDEFINITIONS"))
                    {
                        tr.Commit();
                        return;
                    }

                    var pdfDict = (DBDictionary)tr.GetObject(named.GetAt("ACAD_PDFDEFINITIONS"), OpenMode.ForWrite);
                    int detachedCount = 0;

                    // Create a list of entries to remove to avoid modifying the collection while iterating
                    var entriesToRemove = new List<string>();

                    foreach (DBDictionaryEntry entry in pdfDict)
                    {
                        if (_pdfDefinitionsToDetach.Contains(entry.Value))
                        {
                            entriesToRemove.Add(entry.Key);
                            try
                            {
                                // Erasing the definition object itself
                                var defObj = tr.GetObject(entry.Value, OpenMode.ForWrite);
                                if (!defObj.IsErased)
                                {
                                    defObj.Erase();
                                }
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage($"\nFailed to erase PDF definition object: {ex.Message}");
                            }
                        }
                    }

                    // Remove entries from the dictionary
                    foreach (string key in entriesToRemove)
                    {
                        try
                        {
                            pdfDict.Remove(key);
                            detachedCount++;
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nFailed to remove PDF dictionary entry '{key}': {ex.Message}");
                        }
                    }

                    ed.WriteMessage($"\nSuccessfully detached {detachedCount} PDF definition(s).");
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError during PDF detachment: {ex.Message}");
            }
            finally
            {
                _pdfDefinitionsToDetach.Clear();
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

                doc.SendStringToExecute($"{x},{y}\n", true, false, false);
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
                            double oleWidth = Math.Max(1e-8, ext.MaxPoint.X - ext.MinPoint.X);
                            double oleHeight = Math.Max(1e-8, ext.MaxPoint.Y - ext.MinPoint.Y);

                            var destU = target.U;
                            var destV = target.V;
                            double destWidth = destU.Length;
                            double destHeight = destV.Length;

                            if (destWidth < 1e-8 || destHeight < 1e-8 || oleWidth < 1e-8 || oleHeight < 1e-8)
                            {
                                ed.WriteMessage($"\nSkipping transform: invalid geometry for {target.OriginalEntityId}.");
                                ole.Erase();
                                tr.Commit();
                                return;
                            }

                            // Preserve aspect ratio
                            double oleAspectRatio = oleWidth / oleHeight;
                            double destAspectRatio = destWidth / destHeight;

                            Vector3d finalU = destU;
                            Vector3d finalV = destV;
                            Point3d finalPos = target.Pos;

                            if (Math.Abs(oleAspectRatio - destAspectRatio) > 1e-6)
                            {
                                // Aspect ratios differ, so we need to adjust.
                                // We will fit the OLE object inside the destination box.
                                if (oleAspectRatio > destAspectRatio)
                                {
                                    // OLE is wider than destination. Fit to width.
                                    double newHeight = destWidth / oleAspectRatio;
                                    finalV = destV.GetNormal() * newHeight;
                                    // Center vertically
                                    finalPos = target.Pos + (destV - finalV) / 2.0;
                                }
                                else
                                {
                                    // OLE is taller than destination. Fit to height.
                                    double newWidth = destHeight * oleAspectRatio;
                                    finalU = destU.GetNormal() * newWidth;
                                    // Center horizontally
                                    finalPos = target.Pos + (destU - finalU) / 2.0;
                                }
                            }

                            Vector3d destNormal = finalU.CrossProduct(finalV);
                            if (destNormal.Length < 1e-8)
                            {
                                destNormal = Vector3d.ZAxis;
                            }

                            Point3d oleOrigin = ext.MinPoint;
                            Vector3d oleU = new Vector3d(oleWidth, 0.0, 0.0);
                            Vector3d oleV = new Vector3d(0.0, oleHeight, 0.0);
                            Vector3d oleNormal = oleU.CrossProduct(oleV);

                            var align = Matrix3d.AlignCoordinateSystem(
                                oleOrigin,
                                oleU,
                                oleV,
                                oleNormal,
                                finalPos, // Use adjusted position
                                finalU,   // Use adjusted U-vector
                                finalV,   // Use adjusted V-vector
                                destNormal);

                            ole.TransformBy(align);

                            var finalExt = ole.GeometricExtents;
                            // ed.WriteMessage($"\n[DEBUG] Transform complete: src=({oleWidth:F4},{oleHeight:F4}) dest=({destWidth:F4},{destHeight:F4}) final=({finalExt.MaxPoint.X - finalExt.MinPoint.X:F4},{finalExt.MaxPoint.Y - finalExt.MinPoint.Y:F4}) pos=({finalPos.X:F4},{finalPos.Y:F4},{finalPos.Z:F4})");

                            try { ole.Layer = "0"; } catch { }

                            // *** MODIFICATION: Match draw order of the original entity ***
                            try
                            {
                                // Get the BlockTableRecord that owns the OLE (and the original entity)
                                var btr = tr.GetObject(ole.OwnerId, OpenMode.ForWrite) as BlockTableRecord;
                                if (btr != null)
                                {
                                    var dot = tr.GetObject(btr.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;
                                    if (dot != null)
                                    {
                                        var idsToMove = new ObjectIdCollection();
                                        idsToMove.Add(ole.ObjectId);
                                        // Move the new OLE to be right above the original entity it is replacing
                                        dot.MoveAbove(idsToMove, target.OriginalEntityId);
                                    }
                                }
                            }
                            catch (System.Exception exOrder)
                            {
                                ed.WriteMessage($"\nWarning: Could not match draw order: {exOrder.Message}");
                            }
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
                            originalEnt.Erase(); // Erase the original after setting draw order
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

                // This is the fix: Proactively manage memory and UI thread responsiveness
                // after each paste operation to prevent a large, blocking freeze.
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Thread.Sleep(50); // A brief pause to let message pumps catch up.

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
                // *** MODIFICATION: Call PDF detachment ***
                DetachPdfDefinitions(db, ed);
                DetachXrefs(db, ed);

                // This is the fix: Safely manage window focus before closing PowerPoint.
                // This helps prevent freezes by ensuring AutoCAD has focus and PowerPoint is
                // in a stable (minimized) state before issuing the quit command via COM.
                WindowOrchestrator.EndPptInteraction();

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

        private static bool EnsurePlacementSpace(Document doc, ImagePlacement placement, Editor ed)
        {
            if (doc == null || placement == null || placement.TargetBtrId.IsNull)
            {
                return true;
            }

            var db = doc.Database;
            if (db.CurrentSpaceId == placement.TargetBtrId)
            {
                return true;
            }

            try
            {
                string layoutName = null;
                using (var tr = db.TransactionManager.StartOpenCloseTransaction())
                {
                    var btr = tr.GetObject(placement.TargetBtrId, OpenMode.ForRead) as BlockTableRecord;
                    if (btr == null || !btr.IsLayout)
                    {
                        ed?.WriteMessage("\nSkipping placement: owner space is not a layout or is unavailable.");
                        return false;
                    }

                    var layout = tr.GetObject(btr.LayoutId, OpenMode.ForRead) as Layout;
                    if (layout == null)
                    {
                        ed?.WriteMessage("\nSkipping placement: layout information could not be resolved.");
                        return false;
                    }

                    layoutName = layout.LayoutName;
                }

                var lm = LayoutManager.Current;
                if (lm == null)
                {
                    ed?.WriteMessage("\nUnable to activate layout manager for placement.");
                    return false;
                }

                if (!string.Equals(lm.CurrentLayout, layoutName, StringComparison.OrdinalIgnoreCase))
                {
                    lm.CurrentLayout = layoutName;
                }

                return db.CurrentSpaceId == placement.TargetBtrId;
            }
            catch (System.Exception ex)
            {
                ed?.WriteMessage($"\nFailed to switch to target layout: {ex.Message}");
                return false;
            }
        }

        private static void ProcessNextPaste(Document doc, Editor ed)
        {
            if (!_isEmbeddingProcessActive || _pending.Count == 0)
            {
                if (_isEmbeddingProcessActive) FinishEmbeddingRun(doc, ed, doc.Database);
                return;
            }
            var target = _pending.Peek();

            if (!EnsurePlacementSpace(doc, target, ed))
            {
                _pending.Dequeue();
                _activePlacement = null;
                _activePasteDocument = null;
                ed?.WriteMessage("\nSkipping queued item because its target space could not be activated.");
                ProcessNextPaste(doc, ed);
                return;
            }

            bool clipboardReady;
            if (target.ClipBoundary != null)
            {
                clipboardReady = PrepareClipboardWithCroppedImage(target, ed);
            }
            else
            {
                clipboardReady = PrepareClipboardWithImageShared(target, ed);
            }

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

    }
}