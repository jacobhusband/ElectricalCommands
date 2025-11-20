using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Globalization;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
// WinForms SendKeys for OLE Text Size dialog automation
using System.Windows.Forms;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        internal enum PlacementSource
        {
            Unknown = 0,
            Pdf = 1,
            Xref = 2
        }

        // Win32 helpers for OLE Text Size dialog
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint BM_CLICK = 0x00F5;
        private const uint WM_COMMAND = 0x0111;
        private const int IDOK = 1;

        // Shared embedding state
        private class ImagePlacement
        {
            public string Path;
            public Point3d Pos;
            public Vector3d U;
            public Vector3d V;
            public ObjectId OriginalEntityId;
            public ObjectId TargetBtrId = ObjectId.Null;
            public Point2d[] ClipBoundary;
            public PlacementSource Source = PlacementSource.Unknown;
        }

        private static readonly Queue<ImagePlacement> _pending = new Queue<ImagePlacement>();
        private static ObjectId _lastPastedOle = ObjectId.Null;

        private static bool _handlersAttached = false;
        private static dynamic _pptAppShared = null;
        private static dynamic _pptPresentationShared = null;
        private static ObjectId _savedClayer = ObjectId.Null;

        private static readonly HashSet<ObjectId> _imageDefsToPurge = new HashSet<ObjectId>();
        internal static readonly HashSet<ObjectId> _pdfDefinitionsToDetach = new HashSet<ObjectId>();
        private static readonly HashSet<ObjectId> _xrefsToDetach = new HashSet<ObjectId>();
        private static HashSet<ObjectId> _pdfsToDetach = new HashSet<ObjectId>();

        private static bool _chainFinalizeAfterEmbed = false;
        private static bool _isEmbeddingProcessActive = false;
        private static bool _isCleanSheetWorkflowActive = false;
        internal static bool _skipLayerFreezing = false;

        private static ImagePlacement _activePlacement = null;
        private static Document _activePasteDocument = null;
        private static bool _waitingForPasteStart = false;
        private static bool _pastePointHandlerAttached = false;
        private static ObjectId _finalPastedOleForZoom = ObjectId.Null;

        // ------------------------------------------------------------------------------------
        // Layer utility (shared)
        // ------------------------------------------------------------------------------------
        internal static void EnsureAllLayersVisibleAndUnlocked(Database db, Editor ed)
        {
            ed.WriteMessage("\nEnsuring all layers are visible and unlocked...");
            int unlocked = 0, thawed = 0, turnedOn = 0;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    ObjectId originalClayerId = db.Clayer;

                    if (lt.Has("0") && db.Clayer != lt["0"])
                        db.Clayer = lt["0"];

                    foreach (ObjectId layerId in lt)
                    {
                        if (layerId.IsErased) continue;
                        try
                        {
                            var ltr = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                            if (ltr.IsLocked) { ltr.IsLocked = false; unlocked++; }
                            if (ltr.IsFrozen) { ltr.IsFrozen = false; thawed++; }
                            if (ltr.IsOff) { ltr.IsOff = false; turnedOn++; }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nCould not modify layer {layerId}: {ex.Message}");
                        }
                    }

                    try
                    {
                        var originalLtr = (LayerTableRecord)tr.GetObject(originalClayerId, OpenMode.ForRead);
                        if (!originalLtr.IsErased && !originalLtr.IsFrozen)
                            db.Clayer = originalClayerId;
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
                        IntPtr dlg = FindWindow("#32770", "OLE Text Size");
                        if (dlg != IntPtr.Zero)
                        {
                            IntPtr ok = FindWindowEx(dlg, IntPtr.Zero, "Button", "OK");
                            if (ok != IntPtr.Zero)
                            {
                                SendMessage(ok, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                            }
                            else
                            {
                                SendMessage(dlg, WM_COMMAND, (IntPtr)IDOK, IntPtr.Zero);
                                SetForegroundWindow(dlg);
                                System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                            }
                            return;
                        }
                    }
                    catch { }
                    Thread.Sleep(150);
                }
            });
        }

        // ------------------------------------------------------------------------------------
        // PowerPoint lifecycle (shared)
        // ------------------------------------------------------------------------------------
        private static bool EnsurePowerPoint(Editor ed)
        {
            try
            {
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
                                if (_pptPresentationShared.Slides.Count < 1)
                                    _pptPresentationShared.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
                                return true;
                            }
                            catch
                            {
                                _pptPresentationShared = presentations.Add(MsoTriState.msoFalse);
                                _pptPresentationShared.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
                                return true;
                            }
                        }

                        _pptPresentationShared = presentations.Add(MsoTriState.msoFalse);
                        _pptPresentationShared.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
                        return true;
                    }
                    catch
                    {
                        try { Marshal.FinalReleaseComObject(_pptAppShared); } catch { }
                        _pptAppShared = null;
                        _pptPresentationShared = null;
                    }
                }

                Type pptType = Type.GetTypeFromProgID("PowerPoint.Application");
                if (pptType == null)
                {
                    ed.WriteMessage("\nPowerPoint is not installed (COM ProgID not found).");
                    return false;
                }

                _pptAppShared = Activator.CreateInstance(pptType);
                try { _pptAppShared.Visible = true; } catch { }

                var presNew = _pptAppShared.Presentations;
                _pptPresentationShared = presNew.Add(MsoTriState.msoFalse);
                _pptPresentationShared.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
                return true;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to start or access PowerPoint: {ex.Message}");
                ClosePowerPoint(ed);
                return false;
            }
        }

        private static void ClosePowerPoint(Editor ed = null)
        {
            try
            {
                if (_pptPresentationShared != null)
                {
                    try { _pptPresentationShared.Saved = MsoTriState.msoTrue; } catch { }
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

        // ------------------------------------------------------------------------------------
        // Clipboard via PPT (shared)
        // ------------------------------------------------------------------------------------
        private static bool PrepareClipboardWithImageShared(ImagePlacement placement, Editor ed)
        {
            dynamic slide = null;
            dynamic shapes = null;
            Shape pic = null;

            try
            {
                if (!EnsurePowerPoint(ed)) return false;

                slide = _pptPresentationShared.Slides[1];
                shapes = slide.Shapes;

                for (int i = shapes.Count; i >= 1; i--)
                {
                    try { shapes[i].Delete(); } catch { }
                }

                pic = shapes.AddPicture(placement.Path, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10);
                pic.Copy();
                return true;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nFailed to prepare clipboard for '{placement.Path}': {ex.Message}");
                return false;
            }
            finally
            {
                if (pic != null) Marshal.ReleaseComObject(pic);
                if (shapes != null) Marshal.ReleaseComObject(shapes);
                if (slide != null) Marshal.ReleaseComObject(slide);
            }
        }

        // ------------------------------------------------------------------------------------
        // Path resolution (shared)
        // ------------------------------------------------------------------------------------
        internal static string ResolveImagePath(Database db, string rawPath)
        {
            if (string.IsNullOrWhiteSpace(rawPath)) return null;
            string p = rawPath.Trim().Trim('"').Replace('/', '\\');
            if (Path.IsPathRooted(p) && File.Exists(p)) return p;

            try
            {
                string dbFile = db?.Filename;
                if (!string.IsNullOrWhiteSpace(dbFile))
                {
                    string baseDir = Path.GetDirectoryName(dbFile);
                    if (!string.IsNullOrWhiteSpace(baseDir))
                    {
                        if (Path.IsPathRooted(p) &&
                            p.StartsWith("\\") &&
                            !p.StartsWith("\\\\"))
                        {
                            string drive = Path.GetPathRoot(baseDir);
                            string combined = Path.Combine(drive ?? string.Empty, p.TrimStart('\\'));
                            string full = Path.GetFullPath(combined);
                            if (File.Exists(full)) return full;
                        }

                        string candidate = Path.GetFullPath(Path.Combine(baseDir, p));
                        if (File.Exists(candidate)) return candidate;
                    }
                }
            }
            catch { }

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

        // ------------------------------------------------------------------------------------
        // Shared event handlers (PDF pipeline)
        // ------------------------------------------------------------------------------------
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

        private static void Db_ObjectAppended(object sender, ObjectEventArgs e)
        {
            if (e.DBObject is Ole2Frame)
                _lastPastedOle = e.DBObject.ObjectId;
        }

        private static void Doc_CommandWillStart(object sender, CommandEventArgs e)
        {
            if (!string.Equals(e.GlobalCommandName, "PASTECLIP", StringComparison.OrdinalIgnoreCase))
                return;
            if (!_waitingForPasteStart || _activePlacement == null || _pastePointHandlerAttached)
                return;

            _activePasteDocument ??= AutoCADApp.DocumentManager.MdiActiveDocument;
            AutoCADApp.Idle += Application_OnIdleSendPastePoint;
            _pastePointHandlerAttached = true;
        }

        private static void Application_OnIdleSendPastePoint(object sender, EventArgs e)
        {
            try
            {
                AutoCADApp.Idle -= Application_OnIdleSendPastePoint;
                _pastePointHandlerAttached = false;

                var doc = _activePasteDocument ?? AutoCADApp.DocumentManager.MdiActiveDocument;
                var placement = _activePlacement;
                if (doc == null || placement == null) return;
                if (!string.Equals(doc.CommandInProgress, "PASTECLIP", StringComparison.OrdinalIgnoreCase)) return;

                string x = placement.Pos.X.ToString("G17", CultureInfo.InvariantCulture);
                string y = placement.Pos.Y.ToString("G17", CultureInfo.InvariantCulture);
                doc.SendStringToExecute($"{x},{y}\n", true, false, false);
            }
            finally
            {
                _waitingForPasteStart = false;
            }
        }

        private static void Doc_CommandEnded(object sender, CommandEventArgs e)
        {
            if (!_isEmbeddingProcessActive ||
                !string.Equals(e.GlobalCommandName, "PASTECLIP", StringComparison.OrdinalIgnoreCase))
                return;

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
                return;

            if (_lastPastedOle.IsNull)
            {
                ed.WriteMessage("\nSkipping an item because no OLE object was created.");
                if (_pending.Count > 0) _pending.Dequeue();
                _activePlacement = null;

                if (_pending.Count > 0)
                    ProcessNextPaste(doc, ed);
                else
                    FinishEmbeddingRun(doc, ed, db);

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

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ole = tr.GetObject(_lastPastedOle, OpenMode.ForWrite) as Ole2Frame;
                if (ole != null && target.Source == PlacementSource.Pdf)
                {
                    try
                    {
                        ApplyPdfPlacementTransform(tr, ed, ole, target);
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nFailed to transform pasted OLE (PDF): {ex.Message}");
                    }
                }

                try
                {
                    var originalEnt = tr.GetObject(target.OriginalEntityId, OpenMode.ForWrite, false) as Entity;
                    if (originalEnt != null)
                    {
                        if (originalEnt is RasterImage imgEnt && !imgEnt.ImageDefId.IsNull)
                            _imageDefsToPurge.Add(imgEnt.ImageDefId);

                        originalEnt.Erase();
                    }
                }
                catch { }

                tr.Commit();
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            Thread.Sleep(50);

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

        // ------------------------------------------------------------------------------------
        // CORRECTED: PDF-specific placement transform
        // Maps the OLE's current extents exactly into the underlay's world-space rectangle.
        // This fixes scaling so embedded images match the original PDF underlay size.
        // ------------------------------------------------------------------------------------
        private static void ApplyPdfPlacementTransform(
            Transaction tr,
            Editor ed,
            Ole2Frame ole,
            ImagePlacement target)
        {
            // 1) Measure pasted OLE extents (source rectangle in world coords)
            var ext = ole.GeometricExtents;
            Point3d srcMin = ext.MinPoint;
            Point3d srcMax = ext.MaxPoint;
            double srcWidth = srcMax.X - srcMin.X;
            double srcHeight = srcMax.Y - srcMin.Y;

            if (srcWidth < 1e-8 || srcHeight < 1e-8)
            {
                ed.WriteMessage("\nSkipping PDF transform: invalid OLE extents.");
                ole.Erase();
                return;
            }

            // 2) Compute desired target rectangle from placement vectors
            // target.Pos = origin; U/V are world vectors representing width/height directions.
            Point3d t0 = target.Pos;
            Point3d t1 = target.Pos + target.U;
            Point3d t3 = target.Pos + target.V;

            Vector3d destU = t1 - t0;
            Vector3d destV = t3 - t0;
            double destWidth = destU.Length;
            double destHeight = destV.Length;

            if (destWidth < 1e-8 || destHeight < 1e-8)
            {
                ed.WriteMessage("\nSkipping PDF transform: invalid target geometry.");
                ole.Erase();
                return;
            }

            // Normalize destination axes
            Vector3d uDir = destU / destWidth;
            Vector3d vDir = destV / destHeight;
            Vector3d destNormal = uDir.CrossProduct(vDir);
            if (destNormal.Length < 1e-8)
                destNormal = Vector3d.ZAxis;

            // 3) Build 2D affine mapping:
            //    Map OLE's axis-aligned rectangle [srcMin, srcMax] to oriented rectangle at t0 with size destWidth/destHeight along uDir/vDir.
            //
            //    Steps:
            //    - Translate OLE so srcMin -> origin
            //    - Scale X,Y by (destWidth/srcWidth, destHeight/srcHeight)
            //    - Rotate/align X->uDir, Y->vDir
            //    - Translate origin -> t0

            // Translate to origin
            Matrix3d toOrigin = Matrix3d.Displacement(srcMin.GetAsVector().Negate());

            // Non-uniform scale to match sizes
            double sx = destWidth / srcWidth;
            double sy = destHeight / srcHeight;

            double[] scaleData =
            {
                sx, 0,  0, 0,
                0,  sy, 0, 0,
                0,  0,  1, 0,
                0,  0,  0, 1
            };
            Matrix3d scale = new Matrix3d(scaleData);

            // After scale: basis vectors
            Vector3d scaledX = new Vector3d(srcWidth * sx, 0, 0); // length destWidth
            Vector3d scaledY = new Vector3d(0, srcHeight * sy, 0); // length destHeight

            // Build a matrix that maps:
            //  - origin -> t0
            //  - scaledX -> destU
            //  - scaledY -> destV
            //
            // That is AlignCoordinateSystem from:
            //   origin, scaledX, scaledY, scaledX x scaledY
            // to:
            //   t0, destU, destV, destNormal

            Vector3d scaledNormal = scaledX.CrossProduct(scaledY);
            if (scaledNormal.Length < 1e-12)
                scaledNormal = destNormal;

            Matrix3d align = Matrix3d.AlignCoordinateSystem(
                Point3d.Origin,
                scaledX,
                scaledY,
                scaledNormal,
                t0,
                destU,
                destV,
                destNormal
            );

            // Apply: align * scale * toOrigin
            ole.TransformBy(align * scale * toOrigin);

            try { ole.Layer = "0"; } catch { }

            // Keep draw order relative to original underlay if possible
            try
            {
                if (!target.OriginalEntityId.IsNull)
                {
                    var origEnt = tr.GetObject(target.OriginalEntityId, OpenMode.ForRead, false) as Entity;
                    if (origEnt != null && !origEnt.IsErased)
                    {
                        var btr = tr.GetObject(ole.OwnerId, OpenMode.ForWrite) as BlockTableRecord;
                        if (btr != null &&
                            btr.DrawOrderTableId.IsValid &&
                            origEnt.OwnerId == btr.ObjectId)
                        {
                            var dot = tr.GetObject(btr.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;
                            if (dot != null)
                            {
                                var ids = new ObjectIdCollection(new[] { ole.ObjectId });
                                dot.MoveAbove(ids, origEnt.ObjectId);
                            }
                        }
                    }
                }
            }
            catch (System.Exception exOrder)
            {
                ed.WriteMessage($"\nWarning: Could not adjust draw order (PDF): {exOrder.Message}");
            }
        }

        // ------------------------------------------------------------------------------------
        // Final cleanup and helpers (shared)
        // ------------------------------------------------------------------------------------
        private static void FinalCleanupOnIdle(object sender, EventArgs e)
        {
            // Unsubscribe immediately to prevent re-entry
            AutoCADApp.Idle -= FinalCleanupOnIdle;

            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            bool shouldRun = _isEmbeddingProcessActive || _isCleanSheetWorkflowActive || _chainFinalizeAfterEmbed;
            if (doc == null || !shouldRun)
                return;

            // CRITICAL FIX: You must lock the document when modifying the Database from an Idle event.
            using (doc.LockDocument())
            {
                var db = doc.Database;
                var ed = doc.Editor;

                try
                {
                    ed.WriteMessage("\nPerforming final cleanup for embedding process...");

                    // 1. Reset flags
                    _isEmbeddingProcessActive = false;
                    _finalPastedOleForZoom = ObjectId.Null;

                    // 2. Detach event handlers
                    DetachHandlersForXrefs(db, doc); // Ensure this matches your handler detachment method name

                    // 3. Execute Database Write Operations (Safe now due to LockDocument)
                    DetachXrefs(db, ed);
                    PurgeEmbeddedImageDefs(db, ed);
                    DetachPdfDefinitions(db, ed);

                    // 4. Clean up external apps
                    WindowOrchestrator.EndPptInteraction();
                    ClosePowerPoint(ed);

                    // 5. Restore Layer
                    try
                    {
                        if (!_savedClayer.IsNull && _savedClayer.IsValid)
                            db.Clayer = _savedClayer;
                    }
                    catch { }
                    _savedClayer = ObjectId.Null;
                    _skipLayerFreezing = false;

                    // 6. Queue Visual Updates (Regen/Zoom)
                    // SendStringToExecute is async, so we queue it here.
                    try { doc.SendStringToExecute("_.REGEN ", false, false, true); } catch { }

                    // 7. Chain commands if needed
                    if (_chainFinalizeAfterEmbed)
                    {
                        _chainFinalizeAfterEmbed = false;
                        ed.WriteMessage("\nEMBEDIMAGES complete. Chaining FINALIZE and DETACHREMAININGXREFS...");
                        doc.SendStringToExecute("_.FINALIZE _.DETACHREMAININGXREFS ", true, false, false);
                    }

                    if (_isCleanSheetWorkflowActive)
                    {
                        _isCleanSheetWorkflowActive = false;
                        ed.WriteMessage("\nEMBEDIMAGES complete. Chaining next commands...");
                        doc.SendStringToExecute("_.EMBEDPDFS _.CLEANPS _.VP2PL _.FINALIZE _.DETACHREMAININGXREFS _.ZOOMTOLASTTB ", true, false, false);
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nAn error occurred during final cleanup: {ex.Message}");
                    // Ensure flags are reset even on error
                    _isCleanSheetWorkflowActive = false;
                    _chainFinalizeAfterEmbed = false;
                    _skipLayerFreezing = false;
                }
            }
        }

        private static void PurgeEmbeddedImageDefs(Database db, Editor ed)
        {
            if (_imageDefsToPurge.Count == 0) return;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var named = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                    if (!named.Contains("ACAD_IMAGE_DICT")) return;

                    var imageDict = (DBDictionary)tr.GetObject(named.GetAt("ACAD_IMAGE_DICT"), OpenMode.ForWrite);
                    var keysToRemove = new List<string>();

                    foreach (DBDictionaryEntry entry in imageDict)
                    {
                        if (_imageDefsToPurge.Contains(entry.Value))
                            keysToRemove.Add(entry.Key);
                    }

                    foreach (var key in keysToRemove)
                    {
                        try { imageDict.Remove(key); } catch { }
                    }

                    int purged = 0;
                    foreach (var defId in _imageDefsToPurge)
                    {
                        try
                        {
                            var defObj = tr.GetObject(defId, OpenMode.ForWrite, false);
                            if (defObj != null && !defObj.IsErased)
                            {
                                defObj.Erase();
                                purged++;
                            }
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

        /// <summary>
        /// Removes the PDF Definitions from the Named Objects Dictionary (ACAD_PDF_DEFINITIONS).
        /// This effectively "Detaches" the PDF from the External References palette.
        /// </summary>
        private static void DetachPdfDefinitions(Database db, Editor ed)
        {
            if (_pdfsToDetach == null || _pdfsToDetach.Count == 0)
                return;

            ed.WriteMessage($"\nDetaching {_pdfsToDetach.Count} PDF XREF(s)...");

            using (var tr = db.TransactionManager.StartTransaction())
            {
                DBDictionary pdfDict = null;

                try
                {
                    var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                    if (nod.Contains("ACAD_PDF_DEFINITIONS"))
                    {
                        pdfDict = (DBDictionary)tr.GetObject(nod.GetAt("ACAD_PDF_DEFINITIONS"), OpenMode.ForWrite);
                    }
                }
                catch { }

                int detached = 0;

                foreach (var id in _pdfsToDetach.ToList())
                {
                    try
                    {
                        // Remove dictionary entries that point at this definition
                        if (pdfDict != null)
                        {
                            var keys = new List<string>();
                            foreach (DBDictionaryEntry entry in pdfDict)
                            {
                                if (entry.Value == id)
                                    keys.Add(entry.Key);
                            }

                            foreach (var key in keys)
                            {
                                try { pdfDict.Remove(key); } catch { }
                            }
                        }

                        // Some AutoCAD versions treat PDF underlays like XREFs; try both detach paths.
                        try { db.DetachXref(id); } catch { }

                        try
                        {
                            var obj = tr.GetObject(id, OpenMode.ForWrite, false);
                            if (obj != null && !obj.IsErased)
                                obj.Erase();
                        }
                        catch { }

                        detached++;
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n[!] Error detaching PDF reference {id}: {ex.Message}");
                    }
                }

                tr.Commit();

                if (detached > 0)
                    ed.WriteMessage($"\n - Detached {detached} PDF reference(s).");
            }

            _pdfsToDetach.Clear();
        }

        private static void DetachXrefs(Database db, Editor ed)
        {
            // Ensure the list exists and has items
            if (_xrefsToDetach == null || _xrefsToDetach.Count == 0)
                return;

            // Use a HashSet to ensure we don't try to detach the same XREF twice 
            // (if multiple images lived in one XREF)
            var uniqueXrefs = new HashSet<ObjectId>(_xrefsToDetach);

            ed.WriteMessage($"\nProcessing {uniqueXrefs.Count} XREF(s) for detachment...");

            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var xrefId in uniqueXrefs)
                {
                    try
                    {
                        // 1. Check validity
                        if (xrefId.IsNull || xrefId.IsErased)
                            continue;

                        // 2. Get name for logging & verify it's actually an XREF
                        string xrefName = "Unknown";
                        try
                        {
                            var btr = (BlockTableRecord)tr.GetObject(xrefId, OpenMode.ForRead);
                            if (!btr.IsFromExternalReference)
                            {
                                ed.WriteMessage("\nSkipping detach: Object is not an external reference.");
                                continue;
                            }
                            xrefName = btr.Name;
                        }
                        catch
                        {
                            // If we can't read the BTR, we likely can't detach it safely
                            continue;
                        }

                        // 3. Perform the Detach
                        // Note: DetachXref is a Database method, not a Transaction method.
                        db.DetachXref(xrefId);
                        ed.WriteMessage($"\n - Successfully detached XREF: {xrefName}");
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n [!] Failed to detach XREF: {ex.Message}");
                    }
                }
                tr.Commit();
            }

            // Clear the list so we don't try to detach them again later
            _xrefsToDetach.Clear();
        }

        private static void FinishEmbeddingRun(Document doc, Editor ed, Database db)
        {
            AutoCADApp.Idle += FinalCleanupOnIdle;
        }

        private static string GetLayoutNameFromBtrId(Database db, ObjectId btrId)
        {
            if (btrId.IsNull) return "Model";

            using (var tr = db.TransactionManager.StartOpenCloseTransaction())
            {
                var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                if (btr != null && btr.IsLayout)
                {
                    var layout = tr.GetObject(btr.LayoutId, OpenMode.ForRead) as Layout;
                    return layout?.LayoutName ?? "Model";
                }
            }

            return "Model";
        }

        private static void ProcessNextPaste(Document doc, Editor ed)
        {
            if (!_isEmbeddingProcessActive || _pending.Count == 0)
            {
                if (_isEmbeddingProcessActive)
                    FinishEmbeddingRun(doc, ed, doc.Database);
                return;
            }

            var target = _pending.Peek();
            var db = doc.Database;

            string targetLayoutName = GetLayoutNameFromBtrId(db, target.TargetBtrId);
            if (string.IsNullOrEmpty(targetLayoutName))
            {
                _pending.Dequeue();
                ProcessNextPaste(doc, ed);
                return;
            }

            if (!PrepareClipboardWithImageShared(target, ed))
            {
                _pending.Dequeue();
                ProcessNextPaste(doc, ed);
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

            string cmd = $"_.-LAYOUT S \"{targetLayoutName}\"\n_.PASTECLIP\n";
            ed.WriteMessage($"\nActivating layout '{targetLayoutName}' for pasting...");
            doc.SendStringToExecute(cmd, true, false, false);
        }

        // ------------------------------------------------------------------------------------
        // Rotation + preflight helpers (used by XREF pipeline)
        // ------------------------------------------------------------------------------------
        internal static Bitmap RotateImage(System.Drawing.Image image, float angle)
        {
            if (image == null) throw new ArgumentNullException(nameof(image));

            const double toRad = Math.PI / 180.0;
            double angleRad = angle * toRad;
            double cos = Math.Abs(Math.Cos(angleRad));
            double sin = Math.Abs(Math.Sin(angleRad));

            int newWidth = (int)Math.Round(image.Width * cos + image.Height * sin);
            int newHeight = (int)Math.Round(image.Width * sin + image.Height * cos);

            var rotatedBmp = new Bitmap(newWidth, newHeight);
            rotatedBmp.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (Graphics g = Graphics.FromImage(rotatedBmp))
            {
                g.Clear(Color.Transparent);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;

                g.TranslateTransform(newWidth / 2f, newHeight / 2f);
                g.RotateTransform(angle);
                g.TranslateTransform(-image.Width / 2f, -image.Height / 2f);
                g.DrawImage(image, new System.Drawing.Point(0, 0));
            }

            return rotatedBmp;
        }

        internal static string PreflightRasterForPpt(string srcPath, double rotationInDegrees)
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "AutoCADCleanupTool", "embed");
            Directory.CreateDirectory(tempDir);

            string outPath = Path.Combine(
                tempDir,
                Path.GetFileNameWithoutExtension(srcPath) +
                "_ppt_" +
                Guid.NewGuid().ToString("N").Substring(0, 8) +
                ".png");

            using (var origImage = System.Drawing.Image.FromFile(srcPath))
            {
                try
                {
                    var fd = new FrameDimension(origImage.FrameDimensionsList[0]);
                    if (origImage.GetFrameCount(fd) > 1)
                        origImage.SelectActiveFrame(fd, 0);
                }
                catch { }

                System.Drawing.Image imageToProcess = origImage;
                Bitmap rotatedBitmap = null;

                if (Math.Abs(rotationInDegrees) > 0.1)
                {
                    rotatedBitmap = RotateImage(origImage, (float)rotationInDegrees);
                    imageToProcess = rotatedBitmap;
                }

                int maxSide = 8000;
                int w = imageToProcess.Width;
                int h = imageToProcess.Height;
                double scale = 1.0;

                if (Math.Max(w, h) > maxSide)
                    scale = (double)maxSide / Math.Max(w, h);

                int targetW = Math.Max(1, (int)Math.Round(w * scale));
                int targetH = Math.Max(1, (int)Math.Round(h * scale));

                using (var finalBmp = new Bitmap(targetW, targetH, PixelFormat.Format32bppArgb))
                using (var g = Graphics.FromImage(finalBmp))
                {
                    g.Clear(Color.Transparent);
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.DrawImage(imageToProcess, new Rectangle(0, 0, targetW, targetH));
                    finalBmp.Save(outPath, ImageFormat.Png);
                }

                rotatedBitmap?.Dispose();
            }

            return outPath;
        }
    }
}
