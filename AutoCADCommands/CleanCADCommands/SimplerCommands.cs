using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Globalization;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using System.Diagnostics;

// Required for in-memory image rotation
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;


namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint BM_CLICK = 0x00F5;

        // *** MODIFICATION: Restored the missing method ***
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

        private class ImagePlacement
        {
            public string Path; // Path to the pre-rotated temporary image
            public Point3d Pos; // Original insertion point
            public Vector3d U;   // Original U vector (defines width and rotation)
            public Vector3d V;   // Original V vector (defines height)
            public ObjectId OriginalEntityId;
            public ObjectId TargetBtrId = ObjectId.Null;
            public Point2d[] ClipBoundary;
        }

        private static readonly Queue<ImagePlacement> _pending = new Queue<ImagePlacement>();
        private static ObjectId _lastPastedOle = ObjectId.Null;
        private static dynamic _pptAppShared = null;
        private static dynamic _pptPresentationShared = null;
        private static ObjectId _savedClayer = ObjectId.Null;
        private static readonly HashSet<ObjectId> _imageDefsToPurge = new HashSet<ObjectId>();
        internal static readonly HashSet<ObjectId> _pdfDefinitionsToDetach = new HashSet<ObjectId>();
        private static bool _chainFinalizeAfterEmbed = false;
        private static ImagePlacement _activePlacement = null;
        private static Document _activePasteDocument = null;
        private static bool _waitingForPasteStart = false;
        private static bool _pastePointHandlerAttached = false;
        private static bool _isEmbeddingProcessActive = false;
        private static bool _isCleanSheetWorkflowActive = false;
        internal static bool _skipLayerFreezing = false;
        private static ObjectId _finalPastedOleForZoom = ObjectId.Null;
        private static readonly HashSet<ObjectId> _xrefsToDetach = new HashSet<ObjectId>();

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
                                {
                                    _pptPresentationShared.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
                                }
                                return true;
                            }
                            catch
                            {
                                _pptPresentationShared = presentations.Add(MsoTriState.msoFalse);
                                _pptPresentationShared.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
                                return true;
                            }
                        }
                        else
                        {
                            _pptPresentationShared = presentations.Add(MsoTriState.msoFalse);
                            _pptPresentationShared.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
                            return true;
                        }
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
                    ed.WriteMessage("\nPowerPoint is not installed.");
                    return false;
                }
                _pptAppShared = Activator.CreateInstance(pptType);
                try { _pptAppShared.Visible = true; } catch { }
                var pres = _pptAppShared.Presentations;
                _pptPresentationShared = pres.Add(MsoTriState.msoFalse);
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
            if (_pptPresentationShared != null)
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
            if (_pptAppShared != null)
            {
                try { _pptAppShared.Quit(); } catch { }
                try { Marshal.FinalReleaseComObject(_pptAppShared); } catch { }
                _pptAppShared = null;
            }
        }

        /// <summary>
        /// Places the pre-rotated image onto the clipboard via PowerPoint.
        /// PowerPoint is used only as a converter to a high-quality OLE object.
        /// </summary>
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

                // Clear previous shapes from the slide
                for (int i = shapes.Count; i >= 1; i--)
                {
                    try { shapes[i].Delete(); } catch { }
                }

                // Add the pre-rotated picture and copy it. No rotation is done here.
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
                // Release COM objects
                if (pic != null) Marshal.ReleaseComObject(pic);
                if (shapes != null) Marshal.ReleaseComObject(shapes);
                if (slide != null) Marshal.ReleaseComObject(slide);
            }
        }

        private static string ResolveImagePath(Database db, string rawPath)
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
                        if (Path.IsPathRooted(p) && p.StartsWith("\\") && !p.StartsWith("\\\\"))
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
                    if (!string.IsNullOrWhiteSpace(found) && File.Exists(found)) return found;
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
                        {
                            keysToRemove.Add(entry.Key);
                        }
                    }

                    foreach (string key in keysToRemove)
                    {
                        try
                        {
                            imageDict.Remove(key);
                        }
                        catch { }
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
                    {
                        ed.WriteMessage($"\nPurged {purged} unused image definition(s).");
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

        private static void DetachPdfDefinitions(Database db, Editor ed)
        {
            if (_pdfDefinitionsToDetach.Count == 0) return;
            ed.WriteMessage($"\nAttempting to detach {_pdfDefinitionsToDetach.Count} PDF definition(s)...");
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var named = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                    if (!named.Contains("ACAD_PDFDEFINITIONS")) return;

                    var pdfDict = (DBDictionary)tr.GetObject(named.GetAt("ACAD_PDFDEFINITIONS"), OpenMode.ForWrite);
                    var entriesToRemove = new List<string>();

                    foreach (DBDictionaryEntry entry in pdfDict)
                    {
                        if (_pdfDefinitionsToDetach.Contains(entry.Value))
                        {
                            entriesToRemove.Add(entry.Key);
                            try
                            {
                                var defObj = tr.GetObject(entry.Value, OpenMode.ForWrite);
                                if (!defObj.IsErased) defObj.Erase();
                            }
                            catch { }
                        }
                    }

                    int detachedCount = 0;
                    foreach (string key in entriesToRemove)
                    {
                        try
                        {
                            var defId = pdfDict.GetAt(key);
                            pdfDict.Remove(key);
                            var defObj = tr.GetObject(defId, OpenMode.ForWrite);
                            defObj.Erase();
                            detachedCount++;
                        }
                        catch { }
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
            if (e.DBObject is Ole2Frame)
            {
                _lastPastedOle = e.DBObject.ObjectId;
            }
        }

        private static void Doc_CommandWillStart(object sender, CommandEventArgs e)
        {
            if (!string.Equals(e.GlobalCommandName, "PASTECLIP", StringComparison.OrdinalIgnoreCase)) return;
            if (!_waitingForPasteStart || _activePlacement == null || _pastePointHandlerAttached) return;
            _activePasteDocument ??= AutoCADApp.DocumentManager.MdiActiveDocument;
            AutoCADApp.Idle += Application_OnIdleSendPastePoint;
            _pastePointHandlerAttached = true;
        }

        private static bool EnsurePlacementSpace(Document doc, ImagePlacement placement, Editor ed)
        {
            if (placement.TargetBtrId.IsNull || doc.Database.CurrentSpaceId == placement.TargetBtrId)
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
                ed?.WriteMessage($"\nFailed to switch to target layout: {ex.Message}");
                return false;
            }
        }

        private static void Doc_CommandEnded(object sender, CommandEventArgs e)
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

            if (_pending.Count == 0 && _activePlacement == null) return;

            if (_lastPastedOle.IsNull)
            {
                ed.WriteMessage("\nSkipping an item because no OLE object was created.");
                if (_pending.Count > 0) _pending.Dequeue();
                _activePlacement = null;
                if (_pending.Count > 0) ProcessNextPaste(doc, ed);
                else FinishEmbeddingRun(doc, ed, db);
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
                if (ole != null)
                {
                    try
                    {
                        var oleExtents = ole.GeometricExtents;
                        Point3d oleMin = oleExtents.MinPoint;
                        double oleWidth = oleExtents.MaxPoint.X - oleMin.X;
                        double oleHeight = oleExtents.MaxPoint.Y - oleMin.Y;

                        if (oleWidth < 1e-8 || oleHeight < 1e-8)
                        {
                            ole.Erase();
                            tr.Commit();
                            return;
                        }

                        Point3d p0 = target.Pos;
                        Point3d p1 = target.Pos + target.U;
                        Point3d p2 = target.Pos + target.U + target.V;
                        Point3d p3 = target.Pos + target.V;

                        double minX = Math.Min(Math.Min(p0.X, p1.X), Math.Min(p2.X, p3.X));
                        double minY = Math.Min(Math.Min(p0.Y, p1.Y), Math.Min(p2.Y, p3.Y));
                        double maxX = Math.Max(Math.Max(p0.X, p1.X), Math.Max(p2.X, p3.X));
                        double maxY = Math.Max(Math.Max(p0.Y, p1.Y), Math.Max(p2.Y, p3.Y));

                        Point3d targetMin = new Point3d(minX, minY, p0.Z);
                        double targetWidth = maxX - minX;
                        double targetHeight = maxY - minY;

                        if (targetWidth < 1e-8 || targetHeight < 1e-8)
                        {
                            ole.Erase();
                            tr.Commit();
                            return;
                        }

                        double scaleX = targetWidth / oleWidth;
                        double scaleY = targetHeight / oleHeight;

                        Matrix3d translateToOrigin = Matrix3d.Displacement(oleMin.GetAsVector().Negate());
                        double[] scaleData = { scaleX, 0, 0, 0, 0, scaleY, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1 };
                        Matrix3d scaleMatrix = new Matrix3d(scaleData);
                        Matrix3d translateToTarget = Matrix3d.Displacement(targetMin.GetAsVector());
                        Matrix3d transformMatrix = translateToTarget * scaleMatrix * translateToOrigin;

                        ole.TransformBy(transformMatrix);

                        try { ole.Layer = "0"; } catch { }
                        var btr = (BlockTableRecord)tr.GetObject(ole.OwnerId, OpenMode.ForWrite);
                        var dot = (DrawOrderTable)tr.GetObject(btr.DrawOrderTableId, OpenMode.ForWrite);
                        dot.MoveAbove(new ObjectIdCollection { ole.ObjectId }, target.OriginalEntityId);
                    }
                    catch (System.Exception ex1)
                    {
                        ed.WriteMessage($"\nFailed to transform pasted OLE: {ex1.Message}");
                    }
                }

                var originalEnt = tr.GetObject(target.OriginalEntityId, OpenMode.ForWrite, false) as Entity;
                if (originalEnt != null)
                {
                    if (originalEnt is RasterImage imgEnt && !imgEnt.ImageDefId.IsNull)
                    {
                        _imageDefsToPurge.Add(imgEnt.ImageDefId);
                    }
                    originalEnt.Erase();
                }
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

        public static bool TryGetTitleBlockOutlinePointsForEmbed(Database db, out Point3d[] poly)
        {
            AutoCADApp.Idle -= FinalCleanupOnIdle;
            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null || !_isEmbeddingProcessActive) return;
            var db = doc.Database;
            var ed = doc.Editor;

            try
            {
                ed.WriteMessage("\nPerforming final cleanup for embedding process...");
                _isEmbeddingProcessActive = false;
                _finalPastedOleForZoom = ObjectId.Null;
                DetachHandlers(db, doc);
                DetachXrefs(db, ed);
                try { doc.SendStringToExecute("_.REGEN ", false, false, true); } catch { }
                PurgeEmbeddedImageDefs(db, ed);
                DetachPdfDefinitions(db, ed);
                WindowOrchestrator.EndPptInteraction();
                ClosePowerPoint(ed);
                try
                {
                    if (!_savedClayer.IsNull) db.Clayer = _savedClayer;
                }
                catch { }
                _savedClayer = ObjectId.Null;
                _skipLayerFreezing = false;

                if (_chainFinalizeAfterEmbed)
                {
                    _chainFinalizeAfterEmbed = false;
                    ed.WriteMessage("\nEMBEDFROMXREFS complete. Chaining FINALIZE and DETACHREMAININGXREFS...");
                    doc.SendStringToExecute("_.FINALIZE _.DETACHREMAININGXREFS ", true, false, false);
                }
                else if (_isCleanSheetWorkflowActive)
                {
                    _isCleanSheetWorkflowActive = false;
                    ed.WriteMessage("\nEMBEDFROMXREFS complete. Chaining next commands...");
                    doc.SendStringToExecute("_.EMBEDFROMPDFS _.CLEANPS _.VP2PL _.FINALIZE _.DETACHREMAININGXREFS _.ZOOMTOLASTTB ", true, false, false);
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nAn error occurred during final cleanup: {ex.Message}");
                _isCleanSheetWorkflowActive = false;
                _chainFinalizeAfterEmbed = false;
                _skipLayerFreezing = false;
            }
        }

        private static void DetachXrefs(Database db, Editor ed)
        {
            if (_xrefsToDetach.Count == 0) return;
            ed.WriteMessage($"\nAttempting to detach {_xrefsToDetach.Count} XREF(s)...");
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId xrefId in _xrefsToDetach)
                {
                    try
                    {
                        var btr = tr.GetObject(xrefId, OpenMode.ForRead) as BlockTableRecord;
                        if (btr != null && btr.IsFromExternalReference && btr.IsResolved)
                        {
                            db.DetachXref(xrefId);
                            ed.WriteMessage($"\nSuccessfully detached XREF: {btr.Name}");
                        }
                    }

                    if (best != null)
                    {
                        var ex = best.GeometricExtents;
                        poly = new[] {
                            new Point3d(ex.MinPoint.X, ex.MinPoint.Y, 0), new Point3d(ex.MaxPoint.X, ex.MinPoint.Y, 0),
                            new Point3d(ex.MaxPoint.X, ex.MaxPoint.Y, 0), new Point3d(ex.MinPoint.X, ex.MaxPoint.Y, 0)
                        };
                        return true;
                    }

                    poly = new[] {
                        new Point3d(db.Pextmin.X, db.Pextmin.Y, 0), new Point3d(db.Pextmax.X, db.Pextmin.Y, 0),
                        new Point3d(db.Pextmax.X, db.Pextmax.Y, 0), new Point3d(db.Pextmin.X, db.Pextmax.Y, 0)
                    };
                    return true;
                }
            }
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
                return "Model";
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
            string commandString = $"_.-LAYOUT S \"{targetLayoutName}\"\n_.PASTECLIP\n";
            ed.WriteMessage($"\nActivating layout '{targetLayoutName}' for pasting...");
            doc.SendStringToExecute(commandString, true, false, false);
        }

        // ====================================================================
        // NEW HELPER METHODS FOR IN-MEMORY IMAGE ROTATION
        // ====================================================================

        /// <summary>
        /// Rotates an image in memory around its center, creating a new bitmap
        /// large enough to contain the entire rotated image without clipping.
        /// </summary>
        /// <param name="image">The source image.</param>
        /// <param name="angle">The rotation angle in degrees.</param>
        /// <returns>A new, rotated Bitmap object.</returns>
        private static Bitmap RotateImage(System.Drawing.Image image, float angle)
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

                // Transformation:
                // 1. Move rotation point to the center of the new, larger bitmap.
                g.TranslateTransform(newWidth / 2f, newHeight / 2f);
                // 2. Rotate the coordinate system.
                g.RotateTransform(angle);
                // 3. Move the origin back by half the *original* image's size.
                g.TranslateTransform(-image.Width / 2f, -image.Height / 2f);
                // 4. Draw the original image at the transformed (0,0).
                g.DrawImage(image, new System.Drawing.Point(0, 0));
            }

            return rotatedBmp;
        }

        /// <summary>
        /// Prepares a raster image for embedding. It now handles pre-rotation of the
        /// image content in memory before saving it to a temporary file.
        /// </summary>
        /// <param name="srcPath">Path to the original image file.</param>
        /// <param name="rotationInDegrees">The rotation to apply.</param>
        /// <returns>Path to the new, temporary, pre-rotated PNG file.</returns>
        private static string PreflightRasterForPpt(string srcPath, double rotationInDegrees)
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "AutoCADCleanupTool", "embed");
            Directory.CreateDirectory(tempDir);
            string outPath = Path.Combine(
                tempDir,
                Path.GetFileNameWithoutExtension(srcPath) + "_ppt_" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".png"
            );

            using (var origImage = System.Drawing.Image.FromFile(srcPath))
            {
                try
                {
                    var fd = new FrameDimension(origImage.FrameDimensionsList[0]);
                    if (origImage.GetFrameCount(fd) > 1)
                        origImage.SelectActiveFrame(fd, 0);
                }
                catch { /* not multi-frame; ignore */ }

                System.Drawing.Image imageToProcess = origImage;
                Bitmap rotatedBitmap = null;

                // If rotation is significant, create a new rotated bitmap in memory.
                if (Math.Abs(rotationInDegrees) > 0.1)
                {
                    rotatedBitmap = RotateImage(origImage, (float)rotationInDegrees);
                    imageToProcess = rotatedBitmap;
                }

                // Resize the image (original or rotated) if it's too large.
                int maxSide = 8000;
                int w = imageToProcess.Width;
                int h = imageToProcess.Height;
                double scale = 1.0;
                if (Math.Max(w, h) > maxSide)
                    scale = (double)maxSide / Math.Max(w, h);

                int targetW = Math.Max(1, (int)Math.Round(w * scale));
                int targetH = Math.Max(1, (int)Math.Round(h * scale));

                // Save the final processed image to the temporary PNG file.
                using (var finalBmp = new Bitmap(targetW, targetH, PixelFormat.Format32bppArgb))
                using (var g = Graphics.FromImage(finalBmp))
                {
                    g.Clear(Color.Transparent);
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.DrawImage(imageToProcess, new Rectangle(0, 0, targetW, targetH));
                    finalBmp.Save(outPath, ImageFormat.Png);
                }

                // Clean up the intermediate rotated bitmap if it was created.
                rotatedBitmap?.Dispose();
            }

            return outPath;
        }
    }
}