using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        private static bool IsApprovalMark(string value)
        {
            return Regex.IsMatch(value ?? "", @"stamp|signature|(?:^|[^a-z0-9])sig(?:$|[^a-z0-9])", RegexOptions.IgnoreCase);
        }

        private static object RemoveTitleblockMarks(Database db)
        {
            var removed = new List<object>();
            var detach = new List<ObjectId>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var table = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId blockId in table)
                {
                    var block = (BlockTableRecord)tr.GetObject(blockId, OpenMode.ForRead);
                    if (block.IsFromExternalReference || block.IsDependent)
                    {
                        if (!block.IsDependent && IsApprovalMark(block.Name + " " + Path.GetFileName(block.PathName))) detach.Add(blockId);
                        continue;
                    }
                    foreach (ObjectId id in block)
                    {
                        var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (entity == null) continue;
                        string source = "";
                        string markName = "";
                        if (entity is BlockReference reference)
                        {
                            var definition = (BlockTableRecord)tr.GetObject(reference.IsDynamicBlock ? reference.DynamicBlockTableRecord : reference.BlockTableRecord, OpenMode.ForRead);
                            source = definition.Name + " " + definition.PathName;
                            markName = definition.Name + " " + Path.GetFileName(definition.PathName);
                        }
                        else if (entity is RasterImage image && !(image is Wipeout))
                        {
                            source = ((RasterImageDef)tr.GetObject(image.ImageDefId, OpenMode.ForRead)).SourceFileName;
                            markName = Path.GetFileName(source);
                            if (Regex.IsMatch(source, @"(?:^|[\\/])signatures(?:[\\/]|$)", RegexOptions.IgnoreCase)) markName += " signature";
                        }
                        if (!IsApprovalMark(entity.Layer) && !IsApprovalMark(markName)) continue;
                        var layer = (LayerTableRecord)tr.GetObject(entity.LayerId, OpenMode.ForRead);
                        bool locked = layer.IsLocked;
                        if (locked) { layer.UpgradeOpen(); layer.IsLocked = false; }
                        removed.Add(new { handle = id.Handle.ToString(), kind = entity.GetType().Name, layer = entity.Layer, source });
                        entity.UpgradeOpen();
                        entity.Erase();
                        if (locked) layer.IsLocked = true;
                    }
                }
                tr.Commit();
            }
            foreach (var id in detach) db.DetachXref(id);
            return new { removed, detachedDefinitions = detach.Count,
                detection = "Named stamp/signature blocks, XREFs, image paths and layers. Unnamed exploded marks require visual review." };
        }

        private static object PlotCleanReview(CleanJob job)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            if (!(job.Width > 0 && job.Width <= 1000 && job.Height > 0 && job.Height <= 1000))
                throw new InvalidOperationException("Review plotting requires valid sheet dimensions.");
            if (Directory.Exists(job.Output)) throw new IOException("Plot output directory already exists.");
            Directory.CreateDirectory(job.Output);
            var layouts = new List<Tuple<int, string, ObjectId>>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (DBDictionaryEntry entry in (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead))
                {
                    var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    if (layout.ModelType) continue;
                    if (job.Layouts != null && !job.Layouts.Contains(layout.LayoutName)) continue;
                    var block = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);
                    if (job.Layouts == null && !block.Cast<ObjectId>().Any(id => !(tr.GetObject(id, OpenMode.ForRead) is Viewport))) continue;
                    layouts.Add(Tuple.Create(layout.TabOrder, layout.LayoutName, entry.Value));
                }
                tr.Commit();
            }
            if (layouts.Count == 0 || (job.Layouts != null && layouts.Count != job.Layouts.Length))
                throw new InvalidOperationException("Review layout list is empty or changed during cleanup.");
            var results = new List<object>();
            string originalLayout = LayoutManager.Current.CurrentLayout;
            object background = Application.GetSystemVariable("BACKGROUNDPLOT");
            object transparency = Application.GetSystemVariable("PLOTTRANSPARENCYOVERRIDE");
            object pdfFrame = Application.GetSystemVariable("PDFFRAME"), imageFrame = Application.GetSystemVariable("IMAGEFRAME");
            try
            {
                Application.SetSystemVariable("BACKGROUNDPLOT", 0);
                Application.SetSystemVariable("PLOTTRANSPARENCYOVERRIDE", 0);
                Application.SetSystemVariable("PDFFRAME", 0);
                Application.SetSystemVariable("IMAGEFRAME", 0);
                foreach (var item in layouts.OrderBy(x => x.Item1))
                {
                    LayoutManager.Current.CurrentLayout = item.Item2;
                    using (var settings = new PlotSettings(false))
                    {
                        using (var tr = db.TransactionManager.StartTransaction())
                        {
                            settings.CopyFrom((Layout)tr.GetObject(item.Item3, OpenMode.ForRead));
                            tr.Commit();
                        }
                        var validator = PlotSettingsValidator.Current;
                        validator.SetPlotConfigurationName(settings, "DWG To PDF.pc3", null);
                        validator.RefreshLists(settings);
                        validator.SetPlotPaperUnits(settings, PlotPaperUnit.Inches);
                        string size = string.Format(System.Globalization.CultureInfo.InvariantCulture, "({0:F2}_x_{1:F2}_Inches)", Math.Min(job.Width, job.Height), Math.Max(job.Width, job.Height));
                        string media = validator.GetCanonicalMediaNameList(settings).Cast<string>()
                            .Where(name => name.EndsWith(size, StringComparison.OrdinalIgnoreCase))
                            .OrderByDescending(name => name.IndexOf("full_bleed", StringComparison.OrdinalIgnoreCase) >= 0).FirstOrDefault();
                        if (media == null) throw new InvalidOperationException("PDF driver has no matching paper size: " + size);
                        validator.SetCanonicalMediaName(settings, media);
                        validator.SetPlotRotation(settings, job.Width >= job.Height ? PlotRotation.Degrees090 : PlotRotation.Degrees000);
                        validator.SetPlotType(settings, Autodesk.AutoCAD.DatabaseServices.PlotType.Layout);
                        validator.SetUseStandardScale(settings, true);
                        validator.SetStdScaleType(settings, StdScaleType.StdScale1To1);
                        // Layout plotting disables centering. SetPlotCentered itself
                        // rejects Layout mode, even when passed false.
                        validator.SetPlotOrigin(settings, Point2d.Origin);
                        const string plotStyle = "510-monochrome.ctb";
                        string style = validator.GetPlotStyleSheetList().Cast<string>()
                            .FirstOrDefault(name => string.Equals(name, plotStyle, StringComparison.OrdinalIgnoreCase));
                        if (style == null)
                            throw new InvalidOperationException("Required plot style is unavailable: " + plotStyle + ". Install it in this AutoCAD version's Plot Styles folder.");
                        validator.SetCurrentStyleSheet(settings, style);
                        settings.PlotPlotStyles = true;
                        settings.PrintLineweights = true;
                        settings.ScaleLineweights = false;
                        settings.DrawViewportsFirst = true;
                        settings.PlotTransparency = false;
                        settings.PlotHidden = false;
                        settings.ShadePlot = PlotSettingsShadePlotType.AsDisplayed;
                        settings.ShadePlotResLevel = ShadePlotResLevel.Normal;
                        string file = Path.Combine(job.Output, results.Count.ToString("D4") + ".pdf");
                        using (var info = new PlotInfo { Layout = item.Item3, OverrideSettings = settings })
                        using (var check = new PlotInfoValidator { MediaMatchingPolicy = MatchingPolicy.MatchEnabled })
                        {
                            check.Validate(info);
                            using (var engine = PlotFactory.CreatePublishEngine())
                            using (var page = new PlotPageInfo())
                            {
                                engine.BeginPlot(null, null);
                                engine.BeginDocument(info, doc.Name, null, 1, true, file);
                                engine.BeginPage(page, info, true, null);
                                engine.BeginGenerateGraphics(null);
                                engine.EndGenerateGraphics(null);
                                engine.EndPage(null);
                                engine.EndDocument(null);
                                engine.EndPlot(null);
                            }
                        }
                        if (!File.Exists(file) || new FileInfo(file).Length == 0) throw new IOException("Plot produced no PDF: " + item.Item2);
                        results.Add(new { layout = item.Item2, file, media = settings.CanonicalMediaName });
                    }
                }
            }
            finally
            {
                Application.SetSystemVariable("BACKGROUNDPLOT", background);
                Application.SetSystemVariable("PLOTTRANSPARENCYOVERRIDE", transparency);
                Application.SetSystemVariable("PDFFRAME", pdfFrame);
                Application.SetSystemVariable("IMAGEFRAME", imageFrame);
                LayoutManager.Current.CurrentLayout = originalLayout;
            }
            return new { pages = results, profile = "DWG To PDF.pc3; 510-monochrome.ctb; Layout; 1:1 inches; offset 0,0; full bleed matching confirmed sheet size; lineweights; paperspace last; transparency off; frames off" };
        }
    }
}
