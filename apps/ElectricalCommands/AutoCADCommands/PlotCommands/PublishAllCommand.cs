using Acies.AutoCAD.Shared;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.IO;
using System.Linq;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace ElectricalCommands
{
    public partial class PlotCommands
    {
        private const string PdfDeviceName = "DWG To PDF.pc3";
        private const string PlotStyleName = "510-monochrome.ctb";

        private sealed class PublishLayoutInfo
        {
            internal ObjectId LayoutId { get; set; }
            internal string Name { get; set; }
            internal int TabOrder { get; set; }
        }

        [CommandMethod("PA", CommandFlags.Modal)]
        public static void PublishAllLayouts()
        {
            Document doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                if (!TryGetOrCreateTitleBlockSettings(doc, ed, db, out ProjectTitleBlockSettings titleBlockSettings))
                {
                    return;
                }

                if (!ProjectTitleBlockSettingsStore.TryClassifySheetSize(
                    titleBlockSettings,
                    out SheetSizeProfile profile,
                    out string sizeFailure))
                {
                    ed.WriteMessage($"\nPA: {sizeFailure}");
                    if (!PromptToResetTitleBlock(ed) ||
                        !ProjectTitleBlockSettingsStore.TryPromptAndSave(
                            doc,
                            ed,
                            db,
                            out titleBlockSettings,
                            out _,
                            out _))
                    {
                        return;
                    }

                    if (!ProjectTitleBlockSettingsStore.TryClassifySheetSize(
                        titleBlockSettings,
                        out profile,
                        out sizeFailure))
                    {
                        ed.WriteMessage($"\nPA cancelled: {sizeFailure}");
                        return;
                    }
                }

                string drawingPath = ResolveSavedDrawingPath(doc, db);
                if (string.IsNullOrWhiteSpace(drawingPath))
                {
                    ed.WriteMessage("\nPA: Save the active drawing before publishing.");
                    return;
                }

                List<PublishLayoutInfo> layouts = GetPaperSpaceLayouts(db);
                if (layouts.Count == 0)
                {
                    ed.WriteMessage("\nPA: The active drawing has no paper-space layouts to publish.");
                    return;
                }

                string outputPath = PromptForOutputPdf(ed, drawingPath);
                if (string.IsNullOrWhiteSpace(outputPath))
                {
                    ed.WriteMessage("\nPA cancelled.");
                    return;
                }

                string selectedMedia = ConfigureLayoutsForPdf(db, ed, layouts, profile);
                PublishLayoutsToPdf(drawingPath, outputPath, layouts);

                ed.WriteMessage(
                    $"\nPA: Published {layouts.Count} layout(s) to '{outputPath}'." +
                    $"\nSheet size: {profile.DisplayName}. Media: {selectedMedia}.");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage($"\nPA failed: {ex.Message}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nPA failed: {ex.Message}");
            }
        }

        private static bool TryGetOrCreateTitleBlockSettings(
            Document doc,
            Editor ed,
            Database db,
            out ProjectTitleBlockSettings settings)
        {
            if (ProjectTitleBlockSettingsStore.TryLoad(
                doc,
                db,
                out settings,
                out string projectRoot,
                out _,
                out _))
            {
                ed.WriteMessage($"\nPA: Using the saved titleblock boundary for project '{projectRoot}'.");
                return true;
            }

            var options = new PromptKeywordOptions(
                $"\nNo project titleblock boundary is saved. Run {ProjectTitleBlockSettingsStore.SetCommandName} now?")
            {
                AllowNone = true,
                AppendKeywordsToMessage = true
            };
            options.Keywords.Add("Yes");
            options.Keywords.Add("No");
            options.Keywords.Default = "Yes";

            PromptResult result = ed.GetKeywords(options);
            bool shouldSet =
                result.Status == PromptStatus.None ||
                (result.Status == PromptStatus.OK && string.Equals(result.StringResult, "Yes", StringComparison.OrdinalIgnoreCase));

            if (!shouldSet)
            {
                settings = null;
                ed.WriteMessage(
                    $"\nPA cancelled. Run {ProjectTitleBlockSettingsStore.SetCommandName} and then run PA again.");
                return false;
            }

            return ProjectTitleBlockSettingsStore.TryPromptAndSave(
                doc,
                ed,
                db,
                out settings,
                out _,
                out _);
        }

        private static bool PromptToResetTitleBlock(Editor ed)
        {
            var options = new PromptKeywordOptions(
                $"\nReset the project boundary with {ProjectTitleBlockSettingsStore.SetCommandName} now?")
            {
                AllowNone = true,
                AppendKeywordsToMessage = true
            };
            options.Keywords.Add("Yes");
            options.Keywords.Add("No");
            options.Keywords.Default = "Yes";

            PromptResult result = ed.GetKeywords(options);
            return result.Status == PromptStatus.None ||
                   (result.Status == PromptStatus.OK && string.Equals(result.StringResult, "Yes", StringComparison.OrdinalIgnoreCase));
        }

        private static string ResolveSavedDrawingPath(Document doc, Database db)
        {
            var candidates = new[] { doc?.Name, db?.Filename };
            foreach (string candidate in candidates)
            {
                if (string.IsNullOrWhiteSpace(candidate)) continue;
                try
                {
                    string path = Path.GetFullPath(candidate.Trim().Trim('"'));
                    if (!Path.IsPathRooted(path)) continue;
                    if (!string.Equals(Path.GetExtension(path), ".dwg", StringComparison.OrdinalIgnoreCase)) continue;
                    if (!Directory.Exists(Path.GetDirectoryName(path))) continue;
                    if (!File.Exists(path)) continue;
                    return path;
                }
                catch
                {
                    // Try the next path exposed by AutoCAD.
                }
            }

            return string.Empty;
        }

        private static List<PublishLayoutInfo> GetPaperSpaceLayouts(Database db)
        {
            var layouts = new List<PublishLayoutInfo>();
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var layoutDictionary = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                if (layoutDictionary == null) return layouts;

                foreach (DBDictionaryEntry entry in layoutDictionary)
                {
                    var layout = tr.GetObject(entry.Value, OpenMode.ForRead, false) as Layout;
                    if (layout == null || layout.ModelType) continue;

                    layouts.Add(new PublishLayoutInfo
                    {
                        LayoutId = entry.Value,
                        Name = layout.LayoutName,
                        TabOrder = layout.TabOrder
                    });
                }

                tr.Commit();
            }

            return layouts
                .OrderBy(layout => layout.TabOrder)
                .ThenBy(layout => layout.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string PromptForOutputPdf(Editor ed, string drawingPath)
        {
            string drawingFolder = Path.GetDirectoryName(drawingPath) ?? string.Empty;
            string drawingName = Path.GetFileNameWithoutExtension(drawingPath);

            var options = new PromptSaveFileOptions("\nSelect the multi-sheet PDF output file: ")
            {
                DialogCaption = "Publish All Layouts to PDF",
                DialogName = "PA",
                Filter = "PDF files (*.pdf)|*.pdf",
                FilterIndex = 0,
                InitialDirectory = drawingFolder,
                InitialFileName = drawingName + ".pdf",
                PreferCommandLine = false,
                DeriveInitialFilenameFromDrawingName = false,
                ForceOverwriteWarningForScriptsAndLisp = true
            };

            PromptFileNameResult result = ed.GetFileNameForSave(options);
            if (result.Status != PromptStatus.OK || string.IsNullOrWhiteSpace(result.StringResult))
            {
                return string.Empty;
            }

            string outputPath = result.StringResult;
            if (!string.Equals(Path.GetExtension(outputPath), ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                outputPath += ".pdf";
            }
            return outputPath;
        }

        private static string ConfigureLayoutsForPdf(
            Database db,
            Editor ed,
            IReadOnlyList<PublishLayoutInfo> layouts,
            SheetSizeProfile profile)
        {
            PlotSettingsValidator validator = PlotSettingsValidator.Current;
            string selectedMedia = string.Empty;
            bool warnedMissingStyle = false;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (PublishLayoutInfo layoutInfo in layouts)
                {
                    var layout = tr.GetObject(layoutInfo.LayoutId, OpenMode.ForWrite, false) as Layout;
                    if (layout == null) continue;

                    using (var settings = new PlotSettings(layout.ModelType))
                    {
                        settings.CopyFrom(layout);
                        string mediaName = ConfigurePdfPlotSettings(validator, settings, profile, out bool plotStyleApplied);
                        if (string.IsNullOrWhiteSpace(selectedMedia))
                        {
                            selectedMedia = mediaName;
                        }

                        if (!plotStyleApplied && !warnedMissingStyle)
                        {
                            ed.WriteMessage(
                                $"\n[Warning] Plot style '{PlotStyleName}' was not available; existing layout plot styles were preserved.");
                            warnedMissingStyle = true;
                        }

                        layout.CopyFrom(settings);
                    }
                }

                tr.Commit();
            }

            return selectedMedia;
        }

        private static string ConfigurePdfPlotSettings(
            PlotSettingsValidator validator,
            PlotSettings settings,
            SheetSizeProfile profile,
            out bool plotStyleApplied)
        {
            plotStyleApplied = false;
            string preferredMedia = GetExpectedCanonicalMediaName(profile);

            try
            {
                validator.SetPlotConfigurationName(settings, PdfDeviceName, preferredMedia);
            }
            catch
            {
                validator.SetPlotConfigurationName(settings, PdfDeviceName, null);
                validator.RefreshLists(settings);
                preferredMedia = FindPreferredMediaName(validator, settings, profile);
                if (!string.IsNullOrWhiteSpace(preferredMedia))
                {
                    validator.SetCanonicalMediaName(settings, preferredMedia);
                }
                else
                {
                    validator.SetClosestMediaName(
                        settings,
                        profile.ShortSideInches,
                        profile.LongSideInches,
                        PlotPaperUnit.Inches,
                        false);
                }
            }

            validator.SetPlotPaperUnits(settings, PlotPaperUnit.Inches);
            validator.SetPlotRotation(settings, PlotRotation.Degrees090);
            validator.SetPlotType(settings, Autodesk.AutoCAD.DatabaseServices.PlotType.Layout);
            validator.SetUseStandardScale(settings, true);
            validator.SetStdScaleType(settings, StdScaleType.StdScale1To1);
            validator.SetPlotOrigin(settings, new Autodesk.AutoCAD.Geometry.Point2d(0.0, 0.0));
            validator.SetPlotCentered(settings, false);

            settings.PlotPlotStyles = true;
            settings.PrintLineweights = true;
            settings.ScaleLineweights = false;

            try
            {
                validator.RefreshLists(settings);
                StringCollection styles = validator.GetPlotStyleSheetList();
                string matchingStyle = styles
                    .Cast<string>()
                    .FirstOrDefault(style => string.Equals(style, PlotStyleName, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrWhiteSpace(matchingStyle))
                {
                    validator.SetCurrentStyleSheet(settings, matchingStyle);
                    plotStyleApplied = true;
                }
            }
            catch
            {
                // Preserve the layout's existing style when the configured CTB is unavailable.
            }

            return settings.CanonicalMediaName;
        }

        private static string FindPreferredMediaName(
            PlotSettingsValidator validator,
            PlotSettings settings,
            SheetSizeProfile profile)
        {
            StringCollection mediaNames = validator.GetCanonicalMediaNameList(settings);
            string shortToken = profile.ShortSideInches.ToString("0.00", CultureInfo.InvariantCulture);
            string longToken = profile.LongSideInches.ToString("0.00", CultureInfo.InvariantCulture);

            string bestName = string.Empty;
            int bestScore = int.MinValue;
            for (int index = 0; index < mediaNames.Count; index++)
            {
                string canonical = mediaNames[index] ?? string.Empty;
                string local = string.Empty;
                try { local = validator.GetLocaleMediaName(settings, index) ?? string.Empty; } catch { }

                string combined = (canonical + " " + local).ToLowerInvariant();
                int score = 0;
                if (combined.Contains(shortToken.ToLowerInvariant()) && combined.Contains(longToken.ToLowerInvariant())) score += 200;
                if (combined.Contains("inch")) score += 60;
                if (combined.Contains("full") && combined.Contains("bleed")) score += 80;
                if (combined.Contains(profile.Key.ToLowerInvariant().Replace('_', ' '))) score += 20;

                if (score > bestScore)
                {
                    bestScore = score;
                    bestName = canonical;
                }
            }

            return bestScore >= 200 ? bestName : string.Empty;
        }

        private static string GetExpectedCanonicalMediaName(SheetSizeProfile profile)
        {
            switch (profile.Key)
            {
                case "ANSI_D":
                    return "ANSI_full_bleed_D_(22.00_x_34.00_Inches)";
                case "ARCH_D":
                    return "ARCH_full_bleed_D_(24.00_x_36.00_Inches)";
                case "ARCH_E1":
                    return "ARCH_full_bleed_E1_(30.00_x_42.00_Inches)";
                case "ARCH_E":
                    return "ARCH_full_bleed_E_(36.00_x_48.00_Inches)";
                default:
                    return string.Empty;
            }
        }

        private static void PublishLayoutsToPdf(
            string drawingPath,
            string outputPath,
            IReadOnlyList<PublishLayoutInfo> layouts)
        {
            string temporaryDsd = Path.Combine(
                Path.GetTempPath(),
                "acies-publish-" + Guid.NewGuid().ToString("N") + ".dsd");

            object previousBackgroundPlot = null;
            try
            {
                using (var entries = new DsdEntryCollection())
                {
                    foreach (PublishLayoutInfo layout in layouts)
                    {
                        using (var entry = new DsdEntry())
                        {
                            entry.DwgName = drawingPath;
                            entry.Layout = layout.Name;
                            entry.Title = Path.GetFileNameWithoutExtension(drawingPath) + " - " + layout.Name;
                            entry.Nps = string.Empty;
                            entry.NpsSourceDwg = string.Empty;
                            entries.Add(entry);
                        }
                    }

                    using (var dsd = new DsdData())
                    {
                        dsd.DestinationName = outputPath;
                        dsd.ProjectPath = Path.GetDirectoryName(outputPath) ?? string.Empty;
                        dsd.LogFilePath = Path.ChangeExtension(outputPath, ".publish.log");
                        dsd.SheetType = SheetType.MultiPdf;
                        dsd.SheetSetName = Path.GetFileNameWithoutExtension(outputPath);
                        dsd.IsSheetSet = true;
                        dsd.IsHomogeneous = true;
                        dsd.NoOfCopies = 1;
                        dsd.PromptForDwfName = false;
                        dsd.SetDsdEntryCollection(entries);
                        dsd.WriteDsd(temporaryDsd);
                    }
                }

                previousBackgroundPlot = AutoCADApp.GetSystemVariable("BACKGROUNDPLOT");
                AutoCADApp.SetSystemVariable("BACKGROUNDPLOT", 0);

                using (var publishData = new DsdData())
                {
                    publishData.ReadDsd(temporaryDsd);
                    using (PlotConfig pdfConfig = PlotConfigManager.SetCurrentConfig(PdfDeviceName))
                    {
                        AutoCADApp.Publisher.PublishExecute(publishData, pdfConfig);
                    }
                }
            }
            finally
            {
                if (previousBackgroundPlot != null)
                {
                    try { AutoCADApp.SetSystemVariable("BACKGROUNDPLOT", previousBackgroundPlot); } catch { }
                }

                if (File.Exists(temporaryDsd))
                {
                    try { File.Delete(temporaryDsd); } catch { }
                }
            }
        }
    }
}
