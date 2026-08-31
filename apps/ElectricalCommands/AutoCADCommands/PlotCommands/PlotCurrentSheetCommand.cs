using Acies.AutoCAD.Shared;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace ElectricalCommands
{
    public partial class PlotCommands
    {
        /// <summary>
        /// Plots the project titleblock window in the active layout. The boundary and
        /// paper size are shared by every drawing below the same project root.
        /// </summary>
        [CommandMethod("PLOTSHEET", CommandFlags.Modal)]
        [CommandMethod("PSHEET", CommandFlags.Modal)]
        public static void PlotCurrentSheet()
        {
            Document doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                if (!TryResolveCurrentSheetSettings(
                    doc,
                    ed,
                    db,
                    out ProjectTitleBlockSettings titleBlockSettings,
                    out SheetSizeProfile profile))
                {
                    return;
                }

                string drawingPath = ResolveSavedDrawingPath(doc, db);
                if (string.IsNullOrWhiteSpace(drawingPath))
                {
                    ed.WriteMessage("\nPLOTSHEET: Save the active drawing inside its project folder first.");
                    return;
                }

                string layoutName = LayoutManager.Current?.CurrentLayout ?? "Model";
                string outputPath = PromptForCurrentSheetOutput(ed, drawingPath, layoutName);
                if (string.IsNullOrWhiteSpace(outputPath))
                {
                    ed.WriteMessage("\nPLOTSHEET cancelled.");
                    return;
                }

                PlotCurrentLayoutWindow(doc, db, titleBlockSettings, profile, outputPath, out string mediaName);
                ed.WriteMessage(
                    $"\nPLOTSHEET: Plotted '{layoutName}' to '{outputPath}'." +
                    $"\nSheet size: {profile.DisplayName}. Media: {mediaName}.");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage($"\nPLOTSHEET failed: {ex.Message}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nPLOTSHEET failed: {ex.Message}");
            }
        }

        private static bool TryResolveCurrentSheetSettings(
            Document doc,
            Editor ed,
            Database db,
            out ProjectTitleBlockSettings settings,
            out SheetSizeProfile profile)
        {
            profile = null;
            if (ProjectTitleBlockSettingsStore.TryLoad(
                doc,
                db,
                out settings,
                out string projectRoot,
                out string settingsPath,
                out _))
            {
                if (ProjectTitleBlockSettingsStore.TryClassifySheetSize(settings, out profile, out _))
                {
                    ed.WriteMessage($"\nPLOTSHEET: Using the saved titleblock boundary for project '{projectRoot}'.");
                    return true;
                }

                if (ProjectTitleBlockSettingsStore.TryDetectSheetSizeFromProjectPdfs(
                    doc,
                    db,
                    out profile,
                    out string detectedPdf,
                    out string detectionFailure))
                {
                    if (!ProjectTitleBlockSettingsStore.TryApplyDetectedSheetSize(
                        settingsPath,
                        settings,
                        profile,
                        detectedPdf,
                        out string saveFailure))
                    {
                        ed.WriteMessage($"\n[Warning] PLOTSHEET could not cache the detected paper size: {saveFailure}");
                    }

                    ed.WriteMessage(
                        $"\nPLOTSHEET: Detected {profile.DisplayName} from '{detectedPdf}'." +
                        " The saved project boundary will be used for the plot window.");
                    return true;
                }

                ed.WriteMessage($"\nPLOTSHEET: {detectionFailure}");
                ed.WriteMessage("\nThe saved boundary does not identify a supported paper size. Select the outer titleblock corners again.");
            }
            else
            {
                ed.WriteMessage("\nPLOTSHEET: No project titleblock boundary is saved.");
            }

            SheetSizeProfile detectedProfile = null;
            string sourcePdf = string.Empty;
            if (ProjectTitleBlockSettingsStore.TryDetectSheetSizeFromProjectPdfs(
                doc,
                db,
                out detectedProfile,
                out sourcePdf,
                out string pdfFailure))
            {
                ed.WriteMessage(
                    $"\nPLOTSHEET: Detected {detectedProfile.DisplayName} from '{sourcePdf}'." +
                    " Select the titleblock corners to define the reusable plot window.");
            }
            else
            {
                ed.WriteMessage(
                    $"\nPLOTSHEET: {pdfFailure}" +
                    " Select the titleblock corners; the selected dimensions will determine the paper size.");
            }

            if (!ProjectTitleBlockSettingsStore.TryPromptAndSave(
                doc,
                ed,
                db,
                out settings,
                out projectRoot,
                out settingsPath,
                detectedProfile,
                sourcePdf))
            {
                return false;
            }

            if (!ProjectTitleBlockSettingsStore.TryClassifySheetSize(settings, out profile, out string sizeFailure))
            {
                ed.WriteMessage($"\nPLOTSHEET cancelled: {sizeFailure}");
                return false;
            }

            ed.WriteMessage(
                $"\nPLOTSHEET: Saved the project titleblock boundary at '{settingsPath}'." +
                " Future plots in this project will reuse it.");
            return true;
        }

        private static string PromptForCurrentSheetOutput(Editor ed, string drawingPath, string layoutName)
        {
            string drawingFolder = Path.GetDirectoryName(drawingPath) ?? string.Empty;
            string drawingName = Path.GetFileNameWithoutExtension(drawingPath);
            string safeLayoutName = string.Equals(layoutName, "Model", StringComparison.OrdinalIgnoreCase)
                ? string.Empty
                : " - " + SanitizeFileName(layoutName);

            var options = new PromptSaveFileOptions("\nSelect the current-sheet PDF output file: ")
            {
                DialogCaption = "Plot Current Sheet to PDF",
                DialogName = "PLOTSHEET",
                Filter = "PDF files (*.pdf)|*.pdf",
                FilterIndex = 0,
                InitialDirectory = drawingFolder,
                InitialFileName = drawingName + safeLayoutName + ".pdf",
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
            return string.Equals(Path.GetExtension(outputPath), ".pdf", StringComparison.OrdinalIgnoreCase)
                ? outputPath
                : outputPath + ".pdf";
        }

        private static string SanitizeFileName(string value)
        {
            string sanitized = value ?? string.Empty;
            foreach (char invalidCharacter in Path.GetInvalidFileNameChars())
            {
                sanitized = sanitized.Replace(invalidCharacter, '_');
            }
            return sanitized.Trim();
        }

        private static void PlotCurrentLayoutWindow(
            Document doc,
            Database db,
            ProjectTitleBlockSettings titleBlockSettings,
            SheetSizeProfile profile,
            string outputPath,
            out string mediaName)
        {
            if (PlotFactory.ProcessPlotState != ProcessPlotState.NotPlotting)
            {
                throw new InvalidOperationException("Another plot or publish operation is already in progress.");
            }

            ObjectId layoutId = LayoutManager.Current.GetLayoutId(LayoutManager.Current.CurrentLayout);
            Extents2d plotWindow = GetPlotWindowInDisplayCoordinates(doc.Editor, titleBlockSettings);
            PlotSettingsValidator validator = PlotSettingsValidator.Current;
            PlotSettings plotSettings;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var layout = tr.GetObject(layoutId, OpenMode.ForRead, false) as Layout;
                if (layout == null)
                {
                    throw new InvalidOperationException("The active layout could not be opened.");
                }

                plotSettings = new PlotSettings(layout.ModelType);
                plotSettings.CopyFrom(layout);
                tr.Commit();
            }

            using (plotSettings)
            {
                mediaName = ConfigurePdfPlotSettings(validator, plotSettings, profile, out _);
                validator.SetPlotType(plotSettings, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
                validator.SetPlotWindowArea(plotSettings, plotWindow);
                validator.SetUseStandardScale(plotSettings, true);
                validator.SetStdScaleType(plotSettings, StdScaleType.ScaleToFit);
                validator.SetPlotCentered(plotSettings, true);

                using (var plotInfo = new PlotInfo())
                {
                    plotInfo.Layout = layoutId;
                    plotInfo.OverrideSettings = plotSettings;

                    var infoValidator = new PlotInfoValidator
                    {
                        MediaMatchingPolicy = MatchingPolicy.MatchEnabled
                    };
                    infoValidator.Validate(plotInfo);

                    using (PlotEngine engine = PlotFactory.CreatePublishEngine())
                    using (var progress = new PlotProgressDialog(false, 1, true))
                    {
                        progress.set_PlotMsgString(PlotMessageIndex.DialogTitle, "ACIES Plot Current Sheet");
                        progress.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel plot");
                        progress.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel sheet");
                        progress.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Plot progress");
                        progress.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet progress");
                        progress.LowerPlotProgressRange = 0;
                        progress.UpperPlotProgressRange = 100;
                        progress.PlotProgressPos = 0;

                        progress.OnBeginPlot();
                        progress.IsVisible = true;
                        engine.BeginPlot(progress, null);
                        engine.BeginDocument(plotInfo, doc.Name, null, 1, true, outputPath);

                        using (var pageInfo = new PlotPageInfo())
                        {
                            progress.OnBeginSheet();
                            progress.LowerSheetProgressRange = 0;
                            progress.UpperSheetProgressRange = 100;
                            progress.SheetProgressPos = 0;

                            engine.BeginPage(pageInfo, plotInfo, true, null);
                            engine.BeginGenerateGraphics(null);
                            engine.EndGenerateGraphics(null);
                            engine.EndPage(null);

                            progress.SheetProgressPos = 100;
                            progress.OnEndSheet();
                        }

                        engine.EndDocument(null);
                        progress.PlotProgressPos = 100;
                        engine.EndPlot(null);
                        progress.OnEndPlot();
                    }
                }
            }
        }

        private static Extents2d GetPlotWindowInDisplayCoordinates(
            Editor ed,
            ProjectTitleBlockSettings titleBlockSettings)
        {
            Point3d lowerLeft = titleBlockSettings.LowerLeft.ToPoint3d();
            Point3d upperRight = titleBlockSettings.UpperRight.ToPoint3d();

            using (ViewTableRecord view = ed.GetCurrentView())
            {
                Matrix3d worldToDisplay = Matrix3d.PlaneToWorld(view.ViewDirection);
                worldToDisplay = Matrix3d.Displacement(view.Target - Point3d.Origin) * worldToDisplay;
                worldToDisplay = Matrix3d.Rotation(-view.ViewTwist, view.ViewDirection, view.Target) * worldToDisplay;
                worldToDisplay = worldToDisplay.Inverse();

                Point3d first = lowerLeft.TransformBy(worldToDisplay);
                Point3d second = upperRight.TransformBy(worldToDisplay);
                return new Extents2d(
                    Math.Min(first.X, second.X),
                    Math.Min(first.Y, second.Y),
                    Math.Max(first.X, second.X),
                    Math.Max(first.Y, second.Y));
            }
        }
    }
}
