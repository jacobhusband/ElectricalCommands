using Acies.AutoCAD.Shared;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;

namespace AutoCADCleanupTool
{
    public partial class CleanupCommands
    {
        [CommandMethod("SETTITLEBLOCK", CommandFlags.Modal)]
        [CommandMethod("STB", CommandFlags.Modal)]
        public static void SetProjectTitleBlockBoundary()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var db = doc.Database;

            if (!ProjectTitleBlockSettingsStore.TryPromptAndSave(
                doc,
                ed,
                db,
                out ProjectTitleBlockSettings settings,
                out string projectRoot,
                out string settingsPath))
            {
                return;
            }

            string sizeDescription = $"{settings.Width:0.###} x {settings.Height:0.###}";
            if (ProjectTitleBlockSettingsStore.TryClassifySheetSize(settings, out SheetSizeProfile profile, out _))
            {
                sizeDescription += $" ({profile.DisplayName})";
            }

            ed.WriteMessage(
                $"\nSaved project titleblock boundary {sizeDescription}." +
                $"\nProject root: {projectRoot}" +
                $"\nSettings file: {settingsPath}");
        }
    }
}
