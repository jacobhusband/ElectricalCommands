using System;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AcApplication = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    // =========================================================================
    // MASTER SWITCH COMMAND
    // =========================================================================
    [CommandMethod("SW", CommandFlags.Modal)]
    public static void PlaceSwitchMaster()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null) return;

      SwitchPlacementEngine.RunPlacementLoop(doc.Editor, doc.Database, SwitchType.Standard, SwitchOrientation.North);
    }

    // =========================================================================
    // STANDARD SWITCH COMMANDS (SWN, SWE, SWS, SWW)
    // =========================================================================
    [CommandMethod("SWN", CommandFlags.Modal)]
    public static void PlaceSwitchNorth()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null) return;
      SwitchPlacementEngine.RunPlacementLoop(doc.Editor, doc.Database, SwitchType.Standard, SwitchOrientation.North);
    }

    [CommandMethod("SWE", CommandFlags.Modal)]
    public static void PlaceSwitchEast()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null) return;
      SwitchPlacementEngine.RunPlacementLoop(doc.Editor, doc.Database, SwitchType.Standard, SwitchOrientation.East);
    }

    [CommandMethod("SWS", CommandFlags.Modal)]
    public static void PlaceSwitchSouth()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null) return;
      SwitchPlacementEngine.RunPlacementLoop(doc.Editor, doc.Database, SwitchType.Standard, SwitchOrientation.South);
    }

    [CommandMethod("SWW", CommandFlags.Modal)]
    public static void PlaceSwitchWest()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null) return;
      SwitchPlacementEngine.RunPlacementLoop(doc.Editor, doc.Database, SwitchType.Standard, SwitchOrientation.West);
    }

    // =========================================================================
    // DIMMER SWITCH COMMANDS (DSWN, DSWE, DSWS, DSWW / DMN, DME, DMS, DMW)
    // =========================================================================
    [CommandMethod("DSWN", CommandFlags.Modal)]
    public static void PlaceDimmerSwitchNorth()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null) return;
      SwitchPlacementEngine.RunPlacementLoop(doc.Editor, doc.Database, SwitchType.Dimmer, SwitchOrientation.North);
    }

    [CommandMethod("DMN", CommandFlags.Modal)]
    public static void PlaceDimmerSwitchNorthAlias() => PlaceDimmerSwitchNorth();

    [CommandMethod("DSWE", CommandFlags.Modal)]
    public static void PlaceDimmerSwitchEast()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null) return;
      SwitchPlacementEngine.RunPlacementLoop(doc.Editor, doc.Database, SwitchType.Dimmer, SwitchOrientation.East);
    }

    [CommandMethod("DME", CommandFlags.Modal)]
    public static void PlaceDimmerSwitchEastAlias() => PlaceDimmerSwitchEast();

    [CommandMethod("DSWS", CommandFlags.Modal)]
    public static void PlaceDimmerSwitchSouth()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null) return;
      SwitchPlacementEngine.RunPlacementLoop(doc.Editor, doc.Database, SwitchType.Dimmer, SwitchOrientation.South);
    }

    [CommandMethod("DMS", CommandFlags.Modal)]
    public static void PlaceDimmerSwitchSouthAlias() => PlaceDimmerSwitchSouth();

    [CommandMethod("DSWW", CommandFlags.Modal)]
    public static void PlaceDimmerSwitchWest()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null) return;
      SwitchPlacementEngine.RunPlacementLoop(doc.Editor, doc.Database, SwitchType.Dimmer, SwitchOrientation.West);
    }

    [CommandMethod("DMW", CommandFlags.Modal)]
    public static void PlaceDimmerSwitchWestAlias() => PlaceDimmerSwitchWest();

    // =========================================================================
    // OCCUPANCY SWITCH COMMANDS (OSWN, OSWE, OSWS, OSWW / OSN, OSE, OSS, OSW)
    // =========================================================================
    [CommandMethod("OSWN", CommandFlags.Modal)]
    public static void PlaceOccupancySwitchNorth()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null) return;
      SwitchPlacementEngine.RunPlacementLoop(doc.Editor, doc.Database, SwitchType.Occupancy, SwitchOrientation.North);
    }

    [CommandMethod("OSN", CommandFlags.Modal)]
    public static void PlaceOccupancySwitchNorthAlias() => PlaceOccupancySwitchNorth();

    [CommandMethod("OSWE", CommandFlags.Modal)]
    public static void PlaceOccupancySwitchEast()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null) return;
      SwitchPlacementEngine.RunPlacementLoop(doc.Editor, doc.Database, SwitchType.Occupancy, SwitchOrientation.East);
    }

    [CommandMethod("OSE", CommandFlags.Modal)]
    public static void PlaceOccupancySwitchEastAlias() => PlaceOccupancySwitchEast();

    [CommandMethod("OSWS", CommandFlags.Modal)]
    public static void PlaceOccupancySwitchSouth()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null) return;
      SwitchPlacementEngine.RunPlacementLoop(doc.Editor, doc.Database, SwitchType.Occupancy, SwitchOrientation.South);
    }

    [CommandMethod("OSS", CommandFlags.Modal)]
    public static void PlaceOccupancySwitchSouthAlias() => PlaceOccupancySwitchSouth();

    [CommandMethod("OSWW", CommandFlags.Modal)]
    public static void PlaceOccupancySwitchWest()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null) return;
      SwitchPlacementEngine.RunPlacementLoop(doc.Editor, doc.Database, SwitchType.Occupancy, SwitchOrientation.West);
    }

    [CommandMethod("OSW", CommandFlags.Modal)]
    public static void PlaceOccupancySwitchWestAlias() => PlaceOccupancySwitchWest();

    // =========================================================================
    // CONFIGURATION COMMANDS (SWSETUP, SWCONFIG, SWGUI)
    // =========================================================================
    [CommandMethod("SWSETUP", CommandFlags.Modal)]
    public static void RunSwitchSetup()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null) return;

      SwitchSetupWizard.RunSetup(doc.Editor, doc.Database);
    }

    [CommandMethod("SWCONFIG", CommandFlags.Modal)]
    public static void RunSwitchConfigAlias() => RunSwitchSetup();

    [CommandMethod("SWGUI", CommandFlags.Modal)]
    public static void OpenSwitchConfigGui()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null) return;

      try
      {
        var window = new SwitchConfigurationWindow();
        AcApplication.ShowModelessWindow(window);
      }
      catch (System.Exception ex)
      {
        doc.Editor.WriteMessage($"\nUnable to open Switch Configuration window: {ex.Message}");
      }
    }
  }

  public static class SwitchSetupWizard
  {
    public static void RunSetup(
      Editor ed,
      Database db,
      SwitchType? defaultType = null,
      SwitchOrientation? defaultOrient = null)
    {
      var projectSettings = SwitchConfigurationStore.LoadSettings(db);

      while (true)
      {
        PromptKeywordOptions typeOpts = new PromptKeywordOptions(
          "\nSwitch Setup: Select type to configure [Standard/Dimmer/Occupancy/Review/Defaults/Exit] <Standard>: ");
        typeOpts.Keywords.Add("Standard");
        typeOpts.Keywords.Add("Dimmer");
        typeOpts.Keywords.Add("Occupancy");
        typeOpts.Keywords.Add("Review");
        typeOpts.Keywords.Add("Defaults");
        typeOpts.Keywords.Add("Exit");
        typeOpts.Keywords.Default = defaultType.HasValue ? defaultType.Value.ToString() : "Standard";

        var typeRes = ed.GetKeywords(typeOpts);
        if (typeRes.Status != PromptStatus.OK || typeRes.StringResult == "Exit")
        {
          break;
        }

        if (typeRes.StringResult == "Review")
        {
          PrintSettingsSummary(ed, projectSettings);
          continue;
        }

        if (typeRes.StringResult == "Defaults")
        {
          HandleDefaultsMenu(ed, db, ref projectSettings);
          continue;
        }

        if (!Enum.TryParse<SwitchType>(typeRes.StringResult, true, out var selectedType))
        {
          continue;
        }

        var typeConfig = projectSettings.GetTypeConfig(selectedType);

        PromptKeywordOptions methodOpts = new PromptKeywordOptions(
          $"\nConfiguring [{typeConfig.DisplayName}]: [Auto-Derive-from-1-Sample / Pick-4-Orientations / Review / Back] <Auto-Derive-from-1-Sample>: ");
        methodOpts.Keywords.Add("Auto-Derive-from-1-Sample");
        methodOpts.Keywords.Add("Pick-4-Orientations");
        methodOpts.Keywords.Add("Review");
        methodOpts.Keywords.Add("Back");
        methodOpts.Keywords.Default = "Auto-Derive-from-1-Sample";

        var methodRes = ed.GetKeywords(methodOpts);
        if (methodRes.Status != PromptStatus.OK || methodRes.StringResult == "Back")
        {
          continue;
        }

        if (methodRes.StringResult == "Review")
        {
          PrintTypeSummary(ed, typeConfig);
          continue;
        }

        if (methodRes.StringResult == "Auto-Derive-from-1-Sample")
        {
          PromptKeywordOptions orientOpts = new PromptKeywordOptions(
            "\nWhich orientation is your sample switch? [North/East/South/West] <North>: ");
          orientOpts.Keywords.Add("North");
          orientOpts.Keywords.Add("East");
          orientOpts.Keywords.Add("South");
          orientOpts.Keywords.Add("West");
          orientOpts.Keywords.Default = defaultOrient.HasValue ? defaultOrient.Value.ToString() : "North";

          var orientRes = ed.GetKeywords(orientOpts);
          if (orientRes.Status != PromptStatus.OK) continue;

          Enum.TryParse<SwitchOrientation>(orientRes.StringResult, true, out var sampleOrient);

          ed.WriteMessage($"\n--- Sample Selection for {typeConfig.DisplayName} ({sampleOrient}) ---");
          if (SwitchPlacementEngine.CaptureFromSelection(ed, db, out var capturedConfig, out string captureErr))
          {
            SwitchConfigurationStore.AutoDeriveOrientations(
              capturedConfig,
              sampleOrient,
              out var north,
              out var east,
              out var south,
              out var west);

            typeConfig.North = north;
            typeConfig.East = east;
            typeConfig.South = south;
            typeConfig.West = west;

            if (SwitchConfigurationStore.SaveSettings(db, projectSettings, out string saveErr))
            {
              ed.WriteMessage($"\nSuccessfully auto-derived and saved all 4 orientations for {typeConfig.DisplayName} project-wide!");
              PrintTypeSummary(ed, typeConfig);
            }
            else
            {
              ed.WriteMessage($"\nError saving settings: {saveErr}");
            }
          }
          else
          {
            ed.WriteMessage($"\nCapture cancelled: {captureErr}");
          }
        }
        else if (methodRes.StringResult == "Pick-4-Orientations")
        {
          CaptureIndividualOrientations(ed, db, projectSettings, typeConfig);
        }
      }
    }

    private static void CaptureIndividualOrientations(
      Editor ed,
      Database db,
      ProjectSwitchSettings projectSettings,
      SwitchTypeConfig typeConfig)
    {
      SwitchOrientation[] orientations = { SwitchOrientation.North, SwitchOrientation.East, SwitchOrientation.South, SwitchOrientation.West };

      foreach (var orient in orientations)
      {
        PromptKeywordOptions opt = new PromptKeywordOptions(
          $"\nReady to capture sample for [{orient}]? [Capture/Skip/Cancel] <Capture>: ");
        opt.Keywords.Add("Capture");
        opt.Keywords.Add("Skip");
        opt.Keywords.Add("Cancel");
        opt.Keywords.Default = "Capture";

        var res = ed.GetKeywords(opt);
        if (res.Status != PromptStatus.OK || res.StringResult == "Cancel")
        {
          break;
        }

        if (res.StringResult == "Skip")
        {
          continue;
        }

        if (SwitchPlacementEngine.CaptureFromSelection(ed, db, out var config, out string err))
        {
          typeConfig.SetOrientation(orient, config);
          SwitchConfigurationStore.SaveSettings(db, projectSettings);
          ed.WriteMessage($"\nConfigured {orient} for {typeConfig.DisplayName}.");
        }
        else
        {
          ed.WriteMessage($"\nSkipped {orient}: {err}");
        }
      }

      ed.WriteMessage($"\nUpdated {typeConfig.DisplayName} configurations.");
      PrintTypeSummary(ed, typeConfig);
    }

    private static void HandleDefaultsMenu(Editor ed, Database db, ref ProjectSwitchSettings currentSettings)
    {
      PromptKeywordOptions opt = new PromptKeywordOptions(
        "\nDefaults: [Load-Global-Defaults / Save-Current-As-Global / Back] <Load-Global-Defaults>: ");
      opt.Keywords.Add("Load-Global-Defaults");
      opt.Keywords.Add("Save-Current-As-Global");
      opt.Keywords.Add("Back");
      opt.Keywords.Default = "Load-Global-Defaults";

      var res = ed.GetKeywords(opt);
      if (res.Status != PromptStatus.OK || res.StringResult == "Back") return;

      if (res.StringResult == "Load-Global-Defaults")
      {
        var loaded = SwitchConfigurationStore.LoadGlobalDefaults();
        if (loaded != null)
        {
          currentSettings = loaded;
          SwitchConfigurationStore.SaveSettings(db, currentSettings, out string err);
          ed.WriteMessage("\nGlobal default switch configurations loaded and applied to project!");
          PrintSettingsSummary(ed, currentSettings);
        }
        else
        {
          ed.WriteMessage($"\nNo global defaults file found at {SwitchConfigurationStore.GetGlobalDefaultsPath()}.");
        }
      }
      else if (res.StringResult == "Save-Current-As-Global")
      {
        if (SwitchConfigurationStore.SaveGlobalDefaults(currentSettings, out string err))
        {
          ed.WriteMessage($"\nCurrent switch configurations saved as global default template at: {SwitchConfigurationStore.GetGlobalDefaultsPath()}");
        }
        else
        {
          ed.WriteMessage($"\nFailed to save global defaults: {err}");
        }
      }
    }

    private static void PrintSettingsSummary(Editor ed, ProjectSwitchSettings settings)
    {
      StringBuilder sb = new StringBuilder();
      sb.AppendLine("\n================= PROJECT SWITCH CONFIGURATION =================");
      if (!string.IsNullOrWhiteSpace(settings.SourceDrawing))
      {
        sb.AppendLine($"Source DWG: {settings.SourceDrawing}");
      }
      if (!string.IsNullOrWhiteSpace(settings.UpdatedUtc))
      {
        sb.AppendLine($"Last Updated: {settings.UpdatedUtc}");
      }

      foreach (var kvp in settings.Switches)
      {
        var tc = kvp.Value;
        sb.AppendLine($"\n[{tc.DisplayName}] (Has Config: {tc.HasAnyConfiguration})");
        AppendOrientationSummary(sb, "North", tc.North);
        AppendOrientationSummary(sb, "East", tc.East);
        AppendOrientationSummary(sb, "South", tc.South);
        AppendOrientationSummary(sb, "West", tc.West);
      }

      sb.AppendLine("\nQuick Commands:");
      sb.AppendLine("  Switch:    SWN, SWE, SWS, SWW (Hub: SW)");
      sb.AppendLine("  Dimmer:    DSWN, DSWE, DSWS, DSWW (Aliases: DMN, DME, DMS, DMW)");
      sb.AppendLine("  Occupancy: OSWN, OSWE, OSWS, OSWW (Aliases: OSN, OSE, OSS, OSW)");
      sb.AppendLine("  Setup:     SWSETUP / SWCONFIG, GUI: SWGUI");
      sb.AppendLine("================================================================");

      ed.WriteMessage(sb.ToString());
    }

    private static void PrintTypeSummary(Editor ed, SwitchTypeConfig tc)
    {
      StringBuilder sb = new StringBuilder();
      sb.AppendLine($"\n--- Status for {tc.DisplayName} ---");
      AppendOrientationSummary(sb, "North", tc.North);
      AppendOrientationSummary(sb, "East", tc.East);
      AppendOrientationSummary(sb, "South", tc.South);
      AppendOrientationSummary(sb, "West", tc.West);
      ed.WriteMessage(sb.ToString());
    }

    private static void AppendOrientationSummary(StringBuilder sb, string label, SwitchOrientationConfig cfg)
    {
      if (cfg == null || !cfg.IsConfigured)
      {
        sb.AppendLine($"  {label,-6}: (Not Configured)");
      }
      else
      {
        double rotDeg = cfg.Block.Rotation * 180.0 / Math.PI;
        int txtCount = cfg.TextObjects?.Count ?? 0;
        string txtDetails = "";
        if (txtCount > 0 && cfg.TextObjects[0] != null)
        {
          txtDetails = $", Text: \"{cfg.TextObjects[0].TextString}\" (offset: {cfg.TextObjects[0].RelativeOffset.X:0.##}, {cfg.TextObjects[0].RelativeOffset.Y:0.##})";
        }
        sb.AppendLine($"  {label,-6}: Block '{cfg.Block.BlockName}' (Rot: {rotDeg:0}°){txtDetails}");
      }
    }
  }
}
