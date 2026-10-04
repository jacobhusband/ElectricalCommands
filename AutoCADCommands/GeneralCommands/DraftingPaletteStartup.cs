using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using Autodesk.AutoCAD.Runtime;

[assembly: ExtensionApplication(typeof(ElectricalCommands.DraftingPaletteStartup))]

namespace ElectricalCommands
{
  /// <summary>
  /// Opens the Electrical Drafting palette when AutoCAD finishes starting. The
  /// GeneralCommands bundle already loads at startup, so this is all it takes.
  /// </summary>
  public sealed class DraftingPaletteStartup : IExtensionApplication
  {
    public void Initialize()
    {
      try
      {
        // The headless console (accoreconsole) has no palettes, and loading
        // the palette types there would fail.
        if (!string.Equals(
          Process.GetCurrentProcess().ProcessName,
          "acad",
          StringComparison.OrdinalIgnoreCase))
        {
          return;
        }

        // The checkbox at the bottom of the palette turns this off.
        if (!DraftingPaletteSettings.ReadShowAtStartup())
        {
          return;
        }

        ScheduleShow();
      }
      catch
      {
        // Never block the plugin from loading over a palette.
      }
    }

    public void Terminate()
    {
    }

    // Kept separate so the AutoCAD UI types are only touched inside Initialize's try.
    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void ScheduleShow()
    {
      // Initialize runs while AutoCAD is still starting up; Idle fires once
      // its window and command line are ready.
      Autodesk.AutoCAD.ApplicationServices.Application.Idle += ShowOnFirstIdle;
    }

    private static void ShowOnFirstIdle(object sender, EventArgs e)
    {
      Autodesk.AutoCAD.ApplicationServices.Application.Idle -= ShowOnFirstIdle;
      try
      {
        DraftingPalette.Show();
      }
      catch (System.Exception ex)
      {
        Autodesk.AutoCAD.ApplicationServices.Document document =
          Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        document?.Editor.WriteMessage(
          "\nThe Electrical Drafting palette could not open: " + ex.Message +
          " Run DRAFTPALETTE to try again.");
      }
    }
  }
}
