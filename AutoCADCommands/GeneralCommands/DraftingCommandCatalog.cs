using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands
{
  /// <summary>
  /// One row on the drafting palette: a command, the other names it answers
  /// to, and what it does. A row with variants (for example N/E/S/W) shows one
  /// small button per variant instead of a single button.
  /// </summary>
  internal sealed class DraftingCommand
  {
    internal DraftingCommand(
      string title,
      string command,
      string description,
      string[] aliases,
      string[] variantLabels = null,
      string[] variantCommands = null)
    {
      Title = title;
      Command = command;
      Description = description;
      Aliases = aliases ?? new string[0];
      VariantLabels = variantLabels ?? new string[0];
      VariantCommands = variantCommands ?? new string[0];
    }

    internal string Title { get; }

    /// <summary>The command a single-button row runs.</summary>
    internal string Command { get; }

    internal string Description { get; }
    internal string[] Aliases { get; }
    internal string[] VariantLabels { get; }
    internal string[] VariantCommands { get; }
    internal bool HasVariants => VariantCommands.Length > 0;

    /// <summary>Every name this row answers to.</summary>
    internal IEnumerable<string> AllNames =>
      new[] { Command }
        .Concat(Aliases)
        .Concat(VariantCommands)
        .Where(name => !string.IsNullOrEmpty(name));

    /// <summary>The shortest name to type, shown beside the button.</summary>
    internal string TypedName =>
      HasVariants
        ? string.Join(" ", VariantCommands)
        : new[] { Command }.Concat(Aliases).OrderBy(name => name.Length).First();
  }

  internal sealed class DraftingCommandGroup
  {
    internal DraftingCommandGroup(
      string title,
      bool expandedByDefault,
      IReadOnlyList<DraftingCommand> commands)
    {
      Title = title;
      ExpandedByDefault = expandedByDefault;
      Commands = commands;
    }

    internal string Title { get; }
    internal bool ExpandedByDefault { get; }
    internal IReadOnlyList<DraftingCommand> Commands { get; }
  }

  /// <summary>
  /// Every command a drafter runs by hand, grouped for the drafting palette.
  /// When you add a command, add it here so it can be found and explained;
  /// anything registered but missing from this list (and from
  /// <see cref="NotShownCommands"/>) is picked up at runtime and shown under
  /// "Other Commands" instead.
  /// </summary>
  internal static class DraftingCommandCatalog
  {
    internal const string OtherCommandsTitle = "Other Commands";
    private const string NoDescription =
      "Not described yet. Add this command to DraftingCommandCatalog.cs so the palette can explain it.";

    private static DraftingCommand Cmd(
      string title,
      string command,
      string description,
      params string[] aliases)
    {
      return new DraftingCommand(title, command, description, aliases);
    }

    private static DraftingCommand Variants(
      string title,
      string description,
      string[] labels,
      string[] commands,
      params string[] aliases)
    {
      return new DraftingCommand(title, null, description, aliases, labels, commands);
    }

    private static readonly string[] Directions = { "N", "E", "S", "W" };

    internal static readonly IReadOnlyList<DraftingCommandGroup> Groups =
      new List<DraftingCommandGroup>
      {
        new DraftingCommandGroup("Place Symbols", true, new[]
        {
          Cmd("Receptacle", "R",
            "Places scaled receptacle blocks one after another: pick the location, then a point that sets the orientation. Press S to change the receptacle type and Enter to finish. Uses the drawing scale, or the active viewport's scale when you run it inside a viewport."),
          Cmd("Junction Box", "JBOX",
            "Places a scaled junction box symbol. Choose Plain, Switch, or Wallmount with the prompt keywords; for the last two, pick the side (north, east, south, or west) the switch or wall mount sits on."),
          Cmd("Data", "DATA",
            "Places a scaled cyan triangle data symbol, then lets you pick the direction it points (north, east, south, or west)."),
          Cmd("Disconnect", "DISCONNECT",
            "Places a scaled disconnect symbol, then lets you pick the side its handle faces."),
          Cmd("Switch", "SWITCH",
            "Places a scaled switch symbol: Standard ($), Dimmer ($D), or Occupancy sensor ($OS), chosen with the prompt keywords."),
          Cmd("Home Run", "HOMERUN",
            "Draws a home-run arrow with the panel label in four picks: arrow tip, bend, base, then text. Needs the panel name, panel location, and scale to be set first.",
            "HR"),
        }),

        new DraftingCommandGroup("Room Templates", true, new[]
        {
          Cmd("Capture Room Template", "RTCAPTURE",
            "Saves a room you have drawn as a reusable template. Name it, then select each group of symbols, keyed notes, and leaders, with a base point and the direction the group faces into the room. Circuit labels next to receptacles are captured too.",
            "RTC"),
          Cmd("Place Room Template", "RTPLACE",
            "Places a saved room template: choose a template, then position each group in turn with a live preview. Space or Enter rotates 90 degrees; the Wall, Skip, Back, and Done keywords align to a wall, leave a group out, redo the previous group, or finish. Needs the panel name and linked panel schedule when the template includes circuits.",
            "RTP"),
          Cmd("List Room Templates", "RTLIST",
            "Lists the room templates saved in this drawing and lets you delete one by name.",
            "RTD"),
        }),

        new DraftingCommandGroup("Circuiting", true, new[]
        {
          Cmd("Circuit Receptacles", "RC",
            "Circuits receptacles and writes the circuits into the linked panel schedule. Select the receptacles first, or run it and pick them. They are grouped by room (from Label Rooms) up to the RC Max kVA per circuit. Selecting a single receptacle opens the dedicated-equipment picker instead."),
          Cmd("Label Rooms", "AREALABEL",
            "Select room polylines and name them in a review window. Room names, areas, and locations are stored on the polylines so RC can group receptacles by room, and grouped totals are exported to AreaLabel.json.",
            "QA"),
          Cmd("Switch Circuits", "SWITCHCIRCUIT",
            "Opens a dockable palette that assigns non-repeating lowercase switch-control circuit letters to grouped fixture-tag text.",
            "SCIRCUIT"),
          Cmd("Set Spares", "SETSPARES",
            "Sets how many spare circuits are reserved in the linked panel schedule for the current panel.",
            "SPSPARES", "SETSPANELSPARES"),
        }),

        new DraftingCommandGroup("Annotation", true, new[]
        {
          Cmd("Key Note", "KN",
            "Places a keyed-note symbol in a paper-space layout, sized to the active viewport's scale. Enter the note number, then drag the symbol into place."),
          Cmd("Key Note (Same Number)", "KNV",
            "Like KN, but reuses the last note number so you can place several of the same note. Type V at the prompt to change the number."),
          Cmd("Key Note Leader", "KNL",
            "Places a keyed-note symbol with an attached leader: pick the leader tip, then drag the note into place. The leader ends in an arrow or an open circle (Arrow/Circle keywords)."),
          Cmd("Key Note Table", "KEYNOTETABLE",
            "Creates a compact borderless keyed-note table, or renumbers a selected table from 1 and adds any missing keyed-note symbols.",
            "KNTABLE", "KNT"),
          Cmd("General Note Table", "GENERALNOTETABLE",
            "Converts MText note columns or grouped text lines into a borderless numbered table, with overflow into extra columns controlled by three boundary points.",
            "GNTABLE", "GNT"),
          Cmd("Scaled Text", "SCALEDTEXT",
            "Places MText sized for the drawing scale (4.5\" at 1/4\" = 1'-0\") and lets you pick the direction it runs: north, east, south, or west. It always reads upright.",
            "TXT"),
          Cmd("Add Note", "ADDNOTE",
            "Pick a saved standard note from your note library, then click existing MText to append it, or click empty space to create new MText."),
          Cmd("Revision Cloud", "REV",
            "Draws a revision cloud and its delta tag on new layers, with a live preview."),
        }),

        new DraftingCommandGroup("Switch Assemblies", false, new[]
        {
          Cmd("Switch Assembly", "SW",
            "Places a configured switch assembly (symbol plus subscript text). At the prompt you can change the orientation (N/E/S/W), the type (Standard/Dimmer/Occupancy), the subscript, or open setup."),
          Cmd("Switch Setup", "SWSETUP",
            "Setup wizard: samples switch blocks and text from the drawing so SW can reproduce them, derives the orientations automatically, and saves the result project-wide.",
            "SWCONFIG"),
          Cmd("Switch Manager", "SWGUI",
            "Opens the Switch Configuration manager window."),
          Variants("Standard Switch",
            "Places a configured Standard switch facing north, east, south, or west (SWN, SWE, SWS, SWW).",
            Directions, new[] { "SWN", "SWE", "SWS", "SWW" }),
          Variants("Dimmer Switch",
            "Places a configured Dimmer switch facing north, east, south, or west (DSWN, DSWE, DSWS, DSWW; also DMN, DME, DMS, DMW).",
            Directions, new[] { "DSWN", "DSWE", "DSWS", "DSWW" },
            "DMN", "DME", "DMS", "DMW"),
          Variants("Occupancy Switch",
            "Places a configured Occupancy switch facing north, east, south, or west (OSWN, OSWE, OSWS, OSWW; also OSN, OSE, OSS, OSW).",
            Directions, new[] { "OSWN", "OSWE", "OSWS", "OSWW" },
            "OSN", "OSE", "OSS", "OSW"),
        }),

        new DraftingCommandGroup("Schedules & Lighting", false, new[]
        {
          Cmd("Lighting Fixture Schedule", "LFS",
            "Opens the lighting fixture schedule editor: copy from an existing schedule table, place a new one, and keep the linked table in sync.",
            "LFSOPEN"),
          Cmd("Control Schedule", "CSCHED",
            "Opens the Control Schedule editor to load, create, and update control schedule tables.",
            "CONTROLSCHEDULE"),
          Cmd("Light Plan Scan", "LIGHTPLANSCAN",
            "Scans room-boundary polylines and light fixture blocks and exports ACIESLightingPlan.snapshot.json next to the drawing.",
            "LPSCAN"),
          Cmd("Light Plan Apply", "LIGHTPLANAPPLY",
            "Applies reviewed lighting fixture tags from ACIESLightingPlan.instructions.json, updating tags it made earlier without duplicating them.",
            "LPAPPLY"),
        }),

        new DraftingCommandGroup("Text Tools", false, new[]
        {
          Cmd("Count Text", "TEXTCOUNT",
            "Counts selected TEXT, MTEXT, and attribute values, groups identical strings, reports the totals, and saves a timestamped text report."),
          Cmd("Select Text", "TEXTSELECT",
            "Filters a mixed selection so only TEXT and MTEXT objects stay selected."),
          Cmd("Sum Text", "SUMTEXT",
            "Totals selected room square-footage text, labels the total above the selection, and exports grouped totals to AreaLabel.json (Label Rooms first is recommended)."),
          Cmd("Replace Text", "TEXTREPLACE",
            "Replaces the contents of the selected text with a new value you enter.",
            "TN"),
          Cmd("Renumber Text", "TEXTINCREMENT",
            "Renumbers selected text using a prefix and a start/end range, with an odd/even filter.",
            "TI"),
          Cmd("Add To Text", "TEXTADD",
            "Adds an integer you specify to the numbers inside the selected text.",
            "TA"),
          Cmd("Sum Lengths", "SUMLENGTHS",
            "Sums the lengths of the selected lines, arcs, and polylines and reports the total."),
        }),

        new DraftingCommandGroup("Drawing Tools", false, new[]
        {
          Cmd("Object Mask", "OBJMASK",
            "Creates wipeout objects behind selected text, MText, tables, or polylines to mask the background.",
            "OM"),
          Cmd("Viewport From Region", "VPFROMREG",
            "Pick two corners of a model-space region to create a paper-space viewport for it, choosing the layout and an architectural scale.",
            "QVP"),
          Variants("Rotate",
            "Rotates the selected objects by 90, 180, 270, or 360 degrees around a base point you pick (R90, R180, R270, R360).",
            new[] { "90", "180", "270", "360" }, new[] { "R90", "R180", "R270", "R360" }),
          Cmd("Layer Red", "LAYRED",
            "Sets a selected layer, region entities, or XREF-related layers to red, depending on the mode you choose.",
            "LR"),
          Cmd("Layer Green", "LAYGREEN",
            "Sets every layer belonging to a selected XREF, and its nested XREFs, to green.",
            "LG"),
          Cmd("Freeze / Thaw Layers", "OPENLAYERS",
            "Freezes or thaws selected layers across the selected DWGs that are open in this AutoCAD session.",
            "OLAYERS", "OPENFREEZETHAW"),
          Cmd("Quick Xref", "QUICKXREF",
            "Pick a DWG from the project's Xrefs folder (found by searching up from the drawing) and attach it as an XREF at a point you pick. The drawing must be saved first.",
            "QXR"),
          Cmd("Project Checklist", "CHECKLIST",
            "Opens the project checklist that guides drafting and design steps. Progress is shared through a file in the drawing's project folder.",
            "CHECKLISTS", "PROJECTCHECKLIST", "PCL"),
        }),

        new DraftingCommandGroup("Sheet Prep & Plot", false, new[]
        {
          Cmd("Clean Sheet", "CLEANCAD",
            "Cleans the entire sheet: embeds XREF content and runs the cleanup steps."),
          Cmd("Clean Title Block", "CLEANTBLK",
            "Cleans the title block by exploding blocks, keeping only the title block, detaching XREFs, and embedding images."),
          Cmd("Bind All Xrefs", "BINDALLXREFS",
            "Removes every data link, PDF underlay, and image, then binds every DWG XREF whose file is found. XREFs with missing files stay attached and are reported.",
            "BAX"),
          Cmd("Embed Images", "EMBEDIMAGES",
            "Embeds raster images from XREFs into the drawing as OLE objects using PowerPoint, preserving orientation."),
          Cmd("Embed PDFs", "EMBEDPDFS",
            "Embeds PDF underlays by converting them to PNG and then to OLE objects using PowerPoint."),
          Cmd("Drawing Cleanup", "CLEANUP",
            "Runs SETBYLAYER on all objects, then PURGE ALL, then AUDIT with fixes to clean the current drawing."),
          Cmd("Set Title Block", "SETTITLEBLOCK",
            "Saves the project-wide title block boundary and sheet size that the plot commands use.",
            "STB"),
          Cmd("Pre-Flight QA/QC", "PREFLIGHT",
            "Opens the Pre-Flight QA/QC engine: a project scope questionnaire and drawing audit, with its state saved in the project folder. Save the drawing first.",
            "PFL", "QAQC", "PROJECTAUDIT", "SCOPECHECK", "DYNAMICCHECKLIST"),
          Cmd("Plot Sheet", "PLOTSHEET",
            "Plots the current sheet window to PDF using the saved title block boundary and the paper size detected from it.",
            "PSHEET"),
          Cmd("Publish All Layouts", "PA",
            "Publishes all layouts, in tab order, using the saved title block settings and the 510-monochrome plot style."),
          Variants("Plot To PDF",
            "Plots the current layout to PDF at 36x48, 30x42, 24x36, or 22x34 inches using DWG to PDF.pc3 (P36, P30, P24, P22).",
            new[] { "36", "30", "24", "22" }, new[] { "P36", "P30", "P24", "P22" }),
          Cmd("Plot And Move", "PLOTMOVE",
            "Pick a point, queue a 30x42 PDF plot of the current layout, and move every object in the current space so the picked point lands at the origin.",
            "QM"),
          Cmd("Insert T24 Form", "T24",
            "Inserts a T24 PDF form and automatically adds the embedded text and images on its final page."),
          Cmd("Insert PDF Pages", "PDFSHEETS",
            "Inserts every page of a selected multi-page PDF as an underlay, laid out in a grid.",
            "QP"),
          Cmd("Insert PNG Pages", "IMGSHEETS",
            "Inserts all PNG pages from a selected folder into paper space in a grid layout.",
            "QI"),
        }),
      };

    /// <summary>
    /// Commands the Drawing Settings rows already cover, so they are not
    /// listed a second time.
    /// </summary>
    private static readonly string[] SettingsCommands =
    {
      "SETPANELNAME", "SPN",
      "SETSCALE", "SS",
      "SETPANELLOCATION", "SPL",
      "SETPANELSCHEDULE", "SPS",
    };

    /// <summary>
    /// Registered commands that are deliberately not on the palette because
    /// they are not run by hand while drafting.
    /// </summary>
    private static readonly string[] NotShownCommands =
    {
      // The palette itself.
      "DRAFTPALETTE", "EDP", "HRSETTINGS", "HRS",

      // Steps CLEANCAD and FINALIZE queue for themselves.
      "FINALIZE", "-FINALIZE-CLEANUP", "-FINALIZE-PURGEDEFS",
      "DETACHREMAININGXREFS", "REMOVEREMAININGXREFS", "KEEPONLYTITLEBLOCKMS",
      "CLEANPS", "VP2PL", "ZOOMTOLASTTB",

      // Headless and batch entry points.
      "CLEANCAD2", "CLEANTBLK2", "ACIESCLEANJOB", "VP2PLHEADLESS",
      "-EXPORTSCHEDULES", "-SWITCHCIRCUITAPPLY",

      // Data export and developer helpers.
      "EXPORTSCHEDULES", "TEXTATTR",

      // Companions the schedule editors and RC run for themselves.
      "CSCHEDPLACE", "CSCHEDUPDATE", "CSCHEDLOAD", "-RCCOMPLETEDEDICATED",
    };

    private static HashSet<string> BuildKnownNames()
    {
      var known = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
      foreach (string name in SettingsCommands.Concat(NotShownCommands))
      {
        known.Add(name);
      }
      foreach (DraftingCommandGroup group in Groups)
      {
        foreach (DraftingCommand command in group.Commands)
        {
          foreach (string name in command.AllNames)
          {
            known.Add(name);
          }
        }
      }
      return known;
    }

    /// <summary>
    /// Finds commands registered by the ACIES plugins that the catalog does not
    /// mention, so a command added later still shows up on the palette.
    /// </summary>
    internal static DraftingCommandGroup DiscoverUnlistedCommands()
    {
      var unlisted = new List<DraftingCommand>();
      try
      {
        HashSet<string> known = BuildKnownNames();
        foreach (Assembly assembly in AppDomain.CurrentDomain.GetAssemblies())
        {
          string assemblyName = assembly.GetName().Name ?? string.Empty;
          if (!assemblyName.StartsWith("AutoCADCommands.", StringComparison.Ordinal))
          {
            continue;
          }

          foreach (MethodInfo method in GetMethods(assembly))
          {
            List<string> names;
            try
            {
              names = method
                .GetCustomAttributes(typeof(CommandMethodAttribute), false)
                .Cast<CommandMethodAttribute>()
                .Select(attribute => attribute.GlobalName)
                .Where(name => !string.IsNullOrEmpty(name))
                .ToList();
            }
            catch
            {
              continue;
            }

            if (names.Count == 0 || names.Any(known.Contains))
            {
              continue;
            }

            string primary = names[0];
            var aliases = names.Skip(1).ToArray();
            unlisted.Add(new DraftingCommand(primary, primary, NoDescription, aliases));
            foreach (string name in names)
            {
              known.Add(name);
            }
          }
        }
      }
      catch
      {
        // Discovery only adds to the palette; the catalog stands on its own.
      }

      return new DraftingCommandGroup(
        OtherCommandsTitle,
        false,
        unlisted.OrderBy(command => command.Title, StringComparer.OrdinalIgnoreCase).ToList());
    }

    private static IEnumerable<MethodInfo> GetMethods(Assembly assembly)
    {
      Type[] types;
      try
      {
        types = assembly.GetTypes();
      }
      catch (ReflectionTypeLoadException ex)
      {
        types = ex.Types.Where(type => type != null).ToArray();
      }
      catch
      {
        yield break;
      }

      const BindingFlags flags =
        BindingFlags.Public |
        BindingFlags.NonPublic |
        BindingFlags.Static |
        BindingFlags.Instance |
        BindingFlags.DeclaredOnly;
      foreach (Type type in types)
      {
        MethodInfo[] methods;
        try
        {
          methods = type.GetMethods(flags);
        }
        catch
        {
          continue;
        }
        foreach (MethodInfo method in methods)
        {
          yield return method;
        }
      }
    }
  }
}
