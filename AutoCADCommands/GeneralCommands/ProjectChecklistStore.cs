using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ElectricalCommands
{
  internal static class ProjectChecklistStore
  {
    private const string AppDataFolderName = "ProjectManagementApp";
    private const string StateFileName = "checklist_state.json";
    private const string CustomChecklistsFileName = "checklists.json";

    internal static string ResolveAppDataFolder()
    {
      string userProfile = NormalizePlainText(Environment.GetEnvironmentVariable("USERPROFILE"));
      if (string.IsNullOrWhiteSpace(userProfile))
      {
        userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
      }
      if (string.IsNullOrWhiteSpace(userProfile))
      {
        userProfile = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
      }

      string appFolder = Path.Combine(userProfile, "Documents", AppDataFolderName);
      try
      {
        if (!Directory.Exists(appFolder))
        {
          Directory.CreateDirectory(appFolder);
        }
      }
      catch
      {
      }
      return appFolder;
    }

    internal static string RequireSavedDwgPath(Document doc, Database db)
    {
      string resolvedPath = ResolveBestDwgPath(doc, db);
      if (string.IsNullOrWhiteSpace(resolvedPath))
      {
        throw new InvalidOperationException(
          "Active drawing must be saved to a project folder before opening checklists."
        );
      }
      return resolvedPath;
    }

    internal static string ResolveBestDwgPath(Document doc, Database db)
    {
      var candidates = new List<string>();
      if (doc != null && !string.IsNullOrWhiteSpace(doc.Name))
      {
        candidates.Add(doc.Name);
      }
      if (db != null && !string.IsNullOrWhiteSpace(db.Filename))
      {
        candidates.Add(db.Filename);
      }

      return candidates
        .Select(NormalizePath)
        .Where(path => !string.IsNullOrWhiteSpace(path))
        .Where(path => string.Equals(Path.GetExtension(path), ".dwg", StringComparison.OrdinalIgnoreCase))
        .FirstOrDefault(path => !IsTemporaryPath(path))
        ?? candidates
          .Select(NormalizePath)
          .Where(path => !string.IsNullOrWhiteSpace(path))
          .FirstOrDefault(path => string.Equals(Path.GetExtension(path), ".dwg", StringComparison.OrdinalIgnoreCase))
        ?? string.Empty;
    }

    internal static string ResolveDrawingFolder(string dwgPath)
    {
      if (string.IsNullOrWhiteSpace(dwgPath))
      {
        return string.Empty;
      }

      try
      {
        return Path.GetDirectoryName(NormalizePath(dwgPath)) ?? string.Empty;
      }
      catch
      {
        return string.Empty;
      }
    }

    internal static string ResolveStateFilePath(string folderPath)
    {
      if (string.IsNullOrWhiteSpace(folderPath))
      {
        return string.Empty;
      }
      return Path.Combine(folderPath, StateFileName);
    }

    private const string GlobalLastActiveFileName = "last_active_checklist.txt";

    internal static string LoadGlobalLastActiveChecklist()
    {
      try
      {
        string filePath = Path.Combine(ResolveAppDataFolder(), GlobalLastActiveFileName);
        if (File.Exists(filePath))
        {
          string content = NormalizePlainText(File.ReadAllText(filePath, Encoding.UTF8));
          if (!string.IsNullOrWhiteSpace(content))
          {
            return content;
          }
        }
      }
      catch
      {
      }
      return string.Empty;
    }

    internal static void SaveGlobalLastActiveChecklist(string checklistId)
    {
      if (string.IsNullOrWhiteSpace(checklistId))
      {
        return;
      }

      try
      {
        string appFolder = ResolveAppDataFolder();
        if (!Directory.Exists(appFolder))
        {
          Directory.CreateDirectory(appFolder);
        }
        string filePath = Path.Combine(appFolder, GlobalLastActiveFileName);
        File.WriteAllText(filePath, NormalizePlainText(checklistId), Encoding.UTF8);
      }
      catch
      {
      }
    }

    internal static FolderChecklistState LoadFolderState(string folderPath)
    {
      string statePath = ResolveStateFilePath(folderPath);
      FolderChecklistState state = null;

      if (!string.IsNullOrWhiteSpace(statePath) && File.Exists(statePath))
      {
        try
        {
          string json = File.ReadAllText(statePath, Encoding.UTF8);
          state = JsonConvert.DeserializeObject<FolderChecklistState>(json);
        }
        catch
        {
        }
      }

      state = NormalizeFolderState(state ?? new FolderChecklistState());

      if (string.IsNullOrWhiteSpace(state.LastActiveChecklistId))
      {
        state.LastActiveChecklistId = LoadGlobalLastActiveChecklist();
      }

      return state;
    }

    internal static void SaveFolderState(string folderPath, FolderChecklistState state)
    {
      string statePath = ResolveStateFilePath(folderPath);
      if (string.IsNullOrWhiteSpace(statePath))
      {
        return;
      }

      try
      {
        if (!Directory.Exists(folderPath))
        {
          Directory.CreateDirectory(folderPath);
        }

        FolderChecklistState normalized = NormalizeFolderState(state);
        normalized.LastModifiedUtc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
        normalized.Version = Math.Max(1, normalized.Version + 1);

        if (!string.IsNullOrWhiteSpace(normalized.LastActiveChecklistId))
        {
          SaveGlobalLastActiveChecklist(normalized.LastActiveChecklistId);
        }

        string json = JsonConvert.SerializeObject(normalized, Formatting.Indented);
        string tempPath = statePath + ".tmp";
        File.WriteAllText(tempPath, json, Encoding.UTF8);

        if (File.Exists(statePath))
        {
          File.Delete(statePath);
        }
        File.Move(tempPath, statePath);
      }
      catch
      {
        try
        {
          // Fallback direct write if atomic replace failed
          string json = JsonConvert.SerializeObject(NormalizeFolderState(state), Formatting.Indented);
          File.WriteAllText(statePath, json, Encoding.UTF8);
        }
        catch
        {
        }
      }
    }

    internal static List<ChecklistDefinition> LoadAllChecklists(string folderPath)
    {
      var checklists = new List<ChecklistDefinition>(GetBuiltInChecklists());
      var seenIds = new HashSet<string>(checklists.Select(c => c.Id), StringComparer.OrdinalIgnoreCase);

      // Check for folder-level custom checklists (checklists.json or checklists_custom.json)
      var customCandidates = new List<string>();
      if (!string.IsNullOrWhiteSpace(folderPath))
      {
        customCandidates.Add(Path.Combine(folderPath, CustomChecklistsFileName));
        customCandidates.Add(Path.Combine(folderPath, "checklists_custom.json"));
      }
      customCandidates.Add(Path.Combine(ResolveAppDataFolder(), CustomChecklistsFileName));

      foreach (string customPath in customCandidates)
      {
        if (File.Exists(customPath))
        {
          try
          {
            string json = File.ReadAllText(customPath, Encoding.UTF8);
            var filePayload = JsonConvert.DeserializeObject<ChecklistDefinitionsFile>(json);
            foreach (var customDef in filePayload?.Checklists ?? new List<ChecklistDefinition>())
            {
              var normalized = NormalizeDefinition(customDef);
              if (!string.IsNullOrWhiteSpace(normalized.Id) && !seenIds.Contains(normalized.Id))
              {
                seenIds.Add(normalized.Id);
                checklists.Add(normalized);
              }
            }
          }
          catch
          {
          }
        }
      }

      return checklists;
    }

    public static List<ChecklistDefinition> GetBuiltInChecklists()
    {
      return new List<ChecklistDefinition>
      {
        CreatePreFlightElectricalChecklist(),
        CreateElectricalGeneralChecklist(),
        CreateAciesElectricalChecklist(),
        CreateAciesMechanicalChecklist(),
        CreateAciesPlumbingChecklist()
      };
    }

    private static ChecklistDefinition CreatePreFlightElectricalChecklist()
    {
      var def = new ChecklistDefinition
      {
        Id = "pre_flight_electrical",
        Name = "Pre-flight Electrical Checklist"
      };

      def.AddSubheader("General coordination");
      def.AddItem("Check relevant codes that apply to the project based on local city, state, and national codes. (CEC 90.2, 90.4)");
      def.AddItem("Check if the tenant space has existing mechanical units on the roof that are not powered from tenant electrical panels. (CEC 430.102, 440.14)");
      def.AddItem("Check for unmentioned items that could need power (interior signage). (CEC 600.6)");
      def.AddItem("Check for occ sensor on ceiling instead of wall in storage room. (Title 24, Part 6 Section 130.1(c)1)");
      def.AddItem("Check for dedicated service receptacle within 25ft of electrical panels. (CEC 210.63(B)(2), 110.26(E))");
      def.AddItem("Check for junction box indicated as wall mount for hand dryers. (CEC 314.23, 314.29)");
      def.AddItem("Check for adequate space for electrical panels, relocate to BOH corridors as necessary, avoid storage rooms, IT server racks. (CEC 110.26(A), 110.26(E))");
      def.AddItem("Check for food waste disposer for all sinks with outlet under counter and switch above (confirm with plumbing as it is not a requirement). (CEC 422.16(B)(1), 422.31(B))");
      def.AddItem("Check for furniture systems and include note to \"verify point of connection for furniture systems\". (CEC 605)");

      def.AddSubheader("Power, receptacles & equipment");
      def.AddItem("Check for controlled receptacles in office, lobby, kitchen, printer/copy room, conference room, meeting room. Modular furniture workstations need at least one controlled receptacle per workstation. (Title 24, Part 6 Section 130.5(d))");
      def.AddItem("Check for tamper proof receptacles in areas where children may be present: business offices, lobbies, waiting areas, theaters, auditoriums, gyms, bowling alleys, bus stations, airports, train stations. (CEC 406.12)");
      def.AddItem("Check for rooftop mechanical units shown on RCP, ensure they are dashed in appearance and noted to go on the roof along with rooftop receptacle. (CEC 210.63(A), 440.14)");
      def.AddItem("Check for return or supply air system over 2000CFM for duct smoke requirement. Any mechanical units 2000CFM or over should get duct smoke. (Title 24, Part 2 CBC Section 907.2.12.1.2)");
      def.AddItem("Check for hand dryer specification, if not add note to confirm exact breaker size with manufacturer. (CEC 110.3(B))");
      def.AddItem("Check for GFCI protection at required nonresidential receptacle locations (kitchens, outdoor, rooftops, within 6 ft of sinks, etc.). (CEC 210.8(B))");
      def.AddItem("Check meeting rooms for required receptacle outlets, including floor outlets when room size thresholds are met. (CEC 210.65)");

      def.AddSubheader("Distribution & single-line");
      def.AddItem("Check for kAIC rating shown on single line main switchboard. (CEC 110.9, 110.10)");
      def.AddItem("Check voltage drop requirements and calculations for feeders and branch circuits (<=5% combined). (Title 24, Part 6 Section 130.5(c))");
      def.AddItem("Check AIC ratings for panels and transformers on single line. (CEC 110.9, 110.10)");

      def.AddSubheader("Lighting controls & Title 24");
      def.AddItem("Check for >4000W in space needing demand response. Provide necessary software and device(s) to automatically reducing the lighting power by at least 15% upon receiving a demand response signal. (Title 24, Part 6 Section 110.12)");
      def.AddItem("Check for daylight harvesting in daylit zones. Provide daylight sensor & room controller for automatic dimming of light fixtures. (Title 24, Part 6 Section 130.1(d))");
      def.AddItem("Check for dimming in all rooms that are BOTH >100sqft AND >0.5W/sqft.");
      def.AddItem("Check for occupancy sensors which are required in offices, conference & meeting rooms, classrooms, restrooms, multipurpose rooms, warehouses, library aisles, corridors & stairwells.");
      def.AddItem("Check for time-switch controls which are allowed in lobbies, retail sales floors, commercial kitchens, auditoriums & theaters, and large multipurpose rooms.");
      def.AddItem("Check for 2-hour bypass switch in each room controlled by automatic time controlled on/off. Provide 1 bypass switch per 5000sqft of room size.");
      def.AddItem("Check for timeclock to indoor light fixtures that has manual override up to 2-hours, battery or internal memory capable of storing schedule for at least 7 days if power goes out.");
      def.AddItem("Check for astronomical timeclock, photosensor located on the roof, and contactors with signage, exterior / outdoor lights, pole lights, etc. (Section 130.2(c)2.B)");
      def.AddItem("Check for spaces that don't require multilevel controls such as: rooms under 100sqft, restrooms, rooms with only one luminaire.");

      return def;
    }

    private static ChecklistDefinition CreateElectricalGeneralChecklist()
    {
      var def = new ChecklistDefinition
      {
        Id = "electrical_general",
        Name = "Electrical General Checklist"
      };

      def.AddSubheader("General setup");
      def.AddItem("Reference Manager");
      def.AddItem("\"cleanup\" command");
      def.AddItem("XREF BGs");
      def.AddItem("Sheet layout & Titleblock");
      def.AddItem("Scope of work");
      def.AddItem("Specification sheet (CA or other state-specific)");
      def.AddItem("Sheet Index");

      def.AddSubheader("Distribution & major equipment");
      def.AddItem("MSB / panelboards / meter");
      def.AddItem("SLD");
      def.AddItem("Panel schedules");
      def.AddItem("EV chargers");
      def.AddItem("Solar");

      def.AddSubheader("Lighting plans & controls");
      def.AddItem("Arch RCP notes");
      def.AddItem("Title 24");
      def.AddItem("Light fixture schedule (check site photos of existing lights and compare to archived bank standard)");
      def.AddItem("Light control schedule");
      def.AddItem("Out-of-scope");
      def.AddItem("Daylight zones");
      def.AddItem("Timeclock");
      def.AddItem("EM lights");
      def.AddItem("Symbol list");
      def.AddItem("Light power, controls, notes");

      def.AddSubheader("Power plans & schedules");
      def.AddItem("Arch power notes");
      def.AddItem("Out-of-scope");
      def.AddItem("Equipment schedule");
      def.AddItem("IG/GFI/WP/Controlled/General power outlets");
      def.AddItem("Symbol list");

      def.AddSubheader("Site, roof & specialty systems");
      def.AddItem("Solar equipment");
      def.AddItem("Solar meter");
      def.AddItem("HVAC equipment");
      def.AddItem("Roof outlets");
      def.AddItem("Solar zone");
      def.AddItem("Solar stub");
      def.AddItem("Heaters");
      def.AddItem("Pumps");
      def.AddItem("Lights");
      def.AddItem("Signs");
      def.AddItem("EV chargers");
      def.AddItem("Equipment");
      def.AddItem("Photometric");

      return def;
    }

    private static ChecklistDefinition CreateAciesElectricalChecklist()
    {
      var def = new ChecklistDefinition
      {
        Id = "acies_electrical",
        Name = "Electrical - ACIES Checklist"
      };

      def.AddSubheader("Electrical General");
      def.AddItem("Latest state code requirement; year.");
      def.AddItem("Drawing schedule to match sheet numbers.");
      def.AddItem("Proper scale on drawings.");
      def.AddItem("Coordinate drawings presentation, sheet naming, sheet numbering, and scales with Mechanical & Plumbing for consistency.");
      def.AddItem("Lighting control diagram & specifications.");
      def.AddItem("Latest specification or applicable code.");
      def.AddItem("Receptacle to be 18\" top of box, and switches to be 48\" top of box mounting height.");
      def.AddItem("Utilize online T-24 forms (CA) or Comchecks (other states).");
      def.AddItem("Check adopted energy code and municipal code for specific requirements/amendments.");

      def.AddSubheader("Electrical Lighting");
      def.AddItem("Sequence of operation.");
      def.AddItem("Light fixture voltage / dimmable / 0-10V.");
      def.AddItem("Light fixture tag match lighting plan.");
      def.AddItem("Track current limiter rating at 80% loading.");
      def.AddItem("Location of override switch for exterior lighting controls.");
      def.AddItem("Check night light (NL).");
      def.AddItem("Separate lighting control for daylit area.");
      def.AddItem("Power pack for low-voltage sensor.");
      def.AddItem("Location of photocell or photo-sensor.");
      def.AddItem("Lighting under stairway & loading doors.");
      def.AddItem("Light switch (OS and manual on/off) location.");
      def.AddItem("Exterior light specification, UL listed for wet location. BUG rating.");

      def.AddSubheader("HVAC coord");
      def.AddItem("Provide motion sensor and light in attic space / near roof ladder opening, CMC 304.3.2.");
      def.AddItem("Toilet EF control. Interlock?");

      def.AddSubheader("T24 & Energy");
      def.AddItem("Dimming control for >=100SF with general lighting 0.5w/sf per T24.");
      def.AddItem("Demand response design for lighting power 4000W subject to dimming per T24 110.12(c) and controlled receptacles per T24 110.12(e).");

      def.AddSubheader("EM Lighting");
      def.AddItem("EM lighting (interior & exterior) per CBC 1008.");
      def.AddItem("EM lighting relays, UL924 to bypass lighting controls.");
      def.AddItem("EM light circuiting per 700.12.");
      def.AddItem("Notes on EM circuit ahead of switches.");
      def.AddItem("Sufficient EM lighting (1fc avg, 0.1fc min per CBC 1008.3.5).");
      def.AddItem("Exterior EM light above exit doors.");

      def.AddSubheader("Electrical Room & Service");
      def.AddItem("Clearance in front of panels and switchboard.");
      def.AddItem("Space for pull section at switchboard.");
      def.AddItem("2 exit doors for equipment 1200A or above and more than 6' wide; panic hardware.");
      def.AddItem("Provide receptacle within 15' of MSB.");
      def.AddItem("6\" wall for recessed panel.");
      def.AddItem("No ductwork in EL room.");
      def.AddItem("Solar ready - conduits to roof, solar PV space.");

      def.AddSubheader("Single Line & Power");
      def.AddItem("Single Line Diagram matches panel schedules.");
      def.AddItem("Bus sizing, main breaker sizing, and AIC ratings.");
      def.AddItem("Grounding electrode system per NEC 250 (ground rod, cold water, building steel).");
      def.AddItem("Transformer overcurrent protection & secondary protection.");
      def.AddItem("GFCI protection for counter / wet locations (6ft of sinks).");

      return def;
    }

    private static ChecklistDefinition CreateAciesMechanicalChecklist()
    {
      var def = new ChecklistDefinition
      {
        Id = "acies_mechanical",
        Name = "Mechanical - ACIES Checklist"
      };

      def.AddSubheader("Mechanical General");
      def.AddItem("Parapet height vs. RTU height + curb. Place all equipment at least 10 feet away from roof edge if parapet is < 42\".");
      def.AddItem("RTU voltage and refrigerant specification (R-32 or R-454B).");
      def.AddItem("10' liner from unit for noise attenuation.");
      def.AddItem("Check EF noise level.");
      def.AddItem("Demand control ventilation / CO2 sensor requirement.");
      def.AddItem("Economizer requirement & barometric relief.");
      def.AddItem("Smoke detector requirement (supply/return, >=2000 CFM).");
      def.AddItem("Drawing schedule to match sheet numbers.");
      def.AddItem("Ventilation rate per occupancy.");
      def.AddItem("GFI for RTU and disconnect switch coordination.");
      def.AddItem("Smart T-stat specifications (T24 compliant, demand response).");

      def.AddSubheader("Mechanical Floor Plan");
      def.AddItem("Ductwork available clearance.");
      def.AddItem("Cooling in elevator machine room outside EMR.");
      def.AddItem("Flue and combustion air for gas water heater.");
      def.AddItem("Fire wall protection @ shaft, electrical & mechanical room.");
      def.AddItem("Outside air provision for ground floor tenants.");
      def.AddItem("Toilet exhaust provision for ground floor tenants.");
      def.AddItem("T-stat location matching serving unit.");
      def.AddItem("Provide bug screen for all exterior grilles/louvers.");

      return def;
    }

    private static ChecklistDefinition CreateAciesPlumbingChecklist()
    {
      var def = new ChecklistDefinition
      {
        Id = "acies_plumbing",
        Name = "Plumbing - ACIES Checklist"
      };

      def.AddSubheader("Plumbing General");
      def.AddItem("General notes to be county or city specific.");
      def.AddItem("Check state code if seismic restraints are required for plumbing.");
      def.AddItem("Check if project site falls under IECC code.");
      def.AddItem("Natural gas piping verification for project jurisdiction.");
      def.AddItem("Coordinate EWH voltage and balanced load requirement.");
      def.AddItem("ADA compliance notes for fixtures and clearances.");
      def.AddItem("Drawing schedule to match sheet numbers.");
      def.AddItem("Pipe material specification.");
      def.AddItem("Access panel for clean out notes.");
      def.AddItem("Water heater sizing and GPH rise requirements.");
      def.AddItem("Fixture / faucets to comply with CalGreen and CPC flow rates.");

      def.AddSubheader("Plumbing Floor Plan");
      def.AddItem("Water heater location and seismic strapping.");
      def.AddItem("Floor drains and floor sinks trap primers.");
      def.AddItem("Backflow preventer location and drainage.");
      def.AddItem("Grease interceptor / trap sizing and location.");
      def.AddItem("Roof drain & overflow drain piping coordination.");

      return def;
    }

    internal static FolderChecklistState NormalizeFolderState(FolderChecklistState state)
    {
      var normalized = new FolderChecklistState
      {
        Version = state?.Version ?? 1,
        LastModifiedUtc = string.IsNullOrWhiteSpace(state?.LastModifiedUtc)
          ? DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)
          : state.LastModifiedUtc,
        LastActiveChecklistId = NormalizePlainText(state?.LastActiveChecklistId)
      };

      foreach (var kvp in state?.Checklists ?? new Dictionary<string, FolderChecklistEntry>())
      {
        string checklistId = NormalizePlainText(kvp.Key);
        if (string.IsNullOrWhiteSpace(checklistId))
        {
          continue;
        }

        var entry = kvp.Value ?? new FolderChecklistEntry();
        var normalizedEntry = new FolderChecklistEntry
        {
          LastModifiedUtc = string.IsNullOrWhiteSpace(entry.LastModifiedUtc)
            ? normalized.LastModifiedUtc
            : entry.LastModifiedUtc
        };

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string rawItem in entry.CompletedItems ?? new List<string>())
        {
          string itemKey = NormalizePlainText(rawItem);
          if (!string.IsNullOrWhiteSpace(itemKey) && !seen.Contains(itemKey))
          {
            seen.Add(itemKey);
            normalizedEntry.CompletedItems.Add(itemKey);
          }
        }

        foreach (var note in entry.ItemNotes ?? new Dictionary<string, string>())
        {
          string itemKey = NormalizePlainText(note.Key);
          if (!string.IsNullOrWhiteSpace(itemKey))
          {
            normalizedEntry.ItemNotes[itemKey] = note.Value ?? string.Empty;
          }
        }

        normalized.Checklists[checklistId] = normalizedEntry;
      }

      return normalized;
    }

    internal static ChecklistDefinition NormalizeDefinition(ChecklistDefinition rawDef)
    {
      var def = new ChecklistDefinition
      {
        Id = NormalizePlainText(rawDef?.Id),
        Name = NormalizePlainText(rawDef?.Name),
        Category = NormalizePlainText(rawDef?.Category)
      };

      if (string.IsNullOrWhiteSpace(def.Name))
      {
        def.Name = string.IsNullOrWhiteSpace(def.Id) ? "Untitled Checklist" : def.Id;
      }
      if (string.IsNullOrWhiteSpace(def.Id))
      {
        def.Id = Regex.Replace(def.Name.ToLowerInvariant(), @"[^a-z0-9_]+", "_").Trim('_');
      }

      int order = 0;
      foreach (var item in rawDef?.Items ?? new List<ChecklistItem>())
      {
        order++;
        string itemText = (item?.Text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(itemText))
        {
          continue;
        }

        string itemId = NormalizePlainText(item?.Id);
        if (string.IsNullOrWhiteSpace(itemId))
        {
          itemId = itemText;
        }

        def.Items.Add(new ChecklistItem
        {
          Id = itemId,
          Text = itemText,
          IsSubheader = item != null && item.IsSubheader,
          Order = item?.Order ?? order
        });
      }

      return def;
    }

    internal static string NormalizePlainText(string value)
    {
      return (value ?? string.Empty).Trim();
    }

    private static string NormalizePath(string path)
    {
      string trimmed = NormalizePlainText(path).Trim('"').Trim('\'');
      if (string.IsNullOrWhiteSpace(trimmed))
      {
        return string.Empty;
      }

      try
      {
        return Path.GetFullPath(trimmed);
      }
      catch
      {
        return trimmed;
      }
    }

    private static bool IsTemporaryPath(string path)
    {
      if (string.IsNullOrWhiteSpace(path))
      {
        return true;
      }

      string tempPath = NormalizePath(Path.GetTempPath());
      string normalized = NormalizePath(path);
      return !string.IsNullOrWhiteSpace(tempPath) &&
        normalized.StartsWith(tempPath, StringComparison.OrdinalIgnoreCase);
    }
  }

  internal sealed class ChecklistDefinitionsFile
  {
    [JsonProperty("checklists")]
    public List<ChecklistDefinition> Checklists { get; set; } = new List<ChecklistDefinition>();
  }

  internal sealed class ChecklistDefinition
  {
    [JsonProperty("id")]
    public string Id { get; set; }

    [JsonProperty("name")]
    public string Name { get; set; }

    [JsonProperty("category")]
    public string Category { get; set; }

    [JsonProperty("items")]
    public List<ChecklistItem> Items { get; set; } = new List<ChecklistItem>();

    internal void AddSubheader(string title)
    {
      Items.Add(new ChecklistItem
      {
        Id = $"subheader_{Items.Count + 1}",
        Text = title,
        IsSubheader = true,
        Order = Items.Count + 1
      });
    }

    internal void AddItem(string text)
    {
      Items.Add(new ChecklistItem
      {
        Id = text,
        Text = text,
        IsSubheader = false,
        Order = Items.Count + 1
      });
    }
  }

  internal sealed class ChecklistItem
  {
    [JsonProperty("id")]
    public string Id { get; set; }

    [JsonProperty("text")]
    public string Text { get; set; }

    [JsonProperty("isSubheader")]
    public bool IsSubheader { get; set; }

    [JsonProperty("order")]
    public int Order { get; set; }
  }

  internal sealed class FolderChecklistState
  {
    [JsonProperty("version")]
    public long Version { get; set; } = 1;

    [JsonProperty("lastModifiedUtc")]
    public string LastModifiedUtc { get; set; }

    [JsonProperty("lastActiveChecklistId")]
    public string LastActiveChecklistId { get; set; }

    [JsonProperty("checklists")]
    public Dictionary<string, FolderChecklistEntry> Checklists { get; set; } =
      new Dictionary<string, FolderChecklistEntry>(StringComparer.OrdinalIgnoreCase);
  }

  internal sealed class FolderChecklistEntry
  {
    [JsonProperty("completedItems")]
    public List<string> CompletedItems { get; set; } = new List<string>();

    [JsonProperty("itemNotes")]
    public Dictionary<string, string> ItemNotes { get; set; } =
      new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    [JsonProperty("lastModifiedUtc")]
    public string LastModifiedUtc { get; set; }
  }
}
