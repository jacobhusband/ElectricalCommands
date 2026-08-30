using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ElectricalCommands
{
  internal sealed class DedicatedEquipmentLoad
  {
    [JsonProperty("description")]
    public string Description { get; set; } = string.Empty;

    [JsonProperty("aliases")]
    public List<string> Aliases { get; set; } = new List<string>();

    [JsonProperty("kva")]
    public double Kva { get; set; }

    [JsonProperty("voltage")]
    public int Voltage { get; set; } = 120;

    [JsonProperty("poles")]
    public int Poles { get; set; } = 1;

    [JsonProperty("mcaAmps")]
    public double? McaAmps { get; set; }

    [JsonProperty("mocpAmps")]
    public int? MocpAmps { get; set; }

    [JsonProperty("loadTypeCode")]
    public string LoadTypeCode { get; set; } = "D";
  }

  internal sealed class DedicatedEquipmentCatalogFile
  {
    [JsonProperty("version")]
    public int Version { get; set; } = 1;

    [JsonProperty("items")]
    public List<DedicatedEquipmentLoad> Items { get; set; } =
      new List<DedicatedEquipmentLoad>();
  }

  internal static class DedicatedEquipmentCatalog
  {
    private const string CatalogFileName = "dedicated_equipment_loads.json";

    internal static bool TryFind(
      string description,
      out DedicatedEquipmentLoad equipment,
      out string catalogPath)
    {
      catalogPath = ResolveCatalogPath();
      DedicatedEquipmentCatalogFile catalog = LoadOrCreate(catalogPath);
      string requested = Normalize(description);
      equipment = null;
      int bestMatchLength = -1;

      foreach (DedicatedEquipmentLoad candidate in catalog.Items)
      {
        foreach (string name in GetNames(candidate))
        {
          string normalizedName = Normalize(name);
          if (normalizedName.Length == 0)
          {
            continue;
          }

          if (string.Equals(
            requested,
            normalizedName,
            StringComparison.OrdinalIgnoreCase))
          {
            equipment = candidate;
            return true;
          }

          if (normalizedName.Length >= 4 &&
              requested.IndexOf(
                normalizedName,
                StringComparison.OrdinalIgnoreCase) >= 0 &&
              normalizedName.Length > bestMatchLength)
          {
            equipment = candidate;
            bestMatchLength = normalizedName.Length;
          }
        }
      }
      return equipment != null;
    }

    internal static string SaveOrUpdate(DedicatedEquipmentLoad equipment)
    {
      if (equipment == null || string.IsNullOrWhiteSpace(equipment.Description))
      {
        throw new ArgumentException(
          "A dedicated-equipment description is required.",
          nameof(equipment));
      }

      string catalogPath = ResolveCatalogPath();
      DedicatedEquipmentCatalogFile catalog = LoadOrCreate(catalogPath);
      string key = Normalize(equipment.Description);
      int existingIndex = catalog.Items.FindIndex(
        item => string.Equals(
          Normalize(item.Description),
          key,
          StringComparison.OrdinalIgnoreCase));
      if (existingIndex >= 0)
      {
        catalog.Items[existingIndex] = equipment;
      }
      else
      {
        catalog.Items.Add(equipment);
      }

      catalog.Items = catalog.Items
        .OrderBy(item => item.Description, StringComparer.OrdinalIgnoreCase)
        .ToList();
      Save(catalogPath, catalog);
      return catalogPath;
    }

    private static DedicatedEquipmentCatalogFile LoadOrCreate(
      string catalogPath)
    {
      bool catalogExists = File.Exists(catalogPath);
      if (catalogExists)
      {
        try
        {
          string json = File.ReadAllText(catalogPath, Encoding.UTF8);
          DedicatedEquipmentCatalogFile existing =
            JsonConvert.DeserializeObject<DedicatedEquipmentCatalogFile>(
              json);
          if (existing?.Items != null && existing.Items.Count > 0)
          {
            return existing;
          }
        }
        catch
        {
          // Keep the command usable with built-in defaults if a user-edited
          // catalog is temporarily invalid. Saving a custom preset will repair it.
        }
      }

      DedicatedEquipmentCatalogFile defaults = CreateDefaultCatalog();
      if (catalogExists)
      {
        return defaults;
      }
      try
      {
        Save(catalogPath, defaults);
      }
      catch
      {
        // A read-only profile still gets the in-memory default catalog.
      }
      return defaults;
    }

    private static void Save(
      string catalogPath,
      DedicatedEquipmentCatalogFile catalog)
    {
      string directory = Path.GetDirectoryName(catalogPath);
      if (!Directory.Exists(directory))
      {
        Directory.CreateDirectory(directory);
      }
      string temporaryPath = catalogPath + ".tmp";
      try
      {
        File.WriteAllText(
          temporaryPath,
          JsonConvert.SerializeObject(catalog, Formatting.Indented),
          new UTF8Encoding(false));
        if (File.Exists(catalogPath))
        {
          File.Replace(
            temporaryPath,
            catalogPath,
            catalogPath + ".bak",
            true);
        }
        else
        {
          File.Move(temporaryPath, catalogPath);
        }
      }
      finally
      {
        if (File.Exists(temporaryPath))
        {
          try
          {
            File.Delete(temporaryPath);
          }
          catch
          {
          }
        }
      }
    }

    private static string ResolveCatalogPath()
    {
      return Path.Combine(
        ProjectChecklistStore.ResolveAppDataFolder(),
        CatalogFileName);
    }

    private static IEnumerable<string> GetNames(
      DedicatedEquipmentLoad equipment)
    {
      yield return equipment?.Description ?? string.Empty;
      foreach (string alias in equipment?.Aliases ?? new List<string>())
      {
        yield return alias;
      }
    }

    private static string Normalize(string value)
    {
      return Regex.Replace(
        value ?? string.Empty,
        @"[^A-Z0-9]+",
        string.Empty,
        RegexOptions.IgnoreCase).ToUpperInvariant();
    }

    private static DedicatedEquipmentCatalogFile CreateDefaultCatalog()
    {
      return new DedicatedEquipmentCatalogFile
      {
        Items = new List<DedicatedEquipmentLoad>
        {
          Create("COFFEE MAKER", 1.50, "COFFEE MACHINE"),
          Create(
            "DISHWASHER",
            1.80,
            "DISH WASHER",
            "DISHERWASHER",
            "DW"),
          Create("DRYER", 1.80, "GAS DRYER"),
          Create(
            "GARBAGE DISPOSAL",
            1.20,
            "GARBAGE DISPOSER",
            "FOOD WASTE DISPOSAL",
            "FOOD WASTE DISPOSER",
            "DISPOSAL"),
          Create("ICE MAKER", 0.50, "ICEMAKER"),
          Create("MICROWAVE", 1.50, "MICROWAVE OVEN"),
          Create("RANGE HOOD", 0.50, "HOOD"),
          Create("REFRIGERATOR", 0.50, "FRIDGE", "REFRIG"),
          Create("FREEZER", 0.50),
          Create("WASHER", 1.80, "WASHING MACHINE", "CLOTHES WASHER"),
          Create("WATER COOLER", 0.50, "DRINKING WATER COOLER"),
        },
      };
    }

    private static DedicatedEquipmentLoad Create(
      string description,
      double kva,
      params string[] aliases)
    {
      return new DedicatedEquipmentLoad
      {
        Description = description,
        Aliases = new List<string>(aliases ?? new string[0]),
        Kva = kva,
        Voltage = 120,
        Poles = 1,
        McaAmps = null,
        MocpAmps = 20,
        LoadTypeCode = "D",
      };
    }
  }
}
