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

    internal static List<DedicatedEquipmentLoad> LoadAll(
      out string catalogPath)
    {
      catalogPath = ResolveCatalogPath();
      DedicatedEquipmentCatalogFile catalog = LoadOrCreate(catalogPath);
      return catalog.Items
        .Where(item => item != null)
        .Select(Copy)
        .OrderBy(item => item.Description, StringComparer.OrdinalIgnoreCase)
        .ToList();
    }

    internal static string SaveOrUpdate(
      DedicatedEquipmentLoad equipment,
      string originalDescription = null)
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
      string originalKey = Normalize(originalDescription);
      int originalIndex = originalKey.Length == 0
        ? -1
        : catalog.Items.FindIndex(
          item => string.Equals(
            Normalize(item.Description),
            originalKey,
            StringComparison.OrdinalIgnoreCase));
      int matchingDescriptionIndex = catalog.Items.FindIndex(
        item => string.Equals(
          Normalize(item.Description),
          key,
          StringComparison.OrdinalIgnoreCase));
      if (originalIndex >= 0 &&
          matchingDescriptionIndex >= 0 &&
          matchingDescriptionIndex != originalIndex)
      {
        throw new InvalidOperationException(
          $"A {equipment.Description} preset already exists.");
      }

      int saveIndex = originalIndex >= 0
        ? originalIndex
        : matchingDescriptionIndex;
      if (saveIndex >= 0)
      {
        catalog.Items[saveIndex] = Copy(equipment);
      }
      else
      {
        catalog.Items.Add(Copy(equipment));
      }

      catalog.Items = catalog.Items
        .OrderBy(item => item.Description, StringComparer.OrdinalIgnoreCase)
        .ToList();
      Save(catalogPath, catalog);
      return catalogPath;
    }

    internal static string Remove(string description)
    {
      string key = Normalize(description);
      if (key.Length == 0)
      {
        throw new ArgumentException(
          "A preset description is required.",
          nameof(description));
      }

      string catalogPath = ResolveCatalogPath();
      DedicatedEquipmentCatalogFile catalog = LoadOrCreate(catalogPath);
      int removed = catalog.Items.RemoveAll(item =>
        string.Equals(
          Normalize(item?.Description),
          key,
          StringComparison.OrdinalIgnoreCase));
      if (removed == 0)
      {
        throw new InvalidOperationException(
          $"The {description} preset no longer exists.");
      }

      Save(catalogPath, catalog);
      return catalogPath;
    }

    private static DedicatedEquipmentCatalogFile LoadOrCreate(
      string catalogPath)
    {
      DedicatedEquipmentCatalogFile defaults = CreateDefaultCatalog();
      bool catalogExists = File.Exists(catalogPath);
      if (catalogExists)
      {
        try
        {
          string json = File.ReadAllText(catalogPath, Encoding.UTF8);
          DedicatedEquipmentCatalogFile existing =
            JsonConvert.DeserializeObject<DedicatedEquipmentCatalogFile>(
              json);
          if (existing?.Items != null)
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

    private static DedicatedEquipmentLoad Copy(
      DedicatedEquipmentLoad equipment)
    {
      if (equipment == null)
      {
        return new DedicatedEquipmentLoad();
      }
      return new DedicatedEquipmentLoad
      {
        Description = equipment.Description ?? string.Empty,
        Aliases = new List<string>(
          equipment.Aliases ?? new List<string>()),
        Kva = equipment.Kva,
        Voltage = equipment.Voltage,
        Poles = equipment.Poles,
        McaAmps = equipment.McaAmps,
        MocpAmps = equipment.MocpAmps,
        LoadTypeCode = equipment.LoadTypeCode ?? "D",
      };
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
          Create(
            "KITCHEN COUNTERTOP OUTLET",
            0.18,
            "COUNTER",
            "COUNTERTOP",
            "COUNTER TOP",
            "KITCHEN COUNTER",
            "KITCHEN COUNTERTOP",
            "KITCHEN COUNTER TOP",
            "KITCHEN COUNTERTOP RECEPTACLE",
            "KITCHEN COUNTER RECEPTACLE",
            "COUNTERTOP RECEPTACLE",
            "COUNTER RECEPTACLE",
            "COUNTERTOP OUTLET",
            "COUNTER OUTLET",
            "KITCHEN OUTLET",
            "KITCHEN RECEPTACLE"),
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
