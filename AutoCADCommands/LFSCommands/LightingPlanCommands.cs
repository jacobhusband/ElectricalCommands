using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private const string LightingPlanSnapshotFileName = "ACIESLightingPlan.snapshot.json";
    private const string LightingPlanInstructionFileName = "ACIESLightingPlan.instructions.json";
    internal const string LightingPlanConfigFileName = "ACIESLightingPlan.config.json";
    private const string LightingPlanRegAppName = "ACIES_LIGHTING_PLAN";

    [CommandMethod("LIGHTPLANSCAN", CommandFlags.Modal)]
    public void LightingPlanScan()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null)
      {
        return;
      }

      Editor editor = doc.Editor;
      string drawingPath = NormalizeDrawingPath(doc.Name);
      if (string.IsNullOrWhiteSpace(drawingPath) || !Path.IsPathRooted(drawingPath))
      {
        editor.WriteMessage("\nSave the drawing before running LIGHTPLANSCAN.");
        return;
      }

      try
      {
        string drawingFolder = Path.GetDirectoryName(drawingPath);
        LightingPlanCadConfig config = LightingPlanCadConfig.LoadOrCreate(drawingFolder);
        LightingPlanSnapshot snapshot;
        using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
        {
          snapshot = ScanLightingPlan(doc.Database, transaction, drawingPath, config);
          transaction.Commit();
        }

        string snapshotPath = Path.Combine(drawingFolder, LightingPlanSnapshotFileName);
        WriteLightingPlanJson(snapshotPath, snapshot);
        editor.WriteMessage(
          $"\nLIGHTPLANSCAN exported {snapshot.Fixtures.Count} fixture(s) and " +
          $"{snapshot.Rooms.Count} room boundary(ies) to:\n{snapshotPath}"
        );
        if (snapshot.Fixtures.Count == 0 || snapshot.Rooms.Count == 0)
        {
          editor.WriteMessage(
            $"\nReview {LightingPlanConfigFileName} beside the drawing, update layer/block patterns, and scan again."
          );
        }
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage($"\nLIGHTPLANSCAN failed: {ex.Message}");
      }
    }

    [CommandMethod("LPSCAN", CommandFlags.Modal)]
    public void LightingPlanScanAlias()
    {
      LightingPlanScan();
    }

    [CommandMethod("LIGHTPLANAPPLY", CommandFlags.Modal)]
    public void LightingPlanApply()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null)
      {
        return;
      }

      Editor editor = doc.Editor;
      string drawingPath = NormalizeDrawingPath(doc.Name);
      if (string.IsNullOrWhiteSpace(drawingPath) || !Path.IsPathRooted(drawingPath))
      {
        editor.WriteMessage("\nSave the drawing before running LIGHTPLANAPPLY.");
        return;
      }

      string drawingFolder = Path.GetDirectoryName(drawingPath);
      string instructionPath = Path.Combine(drawingFolder, LightingPlanInstructionFileName);
      if (!File.Exists(instructionPath))
      {
        editor.WriteMessage(
          $"\nNo {LightingPlanInstructionFileName} was found. Import the scan and prepare CAD tags in ACIES first."
        );
        return;
      }

      try
      {
        LightingPlanGenerationInstructions instructions =
          JsonConvert.DeserializeObject<LightingPlanGenerationInstructions>(
            File.ReadAllText(instructionPath, Encoding.UTF8)
          );
        if (instructions == null || instructions.Tags == null)
        {
          throw new InvalidDataException("The lighting plan instruction file is empty or invalid.");
        }
        if (!string.Equals(instructions.Kind, "acies-lighting-plan-generation", StringComparison.OrdinalIgnoreCase))
        {
          throw new InvalidDataException("The JSON file is not an ACIES lighting plan generation file.");
        }
        string instructionDrawingPath = NormalizeDrawingPath(instructions.DrawingPath);
        if (
          !string.IsNullOrWhiteSpace(instructionDrawingPath) &&
          !string.Equals(instructionDrawingPath, drawingPath, StringComparison.OrdinalIgnoreCase)
        )
        {
          throw new InvalidDataException(
            $"Instructions were prepared for '{instructionDrawingPath}', not the active drawing."
          );
        }
        string activeFingerprint = doc.Database.FingerprintGuid.ToString();
        if (
          !string.IsNullOrWhiteSpace(instructions.SourceFingerprint) &&
          !string.Equals(
            instructions.SourceFingerprint,
            activeFingerprint,
            StringComparison.OrdinalIgnoreCase
          )
        )
        {
          throw new InvalidDataException(
            "The drawing fingerprint changed after the scan. Run LIGHTPLANSCAN and review the new snapshot before applying tags."
          );
        }

        LightingPlanCadConfig config = LightingPlanCadConfig.LoadOrCreate(drawingFolder);
        LightingPlanApplyResult result;
        using (doc.LockDocument())
        using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
        {
          result = ApplyLightingPlanTags(doc.Database, transaction, instructions, config);
          transaction.Commit();
        }

        editor.WriteMessage(
          $"\nLIGHTPLANAPPLY complete: {result.Created} created, {result.Updated} updated, " +
          $"{result.Removed} stale ACIES tag(s) removed."
        );
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage($"\nLIGHTPLANAPPLY failed: {ex.Message}");
      }
    }

    [CommandMethod("LPAPPLY", CommandFlags.Modal)]
    public void LightingPlanApplyAlias()
    {
      LightingPlanApply();
    }

    private static LightingPlanSnapshot ScanLightingPlan(
      Database database,
      Transaction transaction,
      string drawingPath,
      LightingPlanCadConfig config
    )
    {
      BlockTable blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
      BlockTableRecord modelSpace = (BlockTableRecord)transaction.GetObject(
        blockTable[BlockTableRecord.ModelSpace],
        OpenMode.ForRead
      );

      var textCandidates = new List<LightingPlanTextCandidate>();
      foreach (ObjectId objectId in modelSpace)
      {
        Entity entity = transaction.GetObject(objectId, OpenMode.ForRead, false) as Entity;
        if (entity is DBText dbText)
        {
          textCandidates.Add(
            new LightingPlanTextCandidate
            {
              Layer = dbText.Layer,
              Text = NormalizeCadText(dbText.TextString),
              Position = dbText.Position,
            }
          );
        }
        else if (entity is MText mtext)
        {
          textCandidates.Add(
            new LightingPlanTextCandidate
            {
              Layer = mtext.Layer,
              Text = NormalizeCadText(mtext.Text),
              Position = mtext.Location,
            }
          );
        }
      }

      var rooms = new List<LightingPlanRoomSnapshot>();
      var fixtures = new List<LightingPlanFixtureSnapshot>();
      foreach (ObjectId objectId in modelSpace)
      {
        Entity entity = transaction.GetObject(objectId, OpenMode.ForRead, false) as Entity;
        if (entity is Polyline polyline &&
            polyline.Closed &&
            MatchesAny(polyline.Layer, config.RoomBoundaryLayerPatterns))
        {
          LightingPlanRoomSnapshot room = CreateRoomSnapshot(polyline, textCandidates, config);
          rooms.Add(room);
          continue;
        }

        if (entity is BlockReference blockReference)
        {
          Dictionary<string, string> attributes = ReadBlockAttributes(blockReference, transaction);
          string blockName = ResolveBlockName(blockReference, transaction);
          if (!IsFixtureCandidate(blockReference.Layer, blockName, attributes, config))
          {
            continue;
          }
          fixtures.Add(CreateFixtureSnapshot(blockReference, blockName, attributes, config));
        }
      }

      return new LightingPlanSnapshot
      {
        SchemaVersion = "1.0.0",
        Drawing = new LightingPlanDrawingSnapshot
        {
          Path = drawingPath,
          Name = Path.GetFileName(drawingPath),
          Fingerprint = database.FingerprintGuid.ToString(),
          Units = database.Insunits.ToString(),
          ScannedAtUtc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture),
        },
        Config = config.ToSnapshotConfig(),
        Rooms = rooms,
        Fixtures = fixtures,
      };
    }

    private static LightingPlanRoomSnapshot CreateRoomSnapshot(
      Polyline polyline,
      List<LightingPlanTextCandidate> texts,
      LightingPlanCadConfig config
    )
    {
      var boundary = new List<LightingPlanPoint>();
      for (int index = 0; index < polyline.NumberOfVertices; index++)
      {
        Point2d vertex = polyline.GetPoint2dAt(index);
        boundary.Add(new LightingPlanPoint { X = vertex.X, Y = vertex.Y, Z = polyline.Elevation });
      }
      Point3d center = GetAveragePoint(boundary);
      List<LightingPlanTextCandidate> inside = texts
        .Where(candidate => PointInsidePolygon(candidate.Position, boundary))
        .ToList();
      LightingPlanTextCandidate nameText = FindNearestText(
        inside.Where(candidate => MatchesAny(candidate.Layer, config.RoomNameLayerPatterns)),
        center
      );
      LightingPlanTextCandidate typeText = FindNearestText(
        inside.Where(candidate => MatchesAny(candidate.Layer, config.RoomTypeLayerPatterns)),
        center
      );
      if (nameText == null)
      {
        nameText = FindNearestText(inside, center);
      }

      double area = 0.0;
      try
      {
        area = polyline.Area;
      }
      catch
      {
        area = CalculatePolygonArea(boundary);
      }
      string roomName = nameText?.Text ?? string.Empty;
      string roomType = typeText?.Text ?? roomName;
      return new LightingPlanRoomSnapshot
      {
        Id = polyline.Handle.ToString(),
        Handle = polyline.Handle.ToString(),
        Name = roomName,
        RoomType = roomType,
        Layer = polyline.Layer,
        Boundary = boundary,
        Area = area,
      };
    }

    private static LightingPlanFixtureSnapshot CreateFixtureSnapshot(
      BlockReference blockReference,
      string blockName,
      Dictionary<string, string> attributes,
      LightingPlanCadConfig config
    )
    {
      string handle = blockReference.Handle.ToString();
      return new LightingPlanFixtureSnapshot
      {
        Id = handle,
        Handle = handle,
        BlockName = blockName,
        Layer = blockReference.Layer,
        Position = LightingPlanPoint.From(blockReference.Position),
        Rotation = blockReference.Rotation,
        Attributes = attributes,
        Mark = FirstAttribute(attributes, config.MarkAttributeAliases),
        Panel = FirstAttribute(attributes, config.PanelAttributeAliases),
        Circuit = FirstAttribute(attributes, config.CircuitAttributeAliases),
        ControlZone = FirstAttribute(attributes, config.ControlZoneAttributeAliases),
        Watts = FirstAttribute(attributes, config.WattsAttributeAliases),
      };
    }

    private static Dictionary<string, string> ReadBlockAttributes(
      BlockReference blockReference,
      Transaction transaction
    )
    {
      var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
      foreach (ObjectId attributeId in blockReference.AttributeCollection)
      {
        AttributeReference attribute = transaction.GetObject(
          attributeId,
          OpenMode.ForRead,
          false
        ) as AttributeReference;
        if (attribute == null || string.IsNullOrWhiteSpace(attribute.Tag))
        {
          continue;
        }
        result[attribute.Tag.Trim().ToUpperInvariant()] = (attribute.TextString ?? string.Empty).Trim();
      }
      return result;
    }

    private static string ResolveBlockName(BlockReference blockReference, Transaction transaction)
    {
      ObjectId definitionId = blockReference.IsDynamicBlock
        ? blockReference.DynamicBlockTableRecord
        : blockReference.BlockTableRecord;
      BlockTableRecord definition = transaction.GetObject(definitionId, OpenMode.ForRead) as BlockTableRecord;
      return definition?.Name ?? string.Empty;
    }

    private static bool IsFixtureCandidate(
      string layer,
      string blockName,
      Dictionary<string, string> attributes,
      LightingPlanCadConfig config
    )
    {
      if (MatchesAny(layer, config.FixtureLayerPatterns) ||
          MatchesAny(blockName, config.FixtureBlockPatterns))
      {
        return true;
      }
      if (!config.IncludeAttributedFixtureBlocks)
      {
        return false;
      }
      bool hasMark = !string.IsNullOrWhiteSpace(FirstAttribute(attributes, config.MarkAttributeAliases));
      bool hasElectricalData =
        !string.IsNullOrWhiteSpace(FirstAttribute(attributes, config.PanelAttributeAliases)) ||
        !string.IsNullOrWhiteSpace(FirstAttribute(attributes, config.CircuitAttributeAliases)) ||
        !string.IsNullOrWhiteSpace(FirstAttribute(attributes, config.WattsAttributeAliases));
      return hasMark && hasElectricalData;
    }

    private static LightingPlanApplyResult ApplyLightingPlanTags(
      Database database,
      Transaction transaction,
      LightingPlanGenerationInstructions instructions,
      LightingPlanCadConfig config
    )
    {
      EnsureRegApp(database, transaction, LightingPlanRegAppName);
      string layerName = !string.IsNullOrWhiteSpace(instructions.TagLayer)
        ? instructions.TagLayer.Trim()
        : config.TagLayer;
      EnsureLayer(database, transaction, layerName);

      BlockTable blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
      BlockTableRecord modelSpace = (BlockTableRecord)transaction.GetObject(
        blockTable[BlockTableRecord.ModelSpace],
        OpenMode.ForWrite
      );
      Dictionary<string, MText> existing = FindExistingLightingPlanTags(modelSpace, transaction);
      var desiredKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
      int created = 0;
      int updated = 0;

      foreach (LightingPlanTagInstruction instruction in instructions.Tags)
      {
        if (instruction == null ||
            string.IsNullOrWhiteSpace(instruction.GenerationKey) ||
            instruction.Position == null ||
            string.IsNullOrWhiteSpace(instruction.Text))
        {
          continue;
        }
        string key = instruction.GenerationKey.Trim();
        desiredKeys.Add(key);
        string instructionLayer = string.IsNullOrWhiteSpace(instruction.Layer)
          ? layerName
          : instruction.Layer.Trim();
        EnsureLayer(database, transaction, instructionLayer);
        if (existing.TryGetValue(key, out MText current) && current != null && !current.IsErased)
        {
          if (!current.IsWriteEnabled)
          {
            current.UpgradeOpen();
          }
          ApplyTagProperties(current, instruction, instructionLayer, config, database);
          SetLightingPlanXData(current, key, instructions.SourceFingerprint, instruction.SourceHandle);
          updated++;
        }
        else
        {
          var tag = new MText();
          tag.SetDatabaseDefaults(database);
          ApplyTagProperties(tag, instruction, instructionLayer, config, database);
          modelSpace.AppendEntity(tag);
          transaction.AddNewlyCreatedDBObject(tag, true);
          SetLightingPlanXData(tag, key, instructions.SourceFingerprint, instruction.SourceHandle);
          created++;
        }
      }

      int removed = 0;
      if (config.RemoveStaleGeneratedTags)
      {
        foreach (KeyValuePair<string, MText> pair in existing)
        {
          if (desiredKeys.Contains(pair.Key) || pair.Value == null || pair.Value.IsErased)
          {
            continue;
          }
          if (!pair.Value.IsWriteEnabled)
          {
            pair.Value.UpgradeOpen();
          }
          pair.Value.Erase();
          removed++;
        }
      }
      return new LightingPlanApplyResult { Created = created, Updated = updated, Removed = removed };
    }

    private static void ApplyTagProperties(
      MText tag,
      LightingPlanTagInstruction instruction,
      string layerName,
      LightingPlanCadConfig config,
      Database database
    )
    {
      tag.Contents = instruction.Text;
      tag.Location = instruction.Position.ToPoint3d();
      tag.Layer = layerName;
      tag.Attachment = AttachmentPoint.MiddleCenter;
      double textHeight = config.TagTextHeight > 0 ? config.TagTextHeight : database.Textsize;
      if (textHeight > 0)
      {
        tag.TextHeight = textHeight;
      }
    }

    private static Dictionary<string, MText> FindExistingLightingPlanTags(
      BlockTableRecord modelSpace,
      Transaction transaction
    )
    {
      var result = new Dictionary<string, MText>(StringComparer.OrdinalIgnoreCase);
      foreach (ObjectId objectId in modelSpace)
      {
        MText tag = transaction.GetObject(objectId, OpenMode.ForRead, false) as MText;
        if (tag == null)
        {
          continue;
        }
        string generationKey = ReadLightingPlanGenerationKey(tag);
        if (!string.IsNullOrWhiteSpace(generationKey) && !result.ContainsKey(generationKey))
        {
          result[generationKey] = tag;
        }
      }
      return result;
    }

    private static string ReadLightingPlanGenerationKey(Entity entity)
    {
      using (ResultBuffer buffer = entity.GetXDataForApplication(LightingPlanRegAppName))
      {
        if (buffer == null)
        {
          return string.Empty;
        }
        TypedValue[] values = buffer.AsArray();
        return values
          .Where(value => value.TypeCode == 1000)
          .Select(value => Convert.ToString(value.Value, CultureInfo.InvariantCulture))
          .FirstOrDefault() ?? string.Empty;
      }
    }

    private static void SetLightingPlanXData(
      Entity entity,
      string generationKey,
      string sourceFingerprint,
      string sourceHandle
    )
    {
      using (var buffer = new ResultBuffer(
        new TypedValue(1001, LightingPlanRegAppName),
        new TypedValue(1000, generationKey ?? string.Empty),
        new TypedValue(1000, sourceFingerprint ?? string.Empty),
        new TypedValue(1000, sourceHandle ?? string.Empty)
      ))
      {
        entity.XData = buffer;
      }
    }

    private static void EnsureRegApp(Database database, Transaction transaction, string name)
    {
      RegAppTable table = (RegAppTable)transaction.GetObject(database.RegAppTableId, OpenMode.ForRead);
      if (table.Has(name))
      {
        return;
      }
      table.UpgradeOpen();
      var record = new RegAppTableRecord { Name = name };
      table.Add(record);
      transaction.AddNewlyCreatedDBObject(record, true);
    }

    private static void EnsureLayer(Database database, Transaction transaction, string name)
    {
      LayerTable table = (LayerTable)transaction.GetObject(database.LayerTableId, OpenMode.ForRead);
      if (table.Has(name))
      {
        return;
      }
      table.UpgradeOpen();
      var record = new LayerTableRecord { Name = name };
      table.Add(record);
      transaction.AddNewlyCreatedDBObject(record, true);
    }

    private static string NormalizeDrawingPath(string path)
    {
      if (string.IsNullOrWhiteSpace(path))
      {
        return string.Empty;
      }
      try
      {
        return Path.GetFullPath(path.Trim());
      }
      catch
      {
        return path.Trim();
      }
    }

    internal static void WriteLightingPlanJson(string path, object payload)
    {
      var settings = new JsonSerializerSettings
      {
        ContractResolver = new CamelCasePropertyNamesContractResolver(),
        Formatting = Formatting.Indented,
        NullValueHandling = NullValueHandling.Include,
      };
      string json = JsonConvert.SerializeObject(payload, settings);
      string temporaryPath = path + ".tmp";
      File.WriteAllText(temporaryPath, json, new UTF8Encoding(false));
      if (File.Exists(path))
      {
        File.Replace(temporaryPath, path, null);
      }
      else
      {
        File.Move(temporaryPath, path);
      }
    }

    private static string NormalizeCadText(string value)
    {
      string text = (value ?? string.Empty).Replace("\\P", " ").Replace("\r", " ").Replace("\n", " ");
      text = Regex.Replace(text, @"\\[A-Za-z][^;]*;", string.Empty);
      return Regex.Replace(text, @"\s+", " ").Trim();
    }

    private static string FirstAttribute(
      Dictionary<string, string> attributes,
      IEnumerable<string> aliases
    )
    {
      if (attributes == null || aliases == null)
      {
        return string.Empty;
      }
      foreach (string alias in aliases)
      {
        if (attributes.TryGetValue((alias ?? string.Empty).Trim(), out string value) &&
            !string.IsNullOrWhiteSpace(value))
        {
          return value.Trim();
        }
      }
      return string.Empty;
    }

    private static bool MatchesAny(string value, IEnumerable<string> patterns)
    {
      if (string.IsNullOrWhiteSpace(value) || patterns == null)
      {
        return false;
      }
      foreach (string pattern in patterns)
      {
        string normalizedPattern = (pattern ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalizedPattern))
        {
          continue;
        }
        string regexPattern = "^" + Regex.Escape(normalizedPattern)
          .Replace("\\*", ".*")
          .Replace("\\?", ".") + "$";
        if (Regex.IsMatch(value, regexPattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
        {
          return true;
        }
      }
      return false;
    }

    private static Point3d GetAveragePoint(List<LightingPlanPoint> boundary)
    {
      if (boundary == null || boundary.Count == 0)
      {
        return Point3d.Origin;
      }
      return new Point3d(
        boundary.Average(point => point.X),
        boundary.Average(point => point.Y),
        boundary.Average(point => point.Z)
      );
    }

    private static LightingPlanTextCandidate FindNearestText(
      IEnumerable<LightingPlanTextCandidate> candidates,
      Point3d center
    )
    {
      return candidates
        .Where(candidate => candidate != null && !string.IsNullOrWhiteSpace(candidate.Text))
        .OrderBy(candidate => candidate.Position.DistanceTo(center))
        .FirstOrDefault();
    }

    private static bool PointInsidePolygon(Point3d point, List<LightingPlanPoint> boundary)
    {
      if (boundary == null || boundary.Count < 3)
      {
        return false;
      }
      bool inside = false;
      for (int current = 0, previous = boundary.Count - 1; current < boundary.Count; previous = current++)
      {
        LightingPlanPoint a = boundary[current];
        LightingPlanPoint b = boundary[previous];
        bool intersects = ((a.Y > point.Y) != (b.Y > point.Y)) &&
          (point.X < (b.X - a.X) * (point.Y - a.Y) / (b.Y - a.Y) + a.X);
        if (intersects)
        {
          inside = !inside;
        }
      }
      return inside;
    }

    private static double CalculatePolygonArea(List<LightingPlanPoint> boundary)
    {
      if (boundary == null || boundary.Count < 3)
      {
        return 0.0;
      }
      double doubledArea = 0.0;
      for (int index = 0; index < boundary.Count; index++)
      {
        LightingPlanPoint current = boundary[index];
        LightingPlanPoint next = boundary[(index + 1) % boundary.Count];
        doubledArea += current.X * next.Y - next.X * current.Y;
      }
      return Math.Abs(doubledArea) / 2.0;
    }
  }

  internal sealed class LightingPlanCadConfig
  {
    public string SchemaVersion { get; set; } = "1.0.0";
    public List<string> FixtureLayerPatterns { get; set; } = new List<string>
    {
      "E-LITE*", "E-LIGHT*", "E-POWR-LITE*",
    };
    public List<string> FixtureBlockPatterns { get; set; } = new List<string>
    {
      "*LIGHT*", "*LITE*", "*FIXTURE*",
    };
    public List<string> RoomBoundaryLayerPatterns { get; set; } = new List<string>
    {
      "E-LITE-ROOM*", "E-ROOM*", "A-AREA*", "A-ROOM*",
    };
    public List<string> RoomNameLayerPatterns { get; set; } = new List<string>
    {
      "E-LITE-ROOM-TEXT*", "A-ROOM*", "A-ANNO-TEXT*",
    };
    public List<string> RoomTypeLayerPatterns { get; set; } = new List<string>
    {
      "E-LITE-ROOM-TYPE*",
    };
    public List<string> MarkAttributeAliases { get; set; } = new List<string>
    {
      "MARK", "TYPE", "FIXTURE_TYPE", "FIXTURETYPE",
    };
    public List<string> PanelAttributeAliases { get; set; } = new List<string>
    {
      "PANEL", "PANEL_NAME", "PANELNAME",
    };
    public List<string> CircuitAttributeAliases { get; set; } = new List<string>
    {
      "CIRCUIT", "CKT", "CIRCUIT_NUMBER", "CIRCUITNUMBER",
    };
    public List<string> ControlZoneAttributeAliases { get; set; } = new List<string>
    {
      "CONTROL_ZONE", "CONTROLZONE", "CONTROL", "ZONE",
    };
    public List<string> WattsAttributeAliases { get; set; } = new List<string>
    {
      "WATTS", "WATTAGE", "LOAD_WATTS", "LOADWATTS",
    };
    public bool IncludeAttributedFixtureBlocks { get; set; } = true;
    public string TagLayer { get; set; } = "E-LITE-TAGS-ACIES";
    public double TagTextHeight { get; set; } = 0.0;
    public LightingPlanPoint TagOffset { get; set; } = new LightingPlanPoint { X = 1.5, Y = 1.5, Z = 0.0 };
    public bool RemoveStaleGeneratedTags { get; set; } = true;

    internal static LightingPlanCadConfig LoadOrCreate(string drawingFolder)
    {
      string path = Path.Combine(drawingFolder, GeneralCommands.LightingPlanConfigFileName);
      LightingPlanCadConfig config = null;
      if (File.Exists(path))
      {
        config = JsonConvert.DeserializeObject<LightingPlanCadConfig>(
          File.ReadAllText(path, Encoding.UTF8)
        );
      }
      config = config ?? new LightingPlanCadConfig();
      config.Normalize();
      if (!File.Exists(path))
      {
        GeneralCommands.WriteLightingPlanJson(path, config);
      }
      return config;
    }

    internal object ToSnapshotConfig()
    {
      return new
      {
        fixtureLayerPatterns = FixtureLayerPatterns,
        fixtureBlockPatterns = FixtureBlockPatterns,
        roomBoundaryLayerPatterns = RoomBoundaryLayerPatterns,
        roomNameLayerPatterns = RoomNameLayerPatterns,
        roomTypeLayerPatterns = RoomTypeLayerPatterns,
        tagLayer = TagLayer,
        tagOffset = TagOffset,
      };
    }

    private void Normalize()
    {
      FixtureLayerPatterns = FixtureLayerPatterns ?? new List<string>();
      FixtureBlockPatterns = FixtureBlockPatterns ?? new List<string>();
      RoomBoundaryLayerPatterns = RoomBoundaryLayerPatterns ?? new List<string>();
      RoomNameLayerPatterns = RoomNameLayerPatterns ?? new List<string>();
      RoomTypeLayerPatterns = RoomTypeLayerPatterns ?? new List<string>();
      MarkAttributeAliases = NormalizeAliases(MarkAttributeAliases, "MARK");
      PanelAttributeAliases = NormalizeAliases(PanelAttributeAliases, "PANEL");
      CircuitAttributeAliases = NormalizeAliases(CircuitAttributeAliases, "CIRCUIT");
      ControlZoneAttributeAliases = NormalizeAliases(ControlZoneAttributeAliases, "CONTROL_ZONE");
      WattsAttributeAliases = NormalizeAliases(WattsAttributeAliases, "WATTS");
      TagLayer = string.IsNullOrWhiteSpace(TagLayer) ? "E-LITE-TAGS-ACIES" : TagLayer.Trim();
      TagOffset = TagOffset ?? new LightingPlanPoint { X = 1.5, Y = 1.5, Z = 0.0 };
    }

    private static List<string> NormalizeAliases(List<string> values, string fallback)
    {
      List<string> result = (values ?? new List<string>())
        .Where(value => !string.IsNullOrWhiteSpace(value))
        .Select(value => value.Trim().ToUpperInvariant())
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToList();
      if (result.Count == 0)
      {
        result.Add(fallback);
      }
      return result;
    }
  }

  internal sealed class LightingPlanSnapshot
  {
    public string SchemaVersion { get; set; }
    public LightingPlanDrawingSnapshot Drawing { get; set; }
    public object Config { get; set; }
    public List<LightingPlanRoomSnapshot> Rooms { get; set; } = new List<LightingPlanRoomSnapshot>();
    public List<LightingPlanFixtureSnapshot> Fixtures { get; set; } = new List<LightingPlanFixtureSnapshot>();
  }

  internal sealed class LightingPlanDrawingSnapshot
  {
    public string Path { get; set; }
    public string Name { get; set; }
    public string Fingerprint { get; set; }
    public string Units { get; set; }
    public string ScannedAtUtc { get; set; }
  }

  internal sealed class LightingPlanRoomSnapshot
  {
    public string Id { get; set; }
    public string Handle { get; set; }
    public string Name { get; set; }
    public string RoomType { get; set; }
    public string Layer { get; set; }
    public List<LightingPlanPoint> Boundary { get; set; } = new List<LightingPlanPoint>();
    public double Area { get; set; }
  }

  internal sealed class LightingPlanFixtureSnapshot
  {
    public string Id { get; set; }
    public string Handle { get; set; }
    public string BlockName { get; set; }
    public string Layer { get; set; }
    public LightingPlanPoint Position { get; set; }
    public double Rotation { get; set; }
    public Dictionary<string, string> Attributes { get; set; } = new Dictionary<string, string>();
    public string Mark { get; set; }
    public string Panel { get; set; }
    public string Circuit { get; set; }
    public string ControlZone { get; set; }
    public string Watts { get; set; }
  }

  internal sealed class LightingPlanGenerationInstructions
  {
    public string SchemaVersion { get; set; }
    public string Kind { get; set; }
    public string SourceFingerprint { get; set; }
    public string DrawingPath { get; set; }
    public string TagLayer { get; set; }
    public List<LightingPlanTagInstruction> Tags { get; set; } = new List<LightingPlanTagInstruction>();
  }

  internal sealed class LightingPlanTagInstruction
  {
    public string GenerationKey { get; set; }
    public string SourceHandle { get; set; }
    public string FixtureId { get; set; }
    public string Mark { get; set; }
    public string Panel { get; set; }
    public string Circuit { get; set; }
    public string ControlZone { get; set; }
    public string RoomId { get; set; }
    public string RoomName { get; set; }
    public string Text { get; set; }
    public string Layer { get; set; }
    public LightingPlanPoint Position { get; set; }
  }

  internal sealed class LightingPlanPoint
  {
    public double X { get; set; }
    public double Y { get; set; }
    public double Z { get; set; }

    internal static LightingPlanPoint From(Point3d point)
    {
      return new LightingPlanPoint { X = point.X, Y = point.Y, Z = point.Z };
    }

    internal Point3d ToPoint3d()
    {
      return new Point3d(X, Y, Z);
    }
  }

  internal sealed class LightingPlanTextCandidate
  {
    public string Layer { get; set; }
    public string Text { get; set; }
    public Point3d Position { get; set; }
  }

  internal sealed class LightingPlanApplyResult
  {
    public int Created { get; set; }
    public int Updated { get; set; }
    public int Removed { get; set; }
  }
}
