using System;
using System.IO;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;

public static class RoomTemplateProbe
{
  private static readonly BindingFlags Hidden = BindingFlags.Static | BindingFlags.NonPublic;
  private static Type Commands;
  private static Type Geometry;
  private static Type Jig;
  private static readonly StringBuilder Results = new StringBuilder();
  private static HashSet<ObjectId> Baseline;
  private static object CommandTemplate;

  [CommandMethod("RTVIEWPROBE")]
  public static void ViewProbe()
  {
    var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
    var report = new StringBuilder();
    try
    {
      var assembly = Assembly.LoadFrom(Environment.GetEnvironmentVariable("ACIES_RT_ASSEMBLY"));
      Commands = assembly.GetType("ElectricalCommands.GeneralCommands");
      using (var original = doc.Editor.GetCurrentView())
      {
        foreach (bool commit in new[] { false, true })
        {
          using (var transaction = doc.Database.TransactionManager.StartTransaction())
          {
            var table = (BlockTable)transaction.GetObject(doc.Database.BlockTableId, OpenMode.ForWrite);
            var space = new BlockTableRecord { Name = "*U" };
            table.Add(space); transaction.AddNewlyCreatedDBObject(space, true);
            var line = new Line(Point3d.Origin, new Point3d(1, 1, 0));
            space.AppendEntity(line); transaction.AddNewlyCreatedDBObject(line, true);
            ObjectId previewSpaceId = space.ObjectId;
            ObjectId previewEntityId = line.ObjectId;
            using (var moved = doc.Editor.GetCurrentView())
            {
              moved.CenterPoint = new Point2d(350, 450);
              moved.Width = 800; moved.Height = 600;
              doc.Editor.SetCurrentView(moved);
            }
            if (commit)
            {
              Commands.GetMethod("CleanupTemplatePreview", Hidden).Invoke(null, new object[] { transaction, space });
              Require(previewSpaceId.IsErased && previewEntityId.IsErased, "temporary preview objects erased");
              transaction.Commit();
            }
          }
          using (var view = doc.Editor.GetCurrentView())
          {
            report.AppendLine((commit ? "COMMIT" : "ABORT") + ": center=" + view.CenterPoint + " width=" + view.Width + " height=" + view.Height);
            if (commit) Require((view.CenterPoint - new Point2d(350, 450)).Length < 1e-6, "preview cleanup preserves latest view center");
          }
          doc.Editor.SetCurrentView(original);
        }
      }
      report.AppendLine("PASS: production preview cleanup erases temporary objects and preserves current view");
    }
    catch (System.Exception ex) { report.AppendLine("FAIL: " + ex); }
    File.WriteAllText(Path.Combine(Environment.GetEnvironmentVariable("ACIES_RT_OUTPUT"), "view.txt"), report.ToString());
  }

  [CommandMethod("RTCOMMANDSETUP")]
  public static void CommandSetup()
  {
    var assembly = Assembly.LoadFrom(Environment.GetEnvironmentVariable("ACIES_RT_ASSEMBLY"));
    Commands = assembly.GetType("ElectricalCommands.GeneralCommands");
    Geometry = assembly.GetType("ElectricalCommands.RoomTemplateGeometry");
    Jig = assembly.GetType("ElectricalCommands.RoomTemplateGroupJig");
    var db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
    var store = assembly.GetType("ElectricalCommands.RoomTemplateStore");
    CommandTemplate = ((IList)store.GetMethod("ReadAll", Hidden).Invoke(null, new object[] { db }))[0];
    Set(CommandTemplate, "Name", "RT_PROBE");
    ((IList)Get(CommandTemplate, "CircuitLabels")).Clear();
    var groups = (IList)Get(CommandTemplate, "Groups");
    // Keep command-flow expectations independent of later edits to the user's
    // saved template. The geometry round-trip tests still use its actual groups.
    groups.Clear();
    Type groupType = groups.GetType().GetGenericArguments()[0];
    groups.Add(JsonConvert.DeserializeObject("{\"Members\":[{\"Kind\":1,\"BlockName\":\"DATA\",\"Layer\":\"EL-N-PWR\",\"ScaleRatio\":-1},{\"Kind\":2,\"BlockName\":\"KEYED_NOTE\",\"Layer\":\"keyed-notes\",\"ScaleRatio\":-1,\"ParentIndex\":0,\"X\":-0.4,\"Attributes\":{\"KEYNO\":\"T4\"}}]}", groupType));
    groups.Add(JsonConvert.DeserializeObject("{\"Members\":[{\"Kind\":0,\"BlockName\":\"RECEPT\",\"Layer\":\"EL-N-PWR\",\"ScaleRatio\":-1,\"VisibilityState\":\"DUPLEX\"}]}", groupType));
    foreach (object group in groups)
      foreach (object member in (IList)Get(group, "Members")) Set(member, "CircuitSlot", -1);
    string circuitScenario = Environment.GetEnvironmentVariable("ACIES_RT_CIRCUIT_TEST");
    if (!string.IsNullOrEmpty(circuitScenario))
    {
      ((IList)Get(CommandTemplate, "CircuitLabels")).Add("Circuit 1");
      foreach (object group in groups)
        foreach (object member in (IList)Get(group, "Members"))
          if (Get(member, "Kind").ToString() == "Receptacle") Set(member, "CircuitSlot", 0);
      // Only the disposable test drawing is changed. Placement and declining
      // circuiting must succeed even when no panel configuration is available.
      using (var tr = db.TransactionManager.StartTransaction())
      {
        var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
        if (nod.Contains("ACIES_ELECTRICAL_COMMAND_SETTINGS"))
        {
          var settings = (DBDictionary)tr.GetObject(nod.GetAt("ACIES_ELECTRICAL_COMMAND_SETTINGS"), OpenMode.ForWrite);
          foreach (string key in new[] { "PANEL_NAME", "PANEL_SCHEDULE" })
            if (settings.Contains(key))
            {
              ObjectId id = settings.GetAt(key);
              settings.Remove(key);
              tr.GetObject(id, OpenMode.ForWrite).Erase();
            }
        }
        tr.Commit();
      }
    }
    store.GetMethod("Save", Hidden).Invoke(null, new[] { db, CommandTemplate });
    using (var tr = db.TransactionManager.StartTransaction())
    {
      var table = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
      var model = (BlockTableRecord)tr.GetObject(table[BlockTableRecord.ModelSpace], OpenMode.ForRead);
      Baseline = new HashSet<ObjectId>(model.Cast<ObjectId>());
    }
    var app = typeof(Autodesk.AutoCAD.ApplicationServices.Core.Application);
    Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("OSMODE", 0);
    Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("SNAPMODE", 0);
    Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("ORTHOMODE", 0);
  }

  [CommandMethod("RTCOMMANDASSERT")]
  public static void CommandAssert()
  {
    var db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
    var report = new StringBuilder();
    try
    {
      var groups = (IList)Get(CommandTemplate, "Groups");
      using (var tr = db.TransactionManager.StartTransaction())
      {
        var table = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        var model = (BlockTableRecord)tr.GetObject(table[BlockTableRecord.ModelSpace], OpenMode.ForRead);
        var inserted = model.Cast<ObjectId>().Where(id => !Baseline.Contains(id) && !id.IsErased)
          .Select(id => (BlockReference)tr.GetObject(id, OpenMode.ForRead)).ToList();
        Require(inserted.Count == 3, "RTP/Back commits exactly two groups with three blocks; actual=" + inserted.Count);
        for (int index = 0; index < 2; index++)
        {
          var resolved = (IList)Geometry.GetMethod("Resolve", Hidden).Invoke(null,
            new object[] { groups[index], new Point3d(100 + 100 * index, 100 + 100 * index, 0), 0.0, 0.0, (double)Get(CommandTemplate, "CaptureBlockScale") });
          foreach (object item in resolved)
          {
            Point3d point = (Point3d)Get(item, "Position");
            var block = inserted.Single(candidate => candidate.Position.DistanceTo(point) < 1e-6);
            Require(block.ScaleFactors.Equals((Scale3d)Get(item, "ScaleFactors")), "RTP commits saved signed scales");
          }
        }
      }
      report.AppendLine("PASS: actual RTP command, Back replacement, preview rollback and final insertion");
    }
    catch (System.Exception ex) { report.AppendLine("FAIL: " + ex); }
    string scenario = Environment.GetEnvironmentVariable("ACIES_RT_CIRCUIT_TEST");
    string reportName = string.IsNullOrEmpty(scenario) ? "command" : "circuit-" + scenario;
    File.WriteAllText(Path.Combine(Environment.GetEnvironmentVariable("ACIES_RT_OUTPUT"), reportName + ".txt"), report.ToString());
  }

  [CommandMethod("RTVERIFY")]
  public static void Verify()
  {
    Results.Clear();
    var previous = HostApplicationServices.WorkingDatabase;
    try
    {
      var assembly = Assembly.LoadFrom(Environment.GetEnvironmentVariable("ACIES_RT_ASSEMBLY"));
      Commands = assembly.GetType("ElectricalCommands.GeneralCommands");
      Geometry = assembly.GetType("ElectricalCommands.RoomTemplateGeometry");
      Jig = assembly.GetType("ElectricalCommands.RoomTemplateGroupJig");
      using (var db = new Database(false, true))
      {
        db.ReadDwgFile(Environment.GetEnvironmentVariable("ACIES_RT_SOURCE"), FileOpenMode.OpenForReadAndAllShare, false, null);
        db.CloseInput(true);
        HostApplicationServices.WorkingDatabase = db;
        var store = assembly.GetType("ElectricalCommands.RoomTemplateStore");
        var templates = (IList)store.GetMethod("ReadAll", Hidden).Invoke(null, new object[] { db });
        Require(templates.Count > 0, "real drawing has saved templates");
        foreach (object template in templates)
        {
          var groups = (IList)Get(template, "Groups");
          foreach (object group in groups)
            foreach (double angle in new[] { 0.0, Math.PI / 2, 0.6457718232379019 })
              VerifyGroup(db, group, (double)Get(template, "CaptureBlockScale"), angle, true);
        }
        Results.AppendLine("PASS: saved CLASSROOM groups at 0, 90 and 37 degrees; attributes, visibility, dynamic properties and preview/final parity");

        // Capture actual mirrored and dynamic references, then reconstruct them at
        // their original frame. All modifications are confined to a read-only load.
        var originals = new List<ObjectId>();
        using (var tr = db.TransactionManager.StartTransaction())
        {
          var table = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
          var model = (BlockTableRecord)tr.GetObject(table[BlockTableRecord.ModelSpace], OpenMode.ForRead);
          var found = new HashSet<string>();
          foreach (ObjectId id in model)
          {
            if (!(tr.GetObject(id, OpenMode.ForRead) is BlockReference block)) continue;
            var definition = (BlockTableRecord)tr.GetObject(block.IsDynamicBlock ? block.DynamicBlockTableRecord : block.BlockTableRecord, OpenMode.ForRead);
            if (!new[] { "KEYED_NOTE", "DATA", "RECEPT", "Receptacles", "os" }.Contains(definition.Name)) continue;
            string key = definition.Name + (block.ScaleFactors.X < 0 ? " mirrored" : " positive");
            if (found.Add(key)) originals.Add(id);
          }
        }
        object receptacles = Activator.CreateInstance(typeof(List<>).MakeGenericType(Commands.GetNestedType("CapturedReceptacle", BindingFlags.NonPublic)));
        object captured = Commands.GetMethod("BuildTemplateGroup", Hidden).Invoke(null,
          new object[] { db, originals, "Round trip", new Point3d(10, 20, 0), null, 0.4, 96.0, receptacles });
        var serialized = JsonConvert.SerializeObject(captured);
        captured = JsonConvert.DeserializeObject(serialized, captured.GetType());
        var resolved = (IList)Geometry.GetMethod("Resolve", Hidden).Invoke(null,
          new object[] { captured, new Point3d(10, 20, 0), 0.4, 0.0, 96.0 });
        using (var tr = db.TransactionManager.StartTransaction())
        {
          var table = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
          var owner = new BlockTableRecord { Name = "*U" };
          table.Add(owner); tr.AddNewlyCreatedDBObject(owner, true);
          foreach (object item in resolved)
          {
            object member = Get(item, "Member");
            string name = (string)Get(member, "BlockName");
            var block = Create(db, tr, owner, table[name], item);
            BlockReference source = originals.Select(id => (BlockReference)tr.GetObject(id, OpenMode.ForRead))
              .First(candidate => candidate.Position.DistanceTo(block.Position) < 1e-6);
            Require(MatrixNear(block.BlockTransform, source.BlockTransform), name + " captured transform");
            CompareAttributes(tr, source, block, name + " captured attributes");
            CompareDynamic(source, block);
          }
        }
        Results.AppendLine("PASS: capture/JSON/load/placement round trip of " + originals.Count + " actual positive and mirrored blocks, including readable MIRRTEXT attributes");
        VerifyGroup(db, captured, 96, Math.PI / 2, false);
        VerifyLeader(assembly);
        VerifyData(assembly, db);
      }
      Results.AppendLine("ALL TESTS PASSED");
    }
    catch (System.Exception ex) { Results.AppendLine("FAIL: " + ex); }
    finally
    {
      HostApplicationServices.WorkingDatabase = previous;
      File.WriteAllText(Path.Combine(Environment.GetEnvironmentVariable("ACIES_RT_OUTPUT"), "verify.txt"), Results.ToString());
    }
  }

  private static void VerifyGroup(Database db, object group, double scale, double angle, bool legacy)
  {
    var resolved = (IList)Geometry.GetMethod("Resolve", Hidden).Invoke(null,
      new object[] { group, new Point3d(100, 200, 0), 0.0, angle, scale });
    using (var tr = db.TransactionManager.StartTransaction())
    {
      var table = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
      var owner = new BlockTableRecord { Name = "*U" };
      table.Add(owner); tr.AddNewlyCreatedDBObject(owner, true);
      foreach (object item in resolved)
      {
        object member = Get(item, "Member");
        if (Get(member, "Kind").ToString() == "Leader") continue;
        string name = (string)Get(member, "BlockName");
        ObjectId definition = table[name];
        var final = Create(db, tr, owner, definition, item);
        object neutral = Activator.CreateInstance(item.GetType());
        Set(neutral, "Member", member); Set(neutral, "Position", Point3d.Origin);
        Set(neutral, "Rotation", 0.0); Set(neutral, "ScaleFactors", new Scale3d(1));
        var preview = Create(db, tr, owner, definition, neutral);
        Jig.GetMethod("UpdatePreview", Hidden).Invoke(null, new[] { preview, item });
        Require(MatrixNear(preview.BlockTransform, final.BlockTransform), name + " preview transform");
        CompareAttributes(tr, final, preview, name + " preview attributes");
        CompareDynamic(final, preview);
        // A second update must not accumulate transformations or move attributes.
        Jig.GetMethod("UpdatePreview", Hidden).Invoke(null, new[] { preview, item });
        CompareAttributes(tr, final, preview, name + " repeated preview");
        if (legacy && Get(member, "Kind").ToString() == "Note" && Get(member, "ScaleRatioY") == null)
        {
          foreach (ObjectId id in final.AttributeCollection)
          {
            var ar = (AttributeReference)tr.GetObject(id, OpenMode.ForRead);
            Require(!ar.IsMirroredInX && !ar.IsMirroredInY, "legacy keyed note readable");
            Require(ar.AlignmentPoint.DistanceTo(final.Position) < 1e-6, "legacy keyed note centered");
          }
        }
      }
      Commands.GetMethod("CleanupTemplatePreview", Hidden).Invoke(null, new object[] { tr, owner });
      Require(owner.IsErased, "preview definition cleaned up");
      tr.Commit();
    }
  }

  private static void VerifyLeader(Assembly assembly)
  {
    string json = "{\"Members\":[{\"Kind\":1,\"X\":3,\"Y\":4},{\"Kind\":2,\"ParentIndex\":0,\"X\":0.4,\"Y\":0.1},{\"Kind\":3,\"ParentIndex\":0,\"NoteIndex\":1,\"LeaderPath\":[{\"X\":0,\"Y\":0},{\"X\":0.2,\"Y\":0.3},{\"X\":0.35,\"Y\":0.1}]}]}";
    object group = JsonConvert.DeserializeObject(json, assembly.GetType("ElectricalCommands.TemplateGroup"));
    var resolved = (IList)Geometry.GetMethod("Resolve", Hidden).Invoke(null,
      new object[] { group, Point3d.Origin, 0.0, 0.0, 10.0 });
    Require(((Point3d)Get(resolved[1], "Position")).DistanceTo(new Point3d(7, 5, 0)) < 1e-8, "saved note center unchanged");
    var path = (Point3d[])Get(resolved[2], "LeaderPath");
    Require(path.Length == 3 && path[1].DistanceTo(new Point3d(5, 7, 0)) < 1e-8, "saved leader elbow unchanged");
    Results.AppendLine("PASS: captured leader path and note center preserved");
  }

  private static void VerifyData(Assembly assembly, Database db)
  {
    Type data = assembly.GetType("ElectricalCommands.DataGeometry");
    Type sideType = assembly.GetType("ElectricalCommands.SymbolSide");
    var expected = new Dictionary<string, Vector3d> { { "North", Vector3d.YAxis },
      { "East", Vector3d.XAxis }, { "South", -Vector3d.YAxis }, { "West", -Vector3d.XAxis } };
    string name = (string)Commands.GetField("DataBlockName", Hidden).GetRawConstantValue();
    Require(name != "DATA", "new DATA definition avoids imported dynamic block collision");
    using (var tr = db.TransactionManager.StartTransaction())
    {
      object north = Enum.Parse(sideType, "North");
      var geometry = (List<Entity>)data.GetMethod("CreateEntities", Hidden).Invoke(null, new[] { db, north });
      var table = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
      var definition = new BlockTableRecord { Name = name + "_PROBE" };
      ObjectId definitionId = table.Add(definition); tr.AddNewlyCreatedDBObject(definition, true);
      foreach (var entity in geometry) { definition.AppendEntity(entity); tr.AddNewlyCreatedDBObject(entity, true); }
      foreach (var pair in expected)
      {
        object side = Enum.Parse(sideType, pair.Key);
        double rotation = (double)data.GetMethod("ResolveRotation", Hidden).Invoke(null, new[] { side });
        var preview = (List<Entity>)data.GetMethod("CreateEntities", Hidden).Invoke(null, new[] { db, side });
        using (var block = new BlockReference(Point3d.Origin, definitionId) { Rotation = rotation })
        {
          var triangle = (Polyline)geometry[0];
          Point3d apex = triangle.GetPoint3dAt(2).TransformBy(block.BlockTransform);
          Require((apex - Point3d.Origin).GetNormal().DotProduct(pair.Value) > 0.999999, "DATA " + pair.Key + " direction");
          for (int i = 0; i < 3; i++)
            Require(((Polyline)preview[0]).GetPoint3dAt(i).DistanceTo(triangle.GetPoint3dAt(i).TransformBy(block.BlockTransform)) < 1e-8, "DATA preview/insertion parity");
          foreach (Entity entity in preview) entity.Dispose();
        }
        if (pair.Key == "North") Require(rotation == 0, "DATA zero points north");
        if (pair.Key == "South") Require(Math.Abs(rotation - Math.PI) < 1e-8, "DATA 180 points south");
      }
    }
    Results.AppendLine("PASS: DATA cardinal directions, 0 north/180 south, preview parity and imported-definition isolation");
  }

  private static BlockReference Create(Database db, Transaction tr, BlockTableRecord owner, ObjectId definition, object item)
    => (BlockReference)Commands.GetMethod("CreateTemplateBlock", Hidden).Invoke(null, new[] { db, tr, owner, definition, item });
  private static object Get(object obj, string name) => obj.GetType().GetProperty(name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).GetValue(obj);
  private static void Set(object obj, string name, object value) => obj.GetType().GetProperty(name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).SetValue(obj, value);
  private static bool MatrixNear(Matrix3d a, Matrix3d b) => a.ToArray().Zip(b.ToArray(), (x, y) => Math.Abs(x - y) < 1e-6).All(x => x);
  private static void Require(bool ok, string message) { if (!ok) throw new System.Exception(message); }
  private static void CompareAttributes(Transaction tr, BlockReference a, BlockReference b, string context)
  {
    Require(a.AttributeCollection.Count == b.AttributeCollection.Count, context + " count");
    var originals = a.AttributeCollection.Cast<ObjectId>().Select(id => (AttributeReference)tr.GetObject(id, OpenMode.ForRead)).ToDictionary(ar => ar.Tag);
    foreach (ObjectId id in b.AttributeCollection)
    {
      var actual = (AttributeReference)tr.GetObject(id, OpenMode.ForRead);
      var expected = originals[actual.Tag];
      Require(actual.TextString == expected.TextString, context + " value");
      Require(actual.Position.DistanceTo(expected.Position) < 1e-6, context + " tag=" + actual.Tag + " position " + actual.Position + " expected " + expected.Position
        + " alignment=" + actual.AlignmentPoint + "/" + expected.AlignmentPoint + " style=" + actual.TextStyleId.Handle + "/" + expected.TextStyleId.Handle
        + " mtext=" + actual.IsMTextAttribute + "/" + expected.IsMTextAttribute + " mirror=" + actual.IsMirroredInX + "," + actual.IsMirroredInY + "/" + expected.IsMirroredInX + "," + expected.IsMirroredInY);
      if (!actual.IsDefaultAlignment) Require(actual.AlignmentPoint.DistanceTo(expected.AlignmentPoint) < 1e-6, context + " alignment");
      Require(Math.Abs(actual.Height - expected.Height) < 1e-6, context + " height");
      Require((actual.Normal - expected.Normal).Length < 1e-6, context + " normal");
      Require(Math.Abs(Math.Sin((actual.Rotation - expected.Rotation) / 2)) < 1e-6, context + " rotation");
      Require(actual.IsMirroredInX == expected.IsMirroredInX && actual.IsMirroredInY == expected.IsMirroredInY, context + " mirroring");
    }
  }
  private static void CompareDynamic(BlockReference a, BlockReference b)
  {
    if (!a.IsDynamicBlock) return;
    var expected = a.DynamicBlockReferencePropertyCollection.Cast<DynamicBlockReferenceProperty>().ToDictionary(p => p.PropertyName, p => p.Value);
    foreach (DynamicBlockReferenceProperty p in b.DynamicBlockReferencePropertyCollection)
    {
      if (p.ReadOnly || !(p.Value is string || p.Value is double || p.Value is int || p.Value is short)) continue;
      bool same = p.Value is double number && expected[p.PropertyName] is double original
        ? Math.Abs(number - original) < 1e-8
        : Convert.ToString(p.Value) == Convert.ToString(expected[p.PropertyName]);
      Require(same, "dynamic " + p.PropertyName + " actual=" + p.Value + " expected=" + expected[p.PropertyName]);
    }
  }

  [CommandMethod("RTINSPECT")]
  public static void Inspect()
  {
    var report = new StringBuilder();
    using (var db = new Database(false, true))
    {
      db.ReadDwgFile(Environment.GetEnvironmentVariable("ACIES_RT_SOURCE"), FileOpenMode.OpenForReadAndAllShare, false, null);
      db.CloseInput(true);
      using (var tr = db.TransactionManager.StartTransaction())
      {
        var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
        if (nod.Contains("ACIES_ROOM_TEMPLATES"))
        {
          var templates = (DBDictionary)tr.GetObject(nod.GetAt("ACIES_ROOM_TEMPLATES"), OpenMode.ForRead);
          foreach (DBDictionaryEntry entry in templates)
          {
            report.AppendLine("TEMPLATE " + entry.Key);
            var record = (Xrecord)tr.GetObject(entry.Value, OpenMode.ForRead);
            foreach (var value in record.Data.AsArray()) if (value.Value is string text) report.Append(text);
            report.AppendLine();
          }
        }
        var table = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        if (table.Has("KEYED_NOTE"))
        {
          var definition = (BlockTableRecord)tr.GetObject(table["KEYED_NOTE"], OpenMode.ForRead);
          report.AppendLine("NOTE ORIGIN " + definition.Origin);
          foreach (ObjectId id in definition)
            if (tr.GetObject(id, OpenMode.ForRead) is AttributeDefinition ad)
              report.AppendLine("DEF " + ad.Tag + " pos=" + ad.Position + " align=" + ad.AlignmentPoint + " rot=" + ad.Rotation + " normal=" + ad.Normal + " mirrored=" + ad.IsMirroredInX + "," + ad.IsMirroredInY);
        }
        var model = (BlockTableRecord)tr.GetObject(table[BlockTableRecord.ModelSpace], OpenMode.ForRead);
        foreach (ObjectId id in model)
        {
          if (!(tr.GetObject(id, OpenMode.ForRead) is BlockReference block)) continue;
          var definition = (BlockTableRecord)tr.GetObject(block.IsDynamicBlock ? block.DynamicBlockTableRecord : block.BlockTableRecord, OpenMode.ForRead);
          if (definition.Name != "KEYED_NOTE") continue;
          report.AppendLine("NOTE " + id.Handle + " pos=" + block.Position + " scale=" + block.ScaleFactors + " rot=" + block.Rotation + " normal=" + block.Normal);
          foreach (ObjectId attributeId in block.AttributeCollection)
          {
            var ar = (AttributeReference)tr.GetObject(attributeId, OpenMode.ForRead);
            report.AppendLine("  ATTR " + ar.TextString + " pos=" + ar.Position + " align=" + ar.AlignmentPoint + " rot=" + ar.Rotation + " normal=" + ar.Normal + " mirrored=" + ar.IsMirroredInX + "," + ar.IsMirroredInY);
          }
        }
      }
    }
    File.WriteAllText(Path.Combine(Environment.GetEnvironmentVariable("ACIES_RT_OUTPUT"), "inspect.txt"), report.ToString());
  }
}
