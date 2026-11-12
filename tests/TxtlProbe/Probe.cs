using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

// Only run in a disposable Core Console drawing. Exercises the production
// command with scripted prompts and inspects the entities it actually creates.
public static class TxtlProbe
{
  private static string _scenario;
  private static ObjectId _layer;

  [CommandMethod("TXTLFIXTURE")]
  public static void Fixture()
  {
    var doc = Application.DocumentManager.MdiActiveDocument;
    _scenario = doc.Editor.GetString("\nScenario: ").StringResult;
    var db = doc.Database;
    db.TileMode = true;
    Application.SetSystemVariable("OSMODE", 0);
    Application.SetSystemVariable("3DOSMODE", 0);
    Application.SetSystemVariable("ORTHOMODE", 0);
    Application.SetSystemVariable("SNAPMODE", 0);
    db.Dimscale = 23;
    doc.Editor.CurrentUserCoordinateSystem = _scenario == "Ucs"
      ? Matrix3d.Rotation(Math.PI / 4, Vector3d.ZAxis, Point3d.Origin)
      : Matrix3d.Identity;
    using (var tr = db.TransactionManager.StartTransaction())
    {
      var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
      foreach (ObjectId id in space) ((Entity)tr.GetObject(id, OpenMode.ForWrite)).Erase();
      var layers = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);
      if (!layers.Has("TXTL_PROBE"))
      {
        var layer = new LayerTableRecord { Name = "TXTL_PROBE", Color = Color.FromColorIndex(ColorMethod.ByAci, 3) };
        _layer = layers.Add(layer);
        tr.AddNewlyCreatedDBObject(layer, true);
      }
      else _layer = layers["TXTL_PROBE"];
      db.Clayer = _layer;
      var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
      if (nod.Contains("ACIES_ELECTRICAL_COMMAND_SETTINGS"))
      {
        var settings = (DBDictionary)tr.GetObject(nod.GetAt("ACIES_ELECTRICAL_COMMAND_SETTINGS"), OpenMode.ForWrite);
        if (settings.Contains("DRAWING_SCALE"))
        {
          ObjectId id = settings.GetAt("DRAWING_SCALE");
          settings.Remove("DRAWING_SCALE");
          tr.GetObject(id, OpenMode.ForWrite).Erase();
        }
      }
      tr.Commit();
    }
    if (_scenario != "MissingScale")
    {
      var assembly = Assembly.LoadFrom(Environment.GetEnvironmentVariable("ACIES_TXTL_ASSEMBLY"));
      assembly.GetType("ElectricalCommands.ElectricalDrawingSettingsStore")
        .GetMethod("WriteScale", BindingFlags.Static | BindingFlags.Public)
        .Invoke(null, new object[] { db, _scenario == "EighthScale" ? 0.125 : 0.25, "probe scale" });
    }
  }

  [CommandMethod("TXTLASSERT")]
  public static void Assert()
  {
    string report;
    try
    {
      var db = Application.DocumentManager.MdiActiveDocument.Database;
      using (var tr = db.TransactionManager.StartTransaction())
      {
        var texts = new List<MText>();
        var leaders = new List<Leader>();
        var space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
        foreach (ObjectId id in space)
        {
          Entity entity = (Entity)tr.GetObject(id, OpenMode.ForRead);
          if (entity is MText text) texts.Add(text);
          else if (entity is Leader leader) leaders.Add(leader);
          else throw new InvalidOperationException("Unexpected entity: " + entity.GetType().Name);
        }
        bool empty = _scenario == "MissingScale" || _scenario == "Blank" ||
          _scenario == "Cancel" || _scenario == "Undo";
        Require(texts.Count == (empty ? 0 : 1) && leaders.Count == (empty ? 0 : 1), "entity count");
        if (!empty)
        {
          var text = texts[0];
          var leader = leaders[0];
          double height = _scenario == "EighthScale" ? 9 : 4.5;
          Require(Math.Abs(text.TextHeight - height) < 1e-8, "TXT text height");
          Require(text.Text == "TEST {A} \\P", "literal text and spaces");
          Require(((TextStyleTableRecord)tr.GetObject(text.TextStyleId, OpenMode.ForRead)).Name == "ARIALNARROW-1-8", "TXT style");
          Require(text.LayerId == _layer && leader.LayerId == _layer && db.Clayer == _layer, "current layer preserved");
          bool vertical = _scenario == "North" || _scenario == "South";
          bool backward = _scenario == "West" || _scenario == "South";
          double rotation = vertical ? Math.PI / 2 : (_scenario == "Ucs" ? Math.PI / 4 : 0);
          Require(Math.Abs(text.Rotation - rotation) < 1e-8, "direction and UCS");
          Require(text.Attachment == (backward ? AttachmentPoint.MiddleRight : AttachmentPoint.MiddleLeft), "TXT attachment");
          Require(leader.HasArrowHead && !leader.IsSplined, "straight arrowed leader");
          Require(Math.Abs(leader.Dimasz - height) < 1e-8 && Math.Abs(leader.Dimscale - 1) < 1e-8, "arrow size ignores DIMSCALE");
          Require(leader.NumVertices >= 2, "leader vertices");
          for (int i = 1; i < leader.NumVertices; i++)
            Require(leader.VertexAt(i).DistanceTo(leader.VertexAt(i - 1)) > 1e-7, "nonzero leader legs");
          if (_scenario == "Circle")
          {
            Require(!leader.Dimldrblk.IsNull && ((BlockTableRecord)tr.GetObject(leader.Dimldrblk, OpenMode.ForRead)).Name == "KN_LEADER_CIRCLE", "KNL circle end");
          }
          else Require(leader.Dimldrblk.IsNull, "standard arrow end");
          // The attachment is beyond the text's nearer edge and has a gap.
          Vector3d baseline = new Vector3d(Math.Cos(rotation), Math.Sin(rotation), 0);
          Vector3d across = Vector3d.ZAxis.CrossProduct(baseline);
          Vector3d offset = leader.VertexAt(leader.NumVertices - 1) - text.Location;
          double x = offset.DotProduct(baseline);
          double y = offset.DotProduct(across);
          double minX = backward ? -text.ActualWidth : 0;
          double maxX = backward ? 0 : text.ActualWidth;
          Require(x < minX - 1e-7 || x > maxX + 1e-7 || Math.Abs(y) > text.ActualHeight / 2 + 1e-7, "leader clears text");
        }
        tr.Commit();
      }
      report = "PASS: " + _scenario;
    }
    catch (System.Exception ex) { report = "FAIL: " + _scenario + ": " + ex; }
    File.AppendAllText(Environment.GetEnvironmentVariable("ACIES_TXTL_REPORT"), report + Environment.NewLine);
    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n" + report + "\n");
  }

  private static void Require(bool condition, string message)
  {
    if (!condition) throw new InvalidOperationException(message);
  }
}
