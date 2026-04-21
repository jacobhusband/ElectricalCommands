using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private const string KnBlockName = "KEYED_NOTE";
    private const string KnLayerName = "keyed-notes";
    private const short KnLayerColor = 2;
    private const string KnAttributeTag = "KEYNO";
    private const string KnAttributePrompt = "Enter No.";
    private const string KnTextStyleName = "ARIALNARROW-1-8";
    private const string KnTextStyleFont = "ARIALN.TTF";
    private const double KnAttributeHeight = 0.09375;

    [CommandMethod("KN", CommandFlags.Modal)]
    public void KeyedNoteCommand()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      ed.WriteMessage("\n[KN / C# jig]");

      if (db.TileMode)
      {
        ed.WriteMessage("\nKN requires a paperspace layout. Switch to a layout tab and run again.");
        return;
      }
      bool inViewportEditing =
        System.Convert.ToInt16(Application.GetSystemVariable("CVPORT")) > 1;

      try
      {
        EnsureKnLayer(db);
        ObjectId textStyleId = EnsureKnTextStyle(db);
        if (!EnsureKnBlockDefinition(db, textStyleId))
        {
          ed.WriteMessage($"\nFailed to prepare canonical {KnBlockName} block definition.");
          return;
        }
        DumpKnBlockDiagnostics(db, ed);
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nKN setup error: {ex.Message}");
        return;
      }

      string keyValue = PromptKnKeyValue(ed);
      if (keyValue == null)
      {
        ed.WriteMessage("\nKN canceled.");
        return;
      }

      double scaleDenom;
      if (inViewportEditing)
      {
        scaleDenom = ResolveViewportScaleDenominator(ed, db);
        if (scaleDenom <= 0.0)
        {
          ed.WriteMessage("\nKN canceled: Could not resolve active viewport scale.");
          return;
        }
        string ratioScale = FormatRatio(scaleDenom);
        TrySetCannoscale(ratioScale);
        ed.WriteMessage($"\nUsing viewport scale: {ratioScale} (BR scale factor: {scaleDenom})");
      }
      else
      {
        scaleDenom = 1.0;
        ed.WriteMessage("\nPlacing in paperspace at 1:1 (BR scale factor: 1).");
      }

      ObjectId blockDefId;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        if (!bt.Has(KnBlockName))
        {
          ed.WriteMessage($"\nKN failed: {KnBlockName} block is missing.");
          tr.Commit();
          return;
        }
        blockDefId = bt[KnBlockName];
        tr.Commit();
      }

      ObjectId oldLayer = db.Clayer;
      try
      {
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
          if (lt.Has(KnLayerName))
          {
            db.Clayer = lt[KnLayerName];
          }
          tr.Commit();
        }

        BlockReference br = new BlockReference(Point3d.Origin, blockDefId)
        {
          Layer = KnLayerName,
          ScaleFactors = new Scale3d(scaleDenom),
          Rotation = 0.0
        };

        PromptResult jigResult;
        KeyNoteJig jig = new KeyNoteJig(br, scaleDenom);
        try
        {
          while (true)
          {
            jigResult = ed.Drag(jig);
            if (jigResult.Status == PromptStatus.Keyword)
            {
              if (jig.ApplyKeyword(jigResult.StringResult))
              {
                ed.WriteMessage($"\nAnchor: {jig.CurrentAnchor}");
              }
              continue;
            }
            break;
          }
        }
        catch (System.Exception ex)
        {
          br.Dispose();
          ed.WriteMessage($"\nKN jig error: {ex.Message}");
          return;
        }

        if (jigResult.Status != PromptStatus.OK)
        {
          br.Dispose();
          ed.WriteMessage("\nKN canceled.");
          return;
        }

        try
        {
          using (Transaction tr = db.TransactionManager.StartTransaction())
          {
            BlockTableRecord space = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            space.AppendEntity(br);
            tr.AddNewlyCreatedDBObject(br, true);

            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blockDefId, OpenMode.ForRead);
            bool attributeUpdated = false;
            foreach (ObjectId id in btr)
            {
              AttributeDefinition ad = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
              if (ad == null || ad.Constant) continue;

              AttributeReference ar = new AttributeReference();
              ar.SetAttributeFromBlock(ad, br.BlockTransform);
              ar.TextString = keyValue;
              br.AttributeCollection.AppendAttribute(ar);
              tr.AddNewlyCreatedDBObject(ar, true);
              attributeUpdated = true;
            }

            if (!attributeUpdated)
            {
              ed.WriteMessage($"\nWarning: {KnBlockName} has no editable attributes to update.");
            }

            tr.Commit();
          }

          ed.WriteMessage($"\nInserted {KnBlockName} with value {keyValue} on layer {KnLayerName}.");
        }
        catch (System.Exception ex)
        {
          ed.WriteMessage($"\nKN placement error: {ex.Message}");
        }
      }
      finally
      {
        if (!oldLayer.IsNull)
        {
          try { db.Clayer = oldLayer; } catch { }
        }
      }
    }

    private static void EnsureKnLayer(Database db)
    {
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
        if (lt.Has(KnLayerName))
        {
          LayerTableRecord existing = (LayerTableRecord)tr.GetObject(lt[KnLayerName], OpenMode.ForWrite);
          if (existing.Color.ColorMethod != ColorMethod.ByAci || existing.Color.ColorIndex != KnLayerColor)
          {
            existing.Color = Color.FromColorIndex(ColorMethod.ByAci, KnLayerColor);
          }
        }
        else
        {
          lt.UpgradeOpen();
          LayerTableRecord ltr = new LayerTableRecord
          {
            Name = KnLayerName,
            Color = Color.FromColorIndex(ColorMethod.ByAci, KnLayerColor)
          };
          lt.Add(ltr);
          tr.AddNewlyCreatedDBObject(ltr, true);
        }
        tr.Commit();
      }
    }

    private static ObjectId EnsureKnTextStyle(Database db)
    {
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        TextStyleTable tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
        if (tst.Has(KnTextStyleName))
        {
          ObjectId existingId = tst[KnTextStyleName];
          tr.Commit();
          return existingId;
        }

        tst.UpgradeOpen();
        TextStyleTableRecord ts = new TextStyleTableRecord
        {
          Name = KnTextStyleName,
          FileName = KnTextStyleFont,
          BigFontFileName = string.Empty,
          XScale = 1.0,
          ObliquingAngle = 0.0,
          TextSize = 0.0
        };
        ObjectId id = tst.Add(ts);
        tr.AddNewlyCreatedDBObject(ts, true);
        tr.Commit();
        return id;
      }
    }

    private static bool EnsureKnBlockDefinition(Database db, ObjectId textStyleId)
    {
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        BlockTableRecord btr;

        if (bt.Has(KnBlockName))
        {
          btr = (BlockTableRecord)tr.GetObject(bt[KnBlockName], OpenMode.ForWrite);
          if (btr.Annotative != AnnotativeStates.False)
          {
            btr.Annotative = AnnotativeStates.False;
          }
          tr.Commit();
          return true;
        }

        bt.UpgradeOpen();
        btr = new BlockTableRecord
        {
          Name = KnBlockName,
          Origin = Point3d.Origin
        };
        bt.Add(btr);
        tr.AddNewlyCreatedDBObject(btr, true);

        Polyline hex = new Polyline();
        hex.AddVertexAt(0, new Point2d(-0.0625998, -0.108426), 0.0, 0.0, 0.0);
        hex.AddVertexAt(1, new Point2d(0.0625998, -0.108426), 0.0, 0.0, 0.0);
        hex.AddVertexAt(2, new Point2d(0.1252, 0.0), 0.0, 0.0, 0.0);
        hex.AddVertexAt(3, new Point2d(0.0625998, 0.108426), 0.0, 0.0, 0.0);
        hex.AddVertexAt(4, new Point2d(-0.0625998, 0.108426), 0.0, 0.0, 0.0);
        hex.AddVertexAt(5, new Point2d(-0.1252, 0.0), 0.0, 0.0, 0.0);
        hex.Closed = true;
        btr.AppendEntity(hex);
        tr.AddNewlyCreatedDBObject(hex, true);
        hex.Layer = KnLayerName;
        hex.ColorIndex = 256;

        AttributeDefinition ad = new AttributeDefinition
        {
          Tag = KnAttributeTag,
          Prompt = KnAttributePrompt,
          TextString = "XXX",
          Height = KnAttributeHeight,
          TextStyleId = textStyleId,
          HorizontalMode = TextHorizontalMode.TextCenter,
          VerticalMode = TextVerticalMode.TextVerticalMid,
          AlignmentPoint = Point3d.Origin
        };
        btr.AppendEntity(ad);
        tr.AddNewlyCreatedDBObject(ad, true);
        ad.Layer = KnLayerName;
        ad.ColorIndex = 256;
        ad.AdjustAlignment(db);

        tr.Commit();
        return true;
      }
    }

    private static void DumpKnBlockDiagnostics(Database db, Editor ed)
    {
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        if (!bt.Has(KnBlockName))
        {
          ed.WriteMessage($"\n[KN diag] block {KnBlockName} missing");
          tr.Commit();
          return;
        }

        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[KnBlockName], OpenMode.ForRead);
        ed.WriteMessage($"\n[KN diag] block {KnBlockName}: Annotative={btr.Annotative}, Origin={btr.Origin}, HasPreviewIcon={btr.HasPreviewIcon}");
        int count = 0;
        foreach (ObjectId id in btr)
        {
          Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
          if (ent == null) continue;
          count++;
          string extra = string.Empty;
          try { extra = $", Annotative={ent.Annotative}, Visible={ent.Visible}"; } catch { }
          ed.WriteMessage($"\n[KN diag]   entity {count}: {ent.GetType().Name} on layer '{ent.Layer}' color={ent.ColorIndex}{extra}");
        }
        ed.WriteMessage($"\n[KN diag] total entities: {count}");
        tr.Commit();
      }
    }

    private static string PromptKnKeyValue(Editor ed)
    {
      while (true)
      {
        PromptStringOptions pso = new PromptStringOptions("\nEnter keyed note number (e.g., 1, 2, 3): ")
        {
          AllowSpaces = false
        };
        PromptResult pr = ed.GetString(pso);
        if (pr.Status != PromptStatus.OK) return null;

        string value = pr.StringResult?.Trim() ?? string.Empty;
        if (value.Length == 0) return null;

        if (IsPositiveInteger(value)) return value;

        ed.WriteMessage("\nInvalid value. Enter a positive whole number such as 1, 2, or 3.");
      }
    }

    private static bool IsPositiveInteger(string value)
    {
      if (string.IsNullOrEmpty(value)) return false;
      foreach (char c in value)
      {
        if (c < '0' || c > '9') return false;
      }
      return int.TryParse(value, out int parsed) && parsed > 0;
    }

    private static double ResolveViewportScaleDenominator(Editor ed, Database db)
    {
      double customScale = 0.0;

      try
      {
        ObjectId vpId = ed.CurrentViewportObjectId;
        if (!vpId.IsNull)
        {
          using (Transaction tr = db.TransactionManager.StartTransaction())
          {
            Viewport vp = tr.GetObject(vpId, OpenMode.ForRead) as Viewport;
            if (vp != null && vp.CustomScale > 0.0)
            {
              customScale = vp.CustomScale;
            }
            tr.Commit();
          }
        }
      }
      catch { }

      if (customScale <= 0.0)
      {
        try
        {
          object v = Application.GetSystemVariable("VPSCALE");
          if (v != null) customScale = System.Convert.ToDouble(v);
        }
        catch { }
      }

      if (customScale <= 0.0) return 0.0;
      return 1.0 / customScale;
    }

    private static string FormatRatio(double scaleDenom)
    {
      int denomRounded = (int)Math.Round(scaleDenom);
      if (Math.Abs(scaleDenom - denomRounded) < 1e-6)
      {
        return $"1:{denomRounded}";
      }
      return $"1:{scaleDenom.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)}";
    }

    private static bool TrySetCannoscale(string ratio)
    {
      try
      {
        Application.SetSystemVariable("CANNOSCALE", ratio);
        return true;
      }
      catch
      {
        return false;
      }
    }
  }
}
