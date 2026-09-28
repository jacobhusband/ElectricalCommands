using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace ElectricalCommands
{
  public static class SwitchPlacementEngine
  {
    public static bool CaptureFromSelection(
      Editor ed,
      Database db,
      out SwitchOrientationConfig config,
      out string error)
    {
      config = null;
      error = string.Empty;

      PromptSelectionOptions selOpts = new PromptSelectionOptions
      {
        MessageForAdding = "\nSelect switch block and associated text object(s): ",
        AllowDuplicates = false
      };

      PromptSelectionResult selRes = ed.GetSelection(selOpts);
      if (selRes.Status != PromptStatus.OK || selRes.Value == null || selRes.Value.Count == 0)
      {
        error = "Selection cancelled or empty.";
        return false;
      }

      BlockReference selectedBlock = null;
      List<ObjectId> textIds = new List<ObjectId>();

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        foreach (SelectedObject so in selRes.Value)
        {
          if (so == null) continue;
          Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
          if (ent is BlockReference br)
          {
            if (selectedBlock == null)
            {
              selectedBlock = br;
            }
          }
          else if (ent is DBText || ent is MText)
          {
            textIds.Add(ent.ObjectId);
          }
        }

        if (selectedBlock == null)
        {
          error = "No block reference was selected. Please select a switch block and any associated text.";
          return false;
        }

        Point3d defaultBasePoint = selectedBlock.Position;
        PromptPointOptions ptOpts = new PromptPointOptions("\nSpecify insertion base point for switch [Press Enter to use block insertion point]: ")
        {
          AllowNone = true,
          UseBasePoint = false
        };

        PromptPointResult ptRes = ed.GetPoint(ptOpts);
        Point3d basePoint;
        if (ptRes.Status == PromptStatus.None || ptRes.Status != PromptStatus.OK)
        {
          basePoint = defaultBasePoint;
        }
        else
        {
          Matrix3d ucsToWcs = ed.CurrentUserCoordinateSystem;
          basePoint = ptRes.Value.TransformBy(ucsToWcs);
        }

        // Capture Block definition
        string effectiveBlockName = GetEffectiveBlockName(selectedBlock, tr);
        var blockDef = new SwitchBlockDefinition
        {
          BlockName = effectiveBlockName,
          Rotation = selectedBlock.Rotation,
          ScaleX = selectedBlock.ScaleFactors.X,
          ScaleY = selectedBlock.ScaleFactors.Y,
          ScaleZ = selectedBlock.ScaleFactors.Z,
          Layer = selectedBlock.Layer,
          VisibilityState = GetVisibilityState(selectedBlock)
        };

        // Capture attributes
        if (selectedBlock.AttributeCollection != null)
        {
          foreach (ObjectId attId in selectedBlock.AttributeCollection)
          {
            if (attId.IsNull) continue;
            var att = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
            if (att != null)
            {
              blockDef.Attributes.Add(new SwitchAttributeDefinition
              {
                Tag = att.Tag,
                TextString = att.TextString,
                RelativeOffset = StoredVector3d.FromVector3d(att.Position - basePoint)
              });
            }
          }
        }

        // Capture Text objects
        var textDefs = new List<SwitchTextDefinition>();
        foreach (ObjectId txtId in textIds)
        {
          Entity ent = tr.GetObject(txtId, OpenMode.ForRead) as Entity;
          if (ent is DBText dt)
          {
            Point3d textPos = (dt.HorizontalMode == TextHorizontalMode.TextLeft && dt.VerticalMode == TextVerticalMode.TextBase)
              ? dt.Position
              : dt.AlignmentPoint;

            string styleName = GetTextStyleName(db, dt.TextStyleId, tr);
            textDefs.Add(new SwitchTextDefinition
            {
              IsMText = false,
              TextString = dt.TextString,
              TextStyleName = styleName,
              Height = dt.Height,
              Rotation = dt.Rotation,
              Layer = dt.Layer,
              ColorIndex = dt.Color.ColorIndex,
              HorizontalMode = (int)dt.HorizontalMode,
              VerticalMode = (int)dt.VerticalMode,
              AttachmentPoint = (int)dt.Justify,
              RelativeOffset = StoredVector3d.FromVector3d(textPos - basePoint)
            });
          }
          else if (ent is MText mt)
          {
            string styleName = GetTextStyleName(db, mt.TextStyleId, tr);
            textDefs.Add(new SwitchTextDefinition
            {
              IsMText = true,
              TextString = mt.Contents,
              TextStyleName = styleName,
              Height = mt.TextHeight,
              Rotation = mt.Rotation,
              Layer = mt.Layer,
              ColorIndex = mt.Color.ColorIndex,
              Width = mt.Width,
              AttachmentPoint = (int)mt.Attachment,
              RelativeOffset = StoredVector3d.FromVector3d(mt.Location - basePoint)
            });
          }
        }

        config = new SwitchOrientationConfig
        {
          Block = blockDef,
          TextObjects = textDefs
        };

        tr.Commit();
        return true;
      }
    }

    public static bool PlaceSwitch(
      Database db,
      Editor ed,
      ProjectSwitchSettings projectSettings,
      SwitchType type,
      SwitchOrientation orientation,
      Point3d ucsPoint,
      string customText = null,
      string overrideSubscript = null)
    {
      if (db == null || ed == null || projectSettings == null)
      {
        return false;
      }

      var typeConfig = projectSettings.GetTypeConfig(type);
      var orientConfig = typeConfig.GetOrientation(orientation);

      if (orientConfig == null || !orientConfig.IsConfigured)
      {
        ed.WriteMessage($"\n{typeConfig.DisplayName} ({orientation}) is not configured yet. Run SWSETUP to configure.");
        return false;
      }

      Matrix3d ucsToWcs = ed.CurrentUserCoordinateSystem;
      Point3d wcsPoint = ucsPoint.TransformBy(ucsToWcs);

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        string blockName = orientConfig.Block.BlockName;

        if (!bt.Has(blockName))
        {
          // Attempt to clone from source drawing
          if (!string.IsNullOrWhiteSpace(projectSettings.SourceDrawing) &&
              SwitchConfigurationStore.TryImportBlockDefinition(db, projectSettings.SourceDrawing, blockName, out _))
          {
            bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
          }

          if (!bt.Has(blockName))
          {
            ed.WriteMessage($"\nBlock definition '{blockName}' not found in drawing.");
            return false;
          }
        }

        ObjectId blockDefId = bt[blockName];
        BlockTableRecord currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

        // 1. Create and insert BlockReference
        BlockReference blockRef = new BlockReference(wcsPoint, blockDefId);
        blockRef.SetDatabaseDefaults(db);

        if (!string.IsNullOrWhiteSpace(orientConfig.Block.Layer) && LayerExists(db, orientConfig.Block.Layer, tr))
        {
          blockRef.Layer = orientConfig.Block.Layer;
        }

        blockRef.ScaleFactors = new Scale3d(
          orientConfig.Block.ScaleX,
          orientConfig.Block.ScaleY,
          orientConfig.Block.ScaleZ);

        blockRef.Rotation = orientConfig.Block.Rotation;

        currentSpace.AppendEntity(blockRef);
        tr.AddNewlyCreatedDBObject(blockRef, true);

        // Apply dynamic visibility if applicable
        if (!string.IsNullOrWhiteSpace(orientConfig.Block.VisibilityState))
        {
          SetVisibilityState(blockRef, orientConfig.Block.VisibilityState);
        }

        // Add attributes from block definition
        AddAttributes(tr, blockDefId, blockRef, orientConfig.Block.Attributes);

        // 2. Insert associated Text Objects
        if (orientConfig.TextObjects != null && orientConfig.TextObjects.Count > 0)
        {
          for (int i = 0; i < orientConfig.TextObjects.Count; i++)
          {
            var txtDef = orientConfig.TextObjects[i];
            Vector3d offset = txtDef.RelativeOffset.ToVector3d();
            Point3d textLoc = wcsPoint + offset;

            string textContent = txtDef.TextString;
            if (!string.IsNullOrWhiteSpace(overrideSubscript))
            {
              textContent = overrideSubscript;
            }
            else if (!string.IsNullOrWhiteSpace(customText))
            {
              textContent = customText;
            }

            ObjectId textStyleId = ResolveTextStyleId(db, txtDef.TextStyleName, tr);

            if (txtDef.IsMText)
            {
              MText mtext = new MText
              {
                Contents = textContent,
                Location = textLoc,
                TextHeight = txtDef.Height,
                Rotation = txtDef.Rotation,
                Attachment = (AttachmentPoint)txtDef.AttachmentPoint
              };
              mtext.SetDatabaseDefaults(db);
              if (!textStyleId.IsNull) mtext.TextStyleId = textStyleId;
              if (!string.IsNullOrWhiteSpace(txtDef.Layer) && LayerExists(db, txtDef.Layer, tr))
              {
                mtext.Layer = txtDef.Layer;
              }
              if (txtDef.ColorIndex > 0 && txtDef.ColorIndex < 256)
              {
                mtext.Color = Color.FromColorIndex(ColorMethod.ByAci, (short)txtDef.ColorIndex);
              }
              currentSpace.AppendEntity(mtext);
              tr.AddNewlyCreatedDBObject(mtext, true);
            }
            else
            {
              DBText dbText = new DBText
              {
                TextString = textContent,
                Height = txtDef.Height,
                Rotation = txtDef.Rotation,
                HorizontalMode = (TextHorizontalMode)txtDef.HorizontalMode,
                VerticalMode = (TextVerticalMode)txtDef.VerticalMode
              };
              dbText.SetDatabaseDefaults(db);
              if (!textStyleId.IsNull) dbText.TextStyleId = textStyleId;
              if (!string.IsNullOrWhiteSpace(txtDef.Layer) && LayerExists(db, txtDef.Layer, tr))
              {
                dbText.Layer = txtDef.Layer;
              }
              if (txtDef.ColorIndex > 0 && txtDef.ColorIndex < 256)
              {
                dbText.Color = Color.FromColorIndex(ColorMethod.ByAci, (short)txtDef.ColorIndex);
              }

              if (dbText.HorizontalMode == TextHorizontalMode.TextLeft && dbText.VerticalMode == TextVerticalMode.TextBase)
              {
                dbText.Position = textLoc;
              }
              else
              {
                dbText.Position = textLoc;
                dbText.AlignmentPoint = textLoc;
              }

              currentSpace.AppendEntity(dbText);
              tr.AddNewlyCreatedDBObject(dbText, true);
              dbText.AdjustAlignment(db);
            }
          }
        }

        tr.Commit();
        return true;
      }
    }

    public static void RunPlacementLoop(
      Editor ed,
      Database db,
      SwitchType type,
      SwitchOrientation orientation)
    {
      var projectSettings = SwitchConfigurationStore.LoadSettings(db);
      var typeConfig = projectSettings.GetTypeConfig(type);

      if (!typeConfig.GetOrientation(orientation).IsConfigured)
      {
        PromptKeywordOptions setupPrompt = new PromptKeywordOptions(
          $"\n{typeConfig.DisplayName} ({orientation}) is not configured yet. Would you like to configure it now? ")
        {
          AllowNone = true
        };
        setupPrompt.Keywords.Add("Yes");
        setupPrompt.Keywords.Add("No");
        setupPrompt.Keywords.Default = "Yes";

        var setupRes = ed.GetKeywords(setupPrompt);
        if (setupRes.Status == PromptStatus.OK && setupRes.StringResult == "Yes")
        {
          SwitchSetupWizard.RunSetup(ed, db, type, orientation);
          projectSettings = SwitchConfigurationStore.LoadSettings(db);
        }
        else
        {
          return;
        }
      }

      int placedCount = 0;
      string overrideText = null;

      while (true)
      {
        typeConfig = projectSettings.GetTypeConfig(type);
        var orientConfig = typeConfig.GetOrientation(orientation);

        string promptMsg = $"\nSpecify insertion point for {typeConfig.DisplayName} [{orientation}]" +
          (placedCount > 0 ? " <Enter to finish>" : "") +
          " or [North/East/South/West/Type/Subscript/Setup]: ";

        PromptPointOptions ptOpts = new PromptPointOptions(promptMsg)
        {
          AllowNone = true
        };

        ptOpts.Keywords.Add("North");
        ptOpts.Keywords.Add("East");
        ptOpts.Keywords.Add("South");
        ptOpts.Keywords.Add("West");
        ptOpts.Keywords.Add("Type");
        ptOpts.Keywords.Add("Subscript");
        ptOpts.Keywords.Add("Setup");

        PromptPointResult ptRes = ed.GetPoint(ptOpts);

        if (ptRes.Status == PromptStatus.None || ptRes.Status == PromptStatus.Cancel)
        {
          break;
        }

        if (ptRes.Status == PromptStatus.Keyword)
        {
          string kw = ptRes.StringResult.ToUpperInvariant();
          if (kw == "NORTH" || kw == "N")
          {
            orientation = SwitchOrientation.North;
            continue;
          }
          if (kw == "EAST" || kw == "E")
          {
            orientation = SwitchOrientation.East;
            continue;
          }
          if (kw == "SOUTH" || kw == "S")
          {
            orientation = SwitchOrientation.South;
            continue;
          }
          if (kw == "WEST" || kw == "W")
          {
            orientation = SwitchOrientation.West;
            continue;
          }
          if (kw == "TYPE" || kw == "T")
          {
            PromptKeywordOptions typeOpts = new PromptKeywordOptions("\nSelect Switch Type [Standard/Dimmer/Occupancy] <Standard>: ");
            typeOpts.Keywords.Add("Standard");
            typeOpts.Keywords.Add("Dimmer");
            typeOpts.Keywords.Add("Occupancy");
            typeOpts.Keywords.Default = type.ToString();
            var typeRes = ed.GetKeywords(typeOpts);
            if (typeRes.Status == PromptStatus.OK)
            {
              if (Enum.TryParse<SwitchType>(typeRes.StringResult, true, out var newType))
              {
                type = newType;
              }
            }
            continue;
          }
          if (kw == "SUBSCRIPT" || kw == "SUB")
          {
            PromptStringOptions strOpts = new PromptStringOptions("\nEnter custom text/subscript for placed switch (or Enter to reset to default): ")
            {
              AllowSpaces = true
            };
            var strRes = ed.GetString(strOpts);
            if (strRes.Status == PromptStatus.OK)
            {
              overrideText = string.IsNullOrWhiteSpace(strRes.StringResult) ? null : strRes.StringResult.Trim();
              ed.WriteMessage(overrideText == null ? "\nSubscript reset to default." : $"\nSubscript set to '{overrideText}'.");
            }
            continue;
          }
          if (kw == "SETUP")
          {
            SwitchSetupWizard.RunSetup(ed, db, type, orientation);
            projectSettings = SwitchConfigurationStore.LoadSettings(db);
            continue;
          }
        }

        if (ptRes.Status == PromptStatus.OK)
        {
          if (PlaceSwitch(db, ed, projectSettings, type, orientation, ptRes.Value, overrideSubscript: overrideText))
          {
            placedCount++;
          }
        }
      }

      if (placedCount > 0)
      {
        ed.WriteMessage($"\nPlaced {placedCount} switch(es).");
      }
    }

    private static string GetEffectiveBlockName(BlockReference blkRef, Transaction tr)
    {
      if (blkRef.IsDynamicBlock)
      {
        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blkRef.DynamicBlockTableRecord, OpenMode.ForRead);
        return btr.Name;
      }
      return blkRef.Name;
    }

    private static string GetVisibilityState(BlockReference blkRef)
    {
      if (!blkRef.IsDynamicBlock) return string.Empty;
      foreach (DynamicBlockReferenceProperty prop in blkRef.DynamicBlockReferencePropertyCollection)
      {
        if ((prop.PropertyName ?? string.Empty).IndexOf("VISIBILITY", StringComparison.OrdinalIgnoreCase) >= 0)
        {
          return Convert.ToString(prop.Value) ?? string.Empty;
        }
      }
      return string.Empty;
    }

    private static void SetVisibilityState(BlockReference blkRef, string stateName)
    {
      if (!blkRef.IsDynamicBlock || string.IsNullOrWhiteSpace(stateName)) return;
      foreach (DynamicBlockReferenceProperty prop in blkRef.DynamicBlockReferencePropertyCollection)
      {
        if (prop.ReadOnly) continue;
        if ((prop.PropertyName ?? string.Empty).IndexOf("VISIBILITY", StringComparison.OrdinalIgnoreCase) >= 0)
        {
          try
          {
            prop.Value = stateName;
            return;
          }
          catch
          {
          }
        }
      }
    }

    private static void AddAttributes(
      Transaction tr,
      ObjectId blockDefId,
      BlockReference blockRef,
      List<SwitchAttributeDefinition> storedAttributes)
    {
      BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blockDefId, OpenMode.ForRead);
      if (!btr.HasAttributeDefinitions) return;

      var attrMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
      if (storedAttributes != null)
      {
        foreach (var sa in storedAttributes)
        {
          if (!string.IsNullOrWhiteSpace(sa.Tag))
          {
            attrMap[sa.Tag] = sa.TextString;
          }
        }
      }

      foreach (ObjectId id in btr)
      {
        AttributeDefinition attDef = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
        if (attDef != null && !attDef.Constant)
        {
          AttributeReference attRef = new AttributeReference();
          attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);
          if (attrMap.TryGetValue(attDef.Tag, out string customVal))
          {
            attRef.TextString = customVal;
          }
          blockRef.AttributeCollection.AppendAttribute(attRef);
          tr.AddNewlyCreatedDBObject(attRef, true);
        }
      }
    }

    private static string GetTextStyleName(Database db, ObjectId textStyleId, Transaction tr)
    {
      if (textStyleId.IsNull) return "Standard";
      try
      {
        var ts = tr.GetObject(textStyleId, OpenMode.ForRead) as TextStyleTableRecord;
        return ts?.Name ?? "Standard";
      }
      catch
      {
        return "Standard";
      }
    }

    private static ObjectId ResolveTextStyleId(Database db, string styleName, Transaction tr)
    {
      if (string.IsNullOrWhiteSpace(styleName)) return ObjectId.Null;
      try
      {
        TextStyleTable tst = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
        if (tst.Has(styleName))
        {
          return tst[styleName];
        }
      }
      catch
      {
      }
      return ObjectId.Null;
    }

    private static bool LayerExists(Database db, string layerName, Transaction tr)
    {
      if (string.IsNullOrWhiteSpace(layerName)) return false;
      try
      {
        LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
        return lt.Has(layerName);
      }
      catch
      {
        return false;
      }
    }
  }
}
