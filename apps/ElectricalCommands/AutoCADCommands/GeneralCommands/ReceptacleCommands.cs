using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private const string ReceptBlockName = "RECEPT";
    private const string ReceptNoteLayerName = "EL-NOTES";
    private const short ReceptNoteLayerColor = 4;
    private const double ReceptModelInchesPerFoot = 12.0;
    private const double ReceptNoteGapInPaperInches = 0.0625;

    [CommandMethod("RE", CommandFlags.Modal)]
    public static void InsertReceptEast()
    {
      InsertRecept(
        "east",
        270.0,
        0.0,
        AttachmentPoint.MiddleLeft);
    }

    [CommandMethod("RN", CommandFlags.Modal)]
    public static void InsertReceptNorth()
    {
      InsertRecept(
        "north",
        0.0,
        270.0,
        AttachmentPoint.MiddleRight);
    }

    [CommandMethod("RW", CommandFlags.Modal)]
    public static void InsertReceptWest()
    {
      InsertRecept(
        "west",
        90.0,
        0.0,
        AttachmentPoint.MiddleRight);
    }

    [CommandMethod("RS", CommandFlags.Modal)]
    public static void InsertReceptSouth()
    {
      InsertRecept(
        "south",
        180.0,
        270.0,
        AttachmentPoint.MiddleLeft);
    }

    private static void InsertRecept(
      string direction,
      double rotationDegrees,
      double noteRotationDegrees,
      AttachmentPoint noteAttachment)
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      if (doc == null)
      {
        return;
      }

      Database db = doc.Database;
      Editor ed = doc.Editor;

      if (!ElectricalDrawingSettingsStore.TryReadScale(db, out var scale))
      {
        ed.WriteMessage(
          "\nReceptacle insertion requires a drawing scale. " +
          "Run SETSCALE (SS) first.");
        return;
      }

      if (!ElectricalDrawingSettingsStore.TryReadPanelName(
        db,
        out string panelName))
      {
        ed.WriteMessage(
          "\nReceptacle insertion requires a panel name. " +
          "Run SETPANELNAME (SPN) first.");
        return;
      }

      PromptStringOptions circuitOptions = new PromptStringOptions(
        $"\nEnter circuit number for panel {panelName} " +
        "(for example 19 or 1/5): ")
      {
        AllowSpaces = false
      };
      PromptResult circuitResult = ed.GetString(circuitOptions);
      if (circuitResult.Status != PromptStatus.OK)
      {
        ed.WriteMessage("\nReceptacle insertion canceled.");
        return;
      }

      string circuitNumber =
        (circuitResult.StringResult ?? string.Empty).Trim().TrimStart('-');
      if (circuitNumber.Length == 0)
      {
        ed.WriteMessage("\nCircuit number cannot be blank.");
        return;
      }

      string panelCircuitLabel =
        BuildPanelLabel(panelName) + circuitNumber;
      double blockScale = ResolveReceptBlockScale(
        scale.PaperInchesPerModelFoot);
      double textHeight = ResolveHomerunSymbolSize(
        scale.PaperInchesPerModelFoot);
      double noteGap = ReceptNoteGapInPaperInches * blockScale;

      PromptPointOptions pointOptions = new PromptPointOptions(
        $"\nSpecify insertion point for {direction} receptacle: ")
      {
        AllowNone = false
      };

      PromptPointResult pointResult = ed.GetPoint(pointOptions);
      if (pointResult.Status != PromptStatus.OK)
      {
        return;
      }

      Point3d insertionPoint =
        pointResult.Value.TransformBy(ed.CurrentUserCoordinateSystem);
      double rotationRadians = rotationDegrees * Math.PI / 180.0;

      try
      {
        ObjectId noteLayerId = EnsureReceptNoteLayer(db);
        ObjectId textStyleId = EnsureHomerunTextStyle(db);

        using (Transaction transaction = db.TransactionManager.StartTransaction())
        {
          BlockTable blockTable = (BlockTable)transaction.GetObject(
            db.BlockTableId,
            OpenMode.ForRead);

          if (!blockTable.Has(ReceptBlockName))
          {
            ed.WriteMessage(
              $"\nCannot insert receptacle: block \"{ReceptBlockName}\" " +
              "is not defined in the current drawing.");
            return;
          }

          BlockTableRecord currentSpace = (BlockTableRecord)transaction.GetObject(
            db.CurrentSpaceId,
            OpenMode.ForWrite);
          ObjectId blockDefinitionId = blockTable[ReceptBlockName];

          BlockReference blockReference = new BlockReference(
            insertionPoint,
            blockDefinitionId);
          blockReference.SetDatabaseDefaults(db);
          blockReference.ScaleFactors = new Scale3d(blockScale);
          blockReference.Rotation = rotationRadians;

          currentSpace.AppendEntity(blockReference);
          transaction.AddNewlyCreatedDBObject(blockReference, true);

          AddDefaultAttributes(
            transaction,
            blockDefinitionId,
            blockReference);

          Vector3d noteDirection = Vector3d.YAxis
            .RotateBy(rotationRadians, Vector3d.ZAxis);
          Point3d noteLocation = ResolveReceptNoteLocation(
            blockReference,
            insertionPoint,
            noteDirection,
            noteGap);

          MText panelLabel = CreateReceptPanelLabel(
            db,
            noteLocation,
            noteRotationDegrees * Math.PI / 180.0,
            textHeight,
            EscapeMTextPlainText(panelCircuitLabel),
            noteLayerId,
            textStyleId,
            noteAttachment);
          currentSpace.AppendEntity(panelLabel);
          transaction.AddNewlyCreatedDBObject(panelLabel, true);

          transaction.Commit();
        }

        ed.WriteMessage(
          $"\nInserted {ReceptBlockName} facing {direction} at " +
          $"{scale.DisplayText} (X/Y/Z scale {FormatNumber(blockScale)}) " +
          $"with circuit label {panelCircuitLabel}.");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage(
          $"\nUnable to insert {ReceptBlockName}: {ex.Message}");
      }
    }

    private static double ResolveReceptBlockScale(
      double paperInchesPerModelFoot)
    {
      return ReceptModelInchesPerFoot / paperInchesPerModelFoot;
    }

    private static ObjectId EnsureReceptNoteLayer(Database database)
    {
      using (Transaction transaction =
        database.TransactionManager.StartTransaction())
      {
        LayerTable layers = (LayerTable)transaction.GetObject(
          database.LayerTableId,
          OpenMode.ForRead);

        if (layers.Has(ReceptNoteLayerName))
        {
          ObjectId existingId = layers[ReceptNoteLayerName];
          transaction.Commit();
          return existingId;
        }

        layers.UpgradeOpen();
        LayerTableRecord layer = new LayerTableRecord
        {
          Name = ReceptNoteLayerName,
          Color = Color.FromColorIndex(
            ColorMethod.ByAci,
            ReceptNoteLayerColor)
        };
        ObjectId layerId = layers.Add(layer);
        transaction.AddNewlyCreatedDBObject(layer, true);
        transaction.Commit();
        return layerId;
      }
    }

    private static Point3d ResolveReceptNoteLocation(
      BlockReference blockReference,
      Point3d insertionPoint,
      Vector3d noteDirection,
      double gap)
    {
      try
      {
        Extents3d extents = blockReference.GeometricExtents;
        Point3d center = new Point3d(
          (extents.MinPoint.X + extents.MaxPoint.X) / 2.0,
          (extents.MinPoint.Y + extents.MaxPoint.Y) / 2.0,
          insertionPoint.Z);
        double halfWidth =
          (extents.MaxPoint.X - extents.MinPoint.X) / 2.0;
        double halfHeight =
          (extents.MaxPoint.Y - extents.MinPoint.Y) / 2.0;
        double distanceToEdge =
          Math.Abs(noteDirection.X) * halfWidth +
          Math.Abs(noteDirection.Y) * halfHeight;

        return center + noteDirection * (distanceToEdge + gap);
      }
      catch
      {
        return insertionPoint + noteDirection * gap;
      }
    }

    private static MText CreateReceptPanelLabel(
      Database database,
      Point3d location,
      double rotation,
      double textHeight,
      string contents,
      ObjectId layerId,
      ObjectId textStyleId,
      AttachmentPoint attachment)
    {
      MText text = new MText();
      text.SetDatabaseDefaults(database);
      text.LayerId = layerId;
      text.ColorIndex = 256;
      text.Location = location;
      text.Contents = contents;
      text.Attachment = attachment;
      text.TextHeight = textHeight;
      text.Width = 0.0;
      text.Rotation = rotation;
      text.Annotative = AnnotativeStates.False;
      text.TextStyleId = textStyleId;
      return text;
    }

    private static void AddDefaultAttributes(
      Transaction transaction,
      ObjectId blockDefinitionId,
      BlockReference blockReference)
    {
      BlockTableRecord blockDefinition =
        (BlockTableRecord)transaction.GetObject(
          blockDefinitionId,
          OpenMode.ForRead);

      foreach (ObjectId entityId in blockDefinition)
      {
        AttributeDefinition attributeDefinition =
          transaction.GetObject(
            entityId,
            OpenMode.ForRead) as AttributeDefinition;

        if (attributeDefinition == null || attributeDefinition.Constant)
        {
          continue;
        }

        AttributeReference attributeReference = new AttributeReference();
        attributeReference.SetAttributeFromBlock(
          attributeDefinition,
          blockReference.BlockTransform);
        attributeReference.TextString = attributeDefinition.TextString;

        blockReference.AttributeCollection.AppendAttribute(attributeReference);
        transaction.AddNewlyCreatedDBObject(attributeReference, true);
      }
    }
  }
}
