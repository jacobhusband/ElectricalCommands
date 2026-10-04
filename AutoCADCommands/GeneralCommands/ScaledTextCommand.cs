using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    // Remembered for the AutoCAD session so repeating the command reuses the last direction.
    private static SymbolSide _lastScaledTextSide = SymbolSide.East;

    [CommandMethod("SCALEDTEXT", CommandFlags.Modal)]
    [CommandMethod("TXT", CommandFlags.Modal)]
    public static void InsertScaledText()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      if (!ElectricalDrawingSettingsStore.TryReadScale(db, out var scale))
      {
        ed.WriteMessage(
          "\nScaled text requires a drawing scale. Run SETSCALE (SS) first.");
        return;
      }

      PromptResult textResult = ed.GetString(new PromptStringOptions("\nEnter text: ")
      {
        AllowSpaces = true,
      });
      if (textResult.Status != PromptStatus.OK)
      {
        ed.WriteMessage("\nSCALEDTEXT canceled.");
        return;
      }
      string text = (textResult.StringResult ?? string.Empty).Trim();
      if (text.Length == 0)
      {
        ed.WriteMessage("\nText cannot be blank.");
        return;
      }

      string contents = EscapeMTextPlainText(text);
      double textHeight = ResolveHomerunSymbolSize(scale.PaperInchesPerModelFoot);
      double previewScale = ResolveReceptBlockScale(scale.PaperInchesPerModelFoot);
      double ucsRotation = SymbolGeometry.ResolveUcsRotation(ed);
      ObjectId previewStyleId = ResolveExistingTextStyleId(db, HomerunTextStyleName);
      SymbolSide side = _lastScaledTextSide;

      // Previews are built in plotted inches like the symbols, then scaled up by the jigs.
      Func<SymbolSide, List<Entity>> createPreview = previewSide => new List<Entity>
      {
        CreateScaledText(db, contents, previewSide, textHeight / previewScale, previewStyleId),
      };

      Point3d location;
      using (SymbolLocationJig locationJig = new SymbolLocationJig(
        createPreview(side),
        "\nSpecify text location: ",
        null,
        previewScale,
        ucsRotation))
      {
        if (ed.Drag(locationJig).Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nSCALEDTEXT canceled.");
          return;
        }
        location = locationJig.Location;
      }

      if (!TryPromptSymbolSide(
        ed,
        createPreview,
        side,
        "\nSpecify text direction or [North/East/South/West]: ",
        location,
        previewScale,
        ucsRotation,
        textHeight,
        out side))
      {
        ed.WriteMessage("\nSCALEDTEXT canceled.");
        return;
      }

      try
      {
        ObjectId textStyleId = EnsureHomerunTextStyle(db);
        using (Transaction transaction = db.TransactionManager.StartTransaction())
        {
          BlockTableRecord currentSpace = (BlockTableRecord)transaction.GetObject(
            db.CurrentSpaceId,
            OpenMode.ForWrite);

          MText mtext = CreateScaledText(db, contents, side, textHeight, textStyleId);
          mtext.TransformBy(
            SymbolGeometry.CreateInsertTransform(location, ucsRotation, 1.0));

          currentSpace.AppendEntity(mtext);
          transaction.AddNewlyCreatedDBObject(mtext, true);
          transaction.Commit();
        }

        _lastScaledTextSide = side;
        ed.WriteMessage(
          $"\nPlaced {FormatNumber(textHeight)}\" MText running " +
          $"{side.ToString().ToLowerInvariant()} at {scale.DisplayText}.");
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nUnable to place the text: {ex.Message}");
      }
    }

    // Creates MText anchored at the origin that runs toward the given side while
    // staying readable from the bottom or right of the sheet.
    private static MText CreateScaledText(
      Database database,
      string contents,
      SymbolSide side,
      double textHeight,
      ObjectId textStyleId)
    {
      bool vertical = side == SymbolSide.North || side == SymbolSide.South;
      bool runsBackward = side == SymbolSide.West || side == SymbolSide.South;

      MText text = new MText();
      text.SetDatabaseDefaults(database);
      text.Location = Point3d.Origin;
      text.Contents = contents;
      text.TextHeight = textHeight;
      text.Width = 0.0;
      text.Rotation = vertical ? Math.PI / 2.0 : 0.0;
      text.Attachment = runsBackward
        ? AttachmentPoint.MiddleRight
        : AttachmentPoint.MiddleLeft;
      if (!textStyleId.IsNull)
      {
        text.TextStyleId = textStyleId;
      }
      return text;
    }
  }
}
