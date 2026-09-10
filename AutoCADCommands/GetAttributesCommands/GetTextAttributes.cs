using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;

namespace ElectricalCommands
{
  using static ElectricalCommands.Globals;
  public partial class GeneralCommands
  {
    [CommandMethod("TEXTATTR")]
    public void GETTEXTATTRIBUTES()
    {
      var (doc, db, ed) = GetGlobals();

      var textId = SelectTextObject();
      if (textId.IsNull)
      {
        ed.WriteMessage("\nNo text object selected.");
        return;
      }

      var textObject = GetTextObject(textId);
      if (textObject == null)
      {
        ed.WriteMessage("\nFailed to get text object.");
        return;
      }

      var coordinate = GetCoordinate();

      var startPoint = new Point3d(
          textObject.Position.X - coordinate.X,
          textObject.Position.Y - coordinate.Y,
          0
      );

      string startXStr =
          startPoint.X == 0
              ? ""
              : (startPoint.X > 0 ? $" + {startPoint.X}" : $" - {-startPoint.X}");
      string startYStr =
          startPoint.Y == 0
              ? ""
              : (startPoint.Y > 0 ? $" + {startPoint.Y}" : $" - {-startPoint.Y}");

      var formattedText =
          $"CreateAndPositionText(tr, \"{textObject.TextString}\", \"{textObject.TextStyleName}\", {textObject.Height}, {textObject.WidthFactor}, {textObject.Color.ColorIndex}, \"{textObject.Layer}\", new Point3d(endPoint.X{startXStr}, endPoint.Y{startYStr}, 0));";

      var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
      var filePath = Path.Combine(desktopPath, "TextAttributes.txt");

      SaveTextToFile(formattedText, filePath);
      ed.WriteMessage($"\nText attributes saved to file: {filePath}");
    }

    private void SaveTextToFile(string text, string filePath)
    {
      using (StreamWriter writer = new StreamWriter(filePath, true))
      {
        writer.WriteLine(text);
      }
    }

    private ObjectId SelectTextObject()
    {
      var (doc, db, ed) = GetGlobals();

      var promptOptions = new PromptEntityOptions("\nSelect text object: ");
      promptOptions.SetRejectMessage("\nInvalid: must be DBText");
      promptOptions.AddAllowedClass(typeof(DBText), true);
      var promptResult = ed.GetEntity(promptOptions);

      if (promptResult.Status == PromptStatus.OK)
        return promptResult.ObjectId;

      return ObjectId.Null;
    }

    private DBText GetTextObject(ObjectId objectId)
    {
      using (var tr = objectId.Database.TransactionManager.StartTransaction())
      {
        var textObject = tr.GetObject(objectId, OpenMode.ForRead) as DBText;
        if (textObject != null)
          return textObject;

        return null;
      }
    }

    private Point3d GetCoordinate()
    {
      var (doc, _, ed) = GetGlobals();

      var promptOptions = new PromptPointOptions("\nSelect a coordinate: ");
      var promptResult = ed.GetPoint(promptOptions);

      if (promptResult.Status == PromptStatus.OK)
        return promptResult.Value;

      return new Point3d(0, 0, 0);
    }
  }
}
