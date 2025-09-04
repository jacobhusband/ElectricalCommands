using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;

namespace ElectricalCommands {
  public partial class GeneralCommands {
    [CommandMethod("GETLINEATTRIBUTES")]
    public void GETLINEATTRIBUTES() {
      var (doc, db, ed) = GeneralCommands.GetGlobals();

      PromptEntityOptions linePromptOptions = new PromptEntityOptions("\nSelect a line: ");
      linePromptOptions.SetRejectMessage("\nSelected object is not a line.");
      linePromptOptions.AddAllowedClass(typeof(Line), true);

      PromptEntityResult lineResult = ed.GetEntity(linePromptOptions);
      if (lineResult.Status != PromptStatus.OK) {
        ed.WriteMessage("\nNo line selected.");
        return;
      }

      using (Transaction tr = db.TransactionManager.StartTransaction()) {
        Line line = tr.GetObject(lineResult.ObjectId, OpenMode.ForRead) as Line;
        if (line == null) {
          ed.WriteMessage("\nSelected object is not a line.");
          return;
        }

        PromptPointOptions startPointOptions = new PromptPointOptions(
            "\nSelect the reference point: "
        );
        PromptPointResult startPointResult = ed.GetPoint(startPointOptions);
        if (startPointResult.Status != PromptStatus.OK) {
          ed.WriteMessage("\nNo reference point selected.");
          return;
        }

        Point3d startPoint = startPointResult.Value;
        Vector3d vector = line.EndPoint - line.StartPoint;

        var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        var filePath = Path.Combine(desktopPath, "LineAttributes.txt");

        SaveLineAttributesToFile(line, startPoint, vector, filePath);

        ed.WriteMessage($"\nLine attributes saved to file: {filePath}");
      }
    }

    private void SaveLineAttributesToFile(Line line, Point3d startPoint, Vector3d vector, string filePath) {
      using (StreamWriter writer = new StreamWriter(filePath, true)) {
        double startX = line.StartPoint.X - startPoint.X;
        double startY = line.StartPoint.Y - startPoint.Y;
        double endX = line.EndPoint.X - startPoint.X;
        double endY = line.EndPoint.Y - startPoint.Y;

        string startXStr =
            startX == 0 ? "" : (startX > 0 ? $" + {startX}" : $" - {-startX}");
        string startYStr =
            startY == 0 ? "" : (startY > 0 ? $" + {startY}" : $" - {-startY}");
        string endXStr = endX == 0 ? "" : (endX > 0 ? $" + {endX}" : $" - {-endX}");
        string endYStr = endY == 0 ? "" : (endY > 0 ? $" + {endY}" : $" - {-endY}");

        writer.WriteLine(
            $"CreateLine(tr, btr, endPoint.X{startXStr}, endPoint.Y{startYStr}, endPoint.X{endXStr}, endPoint.Y{endYStr}, \"{line.Layer}\");"
        );
      }
    }
  }
}
