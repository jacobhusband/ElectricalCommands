using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    // Like ETAG, but the tag text is read from an existing tag: box it with two
    // clicks, then pick the point to tag and the tag location.
    [CommandMethod("ETAGOCR", CommandFlags.Modal)]
    public void EquipmentTagOcrCommand()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      if (!TryResolveKnPlacementScale(ed, db, "ETAGOCR", out double scaleDenom)) return;

      PromptPointResult firstResult = ed.GetPoint(
        new PromptPointOptions("\nSpecify first corner of the existing tag: "));
      if (firstResult.Status != PromptStatus.OK)
      {
        ed.WriteMessage("\nETAGOCR canceled.");
        return;
      }
      // The capture uses where the mouse was clicked, not the picked point, so
      // object snaps do not shift the box.
      System.Drawing.Point firstScreen = System.Windows.Forms.Cursor.Position;

      PromptPointResult secondResult = ed.GetCorner(
        new PromptCornerOptions("\nSpecify opposite corner: ", firstResult.Value));
      if (secondResult.Status != PromptStatus.OK)
      {
        ed.WriteMessage("\nETAGOCR canceled.");
        return;
      }
      System.Drawing.Point secondScreen = System.Windows.Forms.Cursor.Position;

      System.Drawing.Bitmap capture = null;
      string tagValue;
      try
      {
        ed.UpdateScreen();
        capture = EquipmentTagOcr.CaptureScreen(firstScreen, secondScreen);

        // Real text in the box is exact, so prefer it over OCR.
        tagValue = ReadEquipmentTagFromDrawing(ed, db, firstResult.Value, secondResult.Value);
        if (!string.IsNullOrEmpty(tagValue))
        {
          ed.WriteMessage($"\nRead tag from drawing text: {tagValue}");
        }
        else if (capture == null)
        {
          ed.WriteMessage("\nThe picked box is too small to read.");
        }
        else
        {
          tagValue = EquipmentTagOcr.ReadImage(capture, out string ocrError);
          ed.WriteMessage(string.IsNullOrEmpty(tagValue)
            ? $"\nCould not read the tag: {ocrError}."
            : $"\nOCR read the tag as: {tagValue}");
        }
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nCould not read the tag: {ex.Message}");
        tagValue = null;
      }
      finally
      {
        capture?.Dispose();
      }

      if (string.IsNullOrEmpty(tagValue))
      {
        tagValue = PromptEquipmentTagValue(ed, null);
        if (string.IsNullOrEmpty(tagValue))
        {
          ed.WriteMessage("\nETAGOCR canceled.");
          return;
        }
      }

      PlaceEquipmentTag(db, ed, "ETAGOCR", scaleDenom, tagValue, true);
    }

    // Text, MText, and visible attributes lying wholly inside the box, read top
    // to bottom and joined the way OCR lines are.
    private static string ReadEquipmentTagFromDrawing(
      Editor ed,
      Database db,
      Point3d corner1,
      Point3d corner2)
    {
      SelectionFilter filter = new SelectionFilter(new[]
      {
        new TypedValue((int)DxfCode.Start, "TEXT,MTEXT,INSERT")
      });
      PromptSelectionResult selection = ed.SelectWindow(corner1, corner2, filter);
      if (selection.Status != PromptStatus.OK) return null;

      List<(Point3d Position, string Text)> items = new List<(Point3d, string)>();
      using (Transaction tr = db.TransactionManager.StartOpenCloseTransaction())
      {
        foreach (ObjectId id in selection.Value.GetObjectIds())
        {
          switch (tr.GetObject(id, OpenMode.ForRead))
          {
            case DBText text:
              items.Add((text.Position, text.TextString));
              break;
            case MText mtext:
              items.Add((mtext.Location, mtext.Text));
              break;
            case BlockReference br:
              foreach (ObjectId attributeId in br.AttributeCollection)
              {
                if (tr.GetObject(attributeId, OpenMode.ForRead) is AttributeReference attribute
                  && !attribute.Invisible)
                {
                  items.Add((attribute.Position, attribute.TextString));
                }
              }
              break;
          }
        }
        tr.Commit();
      }

      IEnumerable<string> lines = items
        .OrderByDescending(item => item.Position.Y)
        .ThenBy(item => item.Position.X)
        .SelectMany(item => (item.Text ?? string.Empty).Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));
      return EquipmentTagOcr.CleanTagText(lines, false);
    }
  }
}
