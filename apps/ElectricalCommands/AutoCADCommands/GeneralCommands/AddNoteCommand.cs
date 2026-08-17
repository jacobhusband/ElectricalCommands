using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private const double AddNoteDefaultTextHeight = 0.1;
    private const double AddNoteDefaultWidthFactor = 80.0;

    [CommandMethod("ADDNOTE", CommandFlags.Modal)]
    public void AddNote()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null)
      {
        throw new InvalidOperationException("No active AutoCAD document is available.");
      }

      string stage = "notes load";
      try
      {
        AddNoteLoadResult loadResult = AddNoteStore.Load();
        if (loadResult.Tabs == null || loadResult.Tabs.Count == 0)
        {
          ed.WriteMessage($"\nADDNOTE: {loadResult.StatusMessage}");
          return;
        }

        stage = "note selection";
        string selectedNote = PromptForAddNoteText(loadResult.Tabs);
        if (string.IsNullOrWhiteSpace(selectedNote))
        {
          ed.WriteMessage("\nADDNOTE canceled.");
          return;
        }

        stage = "insertion point selection";
        PromptPointOptions pointOptions = new PromptPointOptions(
          "\nClick existing MText to append note, or click empty space to create MText: "
        );
        PromptPointResult pointResult = ed.GetPoint(pointOptions);
        if (pointResult.Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nADDNOTE canceled.");
          return;
        }

        stage = "drawing update";
        AddNoteToDrawing(db, ed, pointResult.Value, selectedNote);
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nADDNOTE error during {stage}: {ex.Message}");
      }
    }

    private static string PromptForAddNoteText(List<AddNoteTab> tabs)
    {
      AddNotePickerWindow window = new AddNotePickerWindow(tabs);
      bool? result = window.ShowDialog();
      return result == true ? window.SelectedNoteText : string.Empty;
    }

    private static void AddNoteToDrawing(
      Database db,
      Editor ed,
      Point3d point,
      string plainNoteText
    )
    {
      string mtextNote = ConvertPlainTextToMTextContents(plainNoteText);
      if (string.IsNullOrWhiteSpace(mtextNote))
      {
        ed.WriteMessage("\nADDNOTE: selected note is empty.");
        return;
      }

      ObjectId mtextId = FindMTextAtPoint(db, ed, point);
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        if (!mtextId.IsNull)
        {
          MText existing = tr.GetObject(mtextId, OpenMode.ForWrite) as MText;
          if (existing != null)
          {
            existing.Contents = AppendMTextContents(existing.Contents, mtextNote);
            tr.Commit();
            ed.WriteMessage("\nADDNOTE: note appended to MText.");
            return;
          }
        }

        BlockTableRecord currentSpace =
          tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
        if (currentSpace == null)
        {
          throw new InvalidOperationException("Current drawing space could not be opened.");
        }

        double textHeight = ResolveCurrentTextHeight(db);
        MText newText = new MText
        {
          Location = point,
          Contents = mtextNote,
          TextHeight = textHeight,
          Width = Math.Max(textHeight * AddNoteDefaultWidthFactor, textHeight),
          Attachment = AttachmentPoint.TopLeft,
          TextStyleId = db.Textstyle,
        };
        newText.SetDatabaseDefaults(db);
        newText.TextStyleId = db.Textstyle;
        newText.LayerId = db.Clayer;

        currentSpace.AppendEntity(newText);
        tr.AddNewlyCreatedDBObject(newText, true);
        tr.Commit();
        ed.WriteMessage("\nADDNOTE: note MText created.");
      }
    }

    private static ObjectId FindMTextAtPoint(Database db, Editor ed, Point3d point)
    {
      double pickboxSize = ResolvePickboxWorldSize(ed);
      Point3d min = new Point3d(point.X - pickboxSize, point.Y - pickboxSize, point.Z);
      Point3d max = new Point3d(point.X + pickboxSize, point.Y + pickboxSize, point.Z);
      SelectionFilter filter = new SelectionFilter(new[]
      {
        new TypedValue((int)DxfCode.Start, "MTEXT"),
      });

      PromptSelectionResult result = ed.SelectCrossingWindow(min, max, filter);
      if (result.Status != PromptStatus.OK || result.Value == null)
      {
        return ObjectId.Null;
      }

      ObjectId[] ids = result.Value.GetObjectIds();
      if (ids.Length == 0)
      {
        return ObjectId.Null;
      }

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        ObjectId bestId = ObjectId.Null;
        double bestDistance = double.MaxValue;
        foreach (ObjectId id in ids)
        {
          MText mtext = tr.GetObject(id, OpenMode.ForRead) as MText;
          if (mtext == null)
          {
            continue;
          }

          double distance = ResolveMTextPickDistance(mtext, point, pickboxSize);
          if (distance < bestDistance)
          {
            bestDistance = distance;
            bestId = id;
          }
        }

        tr.Commit();
        return bestId;
      }
    }

    private static double ResolveMTextPickDistance(MText mtext, Point3d point, double tolerance)
    {
      try
      {
        Extents3d extents = mtext.GeometricExtents;
        if (
          point.X >= extents.MinPoint.X - tolerance &&
          point.X <= extents.MaxPoint.X + tolerance &&
          point.Y >= extents.MinPoint.Y - tolerance &&
          point.Y <= extents.MaxPoint.Y + tolerance
        )
        {
          return 0.0;
        }
      }
      catch
      {
      }

      return mtext.Location.DistanceTo(point);
    }

    private static double ResolvePickboxWorldSize(Editor ed)
    {
      const double fallbackFraction = 1.0 / 400.0;
      try
      {
        ViewTableRecord view = ed.GetCurrentView();
        double screenWidth = ResolveScreenWidth();
        double pickboxPixels = ResolveSystemVariableDouble("PICKBOX", 8.0);
        if (view != null && view.Width > 0.0 && screenWidth > 0.0)
        {
          return Math.Max(view.Width * pickboxPixels / screenWidth, view.Width * fallbackFraction);
        }
      }
      catch
      {
      }

      try
      {
        ViewTableRecord view = ed.GetCurrentView();
        if (view != null && view.Width > 0.0)
        {
          return view.Width * fallbackFraction;
        }
      }
      catch
      {
      }

      return 1.0;
    }

    private static double ResolveScreenWidth()
    {
      object value = Application.GetSystemVariable("SCREENSIZE");
      if (value is Point2d point2d)
      {
        return point2d.X;
      }
      if (value is Point3d point3d)
      {
        return point3d.X;
      }
      return 1000.0;
    }

    private static double ResolveSystemVariableDouble(string name, double fallback)
    {
      try
      {
        object value = Application.GetSystemVariable(name);
        if (value is short shortValue)
        {
          return shortValue;
        }
        if (value is int intValue)
        {
          return intValue;
        }
        if (value is double doubleValue)
        {
          return doubleValue;
        }
        if (double.TryParse(Convert.ToString(value), out double parsed))
        {
          return parsed;
        }
      }
      catch
      {
      }

      return fallback;
    }

    private static double ResolveCurrentTextHeight(Database db)
    {
      double textHeight = db.Textsize;
      return textHeight > 0.0 ? textHeight : AddNoteDefaultTextHeight;
    }

    private static string AppendMTextContents(string existingContents, string mtextNote)
    {
      string existing = (existingContents ?? string.Empty).Trim();
      return string.IsNullOrWhiteSpace(existing)
        ? mtextNote
        : $"{existing}\\P\\P{mtextNote}";
    }

    private static string ConvertPlainTextToMTextContents(string plainText)
    {
      string normalized = (plainText ?? string.Empty)
        .Replace("\r\n", "\n")
        .Replace("\r", "\n")
        .Trim();
      if (string.IsNullOrWhiteSpace(normalized))
      {
        return string.Empty;
      }

      string[] lines = normalized.Split('\n');
      return string.Join("\\P", lines.Select(EscapeMTextPlainText));
    }

    private static string EscapeMTextPlainText(string value)
    {
      var builder = new StringBuilder();
      foreach (char ch in value ?? string.Empty)
      {
        if (ch == '\\' || ch == '{' || ch == '}')
        {
          builder.Append('\\');
        }
        builder.Append(ch);
      }
      return builder.ToString();
    }
  }
}
