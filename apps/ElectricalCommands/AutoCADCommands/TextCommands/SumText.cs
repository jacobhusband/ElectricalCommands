using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ElectricalCommands
{
  public class SumTextCommand
  {
    private static readonly Regex SquareFeetRegex = new Regex(
        @"(?<value>[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)\s*(?:sq(?:uare)?\.?\s*(?:ft|feet)\b\.?|sf\b)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase
    );

    private static readonly Regex TotalLabelRegex = new Regex(
        @"^\s*total\s*:",
        RegexOptions.Compiled | RegexOptions.IgnoreCase
    );

    [CommandMethod("SUMTEXT", CommandFlags.UsePickSet)]
    public void SumText()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null)
      {
        throw new InvalidOperationException("No active AutoCAD document is available.");
      }

      try
      {
        PromptSelectionResult selection = GetTextSelection(ed);
        if (selection.Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nNo text objects selected.");
          return;
        }

        List<SelectedText> selectedTexts = new List<SelectedText>();
        List<RoomInfo> roomInfoList = new List<RoomInfo>();
        double totalSquareFeet = 0.0;

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          foreach (SelectedObject selectedObject in selection.Value)
          {
            Entity entity = tr.GetObject(selectedObject.ObjectId, OpenMode.ForRead) as Entity;
            if (!(entity is DBText) && !(entity is MText))
              continue;

            string text = GetText(entity).Trim();
            selectedTexts.Add(
                new SelectedText
                {
                  Entity = entity,
                  Text = text,
                  Position = GetTextPosition(entity)
                }
            );
          }

          List<SelectedText> areaTexts = selectedTexts
              .Where(t => !IsTotalLabel(t.Text))
              .Select(
                  t =>
                  {
                    t.SquareFeet = ExtractSquareFeet(t.Text);
                    return t;
                  }
              )
              .Where(t => t.SquareFeet.HasValue)
              .ToList();

          if (areaTexts.Count == 0)
          {
            ed.WriteMessage(
                "\nNo square-footage text was found. Select room names and area labels such as \"125 sq ft\"."
            );
            return;
          }

          totalSquareFeet = areaTexts.Sum(t => t.SquareFeet.Value);

          List<SelectedText> roomNameTexts = selectedTexts
              .Where(t => !t.SquareFeet.HasValue && !IsTotalLabel(t.Text))
              .ToList();

          foreach (SelectedText areaText in areaTexts)
          {
            SelectedText nearestRoom = FindNearestText(areaText, roomNameTexts);
            if (nearestRoom != null)
            {
              roomInfoList.Add(
                  new RoomInfo
                  {
                    RoomType = nearestRoom.Text,
                    SquareFeet = areaText.SquareFeet.Value
                  }
              );
            }
          }

          Point3d labelPosition = GetLabelPosition(selectedTexts, db.Textsize);
          double textHeight = GetLabelTextHeight(selectedTexts, db.Textsize);
          AddTotalLabel(tr, db, labelPosition, textHeight, totalSquareFeet);

          tr.Commit();
        }

        List<RoomInfo> combinedRooms = roomInfoList
            .GroupBy(r => r.RoomType, StringComparer.OrdinalIgnoreCase)
            .Select(
                group =>
                  new RoomInfo
                  {
                    RoomType = group.First().RoomType,
                    SquareFeet = group.Sum(room => room.SquareFeet)
                  }
            )
            .ToList();

        combinedRooms.Add(
            new RoomInfo
            {
              RoomType = "TOTAL",
              SquareFeet = totalSquareFeet
            }
        );

        string drawingDirectory = Path.GetDirectoryName(db.Filename);
        if (!string.IsNullOrWhiteSpace(drawingDirectory))
        {
          string jsonPath = Path.Combine(drawingDirectory, "T24Output.json");
          string json = JsonConvert.SerializeObject(combinedRooms, Formatting.Indented);
          File.WriteAllText(jsonPath, json);
          ed.WriteMessage($"\nExported room information to: {jsonPath}");
        }
        else
        {
          ed.WriteMessage("\nThe drawing has not been saved, so T24Output.json was not exported.");
        }

        ed.WriteMessage(
            $"\nTotal square footage: {FormatSquareFeet(totalSquareFeet)} sqft. Added a total label above the selection."
        );
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage($"\nSUMTEXT error: {ex.Message}");
      }
      finally
      {
        ed.SetImpliedSelection(new ObjectId[0]);
      }
    }

    private static PromptSelectionResult GetTextSelection(Editor ed)
    {
      PromptSelectionResult selection = ed.SelectImplied();
      if (selection.Status == PromptStatus.OK)
        return selection;

      PromptSelectionOptions options = new PromptSelectionOptions
      {
        MessageForAdding = "Select room names and square-footage text: ",
        AllowDuplicates = false,
        RejectObjectsOnLockedLayers = true
      };
      SelectionFilter filter = new SelectionFilter(
          new[] { new TypedValue((int)DxfCode.Start, "TEXT,MTEXT") }
      );
      return ed.GetSelection(options, filter);
    }

    private static string GetText(Entity entity)
    {
      DBText dbText = entity as DBText;
      if (dbText != null)
        return dbText.TextString ?? string.Empty;

      MText mText = entity as MText;
      return mText?.Text ?? string.Empty;
    }

    private static Point3d GetTextPosition(Entity entity)
    {
      DBText dbText = entity as DBText;
      if (dbText != null)
        return dbText.Position;

      MText mText = entity as MText;
      return mText?.Location ?? Point3d.Origin;
    }

    private static double? ExtractSquareFeet(string text)
    {
      double total = 0.0;
      bool foundValue = false;

      foreach (Match match in SquareFeetRegex.Matches(text ?? string.Empty))
      {
        string valueText = match.Groups["value"].Value.Replace(",", string.Empty);
        double value;
        if (
          double.TryParse(
              valueText,
              NumberStyles.Float,
              CultureInfo.InvariantCulture,
              out value
          )
        )
        {
          total += value;
          foundValue = true;
        }
      }

      return foundValue ? (double?)total : null;
    }

    private static bool IsTotalLabel(string text)
    {
      return TotalLabelRegex.IsMatch(text ?? string.Empty);
    }

    private static SelectedText FindNearestText(
        SelectedText areaText,
        IEnumerable<SelectedText> roomNameTexts
    )
    {
      return roomNameTexts
          .OrderBy(candidate => areaText.Position.DistanceTo(candidate.Position))
          .FirstOrDefault();
    }

    private static Point3d GetLabelPosition(
        IEnumerable<SelectedText> selectedTexts,
        double defaultTextHeight
    )
    {
      bool hasExtents = false;
      Point3d minPoint = Point3d.Origin;
      Point3d maxPoint = Point3d.Origin;

      foreach (SelectedText selectedText in selectedTexts)
      {
        try
        {
          Extents3d extents = selectedText.Entity.GeometricExtents;
          if (!hasExtents)
          {
            minPoint = extents.MinPoint;
            maxPoint = extents.MaxPoint;
            hasExtents = true;
          }
          else
          {
            minPoint = new Point3d(
                Math.Min(minPoint.X, extents.MinPoint.X),
                Math.Min(minPoint.Y, extents.MinPoint.Y),
                Math.Min(minPoint.Z, extents.MinPoint.Z)
            );
            maxPoint = new Point3d(
                Math.Max(maxPoint.X, extents.MaxPoint.X),
                Math.Max(maxPoint.Y, extents.MaxPoint.Y),
                Math.Max(maxPoint.Z, extents.MaxPoint.Z)
            );
          }
        }
        catch (System.Exception)
        {
          // Some text objects can lack geometric extents; their insertion points are used below.
        }
      }

      double textHeight = GetLabelTextHeight(selectedTexts, defaultTextHeight);
      if (hasExtents)
      {
        return new Point3d(
            (minPoint.X + maxPoint.X) / 2.0,
            maxPoint.Y + (textHeight * 2.0),
            maxPoint.Z
        );
      }

      List<Point3d> positions = selectedTexts.Select(t => t.Position).ToList();
      return new Point3d(
          positions.Average(p => p.X),
          positions.Max(p => p.Y) + (textHeight * 2.0),
          positions.Max(p => p.Z)
      );
    }

    private static double GetLabelTextHeight(
        IEnumerable<SelectedText> selectedTexts,
        double defaultTextHeight
    )
    {
      double height = selectedTexts
          .Select(
              selectedText =>
              {
                DBText dbText = selectedText.Entity as DBText;
                if (dbText != null)
                  return dbText.Height;

                MText mText = selectedText.Entity as MText;
                return mText?.TextHeight ?? 0.0;
              }
          )
          .Where(value => value > 0.0)
          .DefaultIfEmpty(defaultTextHeight > 0.0 ? defaultTextHeight : 1.0)
          .Max();

      return height > 0.0 ? height : 1.0;
    }

    private static void AddTotalLabel(
        Transaction tr,
        Database db,
        Point3d position,
        double textHeight,
        double totalSquareFeet
    )
    {
      DBText totalText = new DBText
      {
        Height = textHeight,
        TextString = $"total: {FormatSquareFeet(totalSquareFeet)}sqft",
        Rotation = 0.0,
        HorizontalMode = TextHorizontalMode.TextCenter,
        VerticalMode = TextVerticalMode.TextVerticalMid,
        Position = position,
        AlignmentPoint = position,
        LayerId = db.Clayer,
        TextStyleId = db.Textstyle
      };

      BlockTableRecord currentSpace = (BlockTableRecord)
          tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
      currentSpace.AppendEntity(totalText);
      tr.AddNewlyCreatedDBObject(totalText, true);
    }

    private static string FormatSquareFeet(double squareFeet)
    {
      return squareFeet.ToString("0.##", CultureInfo.InvariantCulture);
    }

    private class SelectedText
    {
      public Entity Entity { get; set; }
      public string Text { get; set; }
      public Point3d Position { get; set; }
      public double? SquareFeet { get; set; }
    }
  }

  public class RoomInfo
  {
    public string RoomType { get; set; }
    public double SquareFeet { get; set; }
  }
}
