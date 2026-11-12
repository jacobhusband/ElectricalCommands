using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using System;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    // Text and spacing shared by keyed-note and general-note tables. Text
    // matches the LFS and EQS schedules: ARIALNARROW-1-8 in color 4, a 1/4"
    // title over 3/32" text. Sizes are plotted inches; tables multiply them
    // by their scale (text height / 3/32").
    private const double NoteTableTitleHeight = 0.25;
    private const double NoteTableTextHeight = 0.09375;
    private const short NoteTableTextColorIndex = 4;
    private const double NoteTableNoteGap = 0.125;
    private const double NoteTableNumberPad = 0.03;
    private const double NoteTableTextIndent = 0.06;

    internal static string GetNoteTableTitle(NoteTableKind kind)
    {
      return kind == NoteTableKind.Keyed ? "KEYED NOTES" : "GENERAL NOTES";
    }

    /// <summary>
    /// Merges row 0 across the table and writes the title into it. The table
    /// must already have the extra row.
    /// </summary>
    private static void SetNoteTableTitle(Table table, NoteTableKind kind)
    {
      table.MergeCells(CellRange.Create(table, 0, 0, 0, table.Columns.Count - 1));
      EnsureKeyNoteTableCellContent(table, 0, 0);
      table.SetTextString(0, 0, 0, GetNoteTableTitle(kind));
    }

    /// <summary>
    /// A note table's title row is row 0 merged across every column; renumber
    /// and import skip it like any other merged header.
    /// </summary>
    private static bool IsNoteTableTitleRow(Table table, int row)
    {
      if (row != 0 || table.Columns.Count < 2)
      {
        return false;
      }
      try
      {
        CellRange range = table.Cells[0, 0].GetMergeRange();
        return range != null &&
          range.LeftColumn == 0 &&
          range.RightColumn == table.Columns.Count - 1 &&
          range.TopRow == 0 &&
          range.BottomRow == 0;
      }
      catch (Autodesk.AutoCAD.Runtime.Exception)
      {
        return false;
      }
    }

    /// <summary>
    /// Applies the shared text style, color, size, alignment and cell spacing
    /// to every cell. Columns come in pairs: number (hexagon or "1.") then
    /// note text. Each note's first text line is centered on its hexagon, and
    /// each note gets the same gap below it.
    /// </summary>
    private static void ApplyNoteTableStyle(
      Table table,
      NoteTableKind kind,
      double scale,
      ObjectId textStyleId)
    {
      scale = scale > 1e-9 ? scale : 1.0;
      Color textColor = Color.FromColorIndex(ColorMethod.ByAci, NoteTableTextColorIndex);
      double textHeight = NoteTableTextHeight * scale;
      double gap = NoteTableNoteGap * scale;
      double numberPad = NoteTableNumberPad * scale;
      double indent = NoteTableTextIndent * scale;
      double textTop = kind == NoteTableKind.Keyed
        ? Math.Max(0.0, (KeyNoteHexHeight - NoteTableTextHeight) / 2.0) * scale
        : 0.0;
      bool hasTitle = IsNoteTableTitleRow(table, 0);

      table.HorizontalCellMargin = numberPad;
      table.VerticalCellMargin = gap;
      for (int row = 0; row < table.Rows.Count; row++)
      {
        for (int column = 0; column < table.Columns.Count; column++)
        {
          HideKeyNoteCellGrid(table, row, column);
          if (hasTitle && row == 0)
          {
            SetKeyNoteCellMargins(table, row, column, 0.0, gap, 0.0);
            if (column == 0)
            {
              SetNoteTableText(table, row, column, NoteTableTitleHeight * scale, textStyleId, textColor);
              table.Cells[row, column].Alignment = CellAlignment.TopLeft;
            }
            continue;
          }

          bool numberColumn = column % 2 == 0;
          if (numberColumn)
          {
            SetKeyNoteCellMargins(table, row, column, 0.0, gap, numberPad);
            table.Cells[row, column].Alignment = kind == NoteTableKind.Keyed
              ? CellAlignment.TopCenter
              : CellAlignment.TopRight;
            if (kind == NoteTableKind.General)
            {
              SetNoteTableText(table, row, column, textHeight, textStyleId, textColor);
            }
          }
          else
          {
            SetKeyNoteCellMargins(table, row, column, textTop, gap, indent);
            table.Cells[row, column].Alignment = CellAlignment.TopLeft;
            SetNoteTableText(table, row, column, textHeight, textStyleId, textColor);
          }
        }
      }
    }

    /// <summary>
    /// Sets every row to its content plus the gap below each note. Text is
    /// measured with a temporary MText of the cell's width because a table
    /// row's MinimumHeight leaves out the text of text-only cells, which ran
    /// general notes together.
    /// </summary>
    private static void FitNoteTableRowHeights(
      Table table,
      Database database,
      NoteTableKind kind)
    {
      double scale = GetNoteTableScale(table);
      double gap = NoteTableNoteGap * scale;
      double hexHeight = KeyNoteHexHeight * scale;
      bool hasTitle = IsNoteTableTitleRow(table, 0);
      for (int row = 0; row < table.Rows.Count; row++)
      {
        double needed = gap;
        if (hasTitle && row == 0)
        {
          needed = MeasureNoteTableCell(table, database, row, 0, table.Columns.Count - 1) + gap;
        }
        else
        {
          for (int column = 0; column + 1 < table.Columns.Count; column += 2)
          {
            int textColumn = column + 1;
            bool hasSymbol = kind == NoteTableKind.Keyed &&
              table.GetNumberOfContents(row, column) > 0 &&
              !table.GetBlockTableRecordId(row, column, 0).IsNull;
            if (hasSymbol)
            {
              needed = Math.Max(needed, hexHeight + gap);
            }
            else if (kind == NoteTableKind.General)
            {
              needed = Math.Max(needed, MeasureNoteTableCell(table, database, row, column, column) + gap);
            }
            double textTop = table.Cells[row, textColumn].Borders.Top.Margin ?? 0.0;
            double text = MeasureNoteTableCell(table, database, row, textColumn, textColumn);
            if (text > 0.0)
            {
              needed = Math.Max(needed, textTop + text + gap);
            }
          }
        }
        table.Rows[row].Height = needed;
      }
    }

    /// <summary>
    /// Text height / 3/32" for the first note text, which ApplyNoteTableStyle
    /// made the same for every note.
    /// </summary>
    private static double GetNoteTableScale(Table table)
    {
      return TryGetNoteTableTextHeight(table, out double height)
        ? height / NoteTableTextHeight
        : 1.0;
    }

    /// <summary>The text height of the table's first note, if it has one.</summary>
    private static bool TryGetNoteTableTextHeight(Table table, out double height)
    {
      int firstRow = IsNoteTableTitleRow(table, 0) ? 1 : 0;
      for (int row = firstRow; row < table.Rows.Count; row++)
      {
        for (int column = 1; column < table.Columns.Count; column += 2)
        {
          if (table.GetNumberOfContents(row, column) > 0 &&
              !string.IsNullOrWhiteSpace(table.GetTextString(row, column, 0)))
          {
            height = table.GetTextHeight(row, column, 0);
            if (height > 1e-9)
            {
              return true;
            }
          }
        }
      }
      height = 0.0;
      return false;
    }

    /// <summary>
    /// The height of a cell's text wrapped to the cells' width less their
    /// side margins; 0 for an empty or block cell.
    /// </summary>
    private static double MeasureNoteTableCell(
      Table table,
      Database database,
      int row,
      int firstColumn,
      int lastColumn)
    {
      if (table.GetNumberOfContents(row, firstColumn) <= 0 ||
          !table.GetBlockTableRecordId(row, firstColumn, 0).IsNull)
      {
        return 0.0;
      }
      string contents = table.GetTextString(row, firstColumn, 0);
      if (string.IsNullOrWhiteSpace(contents))
      {
        return 0.0;
      }

      double width = 0.0;
      for (int column = firstColumn; column <= lastColumn; column++)
      {
        width += table.Columns[column].Width;
      }
      CellBorders borders = table.Cells[row, firstColumn].Borders;
      width -= (borders.Left.Margin ?? 0.0) + (borders.Right.Margin ?? 0.0);

      using (var text = new MText())
      {
        text.SetDatabaseDefaults(database);
        text.TextStyleId = table.GetTextStyleId(row, firstColumn, 0);
        text.TextHeight = table.GetTextHeight(row, firstColumn, 0);
        text.Width = Math.Max(width, text.TextHeight);
        text.Contents = contents;
        return text.ActualHeight;
      }
    }

    private static void SetNoteTableText(
      Table table,
      int row,
      int column,
      double textHeight,
      ObjectId textStyleId,
      Color textColor)
    {
      if (table.GetNumberOfContents(row, column) <= 0)
      {
        return;
      }
      if (!textStyleId.IsNull)
      {
        table.SetTextStyleId(row, column, 0, textStyleId);
      }
      table.SetTextHeight(row, column, 0, textHeight);
      table.SetContentColor(row, column, 0, textColor);
    }
  }
}
