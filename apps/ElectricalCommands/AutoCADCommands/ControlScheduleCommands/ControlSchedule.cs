using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using AcApplication = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using AcColor = Autodesk.AutoCAD.Colors.Color;
using AcColorMethod = Autodesk.AutoCAD.Colors.ColorMethod;

namespace ElectricalCommands.ControlSchedules
{
  public sealed class ControlScheduleCommands : IExtensionApplication
  {
    private const string Title = "CONTROL SCHEDULE";

    // Styling matched to the LFS (Lighting Fixture Schedule) table template.
    private const string PreferredTextStyleName = "ARIALNARROW-1-8";
    private const double TitleTextHeight = 0.25;
    private const double HeaderTextHeight = 0.125;
    private const double DataTextHeight = 0.09375;
    private const double CellMargin = 0.06;
    private const short TextColorIndex = 4; // cyan
    private const short BorderColorIndex = 1; // red
    private static readonly Guid PaletteGuid = new Guid("7E6394E2-82DE-4D0E-B050-4D8DC05DA178");
    private static readonly string[] Headers =
    {
      "MARK", "DESCRIPTION", "MANUFACTURER", "MODEL NUMBER", "MOUNTING", "NOTES",
    };

    private static PaletteSet s_palette;
    private static ControlScheduleControl s_control;
    private static ControlScheduleData s_pendingSchedule;
    private static ObjectId s_linkedTableId = ObjectId.Null;
    private static Database s_linkedDatabase;

    public void Initialize() { }

    public void Terminate()
    {
      if (s_palette != null)
      {
        s_palette.Dispose();
        s_palette = null;
      }
      s_control = null;
    }

    [CommandMethod("CSCHED", CommandFlags.Modal)]
    public void OpenControlSchedule()
    {
      EnsurePalette();
      s_palette.Visible = true;
      AcApplication.DocumentManager.MdiActiveDocument?.Editor.WriteMessage(
        "\nCSCHED: Control Schedule editor opened."
      );
    }

    [CommandMethod("CONTROLSCHEDULE", CommandFlags.Modal)]
    public void OpenControlScheduleAlias()
    {
      OpenControlSchedule();
    }

    [CommandMethod("CSCHEDPLACE", CommandFlags.Modal | CommandFlags.NoHistory)]
    public void PlaceControlSchedule()
    {
      Document doc = AcApplication.DocumentManager.MdiActiveDocument;
      if (doc == null || s_pendingSchedule == null)
      {
        return;
      }

      PromptPointResult point = doc.Editor.GetPoint("\nSpecify Control Schedule insertion point: ");
      if (point.Status != PromptStatus.OK)
      {
        SetStatus("Placement canceled.");
        return;
      }

      try
      {
        ObjectId id = CreateTable(doc.Database, point.Value, s_pendingSchedule, null);
        LinkTable(doc.Database, id);
        doc.Editor.WriteMessage("\nControl Schedule table created.");
        SetStatus("Table placed and linked to this editor.");
      }
      catch (System.Exception ex)
      {
        doc.Editor.WriteMessage($"\nCSCHED placement error: {ex.Message}");
        SetStatus("Unable to place table. See the AutoCAD command line.");
      }
    }

    [CommandMethod("CSCHEDUPDATE", CommandFlags.Modal | CommandFlags.NoHistory)]
    public void UpdateControlSchedule()
    {
      Document doc = AcApplication.DocumentManager.MdiActiveDocument;
      if (doc == null || s_pendingSchedule == null)
      {
        return;
      }

      ObjectId oldId = ResolveLinkedTable(doc.Database);
      if (oldId.IsNull)
      {
        oldId = PromptForControlScheduleTable(doc.Editor, "\nSelect a Control Schedule table to update: ");
      }
      if (oldId.IsNull)
      {
        SetStatus("Update canceled.");
        return;
      }

      try
      {
        ObjectId newId = ReplaceTable(doc.Database, oldId, s_pendingSchedule);
        LinkTable(doc.Database, newId);
        doc.Editor.WriteMessage("\nControl Schedule table updated.");
        SetStatus("Linked table updated.");
      }
      catch (System.Exception ex)
      {
        doc.Editor.WriteMessage($"\nCSCHED update error: {ex.Message}");
        SetStatus("Unable to update table. See the AutoCAD command line.");
      }
    }

    [CommandMethod("CSCHEDLOAD", CommandFlags.Modal | CommandFlags.NoHistory)]
    public void LoadControlSchedule()
    {
      Document doc = AcApplication.DocumentManager.MdiActiveDocument;
      if (doc == null)
      {
        return;
      }

      ObjectId id = PromptForControlScheduleTable(
        doc.Editor,
        "\nSelect a Control Schedule table to load: "
      );
      if (id.IsNull)
      {
        SetStatus("Load canceled.");
        return;
      }

      try
      {
        ControlScheduleData data = ReadTable(doc.Database, id);
        EnsurePalette();
        s_control.LoadSchedule(data);
        LinkTable(doc.Database, id);
        s_palette.Visible = true;
        doc.Editor.WriteMessage("\nControl Schedule loaded into the editor.");
        SetStatus("Selected table loaded and linked.");
      }
      catch (System.Exception ex)
      {
        doc.Editor.WriteMessage($"\nCSCHED load error: {ex.Message}");
        SetStatus("The selected table is not a valid Control Schedule.");
      }
    }

    internal static void RequestAction(string command, ControlScheduleData schedule)
    {
      Document doc = AcApplication.DocumentManager.MdiActiveDocument;
      if (doc == null)
      {
        SetStatus("Open a drawing before using this action.");
        return;
      }
      s_pendingSchedule = schedule?.Clone() ?? ControlScheduleData.CreateDefault();
      doc.SendStringToExecute(command + " ", true, false, false);
    }

    private static void EnsurePalette()
    {
      if (s_palette != null)
      {
        return;
      }

      s_control = new ControlScheduleControl();
      s_palette = new PaletteSet("Control Schedule", PaletteGuid)
      {
        Style = PaletteSetStyles.ShowAutoHideButton |
          PaletteSetStyles.ShowCloseButton |
          PaletteSetStyles.ShowPropertiesMenu,
        DockEnabled = DockSides.Left | DockSides.Right,
        MinimumSize = new Size(760, 520),
      };
      s_palette.Add("Schedule", s_control);
    }

    private static void LinkTable(Database database, ObjectId id)
    {
      s_linkedDatabase = database;
      s_linkedTableId = id;
      s_control?.SetLinkedHandle(id.IsNull ? string.Empty : id.Handle.ToString());
    }

    private static ObjectId ResolveLinkedTable(Database database)
    {
      if (
        database == null ||
        database != s_linkedDatabase ||
        s_linkedTableId.IsNull ||
        !s_linkedTableId.IsValid ||
        s_linkedTableId.IsErased
      )
      {
        return ObjectId.Null;
      }
      return s_linkedTableId;
    }

    private static void SetStatus(string message)
    {
      s_control?.SetStatus(message);
    }

    private static ObjectId PromptForControlScheduleTable(Editor editor, string prompt)
    {
      var options = new PromptEntityOptions(prompt);
      options.SetRejectMessage("\nSelect an AutoCAD table.");
      options.AddAllowedClass(typeof(Table), true);
      PromptEntityResult result = editor.GetEntity(options);
      return result.Status == PromptStatus.OK ? result.ObjectId : ObjectId.Null;
    }

    private static ObjectId CreateTable(
      Database database,
      Point3d position,
      ControlScheduleData data,
      TableAppearance appearance
    )
    {
      using (Transaction tr = database.TransactionManager.StartTransaction())
      {
        BlockTableRecord space = (BlockTableRecord)tr.GetObject(
          database.CurrentSpaceId,
          OpenMode.ForWrite
        );
        Table table = BuildTable(database, tr, data, position, appearance);
        ObjectId id = space.AppendEntity(table);
        tr.AddNewlyCreatedDBObject(table, true);
        tr.Commit();
        return id;
      }
    }

    private static ObjectId ReplaceTable(Database database, ObjectId oldId, ControlScheduleData data)
    {
      using (Transaction tr = database.TransactionManager.StartTransaction())
      {
        Table oldTable = tr.GetObject(oldId, OpenMode.ForWrite) as Table;
        if (oldTable == null || !IsControlSchedule(oldTable))
        {
          throw new InvalidOperationException("The selected table is not a Control Schedule.");
        }

        var appearance = new TableAppearance
        {
          Layer = oldTable.Layer,
          Rotation = oldTable.Rotation,
          TableStyle = oldTable.TableStyle,
        };
        Table newTable = BuildTable(database, tr, data, oldTable.Position, appearance);
        BlockTableRecord owner = (BlockTableRecord)tr.GetObject(oldTable.OwnerId, OpenMode.ForWrite);
        ObjectId newId = owner.AppendEntity(newTable);
        tr.AddNewlyCreatedDBObject(newTable, true);
        oldTable.Erase();
        tr.Commit();
        return newId;
      }
    }

    private static Table BuildTable(
      Database database,
      Transaction tr,
      ControlScheduleData source,
      Point3d position,
      TableAppearance appearance
    )
    {
      ControlScheduleData data = source?.Normalize() ?? ControlScheduleData.CreateDefault();
      int dataRows = Math.Max(1, data.Rows.Count);
      int generalNotesRow = 2 + dataRows;
      int notesRow = generalNotesRow + 1;
      var table = new Table();
      table.SetDatabaseDefaults(database);
      table.TableStyle = appearance?.TableStyle ?? database.Tablestyle;
      if (!string.IsNullOrWhiteSpace(appearance?.Layer))
      {
        table.Layer = appearance.Layer;
      }
      table.Position = position;
      table.Rotation = appearance?.Rotation ?? 0.0;
      table.SetSize(notesRow + 1, Headers.Length);
      table.HorizontalCellMargin = CellMargin;
      table.VerticalCellMargin = CellMargin;

      double[] widths = { 1.45, 5.35, 2.75, 2.75, 2.25, 4.65 };
      for (int col = 0; col < Headers.Length; col++)
      {
        table.Columns[col].Width = widths[col];
      }
      for (int row = 0; row < table.Rows.Count; row++)
      {
        table.Rows[row].Height = row == 0 ? 0.6 : (row >= generalNotesRow ? 0.78 : 0.375);
      }

      table.MergeCells(CellRange.Create(table, 0, 0, 0, Headers.Length - 1));
      table.MergeCells(CellRange.Create(table, generalNotesRow, 0, generalNotesRow, Headers.Length - 1));
      table.MergeCells(CellRange.Create(table, notesRow, 0, notesRow, Headers.Length - 1));
      SetCell(table, 0, 0, Title, TitleTextHeight, CellAlignment.MiddleCenter);

      for (int col = 0; col < Headers.Length; col++)
      {
        SetCell(table, 1, col, Headers[col], HeaderTextHeight, CellAlignment.MiddleCenter);
      }

      for (int index = 0; index < dataRows; index++)
      {
        ControlScheduleRow item = index < data.Rows.Count ? data.Rows[index] : new ControlScheduleRow();
        string[] values =
        {
          item.Mark, item.Description, item.Manufacturer, item.ModelNumber, item.Mounting, item.Notes,
        };
        for (int col = 0; col < values.Length; col++)
        {
          SetCell(
            table,
            index + 2,
            col,
            values[col],
            DataTextHeight,
            col == 1 ? CellAlignment.MiddleLeft : CellAlignment.MiddleCenter
          );
        }
      }

      SetCell(
        table,
        generalNotesRow,
        0,
        "GENERAL NOTES:\n" + data.GeneralNotes,
        DataTextHeight,
        CellAlignment.MiddleLeft
      );
      SetCell(
        table,
        notesRow,
        0,
        "NOTES:\n" + data.Notes,
        DataTextHeight,
        CellAlignment.MiddleLeft
      );
      ApplyScheduleStyling(table, database, tr);
      table.GenerateLayout();
      return table;
    }

    private static void ApplyScheduleStyling(Table table, Database database, Transaction tr)
    {
      ObjectId textStyleId = ResolveTextStyleId(database, tr, PreferredTextStyleName);
      AcColor textColor = AcColor.FromColorIndex(AcColorMethod.ByAci, TextColorIndex);
      AcColor borderColor = AcColor.FromColorIndex(AcColorMethod.ByAci, BorderColorIndex);

      for (int row = 0; row < table.Rows.Count; row++)
      {
        for (int col = 0; col < table.Columns.Count; col++)
        {
          int r = row;
          int c = col;
          if (!textStyleId.IsNull)
          {
            TryApply(() => table.SetTextStyle(r, c, textStyleId));
          }
          TryApply(() => table.SetContentColor(r, c, textColor));
          foreach (GridLineType gridType in Enum.GetValues(typeof(GridLineType)))
          {
            if (gridType == GridLineType.InvalidGridLine)
            {
              continue;
            }
            GridLineType g = gridType;
            TryApply(() => table.SetGridColor(r, c, g, borderColor));
          }
        }
      }
    }

    private static ObjectId ResolveTextStyleId(Database database, Transaction tr, string styleName)
    {
      var textStyles = (TextStyleTable)tr.GetObject(database.TextStyleTableId, OpenMode.ForRead);
      if (textStyles.Has(styleName))
      {
        return textStyles[styleName];
      }
      foreach (ObjectId id in textStyles)
      {
        var record = tr.GetObject(id, OpenMode.ForRead) as TextStyleTableRecord;
        if (record != null && string.Equals(record.Name, styleName, StringComparison.OrdinalIgnoreCase))
        {
          return id;
        }
      }
      return database.Textstyle;
    }

    private static void TryApply(Action action)
    {
      // Best-effort styling, mirroring the LFS SafeApply behavior: some grid
      // line types are not settable on every cell (merged ranges, margins).
      try
      {
        action();
      }
      catch
      {
      }
    }

    private static void SetCell(
      Table table,
      int row,
      int col,
      string text,
      double textHeight,
      CellAlignment alignment
    )
    {
      table.Cells[row, col].TextString = text ?? string.Empty;
      table.Cells[row, col].TextHeight = textHeight;
      table.Cells[row, col].Alignment = alignment;
    }

    private static ControlScheduleData ReadTable(Database database, ObjectId id)
    {
      using (Transaction tr = database.TransactionManager.StartOpenCloseTransaction())
      {
        Table table = tr.GetObject(id, OpenMode.ForRead) as Table;
        if (table == null || !IsControlSchedule(table))
        {
          throw new InvalidOperationException("The selected table is not a Control Schedule.");
        }

        int generalNotesRow = table.Rows.Count - 2;
        var data = new ControlScheduleData();
        for (int row = 2; row < generalNotesRow; row++)
        {
          data.Rows.Add(
            new ControlScheduleRow
            {
              Mark = CellText(table, row, 0),
              Description = CellText(table, row, 1),
              Manufacturer = CellText(table, row, 2),
              ModelNumber = CellText(table, row, 3),
              Mounting = CellText(table, row, 4),
              Notes = CellText(table, row, 5),
            }
          );
        }
        data.GeneralNotes = RemoveHeading(CellText(table, generalNotesRow, 0), "GENERAL NOTES:");
        data.Notes = RemoveHeading(CellText(table, generalNotesRow + 1, 0), "NOTES:");
        return data.Normalize();
      }
    }

    private static bool IsControlSchedule(Table table)
    {
      return table != null &&
        table.Rows.Count >= 5 &&
        table.Columns.Count == Headers.Length &&
        string.Equals(CellText(table, 0, 0).Trim(), Title, StringComparison.OrdinalIgnoreCase) &&
        Headers.Select((header, col) => string.Equals(
          CellText(table, 1, col).Trim(), header, StringComparison.OrdinalIgnoreCase
        )).All(matches => matches);
    }

    private static string CellText(Table table, int row, int col)
    {
      try
      {
        return table.Cells[row, col].TextString ?? string.Empty;
      }
      catch
      {
        return string.Empty;
      }
    }

    private static string RemoveHeading(string text, string heading)
    {
      string value = text ?? string.Empty;
      return value.StartsWith(heading, StringComparison.OrdinalIgnoreCase)
        ? value.Substring(heading.Length).TrimStart('\r', '\n', ' ')
        : value;
    }

    private sealed class TableAppearance
    {
      internal string Layer { get; set; }
      internal double Rotation { get; set; }
      internal ObjectId TableStyle { get; set; }
    }
  }

  internal sealed class ControlScheduleControl : UserControl
  {
    private readonly DataGridView _grid;
    private readonly TextBox _generalNotes;
    private readonly TextBox _notes;
    private readonly Label _status;
    private readonly Label _linkedTable;

    internal ControlScheduleControl()
    {
      Dock = DockStyle.Fill;
      BackColor = Color.WhiteSmoke;
      var layout = new TableLayoutPanel
      {
        Dock = DockStyle.Fill,
        ColumnCount = 1,
        RowCount = 6,
        Padding = new Padding(8),
      };
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
      layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 82F));
      layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 82F));
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      Controls.Add(layout);

      var heading = new Label
      {
        AutoSize = true,
        Font = new System.Drawing.Font(Font.FontFamily, 14F, FontStyle.Bold),
        Text = "CONTROL SCHEDULE",
        Margin = new Padding(0, 0, 0, 8),
      };
      layout.Controls.Add(heading, 0, 0);

      _grid = new DataGridView
      {
        Dock = DockStyle.Fill,
        AllowUserToAddRows = true,
        AllowUserToDeleteRows = true,
        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
        RowHeadersVisible = false,
        BackgroundColor = Color.White,
      };
      AddColumn("Mark", 45F);
      AddColumn("Description", 150F);
      AddColumn("Manufacturer", 90F);
      AddColumn("Model Number", 90F);
      AddColumn("Mounting", 75F);
      AddColumn("Notes", 130F);
      layout.Controls.Add(_grid, 0, 1);

      _generalNotes = AddNotesBox(layout, 2, "GENERAL NOTES");
      _notes = AddNotesBox(layout, 3, "NOTES");

      var buttons = new FlowLayoutPanel
      {
        AutoSize = true,
        Dock = DockStyle.Fill,
        FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
        WrapContents = true,
        Margin = new Padding(0, 8, 0, 4),
      };
      buttons.Controls.Add(MakeButton("Load Selected", (_, __) => Request("CSCHEDLOAD")));
      buttons.Controls.Add(MakeButton("Place New Table", (_, __) => Request("CSCHEDPLACE")));
      buttons.Controls.Add(MakeButton("Update Linked / Selected", (_, __) => Request("CSCHEDUPDATE")));
      buttons.Controls.Add(MakeButton("ALL CAPS", (_, __) => ConvertAllTextToUppercase()));
      buttons.Controls.Add(MakeButton("Reset Defaults", (_, __) => LoadSchedule(ControlScheduleData.CreateDefault())));
      layout.Controls.Add(buttons, 0, 4);

      var footer = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 2 };
      footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      _status = new Label { AutoSize = true, Text = "Ready." };
      _linkedTable = new Label { AutoSize = true, Text = "Table: not linked" };
      footer.Controls.Add(_status, 0, 0);
      footer.Controls.Add(_linkedTable, 1, 0);
      layout.Controls.Add(footer, 0, 5);
      LoadSchedule(ControlScheduleData.CreateDefault());
    }

    internal void LoadSchedule(ControlScheduleData source)
    {
      ControlScheduleData data = source?.Normalize() ?? ControlScheduleData.CreateDefault();
      _grid.Rows.Clear();
      foreach (ControlScheduleRow row in data.Rows)
      {
        _grid.Rows.Add(
          row.Mark,
          row.Description,
          row.Manufacturer,
          row.ModelNumber,
          row.Mounting,
          row.Notes
        );
      }
      _generalNotes.Text = data.GeneralNotes;
      _notes.Text = data.Notes;
    }

    internal void SetStatus(string message)
    {
      _status.Text = message ?? string.Empty;
    }

    internal void SetLinkedHandle(string handle)
    {
      _linkedTable.Text = string.IsNullOrWhiteSpace(handle)
        ? "Table: not linked"
        : "Table handle: " + handle;
    }

    private void Request(string command)
    {
      ControlScheduleData data = BuildSchedule();
      ControlScheduleCommands.RequestAction(command, data);
      SetStatus("Switching to the AutoCAD command line...");
    }

    private ControlScheduleData BuildSchedule()
    {
      var data = new ControlScheduleData
      {
        GeneralNotes = _generalNotes.Text,
        Notes = _notes.Text,
      };
      foreach (DataGridViewRow gridRow in _grid.Rows)
      {
        if (gridRow.IsNewRow)
        {
          continue;
        }
        data.Rows.Add(
          new ControlScheduleRow
          {
            Mark = Value(gridRow, 0),
            Description = Value(gridRow, 1),
            Manufacturer = Value(gridRow, 2),
            ModelNumber = Value(gridRow, 3),
            Mounting = Value(gridRow, 4),
            Notes = Value(gridRow, 5),
          }
        );
      }
      return data.Normalize();
    }

    private void ConvertAllTextToUppercase()
    {
      foreach (DataGridViewRow row in _grid.Rows)
      {
        if (row.IsNewRow)
        {
          continue;
        }
        foreach (DataGridViewCell cell in row.Cells)
        {
          cell.Value = Convert.ToString(cell.Value)?.ToUpperInvariant() ?? string.Empty;
        }
      }
      _generalNotes.Text = (_generalNotes.Text ?? string.Empty).ToUpperInvariant();
      _notes.Text = (_notes.Text ?? string.Empty).ToUpperInvariant();
      SetStatus("Converted schedule text to uppercase.");
    }

    private static string Value(DataGridViewRow row, int column)
    {
      return Convert.ToString(row.Cells[column].Value) ?? string.Empty;
    }

    private void AddColumn(string text, float weight)
    {
      _grid.Columns.Add(
        new DataGridViewTextBoxColumn
        {
          HeaderText = text,
          Name = text.Replace(" ", string.Empty),
          FillWeight = weight,
          SortMode = DataGridViewColumnSortMode.NotSortable,
        }
      );
    }

    private static TextBox AddNotesBox(TableLayoutPanel parent, int row, string label)
    {
      var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
      panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
      panel.Controls.Add(new Label { AutoSize = true, Text = label, Font = new System.Drawing.Font(SystemFonts.MessageBoxFont, FontStyle.Bold) }, 0, 0);
      var box = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical };
      panel.Controls.Add(box, 0, 1);
      parent.Controls.Add(panel, 0, row);
      return box;
    }

    private static Button MakeButton(string text, EventHandler onClick)
    {
      var button = new Button { AutoSize = true, Text = text, Margin = new Padding(0, 0, 6, 0) };
      button.Click += onClick;
      return button;
    }
  }

  internal sealed class ControlScheduleData
  {
    internal List<ControlScheduleRow> Rows { get; set; } = new List<ControlScheduleRow>();
    internal string GeneralNotes { get; set; } = string.Empty;
    internal string Notes { get; set; } = string.Empty;

    internal static ControlScheduleData CreateDefault()
    {
      return new ControlScheduleData
      {
        Rows = new List<ControlScheduleRow>
        {
          Row("OS", "OCCUPANCY SENSOR - 360 DEGREE", "WATTSTOPPER", "LMPC-100", "CEILING", "DIGITAL PASSIVE IR"),
          Row("DS", "DIMMING PHOTO SENSOR", "WATTSTOPPER", "LMLS-400", "CEILING", "DIGITAL PASSIVE IR"),
          Row("$OS", "DIGITAL WALL BOX OCCUPANCY SENSOR", "WATTSTOPPER", "PW-100", "WALL", "IR TRANSCEIVER (2) RJ45 PORTS"),
          Row("$D", "DIGITAL WALL BOX DIMMING SWITCH", "WATTSTOPPER", "LMDM-101", "WALL", "IR TRANSCEIVER (2) RJ45 PORTS"),
          Row("$", "DIGITAL WALL SWITCH", "WATTSTOPPER", "LMSW-101", "WALL", "IR TRANSCEIVER (2) RJ45 PORTS"),
          Row("RC", "DIGITAL ON/OFF/0-10V DIMMING ROOM CONTROLLER", "WATTSTOPPER", "LMRC-211", "ABOVE CEILING", "0-10V, NOTES 1,2"),
        },
        GeneralNotes = "A. REFER TO ELECTRICAL SPECIFICATIONS AND ADDITIONAL REQUIREMENTS WHICH MAY NOT NECESSARILY BE REFLECTED IN CATALOG NUMBER AND/OR DESCRIPTION IN THE SCHEDULE.",
        Notes = "1. CONTRACTOR SHALL PROVIDE THE APPROPRIATE NUMBER OF ROOM CONTROLLERS FOR THE PROJECT. COORDINATE WITH THE MANUFACTURER FOR THE CORRECT QUANTITY.\r\n2. CONTRACTOR SHALL PROVIDE 4\"X4\" JUNCTION BOX AS REQUIRED.",
      };
    }

    internal ControlScheduleData Normalize()
    {
      Rows = (Rows ?? new List<ControlScheduleRow>())
        .Where(row => row != null && row.HasContent())
        .Select(row => row.Normalize())
        .ToList();
      GeneralNotes = GeneralNotes ?? string.Empty;
      Notes = Notes ?? string.Empty;
      return this;
    }

    internal ControlScheduleData Clone()
    {
      return new ControlScheduleData
      {
        Rows = (Rows ?? new List<ControlScheduleRow>()).Select(row => row.Clone()).ToList(),
        GeneralNotes = GeneralNotes,
        Notes = Notes,
      }.Normalize();
    }

    private static ControlScheduleRow Row(
      string mark,
      string description,
      string manufacturer,
      string model,
      string mounting,
      string notes
    )
    {
      return new ControlScheduleRow
      {
        Mark = mark,
        Description = description,
        Manufacturer = manufacturer,
        ModelNumber = model,
        Mounting = mounting,
        Notes = notes,
      };
    }
  }

  internal sealed class ControlScheduleRow
  {
    internal string Mark { get; set; } = string.Empty;
    internal string Description { get; set; } = string.Empty;
    internal string Manufacturer { get; set; } = string.Empty;
    internal string ModelNumber { get; set; } = string.Empty;
    internal string Mounting { get; set; } = string.Empty;
    internal string Notes { get; set; } = string.Empty;

    internal bool HasContent()
    {
      return new[] { Mark, Description, Manufacturer, ModelNumber, Mounting, Notes }
        .Any(value => !string.IsNullOrWhiteSpace(value));
    }

    internal ControlScheduleRow Normalize()
    {
      Mark = Mark ?? string.Empty;
      Description = Description ?? string.Empty;
      Manufacturer = Manufacturer ?? string.Empty;
      ModelNumber = ModelNumber ?? string.Empty;
      Mounting = Mounting ?? string.Empty;
      Notes = Notes ?? string.Empty;
      return this;
    }

    internal ControlScheduleRow Clone()
    {
      return new ControlScheduleRow
      {
        Mark = Mark,
        Description = Description,
        Manufacturer = Manufacturer,
        ModelNumber = ModelNumber,
        Mounting = Mounting,
        Notes = Notes,
      };
    }
  }
}
