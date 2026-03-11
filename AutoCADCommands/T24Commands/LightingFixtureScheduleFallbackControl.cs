using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ElectricalCommands
{
  internal sealed class LightingFixtureScheduleFallbackControl :
    UserControl,
    ILightingFixtureScheduleEditorControl
  {
    private TextBox _rowsText;
    private TextBox _generalNotes;
    private TextBox _notes;
    private Label _projectValue;
    private Label _dwgValue;
    private Label _tableValue;
    private Label _versionValue;
    private Label _statusValue;
    private Timer _saveTimer;
    private bool _suppressEvents;

    internal LightingFixtureScheduleFallbackControl()
    {
      ExecuteStage(
        "base control initialization",
        () =>
        {
          Dock = DockStyle.Fill;
          BackColor = System.Drawing.SystemColors.Window;
        }
      );

      TableLayoutPanel layout = ExecuteStage("root layout creation", CreateRootLayout);
      ExecuteStage("metadata panel creation", () => InitializeMetadata(layout));
      ExecuteStage("instructions creation", () => InitializeInstructions(layout));
      ExecuteStage("row editor creation", () => InitializeRows(layout));
      ExecuteStage("general notes editor creation", () => InitializeGeneralNotes(layout));
      ExecuteStage("notes editor creation", () => InitializeNotes(layout));
      ExecuteStage("footer creation", () => InitializeFooter(layout));
      ExecuteStage("root layout attachment", () => Controls.Add(layout));
      ExecuteStage("save timer creation", InitializeSaveTimer);
    }

    public event EventHandler<LightingFixtureScheduleSaveRequestedEventArgs> SaveRequested;
    public event EventHandler ReloadRequested;
    public event EventHandler LinkTableRequested;
    public event EventHandler CopyFromTableRequested;
    public event EventHandler<LightingFixtureScheduleSaveRequestedEventArgs> PlaceTableRequested;

    public bool HasPendingLocalEdits { get; private set; }
    public long LoadedVersion { get; private set; }
    public Control HostControl => this;

    public void LoadRecord(
      LightingFixtureScheduleBindingContext context,
      LightingFixtureScheduleStoreRecord record,
      string statusMessage
    )
    {
      _suppressEvents = true;
      try
      {
        _projectValue.Text = context?.ProjectId ?? string.Empty;
        _dwgValue.Text = context?.DwgPath ?? string.Empty;
        _tableValue.Text = string.IsNullOrWhiteSpace(context?.TableHandle)
          ? "(not linked)"
          : context.TableHandle;
        LoadedVersion = record?.Version ?? 0;
        _versionValue.Text = LoadedVersion > 0 ? LoadedVersion.ToString() : "0";

        LightingFixtureScheduleSyncSchedule schedule =
          record?.Schedule ?? GeneralCommands.CreateDefaultLightingFixtureSchedule();
        _rowsText.Text = SerializeRows(schedule.Rows);
        _generalNotes.Text = schedule?.GeneralNotes ?? string.Empty;
        _notes.Text = schedule?.Notes ?? string.Empty;
        HasPendingLocalEdits = false;
        _saveTimer.Stop();
        SetStatus(string.IsNullOrWhiteSpace(statusMessage) ? "Ready." : statusMessage);
      }
      finally
      {
        _suppressEvents = false;
      }
    }

    public void SetStatus(string message)
    {
      _statusValue.Text = string.IsNullOrWhiteSpace(message) ? "Ready." : message;
    }

    public void MarkExternalUpdateAvailable(long version)
    {
      if (HasPendingLocalEdits)
      {
        SetStatus(
          $"Central store has version {version}. Saving local edits will overwrite it."
        );
      }
    }

    private TableLayoutPanel CreateRootLayout()
    {
      var layout = new TableLayoutPanel
      {
        Dock = DockStyle.Fill,
        ColumnCount = 1,
        RowCount = 7,
        Padding = new Padding(8),
      };
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
      layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 112F));
      layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 96F));
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      return layout;
    }

    private void InitializeMetadata(TableLayoutPanel layout)
    {
      var metaPanel = new TableLayoutPanel
      {
        Dock = DockStyle.Top,
        ColumnCount = 2,
        AutoSize = true,
      };
      metaPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      metaPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

      _projectValue = AddMetaRow(metaPanel, 0, "Project");
      _dwgValue = AddMetaRow(metaPanel, 1, "DWG");
      _tableValue = AddMetaRow(metaPanel, 2, "Table");
      _versionValue = AddMetaRow(metaPanel, 3, "Version");
      layout.Controls.Add(metaPanel, 0, 0);
    }

    private void InitializeInstructions(TableLayoutPanel layout)
    {
      layout.Controls.Add(
        new Label
        {
          AutoSize = true,
          Dock = DockStyle.Top,
          Text =
            "Simplified editor. One row per line. Use TAB between columns: MARK, DESCRIPTION, MANUFACTURER, MODEL NUMBER, MOUNTING, VOLTS, WATTS, NOTES.",
          Margin = new Padding(0, 0, 0, 6),
        },
        0,
        1
      );
    }

    private void InitializeRows(TableLayoutPanel layout)
    {
      _rowsText = new TextBox
      {
        Dock = DockStyle.Fill,
        Multiline = true,
        ScrollBars = ScrollBars.Both,
        AcceptsReturn = true,
        AcceptsTab = true,
        WordWrap = false,
      };
      _rowsText.TextChanged += (_, __) => OnLocalEdit();
      layout.Controls.Add(_rowsText, 0, 2);
    }

    private void InitializeGeneralNotes(TableLayoutPanel layout)
    {
      _generalNotes = new TextBox
      {
        Dock = DockStyle.Fill,
        Multiline = true,
        ScrollBars = ScrollBars.Vertical,
      };
      _generalNotes.TextChanged += (_, __) => OnLocalEdit();
      layout.Controls.Add(WrapTextArea("General Notes", _generalNotes), 0, 3);
    }

    private void InitializeNotes(TableLayoutPanel layout)
    {
      _notes = new TextBox
      {
        Dock = DockStyle.Fill,
        Multiline = true,
        ScrollBars = ScrollBars.Vertical,
      };
      _notes.TextChanged += (_, __) => OnLocalEdit();
      layout.Controls.Add(WrapTextArea("Notes", _notes), 0, 4);
    }

    private void InitializeFooter(TableLayoutPanel layout)
    {
      var footer = new TableLayoutPanel
      {
        Dock = DockStyle.Top,
        ColumnCount = 6,
        AutoSize = true,
      };
      footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

      _statusValue = new Label
      {
        Dock = DockStyle.Fill,
        AutoSize = true,
        ForeColor = System.Drawing.SystemColors.GrayText,
        Text = "Ready.",
        Margin = new Padding(0, 8, 8, 0),
      };

      var copyFromTableButton = new Button
      {
        AutoSize = true,
        Text = "Copy From Table",
        Margin = new Padding(4, 0, 0, 0),
      };
      copyFromTableButton.Click += (_, __) => CopyFromTableRequested?.Invoke(this, EventArgs.Empty);

      var placeTableButton = new Button
      {
        AutoSize = true,
        Text = "Place New Table",
        Margin = new Padding(4, 0, 0, 0),
      };
      placeTableButton.Click += (_, __) =>
      {
        _saveTimer.Stop();
        RaisePlaceTableRequested();
      };

      var linkTableButton = new Button
      {
        AutoSize = true,
        Text = "Link Table",
        Margin = new Padding(4, 0, 0, 0),
      };
      linkTableButton.Click += (_, __) => LinkTableRequested?.Invoke(this, EventArgs.Empty);

      var reloadButton = new Button
      {
        AutoSize = true,
        Text = "Reload",
        Margin = new Padding(4, 0, 0, 0),
      };
      reloadButton.Click += (_, __) => ReloadRequested?.Invoke(this, EventArgs.Empty);

      var saveButton = new Button
      {
        AutoSize = true,
        Text = "Save Now",
        Margin = new Padding(4, 0, 0, 0),
      };
      saveButton.Click += (_, __) =>
      {
        _saveTimer.Stop();
        RaiseSaveRequested();
      };

      footer.Controls.Add(_statusValue, 0, 0);
      footer.Controls.Add(copyFromTableButton, 1, 0);
      footer.Controls.Add(placeTableButton, 2, 0);
      footer.Controls.Add(linkTableButton, 3, 0);
      footer.Controls.Add(reloadButton, 4, 0);
      footer.Controls.Add(saveButton, 5, 0);
      layout.Controls.Add(footer, 0, 5);

      layout.Controls.Add(
        new Label
        {
          AutoSize = true,
          Dock = DockStyle.Top,
          ForeColor = System.Drawing.SystemColors.GrayText,
          Text = "Tabs inside a NOTES cell are preserved only in the final column of each line.",
          Margin = new Padding(0, 6, 0, 0),
        },
        0,
        6
      );
    }

    private void InitializeSaveTimer()
    {
      _saveTimer = new Timer { Interval = 600 };
      _saveTimer.Tick += (_, __) =>
      {
        _saveTimer.Stop();
        RaiseSaveRequested();
      };
    }

    private void OnLocalEdit()
    {
      if (_suppressEvents)
      {
        return;
      }

      HasPendingLocalEdits = true;
      SetStatus("Unsaved AutoCAD edits.");
      _saveTimer.Stop();
      _saveTimer.Start();
    }

    private void RaiseSaveRequested()
    {
      if (_suppressEvents)
      {
        return;
      }

      SaveRequested?.Invoke(
        this,
        new LightingFixtureScheduleSaveRequestedEventArgs
        {
          Schedule = BuildSchedule(),
          LoadedVersion = LoadedVersion,
        }
      );
    }

    private void RaisePlaceTableRequested()
    {
      if (_suppressEvents)
      {
        return;
      }

      PlaceTableRequested?.Invoke(
        this,
        new LightingFixtureScheduleSaveRequestedEventArgs
        {
          Schedule = BuildSchedule(),
          LoadedVersion = LoadedVersion,
        }
      );
    }

    private LightingFixtureScheduleSyncSchedule BuildSchedule()
    {
      var schedule = new LightingFixtureScheduleSyncSchedule
      {
        Rows = ParseRows(_rowsText.Text),
        GeneralNotes = _generalNotes.Text ?? string.Empty,
        Notes = _notes.Text ?? string.Empty,
      };
      return GeneralCommands.NormalizeLightingFixtureScheduleForExternalUse(schedule);
    }

    private static List<LightingFixtureScheduleSyncRow> ParseRows(string text)
    {
      var rows = new List<LightingFixtureScheduleSyncRow>();
      string normalized = (text ?? string.Empty)
        .Replace("\r\n", "\n")
        .Replace('\r', '\n');
      foreach (string line in normalized.Split(new[] { '\n' }, StringSplitOptions.None))
      {
        if (string.IsNullOrWhiteSpace(line))
        {
          continue;
        }

        string[] parts = line.Split(new[] { '\t' });
        rows.Add(
          new LightingFixtureScheduleSyncRow
          {
            Mark = GetPart(parts, 0),
            Description = GetPart(parts, 1),
            Manufacturer = GetPart(parts, 2),
            ModelNumber = GetPart(parts, 3),
            Mounting = GetPart(parts, 4),
            Volts = GetPart(parts, 5),
            Watts = GetPart(parts, 6),
            Notes = parts.Length > 7 ? string.Join("\t", parts.Skip(7)) : string.Empty,
          }
        );
      }
      return rows;
    }

    private static string SerializeRows(IReadOnlyList<LightingFixtureScheduleSyncRow> rows)
    {
      if (rows == null || rows.Count == 0)
      {
        return string.Empty;
      }

      var builder = new StringBuilder();
      for (int index = 0; index < rows.Count; index++)
      {
        LightingFixtureScheduleSyncRow row = rows[index] ?? new LightingFixtureScheduleSyncRow();
        string[] parts =
        {
          row.Mark ?? string.Empty,
          row.Description ?? string.Empty,
          row.Manufacturer ?? string.Empty,
          row.ModelNumber ?? string.Empty,
          row.Mounting ?? string.Empty,
          row.Volts ?? string.Empty,
          row.Watts ?? string.Empty,
          row.Notes ?? string.Empty,
        };

        builder.Append(string.Join("\t", parts));
        if (index < rows.Count - 1)
        {
          builder.AppendLine();
        }
      }

      return builder.ToString();
    }

    private static string GetPart(string[] parts, int index)
    {
      if (parts == null || index < 0 || index >= parts.Length)
      {
        return string.Empty;
      }

      return parts[index] ?? string.Empty;
    }

    private static Label AddMetaRow(TableLayoutPanel panel, int rowIndex, string label)
    {
      panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

      var header = new Label
      {
        Text = label,
        AutoSize = true,
        Margin = new Padding(0, 0, 8, 4),
      };
      var value = new Label
      {
        AutoSize = true,
        Margin = new Padding(0, 0, 0, 4),
      };

      panel.Controls.Add(header, 0, rowIndex);
      panel.Controls.Add(value, 1, rowIndex);
      return value;
    }

    private static Control WrapTextArea(string label, TextBox textBox)
    {
      var panel = new TableLayoutPanel
      {
        Dock = DockStyle.Fill,
        ColumnCount = 1,
        RowCount = 2,
        Margin = new Padding(0, 8, 0, 0),
      };
      panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
      panel.Controls.Add(
        new Label
        {
          Text = label,
          Dock = DockStyle.Top,
          AutoSize = true,
          Margin = new Padding(0, 0, 0, 4),
        },
        0,
        0
      );
      panel.Controls.Add(textBox, 0, 1);
      return panel;
    }

    private static void ExecuteStage(string stage, Action action)
    {
      try
      {
        action();
      }
      catch (System.Exception ex)
      {
        throw new InvalidOperationException(
          $"Lighting fixture schedule fallback editor failed during {stage}.",
          ex
        );
      }
    }

    private static T ExecuteStage<T>(string stage, Func<T> action)
    {
      try
      {
        return action();
      }
      catch (System.Exception ex)
      {
        throw new InvalidOperationException(
          $"Lighting fixture schedule fallback editor failed during {stage}.",
          ex
        );
      }
    }
  }
}
