using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ElectricalCommands
{
  internal interface ILightingFixtureScheduleEditorControl : IDisposable
  {
    event EventHandler<LightingFixtureScheduleSaveRequestedEventArgs> SaveRequested;
    event EventHandler ReloadRequested;
    event EventHandler LinkTableRequested;
    event EventHandler CopyFromTableRequested;
    event EventHandler<LightingFixtureScheduleSaveRequestedEventArgs> PlaceTableRequested;

    bool HasPendingLocalEdits { get; }
    long LoadedVersion { get; }
    Control HostControl { get; }

    void LoadRecord(
      LightingFixtureScheduleBindingContext context,
      LightingFixtureScheduleStoreRecord record,
      string statusMessage
    );

    void SetStatus(string message);
    void MarkExternalUpdateAvailable(long version);
  }

  internal sealed class LightingFixtureSchedulePaletteControl :
    UserControl,
    ILightingFixtureScheduleEditorControl
  {
    private readonly List<LightingFixtureScheduleRowCard> _rowCards =
      new List<LightingFixtureScheduleRowCard>();
    private Panel _rowsPanel;
    private TextBox _generalNotes;
    private TextBox _notes;
    private Label _projectValue;
    private Label _dwgValue;
    private Label _tableValue;
    private Label _versionValue;
    private Label _statusValue;
    private Timer _saveTimer;
    private bool _suppressEvents;

    internal LightingFixtureSchedulePaletteControl()
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
      ExecuteStage("row editor creation", () => InitializeRows(layout));
      ExecuteStage("general notes editor creation", () => InitializeGeneralNotes(layout));
      ExecuteStage("notes editor creation", () => InitializeNotes(layout));
      ExecuteStage("footer creation", () => InitializeFooter(layout));
      ExecuteStage("root layout attachment", () => Controls.Add(layout));
      ExecuteStage("save timer creation", InitializeSaveTimer);
      ExecuteStage("initial blank row", () => AddRowCard(new LightingFixtureScheduleSyncRow(), false));
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
        ResetRows(schedule.Rows);
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
        RowCount = 5,
        Padding = new Padding(8),
      };
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
      layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 112F));
      layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 96F));
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

    private void InitializeRows(TableLayoutPanel layout)
    {
      var rowsSection = new TableLayoutPanel
      {
        Dock = DockStyle.Fill,
        ColumnCount = 1,
        RowCount = 3,
        Margin = new Padding(0, 8, 0, 0),
      };
      rowsSection.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      rowsSection.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      rowsSection.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

      var rowsHeader = new TableLayoutPanel
      {
        Dock = DockStyle.Top,
        ColumnCount = 2,
        AutoSize = true,
      };
      rowsHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      rowsHeader.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      rowsHeader.Controls.Add(
        new Label
        {
          Text = "Fixtures",
          AutoSize = true,
          Dock = DockStyle.Left,
          Margin = new Padding(0, 6, 0, 4),
        },
        0,
        0
      );

      var addRowButton = new Button
      {
        AutoSize = true,
        Text = "Add Fixture",
        Margin = new Padding(4, 0, 0, 0),
      };
      addRowButton.Click += (_, __) =>
      {
        AddRowCard(new LightingFixtureScheduleSyncRow(), true);
      };
      rowsHeader.Controls.Add(addRowButton, 1, 0);
      rowsSection.Controls.Add(rowsHeader, 0, 0);

      rowsSection.Controls.Add(
        new Label
        {
          AutoSize = true,
          Dock = DockStyle.Top,
          ForeColor = System.Drawing.SystemColors.GrayText,
          Text = "Use the arrows to reorder fixtures. Each card maps directly to one schedule row.",
          Margin = new Padding(0, 0, 0, 6),
        },
        0,
        1
      );

      _rowsPanel = new Panel
      {
        Dock = DockStyle.Fill,
        AutoScroll = true,
        BorderStyle = BorderStyle.FixedSingle,
        Padding = new Padding(6),
      };
      rowsSection.Controls.Add(_rowsPanel, 0, 2);
      layout.Controls.Add(rowsSection, 0, 1);
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
      layout.Controls.Add(WrapTextArea("General Notes", _generalNotes), 0, 2);
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
      layout.Controls.Add(WrapTextArea("Notes", _notes), 0, 3);
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
      layout.Controls.Add(footer, 0, 4);
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

    private void ResetRows(IEnumerable<LightingFixtureScheduleSyncRow> rows)
    {
      ClearRowCards();
      List<LightingFixtureScheduleSyncRow> normalizedRows =
        (rows ?? Enumerable.Empty<LightingFixtureScheduleSyncRow>()).ToList();
      if (normalizedRows.Count == 0)
      {
        normalizedRows.Add(new LightingFixtureScheduleSyncRow());
      }

      foreach (LightingFixtureScheduleSyncRow row in normalizedRows)
      {
        AddRowCard(row, false);
      }

      RebuildRowCardContainer();
    }

    private void ClearRowCards()
    {
      foreach (LightingFixtureScheduleRowCard rowCard in _rowCards)
      {
        rowCard.Dispose();
      }

      _rowCards.Clear();
      _rowsPanel.Controls.Clear();
    }

    private void AddRowCard(LightingFixtureScheduleSyncRow row, bool markEdited)
    {
      var rowCard = new LightingFixtureScheduleRowCard();
      rowCard.Load(
        row ?? new LightingFixtureScheduleSyncRow(),
        _rowCards.Count + 1
      );
      rowCard.Changed += (_, __) => OnLocalEdit();
      rowCard.MoveUpRequested += (_, __) => MoveRowCard(rowCard, -1);
      rowCard.MoveDownRequested += (_, __) => MoveRowCard(rowCard, 1);
      rowCard.DeleteRequested += (_, __) => DeleteRowCard(rowCard);
      _rowCards.Add(rowCard);
      RebuildRowCardContainer();
      if (markEdited)
      {
        OnLocalEdit();
      }
    }

    private void MoveRowCard(LightingFixtureScheduleRowCard rowCard, int delta)
    {
      int currentIndex = _rowCards.IndexOf(rowCard);
      if (currentIndex < 0)
      {
        return;
      }

      int nextIndex = currentIndex + delta;
      if (nextIndex < 0 || nextIndex >= _rowCards.Count)
      {
        return;
      }

      _rowCards.RemoveAt(currentIndex);
      _rowCards.Insert(nextIndex, rowCard);
      RebuildRowCardContainer();
      OnLocalEdit();
    }

    private void DeleteRowCard(LightingFixtureScheduleRowCard rowCard)
    {
      if (_rowCards.Count <= 1)
      {
        rowCard.Load(new LightingFixtureScheduleSyncRow(), 1);
        OnLocalEdit();
        return;
      }

      _rowCards.Remove(rowCard);
      rowCard.Dispose();
      RebuildRowCardContainer();
      OnLocalEdit();
    }

    private void RebuildRowCardContainer()
    {
      _rowsPanel.SuspendLayout();
      try
      {
        _rowsPanel.Controls.Clear();
        for (int index = _rowCards.Count - 1; index >= 0; index--)
        {
          LightingFixtureScheduleRowCard rowCard = _rowCards[index];
          rowCard.SetDisplayIndex(index + 1, _rowCards.Count);
          _rowsPanel.Controls.Add(rowCard.Root);
        }
      }
      finally
      {
        _rowsPanel.ResumeLayout();
      }
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
        Rows = _rowCards.Select(rowCard => rowCard.Build()).ToList(),
        GeneralNotes = _generalNotes.Text ?? string.Empty,
        Notes = _notes.Text ?? string.Empty,
      };
      return GeneralCommands.NormalizeLightingFixtureScheduleForExternalUse(schedule);
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
          $"Lighting fixture schedule primary editor failed during {stage}.",
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
          $"Lighting fixture schedule primary editor failed during {stage}.",
          ex
        );
      }
    }
  }

  internal sealed class LightingFixtureScheduleRowCard : IDisposable
  {
    private readonly Label _titleLabel;
    private readonly Button _moveUpButton;
    private readonly Button _moveDownButton;
    private readonly Button _deleteButton;
    private readonly TextBox _markTextBox;
    private readonly TextBox _descriptionTextBox;
    private readonly TextBox _manufacturerTextBox;
    private readonly TextBox _modelNumberTextBox;
    private readonly TextBox _mountingTextBox;
    private readonly TextBox _voltsTextBox;
    private readonly TextBox _wattsTextBox;
    private readonly TextBox _notesTextBox;

    internal LightingFixtureScheduleRowCard()
    {
      Root = new Panel
      {
        Dock = DockStyle.Top,
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink,
        BorderStyle = BorderStyle.FixedSingle,
        Margin = new Padding(0, 0, 0, 8),
        Padding = new Padding(8),
      };

      var shell = new TableLayoutPanel
      {
        Dock = DockStyle.Top,
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink,
        ColumnCount = 1,
        RowCount = 5,
      };
      shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      Root.Controls.Add(shell);

      var header = new TableLayoutPanel
      {
        Dock = DockStyle.Top,
        ColumnCount = 5,
        AutoSize = true,
      };
      header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

      _titleLabel = new Label
      {
        AutoSize = true,
        Margin = new Padding(0, 6, 0, 4),
        Text = "Fixture",
      };

      _moveUpButton = new Button
      {
        AutoSize = true,
        Text = "Up",
        Margin = new Padding(4, 0, 0, 0),
      };
      _moveDownButton = new Button
      {
        AutoSize = true,
        Text = "Down",
        Margin = new Padding(4, 0, 0, 0),
      };
      _deleteButton = new Button
      {
        AutoSize = true,
        Text = "Delete",
        Margin = new Padding(4, 0, 0, 0),
      };

      _moveUpButton.Click += (_, __) => MoveUpRequested?.Invoke(this, EventArgs.Empty);
      _moveDownButton.Click += (_, __) => MoveDownRequested?.Invoke(this, EventArgs.Empty);
      _deleteButton.Click += (_, __) => DeleteRequested?.Invoke(this, EventArgs.Empty);

      header.Controls.Add(_titleLabel, 0, 0);
      header.Controls.Add(_moveUpButton, 1, 0);
      header.Controls.Add(_moveDownButton, 2, 0);
      header.Controls.Add(_deleteButton, 3, 0);
      shell.Controls.Add(header, 0, 0);

      _markTextBox = CreateTextBox();
      _descriptionTextBox = CreateTextBox();
      shell.Controls.Add(
        CreateFieldRow("Mark", _markTextBox, "Description", _descriptionTextBox),
        0,
        1
      );

      _manufacturerTextBox = CreateTextBox();
      _modelNumberTextBox = CreateTextBox();
      shell.Controls.Add(
        CreateFieldRow("Manufacturer", _manufacturerTextBox, "Model Number", _modelNumberTextBox),
        0,
        2
      );

      _mountingTextBox = CreateTextBox();
      _voltsTextBox = CreateTextBox();
      _wattsTextBox = CreateTextBox();
      shell.Controls.Add(
        CreateTripleFieldRow("Mounting", _mountingTextBox, "Volts", _voltsTextBox, "Watts", _wattsTextBox),
        0,
        3
      );

      _notesTextBox = CreateTextBox();
      shell.Controls.Add(CreateSingleFieldRow("Notes", _notesTextBox), 0, 4);

      HookChanged(_markTextBox);
      HookChanged(_descriptionTextBox);
      HookChanged(_manufacturerTextBox);
      HookChanged(_modelNumberTextBox);
      HookChanged(_mountingTextBox);
      HookChanged(_voltsTextBox);
      HookChanged(_wattsTextBox);
      HookChanged(_notesTextBox);
    }

    internal event EventHandler Changed;
    internal event EventHandler MoveUpRequested;
    internal event EventHandler MoveDownRequested;
    internal event EventHandler DeleteRequested;

    internal Panel Root { get; }

    internal void Load(LightingFixtureScheduleSyncRow row, int displayIndex)
    {
      LightingFixtureScheduleSyncRow source = row ?? new LightingFixtureScheduleSyncRow();
      _markTextBox.Text = source.Mark ?? string.Empty;
      _descriptionTextBox.Text = source.Description ?? string.Empty;
      _manufacturerTextBox.Text = source.Manufacturer ?? string.Empty;
      _modelNumberTextBox.Text = source.ModelNumber ?? string.Empty;
      _mountingTextBox.Text = source.Mounting ?? string.Empty;
      _voltsTextBox.Text = source.Volts ?? string.Empty;
      _wattsTextBox.Text = source.Watts ?? string.Empty;
      _notesTextBox.Text = source.Notes ?? string.Empty;
      _titleLabel.Text = $"Fixture {displayIndex}";
    }

    internal void SetDisplayIndex(int displayIndex, int totalCount)
    {
      _titleLabel.Text = $"Fixture {displayIndex}";
      _moveUpButton.Enabled = displayIndex > 1;
      _moveDownButton.Enabled = displayIndex < totalCount;
      _deleteButton.Enabled = totalCount > 1;
    }

    internal LightingFixtureScheduleSyncRow Build()
    {
      return new LightingFixtureScheduleSyncRow
      {
        Mark = _markTextBox.Text ?? string.Empty,
        Description = _descriptionTextBox.Text ?? string.Empty,
        Manufacturer = _manufacturerTextBox.Text ?? string.Empty,
        ModelNumber = _modelNumberTextBox.Text ?? string.Empty,
        Mounting = _mountingTextBox.Text ?? string.Empty,
        Volts = _voltsTextBox.Text ?? string.Empty,
        Watts = _wattsTextBox.Text ?? string.Empty,
        Notes = _notesTextBox.Text ?? string.Empty,
      };
    }

    public void Dispose()
    {
      if (Root != null && !Root.IsDisposed)
      {
        Root.Dispose();
      }
    }

    private void HookChanged(TextBox textBox)
    {
      textBox.TextChanged += (_, __) => Changed?.Invoke(this, EventArgs.Empty);
    }

    private static TextBox CreateTextBox()
    {
      return new TextBox
      {
        Dock = DockStyle.Fill,
        Margin = new Padding(0, 0, 0, 6),
      };
    }

    private static Control CreateFieldRow(
      string firstLabel,
      TextBox firstTextBox,
      string secondLabel,
      TextBox secondTextBox
    )
    {
      var row = new TableLayoutPanel
      {
        Dock = DockStyle.Top,
        ColumnCount = 2,
        AutoSize = true,
      };
      row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
      row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
      row.Controls.Add(CreateFieldPanel(firstLabel, firstTextBox), 0, 0);
      row.Controls.Add(CreateFieldPanel(secondLabel, secondTextBox), 1, 0);
      return row;
    }

    private static Control CreateTripleFieldRow(
      string firstLabel,
      TextBox firstTextBox,
      string secondLabel,
      TextBox secondTextBox,
      string thirdLabel,
      TextBox thirdTextBox
    )
    {
      var row = new TableLayoutPanel
      {
        Dock = DockStyle.Top,
        ColumnCount = 3,
        AutoSize = true,
      };
      row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45F));
      row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 27.5F));
      row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 27.5F));
      row.Controls.Add(CreateFieldPanel(firstLabel, firstTextBox), 0, 0);
      row.Controls.Add(CreateFieldPanel(secondLabel, secondTextBox), 1, 0);
      row.Controls.Add(CreateFieldPanel(thirdLabel, thirdTextBox), 2, 0);
      return row;
    }

    private static Control CreateSingleFieldRow(string label, TextBox textBox)
    {
      var row = new TableLayoutPanel
      {
        Dock = DockStyle.Top,
        ColumnCount = 1,
        AutoSize = true,
      };
      row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      row.Controls.Add(CreateFieldPanel(label, textBox), 0, 0);
      return row;
    }

    private static Control CreateFieldPanel(string label, TextBox textBox)
    {
      var panel = new TableLayoutPanel
      {
        Dock = DockStyle.Top,
        ColumnCount = 1,
        RowCount = 2,
        AutoSize = true,
        Margin = new Padding(0, 0, 8, 0),
      };
      panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      panel.Controls.Add(
        new Label
        {
          Text = label,
          Dock = DockStyle.Top,
          AutoSize = true,
          Margin = new Padding(0, 0, 0, 2),
        },
        0,
        0
      );
      panel.Controls.Add(textBox, 0, 1);
      return panel;
    }
  }

  internal sealed class LightingFixtureScheduleBindingContext
  {
    public Document Document { get; set; }
    public Database Database { get; set; }
    public Editor Editor { get; set; }
    public string ProjectId { get; set; }
    public string DwgPath { get; set; }
    public ObjectId TableId { get; set; }
    public string TableHandle { get; set; }
  }

  internal sealed class LightingFixtureScheduleSaveRequestedEventArgs : EventArgs
  {
    public LightingFixtureScheduleSyncSchedule Schedule { get; set; }
    public long LoadedVersion { get; set; }
  }
}
