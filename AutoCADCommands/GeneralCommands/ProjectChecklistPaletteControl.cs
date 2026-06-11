using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ElectricalCommands
{
  internal interface IProjectChecklistEditorControl : IDisposable
  {
    event EventHandler<ProjectChecklistSaveRequestedEventArgs> SaveRequested;
    event EventHandler ReloadRequested;

    bool HasPendingLocalEdits { get; }
    long LoadedVersion { get; }
    Control HostControl { get; }

    void LoadRecord(
      ProjectChecklistBindingContext context,
      IList<ProjectChecklistDefinition> definitions,
      ProjectChecklistStoreRecord record,
      string statusMessage
    );

    void SetStatus(string message);
    void FlushPendingSave();
  }

  internal sealed class ProjectChecklistPaletteControl :
    UserControl,
    IProjectChecklistEditorControl
  {
    private readonly ComboBox _checklistSelector;
    private readonly FlowLayoutPanel _itemsPanel;
    private readonly Label _projectValue;
    private readonly Label _versionValue;
    private readonly Label _definitionsValue;
    private readonly Label _statusValue;
    private readonly Timer _saveTimer;
    private readonly Timer _resizeTimer;
    private bool _suppressEvents;
    private List<ProjectChecklistDefinition> _definitions =
      new List<ProjectChecklistDefinition>();
    private ProjectChecklistState _state = new ProjectChecklistState();
    private string _selectedChecklistId = string.Empty;
    private string _emptyDefinitionMessage = string.Empty;

    internal ProjectChecklistPaletteControl()
    {
      Dock = DockStyle.Fill;
      BackColor = SystemColors.Window;

      var layout = new TableLayoutPanel
      {
        Dock = DockStyle.Fill,
        ColumnCount = 1,
        RowCount = 4,
        Padding = new Padding(8),
      };
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

      var metadata = new TableLayoutPanel
      {
        Dock = DockStyle.Top,
        ColumnCount = 2,
        AutoSize = true,
      };
      metadata.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      metadata.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      _projectValue = AddMetadataRow(metadata, 0, "Project");
      _versionValue = AddMetadataRow(metadata, 1, "Version");
      _definitionsValue = AddMetadataRow(metadata, 2, "Definitions");
      layout.Controls.Add(metadata, 0, 0);

      _checklistSelector = new ComboBox
      {
        Dock = DockStyle.Top,
        DropDownStyle = ComboBoxStyle.DropDownList,
        Margin = new Padding(0, 8, 0, 8),
      };
      _checklistSelector.SelectedIndexChanged += (_, __) =>
      {
        if (_suppressEvents)
        {
          return;
        }

        _selectedChecklistId =
          (_checklistSelector.SelectedItem as ProjectChecklistComboItem)
            ?.Definition
            ?.Id
          ?? string.Empty;
        RenderItems();
      };
      layout.Controls.Add(_checklistSelector, 0, 1);

      _resizeTimer = new Timer { Interval = 80 };
      _resizeTimer.Tick += (_, __) =>
      {
        _resizeTimer.Stop();
        RefreshItemLayout();
      };

      _itemsPanel = new FlowLayoutPanel
      {
        Dock = DockStyle.Fill,
        AutoScroll = true,
        BorderStyle = BorderStyle.FixedSingle,
        FlowDirection = System.Windows.Forms.FlowDirection.TopDown,
        Padding = new Padding(8),
        WrapContents = false,
      };
      _itemsPanel.SizeChanged += (_, __) =>
      {
        if (!_suppressEvents)
        {
          QueueLayoutRefresh();
        }
      };
      layout.Controls.Add(_itemsPanel, 0, 2);

      var footer = new TableLayoutPanel
      {
        Dock = DockStyle.Bottom,
        ColumnCount = 3,
        AutoSize = true,
      };
      footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

      _statusValue = new Label
      {
        Dock = DockStyle.Fill,
        AutoSize = true,
        ForeColor = SystemColors.GrayText,
        Margin = new Padding(0, 8, 8, 0),
        Text = "Ready.",
      };
      var reloadButton = new Button
      {
        AutoSize = true,
        Margin = new Padding(4, 0, 0, 0),
        Text = "Reload",
      };
      reloadButton.Click += (_, __) => ReloadRequested?.Invoke(this, EventArgs.Empty);

      var saveButton = new Button
      {
        AutoSize = true,
        Margin = new Padding(4, 0, 0, 0),
        Text = "Save Now",
      };
      saveButton.Click += (_, __) =>
      {
        _saveTimer.Stop();
        RaiseSaveRequested();
      };

      footer.Controls.Add(_statusValue, 0, 0);
      footer.Controls.Add(reloadButton, 1, 0);
      footer.Controls.Add(saveButton, 2, 0);
      layout.Controls.Add(footer, 0, 3);

      Controls.Add(layout);

      _saveTimer = new Timer { Interval = 600 };
      _saveTimer.Tick += (_, __) =>
      {
        _saveTimer.Stop();
        RaiseSaveRequested();
      };

    }

    public event EventHandler<ProjectChecklistSaveRequestedEventArgs> SaveRequested;
    public event EventHandler ReloadRequested;

    public bool HasPendingLocalEdits { get; private set; }
    public long LoadedVersion { get; private set; }
    public Control HostControl => this;

    public void LoadRecord(
      ProjectChecklistBindingContext context,
      IList<ProjectChecklistDefinition> definitions,
      ProjectChecklistStoreRecord record,
      string statusMessage
    )
    {
      _suppressEvents = true;
      try
      {
        _definitions = (definitions ?? new List<ProjectChecklistDefinition>()).ToList();
        _state = ProjectChecklistStore.NormalizeState(record?.State);
        LoadedVersion = record?.Version ?? 0;
        _emptyDefinitionMessage = _definitions.Count == 0
          ? statusMessage
          : string.Empty;

        _projectValue.Text = context?.ProjectId ?? string.Empty;
        _versionValue.Text = LoadedVersion > 0 ? LoadedVersion.ToString() : "0";
        _definitionsValue.Text = $"{_definitions.Count} checklist(s)";

        string previousSelection = _selectedChecklistId;
        _checklistSelector.Items.Clear();
        foreach (ProjectChecklistDefinition definition in _definitions)
        {
          _checklistSelector.Items.Add(new ProjectChecklistComboItem(definition));
        }

        int selectedIndex = _definitions.FindIndex(
          definition => string.Equals(
            definition.Id,
            previousSelection,
            StringComparison.OrdinalIgnoreCase
          )
        );
        if (selectedIndex < 0 && _definitions.Count > 0)
        {
          selectedIndex = 0;
        }

        if (selectedIndex >= 0)
        {
          _checklistSelector.SelectedIndex = selectedIndex;
          _selectedChecklistId = _definitions[selectedIndex].Id;
        }
        else
        {
          _selectedChecklistId = string.Empty;
        }

        RenderItems();
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

    public void FlushPendingSave()
    {
      if (!HasPendingLocalEdits)
      {
        return;
      }

      _saveTimer.Stop();
      RaiseSaveRequested();
    }

    private static Label AddMetadataRow(TableLayoutPanel panel, int row, string label)
    {
      panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      panel.Controls.Add(
        new Label
        {
          AutoSize = true,
          Font = new System.Drawing.Font(SystemFonts.DefaultFont, FontStyle.Bold),
          Margin = new Padding(0, 0, 8, 3),
          Text = $"{label}:",
        },
        0,
        row
      );

      var value = new Label
      {
        AutoSize = true,
        Dock = DockStyle.Left,
        Margin = new Padding(0, 0, 0, 3),
        Text = string.Empty,
      };
      panel.Controls.Add(value, 1, row);
      return value;
    }

    private void RenderItems()
    {
      _itemsPanel.SuspendLayout();
      try
      {
        _itemsPanel.Controls.Clear();

        ProjectChecklistDefinition definition = _definitions.FirstOrDefault(
          item => string.Equals(
            item.Id,
            _selectedChecklistId,
            StringComparison.OrdinalIgnoreCase
          )
        );

        if (definition == null)
        {
          _itemsPanel.Controls.Add(CreateEmptyLabel(
            string.IsNullOrWhiteSpace(_emptyDefinitionMessage)
              ? "No checklists found. Create checklists in the desktop app first."
              : _emptyDefinitionMessage
          ));
          return;
        }

        List<ProjectChecklistDefinitionItem> items = definition.Items ?? new List<ProjectChecklistDefinitionItem>();
        if (items.Count == 0)
        {
          _itemsPanel.Controls.Add(CreateEmptyLabel("This checklist has no items."));
          return;
        }

        ProjectChecklistStateEntry stateEntry = EnsureStateEntry(definition.Id);
        var completed = new HashSet<string>(
          stateEntry.CompletedItems ?? new List<string>(),
          StringComparer.OrdinalIgnoreCase
        );

        foreach (ProjectChecklistDefinitionItem item in items)
        {
          Control row = item.IsSubheader
            ? CreateSubheader(item.Text)
            : CreateCheckboxRow(definition.Id, item, completed.Contains(item.Id));
          _itemsPanel.Controls.Add(row);
        }
      }
      finally
      {
        _itemsPanel.ResumeLayout();
      }
    }

    private Label CreateEmptyLabel(string text)
    {
      int width = GetChecklistRowWidth();
      Label label = new Label
      {
        AutoSize = false,
        ForeColor = SystemColors.GrayText,
        Margin = new Padding(0, 0, 0, 0),
        Text = text,
      };
      ApplyWrappedLabelSize(label, width);
      return label;
    }

    private Label CreateSubheader(string text)
    {
      string labelText = string.IsNullOrWhiteSpace(text) ? "Section" : text;
      int width = GetChecklistRowWidth();
      System.Drawing.Font font = new System.Drawing.Font(SystemFonts.DefaultFont, FontStyle.Bold);
      Label label = new Label
      {
        AutoSize = false,
        Font = font,
        Margin = new Padding(0, 10, 0, 4),
        Text = labelText,
      };
      ApplyWrappedLabelSize(label, width);
      return label;
    }

    private CheckBox CreateCheckboxRow(
      string checklistId,
      ProjectChecklistDefinitionItem item,
      bool isChecked
    )
    {
      string itemText = string.IsNullOrWhiteSpace(item.Text) ? "Untitled item" : item.Text;
      int width = GetChecklistRowWidth();
      var checkbox = new CheckBox
      {
        AutoSize = false,
        CheckAlign = ContentAlignment.TopLeft,
        Margin = new Padding(0, 3, 0, 3),
        Padding = new Padding(2),
        Text = itemText,
        TextAlign = ContentAlignment.TopLeft,
        Checked = isChecked,
        Tag = new Tuple<string, string>(checklistId, item.Id),
      };
      ApplyWrappedCheckboxSize(checkbox, width);
      checkbox.CheckedChanged += CheckboxChanged;
      return checkbox;
    }

    private void QueueLayoutRefresh()
    {
      if (_resizeTimer == null)
      {
        if (_itemsPanel != null && !_itemsPanel.IsDisposed && _itemsPanel.Controls.Count > 0)
        {
          RefreshItemLayout();
        }
        return;
      }

      _resizeTimer.Stop();
      _resizeTimer.Start();
    }

    private void RefreshItemLayout()
    {
      if (_itemsPanel == null || _itemsPanel.IsDisposed)
      {
        return;
      }

      int width = GetChecklistRowWidth();
      _itemsPanel.SuspendLayout();
      try
      {
        foreach (Control control in _itemsPanel.Controls)
        {
          if (control is CheckBox checkbox)
          {
            ApplyWrappedCheckboxSize(checkbox, width);
          }
          else if (control is Label label)
          {
            ApplyWrappedLabelSize(label, width);
          }
        }
      }
      finally
      {
        _itemsPanel.ResumeLayout();
      }
    }

    private static void ApplyWrappedLabelSize(Label label, int width)
    {
      label.Size = new Size(
        width,
        MeasureWrappedTextHeight(label.Text, label.Font, width)
      );
    }

    private static void ApplyWrappedCheckboxSize(CheckBox checkbox, int width)
    {
      int textWidth = Math.Max(40, width - 30);
      checkbox.Size = new Size(
        width,
        Math.Max(
          28,
          MeasureWrappedTextHeight(checkbox.Text, checkbox.Font, textWidth) + 16
        )
      );
    }

    private int GetChecklistRowWidth()
    {
      int scrollbarAllowance = _itemsPanel.VerticalScroll.Visible
        ? SystemInformation.VerticalScrollBarWidth
        : 0;
      return Math.Max(
        40,
        _itemsPanel.ClientSize.Width -
          _itemsPanel.Padding.Horizontal -
          scrollbarAllowance -
          2
      );
    }

    private static int MeasureWrappedTextHeight(string text, System.Drawing.Font font, int width)
    {
      Size proposed = new Size(Math.Max(40, width), int.MaxValue);
      Size measured = TextRenderer.MeasureText(
        string.IsNullOrWhiteSpace(text) ? " " : text,
        font,
        proposed,
        TextFormatFlags.WordBreak |
          TextFormatFlags.TextBoxControl |
          TextFormatFlags.NoPrefix
      );
      return Math.Max(font.Height + 8, measured.Height + 10);
    }

    private void CheckboxChanged(object sender, EventArgs e)
    {
      if (_suppressEvents)
      {
        return;
      }

      var checkbox = sender as CheckBox;
      var tag = checkbox?.Tag as Tuple<string, string>;
      if (checkbox == null || tag == null)
      {
        return;
      }

      ProjectChecklistStateEntry entry = EnsureStateEntry(tag.Item1);
      string itemId = ProjectChecklistStore.NormalizePlainText(tag.Item2);
      if (string.IsNullOrWhiteSpace(itemId))
      {
        return;
      }

      entry.CompletedItems = entry.CompletedItems ?? new List<string>();
      bool hasItem = entry.CompletedItems.Any(
        existing => string.Equals(existing, itemId, StringComparison.OrdinalIgnoreCase)
      );
      if (checkbox.Checked && !hasItem)
      {
        entry.CompletedItems.Add(itemId);
      }
      else if (!checkbox.Checked && hasItem)
      {
        entry.CompletedItems = entry.CompletedItems
          .Where(existing => !string.Equals(existing, itemId, StringComparison.OrdinalIgnoreCase))
          .ToList();
      }

      OnLocalEdit();
    }

    private ProjectChecklistStateEntry EnsureStateEntry(string checklistId)
    {
      string normalizedChecklistId = ProjectChecklistStore.NormalizePlainText(checklistId);
      ProjectChecklistStateEntry entry = _state.Checklists.FirstOrDefault(
        existing => string.Equals(
          existing.ChecklistId,
          normalizedChecklistId,
          StringComparison.OrdinalIgnoreCase
        )
      );
      if (entry != null)
      {
        return entry;
      }

      entry = new ProjectChecklistStateEntry
      {
        ChecklistId = normalizedChecklistId,
      };
      _state.Checklists.Add(entry);
      return entry;
    }

    private void OnLocalEdit()
    {
      HasPendingLocalEdits = true;
      SetStatus("Unsaved AutoCAD checklist edits.");
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
        new ProjectChecklistSaveRequestedEventArgs
        {
          State = ProjectChecklistStore.NormalizeState(_state),
          LoadedVersion = LoadedVersion,
        }
      );
    }

    private sealed class ProjectChecklistComboItem
    {
      internal ProjectChecklistComboItem(ProjectChecklistDefinition definition)
      {
        Definition = definition;
      }

      internal ProjectChecklistDefinition Definition { get; }

      public override string ToString()
      {
        return Definition?.Name ?? "Untitled checklist";
      }
    }
  }

  internal sealed class ProjectChecklistBindingContext
  {
    public Document Document { get; set; }
    public Database Database { get; set; }
    public Editor Editor { get; set; }
    public string ProjectId { get; set; }
    public string DwgPath { get; set; }
  }

  internal sealed class ProjectChecklistSaveRequestedEventArgs : EventArgs
  {
    public ProjectChecklistState State { get; set; }
    public long LoadedVersion { get; set; }
  }
}
