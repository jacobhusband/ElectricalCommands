using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Windows;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using AcApplication = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace ElectricalCommands
{
  /// <summary>
  /// Modeless palette kept open while drafting: the drawing settings the
  /// electrical commands depend on, plus every drafting command grouped,
  /// searchable, and explained. It opens when AutoCAD starts.
  /// </summary>
  internal static class DraftingPalette
  {
    // Same GUID the Home Run Settings palette used, so an existing docked
    // position and size carry over.
    private static readonly Guid PaletteGuid = new Guid("F3CBAC46-1A30-4EB0-971F-4A33904A0D6B");
    private static PaletteSet _palette;
    private static DraftingPaletteControl _control;
    private static bool _documentEventsAttached;

    internal static void Show()
    {
      EnsureCreated();
      Refresh();
      _palette.Visible = true;
    }

    internal static void Refresh()
    {
      if (_control == null)
      {
        return;
      }

      Document document = AcApplication.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        _control.LoadSnapshot(DraftingPaletteSnapshot.Empty());
        _control.SetEnabled(false);
        _control.SetStatus("Open a drawing to start drafting.");
        return;
      }

      try
      {
        _control.LoadSnapshot(ReadSnapshot(document.Database));
        _control.SetEnabled(true);
      }
      catch (System.Exception ex)
      {
        _control.SetEnabled(false);
        _control.SetStatus("Unable to read drawing settings: " + ex.Message);
      }
    }

    internal static void SetStatus(string message)
    {
      _control?.SetStatus(message);
    }

    /// <summary>
    /// Starts a command in the active drawing and hands keyboard focus back to
    /// it so prompts and keywords can be answered immediately.
    /// </summary>
    internal static void RunCommand(string commandName, string status)
    {
      Document document = AcApplication.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        SetStatus("Open a drawing before running " + commandName + ".");
        return;
      }

      // Only cancel when something is running. An idle ESC would clear the
      // current selection, which RC reads as the receptacles to circuit.
      string cancelRunningCommand = IsCommandActive() ? "\u0003\u0003" : string.Empty;
      document.SendStringToExecute(
        cancelRunningCommand + "_." + commandName + " ",
        true,
        false,
        false
      );
      SetStatus(status);

      try
      {
        document.Window.Focus();
      }
      catch
      {
        // Focus is a convenience; the command is already queued.
      }
    }

    private static bool IsCommandActive()
    {
      try
      {
        return Convert.ToInt32(
          AcApplication.GetSystemVariable("CMDACTIVE"),
          CultureInfo.InvariantCulture
        ) != 0;
      }
      catch
      {
        return false;
      }
    }

    private static void EnsureCreated()
    {
      if (_palette != null)
      {
        return;
      }

      _control = new DraftingPaletteControl();
      // No explicit Size or Dock: AutoCAD restores the palette's saved
      // position, size, and docking from its GUID.
      _palette = new PaletteSet("Electrical Drafting", PaletteGuid)
      {
        Style = PaletteSetStyles.ShowAutoHideButton |
          PaletteSetStyles.ShowCloseButton |
          PaletteSetStyles.ShowPropertiesMenu,
        DockEnabled = DockSides.Left | DockSides.Right,
        MinimumSize = new Size(300, 300),
      };
      _palette.Add("Drafting", _control);

      if (!_documentEventsAttached)
      {
        AcApplication.DocumentManager.DocumentActivated += DocumentManager_DocumentActivated;
        _documentEventsAttached = true;
      }
    }

    private static void DocumentManager_DocumentActivated(
      object sender,
      DocumentCollectionEventArgs e
    )
    {
      if (_palette?.Visible == true)
      {
        Refresh();
      }
    }

    private static DraftingPaletteSnapshot ReadSnapshot(Database database)
    {
      var snapshot = new DraftingPaletteSnapshot();

      if (ElectricalDrawingSettingsStore.TryReadPanelLocation(database, out var location))
      {
        snapshot.LocationText = string.Format(
          CultureInfo.InvariantCulture,
          "{0:0.####}, {1:0.####} ({2})",
          location.Point.X,
          location.Point.Y,
          location.Context
        );
        snapshot.LocationDetail = string.Format(
          CultureInfo.InvariantCulture,
          "{0:0.####}, {1:0.####}, {2:0.####} ({3})",
          location.Point.X,
          location.Point.Y,
          location.Point.Z,
          location.Context
        );
      }

      if (ElectricalDrawingSettingsStore.TryReadScale(database, out var scale))
      {
        snapshot.ScaleText = scale.DisplayText;
      }

      if (ElectricalDrawingSettingsStore.TryReadPanelName(database, out string panelName))
      {
        snapshot.PanelName = panelName;
      }

      if (ElectricalDrawingSettingsStore.TryReadPanelSchedule(database, out var schedule)
        && !string.IsNullOrWhiteSpace(schedule.WorkbookPath))
      {
        snapshot.ScheduleText = Path.GetFileName(schedule.WorkbookPath);
        snapshot.SchedulePath = schedule.WorkbookPath;
      }

      if (ElectricalDrawingSettingsStore.TryReadReceptacleCircuitMaxKva(
        database,
        out double receptacleCircuitMaxKva))
      {
        snapshot.ReceptacleCircuitMaxKva = receptacleCircuitMaxKva;
      }

      GeneralCommands.ResolveHomerunLayerId(database, out string selectedLayerName);
      snapshot.SelectedLayerName = selectedLayerName;

      using (Transaction transaction = database.TransactionManager.StartOpenCloseTransaction())
      {
        LayerTable layers = transaction.GetObject(
          database.LayerTableId,
          OpenMode.ForRead
        ) as LayerTable;
        if (layers != null)
        {
          foreach (ObjectId id in layers)
          {
            if (id.IsErased)
            {
              continue;
            }

            LayerTableRecord layer = transaction.GetObject(
              id,
              OpenMode.ForRead
            ) as LayerTableRecord;
            if (layer != null)
            {
              snapshot.LayerNames.Add(layer.Name);
            }
          }
        }
      }

      snapshot.LayerNames = snapshot.LayerNames
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
        .ToList();
      return snapshot;
    }
  }

  internal sealed class DraftingPaletteControl : UserControl
  {
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessage(IntPtr hWnd, int message, IntPtr wParam, string lParam);

    private const int EmSetCueBanner = 0x1501;
    private const string SearchCue = "Search commands by name or purpose";
    private const int WrapColumns = 60;

    private static readonly string[] StandardScales =
    {
      "1\" = 1'-0\"",
      "3/4\" = 1'-0\"",
      "1/2\" = 1'-0\"",
      "3/8\" = 1'-0\"",
      "1/4\" = 1'-0\"",
      "3/16\" = 1'-0\"",
      "1/8\" = 1'-0\"",
      "3/32\" = 1'-0\"",
      "1/16\" = 1'-0\"",
    };

    private static readonly string[] ReceptacleCircuitMaximums =
    {
      "0.18",
      "0.36",
      "0.54",
      "0.72",
      "0.90",
      "1.08",
      "1.26",
      "1.44",
      "1.62",
      "1.80",
      "1.98",
      "2.16",
      "2.34",
    };

    private sealed class SectionView
    {
      internal string Title;
      internal bool Expanded;
      internal bool IsSettings;
      internal SectionHeader Header;
      internal Control Body;
      internal readonly List<RowView> Rows = new List<RowView>();
    }

    private sealed class RowView
    {
      internal Control[] Controls;
      internal string SearchText;
    }

    private readonly ToolTip _toolTip = new ToolTip
    {
      AutoPopDelay = 20000,
      InitialDelay = 400,
      ReshowDelay = 100,
      ShowAlways = true,
    };
    private readonly List<Control> _documentControls = new List<Control>();
    private readonly List<SectionView> _sections = new List<SectionView>();

    private readonly TextBox _searchBox;
    private readonly TableLayoutPanel _stack;
    private readonly Label _noMatchLabel;
    private readonly Panel _infoPane;
    private readonly TableLayoutPanel _infoLayout;
    private readonly Label _infoTitle;
    private readonly Label _infoNames;
    private readonly Label _infoBody;
    private readonly CheckBox _startupCheckBox;
    private readonly Label _statusLabel;
    private InfoIcon _infoOwner;

    private readonly Label _locationValue;
    private readonly ComboBox _scaleComboBox;
    private readonly TextBox _panelNameTextBox;
    private readonly Label _scheduleValue;
    private readonly ComboBox _receptacleCircuitMaxComboBox;
    private readonly ComboBox _layerComboBox;

    private readonly int _rowHeight;
    private readonly int _iconSize;
    private int _typedColumnWidth;

    internal DraftingPaletteControl()
    {
      AutoScaleMode = AutoScaleMode.Font;
      _rowHeight = Math.Max(Font.Height + 10, 26);
      _iconSize = Math.Max(Font.Height + 2, 16);

      var groups = new List<DraftingCommandGroup>(DraftingCommandCatalog.Groups);
      try
      {
        // Caught here, not inside the method: if AutoCAD's own assemblies
        // cannot load, the failure happens before the method body runs.
        DraftingCommandGroup unlisted = DraftingCommandCatalog.DiscoverUnlistedCommands();
        if (unlisted.Commands.Count > 0)
        {
          groups.Add(unlisted);
        }
      }
      catch
      {
        // The catalog stands on its own.
      }
      _typedColumnWidth = MeasureTypedColumn(groups);

      // The scrolling content, with the search box above it and the info and
      // status lines below it. Fill is added first so the edge-docked
      // controls claim their space before it.
      var scrollHost = new Panel
      {
        AutoScroll = true,
        Dock = DockStyle.Fill,
        Padding = new Padding(8, 4, 8, 0),
      };
      Controls.Add(scrollHost);

      _infoPane = new Panel
      {
        BackColor = SystemColors.Info,
        Dock = DockStyle.Bottom,
        ForeColor = SystemColors.InfoText,
        Padding = new Padding(10, 6, 10, 8),
        Visible = false,
      };
      _infoTitle = new Label
      {
        AutoSize = true,
        Dock = DockStyle.Fill,
        Font = new System.Drawing.Font(Font, FontStyle.Bold),
      };
      _infoNames = new Label
      {
        AutoSize = true,
        ForeColor = SystemColors.GrayText,
      };
      _infoBody = new Label { AutoSize = true };
      var closeInfo = new Label
      {
        AutoSize = true,
        Cursor = Cursors.Hand,
        ForeColor = SystemColors.GrayText,
        Text = "×",
      };
      closeInfo.Click += (_, __) => HideInfo();
      _toolTip.SetToolTip(closeInfo, "Hide this explanation");
      _infoLayout = new TableLayoutPanel
      {
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink,
        ColumnCount = 2,
        Dock = DockStyle.Top,
      };
      _infoLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      _infoLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      _infoLayout.Controls.Add(_infoTitle, 0, 0);
      _infoLayout.Controls.Add(closeInfo, 1, 0);
      _infoLayout.Controls.Add(_infoNames, 0, 1);
      _infoLayout.SetColumnSpan(_infoNames, 2);
      _infoLayout.Controls.Add(_infoBody, 0, 2);
      _infoLayout.SetColumnSpan(_infoBody, 2);
      _infoNames.Margin = new Padding(0, 2, 0, 4);
      _infoPane.Controls.Add(_infoLayout);
      _infoPane.SizeChanged += (_, __) => FitInfoLabels();
      Controls.Add(_infoPane);

      // A per-user preference, so it stays usable with no drawing open.
      _startupCheckBox = new CheckBox
      {
        AutoEllipsis = true,
        Checked = DraftingPaletteSettings.ReadShowAtStartup(),
        Dock = DockStyle.Fill,
        Text = "Open this palette when AutoCAD starts",
      };
      _startupCheckBox.CheckedChanged += StartupCheckBox_CheckedChanged;
      _toolTip.SetToolTip(
        _startupCheckBox,
        "When checked, this palette opens by itself each time AutoCAD starts.\nRun DRAFTPALETTE (EDP) to open it any time.");
      var optionsHost = new Panel
      {
        Dock = DockStyle.Bottom,
        Height = _startupCheckBox.PreferredSize.Height + 10,
        Padding = new Padding(10, 4, 10, 2),
      };
      optionsHost.Controls.Add(_startupCheckBox);
      Controls.Add(optionsHost);

      _statusLabel = new Label
      {
        AutoEllipsis = true,
        BorderStyle = BorderStyle.None,
        Dock = DockStyle.Bottom,
        ForeColor = SystemColors.GrayText,
        Height = 40,
        Padding = new Padding(10, 6, 10, 4),
        Text = "Ready.",
      };
      Controls.Add(_statusLabel);

      _searchBox = new TextBox { Dock = DockStyle.Fill, Name = "SearchBox" };
      _searchBox.HandleCreated += (_, __) =>
        SendMessage(_searchBox.Handle, EmSetCueBanner, (IntPtr)1, SearchCue);
      _searchBox.TextChanged += (_, __) => ApplyFilter();
      _searchBox.KeyDown += (_, e) =>
      {
        if (e.KeyCode == Keys.Escape)
        {
          _searchBox.Clear();
          e.SuppressKeyPress = true;
        }
      };
      _toolTip.SetToolTip(_searchBox, "Type part of a command name, shortcut, or description. Esc clears.");
      var searchHost = new Panel
      {
        Dock = DockStyle.Top,
        Height = _searchBox.PreferredHeight + 14,
        Padding = new Padding(8, 8, 8, 4),
      };
      searchHost.Controls.Add(_searchBox);
      Controls.Add(searchHost);

      _stack = new TableLayoutPanel
      {
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink,
        ColumnCount = 1,
        Dock = DockStyle.Top,
      };
      _stack.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      scrollHost.Controls.Add(_stack);

      _noMatchLabel = new Label
      {
        AutoSize = true,
        ForeColor = SystemColors.GrayText,
        Margin = new Padding(4, 8, 4, 8),
        Visible = false,
      };
      AddStackRow(_noMatchLabel);

      // --- Drawing settings ---------------------------------------------
      SectionView settings = AddSection("Drawing Settings", true, isSettings: true);
      TableLayoutPanel settingsGrid = CreateGrid(4);
      settingsGrid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      settingsGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      settingsGrid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
      settingsGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, _iconSize + 6));
      settings.Body = settingsGrid;
      AttachSection(settings);

      _panelNameTextBox = new TextBox { Dock = DockStyle.Top, Margin = new Padding(0, 3, 8, 8) };
      Button setPanelName = CreateSettingsButton("Set", SetPanelNameButton_Click);
      AddSettingsRow(
        settingsGrid, "Panel", _panelNameTextBox, setPanelName,
        "Panel name", "Command: SETPANELNAME   Shortcut: SPN",
        "Sets the panel name HR and RC use to label home runs and circuits (for example LP-1A). Needed before HR, RC, and linking a panel schedule.");

      _scaleComboBox = new ComboBox
      {
        Dock = DockStyle.Top,
        DropDownStyle = ComboBoxStyle.DropDown,
        Margin = new Padding(0, 3, 8, 8),
      };
      _scaleComboBox.Items.AddRange(StandardScales);
      Button setScale = CreateSettingsButton("Set", SetScaleButton_Click);
      AddSettingsRow(
        settingsGrid, "Scale", _scaleComboBox, setScale,
        "Drawing scale", "Command: SETSCALE   Shortcut: SS",
        "Sets the scale new symbols, text, and home runs are sized for. R and RC also pick the scale up from the active viewport when you run them inside one.");

      _locationValue = CreateValueLabel();
      Button pickLocation = CreateSettingsButton("Pick", PickLocationButton_Click);
      AddSettingsRow(
        settingsGrid, "Location", _locationValue, pickLocation,
        "Panel location", "Command: SETPANELLOCATION   Shortcut: SPL",
        "Marks where the panel is so HR can point its arrow toward it. Pick it in the same space (model space or the viewport) where you draw home runs.");

      _scheduleValue = CreateValueLabel();
      Button linkSchedule = CreateSettingsButton("Link", LinkScheduleButton_Click);
      AddSettingsRow(
        settingsGrid, "Schedule", _scheduleValue, linkSchedule,
        "Panel schedule", "Command: SETPANELSCHEDULE   Shortcut: SPS",
        "Links the panel's Excel schedule workbook. RC writes receptacle circuits into it, so close the workbook in Excel before running RC.");

      _receptacleCircuitMaxComboBox = new ComboBox
      {
        Dock = DockStyle.Top,
        DropDownStyle = ComboBoxStyle.DropDown,
        Margin = new Padding(0, 3, 8, 8),
      };
      _receptacleCircuitMaxComboBox.Items.AddRange(ReceptacleCircuitMaximums);
      Button setMax = CreateSettingsButton("Set", SetReceptacleCircuitMaxButton_Click);
      AddSettingsRow(
        settingsGrid, "RC Max kVA", _receptacleCircuitMaxComboBox, setMax,
        "RC maximum circuit load", "Setting used by RC",
        "The most load RC puts on one receptacle circuit. A duplex counts 0.18 kVA and a quad 0.36 kVA, so the default 1.26 kVA is seven duplex receptacles.");

      _layerComboBox = new ComboBox
      {
        Dock = DockStyle.Top,
        DropDownStyle = ComboBoxStyle.DropDownList,
        Margin = new Padding(0, 3, 8, 8),
        Sorted = true,
      };
      Button setLayer = CreateSettingsButton("Set", SetLayerButton_Click);
      AddSettingsRow(
        settingsGrid, "HR Layer", _layerComboBox, setLayer,
        "Home-run layer", "Setting used by HR",
        "The layer new home runs are drawn on. If it is not set, HR uses the current layer.");

      var refresh = new Button
      {
        AutoSize = true,
        Margin = new Padding(0, 0, 0, 4),
        Text = "Refresh",
      };
      refresh.Click += (_, __) =>
      {
        DraftingPalette.Refresh();
        SetStatus("Settings refreshed from the current drawing.");
      };
      _toolTip.SetToolTip(refresh, "Reload the settings from the current drawing.");
      settingsGrid.Controls.Add(refresh, 1, settingsGrid.RowCount);
      settingsGrid.RowCount++;
      settingsGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));

      // --- Command groups -----------------------------------------------
      foreach (DraftingCommandGroup group in groups)
      {
        BuildCommandSection(group);
      }

      ApplyFilter();
    }

    internal void LoadSnapshot(DraftingPaletteSnapshot snapshot)
    {
      snapshot = snapshot ?? DraftingPaletteSnapshot.Empty();
      SetValueLabel(_locationValue, snapshot.LocationText, snapshot.LocationDetail);
      SetValueLabel(_scheduleValue, snapshot.ScheduleText, snapshot.SchedulePath);
      _scaleComboBox.Text = snapshot.ScaleText ?? string.Empty;
      _panelNameTextBox.Text = snapshot.PanelName ?? string.Empty;
      _receptacleCircuitMaxComboBox.Text =
        snapshot.ReceptacleCircuitMaxKva.ToString(
          "0.00",
          CultureInfo.InvariantCulture);

      _layerComboBox.Items.Clear();
      foreach (string layerName in snapshot.LayerNames)
      {
        _layerComboBox.Items.Add(layerName);
      }

      int selectedIndex = FindComboItem(_layerComboBox, snapshot.SelectedLayerName);
      if (selectedIndex < 0 && _layerComboBox.Items.Count > 0)
      {
        selectedIndex = 0;
      }
      _layerComboBox.SelectedIndex = selectedIndex;
    }

    internal void SetEnabled(bool enabled)
    {
      foreach (Control control in _documentControls)
      {
        control.Enabled = enabled;
      }
    }

    internal void SetStatus(string message)
    {
      _statusLabel.Text = string.IsNullOrWhiteSpace(message) ? "Ready." : message;
    }

    protected override void Dispose(bool disposing)
    {
      if (disposing)
      {
        _toolTip.Dispose();
      }
      base.Dispose(disposing);
    }

    // ----------------------------------------------------------------------
    // Layout helpers
    // ----------------------------------------------------------------------

    private int MeasureTypedColumn(IEnumerable<DraftingCommandGroup> groups)
    {
      int widest = 0;
      foreach (DraftingCommandGroup group in groups)
      {
        foreach (DraftingCommand command in group.Commands.Where(c => !c.HasVariants))
        {
          widest = Math.Max(widest, TextRenderer.MeasureText(command.TypedName, Font).Width);
        }
      }
      return Math.Min(widest + 8, 140);
    }

    private void AddStackRow(Control control)
    {
      control.Dock = control.Dock == DockStyle.None ? DockStyle.Fill : control.Dock;
      _stack.Controls.Add(control, 0, _stack.RowCount);
      _stack.RowCount++;
      _stack.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    }

    private static TableLayoutPanel CreateGrid(int columns)
    {
      return new TableLayoutPanel
      {
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink,
        ColumnCount = columns,
        Dock = DockStyle.Fill,
        Margin = new Padding(0, 4, 0, 8),
        Padding = new Padding(4, 0, 0, 0),
      };
    }

    private SectionView AddSection(string title, bool expanded, bool isSettings = false)
    {
      var section = new SectionView
      {
        Title = title,
        Expanded = expanded,
        IsSettings = isSettings,
      };
      section.Header = new SectionHeader(Font)
      {
        Dock = DockStyle.Fill,
        Margin = new Padding(0, 2, 0, 0),
        Text = title,
      };
      section.Header.Height = Font.Height + 10;
      section.Header.Click += (_, __) =>
      {
        if (HasFilter())
        {
          return;
        }
        section.Expanded = !section.Expanded;
        ApplyFilter();
      };
      _sections.Add(section);
      return section;
    }

    private void AttachSection(SectionView section)
    {
      AddStackRow(section.Header);
      AddStackRow(section.Body);
    }

    private void BuildCommandSection(DraftingCommandGroup group)
    {
      SectionView section = AddSection(group.Title, group.ExpandedByDefault);
      TableLayoutPanel grid = CreateGrid(3);
      grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
      grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, _typedColumnWidth));
      grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, _iconSize + 6));
      section.Body = grid;
      AttachSection(section);

      foreach (DraftingCommand command in group.Commands)
      {
        section.Rows.Add(command.HasVariants
          ? AddVariantRow(grid, command)
          : AddCommandRow(grid, command));
      }
    }

    private RowView AddCommandRow(
      TableLayoutPanel grid,
      DraftingCommand command)
    {
      string typed = command.TypedName;
      var button = new Button
      {
        AutoEllipsis = true,
        Dock = DockStyle.Fill,
        Height = _rowHeight,
        Margin = new Padding(0, 1, 4, 1),
        Text = command.Title,
        TextAlign = ContentAlignment.MiddleLeft,
        UseVisualStyleBackColor = true,
      };
      string status = "Running " + typed + ". Follow the prompts at the command line.";
      button.Click += (_, __) => DraftingPalette.RunCommand(command.Command, status);
      _toolTip.SetToolTip(button, "Run " + command.Command);
      _documentControls.Add(button);

      var typedLabel = new Label
      {
        AutoEllipsis = true,
        Dock = DockStyle.Fill,
        ForeColor = SystemColors.GrayText,
        Margin = new Padding(0),
        Text = typed,
        TextAlign = ContentAlignment.MiddleLeft,
      };
      InfoIcon icon = CreateInfoIcon(command.Title, DescribeNames(command), command.Description);

      int row = grid.RowCount;
      grid.RowCount++;
      grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      grid.Controls.Add(button, 0, row);
      grid.Controls.Add(typedLabel, 1, row);
      grid.Controls.Add(icon, 2, row);

      return new RowView
      {
        Controls = new Control[] { button, typedLabel, icon },
        SearchText = BuildSearchText(command),
      };
    }

    private RowView AddVariantRow(
      TableLayoutPanel grid,
      DraftingCommand command)
    {
      int count = command.VariantCommands.Length;
      var strip = new TableLayoutPanel
      {
        AutoSize = true,
        AutoSizeMode = AutoSizeMode.GrowAndShrink,
        ColumnCount = count + 1,
        Dock = DockStyle.Fill,
        Margin = new Padding(0, 1, 4, 1),
      };
      strip.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

      var title = new Label
      {
        AutoEllipsis = true,
        Dock = DockStyle.Fill,
        Height = _rowHeight,
        Margin = new Padding(0),
        Text = command.Title,
        TextAlign = ContentAlignment.MiddleLeft,
      };
      strip.Controls.Add(title, 0, 0);

      int buttonWidth = 30;
      foreach (string label in command.VariantLabels)
      {
        buttonWidth = Math.Max(buttonWidth, TextRenderer.MeasureText(label, Font).Width + 16);
      }
      for (int i = 0; i < count; i++)
      {
        string variantCommand = command.VariantCommands[i];
        var button = new Button
        {
          Height = _rowHeight,
          Margin = new Padding(2, 0, 0, 0),
          Text = command.VariantLabels[i],
          UseVisualStyleBackColor = true,
          Width = buttonWidth,
        };
        string status = "Running " + variantCommand + ". Follow the prompts at the command line.";
        button.Click += (_, __) => DraftingPalette.RunCommand(variantCommand, status);
        _toolTip.SetToolTip(button, "Run " + variantCommand);
        _documentControls.Add(button);
        strip.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, buttonWidth + 2));
        strip.Controls.Add(button, i + 1, 0);
      }

      InfoIcon icon = CreateInfoIcon(command.Title, DescribeNames(command), command.Description);

      int row = grid.RowCount;
      grid.RowCount++;
      grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      grid.Controls.Add(strip, 0, row);
      grid.SetColumnSpan(strip, 2);
      grid.Controls.Add(icon, 2, row);

      return new RowView
      {
        Controls = new Control[] { strip, icon },
        SearchText = BuildSearchText(command),
      };
    }

    private static string BuildSearchText(DraftingCommand command)
    {
      return string.Join(
        " ",
        new[] { command.Title, command.Description }
          .Concat(command.AllNames))
        .ToLowerInvariant();
    }

    private static string DescribeNames(DraftingCommand command)
    {
      if (command.HasVariants)
      {
        string names = "Commands: " + string.Join(", ", command.VariantCommands);
        return command.Aliases.Length == 0
          ? names
          : names + "   Also: " + string.Join(", ", command.Aliases);
      }

      string shortest = command.TypedName;
      string[] others = new[] { command.Command }
        .Concat(command.Aliases)
        .Where(name => !string.Equals(name, shortest, StringComparison.OrdinalIgnoreCase))
        .ToArray();
      return others.Length == 0
        ? "Command: " + shortest
        : "Command: " + shortest + "   Also: " + string.Join(", ", others);
    }

    // ----------------------------------------------------------------------
    // Searching and collapsing
    // ----------------------------------------------------------------------

    private bool HasFilter()
    {
      return !string.IsNullOrWhiteSpace(_searchBox.Text);
    }

    private void ApplyFilter()
    {
      string[] tokens = (_searchBox.Text ?? string.Empty)
        .ToLowerInvariant()
        .Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
      bool filtering = tokens.Length > 0;
      int totalMatches = 0;

      _stack.SuspendLayout();
      foreach (SectionView section in _sections)
      {
        if (section.IsSettings)
        {
          section.Header.Visible = !filtering;
          section.Body.Visible = !filtering && section.Expanded;
          section.Header.Expanded = section.Expanded;
          section.Header.CountText = string.Empty;
          continue;
        }

        int matches = 0;
        foreach (RowView row in section.Rows)
        {
          bool matched = !filtering || tokens.All(token => row.SearchText.Contains(token));
          foreach (Control control in row.Controls)
          {
            control.Visible = matched;
          }
          if (matched)
          {
            matches++;
          }
        }

        totalMatches += matches;
        bool showSection = !filtering || matches > 0;
        section.Header.Visible = showSection;
        section.Body.Visible = showSection && (filtering || section.Expanded);
        section.Header.Expanded = filtering || section.Expanded;
        section.Header.CountText = filtering
          ? matches + " of " + section.Rows.Count
          : section.Rows.Count.ToString(CultureInfo.InvariantCulture);
      }

      _noMatchLabel.Text = "No commands match \"" + _searchBox.Text.Trim() + "\".";
      _noMatchLabel.Visible = filtering && totalMatches == 0;
      _stack.ResumeLayout(true);

      if (_infoOwner != null && !_infoOwner.Visible)
      {
        HideInfo();
      }
    }

    // ----------------------------------------------------------------------
    // Info icon and explanation pane
    // ----------------------------------------------------------------------

    private InfoIcon CreateInfoIcon(string title, string names, string description)
    {
      var icon = new InfoIcon(Font)
      {
        Anchor = AnchorStyles.None,
        Margin = new Padding(3, 0, 0, 0),
        Size = new Size(_iconSize, _iconSize),
      };
      icon.AccessibleName = "About " + title;
      icon.Click += (_, __) => ToggleInfo(icon, title, names, description);
      _toolTip.SetToolTip(icon, title + "\n" + names + "\n" + Wrap(description, WrapColumns));
      return icon;
    }

    private void ToggleInfo(InfoIcon icon, string title, string names, string description)
    {
      if (_infoOwner == icon)
      {
        HideInfo();
        return;
      }

      if (_infoOwner != null)
      {
        _infoOwner.Active = false;
      }
      _infoOwner = icon;
      icon.Active = true;
      _infoTitle.Text = title;
      _infoNames.Text = names;
      _infoBody.Text = description;
      _infoPane.Visible = true;
      FitInfoLabels();
    }

    private void HideInfo()
    {
      if (_infoOwner != null)
      {
        _infoOwner.Active = false;
        _infoOwner = null;
      }
      _infoPane.Visible = false;
    }

    // Labels only wrap when given a maximum width, and a docked panel does not
    // re-measure itself when that changes, so size the pane to its content here.
    private void FitInfoLabels()
    {
      int width = Math.Max(
        60,
        _infoPane.ClientSize.Width - _infoPane.Padding.Horizontal - 4);
      _infoNames.MaximumSize = new Size(width, 0);
      _infoBody.MaximumSize = new Size(width, 0);
      _infoLayout.PerformLayout();

      int contentHeight = _infoLayout
        .GetPreferredSize(new Size(width + 4, 0))
        .Height;
      int height = contentHeight + _infoPane.Padding.Vertical;
      if (_infoPane.Height != height)
      {
        _infoPane.Height = height;
      }
    }

    private static string Wrap(string text, int columns)
    {
      var wrapped = new StringBuilder();
      int lineLength = 0;
      foreach (string word in (text ?? string.Empty).Split(' '))
      {
        if (lineLength > 0 && lineLength + 1 + word.Length > columns)
        {
          wrapped.Append('\n');
          lineLength = 0;
        }
        else if (lineLength > 0)
        {
          wrapped.Append(' ');
          lineLength++;
        }
        wrapped.Append(word);
        lineLength += word.Length;
      }
      return wrapped.ToString();
    }

    // ----------------------------------------------------------------------
    // Drawing settings rows
    // ----------------------------------------------------------------------

    private Button CreateSettingsButton(string text, EventHandler clickHandler)
    {
      var button = new Button
      {
        AutoSize = true,
        Margin = new Padding(0, 1, 0, 8),
        MinimumSize = new Size(52, 0),
        Text = text,
      };
      button.Click += clickHandler;
      _documentControls.Add(button);
      return button;
    }

    private static Label CreateValueLabel()
    {
      return new Label
      {
        AutoEllipsis = true,
        Dock = DockStyle.Fill,
        Margin = new Padding(0, 6, 8, 8),
        TextAlign = ContentAlignment.MiddleLeft,
      };
    }

    private void SetValueLabel(Label label, string text, string toolTip = null)
    {
      bool isSet = !string.IsNullOrWhiteSpace(text);
      label.Text = isSet ? text : "Not set";
      label.ForeColor = isSet ? SystemColors.ControlText : Color.Firebrick;
      _toolTip.SetToolTip(label, isSet ? (string.IsNullOrEmpty(toolTip) ? text : toolTip) : string.Empty);
    }

    private void AddSettingsRow(
      TableLayoutPanel layout,
      string labelText,
      Control input,
      Control button,
      string infoTitle,
      string infoNames,
      string infoDescription
    )
    {
      int row = layout.RowCount;
      var label = new Label
      {
        AutoSize = true,
        Margin = new Padding(0, 7, 10, 8),
        Text = labelText,
      };
      InfoIcon icon = CreateInfoIcon(infoTitle, infoNames, infoDescription);
      icon.Margin = new Padding(3, 0, 0, 6);
      layout.Controls.Add(label, 0, row);
      layout.Controls.Add(input, 1, row);
      layout.Controls.Add(button, 2, row);
      layout.Controls.Add(icon, 3, row);
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
      layout.RowCount = row + 1;
      _documentControls.Add(input);
    }

    private static int FindComboItem(ComboBox comboBox, string value)
    {
      for (int i = 0; i < comboBox.Items.Count; i++)
      {
        if (string.Equals(
          Convert.ToString(comboBox.Items[i]),
          value,
          StringComparison.OrdinalIgnoreCase
        ))
        {
          return i;
        }
      }
      return -1;
    }

    private void StartupCheckBox_CheckedChanged(object sender, EventArgs e)
    {
      bool wanted = _startupCheckBox.Checked;
      if (!DraftingPaletteSettings.TryWriteShowAtStartup(wanted, out string error))
      {
        // Put the box back so it shows what is actually saved.
        _startupCheckBox.CheckedChanged -= StartupCheckBox_CheckedChanged;
        _startupCheckBox.Checked = !wanted;
        _startupCheckBox.CheckedChanged += StartupCheckBox_CheckedChanged;
        SetStatus("Unable to save the startup setting: " + error);
        return;
      }

      SetStatus(wanted
        ? "The palette will open when AutoCAD starts."
        : "The palette will not open when AutoCAD starts. Run EDP to open it.");
    }

    private void PickLocationButton_Click(object sender, EventArgs e)
    {
      DraftingPalette.RunCommand("SPL", "SPL: pick the panel location in the drawing.");
    }

    private void LinkScheduleButton_Click(object sender, EventArgs e)
    {
      DraftingPalette.RunCommand("SPS", "SPS: choose the panel schedule workbook.");
    }

    private void SetScaleButton_Click(object sender, EventArgs e)
    {
      Document document = AcApplication.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        SetStatus("Open a drawing before setting the scale.");
        return;
      }

      string candidate = (_scaleComboBox.Text ?? string.Empty).Trim();
      if (!GeneralCommands.TryParseDrawingScale(
        candidate,
        out double paperInchesPerFoot,
        out string displayText
      ))
      {
        SetStatus("Enter a scale such as 1/4, 3/16, or 1:48.");
        return;
      }

      try
      {
        using (DocumentLock documentLock = document.LockDocument())
        {
          ElectricalDrawingSettingsStore.WriteScale(
            document.Database,
            paperInchesPerFoot,
            displayText
          );
        }
        DraftingPalette.Refresh();
        SetStatus("Scale set to " + displayText + ".");
        document.Editor.WriteMessage("\nHR scale set to " + displayText + ".");
      }
      catch (System.Exception ex)
      {
        SetStatus("Unable to set the scale: " + ex.Message);
      }
    }

    private void SetPanelNameButton_Click(object sender, EventArgs e)
    {
      Document document = AcApplication.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        SetStatus("Open a drawing before setting the panel name.");
        return;
      }

      string panelName = (_panelNameTextBox.Text ?? string.Empty).Trim();
      if (panelName.Length == 0)
      {
        SetStatus("Panel name cannot be blank.");
        return;
      }

      try
      {
        using (DocumentLock documentLock = document.LockDocument())
        {
          ElectricalDrawingSettingsStore.WritePanelName(document.Database, panelName);
        }
        DraftingPalette.Refresh();
        SetStatus("Panel name set to " + panelName + ".");
        document.Editor.WriteMessage("\nHR panel name set to " + panelName + ".");
      }
      catch (System.Exception ex)
      {
        SetStatus("Unable to set the panel name: " + ex.Message);
      }
    }

    private void SetReceptacleCircuitMaxButton_Click(
      object sender,
      EventArgs e)
    {
      Document document = AcApplication.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        SetStatus("Open a drawing before setting the RC maximum load.");
        return;
      }

      string candidate =
        (_receptacleCircuitMaxComboBox.Text ?? string.Empty).Trim();
      if (!double.TryParse(
            candidate,
            NumberStyles.Float,
            CultureInfo.InvariantCulture,
            out double maximumKva) &&
          !double.TryParse(
            candidate,
            NumberStyles.Float,
            CultureInfo.CurrentCulture,
            out maximumKva))
      {
        SetStatus("Enter an RC maximum load such as 0.90 or 1.26 kVA.");
        return;
      }

      try
      {
        using (DocumentLock documentLock = document.LockDocument())
        {
          ElectricalDrawingSettingsStore.WriteReceptacleCircuitMaxKva(
            document.Database,
            maximumKva);
        }
        DraftingPalette.Refresh();
        SetStatus(
          $"RC maximum circuit load set to {maximumKva:0.00} kVA.");
        document.Editor.WriteMessage(
          $"\nRC maximum circuit load set to {maximumKva:0.00} kVA.");
      }
      catch (System.Exception ex)
      {
        SetStatus("Unable to set the RC maximum load: " + ex.Message);
      }
    }

    private void SetLayerButton_Click(object sender, EventArgs e)
    {
      Document document = AcApplication.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        SetStatus("Open a drawing before setting the HR layer.");
        return;
      }

      string layerName = Convert.ToString(_layerComboBox.SelectedItem) ?? string.Empty;
      if (layerName.Length == 0)
      {
        SetStatus("Select an existing layer.");
        return;
      }

      try
      {
        using (DocumentLock documentLock = document.LockDocument())
        {
          using (Transaction transaction = document.Database.TransactionManager.StartOpenCloseTransaction())
          {
            LayerTable layers = transaction.GetObject(
              document.Database.LayerTableId,
              OpenMode.ForRead
            ) as LayerTable;
            if (layers == null || !layers.Has(layerName))
            {
              throw new InvalidOperationException("The selected layer no longer exists.");
            }
          }

          ElectricalDrawingSettingsStore.WriteHomerunLayer(document.Database, layerName);
        }
        DraftingPalette.Refresh();
        SetStatus("HR layer set to " + layerName + ".");
        document.Editor.WriteMessage("\nHR layer set to " + layerName + ".");
      }
      catch (System.Exception ex)
      {
        SetStatus("Unable to set the HR layer: " + ex.Message);
      }
    }

    // ----------------------------------------------------------------------
    // Custom controls
    // ----------------------------------------------------------------------

    /// <summary>A collapsible group title: triangle, name, and a count.</summary>
    private sealed class SectionHeader : Control
    {
      private readonly System.Drawing.Font _boldFont;
      private bool _hot;
      private bool _expanded;
      private string _countText = string.Empty;

      internal SectionHeader(System.Drawing.Font baseFont)
      {
        SetStyle(
          ControlStyles.UserPaint |
          ControlStyles.AllPaintingInWmPaint |
          ControlStyles.OptimizedDoubleBuffer |
          ControlStyles.ResizeRedraw,
          true);
        Cursor = Cursors.Hand;
        TabStop = false;
        _boldFont = new System.Drawing.Font(baseFont, FontStyle.Bold);
      }

      internal bool Expanded
      {
        get { return _expanded; }
        set
        {
          _expanded = value;
          Invalidate();
        }
      }

      internal string CountText
      {
        get { return _countText; }
        set
        {
          _countText = value ?? string.Empty;
          Invalidate();
        }
      }

      protected override void OnTextChanged(EventArgs e)
      {
        base.OnTextChanged(e);
        Invalidate();
      }

      protected override void OnMouseEnter(EventArgs e)
      {
        base.OnMouseEnter(e);
        _hot = true;
        Invalidate();
      }

      protected override void OnMouseLeave(EventArgs e)
      {
        base.OnMouseLeave(e);
        _hot = false;
        Invalidate();
      }

      protected override void OnPaint(PaintEventArgs e)
      {
        Graphics g = e.Graphics;
        g.Clear(_hot ? SystemColors.ControlLight : SystemColors.Control);
        using (var pen = new Pen(SystemColors.ControlDark))
        {
          g.DrawLine(pen, 0, Height - 1, Width, Height - 1);
        }

        int centerY = Height / 2;
        Point[] triangle = _expanded
          ? new[]
          {
            new Point(6, centerY - 2),
            new Point(14, centerY - 2),
            new Point(10, centerY + 3),
          }
          : new[]
          {
            new Point(8, centerY - 4),
            new Point(8, centerY + 4),
            new Point(13, centerY),
          };
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using (var brush = new SolidBrush(SystemColors.ControlText))
        {
          g.FillPolygon(brush, triangle);
        }

        const TextFormatFlags flags =
          TextFormatFlags.VerticalCenter |
          TextFormatFlags.SingleLine |
          TextFormatFlags.NoPrefix |
          TextFormatFlags.EndEllipsis;
        int countWidth = _countText.Length == 0
          ? 0
          : TextRenderer.MeasureText(_countText, Font).Width + 10;
        TextRenderer.DrawText(
          g,
          Text,
          _boldFont,
          new Rectangle(20, 0, Math.Max(0, Width - 20 - countWidth), Height),
          SystemColors.ControlText,
          flags);
        if (countWidth > 0)
        {
          TextRenderer.DrawText(
            g,
            _countText,
            Font,
            new Rectangle(Width - countWidth, 0, countWidth - 6, Height),
            SystemColors.GrayText,
            flags | TextFormatFlags.Right);
        }
      }

      protected override void Dispose(bool disposing)
      {
        if (disposing)
        {
          _boldFont.Dispose();
        }
        base.Dispose(disposing);
      }
    }

    /// <summary>The round "i" that explains a command.</summary>
    private sealed class InfoIcon : Control
    {
      private readonly System.Drawing.Font _letterFont;
      private bool _hot;
      private bool _active;

      internal InfoIcon(System.Drawing.Font baseFont)
      {
        SetStyle(
          ControlStyles.UserPaint |
          ControlStyles.AllPaintingInWmPaint |
          ControlStyles.OptimizedDoubleBuffer |
          ControlStyles.ResizeRedraw |
          ControlStyles.SupportsTransparentBackColor,
          true);
        BackColor = Color.Transparent;
        Cursor = Cursors.Hand;
        TabStop = false;
        _letterFont = new System.Drawing.Font(baseFont.FontFamily, baseFont.Size - 1f, FontStyle.Bold);
      }

      internal bool Active
      {
        get { return _active; }
        set
        {
          _active = value;
          Invalidate();
        }
      }

      protected override void OnMouseEnter(EventArgs e)
      {
        base.OnMouseEnter(e);
        _hot = true;
        Invalidate();
      }

      protected override void OnMouseLeave(EventArgs e)
      {
        base.OnMouseLeave(e);
        _hot = false;
        Invalidate();
      }

      protected override void OnPaint(PaintEventArgs e)
      {
        base.OnPaint(e);
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        int diameter = Math.Min(Width, Height) - 3;
        var circle = new Rectangle((Width - diameter) / 2, (Height - diameter) / 2, diameter, diameter);
        Color accent = _active || _hot ? SystemColors.HotTrack : SystemColors.GrayText;
        if (_active)
        {
          using (var brush = new SolidBrush(accent))
          {
            g.FillEllipse(brush, circle);
          }
        }
        using (var pen = new Pen(accent, 1.4f))
        {
          g.DrawEllipse(pen, circle);
        }
        TextRenderer.DrawText(
          g,
          "i",
          _letterFont,
          circle,
          _active ? SystemColors.Window : accent,
          TextFormatFlags.HorizontalCenter |
          TextFormatFlags.VerticalCenter |
          TextFormatFlags.NoPadding |
          TextFormatFlags.SingleLine);
      }

      protected override void Dispose(bool disposing)
      {
        if (disposing)
        {
          _letterFont.Dispose();
        }
        base.Dispose(disposing);
      }
    }
  }

  internal sealed class DraftingPaletteSnapshot
  {
    internal string LocationText { get; set; } = string.Empty;
    internal string LocationDetail { get; set; } = string.Empty;
    internal string ScaleText { get; set; } = string.Empty;
    internal string PanelName { get; set; } = string.Empty;
    internal string ScheduleText { get; set; } = string.Empty;
    internal string SchedulePath { get; set; } = string.Empty;
    internal double ReceptacleCircuitMaxKva { get; set; } =
      GeneralCommands.DefaultReceptacleCircuitMaxKva;
    internal string SelectedLayerName { get; set; } = string.Empty;
    internal List<string> LayerNames { get; set; } = new List<string>();

    internal static DraftingPaletteSnapshot Empty()
    {
      return new DraftingPaletteSnapshot();
    }
  }
}
