using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private const double NoteTablePaperTextHeight = 0.09375;
    private const double NoteTablePaperWidth = 4.0;

    [CommandMethod("NOTETABLE", CommandFlags.Modal | CommandFlags.UsePickSet)]
    public static void OpenNoteTable()
    {
      RunNoteTableEditor("NOTETABLE", null);
    }

    [CommandMethod("KEYNOTETABLE", CommandFlags.Modal | CommandFlags.UsePickSet)]
    [CommandMethod("KNTABLE", CommandFlags.Modal | CommandFlags.UsePickSet)]
    [CommandMethod("KNT", CommandFlags.Modal | CommandFlags.UsePickSet)]
    public static void OpenKeyNoteTable()
    {
      RunNoteTableEditor("KEYNOTETABLE", NoteTableKind.Keyed);
    }

    [CommandMethod("GENERALNOTETABLE", CommandFlags.Modal | CommandFlags.UsePickSet)]
    [CommandMethod("GNTABLE", CommandFlags.Modal | CommandFlags.UsePickSet)]
    [CommandMethod("GNT", CommandFlags.Modal | CommandFlags.UsePickSet)]
    public static void OpenGeneralNoteTable()
    {
      RunNoteTableEditor("GENERALNOTETABLE", NoteTableKind.General);
    }

    /// <summary>
    /// Shows the note editor until the user saves, creates a table or cancels.
    /// Imports and table placement need the drawing, so the window closes for
    /// them and opens again afterwards with the same session.
    /// </summary>
    private static void RunNoteTableEditor(string commandName, NoteTableKind? preferredKind)
    {
      Document document = Application.DocumentManager.MdiActiveDocument;
      if (document == null)
      {
        return;
      }
      Database database = document.Database;
      Editor editor = document.Editor;

      ObjectId[] preselectedIds = Array.Empty<ObjectId>();
      PromptSelectionResult implied = editor.SelectImplied();
      if (implied.Status == PromptStatus.OK && implied.Value != null)
      {
        preselectedIds = implied.Value.GetObjectIds();
        editor.SetImpliedSelection(Array.Empty<ObjectId>());
      }

      NoteTableSession session;
      try
      {
        session = LoadNoteTableSession(document, database, editor, preferredKind);
      }
      catch (System.Exception ex)
      {
        editor.WriteMessage($"\n{commandName}: {ex.Message}");
        return;
      }
      if (string.IsNullOrEmpty(session.ProjectFilePath))
      {
        editor.WriteMessage(
          $"\n{commandName}: this drawing has never been saved, so the note tabs " +
          "can't be kept. Save it into the project folder to share them.");
      }

      if (preselectedIds.Length > 0)
      {
        ImportPreselectedNoteObjects(
          database,
          editor,
          commandName,
          session,
          GetSelectedNoteTab(session),
          preselectedIds);
      }

      while (true)
      {
        var window = new NoteTableWindow(session);
        Application.ShowModalWindow(window);
        NoteTableTabItem tab = GetSelectedNoteTab(session);
        switch (window.Action)
        {
          case NoteTableWindowAction.ImportText:
            PromptImportNoteText(database, editor, commandName, session, tab);
            continue;

          case NoteTableWindowAction.ImportTable:
            PromptImportNoteTable(database, editor, session, tab);
            continue;

          case NoteTableWindowAction.Save:
            if (TrySaveNoteTableProject(session))
            {
              editor.WriteMessage(
                $"\nSaved {session.Tabs.Count} note tab" +
                $"{(session.Tabs.Count == 1 ? string.Empty : "s")} to {session.ProjectFilePath}.");
              return;
            }
            continue;

          case NoteTableWindowAction.UpdateTable:
            if (TrySaveNoteTableProject(session) &&
                TryUpdateNoteTable(database, editor, session, tab))
            {
              return;
            }
            continue;

          case NoteTableWindowAction.Create:
            if (TrySaveNoteTableProject(session) &&
                TryCreateNoteTable(database, editor, session, tab))
            {
              return;
            }
            continue;

          default:
            editor.WriteMessage($"\n{commandName} canceled. Nothing was saved.");
            return;
        }
      }
    }

    private static NoteTableSession LoadNoteTableSession(
      Document document,
      Database database,
      Editor editor,
      NoteTableKind? preferredKind)
    {
      // Document.Name is the full path once the drawing has been saved;
      // Database.Filename would be the template for a new drawing.
      string drawingPath = document.Name;
      string projectPath = NoteTableStore.GetProjectFilePath(drawingPath);
      NoteTableProject project = NoteTableStore.Load(projectPath);
      string planTabName = NoteTableStore.DetectPlanTabName(drawingPath);
      NoteTableStore.AddDefaultTabs(project, planTabName);
      if (preferredKind.HasValue &&
          !project.Tabs.Any(tab => tab.Kind == preferredKind.Value))
      {
        project.Tabs.Add(new NoteTableTabData
        {
          Name = preferredKind.Value == NoteTableKind.General
            ? NoteTableStore.GeneralNotesTabName
            : NoteTableStore.KeyNotesTabName,
          Kind = preferredKind.Value,
        });
      }

      NoteTableTabData planTab = planTabName == null
        ? null
        : NoteTableStore.FindTab(project, planTabName);
      NoteTableTabData selected =
        (planTab != null && (!preferredKind.HasValue || planTab.Kind == preferredKind.Value)
          ? planTab
          : null) ??
        project.Tabs.FirstOrDefault(tab =>
          !preferredKind.HasValue || tab.Kind == preferredKind.Value) ??
        project.Tabs[0];

      double scale = ResolveNoteTableScale(editor, database, out string scaleHint);
      var session = new NoteTableSession
      {
        ProjectFilePath = projectPath,
        LoadedFileTimeUtc = NoteTableStore.GetLastWriteTimeUtc(projectPath),
        SelectedTabIndex = project.Tabs.IndexOf(selected),
        TextHeight = NoteTablePaperTextHeight * scale,
        TableWidth = NoteTablePaperWidth * scale,
        ScaleHint = scaleHint,
      };
      foreach (NoteTableTabData tab in project.Tabs)
      {
        session.Tabs.Add(new NoteTableTabItem(tab.Name, tab.Kind, tab.FirstNumber, tab.Notes));
      }
      return session;
    }

    /// <summary>
    /// Drawing units per plotted inch where the table will go: 1 in paper
    /// space, the viewport scale inside a viewport, the SETSCALE scale in
    /// model space.
    /// </summary>
    private static double ResolveNoteTableScale(
      Editor editor,
      Database database,
      out string hint)
    {
      if (!database.TileMode)
      {
        if (Convert.ToInt16(Application.GetSystemVariable("CVPORT")) > 1)
        {
          double denominator = ResolveViewportScaleDenominator(editor, database);
          if (denominator > 0.0)
          {
            hint = $"Sized for the active viewport ({FormatRatio(denominator)}).";
            return denominator;
          }
        }
        hint = "Sized for paper space (3/32\" text).";
        return 1.0;
      }

      if (ElectricalDrawingSettingsStore.TryReadScale(
        database,
        out ElectricalDrawingSettingsStore.ScaleSetting setting))
      {
        hint = $"Sized for {setting.DisplayText} model space.";
        return 12.0 / setting.PaperInchesPerModelFoot;
      }
      hint = "Model space has no drawing scale; run SETSCALE (SS) or type a height.";
      return 1.0;
    }

    private static NoteTableTabItem GetSelectedNoteTab(NoteTableSession session)
    {
      if (session.Tabs.Count == 0)
      {
        return null;
      }
      int index = Math.Max(0, Math.Min(session.SelectedTabIndex, session.Tabs.Count - 1));
      return session.Tabs[index];
    }

    private static bool TrySaveNoteTableProject(NoteTableSession session)
    {
      string path = session.ProjectFilePath;
      if (string.IsNullOrEmpty(path))
      {
        return true;
      }

      if (NoteTableStore.GetLastWriteTimeUtc(path) != session.LoadedFileTimeUtc &&
          System.Windows.MessageBox.Show(
            $"{NoteTableStore.FileName} in this folder was changed by someone else " +
            "after NOTETABLE opened it.\n\nSave anyway and replace their changes?",
            "Note Tables",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Warning) != System.Windows.MessageBoxResult.Yes)
      {
        session.StatusMessage =
          "Not saved: the project's note file changed after NOTETABLE opened it. " +
          "Cancel and run NOTETABLE again to see those changes, or save again to replace them.";
        return false;
      }

      var project = new NoteTableProject
      {
        Tabs = session.Tabs
          .Select(tab => new NoteTableTabData
          {
            Name = tab.Name.Trim(),
            Kind = tab.Kind,
            FirstNumber = tab.FirstNumber,
            Notes = tab.GetNoteTexts().ToList(),
          })
          .ToList(),
      };
      try
      {
        NoteTableStore.Save(path, project);
        session.LoadedFileTimeUtc = NoteTableStore.GetLastWriteTimeUtc(path);
        return true;
      }
      catch (System.Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
      {
        session.StatusMessage = $"Could not save {path}: {ex.Message}";
        return false;
      }
    }

    private static void ImportPreselectedNoteObjects(
      Database database,
      Editor editor,
      string commandName,
      NoteTableSession session,
      NoteTableTabItem tab,
      ObjectId[] ids)
    {
      ObjectId tableId = FindSelectedTableId(database, ids);
      if (!tableId.IsNull)
      {
        ImportNoteTableObject(database, session, tab, tableId);
        return;
      }

      ObjectId[] textIds;
      using (Transaction transaction = database.TransactionManager.StartOpenCloseTransaction())
      {
        textIds = ids
          .Where(id => transaction.GetObject(id, OpenMode.ForRead, false) is MText ||
                       transaction.GetObject(id, OpenMode.ForRead, false) is DBText)
          .ToArray();
        transaction.Commit();
      }
      if (textIds.Length > 0)
      {
        ImportNoteTextObjects(database, editor, commandName, session, tab, textIds);
      }
    }

    private static void PromptImportNoteText(
      Database database,
      Editor editor,
      string commandName,
      NoteTableSession session,
      NoteTableTabItem tab)
    {
      var options = new PromptSelectionOptions
      {
        MessageForAdding = $"\nSelect MText or DBText notes to add to \"{tab.Header}\": ",
        MessageForRemoval = "\nRemove objects: ",
        RejectObjectsFromNonCurrentSpace = true,
      };
      var filter = new SelectionFilter(
        new[] { new TypedValue((int)DxfCode.Start, "TEXT,MTEXT") });
      PromptSelectionResult result = editor.GetSelection(options, filter);
      if (result.Status != PromptStatus.OK || result.Value == null || result.Value.Count == 0)
      {
        session.StatusMessage = "Import canceled; nothing was added.";
        return;
      }
      ImportNoteTextObjects(
        database,
        editor,
        commandName,
        session,
        tab,
        result.Value.GetObjectIds().Distinct().ToArray());
    }

    /// <summary>
    /// Splits the selected text into notes with KNT's or GNT's rules (GNT's
    /// also strips typed "1." numbers) and adds them to the tab.
    /// </summary>
    private static void ImportNoteTextObjects(
      Database database,
      Editor editor,
      string commandName,
      NoteTableSession session,
      NoteTableTabItem tab,
      ObjectId[] textIds)
    {
      var notes = new List<string>();
      var heights = new List<double>();
      var import = new NoteTableImport();
      double tableWidth;
      if (tab.Kind == NoteTableKind.Keyed)
      {
        KeyNoteTablePlan plan = ReadKeyNoteTextSelection(database, editor, commandName, textIds);
        if (plan == null)
        {
          session.StatusMessage = "No notes were imported; the command line says why.";
          return;
        }
        List<KeyNoteEntry> entries = plan.Columns.SelectMany(column => column.Entries).ToList();
        notes.AddRange(entries.Select(entry => entry.PlainText));
        heights.AddRange(entries.Select(entry => entry.TextHeight));
        import.TextStyleId = entries[0].TextStyleId;
        import.ColorIndex = entries[0].ColorIndex;
        import.SetPlacement(
          GetKeyNoteTableDefaultPosition(plan),
          plan.Rotation,
          plan.Layer);
        tableWidth = plan.Columns.Max(column =>
          GetKeyNoteSymbolColumnWidth(column) + column.SourceWidth +
          NoteTableTextIndent * 2.0 * GetKeyNoteColumnScale(column));
      }
      else
      {
        GeneralNoteTablePlan plan = ReadGeneralNoteTextSelection(database, editor, commandName, textIds);
        if (plan == null)
        {
          session.StatusMessage = "No notes were imported; the command line says why.";
          return;
        }
        List<GeneralNoteEntry> entries = plan.Columns.SelectMany(column => column.Entries).ToList();
        notes.AddRange(entries.Select(entry => entry.PlainText));
        heights.AddRange(entries.Select(entry => entry.TextHeight));
        import.TextStyleId = entries[0].TextStyleId;
        import.ColorIndex = entries[0].ColorIndex;
        import.SetPlacement(
          GetGeneralNoteTableDefaultPosition(plan),
          plan.Rotation,
          plan.Layer);
        tableWidth = plan.Columns.Max(column =>
          GetGeneralNoteNumberColumnWidth(column) + column.SourceWidth +
          NoteTableTextIndent * 2.0 * GetGeneralNoteColumnScale(column));
      }

      // The first import decides where "replace in place" puts the table;
      // later ones only add their text to the erase list.
      NoteTableImport existing = tab.Import;
      if (existing != null && existing.HasPlacement)
      {
        existing.SourceTextIds.AddRange(textIds);
      }
      else
      {
        import.SourceTextIds.AddRange(textIds);
        import.TableId = existing?.TableId ?? ObjectId.Null;
        tab.Import = import;
        session.TableWidth = tableWidth;
      }
      session.TextHeight = Median(heights);
      tab.HasImportedText = true;
      tab.HasImportedTable = !tab.Import.TableId.IsNull;
      tab.AddNotes(notes);
      tab.ImportSummary =
        $"Added {notes.Count} note{(notes.Count == 1 ? string.Empty : "s")} from " +
        $"{textIds.Length} text object{(textIds.Length == 1 ? string.Empty : "s")}.";
      session.StatusMessage = null;
    }

    private static void PromptImportNoteTable(
      Database database,
      Editor editor,
      NoteTableSession session,
      NoteTableTabItem tab)
    {
      var options = new PromptEntityOptions(
        $"\nSelect a note table to add to \"{tab.Header}\": ");
      options.SetRejectMessage("\nThat is not a table.");
      options.AddAllowedClass(typeof(Table), false);
      PromptEntityResult result = editor.GetEntity(options);
      if (result.Status != PromptStatus.OK)
      {
        session.StatusMessage = "Import canceled; nothing was added.";
        return;
      }
      ImportNoteTableObject(database, session, tab, result.ObjectId);
    }

    /// <summary>
    /// Adds the notes of a KNT or GNT style table (number column, note column,
    /// repeated for each overflow column) to the tab, top to bottom and then
    /// left to right.
    /// </summary>
    private static void ImportNoteTableObject(
      Database database,
      NoteTableSession session,
      NoteTableTabItem tab,
      ObjectId tableId)
    {
      var notes = new List<string>();
      int? firstNumber = null;
      double textHeight = 0.0;
      var import = new NoteTableImport { TableId = tableId };
      string handle;
      using (Transaction transaction = database.TransactionManager.StartOpenCloseTransaction())
      {
        Table table = (Table)transaction.GetObject(tableId, OpenMode.ForRead);
        handle = table.Handle.ToString();
        if (table.Columns.Count < 2)
        {
          session.StatusMessage =
            $"Table {handle} has only one column; a note table has a number column and a note column.";
          return;
        }

        for (int numberColumn = 0; numberColumn + 1 < table.Columns.Count; numberColumn += 2)
        {
          int textColumn = numberColumn + 1;
          for (int row = 0; row < table.Rows.Count; row++)
          {
            if (IsKeyNoteTableMergedHeader(table, row, numberColumn) ||
                IsKeyNoteTableMergedHeader(table, row, textColumn))
            {
              continue;
            }
            string text = ReadNoteTableCellText(table, row, textColumn);
            if (text.Length == 0)
            {
              continue;
            }
            notes.Add(text);
            if (!firstNumber.HasValue &&
                TryGetCellNumber(table, transaction, row, numberColumn, out int number, out _, out _))
            {
              firstNumber = number;
            }
            if (textHeight <= 0.0)
            {
              textHeight = TryGetNoteTableCellTextHeight(table, row, textColumn);
              import.TextStyleId = TryGetNoteTableCellTextStyle(table, row, textColumn);
            }
          }
        }

        ReadNoteTablePlacement(table, import);
        transaction.Commit();
      }

      if (notes.Count == 0)
      {
        session.StatusMessage = $"Table {handle} has no note text to import.";
        return;
      }

      bool wasEmpty = tab.GetNoteTexts().Count == 0;
      NoteTableImport existing = tab.Import;
      if (existing != null)
      {
        import.SourceTextIds.AddRange(existing.SourceTextIds);
      }
      tab.Import = import;
      tab.HasImportedText = import.SourceTextIds.Count > 0;
      tab.HasImportedTable = true;
      if (wasEmpty && firstNumber.HasValue && firstNumber.Value >= 1)
      {
        tab.FirstNumberText = firstNumber.Value.ToString();
      }
      if (textHeight > 0.0)
      {
        session.TextHeight = textHeight;
      }
      session.TableWidth = import.PairWidth;
      session.ReplaceImportedTable = true;
      tab.AddNotes(notes);
      tab.ImportSummary =
        $"Added {notes.Count} note{(notes.Count == 1 ? string.Empty : "s")} from table {handle}.";
      session.StatusMessage = null;
    }

    /// <summary>
    /// Records where a note table is and how it flows, so a replacement can
    /// take its place. An older table whose notes sit in side-by-side column
    /// pairs without overflow becomes a table that overflows at its height.
    /// </summary>
    private static void ReadNoteTablePlacement(Table table, NoteTableImport placement)
    {
      placement.SetPlacement(table.Position, table.Rotation, table.Layer);
      placement.PairWidth = table.Columns[0].Width + table.Columns[1].Width;
      placement.BreakEnabled = table.BreakEnabled;
      if (placement.BreakEnabled)
      {
        placement.BreakFlowDirection = table.BreakFlowDirection;
        placement.BreakHeight = table.GetBreakHeight(0);
        placement.BreakSpacing = table.GetBreakSpacing();
      }
      else if (table.Columns.Count > 3)
      {
        placement.BreakEnabled = true;
        placement.BreakFlowDirection = TableBreakFlowDirection.Right;
        placement.BreakHeight = table.Height;
        placement.BreakSpacing = KeyNoteTableColumnGap * GetNoteTableScale(table);
      }
    }

    private static string ReadNoteTableCellText(Table table, int row, int column)
    {
      string contents;
      try
      {
        if (table.GetNumberOfContents(row, column) <= 0)
        {
          return string.Empty;
        }
        contents = table.GetTextString(row, column, 0);
      }
      catch (Autodesk.AutoCAD.Runtime.Exception)
      {
        return string.Empty;
      }
      return NoteTableStore.NormalizeNoteText(string.Join(
        "\n",
        SplitRawMTextParagraphs(contents).Select(RemoveMTextFormatting)));
    }

    private static double TryGetNoteTableCellTextHeight(Table table, int row, int column)
    {
      try
      {
        return table.GetTextHeight(row, column, 0);
      }
      catch (Autodesk.AutoCAD.Runtime.Exception)
      {
        return 0.0;
      }
    }

    private static ObjectId TryGetNoteTableCellTextStyle(Table table, int row, int column)
    {
      try
      {
        return table.GetTextStyleId(row, column, 0);
      }
      catch (Autodesk.AutoCAD.Runtime.Exception)
      {
        return ObjectId.Null;
      }
    }

    private static bool TryCreateNoteTable(
      Database database,
      Editor editor,
      NoteTableSession session,
      NoteTableTabItem tab)
    {
      NoteTableImport import = tab.Import;
      bool replaceTable = import != null &&
        !import.TableId.IsNull &&
        !import.TableId.IsErased &&
        session.ReplaceImportedTable;
      var eraseIds = new List<ObjectId>();
      if (replaceTable)
      {
        eraseIds.Add(import.TableId);
      }
      bool eraseText = import != null && session.EraseImportedText;
      if (eraseText)
      {
        eraseIds.AddRange(import.SourceTextIds.Where(id => !id.IsNull && !id.IsErased));
      }

      var request = new NoteTableCreateRequest
      {
        Notes = tab.GetNoteTexts(),
        FirstNumber = tab.FirstNumber,
        TextHeight = session.TextHeight,
        TableWidth = session.TableWidth,
        UseBoundary = session.UseBoundary,
        TextStyleId = import != null && !import.TextStyleId.IsNull && import.TextStyleId.IsValid
          ? import.TextStyleId
          : database.Textstyle,
        ColorIndex = import?.ColorIndex ?? 256,
        Import = import,
        ReplaceTable = replaceTable,
        EraseIds = eraseIds.Distinct().ToList(),
      };

      ObjectId tableId = PlaceNoteTable(database, editor, session, tab, request);
      if (tableId.IsNull)
      {
        return false;
      }
      editor.WriteMessage(
        $"\nCreated {(tab.Kind == NoteTableKind.Keyed ? "keyed" : "general")}-note table " +
        $"{tableId.Handle} from \"{tab.Header}\" with {request.Notes.Count} " +
        $"note{(request.Notes.Count == 1 ? string.Empty : "s")}." +
        (replaceTable ? $" Replaced table {import.TableId.Handle}." : string.Empty) +
        (eraseText ? " Imported text erased." : string.Empty));
      return true;
    }

    /// <summary>
    /// Redraws a picked note table with the tab's notes and numbering. The new
    /// table takes the old one's place, width, overflow and text size, and the
    /// old one is erased in the same step, so one Undo restores it.
    /// </summary>
    private static bool TryUpdateNoteTable(
      Database database,
      Editor editor,
      NoteTableSession session,
      NoteTableTabItem tab)
    {
      var options = new PromptEntityOptions(
        $"\nSelect the note table to update with \"{tab.Header}\": ");
      options.SetRejectMessage("\nThat is not a table.");
      options.AddAllowedClass(typeof(Table), false);
      PromptEntityResult result = editor.GetEntity(options);
      if (result.Status != PromptStatus.OK)
      {
        session.StatusMessage = "Update canceled; no table was changed.";
        return false;
      }

      var placement = new NoteTableImport { TableId = result.ObjectId };
      double textHeight = session.TextHeight;
      using (Transaction transaction = database.TransactionManager.StartOpenCloseTransaction())
      {
        var table = (Table)transaction.GetObject(result.ObjectId, OpenMode.ForRead);
        if (table.Columns.Count < 2)
        {
          session.StatusMessage =
            $"Table {table.Handle} has only one column; a note table has a number column and a note column.";
          return false;
        }
        ReadNoteTablePlacement(table, placement);
        if (TryGetNoteTableTextHeight(table, out double tableTextHeight))
        {
          textHeight = tableTextHeight;
        }
        transaction.Commit();
      }

      var request = new NoteTableCreateRequest
      {
        Notes = tab.GetNoteTexts(),
        FirstNumber = tab.FirstNumber,
        TextHeight = textHeight,
        TableWidth = placement.PairWidth,
        TextStyleId = database.Textstyle,
        ColorIndex = 256,
        Import = placement,
        ReplaceTable = true,
        EraseIds = new List<ObjectId> { result.ObjectId },
      };
      ObjectId tableId = PlaceNoteTable(database, editor, session, tab, request);
      if (tableId.IsNull)
      {
        return false;
      }
      editor.WriteMessage(
        $"\nUpdated note table {result.ObjectId.Handle} with the {request.Notes.Count} " +
        $"note{(request.Notes.Count == 1 ? string.Empty : "s")} in \"{tab.Header}\" " +
        $"(now table {tableId.Handle}).");
      return true;
    }

    /// <summary>
    /// Draws the request as the tab's kind of table. Returns ObjectId.Null,
    /// with the reason in the session status, when it fails or is canceled.
    /// </summary>
    private static ObjectId PlaceNoteTable(
      Database database,
      Editor editor,
      NoteTableSession session,
      NoteTableTabItem tab,
      NoteTableCreateRequest request)
    {
      ObjectId tableId;
      try
      {
        tableId = tab.Kind == NoteTableKind.Keyed
          ? PlaceKeyedNoteTable(database, editor, request)
          : PlaceGeneralNoteTable(database, editor, request);
      }
      catch (System.Exception ex)
      {
        session.StatusMessage = $"The table could not be created: {ex.Message}";
        return ObjectId.Null;
      }
      if (tableId.IsNull)
      {
        session.StatusMessage = "Table placement was canceled. Your notes are still here.";
      }
      return tableId;
    }

    private static ObjectId PlaceKeyedNoteTable(
      Database database,
      Editor editor,
      NoteTableCreateRequest request)
    {
      var column = new KeyNoteSourceColumn { SourceWidth = request.TableWidth };
      foreach (string note in request.Notes)
      {
        column.Entries.Add(new KeyNoteEntry
        {
          PlainText = note,
          TextHeight = request.TextHeight,
          TextStyleId = request.TextStyleId,
          ColorIndex = request.ColorIndex,
        });
      }
      var plan = new KeyNoteTablePlan
      {
        Columns = new List<KeyNoteSourceColumn> { column },
        Rotation = request.Import?.Rotation ?? 0.0,
        Layer = request.Import?.Layer ?? string.Empty,
      };

      Point3d insertionPoint;
      if (request.ReplaceTable)
      {
        NoteTableImport replaced = request.Import;
        insertionPoint = replaced.Position.Value;
        plan.ForcedPairWidth = replaced.PairWidth;
        plan.BreakEnabled = replaced.BreakEnabled;
        plan.BreakFlowDirection = replaced.BreakFlowDirection;
        plan.BreakHeight = replaced.BreakHeight;
        plan.BreakSpacing = replaced.BreakSpacing;
      }
      else if (request.UseBoundary)
      {
        if (!TryPromptKeyNoteBoundary(editor, plan, out KeyNoteTablePlan boundaryPlan, out insertionPoint))
        {
          return ObjectId.Null;
        }
        plan = boundaryPlan;
      }
      else
      {
        if (!TryPromptNoteTableCorner(editor, request.Import?.Position, out insertionPoint))
        {
          return ObjectId.Null;
        }
        plan.ForcedPairWidth = request.TableWidth;
      }

      EnsureKnLayer(database);
      ObjectId knTextStyleId = EnsureKnTextStyle(database);
      if (!EnsureKnBlockDefinition(database, knTextStyleId))
      {
        throw new InvalidOperationException($"Unable to prepare the {KnBlockName} block.");
      }
      return WriteKeyNoteTable(
        database,
        plan,
        insertionPoint,
        request.FirstNumber,
        request.EraseIds,
        request.EraseIds.Count > 0);
    }

    private static ObjectId PlaceGeneralNoteTable(
      Database database,
      Editor editor,
      NoteTableCreateRequest request)
    {
      var column = new GeneralNoteSourceColumn { SourceWidth = request.TableWidth };
      foreach (string note in request.Notes)
      {
        column.Entries.Add(new GeneralNoteEntry
        {
          PlainText = note,
          TextHeight = request.TextHeight,
          TextStyleId = request.TextStyleId,
          ColorIndex = request.ColorIndex,
        });
      }
      var plan = new GeneralNoteTablePlan
      {
        Columns = new List<GeneralNoteSourceColumn> { column },
        Rotation = request.Import?.Rotation ?? 0.0,
        Layer = request.Import?.Layer ?? string.Empty,
      };

      Point3d insertionPoint;
      if (request.ReplaceTable)
      {
        NoteTableImport replaced = request.Import;
        insertionPoint = replaced.Position.Value;
        plan.ForcedPairWidth = replaced.PairWidth;
        plan.BreakEnabled = replaced.BreakEnabled;
        plan.BreakFlowDirection = replaced.BreakFlowDirection;
        plan.BreakHeight = replaced.BreakHeight;
        plan.BreakSpacing = replaced.BreakSpacing;
      }
      else if (request.UseBoundary)
      {
        if (!TryPromptGeneralNoteBoundary(editor, plan, out GeneralNoteTablePlan boundaryPlan, out insertionPoint))
        {
          return ObjectId.Null;
        }
        plan = boundaryPlan;
      }
      else
      {
        if (!TryPromptNoteTableCorner(editor, request.Import?.Position, out insertionPoint))
        {
          return ObjectId.Null;
        }
        plan.ForcedPairWidth = request.TableWidth;
      }

      return WriteGeneralNoteTable(
        database,
        plan,
        insertionPoint,
        request.FirstNumber,
        request.EraseIds,
        request.EraseIds.Count > 0);
    }

    private static bool TryPromptNoteTableCorner(
      Editor editor,
      Point3d? importedPosition,
      out Point3d insertionPoint)
    {
      var options = new PromptPointOptions(importedPosition.HasValue
        ? "\nSpecify table upper-left corner <where the imported notes are>: "
        : "\nSpecify table upper-left corner: ")
      {
        AllowNone = importedPosition.HasValue,
      };
      PromptPointResult result = editor.GetPoint(options);
      if (result.Status == PromptStatus.None && importedPosition.HasValue)
      {
        insertionPoint = importedPosition.Value;
        return true;
      }
      insertionPoint = result.Value;
      return result.Status == PromptStatus.OK;
    }

    private sealed class NoteTableCreateRequest
    {
      internal IReadOnlyList<string> Notes { get; set; }
      internal int FirstNumber { get; set; }
      internal double TextHeight { get; set; }
      internal double TableWidth { get; set; }
      internal bool UseBoundary { get; set; }
      internal ObjectId TextStyleId { get; set; }
      internal short ColorIndex { get; set; }
      internal NoteTableImport Import { get; set; }
      internal bool ReplaceTable { get; set; }
      internal List<ObjectId> EraseIds { get; set; }
    }
  }

  /// <summary>The drawing objects a tab's notes were imported from.</summary>
  internal sealed class NoteTableImport
  {
    internal List<ObjectId> SourceTextIds { get; } = new List<ObjectId>();
    internal ObjectId TableId { get; set; } = ObjectId.Null;
    internal Point3d? Position { get; private set; }
    internal double Rotation { get; private set; }
    internal string Layer { get; private set; } = string.Empty;
    internal ObjectId TextStyleId { get; set; } = ObjectId.Null;
    internal short ColorIndex { get; set; } = 256;
    internal double PairWidth { get; set; }
    internal bool BreakEnabled { get; set; }
    internal TableBreakFlowDirection BreakFlowDirection { get; set; } =
      TableBreakFlowDirection.Right;
    internal double BreakHeight { get; set; }
    internal double BreakSpacing { get; set; }
    internal bool HasPlacement => Position.HasValue;

    internal void SetPlacement(Point3d position, double rotation, string layer)
    {
      Position = position;
      Rotation = rotation;
      Layer = layer ?? string.Empty;
    }
  }
}
