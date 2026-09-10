using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    private const string LayerFallbackBaseName = "ACIES-FREEZE-FALLBACK";

    [CommandMethod("OPENLAYERS", CommandFlags.Session)]
    [CommandMethod("OPENFREEZETHAW", CommandFlags.Session)]
    [CommandMethod("OLAYERS", CommandFlags.Session)]
    public void ManageLayersInOpenDrawings()
    {
      DocumentCollection documents = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
      Document commandDocument = documents.MdiActiveDocument;
      if (commandDocument == null)
      {
        MessageBox.Show(
          "No open AutoCAD drawings are available.",
          "Open Drawing Layer Manager",
          MessageBoxButton.OK,
          MessageBoxImage.Information);
        return;
      }

      var snapshots = new List<OpenDrawingLayerSnapshot>();
      var scanErrors = new List<string>();

      foreach (Document document in documents)
      {
        try
        {
          snapshots.Add(CaptureLayerSnapshot(document));
        }
        catch (System.Exception ex)
        {
          scanErrors.Add($"{GetDrawingDisplayName(document)}: {ex.Message}");
        }
      }

      if (snapshots.Count == 0)
      {
        string details = scanErrors.Count == 0
          ? string.Empty
          : $"\n\n{string.Join("\n", scanErrors)}";
        MessageBox.Show(
          $"No open drawings could be scanned.{details}",
          "Open Drawing Layer Manager",
          MessageBoxButton.OK,
          MessageBoxImage.Warning);
        return;
      }

      var window = new OpenDrawingLayerWindow(snapshots);
      bool? accepted = Autodesk.AutoCAD.ApplicationServices.Application.ShowModalWindow(window);
      if (accepted != true)
      {
        commandDocument.Editor.WriteMessage("\nOPENLAYERS: No layer changes were applied.");
        return;
      }

      var results = new List<OpenDrawingLayerChangeResult>();
      IReadOnlyList<OpenDrawingLayerSnapshot> selectedDrawings = window.SelectedDrawings;
      IReadOnlyDictionary<string, OpenDrawingLayerAction> layerActions = window.LayerActions;
      bool saveDrawingsAfterApplying = window.SaveDrawingsAfterApplying;
      Document originalActiveDocument = documents.MdiActiveDocument;

      try
      {
        foreach (OpenDrawingLayerSnapshot drawing in selectedDrawings)
        {
          results.Add(ApplyLayerChangesInDocumentContext(
            documents,
            drawing,
            layerActions,
            saveDrawingsAfterApplying));
        }
      }
      finally
      {
        if (originalActiveDocument != null && IsDocumentOpen(documents, originalActiveDocument))
        {
          try
          {
            documents.MdiActiveDocument = originalActiveDocument;
          }
          catch (System.Exception ex)
          {
            scanErrors.Add($"Could not restore the originally active drawing: {ex.Message}");
          }
        }
      }

      string summary = BuildLayerChangeSummary(results, scanErrors);
      Document outputDocument = documents.MdiActiveDocument ?? commandDocument;
      outputDocument.Editor.WriteMessage($"\n{summary.Replace(Environment.NewLine, " ")}");

      MessageBoxImage summaryIcon = results.Any(result => result.Errors.Count > 0)
        ? MessageBoxImage.Warning
        : MessageBoxImage.Information;
      MessageBox.Show(
        summary,
        "Open Drawing Layer Manager",
        MessageBoxButton.OK,
        summaryIcon);
    }

    private static OpenDrawingLayerSnapshot CaptureLayerSnapshot(Document document)
    {
      if (document == null) throw new ArgumentNullException(nameof(document));

      var layers = new Dictionary<string, OpenDrawingLayerState>(StringComparer.OrdinalIgnoreCase);
      Database database = document.Database;

      using (document.LockDocument())
      using (Transaction transaction = database.TransactionManager.StartOpenCloseTransaction())
      {
        var layerTable = (LayerTable)transaction.GetObject(database.LayerTableId, OpenMode.ForRead);
        ObjectId currentLayerId = database.Clayer;

        foreach (ObjectId layerId in layerTable)
        {
          if (layerId.IsErased) continue;

          var layer = transaction.GetObject(layerId, OpenMode.ForRead, false) as LayerTableRecord;
          if (layer == null || string.IsNullOrWhiteSpace(layer.Name)) continue;

          layers[layer.Name] = new OpenDrawingLayerState(
            layer.Name,
            layer.IsFrozen,
            layerId == currentLayerId);
        }

        transaction.Commit();
      }

      return new OpenDrawingLayerSnapshot(
        document,
        GetDrawingDisplayName(document),
        document.Name ?? string.Empty,
        layers);
    }

    private static OpenDrawingLayerChangeResult ApplyLayerChangesInDocumentContext(
      DocumentCollection documents,
      OpenDrawingLayerSnapshot drawing,
      IReadOnlyDictionary<string, OpenDrawingLayerAction> actions,
      bool saveAfterApplying)
    {
      var activationFailure = new OpenDrawingLayerChangeResult(drawing.DisplayName);
      if (drawing.Document == null || !IsDocumentOpen(documents, drawing.Document))
      {
        activationFailure.Errors.Add("The drawing is no longer open.");
        return activationFailure;
      }

      try
      {
        if (!ReferenceEquals(documents.MdiActiveDocument, drawing.Document))
        {
          documents.MdiActiveDocument = drawing.Document;
        }

        if (!ReferenceEquals(documents.MdiActiveDocument, drawing.Document))
        {
          activationFailure.Errors.Add("AutoCAD did not activate the selected drawing.");
          return activationFailure;
        }

        activationFailure.WasActivated = true;
        return ApplyLayerChanges(drawing, actions, saveAfterApplying);
      }
      catch (System.Exception ex)
      {
        activationFailure.Errors.Add($"Could not activate the drawing: {ex.Message}");
        return activationFailure;
      }
    }

    private static OpenDrawingLayerChangeResult ApplyLayerChanges(
      OpenDrawingLayerSnapshot drawing,
      IReadOnlyDictionary<string, OpenDrawingLayerAction> actions,
      bool saveAfterApplying)
    {
      var result = new OpenDrawingLayerChangeResult(drawing.DisplayName);
      result.SaveRequested = saveAfterApplying;
      if (drawing.Document == null)
      {
        result.Errors.Add("The drawing is no longer open.");
        return result;
      }

      try
      {
        Document document = drawing.Document;
        Database database = document.Database;
        Database previousWorkingDatabase = HostApplicationServices.WorkingDatabase;

        try
        {
          HostApplicationServices.WorkingDatabase = database;

          using (document.LockDocument())
          {
            using (Transaction transaction = database.TransactionManager.StartTransaction())
            {
              var layerTable = (LayerTable)transaction.GetObject(database.LayerTableId, OpenMode.ForRead);
              var freezeNames = new HashSet<string>(
                actions
                  .Where(pair => pair.Value == OpenDrawingLayerAction.Freeze)
                  .Select(pair => pair.Key),
                StringComparer.OrdinalIgnoreCase);

              var thawNames = new HashSet<string>(
                actions
                  .Where(pair => pair.Value == OpenDrawingLayerAction.Thaw)
                  .Select(pair => pair.Key),
                StringComparer.OrdinalIgnoreCase);

              string currentLayerName = GetCurrentLayerName(database, transaction);
              if (!string.IsNullOrWhiteSpace(currentLayerName) &&
                  freezeNames.Contains(currentLayerName) &&
                  layerTable.Has(currentLayerName))
              {
                ObjectId fallbackLayerId = CreateUniqueFallbackLayer(
                  layerTable,
                  transaction,
                  freezeNames,
                  out string fallbackLayerName);

                database.Clayer = fallbackLayerId;
                result.FallbackLayerName = fallbackLayerName;
                result.PreviousCurrentLayerName = currentLayerName;
              }

              foreach (string layerName in thawNames.OrderBy(name => name, StringComparer.OrdinalIgnoreCase))
              {
                ApplyFrozenState(
                  layerTable,
                  transaction,
                  layerName,
                  shouldFreeze: false,
                  result);
              }

              foreach (string layerName in freezeNames.OrderBy(name => name, StringComparer.OrdinalIgnoreCase))
              {
                ApplyFrozenState(
                  layerTable,
                  transaction,
                  layerName,
                  shouldFreeze: true,
                  result);
              }

              transaction.Commit();
            }

            VerifyLayerChanges(database, actions, result);
            document.Editor.Regen();

            if (saveAfterApplying)
            {
              SaveDrawing(document, database, result);
            }
          }
        }
        finally
        {
          HostApplicationServices.WorkingDatabase = previousWorkingDatabase;
        }
      }
      catch (System.Exception ex)
      {
        result.Errors.Add(ex.Message);
      }

      return result;
    }

    private static void SaveDrawing(
      Document document,
      Database database,
      OpenDrawingLayerChangeResult result)
    {
      string drawingPath = !string.IsNullOrWhiteSpace(database.Filename)
        ? database.Filename
        : document.Name;

      if (string.IsNullOrWhiteSpace(drawingPath) ||
          !Path.IsPathRooted(drawingPath) ||
          !string.Equals(Path.GetExtension(drawingPath), ".dwg", StringComparison.OrdinalIgnoreCase))
      {
        result.Errors.Add("The drawing has not been named; use Save As to save its layer changes.");
        return;
      }

      try
      {
        DwgVersion saveVersion = database.OriginalFileVersion;
        if (saveVersion == DwgVersion.MC0To0 || saveVersion == DwgVersion.Max)
        {
          saveVersion = DwgVersion.Current;
        }

        database.SaveAs(
          drawingPath,
          true,
          saveVersion,
          database.SecurityParameters);
        result.WasSaved = true;
      }
      catch (System.Exception ex)
      {
        result.Errors.Add($"Could not save the drawing: {ex.Message}");
      }
    }

    private static void VerifyLayerChanges(
      Database database,
      IReadOnlyDictionary<string, OpenDrawingLayerAction> actions,
      OpenDrawingLayerChangeResult result)
    {
      using (Transaction transaction = database.TransactionManager.StartOpenCloseTransaction())
      {
        var layerTable = (LayerTable)transaction.GetObject(database.LayerTableId, OpenMode.ForRead);

        foreach (KeyValuePair<string, OpenDrawingLayerAction> action in actions)
        {
          if (!layerTable.Has(action.Key)) continue;

          result.ExpectedLayerCount++;
          ObjectId layerId = layerTable[action.Key];
          var layer = transaction.GetObject(layerId, OpenMode.ForRead, false) as LayerTableRecord;
          bool expectedFrozen = action.Value == OpenDrawingLayerAction.Freeze;
          if (layer != null && layer.IsFrozen == expectedFrozen)
          {
            result.VerifiedLayerCount++;
          }
          else
          {
            string expectedState = expectedFrozen ? "frozen" : "thawed";
            result.Errors.Add($"Verification failed for {action.Key}; expected {expectedState}.");
          }
        }

        if (!string.IsNullOrWhiteSpace(result.FallbackLayerName))
        {
          string verifiedCurrentLayer = GetCurrentLayerName(database, transaction);
          if (!string.Equals(
                verifiedCurrentLayer,
                result.FallbackLayerName,
                StringComparison.OrdinalIgnoreCase))
          {
            result.Errors.Add(
              $"Current-layer verification failed; expected {result.FallbackLayerName} but found {verifiedCurrentLayer}.");
          }
        }

        transaction.Commit();
      }
    }

    private static void ApplyFrozenState(
      LayerTable layerTable,
      Transaction transaction,
      string layerName,
      bool shouldFreeze,
      OpenDrawingLayerChangeResult result)
    {
      if (!layerTable.Has(layerName)) return;

      try
      {
        ObjectId layerId = layerTable[layerName];
        var layer = transaction.GetObject(layerId, OpenMode.ForWrite, false) as LayerTableRecord;
        if (layer == null) return;

        result.MatchedLayerCount++;
        if (layer.IsFrozen == shouldFreeze) return;

        layer.IsFrozen = shouldFreeze;
        if (shouldFreeze)
        {
          result.FrozenLayerCount++;
        }
        else
        {
          result.ThawedLayerCount++;
        }
      }
      catch (System.Exception ex)
      {
        result.Errors.Add($"{layerName}: {ex.Message}");
      }
    }

    private static string GetCurrentLayerName(Database database, Transaction transaction)
    {
      if (database.Clayer.IsNull || database.Clayer.IsErased) return string.Empty;

      var currentLayer = transaction.GetObject(database.Clayer, OpenMode.ForRead, false) as LayerTableRecord;
      return currentLayer?.Name ?? string.Empty;
    }

    private static ObjectId CreateUniqueFallbackLayer(
      LayerTable layerTable,
      Transaction transaction,
      ISet<string> layersToFreeze,
      out string layerName)
    {
      int suffix = 1;
      do
      {
        layerName = suffix == 1
          ? LayerFallbackBaseName
          : $"{LayerFallbackBaseName}-{suffix}";
        suffix++;
      }
      while (layerTable.Has(layerName) || layersToFreeze.Contains(layerName));

      layerTable.UpgradeOpen();
      var fallbackLayer = new LayerTableRecord { Name = layerName };
      ObjectId fallbackLayerId = layerTable.Add(fallbackLayer);
      transaction.AddNewlyCreatedDBObject(fallbackLayer, true);
      return fallbackLayerId;
    }

    private static string GetDrawingDisplayName(Document document)
    {
      if (document == null) return "Unknown drawing";

      try
      {
        string fileName = Path.GetFileName(document.Name);
        return string.IsNullOrWhiteSpace(fileName) ? document.Name : fileName;
      }
      catch
      {
        return string.IsNullOrWhiteSpace(document.Name) ? "Untitled drawing" : document.Name;
      }
    }

    private static bool IsDocumentOpen(DocumentCollection documents, Document target)
    {
      if (documents == null || target == null) return false;

      foreach (Document document in documents)
      {
        if (ReferenceEquals(document, target)) return true;
      }

      return false;
    }

    private static string BuildLayerChangeSummary(
      IReadOnlyCollection<OpenDrawingLayerChangeResult> results,
      IReadOnlyCollection<string> scanErrors)
    {
      int frozen = results.Sum(result => result.FrozenLayerCount);
      int thawed = results.Sum(result => result.ThawedLayerCount);
      int fallbackCount = results.Count(result => !string.IsNullOrWhiteSpace(result.FallbackLayerName));
      int saveRequestedCount = results.Count(result => result.SaveRequested);
      int savedCount = results.Count(result => result.WasSaved);
      int errorCount = results.Sum(result => result.Errors.Count) + scanErrors.Count;

      var summary = new StringBuilder();
      summary.AppendLine($"Processed {results.Count} open drawing(s).");
      summary.AppendLine($"Changed layers: {frozen} frozen, {thawed} thawed.");
      summary.AppendLine("Results by drawing:");
      foreach (OpenDrawingLayerChangeResult result in results)
      {
        string resultState = result.Errors.Count == 0 ? "OK" : "WARNING";
        int changedCount = result.FrozenLayerCount + result.ThawedLayerCount;
        string saveState = result.SaveRequested
          ? result.WasSaved ? "saved" : "not saved"
          : "save not requested";
        summary.AppendLine(
          $"  {result.DrawingName}: {resultState}; {result.VerifiedLayerCount}/{result.ExpectedLayerCount} verified; {changedCount} changed; {saveState}.");
      }

      if (fallbackCount > 0)
      {
        summary.AppendLine($"Created {fallbackCount} fallback layer(s) for drawings whose current layer was frozen:");
        foreach (OpenDrawingLayerChangeResult result in results.Where(item => !string.IsNullOrWhiteSpace(item.FallbackLayerName)))
        {
          summary.AppendLine(
            $"  {result.DrawingName}: {result.PreviousCurrentLayerName} -> {result.FallbackLayerName}");
        }
      }

      if (errorCount > 0)
      {
        summary.AppendLine($"Warnings/errors: {errorCount}");
        foreach (string error in scanErrors.Take(6))
        {
          summary.AppendLine($"  Scan: {error}");
        }

        foreach (OpenDrawingLayerChangeResult result in results)
        {
          foreach (string error in result.Errors.Take(6))
          {
            summary.AppendLine($"  {result.DrawingName}: {error}");
          }
        }
      }

      if (saveRequestedCount > 0)
      {
        summary.Append($"Saved {savedCount} of {saveRequestedCount} selected drawing(s).");
      }
      else
      {
        summary.Append("Drawings remain open and unsaved so the changes can be reviewed.");
      }
      return summary.ToString();
    }
  }

  internal sealed class OpenDrawingLayerSnapshot
  {
    public OpenDrawingLayerSnapshot(
      Document document,
      string displayName,
      string fullName,
      IReadOnlyDictionary<string, OpenDrawingLayerState> layers)
    {
      Document = document;
      DisplayName = displayName;
      FullName = fullName;
      Layers = layers;
      IsSelected = true;
    }

    public Document Document { get; }
    public string DisplayName { get; }
    public string FullName { get; }
    public IReadOnlyDictionary<string, OpenDrawingLayerState> Layers { get; }
    public bool IsSelected { get; set; }
  }

  internal sealed class OpenDrawingLayerState
  {
    public OpenDrawingLayerState(string name, bool isFrozen, bool isCurrent)
    {
      Name = name;
      IsFrozen = isFrozen;
      IsCurrent = isCurrent;
    }

    public string Name { get; }
    public bool IsFrozen { get; }
    public bool IsCurrent { get; }
  }

  internal sealed class OpenDrawingLayerChangeResult
  {
    public OpenDrawingLayerChangeResult(string drawingName)
    {
      DrawingName = drawingName;
    }

    public string DrawingName { get; }
    public int MatchedLayerCount { get; set; }
    public int FrozenLayerCount { get; set; }
    public int ThawedLayerCount { get; set; }
    public int ExpectedLayerCount { get; set; }
    public int VerifiedLayerCount { get; set; }
    public bool WasActivated { get; set; }
    public bool SaveRequested { get; set; }
    public bool WasSaved { get; set; }
    public string PreviousCurrentLayerName { get; set; }
    public string FallbackLayerName { get; set; }
    public List<string> Errors { get; } = new List<string>();
  }
}
