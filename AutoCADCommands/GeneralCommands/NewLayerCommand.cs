using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands
{
  public partial class GeneralCommands
  {
    [CommandMethod("NEWLAYER", CommandFlags.Modal)]
    [CommandMethod("NL", CommandFlags.Modal)]
    public static void CreateNewLayer()
    {
      var (doc, db, ed) = Globals.GetGlobals();
      if (doc == null || db == null || ed == null) return;

      PromptResult nameResult = ed.GetString(new PromptStringOptions("\nEnter new layer name: ")
      {
        AllowSpaces = true,
      });
      if (nameResult.Status != PromptStatus.OK)
      {
        ed.WriteMessage("\nNEWLAYER canceled.");
        return;
      }
      string layerName = (nameResult.StringResult ?? string.Empty).Trim();
      try
      {
        SymbolUtilityServices.ValidateSymbolName(layerName, false);
      }
      catch (Autodesk.AutoCAD.Runtime.Exception)
      {
        ed.WriteMessage($"\n\"{layerName}\" is not a valid layer name.");
        return;
      }

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        var layers = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

        // Like -LAYER Make: an existing layer is just made current, never modified.
        if (layers.Has(layerName))
        {
          db.Clayer = layers[layerName];
          tr.Commit();
          ed.WriteMessage($"\nLayer \"{layerName}\" already exists and is now current.");
          return;
        }

        PromptIntegerResult colorResult = ed.GetInteger(new PromptIntegerOptions(
          "\nEnter color number (1-255) <7>: ")
        {
          AllowNone = true,
          LowerLimit = 1,
          UpperLimit = 255,
          DefaultValue = 7,
          UseDefaultValue = true,
        });
        if (colorResult.Status != PromptStatus.OK && colorResult.Status != PromptStatus.None)
        {
          ed.WriteMessage("\nNEWLAYER canceled.");
          return;
        }
        short colorIndex = (short)(colorResult.Status == PromptStatus.OK ? colorResult.Value : 7);

        var layer = new LayerTableRecord
        {
          Name = layerName,
          Color = Color.FromColorIndex(ColorMethod.ByAci, colorIndex),
        };
        layers.UpgradeOpen();
        ObjectId layerId = layers.Add(layer);
        tr.AddNewlyCreatedDBObject(layer, true);
        db.Clayer = layerId;
        tr.Commit();

        ed.WriteMessage($"\nCreated layer \"{layerName}\" (color {colorIndex}) and set it current.");
      }
    }
  }
}
