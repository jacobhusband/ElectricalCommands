using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

namespace ElectricalCommands {
  public partial class GeneralCommands {
    public static (Autodesk.AutoCAD.ApplicationServices.Document doc, Database db, Editor ed) GetGlobals() {
      var doc = Autodesk
          .AutoCAD
          .ApplicationServices
          .Application
          .DocumentManager
          .MdiActiveDocument;
      var db = doc.Database;
      var ed = doc.Editor;

      return (doc, db, ed);
    }
  }
}

