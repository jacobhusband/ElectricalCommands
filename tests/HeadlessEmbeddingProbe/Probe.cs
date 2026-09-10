using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;
using System.Reflection;
using System.Text;

// Experimental harness only: not included in the installed command bundle.
public static class HeadlessEmbeddingProbe
{
    [CommandMethod("ACIESPDFFIXTURE")]
    public static void PdfFixture()
    {
        string directory = Environment.GetEnvironmentVariable("ACIES_EMBED_PROBE_DIR");
        var db = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
        using (var tr = db.TransactionManager.StartTransaction())
        {
            var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForWrite);
            var dictionary = new DBDictionary();
            nod.SetAt("ACAD_PDFDEFINITIONS", dictionary);
            tr.AddNewlyCreatedDBObject(dictionary, true);
            var definition = new PdfDefinition { SourceFileName = Path.Combine(directory, "form.pdf"), ItemName = "2" };
            dictionary.SetAt("fixture", definition);
            tr.AddNewlyCreatedDBObject(definition, true);
            var table = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var model = (BlockTableRecord)tr.GetObject(table[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
            var reference = new PdfReference { DefinitionId = definition.ObjectId,
                Position = new Autodesk.AutoCAD.Geometry.Point3d(10, 20, 0),
                ScaleFactors = new Autodesk.AutoCAD.Geometry.Scale3d(2), Rotation = Math.PI / 2 };
            model.AppendEntity(reference);
            tr.AddNewlyCreatedDBObject(reference, true);
            tr.Commit();
        }
        db.SaveAs(Path.Combine(directory, "input.dwg"), DwgVersion.Current);
    }

    [CommandMethod("ACIESOLEPROBE")]
    public static void Probe()
    {
        var report = new StringBuilder();
        string directory = Environment.GetEnvironmentVariable("ACIES_EMBED_PROBE_DIR");
        using (var ole = new Ole2Frame())
        {
            foreach (string property in new[] { "OleObject", "Type", "IsLinked", "UserType" })
            {
                try
                {
                    object value = typeof(Ole2Frame).GetProperty(property).GetValue(ole, null);
                    report.AppendLine(property + ": " + (value == null ? "<null>" : value.GetType().FullName + " = " + value));
                    if (property == "OleObject" && value != null)
                    {
                        foreach (PropertyInfo info in value.GetType().GetProperties())
                            report.AppendLine("  " + info.Name + ": " + info.PropertyType + " writable=" + info.CanWrite);
                    }
                }
                catch (System.Exception ex) { report.AppendLine(property + ": " + ex); }
            }
        }
        File.WriteAllText(Path.Combine(directory, "ole-probe.txt"), report.ToString());
    }
}
