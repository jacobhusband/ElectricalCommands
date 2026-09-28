using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace AutoCADCleanupTool
{
    public partial class SimplerCommands
    {
        // Native geometry encoder shared by staged media cleanup and isolated experiments.
        private sealed class PixelRectangle
        {
            public int X, Y, Width, Height, Rgb;
        }

        private sealed class NativeImagePlan
        {
            public ObjectId ImageId;
            public int Width, Height;
            public List<PixelRectangle> Rectangles;
        }

        private static List<PixelRectangle> ReadPixelRectangles(string path, out int width, out int height, int remaining, bool transparent)
        {
            using (var source = new Bitmap(path))
            {
                width = source.Width;
                height = source.Height;
                if ((long)width * height > 2000000)
                    throw new InvalidOperationException("Image exceeds the prototype limit of 2,000,000 pixels: " + path);
                if (source.PixelFormat == PixelFormat.Format1bppIndexed)
                    throw new InvalidOperationException("Bitonal image foreground colors are not supported by the prototype: " + path);
                var rectangles = new List<PixelRectangle>();
                var previous = new Dictionary<Tuple<int, int, int>, PixelRectangle>();
                using (var bitmap = source.Clone(new Rectangle(0, 0, width, height), PixelFormat.Format32bppArgb))
                {
                    var data = bitmap.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                    try
                    {
                        var row = new int[width];
                        for (int y = 0; y < height; y++)
                        {
                            Marshal.Copy(IntPtr.Add(data.Scan0, y * data.Stride), row, 0, width);
                            for (int i = 0; i < width; i++)
                            {
                                uint alpha = (uint)row[i] >> 24;
                                if (alpha != 0 && alpha != 255)
                                    throw new InvalidOperationException("Alpha transparency with partial opacity is not supported: " + path);
                                if (!transparent) row[i] |= unchecked((int)0xff000000);
                            }
                            var current = new Dictionary<Tuple<int, int, int>, PixelRectangle>();
                            for (int x = 0; x < width;)
                            {
                                int color = row[x];
                                if ((uint)color >> 24 == 0) { x++; continue; }
                                int end = x + 1;
                                while (end < width && row[end] == color) end++;
                                var key = Tuple.Create(x, end - x, color);
                                if (previous.TryGetValue(key, out PixelRectangle rectangle)) rectangle.Height++;
                                else
                                {
                                    if (rectangles.Count >= remaining)
                                        throw new InvalidOperationException("Image geometry exceeds the prototype limit of 50,000 solids per drawing.");
                                    rectangle = new PixelRectangle { X = x, Y = y, Width = end - x, Height = 1, Rgb = color };
                                    rectangles.Add(rectangle);
                                }
                                current.Add(key, rectangle);
                                x = end;
                            }
                            previous = current;
                        }
                    }
                    finally { bitmap.UnlockBits(data); }
                }
                return rectangles;
            }
        }

        private static bool HasPartialImageClip(RasterImage image)
        {
            if (!image.IsClipped) return false;
            var boundary = image.GetClipBoundary();
            // AutoCAD enables a whole-image rectangular clip on fresh attachments.
            if (image.ClipBoundaryType != ClipBoundaryType.Rectangle) return true;
            double right = image.ImageWidth - 0.5, top = image.ImageHeight - 0.5;
            var expected = boundary.Count == 2
                ? new[] { new Point2d(-0.5, -0.5), new Point2d(right, top) }
                : new[] { new Point2d(-0.5, -0.5), new Point2d(-0.5, top), new Point2d(right, top), new Point2d(right, -0.5), new Point2d(-0.5, -0.5) };
            if (boundary.Count != expected.Length) return true;
            for (int i = 0; i < expected.Length; i++)
                if (boundary[i].GetDistanceTo(expected[i]) > 1e-8) return true;
            return false;
        }

        private static object EmbedNativeImageGeometry(Database db)
        {
            if (!db.Fillmode) throw new InvalidOperationException("Prototype requires FILLMODE=1 so native solids display filled.");
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var table = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var plans = new List<NativeImagePlan>();
                int count = 0;
                // Validate and decode every image before modifying any entity.
                foreach (ObjectId blockId in table)
                {
                    var owner = (BlockTableRecord)tr.GetObject(blockId, OpenMode.ForRead);
                    if (owner.IsFromExternalReference || owner.IsDependent) continue;
                    foreach (ObjectId id in owner)
                    {
                        var entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (!(entity is RasterImage image) || image is Wipeout) continue;
                        if (image.GetType() != typeof(RasterImage) || HasPartialImageClip(image) ||
                            image.Brightness != 50 || image.Contrast != 50 || image.Fade != 0)
                            throw new InvalidOperationException("Image conversion supports only unclipped raster images with default brightness/contrast and zero fade. Image " + image.Handle + " (clipped=" + image.IsClipped + ", transparency=" + image.ImageTransparency + ", brightness=" + image.Brightness + ", contrast=" + image.Contrast + ", fade=" + image.Fade + ")");
                        var layer = (LayerTableRecord)tr.GetObject(image.LayerId, OpenMode.ForRead);
                        if (layer.IsLocked) throw new InvalidOperationException("Unlock image layer before prototype conversion: " + layer.Name);
                        var axes = image.Orientation;
                        if (Math.Abs(axes.Xaxis.Z) > 1e-10 || Math.Abs(axes.Yaxis.Z) > 1e-10 ||
                            axes.Xaxis.CrossProduct(axes.Yaxis).Length < 1e-10)
                            throw new InvalidOperationException("Prototype supports nondegenerate images parallel to the owner XY plane only.");
                        var definition = (RasterImageDef)tr.GetObject(image.ImageDefId, OpenMode.ForRead);
                        // Use the exact stored reference relative to this DWG; never substitute a same-named image.
                        string path = definition.SourceFileName;
                        if (!Path.IsPathRooted(path)) path = Path.Combine(Path.GetDirectoryName(db.Filename), path);
                        if (!File.Exists(path)) throw new FileNotFoundException("Image source is missing.", path);
                        var rectangles = ReadPixelRectangles(path, out int width, out int height, 50000 - count, image.ImageTransparency);
                        if (Math.Abs(image.ImageWidth - width) > 0.01 || Math.Abs(image.ImageHeight - height) > 0.01)
                            throw new InvalidOperationException("Source image dimensions changed since attachment: " + path);
                        count += rectangles.Count;
                        plans.Add(new NativeImagePlan { ImageId = id, Width = width, Height = height, Rectangles = rectangles });
                    }
                }
                table.UpgradeOpen();
                foreach (var plan in plans)
                {
                    var image = (RasterImage)tr.GetObject(plan.ImageId, OpenMode.ForWrite);
                    var orientation = image.Orientation;
                    // Store owner-space coordinates. The enclosing block's transform (including nested instances) still applies.
                    Func<double, double, Point3d> point = (x, y) => orientation.Origin +
                        orientation.Xaxis * (x / plan.Width) + orientation.Yaxis * (1.0 - y / plan.Height);
                    var block = new BlockTableRecord { Name = "ACIES_IMAGE_" + Guid.NewGuid().ToString("N") };
                    table.Add(block);
                    tr.AddNewlyCreatedDBObject(block, true);
                    foreach (var rectangle in plan.Rectangles)
                    {
                        double right = rectangle.X + rectangle.Width, bottom = rectangle.Y + rectangle.Height;
                        // SOLID uses the third/fourth vertices in crossed order, not polygon perimeter order.
                        var solid = new Solid(point(rectangle.X, bottom), point(right, bottom), point(rectangle.X, rectangle.Y), point(right, rectangle.Y));
                        solid.SetDatabaseDefaults(db);
                        solid.Layer = "0";
                        solid.Color = Autodesk.AutoCAD.Colors.Color.FromRgb((byte)(rectangle.Rgb >> 16), (byte)(rectangle.Rgb >> 8), (byte)rectangle.Rgb);
                        solid.Transparency = new Transparency(TransparencyMethod.ByBlock);
                        block.AppendEntity(solid);
                        tr.AddNewlyCreatedDBObject(solid, true);
                    }
                    var replacement = new BlockReference(Point3d.Origin, block.ObjectId);
                    replacement.SetPropertiesFrom(image);
                    replacement.Visible = image.Visible && image.ShowImage;
                    var owner = (BlockTableRecord)tr.GetObject(image.OwnerId, OpenMode.ForWrite);
                    owner.AppendEntity(replacement);
                    tr.AddNewlyCreatedDBObject(replacement, true);
                    var order = (DrawOrderTable)tr.GetObject(owner.DrawOrderTableId, OpenMode.ForWrite);
                    order.MoveAbove(new ObjectIdCollection(new[] { replacement.ObjectId }), image.ObjectId);
                    image.Erase();
                }
                // All ordinary raster images were converted. Wipeouts have their own internal image definitions.
                var imageDictionaryId = RasterImageDef.GetImageDictionary(db);
                if (!imageDictionaryId.IsNull)
                {
                    var dictionary = (DBDictionary)tr.GetObject(imageDictionaryId, OpenMode.ForRead);
                    foreach (DBDictionaryEntry entry in dictionary)
                    {
                        var definition = tr.GetObject(entry.Value, OpenMode.ForRead) as RasterImageDef;
                        if (definition != null && definition.GetEntityCount(out bool locked) == 0)
                        { definition.UpgradeOpen(); definition.Erase(); }
                    }
                }
                tr.Commit();
                return new { images = plans.Count, solids = count, encoding = "native-truecolor-solids", experimental = true };
            }
        }
    }
}
