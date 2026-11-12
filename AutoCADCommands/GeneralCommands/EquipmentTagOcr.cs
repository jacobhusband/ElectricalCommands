using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace ElectricalCommands
{
  // Reads an equipment tag off the screen for ETAGOCR with the bundled Tesseract
  // engine. (Windows' built-in OCR was tried first but returns nothing for short
  // tags such as P-1.)
  internal static class EquipmentTagOcr
  {
    // Pixels trimmed off each edge of the picked box: AutoCAD's crosshair and
    // rubber-band rectangle sit right on it.
    private const int CaptureInset = 3;
    // RGB distance from the background beyond which a pixel counts as ink.
    private const double InkThreshold = 60.0;
    // Tuned against Arial, Arial Narrow, Courier, and SHX-style (RomanS,
    // Simplex, Txt) tags: 32-40 px letters with strokes at 0.14 of that read best.
    private const double TargetLetterHeight = 40.0;
    private const double TargetStrokeRatio = 0.14;
    private static readonly char[] EdgeJunk = "()[]{}<>|_.,:;'\"`~ ".ToCharArray();

    internal static string DebugFolder => Path.Combine(Path.GetTempPath(), "ETAGOCR");

    internal static Bitmap CaptureScreen(Point corner1, Point corner2)
    {
      int left = Math.Min(corner1.X, corner2.X) + CaptureInset;
      int top = Math.Min(corner1.Y, corner2.Y) + CaptureInset;
      int width = Math.Abs(corner1.X - corner2.X) - 2 * CaptureInset;
      int height = Math.Abs(corner1.Y - corner2.Y) - 2 * CaptureInset;
      if (width < 8 || height < 8) return null;

      Bitmap capture = new Bitmap(width, height, PixelFormat.Format32bppArgb);
      using (Graphics g = Graphics.FromImage(capture))
      {
        g.CopyFromScreen(left, top, 0, 0, new Size(width, height));
      }
      return capture;
    }

    // Cleans up a screen capture and runs OCR on it. Returns null with an error
    // message when nothing could be read. Both images are kept in DebugFolder.
    internal static string ReadImage(Bitmap capture, out string error)
    {
      Directory.CreateDirectory(DebugFolder);
      capture.Save(Path.Combine(DebugFolder, "last-capture.png"), ImageFormat.Png);

      string ocrInputPath = Path.Combine(DebugFolder, "last-ocr-input.png");
      using (Bitmap prepared = PrepareForOcr(capture))
      {
        if (prepared == null)
        {
          error = "the picked box has no visible text";
          return null;
        }
        prepared.Save(ocrInputPath, ImageFormat.Png);
      }

      List<string> lines = RunTesseract(ocrInputPath, out error);
      if (lines == null) return null;

      string text = CleanTagText(lines, true);
      if (string.IsNullOrEmpty(text))
      {
        error = "OCR found no text in the picked box";
        return null;
      }
      error = null;
      return text;
    }

    // Normalizes tag lines: collapses spaces, tightens "EF - 1" to "EF-1", strips
    // stray bracket and punctuation marks left by the tag's outline, and joins a
    // stacked tag ("EF" over "1") with a hyphen.
    internal static string CleanTagText(IEnumerable<string> lines, bool uppercase)
    {
      StringBuilder joined = new StringBuilder();
      foreach (string raw in lines)
      {
        if (raw == null) continue;
        string line = Regex.Replace(raw, @"\s+", " ");
        line = Regex.Replace(line, @"\s*-\s*", "-").Trim(EdgeJunk);
        if (uppercase) line = line.ToUpperInvariant();
        if (line.Length == 0) continue;

        if (joined.Length > 0 && joined[joined.Length - 1] != '-' && line[0] != '-')
        {
          joined.Append('-');
        }
        joined.Append(line);
      }
      return joined.ToString();
    }

    // Turns a capture into dark text on white, enlarged for OCR. Ink is anything
    // far enough from the most common color, so AutoCAD's dark model space and
    // white paper space both work. Shapes that are not text are dropped: ink
    // touching the box edge (crosshair, a clipped outline), a large hollow outline
    // around the text, and long horizontal divider lines.
    internal static Bitmap PrepareForOcr(Bitmap capture)
    {
      int width = capture.Width;
      int height = capture.Height;
      int[] pixels = ReadPixels(capture);
      Color background = FindBackground(pixels);

      double[] distance = new double[pixels.Length];
      bool[] ink = new bool[pixels.Length];
      for (int i = 0; i < pixels.Length; i++)
      {
        distance[i] = ColorDistance(Color.FromArgb(pixels[i]), background);
        ink[i] = distance[i] > InkThreshold;
      }

      bool[] text = KeepTextComponents(ink, width, height);
      if (!Bounds(text, width, height, out _))
      {
        // Everything was filtered out, so the box was probably drawn tight
        // against the text. Use all the ink rather than nothing.
        text = ink;
      }
      if (!Bounds(text, width, height, out Rectangle inkBounds)) return null;

      // Faint, thin text (SHX fonts drawn one pixel wide) comes out as pale
      // gray; stretch the contrast so typical ink is fully black.
      List<double> inkLevels = new List<double>();
      for (int i = 0; i < text.Length; i++)
      {
        if (text[i]) inkLevels.Add(distance[i]);
      }
      inkLevels.Sort();
      double fullInk = Math.Max(InkThreshold, inkLevels[inkLevels.Count / 2]);

      int pad = Math.Max(4, inkBounds.Height / 2);
      Rectangle crop = Rectangle.FromLTRB(
        inkBounds.Left - pad, inkBounds.Top - pad, inkBounds.Right + pad, inkBounds.Bottom + pad);

      using (Bitmap clean = new Bitmap(crop.Width, crop.Height, PixelFormat.Format32bppArgb))
      {
        int[] cleanPixels = new int[crop.Width * crop.Height];
        for (int y = 0; y < crop.Height; y++)
        {
          for (int x = 0; x < crop.Width; x++)
          {
            int sx = x + crop.Left;
            int sy = y + crop.Top;
            int shade = 255;
            if (sx >= 0 && sy >= 0 && sx < width && sy < height
              && (text[sy * width + sx] || TouchesText(text, width, height, sx, sy)))
            {
              // Keep the antialiased edge pixels next to text as partial shades.
              double strength = Math.Min(1.0, distance[sy * width + sx] / fullInk);
              shade = (int)Math.Round(255.0 * (1.0 - strength));
            }
            cleanPixels[y * crop.Width + x] = Color.FromArgb(shade, shade, shade).ToArgb();
          }
        }
        WritePixels(clean, cleanPixels);

        // Scale to the target letter height and thicken thin (SHX) strokes.
        double scale = Math.Max(1.0, Math.Min(8.0, TargetLetterHeight / TallestTextRow(text, width, height)));
        scale = Math.Min(scale, 4000.0 / Math.Max(crop.Width, crop.Height));
        int outWidth = Math.Max(1, (int)Math.Round(crop.Width * scale));
        int outHeight = Math.Max(1, (int)Math.Round(crop.Height * scale));
        double strokeWidth = MedianStrokeWidth(text, width, height) * scale;
        int thicken = (int)Math.Round((TargetLetterHeight * TargetStrokeRatio - strokeWidth) / 2.0);

        Bitmap result = new Bitmap(outWidth, outHeight, PixelFormat.Format32bppArgb);
        using (Graphics g = Graphics.FromImage(result))
        {
          g.Clear(Color.White);
          g.InterpolationMode = InterpolationMode.HighQualityBicubic;
          g.PixelOffsetMode = PixelOffsetMode.HighQuality;
          using (ImageAttributes attributes = new ImageAttributes())
          {
            attributes.SetWrapMode(WrapMode.TileFlipXY);
            g.DrawImage(
              clean,
              new Rectangle(0, 0, outWidth, outHeight),
              0, 0, crop.Width, crop.Height,
              GraphicsUnit.Pixel,
              attributes);
          }
        }
        if (thicken > 0)
        {
          int[] resultPixels = ReadPixels(result);
          WritePixels(result, Darken(resultPixels, outWidth, outHeight, thicken));
        }
        return result;
      }
    }

    // Height of the tallest band of rows holding text, so a stacked tag is
    // scaled by its letter height rather than its overall height.
    private static int TallestTextRow(bool[] text, int width, int height)
    {
      int tallest = 1;
      int run = 0;
      for (int y = 0; y < height; y++)
      {
        bool any = false;
        for (int x = 0; x < width && !any; x++)
        {
          any = text[y * width + x];
        }
        run = any ? run + 1 : 0;
        tallest = Math.Max(tallest, run);
      }
      return tallest;
    }

    // Median length of the horizontal and vertical ink runs. Most runs cross a
    // stroke rather than follow it, so this approximates stroke width.
    private static double MedianStrokeWidth(bool[] text, int width, int height)
    {
      List<int> runs = new List<int>();
      for (int y = 0; y < height; y++)
      {
        int run = 0;
        for (int x = 0; x <= width; x++)
        {
          if (x < width && text[y * width + x]) { run++; continue; }
          if (run > 0) runs.Add(run);
          run = 0;
        }
      }
      for (int x = 0; x < width; x++)
      {
        int run = 0;
        for (int y = 0; y <= height; y++)
        {
          if (y < height && text[y * width + x]) { run++; continue; }
          if (run > 0) runs.Add(run);
          run = 0;
        }
      }
      if (runs.Count == 0) return 1.0;
      runs.Sort();
      return runs[runs.Count / 2];
    }

    // Grows dark strokes by the given radius: each pixel takes the darkest
    // shade within a square around it.
    private static int[] Darken(int[] pixels, int width, int height, int radius)
    {
      int[] rows = new int[pixels.Length];
      for (int y = 0; y < height; y++)
      {
        for (int x = 0; x < width; x++)
        {
          int darkest = 255;
          for (int nx = Math.Max(0, x - radius); nx <= Math.Min(width - 1, x + radius); nx++)
          {
            darkest = Math.Min(darkest, pixels[y * width + nx] & 0xFF);
          }
          rows[y * width + x] = darkest;
        }
      }

      int[] result = new int[pixels.Length];
      for (int y = 0; y < height; y++)
      {
        for (int x = 0; x < width; x++)
        {
          int darkest = 255;
          for (int ny = Math.Max(0, y - radius); ny <= Math.Min(height - 1, y + radius); ny++)
          {
            darkest = Math.Min(darkest, rows[ny * width + x]);
          }
          result[y * width + x] = Color.FromArgb(darkest, darkest, darkest).ToArgb();
        }
      }
      return result;
    }

    private static bool[] KeepTextComponents(bool[] ink, int width, int height)
    {
      bool[] keep = new bool[ink.Length];
      bool[] seen = new bool[ink.Length];
      List<int> component = new List<int>();
      Stack<int> pending = new Stack<int>();

      for (int start = 0; start < ink.Length; start++)
      {
        if (!ink[start] || seen[start]) continue;

        component.Clear();
        seen[start] = true;
        pending.Push(start);
        int minX = width, minY = height, maxX = -1, maxY = -1;
        while (pending.Count > 0)
        {
          int index = pending.Pop();
          component.Add(index);
          int x = index % width;
          int y = index / width;
          minX = Math.Min(minX, x);
          maxX = Math.Max(maxX, x);
          minY = Math.Min(minY, y);
          maxY = Math.Max(maxY, y);

          for (int dy = -1; dy <= 1; dy++)
          {
            for (int dx = -1; dx <= 1; dx++)
            {
              int nx = x + dx;
              int ny = y + dy;
              if (nx < 0 || ny < 0 || nx >= width || ny >= height) continue;
              int next = ny * width + nx;
              if (!ink[next] || seen[next]) continue;
              seen[next] = true;
              pending.Push(next);
            }
          }
        }

        int boxWidth = maxX - minX + 1;
        int boxHeight = maxY - minY + 1;
        bool touchesEdge = minX == 0 || minY == 0 || maxX == width - 1 || maxY == height - 1;
        bool hollowOutline = boxWidth >= 0.7 * width && boxHeight >= 0.7 * height
          && component.Count < 0.15 * boxWidth * boxHeight;
        bool divider = boxWidth >= 0.5 * width && boxHeight * 8 <= boxWidth;
        if (touchesEdge || hollowOutline || divider) continue;

        foreach (int index in component)
        {
          keep[index] = true;
        }
      }
      return keep;
    }

    private static bool TouchesText(bool[] text, int width, int height, int x, int y)
    {
      for (int dy = -1; dy <= 1; dy++)
      {
        for (int dx = -1; dx <= 1; dx++)
        {
          int nx = x + dx;
          int ny = y + dy;
          if (nx >= 0 && ny >= 0 && nx < width && ny < height && text[ny * width + nx]) return true;
        }
      }
      return false;
    }

    private static bool Bounds(bool[] mask, int width, int height, out Rectangle bounds)
    {
      int minX = width, minY = height, maxX = -1, maxY = -1;
      for (int i = 0; i < mask.Length; i++)
      {
        if (!mask[i]) continue;
        int x = i % width;
        int y = i / width;
        minX = Math.Min(minX, x);
        maxX = Math.Max(maxX, x);
        minY = Math.Min(minY, y);
        maxY = Math.Max(maxY, y);
      }
      bounds = maxX < 0 ? Rectangle.Empty : Rectangle.FromLTRB(minX, minY, maxX + 1, maxY + 1);
      return maxX >= 0;
    }

    // The most common color, bucketed to 16 levels per channel.
    private static Color FindBackground(int[] pixels)
    {
      Dictionary<int, int> counts = new Dictionary<int, int>();
      foreach (int argb in pixels)
      {
        int bucket = Bucket(argb);
        counts.TryGetValue(bucket, out int count);
        counts[bucket] = count + 1;
      }

      int best = 0;
      int bestCount = -1;
      foreach (KeyValuePair<int, int> entry in counts)
      {
        if (entry.Value > bestCount)
        {
          best = entry.Key;
          bestCount = entry.Value;
        }
      }

      long r = 0, g = 0, b = 0;
      foreach (int argb in pixels)
      {
        if (Bucket(argb) != best) continue;
        Color c = Color.FromArgb(argb);
        r += c.R;
        g += c.G;
        b += c.B;
      }
      return Color.FromArgb((int)(r / bestCount), (int)(g / bestCount), (int)(b / bestCount));
    }

    private static int Bucket(int argb)
    {
      return ((argb >> 20) & 0xF) << 8 | ((argb >> 12) & 0xF) << 4 | ((argb >> 4) & 0xF);
    }

    private static double ColorDistance(Color a, Color b)
    {
      int dr = a.R - b.R;
      int dg = a.G - b.G;
      int db = a.B - b.B;
      return Math.Sqrt(dr * dr + dg * dg + db * db);
    }

    private static int[] ReadPixels(Bitmap bitmap)
    {
      Rectangle rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
      BitmapData data = bitmap.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
      try
      {
        int[] pixels = new int[bitmap.Width * bitmap.Height];
        for (int y = 0; y < bitmap.Height; y++)
        {
          Marshal.Copy(IntPtr.Add(data.Scan0, y * data.Stride), pixels, y * bitmap.Width, bitmap.Width);
        }
        return pixels;
      }
      finally
      {
        bitmap.UnlockBits(data);
      }
    }

    private static void WritePixels(Bitmap bitmap, int[] pixels)
    {
      Rectangle rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
      BitmapData data = bitmap.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
      try
      {
        for (int y = 0; y < bitmap.Height; y++)
        {
          Marshal.Copy(pixels, y * bitmap.Width, IntPtr.Add(data.Scan0, y * data.Stride), bitmap.Width);
        }
      }
      finally
      {
        bitmap.UnlockBits(data);
      }
    }

    // Returns the recognized lines, or null with an error message. Tags are
    // upper case, so the character set is limited to what tags use.
    private static List<string> RunTesseract(string imagePath, out string error)
    {
      string pluginFolder = Path.GetDirectoryName(typeof(EquipmentTagOcr).Assembly.Location);
      string dataFolder = Path.Combine(pluginFolder, "tessdata");
      if (!File.Exists(Path.Combine(dataFolder, "eng.traineddata")))
      {
        error = $"the OCR data file is missing from {dataFolder}";
        return null;
      }

      // The native engine DLLs ship in the plugin's x64 folder.
      Tesseract.TesseractEnviornment.CustomSearchPath = pluginFolder;
      using (Tesseract.TesseractEngine engine = new Tesseract.TesseractEngine(dataFolder, "eng", Tesseract.EngineMode.LstmOnly))
      using (Tesseract.Pix image = Tesseract.Pix.LoadFromFile(imagePath))
      {
        engine.SetVariable("tessedit_char_whitelist", "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-/.");
        engine.SetVariable("load_system_dawg", "0");
        engine.SetVariable("load_freq_dawg", "0");
        using (Tesseract.Page page = engine.Process(image, Tesseract.PageSegMode.SingleBlock))
        {
          error = null;
          return new List<string>(
            (page.GetText() ?? string.Empty).Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));
        }
      }
    }
  }
}
