/**
 * Windows system OCR helper for pdf2docx (see src/ocr.ts) — the Windows
 * counterpart of vision-ocr.swift, speaking the exact same protocol:
 *
 * Reads a PNG from stdin, writes recognized lines as JSON to stdout:
 *   {"paper":0.93,"lines":[{"t":"text","c":1.0,"b":[x0,y0,x1,y1],"chars":[{"t":"a","b":[...]}]}]}
 * Boxes are normalized 0-1, origin BOTTOM-left, y up (PDF page space).
 * "paper" is the near-white pixel share (photo-vs-document signal).
 *
 * argv[0] (optional): comma-separated BCP-47 language hints, e.g.
 * "zh-Hans,en-US" — tried in order; without it the user-profile languages
 * decide. Windows.Media.Ocr reports NO confidence values, so every line
 * carries c=1.0 — the TS policy layer still applies the paper-share, text
 * coverage, and char-count gates.
 *
 * Exit codes: 0 ok; 2 cannot decode input; 3 recognition failed;
 * 4 NO OCR language available on this machine (server SKUs / stripped
 * installs) — callers treat any non-zero exit as "no engine" and keep the
 * bitmap fallback.
 *
 * Build (no SDK on user machines needed — packaged installers ship the exe;
 * see build-win.mjs, run automatically by electron-builder's preflight):
 *   %WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe /nologo /o
 *     /r:<Windows Kits>\UnionMetadata\<ver>\Windows.winmd
 *     /r:<GAC>\System.Runtime.WindowsRuntime.dll
 *     /out:win-ocr.exe win-ocr.cs
 *
 * Kept to C# 5 (the in-box csc) on purpose: no string interpolation, no ?. —
 * portability over elegance.
 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using Windows.Foundation;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;

static class WinOcr
{
    static int Main(string[] args)
    {
        try
        {
            byte[] png;
            using (var stdin = Console.OpenStandardInput())
            using (var mem = new MemoryStream())
            {
                stdin.CopyTo(mem);
                png = mem.ToArray();
            }
            if (png.Length == 0)
            {
                Console.Error.WriteLine("empty input");
                return 2;
            }

            SoftwareBitmap bitmap;
            try
            {
                bitmap = Decode(png);
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("cannot decode input image: " + e.Message);
                return 2;
            }

            var engine = CreateEngine(args.Length > 0 ? args[0] : null);
            if (engine == null)
            {
                Console.Error.WriteLine("no OCR recognizer language available");
                return 4;
            }

            // the engine caps input dimensions; downscale proportionally when
            // a render exceeds it (boxes stay normalized, so callers never see it)
            var maxDim = (int)OcrEngine.MaxImageDimension;
            if (bitmap.PixelWidth > maxDim || bitmap.PixelHeight > maxDim)
            {
                bitmap = Downscale(bitmap, maxDim);
            }

            OcrResult result;
            try
            {
                result = engine.RecognizeAsync(bitmap).AsTask().Result;
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("recognition failed: " + e.Message);
                return 3;
            }

            double paper = PaperShare(bitmap);
            var json = new StringBuilder();
            json.Append("{\"paper\":").Append(Num(paper)).Append(",\"lines\":[");
            bool firstLine = true;
            foreach (var line in result.Lines)
            {
                var entry = LineJson(line, bitmap.PixelWidth, bitmap.PixelHeight);
                if (entry == null) continue;
                if (!firstLine) json.Append(',');
                firstLine = false;
                json.Append(entry);
            }
            json.Append("]}");
            // BOM-less UTF-8 straight to the stream: Console.OutputEncoding =
            // UTF8 emits a BOM on .NET Framework and JSON parsers choke on it
            byte[] payload = new UTF8Encoding(false).GetBytes(json.ToString());
            using (var stdout = Console.OpenStandardOutput())
            {
                stdout.Write(payload, 0, payload.Length);
            }
            return 0;
        }
        catch (Exception e)
        {
            Console.Error.WriteLine("unexpected failure: " + e);
            return 3;
        }
    }

    static SoftwareBitmap Decode(byte[] png)
    {
        using (var stream = new InMemoryRandomAccessStream())
        {
            stream.WriteAsync(png.AsBuffer()).AsTask().Wait();
            stream.Seek(0);
            var decoder = BitmapDecoder.CreateAsync(stream).AsTask().Result;
            var frame = decoder
                .GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied)
                .AsTask()
                .Result;
            return frame;
        }
    }

    static OcrEngine CreateEngine(string hints)
    {
        if (hints != null && hints.Length > 0)
        {
            foreach (var tag in hints.Split(','))
            {
                try
                {
                    var engine = OcrEngine.TryCreateFromLanguage(new Language(tag.Trim()));
                    if (engine != null) return engine;
                }
                catch (Exception)
                {
                    // malformed tag — try the next hint
                }
            }
        }
        return OcrEngine.TryCreateFromUserProfileLanguages();
    }

    static SoftwareBitmap Downscale(SoftwareBitmap source, int maxDim)
    {
        double scale = Math.Min(
            (double)maxDim / source.PixelWidth,
            (double)maxDim / source.PixelHeight
        );
        int w = Math.Max(1, (int)(source.PixelWidth * scale));
        int h = Math.Max(1, (int)(source.PixelHeight * scale));
        using (var stream = new InMemoryRandomAccessStream())
        {
            var encoder = BitmapEncoder.CreateAsync(BitmapEncoder.PngEncoderId, stream)
                .AsTask()
                .Result;
            encoder.SetSoftwareBitmap(source);
            encoder.BitmapTransform.ScaledWidth = (uint)w;
            encoder.BitmapTransform.ScaledHeight = (uint)h;
            encoder.BitmapTransform.InterpolationMode = BitmapInterpolationMode.Fant;
            encoder.FlushAsync().AsTask().Wait();
            stream.Seek(0);
            var decoder = BitmapDecoder.CreateAsync(stream).AsTask().Result;
            return decoder
                .GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied)
                .AsTask()
                .Result;
        }
    }

    /** near-white pixel share over strided samples (Bgra8: B,G,R,A) */
    static double PaperShare(SoftwareBitmap bitmap)
    {
        int w = bitmap.PixelWidth;
        int h = bitmap.PixelHeight;
        var buffer = new Windows.Storage.Streams.Buffer((uint)(w * h * 4));
        bitmap.CopyToBuffer(buffer);
        byte[] px = buffer.ToArray();
        int total = w * h;
        int stride = Math.Max(1, total / 200000);
        int paper = 0;
        int sampled = 0;
        for (int i = 0; i < total; i += stride)
        {
            int o = i * 4;
            byte b = px[o];
            byte g = px[o + 1];
            byte r = px[o + 2];
            byte min = Math.Min(b, Math.Min(g, r));
            if (min >= 200) paper++;
            sampled++;
        }
        return sampled > 0 ? (double)paper / sampled : 1.0;
    }

    /**
     * One line as a JSON object, or null for empty lines. Word rects become
     * per-char boxes (chars of a word split its rect evenly) so the TS layer
     * can anchor word segments; whitespace in the line text gets zero boxes
     * (the policy layer spans word gaps itself). When the line text cannot
     * be aligned with its words, chars are omitted and the TS layer falls
     * back to advance-weight distribution over the line box.
     */
    static string LineJson(OcrLine line, int pixelW, int pixelH)
    {
        string text = line.Text;
        if (text == null || text.Trim().Length == 0) return null;

        double x0 = double.MaxValue,
            y0 = double.MaxValue,
            x1 = double.MinValue,
            y1 = double.MinValue;
        foreach (var word in line.Words)
        {
            Rect r = word.BoundingRect;
            if (r.X < x0) x0 = r.X;
            if (r.Y < y0) y0 = r.Y;
            if (r.X + r.Width > x1) x1 = r.X + r.Width;
            if (r.Y + r.Height > y1) y1 = r.Y + r.Height;
        }
        if (x1 <= x0 || y1 <= y0) return null;

        var chars = AlignChars(line, text, pixelW, pixelH);

        var json = new StringBuilder();
        json.Append("{\"t\":").Append(Quote(text));
        // Windows.Media.Ocr exposes no confidence — see file header
        json.Append(",\"c\":1.0");
        json.Append(",\"b\":").Append(BoxJson(x0, y0, x1, y1, pixelW, pixelH));
        if (chars != null)
        {
            json.Append(",\"chars\":[");
            for (int i = 0; i < chars.Count; i++)
            {
                if (i > 0) json.Append(',');
                json.Append(chars[i]);
            }
            json.Append(']');
        }
        json.Append('}');
        return json.ToString();
    }

    static List<string> AlignChars(OcrLine line, string text, int pixelW, int pixelH)
    {
        var chars = new List<string>();
        int cursor = 0;
        foreach (var word in line.Words)
        {
            string wt = word.Text;
            if (wt == null || wt.Length == 0) continue;
            int at = text.IndexOf(wt, cursor, StringComparison.Ordinal);
            if (at < 0) return null; // line text and words disagree — let TS distribute
            // whitespace (or any separator) between words: zero boxes
            for (int i = cursor; i < at; i++)
            {
                chars.Add("{\"t\":" + Quote(text.Substring(i, 1)) + ",\"b\":[0,0,0,0]}");
            }
            Rect r = word.BoundingRect;
            // TextElementEnumerator-free split: per UTF-16 unit is fine here —
            // surrogate pairs are rare in OCR output and misalignment just
            // triggers the TS fallback via the glyph-count check
            int n = wt.Length;
            for (int i = 0; i < n; i++)
            {
                double cx0 = r.X + (r.Width * i) / n;
                double cx1 = r.X + (r.Width * (i + 1)) / n;
                chars.Add(
                    "{\"t\":"
                        + Quote(wt.Substring(i, 1))
                        + ",\"b\":"
                        + BoxJson(cx0, r.Y, cx1, r.Y + r.Height, pixelW, pixelH)
                        + "}"
                );
            }
            cursor = at + wt.Length;
        }
        // trailing separators
        for (int i = cursor; i < text.Length; i++)
        {
            chars.Add("{\"t\":" + Quote(text.Substring(i, 1)) + ",\"b\":[0,0,0,0]}");
        }
        return chars;
    }

    /** pixel rect (origin top-left) -> normalized PDF-space box (origin bottom-left) */
    static string BoxJson(double px0, double py0, double px1, double py1, int w, int h)
    {
        double nx0 = px0 / w;
        double nx1 = px1 / w;
        double ny0 = 1.0 - py1 / h;
        double ny1 = 1.0 - py0 / h;
        return "[" + Num(nx0) + "," + Num(ny0) + "," + Num(nx1) + "," + Num(ny1) + "]";
    }

    static string Num(double v)
    {
        return v.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture);
    }

    static string Quote(string s)
    {
        var sb = new StringBuilder("\"");
        foreach (char c in s)
        {
            if (c == '"') sb.Append("\\\"");
            else if (c == '\\') sb.Append("\\\\");
            else if (c < 0x20)
                sb.Append("\\u").Append(((int)c).ToString("x4"));
            else sb.Append(c);
        }
        sb.Append('"');
        return sb.ToString();
    }
}
