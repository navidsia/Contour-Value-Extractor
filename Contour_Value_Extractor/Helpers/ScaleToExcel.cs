using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using OpenCvSharp;
using Excel = Microsoft.Office.Interop.Excel;

namespace Contour_Value_Extractor.Helpers
{
    public static class ScaleToExcel
    {
        private struct ScaleEntry
        {
            public byte R;
            public byte G;
            public byte B;
            public double Temp;

            public ScaleEntry(byte r, byte g, byte b, double temp)
            {
                R = r;
                G = g;
                B = b;
                Temp = temp;
            }
        }

        private struct XY
        {
            public int x;
            public int y;
            public XY(int X, int Y) { x = X; y = Y; }
        }

        private static bool IsWhiteOrBlack(byte r, byte g, byte b)
        {
            int avg = (r + g + b) / 3;
            return (avg > 200) || (avg < 35);
        }

        private static int ToExcelColorInt(byte r, byte g, byte b)
        {
            return r + (g << 8) + (b << 16);
        }

        public static string ExportScaleAndPointsToExcel(
            string scaleImagePath,
            string contourImagePath,
            double tMin,
            double tMax,
            int requestedPoints)
        {
            if (!System.IO.File.Exists(scaleImagePath))
                throw new System.IO.FileNotFoundException("Scale image not found", scaleImagePath);

            if (!System.IO.File.Exists(contourImagePath))
                throw new System.IO.FileNotFoundException("Contour image not found", contourImagePath);

            if (requestedPoints <= 0)
                requestedPoints = 1;


            List<ScaleEntry> scaleTable = BuildScaleTable(scaleImagePath, tMin, tMax);


            int W, H;
            byte[,] validMask;
            int validCount;
            BuildContourMask(contourImagePath, out W, out H, out validMask, out validCount);

            if (validCount <= 0)
                throw new Exception("Contour has no valid (non-white/black) pixels.");

            List<Tuple<int, int>> points = ChoosePoints(validMask, W, H, validCount, requestedPoints);


            List<Tuple<int, int, byte, byte, byte, double>> pointResults =
                ExtractPointTemps(contourImagePath, points, scaleTable);


            string outDir = System.IO.Path.GetDirectoryName(scaleImagePath);

            string contourBase = System.IO.Path.GetFileNameWithoutExtension(contourImagePath);
            string originalBase = contourBase.EndsWith("_contour", StringComparison.OrdinalIgnoreCase)
                ? contourBase.Substring(0, contourBase.Length - "_contour".Length)
                : contourBase;

            string outXlsx = System.IO.Path.Combine(outDir, originalBase + "_scale_and_points.xlsx");

            if (System.IO.File.Exists(outXlsx))
                System.IO.File.Delete(outXlsx);

            WriteExcel(outXlsx, scaleTable, pointResults);

            return outXlsx;
        }

        private static List<ScaleEntry> BuildScaleTable(string scaleImagePath, double tMin, double tMax)
        {
            using (Mat scaleImg = Cv2.ImRead(scaleImagePath, ImreadModes.Color))
            {
                if (scaleImg.Empty())
                    throw new Exception("Could not read scale image.");

                int h = scaleImg.Rows;
                int midX = scaleImg.Cols / 2;

                List<ScaleEntry> raw = new List<ScaleEntry>();

                for (int y = h - 1; y >= 0; y--)
                {
                    Vec3b bgr = scaleImg.At<Vec3b>(y, midX);

                    byte B = bgr.Item0;
                    byte G = bgr.Item1;
                    byte R = bgr.Item2;

                    if (IsWhiteOrBlack(R, G, B))
                        continue;

                    raw.Add(new ScaleEntry(R, G, B, 0.0));
                }

                if (raw.Count < 2)
                    throw new Exception("Not enough valid pixels found in scale middle column.");

                int n = raw.Count;
                List<ScaleEntry> scaleTable = new List<ScaleEntry>(n);

                for (int i = 0; i < n; i++)
                {
                    double frac = (double)i / (double)(n - 1);

                    double temp = tMin + frac * (tMax - tMin);

                    ScaleEntry e = raw[i];
                    scaleTable.Add(new ScaleEntry(e.R, e.G, e.B, temp));
                }

                return scaleTable;
            }
        }

        private static void BuildContourMask(string contourImagePath, out int W, out int H,
            out byte[,] validMask, out int validCount)
        {
            using (Mat contourImg = Cv2.ImRead(contourImagePath, ImreadModes.Color))
            {
                if (contourImg.Empty())
                    throw new Exception("Could not read contour image.");

                H = contourImg.Rows;
                W = contourImg.Cols;

                validMask = new byte[H, W];
                validCount = 0;

                for (int y = 0; y < H; y++)
                {
                    for (int x = 0; x < W; x++)
                    {
                        Vec3b bgr = contourImg.At<Vec3b>(y, x);

                        byte B = bgr.Item0;
                        byte G = bgr.Item1;
                        byte R = bgr.Item2;

                        if (!IsWhiteOrBlack(R, G, B))
                        {
                            validMask[y, x] = 1;
                            validCount++;
                        }
                    }
                }
            }
        }

        private static List<Tuple<int, int>> ChoosePoints(byte[,] validMask, int W, int H, int validCount, int requestedPoints)
        {
            List<Tuple<int, int>> points = new List<Tuple<int, int>>();

            if (requestedPoints >= validCount)
            {
                for (int y = 0; y < H; y++)
                    for (int x = 0; x < W; x++)
                        if (validMask[y, x] == 1)
                            points.Add(Tuple.Create(x, y));

                return points;
            }

            int stride = ChooseBestStrideForValidCount(validCount, requestedPoints);

            for (int y = 0; y < H; y += stride)
            {
                for (int x = 0; x < W; x += stride)
                {
                    XY? p = FindNearestValid(validMask, W, H, x, y, Math.Max(2, stride / 2));
                    if (p != null)
                        points.Add(Tuple.Create(p.Value.x, p.Value.y));
                }
            }

            points = Dedup(points);

            if (points.Count > requestedPoints)
                points = points.GetRange(0, requestedPoints);

            return points;
        }

        private static List<Tuple<int, int, byte, byte, byte, double>> ExtractPointTemps(
            string contourImagePath,
            List<Tuple<int, int>> points,
            List<ScaleEntry> scaleTable)
        {
            List<Tuple<int, int, byte, byte, byte, double>> outList =
                new List<Tuple<int, int, byte, byte, byte, double>>(points.Count);

            using (Mat contourImg = Cv2.ImRead(contourImagePath, ImreadModes.Color))
            {
                if (contourImg.Empty())
                    throw new Exception("Could not read contour image (second read).");

                for (int i = 0; i < points.Count; i++)
                {
                    int x = points[i].Item1;
                    int y = points[i].Item2;

                    Vec3b bgr = contourImg.At<Vec3b>(y, x);
                    byte B = bgr.Item0;
                    byte G = bgr.Item1;
                    byte R = bgr.Item2;

                    if (IsWhiteOrBlack(R, G, B))
                        continue;

                    double temp = FindClosestTemp(scaleTable, R, G, B);
                    outList.Add(Tuple.Create(x, y, R, G, B, temp));
                }
            }

            return outList;
        }

        private static double FindClosestTemp(List<ScaleEntry> scale, byte r, byte g, byte b)
        {
            int best = int.MaxValue;
            double bestTemp = scale[0].Temp;

            for (int i = 0; i < scale.Count; i++)
            {
                int dr = r - scale[i].R;
                int dg = g - scale[i].G;
                int db = b - scale[i].B;
                int d = dr * dr + dg * dg + db * db;

                if (d < best)
                {
                    best = d;
                    bestTemp = scale[i].Temp;
                    if (best == 0) break;
                }
            }

            return bestTemp;
        }

        private static void WriteExcel(
            string outXlsx,
            List<ScaleEntry> scaleTable,
            List<Tuple<int, int, byte, byte, byte, double>> pointResults)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet wsScale = null;
            Excel.Worksheet wsPoints = null;

            try
            {
                app = new Excel.Application();
                app.Visible = false;
                app.DisplayAlerts = false;

                wb = app.Workbooks.Add();

                wsScale = (Excel.Worksheet)wb.Worksheets[1];
                wsScale.Name = "Scale";

                wsScale.Cells[1, 1] = "R";
                wsScale.Cells[1, 2] = "G";
                wsScale.Cells[1, 3] = "B";
                wsScale.Cells[1, 4] = "Temp";

                for (int i = 0; i < scaleTable.Count; i++)
                {
                    int row = i + 2;
                    ScaleEntry e = scaleTable[i];

                    wsScale.Cells[row, 1] = e.R;
                    wsScale.Cells[row, 2] = e.G;
                    wsScale.Cells[row, 3] = e.B;
                    wsScale.Cells[row, 4] = e.Temp;

                    Excel.Range cell = (Excel.Range)wsScale.Cells[row, 4];
                    cell.Interior.Color = ToExcelColorInt(e.R, e.G, e.B);
                }

                wsScale.Columns.AutoFit();

                wsPoints = (Excel.Worksheet)wb.Worksheets.Add(After: wsScale);
                wsPoints.Name = "Points";

                wsPoints.Cells[1, 1] = "X";
                wsPoints.Cells[1, 2] = "Y";
                wsPoints.Cells[1, 3] = "R";
                wsPoints.Cells[1, 4] = "G";
                wsPoints.Cells[1, 5] = "B";
                wsPoints.Cells[1, 6] = "Temp";

                for (int i = 0; i < pointResults.Count; i++)
                {
                    int row = i + 2;
                    var p = pointResults[i];

                    wsPoints.Cells[row, 1] = p.Item1;
                    wsPoints.Cells[row, 2] = p.Item2;
                    wsPoints.Cells[row, 3] = p.Item3;
                    wsPoints.Cells[row, 4] = p.Item4;
                    wsPoints.Cells[row, 5] = p.Item5;
                    wsPoints.Cells[row, 6] = p.Item6;

                    Excel.Range cell = (Excel.Range)wsPoints.Cells[row, 6];
                    cell.Interior.Color = ToExcelColorInt(p.Item3, p.Item4, p.Item5);
                }

                wsPoints.Columns.AutoFit();

                wb.SaveAs(outXlsx);
                wb.Close(false);
                app.Quit();
            }
            finally
            {
                if (wsPoints != null) Marshal.ReleaseComObject(wsPoints);
                if (wsScale != null) Marshal.ReleaseComObject(wsScale);
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (app != null) Marshal.ReleaseComObject(app);

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static int ChooseBestStrideForValidCount(int validCount, int requested)
        {
            double ideal = Math.Sqrt((double)validCount / (double)requested);
            int baseS = Math.Max(1, (int)Math.Round(ideal));

            int bestS = baseS;
            int bestDiff = int.MaxValue;

            for (int s = Math.Max(1, baseS - 10); s <= baseS + 10; s++)
            {
                int approx = Math.Max(1, validCount / (s * s));
                int diff = Math.Abs(approx - requested);
                if (diff < bestDiff)
                {
                    bestDiff = diff;
                    bestS = s;
                }
            }

            return bestS;
        }

        private static XY? FindNearestValid(byte[,] mask, int W, int H, int x0, int y0, int radius)
        {
            if (x0 >= 0 && x0 < W && y0 >= 0 && y0 < H && mask[y0, x0] == 1)
                return new XY(x0, y0);

            int rMax = Math.Max(1, radius);

            for (int r = 1; r <= rMax; r++)
            {
                int yMin = Math.Max(0, y0 - r);
                int yMax = Math.Min(H - 1, y0 + r);
                int xMin = Math.Max(0, x0 - r);
                int xMax = Math.Min(W - 1, x0 + r);

                for (int x = xMin; x <= xMax; x++)
                {
                    if (mask[yMin, x] == 1) return new XY(x, yMin);
                    if (mask[yMax, x] == 1) return new XY(x, yMax);
                }

                for (int y = yMin; y <= yMax; y++)
                {
                    if (mask[y, xMin] == 1) return new XY(xMin, y);
                    if (mask[y, xMax] == 1) return new XY(xMax, y);
                }
            }

            return null;
        }

        private static List<Tuple<int, int>> Dedup(List<Tuple<int, int>> pts)
        {
            HashSet<long> seen = new HashSet<long>();
            List<Tuple<int, int>> outPts = new List<Tuple<int, int>>();

            for (int i = 0; i < pts.Count; i++)
            {
                int x = pts[i].Item1;
                int y = pts[i].Item2;
                long key = (((long)y) << 32) ^ (uint)x;

                if (seen.Add(key))
                    outPts.Add(pts[i]);
            }

            return outPts;
        }
    }
}
