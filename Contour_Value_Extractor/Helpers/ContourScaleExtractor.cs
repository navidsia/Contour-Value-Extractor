// Helpers/ContourScaleExtractor.cs  (FULL FILE)

using System;
using System.Collections.Generic;
using System.Linq;
using OpenCvSharp;

namespace Contour_Value_Extractor.Helpers
{
    public static class ContourScaleExtractor
    {
        private class Candidate
        {
            public int Label;
            public Rect Rect;
            public int Area;
            public double Cx;
            public double Cy;
            public double Aspect;
        }

        public class ExtractionResult
        {
            public Rect ContourRect;
            public Rect ScaleRect;
            public string ContourPath;
            public string ScalePath;
        }

        /// <summary>
        /// Extracts contour (main colored field) and scale (colorbar/legend) from a CFD screenshot.
        /// Saves two crops with minimal white margins (tight crop).
        /// Assumes white background; scale is typically on the left.
        /// </summary>
        public static ExtractionResult ExtractContourAndScale(
            string imagePath,
            string outContourPath,
            string outScalePath,
            int whiteThresh = 245,
            int minArea = 5000,
            int tightPad = 2)
        {
            if (string.IsNullOrWhiteSpace(imagePath))
                throw new ArgumentException("imagePath is null or empty");

            using (Mat img = Cv2.ImRead(imagePath, ImreadModes.Color))
            {
                if (img.Empty())
                    throw new ArgumentException("Could not read image: " + imagePath);

                int W = img.Cols;
                int H = img.Rows;

                using (Mat gray = new Mat())
                using (Mat mask = new Mat())
                using (Mat kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(5, 5)))
                {
                    // Build mask of "non-white" pixels.
                    // Anything darker than whiteThresh is considered part of content.
                    Cv2.CvtColor(img, gray, ColorConversionCodes.BGR2GRAY);
                    Cv2.Threshold(gray, mask, whiteThresh - 1, 255, ThresholdTypes.BinaryInv);

                    // Mild cleanup (avoid destroying thin edges)
                    Cv2.MorphologyEx(mask, mask, MorphTypes.Open, kernel, new Point(-1, -1), 1);
                    Cv2.MorphologyEx(mask, mask, MorphTypes.Close, kernel, new Point(-1, -1), 1);

                    using (Mat labels = new Mat())
                    using (Mat stats = new Mat())
                    using (Mat centroids = new Mat())
                    {
                        // Connected components (labels, stats, centroids are output mats)
                        Cv2.ConnectedComponentsWithStats(mask, labels, stats, centroids);

                        List<Candidate> candidates = new List<Candidate>();

                        // stats rows: one per label (0 = background)
                        for (int i = 1; i < stats.Rows; i++)
                        {
                            int x = stats.Get<int>(i, (int)ConnectedComponentsTypes.Left);
                            int y = stats.Get<int>(i, (int)ConnectedComponentsTypes.Top);
                            int w = stats.Get<int>(i, (int)ConnectedComponentsTypes.Width);
                            int h = stats.Get<int>(i, (int)ConnectedComponentsTypes.Height);
                            int area = stats.Get<int>(i, (int)ConnectedComponentsTypes.Area);

                            if (area < minArea)
                                continue;

                            double cx = centroids.Get<double>(i, 0);
                            double cy = centroids.Get<double>(i, 1);
                            double aspect = (w > 0) ? (double)h / (double)w : 0.0;

                            candidates.Add(new Candidate
                            {
                                Label = i,
                                Rect = new Rect(x, y, w, h),
                                Area = area,
                                Cx = cx,
                                Cy = cy,
                                Aspect = aspect
                            });
                        }

                        if (candidates.Count < 2)
                            throw new Exception("Could not find both contour and scale regions.");

                        // 1) Pick scale: left-most + tall/narrow
                        Candidate scale = candidates
                            .OrderByDescending(c => ScaleScore(c, W))
                            .First();

                        // 2) Pick contour: largest remaining (prefer to the right of scale)
                        List<Candidate> remaining = candidates.Where(c => c.Label != scale.Label).ToList();
                        double rightThreshold = scale.Cx + scale.Rect.Width * 0.5;

                        List<Candidate> remainingRight = remaining.Where(c => c.Cx > rightThreshold).ToList();
                        List<Candidate> pool = (remainingRight.Count > 0) ? remainingRight : remaining;

                        Candidate contour = pool.OrderByDescending(c => c.Area).First();

                        // 3) Tight-crop to remove white margins
                        Rect scaleRect = TightenRectToMask(mask, scale.Rect, W, H, tightPad);
                        Rect contourRect = TightenRectToMask(mask, contour.Rect, W, H, tightPad);

                        // Save crops (overwrite if exist)
                        using (Mat scaleCrop = new Mat(img, scaleRect))
                        using (Mat contourCrop = new Mat(img, contourRect))
                        {
                            Cv2.ImWrite(outScalePath, scaleCrop);
                            Cv2.ImWrite(outContourPath, contourCrop);
                        }

                        return new ExtractionResult
                        {
                            ScaleRect = scaleRect,
                            ContourRect = contourRect,
                            ScalePath = outScalePath,
                            ContourPath = outContourPath
                        };
                    }
                }
            }
        }

        private static double ScaleScore(Candidate c, int imageWidth)
        {
            // Higher score => more likely to be the colorbar
            double leftness = 1.0 - (c.Cx / Math.Max(imageWidth, 1));
            double tallness = Math.Min(c.Aspect, 20.0) / 20.0;
            double narrowness = 1.0 - Math.Min((double)c.Rect.Width / Math.Max(imageWidth, 1), 1.0);

            return 0.55 * leftness + 0.30 * tallness + 0.15 * narrowness;
        }

        /// <summary>
        /// Tightens a rough rectangle to the smallest bounding box of non-zero mask pixels inside it.
        /// This removes white borders around the crop.
        /// </summary>
        private static Rect TightenRectToMask(Mat mask, Rect rough, int imgW, int imgH, int pad)
        {
            Rect r = ClampRect(rough, imgW, imgH);
            if (r.Width <= 2 || r.Height <= 2)
                return r;

            int left = r.X + r.Width - 1;
            int right = r.X;
            int top = r.Y + r.Height - 1;
            int bottom = r.Y;

            bool found = false;

            for (int y = r.Y; y < r.Y + r.Height; y++)
            {
                for (int x = r.X; x < r.X + r.Width; x++)
                {
                    if (mask.At<byte>(y, x) != 0)
                    {
                        found = true;
                        if (x < left) left = x;
                        if (x > right) right = x;
                        if (y < top) top = y;
                        if (y > bottom) bottom = y;
                    }
                }
            }

            if (!found)
                return r;

            left = Math.Max(0, left - pad);
            right = Math.Min(imgW - 1, right + pad);
            top = Math.Max(0, top - pad);
            bottom = Math.Min(imgH - 1, bottom + pad);

            int w = right - left + 1;
            int h = bottom - top + 1;

            if (w < 1) w = 1;
            if (h < 1) h = 1;

            return new Rect(left, top, w, h);
        }

        private static Rect ClampRect(Rect r, int imgW, int imgH)
        {
            int x = Math.Max(0, r.X);
            int y = Math.Max(0, r.Y);
            int w = Math.Min(r.Width, imgW - x);
            int h = Math.Min(r.Height, imgH - y);

            if (w < 1) w = 1;
            if (h < 1) h = 1;

            return new Rect(x, y, w, h);
        }
    }
}
