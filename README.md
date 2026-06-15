# Contour Value Extractor

A desktop application that extracts contour data and scale values from scientific/heatmap-style images and converts them into structured Excel output.

---

## 📌 Overview

This tool processes an input image containing:
- A **contour plot**
- A **color scale (legend)**

It automatically:
1. Detects and separates the contour region and the scale region
2. Extracts the color gradient from the scale
3. Maps colors to numeric values (based on user-defined min/max range)
4. Samples a user-defined number of points from the contour
5. Assigns each point a value based on color similarity to the scale
6. Exports everything into an Excel file

---

## ⚙️ Features

- Automatic contour and scale detection using OpenCV
- Color-to-value mapping from extracted legend
- User-defined sampling of contour points
- Adjustable min/max value range
- Excel export with:
  - Scale table (RGB + values)
  - Extracted contour points (X, Y, RGB, value)
- Visual WPF interface with image preview
- Output folder generation per input image

---

## 🖼️ Input Requirements

The input image should contain:
- A clearly visible **contour plot**
- A visible **color scale/legend**
- Clean contrast between background and data

Best results are achieved with high-resolution scientific plots.

---

## 🧠 How It Works

1. **Preprocessing**
   - Converts image to grayscale
   - Applies thresholding and morphological filtering

2. **Region Detection**
   - Finds connected components
   - Separates:
     - Contour region
     - Scale (legend) region

3. **Scale Extraction**
   - Samples colors from the scale column
   - Maps them linearly between user-defined min/max values

4. **Contour Sampling**
   - Selects N points from valid contour pixels
   - Avoids background (white/black regions)

5. **Value Assignment**
   - Matches each point color to closest scale color

6. **Excel Export**
   - Writes results using Microsoft Excel Interop

---

## 📤 Output

For each input image, a folder is created:
