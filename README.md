# Contour Value Extractor

**Extract numerical values from color-coded contour images.** Contour Value Extractor is a Windows desktop tool that processes scientific/engineering screenshots containing a colored field and a vertical color scale. It locates the contour region and legend, estimates values from pixel colors, samples points across the field, and writes the results to Excel.

Built with **C#**, **WPF**, **.NET Framework 4.8**, **OpenCvSharp**, and **Microsoft Excel Interop**.

## What it does

Given a compatible image containing a contour/heatmap and a vertical color bar, the application:

1. Detects the main colored region and its legend using image processing.
2. Saves tightly cropped versions of both regions.
3. Reads the colors along the legend and associates them with a user-provided numerical range.
4. Samples valid contour pixels and matches their colors to the closest legend colors.
5. Exports the reconstructed scale and the sampled points to an `.xlsx` workbook.

This is useful for recovering approximate data from **CFD visualizations, simulation screenshots, temperature contours, and similar color-mapped scientific plots** when the original numerical data is unavailable.

## Requirements

- Windows
- [.NET Framework 4.8](https://dotnet.microsoft.com/en-us/download/dotnet-framework/net48)
- **Microsoft Excel installed locally** (the application creates workbooks through Excel COM automation)

For development, use Visual Studio with .NET desktop development and the .NET Framework 4.8 targeting pack. The project references these NuGet packages, among others:

- `OpenCvSharp4` and `OpenCvSharp4.runtime.win` (`4.11.0.20250507`)
- `Microsoft.Office.Interop.Excel` (`16.0.18925.20022`)

Visual Studio can restore the remaining dependencies from `packages.config`.

## Build and run

1. Clone the repository:

   ```bash
   git clone https://github.com/navidsia/Contour-Value-Extractor.git
   ```

2. Open `Contour_Value_Extractor.sln` in Visual Studio on Windows.
3. Restore the NuGet packages.
4. Build and run the WPF project. Make sure the OpenCvSharp Windows native runtime is copied to the output directory, as configured by the NuGet package.

## How to use

1. Prepare an image containing a clear **vertical color scale**, usually to the **left** of the main contour, on a mostly white background.
2. Enter **Minimum value** and **Maximum value** for the scale. These can represent temperature or another physical quantity; the program does not impose a unit.
3. Enter a positive **Number of points**. The default is `200`.
4. Click **Choose Image…** and select a `.png`, `.jpg`, `.jpeg`, or `.bmp` image.
5. Review the preview and click **Extract Data**.
6. When processing completes, click **Open Output Folder** to inspect the crops and spreadsheet.

**Example:** For a temperature legend ranging from 273 K at the bottom to 2573 K at the top, enter `273` and `2573`. Keep the source image and legend together in the input screenshot; the application finds and crops them automatically.

### Output files

For an input file named `simulation.png`, the tool creates a folder next to the source image:

```text
simulation.png
extracted_simulation/
  simulation_contour.png
  simulation_scale.png
  simulation_scale_and_points.xlsx
```

The `.xlsx` workbook contains two worksheets:

| Sheet | Columns | Description |
| --- | --- | --- |
| `Scale` | `R`, `G`, `B`, `Temp` | Legend colors and their corresponding estimated values |
| `Points` | `X`, `Y`, `R`, `G`, `B`, `Temp` | Sampled contour-pixel coordinates, RGB values, and estimated values |

The `Temp` cells are filled with the corresponding sampled color for easier inspection. **X and Y are pixel coordinates within the cropped contour image**, with `(0, 0)` at its upper-left corner; they are not physical coordinates or coordinates of the original full screenshot.

The field is named `Temp` in the current code, but values can represent another measurement if its legend provides a numerical range.

**Important:** Running extraction again on the same input filename deletes and regenerates the corresponding output images and workbook. Back up previous results if you need to keep them.

## How the extraction works

### 1. Locate the contour and scale

The `ContourScaleExtractor` class:

- Converts the image to grayscale and builds a mask of non-white content.
- Applies morphological opening and closing to reduce noise.
- Finds connected components and ignores regions smaller than the configured area threshold (currently `5000` pixels from the UI).
- Scores candidate color bars by their **leftmost position, tall/narrow shape**, and width.
- Selects the largest suitable remaining component as the contour, preferring components to the right of the color bar.
- Saves tight crops of both selected regions, with a small margin.

The default image-processing settings from the user interface are a whiteness threshold of `245`, a minimum component area of `5000`, and crop padding of `2` pixels.

### 2. Reconstruct the legend

The `ScaleToExcel` class reads the **middle column** of the cropped color bar from **bottom to top**, skipping pixels considered too dark or too light. It assigns an evenly spaced numerical value between the supplied minimum and maximum to each retained color:

```text
value(i) = minValue + i / (N - 1) * (maxValue - minValue)
```

The current implementation therefore assumes a **vertical, continuous color scale with the minimum at the bottom and maximum at the top**. It does not perform OCR on printed legend numbers or infer the legend range automatically.

### 3. Sample the contour

The program ignores pixels classified as near-black or near-white, then samples the remaining contour pixels across a spatial grid. If a grid location is invalid, it searches nearby for a valid pixel. Duplicate locations are removed, and the result is capped at the requested point count.

The point count is therefore a **target/upper limit**, not a guarantee that exactly that many points will be returned. If the requested count is at least the total number of valid pixels, all valid pixels are used.

### 4. Convert sampled colors to values

For each sampled pixel, the application finds the closest entry in the extracted legend using **squared RGB color distance**:

```text
distance = (R - Rs)^2 + (G - Gs)^2 + (B - Bs)^2
```

The corresponding legend value is recorded as the pixel's estimated value. This is a color lookup method, not a reconstruction of the original simulation field, and its accuracy depends on the source legend and image quality.

## Input recommendations and limitations

For better results:

- Use a lossless **PNG** screenshot when possible; JPEG compression can shift colors.
- Keep the color bar and main contour clearly separated against a **white background**.
- Prefer a tall, narrow legend **left of the contour**, with a continuous gradient.
- Crop away unrelated labels, icons, annotations, or charts if they interfere with region detection.
- Supply the correct numerical minimum and maximum; the application cannot read them from text labels.

Current limitations:

- The automatic region detector assumes a layout similar to CFD screenshots. Multiple disconnected regions, non-white backgrounds, right-hand or horizontal legends, or very small regions can produce incorrect crops or an extraction error.
- Pixels with an average RGB intensity **below 35 or above 200** are excluded from the scale and contour sampling. Very dark or pale legitimate values may therefore be omitted.
- The legend is sampled down its center column with **uniform value spacing**. Discrete, nonlinear, reversed, or irregular legends are not handled explicitly.
- Color matching is nearest-neighbor matching in RGB space; anti-aliasing, text overlays, compression, and color shifts can affect the estimates.
- The program exports pixel coordinates, **not real-world geometry or simulation mesh coordinates**.
- Excel export requires a local Microsoft Excel installation; there is no CSV export or command-line interface in the current project.

## Project structure

```text
Contour_Value_Extractor.sln
Contour_Value_Extractor/
  MainWindow.xaml                    # WPF user interface
  MainWindow.xaml.cs                 # Input validation and processing workflow
  Helpers/
    ContourScaleExtractor.cs         # Region detection and image cropping
    ScaleToExcel.cs                  # Legend lookup, sampling, and Excel export
  Contour_Value_Extractor.csproj
  packages.config                    # NuGet dependencies
```

## Author

Created by **Navid Siamakmanesh**. For questions or feedback, open a GitHub issue.
