// Created by Navid Siamakmanehs
// For additional information or help please contact me at Navid.siamakmanesh@gmail.com
// Enjoy


using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Globalization;
using System.Windows;
using System.Windows.Media.Imaging;
using Contour_Value_Extractor.Helpers;

namespace Contour_Value_Extractor
{
    public partial class MainWindow : Window
    {
        private string _selectedImagePath;
        private string _outputFolder;

        public MainWindow()
        {
            InitializeComponent();
        }











        private void BtnChoose_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Image Files (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp"
            };

            if (dlg.ShowDialog() != true)
                return;

            _selectedImagePath = dlg.FileName;

            // Preview
            var bmp = new BitmapImage();
            bmp.BeginInit();
            bmp.CacheOption = BitmapCacheOption.OnLoad;
            bmp.UriSource = new Uri(_selectedImagePath);
            bmp.EndInit();
            bmp.Freeze();

            ImgPreview.Source = bmp;
            PreviewPlaceholder.Visibility = Visibility.Collapsed;
            PreviewScroll.Visibility = Visibility.Visible;

            BtnExtract.IsEnabled = true;
            TxtStatus.Text = "Image selected.";
        }

        private void BtnExtract_Click(object sender, RoutedEventArgs e)
        {
            // Validate inputs
            if (!TryParseDouble(TxtMin.Text, out double tMin) ||
                !TryParseDouble(TxtMax.Text, out double tMax))
            {
                MessageBox.Show("Min and Max values must be numbers.");
                return;
            }

            if (!int.TryParse(TxtPoints.Text, out int points) || points <= 0)
            {
                MessageBox.Show("Number of points must be a positive integer.");
                return;
            }

            string baseDir = System.IO.Path.GetDirectoryName(_selectedImagePath);
            string baseName = System.IO.Path.GetFileNameWithoutExtension(_selectedImagePath);
            string ext = System.IO.Path.GetExtension(_selectedImagePath);

            _outputFolder = System.IO.Path.Combine(baseDir, "extracted_" + baseName);
            System.IO.Directory.CreateDirectory(_outputFolder);

            string contourPath = System.IO.Path.Combine(_outputFolder, baseName + "_contour" + ext);
            string scalePath = System.IO.Path.Combine(_outputFolder, baseName + "_scale" + ext);

            DeleteIfExists(contourPath);
            DeleteIfExists(scalePath);

            // --- Extract contour & scale ---
            try
            {
                ContourScaleExtractor.ExtractContourAndScale(
                    imagePath: _selectedImagePath,
                    outContourPath: contourPath,
                    outScalePath: scalePath,
                    whiteThresh: 245,
                    minArea: 5000,
                    tightPad: 2
                );
            }
            catch
            {
                MessageBox.Show("The image is not OK. Contact admin.");
                return;
            }

            // --- Excel ---
            try
            {
                string excelExpected = System.IO.Path.Combine(
                    _outputFolder,
                    baseName + "_scale_and_points.xlsx"
                );

                DeleteIfExists(excelExpected);

                ScaleToExcel.ExportScaleAndPointsToExcel(
                    scaleImagePath: scalePath,
                    contourImagePath: contourPath,
                    tMin: tMin,
                    tMax: tMax,
                    requestedPoints: points
                );
            }
            catch
            {
                MessageBox.Show("The image is not OK. Contact admin.");
                return;
            }

            TxtOutputPath.Text = _outputFolder;
            BtnOpenFolder.IsEnabled = true;
            TxtStatus.Text = "Done.";
            MessageBox.Show("Done.");
        }
        private static void OpenFolder(string folderPath)
        {
            if (!System.IO.Directory.Exists(folderPath))
                return;

            Process.Start(new ProcessStartInfo
            {
                FileName = folderPath,
                UseShellExecute = true
            });
        }

        private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
        {
            OpenFolder(_outputFolder);

        }

        private static void DeleteIfExists(string path)
        {
            if (System.IO.File.Exists(path))
                System.IO.File.Delete(path);
        }

        private static bool TryParseDouble(string text, out double value)
        {
            return double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value) ||
                   double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out value);
        }
    }
}
