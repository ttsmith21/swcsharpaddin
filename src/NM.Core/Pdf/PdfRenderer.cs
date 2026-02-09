using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;

namespace NM.Core.Pdf
{
    /// <summary>
    /// Renders PDF pages to PNG images for AI Vision analysis.
    /// Uses an external renderer (Ghostscript or pdftoppm) if available,
    /// or returns raw PDF bytes for services that accept PDF directly (Gemini).
    /// </summary>
    public static class PdfRenderer
    {
        /// <summary>
        /// Returns the raw PDF bytes for the specified page range.
        /// For services like Gemini that accept PDF natively, this avoids
        /// the rendering step entirely — just send the PDF as application/pdf.
        /// </summary>
        public static byte[] GetPdfBytes(string pdfPath)
        {
            if (string.IsNullOrEmpty(pdfPath) || !File.Exists(pdfPath))
                return null;

            try
            {
                return File.ReadAllBytes(pdfPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PdfRenderer] Failed to read PDF: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Renders a specific page of a PDF to a PNG byte array.
        /// Tries Ghostscript first, then pdftoppm (poppler), then returns null.
        /// </summary>
        /// <param name="pdfPath">Path to the PDF file.</param>
        /// <param name="pageIndex">Zero-based page index.</param>
        /// <param name="dpi">Resolution in dots per inch (200 for full page, 300 for detail).</param>
        /// <returns>PNG image bytes, or null if no renderer is available.</returns>
        public static byte[] RenderPageToPng(string pdfPath, int pageIndex = 0, int dpi = 200)
        {
            if (string.IsNullOrEmpty(pdfPath) || !File.Exists(pdfPath))
                return null;

            // Try Ghostscript
            byte[] result = TryGhostscriptRender(pdfPath, pageIndex, dpi);
            if (result != null) return result;

            // Try pdftoppm (poppler-utils)
            result = TryPdftoppmRender(pdfPath, pageIndex, dpi);
            if (result != null) return result;

            System.Diagnostics.Debug.WriteLine(
                "[PdfRenderer] No PDF renderer available (install Ghostscript or poppler-utils for PNG rendering). " +
                "Use GetPdfBytes() for services that accept PDF natively.");
            return null;
        }

        /// <summary>
        /// Renders the title block region (bottom-right quadrant) of a PDF page.
        /// If PNG rendering is available, crops the result. Otherwise returns null.
        /// </summary>
        public static byte[] RenderTitleBlockRegion(string pdfPath, int pageIndex = 0, int dpi = 300)
        {
            // Render full page at higher DPI, then crop
            byte[] fullPage = RenderPageToPng(pdfPath, pageIndex, dpi);
            if (fullPage == null) return null;

            try
            {
                return CropTitleBlock(fullPage);
            }
            catch
            {
                // If cropping fails, return full page
                return fullPage;
            }
        }

        /// <summary>
        /// Checks whether a PNG renderer (Ghostscript or pdftoppm) is available.
        /// </summary>
        public static bool IsPngRendererAvailable()
        {
            return FindGhostscript() != null || FindPdftoppm() != null;
        }

        /// <summary>
        /// Gets the number of pages in a PDF file using PdfPig (already a dependency).
        /// </summary>
        public static int GetPageCount(string pdfPath)
        {
            if (string.IsNullOrEmpty(pdfPath) || !File.Exists(pdfPath))
                return 0;

            try
            {
                using (var doc = UglyToad.PdfPig.PdfDocument.Open(pdfPath))
                {
                    return doc.NumberOfPages;
                }
            }
            catch
            {
                return 0;
            }
        }

        // --- Ghostscript rendering ---
        private static byte[] TryGhostscriptRender(string pdfPath, int pageIndex, int dpi)
        {
            string gs = FindGhostscript();
            if (gs == null) return null;

            try
            {
                string tempPng = Path.GetTempFileName() + ".png";
                int pageNum = pageIndex + 1; // GS uses 1-based pages

                var psi = new ProcessStartInfo
                {
                    FileName = gs,
                    Arguments = $"-dBATCH -dNOPAUSE -dQUIET -sDEVICE=png16m -r{dpi} " +
                                $"-dFirstPage={pageNum} -dLastPage={pageNum} " +
                                $"-sOutputFile=\"{tempPng}\" \"{pdfPath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardError = true
                };

                using (var proc = Process.Start(psi))
                {
                    proc.WaitForExit(30000);
                    if (proc.ExitCode == 0 && File.Exists(tempPng))
                    {
                        byte[] result = File.ReadAllBytes(tempPng);
                        File.Delete(tempPng);
                        return result;
                    }
                }

                if (File.Exists(tempPng)) File.Delete(tempPng);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PdfRenderer] Ghostscript render failed: {ex.Message}");
            }

            return null;
        }

        // --- pdftoppm rendering ---
        private static byte[] TryPdftoppmRender(string pdfPath, int pageIndex, int dpi)
        {
            string pdftoppm = FindPdftoppm();
            if (pdftoppm == null) return null;

            try
            {
                string tempPrefix = Path.Combine(Path.GetTempPath(), "pdfrender_" + Guid.NewGuid().ToString("N"));
                int pageNum = pageIndex + 1;

                var psi = new ProcessStartInfo
                {
                    FileName = pdftoppm,
                    Arguments = $"-png -r {dpi} -f {pageNum} -l {pageNum} \"{pdfPath}\" \"{tempPrefix}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardError = true
                };

                using (var proc = Process.Start(psi))
                {
                    proc.WaitForExit(30000);
                }

                // pdftoppm outputs files like tempPrefix-01.png
                string expectedFile = $"{tempPrefix}-{pageNum:D2}.png";
                if (!File.Exists(expectedFile))
                    expectedFile = $"{tempPrefix}-{pageNum}.png";

                if (File.Exists(expectedFile))
                {
                    byte[] result = File.ReadAllBytes(expectedFile);
                    File.Delete(expectedFile);
                    return result;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PdfRenderer] pdftoppm render failed: {ex.Message}");
            }

            return null;
        }

        private static string FindGhostscript()
        {
            string[] paths =
            {
                @"C:\Program Files\gs\gs10.04.0\bin\gswin64c.exe",
                @"C:\Program Files\gs\gs10.03.1\bin\gswin64c.exe",
                @"C:\Program Files\gs\gs10.02.1\bin\gswin64c.exe",
                @"C:\Program Files\gs\gs10.01.2\bin\gswin64c.exe",
                @"C:\Program Files\gs\gs10.00.0\bin\gswin64c.exe",
                @"C:\Program Files (x86)\gs\gs10.00.0\bin\gswin32c.exe",
            };

            foreach (var path in paths)
            {
                if (File.Exists(path)) return path;
            }

            // Try to find any gs installation
            string gsDir = @"C:\Program Files\gs";
            if (Directory.Exists(gsDir))
            {
                foreach (var dir in Directory.GetDirectories(gsDir, "gs*"))
                {
                    string exe = Path.Combine(dir, "bin", "gswin64c.exe");
                    if (File.Exists(exe)) return exe;
                }
            }

            return null;
        }

        private static string FindPdftoppm()
        {
            string[] paths =
            {
                @"C:\Program Files\poppler\bin\pdftoppm.exe",
                @"C:\poppler\bin\pdftoppm.exe",
            };

            foreach (var path in paths)
            {
                if (File.Exists(path)) return path;
            }

            return null;
        }

        /// <summary>
        /// Crops a PNG image to the title block region (bottom-right ~45% x ~30%).
        /// Uses System.Drawing (GDI+) which is available in .NET Framework.
        /// </summary>
        private static byte[] CropTitleBlock(byte[] pngBytes)
        {
            using (var ms = new MemoryStream(pngBytes))
            using (var bitmap = new Bitmap(ms))
            {
                int cropX = (int)(bitmap.Width * 0.55);
                int cropY = (int)(bitmap.Height * 0.70);
                int cropW = bitmap.Width - cropX;
                int cropH = bitmap.Height - cropY;

                var cropRect = new Rectangle(cropX, cropY, cropW, cropH);

                using (var cropped = bitmap.Clone(cropRect, bitmap.PixelFormat))
                using (var output = new MemoryStream())
                {
                    cropped.Save(output, System.Drawing.Imaging.ImageFormat.Png);
                    return output.ToArray();
                }
            }
        }
    }
}
