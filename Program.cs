using System.Drawing.Imaging;
using Spire.Pdf;
using Spire.Pdf.Graphics;

namespace PDFExtractionTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Helpers.ResetConsoleColor();
            Console.Clear();

            string? pdfPath;
            string? outputDir;

            // User prompt to get PDF file path
            Console.WriteLine();
            pdfPath = Helpers.GetUserInput("Enter PDF file path:");
            pdfPath = Helpers.RemoveQuotations(pdfPath);

            while (true)
            {
                if (!File.Exists(pdfPath) || Path.GetExtension(pdfPath).ToLower() != ".pdf")
                {
                    Helpers.SetConsoleColor("yellow");
                    Console.WriteLine("⚠️ The file does not exist or is not a PDF. Please enter a valid PDF file path.");
                    Helpers.ResetConsoleColor();
                    Console.WriteLine();
                    pdfPath = Helpers.GetUserInput("Enter PDF file path:");
                    pdfPath = Helpers.RemoveQuotations(pdfPath);
                }
                else
                {
                    Console.Clear();
                    Console.WriteLine("📄 Source file: ");
                    Helpers.SetConsoleColor("cyan");
                    Console.WriteLine(pdfPath);
                    Helpers.ResetConsoleColor();
                    Console.WriteLine();
                    break;
                }
            }

            // Create a PdfDocument instance and load the PDF
            // If the file does not exist, keep prompting the user to enter a valid path
            // 'using' keyword properly disposes of 'PdfDocument' (and later 'Image') object to release resources
            using (PdfDocument pdf = new PdfDocument())
            {
                bool isValidPath = false;

                while (!isValidPath)
                {
                    try
                    {
                        pdf.LoadFromFile(pdfPath);
                        isValidPath = true;
                    }
                    catch (Exception e)
                    {
                        Helpers.SetConsoleColor("yellow");
                        Console.WriteLine("⚠️ Invalid file path or file does not exist. Please try again.");
                        Helpers.ResetConsoleColor();
                        pdfPath = Helpers.GetUserInput("Enter PDF path:");
                    }
                }

                // Determine PDF file name without extension
                // Set the path for the output directory
                string fileNameNoExt = Path.GetFileNameWithoutExtension(pdfPath);
                outputDir = pdfPath.Split(".")[0] + "-Images";

                Console.WriteLine("📁 Export to folder: ");
                Helpers.SetConsoleColor("cyan");
                Console.WriteLine(outputDir);
                Helpers.ResetConsoleColor();
                Console.WriteLine();

                // Ensure the output directory exists by creating it
                // Error handling for if creating directory fails (e.g., due to insufficient permissions)
                try
                {
                    Directory.CreateDirectory(outputDir);
                }
                catch (Exception e)
                {
                    Helpers.SetConsoleColor("red");
                    Console.WriteLine($"Failed to create directory {outputDir}: {e.Message}");
                    Helpers.ResetConsoleColor();
                    return;
                }

                // Loop through each page in the PDF
                for (int i = 0; i < pdf.Pages.Count; i++)
                {
                    // Convert all pages to images and set the image Dpi to 500
                    using (System.Drawing.Image image = pdf.SaveAsImage(i, PdfImageType.Bitmap, 500, 500))
                    {
                        // Define the output file path for each new image
                        string newFilePath = Path.Combine(outputDir, $"{fileNameNoExt}_Page{i + 1}.png");

                        // Save images as PNG format to a specified folder 
                        // Console log fail message if image extraction fails
                        try
                        {
                            image.Save(newFilePath, ImageFormat.Png);
                            Helpers.SetConsoleColor("green");
                            Console.Write($"Page {i + 1} saved as ");
                            Helpers.ResetConsoleColor();
                            Console.Write($"{fileNameNoExt}_Page{i + 1}.png");
                            Console.WriteLine();
                        }
                        catch (Exception e)
                        {
                            Helpers.SetConsoleColor("red");
                            Console.WriteLine($"Failed to save Page {i + 1} as an image: {e.Message}");
                            Helpers.ResetConsoleColor();
                        }
                    }
                }

                Console.WriteLine();
                Helpers.SetConsoleColor("green");
                Console.WriteLine("✅ PDF pages extracted and saved as PNG images.\n");
                Helpers.ResetConsoleColor();
            }
        }
    }
}