using System.Drawing.Imaging;
using Spire.Pdf;
using Spire.Pdf.Graphics;

namespace PDFExtractionTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ResetConsoleColor();
            Console.Clear();

            string? pdfPath;
            string? outputDir;

            // User prompt to get PDF file path
            Console.WriteLine();
            pdfPath = GetUserInput("Enter PDF file path:");
            pdfPath = RemoveQuotations(pdfPath);

            while (true)
            {
                if (!File.Exists(pdfPath) || Path.GetExtension(pdfPath).ToLower() != ".pdf")
                {
                    SetConsoleColor("yellow");
                    Console.WriteLine("⚠️ The file does not exist or is not a PDF. Please enter a valid PDF file path.");
                    ResetConsoleColor();
                    Console.WriteLine();
                    pdfPath = GetUserInput("Enter PDF file path:");
                    pdfPath = RemoveQuotations(pdfPath);
                }
                else
                {
                    Console.Clear();
                    Console.WriteLine("📄 Source file: ");
                    SetConsoleColor("cyan");
                    Console.WriteLine(pdfPath);
                    ResetConsoleColor();
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
                        SetConsoleColor("yellow");
                        Console.WriteLine("⚠️ Invalid file path or file does not exist. Please try again.");
                        ResetConsoleColor();
                        pdfPath = GetUserInput("Enter PDF path:");
                    }
                }

                // Determine PDF file name without extension
                // Set the path for the output directory
                string fileNameNoExt = Path.GetFileNameWithoutExtension(pdfPath);
                outputDir = pdfPath.Split(".")[0] + "-Images";

                Console.WriteLine("📁 Export to folder: ");
                SetConsoleColor("cyan");
                Console.WriteLine(outputDir);
                ResetConsoleColor();
                Console.WriteLine();

                // Ensure the output directory exists by creating it
                // Error handling for if creating directory fails (e.g., due to insufficient permissions)
                try
                {
                    Directory.CreateDirectory(outputDir);
                }
                catch (Exception e)
                {
                    SetConsoleColor("red");
                    Console.WriteLine($"Failed to create directory {outputDir}: {e.Message}");
                    ResetConsoleColor();
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
                            SetConsoleColor("green");
                            Console.Write($"Page {i + 1} saved as ");
                            ResetConsoleColor();
                            Console.Write($"{fileNameNoExt}_Page{i + 1}.png");
                            Console.WriteLine();
                        }
                        catch (Exception e)
                        {
                            SetConsoleColor("red");
                            Console.WriteLine($"Failed to save Page {i + 1} as an image: {e.Message}");
                            ResetConsoleColor();
                        }
                    }
                }

                Console.WriteLine();
                SetConsoleColor("green");
                Console.WriteLine("✅ PDF pages extracted and saved as PNG images.\n");
                ResetConsoleColor();
            }
        }

        static string? GetUserInput(string prompt)
        {
            Console.Write(prompt + " ");
            string? inputLine = Console.ReadLine();

            return string.IsNullOrEmpty(inputLine) ? null : inputLine;
        }

        static string? RemoveQuotations(string path)
        {
            return string.IsNullOrEmpty(path) ? null : path.Trim(new char[] { '"' });
        }

        static void SetConsoleColor(string color)
        {
            switch (color)
            {
                case "red":
                    Console.ForegroundColor = ConsoleColor.Red;
                    break;
                case "green":
                    Console.ForegroundColor = ConsoleColor.Green;
                    break;
                case "yellow":
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    break;
                case "blue":
                    Console.ForegroundColor = ConsoleColor.Blue;
                    break;
                case "cyan":
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    break;
                case "magenta":
                    Console.ForegroundColor = ConsoleColor.Magenta;
                    break;
                default:
                    break;
            }
        }

        static void ResetConsoleColor()
        {
            Console.ForegroundColor = ConsoleColor.White;
        }
    }
}