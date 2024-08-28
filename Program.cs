using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Spire.Pdf;
using Spire.Pdf.Graphics;
using static System.Net.Mime.MediaTypeNames;

namespace PDFExtractionTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string? pdfPath;
            string? outputDir;

            pdfPath = GetUserInput("Enter PDF path:");

            while (true)
            {
                if (!File.Exists(pdfPath) || Path.GetExtension(pdfPath).ToLower() != ".pdf")
                {
                    Console.WriteLine("The file does not exist or is not a PDF. Please enter a valid PDF path.");
                    Console.WriteLine();
                    pdfPath = GetUserInput("Enter PDF path:");
                }
                else break;
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
                        Console.WriteLine("Invalid file path or file does not exist. Please try again.");
                        pdfPath = GetUserInput("Enter PDF path:");
                    }
                }

                // Determine PDF file name without extension
                // Set the path for the output directory
                string fileNameNoExt = Path.GetFileNameWithoutExtension(pdfPath);
                outputDir = pdfPath.Split(".")[0] + "-Images";

                Console.WriteLine();
                Console.WriteLine($"Your images will be exported to: {outputDir}");
                Console.WriteLine();

                // Ensure the output directory exists by creating it
                // Error handling for if creating directory fails (e.g., due to insufficient permissions)
                try
                {
                    Directory.CreateDirectory(outputDir);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Failed to create directory {outputDir}: {e.Message}");
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
                            Console.WriteLine($"Page {i + 1} saved as {newFilePath}");
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine($"Failed to save Page {i + 1} as an image: {e.Message}");
                        }
                    }
                }

                Console.WriteLine();
                Console.WriteLine("PDF pages extracted and saved as PNG images.");
            }
        }

        static string? GetUserInput(string prompt)
        {
            Console.Write(prompt + " ");
            string? inputLine = Console.ReadLine();

            return string.IsNullOrEmpty(inputLine) ? null : inputLine;
        }
    }
}