using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFExtractionTest
{
    internal class Helpers
    {
        public static string? GetUserInput(string prompt)
        {
            Console.Write(prompt + " ");
            string? inputLine = Console.ReadLine();

            return string.IsNullOrEmpty(inputLine) ? null : inputLine;
        }

        public static string? RemoveQuotations(string path)
        {
            return string.IsNullOrEmpty(path) ? null : path.Trim(new char[] { '"' });
        }

        public static void SetConsoleColor(string color)
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
                case "cyan":
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    break;
                default:
                    Console.ForegroundColor = ConsoleColor.White;
                    break;
            }
        }

        public static void ResetConsoleColor()
        {
            Console.ForegroundColor = ConsoleColor.White;
        }
    }
}
