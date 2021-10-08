using System;
using System.IO;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;


namespace CreateSpreadSheet
{
    public static class Program
    {
        private static void Main()
        {
            var a = new CreateSpreadSheet();
        }
        private static void CheckFile(string path)
        {
            Console.Write(File.Exists(path) ? "File exist" : "File does not exist");
        }
    }
}