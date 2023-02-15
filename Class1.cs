using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using org.apache.pdfbox;
using java.io;
using OfficeOpenXml;
using System.IO;
using System.Text.RegularExpressions;
using org.apache.pdfbox.util.@operator;

namespace GetTextFromPDFToExcel
{
    public partial class Program
    {
        public static string GetTextPDF(string pdfFile)
        {
            PDDocument doc = null;
            try
            {
                doc = PDDocument.load(pdfFile);
              PDFTextStripper stripper = new PDFTextStripper();
                return stripper.getText(doc);
            }
            finally
            {
                if (doc != null)
                {
                    doc.close();
                }
            }
        }
        public static bool WriteTextToFile(string path, string textToWrite)
        {
            try
            {
                System.IO.File.WriteAllText($"{path}\\OutputText.txt", textToWrite);
                return true;
            }
            catch (Exception)
            {
                System.Console.WriteLine("Exception caught.");
                return false;
            }

        }
        public static bool ConvertToExcel(string path, string fileName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage($"{path}\\new_excel.xlsx"))
            {
                System.Console.WriteLine(path + "\n" + fileName);
                System.Console.ReadKey();
                package.Workbook.Worksheets.Delete("My Sheet");
                var sheet = package.Workbook.Worksheets.Add("My Sheet");
                sheet.Cells[2, 2].Value = "Блюдо";
                sheet.Cells[2, 3].Value = "Описание";
                sheet.Cells[2, 5].Value = "Вес";
                sheet.Cells[2, 6].Value = "Цена";

                using (StreamReader reader = new StreamReader(fileName))
                {
                    int iter = 3;
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        System.Console.WriteLine(line);
                        if (line!="\n" & line!="" & line!=null & line!=" ")
                        {
                            sheet.Cells[iter, 2].Value = line;
                        }
                        
                        //string patternWeightCost = @"[0-9]{2-4}\s*\w+?\.{0,1}\s*[0-9]{2-4}\s*\w+?\.{0,1}";
                        //string patternWeightOrCost = @"[0-9]{2-4}\s*\w\.{0,1}";
                        //if (Regex.IsMatch(line, patternWeightCost, RegexOptions.IgnoreCase))
                        //{
                        //    var regex = new Regex(patternWeightCost);
                        //    var a = regex.Match(line);

                        //    i++;
                        //}
                        iter++;
                    }
                    for (int i = 1; i < iter; i++)
                    {
                        string s = sheet.Cells[i, 2].Value == null? "": sheet.Cells[i, 2].Value.ToString();
                        if (s ==  string.Empty || s == " ")
                        {
                            sheet.Cells[i, 2].Delete(eShiftTypeDelete.EntireRow);
                        }
                    }
                }
                
                // Save to file
                package.Save();
            }
            return true;
        }

    }
}
