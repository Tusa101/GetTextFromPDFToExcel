using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using java.nio.file;


namespace GetTextFromPDFToExcel
{
    public partial class Program
    {
        struct Dir
        {
            public string DirPath { get; set; }
            public string FileName { get; set; }
        }
        static void Main(string[] args)
        {
            while (true)
            {
                Console.Clear();
                DisplayHeader();
                Dir dirWithPDF = new Dir();
                DriveInfo driveNameInfo;
                while (true)
                {
                    Console.WriteLine("First of all, choose the disk (write only one symbol, e.g. 'C'):");
                    char driveName = char.ToUpper(Console.ReadKey(true).KeyChar);
                    try
                    {
                        driveNameInfo = new DriveInfo($"{driveName}:\\");
                        break;
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("There is no such drive, try again!");
                        Thread.Sleep(3000);
                    }
                }
                Console.Clear();
                Console.WriteLine("Please, choose the location of PDF:");
                Console.WriteLine("(Preffered format is dir\\dir\\...\\dir)");
                DirectoryInfo dirInfo = driveNameInfo.RootDirectory;
                Console.Write(dirInfo.Name);
                dirWithPDF.DirPath = $"{dirInfo.Name}{Console.ReadLine()}";
                string[] pdfFiles;
                try
                {
                    pdfFiles = Directory.GetFiles(dirWithPDF.DirPath, "*.pdf");
                }
                catch (Exception)
                {
                    Console.WriteLine("There is no such directory, try again!");
                    continue;
                }
                
                if (pdfFiles.Length == 0)
                {
                    Console.WriteLine("There are no PDFs in this directory. Try again.");
                }
                else
                {
                    Console.WriteLine("There some files found in the directory:");
                    for (int i = 0; i < pdfFiles.Length; i++)
                    {
                        Console.WriteLine($"{i + 1}. {pdfFiles[i]}");
                    }
                    Console.WriteLine("Choose the file to convert:");
                    while (true)
                    {
                        try
                        {
                            dirWithPDF.FileName = pdfFiles[Int32.Parse(Console.ReadLine())-1];
                            break;
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("There is no file on the current position. Try again!");
                        }
                    }
                    Console.WriteLine($"Selected file is: {dirWithPDF.FileName.Substring(dirWithPDF.FileName.LastIndexOf('\\') + 1)}");
                    Console.WriteLine("Extraction started...");
                    string output = GetTextPDF(dirWithPDF.FileName);
                    if(WriteTextToFile(dirWithPDF.DirPath, output))
                    {
                        Console.WriteLine("Extraction ended.");
                        while (true)
                        {
                            Console.WriteLine("Do you want to read the file here or open (press R or O)?");
                            char key = char.ToUpper(Console.ReadKey(true).KeyChar);
                            switch (key)
                            {
                                case 'R':
                                    {
                                        Console.WriteLine(output);
                                        Console.ReadKey();
                                    }
                                    break;
                                case 'O':
                                    {
                                        Console.WriteLine($"{dirWithPDF.DirPath}\\OutputText.txt");
                                        File.Open($"{dirWithPDF.DirPath}OutputText.txt", FileMode.Open, FileAccess.Read);
                                        Console.ReadKey();
                                    }
                                    break;
                                default:
                                    {
                                        Console.WriteLine("No such key, try again!");
                                    }
                                    break;
                            }
                            if (key == 'O' || key == 'R')
                            {
                                break;
                            }
                        }
                    }
                }
            }
        }

        public static bool WriteTextToFile(string path, string textToWrite)
        {
            try
            {
                File.WriteAllText($"{path}\\OutputText.txt", textToWrite);
                return true;
            }
            catch (Exception)
            {
                Console.WriteLine("Exception caught.");
                return false;
            }
            
        }

        private static void DisplayHeader()
        {
            Console.WriteLine("This program is for writing text frrom PDF files into Excel.");
            Console.WriteLine("Created by Tuskaev Alexandr.");
            
        }
    }
}
