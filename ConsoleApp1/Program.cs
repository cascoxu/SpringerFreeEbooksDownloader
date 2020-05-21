using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;

namespace ConsoleApp1
{
    class Program
    {
        public static void Main(string[] args)
        {
            string currentBookTitle="";
            string currentBookUrl="";
            Config config = new Config();

            try
            {
                config = GetConfig();
                List<Book> books = ReadBooksFromExcel(config);            

                for (int i = 0; i < books.Count; i++)
                {
                    //Update currentBookTitle and currentBookUrl for use in case of crash
                    currentBookTitle = books[i].title + " - " + books[i].edition + " - " + books[i].author;
                    currentBookUrl = books[i].openDownloadPage;

                    string savePath = GetSaveBasePath(config, books[i]);                           
                    string bookName = books[i].title + " - " + books[i].edition + " - " + books[i].author;                                 

                    SavePDFIfNoExist(config, savePath, bookName, books[i]);              
                    SaveInfoFileIfNoExist(config, savePath, bookName, books[i]);                   
                }
            }
            catch (Exception Ex)
            {
                WriteErrorFile(config, currentBookTitle, currentBookUrl);
                throw Ex;
            }
        }       

        public static Config GetConfig()
        {
            Config c = new Config();
            StreamReader configReader = new StreamReader("config.txt");

            while (!configReader.EndOfStream)
            {
                string configuration = configReader.ReadLine();
                string[] keyValue = configuration.Split("::");

                c.values.Add(keyValue[0], keyValue[1]);
            }

            configReader.Close();

            return c;
        }

        public static string GetSaveBasePath(Config config, Book book)
        {
            string savePath = config.values[config.BaseDir];
            if (config.values[config.SaveInFoldersByPackageName].ToLower() == "yes")
            {
                savePath += "\\" + book.category;
            }

            if (config.values[config.SaveInFoldersByProductType].ToLower() == "yes")
            {
                savePath += "\\" + book.level;
            }
            return savePath;
        }

        public static List<Book> ReadBooksFromExcel(Config config)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage(new FileInfo(config.values[config.ExcelName]));            
            ExcelWorksheet firstSheet = package.Workbook.Worksheets[config.values[config.SheetName]];

            List<Book> books = new List<Book>();

            int rows = firstSheet.Cells.Rows;
            int start = 2;
            //Starts in 2 to avoid headers
            for (int i = start; i <= rows; i++)
            {
                ExcelRange titleRange = firstSheet.Cells[i, 1];
                ExcelRange editionRange = firstSheet.Cells[i, 3];
                ExcelRange authorRange = firstSheet.Cells[i, 2];
                ExcelRange levelRange = firstSheet.Cells[i, 4];
                ExcelRange yearRange = firstSheet.Cells[i, 5];
                ExcelRange printISBNRange = firstSheet.Cells[i, 7];
                ExcelRange electronicISBNRange = firstSheet.Cells[i, 8];
                ExcelRange languajeRange = firstSheet.Cells[i, 10];
                ExcelRange categoryRange = firstSheet.Cells[i, 12];
                ExcelRange seriesTitleRange = firstSheet.Cells[i, 16];
                ExcelRange openDownloadPageRange = firstSheet.Cells[i, 19];
                ExcelRange ssubjectsRange = firstSheet.Cells[i, 20];

                if (titleRange.GetValue<String>() != null)
                {
                    Book newBook = new Book();
                    newBook.title = titleRange.GetValue<String>();
                    newBook.title = newBook.title.Replace(':', ' ');
                    newBook.title = newBook.title.Replace('/', ' ');
                    newBook.title = newBook.title.Replace('\\', ' ');
                    newBook.edition = editionRange.GetValue<String>();
                    newBook.author = authorRange.GetValue<String>();
                    newBook.level = levelRange.GetValue<String>();
                    newBook.level = newBook.level = newBook.level.Replace('/', '-');
                    newBook.year = yearRange.GetValue<String>();
                    newBook.printISBN = printISBNRange.GetValue<String>();
                    newBook.electronicSBN = electronicISBNRange.GetValue<String>();
                    newBook.languaje = languajeRange.GetValue<String>();
                    newBook.category = categoryRange.GetValue<String>();
                    newBook.seriesTitle = seriesTitleRange.GetValue<String>();
                    newBook.openDownloadPage = openDownloadPageRange.GetValue<String>();
                    newBook.subjects = ssubjectsRange.GetValue<String>();

                    books.Add(newBook);
                }
            }

            return books;
        }

        public static void SavePDFIfNoExist(Config config, string savePath, string bookName, Book book)
        {
            string savePathPDF = savePath + "\\" + bookName + ".pdf";
            WebClient client = new WebClient();

            if (!File.Exists(savePathPDF))
            {
                String content = client.DownloadString(book.openDownloadPage);

                Regex regex = new Regex("pdf(.*)pdf");
                Match match = regex.Match(content);
                if (match.Success)
                {
                    string s = match.Value;
                    s = s.Remove(s.IndexOf("\""));

                    string downloadURL = config.values[config.DownloadURLBase] + s;

                    using (WebClient client2 = new WebClient())
                    {
                        Uri u = new Uri(downloadURL);

                        if (!Directory.Exists(savePath))
                        {
                            Directory.CreateDirectory(savePath);
                        }

                        client2.DownloadFile(downloadURL, savePathPDF);
                    }
                }
            }
        }

        private static void SaveInfoFileIfNoExist(Config config, string savePath, string bookName, Book book)
        {
            string savePathInfo = savePath + "\\" + bookName + ".txt";

            if (config.values[config.SaveBookDataInTxt].ToLower() == "yes")
            {
                if (!File.Exists(savePathInfo))
                {
                    StreamWriter sw = new StreamWriter(savePathInfo);
                    sw.WriteLine("TITLE: " + book.title);
                    sw.WriteLine("EDITION: " + book.edition);
                    sw.WriteLine("AUTHOR: " + book.author);
                    sw.WriteLine("SCHOLAR LEVEL: " + book.level);
                    sw.WriteLine("YEAR: " + book.year);
                    sw.WriteLine("PRINT ISBN: " + book.printISBN);
                    sw.WriteLine("ELECTRONIC ISBN: " + book.electronicSBN);
                    sw.WriteLine("LANGUAJE: " + book.languaje);
                    sw.WriteLine("CATEGORY: " + book.category);
                    sw.WriteLine("SERIESTITLE: " + book.seriesTitle);
                    sw.WriteLine("SUBJECTS: " + book.subjects);
                    sw.WriteLine("OPEN DOWNLOAD PAGE: " + book.openDownloadPage);
                    sw.Close();
                }
            }
        }

        public static void WriteErrorFile(Config config, string currentBookTitle, string currentBookUrl)
        {
            StreamWriter sw = new StreamWriter(config.values[config.BaseDir] + "\\ERROR.txt");
            sw.WriteLine("Error trying to download the following book: " + currentBookTitle);
            sw.WriteLine("Try to download it manually from the following url: " + currentBookUrl);
            sw.WriteLine("Please delete it from the excel file deleting the row and restart the application again.");
            sw.WriteLine("After the restart the application will avoid to download the books downloaded in a previous run.");
            sw.Close();
        }

        public class Book
        {
            public Book()
            {
            }

            public string title; //1
            public string edition; //3
            public string author; //2
            public string level; //4
            public string year; //5
            public string printISBN; //7
            public string electronicSBN; //8
            public string languaje; //10
            public string category; //12
            public string seriesTitle; //16
            public string openDownloadPage; //19
            public string subjects; //20
        }

       
        public class Config
        {
            public Config()
            {
                values = new Dictionary<string, string>();
            }

            public Dictionary<string, string> values;

            public string ExcelName = "ExcelName";
            public string SheetName = "SheetName";
            public string SaveInFoldersByPackageName = "SaveInFoldersByPackageName";
            public string SaveInFoldersByProductType = "SaveInFoldersByProductType";
            public string SaveBookDataInTxt = "SaveBookDataInTxt";
            public string BaseDir = "BaseDir";
            public string DownloadURLBase = "DownloadURLBase";
        }
    }
}
