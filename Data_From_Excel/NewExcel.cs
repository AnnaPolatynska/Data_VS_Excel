using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;

namespace Data_From_Excel
{



    public static class CreateNewExcel<T>
    {
        /// <summary>
        /// Klasa zapisu do formatu .xlsx
        /// </summary>
        /// <param name="kontakty"></param>
        /// <param name="outputDir"></param>
        /// <param name="fileName"></param>
        public static void ExportToExcel(IEnumerable<T> kontakty, DirectoryInfo outputDir, string fileName)
        {
            fileName = FixFileNameExcel(fileName);

            FileInfo f = new FileInfo(outputDir.FullName + @"\" + fileName);
            DeleteFileIfExist(f);
            using (var excelFile = new ExcelPackage(f))
            {
                var worksheet = excelFile.Workbook.Worksheets.Add("Kontakty");
                worksheet.Cells["A1"].LoadFromCollection(Collection: kontakty, PrintHeaders: true);
                worksheet.Cells.AutoFitColumns(0);
                excelFile.Save();
            }
        }// ExportToExcel

        /// <summary>
        /// Kasuje istniejący plik, jeżeli został już utworzony.
        /// </summary>
        /// <param name="newFile"></param>
        /// <returns></returns>
        public static FileInfo DeleteFileIfExist(FileInfo newFile)
        {
            if (newFile.Exists)
            {
                newFile.Delete();  
                newFile = new FileInfo(newFile.FullName);
            }
            return newFile;
        }// DeleteFileIfExist

        /// <summary>
        /// Zmienia nazwę dowolnego pliku na ".xlsx"
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string FixFileNameExcel(string fileName)
        {
            if (Path.HasExtension(fileName))
            {
                string ext = Path.GetExtension(fileName);

                if (ext != ".xlsx")
                {
                    fileName = Path.GetFileNameWithoutExtension(fileName) + ".xlsx";
                }
            }
            else
            {
                fileName += ".xlsx";
            }

            return fileName;
        }//FixFileNameExcel




        public static void RunMyExcel(IEnumerable<T> kontakty, DirectoryInfo outputDir, string fileName)
        {
            fileName = FixFileNameExcel(fileName);
            FileInfo f = new FileInfo(outputDir.FullName + @"\" + fileName);
            DeleteFileIfExist(f);
            using (var excelFile = new ExcelPackage(f))
            {
                var worksheet = excelFile.Workbook.Worksheets.Add("TEST");

                //Nagłówki
                worksheet.Cells[1, 1].Value = "Imie";
                worksheet.Cells[1, 2].Value = "Nazwisko";
                worksheet.Cells[1, 3].Value = "Email";
                worksheet.Cells[1, 4].Value = "Telefon";
                worksheet.Cells[1, 5].Value = "Dane";

                //1 wiersz danych wprowadzanych z kodu
                //worksheet.Cells["A2"].Value = "Jan";
                //worksheet.Cells["B2"].Value = "Kowalski";
                //worksheet.Cells["C2"].Value = "j.kowalski@gmail.com";
                //worksheet.Cells["D2"].Value = "123456";
                //worksheet.Cells["E2"].Value = "Dodatkowe dane";

                worksheet.Cells.AutoFitColumns(0);// dopasowanie szerokości kolumn do tekstu

                worksheet.Cells["A3"].LoadFromCollection(Collection: kontakty, PrintHeaders: true);
                worksheet.Cells.AutoFitColumns(0);
                excelFile.Save();
            }

            //using (var package = new ExcelPackage())
            //{
            //    //Dodanie nowego arkusza do pustego pliku
            //    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Arkusz_test");
              

            //    //właściwości dokumentu
            //    package.Workbook.Properties.Title = "Arkusz_test";
            //    package.Workbook.Properties.Author = "PWM";

            //    var xlFile = Utils.GetFileInfo("Test.xlsx");
            //}

        }// RunMyExcel()


    }// class CreateNewExcel
}