﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Data_From_Excel
{
    class MyExcel
    {
        public static string DB_PATH = @"";
        public static BindingList<Kontakt> EmptyList = new BindingList<Kontakt>();
        // deklaracja Excel
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;

        private static int lastRow = 0;

        /// <summary>
        /// Inicjalizuje Excela.
        /// </summary>
        public static void InitializeExcel()
        {
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(DB_PATH);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1];
            lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
        }//InitializeExcel()

        public static BindingList<Kontakt> ReadMyExcel()
        {
            EmptyList.Clear();
            for (int index = 2; index <= lastRow; index++)
            {
                System.Array MyValues = (System.Array)MySheet.get_Range("A" + index.ToString(), "D" + index.ToString()).Cells.Value;
                EmptyList.Add(new Kontakt
                {
                    Imie = MyValues.GetValue(1, 1).ToString(),
                    Nazwisko = MyValues.GetValue(1, 2).ToString(),
                    Email = MyValues.GetValue(1, 3).ToString(),
                    Telefon = MyValues.GetValue(1, 4).ToString()
                });
            }
            return EmptyList;
        }// ReadMyExcel()

        public static void WriteToExcel(Kontakt kontakt)
        {
            try
            {
                lastRow += 1;
                MySheet.Cells[lastRow, 1] = kontakt.Imie;
                MySheet.Cells[lastRow, 2] = kontakt.Nazwisko;
                MySheet.Cells[lastRow, 3] = kontakt.Email;
                MySheet.Cells[lastRow, 4] = kontakt.Telefon;
                EmptyList.Add(kontakt);
                MyBook.Save();
            }
            catch (Exception ex)
            { }
        }//WriteToExcel

        public static void CloseExcel()
        {
            MyBook.Saved = true;
            MyApp.Quit();
        }// CloseExcel()


    }//class MyExcel
}
