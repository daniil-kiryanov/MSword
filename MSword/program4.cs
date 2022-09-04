/*
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ConsoleApplication13;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text;

namespace ConsoleApplication13
{
    class Price
    {

        public string Name { get; set; }
        public string Cost { get; set; }
        public string Site { get; set; }

    }
    class PrintExel
    {
        public static void ExportToExcel(List<Price> vPices)
        {
            // Загрузить Excel, затем создать новую пустую рабочую книгу
            Excel.Application excelApp = new Excel.Application();

            // Сделать приложение Excel видимым
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            Excel.Worksheets workSheet = excelApp.ActiveSheet;
            //Excel._Worksheet workSheet = excelApp.ActiveSheet;
            // Установить заголовки столбцов в ячейках
            workSheet.Cells[1, "A"] = "NameCompany";
            workSheet.Cells[1, "B"] = "Site";
            workSheet.Cells[1, "C"] = "Cost";

            string parser = File.ReadAllText(@"parser.txt", Encoding.Default);

            int parsers = Convert.ToInt32(parser);
            int row = 1;
            foreach (Price c in vPices)
            {
                row++;
                workSheet.Cells[parsers, "A"] = c.Name;
                workSheet.Cells[parsers, "B"] = c.Site;
                workSheet.Cells[parsers, "C"] = c.Cost;
            }


            excelApp.DisplayAlerts = false;
            workSheet.SaveAs(string.Format(@"{0}\Price.xlsx", Environment.CurrentDirectory));

            excelApp.Quit();

        }

    }

    class Program
    {
        static void Main(string[] args)
        {
            /// Тест записи в эксель
            string a = File.ReadAllText(@"title.txt", Encoding.Default);
            string b = File.ReadAllText(@"asdd.txt", Encoding.Default);
            string c = File.ReadAllText(@"asd.txt", Encoding.Default);
            var ListPricee = new List<Price>();
            ListPricee.Add(new Price { Name = a, Site = b, Cost = c });


            // Записываем в эксель
            PrintExel.ExportToExcel(ListPricee);


        }
    }
}
*/