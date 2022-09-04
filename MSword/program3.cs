/*
using Excel = Microsoft.Office.Interop.Excel;

namespace MSword
{
    internal class program3
    {
        static void Main(string[] args)
        {
            
            //Объявляем приложение
            Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
            //Отобразить Excel
            ex.Visible = true;
            //Количество листов в рабочей книге
            ex.SheetsInNewWorkbook = 2;

            ex.Workbooks.Open(@"C:\1\12.xlsx",
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

            //Добавить рабочую книгу
            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);
            //Отключить отображение окон с сообщениями
            ex.DisplayAlerts = false;
            //Получаем первый лист документа (счет начинается с 1)
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
            //Название листа (вкладки снизу)
            sheet.Name = "Тест";
            //Пример заполнения ячеек
            for (int i = 1; i <= 9; i++)
            {
                for (int j = 1; j < 9; j++)
                    sheet.Cells[i, j] = String.Format("Boom {0} {1}", i, j);
            }
            //Захватываем диапазон ячеек
            Excel.Range range1 = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[9, 9]);
            //Шрифт для диапазона
            range1.Cells.Font.Name = "Tahoma";
            //Размер шрифта для диапазона
            range1.Cells.Font.Size = 10;
            //Захватываем другой диапазон ячеек
            Excel.Range range2 = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[9, 2]);
            range2.Cells.Font.Name = "Times New Roman";

            ex.Application.ActiveWorkbook.SaveAs(@"C:\1\12.xlsx", Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }
    }
}
*/