using Excel = Microsoft.Office.Interop.Excel;
namespace MSword
{
    internal class ExcelProgram
    {
        public static void Main()
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook book = app.Workbooks.Add();
            Excel.Worksheet sheet = (Excel.Worksheet)app.Sheets.Add();
            sheet.Range["A1"].Value = "LoremIpsum";
            book.SaveAs(@"C:\1\1.xlsx");
            book.Close();
            app.Quit();

            /*
            Excel.Application app = null;
            Excel.Workbooks books = null;
            Excel.Workbook book = null;
            Excel.Sheets sheets = null;
            Excel.Worksheet sheet = null;
            Excel.Range range = null;

                try
                {
                    app = new Excel.Application();
                    books = app.Workbooks;
                    book = books.Add();
                    sheets = book.Sheets;
                    sheet = sheets.Add();
                    range = sheet.Range["A1"];
                    range.Value = "Lorem Ipsum";
                    book.SaveAs(@"c:\1\1.xlsx");
                    book.Close();
                    app.Quit();
                }
                finally
                {
                    if (range != null) Marshal.ReleaseComObject(range);
                    if (sheet != null) Marshal.ReleaseComObject(sheet);
                    if (sheets != null) Marshal.ReleaseComObject(sheets);
                    if (book != null) Marshal.ReleaseComObject(book);
                    if (books != null) Marshal.ReleaseComObject(books);
                    if (app != null) Marshal.ReleaseComObject(app);
                }
            */
        }
    }
}
