using Excel = Microsoft.Office.Interop.Excel;

class SomeInteropWithExcel
{
    public void Test()
    {
        Excel.Application xlApp = new Excel.Application { Visible = true };
        Excel.Workbook xlBook = xlApp.Workbooks[1];
        Excel.Worksheet sheet = xlBook.Sheets[1] as Excel.Worksheet;
        int lastRow = sheet.Cells[sheet.Rows.Count, "Q"].End[Excel.XlDirection.xlUp].Row;
        dynamic value = sheet.Cells[lastRow, "Q"].Value;
    }
}