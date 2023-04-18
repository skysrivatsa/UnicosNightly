using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace NightlyUnicos
{
    class ExcelClass
    {
        _Application excel = new Excel.Application();
        Workbook wb;
        Worksheet ws;
        Excel.Hyperlinks links;
        public void OpenExcelMethod(string path)
        {
            wb = excel.Workbooks.Open(path);
        }

        public string GetExcelName()
        {
            return wb.Name.ToString();
        }

        public int GetSheetNumber()
        {
            ws = wb.ActiveSheet;
            return ws.Index;
        }

        public string GetSheetName()
        {
            ws = wb.ActiveSheet;
            return ws.Name.ToString();
        }

        public void SelectSheetNumber(int sheetNumber)
        {
            ws = wb.Worksheets[sheetNumber];
            ws.Select();
        }

        public void CloseExcelMethod()
        {
            wb.Close(false, Type.Missing, Type.Missing);
            excel.Quit();

        }

        public string ExcelReadCell(int i, int j)
        {
            if (ws.Cells[i, j].Value2 != null)
            {
                return ws.Cells[i, j].Value2;
            }

            else
                return "";
        }

        public void ExcelWriteCell(int i, int j, string valu)
        {
            excel.Cells[i, j].value2 = valu;
        }


        public void SaveAsExcelFile(string path)
        {
            excel.DisplayAlerts = false;
            wb.SaveAs(path);

        }

        public void RunMacro(string name)
        {
            excel.Visible = false;
            excel.Run(name);
        }

        public void FindExcelHyperLink()
        {

            System.Threading.Thread.Sleep(1000);
            links = ws.Hyperlinks;
        }

        public void EditExcelHyperLink(int i, string linkAddress, string texttodisplay, int totalNumberOfBuilds)
        {
            System.Threading.Thread.Sleep(1000);
            links[i + totalNumberOfBuilds].TextToDisplay = texttodisplay;
            links[i + totalNumberOfBuilds].Address = linkAddress;
        }

        public void InsertExcelRow(int rowNumber)
        {
            Range line = (Range)ws.Rows[rowNumber];
            line.Insert();
        }

        public void CopyRowFormat()
        {
            ws = wb.ActiveSheet;
            ws.Rows[50].Copy();

        }

    }
}
