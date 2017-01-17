using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

using Excel = Microsoft.Office.Interop.Excel;

namespace XlsFile
{
    class xlsf
    {

        Excel.Application excelApp = new Excel.Application();
        //Excel.Sheets objSheets;

        Excel._Worksheet objSheet;
        //Excel.Range m_Range;, m_Col, m_Row;
        //Excel.Interior m_Cell;
        //Excel.Font m_Font;
        
        //exception

        public xlsf()
        {
            //NewFile();
        }

        public void NewFile()
        {
            excelApp.Workbooks.Add();
            objSheet = (Excel.Worksheet)excelApp.ActiveSheet;
        }

        public void SetVisible(bool IsVisible)
        {
            excelApp.Visible = IsVisible;
        }

        public void AutoFitWidth()
        {
            excelApp.ActiveCell.EntireColumn.AutoFit();
        }

        public void AutoFitHight()
        {
            excelApp.ActiveCell.EntireRow.AutoFit();
        }

        public xlsf SelectCell(string X, int Y) //  ("A", 3)
        {
            objSheet.Cells[Y, X].Select();
            return this;
        }
       

        public xlsf SelectCell(string SelectRange) //("A3") or ("A1:B3")
        {
            objSheet.Range[SelectRange].Select();
            return this;
        }

        public xlsf MoveSelect(int X, int Y)
        {
            excelApp.ActiveCell.Offset[Y, X].Select();
            //get_Offset 是舊版語法
            return this;
        }


        public void SetCell(string CellValue)
        {
            excelApp.ActiveCell.Value = CellValue;
            //Value2是舊版語法
        }

        public xlsf SetCellColor(Color ColorObj)
        {
            excelApp.ActiveCell.Interior.Color = ColorTranslator.ToOle(ColorObj);
            return this;
        }
    }
}
