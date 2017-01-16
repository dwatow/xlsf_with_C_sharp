using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;

namespace XlsFile
{
    class xlsf
    {

        Excel.Application excelApp;
        //Excel.Sheets objSheets;

        Excel._Worksheet objSheet;
        Excel.Range m_Range;//, m_Col, m_Row;
        //Excel.Interior m_Cell;
        //Excel.Font m_Font;
        
        //exception

        public xlsf()
        {
            //NewFile();
        }

        public void NewFile()
        {
            excelApp = new Excel.Application();
            excelApp.Workbooks.Add();
            objSheet = (Excel.Worksheet)excelApp.ActiveSheet;
        }

        public void Visible(bool IsVisible)
        {
            excelApp.Visible = IsVisible;
        }


        public xlsf SelectCell(string X, int Y)
        {
            m_Range = objSheet.Cells[Y, X];
            return this;
        }

        public void SetCell(string CellValue)
        {
            m_Range.Value = CellValue;
        }
    }
}
