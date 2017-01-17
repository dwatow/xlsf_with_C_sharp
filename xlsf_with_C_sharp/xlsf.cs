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

        //Excel.Interior m_Cell;
        //Excel.Font m_Font;
        
        //exception

        //Excel._Worksheet objSheet;
        private Excel.Workbooks CurrWorkbooks
        {
            get { return excelApp.Workbooks; }
        }
        
        private Excel.Workbook CurrWorkbook
        {
            get { return excelApp.ActiveWorkbook; }
        }
        private Excel.Sheets CurrSheets
        {
            get { return excelApp.ActiveWorkbook.Sheets; }
        }

        private Excel._Worksheet CurrSheet
        {
            get { return excelApp.ActiveSheet; }
        }

        //Excel.Range m_Range;, m_Col, m_Row;
        private Excel.Range CurrCell
        {
            get { return excelApp.ActiveCell; }
        }

        private Excel.Range CurrColumnCells
        {
            get { return CurrCell.EntireColumn; }
        }

        private Excel.Range CurrRowCells
        {
            get { return CurrCell.EntireRow; }
        }



        public void NewFile()
        {
            NewBook();
        }

        public void NewBook()
        {
           CurrWorkbooks.Add();
        }

        public void NewSheet()
        {
            CurrSheets.Add(); //往前加，要換成往後加
        }

        public void SetVisible(bool IsVisible)
        {
            excelApp.Visible = IsVisible;
        }

        public void AutoFitWidth()
        {
            CurrColumnCells.AutoFit();
        }

        public void AutoFitHight()
        {
            CurrRowCells.AutoFit();
        }

        #region Select Cell
        public xlsf SelectCell(string X, int Y) //  ("A", 3)
        {
            CurrSheet.Range[Y, X].Select();
            return this;
        }

        public xlsf SelectCell(string CellPosition1, string CellPosition2) //  ("A1", "B3")
        {
            CurrSheet.get_Range(CellPosition1, CellPosition2).Select();
            return this;
        }

        public xlsf SelectCell(string SelectRange) //("A3") or ("A1:B3")
        {
            CurrSheet.Range[SelectRange].Select();
            return this;
        }
        #endregion

        public xlsf MoveSelect(int X, int Y)
        {
            CurrCell.Offset[Y, X].Select();
            //get_Offset 是舊版語法
            return this;
        }


        public void SetCell(string CellValue)
        {
            CurrCell.Value = CellValue;
            //Value2是舊版語法
        }

        public xlsf SetCellColor(Color ColorObj)
        {
            CurrCell.Interior.Color = ColorTranslator.ToOle(ColorObj);
            return this;
        }
    }
}
