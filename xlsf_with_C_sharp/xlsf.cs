using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace XlsFile
{
    class xlsf
    {
        Excel.Application excelApp;
        //Excel.Sheets objSheets;

        //Excel.Interior m_Cell;
        //Excel.Font m_Font;

        //exception

        #region constructor, destructor
        public xlsf()
        {
            excelApp = new Excel.Application();
        }

        public xlsf(object AppProcess)
        {
            //in main.cs
            //xlsf excel_file = new xlsf(Marshal.GetActiveObject("Excel.Application"));
            excelApp = AppProcess as Excel.Application;
        }

        ~xlsf()
        {
            Marshal.FinalReleaseComObject(excelApp);
        }
        #endregion

        #region Workbooks, Workbook
        //Excel._Worksheet objSheet;
        private Excel.Workbooks CurrWorkbooks
        {
            get { return excelApp.Workbooks; }
        }

        private Excel.Workbook CurrWorkbook
        {
            get { return excelApp.ActiveWorkbook; }
        }

        public void NewFile()
        {
            CurrWorkbooks.Add();
        }

        public void OpenFile(string XlsFilePathName)
        {
            CurrWorkbooks.Open(XlsFilePathName);
        }

        public void CloseFile(bool IsCheckSaveFile = true)
        {
            CurrWorkbook.Close(IsCheckSaveFile);
        }

        public void Save()
        {
            CurrWorkbook.Save();
        }

        public void SaveAs(string FilePathName)
        {
            //穩固程式設計
            //以互動方式取消任何儲存或複製活頁簿的方法，都會在程式碼中引發執行階段錯誤。例如，如果您的程序呼叫 SaveAs 方法，但不停用來自 Excel 的提示訊息，而使用者在出現提示時按一下 [取消]，此時 Excel 就會引發執行階段錯誤。
            excelApp.DisplayAlerts = false;
            CurrWorkbook.SaveAs(FilePathName);
            excelApp.DisplayAlerts = true;
        }
        #endregion

        #region Sheet
        private Excel.Sheets CurrSheets
        {
            get { return excelApp.Sheets; }
        }

        private Excel._Worksheet CurrSheet
        {
            get { return excelApp.ActiveSheet; }
        }

        public void NewSheet()
        {
            CurrSheets.Add();
        }

        public void CopySheet()
        {
            CurrSheet.Copy(CurrSheets[SheetTotal()]);
        }

        public void DeleteSheet()
        {
            excelApp.DisplayAlerts = false;
            CurrSheet.Delete();
            excelApp.DisplayAlerts = true;
        }

        public xlsf SelectSheet(int SheetIndex)
        {
            CurrSheets[SheetIndex].Select();
            return this;
        }

        public xlsf SelectSheet(string SheetName)
        {
            CurrSheets[SheetName].Select();
            return this;
        }

        //由SheetNumber 取得SheetName
        public string GetSheetName()
        {
            return CurrSheet.Name;
        }

        public long SheetTotal()
        {
            return CurrSheets.Count;
        }

        #endregion

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

        public void Quit()
        {
            excelApp.Quit();
        }



        public void SetVisible(bool IsVisible)
        {
            excelApp.Visible = IsVisible; //false, 速度快
        }

        #region Cell operator
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
            CurrSheet.Cells[Y, X].Select();
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
        #endregion

    }
}
