using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace XlsFile
{
    class Program
    {
        static void Main(string[] args)
        {
            xlsf excel_file = new xlsf();
            //xlsf excel_file = new xlsf(Marshal.GetActiveObject("Excel.Application"));
            
            //excel_file.NewFile();
            excel_file.OpenFile(@"C:\活頁簿1.xlsx");
            
            excel_file.SetVisible(true);
            
            excel_file.SelectCell("A1").SetCell("ID Number");
            excel_file.SelectCell("A1").AutoFitWidth();
            
            excel_file.SelectCell("B", 1).SetCell("Current Balance");
            excel_file.SelectCell("B", 1).AutoNewLine();

            excel_file.SelectCell("A2").SetCellColor(Color.Beige);

            excel_file.SelectCell("A3", "B4").SetCellColor(Color.Black);
            excel_file.OffsetSelectCell(1, 2).SetCell("Black offset 1, 2");
            excel_file.AutoFitHight();
            
            excel_file.NewSheet();
            excel_file.SelectCell("B1").SetCell("this is B1");
            excel_file.SelectCell("B1").CopyCell("B1", "B5");

            excel_file.SelectCell("C1").SetCell("2016/3/16");
            excel_file.SelectCell("C1").SetFont("Console").SetFontSize(42).SetFontColor(Color.Blue).SetCellBk(Color.Orange);
            string datetime = excel_file.SelectCell("C1").GetCell2DateTime().ToString();
            Console.WriteLine(datetime);
            excel_file.SelectCell("D1").SetCell("this is D1");

            Console.WriteLine("SheetTotal:{0}", excel_file.SheetTotal());
            Console.WriteLine("SheetName:{0}", excel_file.GetSheetName());
            excel_file.SelectSheet(@"工作表1").SetSheetName("new sheet name");
            Console.WriteLine("SheetName:{0}", excel_file.GetSheetName());

            excel_file.SelectSheet(@"工作表3").CopySheet();
            excel_file.SelectCellandSetMerge("B5:B7");
            excel_file.SelectCellandSetMerge("B5:B7");//恢復 無Merge

            excel_file.SelectCell("A1").SetCellHeight(50).SetHorztlAlgmet(XlHAlign.xlHAlignLeft).SetCell(1);

            excel_file.SelectCell("A2").SetCell(2);
            excel_file.SelectCell("A3").SetCellHeight(50).SetCell(3);
            excel_file.SetHorztlAlgmet(XlHAlign.xlHAlignJustify).SetVrticlAlgmet(XlVAlign.xlVAlignTop).SetTextAngle(45);


            excel_file.SelectCell("A4").SetCell(4);
            excel_file.SelectCell("A4").SetCellHeight(36).SetCellWidth(68);
            excel_file.SelectCell("A5").SetCell("=SUM(A1:A4)");
            string str = excel_file.SelectCell("A5").GetCell2Str();
            Console.WriteLine(str);
            long number = excel_file.SelectCell("A5").GetCell2Int();
            Console.WriteLine(number);
            excel_file.SelectCell("A5").SetFontBold(true).SetFontStrkthrgh(true);
            excel_file.SelectSheet(2).CopySheet();
            //excel_file.DeleteSheet(2);

            Console.WriteLine("SheetTotal:{0}", excel_file.SheetTotal());
            excel_file.SelectSheet(@"工作表3").MoveSheet();
            //excel_file.SaveAs(@"C:\321.xlsx");
            //excel_file.CloseFile(false);
            //excel_file.Quit();
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}
