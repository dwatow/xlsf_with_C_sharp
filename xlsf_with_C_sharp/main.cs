using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace XlsFile
{
    class Program
    {
        static void Main(string[] args)
        {
            xlsf excel_file = new xlsf();
            excel_file.NewFile();
            excel_file.SetVisible(true);
            
            excel_file.SelectCell("A1").SetCell("ID Number");
            excel_file.SelectCell("A1").AutoFitWidth();
            
            excel_file.SelectCell("B", 1).SetCell("Current Balance");
            excel_file.SelectCell("B", 1).AutoFitWidth();

            excel_file.SelectCell("A2").SetCellColor(Color.Beige);

            excel_file.SelectCell("A3", "B4").SetCellColor(Color.Black);
            excel_file.MoveSelect(1, 2).SetCell("Black offset 1, 2");
            excel_file.AutoFitHight();
            excel_file.NewSheet();

            Console.WriteLine("SheetTotal:{0}", excel_file.SheetTotal());
            Console.WriteLine("SheetName:{0}", excel_file.GetSheetName());
            excel_file.Quit();
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}
