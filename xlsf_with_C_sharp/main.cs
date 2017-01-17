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

            excel_file.NewSheet();

            //Console.WriteLine("Press any key to exit.");
            //Console.ReadKey();
        }
    }
}
