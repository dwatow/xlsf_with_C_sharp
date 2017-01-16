﻿using System;
using System.Collections.Generic;
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
            excel_file.Visible(true);
            excel_file.SelectCell("A", 1).SetCell("ID Number");
            excel_file.SelectCell("B", 1).SetCell("Current Balance");

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}
