# xlsf_with_C_sharp
A class of xls file with C#

想要仿照 (xlsf_with_MFC)[https://github.com/dwatow/xlsf_with_MFC]的class使用方式
做一個C#版本的class

# xls file with C#
xls file with C#

# 聲明
這是一份學C#沒多久，而寫出來的程式碼。

最大目的希望可以讓目前使用C#的人，可以更愉快一點！^^"

希望這個Class可以拋磚引玉，引起大家對於C#使用心得的話題，討論聊C#的使用心得。

# 作用
從C#加入Excel的參考，簡化它的function name

# 使用平台
* OS: Windows 8.1
* Tool: Visual Studio 2013
* Reference: Excel 2013（其他版本沒有測試過）

# 檔案說明
* main.cs Demo功能的主程式
* xlsf.cs 定義檔

## 參考資料
* MSDN
    * (如何：以程式設計方式將色彩套用至 Excel 範圍)[https://msdn.microsoft.com/zh-tw/library/4zs9xy29.aspx]
    * (逐步解說：Office 程式設計 (C# 和 Visual Basic))[https://msdn.microsoft.com/zh-tw/library/ee342218.aspx]
    * (如何：使用 Visual C# 功能存取 Office Interop 物件 (C# 程式設計指南))[https://msdn.microsoft.com/zh-tw/library/dd264733.aspx]
    * (HOW TO：使用 COM Interop 來建立 Excel 試算表 (C# 程式設計手冊))[https://msdn.microsoft.com/zh-tw/library/ms173186(v=vs.80).aspx]
    * (Application.DisplayAlerts 屬性 (Excel))[https://msdn.microsoft.com/zh-tw/library/office/ff839782.aspx]
* MSDN 使用活頁簿Workbooks (等同於檔案操作)
    * (如何：以程式設計方式建立新活頁簿)[https://msdn.microsoft.com/zh-tw/library/x80526fk.aspx]
    * (如何：以程式設計方式開啟活頁簿)[https://msdn.microsoft.com/zh-tw/library/b3k79a5x.aspx]
    * (如何：以程式設計方式關閉活頁簿)[https://msdn.microsoft.com/zh-tw/library/cd8yh918.aspx]
    * (如何：以程式設計方式儲存活頁簿)[https://msdn.microsoft.com/zh-tw/library/h1e33e36.aspx]
* MSDN 使用工作表Workbook
    * (如何：以程式設計方式在活頁簿中加入新的工作表)[https://msdn.microsoft.com/zh-tw/library/6fczc37s.aspx]
    * (如何：以程式設計方式複製工作表)[https://msdn.microsoft.com/zh-tw/library/ms178800.aspx]
    * (如何：以程式設計方式從活頁簿中刪除工作表)[https://msdn.microsoft.com/zh-tw/library/s9kdkks3.aspx]
    * (如何：以程式設計方式選取工作表)[https://msdn.microsoft.com/zh-tw/library/x62t5306.aspx]
    * (如何：以程式設計方式列印工作表)[https://msdn.microsoft.com/zh-tw/library/czhz96h7.aspx]
    * (如何：以程式設計方式在活頁簿內移動工作表)[https://msdn.microsoft.com/zh-tw/library/xyhf0ksb.aspx]
    * (如何：以程式設計方式隱藏工作表)[https://msdn.microsoft.com/zh-tw/library/x0th45dh.aspx]
* MSDN 使用範圍Range 等同於儲存格
    * (如何：以程式設計方式在程式碼中參考工作表範圍)[https://msdn.microsoft.com/zh-tw/library/3a71yzkw.aspx]
    * (如何：以程式設計方式用遞增 (減) 變化的資料自動填滿範圍)[https://msdn.microsoft.com/zh-tw/library/8c94w5fs.aspx]
    * (如何：以程式設計方式在 Excel 範圍中儲存和擷取日期值)[https://msdn.microsoft.com/zh-tw/library/1ad4d8d6.aspx]
    * (如何：以程式設計方式將樣式套用至活頁簿中的範圍)[https://msdn.microsoft.com/zh-tw/library/f1hh9fza.aspx]


# Sample Code
用起來的code會像這樣
```Cs
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
            excel_file.SelectCell("B", 1).AutoFitWidth();

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
            excel_file.SelectCell("A1").SetCell(1);
            excel_file.SelectCell("A2").SetCell(2);
            excel_file.SelectCell("A3").SetCell(3);
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
            //excel_file.SaveAs(@"C:\Users\kgs_chris\Desktop\321.xlsx");
            //excel_file.CloseFile(false);
            //excel_file.Quit();
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}
```
