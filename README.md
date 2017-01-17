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
* MSDN 使用活頁簿Workbooks (等同於檔案操作)
    * (如何：以程式設計方式建立新活頁簿)[https://msdn.microsoft.com/zh-tw/library/x80526fk.aspx]
    * (如何：以程式設計方式開啟活頁簿)[https://msdn.microsoft.com/zh-tw/library/b3k79a5x.aspx]
    * (如何：以程式設計方式關閉活頁簿)[https://msdn.microsoft.com/zh-tw/library/cd8yh918.aspx]
    * (如何：以程式設計方式儲存活頁簿)[https://msdn.microsoft.com/zh-tw/library/h1e33e36.aspx]
* MSDN 使用工作表Workbook
    * (如何：以程式設計方式在活頁簿中加入新的工作表)[https://msdn.microsoft.com/zh-tw/library/6fczc37s.aspx]
    * (如何：以程式設計方式複製工作表)[https://msdn.microsoft.com/zh-tw/library/ms178800.aspx]


# Sample Code
用起來的code會像這樣
```C#=
namespace XlsFile
{
    xlsf CrossTalkFrom;

    CrossTalkFrom.New();  //開新檔案
    CrossTalkFrom.SetSheetName(1, "CrossTalk值");  //新增sheet, 排在第1位(1起始), 命名為 CrossTalk值

    //////////////////////////////////////////////////////////////////////////
    //填值
    //SelectCell: 選擇儲存格, 填入座標
    //SetCell: 填入值

    CrossTalkFrom.SelectCell("B1").SetCell("Lv");
    CrossTalkFrom.SelectCell("F1").SetCell("x");
    CrossTalkFrom.SelectCell("J1").SetCell("y");
    CrossTalkFrom.SelectCell("C2").SetCell(vChain1[0].GetStrLv());
    CrossTalkFrom.SelectCell("B3").SetCell(vChain1[1].GetStrLv());
    CrossTalkFrom.SelectCell("D3").SetCell(vChain1[2].GetStrLv());
    CrossTalkFrom.SelectCell("C4").SetCell(vChain1[3].GetStrLv());
    CrossTalkFrom.SelectCell("C6").SetCell(vChain1[4].GetStrLv());
    CrossTalkFrom.SelectCell("B7").SetCell(vChain1[5].GetStrLv());
    CrossTalkFrom.SelectCell("D7").SetCell(vChain1[6].GetStrLv());
    CrossTalkFrom.SelectCell("C8").SetCell(vChain1[7].GetStrLv());
    CrossTalkFrom.SelectCell("C10").SetCell(vChain1[8].GetStrLv());
    CrossTalkFrom.SelectCell("B11").SetCell(vChain1[9].GetStrLv());
    CrossTalkFrom.SelectCell("D11").SetCell(vChain1[10].GetStrLv());
    CrossTalkFrom.SelectCell("C12").SetCell(vChain1[11].GetStrLv());
    CrossTalkFrom.SelectCell("G2").SetCell(vChain1[0].GetStrSx());
    CrossTalkFrom.SelectCell("F3").SetCell(vChain1[1].GetStrSx());
    CrossTalkFrom.SelectCell("H3").SetCell(vChain1[2].GetStrSx());
    CrossTalkFrom.SelectCell("G4").SetCell(vChain1[3].GetStrSx());
    CrossTalkFrom.SelectCell("G6").SetCell(vChain1[4].GetStrSx());
    CrossTalkFrom.SelectCell("F7").SetCell(vChain1[5].GetStrSx());
    CrossTalkFrom.SelectCell("H7").SetCell(vChain1[6].GetStrSx());
    CrossTalkFrom.SelectCell("G8").SetCell(vChain1[7].GetStrSx());
    CrossTalkFrom.SelectCell("G10").SetCell(vChain1[8].GetStrSx());
    CrossTalkFrom.SelectCell("F11").SetCell(vChain1[9].GetStrSx());
    CrossTalkFrom.SelectCell("H11").SetCell(vChain1[10].GetStrSx());
    CrossTalkFrom.SelectCell("G12").SetCell(vChain1[11].GetStrSx());
    CrossTalkFrom.SelectCell("K2").SetCell(vChain1[0].GetStrSy());
    CrossTalkFrom.SelectCell("J3").SetCell(vChain1[1].GetStrSy());
    CrossTalkFrom.SelectCell("L3").SetCell(vChain1[2].GetStrSy());
    CrossTalkFrom.SelectCell("K4").SetCell(vChain1[3].GetStrSy());
    CrossTalkFrom.SelectCell("K6").SetCell(vChain1[4].GetStrSy());
    CrossTalkFrom.SelectCell("J7").SetCell(vChain1[5].GetStrSy());
    CrossTalkFrom.SelectCell("L7").SetCell(vChain1[6].GetStrSy());
    CrossTalkFrom.SelectCell("K8").SetCell(vChain1[7].GetStrSy());
    CrossTalkFrom.SelectCell("K10").SetCell(vChain1[8].GetStrSy());
    CrossTalkFrom.SelectCell("J11").SetCell(vChain1[9].GetStrSy());
    CrossTalkFrom.SelectCell("L11").SetCell(vChain1[10].GetStrSy());
    CrossTalkFrom.SelectCell("K12").SetCell(vChain1[11].GetStrSy());

    //////////////////////////////////////////////////////////////////////////
    //畫背景和框線
    //SelectCell: 選取儲存格範圍, 英文和數字可以分開填入，並且參數化
    //SetCellColor: 設定底色
    //SetCellBorder: 設定框(填預設值)

    char cCell = 'B';
    int iCell = 2;
    CrossTalkFrom.SelectCell(cCell, iCell   , cCell + 2, iCell + 2).SetCellColor(15).SetCellBorder();
    CrossTalkFrom.SelectCell(cCell, iCell +4, cCell + 2, iCell + 6).SetCellColor(15).SetCellBorder();
    CrossTalkFrom.SelectCell(cCell + 1, iCell +5).SetCellColor(2);
    CrossTalkFrom.SelectCell(cCell, iCell +8, cCell + 2, iCell +10).SetCellColor(15).SetCellBorder();
    CrossTalkFrom.SelectCell(cCell + 1, iCell +9).SetCellColor(1);

    cCell = 'F';
    CrossTalkFrom.SelectCell(cCell, iCell   , cCell + 2, iCell + 2).SetCellColor(15).SetCellBorder();
    CrossTalkFrom.SelectCell(cCell, iCell +4, cCell + 2, iCell + 6).SetCellColor(15).SetCellBorder();
    CrossTalkFrom.SelectCell(cCell + 1, iCell +5).SetCellColor(2);
    CrossTalkFrom.SelectCell(cCell, iCell +8, cCell + 2, iCell +10).SetCellColor(15).SetCellBorder();
    CrossTalkFrom.SelectCell(cCell + 1, iCell +9).SetCellColor(1);

    cCell = 'J';
    CrossTalkFrom.SelectCell(cCell, iCell   , cCell + 2, iCell + 2).SetCellColor(15).SetCellBorder();
    CrossTalkFrom.SelectCell(cCell, iCell +4, cCell + 2, iCell + 6).SetCellColor(15).SetCellBorder();
    CrossTalkFrom.SelectCell(cCell + 1, iCell +5).SetCellColor(2);
    CrossTalkFrom.SelectCell(cCell, iCell +8, cCell + 2, iCell +10).SetCellColor(15).SetCellBorder();
    CrossTalkFrom.SelectCell(cCell + 1, iCell +9).SetCellColor(1);

    //顯示 操控權還給使用者
    CrossTalkFrom.SetVisible(true);
}
```
