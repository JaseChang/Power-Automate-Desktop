財報重訊公告稿製作
===

流程
---
1. 執行Main:
   - 匯入Pdf檔: 如事務所提供有浮水印字樣，需請事務所將浮水印清除
   - 輸入BS及IS頁數
   - 輸入董事會日期
   - 輸入`季度`
2. 依`季度`執行Subflow:
   - Q1
   - Q2
   - Q3
   - Q4

程式碼
---

<details>
  
<summary>Main </summary>
  
```
@@statistics_TextBlock: '1'
@@statistics_Input_File: '1'
@@statistics_Input_Text: '3'
@@statistics_Action_Submit: '1'
Display.ShowCustomDialog CardTemplateJson: '''{
  \"type\": \"AdaptiveCard\",
  \"version\": \"1.4\",
  \"id\": \"AdaptiveCard\",
  \"body\": [
    {
      \"type\": \"TextBlock\",
      \"id\": \"Title\",
      \"size\": \"large\",
      \"weight\": \"bolder\",
      \"text\": \"${Title_Text}\"
    },
    {
      \"type\": \"Input.File\",
      \"id\": \"PDF\",
      \"isRequired\": true,
      \"label\": \"${PDF_Label}\",
      \"errorMessage\": \"${PDF_ErrorMessage}\"
    },
    {
      \"type\": \"Input.Text\",
      \"id\": \"PageNo\",
      \"isRequired\": true,
      \"label\": \"${PageNo_Label}\"
    },
    {
      \"type\": \"Input.Text\",
      \"id\": \"BOD input\",
      \"isRequired\": true,
      \"label\": \"${BOD_input_Label}\"
    },
    {
      \"type\": \"Input.Text\",
      \"id\": \"Period input\",
      \"isRequired\": true,
      \"label\": \"${Period_input_Label}\"
    }
  ],
  \"actions\": [
    {
      \"type\": \"Action.Submit\",
      \"id\": \"Submit\",
      \"title\": \"${Submit_Title}\"
    }
  ],
  \"FormTitle\": \"${AdaptiveCard_FormTitle}\"
}''' CustomFormData=> CustomFormData ButtonPressed=> ButtonPressed @AdaptiveCard_FormTitle: $'''製作財報公告稿''' @Title_Text: $'''Create new PDF from selected PDF pages''' @PDF_Label: $'''Select the PDF''' @PDF_ErrorMessage: $'''Document not found''' @PageNo_Label: $'''Page number(s) e.g. 1,3,17-24: ''' @BOD_input_Label: $'''董事會日期:11x/xx/xx''' @Period_input_Label: $'''第x季''' @Submit_Title: $'''Extract'''
IF CustomFormData['Period input'] = 1 THEN
    SET Q1 TO $'''113/01/01~113/03/31'''
    CALL Q1
END
IF CustomFormData['Period input'] = 2 THEN
    SET Q2 TO $'''113/01/01~113/06/30'''
    CALL Q2
END
IF CustomFormData['Period input'] = 3 THEN
    SET Q3 TO $'''113/01/01~113/09/30'''
    CALL Q3
END
IF CustomFormData['Period input'] = 4 THEN
    SET Q4 TO $'''113/01/01~113/12/31'''
    CALL Q4
END
```
</details>
  
<details>
<summary>Subflow Q1</summary>

```
Variables.CreateNewList List=> List
Variables.AddItemToList Item: $'''資產總計''' List: List
Variables.AddItemToList Item: $'''負債總計''' List: List
Variables.AddItemToList Item: $'''歸屬母公司業主之權益小計''' List: List
Variables.AddItemToList Item: $'''營業收入''' List: List
Variables.AddItemToList Item: $'''營業毛利''' List: List
Variables.AddItemToList Item: $'''營業淨利''' List: List
Variables.AddItemToList Item: $'''稅前淨利(損)''' List: List
Variables.AddItemToList Item: $'''本期淨利(損)''' List: List
Variables.AddItemToList Item: $'''母公司業主''' List: List
Variables.AddItemToList Item: $'''基本每股盈餘(虧損)''' List: List
IF CustomFormData['Period input'] = 1 THEN
    IF ButtonPressed <> $'''Cancel''' THEN
        Display.SelectFolder Description: $'''Select a folder to save the new PDF file...''' IsTopMost: True SelectedFolder=> DestinationFolder ButtonPressed=> ButtonPressed3
        Pdf.ExtractPages PDFFile: CustomFormData['PDF'] PageSelection: CustomFormData['PageNo'] ExtractedPDFPath: $'''%DestinationFolder%/NewPDFfile.pdf''' IfFileExists: Pdf.IfFileExists.AddSequentialSuffix ExtractedPDFFile=> ExtractedPDF
        Pdf.ExtractTablesFromPDF.ExtractTables PDFFile: ExtractedPDF MultiPageTables: True SetFirstRowAsHeader: True ExtractedPDFTables=> ExtractedPDFTables
        File.Delete Files: ExtractedPDF
        Excel.LaunchExcel.LaunchUnderExistingProcess Visible: True Instance=> ExcelInstance
        SET i TO 0
        Excel.RenameWorksheet.RenameWorksheetWithName Instance: ExcelInstance Name: $'''工作表1''' NewName: $'''Table %i + 1%'''
        LOOP FOREACH CurrentItem IN ExtractedPDFTables
            Excel.WriteToExcel.Write Instance: ExcelInstance Value: ExtractedPDFTables[i].DataTable
            Variables.IncreaseVariable Value: i IncrementValue: 1
            Excel.AddWorksheet Instance: ExcelInstance Name: $'''Table %i + 1%''' WorksheetPosition: Excel.WorksheetPosition.Last
        END
        Excel.SetActiveWorksheet.ActivateWorksheetByName Instance: ExcelInstance Name: $'''Table 3'''
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: List Column: $'''A''' Row: 1
        Excel.SetActiveWorksheet.ActivateWorksheetByName Instance: ExcelInstance Name: $'''Table 3'''
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(A1,\'Table 1\'!B1:D45,3,0)''' Column: $'''B''' Row: 1
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(A2,\'Table 1\'!K1:M45,3,0)''' Column: $'''B''' Row: 2
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(A3,\'Table 1\'!K1:M45,3,0)''' Column: $'''B''' Row: 3
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(\"*營業收入*\",\'Table 2\'!b1:D45,3,0)''' Column: $'''B''' Row: 4
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(A5,\'Table 2\'!B1:D45,3,0)''' Column: $'''B''' Row: 5
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=IFERROR(VLOOKUP(\"*營業淨利*\",\'Table 2\'!B1:D45,3,0),VLOOKUP(\"*營業淨利(損)*\",\'Table 2\'!B1:D45,3,0))''' Column: $'''B''' Row: 6
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=IFERROR(VLOOKUP(\"*稅前淨利*\",\'Table 2\'!B1:D45,3,0),VLOOKUP(\"*稅前淨利(損)*\",\'Table 2\'!B1:D45,3,0))''' Column: $'''B''' Row: 7
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=IFERROR(VLOOKUP(\"*本期淨利*\",\'Table 2\'!B1:D45,3,0),VLOOKUP(\"*本期淨利(損)*\",\'Table 2\'!B1:D45,3,0))''' Column: $'''B''' Row: 8
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(A9,\'Table 2\'!B1:D45,3,0)''' Column: $'''B''' Row: 9
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=IFERROR(VLOOKUP(\"*基本每股盈餘*\",\'Table 2\'!B1:E45,4,0),VLOOKUP(\"*基本每股盈餘(虧損)*\",\'Table 2\'!B1:E45,4,0))''' Column: $'''B''' Row: 10
        DISABLE Excel.GetFirstFreeColumnRow Instance: `%ExcelInstance[Table 1]%` FirstFreeColumn=> FirstFreeColumn FirstFreeRow=> FirstFreeRow
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 1 ReadAsText: False CellValue=> Assets
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 2 ReadAsText: False CellValue=> Liability
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 3 ReadAsText: False CellValue=> ParentEquity
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 4 ReadAsText: False CellValue=> Sales
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 5 ReadAsText: False CellValue=> Grossmargin
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 6 ReadAsText: False CellValue=> Operatingincome
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 7 ReadAsText: False CellValue=> EBIT
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 8 ReadAsText: False CellValue=> Profit
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 9 ReadAsText: False CellValue=> ParentProfit
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 10 ReadAsText: False CellValue=> EPS
        Excel.CloseExcel.Close Instance: ExcelInstance
        Word.LaunchWord.Launch Visible: True Instance=> WordInstance
        Word.WriteToWord.WriteStartOfDocument Instance: WordInstance Text: $'''公司代號:
序號:1
主旨:113年第%CustomFormData['Period input']%季合併財務報告業經提報董事會通過
符合條款-第四條第XX款:31
事實發生日:%CustomFormData['BOD input']%
內容:
1.提報董事會或經董事會決議日期: %CustomFormData['BOD input']%
2.審計委員會通過日期: %CustomFormData['BOD input']%
3.財務報告或年度自結財務資訊報導期間
起訖日期(XXX/XX/XX~XXX/XX/XX):%Q1%
4.1月1日累計至本期止營業收入(仟元):%Sales%
5.1月1日累計至本期止營業毛利(毛損) (仟元):%Grossmargin%
6.1月1日累計至本期止營業利益(損失) (仟元):%Operatingincome%
7.1月1日累計至本期止稅前淨利(淨損) (仟元):%EBIT%
8.1月1日累計至本期止本期淨利(淨損) (仟元):%Profit%
9.1月1日累計至本期止歸屬於母公司業主淨利(損) (仟元):%ParentProfit%
10.1月1日累計至本期止基本每股盈餘(損失) (元):%EPS%
11.期末總資產(仟元):%Assets%
12.期末總負債(仟元):%Liability%
13.期末歸屬於母公司業主之權益(仟元):%ParentEquity%
14.其他應敘明事項:無''' AppendNewLine: False
        Word.ReadFromWord.Read Instance: WordInstance WordData=> WordData
        Word.CloseWord.Close Instance: WordInstance
        System.RunApplication.RunApplication ApplicationPath: $'''C:\\Windows\\System32\\notepad.exe''' WindowStyle: System.ProcessWindowStyle.Normal ProcessId=> AppProcessId
        UIAutomation.PopulateTextField.PopulateTextField TextField: appmask['Window \'未命名 - 記事本\'']['Document \'文字編輯器\''] Text: WordData Mode: UIAutomation.PopulateTextMode.Append ClickType: UIAutomation.PopulateMouseClickType.SingleClick
    END
END

# [ControlRepository][PowerAutomateDesktop]

{
  "ControlRepositorySymbols": [
    {
      "IgnoreImagesOnSerialization": false,
      "Repository": "{\r\n  \"Screens\": [\r\n    {\r\n      \"Controls\": [\r\n        {\r\n          \"AutomationProtocol\": \"uia3\",\r\n          \"ScreenShot\": null,\r\n          \"ElementTypeName\": \"Document\",\r\n          \"InstanceId\": \"44ef9488-65c5-4673-be46-aa9c5a02595d\",\r\n          \"Name\": \"Document '文字編輯器'\",\r\n          \"SelectorCount\": 1,\r\n          \"Selectors\": [\r\n            {\r\n              \"CustomSelector\": null,\r\n              \"Elements\": [\r\n                {\r\n                  \"Attributes\": [\r\n                    {\r\n                      \"Ignore\": false,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Class\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": \"Edit\"\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Enabled\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": true\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Id\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": \"15\"\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Name\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": \"文字編輯器\"\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": true,\r\n                      \"Name\": \"Ordinal\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": 0\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Password\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": false\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Visible\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": true\r\n                    }\r\n                  ],\r\n                  \"CustomValue\": null,\r\n                  \"Ignore\": false,\r\n                  \"Name\": \"Document '文字編輯器'\",\r\n                  \"Tag\": \"document\"\r\n                }\r\n              ],\r\n              \"Ignore\": false,\r\n              \"ImageSelector\": null,\r\n              \"IsCustom\": false,\r\n              \"IsImageBased\": false,\r\n              \"IsWindowsInstance\": false,\r\n              \"Name\": \"Default Selector\",\r\n              \"Properties\": []\r\n            }\r\n          ],\r\n          \"Tag\": \"document\",\r\n          \"ScreenshotPath\": \"controlRepo-screenshots\\\\da229120-5e67-424f-a839-58d4b9834b48.png\"\r\n        }\r\n      ],\r\n      \"Handle\": {\r\n        \"value\": 0\r\n      },\r\n      \"ProcessName\": null,\r\n      \"ScreenShot\": null,\r\n      \"ElementTypeName\": \"Window\",\r\n      \"InstanceId\": \"2a2c0634-d743-4111-b145-677d39e8c67c\",\r\n      \"Name\": \"Window '未命名 - 記事本'\",\r\n      \"SelectorCount\": 1,\r\n      \"Selectors\": [\r\n        {\r\n          \"CustomSelector\": null,\r\n          \"Elements\": [\r\n            {\r\n              \"Attributes\": [\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Class\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": \"Notepad\"\r\n                },\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Enabled\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": true\r\n                },\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Id\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": \"\"\r\n                },\r\n                {\r\n                  \"Ignore\": false,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Name\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": \"未命名 - 記事本\"\r\n                },\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": true,\r\n                  \"Name\": \"Ordinal\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": 0\r\n                },\r\n                {\r\n                  \"Ignore\": false,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Process\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": \"notepad\"\r\n                },\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Visible\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": true\r\n                }\r\n              ],\r\n              \"CustomValue\": null,\r\n              \"Ignore\": false,\r\n              \"Name\": \"Window '未命名 - 記事本'\",\r\n              \"Tag\": \"window\"\r\n            }\r\n          ],\r\n          \"Ignore\": false,\r\n          \"ImageSelector\": null,\r\n          \"IsCustom\": false,\r\n          \"IsImageBased\": false,\r\n          \"IsWindowsInstance\": false,\r\n          \"Name\": \"Default Selector\",\r\n          \"Properties\": []\r\n        }\r\n      ],\r\n      \"Tag\": \"window\",\r\n      \"ScreenshotPath\": \"controlRepo-screenshots\\\\82013743-3a5e-4168-94fe-1a3ca398ed1c.png\"\r\n    }\r\n  ],\r\n  \"Version\": 1\r\n}",
      "ImportMetadata": {
        "DisplayName": "Computer",
        "ConnectionString": "",
        "Type": "Local",
        "DesktopType": "local"
      },
      "Name": "appmask"
    }
  ],
  "ImageRepositorySymbol": {
    "Repository": "{\r\n  \"Folders\": [],\r\n  \"Images\": [],\r\n  \"Version\": 1\r\n}",
    "ImportMetadata": {},
    "Name": "imgrepo"
  },
  "ConnectionReferences": []
}
```
</details>

<details>
<summary>Subflow Q2</summary>

```
Variables.CreateNewList List=> List
Variables.AddItemToList Item: $'''資產總計''' List: List
Variables.AddItemToList Item: $'''負債總計''' List: List
Variables.AddItemToList Item: $'''歸屬母公司業主之權益小計''' List: List
Variables.AddItemToList Item: $'''營業收入''' List: List
Variables.AddItemToList Item: $'''營業毛利''' List: List
Variables.AddItemToList Item: $'''營業淨利''' List: List
Variables.AddItemToList Item: $'''稅前淨利(損)''' List: List
Variables.AddItemToList Item: $'''本期淨利(損)''' List: List
Variables.AddItemToList Item: $'''母公司業主''' List: List
Variables.AddItemToList Item: $'''基本每股盈餘(虧損)''' List: List
IF CustomFormData['Period input'] = 2 THEN
    IF ButtonPressed <> $'''Cancel''' THEN
        Display.SelectFolder Description: $'''Select a folder to save the new PDF file...''' IsTopMost: True SelectedFolder=> DestinationFolder ButtonPressed=> ButtonPressed3
        Pdf.ExtractPages PDFFile: CustomFormData['PDF'] PageSelection: CustomFormData['PageNo'] ExtractedPDFPath: $'''%DestinationFolder%/NewPDFfile.pdf''' IfFileExists: Pdf.IfFileExists.AddSequentialSuffix ExtractedPDFFile=> ExtractedPDF
        Pdf.ExtractTablesFromPDF.ExtractTables PDFFile: ExtractedPDF MultiPageTables: False SetFirstRowAsHeader: False ExtractedPDFTables=> ExtractedPDFTables
        File.Delete Files: ExtractedPDF
        Excel.LaunchExcel.LaunchUnderExistingProcess Visible: True Instance=> ExcelInstance
        SET i TO 0
        Excel.RenameWorksheet.RenameWorksheetWithName Instance: ExcelInstance Name: $'''工作表1''' NewName: $'''Table %i + 1%'''
        LOOP FOREACH CurrentItem IN ExtractedPDFTables
            Excel.WriteToExcel.Write Instance: ExcelInstance Value: ExtractedPDFTables[i].DataTable
            Variables.IncreaseVariable Value: i IncrementValue: 1
            Excel.AddWorksheet Instance: ExcelInstance Name: $'''Table %i + 1%''' WorksheetPosition: Excel.WorksheetPosition.Last
        END
        Excel.SetActiveWorksheet.ActivateWorksheetByName Instance: ExcelInstance Name: $'''Table 3'''
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: List Column: $'''A''' Row: 1
        Excel.SetActiveWorksheet.ActivateWorksheetByName Instance: ExcelInstance Name: $'''Table 3'''
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(A1,\'Table 1\'!B1:D40,3,0)''' Column: $'''B''' Row: 1
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(A2,\'Table 1\'!K1:M40,3,0)''' Column: $'''B''' Row: 2
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(A3,\'Table 1\'!K1:M40,3,0)''' Column: $'''B''' Row: 3
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(\"*營業收入*\",\'Table 2\'!b1:L45,7,0)''' Column: $'''B''' Row: 4
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(\"*營業毛利*\",\'Table 2\'!B1:L45,7,0)''' Column: $'''B''' Row: 5
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=IFERROR(VLOOKUP(\"*營業淨利*\",\'Table 2\'!B1:L45,7,0),VLOOKUP(\"*營業淨利(損)*\",\'Table 2\'!B1:L45,7,0))''' Column: $'''B''' Row: 6
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=IFERROR(VLOOKUP(\"稅前淨利\",\'Table 2\'!B1:L45,7,0),VLOOKUP(\"稅前淨利(損)\",\'Table 2\'!B1:L45,7,0))''' Column: $'''B''' Row: 7
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=IFERROR(VLOOKUP(\"*本期淨利\",\'Table 2\'!B1:L45,7,0),VLOOKUP(\"*本期淨利(損)\",\'Table 2\'!B1:L45,7,0))''' Column: $'''B''' Row: 8
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(\"*母公司業主*\",\'Table 2\'!B1:L45,7,0)''' Column: $'''B''' Row: 9
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=IFERROR(VLOOKUP(\"基本每股盈餘\",\'Table 2\'!B1:L45,8,0),VLOOKUP(\"基本每股盈餘(虧損)\",\'Table 2\'!B1:L45,8,0))''' Column: $'''B''' Row: 10
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 1 ReadAsText: False CellValue=> Assets
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 2 ReadAsText: False CellValue=> Liability
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 3 ReadAsText: False CellValue=> ParentEquity
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 4 ReadAsText: False CellValue=> Sales
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 5 ReadAsText: False CellValue=> Grossmargin
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 6 ReadAsText: False CellValue=> Operatingincome
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 7 ReadAsText: False CellValue=> EBIT
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 8 ReadAsText: False CellValue=> Profit
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 9 ReadAsText: False CellValue=> ParentProfit
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 10 ReadAsText: False CellValue=> EPS
        Text.FromNumber Number: Assets DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> Assets2
        Text.FromNumber Number: Liability DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> Liability2
        Text.FromNumber Number: ParentEquity DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> ParentEquity2
        Text.FromNumber Number: Sales DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> Sale2
        Text.FromNumber Number: Grossmargin DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> Grossmargin2
        Text.FromNumber Number: Operatingincome DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> Operatingincome2
        Text.FromNumber Number: EBIT DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> EBIT2
        Text.FromNumber Number: Profit DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> Profit2
        Text.FromNumber Number: ParentProfit DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> ParentProfit2
        Excel.CloseExcel.Close Instance: ExcelInstance
        Word.LaunchWord.Launch Visible: True Instance=> WordInstance
        Word.WriteToWord.WriteStartOfDocument Instance: WordInstance Text: $'''公司代號:
序號:1
主旨:113年第%CustomFormData['Period input']%季合併財務報告業經提報董事會通過
符合條款-第四條第XX款:31
事實發生日:%CustomFormData['BOD input']%
內容:
1.提報董事會或經董事會決議日期: %CustomFormData['BOD input']%
2.審計委員會通過日期: %CustomFormData['BOD input']%
3.財務報告或年度自結財務資訊報導期間
起訖日期(XXX/XX/XX~XXX/XX/XX):%Q3%
4.1月1日累計至本期止營業收入(仟元):%Sale2%
5.1月1日累計至本期止營業毛利(毛損) (仟元):%Grossmargin2%
6.1月1日累計至本期止營業利益(損失) (仟元):%Operatingincome2%
7.1月1日累計至本期止稅前淨利(淨損) (仟元):%EBIT2%
8.1月1日累計至本期止本期淨利(淨損) (仟元):%Profit2%
9.1月1日累計至本期止歸屬於母公司業主淨利(損) (仟元):%ParentProfit2%
10.1月1日累計至本期止基本每股盈餘(損失) (元):%EPS%
11.期末總資產(仟元):%Assets2%
12.期末總負債(仟元):%Liability2%
13.期末歸屬於母公司業主之權益(仟元):%ParentEquity2%
14.其他應敘明事項:無''' AppendNewLine: False
        Word.ReadFromWord.Read Instance: WordInstance WordData=> WordData
        Word.CloseWord.Close Instance: WordInstance
        System.RunApplication.RunApplication ApplicationPath: $'''C:\\Windows\\System32\\notepad.exe''' WindowStyle: System.ProcessWindowStyle.Normal ProcessId=> AppProcessId
        UIAutomation.PopulateTextField.PopulateTextField TextField: appmask['Window \'未命名 - 記事本\'']['Document \'文字編輯器\''] Text: WordData Mode: UIAutomation.PopulateTextMode.Append ClickType: UIAutomation.PopulateMouseClickType.SingleClick
        DISABLE Display.ShowMessageDialog.ShowMessage Title: $'''財報公告檔''' Message: $'''公司代號:3380
序號:1
主旨:11x年第x季合併財務報告業經提報董事會通過
符合條款-第四條第XX款:31
事實發生日:%CustomFormData['BOD input']%
內容:
1.提報董事會或經董事會決議日期: %CustomFormData['BOD input']%
2.審計委員會通過日期: %CustomFormData['BOD input']%
3.財務報告或年度自結財務資訊報導期間
起訖日期(XXX/XX/XX~XXX/XX/XX):%CustomFormData['Date input']%
4.1月1日累計至本期止營業收入(仟元):%Sales%
5.1月1日累計至本期止營業毛利(毛損) (仟元):%Grossmargin%
6.1月1日累計至本期止營業利益(損失) (仟元):%Operatingincome%
7.1月1日累計至本期止稅前淨利(淨損) (仟元):%EBIT%
8.1月1日累計至本期止本期淨利(淨損) (仟元):%Profit%
9.1月1日累計至本期止歸屬於母公司業主淨利(損) (仟元):%ParentProfit%
10.1月1日累計至本期止基本每股盈餘(損失) (元):%EPS%
11.期末總資產(仟元):%Assets%
12.期末總負債(仟元):%Liability%
13.期末歸屬於母公司業主之權益(仟元):%ParentEquity%
14.其他應敘明事項:無''' Icon: Display.Icon.Information Buttons: Display.Buttons.OK DefaultButton: Display.DefaultButton.Button1 IsTopMost: False ButtonPressed=> ButtonPressed2
    END
END

# [ControlRepository][PowerAutomateDesktop]

{
  "ControlRepositorySymbols": [
    {
      "IgnoreImagesOnSerialization": false,
      "Repository": "{\r\n  \"Screens\": [\r\n    {\r\n      \"Controls\": [\r\n        {\r\n          \"AutomationProtocol\": \"uia3\",\r\n          \"ScreenShot\": null,\r\n          \"ElementTypeName\": \"Document\",\r\n          \"InstanceId\": \"44ef9488-65c5-4673-be46-aa9c5a02595d\",\r\n          \"Name\": \"Document '文字編輯器'\",\r\n          \"SelectorCount\": 1,\r\n          \"Selectors\": [\r\n            {\r\n              \"CustomSelector\": null,\r\n              \"Elements\": [\r\n                {\r\n                  \"Attributes\": [\r\n                    {\r\n                      \"Ignore\": false,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Class\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": \"Edit\"\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Enabled\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": true\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Id\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": \"15\"\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Name\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": \"文字編輯器\"\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": true,\r\n                      \"Name\": \"Ordinal\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": 0\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Password\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": false\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Visible\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": true\r\n                    }\r\n                  ],\r\n                  \"CustomValue\": null,\r\n                  \"Ignore\": false,\r\n                  \"Name\": \"Document '文字編輯器'\",\r\n                  \"Tag\": \"document\"\r\n                }\r\n              ],\r\n              \"Ignore\": false,\r\n              \"ImageSelector\": null,\r\n              \"IsCustom\": false,\r\n              \"IsImageBased\": false,\r\n              \"IsWindowsInstance\": false,\r\n              \"Name\": \"Default Selector\",\r\n              \"Properties\": []\r\n            }\r\n          ],\r\n          \"Tag\": \"document\",\r\n          \"ScreenshotPath\": \"controlRepo-screenshots\\\\da229120-5e67-424f-a839-58d4b9834b48.png\"\r\n        }\r\n      ],\r\n      \"Handle\": {\r\n        \"value\": 0\r\n      },\r\n      \"ProcessName\": null,\r\n      \"ScreenShot\": null,\r\n      \"ElementTypeName\": \"Window\",\r\n      \"InstanceId\": \"2a2c0634-d743-4111-b145-677d39e8c67c\",\r\n      \"Name\": \"Window '未命名 - 記事本'\",\r\n      \"SelectorCount\": 1,\r\n      \"Selectors\": [\r\n        {\r\n          \"CustomSelector\": null,\r\n          \"Elements\": [\r\n            {\r\n              \"Attributes\": [\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Class\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": \"Notepad\"\r\n                },\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Enabled\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": true\r\n                },\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Id\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": \"\"\r\n                },\r\n                {\r\n                  \"Ignore\": false,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Name\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": \"未命名 - 記事本\"\r\n                },\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": true,\r\n                  \"Name\": \"Ordinal\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": 0\r\n                },\r\n                {\r\n                  \"Ignore\": false,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Process\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": \"notepad\"\r\n                },\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Visible\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": true\r\n                }\r\n              ],\r\n              \"CustomValue\": null,\r\n              \"Ignore\": false,\r\n              \"Name\": \"Window '未命名 - 記事本'\",\r\n              \"Tag\": \"window\"\r\n            }\r\n          ],\r\n          \"Ignore\": false,\r\n          \"ImageSelector\": null,\r\n          \"IsCustom\": false,\r\n          \"IsImageBased\": false,\r\n          \"IsWindowsInstance\": false,\r\n          \"Name\": \"Default Selector\",\r\n          \"Properties\": []\r\n        }\r\n      ],\r\n      \"Tag\": \"window\",\r\n      \"ScreenshotPath\": \"controlRepo-screenshots\\\\82013743-3a5e-4168-94fe-1a3ca398ed1c.png\"\r\n    }\r\n  ],\r\n  \"Version\": 1\r\n}",
      "ImportMetadata": {
        "DisplayName": "Computer",
        "ConnectionString": "",
        "Type": "Local",
        "DesktopType": "local"
      },
      "Name": "appmask"
    }
  ],
  "ImageRepositorySymbol": {
    "Repository": "{\r\n  \"Folders\": [],\r\n  \"Images\": [],\r\n  \"Version\": 1\r\n}",
    "ImportMetadata": {},
    "Name": "imgrepo"
  },
  "ConnectionReferences": []
}
```
</details>

<details>
<summary>Subflow Q3</summary>

```
Variables.CreateNewList List=> List
Variables.AddItemToList Item: $'''資產總計''' List: List
Variables.AddItemToList Item: $'''負債總計''' List: List
Variables.AddItemToList Item: $'''歸屬母公司業主之權益小計''' List: List
Variables.AddItemToList Item: $'''營業收入''' List: List
Variables.AddItemToList Item: $'''營業毛利''' List: List
Variables.AddItemToList Item: $'''營業淨利''' List: List
Variables.AddItemToList Item: $'''稅前淨利(損)''' List: List
Variables.AddItemToList Item: $'''本期淨利(損)''' List: List
Variables.AddItemToList Item: $'''母公司業主''' List: List
Variables.AddItemToList Item: $'''基本每股盈餘(虧損)''' List: List
IF CustomFormData['Period input'] = 3 THEN
    IF ButtonPressed <> $'''Cancel''' THEN
        Display.SelectFolder Description: $'''Select a folder to save the new PDF file...''' IsTopMost: True SelectedFolder=> DestinationFolder ButtonPressed=> ButtonPressed3
        Pdf.ExtractPages PDFFile: CustomFormData['PDF'] PageSelection: CustomFormData['PageNo'] ExtractedPDFPath: $'''%DestinationFolder%/NewPDFfile.pdf''' IfFileExists: Pdf.IfFileExists.AddSequentialSuffix ExtractedPDFFile=> ExtractedPDF
        Pdf.ExtractTablesFromPDF.ExtractTables PDFFile: ExtractedPDF MultiPageTables: False SetFirstRowAsHeader: False ExtractedPDFTables=> ExtractedPDFTables
        File.Delete Files: ExtractedPDF
        Excel.LaunchExcel.LaunchUnderExistingProcess Visible: True Instance=> ExcelInstance
        SET i TO 0
        Excel.RenameWorksheet.RenameWorksheetWithName Instance: ExcelInstance Name: $'''工作表1''' NewName: $'''Table %i + 1%'''
        LOOP FOREACH CurrentItem IN ExtractedPDFTables
            Excel.WriteToExcel.Write Instance: ExcelInstance Value: ExtractedPDFTables[i].DataTable
            Variables.IncreaseVariable Value: i IncrementValue: 1
            Excel.AddWorksheet Instance: ExcelInstance Name: $'''Table %i + 1%''' WorksheetPosition: Excel.WorksheetPosition.Last
        END
        Excel.SetActiveWorksheet.ActivateWorksheetByName Instance: ExcelInstance Name: $'''Table 3'''
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: List Column: $'''A''' Row: 1
        Excel.SetActiveWorksheet.ActivateWorksheetByName Instance: ExcelInstance Name: $'''Table 3'''
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(A1,\'Table 1\'!B1:D40,3,0)''' Column: $'''B''' Row: 1
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(A2,\'Table 1\'!K1:M40,3,0)''' Column: $'''B''' Row: 2
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(A3,\'Table 1\'!K1:M40,3,0)''' Column: $'''B''' Row: 3
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(\"*營業收入*\",\'Table 2\'!b1:L45,7,0)''' Column: $'''B''' Row: 4
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(\"*營業毛利*\",\'Table 2\'!B1:L45,7,0)''' Column: $'''B''' Row: 5
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(\"*營業淨利*\",\'Table 2\'!B1:L45,7,0)''' Column: $'''B''' Row: 6
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(\"*稅前淨利(損)*\",\'Table 2\'!B1:L45,7,0)''' Column: $'''B''' Row: 7
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(\"*本期淨利(損)*\",\'Table 2\'!B1:L45,7,0)''' Column: $'''B''' Row: 8
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(\"*母公司業主*\",\'Table 2\'!B1:L45,7,0)''' Column: $'''B''' Row: 9
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(\"*基本每股盈餘(虧損)*\",\'Table 2\'!B1:L45,8,0)''' Column: $'''B''' Row: 10
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 1 ReadAsText: False CellValue=> Assets
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 2 ReadAsText: False CellValue=> Liability
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 3 ReadAsText: False CellValue=> ParentEquity
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 4 ReadAsText: False CellValue=> Sales
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 5 ReadAsText: False CellValue=> Grossmargin
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 6 ReadAsText: False CellValue=> Operatingincome
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 7 ReadAsText: False CellValue=> EBIT
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 8 ReadAsText: False CellValue=> Profit
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 9 ReadAsText: False CellValue=> ParentProfit
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 10 ReadAsText: False CellValue=> EPS
        Text.FromNumber Number: Assets DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> Assets2
        Text.FromNumber Number: Liability DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> Liability2
        Text.FromNumber Number: ParentEquity DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> ParentEquity2
        Text.FromNumber Number: Sales DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> Sale2
        Text.FromNumber Number: Grossmargin DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> Grossmargin2
        Text.FromNumber Number: Operatingincome DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> Operatingincome2
        Text.FromNumber Number: EBIT DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> EBIT2
        Text.FromNumber Number: Profit DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> Profit2
        Text.FromNumber Number: ParentProfit DecimalPlaces: 0 UseThousandsSeparator: True FormattedNumber=> ParentProfit2
        Excel.CloseExcel.Close Instance: ExcelInstance
        Word.LaunchWord.Launch Visible: True Instance=> WordInstance
        Word.WriteToWord.WriteStartOfDocument Instance: WordInstance Text: $'''公司代號:
序號:1
主旨:113年第%CustomFormData['Period input']%季合併財務報告業經提報董事會通過
符合條款-第四條第XX款:31
事實發生日:%CustomFormData['BOD input']%
內容:
1.提報董事會或經董事會決議日期: %CustomFormData['BOD input']%
2.審計委員會通過日期: %CustomFormData['BOD input']%
3.財務報告或年度自結財務資訊報導期間
起訖日期(XXX/XX/XX~XXX/XX/XX):%Q3%
4.1月1日累計至本期止營業收入(仟元):%Sale2%
5.1月1日累計至本期止營業毛利(毛損) (仟元):%Grossmargin2%
6.1月1日累計至本期止營業利益(損失) (仟元):%Operatingincome2%
7.1月1日累計至本期止稅前淨利(淨損) (仟元):%EBIT2%
8.1月1日累計至本期止本期淨利(淨損) (仟元):%Profit2%
9.1月1日累計至本期止歸屬於母公司業主淨利(損) (仟元):%ParentProfit2%
10.1月1日累計至本期止基本每股盈餘(損失) (元):%EPS%
11.期末總資產(仟元):%Assets2%
12.期末總負債(仟元):%Liability2%
13.期末歸屬於母公司業主之權益(仟元):%ParentEquity2%
14.其他應敘明事項:無''' AppendNewLine: False
        Word.ReadFromWord.Read Instance: WordInstance WordData=> WordData
        Word.CloseWord.Close Instance: WordInstance
        System.RunApplication.RunApplication ApplicationPath: $'''C:\\Windows\\System32\\notepad.exe''' WindowStyle: System.ProcessWindowStyle.Normal ProcessId=> AppProcessId
        UIAutomation.PopulateTextField.PopulateTextField TextField: appmask['Window \'未命名 - 記事本\'']['Document \'文字編輯器\''] Text: WordData Mode: UIAutomation.PopulateTextMode.Append ClickType: UIAutomation.PopulateMouseClickType.SingleClick
        DISABLE Display.ShowMessageDialog.ShowMessage Title: $'''財報公告檔''' Message: $'''公司代號:
序號:1
主旨:11x年第x季合併財務報告業經提報董事會通過
符合條款-第四條第XX款:31
事實發生日:%CustomFormData['BOD input']%
內容:
1.提報董事會或經董事會決議日期: %CustomFormData['BOD input']%
2.審計委員會通過日期: %CustomFormData['BOD input']%
3.財務報告或年度自結財務資訊報導期間
起訖日期(XXX/XX/XX~XXX/XX/XX):%CustomFormData['Date input']%
4.1月1日累計至本期止營業收入(仟元):%Sales%
5.1月1日累計至本期止營業毛利(毛損) (仟元):%Grossmargin%
6.1月1日累計至本期止營業利益(損失) (仟元):%Operatingincome%
7.1月1日累計至本期止稅前淨利(淨損) (仟元):%EBIT%
8.1月1日累計至本期止本期淨利(淨損) (仟元):%Profit%
9.1月1日累計至本期止歸屬於母公司業主淨利(損) (仟元):%ParentProfit%
10.1月1日累計至本期止基本每股盈餘(損失) (元):%EPS%
11.期末總資產(仟元):%Assets%
12.期末總負債(仟元):%Liability%
13.期末歸屬於母公司業主之權益(仟元):%ParentEquity%
14.其他應敘明事項:無''' Icon: Display.Icon.Information Buttons: Display.Buttons.OK DefaultButton: Display.DefaultButton.Button1 IsTopMost: False ButtonPressed=> ButtonPressed2
    END
END

# [ControlRepository][PowerAutomateDesktop]

{
  "ControlRepositorySymbols": [
    {
      "IgnoreImagesOnSerialization": false,
      "Repository": "{\r\n  \"Screens\": [\r\n    {\r\n      \"Controls\": [\r\n        {\r\n          \"AutomationProtocol\": \"uia3\",\r\n          \"ScreenShot\": null,\r\n          \"ElementTypeName\": \"Document\",\r\n          \"InstanceId\": \"44ef9488-65c5-4673-be46-aa9c5a02595d\",\r\n          \"Name\": \"Document '文字編輯器'\",\r\n          \"SelectorCount\": 1,\r\n          \"Selectors\": [\r\n            {\r\n              \"CustomSelector\": null,\r\n              \"Elements\": [\r\n                {\r\n                  \"Attributes\": [\r\n                    {\r\n                      \"Ignore\": false,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Class\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": \"Edit\"\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Enabled\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": true\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Id\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": \"15\"\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Name\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": \"文字編輯器\"\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": true,\r\n                      \"Name\": \"Ordinal\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": 0\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Password\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": false\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Visible\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": true\r\n                    }\r\n                  ],\r\n                  \"CustomValue\": null,\r\n                  \"Ignore\": false,\r\n                  \"Name\": \"Document '文字編輯器'\",\r\n                  \"Tag\": \"document\"\r\n                }\r\n              ],\r\n              \"Ignore\": false,\r\n              \"ImageSelector\": null,\r\n              \"IsCustom\": false,\r\n              \"IsImageBased\": false,\r\n              \"IsWindowsInstance\": false,\r\n              \"Name\": \"Default Selector\",\r\n              \"Properties\": []\r\n            }\r\n          ],\r\n          \"Tag\": \"document\",\r\n          \"ScreenshotPath\": \"controlRepo-screenshots\\\\da229120-5e67-424f-a839-58d4b9834b48.png\"\r\n        }\r\n      ],\r\n      \"Handle\": {\r\n        \"value\": 0\r\n      },\r\n      \"ProcessName\": null,\r\n      \"ScreenShot\": null,\r\n      \"ElementTypeName\": \"Window\",\r\n      \"InstanceId\": \"2a2c0634-d743-4111-b145-677d39e8c67c\",\r\n      \"Name\": \"Window '未命名 - 記事本'\",\r\n      \"SelectorCount\": 1,\r\n      \"Selectors\": [\r\n        {\r\n          \"CustomSelector\": null,\r\n          \"Elements\": [\r\n            {\r\n              \"Attributes\": [\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Class\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": \"Notepad\"\r\n                },\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Enabled\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": true\r\n                },\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Id\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": \"\"\r\n                },\r\n                {\r\n                  \"Ignore\": false,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Name\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": \"未命名 - 記事本\"\r\n                },\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": true,\r\n                  \"Name\": \"Ordinal\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": 0\r\n                },\r\n                {\r\n                  \"Ignore\": false,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Process\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": \"notepad\"\r\n                },\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Visible\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": true\r\n                }\r\n              ],\r\n              \"CustomValue\": null,\r\n              \"Ignore\": false,\r\n              \"Name\": \"Window '未命名 - 記事本'\",\r\n              \"Tag\": \"window\"\r\n            }\r\n          ],\r\n          \"Ignore\": false,\r\n          \"ImageSelector\": null,\r\n          \"IsCustom\": false,\r\n          \"IsImageBased\": false,\r\n          \"IsWindowsInstance\": false,\r\n          \"Name\": \"Default Selector\",\r\n          \"Properties\": []\r\n        }\r\n      ],\r\n      \"Tag\": \"window\",\r\n      \"ScreenshotPath\": \"controlRepo-screenshots\\\\82013743-3a5e-4168-94fe-1a3ca398ed1c.png\"\r\n    }\r\n  ],\r\n  \"Version\": 1\r\n}",
      "ImportMetadata": {
        "DisplayName": "Computer",
        "ConnectionString": "",
        "Type": "Local",
        "DesktopType": "local"
      },
      "Name": "appmask"
    }
  ],
  "ImageRepositorySymbol": {
    "Repository": "{\r\n  \"Folders\": [],\r\n  \"Images\": [],\r\n  \"Version\": 1\r\n}",
    "ImportMetadata": {},
    "Name": "imgrepo"
  },
  "ConnectionReferences": []
}
```
</details>

<details>
<summary>Subflow Q4</summary>

```
Variables.CreateNewList List=> List
Variables.AddItemToList Item: $'''資產總計''' List: List
Variables.AddItemToList Item: $'''負債總計''' List: List
Variables.AddItemToList Item: $'''歸屬母公司業主之權益小計''' List: List
Variables.AddItemToList Item: $'''營業收入''' List: List
Variables.AddItemToList Item: $'''營業毛利''' List: List
Variables.AddItemToList Item: $'''營業淨利''' List: List
Variables.AddItemToList Item: $'''稅前淨利(損)''' List: List
Variables.AddItemToList Item: $'''本期淨利(損)''' List: List
Variables.AddItemToList Item: $'''母公司業主''' List: List
Variables.AddItemToList Item: $'''基本每股盈餘(虧損)''' List: List
IF CustomFormData['Period input'] = 4 THEN
    IF ButtonPressed <> $'''Cancel''' THEN
        Display.SelectFolder Description: $'''Select a folder to save the new PDF file...''' IsTopMost: True SelectedFolder=> DestinationFolder ButtonPressed=> ButtonPressed3
        Pdf.ExtractPages PDFFile: CustomFormData['PDF'] PageSelection: CustomFormData['PageNo'] ExtractedPDFPath: $'''%DestinationFolder%/NewPDFfile.pdf''' IfFileExists: Pdf.IfFileExists.AddSequentialSuffix ExtractedPDFFile=> ExtractedPDF
        Pdf.ExtractTablesFromPDF.ExtractTables PDFFile: ExtractedPDF MultiPageTables: True SetFirstRowAsHeader: True ExtractedPDFTables=> ExtractedPDFTables
        File.Delete Files: ExtractedPDF
        Excel.LaunchExcel.LaunchUnderExistingProcess Visible: True Instance=> ExcelInstance
        SET i TO 0
        Excel.RenameWorksheet.RenameWorksheetWithName Instance: ExcelInstance Name: $'''工作表1''' NewName: $'''Table %i + 1%'''
        LOOP FOREACH CurrentItem IN ExtractedPDFTables
            Excel.WriteToExcel.Write Instance: ExcelInstance Value: ExtractedPDFTables[i].DataTable
            Variables.IncreaseVariable Value: i IncrementValue: 1
            Excel.AddWorksheet Instance: ExcelInstance Name: $'''Table %i + 1%''' WorksheetPosition: Excel.WorksheetPosition.Last
        END
        Excel.SetActiveWorksheet.ActivateWorksheetByName Instance: ExcelInstance Name: $'''Table 3'''
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: List Column: $'''A''' Row: 1
        Excel.SetActiveWorksheet.ActivateWorksheetByName Instance: ExcelInstance Name: $'''Table 3'''
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(A1,\'Table 1\'!B1:D45,3,0)''' Column: $'''B''' Row: 1
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(A2,\'Table 1\'!K1:M45,3,0)''' Column: $'''B''' Row: 2
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(A3,\'Table 1\'!K1:M45,3,0)''' Column: $'''B''' Row: 3
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(\"*營業收入*\",\'Table 2\'!b1:D45,3,0)''' Column: $'''B''' Row: 4
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(A5,\'Table 2\'!B1:D45,3,0)''' Column: $'''B''' Row: 5
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=IFERROR(VLOOKUP(\"*營業淨利*\",\'Table 2\'!B1:D45,3,0),VLOOKUP(\"*營業淨利(損)*\",\'Table 2\'!B1:D45,3,0))''' Column: $'''B''' Row: 6
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=IFERROR(VLOOKUP(\"*稅前淨利*\",\'Table 2\'!B1:D45,3,0),VLOOKUP(\"*稅前淨利(損)*\",\'Table 2\'!B1:D45,3,0))''' Column: $'''B''' Row: 7
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=IFERROR(VLOOKUP(\"*本期淨利*\",\'Table 2\'!B1:D45,3,0),VLOOKUP(\"*本期淨利(損)*\",\'Table 2\'!B1:D45,3,0))''' Column: $'''B''' Row: 8
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=VLOOKUP(A9,\'Table 2\'!B1:D45,3,0)''' Column: $'''B''' Row: 9
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''=IFERROR(VLOOKUP(\"*基本每股盈餘*\",\'Table 2\'!B1:E45,4,0),VLOOKUP(\"*基本每股盈餘(虧損)*\",\'Table 2\'!B1:E45,4,0))''' Column: $'''B''' Row: 10
        DISABLE Excel.GetFirstFreeColumnRow Instance: `%ExcelInstance[Table 1]%` FirstFreeColumn=> FirstFreeColumn FirstFreeRow=> FirstFreeRow
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 1 ReadAsText: False CellValue=> Assets
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 2 ReadAsText: False CellValue=> Liability
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 3 ReadAsText: False CellValue=> ParentEquity
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 4 ReadAsText: False CellValue=> Sales
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 5 ReadAsText: False CellValue=> Grossmargin
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 6 ReadAsText: False CellValue=> Operatingincome
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 7 ReadAsText: False CellValue=> EBIT
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 8 ReadAsText: False CellValue=> Profit
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 9 ReadAsText: False CellValue=> ParentProfit
        Excel.ReadFromExcel.ReadCell Instance: ExcelInstance StartColumn: $'''B''' StartRow: 10 ReadAsText: False CellValue=> EPS
        Excel.CloseExcel.Close Instance: ExcelInstance
        Word.LaunchWord.Launch Visible: True Instance=> WordInstance
        Word.WriteToWord.WriteStartOfDocument Instance: WordInstance Text: $'''公司代號:
序號:1
主旨:113年第%CustomFormData['Period input']%季合併財務報告業經提報董事會通過
符合條款-第四條第XX款:31
事實發生日:%CustomFormData['BOD input']%
內容:
1.提報董事會或經董事會決議日期: %CustomFormData['BOD input']%
2.審計委員會通過日期: %CustomFormData['BOD input']%
3.財務報告或年度自結財務資訊報導期間
起訖日期(XXX/XX/XX~XXX/XX/XX):%Q1%
4.1月1日累計至本期止營業收入(仟元):%Sales%
5.1月1日累計至本期止營業毛利(毛損) (仟元):%Grossmargin%
6.1月1日累計至本期止營業利益(損失) (仟元):%Operatingincome%
7.1月1日累計至本期止稅前淨利(淨損) (仟元):%EBIT%
8.1月1日累計至本期止本期淨利(淨損) (仟元):%Profit%
9.1月1日累計至本期止歸屬於母公司業主淨利(損) (仟元):%ParentProfit%
10.1月1日累計至本期止基本每股盈餘(損失) (元):%EPS%
11.期末總資產(仟元):%Assets%
12.期末總負債(仟元):%Liability%
13.期末歸屬於母公司業主之權益(仟元):%ParentEquity%
14.其他應敘明事項:無''' AppendNewLine: False
        Word.ReadFromWord.Read Instance: WordInstance WordData=> WordData
        Word.CloseWord.Close Instance: WordInstance
        System.RunApplication.RunApplication ApplicationPath: $'''C:\\Windows\\System32\\notepad.exe''' WindowStyle: System.ProcessWindowStyle.Normal ProcessId=> AppProcessId
        UIAutomation.PopulateTextField.PopulateTextField TextField: appmask['Window \'未命名 - 記事本\'']['Document \'文字編輯器\''] Text: WordData Mode: UIAutomation.PopulateTextMode.Append ClickType: UIAutomation.PopulateMouseClickType.SingleClick
    END
END

# [ControlRepository][PowerAutomateDesktop]

{
  "ControlRepositorySymbols": [
    {
      "IgnoreImagesOnSerialization": false,
      "Repository": "{\r\n  \"Screens\": [\r\n    {\r\n      \"Controls\": [\r\n        {\r\n          \"AutomationProtocol\": \"uia3\",\r\n          \"ScreenShot\": null,\r\n          \"ElementTypeName\": \"Document\",\r\n          \"InstanceId\": \"44ef9488-65c5-4673-be46-aa9c5a02595d\",\r\n          \"Name\": \"Document '文字編輯器'\",\r\n          \"SelectorCount\": 1,\r\n          \"Selectors\": [\r\n            {\r\n              \"CustomSelector\": null,\r\n              \"Elements\": [\r\n                {\r\n                  \"Attributes\": [\r\n                    {\r\n                      \"Ignore\": false,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Class\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": \"Edit\"\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Enabled\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": true\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Id\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": \"15\"\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Name\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": \"文字編輯器\"\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": true,\r\n                      \"Name\": \"Ordinal\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": 0\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Password\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": false\r\n                    },\r\n                    {\r\n                      \"Ignore\": true,\r\n                      \"IsOrdinal\": false,\r\n                      \"Name\": \"Visible\",\r\n                      \"Operation\": \"EqualTo\",\r\n                      \"Value\": true\r\n                    }\r\n                  ],\r\n                  \"CustomValue\": null,\r\n                  \"Ignore\": false,\r\n                  \"Name\": \"Document '文字編輯器'\",\r\n                  \"Tag\": \"document\"\r\n                }\r\n              ],\r\n              \"Ignore\": false,\r\n              \"ImageSelector\": null,\r\n              \"IsCustom\": false,\r\n              \"IsImageBased\": false,\r\n              \"IsWindowsInstance\": false,\r\n              \"Name\": \"Default Selector\",\r\n              \"Properties\": []\r\n            }\r\n          ],\r\n          \"Tag\": \"document\",\r\n          \"ScreenshotPath\": \"controlRepo-screenshots\\\\da229120-5e67-424f-a839-58d4b9834b48.png\"\r\n        }\r\n      ],\r\n      \"Handle\": {\r\n        \"value\": 0\r\n      },\r\n      \"ProcessName\": null,\r\n      \"ScreenShot\": null,\r\n      \"ElementTypeName\": \"Window\",\r\n      \"InstanceId\": \"2a2c0634-d743-4111-b145-677d39e8c67c\",\r\n      \"Name\": \"Window '未命名 - 記事本'\",\r\n      \"SelectorCount\": 1,\r\n      \"Selectors\": [\r\n        {\r\n          \"CustomSelector\": null,\r\n          \"Elements\": [\r\n            {\r\n              \"Attributes\": [\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Class\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": \"Notepad\"\r\n                },\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Enabled\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": true\r\n                },\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Id\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": \"\"\r\n                },\r\n                {\r\n                  \"Ignore\": false,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Name\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": \"未命名 - 記事本\"\r\n                },\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": true,\r\n                  \"Name\": \"Ordinal\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": 0\r\n                },\r\n                {\r\n                  \"Ignore\": false,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Process\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": \"notepad\"\r\n                },\r\n                {\r\n                  \"Ignore\": true,\r\n                  \"IsOrdinal\": false,\r\n                  \"Name\": \"Visible\",\r\n                  \"Operation\": \"EqualTo\",\r\n                  \"Value\": true\r\n                }\r\n              ],\r\n              \"CustomValue\": null,\r\n              \"Ignore\": false,\r\n              \"Name\": \"Window '未命名 - 記事本'\",\r\n              \"Tag\": \"window\"\r\n            }\r\n          ],\r\n          \"Ignore\": false,\r\n          \"ImageSelector\": null,\r\n          \"IsCustom\": false,\r\n          \"IsImageBased\": false,\r\n          \"IsWindowsInstance\": false,\r\n          \"Name\": \"Default Selector\",\r\n          \"Properties\": []\r\n        }\r\n      ],\r\n      \"Tag\": \"window\",\r\n      \"ScreenshotPath\": \"controlRepo-screenshots\\\\82013743-3a5e-4168-94fe-1a3ca398ed1c.png\"\r\n    }\r\n  ],\r\n  \"Version\": 1\r\n}",
      "ImportMetadata": {
        "DisplayName": "Computer",
        "ConnectionString": "",
        "Type": "Local",
        "DesktopType": "local"
      },
      "Name": "appmask"
    }
  ],
  "ImageRepositorySymbol": {
    "Repository": "{\r\n  \"Folders\": [],\r\n  \"Images\": [],\r\n  \"Version\": 1\r\n}",
    "ImportMetadata": {},
    "Name": "imgrepo"
  },
  "ConnectionReferences": []
}

```
</details>
