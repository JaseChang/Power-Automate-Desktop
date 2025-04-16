### 變數{Variable}
---
- 設定變數{x}

  說明：
  ```
  將簡單資料，如文字、數值、布林值，或進階資料，如清單、資料表、自訂物件，
  設定成變數 Variable
  ```
  
  應用：

  ```
  1. 數值變數運算(加減乘除)
      %NewVar+NewVar2%、%NewVar-NewVar2%、%NewVar*NewVar2%、%NewVar/NewVar2%
  2. 布林值變數（是或否），用於條件判斷
      %NewVar>NewVar2%、％ＮewVar=NewVar2%
  ```
- 建立清單+新增項目至清單
  
  說明：
  ```
  建立一個新清單，並建立項目至清單裡
  ```

  應用：
  ```
  1.表頭或資料表建立
  ```

![upload_60d003024aadb24e53c00001e15d48c7](https://github.com/user-attachments/assets/cc9effd0-1649-4b33-b874-7de26c2426b7)
![upload_47da88b8dc533bdb60401c279ec75185](https://github.com/user-attachments/assets/1af8db6d-86c7-459b-917b-0deede18d09b)

```
        Variables.CreateNewList List=> List
        Variables.AddItemToList Item: $'''A''' List: List
        Variables.AddItemToList Item: $'''B''' List: List
        Variables.AddItemToList Item: $'''C''' List: List    
        Variables.CreateNewList List=> List2
        Variables.AddItemToList Item: 1 List: List2
        Variables.AddItemToList Item: 2 List: List2
        Variables.AddItemToList Item: 3 List: List2
        SET Table TO {List, List2 }
```
    
### 迴圈相關應用
---  

### 文字相關應用
---
- 將文字設定成數字
  
  說明:
  ```
  將只包含數字的文字變數，轉換為數值變數
  ```
  應用:
  ```
  如顯示輸入對話方塊動作中，產生的變數 "UserInput" 預設為文字，如需對該變數做運算，得將轉換為數值變數
  ```
### Excel相關應用
---
- 從 Excel 工作表中取得第一個可用欄/可用列

  說明:
  ```
    將資料加入至工作表時，需先擷取工作表中第一個空白欄位
  ```
  應用:
  
    ![image](https://hackmd.io/_uploads/SyXeSOBoC.png)

- {x}設定變數+設定使用中Excel工作表
  
  說明：
  
  ```
  Excel中有多個工作表時，可選取指定的工作表
  ```
  
  應用：
  
  ```
  將工作表"Product"設定為使用中工作表
  ```
![image](https://github.com/user-attachments/assets/debb21da-2d56-423c-87d1-5e27d92bf14b)
![image](https://github.com/user-attachments/assets/99b9ff87-1e43-4b3e-8d7f-406ce54983b5)

- 選取Excel工作表的儲存格
  
  說明：
  ```
  指定RPA要進行處理之儲存格範圍
  ```
<img width="418" alt="image" src="https://github.com/user-attachments/assets/951bbd03-2e9c-4605-bdac-61d5de2f5943" />

-將選取之儲存格複製至工作表

  說明：

  ```
  將複製之儲存格複製至指定欄位或範圍
  ```

- <img width="409" alt="image" src="https://github.com/user-attachments/assets/91369b4d-01d8-4e30-bc99-67722845faad" />


- For each 迴圈+寫入Excel工作表

  說明：
  
  ```
  重複寫入Exel工作表
  ```
  
  應用：
  
  ```
  重複將B欄Price及C欄Cost文字轉換成數值，並做運算填入D欄Profit
  ```
 ![image](https://github.com/user-attachments/assets/c6cf9301-b002-4452-bf02-a08358ae0789)
![image](https://github.com/user-attachments/assets/262c443f-ebd3-44a8-b382-1ce3ddf48eb6)


- 執行Excel 巨集
  
  說明：
  ```
    啟動附檔名為.xlsm之Excel開啟巢狀處裡及載入增益集及巨集，設定執行Excel 巨集之名稱
  ```
    ![image](https://hackmd.io/_uploads/B1kjVyriC.png)
    ![image](https://hackmd.io/_uploads/BkL9HJHoR.png)
    ![image](https://hackmd.io/_uploads/H1N2r1Hs0.png)
    
  應用：
  
  ```
  VBA資料整理
  ```
  ```
    SET Path TO $'''C:\\Users\\jasej\\Desktop\\VBA Test'''
    SET ActiveSheet TO $'''Product'''
    DateTime.GetCurrentDateTime.Local DateTimeFormat: DateTime.DateTimeFormat.DateAndTime CurrentDateTime=> CurrentDateTime
    Text.ConvertDateTimeToText.FromCustomDateTime DateTime: CurrentDateTime CustomFormat: $'''yyymmddhhmm''' Result=> FormattedDateTime
    Excel.LaunchExcel.LaunchAndOpen Path: $'''%Path%\\VBA Test.xlsm''' Visible: True ReadOnly: False LoadAddInsAndMacros: True Instance=> ExcelInstance
    Excel.SetActiveWorksheet.ActivateWorksheetByName Instance: ExcelInstance Name: ActiveSheet
    Excel.ReadFromExcel.ReadAllCells Instance: ExcelInstance ReadAsText: False FirstLineIsHeader: True RangeValue=> ExcelDataProduct
    Excel.InsertColumn Instance: ExcelInstance Column: $'''D'''
    Excel.InsertColumn Instance: ExcelInstance Column: $'''F'''
    Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''Profit''' Column: $'''D''' Row: 1
    Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: $'''Total Profit''' Column: $'''F''' Row: 1
    SET RowNumber TO 2
    LOOP FOREACH CurrentProduct IN ExcelDataProduct
        Text.ToNumber Text: CurrentProduct['Price'] Number=> Price
        Text.ToNumber Text: CurrentProduct['Cost'] Number=> Cost
        Text.ToNumber Text: CurrentProduct['Unit sold'] Number=> Unitsold
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: Price - Cost Column: $'''D''' Row: RowNumber
        Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: (Price - Cost) * Unitsold Column: $'''F''' Row: RowNumber
        SET RowNumber TO RowNumber + 1
    END
    Excel.RunMacro Instance: ExcelInstance Macro: $'''FormatTable'''
    Excel.ResizeColumnsOrRows.AutofitAllColumns Instance: ExcelInstance
    Excel.CloseExcel.CloseAndSaveAs Instance: ExcelInstance DocumentFormat: Excel.ExcelFormat.FromExtension DocumentPath: $'''%Path%\\Output_%FormattedDateTime%.xlsx'''
  ```
   
   

Outlook相關應用
---
- 擷取來自Outlook訊息
  
  說明：
  ```
  擷取來自Outlook Mail
  ```
  應用
  ```
  擷取業務匯款通知信件，
  ```
  <img width="468" alt="image" src="https://github.com/user-attachments/assets/f80d6ce4-dd45-424b-9ee4-959bdf632adf" />
  <img width="403" alt="image" src="https://github.com/user-attachments/assets/dd9fe5f4-3c57-420c-bdd4-c3e0d83cb7d5" />

### PDF相關應用
---

### 日期時間相關應用
---
- 取得目前日期及時間+將日期轉換為文字

  應用：

  ```
  將檔名加上日期時間
  ```
![upload_906c2edab3ad50cf5a4696acb66746c0](https://github.com/user-attachments/assets/8378d4df-0a90-43c1-997f-403c8f9f5d96)

![upload_4620850e009d7cd7bb052eb7cea9a149](https://github.com/user-attachments/assets/41b14bda-42ea-4096-a1a6-8fcb59ae1c87)


### SAP登入
---
說明：

  ```
  SAP自動登入
  ```

  ![image](https://hackmd.io/_uploads/HyBTeyioR.png)
   
