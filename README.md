# VBA-Tool
'This code is to create a VBA tool to combine all the excel files with the same format into a unique file
'このコードは、同じ形式のすべてのExcelファイルを1つのユニークなファイルに統合するツールを作成するためのものです。

Sub CopyDuLieu()


Dim wbOutput, wbInput As Workbook
Dim selectFiles As Variant
Dim iFileNum, isSheetnum As Integer
Dim iLastRowInput, iLastRowOutput As Long 'As the Number of Excel rows is large
Dim Columnname As Integer
Columnname = 0


Application.DisplayAlerts = False
ApplicationScreenUpdating = False

'Step1: create a file for combination
'ステップ1：統合用のファイルを作成する
Workbooks.Add 'Create a file ファイルを作成する
ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path 'Rename the file ファイルの名前を変更する
Set wbOutput = ActiveWorkbook


'Step 2: Open many files
'ステップ2: 複数のファイルを開く
selectFiles = Application.GetOpenFilename(Filefilter:="Excel File (*.xls*),*.xlsx*", MultiSelect:=True)

'Step3: Open each file to check the number of excel sheets
'ステップ3：各ファイルを開き、Excelシートの数を確認する

For iFileNum = 1 To UBound(selectFiles)
 Set wbInput = Workbooks.Opem(slectFiles(iFileNum))
    isSheetnum = wbInput.Worksheets.Count
    
'Step 4: Check if the sheet has no value. If it does, then ignore the sheet
'ステップ4: シートに値がないかを確認します｡値がない場合は､そのシートを無視します｡

 For i = 1 To iSheetNum
    If wbInput.Sheets(i).Cells(Left(ThisWorkbook.Sheets(1).Cells(7, 4), 2)) <> " " Then
    
    
'Step 5: Find out the last row for each file
'ステップ5：各ファイルの最終行を特定する

iLastRowInput = wbInput.Sheets(i).Range(ThisWorkbook.Sheets(1).Cells(9, 4) & Rows.Count).End(xlUp).Row
iLastRowOutput = wbOutput.Sheets(i).Range(ThisWorkbook.Sheets(1).Cells(9, 4) & Rows.Count).End(xlUp).Row

'Step 6: Copy the column name
'ステップ6: 列名をコピーする

    If tieude = 0 Then
        'Step 6+7 copy both the column name and the data for the first file
        'ステップ6+7：最初のファイルについては、列名とデータの両方をコピーする
        
        wbInput.Sheets(i).Range(ThisWorkbook.Sheets(1).Cells(11, 4) & ":" & _
        ThisWorkbook.Sheets(1).Cells(9, 5) & iLastRowInput).Copy _
        Destination:=wbOutput.Sheets(1).Range(ThisWorkbook.Sheets(1).Cells(9, 4) & iLastRowOutput + 1)
    Else
    'Buoc 7: Copy only the data from the second time
    'ステップ7：2回目以降はデータのみをコピーする
        wbInput.Sheets(i).Range(ThisWorkbook.Sheets(1).Cells(10, 5) + 1 & ":" & _
        ThisWorkbook.Sheets(1).Cells(9, 5) & iLastRowInput).Copy _
        Destination:=wbOutput.Sheets(1).Range(ThisWorkbook.Sheets(1).Cells(9, 4) & iLastRowOutput + 1)

    End If
    Columnname = 1

End If
Next

wbInput.Close

Next
MsgBox "UpdateDone"

End Sub

