VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ValidateAndHighlight()
 Dim ws As Worksheet
 Dim lastRow As Long
 Dim eanRange As Range
 Dim cell As Range
 Dim selectedColumn As Range
 Dim columnLetter As String
 ' Prompt the user to enter the column letter
 columnLetter = InputBox("Enter the column letter containing EAN codes (e.g., A, B, C)", "Column
Selection ")"
 ' Check if the entered column letter is valid
 If Not IsValidColumn(columnLetter) Then
 MsgBox "Invalid column letter. Exiting macro."
 Exit Sub
 End If
 ' Set the active sheet
 Set ws = ActiveSheet
 ' Find the last row in the selected column
 lastRow = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row
 ' Set the range of EAN codes
 Set eanRange = ws.Range(columnLetter & "2:" & columnLetter & lastRow)
 ' Loop through each cell in the range
 For Each cell In eanRange
 ' Skip blank cells
 If Not IsEmpty(cell) Then
 Dim ean As String
 ean = CStr(cell.Value)
 ' Check validity of EANs and add zeros
 If IsNumeric(ean) And Len(ean) <= 13 And Left(ean, 1) <> "2" And Mid(ean, 1, 3) <> "000" And Mid(ean, 3, 5) <> "00000" And Mid(ean, 8, 5) <> "00000" Then
 ' Add zeros to make it at least 13 digits long
 Dim newEAN As String
 newEAN = WorksheetFunction.Rept("0", 13 - Len(ean)) & ean
 ' Check validity of the new EAN
 If IsNumeric(newEAN) And Len(newEAN) = 13 And Left(newEAN, 1) <> "2" And Mid(newEAN, 1, 3) <> "000" And Mid(newEAN, 3, 5) <> "00000" And Mid(newEAN, 8, 5) <> "00000" Then
 ' New EAN is valid
 cell.Value = newEAN
 Else
 ' Invalid EAN, highlight the cell
 With cell.Interior
 .Pattern = xlSolid
 .Color = RGB(128, 0, 128) ' Purple background color
 End With
 With cell.Font
 .Color = RGB(255, 255, 255) ' White text color
 End With
 End If
 Else
 ' Invalid EAN, highlight the cell
 With cell.Interior
 .Pattern = xlSolid
 .Color = RGB(128, 0, 128) ' Purple background color
 End With
 With cell.Font
 .Color = RGB(255, 255, 255) ' White text color
 End With
 End If
 End If
 Next cell
End Sub
Function IsValidColumn(ByVal columnLetter As String) As Boolean
 ' Check if the column letter is between A and XFD (the last column)
 If Len(columnLetter) = 1 And columnLetter Like "[A-Za-z]" Then
 IsValidColumn = True
 Else
 IsValidColumn = False
 End If
End Function
