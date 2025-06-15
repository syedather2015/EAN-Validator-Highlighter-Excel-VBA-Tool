# 🧾 EAN Validator & Highlighter – Excel VBA Tool

This Excel macro streamlines the process of validating **EAN (European Article Numbers)** by checking formatting rules, ensuring correct length, and automatically applying **13-digit padding**. Invalid entries are visually highlighted, making it easier to spot and correct them during product data onboarding or audit processes.

> Developed by **Syed Ather Rizvi** for product data integrity and quality assurance.

---

## ⚙️ Features

- ✅ Prompts user to select the column with EANs
- 🔢 Automatically pads EANs with leading zeros to make them 13 digits
- 🧠 Applies rule-based checks (no leading "2", no "00000" patterns, etc.)
- 🎨 Highlights invalid EANs with **blue background** and **white text**
- 📊 Works on the **active worksheet** – no setup required

---

## 🧪 Validation Rules Applied

A valid EAN must:
- Be numeric
- Be 13 digits or less (padded if needed)
- **Not start with `2`**
- **Not contain `00000`** in positions:
  - 1–3
  - 3–7
  - 8–13

---

## 🔁 How to Use

1. Open your Excel file with EAN data
2. Press `Alt + F11` to open the **VBA Editor**
3. Insert a new Module and paste the macro code
4. Run `ValidateAndHighlight` from the Macro window
5. When prompted, enter the **column letter** where EANs exist (e.g., `B`, `D`, etc.)

---

## 💡 Use Cases

- Product onboarding & GTIN validation  
- Retail or eCommerce data audits  
- Supplier catalog quality checks  
- Marketing intelligence & pricing validation  

---

## 📌 Visual Output

- Valid EANs are auto-corrected to 13 digits
- Invalid EANs are marked **blue** with **white text** for easy identification

---

## 📄 Code Highlights

```vba
Sub ValidateAndHighlight()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim eanRange As Range
    Dim cell As Range
    Dim columnLetter As String
    
    ' Prompt the user to enter the column letter
    columnLetter = InputBox("Enter the column letter containing EAN codes (e.g., A, B, C)", "Column Selection")
    
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
            
            ' Check validity of EAN
            If IsValidEAN(ean) Then
                ' Add leading zeros to make it 13 digits if necessary
                Dim newEAN As String
                newEAN = WorksheetFunction.Rept("0", 13 - Len(ean)) & ean
                
                ' If new EAN is still valid after padding, update the value
                If IsValidEAN(newEAN) Then
                    cell.Value = newEAN
                Else
                    HighlightCell cell
                End If
            Else
                HighlightCell cell
            End If
        End If
    Next cell
End Sub

Function IsValidEAN(ByVal ean As String) As Boolean
    ' Check if the EAN is valid based on given conditions
    If IsNumeric(ean) And Len(ean) <= 13 And _
       Left(ean, 1) <> "2" And _
       Mid(ean, 1, 3) <> "000" And _
       Mid(ean, 3, 5) <> "00000" And _
       Mid(ean, 8, 5) <> "00000" Then
        IsValidEAN = True
    Else
        IsValidEAN = False
    End If
End Function

Sub HighlightCell(ByRef cell As Range)
    ' Highlight the invalid cell with Blue color and white font
    With cell.Interior
        .Pattern = xlSolid
        .Color = RGB(0, 176, 240) ' Blue background color (#00B0F0)
    End With
    With cell.Font
        .Color = RGB(255, 255, 255) ' White text color
    End With
End Sub

Function IsValidColumn(ByVal columnLetter As String) As Boolean
    ' Check if the column letter is between A and XFD (the last column)
    If Len(columnLetter) = 1 And columnLetter Like "[A-Za-z]" Then
        IsValidColumn = True
    Else
        IsValidColumn = False
    End If
End Function
