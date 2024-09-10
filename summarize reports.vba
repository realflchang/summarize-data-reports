' AP Summary Macro to prepare Commission Statement Report for check processing
' By Frank Chang    realflchang@yahoo.com
' Edited 2/15/2022 - no longer has Sheet1 requirements, can run any number of times

Sub RenameSheets()

    Dim WS As Worksheet
    Dim SheetNum As String
    Dim SumStartRow As Integer
        SumStartRow = 6  ' Sum sheet, starting row of list of worksheets

    Dim iSheet, i As Integer
        iSheet = 0
    Dim istart, iend As Integer
    
    Dim shtAcctCell As String
        shtAcctCell = "C1"   ' cell with account code
        
    Dim shtBalCell As String
        shtBalCell = "J8"    ' cell with account balance

    ' In case a worksheet already has 1, 2, 3, etc., just change them to something else to avoid conflict
    For Each WS In Sheets

        If WS.Name <> "Sum" Then
            iSheet = iSheet + 1
            WS.Name = "ren#" & CStr(iSheet)
        End If
    Next
    iSheet = 0
    
    For Each WS In Sheets

        If WS.Name <> "Sum" Then
            iSheet = iSheet + 1
            WS.Name = CStr(iSheet)
            i = iSheet
            
            ' Populate Sum worksheet
            Sheets("Sum").Cells(i + SumStartRow - 1, 1).Value = (i)
            Sheets("Sum").Cells(i + SumStartRow - 1, 2).NumberFormat = "General"
            Sheets("Sum").Cells(i + SumStartRow - 1, 2).Formula = "='" & (i) & "'!" & shtAcctCell 'Account Code
            Sheets("Sum").Cells(i + SumStartRow - 1, 3).Formula = "='" & (i) & "'!" & shtBalCell 'Stmt Balance
            
        End If
    Next

    istart = CStr(iSheet) + SumStartRow
    iend = CStr(iSheet) + SumStartRow + 3
    For i = istart To iend ' Continue Iterate rows, make next 5 rows blank
        Sheets("Sum").Cells(i, 1).Value = ""
        Sheets("Sum").Cells(i, 2).NumberFormat = "General"
        Sheets("Sum").Cells(i, 2).Value = ""
        Sheets("Sum").Cells(i, 3).Value = ""
    Next


End Sub

