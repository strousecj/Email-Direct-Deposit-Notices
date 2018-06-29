'Function to format and merge Payment Info and Invoice information workbooks exported from Foundation accounting software
'into a single sheet.
Sub FormatDirectDepositSheet()
    Application.ScreenUpdating = False 'Prevents screen flashing as macros run
    Application.Run "PERSONAL.XLSB!BrokerPaymentDeleteColumns"
    Application.Run "PERSONAL.XLSB!RenameColumns"
    Application.Run "PERSONAL.XLSB!AutoFillTruckerEmail"
    Application.Run "PERSONAL.XLSB!FillInvoiceNum"
    Range("F5").Select
    Application.ScreenUpdating = True 'Updates screen to reflect changes made by macro
End Sub

Private Sub BrokerPaymentDeleteColumns()
    Range("A:A,B:B,I:I,J:J,K:K,L:L,M:M,N:N,O:O,P:P").Select
    Range("P1").Activate
    Selection.Delete shift:=xlToLeft
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Email"
    Rows("1:1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Bold = True
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "m/d/yyyy"
    Range("D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Range("E9").Select
End Sub

Private Sub RenameColumns()
    Dim cell As Range

    ActiveSheet.UsedRange
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Vendor #"
    Range("A:A").Select
    Selection.HorizontalAlignment = xlCenter
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Payment Date"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Payment Name"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Deposit Amount"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Bank Routing #"
    Range("E:E").Select
    Selection.HorizontalAlignment = xlCenter
    
    Range("F:F").Select
    Selection.HorizontalAlignment = xlCenter
    Range("F1").EntireColumn.Insert
    Range("F2", Selection.End(xlDown)).Select
    Selection.Formula = "=If(LEN(G2)>4,Right(G2, 4),"""")"
    Range("F2", Selection.End(xlDown)).Select
    Selection.Value = Selection.Value
    Range("F2", Selection.End(xlDown)).Select
    Selection.NumberFormat = "@"
    For Each cell In Range("F2", Range("F" & Rows.Count).End(xlUp))
        If Len(cell) < 4 Then cell.Value = "0" & cell.Value
    Next cell
    
    Range("G1").Activate
    Selection.Delete shift:=xlToLeft
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Bank Account #"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Email Address"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Invoice #"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Company Name"
    
    Range("G14").Select
End Sub

Private Sub AutoFillTruckerEmail()
    Dim w1 As Worksheet, w2 As Worksheet, wb As Workbook
    Dim c As Range, nameCell As Long, emailCell As String
    Dim i As Integer
    
    i = 2 'Row to place email into in w1
    Set w1 = Workbooks("DDPaymentInfo.xlsx").Worksheets("DDPaymentInfo")
    Set wb = Workbooks.Open("C:\Users\Chandler.strouse\Desktop\VendorList.xlsx")
    Set w2 = wb.Sheets("VendorList") 'Selects the broker list sheet in the workbook
    
    For Each c In w1.Range("C2", w1.Range("C" & Rows.Count).End(xlUp)) 'c is equal broker's name
        nameCell = 0 'Initializes value to row 0
        On Error Resume Next
        nameCell = Application.Match(c, w2.Columns("B"), 0) 'matches columns containing vendor names
        emailCell = w2.Range("C" & nameCell) 'range is column containing email addresses
        On Error GoTo 0
        If nameCell <> 0 Then w1.Range("G" & i).Value = emailCell 'range is column to place email addresses into
        i = i + 1 'Increments i to place email in next row
    Next c
    
    wb.Close (False) 'Closes workbook, doesn't save changes
End Sub

Private Sub FillInvoiceNum()
    Dim w1 As Worksheet, w2 As Worksheet, wb As Workbook
    Dim c As Range, nameCell As Long, invoiceNum As String
    Dim tempC As String, invoiceAmount As Double
    Dim i As Integer, x As Integer
    
    i = 2 'Row to place invoice into in w1
    Set w1 = Workbooks("DDPaymentInfo.xlsx").Worksheets("DDPaymentInfo")
    Set wb = Workbooks.Open("C:\Users\Chandler.strouse\Desktop\DDInvoices.xlsx")
    Set w2 = wb.Sheets("DDInvoices") 'Selects the broker list sheet in the workbook
    
    nameCell = 0 'Initializes value to row 0
    tempC = ""
    For Each c In w1.Range("C2", w1.Range("C" & Rows.Count).End(xlUp)) 'c is equal to broker's name
        On Error Resume Next
        If x > 0 Then
            x = x - 1
        Else
            Do
                If tempC = c.Value2 Then
                    x = x + 1
                    w1.Rows(i - 1).EntireRow.Copy
                    w1.Rows(i).Insert shift:=xlUp
                Else
                    x = 0
                End If
                tempC = c.Value2
                nameCell = (Application.Match(c, w2.Columns("E"), 0)) + x 'matches columns containing vendor names
                invoiceNum = w2.Range("J" & (nameCell)) 'range is column containing invoice #'s
                invoiceAmount = w2.Range("Q" & nameCell) 'range is column containg invoice amounts
                On Error GoTo 0
                If nameCell <> 0 Then
                    w1.Range("H" & i).Value = invoiceNum 'range is column to place invoice # into
                    w1.Range("D" & i).Value = invoiceAmount 'range is column to place invoice amount into
                End If
                i = i + 1 'Increments i to move to next row
            Loop While tempC = w2.Range("E" & (nameCell + 1))
        End If
    Next c
    
    wb.Close (False) 'Closes workbook, doesn't save changes
    Application.CutCopyMode = False 'removes marching ants from copying rows
    Range("A:I").Select 'Selects Email column
    Selection.Columns.AutoFit 'Autofits all columns
End Sub
