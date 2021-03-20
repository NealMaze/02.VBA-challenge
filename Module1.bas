Attribute VB_Name = "Module1"
' Helper Classes
' Run Module2 instead

'Global Variables
Public PrintRow As Integer
Public PrintColumn As Single
Public StrPrintCell As String
Public PrintValue As Double
Public WorkRow As Double
Public SecWorkRow As Double
Public WorkTicker As String
Public OpenValue As Single
Public CloseValue As Single
Public Yearly As Double

Sub SetUp():
    Range("I1:S2837").Value = Null
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Range("I1:S1").Font.Bold = True
    Range("I2:S2837").Interior.ColorIndex = 0
    Columns("H").ColumnWidth = 4
    Columns("I").ColumnWidth = 6.5
    Columns("J").ColumnWidth = 12.5
    Columns("K").ColumnWidth = 14.5
    Columns("L").ColumnWidth = 18
    Columns("M").ColumnWidth = 18
    Columns("N").ColumnWidth = 18
    Columns("O").ColumnWidth = 18
    Columns("P").ColumnWidth = 18
    Columns("Q").ColumnWidth = 18
    
    
    PrintRow = 2
    PrintColumn = 9
    WorkRow = 1
    SecWorkRow = 1
    WorkTicker = "Not Initialized"
End Sub

Sub IncrementPrintCell():
    If PrintRow < 2 Then
        PrintRow = 2
        PrintColumn = 9
    ElseIf PrintColumn < 12 Then
        PrintColumn = PrintColumn + 1
    Else
        PrintRow = PrintRow + 1
        PrintColumn = 9
    End If
End Sub

Sub IncrementWorkRange():
    WorkRow = SecWorkRow + 1
    WorkTicker = Cells(WorkRow, 1).Value
    SecWorkRow = WorkRow
    If WorkTicker <> "" Then
        While WorkTicker = Cells(SecWorkRow + 1, 1).Value
            SecWorkRow = SecWorkRow + 1
            WorkTicker = Cells(WorkRow, 1).Value
        Wend
    End If
End Sub

Sub PrintTicker():
    Cells(PrintRow, PrintColumn) = WorkTicker
    Call IncrementPrintCell
End Sub

Sub PrintYearly():
    CloseValue = Cells(SecWorkRow, 6).Value
    OpenValue = Cells(WorkRow, 3).Value
    Yearly = CloseValue - OpenValue
    Yearly = Round(Yearly, 2)
    Cells(PrintRow, PrintColumn).Value = Yearly
    Cells(PrintRow, PrintColumn).NumberFormat = "General"
    If Yearly > 0 Then
        Cells(PrintRow, PrintColumn).Interior.ColorIndex = 4
    ElseIf Yearly < 0 Then
        Cells(PrintRow, PrintColumn).Interior.ColorIndex = 3
    End If
    Call IncrementPrintCell
End Sub


Sub PrintPercent():
    Dim Percent As Double
    If OpenValue = 0 Then
        OpenValue = 1
        CloseValue = 1
    End If
    Percent = (CloseValue / OpenValue) - 1
    Cells(PrintRow, PrintColumn).Value = Percent
    Cells(PrintRow, PrintColumn).NumberFormat = "0.00%"
    Call IncrementPrintCell
End Sub

Sub PrintTotal():
    Dim Volume As LongLong
    Dim Total As LongLong
    Dim x As LongLong
    
    For x = WorkRow To SecWorkRow
        Volume = Cells(x, 7).Value
        Total = Total + Volume
    Next
    Cells(PrintRow, PrintColumn).Value = Total
    Cells(PrintRow, PrintColumn).NumberFormat = "General"
    Call IncrementPrintCell
End Sub



Sub OneSheet():
    Call SetUp
    While WorkTicker <> ""
        Call IncrementWorkRange
        Call PrintTicker
        Call PrintYearly
        Call PrintPercent
        Call PrintTotal
    Wend
End Sub

Sub RunSheet(ws As Worksheet):
    Call SetUp
    While WorkTicker <> ""
        Call IncrementWorkRange
        Call PrintTicker
        Call PrintYearly
        Call PrintPercent
        Call PrintTotal
    Wend
End Sub

'1:37







































