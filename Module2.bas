Attribute VB_Name = "Module2"
'Run

Sub ButtonOne():
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Select
        Call RunSheet(ws)
    Next
End Sub
'1:38
