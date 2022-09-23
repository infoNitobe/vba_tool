Attribute VB_Name = "Module1"
Option Explicit

Sub CopyFileRename()
    Dim src_file_path As String
    Dim ws As Worksheet
    Dim num_file As Integer
    Dim src_file_name As String

    Set ws = Worksheets(1)
    src_file_name = ws.Cells(3, "C").Value
    src_file_path = ThisWorkbook.Path & "\" & src_file_name
    num_file = ws.Cells(3, "B").End(xlDown).Row - 2
    
    Dim dest_file_path As String
    Dim i As Integer
    For i = 0 To num_file - 1
        dest_file_path = ThisWorkbook.Path & "\" & ws.Cells(3 + i, "B").Value
        FileCopy src_file_path, dest_file_path
    Next
End Sub
