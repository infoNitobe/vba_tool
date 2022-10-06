Attribute VB_Name = "Module1"
Option Explicit

Sub copyMatchingRow()
    Dim ws_tar As Worksheet: Set ws_tar = Worksheets(1)
    Dim ws_src As Worksheet: Set ws_src = Worksheets(2)
    Dim ws_setting As Worksheet: Set ws_setting = Worksheets(3)
    Dim row_offset_tar As Integer: row_offset_tar = ws_setting.Cells(3, "B")
    Dim col_offset_tar As Integer: col_offset_tar = ws_setting.Cells(4, "B")
    Dim col_num_tar As Integer: col_num_tar = ws_setting.Cells(5, "B")
        Const CELL_WITHOUT_SEARCH_TARGET As Integer = 2
    'åüçıëŒè€ÇÃíPåÍêîÇéZèo
    Dim row_src As Integer
    row_src = ws_src.Cells(Rows.Count, "B").End(xlUp).Row - CELL_WITHOUT_SEARCH_TARGET
    Dim cell_found As Range
    Dim i As Integer
    Dim row_pasted_src As Integer
    Dim row_copied_tar As Integer
    Const row_offset_src As Integer = 3
    Const col_pasted_src As Integer = 4
    
    For i = 1 To row_src
        Set cell_found = ws_tar.Columns(col_offset_tar).Find(ws_src.Cells(2 + i, "B"), LookAt:=xlWhole)
        row_pasted_src = row_offset_src + i - 1
        If cell_found Is Nothing Then
            ws_src.Cells(row_pasted_src, "C").Value = "nothing"
        ElseIf Not (cell_found Is Nothing) Then
            row_copied_tar = cell_found.Row
            ws_src.Cells(row_pasted_src, "C").Value = cell_found.Row & "ÅA" & cell_found.Column
            ws_src.Range(ws_src.Cells(row_pasted_src, col_pasted_src), ws_src.Cells(row_pasted_src, col_pasted_src + col_num_tar)).Value = _
                ws_tar.Range(ws_tar.Cells(row_copied_tar, 2), ws_tar.Cells(row_copied_tar, 2 + col_num_tar)).Value
        End If
    Next i
    
End Sub


