VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub DeleteMatchingLines()
    Dim ws_source As Worksheet
    Dim row_offset As Integer
    Dim col_offset As Integer
    Dim row_source As Integer
    Const CELL_WITHOUT_SEARCH_TARGET As Integer = 2
    Set ws_source = Worksheets(2)
    row_offset = ws_source.Cells(3, "B")
    col_offset = ws_source.Cells(4, "B")
    '検索対象の単語数を算出
    row_source = Worksheets(2).Cells(Rows.Count, "D").End(xlUp).Row - CELL_WITHOUT_SEARCH_TARGET

    Dim ws_target As Worksheet
    Set ws_target = Worksheets(1)
    Dim CheckCells As Range
    Dim i As Integer
    '比較元シートのセルと同じセルが比較先シートにある場合、比較先シートの該当行を削除
    For i = 1 To row_source
        Set CheckCells = ws_target.Columns(col_offset).Find(ws_source.Cells(2 + i, "D"))
        If Not (CheckCells Is Nothing) Then
            Worksheets(1).Rows(CheckCells.Row).Delete
        End If
    Next i
    
End Sub
