Attribute VB_Name = "Module1"
Option Explicit

Sub MergeCells()
    Dim i As Integer
    Dim num_row As Integer
    Dim ws As Worksheet
    Dim upper_init_pos As Integer
    Dim under_init_pos As Integer
    Dim left_init_pos As Integer
    Dim right_init_pos As Integer
    Dim num_row_joined_cell As Integer
    Dim num_column_joined_cell
    Set ws = Worksheets(2)
    num_row_joined_cell = ws.Range("C4").Value
    num_column_joined_cell = ws.Range("C5").Value
    num_row = ws.Range("C2").Value
    upper_init_pos = 1
    under_init_pos = num_row_joined_cell
    left_init_pos = 1
    right_init_pos = num_column_joined_cell
    
    's•ûŒü
    Dim row_insertion_position As Integer
    Dim i_row_insert As Integer
    Dim i_column As Integer
    Dim row_offset As Integer
    'hack:1‚Í•Ï”‚ğg‚Á‚Ä‰Â•Ï‚É‚·‚é
    For i = 1 To num_row
        '‘}“ü
        row_insertion_position = 1 + (i - 1) * num_row_joined_cell
        Rows(row_insertion_position).Select
        For i_row_insert = 1 To (num_row_joined_cell - 1)
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Next
        'Œ‹‡
        row_offset = (i - 1) * num_row_joined_cell
        For i_column = 1 To num_row
            Range(Cells(upper_init_pos + row_offset, i_column), Cells(under_init_pos + row_offset, i_column)).Select
            Selection.Merge
        Next
    Next
    
    '—ñ•ûŒü
    Dim column_insertion_position As Integer
    Dim i_column_insert As Integer
    Dim i_row As Integer
    Dim column_offset As Integer
    Dim num_column As Integer
    num_column = ws.Range("C3").Value
    'hack:1‚Í•Ï”‚ğg‚Á‚Ä‰Â•Ï‚É‚·‚é
    For i = 1 To num_column
        '‘}“ü
        column_insertion_position = 1 + (i - 1) * num_column_joined_cell
        Columns(column_insertion_position).Select
        For i_column_insert = 1 To (num_column_joined_cell - 1)
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Next
        'Œ‹‡
        column_offset = (i - 1) * num_column_joined_cell
        For i_row = 1 To num_column
            row_offset = (i_row - 1) * num_row_joined_cell
            Range(Cells(upper_init_pos + row_offset, left_init_pos + column_offset), Cells(upper_init_pos + row_offset, right_init_pos + column_offset)).Select
            Selection.Merge
        Next
        Range(Cells(left_init_pos + column_offset, 2), Cells(left_init_pos + column_offset, 2)).Select
        Selection.Merge
    Next
End Sub
