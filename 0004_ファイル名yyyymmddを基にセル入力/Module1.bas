Attribute VB_Name = "Module1"
Sub シート内のファイル名からフォルダを生成()
    Dim my_fso As Object, my_path As String
    Set my_fso = CreateObject("Scripting.FileSystemObject")
    my_fld_path = ThisWorkbook.Path
    
    For i = 1 To my_range.Rows.Count
        my_fso.CreateFolder (my_fld_path & "\" & Worksheets(1).Cells(i, 1).Value)
    Next
End Sub
