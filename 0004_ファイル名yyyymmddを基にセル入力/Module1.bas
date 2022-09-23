Attribute VB_Name = "Module1"
Option Explicit

Type WritingInfo
    dest_row As Integer
    dest_column As String
    suffix As String
End Type

Sub InputCellFromFileName()
    Dim fso As Object
    Dim fld_path As String
    Dim file_name As String
    Dim fld_name As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    fld_name = ActiveSheet.Cells(3, "B")
    fld_path = ThisWorkbook.Path & "\" & fld_name
    
    'WritingInfoの初期化
    Dim yyyy_Info As WritingInfo
    Dim mm_Info As WritingInfo
    Dim dd_Info As WritingInfo
    yyyy_Info.dest_row = ActiveSheet.Cells(3, "E")
    yyyy_Info.dest_column = ActiveSheet.Cells(4, "E")
    yyyy_Info.suffix = ActiveSheet.Cells(5, "E")
    mm_Info.dest_row = ActiveSheet.Cells(3, "F")
    mm_Info.dest_column = ActiveSheet.Cells(4, "F")
    mm_Info.suffix = ActiveSheet.Cells(5, "F")
    dd_Info.dest_row = ActiveSheet.Cells(3, "G")
    dd_Info.dest_column = ActiveSheet.Cells(4, "G")
    dd_Info.suffix = ActiveSheet.Cells(5, "G")
    
    Dim file As Object
    Dim delimiter_position As Integer
    Dim file_yyyymmdd As String
    Dim file_yyyy As String, file_mm As String, file_dd As String
    Dim wb As Workbook
    With fso
        For Each file In .GetFolder(fld_path).Files
            'ファイル名関連処理
            delimiter_position = InStr(file.Name, "_")
            file_yyyymmdd = Mid(file.Name, delimiter_position + 1, 8)
            file_yyyy = Mid(file_yyyymmdd, 1, 4)
            file_mm = Mid(file_yyyymmdd, 5, 2)
            file_dd = Mid(file_yyyymmdd, 7, 2)
            'ブックに書き込み
            Set wb = Workbooks.Open(file)
            wb.Sheets(1).Cells(yyyy_Info.dest_row, yyyy_Info.dest_column) = file_yyyy + yyyy_Info.suffix
            wb.Sheets(1).Cells(mm_Info.dest_row, mm_Info.dest_column) = file_mm + yyyy_Info.suffix
            wb.Sheets(1).Cells(dd_Info.dest_row, dd_Info.dest_column) = file_dd + yyyy_Info.suffix
            wb.Close SaveChanges:=True
        Next
    End With
End Sub

