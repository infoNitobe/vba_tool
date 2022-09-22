Attribute VB_Name = "Module2"
'全モジュールエクスポート
'
'Excelの設定を以下の通りに変更すること
'１）オプション -> セキュリティーセンター -> [セキュリティーセンターの設定]ボタン押下
'２）マクロ設定（左ペイン） -> [VBAプロジェクトオブジェクトモデルへのアクセスを信頼する]　チェックON
Public Sub ExportAllModule()
    Dim destDir As String
    
    With CreateObject("WScript.Shell")
        destDir = .SpecialFolders("Desktop") & "\modules"
    End With
    If Dir(destDir, vbDirectory) = "" Then
        Call MkDir(destDir)
    End If
    
    With ActiveWorkbook.VBProject
        Const vbext_ct_StdModule As Variant = 1
        Const vbext_ct_MSForm As Variant = 2
        Const vbext_ct_ClassModule As Variant = 3
        
        Dim ext As String
        Dim c As Object
        For Each c In .VBComponents
            Select Case c.Type
            Case vbext_ct_StdModule
                ext = ".bas"
            Case vbext_ct_MSForm
                ext = ".frm"
            Case vbext_ct_ClassModule
                ext = ".cls"
            Case Else
                ext = Empty
            End Select
            
            If ext <> Empty Then
                Call c.Export(destDir & "\" & c.Name & ext)
            End If
            
        Next
    
    End With
    
    MsgBox "モジュールのエクスポートを完了しました。" & vbNewLine & destDir
End Sub
