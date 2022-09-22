Attribute VB_Name = "Module2"
'�S���W���[���G�N�X�|�[�g
'
'Excel�̐ݒ���ȉ��̒ʂ�ɕύX���邱��
'�P�j�I�v�V���� -> �Z�L�����e�B�[�Z���^�[ -> [�Z�L�����e�B�[�Z���^�[�̐ݒ�]�{�^������
'�Q�j�}�N���ݒ�i���y�C���j -> [VBA�v���W�F�N�g�I�u�W�F�N�g���f���ւ̃A�N�Z�X��M������]�@�`�F�b�NON
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
    
    MsgBox "���W���[���̃G�N�X�|�[�g���������܂����B" & vbNewLine & destDir
End Sub
