'Sub ExportAll(target_file_path As String, output_folder_path As String)

'����
'<target_file_path>�o�͑Ώۂ̃t�@�C�����w�肵�Ă��������B
'<output_folder_path>�o�͐���w�肵�Ă��������B


Dim objParams, strFullPath, strFileName, objExcel, objWorkBook
Dim objTempComponent, strCode
Dim strExportPath
Dim FSO

strFullPath = ""
strExportPath = ""
strFileName = ""
strFilePath = ""

Set FSO = CreateObject("Scripting.FileSystemObject")
Set objParams = WScript.Arguments



'If objParams.Count <> 2 Then

'	Msgbox "����������܂���B�o�͐���w�肵�Ă��������B"
'	WScript.Quit 0
	
'end if


strFullPath = objParams.item(0)     '�G�N�X�|�[�g�Ώۂ̃t�@�C���p�X
'strFullPath = target_file_path
    
strExportPath = objParams.Item(1)  '�G�N�X�|�[�g��̃p�X�������Ŏw�肷��B
'strExportPath = "C:\Users\yasuhisa-sasahara\Documents\dev\src_text"

strFileName = FSO.GetFileName(strFullPath)
strFilePath = FSO.GetParentFolderName(strFullPath)
'WScript.Echo "strFullPath---->" & strFullPath
'WScript.Echo "strFileName---->" & strFileName
'WScript.Echo "strFilePath---->" & strFilePath
'WScript.Echo "strExportPath---->" & strExportPath





'Excel���쏀��
Set objExcel = CreateObject("Excel.Application")

'��Ԃ�ύX����B
objExcel.Visible = False
objExcel.DisplayAlerts = False
objExcel.EnableEvents = False

'�}�N���������̏�ԂŊJ��
'�����߂������I�IobjExcel.AutomationSecurity = msoAutomationSecurityForceDisable
Set objWorkBook = objExcel.Workbooks.Open(strFullPath)


'�\�[�X���G�N�X�|�[�g����
Call ExportSource(objWorkBook, strExportPath)

'Excel���N���[�Y
Set FSO = Nothing
Set objParams = Nothing

'��Ԃ�߂�
objExcel.DisplayAlerts = True
objExcel.EnableEvents = True
objWorkBook.Close False



objExcel.Quit
Set objWorkBook = Nothing
Set objExcel = Nothing

'End Sub


'--------------------------------------------------------------------------
'�\�[�X���G�N�X�|�[�g����
'--------------------------------------------------------------------------
Sub ExportSource(ByRef objWorkBook, strExportPath)

    Dim ErrNumber

    On Error Resume Next
    Set a = objWorkBook.VBProject.VBComponents
    ErrNumber = Err.Number
    On Error GoTo 0
    
    If ErrNumber = 50289 Then
        'OpenProject ("erp3707")
    End If
    

    For Each TempComponent In objWorkBook.VBProject.VBComponents
        If TempComponent.CodeModule.CountOfDeclarationLines <> TempComponent.CodeModule.CountOfLines Then
        
            Select Case TempComponent.Type
                'STANDARD_MODULE
                Case 1
                    TempComponent.Export strExportPath & "\" & objWorkBook.Name & "_" & TempComponent.Name & ".bas"
                'CLASS_MODULE
                Case 2
                    TempComponent.Export strExportPath & "\" & objWorkBook.Name & "_" & TempComponent.Name & ".cls"
                'USER_FORM
                Case 3
                    TempComponent.Export strExportPath & "\" & objWorkBook.Name & "_" & TempComponent.Name & ".frm"
                'SHEET��ThisWorkBook
                Case 100
                    TempComponent.Export strExportPath & "\" & objWorkBook.Name & "_" & TempComponent.Name & ".bas"
                Case Else
                	Msgbox TempComponent.Name
            End Select
            With TempComponent.CodeModule
                'Code = .Lines(1, .CountOfLines)
                'Code = .Lines(.CountOfDeclarationLines + 1, .CountOfLines - .CountOfDeclarationLines + 1)
            End With
            
        Else
            
            'DeclareStatement�Ȃǂ͂�������ʂ�
            '���̂��߂ɔ�r���Ă��邩�킩��Ȃ�
            '�����炭�A�R�[�h���Ȃ����W���[���͏Ȃ��Ă��銴��
            'Msgbox TempComponent.Name
            
             Select Case TempComponent.Type
                'STANDARD_MODULE
                Case 1
                    TempComponent.Export strExportPath & "\" & objWorkBook.Name & "_" & TempComponent.Name & ".bas"
                'CLASS_MODULE
                Case 2
                    TempComponent.Export strExportPath & "\" & objWorkBook.Name & "_" & TempComponent.Name & ".cls"
                'USER_FORM
                Case 3
                    TempComponent.Export strExportPath & "\" & objWorkBook.Name & "_" & TempComponent.Name & ".frm"
                'SHEET��ThisWorkBook
                Case 100
                    TempComponent.Export strExportPath & "\" & objWorkBook.Name & "_" & TempComponent.Name & ".bas"
                Case Else
                	Msgbox TempComponent.Name
            End Select
            With TempComponent.CodeModule
                'Code = .Lines(1, .CountOfLines)
                'Code = .Lines(.CountOfDeclarationLines + 1, .CountOfLines - .CountOfDeclarationLines + 1)
            End With
            
            
        End If
    Next

End Sub