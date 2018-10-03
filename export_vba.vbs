'=======================================================================
'�g�p���@
'=======================================================================
'1 target_file_path:    �\�[�X���o�͂���Ώۃt�@�C�����w�肵�Ă��������B
'2 output_folder_path:  �o�͐���w�肵�Ă��������B

'=======================================================================
'�g�p���̃g���u���V���[�e�B���O
'=======================================================================
'Q �v���O���~���O�ɂ�� Visual Basic �v���W�F�N�g�ւ̃A�N�Z�X�͐M�����Ɍ����܂��Ƃ́H
'A Excel�̃I�v�V�����@�ˁ@�Z�L�����e�B�Z���^�[�@�ˁ@�}�N���̐ݒ�@�ˁ@VBA�v���W�F�N�g�I�u�W�F�N�g�ւ̃A�N�Z�X��M������K�v������܂��B

'=======================================================================
'�ϐ��E�����l��`
'=======================================================================
Dim objParams, strFullPath, strFileName, objExcel
Dim objTempComponent, strCode
Dim strExportPath
Dim FSO

strFullPath = ""
strExportPath = ""
strFileName = ""
strFilePath = ""

'=======================================================================
'�����`�F�b�N
'=======================================================================
Set objParams = WScript.Arguments

'If objParams.Count <> 2 Then

'	Msgbox "����������܂���B�o�͐���w�肵�Ă��������B"
'	WScript.Quit 0
	
'end if

'=======================================================================
'�����ݒ�
'=======================================================================
strFullPath = objParams.item(0)     '�G�N�X�|�[�g�Ώۂ̃t�@�C���p�X
strExportPath = objParams.Item(1)  '�G�N�X�|�[�g��̃p�X�������Ŏw�肷��B

Set FSO = CreateObject("Scripting.FileSystemObject")
strFileName = FSO.GetFileName(strFullPath)
strFilePath = FSO.GetParentFolderName(strFullPath)

Set objParams = Nothing
Set FSO = Nothing

'=======================================================================
'�I�u�W�F�N�g������
'=======================================================================
'Excel���쏀��
Set objExcel = CreateObject("Excel.Application")

'��Ԃ�ύX����B
objExcel.Visible = False
objExcel.DisplayAlerts = False
objExcel.EnableEvents = False

'�}�N���������̏�ԂŊJ���@�����܂������Ȃ��̂Ŗ���
'objExcel.AutomationSecurity = msoAutomationSecurityForceDisable


'=======================================================================
'�\�[�X�G�N�X�|�[�g
'=======================================================================
'�\�[�X���G�N�X�|�[�g����
Call ExportSource(strFullPath, strExportPath)

'=======================================================================
'���ߏ���
'=======================================================================
objExcel.DisplayAlerts = True
objExcel.EnableEvents = True

objExcel.Quit
Set objExcel = Nothing



'--------------------------------------------------------------------------
'�\�[�X���G�N�X�|�[�g����
'--------------------------------------------------------------------------
Sub ExportSource(strFullPath, strExportPath)

    Dim ErrNumber
    Dim objWorkBook

    Set objWorkBook = objExcel.Workbooks.Open(strFullPath)

    On Error Resume Next
    Set a = objWorkBook.VBProject.VBComponents
    ErrNumber = Err.Number
    On Error GoTo 0
    
    If ErrNumber = 50289 Then
        'OpenProject ("erp3707") �Ȃ񂾂�������
    End If
    

    For Each TempComponent In objWorkBook.VBProject.VBComponents
        If TempComponent.CodeModule.CountOfDeclarationLines <> TempComponent.CodeModule.CountOfLines Then
        
            Select Case TempComponent.Type
                'STANDARD_MODULE
                Case 1
                    TempComponent.Export strExportPath & "\" & TempComponent.Name & ".bas"
                'CLASS_MODULE
                Case 2
                    TempComponent.Export strExportPath & "\" & TempComponent.Name & ".cls"
                'USER_FORM
                Case 3
                    TempComponent.Export strExportPath & "\" & TempComponent.Name & ".frm"
                'SHEET��ThisWorkBook
                Case 100
                    TempComponent.Export strExportPath & "\" & TempComponent.Name & ".bas"
                '����ȊO�͑z�肵�Ă��Ȃ��̂ŃG���[
                Case Else
                	Msgbox TempComponent.Name
            End Select

            '�R�[�h�s�����o�����Ƃ��Ă����̂���
            With TempComponent.CodeModule
                'Code = .Lines(1, .CountOfLines)
                'Code = .Lines(.CountOfDeclarationLines + 1, .CountOfLines - .CountOfDeclarationLines + 1)
            End With
            
        Else
            
            'DeclareStatement�Ȃǂ͂�������ʂ�
            '���̂��߂ɔ�r���Ă��邩�킩��Ȃ�
            '�����炭�A�R�[�h���Ȃ����W���[���͏Ȃ��Ă��銴��
            
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
                '����ȊO�͑z�肵�Ă��Ȃ��̂ŃG���[
                Case Else
                	Msgbox TempComponent.Name
            End Select

            '�R�[�h�s�����o�����Ƃ��Ă����̂���
            With TempComponent.CodeModule
                'Code = .Lines(1, .CountOfLines)
                'Code = .Lines(.CountOfDeclarationLines + 1, .CountOfLines - .CountOfDeclarationLines + 1)
            End With
            
        End If
    Next

    objWorkBook.Close False
    Set objWorkBook = Nothing

End Sub
