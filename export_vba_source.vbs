'Sub ExportAll(target_file_path As String, output_folder_path As String)

'引数
'<target_file_path>出力対象のファイルを指定してください。
'<output_folder_path>出力先を指定してください。


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

'	Msgbox "引数が足りません。出力先を指定してください。"
'	WScript.Quit 0
	
'end if


strFullPath = objParams.item(0)     'エクスポート対象のファイルパス
'strFullPath = target_file_path
    
strExportPath = objParams.Item(1)  'エクスポート先のパスを引数で指定する。
'strExportPath = "C:\Users\yasuhisa-sasahara\Documents\dev\src_text"

strFileName = FSO.GetFileName(strFullPath)
strFilePath = FSO.GetParentFolderName(strFullPath)
'WScript.Echo "strFullPath---->" & strFullPath
'WScript.Echo "strFileName---->" & strFileName
'WScript.Echo "strFilePath---->" & strFilePath
'WScript.Echo "strExportPath---->" & strExportPath





'Excel操作準備
Set objExcel = CreateObject("Excel.Application")

'状態を変更する。
objExcel.Visible = False
objExcel.DisplayAlerts = False
objExcel.EnableEvents = False

'マクロが無効の状態で開く
'※だめだった！！objExcel.AutomationSecurity = msoAutomationSecurityForceDisable
Set objWorkBook = objExcel.Workbooks.Open(strFullPath)


'ソースをエクスポートする
Call ExportSource(objWorkBook, strExportPath)

'Excelをクローズ
Set FSO = Nothing
Set objParams = Nothing

'状態を戻す
objExcel.DisplayAlerts = True
objExcel.EnableEvents = True
objWorkBook.Close False



objExcel.Quit
Set objWorkBook = Nothing
Set objExcel = Nothing

'End Sub


'--------------------------------------------------------------------------
'ソースをエクスポートする
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
                'SHEETとThisWorkBook
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
            
            'DeclareStatementなどはこっちを通る
            '何のために比較しているかわからない
            'おそらく、コードがないモジュールは省いている感じ
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
                'SHEETとThisWorkBook
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