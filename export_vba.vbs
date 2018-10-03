'=======================================================================
'使用方法
'=======================================================================
'1 target_file_path:    ソースを出力する対象ファイルを指定してください。
'2 output_folder_path:  出力先を指定してください。

'=======================================================================
'使用時のトラブルシューティング
'=======================================================================
'Q プログラミングによる Visual Basic プロジェクトへのアクセスは信頼性に欠けますとは？
'A Excelのオプション　⇒　セキュリティセンター　⇒　マクロの設定　⇒　VBAプロジェクトオブジェクトへのアクセスを信頼する必要があります。

'=======================================================================
'変数・初期値定義
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
'引数チェック
'=======================================================================
Set objParams = WScript.Arguments

'If objParams.Count <> 2 Then

'	Msgbox "引数が足りません。出力先を指定してください。"
'	WScript.Quit 0
	
'end if

'=======================================================================
'引数設定
'=======================================================================
strFullPath = objParams.item(0)     'エクスポート対象のファイルパス
strExportPath = objParams.Item(1)  'エクスポート先のパスを引数で指定する。

Set FSO = CreateObject("Scripting.FileSystemObject")
strFileName = FSO.GetFileName(strFullPath)
strFilePath = FSO.GetParentFolderName(strFullPath)

Set objParams = Nothing
Set FSO = Nothing

'=======================================================================
'オブジェクト初期化
'=======================================================================
'Excel操作準備
Set objExcel = CreateObject("Excel.Application")

'状態を変更する。
objExcel.Visible = False
objExcel.DisplayAlerts = False
objExcel.EnableEvents = False

'マクロが無効の状態で開く　※うまく動かないので無視
'objExcel.AutomationSecurity = msoAutomationSecurityForceDisable


'=======================================================================
'ソースエクスポート
'=======================================================================
'ソースをエクスポートする
Call ExportSource(strFullPath, strExportPath)

'=======================================================================
'締め処理
'=======================================================================
objExcel.DisplayAlerts = True
objExcel.EnableEvents = True

objExcel.Quit
Set objExcel = Nothing



'--------------------------------------------------------------------------
'ソースをエクスポートする
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
        'OpenProject ("erp3707") なんだっけこれ
    End If

								Dim objFso
								Dim objFile
								
    Set objFso = CreateObject("Scripting.FileSystemObject")
									Set objFile = objFso.OpenTextFile(strExportPath & "\worksheetinfo.txt", 2, True)

If Err.Number > 0 Then
    WScript.Echo "Open Error"
Else
    objFile.WriteLine "書き込む文字列です。"
End If

objFile.Close
Set objFile = Nothing
Set objFso = Nothing

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
                'SHEETとThisWorkBook
                Case 100
                    TempComponent.Export strExportPath & "\" & TempComponent.Name & ".bas"
                'それ以外は想定していないのでエラー
                Case Else
                	Msgbox TempComponent.Name
            End Select

            'コード行数を出そうとしていたのかも
            With TempComponent.CodeModule
                'Code = .Lines(1, .CountOfLines)
                'Code = .Lines(.CountOfDeclarationLines + 1, .CountOfLines - .CountOfDeclarationLines + 1)
            End With
            
        Else
            
            'DeclareStatementなどはこっちを通る
            '何のために比較しているかわからない
            'おそらく、コードがないモジュールは省いている感じ
            
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
                'それ以外は想定していないのでエラー
                Case Else
                	Msgbox TempComponent.Name
            End Select

            'コード行数を出そうとしていたのかも
            With TempComponent.CodeModule
                'Code = .Lines(1, .CountOfLines)
                'Code = .Lines(.CountOfDeclarationLines + 1, .CountOfLines - .CountOfDeclarationLines + 1)
            End With
            
        End If
    Next

    objWorkBook.Close False
    Set objWorkBook = Nothing

End Sub
