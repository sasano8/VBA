Attribute VB_Name = "VBACommon"
Option Explicit

Private IsInitLog As Boolean
Private LogId As Long
Private LogHeader As String
Private LogMsgCash As String

Public Function ExistWorkbook(ByVal StrBookName As String) As Boolean

    Dim wb As Workbook
    Dim IsExist As Boolean
    
    For Each wb In Workbooks
        If wb.Name = StrBookName Then
            IsExist = True
            Exit For
        End If
    Next
    
    ExistWorkbook = IsExist

End Function



'############################################################
'状態変更
'############################################################
Public Function StartProcessing()
    Call TransitToProcessingState(True)
End Function

Public Function EndProcessing()
    Call TransitToProcessingState(False)
End Function

'処理開始時はTrue
'処理終了時はFalse
Private Function TransitToProcessingState(ByVal mode As Boolean)
    
    '処理開始（True）の場合、
    '各機能を無効（false）にする。
    '処理終了（False）の場合、
    '各機能を有効（True）にする。
    mode = Not mode
    
    application.ScreenUpdating = mode
    application.EnableEvents = mode
    application.DisplayAlerts = mode
    
    If mode Then
        application.Calculation = xlCalculationAutomatic
    Else
        application.Calculation = xlCalculationManual
    End If

End Function


Private Function OutLogInit(ByVal header As String)

    LogId = 0
    LogMsgCash = ""
    LogHeader = header
    IsInitLog = True
    
End Function

Public Function LogStart(ByVal header As String)

    Call OutLogInit(header)
    Call LogOutput(header)

End Function

'ログを出力します。
'Idは同じ手順を実行すれば、同じIdが出力されるため、
'それにより問題個所を解析することができます。
'ウォッチ式の追加などで、デバッグしたい個所のIdを条件に指定すれば一時停止できます。
Public Function LogOutput(ByVal msg As String)

    If IsInitLog = False Then Err.Raise -1, , "ログ出力が有効化されていません。"
    
    If LogId > 1 Then
    
        LogId = LogId + 1
        UpdateStatusBar LogId & " " & Now & ":  " & LogMsgCash & "　終了"
    
        LogId = LogId + 1
        UpdateStatusBar LogId & " " & Now & ":  " & msg & "　開始"
        LogMsgCash = msg
    
    ElseIf LogId = 1 Then
    
        LogId = LogId + 1
        UpdateStatusBar LogId & " " & Now & ":  " & msg & "　開始"
        LogMsgCash = msg
    
    ElseIf LogId = 0 Then
        
        LogId = LogId + 1
        UpdateStatusBar "###################################"
        UpdateStatusBar LogId & " " & Now & ":  " & LogHeader & "　開始"
        LogMsgCash = ""
    
    Else
    
    End If
    
End Function

Public Function LogEnd()

    LogId = LogId + 1
    UpdateStatusBar LogId & " " & Now & ":  " & LogMsgCash & "　終了"

    LogId = LogId + 1
    UpdateStatusBar LogId & " " & Now & ":  " & LogHeader & "　終了"
    UpdateStatusBar "###################################"
    UpdateStatusBar ""
    
    LogMsgCash = ""

    LogId = 0
    LogMsgCash = ""
    LogHeader = ""
    IsInitLog = False

End Function

Private Function UpdateStatusBar(ByVal msg As String)
    Debug.Print msg
    application.StatusBar = msg
End Function


