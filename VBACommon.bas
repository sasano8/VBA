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
'��ԕύX
'############################################################
Public Function StartProcessing()
    Call TransitToProcessingState(True)
End Function

Public Function EndProcessing()
    Call TransitToProcessingState(False)
End Function

'�����J�n����True
'�����I������False
Private Function TransitToProcessingState(ByVal mode As Boolean)
    
    '�����J�n�iTrue�j�̏ꍇ�A
    '�e�@�\�𖳌��ifalse�j�ɂ���B
    '�����I���iFalse�j�̏ꍇ�A
    '�e�@�\��L���iTrue�j�ɂ���B
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

'���O���o�͂��܂��B
'Id�͓����菇�����s����΁A����Id���o�͂���邽�߁A
'����ɂ���������͂��邱�Ƃ��ł��܂��B
'�E�H�b�`���̒ǉ��ȂǂŁA�f�o�b�O����������Id�������Ɏw�肷��Έꎞ��~�ł��܂��B
Public Function LogOutput(ByVal msg As String)

    If IsInitLog = False Then Err.Raise -1, , "���O�o�͂��L��������Ă��܂���B"
    
    If LogId > 1 Then
    
        LogId = LogId + 1
        UpdateStatusBar LogId & " " & Now & ":  " & LogMsgCash & "�@�I��"
    
        LogId = LogId + 1
        UpdateStatusBar LogId & " " & Now & ":  " & msg & "�@�J�n"
        LogMsgCash = msg
    
    ElseIf LogId = 1 Then
    
        LogId = LogId + 1
        UpdateStatusBar LogId & " " & Now & ":  " & msg & "�@�J�n"
        LogMsgCash = msg
    
    ElseIf LogId = 0 Then
        
        LogId = LogId + 1
        UpdateStatusBar "###################################"
        UpdateStatusBar LogId & " " & Now & ":  " & LogHeader & "�@�J�n"
        LogMsgCash = ""
    
    Else
    
    End If
    
End Function

Public Function LogEnd()

    LogId = LogId + 1
    UpdateStatusBar LogId & " " & Now & ":  " & LogMsgCash & "�@�I��"

    LogId = LogId + 1
    UpdateStatusBar LogId & " " & Now & ":  " & LogHeader & "�@�I��"
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


