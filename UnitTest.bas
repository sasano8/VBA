Attribute VB_Name = "UnitTest"


Public Function AreEquel(ByVal fact, ByVal expect, Optional testname As String)

    If expect = fact Then
        Debug.Print "true"
    Else
        Debug.Print "false"
        Call CatchFail
    End If

End Function

Public Function AreNotEquel(ByVal fact, ByVal expect, Optional testname As String)

    If expect <> fact Then
        Debug.Print "true"
    Else
        Debug.Print "false"
        Call CatchFail
    End If

End Function


Private Function CatchFail()
    Call Err.Raise(-1)
End Function

Public Function DebugMaker(Optional msg As String = "")
    Static count As Long
    
    count = count + 1

    Debug.Print Format(count, "0000") & " : " & Now & "   " & msg
    

End Function

Public Function DebugStop()

    MsgBox "デバッグ　ESCで中断してね"
    Do While True
    Loop

End Function

