Attribute VB_Name = "HedonLog"

'Require:
    'Hedon.bas
    'HedonArray.bas
    'HedonFile.bas


'Logov�n�
'-------------------------------------------------------
Public Sub Log(ByVal Msg As String, Optional ByVal NewRow As Boolean = True, Optional ByVal Priority As String = "Minor")
    'Priority jsou: Minor, Normal, Major, Critical, Debug
    Dim User As String
    If Len(Msg) > 0 Then
        If NewRow Then
            If Len(Tis.User.Acc) > 0 Then User = Tis.User.Acc Else User = "PC(" & Environ$("Username") & ")"
            Call TextToFile(FilePath("Logs"), Application.ThisWorkbook.Name & Chr(9) & Format(Now, "dd.mm.yyyy hh:nn:ss") & Chr(9) & User & Chr(9) & Priority & Chr(9) & Msg, True, False)
        Else
            Call TextToFile(FilePath("Logs"), Msg, False, False)
        End If
    End If
End Sub

Public Function MsgLog(Optional ByVal Percent As Double = Empty, Optional ByVal Msg As String, Optional ByVal NewRow As Boolean = True, Optional ByVal Priority As String = "Minor", Optional ByVal Title As String, Optional ByVal Typ As Long, Optional Default As Variant = Empty, Optional PasswordChar As String) As String
    Dim BofLogBar As Boolean, Form As Object, Rtrtn As String
    Static MsgText As Variant
    If IsBlank(MsgText) Then MsgText = StrToList("Null;OK;Storno;Zp�t;Znovu;Ignorovat;Ano;Ne", ";")
    
    If Priority = "Major" Then Call Tis.User.AddStat("Zvajda")
    If Priority = "Critical" Then Call Tis.User.AddStat("Zvajda", 10)
    
    For Each Form In VBA.UserForms
        If Form.Name = "fLogBar" Then
            BofLogBar = True
            Exit For
        End If
    Next Form
    
    Call Log(Msg, NewRow, Priority)
    If BofLogBar Then Call fLogBar.Eke(Percent, Msg, NewRow)
    If (((Len(Title) > 0) Or (Typ > 0)) And (Len(Msg) > 0)) Then
        If Typ = 1024 Then
            MsgLog = fInBox.Eke(Msg, Title, Default, PasswordChar)
        Else
            On Error Resume Next
            MsgLog = MsgText(CLng(MsgBox(Msg, Typ, Title)))
        End If
        Call Log(" ..." & HidePass(MsgLog, PasswordChar), False, Priority)
        If BofLogBar Then Call fLogBar.Eke(Percent, " ..." & MsgLog, False)
    End If
End Function
