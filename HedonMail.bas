Attribute VB_Name = "HedonMail"

'Require:
    'Hedon.bas
    'HedonArray.bas
    'HedonString.bas


Public Function SendMail(ByVal QuickSend As Boolean, ByVal toMail As String, ByVal Subject As String, ByVal Body As String, Optional ByVal From As String, Optional ByVal Copy As String, Optional ByVal Attach As String, Optional ByVal Log As Boolean = True, Optional ByVal Percent As Double = 100, Optional Msg As String) As Boolean
    'Odes�l� email pomoc� outlooku
    Dim OutApp As Object, OutMail As Object, vStr As String
    
    SendMail = False
    If Len(toMail) > 0 Then
        On Error GoTo Out
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(olMailItem)
        With OutMail
            .SentOnBehalfOfName = From
            .To = toMail
            .Cc = Copy
            .Subject = Subject
            .Body = Body
            If Len(Attach) > 0 Then .Attachments.Add Attach, olByValue, 1, ""
            .Display
            If QuickSend Then
                .Send
            ElseIf "Ano" = MsgLog(Msg:=SConcatenate(Msg, toMail, ": ") & "Odeslat email?", Priority:="Normal", Title:="Poslat?", Typ:=vbQuestion + vbYesNo) Then
                .Send
            End If
            Sleep (1000)
        End With
        Call Tis.User.AddStat("Mail")
        SendMail = True
    End If
Out:
    If Log Then
        If SendMail Then Call MsgLog(Percent, SConcatenate(Msg, toMail, ": ") & " odesl�n", True, "Normal") Else Call MsgLog(Percent, SConcatenate(Msg, toMail, ": ") & " se nepoda�ilo odeslat!", True, "Major")
    End If
    Set OutMail = Nothing
    Set OutApp = Nothing
End Function

Public Function SendMailTbl(ByVal QuickSend As Boolean, ByVal Table As cTable, Optional ByVal Tag As String, Optional ByVal Log As Boolean = True, Optional ByVal Percent As Double = 100, Optional ByVal Msg As String) As cTable
    'Odes�l� tabulku s emaily, vrac� emaily, kter� se nepoda�ilo obeslat
    'Hledan� popisky sloupc�: "From", "To", "Copy", "Subject", "Attach", "Body"
    
    Dim mFrom As Long, mTo As Long, mCopy As Long, mSubject As Long, mAttach As Long, mBody As Long
    Dim Row As Long, count As Long, cMsg As String, Per As Double
    
    Set SendMailTbl = New cTable
    
    With SendMailTbl
        If Not .Paste(Table) Then GoTo Out
        If Not TablesNotEmpty(.Body) Then GoTo Out
        
        mTo = .FindTagCol("To")
        mBody = .FindTagCol("Body")
        If mTo < 0 Or mBody < 0 Then GoTo Out
        mFrom = .FindTagCol("From")
        mCopy = .FindTagCol("Copy")
        mSubject = .FindTagCol("Subject")
        mAttach = .FindTagCol("Attach")
        
        Row = 0
        Per = GetOnePer(0, .CountRow, Percent)
        If vU(.Body) > 0 Then count = 1
        Call MsgLog(Per, "Odes�l�m " & Inflect("mail", .CountRow, 1))
        Do While TablesNotEmpty(.Body, Row)
            If count > 0 Then
                cMsg = count
                count = count + 1
            End If
            If SendMail(QuickSend, .Cell(Row, mTo), .Cell(Row, mSubject), .Cell(Row, mBody), .Cell(Row, mFrom), .Cell(Row, mCopy), .Cell(Row, mAttach), True, Per, SConcatenate(Msg, cMsg, ".")) Then Call .DelRow(Row) Else Row = Row + 1
        Loop
        Call fOutTbl.Eke(.FullContent, "Neobeslan� emaily", mBody)
    End With
    
    Exit Function
    
Out:
    Call MsgLog(Percent, Msg & "Nenalezeny ��dn� emaily k obesl�n�", True, "Critical", "Emaily nenalezeny!", vbCritical)
    
End Function
