Attribute VB_Name = "HedonWeb"

'Require:
    'Hedon.bas
    'HedonArray.bas


Public Function Waitfor(ByRef IE As Object, Optional ByVal Wait As Long = 300, Optional ByVal sWait As Boolean, Optional ByVal wKeyWord As String, Optional ByVal wTagBody As String = "", Optional ByVal wTagRow As String = "", Optional ByVal wTagCol As String = "", Optional ByVal wInBody As String = "", Optional ByVal wInRow As String = "", Optional ByVal wInCol As String = "") As String
    Dim i As Long, Cube As Variant, Doc As Variant

    Do
        For i = 1 To Wait
            If ((Not IE.Busy) And (IE.readyState = 4)) Then
                If IE.Document.Title = "Chyba certifik�tu: Navigace je blokov�na." Then
                    Doc = ParseDoc(IE.Document, "A", "Pokra�ovat na tento web (nedoporu�ujeme)")
                    If Not IsBlank(Doc) Then Doc(vL(Doc)).Click
                Else
                    Waitfor = "OK"
                    If Len(wKeyWord & wTagBody & wTagRow & wTagCol & wInBody & wInRow & wInCol) > 0 Then
                        Cube = ParseTable(IE.Document, 2, wTagBody, wTagRow, wTagCol, wInBody, wInRow, wInCol)
                        If CubesNotEmpty(Cube) Then
                            If ((InStr(1, CStr(Cube(0)(0)(0)), wKeyWord, vbTextCompare) > 0) = sWait) Then Exit Function
                        ElseIf Not sWait Then
                            Exit Function
                        End If
                    Else
                        Exit Function
                    End If
                End If
            End If
            Sleep 100
        Next i
        Waitfor = MsgLog(Msg:="Prohl�e� je p��li� dlouho bez odezvy!", Typ:=vbExclamation + vbAbortRetryIgnore, Title:="P��li� dlouho bez odezvy", Priority:="Major")
        If Waitfor = "Znovu" Then IE.Refresh
    Loop Until Waitfor <> "Znovu"
    
End Function

Public Function SetIE(Optional ByVal URL As String, Optional ByRef IE As Object = Nothing, Optional ByVal Visible As Boolean = False, Optional ByVal Wait As Boolean = True, Optional ByVal Clear As Boolean = True, Optional ByVal Height As Long, Optional ByVal Width As Long, Optional ByVal Resizable As Boolean = True) As String
    Dim Default As Boolean, dURL As String
    SetIE = "Zp�t"
    On Error GoTo Reset
    If IE Is Nothing Then
Reset:
        Default = True
        Set IE = Nothing
        Set IE = CreateObject("InternetExplorer.Application")
        If Len(URL) = 0 Then URL = "/"
    Else
        dURL = CStr(IE.Document.URL)
    End If
    With IE
        .Visible = Visible Or Debugmode
        .Silent = Visible
        If Len(URL) > 0 And dURL <> URL Then .Navigate URL
        If Default Then .AddressBar = Not Clear Or Debugmode
        If Default Then .MenuBar = Not Clear Or Debugmode
        If Default Then .Toolbar = Not Clear Or Debugmode
        If Default Then .Resizable = Resizable Or Debugmode
        If Width > 0 Then .Width = Width
        If Height > 0 Then .Height = Height
        If Wait Then SetIE = Waitfor(IE) Else SetIE = "OK"
        .Visible = Visible Or Debugmode
    End With
End Function

Public Function ParseDoc(ByVal Document As Object, ByVal ParseTag As String, Optional ByVal InDoc As String, Optional ByVal FeedBack As Integer = 0, Optional ByVal count As Long = 0) As Variant
    'Parsuje html dokument (Document), podle zna�ky (ParseTag), podle krit�ria (InDoc), a maxim�ln�ho po�tu (count)
    'Feedback ur�uje co se vrac� 1 = text, 2 = cel� html, ostatn� - cel� object
    Dim VarInx As Variant, i As Long, Doc As Long, Inx As Object
    
    If Not Document Is Nothing Then
        i = 0
        For Each Inx In Document.getElementsByTagName(ParseTag)
            If ((Len(InDoc) = 0) Or (InStr(1, Inx.outerhtml, InDoc)) > 0) Then
                Call VaRedim(VarInx, i)
                If FeedBack = 1 Then
                    VarInx(i) = STrim(Inx.innertext)
                ElseIf FeedBack = 2 Then
                    VarInx(i) = Inx.outerhtml
                Else
                    Set VarInx(i) = Inx
                End If
                i = i + 1
                If i = count Then Exit For
            End If
        Next Inx
    End If
    ParseDoc = VarInx
End Function

Public Function ParseTable(ByVal Document As Object, Optional ByVal FeedBack As Integer = 0, Optional ByVal TagBody As String = "tbody", Optional ByVal tagRow As String = "tr", Optional ByVal TagColumn As String = "td", Optional ByVal InBody As String, Optional ByVal InRow As String, Optional ByVal InColumn As String)
    'Parsuje html dokument (Document) podle t�� vno�en�ch zna�ek a t�i vno�en�ch krit�ri� do podoby sou�adnicov� tabulky.
    'Vrac� vno�en� seznam kde prvn� ��slo ozna�uje po�ad� tabulky, druh� po�ad� ��dku a t�et� po�ad� sloupce
    'Feedback ur�uje co se vrac� 1 = text, 2 = cel� html, ostatn� - cel� object

    Dim VarInx As Variant, tV As Variant, rV As Variant
    
    VarInx = ParseDoc(Document, TagBody, InBody)
    If IsBlank(VarInx) Then Exit Function
    
    On Error Resume Next
    For Body = vL(VarInx) To vU(VarInx)
        VarInx(Body) = ParseDoc(VarInx(Body), tagRow, InRow)
        For Row = vL(VarInx(Body)) To vU(VarInx(Body))
            rV = ParseDoc(VarInx(Body)(Row), TagColumn, InColumn, FeedBack)
            If Not IsBlank(rV) Then Call VarAdd(tV, rV)
        Next Row
        If Not IsBlank(tV) Then Call VarAdd(ParseTable, tV)
    Next Body
    
End Function

Public Function SClick(ByVal Element As Object, Optional ByVal Events As Boolean = False)
    If Not Events Then Element.removeAttribute ("onclick")
    If Not Events Then Element.setAttribute "onclick", "return true"
    Element.Click
End Function

Public Function SetWebForm(ByRef IE As Object, Optional ByVal InTag As String = "", Optional ByVal Handle As String = "", Optional ByVal Tag As String = "input", Optional ByVal Document As Object, Optional ByVal Events As Boolean = Empty, Optional ByVal Wait As Long = 50) As String
    'Vypl�uje webov� formul�� "Document", kde vypln� pole ozna�en� Tagem a up�esn�n� "InTag" �et�zcem.
    'Pokud je Handle pr�zdn� pak pouze klik�, Wait slou�� jako prodleva p�ed dal��m pokusem
    Dim List As Variant, i As Long
    
    Do
        For i = 1 To Wait
            On Error GoTo Out
            SetWebForm = Waitfor(IE)
            If SetWebForm <> "OK" Then Exit Function
            If Document Is Nothing Then List = ParseDoc(IE.Document, Tag, InTag, 0, 1) Else List = ParseDoc(Document, Tag, InTag, 0, 1)
            If VarsNotEmpty(List, 0) Then
                If Len(Handle) = 0 Then
                    If Events <> True Then
                        List(vL(List)).removeAttribute ("onclick")
                        List(vL(List)).setAttribute "onclick", "return true"
                    End If
                    List(vL(List)).Click
                Else
                    List(vL(List)).Value = Handle
                    If Events <> False Then List(vL(List)).FireEvent ("onchange")
                End If
                SetWebForm = Waitfor(IE)
                Exit Function
            End If
            Sleep (200)
        Next i
Out:
        SetWebForm = MsgLog(Msg:="Selhala komunikace s prohl�e�em", Typ:=vbExclamation + vbAbortRetryIgnore, Title:="P��li� dlouho bez odezvy", Priority:="Major")
    Loop Until SetWebForm <> "Znovu"
End Function


Public Sub IEQuit(ByRef IE As Object, Optional ByVal CloseMethod As Boolean = True)
    'Vypne bezpe�n� IE, nebo jej zviditeln�
    If Not IE Is Nothing Then
        If CloseMethod Then
            On Error GoTo skip
            IE.Quit
        Else
            IE.Visible = True
        End If
skip:
        Set IE = Nothing
    End If
End Sub

