Attribute VB_Name = "HedonArray"

'Require:
    'Hedon.bas
    'HedonString.bas


Public Function VarsNotEmpty(ByVal Var As Variant, Optional ByVal Bound As Long = -1) As Boolean
    If Not IsBlank(Var) Then VarsNotEmpty = ((Bound < 0) Or ((Bound <= vU(Var)) And (Bound >= vL(Var)))) Else VarsNotEmpty = False
End Function

Public Function TablesNotEmpty(ByVal Table As Variant, Optional ByVal Bound1 As Long = -1, Optional ByVal Bound2 As Long = -1) As Boolean
    TablesNotEmpty = False
    If VarsNotEmpty(Table, Bound1) Then
        If Bound1 < 0 Then Bound1 = vL(Table)
        TablesNotEmpty = VarsNotEmpty(Table(Bound1), Bound2)
    End If
End Function

Public Function CubesNotEmpty(ByVal Cube As Variant, Optional ByVal Bound1 As Long = -1, Optional ByVal Bound2 As Long = -1, Optional ByVal Bound3 As Long = -1) As Boolean
    CubesNotEmpty = False
    If VarsNotEmpty(Cube, Bound1) Then
        If Bound1 < 0 Then Bound1 = vL(Cube)
        CubesNotEmpty = TablesNotEmpty(Cube(Bound1), Bound2, Bound3)
    End If
End Function

Public Function VeVar(ByVal Num As Long, ByVal Ads As Variant, Optional Rescue As Long = -1, Optional Default As String) As String
    'Vrac� hodnotu z pole, pokud nenajde pos� se nastavit rescue (-2 prvn� hodnota, -3 posledn� hodnota) jinak nastav� default
    VeVar = Default
    If Rescue = -2 Then Rescue = vL(Ads)
    If Rescue = -3 Then Rescue = vU(Ads)
    If Num > -1 And VarsNotEmpty(Ads, Num) Then
        VeVar = CStr(Ads(Num))
    ElseIf Rescue > -1 And VarsNotEmpty(Ads, Rescue) Then
        VeVar = CStr(Ads(Rescue))
    End If
End Function

Public Function StrToList(ByVal Text As String, Optional ByVal Delimiter As String, Optional STrims As Boolean = True) As Variant
    'P�evede List na Text, jako odd�lova� pou��v� StrDelimiter, defaultn� enter
    'Je mo�n� pou��vat na vstupu delimiteru funkci Ascii(#)
    Dim vList As Variant, Znak As Long
    Dim zac As Long, kon As Long
    zac = 1
    kon = 1
    
    If Len(Delimiter) = 0 Then Delimiter = Chr(10)
    
    Text = Trim(Text)
    While zac <= Len(Text)
        kon = InStr(zac, Text, Delimiter, vbBinaryCompare)
        If ((kon = 0) And (zac <= Len(Text))) Then kon = Len(Text) + 1
        If kon - zac >= 0 Then
            If STrims Then Call VarSet(vList, STrim(Mid(Text, zac, kon - zac))) Else Call VarSet(vList, Trim(Mid(Text, zac, kon - zac)))
        End If
        zac = kon + 1
    Wend
    StrToList = vList
End Function

Public Function ListToStr(ByVal List As Variant, Optional ByVal Delimiter As String) As String
    'P�evede List na Text, jako odd�lova� pou��v� StrDelimiter, defaultn� enter
    'Je mo�n� pou��vat na vstupu delimiteru funkci Ascii(#)
    Dim Inx As Long
    
    If IsBlank(List) Then Exit Function
    If Len(Delimiter) = 0 Then Delimiter = vbCrLf
    
    For Inx = vL(List) To vU(List)
        ListToStr = SConcatenate(ListToStr, List(Inx), Delimiter)
    Next Inx
End Function

Public Function TableToStr(ByVal Table As Variant, Optional ByVal RowDelimiter As String, Optional ByVal ColDelimiter As String) As String
    'Konvertuje Table do Stringu
    Dim Row As Long
    
    If Len(ColDelimiter) = 0 Then ColDelimiter = Chr(9)
    If Len(RowDelimiter) = 0 Then RowDelimiter = vbCrLf
    
    If TablesNotEmpty(Table) Then
        For Row = vL(Table) To vU(Table)
            TableToStr = SConcatenate(TableToStr, ListToStr(Table(Row), ColDelimiter), RowDelimiter, True)
        Next Row
    End If
    
End Function

Public Function StrToTable(ByVal str As String, Optional ByVal RowDelimiter As String, Optional ByVal ColDelimiter As String, Optional ByVal ColEmbed As Boolean = False, Optional ByVal ColSensitivity As Integer) As Variant
    'Konvertuje String do Tablu
    Dim Row As Long, Var As Variant
    
    If Len(str) = 0 Then Exit Function
    
    Var = StrToList(str, RowDelimiter, False)
    
    If ColEmbed Then
        If Len(ColDelimiter) = 0 Then ColDelimiter = " "
        Call DeEmbedList(Var, ColDelimiter, ColSensitivity)
    Else
        If Len(ColDelimiter) = 0 Then ColDelimiter = Chr(9)
        For Row = vL(Var) To vU(Var)
            Var(Row) = StrToList(Var(Row), ColDelimiter)
        Next Row
    End If
    StrToTable = Var
End Function


Public Sub ReadTable(Table As Variant)
    If Not IsBlank(Table) Then
        For X = vL(Table) To vU(Table)
            If Not IsBlank(Table(X)) Then
                For y = vL(Table(X)) To vU(Table(X))
                   Call MsgBox(Table(X)(y), vbOKOnly, "X=" & X & "; Y=" & y)
                Next y
            End If
        Next X
    End If
End Sub

Public Sub ReadCube(Cube As Variant)
    If Not IsBlank(Cube) Then
        For X = vL(Cube) To vU(Cube)
            If Not IsBlank(Cube(X)) Then
                For y = vL(Cube(X)) To vU(Cube(X))
                    If Not IsBlank(Cube(X)(y)) Then
                        For z = vL(Cube(X)(y)) To vU(Cube(X)(y))
                            er = MsgBox(Cube(X)(y)(z), vbOKOnly, "X=" & X & "; Y=" & y & "; Z=" & z)
                        Next z
                    End If
                Next y
            End If
        Next X
    End If
End Sub

Public Function VarConcatenate(ByVal FirstTable As Variant, ByVal SecondTable As Variant) As Variant
    'Spoj� dva listy do jednoho
    If Not IsBlank(SecondTable) Then
        For Each Inx In SecondTable
            Call VarSet(FirstTable, Inx)
        Next Inx
    End If
    VarConcatenate = FirstTable
End Function

Public Function VarMake(Optional ByVal Arg0 = Empty, Optional ByVal Arg1 = Empty, Optional ByVal Arg2 = Empty, Optional ByVal Arg3 = Empty, Optional ByVal Arg4 = Empty, Optional ByVal Arg5 = Empty, Optional ByVal Arg6 = Empty, Optional ByVal Arg7 = Empty, Optional ByVal Arg8 = Empty, Optional ByVal Arg9 = Empty, Optional ByVal Arg10 = Empty) As Variant
    'Se�ad� argumenty do variant
    Dim nTbl As Variant
    If Not IsBlank(Arg0) Then Call VarSet(nTbl, Arg0, 0)
    If Not IsBlank(Arg1) Then Call VarSet(nTbl, Arg1, 1)
    If Not IsBlank(Arg2) Then Call VarSet(nTbl, Arg2, 2)
    If Not IsBlank(Arg3) Then Call VarSet(nTbl, Arg3, 3)
    If Not IsBlank(Arg4) Then Call VarSet(nTbl, Arg4, 4)
    If Not IsBlank(Arg5) Then Call VarSet(nTbl, Arg5, 5)
    If Not IsBlank(Arg6) Then Call VarSet(nTbl, Arg6, 6)
    If Not IsBlank(Arg7) Then Call VarSet(nTbl, Arg7, 7)
    If Not IsBlank(Arg8) Then Call VarSet(nTbl, Arg8, 8)
    If Not IsBlank(Arg9) Then Call VarSet(nTbl, Arg9, 9)
    If Not IsBlank(Arg10) Then Call VarSet(nTbl, Arg10, 10)
    VarMake = nTbl
End Function

Public Function VarGet(ByRef Var As Variant, Optional ByVal Num As Long = -1) As Variant

    If Num < 0 Then Num = vU(Var) + 1 + Num
    VarGet = Var(Num)

End Function

Public Function VarSet(ByRef Var As Variant, ByVal Add, Optional ByVal AsNum As Long = -1) As Variant
    'Nahrazuje polo�ku seznamu na konec nebo na AsNum
    Call VaRedim(Var, AsNum)

    If IsObject(Add) Then Set Var(AsNum) = Add Else Var(AsNum) = Add
    VarSet = Var
End Function

Public Function VarAdd(ByRef Var As Variant, ByVal Add, Optional ByVal AsNum As Long = -1) As Variant
    'Vkl�d� polo�ku do seznamu na pozici AsNum a posouv� ostatn� polo�ky d�l
    Dim i As Long
    
    If AsNum > -1 And AsNum <= vU(Var) Then
        Call VaRedim(Var)
        For i = vU(Var) To (AsNum + 1) Step -1
            VarSet Var, Var(i - 1), i
        Next i
    End If
    VarAdd = VarSet(Var, Add, AsNum)
End Function
Public Function VarPlace(ByRef Var As Variant, ByVal Add, Optional ByVal AsNum As Long = -1)
    'Zapisuje do seznamu, ale pokud na m�st� ji� existuje hodnota nenahrazuje ani nep�id�v�
    'V defaultn�m nastaven� p�id� hodnotu jen pokud v seznamu dan� hodnota je�t� nen�
    Dim i
    
    If Not IsBlank(Var) Then
        If AsNum < 0 Then
            For Each i In Var
                If i = Add Then Exit Function
            Next i
        ElseIf ((vL(Var) <= AsNum) And (AsNum <= vU(Var))) Then
            If Not IsBlank(Var(AsNum)) Then Exit Function
        End If
    End If
    
    VarPlace = VarSet(Var, Add, AsNum)
    
End Function

Public Function VaRedim(ByRef Var As Variant, Optional ByRef AsNum As Long = -1)
    'Zv�t�i pole o 1 nebo na AsNum
    If IsBlank(Var) Then
        If AsNum < 0 Then AsNum = 0
        ReDim Var(0 To AsNum)
    Else
        If AsNum < 0 Then AsNum = vU(Var) + 1
        If AsNum > vU(Var) Then ReDim Preserve Var(vL(Var) To AsNum)
        If AsNum < vL(Var) Then ReDim Preserve Var(AsNum To vU(Var))
    End If
    VaRedim = Var
End Function

Public Function VarCount(ByRef Var As Variant) As Long
    'Se�te po�et z�znam� v poli Var
    If Not IsBlank(Var) Then VarCount = vU(Var) - vL(Var) + 1
End Function

Public Function VarDel(ByRef Var As Variant, Optional AsNum As Long = -1)
    'Vyma�e ��st AsNum pole Var
    Dim i As Long
    If ((Not IsBlank(Var)) And (AsNum >= 0)) Then
        If ((vL(Var) = AsNum) And (vU(Var) = AsNum)) Then
            Var = Empty
        ElseIf ((AsNum >= vL(Var)) And (AsNum <= vU(Var))) Then
            If AsNum <> vU(Var) Then
                For i = AsNum To vU(Var) - 1
                    Var(i) = Var(i + 1)
                Next i
            End If
            ReDim Preserve Var(vL(Var) To vU(Var) - 1)
        End If
    End If
    VarDel = Var
End Function

Public Function VarFind(ByRef Var As Variant, Find, Optional ByVal Direction As Boolean = True) As Long
    'Prohled� pole a vr�t� prvn� shodu s Find
    'Direction, True = hled� zespodu, False hled� shora
    Dim i As Long, v1 As Long, v2 As Long, vS As Long
    VarFind = -1
    If Not IsBlank(Var) Then
        If Direction Then v1 = vL(Var) Else v1 = vU(Var)
        If Direction Then v2 = vU(Var) Else v2 = vL(Var)
        If Direction Then vS = 1 Else vS = -1
        For i = v1 To v2 Step vS
            If Var(i) = Find Then
                VarFind = i
                Exit Function
            End If
        Next i
    End If
End Function

Public Function VarCompare(ByVal Var1 As Variant, ByVal Var2 As Variant, Optional Full As Boolean)
    Dim Row As Long, vFind As Long
    VarCompare = True
    If Not IsBlank(Var2) Then
        For Row = vU(Var2) To vL(Var2) Step -1
            vFind = VarFind(Var1, Var2(Row), False)
            If vFind = -1 Then
                VarCompare = False
                Exit Function
            ElseIf Full Then
                Call VarDel(Var1, vFind)
            End If
        Next Row
    End If
    If Full And Not IsBlank(Var1) Then VarCompare = False
End Function

Public Function FtrList(ByVal List As Variant, ByVal Filter As Variant) As Variant
    Dim i As Long
    If Not (IsBlank(List) Or IsBlank(Filter)) Then
        i = vL(List)
        Do While i <= vU(List)
            If VarFind(Filter, List(i)) >= 0 Then Call VarDel(List, i) Else i = i + 1
        Loop
    End If
    FtrList = List
End Function

Public Function FtrTable(ByVal Table As Variant, Optional ByVal WantRow As Variant = Empty, Optional ByVal WantCol As Variant = Empty, Optional ByVal InRow As String, Optional ByVal ExRow As String, Optional ByVal InCol As String, Optional ByVal ExCol As String) As Variant
    Dim Row As Long, Col As Long, Add As Boolean
    
    If Not TablesNotEmpty(Table) Then Exit Function
    
    'Chyb� filtrov�n� �et�zce v sloupci InCol & ExCol
    If IsBlank(WantCol) Then
        For Col = vL(Table(vL(Table))) To vU(Table(vL(Table)))
            Call VarSet(WantCol, Col)
        Next Col
    End If
    For Row = vL(Table) To vU(Table)
        Want = vL(WantRow)
        Do While VarsNotEmpty(WantRow, Want)
            If Row = CLng(GetNum(CStr(WantRow(Want)))) Then Exit Do
            Want = Want + 1
        Loop
        If Want <= vU(WantRow) And SInStr(ListToStr(Table(Row), ""), InRow, ExRow) Then
            varCol = Empty
            For Col = vL(Table(Row)) To vU(Table(Row))
                Want = vL(WantCol)
                Do While VarsNotEmpty(WantCol, Want)
                    If Col = CLng(GetNum(CStr(WantCol(Want)))) Then Exit Do
                    Want = Want + 1
                Loop
                If Want <= vU(WantCol) Then Call VarAdd(varCol, Table(Row)(Col))
            Next Col
            If Not IsBlank(varCol) Then Call VarAdd(FtrTable, varCol)
        End If
    Next Row
End Function

Public Function vL(ByVal Var As Variant) As Long
    If IsBlank(Var) Then vL = -1 Else vL = LBound(Var)
End Function

Public Function vU(ByVal Var As Variant) As Long
    If IsBlank(Var) Then vU = -1 Else vU = UBound(Var)
End Function

Public Function CutTable(ByVal Table As Variant) As Variant
    'O��zne Table na tabulku od ��dku 0 a sloupce 0
    Dim Row As Long, Col As Long, MinL As Long, MaxL As Long, MinU As Long, MaxU As Long
    Dim nCol As Long, nRow As Variant
    
    If TablesNotEmpty(Table) Then
        Table = CutVar(Table)
        MinL = vL(Table(vL(Table)))
        MinU = vL(Table(vU(Table)))
        For Row = vL(Table) + 1 To vU(Table)
            MinL = MinNum(MinL, vL(Table(Row)))
            MaxL = MaxNum(MaxL, vL(Table(Row)))
            MinU = MinNum(MinU, vU(Table(Row)))
            MaxU = MaxNum(MaxU, vU(Table(Row)))
        Next Row
        If MaxL > 0 Or MaxU > MinU Then
            For Row = vL(Table) To vU(Table)
                If vU(Table(Row)) < MaxU Then
                    nRow = Table(Row)
                    ReDim Preserve nRow(0 To MaxU)
                    Table(Row) = nRow
                End If
                If vL(Table(Row)) > 0 Then
                    nCol = 0
                    ReDim nRow(0 To MaxU - MinL)
                    For Col = MinL To MaxU
                        If Col >= vL(Table(Row)) Or Col <= vU(Table(Row)) Then Call VarSet(nRow, Table(Row)(Col), nCol)
                        nCol = nCol + 1
                    Next Col
                    Table(Row) = nRow
                End If
            Next Row
        End If
    End If
    CutTable = Table
End Function

Public Function CutVar(ByVal Var As Variant) As Variant
    'O��zne Var na Pole od 0
    Dim outVar As Variant, i As Long
    If Not IsBlank(Var) Then
        If vL(Var) > 0 Then
            For i = vL(Var) To vU(Var)
                Call VarAdd(outVar, Var(i))
            Next i
            CutVar = outVar
        Else
            CutVar = Var
        End If
    End If
    
End Function


Public Function CompressTable(ByRef Table As Variant, Optional ByVal DomCol As Long, Optional ByVal Delimiter As String) As Variant
    'Vy�ist� dvojit� pole (Table) od pr�zdn�ch pol� a duplicit podle dominantn�ho sloupce (DomCol) defaultn� podle prvn�ho.
    'Delimiter nastavuje odd�len� textu mezi spojen�mi ��dky
    Dim vrtRow As Variant, lntCol As Long, vroTab As Variant, lnoRow As Variant, lnoCol As Long, News As Boolean
    
    If Len(Delimiter) = 0 Then Delimiter = ", "
    
    If Not IsBlank(Table) Then
        DomCol = MathFrame(DomCol - 1, vL(Table), vU(Table))
        For Each vrtRow In Table
            If VarsNotEmpty(vrtRow, DomCol) Then
                If Len(Trim(vrtRow(DomCol))) > 0 Then
                    News = True
                    If VarsNotEmpty(vroTab, 0) Then
                        For lnoRow = vL(vroTab) To vU(vroTab)
                            If vroTab(lnoRow)(0) = Trim(vrtRow(DomCol)) Then
                                lnoCol = 1
                                For lntCol = vL(vrtRow) To vU(vrtRow)
                                    If lntCol <> DomCol Then
                                        vroTab(lnoRow)(lnoCol) = SConcatenate(vroTab(lnoRow)(lnoCol), Trim(vrtRow(lntCol)), Delimiter, False)
                                        lnoCol = lnoCol + 1
                                    End If
                                Next lntCol
                                News = False
                                Exit For
                            End If
                        Next lnoRow
                    End If
                    If News Then
                        vroRow = Empty
                        Call VarSet(vroRow, vrtRow(DomCol))
                        For lntCol = vL(vrtRow) To vU(vrtRow)
                            If lntCol <> DomCol Then
                                Call VarSet(vroRow, vrtRow(lntCol))
                            End If
                        Next lntCol
                        Call VarSet(vroTab, vroRow)
                    End If
                End If
            End If
        Next vrtRow
    End If
    CompressTable = vroTab
End Function



