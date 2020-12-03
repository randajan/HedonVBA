Attribute VB_Name = "HedonString"

'Require:
    'Hedon.bas
    'HedonArray.bas



Public Function Compare(ByVal String1 As String, ByVal String2 As String, Optional ByVal Execute As String, Optional ByVal Exact As Boolean = False) As Double
    Compare = (PartCompare(String1, String2, Execute, Exact) + PartCompare(String2, String1, Execute, Exact)) / 2
End Function

Public Function PartCompare(ByVal String1 As String, ByVal String2 As String, Optional ByVal Execute As String, Optional ByVal Exact As Boolean = False) As Double

    Dim Str1Len As Long, Str2Len As Long, ItrLen As Long, i As Long
    Dim Str1 As String, str2 As String, sPart As String
    
    Str1 = String1
    str2 = String2
    
    For i = 1 To Len(Execute)
        Str1 = Replace(Str1, Mid(Execute, i, 1), "")
        str2 = Replace(str2, Mid(Execute, i, 1), "")
    Next i
    
    Str1Len = Len(Str1)
    Str2Len = Len(str2)
    ItrLen = Str1Len
    
    If (Str1Len > 0) And (Str2Len > 0) Then
        While ItrLen > 0
            i = 1
            While (ItrLen + i - 1) <= Str1Len
                sPart = Mid(Str1, i, ItrLen)
                If ((InStr(1, str2, sPart, vbBinaryCompare) > 0) And Exact) Or ((InStr(1, str2, sPart, vbTextCompare) > 0) And Not Exact) Then
                    str2 = Replace(str2, sPart, "", 1, 1, vbTextCompare)
                    PartCompare = PartCompare + ParabolicPercentil((Len(sPart) / Str2Len) * 100)
                End If
                i = i + 1
            Wend
            ItrLen = ItrLen - 1
        Wend
    End If
    PartCompare = ParabolicPercentil(PartCompare) / 100
End Function

Public Function SInStr(ByVal Text As String, Optional ByVal InText As String, Optional ByVal ExText As String) As Boolean
    'Vrátí true pokud v textu èást InText je a ExText není
    SInStr = ((Len(InText) = 0) Or (InStr(1, Text, InText, vbTextCompare) > 0)) And ((Len(ExText) = 0) Or (InStr(1, Text, ExText, vbTextCompare) = 0))
End Function

Public Function BoltoLong(ByVal Bol0 As Boolean, Optional ByVal Bol1 As Boolean = False, Optional ByVal Bol2 As Boolean = False, Optional ByVal Bol3 As Boolean = False, Optional ByVal Bol4 As Boolean = False, Optional ByVal Bol5 As Boolean = False, Optional ByVal Bol6 As Boolean = False, Optional ByVal Bol7 As Boolean = False, Optional ByVal Bol8 As Boolean = False, Optional ByVal Bol9 As Boolean = False) As Long
    If Bol0 Then BoltoLong = 1
    If Bol1 Then BoltoLong = BoltoLong + 1
    If Bol2 Then BoltoLong = BoltoLong + 1
    If Bol3 Then BoltoLong = BoltoLong + 1
    If Bol4 Then BoltoLong = BoltoLong + 1
    If Bol5 Then BoltoLong = BoltoLong + 1
    If Bol6 Then BoltoLong = BoltoLong + 1
    If Bol7 Then BoltoLong = BoltoLong + 1
    If Bol8 Then BoltoLong = BoltoLong + 1
    If Bol9 Then BoltoLong = BoltoLong + 1
End Function

Public Function Pack(Value As String, Optional ByVal Sign As Long = 34) As String
    Pack = Chr(Sign) & Value & Chr(Sign)
End Function

Public Function IsPack(Text As String, Pack As String, Value As String, Sign As Long) As Boolean
    IsPack = InStr(1, Text, Pack & "=" & Chr(Sign) & Value & Chr(Sign)) > 0
End Function

Public Function GetPack(Text As String, Name As String, Optional Sign As Long = 34)
    GetPack = GetString(Text, Name & Chr(Sign), Chr(Sign))
End Function

Public Function HidePass(ByVal Text As String, Optional ByVal PasswordChar As String) As String
    Dim i As Long
    If Len(PasswordChar) > 0 Then
        For i = 1 To Len(Text)
            HidePass = HidePass & PasswordChar
        Next i
    Else
        HidePass = Text
    End If
End Function

Public Function STrim(ByVal Text As String) As String
    Dim i As Long, c As String
    Text = Trim(Text)
    
    Do While i < Len(Text)
        i = i + 1
        c = Mid(Text, i, 1)
        If c <> vbCrLf And c <> vbNewLine And c <> vbCr And c <> vbLf And c <> vbTab Then STrim = STrim & c
    Loop
End Function

Public Function SMid(ByVal Text As String, ByVal Start As Long, ByVal Final As Long) As String
    Dim Top As Long
    Top = Len(Text)
    If Start <= 0 Then Start = 1
    If Final > Top Then Final = Top
    If ((Start <= Top) And (Final > 0)) Then
        If Start <= Final Then SMid = Mid(Text, Start, Final - Start + 1) Else SMid = Mid(Text, Final, Start - Final + 1)
    End If
End Function

Public Function PartStr(ByVal Text As String, ByVal Lenght As Long) As String
    If Len(Text) > Lenght Then PartStr = Mid(Text, 1, Lenght) & "..." Else PartStr = Text
End Function

Public Function StrRepeat(ByVal Str As String, ByVal Repeat As Long) As String
    Dim i As Long, r As String
    For i = 1 To Repeat
        StrRepeat = StrRepeat & Str
    Next i
End Function

Public Function Inflect(ByVal Root As String, ByVal Num As Long, ByVal Ads As Long, Optional WithNum As Boolean = True) As String
    'Skloòuje podle èislovky a vzoru dle èísla Ads (nutno pøedem definovat)
    Static AdsTbl As Variant
    If IsBlank(AdsTbl) Then
        'Vzory první je Ads = 0
        'Vzor Minuta (0)
        Call VarAdd(AdsTbl, StrToList(";a;y;", ";"))
        'Vzor Mail (1)
        Call VarAdd(AdsTbl, StrToList("ù;;y;", ";"))
        'Vzor Den (2)
        Call VarAdd(AdsTbl, StrToList("ní;en;ny;", ";"))
        'Vzor Obesláno (3)
        Call VarAdd(AdsTbl, StrToList("o;;y;", ";"))
    End If
    If WithNum Then Inflect = Num & " "
    Inflect = Inflect & Root
    If Num > 4 Then Num = 0
    If TablesNotEmpty(AdsTbl, Ads) Then Inflect = Inflect & VeVar(Num, AdsTbl(Ads), -3)
End Function

Public Function GetBracketValue(ByVal Text As String, Optional ByVal LBracket = "(", Optional ByVal RBracket = ")") As String
    Dim fSt As Long, lSt As Long, lC As Long, p As Long
    
    If Len(Text) = 0 Then Exit Function
    If Len(LBrackets) = 0 Then LBracket = "("
    If Len(RBrackets) = 0 Then LBracket = ")"

    For p = 1 To Len(Text)
        If Mid(Text, p, Len(LBracket)) = LBracket Then
            If lC = 0 Then fSt = p + 1
            lC = lC + 1
        ElseIf Mid(Text, p, Len(RBracket)) = RBracket Then
            lSt = p
            If lC = 1 Then
                GetBracketValue = Mid(Text, fSt, p - fSt)
                Exit Function
            End If
            lC = lC - 1
        End If
    Next p
    
    If lSt > fSt Then GetBracketValue = Mid(Text, fSt, lSt - fSt)

End Function

Public Function GetString(ByVal Text As String, ByVal Start As String, Optional ByVal Final As String = "")
    
    Dim MidStart As Long
    Dim MidFinal As Long
    
    GetString = ""
    MidStart = InStr(1, Text, Start)
    
    If MidStart <= 0 Then Exit Function
    
    MidStart = MidStart + Len(Start)
    
    If (Final = "") Then
        MidFinal = Len(Text) + 1
    Else
        MidFinal = InStr(MidStart, Text, Final)
    End If
    
    If MidFinal > 0 Then GetString = Mid(Text, MidStart, MidFinal - MidStart)
End Function

Public Function TextFunctionAscii(ByVal Text As String) As String
    Dim Znak As Long
    Znak = CLng(GetNum(GetString(Text, "Ascii(", ")")))
    If ((Znak > 0) And (Znak < 128)) Then TextFunctionAscii = Chr(Znak) Else TextFunctionAscii = Text
End Function

Public Function SConcatenate(ByVal Text1 As String, ByVal Text2 As String, Optional ByVal Delimiter As String, Optional ByVal Silly As Boolean = True)
    'Spojí dva textové øetìzce a oddìlí je Delimiterem
    Text1 = Trim(Text1)
    Text2 = Trim(Text2)
    If Len(Delimiter) = 0 Then Delimiter = ", "
    
    If Not Silly Then Silly = (InStr(1, Text1, Text2, vbBinaryCompare) <= 0)
    
    If Silly Then
        If ((Len(Text1) > 0) And (Len(Text2) > 0)) Then Text1 = Text1 & Delimiter
        SConcatenate = Text1 & Text2
    Else
        SConcatenate = Text1
    End If
End Function

Public Function ConcatIfs(ByVal Delimiter As String, ByRef Concat As Range, ByRef Criterium_Range1 As Range, ByVal Criterium_Value1 As String, Optional ByVal Uniq As Boolean = False) As String

    Dim Cell As Range, crit1 As Range
    
    Dim Finish As Long
    Dim Counter As Long
    
    Finish = Concat.End(xlDown).Row
    
    For Counter = 1 To Finish
        
       Set Cell = Concat.Cells(Counter, 1)
       Set crit1 = Criterium_Range1.Cells(Counter, 1)
       If (crit1.Value <> Criterium_Value1) Then GoTo Continue
       
       ConcatIfs = SConcatenate(ConcatIfs, Cell.Value, Delimiter, Not Uniq)
       
Continue:
    Next Counter
    
End Function

Public Function SReplace(ByVal Expression As String, ByVal Find As String, ByVal Replace As String, Optional Start As Long = 1, Optional count As Long = -1, Optional WholeWord As Boolean = False, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare)
    'K funkci Replace pøidává pøepínaè "WholeWord", které vyhledává jen absolutní slova a èísla
    Dim Final As Long
    
    If Len(Find) > 0 Then
        If Start <= 0 Then Start = 1
        Final = 1
        Do
            Start = InStr(Start, Expression, Find, Compare)
            If Start = 0 Then Exit Do
            If Not (WholeWord And (IsChar(Expression, Start - 1) Or IsChar(Expression, Start + Len(Find)))) Then
                SReplace = SReplace & Mid(Expression, Final, Start - Final) & Replace
                Final = Start + Len(Find)
            End If
            Start = Start + Len(Find)
            count = count - 1
        Loop Until count = 0
        SReplace = SReplace & Mid(Expression, Final, Len(Expression) - Final + 1)
    End If
End Function

Public Function SReplaceList(ByVal Text As String, ByVal tbl As Variant, ByVal SearchCol As Long, Optional ByVal ReplaceCol As Long = -1, Optional WholeWord As Boolean = False)
    Dim i As Long, s As String, r As String
    
    SReplaceList = Text
    If Len(SReplaceList) > 0 Then
        For i = vL(tbl) To vU(tbl)
            If TablesNotEmpty(tbl, i, SearchCol) Then
                s = CStr(tbl(i)(SearchCol))
                If TablesNotEmpty(tbl, i, ReplaceCol) And ReplaceCol >= 0 Then r = CStr(tbl(i)(ReplaceCol)) Else r = ""
                SReplaceList = SReplace(SReplaceList, s, r, 1, -1, WholeWord)
            End If
        Next i
    End If
End Function

Public Function ConTable(ByVal str As String, ByVal Table As Variant, Optional ByVal Default As String, Optional ByVal Row As Long = 2) As String
    'Vrátí øádek Row od nejpodobnìjší vstupní hodnoty v seznamu hodnot Table, nebo defaultní nastavení Default
    Dim Col As Long, Cmpr As Double, TopNum As Double
    ConTable = Default
    Row = MathFrame(Row, 2)
    TopNum = 0.5
    If ((IsBlank(Table)) Or (Len(str) = 0)) Then Exit Function
    If VarsNotEmpty(Table(vL(Table)), vL(Table(vL(Table))) + Row - 1) Then
        For Col = vL(Table(vL(Table))) To vU(Table(vL(Table)))
            Call VaRedim(Table(vL(Table) + Row - 1), vU(Table(vL(Table))))
            Cmpr = Compare(ClearDiak(CStr(Table(vL(Table))(Col))), ClearDiak(str), "_,.-/+* ")
            If Cmpr > TopNum Then
                TopNum = Cmpr
                ConTable = CStr(Table(vL(Table) + Row - 1)(Col))
            End If
        Next Col
    End If
End Function

Public Function IsCharIP(From As String, Num As Long)
    If ((Num > 0) And (Num <= (Len(From)))) Then
        If (((Asc(Mid(From, Num, 1))) >= 48) And ((Asc(Mid(From, Num, 1))) <= 57)) Then
            IsCharIP = 2
        ElseIf (Asc(Mid(From, Num, 1))) = 46 Then
            IsCharIP = 1
        Else
            IsCharIP = 0
        End If
    Else
        IsCharIP = 0
    End If
End Function

Public Function IsCharEm(From, Num)
    If ((Num > 0) And (Num <= (Len(From)))) Then
        If (((Asc(Mid(From, Num, 1))) >= 65) And ((Asc(Mid(From, Num, 1))) <= 90)) Or (((Asc(Mid(From, Num, 1))) >= 97) And ((Asc(Mid(From, Num, 1))) <= 122)) Then
            IsCharEm = 5
        ElseIf (((Asc(Mid(From, Num, 1))) >= 48) And ((Asc(Mid(From, Num, 1))) <= 57)) Then
            IsCharEm = 4
        ElseIf (Asc(Mid(From, Num, 1))) = 45 Then
            IsCharEm = 3
        ElseIf (Asc(Mid(From, Num, 1))) = 46 Then
            IsCharEm = 2
        ElseIf (Asc(Mid(From, Num, 1))) = 95 Then
            IsCharEm = 1
        Else
            IsCharEm = 0
        End If
    Else
        IsCharEm = 0
    End If
End Function

Public Function GetEmail(ByVal Text As String) As Variant
   
    Dim strZavLoc As Integer, strAr As Integer, strLastAr As Integer, strStart As Integer, strEnd As Integer, strNum As Integer
    Dim varMail As Variant
    
    If Len(Text) > 0 Then
        strEnd = 0
        strZavLoc = InStr(1, Text, "@", 1)
        While ((strZavLoc > 1) And (strZavLoc < ((Len(Text)) - 3)))
            strLastChar = 3
            strAr = (strZavLoc - 1)
            While (((IsCharEm(Text, strAr)) >= 4) Or (strLastAr >= 4)) And ((IsCharEm(Text, strAr)) > 0) And (strAr >= strEnd)
                strLastAr = IsCharEm(Text, strAr)
                strAr = strAr - 1
            Wend
            strStart = strAr + 1
            If IsCharEm(Text, strStart) < 4 Then
             strStart = strStart + 1
            End If
            strAr = (strZavLoc + 1)
            strLastAr = 3
            While (((IsCharEm(Text, strAr)) >= 4) Or (strLastAr >= 4)) And ((IsCharEm(Text, strAr)) > 2)
                strLastAr = IsCharEm(Text, strAr)
                strAr = strAr + 1
            Wend
            If ((IsCharEm(Text, strAr) = 2) And ((strAr - strZavLoc) > 2) And (strLastAr >= 4)) Then
                strAr = strAr + 1
                strLastAr = strAr
                While (IsCharEm(Text, strAr)) = 5
                    strAr = strAr + 1
                Wend
                If (strStart < strZavLoc) And (InStr(1, (Mid(Text, strStart, strAr - strStart)), "@", 1) <> 0) Then
                    Call VarPlace(varMail, Mid(Text, strStart, strAr - strStart))
                    strEnd = strAr
                End If
            End If
            strZavLoc = InStr((strZavLoc + 1), Text, "@", 1)
        Wend
    End If
    
    GetEmail = varMail
    
End Function

Public Function StEn(Str1 As String) As String
    StEn = Str1 & vbCrLf
End Function

Function ClearDiak(ByVal str As String) As String
    'Nahradí znaky diakritiky
    Const cz As String = "áÁèÈïÏéÉìÌíÍòÒóÓøØšŠúÚùÙýÝžŽ"
    Const EN As String = "aAcCdDeEeEiInNoOrRsStTuUuUyYzZ"
    Dim i As Long
    If Len(str) = 0 Then Exit Function
    For i = 1 To Len(cz)
        str = Replace(str, Mid(cz, i, 1), Mid(EN, i, 1))
    Next i
    ClearDiak = str
End Function

Public Function ClearChar(ByVal Text As String, Optional ByVal Keep As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789") As String
    Dim i As Long, Char As String
    If Len(Text) = 0 Then Exit Function
    For i = 1 To Len(Text)
        Char = Mid(Text, i, 1)
        If InStr(1, Keep, Char) > 0 Then ClearChar = ClearChar & Char
    Next i
End Function

Public Function IsChar(ByVal Text As String, ByVal AsNum As Long, Optional ByVal Keep As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789") As Boolean
    If Len(Text) > 0 And AsNum > 0 And AsNum <= Len(Text) Then IsChar = InStr(1, Keep, Mid(Text, AsNum, 1)) > 0 Else IsChar = False
End Function

Public Function GetIP(ByVal Text As String) As Variant
    
    Dim IP As String, VarIP As Variant
    Dim Sec As Integer, Oktet As Integer, Aktloc As Integer
    Dim i As Long, count As Long, t As Long
    Dim Default As Boolean, newIp As Boolean
    
    Default = True
    
    For i = 1 To Len(Text) + 1
        If Default Then
            IP = ""
            Sec = 0
            Oktet = 0
            newIp = False
            Default = False
        End If
        If i <= Len(Text) Then Aktloc = IsCharIP(Text, i) Else Aktloc = 0
        If Aktloc = 2 Then
            If Sec < 3 Then
                IP = IP & Mid(Text, i, 1)
                Sec = Sec + 1
            ElseIf ((Sec >= 3) And (Oktet = 0)) Then
                IP = Mid(IP, 2, Len(IP) - 1) & Mid(Text, i, 1)
            Else: Default = True
            End If
        ElseIf ((Aktloc = 1) And (Sec > 0)) Then
            IP = IP & "."
            Sec = 0
            Oktet = Oktet + 1
        ElseIf ((Aktloc = 0) And (Oktet = 3) And (Sec > 0)) Then
            newIp = True
            Default = True
        Else: Default = True
        End If
        If (newIp) Or ((Oktet = 3) And (Sec = 3)) Then
            Call VarPlace(VarIP, IP)
            Default = True
        End If
    Next
    
    GetIP = VarIP
    
End Function



Public Function GetDate(Str1 As String)
    Dim i As Long, Znak As Long
    Dim Serial As String, Last As String, Value As Long
    Dim Rok As Long, Mesic As Long, Den As Long, Hodina As Long, Minuta As Long
        
    For i = 1 To Len(Str1) + 1
        If i <= Len(Str1) Then Znak = Asc(Mid(Str1, i, 1)) Else Znak = 0
        If ((Znak >= 48) And (Znak <= 57)) Then
            If Len(Serial) = 4 Then Serial = Mid(Serial, 2, 3) & Chr(Znak) Else Serial = Serial & Chr(Znak)
            If Serial = "" Then Value = 0 Else Value = CLng(Serial)
        ElseIf (((Last = "Den") Or (Last = "Rok")) And (Value <= 12) And (Mesic = 0)) Then
            Mesic = Value
            Last = "Mesic"
            Serial = ""
            Value = 0
        ElseIf (((Znak >= 45) Or (Znak <= 47)) And (Len(Serial) <= 2) And (Value <= 31) And (Den = 0)) Then
            Den = Value
            Last = "Den"
            Serial = ""
            Value = 0
        ElseIf ((Znak = 58) And (Len(Serial) <= 2) And (Value < 24) And (Hodina = 0)) Then
            Hodina = Value
            Last = "Hodina"
            Serial = ""
            Value = 0
        ElseIf ((Last = "Hodina") And (Len(Serial) = 2) And (Value < 60) And (Minuta = 0)) Then
            Minuta = Value
            Last = "Minuta"
            Serial = ""
            Value = 0
        ElseIf ((Last <> "Hodina") And (Value > 1) And (Rok = 0)) Then
            Rok = Value
            Last = "Rok"
            Serial = ""
            Value = 0
        End If
    Next i
    
    If Rok = 0 Then
        Rok = Year(Now)
        If ((Month(Now) > 6) And (Mesic >= 1) And (Mesic <= 6)) Then Rok = Rok + 1
    ElseIf Rok < 99 Then
        Rok = Rok + (Year(Now) \ 100) * 100
    End If
    If Mesic = 0 Then Mesic = Month(Now)
    If Den = 0 Then Den = Day(Now)
    
    GetDate = (DateSerial(Rok, Mesic, Den) + TimeSerial(Hodina, Minuta, 0))
    
End Function




Public Function TimeMod(ByVal Time As Double, Optional ByVal mSec As Long = 3600, Optional ByVal mMin As Long = 360, Optional ByVal mHour As Long = 24, Optional ByVal mDay As Long = -1, Optional ByVal mWeek As Long = 0) As String
    'Pøevede excelové èíslo oznaèující èas do èitelné textové podoby
    Dim Sec As Long, Min As Long, Hour As Long, Day As Long, Week As Long
    
    If Time > 0 Then Time = Int(Time * 24 * 60 * 60) Else Exit Function
    If mSec <> 0 Then
        If Time < mSec Or mSec = -1 Then Sec = Time Else Sec = Time Mod 60
    End If
    Time = Int(Time - Sec) / 60
    If Time > 0 And mMin <> 0 Then
        If Time < mMin Or mMin = -1 Then Min = Time Else Min = Time Mod 60
    End If
    Time = Int(Time - Min) / 60
    If Time > 0 And mHour <> 0 Then
        If Time < mHour Or mHour = -1 Then Hour = Time Else Hour = Time Mod 24
    End If
    Time = Int(Time - Hour) / 24
    If Time > 0 And mDay <> 0 Then
        If Time < mDay Or mDay = -1 Then Day = Time Else Day = Time Mod 7
    End If
    If Time > 0 Then Week = Int(Time - Day) / 7
    
    
    If Week > 0 Then TimeMod = SConcatenate(TimeMod, Inflect("týd", Week, 2), ", ")
    If Day > 0 Then TimeMod = SConcatenate(TimeMod, Inflect("d", Day, 2), ", ")
    If Hour > 0 Then TimeMod = SConcatenate(TimeMod, Inflect("hodin", Hour, 0), ", ")
    If Min > 0 Then TimeMod = SConcatenate(TimeMod, Inflect("minut", Min, 0), ", ")
    If Sec > 0 Then TimeMod = SConcatenate(TimeMod, Inflect("vteøin", Sec, 0), ", ")

End Function

Function NumToLetter(ByVal Number As Long) As String
    Dim c As Byte, s As String
    Do
        c = ((Number - 1) Mod 26)
        s = Chr(c + 65) & s
        Number = (Number - c) \ 26
    Loop While Number > 0
    NumToLetter = s
End Function

Public Function GetNum(ByVal Strin As String, Optional ByVal Default As Double) As Double
    Dim Znak As String, GVar As String, i As Long, Strife As Boolean, Found As Boolean
    Strife = False
    Found = False
    For i = 1 To Len(Strin)
        Znak = Asc(Mid(Strin, i, 1))
        If ((Znak >= 48) And (Znak <= 57)) Then
            GVar = GVar & Chr(Znak)
            Found = True
        ElseIf (((Znak = 44) Or (Znak = 46)) And Found And Not Strife) Then
            GVar = GVar & ","
            Strife = True
        ElseIf Found Then
            Exit For
        End If
    Next i
    If GVar = "" Then GetNum = Default Else GetNum = CDbl(GVar)
End Function

