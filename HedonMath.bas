Attribute VB_Name = "HedonMath"

'Require:
    'Hedon.bas


Public Function MathFrame(ByVal Num, Optional ByVal Min = Empty, Optional ByVal Max = Empty)
    If Not IsBlank(Min) Then
        If Num < Min Then Num = Min
    End If
    If Not IsBlank(Max) Then
        If Num > Max Then Num = Max
    End If
    MathFrame = Num
End Function
Public Function PerToVal(ByVal NewPercent As Double, ByVal MinValue As Double, ByVal MaxValue As Double) As Double
    PerToVal = MathFrame(((MaxValue - MinValue) / 100 * NewPercent) + MinValue, MinValue, MaxValue)
End Function

Public Function ValToPer(ByVal NewValue As Double, ByVal MinValue As Double, ByVal MaxValue As Double, Optional ByVal Per As Double = 100) As Double
    MinValue = MathFrame(MinValue, Max:=MaxValue)
    ValToPer = ((Per / (MaxValue + 1 - MinValue)) * NewValue)
End Function

Public Function GetOnePer(ByVal MinValue As Double, ByVal MaxValue As Double, Optional ByVal Per As Double = 100) As Double
    MinValue = MathFrame(MinValue, Max:=MaxValue)
    GetOnePer = Per / (MaxValue + 1 - MinValue)
End Function
Public Function ParabolicPercentil(Percent As Double) As Double
    ParabolicPercentil = (Sin((PerToVal(Percent, 0, 180) - 90) / 57.2957795130823) + 1) * 50
End Function

Public Function MaxNum(ByVal Num1, ByVal Num2)
    If Num1 > Num2 Then MaxNum = Num1 Else MaxNum = Num2
End Function

Public Function MinNum(ByVal Num1, ByVal Num2)
    If Num1 < Num2 Then MinNum = Num1 Else MinNum = Num2
End Function

Public Function Counter(Optional ByVal Tag As String, Optional ByVal Wripe As Boolean = True, Optional ByVal AddCount As Long = 1) As Long
    Dim Reset As Boolean
    Static cList As Variant, cRow As Long, cRtrn As Long
    
    Reset = False
    Do
        If TablesNotEmpty(cList, cRow, 1) Then
            If cList(cRow)(0) = Tag Then Exit Do
        Else
            Call VarSet(cList, VarMake(Tag, 0), cRow)
            Exit Do
        End If
        If Reset Or cRow = 0 Then cRow = cRow + 1 Else cRow = 0
        Reset = True
    Loop
    If Reset Then cRtrn = CLng(cList(cRow)(1))
    cRtrn = cRtrn + AddCount
    If Wripe Then cList(cRow)(1) = cRtrn

    Counter = cRtrn
End Function
