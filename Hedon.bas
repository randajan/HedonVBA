Attribute VB_Name = "Hedon"

'Require:






Global Const DD = vbCrLf & vbCrLf

#If VBA7 Then
    Declare PtrSafe Global Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Declare Global Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'Main LIBRARY

Public Function IsBlank(ByRef Variable) As Boolean
    IsBlank = IsEmpty(Variable) Or IsMissing(Variable)
End Function




