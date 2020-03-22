Attribute VB_Name = "Hedon"

'Require:






Global Const DD = vbCrLf & vbCrLf

'Main LIBRARY

Public Function IsBlank(ByRef Variable) As Boolean
    IsBlank = IsEmpty(Variable) Or IsMissing(Variable)
End Function




