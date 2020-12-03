Attribute VB_Name = "HedonFormat"

'Require:
    'Hedon.bas
    'HeadonString.bas


Public Function FormatFolderName(ByVal FolderName As String, Optional ByVal Include = "")
    FormatFolderName = STrim(ClearChar(FolderName, "áÁèÈïÏéÉìÌíÍòÒóÓøØšŠúÚùÙýÝžŽABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 _,.-+()!" & Include))
End Function

Public Function FormatQRInvoice(ByVal IBAN As String, ByVal Amount As String, ByVal vS As String, Optional ByVal Ks As String = 308)
    FormatQRInvoice = "SPD*1.0*ACC:" & IBAN & "*AM:" & Amount & "*CC:CZK*MSG:QR*X-KS:" & Ks & "*X-VS:" & vS
End Function

Public Function FormatZIP(ByVal ZIP As String) As String
    Dim dZIP As Double
    dZIP = GetNum(Replace(ZIP, " ", ""))
    If dZIP > 0 Then FormatZIP = SConcatenate(Mid(dZIP, 1, 3), Mid(dZIP, 4, 2), " ")
End Function

Public Function FormatCounter(ByVal Count As Long, ByVal Root As String, Optional ByVal Count1 As String = "", Optional ByVal Count2 As String = "", Optional ByVal Count5 As String = "")
    Dim sAdd As String, lAbsCount As Long
    
    lAbsCount = Abs(Count)
    
    If lAbsCount = 1 Then
        sAdd = Count1
    ElseIf lAbsCount = 2 Or lAbsCount = 3 Or lAbsCount = 4 Then
        sAdd = Count2
    Else
        sAdd = Count5
    End If
    
    FormatCounter = Root & sAdd
End Function
