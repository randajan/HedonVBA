Attribute VB_Name = "HedonFile"

'Require:
    'Hedon.bas


Public Function IsDir(ByVal Path As String)
    IsDir = Len(Dir(Path, vbDirectory)) > 0
End Function

Public Function CreateDir(ByVal Path As String, Optional ByVal Recurse = False)
    Dim Fragment As Variant
    Dim Parent As String
    
    If Not IsDir(Path) Then
        
        Fragment = StrToList(Path, "\")
        VarDel Fragment, vU(Fragment)
        Parent = ListToStr(Fragment, "\")
        
        If Recurse Then CreateDir Parent, True
        If IsDir(Parent) Then MkDir Path
        
    End If
    
    CreateDir = IsDir(Path)
    
End Function

Function FolderContent(ByVal Path As String, Optional ByVal FullPath As Boolean = False, Optional ByRef Mimes As Variant) As Variant
 
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim Name As String
    
    If Not IsDir(Path) Then Exit Function
     
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(Path)
     
    For Each oFile In oFolder.Files
        Name = oFile.Name
        If FullPath Then Name = Path & "\" & Name
        If (IsBlank(Mimes) Or VarFind(Mimes, GetString(Name, ".", "")) >= 0) Then VarAdd FolderContent, Name
    Next oFile
 
End Function


Public Function TextToFile(ByVal FullPath As String, ByVal Text As String, Optional ByVal NewRow As Boolean = True, Optional Visible As Boolean = True, Optional Reset As Boolean = False) As Boolean
    Dim FSO As Object, oFi As Object, vFi As Object
    
    TextToFile = False
    If Len(FullPath) = 0 Then Exit Function
    Set FSO = CreateObject("Scripting.FileSystemObject")
    With FSO
    
        On Error Resume Next
        If (.FileExists(FullPath) Or Len(Dir(FullPath)) > 0) Then Set vFi = .GetFile(FullPath)
        
        If vFi Is Nothing Or Reset Then
            If Not vFi Is Nothing Then Call vFi.Delete
            Set oFi = .CreateTextFile(FullPath, True)
            oFi.Close
        End If

        Set vFi = .GetFile(FullPath)
        If Visible Then vFi.Attributes = vFi.Attributes And (Not vbHidden) Else vFi.Attributes = vFi.Attributes Or vbHidden
        
        Set oFi = .OpenTextFile(FullPath, 8)
        If NewRow And Not Reset Then oFi.Write vbCrLf
        oFi.Write Text
        oFi.Close

        Set FSO = Nothing
        Set oFi = Nothing
        Set vFi = Nothing
        TextToFile = True
    End With
    
End Function

Public Function FileToText(ByVal FullPath As String) As String
    Dim FSO As Object, oFi As Object, vFi As Object
    
    If Len(FullPath) = 0 Then Exit Function
    Set FSO = CreateObject("Scripting.FileSystemObject")
    With FSO

        On Error Resume Next
        If (.FileExists(FullPath) Or Len(Dir(FullPath)) > 0) Then Set vFi = .GetFile(FullPath)
        
        If Not vFi Is Nothing Then
            Set oFi = .OpenTextFile(FullPath, 1)
            FileToText = oFi.ReadAll
            oFi.Close
        End If

        Set FSO = Nothing
        Set oFi = Nothing
        Set vFi = Nothing
    End With
    
End Function

Public Function FilePath(ByVal Name As String, Optional Mime As String = "txt", Optional Path As String) As String
    If Len(Name) = 0 Then Exit Function
    If Len(Mime) = 0 Then Mime = "txt"
    If Len(Path) = 0 Then Path = Application.ThisWorkbook.Path & "\"
    FilePath = Path & Name & "." & Mime
End Function


