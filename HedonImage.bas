Attribute VB_Name = "HedonImage"

'Require:
    'Hedon.bas
    'HedonArray.bas
    'HedonFile.bas


Public Function LoadImages(ByRef Target As Range, ByVal Path As String, Optional ByRef Mimes As Variant) As Variant
    Dim Files As Variant
    Dim Img As Object
    Dim Cell As Range
    Dim i As Long
     
    If IsBlank(Mimes) Then Mimes = VarMake("png", "jpg", "bmp", "gif")
    
    Files = FolderContent(Path, False, Mimes)
    If IsBlank(Files) Then Exit Function
    Set Cell = Target.Cells
    
    Application.ScreenUpdating = False
    For i = vL(Files) To vU(Files)
        Set Img = Target.Worksheet.Pictures.Insert(Path & "\" & Files(i)) '
        VarAdd LoadImages, Img
        With Img
            .Name = Files(i)
            .Top = Cell.Top + i * 10
            .Left = Cell.Left + i * 20
            .Width = 100
            .Height = 100
        End With
    Next i
    Application.ScreenUpdating = True

End Function

Public Sub ClearImages(ByRef Target As Range)
    
    Dim Sheet
    Dim xPicRg As Range
    Dim xPic As Picture
    Dim xRg As Range
    
    Set Sheet = Target.Worksheet
    
    Application.ScreenUpdating = False
    For Each xPic In Sheet.Pictures
        Set xPicRg = Sheet.Range(xPic.TopLeftCell.Address & ":" & xPic.BottomRightCell.Address)
        If Not Intersect(Target, xPicRg) Is Nothing Then xPic.Delete
    Next
    Application.ScreenUpdating = True

    
End Sub

Public Sub SaveImages(ByRef Sheet As Worksheet, ByVal Path As String)

    Dim ppt As Object, ps As Variant, slide As Variant
    Dim Img As Shape, Name As String
    Dim cImg As Long

    Set ppt = CreateObject("PowerPoint.application")
    Set ps = ppt.presentations.Add
    Set slide = ps.slides.Add(1, 1)
    
    cImg = 0
    For Each Img In Sheet.Shapes
        Name = Path & "\" & cImg & ".png"
        Img.Copy
        With slide
            .Shapes.Paste
            .Shapes(.Shapes.count).Export Name, 2
            .Shapes(.Shapes.count).Delete
        End With
        cImg = cImg + 1
    Next Img

    With ps
        .Saved = True
        .Close
    End With
    ppt.Quit
    Set ppt = Nothing

End Sub

Public Function GridImages(ByRef Imgs As Variant, ByVal Height As Double, ByVal Width As Double, Optional Top As Double = 0, Optional Left As Double = 0, Optional ByVal Margin As Double = 2)
    Dim Img As Object
    
    Dim vImgs As Variant, vRows As Variant, vRowsWidth As Variant
    Dim cLowRow As Long
    
    Dim cHeight As Double, cWidth As Double
    Dim cRow As Long, cRows As Long, cImg As Long, i As Long

    If Not VarsNotEmpty(Imgs) Then Exit Function
    
    'Sort Images and count widths
    cWidth = 0
    For cImg = vL(Imgs) To vU(Imgs)
        Set Img = Imgs(cImg)
        Img.Height = Height
        cWidth = cWidth + Img.Width
        
        'Sort Images by width
        i = -1
        If VarsNotEmpty(vImgs) Then
            For i = vL(vImgs) To vU(vImgs)
                If i >= 0 And vImgs(i).Width < Img.Width Then Exit For
            Next i
        End If
        VarAdd vImgs, Img, j
    Next cImg
    
    'Create rows
    cRows = Round(Sqr(cWidth / Width))
    If (cRows = 0) Then cRows = 1
    cHeight = Height / cRows
    
    For cRow = 0 To cRows - 1
        VarSet vRowsWidth, 0, cRow
        VaRedim vRows, cRow
    Next cRow
    
    'Fill rows
    For cImg = vL(vImgs) To vU(vImgs)
        Set Img = Imgs(cImg)
        cLowRow = 0
        For cRow = 1 To cRows - 1
            If (vRowsWidth(cRow) < vRowsWidth(cLowRow)) Then cLowRow = cRow
        Next cRow
        VarAdd vRows(cLowRow), Img
        vRowsWidth(cLowRow) = vRowsWidth(cLowRow) + Img.Width
    Next cImg

    'Arrange images in rows
    For cRow = vL(vRows) To vU(vRows)
        RowImages vRows(cRow), cHeight, Width, Top + cRow * cHeight, Left, Margin
    Next cRow
    
    GridImages = Imgs
End Function


Public Function RowImages(ByRef Imgs As Variant, ByVal Height As Double, ByVal Width As Double, Optional Top As Double = 0, Optional Left As Double = 0, Optional ByVal Margin As Double = 2)
    Dim Img As Object
    
    Dim cWidth As Double, cLeft As Double
    Dim dRatio As Double
    Dim cImg As Long

    If Not VarsNotEmpty(Imgs) Then Exit Function
    
    'Count width
    cWidth = 0
    For cImg = vL(Imgs) To vU(Imgs)
        Set Img = Imgs(cImg)
        Img.Height = Height
        cWidth = cWidth + Img.Width
    Next cImg
    
    'Resize ratio
    dRatio = Width / cWidth
    

    cLeft = Left
    If (dRatio > 1) Then cLeft = cLeft + (Width - cWidth) / 2
    
    For cImg = vL(Imgs) To vU(Imgs)
        Set Img = Imgs(cImg)
        Img.Width = Img.Width * dRatio - 2 * Margin
        If (Img.Height > Height - 2 * Margin) Then Img.Height = Height - 2 * Margin
        Img.Top = Top + (Height - Img.Height) / 2
        Img.Left = cLeft + Margin
        cLeft = Img.Left + Img.Width + Margin
    Next cImg
    
    RowImages = Imgs
End Function


