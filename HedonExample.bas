Attribute VB_Name = "HedonExample"

'Require:
    'Hedon.bas
    'HedonWeb.bas


Public Sub school()
    Dim scl As Worksheet, IE As Object, tbl As Variant
    Dim dRow As Long, r As String

    Const dmn = "https://www.stredniskoly.cz/skola/"
    Set scl = Application.ThisWorkbook.Worksheets("school")

    dRow = 1

    
    Do
        If (scl.Cells(dRow, 6).Value = "") Then
            r = SetIE(dmn & scl.Cells(dRow, 1) & ".html", IE, True)
            scl.Cells(dRow, 6).Value = ListToStr(ParseDoc(IE.Document, "td", "detail-hwww", 2, 1))
            scl.Cells(dRow, 7).Value = ListToStr(ParseDoc(IE.Document, "td", "detail-hmail", 2, 1))
            scl.Cells(dRow, 8).Value = ListToStr(ParseDoc(IE.Document, "td", "detail-hphone", 2, 1))
        End If
        If (scl.Cells(dRow, 6).Value = "") Then
            tbl = ParseDoc(IE.Document, "table", "id=" & Pack("udaje"), 0, 1)
            If VarsNotEmpty(tbl) Then
                scl.Cells(dRow, 8).Value = ListToStr(ParseDoc(tbl(0), "tr", "Telefonn� ��slo:", 2, 1))
                scl.Cells(dRow, 6).Value = ListToStr(ParseDoc(tbl(0), "tr", "Internetov� adresa:", 2, 1))
                scl.Cells(dRow, 7).Value = ListToStr(ParseDoc(tbl(0), "tr", "Emailov� adresa:", 2, 1))
            End If

        End If
        dRow = dRow + 1
    Loop Until scl.Cells(dRow, 1) = ""
End Sub



