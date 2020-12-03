Attribute VB_Name = "HedonSheet"

'Require:
    'Hedon.bas


Public Function TagSheet(ByRef Target As Range) As String
    TagSheet = Target.Worksheet.Name
End Function
Public Function TagRange(ByRef Target As Range, Optional ByVal Full As Boolean = False) As String
    If Full Then TagRange = "'" & Target.Worksheet.Name & "'!" & Target.Address Else TagRange = Target.Address
End Function

Public Function ByTag(ByRef Target As Range) As Range
    Set ByTag = Range(Target.Cells(1, 1).Value)
End Function

'Save sheet as new file
Public Sub SaveSheet(ByRef Sheet As Worksheet, ByVal Path As String)
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    
    With wb
        Sheet.Copy After:=.Worksheets(.Worksheets.count)
        Application.DisplayAlerts = False
        .Worksheets(1).Delete
        Application.DisplayAlerts = True
        .SaveAs Path
        .Close False
    End With
End Sub

'Get cell in Column by Value
Public Function FindInRow(ByRef Row As Range, ByVal Value As String) As Range

    Dim i As Long
    
    For i = 1 To Row.Cells(1, Row.Columns.count).End(xlToLeft).Column
        If Row.Cells(1, i).Value = Value Then
            Set FindInRow = Row.Cells(1, i)
            Exit Function
        End If
    Next i

End Function

'Get cell in Column by Value
Public Function FindInColumn(ByRef Column As Range, ByVal Value As String) As Range

    Dim i As Long
    
    For i = 1 To Column.Cells(Column.Rows.count, 1).End(xlUp).Row
        If Column.Cells(i, 1).Value = Value Then
            Set FindInColumn = Column.Cells(i, 1)
            Exit Function
        End If
    Next i
    
End Function


Public Function GetValueAround(ByRef Cell As Range, Optional ByVal RowOffset As Long = 0, Optional ByVal ColumnOffset As Long = 0) As String
    GetValueAround = ""
    If Cell Is Nothing Then Exit Function
    GetValueAround = Cell.Offset(RowOffset, ColumnOffset).Value
End Function

Public Function GetValueNextTo(ByRef Cell As Range, Optional ByVal ColumnOffset As Long = 1, Optional ByVal SeekDistance As Long = 0)
    Dim Value As String
    
    SeekDistance = SeekDistance + ColumnOffset
    
    Do While (Value = "" And ColumnOffset <= SeekDistance)
        Value = GetValueAround(Cell, 0, ColumnOffset)
        ColumnOffset = ColumnOffset + 1
    Loop
    
    GetValueNextTo = Value
End Function

Public Function GetValueBellow(ByRef Cell As Range, Optional ByVal RowOffset As Long = 1, Optional ByVal SeekDistance As Long = 0)
    Dim Value As String
    
    SeekDistance = SeekDistance + RowOffset
    
    Do While (Value = "" And RowOffset <= SeekDistance)
        Value = GetValueAround(Cell, RowOffset, 0)
        RowOffset = RowOffset + 1
    Loop
    
    GetValueBellow = Value
End Function

Public Function SetSheet(Name As String, Optional Reset As Boolean = False, Optional ByVal Visible As Boolean = True) As Long
    Dim i As Long
    With Application.ThisWorkbook
        For i = 1 To .Sheets.count
            If .Sheets(i).Name = Name Then
                If Reset Then
                    Application.DisplayAlerts = False
                    .Sheets(i).Delete
                    Application.DisplayAlerts = True
                    Exit For
                Else
                    SetSheet = i
                    .Sheets(i).Visible = Visible
                    Exit Function
                End If
            End If
        Next i
        .Sheets.Add(After:=.Sheets(.Sheets.count)).Name = Name
        .Sheets(Name).Visible = Visible
        SetSheet = .Sheets.count
    End With
End Function

Public Function GetDB(DB As Range) As Range
    Dim r As Long, c As Long, cn As Long
    If DB Is Nothing Then Exit Function
    With DB
        For r = 1 To 100
            cr = r
            For c = 1 To MinNum(.Columns.count, r)
                If Not IsEmpty(.Cells(cr, c)) Then GoTo Found
                cr = cr - 1
            Next c
        Next r
        Exit Function
Found:
        Set GetDB = Range(.Cells(cr, c), .Cells((Range(.Cells(.Rows.count, c), .Cells(.Rows.count, c)).End(xlUp).Row), (Range(.Cells(cr, .Columns.count), .Cells(cr, .Columns.count)).End(xlToLeft).Column)))
    End With
End Function

Public Function SheetToRange(ByRef Sheet As Worksheet) As Range
    With Sheet
        Set SheetToRange = Range(.Cells(1, 1), .Cells(.Rows.count, .Columns.count))
    End With
End Function

Public Function RangeToTable(ByVal Range As Range, Optional ByVal WantRow As Variant = Empty, Optional ByVal WantCol As Variant = Empty) As Variant
    Dim Row As Long, Col As Long, Want
    Dim Add As Boolean
    Dim varRow As Variant, varCol As Variant
    
    If Range Is Nothing Then Exit Function

    With Range
        If ((.Rows.count > 0) And (.Columns.count > 0)) Then
            For Row = 1 To .Rows.count
                Add = True
                If Not IsBlank(WantRow) Then
                    Want = vL(WantRow)
                    Do While VarsNotEmpty(WantRow, Want)
                        If Row = CLng(GetNum(CStr(WantRow(Want)))) Then Exit Do
                        Want = Want + 1
                        If Want > vU(WantRow) Then Add = False
                    Loop
                End If
                If Add Then
                    varCol = Empty
                    For Col = 1 To .Columns.count
                        If Not IsBlank(WantCol) Then
                            Want = vL(WantCol)
                            Do While VarsNotEmpty(WantCol, Want)
                                If Col = CLng(GetNum(CStr(WantCol(Want)))) Then Exit Do
                                Want = Want + 1
                                If Want > vU(WantCol) Then Add = False
                            Loop
                        End If
                        If Add Then Call VarSet(varCol, .Cells(Row, Col).Value)
                    Next Col
                    Call VarSet(varRow, varCol)
                End If
            Next Row
            RangeToTable = varRow
        End If
    End With

End Function

Public Function ZPocet(Vyber As Range) As Long
    'Vrací po?et nestejných položek z databáze
    
    Dim NewVal As Boolean
    Dim Table As Variant
    ReDim Table(0 To 0)
    
    For i = 1 To Vyber.Columns.count
        For y = 1 To Vyber.Rows.count
            NewVal = True
            For X = vL(Table) To vU(Table)
                'zn = MsgBox(Vyber.Cells(y, i) & " =? " & Table(x), vbOKOnly, x)
                If (Vyber.Cells(y, i) = Table(X)) Then
                    NewVal = False
                    Exit For
                End If
            Next X
            If NewVal Then
                ReDim Preserve Table(0 To X)
                Table(X - 1) = Vyber.Cells(y, i)
            End If
        Next y
    Next i
    
    ZPocet = vU(Table)
    
End Function