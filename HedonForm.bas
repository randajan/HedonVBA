Attribute VB_Name = "HedonForm"

'Require:
    'Hedon.bas


Option Explicit
'API functions
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" _
                                       Alias "SetWindowLongA" _
                                       (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" _
                                      (ByVal hwnd As Long, _
                                       ByVal hWndInsertAfter As Long, _
                                       ByVal X As Long, _
                                       ByVal y As Long, _
                                       ByVal cx As Long, _
                                       ByVal cy As Long, _
                                       ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" _
                                    Alias "FindWindowA" _
                                    (ByVal lpClassName As String, _
                                     ByVal lpWindowName As String) As Long
Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32" _
                                     Alias "SendMessageA" _
                                     (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
                                     
'Constants
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const GWL_EXSTYLE = (-20)
Private Const HWND_TOP = 0
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_SHOWWINDOW = &H40
Private Const MF_BYPOSITION = &H400&
Private Const WS_SYSMENU = &H80000
Private Const WS_EX_APPWINDOW = &H40000
Private Const GWL_STYLE = (-16)
Private Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_FULLSIZING = &H70000
Private Const SWP_FRAMECHANGED = &H20
Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0&
Private Const ICON_BIG = 1&

Public Sub SForm(Form As Object, Optional Sizing As Boolean = False, Optional Min As Boolean = True, Optional XBut As Boolean = True, Optional ByVal Icon As Long)
   Dim hwnd As Long
   hwnd = FindWindow(vbNullString, Form.Caption)
   Call SetButton(hwnd, Sizing, Min, XBut)
   Call AppTasklist(hwnd)
   If Icon > 0 Then Call AddIcon(hwnd, Icon)
End Sub

Private Sub AddIcon(hwnd As Long, hIcon As Long)
'Add an icon on the titlebar
    Dim lngRet As Long
    lngRet = SendMessage(hwnd, WM_SETICON, ICON_SMALL, ByVal hIcon)
    lngRet = SendMessage(hwnd, WM_SETICON, ICON_BIG, ByVal hIcon)
    lngRet = DrawMenuBar(hwnd)
End Sub

Public Sub SetButton(ByVal hwnd As Long, Optional Sizing As Boolean = False, Optional Min As Boolean = True, Optional XBut As Boolean = True)
    If Min Then SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_MINIMIZEBOX
    If Sizing Then
        SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_MAXIMIZEBOX
        SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_FULLSIZING
    End If
    If Not XBut Then RemoveMenu GetSystemMenu(hwnd, 0), 6, MF_BYPOSITION
    If Not (Min Or Sizing Or XBut) Then SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) And Not WS_SYSMENU
End Sub

Private Sub AppTasklist(hwnd As Long)
'Add this userform into the Task bar
    Dim WStyle As Long
    Dim Result As Long
    WStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    WStyle = WStyle Or WS_EX_APPWINDOW
    Result = SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, _
                          SWP_NOMOVE Or _
                          SWP_NOSIZE Or _
                          SWP_NOACTIVATE Or _
                          SWP_HIDEWINDOW)
    Result = SetWindowLong(hwnd, GWL_EXSTYLE, WStyle)
    Result = SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, _
                          SWP_NOMOVE Or _
                          SWP_NOSIZE Or _
                          SWP_NOACTIVATE Or _
                          SWP_SHOWWINDOW)
End Sub

Public Sub CopyListbox(ByRef ListBox As Object, Optional ByVal RowDelimiter As String, Optional ByVal ColDelimiter As String)
    'Kopíruje vybrané prvky seznamu Listbox do Clipboardu
    Dim MyData As New DataObject
    Set MyData = New DataObject
    
    On Error Resume Next
    MyData.Clear
    MyData.SetText CStr(TableToStr(GetListBox(ListBox, JustSelected:=True), RowDelimiter, ColDelimiter))
    MyData.PutInClipboard
    
End Sub

Public Sub PasteListbox(ByRef ListBox As Object, Optional ByVal RowDelimiter As String, Optional ByVal ColDelimiter As String, Optional ByVal ColEmbed As Boolean = False, Optional ByVal ColSensitivity As Integer)
    'Vkládá Clipboardu do Listboxu
    Dim MyData As New DataObject
    Set MyData = New DataObject
    
    MyData.GetFromClipboard
    On Error Resume Next
    If Not IsBlank(MyData.GetText) Then Call SetListBox(ListBox, StrToTable(CStr(MyData.GetText), RowDelimiter, ColDelimiter, ColEmbed, ColSensitivity))
    
End Sub

Public Sub DelListbox(ByRef ListBox As Object, Optional ByVal DelAll = False)
    'Maže vybrané nebo všechny prvky seznamu Listbox
    Dim varRow As Variant, varCol As Variant, Del As Boolean
    Dim i As Integer, c As Integer
    
    Del = False
    ReDim varCol(0 To ListBox.ColumnCount - 1)
    For i = 0 To ListBox.ListCount - 1
        If ListBox.Selected(i) Or DelAll Then
            Del = True
        Else
            For c = 0 To ListBox.ColumnCount - 1
                varCol(c) = ListBox.List(i, c)
            Next c
            Call VarSet(varRow, varCol)
        End If
    Next i
    If Del Then Call SetListBox(ListBox, varRow)

End Sub
Public Sub SelectListBox(ByRef ListBox As Object, Optional ByVal Typ As String = "All")
    Dim i As Long
    For i = 0 To ListBox.ListCount - 1
        ListBox.Selected(i) = ((Not ListBox.Selected(i)) Or (Typ = "All"))
    Next i
End Sub

Public Sub SetContMenu(ByRef ContMenu As Object, ByRef Parent As Object, ByVal RowColTab As Variant, ByVal X As Single, ByVal y As Single, ByVal Visible As Boolean)
    'Nastavuje kontextové Menu po kliknutí myší
    Dim ColTab(0 To 0) As Variant, Row As Long
    
    With ContMenu
        .Tag = Parent.Name
        If ((Not IsBlank(RowColTab)) And Visible) Then
            Call SetListBox(ContMenu, RowColTab, True)
            .Left = Parent.Left + X - 3
            .Top = Parent.Top + y - 2
            .Visible = Visible
        Else
            .Visible = False
        End If
    End With
End Sub

Public Sub ChangePage(ByRef MultiPage As Object, ByVal Page As String)
    'Pøelistuje Multipage na stranu s názvem Page
    MultiPage.Value = CLng(MultiPage.Pages(CStr(Page)).index)
End Sub

Public Sub SWForm(ByRef Switch As Object, ByVal State As Boolean, Optional ByVal Default As Boolean = False)
    'Aktivuje nebo deaktivuje zaškrtávací checkbox
    Switch.Enabled = State
    If Not State Then Switch.Value = Default
End Sub

Public Sub SetListBox(ListBox As Object, RowColTab As Variant, Optional AutoSize As Boolean = False)
    'Vpisuje hodnoty seøazené v maticovém seznamu "RowColTab" do ListBoxu. Volba autosize mìní celkovou velikost listboxu na míru všech hodnot
    Dim WCol As String, WholeWidth As Double, RowTab As Variant, ColTab, Row As Long, Col As Long, i As Long
    Dim bfr As Worksheet, Testr As String
    Static ListBoxBufferTab As Variant

    On Error Resume Next
    
    With ListBox
        If IsBlank(ListBoxBufferTab) Then
            ReDim ListBoxBufferTab(0 To 1)
        End If
        
        For i = vL(ListBoxBufferTab) To vU(ListBoxBufferTab) + 1
            If i <= vU(ListBoxBufferTab) Then
                If ListBoxBufferTab(i) = CStr(.Name) Then
                    Exit For
                End If
            Else
                ReDim Preserve ListBoxBufferTab(0 To i + 1)
                ListBoxBufferTab(i) = CStr(.Name)
                ListBoxBufferTab(i + 1) = "ListBoxBuffer" & i / 2
                Exit For
            End If
        Next i
        
        Set bfr = Application.ThisWorkbook.Sheets(SetSheet(CStr(ListBoxBufferTab(i + 1)), True, False))
    
        .Enabled = False
        .ListIndex = -1
        .RowSource = ""
        .ColumnCount = 1
        If TablesNotEmpty(RowColTab) Then
            For Each RowTab In RowColTab
                Col = 0
                Row = Row + 1
                For Each ColTab In RowTab
                    Col = Col + 1
                    bfr.Cells(Row, Col) = ColTab
                Next ColTab
                .ColumnCount = CLng(MaxNum(.ColumnCount, Col))
            Next RowTab
            
            With bfr.Cells
                With .Font
                    If ListBox.Font.Bold And ListBox.Font.Italic Then
                        .FontStyle = "Tuènì kurzíva"
                    ElseIf ListBox.Font.Bold Then
                        .FontStyle = "Tuènì"
                    ElseIf ListBox.Font.Italic Then
                        .FontStyle = "Kurzíva"
                    Else
                        .FontStyle = "Obyèejnì"
                    End If
                    .Name = ListBox.Font.Name
                    .Size = ListBox.Font.Size
                    .Strikethrough = ListBox.Font.Strikethrough
                    .Underline = ListBox.Font.Underline
                End With
                .EntireColumn.AutoFit
            End With

            For Col = 1 To .ColumnCount
                WholeWidth = WholeWidth + (bfr.Columns(Col).ColumnWidth * 5.1) + 8.9
                WCol = WCol & Replace((bfr.Columns(Col).ColumnWidth * 5.1) + 8.9, ",", ".") & " pt; "
            Next Col
            
            .RowSource = CStr("'[" & Application.ThisWorkbook.Name & "]" & bfr.Name & "'!" & Range(bfr.Cells(1, 1), bfr.Cells(Row, .ColumnCount)).Address)
            .ColumnWidths = WCol
            .Enabled = True
            If AutoSize Then
                .Height = ((bfr.Rows(1).RowHeight - (.Font.Size * 0.06)) * Row) + 3
                .Width = WholeWidth + 8.9
            End If
        Else
            Application.DisplayAlerts = False
            bfr.Delete
            Application.DisplayAlerts = True
            .Enabled = True
        End If
    End With
End Sub

Public Sub DeEmbedList(ByRef strList As Variant, Optional EmbChar As String, Optional Sensitivity As Integer)
    Dim Row As Long, RowCount As Long, i As Long, y As Long, strRow As String, Temb As Boolean, dVar As Variant, Ix As Double, ColVar As Variant
    EmbChar = TextFunctionAscii(EmbChar, " ")
    If ((Sensitivity <= 0) Or (Sensitivity > 100) Or (IsBlank(Sensitivity))) Then Sensitivity = 95
    If Not IsBlank(strList) Then
        Call VarSet(dVar, 1)
        i = 2
        Do
            Ix = 0
            RowCount = 0
            For Row = vL(strList) To vU(strList)
                strRow = CStr(strList(Row))
                If Len(strRow) >= i Then
                    RowCount = RowCount + 1
                    Temb = Mid(strRow, i, 1) = EmbChar
                    If Mid(strRow, i - 1, 1) = EmbChar Then
                        If Temb Then Ix = Ix + 75 Else Ix = Ix + 100
                    Else
                        If Temb Then Ix = Ix + 50 Else Ix = Ix + 25
                    End If
                End If
            Next Row
            If RowCount > 0 Then
                Ix = (Ix / RowCount)
                If Ix > Sensitivity Then Call VarSet(dVar, i)
            End If
            i = i + 1
        Loop Until RowCount = 0
        Call VarSet(dVar, i)
        For Row = vL(strList) To vU(strList)
            ColVar = Empty
            For i = vL(dVar) + 1 To vU(dVar)
                strRow = Trim(SMid(strList(Row), dVar(i - 1), dVar(i) - 1))
                y = Len(strRow)
                Do While y > 0
                    Temb = Mid(strRow, y, 1) = EmbChar
                    If Not Temb Then Exit Do
                    y = y - 1
                Loop
                Call VarSet(ColVar, Trim(SMid(strRow, 1, y)))
            Next i
            strList(Row) = ColVar
        Next Row
    End If
End Sub

Public Function GetListBox(ByRef ListBox As Object, Optional ByVal MinRow As Long = -1, Optional ByVal MaxRow As Long = -1, Optional ByVal minCol As Long = -1, Optional ByVal maxCol As Long = -1, Optional ByVal JustSelected As Boolean = False) As Variant
    Dim Row As Long, Col As Long, RowVar As Variant, ColVar As Variant
    With ListBox
        If .ListCount > 0 Then
            If MaxRow < 0 Then MaxRow = .ListCount - 1
            If MinRow < 0 Then MinRow = 0
            MinRow = CLng(MathFrame(MinRow, 0, .ListCount - 1))
            MaxRow = CLng(MathFrame(MaxRow, MinRow, .ListCount - 1))
            For Row = MinRow To MaxRow
                On Error Resume Next
                If Not JustSelected Or .Selected(Row) Then
                    If maxCol < 0 Then maxCol = .ColumnCount - 1
                    If minCol < 0 Then minCol = 0
                    minCol = CLng(MathFrame(minCol, 0, .ColumnCount - 1))
                    maxCol = CLng(MathFrame(maxCol, minCol, .ColumnCount - 1))
                    ColVar = Empty
                    For Col = minCol To maxCol
                        Call VarSet(ColVar, .List(Row, Col))
                    Next Col
                    Call VarSet(RowVar, ColVar)
                End If
            Next Row
            GetListBox = RowVar
        End If
    End With
End Function



