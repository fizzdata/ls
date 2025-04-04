Attribute VB_Name = "Module1"
Dim sRep As String
Sub SwapYellow(Byval lP1 As Long, Byval lP2 As Long)
    If IsYellow(lP1) = True And IsYellow(lP2) = True Then
    Elseif IsYellow(lP1) = True Then
        SetYellow lP2
        ResetYellow lP1
    Elseif IsYellow(lP2) = True Then
        SetYellow lP1
        ResetYellow lP2
    End If
End Sub
Sub SwapPagesModule(Byval lP1 As Long, Byval lP2 As Long)
    Dim r1 As Range, r2 As Range, r3 As Range
    '    Set r1 = ThisWorkbook.Names("Page_" & lP1).RefersToRange    ' prva strana
    '    Set r2 = ThisWorkbook.Names("Page_" & lP2).RefersToRange    ' druga strana

    Set r1 = GetPageRange(lP1)
    Set r2 = GetPageRange(lP2)

    Set r3 = ThisWorkbook.Names("swapplace").RefersToRange    ' temp mesto
    r3.ClearContents
    r3.ClearComments
    r3.UnMerge
    r2.Copy
    r3.Cells(1, 1).PasteSpecial
    r2.ClearContents
    r2.ClearComments
    r2.UnMerge
    r1.Copy
    r2.Cells(1, 1).PasteSpecial
    r1.ClearContents
    r1.ClearComments
    r1.UnMerge
    r3.Copy
    r1.Cells(1, 1).PasteSpecial
    r3.ClearContents
    r3.ClearComments
    r3.UnMerge
    Application.CutCopyMode = False
    'bleed, ako ga ima
    If IsBleed(lP1) = True Or IsBleed(lP2) = True Then
        'zameni bleedove
        SwapBleedsModule lP1, lP2
    End If
    'strane markirane ko kolor
    If IsYellow(lP1) = True Or IsYellow(lP2) = True Then
        'zameni bleedove
        SwapYellow lP1, lP2
    End If
End Sub
Sub SwapBleedsModule(Byval lP1 As Long, Byval lP2 As Long)

    If IsBleed(lP1) = True And IsBleed(lP2) = True Then
    Elseif IsBleed(lP1) = True Then
        SetBleed lP2
        ResetBleed lP1
    Elseif IsBleed(lP2) = True Then
        SetBleed lP1
        ResetBleed lP2
    End If

End Sub
Function IsEmptyPage2(pn As Long) As Boolean
    'gleda i format
    Dim r1 As Range, rc As Range
    Dim b As Boolean
    b = True
    'Set r1 = ThisWorkbook.Names("Page_" & pn).RefersToRange
    Set r1 = GetPageRange(pn)
    For Each rc In r1.Cells
        If Len(rc.value) > 0 Then
            b = False
         Exit For
        End If
        If rc.MergeCells = True Then
            b = False
         Exit For
        End If
        If rc.Interior.Pattern <> xlNone Then
            b = False
         Exit For
        End If
        If rc.Interior.TintAndShade <> 0 Then
            b = False
         Exit For
        End If
    Next rc
    IsEmptyPage2 = b
End Function
Function IsEmptyPage(pn As Long) As Boolean
    Dim r1 As Range, rc As Range
    Dim b As Boolean
    b = True
    'Set r1 = ThisWorkbook.Names("Page_" & pn).RefersToRange
    Set r1 = GetPageRange(pn)
    For Each rc In r1.Cells
        If Len(rc.value) > 0 Then
            b = False
         Exit For
        End If
    Next rc
    IsEmptyPage = b
End Function
Sub MyDeletePages()
    'brise samo Do max pageta
    If IDP Then Exit Sub
        '    If InName = False Then
        '        MsgBox "Select page where you want To start delete"
        '        Exit Sub
        '    End If

        Dim X As Long
        Dim lFilledPage As Long
        Dim lPagesToDelete As Long
        Dim lActivePage As Long
        Dim lPageReserve As Long
        Dim lMaxPage As Long
        Dim lMinPage As Long
        Dim bOdZadnjeNapunjeneStrane As Boolean
        Dim ws As Worksheet
        '    Dim keysWs As Worksheet
        '    Dim matchWs As Worksheet

        'On Error Resume Next
        'bookPageCount = UBound(PagesArr)
        'If Err.Number <> 0 Then
        'On Error Goto 0
        '    PopulatePageArray
        'End If
        'On Error Goto 0

        If Not MyRefreshLayout Then ' refresh And draw NewLayout before swap
         Exit Sub
        End If

        Set ws1 = ThisWorkbook.Sheets("NewLayout")
        If Application.activeSheet Is ws1 Then
            lActivePage = GetActivePage
        Else
            lActivePage = 0
        End If

        NoOfPages = CInt(ThisWorkbook.Sheets("Settings").Range("B11"))
        If NoOfPages = 0 Then Exit Sub

            frmDelPages.Show (vbModal)
            If frmDelPages.GetRetVal <> 1 Then
                Call Unload(frmDelPages)
             Exit Sub
            End If
            vPageFrom = frmDelPages.GetPageFrom
            vPageTo = frmDelPages.GetPageTo
            vPageNo = frmDelPages.GetPageNo
            Call Unload(frmDelPages)
            '    If lActivePage < 2 Or lActivePage > NoOfPages Then
            '        lActivePage = Int(val(InputBox("Enter Page No To start deleting from:", , 2)))
            '        If lActivePage = 0 Then Exit Sub
            '
            '    End If
            If vPageNo = 0 Then
                lActivePage = vPageFrom
            Else
                lActivePage = vPageNo
            End If
            If lActivePage < 2 Or lActivePage > NoOfPages Then
                MsgBox "You can Not delete pages outside of book" & vbCrLf & "(Page: " & 2 & " To " & NoOfPages & ")"
             Exit Sub
            End If
            'lPagesToDelete = Int(val(InputBox("Enter number of pages To delete:", , 1)))
            If vPageNo = 0 Then
                lPagesToDelete = vPageTo - vPageFrom + 1
            Else
                lPagesToDelete = 1
            End If
            If lPagesToDelete = 0 Then Exit Sub
                If (lActivePage + lPagesToDelete - 1) > NoOfPages Then
                    MsgBox "You can Not delete pages outside of book" & vbCrLf & "(You want To delete up To " & (lActivePage + lPagesToDelete - 1) & " page in book of " & NoOfPages & " pages)"
                 Exit Sub
                End If

                ''Application.EnableEvents = False
                '    enable (False)
                '    StoreActiveSheets
                'Application.ScreenUpdating = False
                'commandOpRunning = True
                Set ws = ThisWorkbook.Sheets("Main List")
                'Set keysWs = ThisWorkbook.Sheets("KeysSheet")
                'Set matchWs = ThisWorkbook.Sheets("match")
                Application.ScreenUpdating = False
                Application.EnableEvents = False
                Dim sIssue As String
                'sIssue = ThisWorkbook.Worksheets("Main List").Range("CurrentIssue").value
                sIssue = GetCurrentIssue
                result = MyCheckLockAndReconcile("You can Not delete pages in this issue now", True, False)
                If Not result Then
                    'MyRestoreContext
                    Application.ScreenUpdating = True
                    Application.EnableEvents = True
                 Exit Sub
                End If
                'MySaveContext
                LockIssue sIssue


                'lRow = lastNonBlankRow(ws, "A")
                'lRow = GetNewRow - 1
                lRow = GetLastRow(ws)
                'Call DeletePagesFromPageArray(lPagesToDelete, lActivePage)
                For r = 2 To lRow
                    If ws.Range("G" & r) <> "" Then
                        '            r1 = findRowInKeysSheet(ws.Range("A" & r) & ws.Range("B" & r) & ws.Range("C" & r))
                        '            r1 = findRowInKeysSheet(ws.Range("M" & r))
                        'r1 = matchWs.Range("C" & r)
                        'If (r1 <> 0) Then
                        If CInt(ws.Range("G" & r)) >= lActivePage And CInt(ws.Range("G" & r)) < lActivePage + lPagesToDelete Then
                            ws.Range("G" & r & ":H" & r).ClearContents
                            '                        keysWs.Range("E" & r1 & ":N" & r1).ClearContents
                            ws.Cells(r, AdLastChangeCol + 1) = MakeID
                            gMainListArr(r - 1, 7) = ""
                            gMainListArr(r - 1, 8) = ""
                            gMainListArr(r - 1, 14) = ws.Cells(r, AdLastChangeCol + 1)
                        End If
                        If CInt(ws.Range("G" & r)) >= lActivePage + lPagesToDelete Then
                            ws.Range("G" & r) = CInt(ws.Range("G" & r)) - lPagesToDelete
                            ws.Cells(r, AdLastChangeCol + 1) = MakeID
                            gMainListArr(r - 1, 7) = ws.Range("G" & r)
                            gMainListArr(r - 1, 14) = ws.Cells(r, AdLastChangeCol + 1)
                            '                    keysWs.Range("E" & r1) = ws.Range("G" & r)
                        End If
                        'End If
                    End If
                Next r
                ' ThisWorkbook.Sheets("Settings").Range("B11") = NoOfPages - lPagesToDelete
                ' Application.StatusBar = "Pages deleted - Refreshing pages now"
                '    RestoreActiveSheets
                '    DoRefreshLayout ' Added by Hussam
                '    enable (True)
                'If (False) Then
                ' If Application.ActiveSheet Is ws1 Then
                '    Call DrawPages(lActivePage - 1)
                ' Else
                '    gNeedRedraw = True
                ' End If
                'End If
                maxPage = Application.WorksheetFunction.Max(ws.Range("G2:G" & lRow))
                If ThisWorkbook.Worksheets("Settings").Range("MaxNoPages").value <> maxPage Then
                    ThisWorkbook.Worksheets("Settings").Range("MaxNoPages").value = maxPage
                End If
                'Application.StatusBar = "Rechecking Pages ..."
                'CheckPages
                r = PagesStartRow + ((lActivePage - 2) \ 8) * RowsPerPage
                c = startCol + ((lActivePage - 2) Mod 8) * columnsPerPage
                'Application.EnableEvents = True
                'enable (True)
                'Application.ScreenUpdating = True

                'ThisWorkbook.Sheets("Layout").Activate

                MySaveIssue (False)
                UnlockIssue sIssue

                drawPages (2)
                'MyRestoreContext
                If Application.activeSheet Is ws1 Then
                    On Error Resume Next
                    ThisWorkbook.Sheets("NewLayout").Cells(r, c).Select
                    On Error Goto 0
                    End If
                    'Application.StatusBar = "Ready"

                    Application.ScreenUpdating = True
                    Application.EnableEvents = True

                    'commandOpRunning = False
End Sub

Sub DeletePages2()
    'brise samo Do max pageta
    If IDP Then Exit Sub
        If InName = False Then
            MsgBox "Select page where you want To start delete"
         Exit Sub
        End If

        Dim X As Long
        Dim lFilledPage As Long
        Dim lPagesToDelete As Long
        Dim lActivePage As Long
        Dim lPageReserve As Long
        Dim lMaxPage As Long
        Dim lMinPage As Long
        Dim bOdZadnjeNapunjeneStrane As Boolean

        If ThisWorkbook.Worksheets("Settings").Range("FindLastFilledCell").value = "Yes" Then
            bOdZadnjeNapunjeneStrane = True
        End If

        lMaxPage = ThisWorkbook.Worksheets("Settings").Range("MaxNoPages").value
        lMinPage = ThisWorkbook.Worksheets("Settings").Range("MinPageNo").value

        lActivePage = AktivnaStrana
        If lActivePage < lMinPage Or lActivePage > lMaxPage Then
            MsgBox "You can Not delete pages outside of book" & vbCrLf & "(Page: " & lMinPage & " To " & lMaxPage & ")"
         Exit Sub
        End If


        lPagesToDelete = Int(val(InputBox("Enter number of pages To delete:", , 1)))
        If lPagesToDelete = 0 Then Exit Sub

            If (lActivePage + lPagesToDelete - 1) > lMaxPage Then
                MsgBox "You can Not delete pages outside of book" & vbCrLf & "(You want To delete up To " & (lActivePage + lPagesToDelete - 1) & " page in book of " & lMaxPage & " pages)"
             Exit Sub
            End If

            If bOdZadnjeNapunjeneStrane = True Then
                lFilledPage = lMinPage
                'pocni sa krajnjom stranom KNJIGE pa smanjuj
                For X = lMaxPage To lMinPage Step -1
                    If IsEmptyPage2(X) = False Then
                        lFilledPage = X
                     Exit For
                    End If
                Next X
            End If

            '    lPageReserve = 849 - lFilledPage
            '    If lBookSize < lPagesToDelete Then
            '        MsgBox "You can Not delete " & lPagesToDelete & " in book of " & lBookSize & " pages!"
            '
            '        Exit Sub
            '    End If
            Dim r1 As Range, r2 As Range, r3 As Range
            'obrisi strane
            For X = lActivePage To lActivePage + lPagesToDelete - 1
                'Set r1 = ThisWorkbook.Names("Page_" & x).RefersToRange
                Set r1 = GetPageRange(X)
                r1.ClearContents
                r1.UnMerge
                With r1.Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                If IsBleed(X) = True Then
                    'obrisi bleedove
                    ResetBleed X
                End If
                'strane markirane ko kolor
                If IsYellow(X) = True Then
                    ResetYellow X
                End If
            Next X
            If bOdZadnjeNapunjeneStrane = False Then
                For X = (lActivePage + lPagesToDelete) To lMaxPage
                    SwapPagesModule X, X - lPagesToDelete
                Next X
            Else
                For X = lActivePage To lFilledPage
                    SwapPagesModule X, X + lPagesToDelete
                Next X
            End If
            'ThisWorkbook.Names("Page_" & lActivePage).RefersToRange.Select
            Set r1 = GetPageRange(lActivePage)
            r1.Select
End Sub

Sub DeletePages()
    If IDP Then Exit Sub
        If InName = False Then
            MsgBox "Select page where you want To start delete"
         Exit Sub
        End If

        Dim X As Long
        Dim lFilledPage As Long
        Dim lPagesToAdd As Long
        Dim lActivePage As Long
        Dim lPageReserve As Long
        Dim lBookSize As Long
        lBookSize = ThisWorkbook.Worksheets("Settings").Range("MaxNoPages").value
        lActivePage = AktivnaStrana
        If lActivePage < 2 Then Exit Sub
            lPagesToAdd = Int(val(InputBox("Enter number of pages To delete:", , 1)))
            If lPagesToAdd = 0 Then Exit Sub
                lFilledPage = 849    'pocni sa krajnjom stranom pa smanjuj
                For X = 849 To 2 Step -1
                    If IsEmptyPage(X) = False Then
                        lFilledPage = X
                     Exit For
                    End If
                Next X

                lPageReserve = 849 - lFilledPage
                If lBookSize < lPagesToAdd Then
                    MsgBox "You can Not delete " & lPagesToAdd & " in book of " & lBookSize & " pages!"

                 Exit Sub
                End If
                Dim r1 As Range, r2 As Range, r3 As Range
                'obrisi strane
                For X = lActivePage To lActivePage + lPagesToAdd - 1
                    'Set r1 = ThisWorkbook.Names("Page_" & x).RefersToRange
                    Set r1 = GetPageRange(X)
                    r1.ClearContents
                    r1.UnMerge
                    With r1.Interior
                        .Pattern = xlNone
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    If IsBleed(X) = True Then
                        'obrisi bleedove
                        ResetBleed X
                    End If
                    'strane markirane ko kolor
                    If IsYellow(X) = True Then
                        ResetYellow X
                    End If
                Next X

                For X = lActivePage To lFilledPage
                    SwapPagesModule X, X + lPagesToAdd
                Next X
                'ThisWorkbook.Names("Page_" & lActivePage).RefersToRange.Select
                Set r1 = GetPageRange(lActivePage)
                r1.Select
End Sub

Sub MyAddPages()
    If IDP Then Exit Sub
        '    If InName = False Then     'comment by hosam
        '        MsgBox "Select the page right before the New pages"
        '        Exit Sub
        '    End If
        Dim X As Long
        Dim ws As Worksheet
        '    Dim keysWs As Worksheet
        '    Dim matchWs As Worksheet
        Dim lFilledPage As Long
        Dim lPagesToAdd As Long
        Dim lActivePage As Long
        Dim lPageReserve As Long

        If Not MyRefreshLayout Then ' refresh And draw NewLayout before swap
         Exit Sub
        End If
        Set ws1 = ThisWorkbook.Sheets("NewLayout")
        If Application.activeSheet Is ws1 Then
            lActivePage = GetActivePage
        Else
            lActivePage = 0
        End If
        NoOfPages = CInt(ThisWorkbook.Sheets("Settings").Range("B11"))
        If NoOfPages = 0 Then Exit Sub
            'If lActivePage < 2 Or lActivePage > NoOfPages Then Exit Sub
            If lActivePage < 2 Or lActivePage > NoOfPages Then
                lActivePage = Int(val(InputBox("Enter Page No To add the pages after:", , 1)))
                If lActivePage = 0 Then Exit Sub

                End If
                lPagesToAdd = Int(val(InputBox("Enter number of pages To add:", , 25)))
                If lPagesToAdd = 0 Then Exit Sub

                    'Application.EnableEvents = False
                    'Application.ScreenUpdating = False
                    '    enable (False)
                    '    StoreActiveSheets
                    'commandOpRunning = True
                    Set ws = ThisWorkbook.Sheets("Main List")
                    'Set keysWs = ThisWorkbook.Sheets("KeysSheet")
                    'Set matchWs = ThisWorkbook.Sheets("match")
                    'lRow = lastNonBlankRow(ws, "A")

                    Application.ScreenUpdating = False
                    Application.EnableEvents = False
                    Dim sIssue As String
                    'sIssue = ThisWorkbook.Worksheets("Main List").Range("CurrentIssue").value
                    sIssue = GetCurrentIssue
                    result = MyCheckLockAndReconcile("You can Not add pages in this issue now", True, False)
                    If Not result Then
                        'MyRestoreContext
                        Application.ScreenUpdating = True
                        Application.EnableEvents = True
                     Exit Sub
                    End If
                    'MySaveContext
                    LockIssue sIssue
                    'lRow = GetNewRow - 1
                    lRow = GetLastRow(ws)
                    For r = 2 To lRow
                        If ws.Range("G" & r) <> "" Then
                            If CInt(ws.Range("G" & r)) > lActivePage Then
                                ws.Range("G" & r) = CInt(ws.Range("G" & r)) + lPagesToAdd
                                ws.Cells(r, AdLastChangeCol + 1) = MakeID
                                gMainListArr(r - 1, 7) = ws.Range("G" & r)
                                gMainListArr(r - 1, 14) = ws.Cells(r, AdLastChangeCol + 1)
                                '                r1 = findRowInKeysSheet(ws.Range("A" & r) & ws.Range("B" & r) & ws.Range("C" & r))
                                ''                r1 = findRowInKeysSheet(ws.Range("M" & r))
                                ''                'r1 = matchWs.Range("C" & r)
                                ''                If (r1 <> 0) Then
                                ''                    keysWs.Range("E" & r1) = ws.Range("G" & r)
                                ''                End If
                            End If
                        End If
                    Next r
                    'ThisWorkbook.Sheets("Settings").Range("B11") = NoOfPages + lPagesToAdd
                    ' Application.StatusBar = "Adding Pages... "
                    '' Call AddPagesToPageArray(lPagesToAdd, lActivePage)
                    'PopulatePageArray
                    'Application.StatusBar = "Pages added - Refreshing pages now"

                    ' RestoreActiveSheets
                    ''    DoRefreshLayout ' Added by Hussam
                    ''    enable (True)

                    '' If (False) Then
                    ''    If Application.ActiveSheet Is ws1 Then
                    ''       Call DrawPages(lActivePage + 1)
                    ''    Else
                    ''       gNeedRedraw = True
                    ''    End If
                    '' End If
                    maxPage = Application.WorksheetFunction.Max(ws.Range("G2:G" & lRow))
                    If ThisWorkbook.Worksheets("Settings").Range("MaxNoPages").value <> maxPage Then
                        ThisWorkbook.Worksheets("Settings").Range("MaxNoPages").value = maxPage
                    End If
                    ' Application.StatusBar = "Rechecking Pages ..."
                    'CheckPages
                    r = PagesStartRow + ((lActivePage - 2) \ 8) * RowsPerPage
                    c = startCol + ((lActivePage - 2) Mod 8) * columnsPerPage

                    'Application.EnableEvents = True
                    'Application.ScreenUpdating = True
                    'enable (True)
                    'ThisWorkbook.Sheets("Layout").Activate

                    MySaveIssue (False)
                    UnlockIssue sIssue

                    drawPages (2)
                    ' MyRestoreContext
                    If Application.activeSheet Is ws1 Then
                        On Error Resume Next
                        ThisWorkbook.Sheets("NewLayout").Cells(r, c).Select
                        On Error Goto 0
                        End If
                        ' Application.StatusBar = "Ready"
                        'commandOpRunning = False
                        Application.ScreenUpdating = True
                        Application.EnableEvents = True

End Sub
Sub AddPages()
    If IDP Then Exit Sub
        If InName = False Then
            MsgBox "Select the page right before the New pages"
         Exit Sub
        End If
        Dim X As Long
        Dim lFilledPage As Long
        Dim lPagesToAdd As Long
        Dim lActivePage As Long
        Dim lPageReserve As Long
        lActivePage = AktivnaStrana
        If lActivePage < 2 Then Exit Sub
            lPagesToAdd = Int(val(InputBox("Enter number of pages To add:", , 25)))
            If lPagesToAdd = 0 Then Exit Sub
                lFilledPage = 849    'pocni sa krajnjom stranom pa smanjuj
                For X = 849 To 2 Step -1
                    If IsEmptyPage(X) = False Then
                        lFilledPage = X
                     Exit For
                    End If
                Next X
                lPageReserve = 849 - lFilledPage
                If lPageReserve < lPagesToAdd Then
                    MsgBox "No place To add " & lPagesToAdd & " pages!"
                 Exit Sub
                End If

                Dim r1 As Range, r2 As Range, r3 As Range
                'prvo gurni sve strane od aktivne+1 To poslednje zauzete
                For X = lFilledPage To lActivePage + 1 Step -1
                    SwapPagesModule X, X + lPagesToAdd
                Next X


                If lFilledPage + lPagesToAdd > ThisWorkbook.Worksheets("Settings").Range("MaxNoPages").value Then
                    ThisWorkbook.Worksheets("Settings").Range("MaxNoPages").value = lFilledPage + lPagesToAdd
                    CheckPages
                End If
End Sub
Sub ExportNumChart()
    Const FName As String = "C:\Numbers.jpg"
    Dim pic_rng As Range
    Dim ShTemp As Worksheet
    Dim ChTemp As Chart
    Dim PicTemp As Picture
    Application.ScreenUpdating = False
    Set pic_rng = Worksheets("Layout").Range("Page_2")
    Set ShTemp = Worksheets.Add
    Charts.Add
    ActiveChart.Location Where:=xlLocationAsObject, Name:=ShTemp.Name
    Set ChTemp = ActiveChart
    pic_rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    ChTemp.Paste
    Set PicTemp = Selection
    With ChTemp.Parent
        .Width = PicTemp.Width + 8
        .Height = PicTemp.Height + 8
    End With
    ChTemp.Export Filename:="c:\Users\shonius\Documents\Rad\Elance\2013-06-25  Excel To Indesign4\Numbers.jpg", FilterName:="jpg"
    'UserForm1.Image1.Picture = LoadPicture(FName)
    'Kill FName
    Application.DisplayAlerts = False
    ShTemp.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
Sub ListAdsFromLayoutSheet()
    If IDP Then Exit Sub
        ListAds
        ThisWorkbook.Sheets("Layout").Activate
End Sub
Sub FilterAds()
    'remove + sign And text after it from ad list And move results in New Excel file
    Dim r As Range, rResiz As Range, r2 As Range
    Dim X As Long, Y As Long
    Dim SourceAr()
    Dim wbNew As Workbook

    ThisWorkbook.Sheets("Report").Activate
    Set r = Range("g9", Range("h10000").End(xlUp))
    r.Select
    Set rResiz = r.Resize(columnsize:=3)
    If r.Cells.count < 1 Then
     Exit Sub
    End If
    rResiz.Select
    SourceAr = rResiz
    For X = LBound(SourceAr) To UBound(SourceAr)
        Y = InStr(1, SourceAr(X, 1), "+")
        If Y > 0 Then
            SourceAr(X, 1) = Left(SourceAr(X, 1), Y - 1)
        End If
    Next X

    Set wbNew = Workbooks.Add
    wbNew.Worksheets(1).Activate
    wbNew.Worksheets(1).Range("a1").value = "Ad:"
    wbNew.Worksheets(1).Range("b1").value = "Page:"
    wbNew.Worksheets(1).Range("c1").value = "Ad file:"
    wbNew.Worksheets(1).Range("a1:c1").Font.Bold = True
    Set r2 = wbNew.Worksheets(1).Range("a2")
    r2.Resize(UBound(SourceAr), 3).value = SourceAr
    wbNew.Worksheets(1).columns("A:C").EntireColumn.AutoFit
End Sub
Sub ListAds(Optional bAllPages As Boolean = False)    'kad se uvoze rezervacije potrebne su sve strane
    If IDP Then Exit Sub
        Dim nName As Name
        'Dim cPages As New Collection
        Dim cPUnits As New Collection
        Dim cPUnitsSizes As New Collection
        Dim cpErrors As New Collection
        Dim cErrors As Collection
        Dim cUnits As Collection
        Dim cUnitsSizes As Collection
        Dim cUnitsColumns As Collection
        Dim cUnitsPositions As Collection

        Dim cAds As New Collection
        Dim cFiles As New Collection
        Dim cOnPage As New Collection

        Dim X As Long, Y As Long, z As Long
        Dim x2 As Long    'brojac za imena
        Dim r As Range, rc As Range
        Dim r2 As Range

        Dim rCaption As Range, rTemp As Range

        Dim bAtLeastOneUnitOnPage As Boolean
        Dim lUnitsTotalPerPage As Long
        Dim lUnitsFilledPerPage As Long
        Dim lFilesPerPage As Long
        Dim lPercent As Long


        Dim aSize(1 To 7)    'row,col,Indesign width,height,filename,adname
        Dim aPos(1 To 2)    'top ,left
        Dim vc As Variant
        Dim lMinPage As Long
        Dim lMaxPages As Long
        Dim sFile As String

        '****vars For report
        Dim lEmptyPages As Long
        Dim lCompletedPages As Long
        Dim lPagesInProgress As Long
        Dim lErrorPages As Long
        Dim lErrorsPerPage As Long
        Dim rRep As Range
        Dim rRep2 As Range    'abecedno sortiranje
        Dim x3 As Long    'counter For report
        '****
        ThisWorkbook.Sheets("Layout").Activate
        If bAllPages = False Then
            lMinPage = Range("MinPageNo").value
            lMaxPages = Range("MaxNoPages").value
        Else
            lMinPage = 2
            lMaxPages = 401
        End If

        For x2 = lMinPage To lMaxPages
            '        SetPageRangeAddress (x2)
            '        Set NName = ThisWorkbook.Names("Page_" & x2)

            Set cUnits = New Collection
            Set cUnitsSizes = New Collection

            '            tmpStr = NName.RefersTo
            '            vArr = Split(tmpStr, "!")
            '            Set r = ThisWorkbook.Sheets("NewLayout").Range(vArr(1))
            ' Set r = NName.RefersToRange
            Set r = GetPageRange(x2)
            For Each rc In r.Cells
                If rc.MergeCells = True Then
                    On Error Resume Next
                    cUnits.Add rc.MergeArea, rc.MergeArea.Address
                    On Error Goto 0
                    Else
                        cUnits.Add rc
                    End If
                Next rc

                cPUnits.Add cUnits

                For Each vc In cUnits
                    Set r2 = vc
                    'ad name
                    aSize(6) = CStr(r2.Cells(1, 1).value)
                    'file name
                    aSize(5) = GetPathFromComment(r2.Cells(1, 1))

                    If aSize(6) <> "" Then
                        cAds.Add aSize(6)
                        cFiles.Add aSize(5)
                        cOnPage.Add x2
                    End If
                Next vc


            Next x2


            With ThisWorkbook.Sheets("Report")
                With .Range("d8:f3200")
                    .ClearContents
                    .Font.ColorIndex = xlAutomatic
                End With

                With .Range("g8:i3200")
                    .ClearContents
                    .Font.ColorIndex = xlAutomatic
                End With

                .Range("d7").value = "Ad:"
                .Range("e7").value = "Page:"
                .Range("f7").value = "Ad file:"

                .Range("g7").value = "Ad:"
                .Range("h7").value = "Page:"
                .Range("i7").value = "Ad file:"
            End With

            Set rRep = ThisWorkbook.Sheets("Report").Range("D8")

            For X = 1 To cAds.count
                rRep.Offset(X, 0).value = cAds(X)
                rRep.Offset(X, 1).value = cOnPage(X)
                rRep.Offset(X, 2).value = cFiles(X)
                If cFiles(X) <> "" Then
                    If DirU(cFiles(X)) = "" Then
                        With rRep.Offset(X, 2).Font
                            .Color = -16776961
                            .TintAndShade = 0
                        End With
                    End If

                End If
            Next X
            ThisWorkbook.Sheets("Report").Activate
            Range("d7").Select
            If cAds.count < 1 Then Exit Sub
                Range("d9", rRep.Offset(X - 1, 2)).Select

                Range("d9:d3200").NumberFormat = "@"
                Range("d9", rRep.Offset(X - 1, 2)).Copy
                Range("g9").PasteSpecial
                ThisWorkbook.Worksheets("Report").Sort.SortFields.Clear
                ThisWorkbook.Worksheets("Report").Sort.SortFields.Add key:=Range("G9"), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                With ThisWorkbook.Worksheets("Report").Sort
                    .SetRange Range("G9", rRep.Offset(X - 1, 5))
                    .header = xlNo
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With

                With Range("G9", rRep.Offset(X - 1, 3)).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ThemeColor = 6
                    '        .TintAndShade = 0.399945066682943
                    .Weight = xlThin
                End With
                Range("d7").Select
                Application.CutCopyMode = False
End Sub

Sub MarkYellow()
    If InName Then
        If ActiveCell.Interior.Color = 65535 Then
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Else
            With ActiveCell.Interior
                Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
    Else
        MsgBox "Select cell in page"
    End If

End Sub
Sub CheckPagesfromReportSheet()
    If IDP Then Exit Sub
        CheckPages
        ThisWorkbook.Sheets("Report").Activate
End Sub
Sub CheckPages()
    If IDP Then Exit Sub
        Dim nName As Name
        'Dim cPages As New Collection
        Dim cPUnits As New Collection
        Dim cPUnitsSizes As New Collection
        Dim cpErrors As New Collection
        Dim cErrors As Collection
        Dim cUnits As Collection
        Dim cUnitsSizes As Collection
        Dim cUnitsColumns As Collection
        Dim cUnitsPositions As Collection
        Dim X As Long, Y As Long, z As Long
        Dim x2 As Long    'brojac za imena
        Dim r As Range, rc As Range
        Dim r2 As Range

        Dim rCaption As Range, rTemp As Range

        Dim bAtLeastOneUnitOnPage As Boolean
        Dim lUnitsTotalPerPage As Long
        Dim lUnitsFilledPerPage As Long
        Dim lFilesPerPage As Long
        Dim lPercent As Long


        Dim aSize(1 To 6)    'row,col,Indesign width,height,filename, ad name
        Dim aPos(1 To 2)    'top ,left
        Dim vc As Variant
        Dim lMaxPages As Long
        Dim lMinPages As Long
        Dim sFile As String

        '****vars For report
        Dim lEmptyPages As Long
        Dim lCompletedPages As Long
        Dim lPagesInProgress As Long
        Dim lErrorPages As Long
        Dim lErrorsPerPage As Long
        Dim rRep As Range
        Dim x3 As Long    'counter For report
        Dim ws As Worksheet

        'On Error Resume Next
        'bookPageCount = UBound(PagesArr)
        'If Err.Number <> 0 Then
        'On Error Goto 0
        '    PopulatePageArray
        'End If
        'On Error Goto 0
        '
        '****
        'Set ws = ThisWorkbook.Sheets("NewLayout")
        Application.ScreenUpdating = False
        Application.EnableEvents = False

        ThisWorkbook.Sheets("NewLayout").Activate
        lMaxPages = ThisWorkbook.Sheets("Settings").Range("MaxNoPages").value
        lMinPages = ThisWorkbook.Sheets("Settings").Range("MinPageNo").value
        lMinPage = 2
        lMaxPages = ThisWorkbook.Sheets("Settings").Range("B11").value
        '    lMinPages = 2
        '    lMaxPages = UBound(PagesArr)
        If lMinPages < 2 Then lMinPages = 2
            '    If lMaxPages < 2 Or lMaxPages > UBound(PagesArr) Then lMaxPages = UBound(PagesArr)
            For x2 = lMinPages To lMaxPages
                '        Set NName = ThisWorkbook.Names("Page_" & x2)
                '        bAtLeastOneUnitOnPage = False
                '        For J = 1 To 8
                '            If PagesArr(x2).AssignedAdRow(J) <> 0 Then
                '                bAtLeastOneUnitOnPage = True
                '                Exit For
                '            End If
                '        Next J
                '''        bAtLeastOneUnitOnPage = PageHasAllocations(x2)
                '        SetPageRangeAddress (x2)
                '        Set NName = ThisWorkbook.Names("Page_" & x2)
                '''        If bAtLeastOneUnitOnPage Then
                Set cUnits = New Collection
                Set cUnitsSizes = New Collection
                '''            tmpStr = NName.RefersTo
                '''            vArr = Split(tmpStr, "!")
                '''            Set r = ThisWorkbook.Sheets("NewLayout").Range(vArr(1))


                '            Set r = NName.RefersToRange
                Set r = GetPageRange(x2)
                For Each rc In r.Cells
                    If rc.MergeCells = True Then
                        On Error Resume Next
                        cUnits.Add rc.MergeArea, rc.MergeArea.Address
                        On Error Goto 0
                        Else
                            cUnits.Add rc
                        End If
                    Next rc

                    cPUnits.Add cUnits
                    lErrorsPerPage = 0
                    lUnitsFilledPerPage = 0
                    lUnitsTotalPerPage = cUnits.count
                    bAtLeastOneUnitOnPage = False

                    For Each vc In cUnits
                        Set r2 = vc

                        'file name
                        aSize(5) = GetPathFromComment(r2.Cells(1))
                        aSize(6) = CStr(r2.Cells(1, 1).value)
                        'determine color
                        If aSize(5) <> "" And aSize(6) <> "" Then
                            'normalan slucaj: i celija i komentar postoje
                            aSize(5) = DirU(CStr(aSize(5)))
                            If aSize(5) = "" Then    'file does Not exist, altough it is listed in commeny - put patern
                                With r2.Cells(1, 1).Interior
                                    .Pattern = xlGray16
                                    .PatternColor = 255
                                End With
                                lErrorsPerPage = lErrorsPerPage + 1
                                lErrorPages = lErrorPages + 1
                                bAtLeastOneUnitOnPage = True    ' mark presence altough file is missing To draw gradient
                            Else    'file is here
                                With r2.Cells(1, 1).Interior
                                    .Pattern = xlSolid
                                    .PatternColor = 255
                                End With
                                lUnitsFilledPerPage = lUnitsFilledPerPage + 1
                                bAtLeastOneUnitOnPage = True
                            End If

                            cUnitsSizes.Add aSize
                        Elseif aSize(5) <> "" And aSize(6) = "" Then
                            With r2.Cells(1, 1).Interior
                                .Pattern = xlGray16
                                .PatternColor = 255
                            End With
                            lErrorsPerPage = lErrorsPerPage + 1
                            lErrorPages = lErrorPages + 1
                            bAtLeastOneUnitOnPage = True    ' mark presence altough file is missing To draw gradient
                        Elseif r2.Cells(1).Comment Is Nothing = False And aSize(6) = "" Then
                            With r2.Cells(1, 1).Interior
                                .Pattern = xlGray16
                                .PatternColor = 255
                            End With
                            lErrorsPerPage = lErrorsPerPage + 1
                            lErrorPages = lErrorPages + 1
                            bAtLeastOneUnitOnPage = True    ' mark presence altough file is missing To draw gradient

                        Else
                            If aSize(6) <> "" Then
                                With r2.Cells(1, 1).Interior
                                    .Pattern = xlGray8
                                    .PatternColor = 255
                                End With
                                lErrorsPerPage = lErrorsPerPage + 1
                                lErrorPages = lErrorPages + 1
                                bAtLeastOneUnitOnPage = True

                            Else
                                'koristi da ovde resetujes celiju ako ima nekog formata a prazna je
                                If r2.Cells(1, 1).Interior.Pattern <> xlNone Then    ' Or r2.Cells(1, 1).Interior.TintAndShade <> 0 'mozda?
                                    With r2.Cells(1, 1).Interior
                                        .Pattern = xlNone
                                        .TintAndShade = 0
                                        .PatternTintAndShade = 0
                                    End With
                                End If
                            End If
                        End If
                    Next vc
                    Set rTemp = r.Range("a1")
                    Set rCaption = rTemp.Offset(-1, 0)
                    'End If
                    If bAtLeastOneUnitOnPage = False Then
                        lEmptyPages = lEmptyPages + 1
                        lPercent = 0
                        cpErrors.Add "Empty page"
                        'cpErrors.Add NName.Name
                        cpErrors.Add "Page_" & x2
                        cpErrors.Add "Page " & x2
                        '            Call FillGradient(rCaption, lPercent)
                        Call FillGradient(r, lPercent)

                    Elseif lUnitsFilledPerPage = lUnitsTotalPerPage Then
                        lCompletedPages = lCompletedPages + 1
                        lPercent = 100
                        cpErrors.Add "Completed page"
                        'cpErrors.Add NName.Name
                        cpErrors.Add "Page_" & x2
                        cpErrors.Add "Page " & x2
                        Call FillGradient(rCaption, lPercent)
                    Elseif bAtLeastOneUnitOnPage = True Then
                        lPagesInProgress = lPagesInProgress + 1
                        lPercent = (lUnitsFilledPerPage / lUnitsTotalPerPage) * 100
                        If lPercent = 0 Then lPercent = 1    'To draw gradient If page contain only errors
                            If lErrorsPerPage < 1 Then
                                cpErrors.Add "Page in progress"
                                'cpErrors.Add NName.Name
                                cpErrors.Add "Page_" & x2
                                cpErrors.Add "Page " & x2
                            Else
                                cpErrors.Add "Page in progress (With " & lErrorsPerPage & " error[s])"
                                'cpErrors.Add NName.Name
                                cpErrors.Add "Page_" & x2
                                cpErrors.Add "Page " & x2
                            End If
                            Call FillGradient(rCaption, lPercent)
                        End If
                        'cPUnitsSizes.Add cUnitsSizes
                    Next x2
                    MarkLastPage
                    With ThisWorkbook.Sheets("Report")
                        .Range("a1:b1100").ClearContents
                        '.Range("a8:b1000").Clear
                        .Range("a1").value = "Overview:"
                        .Range("a2").value = lCompletedPages & " Completed pages"
                        .Range("b2").value = lCompletedPages / (lMaxPages - 1)
                        .Range("a3").value = lEmptyPages & " Empty pages"
                        .Range("b3").value = lEmptyPages / (lMaxPages - 1)
                        .Range("a4").value = lPagesInProgress & " Pages in progress"
                        .Range("b4").value = lPagesInProgress / (lMaxPages - 1)
                        .Range("a5").value = lErrorPages & " Pages With errors"
                        .Range("b5").value = lErrorPages / (lMaxPages - 1)
                        .Range("a7").value = "Details:"
                    End With

                    Set rRep = ThisWorkbook.Sheets("Report").Range("a8")
                    For X = 1 To cpErrors.count Step 3
                        If cpErrors(X) <> "Completed page" Then
                            x3 = x3 + 1
                            rRep.Offset(x3, 0).value = cpErrors(X + 2) & " " & cpErrors(X)
                            rRep.Offset(x3, 0).Hyperlinks.Add Anchor:=rRep.Offset(x3, 0), Address:="", SubAddress:=cpErrors(X + 1)
                            'Debug.Print cpErrors(x + 2), cpErrors(x)
                        End If
                    Next X
                    Application.ScreenUpdating = True
                    Application.EnableEvents = True

End Sub
Sub FillGradient(r As Range, Byval lPrecent As Long)
    'make range gradient according To percent

    If lPrecent < 100 And lPrecent > 0 Then
        With r.Interior
            .Pattern = xlPatternLinearGradient
            .Gradient.Degree = 0
            .Gradient.ColorStops.Clear
        End With
        With r.Interior.Gradient.ColorStops.Add(0)
            .Color = vbGreen
            .TintAndShade = 0
        End With
        With r.Interior.Gradient.ColorStops.Add(lPrecent / 100)
            .Color = vbGreen
            .TintAndShade = 0
        End With
        With r.Interior.Gradient.ColorStops.Add(1)
            .Color = 255
            .TintAndShade = 0
        End With
    Elseif lPrecent = 100 Then
        With r.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 11263438
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    Elseif lPrecent = 0 Then
        With r.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
End Sub


Sub BuildLayout()
    If IDP Then Exit Sub
        Const StrokeWidth As Double = 1.38888888888889E-02
        Dim nName As Name
        Dim cPages As New Collection
        Dim cPUnits As New Collection
        Dim cPUnitsSizes As New Collection
        Dim cpUnitsPositions As New Collection
        Dim cUnits As Collection
        Dim cUnitsSizes As Collection
        Dim cUnitsColumns As Collection
        Dim cUnitsPositions As Collection
        Dim X As Long, Y As Long, z As Long
        Dim x2 As Long    'brojac za imena
        Dim r As Range, rc As Range
        Dim r2 As Range
        Dim aSize(1 To 6)    'row,col,Indesign width,height,filename,jel bled [onda 1, inace 0], ad name
        Dim aPos(1 To 2)    'top ,left
        Dim vc As Variant
        Dim lMaxPages As Long

        Dim iDA   ' As InDesign.Application
        Dim iDDoc   ' As InDesign.Document
        Dim iDPage    'As InDesign.Page
        Dim idColor    'As InDesign.Color
        Dim iDRectangle    ' As InDesign.Rectangle
        Dim myY1, myX1, myY2, myX2    'geometric bounds For rectangle
        Dim sFile As String
        Dim sTempFile As String    ' ako nema fajla [a upisano je nesto ko ad, ubaci To ko text]


        Dim lMinPage As Long
        Dim lNoOfPages As Long    'number of exported pages

        Dim bOddPage As Boolean    '?
        Dim bCS3fix As Boolean    'Do Not empty frames If cs3fix is active


        BuildNames ' added by Hussam

        If ThisWorkbook.Sheets("Settings").Range("CS3fix") = "Yes" Then bCS3fix = True

            If ThisWorkbook.Sheets("Settings").Range("IndesignTemplate").value = "" Then
                MsgBox "InDesign template is Not defined!"
             Exit Sub
            Elseif Dir(ThisWorkbook.Sheets("Settings").Range("IndesignTemplate").value) = "" Then
                MsgBox "InDesign template: " & ThisWorkbook.Sheets("Settings").Range("IndesignTemplate").value & " is Not found!"
             Exit Sub
            End If
            If DirExists("C:\regkey") = False Then
                MsgBox "Folder ""C:\regkey"" must be present!"
             Exit Sub
            End If

            Dim dHbleed As Double, dVbleed As Double    'za bleed
            dHbleed = val(ThisWorkbook.Sheets("Settings").Range("Bled_Size_Horizontal").value)
            dVbleed = val(ThisWorkbook.Sheets("Settings").Range("Bled_Size_Vertical").value)

            'ThisWorkbook.Sheets("Layout").Activate
            ThisWorkbook.Sheets("NewLayout").Activate

            lMaxPages = ThisWorkbook.Sheets("Settings").Range("MaxNoPages").value

            'add dummy pages at 1, To keep counters sinhronized. It was needed when Yaakov asked For removing of first page
            Application.StatusBar = "Preparing..."
            cPages.Add 0&
            cPUnits.Add 0&
            cPUnitsSizes.Add 0&
            cpUnitsPositions.Add 0&

            lMinPage = ThisWorkbook.Sheets("Settings").Range("MinPageNo").value
            lNoOfPages = lMaxPages - lMinPage + 1
            For x2 = lMinPage To lMaxPages
                '        SetPageRangeAddress (x2)
                '        Set NName = ThisWorkbook.Names("Page_" & x2)

                Set cUnits = New Collection
                Set cUnitsSizes = New Collection
                Set cUnitsPositions = New Collection

                '            tmpStr = NName.RefersTo
                '            vArr = Split(tmpStr, "!")
                '            Set r = ThisWorkbook.Sheets("NewLayout").Range(vArr(1))


                'Set r = NName.RefersToRange
                Set r = GetPageRange(x2)

                cPages.Add r
                For Each rc In r.Cells
                    If rc.MergeCells = True Then
                        On Error Resume Next
                        cUnits.Add rc.MergeArea, rc.MergeArea.Address
                        On Error Goto 0
                        Else
                            cUnits.Add rc
                        End If
                    Next rc
                    cPUnits.Add cUnits
                    For Each vc In cUnits
                        Set r2 = vc
                        aSize(1) = r2.rows.count
                        aSize(2) = r2.columns.count
                        aSize(3) = GetSizes(aSize(2), aSize(1))(1)
                        aSize(4) = GetSizes(aSize(2), aSize(1))(2)
                        'file name
                        aSize(5) = GetPathFromComment(r2)
                        'pokupi ime reklame
                        aSize(7) = r2.Cells(1).value
                        'old way
                        'aSize(5) = CStr(r2.Cells(1, 1).Value)
                        'determine color
                        'ne treba vise nema vise color i bw vadjenja, ime fajla i To je To
                        '            If r2.Cells(1, 1).Interior.Color = 65535 Then
                        '                'color
                        '                If aSize(5) <> "" Then
                        '
                        '
                        '                End If
                        '            Else    'bw
                        '                If aSize(5) <> "" Then
                        '                    sTempFile = GetFirstFile(CStr(aSize(5)), False)
                        '                    If sTempFile <> "" Then
                        '                        aSize(5) = sTempFile
                        '                    Else
                        '
                        '
                        '                    End If
                        '
                        '                End If
                        '
                        '            End If
                        If IsBleed(x2) = True Then
                            If aSize(1) = 4 And aSize(2) = 2 Then    'full page
                                aSize(3) = GetBleedPageSizes()(1) + dHbleed + dHbleed
                                aSize(4) = GetBleedPageSizes()(2) + dVbleed + dVbleed
                            End If
                        End If
                        cUnitsSizes.Add aSize
                        aPos(1) = GetPositions(r2.Cells(1, 1), r)(1)
                        aPos(2) = GetPositions(r2.Cells(1, 1), r)(2)

                        If aSize(1) = 4 And aSize(2) = 2 Then    'full page
                            aPos(1) = ThisWorkbook.Sheets("Settings").Range("Full_Page_Position_1_Top").value
                            aPos(2) = ThisWorkbook.Sheets("Settings").Range("Full_Page_Position_1_Left").value
                        End If
                        If IsBleed(x2) = True Then
                            If aSize(1) = 4 And aSize(2) = 2 Then    'full page
                                '                    aPos(1) = aPos(1) - dVbleed
                                '                    aPos(2) = aPos(2) - dHbleed
                                aPos(1) = ThisWorkbook.Sheets("Settings").Range("f24").value
                                aPos(2) = ThisWorkbook.Sheets("Settings").Range("g24").value
                            End If
                        End If
                        cUnitsPositions.Add aPos
                    Next vc
                    cPUnitsSizes.Add cUnitsSizes
                    cpUnitsPositions.Add cUnitsPositions

                Next x2

                'id fun

                Dim sw As New StopWatch
                sw.StartTimer
                Application.StatusBar = "Starting InDesign"
                Set iDA = CreateObject("InDesign.Application")
                Dim idalerts    'da vratim nivo kukanja na pocetni nivo
                Dim idunitsHOR    'horizontalne jedinice
                Dim idunitsVER    'vertikalne jedinice

                idalerts = iDA.ScriptPreferences.UserInteractionLevel
                iDA.ScriptPreferences.UserInteractionLevel = 1699640946    'idUserInteractionLevels.idNeverInteract
                ' iDA.ScriptPreferences.EnableRedraw = False
                'iDA.Visible = True
                Set iDDoc = iDA.Open(Range("IndesignTemplate").value)
                'Set iDDoc = iDA.Open(Range("IndesignTemplate").Value, False) 'otvara dok skrivenim prozorom. oko duplo brze!
                iDA.Windows(1).Minimize    'ako je prozor skriven ovo pravi gresku!
                ThisWorkbook.Activate
                'Set jedinice mere
                idunitsHOR = iDDoc.ViewPreferences.HorizontalMeasurementUnits
                iDDoc.ViewPreferences.HorizontalMeasurementUnits = 2053729891    'idMeasurementUnits.idInches
                idunitsVER = iDDoc.ViewPreferences.HorizontalMeasurementUnits
                iDDoc.ViewPreferences.VerticalMeasurementUnits = 2053729891    'idMeasurementUnits.idInches
                ThisWorkbook.Activate

                Application.StatusBar = "Adding pages To InDesign"
                For X = lMinPage To lMaxPages
                    iDDoc.Pages.Add 1701733408    'idAtEnd
                    Application.StatusBar = "Adding pages To InDesign: " & X & " of " & lMaxPages
                Next X

                ' We'll need To create a color. Check To see If the color already exists.
                On Error Resume Next
                Set idColor = iDDoc.Colors.Item("YaakovBlack")
                If Error.Number <> 0 Then
                    Set idColor = iDDoc.Colors.Add
                    idColor.Name = "YaakovBlack"
                    idColor.Model = idColorModel.idProcess
                    idColor.ColorValue = Array(0, 0, 0, 100)
                    Error.Clear
                End If
                ' Resume normal error handling.
                On Error Goto 0

                    'peneh-start
                    ' Get all page items in the document
                    Set pageItems = iDDoc.pageItems

                    ' Redefine the array To hold all page items
                    ReDim sortedPageItems(1 To pageItems.count)

                    ' Populate the array With page items
                    For i = 1 To pageItems.count
                        Set sortedPageItems(i) = pageItems(i)
                    Next i

                    ' Sort the array by page item id using Bubble Sort (Or another sort method)
                    For i = 1 To UBound(sortedPageItems) - 1
                        For J = i + 1 To UBound(sortedPageItems)
                            If sortedPageItems(i).ID > sortedPageItems(J).ID Then
                                ' Swap the items
                                Set Temp = sortedPageItems(i)
                                Set sortedPageItems(i) = sortedPageItems(J)
                                Set sortedPageItems(J) = Temp
                            End If
                        Next J
                    Next i

                    'hendle cell color
                    articleColor = ThisWorkbook.Sheets("Additional_settings").Range("b7").Interior.Color

                    If articleColor = 0 Then
                        MsgBox "The Article Color have no fill."

                    End If

                    For X = 2 To lNoOfPages + 1
                        Application.StatusBar = "Creating boxes For page: " & X & " (of " & lNoOfPages & ")"
                        Set iDPage = iDDoc.Pages(X)

                        For Y = 1 To cPUnits(X).count
                            Set vc = cPUnits(X).Item(Y)
                            myY1 = cpUnitsPositions(X).Item(Y)(1)
                            myY2 = myY1 + cPUnitsSizes(X).Item(Y)(4)
                            myX1 = cpUnitsPositions(X).Item(Y)(2)
                            myX2 = myX1 + cPUnitsSizes(X).Item(Y)(3)
                            sFile = cPUnitsSizes(X).Item(Y)(5)
                            sTempFile = cPUnitsSizes(X).Item(Y)(7)

                            ' Get current cell color
                            currentColor = vc.Interior.Color
                            Debug.Print "current color is: " & currentColor & " article Color is: " & articleColor

                            'If is Not an Article
                            If currentColor <> articleColor Then
                                ' create Rectangle And place the ad in it
                                Set iDRectangle = iDPage.Rectangles.Add
                                iDRectangle.GeometricBounds = Array(CDbl(myY1), CDbl(myX1), CDbl(myY2), CDbl(myX2))

                                'Set Rectangle stroke
                                If ThisWorkbook.Worksheets("Settings").Range("BlackBorder").value > 0 Then
                                    iDRectangle.StrokeWeight = ThisWorkbook.Worksheets("Settings").Range("BlackBorder").value
                                    iDRectangle.StrokeColor = iDDoc.Swatches.Item("YaakovBlack")
                                    iDRectangle.StrokeAlignment = 1936998729 'idInsideAlignment
                                Else
                                    iDRectangle.StrokeWeight = 0
                                End If

                                'Set the source file
                                If Len(sFile) > 0 And DirU(sFile) <> "" Then

                                    If getExt(sFile) <> "indd" Then
                                        iDRectangle.place sFile
                                        iDRectangle.Fit 1668575078 'idFitOptions.idContentToFrame
                                    Else
                                        iDRectangle.place sFile, True
                                        iDRectangle.Fit 1668575078 'idFitOptions.idContentToFrame
                                    End If
                                Else
                                    'Set options For frame To fit when place file manually




                                    If bCS3fix = False Then
                                        iDRectangle.FrameFittingOptions.AutoFit = True
                                        iDRectangle.FrameFittingOptions.FittingOnEmptyFrame = 1668575078    'idEmptyFrameFittingOptions.idContentToFrame
                                    End If

                                    If Len(sFile) > 0 And DirU(sFile) = "" Then
                                        If Range("WriteText").value = "Yes" Then
                                            ' Str2TXT CStr(sFile), ThisWorkbook.Path & "\mfile.txt"
                                            Str2TXT CStr(sFile), "C:\regkey\mfile.txt"
                                            'iDRectangle.Place ThisWorkbook.Path & "\mfile.txt"
                                            iDRectangle.place "C:\regkey\mfile.txt"
                                        End If
                                    Elseif Len(sFile) = 0 And Len(sTempFile) > 0 Then
                                        If Range("WriteText").value = "Yes" Then
                                            ' Str2TXT CStr(sFile), ThisWorkbook.Path & "\mfile.txt"
                                            Str2TXT CStr(sTempFile), "C:\regkey\mfile.txt"
                                            'iDRectangle.Place ThisWorkbook.Path & "\mfile.txt"
                                            iDRectangle.place "C:\regkey\mfile.txt"
                                        End If
                                    End If
                                End If

                                ' Else If its Not a AD nor an Article
                            Elseif currentColor = articleColor Or LCase(Left(sFile, 7)) = "article" Then
                                tagName = sTempFile

                                ' Get all page items in the document
                                ' Set pageItems = iDDoc.PageItems

                                ' Loop through all page items
                                For i = 1 To UBound(sortedPageItems)
                                    'Set pageItem = sortedPageItems(i).Item(i)

                                    Debug.Print ("checking If Element ID " & sortedPageItems(i).ID & " has Tag " & tagName)
                                    ' Check If the page item has an associated XML element
                                    On Error Resume Next
                                    Set xmlTag = sortedPageItems(i).AssociatedXMLElement
                                    On Error Goto 0


                                        If Not xmlTag Is Nothing Then
                                            ' Check If the XML tag name matches Tag Name
                                            If xmlTag.MarkupTag.Name = tagName Then

                                                tagExists = True


                                                ' Check If this is same tag name than the previous one
                                                If lastTagName = tagName Then
                                                    offsetPage = offsetPage ' Keep offset from previous one
                                                    oldPage = sortedPageItems(i).ParentPage.DocumentOffset
                                                    moveOnPage = False

                                                Else
                                                    ' Get the current page number (1-based in VBA)
                                                    oldPage = sortedPageItems(i).ParentPage.DocumentOffset

                                                    ' Define the offset , Current page - where it is -oldPage
                                                    offsetPage = X - oldPage

                                                    lastTagName = tagName ' Update the lastTagName
                                                    moveOnPage = False

                                                End If



                                                newPageNumber = oldPage + offsetPage
                                                'Debug.Print ("x is now on page " & X & " Old Page is " & oldPage & " So x-oldpage = " & offsetPage & ", New page number is " & newPageNumber)
                                                ' Ensure the New page number is valid
                                                If newPageNumber < 1 Or newPageNumber > iDDoc.Pages.count Then
                                                    MsgBox "Can't move further backwards Or beyond the document."
                                                 Exit Sub
                                                End If

                                                ' Get the target page based on the offset
                                                Set targetPage = iDDoc.Pages(newPageNumber)


                                                ' Calculate the New position based on current rectangle
                                                Dim bounds As Variant
                                                bounds = sortedPageItems(i).GeometricBounds
                                                Dim currentX As Double, currentY As Double
                                                currentX = bounds(1)
                                                currentY = bounds(0)

                                                If moveOnPage = True Then
                                                    ' Calculate the offset To move the items
                                                    Dim offsetX As Double, offsetY As Double
                                                    offsetX = myX1 - currentX
                                                    offsetY = myY1 - currentY

                                                Else

                                                    offsetX = currentX
                                                    offsetY = currentY

                                                End If

                                                ' Move the page item To the target page And restore its geometric bounds
                                                sortedPageItems(i).Move targetPage
                                                sortedPageItems(i).Move Array(offsetX, offsetY)

                                                Application.StatusBar = "Moving Around Elements"

                                                Debug.Print "Moved Element ID: " & sortedPageItems(i).ID; " With tagName " & tagName & " from page " & oldPage & " To page " & newPageNumber & " Page Item is " & i
                                            End If

                                        End If
                                    Next i
                                    If tagExists Then

                                        'Debug.Print "Tag '" & tagName & "' DOES exist."
                                    Else
                                        MsgBox "Tag '" & tagName & "' does Not exist."
                                     Exit Sub
                                    End If
                                End If


                            Next Y
                        Next X

                        ' Loop through all the pages in the document
                        For p = iDDoc.Pages.count To 1 Step -1 ' Loop backwards To avoid skipping pages
                            ' If the page index is greater than Or equal To deleteAfter
                            If p > lNoOfPages + 1 Then
                                iDDoc.Pages(p).Delete
                                Application.StatusBar = "Deleting extra Pages"

                            End If
                        Next p







                        'peneh-end
                        'vrati jedinice mere
                        iDDoc.ViewPreferences.HorizontalMeasurementUnits = idunitsHOR
                        iDDoc.ViewPreferences.VerticalMeasurementUnits = idunitsVER
                        'vrati kukanje za script
                        iDA.ScriptPreferences.UserInteractionLevel = idalerts
                        Application.StatusBar = "Saving InDesign file"
                        '    sIssue = ThisWorkbook.Worksheets("Main List").Range("CurrentIssue").value
                        sIssue = GetCurrentIssue
                        iDDoc.Save ThisWorkbook.Path & "\issue " & sIssue & "-" & Format(Now, "yy-mm-dd-hh-nn") & ".indd"
                        ' iDDoc.Windows.Add 'dodaje prozor, tj pokazuje skriveni dokument
                        'iDDoc.Close idSaveOptions.idNo
                        Application.StatusBar = "Done"
                        ThisWorkbook.Activate
                        MsgBox "Done For: " & sw.EndTimer / 1000 & " seconds."

End Sub
Sub NewBuildLayout()

    checkTags = True
    Set processedLayers = New Collection
    Set processedTags = New Collection

    If IDP Then Exit Sub
        Const StrokeWidth As Double = 1.38888888888889E-02
        Dim nName As Name
        Dim cPages As New Collection
        Dim cPUnits As New Collection
        Dim cPUnitsSizes As New Collection
        Dim cpUnitsPositions As New Collection
        Dim cUnits As Collection
        Dim cUnitsSizes As Collection
        Dim cUnitsColumns As Collection
        Dim cUnitsPositions As Collection
        Dim X As Long, Y As Long, z As Long
        Dim x2 As Long    'brojac za imena
        Dim r As Range, rc As Range
        Dim r2 As Range
        Dim aSize(1 To 7)    'row,col,Indesign width,height,filename,jel bled [onda 1, inace 0], ad name
        Dim aPos(1 To 2)    'top ,left
        Dim vc As Variant
        Dim lMaxPages As Long

        Dim iDA   ' As InDesign.Application
        Dim iDDoc   ' As InDesign.Document
        Dim iDPage    'As InDesign.Page
        Dim idColor    'As InDesign.Color
        Dim iDRectangle    ' As InDesign.Rectangle
        Dim myY1, myX1, myY2, myX2    'geometric bounds For rectangle
        Dim sFile As String
        Dim sTempFile As String    ' ako nema fajla [a upisano je nesto ko ad, ubaci To ko text]


        Dim lMinPage As Long
        Dim lNoOfPages As Long    'number of exported pages

        Dim bOddPage As Boolean    '?
        Dim bCS3fix As Boolean    'Do Not empty frames If cs3fix is active


        BuildNames ' added by Hussam

        If ThisWorkbook.Sheets("Settings").Range("CS3fix") = "Yes" Then bCS3fix = True

            If ThisWorkbook.Sheets("Settings").Range("IndesignTemplate").value = "" Then
                MsgBox "InDesign template is Not defined!"
             Exit Sub
            Elseif Dir(ThisWorkbook.Sheets("Settings").Range("IndesignTemplate").value) = "" Then
                MsgBox "InDesign template: " & ThisWorkbook.Sheets("Settings").Range("IndesignTemplate").value & " is Not found!"
             Exit Sub
            End If
            If DirExists("C:\regkey") = False Then
                MsgBox "Folder ""C:\regkey"" must be present!"
             Exit Sub
            End If

            Dim dHbleed As Double, dVbleed As Double    'za bleed
            dHbleed = val(ThisWorkbook.Sheets("Settings").Range("Bled_Size_Horizontal").value)
            dVbleed = val(ThisWorkbook.Sheets("Settings").Range("Bled_Size_Vertical").value)

            'ThisWorkbook.Sheets("Layout").Activate
            ThisWorkbook.Sheets("NewLayout").Activate
            If LCase(ThisWorkbook.Sheets("Settings").Range("Use_Export_Page_Settings").value) = "yes" Then
                lMinPage = ThisWorkbook.Sheets("Settings").Range("MinPageNo").value
                lMaxPages = ThisWorkbook.Sheets("Settings").Range("MaxNoPages").value
            Else
                lMinPage = 2
                lMaxPages = ThisWorkbook.Sheets("Settings").Range("B11").value
            End If

            'add dummy pages at 1, To keep counters sinhronized. It was needed when Yaakov asked For removing of first page
            Application.StatusBar = "Preparing..."
            cPages.Add 0&
            cPUnits.Add 0&
            cPUnitsSizes.Add 0&
            cpUnitsPositions.Add 0&


            lNoOfPages = lMaxPages - lMinPage + 1
            For x2 = lMinPage To lMaxPages
                '        SetPageRangeAddress (x2)
                '        Set NName = ThisWorkbook.Names("Page_" & x2)

                Set cUnits = New Collection
                Set cUnitsSizes = New Collection
                Set cUnitsPositions = New Collection
                '        tmpStr = NName.RefersTo
                '        vArr = Split(tmpStr, "!")
                '        Set r = ThisWorkbook.Sheets("NewLayout").Range(vArr(1))

                'Set r = NName.RefersToRange
                Set r = GetPageRange(x2)

                cPages.Add r
                For Each rc In r.Cells
                    If rc.MergeCells = True Then
                        On Error Resume Next
                        cUnits.Add rc.MergeArea, rc.MergeArea.Address
                        On Error Goto 0
                        Else
                            cUnits.Add rc
                        End If
                    Next rc
                    cPUnits.Add cUnits
                    For Each vc In cUnits
                        Set r2 = vc
                        aSize(1) = r2.rows.count
                        aSize(2) = r2.columns.count
                        aSize(3) = GetSizes(aSize(2), aSize(1))(1)
                        aSize(4) = GetSizes(aSize(2), aSize(1))(2)
                        'file name
                        aSize(5) = GetPathFromComment(r2)
                        'pokupi ime reklame
                        aSize(7) = r2.Cells(1).value

                        If IsBleed(x2) = True Then
                            If aSize(1) = 4 And aSize(2) = 2 Then    'full page
                                aSize(3) = GetBleedPageSizes()(1) + dHbleed + dHbleed
                                aSize(4) = GetBleedPageSizes()(2) + dVbleed + dVbleed
                            End If
                        End If
                        cUnitsSizes.Add aSize
                        aPos(1) = GetPositions(r2.Cells(1, 1), r)(1)
                        aPos(2) = GetPositions(r2.Cells(1, 1), r)(2)
                        If aSize(1) = 4 And aSize(2) = 2 Then    'full page
                            aPos(1) = ThisWorkbook.Sheets("Settings").Range("Full_Page_Position_1_Top").value
                            aPos(2) = ThisWorkbook.Sheets("Settings").Range("Full_Page_Position_1_Left").value
                        End If
                        If IsBleed(x2) = True Then
                            If aSize(1) = 4 And aSize(2) = 2 Then    'full page
                                '                    aPos(1) = aPos(1) - dVbleed
                                '                    aPos(2) = aPos(2) - dHbleed
                                aPos(1) = ThisWorkbook.Sheets("Settings").Range("f24").value
                                aPos(2) = ThisWorkbook.Sheets("Settings").Range("g24").value
                            End If
                        End If
                        cUnitsPositions.Add aPos
                    Next vc
                    cPUnitsSizes.Add cUnitsSizes
                    cpUnitsPositions.Add cUnitsPositions

                Next x2

                'id fun

                Dim sw As New StopWatch
                sw.StartTimer
                Application.StatusBar = "Starting InDesign"
                Set iDA = CreateObject("InDesign.Application")
                Dim idalerts    'da vratim nivo kukanja na pocetni nivo
                Dim idunitsHOR    'horizontalne jedinice
                Dim idunitsVER    'vertikalne jedinice

                idalerts = iDA.ScriptPreferences.UserInteractionLevel
                iDA.ScriptPreferences.UserInteractionLevel = 1699640946    'idUserInteractionLevels.idNeverInteract
                ' iDA.ScriptPreferences.EnableRedraw = False
                'iDA.Visible = True
                Set iDDoc = iDA.Open(Range("IndesignTemplate").value)
                'Set iDDoc = iDA.Open(Range("IndesignTemplate").Value, False) 'otvara dok skrivenim prozorom. oko duplo brze!
                iDA.Windows(1).Minimize    'ako je prozor skriven ovo pravi gresku!
                ThisWorkbook.Activate
                'Set jedinice mere
                idunitsHOR = iDDoc.ViewPreferences.HorizontalMeasurementUnits
                iDDoc.ViewPreferences.HorizontalMeasurementUnits = 2053729891    'idMeasurementUnits.idInches
                idunitsVER = iDDoc.ViewPreferences.HorizontalMeasurementUnits
                iDDoc.ViewPreferences.VerticalMeasurementUnits = 2053729891    'idMeasurementUnits.idInches
                ThisWorkbook.Activate

                Application.StatusBar = "Adding pages To InDesign"
                For X = lMinPage To lMaxPages
                    iDDoc.Pages.Add 1650945639    'idAtBegining
                    Application.StatusBar = "Adding pages To InDesign: " & X & " of " & lMaxPages
                Next X

                ' We'll need To create a color. Check To see If the color already exists.
                On Error Resume Next
                Set idColor = iDDoc.Colors.Item("YaakovBlack")
                If Error.Number <> 0 Then
                    Set idColor = iDDoc.Colors.Add
                    idColor.Name = "YaakovBlack"
                    idColor.Model = idColorModel.idProcess
                    idColor.ColorValue = Array(0, 0, 0, 100)
                    Error.Clear
                End If
                ' Resume normal error handling.
                On Error Goto 0

                    'hendle cell color
                    articleColor = ThisWorkbook.Sheets("Additional_settings").Range("b7").Interior.Color

                    If articleColor = 0 Then
                        MsgBox "The Article Color have no fill."

                    End If

                    ' Add a New layer To the document
                    LSPLayerName = "LSP"

                    layerExists = False
                    For i = 1 To iDDoc.Layers.count
                        If iDDoc.Layers.Item(i).Name = LSPLayerName Then
                            layerExists = True
                         Exit For
                        End If
                    Next i

                    ' Create the layer If it doesn't exist
                    If Not layerExists Then
                        Set LSPLayer = iDDoc.Layers.Add
                        LSPLayer.Name = LSPLayerName
                    Else
                        Set LSPLayer = iDDoc.Layers.Item(LSPLayerName)
                    End If


                    For X = 2 To lNoOfPages + 1
                        Application.StatusBar = "Creating boxes For page: " & X & " (of " & lNoOfPages & ")"
                        Set iDPage = iDDoc.Pages(X)

                        For Y = 1 To cPUnits(X).count
                            Set vc = cPUnits(X).Item(Y)
                            myY1 = cpUnitsPositions(X).Item(Y)(1)
                            myY2 = myY1 + cPUnitsSizes(X).Item(Y)(4)
                            myX1 = cpUnitsPositions(X).Item(Y)(2)
                            myX2 = myX1 + cPUnitsSizes(X).Item(Y)(3)
                            sFile = cPUnitsSizes(X).Item(Y)(5)
                            sTempFile = cPUnitsSizes(X).Item(Y)(7)

                            ' Get current cell color
                            currentColor = vc.Interior.Color


                            If currentColor <> articleColor Then 'If is Not an Article

                                ' create Rectangle And place the ad in it
                                Set iDRectangle = iDPage.Rectangles.Add(LSPLayer)
                                iDRectangle.GeometricBounds = Array(CDbl(myY1), CDbl(myX1), CDbl(myY2), CDbl(myX2))

                                'Set Rectangle stroke
                                If ThisWorkbook.Worksheets("Settings").Range("BlackBorder").value > 0 Then
                                    iDRectangle.StrokeWeight = ThisWorkbook.Worksheets("Settings").Range("BlackBorder").value
                                    iDRectangle.StrokeColor = iDDoc.Swatches.Item("YaakovBlack")
                                    iDRectangle.StrokeAlignment = 1936998729 'idInsideAlignment
                                Else
                                    iDRectangle.StrokeWeight = 0
                                End If

                                'Set the source file
                                If Len(sFile) > 0 And DirU(sFile) <> "" Then

                                    If getExt(sFile) <> "indd" Then
                                        iDRectangle.place sFile
                                        iDRectangle.Fit 1668575078 'idFitOptions.idContentToFrame
                                    Else
                                        iDRectangle.place sFile, True
                                        iDRectangle.Fit 1668575078 'idFitOptions.idContentToFrame
                                    End If
                                Else
                                    'Set options For frame To fit when place file manually

                                    If bCS3fix = False Then
                                        iDRectangle.FrameFittingOptions.AutoFit = True
                                        iDRectangle.FrameFittingOptions.FittingOnEmptyFrame = 1668575078    'idEmptyFrameFittingOptions.idContentToFrame
                                    End If

                                    If Len(sFile) > 0 And DirU(sFile) = "" Then
                                        If Range("WriteText").value = "Yes" Then
                                            ' Str2TXT CStr(sFile), ThisWorkbook.Path & "\mfile.txt"
                                            Str2TXT CStr(sFile), "C:\regkey\mfile.txt"
                                            'iDRectangle.Place ThisWorkbook.Path & "\mfile.txt"
                                            iDRectangle.place "C:\regkey\mfile.txt"
                                        End If
                                    Elseif Len(sFile) = 0 And Len(sTempFile) > 0 Then
                                        If Range("WriteText").value = "Yes" Then
                                            ' Str2TXT CStr(sFile), ThisWorkbook.Path & "\mfile.txt"
                                            Str2TXT CStr(sTempFile), "C:\regkey\mfile.txt"
                                            'iDRectangle.Place ThisWorkbook.Path & "\mfile.txt"
                                            iDRectangle.place "C:\regkey\mfile.txt"
                                        End If
                                    End If
                                End If

                            Elseif currentColor = articleColor Then ' Else If its Not a AD nor an Article
                                '************************** LAYER BEGIN ***********************************

                                elementName = sTempFile


                                ' Check If layer has already been processed
                                LayerProcessed = False ' Skip the rest of the processing For this layer


                                On Error Resume Next
                                processedLayers.Add elementName, elementName
                                If Err.Number = 457 Then ' Layer already processed (duplicate key)
                                    Err.Clear
                                    Debug.Print "Skipping layer: " & elementName
                                    LayerProcessed = True ' Skip the rest of the processing For this layer
                                End If

                                layerExists = False

                                Debug.Print ("checking If File has Layer " & elementName)

                                'checking For matching layer name
                                For i = 1 To iDDoc.Layers.count
                                    If iDDoc.Layers.Item(i).Name = elementName Then
                                        Set layer = iDDoc.Layers.Item(i)
                                        layerExists = True
                                     Exit For
                                    End If
                                Next i

                                If layerExists And LayerProcessed = False Then

                                    Application.StatusBar = "Moving layer " & elementName

                                    Set layerItems = layer.pageItems

                                    layerGroup = False

                                    If layerItems.count > 1 Then
                                        layerGroup = True
                                    End If
                                    ' Redefine the array To hold all page items
                                    ReDim sortedLayerItems(1 To layerItems.count)

                                    ' Populate the array With page items
                                    For i = 1 To layerItems.count
                                        Set sortedLayerItems(i) = layerItems(i)
                                    Next i

                                    ' Sort the array by page item id 
                                    For i = 1 To UBound(sortedLayerItems) - 1
                                        For J = i + 1 To UBound(sortedLayerItems)
                                            If sortedLayerItems(i).ID > sortedLayerItems(J).ID Then
                                                ' Swap the items
                                                Set Temp = sortedLayerItems(i)
                                                Set sortedLayerItems(i) = sortedLayerItems(J)
                                                Set sortedLayerItems(J) = Temp
                                            End If
                                        Next J
                                    Next i

                                    ' Move all items in the matched layer To the New location
                                    For J = 1 To UBound(sortedLayerItems)
                                        Set pageItem = sortedLayerItems(J)

                                        ' Calculate the New position based on current position
                                        Dim bounds As Variant
                                        bounds = pageItem.GeometricBounds
                                        Dim currentX As Double, currentY As Double
                                        currentX = bounds(1)
                                        currentY = bounds(0)

                                        ' Calculate the offset To move the items
                                        Dim offsetX As Double, offsetY As Double

                                        oldPage = pageItem.ParentPage.DocumentOffset ' Get the current page number (1-based in VB

                                        If lastelementName = elementName Then ' If its the second With the same layer
                                            offsetPage = offsetPage ' Keep offset from previous one
                                        Else
                                            offsetPage = X - oldPage ' Define the offset , Current page - where it is -oldPage

                                            lastelementName = elementName ' Update the lastelementName
                                        End If

                                        If layerGroup = True Then

                                            layerMoveOnPage = False
                                        Else
                                            layerMoveOnPage = True
                                        End If

                                        ' Move the page item by the calculated offset
                                        newPageNumber = oldPage + offsetPage


                                        If newPageNumber < 1 Or newPageNumber > iDDoc.Pages.count Then
                                            MsgBox "Can't move further backwards Or beyond the document."
                                         Exit Sub
                                        End If

                                        Set targetPage = iDDoc.Pages(newPageNumber)

                                        If layerMoveOnPage = True Then
                                            ' Calculate the offset To move the items
                                            offsetX = myX1 - currentX
                                            offsetY = myY1 - currentY

                                        Else

                                            offsetX = currentX
                                            offsetY = currentY

                                        End If

                                        pageItem.Move targetPage
                                        pageItem.Move Array(offsetX, offsetY)
                                        Debug.Print "Moved Page Item: " & pageItem.ID & " within Layer " & elementName & " from page " & oldPage & " To page " & newPageNumber

                                    Next J
                                Elseif checkTags = True And LayerProcessed = False Then 'If layer dosnt exsist, chck For xml Tag

                                    Debug.Print ("Layer " & elementName & " was Not found")

                                    '************************** LAYER END | XML TAG BEGIN ***********************************

                                    ' Check If layer has already been processed
                                    TagProcessed = False ' Skip the rest of the processing For this layer


                                    On Error Resume Next
                                    processedTags.Add elementName, elementName
                                    If Err.Number = 457 Then ' Layer already processed (duplicate key)
                                        Err.Clear
                                        Debug.Print "Skipping layer: " & elementName
                                        TagProcessed = True ' Skip the rest of the processing For this layer
                                    End If


                                    ' Get all page items in the document
                                    Set pageItems = iDDoc.pageItems

                                    ' Redefine the array To hold all page items
                                    ReDim sortedPageItems(1 To pageItems.count)

                                    ' Populate the array With page items
                                    For i = 1 To pageItems.count
                                        Set sortedPageItems(i) = pageItems(i)
                                    Next i

                                    ' Sort the array by page item id using Bubble Sort (Or another sort method)
                                    For i = 1 To UBound(sortedPageItems) - 1
                                        For J = i + 1 To UBound(sortedPageItems)
                                            If sortedPageItems(i).ID > sortedPageItems(J).ID Then
                                                ' Swap the items
                                                Set Temp = sortedPageItems(i)
                                                Set sortedPageItems(i) = sortedPageItems(J)
                                                Set sortedPageItems(J) = Temp
                                            End If
                                        Next J
                                    Next i
                                    Debug.Print ("checking If Tag " & elementName & " exsist")


                                    ' Loop through all page items
                                    For i = 1 To UBound(sortedPageItems)
                                        'Set pageItem = sortedPageItems(i).Item(i)

                                        ' Check If the page item has an associated XML element
                                        On Error Resume Next
                                        Set xmlTag = sortedPageItems(i).AssociatedXMLElement

                                        On Error Goto 0

                                            If Not xmlTag Is Nothing And TagProcessed = False Then

                                                ' Check If the XML tag name matches Tag Name
                                                If xmlTag.MarkupTag.Name = elementName Then

                                                    Application.StatusBar = "Moving tag " & elementName

                                                    tagExists = True


                                                    ' Check If this is same tag name than the previous one
                                                    If lastelementName = elementName Then
                                                        offsetPage = offsetPage ' Keep offset from previous one
                                                        oldPage = sortedPageItems(i).ParentPage.DocumentOffset
                                                        tagMoveOnPage = False

                                                    Else
                                                        ' Get the current page number (1-based in VBA)
                                                        oldPage = sortedPageItems(i).ParentPage.DocumentOffset

                                                        ' Define the offset , Current page - where it is -oldPage
                                                        offsetPage = X - oldPage

                                                        lastelementName = elementName ' Update the lastelementName
                                                        tagMoveOnPage = False

                                                    End If



                                                    newPageNumber = oldPage + offsetPage
                                                    'Debug.Print ("x is now on page " & X & " Old Page is " & oldPage & " So x-oldpage = " & offsetPage & ", New page number is " & newPageNumber)
                                                    ' Ensure the New page number is valid
                                                    If newPageNumber < 1 Or newPageNumber > iDDoc.Pages.count Then
                                                        MsgBox "Can't move further backwards Or beyond the document."
                                                     Exit Sub
                                                    End If

                                                    ' Get the target page based on the offset
                                                    Set targetPage = iDDoc.Pages(newPageNumber)


                                                    ' Calculate the New position based on current rectangle
                                                    bounds = sortedPageItems(i).GeometricBounds
                                                    currentX = bounds(1)
                                                    currentY = bounds(0)

                                                    If tagMoveOnPage = True Then
                                                        ' Calculate the offset To move the items
                                                        offsetX = myX1 - currentX
                                                        offsetY = myY1 - currentY

                                                    Else

                                                        offsetX = currentX
                                                        offsetY = currentY

                                                    End If

                                                    ' Move the page item To the target page And restore its geometric bounds
                                                    sortedPageItems(i).Move targetPage
                                                    sortedPageItems(i).Move Array(offsetX, offsetY)


                                                    Debug.Print "Moved Element ID: " & sortedPageItems(i).ID; " With Tag Name " & elementName & " from page " & oldPage & " To page " & newPageNumber & " Page Item is " & i
                                                End If 'XML tag name matches Tag Name

                                            End If 'If Not xmlTag Is Nothing


                                        Next i

                                        '************************** XML TAG END ***********************************

                                    End If ' If layerExists
                                    If tagExists = True Or layerExists = True Then

                                        'Debug.Print "Tag '" & elementName & "' DOES exist."
                                    Else
                                        MsgBox "Tag Or Layer '" & elementName & "' does Not exist."
                                        'Exit Sub
                                    End If
                                End If 'If is Not an Article

                            Next Y
                        Next X

                        ' Loop through all the pages in the document
                        For p = iDDoc.Pages.count To 1 Step -1 ' Loop backwards To avoid skipping pages
                            ' If the page index is greater than Or equal To deleteAfter
                            If p > lNoOfPages + 1 Then
                                iDDoc.Pages(p).Delete
                            End If
                        Next p

                        'vrati jedinice mere
                        iDDoc.ViewPreferences.HorizontalMeasurementUnits = idunitsHOR
                        iDDoc.ViewPreferences.VerticalMeasurementUnits = idunitsVER
                        'vrati kukanje za script
                        iDA.ScriptPreferences.UserInteractionLevel = idalerts
                        Application.StatusBar = "Saving InDesign file"
                        '    sIssue = ThisWorkbook.Worksheets("Main List").Range("CurrentIssue").value
                        sIssue = GetCurrentIssue
                        iDDoc.Save ThisWorkbook.Path & "\issue " & sIssue & "-" & Format(Now, "yy-mm-dd-hh-nn") & ".indd"


                        ' iDDoc.Windows.Add 'dodaje prozor, tj pokazuje skriveni dokument
                        'iDDoc.Close idSaveOptions.idNo
                        Application.StatusBar = "Done"
                        ThisWorkbook.Activate
                        MsgBox "Done For: " & sw.EndTimer / 1000 & " seconds."

End Sub
Function GetPositionsRedni(rSout As Range, rSearchIn As Range) As Long 'vraca od 1 Do 8
    Dim ap(1 To 2)
    Dim X As Long
    Dim r As Range, rc As Range
    Set r = Range("PositionForID")
    Dim lPos As Long
    For X = 1 To rSearchIn.Cells.count
        Set rc = rSearchIn.Cells(X)
        If rc.Address = rSout.Address Then
            lPos = X
        End If
    Next X
    '    For Each rc In r.Cells
    '        If rc.Value = lPos Then
    '            ap(1) = rc.Offset(0, 1).Value
    '            ap(2) = rc.Offset(0, 2).Value
    '            Exit For
    '        End If
    '    Next rc

    GetPositionsRedni = lPos
End Function
Private Function GetPositions(rSout As Range, rSearchIn As Range) As Variant()
    Dim ap(1 To 2)
    Dim X As Long
    Dim r As Range, rc As Range
    Set r = ThisWorkbook.Worksheets("Settings").Range("PositionForID")
    Dim lPos As Long
    For X = 1 To rSearchIn.Cells.count
        Set rc = rSearchIn.Cells(X)
        If rc.Address = rSout.Address Then
            lPos = X
        End If
    Next X
    For Each rc In r.Cells
        If rc.value = lPos Then
            ap(1) = rc.Offset(0, 1).value
            ap(2) = rc.Offset(0, 2).value
         Exit For
        End If
    Next rc

    GetPositions = ap
End Function
Function GetSizesNames(lCols, lRows) As String
    Dim ad(1 To 2)
    Dim r As Range, rc As Range
    Set r = ThisWorkbook.Worksheets("Settings").Range("Type_Size")
    For Each rc In r.Cells
        If rc.Offset(0, 1).value = lCols And rc.Offset(0, 2).value = lRows Then
            ad(1) = rc.Offset(0, 5).value

         Exit For
        End If
    Next rc
    GetSizesNames = ad(1)
End Function
Private Function GetSizes(lCols, lRows) As Variant()
    Dim ad(1 To 2)
    Dim r As Range, rc As Range
    Set r = Range("Type_Size")
    For Each rc In r.Cells
        If rc.Offset(0, 1).value = lCols And rc.Offset(0, 2).value = lRows Then
            ad(1) = rc.Offset(0, 3).value
            ad(2) = rc.Offset(0, 4).value
         Exit For
        End If
    Next rc
    GetSizes = ad


End Function
Private Function GetBleedPageSizes() As Variant()
    Dim ad(1 To 2)
    Dim r As Range, rc As Range
    Set r = Range("Type_Size")
    For Each rc In r.Cells
        If rc.Offset(0, 5).value = "FB" Then
            ad(1) = rc.Offset(0, 3).value
            ad(2) = rc.Offset(0, 4).value
         Exit For
        End If
    Next rc
    GetBleedPageSizes = ad

End Function

Function GetFirstFile(Byval sName As String, Optional bColor As Boolean = False) As String
    Dim sTemp As String
    Dim sSuffix As String
    Dim sBaseFolder As String

    sSuffix = " -col"
    sBaseFolder = Range("Base_folder").value
    sBaseFolder = TrailingSlash(sBaseFolder)
    If Len(Trim(sBaseFolder)) < 3 Then
        GetFirstFile = ""
    End If
    Dim X As Long
    Dim aFN()
    aFN = Array(".jpg", ".tif", ".psd", ".pdf")

    If bColor = False Then
        For X = LBound(aFN) To UBound(aFN)
            sTemp = sBaseFolder & sName & aFN(X)
            If DirU(sTemp) <> "" Then
                GetFirstFile = sTemp
             Exit Function
            End If
        Next X
    Else
        For X = LBound(aFN) To UBound(aFN)
            sTemp = sBaseFolder & sName & sSuffix & aFN(X)
            If DirU(sTemp) <> "" Then
                GetFirstFile = sTemp
             Exit Function
            End If
        Next X
    End If

End Function

Public Sub MarkLastPage()
    Dim lPages As Long
    Dim r As Range, r2 As Range, r3 As Range, R4 As Range
    '    If Range("MaxNoPages").Value = "" Then
    '        Exit Sub
    '    End If
    '    lPages = Range("MaxNoPages").Value

    lPages = ThisWorkbook.Sheets("Settings").Range("B11").value

    '    If lPages < 2 Or lPages > UBound(PagesArr) Then
    '        lPages = UBound(PagesArr)
    '    End If
    On Error Goto skipEOB
        With Range("EndOfBook").Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
 skipEOB:
        'Set r = ThisWorkbook.Names("Page_" & lPages).RefersToRange
        On Error Goto 0
            Set ws = ThisWorkbook.Sheets("NewLayout")
            '
            '    PageR = PagesStartRow + ((lPages - 2) \ 8) * RowsPerPage
            '    PageC = startCol + ((lPages - 2) Mod 8) * columnsPerPage
            '    Set r = ws.Range(ws.Cells(PageR + 1, PageC), ws.Cells(PageR + RowsPerPage - 2, PageC + 1))

            Set r = GetPageRange(lPages)

            'Debug.Print r.columns.Count
            Set r2 = r.Cells(4, 2).Offset(0, 1)
            Set r3 = r2.Offset(-3, 0)
            Set R4 = Range(r2.Cells(1, 1), r3.Cells(1, 1))
            ThisWorkbook.Names("EndOfBook").RefersTo = "=NewLayout!" & R4.Address
            With R4.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = vbBlue
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
End Sub

