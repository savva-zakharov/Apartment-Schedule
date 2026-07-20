Option Explicit

' ============================================================================
' UNIT SHORT MODULE
' Generates the condensed "Short" schedule (level/block/whole-scheme totals
' only, no per-unit rows) from wsSource data. Bed count breakdowns are split
' by dwelling type, e.g. "2 Bed Apartment", "2 Bed House" and "2 Bed Duplex"
' are tallied as separate columns.
'
' The Short schedule's own column layout is a custom table defined by
' template row 18 (mirroring how row 9 defines the Long schedule's layout -
' see Unit Schedule.bas). Any field named there (NO, GIFA, BEDS, etc.) is
' placed in wsShort at that column; fields not listed are simply omitted.
' The bed/dwelling-type tally columns are inserted at the "MIX" column,
' expanding it with extra columns as needed to fit every combination found.
'
' Reuses the generic setup routines from Unit Schedule.bas (ImportData,
' SortSchedule, DeleteInvalidRows, ApplyDwellingStandards,
' AddTenPercentIndicator), the dwelling-type classifier GetUnitTitle from
' Unit Stats.bas, and the shared helpers in Common Utilities.bas.
' ============================================================================

Sub GenerateUnitShort()
    Dim wsSource As Worksheet
    Dim wsWork As Worksheet
    Dim wsShort As Worksheet
    Dim wsTemplate As Worksheet
    Dim headerMap As Object
    Dim lastRow As Long
    Dim lastCol As Long
    Dim filePath As String
    Dim currentDate As String
    Dim ws As Worksheet

    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' Delete existing Short sheets (including any leftover working sheet from a failed run)
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, "Short", vbTextCompare) > 0 Then
            ws.Delete
        End If
    Next ws

    ' Set worksheets
    Set wsSource = ThisWorkbook.Sheets("sourceData")
    Set wsTemplate = ThisWorkbook.Sheets("template")

    ' Import data from file
    filePath = Trim(wsTemplate.Range("AA5").Value)
    If Left(filePath, 1) = """" And Right(filePath, 1) = """" Then
        filePath = Mid(filePath, 2, Len(filePath) - 2)
    End If

    If Dir(filePath) = "" Or Dir(filePath) = "NA" Then
        MsgBox "File not found:" & vbCrLf & filePath, vbExclamation
        GoTo Cleanup
    End If

    Call ImportData(wsSource, filePath)

    currentDate = Format(Date, "yy-mm-dd")

    ' Temporary working sheet: sorted, standards-applied per-unit data.
    ' Deleted before this sub exits - only wsShort is left behind.
    Set wsWork = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsWork.Name = "ShortWork_tmp"

    ' Build header map and copy source columns into the working sheet
    Set headerMap = BuildHeaderMap(wsTemplate, 9)
    Call CopyColumnsByHeader(wsSource, wsWork, wsTemplate, 1, 9)

    lastCol = GetLastColumnFromHeaderMap(headerMap)

    ' Sort by ZONE > BLOK > LEVL > NO
    Call SortSchedule(wsWork, headerMap, lastCol)

    ' Delete rows with XX in BLOK/LEVL/ZONE/NO
    Call DeleteInvalidRows(wsWork, headerMap)

    ' Apply dwelling lookup colours and minimum standards
    Call ApplyDwellingStandards(wsWork, wsTemplate, headerMap)

    ' Add 10% area indicator
    Call AddTenPercentIndicator(wsWork, headerMap)

    ' Create the Short worksheet
    Set wsShort = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsShort.Name = "Short " & currentDate

    ' The Short schedule's own column layout - a custom table defined by row 18,
    ' the same way row 9 defines the Long schedule's layout
    Dim headerMapShort As Object
    Set headerMapShort = BuildHeaderMap(wsTemplate, 18)

    ' Build the condensed level/block/whole-scheme summary
    Dim shortLastCol As Long
    Dim tallyLabels As Collection ' Array(column, label, reference colour) per tallied bed/type combo
    Set tallyLabels = New Collection
    shortLastCol = BuildShortSummary(wsWork, wsShort, headerMap, headerMapShort, lastCol, tallyLabels)

    ' Discard the working sheet
    Application.DisplayAlerts = False
    wsWork.Delete
    Application.DisplayAlerts = True

    ' Copy headers from template
    wsTemplate.Range("A10:N17").Copy
    wsShort.Range("A1").Insert Shift:=xlDown
    wsTemplate.Range("BA1:BR8").Copy
    wsShort.Range("S1").Insert Shift:=xlDown

    ' Before labelling the mix columns, shift everything in rows 7 & 8 to the
    ' right of MIX out of the way, to make room for the extra mix columns
    If tallyLabels.Count > 1 Then
        Dim mixStartCol As Long, mixEndCol As Long, shiftBy As Long, headerLastCol As Long
        mixStartCol = tallyLabels(1)(0)
        shiftBy = tallyLabels.Count - 1
        mixEndCol = mixStartCol + shiftBy

        headerLastCol = Application.WorksheetFunction.Max( _
            wsShort.Cells(7, wsShort.columns.Count).End(xlToLeft).Column, _
            wsShort.Cells(8, wsShort.columns.Count).End(xlToLeft).Column)

        If headerLastCol > mixStartCol Then
            wsShort.Range(wsShort.Cells(7, mixStartCol + 1), wsShort.Cells(8, headerLastCol)).Cut _
                Destination:=wsShort.Cells(7, mixStartCol + 1 + shiftBy)
        End If

        ' Fill the gap: copy the MIX column's formatting into every newly
        ' vacated column, then merge and centre row 7 across the whole mix span
        wsShort.Range(wsShort.Cells(7, mixStartCol), wsShort.Cells(8, mixStartCol)).Copy
        wsShort.Range(wsShort.Cells(7, mixStartCol + 1), wsShort.Cells(8, mixEndCol)).PasteSpecial Paste:=xlPasteFormats

        With wsShort.Range(wsShort.Cells(7, mixStartCol), wsShort.Cells(7, mixEndCol))
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        Application.CutCopyMode = False
    End If

    ' Label the tallied bed/type columns in the header row, coloured to match
    ' the fill used for that dwelling type in the schedule itself
    Dim tallyEntry As Variant
    For Each tallyEntry In tallyLabels
        With wsShort.Cells(8, tallyEntry(0))
            .Value = tallyEntry(1)
            .Interior.Color = tallyEntry(2)
            .Font.Bold = True
            .WrapText = True
        End With
    Next tallyEntry

    ' Add timestamp
    wsShort.Range("E5").Value = FormatDateWithSuffix(Date)

    ' Set print area
    lastRow = wsShort.Cells(wsShort.rows.Count, "B").End(xlUp).row
    wsShort.PageSetup.PrintArea = "A1:" & ColumnToLetter(shortLastCol) & (lastRow + 2)

    ' Activate and show print preview
    wsShort.Activate
    Application.CutCopyMode = False
    ActiveWindow.View = xlPageBreakPreview
    With wsShort.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$7:$9"
    End With
    wsShort.Range("A1").Select

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

End Sub

' ============================================================================
' Build the condensed Short schedule: one row per level, one totals row per
' block, and a whole-scheme totals row. No per-unit detail rows are written.
' headerMapShort (template row 18) defines wsShort's own custom column
' layout, decoupled from headerMap (row 9), which describes wsWork's layout.
' Returns the right-most column used, for use in borders and the print area.
' ============================================================================
Function BuildShortSummary(wsWork As Worksheet, wsShort As Worksheet, _
                     headerMap As Object, headerMapShort As Object, lastCol As Long, _
                     ByRef tallyLabels As Collection) As Long

    Dim lastRowWork As Long
    Dim i As Long
    Dim iShort As Long
    Dim c As Long
    Dim currentLevel As Variant, previousLevel As Variant
    Dim currentBlock As Variant, previousBlock As Variant
    Dim currentZone As Variant, previousZone As Variant
    Dim levelStartRow As Long, shortBlockStartRow As Long
    Dim hasZone As Boolean, hasBlock As Boolean, hasLevel As Boolean
    Dim levelChanged As Boolean, blockChanged As Boolean, zoneChanged As Boolean
    Dim statCols As Collection        ' Array(wsWork read column, wsShort write column, isNoField)
    Dim workTallyCols As Collection   ' bed count + dwelling type tally columns in wsWork (scratch positions)
    Dim shortTallyCols As Collection  ' bed count + dwelling type tally columns in wsShort (at/after the MIX column)
    Dim percentCalcColumns As Collection
    Dim shortChangeBlock As Collection
    Dim groupDict As Object          ' "bedCount|DwellingType" -> Array(bedCount, DwellingType, wsWork row)
    Dim groupKeys As Variant, bVal As Variant
    Dim dwellingType As String, groupKey As String
    Dim b1 As Long, b2 As Long, tempKey As Variant
    Dim workTallyStartCol As Long
    Dim mixCol As Long, numGroups As Long
    Dim re1 As Object
    Dim blockTitle As String
    Dim shortLastCol As Long

    hasZone = headerMap.Exists("ZONE")
    hasBlock = headerMap.Exists("BLOK")
    hasLevel = headerMap.Exists("LEVL")

    lastRowWork = wsWork.Cells(wsWork.rows.Count, 1).End(xlUp).row

    ' --- Find unique (bedroom count, dwelling type) combinations ---
    ' e.g. "2 Bed Apartment", "2 Bed House" and "2 Bed Duplex" are tallied separately.
    Set groupDict = CreateObject("Scripting.Dictionary")

    If headerMap.Exists("BEDS") Then
        For i = 2 To lastRowWork
            bVal = wsWork.Cells(i, GetColByHeader(headerMap, "BEDS")).Value
            If IsNumeric(bVal) And Len(bVal) > 0 Then
                bVal = CDbl(bVal)
                dwellingType = GetUnitTitle(wsWork, i, headerMap)
                groupKey = bVal & "|" & dwellingType
                If Not groupDict.Exists(groupKey) Then groupDict.Add groupKey, Array(bVal, dwellingType, i)
            End If
        Next i
    End If

    ' Sort groups by bedroom count, then dwelling type, ascending
    groupKeys = groupDict.Keys
    For b1 = LBound(groupKeys) To UBound(groupKeys) - 1
        For b2 = b1 + 1 To UBound(groupKeys)
            Dim itemA As Variant, itemB As Variant
            itemA = groupDict(groupKeys(b1))
            itemB = groupDict(groupKeys(b2))
            If itemA(0) > itemB(0) Or (itemA(0) = itemB(0) And itemA(1) > itemB(1)) Then
                tempKey = groupKeys(b1)
                groupKeys(b1) = groupKeys(b2)
                groupKeys(b2) = tempKey
            End If
        Next b2
    Next b1
    numGroups = UBound(groupKeys) - LBound(groupKeys) + 1

    ' --- Place the tally columns at the "MIX" column, expanding it with extra
    ' columns as needed to fit every (bed count, dwelling type) combination ---
    If headerMapShort.Exists("MIX") Then
        mixCol = GetColByHeader(headerMapShort, "MIX")
        If numGroups > 1 Then
            wsShort.columns(mixCol + 1).Resize(, numGroups - 1).Insert Shift:=xlToRight
            ' Shift any custom-table field positioned after MIX to account for the new columns
            Dim hKey As Variant
            For Each hKey In headerMapShort.Keys
                If headerMapShort(hKey) > mixCol Then
                    headerMapShort(hKey) = headerMapShort(hKey) + (numGroups - 1)
                End If
            Next hKey
        End If
    Else
        ' No MIX column defined in the custom table - fall back to placing
        ' tallies just past it so nothing gets overwritten
        mixCol = GetLastColumnFromHeaderMap(headerMapShort) + 2
    End If

    ' --- Main-stat columns: read from wsWork (row 9 layout), write to wsShort
    ' (row 18 custom-table layout) ---
    Dim statFields As Variant
    statFields = Array("NO", "GIFA", "MINAREA", "BEDS", "PERS", "DUAL", "MINPAS", "PAS", "MINCAS", "MIN10")
    Dim f As Long
    Dim statField As String
    Set statCols = New Collection
    For f = LBound(statFields) To UBound(statFields)
        statField = CStr(statFields(f))
        If headerMap.Exists(statField) And headerMapShort.Exists(statField) Then
            statCols.Add Array(GetColByHeader(headerMap, statField), _
                                GetColByHeader(headerMapShort, statField), _
                                statField = "NO")
        End If
    Next f

    Set percentCalcColumns = New Collection
    If headerMapShort.Exists("MIN10") Then percentCalcColumns.Add GetColByHeader(headerMapShort, "MIN10")
    If headerMapShort.Exists("DUAL") Then percentCalcColumns.Add GetColByHeader(headerMapShort, "DUAL")

    Set workTallyCols = New Collection
    Set shortTallyCols = New Collection
    Dim tally As Object ' "bedCount|DwellingType" -> wsWork tally column
    Set tally = CreateObject("Scripting.Dictionary")

    workTallyStartCol = lastCol + 2
    shortLastCol = GetLastColumnFromHeaderMap(headerMapShort)

    Dim grp As Variant
    Dim refColor As Long
    For b1 = LBound(groupKeys) To UBound(groupKeys)
        grp = groupDict(groupKeys(b1))
        tally.Add groupKeys(b1), workTallyStartCol + b1
        workTallyCols.Add workTallyStartCol + b1
        shortTallyCols.Add mixCol + b1
        percentCalcColumns.Add mixCol + b1
        refColor = wsWork.Cells(grp(2), 1).Interior.Color
        tallyLabels.Add Array(mixCol + b1, grp(0) & " Bed " & grp(1), refColor)
        If mixCol + b1 > shortLastCol Then shortLastCol = mixCol + b1
    Next b1

    ' Flag each unit row with a 1 in its (bed count, dwelling type) tally column
    If headerMap.Exists("BEDS") Then
        For i = 2 To lastRowWork
            bVal = wsWork.Cells(i, GetColByHeader(headerMap, "BEDS")).Value
            If IsNumeric(bVal) And Len(bVal) > 0 Then
                dwellingType = GetUnitTitle(wsWork, i, headerMap)
                groupKey = CDbl(bVal) & "|" & dwellingType
                If tally.Exists(groupKey) Then
                    wsWork.Cells(i, tally(groupKey)).Value = 1
                End If
            End If
        Next i
    End If

    ' Combined column list used when summing block/whole-scheme totals within wsShort
    Dim shortSumColumns As Collection
    Set shortSumColumns = New Collection
    For c = 1 To statCols.Count
        shortSumColumns.Add statCols(c)(1)
    Next c
    For c = 1 To shortTallyCols.Count
        shortSumColumns.Add shortTallyCols(c)
    Next c

    Set re1 = CreateObject("VBScript.RegExp")
    With re1
        .Pattern = "^[A-Za-z0-9]{1,2}$"
        .IgnoreCase = True
        .Global = False
    End With

    Set shortChangeBlock = New Collection

    Dim readCol As Long, writeCol As Long, isNoField As Boolean
    Dim workCol As Long, shortCol As Long

    If Not hasLevel And Not hasBlock And Not hasZone Then
        ' No grouping columns available - summarize the whole dataset as one group
        iShort = 2
        wsShort.Cells(iShort, "B").Value = "Units"
        For c = 1 To statCols.Count
            readCol = statCols(c)(0)
            writeCol = statCols(c)(1)
            isNoField = statCols(c)(2)
            wsShort.Cells(iShort, writeCol).Value = AggregateWorkRange(wsWork, readCol, 2, lastRowWork, isNoField)
        Next c
        For c = 1 To workTallyCols.Count
            workCol = workTallyCols(c)
            shortCol = shortTallyCols(c)
            wsShort.Cells(iShort, shortCol).Value = AggregateWorkRange(wsWork, workCol, 2, lastRowWork)
        Next c
        shortChangeBlock.Add iShort
        iShort = iShort + 1
    Else
        iShort = 2
        i = 2
        If hasLevel Then previousLevel = wsWork.Cells(2, GetColByHeader(headerMap, "LEVL")).Value
        If hasBlock Then previousBlock = wsWork.Cells(2, GetColByHeader(headerMap, "BLOK")).Value
        If hasZone Then previousZone = wsWork.Cells(2, GetColByHeader(headerMap, "ZONE")).Value
        levelStartRow = 2
        shortBlockStartRow = 2

        Do While True
            If hasLevel Then currentLevel = wsWork.Cells(i, GetColByHeader(headerMap, "LEVL")).Value
            If hasBlock Then currentBlock = wsWork.Cells(i, GetColByHeader(headerMap, "BLOK")).Value
            If hasZone Then currentZone = wsWork.Cells(i, GetColByHeader(headerMap, "ZONE")).Value

            levelChanged = (hasLevel And currentLevel <> previousLevel)
            blockChanged = (hasBlock And currentBlock <> previousBlock)
            zoneChanged = (hasZone And currentZone <> previousZone)

            If levelChanged Or blockChanged Or zoneChanged Then
                ' Write the level summary row
                wsShort.Cells(iShort, "B").Value = previousLevel
                For c = 1 To statCols.Count
                    readCol = statCols(c)(0)
                    writeCol = statCols(c)(1)
                    isNoField = statCols(c)(2)
                    wsShort.Cells(iShort, writeCol).Value = AggregateWorkRange(wsWork, readCol, levelStartRow, i - 1, isNoField)
                Next c
                For c = 1 To workTallyCols.Count
                    workCol = workTallyCols(c)
                    shortCol = shortTallyCols(c)
                    wsShort.Cells(iShort, shortCol).Value = AggregateWorkRange(wsWork, workCol, levelStartRow, i - 1)
                Next c
                iShort = iShort + 1

                If blockChanged Then
                    If hasBlock And re1.Test(CStr(previousBlock)) Then
                        blockTitle = "Block " & previousBlock & " Summary"
                    ElseIf hasBlock Then
                        blockTitle = previousBlock & " Summary"
                    Else
                        blockTitle = "Summary"
                    End If

                    With wsShort.Cells(shortBlockStartRow - 1, "B")
                        .Value = blockTitle
                        .Font.Bold = True
                        .Font.Color = RGB(0, 176, 240)
                        .Font.Name = "Calibri"
                        .HorizontalAlignment = xlLeft
                    End With

                    Call sumColumnsSub(wsShort, shortSumColumns, shortBlockStartRow, iShort, 0, False)
                    Call percentColumnsSub(wsShort, percentCalcColumns, iShort, 0)

                    Call drawBorderThickOutline(wsShort.Range( _
                        wsShort.Cells(shortBlockStartRow, "A"), _
                        wsShort.Cells(iShort - 1, shortLastCol)))

                    shortChangeBlock.Add iShort

                    iShort = iShort + 4 ' total row + percent row + 2 blank spacer rows before the next block
                    shortBlockStartRow = iShort

                    previousBlock = currentBlock
                End If

                levelStartRow = i
            End If

            If hasLevel Then previousLevel = currentLevel
            If hasZone Then previousZone = currentZone

            i = i + 1

            If wsWork.Cells(i - 1, 1).Value = 0 Or i > 100000 Then Exit Do
        Loop
    End If

    ' --- Whole scheme summary ---
    Call drawBorderLine(wsShort, iShort, shortLastCol)

    With wsShort.Cells(iShort - 1, "B")
        .Value = "Whole Scheme Summary"
        .Font.Bold = True
        .Font.Color = RGB(0, 176, 240)
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlLeft
    End With

    If shortChangeBlock.Count > 0 Then
        Call sumColumnsRowsSub(wsShort, shortSumColumns, shortChangeBlock, iShort)
    Else
        Call sumColumnsSub(wsShort, shortSumColumns, 2, iShort, 0, False)
    End If
    Call percentColumnsSub(wsShort, percentCalcColumns, iShort, 0)

    BuildShortSummary = shortLastCol

End Function

' ============================================================================
' Sum (or, for the unit-label NO column, count) a column across a contiguous
' row range on the working sheet. Returns 0 for an empty/invalid range
' instead of raising an error.
' ============================================================================
Function AggregateWorkRange(ws As Worksheet, col As Long, startRow As Long, endRow As Long, _
                           Optional useCount As Boolean = False) As Double
    If endRow < startRow Or col < 1 Then
        AggregateWorkRange = 0
    ElseIf useCount Then
        AggregateWorkRange = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(startRow, col), ws.Cells(endRow, col)))
    Else
        AggregateWorkRange = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, col), ws.Cells(endRow, col)))
    End If
End Function
