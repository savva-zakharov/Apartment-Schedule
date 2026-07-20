Option Explicit

' ============================================================================
' UNIT SHORT MODULE
' Generates the condensed "Short" schedule (level/block/whole-scheme totals
' only, no per-unit rows) from wsSource data. Bed count breakdowns are split
' by dwelling type, e.g. "2 Bed Apartment", "2 Bed House" and "2 Bed Duplex"
' are tallied as separate columns.
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

    ' Build the condensed level/block/whole-scheme summary
    Dim shortLastCol As Long
    shortLastCol = BuildShortSummary(wsWork, wsShort, headerMap, lastCol)

    ' Discard the working sheet
    Application.DisplayAlerts = False
    wsWork.Delete
    Application.DisplayAlerts = True

    ' Copy headers from template
    wsTemplate.Range("A10:N17").Copy
    wsShort.Range("A1").Insert Shift:=xlDown
    wsTemplate.Range("BA1:BR8").Copy
    wsShort.Range("S1").Insert Shift:=xlDown

    ' Add timestamp
    wsShort.Range("E5").Value = FormatDateWithSuffix(Date)

    ' Set print area
    lastRow = wsShort.Cells(wsShort.rows.Count, "B").End(xlUp).row
    wsShort.PageSetup.PrintArea = "A1:" & ColumnToLetter(shortLastCol) & lastRow

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
' Returns the right-most column used (main columns or bed tallies, whichever
' extends further), for use in borders and the print area.
' ============================================================================
Function BuildShortSummary(wsWork As Worksheet, wsShort As Worksheet, _
                     headerMap As Object, lastCol As Long) As Long

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
    Dim sumColumns As Collection      ' main-stat columns, shared position in wsWork and wsShort
    Dim workTallyCols As Collection   ' bed count + dwelling type tally columns in wsWork (scratch positions)
    Dim shortTallyCols As Collection  ' bed count + dwelling type tally columns in wsShort (fixed positions, matches original layout)
    Dim percentCalcColumns As Collection
    Dim shortChangeBlock As Collection
    Dim groupDict As Object          ' "bedCount|DwellingType" -> Array(bedCount, DwellingType)
    Dim groupKeys As Variant, bVal As Variant
    Dim dwellingType As String, groupKey As String
    Dim b1 As Long, b2 As Long, tempKey As Variant
    Dim workTallyStartCol As Long
    Const shortTallyStartCol As Long = 3 ' Column C, matches the original wsShort layout
    Dim re1 As Object
    Dim blockTitle As String
    Dim shortLastCol As Long
    Dim noCol As Long

    hasZone = headerMap.Exists("ZONE")
    hasBlock = headerMap.Exists("BLOK")
    hasLevel = headerMap.Exists("LEVL")

    lastRowWork = wsWork.Cells(wsWork.rows.Count, 1).End(xlUp).row

    ' --- Main-stat columns (identical positions in wsWork and wsShort) ---
    ' NO holds unit labels (not summable), so it is counted instead of summed
    Set sumColumns = New Collection
    If headerMap.Exists("NO") Then
        noCol = GetColByHeader(headerMap, "NO")
        sumColumns.Add noCol
    Else
        noCol = 0
    End If
    If headerMap.Exists("GIFA") Then sumColumns.Add GetColByHeader(headerMap, "GIFA")
    If headerMap.Exists("minAREA") Then sumColumns.Add GetColByHeader(headerMap, "MINAREA")
    If headerMap.Exists("BEDS") Then sumColumns.Add GetColByHeader(headerMap, "BEDS")
    If headerMap.Exists("PERS") Then sumColumns.Add GetColByHeader(headerMap, "PERS")
    If headerMap.Exists("DUAL") Then sumColumns.Add GetColByHeader(headerMap, "DUAL")
    If headerMap.Exists("minPAS") Then sumColumns.Add GetColByHeader(headerMap, "MINPAS")
    If headerMap.Exists("PAS") Then sumColumns.Add GetColByHeader(headerMap, "PAS")
    If headerMap.Exists("minCAS") Then sumColumns.Add GetColByHeader(headerMap, "MINCAS")
    If headerMap.Exists("min10") Then sumColumns.Add GetColByHeader(headerMap, "MIN10")

    Set percentCalcColumns = New Collection
    If headerMap.Exists("min10") Then percentCalcColumns.Add GetColByHeader(headerMap, "MIN10")
    If headerMap.Exists("DUAL") Then percentCalcColumns.Add GetColByHeader(headerMap, "DUAL")

    ' --- Find unique (bedroom count, dwelling type) combinations and set up tally columns ---
    ' e.g. "2 Bed Apartment", "2 Bed House" and "2 Bed Duplex" are tallied separately.
    Set groupDict = CreateObject("Scripting.Dictionary")

    If headerMap.Exists("BEDS") Then
        For i = 2 To lastRowWork
            bVal = wsWork.Cells(i, GetColByHeader(headerMap, "BEDS")).Value
            If IsNumeric(bVal) And Len(bVal) > 0 Then
                bVal = CDbl(bVal)
                dwellingType = GetUnitTitle(wsWork, i, headerMap)
                groupKey = bVal & "|" & dwellingType
                If Not groupDict.Exists(groupKey) Then groupDict.Add groupKey, Array(bVal, dwellingType)
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

    Set workTallyCols = New Collection
    Set shortTallyCols = New Collection
    Dim tally As Object ' "bedCount|DwellingType" -> wsWork tally column
    Set tally = CreateObject("Scripting.Dictionary")

    workTallyStartCol = lastCol + 2
    shortLastCol = lastCol

    Dim grp As Variant
    For b1 = LBound(groupKeys) To UBound(groupKeys)
        grp = groupDict(groupKeys(b1))
        tally.Add groupKeys(b1), workTallyStartCol + b1
        workTallyCols.Add workTallyStartCol + b1
        shortTallyCols.Add shortTallyStartCol + b1
        percentCalcColumns.Add shortTallyStartCol + b1
        wsShort.Cells(1, shortTallyStartCol + b1).Value = grp(0) & " BED " & UCase(grp(1))
        If shortTallyStartCol + b1 > shortLastCol Then shortLastCol = shortTallyStartCol + b1
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
    For c = 1 To sumColumns.Count
        shortSumColumns.Add sumColumns(c)
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

    If Not hasLevel And Not hasBlock And Not hasZone Then
        ' No grouping columns available - summarize the whole dataset as one group
        iShort = 2
        wsShort.Cells(iShort, "B").Value = "Units"
        For c = 1 To sumColumns.Count
            wsShort.Cells(iShort, sumColumns(c)).Value = AggregateWorkRange(wsWork, sumColumns(c), 2, lastRowWork, sumColumns(c) = noCol)
        Next c
        For c = 1 To workTallyCols.Count
            wsShort.Cells(iShort, shortTallyCols(c)).Value = AggregateWorkRange(wsWork, workTallyCols(c), 2, lastRowWork)
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
                For c = 1 To sumColumns.Count
                    wsShort.Cells(iShort, sumColumns(c)).Value = AggregateWorkRange(wsWork, sumColumns(c), levelStartRow, i - 1, sumColumns(c) = noCol)
                Next c
                For c = 1 To workTallyCols.Count
                    wsShort.Cells(iShort, shortTallyCols(c)).Value = AggregateWorkRange(wsWork, workTallyCols(c), levelStartRow, i - 1)
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
                        wsShort.Cells(iShort, shortLastCol)))

                    shortChangeBlock.Add iShort

                    iShort = iShort + 3
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
