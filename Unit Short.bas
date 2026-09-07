Option Explicit

' ============================================================================
' UNIT SHORT MODULE
' Generates the condensed "Short" schedule (level/block/whole-scheme totals
' only, no per-unit rows) from wsSource data. Bed count breakdowns are split
' by person count and dwelling type, e.g. "2B 3P Apartment" and "2B 4P
' Apartment" are tallied as separate columns.
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

    ' Paste the template header art onto the still-blank sheet FIRST, before
    ' BuildShortSummary makes any column inserts for the mix columns. A real
    ' Excel column insert shifts every row in the affected columns uniformly
    ' - including this header art - so it stays correctly aligned with the
    ' data without needing any separate row/column shifting logic afterwards.
    wsTemplate.Range("A10:AB17").Copy
    wsShort.Range("A1").PasteSpecial Paste:=xlPasteAll
    wsTemplate.Range("BA1:BR8").Copy
    wsShort.Range("S1").PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False

    ' The Short schedule's own column layout - a custom table defined by row 18,
    ' the same way row 9 defines the Long schedule's layout
    Dim headerMapShort As Object
    Set headerMapShort = BuildHeaderMap(wsTemplate, 18)

    ' Build the condensed level/block/whole-scheme summary
    Dim shortLastCol As Long
    Dim tallyLabels As Collection ' Array(column, label, reference colour) per tallied bed/type combo
    Dim mixZones As Collection    ' Array(start column, column count, zone name) per expanded mix zone
    Set tallyLabels = New Collection
    Set mixZones = New Collection
    shortLastCol = BuildShortSummary(wsWork, wsShort, wsTemplate, headerMap, headerMapShort, lastCol, tallyLabels, mixZones)

    ' Discard the working sheet
    Application.DisplayAlerts = False
    wsWork.Delete
    Application.DisplayAlerts = True

    ' Fill the gap left by each mix zone's expansion. The column insert made
    ' during BuildShortSummary already shifted the header art along with the
    ' data, so all that's left is: format the newly inserted (still blank)
    ' columns to match the zone's first column, and merge/centre row 7's title
    ' across the whole span.
    Dim zi As Long, zoneStartCol As Long, zoneShiftBy As Long, zoneEndCol As Long
    Dim zoneData As Variant
    For zi = 1 To mixZones.Count
        zoneData = mixZones(zi)
        zoneStartCol = zoneData(0)
        zoneShiftBy = zoneData(1) - 1
        zoneEndCol = zoneStartCol + zoneShiftBy

        wsShort.Range(wsShort.Cells(7, zoneStartCol), wsShort.Cells(8, zoneStartCol)).Copy
        wsShort.Range(wsShort.Cells(7, zoneStartCol + 1), wsShort.Cells(8, zoneEndCol)).PasteSpecial Paste:=xlPasteFormats

        With wsShort.Range(wsShort.Cells(7, zoneStartCol), wsShort.Cells(7, zoneEndCol))
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        ' The Total Mix zone gets its own title, replacing whatever the
        ' template originally had over the single TMIX column
        If zoneData(2) = "TMIX" Then
            With wsShort.Cells(7, zoneStartCol)
                .Value = "Total Mix"
            End With
        End If
    Next zi
    Application.CutCopyMode = False

    ' Label the tallied bed/type columns in the header row, coloured to match
    ' the fill used for that dwelling type in the schedule itself (Total Mix
    ' columns carry no dwelling-type colour, signalled by a negative value)
    Dim tallyEntry As Variant
    For Each tallyEntry In tallyLabels
        With wsShort.Cells(8, tallyEntry(0))
            .Value = tallyEntry(1)
            If tallyEntry(2) >= 0 Then .Interior.Color = tallyEntry(2)
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
Function BuildShortSummary(wsWork As Worksheet, wsShort As Worksheet, wsTemplate As Worksheet, _
                     headerMap As Object, headerMapShort As Object, lastCol As Long, _
                     ByRef tallyLabels As Collection, ByRef mixZones As Collection) As Long

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
    Dim workTotalTallyCols As Collection ' bed count only (Total Mix) tally columns in wsWork (scratch positions)
    Dim tmixTallyCols As Collection      ' bed count only (Total Mix) tally columns in wsShort (at/after the TMIX column)
    Dim percentCalcColumns As Collection
    Dim shortChangeBlock As Collection
    Dim groupDict As Object          ' "bedCount|persCount|DwellingType" -> Array(bedCount, persCount, DwellingType, wsWork row)
    Dim groupKeys As Variant, bVal As Variant, persVal As Double
    Dim totalDict As Object          ' bedCount -> Array(bedCount, wsWork row)
    Dim totalKeys As Variant
    Dim dwellingType As String, groupKey As String
    Dim b1 As Long, b2 As Long, tempKey As Variant
    Dim workTallyStartCol As Long
    Dim mixCol As Long, numGroups As Long
    Dim tmixCol As Long, numBedGroups As Long
    Dim mixExists As Boolean, tmixExists As Boolean
    Dim re1 As Object
    Dim blockTitle As String
    Dim shortLastCol As Long
    Dim headerMapShortRaw As Object  ' snapshot of headerMapShort before any MIX/TMIX column inserts
    Dim formatSourceMap As Object    ' wsShort column (final) -> wsTemplate row-18 column (raw), for formatting

    hasZone = headerMap.Exists("ZONE")
    hasBlock = headerMap.Exists("BLOK")
    hasLevel = headerMap.Exists("LEVL")

    lastRowWork = wsWork.Cells(wsWork.rows.Count, 1).End(xlUp).row

    ' Snapshot the custom table's column layout before it gets mutated by the
    ' MIX/TMIX column inserts below, so per-level rows can later copy each
    ' field's formatting from its original template cell to its final column.
    Set headerMapShortRaw = CreateObject("Scripting.Dictionary")
    Dim hKeySnap As Variant
    For Each hKeySnap In headerMapShort.Keys
        headerMapShortRaw.Add hKeySnap, headerMapShort(hKeySnap)
    Next hKeySnap

    ' --- Find unique (bedroom count, person count, dwelling type) combinations ---
    ' e.g. "2B 3P Apartment" and "2B 4P Apartment" are tallied as separate columns.
    Set groupDict = CreateObject("Scripting.Dictionary")

    If headerMap.Exists("BEDS") Then
        For i = 2 To lastRowWork
            bVal = wsWork.Cells(i, GetColByHeader(headerMap, "BEDS")).Value
            If IsNumeric(bVal) And Len(bVal) > 0 Then
                bVal = CDbl(bVal)
                dwellingType = GetUnitTitle(wsWork, i, headerMap)
                persVal = GetPersValue(wsWork, i, headerMap)
                groupKey = bVal & "|" & persVal & "|" & dwellingType
                If Not groupDict.Exists(groupKey) Then groupDict.Add groupKey, Array(bVal, persVal, dwellingType, i)
            End If
        Next i
    End If

    ' Sort groups by dwelling type (Apartment, then Duplex, then House, always
    ' in that order regardless of bedroom count), then by bedroom count, then
    ' by person count, all ascending
    groupKeys = groupDict.Keys
    For b1 = LBound(groupKeys) To UBound(groupKeys) - 1
        For b2 = b1 + 1 To UBound(groupKeys)
            Dim itemA As Variant, itemB As Variant
            Dim rankA As Long, rankB As Long
            Dim typeA As String, typeB As String
            itemA = groupDict(groupKeys(b1))
            itemB = groupDict(groupKeys(b2))
            typeA = CStr(itemA(2))
            typeB = CStr(itemB(2))
            rankA = DwellingTypeRank(typeA)
            rankB = DwellingTypeRank(typeB)
            If rankA > rankB _
            Or (rankA = rankB And itemA(0) > itemB(0)) _
            Or (rankA = rankB And itemA(0) = itemB(0) And itemA(1) > itemB(1)) Then
                tempKey = groupKeys(b1)
                groupKeys(b1) = groupKeys(b2)
                groupKeys(b2) = tempKey
            End If
        Next b2
    Next b1
    numGroups = UBound(groupKeys) - LBound(groupKeys) + 1

    ' --- Find unique bedroom counts for the Total Mix (bed count only,
    ' regardless of dwelling type) ---
    Set totalDict = CreateObject("Scripting.Dictionary")

    If headerMap.Exists("BEDS") Then
        For i = 2 To lastRowWork
            bVal = wsWork.Cells(i, GetColByHeader(headerMap, "BEDS")).Value
            If IsNumeric(bVal) And Len(bVal) > 0 Then
                bVal = CDbl(bVal)
                If Not totalDict.Exists(bVal) Then totalDict.Add bVal, Array(bVal, i)
            End If
        Next i
    End If

    ' Sort by bedroom count ascending
    totalKeys = totalDict.Keys
    For b1 = LBound(totalKeys) To UBound(totalKeys) - 1
        For b2 = b1 + 1 To UBound(totalKeys)
            If totalKeys(b1) > totalKeys(b2) Then
                tempKey = totalKeys(b1)
                totalKeys(b1) = totalKeys(b2)
                totalKeys(b2) = tempKey
            End If
        Next b2
    Next b1
    numBedGroups = UBound(totalKeys) - LBound(totalKeys) + 1

    ' --- Place the tally columns at the "MIX" column, expanding it with extra
    ' columns as needed to fit every (bed count, dwelling type) combination.
    ' Total Mix is optional: only created when a "TMIX" column is defined,
    ' expanding it the same way to fit every bedroom count found. ---
    mixExists = headerMapShort.Exists("MIX")
    tmixExists = headerMapShort.Exists("TMIX")

    If mixExists Then
        mixCol = GetColByHeader(headerMapShort, "MIX")
        If numGroups > 1 Then
            wsShort.columns(mixCol + 1).Resize(, numGroups - 1).Insert Shift:=xlToRight
            Call ShiftHeaderMapShortColumns(headerMapShort, mixCol, numGroups - 1)
        End If
    Else
        ' No MIX column defined in the custom table - fall back to placing
        ' tallies just past it so nothing gets overwritten
        mixCol = GetLastColumnFromHeaderMap(headerMapShort) + 2
    End If

    If tmixExists Then
        ' Re-fetch TMIX's position now, in case the MIX insert above shifted it
        tmixCol = GetColByHeader(headerMapShort, "TMIX")
        If numBedGroups > 1 Then
            wsShort.columns(tmixCol + 1).Resize(, numBedGroups - 1).Insert Shift:=xlToRight
            Call ShiftHeaderMapShortColumns(headerMapShort, tmixCol, numBedGroups - 1)
            ' TMIX may have sat to the left of MIX - re-resolve MIX in case it just shifted
            If mixExists Then mixCol = GetColByHeader(headerMapShort, "MIX")
        End If
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
        refColor = wsWork.Cells(grp(3), 1).Interior.Color
        tallyLabels.Add Array(mixCol + b1, grp(0) & "B " & grp(1) & "P " & grp(2), refColor)
        If mixCol + b1 > shortLastCol Then shortLastCol = mixCol + b1
    Next b1
    If numGroups > 1 Then mixZones.Add Array(mixCol, numGroups, "MIX")

    ' --- Total Mix tally columns: bed count only, regardless of dwelling type ---
    Set workTotalTallyCols = New Collection
    Set tmixTallyCols = New Collection
    Dim totalTally As Object ' bedCount -> wsWork tally column
    Set totalTally = CreateObject("Scripting.Dictionary")

    If tmixExists Then
        Dim workTotalTallyStartCol As Long
        Dim totalItem As Variant
        workTotalTallyStartCol = workTallyStartCol + numGroups

        For b1 = LBound(totalKeys) To UBound(totalKeys)
            totalItem = totalDict(totalKeys(b1))
            totalTally.Add totalKeys(b1), workTotalTallyStartCol + b1
            workTotalTallyCols.Add workTotalTallyStartCol + b1
            tmixTallyCols.Add tmixCol + b1
            percentCalcColumns.Add tmixCol + b1
            tallyLabels.Add Array(tmixCol + b1, totalItem(0) & " Bed Total", -1)
            If tmixCol + b1 > shortLastCol Then shortLastCol = tmixCol + b1
        Next b1
        If numBedGroups > 1 Then mixZones.Add Array(tmixCol, numBedGroups, "TMIX")
    End If

    ' --- Map each wsShort column back to the template row-18 cell whose
    ' formatting it should inherit, so per-level rows can be styled to match
    ' the custom table regardless of how the MIX/TMIX inserts shifted things ---
    Set formatSourceMap = CreateObject("Scripting.Dictionary")
    Dim fmKey As Variant, finalColForKey As Long
    For Each fmKey In headerMapShortRaw.Keys
        finalColForKey = GetColByHeader(headerMapShort, CStr(fmKey))
        If Not formatSourceMap.Exists(finalColForKey) Then
            formatSourceMap.Add finalColForKey, headerMapShortRaw(fmKey)
        End If
    Next fmKey

    ' The extra MIX/TMIX tally columns don't have their own row-18 entry -
    ' they all inherit the original single "MIX"/"TMIX" cell's formatting
    If mixExists Then
        Dim rawMixColForFormat As Long
        rawMixColForFormat = headerMapShortRaw("MIX")
        For b1 = 0 To numGroups - 1
            If Not formatSourceMap.Exists(mixCol + b1) Then formatSourceMap.Add mixCol + b1, rawMixColForFormat
        Next b1
    End If
    If tmixExists Then
        Dim rawTmixColForFormat As Long
        rawTmixColForFormat = headerMapShortRaw("TMIX")
        For b1 = 0 To numBedGroups - 1
            If Not formatSourceMap.Exists(tmixCol + b1) Then formatSourceMap.Add tmixCol + b1, rawTmixColForFormat
        Next b1
    End If

    ' Flag each unit row with a 1 in its (bed count, person count, dwelling
    ' type) tally column, and separately in its (bed count only) Total Mix
    ' tally column
    If headerMap.Exists("BEDS") Then
        For i = 2 To lastRowWork
            bVal = wsWork.Cells(i, GetColByHeader(headerMap, "BEDS")).Value
            If IsNumeric(bVal) And Len(bVal) > 0 Then
                dwellingType = GetUnitTitle(wsWork, i, headerMap)
                persVal = GetPersValue(wsWork, i, headerMap)
                groupKey = CDbl(bVal) & "|" & persVal & "|" & dwellingType
                If tally.Exists(groupKey) Then
                    wsWork.Cells(i, tally(groupKey)).Value = 1
                End If
                If tmixExists Then
                    If totalTally.Exists(CDbl(bVal)) Then
                        wsWork.Cells(i, totalTally(CDbl(bVal))).Value = 1
                    End If
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
    For c = 1 To tmixTallyCols.Count
        shortSumColumns.Add tmixTallyCols(c)
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
        iShort = 9 ' the template header occupies wsShort rows 1-8
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
        For c = 1 To workTotalTallyCols.Count
            workCol = workTotalTallyCols(c)
            shortCol = tmixTallyCols(c)
            wsShort.Cells(iShort, shortCol).Value = AggregateWorkRange(wsWork, workCol, 2, lastRowWork)
        Next c
        Call ApplyRowFormatting(wsTemplate, wsShort, formatSourceMap, iShort)
        shortChangeBlock.Add iShort
        iShort = iShort + 1
    Else
        iShort = 9 ' the template header occupies wsShort rows 1-8
        i = 2
        If hasLevel Then previousLevel = wsWork.Cells(2, GetColByHeader(headerMap, "LEVL")).Value
        If hasBlock Then previousBlock = wsWork.Cells(2, GetColByHeader(headerMap, "BLOK")).Value
        If hasZone Then previousZone = wsWork.Cells(2, GetColByHeader(headerMap, "ZONE")).Value
        levelStartRow = 2
        shortBlockStartRow = 9

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
                For c = 1 To workTotalTallyCols.Count
                    workCol = workTotalTallyCols(c)
                    shortCol = tmixTallyCols(c)
                    wsShort.Cells(iShort, shortCol).Value = AggregateWorkRange(wsWork, workCol, levelStartRow, i - 1)
                Next c
                Call ApplyRowFormatting(wsTemplate, wsShort, formatSourceMap, iShort)
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
        Call sumColumnsSub(wsShort, shortSumColumns, 9, iShort, 0, False)
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

' ============================================================================
' Shift every headerMapShort column position greater than afterCol by
' shiftAmount, in place - keeps the custom-table field map correct after a
' column insert on wsShort.
' ============================================================================
Sub ShiftHeaderMapShortColumns(headerMapShort As Object, afterCol As Long, shiftAmount As Long)
    Dim hKey As Variant
    For Each hKey In headerMapShort.Keys
        If headerMapShort(hKey) > afterCol Then
            headerMapShort(hKey) = headerMapShort(hKey) + shiftAmount
        End If
    Next hKey
End Sub

' ============================================================================
' Apply each column's template row-18 formatting to a just-written per-level
' row in wsShort, using formatSourceMap (wsShort column -> template row-18
' column) to account for any shift caused by the MIX/TMIX column inserts.
' ============================================================================
Sub ApplyRowFormatting(wsTemplate As Worksheet, wsShort As Worksheet, _
                      formatSourceMap As Object, targetRow As Long)
    Dim fsKey As Variant
    For Each fsKey In formatSourceMap.Keys
        Call CopyCellFormat(wsTemplate.Cells(18, formatSourceMap(fsKey)), wsShort.Cells(targetRow, CLng(fsKey)))
    Next fsKey
End Sub

' ============================================================================
' Copy font, fill (if not plain white), number format, alignment and wrap
' from one cell to another.
' ============================================================================
Sub CopyCellFormat(sourceCell As Range, destCell As Range)
    With destCell
        .Font.Name = sourceCell.Font.Name
        .Font.Size = sourceCell.Font.Size
        .Font.Bold = sourceCell.Font.Bold
        .Font.Italic = sourceCell.Font.Italic
        .Font.Color = sourceCell.Font.Color
        If sourceCell.Interior.Color <> 16777215 Then
            .Interior.Color = sourceCell.Interior.Color
        End If
        .NumberFormat = sourceCell.NumberFormat
        .HorizontalAlignment = sourceCell.HorizontalAlignment
        .VerticalAlignment = sourceCell.VerticalAlignment
        .WrapText = sourceCell.WrapText
    End With
End Sub

' ============================================================================
' Fixed sort order for mix columns: Apartment, then Duplex, then House,
' regardless of bedroom count. Anything else (e.g. GetUnitTitle's "Unit"
' fallback) sorts last.
' ============================================================================
Function DwellingTypeRank(dwellingType As String) As Long
    Select Case UCase(dwellingType)
        Case "APARTMENT": DwellingTypeRank = 0
        Case "DUPLEX": DwellingTypeRank = 1
        Case "HOUSE": DwellingTypeRank = 2
        Case Else: DwellingTypeRank = 3
    End Select
End Function

' ============================================================================
' Read a row's PERS (person count) value. Returns 0 if the PERS column is
' missing or the cell isn't numeric, so callers don't need to special-case it.
' ============================================================================
Function GetPersValue(ws As Worksheet, row As Long, headerMap As Object) As Double
    Dim v As Variant
    If headerMap.Exists("PERS") Then
        v = ws.Cells(row, GetColByHeader(headerMap, "PERS")).Value
        If IsNumeric(v) And Len(v) > 0 Then
            GetPersValue = CDbl(v)
            Exit Function
        End If
    End If
    GetPersValue = 0
End Function
