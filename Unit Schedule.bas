Option Explicit

' ============================================================================
' UNIT SCHEDULE MODULE
' Generates formatted unit schedule from wsSource data
' Organizes by Level > Block > Zone with summary rows
' ============================================================================

Sub GenerateUnitSchedule()
    Dim wsSource As Worksheet
    Dim wsLong As Worksheet
    Dim wsTemplate As Worksheet
    Dim headerMap As Object
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim filePath As String
    Dim currentDate As String
    Dim ws As Worksheet
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' Delete existing Long schedule sheets
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, "Long", vbTextCompare) > 0 Then
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
    
    ' Create wsLong worksheet
    currentDate = Format(Date, "yy-mm-dd")
    Set wsLong = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsLong.Name = "Long " & currentDate

    ' Build header maps
    Set headerMap = BuildHeaderMap(wsTemplate, 9)

    ' Unit types such as "1E4" are text, not 1 x 10^4. Text format the TYPE
    ' column before any data reaches it so Excel keeps the string it is given.
    If headerMap.Exists("TYPE") Then wsLong.columns(headerMap("TYPE")).NumberFormat = "@"

    ' Copy columns from wsSource to wsLong
    Call CopyColumnsByHeader(wsSource, wsLong, wsTemplate, 1, 9)
    
    lastCol = GetLastColumnFromHeaderMap(headerMap)
    
    ' Sort by ZONE > BLOK > LEVL
    Call SortSchedule(wsLong, headerMap, lastCol)
    
    ' Delete rows with XX in BLOK or LEVL
    Call DeleteInvalidRows(wsLong, headerMap)
    
    ' Apply dwelling lookup colors and standards
    Call ApplyDwellingStandards(wsLong, wsTemplate, headerMap)
    
    ' Add 10% indicator
    Call AddTenPercentIndicator(wsLong, headerMap)

    ' Add aggregate bedroom area (sum of BED1-BED5) per unit
    Call AddAggregateBedroomArea(wsLong, headerMap)

    ' Format schedule with level/block/zone summaries
    Call FormatScheduleWithSummaries(wsLong, wsTemplate, headerMap, lastCol)
    
    ' Copy headers from template
    
    wsLong.rows(1).Delete
    wsLong.rows(1).Resize(3).Insert Shift:=xlDown
    wsTemplate.rows("1:8").Copy
    wsLong.Range("A1").Insert Shift:=xlDown
    
    ' Add timestamp
    wsLong.Range("E5").Value = FormatDateWithSuffix(Date)
    
    ' Set print area
    lastRow = wsLong.Cells(wsLong.rows.Count, 1).End(xlUp).row
    wsLong.PageSetup.PrintArea = "A1:" & ColumnToLetter(lastCol) & lastRow + 1
    
    ' Activate and show print preview
    wsLong.Activate
    ActiveWindow.View = xlPageBreakPreview
    With wsLong.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$7:$9"
    End With
    
Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
End Sub

' ============================================================================
' Import data from CSV/TXT file into wsSource
' ============================================================================
' textHeaders lists the headers that must be imported as text instead of being
' parsed by Excel - defaults to TYPE, whose values ("1E4") are otherwise read
' as scientific notation and stored as the number 10000
Sub ImportData(wsSource As Worksheet, filePath As String, Optional textHeaders As Variant)
    Dim qt As QueryTable
    Dim isCsv As Boolean
    Dim colDataTypes As Variant
    
    If IsMissing(textHeaders) Then textHeaders = Array("TYPE")
    
    Application.DisplayAlerts = False
    wsSource.Cells.Clear
    
    ' Remove existing QueryTables
    For Each qt In wsSource.QueryTables
        qt.Delete
    Next qt
    
    isCsv = (LCase(Right(filePath, 4)) = ".csv")
    colDataTypes = BuildTextImportColumnTypes(filePath, isCsv, textHeaders)
    
    With wsSource.QueryTables.Add( _
        Connection:="TEXT;" & filePath, _
        Destination:=wsSource.Range("A1"))
        
        .TextFileParseType = xlDelimited
        .TextFileTabDelimiter = Not isCsv
        .TextFileCommaDelimiter = isCsv
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .AdjustColumnWidth = True
        ' Must be set before Refresh, and only when the header line could be
        ' read - otherwise leave Excel's own column guessing in place
        If Not IsEmpty(colDataTypes) Then .TextFileColumnDataTypes = colDataTypes
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    
    Application.DisplayAlerts = True
End Sub

' ============================================================================
' Sort schedule by ZONE > BLOK > LEVL > NO
' ============================================================================
Sub SortSchedule(ws As Worksheet, headerMap As Object, lastCol As Long)
    Dim lastRow As Long
    Dim zoneCol As Long, blokCol As Long, levelCol As Long, noCol As Long

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).row

    zoneCol = GetColByHeader(headerMap, "ZONE")
    blokCol = GetColByHeader(headerMap, "BLOK")
    levelCol = GetColByHeader(headerMap, "LEVL")
    noCol = FindColumnByAliases(headerMap, Array("NO", "NO.", "NUM", "UNIT NO", "UNIT NO."))

    With ws.Sort
        .SortFields.Clear

        ' Sort by ZONE if header exists
        If headerMap.Exists("ZONE") And zoneCol > 0 Then
            .SortFields.Add key:=ws.Range(ws.Cells(2, zoneCol), ws.Cells(lastRow, zoneCol)), _
                Order:=xlAscending
        End If

        ' Sort by BLOK if header exists
        If headerMap.Exists("BLOK") And blokCol > 0 Then
            .SortFields.Add key:=ws.Range(ws.Cells(2, blokCol), ws.Cells(lastRow, blokCol)), _
                Order:=xlAscending
        End If

        ' Sort by LEVL if header exists
        If headerMap.Exists("LEVL") And levelCol > 0 Then
            .SortFields.Add key:=ws.Range(ws.Cells(2, levelCol), ws.Cells(lastRow, levelCol)), _
                Order:=xlAscending
        End If

        ' Sort by NO (unit number) if header exists, under any common spelling
        If noCol > 0 Then
            .SortFields.Add key:=ws.Range(ws.Cells(2, noCol), ws.Cells(lastRow, noCol)), _
                Order:=xlAscending
        End If

        .SetRange ws.Range("A1:" & ColumnToLetter(lastCol) & lastRow)
        .Header = xlYes
        .Apply
    End With
End Sub

' ============================================================================
' Delete rows with XX in BLOK or LEVL columns
' Safely handles missing columns
' ============================================================================
Sub DeleteInvalidRows(ws As Worksheet, headerMap As Object)
    Dim lastRow As Long
    Dim blokCol As Long, levelCol As Long, zoneCol As Long, noCol As Long
    Dim i As Long
    Dim hasBlok As Boolean, hasLevel As Boolean, hasZone As Boolean, hasNo As Boolean

    hasBlok = headerMap.Exists("BLOK")
    hasLevel = headerMap.Exists("LEVL")
    hasZone = headerMap.Exists("ZONE")

    If hasBlok Then blokCol = GetColByHeader(headerMap, "BLOK")
    If hasLevel Then levelCol = GetColByHeader(headerMap, "LEVL")
    If hasZone Then zoneCol = GetColByHeader(headerMap, "ZONE")

    noCol = FindColumnByAliases(headerMap, Array("NO", "NO.", "NUM", "UNIT NO", "UNIT NO."))
    hasNo = (noCol > 0)

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).row

    For i = lastRow To 2 Step -1
        Dim isInvalid As Boolean
        isInvalid = False
        
        ' Nested Ifs, not "hasX And ws.Cells(...)" - VBA's And does not short-circuit,
        ' so the Cells() read would still run (and error on column 0) even when hasX is False
        If hasBlok Then
            If ws.Cells(i, blokCol).Value = "XX" Then isInvalid = True
        End If
        If hasLevel Then
            If ws.Cells(i, levelCol).Value = "XX" Then isInvalid = True
        End If
        If hasZone Then
            If ws.Cells(i, zoneCol).Value = "XX" Then isInvalid = True
        End If
        If hasNo Then
            If ws.Cells(i, noCol).Value = "XX" Then isInvalid = True
        End If
        
        If isInvalid Then ws.rows(i).Delete
    Next i
End Sub

' ============================================================================
' Apply dwelling lookup colors and minimum standards
' Safely handles missing columns
' ============================================================================
Sub ApplyDwellingStandards(wsData As Worksheet, wsTemplate As Worksheet, _
                          headerMap As Object)
    Dim lastRow As Long
    Dim i As Long
    Dim rng As Range
    Dim descCol As Long, areaCol As Long, minAreaCol As Long
    Dim pasCol As Long, minPasCol As Long
    Dim hasDesc As Boolean, hasArea As Boolean, hasMinArea As Boolean
    Dim hasPas As Boolean, hasMinPas As Boolean

    ' Check which columns exist
    hasDesc = headerMap.Exists("BEDTYPE")
    hasArea = headerMap.Exists("GIFA")
    hasMinArea = headerMap.Exists("minAREA")
    hasPas = headerMap.Exists("PAS")
    hasMinPas = headerMap.Exists("minPAS")

    ' Get column numbers for existing columns
    If hasDesc Then descCol = GetColByHeader(headerMap, "BEDTYPE")
    If hasArea Then areaCol = GetColByHeader(headerMap, "GIFA")
    If hasMinArea Then minAreaCol = GetColByHeader(headerMap, "minAREA")
    If hasPas Then pasCol = GetColByHeader(headerMap, "PAS")
    If hasMinPas Then minPasCol = GetColByHeader(headerMap, "minPAS")

    lastRow = wsData.Cells(wsData.rows.Count, 1).End(xlUp).row

    For i = 2 To lastRow
        Set rng = wsData.Range(wsData.Cells(i, "A"), wsData.Cells(i, GetLastColumnFromHeaderMap(headerMap)))

        Dim dwellingType As String
        If hasDesc Then
            dwellingType = wsData.Cells(i, descCol).Value
        Else
            dwellingType = "Apartment" ' Default if no BEDTYPE column
        End If

        Select Case True
            Case InStr(1, UCase(dwellingType), "HOUSE") > 0
                Call ApplyDwellingLookup(wsData, wsTemplate, i, dwellingType, rng, headerMap)

            Case InStr(1, UCase(dwellingType), "DUPLEX") > 0 _
              Or InStr(1, UCase(dwellingType), "DUP") > 0
                Call ApplyDwellingLookup(wsData, wsTemplate, i, dwellingType, rng, headerMap)

            Case InStr(1, UCase(dwellingType), "APARTMENT") > 0 _
              Or InStr(1, UCase(dwellingType), "APT") > 0
                Call ApplyDwellingLookup(wsData, wsTemplate, i, dwellingType, rng, headerMap)
        End Select

        ' GFA Check - only if both columns exist
        If hasMinArea And hasArea Then
            If Val(wsData.Cells(i, minAreaCol).Value) > Val(wsData.Cells(i, areaCol).Value) _
               And Val(wsData.Cells(i, areaCol).Value) > 0 Then
                wsData.Cells(i, minAreaCol).Interior.Color = RGB(255, 0, 0)
            End If
        End If

        ' Private Amenity Check - only if both columns exist
        If hasPas And hasMinPas Then
            If Val(wsData.Cells(i, pasCol).Value) < Val(wsData.Cells(i, minPasCol).Value) _
               And Val(wsData.Cells(i, minPasCol).Value) > 0 Then
                wsData.Cells(i, pasCol).Interior.Color = RGB(255, 0, 0)
            End If
        End If
    Next i
End Sub

' ============================================================================
' Add 10% area indicator
' Safely handles missing columns
' ============================================================================
Sub AddTenPercentIndicator(ws As Worksheet, headerMap As Object)
    Dim lastRow As Long
    Dim i As Long
    Dim areaCol As Long, minAreaCol As Long, min10Col As Long
    Dim areaExt As Double, areaCur As Double

    ' Exit if required columns don't exist
    If Not headerMap.Exists("GIFA") Or Not headerMap.Exists("MIN10") Or Not headerMap.Exists("MINAREA") Then Exit Sub

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).row

    areaCol = GetColByHeader(headerMap, "GIFA")
    minAreaCol = GetColByHeader(headerMap, "minAREA")
    min10Col = GetColByHeader(headerMap, "MIN10")

    ws.Cells(1, min10Col).Value = "MIN10"

    For i = 2 To lastRow
        areaExt = Val(ws.Cells(i, minAreaCol).Value) * 1.1
        areaCur = Val(ws.Cells(i, areaCol).Value)

        If areaCur > areaExt Then
            ws.Cells(i, min10Col).Value = "1"
        Else
            ws.Cells(i, min10Col).Value = "0"
        End If
    Next i
End Sub

' ============================================================================
' Compute each unit's aggregate bedroom area (sum of BED1-BED5) into AGBED.
' Safely handles missing columns - AGBED is skipped entirely if not present,
' and each BED1-BED5 field is only added if it exists.
' ============================================================================
Sub AddAggregateBedroomArea(ws As Worksheet, headerMap As Object)
    Dim lastRow As Long
    Dim i As Long
    Dim agbedCol As Long
    Dim unitBedroomArea As Double
    Dim bedFields As Variant
    Dim f As Long

    If Not headerMap.Exists("AGBED") Then Exit Sub

    agbedCol = GetColByHeader(headerMap, "AGBED")
    bedFields = Array("BED1", "BED2", "BED3", "BED4", "BED5")

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).row

    For i = 2 To lastRow
        unitBedroomArea = 0
        For f = LBound(bedFields) To UBound(bedFields)
            If headerMap.Exists(CStr(bedFields(f))) Then
                unitBedroomArea = unitBedroomArea + Val(ws.Cells(i, GetColByHeader(headerMap, CStr(bedFields(f)))).Value)
            End If
        Next f
        ws.Cells(i, agbedCol).Value = unitBedroomArea
    Next i
End Sub

' ============================================================================
' Format schedule with level, block, and zone summaries
' Handles missing ZONE/BLOK/LEVL columns gracefully
' ============================================================================
Sub FormatScheduleWithSummaries(ws As Worksheet, wsTemplate As Worksheet, _
                               headerMap As Object, lastCol As Long)
    Dim lastRow As Long
    Dim i As Long
    Dim currentLevel As Variant, previousLevel As Variant
    Dim currentBlock As Variant, previousBlock As Variant
    Dim currentZone As Variant, previousZone As Variant
    Dim levelStartRow As Long, blockStartRow As Long, zoneStartRow As Long
    Dim changeLevel As Collection, changeBlock As Collection, changeZone As Collection
    Dim sumColumns As Collection, percentCalcColumns As Collection
    Dim bedCountsDict As Object, tally As Object
    Dim bedKeys As Variant, bCount As Variant
    Dim colLet As Long, resColLet As Long, b1 As Long, b2 As Long, tempB As Variant
    Dim levelRange As Range, levelTitle As String
    Dim re As Object, re1 As Object
    Dim regexPattern As String
    Dim hasZone As Boolean, hasBlock As Boolean, hasLevel As Boolean
    Dim finalSummaryCollection As Collection
    
    ' Check which grouping columns exist
    hasZone = headerMap.Exists("ZONE")
    hasBlock = headerMap.Exists("BLOK")
    hasLevel = headerMap.Exists("LEVL")
    
    ' Initialize collections
    Set changeLevel = New Collection
    Set changeBlock = New Collection
    Set changeZone = New Collection
    Set finalSummaryCollection = New Collection

    ' Setup sum columns with safety checks
    Set sumColumns = New Collection
    Dim noColForSum As Long
    noColForSum = FindColumnByAliases(headerMap, Array("NO", "NO.", "NUM", "UNIT NO", "UNIT NO."))
    If noColForSum > 0 Then sumColumns.Add noColForSum
    If headerMap.Exists("GIFA") Then sumColumns.Add GetColByHeader(headerMap, "GIFA")
    If headerMap.Exists("minAREA") Then sumColumns.Add GetColByHeader(headerMap, "MINAREA")
    If headerMap.Exists("BEDS") Then sumColumns.Add GetColByHeader(headerMap, "BEDS")
    If headerMap.Exists("PERS") Then sumColumns.Add GetColByHeader(headerMap, "PERS")
    If headerMap.Exists("DUAL") Then sumColumns.Add GetColByHeader(headerMap, "DUAL")
    If headerMap.Exists("minPAS") Then sumColumns.Add GetColByHeader(headerMap, "MINPAS")
    If headerMap.Exists("PAS") Then sumColumns.Add GetColByHeader(headerMap, "PAS")
    If headerMap.Exists("minCAS") Then sumColumns.Add GetColByHeader(headerMap, "MINCAS")
    If headerMap.Exists("min10") Then sumColumns.Add GetColByHeader(headerMap, "MIN10")

    ' Find unique bedroom counts and setup tally columns
    Set bedCountsDict = CreateObject("Scripting.Dictionary")
    Set tally = CreateObject("Scripting.Dictionary")
    ' Set sumTypeColumns = New Collection
    Set percentCalcColumns = New Collection

    lastRow = ws.Cells(ws.rows.Count, "C").End(xlUp).row

    Dim bVal As Variant
    If headerMap.Exists("BEDS") Then
        For i = 2 To lastRow
            bVal = ws.Cells(i, GetColByHeader(headerMap, "BEDS")).Value
            If IsNumeric(bVal) And Len(bVal) > 0 Then
                bVal = CDbl(bVal)
                If Not bedCountsDict.Exists(bVal) Then bedCountsDict.Add bVal, bVal
            End If
        Next i
    End If

    ' Sort bedroom keys
    bedKeys = bedCountsDict.Keys
    For b1 = LBound(bedKeys) To UBound(bedKeys) - 1
        For b2 = b1 + 1 To UBound(bedKeys)
            If bedKeys(b1) > bedKeys(b2) Then
                tempB = bedKeys(b1)
                bedKeys(b1) = bedKeys(b2)
                bedKeys(b2) = tempB
            End If
        Next b2
    Next b1

    ' Create tally columns
    Dim startCol As Long
    startCol = lastCol + 2

    For b1 = LBound(bedKeys) To UBound(bedKeys)
        bCount = bedKeys(b1)
        If IsNumeric(bCount) And bCount > 0 Then
            colLet = startCol + b1
            tally.Add bCount, colLet
            ws.Cells(1, colLet).Value = bCount & " BED"
            sumColumns.Add colLet
            percentCalcColumns.Add colLet
        End If
    Next b1

    If headerMap.Exists("min10") Then percentCalcColumns.Add GetColByHeader(headerMap, "min10")
    If headerMap.Exists("DUAL") Then percentCalcColumns.Add GetColByHeader(headerMap, "DUAL")

    'Apply formating from template to new columns

    Call CopyRowFormattingDown(wsTemplate, 9, ws, 3, lastRow, 1, lastCol)

    ' Initialize regex
    Set re1 = CreateObject("VBScript.RegExp")
    With re1
        .Pattern = "^[A-Za-z0-9]{1,2}$"
        .IgnoreCase = True
        .Global = False
    End With

    Set re = CreateObject("VBScript.RegExp")
    regexPattern = Trim(wsTemplate.Range("X2").Value)
    If Len(regexPattern) = 0 Then regexPattern = ".*"
    With re
        .Global = False
        .IgnoreCase = True
        .Pattern = regexPattern
    End With

    ' Initialize loop variables with safety checks
    
    
    ws.rows(2).Resize(1).Insert Shift:=xlDown

    
    i = 3
    If hasLevel Then previousLevel = ws.Cells(5, GetColByHeader(headerMap, "LEVL")).Value
    If hasBlock Then previousBlock = ws.Cells(5, GetColByHeader(headerMap, "BLOK")).Value
    If hasZone Then previousZone = ws.Cells(5, GetColByHeader(headerMap, "ZONE")).Value
    levelStartRow = i
    blockStartRow = i
    zoneStartRow = i

    ' Main loop through schedule
    Do While True
        'add counts for bedroom counts

        If headerMap.Exists("BEDS") Then
            bVal = ws.Cells(i, GetColByHeader(headerMap, "BEDS")).Value
            If IsNumeric(bVal) And tally.Exists(CDbl(bVal)) Then
                ws.Cells(i, tally(CDbl(bVal))).Value = 1
            End If
        End If
        
    
        ' Get current values with safety checks
        If hasLevel Then currentLevel = ws.Cells(i, GetColByHeader(headerMap, "LEVL")).Value
        If hasBlock Then currentBlock = ws.Cells(i, GetColByHeader(headerMap, "BLOK")).Value
        If hasZone Then currentZone = ws.Cells(i, GetColByHeader(headerMap, "ZONE")).Value

        ' Check for changes in any grouping level
        Dim levelChanged As Boolean, blockChanged As Boolean, zoneChanged As Boolean
        levelChanged = (hasLevel And currentLevel <> previousLevel)
        blockChanged = (hasBlock And currentBlock <> previousBlock)
        zoneChanged = (hasZone And currentZone <> previousZone)
        
        ' Handle missing columns - treat as "changed" to trigger summary if needed
        If Not hasLevel And Not hasBlock And Not hasZone Then
            ' No grouping columns - just process all rows without summaries
            levelChanged = False
        ElseIf Not hasLevel And hasBlock Then
            levelChanged = False ' Skip level changes if no level column
        ElseIf Not hasBlock And hasLevel Then
            blockChanged = False ' Skip block changes if no block column
        End If

        If levelChanged Or blockChanged Or zoneChanged Then
            ' Insert separator rows
            ws.rows(i).Resize(3).Insert Shift:=xlDown
            ws.rows(i).Resize(3).Interior.ColorIndex = -4142

            ' Add summaries for level change
            Call sumColumnsSub(ws, sumColumns, levelStartRow, i, 0, noColForSum > 0)
            ' Call sumColumnsSub(ws, sumTypeColumns, levelStartRow, i, 4, False)
            Call percentColumnsSub(ws, percentCalcColumns, i, 0)

            ' Add borders
            Set levelRange = ws.Range(ws.Cells(i - 1, "A"), ws.Cells(levelStartRow, lastCol))
            Call drawBorderThickOutline(levelRange)

            ' Add level title (adapt based on available columns)
            If hasBlock And hasLevel Then
                If re1.Test(previousBlock) Then
                    levelTitle = "Block " & previousBlock & " Level " & previousLevel
                Else
                    levelTitle = previousBlock
                End If
            ElseIf hasBlock Then
                levelTitle = "Block " & previousBlock
            ElseIf hasLevel Then
                levelTitle = "Level " & previousLevel
            ElseIf hasZone Then
                levelTitle = "Zone " & previousZone
            Else
                levelTitle = "Units"
            End If

            With ws.Cells(levelStartRow - 1, "B")
                .Value = levelTitle
                .Font.Bold = True
                .Font.Color = RGB(0, 176, 240)
                .Font.Name = "Calibri"
                .HorizontalAlignment = xlLeft
            End With

            changeLevel.Add i

            ' Handle block/zone changes
            If blockChanged Or zoneChanged Then
                Dim blockEndRow As Long, blockTitle As String
                blockEndRow = i - 1

                If hasBlock Then
                    If re1.Test(previousBlock) Then
                        blockTitle = "Block " & previousBlock & " Summary"
                    Else
                        blockTitle = previousBlock & " Summary"
                    End If
                Else
                    blockTitle = "Summary"
                End If

                ws.rows(i).Resize(3).Interior.ColorIndex = -4142

                ' Only add a block summary section if we have meaningful grouping.
                ' Either way, this block's contribution must reach changeBlock so
                ' the zone summary (below) totals correctly even when no block
                ' summary row was actually written.
                If (Not hasLevel Or previousLevel <> 0) And previousLevel <> "NA" Then
                    i = i + 3
                    ws.rows(i).Resize(3).Insert Shift:=xlDown

                    Call drawBorderLine(ws, i, lastCol)

                    With ws.Cells(i - 1, "B")
                        .Value = blockTitle
                        .Font.Bold = True
                        .Font.Color = RGB(0, 176, 240)
                        .Font.Name = "Calibri"
                        .HorizontalAlignment = xlLeft
                    End With

                    ' Call sumColumnsSub(ws, sumTypeColumns, blockStartRow, i, 4, False)
                    Call percentColumnsSub(ws, percentCalcColumns, i, 0)
                    Call sumColumnsRowsSub(ws, sumColumns, changeLevel, i)

                    changeBlock.Add i
                Else
                    ' No block summary written - fold this block's own level
                    ' rows straight into the zone total instead of losing them
                    Dim clItem As Variant
                    For Each clItem In changeLevel
                        changeBlock.Add clItem
                    Next clItem
                End If

                ' Handle zone changes - always, regardless of whether a block
                ' summary was written above
                If zoneChanged And hasZone Then
                    i = i + 3
                    ws.rows(i).Resize(3).Insert Shift:=xlDown
                    changeZone.Add i
                    Call drawBorderLine(ws, i, lastCol)

                    With ws.Cells(i - 1, "B")
                        .Value = "Zone " & previousZone & " Summary"
                        .Font.Bold = True
                        .Font.Color = RGB(0, 176, 240)
                        .Font.Name = "Calibri"
                        .HorizontalAlignment = xlLeft
                    End With

                    Call sumColumnsRowsSub(ws, sumColumns, changeBlock, i)
                    Set changeBlock = New Collection
                End If

                previousZone = currentZone

                blockStartRow = i + 1
                Set changeLevel = New Collection
                If hasBlock Then previousBlock = currentBlock
            End If

            i = i + 3
            levelStartRow = i
            lastRow = lastRow + 3
        End If

        If hasLevel Then previousLevel = currentLevel
        i = i + 1

        If ws.Cells(i - 1, 1).Value = 0 Or i > 100000 Then Exit Do
    Loop

    ' Add whole scheme summary
    If Not hasLevel And Not hasBlock And Not hasZone Then
        ' No grouping - just add summary at the end
        Dim rng As Range
        i = i - 2
        Set rng = ws.Range(ws.Cells(5, 1), ws.Cells(i, lastCol))
        Call drawBorderThickOutline(rng)
        i = i + 3
    End If

    i = i - 1
    Call drawBorderLine(ws, i, lastCol)   

    With ws.Cells(i - 1, "B")
        .Value = "Whole Scheme Summary"
        .Font.Bold = True
        .Font.Color = RGB(0, 176, 240)
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlLeft
    End With

    ' Call sumColumnsSub(ws, sumTypeColumns, 4, i, 4, False)
    Call percentColumnsSub(ws, percentCalcColumns, i, 0)
    
    ' Determine which collection to use for final summary based on available columns
    If hasZone And changeZone.Count > 0 Then
        Call sumColumnsRowsSub(ws, sumColumns, changeZone, i)
    ElseIf hasBlock And changeBlock.Count > 0 Then
        Call sumColumnsRowsSub(ws, sumColumns, changeBlock, i)
    ElseIf hasLevel And changeLevel.Count > 0 Then
        Call sumColumnsRowsSub(ws, sumColumns, changeLevel, i)
    Else
        ' No grouping - just sum all rows
        Call sumColumnsSub(ws, sumColumns, 4, i, 0, noColForSum > 0)
    End If

    ' Final formatting
    Call FormatScheduleColumns(ws, headerMap, lastCol)

End Sub

' ============================================================================
' Apply final column formatting
' ============================================================================
Sub FormatScheduleColumns(ws As Worksheet, headerMap As Object, lastCol As Long)
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
    
    If headerMap.Exists("BEDTYPE") Then
        With ws.Range(ws.Cells(1, GetColByHeader(headerMap, "BEDTYPE")), ws.Cells(lastRow, GetColByHeader(headerMap, "BEDTYPE")))
            .HorizontalAlignment = xlLeft
            .Font.Bold = False
        End With
        ws.columns(GetColByHeader(headerMap, "BEDTYPE")).AutoFit
    End If
       
    
    
    

End Sub






