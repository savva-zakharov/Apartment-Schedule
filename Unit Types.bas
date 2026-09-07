Option Explicit

' ============================================================================
' UNIT TYPES MODULE
' Generates formatted unit types from wsSource data
' ============================================================================

Sub UnitTypes()
    Dim wsSource As Worksheet
    Dim wsTypes As Worksheet
    Dim wsTemplate As Worksheet
    Dim wsStats As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim filePath As String
    Dim typeDict As Object
    Set typeDict = CreateObject("Scripting.Dictionary")
    Dim currentDate As String
    currentDate = Format(Date, "yy-mm-dd") ' You can change format here
    Dim rng As Range
        ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Dim ws As Worksheet


    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, "Types", vbTextCompare) > 0 Or InStr(1, ws.Name, "Stats", vbTextCompare) > 0 Then
            ws.Delete
        End If
    Next ws

        ' Set the original worksheet
    Set wsSource = ThisWorkbook.Sheets("sourceData") ' Change to your original sheet name if needed
    
    ' Set the tempalte worksheet
    Set wsTemplate = ThisWorkbook.Sheets("template") ' Change to your original sheet name if needed

    filePath = Trim(wsTemplate.Range("AA5").Value)
    
    
    ' Remove surrounding double quotes if present (tolerant of quoted paths)
    If Left(filePath, 1) = """" And Right(filePath, 1) = """" Then
        filePath = Mid(filePath, 2, Len(filePath) - 2)
    End If

    If Dir(filePath) = "" Or Dir(filePath) = "NA" Then
        MsgBox "File not found:" & vbCrLf & filePath, vbExclamation
        Exit Sub
    Else:
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        wsSource.Cells.Clear
            ' Remove any existing QueryTables (important!)
        Dim qt As QueryTable
        For Each qt In wsSource.QueryTables
            qt.Delete
        Next qt

        Dim isCsv As Boolean
        isCsv = (LCase(Right(filePath, 4)) = ".csv")

        With wsSource.QueryTables.Add( _
            Connection:="TEXT;" & filePath, _
            Destination:=wsSource.Range("A1"))

            .TextFileParseType = xlDelimited
            .TextFileTabDelimiter = Not isCsv
            .TextFileCommaDelimiter = isCsv
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .AdjustColumnWidth = True
            .Refresh BackgroundQuery:=False
            .Delete ' remove query but keep data
        End With
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True

    End If

    ' Create a new worksheet for the unit types
    Set wsTypes = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsTypes.Name = wsTypes.Name & " Types " & currentDate

    ' Create a new worksheet for the unit stat blocks
    Set wsStats = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsStats.Name = wsStats.Name & " Stats " & currentDate

    ' Build header map for wsSource (header row = 1)
    Dim sourceHeaderMap As Object
    Set sourceHeaderMap = BuildHeaderMap(wsSource, 1)

    ' Build header map for wsTypes (header row = 29)
    Dim typeHeaderMap As Object
    Set typeHeaderMap = BuildHeaderMap(wsTemplate, 29)

    Dim typeLastCol As Long
    typeLastCol = GetLastColumnFromHeaderMap(typeHeaderMap)

    ' Copy headers to wsTypes
    Call copyCellsByHeader(wsTemplate, wsTypes, 28, 1, typeHeaderMap, typeHeaderMap)

    

    '############################
    '## find unique unit types ##
    '############################
    Dim reTypes As Object
    Set reTypes = CreateObject("VBScript.RegExp")
    
    Dim regexPattern As String
    
    ' Read the cell
    regexPattern = Trim(wsTemplate.Range("X3").Value)
    
    ' Check if empty and assign default
    If Len(regexPattern) = 0 Then
        regexPattern = ".*"    ' default regex: matches anything
    End If
    
    ' Apply to your regex object
    With reTypes
        .Global = False
        .IgnoreCase = True
        .Pattern = regexPattern
    End With

    Dim unitType
    Dim unitKey
    Dim tempArr

    lastRow = wsSource.Cells(wsSource.rows.Count, 1).End(xlUp).row

    ' Columns used to disregard invalid rows (marked "XX")
    Dim levlCol As Long, blokCol As Long, noCol As Long
    Dim hasLevl As Boolean, hasBlok As Boolean, hasNo As Boolean
    hasLevl = sourceHeaderMap.Exists("LEVL")
    hasBlok = sourceHeaderMap.Exists("BLOK")
    hasNo = sourceHeaderMap.Exists("NO")
    If hasLevl Then levlCol = GetColByHeader(sourceHeaderMap, "LEVL")
    If hasBlok Then blokCol = GetColByHeader(sourceHeaderMap, "BLOK")
    If hasNo Then noCol = GetColByHeader(sourceHeaderMap, "NO")

    For i = 2 To lastRow
        ' Disregard rows marked "XX" in LEVL, BLOK or NO.
        ' Nested Ifs, not "hasX And wsSource.Cells(...)" - VBA's And does not
        ' short-circuit, so the Cells() read would still run (and error on
        ' column 0) even when hasX is False
        Dim isXXRow As Boolean
        isXXRow = False
        If hasLevl Then
            If wsSource.Cells(i, levlCol).Value = "XX" Then isXXRow = True
        End If
        If hasBlok Then
            If wsSource.Cells(i, blokCol).Value = "XX" Then isXXRow = True
        End If
        If hasNo Then
            If wsSource.Cells(i, noCol).Value = "XX" Then isXXRow = True
        End If
        If isXXRow Then GoTo NextRow

        unitType = wsSource.Cells(i, GetColByHeader(sourceHeaderMap, "TYPE")).Value

        If Len(unitType) > 0 And reTypes.Test(unitType) Then
    
            ' Use regex match as the dictionary key
            unitKey = UCase(Trim(reTypes.Execute(unitType)(0)))
    
            If Not typeDict.Exists(unitKey) Then
                ' Store count = 1 and first row = i
                typeDict.Add unitKey, Array(1, i)
            Else
                ' Increment count
                tempArr = typeDict(unitKey)
                tempArr(0) = tempArr(0) + 1
                typeDict(unitKey) = tempArr
            End If
        End If
NextRow:
    Next i


    Dim outputRow As Long
    outputRow = 2

    Dim typeKeys As Variant
    typeKeys = typeDict.Keys
    Dim typeItems As Variant
    typeItems = typeDict.Items
    Dim sourceRow As Long
    Dim key
    '############################################
    '# Copy Unique Unit Types and apply colours #
    '############################################

    
    Dim minAreaCol As Long
    minAreaCol = GetColByHeader(typeHeaderMap, "minAREA")
    Dim descCol As Long
    descCol = GetColByHeader(typeHeaderMap, "BEDTYPE")
    Dim areaCol As Long
    areaCol = GetColByHeader(typeHeaderMap, "GIFA")
    Dim pasCol As Long
    pasCol = GetColByHeader(typeHeaderMap, "PAS")
    Dim minPasCol As Long
    minPasCol = GetColByHeader(typeHeaderMap, "minPAS")
    Dim min10Col As Long
    min10Col = GetColByHeader(typeHeaderMap, "min10")
        
    Dim unitBedroomArea As Double
    Dim percentFormula As String
    Dim totalUnits As Long

    For Each key In typeKeys
        totalUnits = totalUnits + typeDict(key)(0)
    Next key

    For key = LBound(typeKeys) To UBound(typeKeys)
        Dim itemArray As Variant

        ' ? Step 1: Extract the inner array first
        itemArray = typeItems(key)

        ' ? Step 2: Then access the element
        sourceRow = itemArray(1)
        
        Call copyCellsByHeader(wsSource, wsTypes, sourceRow, outputRow, sourceHeaderMap, typeHeaderMap)


        ' overwrite column A with count
        wsTypes.Cells(outputRow, GetColByHeader(typeHeaderMap, "No")).Value = typeItems(key)(0)
        ' overwrite column E with combined unit type
        wsTypes.Cells(outputRow, GetColByHeader(typeHeaderMap, "Type")).Value = typeKeys(key)

        ' CALCULATE UNIT %
        If TypeHeaderMap.Exists("%") Then
            wsTypes.Cells(outputRow, GetColByHeader(typeHeaderMap, "%")).Value =  Format(typeItems(key)(0) / totalUnits, "0%")

        End If

        Set rng = wsTypes.Range(wsTypes.Cells(outputRow, 3), wsTypes.Cells(outputRow, typeLastCol))
        
        Dim dwellingType As String
        dwellingType = wsTypes.Cells(outputRow, descCol).Value

        Select Case True

            ' HOUSES
            Case InStr(1, UCase(dwellingType), "HOUSE") > 0
                Call ApplyDwellingLookup(wsTypes, wsTemplate, outputRow, dwellingType, rng, typeHeaderMap)

            ' DUPLEX
            Case InStr(1, UCase(dwellingType), "DUPLEX") > 0 _
              Or InStr(1, UCase(dwellingType), "DUP") > 0
                Call ApplyDwellingLookup(wsTypes, wsTemplate, outputRow, dwellingType, rng, typeHeaderMap)

            ' APARTMENTS
            Case InStr(1, UCase(dwellingType), "APARTMENT") > 0 _
              Or InStr(1, UCase(dwellingType), "APT") > 0
                Call ApplyDwellingLookup(wsTypes, wsTemplate, outputRow, dwellingType, rng, typeHeaderMap)

        End Select

        If typeHeaderMap.Exists("AGBED") Then
            unitBedroomArea = 0
            If typeHeaderMap.Exists("BED1") Then
                unitBedroomArea = unitBedroomArea + Val(wsTypes.Cells(outputRow, GetColByHeader(typeHeaderMap, "BED1")).Value)
            End If
            If typeHeaderMap.Exists("BED2") Then
                unitBedroomArea = unitBedroomArea + Val(wsTypes.Cells(outputRow, GetColByHeader(typeHeaderMap, "BED2")).Value)
            End If
            If typeHeaderMap.Exists("BED3") Then
                unitBedroomArea = unitBedroomArea + Val(wsTypes.Cells(outputRow, GetColByHeader(typeHeaderMap, "BED3")).Value)
            End If
            If typeHeaderMap.Exists("BED4") Then
                unitBedroomArea = unitBedroomArea + Val(wsTypes.Cells(outputRow, GetColByHeader(typeHeaderMap, "BED4")).Value)
            End If
            If typeHeaderMap.Exists("BED5") Then
                unitBedroomArea = unitBedroomArea + Val(wsTypes.Cells(outputRow, GetColByHeader(typeHeaderMap, "BED5")).Value)
            End If
            wsTypes.Cells(outputRow, GetColByHeader(typeHeaderMap, "AGBED")).Value = unitBedroomArea
        End If


        ' COMPLIANCE CHECK – cell-level only
        
        ' GFA CHECK - check that the floor area matches the minimum area requirement for the unit type
        If typeHeaderMap.Exists("minAREA") And typeHeaderMap.Exists("GIFA") Then
            If Val(wsTypes.Cells(outputRow, minAreaCol).Value) > Val(wsTypes.Cells(outputRow, areaCol).Value) _
            And Val(wsTypes.Cells(outputRow, typeHeaderMap("GIFA")).Value) > 0 Then
                wsTypes.Cells(outputRow, minAreaCol).Interior.Color = RGB(255, 0, 0)
            End If
        End If

        '10%
        If typeHeaderMap.Exists("min10") Then
            If Val(wsTypes.Cells(outputRow, areaCol).Value) > Val(wsTypes.Cells(outputRow, minAreaCol).Value) * 1.1 Then
                wsTypes.Cells(outputRow, min10Col).Value = 1
            Else
                wsTypes.Cells(outputRow, GetColByHeader(typeHeaderMap, "min10")).Value = 0
            End If
        End If            

        ' PRIVATE AMENITY AREA CHECK - check that the amenity area meets the minimum requirement for the unit type
        If typeHeaderMap.Exists("minPAS") And typeHeaderMap.Exists("PAS") Then
            If Val(wsTypes.Cells(outputRow, pasCol).Value) < Val(wsTypes.Cells(outputRow, minPasCol).Value) _
            And Val(wsTypes.Cells(outputRow, minPasCol).Value) > 0 Then
                wsTypes.Cells(outputRow, pasCol).Interior.Color = RGB(255, 0, 0)
            End If
        End If


    
        outputRow = outputRow + 1
    Next key

    '######################
    '# wsTYPES FORMATTING #
    '######################

    With wsTypes.Sort
        .SortFields.Clear
        .SortFields.Add key:=wsTypes.Range(wsTypes.Cells(2, GetColByHeader(typeHeaderMap, "TYPE")), wsTypes.Cells(lastRow, GetColByHeader(typeHeaderMap, "TYPE"))), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        .SetRange wsTypes.Range(wsTypes.Cells(2, 1), wsTypes.Cells(lastRow, typeLastCol))
        .Header = xlNo
        .Apply
    End With

    wsTypes.Cells(outputRow, 1).Value = totalUnits
    

    Set rng = wsTypes.Range(wsTypes.Cells(2, 1), wsTypes.Cells(outputRow - 1, typeLastCol))
    Call drawBorderThickOutline(rng)

    ' Call FormatColumnsByPattern(wsTypes, typeHeaderMap, "MIN", True, True, RGB(128, 128, 128))
    lastRow = wsTypes.Cells(wsSource.rows.Count, 1).End(xlUp).row
    Call CopyRowFormattingDown(wsTemplate, 29, wsTypes, 1, lastRow, 1, typeLastCol)

    wsTypes.rows(1).Clear

    wsTemplate.Range(wsTemplate.Cells(20, 1), wsTemplate.Cells(28, typeLastCol)).Copy
    wsTypes.Range("A1").Insert Shift:=xlDown

    lastRow = wsTypes.Cells(wsSource.rows.Count, 1).End(xlUp).row

    Dim lastColLetter As String
    lastColLetter = ColumnToLetter(typeLastCol)

    wsTypes.PageSetup.PrintArea = "A1:" & lastColLetter & lastRow
    wsTypes.Activate
    ActiveWindow.View = xlPageBreakPreview
    With wsTypes.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False ' Can be 1 or left as False to auto-scale height
    End With

End Sub
