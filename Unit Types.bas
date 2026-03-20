Option Explicit

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

        ' Build header map for wsTypes (header row = 61)
    Dim tempHeaderMap As Object
    Set tempHeaderMap = BuildHeaderMap(wsTemplate, 61)

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

    For i = 2 To lastRow
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
    Next i
    

    Dim outputRow As Long
    outputRow = 2
    
    Dim typeKeys As Variant
    typeKeys = typeDict.Keys
    Dim typeItems As Variant
    typeItems = typeDict.Items
    Dim sourceRow As Long
    Dim key
    Dim rng As Range
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
    
        Select Case True
    
            ' HOUSES
            Case InStr(1, UCase(wsTypes.Cells(outputRow, descCol).Value), "HOUSE") > 0
                Call ApplyDwellingLookup(wsTypes, wsTemplate, outputRow, "A81:A89", rng, typeHeaderMap, tempHeaderMap)
    
            ' DUPLEX
            Case InStr(1, UCase(wsTypes.Cells(outputRow, descCol).Value), "DUPLEX") > 0 _
              Or InStr(1, UCase(wsTypes.Cells(outputRow, descCol).Value), "DUP") > 0
                Call ApplyDwellingLookup(wsTypes, wsTemplate, outputRow, "A71:A76", rng, typeHeaderMap, tempHeaderMap)
    
            ' APARTMENTS
            Case InStr(1, UCase(wsTypes.Cells(outputRow, descCol).Value), "APARTMENT") > 0 _
              Or InStr(1, UCase(wsTypes.Cells(outputRow, descCol).Value), "APT") > 0
                Call ApplyDwellingLookup(wsTypes, wsTemplate, outputRow, "A62:A67", rng, typeHeaderMap, tempHeaderMap)
    
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

    Call FormatColumnsByPattern(wsTypes, typeHeaderMap, "MIN", True, True, RGB(128, 128, 128))

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
    

    '#######################################
    '# Creating Unit Stat Blocks Worksheet #
    '#######################################


    Dim iStats As Long
    iStats = 2

    Dim unitTitle As String
    ' Dim unitType as String
    Dim unitBeds As String
    Dim unitPers As String
    Dim unitStartRow As Long

For i = 10 To lastRow
    unitStartRow = iStats
    ' unitTitle = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BEDTYPE")).Value
    unitType = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "Type")).Value


    Select Case True
    
        ' HOUSES
        Case InStr(1, UCase(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BEDTYPE")).Value), "HOUSE") > 0
            unitTitle = "House"
        ' DUPLEX
        Case InStr(1, UCase(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BEDTYPE")).Value), "DUPLEX") > 0 _
            Or InStr(1, UCase(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BEDTYPE")).Value), "DUP") > 0
            unitTitle = "Duplex"
        ' APARTMENTS
        Case InStr(1, UCase(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BEDTYPE")).Value), "APARTMENT") > 0 _
            Or InStr(1, UCase(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BEDTYPE")).Value), "APT") > 0
            unitTitle = "Apartment"
    
    End Select
    
    With wsStats.Range(wsStats.Cells(iStats, 1), wsStats.Cells(iStats, 3))
        .Merge
        .Value = unitTitle & " Type " & unitType
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With
    iStats = iStats + 1

    
    ' Sub-header row
    unitBeds = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BEDS")).Value
    unitPers = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "PERS")).Value
    wsStats.Cells(iStats, 1).Value = unitBeds & "Bed / " & unitPers & "P " & unitTitle
    wsStats.Cells(iStats, 2).Value = "Target"
    wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
    wsStats.Cells(iStats, 3).Value = "Proposed"
    iStats = iStats + 1

    ' Gross Floor Area
    wsStats.Cells(iStats, 1).Value = "Gross Floor Area - sqm"
    wsStats.Cells(iStats, 2).Value = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "minAREA")).Value
    wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
    wsStats.Cells(iStats, 3).Value = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "GIFA")).Value
    iStats = iStats + 1

    ' Private Amenity Area
    wsStats.Cells(iStats, 1).Value = "Private Amenity Area - sqm"
    wsStats.Cells(iStats, 2).Value = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "minPAS")).Value
    wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
    wsStats.Cells(iStats, 3).Value = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "PAS")).Value
    iStats = iStats + 1

    ' Main Living Area
    wsStats.Cells(iStats, 1).Value = "Main Living Room - sqm"
    If wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "PERS")).Value > 5 Then
        wsStats.Cells(iStats, 2).Value = 15
    ElseIf wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "PERS")).Value > 2 Then
        wsStats.Cells(iStats, 2).Value = 13
    Else
        wsStats.Cells(iStats, 2).Value = 11
    End If



    wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
    iStats = iStats + 1

    ' Aggregate Living Area
    wsStats.Cells(iStats, 1).Value = "Aggregate Living Area - sqm"
    wsStats.Cells(iStats, 2).Value = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "MINLVNG")).Value
    wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
    wsStats.Cells(iStats, 3).Value = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "LVNG")).Value
    iStats = iStats + 1

    ' Aggregate Bedroom Area
    wsStats.Cells(iStats, 1).Value = "Aggregate Bedroom Area - sqm"
    unitBedroomArea = Val(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED1")).Value) + _
                      Val(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED2")).Value) + _
                      Val(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED3")).Value) + _
                      Val(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED4")).Value)

    wsStats.Cells(iStats, 2).Value = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "MINAGBED")).Value
    wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
    wsStats.Cells(iStats, 3).Value = unitBedroomArea
    iStats = iStats + 1

    'Individual Bedroom Areas
    If Val(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED1")).Value) > 0 Then
        wsStats.Cells(iStats, 1).Value = "Main Bedroom - sqm"
        wsStats.Cells(iStats, 2).Value = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "MINBED1")).Value
        wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
        wsStats.Cells(iStats, 3).Value = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED1")).Value
        iStats = iStats + 1
    
    End If
    If Val(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED2")).Value) > 0 Then
        If Val(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED2")).Value) >= 11.4 Then
            wsStats.Cells(iStats, 1).Value = "Double Bedroom Area - sqm"
        Else
            wsStats.Cells(iStats, 1).Value = "Single Bedroom Area - sqm"
        End If

        wsStats.Cells(iStats, 2).Value = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "MINBED2")).Value
        wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
        wsStats.Cells(iStats, 3).Value = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED2")).Value
        iStats = iStats + 1
        If Val(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED2")).Value) >= 11.4 Then
            wsStats.Cells(iStats, 1).Value = "Double Bedroom Width - m"
            wsStats.Cells(iStats, 2).Value = "2.8"
            wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
        Else
            wsStats.Cells(iStats, 1).Value = "Single Bedroom Width - m"
            wsStats.Cells(iStats, 2).Value = "2.1"
            wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
        End If
        iStats = iStats + 1
    End If
    If Val(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED3")).Value) > 0 Then
        If Val(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED3")).Value) >= 11.4 Then
            wsStats.Cells(iStats, 1).Value = "Double Bedroom Area - sqm"
        Else
            wsStats.Cells(iStats, 1).Value = "Single Bedroom Area - sqm"
        End If
        wsStats.Cells(iStats, 2).Value = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "MINBED3")).Value
        wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
        wsStats.Cells(iStats, 3).Value = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED3")).Value
        iStats = iStats + 1
        If Val(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED3")).Value) >= 11.4 Then
            wsStats.Cells(iStats, 1).Value = "Double Bedroom Width - m"
            wsStats.Cells(iStats, 2).Value = "2.8"
            wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
        Else
            wsStats.Cells(iStats, 1).Value = "Single Bedroom Width - m"
            wsStats.Cells(iStats, 2).Value = "2.1"
            wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
        End If
        iStats = iStats + 1
    End If
    If Val(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED4")).Value) > 0 Then
        If Val(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED4")).Value) >= 11.4 Then
            wsStats.Cells(iStats, 1).Value = "Double Bedroom Area - sqm"
        Else
            wsStats.Cells(iStats, 1).Value = "Single Bedroom Area - sqm"
        End If
        wsStats.Cells(iStats, 2).Value = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "MINBED4")).Value
        wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
        wsStats.Cells(iStats, 3).Value = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED4")).Value
        iStats = iStats + 1
        If Val(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "BED4")).Value) >= 11.4 Then
            wsStats.Cells(iStats, 1).Value = "Double Bedroom Width - m"
            wsStats.Cells(iStats, 2).Value = "2.8"
            wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
        Else
            wsStats.Cells(iStats, 1).Value = "Single Bedroom Width - m"
            wsStats.Cells(iStats, 2).Value = "2.1"
            wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
        End If
        iStats = iStats + 1
    End If
    If Val(wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "STOR")).Value) > 0 Then
        wsStats.Cells(iStats, 1).Value = "Min. Storage Space"
        wsStats.Cells(iStats, 2).Value = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "MINSTOR")).Value
        wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
        wsStats.Cells(iStats, 3).Value = wsTypes.Cells(i, GetColByHeader(typeHeaderMap, "STOR")).Value
        iStats = iStats + 1
    End If



    Set rng = wsStats.Range(wsStats.Cells(unitStartRow, 1), wsStats.Cells(iStats - 1, 3))
    Call DrawTableWithHeader(rng)


    iStats = iStats + 2
Next i

    With wsStats
        columns(1).AutoFit
        columns(2).AutoFit
        columns(3).AutoFit
        columns(2).HorizontalAlignment = xlCenter
        columns(3).HorizontalAlignment = xlCenter
    End With
    
    
    
End Sub

Sub copyCellsByHeader(wsSource As Worksheet, wsDest As Worksheet, srcRow As Long, targetRow As Long, srcHeaderMap As Object, targetHeaderMap As Object)
    
    Dim key As Variant
    Dim srcCol As Long
    Dim tgtCol As Long
    Dim matchCount As Long
    
    ' --- 1. VALIDATE OBJECTS ---
    If wsSource Is Nothing Or wsDest Is Nothing Then
        Debug.Print "Error: Worksheet object is Nothing."
        Exit Sub
    End If
    
    If srcHeaderMap Is Nothing Or targetHeaderMap Is Nothing Then
        Debug.Print "Error: Header map dictionary is Nothing."
        Exit Sub
    End If
    
    ' --- 2. VALIDATE ROWS ---
    If srcRow < 1 Or targetRow < 1 Or srcRow > 1048576 Or targetRow > 1048576 Then
        Debug.Print "Error: Row numbers out of range."
        Exit Sub
    End If
    
    ' --- 3. PERFORMANCE SETTINGS ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler
    
    ' --- 4. LOOP AND COPY ---
    For Each key In srcHeaderMap.Keys
        ' Only copy if header exists in both maps
        If targetHeaderMap.Exists(key) Then
            
            ' Convert dictionary values (e.g., "A" or 1) to Column Numbers
            srcCol = ColumnToNumber(srcHeaderMap(key))
            tgtCol = ColumnToNumber(targetHeaderMap(key))
            
            ' Validate columns are within Excel bounds (1 to 16,384)
            If srcCol > 0 And tgtCol > 0 And srcCol <= 16384 And tgtCol <= 16384 Then
                wsDest.Cells(targetRow, tgtCol).Value2 = wsSource.Cells(srcRow, srcCol).Value2
                matchCount = matchCount + 1
            Else
                Debug.Print "Skipping key '" & key & "': Invalid column mapping."
            End If
        End If
    Next key
    
    ' --- 5. RESTORE SETTINGS ---
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Debug.Print "Copy complete: " & matchCount & " columns matched."
    Exit Sub

ErrorHandler:
    Debug.Print "Runtime Error " & Err.Number & ": " & Err.Description
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Function ColumnToNumber(colInput As String) As Long
    Dim i As Long
    Dim result As Long
    Dim char As String
    
    If Trim(colInput) = "" Then
        ColumnToNumber = 0
        Exit Function
    End If
    
    colInput = Trim(colInput)
    
    ' If numeric, return as-is
    If IsNumeric(colInput) Then
        ColumnToNumber = CLng(colInput)
        Exit Function
    End If
    
    ' If alphabetic, convert
    colInput = UCase(colInput)
    result = 0
    
    For i = 1 To Len(colInput)
        char = Mid(colInput, i, 1)
        If char < "A" Or char > "Z" Then
            ColumnToNumber = 0 ' Invalid character found
            Exit Function
        End If
        result = result * 26 + (Asc(char) - 64)
    Next i
    
    ColumnToNumber = result
End Function

Function ColumnToLetter(colInput As Variant) As String
    Dim colNum As Long
    Dim result As String
    Dim tempNum As Long
    
    ' --- 1. HANDLE EMPTY OR NULL INPUT ---
    If IsEmpty(colInput) Or IsNull(colInput) Then
        ColumnToLetter = ""
        Exit Function
    End If
    
    ' --- 2. CONVERT TO STRING FOR VALIDATION ---
    Dim strInput As String
    strInput = Trim(CStr(colInput))
    
    If strInput = "" Then
        ColumnToLetter = ""
        Exit Function
    End If
    
    ' --- 3. DETERMINE IF INPUT IS NUMERIC OR ALPHABETIC ---
    If IsNumeric(strInput) Then
        colNum = CLng(strInput)
        
        ' Validate column number range (1 to 16,384)
        If colNum < 1 Or colNum > 16384 Then
            ColumnToLetter = ""
            Exit Function
        End If
    Else
        ' Convert letter(s) to number first
        colNum = ColLetterToNumber(strInput)
        
        ' If conversion failed (returned 0), invalid input
        If colNum = 0 Then
            ColumnToLetter = ""
            Exit Function
        End If
    End If
    
    ' --- 4. CONVERT COLUMN NUMBER TO LETTER (Math-Based) ---
    result = ""
    tempNum = colNum
    
    Do While tempNum > 0
        tempNum = tempNum - 1
        result = Chr(65 + (tempNum Mod 26)) & result
        tempNum = tempNum \ 26
    Loop
    
    ColumnToLetter = result
End Function

' Build a dictionary mapping uppercase header names to column numbers
Function BuildHeaderMap(ws As Worksheet, headerRow As Long) As Object
    Dim headerMap As Object
    Set headerMap = CreateObject("Scripting.Dictionary")
    
    Dim col As Long
    Dim headerName As String
    
    For col = 1 To 100
        If ws.Cells(headerRow, col).Value <> "" Then
            headerName = UCase(Trim(ws.Cells(headerRow, col).Value))
            headerMap.Add headerName, col
        Else
            
        End If
    Next col
    
    Set BuildHeaderMap = headerMap
End Function

Function GetLastColumnFromHeaderMap(headerMap As Object) As Long
    Dim key As Variant
    Dim colNum As Long
    Dim maxCol As Long
    Dim colValue As Variant
    
    ' --- 1. VALIDATE DICTIONARY ---
    If headerMap Is Nothing Then
        GetLastColumnFromHeaderMap = 0
        Exit Function
    End If
    
    If headerMap.Count = 0 Then
        GetLastColumnFromHeaderMap = 0
        Exit Function
    End If
    
    ' --- 2. LOOP THROUGH ALL VALUES ---
    maxCol = 0
    
    For Each key In headerMap.Keys
        colValue = headerMap(key)
        
        ' Convert to column number (handles letters "A" or numbers 1)
        colNum = colValue
        
        ' Track maximum
        If colNum > maxCol Then
            maxCol = colNum
        End If
    Next key
    
    ' --- 3. RETURN RESULT ---
    GetLastColumnFromHeaderMap = maxCol
End Function

Sub FormatColumnsByPattern(ws As Worksheet, headerMap As Object, _
                          searchPattern As String, _
                          Optional applyColor As Boolean = True, _
                          Optional applyBold As Boolean = True, _
                          Optional fontColor As Variant = -1, _
                          Optional headerRow As Long = 1)
    
    Dim key As Variant
    Dim colNum As Long
    Dim colValue As Variant
    Dim formatCount As Long
    Dim actualColor As Long
    
    ' --- 1. SET DEFAULT COLOR IF NOT PROVIDED ---
    If fontColor = -1 Then
        actualColor = RGB(83, 141, 213)  ' Blue (default)
    Else
        actualColor = fontColor
    End If
    
    ' --- 2. VALIDATE INPUTS ---
    If ws Is Nothing Or headerMap Is Nothing Then
        Debug.Print "Error: Worksheet or HeaderMap is Nothing."
        Exit Sub
    End If
    
    If headerMap.Count = 0 Then
        Debug.Print "Warning: HeaderMap is empty."
        Exit Sub
    End If
    
    If Len(Trim(searchPattern)) = 0 Then
        Debug.Print "Error: Search pattern is empty."
        Exit Sub
    End If
    
    ' --- 3. LOOP THROUGH DICTIONARY ---
    formatCount = 0
    
    For Each key In headerMap.Keys
        If InStr(1, CStr(key), searchPattern, vbTextCompare) > 0 Then
            
            'colValue = headerMap(key)
            'colNum = ColumnToNumber(colValue)
            colNum = headerMap(key)
            
            If colNum > 0 And colNum <= 16384 Then
                With ws.columns(colNum)
                    If applyBold Then
                        .Font.Bold = True
                    End If
                    
                    If applyColor Then
                        .Font.Color = actualColor
                    End If
                End With
                
                formatCount = formatCount + 1
                Debug.Print "Formatted column " & colNum & " for key: " & key
            End If
        End If
    Next key
    
    Debug.Print "Format complete: " & formatCount & " columns formatted for pattern '" & searchPattern & "'"
End Sub

Sub ApplyDwellingLookup(wsData As Worksheet, _
    wsTemplate As Worksheet, _
    rowNum As Long, _
    lookupRange As String, _
    rngRow As Range, _
    headerMap As Object, _
    tempHeaderMap As Object)

    Dim bedCount As Long
    Dim personCount As Long
    Dim lookupKey As String
    Dim foundRow As Range
    Dim tallyCol As Long

    bedCount = wsData.Cells(rowNum, GetColByHeader(headerMap, "BEDS")).Value
    personCount = wsData.Cells(rowNum, GetColByHeader(headerMap, "PERS")).Value

    lookupKey = bedCount & "b " & personCount & "p"

    Set foundRow = wsTemplate.Range(lookupRange).Find( _
                        What:=lookupKey, _
                        LookAt:=xlWhole, _
                        MatchCase:=False)

    If foundRow Is Nothing Then
        rngRow.Interior.Color = RGB(255, 0, 0)
        Exit Sub
    End If

    ' Apply template colour
    rngRow.Interior.Color = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "COLOUR")).Interior.Color

    ' Set minimums
    If headerMap.Exists("MINAREA") Then
        wsData.Cells(rowNum, headerMap("MINAREA")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINAREA")).Value ' Min Area
    End If
    If headerMap.Exists("MINPAS") Then
        wsData.Cells(rowNum, headerMap("MINPAS")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINPAS")).Value ' Min PAS
    End If
    If headerMap.Exists("MINCAS") Then
        wsData.Cells(rowNum, headerMap("MINCAS")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINCAS")).Value ' Min CAS
    End If
    If headerMap.Exists("MINAGBED") Then
        wsData.Cells(rowNum, headerMap("MINAGBED")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINAGBED")).Value ' Min Agregate Bedroom Area
    End If
    If headerMap.Exists("MINLVNG") Then
        wsData.Cells(rowNum, headerMap("MINLVNG")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINLVNG")).Value ' Min Living Area
    End If
    If headerMap.Exists("MINSTOR") Then
        wsData.Cells(rowNum, headerMap("MINSTOR")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINSTOR")).Value ' Min Storage Area
    End If
    If headerMap.Exists("MINBED1") Then
        wsData.Cells(rowNum, headerMap("MINBED1")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINBED1")).Value ' Min Bedroom 1 Area
    End If
    If headerMap.Exists("MINBED2") Then
        wsData.Cells(rowNum, headerMap("MINBED2")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINBED2")).Value ' Min Bedroom 2 Area
    End If
    If headerMap.Exists("MINBED3") Then
        wsData.Cells(rowNum, headerMap("MINBED3")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINBED3")).Value ' Min Bedroom 3 Area
    End If
    If headerMap.Exists("MINBED4") Then
        wsData.Cells(rowNum, headerMap("MINBED4")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINBED4")).Value ' Min Bedroom 4 Area
    End If
    If headerMap.Exists("MINBED5") Then
        wsData.Cells(rowNum, headerMap("MINBED5")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINBED5")).Value ' Min Bedroom 5 Area
    End If
    If headerMap.Exists("MINMAIN") Then
        wsData.Cells(rowNum, headerMap("MINMAIN")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINMAIN")).Value ' Min Main Area
    End If

End Sub

Sub drawBorderThickOutline(rng As Range)
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    ' Add a thick exterior border
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
End Sub

Sub DrawTableWithHeader(rng As Range)

    Dim topRow As Range
    Set topRow = rng.rows(1)
    topRow.Interior.Color = RGB(191, 191, 191)
    topRow.Font.Bold = True
    topRow.Font.Color = RGB(0, 0, 0)
    topRow.HorizontalAlignment = xlCenter
    topRow.VerticalAlignment = xlCenter
    
    Set topRow = rng.rows(2)
    topRow.Interior.Color = RGB(217, 217, 217)

    Dim dataRange As Range

    'add grid to everythign except the headers
    If rng.rows.Count > 2 Then
        Set dataRange = rng.Offset(2, 0).Resize(rng.rows.Count - 2, rng.columns.Count)
        
        With dataRange.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With dataRange.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With dataRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    End If

        ' Add a thick exterior border
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With

End Sub

Function GetColByHeader(headerMap As Object, headerName As String) As Long
    If headerMap.Exists(UCase(headerName)) Then
        GetColByHeader = headerMap(UCase(headerName))
    Else
        GetColByHeader = 0
    End If
End Function

Function ColLetterToNumber(colInput As String) As Long
    Dim i As Long
    Dim result As Long
    Dim char As String
    
    If Trim(colInput) = "" Then
        ColLetterToNumber = 0
        Exit Function
    End If
    
    colInput = Trim(colInput)
    
    ' If numeric, return as-is
    If IsNumeric(colInput) Then
        ColLetterToNumber = CLng(colInput)
        Exit Function
    End If
    
    ' If alphabetic, convert
    colInput = UCase(colInput)
    result = 0
    
    For i = 1 To Len(colInput)
        char = Mid(colInput, i, 1)
        If char < "A" Or char > "Z" Then
            ColLetterToNumber = 0 ' Invalid character found
            Exit Function
        End If
        result = result * 26 + (Asc(char) - 64)
    Next i
    
    ColLetterToNumber = result
End Function

