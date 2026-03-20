Option Explicit

' ============================================================================
' UNIT STATS MODULE - STANDALONE
' Generates unit statistics by reading directly from wsSource
' Can be run independently without Unit Types module
' ============================================================================

Sub GenerateUnitStats()
    Dim wsSource As Worksheet
    Dim wsStats As Worksheet
    Dim wsTemplate As Worksheet
    Dim sourceHeaderMap As Object
    Dim typeDict As Object
    Dim lastRow As Long
    Dim i As Long
    Dim currentDate As String
    Dim ws As Worksheet
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' Delete any existing Stats sheets
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, "Stats", vbTextCompare) > 0 Then
            ws.Delete
        End If
    Next ws
    
    ' Set source worksheet
    Set wsSource = ThisWorkbook.Sheets("sourceData")
    
    ' Set template worksheet
    Set wsTemplate = ThisWorkbook.Sheets("template")
    
    ' Build header maps
    Set sourceHeaderMap = BuildHeaderMap(wsSource, 1)

    ' Create stats worksheet
    currentDate = Format(Date, "yy-mm-dd")
    Set wsStats = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsStats.Name = "Stats " & currentDate
    
    ' Find unique unit types from wsSource
    Set typeDict = CreateObject("Scripting.Dictionary")
    Call FindUniqueUnitTypes(wsSource, sourceHeaderMap, wsTemplate, typeDict)

    ' Generate stats tables
    Call BuildStatsTables(wsSource, wsStats, wsTemplate, sourceHeaderMap, typeDict)
    
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
End Sub

' ============================================================================
' Find unique unit types from wsSource using regex pattern from template
' ============================================================================
Sub FindUniqueUnitTypes(wsSource As Worksheet, sourceHeaderMap As Object, _
                       wsTemplate As Worksheet, typeDict As Object)
    
    Dim reTypes As Object
    Set reTypes = CreateObject("VBScript.RegExp")
    
    Dim regexPattern As String
    regexPattern = Trim(wsTemplate.Range("X3").Value)
    
    If Len(regexPattern) = 0 Then
        regexPattern = ".*"
    End If
    
    With reTypes
        .Global = False
        .IgnoreCase = True
        .Pattern = regexPattern
    End With
    
    Dim lastRow As Long
    Dim i As Long
    Dim unitType As String
    Dim unitKey As String
    Dim tempArr As Variant
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        unitType = wsSource.Cells(i, GetColByHeader(sourceHeaderMap, "TYPE")).Value
        
        If Len(unitType) > 0 And reTypes.Test(unitType) Then
            unitKey = UCase(Trim(reTypes.Execute(unitType)(0)))
            
            If Not typeDict.Exists(unitKey) Then
                typeDict.Add unitKey, Array(1, i)
            Else
                tempArr = typeDict(unitKey)
                tempArr(0) = tempArr(0) + 1
                typeDict(unitKey) = tempArr
            End If
        End If
    Next i
    
End Sub

' ============================================================================
' Build stats tables for each unique unit type
' ============================================================================
Sub BuildStatsTables(wsSource As Worksheet, wsStats As Worksheet, _
                    wsTemplate As Worksheet, sourceHeaderMap As Object, _
                    typeDict As Object)

    Dim iStats As Long
    Dim unitStartRow As Long
    Dim unitTitle As String
    Dim unitType As String
    Dim unitBeds As String
    Dim unitPers As String
    Dim unitBedroomArea As Double
    Dim rng As Range
    Dim key As Variant
    Dim typeItems As Variant
    Dim sourceRow As Long
    Dim totalUnits As Long
    Dim dwellingType As String
    Dim bedCount As Long
    Dim personCount As Long

    iStats = 2

    ' Calculate total units
    For Each key In typeDict.Keys
        totalUnits = totalUnits + typeDict(key)(0)
    Next key

    ' Loop through each unique unit type
    For Each key In typeDict.Keys
        unitStartRow = iStats
        typeItems = typeDict(key)
        sourceRow = typeItems(1)

        ' Get unit type from source
        unitType = key

        ' Determine unit title and dwelling type from source
        dwellingType = GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BEDTYPE")
        unitTitle = GetUnitTitle(wsSource, sourceRow, sourceHeaderMap)

        ' Get beds and persons for lookup
        bedCount = Val(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BEDS"))
        personCount = Val(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "PERS"))

        ' Title row
        With wsStats.Range(wsStats.Cells(iStats, 1), wsStats.Cells(iStats, 3))
            .Merge
            .Value = unitTitle & " Type " & unitType
            .Font.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
        End With
        iStats = iStats + 1

        unitBeds = bedCount
        unitPers = personCount

        ' Sub-header row
        wsStats.Cells(iStats, 1).Value = unitBeds & "Bed / " & unitPers & "P " & unitTitle
        wsStats.Cells(iStats, 2).Value = "Target"
        wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
        wsStats.Cells(iStats, 3).Value = "Proposed"
        iStats = iStats + 1

        ' Gross Floor Area
        wsStats.Cells(iStats, 1).Value = "Gross Floor Area - sqm"
        wsStats.Cells(iStats, 2).Value = GetLookupValue(wsTemplate, dwellingType, bedCount, personCount, "MINAREA")
        wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
        wsStats.Cells(iStats, 3).Value = GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "GIFA")
        iStats = iStats + 1

        ' Private Amenity Area
        wsStats.Cells(iStats, 1).Value = "Private Amenity Area - sqm"
        wsStats.Cells(iStats, 2).Value = GetLookupValue(wsTemplate, dwellingType, bedCount, personCount, "MINPAS")
        wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
        wsStats.Cells(iStats, 3).Value = GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "PAS")
        iStats = iStats + 1

        ' Main Living Area
        wsStats.Cells(iStats, 1).Value = "Main Living Room - sqm"
        wsStats.Cells(iStats, 2).Value = GetLookupValue(wsTemplate, dwellingType, bedCount, personCount, "MINMAIN")
        wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
        iStats = iStats + 1

        ' Aggregate Living Area
        wsStats.Cells(iStats, 1).Value = "Aggregate Living Area - sqm"
        wsStats.Cells(iStats, 2).Value = GetLookupValue(wsTemplate, dwellingType, bedCount, personCount, "MINLVNG")
        wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
        wsStats.Cells(iStats, 3).Value = GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "LVNG")
        iStats = iStats + 1

        ' Aggregate Bedroom Area
        wsStats.Cells(iStats, 1).Value = "Aggregate Bedroom Area - sqm"
        unitBedroomArea = Val(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED1")) + _
                        Val(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED2")) + _
                        Val(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED3")) + _
                        Val(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED4"))

        wsStats.Cells(iStats, 2).Value = GetLookupValue(wsTemplate, dwellingType, bedCount, personCount, "MINAGBED")
        wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
        wsStats.Cells(iStats, 3).Value = unitBedroomArea
        iStats = iStats + 1

        ' Individual Bedroom Areas
        If Val(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED1")) > 0 Then
            wsStats.Cells(iStats, 1).Value = "Main Bedroom - sqm"
            wsStats.Cells(iStats, 2).Value = GetLookupValue(wsTemplate, dwellingType, bedCount, personCount, "MINBED1")
            wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
            wsStats.Cells(iStats, 3).Value = GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED1")
            iStats = iStats + 1
        End If

        If Val(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED2")) > 0 Then
            If Val(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED2")) >= 11.4 Then
                wsStats.Cells(iStats, 1).Value = "Double Bedroom Area - sqm"
            Else
                wsStats.Cells(iStats, 1).Value = "Single Bedroom Area - sqm"
            End If

            wsStats.Cells(iStats, 2).Value = GetLookupValue(wsTemplate, dwellingType, bedCount, personCount, "MINBED2")
            wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
            wsStats.Cells(iStats, 3).Value = GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED2")
            iStats = iStats + 1

            If Val(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED2")) >= 11.4 Then
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

        If Val(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED3")) > 0 Then
            If Val(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED3")) >= 11.4 Then
                wsStats.Cells(iStats, 1).Value = "Double Bedroom Area - sqm"
            Else
                wsStats.Cells(iStats, 1).Value = "Single Bedroom Area - sqm"
            End If
            wsStats.Cells(iStats, 2).Value = GetLookupValue(wsTemplate, dwellingType, bedCount, personCount, "MINBED3")
            wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
            wsStats.Cells(iStats, 3).Value = GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED3")
            iStats = iStats + 1

            If Val(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED3")) >= 11.4 Then
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

        If Val(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED4")) > 0 Then
            If Val(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED4")) >= 11.4 Then
                wsStats.Cells(iStats, 1).Value = "Double Bedroom Area - sqm"
            Else
                wsStats.Cells(iStats, 1).Value = "Single Bedroom Area - sqm"
            End If
            wsStats.Cells(iStats, 2).Value = GetLookupValue(wsTemplate, dwellingType, bedCount, personCount, "MINBED4")
            wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
            wsStats.Cells(iStats, 3).Value = GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED4")
            iStats = iStats + 1

            If Val(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BED4")) >= 11.4 Then
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

        If Val(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "STOR")) > 0 Then
            wsStats.Cells(iStats, 1).Value = "Min. Storage Space"
            wsStats.Cells(iStats, 2).Value = GetLookupValue(wsTemplate, dwellingType, bedCount, personCount, "MINSTOR")
            wsStats.Cells(iStats, 2).Font.Color = RGB(83, 141, 213)
            wsStats.Cells(iStats, 3).Value = GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "STOR")
            iStats = iStats + 1
        End If

        ' Draw table with header
        Set rng = wsStats.Range(wsStats.Cells(unitStartRow, 1), wsStats.Cells(iStats - 1, 3))
        Call DrawTableWithHeader(rng)

        iStats = iStats + 2
    Next key

    ' Auto-fit columns
    With wsStats
        .Columns(1).AutoFit
        .Columns(2).AutoFit
        .Columns(3).AutoFit
        .Columns(2).HorizontalAlignment = xlCenter
        .Columns(3).HorizontalAlignment = xlCenter
    End With

End Sub

' ============================================================================
' Helper function to get unit title based on dwelling type
' ============================================================================
Function GetUnitTitle(wsSource As Worksheet, sourceRow As Long, _
                     sourceHeaderMap As Object) As String
    
    Dim bedType As String
    bedType = UCase(GetSourceValue(wsSource, sourceRow, sourceHeaderMap, "BEDTYPE"))
    
    Select Case True
        Case InStr(1, bedType, "HOUSE") > 0
            GetUnitTitle = "House"
        Case InStr(1, bedType, "DUPLEX") > 0 Or InStr(1, bedType, "DUP") > 0
            GetUnitTitle = "Duplex"
        Case InStr(1, bedType, "APARTMENT") > 0 Or InStr(1, bedType, "APT") > 0
            GetUnitTitle = "Apartment"
        Case Else
            GetUnitTitle = "Unit"
    End Select
    
End Function

' ============================================================================
' Helper function to safely get value from source by header name
' ============================================================================
Function GetSourceValue(wsSource As Worksheet, sourceRow As Long, _
                       sourceHeaderMap As Object, headerName As String) As Variant
    
    Dim colNum As Long
    colNum = GetColByHeader(sourceHeaderMap, headerName)
    
    If colNum >= 1 And colNum <= 16384 Then
        GetSourceValue = wsSource.Cells(sourceRow, colNum).Value
    Else
        GetSourceValue = ""
    End If
    
End Function
