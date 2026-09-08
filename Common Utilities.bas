Option Explicit

' ============================================================================
' COMMON UTILITY FUNCTIONS FOR APARTMENT SCHEDULE
' Shared helper functions used across multiple modules
' ============================================================================

' ----------------------------------------------------------------------------
' Column Conversion Functions
' ----------------------------------------------------------------------------

' Convert column letter to number (e.g., "A" -> 1, "Z" -> 26, "AA" -> 27)
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

' Convert column number to letter (e.g., 1 -> "A", 26 -> "Z", 27 -> "AA")
Function ColumnToLetter(colInput As Variant) As String
    Dim colNum As Long
    Dim result As String
    Dim tempNum As Long

    ' Handle empty or null input
    If IsEmpty(colInput) Or IsNull(colInput) Then
        ColumnToLetter = ""
        Exit Function
    End If

    ' Convert to string for validation
    Dim strInput As String
    strInput = Trim(CStr(colInput))

    If strInput = "" Then
        ColumnToLetter = ""
        Exit Function
    End If

    ' Determine if input is numeric or alphabetic
    If IsNumeric(strInput) Then
        colNum = CLng(strInput)
        ' Validate column number range (1 to 16,384)
        If colNum < 1 Or colNum > 16384 Then
            ColumnToLetter = ""
            Exit Function
        End If
    Else
        ' Convert letter(s) to number first
        colNum = ColumnToNumber(strInput)
        ' If conversion failed (returned 0), invalid input
        If colNum = 0 Then
            ColumnToLetter = ""
            Exit Function
        End If
    End If

    ' Convert column number to letter (math-based)
    result = ""
    tempNum = colNum

    Do While tempNum > 0
        tempNum = tempNum - 1
        result = Chr(65 + (tempNum Mod 26)) & result
        tempNum = tempNum \ 26
    Loop

    ColumnToLetter = result
End Function

' ----------------------------------------------------------------------------
' Text Import Functions
' ----------------------------------------------------------------------------

' Build the TextFileColumnDataTypes array for a delimited text/CSV QueryTable
' import. Every column defaults to xlGeneralFormat except the ones whose header
' matches an entry in textHeaders, which are forced to xlTextFormat. Without
' this Excel parses values like "1E4" as 1 x 10^4 and stores the number 10000.
' Returns Empty when the header line cannot be read - callers should then leave
' TextFileColumnDataTypes alone and let Excel guess as before.
Function BuildTextImportColumnTypes(filePath As String, isCsv As Boolean, _
                                    textHeaders As Variant) As Variant
    Dim fileNum As Integer
    Dim headerLine As String
    Dim fields As Variant
    Dim colTypes() As Variant
    Dim delimiter As String
    Dim headerName As String
    Dim i As Long
    Dim j As Long

    BuildTextImportColumnTypes = Empty

    If Len(Trim(filePath)) = 0 Then Exit Function
    If IsEmpty(textHeaders) Then Exit Function

    On Error GoTo ErrorHandler

    ' Read the first non-blank line - that is the header row the import will use
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, headerLine
        If Len(Trim(headerLine)) > 0 Then Exit Do
    Loop
    Close #fileNum

    If Len(Trim(headerLine)) = 0 Then Exit Function

    ' Drop a UTF-8 BOM so the first header still matches
    If Left(headerLine, 3) = Chr(239) & Chr(187) & Chr(191) Then
        headerLine = Mid(headerLine, 4)
    End If

    If isCsv Then
        delimiter = ","
    Else
        delimiter = vbTab
    End If

    fields = Split(headerLine, delimiter)
    ReDim colTypes(LBound(fields) To UBound(fields))

    For i = LBound(fields) To UBound(fields)
        colTypes(i) = xlGeneralFormat

        headerName = Trim(fields(i))
        ' Strip the text qualifier quotes the import strips anyway
        If Len(headerName) >= 2 Then
            If Left(headerName, 1) = """" And Right(headerName, 1) = """" Then
                headerName = Trim(Mid(headerName, 2, Len(headerName) - 2))
            End If
        End If
        headerName = UCase(headerName)

        For j = LBound(textHeaders) To UBound(textHeaders)
            If headerName = UCase(Trim(CStr(textHeaders(j)))) Then
                colTypes(i) = xlTextFormat
                Exit For
            End If
        Next j
    Next i

    BuildTextImportColumnTypes = colTypes
    Exit Function

ErrorHandler:
    Debug.Print "BuildTextImportColumnTypes failed for '" & filePath & "': " & Err.Description
    On Error Resume Next
    Close #fileNum
    BuildTextImportColumnTypes = Empty
End Function

' ----------------------------------------------------------------------------
' Header Marker Functions
' ----------------------------------------------------------------------------

' A header cell on a mapping row can be tagged with trailing marker characters
' that say what the summary rows should do with that column:
'
'   sigma  - total the column     e.g. a cell reading GIFA followed by a sigma
'   %      - percentage of total  e.g. a cell reading DUAL followed by a %
'
' Both can be combined on one cell, in either order. The mapping row (row 9 for
' the schedule, row 29 for the types block) is only ever read, never copied to
' the output, so the markers cost nothing on the printed sheet.
'
' The sigma is matched by code point rather than written literally: VBA exports
' .bas files as Windows-1252, which has no sigma, so a literal one would not
' survive an export/import round trip.
Private Function IsSumMarkerChar(ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function

    Select Case AscW(ch)
        Case &H3A3   ' greek capital letter sigma
            IsSumMarkerChar = True
        Case &H3C3   ' greek small letter sigma
            IsSumMarkerChar = True
        Case &H2211  ' n-ary summation, what Insert > Symbol tends to give
            IsSumMarkerChar = True
    End Select
End Function

Private Function IsPercentMarkerChar(ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function

    Select Case AscW(ch)
        Case &H25    ' percent sign
            IsPercentMarkerChar = True
        Case &HFF05  ' fullwidth percent sign
            IsPercentMarkerChar = True
    End Select
End Function

Private Function IsMarkerChar(ch As String) As Boolean
    IsMarkerChar = IsSumMarkerChar(ch) Or IsPercentMarkerChar(ch)
End Function

' Split a header cell into the column name and the run of markers at its end
Private Sub SplitHeaderMarkers(ByVal headerText As String, _
                               ByRef headerName As String, _
                               ByRef markers As String)
    Dim s As String

    s = Trim(headerText)
    markers = ""

    ' Len(s) > 1 stops the whole cell being eaten: a header that is nothing but
    ' a marker is a column literally named "%" - the types block has one - not a
    ' tagged column
    Do While Len(s) > 1
        If Not IsMarkerChar(Right(s, 1)) Then Exit Do
        markers = Right(s, 1) & markers
        s = Trim(Left(s, Len(s) - 1))
    Loop

    headerName = s
End Sub

' Header text with any trailing markers removed - use this everywhere a header
' cell is read, so a tagged column still matches its plain name
Function StripHeaderMarkers(headerText As String) As String
    Dim headerName As String
    Dim markers As String

    SplitHeaderMarkers headerText, headerName, markers
    StripHeaderMarkers = headerName
End Function

' True when a header cell is tagged for summing
Function HasSumMarker(headerText As String) As Boolean
    Dim headerName As String
    Dim markers As String
    Dim i As Long

    SplitHeaderMarkers headerText, headerName, markers

    For i = 1 To Len(markers)
        If IsSumMarkerChar(Mid(markers, i, 1)) Then
            HasSumMarker = True
            Exit Function
        End If
    Next i
End Function

' True when a header cell is tagged for a percentage of total
Function HasPercentMarker(headerText As String) As Boolean
    Dim headerName As String
    Dim markers As String
    Dim i As Long

    SplitHeaderMarkers headerText, headerName, markers

    For i = 1 To Len(markers)
        If IsPercentMarkerChar(Mid(markers, i, 1)) Then
            HasPercentMarker = True
            Exit Function
        End If
    Next i
End Function

' Columns tagged with one kind of marker on a header row, left to right
Private Function BuildMarkedColumns(ws As Worksheet, headerRow As Long, _
                                    wantPercent As Boolean) As Collection
    Dim result As Collection
    Dim col As Long
    Dim cellText As String
    Dim isMarked As Boolean

    Set result = New Collection
    Set BuildMarkedColumns = result

    If ws Is Nothing Then Exit Function
    If headerRow < 1 Or headerRow > 1048576 Then Exit Function

    ' Same 1-100 sweep BuildHeaderMap uses, so the two always agree
    On Error Resume Next
    For col = 1 To 100
        cellText = CStr(ws.Cells(headerRow, col).Value)
        If wantPercent Then
            isMarked = HasPercentMarker(cellText)
        Else
            isMarked = HasSumMarker(cellText)
        End If
        If isMarked Then result.Add col
    Next col
    On Error GoTo 0
End Function

' Collect the columns tagged for summing on a header row, left to right.
' Returns an empty collection when the row carries no markers at all, which
' callers use to fall back to their own defaults - so an untagged template
' keeps behaving exactly as it did before.
Function BuildSumColumns(ws As Worksheet, headerRow As Long) As Collection
    Set BuildSumColumns = BuildMarkedColumns(ws, headerRow, False)
End Function

' Collect the columns tagged for a percentage of total, same contract
Function BuildPercentColumns(ws As Worksheet, headerRow As Long) As Collection
    Set BuildPercentColumns = BuildMarkedColumns(ws, headerRow, True)
End Function

' ----------------------------------------------------------------------------
' Header Map Functions
' ----------------------------------------------------------------------------

' Build a dictionary mapping uppercase header names to column numbers
' Scans columns 1-100 for non-empty header cells
Function BuildHeaderMap(ws As Worksheet, headerRow As Long) As Object
    Dim headerMap As Object
    Set headerMap = CreateObject("Scripting.Dictionary")

    Dim col As Long
    Dim headerName As String

    ' Validate inputs
    If ws Is Nothing Then
        Set BuildHeaderMap = headerMap
        Exit Function
    End If

    If headerRow < 1 Or headerRow > 1048576 Then
        Set BuildHeaderMap = headerMap
        Exit Function
    End If

    On Error Resume Next
    For col = 1 To 100
        If ws.Cells(headerRow, col).Value <> "" Then
            ' Strip the markers so a tagged cell maps under its plain name
            headerName = UCase(StripHeaderMarkers(CStr(ws.Cells(headerRow, col).Value)))
            If Len(headerName) > 0 Then
                If Not headerMap.Exists(headerName) Then
                    headerMap.Add headerName, col
                End If
            End If
        End If
    Next col
    On Error GoTo 0

    Set BuildHeaderMap = headerMap
End Function

' Get column number from header name using the header map
' Returns -1 if header not found (invalid column - must be checked before use)
Function GetColByHeader(headerMap As Object, headerName As String) As Long
    If headerMap Is Nothing Then
        GetColByHeader = 1
        Exit Function
    End If

    If headerMap.Exists(UCase(headerName)) Then
        GetColByHeader = headerMap(UCase(headerName))
    Else
        GetColByHeader = 1 ' Default to column 1 if header not found (or could raise an error)
    End If
End Function

' Find a column by trying several possible header spellings, case-insensitive
' (e.g. "NO", "NO.", "NUM" for a unit-number column). Returns 0 if none of
' the aliases exist in headerMap - callers should check for 0 rather than
' relying on GetColByHeader's column-1 fallback.
Function FindColumnByAliases(headerMap As Object, aliases As Variant) As Long
    Dim idx As Long

    If headerMap Is Nothing Then
        FindColumnByAliases = 0
        Exit Function
    End If

    For idx = LBound(aliases) To UBound(aliases)
        If headerMap.Exists(UCase(CStr(aliases(idx)))) Then
            FindColumnByAliases = GetColByHeader(headerMap, CStr(aliases(idx)))
            Exit Function
        End If
    Next idx

    FindColumnByAliases = 0
End Function

' Get the maximum column number from a header map
Function GetLastColumnFromHeaderMap(headerMap As Object) As Long
    Dim key As Variant
    Dim colNum As Long
    Dim maxCol As Long
    Dim colValue As Variant

    ' Validate dictionary
    If headerMap Is Nothing Then
        GetLastColumnFromHeaderMap = 0
        Exit Function
    End If

    If headerMap.Count = 0 Then
        GetLastColumnFromHeaderMap = 0
        Exit Function
    End If

    ' Loop through all values and track maximum
    maxCol = 0

    For Each key In headerMap.Keys
        colValue = headerMap(key)
        colNum = colValue

        If colNum > maxCol Then
            maxCol = colNum
        End If
    Next key

    GetLastColumnFromHeaderMap = maxCol
End Function

' Get maximum value from a dictionary
Function mapMaxValue(dict As Object) As Variant
    Dim key As Variant
    Dim maxValue As Variant
    Dim firstKey As Variant

    If dict Is Nothing Then
        mapMaxValue = 0
        Exit Function
    End If

    If dict.Count = 0 Then
        mapMaxValue = 0
        Exit Function
    End If

    firstKey = dict.Keys()(0)
    maxValue = dict(firstKey)

    For Each key In dict.Keys
        If dict(key) > maxValue Then
            maxValue = dict(key)
        End If
    Next key

    mapMaxValue = maxValue
End Function

' ----------------------------------------------------------------------------
' Date Functions
' ----------------------------------------------------------------------------

' Format date with ordinal suffix (e.g., "1st January 2024")
Function FormatDateWithSuffix(dt As Date) As String
    Dim dayNum As Integer
    Dim suffix As String

    dayNum = Day(dt)

    ' Determine the suffix
    Select Case dayNum
        Case 1, 21, 31: suffix = "st"
        Case 2, 22: suffix = "nd"
        Case 3, 23: suffix = "rd"
        Case Else: suffix = "th"
    End Select

    FormatDateWithSuffix = dayNum & suffix & " " & Format(dt, "mmmm yyyy")
End Function

' ----------------------------------------------------------------------------
' Color Functions
' ----------------------------------------------------------------------------

' Convert hex color to VBA Long color (e.g., "#FF5733" -> RGB value)
Function HEX(hexColor As String) As Long
    Dim r As Integer, g As Integer, b As Integer

    ' Remove "#" if it exists
    If Left(hexColor, 1) = "#" Then
        hexColor = Mid(hexColor, 2)
    End If

    ' Validate hex color length
    If Len(hexColor) <> 6 Then
        Err.Raise vbObjectError + 513, , "Invalid hex color format. Must be 6 characters like '#FF5733'."
    End If

    ' Convert hex to RGB
    On Error GoTo ErrorHandler
    r = CInt("&H" & Mid(hexColor, 1, 2))
    g = CInt("&H" & Mid(hexColor, 3, 2))
    b = CInt("&H" & Mid(hexColor, 5, 2))
    HEX = RGB(r, g, b)
    Exit Function

ErrorHandler:
    HEX = RGB(255, 255, 255) ' fallback to white on error
    MsgBox "Invalid HEX color: " & hexColor, vbExclamation
End Function

' ----------------------------------------------------------------------------
' Formatting Functions
' ----------------------------------------------------------------------------

' Format columns by header pattern (e.g., "MIN" formats all columns with "MIN" in header)
Sub FormatColumnsByPattern(ws As Worksheet, headerMap As Object, _
                          searchPattern As String, _
                          Optional applyColor As Boolean = True, _
                          Optional applyBold As Boolean = True, _
                          Optional fontColor As Variant = -1, _
                          Optional headerRow As Long = 1)

    Dim key As Variant
    Dim colNum As Long
    Dim formatCount As Long
    Dim actualColor As Long

    ' Set default color if not provided
    If fontColor = -1 Then
        actualColor = RGB(83, 141, 213)  ' Blue (default)
    Else
        actualColor = fontColor
    End If

    ' Validate inputs
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

    ' Loop through dictionary and format matching columns
    formatCount = 0

    For Each key In headerMap.Keys
        If InStr(1, CStr(key), searchPattern, vbTextCompare) > 0 Then
            colNum = headerMap(key)

            If colNum >= 1 And colNum <= 16384 Then
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

' ----------------------------------------------------------------------------
' Border Functions
' ----------------------------------------------------------------------------

' Draw thick outline border around a range
Sub drawBorderThickOutline(rng As Range)
    On Error Resume Next
    
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
    ' Add thick exterior border
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
    
    On Error GoTo 0
End Sub

' Draw a horizontal line across a row
Sub drawBorderLine(ws As Worksheet, i As Long, c As Long)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(i, "A"), ws.Cells(i, c))

    On Error Resume Next
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    On Error GoTo 0

    rng.Font.Bold = True
End Sub

' Draw a table with header rows (first two rows styled differently)
Sub DrawTableWithHeader(rng As Range)
    Dim topRow As Range
    Dim dataRange As Range

    On Error Resume Next

    ' First header row - darker grey
    Set topRow = rng.rows(1)
    topRow.Interior.Color = RGB(191, 191, 191)
    topRow.Font.Bold = True
    topRow.Font.Color = RGB(0, 0, 0)
    topRow.HorizontalAlignment = xlCenter
    topRow.VerticalAlignment = xlCenter

    ' Second header row - lighter grey
    Set topRow = rng.rows(2)
    topRow.Interior.Color = RGB(217, 217, 217)

    ' Add grid to everything except the headers
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

    ' Exterior border
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

    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' Formatting Copy Functions
' ----------------------------------------------------------------------------

' Copy all formatting from a source row and propagate it down to all data rows
' in the target worksheet. Copies font, interior, number format, alignment, etc.
Sub CopyRowFormattingDown(wsSource As Worksheet, sourceRow As Long, _
                          wsTarget As Worksheet, targetStartRow As Long, _
                          targetEndRow As Long, _
                          Optional startCol As Long = 1, _
                          Optional endCol As Long = 0)

    Dim col As Long
    Dim sourceRange As Range
    Dim targetRange As Range
    Dim lastCol As Long

    ' Validate worksheets
    If wsSource Is Nothing Or wsTarget Is Nothing Then
        Debug.Print "Error: Worksheet object is Nothing."
        Exit Sub
    End If

    ' Validate rows
    If sourceRow < 1 Or targetStartRow < 1 Or sourceRow > 1048576 Then
        Debug.Print "Error: Row numbers out of range."
        Exit Sub
    End If

    ' Determine end column if not specified
    If endCol = 0 Then
        lastCol = wsSource.Cells(sourceRow, wsSource.columns.Count).End(xlToLeft).Column
        If lastCol < 1 Then lastCol = wsTarget.UsedRange.columns.Count
        endCol = lastCol
    End If

    ' Optimize performance
    Application.ScreenUpdating = False

    ' Copy formatting column by column
    For col = startCol To endCol
        Set sourceRange = wsSource.Cells(sourceRow, col)

        For targetStartRow = 1 To targetEndRow
            Set targetRange = wsTarget.Cells(targetStartRow, col)

            ' Copy font properties
            With targetRange.Font
                .Name = sourceRange.Font.Name
                .Size = sourceRange.Font.Size
                .Bold = sourceRange.Font.Bold
                .Italic = sourceRange.Font.Italic
                .Underline = sourceRange.Font.Underline
                .Color = sourceRange.Font.Color
                .Strikethrough = sourceRange.Font.Strikethrough
                .Superscript = sourceRange.Font.Superscript
                .Subscript = sourceRange.Font.Subscript
                .OutlineFont = sourceRange.Font.OutlineFont
                .Shadow = sourceRange.Font.Shadow
                .Color = sourceRange.Font.Color
            End With

            ' Copy interior (fill) properties if there is a colour applies
            If sourceRange.Interior.Color <> 16777215 Then
            With targetRange.Interior
                .Color = sourceRange.Interior.Color
                .Pattern = sourceRange.Interior.Pattern
                .PatternColor = sourceRange.Interior.PatternColor
                .ThemeColor = sourceRange.Interior.ThemeColor
                .TintAndShade = sourceRange.Interior.TintAndShade
            End With
            End If

            ' Copy other formatting
            targetRange.NumberFormat = sourceRange.NumberFormat
            targetRange.HorizontalAlignment = sourceRange.HorizontalAlignment
            targetRange.VerticalAlignment = sourceRange.VerticalAlignment
            targetRange.WrapText = sourceRange.WrapText
            targetRange.Orientation = sourceRange.Orientation
            targetRange.IndentLevel = sourceRange.IndentLevel
            targetRange.ShrinkToFit = sourceRange.ShrinkToFit
            targetRange.ReadingOrder = sourceRange.ReadingOrder
        Next targetStartRow
    Next col

    ' Restore settings
    Application.ScreenUpdating = True
End Sub

' ----------------------------------------------------------------------------
' Data Copy Functions
' ----------------------------------------------------------------------------

' Copy cells from source to destination by matching headers
Sub copyCellsByHeader(wsSource As Worksheet, wsDest As Worksheet, _
                     srcRow As Long, targetRow As Long, _
                     srcHeaderMap As Object, targetHeaderMap As Object)

    Dim key As Variant
    Dim srcCol As Long
    Dim tgtCol As Long
    Dim matchCount As Long

    ' Validate objects
    If wsSource Is Nothing Or wsDest Is Nothing Then
        Debug.Print "Error: Worksheet object is Nothing."
        Exit Sub
    End If

    If srcHeaderMap Is Nothing Or targetHeaderMap Is Nothing Then
        Debug.Print "Error: Header map dictionary is Nothing."
        Exit Sub
    End If

    ' Validate rows
    If srcRow < 1 Or targetRow < 1 Or srcRow > 1048576 Or targetRow > 1048576 Then
        Debug.Print "Error: Row numbers out of range."
        Exit Sub
    End If

    ' Performance settings
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler

    ' Loop and copy
    For Each key In srcHeaderMap.Keys
        ' Only copy if header exists in both maps
        If targetHeaderMap.Exists(key) Then
            ' Convert dictionary values (e.g., "A" or 1) to Column Numbers
            srcCol = ColumnToNumber(srcHeaderMap(key))
            tgtCol = ColumnToNumber(targetHeaderMap(key))

            ' Validate columns are within Excel bounds (1 to 16,384)
            If srcCol >= 1 And tgtCol >= 1 And srcCol <= 16384 And tgtCol <= 16384 Then
                wsDest.Cells(targetRow, tgtCol).Value2 = wsSource.Cells(srcRow, srcCol).Value2
                matchCount = matchCount + 1
            Else
                Debug.Print "Skipping key '" & key & "': Invalid column mapping."
            End If
        End If
    Next key

    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Debug.Print "Copy complete: " & matchCount & " columns matched."
    Exit Sub

ErrorHandler:
    Debug.Print "Runtime Error " & Err.Number & ": " & Err.Description
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

' Copy columns from wsSource to wsDest by matching headers
Sub CopyColumnsByHeader(wsSource As Worksheet, wsDest As Worksheet, _
                       wsTemplate As Worksheet, rowSource As Long, rowDest As Long)

    Dim srcLastRow As Long
    Dim targetCol As Long
    Dim srcHeaderMap As Object
    Dim targetHeaderRange As Range
    Dim headerName As String
    Dim srcCol As Long

    ' Get the last used row in wsSource
    srcLastRow = wsSource.Cells(wsSource.rows.Count, 1).End(xlUp).row

    ' Build header map for wsSource
    Set srcHeaderMap = BuildHeaderMap(wsSource, rowSource)

    ' Define the target header range in wsTemplate
    Set targetHeaderRange = wsTemplate.Range("A" & rowDest & ":Z" & rowDest)

    ' Loop through each target column in wsTemplate
    For targetCol = 1 To targetHeaderRange.columns.Count
        ' Strip the markers too, or a tagged template column stops matching
        ' its source header and never gets imported
        headerName = UCase(StripHeaderMarkers(CStr(wsTemplate.Cells(rowDest, targetCol).Value)))

        If Len(headerName) > 0 Then
            ' Look up source column using the header map
            If srcHeaderMap.Exists(headerName) Then
                srcCol = GetColByHeader(srcHeaderMap, headerName)
            

                If srcCol >= 1 Then
                    ' Found match - copy the entire column (data only, from row 2 onwards)
                    wsSource.Range(wsSource.Cells(2, srcCol), wsSource.Cells(srcLastRow, srcCol)).Copy _
                        Destination:=wsDest.Cells(2, targetCol)
                    ' Copy the header as well
                    wsDest.Cells(1, targetCol).Value = wsSource.Cells(1, srcCol).Value
                End If
            End If
        End If
    Next targetCol
End Sub

' ----------------------------------------------------------------------------
' Sum and Calculation Functions
' ----------------------------------------------------------------------------

' Add SUM/COUNTA formulas to columns
Sub sumColumnsSub(ws As Worksheet, _
                 columns As Collection, _
                 startRow As Long, _
                 endRow As Long, _
                 colOffset As Long, _
                 Optional countFirst As Boolean = False)

    Dim P As Long
    Dim colNum As Long

    For P = 1 To columns.Count
        colNum = columns(P)

        ' Skip invalid column numbers (must be >= 1)
        If colNum < 1 Then GoTo NextCol

        If P = 1 And countFirst = True Then
            With ws.Cells(endRow, colNum + colOffset)
                .Formula = "=COUNTA(" & ColumnToLetter(colNum) & endRow - 1 & ":" & ColumnToLetter(colNum) & startRow & ")"
                .Font.Bold = True
            End With
        Else
            With ws.Cells(endRow, colNum + colOffset)
                .Formula = "=SUM(" & ColumnToLetter(colNum) & endRow - 1 & ":" & ColumnToLetter(colNum) & startRow & ")"
                .Font.Bold = True
            End With
        End If
NextCol:
    Next P
End Sub

' Add SUM formulas across multiple rows (for zone/block summaries)
Sub sumColumnsRowsSub(ws As Worksheet, columns As Collection, _
                     rows As Collection, i As Long)

    Dim P As Long
    Dim q As Long
    Dim colNum As Long
    Dim colLetter As String

    If rows.Count > 0 Then
        For P = 1 To columns.Count
            colNum = columns(P)
            colLetter = ColumnToLetter(colNum)

            ' Declare the formula and start writing it
            Dim blockFormulaString As String
            blockFormulaString = "=SUM("

            For q = 1 To rows.Count
                If q = rows.Count Then
                    ' For the last item, don't add a comma after it
                    blockFormulaString = blockFormulaString & colLetter & rows(q) & ")"
                Else
                    ' For all other items, add a comma between cell references
                    blockFormulaString = blockFormulaString & colLetter & rows(q) & ","
                End If
            Next q

            ws.Cells(i, colNum).Formula = blockFormulaString
        Next P
    End If
End Sub

' Add percentage formulas to columns
Sub percentColumnsSub(ws As Worksheet, columns As Collection, _
                     row As Long, colOffset As Long)

    Dim P As Long

    For P = 1 To columns.Count
        With ws.Cells(row + 1, columns(P))
            .Formula = "=" & ColumnToLetter(columns(P)) & row & "/A" & row
            .Font.Bold = False
            .NumberFormat = "0%"
        End With
    Next P
End Sub

' Link row from source to destination with formulas
Sub linkRow(wsSource As Worksheet, wsDestination As Worksheet, _
           columns As Collection, rowSrc As Long, _
           rowDest As Long, colOffset As Long)

    Dim P As Long
    Dim colNum As Long
    Dim colLetter As String

    For P = 1 To columns.Count
        colNum = columns(P) + colOffset
        colLetter = ColumnToLetter(colNum)

        With wsDestination.Cells(rowDest, colNum)
            .Formula = "='" & wsSource.Name & "'!" & colLetter & rowSrc
            .Font.Bold = False
        End With
    Next P
End Sub

' ----------------------------------------------------------------------------
' Dwelling Lookup Function
' ----------------------------------------------------------------------------

' Find lookup table range by header name in column A
' Returns the range of the lookup table (10 rows below the header)
' Also builds the header map from the header row found
Function FindLookupTable(wsTemplate As Worksheet, lookupName As String, _
                         ByRef headerMap As Object) As Range

    Dim foundCell As Range
    Dim searchRange As Range
    Dim headerRow As Long

    Set searchRange = wsTemplate.columns("A")
    Set foundCell = searchRange.Find(What:=lookupName, LookIn:=xlValues, LookAt:=xlWhole)

    If foundCell Is Nothing Then
        Set FindLookupTable = Nothing
        Set headerMap = Nothing
        Exit Function
    End If

    headerRow = foundCell.row
    Set headerMap = BuildHeaderMap(wsTemplate, headerRow)
    Set FindLookupTable = wsTemplate.rows(headerRow + 1).Resize(10)

End Function

' Lookup a specific value from dwelling lookup tables
' Returns the value from the lookup table based on dwelling type, bed/person config, and key name
Function GetLookupValue(wsTemplate As Worksheet, dwellingType As String, _
                       bedCount As Long, personCount As Long, _
                       lookupKey As String) As Variant

    Dim lookupName As String
    Dim lookupRange As Range
    Dim lookupHeaderMap As Object
    Dim foundRow As Range
    Dim searchKey As String

    ' Determine lookup table name based on dwelling type
    Select Case True
        Case InStr(1, UCase(dwellingType), "HOUSE") > 0
            lookupName = "House Lookup"
        Case InStr(1, UCase(dwellingType), "DUPLEX") > 0 Or InStr(1, UCase(dwellingType), "DUP") > 0
            lookupName = "Duplex Lookup"
        Case InStr(1, UCase(dwellingType), "APARTMENT") > 0 Or InStr(1, UCase(dwellingType), "APT") > 0
            lookupName = "Apartment Lookup"
        Case Else
            GetLookupValue = ""
            Exit Function
    End Select

    ' Find the lookup table dynamically
    Set lookupRange = FindLookupTable(wsTemplate, lookupName, lookupHeaderMap)

    If lookupRange Is Nothing Or lookupHeaderMap Is Nothing Then
        GetLookupValue = ""
        Exit Function
    End If

    ' Build search key (e.g., "1b 2p")
    searchKey = bedCount & "b " & personCount & "p"

    ' Find the row matching the bed/person configuration
    On Error Resume Next
    Set foundRow = lookupRange.Find(What:=searchKey, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0

    If foundRow Is Nothing Then
        GetLookupValue = ""
        Exit Function
    End If

    ' Return the value from the specified column
    If lookupHeaderMap.Exists(UCase(lookupKey)) Then
        GetLookupValue = wsTemplate.Cells(foundRow.row, lookupHeaderMap(UCase(lookupKey))).Value
    Else
        GetLookupValue = ""
    End If

End Function

' Apply dwelling type lookup from template (colors and minimum standards)
' Dynamically finds the lookup table by searching for header in column A
Sub ApplyDwellingLookup(wsData As Worksheet, _
    wsTemplate As Worksheet, _
    rowNum As Long, _
    dwellingType As String, _
    rngRow As Range, _
    headerMap As Object)

    Dim bedCount As Long
    Dim personCount As Long
    Dim lookupKey As String
    Dim foundRow As Range
    Dim lookupRange As Range
    Dim tempHeaderMap As Object
    Dim lookupName As String
    
    ' Determine lookup table name based on dwelling type
    Select Case True
        Case InStr(1, UCase(dwellingType), "HOUSE") > 0
            lookupName = "House Lookup"
        Case InStr(1, UCase(dwellingType), "DUPLEX") > 0 Or InStr(1, UCase(dwellingType), "DUP") > 0
            lookupName = "Duplex Lookup"
        Case InStr(1, UCase(dwellingType), "APARTMENT") > 0 Or InStr(1, UCase(dwellingType), "APT") > 0
            lookupName = "Apartment Lookup"
        Case Else
            rngRow.Interior.Color = RGB(255, 0, 0)
            Exit Sub
    End Select
    
    ' Find the lookup table dynamically
    Set lookupRange = FindLookupTable(wsTemplate, lookupName, tempHeaderMap)
    
    If lookupRange Is Nothing Or tempHeaderMap Is Nothing Then
        rngRow.Interior.Color = RGB(255, 0, 0)
        Exit Sub
    End If
    
    bedCount = wsData.Cells(rowNum, GetColByHeader(headerMap, "BEDS")).Value
    personCount = wsData.Cells(rowNum, GetColByHeader(headerMap, "PERS")).Value
    
    lookupKey = bedCount & "b " & personCount & "p"
    
    On Error Resume Next
    Set foundRow = lookupRange.Find( _
                        What:=lookupKey, _
                        LookAt:=xlWhole, _
                        MatchCase:=False)
    On Error GoTo 0
    
    If foundRow Is Nothing Then
        rngRow.Interior.Color = RGB(255, 0, 0)
        Exit Sub
    End If
    
    ' Apply template colour
    If tempHeaderMap.Exists("COLOUR") Then
        rngRow.Interior.Color = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "COLOUR")).Interior.Color
    End If
    
    ' Set minimums
    If headerMap.Exists("MINAREA") And tempHeaderMap.Exists("MINAREA") Then
        wsData.Cells(rowNum, headerMap("MINAREA")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINAREA")).Value
    End If
    If headerMap.Exists("MINPAS") And tempHeaderMap.Exists("MINPAS") Then
        wsData.Cells(rowNum, headerMap("MINPAS")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINPAS")).Value
    End If
    If headerMap.Exists("MINCAS") And tempHeaderMap.Exists("MINCAS") Then
        wsData.Cells(rowNum, headerMap("MINCAS")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINCAS")).Value
    End If
    If headerMap.Exists("MINAGBED") And tempHeaderMap.Exists("MINAGBED") Then
        wsData.Cells(rowNum, headerMap("MINAGBED")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINAGBED")).Value
    End If
    If headerMap.Exists("MINLVNG") And tempHeaderMap.Exists("MINLVNG") Then
        wsData.Cells(rowNum, headerMap("MINLVNG")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINLVNG")).Value
    End If
    If headerMap.Exists("MINSTOR") And tempHeaderMap.Exists("MINSTOR") Then
        wsData.Cells(rowNum, headerMap("MINSTOR")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINSTOR")).Value
    End If
    If headerMap.Exists("MINBED1") And tempHeaderMap.Exists("MINBED1") Then
        wsData.Cells(rowNum, headerMap("MINBED1")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINBED1")).Value
    End If
    If headerMap.Exists("MINBED2") And tempHeaderMap.Exists("MINBED2") Then
        wsData.Cells(rowNum, headerMap("MINBED2")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINBED2")).Value
    End If
    If headerMap.Exists("MINBED3") And tempHeaderMap.Exists("MINBED3") Then
        wsData.Cells(rowNum, headerMap("MINBED3")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINBED3")).Value
    End If
    If headerMap.Exists("MINBED4") And tempHeaderMap.Exists("MINBED4") Then
        wsData.Cells(rowNum, headerMap("MINBED4")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINBED4")).Value
    End If
    If headerMap.Exists("MINBED5") And tempHeaderMap.Exists("MINBED5") Then
        wsData.Cells(rowNum, headerMap("MINBED5")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINBED5")).Value
    End If
    If headerMap.Exists("MINMAIN") And tempHeaderMap.Exists("MINMAIN") Then
        wsData.Cells(rowNum, headerMap("MINMAIN")).Value = wsTemplate.Cells(foundRow.row, GetColByHeader(tempHeaderMap, "MINMAIN")).Value
    End If

End Sub





