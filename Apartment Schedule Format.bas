Function ColLetterToNumber(colLetter As String) As Long
    ColLetterToNumber = Range(colLetter & "1").Column
End Function

Sub ProduceHQA()

    Dim wsOriginal As Worksheet
    Dim wsLong As Worksheet
    Dim wsShort As Worksheet
    Dim wsTypes As Worksheet
    Dim wsTemplate As Worksheet
    Dim wsBlocks As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim iShort As Long
    Dim iBlocks As Long
    Dim currentLevel As Variant
    Dim previousLevel As Variant
    
    Dim typeDict As Object
'    Dim blockDict As Object
    Set typeDict = CreateObject("Scripting.Dictionary")
        
    Dim currentDate As String
    currentDate = Format(Date, "yy-mm-dd") ' You can change format here
    
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, "Sheet", vbTextCompare) > 0 Then
            ws.Delete
        End If
    Next ws
    
    Application.DisplayAlerts = True
    
    ' Set the original worksheet
    Set wsOriginal = ThisWorkbook.Sheets("sourceData") ' Change to your original sheet name if needed
    
    ' Set the tempalte worksheet
    Set wsTemplate = ThisWorkbook.Sheets("template") ' Change to your original sheet name if needed
    
    ' Create a new worksheet for the Long schedule output
    Set wsShort = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    wsShort.Name = wsShort.Name & " Short " & currentDate
    
    ' Create a new worksheet for the short schedule output
    Set wsLong = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    wsLong.Name = wsLong.Name & " Long " & currentDate
    
    ' Create a new worksheet for the unit types
    Set wsTypes = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    wsTypes.Name = wsTypes.Name & " Types " & currentDate
    
    ' Create a new worksheet for the blocks summary
    Set wsBlocks = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    wsBlocks.Name = wsBlocks.Name & " Blocks " & currentDate
    
    ' Copy all data from the original worksheet to the new worksheet
    wsOriginal.Cells.Copy Destination:=wsLong.Cells(1, 1)
    
    'Change the dates in the template
    
    wsTemplate.Cells(5, 4).Value = FormatDateWithSuffix(Date)
    wsTemplate.Cells(14, 4).Value = FormatDateWithSuffix(Date)
    wsTemplate.Cells(24, 4).Value = FormatDateWithSuffix(Date)
        
    'Delete columnd A and B
    wsLong.columns("A:B").Delete
    
    ' Find the last used row in column C of the new sheet
    lastRow = wsLong.Cells(wsLong.rows.Count, "C").End(xlUp).row
    
    ' Delete rows where column C equals 0
    For i = lastRow To 2 Step -1 ' Start from the bottom to avoid skiplevelStartRowng rows
        If wsLong.Cells(i, "D").Value = 0 Then
            wsLong.rows(i).Delete
        End If
    Next i
    
     ' Find the last row with data in column A (assuming columns A to H are populated)
    lastRow = wsLong.Cells(wsLong.rows.Count, "A").End(xlUp).row
    
    ' Define the range of data you want to sort (A1 to K[lastRow])
    With wsLong.Sort
        .SortFields.Clear ' Clear any previous sort fields
        
        ' Sort by Column D (6th column), then by Column E (7th column), then by Column F (8th column)
        .SortFields.Add key:=wsLong.Range("D2:D" & lastRow), Order:=xlAscending ' Column D
        .SortFields.Add key:=wsLong.Range("E2:E" & lastRow), Order:=xlAscending ' Column E
        .SortFields.Add key:=wsLong.Range("F2:F" & lastRow), Order:=xlAscending ' Column F
        
        ' Apply the sorting to the range from A1 to H[lastRow]
        .SetRange wsLong.Range("A1:O" & lastRow)
        
        ' Apply the sort
        .Header = xlYes ' Assuming your data includes headers
        .Apply
    End With
    
    ' Recalculate the last row after deletions
    lastRow = wsLong.Cells(wsLong.rows.Count, "C").End(xlUp).row
    
    ' Move columns H, F, and G to A, B, and C respectively
    
        
    wsLong.columns("F:F").Cut
    wsLong.columns("A:A").Insert Shift:=xlToRight
    wsLong.columns("E:E").Cut
    wsLong.columns("B:B").Insert Shift:=xlToRight
    wsLong.columns("F:F").Cut
    wsLong.columns("C:C").Insert Shift:=xlToRight
    wsLong.columns("O:O").Cut
    wsLong.columns("F:F").Insert Shift:=xlToRight
    wsLong.columns("K:K").Copy
    wsLong.columns("L:L").PasteSpecial Paste:=xlPasteAll
    
    wsLong.Cells(1, "F").Value = "MIN.AREA"
    wsLong.Cells(1, "K").Value = "MIN.PR.AM"
    wsLong.Cells(1, "M").Value = "MIN.COM"
    wsLong.Cells(1, "N").Value = "10%+"
    wsLong.Cells(1, "Q").Value = "1 BED"
    wsLong.Cells(1, "R").Value = "2 BED"
    wsLong.Cells(1, "S").Value = "3 BED"

    Dim tally As Object
    Set tally = CreateObject("Scripting.Dictionary")
    tally.Add 1, "Q"
    tally.Add 2, "R"
    tally.Add 3, "S"
    tally.Add 4, "T"
    Dim rng As Range
    
    '#################################
    '## APPLY COLOURS AND STANDARDS ##
    '#################################
    
    For i = 2 To lastRow
        Set rng = wsLong.Range(wsLong.Cells(i, "A"), wsLong.Cells(i, "N"))
    
        Select Case True
    
            ' HOUSES
            Case InStr(1, UCase(wsLong.Cells(i, "D").Value), "HOUSE") > 0
                Call ApplyDwellingLookup(wsLong, wsTemplate, i, "T27:T32", rng, tally)
    
            ' DUPLEX
            Case InStr(1, UCase(wsLong.Cells(i, "D").Value), "DUPLEX") > 0 _
              Or InStr(1, UCase(wsLong.Cells(i, "D").Value), "DUP") > 0
                Call ApplyDwellingLookup(wsLong, wsTemplate, i, "T17:T22", rng, tally)
    
            ' APARTMENTS
            Case InStr(1, UCase(wsLong.Cells(i, "D").Value), "APARTMENT") > 0 _
              Or InStr(1, UCase(wsLong.Cells(i, "D").Value), "APT") > 0
                Call ApplyDwellingLookup(wsLong, wsTemplate, i, "T8:T13", rng, tally)
    
        End Select
    
        ' COMPLIANCE CHECK – cell-level only
    
        If Val(wsLong.Cells(i, "G").Value) < Val(wsLong.Cells(i, "F").Value) _
           And Val(wsLong.Cells(i, "F").Value) > 0 Then
            wsLong.Cells(i, "G").Interior.Color = RGB(255, 0, 0)
        End If
    
        If Val(wsLong.Cells(i, "L").Value) < Val(wsLong.Cells(i, "K").Value) _
           And Val(wsLong.Cells(i, "K").Value) > 0 Then
            wsLong.Cells(i, "L").Interior.Color = RGB(255, 0, 0)
        End If
    
    Next i

    'Add a +10% indicator to every row
    For i = 2 To lastRow
         areaExt = wsLong.Cells(i, "F").Value * 1.1
         areaCur = wsLong.Cells(i, "G").Value
         
         If areaCur > areaExt Then
            wsLong.Cells(i, "N").Value = "1"
         Else
            wsLong.Cells(i, "N").Value = "0"
         End If
    
    Next i
    
   
    
    'find the last row
    lastRow = wsLong.Cells(wsLong.rows.Count, "E").End(xlUp).row
    
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

    
    For i = 2 To lastRow
        unitType = wsLong.Cells(i, 5).Value
    
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
    iBlocks = 2
    
    Dim typeKeys As Variant
    typeKeys = typeDict.Keys
    Dim typeItems As Variant
    typeItems = typeDict.Items
    
    For key = LBound(typeKeys) To UBound(typeKeys)
    
        With wsTypes.rows(outputRow)
            .Value = wsLong.rows(typeItems(key)(1)).Value
        End With
        
        wsTypes.Range(wsTypes.Cells(outputRow, 3), wsTypes.Cells(outputRow, 14)).Interior.Color = _
        wsLong.Cells(typeItems(key)(1), 1).Interior.Color

    
        ' overwrite column A with count (same as your original code)
        wsTypes.Cells(outputRow, 1).Value = typeItems(key)(0)
        ' overwrite column E with combined unit type
        wsTypes.Cells(outputRow, "E").Value = typeKeys(key)
    
        outputRow = outputRow + 1
    Next key

    
    wsTypes.Range("B:C").ClearContents
    wsTypes.Range("Q:S").ClearContents
    wsTypes.Range("A:C").Interior.ColorIndex = xlColorIndexNone
    
    lastRow = wsTypes.Cells(wsTypes.rows.Count, "E").End(xlUp).row
    With wsTypes.Sort
        .SortFields.Clear
        .SortFields.Add key:=wsTypes.Range("E2:E" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        .SetRange wsTypes.Range("A2:N" & lastRow)
        .Header = xlNo
        .Apply
    End With
    
    Call drawBorderThickOutline(wsTypes.Range("A2:N" & lastRow))
    
    wsTypes.Cells(lastRow + 1, 1).Formula = "=SUM(A2:A" & lastRow & ")"
    
    With wsTypes.columns("F").Font
        .Color = RGB(128, 128, 128) ' Grey text
        .Bold = True
    End With
    With wsTypes.columns("K").Font
        .Color = RGB(128, 128, 128) ' Grey text
        .Bold = True
    End With
    With wsTypes.columns("M").Font
        .Color = RGB(128, 128, 128) ' Grey text
        .Bold = True
    End With
    
    wsTemplate.rows("20:27").Copy
    wsTypes.Range("A1").Insert Shift:=xlDown
    
    
    lastRow = wsTypes.Cells(wsTypes.rows.Count, "E").End(xlUp).row
    
    wsTypes.PageSetup.PrintArea = "A1:N" & lastRow + 1
    
    
    'enable print preview
    wsTypes.Activate
    ActiveWindow.View = xlPageBreakPreview
    With wsTypes.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False ' Can be 1 or left as False to auto-scale height
    End With
    With wsTypes.PageSetup
        .PrintTitleRows = "$7:$9"
    End With
    
    
    
    
    
    
    
 
    
    
    
    
    
    'find the last row
    lastRow = wsLong.Cells(wsLong.rows.Count, "C").End(xlUp).row
    
        
    ' Insert 2 empty rows
    wsLong.rows(2).Resize(3).Insert Shift:=xlDown
    
    'set up collections for tracking level and floor changes
    Dim changeLevel As Collection
    Set changeLevel = New Collection    'collection for tracking level changes
    Dim changeBlock As Collection
    Set changeBlock = New Collection    'collection for tracking block changes
    Dim shortChangeBlock As Collection  'collection for tracking block changes in the wsShort Schedule
    Set shortChangeBlock = New Collection
    Dim blocksChangeBlock As Collection 'collection for tracking block changes in the wsBlocks Schedule
    Set blocksChangeBlock = New Collection
    
    'set up columnd to run summs on
    'wsLong columns
    
    Dim sumTypeColumns As Collection
    Set sumTypeColumns = New Collection
    sumTypeColumns.Add "Q"
    sumTypeColumns.Add "R"
    sumTypeColumns.Add "S"
    sumTypeColumns.Add "T"
    
    Dim sumTypeResultColumns As Collection
    Set sumTypeResultColumns = New Collection
    sumTypeResultColumns.Add "U"
    sumTypeResultColumns.Add "V"
    sumTypeResultColumns.Add "W"
    sumTypeResultColumns.Add "X"
    
    Dim shortSumTypeColumns As Collection
    Set shortSumTypeColumns = New Collection
    shortSumTypeColumns.Add "C"
    shortSumTypeColumns.Add "D"
    shortSumTypeColumns.Add "E"
    shortSumTypeColumns.Add "F"
    
    Dim percentCalcColumns As Collection
    Set percentCalcColumns = New Collection
    percentCalcColumns.Add "J"
    percentCalcColumns.Add "N"
    percentCalcColumns.Add "U"
    percentCalcColumns.Add "V"
    percentCalcColumns.Add "W"
    percentCalcColumns.Add "X"
    
    Dim blocksSumColumns As Collection
    Set blocksSumColumns = New Collection
    blocksSumColumns.Add "A"
    blocksSumColumns.Add "F"
    blocksSumColumns.Add "H"
    blocksSumColumns.Add "I"
    blocksSumColumns.Add "J"
    blocksSumColumns.Add "K"
    blocksSumColumns.Add "L"
    blocksSumColumns.Add "M"
    blocksSumColumns.Add "N"
    
    
    Dim sumColumns As Collection
    Set sumColumns = New Collection    'collection for columns to add
    
    'set up which columnds need to be summed, first columnd will use COUNTA instead if last comment is set to TRUE
    sumColumns.Add "A"
    sumColumns.Add "F"
    sumColumns.Add "G"
    sumColumns.Add "H"
    sumColumns.Add "I"
    sumColumns.Add "J"
    sumColumns.Add "K"
    sumColumns.Add "L"
    sumColumns.Add "M"
    sumColumns.Add "N"
    
    'initiating the loop and strating parameters
    i = 5 'start on row 5
    iShort = 5 'start filling out wsShort on row 5
    
    Dim levelStartRow As Long
    Dim blockStartRow As Long
    
    
    Dim shortBlockStartRow As Long
    
    previousLevel = wsLong.Cells(5, "C").Value 'take the initial level name
    previousBlock = wsLong.Cells(5, "B").Value 'take the initial block name
    levelStartRow = 5 'take the initial level start postion
    blockStartRow = 5 'take the initial block start postion
    shortBlockStartRow = 5
    
    Do While True
        currentLevel = wsLong.Cells(i, "C").Value
        currentBlock = wsLong.Cells(i, "B").Value
        
        If currentLevel <> previousLevel Or currentBlock <> previousBlock Then
            
            'record change rows
            
          
            ' Insert 3 empty rows
            wsLong.rows(i).Resize(3).Insert Shift:=xlDown
            
            ' Clear fill color for the inserted rows
            wsLong.rows(i).Resize(3).Interior.ColorIndex = -4142 ' No Fill
            
            ' Add SUM formulas for columns I-K up until the previous empty row
             
            Call sumColumnsSub(wsLong, sumColumns, levelStartRow, i, 0, True) 'summing up the main stats
            
            Call sumColumnsSub(wsLong, sumTypeColumns, levelStartRow, i, 4, False) 'summing up the apartment types and offsetting the result

            'calculate %s~
            
            Call percentColumnsSub(wsLong, percentCalcColumns, i, 0)
            
            ' Apply borders for the block range (A to N, blockStartRow to i-1)
            Dim levelRange As Range
            
            Set levelRange = wsLong.Range(wsLong.Cells(i - 1, "A"), wsLong.Cells(levelStartRow, "N"))
            
            Call drawBorderThickOutline(levelRange)


            ' Add title to floors
            
            Dim levelTitle As String
            Dim re1 As Object
            Set re1 = CreateObject("VBScript.RegExp")
            
            With re1
                .Pattern = "^[A-Za-z0-9]{1,2}$"
                .IgnoreCase = True
                .Global = False
            End With
            
            If re1.Test(previousBlock) Then
                ' string matches 1–2 alphanumeric characters
                levelTitle = "Block " & previousBlock & " Level " & previousLevel
            Else
'                levelTitle = previousBlock
            End If
            
            
            With wsLong.Cells(levelStartRow - 1, "B")
                .Value = levelTitle
                .Font.Bold = True
                .Font.Color = RGB(0, 176, 240)
                .Font.Name = "Calibri"
                .HorizontalAlignment = xlLeft
            End With
            
            
            'link data to wsShort
            Call linkRow(wsLong, wsShort, sumColumns, i, iShort, 0) ' link the level summary row to the wsShort Schedule
            Call linkRow(wsLong, wsShort, sumTypeResultColumns, i, iShort, -18)
            With wsShort.Cells(iShort, "B")
                .Value = previousLevel
                .Font.Bold = False
            End With
            
            iShort = iShort + 1
            
                       
                        
            changeLevel.Add i
            
            'if the blocks change
            
            If currentBlock <> previousBlock Then
            
                Dim blockEndRow As Long
                blockEndRow = i - 1   ' last data row of the block
                
                ' Add block summary title
                Dim blockTitle As String
                If re1.Test(previousBlock) Then
                    ' string matches 1–2 alphanumeric characters
                    blockTitle = "Block " & previousBlock & " Summary"
                Else
                    blockTitle = previousBlock & " Summary"
                End If
            

                
                ' Clear fill color for the inserted rows
                wsLong.rows(i).Resize(3).Interior.ColorIndex = -4142 ' No Fill
                
                'write a summary only if there are multiple levels in a block / area
                If previousLevel <> 0 Then
                    i = i + 3
                 
                    ' Insert 3 empty rows
                    wsLong.rows(i).Resize(3).Insert Shift:=xlDown
                    
                    changeBlock.Add i
                
                    Call drawBorderLine(wsLong, i)     ' add summary line
                    
                    With wsLong.Cells(i - 1, "B")
                        .Value = blockTitle
                        .Font.Bold = True
                        .Font.Color = RGB(0, 176, 240)
                        .Font.Name = "Calibri"
                        .HorizontalAlignment = xlLeft
                    End With
                                 
                    Call sumColumnsSub(wsLong, sumTypeColumns, blockStartRow, i, 4, False) 'summing up the apartment types and offsetting the result
    
                    'calculate %s~
                
                    Call percentColumnsSub(wsLong, percentCalcColumns, i, 0)
                                    
                    Call sumColumnsRowsSub(wsLong, sumColumns, changeLevel, i)
                Else
                    changeBlock.Add i
                End If
                
                '########################
                '## updates to wsShort ##
                '########################
                
                'add block titles
                With wsShort.Cells(shortBlockStartRow - 1, "B")
                    .Value = blockTitle
                    .Font.Bold = True
                    .Font.Color = RGB(0, 176, 240)
                    .Font.Name = "Calibri"
                    .HorizontalAlignment = xlLeft
                End With
                '
                Call sumColumnsSub(wsShort, sumColumns, shortBlockStartRow, iShort, 0, False)
                Call sumColumnsSub(wsShort, shortSumTypeColumns, shortBlockStartRow, iShort, 0, False)
                
                Dim shortBlockRange As Range
                
                Set shortBlockRange = wsShort.Range(wsShort.Cells(iShort - 1, "A"), wsShort.Cells(shortBlockStartRow, "N"))
                
                
                
                
                '#########################
                '## updates to wsBlocks ##
                '#########################
                
                With wsBlocks.Cells(iBlocks - 1, "B")
                    .Value = blockTitle
                    .Font.Bold = True
                    .Font.Color = RGB(0, 176, 240)
                    .Font.Name = "Calibri"
                    .HorizontalAlignment = xlLeft
                End With
                
                Dim blockStartRowBlocks As Long
                blockStartRowBlocks = iBlocks
                
                Dim blockDict As Object
                Set blockDict = CreateObject("Scripting.Dictionary")
                
                Dim re As Object
                Set re = CreateObject("VBScript.RegExp")
                're.Pattern = "^\d?[A-Za-z]"
                
                ' Read the cell
                regexPattern = Trim(wsTemplate.Range("X2").Value)
                
                ' Check if empty and assign default
                If Len(regexPattern) = 0 Then
                    regexPattern = ".*"    ' default regex: matches anything
                End If
                
                ' Apply to your regex object
                With re
                    .Global = False
                    .IgnoreCase = True
                    .Pattern = regexPattern
                End With

                re.Global = False
                
                Dim o As Long
                'Dim unitKey As String
                Dim floorArea As Double
                Dim aspect As Long
                Dim amenityArea As Double
                
                For o = blockStartRow To blockEndRow
                    unitType = wsLong.Cells(o, 5).Value
                    floorArea = 0
                    amenityArea = 0
                    aspect = 0
                
                    If IsNumeric(wsLong.Cells(o, "G").Value) Then
                        floorArea = wsLong.Cells(o, "G").Value
                    End If
                    If IsNumeric(wsLong.Cells(o, "J").Value) Then
                        aspect = wsLong.Cells(o, "J").Value
                    End If
                    If IsNumeric(wsLong.Cells(o, "L").Value) Then
                        amenityArea = wsLong.Cells(o, "L").Value
                    End If
                
                    If Len(unitType) > 0 And re.Test(unitType) Then
                        unitKey = UCase(Trim(re.Execute(unitType)(0)))
                
                        If Not blockDict.Exists(unitKey) Then
                            ' count, first row, total floor area
                            blockDict.Add unitKey, Array(1, o, floorArea, aspect, amenityArea)
                        Else
                            tempArr = blockDict(unitKey)
                            tempArr(0) = tempArr(0) + 1          ' count
                            tempArr(2) = tempArr(2) + floorArea ' sum of floor areas
                            tempArr(3) = tempArr(3) + aspect    ' sum of aspect
                            tempArr(4) = tempArr(4) + amenityArea    ' sum of aspect
                            blockDict(unitKey) = tempArr
                        End If
                    End If
                Next o

                
                ' Calculate total units in block for %
                Dim totalUnits As Long
                totalUnits = 0
                Dim k As Variant
                For Each k In blockDict.Keys
                    totalUnits = totalUnits + blockDict(k)(0)
                Next k
                
                Dim blockKeys As Variant
                Dim iKey As Long, jKey As Long
                Dim tempKey As String
                
                ' Get dictionary keys
                blockKeys = blockDict.Keys
                
                ' Sort keys alphabetically (A-Z) using simple bubble sort
                For iKey = LBound(blockKeys) To UBound(blockKeys) - 1
                    For jKey = iKey + 1 To UBound(blockKeys)
                        If blockKeys(iKey) > blockKeys(jKey) Then
                            tempKey = blockKeys(iKey)
                            blockKeys(iKey) = blockKeys(jKey)
                            blockKeys(jKey) = tempKey
                        End If
                    Next jKey
                Next iKey

                
                Dim srcRow As Range, dstRow As Range
                Dim c As Long
                
                ' Write block summary rows
                For key = LBound(blockKeys) To UBound(blockKeys)
                    Set srcRow = wsLong.rows(blockDict(blockKeys(key))(1))
                    Set dstRow = wsBlocks.rows(iBlocks)
                    
                    ' Copy values
                    dstRow.Value = srcRow.Value
                    
                    ' Copy formatting for columns D:N
                    For c = 4 To 14
                        With dstRow.Cells(1, c)
                            .Interior.Color = srcRow.Cells(1, c).Interior.Color
                        End With
                    Next c
                    
                    ' Column A = count of units
                    wsBlocks.Cells(iBlocks, 1).Value = blockDict(blockKeys(key))(0)
                    
                    ' Column B = % of total units (integer percent)
                    If totalUnits > 0 Then
                        wsBlocks.Cells(iBlocks, 2).Value = Format(blockDict(blockKeys(key))(0) / totalUnits, "0%")
                    Else
                        wsBlocks.Cells(iBlocks, 2).Value = "0%"
                    End If
                    
                    Select Case wsBlocks.Cells(iBlocks, "H").Value

                        Case "1"
                            wsBlocks.Cells(iBlocks, "Q").Value = blockDict(blockKeys(key))(0)
                
                        Case "2"
                            wsBlocks.Cells(iBlocks, "R").Value = blockDict(blockKeys(key))(0)
                
                        Case "3"
                            wsBlocks.Cells(iBlocks, "S").Value = blockDict(blockKeys(key))(0)
                
                        Case Else
                            wsBlocks.Cells(iBlocks, "T").Value = blockDict(blockKeys(key))(0)
                
                    End Select
                    
'                    With wsBlocks.Cells(iBlocks, "F")
'                        .Value = wsBlocks.Cells(iBlocks, 7).Value * blockDict(blockKeys(key))(0)
'                    End With
                    
                    
                    wsBlocks.Cells(iBlocks, "F").Value = blockDict(blockKeys(key))(2)
                    
                    
                    With wsBlocks.Cells(iBlocks, "H")
                        .Value = wsBlocks.Cells(iBlocks, "H").Value * blockDict(blockKeys(key))(0)
                    End With
                    With wsBlocks.Cells(iBlocks, "I")
                        .Value = wsBlocks.Cells(iBlocks, "I").Value * blockDict(blockKeys(key))(0)
                    End With
                    'aspect
                    wsBlocks.Cells(iBlocks, "J").Value = blockDict(blockKeys(key))(3)
                    
                    With wsBlocks.Cells(iBlocks, "K")
                        .Value = wsBlocks.Cells(iBlocks, "K").Value * blockDict(blockKeys(key))(0)
                    End With
                    
                    wsBlocks.Cells(iBlocks, "L").Value = blockDict(blockKeys(key))(4)
                    
                    With wsBlocks.Cells(iBlocks, "M")
                        .Value = wsBlocks.Cells(iBlocks, "M").Value * blockDict(blockKeys(key))(0)
                    End With
                    With wsBlocks.Cells(iBlocks, "N")
                        .Value = wsBlocks.Cells(iBlocks, "N").Value * blockDict(blockKeys(key))(0)
                    End With
                    
                    
                    'tally unit bedroom types
                    
                    
                    
                    ' Clear column C
                    wsBlocks.Cells(iBlocks, 3).ClearContents
                    
                    ' Column E = unit type
                    wsBlocks.Cells(iBlocks, 5).Value = blockKeys(key)
                    
                    iBlocks = iBlocks + 1
                Next key

                
                ' Draw borders for the block
                Call drawBorderLine(wsBlocks, iBlocks)
                Call sumColumnsSub(wsBlocks, blocksSumColumns, blockStartRowBlocks, iBlocks, 0, False)
                Call sumColumnsSub(wsBlocks, sumTypeColumns, blockStartRowBlocks, iBlocks, 4, False)
                Call percentColumnsSub(wsBlocks, percentCalcColumns, iBlocks, 0) 'add %
                
                Dim blockRange As Range
                Set blockRange = wsBlocks.Range( _
                    wsBlocks.Cells(iBlocks - 1, "A"), _
                    wsBlocks.Cells(blockStartRowBlocks, "N") _
                )
                
                Call drawBorderThickOutline(blockRange)
                
                




                'format wsLong
                
                Call drawBorderThickOutline(shortBlockRange)
                Call percentColumnsSub(wsShort, percentCalcColumns, iShort, 0)
                Call percentColumnsSub(wsShort, shortSumTypeColumns, iShort, 0)
                
                
                shortChangeBlock.Add iShort
                blocksChangeBlock.Add iBlocks
                iShort = iShort + 3
                iBlocks = iBlocks + 3
                
                shortBlockStartRow = iShort
                
                
                blockStartRow = i + 1
                
                
                Set changeLevel = New Collection
                previousBlock = currentBlock
            End If
            
            i = i + 3 ' Skip the inserted empty rows
            levelStartRow = i
            lastRow = lastRow + 3
        End If
                  
        previousLevel = currentLevel
                
        i = i + 1
        
        If wsLong.Cells(i - 1, 1).Value = 0 Then
            Exit Do
        End If
        
        If i > 100000 Then
            Exit Do
        End If
    Loop
    
    
    'start filling out the overall summary but offsetting the position
    i = i - 1
            
    'add a summary line
    Call drawBorderLine(wsLong, i)
        
    'add a title
    With wsLong.Cells(i - 1, "B")
        .Value = "Whole Scheme Summary"
        .Font.Bold = True
        .Font.Color = RGB(0, 176, 240)
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlLeft
    End With
    
    Call sumColumnsSub(wsLong, sumTypeColumns, 4, i, 4, False)
    Call percentColumnsSub(wsLong, percentCalcColumns, i, 0)
    
    'start summing up the main stats
    
    Call sumColumnsRowsSub(wsLong, sumColumns, changeBlock, i)
    
    
    
    'fit out wsShort Totals
    
    'add a summary line
    Call drawBorderLine(wsShort, iShort)
        
    'add a title
    With wsShort.Cells(iShort - 1, "B")
        .Value = "Whole Scheme Summary"
        .Font.Bold = True
        .Font.Color = RGB(0, 176, 240)
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlLeft
    End With
    
    Call sumColumnsRowsSub(wsShort, sumColumns, shortChangeBlock, iShort) 'sum up main stats
    Call sumColumnsRowsSub(wsShort, shortSumTypeColumns, shortChangeBlock, iShort) 'sum up unit types
    Call percentColumnsSub(wsShort, percentCalcColumns, iShort, 0) 'add %
    Call percentColumnsSub(wsShort, shortSumTypeColumns, iShort, 0) 'add %
    
    'fit out wsBlocks Totals
    
    'add a summary line
    Call drawBorderLine(wsBlocks, iBlocks)
    
    'add a title
    With wsBlocks.Cells(iBlocks - 1, "B")
        .Value = "Whole Scheme Summary"
        .Font.Bold = True
        .Font.Color = RGB(0, 176, 240)
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlLeft
    End With
    
    Call sumColumnsRowsSub(wsBlocks, blocksSumColumns, blocksChangeBlock, iBlocks) 'sum up main stats
    Call sumColumnsSub(wsBlocks, sumTypeColumns, 1, iBlocks, 0, False) 'sum up unit types
    Call percentColumnsSub(wsBlocks, percentCalcColumns, iBlocks, 0) 'add %
    Call percentColumnsSub(wsBlocks, sumTypeColumns, iBlocks, 0) 'add %
  
    
    'center align columns A B C E
    With wsLong
        .Range("A1:A" & .Cells(.rows.Count, "A").End(xlUp).row).HorizontalAlignment = xlCenter
'        .Range("B1:B" & .Cells(.rows.Count, "B").End(xlUp).row).HorizontalAlignment = xlCenter
        .Range("C1:C" & .Cells(.rows.Count, "C").End(xlUp).row).HorizontalAlignment = xlCenter
        .Range("E1:E" & .Cells(.rows.Count, "E").End(xlUp).row).HorizontalAlignment = xlCenter
    End With
    With wsShort
'        .Range("B1:B" & .Cells(.rows.Count, "B").End(xlUp).row).HorizontalAlignment = xlCenter
    End With
    
    
    'some final fomatting
    columns("E").Font.Bold = True 'bold text for apartment type
    With wsLong.columns("F").Font
        .Color = RGB(128, 128, 128) ' Grey text
        .Bold = True
    End With
    With wsLong.columns("K").Font
        .Color = RGB(128, 128, 128) ' Grey text
        .Bold = True
    End With
    With wsLong.columns("M").Font
        .Color = RGB(128, 128, 128) ' Grey text
        .Bold = True
    End With
    
'    With wsShort.columns("F").Font
'        .Color = RGB(128, 128, 128) ' Grey text
'        .Bold = True
'    End With
    With wsShort.columns("K").Font
        .Color = RGB(128, 128, 128) ' Grey text
        .Bold = True
    End With
    With wsShort.columns("M").Font
        .Color = RGB(128, 128, 128) ' Grey text
        .Bold = True
    End With
    
    With wsBlocks.columns("E").Font
        .Bold = True
    End With
'    With wsBlocks.columns("F").Font
'        .Color = RGB(128, 128, 128) ' Grey text
'        .Bold = True
'    End With
    With wsBlocks.columns("K").Font
        .Color = RGB(128, 128, 128) ' Grey text
        .Bold = True
    End With
    With wsBlocks.columns("M").Font
        .Color = RGB(128, 128, 128) ' Grey text
        .Bold = True
    End With
    
    
    
    
    
    
    'copy the headers fropm wsTemplate
    wsTemplate.rows("1:8").Copy
    wsLong.Range("A1").Insert Shift:=xlDown
    wsLong.rows("9:9").Delete
    With wsLong.Range("O1:Z8")
        .Clear
        .Borders.LineStyle = xlNone
    End With
    
    wsTemplate.rows("10:17").Copy
    wsShort.Range("A1").Insert Shift:=xlDown
    With wsShort.Range("O1:Z8")
        .Clear
        .Borders.LineStyle = xlNone
    End With

    wsTemplate.rows("29:36").Copy
    wsBlocks.Range("A1").Insert Shift:=xlDown
    With wsBlocks.Range("O1:Z8")
        .Clear
        .Borders.LineStyle = xlNone
    End With
    
    With wsTypes.Range("O1:Z8")
        .Clear
        .Borders.LineStyle = xlNone
    End With
    'timestamp
    
    Dim d As Date
    Dim suffix As String
    Dim dateFormated As String
    
    d = Date
    
    Select Case Day(d)
        Case 1, 21, 31: suffix = "st"
        Case 2, 22:     suffix = "nd"
        Case 3, 23:     suffix = "rd"
        Case Else:      suffix = "th"
    End Select
    
    dateFormated = Day(d) & suffix & " " & _
                           Format(d, "mmmm yyyy")
    wsLong.Range("E5").Value = dateFormated
    wsShort.Range("E5").Value = dateFormated
    wsTypes.Range("E5").Value = dateFormated
    wsBlocks.Range("E5").Value = dateFormated

    
    'set print areas
    lastRow = wsLong.Cells(wsLong.rows.Count, "N").End(xlUp).row
    wsLong.PageSetup.PrintArea = "A1:N" & lastRow
    
    lastRow = wsShort.Cells(wsShort.rows.Count, "N").End(xlUp).row
    wsShort.PageSetup.PrintArea = "A1:N" & lastRow
    
    lastRow = wsBlocks.Cells(wsBlocks.rows.Count, "N").End(xlUp).row
    wsBlocks.PageSetup.PrintArea = "A1:N" & lastRow
        
    'enable print preview
    wsLong.Activate
    ActiveWindow.View = xlPageBreakPreview
    With wsLong.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False ' Can be 1 or left as False to auto-scale height
    End With
    With wsLong.PageSetup
        .PrintTitleRows = "$7:$9"
    End With
    
    wsShort.Activate
    ActiveWindow.View = xlPageBreakPreview
    With wsShort.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False ' Can be 1 or left as False to auto-scale height
    End With
    
    With wsShort.PageSetup
        .PrintTitleRows = "$7:$9"
    End With
    
    wsBlocks.Activate
    ActiveWindow.View = xlPageBreakPreview
    With wsBlocks.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False ' Can be 1 or left as False to auto-scale height
    End With
    
    With wsBlocks.PageSetup
        .PrintTitleRows = "$7:$9"
    End With
    

    wsLong.Activate
    Application.CutCopyMode = False
    wsLong.Range("A1").Select
    
    wsShort.Activate
    Application.CutCopyMode = False
    wsShort.Range("A1").Select
    
    wsTypes.Activate
    Application.CutCopyMode = False
    wsTypes.Range("A1").Select
    
   
    
End Sub


Sub sumColumnsSub(ws As Worksheet, columns As Collection, startRow As Long, endRow As Long, colOffset As Long, Optional countFirst As Boolean = False)

        For p = 1 To columns.Count
        
        If p = 1 And countFirst = True Then
        With ws.Cells(endRow, ColLetterToNumber(columns(p)) + colOffset)
                .Formula = "=COUNTA(" & columns(p) & endRow - 1 & ":" & columns(p) & startRow & ")"
                .Font.Bold = True
        End With
        Else
        With ws.Cells(endRow, ColLetterToNumber(columns(p)) + colOffset)
                .Formula = "=SUM(" & columns(p) & endRow - 1 & ":" & columns(p) & startRow & ")"
                .Font.Bold = True
        End With
        End If
            
        Next p
End Sub

Sub sumColumnsRowsSub(ws As Worksheet, columns As Collection, rows As Collection, i As Long)
        For p = 1 To columns.Count
       
       'Declare the formula and start writing it
       Dim blockFormulaString As String
       blockFormulaString = "=SUM("
       
       For q = 1 To rows.Count
           If q = rows.Count Then
               ' For the last item, don't add a comma after it
               blockFormulaString = blockFormulaString & columns(p)
               blockFormulaString = blockFormulaString & rows(q) & ")"
           Else
               ' For all other items, add a comma between cell references
               blockFormulaString = blockFormulaString & columns(p)
               blockFormulaString = blockFormulaString & rows(q) & ","
           End If
       Next q
       
       
       ws.Range(columns(p) & i).Formula = blockFormulaString
    
       Next p
End Sub



Sub percentColumnsSub(ws As Worksheet, columns As Collection, row As Long, colOffset As Long)

        For p = 1 To columns.Count
        
        With ws.Cells(row + 1, columns(p))
            .Formula = "=" & columns(p) & row & "/A" & row
            .Font.Bold = False
            .NumberFormat = "0%"
        End With
            
        Next p
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

Sub drawBorderLine(ws As Worksheet, i As Long)
     Set rng = ws.Range(ws.Cells(i, "A"), ws.Cells(i, "N"))
    
     With rng.Borders(xlEdgeTop)
         .LineStyle = xlContinuous
         .ColorIndex = 0
         .TintAndShade = 0
         .Weight = xlThick
     End With
     
    rng.Font.Bold = True
End Sub

Sub linkRow(wsSource As Worksheet, wsDestination As Worksheet, columns As Collection, rowSrc As Long, rowDest As Long, colOffset As Long)
    For p = 1 To columns.Count
        With wsDestination.Cells(rowDest, ColLetterToNumber(columns(p)) + colOffset)
            .Formula = "='" & wsSource.Name & "'!" & columns(p) & rowSrc
            .Font.Bold = False
        End With
    Next p
End Sub

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

    ' Build the final string
    FormatDateWithSuffix = dayNum & suffix & " " & Format(dt, "mmmm yyyy")
End Function

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

Public Sub ApplyDwellingLookup( _
    wsData As Worksheet, _
    wsTemplate As Worksheet, _
    rowNum As Long, _
    lookupRange As String, _
    rngRow As Range, _
    tallyMap As Object _
)

    Dim bedCount As Long
    Dim personCount As Long
    Dim lookupKey As String
    Dim foundRow As Range
    Dim tallyCol As String

    bedCount = wsData.Cells(rowNum, "H").Value
    personCount = wsData.Cells(rowNum, "I").Value

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
    rngRow.Interior.Color = wsTemplate.Cells(foundRow.row, "U").Interior.Color

    ' Set minimums
    wsData.Cells(rowNum, "F").Value = wsTemplate.Cells(foundRow.row, "V").Value ' Min Area
    wsData.Cells(rowNum, "K").Value = wsTemplate.Cells(foundRow.row, "W").Value ' Min PAS
    wsData.Cells(rowNum, "M").Value = wsTemplate.Cells(foundRow.row, "X").Value ' Min CAS

    ' Tally
    If tallyMap.Exists(bedCount) Then
        tallyCol = tallyMap(bedCount)
        wsData.Cells(rowNum, tallyCol).Value = 1
    End If

End Sub




