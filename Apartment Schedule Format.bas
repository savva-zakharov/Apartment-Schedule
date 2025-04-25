Attribute VB_Name = "Module1"
Function ColLetterToNumber(colLetter As String) As Long
    ColLetterToNumber = Range(colLetter & "1").Column
End Function

Sub ProduceHQA()

    Dim wsOriginal As Worksheet
    Dim wsLong As Worksheet
    Dim wsShort As Worksheet
    Dim wsTypes As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim iShort As Long
    Dim currentLevel As Variant
    Dim previousLevel As Variant
    
    Dim typeDict As Object
    Set typeDict = CreateObject("Scripting.Dictionary")
        
    Dim currentDate As String
    currentDate = Format(Date, "yy-mm-dd") ' You can change format here
    
    
    
    
    
    ' Set the original worksheet
    Set wsOriginal = ThisWorkbook.Sheets("sourceData") ' Change to your original sheet name if needed
    
    ' Set the tempalte worksheet
    Set wsTemplate = ThisWorkbook.Sheets("template") ' Change to your original sheet name if needed
    
    ' Create a new worksheet for the Long schedule output
    Set wsShort = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    wsShort.Name = wsShort.Name & " Short " & currentDate
    
    ' Create a new worksheet for the short schedule output
    Set wsLong = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ' wsLong.Name = "ModifiedData" ' You can change the name of the new sheet here
    
    wsLong.Name = wsLong.Name & " Long " & currentDate
    
    ' Create a new worksheet for the unit types
    Set wsTypes = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ' wsLong.Name = "ModifiedData" ' You can change the name of the new sheet here
    
    wsTypes.Name = wsTypes.Name & " Types " & currentDate
    
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
    

    ' Apply color formatting based on column E values
    For i = 2 To lastRow
        Set rng = wsLong.Range(wsLong.Cells(i, "A"), wsLong.Cells(i, "N"))
        Select Case Mid(wsLong.Cells(i, "E").Value, 1, 1)
            Case "1"
                rng.Interior.Color = RGB(217, 233, 248) ' Light Blue
            Case "2"
                rng.Interior.Color = RGB(255, 242, 204) ' Light Yellow
            Case "3"
                rng.Interior.Color = RGB(251, 226, 213) ' Pink
            Case "D"
                rng.Interior.Color = RGB(211, 177, 194) ' Light Purple
        End Select
    Next i
          
    wsLong.Cells(1, "F").Value = "MIN.AREA"
    wsLong.Cells(1, "K").Value = "MIN.PR.AM"
    wsLong.Cells(1, "M").Value = "MIN.COM"
    wsLong.Cells(1, "N").Value = "10%+"
    wsLong.Cells(1, "Q").Value = "1 BED"
    wsLong.Cells(1, "R").Value = "2 BED"
    wsLong.Cells(1, "S").Value = "3 BED"
        
    ' Apply Apartment Minimum Values and add the type to the matrix
    
    For i = 2 To lastRow
        Select Case wsLong.Cells(i, "H")
            Case "1"
                wsLong.Cells(i, "F").Value = "45"    'set min apartment area
                wsLong.Cells(i, "K").Value = "5"     'set min private open space area
                'wsLong.Cells(i, "L").Value = "5"     'set actual private open space area
                wsLong.Cells(i, "M").Value = "5"     'set min communal open space area
                
                wsLong.Cells(i, "Q").Value = "1"     'set unit type tally
                
            Case "2"
                wsLong.Cells(i, "F").Value = "73"    'set min apartment area
                wsLong.Cells(i, "K").Value = "7"     'set min private open space area
                'wsLong.Cells(i, "L").Value = "7"     'set actual private open space area
                wsLong.Cells(i, "M").Value = "7"     'set min communal open space area
                
                wsLong.Cells(i, "R").Value = "1"     'set unit type tally
                
            Case "3"
                wsLong.Cells(i, "F").Value = "90"    'set min apartment area
                wsLong.Cells(i, "K").Value = "9"     'set min private open space area
                'wsLong.Cells(i, "L").Value = "9"     'set actual private open space area
                wsLong.Cells(i, "M").Value = "9"     'set min communal open space area
                
                wsLong.Cells(i, "S").Value = "1"     'set unit type tally
        End Select
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
    lastRow = wsLong.Cells(wsLong.rows.Count, "C").End(xlUp).row
    
    'find unique unit types
    
    For i = 2 To lastRow
        unitType = wsLong.Cells(i, 5).Value
        
        If Not IsEmpty(unitType) Then
            If Not typeDict.exists(unitType) = True Then
                ' Store count = 1 and first row = i
                typeDict.Add unitType, Array(1, i)
            Else
                ' Increment count
                tempArr = typeDict(unitType)
                tempArr(0) = tempArr(0) + 1
                typeDict(unitType) = tempArr
            End If
        End If
    Next i
    

    Dim outputRow As Long
    outputRow = 2
    
    Dim typeKeys As Variant
    typeKeys = typeDict.Keys
    Dim typeItems As Variant
    typeItems = typeDict.items
    
    Dim key As Long
    For key = 0 To UBound(typeKeys) - LBound(typeKeys)
        wsLong.rows(typeItems(key)(1)).Copy
        wsTypes.rows(outputRow).PasteSpecial Paste:=xlPasteAll
        wsTypes.Cells(outputRow, 1).Value = typeItems(key)(0)
        
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
    
    Dim sumTypeColumns As Collection
    Set sumTypeColumns = New Collection
    sumTypeColumns.Add "Q"
    sumTypeColumns.Add "R"
    sumTypeColumns.Add "S"
    
    Dim sumTypeResultColumns As Collection
    Set sumTypeResultColumns = New Collection
    sumTypeResultColumns.Add "T"
    sumTypeResultColumns.Add "U"
    sumTypeResultColumns.Add "V"
    
    Dim shortSumTypeColumns As Collection
    Set shortSumTypeColumns = New Collection
    shortSumTypeColumns.Add "C"
    shortSumTypeColumns.Add "D"
    shortSumTypeColumns.Add "E"
    
    Dim percentCalcColumns As Collection
    Set percentCalcColumns = New Collection
    percentCalcColumns.Add "J"
    percentCalcColumns.Add "N"
    percentCalcColumns.Add "T"
    percentCalcColumns.Add "U"
    percentCalcColumns.Add "V"
    
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
        
        If currentLevel <> previousLevel Then
            
            'record change rows
            
          
            ' Insert 3 empty rows
            wsLong.rows(i).Resize(3).Insert Shift:=xlDown
            
            ' Clear fill color for the inserted rows
            wsLong.rows(i).Resize(3).Interior.ColorIndex = -4142 ' No Fill
            
            ' Add SUM formulas for columns I-K up until the previous empty row
             
            Call sumColumnsSub(wsLong, sumColumns, levelStartRow, i, 0, True) 'summing up the main stats
            
            Call sumColumnsSub(wsLong, sumTypeColumns, levelStartRow, i, 3, False) 'summing up the apartment types and offsetting the result

            'calculate %s~
            
            Call percentColumnsSub(wsLong, percentCalcColumns, i, 0)
            
            ' Apply borders for the block range (A to N, blockStartRow to i-1)
            Dim levelRange As Range
            
            Set levelRange = wsLong.Range(wsLong.Cells(i - 1, "A"), wsLong.Cells(levelStartRow, "N"))
            
            Call drawBorderThickOutline(levelRange)


            ' Add title to floors
            With wsLong.Cells(levelStartRow - 1, "B")
                .Value = "Block " & previousBlock & " Level " & previousLevel
                .Font.Bold = True
                .Font.Color = RGB(0, 176, 240)
                .Font.Name = "Calibri"
            End With
            
            
            'link data to wsShort
            Call linkRow(wsLong, wsShort, sumColumns, i, iShort, 0) ' link the level summary row to the wsShort Schedule
            Call linkRow(wsLong, wsShort, sumTypeResultColumns, i, iShort, -17)
            With wsShort.Cells(iShort, "B")
                .Value = previousLevel
                .Font.Bold = False
            End With
            
            iShort = iShort + 1
            
                       
                        
            changeLevel.Add i
            
            'if the blocks change
            
            If currentBlock <> previousBlock Then
            
                 i = i + 3
                 
                ' Insert 3 empty rows
                wsLong.rows(i).Resize(3).Insert Shift:=xlDown
                
                changeBlock.Add i
                
                ' Clear fill color for the inserted rows
                wsLong.rows(i).Resize(3).Interior.ColorIndex = -4142 ' No Fill
                
                Call drawBorderLine(wsLong, i)     ' add summary line
                
                With wsLong.Cells(i - 1, "B")
                    .Value = "Block " & previousBlock & " Summary"
                    .Font.Bold = True
                    .Font.Color = RGB(0, 176, 240)
                    .Font.Name = "Calibri"
                End With
                             
                Call sumColumnsSub(wsLong, sumTypeColumns, blockStartRow, i, 3, False) 'summing up the apartment types and offsetting the result

                'calculate %s~
            
                Call percentColumnsSub(wsLong, percentCalcColumns, i, 0)
                
                blockStartRow = i + 1
                
                Call sumColumnsRowsSub(wsLong, sumColumns, changeLevel, i)
                

                'do updates to wsShort
                'add block titles
                With wsShort.Cells(shortBlockStartRow - 1, "B")
                    .Value = "Block " & previousBlock & " Summary"
                    .Font.Bold = True
                    .Font.Color = RGB(0, 176, 240)
                    .Font.Name = "Calibri"
                End With
                '
                Call sumColumnsSub(wsShort, sumColumns, shortBlockStartRow, iShort, 0, False)
                Call sumColumnsSub(wsShort, shortSumTypeColumns, shortBlockStartRow, iShort, 0, False)
                
                Dim shortBlockRange As Range
                
                Set shortBlockRange = wsShort.Range(wsShort.Cells(iShort - 1, "A"), wsShort.Cells(shortBlockStartRow, "N"))
                
                Call drawBorderThickOutline(shortBlockRange)
                shortChangeBlock.Add iShort
                Call percentColumnsSub(wsShort, percentCalcColumns, iShort, 0)
                Call percentColumnsSub(wsShort, shortSumTypeColumns, iShort, 0)
                                
                iShort = iShort + 3
                
                shortBlockStartRow = iShort
                
                
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
    End With
    
    Call sumColumnsSub(wsLong, sumTypeColumns, 4, i, 3, False)
    
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
    End With
    
    Call sumColumnsRowsSub(wsShort, sumColumns, shortChangeBlock, iShort) 'sum up main stats
    Call sumColumnsRowsSub(wsShort, shortSumTypeColumns, shortChangeBlock, iShort) 'sum up unit types
    Call percentColumnsSub(wsShort, percentCalcColumns, iShort, 0) 'add %
    Call percentColumnsSub(wsShort, shortSumTypeColumns, iShort, 0) 'add %
    
    
  
    
    'center align columns A B C E
    With wsLong
        .Range("A1:A" & .Cells(.rows.Count, "A").End(xlUp).row).HorizontalAlignment = xlCenter
        .Range("B1:B" & .Cells(.rows.Count, "B").End(xlUp).row).HorizontalAlignment = xlCenter
        .Range("C1:C" & .Cells(.rows.Count, "C").End(xlUp).row).HorizontalAlignment = xlCenter
        .Range("E1:E" & .Cells(.rows.Count, "E").End(xlUp).row).HorizontalAlignment = xlCenter
    End With
    With wsShort
        .Range("B1:B" & .Cells(.rows.Count, "B").End(xlUp).row).HorizontalAlignment = xlCenter
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
    
    With wsShort.columns("F").Font
        .Color = RGB(128, 128, 128) ' Grey text
        .Bold = True
    End With
    With wsShort.columns("K").Font
        .Color = RGB(128, 128, 128) ' Grey text
        .Bold = True
    End With
    With wsShort.columns("M").Font
        .Color = RGB(128, 128, 128) ' Grey text
        .Bold = True
    End With
    
    
    
    
    
    
    'copy the headers fropm wsTemplate
    wsTemplate.rows("1:8").Copy
    wsLong.Range("A1").Insert Shift:=xlDown
    wsLong.rows("9:9").Delete
    
    wsTemplate.rows("10:17").Copy
    wsShort.Range("A1").Insert Shift:=xlDown
    wsShort.rows("9:11").Delete
    
    'set print areas
    lastRow = wsLong.Cells(wsLong.rows.Count, "N").End(xlUp).row
    wsLong.PageSetup.PrintArea = "A1:N" & lastRow
    
    lastRow = wsShort.Cells(wsShort.rows.Count, "N").End(xlUp).row
    wsShort.PageSetup.PrintArea = "A1:N" & lastRow
    
    
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
    

    wsLong.Activate
    Application.CutCopyMode = False
    wsLong.Range("A1").Select
    
    wsShort.Activate
    Application.CutCopyMode = False
    wsShort.Range("A1").Select
    
    wsTypes.Activate
    Application.CutCopyMode = False
    wsTypes.Range("A1").Select
    
    'wsLong.Columns("O:O").Cut
    'wsLong.Columns("F:F").Insert Shift:=xlToRight
    'wsLong.Columns("K:K").Copy
    'wsLong.Columns("L:L").PasteSpecial Paste:=xlPasteAll
    'wsTemplate.Range("A1:N2").Copy Destination:=wsShort.Range("A1") 'copy the header from template
    
    
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


