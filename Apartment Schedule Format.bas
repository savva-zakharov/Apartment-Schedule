Attribute VB_Name = "Module1"

Sub ModifySpreadsheetToNewSheet()

    Dim wsOriginal As Worksheet
    Dim wsNew As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentLevel As Variant
    Dim previousLevel As Variant
    
    ' Set the original worksheet
    Set wsOriginal = ThisWorkbook.Sheets("Sheet1") ' Change to your original sheet name if needed
    
    ' Create a new worksheet for the output
    Set wsNew = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ' wsNew.Name = "ModifiedData" ' You can change the name of the new sheet here
    
    ' Copy all data from the original worksheet to the new worksheet
    wsOriginal.Cells.Copy Destination:=wsNew.Cells(1, 1)
    
    ' Find the last used row in column C of the new sheet
    lastRow = wsNew.Cells(wsNew.Rows.Count, "C").End(xlUp).Row
    
    ' Delete rows where column C equals 0
    For i = lastRow To 2 Step -1 ' Start from the bottom to avoid skiplevelStartRowng rows
        If wsNew.Cells(i, "F").Value = 0 Then
            wsNew.Rows(i).Delete
        End If
    Next i
    
     ' Find the last row with data in column A (assuming columns A to H are populated)
    lastRow = wsNew.Cells(wsNew.Rows.Count, "A").End(xlUp).Row
    
    ' Define the range of data you want to sort (A1 to K[lastRow])
    With wsNew.Sort
        .SortFields.Clear ' Clear any previous sort fields
        
        ' Sort by Column F (6th column), then by Column G (7th column), then by Column H (8th column)
        .SortFields.Add Key:=wsNew.Range("F2:F" & lastRow), Order:=xlAscending ' Column F
        .SortFields.Add Key:=wsNew.Range("G2:G" & lastRow), Order:=xlAscending ' Column G
        .SortFields.Add Key:=wsNew.Range("H2:H" & lastRow), Order:=xlAscending ' Column H
        
        ' Apply the sorting to the range from A1 to H[lastRow]
        .SetRange wsNew.Range("A1:O" & lastRow)
        
        ' Apply the sort
        .Header = xlYes ' Assuming your data includes headers
        .Apply
    End With
    
    ' Recalculate the last row after deletions
    lastRow = wsNew.Cells(wsNew.Rows.Count, "C").End(xlUp).Row
    

    ' Apply color formatting based on column D values
    For i = 2 To lastRow
        Set Rng = wsNew.Range(wsNew.Cells(i, "A"), wsNew.Cells(i, "P"))
        Select Case Mid(wsNew.Cells(i, "D").Value, 1, 1)
            Case "1"
                Rng.Interior.Color = RGB(217, 225, 242)     ' Light Blue
            Case "2"
                Rng.Interior.Color = RGB(255, 242, 204) ' Light Yellow
            Case "3"
                Rng.Interior.Color = RGB(242, 225, 217) ' levelStartRownk
            Case "D"
                Rng.Interior.Color = RGB(211, 177, 194) ' Light Purple
        End Select
    Next i
          
    wsNew.Cells(1, "L").Value = "MIN.AREA"
    wsNew.Cells(1, "M").Value = "MIN.PR.AM"
    wsNew.Cells(1, "N").Value = "PR.AM."
    wsNew.Cells(1, "O").Value = "MIN.COM"
    wsNew.Cells(1, "P").Value = "10%+"
    wsNew.Cells(1, "Q").Value = "1 BED"
    wsNew.Cells(1, "R").Value = "2 BED"
    wsNew.Cells(1, "S").Value = "3 BED"
        
    ' Apply Apartment Minimum Values
    For i = 2 To lastRow
        Select Case wsNew.Cells(i, "I")
            Case "1"
                wsNew.Cells(i, "L").Value = "45"
                With wsNew.Cells(i, "L").Font
                    .Color = RGB(90, 90, 90)
                    .Bold = True
                End With
                wsNew.Cells(i, "M").Value = "5"
                With wsNew.Cells(i, "M").Font
                    .Color = RGB(90, 90, 90)
                    .Bold = True
                End With
                wsNew.Cells(i, "N").Value = "5"
                wsNew.Cells(i, "O").Value = "5"
                With wsNew.Cells(i, "O").Font
                    .Color = RGB(90, 90, 90)
                    .Bold = True
                End With
                
                wsNew.Cells(i, "Q").Value = "1"
                
            Case "2"
                wsNew.Cells(i, "L").Value = "73"
                With wsNew.Cells(i, "L").Font
                    .Color = RGB(90, 90, 90)
                    .Bold = True
                End With
                wsNew.Cells(i, "M").Value = "7"
                With wsNew.Cells(i, "M").Font
                    .Color = RGB(90, 90, 90)
                    .Bold = True
                End With
                wsNew.Cells(i, "N").Value = "7"
                wsNew.Cells(i, "O").Value = "7"
                With wsNew.Cells(i, "O").Font
                    .Color = RGB(90, 90, 90)
                    .Bold = True
                End With
                
                wsNew.Cells(i, "R").Value = "1"
                
            Case "3"
                wsNew.Cells(i, "L").Value = "90"
                With wsNew.Cells(i, "L").Font
                    .Color = RGB(90, 90, 90)
                    .Bold = True
                End With
                wsNew.Cells(i, "M").Value = "9"
                With wsNew.Cells(i, "M").Font
                    .Color = RGB(90, 90, 90)
                    .Bold = True
                End With
                wsNew.Cells(i, "N").Value = "9"
                wsNew.Cells(i, "O").Value = "9"
                With wsNew.Cells(i, "O").Font
                    .Color = RGB(90, 90, 90)
                    .Bold = True
                End With
                
                wsNew.Cells(i, "S").Value = "1"
        End Select
    Next i
    
        
    
    'Add a +10% indicator to every row
    For i = 2 To lastRow
         areaExt = wsNew.Cells(i, "L").Value * 1.1
         areaCur = wsNew.Cells(i, "E").Value
         
         If areaCur > areaExt Then
            wsNew.Cells(i, "P").Value = "1"
         Else
            wsNew.Cells(i, "P").Value = "0"
         End If
    
    Next i
    
    
   
    
    'find the last row
    lastRow = wsNew.Cells(wsNew.Rows.Count, "C").End(xlUp).Row
        
    ' Insert 2 empty rows
    wsNew.Rows(2).Resize(2).Insert Shift:=xlDown
    
    'set up collections for tracking level and floor changes
    Set changeLevel = New Collection    'collection for tracking level changes
    Set changeBlock = New Collection    'collection for tracking block changes
    Set sumColumns = New Collection    'collection for columns to add
    
    'set up which columnds need to be summed
    sumColumns.Add "H"
    sumColumns.Add "I"
    sumColumns.Add "J"
    sumColumns.Add "K"
    sumColumns.Add "L"
    sumColumns.Add "M"
    sumColumns.Add "N"
    sumColumns.Add "O"
    sumColumns.Add "E"
    sumColumns.Add "P"
    
    'initiating the loop and strating parameters
    previousLevel = wsNew.Cells(4, "G").Value
    previousBlock = wsNew.Cells(4, "F").Value
    i = 4
    levelStartRow = 4
    blockStartRow = 4
    last = 1000
    'For i = 3 To lastRow
    Do While True
        currentLevel = wsNew.Cells(i, "G").Value
        currentBlock = wsNew.Cells(i, "F").Value
        
        If currentLevel <> previousLevel Then
            
            'record change rows
            
          
            ' Insert 3 empty rows
            wsNew.Rows(i).Resize(3).Insert Shift:=xlDown
            
            ' Clear fill color for the inserted rows
            wsNew.Rows(i).Resize(3).Interior.ColorIndex = -4142 ' No Fill
            
            ' Add SUM formulas for columns I-K up until the previous empty row
             ' sum up areas
             
                   
            For p = 1 To sumColumns.Count
            
            If p = 1 Then
            With wsNew.Cells(i, sumColumns(p))
                    .Formula = "=COUNTA(" & sumColumns(p) & i - 1 & ":" & sumColumns(p) & levelStartRow & ")"
                    .Font.Bold = True
            End With
            Else
            With wsNew.Cells(i, sumColumns(p))
                    .Formula = "=SUM(" & sumColumns(p) & i - 1 & ":" & sumColumns(p) & levelStartRow & ")"
                    .Font.Bold = True
            End With
            End If
                
            Next p
            
            ' sum up apartment types
            With wsNew.Cells(i, "T")
                .Formula = "=SUM(Q" & i - 1 & ":Q" & levelStartRow & ")"
                .Font.Bold = True
            End With
            With wsNew.Cells(i, "U")
                .Formula = "=SUM(R" & i - 1 & ":R" & levelStartRow & ")"
                .Font.Bold = True
            End With
            With wsNew.Cells(i, "V")
                .Formula = "=SUM(S" & i - 1 & ":S" & levelStartRow & ")"
                .Font.Bold = True
            End With
            'calculate %s
            With wsNew.Cells(i + 1, "T")
                .Formula = "=T" & i & "/H" & i
                .Font.Bold = False
                .NumberFormat = "0%"
            End With
            With wsNew.Cells(i + 1, "U")
                .Formula = "=U" & i & "/H" & i
                .Font.Bold = False
                .NumberFormat = "0%"
            End With
            With wsNew.Cells(i + 1, "V")
                .Formula = "=V" & i & "/H" & i
                .Font.Bold = False
                .NumberFormat = "0%"
            End With
            With wsNew.Cells(i + 1, "P")
                .Formula = "=P" & i & "/H" & i
                .Font.Bold = False
                .NumberFormat = "0%"
            End With
            With wsNew.Cells(i + 1, "K")
                .Formula = "=K" & i & "/H" & i
                .Font.Bold = False
                .NumberFormat = "0%"
            End With
            
            ' sum up total apartments
            With wsNew.Cells(i, "H")
                .Formula = "=COUNTA(H" & i - 1 & ":H" & levelStartRow & ")"
                .Font.Bold = True
            End With
            
            
            ' Apply borders for the block range (A to O, blockStartRow to i-1)
            Set Rng = wsNew.Range(wsNew.Cells(i - 1, "A"), wsNew.Cells(levelStartRow, "P"))
            
            With Rng.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Rng.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Rng.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThick
            End With
            With Rng.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThick
            End With
            ' Add a thick exterior border
            With Rng.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThick
            End With


            ' Add title to floors
            With wsNew.Cells(levelStartRow - 1, "F")
                .Value = "Block " & previousBlock & " Level " & previousLevel
                .Font.Bold = True
                .Font.Color = RGB(0, 176, 240)
                .Font.Name = "Calibri"
            End With



            changeLevel.Add i
            
            'if the blocks change
            
            If currentBlock <> previousBlock Then
            
                 i = i + 3
                 
                ' Insert 3 empty rows
                wsNew.Rows(i).Resize(3).Insert Shift:=xlDown
                
                changeBlock.Add i
                
                ' Clear fill color for the inserted rows
                wsNew.Rows(i).Resize(3).Interior.ColorIndex = -4142 ' No Fill
                
                ' add summary line
                Set Rng = wsNew.Range(wsNew.Cells(i, "A"), wsNew.Cells(i, "P"))
               
                With Rng.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThick
                End With
                
                Rng.Font.Bold = True
                
                With wsNew.Cells(i - 1, "F")
                    .Value = "Block " & previousBlock & " Summary"
                    .Font.Bold = True
                    .Font.Color = RGB(0, 176, 240)
                    .Font.Name = "Calibri"
                End With
                
                ' sum up apartment types
                With wsNew.Cells(i, "T")
                    .Formula = "=SUM(Q" & i - 1 & ":Q" & blockStartRow & ")"
                    .Font.Bold = True
                End With
                With wsNew.Cells(i, "U")
                    .Formula = "=SUM(R" & i - 1 & ":R" & blockStartRow & ")"
                    .Font.Bold = True
                End With
                With wsNew.Cells(i, "V")
                    .Formula = "=SUM(S" & i - 1 & ":S" & blockStartRow & ")"
                    .Font.Bold = True
                End With
                'calculate %s
                With wsNew.Cells(i + 1, "T")
                    .Formula = "=T" & i & "/H" & i
                    .Font.Bold = False
                    .NumberFormat = "0%"
                End With
                With wsNew.Cells(i + 1, "U")
                    .Formula = "=U" & i & "/H" & i
                    .Font.Bold = False
                    .NumberFormat = "0%"
                End With
                With wsNew.Cells(i + 1, "V")
                    .Formula = "=V" & i & "/H" & i
                    .Font.Bold = False
                    .NumberFormat = "0%"
                End With
                With wsNew.Cells(i + 1, "P")
                    .Formula = "=P" & i & "/H" & i
                    .Font.Bold = False
                    .NumberFormat = "0%"
                End With
                With wsNew.Cells(i + 1, "K")
                    .Formula = "=K" & i & "/H" & i
                    .Font.Bold = False
                    .NumberFormat = "0%"
                End With
                
                blockStartRow = i + 1
                
                For p = 1 To sumColumns.Count
                
                'Declare the formula and start writing it
                Dim blockFormulaString As String
                blockFormulaString = "=SUM("
                
                For q = 1 To changeLevel.Count
                    If q = changeLevel.Count Then
                        ' For the last item, don't add a comma after it
                        blockFormulaString = blockFormulaString & sumColumns(p)
                        blockFormulaString = blockFormulaString & changeLevel(q) & ")"
                    Else
                        ' For all other items, add a comma between cell references
                        blockFormulaString = blockFormulaString & sumColumns(p)
                        blockFormulaString = blockFormulaString & changeLevel(q) & ","
                    End If
                Next q
                
                wsNew.Range(sumColumns(p) & i).Formula = blockFormulaString
             
                Next p
                
                Set changeLevel = New Collection
                previousBlock = currentBlock
            End If
            
            
            
            
            
            i = i + 3 ' Skip the inserted empty rows
            levelStartRow = i
            lastRow = lastRow + 3
        End If
        
  
        
        
        previousLevel = currentLevel
                
        i = i + 1
        
        If wsNew.Cells(i - 1, 1).Value = 0 Then
            Exit Do
        End If
        
        If i > 100000 Then
            Exit Do
        End If
    Loop
    
    
    'start filling out the overall summary but offsetting the position
    i = i - 1
            
    'add a summary line
    Set Rng = wsNew.Range(wsNew.Cells(i, "A"), wsNew.Cells(i, "P"))
    
    With Rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    
    Rng.Font.Bold = True
    
    'add a title
    With wsNew.Cells(i - 1, "F")
        .Value = "Whole Scheme " & previousBlock & " Summary"
        .Font.Bold = True
        .Font.Color = RGB(0, 176, 240)
        .Font.Name = "Calibri"
    End With
    
    ' sum up apartment types
    With wsNew.Cells(i, "T")
        .Formula = "=SUM(Q" & i - 1 & ":Q" & "2" & ")"
        .Font.Bold = True
    End With
    With wsNew.Cells(i, "U")
        .Formula = "=SUM(R" & i - 1 & ":R" & "2" & ")"
        .Font.Bold = True
    End With
    With wsNew.Cells(i, "V")
        .Formula = "=SUM(S" & i - 1 & ":S" & "2" & ")"
        .Font.Bold = True
    End With
    'calculate %s
    With wsNew.Cells(i + 1, "T")
        .Formula = "=T" & i & "/H" & i
        .Font.Bold = False
        .NumberFormat = "0%"
    End With
    With wsNew.Cells(i + 1, "U")
        .Formula = "=U" & i & "/H" & i
        .Font.Bold = False
        .NumberFormat = "0%"
    End With
    With wsNew.Cells(i + 1, "V")
        .Formula = "=V" & i & "/H" & i
        .Font.Bold = False
        .NumberFormat = "0%"
    End With
    With wsNew.Cells(i + 1, "P")
        .Formula = "=P" & i & "/H" & i
        .Font.Bold = False
        .NumberFormat = "0%"
    End With
    With wsNew.Cells(i + 1, "K")
        .Formula = "=K" & i & "/H" & i
        .Font.Bold = False
        .NumberFormat = "0%"
    End With
    
    
    'start summing up the main stats
    For p = 1 To sumColumns.Count
        
        'Declare the formula and start writing it
        Dim projectFormulaString As String
        projectFormulaString = "=SUM("
        
        For q = 1 To changeBlock.Count
            If q = changeBlock.Count Then
                ' For the last item, don't add a comma after it
                projectFormulaString = projectFormulaString & sumColumns(p)
                projectFormulaString = projectFormulaString & changeBlock(q) & ")"
            Else
                ' For all other items, add a comma between cell references
                projectFormulaString = projectFormulaString & sumColumns(p)
                projectFormulaString = projectFormulaString & changeBlock(q) & ","
            End If
        Next q
        
        wsNew.Range(sumColumns(p) & i).Formula = projectFormulaString
    
    Next p
    
    

    
    ' Move columns H, F, and G to A, B, and C respectively
    
    wsNew.Columns("H:H").Cut
    wsNew.Columns("A:A").Insert Shift:=xlToRight
    wsNew.Columns("G:G").Cut
    wsNew.Columns("B:B").Insert Shift:=xlToRight
    wsNew.Columns("H:H").Cut
    wsNew.Columns("C:C").Insert Shift:=xlToRight
    wsNew.Columns("L:L").Cut
    wsNew.Columns("H:H").Insert Shift:=xlToRight
    
    'Delete columns D and E in the new sheet
    
    wsNew.Columns("D:E").Delete
    
    With wsNew
        .Range("A1:A" & .Cells(.Rows.Count, "A").End(xlUp).Row).HorizontalAlignment = xlCenter
        .Range("B1:B" & .Cells(.Rows.Count, "B").End(xlUp).Row).HorizontalAlignment = xlCenter
        .Range("C1:C" & .Cells(.Rows.Count, "C").End(xlUp).Row).HorizontalAlignment = xlCenter
        .Range("E1:E" & .Cells(.Rows.Count, "E").End(xlUp).Row).HorizontalAlignment = xlCenter
    End With
    
    
    
End Sub

