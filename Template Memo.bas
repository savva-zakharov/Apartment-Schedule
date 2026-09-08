Option Explicit

' ============================================================================
' TEMPLATE MEMO MODULE
' Stamps a short set of notes onto the template sheet, so whoever edits the
' mapping rows sees the rules without leaving Excel.
'
' Non-destructive: a note sits on top of a cell and never changes its value.
' Re-run WriteTemplateMemo to refresh the notes after editing the text below,
' or ClearTemplateMemo to take them off again.
'
' The long version of these notes is TEMPLATE-ROW.md in the repository.
' ============================================================================

' Cells the notes are attached to
Private Const MEMO_CELLS As String = "A9,A18,A29,AA5,X3"

Public Sub WriteTemplateMemo()
    Dim wsTemplate As Worksheet
    Dim sigma As String
    Dim mappingNote As String
    Dim failures As Long

    Set wsTemplate = GetTemplateSheet()
    If wsTemplate Is Nothing Then Exit Sub

    ' Written as a code point, not a literal: VBA exports .bas files as
    ' Windows-1252, which has no sigma to export
    sigma = ChrW(&H3A3)

    ' The part that is true of all three mapping rows
    mappingNote = _
        "Each cell names one column of the output sheet. The name must" & vbLf & _
        "match an attribute tag in the exported .txt. Position sets the" & vbLf & _
        "column order. Leave a name out to drop the column." & vbLf & _
        "This row is read, never printed." & vbLf & vbLf & _
        "Add to the end of a cell:" & vbLf & _
        "    " & sigma & "    total the column on the summary rows" & vbLf & _
        "    %    add a percentage-of-units row underneath" & vbLf & _
        "Both together is fine, in either order." & vbLf & _
        "Type the sigma with =UNICHAR(931)." & vbLf & vbLf & _
        "NO is always the unit count, marked or not." & vbLf & _
        "Marking any one cell switches off the built-in default list," & vbLf & _
        "so mark every column you want in one pass." & vbLf & vbLf & _
        "Full notes: TEMPLATE-ROW.md"

    If Not SetCellNote(wsTemplate, "A9", _
        "MAPPING ROW - Long schedule, one row per unit" & vbLf & _
        "Read by GenerateUnitSchedule. Printed headings: rows 1-8." & vbLf & vbLf & _
        mappingNote) Then failures = failures + 1

    If Not SetCellNote(wsTemplate, "A18", _
        "MAPPING ROW - Short schedule, totals only" & vbLf & _
        "Read by GenerateUnitShort. Printed headings: rows 10-17." & vbLf & vbLf & _
        mappingNote & vbLf & vbLf & _
        "A field can only appear here if it is also named in row 9." & vbLf & _
        "MIX is where the bed/type tally block goes; TMIX is the" & vbLf & _
        "optional total-mix block. Both widen to fit the scheme.") Then failures = failures + 1

    If Not SetCellNote(wsTemplate, "A29", _
        "MAPPING ROW - Unit types block, one row per unique type" & vbLf & _
        "Read by UnitTypes. Printed headings: rows 20-28." & vbLf & vbLf & _
        mappingNote & vbLf & vbLf & _
        "A cell holding only % is a column named %, not a marker.") Then failures = failures + 1

    If Not SetCellNote(wsTemplate, "AA5", _
        "Full path to the .txt or .csv exported from AutoCAD." & vbLf & _
        "Surrounding quotes are stripped, so a path pasted straight" & vbLf & _
        "out of Explorer works as-is." & vbLf & vbLf & _
        "Every macro reads this one cell.") Then failures = failures + 1

    If Not SetCellNote(wsTemplate, "X3", _
        "Pattern used to reduce a TYPE value to the key the types" & vbLf & _
        "block groups on. Leave it empty to group on the whole value.") Then failures = failures + 1

    If failures = 0 Then
        MsgBox "Template notes written to " & MEMO_CELLS & ".", vbInformation
    Else
        MsgBox failures & " of the notes could not be written." & vbCrLf & _
               "Is the template sheet protected?", vbExclamation
    End If
End Sub

Public Sub ClearTemplateMemo()
    Dim wsTemplate As Worksheet
    Dim cellRefs As Variant
    Dim i As Long

    Set wsTemplate = GetTemplateSheet()
    If wsTemplate Is Nothing Then Exit Sub

    cellRefs = Split(MEMO_CELLS, ",")

    On Error Resume Next
    For i = LBound(cellRefs) To UBound(cellRefs)
        With wsTemplate.Range(Trim(cellRefs(i)))
            If Not .Comment Is Nothing Then .Comment.Delete
        End With
    Next i
    On Error GoTo 0

    MsgBox "Template notes removed.", vbInformation
End Sub

' ----------------------------------------------------------------------------
' Helpers
' ----------------------------------------------------------------------------

Private Function GetTemplateSheet() As Worksheet
    On Error Resume Next
    Set GetTemplateSheet = ThisWorkbook.Sheets("template")
    On Error GoTo 0

    If GetTemplateSheet Is Nothing Then
        MsgBox "No sheet named 'template' in this workbook.", vbExclamation
    End If
End Function

' Replace the note on one cell. Returns False if the cell would not take it,
' which is usually sheet protection.
Private Function SetCellNote(ws As Worksheet, cellRef As String, _
                             noteText As String) As Boolean
    Dim boxArea As Double

    On Error GoTo Failed

    With ws.Range(cellRef)
        If Not .Comment Is Nothing Then .Comment.Delete

        .AddComment
        .Comment.Text Text:=noteText
        .Comment.Visible = False

        With .Comment.Shape
            .TextFrame.AutoSize = True

            ' AutoSize lays long text out as one very wide line - clamp the
            ' width and hand back the height it needs at that width
            If .Width > 300 Then
                boxArea = .Width * .Height
                .Width = 300
                .Height = (boxArea / 300) * 1.5
            End If
        End With
    End With

    SetCellNote = True
    Exit Function

Failed:
    Debug.Print "Could not write note to " & cellRef & ": " & Err.Description
    SetCellNote = False
End Function
