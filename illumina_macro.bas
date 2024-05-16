
'below pulled from http://www.cpearson.com/excel/findall.aspx
Function FindAll(SearchRange As Range, _
                FindWhat As Variant, _
               Optional LookIn As XlFindLookIn = xlValues, _
                Optional LookAt As XlLookAt = xlWhole, _
                Optional SearchOrder As XlSearchOrder = xlByRows, _
                Optional MatchCase As Boolean = False, _
                Optional BeginsWith As String = vbNullString, _
                Optional EndsWith As String = vbNullString, _
                Optional BeginEndCompare As VbCompareMethod = vbTextCompare) As Range
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FindAll
' This searches the range specified by SearchRange and returns a Range object
' that contains all the cells in which FindWhat was found. The search parameters to
' this function have the same meaning and effect as they do with the
' Range.Find method. If the value was not found, the function return Nothing. If
' BeginsWith is not an empty string, only those cells that begin with BeginWith
' are included in the result. If EndsWith is not an empty string, only those cells
' that end with EndsWith are included in the result. Note that if a cell contains
' a single word that matches either BeginsWith or EndsWith, it is included in the
' result.  If BeginsWith or EndsWith is not an empty string, the LookAt parameter
' is automatically changed to xlPart. The tests for BeginsWith and EndsWith may be
' case-sensitive by setting BeginEndCompare to vbBinaryCompare. For case-insensitive
' comparisons, set BeginEndCompare to vbTextCompare. If this parameter is omitted,
' it defaults to vbTextCompare. The comparisons for BeginsWith and EndsWith are
' in an OR relationship. That is, if both BeginsWith and EndsWith are provided,
' a match if found if the text begins with BeginsWith OR the text ends with EndsWith.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim FoundCell As Range
Dim FirstFound As Range
Dim LastCell As Range
Dim ResultRange As Range
Dim XLookAt As XlLookAt
Dim Include As Boolean
Dim CompMode As VbCompareMethod
Dim Area As Range
Dim MaxRow As Long
Dim MaxCol As Long
Dim BeginB As Boolean
Dim EndB As Boolean


CompMode = BeginEndCompare
If BeginsWith <> vbNullString Or EndsWith <> vbNullString Then
    XLookAt = xlPart
Else
    XLookAt = LookAt
End If

' this loop in Areas is to find the last cell
' of all the areas. That is, the cell whose row
' and column are greater than or equal to any cell
' in any Area.

For Each Area In SearchRange.Areas
    With Area
        If .Cells(.Cells.Count).Row > MaxRow Then
            MaxRow = .Cells(.Cells.Count).Row
        End If
        If .Cells(.Cells.Count).Column > MaxCol Then
            MaxCol = .Cells(.Cells.Count).Column
        End If
    End With
Next Area
Set LastCell = SearchRange.Worksheet.Cells(MaxRow, MaxCol)

On Error GoTo 0
Set FoundCell = SearchRange.Find(what:=FindWhat, _
        after:=LastCell, _
        LookIn:=LookIn, _
        LookAt:=XLookAt, _
        SearchOrder:=SearchOrder, _
        MatchCase:=MatchCase)

If Not FoundCell Is Nothing Then
    Set FirstFound = FoundCell
    Do Until False ' Loop forever. We'll "Exit Do" when necessary.
        Include = False
        If BeginsWith = vbNullString And EndsWith = vbNullString Then
            Include = True
        Else
            If BeginsWith <> vbNullString Then
                If StrComp(Left(FoundCell.Text, Len(BeginsWith)), BeginsWith, BeginEndCompare) = 0 Then
                    Include = True
                End If
            End If
            If EndsWith <> vbNullString Then
                If StrComp(Right(FoundCell.Text, Len(EndsWith)), EndsWith, BeginEndCompare) = 0 Then
                    Include = True
                End If
            End If
        End If
        If Include = True Then
            If ResultRange Is Nothing Then
                Set ResultRange = FoundCell
            Else
                Set ResultRange = Application.Union(ResultRange, FoundCell)
            End If
        End If
        Set FoundCell = SearchRange.FindNext(after:=FoundCell)
        If (FoundCell Is Nothing) Then
            Exit Do
        End If
        If (FoundCell.Address = FirstFound.Address) Then
            Exit Do
        End If

    Loop
End If
    
Set FindAll = ResultRange

End Function
Sub remove_nondata_rows()
'
' removes nondata rows
'

'
    Range("A1").Select
    Dim rng1 As Range
    Set rng1 = Range("A:A").Find("https", , xlValues, xlPart)
    rng1.Select
    ActiveSheet.Range(Selection, Selection).EntireRow.Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Delete Shift:=xlUp


'
' bolds header

    Rows("1:1").Select
    Selection.Font.Bold = True
    Range("A1").Select
End Sub
Sub delete_unwanted_columns()
'
' delete_unwanted_columns Macro
'

'
    
    Dim rng2 As Range
    Set rng2 = FindAll(Range("1:1"), "p-value", xlValues, xlPart)
    Dim rng3 As Range
    Set rng3 = FindAll(Range("1:1"), "type", xlValues, xlPart)
    
    With ActiveSheet
        Set Rng = Union(rng2, rng3)
    End With
    
    Rng.EntireColumn.Select
    Selection.Delete
    Range("A1").Select
    
End Sub


Sub add_adjusted_score_columns()
'
' add_adjusted_score_columns Macro
'

'

    Dim rng4 As Range
    Set rng4 = FindAll(Range("1:1"), "Score", xlValues, xlPart, , True)
    
    rng4.EntireColumn.Select
    Selection.Insert Shift:=xlToRight
    
    
    rng4.Select
    Dim rng5 As Range
    Set rng5 = rng4.Offset(0, -1)
    rng5.Select
    
    
    Selection.Formula = "Adjusted Score"
    Selection.EntireColumn.Interior.Color = 65535
    
'    Selection.Interior.Color = 65535
    
'    ActiveCell.Select
'    ActiveCell.FormulaR1C1 = "Adjusted Score"
'    ActiveCell.Columns("A:A").EntireColumn.Select

End Sub
