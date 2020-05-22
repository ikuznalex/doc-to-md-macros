

Sub Word2Wiki()
    
    Application.ScreenUpdating = False
    
    'Heading 1 to Heading 5
    ConvertParagraphStyle wdStyleHeading1, "# "
    ConvertParagraphStyle wdStyleHeading2, "## "
    ConvertParagraphStyle wdStyleHeading3, "### "
    ConvertParagraphStyle wdStyleHeading4, "#### "
    ConvertParagraphStyle wdStyleHeading5, "##### "
    
    ConvertItalic
    ConvertBold
    ConvertUnderline
    
    ConvertLists
    ConvertTables
    
    ' Copy to clipboard
    ActiveDocument.Content.Copy
    
    Application.ScreenUpdating = True
End Sub
 
Private Sub ConvertParagraphStyle(styleToReplace As WdBuiltinStyle, _
                                    preText As String)
    Dim normalStyle As Style
    Set normalStyle = ActiveDocument.Styles(wdStyleNormal)
    
    ActiveDocument.Select
    
    With Selection.Find
    
        .ClearFormatting
        .Style = ActiveDocument.Styles(styleToReplace)
        .Text = ""
        
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .Forward = True
        .Wrap = wdFindContinue
        
        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore preText
                End If
                
                .Style = normalStyle
            End With
        Loop
    End With
End Sub

Private Sub ConvertBold()
    ActiveDocument.Select
    
    With Selection.Find
    
        .ClearFormatting
        .Font.Bold = True
        .Text = ""
        
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .Forward = True
        .Wrap = wdFindContinue
        
        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
              .Font.Bold = False
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore "**"
                    .InsertAfter "**"
                End If
                
                .Font.Bold = False
            End With
        Loop
    End With
End Sub
 
Private Sub ConvertItalic()
    ActiveDocument.Select
    
    With Selection.Find
    
        .ClearFormatting
        .Font.Italic = True
        .Text = ""
        
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .Forward = True
        .Wrap = wdFindContinue
        
        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Font.Italic = False
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore "_"
                    .InsertAfter "_"
                End If
                
                .Font.Italic = False
            End With
        Loop
    End With
End Sub
 
Private Sub ConvertUnderline()
    ActiveDocument.Select
    
    With Selection.Find
    
        .ClearFormatting
        .Font.Underline = True
        .Text = ""
        
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .Forward = True
        .Wrap = wdFindContinue
        
        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Font.Underline = False
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                                       
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore "<u>"
                    .InsertAfter "</u>"
                End If
                
                .Font.Underline = False
            End With
        Loop
    End With
End Sub
 
Private Sub ConvertLists()
   Dim para As Paragraph
   Dim i As Long
    For Each para In ActiveDocument.ListParagraphs
        With para.Range
            .InsertBefore " "
            For i = 1 To .ListFormat.ListLevelNumber
                If .ListFormat.ListType = wdListBullet Then
                    .InsertBefore "*"
                End If
                If .ListFormat.ListType = wdListSimpleNumbering Then
                     ' This for numbered list
                    .InsertBefore CStr(i) + "."
                End If
            Next i
            .ListFormat.RemoveNumbers
        End With
    Next para
End Sub
 
Private Sub ConvertTables()

    Dim myRange As Word.Range
    Dim tTable  As Word.Table
    Dim tRow    As Word.Row
    Dim tCell   As Word.Cell
    Dim strText As String
    
    Dim i   As Long
    Dim j   As Long
    Dim k   As Long
    Dim l   As Long
    
    For Each tTable In ActiveDocument.Tables
            
            'Memorize table text
            ReDim x(1 To tTable.Rows.Count, 1 To tTable.Columns.Count)
            i = 0
            For Each tRow In tTable.Rows
                i = i + 1
                j = 0
                For Each tCell In tRow.Cells
                    j = j + 1
                    strText = tCell.Range.Text
                    x(i, j) = Left(strText, Len(strText) - 2)
                Next tCell
            Next tRow
             
           
            
            'Delete table and position after table
            Set myRange = tTable.Range
            myRange.Collapse Direction:=wdCollapseEnd
            tTable.Delete
            
            'Rewrite table with memorized text
            'Add one more row for header
            i = i + 1
            myRange.InsertParagraphAfter
            For k = 1 To i
                For l = 1 To j
                    If k = 1 Then
                        myRange.InsertAfter " | " + x(k, l)
                    ElseIf k = 2 Then
                        myRange.InsertAfter " |--"
                    Else
                        myRange.InsertAfter " | " + x(k - 1, l)
                    End If
                Next l
                myRange.InsertParagraphAfter
            Next k
            myRange.InsertParagraphAfter
            
    Next tTable

End Sub

