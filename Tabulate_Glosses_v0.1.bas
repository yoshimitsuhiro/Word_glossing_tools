Attribute VB_Name = "Tabulate_Glosses"
Sub Tabulate_Glosses()

    ' Declare only used variables
    Dim paras As Paragraphs
    Dim paraRanges() As Range
    Dim lineTexts() As String
    Dim words() As Variant
    Dim numWords() As Integer
    Dim numWordsPerLine As Integer
    Dim numParas As Integer
    Dim wordLengths() As Variant
    Dim wordWidths() As Variant
    Dim maxWordWidths() As Single
    Dim tabStops() As Single
    Dim i As Integer, j As Integer

    Dim wordRange As Range
    Dim rangeText As String
    Dim rangeLength As Integer
    Dim startChar As Integer
    Dim startPosition As Single
    Dim endPosition As Single
    Dim widthInPoints As Single
    Dim widthInMM As Single
    
    Dim indentStr As String
    Dim indentInPoints As Single
    Dim indentInMM As Single
    Dim interval As Single
    Dim firstIndentInPoints As Single
    Dim firstIndentInMM As Single
    Dim leftIndentInPoints As Single
    Dim leftIndentInMM As Single

    Dim result As Boolean
    Dim output As String

    Application.ScreenUpdating = False

    ' Prompt user for indent and interval
    result = Tabulate_Glosses_Prompt.ShowForm(indentStr, interval)
    If Not result Then
        MsgBox "Operation cancelled."
        End
    End If

    If indentStr = "Auto" Then
        indentInPoints = Selection.Paragraphs(1).Format.LeftIndent
        indentInMM = indentInPoints * 0.352778
    Else
        indentInMM = CSng(indentStr)
        indentInPoints = indentInMM / 0.352778
    End If

    Set paras = Selection.Paragraphs
    numParas = paras.Count
    ReDim paraRanges(1 To numParas)

    ' Format and clean each paragraph
    For i = 1 To numParas
        DoEvents
        Set paraRanges(i) = paras(i).Range
        With paraRanges(i)
            .Paragraphs(1).tabStops.ClearAll
            With .Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = True
                .Text = vbTab
                .Replacement.Text = " "
                .Execute Replace:=wdReplaceAll
                .Text = "[ ]{2,}"
                .Replacement.Text = " "
                .Execute Replace:=wdReplaceAll
            End With
            If Left(.Text, 1) = " " Then .Characters(1).Text = ""
            If Right(.Text, 2) = " " & vbCr Then .Characters(.Characters.Count - 1).Text = ""
        End With
    Next i

    ' Remove line break chars
    ReDim lineTexts(1 To numParas)
    For i = 1 To numParas
        lineTexts(i) = paraRanges(i).Text
        Do While Right(lineTexts(i), 1) = Chr(13) Or Right(lineTexts(i), 1) = Chr(11)
            lineTexts(i) = Left(lineTexts(i), Len(lineTexts(i)) - 1)
        Loop
    Next i

    ' Split each line into words
    ReDim words(1 To numParas)
    ReDim numWords(1 To numParas)
    For i = 1 To numParas
        words(i) = Split(lineTexts(i), " ")
        numWords(i) = UBound(words(i)) + 1
    Next i

    ' Check to see if number of elements is the same for each line
    numWordsPerLine = numWords(1)
    For i = 1 To UBound(numWords)
        If numWords(i) <> numWordsPerLine Then
            MsgBox "The number of elements must be the same across all lines."
            Exit Sub
        End If
    Next i
    
    If numWordsPerLine < 2 Then
        MsgBox "There must be at least two elements on each line."
        Exit Sub
    End If

    ' Get word lengths (in characters) and widths (in MM)
    ReDim wordLengths(1 To numParas, 1 To numWordsPerLine)
    ReDim wordWidths(1 To numParas, 1 To numWordsPerLine)

    For i = 1 To numParas
        startChar = 0
        With paraRanges(i).Paragraphs(1).Format
            firstIndentInPoints = .FirstLineIndent
            firstIndentInMM = firstIndentInPoints * 0.352778
            leftIndentInPoints = .LeftIndent
            leftIndentInMM = leftIndentInPoints * 0.352778
        End With
        For j = 1 To numWordsPerLine
            Set wordRange = paraRanges(i).Duplicate
            wordLengths(i, j) = Len(words(i)(j - 1))
            wordRange.MoveStart wdCharacter, startChar
            wordRange.MoveEnd wdCharacter, wordRange.Characters.Count * -1 + wordLengths(i, j)
            startPosition = wordRange.Information(wdHorizontalPositionRelativeToPage)
            wordRange.Collapse Direction:=wdCollapseEnd
            endPosition = wordRange.Information(wdHorizontalPositionRelativeToPage)
            widthInPoints = endPosition - startPosition
            widthInMM = widthInPoints * 0.352778
            wordWidths(i, j) = widthInMM
            If j = 1 Then wordWidths(i, j) = wordWidths(i, j) - leftIndentInMM - firstIndentInMM
            startChar = startChar + wordLengths(i, j) + 1
        Next j
    Next i

    ' Find widest element in each line
    ReDim maxWordWidths(1 To numWordsPerLine - 1)
    For j = 1 To numWordsPerLine - 1
        For i = 1 To numParas
            If wordWidths(i, j) > maxWordWidths(j) Then maxWordWidths(j) = wordWidths(i, j)
        Next i
    Next j

    ' Calculate tab stops
    ReDim tabStops(1 To numWordsPerLine - 1)
    For j = 1 To numWordsPerLine - 1
        If j = 1 Then
            tabStops(j) = indentInMM + maxWordWidths(j) + interval
        Else
            tabStops(j) = tabStops(j - 1) + maxWordWidths(j) + interval
        End If
    Next j

    ' Output tab stops
    output = "Tab Stops:" & vbCrLf
    For i = 1 To UBound(tabStops)
        output = output & "Element " & i & ": " & tabStops(i) & vbCrLf
    Next i
    MsgBox output

    ' Reformat selection
    For i = 1 To numParas
        paraRanges(i).Paragraphs(1).Format.LeftIndent = indentInPoints
    Next i

    For i = 1 To numParas
        With paraRanges(i).Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .Text = " "
            .Replacement.Text = vbTab
            .Execute Replace:=wdReplaceAll
        End With
    Next i

    For i = 1 To numParas
        For j = 1 To numWordsPerLine - 1
            paraRanges(i).ParagraphFormat.tabStops.Add _
                Position:=MillimetersToPoints(tabStops(j)), _
                Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
        Next j
    Next i

    Application.ScreenUpdating = True
    DoEvents

End Sub


