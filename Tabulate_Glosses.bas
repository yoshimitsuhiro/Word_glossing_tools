Attribute VB_Name = "Tabulate_Glosses"
Sub Tabulate_Glosses()

    ' Declare only used variables
    Dim selectionRange As Range
    Dim paraRanges() As Range
    Dim lineTexts() As String
    Dim splitWords() As Variant
    Dim words() As Variant
    Dim numWords() As Integer
    Dim numWordsPerLine As Integer
    Dim numParas As Integer
    Dim wordLengths() As Variant
    Dim wordWidths() As Variant
    Dim maxWordWidths() As Single
    Dim tabStops() As Single
    Dim i As Integer, j As Integer, k As Integer
    
    Dim wordRange As Range
    Dim rangeText As String
    Dim rangeLength As Integer
    Dim char As String
    Dim combiningChars As Integer
    Dim startChar As Integer
    Dim startPosition As Single
    Dim endPosition As Single
    Dim widthInPoints As Single
    Dim widthInMM As Single
    Dim combinedWithInPoints As Single
    Dim combinedWithInMM As Single
    Dim usableWidthInPoints As Single
    Dim usableWidthInMM As Single
    
    Dim indentStr As String
    Dim indentInPoints As Single
    Dim indentInMM As Single
    Dim interval As Single
    Dim intervalStr As String
    Dim autoInterval As Single
    Dim maxInterval As Single
    Dim firstIndentInPoints As Single
    Dim firstIndentInMM As Single
    Dim leftIndentInPoints As Single
    Dim leftIndentInMM As Single

    Dim result As Boolean
    Dim output As String

    Application.ScreenUpdating = False

    ' Set default values for indent and interval
    indentStr = "Auto" ' Set to "Auto" to calculate indent based on first line
    intervalStr = "Auto" ' Set to "Auto" to calculate widest interval without linewrapping
    maxInterval = 10 ' The maximum interval auto will allow

    ' Prompt user for indent and interval
    result = Tabulate_Glosses_Prompt.ShowForm(indentStr, intervalStr, maxInterval)
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

    Set selectionRange = Selection.Range

    numParas = selectionRange.Paragraphs.Count
    ReDim paraRanges(1 To numParas)

    ' Format and clean each paragraph
    For i = 1 To numParas
        Set paraRanges(i) = selectionRange.Paragraphs(i).Range
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
    ReDim splitWords(1 To numParas)
    ReDim numWords(1 To numParas)
    For i = 1 To numParas
        splitWords(i) = Split(lineTexts(i), " ")
        numWords(i) = UBound(splitWords(i)) + 1
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

    ' Create 2d array of lines and words
    ReDim words(1 To numParas, 1 To numWordsPerLine)
    For i = 1 To numParas
        For j = 1 To numWordsPerLine
            words(i, j) = splitWords(i)(j - 1)
        Next j
    Next i

'    ' Output word list
'    output = "Words:" & vbCrLf
'    For i = 1 To numParas
'        output = output & "Line " & i & ": "
'        For j = 1 To numWordsPerLine
'            output = output & words(i, j) & " "
'        Next j
'        output = output & vbCrLf
'    Next i
'
'    Tabulate_Glosses_Results.ShowText output

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
            wordLengths(i, j) = Len(words(i, j))
            
            'Test for combining characters
            combiningChars = 0
            For k = 1 To wordLengths(i, j)
                char = Mid(words(i, j), k, 1)
                If CombiningChar(char) Then
                    combiningChars = combiningChars + 1
                End If
            Next k
            wordLengths(i, j) = wordLengths(i, j) - combiningChars
            
            wordRange.MoveStart wdCharacter, startChar
            wordRange.MoveEnd wdCharacter, wordRange.Characters.Count * -1 + wordLengths(i, j)
            
            startPosition = wordRange.Information(wdHorizontalPositionRelativeToPage)
            wordRange.Collapse Direction:=wdCollapseEnd
            endPosition = wordRange.Information(wdHorizontalPositionRelativeToPage)
            
            widthInPoints = endPosition - startPosition
            widthInMM = widthInPoints * 0.352778
            wordWidths(i, j) = widthInMM
            
            If j = 1 Then
                wordWidths(i, j) = wordWidths(i, j) - leftIndentInMM - firstIndentInMM
            End If
            
            startChar = startChar + wordLengths(i, j) + 1
            currentPos = wordStart + wordLengths(i, j)
        Next j
    Next i

    ' Find widest element in each line
    ReDim maxWordWidths(1 To numWordsPerLine)
    For j = 1 To numWordsPerLine
        For i = 1 To numParas
            If wordWidths(i, j) > maxWordWidths(j) Then maxWordWidths(j) = wordWidths(i, j)
        Next i
    Next j

    If intervalStr = "Auto" Then
        ' Get usable width of page
        Get_Usable_Width usableWidthInPoints, usableWidthInMM
        usableWidthInPoints = usableWidthInPoints - indentInPoints
        usableWidthInMM = usableWidthInMM - indentInMM
        
        ' Auto calculate interval
        combinedWidthInPoints = 0
        combinedWidthInMM = 0
        For i = 1 To numWordsPerLine
            combinedWidthInMM = combinedWidthInMM + maxWordWidths(i)
        Next i
        combinedWidthInPoints = combinedWidthInMM / 0.352778
        
        usableWidthInMM = usableWidthInMM - combinedWidthInMM - 1
        
        interval = usableWidthInMM / (numWordsPerLine - 1)
        
        If interval > maxInterval Then
            interval = maxInterval
        End If
    Else
        interval = CSng(intervalStr)
    End If
    
    ' Calculate tab stops
    ReDim tabStops(1 To numWordsPerLine - 1)
    For j = 1 To numWordsPerLine - 1
        If j = 1 Then
            tabStops(j) = indentInMM + maxWordWidths(j) + interval
        Else
            tabStops(j) = tabStops(j - 1) + maxWordWidths(j) + interval
        End If
    Next j

    ' Output results
    output = "Interval = " & Round(interval, 1) & vbCrLf & vbCrLf & "Tab Stops:" & vbCrLf
    For i = 1 To numWordsPerLine
        If i = 1 Then
            output = output & "Element " & i & ": " & Round(indentInMM, 1) & " (indent)" & vbCrLf
        ElseIf i = numWordsPerLine Then
            output = output & "Element " & i & ": 0 (final element)"
        Else
            output = output & "Element " & i & ": " & Round(tabStops(i - 1), 1) & vbCrLf
        End If
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

Function CombiningChar(char As String) As Boolean
    Dim code As Long
    code = AscW(char)
    CombiningChar = (code >= &H300 And code <= &H36F)
End Function

Sub Get_Usable_Width(ByRef usableWidthInPoints As Single, ByRef usableWidthInMM As Single)
    Dim section As section
    Dim totalWidthInPoints As Single
    Dim leftMarginInPoints As Single
    Dim rightMarginInPoints As Single

    ' Get the section of the current selection
    Set section = Selection.Range.Sections(1)

    With section.PageSetup
        totalWidthInPoints = .PageWidth
        leftMarginInPoints = .LeftMargin
        rightMarginInPoints = .RightMargin
    End With

    usableWidthInPoints = totalWidthInPoints - leftMarginInPoints - rightMarginInPoints
    usableWidthInMM = usableWidthInPoints * 0.352778 ' Convert points to mm

End Sub

Sub Wait(ByVal Seconds As Single)
    Dim CurrentTimer As Variant
    CurrentTimer = Timer
    Do While Timer < CurrentTimer + Seconds
    Loop
End Sub
