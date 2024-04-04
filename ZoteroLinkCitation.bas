Private Function isSupportedStyle() As Boolean
    isSupportedStyle = ExtractZoteroStyleId() = "molecular-plant"
End Function

Private Function ExtractZoteroStyleId() As String
    Dim zoteroData As String
    Dim styleId As String

    zoteroData = ActiveDocument.CustomDocumentProperties("ZOTERO_PREF_1").Value

    styleId = ExtractStyleIdFromXML(zoteroData)

    Dim segments() As String
    segments = Split(styleId, "/")
    styleId = segments(UBound(segments))

    ExtractZoteroStyleId = styleId
End Function

Private Function ExtractStyleIdFromXML(xmlContent As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim styleId As String
    
    startPos = InStr(xmlContent, "style id=")
    If startPos > 0 Then
        startPos = startPos + Len("style id=") + 1
        endPos = InStr(startPos, xmlContent, """")
        If endPos > startPos Then
            styleId = Mid(xmlContent, startPos, endPos - startPos)
        End If
    End If
    
    ExtractStyleIdFromXML = styleId
End Function

Private Function RemoveSpecifiedHtmlTags(inputString As String, tagsToRemove As Variant) As String
    Dim regex As Object
    Dim tag As Variant

    Set regex = CreateObject("VBScript.RegExp")

    For Each tag In tagsToRemove
        With regex
            .Global = True
            .IgnoreCase = True
            .Pattern = "</?" & tag & ".*?>"
            inputString = .Replace(inputString, "")
        End With
    Next tag
    
    RemoveSpecifiedHtmlTags = inputString
End Function

Private Function RemoveHtmlTags(inputString As String) As String
    Dim tagsToRemove() As Variant
    tagsToRemove = Array("i", "sub", "sup")

    RemoveHtmlTags = RemoveSpecifiedHtmlTags(inputString, tagsToRemove)
End Function

Private Function CleanTitleAnchor(ByVal title As String) As String
    Dim charsToReplace As Variant
    Dim replacementChar As String
    Dim i As Integer
    
    ' List of characters to replace
    charsToReplace = Array(" ", "#", "&", ":", ",", "-", ".", "(", ")", "?", "!")
    ' Character used for replacement
    replacementChar = "_"
    
    ' Loop through the array of characters to replace
    For i = 0 To UBound(charsToReplace)
        title = Replace(title, charsToReplace(i), replacementChar)
    Next i
    
    ' Truncate to 40 characters
    CleanTitleAnchor = Left(title, 40)
End Function

Private Function ExtractCitations(ByVal inputString As String) As String()
    Dim parts() As String
    Dim i As Integer

    inputString = Mid(inputString, 2, Len(inputString) - 2)

    ' Split the string by semicolon
    parts = Split(inputString, ";")

    ' Trim spaces from each part
    For i = LBound(parts) To UBound(parts)
        parts(i) = Trim(parts(i))
    Next i

    ' Return the array of trimmed parts
    ExtractCitations = parts
End Function

Private Function ExtractTitlesFromJSON(jsonString As String) As String()
    Dim startPos As Long
    Dim endPos As Long
    Dim currentPos As Long
    Dim Titles() As String
    Dim titleCount As Integer
    Dim currentTitle As String

    ' Initialize variables
    titleCount = -1 ' Start from -1 so that the first title will be at index 0
    currentPos = 1

    ' Loop through the string to find titles
    Do
        ' Find the start of the title
        startPos = InStr(currentPos, jsonString, """title"":""")
        If startPos = 0 Then Exit Do

        ' Adjust position to start of title text
        startPos = startPos + Len("""title"":""")

        ' Find the end of the title
        endPos = InStr(startPos, jsonString, """")
        If endPos = 0 Then Exit Do

        ' Extract the title
        currentTitle = Mid(jsonString, startPos, endPos - startPos)

        ' Increment titleCount for 0-based array
        titleCount = titleCount + 1

        ' Resize the array to accommodate the new title
        ' Note: ReDim Preserve can only resize the last dimension of a multi-dimensional array
        ReDim Preserve Titles(titleCount)
        
        ' Assign the currentTitle to the correct index in the array
        Titles(titleCount) = RemoveHtmlTags(currentTitle)

        ' Update the current position for the next search
        currentPos = endPos + 1
    Loop

    ' Return the array of titles
    ExtractTitlesFromJSON = Titles
End Function

Private Sub AssertArrayLengthsEqual(array1 As Variant, array2 As Variant)
    If Not UBound(array1) - LBound(array1) = UBound(array2) - LBound(array2) Then
        MsgBox "Assertion Failed: The lengths of the two arrays are not equal.", vbCritical, "Assertion Failed"
        Err.Raise Number:=vbObjectError + 513, Description:="Array length assertion failed."
    End If
End Sub

Public Sub ZoteroLinkCitation()
    If Not isSupportedStyle() Then
        MsgBox "The current citation style is not yet supported: " & ExtractZoteroStyleId(), vbCritical, "Error"
        Exit Sub
    End if

    ' Declare variables for start and end positions
    Dim nStart&, nEnd&
    ' Capture current selection positions
    nStart = Selection.Start
    nEnd = Selection.End
    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    
    ' Variables for processing
    Dim citation As String
    Dim title As String
    Dim titleAnchor As String
    Dim style As String
    Dim fieldCode As String
    Dim numOrYear As String
    Dim pos&, n1&, n2&
    Dim Response As Integer
    
    ' Show field codes to manipulate Zotero fields
    ActiveWindow.View.ShowFieldCodes = True
    ' Prepare to find Zotero bibliography field,
    ' "^d" instructing Word to search for any field in the document.
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "^d ADDIN ZOTERO_BIBL"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    ' Execute find operation
    Selection.Find.Execute

    ' Bookmark the Zotero bibliography for later reference
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:="Zotero_Bibliography"
        .DefaultSorting = wdSortByName
        .ShowHidden = True
    End With
    ' Hide field codes
    ActiveWindow.View.ShowFieldCodes = False

    ' Iterate through all fields in the document
    For Each aField In ActiveDocument.Fields
        ' Check if the field is a Zotero citation
        If InStr(aField.Code, "ADDIN ZOTERO_ITEM") > 0 Then
            If MsgBox("Found " & aField.Result.Text, vbYesNo + vbQuestion, "Continue?") = vbNo Then
                Exit For
            End If

            Dim Citations() As String
            Dim Titles() As String
            Citations = ExtractCitations(aField.Result.Text)
            Titles = ExtractTitlesFromJSON(aField.Code)
            AssertArrayLengthsEqual Citations, Titles

            For i = 0 To UBound(Citations)
                citation = Citations(i)
                title = Titles(i)

                ' Create a sanitized anchor name from the title
                titleAnchor = CleanTitleAnchor(title)

                Dim rngBibliography As Range

                ' First, get the range of the bookmark "Zotero_Bibliography"
                Set rngBibliography = ActiveDocument.Bookmarks("Zotero_Bibliography").Range

                With rngBibliography.Find
                    .ClearFormatting
                    .Text = Left(title, 255)
                    .Forward = True
                    .Wrap = wdFindStop ' Stop when reaching the end of the range
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .Execute
                End With

                ' Check if the text was found
                If rngBibliography.Find.Found Then
                    ' Create a new range object to represent the found paragraph
                    Dim rngFound As Range
                    Set rngFound = rngBibliography.Paragraphs(1).Range
                    
                    ' Add a bookmark to the found range
                    ActiveDocument.Bookmarks.Add Range:=rngFound, Name:=titleAnchor
                Else
                    If MsgBox("Not found in bibliography:" & vbCrLf & title & vbCrLf & vbCrLf & _
                                "Do you want to continue with the next Zotero citation?", _
                                    vbYesNo + vbCritical, "Error") = vbNo Then
                        Exit Sub
                    Else
                        GoTo SkipToNextCitation
                    End If
                End If

                Dim rng As Range
                Set rng = aField.Result

                With rng.Find
                    .Text = citation ' Assuming 'citation' is the variable holding the text to find
                    .Forward = True
                    .Wrap = wdFindStop ' Ensures the search does not wrap around the document
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                    Dim found As Boolean
                    found = .Execute
                End With

                If Not found Then
                    MsgBox "Not found the citation " & citation, vbOKOnly + vbCritical, "Error"
                    Exit Sub
                Else
                    Dim hp As Hyperlink
                    Set hp = ActiveDocument.Hyperlinks.Add(Anchor:=rng, SubAddress:=titleAnchor)
                    hp.Range.style = ActiveDocument.Styles("交叉引用")
                End If

            SkipToNextCitation:
            Next i
        End If
    Next aField
    ' Restore the original selection
    ActiveDocument.Range(nStart, nEnd).Select
    ' Re-enable screen updating
    Application.ScreenUpdating = True
End Sub
