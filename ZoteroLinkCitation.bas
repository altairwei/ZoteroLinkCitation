Attribute VB_Name = "ZoteroLinkCitation"
' An MS Word macro that links author-date or number style citations to their bibliography entry.
' altair_wei@outlook.com
' https://github.com/altairwei/ZoteroLinkCitation

Option Explicit

'-------------------------------------------------------------------
' VBA JSON Parser
' https://medium.com/swlh/excel-vba-parse-json-easily-c2213f4d8e7a
'-------------------------------------------------------------------

Private p&, token, dic
Private Function ParseJSON(json$, Optional key$ = "obj") As Object
    p = 1
    token = Tokenize(json)
    Set dic = CreateObject("Scripting.Dictionary")
    If token(p) = "{" Then ParseObj key Else ParseArr key
    Set ParseJSON = dic
End Function

Private Function ParseObj(key$)
    Do: p = p + 1
        Select Case token(p)
            Case "]"
            Case "[":  ParseArr key
            Case "{"
                       If token(p + 1) = "}" Then
                           p = p + 1
                           dic.Add key, "null"
                       Else
                           ParseObj key
                       End If
                
            Case "}":  key = ReducePath(key): Exit Do
            Case ":":  key = key & "." & token(p - 1)
            Case ",":  key = ReducePath(key)
            Case Else: If token(p + 1) <> ":" Then dic.Add key, token(p)
        End Select
    Loop
End Function

Private Function ParseArr(key$)
    Dim e&
    Do: p = p + 1
        Select Case token(p)
            Case "}"
            Case "{":  ParseObj key & ArrayID(e)
            Case "[":  ParseArr key
            Case "]":  Exit Do
            Case ":":  key = key & ArrayID(e)
            Case ",":  e = e + 1
            Case Else: dic.Add key & ArrayID(e), token(p)
        End Select
    Loop
End Function

Private Function Tokenize(s$)
    Const Pattern = """(([^""\\]|\\.)*)""|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?|\w+|[^\s""']+?"
    Tokenize = RExtract(s, Pattern, True)
End Function

Private Function RExtract(s$, Pattern, Optional bGroup1Bias As Boolean, Optional bGlobal As Boolean = True)
  Dim c&, m, n
  Dim v()
  With CreateObject("vbscript.regexp")
    .Global = bGlobal
    .MultiLine = False
    .IgnoreCase = True
    .Pattern = Pattern
    If .TEST(s) Then
      Set m = .Execute(s)
      ReDim v(1 To m.Count)
      For Each n In m
        c = c + 1
        v(c) = n.value
        If bGroup1Bias Then If Len(n.submatches(0)) Or n.value = """""" Then v(c) = n.submatches(0)
      Next
    End If
  End With
  RExtract = v
End Function

Private Function ArrayID$(e)
    ArrayID = "(" & e & ")"
End Function

Private Function ReducePath$(key$)
    If InStr(key, ".") Then ReducePath = Left(key, InStrRev(key, ".") - 1) Else ReducePath = key
End Function

Function GetFilteredValues(dic, match)
    Dim c&, i&, v, w
    v = dic.keys
    ReDim w(1 To dic.Count)
    For i = 0 To UBound(v)
        If v(i) Like match Then
            c = c + 1
            w(c) = dic(v(i))
        End If
    Next
    ReDim Preserve w(1 To c)
    GetFilteredValues = w
End Function

Function GetFilteredTable(dic, cols)
    Dim c&, i&, j&, v, w, z
    v = dic.keys
    z = GetFilteredValues(dic, cols(0))
    ReDim w(1 To UBound(z), 1 To UBound(cols) + 1)
    For j = 1 To UBound(cols) + 1
         z = GetFilteredValues(dic, cols(j - 1))
         For i = 1 To UBound(z)
            w(i, j) = z(i)
         Next
    Next
    GetFilteredTable = w
End Function

'-------------------------------------------------------------------
' ZoteroLinkCitation Utilities
'-------------------------------------------------------------------

Private Function isSupportedStyle(style As String) As Boolean
    isSupportedStyle = style = "molecular-plant"
End Function

Private Sub QuickSort(arr As Variant, inLow As Long, inHigh As Long)
    Dim pivot As String
    Dim tmpSwap As Variant
    Dim low As Long
    Dim high As Long
    
    low = inLow
    high = inHigh
    pivot = arr((low + high) \ 2)
    
    While (low <= high)
        While (arr(low) < pivot And low < inHigh)
            low = low + 1
        Wend
        
        While (pivot < arr(high) And high > inLow)
            high = high - 1
        Wend
        
        If (low <= high) Then
            tmpSwap = arr(low)
            arr(low) = arr(high)
            arr(high) = tmpSwap
            low = low + 1
            high = high - 1
        End If
    Wend
    
    If (inLow < high) Then QuickSort arr, inLow, high
    If (low < inHigh) Then QuickSort arr, low, inHigh
End Sub

Private Function ExtractZoteroPrefData() As String
    Dim prop As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    For Each prop In ActiveDocument.CustomDocumentProperties
        If Left(prop.Name, 11) = "ZOTERO_PREF" Then
            dict(prop.Name) = prop.Value
        End If
    Next prop
    
    Dim sortedKeys As Variant
    sortedKeys = dict.Keys
    Call QuickSort(sortedKeys, LBound(sortedKeys), UBound(sortedKeys))

    Dim concatenatedValues As String
    Dim key As Variant
    For Each key In sortedKeys
        concatenatedValues = concatenatedValues & dict(key)
    Next key

    ExtractZoteroPrefData = concatenatedValues
End Function

Private Function GetZoteroPrefs() As Object
    Dim zoteroData As String
    zoteroData = ExtractZoteroPrefData()

    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")

    xmlDoc.Async = False
    xmlDoc.LoadXML zoteroData

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    If xmlDoc.ParseError.ErrorCode <> 0 Then
        MsgBox "XML Parse Error: " & xmlDoc.ParseError.Reason
        Set GetZoteroPrefs = dict
        Exit Function
    End If

    Dim dataElem As Object
    Set dataElem = xmlDoc.SelectSingleNode("//data")
    If Not dataElem Is Nothing Then
        dict("data-version") = dataElem.getAttribute("data-version")
        dict("zotero-version") = dataElem.getAttribute("zotero-version")
    End If
    
    Dim sessionElem As Object
    Set sessionElem = xmlDoc.SelectSingleNode("//session")
    If Not sessionElem Is Nothing Then
        dict("session-id") = sessionElem.getAttribute("id")
    End If

    Dim styleElem As Object
    Set styleElem = xmlDoc.SelectSingleNode("//style")
    If Not styleElem Is Nothing Then
        Dim segments() As String
        segments = Split(styleElem.getAttribute("id"), "/")
        dict("style-id") = segments(UBound(segments))
        dict("hasBibliography") = styleElem.getAttribute("hasBibliography")
        dict("bibliographyStyleHasBeenSet") = styleElem.getAttribute("bibliographyStyleHasBeenSet")
    End If

    Dim prefElem As Object
    Set prefElem = xmlDoc.SelectSingleNode("//prefs/pref[@name='fieldType']")
    If Not prefElem Is Nothing Then
        dict("pref-fieldType") = prefElem.getAttribute("value")
    End If

    Set GetZoteroPrefs = dict
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

Private Function ParseCSLCitationJson(ByVal code As String) As Object
    Dim jsonObj As Object
    Set jsonObj = ParseJSON(Trim(Replace(code, "ADDIN ZOTERO_ITEM CSL_CITATION", "")), "CSL")
    Set ParseCSLCitationJson = jsonObj
End Function

'-------------------------------------------------------------------
' ZoteroLinkCitation Macro
'-------------------------------------------------------------------

Public Sub ZoteroLinkCitation()
    ' Do not support Bookmark-type citations
    Dim prefs As Object
    Set prefs = GetZoteroPrefs()
    If Not prefs("pref-fieldType") = "Field" Then
        MsgBox "Only support 'Fields' type citations", vbCritical, "Error"
        Exit Sub
    End If

    Dim styleId As String
    styleId = prefs("style-id")
    If Not isSupportedStyle(styleId) Then
        MsgBox "The current citation style is not yet supported: " & styleId, vbCritical, "Error"
        Exit Sub
    End If

    Dim userTextStyle As String
    userTextStyle = InputBox(title := "Set a style for linked citations?", _
                             prompt := "If you want to set a certain style for linked citations," & _
                                        " enter the name of that style below.")

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
    Dim aField, iCount As Integer
    For Each aField In ActiveDocument.Fields
        ' Check if the field is a Zotero citation
        If aField.Type = wdFieldAddin Then
            If InStr(aField.Code, "ADDIN ZOTERO_ITEM") > 0 Then
                If MsgBox("Processed " & iCount & " citations, and found the next group:" & vbCrLf & vbCrLf & _ 
                            aField.Result.Text & vbCrLf & vbCrLf & "Do you want to continue?", _
                            vbYesNo + vbQuestion, "Continue?") = vbNo Then
                    Exit For
                End If

                Dim Citations() As String
                Dim Titles() As String
                Citations = ExtractCitations(aField.Result.Text)
                Titles = ExtractTitlesFromJSON(aField.Code)
                AssertArrayLengthsEqual Citations, Titles

                Dim i
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
                            GoTo ExitTheMacro
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
                        GoTo ExitTheMacro
                    Else
                        Dim hp As Hyperlink
                        Set hp = ActiveDocument.Hyperlinks.Add(Anchor:=rng, SubAddress:=titleAnchor)
                        If userTextStyle <> "" Then
                            hp.Range.style = ActiveDocument.Styles(userTextStyle)
                        End If
                    End If

                    iCount = iCount + 1

                SkipToNextCitation:
                Next i
            End If
        End If
    Next aField

ExitTheMacro:

    MsgBox "Linked " & iCount & " Zotero citations.", vbInformation, "Finish"

    ' Restore the original selection
    ActiveDocument.Range(nStart, nEnd).Select
    ' Re-enable screen updating
    Application.ScreenUpdating = True
End Sub
