Attribute VB_Name = "ZoteroLinkCitation"
' An MS Word macro that links author-date or number style citations to their bibliography entry.
' altair_wei@outlook.com
' https://github.com/altairwei/ZoteroLinkCitation

' Option Explicit

'-------------------------------------------------------------------
' VBA JSON Parser
' https://medium.com/swlh/excel-vba-parse-json-easily-c2213f4d8e7a
'-------------------------------------------------------------------

Private p&, myTokens, dic
Private Function ParseJSON(json$, Optional key$ = "obj") As Object
    p = 1
    myTokens = Tokenize(json)
    Set dic = CreateObject("Scripting.Dictionary")
    If myTokens(p) = "{" Then ParseObj key Else ParseArr key
    Set ParseJSON = dic
End Function

Private Function ParseObj(key$)
    Do: p = p + 1
        Select Case myTokens(p)
            Case "]"
            Case "[":  ParseArr key
            Case "{"
                       If myTokens(p + 1) = "}" Then
                           p = p + 1
                           dic.Add key, "null"
                       Else
                           ParseObj key
                       End If
                
            Case "}":  key = ReducePath(key): Exit Do
            Case ":":  key = key & "." & myTokens(p - 1)
            Case ",":  key = ReducePath(key)
            Case Else: If myTokens(p + 1) <> ":" Then dic.Add key, myTokens(p)
        End Select
    Loop
End Function

Private Function ParseArr(key$)
    Dim e&
    Do: p = p + 1
        Select Case myTokens(p)
            Case "}"
            Case "{":  ParseObj key & ArrayID(e)
            Case "[":  ParseArr key
            Case "]":  Exit Do
            Case ":":  key = key & ArrayID(e)
            Case ",":  e = e + 1
            Case Else: dic.Add key & ArrayID(e), myTokens(p)
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

'-------------------------------------------------------------------
' ZoteroLinkCitation Utilities
'-------------------------------------------------------------------

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

'-------------------------------------------------------------------
' ZoteroLinkCitation Macro
'-------------------------------------------------------------------

Private Function ParseCSLCitationJson(ByVal code As String) As Object
    Dim jsonObj As Object
    Set jsonObj = ParseJSON(Trim(Replace(code, "ADDIN ZOTERO_ITEM CSL_CITATION", "")), "CSL")
    Set ParseCSLCitationJson = jsonObj
End Function

Public Sub ZoteroLinkCitation()
    Dim styleId As String
    styleId = ExtractZoteroStyleId()

    If Not isSupportedStyle() Then
        MsgBox "The current citation style is not yet supported: " & styleId, vbCritical, "Error"
        Exit Sub
    End If

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
    Dim aField
    For Each aField In ActiveDocument.Fields
        ' Check if the field is a Zotero citation
        If aField.Type = wdFieldAddin And InStr(aField.Code, "ADDIN ZOTERO_ITEM") > 0 Then
            If MsgBox("Found " & aField.Result.Text, vbYesNo + vbQuestion, "Continue?") = vbNo Then
                Exit For
            End If

            Dim obj
            Set obj = ParseCSLCitationJson(aField.Code)
            MsgBox obj("CSL.citationItems(0).itemData.title")

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
                    ' hp.Range.style = ActiveDocument.Styles("")
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
