Attribute VB_Name = "ZoteroLinkCitation"
' An MS Word macro that links author-date or number style citations to their bibliography entry.
' altair_wei@outlook.com
' https://github.com/altairwei/ZoteroLinkCitation

Option Explicit

Type Citation
    BibPattern As String
    Start As Long
    End As Long
End Type

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
            Case "[":  ParseArr key & ArrayID(e)
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
  With CreateObject("VBScript.RegExp")
    .Global = bGlobal
    .MultiLine = False
    .IgnoreCase = True
    .Pattern = Pattern
    If .TEST(s) Then
      Set m = .Execute(s)
      ReDim v(1 To m.count)
      For Each n In m
        c = c + 1
        v(c) = n.Value
        If bGroup1Bias Then If Len(n.submatches(0)) Or n.Value = """""" Then v(c) = n.submatches(0)
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
    v = dic.Keys
    ReDim w(1 To dic.count)
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
    v = dic.Keys
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

Private Function GetZoteroPrefsFromXml(zoteroData As String) As Object
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

    Set GetZoteroPrefsFromXml = dict
End Function

Private Function GetZoteroPrefsFromJson(zoteroData As String) As Object
    Dim jsonObj As Object
    Set jsonObj = ParseJSON(Trim(zoteroData), "prefs")

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    dict("data-version") = jsonObj("prefs.dataVersion")
    dict("zotero-version") = jsonObj("prefs.zoteroVersion")
    dict("session-id") = jsonObj("prefs.sessionID")

    Dim segments() As String
    segments = Split(jsonObj("prefs.style.styleID"), "/")
    dict("style-id") = segments(UBound(segments))
    dict("hasBibliography") = jsonObj("prefs.style.hasBibliography")
    dict("bibliographyStyleHasBeenSet") = jsonObj("prefs.style.bibliographyStyleHasBeenSet")

    dict("pref-fieldType") = jsonObj("prefs.prefs.fieldType")

    Set GetZoteroPrefsFromJson = dict
End Function

Private Function GetZoteroPrefs() As Object
    Dim zoteroData As String
    zoteroData = ExtractZoteroPrefData()

    Dim firstChar As String
    firstChar = Left(Trim(zoteroData), 1)

    Select Case firstChar
        Case "<"
            Set GetZoteroPrefs = GetZoteroPrefsFromXml(zoteroData)
        Case "{", "["
            Set GetZoteroPrefs = GetZoteroPrefsFromJson(zoteroData)
        Case Else
            Err.Raise vbObjectError + 1, "GetZoteroPrefs", "Can not find Zotero Preference in this document."
    End Select

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

Function SimpleHash(ByVal inputString As String) As String
    Dim i As Long
    Dim hashValue As Long

    For i = 1 To Len(inputString)
        hashValue = hashValue + (Asc(Mid(inputString, i, 1)) * i)
    Next i

    Dim modValue As Long
    modValue = 100
    hashValue = hashValue Mod modValue

    If hashValue < 0 Then
        hashValue = hashValue + modValue
    End If

    SimpleHash = Format$(hashValue, "000")
End Function

Private Function ConvertToBookmarkName(ByVal str As String) As String
    Dim result As String
    Dim i As Integer

    ' Replace illegal characters
    result = Replace(str, " ", "_")
    For i = 1 To Len(result)
        ' Check each character and replace if not alphanumeric or underscore
        If Not (Mid(result, i, 1) Like "[A-Za-z0-9_]") Then
            Mid(result, i, 1) = "_"
        End If
    Next i

    ' Avoid starting with a digit
    If Left(result, 1) Like "[0-9]" Then
        result = "_" & result
    End If

    ' Limit the length to 40 characters
    If Len(result) > 40 Then
        result = Left(result, 36)
        result = result & "_" & SimpleHash(str)
    End If

    ConvertToBookmarkName = result
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

Function StyleExists(ByVal styleToTest As String, ByVal docToTest As Word.Document) As Boolean
    Dim testStyle As Word.style
    On Error Resume Next
    Set testStyle = docToTest.Styles(styleToTest)
    StyleExists = Not testStyle Is Nothing
End Function

'-------------------------------------------------------------------
' Citation Style Handler
'-------------------------------------------------------------------

' Such as (Dweba et al., 2017; Hu et al., 2022; Moonjely et al., 2023)
Private Sub ExtractAuthorYearCitations(field As field, ByRef citations() As Citation, _
        Optional onlyYear As Boolean = False, Optional multiRefCommaSep As Boolean = True)
    Dim targetRange As Range, charRange As Range
    Set targetRange = field.result
    Set charRange = targetRange.Duplicate
    charRange.Collapse wdCollapseStart

    ReDim citations(0)
    Dim rangeIndex As Long
    rangeIndex = -1

    Dim inCitation As Boolean, nComma As Integer, beginYear As Boolean
    inCitation = False
    nComma = 0
    beginYear = False

    Dim json As Object
    Set json = ParseCSLCitationJson(field.code)

    Dim startChar As Long, endChar As Long

    Dim i As Long
    For i = 1 To targetRange.Characters.count
        charRange.Start = targetRange.Start + i - 1
        charRange.End = targetRange.Start + i

        ' Start of full author-year citation
        If charRange.Text = "(" And Not onlyYear Then
            inCitation = True
            startChar = charRange.Start + 1

        ' Start of year citation
        ElseIf charRange.Text Like "[0-9]" Then
            beginYear = True

            If onlyYear And Not inCitation Then
                inCitation = True
                startChar = charRange.Start
            End If

        ' Check multiple citations of same author
        ElseIf multiRefCommaSep And charRange.Text = "," Then
            nComma = nComma + 1
            If nComma > 1 And beginYear Then
                GoTo CreateCitationObject
            End If

        ' End of citation
        ElseIf charRange.Text = ";" Or charRange.Text = ")" Then
            beginYear = False
            If multiRefCommaSep Then nComma = 0

CreateCitationObject:
            If inCitation Then
                endChar = charRange.Start

                rangeIndex = rangeIndex + 1
                If rangeIndex > UBound(citations) Then
                    ReDim Preserve citations(0 To rangeIndex)
                End If

                citations(rangeIndex).Start = startChar
                citations(rangeIndex).End = endChar
                citations(rangeIndex).BibPattern = RemoveHtmlTags( _
                    json("CSL.citationItems(" & rangeIndex & ").itemData.title"))

                inCitation = False
            End If

            ' Skip space after delimiter
            If (charRange.Text = ";" Or charRange.Text = ",") And Not onlyYear Then
                i = i + 1
                startChar = endChar + 2
                inCitation = True
            End If

        End If
    Next i

    ' Resize the array to fit the number of found ranges
    ReDim Preserve citations(0 To rangeIndex)
End Sub

' Such as [1], [2], [3] etc.
Private Sub ExtractNumberInBrackets(field As field, ByRef citations() As Citation, Optional bracket As String = "[]")
    Dim targetRange As Range, charRange As Range
    Set targetRange = field.result
    Set charRange = targetRange.Duplicate
    charRange.Collapse wdCollapseStart

    ReDim citations(0)
    Dim rangeIndex As Long
    rangeIndex = -1

    Dim startBracket As String, endBracket As String
    startBracket = Left(bracket, 1)
    endBracket = Right(bracket, 1)

    Dim inBrackets As Boolean
    inBrackets = False

    Dim json As Object
    Set json = ParseCSLCitationJson(field.code)

    Dim startChar As Long, endChar As Long

    Dim i As Long
    For i = 1 To targetRange.Characters.count
        charRange.Start = targetRange.Start + i - 1
        charRange.End = targetRange.Start + i

        If charRange.Text = startBracket Then
            inBrackets = True
            startChar = charRange.Start + 1 ' Start after the bracket
        ElseIf charRange.Text = endBracket And inBrackets Then
            If startChar < endChar Then
                rangeIndex = rangeIndex + 1
                If rangeIndex > UBound(citations) Then
                    ReDim Preserve citations(0 To rangeIndex)
                End If
                With citations(rangeIndex)
                    .Start = startChar
                    .End = endChar
                    .BibPattern = RemoveHtmlTags( _
                        json("CSL.citationItems(" & rangeIndex & ").itemData.title"))
                End With
            End If
            inBrackets = False
        ElseIf inBrackets And IsNumeric(charRange.Text) Then
            endChar = charRange.End ' Update end if still in brackets and character is numeric
        End If
    Next i

    ' Resize the array to fit the number of found ranges
    ReDim Preserve citations(0 To rangeIndex)

End Sub

' Such as [47,98,100鈥?02]
Private Sub ExtractSerialNumberCitations(field As field, ByRef citations() As Citation, Optional border = "")
    Dim targetRange As Range, charRange As Range
    Set targetRange = field.result
    Set charRange = targetRange.Duplicate
    charRange.Collapse wdCollapseStart

    ReDim citations(0)
    Dim rangeIndex As Long, citOrder As Long
    rangeIndex = -1
    citOrder = -1

    Dim startBorder As String, endBorder As String
    startBorder = Left(border, 1)
    endBorder = Right(border, 1)

    Dim inCitation As Boolean
    inCitation = False

    Dim lastNum As Long
    lastNum = 0

    Dim json As Object
    Set json = ParseCSLCitationJson(field.code)

    Dim startChar As Long, endChar As Long

    Dim currentChar As String
    Dim citationText As String

    Dim i As Long, RL As Long
    RL = targetRange.Characters.count

    ' Add a pseudo-border to the citation text without borders
    If Len(endBorder) = 0 Then
        RL = RL + 1
        endBorder = "]"
    End If

    For i = 1 To RL
        charRange.Start = targetRange.Start + i - 1
        charRange.End = targetRange.Start + i

        If i <= targetRange.Characters.count Then
            currentChar = charRange.Text
        Else
            ' Point to the psuedo-border
            currentChar = endBorder
        End If

        If currentChar Like "[0-9]" And Not inCitation Then
            inCitation = True
            startChar = charRange.Start
            citationText = currentChar

        ' ChrW(8211) means the character "en dash"
        ElseIf currentChar = "," Or currentChar = endBorder Or currentChar = ChrW(8211) Then

            If currentChar = ChrW(8211) Then
                lastNum = CLng(citationText)
            End If

            If inCitation Then
                endChar = charRange.Start

                rangeIndex = rangeIndex + 1
                If rangeIndex > UBound(citations) Then
                    ReDim Preserve citations(0 To rangeIndex)
                End If

                If (currentChar = "," Or currentChar = endBorder) And lastNum > 0 Then
                    citOrder = citOrder + CLng(citationText) - lastNum
                Else
                    citOrder = citOrder + 1
                End If

                citations(rangeIndex).Start = startChar
                citations(rangeIndex).End = endChar
                citations(rangeIndex).BibPattern = RemoveHtmlTags( _
                    json("CSL.citationItems(" & citOrder & ").itemData.title"))

                If Len(citations(rangeIndex).BibPattern) = 0 Then
                    Err.Raise vbObjectError + 1, "ExtractCitations", "Can not find citation CSL data"
                End If

                inCitation = False
            End If

            If currentChar = "," Or currentChar = endBorder Then
                lastNum = 0
            End If

        ElseIf inCitation Then
            citationText = citationText & currentChar

        End If

    Next i

    If rangeIndex >= 0 Then
        ReDim Preserve citations(0 To rangeIndex)
    Else
        ' 没有找到任何引用，初始化为空数组
        ReDim citations(-1 To -1)
    End If

End Sub

'-------------------------------------------------------------------
' Supported Citation Styles
'-------------------------------------------------------------------

Private Function isSupportedStyle(ByVal style As String) As Boolean
    Dim predefinedList As String
    predefinedList = "|" & _
        "molecular-plant|ieee|apa|vancouver|american-chemical-society|" & _
        "american-medical-association|nature|american-political-science-association|" & _
        "american-sociological-association|chicago-author-date|bmc-medicine|" & _
        "china-national-standard-gb-t-7714-2015-numeric|" & _
        "china-national-standard-gb-t-7714-2015-author-date|" & _
        "harvard-cite-them-right|elsevier-harvard|modern-language-association|" & _
        "multidisciplinary-digital-publishing-institute|"
    style = "|" & style & "|"
    isSupportedStyle = InStr(1, predefinedList, style, vbTextCompare) > 0
End Function

Private Sub ExtractCitations(field As field, ByRef citations() As Citation, style As String)
    Select Case style
        Case "molecular-plant", "chicago-author-date", "modern-language-association"
            Call ExtractAuthorYearCitations(field, citations, onlyYear:=False, multiRefCommaSep:=False)

        Case "apa", "china-national-standard-gb-t-7714-2015-author-date", _
             "american-political-science-association", "american-sociological-association", _
             "harvard-cite-them-right"
            Call ExtractAuthorYearCitations(field, citations, onlyYear:=True, multiRefCommaSep:=True)

        Case "elsevier-harvard"
            Call ExtractAuthorYearCitations(field, citations, onlyYear:=False, multiRefCommaSep:=True)

        Case "ieee"
            Call ExtractNumberInBrackets(field, citations, "[]")

        Case "vancouver"
            Call ExtractSerialNumberCitations(field, citations, "()")

        Case "china-national-standard-gb-t-7714-2015-numeric", "bmc-medicine", "multidisciplinary-digital-publishing-institute"
            Call ExtractSerialNumberCitations(field, citations, "[]")

        Case "american-chemical-society", "american-medical-association", "nature"
            Call ExtractSerialNumberCitations(field, citations, "")

        Case Else
            Err.Raise vbObjectError + 1, "ExtractCitations", "Citation style not recognized"
    End Select
End Sub

'-------------------------------------------------------------------
' ZoteroLinkCitation Macro
'-------------------------------------------------------------------

Public Sub ZoteroLinkCitationWithinSelection()
    If Selection.Fields.count > 0 Then
        Dim originalRng As Range
        Set originalRng = Selection.Range

        Application.ScreenUpdating = False

        Dim targetFields As New Collection
        Dim fld As field

        For Each fld In Selection.Fields
            targetFields.Add fld
        Next fld

        Call ZoteroLinkCitation(targetFields, False, False)

        ' Restore the original selection
        ActiveWindow.ScrollIntoView originalRng, True
        originalRng.Select

        Application.ScreenUpdating = True
    End If
End Sub

Public Sub ZoteroLinkCitationAll()
    Dim originalRng As Range
    Set originalRng = Selection.Range

    Dim debugging As Boolean
    debugging = (MsgBox("Do you want run in debug mode?", vbYesNo + vbQuestion, "Debug?") = vbYes)

    ' Disable screen updating for performance
    Application.ScreenUpdating = False

    Call ZoteroLinkCitation(ActiveDocument.Fields, debugging)

    ' Restore the original selection
    ActiveWindow.ScrollIntoView originalRng, True
    originalRng.Select

    ' Re-enable screen updating
    Application.ScreenUpdating = True
    Exit Sub
End Sub

Private Sub ZoteroLinkCitation(targetFields, Optional debugging As Boolean = False, Optional notify As Boolean = True)
    ' Do not support Bookmark-type citations
    Dim prefs As Object
    Set prefs = GetZoteroPrefs()
    If Not prefs("pref-fieldType") = "Field" Then
        MsgBox "Only support 'Fields' type citations" & vbCrLf & vbCrLf & _
            "Click on Document Preferences in Word, then expand the Advanced Options section. " & _
            "This will allow you to verify whether the 'Store citation as bookmarks' option was accidentally enabled.", _
            vbCritical, "Error"
        Exit Sub
    End If

    Dim styleId As String
    styleId = prefs("style-id")
    If Not isSupportedStyle(styleId) Then
        MsgBox "The current citation style is not yet supported: " & styleId, vbCritical, "Error"
        Exit Sub
    End If

    Dim userTextStyle As String

    If notify Then
        Dim resp As String
        resp = InputBox(title:="Set an MS Word style for hyperlinks?", _
                        prompt:="If you want to set a certain style for hyperlinks," & _
                                    " enter the name of that style below.")
        If StyleExists(resp, ActiveDocument) Then userTextStyle = resp
    End If

    Dim i As Long
    Dim bibField As field
    Set bibField = Nothing

    ' Find the Zotero bibliography field
    For i = ActiveDocument.Fields.count To 1 Step -1
        If ActiveDocument.Fields(i).Type = wdFieldAddin Then
            If InStr(ActiveDocument.Fields(i).code, "ADDIN ZOTERO_BIBL") > 0 Then
                Set bibField = ActiveDocument.Fields(i)
                Exit For
            End If
        End If
    Next i

    If bibField Is Nothing Then
        Err.Raise vbObjectError + 513, , "Can not find Zotero bibliography field."
    End If

    ' Iterate through all fields in the document
    Dim aField As field, iCount As Integer
    For Each aField In targetFields
        ' Check if the field is a Zotero citation
        If aField.Type = wdFieldAddin Then
            If InStr(aField.code, "ADDIN ZOTERO_ITEM") > 0 Then

                If debugging Then
                    ' Focus to next field
                    Application.ScreenUpdating = True
                    ActiveWindow.ScrollIntoView aField.result, True
                    aField.result.Select

                    ' Update the document
                    DoEvents

                    If MsgBox("Processed " & iCount & " citations, and found the next group:" & vbCrLf & vbCrLf & _
                                aField.result.Text & vbCrLf & vbCrLf & "Do you want to continue?", _
                                vbYesNo + vbQuestion, "Continue?") = vbNo Then
                        Exit For
                    End If

                    Application.ScreenUpdating = False
                End If

                Dim cit As Citation, cits() As Citation
                Call ExtractCitations(aField, cits, styleId)

                ' Locate all citations in the field
                Dim tempBookmarkName As String
                For i = 0 To UBound(cits)
                    cit = cits(i)
                    Dim rng As Range
                    Set rng = aField.result.Document.Range(Start:=cit.Start, End:=cit.End)
                    tempBookmarkName = "ZoteroLinkCitationTempBookmark" & i
                    ActiveDocument.Bookmarks.Add Name:=tempBookmarkName, Range:=rng
                Next i

                ' Link citations to bibliography
                For i = 0 To UBound(cits)
                    cit = cits(i)

                    Dim title As String
                    title = cit.BibPattern

                    ' Create a sanitized anchor name from the title
                    Dim titleAnchor As String
                    titleAnchor = ConvertToBookmarkName(title)

                    ' Get the range of Zotero bibliography
                    Dim rngBibliography As Range
                    Set rngBibliography = bibField.result

                    With rngBibliography.Find
                        .ClearFormatting
                        .Text = Left(title, 255)
                        .Forward = True
                        ' .MatchPhrase = True
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
                        ' Ensure that the Range does not extend to the end of the bibliography field
                        rngFound.End = rngFound.End - 1
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

                    ' Create hyperlink according to temporary bookmark
                    Dim hp As Hyperlink
                    Set hp = ActiveDocument.Hyperlinks.Add( _
                        Anchor:=ActiveDocument.Bookmarks("ZoteroLinkCitationTempBookmark" & i).Range, _
                        SubAddress:=titleAnchor, ScreenTip:="")

                    ' Apply text style to the hyperlink
                    If userTextStyle <> "" Then
                        hp.Range.style = ActiveDocument.Styles(userTextStyle)
                    End If

                    iCount = iCount + 1

SkipToNextCitation:
                    ActiveDocument.Bookmarks("ZoteroLinkCitationTempBookmark" & i).Delete

                Next i

            End If
        End If
    Next aField

ExitTheMacro:

    If notify Then MsgBox "Linked " & iCount & " Zotero citations.", vbInformation, "Finish"

End Sub


