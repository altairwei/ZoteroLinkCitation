Option Explicit

' =========================================================
' ZoteroLinkCitation - Mac-compatible edition
' Links Zotero in-text citations to corresponding bibliography entries
' Works without Windows-only CreateObject COM components.
'
' Notes / limitations:
' - Best-effort matching (uses DOI first, then title substring fallback)
' - Supports common author-date and numeric bracket/paren formats by linking
'   each semicolon-separated citation segment inside () or [].
' - If bibliography entries do not contain DOI or title, linking may fail.
' =========================================================

Private Type CitationSpan
    BibKey As String
    startPos As Long
    EndPos As Long
End Type

' ----------------------------
' Entry points
' ----------------------------
Public Sub ZoteroLinkCitationAll_Mac()
    Application.ScreenUpdating = False
    On Error GoTo CleanFail

    Call LinkAllZoteroCitationsInRange(ActiveDocument.Range)

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Macro failed: " & Err.Description, vbCritical, "ZoteroLinkCitation (Mac)"
End Sub

Public Sub ZoteroLinkCitationWithinSelection_Mac()
    If Selection Is Nothing Then Exit Sub
    If Selection.Range.Fields.Count = 0 Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo CleanFail

    Call LinkAllZoteroCitationsInRange(Selection.Range)

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Macro failed: " & Err.Description, vbCritical, "ZoteroLinkCitation (Mac)"
End Sub

' ----------------------------
' Core logic
' ----------------------------
Private Sub LinkAllZoteroCitationsInRange(ByVal scopeRng As Range)
    Dim bibField As field
    Set bibField = FindZoteroBibliographyField()

    If bibField Is Nothing Then
        Err.Raise vbObjectError + 513, , "Cannot find Zotero bibliography field. Ensure bibliography is still Zotero-linked (not unlinked/plain text)."
    End If

    Dim fld As field
    Dim linkedCount As Long
    linkedCount = 0

    For Each fld In scopeRng.Fields
        If IsZoteroCitationField(fld) Then
            linkedCount = linkedCount + LinkOneZoteroCitationField(fld, bibField)
        End If
    Next fld

    MsgBox "Linked " & linkedCount & " citation(s).", vbInformation, "ZoteroLinkCitation (Mac)"
End Sub

Private Function FindZoteroBibliographyField() As field
    Dim i As Long
    Dim codeText As String

    For i = ActiveDocument.Fields.Count To 1 Step -1
        If ActiveDocument.Fields(i).Type = wdFieldAddin Then
            codeText = ActiveDocument.Fields(i).code.text

            ' Mac Word/Zotero field code variants: be flexible
            If InStr(1, codeText, "ZOTERO_BIBL", vbTextCompare) > 0 _
            Or InStr(1, codeText, "ZOTERO_BIB", vbTextCompare) > 0 _
            Or InStr(1, codeText, "BIBLIOGRAPHY", vbTextCompare) > 0 Then
                Set FindZoteroBibliographyField = ActiveDocument.Fields(i)
                Exit Function
            End If
        End If
    Next i

    Set FindZoteroBibliographyField = Nothing
End Function

Private Function IsZoteroCitationField(ByVal fld As field) As Boolean
    If fld.Type <> wdFieldAddin Then
        IsZoteroCitationField = False
        Exit Function
    End If
    IsZoteroCitationField = (InStr(1, fld.code.text, "ADDIN ZOTERO_ITEM", vbTextCompare) > 0)
End Function

' Links one Zotero citation FIELD (which may contain multiple cites like (A 2020; B 2021))
' Returns how many individual citations were linked.
Private Function LinkOneZoteroCitationField(ByVal citeField As field, ByVal bibField As field) As Long
    Dim cits() As CitationSpan
    cits = ExtractCitationSpansFromFieldResult(citeField.result)

    Dim keys As Variant
    keys = ExtractKeysFromCSLJson(citeField.code.text) ' DOI first, then title fallback

    If VariantArrayLen(keys) = 0 Then
        LinkOneZoteroCitationField = 0
        Exit Function
    End If

    Dim n As Long
    n = MinLong(UBound(cits) - LBound(cits) + 1, VariantArrayLen(keys))

    If n <= 0 Then
        LinkOneZoteroCitationField = 0
        Exit Function
    End If

    Dim i As Long
    Dim tmpBookmark As String
    Dim bibBookmark As String
    Dim spanRng As Range
    Dim bibRng As Range
    Dim found As Boolean
    Dim needle As String

    For i = 0 To n - 1
        tmpBookmark = "ZLC_TMP_" & Format(i, "000") & "_" & CLng(Timer * 1000)

        Set spanRng = citeField.result.Document.Range(Start:=cits(i).startPos, End:=cits(i).EndPos)
        If spanRng.Start < spanRng.End Then
            ActiveDocument.Bookmarks.Add Name:=tmpBookmark, Range:=spanRng
        Else
            GoTo NextI
        End If

        Set bibRng = bibField.result.Duplicate

        needle = CStr(keys(i))
        needle = Trim$(needle)
        If Len(needle) = 0 Then GoTo NextI

        found = FindInRange(bibRng, Left$(needle, 255))

        If found Then
            bibBookmark = MakeBookmarkName(needle)

            Dim paraRng As Range
            Set paraRng = bibRng.Paragraphs(1).Range
            If paraRng.End > paraRng.Start Then paraRng.End = paraRng.End - 1

            bibBookmark = EnsureUniqueBookmarkName(bibBookmark)
            ActiveDocument.Bookmarks.Add Name:=bibBookmark, Range:=paraRng

            ActiveDocument.Hyperlinks.Add Anchor:=ActiveDocument.Bookmarks(tmpBookmark).Range, _
                                          SubAddress:=bibBookmark, _
                                          ScreenTip:=""

            LinkOneZoteroCitationField = LinkOneZoteroCitationField + 1
        End If

NextI:
        On Error Resume Next
        ActiveDocument.Bookmarks(tmpBookmark).Delete
        On Error GoTo 0
    Next i
End Function

' ----------------------------
' Variant array helpers
' ----------------------------
Private Function VariantArrayLen(ByVal v As Variant) As Long
    On Error GoTo Nope
    If Not IsArray(v) Then GoTo Nope
    If UBound(v) < LBound(v) Then
        VariantArrayLen = 0
    Else
        VariantArrayLen = UBound(v) - LBound(v) + 1
    End If
    Exit Function
Nope:
    VariantArrayLen = 0
End Function

Private Function AppendToVariantArray(ByVal v As Variant, ByVal s As String) As Variant
    Dim n As Long
    n = VariantArrayLen(v)

    Dim out As Variant
    If n = 0 Then
        out = Array(s)
    Else
        out = v
        ReDim Preserve out(0 To n)
        out(n) = s
    End If

    AppendToVariantArray = out
End Function

' ----------------------------
' Citation span extraction from displayed text (field.Result)
' ----------------------------
Private Function ExtractCitationSpansFromFieldResult(ByVal resRng As Range) As CitationSpan()
    Dim txt As String
    txt = resRng.text

    Dim spans() As CitationSpan
    ReDim spans(0)

    Dim i As Long
    Dim inBlock As Boolean
    Dim blockStart As Long

    inBlock = False
    blockStart = 0

    For i = 1 To Len(txt)
        Dim ch As String
        ch = Mid$(txt, i, 1)

        If (ch = "(" Or ch = "[") And Not inBlock Then
            inBlock = True
            blockStart = i + 1
        ElseIf (ch = ")" Or ch = "]") And inBlock Then
            Call AddDelimitedSpans(txt, blockStart, i - 1, resRng.Start, spans)
            inBlock = False
        End If
    Next i

    ' Fallback: treat whole range as a single block
    If UBound(spans) = 0 And spans(0).startPos = 0 And spans(0).EndPos = 0 Then
        Call AddDelimitedSpans(txt, 1, Len(txt), resRng.Start, spans)
    End If

    ExtractCitationSpansFromFieldResult = spans
End Function

Private Sub AddDelimitedSpans(ByVal fullText As String, ByVal s As Long, ByVal e As Long, _
                             ByVal baseDocStart As Long, ByRef spans() As CitationSpan)
    Dim block As String
    block = Mid$(fullText, s, e - s + 1)

    Dim parts() As String
    parts = Split(block, ";")

    Dim p As Long
    Dim runningOffset As Long
    runningOffset = 0

    For p = LBound(parts) To UBound(parts)
        Dim rawPart As String
        rawPart = parts(p)

        Dim part As String
        part = Trim$(rawPart)
        If Len(part) = 0 Then GoTo NextP

        Dim localStart As Long
        localStart = InStr(1 + runningOffset, block, rawPart, vbBinaryCompare)
        If localStart = 0 Then
            localStart = InStr(1 + runningOffset, block, part, vbBinaryCompare)
            If localStart = 0 Then GoTo NextP
        End If

        Dim localEnd As Long
        localEnd = localStart + Len(rawPart) - 1

        Dim docStart As Long, docEnd As Long
        docStart = baseDocStart + (s - 1) + localStart - 1
        docEnd = baseDocStart + (s - 1) + localEnd

        Call AppendSpan(spans, docStart, docEnd)

        runningOffset = localEnd

NextP:
    Next p
End Sub

Private Sub AppendSpan(ByRef spans() As CitationSpan, ByVal docStart As Long, ByVal docEnd As Long)
    Dim idx As Long
    If UBound(spans) = 0 And spans(0).startPos = 0 And spans(0).EndPos = 0 Then
        idx = 0
    Else
        idx = UBound(spans) + 1
        ReDim Preserve spans(idx)
    End If
    spans(idx).startPos = docStart
    spans(idx).EndPos = docEnd
    spans(idx).BibKey = ""
End Sub

' ----------------------------
' CSL JSON key extraction (no regex, no dictionary)
' Strategy: extract DOI if available, else title.
' Returns Variant array of strings aligned with citationItems order.
' ----------------------------
Private Function ExtractKeysFromCSLJson(ByVal fieldCode As String) As Variant
    Dim json As String
    json = fieldCode

    json = Replace(json, "ADDIN ZOTERO_ITEM CSL_CITATION", "", 1, -1, vbTextCompare)
    json = Trim$(json)

    Dim dois As Variant
    dois = ExtractJsonStringValues(json, """DOI""")

    Dim titles As Variant
    titles = ExtractJsonStringValues(json, """title""")

    ' Prefer DOI when present, otherwise fallback to title for that index.
    Dim out As Variant
    out = Array()

    Dim n As Long
    n = MaxLong(VariantArrayLen(dois), VariantArrayLen(titles))
    If n = 0 Then
        ExtractKeysFromCSLJson = out
        Exit Function
    End If

    Dim i As Long
    For i = 0 To n - 1
        Dim key As String
        key = ""

        If VariantArrayLen(dois) > i Then
            key = Trim$(CStr(dois(i)))
        End If

        If Len(key) = 0 And VariantArrayLen(titles) > i Then
            key = Trim$(CStr(titles(i)))
        End If

        key = StripBasicHtml(JsonUnescape(key))
        key = Trim$(key)

        If Len(key) > 0 Then
            out = AppendToVariantArray(out, key)
        Else
            ' Keep alignment by adding empty string
            out = AppendToVariantArray(out, "")
        End If
    Next i

    ExtractKeysFromCSLJson = out
End Function

' Extract values for a JSON key like """title""" or """DOI""" by scanning for: "key": "value"
Private Function ExtractJsonStringValues(ByVal json As String, ByVal keyToken As String) As Variant
    Dim out As Variant
    out = Array()

    Dim pos As Long
    pos = 1

    Do
        pos = InStr(pos, json, keyToken, vbTextCompare)
        If pos = 0 Then Exit Do

        Dim colonPos As Long
        colonPos = InStr(pos, json, ":", vbBinaryCompare)
        If colonPos = 0 Then Exit Do

        Dim q1 As Long
        q1 = InStr(colonPos + 1, json, """", vbBinaryCompare)
        If q1 = 0 Then Exit Do

        Dim q2 As Long
        q2 = FindClosingJsonQuote(json, q1 + 1)
        If q2 = 0 Then Exit Do

        Dim raw As String
        raw = Mid$(json, q1 + 1, q2 - q1 - 1)

        out = AppendToVariantArray(out, raw)

        pos = q2 + 1
    Loop

    ExtractJsonStringValues = out
End Function

Private Function FindClosingJsonQuote(ByVal s As String, ByVal startPos As Long) As Long
    Dim i As Long
    Dim escaped As Boolean
    escaped = False

    For i = startPos To Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)

        If escaped Then
            escaped = False
        Else
            If ch = "\" Then
                escaped = True
            ElseIf ch = """" Then
                FindClosingJsonQuote = i
                Exit Function
            End If
        End If
    Next i

    FindClosingJsonQuote = 0
End Function

Private Function JsonUnescape(ByVal s As String) As String
    s = Replace(s, "\""", """")
    s = Replace(s, "\\", "\")
    s = Replace(s, "\/", "/")
    s = Replace(s, "\n", vbLf)
    s = Replace(s, "\r", vbCr)
    s = Replace(s, "\t", vbTab)
    JsonUnescape = s
End Function

Private Function StripBasicHtml(ByVal s As String) As String
    s = Replace(s, "<i>", "", 1, -1, vbTextCompare)
    s = Replace(s, "</i>", "", 1, -1, vbTextCompare)
    s = Replace(s, "<sub>", "", 1, -1, vbTextCompare)
    s = Replace(s, "</sub>", "", 1, -1, vbTextCompare)
    s = Replace(s, "<sup>", "", 1, -1, vbTextCompare)
    s = Replace(s, "</sup>", "", 1, -1, vbTextCompare)
    StripBasicHtml = s
End Function

' ----------------------------
' Word Find helpers (Mac-safe)
' ----------------------------
Private Function FindInRange(ByRef rng As Range, ByVal needle As String) As Boolean
    If Len(needle) = 0 Then
        FindInRange = False
        Exit Function
    End If

    With rng.Find
        .ClearFormatting
        .text = needle
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        FindInRange = .Execute
    End With
End Function

' ----------------------------
' Bookmark name utilities
' ----------------------------
Private Function MakeBookmarkName(ByVal s As String) As String
    Dim result As String
    Dim i As Long

    result = Replace(s, " ", "_")

    For i = 1 To Len(result)
        Dim ch As String
        ch = Mid$(result, i, 1)
        If Not (ch Like "[A-Za-z0-9_]") Then
            Mid$(result, i, 1) = "_"
        End If
    Next i

    If Left$(result, 1) Like "[0-9]" Then result = "_" & result

    If Len(result) > 36 Then
        result = Left$(result, 36) & "_" & Right$("000" & CStr(SimpleHash(s)), 3)
    End If

    MakeBookmarkName = result
End Function

Private Function EnsureUniqueBookmarkName(ByVal baseName As String) As String
    Dim nameTry As String
    nameTry = baseName

    Dim k As Long
    k = 1
    Do While ActiveDocument.Bookmarks.Exists(nameTry)
        nameTry = Left$(baseName, 30) & "_" & k
        k = k + 1
        If k > 9999 Then Exit Do
    Loop

    EnsureUniqueBookmarkName = nameTry
End Function

Private Function SimpleHash(ByVal s As String) As Long
    Dim i As Long, h As Long
    h = 0
    For i = 1 To Len(s)
        h = h + (Asc(Mid$(s, i, 1)) * i)
    Next i
    SimpleHash = (h Mod 1000)
End Function

' ----------------------------
' Misc
' ----------------------------
Private Function MinLong(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MinLong = a Else MinLong = b
End Function

Private Function MaxLong(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxLong = a Else MaxLong = b
End Function

' ----------------------------
' Debug helper (optional)
' Lists all ADDIN fields so you can see what Zotero inserted on your Mac.
' ----------------------------
Public Sub Debug_ListAddinFields()
    Dim i As Long, fld As field
    Dim out As String
    out = ""

    For i = 1 To ActiveDocument.Fields.Count
        Set fld = ActiveDocument.Fields(i)
        If fld.Type = wdFieldAddin Then
            out = out & i & ": " & Left$(Replace(fld.code.text, vbCr, " "), 160) & vbCrLf & vbCrLf
        End If
    Next i

    If Len(out) = 0 Then
        MsgBox "No wdFieldAddin fields found.", vbInformation
    Else
        MsgBox out, vbInformation, "ADDIN fields (preview)"
    End If
End Sub

