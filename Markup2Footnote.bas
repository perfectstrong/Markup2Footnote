Attribute VB_Name = "Markup2Footnote"
Sub Markup2Footnote()
    Dim ct As Range
    Set ct = ActiveDocument.Content
    Dim TextToFootnote As String, OpenMark As String, CloseMark As String
    Dim CountFootnote As Long
    CountFootnote = 0

    With ct.Find
        .ClearFormatting
        .Wrap = wdFindContinue
        .Forward = True
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
        .Text = "\(\(*\)\)"
    End With
    

    Do
        ct.Find.Execute
        If ct.Find.Found = True Then
            ' Extract text
            TextToFootnote = Replace(ct.Text, "((", "")
            TextToFootnote = Replace(TextToFootnote, "))", "")
            CountFootnote = CountFootnote + 1
            ' Clear found text
            ct.Text = ""
            ct.Collapse Direction:=wdCollapseEnd
            ' Add footnotes
            ActiveDocument.Footnotes.Add Range:=ct, Text:=TextToFootnote
        End If
    Loop While ct.Find.Found
    MsgBox "Number of footnotes converted: " & CountFootnote

End Sub
