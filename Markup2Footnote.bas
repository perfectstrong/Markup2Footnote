Attribute VB_Name = "Markup2Footnote"
Sub Markup2Footnote()
    Application.ScreenUpdating = False

    Dim ct As Range
    Set ct = ActiveDocument.Content
    Dim TextToFootnote As String
    Dim CountFootnote As Long
    CountFootnote = 0

    With ct.Find
        .ClearFormatting
        .Wrap = wdFindStop
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
    

        Do While (ct.Find.Execute = True)
            ' Rip off markup
            ActiveDocument.Range(Start:=ct.Start, End:=ct.Start + 2).Delete
            ActiveDocument.Range(Start:=ct.End - 2, End:=ct.End).Delete
            ct.Select
            ' Cut selection to preserve formatting
            Selection.Cut
            ' Delete old text
            ct.Text = ""
            CountFootnote = CountFootnote + 1
            
            ' Add footnotes
            ActiveDocument.Footnotes.Add Range:=ct
            Selection.PasteAndFormat Type:=wdFormatOriginalFormatting
    Loop
    
    Application.ScreenUpdating = True
    MsgBox "Number of footnotes converted: " & CountFootnote
End Sub
