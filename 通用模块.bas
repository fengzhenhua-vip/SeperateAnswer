Attribute VB_Name = "通用模块"
Sub 删除答案通用版()
    Dim TempPar As Paragraph
    Dim RemoveOn As Boolean
    For n = 1 To ActiveDocument.InlineShapes.Count                      '保留以图片作为标题的情况
        If ActiveDocument.InlineShapes(n).Width > 300 Then
            ActiveDocument.InlineShapes(n).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.TypeText Text:="【【"
        End If
    Next
    For Each TempPar In ActiveDocument.Paragraphs
        If InStr(Mid(TempPar.Range, 1, 4), "【") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "[") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "考点") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "A组") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "B组") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "C组") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "D组") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "例") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "一 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "一、") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "二 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "二、") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "三 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "三、") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "四 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "四、") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "五 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "五、") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "六 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "六、") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "七 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "七、") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "八 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "八、") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "九 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "九、") > 0 Then
           RemoveOn = False
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 1)) And (InStr(Mid(TempPar.Range, 2, 1), "．") > 0 Or InStr(Mid(TempPar.Range, 2, 1), ".") > 0) Then
            RemoveOn = False
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 2)) And (InStr(Mid(TempPar.Range, 3, 1), "．") > 0 Or InStr(Mid(TempPar.Range, 3, 1), ".") > 0) Then
            RemoveOn = False
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 3)) And (InStr(Mid(TempPar.Range, 4, 1), "．") > 0 Or InStr(Mid(TempPar.Range, 4, 1), ".") > 0) Then
            RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "答案") > 0 Then
            RemoveOn = True
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "解析") > 0 Then
            RemoveOn = True
        End If
        If RemoveOn = True Then
            TempPar.Range.Delete
        End If
    Next
    CommandBars("Navigation").Visible = False                   '去除“保留图片”标记
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "【【"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
