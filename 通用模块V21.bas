Attribute VB_Name = "通用模块V21"
' 名称：清除答案通用版
' 版本：V2.1
' 作者：冯振华
' 单位：平原县第一中学
' 日志：增加探测式删除题源信息，其具备准确和可拓展性，考虑删除正则表达式的实现方式，升级版本号V2.1

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
    Selection.Find.ClearFormatting                                  '去除“保留图片”标记
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

Sub 删除题源1()
' 此版是非正则表达式版，其首先探测题源结构，然后删除。写此程序的原因在于有的资料上出现了类似于(2020XXX(下)XX)内含小括号的题源信息，虽相较于正则表达式版复杂些，却可以
' 根据具体的资料结构添加对应的题源形式，因此更加准确。同时方便追加题源结构，拓展支持的格式，所以考虑将其列入通用程序
    Dim j, k, m As Integer
    Dim RepStr As String
    k = 5: j = k + 1
    For n = 1 To ActiveDocument.Paragraphs.Count
        If InStr(Mid(ActiveDocument.Paragraphs(n).Range, 3, 2), "(") > 0 Then
            Do Until InStr(Mid(ActiveDocument.Paragraphs(n).Range, 1, k), "(") = 0
                k = k - 1
            Loop
            j = k + 1
            Do Until InStr(Mid(ActiveDocument.Paragraphs(n).Range, 1, j), ")") > 0
                j = j + 1
            Loop
            m = j
            If InStr(Mid(ActiveDocument.Paragraphs(n).Range, k + 2, m - k + 1), "(") > 0 Then
                Do Until InStr(Mid(ActiveDocument.Paragraphs(n).Range, m + 1, j - m + 2), ")") > 0
                    j = j + 1
                Loop
                j = j + 2
            End If
                RepStr = Mid(ActiveDocument.Paragraphs(n).Range, k + 1, j - k)
        End If
        If InStr(RepStr, "20") > 0 Or InStr(RepStr, "19") > 0 Then
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = RepStr
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchByte = False
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchWildcards = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
        End If
    Next
End Sub

Sub 删除题源2()
' 此版为正则表达式版，其简洁明了，但是仅始于只用一层小括号包围的题源结构，如(2020XXX),如果题源结构发生变化则不适用此程序，这也是考虑在通用版本中不采用此程序的原因
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\(20*\)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "\(19*\)"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
