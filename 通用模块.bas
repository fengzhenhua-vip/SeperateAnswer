Attribute VB_Name = "ͨ��ģ��"
Sub ɾ����ͨ�ð�()
    Dim TempPar As Paragraph
    Dim RemoveOn As Boolean
    For n = 1 To ActiveDocument.InlineShapes.Count                      '������ͼƬ��Ϊ��������
        If ActiveDocument.InlineShapes(n).Width > 300 Then
            ActiveDocument.InlineShapes(n).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.TypeText Text:="����"
        End If
    Next
    For Each TempPar In ActiveDocument.Paragraphs
        If InStr(Mid(TempPar.Range, 1, 4), "��") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "[") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "A��") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "B��") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "C��") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "D��") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "��") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "һ ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "һ��") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "�ġ�") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "�塢") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "�ߡ�") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "�ˡ�") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "�š�") > 0 Then
           RemoveOn = False
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 1)) And (InStr(Mid(TempPar.Range, 2, 1), "��") > 0 Or InStr(Mid(TempPar.Range, 2, 1), ".") > 0) Then
            RemoveOn = False
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 2)) And (InStr(Mid(TempPar.Range, 3, 1), "��") > 0 Or InStr(Mid(TempPar.Range, 3, 1), ".") > 0) Then
            RemoveOn = False
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 3)) And (InStr(Mid(TempPar.Range, 4, 1), "��") > 0 Or InStr(Mid(TempPar.Range, 4, 1), ".") > 0) Then
            RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "��") > 0 Then
            RemoveOn = True
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
            RemoveOn = True
        End If
        If RemoveOn = True Then
            TempPar.Range.Delete
        End If
    Next
    CommandBars("Navigation").Visible = False                   'ȥ��������ͼƬ�����
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "����"
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
