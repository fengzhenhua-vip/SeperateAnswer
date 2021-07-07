Attribute VB_Name = "ͨ��ģ��V21"
' ���ƣ������ͨ�ð�
' �汾��V2.1
' ���ߣ�����
' ��λ��ƽԭ�ص�һ��ѧ
' ��־������̽��ʽɾ����Դ��Ϣ����߱�׼ȷ�Ϳ���չ�ԣ�����ɾ��������ʽ��ʵ�ַ�ʽ�������汾��V2.1

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
    Selection.Find.ClearFormatting                                  'ȥ��������ͼƬ�����
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

Sub ɾ����Դ1()
' �˰��Ƿ�������ʽ�棬������̽����Դ�ṹ��Ȼ��ɾ����д�˳����ԭ�������е������ϳ�����������(2020XXX(��)XX)�ں�С���ŵ���Դ��Ϣ���������������ʽ�渴��Щ��ȴ����
' ���ݾ�������Ͻṹ��Ӷ�Ӧ����Դ��ʽ����˸���׼ȷ��ͬʱ����׷����Դ�ṹ����չ֧�ֵĸ�ʽ�����Կ��ǽ�������ͨ�ó���
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

Sub ɾ����Դ2()
' �˰�Ϊ������ʽ�棬�������ˣ����ǽ�ʼ��ֻ��һ��С���Ű�Χ����Դ�ṹ����(2020XXX),�����Դ�ṹ�����仯�����ô˳�����Ҳ�ǿ�����ͨ�ð汾�в����ô˳����ԭ��
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
