Attribute VB_Name = "SeperateAnswerģ��"
'��Ŀ������ϰ����𰸷������
'���ߣ�����
'�汾��V1.0
Public AnswerTitle As String
Sub SeperateAnswer()
    Application.ScreenUpdating = False
    AnswerTitle = Mid(ActiveDocument.Paragraphs(1).Range, 1, Len(ActiveDocument.Paragraphs(1).Range) - 1) & "�ο���"
    Call ����ҳüҳ��
    Call ɾ������
    Call ����ϰ��
    Call ɾ����
    Call ��Ŀ��������
    Call ��ʽ��ѡ����ѡ��
    Call ��ͼƬ����
    Call У���м��
    Selection.HomeKey Unit:=wdStory
    ActiveDocument.Save
    Application.ScreenUpdating = True
End Sub

Sub ����ϰ��()
    Dim doc As Document
    Dim rngDoc As Range
    Dim i, AnswerBegin As Integer
    AnswerBegin = 0
    For Each TempPar In ActiveDocument.Paragraphs
        i = i + 1
        If InStr(TempPar.Range, "�����㼯ѵ��") > 0 Then
            AnswerBegin = i
        End If
    Next
    If AnswerBegin = 0 Then
        AnswerBegin = 2
    End If
    Set doc = ActiveDocument
    doc.Range(Start:=doc.Paragraphs(AnswerBegin).Range.Start, End:=doc.Paragraphs(doc.Paragraphs.Count).Range.End).Copy
    Selection.EndKey Unit:=wdStory
    Selection.InsertBreak Type:=wdPageBreak
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Size = 16
    Selection.Font.Bold = wdToggle
    Selection.Font.Name = "����"
    Selection.TypeText Text:=AnswerTitle
    Selection.TypeParagraph
    Selection.Paste
End Sub

Sub ɾ������()
    Dim i As Integer
    Dim TempPar As Paragraph
    Dim TempLine As Line
    For Each TempPar In ActiveDocument.Paragraphs
        If Len(TempPar.Range) = 1 Then
            n = n + 1
           TempPar.Range.Delete
        End If
    Next
End Sub

' ����Ϊ���ú�
'
Sub ɾ����()
    Dim i As Integer
    Dim TempPar As Paragraph
    Dim TempLine As Line
    Dim RemoveOn, AnswerOn As Boolean
    Dim TiHao As String
    AnswerOn = False
    For Each TempPar In ActiveDocument.Paragraphs
        If InStr(TempPar.Range, AnswerTitle) > 0 Then
           AnswerOn = True
           RemoveOn = False
        End If
        If IsNumeric(Mid(TempPar.Range, 1, 1)) And InStr(Mid(TempPar.Range, 2, 1), ".") > 0 Then
           If AnswerOn = False Then
            RemoveOn = False
           Else
            RemoveOn = True
           End If
           TiHao = Mid(TempPar.Range, 1, 2)
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 2)) And InStr(Mid(TempPar.Range, 3, 1), ".") > 0 Then
           If AnswerOn = False Then
            RemoveOn = False
           Else
            RemoveOn = True
           End If
           TiHao = Mid(TempPar.Range, 1, 3)
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 3)) And InStr(Mid(TempPar.Range, 4, 1), ".") > 0 Then
           If AnswerOn = False Then
            RemoveOn = False
           Else
            RemoveOn = True
           End If
           TiHao = Mid(TempPar.Range, 1, 4)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "��") > 0 Then
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
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "[��ʦ") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "һ��") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�ġ�") > 0 Then
           RemoveOn = False
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "��") > 0 Then
           If AnswerOn = False Then
            RemoveOn = True
           Else
            RemoveOn = False
            TempPar.Range.InsertBefore Text:=TiHao
           End If
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
           If AnswerOn = False Then
            RemoveOn = True
           Else
            RemoveOn = False
           End If
        End If
        If RemoveOn = True Then
            TempPar.Range.Delete
        End If
    Next
End Sub
Sub ��Ŀ��������()
    Dim i As Integer
    Dim TempPar As Paragraph
    Dim TempLine As Line
    Dim SuoJinOn As Boolean
    Dim TMInteger As Single
    For Each TempPar In ActiveDocument.Paragraphs
        If IsNumeric(Mid(TempPar.Range, 1, 1)) And InStr(Mid(TempPar.Range, 2, 1), ".") > 0 Then
           SuoJinOn = True
           TMInteger = 1
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 2)) And InStr(Mid(TempPar.Range, 3, 1), ".") > 0 Then
           SuoJinOn = True
           TMInteger = 1.5
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 3)) And InStr(Mid(TempPar.Range, 4, 1), ".") > 0 Then
           SuoJinOn = True
           TMInteger = 2
        Else
           SuoJinOn = False
        End If
        If SuoJinOn = True Then
            TempPar.Range.Select
            Call ��������(TMInteger)
        End If
    Next
End Sub
Sub ��������(InInteger)
'
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0.18)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(-0.18)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = -InInteger
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
'        .CollapsedByDefault = False 'Ϊ�˼���office2007ɾ������
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = False
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
End Sub


Sub ��ʽ��ѡ����ѡ��()
' �д���չΪ����ǿ���ѡ���жϹ���
    Dim i, j, k As Integer
    Dim TempPar As Paragraph
    Dim TempLine As Line
    Dim SuoJinOn As Boolean
    Dim TMInteger As Single
    For Each TempPar In ActiveDocument.Paragraphs
        SuoJinOn = False
        If IsNumeric(Mid(TempPar.Range, 1, 1)) And InStr(Mid(TempPar.Range, 2, 1), ".") > 0 Then
           SuoJinOn = True
           TMInteger = 1
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 2)) And InStr(Mid(TempPar.Range, 3, 1), ".") > 0 Then
           SuoJinOn = True
           TMInteger = 1.5
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 3)) And InStr(Mid(TempPar.Range, 4, 1), ".") > 0 Then
           SuoJinOn = True
           TMInteger = 2
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "��") > 0 Then
           SuoJinOn = True
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
           SuoJinOn = True
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "A��") > 0 Then
           SuoJinOn = True
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "B��") > 0 Then
           SuoJinOn = True
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "C��") > 0 Then
           SuoJinOn = True
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "D��") > 0 Then
           SuoJinOn = True
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "[��ʦ") > 0 Then
           SuoJinOn = True
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "ר��") > 0 Then
           SuoJinOn = True
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "һ��") > 0 Then
           SuoJinOn = True
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
           SuoJinOn = True
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
           SuoJinOn = True
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�ġ�") > 0 Then
           SuoJinOn = True
        End If
        If InStr(TempPar.Range, "A.") > 0 Or InStr(TempPar.Range, "B.") > 0 Or InStr(TempPar.Range, "C.") > 0 Or InStr(TempPar.Range, "D.") > 0 Then
            TempPar.Range.Select
            Call ѡ������(TMInteger)
        ElseIf SuoJinOn = False Then
            TempPar.Range.Select
            Call ѡ������(0)
            Call ѡ������(TMInteger)
        End If
    Next
End Sub
Sub ѡ������(ChoiceInteger)
'
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0.18)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = ChoiceInteger
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
  '      .CollapsedByDefault = False
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = False
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
End Sub

Sub ����ҳüҳ��()
    WordBasic.RemoveHeader
    WordBasic.RemoveFooter
    ActiveDocument.Sections.PageSetup.DifferentFirstPageHeaderFooter = True
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    With Selection.ParagraphFormat
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.TypeText Text:="������" & vbTab & "�༶��" & vbTab
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "DATE  \@ ""yyyy-MM-dd"" ", PreserveFormatting:=True
    WordBasic.GoToFooter
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="��"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "PAGE  \* Arabic ", PreserveFormatting:=True
    Selection.TypeText Text:="ҳ ��"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "NUMPAGES  \* Arabic ", PreserveFormatting:=True
    Selection.TypeText Text:="ҳ"
    ActiveWindow.ActivePane.View.NextHeaderFooter
    WordBasic.RemoveHeader
    WordBasic.RemoveFooter
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="��"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "PAGE  \* Arabic ", PreserveFormatting:=True
    Selection.TypeText Text:="ҳ ��"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "NUMPAGES  \* Arabic ", PreserveFormatting:=True
    Selection.TypeText Text:="ҳ"
    WordBasic.GoToHeader
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:=Mid(ActiveDocument.Paragraphs(1).Range, 1, Len(ActiveDocument.Paragraphs(1).Range) - 1)
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub
Sub ��ͼƬ����()
    Dim TempPic As InlineShapes
    For n = 1 To ActiveDocument.InlineShapes.Count
        ActiveDocument.InlineShapes(n).Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Next
End Sub
Sub У���м��()

    Selection.WholeStory
    With Selection.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .DisableLineHeightGrid = True
        .WordWrap = True
    End With
End Sub

