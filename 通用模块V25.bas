Attribute VB_Name = "ͨ��ģ��V25"
' ���ƣ������ͨ�ð�
' �汾��V2.5
' ���ߣ�����
' ��λ��ƽԭ�ص�һ��ѧ
' ��־������̽��ʽɾ����Դ��Ϣ����߱�׼ȷ�Ϳ���չ�ԣ�����ɾ��������ʽ��ʵ�ַ�ʽ�������汾��V2.1
' ��־��������ռ��а�ģ�飬���ڸ�����𰸺�����ո��Ƶ����ݣ��Ӷ����ٳ��ֱ���ʱ���ڱ������һ�����ʾ�����汾��V2.2
'       ͬʱ������̽��ʽɾ����Դ��Ϣģ��Ϊͨ�ð汾
' ��־���Ż�����ع��ܣ�ͬʱ����ת����ϸ�ڣ������汾��V2.3
' ��־���ص��Ż�ɾ����Դͨ�ð棬V2.3��һ���޸Ĺ����еĲ������棬���ڲ����˹�����������ΪV2.4
' ��־������Ŀ������ѡ������ͳһ����һ��ģ�飺��Ŀ�������� ������������Ŀ����ģ��
'       �Ż���ҳü��ҳ�ŵ����ã����ڡ����̡̳�����������֧�֣�ͬʱ�е��ļ����Ʋ�һ���Ǳ������ݣ�����ҳü��Ӧ�����ļ��ڲ�ȡ��
'       ��ǿ�����̡̳��й��ڡ���ʽ���ṹ��̽�⴦�������汾��V2.5
'
' ClearOfficeClipBoard��Դ��https://stackoverflow.com/questions/14440274/cant-clear-office-clipboard-with-vba
' ClearOfficeClipBoard˵����������������ַ����ʱ���������޸�ʹ֮������ȷѡ��VBA�汾���У�ԭʼ�汾ԭ����ȷ�����ǽṹ����
'
Public AnswerTitle As String
Public YeMeiTitle As String
Public myVBA7 As Integer
Private Declare PtrSafe Function AccessibleChildren Lib "oleacc" (ByVal paccContainer As Office.IAccessible, _
                                                                  ByVal iChildStart As Long, ByVal cChildren As Long, _
                                                                  ByRef rgvarChildren As Any, ByRef pcObtained As Long) As Long
Sub ����������()
    Application.ScreenUpdating = False
    Call ȡ����Ŀ����
    Call ͳһ����
    Call ���ı�ʽ��ʽ
    Call ����ҳüҳ��
    Call ɾ������
    Call �淶���
    Call ����ϰ��
    Call �����ͨ�ð�
    Call ��Ŀ��������
    Call ��ͼƬ����
    Call У���м��
    Call ɾ����Դͨ�ð�
    Call �𰸵�����ҳ
    Call ClearOfficeClipBoard                           '��ռ��а�
    Selection.HomeKey Unit:=wdStory
'    ActiveDocument.Save
    Application.ScreenUpdating = True
End Sub
Sub ͳһ����()
    Dim objEq As OMath
    Selection.WholeStory
    Selection.Font.Name = "����"
    Selection.HomeKey Unit:=wdStory
' ����ȫ��ͳһΪ�����塱������ѧ��ʽҲ���ŷ����˱仯�����Դ˴����������ѧ��ʽ�ٻָ�Ϊ��ѧ����
    For Each objEq In ActiveDocument.OMaths
        objEq.Range.Font.Name = "Cambria Math"
    Next
End Sub
Sub ���ı�ʽ��ʽ()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\[��ʽ([0-9])\]"
        .Replacement.Text = "��ʽ\1"
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
        .Text = "\[��ʽ([0-9])��([0-9])\]"
        .Replacement.Text = "��ʽ\1.\2"
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
Sub ȡ����Ŀ����()
    Dim TempPar As Paragraph
    For Each TempPar In ActiveDocument.Paragraphs
        If InStr(TempPar.Range, "��1��") Then
           YeMeiTitle = Mid(TempPar.Range, 1, Len(TempPar.Range) - 1)
        End If
    Next
    If YeMeiTitle = "" Then
        YeMeiTitle = Mid(ActiveDocument.Paragraphs(1).Range, 1, Len(ActiveDocument.Paragraphs(1).Range) - 1)
    End If
    AnswerTitle = YeMeiTitle & "���ο��𰸡�"              '���òο��𰸸�ʽ
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
    Selection.TypeText Text:=YeMeiTitle
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub
Sub ɾ������()
    Dim TempPar As Paragraph
    For Each TempPar In ActiveDocument.Paragraphs
        If Len(TempPar.Range) = 1 Then
           TempPar.Range.Delete
        End If
    Next
End Sub
Sub ����ϰ��()
    Dim Doc As Document
    Dim rngDoc As Range
    Dim i, AnswerBegin As Integer
    AnswerBegin = 2
    For Each TempPar In ActiveDocument.Paragraphs
        i = i + 1
        If InStr(TempPar.Range, "�����㼯ѵ��") > 0 Then
            AnswerBegin = i
        ElseIf InStr(TempPar.Range, "��������ѵ��") > 0 Then
            AnswerBegin = i
        ElseIf InStr(TempPar.Range, "�µ���ͨ") > 0 Then
            AnswerBegin = i
        End If
    Next
    Set Doc = ActiveDocument
    Doc.Range(Start:=Doc.Paragraphs(AnswerBegin).Range.Start, End:=Doc.Paragraphs(Doc.Paragraphs.Count).Range.End).Copy
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
Public Sub ClearOfficeClipBoard()
    If VBA7 Then
        myVBA7 = 1
    Else
        myVBA7 = 0
    End If
    Dim cmnB, IsVis As Boolean, j As Long, Arr As Variant
    Arr = Array(4, 7, 2, 0)                                                     '4 and 2 for 32 bit, 7 and 0 for 64 bit
    Set cmnB = Application.CommandBars("Office Clipboard")
    IsVis = cmnB.Visible
    If Not IsVis Then
        cmnB.Visible = True
        DoEvents
    End If
    For j = 1 To Arr(0 + myVBA7)
        AccessibleChildren cmnB, Choose(j, 0, 3, 0, 3, 0, 3, 1), 1, cmnB, 1
    Next
    cmnB.accDoDefaultAction CLng(Arr(2 + myVBA7))
    Application.CommandBars("Office Clipboard").Visible = IsVis
End Sub

Sub �𰸵�����ҳ()
    Dim TempPar As Paragraph
    For Each TempPar In ActiveDocument.Paragraphs
        If InStr(TempPar.Range, AnswerTitle) > 0 Then
            TempPar.Range.Select
            Selection.HomeKey Unit:=wdLine
            Selection.InsertBreak Type:=wdPageBreak
        End If
    Next
End Sub

Sub �����ͨ�ð�()
    Dim TempPar As Paragraph
    Dim RemoveOn As Boolean
    Dim IsAnswer As Boolean
    Dim DuDianShuTong As Boolean
    Dim TiHao As String
    IsAnswer = False
    For n = 1 To ActiveDocument.InlineShapes.Count                      '������ͼƬ��Ϊ��������
        If ActiveDocument.InlineShapes(n).Width > 350 Then
            ActiveDocument.InlineShapes(n).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.TypeText Text:="����"
        End If
    Next
    For Each TempPar In ActiveDocument.Paragraphs
        If InStr(TempPar.Range, AnswerTitle) > 0 And Len(AnswerTitle) > 0 Then
            IsAnswer = True
            RemoveOn = False
        End If
        If InStr(TempPar.Range, "�µ���ͨ") > 0 Then
            DuDianShuTong = True
        End If
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
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "��") > 0 Then
            If IsAnswer = True Then
                RemoveOn = True
            Else
                RemoveOn = False
            End If
            TiHao = Mid(TempPar.Range, 1, 2)                                                                               'һ��ÿ�鲻�ᳬ��9������
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 1)) And (InStr(Mid(TempPar.Range, 2, 1), "��") > 0 Or InStr(Mid(TempPar.Range, 2, 1), ".") > 0) Then
            If IsAnswer = True Then
                RemoveOn = True
                TiHao = Mid(TempPar.Range, 1, 2)
            Else
                RemoveOn = False
            End If
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 2)) And (InStr(Mid(TempPar.Range, 3, 1), "��") > 0 Or InStr(Mid(TempPar.Range, 3, 1), ".") > 0) Then
            If IsAnswer = True Then
                RemoveOn = True
                TiHao = Mid(TempPar.Range, 1, 3)
            Else
                RemoveOn = False
            End If
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 3)) And (InStr(Mid(TempPar.Range, 4, 1), "��") > 0 Or InStr(Mid(TempPar.Range, 4, 1), ".") > 0) Then
            If IsAnswer = True Then
                RemoveOn = True
                TiHao = Mid(TempPar.Range, 1, 4)
            Else
                RemoveOn = False
            End If
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "��ʽ") > 0 Then
            If IsAnswer = True Then
                RemoveOn = True
                If InStr(Mid(TempPar.Range, 1, 5), ".") Then
                    TiHao = Mid(TempPar.Range, 1, 5)                    '��ʽ�ĸ���Ҳ����9��֮�ڣ����Կ�������ʽ����
                Else
                    TiHao = Mid(TempPar.Range, 1, 3)
                End If
            Else
                  RemoveOn = False
            End If
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "��") > 0 Then
            If IsAnswer = True Then
                RemoveOn = False
                If DuDianShuTong = True Then
                    DuDianShuTong = False
                Else
                    TempPar.Range.InsertBefore Text:=TiHao
                End If
            Else
                RemoveOn = True
            End If
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
            If IsAnswer = True Then
                RemoveOn = False
            Else
                RemoveOn = True
            End If
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
Sub �淶���()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "��"
        .Replacement.Text = "."
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
End Sub
Sub ɾ����Դͨ�ð�()
' д�˳����ԭ�������е������ϳ�����������(2020XXX(��)XX)�ں�С���ŵ���Դ��Ϣ�����ݾ�������Ͻṹ��Ӷ�Ӧ����Դ��ʽ����˸���׼ȷ��
' ͬʱ����׷����Դ�ṹ����չ֧�ֵĸ�ʽ�����Կ��ǽ�������ͨ�ó���
    Dim j, k, m As Integer
    Dim RepStr As String
' ����������ʽ������Դ��ͷͳһ�������ַ�"DELETE"
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\([1-2][0-9][0-9][0-9]"
        .Replacement.Text = "DELETE"
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
        .Text = "\[([1-2][0-9][0-9][0-9]*)\]"
        .Replacement.Text = "(\1)"
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
' ����"DELETE"�Ķ������ִ��ɾ����Դ�Ķ���
    For n = 1 To ActiveDocument.Paragraphs.Count
        If InStr(ActiveDocument.Paragraphs(n).Range, "DELETE") Then
            k = 1: RepStr = ""
            Do Until InStr(Mid(ActiveDocument.Paragraphs(n).Range, 1, k), "DELETE") > 0
                k = k + 1
            Loop
            k = k - 5
            j = k + 1
            Do Until InStr(Mid(ActiveDocument.Paragraphs(n).Range, k, j - k), ")") > 0
                j = j + 1
            Loop
            If j - k - 1 > 0 Then
                If InStr(Mid(ActiveDocument.Paragraphs(n).Range, k + 1, j - k - 1), "(") > 0 Then
                    j = j + 1
                    m = j
                    Do Until InStr(Mid(ActiveDocument.Paragraphs(n).Range, m, j - m), ")") > 0
                        j = j + 1
                    Loop
                End If
            End If
            RepStr = Mid(ActiveDocument.Paragraphs(n).Range, k, j - k)
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

Sub ��������(LeftNum, RightNum, FirstNum)
'
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
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
        .CharacterUnitLeftIndent = LeftNum
        .CharacterUnitRightIndent = RightNum
        .CharacterUnitFirstLineIndent = FirstNum
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
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
Sub ��Ŀ��������()
' ����������ģ��
    Dim TempPar As Paragraph
    Dim LeftInd, RightInd, FirstInd As Single  '�����������������������Ա��պ�����ʹ��
    For Each TempPar In ActiveDocument.Paragraphs
        If Len(TempPar.Range) = 1 Then
            FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "��") > 0 And InStr(Mid(TempPar.Range, 1, 4), "��") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "��") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
            LeftOn = False: RightOn = False: FirstOn = False
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "֪ʶ��") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "A��") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "B��") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "C��") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "D��") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "[��ʦ") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "ר��") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 1), "ͼ") > 0 Then         '��ЩͼƬ�����±꣬�����������Ӧ����������Ҳ��������
 '            FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "ʱ��:") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "Ӧ��һ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "Ӧ�ö�") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "Ӧ����") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "����ƪ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "����ƪ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "�ۺ�ƪ") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "Ӧ��ƪ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "����ƪ") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "һ ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "һ��") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call ��������(0, 0, 0)
             Call ��������(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call ��������(0, 0, 0)
             Call ��������(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call ��������(0, 0, 0)
             Call ��������(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "�ġ�") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call ��������(0, 0, 0)
             Call ��������(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "�塢") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call ��������(0, 0, 0)
             Call ��������(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call ��������(0, 0, 0)
             Call ��������(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "�ߡ�") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call ��������(0, 0, 0)
             Call ��������(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "�ˡ�") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call ��������(0, 0, 0)
             Call ��������(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "�š�") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call ��������(0, 0, 0)
             Call ��������(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "��ʽ") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call ��������(0, 0, 0)
             Call ��������(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "��") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call ��������(0, 0, 0)
             Call ��������(0, 0, -FirstInd)
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 1)) And InStr(Mid(TempPar.Range, 2, 1), ".") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call ��������(0, 0, 0)
             Call ��������(0, 0, -FirstInd)
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 2)) And InStr(Mid(TempPar.Range, 3, 1), ".") > 0 Then
             FirstInd = 1.5
             TempPar.Range.Select
             Call ��������(0, 0, 0)
             Call ��������(0, 0, -FirstInd)
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 3)) And InStr(Mid(TempPar.Range, 4, 1), ".") > 0 Then
             FirstInd = 2
             TempPar.Range.Select
             Call ��������(0, 0, 0)
             Call ��������(0, 0, -FirstInd)
        ElseIf FirstInd > 0 Then
             TempPar.Range.Select
             Call ��������(0, 0, 0)
             Call ��������(FirstInd, 0, 0)
        End If
    Next
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
