Attribute VB_Name = "ͨ��ģ��V27"
' ���ƣ������ͨ�ð�
' �汾��V27
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
' ��־�������˴�������ͬʱ����MsOfficeλ��32λ��64λ���жϽ���������ʹ֮����MsOffice2019��������ǰ�汾�������汾��V26
' ��־������SASReplaceģ�飬���˴��룬����Ч�ʣ������汾��V27
'
' ClearOfficeClipBoard��Դ��https://stackoverflow.com/questions/14440274/cant-clear-office-clipboard-with-vba
' ClearOfficeClipBoard˵����������������ַ����ʱ���������޸�ʹ֮������ȷѡ��VBA�汾���У�ԭʼ�汾ԭ����ȷ�����ǽṹ����
'
Public TempPar As Paragraph
Public AnswerTitle As String
Public YeMeiTitle As String
Public myVBA7 As Integer
Private Declare PtrSafe Function AccessibleChildren Lib "oleacc" (ByVal paccContainer As Office.IAccessible, _
                                                                  ByVal iChildStart As Long, ByVal cChildren As Long, _
                                                                  ByRef rgvarChildren As Any, ByRef pcObtained As Long) As Long
Sub ����������()
    Application.ScreenUpdating = False
    Call ȡ����Ŀ����
    Call UnifiedFont
    Call ���ı�ʽ��ʽ
    Call ����ҳüҳ��
    Call ɾ������
    Call �淶���
    Call ����ϰ��
    Call �����ͨ�ð�
    Call ��Ŀ��������
    Call У���м��
    Call ɾ����Դͨ�ð�
    Call �𰸵�����ҳ
    Call ClearOfficeClipBoard                           '��ռ��а�
    Selection.HomeKey Unit:=wdStory
'    ActiveDocument.Save
    Application.ScreenUpdating = True
End Sub
Public Sub ClearOfficeClipBoard()
' 2021/7/26����MsOffice2019�����д˴���VBA7���жϷ�ʽʧЧ�����Ը����˴˴��жϴ��룬�ο�������
    Dim sText As String
    sText = Environ("PROCESSOR_ARCHITECTURE")
    Debug.Print sText
    If sText Like "*64*" Then
        myVBA7 = 1
    ElseIf sText Like "*86*" Then
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
Sub UnifiedFont()
    Dim objEq As OMath
    ActiveDocument.Range.Font.Name = "����"
    For Each objEq In ActiveDocument.OMaths
        objEq.Range.Font.Name = "Cambria Math"
    Next
End Sub
Sub SASReplace(THText, THReplace, THBool)
        With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = THText
        .Replacement.Text = THReplace
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = THBool
        .Execute Replace:=wdReplaceAll
    End With
End Sub
Sub ����ҳüҳ��()
    Selection.HomeKey Unit:=wdStory
    WordBasic.RemoveHeader
    WordBasic.RemoveFooter
    ActiveDocument.Sections.PageSetup.DifferentFirstPageHeaderFooter = True
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
    With Selection
        .TypeText Text:="������" & vbTab & "�༶��" & vbTab
        .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "DATE  \@ ""yyyy��M��d��"" ", PreserveFormatting:=True
    End With
    WordBasic.GoToFooter
    With Selection
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .TypeText Text:="��"
        .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
            "PAGE  \* Arabic ", PreserveFormatting:=True
        .TypeText Text:="ҳ ��"
        .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
            "NUMPAGES  \* Arabic ", PreserveFormatting:=True
        .TypeText Text:="ҳ"
    End With
    If Application.Selection.Information(wdNumberOfPagesInDocument) > 1 Then   '���ҳ�����1��ִ�еڶ�ҳ��ҳüҳ������
        ActiveWindow.ActivePane.View.NextHeaderFooter
        WordBasic.RemoveHeader
        WordBasic.RemoveFooter
        With Selection
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .TypeText Text:="��"
            .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
                "PAGE  \* Arabic ", PreserveFormatting:=True
            .TypeText Text:="ҳ ��"
            .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
                "NUMPAGES  \* Arabic ", PreserveFormatting:=True
            .TypeText Text:="ҳ"
        End With
        WordBasic.GoToHeader
        With Selection
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .TypeText Text:=YeMeiTitle
        End With
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub
Sub ���ı�ʽ��ʽ()
    Call SASReplace("\[��ʽ([0-9])\]", "��ʽ\1", True)
    Call SASReplace("\[��ʽ([0-9])��([0-9])\]", "��ʽ\1.\2", True)
End Sub
Sub ȡ����Ŀ����()
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

Sub ɾ������()
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
    With Selection
        .EndKey Unit:=wdStory
        .InsertBreak Type:=wdPageBreak
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Size = 16
        .Font.Bold = wdToggle
        .Font.Name = "����"
        .TypeText Text:=AnswerTitle
        .TypeParagraph
        .Paste
    End With
End Sub

Sub �𰸵�����ҳ()
    For Each TempPar In ActiveDocument.Paragraphs
        If InStr(TempPar.Range, AnswerTitle) > 0 Then
            TempPar.Range.Select
            Selection.HomeKey Unit:=wdLine
            Selection.InsertBreak Type:=wdPageBreak
        End If
    Next
End Sub

Sub �����ͨ�ð�()
    Dim RemoveOn As Boolean
    Dim IsAnswer As Boolean
    Dim DuDianShuTong As Boolean
    Dim TiHao As String
    Dim n As Integer
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
            IsAnswer = True: RemoveOn = False
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
    Call SASReplace("����", "", False)                              'ȥ��������ͼƬ�����
End Sub

Sub �淶���()
    Call SASReplace("��", ".", False)
End Sub
Sub ɾ����Դͨ�ð�()
' д�˳����ԭ�������е������ϳ�����������(2020XXX(��)XX)�ں�С���ŵ���Դ��Ϣ�����ݾ�������Ͻṹ��Ӷ�Ӧ����Դ��ʽ����˸���׼ȷ��
' ͬʱ����׷����Դ�ṹ����չ֧�ֵĸ�ʽ�����Կ��ǽ�������ͨ�ó���
    Dim j, k, m As Integer
    Dim RepStr As String
' ����������ʽ������Դ��ͷͳһ�������ַ�"DELETE"
    Call SASReplace("\([1-2][0-9][0-9][0-9]", "DELETE", True)
    Call SASReplace("\[([1-2][0-9][0-9][0-9]*)\]", "(\1)", True)
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
            Call SASReplace(RepStr, "", False)
        End If
    Next
End Sub

Sub ��������(LeftNum, RightNum, FirstNum)
    With TempPar.Range.ParagraphFormat
        .CharacterUnitLeftIndent = LeftNum
        .CharacterUnitRightIndent = RightNum
        .CharacterUnitFirstLineIndent = FirstNum
    End With
End Sub
Sub ��Ŀ��������()
    Dim TiHaoNum As Single
    For Each TempPar In ActiveDocument.Paragraphs
        If InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
             Call ��������(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "֪ʶ��") > 0 Then
             Call ��������(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "һ ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "һ��") > 0 Then
             Call ��������(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
             Call ��������(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
             Call ��������(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "�ġ�") > 0 Then
             Call ��������(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "�塢") > 0 Then
             Call ��������(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
             Call ��������(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "�ߡ�") > 0 Then
             Call ��������(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "�ˡ�") > 0 Then
             Call ��������(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "�� ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "�š�") > 0 Then
             Call ��������(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "��ʽ") > 0 Then
             Call ��������(0, 0, -1): TiHaoNum = 1.3
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "��") > 0 Then
             Call ��������(0, 0, -1): TiHaoNum = 1.3
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 1)) And InStr(Mid(TempPar.Range, 2, 1), ".") > 0 Then
             Call ��������(0, 0, -1): TiHaoNum = 1.3
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 2)) And InStr(Mid(TempPar.Range, 3, 1), ".") > 0 Then
             Call ��������(0, 0, -1.5): TiHaoNum = 1.5
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 3)) And InStr(Mid(TempPar.Range, 4, 1), ".") > 0 Then
             Call ��������(0, 0, -2): TiHaoNum = 1.7
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "A.") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "B.") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "C.") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "D.") > 0 Then
             Call ��������(0, 0, TiHaoNum)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "(1)") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "(2)") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "(3)") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "(4)") > 0 Then
             Call ��������(0, 0, -TiHaoNum)
             TempPar.Range.ParagraphFormat.CharacterUnitLeftIndent = TiHaoNum
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "����") > 0 Then
             Call ��������(0, 0, -1)
        End If
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
