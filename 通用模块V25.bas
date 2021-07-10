Attribute VB_Name = "通用模块V25"
' 名称：清除答案通用版
' 版本：V2.5
' 作者：冯振华
' 单位：平原县第一中学
' 日志：增加探测式删除题源信息，其具备准确和可拓展性，考虑删除正则表达式的实现方式，升级版本号V2.1
' 日志：增加清空剪切板模块，用于复制完答案后再清空复制的内容，从而不再出现保存时关于保留最后一项的提示升级版本号V2.2
'       同时采用了探测式删除题源信息模块为通用版本
' 日志：优化了相关功能，同时增加转换的细节，升级版本号V2.3
' 日志：重点优化删除题源通用版，V2.3是一个修改过程中的不完整版，由于补充了功能所以升级为V2.4
' 日志：将题目缩进和选项缩进统一到了一个模块：题目悬挂缩进 ，并精简了题目缩进模块
'       优化了页眉和页脚的设置，对于《金版教程》的设置增加支持，同时有的文件名称不一定是本节内容，所以页眉名应当从文件内部取定
'       增强《金版教程》中关于“变式”结构的探测处理，升级版本号V2.5
'
' ClearOfficeClipBoard来源：https://stackoverflow.com/questions/14440274/cant-clear-office-clipboard-with-vba
' ClearOfficeClipBoard说明：在引用上述网址代码时，我做了修改使之可以正确选择VBA版本运行，原始版本原理正确，但是结构有误
'
Public AnswerTitle As String
Public YeMeiTitle As String
Public myVBA7 As Integer
Private Declare PtrSafe Function AccessibleChildren Lib "oleacc" (ByVal paccContainer As Office.IAccessible, _
                                                                  ByVal iChildStart As Long, ByVal cChildren As Long, _
                                                                  ByRef rgvarChildren As Any, ByRef pcObtained As Long) As Long
Sub 分离主程序()
    Application.ScreenUpdating = False
    Call 取得题目名称
    Call 统一字体
    Call 更改变式格式
    Call 设置页眉页脚
    Call 删除空行
    Call 规范标点
    Call 复制习题
    Call 分离答案通用版
    Call 题目悬挂缩进
    Call 简单图片处理
    Call 校正行间距
    Call 删除题源通用版
    Call 答案单独分页
    Call ClearOfficeClipBoard                           '清空剪切板
    Selection.HomeKey Unit:=wdStory
'    ActiveDocument.Save
    Application.ScreenUpdating = True
End Sub
Sub 统一字体()
    Dim objEq As OMath
    Selection.WholeStory
    Selection.Font.Name = "宋体"
    Selection.HomeKey Unit:=wdStory
' 由于全部统一为“宋体”所，数学公式也跟着发生了变化，所以此处针对所有数学公式再恢复为数学字体
    For Each objEq In ActiveDocument.OMaths
        objEq.Range.Font.Name = "Cambria Math"
    Next
End Sub
Sub 更改变式格式()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\[变式([0-9])\]"
        .Replacement.Text = "变式\1"
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
        .Text = "\[变式([0-9])－([0-9])\]"
        .Replacement.Text = "变式\1.\2"
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
Sub 取得题目名称()
    Dim TempPar As Paragraph
    For Each TempPar In ActiveDocument.Paragraphs
        If InStr(TempPar.Range, "第1讲") Then
           YeMeiTitle = Mid(TempPar.Range, 1, Len(TempPar.Range) - 1)
        End If
    Next
    If YeMeiTitle = "" Then
        YeMeiTitle = Mid(ActiveDocument.Paragraphs(1).Range, 1, Len(ActiveDocument.Paragraphs(1).Range) - 1)
    End If
    AnswerTitle = YeMeiTitle & "【参考答案】"              '设置参考答案格式
End Sub
Sub 设置页眉页脚()
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
    Selection.TypeText Text:="姓名：" & vbTab & "班级：" & vbTab
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "DATE  \@ ""yyyy-MM-dd"" ", PreserveFormatting:=True
    WordBasic.GoToFooter
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="第"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "PAGE  \* Arabic ", PreserveFormatting:=True
    Selection.TypeText Text:="页 共"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "NUMPAGES  \* Arabic ", PreserveFormatting:=True
    Selection.TypeText Text:="页"
    ActiveWindow.ActivePane.View.NextHeaderFooter
    WordBasic.RemoveHeader
    WordBasic.RemoveFooter
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="第"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "PAGE  \* Arabic ", PreserveFormatting:=True
    Selection.TypeText Text:="页 共"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "NUMPAGES  \* Arabic ", PreserveFormatting:=True
    Selection.TypeText Text:="页"
    WordBasic.GoToHeader
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:=YeMeiTitle
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub
Sub 删除空行()
    Dim TempPar As Paragraph
    For Each TempPar In ActiveDocument.Paragraphs
        If Len(TempPar.Range) = 1 Then
           TempPar.Range.Delete
        End If
    Next
End Sub
Sub 复制习题()
    Dim Doc As Document
    Dim rngDoc As Range
    Dim i, AnswerBegin As Integer
    AnswerBegin = 2
    For Each TempPar In ActiveDocument.Paragraphs
        i = i + 1
        If InStr(TempPar.Range, "【考点集训】") > 0 Then
            AnswerBegin = i
        ElseIf InStr(TempPar.Range, "【基础集训】") > 0 Then
            AnswerBegin = i
        ElseIf InStr(TempPar.Range, "堵点疏通") > 0 Then
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
    Selection.Font.Name = "宋体"
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

Sub 答案单独分页()
    Dim TempPar As Paragraph
    For Each TempPar In ActiveDocument.Paragraphs
        If InStr(TempPar.Range, AnswerTitle) > 0 Then
            TempPar.Range.Select
            Selection.HomeKey Unit:=wdLine
            Selection.InsertBreak Type:=wdPageBreak
        End If
    Next
End Sub

Sub 分离答案通用版()
    Dim TempPar As Paragraph
    Dim RemoveOn As Boolean
    Dim IsAnswer As Boolean
    Dim DuDianShuTong As Boolean
    Dim TiHao As String
    IsAnswer = False
    For n = 1 To ActiveDocument.InlineShapes.Count                      '保留以图片作为标题的情况
        If ActiveDocument.InlineShapes(n).Width > 350 Then
            ActiveDocument.InlineShapes(n).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.TypeText Text:="【【"
        End If
    Next
    For Each TempPar In ActiveDocument.Paragraphs
        If InStr(TempPar.Range, AnswerTitle) > 0 And Len(AnswerTitle) > 0 Then
            IsAnswer = True
            RemoveOn = False
        End If
        If InStr(TempPar.Range, "堵点疏通") > 0 Then
            DuDianShuTong = True
        End If
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
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "例") > 0 Then
            If IsAnswer = True Then
                RemoveOn = True
            Else
                RemoveOn = False
            End If
            TiHao = Mid(TempPar.Range, 1, 2)                                                                               '一般每组不会超过9个例题
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 1)) And (InStr(Mid(TempPar.Range, 2, 1), "．") > 0 Or InStr(Mid(TempPar.Range, 2, 1), ".") > 0) Then
            If IsAnswer = True Then
                RemoveOn = True
                TiHao = Mid(TempPar.Range, 1, 2)
            Else
                RemoveOn = False
            End If
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 2)) And (InStr(Mid(TempPar.Range, 3, 1), "．") > 0 Or InStr(Mid(TempPar.Range, 3, 1), ".") > 0) Then
            If IsAnswer = True Then
                RemoveOn = True
                TiHao = Mid(TempPar.Range, 1, 3)
            Else
                RemoveOn = False
            End If
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 3)) And (InStr(Mid(TempPar.Range, 4, 1), "．") > 0 Or InStr(Mid(TempPar.Range, 4, 1), ".") > 0) Then
            If IsAnswer = True Then
                RemoveOn = True
                TiHao = Mid(TempPar.Range, 1, 4)
            Else
                RemoveOn = False
            End If
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "变式") > 0 Then
            If IsAnswer = True Then
                RemoveOn = True
                If InStr(Mid(TempPar.Range, 1, 5), ".") Then
                    TiHao = Mid(TempPar.Range, 1, 5)                    '变式的个数也是在9个之内，所以可以以左式设置
                Else
                    TiHao = Mid(TempPar.Range, 1, 3)
                End If
            Else
                  RemoveOn = False
            End If
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "答案") > 0 Then
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
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "解析") > 0 Then
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
Sub 规范标点()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "．"
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
Sub 删除题源通用版()
' 写此程序的原因在于有的资料上出现了类似于(2020XXX(下)XX)内含小括号的题源信息，根据具体的资料结构添加对应的题源形式，因此更加准确。
' 同时方便追加题源结构，拓展支持的格式，所以考虑将其列入通用程序
    Dim j, k, m As Integer
    Dim RepStr As String
' 借用正则表达式，将题源开头统一到特殊字符"DELETE"
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
' 含有"DELETE"的段落才是执行删除题源的段落
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

Sub 段落缩进(LeftNum, RightNum, FirstNum)
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
Sub 题目悬挂缩进()
' 改良版缩进模块
    Dim TempPar As Paragraph
    Dim LeftInd, RightInd, FirstInd As Single  '多设置了左缩进、右缩进以备日后升级使用
    For Each TempPar In ActiveDocument.Paragraphs
        If Len(TempPar.Range) = 1 Then
            FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "第") > 0 And InStr(Mid(TempPar.Range, 1, 4), "讲") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "【") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "考点") > 0 Then
            LeftOn = False: RightOn = False: FirstOn = False
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "知识点") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "A组") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "B组") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "C组") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "D组") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "[教师") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "专题") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 1), "图") > 0 Then         '有些图片存在下标，而此种情况不应当缩进，但也不必置零
 '            FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "时间:") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "应用一") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "应用二") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "应用三") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "备考篇") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "基础篇") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "综合篇") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "应用篇") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "创新篇") > 0 Then
             FirstInd = 0
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "一 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "一、") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call 段落缩进(0, 0, 0)
             Call 段落缩进(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "二 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "二、") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call 段落缩进(0, 0, 0)
             Call 段落缩进(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "三 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "三、") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call 段落缩进(0, 0, 0)
             Call 段落缩进(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "四 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "四、") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call 段落缩进(0, 0, 0)
             Call 段落缩进(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "五 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "五、") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call 段落缩进(0, 0, 0)
             Call 段落缩进(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "六 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "六、") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call 段落缩进(0, 0, 0)
             Call 段落缩进(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "七 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "七、") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call 段落缩进(0, 0, 0)
             Call 段落缩进(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "八 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "八、") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call 段落缩进(0, 0, 0)
             Call 段落缩进(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "九 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "九、") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call 段落缩进(0, 0, 0)
             Call 段落缩进(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "变式") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call 段落缩进(0, 0, 0)
             Call 段落缩进(0, 0, -FirstInd)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "例") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call 段落缩进(0, 0, 0)
             Call 段落缩进(0, 0, -FirstInd)
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 1)) And InStr(Mid(TempPar.Range, 2, 1), ".") > 0 Then
             FirstInd = 1
             TempPar.Range.Select
             Call 段落缩进(0, 0, 0)
             Call 段落缩进(0, 0, -FirstInd)
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 2)) And InStr(Mid(TempPar.Range, 3, 1), ".") > 0 Then
             FirstInd = 1.5
             TempPar.Range.Select
             Call 段落缩进(0, 0, 0)
             Call 段落缩进(0, 0, -FirstInd)
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 3)) And InStr(Mid(TempPar.Range, 4, 1), ".") > 0 Then
             FirstInd = 2
             TempPar.Range.Select
             Call 段落缩进(0, 0, 0)
             Call 段落缩进(0, 0, -FirstInd)
        ElseIf FirstInd > 0 Then
             TempPar.Range.Select
             Call 段落缩进(0, 0, 0)
             Call 段落缩进(FirstInd, 0, 0)
        End If
    Next
End Sub


Sub 简单图片处理()
    Dim TempPic As InlineShapes
    For n = 1 To ActiveDocument.InlineShapes.Count
        ActiveDocument.InlineShapes(n).Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Next
End Sub
Sub 校正行间距()

    Selection.WholeStory
    With Selection.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .DisableLineHeightGrid = True
        .WordWrap = True
    End With
End Sub
