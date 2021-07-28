Attribute VB_Name = "通用模块V27"
' 名称：分离答案通用版
' 版本：V27
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
' 日志：精简了大量程序，同时对于MsOffice位数32位和64位的判断进行了升级使之适配MsOffice2019及兼容以前版本，升级版本号V26
' 日志：增加SASReplace模块，简化了代码，提升效率，升级版本号V27
'
' ClearOfficeClipBoard来源：https://stackoverflow.com/questions/14440274/cant-clear-office-clipboard-with-vba
' ClearOfficeClipBoard说明：在引用上述网址代码时，我做了修改使之可以正确选择VBA版本运行，原始版本原理正确，但是结构有误
'
Public TempPar As Paragraph
Public AnswerTitle As String
Public YeMeiTitle As String
Public myVBA7 As Integer
Private Declare PtrSafe Function AccessibleChildren Lib "oleacc" (ByVal paccContainer As Office.IAccessible, _
                                                                  ByVal iChildStart As Long, ByVal cChildren As Long, _
                                                                  ByRef rgvarChildren As Any, ByRef pcObtained As Long) As Long
Sub 分离主程序()
    Application.ScreenUpdating = False
    Call 取得题目名称
    Call UnifiedFont
    Call 更改变式格式
    Call 设置页眉页脚
    Call 删除空行
    Call 规范标点
    Call 复制习题
    Call 分离答案通用版
    Call 题目悬挂缩进
    Call 校正行间距
    Call 删除题源通用版
    Call 答案单独分页
    Call ClearOfficeClipBoard                           '清空剪切板
    Selection.HomeKey Unit:=wdStory
'    ActiveDocument.Save
    Application.ScreenUpdating = True
End Sub
Public Sub ClearOfficeClipBoard()
' 2021/7/26升级MsOffice2019后运行此代码VBA7的判断方式失效，所以更改了此处判断代码，参考了网络
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
    ActiveDocument.Range.Font.Name = "宋体"
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
Sub 设置页眉页脚()
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
        .TypeText Text:="姓名：" & vbTab & "班级：" & vbTab
        .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "DATE  \@ ""yyyy年M月d日"" ", PreserveFormatting:=True
    End With
    WordBasic.GoToFooter
    With Selection
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .TypeText Text:="第"
        .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
            "PAGE  \* Arabic ", PreserveFormatting:=True
        .TypeText Text:="页 共"
        .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
            "NUMPAGES  \* Arabic ", PreserveFormatting:=True
        .TypeText Text:="页"
    End With
    If Application.Selection.Information(wdNumberOfPagesInDocument) > 1 Then   '如果页码大于1则执行第二页的页眉页脚设置
        ActiveWindow.ActivePane.View.NextHeaderFooter
        WordBasic.RemoveHeader
        WordBasic.RemoveFooter
        With Selection
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .TypeText Text:="第"
            .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
                "PAGE  \* Arabic ", PreserveFormatting:=True
            .TypeText Text:="页 共"
            .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
                "NUMPAGES  \* Arabic ", PreserveFormatting:=True
            .TypeText Text:="页"
        End With
        WordBasic.GoToHeader
        With Selection
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .TypeText Text:=YeMeiTitle
        End With
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub
Sub 更改变式格式()
    Call SASReplace("\[变式([0-9])\]", "变式\1", True)
    Call SASReplace("\[变式([0-9])－([0-9])\]", "变式\1.\2", True)
End Sub
Sub 取得题目名称()
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

Sub 删除空行()
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
    With Selection
        .EndKey Unit:=wdStory
        .InsertBreak Type:=wdPageBreak
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Size = 16
        .Font.Bold = wdToggle
        .Font.Name = "宋体"
        .TypeText Text:=AnswerTitle
        .TypeParagraph
        .Paste
    End With
End Sub

Sub 答案单独分页()
    For Each TempPar In ActiveDocument.Paragraphs
        If InStr(TempPar.Range, AnswerTitle) > 0 Then
            TempPar.Range.Select
            Selection.HomeKey Unit:=wdLine
            Selection.InsertBreak Type:=wdPageBreak
        End If
    Next
End Sub

Sub 分离答案通用版()
    Dim RemoveOn As Boolean
    Dim IsAnswer As Boolean
    Dim DuDianShuTong As Boolean
    Dim TiHao As String
    Dim n As Integer
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
            IsAnswer = True: RemoveOn = False
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
    Call SASReplace("【【", "", False)                              '去除“保留图片”标记
End Sub

Sub 规范标点()
    Call SASReplace("．", ".", False)
End Sub
Sub 删除题源通用版()
' 写此程序的原因在于有的资料上出现了类似于(2020XXX(下)XX)内含小括号的题源信息，根据具体的资料结构添加对应的题源形式，因此更加准确。
' 同时方便追加题源结构，拓展支持的格式，所以考虑将其列入通用程序
    Dim j, k, m As Integer
    Dim RepStr As String
' 借用正则表达式，将题源开头统一到特殊字符"DELETE"
    Call SASReplace("\([1-2][0-9][0-9][0-9]", "DELETE", True)
    Call SASReplace("\[([1-2][0-9][0-9][0-9]*)\]", "(\1)", True)
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
            Call SASReplace(RepStr, "", False)
        End If
    Next
End Sub

Sub 段落缩进(LeftNum, RightNum, FirstNum)
    With TempPar.Range.ParagraphFormat
        .CharacterUnitLeftIndent = LeftNum
        .CharacterUnitRightIndent = RightNum
        .CharacterUnitFirstLineIndent = FirstNum
    End With
End Sub
Sub 题目悬挂缩进()
    Dim TiHaoNum As Single
    For Each TempPar In ActiveDocument.Paragraphs
        If InStr(Mid(TempPar.Range, 1, 4), "考点") > 0 Then
             Call 段落缩进(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "知识点") > 0 Then
             Call 段落缩进(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "一 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "一、") > 0 Then
             Call 段落缩进(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "二 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "二、") > 0 Then
             Call 段落缩进(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "三 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "三、") > 0 Then
             Call 段落缩进(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "四 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "四、") > 0 Then
             Call 段落缩进(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "五 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "五、") > 0 Then
             Call 段落缩进(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "六 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "六、") > 0 Then
             Call 段落缩进(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "七 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "七、") > 0 Then
             Call 段落缩进(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "八 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "八、") > 0 Then
             Call 段落缩进(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "九 ") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "九、") > 0 Then
             Call 段落缩进(0, 0, -1)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "变式") > 0 Then
             Call 段落缩进(0, 0, -1): TiHaoNum = 1.3
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "例") > 0 Then
             Call 段落缩进(0, 0, -1): TiHaoNum = 1.3
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 1)) And InStr(Mid(TempPar.Range, 2, 1), ".") > 0 Then
             Call 段落缩进(0, 0, -1): TiHaoNum = 1.3
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 2)) And InStr(Mid(TempPar.Range, 3, 1), ".") > 0 Then
             Call 段落缩进(0, 0, -1.5): TiHaoNum = 1.5
        ElseIf IsNumeric(Mid(TempPar.Range, 1, 3)) And InStr(Mid(TempPar.Range, 4, 1), ".") > 0 Then
             Call 段落缩进(0, 0, -2): TiHaoNum = 1.7
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "A.") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "B.") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "C.") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "D.") > 0 Then
             Call 段落缩进(0, 0, TiHaoNum)
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "(1)") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "(2)") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "(3)") > 0 Or InStr(Mid(TempPar.Range, 1, 4), "(4)") > 0 Then
             Call 段落缩进(0, 0, -TiHaoNum)
             TempPar.Range.ParagraphFormat.CharacterUnitLeftIndent = TiHaoNum
        ElseIf InStr(Mid(TempPar.Range, 1, 4), "解析") > 0 Then
             Call 段落缩进(0, 0, -1)
        End If
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
