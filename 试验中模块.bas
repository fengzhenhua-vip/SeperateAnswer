Attribute VB_Name = "试验中模块"
''目前为测试功能，暂时不加入系统
'
'
Sub 格式化图片()
    Call 移动内嵌图片
    Call 转换为四周型
    Call 设置图片右对齐
'    Call 添加题注
'    Call 设置题注格式
    Call 删除空行
End Sub

Sub 移动内嵌图片()
'
    Dim TempPic As InlineShapes
    For n = 1 To ActiveDocument.InlineShapes.Count
        ActiveDocument.InlineShapes(n).Select
        Selection.Cut
        Selection.MoveUp Unit:=wdParagraph, Count:=1
        Selection.MoveRight Unit:=wdCharacter, Count:=10
        Selection.PasteAndFormat (wdPasteDefault)
    Next

End Sub
Sub 转换为四周型()
    For n = 1 To ActiveDocument.InlineShapes.Count
        ActiveDocument.InlineShapes(1).ConvertToShape.WrapFormat.Type = 0
    Next
End Sub

Sub 设置图片右对齐()
    Dim oShape As Variant
    For Each oShape In ActiveDocument.Shapes
        oShape.Select
        Selection.ShapeRange.RelativeHorizontalPosition = _
            wdRelativeHorizontalPositionMargin
        Selection.ShapeRange.RelativeVerticalPosition = _
            wdRelativeVerticalPositionLine
        Selection.ShapeRange.RelativeHorizontalSize = _
            wdRelativeHorizontalSizeMargin
        Selection.ShapeRange.RelativeVerticalSize = wdRelativeVerticalSizeMargin
        Selection.ShapeRange.Left = wdShapeRight
        Selection.ShapeRange.LeftRelative = wdShapePositionRelativeNone
        Selection.ShapeRange.Top = CentimetersToPoints(0.15)
        Selection.ShapeRange.TopRelative = wdShapePositionRelativeNone
        Selection.ShapeRange.WidthRelative = wdShapeSizeRelativeNone
        Selection.ShapeRange.HeightRelative = wdShapeSizeRelativeNone
        Selection.ShapeRange.LockAnchor = False
        Selection.ShapeRange.LayoutInCell = True
        Selection.ShapeRange.WrapFormat.AllowOverlap = True
        Selection.ShapeRange.WrapFormat.Side = wdWrapLeft
        Selection.ShapeRange.WrapFormat.DistanceTop = CentimetersToPoints(0)
        Selection.ShapeRange.WrapFormat.DistanceBottom = CentimetersToPoints(0)
        Selection.ShapeRange.WrapFormat.DistanceLeft = CentimetersToPoints(0.32)
        Selection.ShapeRange.WrapFormat.DistanceRight = CentimetersToPoints(0.32)
        Selection.ShapeRange.WrapFormat.Type = wdWrapSquare
    Next
End Sub
Sub 添加题注()
    For n = 1 To ActiveDocument.Shapes.Count
        ActiveDocument.Shapes.Range(n).Select
        Selection.InsertCaption Label:="图", TitleAutoText:="InsertCaption3", Title _
        :="", Position:=wdCaptionPositionBelow, ExcludeLabel:=0
    Next
End Sub
Sub 设置题注格式()
'
    With ActiveDocument.Styles("题注").ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 13.5
        .Alignment = wdAlignParagraphCenter
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .CollapsedByDefault = False
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
    ActiveDocument.Styles("题注").NoSpaceBetweenParagraphsOfSameStyle = False
    With ActiveDocument.Styles("题注")
        .AutomaticallyUpdate = False
        .BaseStyle = "正文"
        .NextParagraphStyle = "正文"
    End With
End Sub

