Attribute VB_Name = "������ģ��"
''ĿǰΪ���Թ��ܣ���ʱ������ϵͳ
'
'
Sub ��ʽ��ͼƬ()
    Call �ƶ���ǶͼƬ
    Call ת��Ϊ������
    Call ����ͼƬ�Ҷ���
'    Call �����ע
'    Call ������ע��ʽ
    Call ɾ������
End Sub

Sub �ƶ���ǶͼƬ()
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
Sub ת��Ϊ������()
    For n = 1 To ActiveDocument.InlineShapes.Count
        ActiveDocument.InlineShapes(1).ConvertToShape.WrapFormat.Type = 0
    Next
End Sub

Sub ����ͼƬ�Ҷ���()
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
Sub �����ע()
    For n = 1 To ActiveDocument.Shapes.Count
        ActiveDocument.Shapes.Range(n).Select
        Selection.InsertCaption Label:="ͼ", TitleAutoText:="InsertCaption3", Title _
        :="", Position:=wdCaptionPositionBelow, ExcludeLabel:=0
    Next
End Sub
Sub ������ע��ʽ()
'
    With ActiveDocument.Styles("��ע").ParagraphFormat
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
    ActiveDocument.Styles("��ע").NoSpaceBetweenParagraphsOfSameStyle = False
    With ActiveDocument.Styles("��ע")
        .AutomaticallyUpdate = False
        .BaseStyle = "����"
        .NextParagraphStyle = "����"
    End With
End Sub

