' ============================================================================
' 上海交通大学MBA学位论文样式一键导入宏
' 功能：自动创建并配置论文所需的所有字体样式、段落级别、自动编号
' 适用：Microsoft Word 2016/2019/2021 (Office 365)
' 版本：v3.0
' 使用方法：在Word中按 Alt+F11 打开VBA编辑器，导入此文件，运行主宏
' ============================================================================

Option Explicit


Sub 导入上海交大MBA论文样式()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Application.ScreenUpdating = False
    
    Call 设置页面格式(doc)
    Call 设置默认字体(doc)
    Call 创建正文样式(doc)
    Call 创建标题样式(doc)
    Call 创建摘要样式(doc)
    Call 创建目录样式(doc)
    Call 创建图表样式(doc)
    Call 创建参考文献样式(doc)
    Call 创建封面样式(doc)
    Call 创建页眉页脚样式(doc)
    Call 创建特殊样式(doc)
    Call 创建编号模板(doc)
    Call 强制修正样式属性(doc)
    
    Application.ScreenUpdating = True
    
    MsgBox "上海交通大学MBA学位论文样式导入完成！" & vbCrLf & vbCrLf & _
           "已创建内容：" & vbCrLf & _
           "  - 全部字体样式（含大纲级别）" & vbCrLf & _
           "  - 章/节/小节自动编号（第一章 → 1.1 → 1.1.1）" & vbCrLf & _
           "  - 分项编号（(1)(2)(3) / ①②③）" & vbCrLf & _
           "  - 图编号（图1、图2…全文连续）" & vbCrLf & _
           "  - 表编号（表1、表2…全文连续）" & vbCrLf & _
           "  - 公式编号（(1-1) 按章编排，右对齐）" & vbCrLf & _
           "  - 附录编号（附录1、附录2…连续）" & vbCrLf & _
           "  - 参考文献编号（[1] 全文顺序）" & vbCrLf & vbCrLf & _
           "公式使用：在MBA公式样式段落中按 Tab → 公式 → Tab → 编号", _
           vbInformation, "样式导入成功"

    If MsgBox("是否删除文档中与论文无关的其他样式？" & vbCrLf & vbCrLf & _
              "点击 ""是"" 将删除所有非MBA自定义样式和内置样式，仅保留MBA论文样式。" & vbCrLf & _
              "点击 ""否"" 保留所有现有样式。", _
              vbYesNo + vbQuestion, "清理无关样式") = vbYes Then
        Call 清除无关样式(doc)
        MsgBox "无关样式清理完成！仅保留了MBA论文相关样式。", vbInformation, "清理完成"
    End If
End Sub

Private Sub 设置页面格式(doc As Document)
    With doc.PageSetup
        .PaperSize = wdPaperA4
        .TopMargin = 28.35 * 3.5
        .BottomMargin = 28.35 * 4
        .LeftMargin = 28.35 * 2.8
        .RightMargin = 28.35 * 2.8
        .HeaderDistance = 28.35 * 2.5
        .FooterDistance = 28.35 * 3
        .Gutter = 0
        .MirrorMargins = False
    End With
End Sub

Private Sub 设置默认字体(doc As Document)
    With doc.Styles(wdStyleNormal).Font
        .Name = "Times New Roman"
        .NameFarEast = "宋体"
        .Size = 10.5
        .Bold = False
        .Italic = False
        .Color = wdColorBlack
    End With
    With doc.Styles(wdStyleNormal).ParagraphFormat
        .LineSpacingRule = wdLineSpace1pt5
        .Alignment = wdAlignParagraphJustify
        .FirstLineIndent = 21
        .SpaceBefore = 0
        .SpaceAfter = 0
        .OutlineLevel = wdOutlineLevelBodyText
    End With
End Sub

Private Sub 创建正文样式(doc As Document)
    Dim s As Style
    
    On Error Resume Next
    Set s = doc.Styles("MBA正文")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA正文", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 10.5
            .Bold = False
        End With
        With .ParagraphFormat
            .LineSpacingRule = wdLineSpace1pt5
            .Alignment = wdAlignParagraphJustify
            .FirstLineIndent = 21
            .SpaceBefore = 0
            .SpaceAfter = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA正文无缩进")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA正文无缩进", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 10.5
            .Bold = False
        End With
        With .ParagraphFormat
            .LineSpacingRule = wdLineSpace1pt5
            .Alignment = wdAlignParagraphJustify
            .FirstLineIndent = 0
            .SpaceBefore = 0
            .SpaceAfter = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA论文正文有缩进")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA论文正文有缩进", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 10.5
            .Bold = False
        End With
        With .ParagraphFormat
            .LineSpacingRule = wdLineSpace1pt5
            .Alignment = wdAlignParagraphJustify
            .FirstLineIndent = 21
            .SpaceBefore = 0
            .SpaceAfter = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
End Sub

Private Sub 创建标题样式(doc As Document)
    Dim s As Style
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA标题章")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA标题章", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "黑体"
            .Size = 16
            .Bold = True
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = 0
            .SpaceAfter = 32
            .FirstLineIndent = 0
            .PageBreakBefore = True
            .OutlineLevel = wdOutlineLevel1
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA标题节")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA标题节", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "黑体"
            .Size = 14
            .Bold = True
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 0
            .FirstLineIndent = 0
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevel2
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA标题小节")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA标题小节", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "黑体"
            .Size = 12
            .Bold = True
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 15.75
            .SpaceAfter = 0
            .FirstLineIndent = 0
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevel3
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA标题条目")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA标题条目", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 10.5
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 6
            .SpaceAfter = 3
            .FirstLineIndent = 0
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA项目编号")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA项目编号", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 10.5
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 3
            .SpaceAfter = 3
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
End Sub

Private Sub 创建摘要样式(doc As Document)
    Dim s As Style
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA中文摘要页题目")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA中文摘要页题目", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "黑体"
            .Size = 16
            .Bold = True
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 12
            .FirstLineIndent = 0
            .PageBreakBefore = True
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA中文摘要标题")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA中文摘要标题", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "黑体"
            .Size = 16
            .Bold = True
            .Spacing = 10.5
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 12
            .SpaceAfter = 12
            .FirstLineIndent = 0
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA中文摘要正文")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA中文摘要正文", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 14
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphJustify
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 0
            .FirstLineIndent = 21
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA中文关键词标签")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA中文关键词标签", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "黑体"
            .Size = 14
            .Bold = True
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 12
            .SpaceAfter = 0
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA中文关键词内容")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA中文关键词内容", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 14
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 0
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA英文摘要页题目")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA英文摘要页题目", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 16
            .Bold = True
            .AllCaps = True
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 12
            .FirstLineIndent = 0
            .PageBreakBefore = True
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA英文摘要标题")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA英文摘要标题", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 16
            .Bold = True
            .AllCaps = True
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 12
            .SpaceAfter = 12
            .FirstLineIndent = 0
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA英文摘要正文")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA英文摘要正文", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 14
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphJustify
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 0
            .FirstLineIndent = 36
            .LeftIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA英文关键词标签")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA英文关键词标签", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 14
            .Bold = True
            .AllCaps = True
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 12
            .SpaceAfter = 0
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA英文关键词内容")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA英文关键词内容", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 14
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 0
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
End Sub

Private Sub 创建目录样式(doc As Document)
    Dim s As Style
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA目录标题")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA目录标题", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "黑体"
            .Size = 16
            .Bold = True
            .Spacing = 10.5
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 12
            .SpaceAfter = 24
            .FirstLineIndent = 0
            .PageBreakBefore = True
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA目录一级")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA目录一级", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "黑体"
            .Size = 12
            .Bold = True
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 3
            .SpaceAfter = 3
            .FirstLineIndent = 0
            .TabStops.ClearAll
            .TabStops.Add Position:=28.35 * 15, Alignment:=wdAlignTabRight, Leader:=wdTabLeaderDots
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA目录二级")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA目录二级", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 12
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 3
            .SpaceAfter = 3
            .FirstLineIndent = 10.5
            .TabStops.ClearAll
            .TabStops.Add Position:=28.35 * 15, Alignment:=wdAlignTabRight, Leader:=wdTabLeaderDots
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA目录三级")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA目录三级", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 12
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 3
            .SpaceAfter = 3
            .FirstLineIndent = 21
            .TabStops.ClearAll
            .TabStops.Add Position:=28.35 * 15, Alignment:=wdAlignTabRight, Leader:=wdTabLeaderDots
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
End Sub

Private Sub 创建图表样式(doc As Document)
    Dim s As Style
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA图题")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA图题", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "楷体"
            .Size = 10.5
            .Bold = False
            .Spacing = 0
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 6
            .SpaceAfter = 0
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA表题")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA表题", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "楷体"
            .Size = 10.5
            .Bold = False
            .Spacing = 0
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 12
            .SpaceAfter = 0
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA英文图题")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA英文图题", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 10.5
            .Bold = False
            .Spacing = 0
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 12
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA英文表题")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA英文表题", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 10.5
            .Bold = False
            .Spacing = 0
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 6
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA图表来源")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA图表来源", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 9
            .Bold = False
            .Spacing = 0
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 0
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA公式")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA公式", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 10.5
            .Bold = False
            .Italic = True
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 6
            .SpaceAfter = 6
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
        .ParagraphFormat.TabStops.Add Position:=28.35 * 7.7, Alignment:=wdAlignTabCenter
        .ParagraphFormat.TabStops.Add Position:=28.35 * 15.4, Alignment:=wdAlignTabRight
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA公式编号")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA公式编号", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 10.5
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphRight
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 0
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
End Sub

Private Sub 创建参考文献样式(doc As Document)
    Dim s As Style
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA参考文献标题")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA参考文献标题", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "黑体"
            .Size = 16
            .Bold = True
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 24
            .SpaceAfter = 12
            .FirstLineIndent = 0
            .PageBreakBefore = True
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA参考文献正文")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA参考文献正文", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 10.5
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 0
            .FirstLineIndent = -21
            .LeftIndent = 21
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
End Sub

Private Sub 创建封面样式(doc As Document)
    Dim s As Style
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA封面题目")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA封面题目", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "黑体"
            .Size = 22
            .Bold = True
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 12
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA封面英文题目")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA封面英文题目", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 22
            .Bold = True
            .AllCaps = True
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 12
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA封面信息标签")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA封面信息标签", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "黑体"
            .Size = 16
            .Bold = True
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 6
            .SpaceAfter = 6
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA封面信息内容")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA封面信息内容", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "黑体"
            .Size = 16
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 6
            .SpaceAfter = 6
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
End Sub

Private Sub 创建页眉页脚样式(doc As Document)
    Dim s As Style
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA页眉")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA页眉", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 9
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = 0
            .SpaceAfter = 0
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA页脚")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA页脚", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 9
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = 0
            .SpaceAfter = 0
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
End Sub

Private Sub 创建特殊样式(doc As Document)
    Dim s As Style
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA致谢标题")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA致谢标题", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "黑体"
            .Size = 16
            .Bold = True
            .Spacing = 10.5
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 24
            .SpaceAfter = 12
            .FirstLineIndent = 0
            .PageBreakBefore = True
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA致谢正文")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA致谢正文", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 10.5
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphJustify
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 0
            .FirstLineIndent = 21
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA附录标题")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA附录标题", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "黑体"
            .Size = 16
            .Bold = True
            .Spacing = 10.5
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 24
            .SpaceAfter = 12
            .FirstLineIndent = 0
            .PageBreakBefore = True
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA附录正文")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA附录正文", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 10.5
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphJustify
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 0
            .FirstLineIndent = 21
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA附录编号")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA附录编号", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "黑体"
            .Size = 14
            .Bold = True
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 12
            .SpaceAfter = 12
            .FirstLineIndent = 0
            .PageBreakBefore = False
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA声明标题")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA声明标题", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "黑体"
            .Size = 16
            .Bold = True
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 24
            .SpaceAfter = 24
            .FirstLineIndent = 0
            .PageBreakBefore = True
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA声明正文")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA声明正文", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 10.5
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphJustify
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 0
            .FirstLineIndent = 21
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
    
    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA脚注")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA脚注", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "宋体"
            .Size = 9
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = 0
            .SpaceAfter = 0
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With


    Set s = Nothing
    On Error Resume Next
    Set s = doc.Styles("MBA续表")
    On Error GoTo 0
    If s Is Nothing Then Set s = doc.Styles.Add(Name:="MBA续表", Type:=wdStyleTypeParagraph)
    With s
        With .Font
            .Name = "Times New Roman"
            .NameFarEast = "楷体"
            .Size = 10.5
            .Bold = False
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 12
            .SpaceAfter = 0
            .FirstLineIndent = 0
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With
End Sub

Private Sub 创建编号模板(doc As Document)
    Dim lt As ListTemplate
    Dim i As Long
    Dim s As Style
    
    Call 解绑非引用编号样式(doc)
    
    Set lt = doc.ListTemplates.Add(OutlineNumbered:=True)
    
    With lt.ListLevels(1)
        .NumberFormat = ChrW(31532) & " " & "%1" & " " & ChrW(31456)
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignCenter
        .TextPosition = 0
        .ResetOnHigher = 0
        .StartAt = 1
        .Font.Name = "Times New Roman"
        .Font.NameFarEast = "黑体"
        .Font.Size = 16
        .Font.Bold = True
    End With
    
    With lt.ListLevels(2)
        .NumberFormat = "%1.%2"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 0
        .ResetOnHigher = 1
        .StartAt = 1
        .Font.Name = "Times New Roman"
        .Font.NameFarEast = "黑体"
        .Font.Size = 14
        .Font.Bold = True
    End With
    
    With lt.ListLevels(3)
        .NumberFormat = "%1.%2.%3"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 0
        .ResetOnHigher = 2
        .StartAt = 1
        .Font.Name = "Times New Roman"
        .Font.NameFarEast = "黑体"
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    With lt.ListLevels(4)
        .NumberFormat = ChrW(65288) & "%4" & ChrW(65289)
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 21
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 42
        .ResetOnHigher = 3
        .StartAt = 1
        .Font.Name = "Times New Roman"
        .Font.NameFarEast = "宋体"
        .Font.Size = 10.5
        .Font.Bold = False
    End With
    
    With lt.ListLevels(5)
        .NumberFormat = "%5"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleNumberInCircle
        .NumberPosition = 21
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 42
        .ResetOnHigher = 4
        .StartAt = 1
        .Font.Name = "Times New Roman"
        .Font.NameFarEast = "宋体"
        .Font.Size = 10.5
    End With
    
    With lt.ListLevels(6)
        .NumberFormat = ChrW(22270) & "%1-" & "%6"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignCenter
        .TextPosition = 0
        .ResetOnHigher = 1
        .StartAt = 1
        .Font.Name = "Times New Roman"
        .Font.NameFarEast = "楷体"
        .Font.Size = 10.5
    End With
    
    With lt.ListLevels(7)
        .NumberFormat = ChrW(34920) & "%1-" & "%7"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignCenter
        .TextPosition = 0
        .ResetOnHigher = 1
        .StartAt = 1
        .Font.Name = "Times New Roman"
        .Font.NameFarEast = "楷体"
        .Font.Size = 10.5
    End With
    
    With lt.ListLevels(8)
        .NumberFormat = "(%1-%8)"
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignRight
        .TextPosition = 0
        .ResetOnHigher = 1
        .StartAt = 1
        .Font.Name = "Times New Roman"
        .Font.Size = 10.5
    End With
    
    With lt.ListLevels(9)
        .NumberFormat = ChrW(38468) & ChrW(24405) & "%9"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignCenter
        .TextPosition = 0
        .ResetOnHigher = 0
        .StartAt = 1
        .Font.Name = "Times New Roman"
        .Font.NameFarEast = "黑体"
        .Font.Size = 14
        .Font.Bold = True
    End With
    
    Call 绑定所有编号(doc, lt)
    
    Dim ltR As ListTemplate
    On Error Resume Next
    Set s = doc.Styles("MBA参考文献正文")
    If Not s Is Nothing Then Set ltR = s.ListTemplate
    On Error GoTo 0
    
    If ltR Is Nothing Then Set ltR = doc.ListTemplates.Add(OutlineNumbered:=True)
    
    With ltR.ListLevels(1)
        .NumberFormat = "[%1]"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 21
        .ResetOnHigher = 0
        .StartAt = 1
        .Font.Name = "Times New Roman"
        .Font.NameFarEast = "宋体"
        .Font.Size = 10.5
        .Font.Bold = False
    End With
    For i = 2 To 9
        With ltR.ListLevels(i)
            .NumberFormat = ""
            .NumberStyle = wdListNumberStyleArabic
            .ResetOnHigher = 0
        End With
    Next i
    
    On Error Resume Next
    Set s = doc.Styles("MBA参考文献正文")
    If Not s Is Nothing Then s.LinkToListTemplate ListTemplate:=ltR, ListLevelNumber:=1
    On Error GoTo 0
    
    Dim ltE As ListTemplate
    Set ltE = doc.ListTemplates.Add(OutlineNumbered:=True)
    
    With ltE.ListLevels(1)
        .NumberFormat = "%1"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 0
        .ResetOnHigher = 0
        .StartAt = 1
        .Font.Name = "Times New Roman"
        .Font.Size = 10.5
    End With
    
    With ltE.ListLevels(2)
        .NumberFormat = "Figure " & "%1-" & "%2"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 0
        .ResetOnHigher = 1
        .StartAt = 1
        .Font.Name = "Times New Roman"
        .Font.NameFarEast = "宋体"
        .Font.Size = 10.5
        .Font.Bold = False
    End With
    For i = 3 To 9
        With ltE.ListLevels(i)
            .NumberFormat = ""
            .NumberStyle = wdListNumberStyleArabic
            .ResetOnHigher = 0
        End With
    Next i
    
    On Error Resume Next
    Set s = doc.Styles("MBA英文图题")
    If Not s Is Nothing Then s.LinkToListTemplate ListTemplate:=ltE, ListLevelNumber:=2
    On Error GoTo 0
    
    Dim ltE2 As ListTemplate
    Set ltE2 = doc.ListTemplates.Add(OutlineNumbered:=True)
    
    With ltE2.ListLevels(1)
        .NumberFormat = "%1"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 0
        .ResetOnHigher = 0
        .StartAt = 1
        .Font.Name = "Times New Roman"
        .Font.Size = 10.5
    End With
    
    With ltE2.ListLevels(2)
        .NumberFormat = "Table " & "%1-" & "%2"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = 0
        .Alignment = wdListLevelAlignLeft
        .TextPosition = 0
        .ResetOnHigher = 1
        .StartAt = 1
        .Font.Name = "Times New Roman"
        .Font.NameFarEast = "宋体"
        .Font.Size = 10.5
        .Font.Bold = False
    End With
    For i = 3 To 9
        With ltE2.ListLevels(i)
            .NumberFormat = ""
            .NumberStyle = wdListNumberStyleArabic
            .ResetOnHigher = 0
        End With
    Next i
    
    On Error Resume Next
    Set s = doc.Styles("MBA英文表题")
    If Not s Is Nothing Then s.LinkToListTemplate ListTemplate:=ltE2, ListLevelNumber:=2
    On Error GoTo 0
End Sub

Private Sub 解绑非引用编号样式(doc As Document)
    Dim s As Style
    Dim n As Variant
    On Error Resume Next
    For Each n In Array("MBA标题章", "MBA标题节", "MBA标题小节", _
                        "MBA标题条目", "MBA项目编号", _
                        "MBA图题", "MBA英文图题", _
                        "MBA表题", "MBA英文表题", _
                        "MBA公式编号", "MBA附录编号")
        Set s = doc.Styles(CStr(n))
        If Not s Is Nothing Then s.LinkToListTemplate ListTemplate:=Nothing
    Next n
    On Error GoTo 0
End Sub

Private Sub 绑定所有编号(doc As Document, lt As ListTemplate)
    Dim s As Style
    On Error Resume Next
    
    Set s = doc.Styles("MBA标题章")
    If Not s Is Nothing Then s.LinkToListTemplate ListTemplate:=lt, ListLevelNumber:=1
    
    Set s = doc.Styles("MBA标题节")
    If Not s Is Nothing Then s.LinkToListTemplate ListTemplate:=lt, ListLevelNumber:=2
    
    Set s = doc.Styles("MBA标题小节")
    If Not s Is Nothing Then s.LinkToListTemplate ListTemplate:=lt, ListLevelNumber:=3
    
    Set s = doc.Styles("MBA标题条目")
    If Not s Is Nothing Then s.LinkToListTemplate ListTemplate:=lt, ListLevelNumber:=4
    
    Set s = doc.Styles("MBA项目编号")
    If Not s Is Nothing Then s.LinkToListTemplate ListTemplate:=lt, ListLevelNumber:=5
    
    Set s = doc.Styles("MBA图题")
    If Not s Is Nothing Then s.LinkToListTemplate ListTemplate:=lt, ListLevelNumber:=6
    
    Set s = doc.Styles("MBA表题")
    If Not s Is Nothing Then s.LinkToListTemplate ListTemplate:=lt, ListLevelNumber:=7
    
    Set s = doc.Styles("MBA公式编号")
    If Not s Is Nothing Then s.LinkToListTemplate ListTemplate:=lt, ListLevelNumber:=8
    
    Set s = doc.Styles("MBA附录编号")
    If Not s Is Nothing Then s.LinkToListTemplate ListTemplate:=lt, ListLevelNumber:=9
    
    On Error GoTo 0
End Sub


Private Sub 强制修正样式属性(doc As Document)
    Dim s As Style
    Dim n As Variant
    
    On Error Resume Next
    
    For Each n In Array("MBA正文", "MBA正文无缩进", "MBA论文正文有缩进", _
                        "MBA中文摘要页题目", "MBA中文摘要标题", _
                        "MBA中文摘要正文", "MBA中文关键词标签", "MBA中文关键词内容", _
                        "MBA英文摘要页题目", "MBA英文摘要标题", "MBA英文摘要正文", _
                        "MBA英文关键词标签", "MBA英文关键词内容", _
                        "MBA目录标题", "MBA目录一级", "MBA目录二级", "MBA目录三级", _
                        "MBA图题", "MBA英文图题", "MBA表题", "MBA英文表题", _
                        "MBA图表来源", "MBA公式", "MBA公式编号", _
                        "MBA参考文献标题", "MBA参考文献正文", _
                        "MBA封面题目", "MBA封面英文题目", "MBA封面信息标签", "MBA封面信息内容", _
                        "MBA页眉", "MBA页脚", _
                        "MBA致谢标题", "MBA致谢正文", _
                        "MBA附录标题", "MBA附录编号", "MBA附录正文", _
                        "MBA声明标题", "MBA声明正文", "MBA脚注", _
                        "MBA项目编号")
        Set s = Nothing
        Set s = doc.Styles(CStr(n))
        If Not s Is Nothing Then
            s.Font.Spacing = 0
            s.Font.Color = wdColorBlack
        End If
    Next n
    
    For Each n In Array("MBA正文", "MBA正文无缩进", "MBA论文正文有缩进", _
                        "MBA中文摘要正文", "MBA中文关键词内容", _
                        "MBA目录二级", "MBA目录三级", _
                        "MBA图表来源", "MBA公式", "MBA公式编号", _
                        "MBA参考文献正文", _
                        "MBA致谢正文", _
                        "MBA附录正文", _
                        "MBA声明正文", _
                        "MBA页眉", "MBA页脚", "MBA脚注", "MBA项目编号", _
                        "MBA英文图题", "MBA英文表题", _
                        "MBA英文摘要标题", "MBA英文摘要页题目", "MBA英文摘要正文", _
                        "MBA英文关键词标签", "MBA英文关键词内容", _
                        "MBA封面英文题目")
        Set s = Nothing
        Set s = doc.Styles(CStr(n))
        If Not s Is Nothing Then
            s.Font.Name = "Times New Roman"
            s.Font.NameFarEast = "宋体"
        End If
    Next n
    
    For Each n In Array("MBA标题章", "MBA标题节", "MBA标题小节", _
                        "MBA中文摘要页题目", "MBA中文摘要标题", "MBA中文关键词标签", _
                        "MBA目录标题", "MBA目录一级", "MBA参考文献标题", _
                        "MBA封面题目", "MBA封面信息标签", "MBA封面信息内容", _
                        "MBA致谢标题", "MBA附录标题", "MBA附录编号", _
                        "MBA声明标题")
        Set s = Nothing
        Set s = doc.Styles(CStr(n))
        If Not s Is Nothing Then
            s.Font.Name = "Times New Roman"
            s.Font.NameFarEast = "黑体"
        End If
    Next n
    
    Set s = doc.Styles("MBA致谢正文")
    If Not s Is Nothing Then s.Font.Size = 10.5
    
    Set s = doc.Styles("MBA声明正文")
    If Not s Is Nothing Then s.Font.Size = 10.5
    
    Set s = doc.Styles("MBA英文摘要正文")
    If Not s Is Nothing Then s.ParagraphFormat.FirstLineIndent = 36
    
    Set s = doc.Styles("MBA图题")
    If Not s Is Nothing Then
        s.Font.Name = "Times New Roman"
        s.Font.NameFarEast = "楷体"
    End If
    
    Set s = doc.Styles("MBA表题")
    If Not s Is Nothing Then
        s.Font.Name = "Times New Roman"
        s.Font.NameFarEast = "楷体"
    End If
    
    Set s = doc.Styles("MBA标题条目")
    If Not s Is Nothing Then
        s.Font.Name = "Times New Roman"
        s.Font.NameFarEast = "宋体"
        s.Font.Bold = False
    End If
    
    On Error GoTo 0
End Sub

Sub 自动应用论文样式()
    Dim para As Paragraph
    Dim txt As String
    Dim cn As Long
    
    Application.ScreenUpdating = False
    cn = 0
    
    For Each para In ActiveDocument.Paragraphs
        txt = Replace(para.Range.Text, Chr(13), "")
        txt = Replace(txt, Chr(10), "")
        txt = Trim(txt)
        If Len(txt) = 0 Then GoTo NP
        
        If txt Like "第?章 *" Or txt Like "第??章 *" Or _
           txt Like "第?章" Or txt Like "第??章" Or _
           txt Like "第?章*" Or txt Like "第??章*" Then
            cn = cn + 1
            para.Style = ActiveDocument.Styles("MBA标题章")
            
        ElseIf txt Like "#.# *" Or txt Like "##.# *" Or _
               txt Like "#.#" Or txt Like "##.#" Then
            If cn > 0 Then para.Style = ActiveDocument.Styles("MBA标题节")
            
        ElseIf txt Like "#.#.# *" Or txt Like "##.#.# *" Or _
               txt Like "#.#.#" Or txt Like "##.#.#" Then
            If cn > 0 Then para.Style = ActiveDocument.Styles("MBA标题小节")
            
        ElseIf txt = "摘要" Then
            para.Style = ActiveDocument.Styles("MBA中文摘要标题")
            
        ElseIf UCase(txt) = "ABSTRACT" Then
            para.Style = ActiveDocument.Styles("MBA英文摘要标题")
            
        ElseIf txt = "目录" Then
            para.Style = ActiveDocument.Styles("MBA目录标题")
            
        ElseIf txt = "参考文献" Then
            para.Style = ActiveDocument.Styles("MBA参考文献标题")
            
        ElseIf txt = "致谢" Then
            para.Style = ActiveDocument.Styles("MBA致谢标题")
            
        ElseIf txt Like "*附录[0-9]*" Or txt Like "*附录1*" Or txt Like "*附录2*" Then
            para.Style = ActiveDocument.Styles("MBA附录编号")
            
        ElseIf txt Like "*图[0-9]*" And Len(txt) < 50 Then
            para.Style = ActiveDocument.Styles("MBA图题")
            
        ElseIf txt Like "*表[0-9]*" And Len(txt) < 50 Then
            para.Style = ActiveDocument.Styles("MBA表题")
            
        ElseIf txt Like "关键词*" Then
            para.Style = ActiveDocument.Styles("MBA中文关键词标签")
            
        ElseIf UCase(txt) Like "KEY WORDS*" Or UCase(txt) Like "KEYWORDS*" Then
            para.Style = ActiveDocument.Styles("MBA英文关键词标签")
            
        ElseIf txt Like "*原创性声明*" Or txt Like "*版权使用授权书*" Then
            para.Style = ActiveDocument.Styles("MBA声明标题")
        End If
        
NP:
    Next para
    
    Application.ScreenUpdating = True
    MsgBox "自动应用样式完成！" & vbCrLf & "已自动识别并应用章、节、小节等标题样式。", vbInformation, "完成"
End Sub

Sub 列出所有MBA样式()
    Dim sty As Style
    Dim doc As Document
    Dim nd As Document
    Dim ol As String
    Dim al As String
    
    Set doc = ActiveDocument
    Set nd = Documents.Add
    
    nd.Content.Text = "上海交通大学MBA学位论文样式清单" & vbCrLf & _
                      "=====================================" & vbCrLf & vbCrLf
    
    For Each sty In doc.Styles
        If Left(sty.NameLocal, 3) = "MBA" Then
            Select Case sty.ParagraphFormat.OutlineLevel
                Case wdOutlineLevel1: ol = "大纲级别1级"
                Case wdOutlineLevel2: ol = "大纲级别2级"
                Case wdOutlineLevel3: ol = "大纲级别3级"
                Case wdOutlineLevel4: ol = "大纲级别4级"
                Case Else: ol = "正文级别"
            End Select
            
            nd.Content.InsertAfter "样式名称: " & sty.NameLocal & vbCrLf
            nd.Content.InsertAfter "大纲级别: " & ol & vbCrLf
            nd.Content.InsertAfter "西文字体: " & sty.Font.Name & vbCrLf
            nd.Content.InsertAfter "中文字体: " & sty.Font.NameFarEast & vbCrLf
            nd.Content.InsertAfter "字号: " & sty.Font.Size & "磅" & vbCrLf
            nd.Content.InsertAfter "加粗: " & IIf(sty.Font.Bold, "是", "否") & vbCrLf
            Select Case sty.ParagraphFormat.Alignment
                Case wdAlignParagraphLeft: al = "左对齐"
                Case wdAlignParagraphCenter: al = "居中"
                Case wdAlignParagraphRight: al = "右对齐"
                Case wdAlignParagraphJustify: al = "两端对齐"
                Case Else: al = "其他"
            End Select
            nd.Content.InsertAfter "对齐: " & al & vbCrLf
            nd.Content.InsertAfter "首行缩进: " & sty.ParagraphFormat.FirstLineIndent & "磅" & vbCrLf
            nd.Content.InsertAfter "行距: " & _
                IIf(sty.ParagraphFormat.LineSpacingRule = wdLineSpace1pt5, "1.5倍", _
                IIf(sty.ParagraphFormat.LineSpacingRule = wdLineSpaceDouble, "双倍", "其他")) & vbCrLf
            nd.Content.InsertAfter "段前: " & sty.ParagraphFormat.SpaceBefore & "磅" & vbCrLf
            nd.Content.InsertAfter "段后: " & sty.ParagraphFormat.SpaceAfter & "磅" & vbCrLf
            nd.Content.InsertAfter "----------------------------------------" & vbCrLf & vbCrLf
        End If
    Next sty
    
    MsgBox "样式清单已生成到新文档！", vbInformation, "完成"
End Sub

Sub 清除所有MBA样式()
    Dim sty As Style
    Dim ct As Long
    
    If MsgBox("确定要删除所有MBA样式吗？此操作不可撤销。", vbYesNo + vbExclamation, "确认删除") = vbNo Then
        Exit Sub
    End If
    
    ct = 0
    For Each sty In ActiveDocument.Styles
        If Left(sty.NameLocal, 3) = "MBA" Then
            On Error Resume Next
            sty.Delete
            On Error GoTo 0
            ct = ct + 1
        End If
    Next sty
    
    MsgBox "已删除 " & ct & " 个MBA样式。", vbInformation, "完成"
End Sub

Private Sub 清除无关样式(doc As Document)
    Dim sty As Style
    Dim styName As String
    Dim totalBefore As Long
    Dim totalAfter As Long
    Dim delCount As Long
    Dim keepCount As Long
    Dim msg As String
    
    On Error Resume Next
    
    totalBefore = doc.Styles.Count
    delCount = 0
    keepCount = 0
    
    For Each sty In doc.Styles
        styName = sty.NameLocal
        
        If sty.BuiltIn Then GoTo NextSty
        If Left(styName, 3) = "MBA" Then GoTo NextSty
        
        If sty.InUse Then
            keepCount = keepCount + 1
            GoTo NextSty
        End If
        
        Err.Clear
        sty.Delete
        If Err.Number = 0 Then
            delCount = delCount + 1
        End If
        Err.Clear
        
NextSty:
    Next sty
    
    totalAfter = doc.Styles.Count
    On Error GoTo 0
    
    msg = "样式清理完成！" & vbCrLf & _
          "清理前样式总数: " & totalBefore & vbCrLf & _
          "清理后样式总数: " & totalAfter & vbCrLf & _
          "已删除空闲样式: " & delCount & vbCrLf & _
          "保留在用样式: " & keepCount
    MsgBox msg, vbInformation, "清理结果"
End Sub
