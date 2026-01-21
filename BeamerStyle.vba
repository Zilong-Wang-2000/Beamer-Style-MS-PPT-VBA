' ===================================================================
' Beamer Like PowerPoint 自动化格式宏
' 功能:
' 1. AddProgressBar: 在幻灯片【最下方】添加一个动态进度条。
' 2. AddSectionNamesToHeader: 在顶部创建完整的 Beamer 风格导航。
' 3. UpdatePageFormat: 统一右下角的页码格式为 "当前页 / 总页数"。
' 4. RunAllFunctions: 全部运行。
'
' 作者: ZilongWang
' 日期: 2025-10-14
' 兼容性: MacOS
' ===================================================================

' Subroutine 1: 添加底部进度条
Sub AddProgressBar()
    Dim X As Long
    Dim S As Shape
    Dim slideHeight As Single
    On Error Resume Next
    With ActivePresentation
        slideHeight = .PageSetup.SlideHeight
        For X = 3 To .Slides.Count - 1
            On Error Resume Next
            Do
                .Slides(X).Shapes("PB").Delete
            Loop Until .Slides(X).Shapes("PB") Is Nothing
            Do
                .Slides(X).Shapes("PC").Delete
            Loop Until .Slides(X).Shapes("PC") Is Nothing
            On Error GoTo 0
            Set S = .Slides(X).Shapes.AddLine(-1, slideHeight - 2, .PageSetup.SlideWidth + 1, slideHeight - 2)
            S.Line.Weight = 3
            S.Line.ForeColor.RGB = RGB(205, 205, 205)
            S.Name = "PB"
            Set S = .Slides(X).Shapes.AddLine(-1, slideHeight - 2, (X - 2) * .PageSetup.SlideWidth / (.Slides.Count - 3) + 1, slideHeight - 2)
            S.Line.Weight = 3
            S.Line.ForeColor.RGB = RGB(50, 100, 200)
            S.Name = "PC"
        Next X
    End With
End Sub

' Subroutine 2: 创建完整的顶部导航系统 (高亮圆圈已增大)
Sub AddSectionNamesToHeader()
    Dim sld As slide
    Dim headerShape As Shape, circleShape As Shape, sepShape As Shape
    Dim i As Long, j As Long
    
    ' --- 数据准备 ---
    Dim sectionNames As New Collection
    Dim sectionStartSlides As New Collection
    Dim sectionSlideCounts As New Collection
    
    For i = 1 To ActivePresentation.SectionProperties.Count
        sectionNames.Add ActivePresentation.SectionProperties.Name(i)
        sectionStartSlides.Add ActivePresentation.SectionProperties.FirstSlide(i), ActivePresentation.SectionProperties.Name(i)
        sectionSlideCounts.Add ActivePresentation.SectionProperties.SlidesCount(i), ActivePresentation.SectionProperties.Name(i)
    Next i
    
    ' --- 筛选章节 (移除目录、首页和尾页) ---
    For i = sectionNames.Count To 1 Step -1
        If StrComp(sectionNames(i), "目录", vbTextCompare) = 0 Then
            sectionNames.Remove i
        End If
    Next i
    If sectionNames.Count > 0 Then sectionNames.Remove 1
    If sectionNames.Count > 0 Then sectionNames.Remove sectionNames.Count

    ' --- 遍历每张幻灯片，绘制导航元素 ---
    For Each sld In ActivePresentation.Slides
        For i = sld.Shapes.Count To 1 Step -1
            If Left(sld.Shapes(i).Name, 17) = "HeaderSectionName" Or _
               Left(sld.Shapes(i).Name, 15) = "HeaderSeparator" Or _
               Left(sld.Shapes(i).Name, 17) = "BeamerSlideCircle" Then
                sld.Shapes(i).Delete
            End If
        Next i

        Dim currentSectionName As String
        currentSectionName = ActivePresentation.SectionProperties.Name(sld.sectionIndex)
        
        If sectionNames.Count > 0 Then
            Dim portion As Single
            portion = ActivePresentation.PageSetup.SlideWidth / sectionNames.Count

            For i = 1 To sectionNames.Count
                Dim sectionTitle As String
                sectionTitle = sectionNames(i)
                
                ' 1. 绘制章节标题
                Set headerShape = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, (i - 1) * portion, 0, portion, 12)
                headerShape.Name = "HeaderSectionName" & i
                With headerShape.TextFrame.TextRange
                    .Text = sectionTitle
                    .Font.Size = 9
                    .Font.NameFarEast = "黑体"
                    .Font.Name = "Times New Roman"
                    .ParagraphFormat.Alignment = ppAlignCenter
                    .Font.Color.RGB = RGB(205, 205, 205)
                End With

                If sectionTitle = currentSectionName Then
                    headerShape.TextFrame.TextRange.Font.Bold = msoTrue
                    headerShape.TextFrame.TextRange.Font.Color.RGB = RGB(240, 180, 50)
                End If
                
                Dim sectionStartIndex As Long
                sectionStartIndex = sectionStartSlides(sectionTitle)
                With headerShape.ActionSettings(ppMouseClick)
                    .Action = ppActionHyperlink
                    .Hyperlink.SubAddress = ActivePresentation.Slides(sectionStartIndex).SlideID & "," & sectionStartIndex & "," & ActivePresentation.Slides(sectionStartIndex).Name
                End With

                ' 2. 在标题下方绘制本章节的导航圆圈
                Dim slidesInSection As Long, circlesTotalWidth As Single, circlesStartLeft As Single
                Dim circleDiameter As Single, highlightedCircleDiameter As Single, circleSpacing As Single, verticalPos As Single
                
                slidesInSection = sectionSlideCounts(sectionTitle)
                circleDiameter = 5 ' 普通圆圈的直径
                highlightedCircleDiameter = 7 ' 高亮圆圈的直径
                circleSpacing = 4
                verticalPos = 16
                
                circlesTotalWidth = (slidesInSection * circleDiameter) + ((slidesInSection - 1) * circleSpacing)
                circlesStartLeft = ((i - 1) * portion) + (portion - circlesTotalWidth) / 2
                
                For j = 1 To slidesInSection
                    Dim targetSlideIndex As Long, currentDiameter As Single, sizeAdjust As Single
                    Dim drawLeft As Single, drawTop As Single
                    targetSlideIndex = sectionStartIndex + j - 1
                    
                    ' 判断使用哪个直径尺寸
                    If sld.SlideIndex = targetSlideIndex Then
                        currentDiameter = highlightedCircleDiameter
                    Else
                        currentDiameter = circleDiameter
                    End If
                    
                    ' 计算尺寸和位置的微调量，以保持视觉居中
                    sizeAdjust = (currentDiameter - circleDiameter) / 2
                    drawLeft = circlesStartLeft + ((j - 1) * (circleDiameter + circleSpacing)) - sizeAdjust
                    drawTop = verticalPos - sizeAdjust
                    
                    ' 用计算好的尺寸和位置绘制圆圈
                    Set circleShape = sld.Shapes.AddShape(msoShapeOval, drawLeft, drawTop, currentDiameter, currentDiameter)
                    circleShape.Name = "BeamerSlideCircle" & i & "_" & j

                    With circleShape.ActionSettings(ppMouseClick)
                        .Action = ppActionHyperlink
                        .Hyperlink.SubAddress = ActivePresentation.Slides(targetSlideIndex).SlideID & "," & targetSlideIndex & "," & ActivePresentation.Slides(targetSlideIndex).Name
                    End With
                    
                    If sld.SlideIndex = targetSlideIndex Then
                        circleShape.Fill.Visible = msoTrue
                        circleShape.Fill.ForeColor.RGB = RGB(180, 180, 180)
                        circleShape.Line.Visible = msoFalse
                    Else
                        circleShape.Fill.Visible = msoFalse
                        circleShape.Line.Visible = msoTrue
                        circleShape.Line.ForeColor.RGB = RGB(205, 205, 205)
                        circleShape.Line.Weight = 1
                    End If
                Next j

                ' 3. 绘制分隔符
                If i < sectionNames.Count Then
                    Set sepShape = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, i * portion - 10, 6, 20, 12)
                    sepShape.Name = "HeaderSeparator" & i
                    With sepShape.TextFrame.TextRange
                        .Text = "|"
                        .Font.Size = 10
                        .Font.NameFarEast = "黑体"
                        .Font.Name = "Times New Roman"
                        .ParagraphFormat.Alignment = ppAlignCenter
                        .Font.Color.RGB = RGB(205, 205, 205)
                    End With
                End If
            Next i
        End If
    Next sld
End Sub


' Subroutine 3: 更新页码格式
Sub UpdatePageFormat()
    Dim sld As slide
    Dim shp As Shape
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Type = msoPlaceholder Then
                If shp.PlaceholderFormat.Type = ppPlaceholderSlideNumber Then
                    With shp
                        .TextFrame.TextRange.Text = sld.SlideIndex & " / " & ActivePresentation.Slides.Count
                        .Width = 60
                        .Left = ActivePresentation.PageSetup.SlideWidth - .Width
                        With .TextFrame.TextRange.Font
                            .NameFarEast = "黑体"
                            .Name = "Times New Roman"
                            .Size = 14
                            .Color.RGB = RGB(25, 25, 25)
                        End With
                    End With
                End If
            End If
        Next shp
    Next sld
End Sub

' 主控宏: 运行所有格式化功能
Sub RunAllFunctions()
    AddProgressBar
    AddSectionNamesToHeader 
    UpdatePageFormat
End Sub
