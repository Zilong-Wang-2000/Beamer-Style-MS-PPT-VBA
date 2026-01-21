' ===================================================================
' Beamer Like PowerPoint 自动化格式宏
' 功能:
' 1. 将在 排列-选择窗格 中命名为 TargetImage 的图片所在幻灯片作为模板，插入目标路径中的一系列 frame_*.png 图片以模拟动画效果
'
' 作者: ZilongWang
' 日期: 2026-1-21
' 兼容性: MacOS
' ===================================================================

Sub GenerateFramesAfterTemplate()
    Dim folderPath As String
    Dim fileName As String
    Dim pptSlide As Slide
    Dim targetShape As Shape
    Dim templateSlide As Slide
    Dim newImg As Shape
    Dim sldRange As SlideRange
    Dim foundTemplate As Boolean
    Dim insertIndex As Integer
    
    ' --- 路径设置 ---
    folderPath = "/Users/wangzilong/Downloads/intersection/"
    
    ' --- 1. 定位模板及其位置 ---
    foundTemplate = False
    For Each pptSlide In ActivePresentation.Slides
        For Each targetShape In pptSlide.Shapes
            If targetShape.Name = "TargetImage" Then
                Set templateSlide = pptSlide
                ' 记录模板的当前索引
                insertIndex = pptSlide.SlideIndex
                foundTemplate = True
                Exit For
            End If
        Next targetShape
        If foundTemplate Then Exit For
    Next pptSlide
    
    If Not foundTemplate Then
        MsgBox "未找到名为 'TargetImage' 的模板占位符。", vbCritical
        Exit Sub
    End If

    ' --- 2. 遍历并按顺序插入 ---
    fileName = Dir(folderPath & "frame_*.png")
    
    If fileName = "" Then
        MsgBox "未找到图片，请检查路径。"
        Exit Sub
    End If
    
    Do While fileName <> ""
        templateSlide.Copy
        
        ' --- 在 insertIndex + 1 的位置粘贴 ---
        Set sldRange = ActivePresentation.Slides.Paste(insertIndex + 1)
        Set pptSlide = sldRange(1)
        
        ' 更新插入位置索引，保证序列按 frame_1, frame_2 排序
        insertIndex = pptSlide.SlideIndex
        
        ' 替换图片
        For Each targetShape In pptSlide.Shapes
            If targetShape.Name = "TargetImage" Then
                Set newImg = pptSlide.Shapes.AddPicture(folderPath & fileName, _
                             msoFalse, msoCTrue, _
                             targetShape.Left, targetShape.Top, targetShape.Width, targetShape.Height)
                
                newImg.Name = "TargetImage"
                targetShape.Delete
                Exit For
            End If
        Next targetShape
        
        fileName = Dir()
    Loop
    
    MsgBox "处理完成！新幻灯片已插入到原模板页之后。"
End Sub

