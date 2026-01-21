' ===================================================================
' Beamer Like PowerPoint 自动化格式宏
' 功能:
' 1. 将在 排列-选择窗格 中命名为 TargetImage 的图片所在幻灯片作为模板，插入目标路径中的一系列 frame_*.png 图片以模拟动画效果
'
' 作者: ZilongWang
' 日期: 2026-1-21
' 兼容性: MacOS
' ===================================================================

Sub GenerateFramesForMac_Final()
    Dim folderPath As String
    Dim fileName As String
    Dim pptSlide As Slide
    Dim targetShape As Shape
    Dim templateSlide As Slide
    Dim newImg As Shape
    Dim sldRange As SlideRange
    
    ' --- 路径设置 ---
    ' 提示：确保路径以 / 结尾，且用户名正确
    folderPath = "/Users/wangzilong/Downloads/intersection/"
    
    fileName = Dir(folderPath & "frame_*.png")
    
    If fileName = "" Then
        MsgBox "未找到图片，请检查路径和文件名。"
        Exit Sub
    End If
    
    ' 设置第一页为模板
    Set templateSlide = ActivePresentation.Slides(1)
    
    
    Do While fileName <> ""
        templateSlide.Copy
        
        ' --- 核心修正点 ---
        ' Paste 返回 SlideRange，我们通过 (1) 提取其中的 Slide 对象
        Set sldRange = ActivePresentation.Slides.Paste(ActivePresentation.Slides.Count + 1)
        Set pptSlide = sldRange(1)
        
        ' 寻找并替换图片
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
    
    MsgBox "所有数据帧已成功导入！"
End Sub

