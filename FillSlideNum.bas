Attribute VB_Name = "FillSlideNum"

    'https://www.youtube.com/@ExperiencesSharing_Walker
    Const bBig As Boolean = True
    Const fontSize As Integer = 14
    Const nLeft As Integer = 18
    Const nTop As Integer = 490
    Const nNumGap As Integer = 4
    Const nWidth_Date As Integer = 50
    Const nWidth_Num As Integer = 80
    Const nHeight As Integer = 30

Sub EmbedSlideNumbers()
    Dim pptPresentation As Object
    Dim slide As Object
    Dim totalSlides As Integer
    Dim slideNumber As Integer
    Dim strNumName As String
    Dim strDateName As String
    Dim nStart As Integer
    Dim nResult As Integer
    
    ' ��ȡ��ǰ�򿪵� PowerPoint ��ʾ�ĸ�
    Set pptPresentation = ActivePresentation
    
    ' ��ȡ�õ�Ƭ������
    totalSlides = pptPresentation.Slides.count
    
    ' ����ÿ���õ�Ƭ��Ƕ��õ�Ƭ���
    slideNumber = 0
    For Each slide In pptPresentation.Slides
        slideNumber = slideNumber + 1
        strNumName = FindNumName(slide)
        strDateName = FindDateName(slide)
        If strNumName <> "" And strDateName <> "" Then
            If bBig Then
                'ǰ���ȥ 2 ҳ��������Ŀ¼
                '����ǰȥ 3 ҳ��������л
                nStart = 2
            Else
                'ǰ���ȥ 1 ҳ������
                '����ǰȥ 2 ҳ��������л
                nStart = 1
            End If
            If slideNumber > nStart And slideNumber <> totalSlides Then
                slide.Shapes(strNumName).TextFrame.TextRange.Text = slideNumber - nStart & " / " & totalSlides - (nStart + 1)
                Call SetShapeData(slide.Shapes(strDateName), nLeft, nTop, nWidth_Date, nHeight)
                Call SetShapeData(slide.Shapes(strNumName), nLeft + nWidth_Date, nTop - nNumGap, nWidth_Num, nHeight)
                slide.Shapes(strDateName).Visible = True
                slide.Shapes(strNumName).Visible = True
            Else
                '���桢Ŀ¼����л�������ҳ��
                slide.Shapes(strNumName).Visible = False
                slide.Shapes(strDateName).Visible = False
            End If
        End If
    Next slide
    nResult = MsgBox("Normalization complete.", vbOKOnly)
End Sub

Sub SetShapeData(ByRef shp As Shape, nLeft As Integer, nTop As Integer, nWidth As Integer, nHeight As Integer)
    shp.Left = nLeft
    shp.Top = nTop
    shp.Width = nWidth
    shp.Height = nHeight
    shp.TextFrame.TextRange.Font.Size = fontSize
End Sub

Function FindNumName(ByRef pptSlide As slide) As String
    Dim shp As Shape
    Dim strData As String
    For Each shp In pptSlide.Shapes
        If InStr(shp.Name, "Slide Number") > 0 Or InStr(shp.Name, "���ռλ��") > 0 Then
            strData = shp.Name
            FindNumName = strData
            Exit Function
        End If
    Next shp
    FindNumName = ""
End Function

Function FindDateName(ByRef pptSlide As slide) As String
    Dim shp As Shape
    Dim strData As String
    For Each shp In pptSlide.Shapes
        If InStr(shp.Name, "Date Placeholder") > 0 Then
            strData = shp.Name
            FindDateName = strData
            Exit Function
        End If
    Next shp
    FindDateName = ""
End Function




