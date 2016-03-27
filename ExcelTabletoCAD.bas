Attribute VB_Name = "excel表格插入CAD"
Public Sub ExcelTabletoCAD()
    'Excel表格到CAD
    'On Error Resume Next
    '对AutoCAD部件的引用，方法如下（文中以'开头的语句为注释）：
    Dim acadApp As Object '声明AutoCAD应用程序对象变量
    Dim circleObj As Object, textObj As Object '声明AutoCAD中的对象变量,圆,文本
    Dim lineobj As Object, layerObj As Object '声明AutoCAD中的对象变量,直线,图层
    'On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application") '若AutoCAD已运行则获得它的对象实例
    '        If Err Then  '如果AutoCAD没有运行
    '            Err.Clear
    '            Set acadApp = CreateObject("AutoCAD.Application") '创建AutoCAD应用程序对象实例
    '                If Err Then  '若没有安装AutoCAD
    '                MsgBox Err.Description
    '                Exit Sub
    '                End If
    '        End If
    acadApp.Visible = True '从Excel中的“计算”表中读取各导线点的坐标，在AutoCAD中展点，方法如下： '建立新图层,层名"点",层颜色为红色,并置为当前层
    Dim acadDoc As AcadDocument
    Set acadDoc = acadApp.ActiveDocument
    '  AcadApp.ActiveDocument.ActiveSpace = acModelSpace
    '**********选择cad中的多段线
    
    ' 连接Excel应用程序
    Dim xlApp As Excel.Application
    Set xlApp = GetObject(, "Excel.Application")
    If Err Then
        MsgBox " Excel 应用程序没有运行。请启动 Excel 并重新运行程序。"
        Exit Sub
    End If
    Dim xlSheet As Worksheet
    Set xlSheet = xlApp.ActiveSheet
    
    ' 当初考虑将表格做成块的方式，可以根据需要取舍。
    'Dim iPt(0 To 2) As Double
    'iPt(0) = 0: iPt(1) = 0: iPt(2) = 0
    Dim BlockObj As AcadBlock
    Set BlockObj = acadDoc.Blocks("*Model_Space")
    Dim iPt As Variant
    iPt = acadDoc.Utility.GetPoint(, "指定表格的插入点: ")
    If IsEmpty(iPt) Then Exit Sub
    Dim xlRange As Range
    Debug.Print xlSheet.UsedRange.Address
    For Each xlRange In xlSheet.UsedRange
        AddLine BlockObj, iPt, xlRange
        AddText BlockObj, iPt, xlRange
    Next
    Set xlRange = Nothing
    Set xlSheet = Nothing
    Set xlApp = Nothing
End Sub
'边框线条粗细
Function LineWidth(ByVal xlBorder As Border) As Double
    Select Case xlBorder.Weight
    Case xlThin
        LineWidth = 0
    Case xlMedium
        LineWidth = 0.35
    Case xlThick
        LineWidth = 0.7
    Case Else
        LineWidth = 0
    End Select
End Function

'边框线条颜色，处理的颜色不全，请自己添加
Function LineColor(ByVal xlBorder As Border) As Integer
    Select Case xlBorder.ColorIndex
    Case xlAutomatic
        LineColor = acByLayer
    Case 3
        LineColor = acRed
    Case 4
        LineColor = acGreen
    Case 5
        LineColor = acBlue
    Case 6
        LineColor = acYellow
    Case 8
        LineColor = acCyan
    Case 9
        LineColor = acMagenta
    Case Else
        LineColor = acByLayer
    End Select
End Function

'给制边框
Sub AddLine(ByRef BlockObj As AcadBlock, ByVal iPt As Variant, ByVal xlRange As Range)
    If xlRange.Borders(xlEdgeLeft).LineStyle = xlNone _
        And xlRange.Borders(xlEdgeBottom).LineStyle = xlNone _
        And xlRange.Borders(xlEdgeRight).LineStyle = xlNone _
        And xlRange.Borders(xlEdgeTop).LineStyle = xlNone Then Exit Sub
        Dim rl As Double
        Dim rt As Double
        Dim rw As Double
        Dim rh As Double
        rl = PToM(xlRange.Left)
        rt = PToM(xlRange.Top)
        rw = PToM(xlRange.Width)
        rh = PToM(xlRange.Height)
        Dim pPt(0 To 3) As Double
        Dim pLineObj As AcadLWPolyline
        
        ' 左边框的处理，仅第一列才做处理。
        If xlRange.Borders(xlEdgeLeft).LineStyle <> xlNone And xlRange.Column = 1 Then
            pPt(0) = iPt(0) + rl: pPt(1) = iPt(1) - rt
            pPt(2) = iPt(0) + rl: pPt(3) = iPt(1) - (rt + rh)
            Set pLineObj = BlockObj.AddLightWeightPolyline(pPt)
            pLineObj.ConstantWidth = LineWidth(xlRange.Borders(xlEdgeLeft))
            pLineObj.color = LineColor(xlRange.Borders(xlEdgeLeft))
        End If
        
        ' 下边框的处理，对于合并单元格，只处理最后一行。
        If xlRange.Borders(xlEdgeBottom).LineStyle <> xlNone And (xlRange.Row = xlRange.MergeArea.Row + xlRange.MergeArea.Rows.Count - 1) Then
            pPt(0) = iPt(0) + rl: pPt(1) = iPt(1) - (rt + rh)
            pPt(2) = iPt(0) + rl + rw: pPt(3) = iPt(1) - (rt + rh)
            Set pLineObj = BlockObj.AddLightWeightPolyline(pPt)
            pLineObj.ConstantWidth = LineWidth(xlRange.Borders(xlEdgeBottom))
            pLineObj.color = LineColor(xlRange.Borders(xlEdgeBottom))
        End If
        
        ' 右边框的处理，对于合并单元格，只处理最后一列。
        If xlRange.Borders(xlEdgeRight).LineStyle <> xlNone And (xlRange.Column >= xlRange.MergeArea.Column + xlRange.MergeArea.Columns.Count - 1) Then
            pPt(0) = iPt(0) + rl + rw: pPt(1) = iPt(1) - (rt + rh)
            pPt(2) = iPt(0) + rl + rw: pPt(3) = iPt(1) - rt
            Set pLineObj = BlockObj.AddLightWeightPolyline(pPt)
            pLineObj.ConstantWidth = LineWidth(xlRange.Borders(xlEdgeRight))
            pLineObj.color = LineColor(xlRange.Borders(xlEdgeRight))
        End If
        
        ' 上边框的处理，仅第一行才做处理。
        If xlRange.Borders(xlEdgeTop).LineStyle <> xlNone And xlRange.Row = 1 Then
            pPt(0) = iPt(0) + rl + rw: pPt(1) = iPt(1) - rt
            pPt(2) = iPt(0) + rl: pPt(3) = iPt(1) - rt
            Set pLineObj = BlockObj.AddLightWeightPolyline(pPt)
            pLineObj.ConstantWidth = LineWidth(xlRange.Borders(xlEdgeTop))
            pLineObj.color = LineColor(xlRange.Borders(xlEdgeTop))
        End If
        Set pLineObj = Nothing
    End Sub
    
    '给制文本
    
    Sub AddText(ByRef BlockObj As AcadBlock, ByVal InsertionPoint As Variant, ByVal xlRange As Range)
        '对AutoCAD部件的引用，方法如下（文中以'开头的语句为注释）：
        Dim acadApp As Object '声明AutoCAD应用程序对象变量
        Dim circleObj As Object, textObj As Object '声明AutoCAD中的对象变量,圆,文本
        Dim lineobj As Object, layerObj As Object '声明AutoCAD中的对象变量,直线,图层
        'On Error Resume Next
        Set acadApp = GetObject(, "AutoCAD.Application") '若AutoCAD已运行则获得它的对象实例
        '        If Err Then  '如果AutoCAD没有运行
        '            Err.Clear
        '            Set acadApp = CreateObject("AutoCAD.Application") '创建AutoCAD应用程序对象实例
        '                If Err Then  '若没有安装AutoCAD
        '                MsgBox Err.Description
        '                Exit Sub
        '                End If
        '        End If
        acadApp.Visible = True '从Excel中的“计算”表中读取各导线点的坐标，在AutoCAD中展点，方法如下： '建立新图层,层名"点",层颜色为红色,并置为当前层
        Dim acadDoc As AcadDocument
        Set acadDoc = acadApp.ActiveDocument
        '  AcadApp.ActiveDocument.ActiveSpace = acModelSpace
        '**********选择cad中的多段线
        
        
        If xlRange.text = "" Then Exit Sub
        Dim rl As Double
        Dim rt As Double
        Dim rw As Double
        Dim rh As Double
        rl = PToM(xlRange.Left)
        rt = PToM(xlRange.Top)
        rw = PToM(xlRange.MergeArea.Width)
        rh = PToM(xlRange.MergeArea.Height)
        Dim i As Integer
        Dim s As String
        For i = 1 To Len(xlRange.text) '将EXCEL的换行符替换成P，注如果是在R2002以上可使用Replace函数。
            If Asc(Mid(xlRange.text, i, 1)) = 10 Then
                s = s & "P"
            Else
                s = s & Mid(xlRange.text, i, 1)
            End If
        Next
        Dim iPt(0 To 2) As Double
        iPt(0) = InsertionPoint(0) + rl: iPt(1) = InsertionPoint(1) - rt: iPt(2) = 0
        Dim mTextObj As AcadMText
        Set mTextObj = BlockObj.AddMText(iPt, rw, s) '"{f" & xlRange.Font.Name & ";" & s & "}")
        mTextObj.LineSpacingFactor = 0.75
        mTextObj.Height = PToM(xlRange.Font.Size)
        
        ' 处理文字的对齐方式
        Dim tPt As Variant
        If xlRange.VerticalAlignment = xlTop And (xlRange.HorizontalAlignment = xlLeft Or xlRange.HorizontalAlignment = xlGeneral) Then
            mTextObj.AttachmentPoint = acAttachmentPointTopLeft
            tPt = iPt
        ElseIf xlRange.VerticalAlignment = xlTop And xlRange.HorizontalAlignment = xlCenter Then
            mTextObj.AttachmentPoint = acAttachmentPointTopCenter
            tPt = acadDoc.Utility.PolarPoint(iPt, 0, rw / 2)
        ElseIf xlRange.VerticalAlignment = xlTop And xlRange.HorizontalAlignment = xlRight Then
            mTextObj.AttachmentPoint = acAttachmentPointTopRight
            tPt = acadDoc.Utility.PolarPoint(iPt, 0, rw)
        ElseIf xlRange.VerticalAlignment = xlCenter And (xlRange.HorizontalAlignment = xlLeft _
            Or xlRange.HorizontalAlignment = xlGeneral) Then
            mTextObj.AttachmentPoint = acAttachmentPointMiddleLeft
            tPt = acadDoc.Utility.PolarPoint(iPt, -1.5707963, rh / 2)
        ElseIf xlRange.VerticalAlignment = xlCenter And xlRange.HorizontalAlignment = xlCenter Then
            mTextObj.AttachmentPoint = acAttachmentPointMiddleCenter
            tPt = acadDoc.Utility.PolarPoint(iPt, -1.5707963, rh / 2)
            tPt = acadDoc.Utility.PolarPoint(tPt, 0, rw / 2)
        ElseIf xlRange.VerticalAlignment = xlCenter And xlRange.HorizontalAlignment = xlRight Then
            mTextObj.AttachmentPoint = acAttachmentPointMiddleRight
            tPt = acadDoc.Utility.PolarPoint(iPt, -1.5707963, rh / 2)
            tPt = acadDoc.Utility.PolarPoint(tPt, 0, rw / 2)
        ElseIf xlRange.VerticalAlignment = xlBottom And (xlRange.HorizontalAlignment = xlLeft _
            Or xlRange.HorizontalAlignment = xlGeneral) Then
            mTextObj.AttachmentPoint = acAttachmentPointBottomLeft
            tPt = acadDoc.Utility.PolarPoint(iPt, -1.5707963, rh)
        ElseIf xlRange.VerticalAlignment = xlBottom And xlRange.HorizontalAlignment = xlCenter Then
            mTextObj.AttachmentPoint = acAttachmentPointBottomCenter
            tPt = acadDoc.Utility.PolarPoint(iPt, -1.5707963, rh)
            tPt = acadDoc.Utility.PolarPoint(tPt, 0, rw / 2)
        ElseIf xlRange.VerticalAlignment = xlBottom And xlRange.HorizontalAlignment = xlRight Then
            mTextObj.AttachmentPoint = acAttachmentPointBottomRight
            tPt = acadDoc.Utility.PolarPoint(iPt, -1.5707963, rh)
            tPt = acadDoc.Utility.PolarPoint(tPt, 0, rw)
        End If
        mTextObj.InsertionPoint = tPt
        Set mTextObj = Nothing
    End Sub
    
    ' 磅换算成毫米
    
    ' 注：意义不大，转换的尺寸有偏差，最好自己设定一个转换规则。
    
    Function PToM(ByVal Points As Double) As Double
        PToM = Points * 0.3527778
    End Function

