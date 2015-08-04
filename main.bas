Attribute VB_Name = "main"
Sub plotter()
    'Created by Patrick Moore
    'Free to use/modify/copy/change

    '---------DECLARE AND SET SCALE VARIABLES-----------
    
    'declare input scale variables
    Dim dblDomXMin As Double
    Dim dblDomXMax As Double
    Dim dblDomYMin As Double
    Dim dblDomYMax As Double
    
    'set input scale variables to values
    dblDomXMin = Range("domXMin").Value
    dblDomXMax = Range("domXMax").Value
    dblDomYMin = Range("domYMin").Value
    dblDomYMax = Range("domYMax").Value
    
    '---------DECLARE AND SET RANGE/DATA VARIABLES-----------
    
    'declare string variables to hold range locations for data
    Dim strXPoints As String
    Dim strYPoints As String
    Dim strLabels As String
    Dim strColors As String
    
    'set string variables to cell value that contains location of data
    strXPoints = Range("xRange").Value
    strYPoints = Range("yRange").Value
    strLabels = Range("labels").Value
    strColors = Range("colors").Value

    'declare arrays to hold data
    Dim arrX As Variant
    Dim arrY As Variant
    Dim arrLabels As Variant
    Dim arrColors As Variant
    Dim arrSkip As Variant
    
    'set arrays = data ranges
    arrX = Range(strXPoints)
    arrY = Range(strYPoints)
    arrLabels = Range(strLabels)
    arrColors = Range(strColors)
    arrSkip = Range("tblData[Skip]")
        
    '------------------PLOT POINTS------------------
    
    'declare integer to hold previous count
    Dim intPrev As Integer
    Dim dblX As Double
    Dim dblY As Double
    Dim colShapes As Collection
    'declare and set shape variable
    Dim objRect As Object
    Set objRect = ActiveSheet.Shapes("rctOuter")
    Dim objTemp As Object
    Set colShapes = New Collection
    
    
    'for each value in data array
    For X = 1 To UBound(arrX)
        If arrSkip(X, 1) <> "Y" Then
        'set intPrev to previous count
        intPrev = prevDrawn(arrX, arrY, X)
        
        'set scales
        dblX = scaler(CDbl(arrX(X, 1)), dblDomXMin, dblDomXMax, 0, objRect.width)
        dblY = scaler(CDbl(arrY(X, 1)), dblDomYMin, dblDomYMax, 0, objRect.height)
        
        'draw circles
        Call drawCircles(dblX, dblY, objRect.Left, objRect.Top, objRect.height, 15, intPrev, CStr(arrLabels(X, 1)), CInt(arrColors(X, 1)))
        Set objTemp = Selection.ShapeRange
        
        colShapes.Add objTemp.Name
        End If
    Next
    
    Dim arrShapes() As Variant
    
    ReDim arrShapes(colShapes.Count + 1)
    
    
    For X = 1 To colShapes.Count
        arrShapes(X) = colShapes(X)
        Debug.Print arrShapes(X)
    Next
    
    ActiveSheet.Shapes.Range(arrShapes).Group.Select
    
    
    
    
End Sub




Sub drawCircles(X As Double, Y As Double, bufferX As Double, bufferY As Double, height As Double, R As Double, OffsetCount As Integer, strID As String, intColor As Integer)

    Dim dblOffset As Double
    dblOffset = 10
    Dim strColor As String
    strColor = "msoThemeColorAccent" & CStr(intColor)
    'draw circle on rectangle
    ActiveSheet.Shapes.AddShape(msoShapeOval, (bufferX + X - (R / 2) + (dblOffset * OffsetCount)), bufferY + (height - Y - (R / 2) + (dblOffset * OffsetCount)), R, R).Select
    With Selection.ShapeRange
    If intColor <> 0 Then
    .TextFrame2.TextRange.Characters.Text = strID
    .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End If
    Select Case intColor
        Case 1
            .Fill.ForeColor.RGB = RGB(59, 89, 152)
        Case 2
            .Fill.ForeColor.RGB = RGB(139, 153, 150)
        Case 3
            .Fill.ForeColor.RGB = RGB(255, 207, 57)
        Case 4
            .Fill.ForeColor.RGB = RGB(102, 200, 90)
        Case 5
            .Fill.ForeColor.RGB = RGB(227, 101, 101)
        Case 6
            .Fill.ForeColor.RGB = RGB(255, 175, 45)
        Case 0
            .Fill.ForeColor.RGB = RGB(225, 225, 225)
            .Fill.Transparency = 0.6
        Case Else
            .Fill.ForeColor.RGB = RGB(155, 89, 152)
    End Select
    
    
    
    End With

End Sub

Function scaler(dblVal As Double, dblInpMin As Double, dblInpMax As Double, dblOutMin As Double, dblOutMax As Double) As Double
    ' Perform linear scale and return point
    scaler = dblOutMin + (dblOutMax - dblOutMin) * ((dblVal - dblInpMin) / (dblInpMax - dblInpMin))

End Function


Function prevDrawn(arrX As Variant, arrY As Variant, intCounter As Variant) As Integer

    Dim intRepeat As Integer
    intRepeat = 0
    'for each element in array (Up to the current element)
    For X = 1 To intCounter - 1
        'if element = previous element, increment intRepeat
        If arrX(X, 1) = arrX(intCounter, 1) And arrY(X, 1) = arrY(intCounter, 1) Then
            intRepeat = intRepeat + 1
        End If
    
    Next

    'return intRepeat
    prevDrawn = intRepeat

End Function
