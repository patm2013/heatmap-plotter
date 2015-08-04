Attribute VB_Name = "util"
Sub readArea()
    
    'declare and set shape variable
    Dim objRectOuter As Object
    Set objRectOuter = ActiveSheet.Shapes("rctOuter")
    
    'set ranges = shape parameters
    With objRectOuter
    Range("rangeXMax") = .width
    Range("rangeYMax") = .height
    Range("bufferX") = .Left
    Range("bufferY") = .Top
    End With
 
End Sub

Sub writeArea()

    'declare and set shape variable
    Dim objRectOuter As Object
    Set objRectOuter = ActiveSheet.Shapes("rctOuter")
    
    'set shape parameters = ranges
    With objRectOuter
     .width = Range("rangeXMax")
     .height = Range("rangeYMax")
     .Left = Range("bufferX")
     .Top = Range("bufferY")
    End With

End Sub

Sub domainScale()

    Dim arrX As Variant
    Dim arrY As Variant
    
    Dim strX As String
    Dim strY As String
    
    strX = Range("xRange").Value
    strY = Range("yRange").Value
    
    arrX = Range(strX)
    arrY = Range(strY)

    Range("domXMax").Value = max(arrX)
    Range("domYMax").Value = max(arrY)
    Range("domXMin").Value = min(arrX)
    Range("domYMin").Value = min(arrY)
End Sub
Function max(arr As Variant) As Double

    Dim intHolder

    For X = 1 To UBound(arr)
    
        If X = 1 Then
            intHolder = arr(X, 1)
        End If
        
        If arr(X, 1) > intHolder Then
            intHolder = arr(X, 1)
        End If
    
    Next
    
    max = intHolder

End Function

Function min(arr As Variant) As Double

    Dim intHolder

    For X = 1 To UBound(arr)
    
        If X = 1 Then
            intHolder = arr(X, 1)
        End If
        
        If arr(X, 1) < intHolder Then
            intHolder = arr(X, 1)
        End If
    
    Next
    
    min = intHolder

End Function


Sub deleteShapes()

    ActiveSheet.Shapes.SelectAll
    For Each shp In Selection
    
        If Left(shp.Name, 4) <> "perm" And Left(shp.Name, 3) <> "rct" Then
            shp.Delete
        End If
    
    Next

    Cells(1, 1).Select

End Sub

Sub selectShapes()

    ActiveSheet.Shapes.SelectAll
    For Each shp In Selection
    
        If Left(shp.Name, 4) <> "perm" And Left(shp.Name, 3) <> "rct" Then
            shp.Delete
        End If
    
    Next

    Cells(1, 1).Select

End Sub


Sub resetRect()

    Dim newRect As Object
    Dim oldRect As Object
    
    On Error Resume Next
    Set oldRect = ActiveSheet.Shapes("rctOuter")
    oldRect.Delete
    On Error GoTo 0
    
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 500, 65, 300, 250).Select
    With Selection.ShapeRange
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(255, 0, 0)
        .Line.Transparency = 0
        .Line.DashStyle = msoLineDash
        .Fill.Visible = msoFalse
        .Shadow.Visible = msoFalse
    End With
    
    Set newRect = Selection
    newRect.Name = "rctOuter"

End Sub

Sub drawGrids()

    Dim intGridLines As Integer
    
    intGridLines = CInt(Range("gridLines").Value)
    
    Call drawGridY(intGridLines)
    Call drawGridX(intGridLines)


End Sub



Sub drawGridY(intLines As Integer)

    'declare and set shape variable
    Dim objRect As Object
    Set objRect = ActiveSheet.Shapes("rctOuter")
    
    Dim startX As Double
    Dim endX As Double
    Dim xInterval As Double
    
    xInterval = objRect.width / intLines
    
    For X = 1 To intLines
    
        startX = objRect.Left + (xInterval * X)
        endX = objRect.Left + (xInterval * X)

        ActiveSheet.Shapes.AddConnector(msoConnectorStraight, startX, objRect.Top, endX, (objRect.Top + objRect.height)).Select
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.150000006
            .Transparency = 0
        End With
    
    Next

End Sub

Sub drawGridX(intLines As Integer)

    'declare and set shape variable
    Dim objRect As Object
    Set objRect = ActiveSheet.Shapes("rctOuter")
    
    Dim startY As Double
    Dim endY As Double
    Dim yInterval As Double
    
    yInterval = objRect.height / intLines
    
    For X = 1 To intLines
    
        startY = objRect.Top + (yInterval * X)
        endY = objRect.Top + (yInterval * X)

        ActiveSheet.Shapes.AddConnector(msoConnectorStraight, objRect.Left, startY, (objRect.Left + objRect.width), endY).Select
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = -0.150000006
            .Transparency = 0
        End With
    
    Next

End Sub

