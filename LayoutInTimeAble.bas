Attribute VB_Name = "LayoutInTimeAble"
Option Explicit


Const CWKSNAMEINTIMEABLE As String = "Time Table"
Const CCOLORFRAMETITLE As Long = 12566463    'color : ?RGB(rrr,ggg,bbb) -> Long
Const CCOLORFRAME1 As Long = 15921906
Const CCOLORFRAME2 As Long = 14277081
Const CCOLORNOWTIMERECT As Long = 10213059

Const CFONT As String = "Arial"
Const CFONTSIZE As Integer = 12

Const CTITLE As String = "Prozess"
Const CTITLE0 As String = "Schritt"
Const CTITLE1 As String = "Prozess"
Const CTITLE2 As String = "Verantwortlicher"

Const cSTARTROW As Integer = 3
Const cSTARTCOL As Integer = 1

Const CTOPFRAMEBORDERWIDTH As Integer = 2
Const CLEADINGFRAMEBORDERWIDTH As Integer = 2
Const CTRAILINGFRAMEBORDERWIDTH As Integer = 1

Const CCOLPERMONTH As Integer = 1
Const CROWPERSTEP As Integer = 1
Const cTITLE1COLWIDTH As Double = 3
Const CCELLWIDTH As Integer = 15
Const CCELLHEIGHT As Integer = 30

Public pubTIMERECTHEIGHT As Double

Public Function InitInTimeAble() As InTimeAbleObj
    Dim wks As Worksheet
    Dim inTimeAble As InTimeAbleObj
    Dim nx4InputArray As Variant
    
    Set wks = Worksheets(cWKSNAME1)
    
    Set inTimeAble = New InTimeAbleObj
    nx4InputArray = gatherInTimeAbleData(wks)
    inTimeAble.init nx4InputArray
    
    Set InitInTimeAble = inTimeAble
    pubTIMERECTHEIGHT = CCELLHEIGHT / 2
End Function

Private Function gatherInTimeAbleData(wks As Worksheet) As Variant
    Dim listObj As ListObject
    Dim inputArray As Variant
    Dim returnNx4InputArray As Variant
    Dim title0Col, title1Col, title2Col, dateCol As Integer
    
    Dim lb As Integer
    Dim length, i As Integer
    
    Set listObj = wks.ListObjects(cTABLE)

    With listObj
        inputArray = .DataBodyRange
        title0Col = .ListColumns(CTITLE0COL).index
        title1Col = .ListColumns(CTITLE1COL).index
        title2Col = .ListColumns(CTITLE2COL).index
        dateCol = .ListColumns(CDATECOL).index
        
        lb = LBound(inputArray, 2)
        
        inputArray = arraysStartWith0(inputArray)
        inputArray = removeNARowsFromMatrix(inputArray)
        If isNA(inputArray) Or colIsNA(inputArray, dateCol - lb) Then
            MsgBox "Please enter data in " & cTABLE & " and also date values"
            End
        End If
        length = UBound(inputArray, 1)
        ReDim returnNx4InputArray(0 To length, 0 To 3)
        For i = 0 To length
            returnNx4InputArray(i, 0) = inputArray(i, title0Col - lb)
            returnNx4InputArray(i, 1) = inputArray(i, dateCol - lb)
            returnNx4InputArray(i, 2) = inputArray(i, title1Col - lb)
            returnNx4InputArray(i, 3) = inputArray(i, title2Col - lb)
        Next i
    End With
    
    gatherInTimeAbleData = returnNx4InputArray
End Function

Sub LayoutInTimeAble()
    Dim wks As Worksheet
    Dim inTimeAble As InTimeAbleObj
    Dim timeRect As TimeRectObj
    Dim nx4InputArray As Variant
    Dim yearString As String
    Dim title As String
    
    If WorksheetExists(CWKSNAMEINTIMEABLE) Then
        MsgBox "Timetable does already exist. Please rename or delete the worksheet containing the timetable"
        End
    End If
    
    Set inTimeAble = InitInTimeAble()
    Set wks = Worksheets(cWKSNAME1)
    
On Error Resume Next
    title = wks.Range("Title")
    If title = "" Then
        title = "Title"
    End If
On Error GoTo 0

    Set wks = Worksheets.Add
    With wks
        .name = CWKSNAMEINTIMEABLE
        .Cells.rowHeight = CCELLHEIGHT
        .Cells.ColumnWidth = CCELLWIDTH
        .Cells.Font.name = "Arial"
        .Cells.Interior.color = 16777215
    End With
    
    createFrame wks, inTimeAble
    layoutTimeRects wks, inTimeAble
    layoutCurrentTimeRect wks, inTimeAble
    
    yearString = year(CDate(inTimeAble.minTime))
    If cSTARTROW > 1 Then
        fillCell wks, cSTARTROW - 1, cSTARTCOL + CLEADINGFRAMEBORDERWIDTH, CTITLE & " " & title & " " & yearString, 16777215, xlLeft, 16, True
        wks.Rows(cSTARTROW - 1).AutoFit
    End If
    wks.Visible = True
    wks.Activate
End Sub

Private Sub createFrame(wks As Worksheet, inTimeAble As InTimeAbleObj)
    Dim length As Integer
    Dim width As Integer
    Dim interval As Double
    Dim months As Integer
    Dim startMonth As Integer
    Dim tMonth As String
    Dim bgColor As Long
    
    Dim itemIndex As Integer
    Dim startRow As Integer
    Dim cell As Variant
    Dim i As Integer
    
    interval = inTimeAble.maxTime - inTimeAble.minTime + 30  '+30 -> include first and last month
    months = interval / 30.5
    startMonth = month(CDate(inTimeAble.minTime))
    
    With wks
        For i = 0 To months - 1
            tMonth = getMonth(startMonth + i)
            fillCell wks, cSTARTROW, cSTARTCOL + CLEADINGFRAMEBORDERWIDTH + CCOLPERMONTH - 1 + i, tMonth, CCOLORFRAMETITLE, xlLeft, 12, True
        Next i
        
        width = i + CLEADINGFRAMEBORDERWIDTH + CTRAILINGFRAMEBORDERWIDTH + cSTARTCOL - 1
        length = inTimeAble.calendarItems.Count - 1
        
        
        fillCell wks, cSTARTROW, cSTARTCOL, CTITLE0, CCOLORFRAMETITLE, xlLeft, 12, True
        fillCell wks, cSTARTROW, cSTARTCOL + 1, CTITLE1, CCOLORFRAMETITLE, xlLeft, 12, True
        fillCell wks, cSTARTROW, width, CTITLE2, CCOLORFRAMETITLE, xlLeft, 12, True

        .Range(Cells(cSTARTROW, cSTARTCOL), Cells(cSTARTROW + CTOPFRAMEBORDERWIDTH - 1, width)).Interior.color = CCOLORFRAMETITLE
        For i = 0 To length * CROWPERSTEP Step CROWPERSTEP          'Explanation: sets changing colors of rows
            itemIndex = i / CROWPERSTEP
            If (i / CROWPERSTEP) Mod 2 = 0 Then
                bgColor = CCOLORFRAME1
            Else
                bgColor = CCOLORFRAME2
            End If
            startRow = cSTARTROW + CTOPFRAMEBORDERWIDTH + i
            fillCell wks, startRow, cSTARTCOL, inTimeAble.calendarItems.Item(itemIndex + 1).id
            fillCell wks, startRow, cSTARTCOL + 1, inTimeAble.calendarItems.Item(itemIndex + 1).title1
            fillCell wks, startRow, width, inTimeAble.calendarItems.Item(itemIndex + 1).title2, 1, xlRight
            .Cells(startRow, cSTARTCOL + 1).ColumnWidth = cTITLE1COLWIDTH * CCELLWIDTH
            .Range(.Cells(startRow, cSTARTCOL), .Cells(startRow + CROWPERSTEP - 1, width)).Interior.color = bgColor
            With .Range(.Cells(startRow, cSTARTCOL + 1), .Cells(startRow, cSTARTCOL + 1))
                .WrapText = True
                .Rows.AutoFit
                If .rowHeight < CCELLHEIGHT Then .rowHeight = CCELLHEIGHT
            End With
            For Each cell In .Range(.Cells(startRow, cSTARTCOL + CLEADINGFRAMEBORDERWIDTH), .Cells(startRow + CROWPERSTEP - 1, width - CTRAILINGFRAMEBORDERWIDTH))
                With cell
                    .Borders(xlEdgeLeft).LineStyle = xlDash
                    .Borders(xlEdgeLeft).ThemeColor = 1
                    .Borders(xlEdgeLeft).LineStyle = xlDash
                    .Borders(xlEdgeRight).Weight = xlMedium
                    .Borders(xlEdgeRight).ThemeColor = 1
                    .Borders(xlEdgeRight).Weight = xlMedium
                End With
            Next cell
        Next i
        
    End With
End Sub

Sub layoutTimeRects(wks As Worksheet, inTimeAble As InTimeAbleObj)
    Dim sizes() As SizeObj
    Dim timeRect As TimeRectObj
    Dim colWidth As Double
    Dim rowHeight As Double
    Dim minX As Double
    Dim minY As Double
    Dim curY As Double
    
    Dim length, width As Integer
    Dim i As Integer

    length = inTimeAble.timeRects.Count - 1
    With wks
        colWidth = convertCellWidth(.Cells(1, 1).ColumnWidth)
        rowHeight = .Cells(1, 1).rowHeight
        
        minY = (cSTARTROW + CTOPFRAMEBORDERWIDTH - 1) * rowHeight
        minX = getMinX(wks, inTimeAble)
        curY = minY
        For i = 0 To length
            If i > 0 Then
                curY = curY + .Cells(cSTARTROW + CTOPFRAMEBORDERWIDTH + i - 1, 1).rowHeight          'if rowheights are fitted to their content
            End If
            Set timeRect = inTimeAble.timeRects.Item(i + 1)
            
            sizes = getTimeRectSizes(timeRect, inTimeAble.minTime, curY, minX, colWidth, rowHeight, i + 1)
            createTimeRectShape wks, minX, curY, sizes, i + 1
            createTextBox wks, sizes
        Next i
    End With
    
End Sub

Private Function createTimeRectShape(wks As Worksheet, minX As Double, curY As Double, sizes() As SizeObj, newId As Integer) As shape
    Dim shapeObj As shape
    Dim size As SizeObj
    
    Dim xPos, yPos, yHeight, xWidth As Double
    Dim j, width As Integer
    
    If Not isNA(sizes) Then
           
        width = UBound(sizes)
        For j = 0 To width
            Set size = sizes(j)

            xPos = size.xPos
            yPos = size.yPos
            xWidth = size.xWidth
            yHeight = size.yHeight
            If size.xWidth <> 0 Then
                Set shapeObj = wks.Shapes.AddShape(msoShapeRectangle, xPos, yPos + 2, xWidth, yHeight - 2)         '-2 just for appearance
            Else
                Set shapeObj = wks.Shapes.AddShape(msoShapeIsoscelesTriangle, xPos - (yHeight / 2), yPos + 2, yHeight, yHeight - 2)     'yHeigth for xWidth is no mistake -> appearance; xPos - (yHeight / 2): triangle top has to be shifted half the length of trianlge to the left
            End If
            
            With shapeObj
                .Fill.ForeColor.RGB = size.color
                .Line.ForeColor.RGB = size.color
                .name = newId
            End With
        Next j
    Else
        yHeight = CCELLHEIGHT
        yPos = curY
        Set shapeObj = wks.Shapes.AddTextbox(msoTextOrientationHorizontal, minX, yPos + 2, 100, 2 * (yHeight - 5))
        With shapeObj
            .TextFrame.Characters.text = "No time provided"
            .TextEffect.FontName = CFONT
            .TextFrame.Characters.Font.size = 12
            .TextFrame2.WordWrap = False
            .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
            .name = "999"
        End With

    End If

    Set createTimeRectShape = shapeObj
End Function

Private Function createTextBox(wks As Worksheet, sizes() As SizeObj) As shape
    Dim shapeObj As shape
    Dim size As SizeObj
    
    Dim dateString As String
    
    Dim xPos, yPos, yHeight As Double
    Dim j, width As Integer
    
    If Not isNA(sizes) Then
           
        width = UBound(sizes)
        For j = 0 To width
            Set size = sizes(j)

            yPos = size.yPos + size.yHeight - 2
            If size.xWidth <> 0 Then
                xPos = size.xPos - 8
            Else
                xPos = size.xPos - (size.yHeight / 2) - 8
            End If
            yHeight = size.yHeight
            dateString = size.datum
            Set shapeObj = wks.Shapes.AddTextbox(msoTextOrientationHorizontal, xPos, yPos, 100, yHeight)
            
            With shapeObj
                .TextFrame.Characters.text = dateString
                .TextEffect.FontName = CFONT
                .TextFrame.Characters.Font.size = CFONTSIZE
                .TextFrame2.WordWrap = False
                .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
                .Fill.Visible = msoFalse
                .Line.Visible = msoFalse
            End With
        Next j
    End If

    Set createTextBox = shapeObj
End Function

Function getTimeRectSizes(timeRect As TimeRectObj, minTime As Double, curY As Double, minX As Double, colWidth As Double, rowHeight As Double, id As Integer) As SizeObj()
    Dim sizes() As SizeObj
    Dim size As SizeObj
    
    Dim timeRectSizes As Variant '0: Form, 1: xMinPos, 2: width, 3: color
    Dim yPos As Double

    Dim form As String
    Dim xPosTimeRect As Double
    Dim widthTimeRect As Double
    
    Dim startTime As Double
    Dim dateString As String
    Dim j, width As Integer
    
    timeRectSizes = timeRect.nX3SizesArray
    
    yPos = curY
    If Not isNA(timeRectSizes) Then
        width = UBound(timeRectSizes)
        ReDim sizes(0 To width)
        For j = 0 To width
            xPosTimeRect = timeRectSizes(j, 0)
            widthTimeRect = timeRectSizes(j, 1)
            
            Set size = New SizeObj
            size.xPos = minX + ((xPosTimeRect * colWidth) / 30) 'suppose one cell is 30 days wide
            size.xWidth = (widthTimeRect * colWidth) / 30
 
            size.yPos = yPos
            size.yHeight = pubTIMERECTHEIGHT
            size.color = timeRectSizes(j, 2)
            startTime = minTime + xPosTimeRect
            If size.xWidth <> 0 Then
                
                dateString = formatDate(startTime, minTime) & "-" & formatDate(startTime + widthTimeRect, minTime)
            Else
                dateString = formatDate(startTime, minTime)
            End If
            size.datum = dateString
            Set sizes(j) = size
        Next j
    End If

    getTimeRectSizes = sizes
End Function

Function layoutCurrentTimeRect(wks As Worksheet, inTimeAble As InTimeAbleObj) As shape
    Dim shapeObj As shape
    Dim sizes() As SizeObj
    Dim timeRect As TimeRectObj
    
    Dim colWidth As Double
    Dim rowHeight As Double
    Dim minY As Double
    Dim minX As Double
    
    Dim endYLine As Double
    Dim lastRow As Integer
    
    Dim i As Integer
    
    If Not CDbl(Now()) < inTimeAble.minTime Then
        With wks
            colWidth = convertCellWidth(.Cells(1, 1).ColumnWidth)
            rowHeight = CCELLHEIGHT
            For i = 1 To cSTARTROW
                minY = minY + .Cells(i, 1).rowHeight
            Next i
    
            minX = getMinX(wks, inTimeAble)
            
            Set timeRect = inTimeAble.currentTimeRect
            
            sizes = getTimeRectSizes(timeRect, inTimeAble.minTime, minY, minX, colWidth, rowHeight, 1)
            Set shapeObj = createTimeRectShape(wks, minX, minY, sizes, 0)
            
            With shapeObj
                .Fill.ForeColor.RGB = CCOLORNOWTIMERECT
                .Line.ForeColor.RGB = CCOLORNOWTIMERECT
            End With
            
            lastRow = .Cells(655, cSTARTCOL).End(xlUp).row + CROWPERSTEP
            For i = 1 To lastRow
                endYLine = endYLine + .Cells(i, 1).rowHeight
            Next i
            minX = shapeObj.left + shapeObj.width
            Set shapeObj = .Shapes.AddLine(minX + 2, minY + 2, minX + 2, endYLine)
            shapeObj.name = 1000
            With shapeObj.Line
                .DashStyle = msoLineDash
                .ForeColor.RGB = CCOLORNOWTIMERECT
                .Weight = 3
            End With
        End With
    End If
    Set layoutCurrentTimeRect = shapeObj
End Function

Sub fillCell(wks As Worksheet, row As Integer, col As Integer, Value As Variant, Optional bgColor As Long = 16777215, Optional orientation As Integer = xlLeft, Optional fontSize As Integer = CFONTSIZE, Optional isBold As Boolean = False)
   
    With wks.Cells(row, col)
        .Value = Value
        .Font.FontStyle = CFONT
        .Font.size = fontSize
        .Interior.color = bgColor
        .HorizontalAlignment = orientation
        .VerticalAlignment = xlTop
        .Font.Bold = isBold
    End With
End Sub

Function formatDate(time As Double, startTime As Double) As String
If year(CDate(time)) = year(CDate(startTime)) Then
    formatDate = Format(CDate(time), "dd.mm")
Else
    formatDate = Format(CDate(time), "dd.mm.yyyy")
End If
End Function

Function getMonth(ByVal currentMonth As Integer) As String
    If currentMonth > 12 Then
        currentMonth = currentMonth - 12
        getMonth = MonthName(month(DateValue("01-" & currentMonth & "-2000")))
    Else
        getMonth = MonthName(month(DateValue("01-" & currentMonth & "-2000")))
    End If
End Function

Function getMinX(wks As Worksheet, inTimeAble As InTimeAbleObj) As Double
    Dim colWidth As Double

    colWidth = convertCellWidth(wks.Cells(1, 1).ColumnWidth)
    getMinX = ((cSTARTCOL + CLEADINGFRAMEBORDERWIDTH - 2) + 1 * cTITLE1COLWIDTH) * colWidth + ((Day(inTimeAble.minTime) * colWidth) / 30)
End Function

Function convertCellWidth(cellWidth As Double) As Double
    convertCellWidth = 5.25 * cellWidth + 3.75
End Function



