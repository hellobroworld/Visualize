Attribute VB_Name = "LayoutSwimlane"
Option Explicit
Public Const cWKSNAME1 As String = "Process Description"
Public Const cWKSNAMESWIMLANE As String = "Swimlane"
Public Const cTABLE As String = "ProcessTable"
Public Const CTITLE0COL As String = "#"
Public Const CTITLE1COL As String = "Process Step"
Public Const CTITLE2COL As String = "Who (Responsible Person)"
Public Const CTITLE3COL As String = "With Whom?"
Public Const CTITLE4COL As String = "Quality-Check"
Public Const CDATECOL As String = "When?"

Public Const CFONT As String = "Arial"
Public Const CFONTSIZE As Integer = 12
Const cFONTSIZETEXTBOX As Integer = 9
Const cROWHEIGHT As Double = 15
Const cCOLWIDTH As Double = 15
Const cSTARTROW As Integer = 2
Const cSTARTCOL As Integer = 1

Const cQUALITYCHECK As Long = 1907906
Const cBARCOLOR As Long = 14277081
Const cWHOBARLINECOLOR As Long = 12566463

Const cBARHEIGHT As Double = 45
Const cBARGAP As Double = 5
Const cWHOSWIMBARGAP As Double = 2
Const cWHOBARWIDTH As Double = 15
Const cSWIMBARWIDTH As Double = 1
Const cTEXTBOXGAP As Double = 5

Sub LayoutSwimlane()
    Dim wks As Worksheet: Set wks = Worksheets(cWKSNAME1)
    Dim swimlane As New SwimlaneObj
    Dim swimlaneFrame As SwimlaneFrameObj
    
    Dim inputArray() As Variant
    Dim maxX As Double
    
    If WorksheetExists(cWKSNAMESWIMLANE) Then
        MsgBox "Swimlane does already exist. Please rename or delete the worksheet containing the timetable"
        End
    End If
    
    Set wks = Worksheets(cWKSNAME1)
    'title = wks.Range("Title")
    
    inputArray = gatherSwimlaneData(wks)
    swimlane.init "test", inputArray

    Set wks = Worksheets.Add
    With wks
        .name = cWKSNAMESWIMLANE
        .Cells.rowHeight = cROWHEIGHT
        .Cells.ColumnWidth = cCOLWIDTH
        .Cells.Font.FontStyle = CFONT
        .Cells.Interior.color = 16777215
    End With
    
    Set swimlaneFrame = swimlane.swimlaneFrame
    maxX = layoutTextBoxes(wks, swimlane)
    layoutFrame wks, swimlaneFrame, maxX
    layoutArrows wks, swimlane

End Sub

Private Function layoutTextBoxes(wks As Worksheet, swimlane As SwimlaneObj) As Double
    Dim textBoxShape As shape
    Dim endShape As shape
    
    Dim swimlaneFrame As SwimlaneFrameObj
    Dim whoArray() As String
    Dim textBoxes() As Variant
    Dim swimlaneMatrix() As Variant
    
    Dim process As String
    Dim who As String
    Dim withWhom As String
    Dim qualityCheck As Boolean
    
    Dim minX, minY As Double
    Dim xPos As Double
    Dim yPos As Double
    Dim width As Double
    Dim height As Double
    Dim length As Integer
    Dim numberOfLines() As Integer 'if hasNeighbour -> height *2
    
    Dim yPosInSwimlane As Integer
    Dim lastYPosNeighbourInSwimlane As Integer
    Dim lastXPosInSwimlane As Integer
    Dim lastYPosInSwimlane As Integer
    Dim xPosInSwimlane As Integer
    Dim neighbourPosition As Integer
    Dim neighbourCoversNextOrLastTb As Boolean
    Dim i As Integer
    
    Set swimlaneFrame = swimlane.swimlaneFrame
    whoArray = swimlaneFrame.whoArray
    textBoxes = swimlane.textBoxes
    swimlaneMatrix = swimlane.swimlaneMatrix
    length = UBound(textBoxes)
    
    minX = convertCellWidth(cWHOBARWIDTH) + cSTARTCOL * convertCellWidth(cCOLWIDTH) + cWHOSWIMBARGAP
    minY = cSTARTROW * cROWHEIGHT
    xPos = minX
    numberOfLines = getTextBoxLineCount(wks, cBARHEIGHT, 2 * cBARHEIGHT)
    For i = 0 To length
        process = textBoxes(i, 1)
        who = textBoxes(i, 2)
        withWhom = textBoxes(i, 3)
        yPosInSwimlane = textBoxes(i, 4)
        qualityCheck = check(textBoxes(i, 5), "String") <> ""
        
        xPosInSwimlane = getXPosInMatrix(swimlaneMatrix, yPosInSwimlane, i)
        neighbourPosition = swimlane.isNeighbour(whoArray, yPosInSwimlane, withWhom)
        
        yPos = minY + (yPosInSwimlane + (neighbourPosition - 1) / 2 * Abs(neighbourPosition)) * (cBARHEIGHT + cBARGAP)     'yeah it's more complicated than using some if conditions - but i was bored and it's prettier ;)
        height = (cBARHEIGHT) * (1 + Abs(neighbourPosition))                                         'the essence of this and the last step is to set the right height and position considering if the responsible person performs the task with someone
        
        xPos = xPos + (xPosInSwimlane - lastXPosInSwimlane) * (width + cTEXTBOXGAP)          '(xPosInSwimlane - lastXPosInSwimlane) = {0,1}
        
        neighbourCoversNextOrLastTb = (lastYPosInSwimlane = (yPosInSwimlane + neighbourPosition) Or yPosInSwimlane = (lastYPosInSwimlane + lastYPosNeighbourInSwimlane)) And xPosInSwimlane = lastXPosInSwimlane
        If neighbourCoversNextOrLastTb Then
            xPos = xPos + width + cTEXTBOXGAP
        End If
        
        Set textBoxShape = layoutTextBox(wks, xPos, yPos, height, process, numberOfLines(Abs(neighbourPosition)))
        
        With textBoxShape
            .name = "SwimTextBox" & i
            If .width < width And (xPosInSwimlane - lastXPosInSwimlane) = 0 Then
                xPos = xPos + width - .width
                .left = xPos
            End If
            If qualityCheck Then
                .Line.Weight = 3
                .Line.ForeColor.RGB = cQUALITYCHECK
            End If
        End With
        width = textBoxShape.width
        
        lastYPosNeighbourInSwimlane = neighbourPosition
        lastYPosInSwimlane = yPosInSwimlane
        lastXPosInSwimlane = xPosInSwimlane
    Next i
    
    With textBoxShape
        Set endShape = wks.Shapes.AddShape(msoShapeFlowchartAlternateProcess, .left + .width - 4, .top + 0.25 * .height, 0.5 * .height, 0.5 * .height)
        layoutTextBoxes = .left + .width + 0.5 * .height + 4
    End With
    With endShape
        .Fill.ForeColor.RGB = cQUALITYCHECK
        .Line.ForeColor.RGB = cQUALITYCHECK
    End With
End Function

Function getXPosInMatrix(swimlaneMatrix As Variant, rank As Integer, step As Integer) As Integer
    Dim width As Integer
    Dim i As Integer
    
    width = UBound(swimlaneMatrix, 2)
    
    For i = 0 To width
        If swimlaneMatrix(rank, i) = step Then
            getXPosInMatrix = i
            Exit Function
        End If
    Next i
End Function

Private Function getTextBoxLineCount(wks, nHeight1, nHeight2) As Integer()
    Dim shapeObj As shape
    Dim marginsHeight As Double

    Dim returnArray(0 To 1) As Integer
    
    Set shapeObj = wks.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 20, nHeight1)
    With shapeObj
        marginsHeight = .TextFrame.MarginTop + .TextFrame.MarginTop + 3     'should account for the space between two lines - please finde a better way to get fitting textbox width..
        returnArray(0) = (nHeight1 - marginsHeight) / cFONTSIZETEXTBOX
        returnArray(1) = (nHeight2 - marginsHeight) / cFONTSIZETEXTBOX
        .Delete
    End With
    
    getTextBoxLineCount = returnArray
End Function

Private Sub layoutArrows(wks As Worksheet, swimlane As SwimlaneObj)
    Dim arrowArray() As Variant
    Dim arrowXDirection As Integer
    Dim arrowYDirection As Integer
    Dim arrowForm As Integer
    
    Dim arrowShape As shape
    
    Dim tb0 As shape
    Dim tb1 As shape
    
    Dim arrowWay() As Integer
    Dim connectFirst As Integer
    Dim connectSecond As Integer
    Dim lastConnectSecond As Integer
    
    Dim length As Integer
    Dim i As Integer
    
    arrowArray = swimlane.arrowArray
    length = UBound(arrowArray)
    Set tb0 = getShapeByName(wks, "SwimTextBox0")
    With wks
        For i = 0 To length
            arrowYDirection = arrowArray(i, 0)
            arrowXDirection = arrowArray(i, 1)
            arrowForm = arrowArray(i, 3)
            
            If arrowXDirection = 0 Then
                connectFirst = 2 + arrowYDirection
                connectSecond = 2 - arrowYDirection
            Else
                arrowWay = arrowArray(i, 2)
                connectFirst = arrowWay(0)
                connectSecond = arrowWay(1)
            End If
            
            Set tb1 = getShapeByName(wks, "SwimTextBox" & i + 1)
            If arrowForm = 0 Then
                Set arrowShape = .Shapes.AddConnector(msoConnectorStraight, 0, 0, 0, 0)

            Else
                Set arrowShape = .Shapes.AddConnector(msoConnectorElbow, 0, 0, 0, 0)
            End If
            
            With arrowShape
                .ConnectorFormat.BeginConnect tb0, connectFirst
                If lastConnectSecond = connectFirst Then
                    .left = .left + 5
                End If
                .ConnectorFormat.EndConnect tb1, connectSecond
                .Line.ForeColor.ObjectThemeColor = msoThemeColorText1
                .Line.EndArrowheadStyle = msoArrowheadTriangle
                
                If arrowForm = 0 And arrowXDirection = 0 Then       'only if arrow goes top down
                    If tb0.width <> tb1.width Then
                        .ScaleWidth 0, msoFalse
                        .ConnectorFormat.EndDisconnect
                        If tb1.width - 5 <= tb0.width / 2 Then
                            .left = .left + 0.5 * (tb0.width - tb1.width)
                        End If
                    End If
                End If
            End With
            Set tb0 = tb1
            lastConnectSecond = connectSecond
        Next i
    End With
    
End Sub

Private Sub layoutFrame(wks As Worksheet, swimlaneFrame As SwimlaneFrameObj, maxX As Double)
    Dim shapeObj As shape
    Dim i As Integer
    Dim length As Integer
    Dim xPos As Double
    Dim yMin As Double
    Dim yPos As Double
    Dim width As Double
    Dim height As Double
    Dim whoArray() As String
    
    length = swimlaneFrame.whoBarsCount - 1
    whoArray = swimlaneFrame.whoArray
    xPos = convertCellWidth(cCOLWIDTH) * cSTARTCOL
    yMin = cSTARTROW * cROWHEIGHT
    
    width = convertCellWidth(cWHOBARWIDTH)
    height = cBARHEIGHT
    For i = 0 To length
        yPos = yMin + i * (cBARGAP + height)
        Set shapeObj = wks.Shapes.AddTextbox(msoTextOrientationHorizontal, xPos, yPos, width, height)
        With shapeObj
            .TextFrame.Characters.text = whoArray(i)
            .TextEffect.FontName = CFONT
            .TextFrame.Characters.Font.size = CFONTSIZE
            .TextFrame2.WordWrap = True
            .Fill.ForeColor.RGB = cBARCOLOR
            .Line.ForeColor.RGB = cWHOBARLINECOLOR
            .Line.Weight = 2.25
        End With
        
        Set shapeObj = wks.Shapes.AddShape(msoShapeRectangle, xPos + width + cWHOSWIMBARGAP, yPos, maxX - xPos - width - cWHOSWIMBARGAP, height)
        With shapeObj
            .Fill.ForeColor.RGB = cBARCOLOR
            .Line.ForeColor.RGB = cBARCOLOR
            .ZOrder msoSendToBack
        End With
    Next i

End Sub

Function gatherSwimlaneData(wks As Worksheet) As Variant
    Dim listObj As ListObject
    Dim inputArray As Variant
    Dim returnNx4InputArray As Variant
    Dim title0Col, title1Col, title2Col, title3Col, title4Col As Integer
    Dim lb As Integer
    Dim length, i As Integer
    
    Set listObj = wks.ListObjects(cTABLE)

    With listObj
        inputArray = .DataBodyRange
        title0Col = .ListColumns(CTITLE0COL).index
        title1Col = .ListColumns(CTITLE1COL).index
        title2Col = .ListColumns(CTITLE2COL).index
        title3Col = .ListColumns(CTITLE3COL).index
        title4Col = .ListColumns(CTITLE4COL).index
        
        lb = LBound(inputArray, 2)
        inputArray = arraysStartWith0(inputArray)
        inputArray = removeNARowsFromMatrix(inputArray)
        If isNA(inputArray) Then
            MsgBox "Please enter data in " & cTABLE
            End
        End If
        length = UBound(inputArray, 1)
        ReDim returnNx4InputArray(0 To length, 0 To 4)
        For i = 0 To length
            returnNx4InputArray(i, 0) = inputArray(i, title0Col - lb)   'Step
            returnNx4InputArray(i, 1) = inputArray(i, title1Col - lb)   'Process
            returnNx4InputArray(i, 2) = inputArray(i, title2Col - lb)   'Who
            returnNx4InputArray(i, 3) = inputArray(i, title3Col - lb)   'With Whom
            returnNx4InputArray(i, 4) = inputArray(i, title4Col - lb)   'quality check
        Next i
    End With
    
    gatherSwimlaneData = returnNx4InputArray
End Function

Private Function getShapeByName(wks As Worksheet, name As String) As shape
    Dim shape As shape
    For Each shape In wks.Shapes
        If shape.name = name Then
            Set getShapeByName = shape
            Exit Function
        End If
    Next shape
End Function
