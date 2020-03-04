Attribute VB_Name = "LayoutInPP"
Option Explicit
Const cWKSNAME1 As String = "Prozessbeschreibung"
Const cPPPNAME As String = "Swimlane"
Const CFONT As String = "Arial"
Const CFONTSIZE As Integer = 12
Const cFONTSIZETEXTBOX As Integer = 9

Const cBARCOLOR As Long = 14277081
Const cWHOBARCOLOR As Long = 10066329
Const cWHOBARLINECOLOR As Long = 10066329

Const cBARHEIGHT As Double = 45

Const cBARGAP As Double = 5
Const cWHOSWIMBARGAP As Double = 2
Const cWHOBARWIDTH As Double = 70
Const cSWIMBARWIDTH As Double = 1
Const cTEXTBOXGAP As Double = 5

Const cQUALITYCHECK As Long = 1907906
Const CCOLORNEWSLIDE As Long = 10213059
Const cTEXTBOXCOLOR As Long = 16777215

Sub layoutSwimlaneInPP(whichPresentation As Integer, removeBars As Boolean, Optional fileName As String = "")
    Dim pp As PowerPoint.Application
    Dim ppp As PowerPoint.Presentation
    Dim wks As Worksheet: Set wks = Worksheets(cWKSNAME1)
    Dim swimlane As New SwimlaneObj
    Dim swimlaneFrame As SwimlaneFrameObj
    
    Dim inputArray() As Variant
    Dim unnecessaryWhoBarsArray As Variant
    
    Dim minX As Double
    Dim minY As Double
    
    Set wks = Worksheets(cWKSNAME1)
    'title = wks.Range("Title")
    
    inputArray = gatherSwimlaneData(wks)
    swimlane.init "test", inputArray

    Set pp = New PowerPoint.Application
    If whichPresentation = 0 Then
        Set ppp = pp.Presentations.Add
    Else
        Set ppp = openPPFile(pp, fileName)
    End If

    With pp
        .WindowState = ppWindowMinimized
    End With
    
    minX = 0.03 * ppp.PageSetup.SlideWidth
    minY = 0.2 * ppp.PageSetup.SlideHeight
    Set swimlaneFrame = swimlane.swimlaneFrame
    unnecessaryWhoBarsArray = layoutTextBoxes(ppp, swimlane, minX, minY)
    If removeBars Then
        removeUnnecessaryWhoBars ppp, unnecessaryWhoBarsArray, swimlaneFrame.whoArray
    End If
    layoutArrows ppp, swimlane
    
    With pp
        .WindowState = ppWindowNormal
    End With
End Sub

Private Function layoutTextBoxes(ppp As PowerPoint.Presentation, swimlane As SwimlaneObj, minX As Double, minY As Double) As Variant
    Dim ppS As PowerPoint.Slide
    Dim maxPPSX As Double
    Dim slideCounter As Integer: slideCounter = 1
    Dim textBoxShape As PowerPoint.shape
    Dim endShape As PowerPoint.shape
    Dim swimlaneFrame As SwimlaneFrameObj
    Dim whoArray() As String
    Dim textBoxes() As Variant
    Dim swimlaneMatrix() As Variant
    
    Dim process As String
    Dim who As String
    Dim withWhom As String
    Dim qualityCheck As Boolean
    
    Dim xPos As Double
    Dim yPos As Double
    Dim width As Double
    Dim height As Double
    Dim length As Integer
    
    Dim yPosInSwimlane As Integer
    Dim lastYPosNeighbourInSwimlane As Integer
    Dim lastXPosInSwimlane As Integer
    Dim lastYPosInSwimlane As Integer
    Dim xPosInSwimlane As Integer
    Dim neighbourPosition As Integer
    Dim neighbourCoversNextOrLastTb As Boolean
    
    Dim i As Integer
    
    Dim necessaryWhoBarsCounter As Integer: necessaryWhoBarsCounter = 0
    Dim necessaryWhoBarsArray() As String
    Dim returnCounter As Integer
    Dim returnNecessaryWhoBarsArray() As Variant
    
    Set swimlaneFrame = swimlane.swimlaneFrame
    Set ppS = ppp.Slides.AddSlide(slideCounter, ppp.SlideMaster.CustomLayouts(2))
    
    whoArray = swimlaneFrame.whoArray
    
    textBoxes = swimlane.textBoxes
    swimlaneMatrix = swimlane.swimlaneMatrix
    length = UBound(textBoxes)
    xPos = minX + cWHOBARWIDTH + cWHOSWIMBARGAP
    
    maxPPSX = ppp.PageSetup.SlideWidth - 5
    
    ReDim necessaryWhoBarsArray(0 To 2 * length + 1)      '*2 to acommodate for possible withWhoms - oversize will be removed
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
        
        xPos = xPos + (xPosInSwimlane - lastXPosInSwimlane) * (width + cTEXTBOXGAP)         '(xPosInSwimlane - lastXPosInSwimlane) = {0,1}
        
        neighbourCoversNextOrLastTb = (lastYPosInSwimlane = (yPosInSwimlane + neighbourPosition) Or yPosInSwimlane = (lastYPosInSwimlane + lastYPosNeighbourInSwimlane)) And xPosInSwimlane = lastXPosInSwimlane
        If neighbourCoversNextOrLastTb Then
            xPos = xPos + width + cTEXTBOXGAP
        End If
        
        Set textBoxShape = layoutPPTextBox(ppS, xPos, yPos, height, process)
        
        With textBoxShape
            .name = "SwimTextBox" & i
            .Fill.ForeColor.RGB = cTEXTBOXCOLOR
            .Line.ForeColor.RGB = cTEXTBOXCOLOR
            If .width < width And (xPosInSwimlane - lastXPosInSwimlane) = 0 Then
                xPos = xPos + width - .width
                .left = xPos
            End If
            width = .width
            If qualityCheck Then
                .Line.Weight = 3
                .Line.ForeColor.RGB = cQUALITYCHECK
            End If
        End With
        If maxPPSX < xPos + width Then
            textBoxShape.Delete
            layoutFrame ppS, swimlaneFrame, minX, minY, maxPPSX
            slideCounter = slideCounter + 1
            Set ppS = ppp.Slides.AddSlide(slideCounter, ppp.SlideMaster.CustomLayouts(2))
            xPos = minX + cWHOBARWIDTH + cWHOSWIMBARGAP + 10     'make constant + 2
            Set textBoxShape = layoutPPTextBox(ppS, xPos, yPos, height, process)
            With textBoxShape
                .name = "SwimTextBox" & i
                .Fill.ForeColor.RGB = cTEXTBOXCOLOR
                .Line.ForeColor.RGB = cTEXTBOXCOLOR
            End With
            
            ReDim Preserve returnNecessaryWhoBarsArray(0 To returnCounter)
            returnNecessaryWhoBarsArray(returnCounter) = uniqueInArray(necessaryWhoBarsArray)
            necessaryWhoBarsArray = emptyArray(necessaryWhoBarsArray)
            necessaryWhoBarsCounter = 0
            returnCounter = returnCounter + 1
        End If
        lastYPosNeighbourInSwimlane = neighbourPosition
        lastYPosInSwimlane = yPosInSwimlane
        lastXPosInSwimlane = xPosInSwimlane
        
        necessaryWhoBarsCounter = necessaryWhoBarsCounter + 2
        necessaryWhoBarsArray(necessaryWhoBarsCounter - 2) = who
        necessaryWhoBarsArray(necessaryWhoBarsCounter - 1) = withWhom
    Next i
    layoutFrame ppS, swimlaneFrame, minX, minY, maxPPSX
    
    With textBoxShape
        Set endShape = ppS.Shapes.AddShape(msoShapeFlowchartAlternateProcess, .left + .width - 4, .top + 0.3 * .height, 0.4 * .height, 0.4 * .height)
        layoutTextBoxes = .left + .width + 0.5 * .height + 4
    End With
    With endShape
        .name = "otherObj"
        .Fill.ForeColor.RGB = cQUALITYCHECK
        .Line.ForeColor.RGB = cQUALITYCHECK
    End With
    ReDim Preserve returnNecessaryWhoBarsArray(0 To returnCounter)
    returnNecessaryWhoBarsArray(returnCounter) = uniqueInArray(necessaryWhoBarsArray)
    layoutTextBoxes = returnNecessaryWhoBarsArray
End Function

Private Function layoutPPTextBox(ppS As PowerPoint.Slide, xPos As Double, yPos As Double, nHeight As Double, nText As String) As PowerPoint.shape
    Dim shapeObj As PowerPoint.shape
    
    Dim condition As Double: condition = True
    Dim exitCondition As Integer: exitCondition = 0
    Dim sizeFactor As Double
    nText = removeLinebreaks(nText)
    
    sizeFactor = 1  'use 1, performance vs accuracy -> higher value more accuracy, les performance
    Set shapeObj = ppS.Shapes.AddTextbox(msoTextOrientationHorizontal, xPos, yPos, 10, nHeight)
    With shapeObj
       
        .TextFrame.TextRange.Characters.text = nText
        .TextEffect.FontName = CFONT
        .TextFrame.TextRange.Characters.Font.size = cFONTSIZETEXTBOX

        Do While condition And exitCondition <= 1000

            .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
            If .height > nHeight Then
                .width = .width + 10 / sizeFactor
                condition = True
            Else

                condition = False
            End If
            exitCondition = exitCondition + 1
        Loop
        .height = nHeight
    End With

    Set layoutPPTextBox = shapeObj
End Function

Private Sub layoutFrame(ppS As PowerPoint.Slide, swimlaneFrame As SwimlaneFrameObj, minX As Double, minY As Double, maxX As Double)
    Dim shapeObj As PowerPoint.shape
    Dim i As Integer
    Dim length As Integer
    Dim xPos As Double
    Dim yPos As Double
    Dim width As Double
    Dim height As Double
    Dim whoArray() As String
    
    length = swimlaneFrame.whoBarsCount - 1
    whoArray = swimlaneFrame.whoArray
    xPos = minX
    
    width = cWHOBARWIDTH
    height = cBARHEIGHT
    For i = 0 To length
        yPos = minY + i * (cBARGAP + height)
        Set shapeObj = ppS.Shapes.AddTextbox(msoTextOrientationHorizontal, xPos, yPos, width, height)
        With shapeObj
            .TextFrame.TextRange.Characters.text = whoArray(i)
            .TextEffect.FontName = CFONT
            .TextFrame.TextRange.Characters.Font.size = CFONTSIZE
            .TextFrame2.WordWrap = True
            .height = height
            .Fill.ForeColor.RGB = cWHOBARCOLOR
            .Line.Weight = 2.25
            .Line.ForeColor.RGB = cWHOBARLINECOLOR
            .name = "WhoBar" & i
        End With
        
        Set shapeObj = ppS.Shapes.AddShape(msoShapeRectangle, xPos + cWHOBARWIDTH + cWHOSWIMBARGAP, yPos, maxX - xPos - cWHOBARWIDTH - cWHOSWIMBARGAP, height)
        With shapeObj
            .Fill.ForeColor.RGB = cBARCOLOR
            .Line.ForeColor.RGB = cBARCOLOR
            .ZOrder msoSendToBack
            .name = "Bar" & i
        End With
    Next i

End Sub

Private Sub removeUnnecessaryWhoBars(ppp As PowerPoint.Presentation, necessaryWhoBarsArray As Variant, whoArray As Variant)
    Dim ppS As PowerPoint.Slide
    Dim whoBarToRemove As PowerPoint.shape
    Dim barToRemove As PowerPoint.shape
    Dim shapeObj As PowerPoint.shape
    Dim slideCount As Integer
    Dim length As Integer
    
    Dim necessaryWhoBars As Variant
    Dim height As Double
    Dim top As Double
    
    Dim regexSwim As New RegExp
    Dim regexBar As New RegExp
    Dim regexWho As New RegExp
    Dim regexOtherObj As New RegExp
    
    Dim i As Integer
    Dim j As Integer
    
    regexSwim.Pattern = "SwimTextBox"
    regexBar.Pattern = "Bar"
    regexWho.Pattern = "WhoBar"
    regexOtherObj.Pattern = "otherObj"
    slideCount = UBound(necessaryWhoBarsArray)
    
    For i = 0 To slideCount
        necessaryWhoBars = necessaryWhoBarsArray(i)
        length = UBound(whoArray)
        Set ppS = ppp.Slides(i + 1)
        For j = 0 To length
            If inArray(necessaryWhoBars, whoArray(j)) = -1 Then
                Set whoBarToRemove = getShapeByName(ppS, "WhoBar" & j)
                Set barToRemove = getShapeByName(ppS, "Bar" & j)
                height = barToRemove.height
                top = barToRemove.top
                whoBarToRemove.Delete
                barToRemove.Delete
                For Each shapeObj In ppS.Shapes
                    If shapeObj.top > top And (regexSwim.Test(shapeObj.name) Or regexBar.Test(shapeObj.name) Or regexWho.Test(shapeObj.name) Or regexOtherObj.Test(shapeObj.name)) Then
                        shapeObj.top = shapeObj.top - height - cBARGAP
                    End If
                Next shapeObj
            End If
        Next j
    Next i
End Sub

Private Sub layoutArrows(ppp As PowerPoint.Presentation, swimlane As SwimlaneObj)
    Dim ppS As PowerPoint.Slide
    Dim slideCounter As Integer: slideCounter = 1
    Dim arrowArray() As Variant
    Dim arrowXDirection As Integer
    Dim arrowYDirection As Integer
    Dim arrowForm As Integer
    
    Dim arrowShape As PowerPoint.shape
    Dim arrowShape2 As PowerPoint.shape
    
    Dim tb0 As PowerPoint.shape
    Dim tb1 As PowerPoint.shape
    
    Dim barShape As PowerPoint.shape
    
    Dim arrowWay() As Integer
    Dim connectFirst As Integer
    Dim connectSecond As Integer
    Dim lastConnectSecond As Integer
    
    Dim length As Integer
    Dim i As Integer
    
    arrowArray = swimlane.arrowArray
    length = UBound(arrowArray)
    
    Set ppS = ppp.Slides(slideCounter)
    Set tb0 = getShapeByName(ppS, "SwimTextBox0")
        For i = 0 To length
            With ppS
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
                If arrowForm = 0 Then
                    Set arrowShape = .Shapes.AddConnector(msoConnectorStraight, 0, 0, 0, 0)
                Else
                    Set arrowShape = .Shapes.AddConnector(msoConnectorElbow, 0, 0, 0, 0)
                End If
                
                With arrowShape
                    
                    Set tb1 = getShapeByName(ppS, "SwimTextBox" & i + 1)
                    If tb1 Is Nothing Then
                        .ConnectorFormat.BeginConnect tb0, 4
                        If i < length Then
                            Set barShape = getShapeByName(ppS, "Bar" & arrowArray(i + 1, 4))
                            If barShape Is Nothing Then
                                Set barShape = getShapeByName(ppS, "Bar" & arrowArray(i, 4))
                            End If
                            .ConnectorFormat.EndConnect barShape, 4
                            
                            If arrowForm <> 0 Then
                                .Adjustments.Item(1) = 0.2408
                            End If
                            Set barShape = Nothing
                        End If
                        
                        .Line.ForeColor.RGB = CCOLORNEWSLIDE
                        .Line.Weight = 3
                        
                        slideCounter = slideCounter + 1
                        Set ppS = ppp.Slides(slideCounter)
                        Set arrowShape2 = ppS.Shapes.AddConnector(msoConnectorStraight, 0, 0, 0, 0)
                        With arrowShape2
                            Set tb1 = getShapeByName(ppS, "SwimTextBox" & i + 1)
                            .ConnectorFormat.BeginConnect getShapeByName(ppS, "WhoBar" & arrowArray(i + 1, 4)), 4
                            .ConnectorFormat.EndConnect tb1, 2
                            .Line.ForeColor.RGB = CCOLORNEWSLIDE
                            .Line.Weight = 3
                            .Line.EndArrowheadStyle = msoArrowheadTriangle
                        End With
                    Else
                        .ConnectorFormat.BeginConnect tb0, connectFirst
                        If lastConnectSecond = connectFirst Then
                            .left = .left + 5
                        End If
                        .ConnectorFormat.EndConnect tb1, connectSecond
                        .Line.ForeColor.ObjectThemeColor = msoThemeColorText1
                        If arrowForm = 0 And arrowXDirection = 0 Then       'only if arrow goes top or down
                            If tb0.width <> tb1.width Then
                                .ScaleWidth 0, msoFalse
                                .ConnectorFormat.EndDisconnect
                                If tb1.width - 5 <= tb0.width / 2 Then
                                    .left = .left + 0.5 * (tb0.width - tb1.width)
                                End If
                            End If
                        End If
                    End If
                    .Line.EndArrowheadStyle = msoArrowheadTriangle
                End With
             End With
             lastConnectSecond = connectSecond
             Set tb0 = tb1
        Next i
    
End Sub

Private Function getShapeByName(ppS As PowerPoint.Slide, name As String) As PowerPoint.shape
    Dim shape As PowerPoint.shape
    For Each shape In ppS.Shapes
        If shape.name = name Then
            Set getShapeByName = shape
            Exit Function
        End If
    Next shape
End Function
