VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SwimlaneObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pTitle As String
Private pSwimlaneFrame As SwimlaneFrameObj
Private pTextBoxes() As Variant
Private pSwimlaneMatrix() As Variant
Private pArrowArray() As Variant

Const cTOP = -1
Const cBOTTOM = 1
Const cSTAY = 0
Const cRIGHT = 1
Const cSTREIGHT = 0
Const cCROOKED = 1

Property Get title() As String
    title = pTitle
End Property

Property Get swimlaneFrame() As SwimlaneFrameObj
    Set swimlaneFrame = pSwimlaneFrame
End Property

Property Get textBoxes() As Variant
    textBoxes = pTextBoxes
End Property

Property Get swimlaneMatrix() As Variant
    swimlaneMatrix = pSwimlaneMatrix
End Property

Property Get arrowArray() As Variant
    arrowArray = pArrowArray
End Property

Sub init(nTitle As String, nX4Array As Variant)
    Dim nSwimlaneFrame As New SwimlaneFrameObj
    Dim frameArray() As String
    Dim length As Integer
    Dim i As Integer
    
    pTitle = nTitle
    length = UBound(nX4Array)
    ReDim frameArray(0 To length)
    
    For i = 0 To length
        frameArray(i) = nX4Array(i, 2)
    
    Next i
    
    nSwimlaneFrame.init frameArray
    
    Set pSwimlaneFrame = nSwimlaneFrame
    
    pTextBoxes = filTextBoxes(nX4Array, pSwimlaneFrame.whoArray)
    pSwimlaneMatrix = createSwimlaneMatrix(pTextBoxes, pSwimlaneFrame.whoArray)
    pArrowArray = createArrowArray(pTextBoxes, pSwimlaneMatrix)
End Sub

Private Function filTextBoxes(nX4Array As Variant, whoArray As Variant) As Variant

    Dim length1 As Integer
    Dim length2 As Integer
    Dim rank As Integer
    Dim i, j As Integer

    Dim returnArray As Variant
    
    length1 = UBound(nX4Array)
    length2 = UBound(whoArray)
    
    ReDim returnArray(0 To length1, 0 To 5)
    For i = 0 To length1
        For j = 0 To length2
            If nX4Array(i, 2) = whoArray(j) Then
                rank = j
            End If
        Next j
        returnArray(i, 0) = nX4Array(i, 0)  'Step
        returnArray(i, 1) = nX4Array(i, 1)  'Process
        returnArray(i, 2) = nX4Array(i, 2)  'Who
        returnArray(i, 3) = nX4Array(i, 3)  'with whom
        returnArray(i, 4) = rank
        returnArray(i, 5) = nX4Array(i, 4)  'quality check
    Next i

    filTextBoxes = returnArray
End Function

Private Function createSwimlaneMatrix(textBoxes As Variant, whoArray() As String) As Variant
    Dim length1 As Integer
    Dim yPos0 As Integer
    Dim yPosm1 As Integer       'm = - -> xPosm1 = xPos -1, xPos vom letztem Element
    Dim yPosm2 As Integer
    Dim xPos As Integer
    Dim xPosm1 As Integer
    Dim xPosm2 As Integer
    
    Dim lengthWhoArray As Integer
    
    Dim i As Integer
    
    Dim returnArray() As Variant
    
    lengthWhoArray = UBound(whoArray)
    ReDim returnArray(0 To lengthWhoArray, 0 To 100)
    length1 = UBound(textBoxes, 1)
    
    For i = 0 To length1
        yPos0 = textBoxes(i, 4)
         
        If i > 1 Then
            yPosm1 = textBoxes(i - 2, 4)
            yPosm2 = textBoxes(i - 1, 4)

            'Debug.Print textBoxes(i - 1, 1) 'delete
            xPos = xPositionTextBox(yPosm1, yPosm2, yPos0, returnArray, xPos, xPosm1, xPosm2)
        Else
            If Not IsEmpty(returnArray(yPos0, xPos)) Then
                xPos = xPos + 1
            End If
        End If
        
        If UBound(returnArray, 2) < xPos Then
            returnArray = reDimNxNArray(returnArray, lengthWhoArray, UBound(returnArray, 2) + 20)
        Else
            
            returnArray(yPos0, xPos) = i
            'neighbourPos = isNeighbour(whoArray, yPos0, CStr(textBoxes(i, 3))) 'would add i in both who and with whom position - but doesnt work good with layout logic
            'returnArray(yPos0 + neighbourPos, xPos) = i
        End If
        xPosm2 = xPosm1
        xPosm1 = xPos
    Next i
    
    createSwimlaneMatrix = reDimNxNArray(returnArray, lengthWhoArray, xPos)
    
End Function

Private Function xPositionTextBox(yPosm1 As Integer, yPosm2 As Integer, yPos0 As Integer, nSwimlaneMatrix As Variant, xPos0 As Integer, xPosm1 As Integer, xPosm2 As Integer) As Integer
    Dim pos1 As Integer
    Dim pos2 As Integer
    
    If Not IsEmpty(nSwimlaneMatrix(yPos0, xPos0)) Then
        xPositionTextBox = xPos0 + 1
    'ElseIf xPosm2 < xPosm1 And IsEmpty(nSwimlaneMatrix(yPos0, xPos0)) Then 'would accomodate for x axes, but arrows might make some appearance problems
        'xPositionTextBox = xPos0
    Else
        pos1 = yPosm1 - yPos0
        pos2 = yPosm1 - yPosm2
        If pos1 = 0 Or pos2 = 0 Then
            xPositionTextBox = xPos0
        ElseIf Abs(pos1) > Abs(pos2) And pos1 / Abs(pos1) = pos2 / Abs(pos2) Then           'to determine if textbox has to be shifted one to the right
            xPositionTextBox = xPos0
        Else
            xPositionTextBox = xPos0 + 1
        End If
    End If
End Function

Function isNeighbour(whoArray() As String, whoIndex As Integer, withWhom As String) As Integer      'return: {-1, 0, 1}
    Dim length As Integer
    Dim i As Integer
    Dim index2 As Integer
    index2 = inArray(whoArray, withWhom)
    
    If whoIndex = index2 + 1 And index2 <> -1 Then
        isNeighbour = -1
    ElseIf whoIndex = index2 - 1 Then
        isNeighbour = 1
    Else
        isNeighbour = 0
    End If
End Function

Function createArrowArray(textBoxes As Variant, nSwimlaneMatrix As Variant) As Variant    'return: 0 -> bottom/top, 1 -> left/right, 2 -> arrowWay(), 3 -> form, 4 -> y pos of textbox
    Dim yPos1 As Integer
    Dim yPos0 As Integer
    Dim xPos1 As Integer
    Dim xPos0 As Integer
    
    Dim direction() As Variant
    Dim yAxis As Integer
    Dim xAxis As Integer
    Dim yDirection As Integer
    Dim width As Integer

    Dim j As Integer

    width = UBound(textBoxes, 1)

    ReDim direction(0 To width - 1, 0 To 4)
    For j = 1 To width
        
        yPos0 = textBoxes(j - 1, 4)
        yPos1 = textBoxes(j, 4)
        xPos0 = getXPosInMatrix(nSwimlaneMatrix, yPos0, j - 1)  'pos of current textbox
        xPos1 = getXPosInMatrix(nSwimlaneMatrix, yPos1, j)      'pos of next textbox
        
        direction(j - 1, 4) = yPos0
        
        yAxis = yPos1 - yPos0
        xAxis = xPos1 - xPos0
        
        If yAxis > 0 Then
            yDirection = cBOTTOM
        ElseIf yAxis < 0 Then
            yDirection = cTOP
        Else
            yDirection = cSTAY
        End If
        direction(j - 1, 0) = yDirection
        If xAxis > 0 Then
            direction(j - 1, 1) = cRIGHT
            'Debug.Print textBoxes(j - 1, 1) 'delete
            direction(j - 1, 2) = findArrowWay(xPos0, yPos1, yPos0, yDirection, nSwimlaneMatrix)
        Else
            direction(j - 1, 1) = cSTAY
        End If
        
        If direction(j - 1, 1) = cRIGHT And direction(j - 1, 0) <> cSTAY Then
            direction(j - 1, 3) = cCROOKED
        Else
            direction(j - 1, 3) = cSTREIGHT
        End If
    Next j

    createArrowArray = direction
End Function

Private Function findArrowWay(xPos As Integer, yPos1 As Integer, yPos As Integer, arrowYDirection As Integer, swimlaneMatrix As Variant) As Integer()
    Dim topOk As Boolean: topOk = True
    Dim bottomOk As Boolean: bottomOk = True
    Dim sideOk As Boolean: sideOk = False
    Dim length As Integer
    Dim i As Integer
    Dim connectorBegin As Integer
    Dim connectorEnd As Integer
    
    Dim returnArray(0 To 1) As Integer
    
    If arrowYDirection = 1 Then
        length = UBound(swimlaneMatrix, 1)
        For i = yPos + 1 To length
            If Not IsEmpty(swimlaneMatrix(i, xPos)) Then
                bottomOk = False
            End If
        Next i
        If bottomOk Then
            connectorBegin = 3      'Bottom
            connectorEnd = 2        'Left
            sideOk = True
        End If
    ElseIf arrowYDirection = -1 Then
        For i = 0 To yPos - 1
            If Not IsEmpty(swimlaneMatrix(i, xPos)) Then
                topOk = False
            End If
        Next i
        If topOk Then
            connectorBegin = 1      'Top
            connectorEnd = 2        'Left
            sideOk = True
        End If
    End If
    If Not topOk Or Not bottomOk Then
        topOk = True
        bottomOk = True
        If arrowYDirection = 1 Then
           For i = 0 To yPos1 - 1
               If Not IsEmpty(swimlaneMatrix(i, xPos + 1)) Then
                   bottomOk = False
               End If
           Next i
           If bottomOk Then
               connectorBegin = 4   'Right
               connectorEnd = 1     'Top
               sideOk = True
           End If
        ElseIf arrowYDirection = -1 Then
            For i = yPos1 + 1 To length
               If Not IsEmpty(swimlaneMatrix(i, xPos + 1)) Then
                   topOk = False
               End If
            Next i
            If topOk Then
               connectorBegin = 4   'Rigth
               connectorEnd = 3     'Bottom
               sideOk = True
            End If
        End If
    End If
    If Not sideOk Then
        connectorBegin = 4      'Right
        connectorEnd = 2        'Left
    End If
    
    returnArray(0) = connectorBegin
    returnArray(1) = connectorEnd
    
    findArrowWay = returnArray
End Function

