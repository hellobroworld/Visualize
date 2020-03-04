VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeRectObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const CCOLOR1 As Long = 8355711
Const CCOLOR2 As Long = 1907906
Const CCOLORNOWTIMERECT As Long = 10213059

Private pNx3SizesArray As Variant

Property Get nX3SizesArray() As Variant      '0: Form, 1: xMinPos, 2: width, 3: color
    nX3SizesArray = pNx3SizesArray
End Property

Sub init(minTime As Double, timeArray() As String)
    Dim length As Integer
    Dim returnArray As Variant
  
    Dim value1 As Double
    Dim value2 As Double
    Dim xMin1 As Double
    Dim xMin2 As Double
    Dim width As Double
    Dim form As String
    Dim color As Long
    
    Dim condition As Boolean
    Dim counter As Integer
    Dim i As Integer
    Dim j As Integer
      
    If Not isNA(timeArray) Then
        length = UBound(timeArray)
        
        For i = 0 To length Step 2
            
            If i + 2 <= length Then
                If timeArray(i + 1) <> "," Then
                    condition = True
                End If
            End If
            
            If condition Then
                value1 = check(CDate(timeArray(i)), "Double")
                value2 = check(CDate(timeArray(i + 2)), "Double")
                xMin1 = value1 - minTime
                width = value2 - value1

                i = i + 2
            Else
                value2 = check(timeArray(i), "Double")
                xMin1 = value2 - minTime
                width = 0
            End If
            
            If value2 > CDbl(Now()) Then
                color = CCOLOR1
            ElseIf Format(CDate(timeArray(i)), "dd.mm.yy") = Format(CDate(Now()), "dd.mm.yy") Then
                color = CCOLORNOWTIMERECT
            Else
                color = CCOLOR2
            End If
            
            returnArray = reDimNxNArray(returnArray, counter, 2)
            returnArray(counter, 0) = xMin1
            returnArray(counter, 1) = width
            returnArray(counter, 2) = color
            counter = counter + 1
            condition = False
        Next i
        
        pNx3SizesArray = returnArray
    End If
    
End Sub


