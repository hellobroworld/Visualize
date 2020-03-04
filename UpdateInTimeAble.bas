Attribute VB_Name = "UpdateInTimeAble"
Option Explicit
Const CWKSNAMEINTIMEABLE As String = "Time Table"
Const CCOLORFRAMETITLE As Long = 12566463    'color : ?RGB(rrr,ggg,bbb) -> Long
Const CCOLORFRAME1 As Long = 15921906
Const CCOLORFRAME2 As Long = 14277081

Const CCOLOR1 As Long = 8355711
Const CCOLOR2 As Long = 1907906
Const CCOLORNOWTIMERECT As Long = 10213059

Const CFONT As String = "Arial"
Const CFONTSIZE As Integer = 12
Const cSTARTROW As Integer = 3
Const cSTARTCOL As Integer = 1

Const CTOPFRAMEBORDERWIDTH As Integer = 2
Const CLEADINGFRAMEBORDERWIDTH As Integer = 3
Const CTRAILINGFRAMEBORDERWIDTH As Integer = 1

Const CCOLPERMONTH As Integer = 1
Const CROWPERSTEP As Integer = 2
Const CCELLWIDTH As Integer = 15
Const CCELLHEIGHT As Integer = 15


Sub UpdateCurrentTime()
On Error GoTo noTimeTable
    Dim inTimeAble As InTimeAbleObj: Set inTimeAble = InitInTimeAble()
    Dim wks As Worksheet: Set wks = Worksheets(CWKSNAMEINTIMEABLE)
On Error GoTo 0

    Dim shapeObj As shape
    Dim shapeLine As shape
    Dim shapeI As Variant
    
    Dim xEndPos As Double
    Dim color As Long
    
    For Each shapeI In wks.Shapes
        If shapeI.name = "0" Then
            Set shapeObj = shapeI
        ElseIf shapeI.name = "1000" Then
            shapeI.Delete
        End If
    Next shapeI
    
    If Not shapeObj Is Nothing Then
        shapeObj.Delete
    End If
    Set shapeObj = layoutCurrentTimeRect(wks, inTimeAble)
    If Not shapeObj Is Nothing Then
        xEndPos = shapeObj.left + shapeObj.width
        
        For Each shapeI In wks.Shapes
            If shapeI.name <> "0" And shapeI.name <> "1000" And shapeI.name <> "999" And shapeI.name <> "" Then
            
                If xEndPos < shapeI.left + shapeI.width Then
                    color = CCOLOR1
                Else
                    color = CCOLOR2
                End If
                With shapeI
                    .Fill.ForeColor.RGB = color
                    .Line.ForeColor.RGB = color
                End With
            End If
        Next shapeI
    End If
Exit Sub

noTimeTable:
End Sub
