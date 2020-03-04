Attribute VB_Name = "ComputeTextBoxWidth"
Option Explicit

Const CFONT As String = "Arial"
Const cFONTSIZETEXTBOX As Integer = 9

Function layoutTextBox(wks As Worksheet, xPos As Double, yPos As Double, nHeight As Double, nText As String, numberOfLines As Integer) As shape
    Dim shapeObj As shape
    Dim stringPointWidth As Double
    Dim condition As Double: condition = True
    Dim exitCondition As Integer: exitCondition = 0
    nText = removeLinebreaks(nText)
    stringPointWidth = StrWidth(nText, CFONT, cFONTSIZETEXTBOX)
    Set shapeObj = wks.Shapes.AddTextbox(msoTextOrientationHorizontal, xPos, yPos, 10, nHeight)
    With shapeObj
        stringPointWidth = stringPointWidth + .TextFrame.MarginLeft + .TextFrame.MarginRight
        .width = stringPointWidth / numberOfLines
        .TextFrame.Characters.text = nText
        .TextEffect.FontName = CFONT
        .TextFrame.Characters.Font.size = cFONTSIZETEXTBOX
        .TextFrame2.WordWrap = msoCTrue
        Do While condition And exitCondition <= 100

            .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
            If .height > nHeight Then
                stringPointWidth = stringPointWidth + 4
                .height = nHeight
                .width = stringPointWidth / numberOfLines
                condition = True
            Else
                .height = nHeight
                condition = False
            End If
            exitCondition = exitCondition + 1
        Loop
    End With

    Set layoutTextBox = shapeObj

End Function

Function StrWidth(s As String, sFontName As String, fFontSize As Double) As Double
    ' Returns the approximate width in points of a text string
    ' in a specified font name and font size
    ' Does not account for kerning
    Dim mafChrWid(32 To 127) As Double ' widths of printing characters
    Dim msFontName As String ' font name having these widths
    Dim i As Long
    
    Dim j As Long
    If Len(sFontName) = 0 Then Exit Function
    If sFontName <> msFontName Then
        If Not InitChrWidths(sFontName, mafChrWid) Then Exit Function
    End If
    For i = 1 To Len(s)
        j = Asc(Mid(s, i, 1))
        If j >= 32 And j <= 127 Then
            StrWidth = StrWidth + fFontSize * mafChrWid(j)
        Else
            StrWidth = StrWidth + fFontSize
        End If
    Next i
End Function

Function InitChrWidths(sFontName As String, ByRef mafChrWid() As Double) As Boolean
    Dim i As Long
    
    Select Case sFontName
        Case "Arial"
            For i = 32 To 127
                Select Case i
                    Case 39, 106, 108
                        mafChrWid(i) = 0.1902
                    Case 105, 116
                        mafChrWid(i) = 0.2526
                    Case 32, 33, 44, 46, 47, 58, 59, 73, 91 To 93, 102, 124
                        mafChrWid(i) = 0.3144
                    Case 34, 40, 41, 45, 96, 114, 123, 125
                        mafChrWid(i) = 0.3768
                    Case 42, 94, 118, 120
                        mafChrWid(i) = 0.4392
                    Case 107, 115, 122
                        mafChrWid(i) = 0.501
                    Case 35, 36, 48 To 57, 63, 74, 76, 84, 90, 95, 97 To 101, 103, 104, 110 To 113, 117, 121
                        mafChrWid(i) = 0.5634
                    Case 43, 60 To 62, 70, 126
                        mafChrWid(i) = 0.6252
                    Case 38, 65, 66, 69, 72, 75, 78, 80, 82, 83, 85, 86, 88, 89, 119
                        mafChrWid(i) = 0.6876
                    Case 67, 68, 71, 79, 81
                        mafChrWid(i) = 0.7494
                    Case 77, 109, 127
                        mafChrWid(i) = 0.8118
                    Case 37
                        mafChrWid(i) = 0.936
                    Case 64, 87
                        mafChrWid(i) = 1.0602
                End Select
            Next i
        Case "Consolas"
            For i = 32 To 127
                Select Case i
                    Case 32 To 127
                        mafChrWid(i) = 0.5634
                End Select
            Next i
        Case "Calibri"
            For i = 32 To 127
                Select Case i
                    Case 32, 39, 44, 46, 73, 105, 106, 108
                        mafChrWid(i) = 0.2526
                    Case 40, 41, 45, 58, 59, 74, 91, 93, 96, 102, 123, 125
                        mafChrWid(i) = 0.3144
                    Case 33, 114, 116
                        mafChrWid(i) = 0.3768
                    Case 34, 47, 76, 92, 99, 115, 120, 122
                        mafChrWid(i) = 0.4392
                    Case 35, 42, 43, 60 To 63, 69, 70, 83, 84, 89, 90, 94, 95, 97, 101, 103, 107, 118, 121, 124, 126
                        mafChrWid(i) = 0.501
                    Case 36, 48 To 57, 66, 67, 75, 80, 82, 88, 98, 100, 104, 110 To 113, 117, 127
                        mafChrWid(i) = 0.5634
                    Case 65, 68, 86
                        mafChrWid(i) = 0.6252
                    Case 71, 72, 78, 79, 81, 85
                        mafChrWid(i) = 0.6876
                    Case 37, 38, 119
                        mafChrWid(i) = 0.7494
                    Case 109
                        mafChrWid(i) = 0.8742
                    Case 64, 77, 87
                        mafChrWid(i) = 0.936
                End Select
            Next i
        Case "Tahoma"
        For i = 32 To 127
        Select Case i
        Case 39, 105, 108
        mafChrWid(i) = 0.2526
        Case 32, 44, 46, 102, 106
        mafChrWid(i) = 0.3144
        Case 33, 45, 58, 59, 73, 114, 116
        mafChrWid(i) = 0.3768
        Case 34, 40, 41, 47, 74, 91 To 93, 124
        mafChrWid(i) = 0.4392
        Case 63, 76, 99, 107, 115, 118, 120 To 123, 125
        mafChrWid(i) = 0.501
        Case 36, 42, 48 To 57, 70, 80, 83, 95 To 98, 100, 101, 103, 104, 110 To 113, 117
        mafChrWid(i) = 0.5634
         Case 66, 67, 69, 75, 84, 86, 88, 89, 90
        mafChrWid(i) = 0.6252
        Case 38, 65, 71, 72, 78, 82, 85
        mafChrWid(i) = 0.6876
        Case 35, 43, 60 To 62, 68, 79, 81, 94, 126
        mafChrWid(i) = 0.7494
        Case 77, 119
        mafChrWid(i) = 0.8118
        Case 109
        mafChrWid(i) = 0.8742
        Case 64, 87
        mafChrWid(i) = 0.936
        Case 37, 127
        mafChrWid(i) = 1.0602
        End Select
        Next i
        Case "Lucida Console"
        For i = 32 To 127
        Select Case i
        Case 32 To 127
        mafChrWid(i) = 0.6252
        End Select
        Next i
        
        Case "Times New Roman"
        For i = 32 To 127
        Select Case i
         Case 39, 124
        mafChrWid(i) = 0.1902
        Case 32, 44, 46, 59
        mafChrWid(i) = 0.2526
        Case 33, 34, 47, 58, 73, 91 To 93, 105, 106, 108, 116
        mafChrWid(i) = 0.3144
        Case 40, 41, 45, 96, 102, 114
        mafChrWid(i) = 0.3768
        Case 63, 74, 97, 115, 118, 122
        mafChrWid(i) = 0.4392
        Case 94, 98 To 101, 103, 104, 107, 110, 112, 113, 117, 120, 121, 123, 125
        mafChrWid(i) = 0.501
        Case 35, 36, 42, 48 To 57, 70, 83, 84, 95, 111, 126
        mafChrWid(i) = 0.5634
        Case 43, 60 To 62, 69, 76, 80, 90
        mafChrWid(i) = 0.6252
        Case 65 To 67, 82, 86, 89, 119
        mafChrWid(i) = 0.6876
        Case 68, 71, 72, 75, 78, 79, 81, 85, 88
        mafChrWid(i) = 0.7494
        Case 38, 109, 127
        mafChrWid(i) = 0.8118
        Case 37
        mafChrWid(i) = 0.8742
        Case 64, 77
        mafChrWid(i) = 0.936
        Case 87
        mafChrWid(i) = 0.9984
        End Select
        Next i
    
    Case Else
        MsgBox "Font name """ & sFontName & """ not available!", vbCritical, "StrWidth"
        Exit Function
    End Select
    InitChrWidths = True
End Function


