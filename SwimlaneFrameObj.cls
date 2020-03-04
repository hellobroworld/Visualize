VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SwimlaneFrameObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const cBARHEIGHT As Double = 1
Const cWHOBARWIDTH As Double = 1
Const cSWIMBARWIDTH As Double = 1

Private pWho() As String
Private pWhoBarsCount As Integer

Property Get whoArray() As String()
    whoArray = pWho
End Property

Property Get whoBarsCount() As Integer
    whoBarsCount = pWhoBarsCount
End Property

Sub init(nx1Array() As String)
    pWho = sortWho(nx1Array)
    pWhoBarsCount = UBound(pWho) + 1
End Sub

Private Function sortWho(nx1Array() As String) As String()
    Dim uniqueArray As Variant
    Dim sortedArray As Variant
    Dim length As Integer
    Dim returnArray() As String
    
    Dim i As Integer
    
    uniqueArray = uniqueInArray(nx1Array)
    length = UBound(uniqueArray)
    
    ReDim sortedArray(0 To length, 0 To 1)
    
    For i = 0 To length
        sortedArray(i, 0) = uniqueArray(i)
        sortedArray(i, 1) = conditionWho(CStr(uniqueArray(i)))
    Next i
    
    sortArray sortedArray, 1
    
    ReDim returnArray(0 To length)
    For i = 0 To length
        returnArray(i) = CStr(sortedArray(i, 0))
    Next i
    
    sortWho = returnArray
End Function

Private Function conditionWho(arg As String) As Integer
    Dim regexBR As New RegExp
    Dim regexAR As New RegExp
    Dim regexMA As New RegExp
    Dim regexCxo As New RegExp
    Dim regexL2 As New RegExp
    Dim regexL3 As New RegExp
    Dim regexFK As New RegExp
    Dim regexHRBP As New RegExp
    Dim regexHR As New RegExp
    Dim cond As Boolean
    

    arg = removeWhiteSpaces(arg)
    regexBR.Pattern = "BR"
    regexAR.Pattern = "AR"
    regexMA.Pattern = "MA"
    regexCxo.Pattern = "C.+?O"
    regexL2.Pattern = "L2"
    regexL3.Pattern = "L3"
    regexFK.Pattern = "FK"
    regexHRBP.Pattern = "HR*BP"
    regexHR.Pattern = "HR"
    If regexBR.Test(arg) Then
        conditionWho = 1
    ElseIf regexAR.Test(arg) Then
        conditionWho = 2
    ElseIf regexMA.Test(arg) Then
        conditionWho = 3
    ElseIf regexCxo.Test(arg) Then
        conditionWho = 4
    ElseIf regexL2.Test(arg) Then
        conditionWho = 5
    ElseIf regexL3.Test(arg) Then
        conditionWho = 6
    ElseIf regexFK.Test(arg) Then
        conditionWho = 7
    ElseIf regexHRBP.Test(arg) Then
        conditionWho = 8
    ElseIf regexHR.Test(arg) Then
        conditionWho = 9
    Else
        conditionWho = 10
    End If
End Function
