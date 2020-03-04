VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pMinDate As Double
Private pMaxDate As Double
Private pTimeArray() As String

Property Get minDate() As Double
    minDate = pMinDate
End Property
Property Get maxDate() As Double
    maxDate = pMaxDate
End Property
Property Get timeArray() As String()    'returns date, seperator, date,..
    timeArray = pTimeArray
End Property

Sub init(nTime As String)
    Dim timeArray() As String
    Dim minMaxDate() As Double
    
    If nTime <> "" Then
        pTimeArray = convertStringToTimeArray(nTime)
        minMaxDate = getMinMaxDate(pTimeArray)
        pMinDate = minMaxDate(0)
        pMaxDate = minMaxDate(1)
    Else
        pMinDate = -1
        pMaxDate = -1
    End If
    
End Sub

Private Function convertStringToTimeArray(time As String) As String()
    Dim newString As String
    Dim length As Integer
    
    Dim returnArray() As String
    Dim dateString As String
    Dim seperatorString As String
    Dim iChar As String
    Dim datum As Date
    Dim dateNumber As Double
    
    Dim condition As Boolean
    Dim counter As Integer
    Dim i As Integer
    
    counter = 1
    condition = True
    
    newString = removeWhiteSpaces(time)
    length = Len(newString)

    i = 1
    Do While i <= length
        
        Do While condition And i <= length
            iChar = Mid(newString, i, 1)
            condition = iChar <> "-" And iChar <> "," And iChar <> ";" And iChar <> "–"         'first - and last  – are not the same !!!
            If condition Then
                dateString = dateString & iChar
            Else
                seperatorString = iChar
            End If
            
            i = i + 1
        Loop
        ReDim Preserve returnArray(0 To counter)
        datum = CDate(dateString)
        dateNumber = CDbl(datum)
        returnArray(counter - 1) = dateNumber
        returnArray(counter) = seperatorString
        counter = counter + 2
        dateString = ""
        seperatorString = ""                                            'if i = length seperator String wont be overridden
        condition = True
    Loop
    
    ReDim Preserve returnArray(0 To counter - 3)                          ' -2 which were added to counter but not to returnArray; -1 to cut away last empty string, namly seperator string

    convertStringToTimeArray = returnArray
End Function

Private Function getMinMaxDate(timeArray() As String) As Double()
    Dim length As Integer
    Dim minComparator As Double
    Dim maxComparator As Double
    Dim Value As Double
    Dim i As Integer
    Dim returnArray(0 To 1) As Double
    length = UBound(timeArray)
    minComparator = 1000000
    For i = 0 To length
        Value = check(timeArray(i), "Double")
        If Value > maxComparator Then
            maxComparator = Value
        End If
        If Value < minComparator And Value > 0 Then
            minComparator = Value
        End If
    Next i
    
    returnArray(0) = minComparator
    returnArray(1) = maxComparator
    getMinMaxDate = returnArray
End Function


