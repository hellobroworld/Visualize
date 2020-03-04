VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InTimeAbleObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Notations:
    'arrays always start with 0
    ':arrays: nXnArray (n: length, X, n: width) if n = n -> arbitrary size

Private pMinTime As Double
Private pMaxTime As Double

Private pCalendarItems As Collection
Private pTimeRects As Collection

Property Get minTime() As Double
    minTime = pMinTime
End Property

Property Get maxTime() As Double
    maxTime = pMaxTime
End Property

Property Get currentTime() As Double
    currentTime = CDbl(Now())
End Property

Property Get currentTimeRect() As TimeRectObj
    Dim timeArray(0 To 2) As String

    Set currentTimeRect = New TimeRectObj
    timeArray(0) = "01." & month(CDate(pMinTime)) & "." & year(CDate(pMinTime))
    timeArray(1) = "-"
    timeArray(2) = CStr(Now())
    currentTimeRect.init pMinTime, timeArray
End Property

Property Get calendarItems() As Collection
    Set calendarItems = pCalendarItems
End Property

Property Get timeRects() As Collection
    Set timeRects = pTimeRects
End Property

Sub init(nx4InputArray As Variant)
    Dim minMaxTime() As Double
    
    Set pCalendarItems = gatherCalendarItems(nx4InputArray)
    minMaxTime = getMinMaxTime(pCalendarItems)
    pMinTime = minMaxTime(0)
    pMaxTime = minMaxTime(1)
    
    Set pTimeRects = gatherTimeRects(pCalendarItems, pMinTime)
    
End Sub


Private Function gatherCalendarItems(nx4InputArray As Variant) As Collection

    Dim returnCalendarItems As Collection
    Dim calendItem As CalendarItem
    Dim length As Integer
    Dim id As Integer
    Dim dateString As String
    Dim title1 As String
    Dim title2 As String
    
    Dim i As Integer
    
    length = UBound(nx4InputArray, 1)
    Set returnCalendarItems = New Collection
    
    For i = 0 To length
        Set calendItem = New CalendarItem
        id = nx4InputArray(i, 0)
        dateString = nx4InputArray(i, 1)
        title1 = nx4InputArray(i, 2)
        title2 = nx4InputArray(i, 3)
        calendItem.init i + 1, dateString, title1, title2
        returnCalendarItems.Add calendItem, CStr(i + 1)
    Next i
    
    Set gatherCalendarItems = returnCalendarItems
End Function

Private Function getMinMaxTime(calendItems As Collection) As Double()
    Dim length As Integer
    Dim calendItem As CalendarItem
    Dim minComparator As Double
    Dim maxComparator As Double
    Dim minValue As Double
    Dim maxValue As Double
    
    Dim returnArray(0 To 1) As Double
    Dim i As Integer
    
    length = calendItems.Count
    minComparator = 1000000
    For i = 1 To length
        Set calendItem = calendItems.Item(i)
        If calendItem.dateItem.minDate <> -1 Then
            minValue = calendItem.dateItem.minDate
            maxValue = calendItem.dateItem.maxDate
            If minValue < minComparator Then minComparator = minValue
            
            If maxValue > maxComparator Then maxComparator = maxValue
        End If
    Next i
    
    returnArray(0) = minComparator
    returnArray(1) = maxComparator
    getMinMaxTime = returnArray
End Function

Private Function gatherTimeRects(calendItems As Collection, minTime As Double) As Collection
    Dim returnTimeRects As Collection
    Dim timeRec As TimeRectObj
    Dim timeArray() As String
    Dim length As Integer
    Dim i As Integer
    
    length = calendItems.Count
    Set returnTimeRects = New Collection
    For i = 1 To length
        Set timeRec = New TimeRectObj
        timeArray = calendItems.Item(i).dateItem.timeArray
        timeRec.init minTime, timeArray
        returnTimeRects.Add timeRec, CStr(calendItems.Item(i).id)
    Next i
    
    Set gatherTimeRects = returnTimeRects
End Function






