VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalendarItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pId As Integer
Private pDateObj As DateObj
Private pTitle1 As String
Private pTitle2 As String

Property Get id() As Integer
    id = pId
End Property
Property Get dateItem() As DateObj
    Set dateItem = pDateObj
End Property
Property Get title1() As String
    title1 = pTitle1
End Property
Property Get title2() As String
    title2 = pTitle2
End Property

Sub init(nId As Integer, nDate As String, nTitle1 As String, nTitle2 As String)
    pId = nId
    Set pDateObj = New DateObj
    pDateObj.init nDate
    pTitle1 = nTitle1
    pTitle2 = nTitle2
End Sub
