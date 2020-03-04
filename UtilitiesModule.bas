Attribute VB_Name = "UtilitiesModule"


'FILE OPERATIONS
Function FileExists(filePath As String) As Boolean
Dim TestStr As String
    TestStr = ""
    On Error Resume Next
    TestStr = Dir(filePath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function

Function IsWorkBookOpen(fileName As String)
    Dim file As Long
    Dim ErrNo As Long

    On Error Resume Next
    file = FreeFile()
    Open fileName For Input Lock Read As #file
    Close file
    ErrNo = err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function

Function WorksheetExists(wksName As String) As Boolean

For Each wks In Worksheets
    If wks.name = wksName Then
        WorksheetExists = True
        Exit Function
    End If
Next wks

WorksheetExists = False
End Function

Function openPPFile(pp As PowerPoint.Application, filePath As String) As PowerPoint.Presentation
On Error GoTo err
    
    Set openPPFile = pp.Presentations.Open(filePath)
    Exit Function
err:
MsgBox "Please select valide File"
End
End Function
'------------------------------------------------------------------------------------------------------------------------------------

Function check(dataType As Variant, desiredDataType As String) As Variant

    Select Case desiredDataType
        Case "Integer"
            On Error Resume Next
            check = CInt(dataType)
            On Error GoTo 0
            If TypeName(check) <> desiredDataType Then
                check = 0
            End If
        Case "Single"
            On Error Resume Next
            check = CSng(dataType)
            On Error GoTo 0
            If TypeName(check) <> desiredDataType Then
                check = 0
            End If
        Case "Double"
            On Error Resume Next
            check = CDbl(dataType)
            On Error GoTo 0
            If TypeName(check) <> desiredDataType Then
                check = 0
            End If
        Case "String"
            On Error Resume Next
            check = CStr(dataType)
            On Error GoTo 0
            If TypeName(check) <> desiredDataType Then
                check = ""
            End If
    End Select

End Function
'-----------------------------------------------------------------------------------------------------------------------

''------------------------------------------------------------------------------------------------------------------------------------
'String OPERATIONS
Function removeWhiteSpaces(oString As String) As String
length = Len(oString)

For i = 1 To length
    iChar = Mid(oString, i, 1)
    If iChar <> " " Then
        nString = nString & iChar
    End If
Next i

removeWhiteSpaces = nString
End Function

Function removeLinebreaks(myString As String) As String
    Dim arrayString() As String
    Dim newString As String
    'newString = ""
    For i = 1 To Len(myString) Step 2
        
        iChar = Mid(myString, i, 2)
        If iChar <> vbCrLf And iChar <> vbNewLine And iChar <> Chr(10) And iChar <> Chr(13) Then
            newString = newString & iChar
        End If
    Next i
    
    removeLinebreaks = newString
End Function
'------------------------------------------------------------------------------------------------------------------------------------
'ARRAY OPERATIONS

Function arraysStartWith0(arr As Variant) As Variant
    Dim startLength As Integer
    Dim startWidth As Integer
    Dim length As Integer
    Dim width As Integer
    Dim i As Integer
    Dim j As Integer
    Dim returnArray As Variant
    
    length = UBound(arr, 1)
    width = UBound(arr, 2)
    startLength = LBound(arr, 1)
    startWidth = LBound(arr, 2)
    
    ReDim returnArray(0 To length - startLength, 0 To width - startWidth)
    
    For i = startLength To length
        For j = startWidth To width
            returnArray(i - startLength, j - startWidth) = arr(i, j)
        Next j
    Next i
    
    arraysStartWith0 = returnArray
End Function

Function isNA(arr As Variant) As Boolean
On Error GoTo NA

    x = UBound(arr)
    If x = -1 Then
        isNA = True
    Else
        isNA = False
    End If
    Exit Function

NA:
On Error GoTo 0
isNA = True
End Function

Function removeNARowsFromMatrix(arr As Variant) As Variant
    Dim length As Integer
    Dim width As Integer
    Dim i As Integer
    Dim j As Integer
    Dim naRowCounter As Integer
    Dim rowIsNA As Boolean: rowIsNA = True
    Dim returnArray As Variant

    length = UBound(arr, 1)
    width = UBound(arr, 2)
    
    For i = 0 To length
        For j = 0 To width
            If Not IsEmpty(arr(i, j)) Then
                rowIsNA = False
            End If
        Next j
        If rowIsNA Then
            naRowCounter = naRowCounter + 1
        End If
        rowIsNA = True
    Next i
    
    If length < naRowCounter Then
        returnArray = Empty
    Else
        ReDim returnArray(0 To length - naRowCounter, 0 To width)
        naRowCounter = 0
        For i = 0 To length
            For j = 0 To width
                If Not IsEmpty(arr(i, j)) Then
                    rowIsNA = False
                End If
            Next j
            If Not rowIsNA Then
                filArray returnArray, naRowCounter, arr, i
                naRowCounter = naRowCounter + 1
            End If
            rowIsNA = True
        Next i
    End If
    removeNARowsFromMatrix = returnArray
End Function

Function colIsNA(arr As Variant, col As Integer) As Boolean
    Dim i As Integer
    
    For i = 0 To UBound(arr)
        If Not IsEmpty(arr(i, col)) Then
            isColNA = False
            Exit Function
        End If
    Next i
    
    colIsNA = True
End Function

Function reDimNxNArray(arr As Variant, sizeL As Integer, sizeW As Integer) As Variant
    Dim returnArray As Variant
    
    Dim width As Integer
    Dim length As Integer
    Dim width1 As Integer
    Dim length1 As Integer
    
    Dim i As Integer
    Dim j As Integer
    
    If IsEmpty(arr) Then
        ReDim returnArray(0 To sizeL, 0 To sizeW)
    Else
        length1 = UBound(arr, 1)
        width1 = UBound(arr, 2)
        If sizW > width1 Then
            width = width1
        Else
            width = sizeW
        End If
        If sizeL > length1 Then
            length = length1
        Else
            length = sizeL
        End If
        ReDim returnArray(0 To sizeL, 0 To sizeW)
        For i = 0 To length
            For j = 0 To width
                returnArray(i, j) = arr(i, j)
            Next j
        Next i
    End If
    
    reDimNxNArray = returnArray
End Function

Function emptyArray(arr As Variant) As Variant
    Dim i As Integer
    
    For i = 0 To UBound(arr)
        arr(i) = Emtpy
    Next i
    
    emptyArray = arr
End Function

'returns one array of the appropiate values for each different value in the orderColumn
Function splitArray(arr As Variant, byOrderCol As Integer) As Variant

Dim pivotArray() As Variant
Dim returnArray As Variant


length = UBound(arr, 1)
lLength = LBound(arr, 1)
counter = lLength
For i = lLength To length
    index = inArray(pivotArray, arr(i, byOrderCol))
    If index = -1 Then
        ReDim Preserve pivotArray(lLength To counter)
        pivotArray(counter) = arr(i, byOrderCol)               'infers orderValue
        counter = counter + 1
    End If
Next i

newLength = UBound(pivotArray)
ReDim returnArray(lLength To newLength)
For i = lLength To newLength
    returnArray(i) = filterArray(arr, "argPivot", Array(byOrderCol, pivotArray(i)))
Next i
splitArray = returnArray
End Function

Function uniqueInArray(arr As Variant) As Variant
Dim returnArray() As Variant
Dim counter As Integer

length = UBound(arr)

For i = 0 To length
    If inArray(returnArray, arr(i)) = -1 Then
        ReDim Preserve returnArray(0 To counter)
        returnArray(counter) = arr(i)
        counter = counter + 1
    End If
Next i

uniqueInArray = returnArray
End Function

'returns the index of the first element equal the arg, if arg not in arr then return -1
Function inArray(arr As Variant, arg As Variant) As Integer
On Error GoTo err

For Each i In arr
    If i = arg Then
        inArray = counter
        Exit Function
    End If
    counter = counter + 1
Next i
err:
On Error GoTo 0
inArray = -1
End Function

Sub filArray(ByRef arr1 As Variant, index1 As Integer, arr2 As Variant, index2 As Integer)
    For j = LBound(arr1, 2) To UBound(arr1, 2)
        arr1(index1, j) = arr2(index2, j)
    Next j
End Sub

Sub addHeaderToArray(ByRef arr As Variant, headerNames As Variant)
    Dim returnArray As Variant
    Dim length As Integer
    Dim width As Integer
    length = UBound(arr, 1)
    width = UBound(arr, 2)
    
    ReDim returnArray(0 To length + 1, 0 To width)
    
    For i = 0 To length + 1
        For j = 0 To width
            If i = 0 Then
                returnArray(0, j) = headerNames(j)
            Else
                returnArray(i, j) = arr(i - 1, j)
            End If
        Next j
    Next i
    
    arr = returnArray
End Sub

'Input: Array of Matrices you want to connect -> Array(mat1, mat2); Output: connected Matrix
Function connectMatrix(arrs As Variant) As Variant
Dim returnArray As Variant
Dim width As Integer
Dim widCounter As Integer
length = UBound(arrs(0), 1)

For Each arr In arrs
    width = width + UBound(arr, 2) + 1
Next arr

ReDim returnArray(0 To length, 0 To width - 1)

For Each arr In arrs
    For i = 0 To length
        width = UBound(arr, 2)
        For j = 0 To width
            returnArray(i, widCounter + j) = arr(i, j)
        Next j
    Next i
    widCounter = widCounter + width + 1
Next arr

connectMatrix = returnArray
End Function

Function connectArrays(arrs As Variant) As Variant
Dim returnArray As Variant
Dim width As Integer
width = leng(arrs)
length = leng(arrs(0))
ReDim returnArray(0 To length, 0 To width)

For j = 0 To width
    For i = 0 To length
        returnArray(i, j) = arrs(j)(i)
    Next i

Next j
connectArrays = returnArray
End Function

'Matrix has to be first variable in arrs (logic is the same as in ceonnextMatrix()) + you can determine how many columns need to be left out of the arrays with leaveStartColOutArray
Function connectMatrixWithArray(arrs As Variant, Optional leaveStartColOutArray As Variant = 0) As Variant
Dim returnArray As Variant
Dim width As Integer
Dim widCounter As Integer
Dim leaveStartColOut As Integer
length = UBound(arrs(0), 1)
width = UBound(arrs(0), 2)
For i = LBound(arrs) + 1 To UBound(arrs)
    width = width + UBound(arrs(i), 1) + 1
Next i

ReDim returnArray(0 To length, 0 To width - 1)

For i = LBound(arrs) To UBound(arrs)
    If i = LBound(arrs) Then
        For j = 0 To length
            For k = 0 To UBound(arrs(i), 2)
                returnArray(j, widCounter + k) = arrs(i)(j, k)
            Next k
        Next j
        widCounter = widCounter + k
    Else
        If Not IsEmpty(leaveStartColOutArray) Then leaveStartColOut = leaveStartColOutArray(i - 1)
        For j = 0 To UBound(arrs(i)) - leaveStartColOut
            returnArray(0, widCounter + j) = arrs(i)(j + leaveStartColOut)
        Next j
        widCounter = widCounter + j
    End If
Next i

connectMatrixWithArray = returnArray
End Function

'Conditions: create Function returning bool value if criteria is met and use name of this function as argument:=filterfunction
'e.g:
Function argPivot(argArr As Variant, filterCriteria As Variant) As Boolean
If argArr(filterCriteria(0)) = filterCriteria(1) Then
    argPivot = True
Else
    argPivot = False
End If
End Function
Function filterArray(arr As Variant, filterFunction As String, filterCriteria1 As Variant) As Variant
Dim helpArray As Variant
Dim criteriaArray As Variant

Dim lowlength As Integer
Dim lowwidth As Integer
Dim uplength As Integer
Dim upwidth As Integer
Dim dimCounter As Integer

lowlength = LBound(arr, 1)
lowwidth = LBound(arr, 1)
uplength = UBound(arr, 1)
upwidth = UBound(arr, 2)

ReDim criteriaArray(lowlength To upwidth)
dimCounter = lowlength
For i = lowlength To uplength
    For j = lowwidth To upwidth
        criteriaArray(j) = arr(i, j)
    Next j
    If Application.Run(filterFunction, criteriaArray, filterCriteria1) Then
        dimCounter = dimCounter + 1
    End If
Next i

If dimCounter > lowlength Then
    dimCounter = dimCounter - 1 'one gets added too much after last satisfied criteria
Else    'no value matched criteria
    Exit Function
End If

ReDim helpArray(lowlength To dimCounter, lowwidth To upwidth)
dimCounter = lowlength
For i = lowlength To uplength
    For j = lowwidth To upwidth
        criteriaArray(j) = arr(i, j)
    Next j
    If Application.Run(filterFunction, criteriaArray, filterCriteria1) Then
        For j = lowwidth To upwidth
            helpArray(dimCounter, j) = arr(i, j)
        Next j
        dimCounter = dimCounter + 1
    End If
Next i

filterArray = helpArray
End Function

'can sort 1 to 2 dimensional arrays
Sub sortArray(ByRef oldArray As Variant, Optional sortColumn As Integer = 0)

Dim length As Integer
Dim width As Integer
Dim helpArray As Variant

Dim i As Integer

length = UBound(oldArray, 1)

On Error Resume Next
width = UBound(oldArray, 2)

If width = 0 Then
    On Error GoTo 0
    ReDim helpArray(0 To length, 0)
    For i = LBound(oldArray) To length
        helpArray(i, 0) = oldArray(i)
    Next i
    'NoramlSort helpArray, sortColumn
    QuickSort helpArray, 0, length, sortColumn
    For i = LBound(oldArray) To length
        oldArray(i) = helpArray(i, 0)
    Next i
    Exit Sub
Else
    On Error GoTo 0
    'NoramlSort oldArray, sortColumn
    QuickSort oldArray, 0, length, sortColumn
End If

End Sub

Sub swapIndex(ByRef oldArray As Variant, left As Integer, right As Integer)
    Dim width As Integer
    Dim helpArray As Variant
    
    Dim j As Integer
    
    width = UBound(oldArray, 2)
    
    ReDim helpArray(0 To width)
    
    For j = 0 To width
        helpArray(j) = oldArray(left, j)
        oldArray(left, j) = oldArray(right, j)
        oldArray(right, j) = helpArray(j)
    Next j
End Sub

Sub QuickSort(ByRef arr, left As Integer, right As Integer, sortColumn As Integer)
  Dim varPivot As Variant
  Dim leftPart As Integer
  Dim rightPart As Integer
  Dim formula As Integer
  
  leftPart = left
  rightPart = right
  formula = (left + right) / 2
  varPivot = arr(formula, sortColumn)
  
  Do While leftPart <= rightPart
    Do While arr(leftPart, sortColumn) < varPivot And leftPart < right
      leftPart = leftPart + 1
    Loop
    Do While varPivot < arr(rightPart, sortColumn) And rightPart > left
      rightPart = rightPart - 1
    Loop
    If leftPart <= rightPart Then
      swapIndex arr, leftPart, rightPart
      leftPart = leftPart + 1
      rightPart = rightPart - 1
    End If
  Loop
  If left < rightPart Then QuickSort arr, left, rightPart, sortColumn
  If leftPart < right Then QuickSort arr, leftPart, right, sortColumn
End Sub

Sub NoramlSort(ByRef arr, sortColumn As Integer)

Dim length As Integer

Dim pivot As Variant
Dim oldPivot As Variant

Dim right As Integer

length = UBound(arr, 1)

For i = 0 To length
    pivot = arr(i, sortColumn)
    For j = i + 1 To length
        If arr(j, sortColumn) >= oldPivot And arr(j, sortColumn) < pivot Then
            pivot = arr(j, sortColumn)
            right = j
        End If
    Next j
    
    oldPivot = pivot
    If arr(i, sortColumn) > pivot Then swapIndex arr, i, right
Next i
End Sub

''------------------------------------------------------------------------------------------------------------------------------------
'RANDOM OPERATIONS

Function returnRandomNotNullCol(arr As Variant, row As Integer, width As Integer, Optional startCol As Integer = 0) As Integer
rndNumber = (-1) ^ Int(2 * Rnd + 1)
rndCol = Int((width - startCol + 1) * Rnd + startCol)
j = rndCol
Do While j >= startCol And j <= width
    If check(arr(row, j), "Double", "", 0, "") > 0 Then
        returnRandomNotNullCol = j
        Exit Function
    End If
    j = j + rndNumber
Loop
j = rndCol - rndNumber
Do While j >= startCol And j <= width
    If check(arr(row, j), "Double", "", 0, "") > 0 Then
        returnRandomNotNullCol = j
        Exit Function
    End If
    j = j - rndNumber
Loop
returnRandomNotNullCol = -1
End Function








