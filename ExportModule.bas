Attribute VB_Name = "ExportModule"
Public Sub prcExort()
    Dim objVBComponent As Object
    Dim objWorkbook As Workbook
    Dim strType As String
    Set objWorkbook = ActiveWorkbook
    
    For Each objVBComponent In objWorkbook.VBProject.VBComponents
        With objVBComponent.CodeModule
            Select Case objVBComponent.Type
                Case 1
                    strType = ".bas"
                Case 2, 100
                    strType = ".cls"
                Case 3
                    strType = ".frm"
            End Select
            objWorkbook.VBProject.VBComponents(objVBComponent.name).Export _
                "C:\Users\alexc\Desktop\Bye Generali\Macros\Visualize\" & objVBComponent.name & strType
        End With
    Next
End Sub

