VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "By Alexander Czernik"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8640
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    StartButton.Enabled = False
    PPSwimlaneCheckbox.Enabled = False
    BrowseLocationCheckButton.Enabled = False
    RemoveBarsButton.Enabled = False
    If Not FileExists(Application.ActiveWorkbook.Path & "\Swimlane_Template.pptx") Then
        TemplateCheckButton.Enabled = False
    End If
End Sub

Private Sub StartButton_Click()
    Call StartAutomatisation
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub TimeTableCheckBox_Click()
    If TimeTableCheckBox.Value Or ExcelSwimlaneCheckbox.Value Or PPSwimlaneCheckbox.Value Then
        StartButton.Enabled = True
    Else
        StartButton.Enabled = False
    End If
End Sub

Private Sub ExcelSwimlaneCheckbox_Click()
    If TimeTableCheckBox.Value Or ExcelSwimlaneCheckbox.Value Or PPSwimlaneCheckbox.Value Then
        StartButton.Enabled = True
    Else
        StartButton.Enabled = False
    End If
End Sub

Private Sub PPSwimlaneCheckbox_Click()
    If Not PPSwimlaneCheckbox Then
        NewPresentationCheckButton.Value = False
        TemplateCheckButton.Value = False
        BrowseLocationCheckButton.Value = False
        EnterPathTextfield.text = ""
        
        PPSwimlaneCheckbox.Enabled = False
    Else
        RemoveBarsButton.Enabled = True
    End If
    If TimeTableCheckBox.Value Or ExcelSwimlaneCheckbox.Value Or PPSwimlaneCheckbox.Value Then
        StartButton.Enabled = True
    Else
        StartButton.Enabled = False
    End If
End Sub

Private Sub newPresentationCheckButton_Click()
    If NewPresentationCheckButton.Value Then
        PPSwimlaneCheckbox.Value = True
        PPSwimlaneCheckbox.Enabled = True
    End If
End Sub

Private Sub TemplateCheckButton_Click()
    If TemplateCheckButton.Value Then
        PPSwimlaneCheckbox.Value = True
        PPSwimlaneCheckbox.Enabled = True
    End If
End Sub

Private Sub BrowseFolderButton_Click()
On Error GoTo err
    Dim fileExplorer As FileDialog
    Set fileExplorer = Application.FileDialog(msoFileDialogFilePicker)

    'To allow or disable to multi select
    fileExplorer.AllowMultiSelect = False

    With fileExplorer
        If .Show = -1 Then 'Any file is selected
            EnterPathTextfield.text = .SelectedItems.Item(1)
        Else ' else dialog is cancelled
            MsgBox "You have cancelled the dialogue"
            EnterPathTextfield.text = "" ' when cancelled set blank as file path.
        End If
    End With
err:
    Exit Sub
End Sub

Private Sub BrowseLocationCheckButton_Click()
    PPSwimlaneCheckbox.Value = True
    PPSwimlaneCheckbox.Enabled = True
End Sub

Private Sub EnterPathTextfield_Change()
    BrowseLocationCheckButton.Value = EnterPathTextfield.text <> ""
    PPSwimlaneCheckbox.Value = EnterPathTextfield.text <> ""
    PPSwimlaneCheckbox.Enabled = EnterPathTextfield.text <> ""
End Sub

Private Sub removeBarsButton_Click()

End Sub

Private Sub StartAutomatisation()
'On Error GoTo whoopsie
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
    If ExcelSwimlaneCheckbox.Value Then
        Call LayoutSwimlane.LayoutSwimlane
    End If
    
    If TimeTableCheckBox.Value Then
        Call LayoutInTimeAble.LayoutInTimeAble
    End If
    
    If NewPresentationCheckButton.Value Then
        Call LayoutInPP.layoutSwimlaneInPP(0, RemoveBarsButton.Value)
    ElseIf TemplateCheckButton.Value Then
        Call LayoutInPP.layoutSwimlaneInPP(1, RemoveBarsButton.Value, Application.ActiveWorkbook.Path & "\Swimlane_Template.pptx")
    ElseIf BrowseLocationCheckButton.Value Then
        Call LayoutInPP.layoutSwimlaneInPP(2, RemoveBarsButton.Value, EnterPathTextfield.text)
    End If
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationManual
    End With
    Unload Me
    Exit Sub
    
'whoopsie:
With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationManual
End With
MsgBox "An Error has ocurred"
Unload Me
End
End Sub


