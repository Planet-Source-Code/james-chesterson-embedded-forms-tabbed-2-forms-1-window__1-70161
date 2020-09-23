Attribute VB_Name = "modMultiForm"
Public Declare Function SetParent Lib "user32" _
(ByVal hWndChild As Long, _
ByVal hWndNewParent As Long) As Long

Public CurrentForm As Object

Public Sub LoadForm(oForm As Object)
    Dim Form As Form
    
    Load oForm
    Set CurrentForm = oForm
    Set Form = oForm
    
    Form.Top = frmMain.Label1.Top
    Form.Left = 0
    
    SetParent Form.hWnd, frmMain.hWnd
    Form.Show
    
End Sub

Public Sub SwitchForm(NewForm As Object)
    Dim Form As Form 'Form object
    
    Unload CurrentForm 'Unload the old form that is embedded
    Load NewForm 'Load the new form into memory
    
    Set Form = NewForm 'Access the object NewForm from Form (To access form properties)
    Form.Top = frmMain.Label1.Top 'Align the top of the form
    Form.Left = 0 'Align the left of the form
    
    SetParent Form.hWnd, frmMain.hWnd 'Set the new form as a child of frmMain
    Form.Show 'Show the form
    
    Set CurrentForm = NewForm 'Store the current form in a global so it can be unloaded when the program exits
End Sub

Public Sub UnloadForm()
    Unload CurrentForm
End Sub
