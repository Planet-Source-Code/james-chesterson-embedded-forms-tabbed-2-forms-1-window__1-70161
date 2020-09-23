VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Main form"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.TabStrip tbsTabs 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5953
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sub form 1"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sub form 2"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sub form 3"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    LoadForm frmSub1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm
End Sub

Private Sub tbsTabs_Click()
    'Which tab was clicked?
    Select Case tbsTabs.SelectedItem.Index
    Case 1
        SwitchForm frmSub1
    Case 2
        SwitchForm frmSub2
    Case 3
        SwitchForm frmSub3
    End Select
    'Set focus on the main form
    'frmMain.SetFocus
End Sub
