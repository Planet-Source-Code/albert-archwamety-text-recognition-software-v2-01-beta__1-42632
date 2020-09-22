VERSION 5.00
Begin VB.Form frmTeach 
   Caption         =   "Teach"
   ClientHeight    =   1095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Teach 
      Caption         =   "&Teach"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox TeachText 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   160
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a character to be teach"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmTeach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub Teach_Click()
    frmMain.Main_TeachText = TeachText.Text
    Unload Me
End Sub
