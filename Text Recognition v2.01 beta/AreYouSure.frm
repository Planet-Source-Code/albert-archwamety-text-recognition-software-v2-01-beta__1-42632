VERSION 5.00
Begin VB.Form frmConfirmation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirmation"
   ClientHeight    =   990
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton NOButton 
      Caption         =   "&No"
      Height          =   375
      Left            =   1875
      TabIndex        =   1
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton YESButton 
      Caption         =   "&Yes"
      Height          =   375
      Left            =   315
      TabIndex        =   0
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Are you sure want to Exit ?"
      Height          =   375
      Left            =   195
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public YES As Boolean

'Private Sub Form_Load()
'    Me.Left = (frmMain.Width / 2) - (Me.Width / 2)
'    Me.Top = (frmMain.ScaleHeight / 2) - (Me.Height)
'End Sub

Private Sub NOButton_Click()
    Unload Me
End Sub

Private Sub YESButton_Click()
    YES = True
    Unload Me
End Sub
