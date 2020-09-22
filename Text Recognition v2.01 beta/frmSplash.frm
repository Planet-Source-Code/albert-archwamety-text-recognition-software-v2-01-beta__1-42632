VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   4710
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   7620
      Begin VB.PictureBox picLogo 
         Height          =   2385
         Left            =   510
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   2325
         ScaleWidth      =   1755
         TabIndex        =   2
         Top             =   855
         Width           =   1815
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "LicenseTo Albert Archwamety"
         Height          =   255
         Left            =   270
         TabIndex        =   1
         Tag             =   "LicenseTo"
         Top             =   300
         Width           =   6855
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Text Recognition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   29.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2670
         TabIndex        =   9
         Tag             =   "Product"
         Top             =   1200
         Width           =   4800
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "CompanyProduct"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2505
         TabIndex        =   8
         Tag             =   "CompanyProduct"
         Top             =   765
         Width           =   3000
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Platform Win 9x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4815
         TabIndex        =   7
         Tag             =   "Platform"
         Top             =   2400
         Width           =   2190
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version 0.1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5625
         TabIndex        =   6
         Tag             =   "Version"
         Top             =   2760
         Width           =   1380
      End
      Begin VB.Label lblWarning 
         Caption         =   "Warning"
         Height          =   195
         Left            =   300
         TabIndex        =   3
         Tag             =   "Warning"
         Top             =   3720
         Width           =   6855
      End
      Begin VB.Label lblCompany 
         Caption         =   "Company of AA"
         Height          =   255
         Left            =   4710
         TabIndex        =   5
         Tag             =   "Company"
         Top             =   3330
         Width           =   2415
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright Albert Archwamety "
         Height          =   255
         Left            =   4710
         TabIndex        =   4
         Tag             =   "Copyright"
         Top             =   3120
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

