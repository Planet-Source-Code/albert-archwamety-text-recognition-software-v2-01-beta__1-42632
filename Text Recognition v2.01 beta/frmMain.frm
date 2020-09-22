VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Text Recognition Software (v 2.01 beta)"
   ClientHeight    =   6570
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8730
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMain.frx":0442
   ScaleHeight     =   6570
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox comboRecognise 
      Height          =   315
      Left            =   3120
      TabIndex        =   25
      Text            =   "Few other matches ..."
      Top             =   3840
      Width           =   2535
   End
   Begin MSComctlLib.ProgressBar pbRecognising 
      Height          =   255
      Left            =   3120
      TabIndex        =   24
      Top             =   1320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.ComboBox comboOpen 
      Height          =   315
      ItemData        =   "frmMain.frx":074C
      Left            =   6000
      List            =   "frmMain.frx":074E
      TabIndex        =   23
      Text            =   "Select a character to open ..."
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Frame Status 
      Caption         =   "Status - Tips && Help Window"
      Height          =   1095
      Left            =   120
      TabIndex        =   19
      Top             =   4920
      Width           =   5775
      Begin VB.Label StatusLabel 
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.CommandButton TeachCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   4530
      Width           =   1455
   End
   Begin VB.CommandButton TeachConfirm 
      Caption         =   "C&onfirm"
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   4170
      Width           =   1455
   End
   Begin VB.TextBox TeachText 
      Height          =   315
      Left            =   2640
      TabIndex        =   11
      Top             =   4170
      Width           =   495
   End
   Begin MSComDlg.CommonDialog Teach_CommonDialog 
      Left            =   5520
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog Open_CommonDialog 
      Left            =   5520
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Open 
      Caption         =   "&Open"
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   4200
      Width           =   2535
   End
   Begin VB.CommandButton Teach 
      Caption         =   "&Teach"
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   4680
      Width           =   2535
   End
   Begin VB.CommandButton Recognise 
      Caption         =   "&Recognise"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton ClearScreen 
      Caption         =   "&Clear Screen"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   5640
      Width           =   2535
   End
   Begin VB.CommandButton Exit 
      BackColor       =   &H80000012&
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Frame Description 
      Caption         =   "Description"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8535
      Begin VB.Label Label3 
         Caption         =   "Last Updated was on 22 of May, 2001"
         Height          =   255
         Left            =   2880
         TabIndex        =   21
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "-= Text Recognition Software =-"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Design by Albert Archwamety"
         Height          =   255
         Left            =   6120
         MousePointer    =   2  'Cross
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.PictureBox picboxDrawArea 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DrawWidth       =   100
      FillStyle       =   0  'Solid
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   240
      MouseIcon       =   "frmMain.frx":0750
      MousePointer    =   99  'Custom
      ScaleHeight     =   2505
      ScaleWidth      =   2505
      TabIndex        =   2
      ToolTipText     =   "Please Draw a character"
      Top             =   1320
      Width           =   2535
   End
   Begin VB.PictureBox picboxDatabaseArea 
      BackColor       =   &H80000005&
      Height          =   2535
      Left            =   3120
      ScaleHeight     =   2475
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.PictureBox picboxDataArea 
      BackColor       =   &H80000005&
      Height          =   2535
      Left            =   6000
      ScaleHeight     =   2475
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Frame frameArea1 
      Caption         =   "User's Draw Area"
      Height          =   3015
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   2775
   End
   Begin VB.Frame frameArea2 
      Caption         =   "Database Area"
      Height          =   3015
      Left            =   3000
      TabIndex        =   17
      Top             =   960
      Width           =   2775
   End
   Begin VB.Frame frameArea3 
      Caption         =   "Buffer Data Area"
      Height          =   3015
      Left            =   5880
      TabIndex        =   18
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   5040
      TabIndex        =   22
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label TeachLabelText 
      Caption         =   "Enter a character to be teach"
      Height          =   660
      Left            =   360
      TabIndex        =   12
      Top             =   4215
      Width           =   4815
   End
   Begin VB.Label ResultLabel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   15
      Top             =   6195
      Width           =   5895
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUp_About 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuPopUp_Close 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Main_TeachText As String

Dim strCaption As String
Dim RECOG_EXT As String
Dim DrawNow As Boolean
Dim c As Integer
Dim strData As String
Dim strRECpk As String
Dim arrRawData(100 * 100) As String
Dim arrTagData(100 * 100) As String

Private Sub ClearScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.StatusLabel.Caption = StatusWindow("ClearScreenButton")
End Sub

Private Sub comboOpen_Click()
'    Me.comboOpen.Visible = False

        c = Me.comboOpen.ListIndex
        
        a = 190
        b = 190
        c = c + 1
        d = 0
        Me.picboxDataArea.Cls
            
        strData = ""
            
        If c > 0 Then
            For i = 1 To 10
            For j = 1 To 10
                If Mid(arrRawData(c - 1), d + 1, 1) = vbBlack Then
                    strData = strData & vbBlack
'                    picboxDataArea.PSet (i, j)
                    picboxDataArea.PSet (a, b)
'                    picboxDataArea.Circle (a, b), 110
                    picboxDataArea.Line (a - 110, b - 110)-(a + 110, b - 110)
                    picboxDataArea.Line (a + 110, b - 110)-(a + 110, b + 110)
                    picboxDataArea.Line (a + 110, b + 110)-(a - 110, b + 110)
                    picboxDataArea.Line (a - 110, b + 110)-(a - 110, b - 110)
                    Debug.Print ""
                Else
                    strData = strData & 1
                End If
                d = d + 1
                b = b + (Me.picboxDataArea.Height - 200) / 10
            Next j
            b = 190
            a = a + (Me.picboxDataArea.Width - 200) / 10
            Next i
        End If
            
    Me.Open.Caption = "&Delete"
    Me.Open.Enabled = True
    
End Sub

Private Sub Exit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.StatusLabel.Caption = StatusWindow("ExitButton")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.StatusLabel.Caption = StatusWindow("Form")
End Sub

Private Sub Form_Paint()
Dim a As Integer
Dim b As Integer
Dim d As Integer

a = 190
b = 190
d = 0

If strData <> "" Then
    For i = 1 To 10
        For j = 1 To 10
            If Mid(strData, d + 1, 1) = vbBlack Then
                picboxDataArea.PSet (a, b)
                picboxDataArea.Line (a - 110, b - 110)-(a + 110, b - 110)
                picboxDataArea.Line (a + 110, b - 110)-(a + 110, b + 110)
                picboxDataArea.Line (a + 110, b + 110)-(a - 110, b + 110)
                picboxDataArea.Line (a - 110, b + 110)-(a - 110, b - 110)
            End If
            d = d + 1
            b = b + (Me.picboxDataArea.Height - 200) / 10
        Next j
        b = 190
        a = a + (Me.picboxDataArea.Width - 200) / 10
    Next i
End If

End Sub

Private Sub Open_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.StatusLabel.Caption = StatusWindow("OpenButton")
End Sub

Private Sub picboxDataArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If StatusLabel <> "DataArea" Then
        Me.StatusLabel.Caption = StatusWindow("DataArea")
'    End If
End Sub

Private Sub picboxDatabaseArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If StatusLabel <> "DatabaseArea" Then
        Me.StatusLabel.Caption = StatusWindow("DatabaseArea")
'    End If
End Sub

Private Sub Recognise_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.StatusLabel.Caption = StatusWindow("RecogniseButton")
End Sub

Private Sub Teach_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.StatusLabel.Caption = StatusWindow("TeachButton")
End Sub

Private Sub TeachCancel_Click()
    
    Me.TeachLabelText.FontBold = True
    Me.TeachLabelText.Caption = strCaption
    
    Me.TeachConfirm.Visible = False
    Me.TeachCancel.Visible = False
    Me.TeachText.Visible = False
    
    Me.Open.Visible = True
    Me.Teach.Visible = True
    Me.Recognise.Visible = True
    Me.ClearScreen.Visible = True
    Me.Exit.Visible = True
    
End Sub

Private Sub ClearScreen_Click()
'    MsgBox "Clear Screen...", vbOKOnly, "Run into sub function..."

    strData = ""

    Me.ResultLabel.Caption = ""
    Me.ResultLabel.ToolTipText = ""
    Me.Refresh

    picboxDrawArea.Cls
    picboxDatabaseArea.Cls
    picboxDataArea.Cls

    Me.comboRecognise.Visible = False
    Me.comboOpen.Visible = False
    Me.Open.Caption = "&Open"
    Me.Open.Enabled = True
    
    Me.Teach.Enabled = False
    Me.Recognise.Enabled = False
    Me.ClearScreen.Enabled = False
    
End Sub

Private Sub TeachCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.StatusLabel.Caption = StatusWindow("TeachCancelButton")
End Sub

Private Sub TeachConfirm_Click()
Dim Filename_Database As String
Dim Filename_Teach As String
Dim Buffer_DrawArea As Variant
Dim strTeachText As String
Dim intCounter As Integer
Dim strBuffer As String
'Dim TeachDialog As New frmTeach
'Dim oFile As TextStream

'    Set oFile = New TextStream
    
    FileSystem.ChDir (App.Path)
    picboxDataArea.Cls
    strTeachText = Me.TeachText.Text
    
    intCounter = 0
    
    Filename_Database = "DATA" & RECOG_EXT
    Filename_Teach = Filename_Database
  
    intCounter = intCounter + 1
    
    Call GraspRawData
    
    If strData = "" Then
        Call ClearScreen_Click
        Me.Teach.Enabled = Not Me.Teach.Enabled
        MsgBox "Detect No character was drawn in the Draw Area. Teach operation can not be proceed. ", vbExclamation, "Warning..."
        GoTo TeachConfirm_SkipTeach
    End If

    Open Filename_Teach For Binary As #1
        strBuffer = Space(5)
        Get #1, , strBuffer
    Close #1
    
    If strBuffer = "recPK" Then
' add teaching character as binary
'**
        strRECpk = ""
        strBuffer = ""
        Open Filename_Teach For Binary As #1
            strBuffer = Space(5)
            Get #1, , strBuffer
            strRECpk = strRECpk & strBuffer
            strBuffer = Space(22)
            While Not EOF(1)
                Get #1, , strBuffer
                strRECpk = strRECpk & strBuffer
            Wend
        Close #1
        strRECpk = Mid(strRECpk, 1, Len(strRECpk) - 22)
        i = 3
        strRECpk = strRECpk & strTeachText
        strRECpk = strRECpk & ","
        For j = 1 To 10
            strRECpk = strRECpk & Chr(BinToDec(Mid(strData, i - 2 + ((j - 1) * 10), 2)))
            strRECpk = strRECpk & Chr(BinToDec(Mid(strData, i + 0 + ((j - 1) * 10), 8)))
        Next j
        Open Filename_Teach For Binary As #1
            Put #1, , strRECpk
        Close #1
'**
    Else
' add teaching character as string
'**
        Open Filename_Teach For Append As #1
            Write #1, strTeachText & "," & strData
        Close #1
'**
    End If

'    ** Manually key in teach file into database **
'    MsgBox "Teaching..."

'    On Error GoTo TeachErrorHandler
'    Teach_CommonDialog.DialogTitle = "Teach"
'    Teach_CommonDialog.Filter = "Recognised Files (*.rec)|*.rec|All Files (*.*)|*.*"
'    Teach_CommonDialog.DefaultExt = ".rec"
'    Teach_CommonDialog.InitDir = App.Path
'    Teach_CommonDialog.ShowSave
    
'    If Teach_CommonDialog.FileName <> "" Then
'        Filename_Teach = Teach_CommonDialog.FileName
'        Filename_Teach = Mid(Filename_Teach, InStrRev(Filename_Teach, "\") + 1)
    

'    With oFile
'        If .OpenTextFile(Filename_Teach, ForAppending) Then
'            Call .WriteLine(Me.picboxDrawArea)
'        Else
'            MsgBox "Error opening text file"
'        End If
'
'        .CloseFile
'    End With
'
'    Set oFile = Nothing
    
'TeachErrorHandler:
'    Exit Sub
    
TeachConfirm_SkipTeach:
    
    Me.TeachLabelText.FontBold = True
    Me.TeachLabelText.Caption = strCaption
    
    Me.TeachConfirm.Visible = False
    Me.TeachCancel.Visible = False
    Me.TeachText.Visible = False
    
    Me.Teach.Enabled = False

    Me.Open.Visible = True
    Me.Teach.Visible = True
    Me.Recognise.Visible = True
    Me.ClearScreen.Visible = True
    Me.Exit.Visible = True

End Sub

Private Sub picboxDrawArea_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    DrawNow = True
    
    picboxDrawArea.DrawWidth = 17
    picboxDrawArea.PSet (X, Y)

    If Not Me.Teach.Enabled Then
        Me.comboRecognise.Visible = False
        Me.comboOpen.Visible = False
        Me.Open.Enabled = True
        Me.Open.Caption = "&Open"
        
        Me.Teach.Enabled = True
        Me.Recognise.Enabled = True
        Me.ClearScreen.Enabled = True
    End If
    
    If Not Me.Recognise.Enabled Then
        Me.Recognise.Enabled = True
        Me.comboRecognise.Visible = False
    End If

End Sub

Private Sub picboxDrawArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    picboxDrawArea.DrawStyle = vbSolid
    Me.StatusLabel.Caption = StatusWindow("DrawArea")
    picboxDrawArea.DrawWidth = 17
    If DrawNow Then
        picboxDrawArea.PSet (X, Y)
        
        If Not fMainForm.Teach.Enabled Then
            Me.Teach.Enabled = True
            Me.Recognise.Enabled = True
            Me.ClearScreen.Enabled = True
        End If
    End If
End Sub

Private Sub picboxDrawArea_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawNow = False
    
    Call GraspRawData
End Sub

Private Sub Exit_Click()

'    MsgBox "Are you sure want to Exit?", vbYesNo, "Confirmation"
'    If vbYes Then
'      Unload Me
'    End If

'        frmConfirmation.Left = (frmMain.Width / 2) - (frmConfirmation.Width / 2)
'        frmConfirmation.Top = (frmMain.ScaleHeight / 2) - (frmConfirmation.Height)

        frmConfirmation.Show vbModal
        If frmConfirmation.YES Then
            Unload Me
        End If
End Sub

Private Sub Form_Load()
    Me.Top = 2
    Me.Left = 2
    Me.Width = 8820
    Me.Height = 6945
    RECOG_EXT = ".rec"
    
    strCaption = "Please pay a visit at 'http://come.to/albert.com/' or send found bugs of this software to albertoycc@hotmail.com"
    Me.TeachLabelText.FontBold = True
    Me.TeachLabelText.Caption = strCaption
    
    Me.picboxDataArea.DrawWidth = 2
    Me.picboxDatabaseArea.DrawWidth = 2
    
    Me.Teach.Enabled = False
    Me.Recognise.Enabled = False
    Me.ClearScreen.Enabled = False
    
    Me.TeachText.Visible = False
    Me.TeachConfirm.Visible = False
    Me.TeachCancel.Visible = False
    
    Me.pbRecognising.Visible = False
    Me.pbRecognising.Top = 3840
    Me.pbRecognising.Left = 3120
    Me.comboRecognise.Visible = False
    Me.comboOpen.Visible = False
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnuPopUp
    End If
End Sub

Private Sub Label1_Click()
'Dim WebBrowser As New frmBrowser
'    WebBrowser.Show vbModal

    Me.Label1.FontItalic = Not Me.Label1.FontItalic
    Me.Label1.FontBold = Not Me.Label1.FontBold
    
    If Me.Label1.FontItalic Then
        Me.Label1.Caption = "albertoycc@hotmail.com"
    Else
        Me.Label1.Caption = "Design by Albert Archwamety"
    End If
    
End Sub

Private Sub mnuPopUp_About_Click()
Dim AboutDialog As New frmAbout
    AboutDialog.Show vbModal
End Sub

Private Sub mnuPopUp_Close_Click()
    Call Exit_Click
End Sub

Private Sub Open_Click()
Dim Filename_Open As String
Dim Buffer_DataArea As Variant
Dim boolDeleteItemDetect As Boolean
Dim strBuffer As String

If Me.Open.Caption = "&Delete" Then
    If MsgBox("Are you sure want to Delete current pattern?", vbYesNo + vbQuestion, "Confirmation...") = vbYes Then
        Open Me.Open_CommonDialog.FileName For Output As #1
        i = 0
        While arrTagData(i) <> ""
            If Me.comboOpen.ListIndex = i Or boolDeleteItemDetect Then
                If Me.comboOpen.ListIndex = i Then
                    boolDeleteItemDetect = True
                    Me.comboOpen.RemoveItem (i)
                    If arrTagData(i + 1) <> "" Then
                        Me.comboOpen.ListIndex = i
                    End If
                End If
                arrTagData(i) = arrTagData(i + 1)
                arrRawData(i) = arrRawData(i + 1)
            End If
            
            If arrTagData(i) <> "" Then
                Write #1, arrTagData(i) & "," & arrRawData(i)
            End If
            
            i = i + 1
            
        Wend
        Close #1
    Else
    End If
Else

'Dim oFile As TextStream
Dim strValue As String

'    Set oFile = New TextStream

'    MsgBox "Opening..."

'    On Error GoTo OpenErrorHandler
    Open_CommonDialog.FileName = ""
    Open_CommonDialog.DialogTitle = "Open"
    Open_CommonDialog.Filter = "Recognised Files (*.rec)|*.rec|All Files (*.*)|*.*"
    Open_CommonDialog.DefaultExt = ".rec"
    Open_CommonDialog.InitDir = App.Path
    Open_CommonDialog.ShowOpen

'    If Open_CommonDialog.ShowOpen = vbCancel Then
'        GoTo OpenCancelClicked
'    End If

    If Open_CommonDialog.FileName <> "" Then

        Filename_Open = Open_CommonDialog.FileName
        Filename_Open = Mid(Filename_Open, InStrRev(Filename_Open, "\") + 1)
            
        picboxDataArea.Cls
        
        Open Filename_Open For Binary As #1
            i = 0
            strBuffer = Space(5)
                Get #1, , strBuffer
                arrRawData(i) = strBuffer
                i = i + 1
            strBuffer = Space(22)
            While Not EOF(1)
                Get #1, , strBuffer
                arrRawData(i) = strBuffer
                i = i + 1
            Wend
            arrRawData(i - 1) = ""
            Close #1
            
            i = 0
            strBuffer = ""
            If arrRawData(0) = "recPK" Then
                i = i + 1
                While arrRawData(i) <> ""
                
                    arrTagData(i - 1) = Mid(arrRawData(i), 1, 1)
'                    strBuffer = strBuffer & Mid(strData(i), 1, 1) & vbCrLf
'                    strBuffer = strBuffer & Mid(strData(0), i, 1) & ","
'                    Me.ListBox_List.AddItem i & ". <" & Mid(strData(i), 1, 1) & ">"
                    
                    arrRawData(i - 1) = ""
                    
                    For j = 1 To 10
'                        strBuffer = strBuffer & DecToBin(Asc(Mid(strData(i), 3 + ((j - 1) * 2), 1)), 2)
'                        strBuffer = strBuffer & DecToBin(Asc(Mid(strData(i), 4 + ((j - 1) * 2), 1)), 8)
                        arrRawData(i - 1) = arrRawData(i - 1) & _
                                            DecToBin(Asc(Mid(arrRawData(i), 3 + ((j - 1) * 2), 1)), 2) & _
                                            DecToBin(Asc(Mid(arrRawData(i), 4 + ((j - 1) * 2), 1)), 8)
'                        strBuffer = strBuffer
                    Next j
                    
'                    strBuffer = strBuffer & vbCrLf
                    i = i + 1
                    
                Wend
'                Me.RichTextBox_Text.Text = strBuffer
        '        Me.TextBox_Text.ScrollBars
                
                arrTagData(i - 1) = ""
                arrRawData(i - 1) = ""
   
            Else
            
                Filename_Open = Open_CommonDialog.FileName
                Filename_Open = Mid(Filename_Open, InStrRev(Filename_Open, "\") + 1)
            
                picboxDataArea.Cls
                
                c = 0
                
                Open Filename_Open For Input As #1
                While Not EOF(1)
Open_SkipLine:
                a = 190
                b = 190
                d = 0
                Me.picboxDataArea.Cls
                    If EOF(1) Then
                        GoTo Open_FileClose
                    End If
                    Input #1, arrRawData(c)
                    If Len(arrRawData(c)) < 102 Then
                        GoTo Open_SkipLine
                    End If
                    arrTagData(c) = Mid(arrRawData(c), 1, 1)
                    arrRawData(c) = Mid(arrRawData(c), 3)
        '            Debug.Print arrRawData(c)
                    c = c + 1
                    
        '''                For i = 1 To 10
        '''                For j = 1 To 10
        '''                    If Mid(arrRawData(c - 1), d + 1, 1) = vbBlack Then
        '''                        picboxDataArea.PSet (a, b)
        '''    '                    picboxDataArea.Circle (a, b), 110
        '''    '''                    picboxDataArea.Line (a - 110, b - 110)-(a + 110, b - 110)
        '''    '''                    picboxDataArea.Line (a + 110, b - 110)-(a + 110, b + 110)
        '''    '''                    picboxDataArea.Line (a + 110, b + 110)-(a - 110, b + 110)
        '''    '''                    picboxDataArea.Line (a - 110, b + 110)-(a - 110, b - 110)
        '''                    End If
        '''                    d = d + 1
        '''                    b = b + (Me.picboxDataArea.Height - 200) / 10
        '''                Next j
        '''                b = 190
        '''                a = a + (Me.picboxDataArea.Width - 200) / 10
        '''                Next i
                Wend
Open_FileClose:
                Close #1
            
                arrTagData(c) = ""
                arrRawData(c) = ""
                
            End If
    
    End If
    
'   With oFile
'      .FileName = Filename_Open
'      ' Check For File Too Big - 32K limit on text boxes
'      If .FileTooBig Then
'         Beep
'         MsgBox "File Too Big To Read", , "File Open Error"
'      Else
'         If .OpenTextFile(Filename_Open, ForReading) Then
'            Do Until .AtEndOfStream
'               strValue = strValue & .ReadLine & vbCrLf
'            Loop
'            .CloseFile
'            MsgBox strValue
'         End If
'      End If
'   End With
'   Set oFile = Nothing

'OpenCancelClicked:
'OpenErrorHandler:
'    Exit Sub

    If Open_CommonDialog.FileName <> "" Then
        Me.Open.Enabled = False
        Me.Teach.Enabled = False
        Me.Recognise.Enabled = False
        Me.ClearScreen.Enabled = True
        Me.comboOpen.Visible = True
        Me.comboOpen.Text = "Select a character to open ..."
        
        i = 0
        While Me.comboOpen.List(i) <> ""
            Me.comboOpen.RemoveItem (i)
        Wend
        i = 0
        While arrTagData(i) <> ""
            Me.comboOpen.AddItem i + 1 & ". - <" & arrTagData(i) & ">"
            i = i + 1
        Wend
        
        Dim lngRet As Long
            lngRet = SendMessage(Me.comboOpen.hwnd, _
                                CB_SHOWDROPDOWN, _
                                1, _
                                0&)
            
    End If
    
End If
    
End Sub

Private Sub Recognise_Click()
Dim Filename_Database As String
Dim strRecognised As String
Dim intMatch As Integer
Dim intMaxMatch As Integer
Dim intCounter As Integer
Dim boolFindLastFile As Boolean
Dim boolNoMoreFileLeft As Boolean
Dim buffer As String
Dim Buffer_DatabaseArea As Variant

Dim strBuffer As String

'    MsgBox "Recognising...", vbOKOnly, "Run into sub function..."

'    Call GraspRawData
    
    FileSystem.ChDir (App.Path)
    
    Me.pbRecognising.Visible = True
    
    strRecognised = ""
    intMaxMatch = 0
    intCounter = 0
    c = 0
    
    On Error GoTo Recognise_FileClose
    Filename_Database = "DATA" & RECOG_EXT
    
    If Filename_Database <> "" Then
        Me.picboxDatabaseArea.Cls
        Open Filename_Database For Binary As #1
            i = 0
            strBuffer = Space(5)
                Get #1, , strBuffer
                arrRawData(i) = strBuffer
                i = i + 1
            strBuffer = Space(22)
            While Not EOF(1)
                Get #1, , strBuffer
                arrRawData(i) = strBuffer
                i = i + 1
            Wend
            arrRawData(i - 1) = ""
            Close #1
            
            i = 0
            strBuffer = ""
            If arrRawData(0) = "recPK" Then
                i = i + 1
                
                strRecognised = ""
                intMaxMatch = 0
                intCounter = 0
                c = 0
                
                While arrRawData(i) <> ""
                
                    a = 190
                    b = 190
                    d = 0
                    intMatch = 0
                    Me.picboxDatabaseArea.Cls
                
                    arrTagData(i - 1) = Mid(arrRawData(i), 1, 1)
'                    strBuffer = strBuffer & Mid(strData(i), 1, 1) & vbCrLf
'                    strBuffer = strBuffer & Mid(strData(0), i, 1) & ","
'                    Me.ListBox_List.AddItem i & ". <" & Mid(strData(i), 1, 1) & ">"
                    
                    arrRawData(i - 1) = ""
                    
                    For j = 1 To 10
'                        strBuffer = strBuffer & DecToBin(Asc(Mid(strData(i), 3 + ((j - 1) * 2), 1)), 2)
'                        strBuffer = strBuffer & DecToBin(Asc(Mid(strData(i), 4 + ((j - 1) * 2), 1)), 8)
                        arrRawData(i - 1) = arrRawData(i - 1) & _
                                            DecToBin(Asc(Mid(arrRawData(i), 3 + ((j - 1) * 2), 1)), 2) & _
                                            DecToBin(Asc(Mid(arrRawData(i), 4 + ((j - 1) * 2), 1)), 8)
'                        strBuffer = strBuffer
                    Next j
                    
'                    strBuffer = strBuffer & vbCrLf
                    i = i + 1

                    For ii = 1 To 10
                    For jj = 1 To 10
                        If Mid(arrRawData(c), d + 1, 1) = vbBlack Then
        '                    picboxDataArea.PSet (i, j)
' j2 - mask - finalise
'                            picboxDatabaseArea.PSet (a, b)

        '                    picboxDataArea.Circle (a, b), 110
        '                    picboxDatabaseArea.Line (a - 110, b - 110)-(a + 110, b - 110)
        '                    picboxDatabaseArea.Line (a + 110, b - 110)-(a + 110, b + 110)
        '                    picboxDatabaseArea.Line (a + 110, b + 110)-(a - 110, b + 110)
        '                    picboxDatabaseArea.Line (a - 110, b + 110)-(a - 110, b - 110)
        '                    Debug.Print ""
                            If Mid(strData, d + 1, 1) = vbBlack Then
                                intMatch = intMatch + 1
                            Else
                                intMatch = intMatch - 1
                            End If
                        Else
                            If Mid(strData, d + 1, 1) <> vbBlack Then
                                intMatch = intMatch + 1
                            Else
                                intMatch = intMatch - 1
                            End If
                        End If
                        d = d + 1
                        b = b + (Me.picboxDatabaseArea.Height - 200) / 10
                    Next jj
                    b = 190
                    a = a + (Me.picboxDatabaseArea.Width - 200) / 10
                    Next ii
                    If intMaxMatch < intMatch Then
                        intMaxMatch = intMatch
                        strRecognised = arrTagData(c)
                        intCounter = c
                        Me.pbRecognising.Value = intMaxMatch
                        If intMaxMatch > 90 Then
                            GoTo Recognise_FileClose
                        End If
                    End If
                    c = c + 1
                
                Wend
'                Me.RichTextBox_Text.Text = strBuffer
        '        Me.TextBox_Text.ScrollBars
                
                arrTagData(i - 1) = ""
                arrRawData(i - 1) = ""
   
            Else
            
                strRecognised = ""
                intMaxMatch = 0
                intCounter = 0
                c = 0
                
                Open Filename_Database For Input As #1
                While Not EOF(1)
Recognise_SkipLine:
                a = 190
                b = 190
                d = 0
                intMatch = 0
                Me.picboxDatabaseArea.Cls
                    If EOF(1) Then
                        GoTo Recognise_FileClose
                    End If
                    Input #1, arrRawData(c)
                    If Len(arrRawData(c)) < 102 Then
                        GoTo Recognise_SkipLine
                    End If
                    arrTagData(c) = Mid(arrRawData(c), 1, 1)
                    arrRawData(c) = Mid(arrRawData(c), 3)
        '            Debug.Print arrRawData(c)
                    
                    For i = 1 To 10
                    For j = 1 To 10
                        If Mid(arrRawData(c), d + 1, 1) = vbBlack Then
        '                    picboxDataArea.PSet (i, j)
                            picboxDatabaseArea.PSet (a, b)
        '                    picboxDataArea.Circle (a, b), 110
        '                    picboxDatabaseArea.Line (a - 110, b - 110)-(a + 110, b - 110)
        '                    picboxDatabaseArea.Line (a + 110, b - 110)-(a + 110, b + 110)
        '                    picboxDatabaseArea.Line (a + 110, b + 110)-(a - 110, b + 110)
        '                    picboxDatabaseArea.Line (a - 110, b + 110)-(a - 110, b - 110)
        '                    Debug.Print ""
                            If Mid(strData, d + 1, 1) = vbBlack Then
                                intMatch = intMatch + 1
                            Else
                                intMatch = intMatch - 1
                            End If
                        Else
                            If Mid(strData, d + 1, 1) <> vbBlack Then
                                intMatch = intMatch + 1
                            Else
                                intMatch = intMatch - 1
                            End If
                        End If
                        d = d + 1
                        b = b + (Me.picboxDatabaseArea.Height - 200) / 10
                    Next j
                    b = 190
                    a = a + (Me.picboxDatabaseArea.Width - 200) / 10
                    Next i
                    If intMaxMatch < intMatch Then
                        intMaxMatch = intMatch
                        strRecognised = arrTagData(c)
                        intCounter = c
                        Me.pbRecognising.Value = intMaxMatch
                        If intMaxMatch > 90 Then
                            GoTo Recognise_FileClose
                        End If
                    End If
                    c = c + 1
                Wend
Recognise_FileClose:
                Close #1
            End If
    End If
    
    Me.Recognise.Enabled = False
    Me.pbRecognising.Visible = False
    
    If strRecognised <> "" And intMaxMatch >= 68 Then
        picboxDatabaseArea.Cls
        a = 190
        b = 190
        d = 0
        For i = 1 To 10
        For j = 1 To 10
            If Mid(arrRawData(intCounter), d + 1, 1) = vbBlack Then
                picboxDatabaseArea.PSet (a, b)
                picboxDatabaseArea.Line (a - 110, b - 110)-(a + 110, b - 110)
                picboxDatabaseArea.Line (a + 110, b - 110)-(a + 110, b + 110)
                picboxDatabaseArea.Line (a + 110, b + 110)-(a - 110, b + 110)
                picboxDatabaseArea.Line (a - 110, b + 110)-(a - 110, b - 110)
            End If
            d = d + 1
            b = b + (Me.picboxDatabaseArea.Height - 200) / 10
        Next j
        b = 190
        a = a + (Me.picboxDatabaseArea.Width - 200) / 10
        Next i
    End If
    
    If strRecognised <> "" And intMaxMatch >= 68 Then
        'The highest posible of drawn character is recognised as   'X'
        Me.ResultLabel.Caption = "The highest posible of drawn character is recognised as   '" & strRecognised & "'"
        Me.ResultLabel.ToolTipText = intMaxMatch & "%"
        '& " , " & intMaxMatch & "%"
        
        Me.DrawWidth = 2
    '    me
    '    Me.Circle (6120, 6120), 170, vbYellow
    '    Me.Circle (6120, 6120), 180, vbBlack
        
    '    Me.Circle (5240, 6300), 190, vbYellow
    
        Me.Circle (5180, 6300), 190, vbYellow
        Me.Circle (5180, 6300), 210, vbBlack
    
        Me.comboRecognise.Visible = True
        With Me.comboRecognise
        
        End With
        
    Else
        If MsgBox("No character has been teach OR Character not drawn properly OR User has drawn more than one character OR Run Text Recognition for the 1st. time, Please click 'Yes' to Teach or 'No' to discard drawn character.", vbExclamation + vbYesNo, "Run Text Recognition for the 1st. time?") = vbYes Then
            Call Teach_Click
        Else
            MsgBox intMaxMatch & "% Match with character '" & strRecognised & "'"
        End If
    End If
    
''''    strMatch = "a"
''''    intMatch = 0
''''
''''    Do
''''    intCounter = 0
''''    intPoint = 0
''''    Buffer = Str$(intMatch) & "data"
''''
''''        Do
''''        On Error GoTo Recognise_Anchor_LastFile
''''Recognise_Anchor_Search_UntilNoFile:
''''        Filename_Database = Buffer & intCounter & RECOG_EXT
''''        Open Filename_Database For Input As #1
''''        c = 0
''''            For i = 1 To picboxDrawArea.Width Step 50
''''            For j = 1 To picboxDrawArea.Height Step 50
''''                If EOF(1) Then
''''                    GoTo Recognise_FileClose
''''                End If
''''                Input #1, Buffer_DatabaseArea
''''                RawDatabase(c) = Buffer_DatabaseArea
'''''                Debug.Print RawData(c)
''''                c = c + 1
''''                If RawDatabase(c - 1) = vbBlack Then
''''                    picboxDatabaseArea.PSet (i, j)
''''                    If RawData(c - 1) = vbBlack Then
''''                        intPoint = intPoint + 1
''''                    End If
''''                End If
''''            Next j, i
''''
''''Recognise_FileClose:
''''        Close #1
''''
''''        If intPoint > intMaxPoint Then
''''            intMaxPoint = intPoint
''''            intMaxMatch = intMatch
''''        End If
''''
''''        intCounter = intCounter + 1
''''        GoTo Recognise_Anchor_Search_UntilNoFile
''''Recognise_Anchor_LastFile:
''''        Close #1
''''
''''        boolFindLastFile = True
''''
''''        Loop While Not boolFindLastFile
''''
''''
''''
''''    intMatch = intMatch + 1
''''
''''    If intMatch = 10 Then
''''        boolNoMoreFileLeft = True
''''    End If
''''    Loop While Not boolNoMoreFileLeft
'''''*********
''''
''''    MsgBox "intmaxpoint=" & intMaxPoint
''''    MsgBox "intmaxmatch=" & intMaxMatch

End Sub

Private Sub Teach_Click()

    Me.TeachLabelText.FontBold = False
    Me.TeachLabelText.Caption = "Enter a character to be teach"
    
    Me.Open.Visible = False
    Me.Teach.Visible = False
    Me.Recognise.Visible = False
    Me.ClearScreen.Visible = False
    Me.Exit.Visible = False
    
    Me.TeachConfirm.Visible = True
    Me.TeachCancel.Visible = True
    Me.TeachText.Visible = True

    Me.TeachText.Text = ""
    Me.TeachText.SetFocus
    Me.TeachConfirm.Enabled = False
    
End Sub

Private Sub GraspRawData()
Dim bool1stScan As Boolean
Dim ax As Integer
Dim ay As Integer
Dim bx As Integer
Dim by As Integer

bool1stScan = True
strData = ""

c = 0

Me.picboxDataArea.Cls
For i = 1 To picboxDrawArea.Width Step 100
    For j = 1 To picboxDrawArea.Height Step 100
        If picboxDrawArea.Point(i, j) = vbBlack Then
            picboxDataArea.PSet (i, j)
            If Not bool1stScan Then
                If i < ax Then
                    ax = i
                    End If
                If i > bx Then
                    bx = i
                    End If
                If j < ay Then
                    ay = j
                    End If
                If j > by Then
                    by = j
                    End If
            Else
                bool1stScan = False
                ax = i
                bx = i
                ay = j
                by = j
            End If
        End If
Next j, i

'MsgBox ""

If bx - ax <> 0 And by - ay <> 0 Then

    a = 190
    b = 190
    
    Me.picboxDataArea.Cls
    For i = ax To bx - (bx - ax) / 10 Step (bx - ax) / 10
        For j = ay To by - (by - ay) / 10 Step (by - ay) / 10
            If picboxDrawArea.Point(i, j) = vbBlack Then
                picboxDataArea.PSet (a, b)
    '''            picboxDataArea.Circle (a, b), 110
                picboxDataArea.Line (a - 110, b - 110)-(a + 110, b - 110)
                picboxDataArea.Line (a + 110, b - 110)-(a + 110, b + 110)
                picboxDataArea.Line (a + 110, b + 110)-(a - 110, b + 110)
                picboxDataArea.Line (a - 110, b + 110)-(a - 110, b - 110)
    '                    picboxDataArea.FillStyle = vbSolid
    '                    picboxDataArea.FillColor = vbBlack
    '                    picboxDataArea.fil
    '            RawData(c) = picboxDrawArea.Point(i, j)
                strData = strData & picboxDrawArea.Point(i, j)
                c = c + 1
            Else
    '            RawData(c) = 1
                strData = strData & 1
                c = c + 1
            End If
            b = b + (Me.picboxDataArea.Height - 200) / 10
        Next j
        b = 190
        a = a + (Me.picboxDataArea.Width - 200) / 10
    Next i

End If

End Sub

Private Sub TeachConfirm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.StatusLabel.Caption = StatusWindow("TeachConfirmButton")
End Sub

Private Sub TeachText_GotFocus()
    Me.TeachText.SelStart = 0
    Me.TeachText.SelLength = Len(Me.TeachText.Text)
End Sub

Private Sub TeachText_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(Me.TeachText.Text) = 1 Then
        Me.TeachConfirm.Enabled = True
        Me.TeachConfirm.SetFocus
    ElseIf Len(Me.TeachText.Text) > 1 Then
        Me.TeachText.Text = ""
        Me.TeachConfirm.Enabled = False
    Else
        Me.TeachConfirm.Enabled = False
    End If
End Sub

Private Function StatusWindow(Optional ByVal strBuffer As String) As String
Dim Status_OpenButton As String
Dim Status_TeachButton As String
Dim Status_TeachConfirmButton As String
Dim Status_TeachCancelButton As String
Dim Status_RecogniseButton As String
Dim Status_ClearScreenButton As String
Dim Status_ExitButton As String
Dim Status_DrawArea As String
Dim Status_DatabaseArea As String
Dim Status_DataArea As String
Dim Status_Form As String
    
    Status_OpenButton = "Tips && Help : Open File - Just click the Open button to open FILE then type in your FILENAME - Try it..."
    Status_TeachButton = "Tips && Help : Teach - Just click the Teach button to TEACH then type in your teach CHARACTER - Try it..."
    Status_TeachConfirmButton = "Tips && Help : Confirm the character ENTER in the textbox is match with the drawn character in the Draw Area."
    Status_TeachCancelButton = "Tips && Help : Mispressed or Give up or Do not want to Teach."
    Status_RecogniseButton = "Tips && Help : Recognise Text - Just click the this button to recognise drawn character then the recognised character will display at the BOTTOM. Try it..."
    Status_ClearScreenButton = "Tips && Help : Clear Screen - Just click the this button to CLEAR screen then continue to draw character. Try it..."
    Status_ExitButton = "Tips && Help : Exit this software."
    Status_DrawArea = "Tips && Help : Your mouse pointer is in the Draw Area, just Click and drag to draw a character..."
    Status_DatabaseArea = "Tips && Help : This is Database Area, which it will display a recognised character from the database when you click Recognise Button."
    Status_DataArea = "Tips && Help : This Data Area act as a buffer storage of Draw Area when user draw in Draw Area or to retrieve database data when user click Open Button."
    Status_Form = "Tips && Help : Thanks for using this software. Just move the mouse to the Draw Area, then drag the mouse when you want to draw a character to be recognised. Nice Try ..."
    
    Select Case strBuffer
    Case "OpenButton": strBuffer = Status_OpenButton
    Case "TeachButton": strBuffer = Status_TeachButton
    Case "TeachConfirmButton": strBuffer = Status_TeachConfirmButton
    Case "TeachCancelButton": strBuffer = Status_TeachCancelButton
    Case "RecogniseButton": strBuffer = Status_RecogniseButton
    Case "ClearScreenButton": strBuffer = Status_ClearScreenButton
    Case "ExitButton": strBuffer = Status_ExitButton
    Case "DrawArea": strBuffer = Status_DrawArea
    Case "DatabaseArea": strBuffer = Status_DatabaseArea
    Case "DataArea": strBuffer = Status_DataArea
    Case "Form": strBuffer = Status_Form
    End Select
    
    StatusWindow = strBuffer

End Function

Private Function BinToDec(strBin As String) As Integer

i = Len(strBin)
While i > 0
    If Mid(strBin, i, 1) = "1" Then
        BinToDec = BinToDec + 2 ^ (Len(strBin) - i)
    End If
    i = i - 1
Wend

End Function

Private Function DecToBin(intDec As Integer, intDigit As Integer) As String
Dim intTemp As Integer

While intDec > 0 And intDigit > 0
    intDigit = intDigit - 1
    intTemp = intDec Mod 2
    If intTemp Then
        DecToBin = "1" & DecToBin
        intDec = (intDec - 1) / 2
    Else
        DecToBin = "0" & DecToBin
        intDec = intDec / 2
    End If
Wend

While intDigit
    intDigit = intDigit - 1
    DecToBin = "0" & DecToBin
Wend

End Function

