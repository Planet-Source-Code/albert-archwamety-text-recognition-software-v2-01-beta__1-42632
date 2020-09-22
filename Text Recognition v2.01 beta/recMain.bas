Attribute VB_Name = "recMain"
Public fMainForm As frmMain

Sub Main()

'    Call LoginDialog
    
    frmSplash.Show
    frmSplash.Refresh
    
    Set fMainForm = New frmMain
    Load fMainForm
    
    Unload frmSplash

    fMainForm.Show

End Sub

Private Sub LoginDialog()
Dim fLogin As New frmLogin
    fLogin.Show vbModal
    If Not fLogin.OK Then
        'Login Failed so exit app
        End
    End If
    Unload fLogin
End Sub


