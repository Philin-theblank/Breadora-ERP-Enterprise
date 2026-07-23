'================================================================================
'  frmLogin
'  Purpose: User authentication form
'  Author: Chokcity Technologies
'  Version: 1.0.0
'  Date: 2026-07-23
'================================================================================

Option Explicit

Private Sub UserForm_Initialize()
    'Center form on screen
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    
    'Set focus to username field
    Me.txtUsername.SetFocus
    
    'Set form properties
    Me.Caption = mdlConfiguration.APP_NAME & " v" & mdlConfiguration.APP_VERSION
    Me.lblTitle.Caption = "User Login"
    Me.lblVersion.Caption = "Version " & mdlConfiguration.APP_VERSION
    
    'Clear fields
    Me.txtUsername.Value = ""
    Me.txtPassword.Value = ""
    Me.lblMessage.Caption = ""
    Me.lblMessage.ForeColor = mdlConfiguration.COLOR_DANGER
End Sub

''
' cmdLogin_Click
' Purpose: Handle login button click
'
Private Sub cmdLogin_Click()
    On Error GoTo ErrorHandler
    
    Dim username As String
    Dim password As String
    Dim user As clsUser
    
    'Validate input
    username = Trim(Me.txtUsername.Value)
    password = Me.txtPassword.Value
    
    If Len(username) = 0 Then
        Me.lblMessage.Caption = "Please enter username"
        Me.txtUsername.SetFocus
        Exit Sub
    End If
    
    If Len(password) = 0 Then
        Me.lblMessage.Caption = "Please enter password"
        Me.txtPassword.SetFocus
        Exit Sub
    End If
    
    'Attempt authentication
    Set user = mdlUserManagement.AuthenticateUser(username, password)
    
    If user Is Nothing Then
        Me.lblMessage.Caption = "Login failed. Invalid username or password."
        Me.txtPassword.Value = ""
        Me.txtPassword.SetFocus
        Exit Sub
    End If
    
    'Start session
    If mdlSessionManagement.StartSession(user) Then
        Me.lblMessage.Caption = ""
        Me.lblMessage.ForeColor = mdlConfiguration.COLOR_SUCCESS
        MsgBox "Welcome " & user.FullName & "!", vbInformation, mdlConfiguration.APP_NAME
        
        'Close login form
        Unload Me
        
        'Open main dashboard
        frmDeveloperPanel.Show
    Else
        Me.lblMessage.Caption = "Login failed. Please try again."
        Me.txtPassword.Value = ""
        Me.txtPassword.SetFocus
    End If
    
    Exit Sub
ErrorHandler:
    Me.lblMessage.Caption = "Error: " & Err.Description
    mdlConfiguration.LogAuditTrail username, mdlConfiguration.AUDIT_LOGIN, "frmLogin", 0, "", "", "Failed", Err.Description
End Sub

''
' cmdCancel_Click
' Purpose: Handle cancel button click
'
Private Sub cmdCancel_Click()
    Unload Me
End Sub

''
' txtPassword_KeyDown
' Purpose: Handle Enter key in password field
'
Private Sub txtPassword_KeyDown(ByVal KeyCode As MSForms.ReturnType, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdLogin_Click
    End If
End Sub

''
' txtUsername_KeyDown
' Purpose: Handle Enter key in username field
'
Private Sub txtUsername_KeyDown(ByVal KeyCode As MSForms.ReturnType, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdLogin_Click
    End If
End Sub
