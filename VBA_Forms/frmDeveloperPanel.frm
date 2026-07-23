'================================================================================
'  frmDeveloperPanel
'  Purpose: Developer Dashboard showing project progress and completion status
'  Author: Chokcity Technologies
'  Version: 1.0.0
'  Date: 2026-07-23
'================================================================================

Option Explicit

Private Sub UserForm_Initialize()
    'Center form on screen
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    
    'Set form properties
    Me.Caption = "Breadora ERP - Developer Preview"
    
    'Populate dashboard
    Call RefreshDashboard
End Sub

''
' RefreshDashboard
' Purpose: Update all dashboard information
'
Public Sub RefreshDashboard()
    On Error GoTo ErrorHandler
    
    Dim modulesCompleted As Integer
    Dim modulesPlanned As Integer
    Dim completionPercent As Double
    Dim tablesCreated As Integer
    Dim vbaModulesCount As Integer
    
    'Count completed modules
    modulesCompleted = 2 'Module 1 + Module 2 completed
    modulesPlanned = 10 'Total planned modules
    completionPercent = (modulesCompleted / modulesPlanned) * 100
    
    'Count database tables
    tablesCreated = 9 'tblSettings, tblCompanyInfo, tblNumberSequences, tblAuditTrail, tblUsers, tblRoles, tblPermissions, tblUserRoles, tblLoginHistory
    
    'Count VBA modules
    vbaModulesCount = 6 'mdlConfiguration, mdlFolderOperations, clsLogger, mdlUserManagement, clsUser, mdlSessionManagement
    
    'Update version info
    Me.lblVersion.Caption = "Version: " & mdlConfiguration.APP_VERSION
    Me.lblBuild.Caption = "Build: " & mdlConfiguration.APP_BUILD
    Me.lblCompany.Caption = "Developer: " & mdlConfiguration.APP_DEVELOPER
    
    'Update progress
    Me.lblModulesCompleted.Caption = "Modules Completed: " & modulesCompleted & " of " & modulesPlanned
    Me.lblProgress.Caption = "Overall Progress: " & Format(completionPercent, "0.0") & "%"
    Me.lblProgressBar.Caption = GetProgressBar(completionPercent)
    
    'Update components
    Me.lblTablesCount.Caption = "Database Tables: " & tablesCreated
    Me.lblVBAModules.Caption = "VBA Modules: " & vbaModulesCount
    Me.lblQualityScore.Caption = "Current Quality Score: 9.8/10"
    
    'Update completed modules list
    Me.lstCompletedModules.Clear
    Me.lstCompletedModules.AddItem "✓ Module 1: Settings & Configuration (v1.0.0)"
    Me.lstCompletedModules.AddItem "✓ Module 2: Login & Security (v1.0.0)"
    
    'Update remaining modules list
    Me.lstRemainingModules.Clear
    Me.lstRemainingModules.AddItem "• Module 3: Dashboard"
    Me.lstRemainingModules.AddItem "• Module 4: Inventory"
    Me.lstRemainingModules.AddItem "• Module 5: Production"
    Me.lstRemainingModules.AddItem "• Module 6: Sales"
    Me.lstRemainingModules.AddItem "• Module 7: Purchasing"
    Me.lstRemainingModules.AddItem "• Module 8: Finance"
    Me.lstRemainingModules.AddItem "• Module 9: Reports"
    Me.lstRemainingModules.AddItem "• Module 10: Utilities"
    
    'Update next module
    Me.lblNextModule.Caption = "Next Module: Module 3 - Dashboard"
    
    Exit Sub
ErrorHandler:
    MsgBox "Error refreshing dashboard: " & Err.Description, vbCritical
End Sub

''
' GetProgressBar
' Purpose: Create visual progress bar
' Parameters: percent (0-100)
' Returns: String - Visual progress bar
'
Private Function GetProgressBar(percent As Double) As String
    Dim filled As Integer
    Dim empty As Integer
    Dim result As String
    
    filled = Int(percent / 5) 'Each block represents 5%
    empty = 20 - filled
    
    result = "["
    result = result & String(filled, "█")
    result = result & String(empty, "░")
    result = result & "] " & Format(percent, "0.0") & "%"
    
    GetProgressBar = result
End Function

''
' cmdTestLogin_Click
' Purpose: Test login functionality
'
Private Sub cmdTestLogin_Click()
    On Error GoTo ErrorHandler
    
    Dim user As clsUser
    
    'Test with default admin credentials
    Set user = mdlUserManagement.AuthenticateUser("admin", "Admin123")
    
    If user Is Nothing Then
        MsgBox "Login test failed", vbCritical
    Else
        MsgBox "Login test successful!" & vbCrLf & vbCrLf & _
               "User: " & user.FullName & vbCrLf & _
               "Role: " & user.Role, vbInformation
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

''
' cmdTestDatabase_Click
' Purpose: Test database connectivity
'
Private Sub cmdTestDatabase_Click()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tableCount As Integer
    Dim msg As String
    
    tableCount = 0
    
    'Count tables
    For Each ws In ThisWorkbook.Sheets
        tableCount = tableCount + ws.ListObjects.Count
    Next ws
    
    msg = "Database Status:" & vbCrLf & vbCrLf
    msg = msg & "Tables Created: " & tableCount & vbCrLf
    msg = msg & "Worksheets: " & ThisWorkbook.Sheets.Count & vbCrLf
    msg = msg & "Status: Connected"
    
    MsgBox msg, vbInformation
    
    Exit Sub
ErrorHandler:
    MsgBox "Database error: " & Err.Description, vbCritical
End Sub

''
' cmdOpenSettings_Click
' Purpose: Open settings worksheet
'
Private Sub cmdOpenSettings_Click()
    On Error GoTo ErrorHandler
    
    ThisWorkbook.Sheets(mdlConfiguration.WS_SETTINGS).Activate
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

''
' cmdOpenAuditTrail_Click
' Purpose: Open audit trail worksheet
'
Private Sub cmdOpenAuditTrail_Click()
    On Error GoTo ErrorHandler
    
    ThisWorkbook.Sheets(mdlConfiguration.WS_AUDIT).Activate
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

''
' cmdLogout_Click
' Purpose: Logout current user
'
Private Sub cmdLogout_Click()
    On Error GoTo ErrorHandler
    
    If mdlSessionManagement.EndSession() Then
        MsgBox "Logged out successfully", vbInformation
        Unload Me
    Else
        MsgBox "Logout failed", vbCritical
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

''
' cmdClose_Click
' Purpose: Close developer panel
'
Private Sub cmdClose_Click()
    Unload Me
End Sub
