'================================================================================
'  clsLogger
'  Purpose: Centralized audit logging wrapper for all system activities
'  Author: Chokcity Technologies
'  Version: 1.0.0
'  Date: 2026-07-23
'================================================================================

Option Explicit

''
' LogCreate
' Purpose: Log a record creation
' Parameters: userID, tableName, recordID, newValue, description
' Returns: Boolean - True if logged successfully
'
Public Function LogCreate(userID As Variant, tableName As String, recordID As Variant, _
                          newValue As Variant, Optional description As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    mdlConfiguration.LogAuditTrail userID, mdlConfiguration.AUDIT_CREATE, tableName, _
                                    recordID, description, "", "Success"
    
    LogCreate = True
    Exit Function
ErrorHandler:
    LogCreate = False
End Function

''
' LogUpdate
' Purpose: Log a record update
' Parameters: userID, tableName, recordID, oldValue, newValue, description
' Returns: Boolean - True if logged successfully
'
Public Function LogUpdate(userID As Variant, tableName As String, recordID As Variant, _
                          oldValue As Variant, newValue As Variant, Optional description As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    mdlConfiguration.LogAuditTrail userID, mdlConfiguration.AUDIT_UPDATE, tableName, _
                                    recordID, description, oldValue, "Success"
    
    LogUpdate = True
    Exit Function
ErrorHandler:
    LogUpdate = False
End Function

''
' LogDelete
' Purpose: Log a record deletion
' Parameters: userID, tableName, recordID, oldValue, description
' Returns: Boolean - True if logged successfully
'
Public Function LogDelete(userID As Variant, tableName As String, recordID As Variant, _
                          oldValue As Variant, Optional description As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    mdlConfiguration.LogAuditTrail userID, mdlConfiguration.AUDIT_DELETE, tableName, _
                                    recordID, description, oldValue, "Success"
    
    LogDelete = True
    Exit Function
ErrorHandler:
    LogDelete = False
End Function

''
' LogLogin
' Purpose: Log user login
' Parameters: userID, description
' Returns: Boolean - True if logged successfully
'
Public Function LogLogin(userID As Variant, Optional description As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    mdlConfiguration.LogAuditTrail userID, mdlConfiguration.AUDIT_LOGIN, "", 0, description, "", "Success"
    
    LogLogin = True
    Exit Function
ErrorHandler:
    LogLogin = False
End Function

''
' LogLogout
' Purpose: Log user logout
' Parameters: userID, description
' Returns: Boolean - True if logged successfully
'
Public Function LogLogout(userID As Variant, Optional description As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    mdlConfiguration.LogAuditTrail userID, mdlConfiguration.AUDIT_LOGOUT, "", 0, description, "", "Success"
    
    LogLogout = True
    Exit Function
ErrorHandler:
    LogLogout = False
End Function

''
' LogExport
' Purpose: Log data export
' Parameters: userID, exportType, fileName
' Returns: Boolean - True if logged successfully
'
Public Function LogExport(userID As Variant, exportType As String, fileName As String) As Boolean
    On Error GoTo ErrorHandler
    
    mdlConfiguration.LogAuditTrail userID, mdlConfiguration.AUDIT_EXPORT, "", 0, _
                                    "Exported: " & exportType & " to " & fileName, "", "Success"
    
    LogExport = True
    Exit Function
ErrorHandler:
    LogExport = False
End Function

''
' LogError
' Purpose: Log an error event
' Parameters: userID, action, errorMsg
' Returns: Boolean - True if logged successfully
'
Public Function LogError(userID As Variant, action As String, errorMsg As String) As Boolean
    On Error GoTo ErrorHandler
    
    mdlConfiguration.LogAuditTrail userID, action, "", 0, "", "", "Failed", errorMsg
    
    LogError = True
    Exit Function
ErrorHandler:
    LogError = False
End Function
