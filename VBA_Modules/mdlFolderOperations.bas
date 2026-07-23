'================================================================================
'  mdlFolderOperations
'  Purpose: File and folder operations using FileSystemObject (FIX 1)
'  Author: Chokcity Technologies
'  Version: 1.0.0
'  Date: 2026-07-23
'================================================================================

Option Explicit

''
' CreateFolderIfNotExists
' Purpose: Create a folder if it doesn't already exist using FileSystemObject
' Parameters:
'   - folderPath: Full path to folder (e.g., "C:\Backups\")
' Returns: Boolean - True if folder exists or was created; False if error
' Error Handling: Logs failures to audit trail with descriptive errors
'
Public Function CreateFolderIfNotExists(folderPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim errorMessage As String
    
    'Validate input
    If Len(Trim(folderPath)) = 0 Then
        errorMessage = "Folder path cannot be empty"
        GoTo ErrorHandler
    End If
    
    'Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Check if folder already exists
    If fso.FolderExists(folderPath) Then
        CreateFolderIfNotExists = True
        Set fso = Nothing
        Exit Function
    End If
    
    'Attempt to create folder
    On Error Resume Next
    fso.CreateFolder folderPath
    On Error GoTo ErrorHandler
    
    'Verify creation succeeded
    If Not fso.FolderExists(folderPath) Then
        errorMessage = "Failed to create folder: " & folderPath
        GoTo ErrorHandler
    End If
    
    CreateFolderIfNotExists = True
    Set fso = Nothing
    Exit Function
    
ErrorHandler:
    If Len(errorMessage) = 0 Then
        errorMessage = Err.Description
    End If
    
    mdlConfiguration.LogAuditTrail "System", mdlConfiguration.AUDIT_CREATE, _
                                    "FolderOperations", 0, folderPath, "", "Failed", errorMessage
    
    CreateFolderIfNotExists = False
    If Not (fso Is Nothing) Then Set fso = Nothing
    
End Function

''
' FileExists
' Purpose: Check if a file exists
' Parameters:
'   - filePath: Full path to file
' Returns: Boolean - True if file exists
' Error Handling: Returns False on error
'
Public Function FileExists(filePath As String) As Boolean
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    FileExists = fso.FileExists(filePath)
    
    Set fso = Nothing
    On Error GoTo 0
End Function

''
' DeleteFileIfExists
' Purpose: Delete a file if it exists
' Parameters:
'   - filePath: Full path to file
' Returns: Boolean - True if deleted or didn't exist; False if error
' Error Handling: Logs deletion errors to audit trail
'
Public Function DeleteFileIfExists(filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(filePath) Then
        On Error Resume Next
        fso.DeleteFile filePath, True
        On Error GoTo ErrorHandler
    End If
    
    DeleteFileIfExists = True
    Set fso = Nothing
    Exit Function
    
ErrorHandler:
    mdlConfiguration.LogAuditTrail "System", mdlConfiguration.AUDIT_DELETE, _
                                    "FolderOperations", 0, filePath, "", "Failed", Err.Description
    DeleteFileIfExists = False
    If Not (fso Is Nothing) Then Set fso = Nothing
End Function

''
' GetFileSize
' Purpose: Get file size in bytes
' Parameters:
'   - filePath: Full path to file
' Returns: Long - File size in bytes; -1 if error or doesn't exist
'
Public Function GetFileSize(filePath As String) As Long
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(filePath) Then
        GetFileSize = fso.GetFile(filePath).Size
    Else
        GetFileSize = -1
    End If
    
    Set fso = Nothing
    On Error GoTo 0
End Function

''
' CleanOldBackups
' Purpose: Delete backup files older than specified days
' Parameters:
'   - backupFolder: Path to backups folder
'   - daysToKeep: Number of days to keep backups
'   - filePattern: File pattern (e.g., "Breadora_*.xlsm")
' Returns: Boolean - True if successful
' Error Handling: Logs cleanup results to audit trail
'
Public Function CleanOldBackups(backupFolder As String, daysToKeep As Integer, _
                                 filePattern As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim cutoffDate As Date
    Dim deletedCount As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(backupFolder) Then
        CleanOldBackups = False
        Set fso = Nothing
        Exit Function
    End If
    
    Set folder = fso.GetFolder(backupFolder)
    cutoffDate = DateAdd("d", -daysToKeep, Now())
    deletedCount = 0
    
    'Iterate through files in backup folder
    For Each file In folder.Files
        'Match pattern if specified
        If Len(filePattern) > 0 Then
            If Not file.Name Like filePattern Then
                GoTo NextFile
            End If
        End If
        
        'Check if file is older than cutoff date
        If file.DateLastModified < cutoffDate Then
            On Error Resume Next
            file.Delete True
            On Error GoTo ErrorHandler
            deletedCount = deletedCount + 1
        End If
        
NextFile:
    Next file
    
    mdlConfiguration.LogAuditTrail "System", "CleanupBackup", "FolderOperations", 0, _
                                    deletedCount & " old backups deleted", "", "Success"
    
    CleanOldBackups = True
    Set fso = Nothing
    Exit Function
    
ErrorHandler:
    mdlConfiguration.LogAuditTrail "System", "CleanupBackup", "FolderOperations", 0, _
                                    "", "", "Failed", Err.Description
    CleanOldBackups = False
    If Not (fso Is Nothing) Then Set fso = Nothing
End Function
