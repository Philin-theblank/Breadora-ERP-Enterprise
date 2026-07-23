'================================================================================
'  mdlConfiguration
'  Purpose: Central configuration and settings management for Breadora ERP
'  Author: Chokcity Technologies
'  Version: 1.0.1 (Updated with CTO fixes)
'  Date: 2026-07-23
'================================================================================

Option Explicit

'================================================================================
' APPLICATION CONSTANTS
'================================================================================

Public Const APP_NAME As String = "Breadora ERP Enterprise"
Public Const APP_VERSION As String = "1.0.0"
Public Const APP_BUILD As String = "2026-07-23"
Public Const APP_DEVELOPER As String = "Chokcity Technologies"

'================================================================================
' DATABASE TABLE NAMES
'================================================================================

Public Const TBL_SETTINGS As String = "tblSettings"
Public Const TBL_COMPANY_INFO As String = "tblCompanyInfo"
Public Const TBL_NUMBER_SEQUENCES As String = "tblNumberSequences"
Public Const TBL_AUDIT_TRAIL As String = "tblAuditTrail"
Public Const TBL_USERS As String = "tblUsers"
Public Const TBL_PRODUCTS As String = "tblProducts"
Public Const TBL_RECIPES As String = "tblRecipes"
Public Const TBL_RECIPE_ITEMS As String = "tblRecipeItems"
Public Const TBL_CUSTOMERS As String = "tblCustomers"
Public Const TBL_SUPPLIERS As String = "tblSuppliers"
Public Const TBL_INVENTORY As String = "tblInventory"
Public Const TBL_INVENTORY_TRANSACTIONS As String = "tblInventoryTransactions"
Public Const TBL_PRODUCTION As String = "tblProduction"
Public Const TBL_PRODUCTION_ITEMS As String = "tblProductionItems"
Public Const TBL_PURCHASES As String = "tblPurchases"
Public Const TBL_PURCHASE_ITEMS As String = "tblPurchaseItems"
Public Const TBL_SALES As String = "tblSales"
Public Const TBL_SALE_ITEMS As String = "tblSaleItems"
Public Const TBL_CASHBOOK As String = "tblCashbook"
Public Const TBL_EXPENSES As String = "tblExpenses"

'================================================================================
' WORKSHEET NAMES (Database Layer)
'================================================================================

Public Const WS_SETTINGS As String = "Settings"
Public Const WS_COMPANY As String = "Company"
Public Const WS_SEQUENCES As String = "Sequences"
Public Const WS_AUDIT As String = "AuditTrail"
Public Const WS_USERS As String = "Users"
Public Const WS_PRODUCTS As String = "Products"
Public Const WS_RECIPES As String = "Recipes"
Public Const WS_RECIPE_ITEMS As String = "RecipeItems"
Public Const WS_CUSTOMERS As String = "Customers"
Public Const WS_SUPPLIERS As String = "Suppliers"
Public Const WS_INVENTORY As String = "Inventory"
Public Const WS_INV_TRANSACTIONS As String = "InventoryTransactions"
Public Const WS_PRODUCTION As String = "Production"
Public Const WS_PRODUCTION_ITEMS As String = "ProductionItems"
Public Const WS_PURCHASES As String = "Purchases"
Public Const WS_PURCHASE_ITEMS As String = "PurchaseItems"
Public Const WS_SALES As String = "Sales"
Public Const WS_SALE_ITEMS As String = "SaleItems"
Public Const WS_CASHBOOK As String = "Cashbook"
Public Const WS_EXPENSES As String = "Expenses"

'================================================================================
' DOCUMENT NUMBERING PREFIXES
'================================================================================

Public Const DOC_PREFIX_INVOICE As String = "INV"
Public Const DOC_PREFIX_PURCHASE_ORDER As String = "PO"
Public Const DOC_PREFIX_GRN As String = "GRN"
Public Const DOC_PREFIX_PRODUCTION As String = "PROD"
Public Const DOC_PREFIX_RECEIPT As String = "REC"
Public Const DOC_PREFIX_EXPENSE As String = "EXP"

'================================================================================
' USER ROLES
'================================================================================

Public Const ROLE_ADMIN As String = "Administrator"
Public Const ROLE_MANAGER As String = "Manager"
Public Const ROLE_PRODUCTION As String = "Production Manager"
Public Const ROLE_STOREKEEPER As String = "Storekeeper"
Public Const ROLE_CASHIER As String = "Cashier"
Public Const ROLE_ACCOUNTANT As String = "Accountant"
Public Const ROLE_SALES_REP As String = "Sales Representative"
Public Const ROLE_VIEWER As String = "Viewer"

'================================================================================
' UI COLORS
'================================================================================

Public Const COLOR_PRIMARY As Long = RGB(31, 78, 121)      'Dark Blue
Public Const COLOR_SECONDARY As Long = RGB(192, 0, 0)      'Dark Red
Public Const COLOR_SUCCESS As Long = RGB(0, 112, 64)        'Green
Public Const COLOR_WARNING As Long = RGB(255, 192, 0)       'Gold
Public Const COLOR_DANGER As Long = RGB(192, 0, 0)          'Red
Public Const COLOR_BACKGROUND As Long = RGB(255, 255, 255) 'White
Public Const COLOR_TEXT As Long = RGB(64, 64, 64)           'Dark Gray

'================================================================================
' STATUS VALUES
'================================================================================

Public Const STATUS_ACTIVE As String = "Active"
Public Const STATUS_INACTIVE As String = "Inactive"
Public Const STATUS_PENDING As String = "Pending"
Public Const STATUS_COMPLETED As String = "Completed"
Public Const STATUS_CANCELLED As String = "Cancelled"
Public Const STATUS_DRAFT As String = "Draft"

'================================================================================
' AUDIT ACTION TYPES
'================================================================================

Public Const AUDIT_CREATE As String = "Create"
Public Const AUDIT_UPDATE As String = "Update"
Public Const AUDIT_DELETE As String = "Delete"
Public Const AUDIT_LOGIN As String = "Login"
Public Const AUDIT_LOGOUT As String = "Logout"
Public Const AUDIT_EXPORT As String = "Export"
Public Const AUDIT_IMPORT As String = "Import"
Public Const AUDIT_BACKUP As String = "Backup"
Public Const AUDIT_RESTORE As String = "Restore"

'================================================================================
' INITIALIZATION PROCEDURES
'================================================================================

''
' InitializeApplication
' Purpose: Initialize entire application on first run
' Called from: Workbook_Open event
' Returns: Boolean - True if successful
' Error Handling: Logs all failures to audit trail
'
Public Function InitializeApplication() As Boolean
    On Error GoTo ErrorHandler
    
    'Step 1: Verify or create database structure
    If Not DatabaseExists() Then
        If Not CreateDatabase() Then
            MsgBox "Critical: Database creation failed. Please restore from backup.", vbCritical
            InitializeApplication = False
            Exit Function
        End If
    End If
    
    'Step 2: Validate all tables exist
    If Not ValidateAllTables() Then
        MsgBox "Critical: Database validation failed. Please restore from backup.", vbCritical
        InitializeApplication = False
        Exit Function
    End If
    
    'Step 3: Initialize company info if empty
    If IsCompanyInfoEmpty() Then
        If Not InitializeCompanyInfo() Then
            LogAuditTrail "System", AUDIT_CREATE, TBL_COMPANY_INFO, 0, "", "", "Failed", _
                          "Failed to initialize company info"
        End If
    End If
    
    'Step 4: Initialize number sequences if empty
    If IsNumberSequencesEmpty() Then
        If Not InitializeNumberSequences() Then
            LogAuditTrail "System", AUDIT_CREATE, TBL_NUMBER_SEQUENCES, 0, "", "", "Failed", _
                          "Failed to initialize number sequences"
        End If
    End If
    
    'Step 5: Initialize core settings if empty
    If IsCoreSettingsEmpty() Then
        If Not InitializeCoreSettings() Then
            LogAuditTrail "System", AUDIT_CREATE, TBL_SETTINGS, 0, "", "", "Failed", _
                          "Failed to initialize core settings"
        End If
    End If
    
    'Step 6: Check if backup needed
    CheckAndPerformAutoBackup
    
    InitializeApplication = True
    Exit Function
    
ErrorHandler:
    MsgBox "Initialization Error: " & Err.Description, vbCritical
    LogAuditTrail "System", AUDIT_CREATE, "InitializeApplication", 0, "", "", "Failed", Err.Description
    InitializeApplication = False
End Function

''
' DatabaseExists
' Purpose: Check if database structure exists
' Returns: Boolean - True if database is initialized
'
Private Function DatabaseExists() As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(WS_SETTINGS)
    DatabaseExists = (ws Is Nothing) = False
    On Error GoTo 0
End Function

''
' CreateDatabase
' Purpose: Create all required Excel Tables for database
' Returns: Boolean - True if successful; False if any table creation fails
' Error Handling: Validates each table creation
'
Private Function CreateDatabase() As Boolean
    On Error GoTo ErrorHandler
    
    Dim allSuccess As Boolean
    allSuccess = True
    
    'Create all tables with error checking
    If Not CreateSettingsTable() Then allSuccess = False
    If Not CreateCompanyInfoTable() Then allSuccess = False
    If Not CreateNumberSequencesTable() Then allSuccess = False
    If Not CreateAuditTrailTable() Then allSuccess = False
    
    CreateDatabase = allSuccess
    Exit Function
    
ErrorHandler:
    LogAuditTrail "System", AUDIT_CREATE, "CreateDatabase", 0, "", "", "Failed", Err.Description
    CreateDatabase = False
End Function

''
' CreateSettingsTable
' Purpose: Create tblSettings Excel Table
' Returns: Boolean - True if successful
' Error Handling: Validates table creation and structure
'
Private Function CreateSettingsTable() As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim tbl As ListObject
    
    'Get or create worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_SETTINGS)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = WS_SETTINGS
    End If
    
    'Clear existing data
    ws.Cells.Clear
    
    'Create headers
    ws.Range("A1").Value = "SettingID"
    ws.Range("B1").Value = "SettingKey"
    ws.Range("C1").Value = "SettingValue"
    ws.Range("D1").Value = "SettingDataType"
    ws.Range("E1").Value = "CreatedDate"
    ws.Range("F1").Value = "UpdatedDate"
    ws.Range("G1").Value = "CreatedBy"
    ws.Range("H1").Value = "Status"
    
    'Format headers
    ws.Range("A1:H1").Font.Bold = True
    ws.Range("A1:H1").Interior.Color = COLOR_PRIMARY
    ws.Range("A1:H1").Font.Color = RGB(255, 255, 255)
    
    'Create Excel Table
    Set rng = ws.Range("A1:H1")
    Set tbl = ThisWorkbook.Sheets(WS_SETTINGS).ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = TBL_SETTINGS
    
    'Validate table was created
    If Not TableExists(TBL_SETTINGS) Then
        LogAuditTrail "System", AUDIT_CREATE, TBL_SETTINGS, 0, "", "", "Failed", _
                      "Table validation failed after creation"
        CreateSettingsTable = False
        Exit Function
    End If
    
    CreateSettingsTable = True
    Exit Function
    
ErrorHandler:
    LogAuditTrail "System", AUDIT_CREATE, TBL_SETTINGS, 0, "", "", "Failed", Err.Description
    CreateSettingsTable = False
End Function

''
' CreateCompanyInfoTable
' Purpose: Create tblCompanyInfo Excel Table
' Returns: Boolean - True if successful
' Error Handling: Validates table creation and structure
'
Private Function CreateCompanyInfoTable() As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim tbl As ListObject
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_COMPANY)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = WS_COMPANY
    End If
    
    ws.Cells.Clear
    
    'Create headers
    ws.Range("A1").Value = "CompanyID"
    ws.Range("B1").Value = "CompanyName"
    ws.Range("C1").Value = "CompanyRegistration"
    ws.Range("D1").Value = "Address"
    ws.Range("E1").Value = "City"
    ws.Range("F1").Value = "PostalCode"
    ws.Range("G1").Value = "Phone"
    ws.Range("H1").Value = "Email"
    ws.Range("I1").Value = "Website"
    ws.Range("J1").Value = "Currency"
    ws.Range("K1").Value = "TaxRate"
    ws.Range("L1").Value = "FinancialYearStart"
    ws.Range("M1").Value = "FinancialYearEnd"
    ws.Range("N1").Value = "LogoPath"
    ws.Range("O1").Value = "CreatedDate"
    ws.Range("P1").Value = "UpdatedDate"
    ws.Range("Q1").Value = "CreatedBy"
    
    'Format headers
    ws.Range("A1:Q1").Font.Bold = True
    ws.Range("A1:Q1").Interior.Color = COLOR_PRIMARY
    ws.Range("A1:Q1").Font.Color = RGB(255, 255, 255)
    
    'Create Excel Table
    Set rng = ws.Range("A1:Q1")
    Set tbl = ThisWorkbook.Sheets(WS_COMPANY).ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = TBL_COMPANY_INFO
    
    'Validate table was created
    If Not TableExists(TBL_COMPANY_INFO) Then
        LogAuditTrail "System", AUDIT_CREATE, TBL_COMPANY_INFO, 0, "", "", "Failed", _
                      "Table validation failed after creation"
        CreateCompanyInfoTable = False
        Exit Function
    End If
    
    CreateCompanyInfoTable = True
    Exit Function
    
ErrorHandler:
    LogAuditTrail "System", AUDIT_CREATE, TBL_COMPANY_INFO, 0, "", "", "Failed", Err.Description
    CreateCompanyInfoTable = False
End Function

''
' CreateNumberSequencesTable
' Purpose: Create tblNumberSequences Excel Table
' Returns: Boolean - True if successful
' Error Handling: Validates table creation and structure
'
Private Function CreateNumberSequencesTable() As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim tbl As ListObject
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_SEQUENCES)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = WS_SEQUENCES
    End If
    
    ws.Cells.Clear
    
    'Create headers
    ws.Range("A1").Value = "SequenceID"
    ws.Range("B1").Value = "DocumentType"
    ws.Range("C1").Value = "CurrentSequence"
    ws.Range("D1").Value = "Prefix"
    ws.Range("E1").Value = "Suffix"
    ws.Range("F1").Value = "YearFormat"
    ws.Range("G1").Value = "PadLength"
    ws.Range("H1").Value = "ResetFrequency"
    ws.Range("I1").Value = "LastResetDate"
    ws.Range("J1").Value = "CreatedDate"
    ws.Range("K1").Value = "UpdatedDate"
    ws.Range("L1").Value = "CreatedBy"
    ws.Range("M1").Value = "Status"
    
    'Format headers
    ws.Range("A1:M1").Font.Bold = True
    ws.Range("A1:M1").Interior.Color = COLOR_PRIMARY
    ws.Range("A1:M1").Font.Color = RGB(255, 255, 255)
    
    'Create Excel Table
    Set rng = ws.Range("A1:M1")
    Set tbl = ThisWorkbook.Sheets(WS_SEQUENCES).ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = TBL_NUMBER_SEQUENCES
    
    'Validate table was created
    If Not TableExists(TBL_NUMBER_SEQUENCES) Then
        LogAuditTrail "System", AUDIT_CREATE, TBL_NUMBER_SEQUENCES, 0, "", "", "Failed", _
                      "Table validation failed after creation"
        CreateNumberSequencesTable = False
        Exit Function
    End If
    
    CreateNumberSequencesTable = True
    Exit Function
    
ErrorHandler:
    LogAuditTrail "System", AUDIT_CREATE, TBL_NUMBER_SEQUENCES, 0, "", "", "Failed", Err.Description
    CreateNumberSequencesTable = False
End Function

''
' CreateAuditTrailTable
' Purpose: Create tblAuditTrail Excel Table
' Returns: Boolean - True if successful
' Error Handling: Validates table creation and structure
'
Private Function CreateAuditTrailTable() As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim tbl As ListObject
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_AUDIT)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = WS_AUDIT
    End If
    
    ws.Cells.Clear
    
    'Create headers
    ws.Range("A1").Value = "AuditID"
    ws.Range("B1").Value = "UserID"
    ws.Range("C1").Value = "Action"
    ws.Range("D1").Value = "Module"
    ws.Range("E1").Value = "RecordID"
    ws.Range("F1").Value = "Description"
    ws.Range("G1").Value = "Timestamp"
    ws.Range("H1").Value = "Status"
    ws.Range("I1").Value = "ErrorMessage"
    
    'Format headers
    ws.Range("A1:I1").Font.Bold = True
    ws.Range("A1:I1").Interior.Color = COLOR_PRIMARY
    ws.Range("A1:I1").Font.Color = RGB(255, 255, 255)
    
    'Create Excel Table
    Set rng = ws.Range("A1:I1")
    Set tbl = ThisWorkbook.Sheets(WS_AUDIT).ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = TBL_AUDIT_TRAIL
    
    'Validate table was created
    If Not TableExists(TBL_AUDIT_TRAIL) Then
        LogAuditTrail "System", AUDIT_CREATE, TBL_AUDIT_TRAIL, 0, "", "", "Failed", _
                      "Table validation failed after creation"
        CreateAuditTrailTable = False
        Exit Function
    End If
    
    CreateAuditTrailTable = True
    Exit Function
    
ErrorHandler:
    LogAuditTrail "System", AUDIT_CREATE, TBL_AUDIT_TRAIL, 0, "", "", "Failed", Err.Description
    CreateAuditTrailTable = False
End Function

'================================================================================
' VALIDATION PROCEDURES
'================================================================================

''
' ValidateAllTables
' Purpose: Verify all required Excel Tables exist and have correct structure
' Returns: Boolean - True if all tables valid
'
Private Function ValidateAllTables() As Boolean
    On Error GoTo ErrorHandler
    
    Dim requiredTables As Variant
    Dim i As Integer
    
    requiredTables = Array(TBL_SETTINGS, TBL_COMPANY_INFO, TBL_NUMBER_SEQUENCES, TBL_AUDIT_TRAIL)
    
    For i = LBound(requiredTables) To UBound(requiredTables)
        If Not TableExists(CStr(requiredTables(i))) Then
            LogAuditTrail "System", "Validate", "ValidateAllTables", 0, CStr(requiredTables(i)), "", "Failed", _
                          "Required table missing: " & CStr(requiredTables(i))
            ValidateAllTables = False
            Exit Function
        End If
    Next i
    
    ValidateAllTables = True
    Exit Function
    
ErrorHandler:
    LogAuditTrail "System", "Validate", "ValidateAllTables", 0, "", "", "Failed", Err.Description
    ValidateAllTables = False
End Function

''
' TableExists
' Purpose: Check if an Excel Table exists by name
' Parameters: tableName - Name of table to check
' Returns: Boolean - True if table exists and is valid
'
Public Function TableExists(tableName As String) As Boolean
    On Error Resume Next
    Dim tbl As ListObject
    Dim ws As Worksheet
    Dim i As Integer
    
    'Search all worksheets for table
    For Each ws In ThisWorkbook.Sheets
        For Each tbl In ws.ListObjects
            If tbl.Name = tableName Then
                TableExists = True
                Exit Function
            End If
        Next tbl
    Next ws
    
    TableExists = False
    On Error GoTo 0
End Function

''
' IsCompanyInfoEmpty
' Purpose: Check if company info has been initialized
' Returns: Boolean - True if empty
'
Private Function IsCompanyInfoEmpty() As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Sheets(WS_COMPANY)
    Set tbl = ws.ListObjects(TBL_COMPANY_INFO)
    
    IsCompanyInfoEmpty = (tbl.DataBodyRange.Rows.Count = 0)
    On Error GoTo 0
End Function

''
' IsNumberSequencesEmpty
' Purpose: Check if number sequences have been initialized
' Returns: Boolean - True if empty
'
Private Function IsNumberSequencesEmpty() As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Sheets(WS_SEQUENCES)
    Set tbl = ws.ListObjects(TBL_NUMBER_SEQUENCES)
    
    IsNumberSequencesEmpty = (tbl.DataBodyRange.Rows.Count = 0)
    On Error GoTo 0
End Function

''
' IsCoreSettingsEmpty
' Purpose: Check if core settings have been initialized
' Returns: Boolean - True if empty
'
Private Function IsCoreSettingsEmpty() As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Sheets(WS_SETTINGS)
    Set tbl = ws.ListObjects(TBL_SETTINGS)
    
    IsCoreSettingsEmpty = (tbl.DataBodyRange.Rows.Count = 0)
    On Error GoTo 0
End Function

'================================================================================
' INITIALIZATION DATA PROCEDURES
'================================================================================

''
' InitializeCompanyInfo
' Purpose: Add default company info record
' Returns: Boolean - True if successful
'
Private Function InitializeCompanyInfo() As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    
    Set ws = ThisWorkbook.Sheets(WS_COMPANY)
    Set tbl = ws.ListObjects(TBL_COMPANY_INFO)
    
    Set newRow = tbl.ListRows.Add
    
    With newRow.Range
        .Cells(1, 1).Value = 1 'CompanyID
        .Cells(1, 2).Value = "Chokcity Bakery" 'CompanyName
        .Cells(1, 3).Value = "" 'CompanyRegistration
        .Cells(1, 4).Value = "" 'Address
        .Cells(1, 5).Value = "" 'City
        .Cells(1, 6).Value = "" 'PostalCode
        .Cells(1, 7).Value = "" 'Phone
        .Cells(1, 8).Value = "" 'Email
        .Cells(1, 9).Value = "" 'Website
        .Cells(1, 10).Value = "USD" 'Currency
        .Cells(1, 11).Value = 0.15 'TaxRate
        .Cells(1, 12).Value = DateSerial(Year(Now()), 1, 1) 'FinancialYearStart
        .Cells(1, 13).Value = DateSerial(Year(Now()), 12, 31) 'FinancialYearEnd
        .Cells(1, 14).Value = "" 'LogoPath
        .Cells(1, 15).Value = Now() 'CreatedDate
        .Cells(1, 16).Value = Now() 'UpdatedDate
        .Cells(1, 17).Value = "System" 'CreatedBy
    End With
    
    InitializeCompanyInfo = True
    Exit Function
    
ErrorHandler:
    LogAuditTrail "System", AUDIT_CREATE, TBL_COMPANY_INFO, 0, "", "", "Failed", Err.Description
    InitializeCompanyInfo = False
End Function

''
' InitializeNumberSequences
' Purpose: Add default number sequence records for all document types
' Returns: Boolean - True if successful
'
Private Function InitializeNumberSequences() As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim sequenceID As Integer
    
    Set ws = ThisWorkbook.Sheets(WS_SEQUENCES)
    Set tbl = ws.ListObjects(TBL_NUMBER_SEQUENCES)
    
    sequenceID = 1
    
    'Invoice Sequence
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = sequenceID
        .Cells(1, 2).Value = "Invoice"
        .Cells(1, 3).Value = 1
        .Cells(1, 4).Value = DOC_PREFIX_INVOICE
        .Cells(1, 5).Value = ""
        .Cells(1, 6).Value = "YYYY"
        .Cells(1, 7).Value = 6
        .Cells(1, 8).Value = "Annual"
        .Cells(1, 9).Value = DateSerial(Year(Now()), 1, 1)
        .Cells(1, 10).Value = Now()
        .Cells(1, 11).Value = Now()
        .Cells(1, 12).Value = "System"
        .Cells(1, 13).Value = STATUS_ACTIVE
    End With
    sequenceID = sequenceID + 1
    
    'Purchase Order Sequence
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = sequenceID
        .Cells(1, 2).Value = "Purchase Order"
        .Cells(1, 3).Value = 1
        .Cells(1, 4).Value = DOC_PREFIX_PURCHASE_ORDER
        .Cells(1, 5).Value = ""
        .Cells(1, 6).Value = "YYYY"
        .Cells(1, 7).Value = 6
        .Cells(1, 8).Value = "Annual"
        .Cells(1, 9).Value = DateSerial(Year(Now()), 1, 1)
        .Cells(1, 10).Value = Now()
        .Cells(1, 11).Value = Now()
        .Cells(1, 12).Value = "System"
        .Cells(1, 13).Value = STATUS_ACTIVE
    End With
    sequenceID = sequenceID + 1
    
    'GRN Sequence
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = sequenceID
        .Cells(1, 2).Value = "Goods Received Note"
        .Cells(1, 3).Value = 1
        .Cells(1, 4).Value = DOC_PREFIX_GRN
        .Cells(1, 5).Value = ""
        .Cells(1, 6).Value = "YYYY"
        .Cells(1, 7).Value = 6
        .Cells(1, 8).Value = "Annual"
        .Cells(1, 9).Value = DateSerial(Year(Now()), 1, 1)
        .Cells(1, 10).Value = Now()
        .Cells(1, 11).Value = Now()
        .Cells(1, 12).Value = "System"
        .Cells(1, 13).Value = STATUS_ACTIVE
    End With
    sequenceID = sequenceID + 1
    
    'Production Order Sequence
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = sequenceID
        .Cells(1, 2).Value = "Production Order"
        .Cells(1, 3).Value = 1
        .Cells(1, 4).Value = DOC_PREFIX_PRODUCTION
        .Cells(1, 5).Value = ""
        .Cells(1, 6).Value = "YYYY"
        .Cells(1, 7).Value = 6
        .Cells(1, 8).Value = "Annual"
        .Cells(1, 9).Value = DateSerial(Year(Now()), 1, 1)
        .Cells(1, 10).Value = Now()
        .Cells(1, 11).Value = Now()
        .Cells(1, 12).Value = "System"
        .Cells(1, 13).Value = STATUS_ACTIVE
    End With
    sequenceID = sequenceID + 1
    
    'Receipt Sequence
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = sequenceID
        .Cells(1, 2).Value = "Receipt"
        .Cells(1, 3).Value = 1
        .Cells(1, 4).Value = DOC_PREFIX_RECEIPT
        .Cells(1, 5).Value = ""
        .Cells(1, 6).Value = "YYYY"
        .Cells(1, 7).Value = 6
        .Cells(1, 8).Value = "Annual"
        .Cells(1, 9).Value = DateSerial(Year(Now()), 1, 1)
        .Cells(1, 10).Value = Now()
        .Cells(1, 11).Value = Now()
        .Cells(1, 12).Value = "System"
        .Cells(1, 13).Value = STATUS_ACTIVE
    End With
    sequenceID = sequenceID + 1
    
    'Expense Sequence
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = sequenceID
        .Cells(1, 2).Value = "Expense"
        .Cells(1, 3).Value = 1
        .Cells(1, 4).Value = DOC_PREFIX_EXPENSE
        .Cells(1, 5).Value = ""
        .Cells(1, 6).Value = "YYYY"
        .Cells(1, 7).Value = 6
        .Cells(1, 8).Value = "Annual"
        .Cells(1, 9).Value = DateSerial(Year(Now()), 1, 1)
        .Cells(1, 10).Value = Now()
        .Cells(1, 11).Value = Now()
        .Cells(1, 12).Value = "System"
        .Cells(1, 13).Value = STATUS_ACTIVE
    End With
    
    InitializeNumberSequences = True
    Exit Function
    
ErrorHandler:
    LogAuditTrail "System", AUDIT_CREATE, TBL_NUMBER_SEQUENCES, 0, "", "", "Failed", Err.Description
    InitializeNumberSequences = False
End Function

''
' InitializeCoreSettings
' Purpose: Add core system settings
' Returns: Boolean - True if successful
'
Private Function InitializeCoreSettings() As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim backupPath As String
    
    Set ws = ThisWorkbook.Sheets(WS_SETTINGS)
    Set tbl = ws.ListObjects(TBL_SETTINGS)
    
    'Determine backup path
    backupPath = ThisWorkbook.Path & "\Backups\"
    
    'Setting 1: Last Backup Date
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = 1
        .Cells(1, 2).Value = "LastBackupDate"
        .Cells(1, 3).Value = Format(Now(), "yyyy-mm-dd hh:mm:ss")
        .Cells(1, 4).Value = "DateTime"
        .Cells(1, 5).Value = Now()
        .Cells(1, 6).Value = Now()
        .Cells(1, 7).Value = "System"
        .Cells(1, 8).Value = STATUS_ACTIVE
    End With
    
    'Setting 2: Backup Path
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = 2
        .Cells(1, 2).Value = "BackupPath"
        .Cells(1, 3).Value = backupPath
        .Cells(1, 4).Value = "String"
        .Cells(1, 5).Value = Now()
        .Cells(1, 6).Value = Now()
        .Cells(1, 7).Value = "System"
        .Cells(1, 8).Value = STATUS_ACTIVE
    End With
    
    'Setting 3: Auto Backup Enabled
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = 3
        .Cells(1, 2).Value = "AutoBackupEnabled"
        .Cells(1, 3).Value = "Yes"
        .Cells(1, 4).Value = "Boolean"
        .Cells(1, 5).Value = Now()
        .Cells(1, 6).Value = Now()
        .Cells(1, 7).Value = "System"
        .Cells(1, 8).Value = STATUS_ACTIVE
    End With
    
    'Setting 4: Auto Backup Frequency (in days)
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = 4
        .Cells(1, 2).Value = "AutoBackupFrequency"
        .Cells(1, 3).Value = "1"
        .Cells(1, 4).Value = "Integer"
        .Cells(1, 5).Value = Now()
        .Cells(1, 6).Value = Now()
        .Cells(1, 7).Value = "System"
        .Cells(1, 8).Value = STATUS_ACTIVE
    End With
    
    'Setting 5: Backup Retention Days
    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, 1).Value = 5
        .Cells(1, 2).Value = "BackupRetentionDays"
        .Cells(1, 3).Value = "30"
        .Cells(1, 4).Value = "Integer"
        .Cells(1, 5).Value = Now()
        .Cells(1, 6).Value = Now()
        .Cells(1, 7).Value = "System"
        .Cells(1, 8).Value = STATUS_ACTIVE
    End With
    
    'Create backup folder
    Call mdlFolderOperations.CreateFolderIfNotExists(backupPath)
    
    InitializeCoreSettings = True
    Exit Function
    
ErrorHandler:
    LogAuditTrail "System", AUDIT_CREATE, TBL_SETTINGS, 0, "", "", "Failed", Err.Description
    InitializeCoreSettings = False
End Function

'================================================================================
' UTILITY PROCEDURES
'================================================================================

''
' LogAuditTrail
' Purpose: Log an action to audit trail (STANDARDIZED SIGNATURE - FIX 3)
' Parameters (in order):
'   - userID: User performing action (or "System" for automated actions)
'   - action: Type of action (Create, Update, Delete, Login, etc.)
'   - module: Module/table affected (e.g., "tblProducts", "tblSales")
'   - recordID: Record ID affected (0 for non-record actions)
'   - description: Detailed description of action
'   - oldValue: Previous value (for auditing changes)
'   - status: "Success" or "Failed"
'   - errorMsg: (Optional) Error message if status is "Failed"
' Note: Call signature standardized for consistency across all modules
'
Public Sub LogAuditTrail(userID As Variant, action As String, module As String, _
                         recordID As Variant, description As Variant, _
                         oldValue As Variant, status As String, Optional errorMsg As String = "")
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    
    Set ws = ThisWorkbook.Sheets(WS_AUDIT)
    Set tbl = ws.ListObjects(TBL_AUDIT_TRAIL)
    
    Set newRow = tbl.ListRows.Add
    
    With newRow.Range
        .Cells(1, 1).Value = tbl.ListRows.Count 'AuditID
        .Cells(1, 2).Value = userID 'UserID
        .Cells(1, 3).Value = action 'Action
        .Cells(1, 4).Value = module 'Module
        .Cells(1, 5).Value = recordID 'RecordID
        .Cells(1, 6).Value = description 'Description
        .Cells(1, 7).Value = Now() 'Timestamp
        .Cells(1, 8).Value = status 'Status
        .Cells(1, 9).Value = errorMsg 'ErrorMessage
    End With
    
    Exit Sub
ErrorHandler:
    'Silently fail if audit trail logging fails to prevent cascading errors
End Sub

''
' CheckAndPerformAutoBackup
' Purpose: Check if auto-backup needed and perform if due
'
Private Sub CheckAndPerformAutoBackup()
    On Error GoTo ErrorHandler
    
    Dim lastBackupDate As Date
    Dim backupFrequency As Integer
    Dim backupDue As Boolean
    
    'Get last backup date from settings
    lastBackupDate = CDate(GetSettingValue("LastBackupDate", Format(Now(), "yyyy-mm-dd hh:mm:ss")))
    
    'Get backup frequency
    backupFrequency = CInt(GetSettingValue("AutoBackupFrequency", "1"))
    
    'Check if backup is due
    backupDue = (DateDiff("d", lastBackupDate, Now()) >= backupFrequency)
    
    If backupDue Then
        Call PerformBackup
    End If
    
    Exit Sub
ErrorHandler:
    'Silently fail if backup check fails
End Sub

''
' PerformBackup
' Purpose: Create backup of current workbook using FileSystemObject (FIX 1)
'
Public Sub PerformBackup()
    On Error GoTo ErrorHandler
    
    Dim backupPath As String
    Dim backupFilename As String
    Dim backupFullPath As String
    Dim backupRetention As Integer
    
    backupPath = GetSettingValue("BackupPath", ThisWorkbook.Path & "\Backups\")
    backupRetention = CInt(GetSettingValue("BackupRetentionDays", "30"))
    
    'Create backup directory if it doesn't exist
    If Not mdlFolderOperations.CreateFolderIfNotExists(backupPath) Then
        LogAuditTrail "System", AUDIT_BACKUP, "", 0, "", "", "Failed", "Could not create backup folder"
        Exit Sub
    End If
    
    'Create backup filename with timestamp
    backupFilename = "Breadora_" & Format(Now(), "yyyy-mm-dd_hhmm") & ".xlsm"
    backupFullPath = backupPath & backupFilename
    
    'Copy current workbook to backup location
    On Error Resume Next
    ThisWorkbook.SaveCopyAs backupFullPath
    On Error GoTo ErrorHandler
    
    'Verify backup was created
    If Not mdlFolderOperations.FileExists(backupFullPath) Then
        LogAuditTrail "System", AUDIT_BACKUP, "", 0, backupFullPath, "", "Failed", "Backup file not created"
        Exit Sub
    End If
    
    'Update last backup date in settings
    Call UpdateSettingValue("LastBackupDate", Format(Now(), "yyyy-mm-dd hh:mm:ss"))
    
    'Clean old backups
    Call mdlFolderOperations.CleanOldBackups(backupPath, backupRetention, "Breadora_*.xlsm")
    
    'Log to audit trail
    LogAuditTrail "System", AUDIT_BACKUP, "", 0, backupFullPath, "", "Success"
    
    Exit Sub
ErrorHandler:
    LogAuditTrail "System", AUDIT_BACKUP, "", 0, "", "", "Failed", Err.Description
End Sub

''
' GetSettingValue
' Purpose: Retrieve a setting value by key
' Parameters: settingKey - Name of setting; defaultValue - Fallback if not found
' Returns: Setting value or default if not found
' Error Handling: Returns default on any error
'
Public Function GetSettingValue(settingKey As String, Optional defaultValue As Variant = "") As Variant
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets(WS_SETTINGS)
    Set tbl = ws.ListObjects(TBL_SETTINGS)
    
    'Search for setting key
    For i = 1 To tbl.ListRows.Count
        If tbl.ListRows(i).Range.Cells(1, 2).Value = settingKey Then
            GetSettingValue = tbl.ListRows(i).Range.Cells(1, 3).Value
            Exit Function
        End If
    Next i
    
    'Not found, return default
    GetSettingValue = defaultValue
    Exit Function
    
ErrorHandler:
    GetSettingValue = defaultValue
End Function

''
' UpdateSettingValue
' Purpose: Update or create a setting value
' Parameters: settingKey - Setting name; settingValue - New value
' Error Handling: Logs failures
'
Public Sub UpdateSettingValue(settingKey As String, settingValue As Variant)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Integer
    Dim found As Boolean
    Dim newRow As ListRow
    
    Set ws = ThisWorkbook.Sheets(WS_SETTINGS)
    Set tbl = ws.ListObjects(TBL_SETTINGS)
    
    found = False
    
    'Try to find existing setting
    For i = 1 To tbl.ListRows.Count
        If tbl.ListRows(i).Range.Cells(1, 2).Value = settingKey Then
            tbl.ListRows(i).Range.Cells(1, 3).Value = settingValue
            tbl.ListRows(i).Range.Cells(1, 6).Value = Now()
            found = True
            Exit For
        End If
    Next i
    
    'If not found, create new setting
    If Not found Then
        Set newRow = tbl.ListRows.Add
        With newRow.Range
            .Cells(1, 1).Value = tbl.ListRows.Count
            .Cells(1, 2).Value = settingKey
            .Cells(1, 3).Value = settingValue
            .Cells(1, 4).Value = "String"
            .Cells(1, 5).Value = Now()
            .Cells(1, 6).Value = Now()
            .Cells(1, 7).Value = "System"
            .Cells(1, 8).Value = STATUS_ACTIVE
        End With
    End If
    
    Exit Sub
ErrorHandler:
    LogAuditTrail "System", AUDIT_UPDATE, TBL_SETTINGS, 0, "UpdateSettingValue:" & settingKey, "", "Failed", Err.Description
End Sub

''
' GetCompanyName
' Purpose: Retrieve company name from settings
' Returns: Company name string
' Error Handling: Returns default on error
'
Public Function GetCompanyName() As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Sheets(WS_COMPANY)
    Set tbl = ws.ListObjects(TBL_COMPANY_INFO)
    
    If tbl.ListRows.Count > 0 Then
        GetCompanyName = tbl.ListRows(1).Range.Cells(1, 2).Value
    Else
        GetCompanyName = "Chokcity Bakery"
    End If
    
    Exit Function
ErrorHandler:
    GetCompanyName = "Chokcity Bakery"
End Function

''
' GetCompanyTaxRate
' Purpose: Retrieve company tax rate
' Returns: Tax rate as decimal (e.g., 0.15 for 15%)
' Error Handling: Returns default on error
'
Public Function GetCompanyTaxRate() As Double
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Sheets(WS_COMPANY)
    Set tbl = ws.ListObjects(TBL_COMPANY_INFO)
    
    If tbl.ListRows.Count > 0 Then
        GetCompanyTaxRate = CDbl(tbl.ListRows(1).Range.Cells(1, 11).Value)
    Else
        GetCompanyTaxRate = 0.15
    End If
    
    Exit Function
ErrorHandler:
    GetCompanyTaxRate = 0.15
End Function

''
' GetNextDocumentNumber
' Purpose: Generate next document number in sequence
' Parameters: documentType - Document type (e.g., "Invoice", "Purchase Order")
' Returns: Next document number string (e.g., "INV-2026-000001")
' Error Handling: Returns empty string on error and logs to audit trail
'
Public Function GetNextDocumentNumber(documentType As String) As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Integer
    Dim prefix As String
    Dim nextSeq As Integer
    Dim yearFormat As String
    Dim padLength As Integer
    Dim yearPart As String
    Dim seqPart As String
    
    Set ws = ThisWorkbook.Sheets(WS_SEQUENCES)
    Set tbl = ws.ListObjects(TBL_NUMBER_SEQUENCES)
    
    'Find the sequence record
    For i = 1 To tbl.ListRows.Count
        If tbl.ListRows(i).Range.Cells(1, 2).Value = documentType Then
            'Get values
            prefix = tbl.ListRows(i).Range.Cells(1, 4).Value
            nextSeq = tbl.ListRows(i).Range.Cells(1, 3).Value
            yearFormat = tbl.ListRows(i).Range.Cells(1, 6).Value
            padLength = tbl.ListRows(i).Range.Cells(1, 7).Value
            
            'Increment sequence
            tbl.ListRows(i).Range.Cells(1, 3).Value = nextSeq + 1
            tbl.ListRows(i).Range.Cells(1, 11).Value = Now()
            
            'Build year part
            If yearFormat = "YYYY" Then
                yearPart = CStr(Year(Now()))
            Else
                yearPart = Format(Year(Now()), "YY")
            End If
            
            'Build sequence part
            seqPart = Format(nextSeq, String(padLength, "0"))
            
            'Build final number
            GetNextDocumentNumber = prefix & "-" & yearPart & "-" & seqPart
            
            LogAuditTrail "System", "GenerateNumber", TBL_NUMBER_SEQUENCES, i, _
                          "Generated: " & GetNextDocumentNumber, "", "Success"
            
            Exit Function
        End If
    Next i
    
    'Not found
    GetNextDocumentNumber = ""
    LogAuditTrail "System", "GenerateNumber", TBL_NUMBER_SEQUENCES, 0, _
                  "DocumentType not found: " & documentType, "", "Failed"
    
    Exit Function
ErrorHandler:
    GetNextDocumentNumber = ""
    LogAuditTrail "System", "GenerateNumber", TBL_NUMBER_SEQUENCES, 0, _
                  documentType, "", "Failed", Err.Description
End Function
