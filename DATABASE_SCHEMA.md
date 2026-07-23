# DATABASE_SCHEMA.md

## Breadora ERP Enterprise - Database Schema

### Overview
All data is stored in Excel Tables (ListObjects) for type safety, filtering, and structured references.

---

## Core Tables (Module 1: Settings & Configuration)

### tblSettings
**Purpose**: Centralized key-value configuration storage  
**Worksheet**: Settings  
**Primary Key**: SettingID  
**Rows**: Dynamic (grows with settings)

| Column | Type | Required | Notes |
|--------|------|----------|-------|
| SettingID | Integer | Yes | Auto-increment, Primary Key |
| SettingKey | Text | Yes | Unique configuration key (e.g., "LastBackupDate") |
| SettingValue | Text | Yes | Configuration value (any type stored as text) |
| SettingDataType | Text | Yes | Data type (String, Integer, Currency, Date, Boolean) |
| CreatedDate | DateTime | Yes | When setting was created |
| UpdatedDate | DateTime | Yes | When setting was last updated |
| CreatedBy | Text | Yes | User who created ("System" for automated) |
| Status | Text | Yes | Active/Inactive |

**Pre-Populated Settings**:
```
1. LastBackupDate = Current timestamp
2. BackupPath = [Workbook Path]\Backups\
3. AutoBackupEnabled = Yes
4. AutoBackupFrequency = 1 (daily)
5. BackupRetentionDays = 30
```

**Accessed By**:
- `mdlConfiguration.GetSettingValue(key, default)`
- `mdlConfiguration.UpdateSettingValue(key, value)`

---

### tblCompanyInfo
**Purpose**: Store business entity information  
**Worksheet**: Company  
**Primary Key**: CompanyID  
**Rows**: 1 (Single company per workbook)

| Column | Type | Required | Notes |
|--------|------|----------|-------|
| CompanyID | Integer | Yes | Always 1 (single company design) |
| CompanyName | Text | Yes | Legal business name |
| CompanyRegistration | Text | No | Tax/Registration number |
| Address | Text | No | Physical address |
| City | Text | No | City |
| PostalCode | Text | No | Postal code |
| Phone | Text | No | Primary contact phone |
| Email | Text | No | Primary contact email |
| Website | Text | No | Company website URL |
| Currency | Text | Yes | Default currency code (e.g., "USD") |
| TaxRate | Currency | Yes | Default tax rate (e.g., 0.15 for 15%) |
| FinancialYearStart | Date | Yes | Financial year start date |
| FinancialYearEnd | Date | Yes | Financial year end date |
| LogoPath | Text | No | Path to company logo image |
| CreatedDate | DateTime | Yes | Record creation timestamp |
| UpdatedDate | DateTime | Yes | Record last update timestamp |
| CreatedBy | Text | Yes | User who created record |

**Default Values**:
```
CompanyID: 1
CompanyName: "Chokcity Bakery"
Currency: "USD"
TaxRate: 0.15
FinancialYearStart: January 1
FinancialYearEnd: December 31
```

**Accessed By**:
- `mdlConfiguration.GetCompanyName()`
- `mdlConfiguration.GetCompanyTaxRate()`

---

### tblNumberSequences
**Purpose**: Manage automatic document numbering  
**Worksheet**: Sequences  
**Primary Key**: SequenceID  
**Rows**: 6 (One per document type)

| Column | Type | Required | Notes |
|--------|------|----------|-------|
| SequenceID | Integer | Yes | Primary Key |
| DocumentType | Text | Yes | Document type name (Invoice, Purchase Order, etc.) |
| CurrentSequence | Integer | Yes | Next number to assign |
| Prefix | Text | Yes | Document prefix (INV, PO, GRN, PROD, REC, EXP) |
| Suffix | Text | No | Optional suffix |
| YearFormat | Text | Yes | Year format (YYYY or YY) |
| PadLength | Integer | Yes | Zero-padding length (e.g., 6 for 000001) |
| ResetFrequency | Text | Yes | Reset frequency (Annual, Monthly, Never) |
| LastResetDate | Date | Yes | When counter was last reset |
| CreatedDate | DateTime | Yes | When record was created |
| UpdatedDate | DateTime | Yes | When record was last updated |
| CreatedBy | Text | Yes | User who created record |
| Status | Text | Yes | Active/Inactive |

**Pre-Configured Sequences**:
```
1. Invoice → INV-YYYY-000001
2. Purchase Order → PO-YYYY-000001
3. Goods Received Note → GRN-YYYY-000001
4. Production Order → PROD-YYYY-000001
5. Receipt → REC-YYYY-000001
6. Expense → EXP-YYYY-000001
```

**Accessed By**:
- `mdlConfiguration.GetNextDocumentNumber(documentType)`
  - Auto-increments CurrentSequence
  - Returns formatted number (e.g., "INV-2026-000001")

---

### tblAuditTrail
**Purpose**: Log all system activities for compliance and debugging  
**Worksheet**: AuditTrail  
**Primary Key**: AuditID  
**Rows**: Dynamic (grows with each system action)

| Column | Type | Required | Notes |
|--------|------|----------|-------|
| AuditID | Integer | Yes | Auto-increment, Primary Key |
| UserID | Text | Yes | User ID or "System" for automated actions |
| Action | Text | Yes | Action type (Create, Update, Delete, Login, Logout, Export, Import, Backup, Restore) |
| Module | Text | Yes | Module/table affected (e.g., "tblProducts", "tblSales") |
| RecordID | Integer | No | Record ID affected (0 for non-record actions) |
| Description | Text | No | Detailed description of action |
| Timestamp | DateTime | Yes | When action occurred |
| Status | Text | Yes | Success or Failed |
| ErrorMessage | Text | No | Error details if status is "Failed" |

**Usage**:
- Every data change is logged
- Every user action is tracked
- Failed operations are documented
- Supports compliance audits

**Accessed By**:
- `mdlConfiguration.LogAuditTrail(userID, action, module, recordID, description, oldValue, status, errorMsg)`
- `clsLogger.LogCreate()`, `clsLogger.LogUpdate()`, `clsLogger.LogDelete()`, etc.

---

## Future Tables (Planned Modules)

These tables will be created by their respective modules:

### User Management (Module 2)
- `tblUsers` - User accounts and credentials
- `tblRoles` - Role definitions
- `tblUserRoles` - User-role assignments
- `tblPermissions` - Permission definitions

### Inventory (Module 4)
- `tblProducts` - Product catalog
- `tblInventory` - Current stock levels
- `tblInventoryTransactions` - Stock movement history

### Production (Module 5)
- `tblRecipes` - Bill of Materials (BOM)
- `tblRecipeItems` - Recipe line items
- `tblProduction` - Production orders
- `tblProductionItems` - Production line items

### Sales (Module 6)
- `tblCustomers` - Customer information
- `tblSales` - Sales orders
- `tblSaleItems` - Sales line items

### Purchasing (Module 7)
- `tblSuppliers` - Supplier information
- `tblPurchases` - Purchase orders
- `tblPurchaseItems` - Purchase line items

### Finance (Module 8)
- `tblCashbook` - Cash transactions
- `tblExpenses` - Expense tracking

---

## Table Relationships

```
Configuration Tables:
├── tblSettings (singleton settings)
├── tblCompanyInfo (singleton company)
├── tblNumberSequences (document types)
└── tblAuditTrail (activity log - no foreign keys)

Future Dependencies:
├── tblUsers → tblRoles (role assignment)
├── tblProducts → tblInventory (stock tracking)
├── tblRecipes → tblRecipeItems (BOM structure)
├── tblProduction → tblProductionItems (production detail)
├── tblSales → tblSaleItems (order detail)
└── ... (continued per module)
```

---

## Data Integrity Rules

### Current (Module 1)
- ✅ CompanyID must be 1 (enforced by code)
- ✅ All tables require audit trail logging
- ✅ No NULL values in primary keys
- ✅ All timestamps use `Now()` (server time)
- ✅ CreatedBy is immutable after creation
- ✅ Status values are predefined constants

### Future (Modules 2-10)
- ⏳ Foreign key validation (Module 2+)
- ⏳ Negative stock prevention (Module 4)
- ⏳ Duplicate invoice prevention (Module 6)
- ⏳ Currency consistency (Module 8)
- ⏳ Date range validation (All modules)

---

## Performance Considerations

### Indexing
- Excel Tables automatically index primary keys
- No additional indexing configuration needed

### Query Optimization
- Structured references used throughout
- Avoid full table scans where possible
- Use `.Find()` for lookups instead of loops

### Archive Strategy
- Old audit trail entries retained for compliance
- Future module: Archive old transactions (Module 10)

