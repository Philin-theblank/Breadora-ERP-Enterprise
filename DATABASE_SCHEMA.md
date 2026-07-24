# DATABASE_SCHEMA.md

## Breadora ERP Enterprise - Database Schema

**Current Version**: 0.2.0  
**Tables Created**: 9 of ~30 planned  
**Last Updated**: 2026-07-23

### Overview
All data is stored in Excel Tables (ListObjects) for type safety, filtering, and structured references.

---

## Core Configuration Tables (Module 1)

### tblSettings
**Purpose**: Centralized key-value configuration storage  
**Worksheet**: Settings  
**Primary Key**: SettingID  
**Status**: ✅ Active

| Column | Type | Notes |
|--------|------|-------|
| SettingID | Integer | Auto-increment |
| SettingKey | Text | Unique key |
| SettingValue | Text | Any value |
| SettingDataType | Text | String/Integer/Boolean/Date |
| CreatedDate | DateTime | Creation timestamp |
| UpdatedDate | DateTime | Last update |
| CreatedBy | Text | Administrator |
| Status | Text | Active/Inactive |

**Pre-Populated**:
- LastBackupDate
- BackupPath
- AutoBackupEnabled
- AutoBackupFrequency
- BackupRetentionDays

---

### tblCompanyInfo
**Purpose**: Business entity information  
**Worksheet**: Company  
**Primary Key**: CompanyID  
**Status**: ✅ Active

| Column | Type | Notes |
|--------|------|-------|
| CompanyID | Integer | Always 1 |
| CompanyName | Text | "Chokcity Bakery" |
| CompanyRegistration | Text | Tax ID |
| Address | Text | Street address |
| City | Text | City |
| PostalCode | Text | ZIP code |
| Phone | Text | Contact |
| Email | Text | Email |
| Website | Text | URL |
| Currency | Text | "USD" |
| TaxRate | Currency | 0.15 (15%) |
| FinancialYearStart | Date | Jan 1 |
| FinancialYearEnd | Date | Dec 31 |
| LogoPath | Text | Image path |
| CreatedDate | DateTime | Creation |
| UpdatedDate | DateTime | Last update |
| CreatedBy | Text | "System" |

---

### tblNumberSequences
**Purpose**: Automatic document numbering  
**Worksheet**: Sequences  
**Primary Key**: SequenceID  
**Status**: ✅ Active

| Column | Type | Notes |
|--------|------|-------|
| SequenceID | Integer | Primary Key |
| DocumentType | Text | Invoice, PO, GRN, etc. |
| CurrentSequence | Integer | Next number to assign |
| Prefix | Text | INV, PO, GRN, PROD, REC, EXP |
| Suffix | Text | Optional suffix |
| YearFormat | Text | YYYY or YY |
| PadLength | Integer | Zero-padding (6) |
| ResetFrequency | Text | Annual/Monthly/Never |
| LastResetDate | Date | Last reset |
| CreatedDate | DateTime | Creation |
| UpdatedDate | DateTime | Last update |
| CreatedBy | Text | "System" |
| Status | Text | Active/Inactive |

**Pre-Configured Sequences** (6):
- Invoice → INV-2026-000001
- Purchase Order → PO-2026-000001
- Goods Received Note → GRN-2026-000001
- Production Order → PROD-2026-000001
- Receipt → REC-2026-000001
- Expense → EXP-2026-000001

---

### tblAuditTrail
**Purpose**: Complete activity and compliance log  
**Worksheet**: AuditTrail  
**Primary Key**: AuditID  
**Status**: ✅ Active

| Column | Type | Notes |
|--------|------|-------|
| AuditID | Integer | Auto-increment |
| UserID | Text | User ID or "System" |
| Action | Text | Create/Update/Delete/Login/Logout |
| Module | Text | Affected module |
| RecordID | Integer | Record affected (0 for system) |
| Description | Text | Action details |
| Timestamp | DateTime | When occurred |
| Status | Text | Success/Failed |
| ErrorMessage | Text | Error details if failed |

**Logged Events**:
- All login attempts
- All data changes
- All errors
- All system actions

---

## User Management Tables (Module 2)

### tblUsers
**Purpose**: User accounts and authentication  
**Worksheet**: Users  
**Primary Key**: UserID  
**Status**: ✅ Active

| Column | Type | Notes |
|--------|------|-------|
| UserID | Integer | Primary Key |
| Username | Text | Unique login (case-insensitive) |
| PasswordHash | Text | Hashed (NEVER plaintext) |
| FullName | Text | Display name |
| Email | Text | Contact email |
| Role | Text | FK to tblRoles |
| Status | Text | Active/Inactive |
| CreatedDate | DateTime | Account creation |
| UpdatedDate | DateTime | Last modification |
| CreatedBy | Text | Administrator |
| LastLogin | DateTime | Last successful login |
| FailedAttempts | Integer | Login attempt counter |
| LockedUntil | DateTime | Account lock expiration |
| PasswordExpiry | DateTime | When password expires |

**Default Admin User**:
```
UserID: 1
Username: admin
PasswordHash: [hashed]
FullName: System Administrator
Email: admin@chokcity.local
Role: Administrator
Status: Active
```

**Security Rules**:
- ✅ Passwords hashed, never plaintext
- ✅ Minimum 8 characters
- ✅ Expire after 90 days
- ✅ Lockout after 3 failed attempts
- ✅ 30-minute lockout period

---

### tblRoles
**Purpose**: Security roles and access levels  
**Worksheet**: Roles  
**Primary Key**: RoleID  
**Status**: ✅ Active

| Column | Type | Notes |
|--------|------|-------|
| RoleID | Integer | Primary Key |
| RoleName | Text | Role name |
| Description | Text | Role purpose |
| CreatedDate | DateTime | Creation |
| UpdatedDate | DateTime | Last update |
| CreatedBy | Text | "System" |
| Status | Text | Active/Inactive |

**Pre-Configured Roles** (8):
1. Administrator - Full access
2. Manager - View all, limited admin
3. Production Manager - Production & Inventory
4. Storekeeper - Inventory only
5. Cashier - Sales & Cashbook
6. Accountant - Finance & Reports
7. Sales Representative - Sales only
8. Viewer - View-only

---

### tblPermissions
**Purpose**: Granular permission definitions  
**Worksheet**: Permissions  
**Primary Key**: PermissionID  
**Status**: ✅ Active

| Column | Type | Notes |
|--------|------|-------|
| PermissionID | Integer | Primary Key |
| PermissionName | Text | Permission ID |
| Module | Text | Module name |
| Action | Text | View/Create/Edit/Delete |
| Description | Text | Permission details |
| CreatedDate | DateTime | Creation |
| Status | Text | Active/Inactive |

**Permission Model**:
- Role-based (RBAC)
- Module-level access
- Action-based (View, Create, Edit, Delete)

---

### tblUserRoles (Junction Table)
**Purpose**: Many-to-many user-role assignments  
**Worksheet**: UserRoles  
**Primary Key**: UserRoleID  
**Status**: ✅ Active

| Column | Type | Notes |
|--------|------|-------|
| UserRoleID | Integer | Primary Key |
| UserID | Integer | FK to tblUsers |
| RoleID | Integer | FK to tblRoles |
| AssignedDate | DateTime | Assignment date |
| AssignedBy | Text | Administrator |
| Status | Text | Active/Inactive |

**Purpose**:
- Assign multiple roles per user
- Support for role history
- Deactivate roles without deleting

---

### tblLoginHistory
**Purpose**: Comprehensive login audit trail  
**Worksheet**: LoginHistory  
**Primary Key**: LoginID  
**Status**: ✅ Active

| Column | Type | Notes |
|--------|------|-------|
| LoginID | Integer | Primary Key |
| UserID | Integer | FK to tblUsers |
| Username | Text | Username at login |
| LoginTime | DateTime | Login timestamp |
| LogoutTime | DateTime | Logout timestamp |
| Status | Text | Active/Inactive |
| IPAddress | Text | Client IP (placeholder) |
| Result | Text | Success/Failed |

**Tracks**:
- ✅ All login attempts
- ✅ Failed logins with reasons
- ✅ Session duration
- ✅ User activity timeline

---

## Table Relationships

```
Configuration (Module 1):
├── tblSettings (Configuration)
├── tblCompanyInfo (Business info)
├── tblNumberSequences (Document numbering)
└── tblAuditTrail (Activity log)

User Management (Module 2):
├── tblUsers (User accounts)
├── tblRoles (Security roles)
├── tblPermissions (Permissions)
├── tblUserRoles (User-role assignments)
└── tblLoginHistory (Login audit)

Future Dependencies:
├── tblProducts → tblInventory (Stock tracking)
├── tblRecipes → tblRecipeItems (BOM structure)
├── tblCustomers → tblSales (Orders)
└── ... (Continued per module)
```

---

## Data Integrity Rules

### Implemented
- ✅ Primary key on all tables
- ✅ No NULL in primary keys
- ✅ Timestamps use `Now()`
- ✅ Status values from constants
- ✅ Audit trail logging
- ✅ Password hashing mandatory
- ✅ No plaintext sensitive data

### Planned (Future Modules)
- ⏳ Foreign key validation
- ⏳ Negative stock prevention
- ⏳ Duplicate prevention
- ⏳ Currency consistency
- ⏳ Date range validation

---

## Performance Notes

### Indexing
- Excel Tables auto-index primary keys
- UserID searches optimized
- Username searches optimized

### Query Optimization
- Structured references used
- Avoid full table scans
- Use `.Find()` for lookups

### Scalability
- Design supports 10,000+ users
- No performance issues expected until 100,000+ records
- Archive strategy (Module 10) recommended for 1M+ records

---

## Statistics

| Metric | Value |
|--------|-------|
| Total Tables | 9 |
| Total Columns | 100+ |
| Primary Keys | 9 |
| Foreign Keys (Planned) | 15+ |
| Rows Created (Default) | 15 |
| Max Design Capacity | 1M+ |

---

## Future Enhancements (Module 3+)

Tables to be added:
- tblProducts (Module 4)
- tblInventory (Module 4)
- tblRecipes (Module 5)
- tblCustomers (Module 6)
- tblSuppliers (Module 7)
- tblCashbook (Module 8)
- ... and more

