# security-least-privilege

**Priority:** HIGH  
**Category:** Security & Compliance

## Why It Matters

The principle of least privilege is fundamental to database security:
- Limits blast radius if credentials are compromised
- Prevents accidental data modifications or deletions
- Supports compliance requirements (SOC 2, HIPAA, GDPR)
- Enables audit trails with specific user actions
- Reduces risk of privilege escalation attacks

## Incorrect Code

```sql
-- ❌ BAD: Application using sa account
-- Connection string: "Server=...;User Id=sa;Password=..."

-- ❌ BAD: Application using db_owner role
CREATE USER AppUser WITH PASSWORD = 'password123';
ALTER ROLE db_owner ADD MEMBER AppUser;  -- WAY too many privileges!

-- ❌ BAD: Granting ALL permissions
GRANT ALL ON SCHEMA::dbo TO AppUser;  -- Never do this!

-- ❌ BAD: Using dbo schema for everything
CREATE TABLE dbo.SensitiveData (...);  -- dbo is accessible by db_owner members
```

## Correct Code

```sql
-- ✅ GOOD: Create custom schemas for separation
CREATE SCHEMA app AUTHORIZATION dbo;  -- Application objects
CREATE SCHEMA reports AUTHORIZATION dbo;  -- Reporting objects
CREATE SCHEMA sensitive AUTHORIZATION dbo;  -- Sensitive data

-- ✅ GOOD: Create role-based access
-- Application service account - read/write to app schema only
CREATE ROLE AppServiceRole;
GRANT SELECT, INSERT, UPDATE, DELETE ON SCHEMA::app TO AppServiceRole;
GRANT EXECUTE ON SCHEMA::app TO AppServiceRole;

-- Reporting service - read only
CREATE ROLE ReportingRole;
GRANT SELECT ON SCHEMA::reports TO ReportingRole;
GRANT SELECT ON SCHEMA::app TO ReportingRole;

-- Admin role for DDL operations (used by CI/CD, not app)
CREATE ROLE SchemaAdminRole;
GRANT ALTER, CREATE TABLE, CREATE PROCEDURE ON SCHEMA::app TO SchemaAdminRole;

-- ✅ GOOD: Create users with minimal privileges
CREATE USER AppServiceUser WITH PASSWORD = 'StrongP@ssw0rd!';
ALTER ROLE AppServiceRole ADD MEMBER AppServiceUser;

CREATE USER ReportingUser WITH PASSWORD = 'AnotherStr0ng!';
ALTER ROLE ReportingRole ADD MEMBER ReportingUser;

-- ✅ GOOD: Use contained database users (no server login needed)
CREATE USER ApiUser WITH PASSWORD = 'Str0ngP@ss!' FOR AZURE AD;

-- ✅ GOOD: Grant table-specific permissions when needed
GRANT SELECT (CustomerID, Name, Email) ON app.Customers TO AppServiceRole;
-- Note: Can't SELECT sensitive columns like SSN, CreditCard
```

## Azure AD (Microsoft Entra) Best Practices

```sql
-- ✅ GOOD: Use Azure AD authentication (no passwords in connection strings)
CREATE USER [app-service-identity] FROM EXTERNAL PROVIDER;
ALTER ROLE AppServiceRole ADD MEMBER [app-service-identity];

-- ✅ GOOD: Use Azure AD groups for role management
CREATE USER [AzureAD-Developers-Group] FROM EXTERNAL PROVIDER;
ALTER ROLE db_datareader ADD MEMBER [AzureAD-Developers-Group];

-- ✅ GOOD: Use managed identity for Azure services
-- App Service, Azure Functions, etc. can use system-assigned identity
CREATE USER [my-app-service] FROM EXTERNAL PROVIDER;
GRANT SELECT, INSERT, UPDATE ON SCHEMA::app TO [my-app-service];
```

## Row-Level Security for Multi-Tenant

```sql
-- ✅ GOOD: Implement row-level security for tenant isolation
CREATE FUNCTION security.TenantAccessPredicate(@TenantId INT)
RETURNS TABLE
WITH SCHEMABINDING
AS
RETURN SELECT 1 AS access
WHERE @TenantId = CAST(SESSION_CONTEXT(N'TenantId') AS INT);

CREATE SECURITY POLICY security.TenantPolicy
ADD FILTER PREDICATE security.TenantAccessPredicate(TenantId) ON app.Orders,
ADD FILTER PREDICATE security.TenantAccessPredicate(TenantId) ON app.Customers
WITH (STATE = ON);

-- Application sets tenant context on each connection
-- EXEC sp_set_session_context @key = N'TenantId', @value = @TenantId;
```

## Audit Permissions

```sql
-- Review current permissions
SELECT 
    dp.name AS UserName,
    dp.type_desc AS UserType,
    o.name AS ObjectName,
    p.permission_name,
    p.state_desc AS PermissionState
FROM sys.database_permissions p
JOIN sys.database_principals dp ON p.grantee_principal_id = dp.principal_id
LEFT JOIN sys.objects o ON p.major_id = o.object_id
WHERE dp.name NOT IN ('dbo', 'guest', 'INFORMATION_SCHEMA', 'sys')
ORDER BY dp.name, o.name;

-- Review role memberships
SELECT 
    r.name AS RoleName,
    m.name AS MemberName,
    m.type_desc AS MemberType
FROM sys.database_role_members rm
JOIN sys.database_principals r ON rm.role_principal_id = r.principal_id
JOIN sys.database_principals m ON rm.member_principal_id = m.principal_id
WHERE r.name NOT IN ('public')
ORDER BY r.name, m.name;

-- Find users with excessive privileges
SELECT 
    dp.name AS UserName,
    STRING_AGG(r.name, ', ') AS Roles
FROM sys.database_role_members rm
JOIN sys.database_principals r ON rm.role_principal_id = r.principal_id
JOIN sys.database_principals dp ON rm.member_principal_id = dp.principal_id
WHERE r.name IN ('db_owner', 'db_securityadmin', 'db_ddladmin')
GROUP BY dp.name;
```

## Connection String Configuration

```csharp
// ✅ GOOD: Use different connection strings for different purposes
public class ConnectionStrings
{
    // Regular application operations
    public string Application => "Server=...;User Id=app_service;...";
    
    // Read-only reporting
    public string Reporting => "Server=...;User Id=reporting_user;...;ApplicationIntent=ReadOnly";
    
    // Schema migrations (CI/CD only, not in app)
    public string Migrations => "Server=...;User Id=migration_admin;...";
}
```

## References

- [Azure SQL Database Security Best Practices](https://learn.microsoft.com/en-us/azure/azure-sql/database/security-best-practice)
- [Contained Database Users](https://learn.microsoft.com/en-us/sql/relational-databases/security/contained-database-users-making-your-database-portable)
- [Row-Level Security](https://learn.microsoft.com/en-us/sql/relational-databases/security/row-level-security)
- [Azure AD Authentication](https://learn.microsoft.com/en-us/azure/azure-sql/database/authentication-aad-overview)
