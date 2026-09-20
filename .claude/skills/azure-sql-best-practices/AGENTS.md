# Azure SQL Database Best Practices - Complete Agent Guide

This is the compiled reference document for AI agents. Use this guide when analyzing, writing, or refactoring code that interacts with Azure SQL Database.

## Quick Decision Tree

```
Is this a database query?
├── Does it use SELECT *? → Apply query-avoid-select-star
├── Are parameters concatenated as strings? → Apply query-parameterize (CRITICAL - security)
├── Are functions applied to columns in WHERE? → Apply query-sargable
└── Is it a cursor loop? → Apply query-avoid-cursors

Is this a stored procedure?
├── Missing SET NOCOUNT ON? → Apply tsql-set-nocount
├── No TRY-CATCH? → Apply tsql-error-handling
├── Dynamic SQL with string concatenation? → Apply query-parameterize
└── Long-running transaction? → Apply tsql-transaction-scope

Is this application database code?
├── Creating new connections per request? → Apply connection-pooling
├── No retry logic? → Apply connection-retry-logic
├── Using sa or db_owner? → Apply security-least-privilege
└── Missing async/await? → Apply connection-async

Is this database configuration?
├── Query Store disabled? → Apply config-query-store
├── Auto-tuning disabled? → Apply config-auto-tuning
└── Missing indexes reported? → Apply index-cover-queries
```

---

## Category 1: Query Performance (CRITICAL)

### query-avoid-select-star
**Never use SELECT * in production code.**

❌ BAD:
```sql
SELECT * FROM Orders WHERE CustomerID = @CustomerID;
```

✅ GOOD:
```sql
SELECT OrderID, OrderDate, TotalAmount, Status
FROM Orders 
WHERE CustomerID = @CustomerID;
```

Exception: `EXISTS` subqueries where only existence is checked.

---

### query-parameterize
**Always use parameterized queries. This is both a security AND performance issue.**

❌ BAD (SQL Injection vulnerable):
```python
query = f"SELECT * FROM Users WHERE UserID = {user_id}"
cursor.execute(query)
```

✅ GOOD:
```python
query = "SELECT UserID, Name FROM Users WHERE UserID = ?"
cursor.execute(query, (user_id,))
```

❌ BAD T-SQL:
```sql
SET @SQL = N'SELECT * FROM Orders WHERE CustomerName = ''' + @Name + ''''
EXEC(@SQL)
```

✅ GOOD T-SQL:
```sql
SET @SQL = N'SELECT OrderID FROM Orders WHERE CustomerName = @CustomerName'
EXEC sp_executesql @SQL, N'@CustomerName NVARCHAR(100)', @CustomerName = @Name
```

---

### query-sargable
**Write SARGable predicates that can use indexes.**

❌ BAD (cannot use index):
```sql
WHERE YEAR(OrderDate) = 2024
WHERE LEFT(CustomerName, 3) = 'ABC'
WHERE Price * 1.1 > 100
WHERE ISNULL(Status, 'pending') = 'pending'
```

✅ GOOD (uses index):
```sql
WHERE OrderDate >= '2024-01-01' AND OrderDate < '2025-01-01'
WHERE CustomerName LIKE 'ABC%'
WHERE Price > 100 / 1.1
WHERE (Status = 'pending' OR Status IS NULL)
```

---

### query-avoid-cursors
**Replace cursors with set-based operations.**

❌ BAD:
```sql
DECLARE order_cursor CURSOR FOR SELECT OrderID FROM Orders
OPEN order_cursor
FETCH NEXT FROM order_cursor INTO @OrderID
WHILE @@FETCH_STATUS = 0
BEGIN
    UPDATE Orders SET Processed = 1 WHERE OrderID = @OrderID
    FETCH NEXT FROM order_cursor INTO @OrderID
END
CLOSE order_cursor
DEALLOCATE order_cursor
```

✅ GOOD:
```sql
UPDATE Orders SET Processed = 1 WHERE Processed = 0;
```

---

### query-batch-operations
**Batch INSERT/UPDATE/DELETE operations.**

❌ BAD:
```python
for item in items:
    cursor.execute("INSERT INTO Items (Name) VALUES (?)", item.name)
```

✅ GOOD:
```python
cursor.executemany("INSERT INTO Items (Name) VALUES (?)", 
    [(item.name,) for item in items])
```

Or use table-valued parameters in T-SQL.

---

## Category 2: Indexing Strategy (CRITICAL)

### index-cover-queries
**Create covering indexes to eliminate key lookups.**

❌ BAD (causes key lookup):
```sql
-- Index only on CustomerID
CREATE INDEX IX_Orders_Customer ON Orders(CustomerID);

-- Query needs OrderDate and TotalAmount → Key Lookup!
SELECT OrderID, CustomerID, OrderDate, TotalAmount
FROM Orders WHERE CustomerID = @CustomerID;
```

✅ GOOD (covering index):
```sql
CREATE INDEX IX_Orders_Customer_Covering
ON Orders(CustomerID)
INCLUDE (OrderDate, TotalAmount, OrderID);
```

### index-missing-index-dmv
**Use DMVs to find missing indexes:**
```sql
SELECT 
    'CREATE INDEX IX_' + OBJECT_NAME(mid.object_id) + '_Missing ON ' + 
    mid.statement + ' (' + COALESCE(mid.equality_columns, '') + 
    COALESCE(', ' + mid.inequality_columns, '') + ')' +
    COALESCE(' INCLUDE (' + mid.included_columns + ')', '') AS CreateStatement
FROM sys.dm_db_missing_index_details mid
JOIN sys.dm_db_missing_index_groups mig ON mid.index_handle = mig.index_handle
JOIN sys.dm_db_missing_index_group_stats migs ON mig.index_group_handle = migs.group_handle
ORDER BY migs.avg_total_user_cost * migs.avg_user_impact * migs.user_seeks DESC;
```

---

## Category 3: Security (HIGH)

### security-parameterize-queries
Same as query-parameterize - SQL injection is the #1 database vulnerability.

### security-least-privilege
**Grant minimum required permissions.**

❌ BAD:
```sql
ALTER ROLE db_owner ADD MEMBER AppUser;
GRANT ALL ON SCHEMA::dbo TO AppUser;
```

✅ GOOD:
```sql
CREATE ROLE AppServiceRole;
GRANT SELECT, INSERT, UPDATE, DELETE ON SCHEMA::app TO AppServiceRole;
ALTER ROLE AppServiceRole ADD MEMBER AppUser;
```

### security-encrypt-connections
Always use `Encrypt=True` in connection strings:
```
Server=myserver.database.windows.net;Database=mydb;Encrypt=True;TrustServerCertificate=False;
```

---

## Category 4: Connection Management (HIGH)

### connection-pooling
**Always use connection pooling.**

❌ BAD Node.js:
```javascript
async function getUser(userId) {
    const pool = await sql.connect(config);  // New pool per request!
    const result = await pool.request().query('...');
    pool.close();
}
```

✅ GOOD Node.js:
```javascript
// Create pool once at startup
const poolPromise = new sql.ConnectionPool(config).connect();

async function getUser(userId) {
    const pool = await poolPromise;  // Reuse pool
    const result = await pool.request()
        .input('userId', sql.Int, userId)
        .query('SELECT * FROM Users WHERE UserID = @userId');
    return result.recordset[0];
}
```

### connection-retry-logic
**Implement retry for transient failures.**

```csharp
var options = new SqlRetryLogicOption()
{
    NumberOfTries = 3,
    DeltaTime = TimeSpan.FromSeconds(1),
    MaxTimeInterval = TimeSpan.FromSeconds(20),
    TransientErrors = new[] { 4060, 40197, 40501, 40613, 49918, 49919, 49920, 4221 }
};
connection.RetryLogicProvider = SqlConfigurableRetryFactory.CreateExponentialRetryProvider(options);
```

---

## Category 5: T-SQL Patterns (MEDIUM-HIGH)

### tsql-set-nocount
**Always use SET NOCOUNT ON in stored procedures.**

```sql
CREATE PROCEDURE MyProcedure
AS
BEGIN
    SET NOCOUNT ON;  -- First line after BEGIN
    SET XACT_ABORT ON;  -- Recommended for transaction handling
    
    -- Procedure logic
END;
```

### tsql-schema-qualify
**Always schema-qualify object names.**

❌ BAD:
```sql
SELECT * FROM Orders;
EXEC UpdateStatus;
```

✅ GOOD:
```sql
SELECT * FROM Sales.Orders;
EXEC Sales.UpdateStatus;
```

### tsql-error-handling
**Use TRY-CATCH for error handling.**

```sql
BEGIN TRY
    BEGIN TRANSACTION;
    -- Operations
    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
    THROW;  -- Re-throw the error
END CATCH
```

---

## Category 6: Database Configuration (MEDIUM)

### config-query-store
**Enable and configure Query Store.**

```sql
ALTER DATABASE [YourDatabase] SET QUERY_STORE = ON;
ALTER DATABASE [YourDatabase] SET QUERY_STORE (
    OPERATION_MODE = READ_WRITE,
    MAX_STORAGE_SIZE_MB = 1000,
    QUERY_CAPTURE_MODE = AUTO,
    SIZE_BASED_CLEANUP_MODE = AUTO
);
```

### config-auto-tuning
**Enable automatic tuning.**

```sql
ALTER DATABASE [YourDatabase] 
SET AUTOMATIC_TUNING (FORCE_LAST_GOOD_PLAN = ON);
```

---

## Detection Queries

### Find procedures without SET NOCOUNT ON:
```sql
SELECT SCHEMA_NAME(schema_id) + '.' + name AS ProcedureName
FROM sys.procedures p
JOIN sys.sql_modules m ON p.object_id = m.object_id
WHERE m.definition NOT LIKE '%SET NOCOUNT ON%';
```

### Find queries with key lookups:
```sql
SELECT OBJECT_NAME(i.object_id) AS TableName, i.name AS IndexName,
       ius.user_lookups AS KeyLookups
FROM sys.dm_db_index_usage_stats ius
JOIN sys.indexes i ON ius.object_id = i.object_id AND ius.index_id = i.index_id
WHERE ius.user_lookups > 1000
ORDER BY ius.user_lookups DESC;
```

### Find non-parameterized queries:
```sql
SELECT qt.query_sql_text, COUNT(DISTINCT p.plan_id) AS PlanCount
FROM sys.query_store_query q
JOIN sys.query_store_query_text qt ON q.query_text_id = qt.query_text_id
JOIN sys.query_store_plan p ON q.query_id = p.query_id
GROUP BY q.query_id, qt.query_sql_text
HAVING COUNT(DISTINCT p.plan_id) > 3
ORDER BY PlanCount DESC;
```

---

## Code Review Checklist

- [ ] No `SELECT *` in production code
- [ ] All queries parameterized
- [ ] Connection pooling implemented
- [ ] Retry logic for transient failures
- [ ] `SET NOCOUNT ON` in all stored procedures
- [ ] Appropriate indexes for frequent queries
- [ ] Least privilege database permissions
- [ ] Query Store enabled
- [ ] No dynamic SQL with string concatenation
- [ ] Async/await for database calls
