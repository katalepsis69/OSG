# query-avoid-select-star

**Priority:** CRITICAL  
**Category:** Query Performance

## Why It Matters

Using `SELECT *` in production code is problematic because:
- Returns unnecessary columns, increasing network traffic and memory usage
- Prevents covering index usage, forcing table/clustered index scans
- Breaks applications when table schema changes (new/removed columns)
- Makes code intent unclear and harder to maintain
- Impacts Query Store effectiveness and plan stability

## Incorrect Code

```sql
-- ❌ BAD: Using SELECT * returns all columns
SELECT * 
FROM Sales.Orders 
WHERE CustomerID = @CustomerID;

-- ❌ BAD: SELECT * in subqueries
SELECT CustomerName
FROM Sales.Customers
WHERE CustomerID IN (SELECT * FROM Sales.VIPCustomers);

-- ❌ BAD: SELECT * in CTEs
WITH RecentOrders AS (
    SELECT *
    FROM Sales.Orders
    WHERE OrderDate > DATEADD(DAY, -30, GETDATE())
)
SELECT * FROM RecentOrders;
```

## Correct Code

```sql
-- ✅ GOOD: Explicitly list only needed columns
SELECT OrderID, OrderDate, TotalAmount, Status
FROM Sales.Orders 
WHERE CustomerID = @CustomerID;

-- ✅ GOOD: Explicit columns in subqueries
SELECT CustomerName
FROM Sales.Customers
WHERE CustomerID IN (SELECT CustomerID FROM Sales.VIPCustomers);

-- ✅ GOOD: Explicit columns in CTEs
WITH RecentOrders AS (
    SELECT OrderID, CustomerID, OrderDate, TotalAmount
    FROM Sales.Orders
    WHERE OrderDate > DATEADD(DAY, -30, GETDATE())
)
SELECT OrderID, OrderDate, TotalAmount 
FROM RecentOrders;
```

## Exceptions

- `EXISTS` subqueries: `SELECT *` is acceptable since only existence is checked
  ```sql
  -- This is OK - EXISTS ignores column list
  IF EXISTS (SELECT * FROM Sales.Orders WHERE CustomerID = @ID)
  ```
- Ad-hoc exploration queries during development (never in production code)

## How to Detect

```sql
-- Find procedures with SELECT *
SELECT 
    OBJECT_SCHEMA_NAME(object_id) + '.' + OBJECT_NAME(object_id) AS ProcedureName,
    OBJECT_DEFINITION(object_id) AS Definition
FROM sys.procedures
WHERE OBJECT_DEFINITION(object_id) LIKE '%SELECT[^a-z]%*[^/]%FROM%'
    AND OBJECT_DEFINITION(object_id) NOT LIKE '%EXISTS%(%SELECT%*%'
ORDER BY ProcedureName;
```

## Application Code Detection

Look for patterns in application code:
```python
# Python - pyodbc
cursor.execute("SELECT * FROM Users")  # ❌ BAD

# Python - SQLAlchemy
session.query(User).all()  # Check generated SQL
```

```javascript
// Node.js - mssql
await pool.request().query('SELECT * FROM Users');  // ❌ BAD
```

```csharp
// C# - ADO.NET
command.CommandText = "SELECT * FROM Users";  // ❌ BAD
```

## References

- [Query Processing Architecture Guide](https://learn.microsoft.com/en-us/sql/relational-databases/query-processing-architecture-guide)
- [Azure SQL Database Performance Tuning](https://learn.microsoft.com/en-us/azure/azure-sql/database/performance-guidance)
