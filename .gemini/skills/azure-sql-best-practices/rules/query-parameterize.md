# query-parameterize

**Priority:** CRITICAL  
**Category:** Query Performance, Security

## Why It Matters

Parameterized queries are essential because:
- **Security**: Prevents SQL injection attacks - the #1 database security risk
- **Performance**: Enables query plan caching and reuse
- **Stability**: Reduces plan cache bloat from unique query strings
- **Consistency**: Same logical query uses same cached plan

## Incorrect Code

### T-SQL Stored Procedure
```sql
-- ❌ BAD: Dynamic SQL with string concatenation
CREATE PROCEDURE GetCustomerOrders
    @CustomerName NVARCHAR(100)
AS
BEGIN
    DECLARE @SQL NVARCHAR(MAX);
    SET @SQL = N'SELECT * FROM Orders WHERE CustomerName = ''' + @CustomerName + '''';
    EXEC(@SQL);  -- SQL INJECTION VULNERABLE!
END;
```

### Application Code
```python
# ❌ BAD: String formatting in Python
customer_id = request.args.get('id')
query = f"SELECT * FROM Customers WHERE CustomerID = {customer_id}"
cursor.execute(query)  # SQL INJECTION VULNERABLE!

# ❌ BAD: String concatenation
query = "SELECT * FROM Customers WHERE CustomerID = " + customer_id
cursor.execute(query)
```

```javascript
// ❌ BAD: Template literals in Node.js
const customerId = req.params.id;
const query = `SELECT * FROM Customers WHERE CustomerID = ${customerId}`;
await pool.request().query(query);  // SQL INJECTION VULNERABLE!
```

```csharp
// ❌ BAD: String concatenation in C#
string customerId = Request.QueryString["id"];
string query = "SELECT * FROM Customers WHERE CustomerID = " + customerId;
command.CommandText = query;  // SQL INJECTION VULNERABLE!
```

## Correct Code

### T-SQL Stored Procedure
```sql
-- ✅ GOOD: Parameterized stored procedure
CREATE PROCEDURE GetCustomerOrders
    @CustomerID INT
AS
BEGIN
    SELECT OrderID, OrderDate, TotalAmount
    FROM Sales.Orders
    WHERE CustomerID = @CustomerID;
END;

-- ✅ GOOD: sp_executesql with parameters for dynamic SQL
CREATE PROCEDURE SearchProducts
    @SearchTerm NVARCHAR(100),
    @CategoryID INT = NULL
AS
BEGIN
    DECLARE @SQL NVARCHAR(MAX);
    DECLARE @Params NVARCHAR(500);
    
    SET @SQL = N'SELECT ProductID, ProductName, Price
                 FROM Products
                 WHERE ProductName LIKE @SearchPattern';
    
    IF @CategoryID IS NOT NULL
        SET @SQL = @SQL + N' AND CategoryID = @CatID';
    
    SET @Params = N'@SearchPattern NVARCHAR(100), @CatID INT';
    
    EXEC sp_executesql @SQL, @Params, 
        @SearchPattern = @SearchTerm, 
        @CatID = @CategoryID;
END;
```

### Application Code
```python
# ✅ GOOD: Parameterized query in Python (pyodbc)
customer_id = request.args.get('id')
query = "SELECT CustomerID, Name, Email FROM Customers WHERE CustomerID = ?"
cursor.execute(query, (customer_id,))

# ✅ GOOD: Named parameters
query = "SELECT * FROM Customers WHERE Name = ? AND Status = ?"
cursor.execute(query, (name, status))
```

```javascript
// ✅ GOOD: Parameterized query in Node.js (mssql)
const customerId = req.params.id;
const result = await pool.request()
    .input('customerId', sql.Int, customerId)
    .query('SELECT CustomerID, Name, Email FROM Customers WHERE CustomerID = @customerId');

// ✅ GOOD: Multiple parameters
const result = await pool.request()
    .input('name', sql.NVarChar(100), name)
    .input('status', sql.VarChar(20), status)
    .query('SELECT * FROM Customers WHERE Name = @name AND Status = @status');
```

```csharp
// ✅ GOOD: SqlParameter in C#
using var command = new SqlCommand(
    "SELECT CustomerID, Name, Email FROM Customers WHERE CustomerID = @CustomerID", 
    connection);
command.Parameters.AddWithValue("@CustomerID", customerId);

// ✅ GOOD: Strongly typed parameters
command.Parameters.Add("@CustomerID", SqlDbType.Int).Value = customerId;
```

## How to Detect

### Find Non-Parameterized Queries in Query Store
```sql
-- Queries with literals that should be parameterized
SELECT TOP 20
    q.query_id,
    qt.query_sql_text,
    SUM(rs.count_executions) AS total_executions,
    AVG(rs.avg_duration) / 1000 AS avg_duration_ms
FROM sys.query_store_query q
JOIN sys.query_store_query_text qt ON q.query_text_id = qt.query_text_id
JOIN sys.query_store_plan p ON q.query_id = p.query_id
JOIN sys.query_store_runtime_stats rs ON p.plan_id = rs.plan_id
WHERE qt.query_sql_text LIKE '%WHERE%=%''%'''  -- Literal strings
   OR qt.query_sql_text LIKE '%WHERE%=%[0-9]%' -- Literal numbers without @
GROUP BY q.query_id, qt.query_sql_text
HAVING COUNT(DISTINCT p.plan_id) > 1  -- Multiple plans = not parameterized
ORDER BY total_executions DESC;
```

### Check Plan Cache Bloat
```sql
-- Similar queries that should share a plan
SELECT 
    COUNT(*) AS plan_count,
    SUM(size_in_bytes) / 1024 AS total_kb,
    LEFT(qt.text, 100) AS query_pattern
FROM sys.dm_exec_cached_plans cp
CROSS APPLY sys.dm_exec_sql_text(cp.plan_handle) qt
WHERE qt.text LIKE '%SELECT%FROM%WHERE%'
GROUP BY LEFT(qt.text, 100)
HAVING COUNT(*) > 10
ORDER BY plan_count DESC;
```

## References

- [SQL Injection Prevention](https://learn.microsoft.com/en-us/sql/relational-databases/security/sql-injection)
- [Parameter Sniffing](https://learn.microsoft.com/en-us/sql/relational-databases/query-processing-architecture-guide)
- [sp_executesql](https://learn.microsoft.com/en-us/sql/relational-databases/system-stored-procedures/sp-executesql-transact-sql)
