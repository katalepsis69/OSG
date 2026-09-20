# tsql-set-nocount

**Priority:** MEDIUM-HIGH  
**Category:** T-SQL Patterns

## Why It Matters

`SET NOCOUNT ON` prevents the "N rows affected" messages from being sent after each statement:
- Reduces network traffic between application and database
- Prevents these messages from interfering with application logic
- Improves performance for stored procedures with multiple statements
- Required for some ORM frameworks and data access patterns

## Incorrect Code

```sql
-- ❌ BAD: No SET NOCOUNT ON - sends row count messages
CREATE PROCEDURE UpdateOrderStatus
    @OrderID INT,
    @Status VARCHAR(20)
AS
BEGIN
    -- Each statement sends "(1 row affected)" message
    UPDATE Orders SET Status = @Status WHERE OrderID = @OrderID;
    
    INSERT INTO OrderHistory (OrderID, Status, ChangedDate)
    VALUES (@OrderID, @Status, GETDATE());
    
    UPDATE Orders SET LastModified = GETDATE() WHERE OrderID = @OrderID;
    
    SELECT OrderID, Status, LastModified FROM Orders WHERE OrderID = @OrderID;
END;
-- Network sends: "1 row affected", "1 row affected", "1 row affected", then result set

-- ❌ BAD: Causes issues with some ORMs expecting only result set
CREATE PROCEDURE GetCustomerOrders
    @CustomerID INT
AS
BEGIN
    SELECT o.OrderID, o.OrderDate, o.TotalAmount
    FROM Orders o
    WHERE o.CustomerID = @CustomerID;
END;
-- Some ORMs may misinterpret the row count message
```

## Correct Code

```sql
-- ✅ GOOD: SET NOCOUNT ON at the start
CREATE PROCEDURE UpdateOrderStatus
    @OrderID INT,
    @Status VARCHAR(20)
AS
BEGIN
    SET NOCOUNT ON;  -- Suppress row count messages
    
    UPDATE Orders SET Status = @Status WHERE OrderID = @OrderID;
    
    INSERT INTO OrderHistory (OrderID, Status, ChangedDate)
    VALUES (@OrderID, @Status, GETDATE());
    
    UPDATE Orders SET LastModified = GETDATE() WHERE OrderID = @OrderID;
    
    SELECT OrderID, Status, LastModified FROM Orders WHERE OrderID = @OrderID;
END;
-- Only the final SELECT result set is sent to client

-- ✅ GOOD: Template for all stored procedures
CREATE PROCEDURE [schema].[ProcedureName]
    @Param1 INT,
    @Param2 VARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;  -- Also recommended for transaction handling
    
    BEGIN TRY
        -- Procedure logic here
    END TRY
    BEGIN CATCH
        -- Error handling
        THROW;
    END CATCH
END;

-- ✅ GOOD: If you need row count for logic, capture it
CREATE PROCEDURE UpdateWithRowCount
    @CategoryID INT,
    @NewPrice DECIMAL(10,2),
    @RowsUpdated INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    
    UPDATE Products 
    SET Price = @NewPrice 
    WHERE CategoryID = @CategoryID;
    
    SET @RowsUpdated = @@ROWCOUNT;  -- Capture row count before NOCOUNT suppresses it
END;
```

## When to Use SET NOCOUNT OFF

```sql
-- Rare case: Need row count in output for specific client requirements
CREATE PROCEDURE LegacyImport
    @Data XML
AS
BEGIN
    SET NOCOUNT OFF;  -- Explicitly show row counts (legacy system expects them)
    
    INSERT INTO ImportTable (Data)
    SELECT x.value('.', 'VARCHAR(MAX)')
    FROM @Data.nodes('//Item') AS T(x);
END;
```

## How to Detect Missing SET NOCOUNT ON

```sql
-- Find procedures without SET NOCOUNT ON
SELECT 
    SCHEMA_NAME(o.schema_id) + '.' + o.name AS ProcedureName,
    CASE 
        WHEN m.definition LIKE '%SET NOCOUNT ON%' THEN 'Has SET NOCOUNT ON'
        ELSE 'MISSING SET NOCOUNT ON'
    END AS Status
FROM sys.objects o
JOIN sys.sql_modules m ON o.object_id = m.object_id
WHERE o.type = 'P'  -- Stored Procedures
    AND m.definition NOT LIKE '%SET NOCOUNT ON%'
ORDER BY ProcedureName;

-- Count of procedures missing SET NOCOUNT ON
SELECT 
    COUNT(*) AS ProceduresMissingNoCount
FROM sys.objects o
JOIN sys.sql_modules m ON o.object_id = m.object_id
WHERE o.type = 'P'
    AND m.definition NOT LIKE '%SET NOCOUNT ON%';
```

## Batch Update Script

```sql
-- Generate ALTER statements to add SET NOCOUNT ON
SELECT 
    'ALTER PROCEDURE ' + SCHEMA_NAME(o.schema_id) + '.' + o.name + '
AS
BEGIN
    SET NOCOUNT ON;
    ' + STUFF(m.definition, 1, CHARINDEX('AS', m.definition) + 2, '') AS UpdateScript
FROM sys.objects o
JOIN sys.sql_modules m ON o.object_id = m.object_id
WHERE o.type = 'P'
    AND m.definition NOT LIKE '%SET NOCOUNT ON%';
```

## Application Code Considerations

```csharp
// C# - @@ROWCOUNT is still available via ExecuteNonQuery return value
using var command = new SqlCommand("UpdateOrderStatus", connection);
command.CommandType = CommandType.StoredProcedure;
command.Parameters.AddWithValue("@OrderID", orderId);
command.Parameters.AddWithValue("@Status", status);
int rowsAffected = await command.ExecuteNonQueryAsync();  // Works with SET NOCOUNT ON
```

```python
# Python - Capture row count from cursor
cursor.execute("EXEC UpdateOrderStatus @OrderID=?, @Status=?", order_id, status)
rows_affected = cursor.rowcount  # Still available
```

## References

- [SET NOCOUNT Documentation](https://learn.microsoft.com/en-us/sql/t-sql/statements/set-nocount-transact-sql)
- [Stored Procedure Best Practices](https://learn.microsoft.com/en-us/sql/relational-databases/stored-procedures/stored-procedures-database-engine)
