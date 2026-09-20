# index-cover-queries

**Priority:** CRITICAL  
**Category:** Indexing Strategy

## Why It Matters

A covering index includes all columns needed by a query, allowing the query to be satisfied entirely from the index without accessing the base table. Benefits:
- Eliminates expensive key lookups to the clustered index
- Reduces I/O significantly (index pages are much smaller than table pages)
- Improves query response time dramatically
- Reduces DTU/vCore consumption in Azure SQL Database

## Incorrect Code

```sql
-- ❌ BAD: Query causes key lookup because index doesn't cover all columns
-- Given: CREATE INDEX IX_Orders_CustomerID ON Orders(CustomerID)

SELECT OrderID, CustomerID, OrderDate, TotalAmount
FROM Orders
WHERE CustomerID = @CustomerID;
-- This query needs OrderDate and TotalAmount, but the index only has CustomerID
-- Result: Index Seek + Key Lookup (expensive for many rows)
```

## Correct Code

```sql
-- ✅ GOOD: Covering index includes all columns needed by the query
CREATE NONCLUSTERED INDEX IX_Orders_CustomerID_Covering
ON Orders(CustomerID)
INCLUDE (OrderDate, TotalAmount, OrderID);

-- Now the same query uses Index Seek only (no Key Lookup)
SELECT OrderID, CustomerID, OrderDate, TotalAmount
FROM Orders
WHERE CustomerID = @CustomerID;
```

## INCLUDE vs Key Columns

```sql
-- Key columns: Used in WHERE, JOIN, ORDER BY
-- INCLUDE columns: Only needed in SELECT

-- ❌ BAD: All columns as key columns (wider index, more overhead)
CREATE INDEX IX_Orders_Wide 
ON Orders(CustomerID, OrderDate, TotalAmount, OrderID);

-- ✅ GOOD: Only filter columns as keys, rest in INCLUDE
CREATE INDEX IX_Orders_Optimal
ON Orders(CustomerID)  -- Key: used in WHERE
INCLUDE (OrderDate, TotalAmount, OrderID);  -- Just returned, not filtered

-- ✅ GOOD: Multiple key columns when used for filtering/ordering
CREATE INDEX IX_Orders_DateRange
ON Orders(CustomerID, OrderDate)  -- Both used in WHERE
INCLUDE (TotalAmount, Status);
```

## How to Identify Missing Covering Indexes

```sql
-- Find queries with key lookups that could benefit from covering indexes
SELECT TOP 20
    OBJECT_NAME(i.object_id) AS TableName,
    i.name AS IndexName,
    ius.user_lookups AS KeyLookups,
    ius.user_seeks AS IndexSeeks,
    ius.user_scans AS IndexScans,
    CAST(ius.user_lookups AS FLOAT) / NULLIF(ius.user_seeks + ius.user_scans, 0) AS LookupRatio
FROM sys.dm_db_index_usage_stats ius
JOIN sys.indexes i ON ius.object_id = i.object_id AND ius.index_id = i.index_id
WHERE ius.database_id = DB_ID()
    AND ius.user_lookups > 0
    AND i.type_desc = 'NONCLUSTERED'
ORDER BY ius.user_lookups DESC;

-- DMV for missing index suggestions
SELECT TOP 20
    CONVERT(DECIMAL(18,2), migs.avg_total_user_cost * migs.avg_user_impact * (migs.user_seeks + migs.user_scans)) AS ImprovementMeasure,
    'CREATE INDEX IX_' + OBJECT_NAME(mid.object_id) + '_' + 
        REPLACE(REPLACE(COALESCE(mid.equality_columns, ''), '[', ''), ']', '') AS CreateIndexStatement,
    mid.equality_columns,
    mid.inequality_columns,
    mid.included_columns,
    migs.user_seeks,
    migs.user_scans,
    migs.avg_total_user_cost,
    migs.avg_user_impact
FROM sys.dm_db_missing_index_groups mig
JOIN sys.dm_db_missing_index_group_stats migs ON migs.group_handle = mig.index_group_handle
JOIN sys.dm_db_missing_index_details mid ON mig.index_handle = mid.index_handle
WHERE mid.database_id = DB_ID()
ORDER BY ImprovementMeasure DESC;
```

## Query Store Analysis

```sql
-- Find queries with high logical reads that might benefit from covering indexes
SELECT TOP 20
    q.query_id,
    qt.query_sql_text,
    SUM(rs.avg_logical_io_reads * rs.count_executions) AS total_logical_reads,
    AVG(rs.avg_logical_io_reads) AS avg_logical_reads,
    SUM(rs.count_executions) AS execution_count
FROM sys.query_store_query q
JOIN sys.query_store_query_text qt ON q.query_text_id = qt.query_text_id
JOIN sys.query_store_plan p ON q.query_id = p.query_id
JOIN sys.query_store_runtime_stats rs ON p.plan_id = rs.plan_id
GROUP BY q.query_id, qt.query_sql_text
HAVING SUM(rs.count_executions) > 100
ORDER BY total_logical_reads DESC;
```

## Trade-offs

Consider these factors when creating covering indexes:
- Each INCLUDE column adds storage overhead
- More indexes = slower INSERT/UPDATE/DELETE operations
- Analyze actual query patterns before creating covering indexes
- Use Query Store to validate index effectiveness

## References

- [Covering Indexes](https://learn.microsoft.com/en-us/sql/relational-databases/indexes/create-indexes-with-included-columns)
- [Index Design Guidelines](https://learn.microsoft.com/en-us/sql/relational-databases/sql-server-index-design-guide)
- [Query Store Overview](https://learn.microsoft.com/en-us/sql/relational-databases/performance/monitoring-performance-by-using-the-query-store)
