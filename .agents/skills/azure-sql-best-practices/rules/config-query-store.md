# config-query-store

**Priority:** MEDIUM  
**Category:** Database Configuration

## Why It Matters

Query Store is essential for Azure SQL Database performance management:
- Captures query execution history and performance statistics
- Enables identification of regressed queries
- Supports automatic plan correction
- Provides insights for query tuning
- Essential for troubleshooting performance issues
- Enabled by default in Azure SQL Database (but verify configuration)

## Check Current Configuration

```sql
-- ✅ Check Query Store status and configuration
SELECT 
    actual_state_desc AS CurrentState,
    desired_state_desc AS DesiredState,
    current_storage_size_mb,
    max_storage_size_mb,
    CAST(current_storage_size_mb AS FLOAT) / max_storage_size_mb * 100 AS PercentUsed,
    readonly_reason,
    query_capture_mode_desc,
    size_based_cleanup_mode_desc,
    stale_query_threshold_days,
    max_plans_per_query,
    wait_stats_capture_mode_desc
FROM sys.database_query_store_options;
```

## Incorrect Configuration

```sql
-- ❌ BAD: Query Store disabled
ALTER DATABASE [YourDatabase] SET QUERY_STORE = OFF;

-- ❌ BAD: Capture mode set to NONE
ALTER DATABASE [YourDatabase] 
SET QUERY_STORE (QUERY_CAPTURE_MODE = NONE);

-- ❌ BAD: Storage too small (fills up and goes read-only)
ALTER DATABASE [YourDatabase] 
SET QUERY_STORE (MAX_STORAGE_SIZE_MB = 10);

-- ❌ BAD: Cleanup disabled with small storage
ALTER DATABASE [YourDatabase] 
SET QUERY_STORE (SIZE_BASED_CLEANUP_MODE = OFF, MAX_STORAGE_SIZE_MB = 100);
```

## Correct Configuration

```sql
-- ✅ GOOD: Enable Query Store with optimal settings for Azure SQL
ALTER DATABASE [YourDatabase] SET QUERY_STORE = ON;

ALTER DATABASE [YourDatabase] 
SET QUERY_STORE (
    OPERATION_MODE = READ_WRITE,           -- Actively capturing
    CLEANUP_POLICY = (STALE_QUERY_THRESHOLD_DAYS = 30),  -- Keep 30 days of data
    DATA_FLUSH_INTERVAL_SECONDS = 900,      -- Flush to disk every 15 minutes
    INTERVAL_LENGTH_MINUTES = 60,           -- Aggregate stats hourly
    MAX_STORAGE_SIZE_MB = 1000,             -- 1GB storage (adjust based on workload)
    QUERY_CAPTURE_MODE = AUTO,              -- Capture significant queries only
    SIZE_BASED_CLEANUP_MODE = AUTO,         -- Auto-cleanup when space low
    MAX_PLANS_PER_QUERY = 200,              -- Keep up to 200 plans per query
    WAIT_STATS_CAPTURE_MODE = ON            -- Capture wait statistics
);

-- ✅ GOOD: For high-throughput OLTP workloads, use CUSTOM capture
ALTER DATABASE [YourDatabase] 
SET QUERY_STORE (
    QUERY_CAPTURE_MODE = CUSTOM,
    QUERY_CAPTURE_POLICY = (
        STALE_CAPTURE_POLICY_THRESHOLD = 24 HOURS,
        EXECUTION_COUNT = 30,               -- Capture queries with 30+ executions
        TOTAL_COMPILE_CPU_TIME_MS = 1000,   -- Or high compile CPU
        TOTAL_EXECUTION_CPU_TIME_MS = 100   -- Or high execution CPU
    )
);
```

## Troubleshooting Query Store Issues

```sql
-- Check if Query Store went read-only
SELECT 
    actual_state_desc,
    readonly_reason,
    current_storage_size_mb,
    max_storage_size_mb
FROM sys.database_query_store_options
WHERE actual_state_desc = 'READ_ONLY';

-- If read-only due to space, force cleanup
ALTER DATABASE [YourDatabase] SET QUERY_STORE CLEAR;  -- Clears all data!

-- Better: Reduce data retention or increase storage
ALTER DATABASE [YourDatabase] 
SET QUERY_STORE (
    STALE_QUERY_THRESHOLD_DAYS = 14,  -- Reduce retention
    MAX_STORAGE_SIZE_MB = 2000        -- Or increase storage
);

-- Remove specific query from Query Store
EXEC sys.sp_query_store_remove_query @query_id = 12345;

-- Remove old plans
EXEC sys.sp_query_store_remove_plan @plan_id = 67890;
```

## Using Query Store for Performance Analysis

```sql
-- Top resource-consuming queries
SELECT TOP 20
    q.query_id,
    qt.query_sql_text,
    SUM(rs.count_executions) AS total_executions,
    SUM(rs.avg_duration * rs.count_executions) / 1000000 AS total_duration_sec,
    AVG(rs.avg_duration) / 1000 AS avg_duration_ms,
    SUM(rs.avg_logical_io_reads * rs.count_executions) AS total_logical_reads,
    AVG(rs.avg_logical_io_reads) AS avg_logical_reads
FROM sys.query_store_query q
JOIN sys.query_store_query_text qt ON q.query_text_id = qt.query_text_id
JOIN sys.query_store_plan p ON q.query_id = p.query_id
JOIN sys.query_store_runtime_stats rs ON p.plan_id = rs.plan_id
JOIN sys.query_store_runtime_stats_interval rsi ON rs.runtime_stats_interval_id = rsi.runtime_stats_interval_id
WHERE rsi.start_time > DATEADD(DAY, -7, GETDATE())
GROUP BY q.query_id, qt.query_sql_text
ORDER BY total_duration_sec DESC;

-- Regressed queries (queries that got slower)
SELECT 
    q.query_id,
    qt.query_sql_text,
    rs_old.avg_duration / 1000 AS old_avg_duration_ms,
    rs_new.avg_duration / 1000 AS new_avg_duration_ms,
    (rs_new.avg_duration - rs_old.avg_duration) / rs_old.avg_duration * 100 AS regression_pct
FROM sys.query_store_query q
JOIN sys.query_store_query_text qt ON q.query_text_id = qt.query_text_id
JOIN sys.query_store_plan p ON q.query_id = p.query_id
JOIN sys.query_store_runtime_stats rs_old ON p.plan_id = rs_old.plan_id
JOIN sys.query_store_runtime_stats rs_new ON p.plan_id = rs_new.plan_id
JOIN sys.query_store_runtime_stats_interval rsi_old ON rs_old.runtime_stats_interval_id = rsi_old.runtime_stats_interval_id
JOIN sys.query_store_runtime_stats_interval rsi_new ON rs_new.runtime_stats_interval_id = rsi_new.runtime_stats_interval_id
WHERE rsi_old.start_time BETWEEN DATEADD(DAY, -14, GETDATE()) AND DATEADD(DAY, -7, GETDATE())
    AND rsi_new.start_time > DATEADD(DAY, -7, GETDATE())
    AND rs_new.avg_duration > rs_old.avg_duration * 1.5  -- 50% slower
ORDER BY regression_pct DESC;

-- Force a known good plan
EXEC sys.sp_query_store_force_plan @query_id = 123, @plan_id = 456;

-- Unforce a plan
EXEC sys.sp_query_store_unforce_plan @query_id = 123, @plan_id = 456;
```

## Enable Automatic Tuning

```sql
-- ✅ Enable automatic plan correction
ALTER DATABASE [YourDatabase] 
SET AUTOMATIC_TUNING (FORCE_LAST_GOOD_PLAN = ON);

-- Check automatic tuning recommendations
SELECT 
    reason,
    score,
    JSON_VALUE(details, '$.implementationDetails.script') AS recommendation_script,
    JSON_VALUE(state, '$.currentValue') AS current_state
FROM sys.dm_db_tuning_recommendations;
```

## References

- [Query Store Overview](https://learn.microsoft.com/en-us/sql/relational-databases/performance/monitoring-performance-by-using-the-query-store)
- [Query Store Best Practices](https://learn.microsoft.com/en-us/sql/relational-databases/performance/best-practice-with-the-query-store)
- [Automatic Tuning](https://learn.microsoft.com/en-us/azure/azure-sql/database/automatic-tuning-overview)
