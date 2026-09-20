# Reading EXPLAIN Plans

> Load when: Analyzing query performance, understanding plan nodes.

## How to Get a Plan

```sql
-- PostgreSQL
EXPLAIN (ANALYZE, BUFFERS, FORMAT TEXT) SELECT ...;

-- MySQL
EXPLAIN ANALYZE SELECT ...;
```

## Key Metrics

- **actual time**: First row..last row in ms
- **rows**: Estimated vs actual. Big mismatch = stale statistics, run ANALYZE
- **Buffers**: shared hit (cache) vs shared read (disk)
- **loops**: How many times this node executes

## Node Types

| Node | Signal | Meaning |
|------|--------|---------|
| Index Scan | Good | Using index to find rows |
| Index Only Scan | Best | All data from index, no table access |
| Seq Scan | Warning | Full table scan — OK for small tables |
| Nested Loop | Check | OK if inner has index scan |
| Hash Join | Good | Efficient for large equi-joins |
| Sort | Check | External sort = needs index |
| Bitmap Heap Scan | Good | Index for multiple matches |

## Red Flags

1. Seq Scan on large table with low rows returned → missing index
2. Nested Loop with Seq Scan on inner → add index on join column
3. Sort with `external merge Disk` → add index or increase work_mem
4. Huge gap between estimated and actual rows → run ANALYZE
