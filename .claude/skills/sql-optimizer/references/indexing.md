# Indexing Strategies

> Load when: Choosing index types, composite indexes, partial indexes.

## Composite Index Column Order

The leftmost prefix rule: index on (A, B, C) supports:
- WHERE A = x
- WHERE A = x AND B = y
- WHERE A = x AND B = y AND C = z
- But NOT WHERE B = y alone

**Rule: equality columns first, then range columns.**

```sql
-- Query: WHERE status = 'active' AND created_at > '2024-01-01'
-- Index: (status, created_at) — equality first, range second
CREATE INDEX idx_orders_status_date ON orders (status, created_at);
```

## Partial Indexes

Index only the rows you actually query:

```sql
CREATE INDEX idx_orders_pending ON orders (created_at)
  WHERE status = 'pending';
```

## Covering Indexes

Include non-key columns to avoid table lookups:

```sql
CREATE INDEX idx_users_email ON users (email)
  INCLUDE (name, avatar);
-- Enables Index Only Scan for: SELECT name, avatar FROM users WHERE email = ?
```

## Expression Indexes

```sql
CREATE INDEX idx_users_lower_email ON users (LOWER(email));
CREATE INDEX idx_orders_date ON orders (DATE(created_at));
```

## When NOT to Index

- Tables under 1000 rows (sequential scan is faster)
- Columns with very low cardinality (boolean, status with 2 values)
- Write-heavy tables where read queries are rare
- Columns never used in WHERE, JOIN, or ORDER BY
