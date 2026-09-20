# connection-pooling

**Priority:** HIGH  
**Category:** Connection Management

## Why It Matters

Connection pooling is critical for Azure SQL Database because:
- Opening connections is expensive (authentication, TLS handshake, session setup)
- Azure SQL has connection limits based on service tier
- Pooling reduces latency by reusing existing connections
- Prevents connection exhaustion under load
- Reduces DTU/vCore overhead from connection management

## Incorrect Code

### Python
```python
# ❌ BAD: Creating new connection for each request
def get_user(user_id):
    conn = pyodbc.connect(connection_string)  # New connection each time
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Users WHERE UserID = ?", user_id)
    result = cursor.fetchone()
    conn.close()  # Connection destroyed
    return result

# ❌ BAD: Not using connection pool in SQLAlchemy
engine = create_engine(connection_string, poolclass=NullPool)  # Disables pooling
```

### Node.js
```javascript
// ❌ BAD: Creating new connection for each query
async function getUser(userId) {
    const pool = await sql.connect(config);  // Creates new pool each time!
    const result = await pool.request()
        .input('userId', sql.Int, userId)
        .query('SELECT * FROM Users WHERE UserID = @userId');
    pool.close();  // Pool destroyed
    return result.recordset[0];
}
```

### C#
```csharp
// ❌ BAD: Disabling connection pooling
var connectionString = "Server=...;Pooling=false;";  // Don't do this!

// ❌ BAD: Not disposing connections properly
public User GetUser(int userId)
{
    var connection = new SqlConnection(connectionString);
    connection.Open();
    // ... use connection ...
    // Forgot to close/dispose - connection leak!
    return user;
}
```

## Correct Code

### Python
```python
# ✅ GOOD: Use connection pooling with pyodbc
import pyodbc
from dbutils.pooled_db import PooledDB

# Create a connection pool (do this once at app startup)
pool = PooledDB(
    creator=pyodbc,
    maxconnections=20,  # Max connections in pool
    mincached=5,        # Min idle connections
    maxcached=10,       # Max idle connections
    blocking=True,      # Block when pool exhausted
    dsn=connection_string
)

def get_user(user_id):
    conn = pool.connection()  # Get connection from pool
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT UserID, Name, Email FROM Users WHERE UserID = ?", user_id)
        return cursor.fetchone()
    finally:
        conn.close()  # Returns to pool, not destroyed

# ✅ GOOD: SQLAlchemy with connection pooling (built-in)
from sqlalchemy import create_engine

engine = create_engine(
    connection_string,
    pool_size=10,           # Number of connections to keep
    max_overflow=20,        # Additional connections when pool exhausted
    pool_timeout=30,        # Seconds to wait for connection
    pool_recycle=1800,      # Recycle connections after 30 minutes
    pool_pre_ping=True      # Test connections before use
)
```

### Node.js
```javascript
// ✅ GOOD: Create pool once, reuse globally
const sql = require('mssql');

// Create pool at application startup
const poolPromise = new sql.ConnectionPool({
    server: 'your-server.database.windows.net',
    database: 'your-database',
    user: 'username',
    password: 'password',
    options: {
        encrypt: true,
        trustServerCertificate: false
    },
    pool: {
        max: 20,           // Max connections
        min: 5,            // Min connections  
        idleTimeoutMillis: 30000  // Close idle connections after 30s
    }
}).connect();

// Reuse pool in request handlers
async function getUser(userId) {
    const pool = await poolPromise;  // Get existing pool
    const result = await pool.request()
        .input('userId', sql.Int, userId)
        .query('SELECT UserID, Name, Email FROM Users WHERE UserID = @userId');
    return result.recordset[0];
    // Connection automatically returned to pool
}

// Graceful shutdown
process.on('SIGTERM', async () => {
    const pool = await poolPromise;
    await pool.close();
});
```

### C#
```csharp
// ✅ GOOD: Connection pooling is automatic with proper connection string
// Just ensure Pooling=true (default) and proper pool settings
var connectionString = @"
    Server=your-server.database.windows.net;
    Database=your-database;
    User Id=username;
    Password=password;
    Encrypt=True;
    Pooling=True;
    Min Pool Size=5;
    Max Pool Size=100;
    Connection Timeout=30;
";

// ✅ GOOD: Always use 'using' to ensure proper disposal
public async Task<User> GetUserAsync(int userId)
{
    using var connection = new SqlConnection(_connectionString);
    await connection.OpenAsync();
    
    using var command = new SqlCommand(
        "SELECT UserID, Name, Email FROM Users WHERE UserID = @UserId", 
        connection);
    command.Parameters.AddWithValue("@UserId", userId);
    
    using var reader = await command.ExecuteReaderAsync();
    if (await reader.ReadAsync())
    {
        return new User
        {
            UserId = reader.GetInt32(0),
            Name = reader.GetString(1),
            Email = reader.GetString(2)
        };
    }
    return null;
}  // Connection returned to pool here

// ✅ GOOD: Dependency injection with DbContext in ASP.NET Core
services.AddDbContext<AppDbContext>(options =>
    options.UseSqlServer(connectionString, sqlOptions =>
    {
        sqlOptions.EnableRetryOnFailure(
            maxRetryCount: 3,
            maxRetryDelay: TimeSpan.FromSeconds(30),
            errorNumbersToAdd: null);
    }));
```

## Azure SQL Connection Limits

| Service Tier | Max Concurrent Connections |
|-------------|---------------------------|
| Basic | 30 |
| S0 | 60 |
| S1 | 90 |
| S2 | 120 |
| S3+ | 200 |
| P1 | 450 |
| P2+ | 900+ |

## Monitor Connection Pool Health

```sql
-- Current connections by login/host
SELECT 
    login_name,
    host_name,
    program_name,
    COUNT(*) AS connection_count,
    SUM(CASE WHEN status = 'sleeping' THEN 1 ELSE 0 END) AS idle_connections,
    SUM(CASE WHEN status = 'running' THEN 1 ELSE 0 END) AS active_connections
FROM sys.dm_exec_sessions
WHERE is_user_process = 1
GROUP BY login_name, host_name, program_name
ORDER BY connection_count DESC;

-- Connection pool efficiency
SELECT 
    COUNT(*) AS total_connections,
    SUM(CASE WHEN status = 'sleeping' THEN 1 ELSE 0 END) AS pooled_idle,
    SUM(CASE WHEN status != 'sleeping' THEN 1 ELSE 0 END) AS in_use
FROM sys.dm_exec_sessions
WHERE is_user_process = 1;
```

## References

- [Azure SQL Database Connection Limits](https://learn.microsoft.com/en-us/azure/azure-sql/database/resource-limits-dtu-single-databases)
- [ADO.NET Connection Pooling](https://learn.microsoft.com/en-us/dotnet/framework/data/adonet/sql-server-connection-pooling)
- [Connection Resiliency](https://learn.microsoft.com/en-us/azure/azure-sql/database/troubleshoot-common-connectivity-issues)
