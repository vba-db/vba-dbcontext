# VBA DbContext

A lightweight VBA class module that provides a unified interface for connecting to SQL Server and Access databases using ADODB.  
Supports parameterized queries, transactions, and one-step data synchronization.

## Features

- **Unified Connection**  
  Easily switch between SQL Server and Access by changing one provider setting.
- **Parameterized Queries**  
  Safe handling of text, numbers, dates, booleans, and binary data.
- **Transactions**  
  Begin, commit, and rollback database operations.
- **Data Synchronization**  
  Clone or sync data records between SQL Server and Access with progress reporting.
- **Error Handling**  
  Capture the last error message via the `LastError` property.

## Prerequisites

- Microsoft Office VBA host (Access, Excel, etc.) with ADODB support.
- A reference to the **Microsoft ActiveX Data Objects** library in your VBA project.

## Installation

1. Open your VBA editor (e.g., Access VBA, Excel VBA).
2. Right-click on **Modules** (or Class Modules) and choose **Import File…**.
3. Select `DbContext.cls` from this repository.
4. In **Tools > References**, ensure **Microsoft ActiveX Data Objects 6.1 Library** is checked.
5. Open `DbContext.cls` and replace the placeholder in `CONNECTION_STRING` with your actual connection string if needed.

## Connection String Examples

```vb
' SQL Server (Windows Authentication)
"Provider=SQLOLEDB;Data Source=SERVER_NAME;Initial Catalog=DATABASE_NAME;Integrated Security=SSPI;"

' SQL Server (SQL Authentication)
"Provider=SQLOLEDB;Data Source=SERVER_NAME;Initial Catalog=DATABASE_NAME;User ID=USERNAME;Password=PASSWORD;"

' Azure SQL Database
"Provider=MSOLEDBSQL;Server=tcp:YOUR_SERVER.database.windows.net,1433;Database=YOUR_DATABASE;User ID=YOUR_USER;Password=YOUR_PASSWORD;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"
```

## Usage

```vb
'--- Initialize a new DbContext instance ---
Dim db As New DbContext
db.Initialize DbProvider.SQLServer  ' or DbProvider.Access

'--- Run a SELECT query ---
Dim rs As ADODB.Recordset
Set rs = db.SelectQuery("SELECT * FROM Customers WHERE Country = 'Japan';")
If Not rs Is Nothing Then
    Do While Not rs.EOF
        Debug.Print rs!CustomerName
        rs.MoveNext
    Loop
    rs.Close
End If

'--- Run an UPDATE/INSERT/DELETE query in a transaction ---
Dim success As Boolean
success = db.ExecuteQuery("UPDATE Orders SET ShippedDate = GETDATE() WHERE OrderID = 10248;")
If Not success Then
    Debug.Print "Error: " & db.LastError
End If

'--- Clean up when done ---
db.Terminate
```

## Data Synchronization Examples

```vb
' Clone all records from SQL Server table to Access
Dim result As Boolean
result = db.CloneData(DbProvider.Access, "LocalTable", "ID", "SELECT * FROM RemoteTable;")
If Not result Then Debug.Print db.LastError
```

## Properties & Methods

- `LastError` (String)  
  Returns the last error message encountered.
- `Initialize(provider As DbProvider)`  
  Establishes a connection to the specified database provider.
- `Terminate()`  
  Closes and cleans up the database connection and any open recordsets.  
- `SelectQuery(sql As String) As ADODB.Recordset`  
  Executes a SELECT SQL statement and returns the resulting ADODB.Recordset.  
- `ExecuteQuery(sql As String, Optional useTransaction As Boolean) As Boolean`  
  Executes an action query (INSERT, UPDATE, DELETE), optionally within a transaction, and returns success status.  
- `ExecuteQueryWithOutput(sql As String, Optional useTransaction As Boolean) As ADODB.Recordset`  
  Executes an action query with OUTPUT and returns the resulting ADODB.Recordset, optionally within a transaction.  
- `AddParameter(name As String, value As Variant, Optional dataType As ADODB.DataTypeEnum)`  
  Adds a parameter to the current command for parameterized queries.  
- `ClearParameters()`  
  Clears all parameters from the current command.  
- `InsertQuery(tableName As String, identityField As String, sourceRs As ADODB.Recordset) As ADODB.Recordset`  
  Inserts records from the source recordset into the specified table and returns the inserted rows.  
- `UpdateQuery(tableName As String, identityField As String, sourceRs As ADODB.Recordset) As ADODB.Recordset`  
  Updates records in the specified table based on the source recordset and returns the updated rows.  
- `DeleteQuery(tableName As String, Optional whereClause As String) As Boolean`  
  Deletes records from the specified table matching an optional WHERE clause and returns success status.  
- `BeginTransaction()`  
  Begins a database transaction.  
- `CommitTransaction()`  
  Commits the current database transaction.  
- `RollbackTransaction()`  
  Rolls back the current database transaction.  
- `CloneData(targetProvider As DbProvider, tableName As String, identityField As String, sourceSQL As String, Optional pageNum As Long, Optional sortField As String) As Boolean`  
  Clones data from the specified source SQL query into the target table on the target database, with optional pagination and sorting.  
- `SyncData(targetProvider As DbProvider, sqlQuery As String, updateTableName As String, keyFieldName As String) As Boolean`  
  Synchronizes data by updating target table rows based on the source query results, matching using the specified key field.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.

## Contributing

Contributions, issues, and feature requests are welcome. Please review [CONTRIBUTING.md](CONTRIBUTING.md) before submitting a pull request.
