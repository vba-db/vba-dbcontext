VERSION 0.1 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------------
' DbContext - Database Connection Management Class
'
' Open source library on GitHub: https://github.com/vba-db/vba-dbcontext
' License: MIT License (see LICENSE file)
'
' Description:
'   Provides a unified interface for connecting to SQL Server and Access databases.
'   Supports parameterized queries, transactions, and data synchronization.
'
' Author: Takashi Saito
' Version: 0.1
' ------------------------------------------------------------
' Database Connection Management Class
Option Explicit

Public Enum DbProvider
    SQLServer = 1
    Access = 2
End Enum

Private Const CONNECTION_STRING As String = "{YOUR_CONNECTION_STRING_HERE}"

Private pProvider As DbProvider
Private pConnection As ADODB.Connection
Private pCommand As ADODB.Command
Private pRecordset As ADODB.Recordset
Private pLastError As String

' Database connection property
Public Property Get LastError() As String
    LastError = pLastError
End Property

' Initialize database connection
Public Sub Initialize(Optional ByVal provider As DbProvider = DbProvider.SQLServer)
    On Error GoTo ErrorHandler
    
    pProvider = provider
        
    ' Set connection string according to provider
    Select Case pProvider
        Case DbProvider.SQLServer
            Set pConnection = New ADODB.Connection
            pConnection.ConnectionString = CONNECTION_STRING
            pConnection.Open
            
        Case DbProvider.Access
            Set pConnection = CurrentProject.Connection
            
    End Select
    
    ' Setup command object
    Set pCommand = New ADODB.Command
    pCommand.ActiveConnection = pConnection
    pCommand.CommandType = adCmdText
    
    Exit Sub
    
ErrorHandler:
    pLastError = "Connection Error: " & Err.Description
    Debug.Print pLastError
End Sub

' Terminate database connection
Public Sub Terminate()
    If Not pConnection Is Nothing Then
        If pConnection.State = adStateOpen Then
            pConnection.Close
        End If
        Set pConnection = Nothing
    End If
    
    If Not pRecordset Is Nothing Then
        If pRecordset.State = adStateOpen Then
            pRecordset.Close
        End If
        Set pRecordset = Nothing
    End If
End Sub

' Add parameter
Public Sub AddParameter(ByVal name As String, ByVal value As Variant, _
                       Optional ByVal dataType As ADODB.DataTypeEnum = adVarChar)
    On Error GoTo ErrorHandler
    
    Dim param As ADODB.Parameter
    
    ' Create parameter
    Set param = pCommand.CreateParameter(name, dataType)
    
    ' Handle NULL value
    If IsNull(value) Then
        param.value = Null
    Else
        ' Set value based on data type
        Select Case dataType
            Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar
                param.Size = Len(CStr(value))
                param.value = CStr(value)
            
            Case adInteger, adSmallInt, adTinyInt, adBigInt
                param.value = CLng(value)
            
            Case adDouble, adSingle, adDecimal, adCurrency
                param.value = CDbl(value)
            
            Case adDate, adDBDate, adDBTime, adDBTimeStamp
                param.value = CDate(value)
            
            Case adBoolean
                param.value = CBool(value)
            
            Case Else
                param.value = value
        End Select
    End If
    
    ' Append parameter
    pCommand.Parameters.Append param
    
    Exit Sub
    
ErrorHandler:
    pLastError = "Parameter Addition Error: " & Err.Description
End Sub

' Clear parameters
Public Sub ClearParameters()
    On Error Resume Next
    
    ' Clear existing parameters
    Do While pCommand.Parameters.Count > 0
        pCommand.Parameters.Delete 0
    Loop
End Sub

' General-purpose select method
Public Function SelectQuery(ByVal sqlQuery As String) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    ' Close existing recordset if any
    If Not pRecordset Is Nothing Then
        If pRecordset.State = adStateOpen Then
            pRecordset.Close
        End If
    End If
    
    ' Create new recordset
    Set pRecordset = New ADODB.Recordset
    
    ' Set SQL
    pCommand.CommandText = sqlQuery
    
    pRecordset.CursorLocation = adUseClient
    pRecordset.Open pCommand, , adOpenStatic, adLockReadOnly
    
    Set SelectQuery = pRecordset
    Exit Function
    
ErrorHandler:
    pLastError = "Query Error: " & Err.Description
    Debug.Print pLastError
    Set SelectQuery = Nothing
End Function

' Execute update SQL method (for UPDATE, INSERT, DELETE)
Public Function ExecuteQuery(ByVal sqlQuery As String, _
                           Optional ByVal useTransaction As Boolean = True) As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Begin transaction (optional)
    If useTransaction Then BeginTransaction
    
    ' Set SQL
    pCommand.CommandText = sqlQuery
    
    ' Execute query
    pCommand.Execute
    
    ' Commit transaction (optional)
    If useTransaction Then CommitTransaction
    
    ExecuteQuery = True
    Exit Function
    
ErrorHandler:
    pLastError = "Query Execution Error: " & Err.Description
    Debug.Print pLastError
    
    ' Rollback on error (optional)
    If useTransaction Then
        On Error Resume Next
        RollbackTransaction
    End If
    
    ExecuteQuery = False
End Function

' Execute update SQL with OUTPUT (for UPDATE, INSERT, DELETE) – returns recordset
Public Function ExecuteQueryWithOutput(ByVal sqlQuery As String, _
                           Optional ByVal useTransaction As Boolean = True) As ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    ' Begin transaction (optional)
    If useTransaction Then BeginTransaction
    
    ' Set SQL
    pCommand.CommandText = sqlQuery
    
    ' Execute command
    Set pRecordset = pCommand.Execute
    
    ' Commit transaction (optional)
    If useTransaction Then CommitTransaction
    
    Set ExecuteQueryWithOutput = pRecordset
    Exit Function
    
ErrorHandler:
    pLastError = "Query Execution Error: " & Err.Description
    Debug.Print pLastError
    
    ' Rollback on error (optional)
    If useTransaction Then
        On Error Resume Next
        RollbackTransaction
    End If
    
    Set ExecuteQueryWithOutput = Nothing
End Function

' Insert method from recordset
Public Function InsertQuery(ByVal tableName As String, ByVal identityField As String, ByRef sourceRs As ADODB.Recordset) As ADODB.Recordset
    
    ' Recordset checks
    If sourceRs Is Nothing Then
        pLastError = "Error: Source recordset is Null"
        Debug.Print pLastError
        Set InsertQuery = Nothing
        Exit Function
    End If
    
    If sourceRs.BOF And sourceRs.EOF Then
        pLastError = "Error: Source recordset is empty"
        Debug.Print pLastError
        Set InsertQuery = Nothing
        Exit Function
    End If
    
    If sourceRs.State <> adStateOpen Then
        pLastError = "Error: Source recordset is not open"
        Debug.Print pLastError
        Set InsertQuery = Nothing
        Exit Function
    End If
    
    Dim sqlQuery As String
    Dim fields As String
    Dim values As String
    Dim allValues As String
    Dim outputFields As String
    Dim field As ADODB.field
    Dim skipIdentity As Boolean
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    ' Set skip identity field according to provider
    skipIdentity = (pProvider = DbProvider.SQLServer)
        
    ' Insert data differently based on provider
    Select Case pProvider
        Case DbProvider.SQLServer ' Insert using performance-optimized SQL
        
            ' Retrieve field names
            fields = ""
            For Each field In sourceRs.fields
                ' Skip identity field only for SQL Server
                If Not (skipIdentity And LCase(field.name) = LCase(identityField)) Then
                    fields = fields & field.name
                    If Not field Is sourceRs.fields(sourceRs.fields.Count - 1) Then
                        fields = fields & ", "
                    End If
                End If
            Next field
            
            ' Remove trailing comma
            If Right(fields, 2) = ", " Then fields = Left(fields, Len(fields) - 2)
            
            ' Get values for all records
            allValues = ""
'            sourceRs.MoveFirst is unnecessary as we immediately use the open one, preventing command re-execution.
            Do While Not sourceRs.EOF
                ' Get values of each record
                values = "("
                For Each field In sourceRs.fields
                    ' Skip identity field only for SQL Server
                    If Not (skipIdentity And LCase(field.name) = LCase(identityField)) Then
                        values = values & ConvertToSql(field)
                        If Not field Is sourceRs.fields(sourceRs.fields.Count - 1) Then
                            values = values & ", "
                        End If
                    End If
                Next field
                
                ' Remove trailing comma
                If Right(values, 2) = ", " Then values = Left(values, Len(values) - 2)
                values = values & ")"
                
                ' Add comma between records
                If Not sourceRs.EOF Then
                    allValues = allValues & values & ", "
                End If
                
                sourceRs.MoveNext
            Loop
            
            ' Remove trailing comma
            If Right(allValues, 2) = ", " Then allValues = Left(allValues, Len(allValues) - 2)
            
            ' Build field list for OUTPUT clause
            outputFields = ""
            For Each field In sourceRs.fields
                outputFields = outputFields & "INSERTED." & field.name
                If Not field Is sourceRs.fields(sourceRs.fields.Count - 1) Then
                    outputFields = outputFields & ", "
                End If
            Next field
            
            ' Execute insert and get inserted data
            sqlQuery = "INSERT INTO " & tableName & " (" & fields & ") " & _
                      "OUTPUT " & outputFields & " " & _
                      "VALUES " & allValues
'            Debug.Print sqlQuery
            Set pRecordset = pConnection.Execute(sqlQuery)
            
        Case DbProvider.Access ' Insert using recordset (Access does not support multi-row INSERT)
            ' Execute insert
            Set pRecordset = New ADODB.Recordset
            
            pRecordset.Open tableName, pConnection, adOpenStatic, adLockOptimistic
'            sourceRs.MoveFirst is unnecessary as we immediately use the open one, preventing command re-execution.
            
            Do Until sourceRs.EOF
                
                pRecordset.AddNew
                ' Only set values if field names match
                For Each field In sourceRs.fields
                    On Error Resume Next
                    If Not pRecordset.fields(field.name) Is Nothing Then
                        If field.Type = adBoolean Then
                            pRecordset.fields(field.name).value = IIf(field.value = True, -1, 0)
                        Else
                            pRecordset.fields(field.name).value = field.value
                        End If
                    End If
                    On Error GoTo ErrorHandler
                Next field
                pRecordset.Update
                sourceRs.MoveNext
            Loop
                        
    End Select
            
    Set InsertQuery = pRecordset
        
    Exit Function
    
ErrorHandler:
    pLastError = "Insert Error: " & Err.Description
    Debug.Print pLastError
    Set InsertQuery = Nothing
End Function

' General update method
Public Function UpdateQuery(ByVal tableName As String, ByVal identityField As String, ByRef sourceRs As ADODB.Recordset) As ADODB.Recordset
    
    ' Recordset checks
    If sourceRs Is Nothing Then
        pLastError = "Error: Source recordset is Null"
        Debug.Print pLastError
        Set UpdateQuery = Nothing
        Exit Function
    End If
    
    If sourceRs.BOF And sourceRs.EOF Then
        pLastError = "Error: Source recordset is empty"
        Debug.Print pLastError
        Set UpdateQuery = Nothing
        Exit Function
    End If
    
    If sourceRs.State <> adStateOpen Then
        pLastError = "Error: Source recordset is not open"
        Debug.Print pLastError
        Set UpdateQuery = Nothing
        Exit Function
    End If
    
    Dim sqlQuery As String
    Dim fields As String
    Dim values As String
    Dim outputFields As String
    Dim field As ADODB.field
    
    On Error GoTo ErrorHandler
    
    ' Retrieve field names and values
    fields = ""
    values = ""
    outputFields = ""
    
    For Each field In sourceRs.fields
        If LCase(field.name) <> LCase(identityField) Then
            ' Build field = value pairs for UPDATE clause
            If fields <> "" Then fields = fields & ", "
            fields = fields & field.name & " = " & ConvertToSql(field)
        End If
        
        ' Build field list for OUTPUT clause
        If outputFields <> "" Then outputFields = outputFields & ", "
        outputFields = outputFields & "INSERTED." & field.name
    Next field
    
    ' Build SQL differently according to provider
    Select Case pProvider
        Case DbProvider.SQLServer
            sqlQuery = "UPDATE " & tableName & " " & _
                      "SET " & fields & " " & _
                      "OUTPUT " & outputFields & " " & _
                      "WHERE " & identityField & " = " & sourceRs.fields(identityField).value
            
            Set pRecordset = New ADODB.Recordset
            
            pRecordset.Open sqlQuery, pConnection, adOpenStatic, adLockReadOnly
            
        Case DbProvider.Access
            sqlQuery = "UPDATE " & tableName & " " & _
                      "SET " & fields & " " & _
                      "WHERE " & identityField & " = " & sourceRs.fields(identityField).value
            
            pConnection.Execute sqlQuery
            
            ' Retrieve updated record
            sqlQuery = "SELECT * FROM " & tableName & " WHERE " & identityField & " = " & sourceRs.fields(identityField).value
            Set pRecordset = New ADODB.Recordset
            pRecordset.Open sqlQuery, pConnection, adOpenStatic, adLockReadOnly
    End Select
    
    Set UpdateQuery = pRecordset
    Exit Function
    
ErrorHandler:
    pLastError = "Update Error: " & Err.Description
    Debug.Print pLastError
    Set UpdateQuery = Nothing
End Function

' General delete method
Public Function DeleteQuery(ByVal tableName As String, Optional ByVal whereClause As String = "") As Boolean
    Dim sqlQuery As String
    
    On Error GoTo ErrorHandler
    
    ' Assemble SQL statement
    sqlQuery = "DELETE FROM " & tableName
    If whereClause <> "" Then
        sqlQuery = sqlQuery & " WHERE " & whereClause
    End If
    
    ' Execute query
    pConnection.Execute sqlQuery
    
    DeleteQuery = True
    Exit Function
    
ErrorHandler:
    pLastError = "Delete Error: " & Err.Description
    Debug.Print pLastError
    DeleteQuery = False
End Function

' Transaction management methods
Public Sub BeginTransaction()
    pConnection.BeginTrans
End Sub

Public Sub CommitTransaction()
    pConnection.CommitTrans
End Sub

Public Sub RollbackTransaction()
    pConnection.RollbackTrans
End Sub

' Converts recordset field type to SQL literal format
Private Function ConvertToSql(ByVal field As ADODB.field) As String
    Dim sqlString As String
    Dim value As Variant
    
    value = field.value
    
    ' Recordset through class yields Empty, direct recordset yields Null
    If IsNull(value) Or IsEmpty(value) Then
        sqlString = "NULL"
        ConvertToSql = sqlString
        Exit Function
    End If
    
    Select Case field.Type
        ' String types
        Case adChar, adVarChar, adLongVarChar
            sqlString = "'" & Replace(value, "'", "''") & "'"
            
        ' Unicode string types
        Case adWChar, adVarWChar, adLongVarWChar
            Select Case pProvider
                Case DbProvider.SQLServer
                    sqlString = "N'" & Replace(value, "'", "''") & "'"
                Case DbProvider.Access
                    sqlString = "'" & Replace(value, "'", "''") & "'"
            End Select
            
        ' Date types
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            Select Case pProvider
                Case DbProvider.SQLServer
                    sqlString = "'" & Format(value, "yyyy-mm-dd hh:nn:ss") & "'"
                Case DbProvider.Access
                    sqlString = "#" & Format(value, "yyyy/mm/dd hh:nn:ss") & "#"
            End Select
            
        ' Boolean types
        Case adBoolean
            Select Case pProvider
                Case DbProvider.SQLServer
                    sqlString = IIf(CBool(value), "1", "0")
                Case DbProvider.Access
                    sqlString = IIf(CBool(value), "True", "False")
            End Select
            
        ' Currency types
        Case adCurrency
            sqlString = Replace(CStr(value), ",", ".")
            
        ' Binary types
        Case adBinary, adVarBinary, adLongVarBinary
            sqlString = "0x" & value
            
        ' GUID types
        Case adGUID
            sqlString = "'" & value & "'"
            
        ' Numeric types (integer)
        Case adTinyInt, adSmallInt, adInteger, adBigInt
            sqlString = CStr(value)
            
        ' Numeric types (decimal)
        Case adSingle, adDouble, adDecimal, adNumeric
            sqlString = Replace(CStr(value), ",", ".")
            
        ' Other data types
        Case Else
            sqlString = CStr(value)
            sqlString = Replace(sqlString, ",", ".")
    End Select
    
    ConvertToSql = sqlString
End Function

' Method used mainly for data synchronization from SQL Server to Access
Public Function CloneData(ByVal targetProvider As DbProvider, _
                           ByVal tableName As String, _
                           ByVal identityField As String, _
                           ByVal sourceSQL As String, _
                           Optional pageNum As Long = 0, _
                           Optional sortField As String = "") As Boolean
                           
    On Error GoTo ErrorHandler
    
    Dim targetDb As New DbContext
    Dim sourceRs As ADODB.Recordset
    Dim resultRs As ADODB.Recordset
    Dim sourceCount As Long
    
    ' Initialize progress bar
    SysCmd acSysCmdInitMeter, "Loading data...", 4
    
    ' Without pagination
    If pageNum = 0 Then
    
        ' Retrieve data from source DB (using specified SQL)
        Set sourceRs = SelectQuery(sourceSQL)
        
        ' Check number of source data records
        If sourceRs Is Nothing Then
            pLastError = "Failed to retrieve source data: " & LastError
            CloneData = False
            Exit Function
        End If
        
        sourceCount = sourceRs.recordCount
            
    ' With pagination
    Else
    
        ' Get total record count
        If TempVars!CurrentRecordCount = 0 Then
            
            Set sourceRs = SelectQuery("SELECT COUNT(*) FROM (" & sourceSQL & ") Q;")
            sourceCount = sourceRs(0)
            ' Save record count to global variable
            TempVars!CurrentRecordCount = sourceCount
        
        End If
        
        If sortField = "" Then
            sourceSQL = sourceSQL & " ORDER BY " & identityField & " OFFSET " & (pageNum - 1) * gMaxRows & " ROWS FETCH NEXT " & gMaxRows & " ROWS ONLY;"
        Else
            sourceSQL = sourceSQL & " ORDER BY " & sortField & " OFFSET " & (pageNum - 1) * gMaxRows & " ROWS FETCH NEXT " & gMaxRows & " ROWS ONLY;"
        End If
        
        Set sourceRs = SelectQuery(sourceSQL)
        
        If sourceRs Is Nothing Then
            pLastError = "Failed to retrieve source data: " & LastError
            CloneData = False
            Exit Function
        End If
        
        sourceCount = sourceRs.recordCount
        
    End If
        
    ' Update progress bar
    SysCmd acSysCmdUpdateMeter, 2
        
    ' Prepare target DB
    targetDb.Initialize targetProvider
    
    ' Begin transaction
    targetDb.BeginTransaction
    
    ' Clear target table only for Access
    If targetProvider = DbProvider.Access Then
        If Not targetDb.DeleteQuery(tableName) Then
            pLastError = "Failed to clear target table: " & targetDb.LastError
            CloneData = False
            GoTo RollbackAndCleanUp
        End If
    End If
    
    ' Update progress bar
    SysCmd acSysCmdUpdateMeter, 3
    
    ' If source data count is zero, exit normally (transaction will commit)
    If sourceCount = 0 Then
        pLastError = "No data to transfer"
        CloneData = True
        GoTo CommitAndCleanUp
    End If
    
    ' Insert data
    Set resultRs = targetDb.InsertQuery(tableName, identityField, sourceRs)
    
    ' Update progress bar
    SysCmd acSysCmdUpdateMeter, 4
    
    ' Check results
    If resultRs Is Nothing Then
        pLastError = "Data insertion failed: " & targetDb.LastError
        CloneData = False
        GoTo RollbackAndCleanUp
    End If
    
    If resultRs.recordCount <> sourceCount Then
        pLastError = "Only part of the data was transferred. Transferred: " & resultRs.recordCount & _
                    " rows / Source data: " & sourceCount & " rows"
        CloneData = False
        GoTo RollbackAndCleanUp
    End If
    
    CloneData = True
    
CommitAndCleanUp:
    ' Commit transaction
    targetDb.CommitTransaction
    GoTo FinalCleanUp
    
RollbackAndCleanUp:
    ' Rollback transaction
    targetDb.RollbackTransaction
    
FinalCleanUp:
    
    ' Remove progress bar
    SysCmd acSysCmdRemoveMeter
    
    ' Release resources
    If Not sourceRs Is Nothing Then
        If sourceRs.State = adStateOpen Then sourceRs.Close
        Set sourceRs = Nothing
    End If
    If Not resultRs Is Nothing Then
        If resultRs.State = adStateOpen Then resultRs.Close
        Set resultRs = Nothing
    End If
    If Not targetDb Is Nothing Then targetDb.Terminate
    Exit Function
    
ErrorHandler:
    pLastError = "Unexpected Error: " & Err.Description
    CloneData = False
    ' On error, rollback and clean up
    Resume RollbackAndCleanUp
End Function

' Data synchronization method: updates target data based on selected multiple source data
Public Function SyncData(ByVal targetProvider As DbProvider, _
                        ByVal sqlQuery As String, _
                        ByVal updateTableName As String, _
                        ByVal keyFieldName As String) As Boolean
                        
    Dim targetDb As New DbContext
    Dim sourceRs As ADODB.Recordset
    Dim updateSql As String
    Dim updateFields As String
    Dim field As ADODB.field
    
    On Error GoTo ErrorHandler
    
    ' Retrieve data from source DB
    Set sourceRs = SelectQuery(sqlQuery)
    
    ' Error check for data retrieval
    If sourceRs Is Nothing Then
        pLastError = "Failed to retrieve source data"
        SyncData = False
        Exit Function
    End If
    
    ' Exit if no records exist
    If sourceRs.EOF And sourceRs.BOF Then
        pLastError = "No records to update"
        SyncData = False
        GoTo CleanUp
    End If
    
    ' Prepare target DB
    targetDb.Initialize targetProvider
    targetDb.BeginTransaction
    
    ' Process each record
    sourceRs.MoveFirst
    Do Until sourceRs.EOF
        
        targetDb.UpdateQuery updateTableName, keyFieldName, sourceRs
        
        sourceRs.MoveNext
    Loop
    
    ' Commit transaction
    targetDb.CommitTransaction
    SyncData = True
    
CleanUp:
    ' Release resources
    If Not sourceRs Is Nothing Then
        If sourceRs.State = adStateOpen Then sourceRs.Close
        Set sourceRs = Nothing
    End If
    If Not targetDb Is Nothing Then targetDb.Terminate
    Exit Function
    
ErrorHandler:
    pLastError = "Synchronization Error: " & Err.Description
    
    ' Rollback on error
    If Not targetDb Is Nothing Then
        targetDb.RollbackTransaction
    End If
    
    SyncData = False
    Resume CleanUp
End Function
