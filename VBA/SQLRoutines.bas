Attribute VB_Name = "SQLRoutines"
Option Explicit
    
    Public Property Get MyConnection(ConnectionString As String) As ADODB.Connection
        
        'Checks if a SQL connection is already open
        'If so, returns open SQL connection
        'Otherwise, opens SQL Connection with the parameters provided
        
        Static OpenConnection As ADODB.Connection
        
        Select Case True
            Case OpenConnection Is Nothing, OpenConnection.State = adStateClosed
                Set OpenConnection = New ADODB.Connection
                With OpenConnection
                    .ConnectionString = ConnectionString
                    .Open
                End With
        End Select
        
        Set MyConnection = OpenConnection
        
    End Property
    
    Public Function QueryDBTable(TableName As String, ConnectionString As String, Optional QueryString As String, Optional CloseConnection As Boolean = True) As Variant
    
        'Creates an ADO Command object to execute the provided SQL query
        
        Dim MyRecordSet As ADODB.Recordset
        Dim MyCommand As ADODB.Command: Set MyCommand = New ADODB.Command

        If QueryString = "" Then QueryString = BuildSelectAllQuery(TableName)

        With MyCommand
            .ActiveConnection = MyConnection(ConnectionString)
            .CommandText = QueryString
            Set MyRecordSet = .Execute
            If CloseConnection Then Call .ActiveConnection.Close
        End With
        
        QueryDBTable = MyRecordSet.GetRows
        
    End Function

    Private Property Get BuildSelectAllQuery(TableName As String) As String
        Const SELECT_QUERY As String = "SELECT * FROM @TableName"
        Const TABLE_NAME_PLACEHOLDER AS String = "@TableName"
        BuildSelectAllQuery = SELECT_QUERY
        BuildSelectAllQuery = Replace(BuildSelectAllQuery, TABLE_NAME_PLACEHOLDER, TableName)
    End Property    
