Attribute VB_Name = "SQLRoutines"
Option Explicit
    
    Public Property Get MyConnection() As ADODB.Connection
        
        'Checks if a SQL connection is already open
        'If so, returns open SQL connection
        'Otherwise, opens SQL Connection with the parameters provided
    
        Const CONNECTION_STRING As String = _
            "Provider='SQLOLEDB';" & _
            "Data Source='LUS01-SQL-p03';" & _
            "Initial Catalog='ICIX-ESB-Q';" & _
            "Integrated Security='SSPI'"    'Data Source = Server name, Initial Catalog = SQL Database
        
        Static OpenConnection As ADODB.Connection
        
        Select Case True
            Case OpenConnection Is Nothing, OpenConnection.State = adStateClosed
                Set OpenConnection = New ADODB.Connection
                With OpenConnection
                    .ConnectionString = CONNECTION_STRING
                    .Open
                End With
        End Select
        
        Set MyConnection = OpenConnection
        
    End Property
    
    Public Function ExecuteQuery(SQLQuery As String) As ADODB.Recordset
    
        'Creates an ADO Command object to execute the provided SQL query
        
        Dim MyRecordSet As ADODB.Recordset
        Dim MyCommand As ADODB.Command: Set MyCommand = New ADODB.Command
    
        With MyCommand
            .ActiveConnection = MyConnection
            .CommandText = SQLQuery
            Set MyRecordSet = .Execute
        End With
        
        Set ExecuteQuery = MyRecordSet
        
        Set MyCommand = Nothing
        
    End Function
    
    Public Function FlattenArray(MDArray As Variant) As Variant
        
        'Convert a multidimensional array (MDArray) into a single-dimensional array, i.e. vector
        
        Const VECTOR_INIT As Integer = 0
        Const VECTOR_STEP As Integer = 1
        Const FIRST_DIMENSION As Integer = 1
        
        Dim x As Integer
        Dim Vector As Variant
        
        ReDim Vector(VECTOR_INIT To VECTOR_INIT)
        
        For x = LBound(MDArray, FIRST_DIMENSION) To UBound(MDArray, FIRST_DIMENSION)
            
            Vector(UBound(Vector)) = CStr(MDArray(x, FIRST_DIMENSION - 1))    'Cast to string to avoid data-type issues
            If x <> UBound(MDArray, FIRST_DIMENSION) Then _
                ReDim Preserve Vector(VECTOR_INIT To UBound(Vector) + VECTOR_STEP)
            
        Next x
        
        FlattenArray = Vector
        
        Erase Vector
        
    End Function
    
    Public Function ConstructListString(MyArray As Variant, Optional WrapInSingleQuotes As Boolean = False) As String
    
    '0.0 Setup
        
        Const DELIMITER As String = ", "
        
        Const PARENTHESES As String = "(@Values)"
        Const SINGLE_QUOTES As String = "'@Values'"
        Const VALUES_PLACEHOLDER As String = "@Values"
        
        Dim x As Integer
        Dim MyValue As String, ListString As String
        
    '1.0 Loop through all fields and construct string
    
        For x = LBound(MyArray) To UBound(MyArray)
            MyValue = CStr(MyArray(x))
            If WrapInSingleQuotes Then _
                MyValue = Replace(SINGLE_QUOTES, VALUES_PLACEHOLDER, MyValue)
            ListString = ListString & MyValue
            If x <> UBound(MyArray) Then _
                ListString = ListString & DELIMITER
        Next x
        
        ListString = Replace(PARENTHESES, VALUES_PLACEHOLDER, ListString)
        
        ConstructListString = ListString
    
    End Function
    
