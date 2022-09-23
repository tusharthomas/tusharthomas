VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Const CONN_PROVIDER = "SQLOLEDB"
    Const CONN_SECURITY = "SSPI"

    Private pServerName     As String
    Private pDefaultDB      As String

    Private pMyConnection   As ADODB.Connection

    Private Sub Class_Terminate()
        On Error Resume Next
            pMyConnection.Close
        On Error GoTo 0
    End Sub
    
    Public Sub SetDatabaseName(ServerName As String, DefaultDB As String)
        pServerName = ServerName
        pDefaultDB = DefaultDB
    End Sub
    
    Public Function QueryDBTable(QueryString As String, Optional TransposeArr As Boolean = True) As Variant
        
        Dim MyCommand       As ADODB.Command
        Dim MyRecordSet     As ADODB.Recordset
        Dim MyData          As Variant
        
        Set MyCommand = New ADODB.Command
        
        With MyCommand
            Set .ActiveConnection = MyConn
            .CommandText = QueryString
            Set MyRecordSet = .Execute
        End With
        
        On Error Resume Next
            MyData = MyRecordSet.GetRows
            If Err.Number = 0 Then
                If TransposeArr Then
                    QueryDBTable = TransposeArray(MyData)
                Else
                    QueryDBTable = MyData
                End If
            End If
        On Error GoTo 0

    End Function
    
    Private Function TransposeArray(MyData As Variant) As Variant
        Dim x&, y As Long
        Dim OutputArr As Variant
        ReDim OutputArr(LBound(MyData, 2) To UBound(MyData, 2), LBound(MyData, 1) To UBound(MyData, 1))
        For x = LBound(MyData, 1) To UBound(MyData, 1)
            For y = LBound(MyData, 2) To UBound(MyData, 2)
                OutputArr(y, x) = MyData(x, y)
            Next y
        Next x
        TransposeArray = OutputArr
    End Function
    
    Private Property Get MyConn() As ADODB.Connection
    
        Select Case True
            Case pMyConnection Is Nothing, pMyConnection.State = ObjectStateEnum.adStateClosed
                
                Set pMyConnection = New ADODB.Connection
                
                With pMyConnection
                    .ConnectionString = ConnStr
                    .Open
                End With
                
        End Select
        
        Set MyConn = pMyConnection
        
    End Property

    Private Property Get ConnStr() As String
        
        Dim str As String
        
        If pServerName = "" Or pDefaultDB = "" Then _
            Debug.Print "Warning! Attempt to create connection string without providing server name and/or default database."
                
        str = _
            "Provider='@Provider';" & _
            "Data Source='@Server';" & _
            "Initial Catalog='@Database';" & _
            "Integrated Security='@Security'"
        
        str = Replace(str, "@Provider", CONN_PROVIDER)
        str = Replace(str, "@Server", pServerName)
        str = Replace(str, "@Database", pDefaultDB)
        str = Replace(str, "@Security", CONN_SECURITY)

        ConnStr = str

    End Property
    
    Public Property Get Database() As String:       Database = pDefaultDB:      End Property
    Public Property Get Server() As String:         Server = pServerName:       End Property


        
