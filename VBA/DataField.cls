VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private pDataSourceColumn As Integer
    Private pOutputColumn As Integer
    Private pClauseIndex As ClauseEnum
    Private pDataSource As DataSource
    Private pDataValue As Variant
    Private pName As String
    Private pIsDate As Boolean
    
    Private Sub Class_Initialize()
        pClauseIndex = None
    End Sub
    
    Public Property Get DataSourceColumn() As Integer: DataSourceColumn = pDataSourceColumn: End Property
    Public Property Get OutputColumn() As Integer: OutputColumn = pOutputColumn: End Property
    Public Property Get ClauseIndex() As ClauseEnum: ClauseIndex = pClauseIndex: End Property
    Public Property Get DataSource() As DataSource: Set DataSource = pDataSource: End Property
    Public Property Get DataValue() As Variant: DataValue = pDataValue: End Property
    Public Property Get Name() As String: Name = pName: End Property
    Public Property Get IsDate() As Boolean: IsDate = pIsDate: End Property

    Public Property Let DataSourceColumn(Value As Integer): pDataSourceColumn = Value: End Property
    Public Property Let OutputColumn(Value As Integer): pOutputColumn = Value: End Property
    Public Property Let ClauseIndex(Value As ClauseEnum): pClauseIndex = Value: End Property
    Public Property Set DataSource(Value As DataSource): Set pDataSource = Value: End Property
    Public Property Let DataValue(Value As Variant): pDataValue = Value: End Property
    Public Property Let Name(Value As String): pName = Value: End Property
    Public Property Let IsDate(Value As Boolean): pIsDate = Value: End Property
