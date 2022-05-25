VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Public Enum DataSourceType
        Excel = 1
        Database = 2
        ListBox = 3
    End Enum

'=================================================================================================================================================================================================================================================================================================================================================
' Variables
'=================================================================================================================================================================================================================================================================================================================================================

    Private pSourceObject As Variant            'Worksheet, ListBox, DB Table Name String
    Private pSourceObjectType As DataSourceType

    Private pConnectionString As String         'Only needed for databases

    Private pLoadedData() As Variant
    Private pIsLoaded As Boolean

    Private pMatchData As DataSource            'External data
    Private pMatchColumnNumber As Integer       'Column in external data
    Private pIndexColumnNumber As Integer       'Column to index record in pLoadedData

    Private pRecordRow As Long
    Private pName As String
    
'=================================================================================================================================================================================================================================================================================================================================================
' Data Manipulation Routines
'=================================================================================================================================================================================================================================================================================================================================================

    Public Sub LoadData(Optional ForceLoad As Boolean = False)
        If pIsLoaded = True And ForceLoad = False Then Exit Sub
        Select Case pSourceObjectType
            Case DataSourceType.Excel
                pLoadedData = pSourceObject.DataTable
            Case DataSourceType.Database
                pLoadedData = QueryDBTable(pSourceObject, pConnectionString)
            Case DataSourceType.ListBox
                pLoadedData = ReadLBox(pSourceObject)
        End Select
        pIsLoaded = True
    End Sub
    
    Public Function GetRecordRow(Optional MatchValue As Variant = "") As Boolean
        'Returns True if record row was returned successfully
        Dim IndexColumnData As Variant: IndexColumnData = GetArrayColumn(pLoadedData, pIndexColumnNumber)
        If MatchValue = "" Then 'Try to get value to match and index; if can't be retrieved, exit subroutine
            With pMatchData
                MatchValue = .LoadedData(.RecordRow, pMatchColumnNumber)
            End With
            If MatchValue = "" Then Exit Sub
        End If
        On Error Resume Next
            pRecordRow = 0
            pRecordRow = Application.Match(Trim(CStr(MatchValue)), IndexColumnData, 0)
        On Error GoTo 0
        GetRecordRow =  (pRecordRow <> 0)
    End Sub
    
'=================================================================================================================================================================================================================================================================================================================================================
' Get Statements
'=================================================================================================================================================================================================================================================================================================================================================

    Public Property Get SourceObject() As ListObject: Set pSourceObject = pSourceObject: End Property
    Public Property Get SourceObjectType As DataSourceType: SourceObjectType = pSourceObjectType: End Property
    Public Property Get ConnectionString As String: ConnectionString = pConnectionString: End Property
    Public Property Get LoadedData() As Variant: LoadedData = pLoadedData: End Property
    Public Property Get IsLoaded() As Boolean: IsLoaded = pIsLoaded: End Property
    Public Property Get RecordRow() As Long: RecordRow = pRecordRow: End Property
    Public Property Get MatchData() As DataSource: Set MatchData = pMatchData: End Property
    Public Property Get MatchColumnNumber() As Integer: MatchColumnNumber = pMatchColumnNumber: End Property
    Public Property Get IndexColumnNumber() As Integer: IndexColumnNumber = pIndexColumnNumber: End Property
    Public Property Get Name() As String: Name = pName: End Property
    
'=================================================================================================================================================================================================================================================================================================================================================
' Let Statements
'=================================================================================================================================================================================================================================================================================================================================================
    
    Public Property Let SourceObject(Vaue As String): pSourceObject = Value: End Property 'For DB Table Name
    Public Property Let RecordRow(Value As Long): pRecordRow = Value: End Property
    Public Property Let SourceObjectType(Value As DataSourceType): pSourceObjectType = Value: End Property
    Public Property Let ConnectionString(Value As String): pConnectionString = Value: End Property
    Public Property Let MatchColumnNumber(Value As Integer): pMatchColumnNumber = Value: End Property
    Public Property Let IndexColumnNumber(Value As Integer): pIndexColumnNumber = Value: End Property
    Public Property Let Name(Value As String): pName = Value: End Property
    
'=================================================================================================================================================================================================================================================================================================================================================
' Set Statements
'=================================================================================================================================================================================================================================================================================================================================================
    
    Public Property Set SourceObject(Value As ListObject): Set pSourceObject = Value: End Property 'For WS or Listbox
    Public Property Set MatchData(Value As DataSource): Set pMatchData = Value: End Property
    
