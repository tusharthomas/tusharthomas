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

'=================================================================================================================================================================================================================================================================================================================================================
' Variables
'=================================================================================================================================================================================================================================================================================================================================================

    Private pSourceTable As ListObject
    Private pMyData() As Variant
    Private pIsLoaded As Boolean
    Private pMatchData As DataSource          'External data
    Private pMatchColumnNumber As Integer     'Column in external data
    Private pIndexColumnNumber As Integer     'Column to match to
    Private pRecordRow As Long
    Private pName As String
    
'=================================================================================================================================================================================================================================================================================================================================================
' Data Manipulation Routines
'=================================================================================================================================================================================================================================================================================================================================================

    Public Sub LoadData(Optional ForceLoad As Boolean = False)
        If pIsLoaded = True And ForceLoad = False Then Exit Sub
        pIsLoaded = True
        pMyData = pSourceTable.DataBodyRange.Value2
    End Sub
    
    Public Sub GetRecordRow(Optional MatchValue As Variant = "")
        Dim IndexColumnData As Variant: IndexColumnData = GetArrayColumn(pMyData, pIndexColumnNumber)
        If MatchValue = "" Then MatchValue = pMatchData.MyData(pMatchData.RecordRow, pMatchColumnNumber)
        On Error Resume Next
            pRecordRow = 0
            pRecordRow = Application.Match(Trim(CStr(MatchValue)), IndexColumnData, 0)
        On Error GoTo 0
    End Sub
    
'=================================================================================================================================================================================================================================================================================================================================================
' Get Statements
'=================================================================================================================================================================================================================================================================================================================================================

    Public Property Get SourceTable() As ListObject: Set SourceTable = pSourceTable: End Property
    Public Property Get MyData() As Variant: MyData = pMyData: End Property
    Public Property Get IsLoaded() As Boolean: IsLoaded = pIsLoaded: End Property
    Public Property Get RecordRow() As Long: RecordRow = pRecordRow: End Property
    Public Property Get MatchData() As DataSource: Set MatchData = pMatchData: End Property
    Public Property Get MatchColumnNumber() As Integer: MatchColumnNumber = pMatchColumnNumber: End Property
    Public Property Get IndexColumnNumber() As Integer: IndexColumnNumber = pIndexColumnNumber: End Property
    Public Property Get Name() As String: Name = pName: End Property
    
'=================================================================================================================================================================================================================================================================================================================================================
' Let Statements
'=================================================================================================================================================================================================================================================================================================================================================
    
    Public Property Let RecordRow(Value As Long): pRecordRow = Value: End Property
    Public Property Let MatchColumnNumber(Value As Integer): pMatchColumnNumber = Value: End Property
    Public Property Let IndexColumnNumber(Value As Integer): pIndexColumnNumber = Value: End Property
    Public Property Let Name(Value As String): pName = Value: End Property
    
'=================================================================================================================================================================================================================================================================================================================================================
' Set Statements
'=================================================================================================================================================================================================================================================================================================================================================
    
    Public Property Set SourceTable(Value As ListObject): Set pSourceTable = Value: End Property
    Public Property Set MatchData(Value As DataSource): Set pMatchData = Value: End Property
    
