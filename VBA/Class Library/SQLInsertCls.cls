VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLInsertCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private pTable$, pHeader$, pValues As String

    Private pSQLHelper As SQLCls
    
    Private Sub Class_Initialize()
        Set pSQLHelper = New SQLCls
    End Sub

    Public Sub CreateHeaderString(FieldsArr As Variant)
    
        On Error Resume Next
            Debug.Print UBound(FieldsArr, 2)
            If Err.Number = 0 Then
                Debug.Print "Warning! The array passed into CreateHeaderString has more than one dimension! No string created."
                Exit Sub
            End If
        On Error GoTo 0
    
        pHeader = "(" & Join(FieldsArr, ", ") & ")"
        If InStr(pHeader, ",,") Then _
            Debug.Print "Warning! Check SQL header string: "; pHeader
            
    End Sub
    
    Public Sub CreateHeaderStringFromRange(MyRange As Range)
        
        Dim RowArr As RowArrCls
        
        If MyRange.Rows.Count > 1 Then
            Debug.Print "Warning! Header range must have only one row. Header string not created for INSERT query."
            Exit Sub
        End If
        
        Set RowArr = New RowArrCls
        Call RowArr.GetFromRange(MyRange)
        Call CreateHeaderString(RowArr(RowArr.FirstRow))
        
    End Sub
    
    Public Sub SetDatabaseName(ServerName As String, DefaultDB As String)
        Call pSQLHelper.SetDatabaseName(ServerName, DefaultDB)
    End Sub
    
    Public Sub CreateValuesStringFrom1DArr(ValuesArr As Variant)
        
        On Error Resume Next
            Debug.Print UBound(ValuesArr, 2)
            If Err.Number = 0 Then
                Debug.Print "Warning! The array passed into CreateValuesStringFrom1DArr has more than one dimension! No string created."
                Exit Sub
            End If
        On Error GoTo 0
        
        pValues = pValues & "('" & Join(ValuesArr, "', '") & "')"       'Join array in one string
        pValues = Replace(pValues, ", '',", ", NULL,")                  'Replace empty strings with nulls
        
    End Sub
    
    Public Sub CreateValuesStringFrom2DArr(ValuesArr As Variant)
        
        Dim RowArr As RowArrCls
        
        On Error Resume Next
            Debug.Print UBound(ValuesArr, 2)
            If Err.Number > 0 Then
                Debug.Print "Warning! The array passed into CreateValuesStringFrom2DArr likely only has one dimension! No string created."
                Exit Sub
            End If
        On Error GoTo 0
        
        Set RowArr = New RowArrCls
        
        Call RowArr.GetFrom2DArr(ValuesArr)
        Call CreateValuesStringFromRowArr(RowArr)
        
    End Sub
    
    Public Sub CreateValuesStringFromRowArr(RowArr As RowArrCls)

        Dim Row As Variant

        For Each Row In RowArr.Data
            Call CreateValuesStringFrom1DArr(Row)
            pValues = pValues & "," & vbCrLf
        Next

        pValues = Left(pValues, Len(pValues) - 2)
        pValues = pValues & ";"

    End Sub
    
    Public Sub CreateValuesStringFromRange(MyRange As Range)
        
        Dim RowArr As RowArrCls
        Set RowArr = New RowArrCls
        Call RowArr.GetFromRange(MyRange)
        Call CreateValuesStringFromRowArr(RowArr)
        
    End Sub
    
    Public Sub ClearValuesString()
        pValues = ""
    End Sub
    
    Public Sub InsertRecords()
        
        If pTable = "" Or pHeader = "" Or pValues = "" Then
            Debug.Print "Warning! Attempt to insert records without providing SQL Table or Header & Values. No data inserted."
            Exit Sub
        End If
        
        With pSQLHelper
        
            If .Database = "" Or .Server = "" Then
                Debug.Print "Warning! Attempt to insert records without providing Database or Server information. No data inserted."
                Exit Sub
            End If
        
            Call .QueryDBTable(GetQuery())
        
        End With
        
    End Sub
    
    Private Function GetQuery() As String
        
        Dim str As String
    
        str = _
            "INSERT INTO @Table" & vbCrLf & _
            "@Header" & vbCrLf & _
            "VALUES" & vbCrLf & _
            "@Values"
            
        str = Replace(str, "@Table", pTable)
        str = Replace(str, "@Header", pHeader)
        str = Replace(str, "@Values", pValues)
        
        GetQuery = str
            
    End Function
    
    Public Sub PrintQuery()
        Debug.Print GetQuery()
    End Sub

    Public Property Get Header() As String:     Header = pHeader:               End Property
    Public Property Get Values() As String:     Values = pValues:               End Property
    Public Property Get TableName() As String:  TableName = pTable:             End Property
    Public Property Get Server() As String:     Server = pSQLHelper.Server:     End Property
    Public Property Get DBName() As String:     DBName = pSQLHelper.Database:   End Property
    
    Public Property Let TableName(Value As String):     pTable = Value:     End Property
