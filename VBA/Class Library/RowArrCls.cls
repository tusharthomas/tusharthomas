VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RowArrCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    'Setup: Variables, Constants, and Initialization
        
    Const ROW_INDEX = 1
    Const COL_INDEX = 2
    Const UPPER = 1000000
        
    Private pRows       As Variant
    Private pColInfo    As Variant
    Private pLastRow    As Long
    
    Private Sub Class_Initialize()
        ReDim pRows(0 To UPPER)
        pLastRow = -1
        ReDim pColInfo(0 To 0)
    End Sub

    'Routines Exposed to Caller
    
    Public Sub CreateNewArr(NumberOfColumns As Long)
        ReDim pColInfo(NumberOfColumns - 1)
    End Sub
        
    Public Sub GetFrom2DArr(Arr As Variant)
        
        Dim Row&, Col   As Long
        Dim RowArr      As Variant
        
        Call SetNumRowsAndCols(Arr)
        
        For Row = FirstRow To LastRow
        
            ReDim RowArr(FirstColumn To LastColumn)
            
            For Col = FirstColumn To LastColumn
                RowArr(Col) = Arr(Row, Col)
            Next Col
            
            pRows(Row) = RowArr
            
        Next Row
        
    End Sub
    
    Public Sub GetFromRange(MyRange As Range)
        With MyRange
            If Not IsArray(.Value2) Then
                Debug.Print "Warning! No array created from range " & .Address & " on " & .Parent.Name & " because range reference was scalar."
            Else
                Call GetFrom2DArr(.Value2)
            End If
        End With
    End Sub
    
    Public Function WriteTo2DArr() As Variant
        
        Dim Row&, Col   As Long
        Dim OutputArr   As Variant
        
        ReDim OutputArr(0 To RowCount - 1, 0 To ColumnCount - 1)
        
        For Row = FirstRow To LastRow
            For Col = FirstColumn To LastColumn
                OutputArr(Row, Col) = pRows(Row)(Col)
            Next Col
        Next Row
        
        WriteTo2DArr = OutputArr
        
    End Function
    
    Public Sub WriteToRange(MyRange As Range)
        
        With MyRange
            If .Rows.Count = 1 And .Columns.Count = 1 Then
                Set MyRange = _
                    .Parent.Range( _
                        MyRange, _
                        .Offset(RowCount - .Row, ColumnCount - .Column) _
                    )
            End If
        End With
        
        MyRange.Value2 = WriteTo2DArr()
        
    End Sub
    
    'Row Operations
    
    Public Sub InsertRow(ArrToInsert As Variant, Optional InsertPosition As Long = -1)
    
        Dim Row As Long
    
        If LBound(ArrToInsert) <> FirstColumn Or UBound(ArrToInsert) <> LastColumn Then
            Debug.Print "Warning! Row inserted with call to InsertRow did not match the column count of the existing data." & vbCrLf & _
                        "The row has been resized. Data may have been lost."
            ReDim Preserve ArrToInsert(FirstColumn To LastColumn)
        End If
        
        If InsertPosition = -1 Then InsertPosition = LastRow + 1
        
        pLastRow = LastRow + 1
        
        For Row = LastRow To InsertPosition + 1 Step -1
            pRows(Row) = pRows(Row - 1)
        Next
        
        pRows(InsertPosition) = ArrToInsert
    
    End Sub

    Public Sub DeleteRow(Optional RowToDelete As Long)
    
        Dim Row As Long
        
        If IsMissing(RowToDelete) Then RowToDelete = LastRow
        
        For Row = RowToDelete To LastRow - 1
            pRows(Row) = pRows(Row + 1)
        Next Row
        
        pLastRow = pLastRow - 1
        
    End Sub

    Public Sub SwapRows(Row1 As Long, Row2 As Long)
        Dim OldRow As Variant
        OldRow = pRows(Row1)
        pRows(Row1) = pRows(Row2)
        pRows(Row2) = OldRow
    End Sub
    
    'Column Operations
    
    Public Sub InsertColumn(ArrToInsert As Variant, Optional InsertPosition As Long = -1)

        Dim Row&, Col   As Long
        Dim RowArr      As Variant
        
        If LBound(ArrToInsert) <> FirstRow Or UBound(ArrToInsert) <> LastRow Then
            Debug.Print "Warning! Column inserted with call to InsertCol did not match the row count of the existing data." & vbCrLf & _
                        "The column has been resized. Data may have been lost."
            ReDim Preserve ArrToInsert(FirstRow To LastRow)
        End If
        
        If InsertPosition = -1 Then InsertPosition = LastColumn + 1

        For Row = FirstRow To LastRow

            RowArr = pRows(Row)

            ReDim Preserve RowArr(FirstColumn To LastColumn + 1)

            For Col = LastColumn + 1 To InsertPosition + 1 Step -1
                RowArr(Col) = RowArr(Col - 1)
            Next Col

            RowArr(InsertPosition) = ArrToInsert(Row)
            pRows(Row) = RowArr

        Next Row

        Call UpdateColInfo

    End Sub

    Public Sub DeleteColumn(Optional ColumnToDelete As Long)
    
        Dim Row&, Col   As Long
        Dim RowArr      As Variant
        
        If IsMissing(ColumnToDelete) Then ColumnToDelete = LastColumn
        
        For Row = FirstRow To LastRow

            RowArr = pRows(Row)

            For Col = ColumnToDelete To LastColumn - 1
                RowArr(Col) = RowArr(Col + 1)
            Next Col

            ReDim Preserve RowArr(FirstColumn To LastColumn - 1)

            pRows(Row) = RowArr
            
        Next Row

        Call UpdateColInfo
    
    End Sub
    
    Public Sub SwapColumns(Col1 As Long, Col2 As Long)

        Dim Row         As Long
        Dim OldVal      As Variant

        For Row = FirstRow To LastRow
            OldVal = pRows(Row)(Col1)
            pRows(Row)(Col1) = pRows(Row)(Col2)
            pRows(Row)(Col2) = OldVal
        Next Row

    End Sub

    'Helper routines

    Private Sub SetNumRowsAndCols(Arr As Variant)
        ReDim pRows(LBound(Arr, ROW_INDEX) To UBound(Arr, ROW_INDEX))
        ReDim pColInfo(LBound(Arr, COL_INDEX) To UBound(Arr, COL_INDEX))
        pLastRow = UBound(Arr, ROW_INDEX)
    End Sub

    Private Sub UpdateColInfo()
        ReDim pColInfo(LBound(pRows(FirstRow)) To UBound(pRows(FirstRow)))
    End Sub

    'Properties
    
    Public Property Get Data() As Variant 'Default property
Attribute Data.VB_UserMemId = 0
        Data = pRows
    End Property

    Public Property Get RowCount() As Long:             RowCount = LastRow - FirstRow + 1:              End Property
    Public Property Get ColumnCount() As Long:          ColumnCount = LastColumn - FirstColumn + 1:     End Property
    Public Property Get FirstRow() As Long:             FirstRow = LBound(pRows):                       End Property
    Public Property Get LastRow() As Long:              LastRow = pLastRow:                             End Property
    Public Property Get FirstColumn() As Long:          FirstColumn = LBound(pColInfo):                 End Property
    Public Property Get LastColumn() As Long:           LastColumn = UBound(pColInfo):                  End Property
    Public Property Let NewValue(Row&, Column&, Val):   pRows(Row)(Column) = Val:                       End Property
