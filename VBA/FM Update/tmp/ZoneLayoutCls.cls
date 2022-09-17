VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZoneLayoutCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private pRDC                            As RDCCls

    Private pName                           As String
    Private pOldPriceCol%, pNewPriceCol     As Integer

    Public Sub SetValues(Col As Integer, RDC As RDCCls)
        Set pRDC = RDC
        pName = ExtractName(GetColHeader(Col))
        pOldPriceCol = Col
        pNewPriceCol = pOldPriceCol + 1
    End Sub
    
    Public Sub PrintCols()
        Debug.Print ""
        Debug.Print "Zone Name:"; pName
        Debug.Print "Zone Old Price:"; Common.ColumnLetter(pOldPriceCol)
        Debug.Print "Zone New Price:"; Common.ColumnLetter(pNewPriceCol)
        Debug.Print ""
    End Sub
    
    'Get statements
    Public Property Get ZoneNumber() As String: ZoneNumber = Mid(pName, 3, 1): End Property
    
    Public Property Get OldPriceCol() As Integer: OldPriceCol = pOldPriceCol: End Property
    Public Property Get NewPriceCol() As Integer: NewPriceCol = pNewPriceCol: End Property
    
    Public Property Get OldPrice() As String:   OldPrice = WS.Cells(CurrentRow, pOldPriceCol).Value2:   End Property
    Public Property Get NewPrice() As String:   NewPrice = WS.Cells(CurrentRow, pNewPriceCol).Value2:   End Property
    
    Public Property Get DataRange() As Range
        Set DataRange = _
            WS.Range( _
                WS.Cells(CurrentRow, pOldPriceCol), _
                WS.Cells(CurrentRow, pNewPriceCol) _
                )
    End Property
    
    'Let statements
    Public Property Let OldPrice(Value As String): WS.Cells(CurrentRow, pOldPriceCol).Value2 = Value:   End Property
    Public Property Let NewPrice(Value As String): WS.Cells(CurrentRow, pNewPriceCol).Value2 = Value:   End Property
    
    'Inherited from RDC
    Public Function ExtractName(HeaderName As String) As String:    ExtractName = pRDC.ExtractName(HeaderName):     End Function
    Public Function GetColHeader(Col As Integer) As String:         GetColHeader = pRDC.ColHeader(Col):             End Function
    
    Public Property Get CurrentRow() As Integer: CurrentRow = pRDC.CurrentRow: End Property
    Public Property Get WS() As Worksheet:  Set WS = pRDC.WS:   End Property
