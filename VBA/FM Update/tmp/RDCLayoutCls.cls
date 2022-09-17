VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RDCLayoutCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
    'Stores column numbers for RDC-block columns in the Fresh Memo
    
    Private pName                                                   As String
    Private pOld%, pNew%, pTic%, pVal%, pCom%, pOrd%, pPrevChg      As Integer  'Column location variables
    
    Private pFMLayout   As FMLayoutCls
    Private pRDC        As RDCCls
    
    'Method to set values
    
    Public Function SetValues(StartCol As Integer, FMLayout As FMLayoutCls, RDC As RDCCls) As Integer
    
        Dim x As Integer
        
        Set pFMLayout = FMLayout
        Set pRDC = RDC
    
        pName = ExtractName(ColHeader(StartCol)) 'Extract name from column header
        
        Call GetStartCols(StartCol)
        
        'Zone pricing columns
        x = pVal + 1
        Do Until InStr(ColHeader(x), "Not Orderable")
            If InStr(ColHeader(x), "Price OLD") Then Call AddZone(x)
            x = x + 1
        Loop
        
        'RDC block end, i.e. "Not orderable..." column
        pOrd = x
        
        pPrevChg = FindColumn(pName & " Change From Previous")
        
        SetValues = x + 1
        
    End Function
    
    Private Sub GetStartCols(StartCol As Integer)
        pOld = StartCol
        pNew = pOld + 1
        pTic = pNew + 1
        pCom = pTic + 1
        pVal = pCom + 1
    End Sub
    
    Private Sub AddZone(Col As Integer)
        Dim MyZone As ZoneCls: Set MyZone = New ZoneCls
        Call MyZone.SetValues(Col, pRDC)
        Call pRDC.Zones.Add(MyZone)
    End Sub
    
    Private Function FindColumn(HeaderName As String) As Integer
        FindColumn = pFMLayout.WS.Cells.Find(HeaderName).Column
    End Function
    
    Public Function ExtractName(HeaderName As String) As String
        Const NAME_LEN = 3
        ExtractName = Left(HeaderName, NAME_LEN)
    End Function
    
    Public Sub PrintCols()
    
        Dim x As Integer
        Dim Zone As ZoneCls
        
        Debug.Print ""
        Debug.Print "RDC Name:"; pName
        Debug.Print "RDC Old Price:"; Common.ColumnLetter(pOld)
        Debug.Print "RDC New Price:"; Common.ColumnLetter(pNew)
        Debug.Print "Ticket Type:"; Common.ColumnLetter(pTic)
        Debug.Print "Comments:"; Common.ColumnLetter(pCom)
        Debug.Print "Valid Until:"; Common.ColumnLetter(pVal)
        
        For Each Zone In pRDC.Zones
            Zone.PrintCols
        Next

        Debug.Print "Not Orderable column:"; Common.ColumnLetter(pOrd)
        Debug.Print "Previous Change column:"; Common.ColumnLetter(pPrevChg)
        Debug.Print ""
        
    End Sub
    
    'Get statements
    
    Public Property Get Name() As String: Name = pName: End Property
    Public Property Let Name(Value As String): pName = Value: End Property
    
    Public Property Get OldPriceCol() As Integer:          OldPriceCol = pOld:                  End Property
    Public Property Get NewPriceCol() As Integer:          NewPriceCol = pNew:                  End Property
    Public Property Get TicketTypeCol() As Integer:        TicketTypeCol = pTic:                End Property
    Public Property Get CommentsCol() As Integer:          CommentsCol = pCom:                  End Property
    Public Property Get ValidUntilCol() As Integer:        ValidUntilCol = pVal:                End Property
    Public Property Get OrderableCol() As Integer:         OrderableCol = pOrd:                 End Property
    Public Property Get PreviousChangeCol() As Integer:    PreviousChangeCol = pPrevChg:        End Property
    
    Public Property Get RdcOldPrice() As String:        RdcOldPrice = WS.Cells(CurrentRow, pOld).Value2:                End Property
    Public Property Get RdcNewPrice() As String:        RdcNewPrice = WS.Cells(CurrentRow, pNew).Value2:                End Property
    Public Property Get TicketType() As String:         TicketType = WS.Cells(CurrentRow, pTic).Value2:                 End Property
    Public Property Get Comments() As String:           Comments = WS.Cells(CurrentRow, pCom).Value2:                   End Property
    Public Property Get ValidUntil() As String:         ValidUntil = WS.Cells(CurrentRow, pVal).Value2:                 End Property
    Public Property Get Orderable() As String:          Orderable = UCase(WS.Cells(CurrentRow, pOrd).Value2):           End Property
    Public Property Get PreviousChange() As String:     PreviousChange = WS.Cells(CurrentRow, pPrevChg).Value2:         End Property
    Public Property Get OrderableRange() As Range: Set OrderableRange = WS.Cells(CurrentRow, pOrd): End Property
    
    Public Property Get ColHeader(ColNum As Integer) As String
        With pFMLayout
            ColHeader = Trim(CStr(.WS.Cells(HEADER_ROW, ColNum).Value2))
        End With
    End Property
    
    Public Property Get DataRange() As Range
        Set DataRange = _
            WS.Range( _
                WS.Cells(CurrentRow, pOld), _
                WS.Cells(CurrentRow, pOrd) _
                )
    End Property
    
    Public Property Get OldPriceToValidUntilRange() As Range
        Set OldPriceToValidUntilRange = _
            WS.Range( _
                WS.Cells(CurrentRow, pOld), _
                WS.Cells(CurrentRow, pVal) _
                )
    End Property
    
    'Let statements
    
    Public Property Let RdcOldPrice(Value As String):        WS.Cells(CurrentRow, pOld).Value2 = Value:                 End Property
    Public Property Let RdcNewPrice(Value As String):        WS.Cells(CurrentRow, pNew).Value2 = Value:                 End Property
    Public Property Let TicketType(Value As String):         WS.Cells(CurrentRow, pTic).Value2 = Value:                 End Property
    Public Property Let Comments(Value As String):           WS.Cells(CurrentRow, pCom).Value2 = Value:                 End Property
    Public Property Let ValidUntil(Value As String):         WS.Cells(CurrentRow, pVal).Value2 = Value:                 End Property
    Public Property Let Orderable(Value As String):          WS.Cells(CurrentRow, pOrd).Value2 = Value:                 End Property
    Public Property Let PreviousChange(Value As String):     WS.Cells(CurrentRow, pPrevChg).Value2 = Value:             End Property
    
    'Inherited from FMLayoutCls
    Private Property Get HEADER_ROW() As Integer:       HEADER_ROW = pFMLayout.HEADER_ROW:      End Property
    Public Property Get WS() As Worksheet:              Set WS = pFMLayout.WS:                  End Property
    Public Property Get CurrentRow() As Integer:       CurrentRow = pFMLayout.CurrentRow:      End Property
