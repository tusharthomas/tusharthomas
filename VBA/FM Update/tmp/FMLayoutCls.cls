VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FMLayoutCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    'This class reads RDC/Zone information dynamically from the Fresh Memo

    Const pWS_NAME = "US"
    Const pHEADER_ROW = 3
    Const pDATE_RANGE = "C1"
    
    Const pITEM_NUM_COL = 1
    Const pITEM_DESC_COL = 2
    Const pCASE_CONTENT_COL = 3
    
    Private pBBD%, pPlanogram%, pExpress%, pLimited     As Integer  'Column position variables
    Private pPlanogramRow%, pLastRow                    As Integer
    Private pCurrRow                                    As Integer
    
    Private pMyWS       As Worksheet
    Private pFM         As FreshMemoCls
    
    Public Sub GetWSLayout(FM As FreshMemoCls)
    
        Dim Col As Integer
        Dim RDCInfo As RDCCls
        
        Set pFM = FM
        Call SetWorksheet(FM.WB)
        
        'Get fixed location columns
        Call FindColsByHeader
        
        'Get RDC columns
        Col = pBBD + 1
        Do Until Col >= pPlanogram
            Set RDCInfo = New RDCCls
            Col = RDCInfo.SetValues(Col, Me)
            Call FM.RDCs.Add(RDCInfo)
        Loop
        
        Call ResetCurrentRow
        
    End Sub
    
    Private Sub FindColsByHeader()
    
        With pMyWS.Cells
        
            pBBD = .Find("BBD").Column
            pExpress = .Find("Express Store Change From Previous").Column
            pLimited = .Find("Limited Assortment Store Change From Previous").Column
            
            With .Find("Planogram")
                pPlanogram = .Column
                pPlanogramRow = .Row
            End With
            
        End With
        
    End Sub
    
    Private Sub SetWorksheet(WB As Workbook)
        On Error Resume Next
            Set pMyWS = WB.Worksheets(pWS_NAME)
            If Err.Number > 0 Then
                MsgBox "Error: Could not find a worksheet called " & pWS_NAME & " in the workbook " & WB.Name & ".", vbCritical, "Worksheet Not Found"
                Exit Sub
            End If
            pLastRow = Common.XlLastDataRow(pMyWS, 1)
        On Error GoTo 0
    End Sub
    
    Public Sub ResetCurrentRow()
        pCurrRow = pPlanogramRow + 1
    End Sub
    
    Public Sub MoveToNextRow()
        pCurrRow = pCurrRow + 1
    End Sub
    
    Public Sub ResetChillerNumbering()
        
        ResetCurrentRow
        
        Do Until Not IsNumeric(PlanogramCategory)
            PlanogramCategory = pCurrRow - pPlanogramRow
            MoveToNextRow
        Loop
        
    End Sub
    
    Public Sub PrintCols()
    
        'For debugging
    
        Dim RDC As RDCCls
    
        Debug.Print "BBD column:"; Common.ColumnLetter(pBBD)
        Debug.Print "Number of RDC's: "; pFM.RDCs.Count
    
        For Each RDC In pFM.RDCs
            RDC.PrintCols
        Next
    
        Debug.Print "Planogram column:"; Common.ColumnLetter(pPlanogram)
        Debug.Print "Planogram row:"; Common.ColumnLetter(pPlanogramRow)
        Debug.Print "Limited Assortment column:"; Common.ColumnLetter(pLimited)
        Debug.Print "Express store column:"; Common.ColumnLetter(pExpress)
    
    End Sub
    
    'Get statements
    Private Property Get PlanogramCategory() As String: PlanogramCategory = pMyWS.Cells(pCurrRow, pPlanogram).Value2: End Property
    Public Property Get HEADER_ROW() As Integer:        HEADER_ROW = pHEADER_ROW:       End Property
    
    Public Property Get BBDCol() As Integer:            BBDCol = pBBD:                  End Property
    Public Property Get PlanogramCol() As Integer:      PlanogramCol = pPlanogram:      End Property
    Public Property Get ExpressCol() As Integer:        ExpressCol = pExpress:          End Property
    Public Property Get LimitedCol() As Integer:        LimitedCol = pLimited:          End Property
    Public Property Get PlanogramRow() As Integer:      PlanogramRow = pPlanogramRow:   End Property
    Public Property Get LastRow() As Integer:           LastRow = pLastRow:             End Property
    Public Property Get CurrentRow() As Integer:        CurrentRow = pCurrRow:          End Property
    Public Property Get ITEM_NUM_COL() As Integer:      ITEM_NUM_COL = pITEM_NUM_COL:   End Property
    
    Public Property Get WS() As Worksheet:              Set WS = pMyWS:                 End Property
    
    Public Property Get DATE_RANGE() As String:         DATE_RANGE = pDATE_RANGE:       End Property
    
    Public Property Get ItemNumber() As String:         ItemNumber = pMyWS.Cells(pCurrRow, pITEM_NUM_COL).Value2:      End Property
    
    Public Property Get DataRange() As Range
        Set DataRange = _
            pMyWS.Range( _
                pMyWS.Cells(pPlanogramRow + 1, 1), _
                pMyWS.Cells(pLastRow, pPlanogram - 1) _
                    )
    End Property
    
    Public Property Get ItemInfoRange() As Range
        Set ItemInfoRange = _
            pMyWS.Range( _
                pMyWS.Cells(pCurrRow, pITEM_NUM_COL), _
                pMyWS.Cells(pCurrRow, pCASE_CONTENT_COL) _
                    )
    End Property

    'Let statements
    Private Property Let PlanogramCategory(Value As String): pMyWS.Cells(pCurrRow, pPlanogram).Value2 = Value: End Property
    Public Property Let ItemNumber(Value As String):    pMyWS.Cells(pCurrRow, pITEM_NUM_COL).Value2 = Value: End Property
    Public Property Let Category(Value As String):      pMyWS.Cells(pCurrRow, pPlanogram).Value2 = Value: End Property
    Public Property Let SelectedDate(Value As String):  pMyWS.Range(pDATE_RANGE).Value2 = CStr(CDate(Value)):   End Property
    Public Property Let CaseContent(Value As String):   pMyWS.Cells(pCurrRow, pCASE_CONTENT_COL).Value2 = Value: End Property
    Public Property Let WriteOffDate(Value As String):  CaseContent = Value: End Property
    Public Property Let BestByDate(Value As String):    pMyWS.Cells(pCurrRow, pBBD).Value2 = Value: End Property
    Public Property Let ItemDescription(Value As String):   pMyWS.Cells(pCurrRow, pITEM_DESC_COL).Value2 = Value: End Property
    
    Public Property Let Row(Value As Integer):          pCurrRow = Value:                       End Property
