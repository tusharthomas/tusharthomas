VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FreshMemoCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Const path = "C:\Users\tusha\Documents\projects\MyGithub\tusharthomas\VBA\FM Update\FV Special Offer Master.xlsm"

    Private pWB         As Workbook
    Private pFMLayout   As FMLayoutCls
    Private pRDCs       As Collection
    
    Private pWBIsOpen   As Boolean
    
    Private Sub Class_Initialize()
        
        Set pRDCs = New Collection
        
        Set pWB = FileOps.GetWorkbook(path, pWBIsOpen, False)   'Retrieve workbook
        
        Set pFMLayout = New FMLayoutCls   'Get worksheet layout
        Call pFMLayout.GetWSLayout(Me)
        
    End Sub
    
    Public Sub ClearPreviousChanges()
        Dim Cleaner As FMCleanerCls
        Set Cleaner = New FMCleanerCls
        Call Cleaner.Clean(Me)
    End Sub
    
    Public Function IsSameDate(SelectedDate As String) As Boolean
        IsSameDate = (Trim(SelectedDate) = Trim(CStr(WS.Range(DATE_RANGE).Value2)))
    End Function
    
    Public Function IsFloral() As Boolean
        IsFloral = (Left(ItemNumber, 2) = "99" And Left(ItemNumber, 3) <> "995")
    End Function
    
    Public Function ItemIsOnFreshMemo(ItemNumber As String) As Boolean
        ItemIsOnFreshMemo = (Not IsError(Application.Match(CLng(ItemNumber), WS.Columns(ITEM_NUM_COL), 0)))
    End Function
    
    Public Sub GoToItemRow(ItemNumber As String)
        pFMLayout.Row = Application.Match(CLng(ItemNumber), WS.Columns(ITEM_NUM_COL), 0)
    End Sub
    
    Public Function InsertRowAtCategory(NewCategory As String, NewItemNumber As String) As Boolean
    
        'Attempts to insert row and returns TRUE if successful
    
        Dim InsertRow As Integer
        
        On Error Resume Next
            
            InsertRow = Application.Match(NewCategory, WS.Columns(PlanogramCol), 0)                'Match as string
            If InsertRow = 0 Then _
                InsertRow = Application.Match(CInt(NewCategory), WS.Columns(PlanogramCol), 0)      'Match as numeric
        
        On Error GoTo 0
        
        If InsertRow = 0 Then Exit Function
        
        Call WS.Rows(InsertRow).Insert
        
        pFMLayout.Row = InsertRow
        
        ItemNumber = NewItemNumber
        Category = NewCategory
        
        InsertRowAtCategory = True
    
    End Function
    
    Public Sub SaveChanges()
        Call pWB.Save
    End Sub
    
    Public Sub ReapplyHighlights()
        Dim HighlightHelper As FMHighlightsCls
        Set HighlightHelper = New FMHighlightsCls
        Call HighlightHelper.ApplyHighlights(Me)
    End Sub
    
    Private Sub Class_Terminate()
        On Error Resume Next
            If Not pWBIsOpen Then Call pWB.Close(False)
        On Error GoTo 0
    End Sub
    
    Public Property Get RDCs() As Collection:   Set RDCs = pRDCs:       End Property
    Public Property Get WB() As Workbook:       Set WB = pWB:           End Property

    'Inherited from FMLayoutCls
    Public Sub PrintCols(): Call pFMLayout.PrintCols: End Sub
    Public Sub MoveToNextRow(): Call pFMLayout.MoveToNextRow: End Sub
    Public Sub ResetCurrentRow(): Call pFMLayout.ResetCurrentRow: End Sub
    Public Sub ResetChillerNumbering(): Call pFMLayout.ResetChillerNumbering: End Sub
    
    Private Property Get BBDCol() As Integer:           BBDCol = pFMLayout.BBDCol:                      End Property
    Private Property Get PlanogramCol() As Integer:     PlanogramCol = pFMLayout.PlanogramCol:          End Property
    Private Property Get ExpressCol() As Integer:       ExpressCol = pFMLayout.ExpressCol:              End Property
    Private Property Get LimitedCol() As Integer:       LimitedCol = pFMLayout.LimitedCol:              End Property
    Private Property Get PlanogramRow() As Integer:     PlanogramRow = pFMLayout.PlanogramRow:          End Property
    Private Property Get DATE_RANGE() As String:        DATE_RANGE = pFMLayout.DATE_RANGE:              End Property
    Private Property Get ITEM_NUM_COL() As Integer:     ITEM_NUM_COL = pFMLayout.ITEM_NUM_COL:          End Property
        
    Public Property Get ItemNumber() As String:         ItemNumber = pFMLayout.ItemNumber:              End Property
    Public Property Get CurrentRow() As Integer:        CurrentRow = pFMLayout.CurrentRow:              End Property
    Public Property Get LastRow() As Integer:           LastRow = pFMLayout.LastRow:                    End Property
    Public Property Get DataRange() As Range:           Set DataRange = pFMLayout.DataRange:            End Property
    Public Property Get ItemInfoRange() As Range:       Set ItemInfoRange = pFMLayout.ItemInfoRange:    End Property

    Private Property Get WS() As Worksheet:             Set WS = pFMLayout.WS:                          End Property
    
    Private Property Let ItemNumber(Value As String):   pFMLayout.ItemNumber = Value:                   End Property
    Private Property Let Category(Value As String):     pFMLayout.Category = Value:                     End Property
    Public Property Let SelectedDate(Value As String):  pFMLayout.SelectedDate = Value:                 End Property
    Public Property Let CaseContent(Value As String):   pFMLayout.CaseContent = Value:                  End Property
    Public Property Let WriteOffDate(Value As String):  pFMLayout.WriteOffDate = Value:                 End Property
    Public Property Let BestByDate(Value As String):    pFMLayout.BestByDate = Value:                   End Property
    Public Property Let ItemDescription(Value As String):   pFMLayout.ItemDescription = Value: End Property
