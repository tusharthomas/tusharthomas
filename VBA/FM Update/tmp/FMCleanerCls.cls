VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FMCleanerCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private pFreshMemo As FreshMemoCls
    Private pRDC As RDCCls
    
    Private Sub Class_Initialize()
        Set pRDC = New RDCCls
    End Sub

    Public Sub Clean(FreshMemo As FreshMemoCls)
        
        Dim RDC As RDCCls
        
        Set pFreshMemo = FreshMemo
        
        DataRange.Interior.Color = vbWhite
            
        Call ResetCurrentRow
        Do Until CurrentRow > LastRow
        
            For Each RDC In RDCs
                Call CleanRDC(RDC)
            Next
            
            Call MoveToNextRow
        
        Loop
            
    End Sub
    
    Private Sub CleanRDC(RDC As RDCCls)
    
        Set pRDC = RDC
            
        With RDC
        
            .PreviousChange = ""
            
            'If the previous change was a change to available, then set the current value to nothing
            If .Orderable = OrderableVal(Available) Then .Orderable = ""
        
            If IsFloral Then
            
                If .Orderable = OrderableVal(NewlyOrderable) Then
                    .Orderable = ""
                    .TicketType = ""
                    .Comments = ""
                End If
                
                If IsPriceChange Then Call ClearOldPrices
                
            Else
            
                If IsPriceChange Or .TicketType = TicketTypeVal(STANDARD) Then
                    
                    .TicketType = ""
                    If .Comments <> CommentsVal(Delist) Then .Comments = ""
                    Call ClearOldPrices
                    
                End If
                
            End If

        End With
            
    End Sub
    
    Private Sub ClearOldPrices()
        
        Dim Zone As ZoneCls
        
        pRDC.OldPrice = ""
        
        For Each Zone In pRDC.Zones
            Zone.OldPrice = ""
        Next
        
    End Sub
    
    Private Function IsPriceChange() As Boolean
        Select Case pRDC.Comments
            Case CommentsVal(BacktoNormal), CommentsVal(PermanentPriceChange), CommentsVal(ReprintCorePriceTag)
                IsPriceChange = True
        End Select
    End Function
    
    
    'Inherited from FreshMemoCls
    Private Sub ResetCurrentRow():                  Call pFreshMemo.ResetCurrentRow:        End Sub
    Private Sub MoveToNextRow():                    Call pFreshMemo.MoveToNextRow:          End Sub
    
    Private Function IsFloral() As Boolean:         IsFloral = pFreshMemo.IsFloral:         End Function
    
    Private Property Get CurrentRow() As Integer:   CurrentRow = pFreshMemo.CurrentRow:     End Property
    Private Property Get LastRow() As Integer:      LastRow = pFreshMemo.LastRow:           End Property
    Private Property Get RDCs() As Collection:      Set RDCs = pFreshMemo.RDCs:             End Property
    Private Property Get DataRange() As Range:      Set DataRange = pFreshMemo.DataRange:   End Property

    'Inherited from RDCCls
    Public Function OrderableVal(Value As OrderableEnum) As String: OrderableVal = pRDC.OrderableVal(Value): End Function
    Public Function CommentsVal(Value As CommentEnum) As String: CommentsVal = pRDC.CommentsVal(Value): End Function
    Public Function TicketTypeVal(Value As TicketTypeEnum) As String: TicketTypeVal = pRDC.TicketTypeVal(Value): End Function
