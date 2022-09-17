VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FMHighlightsCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Const HL_YELLOW = 65535
    Const HL_OLIVE = 9359529

    Private pFreshMemo As FreshMemoCls
    Private pChangeCodes As ChangeStringCls
    
    Private Sub Class_Initialize()
        Set pChangeCodes = New ChangeStringCls
    End Sub
    
    Public Sub ApplyHighlights(FreshMemo As FreshMemoCls)
    
        Dim ItemHasHighlights As Boolean
        Dim RDC As RDCCls
    
        Set pFreshMemo = FreshMemo
    
        Call ResetCurrentRow
        Do Until CurrentRow > LastRow
        
            ItemHasHighlights = False
            
            For Each RDC In RDCs
                If RDC.PreviousChange <> "" Then
                    
                    If Not ItemHasHighlights Then
                        Call HighlightItem
                        ItemHasHighlights = True
                    End If
                    
                    Call HighlightRDC(RDC)
                    
                End If
            Next
            
            Call MoveToNextRow
            
        Loop
    
    End Sub
    
    Private Sub HighlightItem()
        ItemInfoRange.Interior.Color = HL_OLIVE
    End Sub
    
    Private Sub HighlightRDC(RDC As RDCCls)
    
        Dim Zone As ZoneCls
        
        'Highlight entire RDC block
        
        If IsMarkedForYellowHighlights(RDC) Then
            RDC.DataRange.Interior.Color = HL_YELLOW
            Exit Sub
        ElseIf RDC.Orderable = RDC.OrderableVal(NewlyOrderable) Then
            RDC.DataRange.Interior.Color = HL_OLIVE
            Exit Sub
        End If
        
        'Highlight specific sections
        
        RDC.OrderableRange.Interior.Color = HL_OLIVE
        
        If HasRDCLevelChange(RDC) Then _
            RDC.OldPriceToValidUntilRange.Interior.Color = HL_OLIVE
        
        For Each Zone In RDC.Zones
            If ZoneHasChange(Zone, RDC) Then _
                Zone.DataRange.Interior.Color = HL_OLIVE
        Next
        
    End Sub
    
    Private Function IsMarkedForYellowHighlights(RDC As RDCCls) As Boolean
        IsMarkedForYellowHighlights = (InStr(RDC.PreviousChange, ChangeCode(MarkForHighlights)) > 0)
    End Function
    
    Private Function HasRDCLevelChange(RDC As RDCCls) As Boolean
        HasRDCLevelChange = (InStr(RDC.PreviousChange, ChangeCode(IsRDCLevelChange)) > 0)
    End Function
    
    Private Function ZoneHasChange(Zone As ZoneCls, RDC As RDCCls) As Boolean
        ZoneHasChange = (InStr(RDC.PreviousChange, Zone.ZoneNumber) > 0)
    End Function

    'Inherited from FreshMemoCls
    Private Sub ResetCurrentRow():                  Call pFreshMemo.ResetCurrentRow:        End Sub
    Private Sub MoveToNextRow():                    Call pFreshMemo.MoveToNextRow:          End Sub
    
    Private Function IsFloral() As Boolean:         IsFloral = pFreshMemo.IsFloral:         End Function
    
    Private Property Get CurrentRow() As Integer:   CurrentRow = pFreshMemo.CurrentRow:     End Property
    Private Property Get LastRow() As Integer:      LastRow = pFreshMemo.LastRow:           End Property
    Private Property Get RDCs() As Collection:      Set RDCs = pFreshMemo.RDCs:             End Property
    Private Property Get ItemInfoRange() As Range:  Set ItemInfoRange = pFreshMemo.ItemInfoRange:    End Property

    'Inherited from ChangeStringCls
    Private Function ChangeCode(ByVal Value As ChangeCodeEnum) As String: ChangeCode = pChangeCodes.ChangeCode(Value): End Function
