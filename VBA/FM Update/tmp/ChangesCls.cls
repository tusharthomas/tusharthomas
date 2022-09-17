VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChangesCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Dim pErrorList As String
    Dim pRDC As RDCCls
    
    Dim pData As ChangeDataCls
    Dim pFreshMemo As FreshMemoCls
    
    Private Sub Class_Initialize()
        Set pData = New ChangeDataCls
        Set pRDC = New RDCCls
    End Sub
    
    Public Sub PushChanges(FreshMemo As FreshMemoCls)
        
        Set pFreshMemo = FreshMemo
        
        Call ClearCompletionMarkers
        
        Do
            Call PushChange
            Call pData.GoToNextChange
        Loop While HasMoreChanges
        
        Call ReportErrors
        
    End Sub
    
    Private Sub PushChange()
    
        Dim MyRDC As RDCCls
        
        If Not ItemIsOnFreshMemo(ItemNumber) Then
            If AddItem = False Then Exit Sub
        Else
            Call pFreshMemo.GoToItemRow(ItemNumber)
        End If
        
        If IsFloral And OSDate <> "" Then pFreshMemo.WriteOffDate = OSDate
        If BBD <> "" Then pFreshMemo.BestByDate = BBD
        If ItemDescription <> "" Then pFreshMemo.ItemDescription = ItemDescription
        If CaseSize <> "" Then pFreshMemo.CaseContent = CaseSize
        
        For Each MyRDC In pFreshMemo.RDCs
            If RDCIsAffected(MyRDC) Then Call PushChangeToRDC(MyRDC)
        Next
        
        Call MarkChangeCompleted
        
    End Sub
    
    Private Function AddItem() As Boolean
    
        Dim MatchValue As String
        
        MatchValue = PlanogramCategory
        If Trim(UCase(PlanogramCategory)) = "CHILLER" Then MatchValue = ChillerNumber
    
        If pFreshMemo.InsertRowAtCategory(MatchValue, ItemNumber) = False Then
            pErrorList = pErrorList & ItemNumber
        Else
            AddItem = True
        End If
    
    End Function
    
    Private Sub PushChangeToRDC(MyRDC As RDCCls)
        
        Call PushPrices(MyRDC)
        Call PushOtherChanges(MyRDC)
        Call BuildChangeSummaryString(MyRDC)
    
    End Sub
    
    Private Sub PushPrices(MyRDC As RDCCls)
    
        Dim Zone As ZoneCls
        
        If ChangeLevel <> ChangeLevelValue(ZoneOnly) Then   'If there is an RDC change
            MyRDC.OldPrice = OldPrice
            MyRDC.NewPrice = NewPrice
        End If
        
        If ChangeLevel <> ChangeLevelValue(RDCOnly) Then
            
            For Each Zone In MyRDC.Zones
                If ZoneIsAffected(Zone) Then
                    Zone.OldPrice = OldPrice
                    Zone.NewPrice = NewPrice
                End If
            Next
            
        End If
        
    End Sub
    
    Private Sub PushOtherChanges(MyRDC As RDCCls)
    
        MyRDC.Comments = ChangeType

        MyRDC.TicketType = TicketType
        MyRDC.ValidUntil = ValidTo

        If InStr(ChangeType, "New") Then MyRDC.Orderable = OrderableVal(NewlyOrderable)
        
        Select Case ChangeType
            
            Case CommentsVal(SpecialOfferFurtherReduction), CommentsVal(SpecialOfferExtension)
                MyRDC.Orderable = ""
                
            Case CommentsVal(Delist)
                MyRDC.TicketType = ""
                MyRDC.ValidUntil = ""
            
        End Select
        
    End Sub

    Private Sub BuildChangeSummaryString(MyRDC As RDCCls)
        Dim ChangeSummary As ChangeStringCls
        Set ChangeSummary = New ChangeStringCls
        Call ChangeSummary.OutputChangeSummaryString(MyRDC, pData)
    End Sub
    
    Private Sub ReportErrors()
        If pErrorList <> "" Then _
            Call MsgBox("Error! The below items did not push to the Fresh Memo: " & pErrorList, vbExclamation, "Error List")
    End Sub
    
    'Inherited from ChangeDataCls
    Private Sub GoToNextChange(): Call pData.GoToNextChange: End Sub
    Private Sub ClearCompletionMarkers(): Call pData.ClearCompletionMarkers: End Sub
    Private Sub MarkChangeCompleted(): Call pData.MarkCompleted: End Sub
    
    Private Function HasMoreChanges() As Boolean: HasMoreChanges = pData.HasMoreChanges: End Function
    Private Function ChangeLevelValue(Value As ChgLvlEnum) As String: ChangeLevelValue = pData.ChangeLevelValue(Value): End Function
    Private Function RDCValue(Value As RDCEnum) As String: RDCValue = pData.RDCValue(Value): End Function
    Private Function ZoneValue(Value As ZoneEnum) As String: ZoneValue = pData.ZoneValue(Value): End Function
    Private Function RDCIsAffected(RDC As RDCCls) As Boolean: RDCIsAffected = pData.RDCIsAffected(RDC): End Function
    Private Function ZoneIsAffected(Zone As ZoneCls) As Boolean: ZoneIsAffected = pData.ZoneIsAffected(Zone): End Function
    
    Public Property Get SelectedDate() As String:   SelectedDate = pData.SelectedDate:      End Property
    
    Private Property Get ItemNumber() As Long:               ItemNumber = pData.ItemNumber:                 End Property
    Private Property Get ItemDescription() As String:        ItemDescription = pData.ItemDescription:       End Property
    Private Property Get OSDate() As String:                 OSDate = pData.OSDate:                         End Property
    Private Property Get CaseSize() As String:               CaseSize = pData.CaseSize:                     End Property
    Private Property Get OldPrice() As String:               OldPrice = pData.OldPrice:                     End Property
    Private Property Get NewPrice() As String:               NewPrice = pData.NewPrice:                     End Property
    Private Property Get ValidFrom() As String:              ValidFrom = pData.ValidFrom:                   End Property
    Private Property Get ValidTo() As String:                ValidTo = pData.ValidTo:                       End Property
    Private Property Get TicketType() As String:             TicketType = pData.TicketType:                 End Property
    Private Property Get ChangeType() As String:             ChangeType = pData.ChangeType:                 End Property
    Private Property Get Comments() As String:               Comments = pData.Comments:                     End Property
    Private Property Get ChangeLevel() As String:            ChangeLevel = pData.ChangeLevel:               End Property
    Private Property Get RDCSelection() As String:           RDCSelection = pData.RDCSelection:             End Property
    Private Property Get ZoneSelection() As String:          ZoneSelection = pData.ZoneSelection:           End Property
    Private Property Get PlanogramCategory() As String:      PlanogramCategory = pData.PlanogramCategory:   End Property
    Private Property Get ChillerNumber() As String:          ChillerNumber = pData.ChillerNumber:           End Property
    Private Property Get MarkYellow() As String:             MarkYellow = pData.MarkYellow:                 End Property
    Private Property Get BBD() As String:                    BBD = pData.BBD:                               End Property
    Private Property Get PushedtoFM() As String:             PushedtoFM = pData.PushedtoFM:                 End Property
    
    'Inherited from FreshMemoCls
    Private Function ItemIsOnFreshMemo(ItemNumber As String) As Boolean: ItemIsOnFreshMemo = pFreshMemo.ItemIsOnFreshMemo(ItemNumber): End Function
    Private Function IsFloral() As Boolean: IsFloral = pFreshMemo.IsFloral: End Function

    'Inherited from RDCCls
    Private Function CommentsVal(Value As CommentEnum) As String: CommentsVal = pRDC.CommentsVal(Value): End Function
    Private Function OrderableVal(Value As OrderableEnum) As String: OrderableVal = pRDC.OrderableVal(Value): End Function
