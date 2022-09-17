VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChangeStringCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    'Builds summary string that is output into "Change From Previous..." column on right of Fresh Memo
    
    Private pChanges    As Variant
    Private pPos        As Integer
    
    Private pChangeData As ChangeDataCls
    Private pRDC        As RDCCls
    
    Public Enum ChangeCodeEnum
        IsRDCLevelChange = -1
        MarkForHighlights = 0
    End Enum
    
    Private Sub Class_Initialize()
        ReDim pChanges(0 To 9)
        pPos = LBound(pChanges)
    End Sub
    
    Public Sub OutputChangeSummaryString(RDC As RDCCls, ChangeData As ChangeDataCls)
    
        Dim Zone As ZoneCls
    
        Set pRDC = RDC
        Set pChangeData = ChangeData
        
        If ChangeLevel <> ChangeLevelValue(ZoneOnly) Then Call AddChange(IsRDCLevelChange)
        If MarkYellow <> "" Then Call AddChange(MarkForHighlights)
        
        For Each Zone In RDC.Zones
            If ZoneIsAffected(Zone) Then Call AddChange(Zone.ZoneNumber)
        Next
        
        pRDC.PreviousChange = BuildString()
    
    End Sub
    
    Private Function BuildString() As String
    
        Dim x           As Integer
        Dim MyString    As String
        
        Dim SortHelper As QuickSortCls: Set SortHelper = New QuickSortCls
        
        ReDim Preserve pChanges(LBound(pChanges) To pPos - 1)
                      
        pChanges = SortHelper.SortArrAsc(pChanges)
        
        MyString = Join(pChanges, ",")
        
        For x = CInt(IsRDCLevelChange) To CInt(MarkForHighlights)
            MyString = Replace(MyString, x, ChangeCode(x))
        Next x
        
        BuildString = MyString
        
    End Function
    
    Public Function ChangeCode(ByVal Value As ChangeCodeEnum) As String
        Select Case Value
            Case IsRDCLevelChange:      ChangeCode = "B"
            Case MarkForHighlights:     ChangeCode = "Y"
        End Select
    End Function

    Private Sub AddChange(Value As String)
        pChanges(pPos) = Value
        pPos = pPos + 1
    End Sub
    
    'Inherited from ChangeDataCls
    Private Sub GoToNextChange(): Call pChangeData.GoToNextChange: End Sub
    Private Sub MarkChangeCompleted(): Call pChangeData.MarkCompleted: End Sub
    
    Private Function HasMoreChanges() As Boolean: HasMoreChanges = pChangeData.HasMoreChanges: End Function
    Private Function ChangeLevelValue(Value As ChgLvlEnum) As String: ChangeLevelValue = pChangeData.ChangeLevelValue(Value): End Function
    Private Function RDCValue(Value As RDCEnum) As String: RDCValue = pChangeData.RDCValue(Value): End Function
    Private Function ZoneValue(Value As ZoneEnum) As String: ZoneValue = pChangeData.ZoneValue(Value): End Function
    Private Function RDCIsAffected(RDC As RDCCls) As Boolean: RDCIsAffected = pChangeData.RDCIsAffected(RDC): End Function
    Private Function ZoneIsAffected(Zone As ZoneCls) As Boolean: ZoneIsAffected = pChangeData.ZoneIsAffected(Zone): End Function

    Public Property Get SelectedDate() As String:   SelectedDate = pChangeData.SelectedDate:      End Property
    
    Private Property Get ItemNumber() As Long:               ItemNumber = pChangeData.ItemNumber:                 End Property
    Private Property Get ItemDescription() As String:        ItemDescription = pChangeData.ItemDescription:       End Property
    Private Property Get OSDate() As String:                 OSDate = pChangeData.OSDate:                         End Property
    Private Property Get CaseSize() As String:               CaseSize = pChangeData.CaseSize:                     End Property
    Private Property Get OldPrice() As String:               OldPrice = pChangeData.OldPrice:                     End Property
    Private Property Get NewPrice() As String:               NewPrice = pChangeData.NewPrice:                     End Property
    Private Property Get ValidFrom() As String:              ValidFrom = pChangeData.ValidFrom:                   End Property
    Private Property Get ValidTo() As String:                ValidTo = pChangeData.ValidTo:                       End Property
    Private Property Get TicketType() As String:             TicketType = pChangeData.TicketType:                 End Property
    Private Property Get ChangeType() As String:             ChangeType = pChangeData.ChangeType:                 End Property
    Private Property Get Comments() As String:               Comments = pChangeData.Comments:                     End Property
    Private Property Get ChangeLevel() As String:            ChangeLevel = pChangeData.ChangeLevel:               End Property
    Private Property Get RDCSelection() As String:           RDCSelection = pChangeData.RDCSelection:             End Property
    Private Property Get ZoneSelection() As String:          ZoneSelection = pChangeData.ZoneSelection:           End Property
    Private Property Get PlanogramCategory() As String:      PlanogramCategory = pChangeData.PlanogramCategory:   End Property
    Private Property Get ChillerNumber() As String:          ChillerNumber = pChangeData.ChillerNumber:           End Property
    Private Property Get MarkYellow() As String:             MarkYellow = pChangeData.MarkYellow:                 End Property
    Private Property Get BBD() As String:                    BBD = pChangeData.BBD:                               End Property
    Private Property Get PushedtoFM() As String:             PushedtoFM = pChangeData.PushedtoFM:                 End Property
