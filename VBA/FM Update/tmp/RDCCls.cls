VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RDCCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
        
    Private pRDCLayout  As RDCLayoutCls
    Private pZoneCol    As Collection
    
    Public Enum OrderableEnum
        NotOrderable
        Available
        NewlyOrderable
    End Enum
    
    Public Enum CommentEnum
        BacktoNormal
        BacktoNormalInOut
        CaseSizeChange
        Delist
        NewCoreItem
        NewInOutItem
        NewItem
        OSExtension
        PermanentPriceChange
        ReprintCorePriceTag
        ReprintInOutPriceTag
        SpecialOffer
        SpecialOfferExtension
        SpecialOfferFurtherReduction
        TemporaryPriceChange
        NotAvailableinRDC
        DigitalCouponSpecialOffer
    End Enum
    
    Public Enum TicketTypeEnum
        LowerPrice
        MyLidl
        NA
        NewItemTicket
        STANDARD
        WhileSuppliesLast
    End Enum
     
    Private Sub Class_Initialize()
       Set pRDCLayout = New RDCLayoutCls
       Set pZoneCol = New Collection
    End Sub
    
    Public Function OrderableVal(Value As OrderableEnum) As String
        Select Case Value
            Case NotOrderable:      OrderableVal = "X"
            Case Available:         OrderableVal = "A"
            Case NewItem:           OrderableVal = "N"
        End Select
    End Function
    
    Public Function CommentsVal(Value As CommentEnum) As String
        Select Case Value
            Case BacktoNormal:                      CommentsVal = "Back to Normal"
            Case BacktoNormalInOut:                 CommentsVal = "Back to Normal (In/Out)"
            Case CaseSizeChange:                    CommentsVal = "Case Size Change"
            Case Delist:                            CommentsVal = "Delist"
            Case NewCoreItem:                       CommentsVal = "New Core Item"
            Case NewInOutItem:                      CommentsVal = "New In/Out Item"
            Case NewItem:                           CommentsVal = "New Item"
            Case OSExtension:                       CommentsVal = "OS Extension"
            Case PermanentPriceChange:              CommentsVal = "Permanent Price Change"
            Case ReprintCorePriceTag:               CommentsVal = "Reprint Core Price Tag"
            Case ReprintInOutPriceTag:              CommentsVal = "Reprint In/Out Price Tag"
            Case SpecialOffer:                      CommentsVal = "Special Offer"
            Case SpecialOfferExtension:             CommentsVal = "Special Offer Extension"
            Case SpecialOfferFurtherReduction:      CommentsVal = "Special Offer Further Reduction"
            Case TemporaryPriceChange:              CommentsVal = "Temporary Price Change"
            Case NotAvailableinRDC:                 CommentsVal = "Not Available in RDC"
            Case DigitalCouponSpecialOffer:         CommentsVal = "Digital Coupon/Special Offer"
        End Select
    End Function
    
    Public Function TicketTypeVal(Value As TicketTypeEnum) As String
        Select Case Value
            Case LowerPrice:            TicketTypeVal = "Lower Price"
            Case MyLidl:                TicketTypeVal = "MyLidl"
            Case NA:                    TicketTypeVal = "NA"
            Case NewItemTicket:         TicketTypeVal = "New"
            Case STANDARD:              TicketTypeVal = "Standard"
            Case WhileSuppliesLast:     TicketTypeVal = "While Supplies Last"
        End Select
    End Function
    
    Public Property Get Zones() As Collection:  Set Zones = pZoneCol:   End Property
    
    'Inherited from RDCLayoutCls
    Public Sub PrintCols(): Call pRDCLayout.PrintCols: End Sub
    
    Public Function SetValues(ColNum As Integer, FMLayout As FMLayoutCls) As Integer: SetValues = pRDCLayout.SetValues(ColNum, FMLayout, Me): End Function
    Public Function ExtractName(HeaderName As String) As String: ExtractName = pRDCLayout.ExtractName(HeaderName): End Function
    
    Public Property Get ColHeader(ColNum As Integer) As String: ColHeader = pRDCLayout.ColHeader(ColNum): End Property
    Public Property Get CurrentRow() As Integer: CurrentRow = pRDCLayout.CurrentRow: End Property
    Public Property Get WS() As Worksheet: Set WS = pRDCLayout.WS: End Property

    Public Property Get Name() As String:               Name = pRDCLayout.Name:                         End Property
    Public Property Get OldPrice() As String:           OldPrice = pRDCLayout.RdcOldPrice:               End Property
    Public Property Get NewPrice() As String:           NewPrice = pRDCLayout.RdcNewPrice:               End Property
    Public Property Get TicketType() As String:         TicketType = pRDCLayout.TicketType:              End Property
    Public Property Get Comments() As String:           Comments = pRDCLayout.Comments:                  End Property
    Public Property Get ValidUntil() As String:         ValidUntil = pRDCLayout.ValidUntil:              End Property
    Public Property Get Orderable() As String:          Orderable = pRDCLayout.Orderable:                End Property
    Public Property Get PreviousChange() As String:     PreviousChange = pRDCLayout.PreviousChange:      End Property
    Public Property Get DataRange() As Range:           Set DataRange = pRDCLayout.DataRange:            End Property
    Public Property Get OldPriceToValidUntilRange() As Range: Set OldPriceToValidUntilRange = pRDCLayout.OldPriceToValidUntilRange: End Property
    Public Property Get OrderableRange() As Range:      Set OrderableRange = pRDCLayout.OrderableRange:     End Property

    Public Property Let OldPrice(Value As String):          pRDCLayout.RdcOldPrice = Value:                 End Property
    Public Property Let NewPrice(Value As String):          pRDCLayout.RdcNewPrice = Value:                 End Property
    Public Property Let TicketType(Value As String):        pRDCLayout.TicketType = Value:                  End Property
    Public Property Let Comments(Value As String):          pRDCLayout.Comments = Value:                    End Property
    Public Property Let ValidUntil(Value As String):        pRDCLayout.ValidUntil = Value:                  End Property
    Public Property Let Orderable(Value As String):         pRDCLayout.Orderable = Value:                   End Property
    Public Property Let PreviousChange(Value As String):    pRDCLayout.PreviousChange = Value:              End Property
