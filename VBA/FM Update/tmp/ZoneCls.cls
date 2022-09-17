VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZoneCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private pZoneLayout As ZoneLayoutCls
    
    Private Sub Class_Initialize()
        Set pZoneLayout = New ZoneLayoutCls
    End Sub
    
    'Inherited from ZoneLayoutCls
    Public Sub SetValues(ColNum As Integer, RDC As RDCCls): Call pZoneLayout.SetValues(ColNum, RDC): End Sub
    Public Sub PrintCols(): Call pZoneLayout.PrintCols: End Sub
    
    Public Property Get ZoneNumber() As String: ZoneNumber = pZoneLayout.ZoneNumber: End Property
    Public Property Get OldPrice() As String:   OldPrice = pZoneLayout.OldPrice:   End Property
    Public Property Get NewPrice() As String:   NewPrice = pZoneLayout.NewPrice:   End Property
    Public Property Get DataRange() As Range:   Set DataRange = pZoneLayout.DataRange:   End Property
    
    Public Property Let OldPrice(Value As String): pZoneLayout.OldPrice = Value:   End Property
    Public Property Let NewPrice(Value As String): pZoneLayout.NewPrice = Value:   End Property
