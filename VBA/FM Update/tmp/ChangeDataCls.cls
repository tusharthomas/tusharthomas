VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChangeDataCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    'Tracks each change to be pushed to the fresh memo

    Const WS_NAME = "FM Changes"
    
    Const pITEM_NUMBER_COL = 1  'Column positions
    Const pITEM_DESCRIPTION_COL = 2
    Const pOS_DATE_COL = 3
    Const pCASE_SIZE_COL = 4
    Const pOLD_PRICE_COL = 5
    Const pNEW_PRICE_COL = 6
    Const pVALID_FROM_COL = 7
    Const pVALID_TO_COL = 8
    Const pTICKET_TYPE_COL = 9
    Const pCHANGE_TYPE_COL = 10
    Const pCOMMENTS_COL = 11
    Const pCHANGE_LEVEL_COL = 12
    Const pRDC_COL = 13
    Const pZONE_COL = 14
    Const pPLANOGRAM_CATEGORY_COL = 15
    Const pCHILLER_NUMBER_COL = 16
    Const pMARK_YELLOW_COL = 17
    Const pBBD_COL = 18
    Const pPUSHED_TO_FM_COL = 19
    
    Const pSTART_ROW = 2
    Const pSTART_COL = 1
    Const pLAST_COL = pPUSHED_TO_FM_COL
    
    Private pMyWS           As Worksheet
    Private pChgData        As Variant  'Data array
    Private pCurrChg        As Integer  'pointer to current row
    Private pLastRow        As Integer
    
    Public Enum ChgLvlEnum
        RDCOnly
        ZoneOnly
        RDCAndZone
    End Enum
    
    Public Enum RDCEnum
        AllRDCs
        PYE
        FRG
        GRM
        PYEFRG
        PYEGRM
        FRGGRM
    End Enum
    
    Public Enum ZoneEnum
        AllZones
        Zone1
        Zone2
        Zone8
        Z1And2
        Z1And8
        Z2And8
        NoZones
    End Enum
    
    Private Sub Class_Initialize()
        Set pMyWS = ThisWorkbook.Worksheets(WS_NAME)
        Call ReadDataFromWS
    End Sub
    
    Private Sub ReadDataFromWS()
        
        pLastRow = Common.XlLastDataRow(pMyWS, 1)
        
        If pLastRow < pSTART_ROW Then
            MsgBox "There is no data on the " & pMyWS.Name & " worksheet! No changes pushed."
            Call Common.ToggleApplicationSettings(True)
            End
        End If
        
        With pMyWS
            pChgData = _
                .Range( _
                    .Cells(pSTART_ROW, pSTART_COL), _
                    .Cells(pLastRow, pLAST_COL) _
                        ).Value2
        End With
        
        pCurrChg = LBound(pChgData, 1)
        
    End Sub
    
    Public Sub GoToNextChange()
        pCurrChg = pCurrChg + 1
    End Sub
    
    Public Sub MarkCompleted()
        pMyWS.Cells(pCurrChg + pSTART_ROW - 1, pPUSHED_TO_FM_COL).Value2 = "X"
    End Sub
    
    Public Sub ClearCompletionMarkers()
        pMyWS.Range( _
            pMyWS.Cells(pSTART_ROW, pPUSHED_TO_FM_COL), _
            pMyWS.Cells(pLastRow, pPUSHED_TO_FM_COL) _
                ).Value2 = ""
    End Sub
    
    Public Function ChangeLevelValue(Value As ChgLvlEnum) As String
        Select Case Value
            Case RDCAndZone:    ChangeLevelValue = "RDC(s) and Zone(s)"
            Case RDCOnly:       ChangeLevelValue = "RDC(s) only"
            Case ZoneOnly:      ChangeLevelValue = "Zone(s) only"
        End Select
    End Function
    
    Public Function RDCValue(Value As RDCEnum) As String
        Select Case Value
            Case AllRDCs:   RDCValue = "ALL"
            Case PYE:       RDCValue = "PYE"
            Case FRG:       RDCValue = "FRG"
            Case GRM:       RDCValue = "GRM"
            Case PYEFRG:    RDCValue = "PYE, FRG"
            Case PYEGRM:    RDCValue = "PYE, GRM"
        End Select
    End Function
    
    Public Function ZoneValue(Value As ZoneEnum) As String
        Select Case Value
            Case AllZones:  ZoneValue = "ALL"
            Case Zone1:     ZoneValue = "1"
            Case Zone2:     ZoneValue = "2"
            Case Zone8:     ZoneValue = "8"
            Case Z1And2:    ZoneValue = "1, 2"
            Case Z1And8:    ZoneValue = "1, 8"
            Case Z2And8:    ZoneValue = "2, 8"
            Case NoZones:   ZoneValue = "NONE"
        End Select
    End Function
    
    Public Function RDCIsAffected(RDC As RDCCls) As Boolean
        RDCIsAffected = (RDCSelection = RDCValue(AllRDCs) Or InStr(RDCSelection, RDC.Name))
    End Function
    
    Public Function ZoneIsAffected(Zone As ZoneCls) As Boolean
        ZoneIsAffected = (ZoneSelection = ZoneValue(AllZones) Or InStr(ZoneSelection, Zone.ZoneNumber))
    End Function
    
    Public Property Get HasMoreChanges() As Boolean
        HasMoreChanges = (pCurrChg <= UBound(pChgData, 1))
    End Property
    
    Public Property Get SelectedDate() As String
        SelectedDate = ThisWorkbook.Worksheets("Dashboard").Range("Q23").Value2
    End Property
    
    Public Property Get ItemNumber() As Long:               ItemNumber = CLng(pChgData(pCurrChg, pITEM_NUMBER_COL)):            End Property
    Public Property Get ItemDescription() As String:        ItemDescription = pChgData(pCurrChg, pITEM_DESCRIPTION_COL):        End Property
    Public Property Get OSDate() As String:                 OSDate = pChgData(pCurrChg, pOS_DATE_COL):                          End Property
    Public Property Get CaseSize() As String:               CaseSize = pChgData(pCurrChg, pCASE_SIZE_COL):                      End Property
    Public Property Get OldPrice() As String:               OldPrice = pChgData(pCurrChg, pOLD_PRICE_COL):                      End Property
    Public Property Get NewPrice() As String:               NewPrice = pChgData(pCurrChg, pNEW_PRICE_COL):                      End Property
    Public Property Get ValidFrom() As String:              ValidFrom = pChgData(pCurrChg, pVALID_FROM_COL):                    End Property
    Public Property Get ValidTo() As String:                ValidTo = pChgData(pCurrChg, pVALID_TO_COL):                        End Property
    Public Property Get TicketType() As String:             TicketType = pChgData(pCurrChg, pTICKET_TYPE_COL):                  End Property
    Public Property Get ChangeType() As String:             ChangeType = pChgData(pCurrChg, pCHANGE_TYPE_COL):                  End Property
    Public Property Get Comments() As String:               Comments = pChgData(pCurrChg, pCOMMENTS_COL):                       End Property
    Public Property Get ChangeLevel() As String:            ChangeLevel = pChgData(pCurrChg, pCHANGE_LEVEL_COL):                End Property
    Public Property Get RDCSelection() As String:           RDCSelection = pChgData(pCurrChg, pRDC_COL):                        End Property
    Public Property Get ZoneSelection() As String:          ZoneSelection = pChgData(pCurrChg, pZONE_COL):                      End Property
    Public Property Get PlanogramCategory() As String:      PlanogramCategory = pChgData(pCurrChg, pPLANOGRAM_CATEGORY_COL):    End Property
    Public Property Get ChillerNumber() As String:          ChillerNumber = pChgData(pCurrChg, pCHILLER_NUMBER_COL):            End Property
    Public Property Get MarkYellow() As String:             MarkYellow = pChgData(pCurrChg, pMARK_YELLOW_COL):                  End Property
    Public Property Get BBD() As String:                    BBD = pChgData(pCurrChg, pBBD_COL):                                 End Property
    Public Property Get PushedtoFM() As String:             PushedtoFM = pChgData(pCurrChg, pPUSHED_TO_FM_COL):                 End Property
