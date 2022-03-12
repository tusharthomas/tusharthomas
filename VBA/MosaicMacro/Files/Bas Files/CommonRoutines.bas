Attribute VB_Name = "CommonRoutines"
Option Explicit

Public Const QUOTE_CHAR_CODE As Integer = 34

Public Sub ToggleApplicationSettings(TurnOn As Boolean, Optional ToggleCalculation As Boolean = True)
    With Application
        .EnableEvents = TurnOn
        .ScreenUpdating = TurnOn
        .DisplayAlerts = TurnOn
        If ToggleCalculation Then
            Application.Calculation = xlCalculationManual
            If TurnOn Then Application.Calculation = xlCalculationAutomatic
        End If
    End With
End Sub

Public Function GetRandomInteger(LowerBound As Integer, UpperBound As Integer) As Integer
    GetRandomInteger = CInt((UpperBound - LowerBound) * Rnd + LowerBound)
End Function

Public Function HasValues(ByVal MyText As String) As Boolean

    MyText = Replace(MyText, vbCr, "")
    MyText = Replace(MyText, vbCrLf, "")
    MyText = Replace(MyText, vbLf, "")
    MyText = Replace(MyText, " ", "")
    MyText = Trim(MyText)
    
    HasValues = (Len(MyText) > 0)

End Function

Public Function ConstructRGBString(Red As String, Green As String, Blue As String) As String

    Const RGB_TEMPLATE As String = "(@Red, @Green, @Blue)"
    Const RED_PLACEHOLDER As String = "@Red"
    Const GREEN_PLACEHOLDER As String = "@Green"
    Const BLUE_PLACEHOLDER As String = "@Blue"

    Dim RGBText As String: RGBText = RGB_TEMPLATE

    RGBText = Replace(RGBText, RED_PLACEHOLDER, Red)
    RGBText = Replace(RGBText, GREEN_PLACEHOLDER, Green)
    RGBText = Replace(RGBText, BLUE_PLACEHOLDER, Blue)
    
    ConstructRGBString = RGBText
    
End Function

Public Function LongToRGB(ColorCode As Long) As String
    
    'Source: https://stackoverflow.com/questions/24132665/return-rgb-values-from-range-interior-color-or-any-other-color-property
    
    Const BYTE_SIZE As Integer = 256
    
    Dim Red As Integer, Green As Integer, Blue As Integer
    
    Red = ColorCode Mod BYTE_SIZE
    Green = ColorCode \ BYTE_SIZE Mod BYTE_SIZE
    Blue = ColorCode \ Application.Power(BYTE_SIZE, 2) Mod BYTE_SIZE
    
    LongToRGB = ConstructRGBString(CStr(Red), CStr(Green), CStr(Blue))
    
End Function

Public Function RGBToLong(ByVal RGBText As String) As Long

    Dim Comma1Position As Integer
    Dim Comma2Position As Integer
    
    Dim Red As Integer, Green As Integer, Blue As Integer
    
    RGBText = Replace(RGBText, "(", "")
    RGBText = Replace(RGBText, ")", "")
    RGBText = Trim(RGBText)
    
    Comma1Position = InStr(1, RGBText, ",")
    Comma2Position = InStr(Comma1Position + 1, RGBText, ",")
    
    Red = Mid(RGBText, 1, Comma1Position - 1)
    Green = Mid(RGBText, Comma1Position + 1, Comma2Position - Comma1Position)
    Blue = Mid(RGBText, Comma2Position + 1, Len(RGBText) - Comma2Position)
    
    RGBToLong = RGB(Red, Green, Blue)
    
End Function

Public Function IsNumericAndValid(DataValue As String, Optional IsColorCode As Boolean = False) As Boolean
    IsNumericAndValid = True
    Select Case True
        Case Not IsNumeric(DataValue), CInt(DataValue) < 0, (IsColorCode And CInt(DataValue) > 255)
            IsNumericAndValid = False
    End Select
End Function
