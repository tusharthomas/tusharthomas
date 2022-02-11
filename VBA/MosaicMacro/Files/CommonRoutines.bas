Attribute VB_Name = "CommonRoutines"
Option Explicit

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

