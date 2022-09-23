VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Public Sub EndRuntimeSafe()
        Call Common.ToggleApplicationSettings(True)
        DoEvents
        End
    End Sub

    Public Function ColumnLetter(ColumnNumber As Integer) As String
        ColumnLetter = Split(Cells(1, ColumnNumber).Address(True, False), "$")(0)
    End Function

    Public Function XlLastDataRow(WS As Worksheet, ColNum As Integer) As Long
        XlLastDataRow = CLng(WS.Cells(WS.Rows.Count, ColNum).End(xlUp).Row)
    End Function

    Public Sub ToggleApplicationSettings(TurnOn As Boolean)
        With Application
            .ScreenUpdating = TurnOn
            .DisplayAlerts = TurnOn
            .EnableEvents = TurnOn
        End With
    End Sub

    Public Function ExtractNumberFromText(TextToProcess As String) As Double
    'Description:_____________________________________________________________________________________________________________________________________________________________________
    '
    '   - Loop through all characters in the provided text, only keeping the first decimal and the remaining numeric characters
    '
    
        Dim Position As Integer
        Dim CharacterAtPosition As String, ExtractedNumber As String
        
        For Position = 1 To Len(TextToProcess)
        
            CharacterAtPosition = Mid(TextToProcess, Position, 1)
            
            If IsNumeric(CharacterAtPosition) = False Then  'Skip if there's already a decimal or if it's not a decimal
                If CharacterAtPosition <> "." Or InStr(1, ExtractedNumber, ".") > 0 Then GoTo NextPosition
                If CharacterAtPosition = "." And Not IsNumeric(Mid(TextToProcess, Position + 1, 1)) Then GoTo NextPosition
            End If
        
            ExtractedNumber = ExtractedNumber & CharacterAtPosition
        
NextPosition:
        
        Next
        
        If IsNumeric(ExtractedNumber) Then ExtractNumberFromText = CDbl(ExtractedNumber)
    
    End Function
