Attribute VB_Name = "a_CallBacks"
Option Explicit

Private Const MAIN_ID As String = "PaintButton"
Private Const CREATE_GRID_ID As String = "CreateButton"
Private Const RESIZE_ID As String = "ResizeButton"
Private Const COLOR_ID As String = "ColorButton"
Private Const SETTINGS_ID As String = "SettingsButton"

Public Sub ProcessCallbacks(Control As IRibbonControl)
    Select Case Control.ID
        Case MAIN_ID
            Call Main
        Case CREATE_GRID_ID
            Call FormatCanvas(True)
        Case RESIZE_ID
            Call ResizeTiles(True)
        Case COLOR_ID
            Call ColorCanvas(True)
        Case SETTINGS_ID
            SettingsForm.Show
    End Select
End Sub
