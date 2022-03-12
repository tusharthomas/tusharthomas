Attribute VB_Name = "a_Messages"
Public Const MAIN_INITIAL_PROMPT As String = "Are you sure you want to create a mosaic on the current worksheet?" & vbCr & vbCr & "This will clear out any current information in the sheet."
Public Const MAIN_INITIAL_TITLE As String = "Are you sure?"

Public Const FORMAT_INITIAL_PROMPT As String = "Are you sure you want to format your current worksheet as a mosaic canvas?" & vbCr & vbCr & "This will clear out any current information in the sheet."
Public Const FORMAT_INITIAL_TITLE As String = "Are you sure?"

Public Const COLOR_INITIAL_PROMPT As String = "Are you sure you want to color all tiles in your canvas? This action cannot be undone."
Public Const COLOR_INITIAL_TITLE As String = "Are you sure?"

Public Const RESIZE_INITIAL_PROMPT As String = "Are you sure you want to randomly resize all tiles in your canvas? This action cannot be undone."
Public Const RESIZE_INITIAL_TITLE As String = "Are you sure?"

Public Const INFINITE_LOOP_PROMPT As String = "Infinite Loop!"
Public Const INFINITE_LOOP_TITLE As String = "Infinite Loop!"

Public Const TOO_FEW_COLORS_PROMPT As String = "Oops, you can't have fewer colors. Add more colors before removing the selected one."
Public Const TOO_FEW_COLORS_TITLE As String = "Too few colors!"

Public Const NO_NAME_PROMPT As String = "Oops, seems you forgot to provide a name for your color."
Public Const NO_NAME_TITLE As String = "No name!"

Public Const EXISTING_COLOR_PROMPT As String = "It looks like this color is already included in the mosaic."
Public Const EXISTING_COLOR_TITLE As String = "Color already included."

Public Const SETTINGS_APPLIED_PROMPT As String = "Successfully saved settings!"
Public Const SETTINGS_APPLIED_TITLE As String = "Settings saved"

Public Const INVALID_INPUT_PROMPT As String = "Oops, looks like one of the inputs is not numeric or formatted incorrectly. Please check the input and try again."
Public Const INVALID_INPUT_TITLE As String = "Invalid data type"

Public Const SMALL_INPUT_PROMPT As String = "Oops, looks like one of the inputs is too large. You can't input a dimension smaller than 5."
Public Const SMALL_INPUT_TITLE As String = "Input too small"

Public Const LARGE_INPUT_PROMPT As String = "Oops, looks like one of the inputs is too large. You can't input a dimension larger than 250."
Public Const LARGE_INPUT_TITLE As String = "Input too large"
