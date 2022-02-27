VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsForm 
   Caption         =   "Settings"
   ClientHeight    =   10800
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6380
   OleObjectBlob   =   "SettingsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Sub ApplySettingsButton_Click()
        
        Dim SettingsText As String
        Dim TileAlgorithmCode As Integer
        
        Dim MyFileSystemObject As FileSystemObject: Set MyFileSystemObject = New FileSystemObject
        Dim MyTextStream As TextStream: Set MyTextStream = MyFileSystemObject.OpenTextFile(MySettings.FilePath, ForWriting, False)
        
        Dim ColorsDictionary As Dictionary: Set ColorsDictionary = CreateColorsDictionary
        
        If Not InputIsValid Then GoTo EndOfSub

        If Me.MergeOption.Value = True Then
            TileAlgorithmCode = MySettings.MERGE_CODE
        Else
            TileAlgorithmCode = MySettings.BISECT_CODE
        End If
        
        SettingsText = MySettings.CreateJSON(Me.NumberOfRowsTextbox.Value, Me.NumberOfColumnsTextbox.Value, Me.TileHeightWidthTextbox.Value, TileAlgorithmCode, ColorsDictionary)
        
        MyTextStream.Write (SettingsText)
        
        Call MsgBox(Prompt:=SETTINGS_APPLIED_PROMPT, Title:=SETTINGS_APPLIED_TITLE, Buttons:=vbOKOnly)
        
EndOfSub:
        
        MyTextStream.Close
        
        Set MyFileSystemObject = Nothing
        Set MyTextStream = Nothing
        Set ColorsDictionary = Nothing
        
    End Sub
    
    Private Function InputIsValid() As Boolean
        
        Const MIN_INPUT As Integer = 5
        Const MAX_INPUT As Integer = 250
        
        Dim MyTextbox As Control
        
        For Each MyTextbox In NumericTextboxes
            
            If Not IsNumeric(MyTextbox.Text) Then
                Call MsgBox(Prompt:=INVALID_INPUT_PROMPT, Title:=INVALID_INPUT_TITLE, Buttons:=vbOKOnly)
                GoTo InvalidInput
            End If
        
            If MyTextbox.Text < MIN_INPUT Then
                Call MsgBox(Prompt:=SMALL_INPUT_PROMPT, Title:=SMALL_INPUT_TITLE, Buttons:=vbOKOnly)
                GoTo InvalidInput
            End If
            
            If MyTextbox.Text > MAX_INPUT Then
                Call MsgBox(Prompt:=LARGE_INPUT_PROMPT, Title:=LARGE_INPUT_TITLE, Buttons:=vbOKOnly)
                GoTo InvalidInput
            End If
            
        Next
        
        InputIsValid = True
        Exit Function
        
InvalidInput:

        InputIsValid = False
        
    End Function
    
    Private Function CreateColorsDictionary() As Scripting.Dictionary
        Dim x As Integer
        Dim ColorName As String
        Dim ColorCode As Long
        Dim MyDictionary As Dictionary: Set MyDictionary = New Dictionary
        With Me.ColorsListBox
            For x = 1 To .ListCount
                ColorName = Replace(.List(x - 1, 0), Chr(QUOTE_CHAR_CODE), "")
                ColorCode = RGBToLong(.List(x - 1, 1))
                MyDictionary.Add ColorName, ColorCode
            Next x
        End With
        Set CreateColorsDictionary = MyDictionary
        Set MyDictionary = Nothing
    End Function

    Private Sub Userform_Initialize()
        Call RetrieveSettings
        Call RetrieveColors
    End Sub

    Private Sub RetrieveSettings()
    
        Application.EnableEvents = False
    
        With Me
        
            .NumberOfRowsTextbox.Value = MySettings.NumberOfRows
            .NumberOfColumnsTextbox.Value = MySettings.NumberOfColumns
            .TileHeightWidthTextbox.Value = MySettings.TileHeightWidthPixels
            
            Select Case MySettings.TileCreationAlgorithm
                Case MySettings.MERGE_CODE
                    Call SetMergeAlgorithm
                Case MySettings.BISECT_CODE
                    Call SetBisectAlgorithm
            End Select
            
        End With
        
        Application.EnableEvents = True
    
    End Sub
    
    Private Sub RetrieveColors()
    
        Dim x As Integer
        
        Dim ColorsDictionary As Scripting.Dictionary: Set ColorsDictionary = MySettings.ParseColorsJSON
        Dim AllKeys As Variant: AllKeys = ColorsDictionary.Keys
        Dim ColorName As String, ColorRGBCode As String
        Dim ColorCode As Long
        
        For x = LBound(AllKeys) To UBound(AllKeys)
            With Me.ColorsListBox
                ColorName = AllKeys(x)
                .AddItem ColorName
                ColorCode = ColorsDictionary.Item(ColorName)
                ColorRGBCode = LongToRGB(ColorCode)
                .List(x, 1) = ColorRGBCode
            End With
        Next x
        
        Set ColorsDictionary = Nothing
        
        Erase AllKeys
        
    End Sub
    
    Private Sub SetMergeAlgorithm()
        With Me
            .MergeOption.Value = True
            .BisectOption.Value = False
        End With
    End Sub
    
    Private Sub SetBisectAlgorithm()
        With Me
            .MergeOption.Value = False
            .BisectOption.Value = True
        End With
    End Sub
    
    Private Sub MergeOption_Click()
        Call SetMergeAlgorithm
    End Sub
    
    Private Sub BisectOption_Click()
        Call SetBisectAlgorithm
    End Sub
    
    Private Sub RemoveColorButton_Click()
    
        Const MINIMUM_COLORS_COUNT As Integer = 4
    
        Dim x As Integer: x = 0
    
        With Me.ColorsListBox
            
            If IsNull(.Value) Then Exit Sub
            If .ListCount = MINIMUM_COLORS_COUNT Then
                Call MsgBox(Prompt:=TOO_FEW_COLORS_PROMPT, Title:=TOO_FEW_COLORS_TITLE, Buttons:=vbExclamation)
                Exit Sub
            End If
            
            Do Until .List(x, 0) = .Value
                x = x + 1
                If x > .ListCount Then Exit Sub
            Loop
            
            .RemoveItem x
            
        End With
        
    End Sub
    
    Private Sub AddColorButton_Click()
    
    '0.0 Setup
    
        Dim RGBText As String
        Dim x As Integer
        
        Call ReplaceBlanks
        RGBText = ConstructRGBString(Me.RedTextbox.Value, Me.GreenTextbox.Value, Me.BlueTextbox.Value)
        
        With Me
        
    '1.0 Check if a name has been provided
            
            If HasValues(.NameTextbox.Value) = False Then
                Call MsgBox(Prompt:=NO_NAME_PROMPT, Title:=NO_NAME_TITLE, Buttons:=vbExclamation)
                Exit Sub
            End If
            
    '2.0 Check if the color is already in the list
            
            With .ColorsListBox
                
                For x = 1 To .ListCount
                    If .List(x - 1, 1) = RGBText Then
                        Call MsgBox(Prompt:=EXISTING_COLOR_PROMPT, Title:=EXISTING_COLOR_TITLE, Buttons:=vbOKOnly)
                        Exit Sub
                    End If
                Next x
                
    '3.0 If checks are passed, add to listbox
    
                .AddItem Chr(QUOTE_CHAR_CODE) & Me.NameTextbox.Value & Chr(QUOTE_CHAR_CODE)
                .List(.ListCount - 1, 1) = RGBText
                
            End With
            
        End With
        
    End Sub
    
    Private Sub ReplaceBlanks()
        
        Dim ColorTextbox As Control
        
        For Each ColorTextbox In ColorTextboxes
            If Trim(ColorTextbox.Text) = "" Then ColorTextbox.Text = 0
        Next
        
        Set ColorTextbox = Nothing
        
    End Sub
    
    Private Property Get ColorTextboxes() As Collection
        Dim MyCollection As Collection: Set MyCollection = New Collection
        With MyCollection
            .Add RedTextbox
            .Add GreenTextbox
            .Add BlueTextbox
        End With
        Set ColorTextboxes = MyCollection
        Set MyCollection = Nothing
    End Property

    Private Property Get NumericTextboxes() As Collection
        Dim MyCollection As Collection: Set MyCollection = New Collection
        With MyCollection
            .Add NumberOfRowsTextbox
            .Add NumberOfColumnsTextbox
            .Add TileHeightWidthTextbox
        End With
        Set NumericTextboxes = MyCollection
        Set MyCollection = Nothing
    End Property
