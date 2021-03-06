Attribute VB_Name = "mdlMain"
 Option Explicit
 
    Dim NumberOfRows As Integer, NumberOfColumns As Integer
    
    Public Property Get MySettings() As AddInSettings
        Static Settings As AddInSettings
        If Settings Is Nothing Then
            Set Settings = New AddInSettings
        End If
        Set MySettings = Settings
    End Property

    Public Sub Main()
    
        If MsgBox(Prompt:=MAIN_INITIAL_PROMPT, Title:=MAIN_INITIAL_TITLE, Buttons:=vbYesNo) <> vbYes Then Exit Sub
        Call ToggleApplicationSettings(False)
        Call GetSettings
        
        Call FormatCanvas
        Call ResizeTiles
        Call ColorCanvas
        
        Call ToggleApplicationSettings(True)
        
    End Sub
    
    Sub FormatCanvas(Optional IsExternallyCalled As Boolean = False)
    
        'Creates grid with evenly-sized squares to fill with colors
    
    '0.0 Setup
        
        Const SQUARE_SIDE_LENGTH_IN_PIXELS As Integer = 64
        
        Const WIDTH_PIXELS_TO_POINTS_RATIO As Double = (64 / 5.18)
        Const HEIGHT_PIXELS_TO_POINTS_RATIO As Double = 2
        
        Const DARK_GRAY_50_PERCENT As LongLong = 7434613 'RGB(117, 113, 113)
        
        Dim x As Integer
        
        Dim MyWS As Worksheet: Set MyWS = ActiveWorkbook.ActiveSheet
        Dim MyBorders As Borders
        
        If IsExternallyCalled Then
            If MsgBox(Prompt:=FORMAT_INITIAL_PROMPT, Title:=FORMAT_INITIAL_TITLE, Buttons:=vbYesNo) <> vbYes Then GoTo EndOfSub
            Call GetSettings
            Call ToggleApplicationSettings(False)
        End If
        
    '1.0 Format canvas
        
        With MyWS
        
            .Cells.Clear
            
        '1.1 Dimension columns and rows
            
            For x = 1 To NumberOfRows
                .Cells(x, 1).RowHeight = SQUARE_SIDE_LENGTH_IN_PIXELS / HEIGHT_PIXELS_TO_POINTS_RATIO   'Convert pixels to points and set value
            Next x
            
            For x = 1 To NumberOfColumns
                .Cells(1, x).ColumnWidth = SQUARE_SIDE_LENGTH_IN_PIXELS / WIDTH_PIXELS_TO_POINTS_RATIO
            Next x
            
        '1.2 Set inactive cells to dark gray color
            
            .Range( _
                .Cells(1, NumberOfColumns + 1), _
                .Cells(1, .Columns.Count) _
                    ).EntireColumn.Interior.Color = DARK_GRAY_50_PERCENT
            
            .Range( _
                .Cells(NumberOfRows + 1, 1), _
                .Cells(.Rows.Count, 1) _
                    ).EntireRow.Interior.Color = DARK_GRAY_50_PERCENT
                    
        '1.3 Apply dark black bordering to all cells in active canvas area
                    
            Set MyBorders = _
                .Range( _
                    .Cells(1, 1), _
                    .Cells(NumberOfRows, NumberOfColumns) _
                        ).Borders
                
            With MyBorders
                .LineStyle = XlLineStyle.xlContinuous
                .Color = vbBlack
                .Weight = XlBorderWeight.xlThick
            End With
            
            .Cells(1, 1).Select
        
        End With
        
    '2.0 Tear down
    
        If IsExternallyCalled Then
            Call ToggleApplicationSettings(True)
        End If
    
EndOfSub:
        
        Set MyBorders = Nothing
        Set MyWS = Nothing
        
    End Sub
    
    Sub ColorCanvas(Optional IsExternallyCalled As Boolean = False)
    
    '0.0 Setup
    
        Dim x As Integer, y As Integer
        Dim Colors As Variant: Colors = MySettings.ColorsArray
        Dim RandomIndex As Integer
        Dim TryCounter As Integer
        
        Dim MyWS As Worksheet: Set MyWS = ActiveWorkbook.ActiveSheet
        
        For x = LBound(Colors) To UBound(Colors)
            Debug.Print Colors(x)
        Next x
    
        If IsExternallyCalled Then
            If MsgBox(Prompt:=COLOR_INITIAL_PROMPT, Title:=COLOR_INITIAL_TITLE, Buttons:=vbYesNo) <> vbYes Then GoTo EndOfSub
            Call GetSettings
            Call ToggleApplicationSettings(False)
        End If
        
    '1.0 Loop through tiles and apply coloring
        
        For y = 1 To NumberOfRows
            For x = 1 To NumberOfColumns
            
                If IsTopLeftCell(MyWS.Cells(y, x)) Then
                
                    TryCounter = 0
                    RandomIndex = GetRandomInteger(LBound(Colors), UBound(Colors))
                    Do Until IsValidColor(CLng(Colors(RandomIndex)), MyWS.Cells(y, x))
                    
                        RandomIndex = GetRandomInteger(LBound(Colors), UBound(Colors))
                        
                        TryCounter = TryCounter + 1
                        If TryCounter > UBound(Colors) * 100 Then
                            GoTo ApplyColor
'                                Call MsgBox(Prompt:=INFINITE_LOOP_PROMPT, Title:=INFINITE_LOOP_TITLE, Buttons:=vbExclamation)
'                                Stop
                        End If
                        
                    Loop
                    
ApplyColor:
                    MyWS.Cells(y, x).Interior.Color = Colors(RandomIndex)
                    
                
                End If
                
            Next x
        Next y
        
    '2.0 Tear down
    
        If IsExternallyCalled Then
            Call ToggleApplicationSettings(False)
        End If
        
EndOfSub:

        Set MyWS = Nothing
    
    End Sub
    
    Function IsTopLeftCell(MyRange As Range) As Boolean
        Dim IsMerged As Boolean: IsMerged = MyRange.MergeCells
        IsTopLeftCell = True
        If IsMerged And MyRange.Address <> MyRange.MergeArea.Cells(1, 1).Address Then IsTopLeftCell = False
    End Function
    
    Function IsValidColor(MyColor As Long, MyRange As Range) As Boolean
    
        'Loop through all perimeter cells to ensure that none match the color provided
        
        Dim TileRows As Integer: TileRows = MyRange.MergeArea.Rows.Count
        Dim TileColumns As Integer: TileColumns = MyRange.MergeArea.Columns.Count
        Dim Top As Integer: Top = MyRange.MergeArea.Cells(1, 1).Row
        Dim Left As Integer: Left = MyRange.MergeArea.Cells(1, 1).Column
        Dim Bottom As Integer: Bottom = Top + TileRows - 1
        Dim Right As Integer: Right = Left + TileColumns - 1
        
        Dim MyWS As Worksheet: Set MyWS = MyRange.Worksheet
        
        Dim x As Integer
        
        For x = Top To Bottom   'Loop through vertical perimeter cells
            If Left > 1 Then
                If MyWS.Cells(x, Left - 1).MergeArea.Cells(1, 1).Interior.Color = MyColor Then Exit Function
            End If
            If Right < NumberOfColumns Then
                If MyWS.Cells(x, Right + 1).MergeArea.Cells(1, 1).Interior.Color = MyColor Then Exit Function
            End If
        Next x
        
        For x = Left To Right    'Loop through horizontal perimeter cells
            If Top > 1 Then
                If MyWS.Cells(Top - 1, x).MergeArea.Cells(1, 1).Interior.Color = MyColor Then Exit Function
            End If
            If Bottom < NumberOfRows Then
                If MyWS.Cells(Bottom + 1, x).MergeArea.Cells(1, 1).Interior.Color = MyColor Then Exit Function
            End If
        Next x
        
        IsValidColor = True
    
    End Function
    
    Sub ResizeTiles(Optional IsExternallyCalled As Boolean = False)
    
    '0.0 Setup
    
        Const MAX_LENGTH As Integer = 5
        
        Dim x As Integer, y As Integer, i As Integer
        Dim Height As Integer, Width As Integer
        Dim IsMerged As Boolean
        
        Dim MyWS As Worksheet: Set MyWS = ActiveWorkbook.ActiveSheet
        Dim MyRange As Range
        
        If IsExternallyCalled Then
            If MsgBox(Prompt:=RESIZE_INITIAL_PROMPT, Title:=RESIZE_INITIAL_TITLE, Buttons:=vbYesNo) <> vbYes Then GoTo EndOfSub
            Call GetSettings
            Call ToggleApplicationSettings(False)
        End If
        
    '1.0 Loop through tiles
        
        For y = 1 To NumberOfRows
            For x = 1 To NumberOfColumns
            
'                Debug.Print MyWS.Cells(y, x).Address
                    
                If MyWS.Cells(y, x).MergeCells = True Then GoTo NextCell
                    
                Height = GetRandomInteger(0, MAX_LENGTH - 1)
                Width = GetRandomInteger(0, MAX_LENGTH - 1)
                
                i = 1
                Do Until i > Width
                    IsMerged = MyWS.Cells(y, x + i).MergeCells
                    If IsMerged Then Width = i - 1
                    i = i + 1
                Loop
                
                Do While x + Width > NumberOfColumns
                    Width = Width - 1
                Loop
                
                Do While y + Height > NumberOfRows
                    Height = Height - 1
                Loop
                
                With MyWS
                    Set MyRange = _
                        .Range( _
                            .Cells(y, x), _
                            .Cells(y + Height, x + Width) _
                                )
'                    Debug.Print MyRange.Address
                    MyRange.Select
                    Call MyRange.Merge
                End With
                
NextCell:
                
            Next x
        Next y
        
        If IsExternallyCalled Then
            Call ToggleApplicationSettings(False)
        End If
        
EndOfSub:
        
        Set MyRange = Nothing
        Set MyWS = Nothing
        
    End Sub
    
    Sub GetSettings()
        NumberOfRows = MySettings.NumberOfRows
        NumberOfColumns = MySettings.NumberOfColumns
    End Sub
    
    Sub Test()
        Dim f As FileSystemObject: Set f = New FileSystemObject
        Dim t As TextStream
        Dim MyText As String
        Set t = f.OpenTextFile("C:\Users\tusha\Documents\projects\MyGithub\tusharthomas\VBA\MosaicMacro\Settings.json", ForReading, False)
        MyText = t.ReadAll
        t.Close
        Stop
        Set f = Nothing
    End Sub
    
