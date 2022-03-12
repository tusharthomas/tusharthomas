VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AddInSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private Const FILE_NAME As String = "Settings.json"

    Private Sub Class_Initialize()
        If Dir(FilePath, vbNormal) = "" Then Call CreateFile
    End Sub

    Public Property Get FilePath() As String
        FilePath = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "") & FILE_NAME
    End Property
    
    Public Property Get SettingsTextStream(Mode As IOMode) As TextStream
        Dim MyPath As String: MyPath = FilePath
        Dim MyFileSystemObject As FileSystemObject: Set MyFileSystemObject = New FileSystemObject
        If Dir(MyPath, vbNormal) = "" Then Call CreateFile
        Set SettingsTextStream = MyFileSystemObject.OpenTextFile(MyPath, Mode, False)
        Set MyFileSystemObject = Nothing
    End Property
    
    Public Property Get SettingsText() As String
    
        Dim MyTextStream As TextStream: Set MyTextStream = SettingsTextStream(ForReading)
        
        SettingsText = MyTextStream.ReadAll
        
        MyTextStream.Close
        
        Set MyTextStream = Nothing
        
    End Property
    
    Private Sub CreateFile()
        
        Dim MyFileSystemObject As FileSystemObject: Set MyFileSystemObject = New FileSystemObject
        Dim MyTextStream As TextStream: Set MyTextStream = MyFileSystemObject.OpenTextFile(FilePath, ForAppending, True)
        
        MyTextStream.Write (DefaultJSON)
        MyTextStream.Close
        
        Set MyTextStream = Nothing
        Set MyFileSystemObject = Nothing
        
    End Sub

    Private Property Get DefaultJSON() As String
    
        Const NUMBER_OF_ROWS_DEFAULT As Integer = 128
        Const NUMBER_OF_COLUMNS_DEFAULT As Integer = 128
        Const TILE_DIMENSIONS_DEFAULT As Integer = 64
        Const ALGORITHM_CODE_DEFAULT As Integer = 1
        
        Dim ColorsDictionary As Scripting.Dictionary: Set ColorsDictionary = New Scripting.Dictionary
        
        With ColorsDictionary
            .Add "Green, Accent 6, Lighter 40%", 9359539
            .Add "Gold, Accent 4", 49407
            .Add "Blue, Accent 1", 12874308
            .Add "Light Blue", 16763183
        End With
        
        DefaultJSON = CreateJSON(NUMBER_OF_ROWS_DEFAULT, NUMBER_OF_COLUMNS_DEFAULT, TILE_DIMENSIONS_DEFAULT, ALGORITHM_CODE_DEFAULT, ColorsDictionary)
        
        Set ColorsDictionary = Nothing
        
    End Property
    
    Public Function CreateJSON(NumberOfRows As Integer, NumberOfColumns As Integer, TileDimension As Integer, AlgorithmCode As Integer, ColorsDictionary As Scripting.Dictionary) As String
    
        Const NUMBER_OF_ROWS_PLACEHOLDER As String = "@NumberOfRows"
        Const NUMBER_OF_COLUMNS_PLACEHOLDER As String = "@NumberOfColumns"
        Const TILE_DIMENSION_PLACEHOLDER As String = "@TileDimension"
        Const ALGORITHM_PLACEHOLDER As String = "@AlgorithmCode"
        Const COLORS_PLACEHOLDER As String = "@Colors"
    
        Dim JSONText As String
        
        JSONText = JSONText & "{" & vbCr
        JSONText = JSONText & "    ""Number of Rows"" : @NumberOfRows," & vbCr
        JSONText = JSONText & "    ""Number of Columns"" : @NumberOfColumns," & vbCr
        JSONText = JSONText & "    ""Tile Height/Width (px.)"" : @TileDimension," & vbCr
        JSONText = JSONText & "    ""Tile Creation Algorithm"" : @AlgorithmCode," & vbCr
        JSONText = JSONText & "    ""Colors"" : [" & vbCr
        JSONText = JSONText & "@Colors" & vbCr
        JSONText = JSONText & "    ]" & vbCr
        JSONText = JSONText & "}"
        
        JSONText = Replace(JSONText, NUMBER_OF_ROWS_PLACEHOLDER, NumberOfRows)
        JSONText = Replace(JSONText, NUMBER_OF_COLUMNS_PLACEHOLDER, NumberOfColumns)
        JSONText = Replace(JSONText, TILE_DIMENSION_PLACEHOLDER, TileDimension)
        JSONText = Replace(JSONText, ALGORITHM_PLACEHOLDER, AlgorithmCode)
        JSONText = Replace(JSONText, COLORS_PLACEHOLDER, ParseColorsDictionary(ColorsDictionary))
        
        CreateJSON = JSONText
        
    End Function
    
    Private Function ParseColorsDictionary(ColorsDictionary As Scripting.Dictionary) As String
    
        'Loops through provided dictionary and creates JSON-formatted text from data
    
        Const NAME_PLACEHOLDER As String = "@Name"
        Const CODE_PLACEHOLDER As String = "@Code"
        
        Dim x As Integer
        Dim ParsedString As String
        Dim MyKey As String
        Dim AllKeys As Variant: AllKeys = ColorsDictionary.Keys
        
        For x = LBound(AllKeys) To UBound(AllKeys)
        
            MyKey = AllKeys(x)
            
            ParsedString = ParsedString & "        {" & vbCr
            ParsedString = ParsedString & "            ""Name"" : ""@Name""," & vbCr
            ParsedString = ParsedString & "            ""Code"" : @Code" & vbCr
            ParsedString = ParsedString & "        }," & vbCr
            
            ParsedString = Replace(ParsedString, NAME_PLACEHOLDER, MyKey)
            ParsedString = Replace(ParsedString, CODE_PLACEHOLDER, ColorsDictionary.Item(MyKey))
            
        Next x
        
        ParsedString = Left(ParsedString, Len(ParsedString) - 2)
        
        ParseColorsDictionary = ParsedString
        
        Erase AllKeys
        
    End Function
    
    Public Function ParseColorsJSON() As Scripting.Dictionary
    
        'Retrieves settings JSON text, loops through colors portion, and creates dictionary using data
    
    '0.0 Setup
    
        Const COLORS_HEADER As String = """Colors"" : ["
        Const NAME_HEADER As String = """Name"" : "
        Const CODE_HEADER As String = """Code"" : "
         
        Dim MySettingsText As String: MySettingsText = SettingsText
        Dim MyLine As String
        Dim ColorName As String
        Dim ColorCode As Long
        Dim StartPosition As Integer, EndPosition As Integer
        Dim x As Integer: x = 1
        
        Dim ColorsDictionary As Scripting.Dictionary: Set ColorsDictionary = New Scripting.Dictionary
        
    '1.0 Extract part of JSON text that contains color information
        
        StartPosition = InStr(1, MySettingsText, COLORS_HEADER)
        StartPosition = InStr(StartPosition, MySettingsText, vbCr) + 1
        
        EndPosition = InStr(StartPosition, MySettingsText, "]")
        
        MySettingsText = Trim(Mid(MySettingsText, StartPosition, EndPosition - StartPosition))
        
    '2.0 Process line-by-line
        
        Do Until HasValues(MySettingsText) = False
            
            MyLine = Mid(MySettingsText, 1, InStr(1, MySettingsText, vbCr)) 'Retrieve one line
            
            If InStr(1, MyLine, NAME_HEADER) = 0 Then GoTo NextLine         'Skip line, if there is no name tag
        
            ColorName = JSONData(NAME_HEADER, MyLine)                       'Retrieve color name
            
            MySettingsText = Replace(MySettingsText, MyLine, "")            'Erase retrieved line and move down one
            MyLine = Mid(MySettingsText, 1, InStr(1, MySettingsText, vbCr)) 'Retrieve next line
            
            ColorCode = JSONData(CODE_HEADER, MyLine)                       'Retrieve color code
            
            ColorsDictionary.Add ColorName, ColorCode
            
NextLine:
            
            MySettingsText = Replace(MySettingsText, MyLine, "")
            If Len(Replace(MySettingsText, Chr(10), "")) = 0 Then MySettingsText = ""
            
            x = x + 1
            
        Loop
        
        Set ParseColorsJSON = ColorsDictionary
        
    '3.0 Tear down
        
        Set ColorsDictionary = Nothing
        
    End Function
    
    Private Function JSONData(FieldName As String, JSONText As String) As String
    
        Dim StartPosition As Integer, EndPosition As Integer
        
        If InStr(1, FieldName, ":") = 0 Then FieldName = """" & FieldName & """ : "
        
        StartPosition = InStr(1, JSONText, FieldName) + Len(FieldName)
        
        EndPosition = InStr(StartPosition, JSONText, vbCr) - 1
        If Mid(JSONText, EndPosition, 1) = "," Then EndPosition = EndPosition - 1
        
        JSONData = Trim(Mid(JSONText, StartPosition, EndPosition - StartPosition + 1))
        
    End Function
    
    Public Property Get NumberOfRows() As Integer
        Const FIELD_NAME As String = "Number of Rows"
        NumberOfRows = JSONData(FIELD_NAME, Me.SettingsText)
    End Property
    
    Public Property Get NumberOfColumns() As Integer
        Const FIELD_NAME As String = "Number of Columns"
        NumberOfColumns = JSONData(FIELD_NAME, Me.SettingsText)
    End Property

    Public Property Get TileHeightWidthPixels() As Integer
        Const FIELD_NAME As String = "Tile Height/Width (px.)"
        TileHeightWidthPixels = JSONData(FIELD_NAME, Me.SettingsText)
    End Property
    
    Public Property Get TileCreationAlgorithm() As Integer
        Const FIELD_NAME As String = "Tile Creation Algorithm"
        TileCreationAlgorithm = JSONData(FIELD_NAME, Me.SettingsText)
    End Property
    
    Public Property Get ColorsArray() As Variant
        ColorsArray = ParseColorsJSON.Items
    End Property
    
    Public Property Get MERGE_CODE() As Integer: MERGE_CODE = 1: End Property
    Public Property Get BISECT_CODE() As Integer: BISECT_CODE = 2: End Property
