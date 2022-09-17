VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Public Function ReadContentsFromFile(FullPath As String, Optional ReplaceStrings As Variant) As String
        
        Const REPLACE_DELIMITER As String = ":"
        
        Dim MyFileSystemObject  As Scripting.FileSystemObject
        Dim MyTextStream        As Scripting.TextStream
        
        Dim x                   As Integer
        Dim ToReplace$, MyVal   As String
        
        Set MyFileSystemObject = New Scripting.FileSystemObject
        
        Set MyTextStream = MyFileSystemObject.OpenTextFile(FullPath)
        ReadContentsFromFile = MyTextStream.ReadAll
        
        If Not IsEmpty(ReplaceStrings) And Not IsMissing(ReplaceStrings) Then
            For x = LBound(ReplaceStrings) To UBound(ReplaceStrings)
                ToReplace = Split(ReplaceStrings(x), REPLACE_DELIMITER)(0)
                MyVal = Split(ReplaceStrings(x), REPLACE_DELIMITER)(1)
                ReadContentsFromFile = Replace(ReadContentsFromFile, ToReplace, MyVal)
            Next x
        End If
        
    End Function
    
    Public Function GetWorkbook(FullPath As String, ByRef WBIsOpen As Boolean, Optional OpenReadOnly As Boolean = True) As Workbook
      
        'Check whether the provided workbook is already open; if not, open it
        'Output a boolean variable WBIsOpen to report whether the workbook was open
        
        Const OPEN_WB_ERROR_PROMPT As String = "ERROR: Failed to open file at the following location: @FilePath"
        Const FILE_PATH_PLACEHOLDER As String = "@FilePath"
        Const OPEN_WB_ERROR_TITLE As String = "Failed to open"
        
        Dim ErrorOnOpen As Boolean
        Dim FileName As String: FileName = ExtractNameFromPath(FullPath)
        
        Dim WB As Workbook
        
        'Check if workbook is already open
        
        For Each WB In Workbooks
            If WB.Name = Trim(FileName) Then
                Set GetWorkbook = WB
                WBIsOpen = True
                GoTo CheckReadOnly
            End If
        Next
        
        WBIsOpen = False
        
        'If not, try to open
        
        On Error Resume Next
        
            Set WB = Workbooks.Open(FileName:=FullPath, UpdateLinks:=False, ReadOnly:=OpenReadOnly)
            
            If Err.Number > 0 Then
        
                Call MsgBox(Prompt:=Replace(OPEN_WB_ERROR_PROMPT, FILE_PATH_PLACEHOLDER, FullPath), Title:=OPEN_WB_ERROR_TITLE, Buttons:=vbOKOnly)
                FullPath = GetFileSelection(FullPath, FileName)
                Set GetWorkbook = Workbooks.Open(FileName:=FullPath, UpdateLinks:=False, ReadOnly:=OpenReadOnly)
            
            End If
            
        On Error GoTo 0
        
CheckReadOnly:
        
        'Check If ReadOnly
                        
        On Error Resume Next
        If Not OpenReadOnly Then
            If WB.ReadOnly Then
                MsgBox "Warning! The resource you are trying to access (" & FullPath & ") is currently being used by somebody else.", vbExclamation, "Read Only"
                    Call Common.ToggleApplicationSettings(True)
                    End
            End If
        End If
        On Error GoTo 0
        
        Set GetWorkbook = WB
        
    End Function
    
    Public Function GetFileSelection(FullPath As String, FileName As String) As String
        Const NO_SELECTION_PROMPT As String = "No selection. Macro runtime will end."
        Const NO_SELECTION_TITLE  As String = "No selection"
        With Application.FileDialog(msoFileDialogFilePicker)
            .Title = FileName
            .InitialFileName = Replace(FullPath, FileName, "")
            .AllowMultiSelect = False
            .Filters.Clear
            .Filters.Add "Excel", "*.xls*"
            .Show
            If .SelectedItems.Count > 0 Then
                GetFileSelection = .SelectedItems(1)
            Else
                Call MsgBox(Prompt:=NO_SELECTION_PROMPT, Title:=NO_SELECTION_TITLE, Buttons:=vbOKOnly)
                End
            End If
        End With
    End Function
    
    Public Function ExtractNameFromPath(FullPath As String) As String
        Dim NumberOfSlashes As Integer: NumberOfSlashes = UBound(Split(FullPath, "\"))
        ExtractNameFromPath = Split(FullPath, "\")(NumberOfSlashes)
    End Function
    
    Public Sub CreatePath(FolderPath As String)
        Dim x As Integer
        Dim path As String
        Dim SubFolders() As String: SubFolders = Split(FolderPath, "\")
        For x = LBound(SubFolders) To UBound(SubFolders)
            If path <> "" Then path = path & "\"
            path = path & SubFolders(x)
            If Dir(path, vbDirectory) = "" Then MkDir path
        Next x
    End Sub
    
    Public Sub MoveMe(TargetFolder As String, Optional NewName As String)
        Dim OldFile As String
        If Right(TargetFolder, 1) <> "\" Then TargetFolder = TargetFolder & "\"
        If Dir(TargetFolder, vbDirectory) = "" Then Call CreatePath(TargetFolder)
        With ThisWorkbook
            OldFile = .FullName
            If NewName = "" Then NewName = .Name
            .SaveAs TargetFolder & NewName
        End With
        Kill OldFile
    End Sub
