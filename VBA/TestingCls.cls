VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestingCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Const DEFAULT_MIN As Integer = 0
    Const DEFAULT_MAX As Integer = 32000
    Const DEFAULT_STRLEN As Integer = 10

    Public Enum TestTypeEnum
        RndInt
        RndStr
    End Enum
    
    Private pStartTime As Single

    Public Function CreateDummyRecord(NumberOfFields As Integer, Optional TestType As TestTypeEnum = -1, Optional WSSafe As Boolean) As Variant
        
        Dim Arr As Variant
        Dim x As Integer
        
        ReDim Arr(NumberOfFields - 1)
        
        For x = LBound(Arr) To UBound(Arr)
            If TestType = -1 Then
                Arr(x) = GetRandVar
            Else
                Select Case TestType
                    Case RndInt:    Arr(x) = GetRandInt
                    Case RndStr:    Arr(x) = GetRandStr(, WSSafe)
                End Select
            End If
        Next x
        
        CreateDummyRecord = Arr
        
    End Function

    Public Function GetRandInt(Optional Min As Integer = DEFAULT_MIN, Optional Max As Integer = DEFAULT_MAX) As Integer
        GetRandInt = CInt(Rnd * (Max - Min)) + Min
    End Function
    
    Public Function GetRandStr(Optional Length As Integer = DEFAULT_STRLEN, Optional WSSafe As Boolean) As String
        Dim x As Integer
        Dim str As String
        For x = 1 To Length
            str = str & Chr(GetRandInt(36, 122))    '36 & 122 correspond to character codes in the ASCII character set
        Next x
        If WSSafe Then str = "'" & str
        GetRandStr = str
    End Function
    
    Public Function GetRandVar(Optional MinNum As Integer = DEFAULT_MIN, Optional MaxNum As Integer = DEFAULT_MAX, Optional StrLen As Integer = DEFAULT_STRLEN) As Variant
        
        Select Case GetRandInt(RndInt, RndStr)
            Case RndInt:        GetRandVar = GetRandInt(MinNum, MaxNum)
            Case RndStr:        GetRandVar = GetRandStr(StrLen)
        End Select
        
    End Function
    
    Public Sub StartTimer():                    pStartTime = Timer:                     End Sub
    Public Function StopTimer() As Single:      StopTimer = Timer - pStartTime:         End Function
