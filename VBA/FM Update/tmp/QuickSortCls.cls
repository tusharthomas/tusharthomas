VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QuickSortCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Private pData As Variant

    Public Function SortArrAsc(MyArr As Variant) As Variant
        pData = MyArr
        Call QuickSort(LBound(pData), UBound(pData))
        SortArrAsc = pData
    End Function

    Private Sub QuickSort(inLow As Integer, inHi As Integer)
    
        Dim pivot   As Variant
        Dim tmpSwap As Variant
        Dim tmpLow  As Integer
        Dim tmpHi   As Integer
        
        tmpLow = inLow
        tmpHi = inHi
    
        pivot = pData((inLow + inHi) \ 2)
    
        While (tmpLow <= tmpHi)
            While (pData(tmpLow) < pivot And tmpLow < inHi)
                tmpLow = tmpLow + 1
            Wend
    
            While (pivot < pData(tmpHi) And tmpHi > inLow)
                tmpHi = tmpHi - 1
            Wend
    
            If (tmpLow <= tmpHi) Then
                tmpSwap = pData(tmpLow)
                pData(tmpLow) = pData(tmpHi)
                pData(tmpHi) = tmpSwap
                tmpLow = tmpLow + 1
                tmpHi = tmpHi - 1
            End If
        Wend
    
        If (inLow < tmpHi) Then QuickSort inLow, tmpHi
        If (tmpLow < inHi) Then QuickSort tmpLow, inHi
        
    End Sub

