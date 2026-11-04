VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Finance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum ExtremumType
    eMin = 0
    eMax = 1
End Enum

Function Sum(ByVal vArr As Variant) As Double
    'Returns the sum of numeric values in a column vector

    Dim lRow                    As Long
    Dim lCol                    As Long
    Dim vValue                  As Variant
    Dim dblSum                  As Double

    If Not IsArray(vArr) Then Exit Function
    If Arr.DimCounter(vArr) = 1 Then vArr = Arr.DimExpand(vArr)
    If Arr.DimCounter(vArr) <> 2 Then Exit Function

    lCol = LBound(vArr, 2)

    For lRow = LBound(vArr, 1) To UBound(vArr, 1)
        vValue = vArr(lRow, lCol)
        If Not IsError(vValue) Then
            If Not IsNull(vValue) Then
                If Not IsEmpty(vValue) Then
                    If IsNumeric(vValue) Then dblSum = dblSum + CDbl(vValue)
                End If
            End If
        End If
    Next lRow

    Sum = dblSum
End Function

Function SquaredSum(ByVal vArr As Variant) As Double
    'Returns the sum of squared numeric values in a column vector

    Dim lRow                    As Long
    Dim lCol                    As Long
    Dim vValue                  As Variant
    Dim dblValue                As Double
    Dim dblSum                  As Double

    If Not IsArray(vArr) Then Exit Function
    If Arr.DimCounter(vArr) = 1 Then vArr = Arr.DimExpand(vArr)
    If Arr.DimCounter(vArr) <> 2 Then Exit Function

    lCol = LBound(vArr, 2)

    For lRow = LBound(vArr, 1) To UBound(vArr, 1)
        vValue = vArr(lRow, lCol)
        If Not IsError(vValue) Then
            If Not IsNull(vValue) Then
                If Not IsEmpty(vValue) Then
                    If IsNumeric(vValue) Then
                        dblValue = CDbl(vValue)
                        dblSum = dblSum + dblValue * dblValue
                    End If
                End If
            End If
        End If
    Next lRow

    SquaredSum = dblSum
End Function

Function Count(ByVal vArr As Variant, Optional ByVal bNumeric As Boolean = False) As Long
    'Counts non-empty values, or only numeric values when bNumeric is True

    Dim lRow                    As Long
    Dim lCol                    As Long
    Dim vValue                  As Variant
    Dim n                       As Long

    If Not IsArray(vArr) Then Exit Function
    If Arr.DimCounter(vArr) = 1 Then vArr = Arr.DimExpand(vArr)
    If Arr.DimCounter(vArr) <> 2 Then Exit Function

    lCol = LBound(vArr, 2)

    For lRow = LBound(vArr, 1) To UBound(vArr, 1)
        vValue = vArr(lRow, lCol)
        If Not IsError(vValue) Then
            If Not IsNull(vValue) Then
                If Not IsEmpty(vValue) Then
                    If Len(CStr(vValue)) > 0 Then
                        If bNumeric Then
                            If IsNumeric(vValue) Then n = n + 1
                        Else
                            n = n + 1
                        End If
                    End If
                End If
            End If
        End If
    Next lRow

    Count = n
End Function

Function Mean(ByVal vArr As Variant) As Double
    'Returns the mean of numeric values in a column vector

    Dim n                       As Long

    If Not IsArray(vArr) Then Exit Function
    If Arr.DimCounter(vArr) = 1 Then vArr = Arr.DimExpand(vArr)
    If Arr.DimCounter(vArr) <> 2 Then Exit Function

    n = Me.Count(vArr, True)
    If n = 0 Then Exit Function

    Mean = Me.Sum(vArr) / n
End Function

Function StDev(ByVal vArr As Variant) As Double
    'Returnds stdof a variant

    Dim lRow                    As Long
    Dim lCol                    As Long
    Dim n                       As Long
    Dim vValue                  As Variant
    Dim dblMean                 As Double
    Dim dblValue                As Double
    Dim dblDelta                As Double
    Dim dblM2                   As Double

    If Not IsArray(vArr) Then Exit Function
    If Arr.DimCounter(vArr) = 1 Then vArr = Arr.DimExpand(vArr)
    If Arr.DimCounter(vArr) <> 2 Then Exit Function

    lCol = LBound(vArr, 2)

    For lRow = LBound(vArr, 1) To UBound(vArr, 1)
        vValue = vArr(lRow, lCol)
        If Not IsError(vValue) Then
            If Not IsNull(vValue) Then
                If Not IsEmpty(vValue) Then
                    If IsNumeric(vValue) Then
                        'welford algoritlhm
                        dblValue = CDbl(vValue)
                        n = n + 1
                        dblDelta = dblValue - dblMean
                        dblMean = dblMean + dblDelta / n
                        dblM2 = dblM2 + dblDelta * (dblValue - dblMean)
                    End If
                End If
            End If
        End If
    Next lRow

    If n = 0 Then Exit Function
    If dblM2 < 0 Then dblM2 = 0

    StDev = Sqr(dblM2 / n)
End Function

Function Extremum(ByVal vArr As Variant, ByVal eExtremum As ExtremumType) As Double
    'Returns the minimum or maximum numeric value in a column vector
    'eMax = Max
    'eMin = Min
    
    Dim lRow                    As Long
    Dim lCol                    As Long
    Dim vValue                  As Variant
    Dim dblExtremum             As Double
    Dim dblValue                As Double
    Dim bFound                  As Boolean

    If Not IsArray(vArr) Then Exit Function
    If Arr.DimCounter(vArr) = 1 Then vArr = Arr.DimExpand(vArr)
    If Arr.DimCounter(vArr) <> 2 Then Exit Function
    If eExtremum <> eMin And eExtremum <> eMax Then Exit Function

    lCol = LBound(vArr, 2)

    For lRow = LBound(vArr, 1) To UBound(vArr, 1)
        vValue = vArr(lRow, lCol)
        If Not IsError(vValue) Then
            If Not IsNull(vValue) Then
                If Not IsEmpty(vValue) Then
                    If IsNumeric(vValue) Then
                        dblValue = CDbl(vValue)
                        If Not bFound Then
                            dblExtremum = dblValue
                            bFound = True
                        ElseIf eExtremum = eMax Then
                            If dblValue > dblExtremum Then dblExtremum = dblValue
                        Else
                            If dblValue < dblExtremum Then dblExtremum = dblValue
                        End If
                    End If
                End If
            End If
        End If
    Next lRow

    If bFound Then Extremum = dblExtremum
End Function



