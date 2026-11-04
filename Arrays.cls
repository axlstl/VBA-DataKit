VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const LTOTALDIMCHECK As Long = 10

Enum FilterType
    eAND = 1
    eOR = 2
End Enum

Enum ResearchType
    eEqual = 1000
    eNotEqual = 1001
    eGreater = 1002
    eGreaterOrEqual = 1003
    eLess = 1004
    eLessOrEqual = 1005
    eTextInclude = 1006
    eTextExclude = 1007
End Enum

Enum SelectionType
    eSelectLine = 2
    eSelectColumn = 1
End Enum

Enum DeletionType
    eDeleteLine = 1
    eDeleteColumn = 2
End Enum

Enum GroupType
    eSum = 1
    eMean = 2
    eCount = 3
    eMin = 4
    eMax = 5
    eStDev = 6
End Enum

Function Exchange(ByRef val_a, ByRef val_b) As Variant
    'Swaps two values passed by reference

    Dim vTemp       As Variant
    Dim temp        As Variant

    temp = val_b
    val_b = val_a
    val_a = temp
End Function

Function DimCounter(ByVal vArr As Variant) As Integer
    'Returns the number of dimensions in an array

    Dim lDim            As Long
    Dim lDim_checker    As Long
    Dim iDim            As Long

    If Not IsArray(vArr) Then
        DimCounter = 0
        Exit Function
    End If

    For iDim = 1 To LTOTALDIMCHECK
        On Error Resume Next
        lDim_checker = UBound(vArr, iDim)
        If Err <> 0 Then
            Err.Clear
            On Error GoTo 0
            DimCounter = iDim - 1
            Exit Function
        End If
    Next iDim
End Function

Function SliceArray(ByVal vArr As Variant, ByVal SelectedType As SelectionType, ByVal lIndex As Long)
    'Extracts a specified row or column from a two-dimensional array

    Dim vArr_Out        As Variant
    Dim lDimResize      As Long
    Dim lSelect         As Long
    Dim lCol            As Long

    If Not IsArray(vArr) Then Exit Function
    If Me.DimCounter(vArr) <> 2 Then Exit Function
    Select Case SelectedType
        Case eSelectLine
            ReDim vArr_Out(1 To 1, LBound(vArr, 2) To UBound(vArr, 2))
            For lSelect = LBound(vArr_Out, SelectedType) To UBound(vArr_Out, SelectedType)
                vArr_Out(1, lSelect) = vArr(lIndex, lSelect)
            Next lSelect
        Case eSelectColumn
            ReDim vArr_Out(LBound(vArr, SelectedType) To UBound(vArr, SelectedType), 1 To 1)
            For lSelect = LBound(vArr, SelectedType) To UBound(vArr, SelectedType)
                vArr_Out(lSelect, 1) = vArr(lSelect, lIndex)
            Next lSelect
    End Select

    If IsArray(vArr_Out) And Not IsEmpty(vArr_Out) Then
        SliceArray = vArr_Out
    End If
End Function

Function Transpose(ByVal vArr As Variant) As Variant
    'Transposes a two-dimensional array by swapping its rows and columns

    Dim vArr_Out        As Variant
    Dim iRow            As Long
    Dim iCol            As Long

    If Not IsArray(vArr) Then Exit Function
    If Me.DimCounter(vArr) <> 2 Then Exit Function
    ReDim vArr_Out(LBound(vArr, 2) To UBound(vArr, 2), LBound(vArr, 1) To UBound(vArr, 1))
    For iRow = LBound(vArr_Out, 1) To UBound(vArr_Out, 1)
        For iCol = LBound(vArr_Out, 2) To UBound(vArr_Out, 2)
            vArr_Out(iRow, iCol) = vArr(iCol, iRow)
        Next iCol
    Next iRow

    Transpose = vArr_Out
End Function

Function Sort(ByVal vArr As Variant, Optional ByVal lCol As Integer = -1, Optional ByVal iSortSens As Integer = 1) As Variant
    'Sorts a one- or two-dimensional array by column using a stable merge sort

    Dim vArr_Out                As Variant
    Dim lIndex()                As Long
    Dim lTemp()                 As Long
    Dim iDim__                  As Integer
    Dim lRow                    As Long
    Dim lCol_Out                As Long

    iDim__ = Me.DimCounter(vArr)

    Select Case iDim__
        Case 1
            vArr = Me.DimExpand(vArr)
        Case 2
        Case Else
            Exit Function
    End Select

    If lCol < 0 Then lCol = LBound(vArr, 2)
    If lCol < LBound(vArr, 2) Or lCol > UBound(vArr, 2) Then Exit Function

    If iSortSens <> 0 And iSortSens <> 1 Then
        If iDim__ = 1 Then
            ReDim vArr_Out(LBound(vArr, 1) To UBound(vArr, 1))
            For lRow = LBound(vArr, 1) To UBound(vArr, 1)
                vArr_Out(lRow) = vArr(lRow, LBound(vArr, 2))
            Next lRow
            Sort = vArr_Out
        Else
            Sort = vArr
        End If
        Exit Function
    End If

    ReDim lIndex(LBound(vArr, 1) To UBound(vArr, 1))
    ReDim lTemp(LBound(vArr, 1) To UBound(vArr, 1))

    For lRow = LBound(vArr, 1) To UBound(vArr, 1)
        lIndex(lRow) = lRow
    Next lRow

    Call MergeSortRows(vArr, lIndex, lTemp, LBound(lIndex), UBound(lIndex), CLng(lCol), iSortSens)

    If iDim__ = 1 Then
        ReDim vArr_Out(LBound(vArr, 1) To UBound(vArr, 1))
        For lRow = LBound(vArr_Out) To UBound(vArr_Out)
            vArr_Out(lRow) = vArr(lIndex(lRow), LBound(vArr, 2))
        Next lRow
    Else
        ReDim vArr_Out(LBound(vArr, 1) To UBound(vArr, 1), LBound(vArr, 2) To UBound(vArr, 2))
        For lRow = LBound(vArr_Out, 1) To UBound(vArr_Out, 1)
            For lCol_Out = LBound(vArr_Out, 2) To UBound(vArr_Out, 2)
                vArr_Out(lRow, lCol_Out) = vArr(lIndex(lRow), lCol_Out)
            Next lCol_Out
        Next lRow
    End If

    Sort = vArr_Out
End Function

Private Sub MergeSortRows(ByRef vArr As Variant, ByRef lIndex() As Long, ByRef lTemp() As Long, _
                          ByVal lFirst As Long, ByVal lLast As Long, ByVal lCol As Long, ByVal iSortSens As Integer)
    'Recursively sorts row indices without moving the source data

    Dim lMiddle As Long

    If lFirst >= lLast Then Exit Sub

    lMiddle = lFirst + (lLast - lFirst) \ 2

    Call MergeSortRows(vArr, lIndex, lTemp, lFirst, lMiddle, lCol, iSortSens)
    Call MergeSortRows(vArr, lIndex, lTemp, lMiddle + 1, lLast, lCol, iSortSens)
    Call MergeSortedRows(vArr, lIndex, lTemp, lFirst, lMiddle, lLast, lCol, iSortSens)
End Sub

Private Sub MergeSortedRows(ByRef vArr As Variant, ByRef lIndex() As Long, ByRef lTemp() As Long, _
                            ByVal lFirst As Long, ByVal lMiddle As Long, ByVal lLast As Long, _
                            ByVal lCol As Long, ByVal iSortSens As Integer)
    'Merges sorted row indices while preserving the original order of equal values

    Dim lLeft                   As Long
    Dim lRight                  As Long
    Dim lOut                    As Long
    Dim bUseRight               As Boolean

    lLeft = lFirst
    lRight = lMiddle + 1
    lOut = lFirst

    Do While lLeft <= lMiddle And lRight <= lLast
        If iSortSens = 1 Then
            bUseRight = (vArr(lIndex(lLeft), lCol) > vArr(lIndex(lRight), lCol))
        Else
            bUseRight = (vArr(lIndex(lLeft), lCol) < vArr(lIndex(lRight), lCol))
        End If

        If bUseRight Then
            lTemp(lOut) = lIndex(lRight)
            lRight = lRight + 1
        Else
            lTemp(lOut) = lIndex(lLeft)
            lLeft = lLeft + 1
        End If
        lOut = lOut + 1
    Loop

    Do While lLeft <= lMiddle
        lTemp(lOut) = lIndex(lLeft)
        lLeft = lLeft + 1
        lOut = lOut + 1
    Loop

    Do While lRight <= lLast
        lTemp(lOut) = lIndex(lRight)
        lRight = lRight + 1
        lOut = lOut + 1
    Loop

    For lOut = lFirst To lLast
        lIndex(lOut) = lTemp(lOut)
    Next lOut
End Sub


Function DimReduce(ByVal vArr As Variant) As Variant
    'Converts a single-row or single-column 2D array into a 1D array

    Dim vArr_Out                As Variant
    Dim lRow                    As Long
    Dim lCol                    As Long

    If Not IsArray(vArr) Then Exit Function
    If Me.DimCounter(vArr) <> 2 Then Exit Function

    If LBound(vArr, 1) = UBound(vArr, 1) Then
        ReDim vArr_Out(LBound(vArr, 2) To UBound(vArr, 2))
        For lCol = LBound(vArr, 2) To UBound(vArr, 2)
            vArr_Out(lCol) = vArr(LBound(vArr, 1), lCol)
        Next lCol
    ElseIf LBound(vArr, 2) = UBound(vArr, 2) Then
        ReDim vArr_Out(LBound(vArr, 1) To UBound(vArr, 1))
        For lRow = LBound(vArr, 1) To UBound(vArr, 1)
            vArr_Out(lRow) = vArr(lRow, LBound(vArr, 2))
        Next lRow
    Else
        Exit Function
    End If

    DimReduce = vArr_Out
End Function

Function DimExpand(ByVal vArr As Variant) As Variant
    'Converts a one-dimensional array into a two-dimensional column array

    Dim vArr_Out        As Variant
    Dim iRow            As Long
    Dim iCol            As Long
    Dim iDim            As Long

    If Not IsArray(vArr) Then Exit Function
    If Me.DimCounter(vArr) <> 1 Then Exit Function
    ReDim vArr_Out(LBound(vArr, 1) To UBound(vArr, 1), 1 To 1)
    iCol = UBound(vArr, 1)
    For iRow = LBound(vArr_Out, 1) To UBound(vArr_Out, 1)
        vArr_Out(iRow, 1) = vArr(iRow)
    Next iRow

    If IsArray(vArr_Out) Then
        DimExpand = vArr_Out
    End If
End Function

Function DimSplit(ByVal vArr As Variant) As Variant
    'Converts a two-dimensional array into a one-dimensional array of row arrays

    Dim vArr_Template As Variant
    Dim lRow As Long
    Dim lCol As Long
    Dim vArr_Out As Variant

    If Not IsArray(vArr) Then Exit Function
    If Me.DimCounter(vArr) <> 2 Then Exit Function
    ReDim vArr_Out(LBound(vArr, 1) To UBound(vArr, 1))
    ReDim vArr_Template(LBound(vArr, 2) To UBound(vArr, 2))
    For lRow = LBound(vArr_Out) To UBound(vArr_Out)
        vArr_Out(lRow) = vArr_Template
        For lCol = LBound(vArr, 2) To UBound(vArr, 2)
            vArr_Out(lRow)(lCol) = vArr(lRow, lCol)
        Next lCol
    Next lRow

    If IsArray(vArr_Out) Then
        DimSplit = vArr_Out
    End If
End Function

Function DimMerge(ByVal vArr As Variant) As Variant
    'Converts an array of row arrays into a two-dimensional array

    Dim vArr_Out As Variant
    Dim lRow As Long
    Dim lCol As Long

    If Not IsArray(vArr) Then Exit Function
    If Me.DimCounter(vArr) <> 1 Then Exit Function
    If Not IsArray(vArr(LBound(vArr))) Then Exit Function
    If Me.DimCounter(vArr(LBound(vArr))) <> 1 Then Exit Function
    ReDim vArr_Out(LBound(vArr) To UBound(vArr), LBound(vArr(LBound(vArr))) To UBound(vArr(LBound(vArr))))
    For lRow = LBound(vArr_Out, 1) To UBound(vArr_Out, 1)
        For lCol = LBound(vArr_Out, 2) To UBound(vArr_Out, 2)
            vArr_Out(lRow, lCol) = vArr(lRow)(lCol)
        Next lCol
    Next lRow

    If IsArray(vArr_Out) Then
        DimMerge = vArr_Out
    End If
End Function

Function Filter(ByVal vArr As Variant, _
                ByVal lFilterColumn1 As Long, ByVal ResearchType1 As ResearchType, ByVal vComparedValues1 As Variant, Optional ByVal IncludeBlank As Boolean = False, _
                Optional ByVal FilterType As FilterType = eAND, _
                Optional ByVal lFilterColumn2 As Long = -1, Optional ByVal ResearchType2 As ResearchType = eEqual, Optional ByVal vComparedValues2 As Variant = Empty) As Variant
    'Filters an array using up to two criteria with a single output allocation

    Dim vArr_Out                As Variant
    Dim vCriterionArray1        As Variant
    Dim vCriterionArray2        As Variant
    Dim lSelectedRows()         As Long
    Dim lRow                    As Long
    Dim lCol                    As Long
    Dim lCount                  As Long
    Dim lOut                    As Long
    Dim bCase1                  As Boolean
    Dim bCaseTotal              As Boolean
    Dim bArray2_Empty           As Boolean
    Dim vValue                  As Variant

    If Not IsArray(vArr) Then Exit Function
    If IsEmpty(vComparedValues1) Then
        Debug.Print ("No filter applied")
        Filter = vArr
        Exit Function
    End If

    Select Case FilterType
        Case eAND, eOR
        Case Else
            Exit Function
    End Select

    Select Case Me.DimCounter(vArr)
        Case 1
            vArr = Me.DimExpand(vArr)
        Case 2
        Case Else
            Exit Function
    End Select

    If lFilterColumn1 < LBound(vArr, 2) Or lFilterColumn1 > UBound(vArr, 2) Then Exit Function

    bArray2_Empty = IsEmpty(vComparedValues2)
    If Not bArray2_Empty Then
        If lFilterColumn2 < LBound(vArr, 2) Or lFilterColumn2 > UBound(vArr, 2) Then Exit Function
    End If

    vCriterionArray1 = BuildCriterionArray(lFilterColumn1, vComparedValues1)
    If Not bArray2_Empty Then vCriterionArray2 = BuildCriterionArray(lFilterColumn2, vComparedValues2)

    ReDim lSelectedRows(1 To UBound(vArr, 1) - LBound(vArr, 1) + 1)

    For lRow = LBound(vArr, 1) To UBound(vArr, 1)
        vValue = vArr(lRow, lFilterColumn1)
        If IsError(vValue) Or IsNull(vValue) Then
            bCase1 = False
        Else
            bCase1 = CheckCriterion(vValue, ResearchType1, vCriterionArray1, IncludeBlank)
        End If

        If bArray2_Empty Then
            bCaseTotal = bCase1
        ElseIf FilterType = eAND Then
            bCaseTotal = False
            If bCase1 Then
                vValue = vArr(lRow, lFilterColumn2)
                If Not IsError(vValue) And Not IsNull(vValue) Then
                    bCaseTotal = CheckCriterion(vValue, ResearchType2, vCriterionArray2, IncludeBlank)
                End If
            End If
        Else
            bCaseTotal = True
            If Not bCase1 Then
                vValue = vArr(lRow, lFilterColumn2)
                If IsError(vValue) Or IsNull(vValue) Then
                    bCaseTotal = False
                Else
                    bCaseTotal = CheckCriterion(vValue, ResearchType2, vCriterionArray2, IncludeBlank)
                End If
            End If
        End If

        If bCaseTotal Then
            lCount = lCount + 1
            lSelectedRows(lCount) = lRow
        End If
    Next lRow

    If lCount = 0 Then Exit Function

    ReDim vArr_Out(1 To lCount, LBound(vArr, 2) To UBound(vArr, 2))

    For lOut = 1 To lCount
        For lCol = LBound(vArr, 2) To UBound(vArr, 2)
            vArr_Out(lOut, lCol) = vArr(lSelectedRows(lOut), lCol)
        Next lCol
    Next lOut

    Filter = vArr_Out
End Function

Private Function BuildCriterionArray(ByVal lColumn As Long, vArr As Variant) As Variant
    'Converts a single criterion or column/row vector to a one-dimensional array

    Dim vArr_Out                As Variant

    If IsArray(vArr) Then
        Select Case Me.DimCounter(vArr)
            Case 1
                BuildCriterionArray = vArr
            Case 2
                BuildCriterionArray = Me.DimReduce(vArr)
        End Select
    Else
        ReDim vArr_Out(1 To 1)
        vArr_Out(1) = vArr
        BuildCriterionArray = vArr_Out
    End If
End Function

Private Function CheckCriterion(vComparedValue As Variant, ResearchType As ResearchType, vCriterionArray As Variant, ByVal IncludeBlank As Boolean) As Variant
    'Checks whether a value satisfies a criterion, including inverse comparisons

    Dim bInvert                 As Boolean
    Dim vCriterionValue         As Variant
    Dim bCase                   As Boolean
    Dim lRow                    As Long

    If IsError(vComparedValue) Or IsNull(vComparedValue) Then Exit Function
    If Not IsArray(vCriterionArray) Then
        CheckCriterion = True
        Exit Function
    End If

    Select Case ResearchType
        Case eEqual, eLess, eGreater, eTextInclude
            bInvert = False
        Case Else
            bInvert = True
    End Select

    If Not IncludeBlank Then
        If IsEmpty(vComparedValue) Then
            CheckCriterion = bInvert
            Exit Function
        End If
    End If

    For lRow = LBound(vCriterionArray) To UBound(vCriterionArray)
        vCriterionValue = vCriterionArray(lRow)
        If IsError(vCriterionValue) Or IsNull(vCriterionValue) Then GoTo NextCriterion

        Select Case ResearchType
            Case eEqual, eNotEqual
                If IsNumeric(vComparedValue) And Len(vComparedValue) > 0 And IsNumeric(vCriterionValue) Then
                    If CDbl(vComparedValue) = CDbl(vCriterionValue) Then
                        bCase = True
                        Exit For
                    End If
                Else
                    If UCase(CStr(vComparedValue)) = UCase(CStr(vCriterionValue)) Then
                        bCase = True
                        Exit For
                    End If
                End If

            Case eLess, eGreaterOrEqual
                If IsNumeric(vComparedValue) And Len(vComparedValue) > 0 And IsNumeric(vCriterionValue) Then
                    If CDbl(vComparedValue) < CDbl(vCriterionValue) Then
                        bCase = True
                        Exit For
                    End If
                Else
                    If UCase(CStr(vComparedValue)) < UCase(CStr(vCriterionValue)) Then
                        bCase = True
                        Exit For
                    End If
                End If

            Case eGreater, eLessOrEqual
                If IsNumeric(vComparedValue) And Len(vComparedValue) > 0 And IsNumeric(vCriterionValue) Then
                    If CDbl(vComparedValue) > CDbl(vCriterionValue) Then
                        bCase = True
                        Exit For
                    End If
                Else
                    If UCase(CStr(vComparedValue)) > UCase(CStr(vCriterionValue)) Then
                        bCase = True
                        Exit For
                    End If
                End If

            Case eTextInclude, eTextExclude
                If InStr(1, CStr(vComparedValue), CStr(vCriterionValue), vbBinaryCompare) Then
                    bCase = True
                    Exit For
                End If
        End Select
NextCriterion:
    Next lRow

    If bInvert Then bCase = Not bCase
    CheckCriterion = bCase
End Function



Function Unique(ByVal vArr As Variant) As Variant
    'Returns the unique values from a column vector in first-appearance order

    Dim oDict                   As New Dictionary
    Dim vArr_Out                As Variant
    Dim vKeys()                 As Variant
    Dim vKey                    As Variant
    Dim lRow                    As Long
    Dim lCol                    As Long
    Dim lCount                  As Long
    Dim bBlank                  As Boolean
    Dim bNull                   As Boolean

    If Not IsArray(vArr) Then Exit Function
    If Me.DimCounter(vArr) = 1 Then vArr = Me.DimExpand(vArr)
    If Me.DimCounter(vArr) <> 2 Then Exit Function

    lCol = LBound(vArr, 2)
    ReDim vKeys(1 To UBound(vArr, 1) - LBound(vArr, 1) + 1)
    oDict.CompareMode = vbBinaryCompare

    For lRow = LBound(vArr, 1) To UBound(vArr, 1)
        vKey = vArr(lRow, lCol)
        If IsError(vKey) Then
            'Excel errors are skipped because they cannot be dictionary keys
        ElseIf IsNull(vKey) Then
            If Not bNull Then
                lCount = lCount + 1
                vKeys(lCount) = Null
                bNull = True
            End If
        ElseIf IsEmpty(vKey) Then
            If Not bBlank Then
                lCount = lCount + 1
                vKeys(lCount) = Empty
                bBlank = True
            End If
        ElseIf Not oDict.Exists(vKey) Then
            lCount = lCount + 1
            oDict.Add vKey, lCount
            vKeys(lCount) = vKey
        End If
    Next lRow

    If lCount = 0 Then Exit Function

    ReDim vArr_Out(1 To lCount, 1 To 1)
    For lRow = 1 To lCount
        vArr_Out(lRow, 1) = vKeys(lRow)
    Next lRow

    Unique = vArr_Out
End Function
Function InsertColumns(ByVal vArr As Variant, ByVal vColumns As Variant, Optional ByVal bRedim As Boolean = True) As Variant
    'Appends columns to a 2D array or overwrites its last columns when bRedim is False

    Dim vArr_Out                As Variant
    Dim lRow                    As Long
    Dim lCol                    As Long
    Dim lRowSource              As Long
    Dim lColSource              As Long
    Dim lStart                  As Long
    Dim lNumColumns             As Long
    Dim iDim                    As Integer

    If Not IsArray(vArr) Or Not IsArray(vColumns) Then Exit Function
    If Me.DimCounter(vArr) <> 2 Then Exit Function

    iDim = Me.DimCounter(vColumns)
    If iDim <> 1 And iDim <> 2 Then Exit Function

    If iDim = 1 Then
        If UBound(vColumns) - LBound(vColumns) <> UBound(vArr, 1) - LBound(vArr, 1) Then Exit Function
        lNumColumns = 1
    Else
        If UBound(vColumns, 1) - LBound(vColumns, 1) <> UBound(vArr, 1) - LBound(vArr, 1) Then Exit Function
        lNumColumns = UBound(vColumns, 2) - LBound(vColumns, 2) + 1
    End If

    If Not bRedim Then
        If lNumColumns > UBound(vArr, 2) - LBound(vArr, 2) + 1 Then Exit Function
    End If

    If bRedim Then
        ReDim vArr_Out(LBound(vArr, 1) To UBound(vArr, 1), LBound(vArr, 2) To UBound(vArr, 2) + lNumColumns)
    Else
        ReDim vArr_Out(LBound(vArr, 1) To UBound(vArr, 1), LBound(vArr, 2) To UBound(vArr, 2))
    End If

    For lRow = LBound(vArr, 1) To UBound(vArr, 1)
        For lCol = LBound(vArr, 2) To UBound(vArr, 2)
            vArr_Out(lRow, lCol) = vArr(lRow, lCol)
        Next lCol
    Next lRow

    lStart = UBound(vArr_Out, 2) - lNumColumns + 1

    For lRow = LBound(vArr_Out, 1) To UBound(vArr_Out, 1)
        lRowSource = lRow - LBound(vArr, 1)
        For lCol = lStart To UBound(vArr_Out, 2)
            If iDim = 1 Then
                vArr_Out(lRow, lCol) = vColumns(LBound(vColumns) + lRowSource)
            Else
                lColSource = LBound(vColumns, 2) + lCol - lStart
                vArr_Out(lRow, lCol) = vColumns(LBound(vColumns, 1) + lRowSource, lColSource)
            End If
        Next lCol
    Next lRow

    InsertColumns = vArr_Out
End Function

Function InsertLines(ByVal vArr As Variant, ByVal vLines As Variant, Optional ByVal bRedim As Boolean = True) As Variant
    'Appends rows to a 2D array or overwrites its last rows when bRedim is False

    Dim vArr_Out                As Variant
    Dim lRow                    As Long
    Dim lCol                    As Long
    Dim lRowSource              As Long
    Dim lColSource              As Long
    Dim lStart                  As Long
    Dim lNumLines               As Long
    Dim lFirstRow               As Long
    Dim lLastRow                As Long
    Dim lFirstCol               As Long
    Dim lLastCol                As Long
    Dim iDim                    As Integer
    Dim iLinesDim               As Integer

    If Not IsArray(vLines) Then
        InsertLines = vArr
        Exit Function
    End If

    If Not IsArray(vArr) Then
        InsertLines = vLines
        Exit Function
    End If

    iDim = Me.DimCounter(vArr)
    iLinesDim = Me.DimCounter(vLines)
    If iDim <> 1 And iDim <> 2 Then Exit Function
    If iLinesDim <> 1 And iLinesDim <> 2 Then Exit Function

    If iDim = 1 Then
        lFirstRow = 1
        lLastRow = 1
        lFirstCol = LBound(vArr)
        lLastCol = UBound(vArr)
    Else
        lFirstRow = LBound(vArr, 1)
        lLastRow = UBound(vArr, 1)
        lFirstCol = LBound(vArr, 2)
        lLastCol = UBound(vArr, 2)
    End If

    If iLinesDim = 1 Then
        If UBound(vLines) - LBound(vLines) <> lLastCol - lFirstCol Then Exit Function
        lNumLines = 1
    Else
        If UBound(vLines, 2) - LBound(vLines, 2) <> lLastCol - lFirstCol Then Exit Function
        lNumLines = UBound(vLines, 1) - LBound(vLines, 1) + 1
    End If

    If Not bRedim Then
        If lNumLines > lLastRow - lFirstRow + 1 Then Exit Function
    End If

    If bRedim Then
        ReDim vArr_Out(lFirstRow To lLastRow + lNumLines, lFirstCol To lLastCol)
    Else
        ReDim vArr_Out(lFirstRow To lLastRow, lFirstCol To lLastCol)
    End If

    For lRow = lFirstRow To lLastRow
        For lCol = lFirstCol To lLastCol
            If iDim = 1 Then
                vArr_Out(lRow, lCol) = vArr(lCol)
            Else
                vArr_Out(lRow, lCol) = vArr(lRow, lCol)
            End If
        Next lCol
    Next lRow

    lStart = UBound(vArr_Out, 1) - lNumLines + 1

    For lRow = lStart To UBound(vArr_Out, 1)
        lRowSource = lRow - lStart
        For lCol = lFirstCol To lLastCol
            If iLinesDim = 1 Then
                vArr_Out(lRow, lCol) = vLines(LBound(vLines) + lCol - lFirstCol)
            Else
                lColSource = LBound(vLines, 2) + lCol - lFirstCol
                vArr_Out(lRow, lCol) = vLines(LBound(vLines, 1) + lRowSource, lColSource)
            End If
        Next lCol
    Next lRow

    InsertLines = vArr_Out
End Function

Function DeleteElement(ByVal vArr As Variant, ByVal lIndexDel As Long, ByVal eDeletionType As DeletionType) As Variant
    'Removes one row or column while preserving the other dimension and its bounds

    Dim vArr_Out                As Variant
    Dim lRow                    As Long
    Dim lCol                    As Long
    Dim lRowSource              As Long
    Dim lColSource              As Long

    If Not IsArray(vArr) Then Exit Function
    If Me.DimCounter(vArr) <> 2 Then Exit Function

    Select Case eDeletionType
        Case eDeleteLine
            If lIndexDel < LBound(vArr, 1) Or lIndexDel > UBound(vArr, 1) Then
                DeleteElement = vArr
                Exit Function
            End If
            If LBound(vArr, 1) = UBound(vArr, 1) Then Exit Function
            ReDim vArr_Out(LBound(vArr, 1) To UBound(vArr, 1) - 1, LBound(vArr, 2) To UBound(vArr, 2))

        Case eDeleteColumn
            If lIndexDel < LBound(vArr, 2) Or lIndexDel > UBound(vArr, 2) Then
                DeleteElement = vArr
                Exit Function
            End If
            If LBound(vArr, 2) = UBound(vArr, 2) Then Exit Function
            ReDim vArr_Out(LBound(vArr, 1) To UBound(vArr, 1), LBound(vArr, 2) To UBound(vArr, 2) - 1)

        Case Else
            Exit Function
    End Select

    For lRow = LBound(vArr_Out, 1) To UBound(vArr_Out, 1)
        lRowSource = lRow
        If eDeletionType = eDeleteLine And lRow >= lIndexDel Then lRowSource = lRow + 1

        For lCol = LBound(vArr_Out, 2) To UBound(vArr_Out, 2)
            lColSource = lCol
            If eDeletionType = eDeleteColumn And lCol >= lIndexDel Then lColSource = lCol + 1
            vArr_Out(lRow, lCol) = vArr(lRowSource, lColSource)
        Next lCol
    Next lRow

    DeleteElement = vArr_Out
End Function


Function FillColumn(ByVal vArr As Variant, ByVal vColumn As Variant, ByVal lIndex As Long) As Variant
    'Fills a specified column using a 1D or 2D column vector

    Dim vArr_Out                As Variant
    Dim lRow                    As Long
    Dim lCol                    As Long
    Dim lSourceRow              As Long
    Dim lSourceCol              As Long

    If Not IsArray(vArr) Or Not IsArray(vColumn) Then Exit Function
    If Me.DimCounter(vArr) = 1 Then vArr = Me.DimExpand(vArr)
    If Me.DimCounter(vArr) <> 2 Then Exit Function
    If Me.DimCounter(vColumn) = 1 Then vColumn = Me.DimExpand(vColumn)
    If Me.DimCounter(vColumn) <> 2 Then Exit Function

    If lIndex < LBound(vArr, 2) Or lIndex > UBound(vArr, 2) Then Exit Function
    If LBound(vColumn, 2) <> UBound(vColumn, 2) Then Exit Function
    If UBound(vColumn, 1) - LBound(vColumn, 1) <> UBound(vArr, 1) - LBound(vArr, 1) Then Exit Function

    ReDim vArr_Out(LBound(vArr, 1) To UBound(vArr, 1), LBound(vArr, 2) To UBound(vArr, 2))
    lSourceCol = LBound(vColumn, 2)

    For lRow = LBound(vArr, 1) To UBound(vArr, 1)
        lSourceRow = LBound(vColumn, 1) + lRow - LBound(vArr, 1)
        For lCol = LBound(vArr, 2) To UBound(vArr, 2)
            If lCol = lIndex Then
                vArr_Out(lRow, lCol) = vColumn(lSourceRow, lSourceCol)
            Else
                vArr_Out(lRow, lCol) = vArr(lRow, lCol)
            End If
        Next lCol
    Next lRow

    FillColumn = vArr_Out
End Function

Function Change_Lbound(ByVal vArr As Variant, ByVal lLbound As Long, ByVal lDim As Long) As Variant
    'Changes the lower bound of a selected dimension without changing its values

    Dim vArr_Out                As Variant
    Dim lRow                    As Long
    Dim lCol                    As Long
    Dim lRowSource              As Long
    Dim lColSource              As Long
    Dim lLast                   As Long
    Dim iDim                    As Integer

    If Not IsArray(vArr) Then Exit Function
    iDim = Me.DimCounter(vArr)

    If iDim = 1 Then
        If lDim <> 1 Then Exit Function
        lLast = lLbound + UBound(vArr) - LBound(vArr)
        ReDim vArr_Out(lLbound To lLast)
        For lRow = lLbound To lLast
            vArr_Out(lRow) = vArr(LBound(vArr) + lRow - lLbound)
        Next lRow
    ElseIf iDim = 2 Then
        Select Case lDim
            Case 1
                lLast = lLbound + UBound(vArr, 1) - LBound(vArr, 1)
                ReDim vArr_Out(lLbound To lLast, LBound(vArr, 2) To UBound(vArr, 2))
            Case 2
                lLast = lLbound + UBound(vArr, 2) - LBound(vArr, 2)
                ReDim vArr_Out(LBound(vArr, 1) To UBound(vArr, 1), lLbound To lLast)
            Case Else
                Exit Function
        End Select

        For lRow = LBound(vArr_Out, 1) To UBound(vArr_Out, 1)
            lRowSource = LBound(vArr, 1) + lRow - LBound(vArr_Out, 1)
            For lCol = LBound(vArr_Out, 2) To UBound(vArr_Out, 2)
                lColSource = LBound(vArr, 2) + lCol - LBound(vArr_Out, 2)
                vArr_Out(lRow, lCol) = vArr(lRowSource, lColSource)
            Next lCol
        Next lRow
    Else
        Exit Function
    End If

    Change_Lbound = vArr_Out
End Function

Function GroupBy(ByVal vArr As Variant, ByVal lIndexKey As Long, ByVal lIndexGroup As Long, ByVal eGroupType As GroupType) As Variant
    'Groups a two-dimensional array and calculates one statistic per key in a single pass

    Dim oDictKey                As New Dictionary
    Dim vArr_Out                As Variant
    Dim vKeys()                 As Variant
    Dim vStats()                As Double
    Dim vKey                    As Variant
    Dim vValue                  As Variant
    Dim lRow                    As Long
    Dim lGroup                  As Long
    Dim lCount                  As Long
    Dim lBlank                  As Long
    Dim lRows                   As Long
    Dim dblValue                As Double
    Dim dblDelta                As Double
    Dim dblVariance             As Double

    If Not IsArray(vArr) Then
        Debug.Print ("Error GroupBy -> vArr is not an array")
        GroupBy = vArr
        Exit Function
    End If

    If Me.DimCounter(vArr) <> 2 Then
        Debug.Print ("Error GroupBy -> vArr must have two dimensions")
        GroupBy = vArr
        Exit Function
    End If

    If lIndexKey < LBound(vArr, 2) Or lIndexKey > UBound(vArr, 2) Then Exit Function
    If lIndexGroup < LBound(vArr, 2) Or lIndexGroup > UBound(vArr, 2) Then Exit Function

    Select Case eGroupType
        Case eSum, eMean, eCount, eMin, eMax, eStDev
        Case Else
            Exit Function
    End Select

    lRows = UBound(vArr, 1) - LBound(vArr, 1) + 1
    ReDim vKeys(1 To lRows)
    ReDim vStats(1 To lRows, 1 To 3)
    oDictKey.CompareMode = vbBinaryCompare

    For lRow = LBound(vArr, 1) To UBound(vArr, 1)
        vKey = vArr(lRow, lIndexKey)

        If IsError(vKey) Then
            Debug.Print ("Error GroupBy -> error value in grouping column")
            Exit Function
        End If
        If IsNull(vKey) Then
            Debug.Print ("Error GroupBy -> Null key")
            Exit Function
        End If

        If IsEmpty(vKey) Then
            If lBlank = 0 Then
                lCount = lCount + 1
                lBlank = lCount
                vKeys(lCount) = vKey
            End If
            lGroup = lBlank
        Else
            If Not oDictKey.Exists(vKey) Then
                lCount = lCount + 1
                oDictKey.Add vKey, lCount
                vKeys(lCount) = vKey
            End If
            lGroup = oDictKey(vKey)
        End If

        vValue = vArr(lRow, lIndexGroup)
        If Not IsError(vValue) Then
            If Not IsNull(vValue) Then
                If eGroupType = eCount Then
                    If Not IsEmpty(vValue) Then
                        If Len(CStr(vValue)) > 0 Then vStats(lGroup, 1) = vStats(lGroup, 1) + 1
                    End If
                ElseIf Not IsEmpty(vValue) And IsNumeric(vValue) Then
                    dblValue = CDbl(vValue)
                    Select Case eGroupType
                        Case eSum
                            vStats(lGroup, 2) = vStats(lGroup, 2) + dblValue
                        Case eMean, eStDev
                            vStats(lGroup, 1) = vStats(lGroup, 1) + 1
                            dblDelta = dblValue - vStats(lGroup, 2)
                            vStats(lGroup, 2) = vStats(lGroup, 2) + dblDelta / vStats(lGroup, 1)
                            If eGroupType = eStDev Then
                                vStats(lGroup, 3) = vStats(lGroup, 3) + dblDelta * (dblValue - vStats(lGroup, 2))
                            End If
                        Case eMin
                            If vStats(lGroup, 1) = 0 Then
                                vStats(lGroup, 2) = dblValue
                            ElseIf dblValue < vStats(lGroup, 2) Then
                                vStats(lGroup, 2) = dblValue
                            End If
                            vStats(lGroup, 1) = vStats(lGroup, 1) + 1
                        Case eMax
                            If vStats(lGroup, 1) = 0 Then
                                vStats(lGroup, 2) = dblValue
                            ElseIf dblValue > vStats(lGroup, 2) Then
                                vStats(lGroup, 2) = dblValue
                            End If
                            vStats(lGroup, 1) = vStats(lGroup, 1) + 1
                    End Select
                End If
            End If
        End If
    Next lRow

    ReDim vArr_Out(1 To lCount, 1 To 2)

    For lRow = 1 To lCount
        vArr_Out(lRow, 1) = vKeys(lRow)
        Select Case eGroupType
            Case eCount
                vArr_Out(lRow, 2) = vStats(lRow, 1)
            Case eStDev
                If vStats(lRow, 1) > 0 Then
                    dblVariance = vStats(lRow, 3) / vStats(lRow, 1)
                    If dblVariance < 0 Then dblVariance = 0
                    vArr_Out(lRow, 2) = Sqr(dblVariance)
                Else
                    vArr_Out(lRow, 2) = 0#
                End If
            Case Else
                vArr_Out(lRow, 2) = vStats(lRow, 2)
        End Select
    Next lRow

    GroupBy = vArr_Out
End Function




