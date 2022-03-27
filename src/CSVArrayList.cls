VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' Copyright © 2022 W. García
' GPL-3.0 license | https://www.gnu.org/licenses/gpl-3.0.html/
' https://ingwilfredogarcia.wordpress.com
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' GENERAL INFO:
' Class module developed to emulate some functionalities from the ArrayList present in some
' most modern languages. The CSVArrayList serve as a container for all the data read from
' CSV files and can be used to manipulate the stored items, or to store data that does not
' come from a CSV file, according to the user's request.
'////////////////////////////////////////////////////////////////////////////////////////////
'#
Option Explicit
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' CONSTANTS:
Private Const RUN As Long = 32
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' VARIABLES:
Private Buffer() As Variant
Private CurrentIndex As Long
Private keysLB As Long
Private keysUB As Long
Private MaxIndex As Long
Private REC_DEPTH As Long         ' IntroSort QuickSort recursions limit
Private tmpBuffer As Variant
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' CONSTANTS:
Private Const d_Apostrophe As String = "'"
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' ENUMERATIONS:
Public Enum SortingAlgorithms
    SA_IntroSort = 0
    SA_Quicksort = 1
    SA_TimSort = 2
    SA_HeapSort = 3
    SA_MergeSort = 4
End Enum
'////////////////////////////////////////////////////////////////////////////////////////////
'#

''' <summary>
''' Initializes this instance.
''' </summary>
Private Sub Class_Initialize()
    CurrentIndex = -1&
    MaxIndex = 1024& - 1&
    ReDim Buffer(0 To MaxIndex)
End Sub

''' <summary>
''' Finalizes this instance.
''' </summary>
Private Sub Class_Terminate()
    Erase Buffer
End Sub

'////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES:
'#
''' <summary>
''' Gets the amount of items stored in the current instance.
''' </summary>
Public Property Get count() As Long
Attribute count.VB_Description = "Gets the amount of items stored in the current instance"
    count = CurrentIndex + 1
End Property

''' <summary>
''' Gets or sets the Item, by its index, from the current instance.
''' </summary>
''' <param name="Index">Item's position in this instance.</param>
Public Property Get item(Index As Long) As Variant
Attribute item.VB_Description = "Gets the Item, by its index, from the current instance."
Attribute item.VB_UserMemId = 0
    Select Case Index
        Case 0 To CurrentIndex
            item = Buffer(Index)
        Case Else
            err.Raise 9
    End Select
End Property

''' <summary>
''' Gets or sets the Item, by its index, from the current instance.
''' </summary>
''' <param name="Index">Item's position in this instance.</param>
''' <param name="aValue">The value to overwrite the element.</param>
Public Property Let item(Index As Long, AValue As Variant)
    Select Case Index
        Case 0 To CurrentIndex
            Buffer(Index) = AValue
        Case Else
            err.Raise 9
    End Select
End Property

''' <summary>
''' Gets or sets the set of elements of the current instance.
''' </summary>
Public Property Get items() As Variant
Attribute items.VB_Description = "Gets an array with all the Items from the current instance."
    If CurrentIndex >= 0 Then
        Dim tmpResult() As Variant
        tmpResult = Buffer
        If MaxIndex <> CurrentIndex Then
            ReDim Preserve tmpResult(0 To CurrentIndex)
        End If
        items = tmpResult
    End If
End Property

''' <summary>
''' Gets or sets the current instance's array.
''' </summary>
Public Property Let items(AValue As Variant)
    Clear
    If IsArray(AValue) Then
        Dim Dim1Pointer As Long
        Dim Dim2Pointer As Long
        Dim tmpRow() As Variant
        
        If MultiDimensional(AValue) Then '2D array expected
            ReDim tmpRow(LBound(AValue, 2) To UBound(AValue, 2))
            For Dim1Pointer = LBound(AValue) To UBound(AValue)
                For Dim2Pointer = LBound(AValue, 2) To UBound(AValue, 2)
                    tmpRow(Dim2Pointer) = AValue(Dim1Pointer, Dim2Pointer)
                Next Dim2Pointer
                Add tmpRow
            Next Dim1Pointer
        Else 'Jagged or 1D array expected
            If IsJaggedArray(AValue) Then
                For Dim1Pointer = LBound(AValue) To UBound(AValue)
                    Add AValue(Dim1Pointer)
                Next Dim1Pointer
            Else
                For Dim1Pointer = LBound(AValue) To UBound(AValue)
                    Add2 AValue
                Next Dim1Pointer
            End If
        End If
    Else
        Add2 AValue
    End If
End Property
'////////////////////////////////////////////////////////////////////////////////////////////
'#

'////////////////////////////////////////////////////////////////////////////////////////////
' METHODS:
'#
''' <summary>
''' Appends a copy of the specified value to the current instance.
''' </summary>
''' <param name="aValue">The value to append.</param>
Public Sub Add(AValue As Variant)
Attribute Add.VB_Description = "Appends a copy of the specified value to the current instance."
    CurrentIndex = CurrentIndex + 1
    On Error GoTo Expand_Buffer
    Buffer(CurrentIndex) = AValue
    Exit Sub
Expand_Buffer:
    MaxIndex = 2 * (MaxIndex + 1) - 1
    ReDim Preserve Buffer(0 To MaxIndex)
    Buffer(CurrentIndex) = AValue
End Sub
''' <summary>
''' Appends a copy, in jagged array fashion, of the specified
''' values to the current instance.
''' </summary>
''' <param name="aValue">The value to append.</param>
Public Sub Add2(ParamArray aValues() As Variant)
Attribute Add2.VB_Description = "Appends a copy, in jagged array fashion, of the specified value to the current instance."
    Dim i As Long
    Dim tmpValue() As Variant
    
    ReDim tmpValue(0 To 0)
    For i = LBound(aValues) To UBound(aValues)
        CurrentIndex = CurrentIndex + 1
        Select Case CurrentIndex
            Case Is <= MaxIndex
                If IsArray(aValues(i)) Then
                    Buffer(CurrentIndex) = aValues(i)
                Else
                    tmpValue(0) = aValues(i)
                    Buffer(CurrentIndex) = tmpValue
                End If
            Case Else
                MaxIndex = 2 * (MaxIndex + 1) - 1
                ReDim Preserve Buffer(0 To MaxIndex)
                If IsArray(aValues(i)) Then
                    Buffer(CurrentIndex) = aValues(i)
                Else
                    tmpValue(0) = aValues(i)
                    Buffer(CurrentIndex) = tmpValue
                End If
        End Select
    Next i
End Sub

''' <summary>
''' Turns the array A=[1..n] into a max-heap in a bottom-up manner.
''' </summary>
''' <param name="leftt">Lower bound on the sorting  operation (base 0).</param>
''' <param name="SortingKey">The column/key for the logical comparisons.</param>
''' <param name="HeapSize">Elements to be sorted.</param>
Private Sub build_max_heap(leftt As Long, _
                        SortingKey As Long, _
                        HeapSize As Long)
    Dim i As Long

    For i = Floor(HeapSize / 2) To 1 Step -1
        max_heapify i, leftt, SortingKey, HeapSize
    Next i
End Sub

''' <summary>
''' Turns the array A=[1..n] into a max-heap in a up-bottom manner.
''' </summary>
''' <param name="leftt">Lower bound on the sorting  operation (base 0).</param>
''' <param name="SortingKey">The column/key for the logical comparisons.</param>
''' <param name="HeapSize">Elements to be sorted.</param>
Private Sub build_min_heap(leftt As Long, _
                        SortingKey As Long, _
                        HeapSize As Long)
    Dim i As Long

    For i = Floor(HeapSize / 2) To 1 Step -1
        min_heapify i, leftt, SortingKey, HeapSize
    Next i
End Sub

''' <summary>
''' Returns the smallest integral value greater than or
''' equal to the specified number.
''' Equivalent to .NET Math.Ceiling function.
''' </summary>
Private Function Ceiling(value As Double) As Long
    Dim tmpResult As Long
    
    tmpResult = Fix(value)
    Ceiling = tmpResult - ((value <> tmpResult) And (value > 0))
End Function

''' <summary>
''' Creates an empty jagged array.
''' </summary>
''' <param name="ArrVar">Output array.</param>
''' <param name="ArraySize">Max row index (base 0).</param>
''' <param name="VectorSize">Max column index (base 0).</param>
Public Sub CreateJagged(ByRef ArrVar() As Variant, ArraySize As Long, VectorSize As Long)
    Dim Vector() As String
    Dim arrPointer As Long
    
    ReDim ArrVar(0 To ArraySize)
    ReDim Vector(0 To VectorSize)

    For arrPointer = 0 To ArraySize
        ArrVar(arrPointer) = Vector()
    Next arrPointer
    Erase Vector
End Sub

''' <summary>
''' Delimits the next sorting block index.
''' </summary>
''' <param name="OldKey">The key/column used in the last sorting.</param>
''' <param name="StartPos">Sorting block start.</param>
''' <param name="MaxIndex">Last index in the sorting operation.</param>
Private Function EndOfSortingBlock(OldKey As Long, _
                                    StartPos As Long, _
                                    MaxIndex As Long) As Long
    Dim i As Long
    Dim switch As Boolean
    Dim tmpRecordIndex As Long
    
    i = StartPos
    tmpRecordIndex = i
    Do While i <= MaxIndex And Not switch
        switch = Buffer(tmpRecordIndex)(OldKey) <> Buffer(i)(OldKey)
        i = i + (1 + switch)
    Loop
    EndOfSortingBlock = i - 1
End Function

''' <summary>
''' Returns the indexes of the fields
''' to be used in the filter operation.
''' </summary>
''' <param name="VarList">List of varibles [f1; f2; ...; fn].</param>
Private Function FieldsToFilter(VarList As String) As Long()
    Dim tmpResult() As Long
    Dim tmpVars() As String
    Dim jC As Long
    Dim opLB As Long
    Dim opUB As Long
    Dim IdxLen As Long
    
    tmpVars() = Split(VarList, "; ")                    'Separates variables [f#; f#...]
    opLB = LBound(tmpVars)
    opUB = UBound(tmpVars)
    ReDim tmpResult(opLB To opUB)
    For jC = opLB To opUB
        IdxLen = LenB(tmpVars(jC)) - (InStrB(1, tmpVars(jC), "f") + 1)
        If IdxLen Then 'At least one digit to extract
            tmpResult(jC) = CLng(MidB$(tmpVars(jC), 3))
        Else 'Input error
            ReDim tmpResult(0)
            tmpResult(0) = -1
            FieldsToFilter = tmpResult
            Exit Function
        End If
    Next jC
    FieldsToFilter = tmpResult
End Function

''' <summary>
''' Returns a filtered array list using the
''' CSVexpressions class module.
''' </summary>
''' <param name="Pattern">The expression to eval.</param>
''' <param name="StartIndex">The index to start filter records (base 1).</param>
Public Function Filter(Pattern As String, startIndex As Long) As CSVArrayList
    Dim Evaluator As CSVexpressions
    Dim rCounter As Long
    Dim FilterFields() As Long
    
    If startIndex < 1 Then Exit Function
    Set Filter = New CSVArrayList
    Set Evaluator = New CSVexpressions
    With Evaluator
        .GallopingMode = False
        .Create Pattern
        FilterFields() = FieldsToFilter(.CurrentVariables)
        For rCounter = startIndex - 1 To CurrentIndex
            On Error Resume Next
            .Eval FilterVarValues(rCounter, FilterFields)
            If .ErrorType = ExpressionErrors.errNone Then
                If err.Number = 0 Then
                    If CBool(.Result) Then
                        Filter.Add Buffer(rCounter) 'Append current record
                    End If
                Else
                    err.Clear
                End If
            End If
        Next rCounter
    End With
End Function

''' <summary>
''' Returns teh fields values to apply
''' filter on given record.
''' </summary>
''' <param name="Pattern">The expression to eval.</param>
Private Function FilterVarValues(RecIndex As Long, FieldsToFilterOn() As Long) As String
    Dim tmpData() As String
    Dim tdLB As Long
    Dim tdUB As Long
    Dim idxIter As Long
    
    tdLB = LBound(FieldsToFilterOn)
    tdUB = UBound(FieldsToFilterOn)
    ReDim tmpData(tdLB To tdUB)
    For idxIter = tdLB To tdUB
        If IsNumeric(Buffer(RecIndex)(FieldsToFilterOn(idxIter) - 1)) Then
            tmpData(idxIter) = "f" & FieldsToFilterOn(idxIter) & "=" _
                                & Buffer(RecIndex)(FieldsToFilterOn(idxIter) - 1)
        Else    'Literal strings
            If IsBoolean(CStr(Buffer(RecIndex)(FieldsToFilterOn(idxIter) - 1))) Then
                tmpData(idxIter) = "f" & FieldsToFilterOn(idxIter) & "=" _
                                & Buffer(RecIndex)(FieldsToFilterOn(idxIter) - 1)
            Else
                tmpData(idxIter) = "f" & FieldsToFilterOn(idxIter) & "=" _
                                & d_Apostrophe & Buffer(RecIndex)(FieldsToFilterOn(idxIter) - 1) & d_Apostrophe
            End If
        End If
    Next idxIter
    FilterVarValues = Join$(tmpData, ";")
End Function

Private Function IsBoolean(ByRef Expression As String) As Boolean
    IsBoolean = (LCase(Expression) = "true")
    If Not IsBoolean Then
        IsBoolean = (LCase(Expression) = "false")
    End If
End Function

''' <summary>
''' Returns the largest integral value less than or
''' equal to the specified number.
''' Equivalent to .NET Math.Floor function.
''' </summary>
Private Function Floor(value As Double) As Long
    Dim tmpResult As Long
    
    tmpResult = Fix(value)
    Floor = tmpResult + ((value <> tmpResult) And (value < 0))
End Function

''' <summary>
''' Gets the ASCII char code for a string at specified position.
''' </summary>
Private Function charAt(str As Variant, d As Long) As Long
    If LenB(str) <= 2 * d Then
        charAt = -1
    Else
        charAt = AscW(MidB$(str, d * 2 - 1, 2))
    End If
End Function

''' <summary>
''' Check sorting bounds.
''' </summary>
Private Function CheckSortingBounds(startIndex As Long, EndIndex As Long) As Boolean
    CheckSortingBounds = True
    If startIndex > CurrentIndex + 1 And startIndex <= 0 Then
        CheckSortingBounds = False
    End If
    If EndIndex > CurrentIndex + 1 And EndIndex <= 0 Then
        CheckSortingBounds = False
    End If
    If startIndex > EndIndex Then
        CheckSortingBounds = False
    End If
End Function

''' <summary>
''' Check sorting keys.
''' </summary>
Private Function CheckSortingKeys(SortingKeys As Variant) As Boolean
    Dim iCounter As Long
    Dim tmpResult As Boolean
    
    tmpResult = True
    If IsArray(SortingKeys) Then
        iCounter = LBound(SortingKeys)
        Do While iCounter <= UBound(SortingKeys) And tmpResult
            If Abs(CLng(SortingKeys(iCounter))) < LBound(Buffer(0)) + 1 Or _
                        Abs(CLng(SortingKeys(iCounter))) > UBound(Buffer(0)) + 1 _
                        Or CLng(SortingKeys(iCounter)) = 0 Then
                tmpResult = False
            End If
            iCounter = iCounter + 1
        Loop
    Else
        If Abs(SortingKeys) < LBound(Buffer(0)) + 1 Or _
                        Abs(SortingKeys) > UBound(Buffer(0)) + 1 _
                        Or CLng(SortingKeys) = 0 Then
            tmpResult = False
        End If
    End If
    CheckSortingKeys = tmpResult
End Function

''' <summary>
''' Reinitializes the current instance.
''' </summary>
Public Sub Clear()
Attribute Clear.VB_Description = "Reinitializes the current instance."
    Erase Buffer
    ReDim Buffer(0 To MaxIndex)
    CurrentIndex = -1
End Sub

''' <summary>
''' Returns a copy of the current instance.
''' </summary>
Public Function Clone() As Object
Attribute Clone.VB_Description = "Returns a copy of the current instance"
    Set Clone = Me
End Function

''' <summary>
''' Concatenates the values from the current instance with the specified values.
''' </summary>
''' <param name="AValues">A variant containing the array, arraylist or value to concatenate.</param>
Public Function Concat(ByRef aValues As Variant) As CSVArrayList
Attribute Concat.VB_Description = "Concatenates the values from the current instance with the specified values."
    On Error GoTo Concat_ErrHandler
    Dim tmpValues As Variant
    
    If TypeName(aValues) = TypeName(Me) Then
        tmpValues = aValues.items
    Else
        tmpValues = aValues
    End If
    If IsArray(tmpValues) Then
        Dim Dim1Pointer As Long
        Dim Dim2Pointer As Long
        Dim tmpRow() As Variant
        
        If MultiDimensional(tmpValues) Then '2D array expected
            ReDim tmpRow(LBound(tmpValues, 2) To UBound(tmpValues, 2))
            For Dim1Pointer = LBound(tmpValues) To UBound(tmpValues)
                For Dim2Pointer = LBound(tmpValues, 2) To UBound(tmpValues, 2)
                    tmpRow(Dim1Pointer, Dim2Pointer) = tmpValues(Dim1Pointer, Dim2Pointer)
                Next Dim2Pointer
                Add tmpRow
            Next Dim1Pointer
        Else 'Jagged or 1D array expected
            If IsJaggedArray(tmpValues) Then
                For Dim1Pointer = LBound(tmpValues) To UBound(tmpValues)
                    Add tmpValues(Dim1Pointer)
                Next Dim1Pointer
            Else
                Add tmpValues
            End If
        End If
    Else
        Add2 tmpValues
    End If
    Set Concat = Me
Concat_ErrHandler:
End Function

''' <summary>
''' Concatenates the values from the current instance with the specified values.
''' CSV generated storage expected
''' </summary>
''' <param name="AValues">Containing an arraylist or value to concatenate.</param>
Public Function Concat2(ByRef aValues As CSVArrayList) As CSVArrayList
    On Error GoTo Concat_ErrHandler
    Dim tmpValues As Variant
    Dim jc2 As Long
    
    tmpValues = aValues.items
    For jc2 = LBound(tmpValues) To UBound(tmpValues)
        Add tmpValues(jc2)
    Next jc2
    Set Concat2 = Me
Concat_ErrHandler:
End Function

''' <summary>
''' Returns a copy of the current instance.
''' </summary>
''' <param name="StartIndex">The index in which the operation will start.</param>
''' <param name="EndIndex">The index in which the operation will end.</param>
Public Function Copy(Optional startIndex As Long = 0, _
                        Optional EndIndex As Long = -1) As CSVArrayList
Attribute Copy.VB_Description = "Returns a copy of the current instance."
    Dim c As Long
    
    Set Copy = New CSVArrayList
    If EndIndex = -1 Then
        EndIndex = CurrentIndex
    End If
    Copy.Reinitialize EndIndex - startIndex + 1
    For c = startIndex To EndIndex
        Copy.Add Buffer(c)
    Next c
End Function

''' <summary>
''' Dumps content from current instance to a one-dimensional or jagged array.
''' </summary>
''' <param name="StartIndex">The index in which the operation will start.</param>
''' <param name="EndIndex">The index in which the operation will end.</param>
Public Function CopyToArray(Optional startIndex As Long = 0, _
                        Optional EndIndex As Long = -1) As Variant()
Attribute CopyToArray.VB_Description = "Returns a copy of the current instance."
    Dim c As Long
    Dim tmpResult() As Variant
    
    If EndIndex = -1 Then
        EndIndex = CurrentIndex
    End If
    ReDim tmpResult(0 To EndIndex - startIndex)
    For c = startIndex To EndIndex
        tmpResult(c) = Buffer(c)
    Next c
    CopyToArray = tmpResult
End Function

''' <summary>
''' Sorts the data using Heaps.
''' Requires all elements to be one-dimensional arrays.
''' </summary>
''' <param name="leftt">The index of the first item to be sorted.</param>
''' <param name="rightt">The index of the last item to be sorted.</param>
''' <param name="SortingKey">The column/key for the logical comparisons.</param>
''' <param name="Descending">Sort order.</param>
Private Sub heapSort(leftt As Long, _
                        rightt As Long, _
                        SortingKey As Long, _
                        Optional Descending As Boolean = False)
    Dim HeapSize As Long
    Dim i As Long
    
    HeapSize = rightt - leftt + 1           'heap-size[A]
    If Descending Then
        build_min_heap leftt, SortingKey, HeapSize
        For i = HeapSize To 2 Step -1
            Swap leftt, leftt + i - 1
            HeapSize = HeapSize - 1
            min_heapify 1, leftt, SortingKey, HeapSize
        Next i
    Else
        build_max_heap leftt, SortingKey, HeapSize
        For i = HeapSize To 2 Step -1
            Swap leftt, leftt + i - 1
            HeapSize = HeapSize - 1
            max_heapify 1, leftt, SortingKey, HeapSize
        Next i
    End If
End Sub

''' <summary>
''' Inserts an Item, at the given Index, in the current instance of the class.
''' </summary>
''' <param name="Index">The index into which the Item'll be inserted.</param>
''' <param name="aValue">The value to be inserted.</param>
Public Sub Insert(Index As Long, AValue As Variant)
Attribute Insert.VB_Description = "Inserts an Item, at the given Index, in the current instance of the class."
    Dim tmpCopy() As Variant
    Dim iCounter As Long
        
    Select Case Index
        Case 0 To CurrentIndex + 1 'Avoids to leave empty items
            Me.Add AValue
            'Checks if the item need to be placed on a previous Index
            If Index < CurrentIndex Then
                tmpCopy() = Buffer
                Erase Buffer
                ReDim Buffer(0 To MaxIndex)
                For iCounter = 0 To Index - 1
                    Buffer(iCounter) = tmpCopy(iCounter)
                    tmpCopy(iCounter) = Null
                Next iCounter
                Buffer(Index) = tmpCopy(CurrentIndex)
                For iCounter = Index To CurrentIndex - 1
                    Buffer(iCounter + 1) = tmpCopy(iCounter)
                    tmpCopy(iCounter) = Null
                Next iCounter
                Erase tmpCopy
            End If
        Case Else
            err.Raise 9
    End Select
End Sub

''' <summary>
''' Sorts tiny array data.
''' </summary>
''' <param name="leftt">The index of the first item to be sorted.</param>
''' <param name="rightt">The index of the last item to be sorted.</param>
''' <param name="SortingKeys">The columns/keys for the logical comparisons.</param>
''' <param name="Descending">Sort order.</param>
Private Sub InsertionSort(leftt As Long, _
                                rightt As Long, _
                                SortingKey As Long, _
                                Descending As Boolean)
    Dim i As Long, j As Long
    
    i = leftt + 1
    Select Case Descending
        Case False
            Do While i <= rightt
                j = i
                Do While j > leftt
                    If Buffer(j)(SortingKey) < Buffer(j - 1)(SortingKey) Then
                        Swap j, j - 1
                    End If
                    j = j - 1
                Loop
                i = i + 1
            Loop
        Case Else
            Do While i <= rightt
                j = i
                Do While j > leftt
                    If Buffer(j)(SortingKey) > Buffer(j - 1)(SortingKey) Then
                        Swap j, j - 1
                    End If
                    j = j - 1
                Loop
                i = i + 1
            Loop
    End Select
End Sub

''' <summary>
''' Returns True if the paseed argument is a jagged array.
''' </summary>
''' <param name="Arr">The array to check.</param>
Public Function IsJaggedArray(Arr As Variant) As Boolean
Attribute IsJaggedArray.VB_Description = "Returns True if the paseed argument is a jagged array."
    On Error GoTo IsJaggedArray_Err_Handler
    If IsArray(Arr) Then
        If Not MultiDimensional(Arr) Then
            Dim BoundingTest As Long
            BoundingTest = LBound(Arr(LBound(Arr)))
            IsJaggedArray = True
        End If
    End If
    Exit Function
IsJaggedArray_Err_Handler:
    IsJaggedArray = False
End Function

''' <summary>
''' Turns a jagged array into a two dim array.
''' </summary>
''' <param name="Arr">The array to check.</param>
Public Sub JaggedToTwoDimArray(ByRef JaggedArray() As Variant, ByRef TwoDimArray() As Variant)
Attribute JaggedToTwoDimArray.VB_Description = "Turns a jagged array into a two dim array."
    Dim UBj1 As Long, LBj1 As Long
    Dim UBj2 As Long, LBj2 As Long
    Dim MaxDim1 As Long, MaxDim2 As Long
    Dim jgdCounter1 As Long, jgdCounter2 As Long
    Dim tmpDimension As Long, tdimCounter As Long
    
    On Error GoTo JaggedToTwoDimArray_Err_Handler
    LBj1 = LBound(JaggedArray)
    UBj1 = UBound(JaggedArray)
    MaxDim1 = Abs(UBj1 - LBj1) 'Dimension One in base 0
    MaxDim2 = Abs(UBound(JaggedArray(LBj1)) - LBound(JaggedArray(LBj1))) 'Dimension Two in base 0
    ReDim TwoDimArray(0 To MaxDim1, 0 To MaxDim2)
    For jgdCounter1 = LBj1 To UBj1
        LBj2 = LBound(JaggedArray(jgdCounter1))
        UBj2 = UBound(JaggedArray(jgdCounter1))
        tdimCounter = 0
        tmpDimension = Abs(UBj2 - LBj2)
        If tmpDimension > MaxDim2 Then
            MaxDim2 = tmpDimension
            ReDim Preserve TwoDimArray(0 To MaxDim1, 0 To MaxDim2)
        End If
        For jgdCounter2 = LBj2 To UBj2
            TwoDimArray(jgdCounter1, tdimCounter) = JaggedArray(jgdCounter1)(jgdCounter2)
            tdimCounter = tdimCounter + 1&
        Next jgdCounter2
        Erase JaggedArray(jgdCounter1) 'Free memory
    Next jgdCounter1
JaggedToTwoDimArray_Err_Handler:
End Sub

''' <summary>
''' To max heapify a subtree rooted with node [i].
''' </summary>
''' <param name="i">Node on the heap tree (base 1).</param>
''' <param name="leftt">Lower bound on the sorting  operation (base 0).</param>
''' <param name="SortingKey">The column/key for the logical comparisons.</param>
''' <param name="HeapSize">Elements to be sorted.</param>
Private Sub max_heapify(i As Long, _
                        leftt As Long, _
                        SortingKey As Long, _
                        HeapSize As Long)
    Dim largest As Long
    Dim l As Long
    Dim R As Long
    
    l = 2 * i                                   'LEFT
    R = 2 * i + 1                               'RIGHT
    largest = i
    If l <= HeapSize Then
        If Buffer(leftt + l - 1)(SortingKey) > Buffer(leftt + i - 1)(SortingKey) Then
            largest = l
        End If
    End If
    If R <= HeapSize Then
        If Buffer(leftt + R - 1)(SortingKey) > Buffer(leftt + largest - 1)(SortingKey) Then
            largest = R
        End If
    End If
    If largest <> i Then
        Swap leftt + i - 1, leftt + largest - 1
        max_heapify largest, leftt, SortingKey, HeapSize
    End If
End Sub

''' <summary>
''' Merges TimSort sorted runs.
''' </summary>
''' <param name="L">The index of the first item to be merged.</param>
''' <param name="m">The index of the middle item to be merged.</param>
''' <param name="R">The index of the last item to be merged.</param>
''' <param name="SortingKey">The column/key for the logical comparisons.</param>
''' <param name="Descending">Sort order.</param>
Private Sub merge(l As Long, m As Long, R As Long, SortingKey As Long, Descending As Boolean)
    
    If Descending Then
        If Buffer(m)(SortingKey) >= Buffer(m + 1)(SortingKey) Then
            Exit Sub
        End If
    Else
        If Buffer(m)(SortingKey) <= Buffer(m + 1)(SortingKey) Then
            Exit Sub
        End If
    End If
    Dim len1 As Long
    Dim len2 As Long
    Dim leftt() As Variant
    Dim rightt() As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    len1 = m - l + 1
    len2 = R - m
    ReDim leftt(len1 - 1)
    ReDim rightt(len2 - 1)
    For i = 0 To len1 - 1
        leftt(i) = Buffer(l + i)
    Next i
    For i = 0 To len2 - 1
        rightt(i) = Buffer(m + 1 + i)
    Next i
    i = 0
    j = 0
    k = l
    If Descending Then
        Do While i < len1 And j < len2
            If leftt(i)(SortingKey) >= rightt(j)(SortingKey) Then
                Buffer(k) = leftt(i)
                i = i + 1
            Else
                Buffer(k) = rightt(j)
                j = j + 1
            End If
            k = k + 1
        Loop
    Else
        Do While i < len1 And j < len2
            If leftt(i)(SortingKey) <= rightt(j)(SortingKey) Then
                Buffer(k) = leftt(i)
                i = i + 1
            Else
                Buffer(k) = rightt(j)
                j = j + 1
            End If
            k = k + 1
        Loop
    End If
    ' Copy remaining elements of leftt, if any
    Do While i < len1
        Buffer(k) = leftt(i)
        k = k + 1
        i = i + 1
    Loop
    ' Copy remaining element of rightt, if any
    Do While j < len2
        Buffer(k) = rightt(j)
        k = k + 1
        j = j + 1
    Loop
End Sub

''' <summary>
''' Sorts the data. Requires all elements to be one-dimensional arrays.
''' Omit pvarMirror, leftt & rightt; they are used internally during recursion.
''' Adapted from:
'''              (c) Ellis Dee, retrieved from vbforums.com on august 21,2021.
''' </summary>
''' <param name="leftt">The index of the first item to be sorted.</param>
''' <param name="rightt">The index of the last item to be sorted.</param>
''' <param name="SortingKeys">The columns/keys for the logical comparisons.</param>
''' <param name="Descending">Sort order.</param>
''' <param name="pvarMirror">Aditional sorting space.</param>
Private Sub MergeSort(leftt As Long, _
                        rightt As Long, _
                        Optional SortingKey As Long = 0, _
                        Optional Descending As Boolean = False, _
                        Optional ByRef pvarMirror As Variant)
                        
    ReDim pvarMirror(leftt To rightt)
    If Descending Then
        MergeSort_Desc leftt, rightt, SortingKey, pvarMirror
    Else
        MergeSort_Asc leftt, rightt, SortingKey, pvarMirror
    End If
End Sub

''' <summary>
''' Sorts the data in ascending order. Requires all elements to be one-dimensional arrays.
''' </summary>
Private Sub MergeSort_Asc(leftt As Long, _
                            rightt As Long, _
                            SortingKey As Long, _
                            pvarMirror As Variant)
    Dim Ulen As Long
    Dim l As Long
    Dim R As Long
    Dim O As Long
    Dim BoolSwitch As Boolean
    
    Ulen = rightt - leftt
    Select Case Ulen
        Case 0
        Case Is <= 8 'insertion sort for tiny array
            InsertionSort leftt, rightt, SortingKey, False
        Case Else
            Ulen = Ulen \ 2 + leftt
            MergeSort_Asc leftt, Ulen, SortingKey, pvarMirror
            MergeSort_Asc Ulen + 1, rightt, SortingKey, pvarMirror
            ' Merge the resulting halves
            l = leftt ' start of first (left) half
            R = Ulen + 1 ' start of second (right) half
            O = leftt ' start of output (mirror array)
            Do
                BoolSwitch = Buffer(R)(SortingKey) < Buffer(l)(SortingKey)
                If BoolSwitch Then
                    pvarMirror(O) = Buffer(R)
                    R = R + 1
                    If R > rightt Then
                        For l = l To Ulen
                            O = O + 1
                            pvarMirror(O) = Buffer(l)
                        Next
                        Exit Do
                    End If
                Else
                    pvarMirror(O) = Buffer(l)
                    l = l + 1
                    If l > Ulen Then
                        For R = R To rightt
                            O = O + 1
                            pvarMirror(O) = Buffer(R)
                        Next
                        Exit Do
                    End If
                End If
                O = O + 1
            Loop
            For O = leftt To rightt
                Buffer(O) = pvarMirror(O)
            Next
    End Select
End Sub

''' <summary>
''' Sorts the data in descending order. Requires all elements to be one-dimensional arrays.
''' </summary>
Private Sub MergeSort_Desc(leftt As Long, _
                            rightt As Long, _
                            SortingKey As Long, _
                            pvarMirror As Variant)
    Dim Ulen As Long
    Dim l As Long
    Dim R As Long
    Dim O As Long
    Dim BoolSwitch As Boolean
    
    Ulen = rightt - leftt
    Select Case Ulen
        Case 0
        Case Is <= 8 'insertion sort for tiny array
            InsertionSort leftt, rightt, SortingKey, True
        Case Else
            Ulen = Ulen \ 2 + leftt
            MergeSort_Desc leftt, Ulen, SortingKey, pvarMirror
            MergeSort_Desc Ulen + 1, rightt, SortingKey, pvarMirror
            ' Merge the resulting halves
            l = leftt ' start of first (left) half
            R = Ulen + 1 ' start of second (right) half
            O = leftt ' start of output (mirror array)
            Do
                BoolSwitch = Buffer(R)(SortingKey) > Buffer(l)(SortingKey)
                If BoolSwitch Then
                    pvarMirror(O) = Buffer(R)
                    R = R + 1
                    If R > rightt Then
                        For l = l To Ulen
                            O = O + 1
                            pvarMirror(O) = Buffer(l)
                        Next
                        Exit Do
                    End If
                Else
                    pvarMirror(O) = Buffer(l)
                    l = l + 1
                    If l > Ulen Then
                        For R = R To rightt
                            O = O + 1
                            pvarMirror(O) = Buffer(R)
                        Next
                        Exit Do
                    End If
                End If
                O = O + 1
            Loop
            For O = leftt To rightt
                Buffer(O) = pvarMirror(O)
            Next
    End Select
End Sub

''' <summary>
''' Computes the min of given values.
''' </summary>
''' <param name="a">First value.</param>
''' <param name="b">Second value.</param>
Private Function Min(a As Long, B As Long) As Long
    If B < a Then
        Min = B
    Else
        Min = a
    End If
End Function

''' <summary>
''' To min heapify a subtree rooted with node [i].
''' </summary>
''' <param name="i">Node on the heap tree (base 1).</param>
''' <param name="SortingKey">The column/key for the logical comparisons.</param>
''' <param name="HeapSize">Elements to be sorted.</param>
Private Sub min_heapify(i As Long, _
                        leftt As Long, _
                        SortingKey As Long, _
                        HeapSize As Long)
    Dim smallest As Long
    Dim l As Long
    Dim R As Long
    
    l = 2 * i                                   'LEFT
    R = 2 * i + 1                               'RIGHT
    smallest = i
    If l <= HeapSize Then
        If Buffer(leftt + l - 1)(SortingKey) < Buffer(leftt + i - 1)(SortingKey) Then
            smallest = l
        End If
    End If
    If R <= HeapSize Then
        If Buffer(leftt + R - 1)(SortingKey) < Buffer(leftt + smallest - 1)(SortingKey) Then
            smallest = R
        End If
    End If
    If smallest <> i Then
        Swap leftt + i - 1, leftt + smallest - 1
        min_heapify smallest, leftt, SortingKey, HeapSize
    End If
End Sub

''' <summary>
''' Checks if an array has more than one dimension.
''' </summary>
''' <param name="CheckArray">The array to check.</param>
Public Function MultiDimensional(CheckArray As Variant) As Boolean
Attribute MultiDimensional.VB_Description = "Checks if an array has more than one dimension."
    On Error GoTo ErrHandler_MultiDimensional

    If UBound(CheckArray, 2) > 0 Then
        MultiDimensional = True
    End If
    Exit Function
ErrHandler_MultiDimensional:
    MultiDimensional = False '1 dimension
End Function

''' <summary>
''' Sorts the data using dual-pivot QuickSort.
''' Requires all elements to be one-dimensional arrays.
''' </summary>
''' <param name="leftt">The index of the first item to be sorted.</param>
''' <param name="rightt">The index of the last item to be sorted.</param>
''' <param name="div">Number of splits (3 is optimal).</param>
''' <param name="SortingKey">The column/key for the logical comparisons.</param>
''' <param name="Descending">Sort order.</param>
''' <param name="UseIntroSort">Instrospective sorting with Heap/Insertion Sort.</param>
Private Sub QuickSort(leftt As Long, _
                        rightt As Long, _
                        div As Long, _
                        SortingKey As Long, _
                        Optional Descending As Boolean = False, _
                        Optional UseIntroSort As Boolean = False)
    If Descending Then
        QuickSort_Desc leftt, rightt, div, SortingKey, UseIntroSort
    Else
        QuickSort_Asc leftt, rightt, div, SortingKey, UseIntroSort
    End If
End Sub

''' <summary>
''' Sorts the data in ascending order. Requires all elements to be one-dimensional arrays.
''' </summary>
Private Sub QuickSort_Asc(leftt As Long, _
                        rightt As Long, _
                        div As Long, _
                        SortingKey As Long, _
                        Optional UseIntroSort As Boolean = False)
    
    Dim Ulen As Long
    
    Ulen = rightt - leftt
    If UseIntroSort Then
        REC_DEPTH = REC_DEPTH - 1
        If REC_DEPTH = 0 Then
            If Ulen > RUN Then
                heapSort leftt, rightt, SortingKey
            End If
            Exit Sub
        End If
    End If
    If Ulen < RUN Then 'insertion sort for tiny array
        InsertionSort leftt, rightt, SortingKey, False
    Else
        Dim Third As Long
        Third = Ulen / div
        '"medians"
        Dim m1 As Long
        Dim m2 As Long
        m1 = leftt + Third
        m2 = rightt - Third
        If m1 <= leftt Then m1 = leftt + 1
        If m2 >= rightt Then m2 = rightt - 1
        If Buffer(m1)(SortingKey) < Buffer(m2)(SortingKey) Then
            Swap m1, leftt
            Swap m2, rightt
        Else
            Swap m1, rightt
            Swap m2, leftt
        End If
        'pivots
        Dim pivot1 As Variant
        Dim pivot2 As Variant
        pivot1 = Buffer(leftt)(SortingKey)
        pivot2 = Buffer(rightt)(SortingKey)
        'pointers
        Dim less As Long
        Dim great As Long
        less = leftt + 1
        great = rightt - 1
        'sorting
        Dim k As Long
        k = less
        Do While k <= great
            If Buffer(k)(SortingKey) < pivot1 Then
                Swap k, less
                less = less + 1
            ElseIf Buffer(k)(SortingKey) > pivot2 Then
                Do While k < great And Buffer(great)(SortingKey) > pivot2
                    great = great - 1
                Loop
                Swap k, great
                great = great - 1
                If Buffer(k)(SortingKey) < pivot1 Then
                    Swap k, less
                    less = less + 1
                End If
            End If
            k = k + 1
        Loop
        'swaps
        Dim dist As Long
        dist = great - less
        If dist < 13 Then div = div + 1
        Swap less - 1, leftt
        Swap great + 1, rightt
        'subarrays
        QuickSort_Asc leftt, less - 2, div, SortingKey, UseIntroSort
        QuickSort_Asc great + 2, rightt, div, SortingKey, UseIntroSort
        'equal elements
        If (dist > Ulen - 13 And pivot1 <> pivot2) Then
            k = less
            Do While k <= great
                If Buffer(k)(SortingKey) = pivot1 Then
                    Swap k, less
                    less = less + 1
                ElseIf Buffer(k)(SortingKey) = pivot2 Then
                    Swap k, great
                    great = great - 1
                    If Buffer(k)(SortingKey) = pivot1 Then
                        Swap k, less
                        less = less + 1
                    End If
                End If
                k = k + 1
            Loop
        End If
        'subarray
        If pivot1 < pivot2 Then
            QuickSort_Asc less, great, div, SortingKey, UseIntroSort
        End If
    End If
End Sub

''' <summary>
''' Sorts the data in descending order. Requires all elements to be one-dimensional arrays.
''' </summary>
Private Sub QuickSort_Desc(leftt As Long, _
                        rightt As Long, _
                        div As Long, _
                        SortingKey As Long, _
                        Optional UseIntroSort As Boolean = False)
                        
    Dim Ulen As Long
    
    Ulen = rightt - leftt
    If UseIntroSort Then
        REC_DEPTH = REC_DEPTH - 1
        If REC_DEPTH = 0 Then
            If Ulen > RUN Then
                heapSort leftt, rightt, SortingKey, True
            End If
            Exit Sub
        End If
    End If
    If Ulen < RUN Then 'insertion sort for tiny array
        InsertionSort leftt, rightt, SortingKey, True
    Else
        Dim Third As Long
        Third = Ulen / div
        '"medians"
        Dim m1 As Long
        Dim m2 As Long
        m1 = leftt + Third
        m2 = rightt - Third
        If m1 <= leftt Then m1 = leftt + 1
        If m2 >= rightt Then m2 = rightt - 1
        If Buffer(m1)(SortingKey) > Buffer(m2)(SortingKey) Then
            Swap m1, leftt
            Swap m2, rightt
        Else
            Swap m1, rightt
            Swap m2, leftt
        End If
        'pivots
        Dim pivot1 As Variant
        Dim pivot2 As Variant
        pivot1 = Buffer(leftt)(SortingKey)
        pivot2 = Buffer(rightt)(SortingKey)
        'pointers
        Dim less As Long
        Dim great As Long
        less = leftt + 1
        great = rightt - 1
        'sorting
        Dim k As Long
        k = less
        Do While k <= great
            If Buffer(k)(SortingKey) > pivot1 Then
                Swap k, less
                less = less + 1
            ElseIf Buffer(k)(SortingKey) < pivot2 Then
                Do While k < great And Buffer(great)(SortingKey) < pivot2
                    great = great - 1
                Loop
                Swap k, great
                great = great - 1
                If Buffer(k)(SortingKey) > pivot1 Then
                    Swap k, less
                    less = less + 1
                End If
            End If
            k = k + 1
        Loop
        'swaps
        Dim dist As Long
        dist = great - less
        If dist < 13 Then div = div + 1
        Swap less - 1, leftt
        Swap great + 1, rightt
        'subarrays
        QuickSort_Desc leftt, less - 2, div, SortingKey, UseIntroSort
        QuickSort_Desc great + 2, rightt, div, SortingKey, UseIntroSort
        'equal elements
        If (dist > Ulen - 13 And pivot1 <> pivot2) Then
            k = less
            Do While k <= great
                If Buffer(k)(SortingKey) = pivot1 Then
                    Swap k, less
                    less = less + 1
                ElseIf Buffer(k)(SortingKey) = pivot2 Then
                    Swap k, great
                    great = great - 1
                    If Buffer(k)(SortingKey) = pivot1 Then
                        Swap k, less
                        less = less + 1
                    End If
                End If
                k = k + 1
            Loop
        End If
        'subarray
        If pivot1 > pivot2 Then
            QuickSort_Desc less, great, div, SortingKey, UseIntroSort
        End If
    End If
End Sub

''' <summary>
''' Reinitializes the current instance of the class and reserves user desired storage space.
''' </summary>
Public Sub Reinitialize(bufferSize As Long)
    Select Case bufferSize
        Case Is >= 0
            CurrentIndex = -1
            MaxIndex = bufferSize - 1
            ReDim Buffer(0 To MaxIndex)
        Case Else
            err.Raise 9
    End Select
End Sub

''' <summary>
''' Removes the Item at specified Index.
''' </summary>
''' <param name="Index">The index of the Item'll be deleted.</param>
Public Sub RemoveAt(Index As Long)
Attribute RemoveAt.VB_Description = "Removes an Item at specified Index."
    Dim tmpCopy() As Variant
    Dim iCounter As Long
        
    Select Case Index
        Case 0 To CurrentIndex
            tmpCopy() = Buffer
            Erase Buffer
            ReDim Buffer(0 To MaxIndex)
            For iCounter = 0 To Index - 1
                Buffer(iCounter) = tmpCopy(iCounter)
                tmpCopy(iCounter) = Null
            Next iCounter
            For iCounter = Index + 1 To CurrentIndex
                Buffer(iCounter - 1) = tmpCopy(iCounter)
                tmpCopy(iCounter) = Null
            Next iCounter
            Erase tmpCopy
            CurrentIndex = CurrentIndex - 1
        Case Else
            err.Raise 9
    End Select
End Sub

''' <summary>
''' Removes a range of Items starting at specified Index.
''' </summary>
''' <param name="Index">The index of the first item to be deleted.</param>
''' <param name="Count">The number of item to be deleted.</param>
Public Sub RemoveRange(Index As Long, count As Long)
Attribute RemoveRange.VB_Description = "Removes a range of Items starting at specified Index."
    Dim tmpCopy() As Variant
    Dim iCounter As Long
    Dim EndIndex As Long
        
    Select Case Index
        Case 0 To CurrentIndex
            EndIndex = Index + count - 1
            If EndIndex > CurrentIndex Then
                EndIndex = CurrentIndex
                count = EndIndex - Index + 1
            End If
            tmpCopy() = Buffer
            Erase Buffer
            ReDim Buffer(0 To MaxIndex)
            For iCounter = 0 To Index - 1
                Buffer(iCounter) = tmpCopy(iCounter)
                tmpCopy(iCounter) = Null
            Next iCounter
            For iCounter = EndIndex + 1 To CurrentIndex
                Buffer(iCounter - count) = tmpCopy(iCounter)
                tmpCopy(iCounter) = Null
            Next iCounter
            CurrentIndex = CurrentIndex - count
            Erase tmpCopy
        Case Else
            err.Raise 9
    End Select
End Sub

''' <summary>
''' Reverse the order of the internal items.
''' </summary>
''' <param name="StartIndex">The index of the first item to be reversed.</param>
''' <param name="EndIndex">The number of item to be reversed.</param>
Public Function Reverse(Optional startIndex As Long = -1, _
                        Optional EndIndex As Long = -1) As Boolean
Attribute Reverse.VB_Description = "Reverse the order of the internal items."
    Dim i As Long
    
    On Error GoTo Reverse_ErrHandler
    If startIndex < 0 Then
        startIndex = 0
    End If
    If EndIndex < 0 Then
        EndIndex = CurrentIndex
    End If
    If EndIndex > startIndex Then
        For i = startIndex To (startIndex + ((EndIndex - startIndex - 1) \ 2))
            Swap i, EndIndex - (i - startIndex)
        Next i
    End If
    Reverse = True
    Exit Function
Reverse_ErrHandler:
    Reverse = False
End Function

''' <summary>
''' Reverse the order of items in the target jagged array.
''' </summary>
''' <param name="StartIndex">The index of the first item to be reversed.</param>
''' <param name="EndIndex">The number of item to be reversed.</param>
''' <param name="TargetArr">Array to reverse in.</param>
Public Function Reverse2(ByRef TargetArr() As Variant, _
                            Optional startIndex As Long = -1, _
                            Optional EndIndex As Long = -1) As Boolean
Attribute Reverse2.VB_Description = "Reverse the order of items in the target array."
    Dim i As Long
    
    On Error GoTo Reverse_ErrHandler
    If startIndex < 0 Then
        startIndex = LBound(TargetArr)
    End If
    If EndIndex < 0 Then
        EndIndex = UBound(TargetArr)
    End If
    If EndIndex > startIndex Then
        For i = startIndex To (startIndex + ((EndIndex - startIndex - 1) \ 2))
            Swap2 TargetArr, i, EndIndex - (i - startIndex)
        Next i
    End If
    Reverse2 = True
    Exit Function
Reverse_ErrHandler:
    Reverse2 = False
End Function
''' <summary>
''' Shrinks the buffer size to avoid extra space reservation.
''' </summary>
Public Sub ShrinkBuffer()
Attribute ShrinkBuffer.VB_Description = "Shrinks the buffer size to avoid extra space reservation."
    If CurrentIndex >= 0 Then
        MaxIndex = CurrentIndex
        ReDim Preserve Buffer(0 To MaxIndex)
    End If
End Sub

''' <summary>
''' Sorts the internal Items. Requires all Items to be one-dimensional arrays.
''' </summary>
''' <param name="fromIndex">The index of the first item to be sorted.</param>
''' <param name="toIndex">The index of the last item to be sorted.</param>
''' <param name="SortingKeys">The columns/keys for the logical comparisons.</param>
''' <param name="Descending">Sort order.</param>
Public Sub Sort(Optional ByVal FromIndex As Long = -1, _
                Optional ByVal ToIndex As Long = -1, _
                Optional ByVal SortingKeys As Variant = 1, _
                Optional ByVal SortAlgorithm As SortingAlgorithms = SortingAlgorithms.SA_IntroSort)
Attribute Sort.VB_Description = "Sorts the data. Requires all elements to be one-dimensional arrays."
    On Error GoTo Sort_Error_Handler:
    If CurrentIndex + 1 > 0 Then
        If (CheckSortingBounds(FromIndex, ToIndex) And CheckSortingKeys(SortingKeys)) Then
            If FromIndex = -1 Then
                FromIndex = 0
            Else
                FromIndex = FromIndex - 1 'Zero based
            End If
            If ToIndex = -1 Then
                ToIndex = CurrentIndex
            Else
                ToIndex = ToIndex - 1
            End If
            Dim i As Long
            Dim ColumnsToSort() As Long
            If IsArray(SortingKeys) Then
                ReDim ColumnsToSort(LBound(SortingKeys) To UBound(SortingKeys))
                For i = LBound(SortingKeys) To UBound(SortingKeys)
                    ColumnsToSort(i) = CLng(SortingKeys(i))
                Next i
            Else
                ReDim ColumnsToSort(0 To 0)
                ColumnsToSort(0) = CLng(SortingKeys)
            End If
            'Sort
            keysLB = LBound(ColumnsToSort)
            keysUB = UBound(ColumnsToSort)
            If SortAlgorithm = SA_IntroSort Then 'Introspective QuickSort
                REC_DEPTH = 3 * Ceiling(Log(ToIndex - FromIndex + 1) / Log(3))
                QuickSort FromIndex, ToIndex, 3, Abs(ColumnsToSort(keysUB)) - 1, ColumnsToSort(keysUB) < 0, True
                For i = keysUB - 1 To keysLB Step -1
                    StableQSort FromIndex, ToIndex, Abs(ColumnsToSort(i)) - 1, _
                                Abs(ColumnsToSort(i + 1)) - 1, ColumnsToSort(i) < 0, True
                Next i
            Else
                If SortAlgorithm = SA_Quicksort Then 'QuickSort
                    QuickSort FromIndex, ToIndex, 3, Abs(ColumnsToSort(keysUB)) - 1, ColumnsToSort(keysUB) < 0
                    For i = keysUB - 1 To keysLB Step -1
                        StableQSort FromIndex, ToIndex, Abs(ColumnsToSort(i)) - 1, _
                                    Abs(ColumnsToSort(i + 1)) - 1, ColumnsToSort(i) < 0
                    Next i
                Else
                    If SortAlgorithm = SA_TimSort Then 'TimSort
                        For i = keysLB To keysUB
                            TimSort FromIndex, ToIndex, Abs(ColumnsToSort(i)) - 1, ColumnsToSort(i) < 0
                        Next i
                    Else
                        If SortAlgorithm = SA_HeapSort Then 'HeapSort
                            heapSort FromIndex, ToIndex, Abs(ColumnsToSort(keysUB)) - 1, ColumnsToSort(keysUB) < 0
                            For i = keysUB - 1 To keysLB Step -1
                                StableHeapSort FromIndex, ToIndex, Abs(ColumnsToSort(i)) - 1, _
                                                Abs(ColumnsToSort(i + 1)) - 1, ColumnsToSort(i) < 0
                            Next i
                        Else 'MergeSort
                            For i = keysLB To keysUB
                                MergeSort FromIndex, ToIndex, Abs(ColumnsToSort(i)) - 1, ColumnsToSort(i) < 0
                            Next i
                        End If
                    End If
                End If
            End If
        End If
    End If
    Exit Sub
Sort_Error_Handler:
End Sub

''' <summary>
''' Sorts the data in a stable way, using QuickSort with new and old sort keys.
''' The user must be sure that the data has been sorted using the
''' given OldKey, if this rule is omitted, an unexpected behavior
''' may occur. Note that all indexes are zero (0) based.
''' </summary>
''' <param name="fromIndex">The index of the first item to be sorted.</param>
''' <param name="toIndex">The index of the last item to be sorted.</param>
''' <param name="NewKey">Actual data sorting key.</param>
''' <param name="OldKey">Previous data sorting key.</param>
''' <param name="Descending">Sort order.</param>
Private Sub StableHeapSort(FromIndex As Long, _
                            ToIndex As Long, _
                            NewKey As Long, _
                            OldKey As Long, _
                            Descending As Boolean)
    Dim FIndex As Long
    Dim lIndex As Long
        
    FIndex = FromIndex
    lIndex = EndOfSortingBlock(OldKey, FIndex, ToIndex)
    Do While FIndex <= ToIndex
        heapSort FIndex, lIndex, NewKey, Descending
        FIndex = lIndex + 1
        lIndex = EndOfSortingBlock(OldKey, FIndex, ToIndex)
    Loop
End Sub

''' <summary>
''' Sorts the data in a stable way, using QuickSort with new and old sort keys.
''' The user must be sure that the data has been sorted using the
''' given OldKey, if this rule is omitted, an unexpected behavior
''' may occur. Note that all indexes are zero (0) based.
''' </summary>
''' <param name="fromIndex">The index of the first item to be sorted.</param>
''' <param name="toIndex">The index of the last item to be sorted.</param>
''' <param name="NewKey">Actual data sorting key.</param>
''' <param name="OldKey">Previous data sorting key.</param>
''' <param name="Descending">Sort order.</param>
''' <param name="UseIntroSort">Instrospective sorting with Heap/Insertion Sort.</param>
Private Sub StableQSort(FromIndex As Long, _
                        ToIndex As Long, _
                        NewKey As Long, _
                        OldKey As Long, _
                        Descending As Boolean, _
                        Optional UseIntroSort As Boolean = False)
    Dim FIndex As Long
    Dim lIndex As Long
        
    FIndex = FromIndex
    lIndex = EndOfSortingBlock(OldKey, FIndex, ToIndex)
    Do While FIndex <= ToIndex
        QuickSort FIndex, lIndex, 3, NewKey, Descending, UseIntroSort
        FIndex = lIndex + 1
        lIndex = EndOfSortingBlock(OldKey, FIndex, ToIndex)
    Loop
End Sub

''' <summary>
''' Swap items in buffer.
''' </summary>
Public Sub Swap(position1 As Long, position2 As Long)
Attribute Swap.VB_Description = "Swap items in buffer."
    If CurrentIndex >= 0 Then
        tmpBuffer = Buffer(position1)
        Buffer(position1) = Buffer(position2)
        Buffer(position2) = tmpBuffer
    End If
End Sub

''' <summary>
''' Swap items in target jagged array.
''' </summary>
Public Sub Swap2(ByRef TargetArr() As Variant, _
                    position1 As Long, position2 As Long)
Attribute Swap2.VB_Description = "Swap items in target jagged array."
    tmpBuffer = TargetArr(position1)
    TargetArr(position1) = TargetArr(position2)
    TargetArr(position2) = tmpBuffer
End Sub

''' <summary>
''' Sorts the data using Iterative Timsort.
''' Requires all elements to be one-dimensional arrays.
''' </summary>
''' <param name="leftt">The index of the first item to be sorted.</param>
''' <param name="rightt">The index of the last item to be sorted.</param>
''' <param name="SortingKey">The columns/keys for the logical comparisons.</param>
''' <param name="Descending">Sort order.</param>
Private Sub TimSort(leftt As Long, rightt As Long, SortingKey As Long, Descending As Boolean)
    Dim Size As Long
    Dim l As Long
    Dim midd As Long
    Dim R As Long
    Dim i As Long
    
    For i = leftt To rightt Step RUN
        InsertionSort i, Min(i + RUN - 1, rightt), SortingKey, Descending
    Next i
    ' Start merging from size RUN (or 32).
    ' It will merge
    ' to form size 32, then 64, 128...
    Size = RUN
    Do While Size - 1 < rightt
        For l = leftt To rightt Step 2 * Size
            midd = Min(l + Size - 1, rightt)
            R = Min(l + 2 * Size - 1, rightt)
            ' merge sub array arr[L.....midd] &
            ' arr[midd+1....R]
            If midd < R Then
                merge l, midd, R, SortingKey, Descending
            End If
        Next l
        Size = 2 * Size
    Loop
End Sub

''' <summary>
''' Turns a two-dimensional array into a jagged array.
''' </summary>
''' <param name="TwoDimArray">Array to be converted.</param>
''' <param name="JaggedArray">Output array.</param>
Public Sub TwoDimToJaggedArray(ByRef TwoDimArray() As Variant, ByRef JaggedArray() As Variant)
Attribute TwoDimToJaggedArray.VB_Description = "Turns a two-dimensional array into a jagged array."
    Dim UBj1 As Long, LBj1 As Long
    Dim UBj2 As Long, LBj2 As Long
    Dim MaxDim1 As Long, MaxDim2 As Long
    Dim jgdCounter1 As Long, jgdCounter2 As Long
    Dim tdimCounter As Long, tdimCounter2 As Long
    
    On Error GoTo TwoDimToJaggedArray_Err_Handler
    If MultiDimensional(TwoDimArray) Then
        LBj1 = LBound(TwoDimArray)
        UBj1 = UBound(TwoDimArray)
        MaxDim1 = Abs(UBj1 - LBj1) 'Dimension One in base 0
        LBj2 = LBound(TwoDimArray, 2)
        UBj2 = UBound(TwoDimArray, 2)
        MaxDim2 = Abs(UBj2 - LBj2) 'Dimension Two in base 0
        '@----------------------------------------------
        'Create the jagged array
        Call CreateJagged(JaggedArray, MaxDim1, MaxDim2)
        '@----------------------------------------------
        'Deconstruct and dump the data
        jgdCounter1 = 0
        For tdimCounter = LBj1 To UBj1
            jgdCounter2 = 0
            For tdimCounter2 = LBj2 To UBj2
                JaggedArray(jgdCounter1)(jgdCounter2) = TwoDimArray(tdimCounter, tdimCounter2)
                TwoDimArray(tdimCounter, tdimCounter2) = vbNullString
                jgdCounter2 = jgdCounter2 + 1&
            Next tdimCounter2
            jgdCounter1 = jgdCounter1 + 1&
        Next tdimCounter
    Else
        LBj1 = LBound(TwoDimArray)
        UBj1 = UBound(TwoDimArray)
        MaxDim1 = Abs(UBj1 - LBj1) 'Dimension One in base 0
        MaxDim2 = 0 'Dimension Two in base 0
        '@----------------------------------------------
        'Create the jagged array
        Call CreateJagged(JaggedArray, MaxDim1, MaxDim2)
        '@----------------------------------------------
        'Deconstruct and dump the data
        jgdCounter1 = 0
        For tdimCounter = LBj1 To UBj1
            JaggedArray(jgdCounter1)(0) = TwoDimArray(tdimCounter)
            TwoDimArray(tdimCounter) = vbNullString
            jgdCounter1 = jgdCounter1 + 1&
        Next tdimCounter
    End If
    Erase TwoDimArray
TwoDimToJaggedArray_Err_Handler:
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////
'#
