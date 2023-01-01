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
' Copyright © 2021-2023 W. García
' GPL-3.0 license | https://www.gnu.org/licenses/gpl-3.0.html/
' https://ingwilfredogarcia.wordpress.com
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' GENERAL INFO:
' Class module developed to emulate some functionalities from the ArrayList present in some
' most modern languages. The CSVArrayList serves as a container for all the data read from
' CSV files and can be used to manipulate the stored items, or to store data that does not
' come from a CSV file, according to the user's request.
'////////////////////////////////////////////////////////////////////////////////////////////
'#
Option Explicit
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' CONSTANTS:
Private Const KEY_TREE_INDICATOR As String = "#[OT in use]!"
Private Const RUN As Long = 32
Private Const SPACE_CHR As String = " "
Private Const UNDERSCORE_CHR As String = "_"
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' TYPES:
Private Type IndexedItem
    ItemValue As Variant
    itemKey As String
    ITree As CSVArrayList
End Type
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' VARIABLES:
Private Buffer() As Variant
Private IndexedBuffer() As IndexedItem
Private CurrentIndex As Long
Private IndexedCurrentIndex As Long
Private keysLB As Long
Private keysUB As Long
Private MaxIndex As Long
Private IndexedMaxIndex As Long
Private tmpBuffer As Variant
Private tmpIndexedBuffer As IndexedItem
Private SORTED_DATA As Boolean
Private LAST_SORTED_FIELD As Long
Private INDEXED_SORTED_DATA As Boolean
Private P_INDEXING As Boolean
Private P_KEYS_TREE As Boolean 'Allows to store elements with the same key in an array list
Private SORTED_KEYS As Boolean
Private NullItem As IndexedItem
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' CONSTANTS:
Private Const d_Apostrophe As String = "'"
Private Const d_comma As String = ","
Private Const d_dot As String = "."
Private Const d_semicolon As String = ";"
'////////////////////////////////////////////////////////////////////////////////////////////
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' ENUMERATIONS:
Public Enum SortingAlgorithms
    SA_Quicksort = 0
    SA_TimSort = 1
    SA_HeapSort = 2
    SA_MergeSort = 3
End Enum
Public Enum JoinType
    JT_LeftJoin = 0
    JT_RightJoin = 1
    JT_InnerJoin = 2
End Enum
Public Enum AggregateFunctions
    AF_Count = 0
    AF_Max = 1
    AF_Min = 2
    AF_Sum = 3
    AF_AVG = 4
    AF_STDEV = 5
    AF_Summary = 6
End Enum
'////////////////////////////////////////////////////////////////////////////////////////////
'#

''' <summary>
''' Initializes this instance.
''' </summary>
Private Sub Class_Initialize()
    CurrentIndex = -1&
    IndexedCurrentIndex = CurrentIndex
    MaxIndex = 1024& - 1&
    IndexedMaxIndex = MaxIndex
    ReDim Buffer(0 To MaxIndex)
    ReDim IndexedBuffer(0 To IndexedMaxIndex)
    SORTED_KEYS = False
End Sub

''' <summary>
''' Finalizes this instance.
''' </summary>
Private Sub Class_Terminate()
    ReDim Buffer(0 To 0)
    ReDim IndexedBuffer(0 To 0)
End Sub

'////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES:
'#
''' <summary>
''' Gets the amount of items stored in the current instance.
''' </summary>
Public Property Get count() As Long
Attribute count.VB_Description = "Gets the amount of items stored in the current instance"
    If Not P_INDEXING Then
        count = CurrentIndex + 1
    Else
        count = IndexedCurrentIndex + 1
    End If
End Property

''' <summary>
''' Gets all indexed Items from the current instance.
''' </summary>
Public Property Get indexedItems() As Variant
    Dim tmpResult() As Variant
    Dim iCounter As Long
    
    ReDim tmpResult(0 To IndexedCurrentIndex)
    For iCounter = 0 To IndexedCurrentIndex
        If IsArray(IndexedBuffer(iCounter).ItemValue) Then
            tmpResult(iCounter) = IndexedBuffer(iCounter).ItemValue
        Else
            If IndexedBuffer(iCounter).ItemValue <> KEY_TREE_INDICATOR Then
                tmpResult(iCounter) = IndexedBuffer(iCounter).ItemValue
            Else
                tmpResult(iCounter) = IndexedBuffer(iCounter).ITree.items
            End If
        End If
    Next iCounter
    indexedItems = tmpResult
End Property

''' <summary>
''' Gets all indexed Items from the current instance.
''' </summary>
Public Property Get keys() As String()
    Dim tmpResult() As String
    Dim iCounter As Long
    
    ReDim tmpResult(0 To IndexedCurrentIndex)
    For iCounter = 0 To IndexedCurrentIndex
        tmpResult(iCounter) = IndexedBuffer(iCounter).itemKey
    Next iCounter
    keys = tmpResult
End Property

''' <summary>
''' Gets the key at given position.
''' </summary>
Public Property Get itemKey(aIndex As Long) As String
    itemKey = IndexedBuffer(aIndex).itemKey
End Property

''' <summary>
''' Gets or sets the Item, by its index, from the current instance.
''' </summary>
''' <param name="Index">Item's position in this instance.</param>
Public Property Get item(index As Long) As Variant
Attribute item.VB_Description = "Gets the Item, by its index, from the current instance."
Attribute item.VB_UserMemId = 0

    If Not P_INDEXING Then
        Select Case index
            Case 0 To CurrentIndex
                item = Buffer(index)
            Case Else
                err.Raise 9
        End Select
    Else
        Select Case index
            Case 0 To IndexedCurrentIndex
                If Not P_KEYS_TREE Then
                    item = IndexedBuffer(index).ItemValue
                Else
                    item = IndexedBuffer(index).ITree.items
                End If
            Case Else
                err.Raise 9
        End Select
    End If
End Property

''' <summary>
''' Gets or sets the Item, by its index, from the current instance.
''' </summary>
''' <param name="Index">Item's position in this instance.</param>
''' <param name="aValue">The value to overwrite the element.</param>
Public Property Let item(index As Long, aValue As Variant)
    Select Case index
        Case 0 To CurrentIndex
            Buffer(index) = aValue
        Case Else
            err.Raise 9
    End Select
End Property

''' <summary>
''' Gets or sets the set of elements of the current instance.
''' </summary>
Public Property Get items() As Variant
Attribute items.VB_Description = "Gets an array with all the Items from the current instance."
    If Not P_INDEXING Then
        If CurrentIndex >= 0 Then
            Dim tmpResult() As Variant
            tmpResult = Buffer
            If MaxIndex <> CurrentIndex Then
                ReDim Preserve tmpResult(0 To CurrentIndex)
            End If
            items = tmpResult
        End If
    Else
        If IndexedCurrentIndex >= 0 Then
            items = indexedItems
        End If
    End If
End Property

''' <summary>
''' Gets or sets the current instance's array.
''' </summary>
Public Property Let items(aValue As Variant)
    Clear
    Add2 aValue
End Property

''' <summary>
''' Option to use this instance with indexed items.
''' </summary>
''' <param name="aOption">True for indexig.</param>
Public Property Let indexing(aOption As Boolean)
    P_INDEXING = aOption
End Property

Public Property Get indexing() As Boolean
    indexing = P_INDEXING
End Property

Public Property Get isSorted() As Boolean
    isSorted = SORTED_DATA
End Property

Public Property Get keyTree() As Boolean
    keyTree = P_KEYS_TREE
End Property

''' <summary>
''' Allows to store and group elements with the same key.
''' </summary>
''' <param name="oEnable">Set to True to create a key tree.</param>
Public Property Let keyTree(oEnable As Boolean)
    P_KEYS_TREE = oEnable
End Property

Public Property Get lastSortedIndex() As Long
    lastSortedIndex = LAST_SORTED_FIELD
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
Public Sub Add(aValue As Variant)
Attribute Add.VB_Description = "Appends a copy of the specified value to the current instance."
    On Error GoTo Expand_Buffer
    Buffer(CurrentIndex + 1) = aValue
    CurrentIndex = CurrentIndex + 1
    SORTED_DATA = False
    Exit Sub
Expand_Buffer:
    MaxIndex = 2 * (MaxIndex + 1) - 1
    ReDim Preserve Buffer(0 To MaxIndex)
    Add aValue
End Sub
''' <summary>
''' Appends a copy, in jagged array fashion, of the specified
''' values to the current instance. Not indexing support.
''' </summary>
''' <param name="aValue">The value to append.</param>
Public Sub Add2(ParamArray aValues() As Variant)
Attribute Add2.VB_Description = "Appends a copy, in jagged array fashion, of the specified value to the current instance."
    Dim i As Long, j As Long, k As Long
    Dim tmpValue() As Variant
    Dim DeltaIdx As Long
    
    For i = LBound(aValues) To UBound(aValues)
        Select Case CurrentIndex + 1
            Case Is <= MaxIndex
                If IsArray(aValues(i)) Then
                    If Is2Darray(aValues(i)) Then
                        DeltaIdx = LBound(aValues(i), 2)
                        ReDim tmpValue(0 To UBound(aValues(i), 2) - LBound(aValues(i), 2))
                        For j = LBound(aValues(i)) To UBound(aValues(i))
                            For k = LBound(aValues(i), 2) To UBound(aValues(i), 2)
                                tmpValue(k - DeltaIdx) = aValues(i)(j, k)
                            Next k
                            Add2 tmpValue
                        Next j
                    Else
                        If IsJaggedArray(aValues(i)) Then
                            CurrentIndex = CurrentIndex - 1
                            For j = LBound(aValues(i)) To UBound(aValues(i))
                                Add2 aValues(i)(j)
                            Next j
                        Else
                            Buffer(CurrentIndex + 1) = aValues(i)
                            CurrentIndex = CurrentIndex + 1
                        End If
                    End If
                Else
                    If TypeName(aValues(i)) = TypeName(Me) Then
                        For j = 0 To aValues(i).count - 1
                            Add2 aValues(i).item(j)
                        Next j
                    Else
                        ReDim tmpValue(0 To 0)
                        tmpValue(0) = aValues(i)
                        Buffer(CurrentIndex + 1) = tmpValue
                        CurrentIndex = CurrentIndex + 1
                    End If
                End If
            Case Else
                MaxIndex = 2 * (MaxIndex + 1) - 1
                ReDim Preserve Buffer(0 To MaxIndex)
                Add2 aValues(i)
        End Select
    Next i
    SORTED_DATA = False
End Sub
''' <summary>
''' Appends a copy of the specified value to the current instance's
''' indexed buffer. If the key exist, the item will be modified.
''' </summary>
''' <param name="Key">Item's key.</param>
''' <param name="iValue">The value to append.</param>
''' <param name="UpdateExistingItems">Udates items stored with the previous key.</param>
Public Sub AddIndexedItem(Key As String, iValue As Variant, _
                            Optional UpdateExistingItems As Boolean = True)
    Dim tmpItem As IndexedItem
    Dim tmpIdx As Long
    
    tmpIdx = KeyIndex(Key)
    If tmpIdx = -1 Then 'The key is not in use (append data)
        IndexedCurrentIndex = IndexedCurrentIndex + 1
        tmpIdx = IndexedCurrentIndex
        If Not P_KEYS_TREE Then
            tmpItem.itemKey = Key
            tmpItem.ItemValue = iValue
        End If
        SORTED_KEYS = False
    Else
        If Not UpdateExistingItems Then
            If Not P_KEYS_TREE Then
                err.Raise 457
                Exit Sub
            End If
        End If
    End If
    If IndexedCurrentIndex > IndexedMaxIndex Then
        IndexedMaxIndex = 2 * (IndexedMaxIndex + 1) - 1
        ReDim Preserve IndexedBuffer(0 To IndexedMaxIndex)
    End If
    If Not P_KEYS_TREE Then
        IndexedBuffer(tmpIdx) = tmpItem
    Else
        IndexedBuffer(tmpIdx).itemKey = Key
        IndexedBuffer(tmpIdx).ItemValue = KEY_TREE_INDICATOR
        If IndexedBuffer(tmpIdx).ITree Is Nothing Then
            Set IndexedBuffer(tmpIdx).ITree = New CSVArrayList
        End If
        IndexedBuffer(tmpIdx).ITree.Add iValue
    End If
    INDEXED_SORTED_DATA = False
End Sub

''' <summary>
''' Searches for an item in the internal CSV records, in a given field,
''' of the current instance when the data is already sorted. Returns
''' the index of the item when found and -1 when missing values.
''' </summary>
''' <param name="Target">The value to be searched.</param>
''' <param name="fldIndex">The field to search in (base 1).</param>
''' <param name="StartRecord">The record to start the search in (base 1).</param>
''' <param name="ReverseSearch">Backward matching. The search will end on the starting index</param>
Private Function BinarySearch(ByRef Target As Variant, fldIndex As Long, _
                                StartRecord As Long, _
                                Optional ReverseSearch As Boolean = False) As Long

    If Not SORTED_DATA Or LAST_SORTED_FIELD <> Abs(fldIndex) Then
        Sort sortingKeys:=Abs(fldIndex)
        BinarySearch = ItemIndex(Target, fldIndex, ReverseSearch:=ReverseSearch)
    Else
        Dim tmpR As Boolean
        Dim BottomIdx As Double
        Dim TopIdx As Double
        Dim MiddleIdx As Double
        Dim idx As Long
        
        idx = Abs(fldIndex) - 1
        If StartRecord - 1 <= CurrentIndex Then
            If Not ReverseSearch Then
                BottomIdx = StartRecord - 1
                TopIdx = CurrentIndex
            Else
                BottomIdx = 0
                TopIdx = StartRecord - 1
            End If
            Do
                MiddleIdx = BottomIdx + Ceiling((TopIdx - BottomIdx) / 2)
                tmpR = (Target = Buffer(MiddleIdx)(idx))
                If Target > Buffer(MiddleIdx)(idx) Then
                    BottomIdx = MiddleIdx + 1
                Else
                    TopIdx = MiddleIdx - 1
                End If
            Loop While Not tmpR And TopIdx >= BottomIdx
            If tmpR Then
                BinarySearch = MiddleIdx
            Else
                BinarySearch = -1
            End If
        Else
            BinarySearch = -1
        End If
    End If
End Function

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
''' Remove duplicates from records. Requires rectangular
''' table input (all records with same fields count)
''' </summary>
''' <param name="Keys">Indexes of the fields used for deduplication.</param>
''' <param name="headers">Indicates if the data has a header record.</param>
Public Function Dedupe(keys As String, Optional Headers As Boolean = True) As CSVArrayList
    On Error GoTo ErrHandler_Dedupe
    Dim aIndexes() As String
    aIndexes() = SplitFieldsOrderStr(keys)
    If CheckForDupIndexes(aIndexes) Then
        Dim IndexesCollection As Collection
        Dim dRecCounter As Long
            
        Set IndexesCollection = New Collection
        With IndexesCollection
            For dRecCounter = Abs(Headers) To CurrentIndex
                On Error Resume Next
                IndexesCollection.Add dRecCounter, DedupeKeysMerge(dRecCounter, aIndexes)
            Next dRecCounter
            Dim colItem As Variant
            Set Dedupe = New CSVArrayList
            For Each colItem In IndexesCollection
                Dedupe.Add Buffer(CLng(colItem))
            Next
        End With
    End If
    Exit Function
ErrHandler_Dedupe:
    Set Dedupe = Nothing
End Function
Private Function DedupeKeysMerge(aIndex As Long, keys As Variant) As String
    Dim tmpResult() As Variant
    Dim kLB As Long
    Dim kUb As Long
    Dim keyCounter As Long
    Dim tgRecord() As Variant
    
    kLB = LBound(keys)
    kUb = UBound(keys)
    ReDim tmpResult(kLB To kUb)
    tgRecord() = Buffer(aIndex)
    For keyCounter = kLB To kUb
        tmpResult(keyCounter) = tgRecord(CLng(keys(keyCounter)))
    Next keyCounter
    DedupeKeysMerge = Join$(tmpResult, "|")
End Function

''' <summary>
''' Delimits the next sorting block index.
''' </summary>
''' <param name="OldKey">The key/column used in the last sorting.</param>
''' <param name="StartPos">Sorting block start.</param>
''' <param name="MaxIndex">Last index in the sorting operation.</param>
''' <param name="Indexed">Work over indexed buffer.</param>
Private Function EndOfSortingBlock(OldKey As Long, _
                                    StartPos As Long, _
                                    MaxIndex As Long, _
                                    Optional Indexed As Boolean) As Long
    Dim i As Long
    Dim Switch As Boolean
    Dim tmpRecordIndex As Long
    
    i = StartPos
    tmpRecordIndex = i
    If Not Indexed Then
        Do While i <= MaxIndex And Not Switch
            Switch = Buffer(tmpRecordIndex)(OldKey) <> Buffer(i)(OldKey)
            i = i + (1 + Switch)
        Loop
    Else
        Do While i <= MaxIndex And Not Switch
            Switch = IndexedBuffer(tmpRecordIndex).ItemValue(OldKey) <> IndexedBuffer(i).ItemValue(OldKey)
            i = i + (1 + Switch)
        Loop
    End If
    EndOfSortingBlock = i - 1
End Function

''' <summary>
''' Returns the indexes of the fields
''' to be used in the filter operation.
''' </summary>
''' <param name="VarList">List of varibles [f1; f2; ...; fn].</param>
Private Function GetIndexesFromVarList(VarList As String, _
                                        Optional hRecord As Variant) As Long()
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
        If InStrB(1, tmpVars(jC), "f") = 0 Then
            If TypeName(hRecord) = "Variant()" Then
                Dim deltaPos As Long
                deltaPos = LBound(hRecord)
                tmpVars(jC) = "f" & CStr(GetFieldIndex(tmpVars(jC), hRecord) - deltaPos + 1)
            End If
        End If
        IdxLen = LenB(tmpVars(jC)) - (InStrB(1, tmpVars(jC), "f") + 1)
        If IdxLen Then 'At least one digit to extract
            tmpResult(jC) = CLng(MidB$(tmpVars(jC), 3))
        Else 'Input error
            ReDim tmpResult(0)
            tmpResult(0) = -1
            GetIndexesFromVarList = tmpResult
            Exit Function
        End If
    Next jC
    GetIndexesFromVarList = tmpResult
End Function

Private Function GetRearrangeSortStr(FieldArray() As Variant, SortedFieldList As CSVArrayList) As String
    Dim fldPos As Long
    Dim srgIcounter As Long
    Dim srgJcounter As Long
    Dim srgTmpBool As Boolean
    Dim tmpData() As Variant
    Dim tmpResult As CSVArrayList
    
    tmpData() = SortedFieldList.items
    Set tmpResult = New CSVArrayList
    For srgIcounter = LBound(tmpData) To UBound(tmpData)
        srgJcounter = LBound(FieldArray)
        fldPos = 0
        Do
            srgTmpBool = (CStr(tmpData(srgIcounter)(0)) = CStr(FieldArray(srgJcounter)))
            If Not srgTmpBool Then
                fldPos = fldPos + 1
            End If
            srgJcounter = srgJcounter + 1
        Loop While Not srgTmpBool And srgJcounter <= UBound(FieldArray)
        tmpResult.Add fldPos
    Next srgIcounter
    GetRearrangeSortStr = Join$(tmpResult.items, ",")
End Function

Private Function FiedlInMergeList(aIndex As Long, indexes() As String) As Boolean
    Dim iDxFld As Long
    Dim tmpBool As Boolean
    Dim IdxUprB As Long
    
    iDxFld = LBound(indexes)
    IdxUprB = UBound(indexes)
    Do
        tmpBool = CDbl(indexes(iDxFld)) = aIndex
        iDxFld = iDxFld + 1
    Loop While iDxFld <= IdxUprB And Not tmpBool
    FiedlInMergeList = tmpBool
End Function

Private Function FieldsToOperate(fieldsString As String) As String()
    FieldsToOperate = SplitFieldsOrderStr(fieldsString)
End Function

Private Function FielsToMergeAreValid(indexes() As String) As Boolean
    Dim iDxFld As Long
    Dim tmpBool As Boolean
    Dim UprB As Long
    Dim IdxUprB As Long
    
    iDxFld = LBound(indexes)
    UprB = UBound(Buffer(0))
    IdxUprB = UBound(indexes)
    Do
        tmpBool = CDbl(indexes(iDxFld)) >= 0 And CDbl(indexes(iDxFld)) <= UprB
        iDxFld = iDxFld + 1
    Loop While iDxFld <= IdxUprB And tmpBool
    FielsToMergeAreValid = tmpBool
End Function

''' <summary>
''' Returns a filtered array list using the
''' CSVexpressions class module.
''' </summary>
''' <param name="Pattern">The expression to eval.</param>
''' <param name="StartIndex">The index to start filter records (base 1).</param>
''' <param name="Exclude"> Exclude records matching "Pattern" if set to "True".</param>
''' <param name="Headers"> Indicates if the data has header record.</param>
Public Function Filter(Pattern As String, startIndex As Long, _
                        Optional Exclude As Boolean = False, _
                        Optional Headers As Boolean = True) As CSVArrayList
    Dim cPattern As String
    Dim endIndex As Long
    Dim Evaluator As CSVexpressions
    Dim evalWithOutVar As Boolean
    Dim rCounter As Long
    Dim TargetFields() As Long
    
    If P_INDEXING Then Exit Function 'Does not work over indexed buffers
    If startIndex < 1 Then Exit Function
    Set Filter = New CSVArrayList
    Set Evaluator = New CSVexpressions
    endIndex = CurrentIndex
    With Evaluator
        If Headers Then
            .Create SwitchUnderscoresAndSpaces(Pattern, Buffer(0))
            evalWithOutVar = (.currentVariables = vbNullString)
            If Not evalWithOutVar Then
                TargetFields() = GetIndexesFromVarList(SwitchUnderscoresAndSpaces(.currentVariables, Buffer(0), ToUnderscores:=False), Buffer(0))
                'Replace
                cPattern = RebuildEvalExpression(.expression, .currentVariables, TargetFields)
                .Create cPattern
            End If
        Else
            .Create Pattern
            evalWithOutVar = (.currentVariables = vbNullString)
            If Not evalWithOutVar Then
                TargetFields() = GetIndexesFromVarList(.currentVariables)
            End If
        End If
        For rCounter = startIndex - 1 To endIndex
            If Not evalWithOutVar Then
                On Error Resume Next
                .Eval GetValuesForVariables(rCounter, TargetFields)
            Else
                On Error Resume Next
                .Eval
            End If
            If .errorType = ExpressionErrors.errNone Then
                If err.Number = 0 Then
                    If CBool(.result) Then
                        If Not Exclude Then
                            Filter.Add Buffer(rCounter) 'Append current record
                        End If
                    Else
                         If Exclude Then
                            Filter.Add Buffer(rCounter) 'Append current record
                        End If
                    End If
                Else
                    err.Clear
                End If
            End If
        Next rCounter
    End With
End Function

''' <summary>
''' Returns the fields values to operate the given record.
''' </summary>
''' <param name="Pattern">The expression to eval.</param>
Private Function GetValuesForVariables(RecIndex As Long, FieldsIndexes() As Long, _
                                Optional useInternalBuffer As Boolean = True, _
                                Optional ByRef tRecord As Variant) As String
    If useInternalBuffer Then
        If Not P_INDEXING Then
            GetValuesForVariables = JoinVariablesAndValues(Buffer(RecIndex), FieldsIndexes)
        Else
            GetValuesForVariables = JoinVariablesAndValues(tRecord, FieldsIndexes)
        End If
    Else
        GetValuesForVariables = JoinVariablesAndValues(tRecord, FieldsIndexes)
    End If
End Function

Private Function IsBoolean(ByRef expression As String) As Boolean
    IsBoolean = (LCase(expression) = "true")
    If Not IsBoolean Then
        IsBoolean = (LCase(expression) = "false")
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

Private Function Ceiling(ByRef value As Double) As Double
    Dim tmpResult As Double
    
    tmpResult = Fix(value)
    Ceiling = tmpResult - ((value <> tmpResult) And (value > 0))
End Function

Private Function CheckForDupIndexes(aIndexes As Variant) As Boolean
    Dim aLB As Long
    Dim tmpCol As Collection
    
    Set tmpCol = New Collection
    On Error GoTo err_handler
    For aLB = LBound(aIndexes) To UBound(aIndexes)
        tmpCol.Add aIndexes(aLB), CStr(aIndexes(aLB))
    Next aLB
    CheckForDupIndexes = True
    Exit Function
err_handler:
    CheckForDupIndexes = False
End Function

''' <summary>
''' Check sorting bounds.
''' </summary>
Private Function CheckSortingBounds(startIndex As Long, endIndex As Long, _
                                    Optional Indexed As Boolean = False) As Boolean
    CheckSortingBounds = True
    If Not Indexed Then
        If startIndex > CurrentIndex + 1 And startIndex <= 0 Then
            CheckSortingBounds = False
        End If
        If endIndex > CurrentIndex + 1 And endIndex <= 0 Then
            CheckSortingBounds = False
        End If
        If startIndex > endIndex Then
            CheckSortingBounds = False
        End If
    Else
        If startIndex > IndexedCurrentIndex + 1 And startIndex <= 0 Then
            CheckSortingBounds = False
        End If
        If endIndex > IndexedCurrentIndex + 1 And endIndex <= 0 Then
            CheckSortingBounds = False
        End If
        If startIndex > endIndex Then
            CheckSortingBounds = False
        End If
    End If
End Function

''' <summary>
''' Check sorting keys.
''' </summary>
Private Function CheckSortingKeys(sortingKeys As Variant, Optional Indexed As Boolean = False) As Boolean
    Dim iCounter As Long
    Dim tmpResult As Boolean
    
    tmpResult = True
    If IsArray(sortingKeys) Then
        iCounter = LBound(sortingKeys)
        Do While iCounter <= UBound(sortingKeys) And tmpResult
            If Not Indexed Then
                If Abs(CLng(sortingKeys(iCounter))) < LBound(Buffer(0)) + 1 Or _
                            Abs(CLng(sortingKeys(iCounter))) > UBound(Buffer(0)) + 1 _
                            Or CLng(sortingKeys(iCounter)) = 0 Then
                    tmpResult = False
                End If
            Else
                If Abs(CLng(sortingKeys(iCounter))) < LBound(IndexedBuffer(0).ItemValue) + 1 Or _
                            Abs(CLng(sortingKeys(iCounter))) > UBound(IndexedBuffer(0).ItemValue) + 1 _
                            Or CLng(sortingKeys(iCounter)) = 0 Then
                    tmpResult = False
                End If
            End If
            iCounter = iCounter + 1
        Loop
    Else
        If Not Indexed Then
            If Abs(sortingKeys) < LBound(Buffer(0)) + 1 Or _
                            Abs(sortingKeys) > UBound(Buffer(0)) + 1 _
                            Or CLng(sortingKeys) = 0 Then
                tmpResult = False
            End If
        Else
            If Abs(sortingKeys) < LBound(IndexedBuffer(0).ItemValue) + 1 Or _
                            Abs(sortingKeys) > UBound(IndexedBuffer(0).ItemValue) + 1 _
                            Or CLng(sortingKeys) = 0 Then
                tmpResult = False
            End If
        End If
    End If
    CheckSortingKeys = tmpResult
End Function

''' <summary>
''' Reinitializes the current instance.
''' </summary>
Public Sub Clear()
Attribute Clear.VB_Description = "Reinitializes the current instance."
    If Not P_INDEXING Then
        Erase Buffer
        ReDim Buffer(0 To MaxIndex)
        CurrentIndex = -1
        SORTED_DATA = False
    Else
        ReDim IndexedBuffer(0 To IndexedMaxIndex)
        IndexedCurrentIndex = -1
        INDEXED_SORTED_DATA = False
    End If
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
    
    If P_INDEXING Then Exit Function 'Not available for indexed lists
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
    SORTED_DATA = False
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
    
    If P_INDEXING Then Exit Function 'Not available for indexed lists
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
                        Optional endIndex As Long = -1) As CSVArrayList
Attribute Copy.VB_Description = "Returns a copy of the current instance."
    Dim c As Long
    
    Set Copy = New CSVArrayList
    If Not P_INDEXING Then
        If endIndex = -1 Then
            endIndex = CurrentIndex
        End If
        Copy.Reinitialize endIndex - startIndex + 1
        For c = startIndex To endIndex
            Copy.Add Buffer(c)
        Next c
    Else
        If endIndex = -1 Then
            endIndex = IndexedCurrentIndex
        End If
        Copy.Reinitialize endIndex - startIndex + 1
        For c = startIndex To endIndex
            If Not P_KEYS_TREE Then
                Copy.AddIndexedItem IndexedBuffer(c).itemKey, IndexedBuffer(c).ItemValue
            Else
                Copy.AddIndexedItem IndexedBuffer(c).itemKey, IndexedBuffer(c).ITree.items
            End If
        Next c
    End If
End Function

''' <summary>
''' Dumps content from current instance to a one-dimensional or jagged array.
''' </summary>
''' <param name="StartIndex">The index in which the operation will start.</param>
''' <param name="EndIndex">The index in which the operation will end.</param>
Public Function CopyToArray(Optional startIndex As Long = 0, _
                        Optional endIndex As Long = -1) As Variant()
Attribute CopyToArray.VB_Description = "Returns a copy of the current instance."
    Dim c As Long
    Dim tmpResult() As Variant
    
    If Not P_INDEXING Then
        If endIndex = -1 Then
            endIndex = CurrentIndex
        End If
        ReDim tmpResult(0 To endIndex - startIndex)
        For c = startIndex To endIndex
        tmpResult(c) = Buffer(c)
        Next c
    Else
        If endIndex = -1 Then
            endIndex = IndexedCurrentIndex
        End If
        ReDim tmpResult(0 To endIndex - startIndex)
        For c = startIndex To endIndex
            If Not P_KEYS_TREE Then
                tmpResult(c) = IndexedBuffer(c).ItemValue
            Else
                tmpResult(c) = IndexedBuffer(c).ITree.items
            End If
        Next c
    End If
    CopyToArray = tmpResult
End Function

''' <summary>
''' Gets an indexed Item, by its key, from the current instance.
''' </summary>
''' <param name="Key">Item's key.</param>
Public Function GetIndexedItem(Key As String) As Variant
    If Key <> vbNullString Then
        Dim tIndex As Long
            
        tIndex = KeyIndex(Key)
        If tIndex > -1 Then
            If IsArray(IndexedBuffer(tIndex).ItemValue) Then
                GetIndexedItem = IndexedBuffer(tIndex).ItemValue
            Else
                If IndexedBuffer(tIndex).ItemValue <> KEY_TREE_INDICATOR Then
                    GetIndexedItem = IndexedBuffer(tIndex).ItemValue
                Else
                    GetIndexedItem = IndexedBuffer(tIndex).ITree.items
                End If
            End If
        End If
    End If
End Function

Private Function GetGroupFieldIndex(strFieldID As String, _
                                    Headers As Boolean, _
                                    headerRecord As Variant, _
                                    deltaPosition As Long) As Long
    If Not IsNumeric(strFieldID) Then
        If Headers Then
            GetGroupFieldIndex = GetFieldIndex(strFieldID, headerRecord) - deltaPosition + 1
        Else
            GetGroupFieldIndex = 1 'Base 1 input
        End If
    Else
        GetGroupFieldIndex = Abs(strFieldID)
    End If
End Function

''' <summary>
''' Groups rows having the same values into summary rows
''' </summary>
''' <param name="Table">Target table to be grouped.</param>
''' <param name="By">Position (1 to N) or name of the field to be grouped.</param>
''' <param name="OperateOn">Position (1 to N) or name of the field to be aggredated.</param>
''' <param name="Alias">Header for the output result field.</param>
''' <param name="AggregateFunct">The aggregate function to be used.</param>
Public Function Group(Table As CSVArrayList, By As Variant, OperateOn As Variant, _
                        Optional Alias As String = vbNullString, _
                        Optional AggregateFunct As AggregateFunctions = AggregateFunctions.AF_Summary, _
                        Optional Headers As Boolean = True) As CSVArrayList
    On Error GoTo grop_errHandler
    Dim deltaPos As Long
    Dim gAvg As Double
    Dim gCount As Long
    Dim gMax As Double
    Dim gMin As Double
    Dim GroupFIdx As Long
    Dim gSum As Double
    Dim iCounter As Long
    Dim itemsCount As Long
    Dim OperateFidx As Long
    Dim Switch As Boolean
    Dim tmpCompItem As Variant
    Dim tmpData As Variant
    Dim tmpNumItem As Double
    Dim tmpRow() As Variant
    Dim tmpRow_T() As Variant
    
    If Table.indexing Then Exit Function 'Not available for indexed lists
    
    If AggregateFunct <> AggregateFunctions.AF_Summary Then
        ReDim tmpRow_T(0 To 1) 'Records will have 2 fields
    Else
        ReDim tmpRow_T(0 To 5) 'Records will have 6 fields
    End If
    deltaPos = LBound(Table.item(0))
    tmpRow = tmpRow_T
    Set Group = New CSVArrayList
    'Save fields indexes
    GroupFIdx = GetGroupFieldIndex(CStr(By), True, Table.item(0), deltaPos)
    OperateFidx = GetGroupFieldIndex(CStr(OperateOn), True, Table.item(0), deltaPos)
    'Save first row
    If Headers Then
        tmpRow(0) = Table.item(0)(GroupFIdx + deltaPos - 1)
        If AggregateFunct <> AggregateFunctions.AF_Summary Then
            If Alias <> vbNullString Then
                tmpRow(1) = Alias
            Else
                tmpRow(1) = "Aggregated Col"
            End If
        Else
            tmpRow(1) = Alias & " Count"
            tmpRow(2) = Alias & " Sum"
            tmpRow(3) = Alias & " Avg"
            tmpRow(4) = Alias & " Min"
            tmpRow(5) = Alias & " Max"
        End If
        Group.Add tmpRow
    End If
    'Sort the data if needed
    If Table.isSorted Then
        If LAST_SORTED_FIELD <> GroupFIdx Then 'Sorting required
            tmpData = Table.Copy.Sort(fromIndex:=1 + Abs(Headers), sortingKeys:=GroupFIdx + deltaPos).items
        Else
            tmpData = Table.Copy.items
        End If
    Else
        tmpData = Table.Copy.Sort(fromIndex:=1 + Abs(Headers), sortingKeys:=GroupFIdx + deltaPos).items
    End If
    'Compute agregates
    If Headers Then
        iCounter = LBound(tmpData) + 1 'Headers worked yet
    Else
        iCounter = LBound(tmpData)
    End If
    itemsCount = UBound(tmpData)
    Do
        tmpRow = tmpRow_T 'Clear data
        tmpCompItem = tmpData(iCounter)(GroupFIdx + deltaPos - 1)
        gSum = 0
        gCount = 0
        gAvg = 0
        gMin = 0
        gMax = 0
        Do
            If IsNumeric(tmpData(iCounter)(OperateFidx + deltaPos - 1)) Then
                tmpNumItem = CDbl(tmpData(iCounter)(OperateFidx + deltaPos - 1))
                If gMin = 0 Then
                    gMin = tmpNumItem
                Else
                    If tmpNumItem < gMin Then
                        gMin = tmpNumItem
                    End If
                End If
                If gMax = 0 Then
                    gMax = tmpNumItem
                Else
                    If tmpNumItem > gMax Then
                        gMax = tmpNumItem
                    End If
                End If
                gSum = gSum + tmpNumItem
            End If
            gCount = gCount + 1
            iCounter = iCounter + 1
            If iCounter <= itemsCount Then
                Switch = (tmpCompItem <> tmpData(iCounter)(GroupFIdx + deltaPos - 1))
            Else
                Switch = True
            End If
        Loop While Not Switch
        gAvg = gSum / gCount
        tmpRow(0) = tmpCompItem
        Select Case AggregateFunct
            Case AggregateFunctions.AF_AVG
                tmpRow(1) = gAvg
            Case AggregateFunctions.AF_Count
                tmpRow(1) = gCount
            Case AggregateFunctions.AF_Max
                tmpRow(1) = gMax
            Case AggregateFunctions.AF_Min
                tmpRow(1) = gMin
            Case AggregateFunctions.AF_Sum
                tmpRow(1) = gSum
            Case AggregateFunctions.AF_Summary
                tmpRow(1) = gCount
                tmpRow(2) = gSum
                tmpRow(3) = gAvg
                tmpRow(4) = gMin
                tmpRow(5) = gMax
        End Select
        Group.Add tmpRow
    Loop While iCounter <= itemsCount
    Exit Function
grop_errHandler:
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
Public Sub Insert(index As Long, aValue As Variant)
    Dim iCounter As Long
        
    Select Case index
        Case 0 To CurrentIndex + 1 'Avoids to leave empty items
            Add aValue
            'Checks if the item need to be placed on a previous Index
            If index < CurrentIndex Then
                iCounter = CurrentIndex
                Do
                    Swap iCounter, iCounter - 1
                    iCounter = iCounter - 1
                Loop While iCounter > index
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
''' <param name="Descending">Sort order.</param>
''' <param name="Indexed">Sort indexed buffer.</param>
Private Sub InsertionSort(leftt As Long, _
                                rightt As Long, _
                                SortingKey As Long, _
                                Descending As Boolean, _
                                Optional Indexed As Boolean = False, _
                                Optional SortKeys As Boolean = False)
    Dim i As Long, j As Long
    
    i = leftt + 1
    If Not SortKeys Then
        Select Case Descending
            Case False
                If Not Indexed Then
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
                Else
                    Do While i <= rightt
                        j = i
                        Do While j > leftt
                            If IndexedBuffer(j).ItemValue(SortingKey) < IndexedBuffer(j - 1).ItemValue(SortingKey) Then
                                Swap j, j - 1, True
                            End If
                            j = j - 1
                        Loop
                        i = i + 1
                    Loop
                End If
            Case Else
                If Not Indexed Then
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
                Else
                    Do While i <= rightt
                        j = i
                        Do While j > leftt
                            If IndexedBuffer(j).ItemValue(SortingKey) > IndexedBuffer(j - 1).ItemValue(SortingKey) Then
                                Swap j, j - 1, True
                            End If
                            j = j - 1
                        Loop
                        i = i + 1
                    Loop
                End If
        End Select
    Else 'Sort keys in ascending order
        Do While i <= rightt
            j = i
            Do While j > leftt
                If IndexedBuffer(j).itemKey < IndexedBuffer(j - 1).itemKey Then
                    Swap j, j - 1, True
                End If
                j = j - 1
            Loop
            i = i + 1
        Loop
    End If
End Sub

''' <summary>
''' Returns True if the paseed argument is a jagged array.
''' </summary>
''' <param name="Arr">The array to check.</param>
Public Function IsJaggedArray(arr As Variant) As Boolean
Attribute IsJaggedArray.VB_Description = "Returns True if the paseed argument is a jagged array."
    On Error GoTo IsJaggedArray_Err_Handler
    If IsArray(arr) Then
        If Not MultiDimensional(arr) Then
            Dim BoundingTest As Long
            BoundingTest = LBound(arr(LBound(arr)))
            IsJaggedArray = True
        End If
    End If
    Exit Function
IsJaggedArray_Err_Handler:
    IsJaggedArray = False
End Function

Private Function Is2Darray(arr As Variant) As Boolean
    Dim d As Long
    
    On Error GoTo err_handler
    d = UBound(arr, 2)
    Is2Darray = True
    Exit Function
err_handler:
    Is2Darray = False
End Function

''' <summary>
''' Inserts a new field with name FieldName into the records of the
''' current instance at the given index. If a formula is given, the
''' field is populated in each record (row) with the result of evaluating
''' the formula on each field.
''' </summary>
''' <param name="aIndex">The index of the field to insert.</param>
''' <param name="FieldName">Name of the new field.</param>
''' <param name="Headers">Indicates whether the data have headers.</param>
''' <param name="Formula">The formula to compute the new field values.</param>
Public Function InsertField(aIndex As Long, _
                        Optional FieldName As String = vbNullString, _
                        Optional Headers As Boolean = True, _
                        Optional Formula As String = vbNullString) As CSVArrayList
    On Error GoTo ErrHandler_InsertField
    Select Case aIndex
        Case 0 To UBound(Buffer(0)) + 1
            Dim cFormula As String
            Dim cpRecord() As Variant
            Dim cpRecordBK() As Variant
            Dim curRecord() As Variant
            Dim Evaluator As CSVexpressions
            Dim evalWithOutVar As Boolean
            Dim fCounter As Long
            Dim fldCount As Long
            Dim rCounter As Long
            Dim TargetFields() As Long
            
            fldCount = UBound(Buffer(0))
            '@--------------------------------------------------------------------------------
            'Reserve storage
            ReDim cpRecord(0 To fldCount + 1)
            cpRecordBK() = cpRecord
            Set Evaluator = New CSVexpressions
            If Formula <> vbNullString Then
                With Evaluator
                    .formatResult = True
                    .Create SwitchUnderscoresAndSpaces(Formula, Buffer(0))
                    evalWithOutVar = (.currentVariables = vbNullString)
                    If Not evalWithOutVar Then
                        TargetFields() = GetIndexesFromVarList( _
                                                                SwitchUnderscoresAndSpaces(.currentVariables, Buffer(0), ToUnderscores:=False), _
                                                                Buffer(0))
                        'Replace
                        cFormula = RebuildEvalExpression(.expression, .currentVariables, TargetFields)
                        .Create cFormula
                    End If
                End With
            End If
            For rCounter = 0 To CurrentIndex
                curRecord() = Buffer(rCounter)
                If Headers Then
                    If FieldName <> vbNullString Then
                        If rCounter = 0 Then
                            '@--------------------------------------------------------------------------------
                            'Add header name
                            cpRecord(aIndex) = FieldName
                        End If
                    End If
                End If
                For fCounter = 0 To fldCount
                    If fCounter < aIndex Then
                        cpRecord(fCounter) = curRecord(fCounter)
                    Else
                        cpRecord(fCounter + 1) = curRecord(fCounter)
                    End If
                Next fCounter
                If Formula <> vbNullString Then
                    If Headers Then
                        If rCounter > 0 Then
                            If Not evalWithOutVar Then
                                On Error Resume Next
                                cpRecord(aIndex) = FormatEvalOutput(Evaluator.Eval(GetValuesForVariables(rCounter, TargetFields)))
                            Else
                                cpRecord(aIndex) = FormatEvalOutput(Evaluator.Eval())
                            End If
                        End If
                    Else
                        If Not evalWithOutVar Then
                            On Error Resume Next
                            cpRecord(aIndex) = FormatEvalOutput(Evaluator.Eval(GetValuesForVariables(rCounter, TargetFields)))
                        Else
                            cpRecord(aIndex) = FormatEvalOutput(Evaluator.Eval())
                        End If
                    End If
                End If
                Buffer(rCounter) = cpRecord
                cpRecord = cpRecordBK()
            Next rCounter
        Case Else
            err.Raise 9
    End Select
    Set InsertField = Me
    Exit Function
ErrHandler_InsertField:
    Set InsertField = Nothing
    err.Clear
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

Private Function JoinVariablesAndValues(ByRef tRecord As Variant, FieldsToJoin() As Long) As String
    Dim tmpData() As String
    Dim tdLB As Long
    Dim tdUB As Long
    Dim idxIter As Long
    Dim DeltaIdx As Long
    
    tdLB = LBound(FieldsToJoin)
    tdUB = UBound(FieldsToJoin)
    DeltaIdx = LBound(tRecord)
    ReDim tmpData(tdLB To tdUB)
    
    For idxIter = tdLB To tdUB
        If IsNumeric(tRecord(DeltaIdx + FieldsToJoin(idxIter) - 1)) Then
            tmpData(idxIter) = "f" & FieldsToJoin(idxIter) & "=" _
                                & tRecord(DeltaIdx + FieldsToJoin(idxIter) - 1)
        Else    'Literal strings
            If IsBoolean(CStr(tRecord(DeltaIdx + FieldsToJoin(idxIter) - 1))) Then
                tmpData(idxIter) = "f" & FieldsToJoin(idxIter) & "=" _
                                    & tRecord(DeltaIdx + FieldsToJoin(idxIter) - 1)
            Else
                tmpData(idxIter) = "f" & FieldsToJoin(idxIter) & "=" _
                                    & d_Apostrophe & tRecord(DeltaIdx + FieldsToJoin(idxIter) - 1) & d_Apostrophe
            End If
        End If
    Next idxIter
    JoinVariablesAndValues = Join$(tmpData, d_semicolon)
End Function

''' <summary>
''' Checks if a given field, of the current instance, exists on a record.
''' Returns False when the key can not be found.
''' </summary>
''' <param name="key">The value to be searched.</param>
''' <param name="tIndex">The field to search in (base 1).</param>
Public Function ItemExist(ByRef Key As Variant, tIndex As Long) As Boolean
    ItemExist = (ItemIndex(Key, tIndex) > -1)
End Function

''' <summary>
''' Use BinarySearch, with a given field, to get the target record index.
''' USE ONLY WITH SORTED DATA
''' </summary>
''' <param name="key">The value to be searched.</param>
''' <param name="tIndex">The field to search in (base 1).</param>
''' <param name="StartRecord">The record to start the search in (base 1).</param>
''' <param name="ReverseSearch">Backward matching. The search will end on the starting index</param>
Public Function ItemIndex(ByRef Key As Variant, tIndex As Long, _
                        Optional StartRecord As Long = 1, _
                        Optional ReverseSearch As Boolean = False) As Long
    If CurrentIndex > -1 Then
        ItemIndex = BinarySearch(Key, tIndex, StartRecord, ReverseSearch)
    Else
        ItemIndex = -1
    End If
End Function

Private Function FormatEvalOutput(ByRef EvalOutput As String) As String
    If InStrB(1, EvalOutput, d_Apostrophe) = 1 Then
        FormatEvalOutput = MidB$(EvalOutput, 3, LenB(EvalOutput) - 4)
    Else
        FormatEvalOutput = EvalOutput
    End If
End Function

Public Function SplitField(aIndex As Long, CharToSplitWith As String) As CSVArrayList
    Dim ColUB As Long
    
    On Error GoTo ErrHandler_SplitField
    ColUB = UBound(Buffer(0))
    Select Case aIndex
        Case 0 To ColUB
            Dim curRecord() As Variant
            Dim cpRecord() As String
            Dim tmpRecord() As Variant
            Dim rCounter As Long
            Dim sfldIndex As Long
            Dim sfldCPindex As Long
            Dim FldDiff As Long
                
            For rCounter = 0 To CurrentIndex
                curRecord() = Buffer(rCounter)
                cpRecord() = Split(curRecord(aIndex), CharToSplitWith)
                FldDiff = UBound(cpRecord) - LBound(cpRecord)
                ReDim tmpRecord(0 To ColUB + FldDiff)
                For sfldIndex = 0 To aIndex - 1
                    tmpRecord(sfldIndex) = curRecord(sfldIndex)
                Next sfldIndex
                For sfldCPindex = LBound(cpRecord) To UBound(cpRecord)
                    tmpRecord(sfldIndex) = cpRecord(sfldCPindex)
                    sfldIndex = sfldIndex + 1
                Next sfldCPindex
                For sfldIndex = aIndex + FldDiff + 1 To UBound(curRecord) + FldDiff
                    tmpRecord(sfldIndex) = curRecord(sfldIndex - FldDiff)
                Next sfldIndex
                Buffer(rCounter) = tmpRecord
            Next rCounter
        Case Else
            GoTo OutOfBounds_SplitField
    End Select
    Set SplitField = Me
ErrHandler_SplitField:
    Exit Function
OutOfBounds_SplitField:
    Set SplitField = Nothing
    err.Raise 9
    Resume ErrHandler_SplitField
End Function

Private Function SplitFieldsOrderStr(fieldsString As String) As String()
    Dim KeyIndex As Long
    Dim rfICounter As Long
    Dim tmpPreResult() As String
    Dim tmpResult As CSVArrayList
    
    Set tmpResult = New CSVArrayList
    tmpResult.indexing = True
    tmpPreResult() = Split(RemoveEndSpaces(fieldsString), d_comma)
    If fieldsString Like "*-*" Then
        Dim rfJCounter As Long
        Dim ROFields() As String
        Dim ROFlb As Long
        Dim ROFub As Long
        Dim stepIdx As Long
        
        For rfICounter = LBound(tmpPreResult) To UBound(tmpPreResult)
            If InStrB(1, tmpPreResult(rfICounter), "-") Then
                ROFields() = Split(tmpPreResult(rfICounter), "-")
                ROFlb = CDbl(ROFields(LBound(ROFields)))
                ROFub = CDbl(ROFields(LBound(ROFields) + 1))
                If ROFlb < ROFub Then
                    stepIdx = 1
                Else
                    stepIdx = -1
                End If
                With tmpResult
                    For rfJCounter = ROFlb To ROFub Step stepIdx
                        KeyIndex = KeyIndex + 1
                        .AddIndexedItem CStr(rfJCounter), Array(KeyIndex)
                    Next rfJCounter
                End With
            Else
                With tmpResult
                    KeyIndex = KeyIndex + 1
                    .AddIndexedItem tmpPreResult(rfICounter), Array(KeyIndex)
                End With
            End If
        Next rfICounter
    Else
        With tmpResult
            For rfICounter = LBound(tmpPreResult) To UBound(tmpPreResult)
                KeyIndex = KeyIndex + 1
                .AddIndexedItem tmpPreResult(rfICounter), Array(KeyIndex)
            Next rfICounter
        End With
    End If
    tmpResult.Sort
    SplitFieldsOrderStr = tmpResult.keys
    Set tmpResult = Nothing
End Function
Private Function formatColumnPredicate(predicate As String, _
                                    ByRef headerRecord As Variant) As Long()
    If IsArray(headerRecord) Then  'Only works on tables.
        Dim tmpResult() As Long
        Dim tmpArr() As String
        Dim fcpCounter As Long
        Dim DeltaIdx As Long
        Dim tmpPredicate As String
        
        tmpPredicate = RemoveEndSpaces(predicate)
        DeltaIdx = LBound(headerRecord)
        If tmpPredicate = "*" Then
            tmpArr = SplitFieldsOrderStr("1" _
                        & "-" & CStr(UBound(headerRecord) - LBound(headerRecord) + 1))
        Else
            tmpArr = SplitFieldsOrderStr(tmpPredicate)
        End If
        ReDim tmpResult(LBound(tmpArr) To UBound(tmpArr))
        For fcpCounter = LBound(tmpArr) To UBound(tmpArr)
            If IsNumeric(tmpArr(fcpCounter)) Then
                tmpResult(fcpCounter) = CLng(tmpArr(fcpCounter))
            Else
                tmpResult(fcpCounter) = GetFieldIndex(tmpArr(fcpCounter), headerRecord) - DeltaIdx + 1
            End If
        Next
        formatColumnPredicate = tmpResult
    End If
End Function
Private Function formatJoinPredicate(predicate As String, _
                                    ByRef headerRecordT1 As Variant, _
                                    ByRef headerRecordT2 As Variant) As String
    If IsArray(headerRecordT1) And IsArray(headerRecordT2) Then 'Only works on tables. Relative indexes.
        Dim cLB As Long
        Dim cUB As Long
        Dim DeltaIdx As Long
        Dim ElmIdx As Long
        Dim exprHelper As CSVexpressions
        Dim headerRecord As Variant
        Dim iterator As Long
        Dim predicateCP As String
        Dim tmpElm As String
        Dim tmpStr As String
        Dim tmpVariables() As String
        Dim varComposition() As String
        Dim SwitchedExpr As Boolean
        
        predicateCP = SwitchUnderscoresAndSpaces(predicate, headerRecordT1, headerRecordT2)
        SwitchedExpr = predicateCP <> predicate
        Set exprHelper = New CSVexpressions
        exprHelper.Create predicateCP
        tmpVariables() = Split(exprHelper.currentVariables, "; ")
        For iterator = LBound(tmpVariables) To UBound(tmpVariables)
            varComposition() = Split(tmpVariables(iterator), d_dot)
            cLB = LBound(varComposition)
            cUB = UBound(varComposition)
            If Not IsNumeric(varComposition(cUB)) Then 'Check after dot symbol
                tmpElm = varComposition(cUB)
                If LCase(varComposition(cLB)) = "t1" Then
                    headerRecord = headerRecordT1
                Else
                    If LCase(varComposition(cLB)) = "t2" Then
                        headerRecord = headerRecordT2
                    End If
                End If
                If Not SwitchedExpr Then
                    ElmIdx = GetFieldIndex(tmpElm, headerRecord)
                Else
                    ElmIdx = GetFieldIndex(SwitchUnderscoresAndSpaces(tmpElm, headerRecordT1, headerRecordT2, False), headerRecord)
                End If
                If ElmIdx > -1 Then
                    DeltaIdx = LBound(headerRecord)
                    varComposition(cUB) = ElmIdx - DeltaIdx + 1
                    tmpStr = Join(varComposition, d_dot)
                    predicateCP = Replace(predicateCP, tmpVariables(iterator), tmpStr)
                End If
            End If
        Next iterator
        Set exprHelper = Nothing
        formatJoinPredicate = predicateCP
    End If
End Function
Private Function RemoveEndSpaces(strValue As String) As String
        Dim i As Long
        Dim tmpStr As String
        Dim sFlag As Boolean
        Dim rCounter As Long
        
        tmpStr = strValue
        i = 1
        Do
            sFlag = (MidB$(tmpStr, i, 2) = SPACE_CHR)
            If Not sFlag Then
                If rCounter < 2 Then
                    tmpStr = StrReverse(MidB$(tmpStr, i))
                    i = 1
                    rCounter = rCounter + 1
                End If
            Else
                i = i + 2
            End If
        Loop While sFlag Or rCounter < 2
        RemoveEndSpaces = tmpStr
End Function
Private Function formatKeyStr(Key As String, ByRef headerRecord As Variant) As String
    If IsNumeric(Key) Then
        formatKeyStr = Key
    Else
        Dim DeltaIdx As Long
        DeltaIdx = LBound(headerRecord)
        formatKeyStr = CStr(GetFieldIndex(RemoveEndSpaces(Key), headerRecord) - DeltaIdx + 1) 'Positional index
    End If
End Function
''' <summary>
''' Serializes the buffer contents to a common string representation.
''' Only one-dimensional arrays and jagged arrays populated with
''' one-dimensional arrays are supported.
''' </summary>
Public Function ToString() As String
    Dim tmpData As String
    Dim j As Long
    
    For j = 0 To CurrentIndex
        If j = 0 Then
            tmpData = SerializeRow(Buffer(j))
        Else
            tmpData = tmpData & d_comma & SerializeRow(Buffer(j))
        End If
    Next j
    ToString = "{" & tmpData & "}"
End Function
Private Function SerializeRow(ByRef rArray As Variant) As String
    Dim i As Long, LB As Long, UB As Long
    Dim tmpData As String, Jagged As Boolean
    
    LB = LBound(rArray)
    UB = UBound(rArray)
    Jagged = IsJaggedArray(rArray)
    If Jagged Then 'Recurse
        For i = LB To UB
            If i = LB Then
                tmpData = SerializeRow(rArray(i))
            Else
                tmpData = tmpData & d_comma & SerializeRow(rArray(i))
            End If
        Next i
    Else 'One dimentional array
        For i = LB To UB
            If i = LB Then
                tmpData = tmpData & "{" & rArray(i)
            Else
                tmpData = tmpData & d_comma & rArray(i)
            End If
            If i = UB Then
                tmpData = tmpData & "}"
            End If
        Next i
    End If
    SerializeRow = tmpData
End Function

''' <summary>
''' Main constructor method. Populates the current instance using values passed
''' as a Java array string ({{*};{*}}).
''' </summary>
''' <param name="StrArray">The string definition used to create the array.</param>
''' <param name="Delimiter">The string that delimits the columns and row elements of the matrix.</param>
Public Sub FromString(ByRef strArray As String, Optional Delimiter As String = d_comma)
    Dim i As Long, j As Long
    Dim StrCopy As String
    Dim StrLen As Long
    Dim VectorEndPos As Long
    Dim VectorStartPos As Long
    Dim prevI As Long, prevJ As Long
    
    StrCopy = strArray
    StrLen = LenB(StrCopy) - 4
    If MidB$(StrCopy, StrLen + 1, 4) = "}}" Then
        StrCopy = MidB$(StrCopy, 3, StrLen + 2)
    Else
        If InStrB(1, StrCopy, "{{") Then 'Missed "}" from input
            Exit Sub
        Else
            StrLen = StrLen + 4
        End If
    End If
    prevJ = -1
    prevI = -1
    Do
        j = InStrB(prevJ + 2, StrCopy, "}")
        If j Then
            prevJ = j
        End If
        i = InStrB(prevI + 2, StrCopy, "{")
        Do While i > 0 And i < prevJ
            If i Then
                prevI = i
                i = InStrB(prevI + 2, StrCopy, "{")
            End If
        Loop
        VectorStartPos = prevI + 2
        VectorEndPos = prevJ - 1
        If Not P_INDEXING Then
            Add Split(MidB$(StrCopy, VectorStartPos, VectorEndPos - VectorStartPos + 1), Delimiter)
        Else
            AddIndexedItem CStr(count + 1), Split(MidB$(StrCopy, VectorStartPos, VectorEndPos - VectorStartPos + 1), Delimiter)
        End If
    Loop While prevJ + 1 < StrLen And i
End Sub
Private Function GetFieldIndex(ByRef FieldName As String, ByRef headerRecord As Variant) As Long

    Dim hIterator As Long
    Dim tmpBool As Boolean
    Dim hUB As Long
    Dim tmpResult As Long
    
    hIterator = LBound(headerRecord)
    tmpResult = -1
    hUB = UBound(headerRecord)
    Do
        tmpBool = (FieldName = CStr(headerRecord(hIterator)))
        If tmpBool Then
            tmpResult = hIterator
        End If
        hIterator = hIterator + 1
    Loop While Not tmpBool And hIterator <= hUB
    GetFieldIndex = tmpResult
End Function
Private Sub FillReqFlds(ByRef fullReqFlds As Variant, _
                        ByRef singleReqFlds As Variant, _
                        startIdx As Long, _
                        IndexDelta As Long) 'Expected arrays
Dim idxLoop1 As Long
Dim idxLoop2 As Long

idxLoop2 = startIdx
For idxLoop1 = LBound(singleReqFlds) To UBound(singleReqFlds)
    fullReqFlds(idxLoop2 - IndexDelta) = singleReqFlds(idxLoop1)
    idxLoop2 = idxLoop2 + 1
Next
End Sub
Private Function BindToEvalTable(ByRef expression As String, _
                                        ByVal expressionVariables As String, _
                                        ByRef resultRecord() As Long, _
                                        boundaryIndex As Long) As String
    Dim tmpData() As String
    Dim tmpDLoop As Long
    Dim tmpVariable() As String
    Dim startIdx As Long
    Dim tmpResult As String
    
    tmpResult = expression
    tmpData() = Split(expressionVariables, "; ")
    For tmpDLoop = LBound(tmpData) To UBound(tmpData)
        tmpVariable() = Split(tmpData(tmpDLoop), d_dot)
        Select Case LCase(tmpVariable(LBound(tmpVariable)))
            Case "t1"
                startIdx = LBound(resultRecord)
            Case "t2"
                startIdx = boundaryIndex + 1
        End Select
        tmpResult = Replace(tmpResult, Join(tmpVariable, d_dot), _
                            "f" & CStr(CLng(tmpVariable(UBound(tmpVariable))) + startIdx))

    Next
    BindToEvalTable = tmpResult
End Function

Private Function SwitchUnderscoresAndSpaces(ByRef cExpression As String, _
                                            ByRef lHeader As Variant, _
                                            Optional rHeader As Variant = 1, _
                                            Optional ToUnderscores As Boolean = True) As String
    Dim tmpResult As String
    Dim tmpArr As Variant
    Dim iterator As Long
    Dim counterH As Long
    
    tmpResult = cExpression
    tmpArr = lHeader
    Do
        For iterator = LBound(tmpArr) To UBound(tmpArr)
            If ToUnderscores Then
                If InStrB(1, tmpArr(iterator), SPACE_CHR) Then
                    tmpResult = Replace(tmpResult, tmpArr(iterator), Join$(Split(tmpArr(iterator), SPACE_CHR), UNDERSCORE_CHR))
                End If
            Else
                If InStrB(1, tmpArr(iterator), SPACE_CHR) Then
                    tmpResult = Replace(tmpResult, Join$(Split(tmpArr(iterator), SPACE_CHR), UNDERSCORE_CHR), tmpArr(iterator))
                End If
            End If
        Next iterator
        counterH = counterH + 1
        If TypeName(rHeader) = "Variant()" Then
            tmpArr = rHeader
        Else
            Exit Do
        End If
    Loop While counterH = 1
    SwitchUnderscoresAndSpaces = tmpResult
End Function

''' <summary>
''' Run a left outer join on the provided data tables.
'''' 1) Use a string such as "{1-2,5,ID};{1-6}" as a predicate of the columns to
'''' indicate the join of columns 1 to 2, 5 and ID of leftTable with
'''' the columns 1 to 6 of rightTable.
'''' 2) Use a string such as "{*};{1-3}" to indicate the union of ALL
'''' columns of leftTable with columns 1 to 3 of rightTable.
''' 3) The predicate must use the dot syntax [t1.#][t1.fieldName] to.
'''' indicate the fields of the table, where t1 refers to the leftTable.
''' 4) The matchKeys predicate must be given as ''''#/$;#/$'''.
''' </summary>
''' <param name="leftTable">First table.</param>.
''' <param name="rightTable">Second table.</param> ''' <param name="rightTable">Second table.
''' <param name="columns">String that specifies the structure of the rows returned.</param>
''' <param name="matchKeys">Priority and preference keys to be matched.</param>
''' <param name="predicate">Condition that must be met when selecting rows.</param>
''' <param name="headers">Must be True if both tables have a header row.</param>
''' All indexes must be given IN BASE 1.
Public Function LeftJoin(ByRef leftTable As CSVArrayList, _
                            ByRef rightTable As CSVArrayList, _
                            Columns As String, _
                            matchKeys As String, _
                            Optional predicate As String = vbNullString, _
                            Optional Headers As Boolean = True) As CSVArrayList
                             
    Dim cColumns() As String
    Dim cKeys() As String
    Dim cKLBidx As Long
    Dim cPredicate As String
    Dim evalRecord() As Variant
    Dim evalRecord_BK() As Variant
    Dim exprHelper As CSVexpressions
    Dim FilterFields() As Long
    Dim fullReqFields() As Long
    Dim leftIdxDelta As Long
    Dim lftTableReqFlds() As Long
    Dim LftTblENDidx As Long
    Dim lJoinIndex As Long
    Dim MatchesForCurrentKey As CSVArrayList
    Dim MatchesCount As Long
    Dim MatchCounter As Long
    Dim rCounter As Long
    Dim refTable As CSVArrayList
    Dim resultRecord() As Variant
    Dim resultRecord_BK() As Variant
    Dim rgtTableReqFlds() As Long
    Dim rightIdxDelta As Long
    Dim sCounter As Long
    Dim tmpIndex As Long
    Dim WrittenHeaders As Boolean
    
    On Error GoTo join_errHandler
    If leftTable.indexing Or rightTable.indexing Then Exit Function 'Not available for indexed lists
    Set LeftJoin = New CSVArrayList
    Set MatchesForCurrentKey = New CSVArrayList
    cColumns() = Split(Columns, d_semicolon)
    lftTableReqFlds() = formatColumnPredicate( _
                                                MidB$(cColumns(LBound(cColumns)), _
                                            3, LenB(cColumns(LBound(cColumns))) - 4), _
                                            leftTable(0))
    rgtTableReqFlds() = formatColumnPredicate( _
                                                MidB$(cColumns(UBound(cColumns)), _
                                            3, LenB(cColumns(UBound(cColumns))) - 4), _
                                            rightTable(0))
    cKeys() = Split(matchKeys, d_semicolon)
    cKLBidx = LBound(cKeys)
    cKeys(cKLBidx) = formatKeyStr(cKeys(cKLBidx), leftTable(0))
    cKeys(cKLBidx + 1) = formatKeyStr(cKeys(cKLBidx + 1), rightTable(0))
    If predicate <> vbNullString Then
        cPredicate = formatJoinPredicate(predicate, leftTable(0), rightTable(0))
    End If
    LftTblENDidx = UBound(lftTableReqFlds)
    ReDim fullReqFields(0 To (LftTblENDidx + UBound(rgtTableReqFlds) + 1))
    ReDim resultRecord(0 To UBound(fullReqFields))
    resultRecord_BK() = resultRecord
    leftIdxDelta = LBound(leftTable(0))
    rightIdxDelta = LBound(rightTable(0))
    FillReqFlds fullReqFields, lftTableReqFlds, leftIdxDelta, leftIdxDelta
    FillReqFlds fullReqFields, rgtTableReqFlds, LftTblENDidx + 1, leftIdxDelta
    Set refTable = rightTable.Copy
    refTable.Sort 1 + Abs(Headers), sortingKeys:=CLng(cKeys(UBound(cKeys))) + rightIdxDelta  'Sort the right table
    Set exprHelper = New CSVexpressions
    With exprHelper
        If predicate <> vbNullString Then
            ReDim evalRecord(0 To (UBound(leftTable(0)) - LBound(leftTable(0)) + UBound(rightTable(0)) - LBound(rightTable(0))) + 1)
            evalRecord_BK() = evalRecord
            .Create cPredicate
            cPredicate = BindToEvalTable(.expression, _
                                    .currentVariables, _
                                    fullReqFields, UBound(leftTable(0)) - leftIdxDelta)
        
            .Create cPredicate
            FilterFields() = GetIndexesFromVarList(.currentVariables)
        End If
        For rCounter = Abs(Headers) To leftTable.count - 1
            'Initial configurations
            MatchesForCurrentKey.Clear
            MatchCounter = 0
            lJoinIndex = -1
            'Find a key occurrence
            tmpIndex = refTable.ItemIndex(leftTable(rCounter)(CLng(cKeys(cKLBidx)) + leftIdxDelta - 1), _
                                            CLng(cKeys(cKLBidx + 1)) + rightIdxDelta, lJoinIndex + 2)
            'Find the first occurrence of the key
            Do While tmpIndex > -1
                lJoinIndex = tmpIndex
                tmpIndex = refTable.ItemIndex(leftTable(rCounter)(CLng(cKeys(cKLBidx)) + leftIdxDelta - 1), _
                                            CLng(cKeys(cKLBidx + 1)) + rightIdxDelta, tmpIndex, True)
            Loop
            If lJoinIndex > -1 Then
                'Save all occurrences
                Do
                    MatchesForCurrentKey.Add lJoinIndex
                    lJoinIndex = lJoinIndex + 1
                Loop While leftTable(rCounter)(CLng(cKeys(cKLBidx)) + leftIdxDelta - 1) _
                            = refTable(lJoinIndex)(CLng(cKeys(cKLBidx + 1)) + rightIdxDelta - 1)
            End If
            MatchesCount = MatchesForCurrentKey.count
            If MatchesCount = 0 Then
                MatchesForCurrentKey.Add -1
            End If
            Do
                If Headers Then
                    If Not WrittenHeaders Then
                        ' Fill headers on left
                        For sCounter = 0 To LftTblENDidx
                            resultRecord(sCounter) = leftTable(0)(lftTableReqFlds(sCounter) _
                                                    + leftIdxDelta - 1)
                        Next sCounter
                        ' Fill headers on right
                        For sCounter = LftTblENDidx + 1 To UBound(resultRecord)
                            resultRecord(sCounter) = refTable(0)(rgtTableReqFlds(sCounter - _
                                                    (LftTblENDidx + 1)) + rightIdxDelta - 1)
                        Next sCounter
                        LeftJoin.Add resultRecord 'Append current record
                        WrittenHeaders = True
                        resultRecord() = resultRecord_BK
                    End If
                End If
                'Fill left table data
                For sCounter = 0 To LftTblENDidx
                    resultRecord(sCounter) = leftTable(rCounter)(lftTableReqFlds(sCounter) _
                                                    + leftIdxDelta - 1)
                Next sCounter
                lJoinIndex = MatchesForCurrentKey(MatchCounter)
                If predicate <> vbNullString Then
                    FillReqFlds evalRecord, leftTable(rCounter), leftIdxDelta, leftIdxDelta
                    If lJoinIndex > -1 Then
                        FillReqFlds evalRecord, refTable(lJoinIndex), UBound(leftTable(0)) + 1, leftIdxDelta
                    End If
                    On Error Resume Next
                    .Eval GetValuesForVariables(rCounter, FilterFields, False, evalRecord)
                    If .errorType = ExpressionErrors.errNone Then
                        If err.Number = 0 Then
                            If CBool(.result) Then
                                'Fill in the data in the table on the right only if the join and predicate are satisfied.
                                If lJoinIndex > -1 Then
                                    For sCounter = LftTblENDidx + 1 To UBound(resultRecord)
                                        resultRecord(sCounter) = refTable(lJoinIndex)(rgtTableReqFlds(sCounter - _
                                                                (LftTblENDidx + 1)) + rightIdxDelta - 1)
                                    Next sCounter
                                End If
                            End If
                        Else
                        err.Clear
                        End If
                    End If
                Else
                    'Fill in the data in the table on the right only if the join is satisfied.
                    If lJoinIndex > -1 Then
                        For sCounter = LftTblENDidx + 1 To UBound(resultRecord)
                            resultRecord(sCounter) = refTable(lJoinIndex)(rgtTableReqFlds(sCounter - _
                                                            (LftTblENDidx + 1)) + rightIdxDelta - 1)
                        Next sCounter
                    End If
                End If
                LeftJoin.Add resultRecord 'Append current record
                resultRecord() = resultRecord_BK
                evalRecord() = evalRecord_BK
                MatchCounter = MatchCounter + 1
            Loop While MatchCounter <= MatchesCount - 1
        Next rCounter
    End With
    Set exprHelper = Nothing
    Set refTable = Nothing
    Exit Function
join_errHandler:
End Function
''' <summary>
''' Run a right outer join on the provided data tables.
'''' 1) Use a string such as "{1-2,5,ID};{1-6}" as a predicate of the columns to
'''' indicate the join of columns 1 to 2, 5 and ID of leftTable with
'''' the columns 1 to 6 of rightTable.
'''' 2) Use a string such as "{*};{1-3}" to indicate the union of ALL
'''' columns of leftTable with columns 1 to 3 of rightTable.
''' 3) The predicate must use the dot syntax [t1.#][t1.fieldName] to.
'''' indicate the fields of the table, where t1 refers to the leftTable.
''' 4) The matchKeys predicate must be given as ''''#/$;#/$'''.
''' </summary>
''' <param name="leftTable">First table.</param>.
''' <param name="rightTable">Second table.</param> ''' <param name="rightTable">Second table.
''' <param name="columns">String that specifies the structure of the rows returned.</param>
''' <param name="matchKeys">Priority and preference keys to be matched.</param>
''' <param name="predicate">Condition that must be met when selecting rows.</param>
''' <param name="headers">Must be True if both tables have a header row.</param>
''' All indexes must be given IN BASE 1.
Public Function RightJoin(ByRef leftTable As CSVArrayList, _
                            ByRef rightTable As CSVArrayList, _
                            Columns As String, _
                            matchKeys As String, _
                            Optional predicate As String = vbNullString, _
                            Optional Headers As Boolean = True) As CSVArrayList
    On Error GoTo join_errHandler
    Dim cColumns() As String
    Dim cKeys() As String
    Dim cKLBidx As Long
    Dim cPredicate As String
    Dim evalRecord() As Variant
    Dim evalRecord_BK() As Variant
    Dim exprHelper As CSVexpressions
    Dim FilterFields() As Long
    Dim fullReqFields() As Long
    Dim leftIdxDelta As Long
    Dim lftTableReqFlds() As Long
    Dim LftTblENDidx As Long
    Dim lJoinIndex As Long
    Dim MatchesForCurrentKey As CSVArrayList
    Dim MatchCounter As Long
    Dim MatchesCount As Long
    Dim rCounter As Long
    Dim refTable As CSVArrayList
    Dim resultRecord() As Variant
    Dim resultRecord_BK() As Variant
    Dim rgtTableReqFlds() As Long
    Dim rightIdxDelta As Long
    Dim sCounter As Long
    Dim tmpIndex As Long
    Dim WrittenHeaders As Boolean
    
    If leftTable.indexing Or rightTable.indexing Then Exit Function 'Not available for indexed lists
    Set RightJoin = New CSVArrayList
    Set MatchesForCurrentKey = New CSVArrayList
    cColumns() = Split(Columns, d_semicolon)
    lftTableReqFlds() = formatColumnPredicate( _
                                                MidB$(cColumns(LBound(cColumns)), _
                                            3, LenB(cColumns(LBound(cColumns))) - 4), _
                                            leftTable(0))
    rgtTableReqFlds() = formatColumnPredicate( _
                                                MidB$(cColumns(UBound(cColumns)), _
                                            3, LenB(cColumns(UBound(cColumns))) - 4), _
                                            rightTable(0))
    cKeys() = Split(matchKeys, d_semicolon)
    cKLBidx = LBound(cKeys)
    cKeys(cKLBidx) = formatKeyStr(cKeys(cKLBidx), leftTable(0))
    cKeys(cKLBidx + 1) = formatKeyStr(cKeys(cKLBidx + 1), rightTable(0))
    If predicate <> vbNullString Then
        cPredicate = formatJoinPredicate(predicate, leftTable(0), rightTable(0))
    End If
    LftTblENDidx = UBound(lftTableReqFlds)
    ReDim fullReqFields(0 To (LftTblENDidx + UBound(rgtTableReqFlds) + 1))
    ReDim resultRecord(0 To UBound(fullReqFields))
    resultRecord_BK() = resultRecord
    leftIdxDelta = LBound(leftTable(0))
    rightIdxDelta = LBound(rightTable(0))
    FillReqFlds fullReqFields, lftTableReqFlds, leftIdxDelta, leftIdxDelta
    FillReqFlds fullReqFields, rgtTableReqFlds, LftTblENDidx + 1, leftIdxDelta
    Set refTable = leftTable.Copy
    refTable.Sort 1 + Abs(Headers), sortingKeys:=CLng(cKeys(LBound(cKeys))) + leftIdxDelta  'Sort the left table
    Set exprHelper = New CSVexpressions
    With exprHelper
        If predicate <> vbNullString Then
            ReDim evalRecord(0 To (UBound(leftTable(0)) - LBound(leftTable(0)) + UBound(rightTable(0)) - LBound(rightTable(0))) + 1)
            evalRecord_BK() = evalRecord
            .Create cPredicate
            cPredicate = BindToEvalTable(.expression, _
                                    .currentVariables, _
                                    fullReqFields, UBound(leftTable(0)) - leftIdxDelta)
        
            .Create cPredicate
            FilterFields() = GetIndexesFromVarList(.currentVariables)
        End If
        For rCounter = Abs(Headers) To rightTable.count - 1
            'Initial configurations
            MatchesForCurrentKey.Clear
            MatchCounter = 0
            lJoinIndex = -1
            'Find a key occurrence
            tmpIndex = refTable.ItemIndex(rightTable(rCounter)(CLng(cKeys(cKLBidx + 1)) + rightIdxDelta - 1), _
                                            CLng(cKeys(cKLBidx)) + leftIdxDelta, lJoinIndex + 2)
            'Find the first occurrence of the key
            Do While tmpIndex > -1
                lJoinIndex = tmpIndex
                tmpIndex = refTable.ItemIndex(rightTable(rCounter)(CLng(cKeys(cKLBidx + 1)) + rightIdxDelta - 1), _
                                            CLng(cKeys(cKLBidx)) + leftIdxDelta, tmpIndex, True)
            Loop
            If lJoinIndex > -1 Then
                'Save all occurrences
                Do
                    MatchesForCurrentKey.Add lJoinIndex
                    lJoinIndex = lJoinIndex + 1
                Loop While rightTable(rCounter)(CLng(cKeys(cKLBidx)) + rightIdxDelta - 1) _
                            = refTable(lJoinIndex)(CLng(cKeys(cKLBidx + 1)) + leftIdxDelta - 1)
            End If
            MatchesCount = MatchesForCurrentKey.count
            If MatchesCount = 0 Then
                MatchesForCurrentKey.Add -1
            End If
            Do
                If Headers Then
                    If Not WrittenHeaders Then
                        ' Fill headers on right
                        For sCounter = LftTblENDidx + 1 To UBound(resultRecord)
                            resultRecord(sCounter) = rightTable(0)(rgtTableReqFlds(sCounter - _
                                                        (LftTblENDidx + 1)) + rightIdxDelta - 1)
                        Next sCounter
                        ' Fill headers on left
                        For sCounter = 0 To LftTblENDidx
                            resultRecord(sCounter) = refTable(0)(lftTableReqFlds(sCounter) _
                                                    + leftIdxDelta - 1)
                        Next sCounter
                        RightJoin.Add resultRecord 'Append current record
                        WrittenHeaders = True
                        resultRecord() = resultRecord_BK
                    End If
                End If
                'Fill right table data
                For sCounter = LftTblENDidx + 1 To UBound(resultRecord)
                    resultRecord(sCounter) = rightTable(rCounter)(rgtTableReqFlds(sCounter - _
                                                        (LftTblENDidx + 1)) + rightIdxDelta - 1)
                Next sCounter
                lJoinIndex = MatchesForCurrentKey(MatchCounter)
                If predicate <> vbNullString Then
                    FillReqFlds evalRecord, rightTable(rCounter), UBound(leftTable(0)) + 1, leftIdxDelta
                    If lJoinIndex > -1 Then
                        FillReqFlds evalRecord, refTable(lJoinIndex), leftIdxDelta, leftIdxDelta
                    End If
                    On Error Resume Next
                    .Eval GetValuesForVariables(rCounter, FilterFields, False, evalRecord)
                    If .errorType = ExpressionErrors.errNone Then
                        If err.Number = 0 Then
                            If CBool(.result) Then
                                'Fill in the data in the table on the left only if the join and predicate are satisfied.
                                If lJoinIndex > -1 Then
                                    For sCounter = 0 To LftTblENDidx
                                        resultRecord(sCounter) = refTable(lJoinIndex)(lftTableReqFlds(sCounter) _
                                                                        + leftIdxDelta - 1)
                                    Next sCounter
                                End If
                            End If
                        Else
                        err.Clear
                        End If
                    End If
                Else
                    'Fill in the data in the table on the left only if the join is satisfied.
                    If lJoinIndex > -1 Then
                        For sCounter = 0 To LftTblENDidx
                            resultRecord(sCounter) = refTable(lJoinIndex)(lftTableReqFlds(sCounter) _
                                                            + leftIdxDelta - 1)
                        Next sCounter
                    End If
                End If
                RightJoin.Add resultRecord 'Append current record
                resultRecord() = resultRecord_BK
                evalRecord() = evalRecord_BK
                MatchCounter = MatchCounter + 1
            Loop While MatchCounter <= MatchesCount - 1
        Next rCounter
    End With
    Set exprHelper = Nothing
    Set refTable = Nothing
    Exit Function
join_errHandler:
End Function
''' <summary>
''' Run an inner join on the provided data tables.
'''' 1) Use a string such as "{1-2,5,ID};{1-6}" as a predicate of the columns to
'''' indicate the join of columns 1 to 2, 5 and ID of leftTable with
'''' the columns 1 to 6 of rightTable.
'''' 2) Use a string such as "{*};{1-3}" to indicate the union of ALL '''' columns of leftTable.
'''' columns of leftTable with columns 1 to 3 of rightTable.
''' 3) The predicate must use the dot syntax [t1.#][t1.fieldName] to.
'''' indicate the fields of the table, where t1 refers to the leftTable.
''' 4) The matchKeys predicate must be given as ''''#/$;#/$'''.
''' </summary>
''' <param name="leftTable">First table.</param>.
''' <param name="rightTable">Second table.</param> ''' <param name="rightTable">Second table.
''' <param name="columns">String that specifies the structure of the rows returned.</param>
''' <param name="matchKeys">Priority and preference keys to be matched.</param>
''' <param name="predicate">Condition that must be met when selecting rows.</param>
''' <param name="headers">Must be True if both tables have a header row.</param>
''' All indexes must be given IN BASE 1.
Public Function InnerJoin(ByRef leftTable As CSVArrayList, _
                            ByRef rightTable As CSVArrayList, _
                            Columns As String, _
                            matchKeys As String, _
                            Optional predicate As String = vbNullString, _
                            Optional Headers As Boolean = True) As CSVArrayList
                             
    Dim cColumns() As String
    Dim cKeys() As String
    Dim cKLBidx As Long
    Dim cPredicate As String
    Dim evalRecord() As Variant
    Dim evalRecord_BK() As Variant
    Dim exprHelper As CSVexpressions
    Dim FilterFields() As Long
    Dim fullReqFields() As Long
    Dim leftIdxDelta As Long
    Dim lftTableReqFlds() As Long
    Dim LftTblENDidx As Long
    Dim lJoinIndex As Long
    Dim MatchesForCurrentKey As CSVArrayList
    Dim MatchCounter As Long
    Dim MatchesCount As Long
    Dim rCounter As Long
    Dim refTable As CSVArrayList
    Dim resultRecord() As Variant
    Dim resultRecord_BK() As Variant
    Dim rgtTableReqFlds() As Long
    Dim rightIdxDelta As Long
    Dim sCounter As Long
    Dim tmpIndex As Long
    Dim WrittenHeaders As Boolean
    
    If leftTable.indexing Or rightTable.indexing Then Exit Function 'Not available for indexed lists
    Set InnerJoin = New CSVArrayList
    Set MatchesForCurrentKey = New CSVArrayList
    cColumns() = Split(Columns, d_semicolon)
    lftTableReqFlds() = formatColumnPredicate( _
                                                MidB$(cColumns(LBound(cColumns)), _
                                            3, LenB(cColumns(LBound(cColumns))) - 4), _
                                            leftTable(0))
    rgtTableReqFlds() = formatColumnPredicate( _
                                                MidB$(cColumns(UBound(cColumns)), _
                                            3, LenB(cColumns(UBound(cColumns))) - 4), _
                                            rightTable(0))
    cKeys() = Split(matchKeys, d_semicolon)
    cKLBidx = LBound(cKeys)
    cKeys(cKLBidx) = formatKeyStr(cKeys(cKLBidx), leftTable(0))
    cKeys(cKLBidx + 1) = formatKeyStr(cKeys(cKLBidx + 1), rightTable(0))
    If predicate <> vbNullString Then
        cPredicate = formatJoinPredicate(predicate, leftTable(0), rightTable(0))
    End If
    LftTblENDidx = UBound(lftTableReqFlds)
    ReDim fullReqFields(0 To (LftTblENDidx + UBound(rgtTableReqFlds) + 1))
    ReDim resultRecord(0 To UBound(fullReqFields))
    resultRecord_BK() = resultRecord
    leftIdxDelta = LBound(leftTable(0))
    rightIdxDelta = LBound(rightTable(0))
    FillReqFlds fullReqFields, lftTableReqFlds, leftIdxDelta, leftIdxDelta
    FillReqFlds fullReqFields, rgtTableReqFlds, LftTblENDidx + 1, leftIdxDelta
    Set refTable = rightTable.Copy
    refTable.Sort 1 + Abs(Headers), sortingKeys:=CLng(cKeys(UBound(cKeys))) + rightIdxDelta  'Sort the right table
    Set exprHelper = New CSVexpressions
    With exprHelper
        If predicate <> vbNullString Then
            ReDim evalRecord(0 To (UBound(leftTable(0)) - LBound(leftTable(0)) + UBound(rightTable(0)) - LBound(rightTable(0))) + 1)
            evalRecord_BK() = evalRecord
            .Create cPredicate
            cPredicate = BindToEvalTable(.expression, _
                                    .currentVariables, _
                                    fullReqFields, UBound(leftTable(0)) - leftIdxDelta)
        
            .Create cPredicate
            FilterFields() = GetIndexesFromVarList(.currentVariables)
        End If
        For rCounter = Abs(Headers) To leftTable.count - 1
                        'Initial configurations
            MatchesForCurrentKey.Clear
            MatchCounter = 0
            lJoinIndex = -1
            'Find a key occurrence
            tmpIndex = refTable.ItemIndex(leftTable(rCounter)(CLng(cKeys(cKLBidx)) + leftIdxDelta - 1), _
                                            CLng(cKeys(cKLBidx + 1)) + rightIdxDelta, lJoinIndex + 2)
            'Find the first occurrence of the key
            Do While tmpIndex > -1
                lJoinIndex = tmpIndex
                tmpIndex = refTable.ItemIndex(leftTable(rCounter)(CLng(cKeys(cKLBidx)) + leftIdxDelta - 1), _
                                            CLng(cKeys(cKLBidx + 1)) + rightIdxDelta, tmpIndex, True)
            Loop
            If lJoinIndex > -1 Then
                'Save all occurrences
                Do
                    MatchesForCurrentKey.Add lJoinIndex
                    lJoinIndex = lJoinIndex + 1
                Loop While leftTable(rCounter)(CLng(cKeys(cKLBidx)) + leftIdxDelta - 1) _
                            = refTable(lJoinIndex)(CLng(cKeys(cKLBidx + 1)) + rightIdxDelta - 1)
            End If
            MatchesCount = MatchesForCurrentKey.count
            If MatchesCount = 0 Then
                MatchesForCurrentKey.Add -1
            End If
            Do
                If Headers Then
                    If Not WrittenHeaders Then
                        ' Fill headers on left
                        For sCounter = 0 To LftTblENDidx
                            resultRecord(sCounter) = leftTable(0)(lftTableReqFlds(sCounter) _
                                                    + leftIdxDelta - 1)
                        Next sCounter
                        ' Fill headers on right
                        For sCounter = LftTblENDidx + 1 To UBound(resultRecord)
                            resultRecord(sCounter) = refTable(0)(rgtTableReqFlds(sCounter - _
                                                    (LftTblENDidx + 1)) + rightIdxDelta - 1)
                        Next sCounter
                        InnerJoin.Add resultRecord 'Append current record
                        WrittenHeaders = True
                        resultRecord() = resultRecord_BK
                    End If
                End If
                lJoinIndex = MatchesForCurrentKey(MatchCounter)
                If predicate <> vbNullString Then
                    FillReqFlds evalRecord, leftTable(rCounter), leftIdxDelta, leftIdxDelta
                    If lJoinIndex > -1 Then
                        FillReqFlds evalRecord, refTable(lJoinIndex), UBound(leftTable(0)) + 1, leftIdxDelta
                    End If
                    On Error Resume Next
                    .Eval GetValuesForVariables(rCounter, FilterFields, False, evalRecord)
                    If .errorType = ExpressionErrors.errNone Then
                        If err.Number = 0 Then
                            If CBool(.result) Then
                                'Fill in the data in the table.
                                If lJoinIndex > -1 Then
                                    For sCounter = 0 To LftTblENDidx
                                        resultRecord(sCounter) = leftTable(rCounter)(lftTableReqFlds(sCounter) _
                                                                    + leftIdxDelta - 1)
                                    Next sCounter
                                    For sCounter = LftTblENDidx + 1 To UBound(resultRecord)
                                        resultRecord(sCounter) = refTable(lJoinIndex)(rgtTableReqFlds(sCounter - _
                                                                (LftTblENDidx + 1)) + rightIdxDelta - 1)
                                    Next sCounter
                                    InnerJoin.Add resultRecord 'Append current record
                                End If
                            End If
                        Else
                        err.Clear
                        End If
                    End If
                Else
                    'Fill in the data in the table.
                    If lJoinIndex > -1 Then
                        For sCounter = 0 To LftTblENDidx
                            resultRecord(sCounter) = leftTable(rCounter)(lftTableReqFlds(sCounter) _
                                                        + leftIdxDelta - 1)
                        Next sCounter
                        For sCounter = LftTblENDidx + 1 To UBound(resultRecord)
                            resultRecord(sCounter) = refTable(lJoinIndex)(rgtTableReqFlds(sCounter - _
                                                            (LftTblENDidx + 1)) + rightIdxDelta - 1)
                        Next sCounter
                        InnerJoin.Add resultRecord 'Append current record
                    End If
                End If
                resultRecord() = resultRecord_BK
                evalRecord() = evalRecord_BK
                MatchCounter = MatchCounter + 1
            Loop While MatchCounter <= MatchesCount - 1
        Next rCounter
    End With
    Set exprHelper = Nothing
    Set refTable = Nothing
End Function

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
    Dim r As Long
    
    l = 2 * i                                   'LEFT
    r = 2 * i + 1                               'RIGHT
    largest = i
    If l <= HeapSize Then
        If Buffer(leftt + l - 1)(SortingKey) > Buffer(leftt + i - 1)(SortingKey) Then
            largest = l
        End If
    End If
    If r <= HeapSize Then
        If Buffer(leftt + r - 1)(SortingKey) > Buffer(leftt + largest - 1)(SortingKey) Then
            largest = r
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
Private Sub merge(l As Long, m As Long, r As Long, SortingKey As Long, Descending As Boolean)
    
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
    len2 = r - m
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

Public Function MergeFields(indexes As String, CharToMergeWith As String) As CSVArrayList
    Dim cpIndexes() As String
    
    cpIndexes() = FieldsToOperate(indexes)
    If CheckForDupIndexes(cpIndexes) Then
        Set MergeFields = MergeFlds(CharToMergeWith, cpIndexes)
    Else
        Set MergeFields = Me
    End If
End Function

Private Function MergeFlds(CharToMergeWith As String, cpIndexes() As String) As CSVArrayList
    On Error GoTo ErrHandler_MergeFields
    Select Case FielsToMergeAreValid(cpIndexes)
        Case True
            Dim BufferColUB As Long
            Dim curRecord() As Variant
            Dim cpRecord() As Variant
            Dim fCounter As Long
            Dim rCounter As Long
            Dim tmpMergeRec As CSVArrayList
            Dim LB As Long
            Dim UB As Long
            Dim storedIndexes As Long
            Dim IndexToSaveIn As Long
            Dim tmpIndex As Long
            
            LB = LBound(cpIndexes)
            UB = UBound(cpIndexes)
            '@--------------------------------------------------------------------------------
            'Shrink storage
            BufferColUB = UBound(Buffer(0))
            ReDim cpRecord(0 To BufferColUB - UB + LB)
            '@--------------------------------------------------------------------------------
            'Reserve storage
            Set tmpMergeRec = New CSVArrayList
            tmpMergeRec.Reinitialize UB - LB
            With tmpMergeRec
                For rCounter = 0 To CurrentIndex
                    curRecord() = Buffer(rCounter)
                    storedIndexes = 0
                    For fCounter = 0 To BufferColUB
                        If FiedlInMergeList(fCounter, cpIndexes) Then
                            .Add curRecord(cpIndexes(storedIndexes))
                            storedIndexes = storedIndexes + 1
                        Else
                            '@--------------------------------------------------------------------------------
                            'Check field position
                            tmpIndex = fCounter - storedIndexes
                            Select Case tmpIndex
                                Case Is < cpIndexes(LB)
                                    cpRecord(tmpIndex) = curRecord(fCounter)
                                Case Else
                                    cpRecord(tmpIndex + 1) = curRecord(fCounter)
                            End Select
                        End If
                    Next fCounter
                    Select Case cpIndexes(LB)
                        '@--------------------------------------------------------------------------------
                        'Take care of bounds
                        Case Is > UBound(cpRecord)
                            IndexToSaveIn = BufferColUB - UB + LB
                        Case Else
                            IndexToSaveIn = cpIndexes(LB)
                    End Select
                    cpRecord(IndexToSaveIn) = Join$(tmpMergeRec.items, CharToMergeWith)
                    Buffer(rCounter) = cpRecord
                    .Clear
                Next rCounter
            End With
        Case Else
            GoTo OutOfBounds_MergeFields
    End Select
    Set MergeFlds = Me
ErrHandler_MergeFields:
    Exit Function
OutOfBounds_MergeFields:
    Set MergeFlds = Nothing
    err.Raise Number:=9
    Resume ErrHandler_MergeFields
End Function

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
    Dim r As Long
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
            r = Ulen + 1 ' start of second (right) half
            O = leftt ' start of output (mirror array)
            Do
                BoolSwitch = Buffer(r)(SortingKey) < Buffer(l)(SortingKey)
                If BoolSwitch Then
                    pvarMirror(O) = Buffer(r)
                    r = r + 1
                    If r > rightt Then
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
                        For r = r To rightt
                            O = O + 1
                            pvarMirror(O) = Buffer(r)
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
    Dim r As Long
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
            r = Ulen + 1 ' start of second (right) half
            O = leftt ' start of output (mirror array)
            Do
                BoolSwitch = Buffer(r)(SortingKey) > Buffer(l)(SortingKey)
                If BoolSwitch Then
                    pvarMirror(O) = Buffer(r)
                    r = r + 1
                    If r > rightt Then
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
                        For r = r To rightt
                            O = O + 1
                            pvarMirror(O) = Buffer(r)
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
Private Function Min(a As Long, b As Long) As Long
    If b < a Then
        Min = b
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
    Dim r As Long
    
    l = 2 * i                                   'LEFT
    r = 2 * i + 1                               'RIGHT
    smallest = i
    If l <= HeapSize Then
        If Buffer(leftt + l - 1)(SortingKey) < Buffer(leftt + i - 1)(SortingKey) Then
            smallest = l
        End If
    End If
    If r <= HeapSize Then
        If Buffer(leftt + r - 1)(SortingKey) < Buffer(leftt + smallest - 1)(SortingKey) Then
            smallest = r
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
Private Sub QuickSort(leftt As Long, _
                        rightt As Long, _
                        div As Long, _
                        SortingKey As Long, _
                        Optional Descending As Boolean = False, _
                        Optional Indexed As Boolean = False)
    If Descending Then
        QuickSort_Desc leftt, rightt, div, SortingKey, Indexed
    Else
        QuickSort_Asc leftt, rightt, div, SortingKey, Indexed
    End If
End Sub

''' <summary>
''' Sorts the data in ascending order. Requires all elements to be one-dimensional arrays.
''' </summary>
Private Sub QuickSort_Asc(leftt As Long, _
                        rightt As Long, _
                        div As Long, _
                        SortingKey As Long, _
                        Optional Indexed As Boolean = False)
    
    Dim Ulen As Long
    
    Ulen = rightt - leftt
    If Ulen < RUN Then 'insertion sort for tiny array
        InsertionSort leftt, rightt, SortingKey, False, Indexed
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
        If Not Indexed Then
            If Buffer(m1)(SortingKey) < Buffer(m2)(SortingKey) Then
                Swap m1, leftt
                Swap m2, rightt
            Else
                Swap m1, rightt
                Swap m2, leftt
            End If
        Else
            If IndexedBuffer(m1).ItemValue(SortingKey) < IndexedBuffer(m2).ItemValue(SortingKey) Then
                Swap m1, leftt, True
                Swap m2, rightt, True
            Else
                Swap m1, rightt, True
                Swap m2, leftt, True
            End If
        End If
        'pivots
        Dim pivot1 As Variant
        Dim pivot2 As Variant
        If Not Indexed Then
            pivot1 = Buffer(leftt)(SortingKey)
            pivot2 = Buffer(rightt)(SortingKey)
        Else
            pivot1 = IndexedBuffer(leftt).ItemValue(SortingKey)
            pivot2 = IndexedBuffer(rightt).ItemValue(SortingKey)
        End If
        'pointers
        Dim less As Long
        Dim great As Long
        less = leftt + 1
        great = rightt - 1
        'sorting
        Dim k As Long
        k = less
        If Not Indexed Then
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
        Else
            Do While k <= great
                If IndexedBuffer(k).ItemValue(SortingKey) < pivot1 Then
                    Swap k, less, True
                    less = less + 1
                ElseIf IndexedBuffer(k).ItemValue(SortingKey) > pivot2 Then
                    Do While k < great And IndexedBuffer(great).ItemValue(SortingKey) > pivot2
                        great = great - 1
                    Loop
                    Swap k, great, True
                    great = great - 1
                    If IndexedBuffer(k).ItemValue(SortingKey) < pivot1 Then
                        Swap k, less, True
                        less = less + 1
                    End If
                End If
                k = k + 1
            Loop
        End If
        'swaps
        Dim dist As Long
        dist = great - less
        If dist < 13 Then div = div + 1
        If Not Indexed Then
            Swap less - 1, leftt
            Swap great + 1, rightt
        Else
            Swap less - 1, leftt, True
            Swap great + 1, rightt, True
        End If
        'subarrays
        QuickSort_Asc leftt, less - 2, div, SortingKey, Indexed
        QuickSort_Asc great + 2, rightt, div, SortingKey, Indexed
        'equal elements
        If (dist > Ulen - 13 And pivot1 <> pivot2) Then
            If Not Indexed Then
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
            Else
                k = less
                Do While k <= great
                    If IndexedBuffer(k).ItemValue(SortingKey) = pivot1 Then
                        Swap k, less, True
                        less = less + 1
                    ElseIf IndexedBuffer(k).ItemValue(SortingKey) = pivot2 Then
                        Swap k, great, True
                        great = great - 1
                        If IndexedBuffer(k).ItemValue(SortingKey) = pivot1 Then
                            Swap k, less, True
                            less = less + 1
                        End If
                    End If
                    k = k + 1
                Loop
            End If
        End If
        'subarray
        If pivot1 < pivot2 Then
            QuickSort_Asc less, great, div, SortingKey, Indexed
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
                        Optional Indexed As Boolean = False)
    
    Dim Ulen As Long
    
    Ulen = rightt - leftt
    If Ulen < RUN Then 'insertion sort for tiny array
        InsertionSort leftt, rightt, SortingKey, True, Indexed
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
        If Not Indexed Then
            If Buffer(m1)(SortingKey) > Buffer(m2)(SortingKey) Then
                Swap m1, leftt
                Swap m2, rightt
            Else
                Swap m1, rightt
                Swap m2, leftt
            End If
        Else
            If IndexedBuffer(m1).ItemValue(SortingKey) > IndexedBuffer(m2).ItemValue(SortingKey) Then
                Swap m1, leftt, True
                Swap m2, rightt, True
            Else
                Swap m1, rightt, True
                Swap m2, leftt, True
            End If
        End If
        'pivots
        Dim pivot1 As Variant
        Dim pivot2 As Variant
        If Not Indexed Then
            pivot1 = Buffer(leftt)(SortingKey)
            pivot2 = Buffer(rightt)(SortingKey)
        Else
            pivot1 = IndexedBuffer(leftt).ItemValue(SortingKey)
            pivot2 = IndexedBuffer(rightt).ItemValue(SortingKey)
        End If
        'pointers
        Dim less As Long
        Dim great As Long
        less = leftt + 1
        great = rightt - 1
        'sorting
        Dim k As Long
        k = less
        If Not Indexed Then
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
        Else
            Do While k <= great
                If IndexedBuffer(k).ItemValue(SortingKey) > pivot1 Then
                    Swap k, less, True
                    less = less + 1
                ElseIf IndexedBuffer(k).ItemValue(SortingKey) < pivot2 Then
                    Do While k < great And IndexedBuffer(great).ItemValue(SortingKey) < pivot2
                        great = great - 1
                    Loop
                    Swap k, great, True
                    great = great - 1
                    If IndexedBuffer(k).ItemValue(SortingKey) > pivot1 Then
                        Swap k, less, True
                        less = less + 1
                    End If
                End If
                k = k + 1
            Loop
        End If
        'swaps
        Dim dist As Long
        dist = great - less
        If dist < 13 Then div = div + 1
        If Not Indexed Then
            Swap less - 1, leftt
            Swap great + 1, rightt
        Else
            Swap less - 1, leftt, True
            Swap great + 1, rightt, True
        End If
        'subarrays
        QuickSort_Desc leftt, less - 2, div, SortingKey, Indexed
        QuickSort_Desc great + 2, rightt, div, SortingKey, Indexed
        'equal elements
        If (dist > Ulen - 13 And pivot1 <> pivot2) Then
            If Not Indexed Then
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
            Else
                k = less
                Do While k <= great
                    If IndexedBuffer(k).ItemValue(SortingKey) = pivot1 Then
                        Swap k, less, True
                        less = less + 1
                    ElseIf IndexedBuffer(k).ItemValue(SortingKey) = pivot2 Then
                        Swap k, great, True
                        great = great - 1
                        If IndexedBuffer(k).ItemValue(SortingKey) = pivot1 Then
                            Swap k, less, True
                            less = less + 1
                        End If
                    End If
                    k = k + 1
                Loop
            End If
        End If
        'subarray
        If pivot1 > pivot2 Then
            QuickSort_Desc less, great, div, SortingKey, Indexed
        End If
    End If
End Sub

''' <summary>
''' Reduces the internal array list to the result by evaluate the
''' ReductionExpression over all items.
''' </summary>
''' <param name="ReductionExpression">The expression to eval and get reduction.</param>
''' <param name="StartIndex">The index to start filter records (base 1).</param>
Public Function Reduce(ReductionExpression As String, startIndex As Long, _
                        endIndex As Long, Optional Headers As Boolean = True) As Variant()
    Dim cPattern As String
    Dim Evaluator As CSVexpressions
    Dim evalWithOutVar As Boolean
    Dim rCounter As Long
    Dim TargetFields() As Long
    Dim tmpElement() As Variant
    Dim tmpResult() As Variant
    
    On Error GoTo Reduce_ErrHandler
    If P_INDEXING Then Exit Function 'Does not work over indexed buffers
    If startIndex < 1 Then Exit Function
    Set Evaluator = New CSVexpressions
    If Headers Then
        If startIndex = 1 Then
            startIndex = startIndex + 1
        End If
    End If
    If endIndex - 1 > CurrentIndex Then
        endIndex = CurrentIndex + 1
    End If
    ReDim tmpResult(0 To endIndex - startIndex)
    ReDim tmpElement(0 To 0)
    With Evaluator
        If Headers Then
            .Create SwitchUnderscoresAndSpaces(ReductionExpression, Buffer(0))
            evalWithOutVar = (.currentVariables = vbNullString)
            If Not evalWithOutVar Then
                TargetFields() = GetIndexesFromVarList(SwitchUnderscoresAndSpaces(.currentVariables, Buffer(0), ToUnderscores:=False), Buffer(0))
                'Replace
                cPattern = RebuildEvalExpression(.expression, .currentVariables, TargetFields)
                .Create cPattern
            End If
        Else
            .Create ReductionExpression
            evalWithOutVar = (.currentVariables = vbNullString)
            If Not evalWithOutVar Then
                TargetFields() = GetIndexesFromVarList(.currentVariables)
            End If
        End If
        For rCounter = startIndex - 1 To endIndex - 1
            If Not evalWithOutVar Then
                On Error Resume Next
                .Eval GetValuesForVariables(rCounter, TargetFields)
            Else
                On Error Resume Next
                .Eval
            End If
            tmpElement(0) = FormatEvalOutput(.result)
            tmpResult(0 + rCounter - startIndex + 1) = tmpElement  'reduce
        Next rCounter
    End With
    Reduce = tmpResult
Reduce_ErrHandler:
    Set Evaluator = Nothing
End Function

Private Function RearrangeCheckFieldsToOperate(fieldsToRearrange() As String) As Boolean
    Dim BufferColUB As Long
    Dim definedIdx As Long
    Dim fldOcount As Long
    Dim jkl As Long
    Dim tmpBool As Boolean
    
    BufferColUB = UBound(Buffer(0))
    For fldOcount = LBound(fieldsToRearrange) To UBound(fieldsToRearrange)
        jkl = -1
        Do
            jkl = jkl + 1
            tmpBool = CLng(fieldsToRearrange(fldOcount)) = jkl
        Loop While jkl <= BufferColUB And Not tmpBool
        If tmpBool Then
            definedIdx = definedIdx + 1
        End If
    Next fldOcount
    RearrangeCheckFieldsToOperate = (BufferColUB = definedIdx - 1)
End Function

''' <summary>
''' Rearranges the fields of the stored data.
''' </summary>
''' <param name="FieldsOrder">
''' String containing the desired order for all fields.
''' A string such as "0-3,5-4,6-11" used as a parameter will leave
''' the position of fields with indexes 0 to 3 unchanged, swap the fields
''' at indexes 5 and 4, and leave all remaining fields in position.
''' </param>
Public Function RearrangeFields(FieldsOrder As String) As CSVArrayList
    Dim BufferColUB As Long
    Dim cpRecord() As Variant
    Dim fCounter As Long
    Dim iRecCnt As Long
    Dim rCounter As Long
    Dim tmpFldOrder() As String
    Dim tmpRecord() As Variant
    
    On Error GoTo ErrHandler_RearrangeFields
    '@----------------------------------------------------------------------------
    'Check given fields arrange
    BufferColUB = UBound(Buffer(0))
    tmpFldOrder() = FieldsToOperate(FieldsOrder)
    If Not RearrangeCheckFieldsToOperate(tmpFldOrder) Then
        GoTo ErrHandler_RearrangeFields
    End If
    '@----------------------------------------------------------------------------
    'Rearrange
    ReDim cpRecord(0 To BufferColUB)
    For rCounter = 0 To CurrentIndex
        tmpRecord() = Buffer(rCounter)
        iRecCnt = 0
        For fCounter = LBound(tmpFldOrder) To UBound(tmpFldOrder)
            cpRecord(iRecCnt) = tmpRecord(CLng(tmpFldOrder(iRecCnt)))
            iRecCnt = iRecCnt + 1
        Next fCounter
        Buffer(rCounter) = cpRecord
    Next rCounter
    Set RearrangeFields = Me
ErrHandler_RearrangeFields:
End Function

Private Function RebuildEvalExpression(ByRef eExpression As String, _
                                        ByRef cVariables As String, _
                                        ByRef TargetFields() As Long) As String
    Dim tmpVars() As String
    Dim i As Long
    
    'Replace
    RebuildEvalExpression = eExpression
    tmpVars() = Split(cVariables, "; ")
    For i = LBound(tmpVars) To UBound(tmpVars)
        RebuildEvalExpression = Replace(RebuildEvalExpression, tmpVars(i), "f" & CStr(TargetFields(i)))
    Next i
End Function
''' <summary>
''' Reinitializes the current instance of the class and reserves user desired storage space.
''' </summary>
Public Sub Reinitialize(bufferSize As Long)
    Select Case bufferSize
        Case Is >= 0
            If Not P_INDEXING Then
                CurrentIndex = -1
                MaxIndex = bufferSize - 1
                ReDim Buffer(0 To MaxIndex)
            Else
                IndexedCurrentIndex = -1
                IndexedMaxIndex = bufferSize - 1
                ReDim IndexedBuffer(0 To IndexedMaxIndex)
            End If
        Case Else
            err.Raise 9
    End Select
End Sub

''' <summary>
''' Removes the Item at specified Index.
''' </summary>
''' <param name="Index">The index of the Item'll be deleted.</param>
Public Sub RemoveAt(index As Long)
Attribute RemoveAt.VB_Description = "Removes an Item at specified Index."
    RemoveRange index, 1
End Sub

''' <summary>
''' Removes the field at specified aIndex in all records.
''' </summary>
''' <param name="aIndex">The index of the field to be deleted.</param>
Public Function RemoveField(aIndex As Long) As CSVArrayList
    Dim ColUB As Long
    
    On Error GoTo ErrHandler_RemoveField
    ColUB = UBound(Buffer(0))
    Select Case aIndex
        Case 0 To ColUB
            Dim curRecord() As Variant
            Dim cpRecord() As Variant
            Dim fCounter As Long
            Dim rCounter As Long
            '@--------------------------------------------------------------------------------
            'Shrink storage
            ReDim cpRecord(0 To ColUB - 1)
            For rCounter = 0 To CurrentIndex
                curRecord() = Buffer(rCounter)
                For fCounter = 0 To ColUB
                    If fCounter < aIndex Then
                        cpRecord(fCounter) = curRecord(fCounter)
                    ElseIf fCounter > aIndex Then
                        cpRecord(fCounter - 1) = curRecord(fCounter)
                    End If
                Next fCounter
                Buffer(rCounter) = cpRecord
            Next rCounter
        Case Else
            GoTo OutOfBounds_RemoveField
    End Select
    Set RemoveField = Me
ErrHandler_RemoveField:
    Exit Function
OutOfBounds_RemoveField:
    Set RemoveField = Nothing
    err.Raise 9
    Resume ErrHandler_RemoveField
End Function

''' <summary>
''' Removes an indexed Item using the specified key.
''' </summary>
''' <param name="Key">Item's key.</param>
Public Sub RemoveIndexedItem(Key As String)
Attribute RemoveIndexedItem.VB_Description = "Removes a range of Items starting at specified Index."
    Dim index As Long
    Dim iCounter As Long
    Dim pointer As Long
    
    index = KeyIndex(Key)
    Select Case index
        Case Is > -1
            IndexedBuffer(index) = NullItem
            For iCounter = index + 1 To IndexedCurrentIndex
                Swap index + pointer, iCounter, True
                pointer = pointer + 1
            Next
            IndexedCurrentIndex = IndexedCurrentIndex - 1
        Case Else
            err.Raise 9
    End Select
End Sub

''' <summary>
''' Removes a range of Items starting at specified Index.
''' </summary>
''' <param name="Index">The index of the first item to be deleted.</param>
''' <param name="Count">The number of item to be deleted.</param>
Public Sub RemoveRange(index As Long, count As Long)
    Dim iCounter As Long
    Dim pointer As Long
    Dim endIndex As Long
    Dim curIdx As Long
    
    If Not P_INDEXING Then
        curIdx = CurrentIndex
    Else
        curIdx = IndexedCurrentIndex
    End If
    Select Case index
        Case 0 To curIdx
            endIndex = index + count - 1
            If Not P_INDEXING Then
                If endIndex > CurrentIndex Then
                    endIndex = CurrentIndex
                    count = endIndex - index + 1
                End If
                For iCounter = index To endIndex
                    Buffer(iCounter) = Null
                Next
                For iCounter = endIndex + 1 To CurrentIndex
                    Swap index + pointer, iCounter
                    pointer = pointer + 1
                Next
                CurrentIndex = CurrentIndex - count
            Else
                If endIndex > IndexedCurrentIndex Then
                    endIndex = IndexedCurrentIndex
                    count = endIndex - index + 1
                End If
                For iCounter = index To endIndex
                    IndexedBuffer(iCounter) = NullItem
                Next
                For iCounter = endIndex + 1 To IndexedCurrentIndex
                    Swap index + pointer, iCounter, True
                    pointer = pointer + 1
                Next
                IndexedCurrentIndex = IndexedCurrentIndex - count
            End If
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
                        Optional endIndex As Long = -1) As Boolean
Attribute Reverse.VB_Description = "Reverse the order of the internal items."
    Dim i As Long
    
    On Error GoTo Reverse_ErrHandler
    If startIndex < 0 Then
        startIndex = 0
    End If
    If endIndex < 0 Then
        If Not P_INDEXING Then
            endIndex = CurrentIndex
        Else
            endIndex = IndexedCurrentIndex
        End If
    End If
    If endIndex > startIndex Then
        For i = startIndex To (startIndex + ((endIndex - startIndex - 1) \ 2))
            Swap i, endIndex - (i - startIndex), P_INDEXING
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
                            Optional endIndex As Long = -1) As Boolean
Attribute Reverse2.VB_Description = "Reverse the order of items in the target array."
    Dim i As Long
    
    On Error GoTo Reverse_ErrHandler
    If startIndex < 0 Then
        startIndex = LBound(TargetArr)
    End If
    If endIndex < 0 Then
        endIndex = UBound(TargetArr)
    End If
    If endIndex > startIndex Then
        For i = startIndex To (startIndex + ((endIndex - startIndex - 1) \ 2))
            Swap2 TargetArr, i, endIndex - (i - startIndex)
        Next i
    End If
    Reverse2 = True
    Exit Function
Reverse_ErrHandler:
    Reverse2 = False
End Function

''' <summary>
''' Searches for a key in the internal indexed records
''' </summary>
''' <param name="TargetKey">The key to be searched.</param>
Public Function KeyExist(ByRef TargetKey As String) As Boolean
    KeyExist = (KeyIndex(TargetKey) > -1)
End Function
''' <summary>
''' Searches for an item in the internal indexed records, using a key,
''' of the current instance when the data is already sorted. Returns
''' the index of the item when found and -1 when missing values.
''' </summary>
''' <param name="TargetKey">The key to be searched.</param>
Public Function KeyIndex(ByRef TargetKey As String) As Long
    If IndexedCurrentIndex = -1 Then 'No items found
        KeyIndex = -1
    Else
        If Not SORTED_KEYS Then
            SortKeys
        End If
        Dim tmpR As Boolean
        Dim BottomIdx As Long
        Dim TopIdx As Long
        Dim MiddleIdx As Long
        
        BottomIdx = 0
        TopIdx = IndexedCurrentIndex
        Do
            MiddleIdx = BottomIdx + Ceiling((TopIdx - BottomIdx) / 2)
            tmpR = (TargetKey = IndexedBuffer(MiddleIdx).itemKey)
            If TargetKey > IndexedBuffer(MiddleIdx).itemKey Then
                BottomIdx = MiddleIdx + 1
            Else
                TopIdx = MiddleIdx - 1
            End If
        Loop While Not tmpR And TopIdx >= BottomIdx
        If tmpR Then
            KeyIndex = MiddleIdx
        Else
            KeyIndex = -1
        End If
    End If
End Function

''' <summary>
''' Moves a field leftward or rightward.
''' </summary>
''' <param name="aIndex">The index of the field to be shifted.</param>
''' <param name="Shift">Number of shifts. Negative values produce leftward shifts.</param>
Public Function ShiftField(aIndex As Long, Shift As Long) As CSVArrayList
    Dim ColUB As Long
    
    On Error GoTo ErrHandler_ShiftField
    ColUB = UBound(Buffer(0))
    Select Case aIndex + Shift
        Case 0 To ColUB
            Dim curRecord() As Variant
            Dim cpRecord() As Variant
            Dim rCounter As Long
            Dim ShiftLB As Long
            Dim ShiftUB As Long
            Dim fCounter As Long
            Dim ShiftFactor As Long
                
            ShiftFactor = 1 * Sgn(Shift)
            If ShiftFactor < 0 Then 'Shift to left
                ShiftLB = aIndex + Shift + 1
                ShiftUB = aIndex
            Else 'Shift to right
                ShiftLB = aIndex
                ShiftUB = aIndex + Shift - 1
            End If
            For rCounter = 0 To CurrentIndex
                curRecord() = Buffer(rCounter)
                cpRecord() = curRecord
                cpRecord(aIndex + Shift) = curRecord(aIndex)
                For fCounter = ShiftLB To ShiftUB
                    cpRecord(fCounter) = curRecord(fCounter + (1 * ShiftFactor))
                Next fCounter
                Buffer(rCounter) = cpRecord
            Next rCounter
        Case Else
            GoTo OutOfBounds_ShiftField
    End Select
    Set ShiftField = Me
ErrHandler_ShiftField:
    Exit Function
OutOfBounds_ShiftField:
    Set ShiftField = Nothing
    err.Raise 9
    Resume ErrHandler_ShiftField
End Function

''' <summary>
''' Moves a record upward or downward
''' </summary>
''' <param name="aIndex">The index of the record to be shifted.</param>
''' <param name="Shift">Number of shifts. Negative values produce upward shifts.</param>
Public Function ShiftRecord(aIndex As Long, Shift As Long) As CSVArrayList
    Dim ColUB As Long
    
    On Error GoTo ErrHandler_ShiftRecord
    ColUB = UBound(Buffer(0))
    Select Case aIndex + Shift
        Case 0 To CurrentIndex
            Dim cpRecord() As Variant
            Dim rCounter As Long
            Dim ShiftLB As Long
            Dim ShiftUB As Long
            Dim ShiftFactor As Long
                
            ShiftFactor = Sgn(Shift)
            If ShiftFactor < 0 Then 'Shift up
                ShiftLB = aIndex
                ShiftUB = aIndex + Shift + 1
            Else 'Shift down
                ShiftLB = aIndex
                ShiftUB = aIndex + Shift - 1
            End If
            cpRecord() = Buffer(aIndex)
            For rCounter = ShiftLB To ShiftUB Step ShiftFactor
                Buffer(rCounter) = Buffer(rCounter + (1 * ShiftFactor))
            Next rCounter
            Buffer(aIndex + Shift) = cpRecord
        Case Else
            GoTo OutOfBounds_ShiftRecord
    End Select
    Set ShiftRecord = Me
ErrHandler_ShiftRecord:
    Exit Function
OutOfBounds_ShiftRecord:
    Set ShiftRecord = Nothing
    err.Raise 9
    Resume ErrHandler_ShiftRecord
End Function

''' <summary>
''' Shrinks the buffer size to avoid extra space reservation.
''' </summary>
Public Sub ShrinkBuffer()
Attribute ShrinkBuffer.VB_Description = "Shrinks the buffer size to avoid extra space reservation."
    If CurrentIndex >= 0 Then
        If Not P_INDEXING Then
            MaxIndex = CurrentIndex
            ReDim Preserve Buffer(0 To MaxIndex)
        Else
            IndexedMaxIndex = IndexedCurrentIndex
            ReDim Preserve IndexedBuffer(0 To IndexedMaxIndex)
        End If
    End If
End Sub

Private Function GetColsToSort(sortingKeys As Variant) As Long()
    Dim i As Long
    Dim tmpResult() As Long
            
    If IsArray(sortingKeys) Then
        ReDim tmpResult(LBound(sortingKeys) To UBound(sortingKeys))
        For i = LBound(sortingKeys) To UBound(sortingKeys)
            tmpResult(i) = CLng(sortingKeys(i))
        Next i
    Else
        ReDim tmpResult(0 To 0)
        tmpResult(0) = CLng(sortingKeys)
    End If
    GetColsToSort = tmpResult
End Function
''' <summary>
''' Sorts the internal Items. Requires all Items to be one-dimensional arrays.
''' The indexes are base 0.
''' </summary>
''' <param name="fromIndex">The index of the first item to be sorted.</param>
''' <param name="toIndex">The index of the last item to be sorted.</param>
''' <param name="SortingKeys">The columns/keys for the logical comparisons.</param>
''' <param name="sortAlgorithm">Algorithm used to sort the data.</param>
Public Function Sort(Optional ByVal fromIndex As Long = -1, _
                Optional ByVal toIndex As Long = -1, _
                Optional ByVal sortingKeys As Variant = 1, _
                Optional ByVal sortAlgorithm As SortingAlgorithms = SortingAlgorithms.SA_Quicksort) As CSVArrayList
    On Error GoTo Sort_Error_Handler:
    Dim tmpIndex As Long
    
    If Not P_INDEXING Then
        tmpIndex = CurrentIndex
    Else
        tmpIndex = IndexedCurrentIndex
    End If
    If tmpIndex + 1 > 0 Then
        If fromIndex = -1 Then
            fromIndex = 0
        Else
            fromIndex = fromIndex - 1 'Zero based
        End If
        If toIndex = -1 Then
            If Not P_INDEXING Then
                toIndex = CurrentIndex
            Else
                toIndex = IndexedCurrentIndex
            End If
        Else
            toIndex = toIndex - 1
        End If
        If (CheckSortingBounds(fromIndex, toIndex, P_INDEXING) And CheckSortingKeys(sortingKeys, P_INDEXING)) Then
            Dim i As Long
            Dim ColumnsToSort() As Long
            ColumnsToSort() = GetColsToSort(sortingKeys)
            'Sort
            keysLB = LBound(ColumnsToSort)
            keysUB = UBound(ColumnsToSort)
            If Not P_INDEXING Then
                If sortAlgorithm = SA_Quicksort Then 'QuickSort
                    QuickSort fromIndex, toIndex, 3, Abs(ColumnsToSort(keysLB)) - 1, ColumnsToSort(keysLB) < 0, P_INDEXING
                    For i = keysLB + 1 To keysUB
                        StableQSort fromIndex, toIndex, Abs(ColumnsToSort(i)) - 1, _
                                    Abs(ColumnsToSort(i - 1)) - 1, ColumnsToSort(i) < 0
                    Next i
                Else
                    If sortAlgorithm = SA_TimSort Then 'TimSort
                        For i = keysLB To keysUB
                            TimSort fromIndex, toIndex, Abs(ColumnsToSort(i)) - 1, ColumnsToSort(i) < 0
                        Next i
                    Else
                        If sortAlgorithm = SA_HeapSort Then 'HeapSort
                            heapSort fromIndex, toIndex, Abs(ColumnsToSort(keysLB)) - 1, ColumnsToSort(keysLB) < 0
                            For i = keysLB + 1 To keysUB
                                StableHeapSort fromIndex, toIndex, Abs(ColumnsToSort(i)) - 1, _
                                                Abs(ColumnsToSort(i - 1)) - 1, ColumnsToSort(i) < 0
                            Next i
                        Else 'MergeSort
                            For i = keysLB To keysUB
                                MergeSort fromIndex, toIndex, Abs(ColumnsToSort(i)) - 1, ColumnsToSort(i) < 0
                            Next i
                        End If
                    End If
                End If
            Else 'Sort indexed items
                QuickSort fromIndex, toIndex, 3, Abs(ColumnsToSort(keysLB)) - 1, ColumnsToSort(keysLB) < 0, P_INDEXING
                For i = keysLB + 1 To keysUB
                    StableQSort fromIndex, toIndex, Abs(ColumnsToSort(i)) - 1, _
                                    Abs(ColumnsToSort(i - 1)) - 1, ColumnsToSort(i) < 0
                Next i
            End If
            If Not P_INDEXING Then
                SORTED_DATA = True
                LAST_SORTED_FIELD = ColumnsToSort(keysUB)
            Else
                INDEXED_SORTED_DATA = True
            End If
        End If
    End If
    Set Sort = Me
    Exit Function
Sort_Error_Handler:
End Function

''' <summary>
''' Sorts the internal items by an specified field.
''' </summary>
''' <param name="fromIndex">The index of the first item to be sorted.</param>
''' <param name="toIndex">The index of the last item to be sorted.</param>
''' <param name="SortingKey">The row/record for the logical comparisons.</param>
''' <param name="sortAlgorithm">Algorithm used to sort the data.</param>
Public Function SortByField(Optional ByVal fromIndex As Long = -1, _
                            Optional ByVal toIndex As Long = -1, _
                            Optional ByVal SortingKey As Long = 1, _
                            Optional ByVal sortAlgorithm As SortingAlgorithms = SortingAlgorithms.SA_Quicksort) As CSVArrayList
    
    Dim CurField As CSVArrayList
    Dim CurFieldArray() As Variant
    Dim sfIcounter As Long
    
    On Error GoTo SortByField_Error_Handler:
    If fromIndex = -1 Then
        fromIndex = 1
    End If
    Set CurField = New CSVArrayList
    CurFieldArray() = Buffer(Abs(SortingKey) - 1)
    If toIndex = -1 Then
        toIndex = UBound(CurFieldArray) + 1
    End If
    For sfIcounter = LBound(CurFieldArray) To UBound(CurFieldArray)
        CurField.Add2 CurFieldArray(sfIcounter)
    Next sfIcounter
    CurField.ShrinkBuffer
    CurField.Sort fromIndex, toIndex, Sgn(SortingKey) * 1, sortAlgorithm
    RearrangeFields GetRearrangeSortStr(CurFieldArray, CurField)
    Set SortByField = Me
SortByField_Error_Handler:
End Function

''' <summary>
''' Sorts the internal Items by its keys using QuickSort. Requires
''' all Items to be one-dimensional arrays. The indexes are base 0.
''' </summary>
''' <param name="fromIndex">The index of the first item to be sorted.</param>
''' <param name="toIndex">The index of the last item to be sorted.</param>
Public Sub SortKeys(Optional ByVal fromIndex As Long = -1, _
                Optional ByVal toIndex As Long = -1)
                
    On Error GoTo Sort_Error_Handler:
    If IndexedMaxIndex + 1 > 0 Then
        If fromIndex = -1 Then
            fromIndex = 0
        Else
            fromIndex = fromIndex - 1 'Zero based
        End If
        If toIndex = -1 Then
            toIndex = IndexedCurrentIndex
        Else
            toIndex = toIndex - 1
        End If
    End If
    If CheckSortingBounds(fromIndex, toIndex, True) Then
        SortKeys_ fromIndex, toIndex, 3
        SORTED_KEYS = True
    End If
    Exit Sub
Sort_Error_Handler:
End Sub
''' <summary>
''' Sorts the keys on current instance.
''' </summary>
Private Sub SortKeys_(leftt As Long, rightt As Long, div As Long)
    
    Dim Ulen As Long

    Ulen = rightt - leftt
    If Ulen < RUN Then 'insertion sort for tiny array
        InsertionSort leftt, rightt, 0, False, True, True
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
        If IndexedBuffer(m1).itemKey < IndexedBuffer(m2).itemKey Then
            Swap m1, leftt, True
            Swap m2, rightt, True
        Else
            Swap m1, rightt, True
            Swap m2, leftt, True
        End If
        'pivots
        Dim pivot1 As Variant
        Dim pivot2 As Variant
        pivot1 = IndexedBuffer(leftt).itemKey
        pivot2 = IndexedBuffer(rightt).itemKey
        'pointers
        Dim less As Long
        Dim great As Long
        less = leftt + 1
        great = rightt - 1
        'sorting
        Dim k As Long
        k = less
        Do While k <= great
            If IndexedBuffer(k).itemKey < pivot1 Then
                Swap k, less, True
                less = less + 1
            ElseIf IndexedBuffer(k).itemKey > pivot2 Then
                Do While k < great And IndexedBuffer(great).itemKey > pivot2
                    great = great - 1
                Loop
                Swap k, great, True
                great = great - 1
                If IndexedBuffer(k).itemKey < pivot1 Then
                    Swap k, less, True
                    less = less + 1
                End If
            End If
            k = k + 1
        Loop
        'swaps
        Dim dist As Long
        dist = great - less
        If dist < 13 Then div = div + 1
        Swap less - 1, leftt, True
        Swap great + 1, rightt, True
        'subarrays
        SortKeys_ leftt, less - 2, div
        SortKeys_ great + 2, rightt, div
        'equal elements
        If (dist > Ulen - 13 And pivot1 <> pivot2) Then
            k = less
            Do While k <= great
                If IndexedBuffer(k).itemKey = pivot1 Then
                    Swap k, less, True
                    less = less + 1
                ElseIf IndexedBuffer(k).itemKey = pivot2 Then
                    Swap k, great, True
                    great = great - 1
                    If IndexedBuffer(k).itemKey = pivot1 Then
                        Swap k, less, True
                        less = less + 1
                    End If
                End If
                k = k + 1
            Loop
        End If
        'subarray
        If pivot1 < pivot2 Then
            SortKeys_ less, great, div
        End If
    End If
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
Private Sub StableHeapSort(fromIndex As Long, _
                            toIndex As Long, _
                            NewKey As Long, _
                            OldKey As Long, _
                            Descending As Boolean)
    Dim FIndex As Long
    Dim lIndex As Long
        
    FIndex = fromIndex
    lIndex = EndOfSortingBlock(OldKey, FIndex, toIndex)
    Do While FIndex <= toIndex
        heapSort FIndex, lIndex, NewKey, Descending
        FIndex = lIndex + 1
        lIndex = EndOfSortingBlock(OldKey, FIndex, toIndex)
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
Private Sub StableQSort(fromIndex As Long, _
                        toIndex As Long, _
                        NewKey As Long, _
                        OldKey As Long, _
                        Descending As Boolean)
    Dim FIndex As Long
    Dim lIndex As Long
        
    FIndex = fromIndex
    lIndex = EndOfSortingBlock(OldKey, FIndex, toIndex, P_INDEXING)
    Do While FIndex <= toIndex
        QuickSort FIndex, lIndex, 3, NewKey, Descending, P_INDEXING
        FIndex = lIndex + 1
        lIndex = EndOfSortingBlock(OldKey, FIndex, toIndex, P_INDEXING)
    Loop
End Sub

''' <summary>
''' Swap items in buffer.
''' </summary>
Public Sub Swap(position1 As Long, position2 As Long, _
                Optional Indexed As Boolean = False)
    If Not Indexed Then
        If CurrentIndex >= 0 Then
            tmpBuffer = Buffer(position1)
            Buffer(position1) = Buffer(position2)
            Buffer(position2) = tmpBuffer
        End If
    Else
        If IndexedCurrentIndex >= 0 Then
            tmpIndexedBuffer = IndexedBuffer(position1)
            IndexedBuffer(position1) = IndexedBuffer(position2)
            IndexedBuffer(position2) = tmpIndexedBuffer
        End If
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
    Dim r As Long
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
            r = Min(l + 2 * Size - 1, rightt)
            ' merge sub array arr[L.....midd] &
            ' arr[midd+1....R]
            If midd < r Then
                merge l, midd, r, SortingKey, Descending
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
