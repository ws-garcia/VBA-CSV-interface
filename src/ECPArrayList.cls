VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ECPArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' Copyright © 2021 W. García
' GPL-3.0 license | https://www.gnu.org/licenses/gpl-3.0.html/
' https://ingwilfredogarcia.wordpress.com
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' GENERAL INFO:
' Class module developed to emulate some functionalities from the ArrayList present in most
' modern languages.
'////////////////////////////////////////////////////////////////////////////////////////////
'#
Option Explicit
Private Buffer() As Variant
Private CurrentIndex As Long
Private MaxIndex As Long
Private tmpBuffer As Variant

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
''' Gets the Item, by its index, from the current instance.
''' </summary>
''' <param name="Index">Item's position in this instance.</param>
Public Property Get item(Index As Long) As Variant
Attribute item.VB_Description = "Gets the Item, by its index, from the current instance."
Attribute item.VB_UserMemId = 0
    Select Case Index
        Case 0 To CurrentIndex
            item = Buffer(Index)
        Case Else
            Err.Raise 9
    End Select
End Property

''' <summary>
''' Gets an array with all the Items from the current instance.
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
    CurrentIndex = CurrentIndex + 1
    On Error GoTo Expand_Buffer
    Buffer(CurrentIndex) = aValue
    Exit Sub
Expand_Buffer:
    MaxIndex = 2 * (MaxIndex + 1) - 1
    ReDim Preserve Buffer(0 To MaxIndex)
    Buffer(CurrentIndex) = aValue
End Sub

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
''' Inserts an Item, at the given Index, in the current instance of the class.
''' </summary>
''' <param name="Index">The index into which the Item'll be inserted.</param>
''' <param name="aValue">The value to be inserted.</param>
Public Sub Insert(Index As Long, aValue As Variant)
Attribute Insert.VB_Description = "Inserts an Item, at the given Index, in the current instance of the class."
    Dim tmpCopy() As Variant
    Dim iCounter As Long
        
    Select Case Index
        Case 0 To CurrentIndex + 1 'Avoids to leave empty items
            Me.Add aValue
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
            Err.Raise 9
    End Select
End Sub

''' <summary>
''' Sorts the data. Requires all elements to be one-dimensional arrays.
''' </summary>
''' <param name="leftt">The index of the first item to be sorted.</param>
''' <param name="rightt">The index of the last item to be sorted.</param>
''' <param name="div">Indicates the number of parts for dividing the data.</param>
''' <param name="SortColumn">The column for the logical comparisons.</param>
''' <param name="Descending">Sort order.</param>
Private Sub Quicksort(leftt As Long, _
                        rightt As Long, _
                        div As Long, _
                        SortColumn As Long, _
                        Optional Descending As Boolean = False)
                        
    Dim Ulen As Long
    Dim i As Long, j As Long

    Static RecursionLevel As Long
    RecursionLevel = RecursionLevel + 1
    Ulen = rightt - leftt
    If Ulen < 27 Then 'insertion sort for tiny array
        i = leftt + 1
        Do While i <= rightt
            j = i
            Do While j > leftt
                If Buffer(j)(SortColumn) < Buffer(j - 1)(SortColumn) Then
                    Swap j, j - 1
                End If
                j = j - 1
            Loop
            i = i + 1
        Loop
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
        If Buffer(m1)(SortColumn) < Buffer(m2)(SortColumn) Then
            Swap m1, leftt
            Swap m2, rightt
        Else
            Swap m1, rightt
            Swap m2, leftt
        End If
        'pivots
        Dim pivot1 As Variant
        Dim pivot2 As Variant
        pivot1 = Buffer(leftt)(SortColumn)
        pivot2 = Buffer(rightt)(SortColumn)
        'pointers
        Dim less As Long
        Dim great As Long
        less = leftt + 1
        great = rightt - 1
        'sorting
        Dim k As Long
        k = less
        Do While k <= great
            If Buffer(k)(SortColumn) < pivot1 Then
                Swap k, less
                less = less + 1
            ElseIf Buffer(k)(SortColumn) > pivot2 Then
                Do While k < great And Buffer(great)(SortColumn) > pivot2
                    great = great - 1
                Loop
                Swap k, great
                great = great - 1
                If Buffer(k)(SortColumn) < pivot1 Then
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
        Quicksort leftt, less - 2, div, SortColumn, Descending
        Quicksort great + 2, rightt, div, SortColumn, Descending
        'equal elements
        If (dist > Ulen - 13 And pivot1 <> pivot2) Then
            k = less
            Do While k <= great
                If Buffer(k)(SortColumn) = pivot1 Then
                    Swap k, less
                    less = less + 1
                ElseIf Buffer(k)(SortColumn) = pivot2 Then
                    Swap k, great
                    great = great - 1
                    If Buffer(k)(SortColumn) = pivot1 Then
                        Swap k, less
                        less = less + 1
                    End If
                End If
                k = k + 1
            Loop
        End If
        'subarray
        If pivot1 < pivot2 Then
            Quicksort less, great, div, SortColumn, Descending
        End If
    End If
    If Descending = True Then
        If RecursionLevel = 1 Then
            ReverseOrder leftt, rightt
        End If
    End If
    RecursionLevel = RecursionLevel - 1
End Sub

''' <summary>
''' Reinitializes the current instance of the class and reserve user desired storage space.
''' </summary>
Public Sub Reinitialize(bufferSize As Long)
Attribute Reinitialize.VB_Description = "Reinitializes the current instance of the class and reserve user desired storage space."
    Select Case bufferSize
        Case Is >= 0
            CurrentIndex = -1
            MaxIndex = bufferSize - 1
            ReDim Buffer(0 To MaxIndex)
        Case Else
            Err.Raise 9
    End Select
End Sub

''' <summary>
''' Removes an Item at specified Index.
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
            Err.Raise 9
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
            Err.Raise 9
    End Select
End Sub

''' <summary>
''' Order items in reverse.
''' </summary>
''' <param name="LB">The index of the first item to be reversed.</param>
''' <param name="UB">The number of item to be reversed.</param>
Private Function ReverseOrder(Optional LB As Long = -1, _
                                Optional UB As Long = -1) As Boolean
                                
    Dim i As Long, tmpVar As Variant
    
    On Error Resume Next
    If LB < 0 Then
        LB = 0
    End If
    If UB < 0 Then
        UB = CurrentIndex
    End If
    If UB > LB Then
        For i = LB To (LB + ((UB - LB - 1) \ 2))
            Swap i, UB - (i - LB)
        Next i
    End If
    If Err.Number = 0 Then
        ReverseOrder = True
    Else
        ReverseOrder = False
    End If
End Function


''' <summary>
''' Shrinks the buffer size to avoid extra space reservation.
''' </summary>
Public Sub ShrinkBuffer()
Attribute ShrinkBuffer.VB_Description = "Shrinks the buffer size to avoid extra space reservation"
    If CurrentIndex >= 0 Then
        MaxIndex = CurrentIndex
        ReDim Preserve Buffer(0 To MaxIndex)
    End If
End Sub

''' <summary>
''' Sorts the data. Requires all elements to be one-dimensional arrays.
''' </summary>
''' <param name="fromIndex">The index of the first item to be sorted.</param>
''' <param name="toIndex">The index of the last item to be sorted.</param>
''' <param name="SortColumn">The column for the logical comparisons.</param>
''' <param name="Descending">Sort order.</param>
Public Sub Sort(Optional ByVal fromIndex As Long = -1, _
                Optional ByVal toIndex As Long = -1, _
                Optional ByVal SortColumn As Long = -1, _
                Optional ByVal Descending As Boolean = False)
    
    On Error GoTo Sort_Error_Handler:
    Dim ParametersWellFormed As Boolean
    
    If CurrentIndex + 1 > 0 Then
        ParametersWellFormed = True
        If fromIndex = -1 Then
            fromIndex = 0
        Else
            If fromIndex <= CurrentIndex + 1 And fromIndex > 0 Then
                fromIndex = fromIndex - 1 'Zero based
            Else
                ParametersWellFormed = False
            End If
        End If
        If toIndex = -1 Then
            toIndex = CurrentIndex
        Else
            If toIndex <= CurrentIndex + 1 And toIndex > 0 Then
                toIndex = toIndex - 1
            Else
                ParametersWellFormed = False
            End If
        End If
        If fromIndex > toIndex Then
            ParametersWellFormed = False
        End If
        If SortColumn = -1 Then
                SortColumn = 0 'Ensures a column to sort
        ElseIf SortColumn < LBound(Buffer(0)) + 1 Or _
                SortColumn > UBound(Buffer(0)) + 1 Then
                ParametersWellFormed = False
        Else
            SortColumn = SortColumn - 1
        End If
        If ParametersWellFormed Then
            Quicksort fromIndex, toIndex, 3, SortColumn, Descending
        End If
    End If
    Exit Sub
Sort_Error_Handler:
End Sub

''' <summary>
''' Swap items in buffer.
''' </summary>
Public Sub Swap(position1 As Long, position2 As Long)
Attribute Swap.VB_Description = "Swap items in buffer"
    If CurrentIndex >= 0 Then
        tmpBuffer = Buffer(position1)
        Buffer(position1) = Buffer(position2)
        Buffer(position2) = tmpBuffer
    End If
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////
'#
