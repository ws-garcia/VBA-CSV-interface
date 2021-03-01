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
