VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private P_VECTOR() As String
Private P_UBOUND As Long
Public Property Get Item(Index As Long) As String
Attribute Item.VB_Description = "Gets data from the internal vector."
Attribute Item.VB_UserMemId = 0
	Item = P_VECTOR(Index)
End Property
Public Property Let Item(Index As Long, Value As String)
	P_VECTOR(Index) = Value
End Property
Public Sub Resize(Value As Long)
Attribute Resize.VB_Description = "Resizes the internal vector preserving previously stored data."
	Redim Preserve P_VECTOR(0 To Value)
	P_UBOUND = Value
End Sub
Public Property Get vUBound() As Long
Attribute vUBound.VB_Description = "Gets the internal vector upper bound."
	vUBound = P_UBOUND
End Property