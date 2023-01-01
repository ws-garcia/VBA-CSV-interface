VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVexpressionsScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' Copyright © 2022-2023 W. García
' GPL-3.0 license | https://www.gnu.org/licenses/gpl-3.0.html/
' https://ingwilfredogarcia.wordpress.com
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
' INFO:
' Class module developed to define constants and variable scopes for expressions, so that several expressions
' can share the same variables without the need to redefine each of them.
'
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' VARIABLES:
Private e As Double
Private PI As Double
Private P_CONSTANTS As ClusterBuffer
Private P_EXPR_VARIABLES As ClusterBuffer
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' TYPES:
Private Type ClusterItem
    index As Long
    name As String
    value As String
    Assigned As Boolean
End Type
Private Type ClusterBuffer
    Capacity As Long
    index As Long
    Storage() As ClusterItem
End Type

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES:


''' <summary>
''' Gets the constants collection. By default this is pre-populated with PI and e.
''' </summary>
Public Property Get Constants() As Collection
    Dim i As Long
    
    Set Constants = New Collection
    For i = 0 To P_CONSTANTS.index
        Constants.Add P_CONSTANTS.Storage(i).value, P_CONSTANTS.Storage(i).name
    Next i
End Property

''' <summary>
''' Gets a string with the variables values used for the last evaluation.
''' </summary>
Public Property Get CurrentVarValues() As String
    Dim i As Long
    Dim tmpResult As String
    
    For i = 0 To P_EXPR_VARIABLES.index
        If tmpResult = vbNullString Then
            tmpResult = P_EXPR_VARIABLES.Storage(i).name & " = " & _
                        P_EXPR_VARIABLES.Storage(i).value
        Else
            tmpResult = tmpResult & "; " & P_EXPR_VARIABLES.Storage(i).name & " = " & _
                        P_EXPR_VARIABLES.Storage(i).value
        End If
    Next i
    CurrentVarValues = tmpResult
End Property

''' <summary>
''' Gets a string with the variables values used for the last evaluation.
''' </summary>
Public Property Get CurrentVariables() As String
    Dim i As Long
    Dim tmpResult As String
    
    For i = 0 To P_EXPR_VARIABLES.index
        If tmpResult = vbNullString Then
            tmpResult = P_EXPR_VARIABLES.Storage(i).name
        Else
            tmpResult = tmpResult & "; " & P_EXPR_VARIABLES.Storage(i).name
        End If
    Next i
    CurrentVariables = tmpResult
End Property

''' <summary>
''' Returns 'True' if all stored variables have a value or when there
''' are no stored variables, 'False' when the value of any variable is missing.
''' </summary>
Public Function DefinedScope() As Boolean
    If P_EXPR_VARIABLES.index > -1 Then
        Dim i As Long
        Dim cviCounter As Long
        
        For i = 0 To P_EXPR_VARIABLES.index
            If P_EXPR_VARIABLES.Storage(i).Assigned Then
                cviCounter = cviCounter + 1
            End If
        Next i
        DefinedScope = (cviCounter = P_EXPR_VARIABLES.index + 1)
    Else
        DefinedScope = True
    End If
End Function

''' <summary>
''' Gets the count of stored variable.
''' </summary>
Public Property Get VariablesCount() As Long
    VariablesCount = P_EXPR_VARIABLES.index + 1
End Property

''' <summary>
''' Gets or sets the current value from/to the given variable.
''' If 'aVarKey' is numeric, the argument is treated as an index.
''' </summary>
Public Property Get VarValue(aVarKey As Variant) As String
    If Not IsNumeric(aVarKey) Then
        Dim ValueIdx As Long
    
        ValueIdx = GetCBItemIdx(P_EXPR_VARIABLES, CStr(aVarKey))
        If ValueIdx > -1 Then
            VarValue = P_EXPR_VARIABLES.Storage(ValueIdx).value
        End If
    Else
        VarValue = P_EXPR_VARIABLES.Storage(CLng(aVarKey)).value
    End If
End Property

Public Property Let VarValue(aVarKey As Variant, aVarValue As String)
    If Not IsNumeric(aVarKey) Then
        Dim ValueIdx As Long
    
        ValueIdx = GetCBItemIdx(P_EXPR_VARIABLES, CStr(aVarKey))
        If ValueIdx > -1 Then
            P_EXPR_VARIABLES.Storage(ValueIdx).value = aVarValue
            P_EXPR_VARIABLES.Storage(ValueIdx).Assigned = True
        End If
    Else
        ValueIdx = CLng(aVarKey)
        If ValueIdx > -1 And ValueIdx <= P_EXPR_VARIABLES.index Then
            P_EXPR_VARIABLES.Storage(ValueIdx).value = aVarValue
            P_EXPR_VARIABLES.Storage(ValueIdx).Assigned = True
        End If
    End If
End Property

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' METHODS:

Public Sub AddConstant(aValue As String, aKey As String)
    Dim ConstIdx As Long
    
    ConstIdx = GetCBItemIdx(P_CONSTANTS, aKey)
    If ConstIdx = -1 Then 'Ensure uniqueness
        AppendToCBbuffer P_CONSTANTS, aKey, aValue
    End If
End Sub

Public Sub AddVariable(ByRef variable As String, ByRef aKey As String)
    Dim varIdx As Long
    
    aKey = Cast(variable)
    varIdx = GetCBItemIdx(P_EXPR_VARIABLES, aKey)
    If varIdx = -1 Then 'Ensure uniqueness
        AppendToCBbuffer P_EXPR_VARIABLES, aKey
        varIdx = GetCBItemIdx(P_CONSTANTS, aKey)
        If varIdx > -1 Then 'Assign the value from constants
            P_EXPR_VARIABLES.Storage(P_EXPR_VARIABLES.index).value = P_CONSTANTS.Storage(varIdx).value
            P_EXPR_VARIABLES.Storage(P_EXPR_VARIABLES.index).Assigned = True
        End If
    End If
End Sub

Private Sub AppendToCBbuffer(ByRef aBuffer As ClusterBuffer, ByRef ItemName As String, Optional ByRef ItemValue As String = vbNullString)
    aBuffer.index = aBuffer.index + 1
    On Error GoTo BufferAppend_errHandler
    aBuffer.Storage(aBuffer.index).index = aBuffer.index
    aBuffer.Storage(aBuffer.index).Assigned = (ItemValue <> vbNullString)
    aBuffer.Storage(aBuffer.index).name = ItemName
    aBuffer.Storage(aBuffer.index).value = ItemValue
    Exit Sub
BufferAppend_errHandler:
    err.Clear
    ExpandCBbuffer aBuffer
    aBuffer.Storage(aBuffer.index).index = aBuffer.index
    aBuffer.Storage(aBuffer.index).Assigned = (ItemValue <> vbNullString)
    aBuffer.Storage(aBuffer.index).name = ItemName
    aBuffer.Storage(aBuffer.index).value = ItemValue
End Sub

Private Function Cast(ByRef expression As String) As String
    If IsPI(expression) Then
        Cast = LCase$(expression)     'Case insensitive for PI
    Else
        Cast = expression
    End If
End Function

Public Sub ConstantsInit()
    InitCBbuffer P_CONSTANTS
    '@--------------------------------------------------------------------
    ' Save predefined constants
    AppendToCBbuffer P_CONSTANTS, "pi", CStr(PI)
    AppendToCBbuffer P_CONSTANTS, "e", CStr(e)
End Sub

Public Function CopyScope(ByRef sourceScope As CSVexpressionsScope) As CSVexpressionsScope
    Dim i As Long
    Dim tmpValue As String
    Dim tmpIdx As Long
    
    For i = 0 To P_EXPR_VARIABLES.index
        tmpIdx = sourceScope.VarIndex(P_EXPR_VARIABLES.Storage(i).name)
        If tmpIdx > -1 Then
            tmpValue = sourceScope.VarValue(tmpIdx)
            If tmpValue <> vbNullString Then
                VarValue(i) = tmpValue
            End If
        End If
    Next
    Set CopyScope = Me
End Function
Private Sub ExpandCBbuffer(ByRef aBuffer As ClusterBuffer)
    aBuffer.Capacity = 2 * (aBuffer.Capacity + 1)
    ReDim Preserve aBuffer.Storage(0 To aBuffer.Capacity - 1)
End Sub

Public Sub FillPredefinedVars()
    Dim i As Long
    Dim tmpIdx As Long
    
    For i = 0 To P_EXPR_VARIABLES.index
        tmpIdx = GetCBItemIdx(P_CONSTANTS, P_EXPR_VARIABLES.Storage(i).name)
        If tmpIdx > -1 Then 'The variable is a defined constant
            P_EXPR_VARIABLES.Storage(i).value = P_CONSTANTS.Storage(tmpIdx).value
            P_EXPR_VARIABLES.Storage(i).Assigned = True
        End If
    Next i
End Sub
Private Function GetCBItemIdx(ByRef cbBuffer As ClusterBuffer, ByRef ItemName As String) As Long
    Dim i As Long
    Dim tmpResult As Boolean
    Dim tmpItemName As String
    
    If LenB(ItemName) Then
        Select Case AscW(ItemName)
            Case 43, 45 'Unary expression
                tmpItemName = MidB$(ItemName, 3)
            Case Else
                tmpItemName = ItemName
        End Select
        If cbBuffer.index > -1 Then
            i = 0
            Do
                tmpResult = (cbBuffer.Storage(i).name = tmpItemName)
                i = i + 1
            Loop While i <= cbBuffer.index And Not tmpResult
        End If
        If tmpResult Then
            GetCBItemIdx = i - 1
        Else
            GetCBItemIdx = -1
        End If
    Else
        GetCBItemIdx = -1
    End If
End Function

Private Sub InitCBbuffer(ByRef aBuffer As ClusterBuffer)
    aBuffer.Capacity = 128
    ReDim aBuffer.Storage(0 To aBuffer.Capacity - 1)
    aBuffer.index = -1
End Sub

Public Function IsConstant(aVarName As String) As Boolean
    IsConstant = GetCBItemIdx(P_CONSTANTS, aVarName) > -1
End Function

Private Function IsPI(ByRef varName As String) As Boolean
    IsPI = varName Like "[Pp][Ii]"
End Function

Public Sub VariablesInit()
    InitCBbuffer P_EXPR_VARIABLES
End Sub

''' <summary>
''' Returns the index of a given variable or constant named 'ItemName'
''' </summary>
Public Function VarIndex(ByRef ItemName As String, _
                            Optional findConstant As Boolean = False) As Long
    If Not findConstant Then
        VarIndex = GetCBItemIdx(P_EXPR_VARIABLES, ItemName)
    Else
        VarIndex = GetCBItemIdx(P_CONSTANTS, ItemName)
    End If
End Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''' <summary>
''' Initializes this instance.
''' </summary>
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
    e = Exp(1)
    ConstantsInit
    VariablesInit
End Sub
