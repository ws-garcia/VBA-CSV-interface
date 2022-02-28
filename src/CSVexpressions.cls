VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVexpressions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
Option Explicit
'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' Copyright © 2022 W. García
' GPL-3.0 license | https://www.gnu.org/licenses/gpl-3.0.html/
' https://ingwilfredogarcia.wordpress.com
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
' GENERAL INFO:
' Class module developed to evaluate mathematical expressions using VBA strings. The
' CSVexpressions class serves as an intermediate between the user interfaces and the main
' VBA/custom functions exposed through it. The module can evaluate mathematical expressions
' such as:
'
'            (a) 5*avg(2;abs(-3-7*tan(5));9)-12*pi-e+(7/sin(30)-4!)*min(cos(30);cos(150))
'            (b) min(cos(sin(30))+2^2;1)
'           *(c) GCD(1280;240;100;30*cos(0);10*DET({{sin(atn(1)*2); 0; 0}; {0; 2; 0}; {0; 0; 3}}))
'
' The main objective of the development of the class is to demonstrate that all mathematical
' expressions can be evaluated computationally using an approach similar to the one we humans
' use: divide the function into sub-expressions, create a symbolic string (lambda) to draw the
' evaluation flow of the expression, split the sub-expressions into chunks of operations
' (tokens) by tokenization, evaluate all the tokens.
'
' For example, the expressions given in (a), (b) and (c) will be translated into lambda strings:
'           (a) 5*avg{6}-12*pi-e+{3}*min{2}
'           (b) min{2}
'          *(c) GCD{4}
'
'               *GCD and DET are user-defined functions (UDF).
'
' Allowed expressions must follow the following grammar:
'
' Expression    =     ([{"("}]  SubExpr [{Operator [{"("}] SubExpr [{")"}]}] [{")"}] | {["("] ["{"] List [{";" List}] ["}"] [")"]}
' SubExpr       =     Token [{Operator Token}]
' Token         =     [{Unary}] Argument [(Operator | Function) ["("] [{Unary}] [Argument] [")"]]
' Argument      =     (List | Variable | Operand)
' List          =     "{" ["{"] SubExpr [{";" SubExpr}] ["}"] "}"
' Unary         =     "-" | "+" | ~
' Operand       =     ({Digit} ["."] [{Digit}] ["E"("-" | "+"){Digit}] | (True | False) | "'"Alphabet"'")
' Variable      =     Alphabet [{Decimal}] [{(Digit | Alphabet)}]
' Alphabet      =     "A-Z" | "a-z"
' Decimal       =     "."
' Digit = "0-9"
' Operator      =     "+" | "-" | "*" | "/" | "\" | "^" | "%" | "!" | "<" | "<=" | "<>" | ">" | ">=" | "=" | "&" | "|" | "||"
' Function      =     "abs" | "sin" | "cos" | "min" |...|[UDF]
'
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
' ABOUT THE ORDER IN WHICH MATHEMATICAL EXPRESSIONS ARE EVALUATED:
' VBA expressions uses the following precedence rules to evaluate mathematical expressions:
'
' 1. ()             Grouping: evaluates functions arguments as well.
' 2. ! - +          Unary operators: exponentiation is the only operation that violates this. Ex.:
'                   -2 ^ 2 = -4 | (-2) ^ 2 = 4.
' 3. ^              Exponentiation: Although Excel and Matlab evaluate nested exponentiations from
'                   left to right, Google, mathematicians and several modern programming languages,
'                   such as Perl, Python and Ruby, evaluate this operation from right to left.
'                   VBA expressions also evals in Python way: a^b^c = a^(b^c).
' 4. * / %          Multiplication, division, modulo: from left to right.
' 5. + -            Addition and subtraction: from left to right.
' 6. < <= <> >= = > Comparison operators.
' 7. ~              Logical negation.
' 8. &              Logical AND.
' 9. ||             Logical XOR.
' 10. |             Logical OR.
'
' Users can enter variables and substitute their values for the calculations. Variable names
' must meet the following requirements:
' 1. Start with a letter.
' 2. End in a letter or number.
'                               "x.1", "number1", "value.a" are valid variable names.
' 3. A variable named "A" is distinct from another variable named "a", since variables are
'    case-sensitive. This rule is broken by the constant PI, since PI=Pi=pi=pI.
' 4. The token "E" cannot be used as variable due this token is reserved for floating point
'    computation. For example, the expression "2.5pi+3.5e" will be evaluated to ~17.3679680,
'    but a expression like "2.5pi+3.5E" will return an error.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
' ABOUT USER-DEFINED FUNCTIONS (UDF):
' Users can register custom modules to expose and use their functions for throght the
' CSVcallBack.cls module. All UDFs must have a single Variant argument that will receive a
' one-dimensional array of strings (one element for each function argument).
'
' VBA expressions can evaluate matrix functions whose arguments are given as vectors, using
' a syntax like [Java](https://www.w3schools.com/java/java_arrays.asp). The following expression
' will calculate the determinant (DET) of a matrix composed of 3 vectors with 3 elements each:
'
' DET({{sin(atn(1)*2); 0; 0}; {0; 2; 0}; {0; 0; 3}})
'
' If a user needs to evaluate an array function that accepts more than one argument, including
' more than one array, all arrays must be passed surrounded by parentheses "({...})".
' For example, a call to a function that emule the SQL IN statement using an array and a reference
' value can be written as follows:
'
'IN_(({{sin(atn(1)*2); 2; 3; 4; 5}});1)
'
' The above will pass this array of strings to the IN_ function:
'
' [{{1;2;3;4;5}}] [1]
'
' However, the array function needs to take care of creating arrays from a string, the ArrayFromString
' method can be used for this purpose.
'
' For illustration pourposes, the UDFunctions.cls module has an implementation of the DET function
' with a usage example of the array handle function. As well, the GCD function is implemented as demo.
'
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' CONSTANTS:
Private PI As Double
Private e As Double
Private Const op_plus As String = "+"
Private Const op_minus As String = "-"
Private Const op_mult As String = "*"
Private Const op_div As String = "/"
Private Const op_intDiv As String = "\"
Private Const op_Factorial As String = "!"
Private Const op_mod As String = "%"
Private Const op_power As String = "^"
Private Const op_equal As String = "="
Private Const op_inequality As String = "<>"
Private Const op_gt As String = ">"
Private Const op_lt As String = "<"
Private Const op_ltequal As String = "<="
Private Const op_gtequal As String = ">="
Private Const op_and As String = "&"
Private Const op_or As String = "|"
Private Const op_xor As String = "||"
Private Const op_neg As String = "~"
Private Const op_AllItems As String = "*+-/^%\=<>&|"
Private Const op_AllNotUnaryItems As String = "*/^%\=<>&|"
Private Const d_lCurly As String = "{"
Private Const d_rCurly As String = "}"
Private Const d_lParenthesis As String = "("
Private Const d_rParenthesis As String = ")"
Private Const d_Apostrophe As String = "'"
'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' VARIABLES:
Private AssignedExpression As Boolean
Private Const BuildinFunctIDList As String = "abs;floor;asin;acos;atn;avg;ceil;cos;exp;gamma;log;lgn;ln;max;min;percent;pow;sgn;sin;sqr;tan"
Private Const BuildinFunctNameList As String = "Absolute;aFloor;ArcSin;ArcCos;ArcTan;Average;aCeiling;Cosin;ExpEuler;Gamma;Logarithm;LgN;LN;Max;Min;Percent;Power;Sign;Sine;SquareRoot;Tangent"
Private ComputedTree As Boolean
Private EvalTree() As ClusterTree
Private ExprToEval As String
Private ExprVariables As ClusterBuffer
Private FunctionsId() As String
Private FunctionsName() As String
Private GeneratedTree As Boolean
Private IsUDFFunction As Boolean
Private LIndexConstruc(0 To 2) As String
Private P_CONSTANTS As ClusterBuffer
Private P_DEGREES As Boolean
Private P_ERRORDESC As String
Private P_ERRTYPE As ExpressionErrors
Private P_EXPRESSION As String
Private P_GALLOPING_MODE As Boolean
Private P_RESULT As Variant
Private P_SEPARATORCHAR As String
Private SubTreeData() As String
Private UserDefFunctions As ClusterBuffer

'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' ENUMERATIONS:
Public Enum OperatorToken
    otNull = 0
    otSum = 1                  '+
    otDiff = 2                 '-
    otMultiplication = 3       '*
    otDivision = 4             '/
    otIntDiv = 5               '\
    otPower = 6                '^
    otMod = 7                  '%
    otEqual = 8                '=
    otNotEqual = 9             '<>
    otGreaterThan = 10         '>
    otLessThan = 11            '<
    otGreaterThanOrEqual = 12  '>=
    otLessThanOrEqual = 13     '<=
    otLogicalAND = 14          '&
    otLogicalOR = 15           '|
    otLogicalXOR = 16          '||
End Enum
Public Enum ExpressionErrors
    errNone = 0
    errUnbalancedBrackets = 1
    errSyntaxError = 2
    errEvalError = 3
    errVariableNotAssigned = 4
End Enum
Private Type Argument
    DefString As String
    FactorialIn As Boolean
    FuncName As String
    FunctionIn As Boolean
    Implicit As Boolean
    LinkedIndex As Long
    LinkedVar As Long
    NegationFlagOn As Boolean
    Operand As String
    UDFFunctionIn As Boolean
End Type
Private Type Token
    Arg1 As Argument
    Arg2 As Argument
    ConstantToken As Boolean
    DefString As String
    EvalResult As String
    Logical As Boolean
    OperationToken As OperatorToken
End Type
Private Type TokenInfo
    OperationToken As OperatorToken
    OperatorLen As Long
    Position As Long
    LogicalToken As Boolean
End Type
Private Type ClusterTree
    Capacity As Long
    ClusterArrBounds() As Long   'Stores the number of rows and columns
    CompArrCluster As Boolean
    CompCluster As Boolean
    EvalResult As String
    Index As Long
    Storage() As Token
End Type
Private Type ClusterItem
    Index As Long
    name As String
    value As String
    Assigned As Boolean
End Type
Private Type ClusterBuffer
    Capacity As Long
    Index As Long
    Storage() As ClusterItem
End Type

'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''' <summary>
''' Initializes this instance.
''' </summary>
Private Sub Class_Initialize()
    AssignedExpression = False
    PI = 4 * Atn(1)
    e = Exp(1)
    P_SEPARATORCHAR = ";"
    P_GALLOPING_MODE = True
    FunctionsId() = Split(BuildinFunctIDList, ";")
    FunctionsName() = Split(BuildinFunctNameList, ";")
    InitCBbuffer P_CONSTANTS
    InitCBbuffer UserDefFunctions
    '@--------------------------------------------------------------------
    ' Save predefined constants
    AppendToCBbuffer P_CONSTANTS, "pi", CStr(PI)
    AppendToCBbuffer P_CONSTANTS, "e", CStr(e)
    '@--------------------------------------------------------------------
    ' Populate linked index constructor
    LIndexConstruc(0) = d_lCurly
    LIndexConstruc(2) = d_rCurly
    '@--------------------------------------------------------------------
    ' Populate building UDFs
    Dim UDFnames() As Variant
    UDFnames() = Array("GCD", "DET")
    DeclareUDF UDFnames, "UserDefFunctions"                 'Declare the Greatest Common Divisor and matrix Determinant functions
                                                            'defined in the UDfunctions class module. This need
                                                            'an instance in the CSVcallBack class module.
    '@--------------------------------------------------------------------
    ' Initialize error messages
    InitializeErrHandler
End Sub

'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' PROPERTIES:


''' <summary>
''' Gets the constants collection. By default this is pre-populated with PI and e.
''' </summary>
Public Property Get Constants() As Collection
    Dim i As Long
    
    Set Constants = New Collection
    For i = 0 To P_CONSTANTS.Index
        Constants.Add P_CONSTANTS.Storage(i).value, P_CONSTANTS.Storage(i).name
    Next i
End Property

''' <summary>
''' Gets a string with the variables values used for the last evaluation.
''' </summary>
Public Property Get CurrentVarValues() As String
    Dim i As Long
    Dim tmpResult As String
    
    For i = 0 To ExprVariables.Index
        If tmpResult = vbNullString Then
            tmpResult = ExprVariables.Storage(i).name & " = " & _
                        ExprVariables.Storage(i).value
        Else
            tmpResult = tmpResult & "; " & ExprVariables.Storage(i).name & " = " & _
                        ExprVariables.Storage(i).value
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
    
    For i = 0 To ExprVariables.Index
        If tmpResult = vbNullString Then
            tmpResult = ExprVariables.Storage(i).name
        Else
            tmpResult = tmpResult & "; " & ExprVariables.Storage(i).name
        End If
    Next i
    CurrentVariables = tmpResult
End Property

''' <summary>
''' Gets or sets the behavior when calculating trigonometric functions. If True, the program
''' will assume argument passed in degrees for all trigonometric functions.
''' </summary>
Public Property Get Degrees() As Boolean
    Degrees = P_DEGREES
End Property

Public Property Let Degrees(aValue As Boolean)
    P_DEGREES = aValue
End Property

''' <summary>
''' Returns the last error message generated.
''' </summary>
Public Property Get ErrorDesc() As String
    ErrorDesc = P_ERRORDESC
End Property

''' <summary>
''' Returns the type of the last error message generated.
''' </summary>
Public Property Get ErrorType() As String
    ErrorType = P_ERRTYPE
End Property

''' <summary>
''' Returns the math expression to be evaluated.
''' </summary>
Public Property Get Expression() As String
    Expression = P_EXPRESSION
End Property

''' <summary>
''' Gets or sets the evaluation in galloping mode. When set
''' to True, the evaluator will discriminate constant tokens
''' from variable tokens.
''' </summary>
Public Property Get GallopingMode() As Boolean
    GallopingMode = P_GALLOPING_MODE
End Property

Public Property Let GallopingMode(aValue As Boolean)
    P_GALLOPING_MODE = aValue
End Property

''' <summary>
''' Gets the parsed status of the actual expression.
''' </summary>
Public Property Get ReadyToEval() As Boolean
    ReadyToEval = GeneratedTree
End Property

''' <summary>
''' Gets the result aftter evaluate the give  expression.
''' </summary>
Public Property Get Result() As String
    Result = P_RESULT
End Property

''' <summary>
''' Gets or sets the character used as functions argumnets, array elements, separator. By
''' default the char used is the semicolon ";".
''' </summary>
Public Property Get SeparatorChar() As String
    SeparatorChar = P_SEPARATORCHAR
End Property

Public Property Let SeparatorChar(aValue As String)
    P_SEPARATORCHAR = aValue
End Property

''' <summary>
''' Gets or sets the current value from/to the given variable.
''' </summary>
Public Property Get VarValue(aVarName As String) As String
    Dim ValueIdx As Long
    
    ValueIdx = GetCBItemIdx(ExprVariables, aVarName)
    If ValueIdx > -1 Then
        VarValue = ExprVariables.Storage(ValueIdx).value
    End If
End Property

Public Property Let VarValue(aVarName As String, aVarValue As String)
    If GeneratedTree Then
        Dim ValueIdx As Long
        
        ValueIdx = GetCBItemIdx(ExprVariables, aVarName)
        If ValueIdx > -1 Then
            ExprVariables.Storage(ValueIdx).value = aVarValue
            ExprVariables.Storage(ValueIdx).Assigned = True
        End If
    End If
End Property

'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' METHODS:
Private Function Absolute(Expression As String) As Double
    Absolute = Abs(CDbl(Expression))
End Function

Private Function aCeiling(ByRef Expression As String) As Double
    aCeiling = Ceiling(CDbl(Expression))
End Function

Public Sub AddConstant(aValue As String, aKey As String)
    Dim ConstIdx As Long
    
    ConstIdx = GetCBItemIdx(P_CONSTANTS, aKey)
    If ConstIdx = -1 Then 'Ensure uniqueness
        AppendToCBbuffer P_CONSTANTS, aKey, aValue
    End If
End Sub

Private Sub AddVariable(ByRef variable As String, ByRef aKey As String)
    Dim VarIdx As Long
    
    CastCase variable, aKey
    VarIdx = GetCBItemIdx(ExprVariables, aKey)
    If VarIdx = -1 Then 'Ensure uniqueness
        AppendToCBbuffer ExprVariables, aKey
        VarIdx = GetCBItemIdx(P_CONSTANTS, aKey)
        If VarIdx > -1 Then 'Assign the value from constants
            ExprVariables.Storage(ExprVariables.Index).value = P_CONSTANTS.Storage(VarIdx).value
            ExprVariables.Storage(ExprVariables.Index).Assigned = True
        End If
    End If
End Sub

Private Function aFloor(ByRef Expression As String) As Double
    aFloor = Floor(CDbl(Expression))
End Function

Private Sub AppendToBuffer(ByRef aBuffer As ClusterTree, ByRef aToken As Token)
    aBuffer.Index = aBuffer.Index + 1
    On Error GoTo BufferAppend_errHandler
    aBuffer.Storage(aBuffer.Index) = aToken
    Exit Sub
BufferAppend_errHandler:
    err.Clear
    ExpandBuffer aBuffer
    aBuffer.Storage(aBuffer.Index) = aToken
End Sub

Private Sub AppendToCBbuffer(ByRef aBuffer As ClusterBuffer, ByRef ItemName As String, Optional ByRef ItemValue As String = vbNullString)
    aBuffer.Index = aBuffer.Index + 1
    On Error GoTo BufferAppend_errHandler
    aBuffer.Storage(aBuffer.Index).Index = aBuffer.Index
    aBuffer.Storage(aBuffer.Index).Assigned = (ItemValue <> vbNullString)
    aBuffer.Storage(aBuffer.Index).name = ItemName
    aBuffer.Storage(aBuffer.Index).value = ItemValue
    Exit Sub
BufferAppend_errHandler:
    err.Clear
    ExpandCBbuffer aBuffer
    aBuffer.Storage(aBuffer.Index).Index = aBuffer.Index
    aBuffer.Storage(aBuffer.Index).Assigned = (ItemValue <> vbNullString)
    aBuffer.Storage(aBuffer.Index).name = ItemName
    aBuffer.Storage(aBuffer.Index).value = ItemValue
End Sub

Private Function ApplyLawOfSigns(ByRef Expression As String) As String
    Dim tmpResult As String
    
    tmpResult = Expression
    Do While tmpResult Like "*[+-][+-]*"
        SignSubstitution tmpResult, "+-", op_minus
        SignSubstitution tmpResult, "-+", op_minus
        SignSubstitution tmpResult, "++", op_plus
        SignSubstitution tmpResult, "--", op_plus
    Loop
    ApplyLawOfSigns = tmpResult
End Function

Private Function ArcCos(ByRef Expression As String) As Double
    Dim tmpEval As Double
        
    tmpEval = CDbl(Expression)
    tmpEval = Atn(-tmpEval / Sqr(-tmpEval * tmpEval + 1)) + 2 * Atn(1)
    If P_DEGREES Then
        tmpEval = tmpEval * 180 / PI
    End If
    ArcCos = tmpEval
End Function

Private Function ArcSin(ByRef Expression As String) As Double
    Dim tmpEval As Double
    
    tmpEval = CDbl(Expression)
    tmpEval = Atn(tmpEval / Sqr(-tmpEval * tmpEval + 1))
    If P_DEGREES Then
        tmpEval = tmpEval * 180 / PI
    End If
    ArcSin = tmpEval
End Function

Private Function ArcTan(ByRef Expression As String) As Double
    Dim tmpEval As Double

    tmpEval = CDbl(Expression)
    tmpEval = Atn(tmpEval)
    If P_DEGREES Then
        tmpEval = tmpEval * 180 / PI
    End If
    ArcTan = tmpEval
End Function

''' <summary>
''' Turns a like Java array string ({{*};{*}}) into a 2D VBA array with n rows and m columns.
''' </summary>
''' <param name="StrArray">The string definition used to create the array.</param>
''' <param name="OutStrArray">The array to return in.</param>
Public Function ArrayFromString(ByRef StrArray As String) As String()
    Dim aIndex As Long
    Dim arrCapacity As Long
    Dim curChar As String
    Dim i As Long, j As Long
    Dim NumCols As Long
    Dim NumRows As Long
    Dim OutStrArray() As String
    Dim StrCopy As String
    Dim StrLen As Long
    Dim tmpStr() As Variant
    Dim VectorEndPos As Long
    Dim VectorOpenFlag As Boolean
    Dim VectorStartPos As Long
    Dim OpenCBrackets As Long
    
    StrCopy = ReconstructLiteralStrings(StrArray, Join$(Split(StrArray, " "), vbNullString))
    StrLen = LenB(StrCopy) - 4
    If MidB$(StrCopy, StrLen + 1, 4) <> "}}" Then 'Missed "}" from input
        Exit Function
    End If
    StrCopy = MidB$(StrCopy, 3, StrLen)
    i = 1
    aIndex = -1
    arrCapacity = 128
    ReDim tmpStr(0 To arrCapacity - 1)
    Do
        curChar = MidB$(StrCopy, i, 2)
        Do While Not curChar = "{"
            i = i + 2
            If i > StrLen Then
                Exit Do
            End If
            curChar = MidB$(StrCopy, i, 2)
        Loop
        VectorOpenFlag = (curChar = "{")
        If VectorOpenFlag Then
            VectorStartPos = i + 2
            OpenCBrackets = OpenCBrackets + 1
        End If
        Do While VectorOpenFlag
            i = i + 2
            curChar = MidB$(StrCopy, i, 2)
            If curChar = "{" Then
                OpenCBrackets = OpenCBrackets + 1
            Else
                If curChar = "}" Then
                    OpenCBrackets = OpenCBrackets - 1
                End If
            End If
            VectorOpenFlag = Not (curChar = "}") Or OpenCBrackets
        Loop
        VectorEndPos = i
        i = i + 2
        aIndex = aIndex + 1
        If aIndex > arrCapacity Then
            arrCapacity = arrCapacity * 2
            ReDim Preserve tmpStr(0 To arrCapacity - 1)
        End If
        tmpStr(aIndex) = Split(MidB$(StrCopy, VectorStartPos, VectorEndPos - VectorStartPos + 1), P_SEPARATORCHAR)
    Loop While i <= StrLen
    NumRows = aIndex
    NumCols = UBound(tmpStr(aIndex))
    ReDim OutStrArray(0 To NumRows, 0 To NumCols)
    For i = 0 To NumRows
        For j = 0 To NumCols
            OutStrArray(i, j) = tmpStr(i)(j)
        Next j
    Next i
    ArrayFromString = OutStrArray
End Function

''' <summary>
''' Turns a mxn array into a like Java array string ({{*};{*}}).
''' </summary>
''' <param name="InputArray">The array to operate.</param>
Public Function ArrayToString(ByRef InputArray As Variant) As String
    Dim i As Long, LB As Long, UB As Long
    Dim j As Long, LB2 As Long, UB2 As Long
    Dim tmpData As String, MultiDimArr As Boolean
    
    If IsArray(InputArray) Then
        LB = LBound(InputArray)
        UB = UBound(InputArray)
        MultiDimArr = Is2Darray(InputArray)
        If MultiDimArr Then
            LB2 = LBound(InputArray, 2)
            UB2 = UBound(InputArray, 2)
        Else
            LB2 = 0
            UB2 = 0
        End If
        For i = LB To UB
            For j = LB2 To UB2
                If j = LB2 Then
                    tmpData = tmpData & "{" & GetArrItm(InputArray, MultiDimArr, i, j)
                Else
                    tmpData = tmpData & P_SEPARATORCHAR & GetArrItm(InputArray, MultiDimArr, i, j)
                End If
                If j = UB2 Then
                    tmpData = tmpData & "}"
                End If
            Next j
            If i < UB Then
                tmpData = tmpData & P_SEPARATORCHAR
            End If
        Next i
        ArrayToString = "{" & tmpData & "}"
    End If
End Function

''' <summary>
''' Turns a two dimensional array into a one dimensional array.
''' </summary>
''' <param name="InputArray">The array to operate.</param>
Private Function Array1DFrom2DArr(ByRef InputArray() As String) As String()
    Dim i As Long, LB As Long, UB As Long
    Dim j As Long, LB2 As Long, UB2 As Long
    Dim tmpData() As String
    
    LB = LBound(InputArray)
    UB = UBound(InputArray)
    LB2 = LBound(InputArray, 2)
    UB2 = UBound(InputArray, 2)
    ReDim tmpData(0 To (UB - LB + 1) * (UB2 - LB2 + 1) - 1)
    For i = LB To UB
        For j = LB2 To UB2
            tmpData(i * (UB2 - LB2 + 1) + j) = InputArray(i, j)
        Next j
    Next i
    Array1DFrom2DArr = tmpData
End Function
Private Function average(ByRef Expression As String) As Double
    Dim g As Long
    Dim tmpData() As String
    Dim tmpEval As Double

    tmpEval = 0
    tmpData() = Split(Expression, P_SEPARATORCHAR)
    For g = LBound(tmpData) To UBound(tmpData)
        tmpEval = tmpEval + CDbl(tmpData(g))
    Next g
    tmpEval = tmpEval / (UBound(tmpData) - LBound(tmpData) + 1)
    average = tmpEval
End Function

Private Sub BottomLevelEval(ByRef aToken As Token)
    If aToken.OperationToken < 8 Then 'Arithmetic operators
        Select Case aToken.OperationToken
            Case OperatorToken.otSum
                aToken.EvalResult = CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn) _
                                    + CastOPtype(aToken.Arg2.Operand, aToken.Arg2.NegationFlagOn)
            Case OperatorToken.otDiff
                aToken.EvalResult = CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn) - _
                                    CastOPtype(aToken.Arg2.Operand, aToken.Arg2.NegationFlagOn)
            Case OperatorToken.otMultiplication
                aToken.EvalResult = CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn) * _
                                    CastOPtype(aToken.Arg2.Operand, aToken.Arg2.NegationFlagOn)
            Case OperatorToken.otDivision
                aToken.EvalResult = CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn) / _
                                    CastOPtype(aToken.Arg2.Operand, aToken.Arg2.NegationFlagOn)
            Case OperatorToken.otPower
                Dim kFctr As Double
                If AscW(aToken.Arg1.DefString) = 45 Then
                    kFctr = -1
                Else
                    kFctr = 1
                End If
                aToken.EvalResult = kFctr * CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn) ^ _
                                            CastOPtype(aToken.Arg2.Operand, aToken.Arg2.NegationFlagOn)
            Case OperatorToken.otMod
                aToken.EvalResult = CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn) Mod _
                                    CastOPtype(aToken.Arg2.Operand, aToken.Arg2.NegationFlagOn)
            Case OperatorToken.otIntDiv
                aToken.EvalResult = Floor(CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn) / _
                                    CastOPtype(aToken.Arg2.Operand, aToken.Arg2.NegationFlagOn))
            Case Else
                If aToken.Logical Then
                    If aToken.Arg1.NegationFlagOn Then
                        If AscW(aToken.Arg1.Operand) <> 126 Then '"~"
                            aToken.EvalResult = Not CBool(aToken.Arg1.Operand)
                        Else
                            aToken.EvalResult = Not CBool(MidB$(aToken.Arg1.Operand, 3))
                        End If
                    Else
                        aToken.EvalResult = CBool(aToken.Arg1.Operand)
                    End If
                Else
                    aToken.EvalResult = aToken.Arg1.Operand
                End If
        End Select
    Else
        If aToken.OperationToken < 14 Then 'Comparison operators
            Select Case aToken.OperationToken
                Case OperatorToken.otEqual
                    aToken.EvalResult = (CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn) = _
                                        CastOPtype(aToken.Arg2.Operand, aToken.Arg2.NegationFlagOn))
                Case OperatorToken.otNotEqual
                    aToken.EvalResult = (CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn) <> _
                                        CastOPtype(aToken.Arg2.Operand, aToken.Arg2.NegationFlagOn))
                Case OperatorToken.otGreaterThan
                    aToken.EvalResult = CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn) > _
                                        CastOPtype(aToken.Arg2.Operand, aToken.Arg2.NegationFlagOn)
                Case OperatorToken.otLessThan
                    aToken.EvalResult = CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn) < _
                                        CastOPtype(aToken.Arg2.Operand, aToken.Arg2.NegationFlagOn)
                Case OperatorToken.otGreaterThanOrEqual
                    aToken.EvalResult = CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn) >= _
                                        CastOPtype(aToken.Arg2.Operand, aToken.Arg2.NegationFlagOn)
                Case Else
                    aToken.EvalResult = CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn) <= _
                                        CastOPtype(aToken.Arg2.Operand, aToken.Arg2.NegationFlagOn)
            End Select
        Else 'Logical operators
            Dim tmpBooleans() As Boolean
            Select Case aToken.OperationToken
                Case OperatorToken.otLogicalAND
                    tmpBooleans() = GetLogicalNeg(aToken)
                    aToken.EvalResult = tmpBooleans(0) And tmpBooleans(1)
                Case OperatorToken.otLogicalOR
                    tmpBooleans() = GetLogicalNeg(aToken)
                    aToken.EvalResult = tmpBooleans(0) Or tmpBooleans(1)
                Case Else
                    tmpBooleans() = GetLogicalNeg(aToken)
                    aToken.EvalResult = tmpBooleans(0) Xor tmpBooleans(1)
            End Select
        End If
    End If
End Sub

Private Sub BuildErrMessage(ErrorType As ExpressionErrors, Optional AditionalContext As String)
    P_ERRTYPE = ErrorType
    Select Case P_ERRTYPE
        Case Is = errUnbalancedBrackets
            P_ERRORDESC = "The expression cannot be evaluated due to unbalanced parenthesis count. " & AditionalContext
        Case Is = errSyntaxError
            P_ERRORDESC = "The expression cannot be evaluated due to a syntax error. " & AditionalContext
        Case Is = errEvalError
            P_ERRORDESC = "An error occurred while trying to evaluate the expression. " & AditionalContext
        Case Is = errVariableNotAssigned
            P_ERRORDESC = "The expression cannot be evaluated due to the existence of variables without assigned values. " & AditionalContext
    End Select
End Sub

Private Function callback(ByRef sClassName As String) As Object
    Set callback = CallByName(New CSVcallBack, sClassName, VbGet)
End Function

Private Sub CastCase(ByRef Expression As String, ByRef outStr As String)
    If Not NotIsPI(Expression) Then
        outStr = LCase$(Expression)     'Case insensitive for PI
    Else
        outStr = Expression
    End If
End Sub

Private Function CastOPtype(ByRef strOperand As String, ByRef Negate As Boolean) As Variant
    If InStrB(1, strOperand, d_Apostrophe) Then 'Literal strings like ['string']
        CastOPtype = strOperand
    Else
        If AscW(strOperand) < 58 Then
            CastOPtype = CDbl(strOperand)
        Else
            If Not Negate Then
                CastOPtype = -1 * CBool(strOperand)
            Else
                CastOPtype = -1 * (Not CBool(strOperand))
            End If
        End If
    End If
End Function

Private Function Ceiling(ByRef value As Double) As Double
    Dim tmpResult As Double
    
    tmpResult = Fix(value)
    Ceiling = tmpResult - ((value <> tmpResult) And (value > 0))
End Function

Private Function CheckArgument(ByRef ArgDefStr As String, ByRef IsFactorial As Boolean) As Boolean
    Dim tmpResult As Boolean
    
    tmpResult = Not OPsymbolInArgument(ArgDefStr, PatternToCheckOn(ArgDefStr))
    If tmpResult Then
        If Not IsFactorial Then     'Check missplaced op_Factorial
            tmpResult = (InStrB(1, ArgDefStr, op_Factorial) = 0)
            If tmpResult Then
                If LenB(ArgDefStr) = 2 Then
                    tmpResult = Not ArgDefStr = op_plus And Not ArgDefStr = op_minus
                End If
            End If
        End If
    End If
    CheckArgument = tmpResult
End Function

Private Function CheckVarValues() As Boolean
    If ExprVariables.Index > -1 And AssignedExpression Then
        Dim i As Long
        Dim cviCounter As Long
        
        For i = 0 To ExprVariables.Index
            If ExprVariables.Storage(i).Assigned Then
                cviCounter = cviCounter + 1
            End If
        Next i
        CheckVarValues = (cviCounter = ExprVariables.Index + 1)
    Else
        CheckVarValues = True
    End If
End Function

''' <summary>
''' Traverses and evaluates the current parse tree.
''' </summary>
Private Function Compute() As String
    Dim B As Long
    Dim t As Long
    Dim i As Long
    Dim OperationIndex As Long
    Dim BaseIndex As Long
    Dim PrevOP1 As String
    Dim PrevOP2 As String
    
    BaseIndex = UBound(SubTreeData) + 1
    For B = LBound(EvalTree) To UBound(EvalTree)              'Loop all subexpressions
        OperationIndex = BaseIndex
        For t = 0 To EvalTree(B).Index                         'Loop all tokens
            OperationIndex = OperationIndex + 1
            If Not EvalTree(B).Storage(t).ConstantToken Then   'Gallop
                If P_GALLOPING_MODE Then
                    PrevOP1 = EvalTree(B).Storage(t).Arg1.Operand
                    PrevOP2 = EvalTree(B).Storage(t).Arg2.Operand
                End If
                GetOperands EvalTree(B).Storage(t), EvalTree(B), BaseIndex
                BottomLevelEval EvalTree(B).Storage(t)
                If P_GALLOPING_MODE Then
                    EvalTree(B).Storage(t).ConstantToken = (PrevOP1 = EvalTree(B).Storage(t).Arg1.Operand And _
                                                            PrevOP2 = EvalTree(B).Storage(t).Arg2.Operand)
                End If
            End If
        Next t
        If Not EvalTree(B).CompCluster Then
            EvalTree(B).EvalResult = EvalTree(B).Storage(t - 1).EvalResult
        Else            'The ClusterTree contains a composite function args as expression
            Dim tmpResult() As String
            ReDim tmpResult(0 To EvalTree(B).Index)
            '@--------------------------------------------------------------------
            ' Loop all sub-expression tokens results
            For i = 0 To EvalTree(B).Index
                tmpResult(i) = EvalTree(B).Storage(i).EvalResult
            Next i
            If Not EvalTree(B).CompArrCluster Then   'Function Argument
                EvalTree(B).EvalResult = Join$(tmpResult, P_SEPARATORCHAR)
            Else        'Array function Argument
                EvalTree(B).EvalResult = JoinArrFunctArg(tmpResult, EvalTree(B).ClusterArrBounds(0), EvalTree(B).ClusterArrBounds(1))
            End If
        End If
    Next B
    Compute = EvalTree(B - 1).EvalResult
    ComputedTree = True
End Function

Private Function Cosin(ByRef Expression As String) As Double
    Dim tmpEval As Double
    
    tmpEval = CDbl(Expression)
    If P_DEGREES Then
        tmpEval = tmpEval * PI / 180
    End If
    Cosin = Cos(tmpEval)
End Function

Private Function CountParentheses(ByRef Expression As String, ByRef ParenthesisChar As String) As Long
    Dim tmpResult As Long
    Dim SearchIndex As Long
    Dim OpenedClosedP As Long
    
    OpenedClosedP = InStrB(1, Expression, ParenthesisChar)
    Do While OpenedClosedP
        tmpResult = tmpResult + 1
        SearchIndex = OpenedClosedP + 2
        OpenedClosedP = InStrB(SearchIndex, Expression, ParenthesisChar)
    Loop
    CountParentheses = tmpResult
End Function

Public Function Create(ByRef aExpression As Variant) As CSVexpressions
    If aExpression <> vbNullString Then
        ExprToEval = RemoveDupNegation(ApplyLawOfSigns(ReconstructLiteralStrings(CStr(aExpression), Join$(Split(aExpression, " "), vbNullString))))
        If ExprToEval <> RemoveDupNegation(ApplyLawOfSigns(ReconstructLiteralStrings(CStr(P_EXPRESSION), Join$(Split(P_EXPRESSION, " "), vbNullString)))) Then
            P_EXPRESSION = aExpression
            VariablesInit ExprToEval
            Parse ExprToEval
        End If
        AssignedExpression = True
    End If
    Set Create = Me
End Function

Public Sub DeclareUDF(ByRef UDFname As Variant, Optional ByRef UDFlib As String = "UserDefFunctions")
    If IsArray(UDFname) Then
        Dim i As Long
        For i = LBound(UDFname) To UBound(UDFname)
            StoreUDF UserDefFunctions, CStr(UDFname(i)), UDFlib    'The Value member points to the function library
        Next i
    Else
        StoreUDF UserDefFunctions, CStr(UDFname), UDFlib
    End If
End Sub

''' <summary>
''' Evaluate a math expression.
''' </summary>
''' <param name="ValuesToEvalIn">The values to use for eval the parse tree.</param>
'@DefaultMember
Public Function Eval(Optional ByRef ValuesToEvalWith As String = vbNullString) As String
Attribute Eval.VB_UserMemId = 0
    If Not AssignedExpression Then
        Exit Function
    Else
        If GeneratedTree Then
            InitializeErrHandler
            If ValuesToEvalWith <> vbNullString Then
                VariableAssignment ValuesToEvalWith
            End If
            If Not ComputedTree Then
                If Not CheckVarValues Then 'Only eval if each variables has its value
                    BuildErrMessage errVariableNotAssigned, "The values of the current variables are: " & CurrentVarValues()
                    Exit Function
                End If
            End If
            On Error GoTo Eval_errHandler
            P_RESULT = Compute()
        End If
    End If
    Eval = P_RESULT
    Exit Function
Eval_errHandler:
    P_RESULT = vbNullString
    Eval = P_RESULT
    BuildErrMessage errEvalError, err.Description & "."
End Function

Private Function EvalFunction(ByRef Argument As String, ByRef FunctionName As String, Optional IsUDF As Boolean = False) As String
    If Not IsUDF Then
        Select Case FunctionName
            Case "Absolute"
                EvalFunction = Absolute(Argument)
            Case "ArcSin"
                EvalFunction = ArcSin(Argument)
            Case "ArcCos"
                EvalFunction = ArcCos(Argument)
            Case "ArcTan"
                EvalFunction = ArcTan(Argument)
            Case "Average"
                EvalFunction = average(Argument)
            Case "Cosin"
                EvalFunction = Cosin(Argument)
            Case "aCeiling"
                EvalFunction = aCeiling(Argument)
            Case "ExpEuler"
                EvalFunction = ExpEuler(Argument)
            Case "aFloor"
                EvalFunction = aFloor(Argument)
            Case "Gamma"
                EvalFunction = tGamma(Argument)
            Case "Logarithm"
                EvalFunction = Logarithm(Argument)
            Case "LgN"
                EvalFunction = LgN(Argument)
            Case "LN"
                EvalFunction = LN(Argument)
            Case "Max"
                EvalFunction = Max(Argument)
            Case "Min"
                EvalFunction = Min(Argument)
            Case "Percent"
                EvalFunction = Percent(Argument)
            Case "Power"
                EvalFunction = Power(Argument)
            Case "Sign"
                EvalFunction = Sign(Argument)
            Case "Sine"
                EvalFunction = Sine(Argument)
            Case "SquareRoot"
                EvalFunction = SquareRoot(Argument)
            Case "Tangent"
                EvalFunction = Tangent(Argument)
            Case Else
                'Rise an error for not found function
        End Select
    Else
        EvalFunction = EvalUDF(FunctionName, Argument)
    End If
End Function

Private Function EvalUDF(ByRef UDFname As String, ByRef Expression As String) As String
    Dim args As Variant
    Dim tmpEval As String
    Dim UDFidx As Long
    
    UDFidx = GetCBItemIdx(UserDefFunctions, UDFname)
    If UDFidx > -1 Then     'Only declared functions are called
        args = SplitArgs(Expression) 'Pass a string array to UDF functions
        tmpEval = CallByName(callback(UserDefFunctions.Storage(UDFidx).value), UDFname, VbMethod, args)
        EvalUDF = tmpEval
    End If
End Function

Private Sub ExpandBuffer(ByRef aBuffer As ClusterTree)
    aBuffer.Capacity = 2 * (aBuffer.Capacity + 1)
    ReDim Preserve aBuffer.Storage(0 To aBuffer.Capacity - 1)
End Sub

Private Sub ExpandCBbuffer(ByRef aBuffer As ClusterBuffer)
    aBuffer.Capacity = 2 * (aBuffer.Capacity + 1)
    ReDim Preserve aBuffer.Storage(0 To aBuffer.Capacity - 1)
End Sub

Private Function ExpEuler(ByRef Expression As String) As Double
    ExpEuler = Exp(CDbl(Expression))
End Function

Private Function Factorial(ByRef Expression As String) As String
    Dim g As Long
    Dim tmpEval As Double
    Dim tmpElm As Double
    Dim Fsgn As Long
    
    If InStrB(1, Expression, P_SEPARATORCHAR) = 0 And InStrB(1, Expression, ".") = 0 Then
        tmpEval = 1
        tmpElm = CDbl(Expression)
        Fsgn = Sgn(tmpElm)
        For g = Abs(tmpElm) To 2 Step -1
            tmpEval = tmpEval * g
        Next g
        Factorial = Fsgn * tmpEval
    Else
        'Code here to rise an error
    End If
End Function

Private Function Floor(ByRef value As Double) As Double
    Dim tmpResult As Double
    
    tmpResult = Fix(value)
    Floor = tmpResult + ((value <> tmpResult) And (value < 0))
End Function

Private Function Gamma(ByRef x As Double) As Double
    'Copyright © 2004, Leonardo Volpi & Foxes Team.
    Dim mantissa As Double, Expo As Double, z As Double
    Dim t As Double, y As Double, e As Long
    If x <= 0 And x - Int(x) = 0 Then 'negative integer
        Gamma = "?": Exit Function
    End If
    z = Abs(x)
    GammaSplit z, mantissa, Expo
    If x < 0 Then
        t = z * Sin(PI * z)
        y = -PI / (mantissa * t)
        e = Int(Log(Abs(y)) / Log(10#))
        mantissa = y * 10 ^ -e
        Expo = e - Expo
    End If
    Gamma = mantissa * 10 ^ Expo
End Function

' gamma  - Lanczos approximation algorithm for gamma function
'Copyright © 2004, Leonardo Volpi & Foxes Team.
Private Sub GammaSplit(ByVal x As Double, ByRef mantissa As Double, ByRef Expo As Double)
Dim z As Double, Cf(14) As Double, w As Double, i As Long, s As Double, p As Double
Const G_ As Double = 4.7421875  '607/128
    z = x - 1
    
    Cf(0) = 0.999999999999997
    Cf(1) = 57.1562356658629
    Cf(2) = -59.5979603554755
    Cf(3) = 14.1360979747417
    Cf(4) = -0.49191381609762
    Cf(5) = 3.39946499848119E-05
    Cf(6) = 4.65236289270486E-05
    Cf(7) = -9.83744753048796E-05
    Cf(8) = 1.58088703224912E-04
    Cf(9) = -2.10264441724105E-04
    Cf(10) = 2.17439618115213E-04
    Cf(11) = -1.64318106536764E-04
    Cf(12) = 8.44182239838528E-05
    Cf(13) = -2.61908384015814E-05
    Cf(14) = 3.68991826595316E-06
    
    w = Exp(G_) / Sqr(PI)
    s = Cf(0)
    For i = 1 To 14
        s = s + Cf(i) / (z + i)
    Next
    s = s / w
    p = Log((z + G_ + 0.5) / Exp(1)) * (z + 0.5) / Log(10)
    'split in mantissa and exponent to avoid overflow
    Expo = Int(p)
    p = p - Int(p)
    mantissa = 10 ^ p * s
    'rescaling
    p = Int(Log(mantissa) / Log(10#))
    mantissa = mantissa * 10 ^ -p
    Expo = Expo + p
End Sub

Private Function GetArithOpInfo(ByRef Expression As String) As TokenInfo
    Dim PowerSymbolPos As Long
    Dim MultSymbolPos As Long
    Dim ModSymbolPos As Long
    Dim DivSymbolPos As Long
    Dim IntDivSymbolPos As Long
    Dim AddSymbolPos As Long
    Dim SubtSymbolPos As Long
    Dim tmpPos As Long
    Dim tmpStr As String
    Dim tmpTestChr As String
    
    PowerSymbolPos = GetPowerSymbolPos(Expression)
    If PowerSymbolPos Then                              'Powers first
        GetArithOpInfo.Position = PowerSymbolPos
        GetArithOpInfo.OperationToken = otPower
        GetArithOpInfo.OperatorLen = LenB(op_power)
    Else
        MultSymbolPos = InStrB(1, Expression, op_mult)
        DivSymbolPos = InStrB(1, Expression, op_div)
        IntDivSymbolPos = InStrB(1, Expression, op_intDiv)
        ModSymbolPos = InStrB(1, Expression, op_mod)
        If NonZero(MultSymbolPos, DivSymbolPos, ModSymbolPos, IntDivSymbolPos) Then
            GetArithOpInfo.Position = MinNonZero(MultSymbolPos, DivSymbolPos, ModSymbolPos, IntDivSymbolPos)    'Priority to the first operator
            Select Case GetArithOpInfo.Position
                Case MultSymbolPos
                    GetArithOpInfo.OperationToken = otMultiplication
                    GetArithOpInfo.OperatorLen = LenB(op_mult)
                Case DivSymbolPos
                    GetArithOpInfo.OperationToken = otDivision
                    GetArithOpInfo.OperatorLen = LenB(op_div)
                Case IntDivSymbolPos
                    GetArithOpInfo.OperationToken = otIntDiv
                    GetArithOpInfo.OperatorLen = LenB(op_intDiv)
                Case Else
                    GetArithOpInfo.OperationToken = otMod
                    GetArithOpInfo.OperatorLen = LenB(op_mod)
            End Select
        Else                                               'Addition and subtraction
            AddSymbolPos = InStrB(1, Expression, op_plus)
            If AddSymbolPos = 1 Then                       'Unary symbol
                AddSymbolPos = InStrB(AddSymbolPos + 2, Expression, op_plus)
            End If
            If AddSymbolPos Then
                tmpPos = AddSymbolPos - 2
                Do While tmpPos >= 1 And AddSymbolPos > 0  'Skip symbol in functions like [{func+arg}]
                    tmpTestChr = MidB$(Expression, tmpPos, 2)
                    If IsLetter(tmpTestChr) Then
                        Do
                            tmpPos = tmpPos - 2
                            If tmpPos = 1 Then
                                Exit Do
                            Else
                                If tmpPos < 1 Then
                                    tmpPos = 1
                                    Exit Do
                                End If
                            End If
                            tmpTestChr = MidB$(Expression, tmpPos, 2)
                        Loop While IsLetter(tmpTestChr)
                        If tmpPos > 1 Then
                            tmpPos = tmpPos + 2
                        End If
                        tmpStr = MidB$(Expression, tmpPos, AddSymbolPos - tmpPos)
                        If GetFunctionName(tmpStr) <> vbNullString Or tmpStr = "E" Then
                            AddSymbolPos = InStrB(AddSymbolPos + 2, Expression, op_plus)
                        Else
                            Exit Do
                        End If
                    Else
                        Exit Do
                    End If
                Loop
            End If
            SubtSymbolPos = InStrB(1, Expression, op_minus)
            If SubtSymbolPos = 1 Then
                SubtSymbolPos = InStrB(SubtSymbolPos + 2, Expression, op_minus)
            End If
            If SubtSymbolPos Then
                tmpPos = SubtSymbolPos - 2
                Do While tmpPos >= 1 And SubtSymbolPos > 0 'Skip symbol in functions like [{func-arg}]
                    tmpTestChr = MidB$(Expression, tmpPos, 2)
                    If IsLetter(tmpTestChr) Then
                        Do
                            tmpPos = tmpPos - 2
                            If tmpPos = 1 Then
                                Exit Do
                            Else
                                If tmpPos < 1 Then
                                    tmpPos = 1
                                    Exit Do
                                End If
                            End If
                            tmpTestChr = MidB$(Expression, tmpPos, 2)
                        Loop While IsLetter(tmpTestChr)
                        If tmpPos > 1 Then
                            tmpPos = tmpPos + 2
                        End If
                        tmpStr = MidB$(Expression, tmpPos, SubtSymbolPos - tmpPos)
                        If GetFunctionName(tmpStr) <> vbNullString Or tmpStr = "E" Then
                            SubtSymbolPos = InStrB(SubtSymbolPos + 2, Expression, op_minus)
                        Else
                            Exit Do
                        End If
                    Else
                        Exit Do
                    End If
                Loop
            End If
            If NonZero(AddSymbolPos, SubtSymbolPos) Then
                GetArithOpInfo.Position = MinNonZero(AddSymbolPos, SubtSymbolPos)     'Priority to the first operator
                Select Case GetArithOpInfo.Position
                    Case AddSymbolPos
                        GetArithOpInfo.OperationToken = otSum
                        GetArithOpInfo.OperatorLen = LenB(op_plus)
                    Case Else
                        GetArithOpInfo.OperationToken = otDiff
                        GetArithOpInfo.OperatorLen = LenB(op_minus)
                End Select
            Else                        'Missing opetaration symbol.
                GetArithOpInfo.Position = -1        'Return a out of bound value
            End If
        End If
    End If
End Function

Private Function GetArrItm(ByRef Arr As Variant, ByRef MultiDimArr As Boolean, _
                            ByRef IdxDim1 As Long, ByRef IdxDim2 As Long) As String
    If MultiDimArr Then
        GetArrItm = CStr(Arr(IdxDim1, IdxDim2))
    Else
        GetArrItm = CStr(Arr(IdxDim1))
    End If
End Function

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
        If cbBuffer.Index > -1 Then
            i = 0
            Do
                tmpResult = (cbBuffer.Storage(i).name = tmpItemName)
                i = i + 1
            Loop While i <= cbBuffer.Index And Not tmpResult
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

Private Function GetRParentPos(ByRef Expression As String, ByRef Index As Long) As Long
    Dim tmpCounter As Long
    Dim SearchIndex As Long
    Dim OpenedSum As Long
    Dim tmpChr As String
    
    SearchIndex = 1
    OpenedSum = 0
    Do
        tmpChr = MidB$(Expression, SearchIndex, 2)
        If tmpChr = d_lParenthesis Then
            OpenedSum = OpenedSum + 1
        Else
            If tmpChr = d_rParenthesis Then
                OpenedSum = OpenedSum - 1
            End If
        End If
        SearchIndex = SearchIndex + 2
    Loop While SearchIndex <= Index
    SearchIndex = 1
    tmpCounter = 0
    Do
        tmpChr = MidB$(Expression, SearchIndex, 2)
        If tmpChr = d_lParenthesis Then
            tmpCounter = tmpCounter + 1
        Else
            If tmpChr = d_rParenthesis Then
                tmpCounter = tmpCounter - 1
            End If
        End If
        SearchIndex = SearchIndex + 2
    Loop Until (tmpCounter = OpenedSum - 1) And (SearchIndex > Index)
    GetRParentPos = SearchIndex - 2
End Function

Private Function GetEvalToken(ByRef Expression As String) As Token
    Dim TokenStart As Long
    Dim TokenEnd As Long
    Dim TokenDet As TokenInfo
    Dim tmpArgs() As String
    
    TokenDet = GetTokenInfo(Expression)
    If TokenDet.Position > 2 Then
        '@--------------------------------------------------------------------
        ' Find token start
        GetTokenStart Expression, TokenDet.Position, TokenStart
        '@--------------------------------------------------------------------
        ' Find token end
        GetTokenEnd Expression, TokenDet.Position, TokenDet.OperatorLen, TokenEnd
        '@--------------------------------------------------------------------
        ' Fill token data
        GetEvalToken.DefString = MidB$(Expression, TokenStart, TokenEnd - TokenStart + 2)
        GetEvalToken.OperationToken = TokenDet.OperationToken
        SplitToken GetEvalToken.DefString, tmpArgs, GetEvalToken.OperationToken
        GetEvalToken.Arg1.DefString = tmpArgs(LBound(tmpArgs))
        GetEvalToken.Arg2.DefString = tmpArgs(UBound(tmpArgs))
        GetEvalToken.Logical = TokenDet.LogicalToken
        GetEvalToken.Arg1.NegationFlagOn = (InStrB(1, GetEvalToken.Arg1.DefString, op_neg) = 1) '~*
        GetEvalToken.Arg2.NegationFlagOn = (InStrB(1, GetEvalToken.Arg2.DefString, op_neg) = 1) '~*
        GetEvalToken.Arg1.FuncName = GetFunctionName(GetEvalToken.Arg1.DefString)
        GetEvalToken.Arg1.UDFFunctionIn = IsUDFFunction
        GetEvalToken.Arg2.FuncName = GetFunctionName(GetEvalToken.Arg2.DefString)
        GetEvalToken.Arg2.UDFFunctionIn = IsUDFFunction
        GetEvalToken.Arg1.FunctionIn = (GetEvalToken.Arg1.FuncName <> vbNullString) Or GetEvalToken.Arg1.UDFFunctionIn
        GetEvalToken.Arg2.FunctionIn = (GetEvalToken.Arg2.FuncName <> vbNullString) Or GetEvalToken.Arg2.UDFFunctionIn
        GetEvalToken.Arg1.LinkedIndex = GetIndex(GetEvalToken.Arg1.DefString)
        GetEvalToken.Arg2.LinkedIndex = GetIndex(GetEvalToken.Arg2.DefString)
        GetEvalToken.Arg1.LinkedVar = GetCBItemIdx(ExprVariables, GetEvalToken.Arg1.DefString)
        GetEvalToken.Arg2.LinkedVar = GetCBItemIdx(ExprVariables, GetEvalToken.Arg2.DefString)
        GetEvalToken.Arg1.Implicit = (GetEvalToken.Arg1.LinkedIndex >= 0)
        GetEvalToken.Arg2.Implicit = (GetEvalToken.Arg2.LinkedIndex >= 0)
        GetEvalToken.Arg1.FactorialIn = (InStrB(1, GetEvalToken.Arg1.DefString, op_Factorial) = LenB(GetEvalToken.Arg1.DefString) - 1)
        GetEvalToken.Arg2.FactorialIn = (InStrB(1, GetEvalToken.Arg2.DefString, op_Factorial) = LenB(GetEvalToken.Arg2.DefString) - 1)
    Else            'Retun values
        GetEvalToken.Arg1.DefString = Expression
        GetEvalToken.Logical = TokenDet.LogicalToken
        GetEvalToken.Arg1.NegationFlagOn = (InStrB(1, GetEvalToken.Arg1.DefString, op_neg) = 1) '~*
        GetEvalToken.DefString = Expression
        GetEvalToken.Arg1.FuncName = GetFunctionName(GetEvalToken.Arg1.DefString)
        GetEvalToken.Arg1.UDFFunctionIn = IsUDFFunction
        GetEvalToken.Arg1.FunctionIn = (GetEvalToken.Arg1.FuncName <> vbNullString) Or GetEvalToken.Arg1.UDFFunctionIn
        GetEvalToken.Arg1.LinkedIndex = GetIndex(GetEvalToken.Arg1.DefString)
        GetEvalToken.Arg1.LinkedVar = GetCBItemIdx(ExprVariables, GetEvalToken.Arg1.DefString)
        GetEvalToken.Arg1.Implicit = (GetEvalToken.Arg1.LinkedIndex >= 0)
        GetEvalToken.Arg1.FactorialIn = (InStrB(1, GetEvalToken.Arg1.DefString, op_Factorial) = LenB(GetEvalToken.Arg1.DefString) - 1)
    End If
End Function

Private Function GetFunctionName(ByRef Expression As String) As String
    Dim EFjCounter As Long
    Dim GFNbool As Boolean
    Dim ExpCopy As String
    Dim tmpPos As Long
    
    ExpCopy = LCase$(Expression)
    For EFjCounter = LBound(FunctionsId) To UBound(FunctionsId)
        tmpPos = InStrB(1, ExpCopy, FunctionsId(EFjCounter))
        If tmpPos Then
            GFNbool = ValidFuntionName(ExpCopy, FunctionsId(EFjCounter), tmpPos)
            If GFNbool Then
                Exit For
            End If
        End If
    Next EFjCounter
    If GFNbool Then
        GetFunctionName = FunctionsName(EFjCounter)
        IsUDFFunction = False
    Else 'Check for UDFs
        Dim i As Long
        For i = 0 To UserDefFunctions.Index
            tmpPos = InStrB(1, ExpCopy, UserDefFunctions.Storage(i).name)
            If tmpPos Then
                GFNbool = ValidFuntionName(ExpCopy, UserDefFunctions.Storage(i).name, tmpPos)
                If GFNbool Then
                    Exit For
                End If
            End If
        Next i
        If Not GFNbool Then
            GetFunctionName = vbNullString
        Else
            GetFunctionName = UserDefFunctions.Storage(i).name
        End If
        IsUDFFunction = GFNbool
    End If
End Function

Private Function GetIndex(ByRef SubstStr As String) As Long
    Dim InitPos As Long
    Dim EndPos As Long
    
    InitPos = InStrB(1, SubstStr, d_lCurly)
    If InitPos Then
        EndPos = InStrB(1, SubstStr, d_rCurly)
        If EndPos Then
            GetIndex = MidB$(SubstStr, InitPos + 2, EndPos - InitPos - 2)
        Else
            GetIndex = -1
        End If
    Else
        GetIndex = -1
    End If
End Function

Private Function GetLCOpInfo(ByRef Expression As String) As TokenInfo
    Dim NotEqualSymbolPos As Long
    Dim GreatterOrEqualSymbolPos As Long
    Dim LessOrEqualSymbolPos As Long
    Dim EqualSymbolPos As Long
    Dim LessThanSymbolPos As Long
    Dim GreatterThanSymbolPos As Long
    Dim LogANDSymbolPos As Long
    Dim LogORSymbolPos As Long
    Dim LogXORSymbolPos As Long
    Dim testChar As String
    
    '@--------------------------------------------------------------------
    ' Comparisons first
    Do 'Discard false positives
        EqualSymbolPos = InStrB(EqualSymbolPos + 1, Expression, op_equal)
        If EqualSymbolPos > 2 Then
            testChar = MidB$(Expression, EqualSymbolPos - 2, 2)
        Else
            testChar = vbNullString
        End If
    Loop While EqualSymbolPos > 0 And testChar Like "[<>]"
    NotEqualSymbolPos = InStrB(1, Expression, op_inequality)
    Do
        GreatterThanSymbolPos = InStrB(GreatterThanSymbolPos + 1, Expression, op_gt)
        If GreatterThanSymbolPos Then
            testChar = MidB$(Expression, GreatterThanSymbolPos + 2, 2)
        Else
            testChar = vbNullString
        End If
    Loop While GreatterThanSymbolPos > 0 And testChar = op_equal
    Do
        LessThanSymbolPos = InStrB(LessThanSymbolPos + 1, Expression, op_lt)
        If LessThanSymbolPos Then
            testChar = MidB$(Expression, LessThanSymbolPos + 2, 2)
        Else
            testChar = vbNullString
        End If
    Loop While LessThanSymbolPos > 0 And testChar = op_equal
    GreatterOrEqualSymbolPos = InStrB(1, Expression, op_gtequal)
    LessOrEqualSymbolPos = InStrB(1, Expression, op_ltequal)
    If NonZero(EqualSymbolPos, NotEqualSymbolPos, GreatterThanSymbolPos, _
                LessThanSymbolPos, GreatterOrEqualSymbolPos, LessOrEqualSymbolPos) Then
        GetLCOpInfo.Position = MinNonZero(EqualSymbolPos, NotEqualSymbolPos, GreatterThanSymbolPos, _
                                    LessThanSymbolPos, GreatterOrEqualSymbolPos, LessOrEqualSymbolPos)    'Priority to the first operator
        Select Case GetLCOpInfo.Position
            Case EqualSymbolPos
                GetLCOpInfo.OperationToken = otEqual
                GetLCOpInfo.OperatorLen = LenB(op_equal)
            Case NotEqualSymbolPos
                GetLCOpInfo.OperationToken = otNotEqual
                GetLCOpInfo.OperatorLen = LenB(op_inequality)
            Case GreatterThanSymbolPos
                GetLCOpInfo.OperationToken = otGreaterThan
                GetLCOpInfo.OperatorLen = LenB(op_gt)
            Case LessThanSymbolPos
                GetLCOpInfo.OperationToken = otLessThan
                GetLCOpInfo.OperatorLen = LenB(op_lt)
            Case GreatterOrEqualSymbolPos
                GetLCOpInfo.OperationToken = otGreaterThanOrEqual
                GetLCOpInfo.OperatorLen = LenB(op_gtequal)
            Case Else
                GetLCOpInfo.OperationToken = otLessThanOrEqual
                GetLCOpInfo.OperatorLen = LenB(op_ltequal)
            End Select
        '@--------------------------------------------------------------------
        ' LogicalOperators
        Else
            '@--------------------------------------------------------------------
            ' Check the expression giving higher precedence to AND, XOR
            LogANDSymbolPos = InStrB(LogANDSymbolPos + 1, Expression, op_and)
            If NonZero(LogANDSymbolPos) Then
                GetLCOpInfo.Position = LogANDSymbolPos
                GetLCOpInfo.OperationToken = otLogicalAND
                GetLCOpInfo.OperatorLen = LenB(op_and)
            Else
                LogXORSymbolPos = InStrB(LogXORSymbolPos + 1, Expression, op_xor)
                If NonZero(LogXORSymbolPos) Then
                    GetLCOpInfo.Position = LogXORSymbolPos
                    GetLCOpInfo.OperationToken = otLogicalXOR
                    GetLCOpInfo.OperatorLen = LenB(op_xor)
                Else
                    Do 'Discard false positives: "||"
                        LogORSymbolPos = InStrB(LogORSymbolPos + 1, Expression, op_or)
                        If LogORSymbolPos Then
                            testChar = MidB$(Expression, LogORSymbolPos + 2, 2)
                        Else
                            testChar = vbNullString
                        End If
                    Loop While LogORSymbolPos > 0 And testChar = op_or
                    If NonZero(LogORSymbolPos) Then
                        GetLCOpInfo.Position = LogORSymbolPos
                        GetLCOpInfo.OperationToken = otLogicalOR
                        GetLCOpInfo.OperatorLen = LenB(op_or)
                    Else                        'Missing opetaration symbol.
                        GetLCOpInfo.Position = -1        'Return a out of bound value
                    End If
                End If
            End If
        End If
End Function

Private Function GetLogicalNeg(ByRef aToken As Token) As Boolean()
    Dim tmpBoolArr(0 To 1) As Boolean
    
    If aToken.Arg1.NegationFlagOn Then
        tmpBoolArr(0) = Not CBool(MidB$(aToken.Arg1.Operand, 3))
    Else
        tmpBoolArr(0) = CBool(aToken.Arg1.Operand)
    End If
    If aToken.Arg2.NegationFlagOn Then
        tmpBoolArr(1) = Not CBool(MidB$(aToken.Arg2.Operand, 3))
    Else
        tmpBoolArr(1) = CBool(aToken.Arg2.Operand)
    End If
    GetLogicalNeg = tmpBoolArr
End Function

Private Function GetLParentPos(ByRef Expression As String, ByRef RelativePosition As Long) As Long
    Dim tmpCounter As Long
    Dim SearchIndex As Long
    Dim OpenedP As Long
    Dim tmpResult As Long
    
    OpenedP = InStrB(1, Expression, d_lParenthesis)
    Do While tmpCounter < RelativePosition
        tmpCounter = tmpCounter + 1
        SearchIndex = OpenedP + 2
        tmpResult = OpenedP
        OpenedP = InStrB(SearchIndex, Expression, d_lParenthesis)
    Loop
    GetLParentPos = tmpResult
End Function

Private Sub GetOperand(ByRef CurToken As Token, ByRef CurArg As Argument, _
                        ByRef CurTree As ClusterTree, ByRef BaseIndex As Long)
                                    
    If CurArg.Implicit Then
        If CurArg.FunctionIn Then 'Implicit function
            If CurArg.FactorialIn Then
                If CurArg.LinkedIndex >= BaseIndex Then 'Data on current tree
                    CurArg.Operand = Factorial(EvalFunction(CurTree.Storage(CurArg.LinkedIndex - BaseIndex).EvalResult, _
                                        CurArg.FuncName, CurArg.UDFFunctionIn))
                Else 'Data is main tree
                    CurArg.Operand = Factorial(EvalFunction(EvalTree(CurArg.LinkedIndex).EvalResult, _
                                        CurArg.FuncName, CurArg.UDFFunctionIn))
                End If
            Else 'Eval
                If CurArg.LinkedIndex >= BaseIndex Then
                    CurArg.Operand = EvalFunction(CurTree.Storage(CurArg.LinkedIndex - BaseIndex).EvalResult, _
                                        CurArg.FuncName, CurArg.UDFFunctionIn)
                Else
                    CurArg.Operand = EvalFunction(EvalTree(CurArg.LinkedIndex).EvalResult, _
                                        CurArg.FuncName, CurArg.UDFFunctionIn)
                End If
            End If
        Else    'Return data
            If CurArg.LinkedIndex >= BaseIndex Then
                If CurArg.FactorialIn Then
                    CurArg.Operand = Factorial(CurTree.Storage(CurArg.LinkedIndex - BaseIndex).EvalResult)
                Else
                    CurArg.Operand = CurTree.Storage(CurArg.LinkedIndex - BaseIndex).EvalResult
                End If
            Else
                If CurArg.FactorialIn Then
                    CurArg.Operand = Factorial(EvalTree(CurArg.LinkedIndex).EvalResult)
                Else
                    CurArg.Operand = EvalTree(CurArg.LinkedIndex).EvalResult
                End If
            End If
        End If
        If AscW(CurArg.DefString) = 45 Then
            CurArg.Operand = ApplyLawOfSigns(op_minus + CurArg.Operand)
        End If
    Else     'Explicit function or data
        If CurArg.LinkedVar > -1 Then     'Variable substitution
            If CurArg.FactorialIn Then            'Operate factorials
                CurArg.Operand = Factorial(ExprVariables.Storage(CurArg.LinkedVar).value)
            Else
                CurArg.Operand = ExprVariables.Storage(CurArg.LinkedVar).value
            End If
            If AscW(CurArg.DefString) = 45 Then
                CurArg.Operand = ApplyLawOfSigns(op_minus + CurArg.Operand)
            End If
        Else
            If CurArg.FactorialIn Then
                CurArg.Operand = Factorial(MidB$(CurArg.DefString, 1, LenB(CurArg.DefString) - 2))
            Else
                CurArg.Operand = CurArg.DefString
            End If
        End If
    End If
End Sub

Private Sub GetOperands(ByRef CurToken As Token, ByRef CurTree As ClusterTree, _
                                    ByRef BaseIndex As Long)
                                    
    '@--------------------------------------------------------------------
    ' Get the first operand
    GetOperand CurToken, CurToken.Arg1, CurTree, BaseIndex
    '@--------------------------------------------------------------------
    ' Get the second operand if required
    If Not CurToken.OperationToken = otNull Then
        GetOperand CurToken, CurToken.Arg2, CurTree, BaseIndex
    End If
End Sub

''' <summary>
''' Gets the operator given an operator token.
''' </summary>
Private Function GetOpSymbol(ByRef OPtoken As OperatorToken) As String
    Select Case OPtoken
        Case OperatorToken.otSum
            GetOpSymbol = op_plus
        Case OperatorToken.otDiff
            GetOpSymbol = op_minus
        Case OperatorToken.otMultiplication
            GetOpSymbol = op_mult
        Case OperatorToken.otDivision
            GetOpSymbol = op_div
        Case OperatorToken.otIntDiv
            GetOpSymbol = op_intDiv
        Case OperatorToken.otMod
            GetOpSymbol = op_mod
        Case OperatorToken.otPower
            GetOpSymbol = op_power
        Case OperatorToken.otEqual
            GetOpSymbol = op_equal
        Case OperatorToken.otNotEqual
            GetOpSymbol = op_inequality
        Case OperatorToken.otGreaterThan
            GetOpSymbol = op_gt
        Case OperatorToken.otLessThan
            GetOpSymbol = op_lt
        Case OperatorToken.otGreaterThanOrEqual
            GetOpSymbol = op_gtequal
        Case OperatorToken.otLessThanOrEqual
            GetOpSymbol = op_ltequal
        Case OperatorToken.otLogicalAND
            GetOpSymbol = op_and
        Case OperatorToken.otLogicalOR
            GetOpSymbol = op_or
        Case OperatorToken.otLogicalXOR
            GetOpSymbol = op_xor
    End Select
End Function

Private Function GetPowerSymbolPos(ByRef Expression As String) As Long
    Dim tmpPos As Long
    Dim tmpResult As Long
    
    tmpPos = InStrB(1, Expression, op_power)
    Do While tmpPos
        tmpResult = tmpPos
        tmpPos = InStrB(tmpResult + 2, Expression, op_power)
    Loop
    GetPowerSymbolPos = tmpResult
End Function

Private Sub GetRootedTree(ByRef SubExpression As String, ByRef tmpReplacement As String, _
                                    ByRef OperationIndex As Long, ByRef outBuffer As ClusterTree)
    Dim vToken As Token
    Dim switch As Boolean
    Dim tmpPos As Long
    Dim OperandInBundle As Boolean
    Dim PrevChar As String
    
    Do
        SubExpression = ApplyLawOfSigns(SubExpression)
        vToken = GetEvalToken(SubExpression)
        '@--------------------------------------------------------------------
        ' Mask worked token
        tmpPos = InStrB(1, SubExpression, vToken.DefString)
        If tmpPos > 2 Then
            PrevChar = MidB$(SubExpression, tmpPos - 2, 2)
            OperandInBundle = (InStrB(1, op_AllItems, PrevChar))
            Do While Not OperandInBundle And tmpPos > 2    'Tokens starts with a operator or with a null string
                tmpPos = InStrB(tmpPos + 2, SubExpression, vToken.DefString)
                PrevChar = MidB$(SubExpression, tmpPos - 2, 2)
                OperandInBundle = (InStrB(1, op_AllItems, PrevChar))
            Loop
        End If
        SubExpression = MidB$(SubExpression, 1, tmpPos - 1) & tmpReplacement & MidB$(SubExpression, tmpPos + LenB(vToken.DefString))
        AppendToBuffer outBuffer, vToken 'Save to target token ClusterTree
        switch = (SubExpression <> tmpReplacement)
        If switch Then
            OperationIndex = OperationIndex + 1
            tmpReplacement = GetSubstStr(OperationIndex)
        End If
    Loop While switch
End Sub

Private Function GetSubstStr(ByRef aValue As Long) As String
    If aValue >= 0 Then
        LIndexConstruc(1) = aValue
        GetSubstStr = Join$(LIndexConstruc, vbNullString)
    End If
End Function

''' <summary>
''' Returns an array with all the sub expressions needed to
''' evaluate the given expression. A string such as {0}
''' indicates that the current token should be evaluated using
''' the value or token residing at index 0 using functions
''' and arithmetic operators.
''' </summary>
''' <param name="Expression">The expression to parse.</param>
Private Function GetSubTreeData(ByRef Expression As String) As String()
    Dim opCount As Long
    Dim CPcount As Long
    
    opCount = CountParentheses(Expression, d_lParenthesis)
    CPcount = CountParentheses(Expression, d_rParenthesis)
    If opCount <> CPcount Then
        Exit Function
    End If
    Dim PSerial() As String
    Dim FSerial() As String
    If opCount > 0 Then
        ReDim PSerial(0 To opCount - 1)
        Dim tgOpenP As Long
        Dim tgClosedP As Long
        Dim bMark As Long
        Dim eMark As Long
        Dim tmpTK As String
        Dim Uspace As Long
        Dim tmpKey As String
        
        tgOpenP = opCount
        tgClosedP = 1
        Uspace = LBound(PSerial)
        Do
            bMark = GetLParentPos(Expression, tgOpenP) + 2
            eMark = GetRParentPos(Expression, bMark - 2)
            tmpTK = MidB$(Expression, bMark, eMark - bMark)
            If UnicToken(PSerial, tmpTK) Then
                PSerial(Uspace) = tmpTK
                Uspace = Uspace + 1
            End If
            tgOpenP = tgOpenP - 1
        Loop While tgOpenP > 0
        Dim CurIdx As Long
        Dim fpIdx As Long
        ReDim Preserve PSerial(LBound(PSerial) To Uspace - 1)
        FSerial() = PSerial
        For CurIdx = UBound(PSerial) To LBound(PSerial) Step -1
            For fpIdx = LBound(PSerial) To UBound(PSerial)
                If fpIdx <> CurIdx Then
                    If InStrB(1, FSerial(fpIdx), PSerial(CurIdx)) Then
                        tmpKey = GetSubstStr(CurIdx)
                        FSerial(fpIdx) = Replace(FSerial(fpIdx), _
                                                    d_lParenthesis & PSerial(CurIdx) & d_rParenthesis, _
                                                    tmpKey)
                    End If
                End If
            Next fpIdx
        Next CurIdx
    Else
        ReDim FSerial(0 To 0)
        FSerial(0) = Expression
    End If
    GetSubTreeData = FSerial
End Function

Private Sub GetTokenEnd(ByRef Expression As String, ByRef startIndex As Long, ByRef OPlen As Long, ByRef outLng As Long)
    Dim curChar As String
    Dim lenExpr As Long
    Dim tmpPos As Long
    Dim varLen As Long
    Dim tmpVar As String
    Dim functionFlag As Boolean
    Dim cFlag As Boolean
    Dim OperandInBundle As Boolean
    
    lenExpr = LenB(Expression)
    outLng = startIndex + OPlen
    curChar = MidB$(Expression, outLng, 2)
    Do
        tmpPos = outLng
        If IsLetter(curChar) Then 'Check for functions [ arg/-funct(...) ] and Sci notation
            Do
                varLen = varLen + 2
                outLng = outLng + 2
                curChar = MidB$(Expression, outLng, 2)
            Loop While IsAlphaNumeric(curChar) And tmpPos < lenExpr
            tmpVar = MidB$(Expression, tmpPos, varLen)
            functionFlag = (GetFunctionName(tmpVar) <> vbNullString)
            If functionFlag Then   'Ignore [+-] symbols for functions
                outLng = outLng + 2
                curChar = MidB$(Expression, outLng, 2)
                If IsPlusOrMinus(curChar) Then
                    outLng = outLng + 2
                    curChar = MidB$(Expression, outLng, 2)
                End If
            Else
                If LenB(tmpVar) = 2 Then  'Ignore [+-] symbols for sci not symbol "E"
                    If AscW(tmpVar) = 69 Then
                        If IsPlusOrMinus(curChar) Then
                            outLng = outLng + 2
                            curChar = MidB$(Expression, outLng, 2)
                        End If
                    End If
                End If
            End If
        Else 'Check operations like a[*/\^]-b
            If InStrB(1, op_AllNotUnaryItems, curChar) Then 'May be unary symbol
                outLng = outLng + 2
            End If
        End If
        cFlag = (tmpPos <> outLng)
        If Not cFlag Then           'Conditional check of the next symbol
            outLng = outLng + 2
            curChar = MidB$(Expression, outLng, 2)
        End If
        OperandInBundle = (InStrB(1, op_AllItems, curChar))
    Loop While Not OperandInBundle And outLng < lenExpr
    If outLng < lenExpr Then
        outLng = outLng - 2
    Else
        outLng = lenExpr
    End If
End Sub

Private Function GetTokenInfo(ByRef Expression As String) As TokenInfo
    Dim tmpResult As TokenInfo
            
    tmpResult = GetArithOpInfo(Expression)
    If tmpResult.Position = -1 Then 'Missing arithmetic opetarators.
        If Expression Like "*[=<>&|]*" Then 'Try with logical operators.
            tmpResult = GetLCOpInfo(Expression)
            tmpResult.LogicalToken = True
        Else
            tmpResult.Position = -1
            If Expression Like "*[Tt][Rr][Uu][Ee]" Then
                tmpResult.LogicalToken = True
            Else
                If Expression Like "*[Ff][Aa][Ll][Ss][Ee]" Then
                    tmpResult.LogicalToken = True
                End If
            End If
        End If
    End If
    GetTokenInfo = tmpResult
End Function

Private Sub GetTokenStart(ByRef Expression As String, ByRef startIndex As Long, ByRef outLng As Long)
    Dim curChar As String
    
    '@--------------------------------------------------------------------
    ' Find token start
    outLng = startIndex - 2
    curChar = MidB$(Expression, outLng, 2)
    Do While (InStrB(1, op_AllItems, curChar) = 0) And outLng > 1
        outLng = outLng - 2
        curChar = MidB$(Expression, outLng, 2)
        If AscW(curChar) = 46 Then  'Dot "."
            outLng = outLng - 2
            If outLng > 0 Then
                curChar = MidB$(Expression, outLng, 2)
            End If
        Else
            If outLng >= 3 Then
                If AscW(MidB$(Expression, outLng - 2, 2)) = 69 Then 'Ignore Sci notation "E"
                    outLng = outLng - 4
                    curChar = MidB$(Expression, outLng, 2)
                End If
            End If
        End If
    Loop
    '@--------------------------------------------------------------------
    ' Check for symbols combination: [*+-/^%\~=<>&|][-~]
    If outLng > 1 Then
        Select Case AscW(curChar)
            Case 45, 126
                Dim PrevChar As String
                PrevChar = MidB$(Expression, outLng - 2, 2)
                If InStrB(1, op_AllNotUnaryItems, PrevChar) Then
                    outLng = outLng - 2
                Else
                    outLng = outLng + 2
                End If
            Case Else
                outLng = outLng + 2
        End Select
    Else
        If outLng < 1 Then
            outLng = 1
        End If
    End If
End Sub

Private Function GoBackToOpSymbol(ByRef Expression As String, ByRef VarStartPos As Long) As Long
    Dim tmpResult As Long
    
    tmpResult = VarStartPos - 2
    Do While tmpResult >= 1
        If ImplicitMultFlag(MidB$(Expression, tmpResult, 2)) Then
            tmpResult = tmpResult - 2
        Else
            Exit Do
        End If
    Loop
    GoBackToOpSymbol = tmpResult
End Function

Private Function ImplicitMultFlag(ByRef Char As String) As Boolean
    If LenB(Char) Then
        Select Case AscW(Char)
            Case 46, 48 To 57
                ImplicitMultFlag = True
            Case Else
                ImplicitMultFlag = False
        End Select
    End If
End Function

Private Sub InitBuffer(ByRef aBuffer As ClusterTree)
    aBuffer.Capacity = 128
    ReDim aBuffer.Storage(0 To aBuffer.Capacity - 1)
    aBuffer.Index = -1
End Sub

Private Sub InitCBbuffer(ByRef aBuffer As ClusterBuffer)
    aBuffer.Capacity = 128
    ReDim aBuffer.Storage(0 To aBuffer.Capacity - 1)
    aBuffer.Index = -1
End Sub

Private Sub InitializeErrHandler()
    P_ERRORDESC = vbNullString
    P_ERRTYPE = errNone
End Sub

Private Function InitializedArray(ByRef Arr As Variant) As Boolean
    Dim UB As Long
    
    On Error GoTo err_handler
    UB = UBound(Arr)
    InitializedArray = True
    Exit Function
err_handler:
    InitializedArray = False
End Function

Private Function Is2Darray(Arr As Variant) As Boolean
    Dim d As Long
    
    On Error GoTo err_handler
    d = UBound(Arr, 2)
    Is2Darray = True
    Exit Function
err_handler:
    Is2Darray = False
End Function

Private Function IsAlphaNumeric(ByRef Char As String) As Boolean
    If LenB(Char) Then
        Select Case AscW(Char)
            Case 48 To 57, 65 To 90, 97 To 122
                IsAlphaNumeric = True
            Case Else
                IsAlphaNumeric = False
        End Select
    End If
End Function

Private Function IsBoolean(ByRef Expression As String) As Boolean
    IsBoolean = (Expression = "true")
    If Not IsBoolean Then
        IsBoolean = (Expression = "false")
    End If
End Function

Private Function IsDigit(ByRef Char As String) As Boolean
    If LenB(Char) Then
        Select Case AscW(Char)
            Case 48 To 57
                IsDigit = True
            Case Else
                IsDigit = False
        End Select
    End If
End Function

Private Function IsExtAlphaNumeric(ByRef Char As String) As Boolean
    If LenB(Char) Then
        Select Case AscW(Char)
            Case 46, 48 To 57, 65 To 90, 97 To 122
                IsExtAlphaNumeric = True
            Case Else
                IsExtAlphaNumeric = False
        End Select
    End If
End Function

Private Function IsLetter(ByRef Char As String) As Boolean
    If LenB(Char) Then
        Select Case AscW(Char)
            Case 65 To 90, 97 To 122
                IsLetter = True
            Case Else
                IsLetter = False
        End Select
    End If
End Function

Private Function IsLikeSciNot(ByRef Chars As String) As Boolean
    If LenB(Chars) Then
        If AscW(Chars) = 69 Then 'Sci not token "E"
            If IsDigit(MidB$(Chars, 3, 2)) Then 'Check 2nd char
                IsLikeSciNot = True
            End If
        Else
            IsLikeSciNot = False
        End If
    End If
End Function

Private Function IsLiteralString(ByRef aString As String) As Boolean
    If LenB(aString) Then
        If AscW(aString) = 39 Then 'Apostrophe
            IsLiteralString = (InStrB(3, aString, d_Apostrophe) = LenB(aString) - 1)
        Else
            IsLiteralString = False
        End If
    End If
End Function

Private Function IsPlusOrMinus(ByRef Char As String) As Boolean
    If LenB(Char) Then
        Select Case AscW(Char)
            Case 43, 45
                IsPlusOrMinus = True
            Case Else
                IsPlusOrMinus = False
        End Select
    End If
End Function

Private Function JoinArrFunctArg(ByRef DecompArray() As String, ByRef MaxRowIndex As Long, ByRef MaxColIndex As Long) As String
    Dim tmpResult As String
    Dim i As Long, j As Long
    
    For i = 0 To MaxRowIndex
        For j = 0 To MaxColIndex
            If j = 0 Then
                tmpResult = tmpResult & d_lCurly & DecompArray(j + (i * (MaxColIndex + 1)))
            Else
                tmpResult = tmpResult & P_SEPARATORCHAR & DecompArray(j + (i * (MaxColIndex + 1)))
            End If
            If j = MaxColIndex Then
                tmpResult = tmpResult & d_rCurly
            End If
        Next j
        If i < MaxRowIndex Then
            tmpResult = tmpResult & P_SEPARATORCHAR
        End If
    Next i
    JoinArrFunctArg = d_lCurly & tmpResult & d_rCurly
End Function

Private Function LgN(ByRef Expression As String) As Double
    Dim tmpData() As String
    Dim tmpEval As Double
    
    tmpEval = 0
    tmpData() = Split(Expression, P_SEPARATORCHAR)
    tmpEval = Log(CDbl(tmpData(LBound(tmpData)))) _
                / Log(CDbl(tmpData(UBound(tmpData))))    'Log(x)/Log(N)
    LgN = tmpEval
End Function

Private Function LN(ByRef Expression As String) As Double
    LN = Log(CDbl(Expression))
End Function

Private Function Logarithm(ByRef Expression As String) As Double
    Logarithm = Log(CDbl(Expression)) / Log(10)
End Function

Private Function Max(ByRef Expression As String) As Double
    Dim g As Long
    Dim tmpData() As String
    Dim tmpEval As Double
    Dim compEval As Double

    tmpData() = Split(Expression, P_SEPARATORCHAR)
    tmpEval = CDbl(tmpData(LBound(tmpData)))
    For g = LBound(tmpData) + 1 To UBound(tmpData)
        compEval = CDbl(tmpData(g))
        If compEval > tmpEval Then
            tmpEval = compEval
        End If
    Next g
    Max = tmpEval
End Function

Private Function Min(ByRef Expression As String) As Double
    Dim g As Long
    Dim tmpData() As String
    Dim tmpEval As Double
    Dim compEval As Double
    
    tmpData() = Split(Expression, P_SEPARATORCHAR)
    tmpEval = CDbl(tmpData(LBound(tmpData)))
    For g = LBound(tmpData) + 1 To UBound(tmpData)
        compEval = CDbl(tmpData(g))
        If compEval < tmpEval Then
            tmpEval = compEval
        End If
    Next g
    Min = tmpEval
End Function

''' <summary>
''' Takes a list of values of type Long and
''' returns the minimum non-zero value.
''' </summary>
''' <param name="Values">The list of values.</param>
Private Function MinNonZero(ParamArray values() As Variant) As Long
    Dim minTmp As Long
    Dim vIndex As Long
    
    vIndex = LBound(values)
    minTmp = values(vIndex)
    For vIndex = vIndex + 1 To UBound(values)
        If values(vIndex) > 0 Then
            If minTmp Then
                If values(vIndex) < minTmp Then
                    minTmp = values(vIndex)
                End If
            Else
                minTmp = values(vIndex)
            End If
        End If
    Next vIndex
    MinNonZero = minTmp
End Function

''' <summary>
''' Takes a list of values of type Long and
''' returns true if one of then is a non-zero value.
''' </summary>
''' <param name="Values">The list of values.</param>
Private Function NonZero(ParamArray values() As Variant) As Boolean
    Dim tmpResult As Boolean
    Dim vIndex As Long
    Dim valuesUB As Long
    
    vIndex = LBound(values)
    valuesUB = UBound(values)
    Do
        tmpResult = (values(vIndex) > 0)
        vIndex = vIndex + 1
    Loop While Not tmpResult And vIndex <= valuesUB
    NonZero = tmpResult
End Function

Private Function NotIsPI(ByRef VarName As String) As Boolean

    NotIsPI = (LCase$(VarName) <> "pi")
End Function

Private Function OPsymbolInArgument(ByRef ArgDefStr As String, ByRef Pattrn As String) As Boolean
    Dim i As Long
    Dim tmpResult As Boolean
    Dim lenStr As Long
    
    i = 1
    lenStr = LenB(ArgDefStr)
    Do
        tmpResult = InStrB(1, MidB$(ArgDefStr, i, 2), Pattrn)
        i = i + 2
    Loop While i <= lenStr And Not tmpResult
    OPsymbolInArgument = tmpResult
End Function

''' <summary>
''' Parses an expression and builds a parse tree. The
''' computed data is ready for traverse and evaluate.
''' </summary>
''' <param name="Expression">The expression to parse.</param>
Private Sub Parse(ByRef Expression As String)
    Dim lambdaText As String
    Dim meLB As Long
    Dim meUB As Long
    Dim meCounter As Long
    Dim SimpleExpr As Boolean
    Dim TreeUB As Long
    Dim LbrCount As Long
    Dim RbrCount As Long
    
    On Error GoTo Parse_errHandler
    InitializeErrHandler
    LbrCount = CountParentheses(Expression, d_lParenthesis)
    RbrCount = CountParentheses(Expression, d_rParenthesis)
    If LbrCount <> RbrCount Then
        If LbrCount > RbrCount Then
            BuildErrMessage errUnbalancedBrackets, "At least one right parenthesis [)] is missing."
        Else
            BuildErrMessage errUnbalancedBrackets, "At least one left parenthesis [(] is missing."
        End If
        Exit Sub
    End If
    err.Clear
    SubTreeData() = GetSubTreeData(Expression)
    lambdaText = SerializeSubTree(Expression, SubTreeData)
    meLB = LBound(SubTreeData)
    meUB = UBound(SubTreeData)
    SimpleExpr = (meUB - meLB = 0 And lambdaText = SubTreeData(meUB))
    TreeUB = meUB + Abs(CLng(Not SimpleExpr)) 'Upper bound for ClusterTree array
    
    ReDim EvalTree(meLB To TreeUB)
    For meCounter = meLB To TreeUB
        InitBuffer EvalTree(meCounter)   'Reserve storage for tree branches/Expressions
        If meCounter < TreeUB Then          'Tokenize branches
            TokenizeSubExpr SubTreeData(meCounter), SubTreeData, EvalTree(meCounter)
        Else
            If Not SimpleExpr Then          'Tokenize main tree/lambda expression
                TokenizeSubExpr lambdaText, SubTreeData, EvalTree(meCounter)
            Else
                TokenizeSubExpr SubTreeData(meCounter), SubTreeData, EvalTree(meCounter)
            End If
        End If
        ShrinkBuffer EvalTree(meCounter)
    Next meCounter
    If ValidTree Then    'Call the validate function
        GeneratedTree = True
    Else
        GeneratedTree = False
        If P_ERRTYPE = errNone Then
            BuildErrMessage errSyntaxError
        End If
    End If
    Exit Sub
Parse_errHandler:
    BuildErrMessage errSyntaxError, err.Description & "."
End Sub

''' <summary>
''' Parses variables in an expression and resolves implicit multiplications.
''' </summary>
''' <param name="Expression">The expression to parse.</param>
Private Sub ParseVariables(ByRef Expression As String)
    Dim curPos As Long
    Dim lenExpr As Long
    Dim varInitPos As Long
    Dim varLen As Long
    Dim tmpChr As String
    Dim tmpVar As String
    Dim tmpPos As Long
    Dim ConstantKey As String
    Dim charsBefore As String
    
    curPos = 1
    lenExpr = LenB(Expression)
    Do
        tmpChr = MidB$(Expression, curPos, 2)
        varLen = 0
        If IsLetter(tmpChr) Or tmpChr = d_Apostrophe Then
            varInitPos = curPos
            If tmpChr <> d_Apostrophe Then
                Do
                    varLen = varLen + 2
                    curPos = curPos + 2
                    tmpChr = MidB$(Expression, curPos, 2)
                Loop While IsExtAlphaNumeric(tmpChr) And curPos < lenExpr
            Else
                curPos = InStrB(varInitPos + 2, Expression, d_Apostrophe)
                varLen = curPos - varInitPos + 2
            End If
            If varInitPos > 2 Then
                tmpPos = varInitPos - 2
            Else
                tmpPos = varInitPos
            End If
            If Not IsDigit(MidB$(Expression, tmpPos, 2)) Then  'Variables start with a letter
                tmpVar = MidB$(Expression, varInitPos, varLen)
                If Not ReservedToken(tmpVar) Then
                    If GetFunctionName(LCase$(tmpVar)) = vbNullString Then
                        If Not IsLiteralString(tmpVar) Then 'Exclude literal strings
                            AddVariable tmpVar, ConstantKey
                        End If
                    End If
                End If
            Else
                tmpVar = MidB$(Expression, varInitPos, varLen)
                If Not ReservedToken(tmpVar) Then
                    tmpPos = GoBackToOpSymbol(Expression, varInitPos) + 1
                    charsBefore = MidB$(Expression, tmpPos + 1, varInitPos - (tmpPos + 1)) 'Chars before variable
                    If IsNumeric(charsBefore) Then 'implied multiplication found
                        If GetFunctionName(LCase$(tmpVar)) = vbNullString Then
                            If Not IsLiteralString(tmpVar) Then
                                AddVariable tmpVar, ConstantKey
                            End If
                        End If
                        Expression = MidB$(Expression, 1, varInitPos - LenB(charsBefore) - 1) & _
                                        charsBefore & op_mult & tmpVar & _
                                        MidB$(Expression, varInitPos + LenB(tmpVar))
                        lenExpr = LenB(Expression)
                    End If
                End If
            End If
        End If
        curPos = curPos + 2
    Loop While curPos < lenExpr
End Sub

Private Function PatternToCheckOn(ByRef ArgDefStr As String) As String
    Select Case AscW(ArgDefStr)
        Case 43, 45
            PatternToCheckOn = op_AllNotUnaryItems
        Case Else
            PatternToCheckOn = op_AllItems
    End Select
End Function

Private Function Percent(ByRef Expression As String) As Double
    Percent = CDbl(Expression) / 100
End Function

Private Function Power(ByRef Expression As String) As Double
    Dim tmpData() As String
    Dim tmpEval As Double
    
    tmpData() = Split(Expression, P_SEPARATORCHAR)
    tmpEval = CDbl(tmpData(LBound(tmpData))) ^ _
                    CDbl(tmpData(LBound(tmpData) + 1))
    Power = tmpEval
End Function

Private Function ReconstructLiteralStrings(InputExpr As String, StoredExp As String) As String
    Dim curPosInput As Long
    Dim lastPosInput As Long
    Dim curPosStored As Long
    Dim lastPosStored As Long
    Dim closingMarkInput As Long
    Dim closingMarkStored As Long
    Dim tmpResult As String
    
    lastPosInput = 1
    lastPosStored = -1
    curPosInput = InStrB(lastPosInput, InputExpr, d_Apostrophe)
    tmpResult = StoredExp
    Do While curPosInput        'Found literal string
        lastPosStored = lastPosStored + 2
        curPosStored = InStrB(lastPosStored, tmpResult, d_Apostrophe)
        closingMarkInput = InStrB(curPosInput + 2, InputExpr, d_Apostrophe)
        closingMarkStored = InStrB(curPosStored + 2, tmpResult, d_Apostrophe)
        tmpResult = MidB$(tmpResult, 1, curPosStored - 1) _
                    & MidB$(InputExpr, curPosInput, closingMarkInput - curPosInput + 2) _
                    & MidB$(tmpResult, closingMarkStored + 2)
        lastPosStored = closingMarkStored + 2
        lastPosInput = closingMarkInput + 2
        curPosInput = InStrB(lastPosInput, InputExpr, d_Apostrophe)
    Loop
    ReconstructLiteralStrings = tmpResult
End Function

Private Function RemoveDupNegation(ByRef Expression As String) As String
    Dim tmpResult As String
    
    tmpResult = Expression
    Do While tmpResult Like "*~~*"
        tmpResult = Replace(tmpResult, "~~", vbNullString)
    Loop
    RemoveDupNegation = tmpResult
End Function

Private Function ReservedToken(ByRef VarName As String) As Boolean
    ReservedToken = True
    If Not IsBoolean(LCase$(VarName)) Then 'Exclude logical values
        If IsLikeSciNot(VarName) Then ' Like "E#*"
            If Not IsDigit(MidB$(VarName, LenB(VarName) - 1)) Then 'Exclude Sci notation exp
                If Not AscW(VarName) = 69 Then  'Exclude Sci notation token E
                    ReservedToken = False
                End If
            End If
        Else
            If Not AscW(VarName) = 69 Then  'Exclude Sci notation token E
                ReservedToken = False
            End If
        End If
    End If
End Function

''' <summary>
''' Returns an implicit string containing the order in which arithmetic operations
''' should be evaluated. The string can point to any part of the source expression
''' using syntax such as {1}*{3}+{0} to represent it.
''' </summary>
''' <param name="Expression">The target expression of the lambda operation.</param>
''' <param name="SubExpressionsData">A set of expressions to derive the lambda string.</param>
Private Function SerializeSubTree(ByRef Expression As String, ByRef SubExpressionsData() As String) As String
    Dim glUb As Long
    Dim glCounter As Long
    Dim tmpResult As String
    Dim tmpKey As String
    
    tmpResult = Expression
    glUb = UBound(SubExpressionsData)
    For glCounter = LBound(SubExpressionsData) To glUb
        If InStrB(1, tmpResult, SubExpressionsData(glCounter)) Then
            tmpKey = GetSubstStr(glCounter)
            tmpResult = Replace(tmpResult, d_lParenthesis & SubExpressionsData(glCounter) & d_rParenthesis, tmpKey)
        End If
    Next
    SerializeSubTree = tmpResult
End Function

Private Sub ShrinkBuffer(ByRef aBuffer As ClusterTree)
    If aBuffer.Index > -1 Then
        aBuffer.Capacity = aBuffer.Index
        ReDim Preserve aBuffer.Storage(0 To aBuffer.Capacity)
    End If
End Sub

Private Function Sign(ByRef Expression As String) As Double
    Sign = (Sgn(CDbl(Expression)))
End Function

Private Sub SignSubstitution(ByRef Expression As String, FindStr As String, Replacement As String)
    If InStrB(1, Expression, FindStr) Then
        Expression = Replace(Expression, FindStr, Replacement)
    End If
End Sub

Private Function Sine(ByRef Expression As String) As Double
    Dim tmpEval As Double
    
    tmpEval = CDbl(Expression)
    If P_DEGREES Then
        tmpEval = tmpEval * PI / 180
    End If
    Sine = Sin(tmpEval)
End Function

''' <summary>
''' Splits an array function argument and returns an array with the bounds
''' of the passed string array. In the returned array, the lower index
''' indicates the argument max row index and the upper index indicates the
''' argument max column index.
''' </summary>
''' <param name="Argument">The array function argumnet ("*{{*}}*").</param>
''' <param name="outArr">Array to be scanned and overwritten.</param>
Private Function SplitArrBranch(ByRef Argument As String, ByRef outArr() As String) As Long()
    Dim tmpArr() As String
    Dim tmpResult(0 To 1) As Long
    
    tmpArr() = ArrayFromString(Argument)
    If InitializedArray(tmpArr) Then 'Transform success
        outArr = Array1DFrom2DArr(tmpArr)
        tmpResult(0) = UBound(tmpArr)   'Rows in the array
        tmpResult(1) = UBound(tmpArr, 2)  'Columns in each row
    Else
        tmpResult(0) = -1   'Return error values
        tmpResult(1) = -1
    End If
    SplitArrBranch = tmpResult
End Function

Private Function SplitArgs(ByRef args As String) As String()
    Dim tmpPos As Long
    Dim ArrStartIdx As Long
    Dim ArrEndIdx As Long
    Dim SargStart As Long
    Dim SargEnd As Long
    Dim tmpResult As String
    Dim LenArgsStr As Long
    
    tmpPos = 1
    LenArgsStr = LenB(args)
    Do
        ArrStartIdx = InStrB(tmpPos, args, "{{")
        If ArrStartIdx = tmpPos Then 'Currrent argument is an array
            ArrEndIdx = InStrB(ArrStartIdx, args, "}}")
            tmpPos = ArrEndIdx + 6
            If tmpResult <> vbNullString Then
                tmpResult = tmpResult & "?" & MidB$(args, ArrStartIdx, ArrEndIdx - ArrStartIdx + 4)
            Else
                tmpResult = MidB$(args, ArrStartIdx, ArrEndIdx - ArrStartIdx + 4)
            End If
        Else
            SargStart = tmpPos
            SargEnd = InStrB(tmpPos, args, P_SEPARATORCHAR)
            If SargEnd = 0 Then
                SargEnd = LenArgsStr + 1
            End If
            tmpPos = SargEnd + 2
            If tmpResult <> vbNullString Then
                tmpResult = tmpResult & "?" & MidB$(args, SargStart, SargEnd - SargStart)
            Else
                tmpResult = MidB$(args, SargStart, SargEnd - SargStart)
            End If
        End If
    Loop While tmpPos < LenArgsStr
    SplitArgs = Split(tmpResult, "?")
End Function

Private Sub SplitToken(ByRef Expression As String, ByRef oArray() As String, ByRef OPtoken As OperatorToken)
    Dim opSymbol As String
    Dim tmpPos As Long
    Dim Position As Long
    ReDim oArray(0 To 1)
    
    opSymbol = GetOpSymbol(OPtoken)
    If OPtoken = OperatorToken.otSum Or OPtoken = OperatorToken.otDiff Then
        tmpPos = InStrB(1, Expression, opSymbol)
        Do While tmpPos
            Position = tmpPos
            tmpPos = InStrB(tmpPos + 2, Expression, opSymbol)
        Loop
        If tmpPos = 1 Then          'unary expression
            oArray(0) = "0"
        Else
            oArray(0) = MidB$(Expression, 1, Position - 1)
        End If
        oArray(1) = MidB$(Expression, Position + 2)
    Else
        oArray() = Split(Expression, opSymbol)
    End If
End Sub

Private Function SquareRoot(ByRef Expression As String) As Double
    SquareRoot = Sqr(CDbl(Expression))
End Function

Private Function Tangent(ByRef Expression As String) As Double
    Dim tmpEval As Double

    tmpEval = CDbl(Expression)
    If P_DEGREES Then
        tmpEval = tmpEval * PI / 180
    End If
    Tangent = Tan(tmpEval)
End Function

Private Sub StoreUDF(ByRef targetBuffer As ClusterBuffer, _
                    ByRef UDFname As String, ByRef UDFlib As String)
    Dim UDFidx As Long
    Dim tmpUDF As String
    Dim tmpUDFlib As String
    
    tmpUDF = LCase$(UDFname)
    UDFidx = GetCBItemIdx(targetBuffer, tmpUDF)
    If UDFidx = -1 Then 'Ensure uniqueness
        tmpUDFlib = LCase$(UDFlib)
        AppendToCBbuffer targetBuffer, tmpUDF, tmpUDFlib
    End If
End Sub

Private Function tGamma(ByRef Expression As String) As Double
    tGamma = Gamma(CDbl(Expression))
End Function

Private Sub TokenizeSubExpr(ByRef Expression As String, ByRef SubExpressionsData() As String, ByRef outBuffer As ClusterTree)
    Dim tmpReplacement As String
    Dim ExpCopy As String
    Dim tmpArgs() As String
    Dim taIcounter As Long
    Dim OperationIndex As Long
    Dim tmpIndex As Long
    
    tmpIndex = UBound(SubExpressionsData) + 1
    OperationIndex = tmpIndex
    ExpCopy = Expression
    tmpReplacement = GetSubstStr(OperationIndex)
    If Not ExpCopy Like "*{{*}}*" Then
        Select Case InStrB(1, ExpCopy, P_SEPARATORCHAR)
            Case 0 'Regular sub-expression
                GetRootedTree ExpCopy, tmpReplacement, OperationIndex, outBuffer
                outBuffer.CompCluster = False
            Case Else 'Composite function argument
                tmpArgs() = Split(ExpCopy, P_SEPARATORCHAR)
                For taIcounter = LBound(tmpArgs) To UBound(tmpArgs)
                    GetRootedTree tmpArgs(taIcounter), tmpReplacement, OperationIndex, outBuffer
                Next taIcounter
                outBuffer.CompCluster = True
                outBuffer.CompArrCluster = False
        End Select
    Else    'Composite array function argument
        outBuffer.ClusterArrBounds = SplitArrBranch(ExpCopy, tmpArgs)
        If outBuffer.ClusterArrBounds(0) <> -1 Then 'Splitting argument success
            For taIcounter = LBound(tmpArgs) To UBound(tmpArgs)
                GetRootedTree tmpArgs(taIcounter), tmpReplacement, OperationIndex, outBuffer
            Next taIcounter
            outBuffer.CompCluster = True
            outBuffer.CompArrCluster = True
        Else
            'Todo: Code here for trap error of missing () in a composite array and standard input
        End If
    End If
End Sub

Private Function UnicToken(ByRef Source() As String, ByRef value As String) As Boolean
    Dim UTiCounter As Long
    Dim UTjCounter As Long
    Dim tmpResult As Boolean
    
    UTiCounter = LBound(Source)
    UTjCounter = UBound(Source)
    Do
        tmpResult = (Source(UTiCounter) = value)
        UTiCounter = UTiCounter + 1
    Loop While Not tmpResult And UTiCounter < UTjCounter
    UnicToken = Not tmpResult
End Function

Private Function ValidFuntionName(ByRef Expression As String, _
                                    ByRef FuntionName As String, _
                                    StartPos As Long) As Boolean
    Dim cLeft As Boolean
    Dim cRight As Boolean
    
    If StartPos > 1 Then
        cLeft = Not IsLetter(MidB$(Expression, StartPos - 2, 2))
    Else
        cLeft = True
    End If
    cRight = Not IsLetter(MidB$(Expression, StartPos + LenB(FuntionName), 2))
    ValidFuntionName = cLeft And cRight
End Function

Private Function ValidTree() As Boolean
    Dim B As Long
    Dim c As Long
    Dim t As Long
    Dim tmpResult As Boolean
    
    B = LBound(EvalTree)
    c = UBound(EvalTree)
    Do                        'Loop all subexpressions
        t = 0
        Do                    'Loop all tokens
            tmpResult = CheckArgument(EvalTree(B).Storage(t).Arg1.DefString, _
                        EvalTree(B).Storage(t).Arg1.FactorialIn)
            If tmpResult Then
                If EvalTree(B).Storage(t).Arg2.DefString <> vbNullString Then
                    tmpResult = CheckArgument(EvalTree(B).Storage(t).Arg2.DefString, _
                                EvalTree(B).Storage(t).Arg2.FactorialIn)
                Else
                    tmpResult = (EvalTree(B).Storage(t).OperationToken = otNull) 'Check for invalid input like "**" "//" ...
                End If
            End If
            t = t + 1
        Loop While t <= EvalTree(B).Index And tmpResult
        B = B + 1
    Loop While B <= c And tmpResult
    If Not tmpResult Then
        BuildErrMessage errSyntaxError, "Misplaced operators have been found. Please check and try again."
    End If
    ValidTree = tmpResult
End Function

''' <summary>
''' Assigns values to the variables defined in the given expression.
''' An expression such as "x = 5; y = -20.5" will assign 5 to the x
''' variable and -20.5 to the y variable. Used by the eval method.
''' </summary>
''' <param name="vString">The list of variables to assign.</param>
Private Sub VariableAssignment(ByRef vString As String)
    Dim tmpAssignment() As String
    Dim tmpValues() As String
    Dim avIcounter As Long
    Dim VarIdx As Long
    
    tmpAssignment() = Split(ReconstructLiteralStrings(vString, Join$(Split(vString, " "), vbNullString)), P_SEPARATORCHAR)
    For avIcounter = LBound(tmpAssignment) To UBound(tmpAssignment)
        tmpValues() = Split(tmpAssignment(avIcounter), "=")
        If tmpValues(UBound(tmpValues)) <> vbNullString Then
            If IsNumeric(tmpValues(UBound(tmpValues))) Or IsLiteralString(tmpValues(UBound(tmpValues))) Then
                VarIdx = GetCBItemIdx(ExprVariables, tmpValues(LBound(tmpValues)))
                If VarIdx > -1 Then
                    ExprVariables.Storage(VarIdx).value = tmpValues(UBound(tmpValues))
                    ExprVariables.Storage(VarIdx).Assigned = True
                End If
            End If
        End If
    Next avIcounter
End Sub

Private Sub VariablesInit(ByRef Expression As String)
    Dim i As Long
    Dim tmpIdx As Long
    
    InitCBbuffer ExprVariables
    ParseVariables Expression
    For i = 0 To ExprVariables.Index
        tmpIdx = GetCBItemIdx(P_CONSTANTS, ExprVariables.Storage(i).name)
        If tmpIdx > -1 Then 'The variable is a defined constant
            ExprVariables.Storage(i).value = P_CONSTANTS.Storage(tmpIdx).value
            ExprVariables.Storage(i).Assigned = True
        End If
    Next i
End Sub
