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
' Copyright © 2022-2023 W. García
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
' Digit         =     "0-9"
' Operator      =     "+" | "-" | "*" | "/" | "\" | "^" | "%" | "!" | "<" | "<=" | "<>" | ">" | ">=" | "=" | "$" | "&" | "|" | "||"
' Function      =     "abs" | "sin" | "cos" | "min" |...|[UDF]
'
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
' ABOUT THE ORDER IN WHICH MATHEMATICAL EXPRESSIONS ARE EVALUATED:
' VBA expressions uses the following precedence rules to evaluate mathematical expressions:
'
' 1. ()               Grouping: evaluates functions arguments as well.
' 2. ! - +            Unary operators: exponentiation is the only operation that violates this. Ex.:
'                     -2 ^ 2 = -4 | (-2) ^ 2 = 4.
' 3. ^                Exponentiation: Although Excel and Matlab evaluate nested exponentiations from
'                     left to right, Google, mathematicians and several modern programming languages,
'                     such as Perl, Python and Ruby, evaluate this operation from right to left.
'                     VBA expressions also evals in Python way: a^b^c = a^(b^c).
' 4. * / %            Multiplication, division, modulo: from left to right.
' 5. + -              Addition and subtraction: from left to right.
' 6. < <= <> >= = > $ Comparison operators.
' 7. ~                Logical negation.
' 8. &                Logical AND.
' 9. ||               Logical XOR.
' 10. |               Logical OR.
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
' For illustration pourposes, the CSVudFunctions.cls module has an implementation of the DET function
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
Private Const op_like As String = "$"
Private Const op_neg As String = "~"
Private Const op_AllItems As String = "*+-/^%\=<>&|$"
Private Const op_AllNotUnaryItems As String = "*/^%\=<>&|$"
Private Const d_lCurly As String = "{"
Private Const d_rCurly As String = "}"
Private Const d_lParenthesis As String = "("
Private Const d_rParenthesis As String = ")"
Private Const d_Apostrophe As String = "'"
Private Const d_Space As String = " "
Private Const e_ValueError As String = "#VALUE!"
Private Const TINY As Double = 1E-20
'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' VARIABLES:
Private AssignedExpression As Boolean
Private BuildinFunctIDList As String
Private BuildinFunctNameList As String
Private ComputedTree As Boolean
Private EvalTree() As ClusterTree
Private ExprToEval As String
Private FunctionsId() As String
Private FunctionsName() As String
Private GeneratedTree As Boolean
Private IsUDFFunction As Boolean
Private LIndexConstruc(0 To 2) As String
Private P_DEGREES As Boolean
Private P_ERRORDESC As String
Private P_ERRTYPE As ExpressionErrors
Private P_EXPRESSION As String
Private P_FORCE_BOOLEAN As Boolean
Private P_FORMATRESULT As Boolean
Private P_GALLOPING_MODE As Boolean
Private P_RESULT As Variant
Private P_SCOPE As CSVexpressionsScope
Private P_SEPARATORCHAR As String
Private SubTreeData() As String
Private UserDefFunctions As ClusterBuffer

'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' ENUMERATIONS:
Public Enum OperatorToken
    otNull = 0
    otSum = 1                    '+
    otDiff = 2                   '-
    otMultiplication = 3         '*
    otDivision = 4               '/
    otIntDiv = 5                 '\
    otPower = 6                  '^
    otMod = 7                    '%
    otEqual = 101                '=
    otNotEqual = 102             '<>
    otGreaterThan = 103          '>
    otLessThan = 104             '<
    otGreaterThanOrEqual = 105   '>=
    otLessThanOrEqual = 106      '<=
    otLike = 107                 '$
    otLogicalAND = 201           '&
    otLogicalOR = 202            '|
    otLogicalXOR = 203           '||
End Enum
Public Enum ExpressionErrors
    errNone = 0
    errUnbalancedBrackets = 1
    errSyntaxError = 2
    errEvalError = 3
    errVariableNotAssigned = 4
End Enum
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' TYPES:
Private Type Argument
    DefString As String
    FactorialIn As Boolean
    funcName As String
    FunctionIn As Boolean
    Implicit As Boolean
    LinkedIndex As Long
    LinkedVar As Long
    NegationFlagOn As Boolean
    Operand As String
    UDFFunctionIn As Boolean
End Type
Private Type token
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
    index As Long
    resulstMap As String
    Storage() As token
End Type
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
    P_GALLOPING_MODE = False
    P_FORMATRESULT = False
    BuildinFunctIDList = "abs;floor;asin;acos;asc;atn;array;avg;ceil;chr;cos;choose;date;dateadd;datediff" & _
                            ";datepart;dateserial;datevalue;day;ddb;det;exp;format;fv;fzero;gamma;hour;iff;inverse;ipmt" & _
                            ";irr;lcase;left;len;log;lgn;ln;ludecomp;lusolve;max;mid;min;minute;mirr;mmult;month;monthname;now" & _
                            ";nper;npv;percent;pmt;ppmt;pow;pv;rate;replace;right;sgn;sin;sln;solve;sqr;switch" & _
                            ";syd;tan;timeserial;timevalue;trim;ucase;weekday;weekdayname;year"
    BuildinFunctNameList = "Absolute;aFloor;ArcSin;ArcCos;ASCII;ArcTan;strArray;Average;aCeiling;ASCIIchr;Cosin" & _
                            ";aChoose;aDate;aDateAdd;aDateDiff;aDatePart;aDateSerial;aDateValue;aDay;aDDB;MatrixDeterminant" & _
                            ";ExpEuler;aFormat;aFV;FunctionZero;Gamma;aHour;aIff;InverseMatrix;aIPMT;aIRR;LowerCase;aLeft;aLen;Logarithm" & _
                            ";LgN;LN;LUdecomposition;LUSolveLinearSystem;Max;Middle;Min;aMinute;aMIRR;MatrixMult;aMonth;aMonthName;aNow;aNPER;aNPV;Percent;aPMT;aPPMT;Power;aPV" & _
                            ";aRATE;aReplace;aRight;Sign;Sine;aSLN;SolveLinearSystem;SquareRoot;aSwitch;aSYD;Tangent;aTimeSerial;aTimeValue" & _
                            ";aTrim;aUcase;aWweekDay;aWeekDayName;aYear"
    FunctionsId() = Split(BuildinFunctIDList, ";")
    FunctionsName() = Split(BuildinFunctNameList, ";")
    Set P_SCOPE = New CSVexpressionsScope
    P_FORCE_BOOLEAN = False
    InitCBbuffer UserDefFunctions
    '@--------------------------------------------------------------------
    ' Populate linked index constructor
    LIndexConstruc(0) = d_lCurly
    LIndexConstruc(2) = d_rCurly
    '@--------------------------------------------------------------------
    ' Populate building UDFs
    Dim UDFnames() As Variant
    UDFnames() = Array("GCD", "Concat")
    DeclareUDF UDFnames, "UserDefFunctions"                 ' Declare the Greatest Common Divisor and the string concatenation
                                                            ' functions defined in the CSVudFunctions class module.
                                                            ' This need an instance in the CSVcallBack class module.
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
Public Property Get constants() As Collection
    Set constants = P_SCOPE.constants
End Property

''' <summary>
''' Gets a string with the variables values used for the last evaluation.
''' </summary>
Public Property Get currentVarValues() As String
    currentVarValues = P_SCOPE.currentVarValues
End Property

''' <summary>
''' Gets a string with the variables values used for the last evaluation.
''' </summary>
Public Property Get currentVariables() As String
    currentVariables = P_SCOPE.currentVariables
End Property

''' <summary>
''' Gets or sets the behavior when calculating trigonometric functions. If True, the program
''' will assume argument passed in degrees for all trigonometric functions.
''' </summary>
Public Property Get degrees() As Boolean
    degrees = P_DEGREES
End Property

Public Property Let degrees(aValue As Boolean)
    P_DEGREES = aValue
End Property

''' <summary>
''' Returns the last error message generated.
''' </summary>
Public Property Get errorDesc() As String
    errorDesc = P_ERRORDESC
End Property

''' <summary>
''' Returns the type of the last error message generated.
''' </summary>
Public Property Get errorType() As ExpressionErrors
    errorType = P_ERRTYPE
End Property

''' <summary>
''' Sets/gets the scope for the current expression.
''' </summary>
Public Property Get evalScope() As CSVexpressionsScope
    Set evalScope = P_SCOPE
End Property

Public Property Set evalScope(aObject As CSVexpressionsScope)
    Set P_SCOPE = aObject
End Property

''' <summary>
''' Returns the math expression to be evaluated.
''' </summary>
Public Property Get expression() As String
    expression = P_EXPRESSION
End Property

''' <summary>
''' When True, a False is returned on evaluation errors.
''' This can be useful when evaluating piecewise functions.
''' </summary>
Public Property Let forceBoolean(aValue As Boolean)
    P_FORCE_BOOLEAN = aValue
End Property

Public Property Get forceBoolean() As Boolean
    forceBoolean = P_FORCE_BOOLEAN
End Property

''' <summary>
''' Indicates if the results will be converted to standard VBA strings
''' </summary>
Public Property Get formatResult() As Boolean
    formatResult = P_FORMATRESULT
End Property

Public Property Let formatResult(aValue As Boolean)
    P_FORMATRESULT = aValue
End Property

''' <summary>
''' Gets or sets the evaluation in galloping mode. When set
''' to True, the evaluator will discriminate constant tokens
''' from variable tokens.
''' </summary>
Public Property Get gallopingMode() As Boolean
    gallopingMode = P_GALLOPING_MODE
End Property

Public Property Let gallopingMode(aValue As Boolean)
    P_GALLOPING_MODE = aValue
End Property

''' <summary>
''' Gets the parsed status of the actual expression.
''' </summary>
Public Property Get readyToEval() As Boolean
    readyToEval = GeneratedTree
End Property

''' <summary>
''' Gets the result aftter evaluate the give  expression.
''' </summary>
Public Property Get result() As String
    If P_FORMATRESULT Then
        result = FormatLiteralString(CStr(P_RESULT), True)
    Else
        result = P_RESULT
    End If
End Property

''' <summary>
''' Gets or sets the character used as functions argumnets, array elements, separator. By
''' default the char used is the semicolon ";".
''' </summary>
Public Property Get separatorChar() As String
    separatorChar = P_SEPARATORCHAR
End Property

Public Property Let separatorChar(aValue As String)
    P_SEPARATORCHAR = aValue
End Property

''' <summary>
''' Gets or sets the current value from/to the given variable.
''' </summary>
Public Property Get varValue(aVarName As String) As String
    varValue = P_SCOPE.varValue(aVarName)
End Property

Public Property Let varValue(aVarName As String, aVarValue As String)
    If GeneratedTree Then
        P_SCOPE.varValue(aVarName) = aVarValue
    End If
End Property

'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' METHODS:
Private Function Absolute(expression As String) As Double
    Absolute = Abs(CDbl(expression))
End Function

Private Function aCeiling(ByRef expression As String) As Double
    aCeiling = Ceiling(CDbl(expression))
End Function

Public Sub AddConstant(aValue As String, aKey As String)
    P_SCOPE.AddConstant aValue, aKey
End Sub

Private Sub AddToMap(aValue As Long, aTarget As ClusterTree)
    If aTarget.resulstMap = vbNullString Then
        aTarget.resulstMap = CStr(aValue)
    Else
        aTarget.resulstMap = aTarget.resulstMap & P_SEPARATORCHAR & aValue
    End If
End Sub
Private Sub AddVariable(ByRef variable As String, ByRef aKey As String)
    P_SCOPE.AddVariable variable, aKey
End Sub

Private Function aFloor(ByRef expression As String) As Double
    aFloor = Floor(CDbl(expression))
End Function

Private Sub AppendToBuffer(ByRef aBuffer As ClusterTree, ByRef aToken As token)
    aBuffer.index = aBuffer.index + 1
    On Error GoTo BufferAppend_errHandler
    aBuffer.Storage(aBuffer.index) = aToken
    Exit Sub
BufferAppend_errHandler:
    err.Clear
    ExpandBuffer aBuffer
    aBuffer.Storage(aBuffer.index) = aToken
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

Private Function ApplyLawOfSigns(ByRef expression As String) As String
    Dim tmpResult As String
    
    tmpResult = expression
    Do While tmpResult Like "*[+-][+-]*"
        SignSubstitution tmpResult, "+-", op_minus
        SignSubstitution tmpResult, "-+", op_minus
        SignSubstitution tmpResult, "++", op_plus
        SignSubstitution tmpResult, "--", op_plus
    Loop
    ApplyLawOfSigns = tmpResult
End Function

Private Function ArcCos(ByRef expression As String) As Double
    Dim tmpEval As Double
        
    tmpEval = CDbl(expression)
    tmpEval = Atn(-tmpEval / Sqr(-tmpEval * tmpEval + 1)) + 2 * Atn(1)
    If P_DEGREES Then
        tmpEval = tmpEval * 180 / PI
    End If
    ArcCos = tmpEval
End Function

Private Function ArcSin(ByRef expression As String) As Double
    Dim tmpEval As Double
    
    tmpEval = CDbl(expression)
    tmpEval = Atn(tmpEval / Sqr(-tmpEval * tmpEval + 1))
    If P_DEGREES Then
        tmpEval = tmpEval * 180 / PI
    End If
    ArcSin = tmpEval
End Function

Private Function ArcTan(ByRef expression As String) As Double
    Dim tmpEval As Double

    tmpEval = CDbl(expression)
    tmpEval = Atn(tmpEval)
    If P_DEGREES Then
        tmpEval = tmpEval * 180 / PI
    End If
    ArcTan = tmpEval
End Function

Private Function Array_(expression As String) As String
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    Dim iCounter As Long
    
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    tmpEval = tmpData(iCounter)
    For iCounter = LB + 1 To UB
        tmpEval = tmpEval & P_SEPARATORCHAR & tmpData(iCounter)
    Next iCounter
    Array_ = d_lCurly & tmpEval & d_rCurly
End Function

''' <summary>
''' Turns a like Java array string ({{*};{*}}) into a 1D  or 2D VBA array with n rows and m columns.
''' </summary>
''' <param name="StrArray">The string definition used to create the array.</param>
''' <param name="OutStrArray">The array to return in.</param>
Public Function ArrayFromString(ByRef strArray As String) As String()
    Dim aIndex As Long
    Dim arrCapacity As Long
    Dim curChar As String
    Dim i As Long, j As Long
    Dim NumCols As Long
    Dim NumRows As Long
    Dim OneDarr As Boolean
    Dim OutStrArray() As String
    Dim StrCopy As String
    Dim StrLen As Long
    Dim tmpStr() As Variant
    Dim VectorEndPos As Long
    Dim VectorOpenFlag As Boolean
    Dim VectorStartPos As Long
    Dim OpenCBrackets As Long
    
    StrCopy = ReconstructLiteralStrings(strArray, Join$(Split(strArray, d_Space), vbNullString))
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
        Do While Not curChar = d_lCurly
            i = i + 2
            If i > StrLen Then
                Exit Do
            End If
            curChar = MidB$(StrCopy, i, 2)
        Loop
        VectorOpenFlag = (curChar = d_lCurly)
        If VectorOpenFlag Then
            VectorStartPos = i + 2
            OpenCBrackets = OpenCBrackets + 1
        End If
        Do While VectorOpenFlag
            i = i + 2
            curChar = MidB$(StrCopy, i, 2)
            If curChar = d_lCurly Then
                OpenCBrackets = OpenCBrackets + 1
            Else
                If curChar = d_rCurly Then
                    OpenCBrackets = OpenCBrackets - 1
                End If
            End If
            VectorOpenFlag = Not (curChar = d_rCurly) Or OpenCBrackets
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
    OneDarr = (NumRows = 0)
    If OneDarr Then
        ReDim OutStrArray(0 To NumCols)
    Else '2D array
        ReDim OutStrArray(0 To NumRows, 0 To NumCols)
    End If
    For i = 0 To NumRows
        For j = 0 To NumCols
            If OneDarr Then
                OutStrArray(j) = tmpStr(i)(j)
            Else
                OutStrArray(i, j) = tmpStr(i)(j)
            End If
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
                    tmpData = tmpData & d_lCurly & GetArrItm(InputArray, MultiDimArr, i, j)
                Else
                    tmpData = tmpData & P_SEPARATORCHAR & GetArrItm(InputArray, MultiDimArr, i, j)
                End If
                If j = UB2 Then
                    tmpData = tmpData & d_rCurly
                End If
            Next j
            If i < UB Then
                tmpData = tmpData & P_SEPARATORCHAR
            End If
        Next i
        ArrayToString = d_lCurly & tmpData & d_rCurly
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
Private Function Asc_(ByRef expression As String) As Long
    Asc_ = AscW(FormatLiteralString(expression, True))
End Function
Private Function average(ByRef expression As String) As Double
    Dim g As Long
    Dim tmpData() As String
    Dim tmpEval As Double

    tmpEval = 0
    tmpData() = Split(expression, P_SEPARATORCHAR)
    For g = LBound(tmpData) To UBound(tmpData)
        tmpEval = tmpEval + CDbl(tmpData(g))
    Next g
    tmpEval = tmpEval / (UBound(tmpData) - LBound(tmpData) + 1)
    average = tmpEval
End Function

Private Sub BottomLevelEval(ByRef aToken As token)
    On Error GoTo BLevelEval_errHanlder
    If aToken.OperationToken < 100 Then 'Arithmetic operators
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
                            aToken.EvalResult = CBool(CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn))
                        Else
                            aToken.EvalResult = CBool(MidB$(CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn), 3))
                        End If
                    Else
                        aToken.EvalResult = CBool(CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn))
                    End If
                Else
                    aToken.EvalResult = aToken.Arg1.Operand
                End If
        End Select
    Else
        If aToken.OperationToken < 200 Then 'Comparison operators
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
                Case OperatorToken.otLessThanOrEqual
                    aToken.EvalResult = CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn) <= _
                                        CastOPtype(aToken.Arg2.Operand, aToken.Arg2.NegationFlagOn)
                Case OperatorToken.otLike
                    aToken.EvalResult = CastOPtype(aToken.Arg1.Operand, aToken.Arg1.NegationFlagOn) Like _
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
    Exit Sub
BLevelEval_errHanlder:
    If Not P_FORCE_BOOLEAN Then
        aToken.EvalResult = e_ValueError
    Else
        aToken.EvalResult = False
    End If
    P_ERRTYPE = errEvalError
End Sub

Private Sub BuildErrMessage(errorType As ExpressionErrors, Optional AditionalContext As String)
    P_ERRTYPE = errorType
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

Private Function CastCase(ByRef expression As String) As String
    If Not NotIsPI(expression) Then
        CastCase = LCase$(expression)     'Case insensitive for PI
    Else
        CastCase = expression
    End If
End Function

Private Function CastOPtype(ByRef strOperand As String, ByRef Negate As Boolean) As Variant
    If strOperand <> vbNullString Then
        Dim OpAscw As Long
        OpAscw = AscW(strOperand)
        Select Case OpAscw
            Case 48 To 57, 43, 45, 46
                CastOPtype = CDbl(strOperand)
            Case Else
                If IsLiteralString(strOperand) Then 'Literal strings like ['string']
                    CastOPtype = FormatLiteralString(strOperand)
                Else
                    If Not Negate Then
                        CastOPtype = -1 * CBool(strOperand)
                    Else
                        CastOPtype = -1 * (Not CBool(strOperand))
                    End If
                End If
        End Select
    Else
        CastOPtype = strOperand
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
    If AssignedExpression Then
        CheckVarValues = P_SCOPE.DefinedScope
    Else
        CheckVarValues = True
    End If
End Function

Private Function Choose(ByRef expression As String) As String
    Dim LB As Long
    Dim UB As Long
    Dim tmpData() As String
    Dim index As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    index = Fix(CDbl(tmpData(LB))) + LB
    If index > LB And (index <= UB) Then
        Choose = tmpData(index)
    Else
        Choose = e_ValueError
    End If
End Function
Private Function Chr_(ByRef expression As String) As String
    Chr_ = ToLiteralString(ChrW(CLng(expression))) 'Expected numeric value
End Function
''' <summary>
''' Traverses and evaluates the current parse tree.
''' </summary>
Private Function Compute() As String
    Dim b As Long
    Dim t As Long
    Dim i As Long
    Dim OperationIndex As Long
    Dim BaseIndex As Long
    Dim PrevOP1 As String
    Dim PrevOP2 As String
    
    BaseIndex = UBound(SubTreeData) + 1
    For b = LBound(EvalTree) To UBound(EvalTree)              'Loop all subexpressions
        OperationIndex = BaseIndex
        For t = 0 To EvalTree(b).index                         'Loop all tokens
            OperationIndex = OperationIndex + 1
            If Not EvalTree(b).Storage(t).ConstantToken Then   'Gallop
                If P_GALLOPING_MODE Then
                    PrevOP1 = EvalTree(b).Storage(t).Arg1.Operand
                    PrevOP2 = EvalTree(b).Storage(t).Arg2.Operand
                End If
                GetOperands EvalTree(b).Storage(t), EvalTree(b), BaseIndex
                BottomLevelEval EvalTree(b).Storage(t)
                If P_GALLOPING_MODE Then
                    EvalTree(b).Storage(t).ConstantToken = (PrevOP1 = EvalTree(b).Storage(t).Arg1.Operand And _
                                                            PrevOP2 = EvalTree(b).Storage(t).Arg2.Operand)
                End If
            End If
        Next t
        If Not EvalTree(b).CompCluster Then
            EvalTree(b).EvalResult = EvalTree(b).Storage(t - 1).EvalResult
        Else            'The ClusterTree contains a composite function args as expression
            Dim tmpResult() As String
            tmpResult() = Split(EvalTree(b).resulstMap, P_SEPARATORCHAR)
            '@--------------------------------------------------------------------
            ' Loop all sub-expression tokens results
            For i = LBound(tmpResult) To UBound(tmpResult)
                tmpResult(i) = EvalTree(b).Storage(CLng(tmpResult(i))).EvalResult
            Next i
            If Not EvalTree(b).CompArrCluster Then   'Function Argument
                EvalTree(b).EvalResult = Join$(tmpResult, P_SEPARATORCHAR)
            Else        'Array function Argument
                EvalTree(b).EvalResult = JoinArrFunctArg(tmpResult, EvalTree(b).ClusterArrBounds(0), EvalTree(b).ClusterArrBounds(1))
            End If
        End If
    Next b
    Compute = EvalTree(b - 1).EvalResult
    ComputedTree = True
End Function

Private Function Cosin(ByRef expression As String) As Double
    Dim tmpEval As Double
    
    tmpEval = CDbl(expression)
    If P_DEGREES Then
        tmpEval = tmpEval * PI / 180
    End If
    Cosin = Cos(tmpEval)
End Function

Private Function CountParentheses(ByRef expression As String, ByRef ParenthesisChar As String) As Long
    Dim tmpResult As Long
    Dim SearchIndex As Long
    Dim OpenedClosedP As Long
    
    OpenedClosedP = InStrB(1, expression, ParenthesisChar)
    Do While OpenedClosedP
        tmpResult = tmpResult + 1
        SearchIndex = OpenedClosedP + 2
        OpenedClosedP = InStrB(SearchIndex, expression, ParenthesisChar)
    Loop
    CountParentheses = tmpResult
End Function

Public Function Create(ByRef aExpression As Variant) As CSVexpressions
    If aExpression <> vbNullString Then
        ExprToEval = FormatEntry(CStr(aExpression))
        If ExprToEval <> FormatEntry(P_EXPRESSION) Then
            P_EXPRESSION = aExpression
            VariablesInit ExprToEval
            Parse ExprToEval
        End If
        AssignedExpression = True
    End If
    Set Create = Me
End Function
Private Function Date_(ByRef expression As String) As String
    Date_ = ToLiteralString(Date)
End Function
Private Function DateAdd_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = ToLiteralString(DateAdd( _
                                FormatLiteralString(tmpData(LB), True), _
                                CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                                FormatLiteralString(tmpData(UB), True) _
                                ))
        Case Else
            tmpEval = e_ValueError
    End Select
    DateAdd_ = tmpEval
End Function
Private Function DateDiff_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = DateDiff( _
                                FormatLiteralString(tmpData(LB), True), _
                                FormatLiteralString(tmpData(LB + 1), True), _
                                FormatLiteralString(tmpData(UB), True) _
                                )
        Case 4
            tmpEval = DateDiff( _
                                FormatLiteralString(tmpData(LB), True), _
                                FormatLiteralString(tmpData(LB + 1), True), _
                                FormatLiteralString(tmpData(LB + 2), True), _
                                CLng(FormatLiteralString(tmpData(UB), True)) _
                                )
        Case 5
            tmpEval = DateDiff( _
                                FormatLiteralString(tmpData(LB), True), _
                                FormatLiteralString(tmpData(LB + 1), True), _
                                FormatLiteralString(tmpData(LB + 2), True), _
                                CLng(FormatLiteralString(tmpData(LB + 3), True)), _
                                CLng(FormatLiteralString(tmpData(UB), True)) _
                                )
        Case Else
            tmpEval = e_ValueError
    End Select
    DateDiff_ = tmpEval
End Function
Private Function DatePart_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 2
            tmpEval = DatePart( _
                                FormatLiteralString(tmpData(LB), True), _
                                FormatLiteralString(tmpData(UB), True) _
                                )
        Case 3
            tmpEval = DatePart( _
                                FormatLiteralString(tmpData(LB), True), _
                                FormatLiteralString(tmpData(LB + 1), True), _
                                CLng(FormatLiteralString(tmpData(UB), True)) _
                                )
        Case 4
            tmpEval = DatePart( _
                                FormatLiteralString(tmpData(LB), True), _
                                FormatLiteralString(tmpData(LB + 1), True), _
                                CLng(FormatLiteralString(tmpData(LB + 2), True)), _
                                CLng(FormatLiteralString(tmpData(UB), True)) _
                                )
        Case Else
            tmpEval = e_ValueError
    End Select
    DatePart_ = tmpEval
End Function
Private Function DateSerial_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = ToLiteralString(DateSerial( _
                                CLng(FormatLiteralString(tmpData(LB), True)), _
                                CLng(FormatLiteralString(tmpData(LB + 1), True)), _
                                CLng(FormatLiteralString(tmpData(UB), True)) _
                                ))
        Case Else
            tmpEval = e_ValueError
    End Select
    DateSerial_ = tmpEval
End Function
Private Function DateValue_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = ToLiteralString(DateValue( _
                                FormatLiteralString(tmpData(LB), True) _
                                ))
        Case Else
            tmpEval = e_ValueError
    End Select
    DateValue_ = tmpEval
End Function
Private Function Day_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = Day( _
                            FormatLiteralString(tmpData(LB), True) _
                            )
        Case Else
            tmpEval = e_ValueError
    End Select
    Day_ = tmpEval
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

Private Function DDB_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 4
            tmpEval = DDB(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case 5
            tmpEval = DDB(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 3), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case Else
            tmpEval = e_ValueError
    End Select
    DDB_ = tmpEval
End Function

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
                    BuildErrMessage errVariableNotAssigned, "The values of the current variables are: " & currentVarValues()
                    Exit Function
                End If
            End If
            On Error GoTo Eval_errHandler
            P_RESULT = Compute()
        End If
    End If
    If P_FORMATRESULT Then 'Format: trim ['] chars
        Eval = FormatLiteralString(CStr(P_RESULT), True)
    Else
        Eval = P_RESULT
    End If
    Exit Function
Eval_errHandler:
    P_RESULT = vbNullString
    Eval = P_RESULT
    BuildErrMessage errEvalError, err.Description & "."
End Function

Private Function EvalFunction(ByRef Argument As String, ByRef FunctionName As String, Optional IsUDF As Boolean = False) As String
    On Error GoTo EvalFunction_errHandler
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
            Case "strArray"
                EvalFunction = Array_(Argument)
            Case "ASCII"
                EvalFunction = Asc_(Argument)
            Case "Average"
                EvalFunction = average(Argument)
            Case "Cosin"
                EvalFunction = Cosin(Argument)
            Case "aChoose"
                EvalFunction = Choose(Argument)
            Case "ASCIIchr"
                EvalFunction = Chr_(Argument)
            Case "aCeiling"
                EvalFunction = aCeiling(Argument)
            Case "aDate"
                EvalFunction = Date_(Argument)
            Case "aDateAdd"
                EvalFunction = DateAdd_(Argument)
            Case "aDateDiff"
                EvalFunction = DateDiff_(Argument)
            Case "aDatePart"
                EvalFunction = DatePart_(Argument)
            Case "aDateSerial"
                EvalFunction = DateSerial_(Argument)
            Case "aDateValue"
                EvalFunction = DateValue_(Argument)
            Case "aDay"
                EvalFunction = Day_(Argument)
            Case "aDDB"
                EvalFunction = DDB_(Argument)
            Case "MatrixDeterminant"
                EvalFunction = Determinant(Argument)
            Case "ExpEuler"
                EvalFunction = ExpEuler(Argument)
            Case "aFloor"
                EvalFunction = aFloor(Argument)
            Case "aFormat"
                EvalFunction = Format_(Argument)
            Case "aFV"
                EvalFunction = FV_(Argument)
            Case "FunctionZero"
                EvalFunction = Zero(Argument)
            Case "Gamma"
                EvalFunction = tGamma(Argument)
            Case "aHour"
                EvalFunction = Hour_(Argument)
            Case "aIff"
                EvalFunction = Iff_(Argument)
            Case "InverseMatrix"
                EvalFunction = InverseMatrix(Argument)
            Case "aIPMT"
                EvalFunction = IPMT_(Argument)
            Case "aIRR"
                EvalFunction = IRR_(Argument)
            Case "aLeft"
                EvalFunction = Left_(Argument)
            Case "aLen"
                EvalFunction = Len_(Argument)
            Case "LowerCase"
                EvalFunction = LCase_(Argument)
            Case "Logarithm"
                EvalFunction = Logarithm(Argument)
            Case "LgN"
                EvalFunction = LgN(Argument)
            Case "LUdecomposition"
                EvalFunction = LUdecomposition(Argument)
            Case "LUSolveLinearSystem"
                EvalFunction = LUsolve(Argument)
            Case "LN"
                EvalFunction = LN(Argument)
            Case "Max"
                EvalFunction = max(Argument)
            Case "Middle"
                EvalFunction = Mid_(Argument)
            Case "Min"
                EvalFunction = Min(Argument)
            Case "aMinute"
                EvalFunction = Minute_(Argument)
            Case "aMIRR"
                EvalFunction = MIRR_(Argument)
            Case "MatrixMult"
                EvalFunction = MMult(Argument)
            Case "aMonthName"
                EvalFunction = MonthName_(Argument)
            Case "aMonth"
                EvalFunction = Month_(Argument)
            Case "aNow"
                EvalFunction = Now_(Argument)
            Case "aNPER"
                EvalFunction = NPER_(Argument)
            Case "aNPV"
                EvalFunction = NPV_(Argument)
            Case "Percent"
                EvalFunction = Percent(Argument)
            Case "aPMT"
                EvalFunction = PMT_(Argument)
            Case "aPPMT"
                EvalFunction = PPMT_(Argument)
            Case "Power"
                EvalFunction = Power(Argument)
            Case "aPV"
                EvalFunction = PV_(Argument)
            Case "aRATE"
                EvalFunction = RATE_(Argument)
            Case "aReplace"
                EvalFunction = Replace_(Argument)
            Case "aRight"
                EvalFunction = Right_(Argument)
            Case "Sign"
                EvalFunction = Sign(Argument)
            Case "Sine"
                EvalFunction = Sine(Argument)
            Case "SolveLinearSystem"
                EvalFunction = Solve(Argument)
            Case "aSLN"
                EvalFunction = SLN_(Argument)
            Case "SquareRoot"
                EvalFunction = SquareRoot(Argument)
            Case "Tangent"
                EvalFunction = Tangent(Argument)
            Case "aSwitch"
                EvalFunction = Switch_(Argument)
            Case "aSYD"
                EvalFunction = SYD_(Argument)
            Case "aTimeSerial"
                EvalFunction = TimeSerial_(Argument)
            Case "aTimeValue"
                EvalFunction = TimeValue_(Argument)
            Case "aTrim"
                EvalFunction = Trim_(Argument)
            Case "aUcase"
                EvalFunction = UCase_(Argument)
            Case "aWweekDay"
                EvalFunction = WeekDay_(Argument)
            Case "aWeekDayName"
                EvalFunction = WeekDayName_(Argument)
            Case "aYear"
                EvalFunction = Year_(Argument)
            Case Else
                'Rise an error for not found function
        End Select
    Else
        EvalFunction = EvalUDF(FunctionName, Argument)
    End If
    Exit Function
EvalFunction_errHandler:
    If Not P_FORCE_BOOLEAN Then
        EvalFunction = e_ValueError
    Else
        EvalFunction = False
    End If
    P_ERRTYPE = errEvalError
End Function

Private Function EvalUDF(ByRef UDFname As String, ByRef expression As String) As String
    Dim args As Variant
    Dim tmpEval As String
    Dim UDFidx As Long
    
    UDFidx = GetCBItemIdx(UserDefFunctions, UDFname)
    If UDFidx > -1 Then     'Only declared functions are called
        args = SplitArgs(expression) 'Pass a string array to UDF functions
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

Private Function ExpEuler(ByRef expression As String) As Double
    ExpEuler = Exp(CDbl(expression))
End Function

Private Function Factorial(ByRef expression As String) As String
    Dim g As Long
    Dim tmpEval As Double
    Dim tmpElm As Double
    Dim Fsgn As Long
    
    If InStrB(1, expression, P_SEPARATORCHAR) = 0 And InStrB(1, expression, ".") = 0 Then
        tmpEval = 1
        tmpElm = CDbl(expression)
        Fsgn = Sgn(tmpElm)
        For g = Abs(tmpElm) To 2 Step -1
            tmpEval = tmpEval * g
        Next g
        Factorial = Fsgn * tmpEval
    Else
        'Code here to rise an error
    End If
End Function
Private Function Format_(ByRef expression As String) As String
     Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = ToLiteralString(tmpData(LB))
        Case 2
            tmpEval = ToLiteralString( _
                                        Format( _
                                                FormatLiteralString(tmpData(LB), True), _
                                                FormatLiteralString(tmpData(UB), True) _
                                                ) _
                                    )
        Case 3
            tmpEval = ToLiteralString( _
                                        Format( _
                                                FormatLiteralString(tmpData(LB), True), _
                                                FormatLiteralString(tmpData(LB + 1), True), _
                                                CLng(FormatLiteralString(tmpData(UB), True)) _
                                                ) _
                                    )
        Case 4
            tmpEval = ToLiteralString( _
                                        Format( _
                                                FormatLiteralString(tmpData(LB), True), _
                                                FormatLiteralString(tmpData(LB + 1), True), _
                                                CLng(FormatLiteralString(tmpData(LB + 2), True)), _
                                                CLng(FormatLiteralString(tmpData(UB), True)) _
                                                ) _
                                    )
        Case Else
            tmpEval = e_ValueError
    End Select
    Format_ = tmpEval
End Function
Private Function Floor(ByRef value As Double) As Double
    Dim tmpResult As Double
    
    tmpResult = Fix(value)
    Floor = tmpResult + ((value <> tmpResult) And (value < 0))
End Function

Private Function FormatLiteralString(ByRef aString As String, _
                                        Optional AutoCheckFormating As Boolean = False) As String
    If Not AutoCheckFormating Then
        FormatLiteralString = MidB$(aString, 3, LenB(aString) - 4)
    Else
        If IsLiteralString(aString) Then
            FormatLiteralString = MidB$(aString, 3, LenB(aString) - 4)
        Else
            FormatLiteralString = aString
        End If
    End If
End Function
Private Function FormatEntry(expression As String) As String
    FormatEntry = ReplaceImpliedMult( _
                                Replace(Replace(Replace(RemoveDupNegation(ApplyLawOfSigns _
                                (ReconstructLiteralStrings(CStr(expression), _
                                Join$(Split(expression, d_Space), vbNullString)))), "()", "('')"), "{{", "({{"), "}}", "}})") _
                                )
End Function

Private Function ToLiteralString(ByRef aString As String) As String
    If Not IsLiteralString(aString) Then
        ToLiteralString = d_Apostrophe & aString & d_Apostrophe
    Else
        ToLiteralString = aString
    End If
End Function

Private Function FV_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = FV(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case 4
            tmpEval = FV(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case 5
            tmpEval = FV(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 3), True)), _
                        CLng((FormatLiteralString(tmpData(UB), True))))
        Case Else
            tmpEval = e_ValueError
    End Select
    FV_ = tmpEval
End Function

''' <summary>
''' Find roots of a univariate function ussing the bisection method.
''' at the interval a <= x <= b.
''' </summary>
''' <param name="aFunction">Inline function.</param>
''' <param name="a">Leftmost interval value.</param>
''' <param name="b">Rightmost interval value.</param>
''' <param name="inc">Increment used in calculations.</param>
''' <param name="epsilon">Tolerance (relative error limit).</param>
Public Function fRoots(ByRef aFunction As String, a As Double, _
                        b As Double, Optional inc As Double = 1, _
                        Optional epsilon As Double = 0.0000000001) As String
    Dim tmpRoot As String
    Dim tmpResult As String
    Dim lx As Double, ux As Double
    
    If inc <> 1 Then
        inc = Abs(inc)
    End If
    If a > b Then
        fRoots = e_ValueError
        Exit Function
    End If
    tmpRoot = fZero(aFunction, a, b, epsilon, False)
    If tmpRoot <> e_ValueError Then      'Function with at least one root
        lx = a
        ux = a + inc
        Do While ux <= b
            On Error Resume Next
            tmpRoot = fZeroMBM(aFunction, lx, ux, epsilon, False)
            If tmpRoot <> e_ValueError Then
                If tmpResult <> vbNullString Then
                    tmpResult = tmpResult & "; " & tmpRoot
                Else
                    tmpResult = tmpRoot
                End If
            End If
            lx = lx + inc
            ux = ux + inc
        Loop
        fRoots = d_lCurly & tmpResult & d_rCurly
    Else
        fRoots = tmpRoot
    End If
End Function

Public Function fZero(ByRef aFunction As String, ByVal a As Double, _
                        ByVal b As Double, Optional epsilon As Double = 0.0000000001, _
                        Optional includeVarNames As Boolean = True, _
                        Optional useBisectionMethod As Boolean = False, _
                        Optional printDebugINFO As Boolean = False) As String
    If Not useBisectionMethod Then
        fZero = fZeroMRF(aFunction, a, b, epsilon, includeVarNames, printDebugINFO)
    Else
        fZero = fZeroMBM(aFunction, a, b, epsilon, includeVarNames, printDebugINFO)
    End If
End Function

''' <summary>
''' Find a zero of a univariate function ussing a modified bisection method.
''' aFunction must be a continuous function f(x) for the interval a <= x <= b.
''' </summary>
''' <param name="aFunction">Inline function.</param>
''' <param name="a">Leftmost interval value.</param>
''' <param name="b">Rightmost interval value.</param>
''' <param name="epsilon">Tolerance (relative error limit).</param>
''' <param name="IncludeVarNames">Include variable name in result.</param>
''' <param name="printDebugINFO">Print information like iteraction count.</param>
Private Function fZeroMBM(ByRef aFunction As String, ByVal a As Double, _
                        ByVal b As Double, Optional epsilon As Double = 0.0000000001, _
                        Optional includeVarNames As Boolean = True, _
                        Optional printDebugINFO As Boolean = False) As String
    Dim absFC As Double, absFD As Double
    Dim aZero As Double
    Dim c As Double, d As Double   'Intermediate points |a---:c:----:d:---b|
    Dim evalCounter As Long
    Dim fa As Double, fb As Double
    Dim fc As Double, fd As Double
    Dim fEvalHelper As CSVexpressions
    Dim i As Long
    Dim k As Double
    Dim segmentLen As Double
    Dim tmpFzeroEval As Double
    Dim tmpResult As Double
    Dim tmpVar() As String
    Dim toleranceFlag As Boolean
    Dim varCounter As Long
    Dim varIdx As Long
    Dim varLB As Long
    
    On Error GoTo fZeroMBM_errHandler
    If a > b Then Exit Function
    Set fEvalHelper = New CSVexpressions
    aZero = 10 * epsilon
    With fEvalHelper
        .Create aFunction
        tmpVar() = Split(.currentVariables, "; ")
        varLB = LBound(tmpVar)
        If UBound(tmpVar) - varLB > 0 Then  'Reject multivariate functions
            For i = varLB To UBound(tmpVar)
                If Not IsConstant(tmpVar(i)) Then
                    varCounter = varCounter + 1
                    varIdx = i
                End If
            Next i
            If varCounter > 1 Then
                fZeroMBM = e_ValueError
                Exit Function
            End If
        End If
        k = 0
        'Evaluate the function at a, b
        .varValue(tmpVar(varIdx)) = a
        fa = CDbl(.Eval): evalCounter = evalCounter + 1
        .varValue(tmpVar(varIdx)) = b
        fb = CDbl(.Eval): evalCounter = evalCounter + 1
        Do
            'Divide the interval into three segments, assumes a < b
            segmentLen = (b - a) / 3
            c = a + segmentLen
            d = b - segmentLen
            'Evaluate the function at c, d
            .varValue(tmpVar(varIdx)) = c
            fc = CDbl(.Eval): evalCounter = evalCounter + 1
            .varValue(tmpVar(varIdx)) = d
            fd = CDbl(.Eval): evalCounter = evalCounter + 1
            If fc = 0 Or fd = 0 Then
                Exit Do
            End If
            If fa * fc < 0 Then 'Root in interval a-c
                b = c
                fb = fc
            ElseIf fc * fd < 0 Then
                a = c
                fa = fc
                b = d
                fb = fd
            ElseIf fd * fb < 0 Then
                a = d
                fa = fd
            Else        'The interval does not satisfy the condition fa*fb<0.
                fZeroMBM = e_ValueError
                GoTo fZeroMBM_terminate
            End If
            absFC = Abs(fc)
            absFD = Abs(fd)
            If absFC < absFD Then
                toleranceFlag = (absFC > epsilon) ': Debug.Print c
            Else
                toleranceFlag = (absFD > epsilon) ': Debug.Print d
            End If
            k = k + 1
        Loop While toleranceFlag And k < 100
        If absFC < absFD Then
            tmpResult = c
            tmpFzeroEval = fc
        Else
            tmpResult = d
            tmpFzeroEval = fd
        End If
        .varValue(tmpVar(varIdx)) = tmpResult
        If Round(tmpFzeroEval, Len(CStr(1 / aZero)) - 1) = 0 Then
            If includeVarNames Then
                fZeroMBM = .currentVarValues
            Else
                fZeroMBM = tmpResult
            End If
        Else
            fZeroMBM = e_ValueError
        End If
    End With
fZeroMBM_terminate:
    Set fEvalHelper = Nothing
    If printDebugINFO Then
        Debug.Print "fZeroMBM: " & vbCrLf & vbTab & vbTab & "Evaluations:"; evalCounter; "; Iteractions:"; k
    End If
    Exit Function
fZeroMBM_errHandler:
    fZeroMBM = e_ValueError
End Function

''' <summary>
''' Find a zero of a univariate function ussing a modified regula falsi method.
''' aFunction must be a continuous function f(x) for the interval a <= x <= b.
''' Credits: A NeamvonK [https://www.ajouronline.com/index.php/AJAS/article/view/2982]
''' </summary>
''' <param name="aFunction">Inline function.</param>
''' <param name="a">Leftmost interval value.</param>
''' <param name="b">Rightmost interval value.</param>
''' <param name="epsilon">Tolerance (relative error limit).</param>
''' <param name="IncludeVarNames">Include variable name in result.</param>
''' <param name="printDebugINFO">Print information like iteraction count.</param>
Private Function fZeroMRF(ByRef aFunction As String, ByVal a As Double, _
                        ByVal b As Double, Optional epsilon As Double = 0.0000000001, _
                        Optional includeVarNames As Boolean = True, _
                        Optional printDebugINFO As Boolean = False) As String
    Dim aZero As Double
    Dim c As Double, d As Double
    Dim evalCounter As Long
    Dim fa As Double, fb As Double
    Dim fc As Double, fd As Double
    Dim fEvalHelper As CSVexpressions
    Dim i As Long
    Dim k As Double, ck As Double
    Dim tmpVar() As String
    Dim toleranceFlag As Boolean
    Dim varCounter As Long
    Dim varIdx As Long
    Dim varLB As Long
    
    On Error GoTo fZeroMRF_errHandler
    If a > b Then Exit Function
    Set fEvalHelper = New CSVexpressions
    aZero = 10 * epsilon
    With fEvalHelper
        .Create aFunction
        tmpVar() = Split(.currentVariables, "; ")
        varLB = LBound(tmpVar)
        If UBound(tmpVar) - varLB > 0 Then  'Reject multivariate functions
            For i = varLB To UBound(tmpVar)
                If Not IsConstant(tmpVar(i)) Then
                    varCounter = varCounter + 1
                    varIdx = i
                End If
            Next i
            If varCounter > 1 Then
                fZeroMRF = e_ValueError
                Exit Function
            End If
        End If
        k = 0
        Do
            'Evaluate the function at a, b
            .varValue(tmpVar(varIdx)) = a
            fa = CDbl(.Eval): evalCounter = evalCounter + 1
            .varValue(tmpVar(varIdx)) = b
            fb = CDbl(.Eval): evalCounter = evalCounter + 1
            'Compute approximation point of the root
            c = (a * fb - b * fa) / (fb - fa)
            'Evaluate the function at approximation
            .varValue(tmpVar(varIdx)) = c
            fc = CDbl(.Eval): evalCounter = evalCounter + 1
            If fa * fc < 0 Then 'Root in interval a-c
                'Projection
                ck = Abs(2 * fc / (b - c))
                d = ((1 + ck) * a * fb - b * fa) / ((1 + ck) * fb - fa)
                'Evaluate the function at projection
                .varValue(tmpVar(varIdx)) = d
                fd = CDbl(.Eval): evalCounter = evalCounter + 1
                If fd * fa < 0 Then
                    b = d
                Else
                    b = c
                    a = d
                End If
            ElseIf fa * fc > 0 Then
                ck = Abs(0.5 * fc / (b - c))
                d = ((1 + ck) * b * fa - a * fb) / ((1 + ck) * fa - fb)
                .varValue(tmpVar(varIdx)) = d
                fd = CDbl(.Eval): evalCounter = evalCounter + 1
                If fd * fa > 0 Then
                    a = d
                Else
                    a = c
                    b = d
                End If
            End If
            toleranceFlag = (Abs(fd) > epsilon)
            k = k + 1
        Loop While toleranceFlag And k < 100
        If Round(fd, Len(CStr(1 / aZero)) - 1) = 0 Then
            If includeVarNames Then
                fZeroMRF = .currentVarValues
            Else
                fZeroMRF = d
            End If
        Else
            fZeroMRF = e_ValueError
        End If
    End With
fZeroMRF_terminate:
    Set fEvalHelper = Nothing
    If printDebugINFO Then
        Debug.Print "fZeroMRF: " & vbCrLf & vbTab & vbTab & "Evaluations:"; evalCounter; "; Iteractions:"; k
    End If
    Exit Function
fZeroMRF_errHandler:
    fZeroMRF = e_ValueError
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

Private Function GetArithOpInfo(ByRef expression As String) As TokenInfo
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
    
    PowerSymbolPos = GetPowerSymbolPos(expression)
    If PowerSymbolPos Then                              'Powers first
        GetArithOpInfo.Position = PowerSymbolPos
        GetArithOpInfo.OperationToken = otPower
        GetArithOpInfo.OperatorLen = LenB(op_power)
    Else
        MultSymbolPos = GetOPeratorSymbolPos(expression, op_mult)
        DivSymbolPos = GetOPeratorSymbolPos(expression, op_div)
        IntDivSymbolPos = GetOPeratorSymbolPos(expression, op_intDiv)
        ModSymbolPos = GetOPeratorSymbolPos(expression, op_mod)
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
            AddSymbolPos = GetOPeratorSymbolPos(expression, op_plus)
            If AddSymbolPos = 1 Then                       'Unary symbol
                AddSymbolPos = GetOPeratorSymbolPos(expression, op_plus, AddSymbolPos + 2)
            End If
            If AddSymbolPos Then
                tmpPos = AddSymbolPos - 2
                Do While tmpPos >= 1 And AddSymbolPos > 0  'Skip symbol in functions like [{func+arg}]
                    tmpTestChr = MidB$(expression, tmpPos, 2)
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
                            tmpTestChr = MidB$(expression, tmpPos, 2)
                        Loop While IsLetter(tmpTestChr)
                        If tmpPos > 1 Then
                            tmpPos = tmpPos + 2
                        End If
                        tmpStr = MidB$(expression, tmpPos, AddSymbolPos - tmpPos)
                        If GetFunctionName(tmpStr) <> vbNullString Or tmpStr = "E" Then
                            AddSymbolPos = GetOPeratorSymbolPos(expression, op_plus, AddSymbolPos + 2)
                        Else
                            Exit Do
                        End If
                    Else
                        Exit Do
                    End If
                Loop
            End If
            SubtSymbolPos = GetOPeratorSymbolPos(expression, op_minus)
            If SubtSymbolPos = 1 Then
                SubtSymbolPos = GetOPeratorSymbolPos(expression, op_minus, SubtSymbolPos + 2)
            End If
            If SubtSymbolPos Then
                tmpPos = SubtSymbolPos - 2
                Do While tmpPos >= 1 And SubtSymbolPos > 0 'Skip symbol in functions like [{func-arg}]
                    tmpTestChr = MidB$(expression, tmpPos, 2)
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
                            tmpTestChr = MidB$(expression, tmpPos, 2)
                        Loop While IsLetter(tmpTestChr)
                        If tmpPos > 1 Then
                            tmpPos = tmpPos + 2
                        End If
                        tmpStr = MidB$(expression, tmpPos, SubtSymbolPos - tmpPos)
                        If GetFunctionName(tmpStr) <> vbNullString Or tmpStr = "E" Then
                            SubtSymbolPos = GetOPeratorSymbolPos(expression, op_minus, SubtSymbolPos + 2)
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

Private Function GetArrItm(ByRef arr As Variant, ByRef MultiDimArr As Boolean, _
                            ByRef IdxDim1 As Long, ByRef IdxDim2 As Long) As String
    If MultiDimArr Then
        GetArrItm = CStr(arr(IdxDim1, IdxDim2))
    Else
        GetArrItm = CStr(arr(IdxDim1))
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

Private Function GetRParentPos(ByRef expression As String, ByRef index As Long) As Long
    Dim tmpCounter As Long
    Dim SearchIndex As Long
    Dim OpenedSum As Long
    Dim tmpChr As String
    
    SearchIndex = 1
    OpenedSum = 0
    Do
        tmpChr = MidB$(expression, SearchIndex, 2)
        If tmpChr = d_lParenthesis Then
            OpenedSum = OpenedSum + 1
        Else
            If tmpChr = d_rParenthesis Then
                OpenedSum = OpenedSum - 1
            End If
        End If
        SearchIndex = SearchIndex + 2
    Loop While SearchIndex <= index
    SearchIndex = 1
    tmpCounter = 0
    Do
        tmpChr = MidB$(expression, SearchIndex, 2)
        If tmpChr = d_lParenthesis Then
            tmpCounter = tmpCounter + 1
        Else
            If tmpChr = d_rParenthesis Then
                tmpCounter = tmpCounter - 1
            End If
        End If
        SearchIndex = SearchIndex + 2
    Loop Until (tmpCounter = OpenedSum - 1) And (SearchIndex > index)
    GetRParentPos = SearchIndex - 2
End Function

Private Function GetEvalToken(ByRef expression As String) As token
    Dim TokenStart As Long
    Dim TokenEnd As Long
    Dim TokenDet As TokenInfo
    Dim tmpArgs() As String
    
    TokenDet = GetTokenInfo(expression)
    If TokenDet.Position > 2 Then
        '@--------------------------------------------------------------------
        ' Find token start
        GetTokenStart expression, TokenDet.Position, TokenStart
        '@--------------------------------------------------------------------
        ' Find token end
        GetTokenEnd expression, TokenDet.Position, TokenDet.OperatorLen, TokenEnd
        '@--------------------------------------------------------------------
        ' Fill token data
        GetEvalToken.DefString = MidB$(expression, TokenStart, TokenEnd - TokenStart + 2)
        GetEvalToken.OperationToken = TokenDet.OperationToken
        SplitToken GetEvalToken.DefString, tmpArgs, GetEvalToken.OperationToken
        GetEvalToken.Arg1.DefString = tmpArgs(LBound(tmpArgs))
        GetEvalToken.Arg2.DefString = tmpArgs(UBound(tmpArgs))
        GetEvalToken.Logical = TokenDet.LogicalToken
        GetEvalToken.Arg1.NegationFlagOn = (InStrB(1, GetEvalToken.Arg1.DefString, op_neg) = 1) '~*
        GetEvalToken.Arg2.NegationFlagOn = (InStrB(1, GetEvalToken.Arg2.DefString, op_neg) = 1) '~*
        GetEvalToken.Arg1.funcName = GetFunctionName(GetEvalToken.Arg1.DefString)
        GetEvalToken.Arg1.UDFFunctionIn = IsUDFFunction
        GetEvalToken.Arg2.funcName = GetFunctionName(GetEvalToken.Arg2.DefString)
        GetEvalToken.Arg2.UDFFunctionIn = IsUDFFunction
        GetEvalToken.Arg1.FunctionIn = (GetEvalToken.Arg1.funcName <> vbNullString) Or GetEvalToken.Arg1.UDFFunctionIn
        GetEvalToken.Arg2.FunctionIn = (GetEvalToken.Arg2.funcName <> vbNullString) Or GetEvalToken.Arg2.UDFFunctionIn
        GetEvalToken.Arg1.LinkedIndex = GetIndex(GetEvalToken.Arg1.DefString)
        GetEvalToken.Arg2.LinkedIndex = GetIndex(GetEvalToken.Arg2.DefString)
        If Not GetEvalToken.Arg1.NegationFlagOn Then
            GetEvalToken.Arg1.LinkedVar = P_SCOPE.VarIndex(CastCase(GetEvalToken.Arg1.DefString))
        Else
            GetEvalToken.Arg1.LinkedVar = P_SCOPE.VarIndex(CastCase(MidB$(GetEvalToken.Arg1.DefString, 3)))
        End If
        If Not GetEvalToken.Arg2.NegationFlagOn Then
            GetEvalToken.Arg2.LinkedVar = P_SCOPE.VarIndex(CastCase(GetEvalToken.Arg2.DefString))
        Else
            GetEvalToken.Arg2.LinkedVar = P_SCOPE.VarIndex(CastCase(MidB$(GetEvalToken.Arg2.DefString, 3)))
        End If
        GetEvalToken.Arg1.Implicit = (GetEvalToken.Arg1.LinkedIndex >= 0)
        GetEvalToken.Arg2.Implicit = (GetEvalToken.Arg2.LinkedIndex >= 0)
        GetEvalToken.Arg1.FactorialIn = (InStrB(1, GetEvalToken.Arg1.DefString, op_Factorial) = LenB(GetEvalToken.Arg1.DefString) - 1)
        GetEvalToken.Arg2.FactorialIn = (InStrB(1, GetEvalToken.Arg2.DefString, op_Factorial) = LenB(GetEvalToken.Arg2.DefString) - 1)
    Else            'Retun values
        GetEvalToken.Arg1.DefString = expression
        GetEvalToken.Logical = TokenDet.LogicalToken
        GetEvalToken.Arg1.NegationFlagOn = (InStrB(1, GetEvalToken.Arg1.DefString, op_neg) = 1) '~*
        GetEvalToken.DefString = expression
        GetEvalToken.Arg1.funcName = GetFunctionName(GetEvalToken.Arg1.DefString)
        GetEvalToken.Arg1.UDFFunctionIn = IsUDFFunction
        GetEvalToken.Arg1.FunctionIn = (GetEvalToken.Arg1.funcName <> vbNullString) Or GetEvalToken.Arg1.UDFFunctionIn
        GetEvalToken.Arg1.LinkedIndex = GetIndex(GetEvalToken.Arg1.DefString)
        If Not GetEvalToken.Arg1.NegationFlagOn Then
            GetEvalToken.Arg1.LinkedVar = P_SCOPE.VarIndex(CastCase(GetEvalToken.Arg1.DefString))
        Else
            GetEvalToken.Arg1.LinkedVar = P_SCOPE.VarIndex(CastCase(MidB$(GetEvalToken.Arg1.DefString, 3)))
        End If
        GetEvalToken.Arg1.Implicit = (GetEvalToken.Arg1.LinkedIndex >= 0)
        GetEvalToken.Arg1.FactorialIn = (InStrB(1, GetEvalToken.Arg1.DefString, op_Factorial) = LenB(GetEvalToken.Arg1.DefString) - 1)
    End If
End Function

Private Function GetFunctionName(ByRef expression As String) As String
    Dim EFjCounter As Long
    Dim GFNbool As Boolean
    Dim ExpCopy As String
    Dim tmpPos As Long
    
    ExpCopy = LCase$(expression)
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
        For i = 0 To UserDefFunctions.index
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

Private Function GetLCOpInfo(ByRef expression As String) As TokenInfo
    Dim NotEqualSymbolPos As Long
    Dim GreatterOrEqualSymbolPos As Long
    Dim LessOrEqualSymbolPos As Long
    Dim EqualSymbolPos As Long
    Dim LessThanSymbolPos As Long
    Dim GreatterThanSymbolPos As Long
    Dim LogANDSymbolPos As Long
    Dim LogORSymbolPos As Long
    Dim LogXORSymbolPos As Long
    Dim LikeSymbolPos As Long
    Dim testChar As String
    
    '@--------------------------------------------------------------------
    ' Comparisons first
    Do 'Discard false positives
        EqualSymbolPos = GetOPeratorSymbolPos(expression, op_equal, EqualSymbolPos + 1)
        If EqualSymbolPos > 2 Then
            testChar = MidB$(expression, EqualSymbolPos - 2, 2)
        Else
            testChar = vbNullString
        End If
    Loop While EqualSymbolPos > 0 And testChar Like "[<>]"
    NotEqualSymbolPos = GetOPeratorSymbolPos(expression, op_inequality)
    Do
        GreatterThanSymbolPos = GetOPeratorSymbolPos(expression, op_gt, GreatterThanSymbolPos + 1)
        If GreatterThanSymbolPos Then
            testChar = MidB$(expression, GreatterThanSymbolPos + 2, 2)
        Else
            testChar = vbNullString
        End If
    Loop While GreatterThanSymbolPos > 0 And testChar = op_equal
    Do
        LessThanSymbolPos = GetOPeratorSymbolPos(expression, op_lt, LessThanSymbolPos + 1)
        If LessThanSymbolPos Then
            testChar = MidB$(expression, LessThanSymbolPos + 2, 2)
        Else
            testChar = vbNullString
        End If
    Loop While LessThanSymbolPos > 0 And testChar = op_equal
    GreatterOrEqualSymbolPos = GetOPeratorSymbolPos(expression, op_gtequal)
    LessOrEqualSymbolPos = GetOPeratorSymbolPos(expression, op_ltequal)
    LikeSymbolPos = InStrB(LessThanSymbolPos + 1, expression, op_like)
    If NonZero(EqualSymbolPos, NotEqualSymbolPos, GreatterThanSymbolPos, _
                LessThanSymbolPos, GreatterOrEqualSymbolPos, LessOrEqualSymbolPos, LikeSymbolPos) Then
        GetLCOpInfo.Position = MinNonZero(EqualSymbolPos, NotEqualSymbolPos, GreatterThanSymbolPos, _
                                    LessThanSymbolPos, GreatterOrEqualSymbolPos, LessOrEqualSymbolPos, _
                                    LikeSymbolPos)                                                         'Priority to the first operator
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
            Case LessOrEqualSymbolPos
                GetLCOpInfo.OperationToken = otLessThanOrEqual
                GetLCOpInfo.OperatorLen = LenB(op_ltequal)
            Case LikeSymbolPos
                GetLCOpInfo.OperationToken = otLike
                GetLCOpInfo.OperatorLen = LenB(op_like)
            End Select
        '@--------------------------------------------------------------------
        ' LogicalOperators
        Else
            '@--------------------------------------------------------------------
            ' Check the expression giving higher precedence to AND, XOR
            LogANDSymbolPos = GetOPeratorSymbolPos(expression, op_and, LogANDSymbolPos + 1)
            If NonZero(LogANDSymbolPos) Then
                GetLCOpInfo.Position = LogANDSymbolPos
                GetLCOpInfo.OperationToken = otLogicalAND
                GetLCOpInfo.OperatorLen = LenB(op_and)
            Else
                LogXORSymbolPos = GetOPeratorSymbolPos(expression, op_xor, LogXORSymbolPos + 1)
                If NonZero(LogXORSymbolPos) Then
                    GetLCOpInfo.Position = LogXORSymbolPos
                    GetLCOpInfo.OperationToken = otLogicalXOR
                    GetLCOpInfo.OperatorLen = LenB(op_xor)
                Else
                    Do 'Discard false positives: "||"
                        LogORSymbolPos = GetOPeratorSymbolPos(expression, op_or, LogORSymbolPos + 1)
                        If LogORSymbolPos Then
                            testChar = MidB$(expression, LogORSymbolPos + 2, 2)
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

Private Function GetLogicalNeg(ByRef aToken As token) As Boolean()
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

Private Function GetLParentPos(ByRef expression As String, ByRef RelativePosition As Long) As Long
    Dim tmpCounter As Long
    Dim SearchIndex As Long
    Dim OpenedP As Long
    Dim tmpResult As Long
    
    OpenedP = InStrB(1, expression, d_lParenthesis)
    Do While tmpCounter < RelativePosition
        tmpCounter = tmpCounter + 1
        SearchIndex = OpenedP + 2
        tmpResult = OpenedP
        OpenedP = InStrB(SearchIndex, expression, d_lParenthesis)
    Loop
    GetLParentPos = tmpResult
End Function

Private Sub GetOperand(ByRef CurArg As Argument, _
                        ByRef CurTree As ClusterTree, _
                        ByRef BaseIndex As Long)
                                    
    If CurArg.Implicit Then
        If CurArg.FunctionIn Then 'Implicit function
            If CurArg.FactorialIn Then
                If CurArg.LinkedIndex >= BaseIndex Then 'Data on current tree
                    CurArg.Operand = Factorial(EvalFunction(CurTree.Storage(CurArg.LinkedIndex - BaseIndex).EvalResult, _
                                        CurArg.funcName, CurArg.UDFFunctionIn))
                Else 'Data is main tree
                    CurArg.Operand = Factorial(EvalFunction(EvalTree(CurArg.LinkedIndex).EvalResult, _
                                        CurArg.funcName, CurArg.UDFFunctionIn))
                End If
            Else 'Eval
                If CurArg.LinkedIndex >= BaseIndex Then
                    CurArg.Operand = EvalFunction(CurTree.Storage(CurArg.LinkedIndex - BaseIndex).EvalResult, _
                                        CurArg.funcName, CurArg.UDFFunctionIn)
                Else
                    CurArg.Operand = EvalFunction(EvalTree(CurArg.LinkedIndex).EvalResult, _
                                        CurArg.funcName, CurArg.UDFFunctionIn)
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
                CurArg.Operand = Factorial(P_SCOPE.varValue(CurArg.LinkedVar))
            Else
                CurArg.Operand = P_SCOPE.varValue(CurArg.LinkedVar)
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

Private Sub GetOperands(ByRef CurToken As token, ByRef CurTree As ClusterTree, _
                                    ByRef BaseIndex As Long)
                                    
    '@--------------------------------------------------------------------
    ' Get the first operand
    GetOperand CurToken.Arg1, CurTree, BaseIndex
    '@--------------------------------------------------------------------
    ' Get the second operand if required
    If Not CurToken.OperationToken = otNull Then
        GetOperand CurToken.Arg2, CurTree, BaseIndex
    End If
End Sub

Private Function IsSymbolInLiteralString(ByRef expression As String, SymbolPos As Long) As Boolean
    Dim flagCounter As Long
    Dim tmpPos As Long
    
    tmpPos = InStrB(1, expression, d_Apostrophe)
    If tmpPos Then
        Do While tmpPos < SymbolPos And tmpPos > 0
            flagCounter = flagCounter + 1
            tmpPos = InStrB(tmpPos + 2, expression, d_Apostrophe)
        Loop
    End If
    IsSymbolInLiteralString = flagCounter And 1
End Function

Private Function GetOPeratorSymbolPos(ByRef expression As String, _
                                    ByRef OperatorSymbol As String, _
                                    Optional StartPosition As Long = 1) As Long
    Dim tmpResult As Long
    
    tmpResult = InStrB(StartPosition, expression, OperatorSymbol)
    Do While IsSymbolInLiteralString(expression, tmpResult)
        tmpResult = InStrB(tmpResult + LenB(OperatorSymbol), expression, OperatorSymbol)
    Loop
    GetOPeratorSymbolPos = tmpResult
End Function

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
        Case OperatorToken.otLike
            GetOpSymbol = op_like
        Case OperatorToken.otLogicalAND
            GetOpSymbol = op_and
        Case OperatorToken.otLogicalOR
            GetOpSymbol = op_or
        Case OperatorToken.otLogicalXOR
            GetOpSymbol = op_xor
    End Select
End Function

Private Function GetPowerSymbolPos(ByRef expression As String) As Long
    Dim tmpPos As Long
    Dim tmpResult As Long

    tmpPos = GetOPeratorSymbolPos(expression, op_power)
    Do While tmpPos
        tmpResult = tmpPos
        tmpPos = GetOPeratorSymbolPos(expression, op_power, tmpPos + 2)
    Loop
    GetPowerSymbolPos = tmpResult
End Function

Private Sub GetRootedTree(ByRef SubExpression As String, ByRef tmpReplacement As String, _
                                    ByRef OperationIndex As Long, ByRef outBuffer As ClusterTree)
    Dim vToken As token
    Dim Switch As Boolean
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
        If tmpPos > 0 Then
            SubExpression = MidB$(SubExpression, 1, tmpPos - 1) & tmpReplacement & MidB$(SubExpression, tmpPos + LenB(vToken.DefString))
        Else
            SubExpression = tmpReplacement
        End If
        AppendToBuffer outBuffer, vToken 'Save to target token ClusterTree
        Switch = (SubExpression <> tmpReplacement)
        OperationIndex = OperationIndex + 1
        tmpReplacement = GetSubstStr(OperationIndex)
    Loop While Switch
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
Private Function GetSubTreeData(ByRef expression As String) As String()
    Dim opCount As Long
    Dim CPcount As Long
    
    opCount = CountParentheses(expression, d_lParenthesis)
    CPcount = CountParentheses(expression, d_rParenthesis)
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
            bMark = GetLParentPos(expression, tgOpenP) + 2
            eMark = GetRParentPos(expression, bMark - 2)
            tmpTK = MidB$(expression, bMark, eMark - bMark)
            If UnicToken(PSerial, tmpTK) Then
                PSerial(Uspace) = tmpTK
                Uspace = Uspace + 1
            End If
            tgOpenP = tgOpenP - 1
        Loop While tgOpenP > 0
        Dim curIdx As Long
        Dim fpIdx As Long
        
        ReDim Preserve PSerial(LBound(PSerial) To Uspace - 1)
        FSerial() = PSerial
        For curIdx = UBound(PSerial) To LBound(PSerial) Step -1
            For fpIdx = LBound(PSerial) To UBound(PSerial)
                If fpIdx <> curIdx Then
                    If InStrB(1, FSerial(fpIdx), PSerial(curIdx)) Then
                        tmpKey = GetSubstStr(curIdx)
                        FSerial(fpIdx) = Replace(FSerial(fpIdx), _
                                                    d_lParenthesis & PSerial(curIdx) & d_rParenthesis, _
                                                    tmpKey)
                    End If
                End If
            Next fpIdx
        Next curIdx
    Else
        ReDim FSerial(0 To 0)
        FSerial(0) = expression
    End If
    GetSubTreeData = FSerial
End Function

Private Sub GetTokenEnd(ByRef expression As String, ByRef startIndex As Long, ByRef OPlen As Long, ByRef outLng As Long)
    Dim curChar As String
    Dim lenExpr As Long
    Dim tmpPos As Long
    Dim varLen As Long
    Dim tmpVar As String
    Dim functionFlag As Boolean
    Dim cFlag As Boolean
    Dim OperandInBundle As Boolean
    
    lenExpr = LenB(expression)
    outLng = startIndex + OPlen
    curChar = MidB$(expression, outLng, 2)
    '@--------------------------------------------------------------------
    ' Skip literal strings
    If curChar = d_Apostrophe Then
        Do
            outLng = outLng + 2
            curChar = MidB$(expression, outLng, 2)
        Loop While curChar <> d_Apostrophe And outLng < lenExpr
        If outLng >= lenExpr Then Exit Sub
    End If
    Do
        tmpPos = outLng
        If IsLetter(curChar) Then 'Check for functions [ arg/-funct(...) ] and Sci notation
            Do
                varLen = varLen + 2
                outLng = outLng + 2
                curChar = MidB$(expression, outLng, 2)
            Loop While IsAlphaNumeric(curChar) And tmpPos < lenExpr
            tmpVar = MidB$(expression, tmpPos, varLen)
            functionFlag = (GetFunctionName(tmpVar) <> vbNullString)
            If functionFlag Then   'Ignore [+-] symbols for functions
                outLng = outLng + 2
                curChar = MidB$(expression, outLng, 2)
                If IsPlusOrMinus(curChar) Then
                    outLng = outLng + 2
                    curChar = MidB$(expression, outLng, 2)
                End If
            Else
                If LenB(tmpVar) = 2 Then  'Ignore [+-] symbols for sci not symbol "E"
                    If AscW(tmpVar) = 69 Then
                        If IsPlusOrMinus(curChar) Then
                            outLng = outLng + 2
                            curChar = MidB$(expression, outLng, 2)
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
            curChar = MidB$(expression, outLng, 2)
        End If
        OperandInBundle = (InStrB(1, op_AllItems, curChar))
    Loop While Not OperandInBundle And outLng < lenExpr
    If outLng < lenExpr Then
        outLng = outLng - 2
    Else
        outLng = lenExpr
    End If
End Sub

Private Function GetTokenInfo(ByRef expression As String) As TokenInfo
    Dim tmpResult As TokenInfo
            
    tmpResult = GetArithOpInfo(expression)
    If tmpResult.Position = -1 Then 'Missing arithmetic opetarators.
        If expression Like "*[=<>&|$]*" Then 'Try with logical operators.
            tmpResult = GetLCOpInfo(expression)
            tmpResult.LogicalToken = Not expression Like "'*[=<>&|$]*'"
        Else
            tmpResult.Position = -1
            If expression Like "*[Tt][Rr][Uu][Ee]" Then
                tmpResult.LogicalToken = True
            Else
                If expression Like "*[Ff][Aa][Ll][Ss][Ee]" Then
                    tmpResult.LogicalToken = True
                Else
                    If expression Like "~*" Then
                        tmpResult.LogicalToken = True
                    End If
                End If
            End If
        End If
    End If
    GetTokenInfo = tmpResult
End Function

Private Sub GetTokenStart(ByRef expression As String, ByRef startIndex As Long, ByRef outLng As Long)
    Dim curChar As String
    
    outLng = startIndex - 2
    curChar = MidB$(expression, outLng, 2)
    '@--------------------------------------------------------------------
    ' Skip literal strings
    If curChar = d_Apostrophe Then
        Do
            outLng = outLng - 2
            curChar = MidB$(expression, outLng, 2)
        Loop While curChar <> d_Apostrophe And outLng > 1
        If outLng = 1 Then Exit Sub
    End If
    Do While (InStrB(1, op_AllItems, curChar) = 0) And outLng > 1
        outLng = outLng - 2
        curChar = MidB$(expression, outLng, 2)
        If AscW(curChar) = 46 Then  'Dot "."
            outLng = outLng - 2
            If outLng > 0 Then
                curChar = MidB$(expression, outLng, 2)
            End If
        Else
            If outLng >= 3 Then
                If AscW(MidB$(expression, outLng - 2, 2)) = 69 Then 'Ignore Sci notation "E"
                    outLng = outLng - 4
                    curChar = MidB$(expression, outLng, 2)
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
                PrevChar = MidB$(expression, outLng - 2, 2)
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

Private Function GoBackToOpSymbol(ByRef expression As String, ByRef VarStartPos As Long) As Long
    Dim tmpResult As Long
    
    tmpResult = VarStartPos - 2
    Do While tmpResult >= 1
        If ImplicitMultFlag(MidB$(expression, tmpResult, 2)) Then
            tmpResult = tmpResult - 2
        Else
            Exit Do
        End If
    Loop
    GoBackToOpSymbol = tmpResult
End Function

Private Function Hour_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = Hour( _
                            FormatLiteralString(tmpData(LB), True) _
                            )
        Case Else
            tmpEval = e_ValueError
    End Select
    Hour_ = tmpEval
End Function
Private Function Iff_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As Boolean
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = CBool( _
                            FormatLiteralString(tmpData(LB), True) _
                            )
        Case Else
            Iff_ = e_ValueError
            Exit Function
    End Select
    If tmpEval Then    'Return the TRUE part
        Iff_ = tmpData(LB + 1)
    Else
        Iff_ = tmpData(UB)
    End If
End Function

Private Function IPMT_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 4
            tmpEval = IPmt(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case 5
            tmpEval = IPmt(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 3), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case 6
            tmpEval = IPmt(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 3), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 4), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case Else
            tmpEval = e_ValueError
    End Select
    IPMT_ = tmpEval
End Function
Private Function DefineIRRfunction(ByRef strArray() As String) As String
    Dim arrLB As Long
    Dim idx As Long
    Dim tmpResult As String
    
    arrLB = LBound(strArray)
    tmpResult = strArray(arrLB)
    For idx = arrLB + 1 To UBound(strArray)
        tmpResult = tmpResult & op_plus & _
                    "(" & FormatLiteralString(strArray(idx), True) _
                    & op_div & "(1+iRate)" & op_power & idx - arrLB & ")"
    Next idx
    DefineIRRfunction = tmpResult
End Function

''' <summary>
''' Returns the determinant for the given matrix
''' </summary>
Private Function Determinant(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            Dim aArray() As Double
            Dim pivotingIdx() As Long
            Dim d As Double             'Interchages parity
            Dim i As Long
            
            aArray() = StringTodblArray(ArrayFromString(tmpData(LB)))
            LUdecomp aArray(), pivotingIdx, d
            For i = LBound(aArray) To UBound(aArray)
                d = d * aArray(i, i)
            Next i
            tmpEval = d
        Case Else
            tmpEval = e_ValueError
    End Select
    Determinant = tmpEval
End Function

Private Function IRR_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    Dim strArray() As String
    
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    strArray() = ArrayFromString(tmpData(LB)) 'Get array of data
    Select Case argsCount
        Case 1, 2
            'Attempt to arrive at a positive IRR solution.
            tmpEval = fZeroMBM(DefineIRRfunction(strArray), -0.2, 0.99, 0.00001, False)
            If argsCount = 2 Then
                If CBool(tmpData(UB)) Then
                    If tmpEval = e_ValueError Then 'Positive IRR not found
                        tmpEval = fZeroMBM(DefineIRRfunction(strArray), -0.99, -0.19, 0.00001, False)
                    End If
                End If
            End If
        Case Else
            tmpEval = e_ValueError
    End Select
    IRR_ = tmpEval
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

Public Property Let ImplicitVarValue(aVarName As String, aVarValue As String)
    If GeneratedTree Then
        Dim exprHelper As CSVexpressions
        
        Set exprHelper = New CSVexpressions
        With exprHelper
            .Create aVarValue
            If .errorType = errNone Then
                Set .evalScope = .evalScope.CopyScope(P_SCOPE)
                P_SCOPE.varValue(aVarName) = .Eval
            End If
        End With
    End If
End Property

Private Sub InitBuffer(ByRef aBuffer As ClusterTree)
    aBuffer.Capacity = 128
    ReDim aBuffer.Storage(0 To aBuffer.Capacity - 1)
    aBuffer.index = -1
End Sub

Private Sub InitCBbuffer(ByRef aBuffer As ClusterBuffer)
    aBuffer.Capacity = 128
    ReDim aBuffer.Storage(0 To aBuffer.Capacity - 1)
    aBuffer.index = -1
End Sub

Private Sub InitializeErrHandler()
    P_ERRORDESC = vbNullString
    P_ERRTYPE = errNone
End Sub

Private Function InitializedArray(ByRef arr As Variant) As Boolean
    Dim UB As Long
    
    On Error GoTo err_handler
    UB = UBound(arr)
    InitializedArray = True
    Exit Function
err_handler:
    InitializedArray = False
End Function

Private Function Inverse(ByRef a() As Double, ByRef indx() As Long, _
                        ByRef d As Double) As Double()
'@----------------------------------------------------------------------------------------
' Numerical recipes in C: the art of scientific computing / William H. Press [et al.]
' - 2nd ed. (p.48)
'-----------------------------------------------------------------------------------------
    Dim y() As Double
    Dim col() As Double
    Dim i As Long
    Dim j As Long
    Dim aUB As Long
    Dim aLB As Long
    Dim n As Long
    
    aUB = UBound(a)
    aLB = LBound(a)
    n = aUB - aLB + 1
    ReDim y(1 To n, 1 To n)
    ReDim col(1 To n)
    LUdecomp a, indx, d
    For j = 1 To n
        For i = 1 To n
            col(i) = 0
        Next i
        col(j) = 1
        LUbkSub a, indx, col
        For i = 1 To n
            If Abs(col(i)) = 0 Then col(i) = 0      'Avoid VBA -0 value
            y(i, j) = col(i)
        Next i
    Next j
    Inverse = y
End Function

Private Function InverseMatrix(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            Dim aArray() As Double
            Dim pivotingIdx() As Long
            Dim d As Double             'Interchages parity
            
            aArray() = StringTodblArray(ArrayFromString(tmpData(LB)))
            tmpEval = ArrayToString(Inverse(aArray, pivotingIdx, d))
        Case Else
            tmpEval = e_ValueError
    End Select
    InverseMatrix = tmpEval
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

Private Function IsBoolean(ByRef expression As String) As Boolean
    IsBoolean = (expression = "true")
    If Not IsBoolean Then
        IsBoolean = (expression = "false")
    End If
End Function
Public Function IsConstant(aVarName As String) As Boolean
    IsConstant = P_SCOPE.IsConstant(aVarName)
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
            Case 46, 48 To 57, 65 To 90, 95, 97 To 122
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
    If aString <> vbNullString Then
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
    
    If MaxColIndex > -1 Then
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
    Else
        tmpResult = d_lCurly & Join$(DecompArray, P_SEPARATORCHAR) & d_rCurly
    End If
    JoinArrFunctArg = d_lCurly & tmpResult & d_rCurly
End Function

Private Function LgN(ByRef expression As String) As Double
    Dim tmpData() As String
    Dim tmpEval As Double
    Dim LB As Long, UB As Long
    
    tmpEval = 0
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    If UB - LB = 1 Then
        tmpEval = Log(CDbl(tmpData(LB))) _
                    / Log(CDbl(tmpData(UB)))    'Log(x)/Log(N)
    Else
        tmpEval = e_ValueError
    End If
    LgN = tmpEval
End Function

Private Function LN(ByRef expression As String) As Double
    LN = Log(CDbl(expression))
End Function
Private Function LCase_(ByRef expression As String) As String
    LCase_ = ToLiteralString(LCase(expression))
End Function
Private Function Len_(ByRef expression As String) As Long
    If IsLiteralString(expression) Then
        Len_ = Len(FormatLiteralString(expression))
    Else
        Len_ = Len(expression)
    End If
End Function
Private Function Left_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = ToLiteralString(Left(FormatLiteralString(tmpData(LB), True), 1))
        Case 2
            tmpEval = ToLiteralString( _
                                        Left( _
                                                FormatLiteralString(tmpData(LB), True), _
                                                CLng(FormatLiteralString(tmpData(UB), True)) _
                                                ) _
                                    )
        Case Else
            tmpEval = e_ValueError
    End Select
    Left_ = tmpEval
End Function
Private Function Logarithm(ByRef expression As String) As Double
    Logarithm = Log(CDbl(expression)) / Log(10)
End Function

''' <summary>
''' Solves the set of n linear equations A.X = B. Here
''' a[1,...,n][1,...,n] is input, not as the matrix A but
''' rather as its LU decomposition, determined by the routine
''' LUdecomp. indx[1,...,n] is input as the permutation vector
''' returned from LUdecomp. b[1,...,n] is input as the right-hand
''' side vector B, and returns with the solution vector X
''' </summary>
Private Sub LUbkSub(ByRef a() As Double, ByRef indx() As Long, ByRef b() As Double)
'@----------------------------------------------------------------------------------------
' Numerical recipes in C: the art of scientific computing / William H. Press [et al.]
' - 2nd ed. (p.47)
'-----------------------------------------------------------------------------------------
    Dim i As Long
    Dim ii As Long
    Dim ip As Long
    Dim j As Long
    Dim n As Long
    Dim sum As Double
    Dim aLB As Long
    Dim aUB As Long
    Dim bLB As Long
    
    aUB = UBound(a)
    aLB = LBound(a)
    bLB = LBound(b)
    n = aUB - aLB + 1
    ii = 0
    For i = 1 To n
        ip = indx(i)
        sum = b(ip + bLB - 1)
        b(ip + bLB - 1) = b(i + bLB - 1)
        If ii Then
            For j = ii To i - 1
                sum = sum - a(i + aLB - 1, j + aLB - 1) * b(j + bLB - 1)
            Next j
        ElseIf sum Then
            ii = i
        End If
        b(i + bLB - 1) = sum
    Next i
    For i = n To 1 Step -1
        sum = b(i + bLB - 1)
        For j = i + 1 To n
            sum = sum - a(i + aLB - 1, j + aLB - 1) * b(j + bLB - 1)
        Next j
        b(i + bLB - 1) = sum / a(i + aLB - 1, i + aLB - 1)
    Next i
End Sub

''' <summary>
''' Given a matrix a[1,...,n][1,...,n], this routine replaces
''' it by the LU decomposition of a rowwise permutation of
''' itself. a is input. a is output, indx[1,...,n] is
''' an output vector that records the rows permutation effected
''' by the partial pivoting; d is output +-1 depending whether
''' the number of row interchanges was even or odd, respectively.
''' This routine is used in combination with LUbkSub to solve
''' linear equations or invert matrix.
''' </summary>
Private Sub LUdecomp(ByRef a() As Double, ByRef indx() As Long, _
                        ByRef d As Double)
'@----------------------------------------------------------------------------------------
' Numerical recipes in C: the art of scientific computing / William H. Press [et al.]
' - 2nd ed. (p.46-47)
'-----------------------------------------------------------------------------------------
    Dim n As Long
    Dim i As Long
    Dim imax As Long
    Dim j As Long
    Dim k As Long
    Dim big As Double
    Dim dum As Double
    Dim sum As Double
    Dim temp As Double
    Dim vv() As Double         'vv stores the implicit scaling of each row
    Dim aLB As Long
    Dim aUB As Long
    
    aUB = UBound(a)
    aLB = LBound(a)
    n = aUB - aLB + 1
    ReDim vv(1 To n)
    ReDim indx(1 To n)
    d = 1                       'No row interchanges yet
    For i = 1 To n              'Loop rows to get the implicit scaling information
        big = 0
        For j = 1 To n
            temp = Abs(a(i + aLB - 1, j + aLB - 1))
            If temp > big Then big = temp
        Next j
        If big = 0 Then         'a is a singular matrix
            Exit Sub
        End If
        'No nonzero largest element
        vv(i) = 1 / big          'Save the scaling
    Next i
    For j = 1 To n               'Loop over columns of Crout's method
        For i = 1 To j - 1
            sum = a(i + aLB - 1, j + aLB - 1)
            For k = 1 To i - 1
                sum = sum - a(i + aLB - 1, k + aLB - 1) * a(k + aLB - 1, j + aLB - 1)
            Next k
            a(i + aLB - 1, j + aLB - 1) = sum
        Next i
        big = 0                  'Initialize for search the largest pivot element
        For i = j To n
            sum = a(i + aLB - 1, j + aLB - 1)
            For k = 1 To j - 1
                sum = sum - a(i + aLB - 1, k + aLB - 1) * a(k + aLB - 1, j + aLB - 1)
            Next k
            a(i + aLB - 1, j + aLB - 1) = sum
            dum = vv(i) * Abs(sum)
            If dum >= big Then
                big = dum
                imax = i
            End If
        Next i
        If j <> imax Then         'Need rows interchange?
            For k = 1 To n
                dum = a(imax + aLB - 1, k + aLB - 1)
                a(imax + aLB - 1, k + aLB - 1) = a(j + aLB - 1, k + aLB - 1)
                a(j + aLB - 1, k + aLB - 1) = dum
            Next k
            d = -1 * d           'Change parity
            vv(imax) = vv(j)     'Change scale factor
        End If
        indx(j) = imax
        If a(j + aLB - 1, j + aLB - 1) = 0 Then    'Avoid division by zero in singular matrix
            a(j + aLB - 1, j + aLB - 1) = TINY
        End If
        If j <> n Then           'Divide by the pivot element
            dum = 1 / a(j + aLB - 1, j + aLB - 1)
            For i = j + 1 To n
                a(i + aLB - 1, j + aLB - 1) = a(i + aLB - 1, j + aLB - 1) * dum
            Next i
        End If
    Next j
    Erase vv
End Sub

''' <summary>
''' Return the LU decomposition for the given matrix
''' </summary>
Private Function LUdecomposition(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            Dim aArray() As Double
            Dim pivotingIdx() As Long
            Dim d As Double             'Interchages parity
            
            aArray() = StringTodblArray(ArrayFromString(tmpData(LB)))
            LUdecomp aArray(), pivotingIdx, d
            tmpEval = ArrayToString(aArray)
        Case Else
            tmpEval = e_ValueError
    End Select
    LUdecomposition = tmpEval
End Function

''' <summary>
''' Solves the linear system AX = B using LU decomposition. The function
''' requires that the <expression> argument be composed of the following:
'''                     1-) An array in text form containing the coefficients of all equations.
'''                     2-) An one dimentional array containing the name of each variable
'''                         (the first name will be applied to the first column of coefficients,
'''                         the second name to the second column and so on).
'''                     3-) An one dimentional array containing the right-hand side of each equation,
'''                         (the first name will be applied to the first column of coefficients,
'''                         the second name to the second column and so on).
'''                     4-) A True or False parameter to decide when to include the variable names
'''                         in the result set.
''' </summary>
Private Function LUsolve(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3, 4
            Dim aArray() As Double
            Dim eqCount As Long
            Dim iCounter As Long
            Dim includeNames As Boolean
            Dim mLB As Long
            Dim mUB As Long
            Dim xArray() As String
            Dim bArray() As Double
            Dim pivotingIdx() As Long
            Dim d As Double             'Interchages parity
            Dim tmpElement As String
            
            aArray() = StringTodblArray(ArrayFromString(tmpData(LB)))
            bArray() = StringTodblArray(ArrayFromString(tmpData(LB + 2)))
            mLB = LBound(aArray)
            mUB = UBound(aArray)
            
            eqCount = mUB - mLB + 1
            LUdecomp aArray, pivotingIdx, d             'LU decomposition
            LUbkSub aArray, pivotingIdx, bArray         'Back substitution, returning in bArray
            If argsCount = 4 Then
                xArray() = ArrayFromString(tmpData(LB + 1))
                includeNames = CBool(tmpData(UB))
            End If
            For iCounter = mLB To mUB
                If iCounter > mLB Then
                    tmpEval = tmpEval & P_SEPARATORCHAR & d_Space
                End If
                If includeNames Then
                    tmpElement = FormatLiteralString(xArray(iCounter)) _
                                & d_Space & op_equal & d_Space _
                                & bArray(iCounter)
                Else
                    tmpElement = bArray(iCounter)
                End If
                tmpEval = tmpEval & tmpElement
            Next iCounter
            If Not includeNames Then
                tmpEval = d_lCurly & d_lCurly & tmpEval & d_rCurly & d_rCurly
            End If
        Case Else
            tmpEval = e_ValueError
    End Select
    LUsolve = tmpEval
End Function
Private Function max(ByRef expression As String) As Double
    Dim g As Long
    Dim tmpData() As String
    Dim tmpEval As Double
    Dim compEval As Double

    tmpData() = Split(expression, P_SEPARATORCHAR)
    tmpEval = CDbl(tmpData(LBound(tmpData)))
    For g = LBound(tmpData) + 1 To UBound(tmpData)
        compEval = CDbl(tmpData(g))
        If compEval > tmpEval Then
            tmpEval = compEval
        End If
    Next g
    max = tmpEval
End Function

Private Function Mid_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 2
            tmpEval = ToLiteralString( _
                                        Mid( _
                                                FormatLiteralString(tmpData(LB), True), _
                                                CLng(FormatLiteralString(tmpData(UB), True)) _
                                                ) _
                                    )
        Case 3
            tmpEval = ToLiteralString( _
                                        Mid( _
                                                FormatLiteralString(tmpData(LB), True), _
                                                CLng(FormatLiteralString(tmpData(LB + 1), True)), _
                                                CLng(FormatLiteralString(tmpData(UB), True)) _
                                                ) _
                                    )
        Case Else
            tmpEval = e_ValueError
    End Select
    Mid_ = tmpEval
End Function

Private Function Min(ByRef expression As String) As Double
    Dim g As Long
    Dim tmpData() As String
    Dim tmpEval As Double
    Dim compEval As Double
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
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

Private Function Minute_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = Minute( _
                            FormatLiteralString(tmpData(LB), True) _
                            )
        Case Else
            tmpEval = e_ValueError
    End Select
    Minute_ = tmpEval
End Function

Private Function MIRR_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    Dim idx As Long
    Dim strArray() As String
    Dim dblArray() As Double
    
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    strArray() = ArrayFromString(tmpData(LB)) 'Get array of data
    ReDim dblArray(LBound(strArray) To UBound(strArray))
    For idx = LBound(strArray) To UBound(strArray) 'Copy to double array
        dblArray(idx) = CDbl(FormatLiteralString(strArray(idx), True))
    Next idx
    Select Case argsCount
        Case 3
            tmpEval = MIRR(dblArray, _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(UB), True)))
        Case Else
            tmpEval = e_ValueError
    End Select
    MIRR_ = tmpEval
End Function

''' <summary>
''' Multiplies 2 matrices A and B when A has multiple
''' rows and columns. Matrix B can be a column vector.
''' </summary>
Private Function MMult(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 2
            tmpEval = ArrayToString(MMult_( _
                                StringTodblArray(ArrayFromString(tmpData(LB))), _
                                StringTodblArray(ArrayFromString(tmpData(UB)))) _
                                )
        Case Else
            tmpEval = e_ValueError
    End Select
    MMult = tmpEval
End Function

Private Function MMult_(ByRef a() As Double, ByRef b() As Double) As Double()
    Dim c() As Double
    Dim LBa As Long
    Dim LBa2 As Long
    Dim LBb As Long
    Dim LBb2 As Long
    Dim UBb As Long
    Dim m As Long
    Dim n As Long
    Dim l As Long
    Dim IsVectorb As Boolean
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim sum As Double
    
    LBa = LBound(a)
    LBa2 = LBound(a, 2)
    LBb = LBound(b)
    UBb = UBound(b)
    n = UBound(a) - LBa + 1
    m = UBound(a, 2) - LBa2 + 1
    If UBb - LBb + 1 <> m Then Exit Function    'Required columns from a equal to rows of b
    IsVectorb = Not Is2Darray(b)
    If Not IsVectorb Then
        LBb2 = LBound(b, 2)
        l = UBound(b, 2) - LBb2 + 1
        ReDim c(LBa To n + LBa - 1, LBb2 To l + LBb2 - 1)
    Else
        ReDim c(LBa To n + LBa - 1, LBb To LBb)
    End If
    For i = 1 To n
        If Not IsVectorb Then
            For j = 1 To l
                sum = 0
                For k = 1 To m
                    sum = sum + a(i + LBa - 1, k + LBa2 - 1) * b(k + LBb - 1, j + LBb2 - 1)
                Next k
                c(i + LBa - 1, j + LBb2 - 1) = sum
            Next j
        Else
            For j = 1 To 1 'One column vector
                sum = 0
                For k = 1 To m
                    sum = sum + a(i + LBa - 1, k + LBa2 - 1) * b(k + LBb - 1)
                Next k
                c(i + LBa - 1, j + LBb - 1) = sum
            Next j
        End If
    Next i
    MMult_ = c
End Function

Private Function Month_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = Month( _
                            FormatLiteralString(tmpData(LB), True) _
                            )
        Case Else
            tmpEval = e_ValueError
    End Select
    Month_ = tmpEval
End Function

Private Function MonthName_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = ToLiteralString(MonthName( _
                                            CLng(FormatLiteralString(tmpData(LB), True)) _
                                            ))
        Case 2
            tmpEval = ToLiteralString(MonthName( _
                                            CLng(FormatLiteralString(tmpData(LB), True)), _
                                            CBool(FormatLiteralString(tmpData(UB), True)) _
                                            ))
        Case Else
            tmpEval = e_ValueError
    End Select
    MonthName_ = tmpEval
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

Private Function NPER_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = NPer(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case 4
            tmpEval = NPer(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case 5
            tmpEval = NPer(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 3), True)), _
                        CLng((FormatLiteralString(tmpData(UB), True))))
        Case Else
            tmpEval = e_ValueError
    End Select
    NPER_ = tmpEval
End Function

Private Function NPV_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As Double
    Dim iRate As Double
    Dim pCounter As Long
    Dim LB As Long, UB As Long
    Dim cLooper As Long
    Dim strArray() As String
    
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 2
            strArray() = ArrayFromString(tmpData(UB)) 'Get array of data
            iRate = CDbl(tmpData(LB))
            pCounter = 1
            For cLooper = LBound(strArray) To UBound(strArray)
                tmpEval = tmpEval + CDbl(strArray(cLooper)) / ((1 + iRate) ^ pCounter)
                pCounter = pCounter + 1
            Next cLooper
        Case Else
            tmpEval = e_ValueError
    End Select
    NPV_ = tmpEval
End Function

Private Function NotIsPI(ByRef varName As String) As Boolean
    NotIsPI = Not varName Like "[Pp][Ii]"
End Function

Private Function Now_(ByRef expression As String) As String
    Now_ = ToLiteralString(Now)
End Function

Private Function OPsymbolInArgument(ByRef ArgDefStr As String, ByRef Pattrn As String) As Boolean
    Dim i As Long
    Dim tmpResult As Boolean
    Dim LenStr As Long
    
    i = 1
    LenStr = LenB(ArgDefStr)
    Do
        tmpResult = InStrB(1, MidB$(ArgDefStr, i, 2), Pattrn)
        i = i + 2
    Loop While i <= LenStr And Not tmpResult
    OPsymbolInArgument = tmpResult
End Function

''' <summary>
''' Parses an expression and builds a parse tree. The
''' computed data is ready for traverse and evaluate.
''' </summary>
''' <param name="Expression">The expression to parse.</param>
Private Sub Parse(ByRef expression As String)
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
    LbrCount = CountParentheses(expression, d_lParenthesis)
    RbrCount = CountParentheses(expression, d_rParenthesis)
    If LbrCount <> RbrCount Then
        If LbrCount > RbrCount Then
            BuildErrMessage errUnbalancedBrackets, "At least one right parenthesis [)] is missing."
        Else
            BuildErrMessage errUnbalancedBrackets, "At least one left parenthesis [(] is missing."
        End If
        Exit Sub
    End If
    err.Clear
    SubTreeData() = GetSubTreeData(expression)
    lambdaText = SerializeSubTree(expression, SubTreeData)
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
Private Sub ParseVariables(ByRef expression As String)
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
    lenExpr = LenB(expression)
    Do
        tmpChr = MidB$(expression, curPos, 2)
        varLen = 0
        If IsLetter(tmpChr) Or tmpChr = d_Apostrophe Then
            varInitPos = curPos
            If tmpChr <> d_Apostrophe Then
                Do
                    varLen = varLen + 2
                    curPos = curPos + 2
                    tmpChr = MidB$(expression, curPos, 2)
                Loop While IsExtAlphaNumeric(tmpChr) And curPos < lenExpr
            Else
                curPos = InStrB(varInitPos + 2, expression, d_Apostrophe)
                varLen = curPos - varInitPos + 2
            End If
            If varInitPos > 2 Then
                tmpPos = varInitPos - 2
            Else
                tmpPos = varInitPos
            End If
            If Not IsDigit(MidB$(expression, tmpPos, 2)) Then  'Variables start with a letter
                tmpVar = MidB$(expression, varInitPos, varLen)
                If Not ReservedToken(tmpVar) Then
                    If GetFunctionName(LCase$(tmpVar)) = vbNullString Then
                        If Not IsLiteralString(tmpVar) Then 'Exclude literal strings
                            AddVariable tmpVar, ConstantKey
                        End If
                    End If
                End If
            Else
                tmpVar = MidB$(expression, varInitPos, varLen)
                If Not ReservedToken(tmpVar) Then
                    tmpPos = GoBackToOpSymbol(expression, varInitPos) + 1
                    charsBefore = MidB$(expression, tmpPos + 1, varInitPos - (tmpPos + 1)) 'Chars before variable
                    If IsNumeric(charsBefore) Then 'implied multiplication found
                        If GetFunctionName(LCase$(tmpVar)) = vbNullString Then
                            If Not IsLiteralString(tmpVar) Then
                                AddVariable tmpVar, ConstantKey
                            End If
                        End If
                        expression = MidB$(expression, 1, varInitPos - LenB(charsBefore) - 1) & _
                                        charsBefore & op_mult & tmpVar & _
                                        MidB$(expression, varInitPos + LenB(tmpVar))
                        lenExpr = LenB(expression)
                    End If
                End If
            End If
        End If
        curPos = curPos + 2
    Loop While curPos < lenExpr
End Sub

Private Function PatternToCheckOn(ByRef ArgDefStr As String) As String
    If LenB(ArgDefStr) Then
        Select Case AscW(ArgDefStr)
            Case 43, 45
                PatternToCheckOn = op_AllNotUnaryItems
            Case Else
                PatternToCheckOn = op_AllItems
        End Select
    Else
        PatternToCheckOn = op_AllItems
    End If
End Function

Private Function Percent(ByRef expression As String) As Double
    Percent = CDbl(expression) / 100
End Function

Private Function PMT_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = Pmt(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case 4
            tmpEval = Pmt(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case 5
            tmpEval = Pmt(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 3), True)), _
                        CLng((FormatLiteralString(tmpData(UB), True))))
        Case Else
            tmpEval = e_ValueError
    End Select
    PMT_ = tmpEval
End Function

Private Function PPMT_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 4
            tmpEval = PPmt(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case 5
            tmpEval = PPmt(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 3), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case 6
            tmpEval = PPmt(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 3), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 4), True)), _
                        CLng((FormatLiteralString(tmpData(UB), True))))
        Case Else
            tmpEval = e_ValueError
    End Select
    PPMT_ = tmpEval
End Function

Private Function Power(ByRef expression As String) As Double
    Dim tmpData() As String
    Dim tmpEval As Double
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    If UB - LB = 1 Then
        tmpEval = CDbl(tmpData(LB)) ^ _
                        CDbl(tmpData(UB))
    Else
        tmpEval = e_ValueError
    End If
    Power = tmpEval
End Function

Private Function PV_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = PV(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case 4
            tmpEval = PV(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case 5
            tmpEval = PV(CDbl(FormatLiteralString(tmpData(LB), True)), _
                        CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                        CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                        CDbl(FormatLiteralString(tmpData(LB + 3), True)), _
                        CLng((FormatLiteralString(tmpData(UB), True))))
        Case Else
            tmpEval = e_ValueError
    End Select
    PV_ = tmpEval
End Function

Private Function RATE_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = Rate(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case 4
            tmpEval = Rate(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case 5
            tmpEval = Rate(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 3), True)), _
                        CLng((FormatLiteralString(tmpData(UB), True))))
        Case 6
            tmpEval = Rate(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 3), True)), _
                         CLng(FormatLiteralString(tmpData(LB + 4), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case Else
            tmpEval = e_ValueError
    End Select
    RATE_ = tmpEval
End Function

Private Function Replace_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = ToLiteralString( _
                                        Replace( _
                                                FormatLiteralString(tmpData(LB), True), _
                                                FormatLiteralString(tmpData(LB + 1), True), _
                                                FormatLiteralString(tmpData(UB), True) _
                                                ) _
                                    )
        Case 4
            tmpEval = ToLiteralString( _
                                        Replace( _
                                                FormatLiteralString(tmpData(LB), True), _
                                                FormatLiteralString(tmpData(LB + 1), True), _
                                                FormatLiteralString(tmpData(LB + 2), True), _
                                                CLng(FormatLiteralString(tmpData(UB), True)) _
                                                ) _
                                    )
        Case 5
            tmpEval = ToLiteralString( _
                                        Replace( _
                                                FormatLiteralString(tmpData(LB), True), _
                                                FormatLiteralString(tmpData(LB + 1), True), _
                                                FormatLiteralString(tmpData(LB + 2), True), _
                                                CLng(FormatLiteralString(tmpData(LB + 3), True)), _
                                                CLng(FormatLiteralString(tmpData(UB), True)) _
                                                ) _
                                    )
        Case 6
            tmpEval = ToLiteralString( _
                                        Replace( _
                                                FormatLiteralString(tmpData(LB), True), _
                                                FormatLiteralString(tmpData(LB + 1), True), _
                                                FormatLiteralString(tmpData(LB + 2), True), _
                                                CLng(FormatLiteralString(tmpData(LB + 3), True)), _
                                                CLng(FormatLiteralString(tmpData(LB + 4), True)), _
                                                CLng(FormatLiteralString(tmpData(UB), True)) _
                                                ) _
                                    )
        Case Else
            tmpEval = e_ValueError
    End Select
    Replace_ = tmpEval
End Function

Private Function ReplaceImpliedMult(expression As String) As String
    Dim tmpPos As Long
    Dim tmpStr As String
    Dim LookupPos As Long
    Dim tmpVar As String
    Dim tmpVarInitPos As Long
    
    LookupPos = 1
    tmpStr = expression
    tmpPos = InStrB(LookupPos, tmpStr, d_lParenthesis)
    Do While tmpPos
        If tmpPos > 2 Then
            If InStr(1, op_AllItems, MidB(tmpStr, tmpPos - 2, 2)) = 0 Then
                tmpVarInitPos = tmpPos
                Do While tmpVarInitPos > 1
                    If (InStrB(1, op_AllItems, MidB$(tmpStr, tmpVarInitPos - 2, 2)) <> 0) Then
                        Exit Do
                    End If
                    tmpVarInitPos = tmpVarInitPos - 2
                Loop
                tmpVar = MidB$(tmpStr, tmpVarInitPos, tmpPos - tmpVarInitPos)
                If GetFunctionName(LCase$(tmpVar)) = vbNullString Then 'Implied multiplication found
                    tmpStr = MidB(tmpStr, 1, tmpPos - 1) & op_mult & MidB(tmpStr, tmpPos)
                End If
            End If
            LookupPos = tmpPos + 4
            tmpPos = InStrB(LookupPos, tmpStr, d_lParenthesis)
        Else
            LookupPos = tmpPos + 2
            tmpPos = InStrB(LookupPos, tmpStr, d_lParenthesis)
        End If
    Loop
    ReplaceImpliedMult = tmpStr
End Function

Private Function Right_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = ToLiteralString(Right(FormatLiteralString(tmpData(LB), True), 1))
        Case 2
            tmpEval = ToLiteralString( _
                                        Right( _
                                                FormatLiteralString(tmpData(LB), True), _
                                                CLng(FormatLiteralString(tmpData(UB), True)) _
                                                ) _
                                    )
        Case Else
            tmpEval = e_ValueError
    End Select
    Right_ = tmpEval
End Function

Private Function ReconstructLiteralStrings(InputExpr As String, StoredExp As String) As String
    Dim curPosInput As Long
    Dim lastPosInput As Long
    Dim curPosStored As Long
    Dim lastPosStored As Long
    Dim closingMarkInput As Long
    Dim closingMarkStored As Long
    Dim tmpResult As String
    
    On Error GoTo Reconstruct_errHandler
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
        lastPosStored = InStrB(closingMarkStored, tmpResult, d_Apostrophe)
        lastPosInput = closingMarkInput + 2
        curPosInput = InStrB(lastPosInput, InputExpr, d_Apostrophe)
    Loop
    ReconstructLiteralStrings = tmpResult
    Exit Function
Reconstruct_errHandler:
End Function

Private Function RemoveDupNegation(ByRef expression As String) As String
    Dim tmpResult As String
    
    tmpResult = expression
    Do While tmpResult Like "*~~*"
        tmpResult = Replace(tmpResult, "~~", vbNullString)
    Loop
    RemoveDupNegation = tmpResult
End Function

Private Function ReservedToken(ByRef varName As String) As Boolean
    ReservedToken = True
    If Not IsBoolean(LCase$(varName)) Then 'Exclude logical values
        If IsLikeSciNot(varName) Then ' Like "E#*"
            If Not IsDigit(MidB$(varName, LenB(varName) - 1)) Then 'Exclude Sci notation exp
                If Not AscW(varName) = 69 Then  'Exclude Sci notation token E
                    ReservedToken = False
                End If
            End If
        Else
            If Not AscW(varName) = 69 Then  'Exclude Sci notation token E
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
Private Function SerializeSubTree(ByRef expression As String, ByRef SubExpressionsData() As String) As String
    Dim glUb As Long
    Dim glCounter As Long
    Dim tmpResult As String
    Dim tmpKey As String
    
    tmpResult = expression
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
    If aBuffer.index > -1 Then
        aBuffer.Capacity = aBuffer.index
        ReDim Preserve aBuffer.Storage(0 To aBuffer.Capacity)
    End If
End Sub

Private Function Sign(ByRef expression As String) As Double
    Sign = (Sgn(CDbl(expression)))
End Function

Private Sub SignSubstitution(ByRef expression As String, FindStr As String, Replacement As String)
    If InStrB(1, expression, FindStr) Then
        expression = Replace(expression, FindStr, Replacement)
    End If
End Sub

Private Function SLN_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = SLN(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case Else
            tmpEval = e_ValueError
    End Select
    SLN_ = tmpEval
End Function

Private Function Sine(ByRef expression As String) As Double
    Dim tmpEval As Double
    
    tmpEval = CDbl(expression)
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
        If Is2Darray(tmpArr) Then
            outArr = Array1DFrom2DArr(tmpArr)
            tmpResult(0) = UBound(tmpArr)   'Rows in the array
            tmpResult(1) = UBound(tmpArr, 2)  'Columns in each row
        Else
            outArr = tmpArr
            tmpResult(0) = UBound(tmpArr)
            tmpResult(1) = -1
        End If
    Else
        tmpResult(0) = -1   'Return error values
        tmpResult(1) = -1
    End If
    SplitArrBranch = tmpResult
End Function

''' <summary>
''' Solves the linear system AX = B using Over Relaxation (SOR)
''' iteration. The function requires that the <expression> argument be composed
''' of the following:
'''                     1-) An array in text form containing the coefficients of all equations.
'''                     2-) An one dimentional array containing the name of each variable
'''                         (the first name will be applied to the first column of coefficients,
'''                         the second name to the second column and so on).
'''                     3-) An one dimentional array containing the right-hand side of each equation,
'''                         (the first name will be applied to the first column of coefficients,
'''                         the second name to the second column and so on).
'''                     4-) A True or False parameter to decide when to include the variable names
'''                         in the result set.
''' The program will iterate until the solution is approximated to 9
''' significant digits or until 500 iterations are completed, whichever comes first.
''' </summary>
Private Function Solve(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    Dim tolerance As Double
    
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3, 4
            Dim cArray() As Double
            Dim eqCount As Long
            Dim iCounter As Long
            Dim includeNames As Boolean
            Dim mLB As Long
            Dim mUB As Long
            Dim nArray() As String
            Dim rArray() As Double
            Dim tmpElement As String
            Dim significantDigits As Long
            Dim xArray() As Double
            
            tolerance = 0.0000000001
            significantDigits = Len(CStr(1 / tolerance)) - 1
            cArray() = StringTodblArray(ArrayFromString(tmpData(LB)))
            rArray() = StringTodblArray(ArrayFromString(tmpData(LB + 2)))
            mLB = LBound(cArray)
            mUB = UBound(cArray)
            
            ReDim xArray(mLB To mUB)
            eqCount = mUB - mLB + 1
            SORiteration eqCount, cArray, rArray, xArray, 500, tolerance, 1
            If argsCount = 4 Then
                nArray() = ArrayFromString(tmpData(LB + 1))
                includeNames = CBool(tmpData(UB))
            End If
            For iCounter = mLB To mUB
                If iCounter > mLB Then
                    tmpEval = tmpEval & P_SEPARATORCHAR & d_Space
                End If
                If includeNames Then
                    tmpElement = FormatLiteralString(nArray(iCounter)) _
                                & d_Space & op_equal & d_Space _
                                & Round(xArray(iCounter), significantDigits)
                Else
                    tmpElement = Round(xArray(iCounter), significantDigits)
                End If
                tmpEval = tmpEval & tmpElement
            Next iCounter
            If Not includeNames Then
                tmpEval = d_lCurly & d_lCurly & tmpEval & d_rCurly & d_rCurly
            End If
        Case Else
            tmpEval = e_ValueError
    End Select
    Solve = tmpEval
End Function
Private Sub SORiteration(n As Long, ByRef a() As Double, ByRef b() As Double, _
                         ByRef x() As Double, iter As Long, tol As Double, omega As Double)
''' Hoffman, J. D. (2001). Numerical methods for engineers and scientists (2nd ed., rev.expanded). Marcel Dekker.
    Dim it As Long
    Dim i As Long, j As Long
    Dim dxmax As Double
    Dim residual As Double
    Dim colLB As Long, rowLB As Long
    
    colLB = LBound(a, 2)
    rowLB = LBound(a)
    For it = 1 To iter
        dxmax = 0
        For i = 1 To n
            residual = b(i + rowLB - 1)
            For j = 1 To n
                residual = residual - a(i + rowLB - 1, j + colLB - 1) * x(j + rowLB - 1)
            Next j
            If Abs(residual) > dxmax Then dxmax = Abs(residual)
            x(i + rowLB - 1) = x(i + rowLB - 1) + omega * residual / a(i + rowLB - 1, i + colLB - 1)
        Next i
        If dxmax < tol Then Exit For
    Next it
End Sub

Private Function SplitArgs(ByRef args As String) As String()
    Dim tmpPos As Long
    Dim curChar As String
    Dim SargStart As Long
    Dim SargEnd As Long
    Dim tmpResult As String
    Dim LenArgsStr As Long
    Dim VectorOpenFlag As Boolean
    Dim OpenCBrackets As Long
    Dim i As Long
    
    tmpPos = 1
    LenArgsStr = LenB(args)
    Do
        VectorOpenFlag = (InStrB(tmpPos, args, d_lCurly) = tmpPos)
        If VectorOpenFlag Then 'Currrent argument is an array
            i = tmpPos
            OpenCBrackets = 1
            Do While VectorOpenFlag And i <= LenArgsStr
                i = i + 2
                curChar = MidB$(args, i, 2)
                If curChar = d_lCurly Then
                    OpenCBrackets = OpenCBrackets + 1
                Else
                    If curChar = d_rCurly Then
                        OpenCBrackets = OpenCBrackets - 1
                    End If
                End If
                VectorOpenFlag = Not (curChar = d_rCurly) Or OpenCBrackets
            Loop
            SargStart = tmpPos
            SargEnd = i + 2
            tmpPos = i + 4
        Else
            SargStart = tmpPos
            SargEnd = InStrB(tmpPos, args, P_SEPARATORCHAR)
            If SargEnd = 0 Then
                SargEnd = LenArgsStr + 1
            End If
            tmpPos = SargEnd + 2
        End If
        If tmpResult <> vbNullString Then
            tmpResult = tmpResult & "?" & MidB$(args, SargStart, SargEnd - SargStart)
        Else
            tmpResult = MidB$(args, SargStart, SargEnd - SargStart)
        End If
    Loop While tmpPos < LenArgsStr
    SplitArgs = Split(tmpResult, "?")
End Function

Private Sub SplitToken(ByRef expression As String, ByRef oArray() As String, ByRef OPtoken As OperatorToken)
    Dim opSymbol As String
    Dim tmpPos As Long
    Dim Position As Long
    ReDim oArray(0 To 1)
    
    opSymbol = GetOpSymbol(OPtoken)
    If OPtoken = OperatorToken.otSum Or OPtoken = OperatorToken.otDiff Then
        tmpPos = InStrB(1, expression, opSymbol)
        Do While tmpPos
            Position = tmpPos
            tmpPos = InStrB(tmpPos + 2, expression, opSymbol)
        Loop
        If tmpPos = 1 Then          'unary expression
            oArray(0) = "0"
        Else
            oArray(0) = MidB$(expression, 1, Position - 1)
        End If
        oArray(1) = MidB$(expression, Position + 2)
    Else
        oArray() = Split(expression, opSymbol)
    End If
End Sub

Private Function SquareRoot(ByRef expression As String) As Double
    SquareRoot = Sqr(CDbl(expression))
End Function

Public Function StringTodblArray(ByRef StringArray() As String) As Double()
    Dim i As Long, LB As Long, UB As Long
    Dim j As Long, LB2 As Long, UB2 As Long
    Dim tmpResult() As Double
    Dim IsVector As Boolean
    
    LB = LBound(StringArray)
    UB = UBound(StringArray)
    IsVector = Not Is2Darray(StringArray)
    If Not IsVector Then
        LB2 = LBound(StringArray, 2)
        UB2 = UBound(StringArray, 2)
        ReDim tmpResult(LB To UB, LB2 To UB2)
    Else
        ReDim tmpResult(LB To UB)
    End If
    For i = LB To UB
        If IsVector Then
            tmpResult(i) = CDbl(StringArray(i))
        Else
            For j = LB2 To UB2
                tmpResult(i, j) = CDbl(StringArray(i, j))
            Next j
        End If
    Next i
    StringTodblArray = tmpResult
End Function

Private Function Switch_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim idx As Long
    Dim LB As Long, UB As Long
    Dim tmpData() As String
    Dim tmpEval As Boolean
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount And 1
        Case 0
            idx = LB
            Do
                tmpEval = CBool( _
                            FormatLiteralString(tmpData(idx), True) _
                            )
                idx = idx + 2
            Loop While Not tmpEval And idx <= UB - 1
        Case Else
            Switch_ = e_ValueError
            Exit Function
    End Select
    If tmpEval Then    'Return the TRUE part
        Switch_ = tmpData(idx - 1)
    Else
        Switch_ = "#Null!"
    End If
End Function

Private Function SYD_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 4
            tmpEval = SYD(CDbl(FormatLiteralString(tmpData(LB), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 1), True)), _
                         CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                        CDbl((FormatLiteralString(tmpData(UB), True))))
        Case Else
            tmpEval = e_ValueError
    End Select
    SYD_ = tmpEval
End Function

Private Function Tangent(ByRef expression As String) As Double
    Dim tmpEval As Double

    tmpEval = CDbl(expression)
    If P_DEGREES Then
        tmpEval = tmpEval * PI / 180
    End If
    Tangent = Tan(tmpEval)
End Function

Private Function TimeSerial_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = ToLiteralString(TimeSerial( _
                                CLng(FormatLiteralString(tmpData(LB), True)), _
                                CLng(FormatLiteralString(tmpData(LB + 1), True)), _
                                CLng(FormatLiteralString(tmpData(UB), True)) _
                                ))
        Case Else
            tmpEval = e_ValueError
    End Select
    TimeSerial_ = tmpEval
End Function

Private Function TimeValue_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = ToLiteralString(TimeValue( _
                                FormatLiteralString(tmpData(LB), True) _
                                ))
        Case Else
            tmpEval = e_ValueError
    End Select
    TimeValue_ = tmpEval
End Function

Private Function Trim_(ByRef expression As String) As String
    Trim_ = ToLiteralString(Trim(FormatLiteralString(expression, True)))
End Function

Private Function UCase_(ByRef expression As String) As String
    UCase_ = ToLiteralString(UCase(expression))
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

Private Function tGamma(ByRef expression As String) As Double
    tGamma = Gamma(CDbl(expression))
End Function

Private Sub TokenizeSubExpr(ByRef expression As String, ByRef SubExpressionsData() As String, ByRef outBuffer As ClusterTree)
    Dim tmpReplacement As String
    Dim ExpCopy As String
    Dim tmpArgs() As String
    Dim taIcounter As Long
    Dim OperationIndex As Long
    Dim tmpIndex As Long
    
    tmpIndex = UBound(SubExpressionsData) + 1
    OperationIndex = tmpIndex
    ExpCopy = expression
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
                    AddToMap outBuffer.index, outBuffer
                Next taIcounter
                outBuffer.CompCluster = True
                outBuffer.CompArrCluster = False
        End Select
    Else    'Composite array function argument
        outBuffer.ClusterArrBounds = SplitArrBranch(ExpCopy, tmpArgs)
        If outBuffer.ClusterArrBounds(0) <> -1 Then 'Splitting argument success
            For taIcounter = LBound(tmpArgs) To UBound(tmpArgs)
                GetRootedTree tmpArgs(taIcounter), tmpReplacement, OperationIndex, outBuffer
                AddToMap outBuffer.index, outBuffer
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

Private Function ValidFuntionName(ByRef expression As String, _
                                    ByRef FuntionName As String, _
                                    StartPos As Long) As Boolean
    Dim cLeft As Boolean
    Dim cRight As Boolean
    
    If StartPos > 1 Then
        cLeft = Not IsLetter(MidB$(expression, StartPos - 2, 2))
    Else
        cLeft = True
    End If
    cRight = Not IsLetter(MidB$(expression, StartPos + LenB(FuntionName), 2))
    ValidFuntionName = cLeft And cRight
End Function

Private Function ValidTree() As Boolean
    Dim b As Long
    Dim c As Long
    Dim t As Long
    Dim tmpResult As Boolean
    
    b = LBound(EvalTree)
    c = UBound(EvalTree)
    Do                        'Loop all subexpressions
        t = 0
        Do                    'Loop all tokens
            tmpResult = CheckArgument(EvalTree(b).Storage(t).Arg1.DefString, _
                        EvalTree(b).Storage(t).Arg1.FactorialIn)
            If tmpResult Then
                If EvalTree(b).Storage(t).Arg2.DefString <> vbNullString Then
                    tmpResult = CheckArgument(EvalTree(b).Storage(t).Arg2.DefString, _
                                EvalTree(b).Storage(t).Arg2.FactorialIn)
                Else
                    tmpResult = (EvalTree(b).Storage(t).OperationToken = otNull) 'Check for invalid input like "**" "//" ...
                End If
            End If
            t = t + 1
        Loop While t <= EvalTree(b).index And tmpResult
        b = b + 1
    Loop While b <= c And tmpResult
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
    Dim UB As Long
    
    tmpAssignment() = Split(ReconstructLiteralStrings(vString, Join$(Split(vString, d_Space), vbNullString)), P_SEPARATORCHAR)
    For avIcounter = LBound(tmpAssignment) To UBound(tmpAssignment)
        tmpValues() = Split(tmpAssignment(avIcounter), "=")
        UB = UBound(tmpValues)
        If tmpValues(UB) <> vbNullString Then
            If IsNumeric(tmpValues(UB)) Or IsBoolean(tmpValues(UB)) Or IsLiteralString(tmpValues(UB)) Then
                P_SCOPE.varValue(tmpValues(LBound(tmpValues))) = tmpValues(UB)
            End If
        End If
    Next avIcounter
End Sub

Private Sub VariablesInit(ByRef expression As String)
    P_SCOPE.VariablesInit
    ParseVariables expression
    P_SCOPE.FillPredefinedVars
End Sub

Private Function WeekDay_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = Weekday( _
                            FormatLiteralString(tmpData(LB), True) _
                            )
        Case 2
            tmpEval = Weekday( _
                            FormatLiteralString(tmpData(LB), True), _
                            CLng(FormatLiteralString(tmpData(UB), True)) _
                            )
        Case Else
            tmpEval = e_ValueError
    End Select
    WeekDay_ = tmpEval
End Function

Private Function WeekDayName_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = ToLiteralString(WeekdayName( _
                            CLng(FormatLiteralString(tmpData(LB), True)) _
                            ))
        Case 2
            tmpEval = ToLiteralString(WeekdayName( _
                            CLng(FormatLiteralString(tmpData(LB), True)), _
                            CBool(FormatLiteralString(tmpData(UB), True)) _
                            ))
        Case 3
            tmpEval = ToLiteralString(WeekdayName( _
                            CLng(FormatLiteralString(tmpData(LB), True)), _
                            CBool(FormatLiteralString(tmpData(LB + 1), True)), _
                            CLng(FormatLiteralString(tmpData(UB), True)) _
                            ))
        Case Else
            tmpEval = e_ValueError
    End Select
    WeekDayName_ = tmpEval
End Function

Private Function Year_(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = Year( _
                            FormatLiteralString(tmpData(LB), True) _
                            )
        Case Else
            tmpEval = e_ValueError
    End Select
    Year_ = tmpEval
End Function

Private Function Zero(ByRef expression As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = fZeroMBM( _
                            FormatLiteralString(tmpData(LB), True) _
                            , CDbl(FormatLiteralString(tmpData(LB + 1), True)) _
                            , CDbl(FormatLiteralString(tmpData(UB), True)))
        Case 4
            tmpEval = fZeroMBM( _
                            FormatLiteralString(tmpData(LB), True) _
                            , CDbl(FormatLiteralString(tmpData(LB + 1), True)) _
                            , CDbl(FormatLiteralString(tmpData(LB + 2), True)), _
                            CBool(FormatLiteralString(tmpData(UB), True)))
        Case Else
            tmpEval = e_ValueError
    End Select
    Zero = tmpEval
End Function

