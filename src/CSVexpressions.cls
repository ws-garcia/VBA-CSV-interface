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
Option Base 0
'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' Copyright © 2022-2024 W. García
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
Private PID2 As Double
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
Private Const Tiny As Double = 1E-20
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
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' QR DECOMPOSITION VARIABLES:
'   /** Array for internal storage of decomposition.
'   @serial internal array storage.
'   */
Private QR() As Double
'   /** Row and column dimensions.
'   @serial column dimension.
'   @serial row dimension.
'   */
Private m As Long
Private n As Long
'   /** Array for internal storage of diagonal of R.
'   @serial diagonal of R.
'   */
Private rDiag() As Double
'   /** Array for internal storage of columns with signs to flip.
'   @serial signs flip in R and Q outputs.
'   */
Private signFlip() As Boolean
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
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
    errMissingArgsOrTooManyArgs = 5
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
    PID2 = PI / 2
    e = Exp(1)
    P_SEPARATORCHAR = ";"
    P_GALLOPING_MODE = True
    P_FORMATRESULT = False
    BuildinFunctIDList = "abs;floor;achisq;asin;acos;aerf;afishf;agauss;asc;anorm;atn;astudt;array;avg;betainv;ceil;chisq;chr;cos;choose;date;dateadd;datediff" & _
                            ";datepart;dateserial;datevalue;day;ddb;det;erf;exp;fishf;fit;format;fv;fzero;gamma;gammaln;gauss;hour;ibeta;iff;inverse;ipmt" & _
                            ";irr;lcase;left;len;log;lgn;ln;lsqrsolve;ludecomp;lusolve;max;mid;min;minute;mirr;mmult;mneg;mround;msum;mtranspose;month;monthname" & _
                            ";norm;now;nper;npv;percent;pmt;ppmt;pow;pv;qr;rate;rem;replace;right;round;sgn;sin;sln;solve;sqr;sqrt;studt;switch" & _
                            ";syd;tan;timeserial;timevalue;tinv;tinv_1t;tinv_2t;trim;ucase;weekday;weekdayname;year"
    BuildinFunctNameList = "Absolute;aFloor;ACHISQ;ArcSin;ArcCos;AERF;AFISHF;AGAUSS;ASCII;ANORM;ArcTan;ASTUDT;strArray;Average;BETAINV;aCeiling;CHISQ;ASCIIchr;Cosin" & _
                            ";aChoose;aDate;aDateAdd;aDateDiff;aDatePart;aDateSerial;aDateValue;aDay;aDDB;MatrixDeterminant;ERF" & _
                            ";ExpEuler;FISHF;CurveFit;aFormat;aFV;FunctionZero;Gamma;GammaLN;GAUSS;aHour;iBETA;aIff;InverseMatrix;aIPMT;aIRR;LowerCase;aLeft;aLen;Logarithm" & _
                            ";LgN;LN;LSQRsolve;LUdecomposition;LUSolveLinearSystem;Max;Middle;Min;aMinute;aMIRR;MatrixMult;MatrixNegation;MatrixRound;MatrixSum;MatrixTranspose;aMonth;aMonthName;NORM" & _
                            ";aNow;aNPER;aNPV;Percent;aPMT;aPPMT;Power;aPV;QRdecomposition;aRATE;REM;aReplace;aRight;ROUND;Sign;Sine;aSLN;SolveLinearSystem;SquareRoot;SquareRoot;STUDT;aSwitch" & _
                            ";aSYD;Tangent;aTimeSerial;aTimeValue;TINV;TINV_1T;TINV_2T;aTrim;aUcase;aWweekDay;aWeekDayName;aYear"
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
                                                            ' functions defined in the UDfunctions class module.
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
Public Property Get Constants() As Collection
    Set Constants = P_SCOPE.Constants
End Property

''' <summary>
''' Gets a string with the variables values used for the last evaluation.
''' </summary>
Public Property Get CurrentVariables() As String
    CurrentVariables = P_SCOPE.CurrentVariables
End Property

''' <summary>
''' Gets a string with the variables values used for the last evaluation.
''' </summary>
Public Property Get CurrentVarValues() As String
    CurrentVarValues = P_SCOPE.CurrentVarValues
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
Public Property Get ErrorType() As ExpressionErrors
    ErrorType = P_ERRTYPE
End Property

''' <summary>
''' Sets/gets the scope for the current expression.
''' </summary>
Public Property Get EvalScope() As CSVexpressionsScope
    Set EvalScope = P_SCOPE
End Property

Public Property Set EvalScope(aObject As CSVexpressionsScope)
    Set P_SCOPE = aObject
End Property

''' <summary>
''' Returns the math expression to be evaluated.
''' </summary>
Public Property Get expression() As String
    expression = P_EXPRESSION
End Property

Public Property Get ForceBoolean() As Boolean
    ForceBoolean = P_FORCE_BOOLEAN
End Property

''' <summary>
''' When True, a False is returned on evaluation errors.
''' This can be useful when evaluating piecewise functions.
''' </summary>
Public Property Let ForceBoolean(aValue As Boolean)
    P_FORCE_BOOLEAN = aValue
End Property

''' <summary>
''' Indicates if the results will be converted to standard VBA strings
''' </summary>
Public Property Get FormatResult() As Boolean
    FormatResult = P_FORMATRESULT
End Property

Public Property Let FormatResult(aValue As Boolean)
    P_FORMATRESULT = aValue
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

Public Property Let ImplicitVarValue(aVarName As String, aVarValue As String)
    If GeneratedTree Then
        Dim exprHelper As CSVexpressions
        
        Set exprHelper = New CSVexpressions
        With exprHelper
            .Create aVarValue
            If .ErrorType = errNone Then
                Set .EvalScope = .EvalScope.CopyScope(P_SCOPE)
                P_SCOPE.VarValue(aVarName) = .Eval
            End If
        End With
    End If
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
    VarValue = P_SCOPE.VarValue(aVarName)
End Property

Public Property Let VarValue(aVarName As String, aVarValue As String)
    If GeneratedTree Then
        P_SCOPE.VarValue(aVarName) = aVarValue
    End If
End Property

'#
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' METHODS:
Private Function Absolute(ByRef expression As String, ByRef fName As String) As Double
    On Error GoTo err_handler
    Absolute = Abs(CDbl(expression))
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function aCeiling(ByRef expression As String, ByRef fName As String) As Double
    On Error GoTo err_handler
    aCeiling = Ceiling(CDbl(expression))
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

''' <summary>
''' Computes the inverse of CHISQ function
''' [(c) John C. Pezzullo](https://statpages.info/scicalc.html)
''' </summary>
''' <param name="p">probability.</param>
''' <param name="n">Degrees of freedom.</param>
Private Function ACHISQ(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 2
            tmpEval = ACHISQ_( _
                            CDbl(tmpData(LB)), _
                            CDbl(tmpData(UB)) _
                            )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    ACHISQ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function ACHISQ_(p As Double, n As Double) As Double
    Dim v As Double, dv As Double, X As Double
    
    v = 0.5: dv = 0.5: X = 0
    Do While (dv > 0.000000000000001)
        X = 1 / v - 1: dv = dv / 2
        If (CHISQ_(X, n) > p) Then
            v = v - dv
        Else
            v = v + dv
        End If
    Loop
    ACHISQ_ = X
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

''' <summary>
''' Computes the inverse of ERF function
''' [(c) John C. Pezzullo](https://statpages.info/scicalc.html)
''' </summary>
''' <param name="p">probability.</param>
Private Function AERF(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = AERF_( _
                            CDbl(tmpData(LB)) _
                            )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    AERF = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function AERF_(p As Double) As Double
    AERF_ = AGAUSS_(p / 2 + 0.5) / Sqr(2)
End Function

''' <summary>
'''  Computes the inverse of FISHF function
''' [(c) John C. Pezzullo](https://statpages.info/scicalc.html)
''' </summary>
''' <param name="p">P-value.</param>
''' <param name="n1">Degrees of freedom 1.</param>
''' <param name="n2">Degrees of freedom 2.</param>
Private Function AFISHF(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = AFISHF_( _
                            CDbl(tmpData(LB)), _
                            CDbl(tmpData(LB + 1)), _
                            CDbl(tmpData(UB)) _
                            )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    AFISHF = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function AFISHF_(p As Double, n1 As Double, n2 As Double) As Double
    Dim v As Double, dv As Double, f As Double
    
    v = 0.5: dv = 0.5: f = 0
    Do While (dv > 0.000000000000001)
        f = 1 / v - 1: dv = dv / 2
        If (FISHF_(f, n1, n2) > p) Then
             v = v - dv
        Else
            v = v + dv
        End If
    Loop
    AFISHF_ = f
End Function

Private Function aFloor(ByRef expression As String, ByRef fName As String) As Double
    aFloor = Floor(CDbl(expression))
End Function

''' <summary>
''' Computes the inverse of GAUSS function
''' [(c) John C. Pezzullo](https://statpages.info/scicalc.html)
''' </summary>
''' <param name="p">probability.</param>
Private Function AGAUSS(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = AGAUSS_( _
                            CDbl(tmpData(LB)) _
                            )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    AGAUSS = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function AGAUSS_(p As Double) As Double
    If (p > 0.5) Then
        AGAUSS_ = Sqr(ACHISQ_(2 * (1 - p), 1))
    Else
        AGAUSS_ = -Sqr(ACHISQ_(2 * p, 1))
    End If
End Function

''' <summary>
''' Computes the inverse of NORM function
''' [(c) John C. Pezzullo](https://statpages.info/scicalc.html)
''' </summary>
''' <param name="p">probability.</param>
Private Function ANORM(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = ANORM_( _
                            CDbl(tmpData(LB)) _
                            )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    ANORM = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function ANORM_(p) As Double
    Dim v As Double, dv As Double, z As Double
    v = 0.5: dv = 0.5: z = 0
    Do While (dv > 0.000000000000001)
        z = 1 / v - 1: dv = dv / 2
        If (NORM_(z) > p) Then
             v = v - dv
        Else
            v = v + dv
        End If
    Loop
    ANORM_ = z
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

Private Function ArcCos(ByRef expression As String, ByRef fName As String) As Double
    Dim tmpEval As Double
    
    On Error GoTo err_handler
    tmpEval = CDbl(expression)
    tmpEval = Atn(-tmpEval / Sqr(-tmpEval * tmpEval + 1)) + 2 * Atn(1)
    If P_DEGREES Then
        tmpEval = tmpEval * 180 / PI
    End If
    ArcCos = tmpEval
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function ArcSin(ByRef expression As String, ByRef fName As String) As Double
    Dim tmpEval As Double
    
    On Error GoTo err_handler
    tmpEval = CDbl(expression)
    tmpEval = Atn(tmpEval / Sqr(-tmpEval * tmpEval + 1))
    If P_DEGREES Then
        tmpEval = tmpEval * 180 / PI
    End If
    ArcSin = tmpEval
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function ArcTan(ByRef expression As String, ByRef fName As String) As Double
    Dim tmpEval As Double
    
    On Error GoTo err_handler
    tmpEval = CDbl(expression)
    tmpEval = Atn(tmpEval)
    If P_DEGREES Then
        tmpEval = tmpEval * 180 / PI
    End If
    ArcTan = tmpEval
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

''' <summary>
'''  Computes the remainder of n/d division
''' </summary>
''' <param name="n">Dividend.</param>
''' <param name="d">Divisor.</param>
Private Function aREM(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 2
            tmpEval = REM_( _
                            CDbl(tmpData(LB)), _
                            CDbl(tmpData(UB)) _
                            )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    aREM = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function aROUND(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = ROUND_( _
                            CDbl(tmpData(LB)) _
                            )
        Case 2
            tmpEval = ROUND_( _
                            CDbl(tmpData(LB)), _
                            CDbl(tmpData(UB)) _
                            )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    aROUND = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
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

''' <summary>
''' Turns a like Java array string ({{*};{*}}) into a 1D  or 2D VBA array with n rows and m columns.
''' </summary>
''' <param name="StrArray">The string definition used to create the array.</param>
Public Function ArrayFromString(ByRef strArray As String) As String()
    Dim aIndex As Long
    Dim arrCapacity As Long
    Dim boolOpenArray As Boolean
    Dim curChar As String
    Dim i As Long, j As Long
    Dim NumCols As Long
    Dim NumRows As Long
    Dim OneDarr As Boolean
    Dim OutStrArray() As String
    Dim OpenCBrackets As Long
    Dim StrCopy As String
    Dim StrLen As Long
    Dim tmpStr() As Variant
    Dim VectorEndPos As Long
    Dim VectorOpenFlag As Boolean
    Dim VectorStartPos As Long
    
    StrCopy = ReconstructLiteralStrings(strArray, Join$(Split(strArray, d_Space), vbNullString))
    StrLen = LenB(StrCopy)
    If MidB$(StrCopy, 1, 4) <> "{{" Then 'Missed "{" from input
        Exit Function
    End If
    If MidB$(StrCopy, StrLen - 3) <> "}}" Then 'Missed "}" from input
        Exit Function
    End If
    i = 1
    aIndex = -1
    arrCapacity = 128
    ReDim tmpStr(0 To arrCapacity - 1)
    Do
        curChar = MidB$(StrCopy, i, 2)
        If curChar = d_lCurly Then
            OpenCBrackets = OpenCBrackets + 1
            If (OpenCBrackets And 1) = 0 Then 'Takes care of array syntax
                VectorStartPos = i + 2
            End If
        Else
            If curChar = d_rCurly Then
                OpenCBrackets = OpenCBrackets - 1
                If (OpenCBrackets And 1) Then 'End of vector
                    VectorEndPos = i
                    aIndex = aIndex + 1
                    If aIndex > arrCapacity Then
                        arrCapacity = arrCapacity * 2
                        ReDim Preserve tmpStr(0 To arrCapacity - 1)
                    End If
                    tmpStr(aIndex) = Split(MidB$(StrCopy, VectorStartPos, _
                                        VectorEndPos - VectorStartPos + 1), P_SEPARATORCHAR)
                End If
            End If
        End If
        boolOpenArray = (OpenCBrackets > 0)
        i = i + 2
    Loop While boolOpenArray Or (i <= StrLen)
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
''' Turns a like Java array string ({{*};{*}}) into VBA jagged array with n rows.
''' </summary>
''' <param name="StrArray">The string definition used to create the array.</param>
Public Function ArrayFromString2(ByRef strArray As String) As Variant()
    Dim aIndex As Long
    Dim arrCapacity As Long
    Dim boolOpenArray As Boolean
    Dim curChar As String
    Dim i As Long
    Dim JaggedArrFlag As Boolean
    Dim NumRows As Long
    Dim StrCopy As String
    Dim StrLen As Long
    Dim tmpStr() As Variant
    Dim VectorEndPos As Long
    Dim VectorStartPos As Long
    Dim OpenCBrackets As Long
    
    StrCopy = ReconstructLiteralStrings(strArray, Join$(Split(strArray, d_Space), vbNullString))
    StrLen = LenB(StrCopy)
    If MidB$(StrCopy, 1, 4) <> "{{" Then 'Missed "{" from input
        Exit Function
    End If
    If MidB$(StrCopy, StrLen - 3) <> "}}" Then 'Missed "}" from input
        Exit Function
    End If
    i = 1
    aIndex = -1
    arrCapacity = 128
    ReDim tmpStr(0 To arrCapacity - 1)
    Do
        curChar = MidB$(StrCopy, i, 2)
        If curChar = d_lCurly Then
            OpenCBrackets = OpenCBrackets + 1
            If Not JaggedArrFlag Then
                JaggedArrFlag = (OpenCBrackets > 2)
            End If
            If (OpenCBrackets And 1) = 0 Then 'Takes care of array syntax
                VectorStartPos = i + 2
            End If
        Else
            If curChar = d_rCurly Then
                OpenCBrackets = OpenCBrackets - 1
                If (OpenCBrackets And 1) Then 'End of vector
                    If Not JaggedArrFlag Then
                        VectorEndPos = i
                    Else
                        VectorStartPos = VectorStartPos - 2
                        VectorEndPos = i + 2
                    End If
                    aIndex = aIndex + 1
                    If aIndex > arrCapacity Then
                        arrCapacity = arrCapacity * 2
                        ReDim Preserve tmpStr(0 To arrCapacity - 1)
                    End If
                    If boolOpenArray And Not JaggedArrFlag Then
                        tmpStr(aIndex) = Split(MidB$(StrCopy, VectorStartPos, _
                                        VectorEndPos - VectorStartPos + 1), P_SEPARATORCHAR)
                    Else
                        If JaggedArrFlag Then
                            tmpStr(aIndex) = ArrayFromString2(MidB$(StrCopy, VectorStartPos, _
                                                        VectorEndPos - VectorStartPos))
                            JaggedArrFlag = False
                        End If
                    End If
                End If
            End If
        End If
        boolOpenArray = (OpenCBrackets > 0)
        i = i + 2
    Loop While boolOpenArray Or (i <= StrLen)
    NumRows = aIndex
    ReDim Preserve tmpStr(0 To aIndex)
    ArrayFromString2 = tmpStr
End Function

''' <summary>
''' Turns a mxn array into a like Java array string ({{*};{*}}).
''' </summary>
''' <param name="InputArray">The array to operate.</param>
Public Function ArrayToString(ByRef InputArray As Variant) As String
    Dim i As Long, LB As Long, UB As Long
    Dim j As Long, LB2 As Long, UB2 As Long
    Dim tmpData As String, MultiDimArr As Boolean
    Dim isJaggedArray As Boolean
    
    If IsArray(InputArray) Then
        LB = LBound(InputArray)
        UB = UBound(InputArray)
        MultiDimArr = Is2Darray(InputArray)
        If MultiDimArr Then
            LB2 = LBound(InputArray, 2)
            UB2 = UBound(InputArray, 2)
        Else
            isJaggedArray = IsArray(InputArray(LB))
            If Not isJaggedArray Then
                LB2 = LB
                UB2 = UB
                UB = LB
            End If
        End If
        For i = LB To UB
            If isJaggedArray And IsTwiceJaggedArr(InputArray, i) Then
                If i = LB Then
                    tmpData = ArrayToString(InputArray(i))
                Else
                    tmpData = tmpData & ArrayToString(InputArray(i))
                End If
            Else
                If isJaggedArray Then
                    LB2 = LBound(InputArray(i))
                    UB2 = UBound(InputArray(i))
                End If
                For j = LB2 To UB2
                    If j = LB2 Then
                        tmpData = tmpData & d_lCurly & GetArrItm(InputArray, MultiDimArr, isJaggedArray, i, j)
                    Else
                        tmpData = tmpData & P_SEPARATORCHAR & GetArrItm(InputArray, MultiDimArr, isJaggedArray, i, j)
                    End If
                    If j = UB2 Then
                        tmpData = tmpData & d_rCurly
                    End If
                Next j
            End If
            If i < UB Then
                tmpData = tmpData & P_SEPARATORCHAR
            End If
        Next i
        ArrayToString = d_lCurly & tmpData & d_rCurly
    End If
End Function

Private Function Array_(ByRef expression As String, ByRef fName As String) As String
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    Dim iCounter As Long
    
    On Error GoTo err_handler
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    tmpEval = tmpData(iCounter)
    For iCounter = LB + 1 To UB
        tmpEval = tmpEval & P_SEPARATORCHAR & tmpData(iCounter)
    Next iCounter
    Array_ = d_lCurly & tmpEval & d_rCurly
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function Asc_(ByRef expression As String, ByRef fName As String) As Long
    On Error GoTo err_handler
    Asc_ = AscW(FormatLiteralString(expression, True))
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

''' <summary>
''' Computes the inverse of STUDT function
''' [(c) John C. Pezzullo](https://statpages.info/scicalc.html)
''' </summary>
''' <param name="p">probability.</param>
''' <param name="n">Degrees of freedom.</param>
Private Function ASTUDT(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 2
            tmpEval = ASTUDT_( _
                            CDbl(tmpData(LB)), _
                            CDbl(tmpData(UB)) _
                            )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    ASTUDT = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function ASTUDT_(p As Double, n As Double) As Double
    Dim v As Double, dv As Double, t As Double
    
    v = 0.5: dv = 0.5: t = 0
    Do While (dv > 0.000000000000001)
        t = 1 / v - 1: dv = dv / 2
        If (STUDT_(t, n) > p) Then
            v = v - dv
        Else
            v = v + dv
        End If
    Loop
    ASTUDT_ = t
End Function

Private Function average(ByRef expression As String, ByRef fName As String) As Double
    Dim g As Long
    Dim tmpData() As String
    Dim tmpEval As Double
    
    On Error GoTo err_handler
    tmpEval = 0
    tmpData() = Split(expression, P_SEPARATORCHAR)
    For g = LBound(tmpData) To UBound(tmpData)
        tmpEval = tmpEval + CDbl(tmpData(g))
    Next g
    tmpEval = tmpEval / (UBound(tmpData) - LBound(tmpData) + 1)
    average = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function betacf(X As Double, A As Double, b As Double) As Double
    Dim fpmin As Double
    Dim m As Long
    Dim m2 As Long
    Dim aa As Double, c As Double
    Dim d As Double, del As Double
    Dim H As Double, qab As Double
    Dim qam As Double, qap As Double
    
    fpmin = 1E-30
    m = 1

'   // These q's will be used in factors that occur in the coefficients
    qab = A + b: qap = A + 1: qam = A - 1
    c = 1: d = 1 - qab * X / qap
    If Abs(d) < fpmin Then d = fpmin
    d = 1 / d
    H = d
    Do While m <= 100
        m2 = 2 * m
        aa = m * (b - m) * X / ((qam + m2) * (A + m2))
'       // One step (the even one) of the recurrence
        d = 1 + aa * d
        If Abs(d) < fpmin Then d = fpmin
        c = 1 + aa / c
        If Abs(c) < fpmin Then c = fpmin
        d = 1 / d: H = H * d * c
        aa = -(A + m) * (qab + m) * X / ((A + m2) * (qap + m2))
'       // Next step of the recurrence (the odd one)
        d = 1 + aa * d
        If Abs(d) < fpmin Then d = fpmin
        c = 1 + aa / c
        If Abs(c) < fpmin Then c = fpmin
        d = 1 / d: del = d * c
        H = H * del
        If Abs(del - 1#) < 0.0000003 Then Exit Do
        m = m + 1
    Loop
    betacf = H
End Function

''' <summary>
''' Returns the inverse cumulative beta inverse probability distribution function.
''' That is, if probability = BETADIST(x,...), BETAINV(probability,...) = x.
''' The cumulative beta distribution can be used in project planning to determine
''' likely completion times given a variability and an expected completion time.
''' [(c) John C. Pezzullo](https://statpages.info/scicalc.html)
''' </summary>
''' <param name="p">Probability.</param>
''' <param name="a">Shape parameter alpha.</param>
''' <param name="b">Shape parameter beta.</param>
Private Function BETAINV(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = BETAINV_( _
                            CDbl(tmpData(LB)), _
                            CDbl(tmpData(LB + 1)), _
                            CDbl(tmpData(UB)))
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    BETAINV = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function BETAINV_(p As Double, A As Double, b As Double) As Double
    Const EPS As Double = 0.00000001
    Dim aL As Double
    Dim a1, b1 As Double
    Dim j As Long
    Dim lna As Double, lnb As Double
    Dim pp As Double, t As Double
    Dim u As Double, err As Double
    Dim X As Double, H As Double
    Dim w As Double, afac As Double
    
    a1 = A - 1: b1 = b - 1
    j = 0
    If p <= 0 Then
        BETAINV_ = 0
        Exit Function
    End If
    If p >= 1 Then
        BETAINV_ = 1
        Exit Function
    End If
    If A >= 1 And b >= 1 Then
        If p < 0.5 Then
            pp = p
        Else
            pp = 1 - p
        End If
        t = Sqr(-2 * Log(pp))
        X = (2.30753 + t * 0.27061) / (1 + t * (0.99229 + t * 0.04481)) - t
        If p < 0.5 Then X = -X
        aL = (X * X - 3) / 6
        H = 2 / (1 / (2 * A - 1) + 1 / (2 * b - 1))
        w = (X * Sqr(aL + H) / H) - (1 / (2 * b - 1) - 1 / (2 * A - 1)) * _
            (aL + 5 / 6 - 2 / (3 * H))
        X = A / (A + b * Exp(2 * w))
    Else
        lna = Log(A / (A + b)): lnb = Log(b / (A + b))
        t = Exp(A * lna) / A: u = Exp(b * lnb) / b
        w = t + u
        If p < t / w Then
            X = (A * w * p) ^ (1 / A)
        Else
            X = 1 - ((b * w * (1 - p)) ^ (1 / b))
        End If
    End If
    afac = -GAMMALN_(A) - GAMMALN_(b) + GAMMALN_(A + b)
    Do While j < 10
        If X = 0 Or X = 1 Then BETAINV_ = X: Exit Function
        err = iBETA_(X, A, b) - p
        t = Exp(a1 * Log(X) + b1 * Log(1 - X) + afac)
        u = err / t
        t = u / (1 - 0.5 * MIN_(1, u * (a1 / X - b1 / (1 - X))))
        X = X - t
        If X <= 0 Then X = 0.5 * (X + t)
        If X >= 1 Then X = 0.5 * (X + t + 1)
        If Abs(t) < EPS * X And j > 0 Then Exit Do
        j = j + 1
    Loop
    BETAINV_ = X
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

Private Sub BuildErrMessage(ErrorType As ExpressionErrors, Optional AditionalContext As String)
    P_ERRTYPE = ErrorType
    Select Case P_ERRTYPE
        Case Is = ExpressionErrors.errUnbalancedBrackets
            P_ERRORDESC = "The expression cannot be evaluated due to unbalanced parenthesis count. " & AditionalContext
        Case Is = ExpressionErrors.errSyntaxError
            P_ERRORDESC = "The expression cannot be evaluated due to a syntax error. " & AditionalContext
        Case Is = ExpressionErrors.errEvalError
            P_ERRORDESC = "An error occurred while trying to evaluate the expression. " & AditionalContext
        Case Is = ExpressionErrors.errVariableNotAssigned
            P_ERRORDESC = "The expression cannot be evaluated due to the existence of variables without assigned values. " & AditionalContext
        Case Is = ExpressionErrors.errMissingArgsOrTooManyArgs
            P_ERRORDESC = "Arguments count differ from predefined one.  " & AditionalContext
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

''' <summary>
''' Computes the one-tailed (right-tail) probability value for a chi-square test
''' (i.e., the area under the chi-square distribution from the chi-square value
''' to positive infinity), given the chi-square value and the degrees of freedom.
''' [(c) John C. Pezzullo](https://statpages.info/scicalc.html)
''' </summary>
''' <param name="x">Chi-square value.</param>
''' <param name="n">Degrees of freedom.</param>
Private Function CHISQ(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 2
            tmpEval = CHISQ_( _
                            CDbl(tmpData(LB)), _
                            CDbl(tmpData(UB)) _
                            )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    CHISQ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function CHISQ_(X As Double, n As Double) As Double
    Dim Q As Double
    Dim p As Double
    Dim k As Double
    Dim t As Double
    Dim A As Double
    
    If (X > 1000 Or n > 1000) Then
        Q = NORM_((((X / n) ^ (1 / 3)) + 2 / (9 * n) - 1) / Sqr(2 / (9 * n))) / 2
        If (X > n) Then
            CHISQ_ = Q: Exit Function
        Else
            CHISQ_ = 1 - Q: Exit Function
        End If
    End If
    p = Exp(-0.5 * X)
    If (REM_(n, 2) = 1) Then
        p = p * Sqr(2 * X / PI)
    End If
    k = n
    Do While (k >= 2)
        p = p * X / k: k = k - 2
    Loop
    t = p
    A = n
    Do While (t > 0.000000000000001 * p)
        A = A + 2
        t = t * X / A: p = p + t
    Loop
    CHISQ_ = 1 - p
End Function

Private Function Choose(ByRef expression As String, ByRef fName As String) As String
    Dim LB As Long
    Dim UB As Long
    Dim tmpData() As String
    Dim index As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    index = Fix(CDbl(tmpData(LB))) + LB
    If index > LB And (index <= UB) Then
        Choose = tmpData(index)
    Else
        Choose = e_ValueError
    End If
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function Chr_(ByRef expression As String, ByRef fName As String) As String
    On Error GoTo err_handler
    Chr_ = ToLiteralString(ChrW(CLng(expression))) 'Expected numeric value
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

''' <summary>
''' Computes the mean for the given column from a 2D array
''' NOTE: PolyFit helper
''' </summary>
Private Function ColumnAvg(vArr() As Double, colIndex As Long) As Double
    Dim i As Long
    Dim n As Long
    Dim tmpResult As Double
    
    n = UBound(vArr) - LBound(vArr) + 1
    For i = 0 To n - 1
        tmpResult = tmpResult + vArr(i, colIndex) / n
    Next i
    ColumnAvg = tmpResult
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
    Dim SkipToken As Boolean
    
    BaseIndex = UBound(SubTreeData) + 1
    For b = LBound(EvalTree) To UBound(EvalTree)              'Loop all subexpressions
        OperationIndex = BaseIndex
        For t = 0 To EvalTree(b).index                         'Loop all tokens
            OperationIndex = OperationIndex + 1
            If P_GALLOPING_MODE Then
                SkipToken = EvalTree(b).Storage(t).ConstantToken And _
                            EvalTree(b).Storage(t).EvalResult <> vbNullString
            Else
                SkipToken = False
            End If
            If Not SkipToken Then 'Gallop
                GetOperands EvalTree(b).Storage(t), EvalTree(b), BaseIndex
                BottomLevelEval EvalTree(b).Storage(t)
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

Private Function Cosin(ByRef expression As String, ByRef fName As String) As Double
    Dim tmpEval As Double
    
    On Error GoTo err_handler
    tmpEval = CDbl(expression)
    If P_DEGREES Then
        tmpEval = tmpEval * PI / 180
    End If
    Cosin = Cos(tmpEval)
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
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

''' <summary>
''' Creates a curve string representation from solver coefficients
''' NOTE: PolyFit helper
''' </summary>
Private Function CurveString(ByRef SolverResult As Variant, cOption As Long) As String
    Select Case cOption
        Case 2 'Exponential [y = a*e^(b*x)]
            CurveString = CStr(Round(SolverResult(0), 4)) & op_mult & "e" & op_power & _
                            d_lParenthesis & Round(SolverResult(1), 4) & op_mult & "x" & d_rParenthesis
        Case 3 'Exponential [y = a*b^x]
            CurveString = CStr(Round(SolverResult(0), 4)) & op_mult & _
                            d_lParenthesis & Round(SolverResult(1), 4) & op_power & "x" & d_rParenthesis
        Case 4 'Power [y = a*x^b]
            CurveString = CStr(Round(SolverResult(0), 4)) & op_mult & "x" & op_power & Round(SolverResult(1), 4)
        Case 5 'Logarithmic [y = a*ln(x) + b]
            Dim sepChr As String
            If SolverResult(0) > 0 Then
                sepChr = op_plus
            End If
            CurveString = CStr(Round(SolverResult(1), 4)) & op_mult & "ln(x)" & sepChr & _
                                Round(SolverResult(0), 4)
    End Select
End Function

''' <summary>
''' Computes linearization for linear like functions
''' NOTE: CurveFit helper
''' </summary>
''' <param name="samplesArr">Data sample containing the observations pairs (x,y).</param>
''' <param name="cOption">Curve type ID.</param>
Private Function DataLinearization(ByRef samplesArr() As Double, cOption As Long) As Double()
    Dim i As Long
    Dim LB As Long
    Dim UB As Long
    Dim LB2 As Long
    Dim UB2 As Long
    Dim tmpResult() As Double
    
    LB = LBound(samplesArr)
    UB = UBound(samplesArr)
    LB2 = LBound(samplesArr, 2)
    UB2 = UBound(samplesArr, 2)
    ReDim tmpResult(LB To UB, LB2 To UB2)
    For i = LB To UB
        Select Case cOption
            Case 2, 3 'Exponential [y = a*e^(b*x) or y = a*b^x]
                tmpResult(i, 1) = Log10(samplesArr(i, 1))
                tmpResult(i, 0) = samplesArr(i, 0)
            Case 4 'Power [y = a*x^b]
                tmpResult(i, 1) = Log10(samplesArr(i, 1))
                tmpResult(i, 0) = Log10(samplesArr(i, 0))
            Case 5 'Logarithmic [y = a*ln(x) + b]
                tmpResult(i, 1) = samplesArr(i, 1)
                tmpResult(i, 0) = Log(samplesArr(i, 0)) 'LN
        End Select
    Next i
    DataLinearization = tmpResult
End Function

Private Function DateAdd_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    DateAdd_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function DateDiff_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    DateDiff_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function DatePart_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    DatePart_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function DateSerial_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    DateSerial_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function DateValue_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    DateValue_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function Date_(ByRef expression As String, ByRef fName As String) As String
    On Error GoTo err_handler
    Date_ = ToLiteralString(Date)
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function Day_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    Day_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function DDB_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    DDB_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
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
Private Function Determinant(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
            
            aArray() = ToDblArray(ArrayFromString(tmpData(LB)))
            LUdecomp aArray(), pivotingIdx, d
            For i = LBound(aArray) To UBound(aArray)
                d = d * aArray(i, i)
            Next i
            tmpEval = d
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    Determinant = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

''' <summary>
''' Finds for constant sub-expressions
''' </summary>
Private Sub DetermineKtokens()
    Dim BaseIndex As Long
    Dim b As Long, t As Long
    
    BaseIndex = UBound(SubTreeData) + 1
    For b = LBound(EvalTree) To UBound(EvalTree)              'Loop all subexpressions
        For t = 0 To EvalTree(b).index                         'Loop all tokens
            EvalTree(b).Storage(t).ConstantToken = IsConstantToken( _
                                                    EvalTree(b).Storage(t), EvalTree(b), BaseIndex _
                                                    )
        Next t
    Next b
End Sub

''' <summary>
''' Computes the value of the error function (i.e., the area under the
''' error function from 0 to x), given the limit of integration x. The
''' error function is also known as the Gauss error function.
''' [(c) John C. Pezzullo](https://statpages.info/scicalc.html)
''' </summary>
''' <param name="x">Limit of integration.</param>
Private Function ERF(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = ERF_( _
                            CDbl(tmpData(LB)) _
                            )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    ERF = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function ERF_(X As Double) As Double
    If (X < 0) Then
        ERF_ = (2 * (GAUSS_(Sqr(2) * X) + 0.5) - 1) 'Use CDF function over GAUSS value
    Else
        ERF_ = (1 - 2 * (GAUSS_(-Sqr(2) * X) + 0.5))
    End If
End Function

''' <summary>
''' Evaluate a math expression.
''' </summary>
''' <param name="ValuesToEvalIn">The values to use for eval the parse tree.</param>
'@DefaultMember
Public Function Eval(Optional ByRef ValuesToEvalWith As String = vbNullString) As String
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
                EvalFunction = Absolute(Argument, FunctionName)
            Case "ACHISQ"
                EvalFunction = ACHISQ(Argument, FunctionName)
            Case "AERF"
                EvalFunction = AERF(Argument, FunctionName)
            Case "AGAUSS"
                EvalFunction = AGAUSS(Argument, FunctionName)
            Case "AFISHF"
                EvalFunction = AFISHF(Argument, FunctionName)
            Case "ANORM"
                EvalFunction = ANORM(Argument, FunctionName)
            Case "ArcSin"
                EvalFunction = ArcSin(Argument, FunctionName)
            Case "ArcCos"
                EvalFunction = ArcCos(Argument, FunctionName)
            Case "ArcTan"
                EvalFunction = ArcTan(Argument, FunctionName)
            Case "ASTUDT"
                EvalFunction = ASTUDT(Argument, FunctionName)
            Case "strArray"
                EvalFunction = Array_(Argument, FunctionName)
            Case "ASCII"
                EvalFunction = Asc_(Argument, FunctionName)
            Case "Average"
                EvalFunction = average(Argument, FunctionName)
            Case "BETAINV"
                EvalFunction = BETAINV(Argument, FunctionName)
            Case "Cosin"
                EvalFunction = Cosin(Argument, FunctionName)
            Case "CHISQ"
                EvalFunction = CHISQ(Argument, FunctionName)
            Case "aChoose"
                EvalFunction = Choose(Argument, FunctionName)
            Case "ASCIIchr"
                EvalFunction = Chr_(Argument, FunctionName)
            Case "aCeiling"
                EvalFunction = aCeiling(Argument, FunctionName)
            Case "aDate"
                EvalFunction = Date_(Argument, FunctionName)
            Case "aDateAdd"
                EvalFunction = DateAdd_(Argument, FunctionName)
            Case "aDateDiff"
                EvalFunction = DateDiff_(Argument, FunctionName)
            Case "aDatePart"
                EvalFunction = DatePart_(Argument, FunctionName)
            Case "aDateSerial"
                EvalFunction = DateSerial_(Argument, FunctionName)
            Case "aDateValue"
                EvalFunction = DateValue_(Argument, FunctionName)
            Case "aDay"
                EvalFunction = Day_(Argument, FunctionName)
            Case "aDDB"
                EvalFunction = DDB_(Argument, FunctionName)
            Case "MatrixDeterminant"
                EvalFunction = Determinant(Argument, FunctionName)
            Case "ERF"
                EvalFunction = ERF(Argument, FunctionName)
            Case "ExpEuler"
                EvalFunction = ExpEuler(Argument, FunctionName)
            Case "FISHF"
                EvalFunction = FISHF(Argument, FunctionName)
            Case "CurveFit"
                EvalFunction = Fit(Argument, FunctionName)
            Case "aFloor"
                EvalFunction = aFloor(Argument, FunctionName)
            Case "aFormat"
                EvalFunction = Format_(Argument, FunctionName)
            Case "aFV"
                EvalFunction = FV_(Argument, FunctionName)
            Case "FunctionZero"
                EvalFunction = Zero(Argument, FunctionName)
            Case "Gamma"
                EvalFunction = tGamma(Argument, FunctionName)
            Case "GammaLN"
                EvalFunction = GAMMALN(Argument, FunctionName)
            Case "GAUSS"
                EvalFunction = GAUSS(Argument, FunctionName)
            Case "aHour"
                EvalFunction = Hour_(Argument, FunctionName)
            Case "iBETA"
                EvalFunction = iBETA(Argument, FunctionName)
            Case "aIff"
                EvalFunction = Iff_(Argument, FunctionName)
            Case "InverseMatrix"
                EvalFunction = InverseMatrix(Argument, FunctionName)
            Case "aIPMT"
                EvalFunction = IPMT_(Argument, FunctionName)
            Case "aIRR"
                EvalFunction = IRR_(Argument, FunctionName)
            Case "aLeft"
                EvalFunction = Left_(Argument, FunctionName)
            Case "aLen"
                EvalFunction = Len_(Argument, FunctionName)
            Case "LowerCase"
                EvalFunction = LCase_(Argument, FunctionName)
            Case "Logarithm"
                EvalFunction = Logarithm(Argument, FunctionName)
            Case "LgN"
                EvalFunction = LgN(Argument, FunctionName)
            Case "LUdecomposition"
                EvalFunction = LUdecomposition(Argument, FunctionName)
            Case "LUSolveLinearSystem"
                EvalFunction = LUsolve(Argument, FunctionName)
            Case "LN"
                EvalFunction = LN(Argument, FunctionName)
            Case "LSQRsolve"
                EvalFunction = LSQRsolve(Argument, FunctionName)
            Case "Max"
                EvalFunction = Max(Argument, FunctionName)
            Case "Middle"
                EvalFunction = Mid_(Argument, FunctionName)
            Case "Min"
                EvalFunction = Min(Argument, FunctionName)
            Case "aMinute"
                EvalFunction = Minute_(Argument, FunctionName)
            Case "aMIRR"
                EvalFunction = MIRR_(Argument, FunctionName)
            Case "MatrixMult"
                EvalFunction = MMult(Argument, FunctionName)
            Case "MatrixNegation"
                EvalFunction = MNEG(Argument, FunctionName)
            Case "MatrixRound"
                EvalFunction = MROUND(Argument, FunctionName)
            Case "MatrixSum"
                EvalFunction = MSum(Argument, FunctionName)
            Case "MatrixTranspose"
                EvalFunction = MTranspose(Argument, FunctionName)
            Case "aMonthName"
                EvalFunction = MonthName_(Argument, FunctionName)
            Case "aMonth"
                EvalFunction = Month_(Argument, FunctionName)
            Case "NORM"
                EvalFunction = NORM(Argument, FunctionName)
            Case "aNow"
                EvalFunction = Now_(Argument, FunctionName)
            Case "aNPER"
                EvalFunction = NPER_(Argument, FunctionName)
            Case "aNPV"
                EvalFunction = NPV_(Argument, FunctionName)
            Case "Percent"
                EvalFunction = Percent(Argument, FunctionName)
            Case "aPMT"
                EvalFunction = PMT_(Argument, FunctionName)
            Case "aPPMT"
                EvalFunction = PPMT_(Argument, FunctionName)
            Case "Power"
                EvalFunction = Power(Argument, FunctionName)
            Case "aPV"
                EvalFunction = PV_(Argument, FunctionName)
            Case "QRdecomposition"
                EvalFunction = QRdec(Argument, FunctionName)
            Case "aRATE"
                EvalFunction = RATE_(Argument, FunctionName)
            Case "REM"
                EvalFunction = aREM(Argument, FunctionName)
            Case "aReplace"
                EvalFunction = Replace_(Argument, FunctionName)
            Case "aRight"
                EvalFunction = Right_(Argument, FunctionName)
            Case "ROUND"
                EvalFunction = aROUND(Argument, FunctionName)
            Case "Sign"
                EvalFunction = Sign(Argument, FunctionName)
            Case "Sine"
                EvalFunction = Sine(Argument, FunctionName)
            Case "SolveLinearSystem"
                EvalFunction = solve(Argument, FunctionName)
            Case "aSLN"
                EvalFunction = SLN_(Argument, FunctionName)
            Case "SquareRoot"
                EvalFunction = SquareRoot(Argument, FunctionName)
            Case "Tangent"
                EvalFunction = Tangent(Argument, FunctionName)
            Case "STUDT"
                EvalFunction = STUDT(Argument, FunctionName)
            Case "aSwitch"
                EvalFunction = Switch_(Argument, FunctionName)
            Case "aSYD"
                EvalFunction = SYD_(Argument, FunctionName)
            Case "aTimeSerial"
                EvalFunction = TimeSerial_(Argument, FunctionName)
            Case "aTimeValue"
                EvalFunction = TimeValue_(Argument, FunctionName)
            Case "TINV"
                EvalFunction = TINV(Argument, FunctionName)
            Case "TINV_1T"
                EvalFunction = TINV_1T(Argument, FunctionName)
            Case "TINV_2T"
                EvalFunction = TINV_2T(Argument, FunctionName)
            Case "aTrim"
                EvalFunction = Trim_(Argument, FunctionName)
            Case "aUcase"
                EvalFunction = UCase_(Argument, FunctionName)
            Case "aWweekDay"
                EvalFunction = WeekDay_(Argument, FunctionName)
            Case "aWeekDayName"
                EvalFunction = WeekDayName_(Argument, FunctionName)
            Case "aYear"
                EvalFunction = Year_(Argument, FunctionName)
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

Private Function ExpEuler(ByRef expression As String, ByRef fName As String) As Double
    On Error GoTo err_handler
    ExpEuler = Exp(CDbl(expression))
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

''' <summary>
''' Computes factorial by recursion.
''' [(c) John C. Pezzullo](https://statpages.info/scicalc.html)
''' </summary>
''' <param name="n">Value to compute the factorial from.</param>
Private Function FACT(n As Double) As Double
    Dim R As Double
    Dim pA As Double, pB As Double
    Dim pC As Double, pD As Double
    
    If n = 0 Or n = 1 Then
        FACT = 1: Exit Function
    End If
    If n < 0 Then
        FACT = FACT(n + 1) / (n + 1): Exit Function
    End If
    If n > 1 Then
        FACT = n * FACT(n - 1): Exit Function
    End If
    If n < 0.5 Then
        R = n
    Else
        R = 1 - n
    End If
    pA = -0.009634403818022 + R * (0.007285315490429 + R * -0.001331461501875)
    pB = 0.1665389907228 + R * (-0.042197630554869 + R * pA)
    pC = -0.655878067489187 + R * (-0.042002698827786 + R * pB)
    pD = 0.577215664819072 + R * pC
    R = 1 / (1 + R * pD)
    If n > 0.5 Then
        R = (n * (1 - n) * PI) / (R * Sin(PI * n))
    End If
    FACT = R
End Function

Private Function Factorial(ByRef expression As String) As String
    On Error GoTo err_handler
    Factorial = FACT(CDbl(expression))
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & "Factorial" & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Sub FillTokenArgProps(ByRef TokenArg As Argument, Optional OnlySetIndexes As Boolean = False)
    If Not OnlySetIndexes Then
        TokenArg.NegationFlagOn = (InStrB(1, TokenArg.DefString, op_neg) = 1) '~*
        TokenArg.funcName = GetFunctionName(TokenArg.DefString)
        TokenArg.UDFFunctionIn = IsUDFFunction
        TokenArg.FunctionIn = (TokenArg.funcName <> vbNullString) Or TokenArg.UDFFunctionIn
        TokenArg.LinkedIndex = GetIndex(TokenArg.DefString)
        If Not TokenArg.NegationFlagOn Then
            TokenArg.LinkedVar = P_SCOPE.VarIndex(CastCase(TokenArg.DefString))
        Else
            TokenArg.LinkedVar = P_SCOPE.VarIndex(CastCase(MidB$(TokenArg.DefString, 3)))
        End If
        TokenArg.Implicit = (TokenArg.LinkedIndex >= 0)
        TokenArg.FactorialIn = (InStrB(1, TokenArg.DefString, op_Factorial) = LenB(TokenArg.DefString) - 1)
    Else
        TokenArg.LinkedIndex = -1
        TokenArg.LinkedVar = -1
    End If
End Sub

''' <summary>
''' Computes the probability value of an F-test, given the F-value,
''' numerator degrees of freedom, and denominator degrees of freedom.
''' [(c) John C. Pezzullo](https://statpages.info/scicalc.html)
''' </summary>
''' <param name="f">F-value.</param>
''' <param name="n1">Degrees of freedom 1.</param>
''' <param name="n2">Degrees of freedom 2.</param>
Private Function FISHF(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = FISHF_( _
                            CDbl(tmpData(LB)), _
                            CDbl(tmpData(LB + 1)), _
                            CDbl(tmpData(UB)) _
                            )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    FISHF = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function FISHF_(f As Double, n1 As Double, n2 As Double) As Double
    Dim X As Double, th As Double
    Dim A As Double, sth As Double
    Dim cth As Double, c As Double
    Dim k As Double
    
    X = n2 / (n1 * f + n2)
    If (REM_(n1, 2) = 0) Then
        FISHF_ = STATCOM(1 - X, n2, n1 + n2 - 4, n2 - 2) * (X ^ (n2 / 2))
        Exit Function
    End If
    If REM_(n2, 2) = 0 Then
        FISHF_ = 1 - STATCOM(X, n1, n1 + n2 - 4, n1 - 2) * ((1 - X) ^ (n1 / 2))
        Exit Function
    End If
    th = Atn(Sqr(n1 * f / n2))
    A = th / PID2: sth = Sin(th): cth = Cos(th)
    If (n2 > 1) Then
        A = A + sth * cth * STATCOM(cth * cth, 2, n2 - 3, -1) / PID2
    End If
    If (n1 = 1) Then FISHF_ = 1 - A: Exit Function
    c = 4 * STATCOM(sth * sth, n2 + 1, n1 + n2 - 4, n2 - 2) * sth * (cth ^ n2) / PI
    If (n2 = 1) Then FISHF_ = 1 - A + c / 2: Exit Function
    k = 2
    Do While (k <= (n2 - 1) / 2)
        c = c * k / (k - 0.5)
        k = k + 1
    Loop
    FISHF_ = 1 - A + c
End Function

''' <summary>
''' Performs a least squares regression curve fitting. The methods requires that the <expression>
''' argument be composed of the following:
'''                     1-) An array in text form containing the observations pairs.
'''                     2-) An integer value representing the fitting curve type:
'''                             1 --> Polynomial [y = a + b*x + c*x^2 +...+ m*x^n]
'''                             2 --> Exponential [y = a*e^(b*x)]
'''                             3 --> Exponential [y = a*b^x]
'''                             4 --> Power [y = a*x^b]
'''                             5 --> Logarithmic [y = a*ln(x) + b]
'''                     3-) An optional integer value representing the polynomial degree
'''                         only if Polynomial fitting option is selected
'''
''' The returned value is an like Java array with the form {{POLYNOMIAL};{R-squared VALUE}}
''' </summary>
Private Function Fit(ByRef expression As String, ByRef fName As String) As String
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    Dim argsCount As Long
    Dim ObservatiosnArray() As Double
    Dim fittingOption As Long
    Dim pDegree As Long
    Dim FittedPolynomial As String
    Dim solverCoeff As Variant
    
    On Error GoTo err_handler
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 2
            ObservatiosnArray() = ToDblArray(ArrayFromString(tmpData(LB)))
            fittingOption = CLng(tmpData(UB))
            If fittingOption = 1 Then 'Fit straight line [y = a + bx]
                pDegree = fittingOption
                FittedPolynomial = PolyFit_(ObservatiosnArray, pDegree)
                tmpEval = d_lCurly & _
                        d_lCurly & FittedPolynomial & d_rCurly & P_SEPARATORCHAR & _
                        d_lCurly & CStr(GoodnessOfFit_(FittedPolynomial, ObservatiosnArray)) & d_rCurly & _
                        d_rCurly
            Else
                solverCoeff = PolyFit_(DataLinearization(ObservatiosnArray, fittingOption), 1, False)
                Select Case fittingOption
                    Case 2 'Exponential [y = a*e^(b*x)]
                        solverCoeff(0) = 10 ^ solverCoeff(0) 'antilog_10 (A)
                        solverCoeff(1) = solverCoeff(1) / Log10(Exp(1)) 'B/log10(e)
                    Case 3 'Exponential [y = a*b^x]
                        solverCoeff(0) = 10 ^ solverCoeff(0) 'antilog_10 (A)
                        solverCoeff(1) = 10 ^ solverCoeff(1) 'antilog_10 (B)
                    Case 4 'Power [y = a*x^b]
                        solverCoeff(0) = 10 ^ solverCoeff(0) 'antilog_10 (A)
                End Select
                FittedPolynomial = CurveString(solverCoeff, fittingOption)
                tmpEval = d_lCurly & _
                            d_lCurly & FittedPolynomial & d_rCurly & P_SEPARATORCHAR & _
                            d_lCurly & CStr(GoodnessOfFit_(FittedPolynomial, ObservatiosnArray)) & d_rCurly & _
                            d_rCurly
            End If
        Case 3 'polynomial fitting
            fittingOption = CLng(tmpData(LB + 1))
            If fittingOption = 1 Then 'Polynomial fitting
                ObservatiosnArray() = ToDblArray(ArrayFromString(tmpData(LB)))
                pDegree = CLng(tmpData(UB))
                FittedPolynomial = PolyFit_(ObservatiosnArray, pDegree)
                tmpEval = d_lCurly & _
                        d_lCurly & FittedPolynomial & d_rCurly & P_SEPARATORCHAR & _
                        d_lCurly & CStr(GoodnessOfFit_(FittedPolynomial, ObservatiosnArray)) & d_rCurly & _
                        d_rCurly
            Else
                tmpEval = e_ValueError
            End If
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    Fit = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function Floor(ByRef value As Double) As Double
    Dim tmpResult As Double
    
    tmpResult = Fix(value)
    Floor = tmpResult + ((value <> tmpResult) And (value < 0))
End Function

Private Function FormatEntry(expression As String) As String
    FormatEntry = ReplaceImpliedMult( _
                                Replace(Replace(Replace(RemoveDupNegation(ApplyLawOfSigns _
                                (ReconstructLiteralStrings(CStr(expression), _
                                Join$(Split(expression, d_Space), vbNullString)))), "()", "('')"), "{{", "({{"), "}}", "}})") _
                                )
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

Private Function Format_(ByRef expression As String, ByRef fName As String) As String
     Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    Format_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
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
Public Function fRoots(ByRef aFunction As String, A As Double, _
                        b As Double, Optional inc As Double = 1, _
                        Optional epsilon As Double = 0.0000000001) As String
    Dim tmpRoot As String
    Dim tmpResult As String
    Dim lx As Double, ux As Double
    
    If inc <> 1 Then
        inc = Abs(inc)
    End If
    If A > b Then
        fRoots = e_ValueError
        Exit Function
    End If
    tmpRoot = fZero(aFunction, A, b, epsilon, False)
    If tmpRoot <> e_ValueError Then      'Function with at least one root
        lx = A
        ux = A + inc
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

Private Function FV_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    FV_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Public Function fZero(ByRef aFunction As String, ByVal A As Double, _
                        ByVal b As Double, Optional epsilon As Double = 0.0000000001, _
                        Optional includeVarNames As Boolean = True, _
                        Optional useBisectionMethod As Boolean = False, _
                        Optional printDebugINFO As Boolean = False) As String
    If Not useBisectionMethod Then
        fZero = fZeroMRF(aFunction, A, b, epsilon, includeVarNames, printDebugINFO)
    Else
        fZero = fZeroMBM(aFunction, A, b, epsilon, includeVarNames, printDebugINFO)
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
Private Function fZeroMBM(ByRef aFunction As String, ByVal A As Double, _
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
    If A > b Then Exit Function
    Set fEvalHelper = New CSVexpressions
    aZero = 10 * epsilon
    With fEvalHelper
        .Create aFunction
        tmpVar() = Split(.CurrentVariables, "; ")
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
        .VarValue(tmpVar(varIdx)) = A
        fa = CDbl(.Eval): evalCounter = evalCounter + 1
        .VarValue(tmpVar(varIdx)) = b
        fb = CDbl(.Eval): evalCounter = evalCounter + 1
        Do
            'Divide the interval into three segments, assumes a < b
            segmentLen = (b - A) / 3
            c = A + segmentLen
            d = b - segmentLen
            'Evaluate the function at c, d
            .VarValue(tmpVar(varIdx)) = c
            fc = CDbl(.Eval): evalCounter = evalCounter + 1
            .VarValue(tmpVar(varIdx)) = d
            fd = CDbl(.Eval): evalCounter = evalCounter + 1
            If fc = 0 Or fd = 0 Then
                Exit Do
            End If
            If fa * fc < 0 Then 'Root in interval a-c
                b = c
                fb = fc
            ElseIf fc * fd < 0 Then
                A = c
                fa = fc
                b = d
                fb = fd
            ElseIf fd * fb < 0 Then
                A = d
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
        .VarValue(tmpVar(varIdx)) = tmpResult
        If Round(tmpFzeroEval, Len(CStr(1 / aZero)) - 1) = 0 Then
            If includeVarNames Then
                fZeroMBM = .CurrentVarValues
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
Private Function fZeroMRF(ByRef aFunction As String, ByVal A As Double, _
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
    Dim k As Double, CK As Double
    Dim tmpVar() As String
    Dim toleranceFlag As Boolean
    Dim varCounter As Long
    Dim varIdx As Long
    Dim varLB As Long
    
    On Error GoTo fZeroMRF_errHandler
    If A > b Then Exit Function
    Set fEvalHelper = New CSVexpressions
    aZero = 10 * epsilon
    With fEvalHelper
        .Create aFunction
        tmpVar() = Split(.CurrentVariables, "; ")
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
            .VarValue(tmpVar(varIdx)) = A
            fa = CDbl(.Eval): evalCounter = evalCounter + 1
            .VarValue(tmpVar(varIdx)) = b
            fb = CDbl(.Eval): evalCounter = evalCounter + 1
            'Compute approximation point of the root
            c = (A * fb - b * fa) / (fb - fa)
            'Evaluate the function at approximation
            .VarValue(tmpVar(varIdx)) = c
            fc = CDbl(.Eval): evalCounter = evalCounter + 1
            If fa * fc < 0 Then 'Root in interval a-c
                'Projection
                CK = Abs(2 * fc / (b - c))
                d = ((1 + CK) * A * fb - b * fa) / ((1 + CK) * fb - fa)
                'Evaluate the function at projection
                .VarValue(tmpVar(varIdx)) = d
                fd = CDbl(.Eval): evalCounter = evalCounter + 1
                If fd * fa < 0 Then
                    b = d
                Else
                    b = c
                    A = d
                End If
            ElseIf fa * fc > 0 Then
                CK = Abs(0.5 * fc / (b - c))
                d = ((1 + CK) * b * fa - A * fb) / ((1 + CK) * fa - fb)
                .VarValue(tmpVar(varIdx)) = d
                fd = CDbl(.Eval): evalCounter = evalCounter + 1
                If fd * fa > 0 Then
                    A = d
                Else
                    A = c
                    b = d
                End If
            End If
            toleranceFlag = (Abs(fd) > epsilon)
            k = k + 1
        Loop While toleranceFlag And k < 100
        If Round(fd, Len(CStr(1 / aZero)) - 1) = 0 Then
            If includeVarNames Then
                fZeroMRF = .CurrentVarValues
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

''' <summary>
''' Computes the value GAMMA of x from factorial function.
''' [(c) John C. Pezzullo](https://statpages.info/scicalc.html)
''' </summary>
''' <param name="x">Value to compute the GAMMA function from.</param>
Private Function GAMMA(X As Double) As Variant
    GAMMA = FACT(X - 1)
End Function

Private Function GAMMALN(ByRef expression As String, ByRef fName As String) As String
    On Error GoTo err_handler
    GAMMALN = GAMMALN_(CDbl(expression))
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

''' <summary>
''' Returns the natural logarithm of the gamma function.
''' [(c) John C. Pezzullo](https://statpages.info/scicalc.html)
''' </summary>
''' <param name="x">Value to compute the GAMMALN function from.</param>
Private Function GAMMALN_(X As Double) As Double
    Dim j As Long
    Dim cof As Variant
    Dim ser As Double
    Dim xx As Double, y As Double
    Dim tmp As Double

    cof = [{76.18009172947146, -86.50532032941677, 24.01409824083091, -1.231739572450155, 0.1208650973866179E-2, -0.5395239384953E-5}]
    ser = 1.00000000019001
    y = X: xx = X
    tmp = X + 5.5
    tmp = tmp - (xx + 0.5) * Log(tmp)
    For j = LBound(cof) To UBound(cof)
        y = y + 1
        ser = ser + cof(j) / (y)
    Next j
    GAMMALN_ = Log(2.506628274631 * ser / xx) - tmp
End Function

''' <summary>
''' Computes the probability that a member of a standard normal population
''' will fall between the mean and z standard deviations from the mean.
''' [(c) John C. Pezzullo](https://statpages.info/scicalc.html) with adjustment specified in
''' https://support.microsoft.com/en-us/office/gauss-function-069f1b4e-7dee-4d6a-a71f-4b69044a6b33
''' The result is 0.5 off from the cumulative distribution function (CDF) for the standard normal
''' given the upper limit of integration z. So CDF (mean:=0, sd:=1, z) = GAUSS(z) + 0.5
''' </summary>
''' <param name="z">Standard deviations from the mean.</param>
Private Function GAUSS(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = GAUSS_( _
                            CDbl(tmpData(LB)) _
                            )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    GAUSS = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function GAUSS_(z As Double) As Double
    If (z < 0) Then
        If (z < -10) Then
            GAUSS_ = 0
        Else
            GAUSS_ = CHISQ_(z * z, 1) / 2
        End If
    Else
        If (z > 10) Then
            GAUSS_ = 1
        Else
            GAUSS_ = 1 - CHISQ_(z * z, 1) / 2
        End If
    End If
    GAUSS_ = GAUSS_ - 0.5
End Function

Private Function get2DimArrayCopy(ByRef BArray() As Double) As Double()
    Dim tmpResult() As Double
    Dim isArray2D As Boolean
    Dim colDim As Long
    Dim i As Long, j As Long
    Dim LB As Long, UB As Long
    Dim LB2 As Long, UB2 As Long
    
    isArray2D = Is2Darray(BArray)
    UB = UBound(BArray)
    LB = LBound(BArray)
    If isArray2D Then
        UB2 = UBound(BArray, 2)
        LB2 = LBound(BArray, 2)
        colDim = UB2 - LB2
    End If
    ReDim tmpResult(0 To UB - LB, 0 To colDim)
    If isArray2D Then
        For i = LB To UB
            For j = LB2 To UB2
                tmpResult(0 + i - LB, 0 + j - LB2) = BArray(i, j)
            Next j
        Next i
    Else
        For i = LB To UB
            tmpResult(0 + i - LB, 0) = BArray(i)
        Next i
    End If
    get2DimArrayCopy = tmpResult
End Function

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
                            ByRef JaggedArray As Boolean, ByRef IdxDim1 As Long, _
                            ByRef IdxDim2 As Long) As String
    If MultiDimArr Then
        GetArrItm = CStr(arr(IdxDim1, IdxDim2))
    Else
        If JaggedArray Then
            GetArrItm = CStr(arr(IdxDim1)(IdxDim2))
        Else
            GetArrItm = CStr(arr(IdxDim2))
        End If
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

Private Function GetEvalToken(ByRef expression As String) As token
    Dim TokenStart As Long
    Dim TokenEnd As Long
    Dim TokenDet As TokenInfo
    Dim tmpArgs() As String
    
    TokenDet = GetTokenInfo(expression)
    If TokenDet.Position > 2 Then
        '@--------------------------------------------------------------------
        ' Delimit token
        GetTokenStart expression, TokenDet.Position, TokenStart
        GetTokenEnd expression, TokenDet.Position, TokenDet.OperatorLen, TokenEnd
        GetEvalToken.DefString = MidB$(expression, TokenStart, TokenEnd - TokenStart + 2)
        GetEvalToken.OperationToken = TokenDet.OperationToken
        SplitToken GetEvalToken.DefString, tmpArgs, GetEvalToken.OperationToken
        GetEvalToken.Logical = TokenDet.LogicalToken
        '@--------------------------------------------------------------------
        ' Define arguments
        GetEvalToken.Arg1.DefString = tmpArgs(LBound(tmpArgs))
        GetEvalToken.Arg2.DefString = tmpArgs(UBound(tmpArgs))
        '@--------------------------------------------------------------------
        ' Fill arguments data
        FillTokenArgProps GetEvalToken.Arg1
        FillTokenArgProps GetEvalToken.Arg2
    Else            'Retun values
        GetEvalToken.DefString = expression
        GetEvalToken.OperationToken = OperatorToken.otNull
        GetEvalToken.Logical = TokenDet.LogicalToken
        '@--------------------------------------------------------------------
        ' Define argument
        GetEvalToken.Arg1.DefString = expression
        '@--------------------------------------------------------------------
        ' Fill argument data
        FillTokenArgProps GetEvalToken.Arg1
        FillTokenArgProps GetEvalToken.Arg2, True
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
                CurArg.Operand = Factorial(P_SCOPE.VarValue(CurArg.LinkedVar))
            Else
                CurArg.Operand = P_SCOPE.VarValue(CurArg.LinkedVar)
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
    Dim prevChar As String
    
    Do
        SubExpression = ApplyLawOfSigns(SubExpression)
        vToken = GetEvalToken(SubExpression)
        '@--------------------------------------------------------------------
        ' Mask worked token
        tmpPos = InStrB(1, SubExpression, vToken.DefString)
        If tmpPos > 2 Then
            prevChar = MidB$(SubExpression, tmpPos - 2, 2)
            OperandInBundle = (InStrB(1, op_AllItems, prevChar))
            Do While Not OperandInBundle And tmpPos > 2    'Tokens starts with a operator or with a null string
                tmpPos = InStrB(tmpPos + 2, SubExpression, vToken.DefString)
                prevChar = MidB$(SubExpression, tmpPos - 2, 2)
                OperandInBundle = (InStrB(1, op_AllItems, prevChar))
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
    Dim cAcsw As Long
    
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
        cAcsw = AscW(curChar)
        Select Case cAcsw
            Case 45, 126
                Dim prevChar As String
                prevChar = MidB$(expression, outLng - 2, 2)
                If InStrB(1, op_AllNotUnaryItems, prevChar) Then
                    If cAcsw = 45 And (prevChar = op_plus Or prevChar = op_minus) Then
                        outLng = outLng - 2
                    Else
                        If cAcsw = 126 And prevChar = op_neg Then
                            outLng = outLng - 2
                        End If
                    End If
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

''' <summary>
''' Computes the googness of a curve fit over The ObserArray must be a 2 columns
''' array in which the first column contains the X (independent variable) values
''' and the second column contains the Y (dependent variable) values
''' </summary>
Private Function GoodnessOfFit_(ByRef Polynomial As String, ByRef ObserArray() As Double) As Double
    Dim exprHelper As CSVexpressions
    Dim averageY As Double
    Dim Rsquared As Double
    Dim seSquare As Double
    Dim stSquare As Double
    Dim i As Long
    Dim n As Long
    
    n = UBound(ObserArray) - LBound(ObserArray) + 1
    Set exprHelper = New CSVexpressions
    With exprHelper
        .Create Polynomial
        averageY = ColumnAvg(ObserArray, 1) 'Mean
        For i = 0 To n - 1
            .Eval "x=" & CStr(ObserArray(i, 0))
            seSquare = seSquare + (ObserArray(i, 1) - CDbl(.result)) ^ 2 'Partial sum squared regression
            stSquare = stSquare + (averageY - ObserArray(i, 1)) ^ 2 'Partial total sum of squares
        Next i
        Rsquared = Round(1 - seSquare / stSquare, 4)
    End With
    GoodnessOfFit_ = Rsquared
    Set exprHelper = Nothing
End Function

Private Function Hour_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    Hour_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

''' <summary>
''' Computes the cumulative distribution function (CDF) for the beta distribution
''' (i.e., the area under the beta distribution from 0 to x), given values of
''' the shape parameters, and the point at which to evaluate the function.
''' [(c) John C. Pezzullo](https://statpages.info/scicalc.html)
''' </summary>
''' <param name="x">Integration limit point.</param>
''' <param name="a">Shape parameter alpha.</param>
''' <param name="b">Shape parameter beta.</param>
Private Function iBETA(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = iBETA_( _
                            CDbl(tmpData(LB)), _
                            CDbl(tmpData(LB + 1)), _
                            CDbl(tmpData(UB)))
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    iBETA = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Public Function iBETAINV(p As Double, A As Double, b As Double) As Double
    Const EPS As Double = 0.00000001
    Dim aL As Double
    Dim a1, b1 As Double
    Dim j As Long
    Dim lna As Double, lnb As Double
    Dim pp As Double, t As Double
    Dim u As Double, err As Double
    Dim X As Double, H As Double
    Dim w As Double, afac As Double

    a1 = A - 1: b1 = b - 1
    j = 0
    If p <= 0 Then
        iBETAINV = 0
        Exit Function
    End If
    If p >= 1 Then
        iBETAINV = 1
        Exit Function
    End If
    If A >= 1 And b >= 1 Then
        If p < 0.5 Then
            pp = p
        Else
            pp = 1 - p
        End If
        t = Sqr(-2 * Log(pp))
        X = (2.30753 + t * 0.27061) / (1 + t * (0.99229 + t * 0.04481)) - t
        If p < 0.5 Then X = -X
        aL = (X * X - 3) / 6
        H = 2 / (1 / (2 * A - 1) + 1 / (2 * b - 1))
        w = (X * Sqr(aL + H) / H) - (1 / (2 * b - 1) - 1 / (2 * A - 1)) * _
            (aL + 5 / 6 - 2 / (3 * H))
        X = A / (A + b * Exp(2 * w))
    Else
        lna = Log(A / (A + b)): lnb = Log(b / (A + b))
        t = Exp(A * lna) / A: u = Exp(b * lnb) / b
        w = t + u
        If p < t / w Then
            X = (A * w * p) ^ (1 / A)
        Else
            X = 1 - ((b * w * (1 - p)) ^ (1 / b))
        End If
    End If
    afac = -GAMMALN_(A) - GAMMALN_(b) + GAMMALN_(A + b)
    Do While j < 10
        If X = 0 Or X = 1 Then iBETAINV = X: Exit Function
        err = iBETA_(X, A, b) - p
        t = Exp(a1 * Log(X) + b1 * Log(1 - X) + afac)
        u = err / t
        t = u / (1 - 0.5 * Min(1, u * (a1 / X - b1 / (1 - X))))
        X = X - t
        If X <= 0 Then X = 0.5 * (X + t)
        If X >= 1 Then X = 0.5 * (X + t + 1)
        If Abs(t) < EPS * X And j > 0 Then Exit Do
        j = j + 1
    Loop
    iBETAINV = X
End Function

Private Function iBETA_(X As Double, A As Double, b As Double) As Variant
    Dim bt As Double

'   // Factors in front of the continued fraction.
    If X = 0 Or X = 1 Then
        bt = 0
    Else
        bt = Exp(GAMMALN_(A + b) - GAMMALN_(A) - GAMMALN_(b) + A * Log(X) + b * Log(1 - X))
    End If
    If X < 0 Or X > 1 Then
        iBETA_ = False: Exit Function
    Else
        If X < (A + 1) / (A + b + 2) Then
'           // Use continued fraction directly.
            iBETA_ = bt * betacf(X, A, b) / A
        Else
'           // else use continued fraction after making the symmetry transformation.
            iBETA_ = 1 - bt * betacf(1 - X, b, A) / b
        End If
    End If
End Function

Private Function Iff_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As Boolean
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
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
    aBuffer.index = -1
End Sub

Private Sub InitCBbuffer(ByRef aBuffer As ClusterBuffer)
    aBuffer.Capacity = 128
    ReDim aBuffer.Storage(0 To aBuffer.Capacity - 1)
    aBuffer.index = -1
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

Private Sub InitializeErrHandler()
    P_ERRORDESC = vbNullString
    P_ERRTYPE = errNone
End Sub

Private Function Inverse(ByRef A() As Double, ByRef indx() As Long, _
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
    
    aUB = UBound(A)
    aLB = LBound(A)
    n = aUB - aLB + 1
    ReDim y(1 To n, 1 To n)
    ReDim col(1 To n)
    LUdecomp A, indx, d
    For j = 1 To n
        For i = 1 To n
            col(i) = 0
        Next i
        col(j) = 1
        LUbkSub A, indx, col
        For i = 1 To n
            If Abs(col(i)) = 0 Then col(i) = 0      'Avoid VBA -0 value
            y(i, j) = col(i)
        Next i
    Next j
    Inverse = y
End Function

Private Function InverseMatrix(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            Dim aArray() As Double
            Dim pivotingIdx() As Long
            Dim d As Double             'Interchages parity
            
            aArray() = ToDblArray(ArrayFromString(tmpData(LB)))
            tmpEval = ArrayToString(Inverse(aArray, pivotingIdx, d))
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    InverseMatrix = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function IPMT_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    IPMT_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function IRR_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    Dim strArray() As String
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    IRR_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
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

''' <summary>
''' Determines if a token/sub-expressions is constant
''' </summary>
Private Function IsConstantToken(ByRef CurToken As token, ByRef CurTree As ClusterTree, _
                                    ByRef BaseIndex As Long) As Boolean
    Dim VariablesInToken As Boolean
    Dim VariablesInRelatedToken1 As Boolean
    Dim VariablesInRelatedToken2 As Boolean
    Dim iCounter As Long
    Dim TokensCount As Long
    
    VariablesInToken = CurToken.Arg1.LinkedVar <> -1 Or CurToken.Arg2.LinkedVar <> -1
    If CurToken.Arg1.Implicit Or CurToken.Arg2.Implicit Then 'Check related token
        If CurToken.Arg1.Implicit Then
            If CurToken.Arg1.LinkedIndex >= BaseIndex Then 'Data on current tree
                VariablesInRelatedToken1 = Not IsConstantToken(CurTree.Storage(CurToken.Arg1.LinkedIndex - BaseIndex), _
                                                            CurTree, BaseIndex)
            Else 'Data is main tree
                VariablesInRelatedToken1 = Not IsConstantToken(EvalTree(CurToken.Arg1.LinkedIndex).Storage(0), _
                                                            EvalTree(CurToken.Arg1.LinkedIndex), BaseIndex)
                iCounter = 1 'Check all tokens in current cluster
                TokensCount = EvalTree(CurToken.Arg1.LinkedIndex).index
                Do While VariablesInRelatedToken1 And iCounter <= TokensCount
                    VariablesInRelatedToken1 = Not IsConstantToken(EvalTree(CurToken.Arg1.LinkedIndex).Storage(iCounter), _
                                                            EvalTree(CurToken.Arg1.LinkedIndex), BaseIndex)
                    iCounter = iCounter + 1
                Loop
            End If
        End If
        If CurToken.Arg2.Implicit Then
            If CurToken.Arg2.LinkedIndex >= BaseIndex Then
                VariablesInRelatedToken2 = Not IsConstantToken(CurTree.Storage(CurToken.Arg2.LinkedIndex - BaseIndex), _
                                                            CurTree, BaseIndex)
            Else
                VariablesInRelatedToken2 = Not IsConstantToken(EvalTree(CurToken.Arg2.LinkedIndex).Storage(0), _
                                                            EvalTree(CurToken.Arg2.LinkedIndex), BaseIndex)
                iCounter = 1
                TokensCount = EvalTree(CurToken.Arg2.LinkedIndex).index
                Do While VariablesInRelatedToken2 And iCounter <= TokensCount
                    VariablesInRelatedToken2 = Not IsConstantToken(EvalTree(CurToken.Arg2.LinkedIndex).Storage(iCounter), _
                                                            EvalTree(CurToken.Arg2.LinkedIndex), BaseIndex)
                    iCounter = iCounter + 1
                Loop
            End If
        End If
        IsConstantToken = Not VariablesInToken And (Not VariablesInRelatedToken1 And Not VariablesInRelatedToken2)
    Else
        IsConstantToken = Not VariablesInToken
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

Private Function IsTwiceJaggedArr(ByRef arr As Variant, elIdx As Long) As Boolean
    On Error GoTo err_handler
    IsTwiceJaggedArr = IsArray(arr(elIdx)(LBound(arr(elIdx))))
    Exit Function
err_handler:
    IsTwiceJaggedArr = False
    err.Clear
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

Private Function LCase_(ByRef expression As String, ByRef fName As String) As String
    On Error GoTo err_handler
    LCase_ = ToLiteralString(LCase(expression))
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function Left_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    Left_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function Len_(ByRef expression As String, ByRef fName As String) As Long
    On Error GoTo err_handler
    If IsLiteralString(expression) Then
        Len_ = Len(FormatLiteralString(expression))
    Else
        Len_ = Len(expression)
    End If
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function LgN(ByRef expression As String, ByRef fName As String) As Double
    Dim tmpData() As String
    Dim tmpEval As Double
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpEval = 0
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    If UB - LB = 1 Then
        tmpEval = Log(CDbl(tmpData(LB))) _
                    / Log(CDbl(tmpData(UB)))    'Log(x)/Log(N)
    Else
        tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End If
    LgN = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function LN(ByRef expression As String, ByRef fName As String) As Double
    On Error GoTo err_handler
    LN = Log(CDbl(expression))
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function Log10(aValue As Double) As Double
    Log10 = Log(aValue) / Log(10)
End Function

Private Function Logarithm(ByRef expression As String, ByRef fName As String) As Double
    On Error GoTo err_handler
    Logarithm = Log(CDbl(expression)) / Log(10)
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

''' <summary: Least squares solution of A*X = B>
''' Returns X that minimizes the two norm of Q*R*X-B
''' @exception  IllegalArgumentException  Matrix row dimensions must agree.
''' @exception  RuntimeException  Matrix is rank deficient.
''' </summary>
''' <param name="A">Matrix of coefficients of the system variables.</param>
''' <param name="b">Right side vector.</param>
Private Function LSQRsolve(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 2
            tmpEval = ArrayToString(QR_LSQsolve( _
                            ToDblArray(ArrayFromString(tmpData(LB))), _
                            ToDblArray(ArrayFromString(tmpData(UB))) _
                            ))
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    LSQRsolve = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

''' <summary>
''' Solves the set of n linear equations A.X = B. Here
''' a[1,...,n][1,...,n] is input, not as the matrix A but
''' rather as its LU decomposition, determined by the routine
''' LUdecomp. indx[1,...,n] is input as the permutation vector
''' returned from LUdecomp. b[1,...,n] is input as the right-hand
''' side vector B, and returns with the solution vector X
''' </summary>
Private Sub LUbkSub(ByRef A() As Double, ByRef indx() As Long, ByRef b() As Double)
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
    
    aUB = UBound(A)
    aLB = LBound(A)
    bLB = LBound(b)
    n = aUB - aLB + 1
    ii = 0
    For i = 1 To n
        ip = indx(i)
        sum = b(ip + bLB - 1)
        b(ip + bLB - 1) = b(i + bLB - 1)
        If ii Then
            For j = ii To i - 1
                sum = sum - A(i + aLB - 1, j + aLB - 1) * b(j + bLB - 1)
            Next j
        ElseIf sum Then
            ii = i
        End If
        b(i + bLB - 1) = sum
    Next i
    For i = n To 1 Step -1
        sum = b(i + bLB - 1)
        For j = i + 1 To n
            sum = sum - A(i + aLB - 1, j + aLB - 1) * b(j + bLB - 1)
        Next j
        b(i + bLB - 1) = sum / A(i + aLB - 1, i + aLB - 1)
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
Private Sub LUdecomp(ByRef A() As Double, ByRef indx() As Long, _
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
    
    aUB = UBound(A)
    aLB = LBound(A)
    n = aUB - aLB + 1
    ReDim vv(1 To n)
    ReDim indx(1 To n)
    d = 1                       'No row interchanges yet
    For i = 1 To n              'Loop rows to get the implicit scaling information
        big = 0
        For j = 1 To n
            temp = Abs(A(i + aLB - 1, j + aLB - 1))
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
            sum = A(i + aLB - 1, j + aLB - 1)
            For k = 1 To i - 1
                sum = sum - A(i + aLB - 1, k + aLB - 1) * A(k + aLB - 1, j + aLB - 1)
            Next k
            A(i + aLB - 1, j + aLB - 1) = sum
        Next i
        big = 0                  'Initialize for search the largest pivot element
        For i = j To n
            sum = A(i + aLB - 1, j + aLB - 1)
            For k = 1 To j - 1
                sum = sum - A(i + aLB - 1, k + aLB - 1) * A(k + aLB - 1, j + aLB - 1)
            Next k
            A(i + aLB - 1, j + aLB - 1) = sum
            dum = vv(i) * Abs(sum)
            If dum >= big Then
                big = dum
                imax = i
            End If
        Next i
        If j <> imax Then         'Need rows interchange?
            For k = 1 To n
                dum = A(imax + aLB - 1, k + aLB - 1)
                A(imax + aLB - 1, k + aLB - 1) = A(j + aLB - 1, k + aLB - 1)
                A(j + aLB - 1, k + aLB - 1) = dum
            Next k
            d = -1 * d           'Change parity
            vv(imax) = vv(j)     'Change scale factor
        End If
        indx(j) = imax
        If A(j + aLB - 1, j + aLB - 1) = 0 Then    'Avoid division by zero in singular matrix
            A(j + aLB - 1, j + aLB - 1) = Tiny
        End If
        If j <> n Then           'Divide by the pivot element
            dum = 1 / A(j + aLB - 1, j + aLB - 1)
            For i = j + 1 To n
                A(i + aLB - 1, j + aLB - 1) = A(i + aLB - 1, j + aLB - 1) * dum
            Next i
        End If
    Next j
    Erase vv
End Sub

''' <summary>
''' Return the LU decomposition for the given matrix
''' </summary>
Private Function LUdecomposition(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            Dim aArray() As Double
            Dim pivotingIdx() As Long
            Dim d As Double             'Interchages parity
            
            aArray() = ToDblArray(ArrayFromString(tmpData(LB)))
            LUdecomp aArray(), pivotingIdx, d
            tmpEval = ArrayToString(aArray)
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    LUdecomposition = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
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
Private Function LUsolve(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
            Dim Xarray() As String
            Dim BArray() As Double
            Dim pivotingIdx() As Long
            Dim d As Double             'Interchages parity
            Dim tmpElement As String
            
            aArray() = ToDblArray(ArrayFromString(tmpData(LB)))
            BArray() = ToDblArray(ArrayFromString(tmpData(LB + 2)))
            mLB = LBound(aArray)
            mUB = UBound(aArray)
            
            eqCount = mUB - mLB + 1
            LUdecomp aArray, pivotingIdx, d             'LU decomposition
            LUbkSub aArray, pivotingIdx, BArray         'Back substitution, returning in bArray
            If argsCount = 4 Then
                Xarray() = ArrayFromString(tmpData(LB + 1))
                includeNames = CBool(tmpData(UB))
            End If
            For iCounter = mLB To mUB
                If iCounter > mLB Then
                    tmpEval = tmpEval & P_SEPARATORCHAR & d_Space
                End If
                If includeNames Then
                    tmpElement = FormatLiteralString(Xarray(iCounter)) _
                                & d_Space & op_equal & d_Space _
                                & BArray(iCounter)
                Else
                    tmpElement = BArray(iCounter)
                End If
                tmpEval = tmpEval & tmpElement
            Next iCounter
            If Not includeNames Then
                tmpEval = d_lCurly & d_lCurly & tmpEval & d_rCurly & d_rCurly
            End If
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    LUsolve = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function Max(ByRef expression As String, ByRef fName As String) As Double
    Dim g As Long
    Dim tmpData() As String
    Dim tmpEval As Double
    Dim compEval As Double
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    tmpEval = CDbl(tmpData(LBound(tmpData)))
    For g = LBound(tmpData) + 1 To UBound(tmpData)
        compEval = CDbl(tmpData(g))
        If compEval > tmpEval Then
            tmpEval = compEval
        End If
    Next g
    Max = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function MAX_(ParamArray aList() As Variant) As Double
    Dim i As Long
    Dim tmpResult As Double
    Dim tmpEl As Double
    
    i = LBound(aList)
    tmpResult = CDbl(aList(i))
    For i = i + 1 To UBound(aList)
        tmpEl = CDbl(aList(i))
        If tmpEl > tmpResult Then tmpResult = tmpEl
    Next i
    MAX_ = tmpResult
End Function

Private Function method1(t As Double, R As Double) As Double
    Dim e As Double, H As Double
    Dim f As Double, i As Double
    Dim inp1 As Double, inp2 As Double
    Dim inp3 As Double, inp4 As Double
    Dim inp5 As Double, inp6 As Double
    Dim inp7 As Double, inp9 As Double
    
    If (0 >= R Or R >= 1) Then
        method1 = 0: Exit Function
    End If
    If (0.5 > R) Then
        method1 = -method1(t, 1 - R): Exit Function
    End If
    e = method3(R)
    inp1 = e ^ 2
    inp2 = (27 * inp1 + 339) * inp1 + 930
    inp2 = inp2 * inp1 - 1782
    inp2 = ((inp2 * inp1 - 765) * inp1 + 17955) / 368640
    inp3 = (79 * inp1 + 776) * inp1 + 1482
    inp3 = inp3 * inp1 - 1920
    inp3 = (inp3 * inp1 - 945) / 92160
    inp4 = (((3 * inp1 + 19) * inp1 + 17) * inp1 - 15) / 384
    inp6 = ((5 * inp1 + 16) * inp1 + 3) / 96
    inp5 = (inp1 + 1) / 4
    inp7 = inp3 + inp2 / t
    inp7 = (inp6 + (inp4 + inp7 / t) / t)
    inp7 = e * (1 + (inp5 + inp7 / t) / t)
    If (t <= (Log10(R) ^ 2) + 3) Then
        Do
            f = method11(t, inp7)
            inp9 = t + 1
            i = Log(inp9 / (t + inp7 * inp7))
            i = i + Log(t / inp9 / 2 / PI) - 1 + (1 / inp9 - 1 / t)
            i = (f - R) / Exp((inp9 * i / 6) / 2)
            inp7 = inp7 + i
            H = method6(i, Abs(Fix(Log10(Abs(inp7)) - 6)))
        Loop While (inp7 And 0 <> H)
    End If
    method1 = inp7
End Function

Private Function method11(y1 As Double, y2 As Double) As Double
    Dim e As Double, rt1 As Double
    Dim l As Double, n As Double
    Dim s As Double, rt2 As Double
    
    l = Atn((y2 / Sqr(y1)) / 1)
    n = Cos(l) ^ 2: s = 1
    For rt2 = y1 - 2 To 2 Step -2
        s = 1 + (rt2 - 1) / rt2 * n * s
    Next rt2
    If REM_(y1, 2) = 0 Then
        e = 0.5
        rt1 = Sin(l) / 2
        method11 = rt1
    Else
        e = 0.5 + l / PI
        If 1 = y1 Then
            rt1 = 0
        Else
            rt1 = Cos(l) * Sin(l) / PI
        End If
        method11 = MAX_(0, 1 - e - rt1 * s)
    End If
End Function

Private Function method12(p2 As Double) As Double
    If p2 <> 0 Then
        method12 = method6(p2, method22(p2))
    Else
        method12 = 0
    End If
End Function

Private Function method2(x1 As Double, x2 As Double) As Variant
    If ((0 >= x1 Or Abs(x1) - Abs(Fix(x1)) <> 0) Or (0 >= x2 Or x2 >= 1)) Then
        method2 = False
    Else
        method2 = method12(method1(x1, x2))
    End If
End Function

Private Function method22(p5 As Double) As Double
    method22 = Abs(Fix(Log10(Abs(p5)) - 7))
End Function

Private Function method3(p3 As Double) As Double
    Dim p4 As Double
    Dim e As Double
    
    p4 = -Log(4 * p3 * (1 - p3))
    e = -3.231081277E-09 + p4 * (3.657763036E-11 + 6.936233982E-13 * p4)
    e = -0.00000104527497 + p4 * (8.360937017E-08 + p4 * e)
    e = 0.000006841218299 + p4 * (0.000005824238515 + p4 * e)
    e = -0.0008364353589 + p4 * (-0.0002250947176 + p4 * e)
    e = Sqr(p4 * (1.570796288 + p4 * (0.03706987906 + p4 * e)))
    If p3 > 0.5 Then
        e = -e
    End If
    method3 = e
End Function

Private Function method6(t As Double, R As Double) As Double
    Dim tmpResult As Double

    tmpResult = t * 10 ^ R
    method6 = Round(tmpResult) / (10 ^ R)
End Function

Private Function Mid_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    Mid_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function Min(ByRef expression As String, ByRef fName As String) As Double
    Dim g As Long
    Dim tmpData() As String
    Dim tmpEval As Double
    Dim compEval As Double
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    tmpEval = CDbl(tmpData(LBound(tmpData)))
    For g = LBound(tmpData) + 1 To UBound(tmpData)
        compEval = CDbl(tmpData(g))
        If compEval < tmpEval Then
            tmpEval = compEval
        End If
    Next g
    Min = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
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

Private Function Minute_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    Minute_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function MIN_(ParamArray aList() As Variant) As Double
    Dim i As Long
    Dim tmpResult As Double
    Dim tmpEl As Double
    
    i = LBound(aList)
    tmpResult = CDbl(aList(i))
    For i = i + 1 To UBound(aList)
        tmpEl = CDbl(aList(i))
        If tmpEl < tmpResult Then tmpResult = tmpEl
    Next i
    MIN_ = tmpResult
End Function

Private Function MIRR_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    Dim idx As Long
    Dim strArray() As String
    Dim dblArray() As Double
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    MIRR_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

''' <summary>
''' Multiplies 2 matrices A and B. Matrix B can be a column vector or
''' an scalar number.
''' </summary>
Private Function MMult(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    Dim tmpOP As Variant
    Dim bIsMatrix As Boolean
    
    On Error GoTo err_handler
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 2
            bIsMatrix = InStrB(1, tmpData(UB), "{{")
            If bIsMatrix Then
                tmpOP = MMULT_( _
                                ToDblArray(ArrayFromString(tmpData(LB))), _
                                ToDblArray(ArrayFromString(tmpData(UB))))
            Else
                tmpOP = MMULT_( _
                                ToDblArray(ArrayFromString(tmpData(LB))), _
                                CDbl(tmpData(UB)))
            End If
            If IsArray(tmpOP) Then
                tmpEval = ArrayToString(tmpOP)
            Else
                tmpEval = CStr(tmpOP)
            End If
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    MMult = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function MMULT_(ByRef A() As Double, ByRef b As Variant) As Variant
    Dim ai As Long, ai2 As Long
    Dim aIsNxMmatrix As Boolean
    Dim aIsVector As Boolean
    Dim aLB As Long, aLB2 As Long
    Dim aUB As Long, aUB2 As Long
    Dim bi As Long, bi2 As Long
    Dim bIsNxMmatrix As Boolean
    Dim bIsVector As Boolean
    Dim bLB As Long, bLB2 As Long
    Dim bUB As Long, bUB2 As Long
    Dim c() As Double
    Dim i As Long, ii As Long
    Dim j As Long, jj As Long
    Dim m As Long 'Columns of the output
    Dim n As Long 'Rows of the output
    Dim OutputIsArray As Boolean
    Dim sum As Double
    
    aIsVector = Not Is2Darray(A) And IsArray(A)
    bIsVector = Not Is2Darray(b) And IsArray(b)
    OutputIsArray = True
    If aIsVector And bIsVector Then     'Compute dot product
        aLB2 = LBound(A): aUB2 = UBound(A)
        bLB2 = LBound(b): bUB2 = UBound(b)
        n = aUB2 - aLB2 + 1
        m = bUB2 - bLB2 + 1
        If m = n Then   'Require equal columns count
            For i = 0 To n - 1 'Loop a() rows
                sum = sum + A(aLB2 + i) * b(bLB2 + i)
            Next i
            OutputIsArray = False
        End If
    ElseIf (aIsVector And Not IsArray(b)) Or _
            (Not aIsVector And Not IsArray(b)) Then '1D or 2D array x scalar
        If IsArray(A) Then
            If aIsVector Then
                aLB = LBound(A): aUB = UBound(A)
                m = aUB - aLB + 1
                ReDim c(0 To m - 1)
                For i = 0 To m - 1
                    c(i) = A(aLB + i) * b
                Next i
            Else
                aLB = LBound(A): aUB = UBound(A)
                aLB2 = LBound(A, 2): aUB2 = UBound(A, 2)
                n = aUB - aLB + 1
                m = aUB2 - aLB2 + 1
                ReDim c(0 To n - 1, 0 To m - 1)
                For i = 0 To n - 1
                    For j = 0 To m - 1
                        c(i, j) = A(aLB + i, aLB2 + j) * b
                    Next j
                Next i
            End If
        End If
    ElseIf aIsVector And IsArray(b) Then '1D array x 2D array
        aLB = 0: aUB = 0
        aLB2 = LBound(A): aUB2 = UBound(A)
        bLB = LBound(b): bUB = UBound(b)
        bLB2 = LBound(b, 2): bUB2 = UBound(b, 2)
        n = aUB - aLB + 1
        m = bUB2 - bLB2 + 1
        If Not (aUB2 - aLB2 + 1) <> (bUB - bLB + 1) Then 'Requires columns from a() equals to rows from b()
            ReDim c(0 To n - 1, 0 To m - 1)
            For i = 0 To n - 1 'Loop a() rows
                For j = 0 To m - 1 'Loop c() columns
                    sum = 0
                    For jj = 0 To aUB2 - aLB2 'Loop a() columns
                        sum = sum + A(aLB + jj) * b(bLB + jj, bLB2 + j)
                    Next jj
                    c(i, j) = sum
                Next j
            Next i
        End If
    ElseIf Not aIsVector And Not bIsVector Then '2D array x 2D array
        aLB = LBound(A): aUB = UBound(A)
        aLB2 = LBound(A, 2): aUB2 = UBound(A, 2)
        bLB = LBound(b): bUB = UBound(b)
        bLB2 = LBound(b, 2): bUB2 = UBound(b, 2)
        n = aUB - aLB + 1
        m = bUB2 - bLB2 + 1
        If Not (aUB2 - aLB2 + 1) <> (bUB - bLB + 1) Then 'Requires columns from a() equals to rows from b()
            ReDim c(0 To n - 1, 0 To m - 1)
            For i = 0 To n - 1 'Loop c() rows
                For j = 0 To m - 1 'Loop c() columns
                    sum = 0
                    For jj = 0 To aUB2 - aLB2 'Loop a() columns
                        sum = sum + A(aLB + i, aLB2 + jj) * b(bLB + jj, bLB2 + j)
                    Next jj
                    c(i, j) = sum
                Next j
            Next i
        End If
    ElseIf Not aIsVector And bIsVector Then '2D array x 1D array
        aLB = LBound(A): aUB = UBound(A)
        aLB2 = LBound(A, 2): aUB2 = UBound(A, 2)
        bLB = 0: bUB = 0
        bLB2 = LBound(b): bUB2 = UBound(b)
        n = aUB - aLB + 1
        m = bUB - bLB + 1
        If Not (aUB2 - aLB2 + 1) <> (bUB - bLB + 1) Then 'Requires columns from a() equals to rows from b()
            ReDim c(0 To n - 1, 0 To m - 1)
            For i = 0 To n - 1 'Loop c() rows
                For j = 0 To m - 1 'Loop c() columns
                    sum = 0
                    For jj = 0 To aUB2 - aLB2 'Loop a() columns
                        sum = sum + A(aLB + i, aLB2 + jj) * b(bLB2 + i)
                    Next jj
                    c(i, j) = sum
                Next j
            Next i
        End If
    End If
    If OutputIsArray Then
        If InitializedArray(c) Then
            MMULT_ = c
        Else
            MMULT_ = e_ValueError
        End If
    Else
        MMULT_ = sum
    End If
End Function

''' <summary>
''' Returns the negation of the given matrix.
''' </summary>
''' <param name="aMatrix">The matrix to negate.</param>
Private Function MNEG(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = ArrayToString(MNEG_( _
                            ToDblArray(ArrayFromString(tmpData(LB))) _
                            ))
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    MNEG = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function MNEG_(ByRef aMatrix() As Double) As Double()
    Dim mLB As Long
    Dim mLB2 As Long
    Dim mUB As Long
    Dim mUB2 As Long
    Dim i As Long, j As Long
    Dim ii As Long, jj As Long
    Dim tmpResult() As Double
    
    If Not IsArray(aMatrix) Then Exit Function
    mLB = LBound(aMatrix)
    mUB = UBound(aMatrix)
    If Is2Darray(aMatrix) Then
        mLB2 = LBound(aMatrix, 2)
        mUB2 = UBound(aMatrix, 2)
        ReDim tmpResult(0 To mUB - mLB, 0 To mUB2 - mLB2)
        For i = mLB To mUB
            jj = 0
            For j = mLB2 To mUB2
                If Abs(aMatrix(i, j)) Then tmpResult(ii, jj) = -1 * aMatrix(i, j)
                jj = jj + 1
            Next j
            ii = ii + 1
        Next i
    Else
        ReDim tmpResult(0 To mUB - mLB)
        For i = mLB To mUB
            If Abs(aMatrix(i, j)) Then tmpResult(ii) = -1 * aMatrix(i)
            ii = ii + 1
        Next
    End If
    MNEG_ = tmpResult
End Function

Private Function MonthName_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    MonthName_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function Month_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    Month_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function MROUND(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = ArrayToString(MROUND_( _
                            ToDblArray(ArrayFromString(tmpData(LB))) _
                            ))
        Case 2
            tmpEval = ArrayToString(MROUND_( _
                            ToDblArray(ArrayFromString(tmpData(LB))), _
                            CDbl(tmpData(UB)) _
                            ))
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    MROUND = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function MROUND_(ByRef aArray() As Double, Optional nDigits As Long = 0) As Double()
    Dim i As Long, j As Long
    Dim LB1 As Long, LB2 As Long
    Dim UB1 As Long, UB2 As Long
    Dim tmpResult() As Double
    
    If Is2Darray(aArray) Then
        UB1 = UBound(aArray): LB1 = LBound(aArray)
        UB2 = UBound(aArray, 2): LB2 = LBound(aArray, 2)
        ReDim tmpResult(0 To UB1 - LB1, 0 To UB2 - LB2)
        For i = LB1 To UB1
            For j = LB2 To UB2
                tmpResult(i - LB1, j - LB2) = Round(aArray(i, j), nDigits)
            Next j
        Next i
    Else
        UB1 = UBound(aArray): LB1 = LBound(aArray)
        ReDim tmpResult(0 To UB1 - LB1)
        For i = LB1 To UB1
            tmpResult(i - LB1) = Round(aArray(i), nDigits)
        Next i
    End If
    MROUND_ = tmpResult
End Function

''' <summary>
''' Sum 2 matrices A and B when A and B has same number of
''' rows and columns. With the last argument set to True
''' A - B is computed instead of A + B
''' </summary>
Private Function MSum(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 2
            tmpEval = ArrayToString(MSum_( _
                                ToDblArray(ArrayFromString(tmpData(LB))), _
                                ToDblArray(ArrayFromString(tmpData(UB)))) _
                                )
        Case 3
            tmpEval = ArrayToString(MSum_( _
                                ToDblArray(ArrayFromString(tmpData(LB))), _
                                ToDblArray(ArrayFromString(tmpData(LB + 1))), _
                                CBool(tmpData(UB))))
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    MSum = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function MSum_(ByRef A() As Double, ByRef b() As Double, _
                        Optional Difference As Boolean = False) As Double()
    Dim c() As Double
    Dim LBa As Long
    Dim LBa2 As Long
    Dim UBb As Long
    Dim LBb As Long
    Dim LBb2 As Long
    Dim m As Long
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim sum As Double
    
    LBa = LBound(A)
    LBa2 = LBound(A, 2)

    n = UBound(A) - LBa + 1
    m = UBound(A, 2) - LBa2 + 1
    LBb = LBound(b)
    LBb2 = LBound(b, 2)
    UBb = UBound(b)
    If UBb - LBb + 1 <> n Then Exit Function    'Required equal numbers of rows
    If UBound(b, 2) - LBb2 + 1 <> m Then Exit Function     'Required equal numbers of columns
    ReDim c(0 To n - 1, 0 To m - 1)
    For i = 0 To n - 1
        For j = 0 To m - 1
            If Not Difference Then
                sum = A(i + LBa, j + LBa2) + b(i + LBb, j + LBb2)
            Else
                sum = A(i + LBa, j + LBa2) - b(i + LBb, j + LBb2)
            End If
            c(i, j) = sum
        Next j
    Next i
    MSum_ = c
End Function

''' <summary>
''' Compute an array transposition
''' </summary>
Private Function MTranspose(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = ArrayToString(MTranspose_( _
                                ToDblArray(ArrayFromString(tmpData(LB)))))
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    MTranspose = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function MTranspose_(ByRef A() As Double) As Double()
    Dim c() As Double
    Dim LBa As Long
    Dim LBa2 As Long
    Dim m As Long
    Dim n As Long
    Dim i As Long
    Dim j As Long
    
    LBa = LBound(A)
    LBa2 = LBound(A, 2)

    n = UBound(A) - LBa + 1
    m = UBound(A, 2) - LBa2 + 1
    ReDim c(0 To m - 1, 0 To n - 1)
    For i = 0 To n - 1
        For j = 0 To m - 1
            c(j, i) = A(LBa + i, LBa2 + j)
        Next j
    Next i
    MTranspose_ = c
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

''' <summary>
''' Computes the two-tailed probability value forstandard normal curve,
''' given a Z-score (i.e., the two-tailed probability from +/- the
''' Z-score to infinity on both tails of the distribution).
''' [(c) John C. Pezzullo](https://statpages.info/scicalc.html)
''' </summary>
''' <param name="z">Z-score.</param>
Private Function NORM(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = NORM_( _
                            CDbl(tmpData(LB)) _
                            )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    NORM = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function NORM_(z As Double) As Double
    Dim Q As Double
    Q = z * z
    If (Abs(z) > 7) Then
        NORM_ = (1 - 1 / Q + 3 / (Q * Q)) * Exp(-Q / 2) / (Abs(z) * Sqr(PID2))
    Else
        NORM_ = CHISQ_(Q, 1)
    End If
End Function

Private Function NotIsPI(ByRef VarName As String) As Boolean
    NotIsPI = Not VarName Like "[Pp][Ii]"
End Function

Private Function Now_(ByRef expression As String, ByRef fName As String) As String
    On Error GoTo err_handler
    Now_ = ToLiteralString(Now)
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function NPER_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    NPER_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function NPV_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As Double
    Dim iRate As Double
    Dim pCounter As Long
    Dim LB As Long, UB As Long
    Dim cLooper As Long
    Dim strArray() As String
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    NPV_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
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
        DetermineKtokens 'Detect constants tokens
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

Private Function Percent(ByRef expression As String, ByRef fName As String) As Double
    On Error GoTo err_handler
    Percent = CDbl(expression) / 100
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function PMT_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    PMT_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

''' <summary>
''' Performs a regression polynomial fitting in which the relationship between the
''' independent variable x and the dependent variable y is modelled as an nth
''' degree polynomial in x in the form y = a + b*x + c*x^2 +...+ m*x^n, where
''' the calculated parameters and the number of equations will be n+1.
''' Furthermore, the power summations for the independent variable will be 2n,
''' while n multiplications are required for the dependent variable.
'''
''' In summary, a table with n + 1 rows and 3*n + 1 columns will be needed to perform
''' the calculations. The products will be allocated aferter 2*n + 1 indexes.
'''
''' </summary>
''' <param name="samplesArr">Data sample containing the observations pairs (x,y).</param>
''' <param name="pDegree">Polynomial degree.</param>
''' <param name="stringOutput">When 'False' the coeficents array is returned.</param>
Private Function PolyFit_(ByRef samplesArr() As Double, pDegree As Long, _
                            Optional stringOutput As Boolean = True) As Variant
    Dim accumulatorArray() As Double
    Dim observCount As Long
    Dim iCounter As Long
    Dim jCounter As Long
    Dim mLB As Long
    Dim mUB As Long
    Dim Xarray() As Double
    Dim strXArray As String
    Dim BArray() As Double
    Dim strBarray As String
    Dim cnArray() As String
    Dim SolverExpr As String
    Dim SolverResult() As String
    Dim FittedPolynomial As String
    
    mLB = LBound(samplesArr)
    mUB = UBound(samplesArr)
    observCount = mUB - mLB + 1
    '|x|y|x^2|x^3|...|x^2n|y*x|y*x^2|...|y*x^n|
    ReDim accumulatorArray(0 To 3 * pDegree) '3*n + 1 columns
    For iCounter = 0 To observCount - 1
        For jCounter = 0 To UBound(accumulatorArray)
            If jCounter <= 1 Then 'Accumulate observations
                accumulatorArray(jCounter) = accumulatorArray(jCounter) _
                                                        + samplesArr(iCounter, jCounter)
            Else
                If jCounter > 1 And jCounter <= 2 * pDegree Then 'Accumulate powers
                    accumulatorArray(jCounter) = accumulatorArray(jCounter) + samplesArr(iCounter, 0) ^ (jCounter) 'SUM(x ^ i)
                Else 'Accumulate products
                    accumulatorArray(jCounter) = accumulatorArray(jCounter) + samplesArr(iCounter, 1) * _
                                                samplesArr(iCounter, 0) ^ (jCounter - 2 * pDegree) 'SUM((x ^ i) * y)
                End If
            End If
        Next jCounter
    Next iCounter
    'Create parameters array
    ReDim Xarray(0 To pDegree, 0 To pDegree)
    For iCounter = 0 To pDegree
        For jCounter = 0 To pDegree
            If iCounter + jCounter = 0 Then
                Xarray(iCounter, jCounter) = observCount
            Else
                If iCounter + jCounter = 1 Then
                    Xarray(iCounter, jCounter) = accumulatorArray(0) 'Sum x
                Else
                    Xarray(iCounter, jCounter) = accumulatorArray(iCounter + jCounter) 'Sum x^i
                End If
            End If
        Next jCounter
    Next iCounter
    'Create B vector
    ReDim BArray(0, 0 To pDegree)
    For iCounter = 0 To pDegree  '2*n + 1 < i < 3*n
        If iCounter = 0 Then 'Sum y
            BArray(0, iCounter) = accumulatorArray(1)
        Else
            BArray(0, iCounter) = accumulatorArray(iCounter + 2 * pDegree)
        End If
    Next iCounter
    'Create vector of coefficient names
    ReDim cnArray(0, 0 To pDegree)
    For iCounter = 0 To pDegree
        cnArray(0, iCounter) = ToLiteralString(ChrW(97 + iCounter))
    Next iCounter
    'Get linear eq. system
    strXArray = ArrayToString(Xarray)
    strBarray = ArrayToString(BArray)
    SolverExpr = strXArray & P_SEPARATORCHAR & _
                ArrayToString(cnArray) & P_SEPARATORCHAR & _
                strBarray & P_SEPARATORCHAR & "False"
    'Solve
    SolverResult = ArrayFromString(LUsolve(SolverExpr, "LUsolve"))
    If stringOutput Then
        'Format
        FittedPolynomial = PolyString(SolverResult, pDegree)
        PolyFit_ = FittedPolynomial
    Else
        PolyFit_ = SolverResult
    End If
End Function

''' <summary>
''' Creates a polynomial string representation from solver coefficients
''' NOTE: PolyFit helper
''' </summary>
Private Function PolyString(ByRef SolverResult() As String, ByRef pDegree As Long) As String
    Dim i As Long
    Dim tmpResult As String
    
    For i = 0 To pDegree
        If i = 0 Then
            tmpResult = CStr(Round(CDbl(SolverResult(i)), 4))
        Else
            If CDbl(SolverResult(i)) > 0 Then
                tmpResult = tmpResult & d_Space & op_plus _
                            & d_Space & Round(CDbl(SolverResult(i)), 4)
            Else
                tmpResult = tmpResult & d_Space & Round(CDbl(SolverResult(i)), 4)
            End If
            If i > 1 Then
                tmpResult = tmpResult & "*x^" & i
            Else
                tmpResult = tmpResult & "*x"
            End If
        End If
    Next i
    PolyString = tmpResult
End Function

Private Function Power(ByRef expression As String, ByRef fName As String) As Double
    Dim tmpData() As String
    Dim tmpEval As Double
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    If UB - LB = 1 Then
        tmpEval = CDbl(tmpData(LB)) ^ _
                        CDbl(tmpData(UB))
    Else
        tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End If
    Power = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function PPMT_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    PPMT_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function PV_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    PV_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

''' <summary QR Decomposition>
''' For an m-by-n matrix A with m >= n, the QR decomposition is an m-by-n
''' orthogonal matrix Q and an n-by-n upper triangular matrix R so that
''' A = Q*R.
'''
''' The QR decompostion always exists, even if the matrix does not have
''' full rank, so the constructor will never fail.  The primary use of the
''' QR decomposition is in the least squares solution of nonsquare systems
''' of simultaneous linear equations.  This will fail if isFullRank()
''' returns false.
'''<NOTE>
''' The original source code was modified to allow the computation of R
''' by enforce signs of the diagonal elements to be positive.
''' https://github.com/fiji/Jama/blob/master/src/main/java/Jama/QRDecomposition.java
''' https://math.nist.gov/javanumerics/jama/
''' </summary>
''' <param name="A">Matrix to decompose.</param>
''' <param name="PositiveDiag">When "True" enforce signs of the diagonal elements to be positive.</param>
Private Function QRdec(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = SplitArgs(expression)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1, 2
            If argsCount = 1 Then
                QR_ ArrayFromString(tmpData(LB))
            Else
                QR_ ArrayFromString(tmpData(LB)), CBool(FormatLiteralString(tmpData(UB), True))
            End If
            tmpEval = d_lCurly _
                        & ArrayToString(QR_getQ) & P_SEPARATORCHAR & _
                        ArrayToString(QR_getR) & _
                    d_rCurly
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    QRdec = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#

'# QR DECOMPOSITION
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub QR_(ByRef A As Variant, Optional PositiveDiag As Boolean = False)
    Dim i As Long, j As Long
    Dim k As Long
    Dim s As Double
    Dim nrm As Double
    
    m = UBound(A) - LBound(A) + 1
    n = UBound(A, 2) - LBound(A, 2) + 1
    QR() = ToDblArray(A)
    ReDim rDiag(0 To n - 1)
    ReDim signFlip(0 To n - 1)
    For k = 0 To n - 1
        '// Compute 2-norm of k-th column without under/overflow.
        nrm = 0
        For i = k To m - 1
            nrm = QR_hypot(nrm, QR(i, k))
        Next i
        If nrm <> 0 Then
            '// Form k-th Householder vector.
            If (QR(k, k) < 0) Then
                nrm = -nrm
            End If
            For i = k To m - 1
                QR(i, k) = QR(i, k) / nrm
            Next i
            QR(k, k) = QR(k, k) + 1#
            '// Apply transformation to remaining columns.
            For j = k + 1 To n - 1
                s = 0#
                For i = k To m - 1
                    s = s + QR(i, k) * QR(i, j)
                Next i
                s = -s / QR(k, k)
                For i = k To m - 1
                    QR(i, j) = QR(i, j) + s * QR(i, k)
                Next i
            Next j
        End If
        rDiag(k) = -nrm
        signFlip(k) = (rDiag(k) < 0) And PositiveDiag
    Next k
End Sub

'   /** Return the Householder vectors
'   @return     Lower trapezoidal matrix whose columns define the reflections
'   */
Private Function QR_getH() As Double()
    Dim H() As Double
    Dim i As Long, j As Long
    
    ReDim H(0 To m - 1, 0 To n - 1)
    For i = 0 To m - 1
        For j = 0 To n - 1
            If i >= j Then
                H(i, j) = QR(i, j)
            Else
                H(i, j) = 0#
            End If
        Next j
    Next i
    QR_getH = H
End Function

'   /** Generate and return the (economy-sized) orthogonal factor
'   @return     Q
'   */
Private Function QR_getQ() As Double()
    Dim Q() As Double
    Dim i As Long, j As Long
    Dim k As Long
    Dim s As Double
    
    ReDim Q(0 To m - 1, 0 To n - 1)
    For k = m - 1 To 0 Step -1
        Q(k, k) = 1#
        For j = k To n - 1
            If (QR(k, k) <> 0) Then
                s = 0#
                For i = k To m - 1
                    s = s + QR(i, k) * Q(i, j)
                Next i
                s = -s / QR(k, k)
                For i = k To m - 1
                    Q(i, j) = Q(i, j) + s * QR(i, k)
                Next i
            End If
        Next j
    Next k
    For i = 0 To m - 1
        For j = 0 To n - 1
            If Abs(Q(i, j)) <> 0 Then
                If signFlip(j) Then
                    Q(i, j) = -1 * Q(i, j)
                End If
            Else
                 Q(i, j) = 0
            End If
        Next j
    Next i
    QR_getQ = Q
End Function

'   /** Return the upper triangular factor
'   @return     R
'   */
Private Function QR_getR() As Double()
    Dim R() As Double
    Dim i As Long, j As Long
    
    ReDim R(0 To n - 1, 0 To n - 1)
    For i = 0 To n - 1
        For j = 0 To n - 1
            If i < j Then
                R(i, j) = QR(i, j)
            ElseIf i = j Then
                R(i, j) = rDiag(i)
            Else
                R(i, j) = 0#
            End If
            If signFlip(i) And (R(i, j) <> 0) Then R(i, j) = -1 * R(i, j)
        Next j
    Next i
    QR_getR = R
End Function

   
Private Function QR_hypot(ParamArray args() As Variant) As Double
    Dim i As Long
    Dim sqrSum As Double
    
    For i = LBound(args) To UBound(args)
        sqrSum = sqrSum + (args(i) * args(i))
    Next i
    QR_hypot = Sqr(sqrSum)
End Function

'   /** Is the matrix full rank?
'   @return     true if R, and hence A, has full rank.
'   */
Private Function QR_isFullRank() As Boolean
    Dim j As Long
    Dim tmpResult As Boolean
    
    Do While j < n And Not tmpResult
        tmpResult = (rDiag(j) = 0)
        j = j + 1
    Loop
    QR_isFullRank = Not tmpResult
End Function

Private Function QR_LSQsolve(ByRef A() As Double, b() As Double) As Double()
    QR_ A
    QR_LSQsolve = QR_solve(b)
End Function

Private Function QR_solve(ByRef b() As Double) As Double()
    If UBound(b) - LBound(b) + 1 <> m Then
        Exit Function '"Matrix row dimensions must agree."
    End If
    If Not QR_isFullRank Then
        Exit Function '"Matrix is rank deficient."
    End If
    Dim X() As Double, nx As Long
    Dim i As Long, j As Long, k As Long
    Dim s As Double, tmpResult() As Double
    
'    // Copy right hand side
    X() = get2DimArrayCopy(b)
    nx = UBound(X, 2) - LBound(X, 2) + 1
'    // Compute Y = transpose(Q)*B
    For k = 0 To n - 1
        For j = 0 To nx - 1
            s = 0
            For i = k To m - 1
                s = s + QR(i, k) * X(i, j)
            Next i
            s = -s / QR(k, k)
            For i = k To m - 1
                X(i, j) = X(i, j) + s * QR(i, k)
            Next i
        Next j
    Next k
'    // Solve R*X = Y
    For k = n - 1 To 0 Step -1
        For j = 0 To nx - 1
            X(k, j) = X(k, j) / rDiag(k)
        Next j
        For i = 0 To k - 1
            For j = 0 To nx - 1
                X(i, j) = X(i, j) - X(k, j) * QR(i, k)
            Next j
        Next i
    Next k
    ReDim tmpResult(0 To n - 1, 0 To nx - 1)
    For i = 0 To n - 1
        For j = 0 To nx - 1
            tmpResult(i, j) = X(i, j)
        Next j
    Next i
    QR_solve = tmpResult
End Function

Private Function RATE_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    RATE_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
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

Private Function REM_(n As Double, d As Double)
    Dim Q As Double
    
    Q = Fix(n / d)
    REM_ = n - d * Q
End Function

Private Function ReplaceImpliedMult(expression As String) As String
    Dim tmpPos As Long
    Dim tmpStr As String
    Dim LookupPos As Long
    Dim tmpVar As String
    Dim tmpVarInitPos As Long
    Dim prevChar As String
    Dim reservedChar As Boolean
    
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
                prevChar = MidB$(tmpStr, tmpPos - 2, 2)
                reservedChar = (prevChar = d_lParenthesis Or InStrB(1, op_AllItems, prevChar))
                If Not reservedChar Then
                    If GetFunctionName(LCase$(tmpVar)) = vbNullString Then  'Implied multiplication found
                        tmpStr = MidB(tmpStr, 1, tmpPos - 1) & op_mult & MidB(tmpStr, tmpPos)
                    End If
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

Private Function Replace_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    Replace_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
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

Private Function Right_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 1
            tmpEval = ToLiteralString(right(FormatLiteralString(tmpData(LB), True), 1))
        Case 2
            tmpEval = ToLiteralString( _
                                        right( _
                                                FormatLiteralString(tmpData(LB), True), _
                                                CLng(FormatLiteralString(tmpData(UB), True)) _
                                                ) _
                                    )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    Right_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function ROUND_(aNumber As Double, Optional aDigits As Double = 0) As Double
    ROUND_ = Round(aNumber, aDigits)
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

Private Function Sign(ByRef expression As String, ByRef fName As String) As Double
    On Error GoTo err_handler
    Sign = (Sgn(CDbl(expression)))
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Sub SignSubstitution(ByRef expression As String, FindStr As String, Replacement As String)
    If InStrB(1, expression, FindStr) Then
        expression = Replace(expression, FindStr, Replacement)
    End If
End Sub

Private Function Sine(ByRef expression As String, ByRef fName As String) As Double
    Dim tmpEval As Double
    
    On Error GoTo err_handler
    tmpEval = CDbl(expression)
    If P_DEGREES Then
        tmpEval = tmpEval * PI / 180
    End If
    Sine = Sin(tmpEval)
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function SLN_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    SLN_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
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
Private Function solve(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    Dim tolerance As Double
    
    On Error GoTo err_handler
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
            Dim Xarray() As Double
            
            tolerance = 0.0000000001
            significantDigits = Len(CStr(1 / tolerance)) - 1
            cArray() = ToDblArray(ArrayFromString(tmpData(LB)))
            rArray() = ToDblArray(ArrayFromString(tmpData(LB + 2)))
            mLB = LBound(cArray)
            mUB = UBound(cArray)
            
            ReDim Xarray(mLB To mUB)
            eqCount = mUB - mLB + 1
            SORiteration eqCount, cArray, rArray, Xarray, 500, tolerance, 1
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
                                & Round(Xarray(iCounter), significantDigits)
                Else
                    tmpElement = Round(Xarray(iCounter), significantDigits)
                End If
                tmpEval = tmpEval & tmpElement
            Next iCounter
            If Not includeNames Then
                tmpEval = d_lCurly & d_lCurly & tmpEval & d_rCurly & d_rCurly
            End If
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    solve = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Sub SORiteration(n As Long, ByRef A() As Double, ByRef b() As Double, _
                         ByRef X() As Double, iter As Long, tol As Double, omega As Double)
''' Hoffman, J. D. (2001). Numerical methods for engineers and scientists (2nd ed., rev.expanded). Marcel Dekker.
    Dim IT As Long
    Dim i As Long, j As Long
    Dim dxmax As Double
    Dim residual As Double
    Dim colLB As Long, rowLB As Long
    
    colLB = LBound(A, 2)
    rowLB = LBound(A)
    For IT = 1 To iter
        dxmax = 0
        For i = 1 To n
            residual = b(i + rowLB - 1)
            For j = 1 To n
                residual = residual - A(i + rowLB - 1, j + colLB - 1) * X(j + rowLB - 1)
            Next j
            If Abs(residual) > dxmax Then dxmax = Abs(residual)
            X(i + rowLB - 1) = X(i + rowLB - 1) + omega * residual / A(i + rowLB - 1, i + colLB - 1)
        Next i
        If dxmax < tol Then Exit For
    Next IT
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

Private Function SquareRoot(ByRef expression As String, ByRef fName As String) As Double
    On Error GoTo err_handler
    SquareRoot = Sqr(CDbl(expression))
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function STATCOM(Q As Double, i As Double, _
                        j As Double, b As Double) As Double
    Dim zz As Double
    Dim z As Double
    Dim k As Double
    
    zz = 1: z = zz: k = i
    Do While (k <= j)
        zz = zz * Q * k / (k - b)
        z = z + zz: k = k + 2
    Loop
    STATCOM = z
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

''' <summary>
''' Computes the two-tailed probability values of a t-test, given the
''' t-value and the degrees of freedom. The one-tailed probabilities
''' can be computed as:
'''                      One-tailed probability (right tail) = STUDT / 2
'''                      One-tailed probability (left tail) = 1 - STUDT / 2
''' [(c) John C. Pezzullo](https://statpages.info/scicalc.html)
''' </summary>
''' <param name="t">t-value.</param>
''' <param name="n">Degrees of freedom.</param>
Private Function STUDT(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 2
            tmpEval = STUDT_( _
                            CDbl(tmpData(LB)), _
                            CDbl(tmpData(UB)) _
                            )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    STUDT = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function STUDT_(ByVal t As Double, n As Double) As Double
    Dim w As Double
    Dim th As Double
    Dim sth As Double
    Dim cth As Double
    
    t = Abs(t)
    w = t / Sqr(n)
    th = Atn(w)
    If (n = 1) Then STUDT_ = 1 - th / PID2: Exit Function
    sth = Sin(th): cth = Cos(th)
    If (REM_(n, 2) = 1) Then
        STUDT_ = 1 - (th + sth * cth * STATCOM(cth * cth, 2, n - 3, -1)) / PID2
    Else
        STUDT_ = 1 - sth * STATCOM(cth * cth, 1, n - 3, -1)
    End If
End Function

Private Function Switch_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim idx As Long
    Dim LB As Long, UB As Long
    Dim tmpData() As String
    Dim tmpEval As Boolean
    
    On Error GoTo err_handler
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
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function SYD_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    SYD_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function Tangent(ByRef expression As String, ByRef fName As String) As Double
    Dim tmpEval As Double
    
    On Error GoTo err_handler
    tmpEval = CDbl(expression)
    If P_DEGREES Then
        tmpEval = tmpEval * PI / 180
    End If
    Tangent = Tan(tmpEval)
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function tGamma(ByRef expression As String, ByRef fName As String) As Double
    On Error GoTo err_handler
    tGamma = GAMMA(CDbl(expression))
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function TimeSerial_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    TimeSerial_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function TimeValue_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    TimeValue_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

''' <summary>
''' Receives the degrees of freedom and the confidence level to
''' compute the one-tailed or two-tailed t-value (Student t-value)
''' with up to 6 significant digits accuracy. Use the tOption
''' parameter to select from two and right one-tailed computation.
''' [(c) iCalculator™](https://www.icalculator.com/)
''' </summary>
''' <param name="confidence">Confidence level.</param>
''' <param name="dof">Shape parameter alpha.</param>
''' <param name="tOption">Select from two and right one-tailed computation.</param>
Private Function TINV(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 3
            tmpEval = TINV_( _
                            CDbl(tmpData(LB)), _
                            CDbl(tmpData(LB + 1)), _
                            CDbl(tmpData(UB)) _
                            )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    TINV = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function TINV_(confidence As Double, dof As Double, _
                            Optional tOption As Single = 2) As Double
    Dim p As Double 'Probability
    
    If confidence > 1 And confidence < 100 Then 'Percentage entry
        confidence = confidence / 100
    End If
    p = (1 - confidence)
    If tOption = 1 Then
        TINV_ = method2(dof, p)
    Else
        If tOption = 2 Then
            TINV_ = method2(dof, p / 2)
        End If
    End If
End Function

''' <summary>
''' Receives the degrees of freedom and the Probability to
''' compute the one-tailed t-value (Student t-value) with up
''' to 6 significant digits accuracy. Parameter p can be passed
''' as 0 < p < 1, or 1 < p < 100
''' </summary>
''' <param name="p">Probability.</param>
''' <param name="dof">Degrees of freedom.</param>
Private Function TINV_1T(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 2
            tmpEval = TINV_( _
                            CDbl(tmpData(LB)), _
                            CDbl(tmpData(UB)), _
                            1 _
                            )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    TINV_1T = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function TINV_1T_(p As Double, dof As Double) As Double
    Dim confidence As Double
    
    If p > 1 Then  'Percentage entry
        p = p / 100
    End If
    confidence = 1 - p
    TINV_1T_ = TINV_(confidence, dof, 1)
End Function

''' <summary>
''' Receives the degrees of freedom and the Probability to
''' compute the two-tailed t-value (Student t-value) with full
''' significant digits accuracy. Parameter p can be passed
''' as 0 < p < 1, or 1 < p < 100
''' </summary>
''' <param name="p">Probability.</param>
''' <param name="dof">Degrees of freedom.</param>
Private Function TINV_2T(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
    tmpData() = Split(expression, P_SEPARATORCHAR)
    LB = LBound(tmpData)
    UB = UBound(tmpData)
    argsCount = UB - LB + 1
    Select Case argsCount
        Case 2
            tmpEval = TINV_2T_( _
                            CDbl(tmpData(LB)), _
                            CDbl(tmpData(UB)) _
                            )
        Case Else
            tmpEval = e_ValueError
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    TINV_2T = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function TINV_2T_(ByVal p As Double, dof As Long) As Double
    Dim X As Double
    
    If p > 1 Then 'Percentage entry
        p = p / 100
    End If
    X = iBETAINV(p, 0.5 * dof, 0.5)
    X = Sqr(dof * (1 - X) / X)
    If p > 0 Then
        TINV_2T_ = X
    Else
        TINV_2T_ = -X
    End If
End Function

Public Function ToDblArray(ByRef aArray As Variant) As Double()
    Dim i As Long, LB As Long, UB As Long
    Dim j As Long, LB2 As Long, UB2 As Long
    Dim tmpResult() As Double
    Dim IsVector As Boolean
    
    LB = LBound(aArray)
    UB = UBound(aArray)
    IsVector = Not Is2Darray(aArray)
    If Not IsVector Then
        LB2 = LBound(aArray, 2)
        UB2 = UBound(aArray, 2)
        ReDim tmpResult(0 To UB - LB, 0 To UB2 - LB2)
    Else
        ReDim tmpResult(0 To UB - LB)
    End If
    For i = LB To UB
        If IsVector Then
            tmpResult(i - LB) = CDbl(aArray(i))
        Else
            For j = LB2 To UB2
                tmpResult(i - LB, j - LB2) = CDbl(aArray(i, j))
            Next j
        End If
    Next i
    ToDblArray = tmpResult
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

Private Function ToLiteralString(ByRef aString As String) As String
    If Not IsLiteralString(aString) Then
        ToLiteralString = d_Apostrophe & aString & d_Apostrophe
    Else
        ToLiteralString = aString
    End If
End Function

Private Function Trim_(ByRef expression As String, ByRef fName As String) As String
    On Error GoTo err_handler
    Trim_ = ToLiteralString(Trim(FormatLiteralString(expression, True)))
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function UCase_(ByRef expression As String, ByRef fName As String) As String
    On Error GoTo err_handler
    UCase_ = ToLiteralString(UCase(expression))
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

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
    Dim tmpVstring As String
    Dim fCurlyPos As Long
    
    tmpVstring = ReconstructLiteralStrings(vString, Join$(Split(vString, d_Space), vbNullString))
    fCurlyPos = InStrB(1, tmpVstring, d_lCurly)
    If fCurlyPos > 0 Then
        If InStrB(fCurlyPos + 1, tmpVstring, d_lCurly) - fCurlyPos = 2 Then 'Full array
            tmpVstring = Replace(tmpVstring, d_rCurly & d_rCurly & P_SEPARATORCHAR, d_rCurly & d_rCurly & P_SEPARATORCHAR & P_SEPARATORCHAR)
        Else 'Vector
            tmpVstring = Replace(tmpVstring, d_rCurly & P_SEPARATORCHAR, d_rCurly & P_SEPARATORCHAR & P_SEPARATORCHAR)
        End If
    Else
        tmpVstring = Replace(tmpVstring, P_SEPARATORCHAR, P_SEPARATORCHAR & P_SEPARATORCHAR)
    End If
    tmpAssignment() = Split(tmpVstring, P_SEPARATORCHAR & P_SEPARATORCHAR)
    For avIcounter = LBound(tmpAssignment) To UBound(tmpAssignment)
        tmpValues() = Split(tmpAssignment(avIcounter), "=")
        UB = UBound(tmpValues)
        If tmpValues(UB) <> vbNullString Then
            If IsNumeric(tmpValues(UB)) Or IsBoolean(tmpValues(UB)) Or IsLiteralString(tmpValues(UB)) Then
                P_SCOPE.VarValue(tmpValues(LBound(tmpValues))) = tmpValues(UB)
            Else
                If MidB(tmpValues(UB), 1, 2) = d_lCurly And _
                        MidB(tmpValues(UB), LenB(tmpValues(UB)) - 1, 2) = d_rCurly Then 'Posible array assigment
                    P_SCOPE.VarValue(tmpValues(LBound(tmpValues))) = tmpValues(UB)
                End If
            End If
        End If
    Next avIcounter
End Sub

Private Sub VariablesInit(ByRef expression As String)
    P_SCOPE.VariablesInit
    ParseVariables expression
    P_SCOPE.FillPredefinedVars
End Sub

''' <summary>
''' Turns a row vector into a proper 2D array.
''' </summary>
''' <param name="tVector">The vertor to transform.</param>
Private Function VectorFormat(ByRef tVector As Variant) As Double()
    Dim vLB As Long
    Dim vLB2 As Long
    Dim vUB As Long
    Dim fromRowToCol As Boolean
    Dim i As Long, j As Long
    Dim tmpResult() As Double
    
    If Not IsArray(tVector) Then Exit Function
    vLB = LBound(tVector)
    vUB = UBound(tVector)
    If Is2Darray(tVector) Then
        tmpResult = tVector
    Else
        ReDim tmpResult(0 To 0, 0 To vUB - vLB)
        For i = vLB To vUB
            tmpResult(0, j) = tVector(i)
            j = j + 1
        Next i
    End If
    VectorFormat = tmpResult
End Function

Private Function WeekDayName_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    WeekDayName_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function WeekDay_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    WeekDay_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function Year_(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    Year_ = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

Private Function Zero(ByRef expression As String, ByRef fName As String) As String
    Dim argsCount As Long
    Dim tmpData() As String
    Dim tmpEval As String
    Dim LB As Long, UB As Long
    
    On Error GoTo err_handler
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
              BuildErrMessage errMissingArgsOrTooManyArgs, d_lCurly & fName & d_rCurly
    End Select
    Zero = tmpEval: Erase tmpData
    Exit Function
err_handler:
    BuildErrMessage errEvalError, d_lCurly & fName & d_rCurly & " | Error#: " & err.number & d_Space & _
                                    d_lParenthesis & err.Description & d_rParenthesis
End Function

