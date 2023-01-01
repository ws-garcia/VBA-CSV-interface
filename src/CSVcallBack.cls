VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVcallBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'#
'////////////////////////////////////////////////////////////////////////////////////////////
' Copyright © 2022-2023 W. García
' GPL-3.0 license | https://www.gnu.org/licenses/gpl-3.0.html/
' https://ingwilfredogarcia.wordpress.com
'////////////////////////////////////////////////////////////////////////////////////////////
'#
' GENERAL INFO:
' Class module developed to provide a way to extend the VBAexpression.cls and allow methods
' call back for user defined functions. See: https://stackoverflow.com/a/48372415
'
' Users can register custom modules to expose and use their functions for throght the
' CSVcallBack.cls module. All UDFs must have a single Variant argument that will receive a
' one-dimensional array of strings (one element for each function argument).
'
' For example, if the functions to be used are in a class module called CSVudFunctions.cls, the
' line ['Public UserDefFunctions As New CSVudFunctions'] will be sufficient to expose all its
' internal functions.
'
' After doing this, the user must "bind" a procedure to the desired methods through the
' DeclareUDF method of the VBAexpression.cls:
'
'                                               DeclareUDF UDFname, UDFlib
'
' In the above expression, the UDFname parameter can be an array of function names or a
' string function name. The UDFlib is a string with the name of the class module to be linked.
'////////////////////////////////////////////////////////////////////////////////////////////
'@--------------------------------------------------------------------
' Expose functions defined in the CSVudFunctions.cls
Public UserDefFunctions As New CSVudFunctions
