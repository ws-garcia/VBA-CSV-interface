Attribute VB_Name = "CSVinterfaceTest"
Option Explicit
Sub TestImportLikeRFC-4180CSV()
    Dim CSVix As CSVinterface
    Dim fileName As String
    Dim tmpCSV As String
	 
    Set CSVix = New CSVinterface
    
    fileName = "C:\RFC-4180_HF.csv"
	 tmpCSV = CSVix.GetDataFromCSV(fileName)
	 Call CSVix.ImportFromCSVString(tmpCSV)
	 If CSVix.ErrNumber <> 0 Then Goto Err_Handler
    Set CSVix = Nothing
	 Exit Sub
Err_Handler:
    Debug.Print "Returned Error #", CSVix.ErrNumber; CSVix.ErrDescription
End Sub

Sub TestImportTSV()
    Dim CSVix As CSVinterface
    Dim fileName As String
    Dim tmpCSV As String
	 
    Set CSVix = New CSVinterface
    
    fileName = "C:\TestTSV.tsv"
	 tmpCSV = CSVix.GetDataFromCSV(fileName)
	 Call CSVix.ImportFromCSVString(tmpCSV)
	 If CSVix.ErrNumber <> 0 Then Goto Err_Handler
    Set CSVix = Nothing
	 Exit Sub
Err_Handler:
    Debug.Print "Returned Error #", CSVix.ErrNumber; CSVix.ErrDescription
End Sub