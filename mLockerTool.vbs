'VisualBasic Script
'@author Cédric LEVEQUE
'@version 1.0

Option Explicit

Dim objArgs, baseDir, traceDir, finalDir, fichier

Set objArgs = WScript.Arguments
baseDir = objArgs(0)
traceDir = objArgs(1)
finalDir = objArgs(2)
fichier = objArgs(3)

Call Programme( fichier, baseDir, traceDir, finalDir )

Sub Programme( myFile, baseDir, traceDir, finalDir )
	On Error Resume Next

    Dim objFSO, objXls, objDoc, objFile, objWord, objExcel, strFile, strTrace, strDOC, strXLSx

    Const wdFormatDocument                    =  0
    Const wdFormatDocument97                  =  0
    Const wdFormatDocumentDefault             = 16
    Const wdFormatDOSText                     =  4
    Const wdFormatDOSTextLineBreaks           =  5
    Const wdFormatEncodedText                 =  7
    Const wdFormatFilteredHTML                = 10
    Const wdFormatFlatXML                     = 19
    Const wdFormatFlatXMLMacroEnabled         = 20
    Const wdFormatFlatXMLTemplate             = 21
    Const wdFormatFlatXMLTemplateMacroEnabled = 22
    Const wdFormatHTML                        =  8
    Const wdFormatPDF                         = 17
    Const wdFormatRTF                         =  6
    Const wdFormatTemplate                    =  1
    Const wdFormatTemplate97                  =  1
    Const wdFormatText                        =  2
    Const wdFormatTextLineBreaks              =  3
    Const wdFormatUnicodeText                 =  7
    Const wdFormatWebArchive                  =  9
    Const wdFormatXML                         = 11
    Const wdFormatXMLDocument                 = 12
    Const wdFormatXMLDocumentMacroEnabled     = 13
    Const wdFormatXMLTemplate                 = 14
    Const wdFormatXMLTemplateMacroEnabled     = 15
    Const wdFormatXPS                         = 18
    Const wdFormatOfficeDocumentTemplate      = 23
    Const wdFormatMediaWiki                   = 24
	Const xlFormatXlsx				  		  = 51
	Const xlFormatXls						  = 56
	Const xlFormatCsv						  =  6
	
	Set objFSO = CreateObject( "Scripting.FileSystemObject" )
    
	Set objWord = CreateObject( "Word.Application" )
	
	Set objExcel = CreateObject( "Excel.Application" )
	
	If LCase( objFSO.GetExtensionName( myFile ) ) = "doc" Then
		With objWord
			objWord.Visible = False

			If objFSO.FileExists( myFile ) Then
				Set objFile = objFSO.GetFile( myFile )
				strFile = objFile.Path
			Else
				WScript.Echo( "ERREUR OUVERTURE FICHIER: Le fichier n'existe pas ou plus" & vbCrLf )
				Exit Sub
			End If

			strTrace = objFSO.BuildPath( traceDir & "\", objFSO.GetBaseName( objFile ) & ".xml" )
					  
			strDOC = objFSO.BuildPath( finalDir & "\", objFSO.GetBaseName(objFile) & ".doc" )

			Call objWord.Documents.Open( strFile, , True )
			Set objDoc = objWord.ActiveDocument
			objDoc.SaveAs strTrace, wdFormatXML
			objDoc.Close
				
			Call objWord.Documents.Open( strTrace )
			Set objDoc = objWord.ActiveDocument
			objDoc.SaveAs strDOC, wdFormatDocument
			objDoc.Close
		End With
	ElseIf LCase( objFSO.GetExtensionName( myFile ) ) = "xls" Then
		With objExcel
			objExcel.Visible = False

			If objFSO.FileExists( myFile ) Then
				Set objFile = objFSO.GetFile( myFile )
				strFile = objFile.Path
			Else
				WScript.Echo( "ERREUR OUVERTURE FICHIER: Le fichier n'existe pas" & vbCrLf )
				Exit Sub
			End If
					  
			strXLSx = objFSO.BuildPath( finalDir & "\", objFSO.GetBaseName(objFile) & ".xlsx" )
			
			Set objXls = objExcel.Workbooks.Open( strFile, , True )
			objExcel.ActiveWorkbook.CheckCompatibility = False
			objXls.SaveAs strXLSx, xlFormatXlsx
			Call objXls.Close( False )

		End With
	Else
		If objFSO.FileExists( myFile ) Then
				Call objFSO.MoveFile( myFile, finalDir & "\" )
			Else
				WScript.Echo( "ERREUR DETECTION FICHIER: Le fichier n'existe pas ou plus" & vbCrLf )
			Exit Sub
		End If
	End If
	objWord.Quit
	objExcel.Quit
End Sub