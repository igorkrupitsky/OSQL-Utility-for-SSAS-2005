Const ForAppending = 8
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim sFolder: sFolder = fso.GetParentFolderName(WScript.ScriptFullName)
Dim oLogFile: Set oLogFile = fso.CreateTextFile(WScript.ScriptFullName & ".log", True)
Dim sUrl
Dim sUserName
Dim sPassword

LoadConfigFile

Dim sFilePath
If WScript.Arguments.Count = 0 Then
	ProcessFolder sFolder & "\XMLA"
Else
	For i = 0 to WScript.Arguments.Count -1
		ProcessFile WScript.Arguments(i)
	Next	
End If

Log "Done!"
oLogFile.Close

'===================================================================
Sub ProcessFolder(sMyFolder)
	If Not fso.FolderExists(sMyFolder) Then
		Log "Folder does not exist: " & sMyFolder
		Exit Sub
	End If
	
	Dim oFolder, oFile
	Set oFolder = fso.GetFolder(sMyFolder)
    For Each oFile In oFolder.Files
        If LCase(Right(oFile.Name, 5)) = ".xmla" Then
			ProcessFile oFile.Path 
        End If
    Next
End Sub

Sub LoadConfigFile()
	Dim sConfigFile: sConfigFile = sFolder & "\XmlaConfig.xml"
	If Not fso.FileExists(sConfigFile) Then
		Log "Configuration file does not exist: " & sConfigFile
		WScript.Quit
	End If
	
	Dim oDoc: Set oDoc = CreateObject("MSXML2.DOMDocument")
	If Not oDoc.Load(sConfigFile) Then
		Log "Configuration file could not be loaded: " & sConfigFile & " " & oDoc.parseError.reason
		WScript.Quit
	End If

	sUrl = oDoc.SelectSingleNode("settings/url").Text
	sUserName = oDoc.SelectSingleNode("settings/user").Text
	sPassword = oDoc.SelectSingleNode("settings/password").Text
End Sub

Sub ProcessFile(sFilePath)
	If Not fso.FileExists(sFilePath) Then
		Log ("File does not exist: " & sFilePath)
		Exit Sub
	End If

	Log "Executing: " & sFilePath

	Dim sConents: sConents = GetFileContents(sFilePath)
	Dim strQuery: strQuery = EncloseTag("Command", sConents)
	Dim strProps: strProps = EncloseTag("Properties", "<PropertyList><Timeout>0</Timeout></PropertyList>")
	Dim sNS: sNS = " xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi = ""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"""

	Dim sPayload: sPayload = "<?xml version='1.0'?>" & _
			"<SOAP-ENV:Envelope " & sNS & ">" & _
			"<SOAP-ENV:Body>" & _
			" <Execute xmlns=""urn:schemas-microsoft-com:xml-analysis"" >" & _
			strQuery + strProps & _
			" </Execute>" & _
			"</SOAP-ENV:Body>" & _
			"</SOAP-ENV:Envelope>"

	Log PostData(sUrl, sPayload, sUserName, sPassword)
	Log ""
End Sub

Private Sub Log(sLine)
	oLogFile.WriteLine sLine
	'WScript.Echo sLine
End Sub

Private Function EncloseTag(sTag, sValue)
	EncloseTag = "<" & sTag & ">" & sValue & "</" & sTag & ">"
End Function

Private Function PostData(sUrl, sData, sUserName, sPassword)
	Dim oHttp
	'Set oHttp = CreateObject("Microsoft.XMLHTTP")
	Set oHttp = CreateObject("MSXML2.ServerXMLHTTP")
    oHttp.setTimeouts 0, 0, 0, 0
	
	oHttp.Open "POST", sUrl, False, sUserName, sPassword
	oHttp.setRequestHeader "SOAPAction", """urn:schemas-microsoft-com:xml-analysis:Execute"""
	oHttp.Send sData
	PostData = oHttp.responseText
	Set oHttp = Nothing
End Function

Public Function GetFileContents(sFilePath)
	Dim sContents
	Const ForReading = 1
	Const TristateMixed = -2
	Set oTextFile = fso.OpenTextFile(sFilePath, ForReading, False, TristateMixed)
	Do While Not oTextFile.AtEndOfStream
		sContents = sContents & oTextFile.ReadLine
	Loop
	oTextFile.Close
	GetFileContents = sContents
End Function
