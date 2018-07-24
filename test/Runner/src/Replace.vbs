Option Explicit
On Error Resume Next

Const ForReading = 1
Const ForWriting = 2
Const vbTextCompare = 1 
Const vbBinaryCompare = 0 

Dim sSrcFile
Dim sDestFile
Dim sSearch
Dim sReplace
Dim oFS
Dim oStream
Dim sContents
Dim bInputUnicode
Dim bOutputUnicode
Dim lOrigLen
Dim bVerbose
Dim bCaseSensitive
Dim lNumReplaces

sSrcFile = WScript.Arguments.Named("f")
sDestFile = WScript.Arguments.Named("d")
If Len(sDestFile) = 0 Then
	sDestFile = sSrcFile
End If
sSearch = TranslateArg(WScript.Arguments.Named("s"))
If WScript.Arguments.Named.Exists("rf") Then
	sReplace = ReadFile(WScript.Arguments.Named("rf"))
Else
	sReplace = TranslateArg(WScript.Arguments.Named("r"))
End If
bOutputUnicode = WScript.Arguments.Named.Exists("u")
bVerbose = WScript.Arguments.Named.Exists("v")
bCaseSensitive = WScript.Arguments.Named.Exists("cs")

If Len(sSrcFile) = 0 Or Len(sSearch) = 0 Then
    WScript.Echo "" & vbCrLf & _
    		"Usage: Replace.vbs /f:SourceFile [/d:DestFile] /s:SearchString [/r:ReplaceWithString] [/rf:ReplaceWithFileContents] [/u] [/v] [/cs]" & vbCrLf & _
    		"" & vbCrLf & _
    		"      SourceFile - the text file to search into" & vbCrLf & _
    		"      DestFile - results file. if not specified defaults to SourceFile" & vbCrLf & _
    		"      SearchString - string to find" & vbCrLf & _
    		"      ReplaceWithString - string to replace with the SearchString. default to empty string"  & vbCrLf & _
    		"      ReplaceWithFileContents - file with contents to replace with the SearchString"  & vbCrLf & _
    		"      /u - save output file in UNICODE "  & vbCrLf & _
    		"      /v - verbose output"  & vbCrLf & _
    		"      /cs - case sensitive compare"  & vbCrLf
    WScript.Quit 1
End If

Set oFS = CreateObject("Scripting.FileSystemObject")
'--- check if input unicode
bInputUnicode = False
bInputUnicode = (oFS.OpenTextFile(sSrcFile, ForReading, False, False).Read(2) = Chr(&hFF) & Chr(&HFE))		
'--- open src file
Set oStream = oFS.OpenTextFile(sSrcFile, ForReading, False, bInputUnicode)
If oStream Is Nothing Then
	WScript.Echo "Error reading from " & sSrcFile & vbCrLf & vbCrLf & Error
	WScript.Quit 2
End If
'--- read src file
sContents = oStream.ReadAll()
oStream.Close
Set oStream = Nothing
'--- replace
lOrigLen = Len(sContents)
If bVerbose Then
	WScript.Echo "Replacing " & sSearch & " -> " & sReplace & IIf(bCaseSensitive, " (case-sensitive)", " (case-insensitive)")
End If
sContents = Replace(sContents, sSearch, sReplace, 1, -1, IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare))
If bVerbose Then	
	If (Len(sSearch) - Len(sReplace)) <> 0 Then
		lNumReplaces = (lOrigLen - Len(sContents)) \ (Len(sSearch) - Len(sReplace))
	Else
		lNumReplaces = 0
	End If
	WScript.Echo lNumReplaces & " replacement(s)"
End If
'--- open dest file
If IsEmpty(bOutputUnicode) Then
	bOutputUnicode = bInputUnicode
End If
Set oStream = oFS.OpenTextFile(sDestFile, ForWriting, True, bOutputUnicode)
If oStream Is Nothing Then
	WScript.Echo "Error writing to " & sDestFile & vbCrLf & vbCrLf & Error
	WScript.Quit 3
End If
'--- write to dest file
oStream.Write sContents
oStream.Close
Set oStream = Nothing
Set oFS = Nothing

WScript.Quit 0

Function TranslateArg(sArg)
	TranslateArg = Replace(Replace(Replace(sArg, "^p", vbCrLf), "^t", vbTab), "^q", """")
End Function

Function IIf(Cond, TruePart, FalsePart)
	If Cond Then
		IIf = TruePart
	Else
		IIf = FalsePart
	End If
End Function

Private Function ReadFile(sFile) 
    Const ForReading = 1
    With CreateObject("Scripting.FileSystemObject")
        ReadFile = .OpenTextFile(sFile, ForReading, False, _
            .OpenTextFile(sFile, ForReading, False, False).Read(2) = Chr(&HFF) & Chr(&HFE)).ReadAll()
    End With
End Function
