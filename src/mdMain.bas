Attribute VB_Name = "mdMain"
'=========================================================================
'
' VbPeg (c) 2018 by wqweto@gmail.com
'
' PEG parser generator for VB6
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Const STD_OUTPUT_HANDLE             As Long = -11&
Private Const STD_ERROR_HANDLE              As Long = -12&
'--- for CreateFile
Private Const GENERIC_WRITE                 As Long = &H40000000
Private Const OPEN_EXISTING                 As Long = 3
Private Const FILE_SHARE_READ               As Long = &H1

Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, lpszDst As Any, ByVal cchDstLength As Long) As Long
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ApiSysAllocString Lib "oleaut32" Alias "SysAllocString" (ByVal Ptr As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function IsTextUnicode Lib "advapi32" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long
Public Declare Function EmptyLongArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbLong, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As Long()
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByVal lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_VERSION           As String = "0.2"

'=========================================================================
' Functions
'=========================================================================

Private Sub Main()
    Dim oOpt           As Variant
    Dim oTree           As cTree
    Dim oParser         As cParser
    Dim oIR             As cIR
    Dim nFile           As Integer
    Dim cOutput         As Collection
    Dim sOutput         As String
    
    On Error GoTo EH
    Set oParser = New cParser
    Set oOpt = GetOpt(SplitArgs(Command$), "o:module:userdata")
    If Not oOpt.Item("-nologo") And Not oOpt.Item("-q") Then
        ConsoleError "VbPeg " & STR_VERSION & " (c) 2018 by wqweto@gmail.com (" & oParser.ParserVersion & ")" & vbCrLf & vbCrLf
    End If
    If LenB(oOpt.Item("error")) <> 0 Then
        ConsoleError "Error in command line: " & oOpt.Item("error") & vbCrLf & vbCrLf
        If Not (oOpt.Item("-h") Or oOpt.Item("-?") Or oOpt.Item("arg1") = "?") Then
            Exit Sub
        End If
    End If
    If oOpt.Item("numarg") = 0 Or oOpt.Item("-h") Or oOpt.Item("-?") Or oOpt.Item("arg1") = "?" Then
        ConsoleError "Usage: %1.exe [options] <in_file.peg>" & vbCrLf & vbCrLf, App.EXEName
        ConsoleError "Options:" & vbCrLf & _
            "  -o OUTFILE      write result to OUTFILE [default: stdout]" & vbCrLf & _
            "  -tree           output parse tree" & vbCrLf & _
            "  -ir             output intermediate represetation" & vbCrLf & _
            "  -public         emit public VB6 class module" & vbCrLf & _
            "  -private        emit private VB6 class module" & vbCrLf & _
            "  -module NAME    VB6 class/module name [default: OUTFILE]" & vbCrLf & _
            "  -userdata TYPE  parser context's UserData member data-type [default: Variant]" & vbCrLf & _
            "  -q              in quiet operation outputs only errors" & vbCrLf & vbCrLf & _
            "If no -tree/-ir is used emits VB6 code. If no -o is used writes result to console. If no -public/-private is used emits standard .bas module." & vbCrLf
        Exit Sub
    End If
    Set oTree = New cTree
    If oParser.Match(ReadTextFile(oOpt.Item("arg1")), UserData:=oTree) = 0 Then
        ConsoleError "Error parsing: %1" & vbCrLf, oParser.LastError
        Exit Sub
    End If
    If Not oTree.OptimizeTree() Then
        ConsoleError "Failed tree optimize: %1" & vbCrLf, oTree.LastError
        Exit Sub
    End If
    If oOpt.Item("-tree") Then
        sOutput = oTree.DumpParseTree
    Else
        Set oIR = New cIR
        If Not oIR.CodeGen(oTree, oOpt.Item("-allrules")) Then
            ConsoleError "Failed codegen: %1" & vbCrLf, oIR.LastError
            Exit Sub
        End If
        If oOpt.Item("-ir") Then
            sOutput = oIR.DumpIrTree
        Else
            If LenB(oOpt.Item("-module")) = 0 Then
                oOpt.Item("-module") = GetFilePart(oOpt.Item("-o"))
            End If
            If Not oIR.EmitCode(cOutput, _
                    oOpt.Item("-public") Or oOpt.Item("-private"), _
                    oOpt.Item("-public"), _
                    CStr(oOpt.Item("-module")), _
                    CStr(oOpt.Item("-userdata"))) Then
                ConsoleError "Failed emit: %1" & vbCrLf, oIR.LastError
                Exit Sub
            End If
            sOutput = ConcatCollection(cOutput, vbCrLf)
        End If
    End If
    '--- write output
    If InIde Then
        Clipboard.Clear
        Clipboard.SetText sOutput
    End If
    If LenB(oOpt.Item("-o")) > 0 Then
        '--- fix output file extension if not supplied
        If InStrRev(oOpt.Item("-o"), "\") >= InStrRev(oOpt.Item("-o"), ".") Then
            oOpt.Item("-o") = oOpt.Item("-o") & IIf(oOpt.Item("-public") Or oOpt.Item("-private"), ".cls", ".bas")
        End If
        SetFileLen oOpt.Item("-o"), Len(sOutput)
        nFile = FreeFile
        Open oOpt.Item("-o") For Binary Access Write Shared As nFile
        Put nFile, , sOutput
        Close nFile
        If Not oOpt.Item("-q") Then
            ConsoleError "File " & oOpt.Item("-o") & " successfully emitted" & vbCrLf
        End If
    Else
        ConsolePrint sOutput
    End If
    Exit Sub
EH:
    ConsoleError "Critical error: " & Err.Description & vbCrLf
End Sub

Private Function GetOpt(vArgs As Variant, Optional OptionsWithArg As String) As Object
    Dim oRetVal         As Object
    Dim lIdx            As Long
    Dim bNoMoreOpt      As Boolean
    Dim vOptArg         As Variant
    Dim vElem           As Variant

    vOptArg = Split(OptionsWithArg, ":")
    Set oRetVal = CreateObject("Scripting.Dictionary")
    With oRetVal
        .CompareMode = vbTextCompare
        For lIdx = 0 To UBound(vArgs)
            Select Case Left$(At(vArgs, lIdx), 1 + bNoMoreOpt)
            Case "-", "/"
                For Each vElem In vOptArg
                    If Mid$(At(vArgs, lIdx), 2, Len(vElem)) = vElem Then
                        If Mid(At(vArgs, lIdx), Len(vElem) + 2, 1) = ":" Then
                            .Item("-" & vElem) = Mid$(At(vArgs, lIdx), Len(vElem) + 3)
                        ElseIf Len(At(vArgs, lIdx)) > Len(vElem) + 1 Then
                            .Item("-" & vElem) = Mid$(At(vArgs, lIdx), Len(vElem) + 2)
                        ElseIf LenB(At(vArgs, lIdx + 1)) <> 0 Then
                            .Item("-" & vElem) = At(vArgs, lIdx + 1)
                            lIdx = lIdx + 1
                        Else
                            .Item("error") = "Option -" & vElem & " requires an argument"
                        End If
                        GoTo Conitnue
                    End If
                Next
                .Item("-" & Mid$(At(vArgs, lIdx), 2)) = True
            Case Else
                .Item("numarg") = .Item("numarg") + 1
                .Item("arg" & .Item("numarg")) = At(vArgs, lIdx)
            End Select
Conitnue:
        Next
    End With
    Set GetOpt = oRetVal
End Function

Public Function ConsolePrint(ByVal sText As String, ParamArray A() As Variant) As String
    ConsolePrint = pvConsoleOutput(GetStdHandle(STD_OUTPUT_HANDLE), sText, CVar(A))
End Function

Public Function ConsoleError(ByVal sText As String, ParamArray A() As Variant) As String
    ConsoleError = pvConsoleOutput(GetStdHandle(STD_ERROR_HANDLE), sText, CVar(A))
End Function

Private Function pvConsoleOutput(ByVal hOut As Long, ByVal sText As String, A As Variant) As String
    Const LNG_PRIVATE   As Long = &HE1B6 '-- U+E000 to U+F8FF - Private Use Area (PUA)
    Dim lIdx            As Long
    Dim sArg            As String
    Dim baBuffer()      As Byte
    Dim dwDummy         As Long

    If LenB(sText) = 0 Then
        Exit Function
    End If
    '--- format
    For lIdx = UBound(A) To LBound(A) Step -1
        sArg = Replace(A(lIdx), "%", ChrW$(LNG_PRIVATE))
        sText = Replace(sText, "%" & (lIdx - LBound(A) + 1), sArg)
    Next
    pvConsoleOutput = Replace(sText, ChrW$(LNG_PRIVATE), "%")
    '--- output
    If hOut = 0 Then
        Debug.Print pvConsoleOutput;
    Else
        ReDim baBuffer(0 To Len(pvConsoleOutput) - 1) As Byte
        If CharToOemBuff(pvConsoleOutput, baBuffer(0), UBound(baBuffer) + 1) Then
            Call WriteFile(hOut, baBuffer(0), UBound(baBuffer) + 1, dwDummy, ByVal 0&)
        End If
    End If
End Function

Public Function SplitArgs(sText As String) As Variant
    Dim vRetVal         As Variant
    Dim lPtr            As Long
    Dim lArgc           As Long
    Dim lIdx            As Long
    Dim lArgPtr         As Long

    If LenB(sText) <> 0 Then
        lPtr = CommandLineToArgvW(StrPtr(sText), lArgc)
    End If
    If lArgc > 0 Then
        ReDim vRetVal(0 To lArgc - 1) As String
        For lIdx = 0 To UBound(vRetVal)
            Call CopyMemory(lArgPtr, ByVal lPtr + 4 * lIdx, 4)
            vRetVal(lIdx) = SysAllocString(lArgPtr)
        Next
    Else
        vRetVal = Split(vbNullString)
    End If
    Call LocalFree(lPtr)
    SplitArgs = vRetVal
End Function

Private Function SysAllocString(ByVal lPtr As Long) As String
    Dim lTemp           As Long

    lTemp = ApiSysAllocString(lPtr)
    Call CopyMemory(ByVal VarPtr(SysAllocString), lTemp, 4)
End Function

Public Function ReadTextFile(sFile As String) As String
    Const ForReading    As Long = 1
    Const BOM_UTF       As String = "ï»¿"   '--- "\xEF\xBB\xBF"
    Const BOM_UNICODE   As String = "ÿþ"    '--- "\xFF\xFE"
    Dim lSize           As Long
    Dim sPrefix         As String
    Dim nFile           As Integer
    Dim sCharset        As String
    Dim oStream         As Object
    
    '--- get file size
    On Error GoTo EH
    If FileExists(sFile) Then
        lSize = FileLen(sFile)
    End If
    If lSize = 0 Then
        Exit Function
    End If
    '--- read first 50 chars
    nFile = FreeFile
    Open sFile For Binary Access Read Shared As nFile
    sPrefix = String$(IIf(lSize < 50, lSize, 50), 0)
    Get nFile, , sPrefix
    Close nFile
    '--- figure out charset
    If Left$(sPrefix, 3) = BOM_UTF Then
        sCharset = "UTF-8"
    ElseIf Left$(sPrefix, 2) = BOM_UNICODE Or IsTextUnicode(ByVal sPrefix, Len(sPrefix), &HFFFF& - 2) <> 0 Then
        sCharset = "Unicode"
    ElseIf InStr(1, sPrefix, "<?xml", vbTextCompare) > 0 And InStr(1, sPrefix, "utf-8", vbTextCompare) > 0 Then
        '--- special xml encoding test
        sCharset = "UTF-8"
    End If
    '--- plain text: direct VB6 read
    If LenB(ReadTextFile) = 0 And LenB(sCharset) = 0 Then
        nFile = FreeFile
        Open sFile For Binary Access Read Shared As nFile
        ReadTextFile = String$(lSize, 0)
        Get nFile, , ReadTextFile
        Close nFile
    End If
    '--- plain text + unicode: use FileSystemObject
    If LenB(ReadTextFile) = 0 And sCharset <> "UTF-8" Then
        On Error Resume Next  '--- checked
        ReadTextFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(sFile, ForReading, False, sCharset = "Unicode").ReadAll()
        On Error GoTo EH
    End If
    '--- plain text + unicode + utf-8: use ADODB.Stream
    If LenB(ReadTextFile) = 0 Then
        Set oStream = CreateObject("ADODB.Stream")
        With oStream
            .Open
            If LenB(sCharset) <> 0 Then
                .Charset = sCharset
            End If
            .LoadFromFile sFile
            ReadTextFile = .ReadText()
        End With
    End If
    Exit Function
EH:
End Function

Public Function FileExists(sFile As String) As Boolean
    If GetFileAttributes(sFile) = -1 Then ' INVALID_FILE_ATTRIBUTES
    Else
        FileExists = True
    End If
End Function

Public Function At(vArray As Variant, ByVal lIdx As Long) As Variant
    On Error GoTo QH
    If lIdx >= LBound(vArray) And lIdx <= UBound(vArray) Then
        At = vArray(lIdx)
    End If
QH:
End Function

Public Function ConcatCollection(oCol As Collection, Optional Separator As String) As String
    Dim lSize           As Long
    Dim vElem           As Variant
    
    For Each vElem In oCol
        lSize = lSize + Len(vElem) + Len(Separator)
    Next
    If lSize > 0 Then
        ConcatCollection = String$(lSize - Len(Separator), 0)
        lSize = 1
        For Each vElem In oCol
            If lSize <= Len(ConcatCollection) Then
                Mid$(ConcatCollection, lSize, Len(vElem) + Len(Separator)) = vElem & Separator
            End If
            lSize = lSize + Len(vElem) + Len(Separator)
        Next
    End If
End Function

Public Property Get InIde() As Boolean
    Debug.Assert pvSetTrue(InIde)
End Property

Private Function pvSetTrue(bValue As Boolean) As Boolean
    bValue = True
    pvSetTrue = True
End Function

Public Function GetFilePart(sFileName As String) As String
    GetFilePart = Mid$(sFileName, InStrRev(sFileName, "\") + 1)
    If InStrRev(GetFilePart, ".") > 0 Then
        GetFilePart = Left$(GetFilePart, InStrRev(GetFilePart, ".") - 1)
    End If
End Function

Public Function SetFileLen(sFile As String, ByVal lSize As Long) As Boolean
    Dim hFile       As Long
    
    hFile = CreateFile(sFile, GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, 0, 0)
    If hFile <> 0 Then
        If SetFilePointer(hFile, lSize, 0, 0) <> -1 Then
            If SetEndOfFile(hFile) <> 0 Then
                SetFileLen = True
            End If
        End If
        Call CloseHandle(hFile)
    End If
End Function
