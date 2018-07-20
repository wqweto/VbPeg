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

#Const HasIVbCollection = False

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
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_VERSION           As String = "0.3.10"

Private m_oParser               As cParser
Private m_oOpt                  As Object

'=========================================================================
' Functions
'=========================================================================

Private Sub Main()
    Dim lExitCode       As Long
    
    lExitCode = Process(SplitArgs(Command$))
    If Not InIde Then
        Call ExitProcess(lExitCode)
    End If
End Sub

Private Function Process(vArgs As Variant) As Long
    Dim oTree           As cTree
    Dim oIR             As cIR
    Dim nFile           As Integer
    Dim sOutput         As String
    Dim vElem           As Variant
    Dim cWarnings       As Collection
    Dim lOffset         As Long
    Dim lIdx            As Long
    
    On Error GoTo EH
    Set m_oParser = New cParser
    Set m_oOpt = GetOpt(vArgs, "o:module:userdata")
    If Not m_oOpt.Item("-nologo") And Not m_oOpt.Item("-q") Then
        ConsoleError "VbPeg " & STR_VERSION & " (c) 2018 by wqweto@gmail.com (" & m_oParser.ParserVersion & ")" & vbCrLf & vbCrLf
    End If
    If LenB(m_oOpt.Item("error")) <> 0 Then
        ConsoleError "Error in command line: " & m_oOpt.Item("error") & vbCrLf & vbCrLf
        If Not (m_oOpt.Item("-h") Or m_oOpt.Item("-?") Or m_oOpt.Item("arg1") = "?") Then
            Exit Function
        End If
    End If
    If m_oOpt.Item("numarg") = 0 Or m_oOpt.Item("-h") Or m_oOpt.Item("-?") Or m_oOpt.Item("arg1") = "?" Then
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
        If m_oOpt.Item("numarg") = 0 Then
            Process = 100
        End If
        Exit Function
    End If
    Set oTree = New cTree
    For lIdx = 1 To m_oOpt.Item("numarg")
        oTree.AddFileToQueue CanonicalPath(m_oOpt.Item("arg" & lIdx))
    Next
    lIdx = 1
    Do While lIdx <= oTree.FileQueue.Count
        If Not m_oOpt.Item("-q") Then
            ConsoleError "%1" & vbCrLf, PathDifference(CurDir$, oTree.FileQueue.Item(lIdx))
        End If
        lOffset = m_oParser.Match(oTree.ReadFile(oTree.FileQueue.Item(lIdx)), UserData:=oTree)
        If LenB(m_oParser.LastError) Then
            ConsoleError "%2: %3: %1" & vbCrLf, m_oParser.LastError, Join(oTree.CalcLine(m_oParser.LastOffset + 1), ":"), IIf(lOffset = 0, "error", "warning")
        End If
        If Not m_oParser.VbPegGetParseErrors() Is Nothing Then
            For Each vElem In m_oParser.VbPegGetParseErrors()
                ConsoleError "%2: %3: %1" & vbCrLf, At(vElem, 0), Join(oTree.CalcLine(At(vElem, 1)), ":"), IIf(lOffset = 0, "error", "warning")
            Next
        End If
        If lOffset = 0 Then
            Process = 1
            Exit Function
        End If
        lIdx = lIdx + 1
    Loop
    If Not oTree.CheckTree(cWarnings) Then
        ConsoleError "%1" & vbCrLf, oTree.LastError
        Process = 2
        Exit Function
    ElseIf Not cWarnings Is Nothing Then
        For Each vElem In cWarnings
            ConsoleError "%2: %3: %1" & vbCrLf, At(vElem, 0), Join(oTree.CalcLine(At(vElem, 1)), ":"), IIf(lOffset = 0, "error", "warning")
        Next
    End If
    If Not oTree.OptimizeTree(cWarnings) Then
        ConsoleError "Optimize failed: %1" & vbCrLf, oTree.LastError
        Process = 3
        Exit Function
    ElseIf Not cWarnings Is Nothing Then
        For Each vElem In cWarnings
            ConsoleError "%2: %3: %1" & vbCrLf, At(vElem, 0), Join(oTree.CalcLine(At(vElem, 1)), ":"), IIf(lOffset = 0, "error", "warning")
        Next
    End If
    If m_oOpt.Item("-tree") Then
        sOutput = oTree.DumpParseTree
    Else
        Set oIR = New cIR
        If Not oIR.CodeGen(oTree, m_oOpt.Item("-allrules")) Then
            ConsoleError "Failed codegen: %1" & vbCrLf, oIR.LastError
            Process = 4
            Exit Function
        End If
        If m_oOpt.Item("-ir") Then
            sOutput = oIR.DumpIrTree
        Else
            If LenB(m_oOpt.Item("-module")) = 0 Then
                m_oOpt.Item("-module") = GetFilePart(m_oOpt.Item("-o"))
            End If
            If Not oIR.EmitCode( _
                    Switch(C_Bool(m_oOpt.Item("-public")), vbTrue, C_Bool(m_oOpt.Item("-private")), vbFalse, True, vbUseDefault), _
                    CStr(m_oOpt.Item("-module")), _
                    CStr(m_oOpt.Item("-userdata")), _
                    sOutput) Then
                ConsoleError "Failed emit: %1" & vbCrLf, oIR.LastError
                Process = 5
                Exit Function
            End If
        End If
    End If
    '--- write output
    If InIde Then
        Clipboard.Clear
        Clipboard.SetText sOutput
    End If
    If LenB(m_oOpt.Item("-o")) > 0 Then
        '--- fix output file extension if not supplied
        If InStrRev(m_oOpt.Item("-o"), "\") >= InStrRev(m_oOpt.Item("-o"), ".") Then
            m_oOpt.Item("-o") = m_oOpt.Item("-o") & IIf(m_oOpt.Item("-public") Or m_oOpt.Item("-private"), ".cls", ".bas")
        End If
        SetFileLen m_oOpt.Item("-o"), Len(sOutput)
        nFile = FreeFile
        Open m_oOpt.Item("-o") For Binary Access Write Shared As nFile
        Put nFile, , sOutput
        Close nFile
        If Not m_oOpt.Item("-q") Then
            ConsoleError "File " & m_oOpt.Item("-o") & " emitted successfully" & vbCrLf
        End If
    Else
        ConsolePrint sOutput
    End If
    Exit Function
EH:
    ConsoleError "Critical error: " & Err.Description & vbCrLf
    Process = 100
End Function

Public Function ConsoleTrace(ByVal lOffset As Long, sRule As String, ByVal lAction As Long, oUserData As cTree) As Boolean
    Const LINE_LEN      As Long = 8
    Const TEXT_LEN      As Long = 60
    Static lLevel       As Long
    Dim sText           As String
    Dim sLine           As String
    
    If C_Bool(m_oOpt.Item("-trace")) Then
        sText = Mid$(m_oParser.VbPegGetContents(), lOffset, TEXT_LEN)
        If InStr(sText, vbCr) > 0 Then
            sText = Left$(sText, InStr(sText, vbCr) - 1)
        End If
        If Len(sText) < TEXT_LEN Then
            sText = sText & Space$(TEXT_LEN - Len(sText))
        End If
        sLine = Join(oUserData.CalcLine(lOffset), ":")
        If Len(sLine) - InStr(sLine, ":") < LINE_LEN Then
            sLine = sLine & Space$(LINE_LEN - Len(sLine) + InStr(sLine, ":"))
        End If
        If lAction = 1 Then
            ConsoleError "%1|%2|%3?%4" & vbCrLf, sLine, sText, Space$(lLevel * 2), sRule
            lLevel = lLevel + 1
        Else
            If lLevel > 0 Then
                lLevel = lLevel - 1
            End If
            If lAction = 2 Then
                ConsoleError "%1|%2|%3=%4" & vbCrLf, sLine, sText, Space$(lLevel * 2), sRule
            Else
                ConsoleError "%1|%2|%3!%4" & vbCrLf, sLine, sText, Space$(lLevel * 2), sRule
            End If
        End If
    End If
End Function

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

#If HasIVbCollection Then
    Public Function SearchCollection(oCol As IVbCollection, Index As Variant) As Boolean
        SearchCollection = (oCol.Item(Index) >= 0)
    End Function
#Else
    Public Function SearchCollection(oCol As Collection, Index As Variant) As Boolean
        On Error GoTo QH
        oCol.Item Index
        SearchCollection = True
QH:
    End Function
#End If

Public Function Zn(sText As String, Optional IfEmptyString As Variant = Null) As Variant
    Zn = IIf(LenB(sText) = 0, IfEmptyString, sText)
End Function

Public Function C_Bool(vValue As Variant) As Boolean
    On Error GoTo QH
    If LenB(vValue) <> 0 Then
        C_Bool = CBool(vValue)
    End If
QH:
End Function

Public Function CanonicalPath(sPath As String) As String
    With CreateObject("Scripting.FileSystemObject")
        CanonicalPath = .GetAbsolutePathName(sPath)
    End With
End Function

Public Function PathDifference(sBase As String, sFolder As String) As String
    Dim vBase           As Variant
    Dim vFolder         As Variant
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    If LCase$(Left$(sBase, 2)) <> LCase$(Left$(sFolder, 2)) Then
        PathDifference = sFolder
    Else
        vBase = Split(sBase, "\")
        vFolder = Split(sFolder, "\")
        For lIdx = 0 To UBound(vFolder)
            If lIdx <= UBound(vBase) Then
                If LCase$(vBase(lIdx)) <> LCase$(vFolder(lIdx)) Then
                    Exit For
                End If
            Else
                Exit For
            End If
        Next
        If lIdx > UBound(vBase) Then
'            PathDifference = "."
        Else
            For lJdx = lIdx To UBound(vBase)
                PathDifference = PathDifference & IIf(LenB(PathDifference) <> 0, "\", vbNullString) & ".."
            Next
        End If
        For lJdx = lIdx To UBound(vFolder)
            PathDifference = PathDifference & IIf(LenB(PathDifference) <> 0, "\", vbNullString) & vFolder(lJdx)
        Next
    End If
End Function

Public Function PathMerge(sBase As String, sFolder As String) As String
    If Mid$(sFolder, 2, 1) = ":" Or Left$(sFolder, 2) = "\\" Then
        PathMerge = sFolder
    ElseIf Left$(sFolder, 1) = "\" Then
        PathMerge = Left$(sBase, 2) & sFolder
    Else
        PathMerge = PathCombine(sBase, sFolder)
    End If
    PathMerge = CanonicalPath(PathMerge)
End Function

Public Function PathCombine(sPath As String, sFile As String) As String
    PathCombine = sPath & IIf(LenB(sPath) <> 0 And Right$(sPath, 1) <> "\" And LenB(sFile) <> 0, "\", vbNullString) & sFile
End Function

