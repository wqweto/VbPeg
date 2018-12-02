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

'--- for CreateFile
Private Const GENERIC_WRITE                 As Long = &H40000000
Private Const OPEN_EXISTING                 As Long = 3
Private Const FILE_SHARE_READ               As Long = &H1
'--- for VirtualProtect
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function ApiSysAllocString Lib "oleaut32" Alias "SysAllocString" (ByVal Ptr As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function IsTextUnicode Lib "advapi32" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long
Public Declare Function EmptyLongArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbLong, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As Long()
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByVal lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_VERSION           As String = "0.4.1"

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
    Dim sOutFile        As String
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
    Set m_oOpt = GetOpt(vArgs, "o:set")
    If Not m_oOpt.Item("-nologo") And Not m_oOpt.Item("-q") Then
        ConsoleError App.ProductName & " " & STR_VERSION & " (c) 2018 by wqweto@gmail.com (" & m_oParser.ParserVersion & ")" & vbCrLf & vbCrLf
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
            "  -emit-tree      output parse tree" & vbCrLf & _
            "  -emit-ir        output intermediate represetation" & vbCrLf & _
            "  -set NAME=VALUE set or modify grammar setting NAME to VALUE" & vbCrLf & _
            "  -q              in quiet operation outputs only errors" & vbCrLf & _
            "  -nologo         suppress startup banner" & vbCrLf & _
            "  -allrules       output all rules (don't skip unused)" & vbCrLf & _
            "  -trace          trace in_file.peg parsing as performed by generator" & vbCrLf & vbCrLf & _
            "If no -emit-xxx is used emits VB6 code. If no -o is used writes result to console." & vbCrLf
        If m_oOpt.Item("numarg") = 0 Then
            Process = 100
        End If
        Exit Function
    End If
    sOutFile = m_oOpt.Item("-o")
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
        If LenB(m_oParser.LastError) <> 0 Then
            ConsoleError "%2: %3: %1" & vbCrLf, m_oParser.LastError, Join(oTree.CalcLine(m_oParser.LastOffset + 1), ":"), IIf(lOffset = 0, "error", "warning")
        End If
        If Not m_oParser.GetParseErrors() Is Nothing Then
            For Each vElem In m_oParser.GetParseErrors()
                ConsoleError "%2: %3: %1" & vbCrLf, At(vElem, 0), Join(oTree.CalcLine(At(vElem, 1)), ":"), IIf(lOffset = 0, "error", "warning")
            Next
        End If
        If lOffset = 0 Then
            Process = 1
            Exit Function
        End If
        lIdx = lIdx + 1
    Loop
    Set cWarnings = Nothing
    If Not oTree.CheckTree(cWarnings) Then
        ConsoleError "%1" & vbCrLf, oTree.LastError
        Process = 2
        Exit Function
    ElseIf Not cWarnings Is Nothing Then
        For Each vElem In cWarnings
            ConsoleError "%2: %3: %1" & vbCrLf, At(vElem, 0), Join(oTree.CalcLine(At(vElem, 1)), ":"), IIf(lOffset = 0, "error", "warning")
        Next
    End If
    Set cWarnings = Nothing
    If Not oTree.OptimizeTree(cWarnings) Then
        ConsoleError "Optimize failed: %1" & vbCrLf, oTree.LastError
        Process = 3
        Exit Function
    ElseIf Not cWarnings Is Nothing Then
        For Each vElem In cWarnings
            ConsoleError "%2: %3: %1" & vbCrLf, At(vElem, 0), Join(oTree.CalcLine(At(vElem, 1)), ":"), IIf(lOffset = 0, "error", "warning")
        Next
    End If
    For lIdx = 0 To m_oOpt.Item("#set")
        vElem = Split2(m_oOpt.Item("-set" & IIf(lIdx > 0, lIdx, vbNullString)), "=")
        If LenB(At(vElem, 0)) <> 0 Then
            oTree.SettingValue(At(vElem, 0)) = At(vElem, 1)
        End If
    Next
    If LenB(oTree.SettingValue(STR_SETTING_MODULENAME)) = 0 And LenB(sOutFile) <> 0 Then
        oTree.SettingValue(STR_SETTING_MODULENAME) = GetFilePart(sOutFile)
    End If
    If m_oOpt.Item("-emit-tree") Then
        sOutput = oTree.DumpParseTree
    Else
        Set oIR = New cIR
        If Not oIR.CodeGen(oTree, m_oOpt.Item("-allrules")) Then
            ConsoleError "Failed codegen: %1" & vbCrLf, oIR.LastError
            Process = 4
            Exit Function
        End If
        If m_oOpt.Item("-emit-ir") Then
            sOutput = oIR.DumpIrTree
        Else
            If Not oIR.EmitCode(sOutput) Then
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
    If LenB(sOutFile) > 0 Then
        '--- fix output file extension if not supplied
        If Right$(sOutFile, 1) <> ":" Then
            If InStrRev(sOutFile, "\") >= InStrRev(sOutFile, ".") Then
                sOutFile = sOutFile & IIf(C_Bool(oTree.SettingValue(STR_SETTING_PUBLIC)) Or C_Bool(oTree.SettingValue(STR_SETTING_PRIVATE)), ".cls", ".bas")
            End If
        End If
        SetFileLen sOutFile, Len(sOutput)
        nFile = FreeFile
        Open sOutFile For Binary Access Write Shared As nFile
        Put nFile, , sOutput
        Close nFile
        If Not m_oOpt.Item("-q") Then
            ConsoleError "File " & sOutFile & " emitted successfully" & vbCrLf
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
        If lAction = 0 Then
            lLevel = lLevel + 1
        Else
            sText = m_oParser.Contents(lOffset, TEXT_LEN)
            If InStr(sText, vbCr) > 0 Then
                sText = Left$(sText, InStr(sText, vbCr) - 1)
            End If
            If Len(sText) < TEXT_LEN Then
                sText = sText & String$(TEXT_LEN - Len(sText), "~")
            End If
            sLine = Join(oUserData.CalcLine(lOffset), ":")
            If Len(sLine) - InStr(sLine, ":") < LINE_LEN Then
                sLine = sLine & Space$(LINE_LEN - Len(sLine) + InStr(sLine, ":"))
            End If
            If lAction = 1 Then
                ConsoleError "%1|%2|%3?%4" & vbCrLf, sLine, sText, Space$(lLevel * 2), sRule
                lLevel = lLevel + 1
            Else
                Debug.Assert lLevel > 0
                lLevel = lLevel - 1
                If lAction = 2 Then
                    Const FOREGROUND_GREEN As Long = &H2
                    Const FOREGROUND_MASK As Long = &HF
                    ConsoleColorError FOREGROUND_GREEN, FOREGROUND_MASK, "%1|%2|%3=%4" & vbCrLf, sLine, sText, Space$(lLevel * 2), sRule
                ElseIf lAction = 3 Then
                    ConsoleError "%1|%2|%3!%4" & vbCrLf, sLine, sText, Space$(lLevel * 2), sRule
                Else
                    ConsoleError "Trace error: lAction=" & lAction & vbCrLf
                End If
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
    Dim sValue          As String

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
                            sValue = Mid$(At(vArgs, lIdx), Len(vElem) + 3)
                        ElseIf Len(At(vArgs, lIdx)) > Len(vElem) + 1 Then
                            sValue = Mid$(At(vArgs, lIdx), Len(vElem) + 2)
                        ElseIf LenB(At(vArgs, lIdx + 1)) <> 0 Then
                            sValue = At(vArgs, lIdx + 1)
                            lIdx = lIdx + 1
                        Else
                            .Item("error") = "Option -" & vElem & " requires an argument"
                        End If
                        If Not .Exists("-" & vElem) Then
                            .Item("-" & vElem) = sValue
                        Else
                            .Item("#" & vElem) = .Item("#" & vElem) + 1
                            .Item("-" & vElem & .Item("#" & vElem)) = sValue
                        End If
                        GoTo Continue
                    End If
                Next
                .Item("-" & Mid$(At(vArgs, lIdx), 2)) = True
            Case Else
                .Item("numarg") = .Item("numarg") + 1
                .Item("arg" & .Item("numarg")) = At(vArgs, lIdx)
            End Select
Continue:
        Next
    End With
    Set GetOpt = oRetVal
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
    If IsArray(vArray) Then
        If lIdx >= LBound(vArray) And lIdx <= UBound(vArray) Then
            At = vArray(lIdx)
        End If
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

Public Function Split2(sText As String, sDelim As String) As Variant
    Dim lPos            As Long
    
    lPos = InStr(sText, sDelim)
    If lPos > 0 Then
        Split2 = Array(Left$(sText, lPos - 1), Mid$(sText, lPos + Len(sDelim)))
    Else
        Split2 = Array(sText)
    End If
End Function

Public Sub PatchMethodProto(ByVal pfn As Long, ByVal lMethodIdx As Long)
    If InIde Then
        '--- note: IDE is not large-address aware
        Call CopyMemory(pfn, ByVal pfn + &H16, 4)
    Else
        Call VirtualProtect(pfn, 12, PAGE_EXECUTE_READWRITE, 0)
    End If
    ' 0: 8B 44 24 04          mov         eax,dword ptr [esp+4]
    ' 4: 8B 00                mov         eax,dword ptr [eax]
    ' 6: FF A0 00 00 00 00    jmp         dword ptr [eax+lMethodIdx*4]
    Call CopyMemory(ByVal pfn, -684575231150992.4725@, 8)
    Call CopyMemory(ByVal (pfn Xor &H80000000) + 8 Xor &H80000000, lMethodIdx * 4, 4)
End Sub
 
Public Function TryGetValue(ByVal oCol As Collection, Index As Variant, RetVal As Variant) As Long
    Const IDX_COLLECTION_ITEM   As Long = 7
    PatchMethodProto AddressOf mdMain.TryGetValue, IDX_COLLECTION_ITEM
    TryGetValue = TryGetValue(oCol, Index, RetVal)
End Function

Public Function SearchCollection(oCol As Collection, Index As Variant, Optional RetVal As Variant) As Boolean
    If Not oCol Is Nothing Then
        SearchCollection = TryGetValue(oCol, Index, RetVal) = 0 ' S_OK
    End If
End Function

