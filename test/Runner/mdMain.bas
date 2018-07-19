Attribute VB_Name = "mdMain"
Option Explicit

Private Const STD_OUTPUT_HANDLE             As Long = -11&
Private Const STD_ERROR_HANDLE              As Long = -12&

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, lpszDst As Any, ByVal cchDstLength As Long) As Long
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function ApiSysAllocString Lib "oleaut32" Alias "SysAllocString" (ByVal Ptr As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function IsTextUnicode Lib "advapi32" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long

Private m_sContents         As String
Private m_laOffsets()       As Long
Private m_sFileName         As String

Private Sub Main()
    Dim oOpt            As Object
    Dim lIdx            As Long
    Dim vResult         As Variant
    Dim lPos            As Long
    
    On Error GoTo EH
    Set oOpt = GetOpt(SplitArgs(Command$))
    For lIdx = 1 To oOpt.Item("numarg")
        lPos = 1
        m_sFileName = oOpt.Item("arg" & lIdx)
        m_sContents = ReadTextFile(m_sFileName)
        m_sFileName = Mid$(m_sFileName, InStrRev(m_sFileName, "\") + 1)
        pvBuildLineInfo m_sContents
        Do While lPos < Len(m_sContents)
            vResult = Empty
            lPos = VbPegMatch(m_sContents, lPos - 1, Result:=vResult)
            If lPos = 0 Then
                ConsolePrint "LastError: %1" & vbCrLf, VbPegLastError
                Exit Do
            End If
            ConsolePrint "Pos: %1", lPos
            If Not IsEmpty(vResult) Then
                ConsolePrint ", Result: %1", C_Str(vResult)
            End If
            If LenB(VbPegLastError) <> 0 Then
                ConsolePrint ", LastError: %1", C_Str(VbPegLastError)
            End If
            ConsolePrint vbCrLf
        Loop
    Next
    Exit Sub
EH:
    ConsoleError "Critical: " & Err.Description & vbCrLf
End Sub

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
    Const BOM_UTF       As String = "?»?"   '--- "\xEF\xBB\xBF"
    Const BOM_UNICODE   As String = "??"    '--- "\xFF\xFE"
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

Public Function C_Str(vValue As Variant) As String
    On Error GoTo QH
    C_Str = CStr(vValue)
QH:
End Function

Public Function ConsoleTrace(ByVal lOffset As Long, sRule As String, ByVal lAction As Long, vUserData As Variant) As Boolean
    Const LINE_LEN      As Long = 8
    Const TEXT_LEN      As Long = 60
    Static lLevel       As Long
    Dim sText           As String
    Dim sLine           As String
    
        sText = Mid$(m_sContents, lOffset, TEXT_LEN)
        If InStr(sText, vbCr) > 0 Then
            sText = Left$(sText, InStr(sText, vbCr) - 1)
        End If
        sText = Replace(sText, vbLf, " ")
        If Len(sText) < TEXT_LEN Then
            sText = sText & Space$(TEXT_LEN - Len(sText))
        End If
        sLine = Join(CalcLine(lOffset), ":")
        If Len(sLine) - InStr(sLine, ":") < LINE_LEN Then
            sLine = sLine & Space$(LINE_LEN - Len(sLine) + InStr(sLine, ":"))
        End If
        If lAction = 1 Then
            ConsolePrint "%1|%2|%3?%4" & vbCrLf, sLine, sText, Space$(lLevel * 2), sRule
            lLevel = lLevel + 1
        Else
            If lLevel > 0 Then
                lLevel = lLevel - 1
            End If
            If lAction = 2 Then
                ConsolePrint "%1|%2|%3=%4" & vbCrLf, sLine, sText, Space$(lLevel * 2), sRule
            Else
                ConsolePrint "%1|%2|%3!%4" & vbCrLf, sLine, sText, Space$(lLevel * 2), sRule
            End If
        End If
End Function

Private Sub pvBuildLineInfo(sSubject As String)
    Dim lIdx            As Long
    
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "\r?\n"
        With .Execute(sSubject)
            ReDim m_laOffsets(0 To .Count) As Long
            For lIdx = 0 To .Count - 1
                With .Item(lIdx)
                    m_laOffsets(lIdx + 1) = .FirstIndex + .Length
                End With
            Next
        End With
    End With
End Sub

Public Function CalcLine(ByVal lOffset As Long) As Variant
    Dim lLower          As Long
    Dim lUpper          As Long
    Dim lMiddle         As Long
    
    lUpper = UBound(m_laOffsets)
    Do While lLower < lUpper
        lMiddle = (lLower + lUpper + 1) \ 2
        If m_laOffsets(lMiddle) < lOffset Then
            lLower = lMiddle
        Else
            lUpper = lMiddle - 1
        End If
    Loop
    CalcLine = Array(m_sFileName, lUpper + 1, lOffset - m_laOffsets(lUpper))
End Function
