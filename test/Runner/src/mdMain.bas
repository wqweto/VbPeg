Attribute VB_Name = "mdMain"
'=========================================================================
'
' VbPeg (c) 2018 by wqweto@gmail.com
'
' PEG parser generator for VB6
'
' mdMain.bas - Test runner global functions
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

'--- for VirtualProtect
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function ApiSysAllocString Lib "oleaut32" Alias "SysAllocString" (ByVal Ptr As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function IsTextUnicode Lib "advapi32" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_sContents         As String
Private m_laOffsets()       As Long
Private m_sFileName         As String

'=========================================================================
' Functions
'=========================================================================

Private Sub Main()
    Dim oOpt            As Object
    Dim oParser         As Object
    Dim lIdx            As Long
    Dim lPos            As Long
    Dim vResult         As Variant
    
    On Error GoTo EH
    Set oOpt = GetOpt(SplitArgs(Command$))
    Set oParser = CreateObjectPrivate(oOpt.Item("arg1"))
    For lIdx = 2 To oOpt.Item("numarg")
        lPos = 1
        m_sFileName = oOpt.Item("arg" & lIdx)
        m_sContents = ReadTextFile(m_sFileName)
        m_sFileName = Mid$(m_sFileName, InStrRev(m_sFileName, "\") + 1)
        pvBuildLineInfo m_sContents
        Do While lPos <= Len(m_sContents)
            vResult = Empty
            lPos = oParser.Match(m_sContents, lPos - 1, UserData:=oParser, Result:=vResult)
            If lPos = 0 Then
                ConsolePrint "LastError: %1, LastOffset: %2" & vbCrLf, oParser.LastError, oParser.LastOffset
                Exit Do
            End If
            ConsolePrint "Pos: %1", lPos
            If Not IsEmpty(vResult) Then
                If TypeName(vResult) = "Dictionary" Then
                    ConsolePrint ", Result: %1", JsonDump(vResult)
                Else
                    ConsolePrint ", Result: %1", C_Str(vResult)
                End If
            End If
            If LenB(oParser.LastError) <> 0 Then
                ConsolePrint ", Warning: %1", oParser.LastError
            End If
            ConsolePrint vbCrLf
        Loop
    Next
    Exit Sub
EH:
    ConsoleError "Critical: " & Err.Description & vbCrLf
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

Public Function C_Str(Value As Variant) As String
    On Error GoTo QH
    C_Str = CStr(Value)
QH:
End Function

Public Function C_Lng(Value As Variant) As Long
    On Error GoTo QH
    C_Lng = CLng(Value)
QH:
End Function

Public Function C_Dbl(Value As Variant) As Double
    On Error GoTo QH
    C_Dbl = CDbl(Value)
QH:
End Function

Public Function C_Bool(Value As Variant) As Boolean
    On Error GoTo QH
    C_Bool = CBool(Value)
QH:
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

Public Function Zn(sText As String, Optional IfEmptyString As Variant = Null) As Variant
    Zn = IIf(LenB(sText) = 0, IfEmptyString, sText)
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

Public Sub AssignVariant(vDest As Variant, vSrc As Variant)
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
End Sub

Public Sub PatchMethodProto(ByVal pfn As Long, ByVal lMethodIdx As Long)
    If App.LogMode = 0 Then
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

