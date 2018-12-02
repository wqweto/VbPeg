VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cZig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Auto-generated on 3.12.2018 1:29:31
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Const LOCALE_USER_DEFAULT           As Long = &H400
Private Const NORM_IGNORECASE               As Long = 1
Private Const CSTR_EQUAL                    As Long = 2

Private Declare Function CompareStringW Lib "kernel32" (ByVal Locale As Long, ByVal dwCmpFlags As Long, lpString1 As Any, ByVal cchCount1 As Long, lpString2 As Any, ByVal cchCount2 As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

Private Type SAFEARRAY1D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As Long
    cElements           As Long
    lLbound             As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const LNG_MAXINT            As Long = 2 ^ 31 - 1

'= generated enum ========================================================

Private Type UcsParserThunkType
    Action              As Long
    CaptureBegin        As Long
    CaptureEnd          As Long
End Type

Private Type UcsParserType
    Contents            As String
    BufArray            As SAFEARRAY1D
    BufData()           As Integer
    BufPos              As Long
    BufSize             As Long
    ThunkData()         As UcsParserThunkType
    ThunkPos            As Long
    CaptureBegin        As Long
    CaptureEnd          As Long
    LastExpected        As String
    LastError           As String
    LastBufPos          As Long
    UserData            As Variant
    VarResult           As Variant
End Type

Private ctx                     As UcsParserType

'=========================================================================
' Properties
'=========================================================================

Property Get LastError() As String
    LastError = ctx.LastError
End Property

Property Get LastOffset() As Long
    LastOffset = ctx.LastBufPos + 1
End Property

Property Get ParserVersion() As String
    ParserVersion = "3.12.2018 1:29:31"
End Property

Property Get Contents(Optional ByVal lOffset As Long = 1, Optional ByVal lSize As Long = LNG_MAXINT) As String
    Contents = Mid$(ctx.Contents, lOffset, lSize)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Match(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant, Optional Result As Variant) As Long
    If BeginMatch(sSubject, StartPos, UserData) Then
        If ParseRoot() Then
            Match = EndMatch(Result)
        Else
            With ctx
                If LenB(.LastError) = 0 Then
                    If LenB(.LastExpected) = 0 Then
                        .LastError = "Fail"
                    Else
                        .LastError = "Expected " & Join(Split(Mid$(.LastExpected, 2, Len(.LastExpected) - 2), vbNullChar), " or ")
                    End If
                End If
            End With
        End If
    End If
End Function

Public Function BeginMatch(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant) As Boolean
    With ctx
        .LastBufPos = 0
        If LenB(sSubject) = 0 Then
            .LastError = "Cannot match empty input"
            Exit Function
        End If
        Call CopyMemory(ByVal VarPtr(.Contents), ByVal VarPtr(sSubject), 4)
        With .BufArray
            .cDims = 1
            .cbElements = 2
            .fFeatures = 1 ' FADF_AUTO
            .pvData = StrPtr(sSubject)
            .cElements = Len(sSubject) + 2 '-- look-ahead chars
        End With
        Call CopyMemory(ByVal ArrPtr(.BufData), VarPtr(.BufArray), 4)
        .BufPos = StartPos
        .BufSize = Len(sSubject)
        ReDim .ThunkData(0 To 4) As UcsParserThunkType
        .ThunkPos = 0
        .CaptureBegin = 0
        .CaptureEnd = 0
        If IsObject(UserData) Then
            Set .UserData = UserData
        Else
            .UserData = UserData
        End If
    End With
    BeginMatch = True
End Function

Public Function EndMatch(Optional Result As Variant) As Long
    With ctx
        If IsObject(.VarResult) Then
            Set Result = .VarResult
        Else
            Result = .VarResult
        End If
        EndMatch = .BufPos + 1
        Call CopyMemory(ByVal VarPtr(.Contents), 0&, 4)
        Call CopyMemory(ByVal ArrPtr(.BufData), 0&, 4)
        .BufPos = 0
        .BufSize = 0
        Erase .ThunkData
        .ThunkPos = 0
        .CaptureBegin = 0
        .CaptureEnd = 0
    End With
End Function

Private Function pvMatchString(sText As String, Optional ByVal CmpFlags As Long) As Boolean
    With ctx
        If .BufPos + Len(sText) <= .BufSize Then
            pvMatchString = CompareStringW(LOCALE_USER_DEFAULT, CmpFlags, ByVal StrPtr(sText), Len(sText), .BufData(.BufPos), Len(sText)) = CSTR_EQUAL
        End If
    End With
End Function

Private Sub pvSetAdvance()
    With ctx
        If .BufPos > .LastBufPos Then
            .LastExpected = vbNullString
            .LastError = vbNullString
            .LastBufPos = .BufPos
        End If
    End With
End Sub

Private Sub pvSetExpected(sText As String)
    With ctx
        If LenB(.LastExpected) = 0 Or .BufPos > .LastBufPos Then
            .LastExpected = vbNullChar & sText & vbNullChar
            .LastBufPos = .BufPos
        ElseIf .BufPos = .LastBufPos Then
            If InStr(.LastExpected, vbNullChar & sText & vbNullChar) = 0 Then
                .LastExpected = .LastExpected & sText & vbNullChar
            End If
        End If
    End With
End Sub

'= generated functions ===================================================

Public Function ParseRoot() As Boolean
    Call Parseskip
    Call ParseContainerMembers
    If Parseeof() Then
        ParseRoot = True
    End If
End Function

Private Sub Parseskip()
    Dim p1101 As Long

    With ctx
        Do
            p1101 = .BufPos
            Select Case .BufData(.BufPos)
            Case 32, 13, 10                         ' [ \r\n]
                .BufPos = .BufPos + 1
            Case Else
                If Not (Parseline_comment()) Then
                    .BufPos = p1101
                    .BufPos = p1101
                    Exit Do
                End If
            End Select
        Loop
        Call pvSetAdvance
    End With
End Sub

Private Sub ParseContainerMembers()
    Dim p17 As Long
    Dim p19 As Long
    Dim p25 As Long
    Dim p33 As Long

    With ctx
        p17 = .BufPos
        If ParseTestDecl() Then
            Call ParseContainerMembers
        Else
            .BufPos = p17
            If ParseTopLevelComptime() Then
                Call ParseContainerMembers
            Else
                .BufPos = p17
                p19 = .BufPos
                If Not (ParseKEYWORD_pub()) Then
                    .BufPos = p19
                End If
                If ParseTopLevelDecl() Then
                    Call ParseContainerMembers
                Else
                    .BufPos = p17
                    p25 = .BufPos
                    If Not (ParseKEYWORD_pub()) Then
                        .BufPos = p25
                    End If
                    If ParseContainerField() Then
                        If ParseCOMMA() Then
                            Call ParseContainerMembers
                        Else
                            .BufPos = p17
                            p33 = .BufPos
                            If Not (ParseKEYWORD_pub()) Then
                                .BufPos = p33
                            End If
                            If Not (ParseContainerField()) Then
                                .BufPos = p17
                            End If
                        End If
                    Else
                        .BufPos = p17
                        p33 = .BufPos
                        If Not (ParseKEYWORD_pub()) Then
                            .BufPos = p33
                        End If
                        If Not (ParseContainerField()) Then
                            .BufPos = p17
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function Parseeof() As Boolean
    Dim p1044 As Long

    With ctx
        p1044 = .BufPos
        If Not (.BufPos < .BufSize) Then
            .BufPos = p1044
            Parseeof = True
        End If
    End With
End Function

Private Function ParseTestDecl() As Boolean
    If ParseKEYWORD_test() Then
        If ParseSTRINGLITERAL() Then
            If ParseBlock() Then
                ParseTestDecl = True
            End If
        End If
    End If
End Function

Private Function ParseTopLevelComptime() As Boolean
    If ParseKEYWORD_comptime() Then
        If ParseBlockExpr() Then
            ParseTopLevelComptime = True
        End If
    End If
End Function

Private Function ParseKEYWORD_pub() As Boolean
    With ctx
        If pvMatchString("pub") Then                ' "pub"
            .BufPos = .BufPos + 3
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_pub = True
            End If
        End If
    End With
End Function

Private Function ParseTopLevelDecl() As Boolean
    Dim p78 As Long
    Dim p57 As Long
    Dim p54 As Long
    Dim p67 As Long
    Dim p73 As Long
    Dim p70 As Long

    With ctx
        p78 = .BufPos
        p57 = .BufPos
        If ParseKEYWORD_export() Then
            GoTo L35
        Else
            .BufPos = p57
        End If
        If ParseKEYWORD_extern() Then
            p54 = .BufPos
            If Not (ParseSTRINGLITERAL()) Then
                .BufPos = p54
            End If
            GoTo L35
        Else
            .BufPos = p57
        End If
        If ParseKEYWORD_inline() Then
            GoTo L35
        Else
            .BufPos = p57
        End If
        .BufPos = p57
L35:
        If ParseFnProto() Then
            p67 = .BufPos
            If ParseSEMICOLON() Then
                ParseTopLevelDecl = True
                Exit Function
            Else
                .BufPos = p67
            End If
            If ParseBlock() Then
                ParseTopLevelDecl = True
                Exit Function
            Else
                .BufPos = p67
            End If
            .BufPos = p78
        Else
            .BufPos = p78
        End If
        p73 = .BufPos
        If ParseKEYWORD_export() Then
            GoTo L37
        Else
            .BufPos = p73
        End If
        If ParseKEYWORD_extern() Then
            p70 = .BufPos
            If Not (ParseSTRINGLITERAL()) Then
                .BufPos = p70
            End If
            GoTo L37
        Else
            .BufPos = p73
        End If
        .BufPos = p73
L37:
        If ParseVarDecl() Then
            ParseTopLevelDecl = True
            Exit Function
        Else
            .BufPos = p78
        End If
        If ParseKEYWORD_use() Then
            If ParseExpr() Then
                If ParseSEMICOLON() Then
                    ParseTopLevelDecl = True
                    Exit Function
                Else
                    .BufPos = p78
                End If
            Else
                .BufPos = p78
            End If
        Else
            .BufPos = p78
        End If
    End With
End Function

Private Function ParseContainerField() As Boolean
    Dim p138 As Long
    Dim p143 As Long

    With ctx
        If ParseIDENTIFIER() Then
            p138 = .BufPos
            If ParseCOLON() Then
                If Not (ParseTypeExpr()) Then
                    .BufPos = p138
                End If
            Else
                .BufPos = p138
            End If
            p143 = .BufPos
            If ParseEQUAL() Then
                If Not (ParseExpr()) Then
                    .BufPos = p143
                End If
            Else
                .BufPos = p143
            End If
            ParseContainerField = True
        End If
    End With
End Function

Private Function ParseCOMMA() As Boolean
    With ctx
        If .BufData(.BufPos) = 44 Then              ' ","
            .BufPos = .BufPos + 1
            Call Parseskip
            Call pvSetAdvance
            ParseCOMMA = True
        End If
    End With
End Function

Private Function ParseKEYWORD_test() As Boolean
    With ctx
        If pvMatchString("test") Then               ' "test"
            .BufPos = .BufPos + 4
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_test = True
            End If
        End If
    End With
End Function

Private Function ParseSTRINGLITERAL() As Boolean
    Dim p1246 As Long
    Dim p1239 As Long

    With ctx
        p1246 = .BufPos
        If .BufData(.BufPos) = 99 Then              ' "c"
            .BufPos = .BufPos + 1
        End If
        If .BufData(.BufPos) = 34 Then              ' "\""
            .BufPos = .BufPos + 1
            Do
                p1239 = .BufPos
                If Not (Parsestring_char()) Then
                    .BufPos = p1239
                    Exit Do
                End If
            Loop
            If .BufData(.BufPos) = 34 Then          ' "\""
                .BufPos = .BufPos + 1
                Call Parseskip
                Call pvSetAdvance
                ParseSTRINGLITERAL = True
                Exit Function
            Else
                .BufPos = p1246
            End If
        Else
            .BufPos = p1246
        End If
        If Parseline_string() Then
            Call Parseskip
            Call pvSetAdvance
            ParseSTRINGLITERAL = True
            Exit Function
        Else
            .BufPos = p1246
        End If
        If Parseline_cstring() Then
            Call Parseskip
            Call pvSetAdvance
            ParseSTRINGLITERAL = True
            Exit Function
        Else
            .BufPos = p1246
        End If
        Call pvSetExpected("STRINGLITERAL")
    End With
End Function

Private Function ParseBlock() As Boolean
    Dim p396 As Long

    With ctx
        If ParseLBRACE() Then
            Do
                p396 = .BufPos
                If Not (ParseStatement()) Then
                    .BufPos = p396
                    Exit Do
                End If
            Loop
            If ParseRBRACE() Then
                ParseBlock = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_comptime() As Boolean
    With ctx
        If pvMatchString("comptime") Then           ' "comptime"
            .BufPos = .BufPos + 8
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_comptime = True
            End If
        End If
    End With
End Function

Private Function ParseBlockExpr() As Boolean
    Dim p261 As Long

    With ctx
        p261 = .BufPos
        If Not (ParseBlockLabel()) Then
            .BufPos = p261
        End If
        If ParseBlock() Then
            ParseBlockExpr = True
        End If
    End With
End Function

Private Function ParseKEYWORD_export() As Boolean
    With ctx
        If pvMatchString("export") Then             ' "export"
            .BufPos = .BufPos + 6
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_export = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_extern() As Boolean
    With ctx
        If pvMatchString("extern") Then             ' "extern"
            .BufPos = .BufPos + 6
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_extern = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_inline() As Boolean
    With ctx
        If pvMatchString("inline") Then             ' "inline"
            .BufPos = .BufPos + 6
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_inline = True
            End If
        End If
    End With
End Function

Private Function ParseFnProto() As Boolean
    Dim p86 As Long
    Dim p92 As Long
    Dim p101 As Long
    Dim p104 As Long
    Dim p107 As Long
    Dim p113 As Long

    With ctx
        p86 = .BufPos
        If Not (ParseFnCC()) Then
            .BufPos = p86
        End If
        If ParseKEYWORD_fn() Then
            p92 = .BufPos
            If Not (ParseIDENTIFIER()) Then
                .BufPos = p92
            End If
            If ParseLPAREN() Then
                Call ParseParamDeclList
                If ParseRPAREN() Then
                    p101 = .BufPos
                    If Not (ParseByteAlign()) Then
                        .BufPos = p101
                    End If
                    p104 = .BufPos
                    If Not (ParseSection()) Then
                        .BufPos = p104
                    End If
                    p107 = .BufPos
                    If Not (ParseEXCLAMATIONMARK()) Then
                        .BufPos = p107
                    End If
                    p113 = .BufPos
                    If ParseKEYWORD_var() Then
                        ParseFnProto = True
                        Exit Function
                    Else
                        .BufPos = p113
                    End If
                    If ParseTypeExpr() Then
                        ParseFnProto = True
                        Exit Function
                    Else
                        .BufPos = p113
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseSEMICOLON() As Boolean
    With ctx
        If .BufData(.BufPos) = 59 Then              ' ";"
            .BufPos = .BufPos + 1
            Call Parseskip
            Call pvSetAdvance
            ParseSEMICOLON = True
        End If
    End With
End Function

Private Function ParseVarDecl() As Boolean
    Dim p117 As Long
    Dim p123 As Long
    Dim p125 As Long
    Dim p127 As Long
    Dim p132 As Long

    With ctx
        p117 = .BufPos
        If ParseKEYWORD_const() Then
            If ParseIDENTIFIER() Then
                p123 = .BufPos
                If ParseCOLON() Then
                    If Not (ParseTypeExpr()) Then
                        .BufPos = p123
                    End If
                Else
                    .BufPos = p123
                End If
                p125 = .BufPos
                If Not (ParseByteAlign()) Then
                    .BufPos = p125
                End If
                p127 = .BufPos
                If Not (ParseSection()) Then
                    .BufPos = p127
                End If
                p132 = .BufPos
                If ParseEQUAL() Then
                    If Not (ParseExpr()) Then
                        .BufPos = p132
                    End If
                Else
                    .BufPos = p132
                End If
                If ParseSEMICOLON() Then
                    ParseVarDecl = True
                    Exit Function
                End If
            End If
        Else
            .BufPos = p117
        End If
        If ParseKEYWORD_var() Then
            If ParseIDENTIFIER() Then
                p123 = .BufPos
                If ParseCOLON() Then
                    If Not (ParseTypeExpr()) Then
                        .BufPos = p123
                    End If
                Else
                    .BufPos = p123
                End If
                p125 = .BufPos
                If Not (ParseByteAlign()) Then
                    .BufPos = p125
                End If
                p127 = .BufPos
                If Not (ParseSection()) Then
                    .BufPos = p127
                End If
                p132 = .BufPos
                If ParseEQUAL() Then
                    If Not (ParseExpr()) Then
                        .BufPos = p132
                    End If
                Else
                    .BufPos = p132
                End If
                If ParseSEMICOLON() Then
                    ParseVarDecl = True
                    Exit Function
                End If
            End If
        Else
            .BufPos = p117
        End If
    End With
End Function

Private Function ParseKEYWORD_use() As Boolean
    With ctx
        If pvMatchString("use") Then                ' "use"
            .BufPos = .BufPos + 3
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_use = True
            End If
        End If
    End With
End Function

Private Function ParseExpr() As Boolean
    Dim p273 As Long

    With ctx
        Do
            p273 = .BufPos
            If Not (ParseKEYWORD_try()) Then
                .BufPos = p273
                Exit Do
            End If
        Loop
        If ParseBoolOrExpr() Then
            ParseExpr = True
        End If
    End With
End Function

Private Function ParseFnCC() As Boolean
    Dim p685 As Long
    Dim p692 As Long

    With ctx
        p685 = .BufPos
        If ParseKEYWORD_nakedcc() Then
            ParseFnCC = True
            Exit Function
        Else
            .BufPos = p685
        End If
        If ParseKEYWORD_stdcallcc() Then
            ParseFnCC = True
            Exit Function
        Else
            .BufPos = p685
        End If
        If ParseKEYWORD_extern() Then
            ParseFnCC = True
            Exit Function
        Else
            .BufPos = p685
        End If
        If ParseKEYWORD_async() Then
            p692 = .BufPos
            If ParseLARROW() Then
                If ParseTypeExpr() Then
                    If Not (ParseRARROW()) Then
                        .BufPos = p692
                    End If
                Else
                    .BufPos = p692
                End If
            Else
                .BufPos = p692
            End If
            ParseFnCC = True
            Exit Function
        Else
            .BufPos = p685
        End If
    End With
End Function

Private Function ParseKEYWORD_fn() As Boolean
    With ctx
        If .BufData(.BufPos) = 102 And .BufData(.BufPos + 1) = 110 Then ' "fn"
            .BufPos = .BufPos + 2
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_fn = True
            End If
        End If
    End With
End Function

Private Function ParseIDENTIFIER() As Boolean
    Dim p1269 As Long
    Dim p1252 As Long
    Dim e1252 As String
    Dim p1258 As Long
    Dim p1255 As Long
    Dim p1264 As Long

    With ctx
        p1269 = .BufPos
        p1252 = .BufPos
        e1252 = .LastExpected
        If Parsekeyword() Then
            .BufPos = p1269
        Else
            .BufPos = p1252
            .LastExpected = e1252
            p1258 = .BufPos
            If .BufData(.BufPos) = 99 Then          ' "c"
                .BufPos = .BufPos + 1
                p1255 = .BufPos
                Select Case .BufData(.BufPos)
                Case 34, 92                         ' ["\\]
                    .BufPos = p1258
                Case Else
                    .BufPos = p1255
                    Do
                        Select Case .BufData(.BufPos)
                        Case 65 To 90, 97 To 122, 48 To 57, 95 ' [A-Za-z0-9_]
                            .BufPos = .BufPos + 1
                        Case Else
                            Exit Do
                        End Select
                    Loop
                    Call Parseskip
                    Call pvSetAdvance
                    ParseIDENTIFIER = True
                    Exit Function
                End Select
            End If
            Select Case .BufData(.BufPos)
            Case 65 To 90, 97, 98, 100 To 122, 95   ' [A-Zabd-z_]
                .BufPos = .BufPos + 1
                Do
                    Select Case .BufData(.BufPos)
                    Case 65 To 90, 97 To 122, 48 To 57, 95 ' [A-Za-z0-9_]
                        .BufPos = .BufPos + 1
                    Case Else
                        Exit Do
                    End Select
                Loop
                Call Parseskip
                Call pvSetAdvance
                ParseIDENTIFIER = True
                Exit Function
            End Select
            .BufPos = p1269
        End If
        If .BufData(.BufPos) = 64 And .BufData(.BufPos + 1) = 34 Then ' "@\""
            .BufPos = .BufPos + 2
            Do
                p1264 = .BufPos
                If Not (Parsestring_char()) Then
                    .BufPos = p1264
                    Exit Do
                End If
            Loop
            If .BufData(.BufPos) = 34 Then          ' "\""
                .BufPos = .BufPos + 1
                Call Parseskip
                Call pvSetAdvance
                ParseIDENTIFIER = True
                Exit Function
            Else
                .BufPos = p1269
            End If
        End If
        Call pvSetExpected("IDENTIFIER")
    End With
End Function

Private Function ParseLPAREN() As Boolean
    With ctx
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call Parseskip
            Call pvSetAdvance
            ParseLPAREN = True
        End If
    End With
End Function

Private Sub ParseParamDeclList()
    Dim p1031 As Long
    Dim p1033 As Long

    With ctx
        Do
            p1031 = .BufPos
            If Not (ParseParamDecl()) Then
                .BufPos = p1031
                Exit Do
            End If
            If Not (ParseCOMMA()) Then
                .BufPos = p1031
                Exit Do
            End If
        Loop
        p1033 = .BufPos
        If Not (ParseParamDecl()) Then
            .BufPos = p1033
        End If
    End With
End Sub

Private Function ParseRPAREN() As Boolean
    With ctx
        If .BufData(.BufPos) = 41 Then              ' ")"
            .BufPos = .BufPos + 1
            Call Parseskip
            Call pvSetAdvance
            ParseRPAREN = True
        End If
    End With
End Function

Private Function ParseByteAlign() As Boolean
    If ParseKEYWORD_align() Then
        If ParseLPAREN() Then
            If ParseExpr() Then
                If ParseRPAREN() Then
                    ParseByteAlign = True
                End If
            End If
        End If
    End If
End Function

Private Function ParseSection() As Boolean
    If ParseKEYWORD_section() Then
        If ParseLPAREN() Then
            If ParseExpr() Then
                If ParseRPAREN() Then
                    ParseSection = True
                End If
            End If
        End If
    End If
End Function

Private Function ParseEXCLAMATIONMARK() As Boolean
    Dim p1349 As Long

    With ctx
        If .BufData(.BufPos) = 33 Then              ' "!"
            .BufPos = .BufPos + 1
            p1349 = .BufPos
            If Not (.BufData(.BufPos) = 61) Then    ' "="
                .BufPos = p1349
                Call Parseskip
                Call pvSetAdvance
                ParseEXCLAMATIONMARK = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_var() As Boolean
    With ctx
        If pvMatchString("var") Then                ' "var"
            .BufPos = .BufPos + 3
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_var = True
            End If
        End If
    End With
End Function

Private Function ParseTypeExpr() As Boolean
    Dim p456 As Long

    With ctx
        Do
            p456 = .BufPos
            If Not (ParsePrefixTypeOp()) Then
                .BufPos = p456
                Exit Do
            End If
        Loop
        If ParseErrorUnionExpr() Then
            ParseTypeExpr = True
        End If
    End With
End Function

Private Function ParseKEYWORD_const() As Boolean
    With ctx
        If pvMatchString("const") Then              ' "const"
            .BufPos = .BufPos + 5
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_const = True
            End If
        End If
    End With
End Function

Private Function ParseCOLON() As Boolean
    With ctx
        If .BufData(.BufPos) = 58 Then              ' ":"
            .BufPos = .BufPos + 1
            Call Parseskip
            Call pvSetAdvance
            ParseCOLON = True
        End If
    End With
End Function

Private Function ParseEQUAL() As Boolean
    Dim p1338 As Long

    With ctx
        If .BufData(.BufPos) = 61 Then              ' "="
            .BufPos = .BufPos + 1
            p1338 = .BufPos
            Select Case .BufData(.BufPos)
            Case 62, 61                             ' [>=]
                '--- do nothing
            Case Else
                .BufPos = p1338
                Call Parseskip
                Call pvSetAdvance
                ParseEQUAL = True
            End Select
        End If
    End With
End Function

Private Function ParseStatement() As Boolean
    Dim p154 As Long
    Dim p146 As Long
    Dim p159 As Long

    With ctx
        p154 = .BufPos
        p146 = .BufPos
        If Not (ParseKEYWORD_comptime()) Then
            .BufPos = p146
        End If
        If ParseVarDecl() Then
            ParseStatement = True
            Exit Function
        Else
            .BufPos = p154
        End If
        If ParseKEYWORD_comptime() Then
            If ParseBlockExprStatement() Then
                ParseStatement = True
                Exit Function
            Else
                .BufPos = p154
            End If
        Else
            .BufPos = p154
        End If
        If ParseKEYWORD_suspend() Then
            p159 = .BufPos
            If ParseSEMICOLON() Then
                ParseStatement = True
                Exit Function
            Else
                .BufPos = p159
            End If
            If ParseBlockExprStatement() Then
                ParseStatement = True
                Exit Function
            Else
                .BufPos = p159
            End If
            .BufPos = p154
        Else
            .BufPos = p154
        End If
        If ParseKEYWORD_defer() Then
            If ParseBlockExprStatement() Then
                ParseStatement = True
                Exit Function
            Else
                .BufPos = p154
            End If
        Else
            .BufPos = p154
        End If
        If ParseKEYWORD_errdefer() Then
            If ParseBlockExprStatement() Then
                ParseStatement = True
                Exit Function
            Else
                .BufPos = p154
            End If
        Else
            .BufPos = p154
        End If
        If ParseIfStatement() Then
            ParseStatement = True
            Exit Function
        Else
            .BufPos = p154
        End If
        If ParseLabeledStatement() Then
            ParseStatement = True
            Exit Function
        Else
            .BufPos = p154
        End If
        If ParseSwitchExpr() Then
            ParseStatement = True
            Exit Function
        Else
            .BufPos = p154
        End If
        If ParseAssignExpr() Then
            If ParseSEMICOLON() Then
                ParseStatement = True
                Exit Function
            Else
                .BufPos = p154
            End If
        Else
            .BufPos = p154
        End If
    End With
End Function

Private Function ParseBlockExprStatement() As Boolean
    Dim p260 As Long

    With ctx
        p260 = .BufPos
        If ParseBlockExpr() Then
            ParseBlockExprStatement = True
            Exit Function
        Else
            .BufPos = p260
        End If
        If ParseAssignExpr() Then
            If ParseSEMICOLON() Then
                ParseBlockExprStatement = True
                Exit Function
            Else
                .BufPos = p260
            End If
        Else
            .BufPos = p260
        End If
    End With
End Function

Private Function ParseKEYWORD_suspend() As Boolean
    With ctx
        If pvMatchString("suspend") Then            ' "suspend"
            .BufPos = .BufPos + 7
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_suspend = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_defer() As Boolean
    With ctx
        If pvMatchString("defer") Then              ' "defer"
            .BufPos = .BufPos + 5
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_defer = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_errdefer() As Boolean
    With ctx
        If pvMatchString("errdefer") Then           ' "errdefer"
            .BufPos = .BufPos + 8
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_errdefer = True
            End If
        End If
    End With
End Function

Private Function ParseIfStatement() As Boolean
    Dim p201 As Long
    Dim p188 As Long
    Dim p186 As Long
    Dim p200 As Long
    Dim p196 As Long

    With ctx
        p201 = .BufPos
        If ParseIfPrefix() Then
            If ParseBlockExpr() Then
                p188 = .BufPos
                If ParseKEYWORD_else() Then
                    p186 = .BufPos
                    If Not (ParsePayload()) Then
                        .BufPos = p186
                    End If
                    If Not (ParseStatement()) Then
                        .BufPos = p188
                    End If
                Else
                    .BufPos = p188
                End If
                ParseIfStatement = True
                Exit Function
            Else
                .BufPos = p201
            End If
        Else
            .BufPos = p201
        End If
        If ParseIfPrefix() Then
            If ParseAssignExpr() Then
                p200 = .BufPos
                If ParseSEMICOLON() Then
                    ParseIfStatement = True
                    Exit Function
                Else
                    .BufPos = p200
                End If
                If ParseKEYWORD_else() Then
                    p196 = .BufPos
                    If Not (ParsePayload()) Then
                        .BufPos = p196
                    End If
                    If ParseStatement() Then
                        ParseIfStatement = True
                        Exit Function
                    Else
                        .BufPos = p200
                    End If
                Else
                    .BufPos = p200
                End If
                .BufPos = p201
            Else
                .BufPos = p201
            End If
        Else
            .BufPos = p201
        End If
    End With
End Function

Private Function ParseLabeledStatement() As Boolean
    Dim p203 As Long
    Dim p208 As Long

    With ctx
        p203 = .BufPos
        If Not (ParseBlockLabel()) Then
            .BufPos = p203
        End If
        p208 = .BufPos
        If ParseBlock() Then
            ParseLabeledStatement = True
            Exit Function
        Else
            .BufPos = p208
        End If
        If ParseLoopStatement() Then
            ParseLabeledStatement = True
            Exit Function
        Else
            .BufPos = p208
        End If
    End With
End Function

Private Function ParseSwitchExpr() As Boolean
    If ParseKEYWORD_switch() Then
        If ParseLPAREN() Then
            If ParseExpr() Then
                If ParseRPAREN() Then
                    If ParseLBRACE() Then
                        Call ParseSwitchProngList
                        If ParseRBRACE() Then
                            ParseSwitchExpr = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Private Function ParseAssignExpr() As Boolean
    Dim p269 As Long

    With ctx
        If ParseExpr() Then
            p269 = .BufPos
            If ParseAssignOp() Then
                If Not (ParseExpr()) Then
                    .BufPos = p269
                End If
            Else
                .BufPos = p269
            End If
            ParseAssignExpr = True
        End If
    End With
End Function

Private Function ParseIfPrefix() As Boolean
    Dim p722 As Long

    With ctx
        If ParseKEYWORD_if() Then
            If ParseLPAREN() Then
                If ParseExpr() Then
                    If ParseRPAREN() Then
                        p722 = .BufPos
                        If Not (ParsePtrPayload()) Then
                            .BufPos = p722
                        End If
                        ParseIfPrefix = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_else() As Boolean
    With ctx
        If pvMatchString("else") Then               ' "else"
            .BufPos = .BufPos + 4
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_else = True
            End If
        End If
    End With
End Function

Private Function ParsePayload() As Boolean
    If ParsePIPE() Then
        If ParseIDENTIFIER() Then
            If ParsePIPE() Then
                ParsePayload = True
            End If
        End If
    End If
End Function

Private Function ParseBlockLabel() As Boolean
    If ParseIDENTIFIER() Then
        If ParseCOLON() Then
            ParseBlockLabel = True
        End If
    End If
End Function

Private Function ParseLoopStatement() As Boolean
    Dim p210 As Long
    Dim p216 As Long

    With ctx
        p210 = .BufPos
        If Not (ParseKEYWORD_inline()) Then
            .BufPos = p210
        End If
        p216 = .BufPos
        If ParseForStatement() Then
            ParseLoopStatement = True
            Exit Function
        Else
            .BufPos = p216
        End If
        If ParseWhileStatement() Then
            ParseLoopStatement = True
            Exit Function
        Else
            .BufPos = p216
        End If
    End With
End Function

Private Function ParseForStatement() As Boolean
    Dim p234 As Long
    Dim p224 As Long
    Dim p233 As Long

    With ctx
        p234 = .BufPos
        If ParseForPrefix() Then
            If ParseBlockExpr() Then
                p224 = .BufPos
                If ParseKEYWORD_else() Then
                    If Not (ParseStatement()) Then
                        .BufPos = p224
                    End If
                Else
                    .BufPos = p224
                End If
                ParseForStatement = True
                Exit Function
            Else
                .BufPos = p234
            End If
        Else
            .BufPos = p234
        End If
        If ParseForPrefix() Then
            If ParseAssignExpr() Then
                p233 = .BufPos
                If ParseSEMICOLON() Then
                    ParseForStatement = True
                    Exit Function
                Else
                    .BufPos = p233
                End If
                If ParseKEYWORD_else() Then
                    If ParseStatement() Then
                        ParseForStatement = True
                        Exit Function
                    Else
                        .BufPos = p233
                    End If
                Else
                    .BufPos = p233
                End If
                .BufPos = p234
            Else
                .BufPos = p234
            End If
        Else
            .BufPos = p234
        End If
    End With
End Function

Private Function ParseWhileStatement() As Boolean
    Dim p255 As Long
    Dim p242 As Long
    Dim p240 As Long
    Dim p254 As Long
    Dim p250 As Long

    With ctx
        p255 = .BufPos
        If ParseWhilePrefix() Then
            If ParseBlockExpr() Then
                p242 = .BufPos
                If ParseKEYWORD_else() Then
                    p240 = .BufPos
                    If Not (ParsePayload()) Then
                        .BufPos = p240
                    End If
                    If Not (ParseStatement()) Then
                        .BufPos = p242
                    End If
                Else
                    .BufPos = p242
                End If
                ParseWhileStatement = True
                Exit Function
            Else
                .BufPos = p255
            End If
        Else
            .BufPos = p255
        End If
        If ParseWhilePrefix() Then
            If ParseAssignExpr() Then
                p254 = .BufPos
                If ParseSEMICOLON() Then
                    ParseWhileStatement = True
                    Exit Function
                Else
                    .BufPos = p254
                End If
                If ParseKEYWORD_else() Then
                    p250 = .BufPos
                    If Not (ParsePayload()) Then
                        .BufPos = p250
                    End If
                    If ParseStatement() Then
                        ParseWhileStatement = True
                        Exit Function
                    Else
                        .BufPos = p254
                    End If
                Else
                    .BufPos = p254
                End If
                .BufPos = p255
            Else
                .BufPos = p255
            End If
        Else
            .BufPos = p255
        End If
    End With
End Function

Private Function ParseForPrefix() As Boolean
    If ParseKEYWORD_for() Then
        If ParseLPAREN() Then
            If ParseExpr() Then
                If ParseRPAREN() Then
                    If ParsePtrIndexPayload() Then
                        ParseForPrefix = True
                    End If
                End If
            End If
        End If
    End If
End Function

Private Function ParseWhilePrefix() As Boolean
    Dim p730 As Long
    Dim p732 As Long

    With ctx
        If ParseKEYWORD_while() Then
            If ParseLPAREN() Then
                If ParseExpr() Then
                    If ParseRPAREN() Then
                        p730 = .BufPos
                        If Not (ParsePtrPayload()) Then
                            .BufPos = p730
                        End If
                        p732 = .BufPos
                        If Not (ParseWhileContinueExpr()) Then
                            .BufPos = p732
                        End If
                        ParseWhilePrefix = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseAssignOp() As Boolean
    Dim p794 As Long

    With ctx
        p794 = .BufPos
        If ParseASTERISKEQUAL() Then
            ParseAssignOp = True
            Exit Function
        Else
            .BufPos = p794
        End If
        If ParseSLASHEQUAL() Then
            ParseAssignOp = True
            Exit Function
        Else
            .BufPos = p794
        End If
        If ParsePERCENTEQUAL() Then
            ParseAssignOp = True
            Exit Function
        Else
            .BufPos = p794
        End If
        If ParsePLUSEQUAL() Then
            ParseAssignOp = True
            Exit Function
        Else
            .BufPos = p794
        End If
        If ParseMINUSEQUAL() Then
            ParseAssignOp = True
            Exit Function
        Else
            .BufPos = p794
        End If
        If ParseLARROW2EQUAL() Then
            ParseAssignOp = True
            Exit Function
        Else
            .BufPos = p794
        End If
        If ParseRARROW2EQUAL() Then
            ParseAssignOp = True
            Exit Function
        Else
            .BufPos = p794
        End If
        If ParseAMPERSANDEQUAL() Then
            ParseAssignOp = True
            Exit Function
        Else
            .BufPos = p794
        End If
        If ParseCARETEQUAL() Then
            ParseAssignOp = True
            Exit Function
        Else
            .BufPos = p794
        End If
        If ParsePIPEEQUAL() Then
            ParseAssignOp = True
            Exit Function
        Else
            .BufPos = p794
        End If
        If ParseASTERISKPERCENTEQUAL() Then
            ParseAssignOp = True
            Exit Function
        Else
            .BufPos = p794
        End If
        If ParsePLUSPERCENTEQUAL() Then
            ParseAssignOp = True
            Exit Function
        Else
            .BufPos = p794
        End If
        If ParseMINUSPERCENTEQUAL() Then
            ParseAssignOp = True
            Exit Function
        Else
            .BufPos = p794
        End If
        If ParseEQUAL() Then
            ParseAssignOp = True
            Exit Function
        Else
            .BufPos = p794
        End If
    End With
End Function

Private Function ParseKEYWORD_try() As Boolean
    With ctx
        If pvMatchString("try") Then                ' "try"
            .BufPos = .BufPos + 3
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_try = True
            End If
        End If
    End With
End Function

Private Function ParseBoolOrExpr() As Boolean
    Dim p283 As Long

    With ctx
        If ParseBoolAndExpr() Then
            Do
                p283 = .BufPos
                If Not (ParseKEYWORD_or()) Then
                    .BufPos = p283
                    Exit Do
                End If
                If Not (ParseBoolAndExpr()) Then
                    .BufPos = p283
                    Exit Do
                End If
            Loop
            ParseBoolOrExpr = True
        End If
    End With
End Function

Private Function ParseBoolAndExpr() As Boolean
    Dim p291 As Long

    With ctx
        If ParseCompareExpr() Then
            Do
                p291 = .BufPos
                If Not (ParseKEYWORD_and()) Then
                    .BufPos = p291
                    Exit Do
                End If
                If Not (ParseCompareExpr()) Then
                    .BufPos = p291
                    Exit Do
                End If
            Loop
            ParseBoolAndExpr = True
        End If
    End With
End Function

Private Function ParseKEYWORD_or() As Boolean
    With ctx
        If .BufData(.BufPos) = 111 And .BufData(.BufPos + 1) = 114 Then ' "or"
            .BufPos = .BufPos + 2
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_or = True
            End If
        End If
    End With
End Function

Private Function ParseCompareExpr() As Boolean
    Dim p299 As Long

    With ctx
        If ParseBitwiseExpr() Then
            p299 = .BufPos
            If ParseCompareOp() Then
                If Not (ParseBitwiseExpr()) Then
                    .BufPos = p299
                End If
            Else
                .BufPos = p299
            End If
            ParseCompareExpr = True
        End If
    End With
End Function

Private Function ParseKEYWORD_and() As Boolean
    With ctx
        If pvMatchString("and") Then                ' "and"
            .BufPos = .BufPos + 3
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_and = True
            End If
        End If
    End With
End Function

Private Function ParseBitwiseExpr() As Boolean
    Dim p307 As Long

    With ctx
        If ParseBitShiftExpr() Then
            Do
                p307 = .BufPos
                If Not (ParseBitwiseOp()) Then
                    .BufPos = p307
                    Exit Do
                End If
                If Not (ParseBitShiftExpr()) Then
                    .BufPos = p307
                    Exit Do
                End If
            Loop
            ParseBitwiseExpr = True
        End If
    End With
End Function

Private Function ParseCompareOp() As Boolean
    Dim p822 As Long

    With ctx
        p822 = .BufPos
        If ParseEQUALEQUAL() Then
            ParseCompareOp = True
            Exit Function
        Else
            .BufPos = p822
        End If
        If ParseEXCLAMATIONMARKEQUAL() Then
            ParseCompareOp = True
            Exit Function
        Else
            .BufPos = p822
        End If
        If ParseLARROW() Then
            ParseCompareOp = True
            Exit Function
        Else
            .BufPos = p822
        End If
        If ParseRARROW() Then
            ParseCompareOp = True
            Exit Function
        Else
            .BufPos = p822
        End If
        If ParseLARROWEQUAL() Then
            ParseCompareOp = True
            Exit Function
        Else
            .BufPos = p822
        End If
        If ParseRARROWEQUAL() Then
            ParseCompareOp = True
            Exit Function
        Else
            .BufPos = p822
        End If
    End With
End Function

Private Function ParseBitShiftExpr() As Boolean
    Dim p315 As Long

    With ctx
        If ParseAdditionExpr() Then
            Do
                p315 = .BufPos
                If Not (ParseBitShiftOp()) Then
                    .BufPos = p315
                    Exit Do
                End If
                If Not (ParseAdditionExpr()) Then
                    .BufPos = p315
                    Exit Do
                End If
            Loop
            ParseBitShiftExpr = True
        End If
    End With
End Function

Private Function ParseBitwiseOp() As Boolean
    Dim p833 As Long
    Dim p839 As Long

    With ctx
        p833 = .BufPos
        If ParseAMPERSAND() Then
            ParseBitwiseOp = True
            Exit Function
        Else
            .BufPos = p833
        End If
        If ParseCARET() Then
            ParseBitwiseOp = True
            Exit Function
        Else
            .BufPos = p833
        End If
        If ParsePIPE() Then
            ParseBitwiseOp = True
            Exit Function
        Else
            .BufPos = p833
        End If
        If ParseKEYWORD_orelse() Then
            ParseBitwiseOp = True
            Exit Function
        Else
            .BufPos = p833
        End If
        If ParseKEYWORD_catch() Then
            p839 = .BufPos
            If Not (ParsePayload()) Then
                .BufPos = p839
            End If
            ParseBitwiseOp = True
            Exit Function
        Else
            .BufPos = p833
        End If
    End With
End Function

Private Function ParseAdditionExpr() As Boolean
    Dim p323 As Long

    With ctx
        If ParseMultiplyExpr() Then
            Do
                p323 = .BufPos
                If Not (ParseAdditionOp()) Then
                    .BufPos = p323
                    Exit Do
                End If
                If Not (ParseMultiplyExpr()) Then
                    .BufPos = p323
                    Exit Do
                End If
            Loop
            ParseAdditionExpr = True
        End If
    End With
End Function

Private Function ParseBitShiftOp() As Boolean
    Dim p846 As Long

    With ctx
        p846 = .BufPos
        If ParseLARROW2() Then
            ParseBitShiftOp = True
            Exit Function
        Else
            .BufPos = p846
        End If
        If ParseRARROW2() Then
            ParseBitShiftOp = True
            Exit Function
        Else
            .BufPos = p846
        End If
    End With
End Function

Private Function ParseMultiplyExpr() As Boolean
    Dim p331 As Long

    With ctx
        If ParsePrefixExpr() Then
            Do
                p331 = .BufPos
                If Not (ParseMultiplyOp()) Then
                    .BufPos = p331
                    Exit Do
                End If
                If Not (ParsePrefixExpr()) Then
                    .BufPos = p331
                    Exit Do
                End If
            Loop
            ParseMultiplyExpr = True
        End If
    End With
End Function

Private Function ParseAdditionOp() As Boolean
    Dim p851 As Long

    With ctx
        p851 = .BufPos
        If ParsePLUS() Then
            ParseAdditionOp = True
            Exit Function
        Else
            .BufPos = p851
        End If
        If ParseMINUS() Then
            ParseAdditionOp = True
            Exit Function
        Else
            .BufPos = p851
        End If
        If ParsePLUS2() Then
            ParseAdditionOp = True
            Exit Function
        Else
            .BufPos = p851
        End If
        If ParsePLUSPERCENT() Then
            ParseAdditionOp = True
            Exit Function
        Else
            .BufPos = p851
        End If
        If ParseMINUSPERCENT() Then
            ParseAdditionOp = True
            Exit Function
        Else
            .BufPos = p851
        End If
    End With
End Function

Private Function ParsePrefixExpr() As Boolean
    Dim p335 As Long

    With ctx
        Do
            p335 = .BufPos
            If Not (ParsePrefixOp()) Then
                .BufPos = p335
                Exit Do
            End If
        Loop
        If ParsePrimaryExpr() Then
            ParsePrefixExpr = True
        End If
    End With
End Function

Private Function ParseMultiplyOp() As Boolean
    Dim p861 As Long

    With ctx
        p861 = .BufPos
        If ParsePIPE2() Then
            ParseMultiplyOp = True
            Exit Function
        Else
            .BufPos = p861
        End If
        If ParseASTERISK() Then
            ParseMultiplyOp = True
            Exit Function
        Else
            .BufPos = p861
        End If
        If ParseSLASH() Then
            ParseMultiplyOp = True
            Exit Function
        Else
            .BufPos = p861
        End If
        If ParsePERCENT() Then
            ParseMultiplyOp = True
            Exit Function
        Else
            .BufPos = p861
        End If
        If ParseASTERISK2() Then
            ParseMultiplyOp = True
            Exit Function
        Else
            .BufPos = p861
        End If
        If ParseASTERISKPERCENT() Then
            ParseMultiplyOp = True
            Exit Function
        Else
            .BufPos = p861
        End If
    End With
End Function

Private Function ParsePrefixOp() As Boolean
    Dim p872 As Long

    With ctx
        p872 = .BufPos
        If ParseEXCLAMATIONMARK() Then
            ParsePrefixOp = True
            Exit Function
        Else
            .BufPos = p872
        End If
        If ParseMINUS() Then
            ParsePrefixOp = True
            Exit Function
        Else
            .BufPos = p872
        End If
        If ParseTILDE() Then
            ParsePrefixOp = True
            Exit Function
        Else
            .BufPos = p872
        End If
        If ParseMINUSPERCENT() Then
            ParsePrefixOp = True
            Exit Function
        Else
            .BufPos = p872
        End If
        If ParseAMPERSAND() Then
            ParsePrefixOp = True
            Exit Function
        Else
            .BufPos = p872
        End If
        If ParseKEYWORD_try() Then
            ParsePrefixOp = True
            Exit Function
        Else
            .BufPos = p872
        End If
        If ParseKEYWORD_await() Then
            ParsePrefixOp = True
            Exit Function
        Else
            .BufPos = p872
        End If
    End With
End Function

Private Function ParsePrimaryExpr() As Boolean
    Dim p344 As Long
    Dim p348 As Long
    Dim p351 As Long
    Dim p362 As Long
    Dim p371 As Long

    With ctx
        p344 = .BufPos
        If ParseAsmExpr() Then
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p344
        End If
        If ParseIfExpr() Then
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p344
        End If
        If ParseKEYWORD_break() Then
            p348 = .BufPos
            If Not (ParseBreakLabel()) Then
                .BufPos = p348
            End If
            p351 = .BufPos
            If Not (ParseExpr()) Then
                .BufPos = p351
            End If
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p344
        End If
        If ParseKEYWORD_cancel() Then
            If ParseExpr() Then
                ParsePrimaryExpr = True
                Exit Function
            Else
                .BufPos = p344
            End If
        Else
            .BufPos = p344
        End If
        If ParseKEYWORD_comptime() Then
            If ParseExpr() Then
                ParsePrimaryExpr = True
                Exit Function
            Else
                .BufPos = p344
            End If
        Else
            .BufPos = p344
        End If
        If ParseKEYWORD_continue() Then
            p362 = .BufPos
            If Not (ParseBreakLabel()) Then
                .BufPos = p362
            End If
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p344
        End If
        If ParseKEYWORD_resume() Then
            If ParseExpr() Then
                ParsePrimaryExpr = True
                Exit Function
            Else
                .BufPos = p344
            End If
        Else
            .BufPos = p344
        End If
        If ParseKEYWORD_return() Then
            p371 = .BufPos
            If Not (ParseExpr()) Then
                .BufPos = p371
            End If
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p344
        End If
        If ParseLabeledExpr() Then
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p344
        End If
        If ParseCurlySuffixExpr() Then
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p344
        End If
        Call pvSetExpected("PrimaryExpr")
    End With
End Function

Private Function ParseAsmExpr() As Boolean
    Dim p606 As Long
    Dim p612 As Long

    With ctx
        If ParseKEYWORD_asm() Then
            p606 = .BufPos
            If Not (ParseKEYWORD_volatile()) Then
                .BufPos = p606
            End If
            If ParseLPAREN() Then
                If ParseSTRINGLITERAL() Then
                    p612 = .BufPos
                    If Not (ParseAsmOutput()) Then
                        .BufPos = p612
                    End If
                    If ParseRPAREN() Then
                        ParseAsmExpr = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseIfExpr() As Boolean
    Dim p384 As Long
    Dim p382 As Long

    With ctx
        If ParseIfPrefix() Then
            If ParseExpr() Then
                p384 = .BufPos
                If ParseKEYWORD_else() Then
                    p382 = .BufPos
                    If Not (ParsePayload()) Then
                        .BufPos = p382
                    End If
                    If Not (ParseExpr()) Then
                        .BufPos = p384
                    End If
                Else
                    .BufPos = p384
                End If
                ParseIfExpr = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_break() As Boolean
    With ctx
        If pvMatchString("break") Then              ' "break"
            .BufPos = .BufPos + 5
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_break = True
            End If
        End If
    End With
End Function

Private Function ParseBreakLabel() As Boolean
    If ParseCOLON() Then
        If ParseIDENTIFIER() Then
            ParseBreakLabel = True
        End If
    End If
End Function

Private Function ParseKEYWORD_cancel() As Boolean
    With ctx
        If pvMatchString("cancel") Then             ' "cancel"
            .BufPos = .BufPos + 6
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_cancel = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_continue() As Boolean
    With ctx
        If pvMatchString("continue") Then           ' "continue"
            .BufPos = .BufPos + 8
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_continue = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_resume() As Boolean
    With ctx
        If pvMatchString("resume") Then             ' "resume"
            .BufPos = .BufPos + 6
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_resume = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_return() As Boolean
    With ctx
        If pvMatchString("return") Then             ' "return"
            .BufPos = .BufPos + 6
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_return = True
            End If
        End If
    End With
End Function

Private Function ParseLabeledExpr() As Boolean
    Dim p387 As Long
    Dim p392 As Long

    With ctx
        p387 = .BufPos
        If Not (ParseBlockLabel()) Then
            .BufPos = p387
        End If
        p392 = .BufPos
        If ParseBlock() Then
            ParseLabeledExpr = True
            Exit Function
        Else
            .BufPos = p392
        End If
        If ParseLoopExpr() Then
            ParseLabeledExpr = True
            Exit Function
        Else
            .BufPos = p392
        End If
    End With
End Function

Private Function ParseCurlySuffixExpr() As Boolean
    Dim p427 As Long

    With ctx
        If ParseTypeExpr() Then
            p427 = .BufPos
            If Not (ParseInitList()) Then
                .BufPos = p427
            End If
            ParseCurlySuffixExpr = True
        End If
    End With
End Function

Private Function ParseLoopExpr() As Boolean
    Dim p401 As Long
    Dim p407 As Long

    With ctx
        p401 = .BufPos
        If Not (ParseKEYWORD_inline()) Then
            .BufPos = p401
        End If
        p407 = .BufPos
        If ParseForExpr() Then
            ParseLoopExpr = True
            Exit Function
        Else
            .BufPos = p407
        End If
        If ParseWhileExpr() Then
            ParseLoopExpr = True
            Exit Function
        Else
            .BufPos = p407
        End If
    End With
End Function

Private Function ParseLBRACE() As Boolean
    With ctx
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Call Parseskip
            Call pvSetAdvance
            ParseLBRACE = True
        End If
    End With
End Function

Private Function ParseRBRACE() As Boolean
    With ctx
        If .BufData(.BufPos) = 125 Then             ' "}"
            .BufPos = .BufPos + 1
            Call Parseskip
            Call pvSetAdvance
            ParseRBRACE = True
        End If
    End With
End Function

Private Function ParseForExpr() As Boolean
    Dim p414 As Long

    With ctx
        If ParseForPrefix() Then
            If ParseExpr() Then
                p414 = .BufPos
                If ParseKEYWORD_else() Then
                    If Not (ParseExpr()) Then
                        .BufPos = p414
                    End If
                Else
                    .BufPos = p414
                End If
                ParseForExpr = True
            End If
        End If
    End With
End Function

Private Function ParseWhileExpr() As Boolean
    Dim p422 As Long
    Dim p420 As Long

    With ctx
        If ParseWhilePrefix() Then
            If ParseExpr() Then
                p422 = .BufPos
                If ParseKEYWORD_else() Then
                    p420 = .BufPos
                    If Not (ParsePayload()) Then
                        .BufPos = p420
                    End If
                    If Not (ParseExpr()) Then
                        .BufPos = p422
                    End If
                Else
                    .BufPos = p422
                End If
                ParseWhileExpr = True
            End If
        End If
    End With
End Function

Private Function ParseInitList() As Boolean
    Dim p451 As Long
    Dim p436 As Long
    Dim p438 As Long
    Dim p446 As Long
    Dim p448 As Long

    With ctx
        p451 = .BufPos
        If ParseLBRACE() Then
            If ParseFieldInit() Then
                Do
                    p436 = .BufPos
                    If Not (ParseCOMMA()) Then
                        .BufPos = p436
                        Exit Do
                    End If
                    If Not (ParseFieldInit()) Then
                        .BufPos = p436
                        Exit Do
                    End If
                Loop
                p438 = .BufPos
                If Not (ParseCOMMA()) Then
                    .BufPos = p438
                End If
                If ParseRBRACE() Then
                    ParseInitList = True
                    Exit Function
                Else
                    .BufPos = p451
                End If
            Else
                .BufPos = p451
            End If
        Else
            .BufPos = p451
        End If
        If ParseLBRACE() Then
            If ParseExpr() Then
                Do
                    p446 = .BufPos
                    If Not (ParseCOMMA()) Then
                        .BufPos = p446
                        Exit Do
                    End If
                    If Not (ParseExpr()) Then
                        .BufPos = p446
                        Exit Do
                    End If
                Loop
                p448 = .BufPos
                If Not (ParseCOMMA()) Then
                    .BufPos = p448
                End If
                If ParseRBRACE() Then
                    ParseInitList = True
                    Exit Function
                Else
                    .BufPos = p451
                End If
            Else
                .BufPos = p451
            End If
        Else
            .BufPos = p451
        End If
        If ParseLBRACE() Then
            If ParseRBRACE() Then
                ParseInitList = True
                Exit Function
            Else
                .BufPos = p451
            End If
        Else
            .BufPos = p451
        End If
    End With
End Function

Private Function ParseFieldInit() As Boolean
    If ParseDOT() Then
        If ParseIDENTIFIER() Then
            If ParseEQUAL() Then
                If ParseExpr() Then
                    ParseFieldInit = True
                End If
            End If
        End If
    End If
End Function

Private Function ParsePrefixTypeOp() As Boolean
    Dim p885 As Long
    Dim p890 As Long
    Dim p909 As Long
    Dim p903 As Long

    With ctx
        p885 = .BufPos
        If ParseQUESTIONMARK() Then
            ParsePrefixTypeOp = True
            Exit Function
        Else
            .BufPos = p885
        End If
        If ParseKEYWORD_promise() Then
            If ParseMINUSRARROW() Then
                ParsePrefixTypeOp = True
                Exit Function
            Else
                .BufPos = p885
            End If
        Else
            .BufPos = p885
        End If
        If ParseArrayTypeStart() Then
            Do
                p890 = .BufPos
                If Not (ParseByteAlign()) Then
                    .BufPos = p890
                    If Not (ParseKEYWORD_const()) Then
                        .BufPos = p890
                        If Not (ParseKEYWORD_volatile()) Then
                            .BufPos = p890
                            .BufPos = p890
                            Exit Do
                        End If
                    End If
                End If
            Loop
            ParsePrefixTypeOp = True
            Exit Function
        Else
            .BufPos = p885
        End If
        If ParsePtrTypeStart() Then
            Do
                p909 = .BufPos
                If ParseKEYWORD_align() Then
                    If ParseLPAREN() Then
                        If ParseExpr() Then
                            p903 = .BufPos
                            If ParseCOLON() Then
                                If ParseINTEGER() Then
                                    If ParseCOLON() Then
                                        If Not (ParseINTEGER()) Then
                                            .BufPos = p903
                                        End If
                                    Else
                                        .BufPos = p903
                                    End If
                                Else
                                    .BufPos = p903
                                End If
                            Else
                                .BufPos = p903
                            End If
                            If Not (ParseRPAREN()) Then
                                .BufPos = p909
                                If Not (ParseKEYWORD_const()) Then
                                    .BufPos = p909
                                    If Not (ParseKEYWORD_volatile()) Then
                                        .BufPos = p909
                                        .BufPos = p909
                                        Exit Do
                                    End If
                                End If
                            End If
                        Else
                            .BufPos = p909
                            If Not (ParseKEYWORD_const()) Then
                                .BufPos = p909
                                If Not (ParseKEYWORD_volatile()) Then
                                    .BufPos = p909
                                    .BufPos = p909
                                    Exit Do
                                End If
                            End If
                        End If
                    Else
                        .BufPos = p909
                        If Not (ParseKEYWORD_const()) Then
                            .BufPos = p909
                            If Not (ParseKEYWORD_volatile()) Then
                                .BufPos = p909
                                .BufPos = p909
                                Exit Do
                            End If
                        End If
                    End If
                Else
                    .BufPos = p909
                    If Not (ParseKEYWORD_const()) Then
                        .BufPos = p909
                        If Not (ParseKEYWORD_volatile()) Then
                            .BufPos = p909
                            .BufPos = p909
                            Exit Do
                        End If
                    End If
                End If
            Loop
            ParsePrefixTypeOp = True
            Exit Function
        Else
            .BufPos = p885
        End If
    End With
End Function

Private Function ParseErrorUnionExpr() As Boolean
    Dim p465 As Long

    With ctx
        If ParseSuffixExpr() Then
            p465 = .BufPos
            If ParseEXCLAMATIONMARK() Then
                If Not (ParseTypeExpr()) Then
                    .BufPos = p465
                End If
            Else
                .BufPos = p465
            End If
            ParseErrorUnionExpr = True
        End If
    End With
End Function

Private Function ParseSuffixExpr() As Boolean
    Dim p484 As Long
    Dim p474 As Long
    Dim p481 As Long

    With ctx
        p484 = .BufPos
        If ParseAsyncPrefix() Then
            If ParsePrimaryTypeExpr() Then
                Do
                    p474 = .BufPos
                    If Not (ParseSuffixOp()) Then
                        .BufPos = p474
                        Exit Do
                    End If
                Loop
                If ParseFnCallArgumnets() Then
                    ParseSuffixExpr = True
                    Exit Function
                Else
                    .BufPos = p484
                End If
            Else
                .BufPos = p484
            End If
        Else
            .BufPos = p484
        End If
        If ParsePrimaryTypeExpr() Then
            Do
                p481 = .BufPos
                If Not (ParseSuffixOp()) Then
                    .BufPos = p481
                    If Not (ParseFnCallArgumnets()) Then
                        .BufPos = p481
                        .BufPos = p481
                        Exit Do
                    End If
                End If
            Loop
            ParseSuffixExpr = True
            Exit Function
        Else
            .BufPos = p484
        End If
    End With
End Function

Private Function ParseAsyncPrefix() As Boolean
    Dim p934 As Long

    With ctx
        If ParseKEYWORD_async() Then
            p934 = .BufPos
            If ParseLARROW() Then
                If ParsePrefixExpr() Then
                    If Not (ParseRARROW()) Then
                        .BufPos = p934
                    End If
                Else
                    .BufPos = p934
                End If
            Else
                .BufPos = p934
            End If
            ParseAsyncPrefix = True
        End If
    End With
End Function

Private Function ParsePrimaryTypeExpr() As Boolean
    Dim p491 As Long

    With ctx
        p491 = .BufPos
        If ParseBUILTININDENTIFIER() Then
            If ParseFnCallArgumnets() Then
                ParsePrimaryTypeExpr = True
                Exit Function
            Else
                .BufPos = p491
            End If
        Else
            .BufPos = p491
        End If
        If ParseCHAR_LITERAL() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseContainerDecl() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseErrorSetDecl() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseFLOAT() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseFnProto() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseGroupedExpr() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseIDENTIFIER() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseIfTypeExpr() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseINTEGER() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseKEYWORD_anyerror() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseKEYWORD_comptime() Then
            If ParseTypeExpr() Then
                ParsePrimaryTypeExpr = True
                Exit Function
            Else
                .BufPos = p491
            End If
        Else
            .BufPos = p491
        End If
        If ParseKEYWORD_error() Then
            If ParseDOT() Then
                If ParseIDENTIFIER() Then
                    ParsePrimaryTypeExpr = True
                    Exit Function
                Else
                    .BufPos = p491
                End If
            Else
                .BufPos = p491
            End If
        Else
            .BufPos = p491
        End If
        If ParseKEYWORD_false() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseKEYWORD_null() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseKEYWORD_promise() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseKEYWORD_true() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseKEYWORD_undefined() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseKEYWORD_unreachable() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseLabeledTypeExpr() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseSTRINGLITERAL() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
        If ParseSwitchExpr() Then
            ParsePrimaryTypeExpr = True
            Exit Function
        Else
            .BufPos = p491
        End If
    End With
End Function

Private Function ParseSuffixOp() As Boolean
    Dim p926 As Long
    Dim p920 As Long
    Dim p918 As Long

    With ctx
        p926 = .BufPos
        If ParseLBRACKET() Then
            If ParseExpr() Then
                p920 = .BufPos
                If ParseDOT2() Then
                    p918 = .BufPos
                    If Not (ParseExpr()) Then
                        .BufPos = p918
                    End If
                Else
                    .BufPos = p920
                End If
                If ParseRBRACKET() Then
                    ParseSuffixOp = True
                    Exit Function
                Else
                    .BufPos = p926
                End If
            Else
                .BufPos = p926
            End If
        Else
            .BufPos = p926
        End If
        If ParseDOT() Then
            If ParseIDENTIFIER() Then
                ParseSuffixOp = True
                Exit Function
            Else
                .BufPos = p926
            End If
        Else
            .BufPos = p926
        End If
        If ParseDOTASTERISK() Then
            ParseSuffixOp = True
            Exit Function
        Else
            .BufPos = p926
        End If
        If ParseDOTQUESTIONMARK() Then
            ParseSuffixOp = True
            Exit Function
        Else
            .BufPos = p926
        End If
    End With
End Function

Private Function ParseFnCallArgumnets() As Boolean
    If ParseLPAREN() Then
        Call ParseExprList
        If ParseRPAREN() Then
            ParseFnCallArgumnets = True
        End If
    End If
End Function

Private Function ParseBUILTININDENTIFIER() As Boolean
    With ctx
        If .BufData(.BufPos) = 64 Then              ' "@"
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 65 To 90, 97 To 122, 95            ' [A-Za-z_]
                .BufPos = .BufPos + 1
                Do
                    Select Case .BufData(.BufPos)
                    Case 65 To 90, 97 To 122, 48 To 57, 95 ' [A-Za-z0-9_]
                        .BufPos = .BufPos + 1
                    Case Else
                        Exit Do
                    End Select
                Loop
                Call Parseskip
                Call pvSetAdvance
                ParseBUILTININDENTIFIER = True
            End Select
        End If
    End With
End Function

Private Function ParseCHAR_LITERAL() As Boolean
    With ctx
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            If Parsechar_char() Then
                If .BufData(.BufPos) = 39 Then      ' "'"
                    .BufPos = .BufPos + 1
                    Call Parseskip
                    Call pvSetAdvance
                    ParseCHAR_LITERAL = True
                    Exit Function
                Else
                    Call pvSetExpected("CHAR_LITERAL")
                End If
            Else
                Call pvSetExpected("CHAR_LITERAL")
            End If
        Else
            Call pvSetExpected("CHAR_LITERAL")
        End If
    End With
End Function

Private Function ParseContainerDecl() As Boolean
    Dim p536 As Long

    With ctx
        p536 = .BufPos
        If ParseKEYWORD_extern() Then
            GoTo L103
        Else
            .BufPos = p536
        End If
        If ParseKEYWORD_packed() Then
            GoTo L103
        Else
            .BufPos = p536
        End If
        .BufPos = p536
L103:
        If ParseContainerDeclAuto() Then
            ParseContainerDecl = True
        End If
    End With
End Function

Private Function ParseErrorSetDecl() As Boolean
    If ParseKEYWORD_error() Then
        If ParseLBRACE() Then
            Call ParseIdentifierList
            If ParseRBRACE() Then
                ParseErrorSetDecl = True
            End If
        End If
    End If
End Function

Private Function ParseFLOAT() As Boolean
    Dim p1138 As Long
    Dim i1110 As Long
    Dim i1114 As Long
    Dim p1118 As Long
    Dim i1120 As Long
    Dim i1125 As Long
    Dim i1129 As Long
    Dim p1133 As Long
    Dim i1135 As Long
    Dim i1141 As Long
    Dim i1145 As Long
    Dim p1149 As Long
    Dim i1151 As Long
    Dim i1155 As Long
    Dim i1159 As Long
    Dim p1163 As Long
    Dim i1165 As Long
    Dim i1170 As Long
    Dim i1178 As Long
    Dim i1182 As Long
    Dim i1190 As Long
    Dim i1194 As Long
    Dim i1202 As Long
    Dim i1205 As Long
    Dim i1213 As Long

    With ctx
        p1138 = .BufPos
        If .BufData(.BufPos) = 48 And .BufData(.BufPos + 1) = 98 Then ' "0b"
            .BufPos = .BufPos + 2
            For i1110 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 48, 49                         ' [01]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If i1110 <> 0 Then
                If .BufData(.BufPos) = 46 Then      ' "."
                    .BufPos = .BufPos + 1
                    For i1114 = 0 To LNG_MAXINT
                        Select Case .BufData(.BufPos)
                        Case 48, 49                 ' [01]
                            .BufPos = .BufPos + 1
                        Case Else
                            Exit For
                        End Select
                    Next
                    If i1114 <> 0 Then
                        p1118 = .BufPos
                        Select Case .BufData(.BufPos)
                        Case 101, 69                ' [eE]
                            .BufPos = .BufPos + 1
                            Select Case .BufData(.BufPos)
                            Case 45, 43             ' [-+]
                                .BufPos = .BufPos + 1
                            End Select
                            For i1120 = 0 To LNG_MAXINT
                                Select Case .BufData(.BufPos)
                                Case 48, 49         ' [01]
                                    .BufPos = .BufPos + 1
                                Case Else
                                    Exit For
                                End Select
                            Next
                            If i1120 <> 0 Then
                            Else
                                .BufPos = p1118
                            End If
                        End Select
                        Call Parseskip
                        Call pvSetAdvance
                        ParseFLOAT = True
                        Exit Function
                    Else
                        .BufPos = p1138
                    End If
                Else
                    .BufPos = p1138
                End If
            Else
                .BufPos = p1138
            End If
        End If
        If .BufData(.BufPos) = 48 And .BufData(.BufPos + 1) = 111 Then ' "0o"
            .BufPos = .BufPos + 2
            For i1125 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If i1125 <> 0 Then
                If .BufData(.BufPos) = 46 Then      ' "."
                    .BufPos = .BufPos + 1
                    For i1129 = 0 To LNG_MAXINT
                        Select Case .BufData(.BufPos)
                        Case 48 To 55               ' [0-7]
                            .BufPos = .BufPos + 1
                        Case Else
                            Exit For
                        End Select
                    Next
                    If i1129 <> 0 Then
                        p1133 = .BufPos
                        Select Case .BufData(.BufPos)
                        Case 101, 69                ' [eE]
                            .BufPos = .BufPos + 1
                            Select Case .BufData(.BufPos)
                            Case 45, 43             ' [-+]
                                .BufPos = .BufPos + 1
                            End Select
                            For i1135 = 0 To LNG_MAXINT
                                Select Case .BufData(.BufPos)
                                Case 48 To 55       ' [0-7]
                                    .BufPos = .BufPos + 1
                                Case Else
                                    Exit For
                                End Select
                            Next
                            If i1135 <> 0 Then
                            Else
                                .BufPos = p1133
                            End If
                        End Select
                        Call Parseskip
                        Call pvSetAdvance
                        ParseFLOAT = True
                        Exit Function
                    Else
                        .BufPos = p1138
                    End If
                Else
                    .BufPos = p1138
                End If
            Else
                .BufPos = p1138
            End If
        End If
        If .BufData(.BufPos) = 48 And .BufData(.BufPos + 1) = 120 Then ' "0x"
            .BufPos = .BufPos + 2
            For i1141 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 48 To 57, 97 To 102, 65 To 70  ' [0-9a-fA-F]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If i1141 <> 0 Then
                If .BufData(.BufPos) = 46 Then      ' "."
                    .BufPos = .BufPos + 1
                    For i1145 = 0 To LNG_MAXINT
                        Select Case .BufData(.BufPos)
                        Case 48 To 57, 97 To 102, 65 To 70 ' [0-9a-fA-F]
                            .BufPos = .BufPos + 1
                        Case Else
                            Exit For
                        End Select
                    Next
                    If i1145 <> 0 Then
                        p1149 = .BufPos
                        Select Case .BufData(.BufPos)
                        Case 112, 80                ' [pP]
                            .BufPos = .BufPos + 1
                            Select Case .BufData(.BufPos)
                            Case 45, 43             ' [-+]
                                .BufPos = .BufPos + 1
                            End Select
                            For i1151 = 0 To LNG_MAXINT
                                Select Case .BufData(.BufPos)
                                Case 48 To 57, 97 To 102, 65 To 70 ' [0-9a-fA-F]
                                    .BufPos = .BufPos + 1
                                Case Else
                                    Exit For
                                End Select
                            Next
                            If i1151 <> 0 Then
                            Else
                                .BufPos = p1149
                            End If
                        End Select
                        Call Parseskip
                        Call pvSetAdvance
                        ParseFLOAT = True
                        Exit Function
                    Else
                        .BufPos = p1138
                    End If
                Else
                    .BufPos = p1138
                End If
            Else
                .BufPos = p1138
            End If
        End If
        For i1155 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1155 <> 0 Then
            If .BufData(.BufPos) = 46 Then          ' "."
                .BufPos = .BufPos + 1
                For i1159 = 0 To LNG_MAXINT
                    Select Case .BufData(.BufPos)
                    Case 48 To 57                   ' [0-9]
                        .BufPos = .BufPos + 1
                    Case Else
                        Exit For
                    End Select
                Next
                If i1159 <> 0 Then
                    p1163 = .BufPos
                    Select Case .BufData(.BufPos)
                    Case 101, 69                    ' [eE]
                        .BufPos = .BufPos + 1
                        Select Case .BufData(.BufPos)
                        Case 45, 43                 ' [-+]
                            .BufPos = .BufPos + 1
                        End Select
                        For i1165 = 0 To LNG_MAXINT
                            Select Case .BufData(.BufPos)
                            Case 48 To 57           ' [0-9]
                                .BufPos = .BufPos + 1
                            Case Else
                                Exit For
                            End Select
                        Next
                        If i1165 <> 0 Then
                        Else
                            .BufPos = p1163
                        End If
                    End Select
                    Call Parseskip
                    Call pvSetAdvance
                    ParseFLOAT = True
                    Exit Function
                Else
                    .BufPos = p1138
                End If
            Else
                .BufPos = p1138
            End If
        Else
            .BufPos = p1138
        End If
        If .BufData(.BufPos) = 48 And .BufData(.BufPos + 1) = 98 Then ' "0b"
            .BufPos = .BufPos + 2
            For i1170 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 48, 49                         ' [01]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If i1170 <> 0 Then
                If .BufData(.BufPos) = 46 Then      ' "."
                    .BufPos = .BufPos + 1
                End If
                Select Case .BufData(.BufPos)
                Case 101, 69                        ' [eE]
                    .BufPos = .BufPos + 1
                    Select Case .BufData(.BufPos)
                    Case 45, 43                     ' [-+]
                        .BufPos = .BufPos + 1
                    End Select
                    For i1178 = 0 To LNG_MAXINT
                        Select Case .BufData(.BufPos)
                        Case 48, 49                 ' [01]
                            .BufPos = .BufPos + 1
                        Case Else
                            Exit For
                        End Select
                    Next
                    If i1178 <> 0 Then
                        Call Parseskip
                        Call pvSetAdvance
                        ParseFLOAT = True
                        Exit Function
                    Else
                        .BufPos = p1138
                    End If
                Case Else
                    .BufPos = p1138
                End Select
            Else
                .BufPos = p1138
            End If
        End If
        If .BufData(.BufPos) = 48 And .BufData(.BufPos + 1) = 111 Then ' "0o"
            .BufPos = .BufPos + 2
            For i1182 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If i1182 <> 0 Then
                If .BufData(.BufPos) = 46 Then      ' "."
                    .BufPos = .BufPos + 1
                End If
                Select Case .BufData(.BufPos)
                Case 101, 69                        ' [eE]
                    .BufPos = .BufPos + 1
                    Select Case .BufData(.BufPos)
                    Case 45, 43                     ' [-+]
                        .BufPos = .BufPos + 1
                    End Select
                    For i1190 = 0 To LNG_MAXINT
                        Select Case .BufData(.BufPos)
                        Case 48 To 55               ' [0-7]
                            .BufPos = .BufPos + 1
                        Case Else
                            Exit For
                        End Select
                    Next
                    If i1190 <> 0 Then
                        Call Parseskip
                        Call pvSetAdvance
                        ParseFLOAT = True
                        Exit Function
                    Else
                        .BufPos = p1138
                    End If
                Case Else
                    .BufPos = p1138
                End Select
            Else
                .BufPos = p1138
            End If
        End If
        If .BufData(.BufPos) = 48 And .BufData(.BufPos + 1) = 120 Then ' "0x"
            .BufPos = .BufPos + 2
            For i1194 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 48 To 57, 97 To 102, 65 To 70  ' [0-9a-fA-F]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If i1194 <> 0 Then
                If .BufData(.BufPos) = 46 Then      ' "."
                    .BufPos = .BufPos + 1
                End If
                Select Case .BufData(.BufPos)
                Case 112, 80                        ' [pP]
                    .BufPos = .BufPos + 1
                    Select Case .BufData(.BufPos)
                    Case 45, 43                     ' [-+]
                        .BufPos = .BufPos + 1
                    End Select
                    For i1202 = 0 To LNG_MAXINT
                        Select Case .BufData(.BufPos)
                        Case 48 To 57, 97 To 102, 65 To 70 ' [0-9a-fA-F]
                            .BufPos = .BufPos + 1
                        Case Else
                            Exit For
                        End Select
                    Next
                    If i1202 <> 0 Then
                        Call Parseskip
                        Call pvSetAdvance
                        ParseFLOAT = True
                        Exit Function
                    Else
                        .BufPos = p1138
                    End If
                Case Else
                    .BufPos = p1138
                End Select
            Else
                .BufPos = p1138
            End If
        End If
        For i1205 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1205 <> 0 Then
            If .BufData(.BufPos) = 46 Then          ' "."
                .BufPos = .BufPos + 1
            End If
            Select Case .BufData(.BufPos)
            Case 101, 69                            ' [eE]
                .BufPos = .BufPos + 1
                Select Case .BufData(.BufPos)
                Case 45, 43                         ' [-+]
                    .BufPos = .BufPos + 1
                End Select
                For i1213 = 0 To LNG_MAXINT
                    Select Case .BufData(.BufPos)
                    Case 48 To 57                   ' [0-9]
                        .BufPos = .BufPos + 1
                    Case Else
                        Exit For
                    End Select
                Next
                If i1213 <> 0 Then
                    Call Parseskip
                    Call pvSetAdvance
                    ParseFLOAT = True
                    Exit Function
                Else
                    .BufPos = p1138
                End If
            Case Else
                .BufPos = p1138
            End Select
        Else
            .BufPos = p1138
        End If
        Call pvSetExpected("FLOAT")
    End With
End Function

Private Function ParseGroupedExpr() As Boolean
    If ParseLPAREN() Then
        If ParseExpr() Then
            If ParseRPAREN() Then
                ParseGroupedExpr = True
            End If
        End If
    End If
End Function

Private Function ParseIfTypeExpr() As Boolean
    Dim p557 As Long
    Dim p555 As Long

    With ctx
        If ParseIfPrefix() Then
            If ParseTypeExpr() Then
                p557 = .BufPos
                If ParseKEYWORD_else() Then
                    p555 = .BufPos
                    If Not (ParsePayload()) Then
                        .BufPos = p555
                    End If
                    If Not (ParseTypeExpr()) Then
                        .BufPos = p557
                    End If
                Else
                    .BufPos = p557
                End If
                ParseIfTypeExpr = True
            End If
        End If
    End With
End Function

Private Function ParseINTEGER() As Boolean
    Dim p1225 As Long
    Dim i1217 As Long
    Dim i1222 As Long
    Dim i1228 As Long
    Dim i1232 As Long

    With ctx
        p1225 = .BufPos
        If .BufData(.BufPos) = 48 And .BufData(.BufPos + 1) = 98 Then ' "0b"
            .BufPos = .BufPos + 2
            For i1217 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 48, 49                         ' [01]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If i1217 <> 0 Then
                Call Parseskip
                Call pvSetAdvance
                ParseINTEGER = True
                Exit Function
            Else
                .BufPos = p1225
            End If
        End If
        If .BufData(.BufPos) = 48 And .BufData(.BufPos + 1) = 111 Then ' "0o"
            .BufPos = .BufPos + 2
            For i1222 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If i1222 <> 0 Then
                Call Parseskip
                Call pvSetAdvance
                ParseINTEGER = True
                Exit Function
            Else
                .BufPos = p1225
            End If
        End If
        If .BufData(.BufPos) = 48 And .BufData(.BufPos + 1) = 120 Then ' "0x"
            .BufPos = .BufPos + 2
            For i1228 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 48 To 57, 97 To 102, 65 To 70  ' [0-9a-fA-F]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If i1228 <> 0 Then
                Call Parseskip
                Call pvSetAdvance
                ParseINTEGER = True
                Exit Function
            Else
                .BufPos = p1225
            End If
        End If
        For i1232 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1232 <> 0 Then
            Call Parseskip
            Call pvSetAdvance
            ParseINTEGER = True
            Exit Function
        Else
            .BufPos = p1225
        End If
        Call pvSetExpected("INTEGER")
    End With
End Function

Private Function ParseKEYWORD_anyerror() As Boolean
    With ctx
        If pvMatchString("anyerror") Then           ' "anyerror"
            .BufPos = .BufPos + 8
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_anyerror = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_error() As Boolean
    With ctx
        If pvMatchString("error") Then              ' "error"
            .BufPos = .BufPos + 5
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_error = True
            End If
        End If
    End With
End Function

Private Function ParseDOT() As Boolean
    Dim p1319 As Long

    With ctx
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
            p1319 = .BufPos
            Select Case .BufData(.BufPos)
            Case 42, 46, 63                         ' [*.?]
                '--- do nothing
            Case Else
                .BufPos = p1319
                Call Parseskip
                Call pvSetAdvance
                ParseDOT = True
            End Select
        End If
    End With
End Function

Private Function ParseKEYWORD_false() As Boolean
    With ctx
        If pvMatchString("false") Then              ' "false"
            .BufPos = .BufPos + 5
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_false = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_null() As Boolean
    With ctx
        If pvMatchString("null") Then               ' "null"
            .BufPos = .BufPos + 4
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_null = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_promise() As Boolean
    With ctx
        If pvMatchString("promise") Then            ' "promise"
            .BufPos = .BufPos + 7
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_promise = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_true() As Boolean
    With ctx
        If pvMatchString("true") Then               ' "true"
            .BufPos = .BufPos + 4
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_true = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_undefined() As Boolean
    With ctx
        If pvMatchString("undefined") Then          ' "undefined"
            .BufPos = .BufPos + 9
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_undefined = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_unreachable() As Boolean
    With ctx
        If pvMatchString("unreachable") Then        ' "unreachable"
            .BufPos = .BufPos + 11
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_unreachable = True
            End If
        End If
    End With
End Function

Private Function ParseLabeledTypeExpr() As Boolean
    Dim p568 As Long
    Dim p563 As Long

    With ctx
        p568 = .BufPos
        If ParseBlockLabel() Then
            If ParseBlock() Then
                ParseLabeledTypeExpr = True
                Exit Function
            Else
                .BufPos = p568
            End If
        Else
            .BufPos = p568
        End If
        p563 = .BufPos
        If Not (ParseBlockLabel()) Then
            .BufPos = p563
        End If
        If ParseLoopTypeExpr() Then
            ParseLabeledTypeExpr = True
            Exit Function
        Else
            .BufPos = p568
        End If
    End With
End Function

Private Function ParseKEYWORD_packed() As Boolean
    With ctx
        If pvMatchString("packed") Then             ' "packed"
            .BufPos = .BufPos + 6
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_packed = True
            End If
        End If
    End With
End Function

Private Function ParseContainerDeclAuto() As Boolean
    If ParseContainerDeclType() Then
        If ParseLBRACE() Then
            Call ParseContainerMembers
            If ParseRBRACE() Then
                ParseContainerDeclAuto = True
            End If
        End If
    End If
End Function

Private Sub ParseIdentifierList()
    Dim p996 As Long
    Dim p998 As Long

    With ctx
        Do
            p996 = .BufPos
            If Not (ParseIDENTIFIER()) Then
                .BufPos = p996
                Exit Do
            End If
            If Not (ParseCOMMA()) Then
                .BufPos = p996
                Exit Do
            End If
        Loop
        p998 = .BufPos
        If Not (ParseIDENTIFIER()) Then
            .BufPos = p998
        End If
    End With
End Sub

Private Function ParseLoopTypeExpr() As Boolean
    Dim p569 As Long
    Dim p575 As Long

    With ctx
        p569 = .BufPos
        If Not (ParseKEYWORD_inline()) Then
            .BufPos = p569
        End If
        p575 = .BufPos
        If ParseForTypeExpr() Then
            ParseLoopTypeExpr = True
            Exit Function
        Else
            .BufPos = p575
        End If
        If ParseWhileTypeExpr() Then
            ParseLoopTypeExpr = True
            Exit Function
        Else
            .BufPos = p575
        End If
    End With
End Function

Private Function ParseForTypeExpr() As Boolean
    Dim p582 As Long

    With ctx
        If ParseForPrefix() Then
            If ParseTypeExpr() Then
                p582 = .BufPos
                If ParseKEYWORD_else() Then
                    If Not (ParseTypeExpr()) Then
                        .BufPos = p582
                    End If
                Else
                    .BufPos = p582
                End If
                ParseForTypeExpr = True
            End If
        End If
    End With
End Function

Private Function ParseWhileTypeExpr() As Boolean
    Dim p590 As Long
    Dim p588 As Long

    With ctx
        If ParseWhilePrefix() Then
            If ParseTypeExpr() Then
                p590 = .BufPos
                If ParseKEYWORD_else() Then
                    p588 = .BufPos
                    If Not (ParsePayload()) Then
                        .BufPos = p588
                    End If
                    If Not (ParseTypeExpr()) Then
                        .BufPos = p590
                    End If
                Else
                    .BufPos = p590
                End If
                ParseWhileTypeExpr = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_switch() As Boolean
    With ctx
        If pvMatchString("switch") Then             ' "switch"
            .BufPos = .BufPos + 6
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_switch = True
            End If
        End If
    End With
End Function

Private Sub ParseSwitchProngList()
    Dim p1003 As Long
    Dim p1005 As Long

    With ctx
        Do
            p1003 = .BufPos
            If Not (ParseSwitchProng()) Then
                .BufPos = p1003
                Exit Do
            End If
            If Not (ParseCOMMA()) Then
                .BufPos = p1003
                Exit Do
            End If
        Loop
        p1005 = .BufPos
        If Not (ParseSwitchProng()) Then
            .BufPos = p1005
        End If
    End With
End Sub

Private Function ParseKEYWORD_asm() As Boolean
    With ctx
        If pvMatchString("asm") Then                ' "asm"
            .BufPos = .BufPos + 3
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_asm = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_volatile() As Boolean
    With ctx
        If pvMatchString("volatile") Then           ' "volatile"
            .BufPos = .BufPos + 8
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_volatile = True
            End If
        End If
    End With
End Function

Private Function ParseAsmOutput() As Boolean
    Dim p620 As Long

    With ctx
        If ParseCOLON() Then
            Call ParseAsmOutputList
            p620 = .BufPos
            If Not (ParseAsmInput()) Then
                .BufPos = p620
            End If
            ParseAsmOutput = True
        End If
    End With
End Function

Private Sub ParseAsmOutputList()
    Dim p1010 As Long
    Dim p1012 As Long

    With ctx
        Do
            p1010 = .BufPos
            If Not (ParseAsmOutputItem()) Then
                .BufPos = p1010
                Exit Do
            End If
            If Not (ParseCOMMA()) Then
                .BufPos = p1010
                Exit Do
            End If
        Loop
        p1012 = .BufPos
        If Not (ParseAsmOutputItem()) Then
            .BufPos = p1012
        End If
    End With
End Sub

Private Function ParseAsmInput() As Boolean
    Dim p643 As Long

    With ctx
        If ParseCOLON() Then
            Call ParseAsmInputList
            p643 = .BufPos
            If Not (ParseAsmCloppers()) Then
                .BufPos = p643
            End If
            ParseAsmInput = True
        End If
    End With
End Function

Private Function ParseAsmOutputItem() As Boolean
    Dim p636 As Long

    With ctx
        If ParseLBRACKET() Then
            If ParseIDENTIFIER() Then
                If ParseRBRACKET() Then
                    If ParseSTRINGLITERAL() Then
                        If ParseLPAREN() Then
                            p636 = .BufPos
                            If ParseMINUSRARROW() Then
                                If ParseTypeExpr() Then
                                    If ParseRPAREN() Then
                                        ParseAsmOutputItem = True
                                        Exit Function
                                    End If
                                Else
                                    .BufPos = p636
                                End If
                            Else
                                .BufPos = p636
                            End If
                            If ParseIDENTIFIER() Then
                                If ParseRPAREN() Then
                                    ParseAsmOutputItem = True
                                    Exit Function
                                End If
                            Else
                                .BufPos = p636
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseLBRACKET() As Boolean
    With ctx
        If .BufData(.BufPos) = 91 Then              ' "["
            .BufPos = .BufPos + 1
            Call Parseskip
            Call pvSetAdvance
            ParseLBRACKET = True
        End If
    End With
End Function

Private Function ParseRBRACKET() As Boolean
    With ctx
        If .BufData(.BufPos) = 93 Then              ' "]"
            .BufPos = .BufPos + 1
            Call Parseskip
            Call pvSetAdvance
            ParseRBRACKET = True
        End If
    End With
End Function

Private Function ParseMINUSRARROW() As Boolean
    With ctx
        If .BufData(.BufPos) = 45 And .BufData(.BufPos + 1) = 62 Then ' "->"
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParseMINUSRARROW = True
        End If
    End With
End Function

Private Sub ParseAsmInputList()
    Dim p1017 As Long
    Dim p1019 As Long

    With ctx
        Do
            p1017 = .BufPos
            If Not (ParseAsmInputItem()) Then
                .BufPos = p1017
                Exit Do
            End If
            If Not (ParseCOMMA()) Then
                .BufPos = p1017
                Exit Do
            End If
        Loop
        p1019 = .BufPos
        If Not (ParseAsmInputItem()) Then
            .BufPos = p1019
        End If
    End With
End Sub

Private Function ParseAsmCloppers() As Boolean
    If ParseCOLON() Then
        Call ParseStringList
        ParseAsmCloppers = True
    End If
End Function

Private Function ParseAsmInputItem() As Boolean
    If ParseLBRACKET() Then
        If ParseIDENTIFIER() Then
            If ParseRBRACKET() Then
                If ParseSTRINGLITERAL() Then
                    If ParseLPAREN() Then
                        If ParseExpr() Then
                            If ParseRPAREN() Then
                                ParseAsmInputItem = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Private Sub ParseStringList()
    Dim p1024 As Long
    Dim p1026 As Long

    With ctx
        Do
            p1024 = .BufPos
            If Not (ParseSTRINGLITERAL()) Then
                .BufPos = p1024
                Exit Do
            End If
            If Not (ParseCOMMA()) Then
                .BufPos = p1024
                Exit Do
            End If
        Loop
        p1026 = .BufPos
        If Not (ParseSTRINGLITERAL()) Then
            .BufPos = p1026
        End If
    End With
End Sub

Private Function ParseWhileContinueExpr() As Boolean
    If ParseCOLON() Then
        If ParseLPAREN() Then
            If ParseAssignExpr() Then
                If ParseRPAREN() Then
                    ParseWhileContinueExpr = True
                End If
            End If
        End If
    End If
End Function

Private Function ParseKEYWORD_section() As Boolean
    With ctx
        If pvMatchString("section") Then            ' "section"
            .BufPos = .BufPos + 7
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_section = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_nakedcc() As Boolean
    With ctx
        If pvMatchString("nakedcc") Then            ' "nakedcc"
            .BufPos = .BufPos + 7
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_nakedcc = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_stdcallcc() As Boolean
    With ctx
        If pvMatchString("stdcallcc") Then          ' "stdcallcc"
            .BufPos = .BufPos + 9
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_stdcallcc = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_async() As Boolean
    With ctx
        If pvMatchString("async") Then              ' "async"
            .BufPos = .BufPos + 5
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_async = True
            End If
        End If
    End With
End Function

Private Function ParseLARROW() As Boolean
    Dim p1357 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            p1357 = .BufPos
            Select Case .BufData(.BufPos)
            Case 60, 61                             ' [<=]
                '--- do nothing
            Case Else
                .BufPos = p1357
                Call Parseskip
                Call pvSetAdvance
                ParseLARROW = True
            End Select
        End If
    End With
End Function

Private Function ParseRARROW() As Boolean
    Dim p1442 As Long

    With ctx
        If .BufData(.BufPos) = 62 Then              ' ">"
            .BufPos = .BufPos + 1
            p1442 = .BufPos
            Select Case .BufData(.BufPos)
            Case 62, 61                             ' [>=]
                '--- do nothing
            Case Else
                .BufPos = p1442
                Call Parseskip
                Call pvSetAdvance
                ParseRARROW = True
            End Select
        End If
    End With
End Function

Private Function ParseParamDecl() As Boolean
    Dim p701 As Long
    Dim p705 As Long

    With ctx
        p701 = .BufPos
        If ParseKEYWORD_noalias() Then
            GoTo L140
        Else
            .BufPos = p701
        End If
        If ParseKEYWORD_comptime() Then
            GoTo L140
        Else
            .BufPos = p701
        End If
        .BufPos = p701
L140:
        p705 = .BufPos
        If ParseIDENTIFIER() Then
            If Not (ParseCOLON()) Then
                .BufPos = p705
            End If
        Else
            .BufPos = p705
        End If
        If ParseParamType() Then
            ParseParamDecl = True
        End If
    End With
End Function

Private Function ParseKEYWORD_noalias() As Boolean
    With ctx
        If pvMatchString("noalias") Then            ' "noalias"
            .BufPos = .BufPos + 7
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_noalias = True
            End If
        End If
    End With
End Function

Private Function ParseParamType() As Boolean
    Dim p713 As Long

    With ctx
        p713 = .BufPos
        If ParseKEYWORD_var() Then
            ParseParamType = True
            Exit Function
        Else
            .BufPos = p713
        End If
        If ParseDOT3() Then
            ParseParamType = True
            Exit Function
        Else
            .BufPos = p713
        End If
        If ParseTypeExpr() Then
            ParseParamType = True
            Exit Function
        Else
            .BufPos = p713
        End If
    End With
End Function

Private Function ParseDOT3() As Boolean
    With ctx
        If pvMatchString("...") Then                ' "..."
            .BufPos = .BufPos + 3
            Call Parseskip
            Call pvSetAdvance
            ParseDOT3 = True
        End If
    End With
End Function

Private Function ParseKEYWORD_if() As Boolean
    With ctx
        If .BufData(.BufPos) = 105 And .BufData(.BufPos + 1) = 102 Then ' "if"
            .BufPos = .BufPos + 2
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_if = True
            End If
        End If
    End With
End Function

Private Function ParsePtrPayload() As Boolean
    Dim p749 As Long

    With ctx
        If ParsePIPE() Then
            p749 = .BufPos
            If Not (ParseASTERISK()) Then
                .BufPos = p749
            End If
            If ParseIDENTIFIER() Then
                If ParsePIPE() Then
                    ParsePtrPayload = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_while() As Boolean
    With ctx
        If pvMatchString("while") Then              ' "while"
            .BufPos = .BufPos + 5
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_while = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_for() As Boolean
    With ctx
        If pvMatchString("for") Then                ' "for"
            .BufPos = .BufPos + 3
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_for = True
            End If
        End If
    End With
End Function

Private Function ParsePtrIndexPayload() As Boolean
    Dim p755 As Long
    Dim p761 As Long

    With ctx
        If ParsePIPE() Then
            p755 = .BufPos
            If Not (ParseASTERISK()) Then
                .BufPos = p755
            End If
            If ParseIDENTIFIER() Then
                p761 = .BufPos
                If ParseCOMMA() Then
                    If Not (ParseIDENTIFIER()) Then
                        .BufPos = p761
                    End If
                Else
                    .BufPos = p761
                End If
                If ParsePIPE() Then
                    ParsePtrIndexPayload = True
                End If
            End If
        End If
    End With
End Function

Private Function ParsePIPE() As Boolean
    Dim p1409 As Long

    With ctx
        If .BufData(.BufPos) = 124 Then             ' "|"
            .BufPos = .BufPos + 1
            p1409 = .BufPos
            Select Case .BufData(.BufPos)
            Case 124, 61                            ' [|=]
                '--- do nothing
            Case Else
                .BufPos = p1409
                Call Parseskip
                Call pvSetAdvance
                ParsePIPE = True
            End Select
        End If
    End With
End Function

Private Function ParseASTERISK() As Boolean
    Dim p1286 As Long

    With ctx
        If .BufData(.BufPos) = 42 Then              ' "*"
            .BufPos = .BufPos + 1
            p1286 = .BufPos
            Select Case .BufData(.BufPos)
            Case 42, 37, 61                         ' [*%=]
                '--- do nothing
            Case Else
                .BufPos = p1286
                Call Parseskip
                Call pvSetAdvance
                ParseASTERISK = True
            End Select
        End If
    End With
End Function

Private Function ParseSwitchProng() As Boolean
    Dim p770 As Long

    With ctx
        If ParseSwitchCase() Then
            If ParseEQUALRARROW() Then
                p770 = .BufPos
                If Not (ParsePtrPayload()) Then
                    .BufPos = p770
                End If
                If ParseAssignExpr() Then
                    ParseSwitchProng = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseSwitchCase() As Boolean
    Dim p783 As Long
    Dim p777 As Long
    Dim p780 As Long

    With ctx
        p783 = .BufPos
        If ParseSwitchItem() Then
            Do
                p777 = .BufPos
                If Not (ParseCOMMA()) Then
                    .BufPos = p777
                    Exit Do
                End If
                If Not (ParseSwitchItem()) Then
                    .BufPos = p777
                    Exit Do
                End If
            Loop
            p780 = .BufPos
            If Not (ParseCOMMA()) Then
                .BufPos = p780
            End If
            ParseSwitchCase = True
            Exit Function
        Else
            .BufPos = p783
        End If
        If ParseKEYWORD_else() Then
            ParseSwitchCase = True
            Exit Function
        Else
            .BufPos = p783
        End If
    End With
End Function

Private Function ParseEQUALRARROW() As Boolean
    With ctx
        If .BufData(.BufPos) = 61 And .BufData(.BufPos + 1) = 62 Then ' "=>"
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParseEQUALRARROW = True
        End If
    End With
End Function

Private Function ParseSwitchItem() As Boolean
    Dim p787 As Long

    With ctx
        If ParseExpr() Then
            p787 = .BufPos
            If ParseDOT3() Then
                If Not (ParseExpr()) Then
                    .BufPos = p787
                End If
            Else
                .BufPos = p787
            End If
            ParseSwitchItem = True
        End If
    End With
End Function

Private Function ParseASTERISKEQUAL() As Boolean
    With ctx
        If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 61 Then ' "*="
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParseASTERISKEQUAL = True
        End If
    End With
End Function

Private Function ParseSLASHEQUAL() As Boolean
    With ctx
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 61 Then ' "/="
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParseSLASHEQUAL = True
        End If
    End With
End Function

Private Function ParsePERCENTEQUAL() As Boolean
    With ctx
        If .BufData(.BufPos) = 37 And .BufData(.BufPos + 1) = 61 Then ' "%="
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParsePERCENTEQUAL = True
        End If
    End With
End Function

Private Function ParsePLUSEQUAL() As Boolean
    With ctx
        If .BufData(.BufPos) = 43 And .BufData(.BufPos + 1) = 61 Then ' "+="
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParsePLUSEQUAL = True
        End If
    End With
End Function

Private Function ParseMINUSEQUAL() As Boolean
    With ctx
        If .BufData(.BufPos) = 45 And .BufData(.BufPos + 1) = 61 Then ' "-="
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParseMINUSEQUAL = True
        End If
    End With
End Function

Private Function ParseLARROW2EQUAL() As Boolean
    With ctx
        If pvMatchString("<<=") Then                ' "<<="
            .BufPos = .BufPos + 3
            Call Parseskip
            Call pvSetAdvance
            ParseLARROW2EQUAL = True
        End If
    End With
End Function

Private Function ParseRARROW2EQUAL() As Boolean
    With ctx
        If pvMatchString(">>=") Then                ' ">>="
            .BufPos = .BufPos + 3
            Call Parseskip
            Call pvSetAdvance
            ParseRARROW2EQUAL = True
        End If
    End With
End Function

Private Function ParseAMPERSANDEQUAL() As Boolean
    With ctx
        If .BufData(.BufPos) = 38 And .BufData(.BufPos + 1) = 61 Then ' "&="
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParseAMPERSANDEQUAL = True
        End If
    End With
End Function

Private Function ParseCARETEQUAL() As Boolean
    With ctx
        If .BufData(.BufPos) = 94 And .BufData(.BufPos + 1) = 61 Then ' "^="
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParseCARETEQUAL = True
        End If
    End With
End Function

Private Function ParsePIPEEQUAL() As Boolean
    With ctx
        If .BufData(.BufPos) = 124 And .BufData(.BufPos + 1) = 61 Then ' "|="
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParsePIPEEQUAL = True
        End If
    End With
End Function

Private Function ParseASTERISKPERCENTEQUAL() As Boolean
    With ctx
        If pvMatchString("*%=") Then                ' "*%="
            .BufPos = .BufPos + 3
            Call Parseskip
            Call pvSetAdvance
            ParseASTERISKPERCENTEQUAL = True
        End If
    End With
End Function

Private Function ParsePLUSPERCENTEQUAL() As Boolean
    With ctx
        If pvMatchString("+%=") Then                ' "+%="
            .BufPos = .BufPos + 3
            Call Parseskip
            Call pvSetAdvance
            ParsePLUSPERCENTEQUAL = True
        End If
    End With
End Function

Private Function ParseMINUSPERCENTEQUAL() As Boolean
    With ctx
        If pvMatchString("-%=") Then                ' "-%="
            .BufPos = .BufPos + 3
            Call Parseskip
            Call pvSetAdvance
            ParseMINUSPERCENTEQUAL = True
        End If
    End With
End Function

Private Function ParseEQUALEQUAL() As Boolean
    With ctx
        If .BufData(.BufPos) = 61 And .BufData(.BufPos + 1) = 61 Then ' "=="
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParseEQUALEQUAL = True
        End If
    End With
End Function

Private Function ParseEXCLAMATIONMARKEQUAL() As Boolean
    With ctx
        If .BufData(.BufPos) = 33 And .BufData(.BufPos + 1) = 61 Then ' "!="
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParseEXCLAMATIONMARKEQUAL = True
        End If
    End With
End Function

Private Function ParseLARROWEQUAL() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 And .BufData(.BufPos + 1) = 61 Then ' "<="
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParseLARROWEQUAL = True
        End If
    End With
End Function

Private Function ParseRARROWEQUAL() As Boolean
    With ctx
        If .BufData(.BufPos) = 62 And .BufData(.BufPos + 1) = 61 Then ' ">="
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParseRARROWEQUAL = True
        End If
    End With
End Function

Private Function ParseAMPERSAND() As Boolean
    Dim p1278 As Long

    With ctx
        If .BufData(.BufPos) = 38 Then              ' "&"
            .BufPos = .BufPos + 1
            p1278 = .BufPos
            If Not (.BufData(.BufPos) = 61) Then    ' "="
                .BufPos = p1278
                Call Parseskip
                Call pvSetAdvance
                ParseAMPERSAND = True
            End If
        End If
    End With
End Function

Private Function ParseCARET() As Boolean
    Dim p1305 As Long

    With ctx
        If .BufData(.BufPos) = 94 Then              ' "^"
            .BufPos = .BufPos + 1
            p1305 = .BufPos
            If Not (.BufData(.BufPos) = 61) Then    ' "="
                .BufPos = p1305
                Call Parseskip
                Call pvSetAdvance
                ParseCARET = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_orelse() As Boolean
    With ctx
        If pvMatchString("orelse") Then             ' "orelse"
            .BufPos = .BufPos + 6
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_orelse = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_catch() As Boolean
    With ctx
        If pvMatchString("catch") Then              ' "catch"
            .BufPos = .BufPos + 5
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_catch = True
            End If
        End If
    End With
End Function

Private Function ParseLARROW2() As Boolean
    Dim p1362 As Long

    With ctx
        If .BufData(.BufPos) = 60 And .BufData(.BufPos + 1) = 60 Then ' "<<"
            .BufPos = .BufPos + 2
            p1362 = .BufPos
            If Not (.BufData(.BufPos) = 61) Then    ' "="
                .BufPos = p1362
                Call Parseskip
                Call pvSetAdvance
                ParseLARROW2 = True
            End If
        End If
    End With
End Function

Private Function ParseRARROW2() As Boolean
    Dim p1447 As Long

    With ctx
        If .BufData(.BufPos) = 62 And .BufData(.BufPos + 1) = 62 Then ' ">>"
            .BufPos = .BufPos + 2
            p1447 = .BufPos
            If Not (.BufData(.BufPos) = 61) Then    ' "="
                .BufPos = p1447
                Call Parseskip
                Call pvSetAdvance
                ParseRARROW2 = True
            End If
        End If
    End With
End Function

Private Function ParsePLUS() As Boolean
    Dim p1420 As Long

    With ctx
        If .BufData(.BufPos) = 43 Then              ' "+"
            .BufPos = .BufPos + 1
            p1420 = .BufPos
            Select Case .BufData(.BufPos)
            Case 37, 43, 61                         ' [%+=]
                '--- do nothing
            Case Else
                .BufPos = p1420
                Call Parseskip
                Call pvSetAdvance
                ParsePLUS = True
            End Select
        End If
    End With
End Function

Private Function ParseMINUS() As Boolean
    Dim p1382 As Long

    With ctx
        If .BufData(.BufPos) = 45 Then              ' "-"
            .BufPos = .BufPos + 1
            p1382 = .BufPos
            Select Case .BufData(.BufPos)
            Case 37, 61, 62                         ' [%=>]
                '--- do nothing
            Case Else
                .BufPos = p1382
                Call Parseskip
                Call pvSetAdvance
                ParseMINUS = True
            End Select
        End If
    End With
End Function

Private Function ParsePLUS2() As Boolean
    With ctx
        If .BufData(.BufPos) = 43 And .BufData(.BufPos + 1) = 43 Then ' "++"
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParsePLUS2 = True
        End If
    End With
End Function

Private Function ParsePLUSPERCENT() As Boolean
    Dim p1431 As Long

    With ctx
        If .BufData(.BufPos) = 43 And .BufData(.BufPos + 1) = 37 Then ' "+%"
            .BufPos = .BufPos + 2
            p1431 = .BufPos
            If Not (.BufData(.BufPos) = 61) Then    ' "="
                .BufPos = p1431
                Call Parseskip
                Call pvSetAdvance
                ParsePLUSPERCENT = True
            End If
        End If
    End With
End Function

Private Function ParseMINUSPERCENT() As Boolean
    Dim p1390 As Long

    With ctx
        If .BufData(.BufPos) = 45 And .BufData(.BufPos + 1) = 37 Then ' "-%"
            .BufPos = .BufPos + 2
            p1390 = .BufPos
            If Not (.BufData(.BufPos) = 61) Then    ' "="
                .BufPos = p1390
                Call Parseskip
                Call pvSetAdvance
                ParseMINUSPERCENT = True
            End If
        End If
    End With
End Function

Private Function ParsePIPE2() As Boolean
    With ctx
        If .BufData(.BufPos) = 124 And .BufData(.BufPos + 1) = 124 Then ' "||"
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParsePIPE2 = True
        End If
    End With
End Function

Private Function ParseSLASH() As Boolean
    Dim p1470 As Long

    With ctx
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            p1470 = .BufPos
            If Not (.BufData(.BufPos) = 61) Then    ' "="
                .BufPos = p1470
                Call Parseskip
                Call pvSetAdvance
                ParseSLASH = True
            End If
        End If
    End With
End Function

Private Function ParsePERCENT() As Boolean
    Dim p1401 As Long

    With ctx
        If .BufData(.BufPos) = 37 Then              ' "%"
            .BufPos = .BufPos + 1
            p1401 = .BufPos
            If Not (.BufData(.BufPos) = 61) Then    ' "="
                .BufPos = p1401
                Call Parseskip
                Call pvSetAdvance
                ParsePERCENT = True
            End If
        End If
    End With
End Function

Private Function ParseASTERISK2() As Boolean
    With ctx
        If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 42 Then ' "**"
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParseASTERISK2 = True
        End If
    End With
End Function

Private Function ParseASTERISKPERCENT() As Boolean
    Dim p1297 As Long

    With ctx
        If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 37 Then ' "*%"
            .BufPos = .BufPos + 2
            p1297 = .BufPos
            If Not (.BufData(.BufPos) = 61) Then    ' "="
                .BufPos = p1297
                Call Parseskip
                Call pvSetAdvance
                ParseASTERISKPERCENT = True
            End If
        End If
    End With
End Function

Private Function ParseTILDE() As Boolean
    With ctx
        If .BufData(.BufPos) = 126 Then             ' "~"
            .BufPos = .BufPos + 1
            Call Parseskip
            Call pvSetAdvance
            ParseTILDE = True
        End If
    End With
End Function

Private Function ParseKEYWORD_await() As Boolean
    With ctx
        If pvMatchString("await") Then              ' "await"
            .BufPos = .BufPos + 5
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_await = True
            End If
        End If
    End With
End Function

Private Function ParseQUESTIONMARK() As Boolean
    With ctx
        If .BufData(.BufPos) = 63 Then              ' "?"
            .BufPos = .BufPos + 1
            Call Parseskip
            Call pvSetAdvance
            ParseQUESTIONMARK = True
        End If
    End With
End Function

Private Function ParseArrayTypeStart() As Boolean
    Dim p944 As Long

    With ctx
        If ParseLBRACKET() Then
            p944 = .BufPos
            If Not (ParseExpr()) Then
                .BufPos = p944
            End If
            If ParseRBRACKET() Then
                ParseArrayTypeStart = True
            End If
        End If
    End With
End Function

Private Function ParsePtrTypeStart() As Boolean
    Dim p950 As Long

    With ctx
        p950 = .BufPos
        If ParseASTERISK() Then
            ParsePtrTypeStart = True
            Exit Function
        Else
            .BufPos = p950
        End If
        If ParseASTERISK2() Then
            ParsePtrTypeStart = True
            Exit Function
        Else
            .BufPos = p950
        End If
        If ParseLBRACKET() Then
            If ParseASTERISK() Then
                If ParseRBRACKET() Then
                    ParsePtrTypeStart = True
                    Exit Function
                Else
                    .BufPos = p950
                End If
            Else
                .BufPos = p950
            End If
        Else
            .BufPos = p950
        End If
    End With
End Function

Private Function ParseKEYWORD_align() As Boolean
    With ctx
        If pvMatchString("align") Then              ' "align"
            .BufPos = .BufPos + 5
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_align = True
            End If
        End If
    End With
End Function

Private Function ParseDOT2() As Boolean
    Dim p1324 As Long

    With ctx
        If .BufData(.BufPos) = 46 And .BufData(.BufPos + 1) = 46 Then ' ".."
            .BufPos = .BufPos + 2
            p1324 = .BufPos
            If Not (.BufData(.BufPos) = 46) Then    ' "."
                .BufPos = p1324
                Call Parseskip
                Call pvSetAdvance
                ParseDOT2 = True
            End If
        End If
    End With
End Function

Private Function ParseDOTASTERISK() As Boolean
    With ctx
        If .BufData(.BufPos) = 46 And .BufData(.BufPos + 1) = 42 Then ' ".*"
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParseDOTASTERISK = True
        End If
    End With
End Function

Private Function ParseDOTQUESTIONMARK() As Boolean
    With ctx
        If .BufData(.BufPos) = 46 And .BufData(.BufPos + 1) = 63 Then ' ".?"
            .BufPos = .BufPos + 2
            Call Parseskip
            Call pvSetAdvance
            ParseDOTQUESTIONMARK = True
        End If
    End With
End Function

Private Sub ParseExprList()
    Dim p1038 As Long
    Dim p1040 As Long

    With ctx
        Do
            p1038 = .BufPos
            If Not (ParseExpr()) Then
                .BufPos = p1038
                Exit Do
            End If
            If Not (ParseCOMMA()) Then
                .BufPos = p1038
                Exit Do
            End If
        Loop
        p1040 = .BufPos
        If Not (ParseExpr()) Then
            .BufPos = p1040
        End If
    End With
End Sub

Private Function ParseContainerDeclType() As Boolean
    Dim p988 As Long
    Dim p965 As Long
    Dim p968 As Long
    Dim p984 As Long
    Dim p983 As Long
    Dim p978 As Long

    With ctx
        p988 = .BufPos
        p965 = .BufPos
        If ParseKEYWORD_struct() Then
            p968 = .BufPos
            If ParseLPAREN() Then
                If ParseExpr() Then
                    If Not (ParseRPAREN()) Then
                        .BufPos = p968
                    End If
                Else
                    .BufPos = p968
                End If
            Else
                .BufPos = p968
            End If
            ParseContainerDeclType = True
            Exit Function
        Else
            .BufPos = p965
        End If
        If ParseKEYWORD_enum() Then
            p968 = .BufPos
            If ParseLPAREN() Then
                If ParseExpr() Then
                    If Not (ParseRPAREN()) Then
                        .BufPos = p968
                    End If
                Else
                    .BufPos = p968
                End If
            Else
                .BufPos = p968
            End If
            ParseContainerDeclType = True
            Exit Function
        Else
            .BufPos = p965
        End If
        .BufPos = p988
        If ParseKEYWORD_union() Then
            p984 = .BufPos
            If ParseLPAREN() Then
                p983 = .BufPos
                If ParseKEYWORD_enum() Then
                    p978 = .BufPos
                    If ParseLPAREN() Then
                        If ParseExpr() Then
                            If Not (ParseRPAREN()) Then
                                .BufPos = p978
                            End If
                        Else
                            .BufPos = p978
                        End If
                    Else
                        .BufPos = p978
                    End If
                    If ParseRPAREN() Then
                        GoTo L153
                    Else
                        .BufPos = p984
                    End If
                    GoTo L154
                Else
                    .BufPos = p983
                End If
                If ParseExpr() Then
                    If ParseRPAREN() Then
                        GoTo L153
                    Else
                        .BufPos = p984
                    End If
                    GoTo L154
                Else
                    .BufPos = p983
                End If
                .BufPos = p984
L154:
            Else
                .BufPos = p984
            End If
L153:
            ParseContainerDeclType = True
            Exit Function
        Else
            .BufPos = p988
        End If
    End With
End Function

Private Function ParseKEYWORD_struct() As Boolean
    With ctx
        If pvMatchString("struct") Then             ' "struct"
            .BufPos = .BufPos + 6
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_struct = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_enum() As Boolean
    With ctx
        If pvMatchString("enum") Then               ' "enum"
            .BufPos = .BufPos + 4
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_enum = True
            End If
        End If
    End With
End Function

Private Function ParseKEYWORD_union() As Boolean
    With ctx
        If pvMatchString("union") Then              ' "union"
            .BufPos = .BufPos + 5
            If Parseend_of_word() Then
                Call pvSetAdvance
                ParseKEYWORD_union = True
            End If
        End If
    End With
End Function

Private Function Parsechar_escape() As Boolean
    Dim p1058 As Long

    With ctx
        p1058 = .BufPos
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 120 Then ' "\\x"
            .BufPos = .BufPos + 2
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
                Select Case .BufData(.BufPos)
                Case 48 To 57, 97 To 102, 65 To 70  ' [0-9a-fA-F]
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    Parsechar_escape = True
                    Exit Function
                End Select
            End Select
        End If
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 117 Then ' "\\u"
            .BufPos = .BufPos + 2
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
                Select Case .BufData(.BufPos)
                Case 48 To 57, 97 To 102, 65 To 70  ' [0-9a-fA-F]
                    .BufPos = .BufPos + 1
                    Select Case .BufData(.BufPos)
                    Case 48 To 57, 97 To 102, 65 To 70 ' [0-9a-fA-F]
                        .BufPos = .BufPos + 1
                        Select Case .BufData(.BufPos)
                        Case 48 To 57, 97 To 102, 65 To 70 ' [0-9a-fA-F]
                            .BufPos = .BufPos + 1
                            Call pvSetAdvance
                            Parsechar_escape = True
                            Exit Function
                        End Select
                    End Select
                End Select
            End Select
        End If
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 85 Then ' "\\U"
            .BufPos = .BufPos + 2
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
                Select Case .BufData(.BufPos)
                Case 48 To 57, 97 To 102, 65 To 70  ' [0-9a-fA-F]
                    .BufPos = .BufPos + 1
                    Select Case .BufData(.BufPos)
                    Case 48 To 57, 97 To 102, 65 To 70 ' [0-9a-fA-F]
                        .BufPos = .BufPos + 1
                        Select Case .BufData(.BufPos)
                        Case 48 To 57, 97 To 102, 65 To 70 ' [0-9a-fA-F]
                            .BufPos = .BufPos + 1
                            Select Case .BufData(.BufPos)
                            Case 48 To 57, 97 To 102, 65 To 70 ' [0-9a-fA-F]
                                .BufPos = .BufPos + 1
                                Select Case .BufData(.BufPos)
                                Case 48 To 57, 97 To 102, 65 To 70 ' [0-9a-fA-F]
                                    .BufPos = .BufPos + 1
                                    Call pvSetAdvance
                                    Parsechar_escape = True
                                    Exit Function
                                End Select
                            End Select
                        End Select
                    End Select
                End Select
            End Select
        End If
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 110, 114, 92, 116, 39, 34          ' [nr\\t'"]
                .BufPos = .BufPos + 1
                Call pvSetAdvance
                Parsechar_escape = True
            End Select
        End If
    End With
End Function

Private Function Parsechar_char() As Boolean
    Dim p1073 As Long

    With ctx
        p1073 = .BufPos
        If Parsechar_escape() Then
            Call pvSetAdvance
            Parsechar_char = True
            Exit Function
        Else
            .BufPos = p1073
        End If
        Select Case .BufData(.BufPos)
        Case 92, 39, 13, 10                         ' [\\'\r\n]
            '--- do nothing
        Case Else
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
                Call pvSetAdvance
                Parsechar_char = True
            End If
        End Select
    End With
End Function

Private Function Parsestring_char() As Boolean
    Dim p1077 As Long

    With ctx
        p1077 = .BufPos
        If Parsechar_escape() Then
            Call pvSetAdvance
            Parsestring_char = True
            Exit Function
        Else
            .BufPos = p1077
        End If
        Select Case .BufData(.BufPos)
        Case 92, 34, 13, 10                         ' [\\"\r\n]
            '--- do nothing
        Case Else
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
                Call pvSetAdvance
                Parsestring_char = True
            End If
        End Select
    End With
End Function

Private Function Parseline_comment() As Boolean
    With ctx
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 47 Then ' "//"
            .BufPos = .BufPos + 2
            Do
                Select Case .BufData(.BufPos)
                Case 13, 10                         ' [\r\n]
                    Exit Do
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                    Else
                        Exit Do
                    End If
                End Select
            Loop
            Call pvSetAdvance
            Parseline_comment = True
        End If
    End With
End Function

Private Function Parseline_string() As Boolean
    Dim i1090 As Long
    Dim p1087 As Long

    With ctx
        For i1090 = 0 To LNG_MAXINT
            p1087 = .BufPos
            If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 92 Then ' "\\\\"
                .BufPos = .BufPos + 2
            Else
                Exit For
            End If
            Do
                Select Case .BufData(.BufPos)
                Case 13, 10                         ' [\r\n]
                    Exit Do
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                    Else
                        Exit Do
                    End If
                End Select
            Loop
            Do
                Select Case .BufData(.BufPos)
                Case 32, 13, 10                     ' [ \r\n]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit Do
                End Select
            Loop
        Next
        If i1090 <> 0 Then
            Call pvSetAdvance
            Parseline_string = True
        End If
    End With
End Function

Private Function Parseline_cstring() As Boolean
    Dim i1098 As Long
    Dim p1095 As Long

    With ctx
        For i1098 = 0 To LNG_MAXINT
            p1095 = .BufPos
            If pvMatchString("c\\") Then            ' "c\\\\"
                .BufPos = .BufPos + 3
            Else
                Exit For
            End If
            Do
                Select Case .BufData(.BufPos)
                Case 13, 10                         ' [\r\n]
                    Exit Do
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                    Else
                        Exit Do
                    End If
                End Select
            Loop
            Do
                Select Case .BufData(.BufPos)
                Case 32, 13, 10                     ' [ \r\n]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit Do
                End Select
            Loop
        Next
        If i1098 <> 0 Then
            Call pvSetAdvance
            Parseline_cstring = True
        End If
    End With
End Function

Private Function Parsekeyword() As Boolean
    Dim p1633 As Long

    With ctx
        p1633 = .BufPos
        If ParseKEYWORD_align() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_and() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_anyerror() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_asm() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_async() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_await() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_break() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_cancel() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_catch() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_comptime() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_const() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_continue() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_defer() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_else() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_enum() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_errdefer() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_error() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_export() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_extern() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_false() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_fn() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_for() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_if() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_inline() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_nakedcc() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_noalias() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_null() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_or() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_orelse() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_packed() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_promise() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_pub() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_resume() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_return() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_section() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_stdcallcc() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_struct() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_suspend() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_switch() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_test() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_true() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_try() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_undefined() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_union() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_unreachable() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_use() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_var() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_volatile() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
        If ParseKEYWORD_while() Then
            Parsekeyword = True
            Exit Function
        Else
            .BufPos = p1633
        End If
    End With
End Function

Private Function Parseend_of_word() As Boolean
    Dim p1481 As Long

    With ctx
        p1481 = .BufPos
        Select Case .BufData(.BufPos)
        Case 97 To 122, 65 To 90, 48 To 57, 95      ' [a-zA-Z0-9_]
            '--- do nothing
        Case Else
            .BufPos = p1481
            Call Parseskip
            Call pvSetAdvance
            Parseend_of_word = True
        End Select
    End With
End Function

