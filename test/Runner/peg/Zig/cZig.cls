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
' Auto-generated on 5.3.2019 16:44:42
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Const LOCALE_USER_DEFAULT           As Long = &H400
Private Const NORM_IGNORECASE               As Long = 1
Private Const CSTR_EQUAL                    As Long = 2

Private Declare Function CompareStringW Lib "kernel32" (ByVal Locale As Long, ByVal dwCmpFlags As Long, lpString1 As Any, ByVal cchCount1 As Long, lpString2 As Any, ByVal cchCount2 As Long) As Long
#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function ArrPtr Lib "vbe7" Alias "VarPtr" (Ptr() As Any) As LongPtr
Private Const NULL_PTR                  As LongPtr = 0
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Const NULL_PTR                  As Long = 0
#End If
#If Win64 Then
Private Const PTR_SIZE                  As Long = 8
#Else
Private Const PTR_SIZE                  As Long = 4
#End If

Private Type SAFEARRAY1D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
#If VBA7 Then
    pvData              As LongPtr
#Else
    pvData              As Long
#End If
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
    ParserVersion = "5.3.2019 16:44:42"
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
        Call CopyMemory(ByVal VarPtr(.Contents), ByVal VarPtr(sSubject), PTR_SIZE)
        With .BufArray
            .cDims = 1
            .fFeatures = 1 ' FADF_AUTO
            .cbElements = 2
            .pvData = StrPtr(sSubject)
            .cElements = Len(sSubject) + 2 '-- look-ahead chars
        End With
        Call CopyMemory(ByVal ArrPtr(.BufData), VarPtr(.BufArray), PTR_SIZE)
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
        Call CopyMemory(ByVal VarPtr(.Contents), NULL_PTR, PTR_SIZE)
        Call CopyMemory(ByVal ArrPtr(.BufData), NULL_PTR, PTR_SIZE)
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
            Exit Sub
        End If
        .BufPos = p17
        If ParseTopLevelComptime() Then
            Call ParseContainerMembers
            Exit Sub
        End If
        .BufPos = p17
        p19 = .BufPos
        If Not (ParseKEYWORD_pub()) Then
            .BufPos = p19
        End If
        If ParseTopLevelDecl() Then
            Call ParseContainerMembers
            Exit Sub
        End If
        .BufPos = p17
        p25 = .BufPos
        If Not (ParseKEYWORD_pub()) Then
            .BufPos = p25
        End If
        If ParseContainerField() Then
            If ParseCOMMA() Then
                Call ParseContainerMembers
                Exit Sub
            End If
            .BufPos = p17
        Else
            .BufPos = p17
        End If
        p33 = .BufPos
        If Not (ParseKEYWORD_pub()) Then
            .BufPos = p33
        End If
        If ParseContainerField() Then
            Exit Sub
        End If
        .BufPos = p17
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
        If Not (ParseKEYWORD_export()) Then
            .BufPos = p57
            If Not (ParseKEYWORD_extern()) Then
                .BufPos = p57
                GoTo L20
            End If
            p54 = .BufPos
            If Not (ParseSTRINGLITERAL()) Then
                .BufPos = p54
            End If
            GoTo L18
L20:
            If Not (ParseKEYWORD_inline()) Then
                .BufPos = p57
            End If
        End If
L18:
        If ParseFnProto() Then
            p67 = .BufPos
            If ParseSEMICOLON() Then
                GoTo L15
            End If
            .BufPos = p67
            If ParseBlock() Then
                GoTo L15
            End If
            .BufPos = p67
            .BufPos = p78
            GoTo L16
        End If
        .BufPos = p78
L16:
        p73 = .BufPos
        If Not (ParseKEYWORD_export()) Then
            .BufPos = p73
            If Not (ParseKEYWORD_extern()) Then
                .BufPos = p73
                GoTo L25
            End If
            p70 = .BufPos
            If Not (ParseSTRINGLITERAL()) Then
                .BufPos = p70
            End If
            GoTo L23
L25:
            .BufPos = p73
        End If
L23:
        If ParseVarDecl() Then
            GoTo L15
        End If
        .BufPos = p78
        If ParseKEYWORD_use() Then
            If ParseExpr() Then
                If ParseSEMICOLON() Then
                    GoTo L15
                End If
                .BufPos = p78
            Else
                .BufPos = p78
            End If
        Else
            .BufPos = p78
        End If
        Exit Function
L15:
        ParseTopLevelDecl = True
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
        Else
            .BufPos = p1246
            GoTo L38
        End If
        Do
            p1239 = .BufPos
            If Not (Parsestring_char()) Then
                .BufPos = p1239
                Exit Do
            End If
        Loop
        If .BufData(.BufPos) = 34 Then              ' "\""
            .BufPos = .BufPos + 1
        Else
            .BufPos = p1246
            GoTo L38
        End If
        Call Parseskip
        Call pvSetAdvance
        ParseSTRINGLITERAL = True
        Exit Function
L38:
        If Not (Parseline_string()) Then
            .BufPos = p1246
            GoTo L37
        End If
        Call Parseskip
        Call pvSetAdvance
        ParseSTRINGLITERAL = True
        Exit Function
L37:
        If Not (Parseline_cstring()) Then
            .BufPos = p1246
            GoTo L36
        End If
        Call Parseskip
        Call pvSetAdvance
        ParseSTRINGLITERAL = True
        Exit Function
L36:
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
                        GoTo L48
                    End If
                    .BufPos = p113
                    If ParseTypeExpr() Then
                        GoTo L48
                    End If
                    .BufPos = p113
                    Exit Function
L48:
                    ParseFnProto = True
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
            GoTo L62
        End If
        .BufPos = p117
        If ParseKEYWORD_var() Then
            GoTo L62
        End If
        .BufPos = p117
        Exit Function
L62:
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
            End If
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
            GoTo L65
        End If
        .BufPos = p685
        If ParseKEYWORD_stdcallcc() Then
            GoTo L65
        End If
        .BufPos = p685
        If ParseKEYWORD_extern() Then
            GoTo L65
        End If
        .BufPos = p685
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
            GoTo L65
        End If
        .BufPos = p685
        Exit Function
L65:
        ParseFnCC = True
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
            GoTo L72
        End If
        .BufPos = p1252
        .LastExpected = e1252
        p1258 = .BufPos
        If .BufData(.BufPos) = 99 Then              ' "c"
            .BufPos = .BufPos + 1
        Else
            GoTo L74
        End If
        p1255 = .BufPos
        Select Case .BufData(.BufPos)
        Case 34, 92                                 ' ["\\]
            GoTo L74
        Case Else
            .BufPos = p1255
        End Select
        GoTo L73
L74:
        Select Case .BufData(.BufPos)
        Case 65 To 90, 97, 98, 100 To 122, 95       ' [A-Zabd-z_]
            .BufPos = .BufPos + 1
        Case Else
            .BufPos = p1269
            GoTo L72
        End Select
L73:
        Do
            Select Case .BufData(.BufPos)
            Case 65 To 90, 97 To 122, 48 To 57, 95  ' [A-Za-z0-9_]
                .BufPos = .BufPos + 1
            Case Else
                Exit Do
            End Select
        Loop
        Call Parseskip
        Call pvSetAdvance
        ParseIDENTIFIER = True
        Exit Function
L72:
        If .BufData(.BufPos) = 64 And .BufData(.BufPos + 1) = 34 Then ' "@\""
            .BufPos = .BufPos + 2
        Else
            GoTo L71
        End If
        Do
            p1264 = .BufPos
            If Not (Parsestring_char()) Then
                .BufPos = p1264
                Exit Do
            End If
        Loop
        If .BufData(.BufPos) = 34 Then              ' "\""
            .BufPos = .BufPos + 1
        Else
            .BufPos = p1269
            GoTo L71
        End If
        Call Parseskip
        Call pvSetAdvance
        ParseIDENTIFIER = True
        Exit Function
L71:
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
            GoTo L87
        End If
        .BufPos = p154
        If ParseKEYWORD_comptime() Then
            If ParseBlockExprStatement() Then
                GoTo L87
            End If
            .BufPos = p154
        Else
            .BufPos = p154
        End If
        If ParseKEYWORD_suspend() Then
            p159 = .BufPos
            If ParseSEMICOLON() Then
                GoTo L87
            End If
            .BufPos = p159
            If ParseBlockExprStatement() Then
                GoTo L87
            End If
            .BufPos = p159
            .BufPos = p154
            GoTo L91
        End If
        .BufPos = p154
L91:
        If ParseKEYWORD_defer() Then
            If ParseBlockExprStatement() Then
                GoTo L87
            End If
            .BufPos = p154
        Else
            .BufPos = p154
        End If
        If ParseKEYWORD_errdefer() Then
            If ParseBlockExprStatement() Then
                GoTo L87
            End If
            .BufPos = p154
        Else
            .BufPos = p154
        End If
        If ParseIfStatement() Then
            GoTo L87
        End If
        .BufPos = p154
        If ParseLabeledStatement() Then
            GoTo L87
        End If
        .BufPos = p154
        If ParseSwitchExpr() Then
            GoTo L87
        End If
        .BufPos = p154
        If ParseAssignExpr() Then
            If ParseSEMICOLON() Then
                GoTo L87
            End If
            .BufPos = p154
        Else
            .BufPos = p154
        End If
        Exit Function
L87:
        ParseStatement = True
    End With
End Function

Private Function ParseBlockExprStatement() As Boolean
    Dim p260 As Long

    With ctx
        p260 = .BufPos
        If ParseBlockExpr() Then
            GoTo L96
        End If
        .BufPos = p260
        If ParseAssignExpr() Then
            If ParseSEMICOLON() Then
                GoTo L96
            End If
            .BufPos = p260
        Else
            .BufPos = p260
        End If
        Exit Function
L96:
        ParseBlockExprStatement = True
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
                GoTo L101
            End If
            .BufPos = p201
        Else
            .BufPos = p201
        End If
        If ParseIfPrefix() Then
            If ParseAssignExpr() Then
                p200 = .BufPos
                If ParseSEMICOLON() Then
                    GoTo L101
                End If
                .BufPos = p200
                If ParseKEYWORD_else() Then
                    p196 = .BufPos
                    If Not (ParsePayload()) Then
                        .BufPos = p196
                    End If
                    If ParseStatement() Then
                        GoTo L101
                    End If
                    .BufPos = p200
                Else
                    .BufPos = p200
                End If
                .BufPos = p201
                GoTo L106
            End If
            .BufPos = p201
        Else
            .BufPos = p201
        End If
L106:
        Exit Function
L101:
        ParseIfStatement = True
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
            GoTo L111
        End If
        .BufPos = p208
        If ParseLoopStatement() Then
            GoTo L111
        End If
        .BufPos = p208
        Exit Function
L111:
        ParseLabeledStatement = True
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
            GoTo L123
        End If
        .BufPos = p216
        If ParseWhileStatement() Then
            GoTo L123
        End If
        .BufPos = p216
        Exit Function
L123:
        ParseLoopStatement = True
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
                GoTo L125
            End If
            .BufPos = p234
        Else
            .BufPos = p234
        End If
        If ParseForPrefix() Then
            If ParseAssignExpr() Then
                p233 = .BufPos
                If ParseSEMICOLON() Then
                    GoTo L125
                End If
                .BufPos = p233
                If ParseKEYWORD_else() Then
                    If ParseStatement() Then
                        GoTo L125
                    End If
                    .BufPos = p233
                Else
                    .BufPos = p233
                End If
                .BufPos = p234
                GoTo L129
            End If
            .BufPos = p234
        Else
            .BufPos = p234
        End If
L129:
        Exit Function
L125:
        ParseForStatement = True
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
                GoTo L132
            End If
            .BufPos = p255
        Else
            .BufPos = p255
        End If
        If ParseWhilePrefix() Then
            If ParseAssignExpr() Then
                p254 = .BufPos
                If ParseSEMICOLON() Then
                    GoTo L132
                End If
                .BufPos = p254
                If ParseKEYWORD_else() Then
                    p250 = .BufPos
                    If Not (ParsePayload()) Then
                        .BufPos = p250
                    End If
                    If ParseStatement() Then
                        GoTo L132
                    End If
                    .BufPos = p254
                Else
                    .BufPos = p254
                End If
                .BufPos = p255
                GoTo L137
            End If
            .BufPos = p255
        Else
            .BufPos = p255
        End If
L137:
        Exit Function
L132:
        ParseWhileStatement = True
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
            GoTo L145
        End If
        .BufPos = p794
        If ParseSLASHEQUAL() Then
            GoTo L145
        End If
        .BufPos = p794
        If ParsePERCENTEQUAL() Then
            GoTo L145
        End If
        .BufPos = p794
        If ParsePLUSEQUAL() Then
            GoTo L145
        End If
        .BufPos = p794
        If ParseMINUSEQUAL() Then
            GoTo L145
        End If
        .BufPos = p794
        If ParseLARROW2EQUAL() Then
            GoTo L145
        End If
        .BufPos = p794
        If ParseRARROW2EQUAL() Then
            GoTo L145
        End If
        .BufPos = p794
        If ParseAMPERSANDEQUAL() Then
            GoTo L145
        End If
        .BufPos = p794
        If ParseCARETEQUAL() Then
            GoTo L145
        End If
        .BufPos = p794
        If ParsePIPEEQUAL() Then
            GoTo L145
        End If
        .BufPos = p794
        If ParseASTERISKPERCENTEQUAL() Then
            GoTo L145
        End If
        .BufPos = p794
        If ParsePLUSPERCENTEQUAL() Then
            GoTo L145
        End If
        .BufPos = p794
        If ParseMINUSPERCENTEQUAL() Then
            GoTo L145
        End If
        .BufPos = p794
        If ParseEQUAL() Then
            GoTo L145
        End If
        .BufPos = p794
        Exit Function
L145:
        ParseAssignOp = True
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
            GoTo L155
        End If
        .BufPos = p822
        If ParseEXCLAMATIONMARKEQUAL() Then
            GoTo L155
        End If
        .BufPos = p822
        If ParseLARROW() Then
            GoTo L155
        End If
        .BufPos = p822
        If ParseRARROW() Then
            GoTo L155
        End If
        .BufPos = p822
        If ParseLARROWEQUAL() Then
            GoTo L155
        End If
        .BufPos = p822
        If ParseRARROWEQUAL() Then
            GoTo L155
        End If
        .BufPos = p822
        Exit Function
L155:
        ParseCompareOp = True
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
            GoTo L157
        End If
        .BufPos = p833
        If ParseCARET() Then
            GoTo L157
        End If
        .BufPos = p833
        If ParsePIPE() Then
            GoTo L157
        End If
        .BufPos = p833
        If ParseKEYWORD_orelse() Then
            GoTo L157
        End If
        .BufPos = p833
        If ParseKEYWORD_catch() Then
            p839 = .BufPos
            If Not (ParsePayload()) Then
                .BufPos = p839
            End If
            GoTo L157
        End If
        .BufPos = p833
        Exit Function
L157:
        ParseBitwiseOp = True
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
            GoTo L161
        End If
        .BufPos = p846
        If ParseRARROW2() Then
            GoTo L161
        End If
        .BufPos = p846
        Exit Function
L161:
        ParseBitShiftOp = True
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
            GoTo L163
        End If
        .BufPos = p851
        If ParseMINUS() Then
            GoTo L163
        End If
        .BufPos = p851
        If ParsePLUS2() Then
            GoTo L163
        End If
        .BufPos = p851
        If ParsePLUSPERCENT() Then
            GoTo L163
        End If
        .BufPos = p851
        If ParseMINUSPERCENT() Then
            GoTo L163
        End If
        .BufPos = p851
        Exit Function
L163:
        ParseAdditionOp = True
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
            GoTo L165
        End If
        .BufPos = p861
        If ParseASTERISK() Then
            GoTo L165
        End If
        .BufPos = p861
        If ParseSLASH() Then
            GoTo L165
        End If
        .BufPos = p861
        If ParsePERCENT() Then
            GoTo L165
        End If
        .BufPos = p861
        If ParseASTERISK2() Then
            GoTo L165
        End If
        .BufPos = p861
        If ParseASTERISKPERCENT() Then
            GoTo L165
        End If
        .BufPos = p861
        Exit Function
L165:
        ParseMultiplyOp = True
    End With
End Function

Private Function ParsePrefixOp() As Boolean
    Dim p872 As Long

    With ctx
        p872 = .BufPos
        If ParseEXCLAMATIONMARK() Then
            GoTo L166
        End If
        .BufPos = p872
        If ParseMINUS() Then
            GoTo L166
        End If
        .BufPos = p872
        If ParseTILDE() Then
            GoTo L166
        End If
        .BufPos = p872
        If ParseMINUSPERCENT() Then
            GoTo L166
        End If
        .BufPos = p872
        If ParseAMPERSAND() Then
            GoTo L166
        End If
        .BufPos = p872
        If ParseKEYWORD_try() Then
            GoTo L166
        End If
        .BufPos = p872
        If ParseKEYWORD_await() Then
            GoTo L166
        End If
        .BufPos = p872
        Exit Function
L166:
        ParsePrefixOp = True
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
        End If
        .BufPos = p344
        If ParseIfExpr() Then
            ParsePrimaryExpr = True
            Exit Function
        End If
        .BufPos = p344
        If Not (ParseKEYWORD_break()) Then
            .BufPos = p344
            GoTo L175
        End If
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
L175:
        If Not (ParseKEYWORD_cancel()) Then
            .BufPos = p344
            GoTo L174
        End If
        If Not (ParseExpr()) Then
            .BufPos = p344
            GoTo L174
        End If
        ParsePrimaryExpr = True
        Exit Function
L174:
        If Not (ParseKEYWORD_comptime()) Then
            .BufPos = p344
            GoTo L173
        End If
        If Not (ParseExpr()) Then
            .BufPos = p344
            GoTo L173
        End If
        ParsePrimaryExpr = True
        Exit Function
L173:
        If Not (ParseKEYWORD_continue()) Then
            .BufPos = p344
            GoTo L171
        End If
        p362 = .BufPos
        If Not (ParseBreakLabel()) Then
            .BufPos = p362
        End If
        ParsePrimaryExpr = True
        Exit Function
L171:
        If Not (ParseKEYWORD_resume()) Then
            .BufPos = p344
            GoTo L170
        End If
        If Not (ParseExpr()) Then
            .BufPos = p344
            GoTo L170
        End If
        ParsePrimaryExpr = True
        Exit Function
L170:
        If Not (ParseKEYWORD_return()) Then
            .BufPos = p344
            GoTo L168
        End If
        p371 = .BufPos
        If Not (ParseExpr()) Then
            .BufPos = p371
        End If
        ParsePrimaryExpr = True
        Exit Function
L168:
        If ParseLabeledExpr() Then
            ParsePrimaryExpr = True
            Exit Function
        End If
        .BufPos = p344
        If ParseCurlySuffixExpr() Then
            ParsePrimaryExpr = True
            Exit Function
        End If
        .BufPos = p344
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
            GoTo L192
        End If
        .BufPos = p392
        If ParseLoopExpr() Then
            GoTo L192
        End If
        .BufPos = p392
        Exit Function
L192:
        ParseLabeledExpr = True
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
            GoTo L197
        End If
        .BufPos = p407
        If ParseWhileExpr() Then
            GoTo L197
        End If
        .BufPos = p407
        Exit Function
L197:
        ParseLoopExpr = True
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
                    GoTo L208
                End If
                .BufPos = p451
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
                    GoTo L208
                End If
                .BufPos = p451
            Else
                .BufPos = p451
            End If
        Else
            .BufPos = p451
        End If
        If ParseLBRACE() Then
            If ParseRBRACE() Then
                GoTo L208
            End If
            .BufPos = p451
        Else
            .BufPos = p451
        End If
        Exit Function
L208:
        ParseInitList = True
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
            GoTo L215
        End If
        .BufPos = p885
        If ParseKEYWORD_promise() Then
            If ParseMINUSRARROW() Then
                GoTo L215
            End If
            .BufPos = p885
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
                            Exit Do
                        End If
                    End If
                End If
            Loop
            GoTo L215
        End If
        .BufPos = p885
        If ParsePtrTypeStart() Then
            Do
                p909 = .BufPos
                If Not (ParseKEYWORD_align()) Then
                    .BufPos = p909
                    GoTo L221
                End If
                If Not (ParseLPAREN()) Then
                    .BufPos = p909
                    GoTo L221
                End If
                If Not (ParseExpr()) Then
                    .BufPos = p909
                    GoTo L221
                End If
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
                    GoTo L221
                End If
                GoTo L220
L221:
                If Not (ParseKEYWORD_const()) Then
                    .BufPos = p909
                    If Not (ParseKEYWORD_volatile()) Then
                        .BufPos = p909
                        Exit Do
                    End If
                End If
L220:
            Loop
            GoTo L215
        End If
        .BufPos = p885
        Exit Function
L215:
        ParsePrefixTypeOp = True
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
                    GoTo L227
                End If
                .BufPos = p484
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
                        Exit Do
                    End If
                End If
            Loop
            GoTo L227
        End If
        .BufPos = p484
        Exit Function
L227:
        ParseSuffixExpr = True
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
                GoTo L234
            End If
            .BufPos = p491
        Else
            .BufPos = p491
        End If
        If ParseCHAR_LITERAL() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseContainerDecl() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseErrorSetDecl() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseFLOAT() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseFnProto() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseGroupedExpr() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseIDENTIFIER() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseIfTypeExpr() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseINTEGER() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseKEYWORD_anyerror() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseKEYWORD_comptime() Then
            If ParseTypeExpr() Then
                GoTo L234
            End If
            .BufPos = p491
        Else
            .BufPos = p491
        End If
        If ParseKEYWORD_error() Then
            If ParseDOT() Then
                If ParseIDENTIFIER() Then
                    GoTo L234
                End If
                .BufPos = p491
            Else
                .BufPos = p491
            End If
        Else
            .BufPos = p491
        End If
        If ParseKEYWORD_false() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseKEYWORD_null() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseKEYWORD_promise() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseKEYWORD_true() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseKEYWORD_undefined() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseKEYWORD_unreachable() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseLabeledTypeExpr() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseSTRINGLITERAL() Then
            GoTo L234
        End If
        .BufPos = p491
        If ParseSwitchExpr() Then
            GoTo L234
        End If
        .BufPos = p491
        Exit Function
L234:
        ParsePrimaryTypeExpr = True
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
                    GoTo L238
                End If
                .BufPos = p926
            Else
                .BufPos = p926
            End If
        Else
            .BufPos = p926
        End If
        If ParseDOT() Then
            If ParseIDENTIFIER() Then
                GoTo L238
            End If
            .BufPos = p926
        Else
            .BufPos = p926
        End If
        If ParseDOTASTERISK() Then
            GoTo L238
        End If
        .BufPos = p926
        If ParseDOTQUESTIONMARK() Then
            GoTo L238
        End If
        .BufPos = p926
        Exit Function
L238:
        ParseSuffixOp = True
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
                End If
                Call pvSetExpected("CHAR_LITERAL")
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
        If Not (ParseKEYWORD_extern()) Then
            .BufPos = p536
            If Not (ParseKEYWORD_packed()) Then
                .BufPos = p536
            End If
        End If
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
        Else
            GoTo L276
        End If
        For i1110 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48, 49                             ' [01]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1110 = 0 Then
            .BufPos = p1138
            GoTo L276
        End If
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
        Else
            .BufPos = p1138
            GoTo L276
        End If
        For i1114 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48, 49                             ' [01]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1114 = 0 Then
            .BufPos = p1138
            GoTo L276
        End If
        p1118 = .BufPos
        Select Case .BufData(.BufPos)
        Case 101, 69                                ' [eE]
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 45, 43                             ' [-+]
                .BufPos = .BufPos + 1
            End Select
            For i1120 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 48, 49                         ' [01]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If Not (i1120 <> 0) Then
                .BufPos = p1118
            End If
        End Select
        Call Parseskip
        Call pvSetAdvance
        ParseFLOAT = True
        Exit Function
L276:
        If .BufData(.BufPos) = 48 And .BufData(.BufPos + 1) = 111 Then ' "0o"
            .BufPos = .BufPos + 2
        Else
            GoTo L272
        End If
        For i1125 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 55                           ' [0-7]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1125 = 0 Then
            .BufPos = p1138
            GoTo L272
        End If
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
        Else
            .BufPos = p1138
            GoTo L272
        End If
        For i1129 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 55                           ' [0-7]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1129 = 0 Then
            .BufPos = p1138
            GoTo L272
        End If
        p1133 = .BufPos
        Select Case .BufData(.BufPos)
        Case 101, 69                                ' [eE]
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 45, 43                             ' [-+]
                .BufPos = .BufPos + 1
            End Select
            For i1135 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If Not (i1135 <> 0) Then
                .BufPos = p1133
            End If
        End Select
        Call Parseskip
        Call pvSetAdvance
        ParseFLOAT = True
        Exit Function
L272:
        If .BufData(.BufPos) = 48 And .BufData(.BufPos + 1) = 120 Then ' "0x"
            .BufPos = .BufPos + 2
        Else
            GoTo L268
        End If
        For i1141 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1141 = 0 Then
            .BufPos = p1138
            GoTo L268
        End If
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
        Else
            .BufPos = p1138
            GoTo L268
        End If
        For i1145 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1145 = 0 Then
            .BufPos = p1138
            GoTo L268
        End If
        p1149 = .BufPos
        Select Case .BufData(.BufPos)
        Case 112, 80                                ' [pP]
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 45, 43                             ' [-+]
                .BufPos = .BufPos + 1
            End Select
            For i1151 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 48 To 57, 97 To 102, 65 To 70  ' [0-9a-fA-F]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If Not (i1151 <> 0) Then
                .BufPos = p1149
            End If
        End Select
        Call Parseskip
        Call pvSetAdvance
        ParseFLOAT = True
        Exit Function
L268:
        For i1155 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1155 = 0 Then
            .BufPos = p1138
            GoTo L264
        End If
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
        Else
            .BufPos = p1138
            GoTo L264
        End If
        For i1159 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1159 = 0 Then
            .BufPos = p1138
            GoTo L264
        End If
        p1163 = .BufPos
        Select Case .BufData(.BufPos)
        Case 101, 69                                ' [eE]
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 45, 43                             ' [-+]
                .BufPos = .BufPos + 1
            End Select
            For i1165 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 48 To 57                       ' [0-9]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If Not (i1165 <> 0) Then
                .BufPos = p1163
            End If
        End Select
        Call Parseskip
        Call pvSetAdvance
        ParseFLOAT = True
        Exit Function
L264:
        If .BufData(.BufPos) = 48 And .BufData(.BufPos + 1) = 98 Then ' "0b"
            .BufPos = .BufPos + 2
        Else
            GoTo L261
        End If
        For i1170 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48, 49                             ' [01]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1170 = 0 Then
            .BufPos = p1138
            GoTo L261
        End If
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
        End If
        Select Case .BufData(.BufPos)
        Case 101, 69                                ' [eE]
            .BufPos = .BufPos + 1
        Case Else
            .BufPos = p1138
            GoTo L261
        End Select
        Select Case .BufData(.BufPos)
        Case 45, 43                                 ' [-+]
            .BufPos = .BufPos + 1
        End Select
        For i1178 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48, 49                             ' [01]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1178 = 0 Then
            .BufPos = p1138
            GoTo L261
        End If
        Call Parseskip
        Call pvSetAdvance
        ParseFLOAT = True
        Exit Function
L261:
        If .BufData(.BufPos) = 48 And .BufData(.BufPos + 1) = 111 Then ' "0o"
            .BufPos = .BufPos + 2
        Else
            GoTo L258
        End If
        For i1182 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 55                           ' [0-7]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1182 = 0 Then
            .BufPos = p1138
            GoTo L258
        End If
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
        End If
        Select Case .BufData(.BufPos)
        Case 101, 69                                ' [eE]
            .BufPos = .BufPos + 1
        Case Else
            .BufPos = p1138
            GoTo L258
        End Select
        Select Case .BufData(.BufPos)
        Case 45, 43                                 ' [-+]
            .BufPos = .BufPos + 1
        End Select
        For i1190 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 55                           ' [0-7]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1190 = 0 Then
            .BufPos = p1138
            GoTo L258
        End If
        Call Parseskip
        Call pvSetAdvance
        ParseFLOAT = True
        Exit Function
L258:
        If .BufData(.BufPos) = 48 And .BufData(.BufPos + 1) = 120 Then ' "0x"
            .BufPos = .BufPos + 2
        Else
            GoTo L255
        End If
        For i1194 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1194 = 0 Then
            .BufPos = p1138
            GoTo L255
        End If
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
        End If
        Select Case .BufData(.BufPos)
        Case 112, 80                                ' [pP]
            .BufPos = .BufPos + 1
        Case Else
            .BufPos = p1138
            GoTo L255
        End Select
        Select Case .BufData(.BufPos)
        Case 45, 43                                 ' [-+]
            .BufPos = .BufPos + 1
        End Select
        For i1202 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1202 = 0 Then
            .BufPos = p1138
            GoTo L255
        End If
        Call Parseskip
        Call pvSetAdvance
        ParseFLOAT = True
        Exit Function
L255:
        For i1205 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1205 = 0 Then
            .BufPos = p1138
            GoTo L252
        End If
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
        End If
        Select Case .BufData(.BufPos)
        Case 101, 69                                ' [eE]
            .BufPos = .BufPos + 1
        Case Else
            .BufPos = p1138
            GoTo L252
        End Select
        Select Case .BufData(.BufPos)
        Case 45, 43                                 ' [-+]
            .BufPos = .BufPos + 1
        End Select
        For i1213 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1213 = 0 Then
            .BufPos = p1138
            GoTo L252
        End If
        Call Parseskip
        Call pvSetAdvance
        ParseFLOAT = True
        Exit Function
L252:
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
        Else
            GoTo L289
        End If
        For i1217 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48, 49                             ' [01]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1217 = 0 Then
            .BufPos = p1225
            GoTo L289
        End If
        Call Parseskip
        Call pvSetAdvance
        ParseINTEGER = True
        Exit Function
L289:
        If .BufData(.BufPos) = 48 And .BufData(.BufPos + 1) = 111 Then ' "0o"
            .BufPos = .BufPos + 2
        Else
            GoTo L288
        End If
        For i1222 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 55                           ' [0-7]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1222 = 0 Then
            .BufPos = p1225
            GoTo L288
        End If
        Call Parseskip
        Call pvSetAdvance
        ParseINTEGER = True
        Exit Function
L288:
        If .BufData(.BufPos) = 48 And .BufData(.BufPos + 1) = 120 Then ' "0x"
            .BufPos = .BufPos + 2
        Else
            GoTo L287
        End If
        For i1228 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1228 = 0 Then
            .BufPos = p1225
            GoTo L287
        End If
        Call Parseskip
        Call pvSetAdvance
        ParseINTEGER = True
        Exit Function
L287:
        For i1232 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i1232 = 0 Then
            .BufPos = p1225
            GoTo L286
        End If
        Call Parseskip
        Call pvSetAdvance
        ParseINTEGER = True
        Exit Function
L286:
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
                GoTo L299
            End If
            .BufPos = p568
        Else
            .BufPos = p568
        End If
        p563 = .BufPos
        If Not (ParseBlockLabel()) Then
            .BufPos = p563
        End If
        If ParseLoopTypeExpr() Then
            GoTo L299
        End If
        .BufPos = p568
        Exit Function
L299:
        ParseLabeledTypeExpr = True
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
            GoTo L308
        End If
        .BufPos = p575
        If ParseWhileTypeExpr() Then
            GoTo L308
        End If
        .BufPos = p575
        Exit Function
L308:
        ParseLoopTypeExpr = True
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
                                    GoTo L329
                                End If
                                .BufPos = p636
                            Else
                                .BufPos = p636
                            End If
                            If ParseIDENTIFIER() Then
                                GoTo L329
                            End If
                            .BufPos = p636
                            Exit Function
L329:
                            If ParseRPAREN() Then
                                ParseAsmOutputItem = True
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
        If Not (ParseKEYWORD_noalias()) Then
            .BufPos = p701
            If Not (ParseKEYWORD_comptime()) Then
                .BufPos = p701
            End If
        End If
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
            GoTo L353
        End If
        .BufPos = p713
        If ParseDOT3() Then
            GoTo L353
        End If
        .BufPos = p713
        If ParseTypeExpr() Then
            GoTo L353
        End If
        .BufPos = p713
        Exit Function
L353:
        ParseParamType = True
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
            GoTo L368
        End If
        .BufPos = p783
        If ParseKEYWORD_else() Then
            GoTo L368
        End If
        .BufPos = p783
        Exit Function
L368:
        ParseSwitchCase = True
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
            GoTo L413
        End If
        .BufPos = p950
        If ParseASTERISK2() Then
            GoTo L413
        End If
        .BufPos = p950
        If ParseLBRACKET() Then
            If ParseASTERISK() Then
                If ParseRBRACKET() Then
                    GoTo L413
                End If
                .BufPos = p950
            Else
                .BufPos = p950
            End If
        Else
            .BufPos = p950
        End If
        Exit Function
L413:
        ParsePtrTypeStart = True
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
            GoTo L425
        End If
        .BufPos = p965
        If ParseKEYWORD_enum() Then
            GoTo L425
        End If
        .BufPos = p965
        .BufPos = p988
        GoTo L422
L425:
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
        GoTo L421
L422:
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
                    GoTo L429
                End If
                .BufPos = p983
                If ParseExpr() Then
                    GoTo L429
                End If
                .BufPos = p983
                .BufPos = p984
                GoTo L428
L429:
                If Not (ParseRPAREN()) Then
                    .BufPos = p984
                End If
            Else
                .BufPos = p984
            End If
L428:
            GoTo L421
        End If
        .BufPos = p988
        Exit Function
L421:
        ParseContainerDeclType = True
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
                    GoTo L436
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
                            GoTo L436
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
                                    GoTo L436
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
                GoTo L436
            End Select
        End If
        Exit Function
L436:
        Call pvSetAdvance
        Parsechar_escape = True
    End With
End Function

Private Function Parsechar_char() As Boolean
    Dim p1073 As Long

    With ctx
        p1073 = .BufPos
        If Parsechar_escape() Then
            GoTo L441
        End If
        .BufPos = p1073
        Select Case .BufData(.BufPos)
        Case 92, 39, 13, 10                         ' [\\'\r\n]
            '--- do nothing
        Case Else
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
                GoTo L441
            End If
        End Select
        Exit Function
L441:
        Call pvSetAdvance
        Parsechar_char = True
    End With
End Function

Private Function Parsestring_char() As Boolean
    Dim p1077 As Long

    With ctx
        p1077 = .BufPos
        If Parsechar_escape() Then
            GoTo L442
        End If
        .BufPos = p1077
        Select Case .BufData(.BufPos)
        Case 92, 34, 13, 10                         ' [\\"\r\n]
            '--- do nothing
        Case Else
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
                GoTo L442
            End If
        End Select
        Exit Function
L442:
        Call pvSetAdvance
        Parsestring_char = True
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
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_and() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_anyerror() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_asm() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_async() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_await() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_break() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_cancel() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_catch() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_comptime() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_const() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_continue() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_defer() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_else() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_enum() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_errdefer() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_error() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_export() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_extern() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_false() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_fn() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_for() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_if() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_inline() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_nakedcc() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_noalias() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_null() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_or() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_orelse() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_packed() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_promise() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_pub() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_resume() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_return() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_section() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_stdcallcc() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_struct() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_suspend() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_switch() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_test() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_true() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_try() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_undefined() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_union() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_unreachable() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_use() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_var() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_volatile() Then
            GoTo L444
        End If
        .BufPos = p1633
        If ParseKEYWORD_while() Then
            GoTo L444
        End If
        .BufPos = p1633
        Exit Function
L444:
        Parsekeyword = True
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

