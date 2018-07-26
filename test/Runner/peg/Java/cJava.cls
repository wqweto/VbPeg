VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Auto-generated on 26.7.2018 19:36:21
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
    ParserVersion = "26.7.2018 19:36:21"
End Property

Property Get Contents(Optional ByVal lOffset As Long = 1, Optional ByVal lSize As Long = -1) As String
    If lSize < 0 Then
        Contents = Mid$(ctx.Contents, lOffset)
    Else
        Contents = Mid$(ctx.Contents, lOffset, lSize)
    End If
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Match(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant, Optional Result As Variant) As Long
    If BeginMatch(sSubject, StartPos, UserData) Then
        If ParseProg() Then
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
        .Contents = sSubject
        ReDim .BufData(0 To Len(sSubject) + 3) As Integer
        Call CopyMemory(.BufData(0), ByVal StrPtr(sSubject), LenB(sSubject))
        .BufPos = StartPos
        .BufSize = Len(sSubject)
        .BufData(.BufSize) = -1 '-- EOF anchor
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
        .Contents = vbNullString
        Erase .BufData
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

Private Sub pvSetError(sText As String)
    With ctx
        If LenB(.LastError) = 0 Or .BufPos > .LastBufPos Then
            .LastError = sText
            .LastBufPos = .BufPos
        ElseIf .BufPos = .LastBufPos Then
            .LastError = .LastError & "; " & sText
        End If
    End With
End Sub

'= generated functions ===================================================

Public Function ParseProg() As Boolean
    If ParsePUBLIC() Then
        If ParseCLASS() Then
            If ParseNAME() Then
                If ParseLCUR() Then
                    If ParsePUBLIC() Then
                        If ParseSTATIC() Then
                            If ParseVOID() Then
                                If ParseMAIN() Then
                                    If ParseLPAR() Then
                                        If ParseSTRING() Then
                                            If ParseLBRA() Then
                                                If ParseRBRA() Then
                                                    If ParseNAME() Then
                                                        If ParseRPAR() Then
                                                            If ParseBlockStmt() Then
                                                                If ParseRCUR() Then
                                                                    ParseProg = True
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Private Function ParsePUBLIC() As Boolean
    With ctx
        If pvMatchString("public") Then             ' "public"
            .BufPos = .BufPos + 6
            Call Parse_
            Call pvSetAdvance
            ParsePUBLIC = True
        End If
    End With
End Function

Private Function ParseCLASS() As Boolean
    With ctx
        If pvMatchString("class") Then              ' "class"
            .BufPos = .BufPos + 5
            Call Parse_
            Call pvSetAdvance
            ParseCLASS = True
        End If
    End With
End Function

Private Function ParseNAME() As Boolean
    With ctx
        Select Case .BufData(.BufPos)
        Case 97 To 122, 65 To 90, 95                ' [a-zA-Z_]
            .BufPos = .BufPos + 1
            Do
                Select Case .BufData(.BufPos)
                Case 97 To 122, 65 To 90, 95, 48 To 57 ' [a-zA-Z_0-9]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit Do
                End Select
            Loop
            Call Parse_
            Call pvSetAdvance
            ParseNAME = True
        End Select
    End With
End Function

Private Function ParseLCUR() As Boolean
    With ctx
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseLCUR = True
        End If
    End With
End Function

Private Function ParseSTATIC() As Boolean
    With ctx
        If pvMatchString("static") Then             ' "static"
            .BufPos = .BufPos + 6
            Call Parse_
            Call pvSetAdvance
            ParseSTATIC = True
        End If
    End With
End Function

Private Function ParseVOID() As Boolean
    With ctx
        If pvMatchString("void") Then               ' "void"
            .BufPos = .BufPos + 4
            Call Parse_
            Call pvSetAdvance
            ParseVOID = True
        End If
    End With
End Function

Private Function ParseMAIN() As Boolean
    With ctx
        If pvMatchString("main") Then               ' "main"
            .BufPos = .BufPos + 4
            Call Parse_
            Call pvSetAdvance
            ParseMAIN = True
        End If
    End With
End Function

Private Function ParseLPAR() As Boolean
    With ctx
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseLPAR = True
        End If
    End With
End Function

Private Function ParseSTRING() As Boolean
    With ctx
        If pvMatchString("String") Then             ' "String"
            .BufPos = .BufPos + 6
            Call Parse_
            Call pvSetAdvance
            ParseSTRING = True
        End If
    End With
End Function

Private Function ParseLBRA() As Boolean
    With ctx
        If .BufData(.BufPos) = 91 Then              ' "["
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseLBRA = True
        End If
    End With
End Function

Private Function ParseRBRA() As Boolean
    With ctx
        If .BufData(.BufPos) = 93 Then              ' "]"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseRBRA = True
        End If
    End With
End Function

Private Function ParseRPAR() As Boolean
    With ctx
        If .BufData(.BufPos) = 41 Then              ' ")"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseRPAR = True
        End If
    End With
End Function

Public Function ParseBlockStmt() As Boolean
    Dim p35 As Long
    Dim p39 As Long

    With ctx
        If ParseLCUR() Then
            Do
                p35 = .BufPos
                If Not (ParseStmt()) Then
                    .BufPos = p35
                    Exit Do
                End If
            Loop
            p39 = .BufPos
            If ParseRCUR() Then
                ParseBlockStmt = True
                Exit Function
            Else
                .BufPos = p39
            End If
            If ParseErrRcblk() Then
                ParseBlockStmt = True
                Exit Function
            Else
                .BufPos = p39
            End If
        End If
    End With
End Function

Private Function ParseRCUR() As Boolean
    With ctx
        If .BufData(.BufPos) = 125 Then             ' "}"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseRCUR = True
        End If
    End With
End Function

Public Function ParseStmt() As Boolean
    Dim p46 As Long

    With ctx
        p46 = .BufPos
        If ParseIfStmt() Then
            ParseStmt = True
            Exit Function
        Else
            .BufPos = p46
        End If
        If ParseWhileStmt() Then
            ParseStmt = True
            Exit Function
        Else
            .BufPos = p46
        End If
        If ParsePrintStmt() Then
            ParseStmt = True
            Exit Function
        Else
            .BufPos = p46
        End If
        If ParseDecStmt() Then
            ParseStmt = True
            Exit Function
        Else
            .BufPos = p46
        End If
        If ParseAssignStmt() Then
            ParseStmt = True
            Exit Function
        Else
            .BufPos = p46
        End If
        If ParseBlockStmt() Then
            ParseStmt = True
            Exit Function
        Else
            .BufPos = p46
        End If
    End With
End Function

Public Function ParseErrRcblk() As Boolean
    Dim p317 As Long
    Dim p316 As Long

    With ctx
        p317 = .BufPos
        p316 = .BufPos
        Call pvSetError("Missing end of block")
        .BufPos = p317
    End With
End Function

Public Function ParseIfStmt() As Boolean
    Dim p57 As Long
    Dim p63 As Long
    Dim p67 As Long
    Dim p71 As Long
    Dim p83 As Long
    Dim p77 As Long

    With ctx
        If ParseIF() Then
            p57 = .BufPos
            If ParseLPAR() Then
                p63 = .BufPos
                If ParseExp() Then
                    p67 = .BufPos
                    If ParseRPAR() Then
                        p71 = .BufPos
                        If ParseStmt() Then
                            p83 = .BufPos
                            If ParseELSE() Then
                                p77 = .BufPos
                                If ParseStmt() Then
                                    ParseIfStmt = True
                                    Exit Function
                                Else
                                    .BufPos = p77
                                End If
                                Call pvSetError("ErrElse")
                                .BufPos = p77
                                .BufPos = p83
                            Else
                                .BufPos = p83
                            End If
                            Call ParseEPSILON
                            ParseIfStmt = True
                            Exit Function
                        Else
                            .BufPos = p71
                        End If
                        Call pvSetError("ErrThen")
                        .BufPos = p71
                        GoTo L6
                    Else
                        .BufPos = p67
                    End If
                    Call pvSetError("ErrRpif")
                    .BufPos = p67
L6:
                    GoTo L7
                Else
                    .BufPos = p63
                End If
                Call pvSetError("ErrCondi")
                .BufPos = p63
L7:
                GoTo L5
            Else
                .BufPos = p57
            End If
            Call pvSetError("ErrLpif")
            .BufPos = p57
L5:
        End If
    End With
End Function

Public Function ParseWhileStmt() As Boolean
    Dim p87 As Long
    Dim p92 As Long
    Dim p96 As Long
    Dim p100 As Long

    With ctx
        If ParseWHILE() Then
            p87 = .BufPos
            If ParseLPAR() Then
                p92 = .BufPos
                If ParseExp() Then
                    p96 = .BufPos
                    If ParseRPAR() Then
                        p100 = .BufPos
                        If ParseStmt() Then
                            ParseWhileStmt = True
                            Exit Function
                        Else
                            .BufPos = p100
                        End If
                        Call pvSetError("ErrBody")
                        .BufPos = p100
                        GoTo L15
                    Else
                        .BufPos = p96
                    End If
                    Call pvSetError("ErrRpw")
                    .BufPos = p96
L15:
                    GoTo L16
                Else
                    .BufPos = p92
                End If
                Call pvSetError("ErrCondw")
                .BufPos = p92
L16:
                GoTo L14
            Else
                .BufPos = p87
            End If
            Call pvSetError("ErrLpw")
            .BufPos = p87
L14:
        End If
    End With
End Function

Public Function ParsePrintStmt() As Boolean
    Dim p141 As Long
    Dim p146 As Long
    Dim p150 As Long
    Dim p154 As Long

    With ctx
        If ParsePRINTLN() Then
            p141 = .BufPos
            If ParseLPAR() Then
                p146 = .BufPos
                If ParseExp() Then
                    p150 = .BufPos
                    If ParseRPAR() Then
                        p154 = .BufPos
                        If ParseSEMI() Then
                            ParsePrintStmt = True
                            Exit Function
                        Else
                            .BufPos = p154
                        End If
                        Call pvSetError("ErrSemip")
                        .BufPos = p154
                        GoTo L24
                    Else
                        .BufPos = p150
                    End If
                    Call pvSetError("ErrRpp")
                    .BufPos = p150
L24:
                    GoTo L25
                Else
                    .BufPos = p146
                End If
                Call pvSetError("ErrEprint")
                .BufPos = p146
L25:
                GoTo L23
            Else
                .BufPos = p141
            End If
            Call pvSetError("ErrLpp")
            .BufPos = p141
L23:
        End If
    End With
End Function

Public Function ParseDecStmt() As Boolean
    Dim p106 As Long
    Dim p118 As Long
    Dim p113 As Long
    Dim p121 As Long

    With ctx
        If ParseINT() Then
            p106 = .BufPos
            If ParseNAME() Then
                p118 = .BufPos
                If ParseASSIGN() Then
                    p113 = .BufPos
                    If ParseExp() Then
                        p121 = .BufPos
                        If ParseSEMI() Then
                            ParseDecStmt = True
                            Exit Function
                        Else
                            .BufPos = p121
                        End If
                        Call pvSetError("ErrSemid")
                        .BufPos = p121
                    Else
                        .BufPos = p113
                    End If
                    Call pvSetError("ErrEdec")
                    .BufPos = p113
                    .BufPos = p118
                Else
                    .BufPos = p118
                End If
                Call ParseEPSILON
                p121 = .BufPos
                If ParseSEMI() Then
                    ParseDecStmt = True
                    Exit Function
                Else
                    .BufPos = p121
                End If
                Call pvSetError("ErrSemid")
                .BufPos = p121
                GoTo L29
            Else
                .BufPos = p106
            End If
            Call pvSetError("ErrNdec")
            .BufPos = p106
L29:
        End If
    End With
End Function

Public Function ParseAssignStmt() As Boolean
    Dim p126 As Long
    Dim p131 As Long
    Dim p135 As Long

    With ctx
        If ParseNAME() Then
            p126 = .BufPos
            If ParseASSIGN() Then
                p131 = .BufPos
                If ParseExp() Then
                    p135 = .BufPos
                    If ParseSEMI() Then
                        ParseAssignStmt = True
                        Exit Function
                    Else
                        .BufPos = p135
                    End If
                    If ParseErrSemia() Then
                        ParseAssignStmt = True
                        Exit Function
                    Else
                        .BufPos = p135
                    End If
                    GoTo L34
                Else
                    .BufPos = p131
                End If
                Call pvSetError("ErrRval")
                .BufPos = p131
L34:
                GoTo L33
            Else
                .BufPos = p126
            End If
            Call pvSetError("ErrAssign")
            .BufPos = p126
L33:
        End If
    End With
End Function

Private Function ParseIF() As Boolean
    With ctx
        If .BufData(.BufPos) = 105 And .BufData(.BufPos + 1) = 102 Then ' "if"
            .BufPos = .BufPos + 2
            Call Parse_
            Call pvSetAdvance
            ParseIF = True
        End If
    End With
End Function

Public Function ParseExp() As Boolean
    Dim p165 As Long
    Dim p162 As Long

    With ctx
        If ParseRelExp() Then
            Do
                p165 = .BufPos
                If Not (ParseEQ()) Then
                    .BufPos = p165
                    Exit Do
                End If
                p162 = .BufPos
                If Not (ParseRelExp()) Then
                    .BufPos = p162
                    Call pvSetError("ErrRelexp")
                    .BufPos = p162
                    .BufPos = p165
                    Exit Do
                End If
            Loop
            ParseExp = True
        End If
    End With
End Function

Private Function ParseELSE() As Boolean
    With ctx
        If pvMatchString("else") Then               ' "else"
            .BufPos = .BufPos + 4
            Call Parse_
            Call pvSetAdvance
            ParseELSE = True
        End If
    End With
End Function

Private Sub ParseEPSILON()
    Call Parse_
L37:
End Sub

Private Function ParseWHILE() As Boolean
    With ctx
        If pvMatchString("while") Then              ' "while"
            .BufPos = .BufPos + 5
            Call Parse_
            Call pvSetAdvance
            ParseWHILE = True
        End If
    End With
End Function

Private Function ParseINT() As Boolean
    With ctx
        If pvMatchString("int") Then                ' "int"
            .BufPos = .BufPos + 3
            Call Parse_
            Call pvSetAdvance
            ParseINT = True
        End If
    End With
End Function

Private Function ParseASSIGN() As Boolean
    With ctx
        If .BufData(.BufPos) = 61 Then              ' "="
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseASSIGN = True
        End If
    End With
End Function

Private Function ParseSEMI() As Boolean
    With ctx
        If .BufData(.BufPos) = 59 Then              ' ";"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseSEMI = True
        End If
    End With
End Function

Public Function ParseErrSemia() As Boolean
    Dim p313 As Long
    Dim p312 As Long

    With ctx
        p313 = .BufPos
        p312 = .BufPos
        Call pvSetError("Missing semicolon in assignment")
        .BufPos = p313
    End With
End Function

Private Function ParsePRINTLN() As Boolean
    With ctx
        If pvMatchString("System.out.println") Then ' "System.out.println"
            .BufPos = .BufPos + 18
            Call Parse_
            Call pvSetAdvance
            ParsePRINTLN = True
        End If
    End With
End Function

Public Function ParseRelExp() As Boolean
    Dim p176 As Long
    Dim p173 As Long

    With ctx
        If ParseAddExp() Then
            Do
                p176 = .BufPos
                If Not (ParseLT()) Then
                    .BufPos = p176
                    Exit Do
                End If
                p173 = .BufPos
                If Not (ParseAddExp()) Then
                    .BufPos = p173
                    Call pvSetError("ErrAddexp")
                    .BufPos = p173
                    .BufPos = p176
                    Exit Do
                End If
            Loop
            ParseRelExp = True
        End If
    End With
End Function

Private Function ParseEQ() As Boolean
    With ctx
        If .BufData(.BufPos) = 61 Then              ' "="
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseEQ = True
        End If
    End With
End Function

Public Function ParseAddExp() As Boolean
    Dim p190 As Long
    Dim p185 As Long
    Dim p187 As Long

    With ctx
        If ParseMulExp() Then
            Do
                p190 = .BufPos
                p185 = .BufPos
                If Not (ParsePLUS()) Then
                    .BufPos = p185
                    If Not (ParseMINUS()) Then
                        .BufPos = p185
                        .BufPos = p190
                        Exit Do
                    End If
                End If
                p187 = .BufPos
                If Not (ParseMulExp()) Then
                    .BufPos = p187
                    Call pvSetError("ErrMulexp")
                    .BufPos = p187
                    .BufPos = p190
                    Exit Do
                End If
            Loop
            ParseAddExp = True
        End If
    End With
End Function

Private Function ParseLT() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseLT = True
        End If
    End With
End Function

Public Function ParseMulExp() As Boolean
    Dim p204 As Long
    Dim p199 As Long
    Dim p201 As Long

    With ctx
        If ParseAtomExp() Then
            Do
                p204 = .BufPos
                p199 = .BufPos
                If Not (ParseTIMES()) Then
                    .BufPos = p199
                    If Not (ParseDIV()) Then
                        .BufPos = p199
                        .BufPos = p204
                        Exit Do
                    End If
                End If
                p201 = .BufPos
                If Not (ParseAtomExp()) Then
                    .BufPos = p201
                    Call pvSetError("ErrAtomexp")
                    .BufPos = p201
                    .BufPos = p204
                    Exit Do
                End If
            Loop
            ParseMulExp = True
        End If
    End With
End Function

Private Function ParsePLUS() As Boolean
    With ctx
        If .BufData(.BufPos) = 43 Then              ' "+"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParsePLUS = True
        End If
    End With
End Function

Private Function ParseMINUS() As Boolean
    With ctx
        If .BufData(.BufPos) = 45 Then              ' "-"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseMINUS = True
        End If
    End With
End Function

Public Function ParseAtomExp() As Boolean
    Dim p219 As Long
    Dim p209 As Long
    Dim p214 As Long

    With ctx
        p219 = .BufPos
        If ParseLPAR() Then
            p209 = .BufPos
            If ParseExp() Then
                p214 = .BufPos
                If ParseRPAR() Then
                    ParseAtomExp = True
                    Exit Function
                Else
                    .BufPos = p214
                End If
                Call pvSetError("ErrRpe")
                .BufPos = p214
                .BufPos = p219
                GoTo L43
            Else
                .BufPos = p209
            End If
            Call pvSetError("ErrParexp")
            .BufPos = p209
            .BufPos = p219
L43:
        Else
            .BufPos = p219
        End If
        If ParseNUMBER() Then
            ParseAtomExp = True
            Exit Function
        Else
            .BufPos = p219
        End If
        If ParseNAME() Then
            ParseAtomExp = True
            Exit Function
        Else
            .BufPos = p219
        End If
    End With
End Function

Private Function ParseTIMES() As Boolean
    With ctx
        If .BufData(.BufPos) = 42 Then              ' "*"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseTIMES = True
        End If
    End With
End Function

Private Function ParseDIV() As Boolean
    With ctx
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseDIV = True
        End If
    End With
End Function

Private Function ParseNUMBER() As Boolean
    Dim i306 As Long

    With ctx
        For i306 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i306 <> 0 Then
            Call Parse_
            Call pvSetAdvance
            ParseNUMBER = True
        End If
    End With
End Function

Private Sub Parse_()
    With ctx
        Do
            Select Case .BufData(.BufPos)
            Case 32, 9, 13, 10                      ' [ \t\r\n]
                .BufPos = .BufPos + 1
            Case Else
                Exit Do
            End Select
        Loop
        Call pvSetAdvance
    End With
End Sub

