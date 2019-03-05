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
' Auto-generated on 5.3.2019 16:44:40
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
    ParserVersion = "5.3.2019 16:44:40"
End Property

Property Get Contents(Optional ByVal lOffset As Long = 1, Optional ByVal lSize As Long = LNG_MAXINT) As String
    Contents = Mid$(ctx.Contents, lOffset, lSize)
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

Private Function ParseBlockStmt() As Boolean
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
                GoTo L15
            End If
            .BufPos = p39
            If ParseErrRcblk() Then
                GoTo L15
            End If
            .BufPos = p39
            Exit Function
L15:
            ParseBlockStmt = True
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

Private Function ParseStmt() As Boolean
    Dim p46 As Long

    With ctx
        p46 = .BufPos
        If ParseIfStmt() Then
            GoTo L17
        End If
        .BufPos = p46
        If ParseWhileStmt() Then
            GoTo L17
        End If
        .BufPos = p46
        If ParsePrintStmt() Then
            GoTo L17
        End If
        .BufPos = p46
        If ParseDecStmt() Then
            GoTo L17
        End If
        .BufPos = p46
        If ParseAssignStmt() Then
            GoTo L17
        End If
        .BufPos = p46
        If ParseBlockStmt() Then
            GoTo L17
        End If
        .BufPos = p46
        Exit Function
L17:
        ParseStmt = True
    End With
End Function

Private Function ParseErrRcblk() As Boolean
    Dim p317 As Long
    Dim p316 As Long

    With ctx
        p317 = .BufPos
        p316 = .BufPos
        Call pvSetError("Missing end of block")
        .BufPos = p317
        Exit Function
        ParseErrRcblk = True
    End With
End Function

Private Function ParseIfStmt() As Boolean
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
                GoTo L26
            End If
            .BufPos = p57
            Call pvSetError("ErrLpif")
            .BufPos = p57
            Exit Function
L26:
            p63 = .BufPos
            If ParseExp() Then
                GoTo L25
            End If
            .BufPos = p63
            Call pvSetError("ErrCondi")
            .BufPos = p63
            Exit Function
L25:
            p67 = .BufPos
            If ParseRPAR() Then
                GoTo L24
            End If
            .BufPos = p67
            Call pvSetError("ErrRpif")
            .BufPos = p67
            Exit Function
L24:
            p71 = .BufPos
            If ParseStmt() Then
                GoTo L23
            End If
            .BufPos = p71
            Call pvSetError("ErrThen")
            .BufPos = p71
            Exit Function
L23:
            p83 = .BufPos
            If ParseELSE() Then
                p77 = .BufPos
                If ParseStmt() Then
                    GoTo L20
                End If
                .BufPos = p77
                Call pvSetError("ErrElse")
                .BufPos = p83
                GoTo L21
            End If
            .BufPos = p83
L21:
            Call ParseEPSILON
            GoTo L20
            Exit Function
L20:
            ParseIfStmt = True
        End If
    End With
End Function

Private Function ParseWhileStmt() As Boolean
    Dim p87 As Long
    Dim p92 As Long
    Dim p96 As Long
    Dim p100 As Long

    With ctx
        If ParseWHILE() Then
            p87 = .BufPos
            If ParseLPAR() Then
                GoTo L31
            End If
            .BufPos = p87
            Call pvSetError("ErrLpw")
            .BufPos = p87
            Exit Function
L31:
            p92 = .BufPos
            If ParseExp() Then
                GoTo L30
            End If
            .BufPos = p92
            Call pvSetError("ErrCondw")
            .BufPos = p92
            Exit Function
L30:
            p96 = .BufPos
            If ParseRPAR() Then
                GoTo L29
            End If
            .BufPos = p96
            Call pvSetError("ErrRpw")
            .BufPos = p96
            Exit Function
L29:
            p100 = .BufPos
            If ParseStmt() Then
                GoTo L28
            End If
            .BufPos = p100
            Call pvSetError("ErrBody")
            .BufPos = p100
            Exit Function
L28:
            ParseWhileStmt = True
        End If
    End With
End Function

Private Function ParsePrintStmt() As Boolean
    Dim p141 As Long
    Dim p146 As Long
    Dim p150 As Long
    Dim p154 As Long

    With ctx
        If ParsePRINTLN() Then
            p141 = .BufPos
            If ParseLPAR() Then
                GoTo L36
            End If
            .BufPos = p141
            Call pvSetError("ErrLpp")
            .BufPos = p141
            Exit Function
L36:
            p146 = .BufPos
            If ParseExp() Then
                GoTo L35
            End If
            .BufPos = p146
            Call pvSetError("ErrEprint")
            .BufPos = p146
            Exit Function
L35:
            p150 = .BufPos
            If ParseRPAR() Then
                GoTo L34
            End If
            .BufPos = p150
            Call pvSetError("ErrRpp")
            .BufPos = p150
            Exit Function
L34:
            p154 = .BufPos
            If ParseSEMI() Then
                GoTo L33
            End If
            .BufPos = p154
            Call pvSetError("ErrSemip")
            .BufPos = p154
            Exit Function
L33:
            ParsePrintStmt = True
        End If
    End With
End Function

Private Function ParseDecStmt() As Boolean
    Dim p106 As Long
    Dim p118 As Long
    Dim p113 As Long
    Dim p121 As Long

    With ctx
        If ParseINT() Then
            p106 = .BufPos
            If ParseNAME() Then
                GoTo L42
            End If
            .BufPos = p106
            Call pvSetError("ErrNdec")
            .BufPos = p106
            Exit Function
L42:
            p118 = .BufPos
            If ParseASSIGN() Then
                p113 = .BufPos
                If ParseExp() Then
                    GoTo L39
                End If
                .BufPos = p113
                Call pvSetError("ErrEdec")
                .BufPos = p118
                GoTo L40
            End If
            .BufPos = p118
L40:
            Call ParseEPSILON
            GoTo L39
            Exit Function
L39:
            p121 = .BufPos
            If ParseSEMI() Then
                GoTo L38
            End If
            .BufPos = p121
            Call pvSetError("ErrSemid")
            .BufPos = p121
            Exit Function
L38:
            ParseDecStmt = True
        End If
    End With
End Function

Private Function ParseAssignStmt() As Boolean
    Dim p126 As Long
    Dim p131 As Long
    Dim p135 As Long

    With ctx
        If ParseNAME() Then
            p126 = .BufPos
            If ParseASSIGN() Then
                GoTo L46
            End If
            .BufPos = p126
            Call pvSetError("ErrAssign")
            .BufPos = p126
            Exit Function
L46:
            p131 = .BufPos
            If ParseExp() Then
                GoTo L45
            End If
            .BufPos = p131
            Call pvSetError("ErrRval")
            .BufPos = p131
            Exit Function
L45:
            p135 = .BufPos
            If ParseSEMI() Then
                GoTo L44
            End If
            .BufPos = p135
            If ParseErrSemia() Then
                GoTo L44
            End If
            .BufPos = p135
            Exit Function
L44:
            ParseAssignStmt = True
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

Private Function ParseExp() As Boolean
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

Private Function ParseErrSemia() As Boolean
    Dim p313 As Long
    Dim p312 As Long

    With ctx
        p313 = .BufPos
        p312 = .BufPos
        Call pvSetError("Missing semicolon in assignment")
        .BufPos = p313
        Exit Function
        ParseErrSemia = True
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

Private Function ParseRelExp() As Boolean
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

Private Function ParseAddExp() As Boolean
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
                        .BufPos = p190
                        Exit Do
                    End If
                End If
                p187 = .BufPos
                If Not (ParseMulExp()) Then
                    .BufPos = p187
                    Call pvSetError("ErrMulexp")
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

Private Function ParseMulExp() As Boolean
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
                        .BufPos = p204
                        Exit Do
                    End If
                End If
                p201 = .BufPos
                If Not (ParseAtomExp()) Then
                    .BufPos = p201
                    Call pvSetError("ErrAtomexp")
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

Private Function ParseAtomExp() As Boolean
    Dim p219 As Long
    Dim p209 As Long
    Dim p214 As Long

    With ctx
        p219 = .BufPos
        If ParseLPAR() Then
            p209 = .BufPos
            If ParseExp() Then
                GoTo L73
            End If
            .BufPos = p209
            Call pvSetError("ErrParexp")
            .BufPos = p219
            GoTo L71
L73:
            p214 = .BufPos
            If ParseRPAR() Then
                GoTo L70
            End If
            .BufPos = p214
            Call pvSetError("ErrRpe")
            .BufPos = p219
            GoTo L71
        End If
        .BufPos = p219
L71:
        If ParseNUMBER() Then
            GoTo L70
        End If
        .BufPos = p219
        If ParseNAME() Then
            GoTo L70
        End If
        .BufPos = p219
        Exit Function
L70:
        ParseAtomExp = True
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

