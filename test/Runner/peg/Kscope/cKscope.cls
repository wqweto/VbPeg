VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cKscope"
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

Private Enum UcsParserActionsEnum
    ucsAct_3_start
    ucsAct_2_start
    ucsAct_1_start
    ucsAct_1_definition
    ucsAct_1_extern
    ucsAct_1_prototype
    ucsAct_2_prototype
    ucsAct_1_binopproto
    ucsAct_1_UNOP
    ucsAct_3_param_list
    ucsAct_2_param_list
    ucsAct_1_param_list
    ucsAct_1_IDENT
    ucsAct_1_BINOP
    ucsAct_1_NUMBER
    ucsAct_2_binexpr
    ucsAct_1_binexpr
    ucsAct_1_unexpr
    ucsAct_1_DOUBLE
    ucsAct_1_invokeexpr
    ucsAct_1_ifexpr
    ucsAct_1_forexpr
    ucsAct_1_varexpr
    ucsAct_3_arg_list
    ucsAct_2_arg_list
    ucsAct_1_arg_list
    ucsAct_3_var_list
    ucsAct_2_var_list
    ucsAct_1_var_list
    ucsAct_1_var_decl
    ucsActVarAlloc = -1
    ucsActVarSet = -2
    ucsActResultClear = -3
    ucsActResultSet = -4
End Enum

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
    VarStack()          As Variant
    VarPos              As Long
End Type

Private m_oBinopInfo        As Scripting.Dictionary
Private m_lBinopMinPrec     As Long

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
        If Parsestart() Then
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
    Dim lIdx            As Long

    With ctx
        ReDim .VarStack(0 To 1024) As Variant
        For lIdx = 0 To .ThunkPos - 1
            Select Case .ThunkData(lIdx).Action
            Case ucsActVarAlloc
                .VarPos = .VarPos + .ThunkData(lIdx).CaptureBegin
            Case ucsActVarSet
                If IsObject(.VarResult) Then
                    Set .VarStack(.VarPos - .ThunkData(lIdx).CaptureBegin) = .VarResult
                Else
                    .VarStack(.VarPos - .ThunkData(lIdx).CaptureBegin) = .VarResult
                End If
            Case ucsActResultClear
                .VarResult = Empty
            Case ucsActResultSet
                With .ThunkData(lIdx)
                    ctx.VarResult = Mid$(ctx.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)
                End With
            Case Else
                With .ThunkData(lIdx)
                    pvImplAction .Action, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin
                End With
            End Select
        Next
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

Private Sub pvPushThunk(ByVal eAction As UcsParserActionsEnum, Optional ByVal lBegin As Long, Optional ByVal lEnd As Long)
    With ctx
        If UBound(.ThunkData) < .ThunkPos Then
            ReDim Preserve .ThunkData(0 To 2 * UBound(.ThunkData)) As UcsParserThunkType
        End If
        With .ThunkData(.ThunkPos)
            .Action = eAction
            .CaptureBegin = lBegin
            .CaptureEnd = lEnd
        End With
        .ThunkPos = .ThunkPos + 1
    End With
End Sub

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

'= generated functions ===================================================

Public Function Parsestart() As Boolean
    Dim i14 As Long
    Dim p12 As Long
    Dim q12 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        Call Parse_
        pvPushThunk ucsActVarSet, 1
        pvPushThunk ucsAct_1_start, .CaptureBegin, .CaptureEnd
        For i14 = 0 To LNG_MAXINT
            p12 = .BufPos
            q12 = .ThunkPos
            pvPushThunk ucsActResultClear
            If Parsestmt() Then
                pvPushThunk ucsActVarSet, 2
            Else
                .BufPos = p12
                .ThunkPos = q12
                Exit For
            End If
            pvPushThunk ucsAct_2_start, .CaptureBegin, .CaptureEnd
        Next
        If i14 <> 0 Then
            pvPushThunk ucsAct_3_start, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -2
            Parsestart = True
        End If
    End With
End Function

Private Sub Parse_()
    Dim p417 As Long

    With ctx
        Do
            p417 = .BufPos
            If Not (ParseLineComment()) Then
                .BufPos = p417
                Select Case .BufData(.BufPos)
                Case 32, 9, 13, 10                  ' [ \t\r\n]
                    .BufPos = .BufPos + 1
                Case Else
                    .BufPos = p417
                    Exit Do
                End Select
            End If
        Loop
        Call pvSetAdvance
    End With
End Sub

Private Function Parsestmt() As Boolean
    Dim p27 As Long
    Dim q27 As Long
    Dim p19 As Long
    Dim q19 As Long
    Dim p24 As Long
    Dim q24 As Long
    Dim p30 As Long
    Dim q30 As Long

    With ctx
        p27 = .BufPos
        q27 = .ThunkPos
        If Parsedefinition() Then
            p19 = .BufPos
            q19 = .ThunkPos
            If Not (ParseSEMI()) Then
                .BufPos = p19
                .ThunkPos = q19
            End If
            GoTo L3
        End If
        .BufPos = p27
        .ThunkPos = q27
        If Parseextern() Then
            p24 = .BufPos
            q24 = .ThunkPos
            If Not (ParseSEMI()) Then
                .BufPos = p24
                .ThunkPos = q24
            End If
            GoTo L3
        End If
        .BufPos = p27
        .ThunkPos = q27
        If Parseexpr() Then
            p30 = .BufPos
            q30 = .ThunkPos
            If Not (ParseSEMI()) Then
                .BufPos = p30
                .ThunkPos = q30
            End If
            GoTo L3
        End If
        .BufPos = p27
        .ThunkPos = q27
        If ParseSEMI() Then
            GoTo L3
        End If
        .BufPos = p27
        .ThunkPos = q27
        Exit Function
L3:
        Parsestmt = True
    End With
End Function

Private Function Parsedefinition() As Boolean
    With ctx
        pvPushThunk ucsActVarAlloc, 2
        If pvMatchString("def") Then                ' "def"
            .BufPos = .BufPos + 3
            If Parsens() Then
                pvPushThunk ucsActResultClear
                If Parseprototype() Then
                    pvPushThunk ucsActVarSet, 1
                    pvPushThunk ucsActResultClear
                    If Parseexpr() Then
                        pvPushThunk ucsActVarSet, 2
                        pvPushThunk ucsAct_1_definition, .CaptureBegin, .CaptureEnd
                        pvPushThunk ucsActVarAlloc, -2
                        Call pvSetAdvance
                        Parsedefinition = True
                    End If
                End If
            End If
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

Private Function Parseextern() As Boolean
    With ctx
        pvPushThunk ucsActVarAlloc, 1
        If pvMatchString("extern") Then             ' "extern"
            .BufPos = .BufPos + 6
            If Parsens() Then
                pvPushThunk ucsActResultClear
                If Parseprototype() Then
                    pvPushThunk ucsActVarSet, 1
                    pvPushThunk ucsAct_1_extern, .CaptureBegin, .CaptureEnd
                    pvPushThunk ucsActVarAlloc, -1
                    Call pvSetAdvance
                    Parseextern = True
                End If
            End If
        End If
    End With
End Function

Private Function Parseexpr() As Boolean
    With ctx
        '--- clear min precedence
        m_lBinopMinPrec = 0
        If Parsebinexpr() Then
            Parseexpr = True
        End If
    End With
End Function

Private Function Parsens() As Boolean
    Dim p420 As Long

    With ctx
        p420 = .BufPos
        Select Case .BufData(.BufPos)
        Case 97 To 122, 65 To 90, 95, 48 To 57      ' [a-zA-Z_0-9]
            '--- do nothing
        Case Else
            .BufPos = p420
            Call Parse_
            Call pvSetAdvance
            Parsens = True
        End Select
    End With
End Function

Private Function Parseprototype() As Boolean
    Dim p79 As Long
    Dim q79 As Long
    Dim p82 As Long
    Dim q82 As Long
    Dim e82 As String

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        p79 = .BufPos
        q79 = .ThunkPos
        If Parsebinopproto() Then
            GoTo L15
        End If
        .BufPos = p79
        .ThunkPos = q79
        If pvMatchString("unary") Then              ' "unary"
            .BufPos = .BufPos + 5
            If Parsens() Then
                pvPushThunk ucsActResultClear
                If ParseUNOP() Then
                    pvPushThunk ucsActVarSet, 1
                    If ParseLPAREN() Then
                        pvPushThunk ucsActResultClear
                        If Parseparam_list() Then
                            pvPushThunk ucsActVarSet, 2
                            If ParseRPAREN() Then
                                pvPushThunk ucsAct_1_prototype, .CaptureBegin, .CaptureEnd
                                GoTo L15
                            End If
                            .BufPos = p79
                            .ThunkPos = q79
                        Else
                            .BufPos = p79
                            .ThunkPos = q79
                        End If
                    Else
                        .BufPos = p79
                        .ThunkPos = q79
                    End If
                Else
                    .BufPos = p79
                    .ThunkPos = q79
                End If
            Else
                .BufPos = p79
                .ThunkPos = q79
            End If
        End If
        p82 = .BufPos
        q82 = .ThunkPos
        e82 = .LastExpected
        If ParseKEYWORD() Then
            .BufPos = p79
            .ThunkPos = q79
        Else
            .BufPos = p82
            .ThunkPos = q82
            .LastExpected = e82
            pvPushThunk ucsActResultClear
            If ParseIDENT() Then
                pvPushThunk ucsActVarSet, 1
                If ParseLPAREN() Then
                    pvPushThunk ucsActResultClear
                    If Parseparam_list() Then
                        pvPushThunk ucsActVarSet, 2
                        If ParseRPAREN() Then
                            pvPushThunk ucsAct_2_prototype, .CaptureBegin, .CaptureEnd
                            GoTo L15
                        End If
                        .BufPos = p79
                        .ThunkPos = q79
                    Else
                        .BufPos = p79
                        .ThunkPos = q79
                    End If
                Else
                    .BufPos = p79
                    .ThunkPos = q79
                End If
            Else
                .BufPos = p79
                .ThunkPos = q79
            End If
        End If
        Exit Function
L15:
        pvPushThunk ucsActVarAlloc, -2
        Call pvSetAdvance
        Parseprototype = True
    End With
End Function

Private Function Parsebinopproto() As Boolean
    With ctx
        pvPushThunk ucsActVarAlloc, 3
        Dim sBinOp          As String
        Dim lOpPrec         As Long
        Dim lOpAssoc        As Long
        If pvMatchString("binary") Then             ' "binary"
            .BufPos = .BufPos + 6
            If Parsens() Then
                pvPushThunk ucsActResultClear
                If ParseBINOP() Then
                    pvPushThunk ucsActVarSet, 1
                    '-- collect local binop info for use on success
                    sBinOp = Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)
                    pvPushThunk ucsActResultClear
                    If ParseNUMBER() Then
                        pvPushThunk ucsActVarSet, 2
                        lOpPrec = C_Lng(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin))
                        lOpAssoc = 1
                        If ParseLPAREN() Then
                            pvPushThunk ucsActResultClear
                            If Parseparam_list() Then
                                pvPushThunk ucsActVarSet, 3
                                If ParseRPAREN() Then
                                    '--- on match immediately update parser binop info
                                    SetBinopInfo sBinOp, lOpPrec, lOpAssoc
                                    pvPushThunk ucsAct_1_binopproto, .CaptureBegin, .CaptureEnd
                                    pvPushThunk ucsActVarAlloc, -3
                                    Call pvSetAdvance
                                    Parsebinopproto = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseUNOP() As Boolean
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        Select Case .BufData(.BufPos)
        Case 45, 43, 33, 126                        ' [-+!~]
            .BufPos = .BufPos + 1
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_UNOP, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseUNOP = True
        End Select
    End With
End Function

Private Function ParseLPAREN() As Boolean
    With ctx
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseLPAREN = True
        End If
    End With
End Function

Private Function Parseparam_list() As Boolean
    Dim p271 As Long
    Dim q271 As Long
    Dim p265 As Long
    Dim q265 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        If ParseIDENT() Then
            pvPushThunk ucsActVarSet, 1
            pvPushThunk ucsAct_1_param_list, .CaptureBegin, .CaptureEnd
            Do
                p271 = .BufPos
                q271 = .ThunkPos
                p265 = .BufPos
                q265 = .ThunkPos
                If Not (ParseCOMMA()) Then
                    .BufPos = p265
                    .ThunkPos = q265
                End If
                pvPushThunk ucsActResultClear
                If ParseIDENT() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p271
                    .ThunkPos = q271
                    Exit Do
                End If
                pvPushThunk ucsAct_2_param_list, .CaptureBegin, .CaptureEnd
            Loop
            pvPushThunk ucsAct_3_param_list, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -2
            Parseparam_list = True
        End If
    End With
End Function

Private Function ParseRPAREN() As Boolean
    With ctx
        If .BufData(.BufPos) = 41 Then              ' ")"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseRPAREN = True
        End If
    End With
End Function

Private Function ParseKEYWORD() As Boolean
    Dim p335 As Long

    With ctx
        If pvMatchString("def") Then                ' "def"
            .BufPos = .BufPos + 3
            GoTo L25
        End If
        If pvMatchString("extern") Then             ' "extern"
            .BufPos = .BufPos + 6
            GoTo L25
        End If
        If .BufData(.BufPos) = 105 And .BufData(.BufPos + 1) = 102 Then ' "if"
            .BufPos = .BufPos + 2
            GoTo L25
        End If
        If pvMatchString("then") Then               ' "then"
            .BufPos = .BufPos + 4
            GoTo L25
        End If
        If pvMatchString("else") Then               ' "else"
            .BufPos = .BufPos + 4
            GoTo L25
        End If
        If pvMatchString("unary") Then              ' "unary"
            .BufPos = .BufPos + 5
            GoTo L25
        End If
        If pvMatchString("for") Then                ' "for"
            .BufPos = .BufPos + 3
            GoTo L25
        End If
        If pvMatchString("binary") Then             ' "binary"
            .BufPos = .BufPos + 6
            GoTo L25
        End If
        If .BufData(.BufPos) = 105 And .BufData(.BufPos + 1) = 110 Then ' "in"
            .BufPos = .BufPos + 2
            GoTo L25
        End If
        If pvMatchString("var") Then                ' "var"
            .BufPos = .BufPos + 3
            GoTo L25
        End If
        Exit Function
L25:
        p335 = .BufPos
        Select Case .BufData(.BufPos)
        Case 97 To 122, 65 To 90, 95, 48 To 57      ' [a-zA-Z_0-9]
            '--- do nothing
        Case Else
            .BufPos = p335
            Call pvSetAdvance
            ParseKEYWORD = True
        End Select
    End With
End Function

Private Function ParseIDENT() As Boolean
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
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
            lCaptureEnd = .BufPos
            Call Parse_
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_IDENT, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseIDENT = True
        End Select
    End With
End Function

Private Function ParseBINOP() As Boolean
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        Select Case .BufData(.BufPos)
        Case 45, 43, 42, 47, 60, 62, 61, 58, 124, 38 ' [-+*/<>=:|&]
            .BufPos = .BufPos + 1
            If .BufData(.BufPos) = 61 Then          ' "="
                .BufPos = .BufPos + 1
            End If
            lCaptureEnd = .BufPos
            Call Parse_
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_BINOP, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseBINOP = True
        End Select
    End With
End Function

Private Function ParseNUMBER() As Boolean
    Dim lCaptureBegin As Long
    Dim i375 As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        For i375 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i375 <> 0 Then
            lCaptureEnd = .BufPos
            Call Parse_
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_NUMBER, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseNUMBER = True
        End If
    End With
End Function

Private Function Parsebinexpr() As Boolean
    Dim p133 As Long
    Dim q133 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 3
        Dim lMinPrec        As Long
        Dim vOpInfo         As Variant
        lMinPrec = m_lBinopMinPrec
        pvPushThunk ucsActResultClear
        If Parseunexpr() Then
            pvPushThunk ucsActVarSet, 1
            Do
                p133 = .BufPos
                q133 = .ThunkPos
                pvPushThunk ucsActResultClear
                If ParseBINOP() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p133
                    .ThunkPos = q133
                    Exit Do
                End If
                '--- get current binop prec & assoc
                vOpInfo = GetBinopInfo(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin))
                If Not (vOpInfo(0) >= lMinPrec) Then
                    .BufPos = p133
                    .ThunkPos = q133
                    Exit Do
                End If
                '--- set next min precedence based on binop prec & assoc
                m_lBinopMinPrec = vOpInfo(0) + vOpInfo(1)
                pvPushThunk ucsActResultClear
                If Parsebinexpr() Then
                    pvPushThunk ucsActVarSet, 3
                Else
                    .BufPos = p133
                    .ThunkPos = q133
                    Exit Do
                End If
                pvPushThunk ucsAct_1_binexpr, .CaptureBegin, .CaptureEnd
            Loop
            pvPushThunk ucsAct_2_binexpr, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -3
            Parsebinexpr = True
        End If
    End With
End Function

Private Function Parseunexpr() As Boolean
    Dim p155 As Long
    Dim q155 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        p155 = .BufPos
        q155 = .ThunkPos
        pvPushThunk ucsActResultClear
        If ParseUNOP() Then
            pvPushThunk ucsActVarSet, 1
            pvPushThunk ucsActResultClear
            If Parseunexpr() Then
                pvPushThunk ucsActVarSet, 2
                pvPushThunk ucsAct_1_unexpr, .CaptureBegin, .CaptureEnd
                GoTo L31
            End If
            .BufPos = p155
            .ThunkPos = q155
        Else
            .BufPos = p155
            .ThunkPos = q155
        End If
        If Parseprimary() Then
            GoTo L31
        End If
        .BufPos = p155
        .ThunkPos = q155
        Exit Function
L31:
        pvPushThunk ucsActVarAlloc, -2
        Parseunexpr = True
    End With
End Function

Private Function Parseprimary() As Boolean
    Dim p162 As Long
    Dim q162 As Long

    With ctx
        p162 = .BufPos
        q162 = .ThunkPos
        If ParseDOUBLE() Then
            GoTo L33
        End If
        .BufPos = p162
        .ThunkPos = q162
        If ParseLPAREN() Then
            If Parseexpr() Then
                If ParseRPAREN() Then
                    GoTo L33
                End If
                .BufPos = p162
                .ThunkPos = q162
            Else
                .BufPos = p162
                .ThunkPos = q162
            End If
        Else
            .BufPos = p162
            .ThunkPos = q162
        End If
        If Parseinvokeexpr() Then
            GoTo L33
        End If
        .BufPos = p162
        .ThunkPos = q162
        If Parseifexpr() Then
            GoTo L33
        End If
        .BufPos = p162
        .ThunkPos = q162
        If Parseforexpr() Then
            GoTo L33
        End If
        .BufPos = p162
        .ThunkPos = q162
        If Parsevarexpr() Then
            GoTo L33
        End If
        .BufPos = p162
        .ThunkPos = q162
        If ParseIDENT() Then
            GoTo L33
        End If
        .BufPos = p162
        .ThunkPos = q162
        Exit Function
L33:
        Parseprimary = True
    End With
End Function

Private Function ParseDOUBLE() As Boolean
    Dim lCaptureBegin As Long
    Dim p394 As Long
    Dim i382 As Long
    Dim i393 As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        p394 = .BufPos
        For i382 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i382 <> 0 Then
            If .BufData(.BufPos) = 46 Then          ' "."
                .BufPos = .BufPos + 1
            End If
            Do
                Select Case .BufData(.BufPos)
                Case 48 To 57                       ' [0-9]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit Do
                End Select
            Loop
            GoTo L36
        End If
        .BufPos = p394
        Do
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit Do
            End Select
        Loop
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
            For i393 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 48 To 57                       ' [0-9]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If i393 <> 0 Then
                GoTo L36
            End If
            .BufPos = p394
        Else
            .BufPos = p394
        End If
        Exit Function
L36:
        lCaptureEnd = .BufPos
        Call Parse_
        .CaptureBegin = lCaptureBegin
        .CaptureEnd = lCaptureEnd
        pvPushThunk ucsAct_1_DOUBLE, lCaptureBegin, lCaptureEnd
        Call pvSetAdvance
        ParseDOUBLE = True
    End With
End Function

Private Function Parseinvokeexpr() As Boolean
    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        If ParseIDENT() Then
            pvPushThunk ucsActVarSet, 1
            If ParseLPAREN() Then
                pvPushThunk ucsActResultClear
                If Parsearg_list() Then
                    pvPushThunk ucsActVarSet, 2
                    If ParseRPAREN() Then
                        pvPushThunk ucsAct_1_invokeexpr, .CaptureBegin, .CaptureEnd
                        pvPushThunk ucsActVarAlloc, -2
                        Parseinvokeexpr = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function Parseifexpr() As Boolean
    Dim p202 As Long
    Dim q202 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 3
        If .BufData(.BufPos) = 105 And .BufData(.BufPos + 1) = 102 Then ' "if"
            .BufPos = .BufPos + 2
            If Parsens() Then
                pvPushThunk ucsActResultClear
                If Parseexpr() Then
                    pvPushThunk ucsActVarSet, 1
                    If pvMatchString("then") Then   ' "then"
                        .BufPos = .BufPos + 4
                        If Parsens() Then
                            pvPushThunk ucsActResultClear
                            If Parseexpr() Then
                                pvPushThunk ucsActVarSet, 2
                                p202 = .BufPos
                                q202 = .ThunkPos
                                If pvMatchString("else") Then ' "else"
                                    .BufPos = .BufPos + 4
                                    If Parsens() Then
                                        pvPushThunk ucsActResultClear
                                        If Parseexpr() Then
                                            pvPushThunk ucsActVarSet, 3
                                        Else
                                            .BufPos = p202
                                            .ThunkPos = q202
                                        End If
                                    Else
                                        .BufPos = p202
                                        .ThunkPos = q202
                                    End If
                                End If
                                pvPushThunk ucsAct_1_ifexpr, .CaptureBegin, .CaptureEnd
                                pvPushThunk ucsActVarAlloc, -3
                                Call pvSetAdvance
                                Parseifexpr = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function Parseforexpr() As Boolean
    Dim p234 As Long
    Dim q234 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 5
        If pvMatchString("for") Then                ' "for"
            .BufPos = .BufPos + 3
            If Parsens() Then
                pvPushThunk ucsActResultClear
                If ParseIDENT() Then
                    pvPushThunk ucsActVarSet, 1
                    If ParseEQ() Then
                        pvPushThunk ucsActResultClear
                        If Parseexpr() Then
                            pvPushThunk ucsActVarSet, 2
                            If ParseCOMMA() Then
                                pvPushThunk ucsActResultClear
                                If Parseexpr() Then
                                    pvPushThunk ucsActVarSet, 3
                                    p234 = .BufPos
                                    q234 = .ThunkPos
                                    If ParseCOMMA() Then
                                        pvPushThunk ucsActResultClear
                                        If Parseexpr() Then
                                            pvPushThunk ucsActVarSet, 4
                                        Else
                                            .BufPos = p234
                                            .ThunkPos = q234
                                        End If
                                    Else
                                        .BufPos = p234
                                        .ThunkPos = q234
                                    End If
                                    If .BufData(.BufPos) = 105 And .BufData(.BufPos + 1) = 110 Then ' "in"
                                        .BufPos = .BufPos + 2
                                        If Parsens() Then
                                            pvPushThunk ucsActResultClear
                                            If Parseexpr() Then
                                                pvPushThunk ucsActVarSet, 5
                                                pvPushThunk ucsAct_1_forexpr, .CaptureBegin, .CaptureEnd
                                                pvPushThunk ucsActVarAlloc, -5
                                                Call pvSetAdvance
                                                Parseforexpr = True
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
    End With
End Function

Private Function Parsevarexpr() As Boolean
    With ctx
        pvPushThunk ucsActVarAlloc, 2
        If pvMatchString("var") Then                ' "var"
            .BufPos = .BufPos + 3
            If Parsens() Then
                pvPushThunk ucsActResultClear
                If Parsevar_list() Then
                    pvPushThunk ucsActVarSet, 1
                    If .BufData(.BufPos) = 105 And .BufData(.BufPos + 1) = 110 Then ' "in"
                        .BufPos = .BufPos + 2
                        If Parsens() Then
                            pvPushThunk ucsActResultClear
                            If Parseexpr() Then
                                pvPushThunk ucsActVarSet, 2
                                pvPushThunk ucsAct_1_varexpr, .CaptureBegin, .CaptureEnd
                                pvPushThunk ucsActVarAlloc, -2
                                Call pvSetAdvance
                                Parsevarexpr = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function Parsearg_list() As Boolean
    Dim p286 As Long
    Dim q286 As Long
    Dim p280 As Long
    Dim q280 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        If Parseexpr() Then
            pvPushThunk ucsActVarSet, 1
            pvPushThunk ucsAct_1_arg_list, .CaptureBegin, .CaptureEnd
            Do
                p286 = .BufPos
                q286 = .ThunkPos
                p280 = .BufPos
                q280 = .ThunkPos
                If Not (ParseCOMMA()) Then
                    .BufPos = p280
                    .ThunkPos = q280
                End If
                pvPushThunk ucsActResultClear
                If Parseexpr() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p286
                    .ThunkPos = q286
                    Exit Do
                End If
                pvPushThunk ucsAct_2_arg_list, .CaptureBegin, .CaptureEnd
            Loop
            pvPushThunk ucsAct_3_arg_list, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -2
            Parsearg_list = True
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

Private Function ParseCOMMA() As Boolean
    With ctx
        If .BufData(.BufPos) = 44 Then              ' ","
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseCOMMA = True
        End If
    End With
End Function

Private Function Parsevar_list() As Boolean
    Dim p302 As Long
    Dim q302 As Long
    Dim p296 As Long
    Dim q296 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        If Parsevar_decl() Then
            pvPushThunk ucsActVarSet, 1
            pvPushThunk ucsAct_1_var_list, .CaptureBegin, .CaptureEnd
            Do
                p302 = .BufPos
                q302 = .ThunkPos
                p296 = .BufPos
                q296 = .ThunkPos
                If Not (ParseCOMMA()) Then
                    .BufPos = p296
                    .ThunkPos = q296
                End If
                pvPushThunk ucsActResultClear
                If Parsevar_decl() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p302
                    .ThunkPos = q302
                    Exit Do
                End If
                pvPushThunk ucsAct_2_var_list, .CaptureBegin, .CaptureEnd
            Loop
            pvPushThunk ucsAct_3_var_list, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -2
            Parsevar_list = True
        End If
    End With
End Function

Private Function Parsevar_decl() As Boolean
    Dim p320 As Long
    Dim q320 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        If ParseIDENT() Then
            pvPushThunk ucsActVarSet, 1
            p320 = .BufPos
            q320 = .ThunkPos
            If ParseEQ() Then
                pvPushThunk ucsActResultClear
                If Parseexpr() Then
                    pvPushThunk ucsActVarSet, 2
                    GoTo L55
                End If
                .BufPos = p320
                .ThunkPos = q320
            Else
                .BufPos = p320
                .ThunkPos = q320
            End If
            pvPushThunk ucsActResultClear
            Call ParseEMPTY
            pvPushThunk ucsActVarSet, 2
            GoTo L55
            Exit Function
L55:
            pvPushThunk ucsAct_1_var_decl, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -2
            Parsevar_decl = True
        End If
    End With
End Function

Private Sub ParseEMPTY()
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        lCaptureEnd = .BufPos
        .CaptureBegin = lCaptureBegin
        .CaptureEnd = lCaptureEnd
    End With
End Sub

Private Function ParseLineComment() As Boolean
    With ctx
        If .BufData(.BufPos) = 35 Then              ' "#"
            .BufPos = .BufPos + 1
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
            ParseLineComment = True
        End If
    End With
End Function

Private Sub pvImplAction(ByVal eAction As UcsParserActionsEnum, ByVal lOffset As Long, ByVal lSize As Long)
    Dim oJson As Object

    With ctx
    Select Case eAction
    Case ucsAct_3_start
           Set .VarResult = .VarStack(.VarPos - 1)
    Case ucsAct_2_start
           Set oJson = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, -1) = .VarStack(.VarPos - 2)
    Case ucsAct_1_start
           JsonItem(oJson, -1) = Empty
                                                                            Set .VarStack(.VarPos - 1) = oJson
    Case ucsAct_1_definition
           JsonItem(oJson, "Tag") = "Def"
                                                                            JsonItem(oJson, "Proto") = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, "Body") = .VarStack(.VarPos - 2)
                                                                            Set .VarResult = oJson
    Case ucsAct_1_extern
           JsonItem(oJson, "Tag") = "Extern"
                                                                            JsonItem(oJson, "Proto") = .VarStack(.VarPos - 1)
                                                                            Set .VarResult = oJson
    Case ucsAct_1_prototype
           JsonItem(oJson, "Tag") = "UnOpProto"
                                                                            JsonItem(oJson, "Oper") = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, "Params") = .VarStack(.VarPos - 2)
                                                                            Set .VarResult = oJson
    Case ucsAct_2_prototype
           JsonItem(oJson, "Tag") = "FuncProto"
                                                                            JsonItem(oJson, "Name") = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, "Params") = .VarStack(.VarPos - 2)
                                                                            Set .VarResult = oJson
    Case ucsAct_1_binopproto
           JsonItem(oJson, "Tag") = "BinOpProto"
                                                                            JsonItem(oJson, "Oper") = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, "Precedence") = .VarStack(.VarPos - 2)
                                                                            JsonItem(oJson, "Params") = .VarStack(.VarPos - 3)
                                                                            Set .VarResult = oJson
    Case ucsAct_1_UNOP
         .VarResult = Mid$(.Contents, lOffset, lSize)
    Case ucsAct_3_param_list
           Set .VarResult = .VarStack(.VarPos - 1)
    Case ucsAct_2_param_list
           Set oJson = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, -1) = .VarStack(.VarPos - 2)
    Case ucsAct_1_param_list
           JsonItem(oJson, -1) = .VarStack(.VarPos - 1)
                                                                            Set .VarStack(.VarPos - 1) = oJson
    Case ucsAct_1_IDENT
         .VarResult = Mid$(.Contents, lOffset, lSize)
    Case ucsAct_1_BINOP
         .VarResult = Mid$(.Contents, lOffset, lSize)
    Case ucsAct_1_NUMBER
         .VarResult = C_Lng(Mid$(.Contents, lOffset, lSize))
    Case ucsAct_2_binexpr
           AssignVariant .VarResult, .VarStack(.VarPos - 1)
    Case ucsAct_1_binexpr
           JsonItem(oJson, "Tag") = "BinOp"
                                                                            JsonItem(oJson, "Left") = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, "Oper") = .VarStack(.VarPos - 2)
                                                                            JsonItem(oJson, "Right") = .VarStack(.VarPos - 3)
                                                                            JsonItem(oJson, "BinopInfo") = GetBinopInfo(C_Str(.VarStack(.VarPos - 2)))
                                                                            Set .VarStack(.VarPos - 1) = oJson
    Case ucsAct_1_unexpr
           JsonItem(oJson, "Tag") = "UnOp"
                                                                            JsonItem(oJson, "Oper") = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, "Expr") = .VarStack(.VarPos - 2)
                                                                            Set .VarResult = oJson
    Case ucsAct_1_DOUBLE
         .VarResult = C_Dbl(Mid$(.Contents, lOffset, lSize))
    Case ucsAct_1_invokeexpr
           JsonItem(oJson, "Tag") = "InvokeExpr"
                                                                            JsonItem(oJson, "Func") = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, "Args") = .VarStack(.VarPos - 2)
                                                                            Set .VarResult = oJson
    Case ucsAct_1_ifexpr
           JsonItem(oJson, "Tag") = "IfExpr"
                                                                            JsonItem(oJson, "Cond") = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, "Then") = .VarStack(.VarPos - 2)
                                                                            JsonItem(oJson, "Else") = .VarStack(.VarPos - 3)
                                                                            Set .VarResult = oJson
    Case ucsAct_1_forexpr
           JsonItem(oJson, "Tag") = "ForExpr"
                                                                            JsonItem(oJson, "Var") = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, "Init") = .VarStack(.VarPos - 2)
                                                                            JsonItem(oJson, "Last") = .VarStack(.VarPos - 3)
                                                                            JsonItem(oJson, "Step") = .VarStack(.VarPos - 4)
                                                                            JsonItem(oJson, "Body") = .VarStack(.VarPos - 5)
                                                                            Set .VarResult = oJson
    Case ucsAct_1_varexpr
           JsonItem(oJson, "Tag") = "VarExpr"
                                                                            JsonItem(oJson, "Vars") = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, "Body") = .VarStack(.VarPos - 2)
                                                                            Set .VarResult = oJson
    Case ucsAct_3_arg_list
           Set .VarResult = .VarStack(.VarPos - 1)
    Case ucsAct_2_arg_list
           Set oJson = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, -1) = .VarStack(.VarPos - 2)
    Case ucsAct_1_arg_list
           JsonItem(oJson, -1) = .VarStack(.VarPos - 1)
                                                                            Set .VarStack(.VarPos - 1) = oJson
    Case ucsAct_3_var_list
           Set .VarResult = .VarStack(.VarPos - 1)
    Case ucsAct_2_var_list
           Set oJson = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, -1) = .VarStack(.VarPos - 2)
    Case ucsAct_1_var_list
           JsonItem(oJson, -1) = .VarStack(.VarPos - 1)
                                                                            Set .VarStack(.VarPos - 1) = oJson
    Case ucsAct_1_var_decl
           JsonItem(oJson, "Tag") = "VarDecl"
                                                                            JsonItem(oJson, "Name") = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, "Expr") = .VarStack(.VarPos - 2)
    End Select
    End With
End Sub

'= user-defined functions ================================================

Public Function GetBinopInfo(sOp As String) As Variant
    If m_oBinopInfo Is Nothing Then
        Set m_oBinopInfo = New Scripting.Dictionary
        m_oBinopInfo.Item("<") = Array(10, 1)
        m_oBinopInfo.Item("+") = Array(20, 1)
        m_oBinopInfo.Item("-") = Array(20, 1)
        m_oBinopInfo.Item("*") = Array(40, 1)
'        m_oBinopInfo.Item("/") = Array(40, 1)
'        m_oBinopInfo.Item("^") = Array(50, 0)
    End If
    If m_oBinopInfo.Exists(sOp) Then
        GetBinopInfo = m_oBinopInfo.Item(sOp)
    Else
        GetBinopInfo = Array(0, 0)
    End If
End Function

Public Function SetBinopInfo(sOp As String, ByVal lPrec As Long,  ByVal lAssoc As Long) As Variant
    If m_oBinopInfo Is Nothing Then
        GetBinopInfo vbNullString
    End If
    m_oBinopInfo.Item(sOp) = Array(lPrec, lAssoc, "Override")
End Function
