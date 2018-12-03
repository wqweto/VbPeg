VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Auto-generated on 3.12.2018 14:33:00
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

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
    ucsAct_2_Setting
    ucsAct_1_Setting
    ucsAct_5_Rule
    ucsAct_4_Rule
    ucsAct_3_Rule
    ucsAct_2_Rule
    ucsAct_1_Rule
    ucsAct_1_ChoiceExpr
    ucsAct_2_SequenceExpr
    ucsAct_1_SequenceExpr
    ucsAct_1_ErrorExpr
    ucsAct_2_ErrorExpr
    ucsAct_2_VariableExpr
    ucsAct_1_VariableExpr
    ucsAct_1_PrefixExpr
    ucsAct_2_PrefixExpr
    ucsAct_3_PrefixExpr
    ucsAct_1_SuffixExpr
    ucsAct_2_SuffixExpr
    ucsAct_3_SuffixExpr
    ucsAct_2_PrimaryExpr
    ucsAct_1_PrimaryExpr
    ucsAct_4_PrimaryExpr
    ucsAct_3_PrimaryExpr
    ucsAct_6_PrimaryExpr
    ucsAct_5_PrimaryExpr
    ucsAct_7_PrimaryExpr
    ucsAct_8_PrimaryExpr
    ucsAct_9_PrimaryExpr
    ucsAct_10_PrimaryExpr
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
    UserData            As cTree
    VarResult           As Variant
    '-- from @members setting
    ParseErrors         As Collection
    ParamRules          As Collection
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
    ParserVersion = "3.12.2018 14:33:00"
End Property

Property Get Contents(Optional ByVal lOffset As Long = 1, Optional ByVal lSize As Long = LNG_MAXINT) As String
    Contents = Mid$(ctx.Contents, lOffset, lSize)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Match(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As cTree, Optional Result As Variant) As Long
    If BeginMatch(sSubject, StartPos, UserData) Then
        If ParseGrammar() Then
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

Public Function BeginMatch(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As cTree) As Boolean
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
        For lIdx = 0 To .ThunkPos - 1
            With .ThunkData(lIdx)
                pvImplAction .Action, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin
            End With
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

Public Function ParseGrammar() As Boolean
    Dim p24 As Long
    Dim q24 As Long
    Dim p6 As Long
    Dim i16 As Long
    Dim p15 As Long
    Dim q15 As Long
    Dim p20 As Long
    Dim q20 As Long

    With ctx
        p24 = .BufPos
        q24 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Grammar", 1, ctx.UserData)
        p6 = .BufPos
        If Not (ParsePreprocess()) Then
            .BufPos = p6
        End If
        Call ParseSpacing
        For i16 = 0 To LNG_MAXINT
            p15 = .BufPos
            q15 = .ThunkPos
            If Not (ParseSetting()) Then
                .BufPos = p15
                .ThunkPos = q15
                If Not (ParseRule()) Then
                    .BufPos = p15
                    .ThunkPos = q15
                    .BufPos = p15
                    .ThunkPos = q15
                    Exit For
                End If
            End If
        Next
        If i16 <> 0 Then
            p20 = .BufPos
            q20 = .ThunkPos
            If ParseEndOfFile() Then
                Call ConsoleTrace(ctx.BufPos + 1, "Grammar", 2, ctx.UserData)
                ParseGrammar = True
                Exit Function
            Else
                .BufPos = p20
                .ThunkPos = q20
            End If
            If AppendError("Expected rule or setting", ctx.BufPos + 1) Then
                Call ConsoleTrace(ctx.BufPos + 1, "Grammar", 2, ctx.UserData)
                ParseGrammar = True
                Exit Function
            End If
            .BufPos = p24
            .ThunkPos = q24
        Else
            .BufPos = p24
            .ThunkPos = q24
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Grammar", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParsePreprocess() As Boolean
    Dim p1244 As Long
    Dim i1238 As Long
    Dim p1234 As Long
    Dim p1226 As Long
    Dim p1228 As Long

    With ctx
        p1244 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Preprocess", 1, ctx.UserData)
        Call ParseSpacing
        For i1238 = 0 To LNG_MAXINT
            p1234 = .BufPos
            If ParseIdentifier() Then
                If ParseParamNC() Then
                    p1226 = .BufPos
                    If Not (ParseLiteralNC()) Then
                        .BufPos = p1226
                    End If
                    Do
                        p1228 = .BufPos
                        If Not (ParseRuleOptNC()) Then
                            .BufPos = p1228
                            Exit Do
                        End If
                    Loop
                    If ParseAssignLiteral() Then
                        '--- this might back-track
                        AddParamRule Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)
                    Else
                        .BufPos = p1234
                        If Not (ParsePrepExpr()) Then
                            .BufPos = p1234
                            If .BufPos < .BufSize Then
                                .BufPos = .BufPos + 1
                                Call ParseSpacing
                            Else
                                .BufPos = p1234
                                Exit For
                            End If
                        End If
                    End If
                Else
                    .BufPos = p1234
                    If Not (ParsePrepExpr()) Then
                        .BufPos = p1234
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                            Call ParseSpacing
                        Else
                            .BufPos = p1234
                            Exit For
                        End If
                    End If
                End If
            Else
                .BufPos = p1234
                If Not (ParsePrepExpr()) Then
                    .BufPos = p1234
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        Call ParseSpacing
                    Else
                        .BufPos = p1234
                        Exit For
                    End If
                End If
            End If
        Next
        If i1238 <> 0 Then
            '-- rewind position
            ctx.BufPos = 0
            Call ConsoleTrace(ctx.BufPos + 1, "Preprocess", 2, ctx.UserData)
            ParsePreprocess = True
            Exit Function
        Else
            .BufPos = p1244
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Preprocess", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Sub ParseSpacing()
    Dim p980 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "Spacing", 1, ctx.UserData)
        Do
            p980 = .BufPos
            If Not (ParseSpace()) Then
                .BufPos = p980
                If Not (ParseLineComment()) Then
                    .BufPos = p980
                    If Not (ParseBlockComment()) Then
                        .BufPos = p980
                        .BufPos = p980
                        Exit Do
                    End If
                End If
            End If
        Loop
        Call ConsoleTrace(ctx.BufPos + 1, "Spacing", 2, ctx.UserData)
    End With
End Sub

Private Function ParseSetting() As Boolean
    Dim p265 As Long
    Dim q265 As Long
    Dim p261 As Long
    Dim q261 As Long

    With ctx
        p265 = .BufPos
        q265 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 1, ctx.UserData)
        If ParseSettingName() Then
            If CheckSettingName(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                pvPushThunk ucsAct_1_Setting, .CaptureBegin, .CaptureEnd
                If ClearErrors(ctx.BufPos + 1) And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                    p261 = .BufPos
                    q261 = .ThunkPos
                    If ParseSettingValue() Then
                        pvPushThunk ucsAct_2_Setting, .CaptureBegin, .CaptureEnd
                        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 2, ctx.UserData)
                        ParseSetting = True
                        Exit Function
                    Else
                        .BufPos = p261
                        .ThunkPos = q261
                    End If
                    If AppendError("Invalid value for setting '" & .UserData.Tag & "'", ctx.BufPos + 1) Then
                        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 2, ctx.UserData)
                        ParseSetting = True
                        Exit Function
                    End If
                    .BufPos = p265
                    .ThunkPos = q265
                Else
                    .BufPos = p265
                    .ThunkPos = q265
                End If
            Else
                .BufPos = p265
                .ThunkPos = q265
            End If
        Else
            .BufPos = p265
            .ThunkPos = q265
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Setting", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseRule() As Boolean
    Dim p61 As Long
    Dim q61 As Long
    Dim p34 As Long
    Dim q34 As Long
    Dim p39 As Long
    Dim q39 As Long
    Dim p44 As Long
    Dim q44 As Long
    Dim p56 As Long
    Dim q56 As Long

    With ctx
        p61 = .BufPos
        q61 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Rule", 1, ctx.UserData)
        If ParseIdentifier() Then
            pvPushThunk ucsAct_1_Rule, .CaptureBegin, .CaptureEnd
            If ClearErrors(ctx.BufPos + 1) And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                p34 = .BufPos
                q34 = .ThunkPos
                If ParseParamBlock() Then
                    pvPushThunk ucsAct_2_Rule, .CaptureBegin, .CaptureEnd
                Else
                    .BufPos = p34
                    .ThunkPos = q34
                End If
                p39 = .BufPos
                q39 = .ThunkPos
                If ParseLiteral() Then
                    pvPushThunk ucsAct_3_Rule, .CaptureBegin, .CaptureEnd
                Else
                    .BufPos = p39
                    .ThunkPos = q39
                End If
                Do
                    p44 = .BufPos
                    q44 = .ThunkPos
                    If Not (ParseRuleOption()) Then
                        .BufPos = p44
                        .ThunkPos = q44
                        Exit Do
                    End If
                    If Not (CheckRuleOptionName(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin))) Then
                        .BufPos = p44
                        .ThunkPos = q44
                        Exit Do
                    End If
                    pvPushThunk ucsAct_4_Rule, .CaptureBegin, .CaptureEnd
                Loop
                If ParseAssignLiteral() Then
                    Call ParseExpression
                    pvPushThunk ucsAct_5_Rule, .CaptureBegin, .CaptureEnd
                    p56 = .BufPos
                    q56 = .ThunkPos
                    If Not (ParseSEMI()) Then
                        .BufPos = p56
                        .ThunkPos = q56
                    End If
                    Call ConsoleTrace(ctx.BufPos + 1, "Rule", 2, ctx.UserData)
                    ParseRule = True
                    Exit Function
                Else
                    .BufPos = p61
                    .ThunkPos = q61
                End If
            Else
                .BufPos = p61
                .ThunkPos = q61
            End If
        Else
            .BufPos = p61
            .ThunkPos = q61
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Rule", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseEndOfFile() As Boolean
    Dim p1098 As Long
    Dim p1094 As Long

    With ctx
        p1098 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "EndOfFile", 1, ctx.UserData)
        p1094 = .BufPos
        If .BufPos < .BufSize Then
            .BufPos = p1098
        Else
            .BufPos = p1094
            Call ConsoleTrace(ctx.BufPos + 1, "EndOfFile", 2, ctx.UserData)
            ParseEndOfFile = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "EndOfFile", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseIdentifier() As Boolean
    Dim p466 As Long
    Dim lCaptureBegin As Long
    Dim p459 As Long
    Dim lCaptureEnd As Long

    With ctx
        p466 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Identifier", 1, ctx.UserData)
        lCaptureBegin = .BufPos
        If ParseIdentStart() Then
            Do
                p459 = .BufPos
                If Not (ParseIdentCont()) Then
                    .BufPos = p459
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "Identifier", 2, ctx.UserData)
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            ParseIdentifier = True
            Exit Function
        Else
            .BufPos = p466
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Identifier", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseParamBlock() As Boolean
    Dim p396 As Long
    Dim p363 As Long
    Dim p373 As Long
    Dim p369 As Long
    Dim lCaptureBegin As Long
    Dim p381 As Long
    Dim lCaptureEnd As Long
    Dim p390 As Long

    With ctx
        p396 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ParamBlock", 1, ctx.UserData)
        p363 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """(""", 0, ctx.UserData)
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """(""", 2, ctx.UserData)
            Do
                p373 = .BufPos
                Do
                    p369 = .BufPos
                    Call ConsoleTrace(ctx.BufPos + 1, "[ \t]", 0, ctx.UserData)
                    Select Case .BufData(.BufPos)
                    Case 32, 9                      ' [ \t]
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[ \t]", 2, ctx.UserData)
                    Case Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[ \t]", 3, ctx.UserData) And False) Then
                            .BufPos = p369
                            Exit Do
                        End If
                    End Select
                Loop
                If Not (ParseEndOfLine()) Then
                    .BufPos = p373
                    Exit Do
                End If
            Loop
            lCaptureBegin = .BufPos
            Do
                p381 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^()]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 40, 41                         ' [()]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^()]", 3, ctx.UserData) And False) Then
                        If Not (ParseParamNC()) Then
                            .BufPos = p381
                            .BufPos = p381
                            Exit Do
                        End If
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[^()]", 2, ctx.UserData)
                    Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[^()]", 3, ctx.UserData) And False) Then
                            If Not (ParseParamNC()) Then
                                .BufPos = p381
                                .BufPos = p381
                                Exit Do
                            End If
                        End If
                    End If
                End Select
            Loop
            lCaptureEnd = .BufPos
            p390 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, """)""", 0, ctx.UserData)
            If .BufData(.BufPos) = 41 Then          ' ")"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, """)""", 2, ctx.UserData)
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "ParamBlock", 2, ctx.UserData)
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call pvSetAdvance
                ParseParamBlock = True
                Exit Function
            End If
            If ConsoleTrace(ctx.BufPos + 1, """)""", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p396
            GoTo L12
        End If
        If ConsoleTrace(ctx.BufPos + 1, """(""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p396
L12:
        If ConsoleTrace(ctx.BufPos + 1, "ParamBlock", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseLiteral() As Boolean
    Dim p496 As Long

    With ctx
        p496 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Literal", 1, ctx.UserData)
        If ParseLiteralNoSpace() Then
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "Literal", 2, ctx.UserData)
            ParseLiteral = True
            Exit Function
        Else
            .BufPos = p496
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Literal", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseRuleOption() As Boolean
    Dim p436 As Long
    Dim p429 As Long

    With ctx
        p436 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "RuleOption", 1, ctx.UserData)
        p429 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """-""", 0, ctx.UserData)
        If .BufData(.BufPos) = 45 Then              ' "-"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """-""", 2, ctx.UserData)
            If ParseIdentifier() Then
                Call ConsoleTrace(ctx.BufPos + 1, "RuleOption", 2, ctx.UserData)
                Call pvSetAdvance
                ParseRuleOption = True
                Exit Function
            Else
                .BufPos = p436
            End If
            GoTo L17
        End If
        If ConsoleTrace(ctx.BufPos + 1, """-""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p436
L17:
        If ConsoleTrace(ctx.BufPos + 1, "RuleOption", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseAssignLiteral() As Boolean
    Dim p1133 As Long
    Dim p1129 As Long

    With ctx
        p1133 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 1, ctx.UserData)
        p1129 = .BufPos
        If ParseLEFTARROW() Then
            Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 2, ctx.UserData)
            ParseAssignLiteral = True
            Exit Function
        Else
            .BufPos = p1129
        End If
        If ParseEQUAL() Then
            Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 2, ctx.UserData)
            ParseAssignLiteral = True
            Exit Function
        Else
            .BufPos = p1129
        End If
        .BufPos = p1133
        If ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Sub ParseExpression()
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "Expression", 1, ctx.UserData)
        Call ParseChoiceExpr
        Call ConsoleTrace(ctx.BufPos + 1, "Expression", 2, ctx.UserData)
L18:
    End With
End Sub

Private Function ParseSEMI() As Boolean
    Dim p1194 As Long
    Dim p1187 As Long

    With ctx
        p1194 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SEMI", 1, ctx.UserData)
        p1187 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """;""", 0, ctx.UserData)
        If .BufData(.BufPos) = 59 Then              ' ";"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """;""", 2, ctx.UserData)
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "SEMI", 2, ctx.UserData)
            Call pvSetAdvance
            ParseSEMI = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """;""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p1194
        If ConsoleTrace(ctx.BufPos + 1, "SEMI", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Sub ParseChoiceExpr()
    Dim p75 As Long
    Dim q75 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceExpr", 1, ctx.UserData)
        Call ParseSequenceExpr
        Do
            p75 = .BufPos
            q75 = .ThunkPos
            If Not (ParseChoiceLiteral()) Then
                .BufPos = p75
                .ThunkPos = q75
                Exit Do
            End If
            Call ParseSequenceExpr
            pvPushThunk ucsAct_1_ChoiceExpr, .CaptureBegin, .CaptureEnd
        Loop
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceExpr", 2, ctx.UserData)
L19:
    End With
End Sub

Private Sub ParseSequenceExpr()
    Dim p92 As Long
    Dim q92 As Long
    Dim p88 As Long
    Dim q88 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "SequenceExpr", 1, ctx.UserData)
        p92 = .BufPos
        q92 = .ThunkPos
        If ParseErrorExpr() Then
            Do
                p88 = .BufPos
                q88 = .ThunkPos
                If Not (ParseErrorExpr()) Then
                    .BufPos = p88
                    .ThunkPos = q88
                    Exit Do
                End If
                pvPushThunk ucsAct_1_SequenceExpr, .CaptureBegin, .CaptureEnd
            Loop
        Else
            .BufPos = p92
            .ThunkPos = q92
            pvPushThunk ucsAct_2_SequenceExpr, .CaptureBegin, .CaptureEnd
        End If
        Call ConsoleTrace(ctx.BufPos + 1, "SequenceExpr", 2, ctx.UserData)
    End With
End Sub

Private Function ParseChoiceLiteral() As Boolean
    Dim p1142 As Long
    Dim p1138 As Long

    With ctx
        p1142 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 1, ctx.UserData)
        p1138 = .BufPos
        If ParseSLASH() Then
            Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 2, ctx.UserData)
            ParseChoiceLiteral = True
            Exit Function
        Else
            .BufPos = p1138
        End If
        If ParseBAR() Then
            Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 2, ctx.UserData)
            ParseChoiceLiteral = True
            Exit Function
        Else
            .BufPos = p1138
        End If
        .BufPos = p1142
        If ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseErrorExpr() As Boolean
    Dim p117 As Long
    Dim q117 As Long
    Dim p103 As Long
    Dim q103 As Long
    Dim p110 As Long
    Dim q110 As Long

    With ctx
        p117 = .BufPos
        q117 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 1, ctx.UserData)
        If ParseVariableExpr() Then
            p103 = .BufPos
            q103 = .ThunkPos
            If ParseCARET() Then
                If ClearErrors(ctx.BufPos + 1) Then
                    p110 = .BufPos
                    q110 = .ThunkPos
                    If ParseIdentifier() Then
                        pvPushThunk ucsAct_1_ErrorExpr, .CaptureBegin, .CaptureEnd
                        GoTo L22
                    Else
                        .BufPos = p110
                        .ThunkPos = q110
                    End If
                    If ParseLiteral() Then
                        pvPushThunk ucsAct_2_ErrorExpr, .CaptureBegin, .CaptureEnd
                        GoTo L22
                    Else
                        .BufPos = p110
                        .ThunkPos = q110
                    End If
                    If AppendError("Expected rule or literal for error label", ctx.BufPos + 1) Then
                        GoTo L22
                    End If
                    .BufPos = p103
                    .ThunkPos = q103
                Else
                    .BufPos = p103
                    .ThunkPos = q103
                End If
            Else
                .BufPos = p103
                .ThunkPos = q103
            End If
L22:
            Call ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 2, ctx.UserData)
            ParseErrorExpr = True
            Exit Function
        Else
            .BufPos = p117
            .ThunkPos = q117
        End If
        If ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseVariableExpr() As Boolean
    Dim p132 As Long
    Dim q132 As Long
    Dim p128 As Long
    Dim q128 As Long

    With ctx
        p132 = .BufPos
        q132 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 1, ctx.UserData)
        p128 = .BufPos
        q128 = .ThunkPos
        If ParseIdentifier() Then
            pvPushThunk ucsAct_1_VariableExpr, .CaptureBegin, .CaptureEnd
            If ParseCOLON() Then
                If ParsePrefixExpr() Then
                    pvPushThunk ucsAct_2_VariableExpr, .CaptureBegin, .CaptureEnd
                    Call ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 2, ctx.UserData)
                    ParseVariableExpr = True
                    Exit Function
                Else
                    .BufPos = p128
                    .ThunkPos = q128
                End If
            Else
                .BufPos = p128
                .ThunkPos = q128
            End If
        Else
            .BufPos = p128
            .ThunkPos = q128
        End If
        If ParsePrefixExpr() Then
            Call ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 2, ctx.UserData)
            ParseVariableExpr = True
            Exit Function
        Else
            .BufPos = p128
            .ThunkPos = q128
        End If
        .BufPos = p132
        .ThunkPos = q132
        If ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCARET() As Boolean
    Dim p1220 As Long
    Dim p1213 As Long

    With ctx
        p1220 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CARET", 1, ctx.UserData)
        p1213 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """^""", 0, ctx.UserData)
        If .BufData(.BufPos) = 94 Then              ' "^"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """^""", 2, ctx.UserData)
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "CARET", 2, ctx.UserData)
            Call pvSetAdvance
            ParseCARET = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """^""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p1220
        If ConsoleTrace(ctx.BufPos + 1, "CARET", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCOLON() As Boolean
    Dim p1181 As Long
    Dim p1174 As Long

    With ctx
        p1181 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "COLON", 1, ctx.UserData)
        p1174 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """:""", 0, ctx.UserData)
        If .BufData(.BufPos) = 58 Then              ' ":"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """:""", 2, ctx.UserData)
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "COLON", 2, ctx.UserData)
            Call pvSetAdvance
            ParseCOLON = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """:""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p1181
        If ConsoleTrace(ctx.BufPos + 1, "COLON", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParsePrefixExpr() As Boolean
    Dim p155 As Long
    Dim q155 As Long
    Dim p145 As Long
    Dim q145 As Long

    With ctx
        p155 = .BufPos
        q155 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 1, ctx.UserData)
        p145 = .BufPos
        q145 = .ThunkPos
        If ParseAND() Then
            If ParseCodeBlock() Then
                pvPushThunk ucsAct_1_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p145
                .ThunkPos = q145
            End If
        Else
            .BufPos = p145
            .ThunkPos = q145
        End If
        If ParseAND() Then
            If ParseSuffixExpr() Then
                pvPushThunk ucsAct_2_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p145
                .ThunkPos = q145
            End If
        Else
            .BufPos = p145
            .ThunkPos = q145
        End If
        If ParseNOT() Then
            If ParseSuffixExpr() Then
                pvPushThunk ucsAct_3_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p145
                .ThunkPos = q145
            End If
        Else
            .BufPos = p145
            .ThunkPos = q145
        End If
        If ParseSuffixExpr() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
            ParsePrefixExpr = True
            Exit Function
        Else
            .BufPos = p145
            .ThunkPos = q145
        End If
        .BufPos = p155
        .ThunkPos = q155
        If ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseAND() As Boolean
    Dim p868 As Long
    Dim p861 As Long

    With ctx
        p868 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "AND", 1, ctx.UserData)
        p861 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """&""", 0, ctx.UserData)
        If .BufData(.BufPos) = 38 Then              ' "&"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """&""", 2, ctx.UserData)
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "AND", 2, ctx.UserData)
            Call pvSetAdvance
            ParseAND = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """&""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p868
        If ConsoleTrace(ctx.BufPos + 1, "AND", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCodeBlock() As Boolean
    Dim p330 As Long
    Dim p295 As Long
    Dim p306 As Long
    Dim p301 As Long
    Dim lCaptureBegin As Long
    Dim p314 As Long
    Dim lCaptureEnd As Long
    Dim p324 As Long

    With ctx
        p330 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CodeBlock", 1, ctx.UserData)
        p295 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """{""", 0, ctx.UserData)
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """{""", 2, ctx.UserData)
            Do
                p306 = .BufPos
                Do
                    p301 = .BufPos
                    Call ConsoleTrace(ctx.BufPos + 1, "[ \t]", 0, ctx.UserData)
                    Select Case .BufData(.BufPos)
                    Case 32, 9                      ' [ \t]
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[ \t]", 2, ctx.UserData)
                    Case Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[ \t]", 3, ctx.UserData) And False) Then
                            .BufPos = p301
                            Exit Do
                        End If
                    End Select
                Loop
                If Not (ParseEndOfLine()) Then
                    .BufPos = p306
                    Exit Do
                End If
            Loop
            lCaptureBegin = .BufPos
            Do
                p314 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^{}]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^{}]", 3, ctx.UserData) And False) Then
                        If Not (ParseCodeNC()) Then
                            .BufPos = p314
                            .BufPos = p314
                            Exit Do
                        End If
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[^{}]", 2, ctx.UserData)
                    Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[^{}]", 3, ctx.UserData) And False) Then
                            If Not (ParseCodeNC()) Then
                                .BufPos = p314
                                .BufPos = p314
                                Exit Do
                            End If
                        End If
                    End If
                End Select
            Loop
            lCaptureEnd = .BufPos
            p324 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, """}""", 0, ctx.UserData)
            If .BufData(.BufPos) = 125 Then         ' "}"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, """}""", 2, ctx.UserData)
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "CodeBlock", 2, ctx.UserData)
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call pvSetAdvance
                ParseCodeBlock = True
                Exit Function
            End If
            If ConsoleTrace(ctx.BufPos + 1, """}""", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p330
            GoTo L25
        End If
        If ConsoleTrace(ctx.BufPos + 1, """{""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p330
L25:
        If ConsoleTrace(ctx.BufPos + 1, "CodeBlock", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSuffixExpr() As Boolean
    Dim p177 As Long
    Dim q177 As Long
    Dim p167 As Long
    Dim q167 As Long

    With ctx
        p177 = .BufPos
        q177 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 1, ctx.UserData)
        If ParsePrimaryExpr() Then
            p167 = .BufPos
            q167 = .ThunkPos
            If ParseQUESTION() Then
                pvPushThunk ucsAct_1_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L30
            Else
                .BufPos = p167
                .ThunkPos = q167
            End If
            If ParseSTAR() Then
                pvPushThunk ucsAct_2_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L30
            Else
                .BufPos = p167
                .ThunkPos = q167
            End If
            If ParsePLUS() Then
                pvPushThunk ucsAct_3_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L30
            Else
                .BufPos = p167
                .ThunkPos = q167
            End If
            .BufPos = p167
            .ThunkPos = q167
L30:
            Call ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 2, ctx.UserData)
            ParseSuffixExpr = True
            Exit Function
        Else
            .BufPos = p177
            .ThunkPos = q177
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseNOT() As Boolean
    Dim p881 As Long
    Dim p874 As Long

    With ctx
        p881 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "NOT", 1, ctx.UserData)
        p874 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """!""", 0, ctx.UserData)
        If .BufData(.BufPos) = 33 Then              ' "!"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """!""", 2, ctx.UserData)
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "NOT", 2, ctx.UserData)
            Call pvSetAdvance
            ParseNOT = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """!""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p881
        If ConsoleTrace(ctx.BufPos + 1, "NOT", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParsePrimaryExpr() As Boolean
    Dim p248 As Long
    Dim q248 As Long
    Dim p209 As Long
    Dim q209 As Long
    Dim p184 As Long
    Dim q184 As Long
    Dim p198 As Long
    Dim q198 As Long
    Dim p188 As Long
    Dim q188 As Long
    Dim p191 As Long
    Dim q191 As Long
    Dim p195 As Long
    Dim q195 As Long
    Dim p208 As Long
    Dim q208 As Long
    Dim p218 As Long
    Dim q218 As Long
    Dim p227 As Long
    Dim q227 As Long

    With ctx
        p248 = .BufPos
        q248 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 1, ctx.UserData)
        p209 = .BufPos
        q209 = .ThunkPos
        If ParseIdentifier() Then
            pvPushThunk ucsAct_1_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p184 = .BufPos
            q184 = .ThunkPos
            If IsParamRule(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                If ParseParamBlock() Then
                    pvPushThunk ucsAct_2_PrimaryExpr, .CaptureBegin, .CaptureEnd
                Else
                    .BufPos = p184
                    .ThunkPos = q184
                End If
            End If
            p198 = .BufPos
            q198 = .ThunkPos
            p188 = .BufPos
            q188 = .ThunkPos
            If Not (ParseParamNC()) Then
                .BufPos = p188
                .ThunkPos = q188
            End If
            p191 = .BufPos
            q191 = .ThunkPos
            If Not (ParseLiteralNC()) Then
                .BufPos = p191
                .ThunkPos = q191
            End If
            Do
                p195 = .BufPos
                q195 = .ThunkPos
                If Not (ParseRuleOptNC()) Then
                    .BufPos = p195
                    .ThunkPos = q195
                    Exit Do
                End If
            Loop
            If ParseAssignLiteral() Then
                .BufPos = p209
                .ThunkPos = q209
            Else
                .BufPos = p198
                .ThunkPos = q198
                Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                ParsePrimaryExpr = True
                Exit Function
            End If
        Else
            .BufPos = p209
            .ThunkPos = q209
        End If
        If ParseOPEN() Then
            If ClearErrors(ctx.BufPos + 1) Then
                p208 = .BufPos
                q208 = .ThunkPos
                Call ParseExpression
                If ParseCLOSE() Then
                    Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                    ParsePrimaryExpr = True
                    Exit Function
                Else
                    .BufPos = p208
                    .ThunkPos = q208
                End If
                If AppendError("Invalid expression", ctx.BufPos + 1) Then
                    Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                    ParsePrimaryExpr = True
                    Exit Function
                End If
                .BufPos = p209
                .ThunkPos = q209
            Else
                .BufPos = p209
                .ThunkPos = q209
            End If
        Else
            .BufPos = p209
            .ThunkPos = q209
        End If
        If ParseLiteralNoSpace() Then
            pvPushThunk ucsAct_3_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p218 = .BufPos
            q218 = .ThunkPos
            If ParseIgnoreCase() Then
                Call ParseSpacing
                pvPushThunk ucsAct_4_PrimaryExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                ParsePrimaryExpr = True
                Exit Function
            Else
                .BufPos = p218
                .ThunkPos = q218
            End If
            Call ParseEMPTY
            Call ParseSpacing
            pvPushThunk ucsAct_4_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
            .BufPos = p209
            .ThunkPos = q209
        Else
            .BufPos = p209
            .ThunkPos = q209
        End If
        If ParseClassNoSpace() Then
            pvPushThunk ucsAct_5_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p227 = .BufPos
            q227 = .ThunkPos
            If ParseIgnoreCase() Then
                Call ParseSpacing
                pvPushThunk ucsAct_6_PrimaryExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                ParsePrimaryExpr = True
                Exit Function
            Else
                .BufPos = p227
                .ThunkPos = q227
            End If
            Call ParseEMPTY
            Call ParseSpacing
            pvPushThunk ucsAct_6_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
            .BufPos = p209
            .ThunkPos = q209
        Else
            .BufPos = p209
            .ThunkPos = q209
        End If
        If ParseDOT() Then
            pvPushThunk ucsAct_7_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p209
            .ThunkPos = q209
        End If
        If ParseCodeBlock() Then
            pvPushThunk ucsAct_8_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p209
            .ThunkPos = q209
        End If
        If ParseBEGIN() Then
            pvPushThunk ucsAct_9_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p209
            .ThunkPos = q209
        End If
        If ParseEND() Then
            pvPushThunk ucsAct_10_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p209
            .ThunkPos = q209
        End If
        .BufPos = p248
        .ThunkPos = q248
        If ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseQUESTION() As Boolean
    Dim p894 As Long
    Dim p887 As Long

    With ctx
        p894 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "QUESTION", 1, ctx.UserData)
        p887 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """?""", 0, ctx.UserData)
        If .BufData(.BufPos) = 63 Then              ' "?"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """?""", 2, ctx.UserData)
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "QUESTION", 2, ctx.UserData)
            Call pvSetAdvance
            ParseQUESTION = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """?""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p894
        If ConsoleTrace(ctx.BufPos + 1, "QUESTION", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSTAR() As Boolean
    Dim p907 As Long
    Dim p900 As Long

    With ctx
        p907 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "STAR", 1, ctx.UserData)
        p900 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """*""", 0, ctx.UserData)
        If .BufData(.BufPos) = 42 Then              ' "*"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """*""", 2, ctx.UserData)
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "STAR", 2, ctx.UserData)
            Call pvSetAdvance
            ParseSTAR = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """*""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p907
        If ConsoleTrace(ctx.BufPos + 1, "STAR", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParsePLUS() As Boolean
    Dim p920 As Long
    Dim p913 As Long

    With ctx
        p920 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "PLUS", 1, ctx.UserData)
        p913 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """+""", 0, ctx.UserData)
        If .BufData(.BufPos) = 43 Then              ' "+"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """+""", 2, ctx.UserData)
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "PLUS", 2, ctx.UserData)
            Call pvSetAdvance
            ParsePLUS = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """+""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p920
        If ConsoleTrace(ctx.BufPos + 1, "PLUS", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseParamNC() As Boolean
    Dim p423 As Long
    Dim p402 As Long
    Dim p408 As Long
    Dim p417 As Long

    With ctx
        p423 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ParamNC", 1, ctx.UserData)
        p402 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """(""", 0, ctx.UserData)
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """(""", 2, ctx.UserData)
            Do
                p408 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^()]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 40, 41                         ' [()]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^()]", 3, ctx.UserData) And False) Then
                        If Not (ParseParamNC()) Then
                            .BufPos = p408
                            .BufPos = p408
                            Exit Do
                        End If
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[^()]", 2, ctx.UserData)
                    Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[^()]", 3, ctx.UserData) And False) Then
                            If Not (ParseParamNC()) Then
                                .BufPos = p408
                                .BufPos = p408
                                Exit Do
                            End If
                        End If
                    End If
                End Select
            Loop
            p417 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, """)""", 0, ctx.UserData)
            If .BufData(.BufPos) = 41 Then          ' ")"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, """)""", 2, ctx.UserData)
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "ParamNC", 2, ctx.UserData)
                Call pvSetAdvance
                ParseParamNC = True
                Exit Function
            End If
            If ConsoleTrace(ctx.BufPos + 1, """)""", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p423
            GoTo L36
        End If
        If ConsoleTrace(ctx.BufPos + 1, """(""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p423
L36:
        If ConsoleTrace(ctx.BufPos + 1, "ParamNC", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseLiteralNC() As Boolean
    Dim p607 As Long
    Dim p603 As Long
    Dim p559 As Long
    Dim p569 As Long
    Dim p567 As Long
    Dim p565 As Long
    Dim p576 As Long
    Dim p583 As Long
    Dim p593 As Long
    Dim p591 As Long
    Dim p589 As Long
    Dim p600 As Long

    With ctx
        p607 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 1, ctx.UserData)
        p603 = .BufPos
        p559 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
            Do
                p569 = .BufPos
                p567 = .BufPos
                p565 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
                If .BufData(.BufPos) = 39 Then      ' "'"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
                    .BufPos = p569
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p567
                If Not (ParseChar()) Then
                    .BufPos = p569
                    Exit Do
                End If
            Loop
            p576 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
            If .BufData(.BufPos) = 39 Then          ' "'"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLiteralNC = True
                Exit Function
            End If
            If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p603
            GoTo L39
        End If
        If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p603
L39:
        p583 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
            Do
                p593 = .BufPos
                p591 = .BufPos
                p589 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
                If .BufData(.BufPos) = 34 Then      ' """
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
                    .BufPos = p593
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p591
                If Not (ParseChar()) Then
                    .BufPos = p593
                    Exit Do
                End If
            Loop
            p600 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
            If .BufData(.BufPos) = 34 Then          ' """
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLiteralNC = True
                Exit Function
            End If
            If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p603
            GoTo L40
        End If
        If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p603
L40:
        .BufPos = p607
        If ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseRuleOptNC() As Boolean
    Dim p454 As Long
    Dim p442 As Long
    Dim p448 As Long

    With ctx
        p454 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "RuleOptNC", 1, ctx.UserData)
        p442 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """-""", 0, ctx.UserData)
        If .BufData(.BufPos) = 45 Then              ' "-"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """-""", 2, ctx.UserData)
            If ParseIdentStart() Then
                Do
                    p448 = .BufPos
                    If Not (ParseIdentCont()) Then
                        .BufPos = p448
                        Exit Do
                    End If
                Loop
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "RuleOptNC", 2, ctx.UserData)
                Call pvSetAdvance
                ParseRuleOptNC = True
                Exit Function
            Else
                .BufPos = p454
            End If
            GoTo L41
        End If
        If ConsoleTrace(ctx.BufPos + 1, """-""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p454
L41:
        If ConsoleTrace(ctx.BufPos + 1, "RuleOptNC", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseOPEN() As Boolean
    Dim p933 As Long
    Dim p926 As Long

    With ctx
        p933 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "OPEN", 1, ctx.UserData)
        p926 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """(""", 0, ctx.UserData)
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """(""", 2, ctx.UserData)
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "OPEN", 2, ctx.UserData)
            Call pvSetAdvance
            ParseOPEN = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """(""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p933
        If ConsoleTrace(ctx.BufPos + 1, "OPEN", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCLOSE() As Boolean
    Dim p946 As Long
    Dim p939 As Long

    With ctx
        p946 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CLOSE", 1, ctx.UserData)
        p939 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """)""", 0, ctx.UserData)
        If .BufData(.BufPos) = 41 Then              ' ")"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """)""", 2, ctx.UserData)
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "CLOSE", 2, ctx.UserData)
            Call pvSetAdvance
            ParseCLOSE = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """)""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p946
        If ConsoleTrace(ctx.BufPos + 1, "CLOSE", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseLiteralNoSpace() As Boolean
    Dim p553 As Long
    Dim p549 As Long
    Dim p502 As Long
    Dim lCaptureBegin As Long
    Dim p515 As Long
    Dim p512 As Long
    Dim p510 As Long
    Dim lCaptureEnd As Long
    Dim p522 As Long
    Dim p528 As Long
    Dim p540 As Long
    Dim p538 As Long
    Dim p536 As Long
    Dim p547 As Long

    With ctx
        p553 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 1, ctx.UserData)
        p549 = .BufPos
        p502 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
            lCaptureBegin = .BufPos
            Do
                p515 = .BufPos
                p512 = .BufPos
                p510 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
                If .BufData(.BufPos) = 39 Then      ' "'"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
                    .BufPos = p515
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p512
                If Not (ParseChar()) Then
                    .BufPos = p515
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            p522 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
            If .BufData(.BufPos) = 39 Then          ' "'"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
                Call ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 2, ctx.UserData)
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call pvSetAdvance
                ParseLiteralNoSpace = True
                Exit Function
            End If
            If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p549
            GoTo L42
        End If
        If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p549
L42:
        p528 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
            lCaptureBegin = .BufPos
            Do
                p540 = .BufPos
                p538 = .BufPos
                p536 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
                If .BufData(.BufPos) = 34 Then      ' """
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
                    .BufPos = p540
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p538
                If Not (ParseChar()) Then
                    .BufPos = p540
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            p547 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
            If .BufData(.BufPos) = 34 Then          ' """
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
                Call ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 2, ctx.UserData)
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call pvSetAdvance
                ParseLiteralNoSpace = True
                Exit Function
            End If
            If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p549
            GoTo L43
        End If
        If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p549
L43:
        .BufPos = p553
        If ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseIgnoreCase() As Boolean
    Dim p621 As Long
    Dim lCaptureBegin As Long
    Dim p614 As Long
    Dim lCaptureEnd As Long

    With ctx
        p621 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "IgnoreCase", 1, ctx.UserData)
        lCaptureBegin = .BufPos
        p614 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[si]", 0, ctx.UserData)
        Select Case .BufData(.BufPos)
        Case 115, 105                               ' [si]
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[si]", 2, ctx.UserData)
            lCaptureEnd = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "IgnoreCase", 2, ctx.UserData)
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            Call pvSetAdvance
            ParseIgnoreCase = True
            Exit Function
        End Select
        If ConsoleTrace(ctx.BufPos + 1, "[si]", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p621
        If ConsoleTrace(ctx.BufPos + 1, "IgnoreCase", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Sub ParseEMPTY()
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "EMPTY", 1, ctx.UserData)
        lCaptureBegin = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """""", 0, ctx.UserData)
        Call ConsoleTrace(ctx.BufPos + 1, """""", 2, ctx.UserData)
        lCaptureEnd = .BufPos
        Call ParseSpacing
        Call ConsoleTrace(ctx.BufPos + 1, "EMPTY", 2, ctx.UserData)
        .CaptureBegin = lCaptureBegin
        .CaptureEnd = lCaptureEnd
    End With
End Sub

Private Function ParseClassNoSpace() As Boolean
    Dim p661 As Long
    Dim p636 As Long
    Dim lCaptureBegin As Long
    Dim p649 As Long
    Dim p646 As Long
    Dim p644 As Long
    Dim lCaptureEnd As Long
    Dim p656 As Long

    With ctx
        p661 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ClassNoSpace", 1, ctx.UserData)
        p636 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """[""", 0, ctx.UserData)
        If .BufData(.BufPos) = 91 Then              ' "["
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """[""", 2, ctx.UserData)
            lCaptureBegin = .BufPos
            Do
                p649 = .BufPos
                p646 = .BufPos
                p644 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, """]""", 0, ctx.UserData)
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, """]""", 2, ctx.UserData)
                    .BufPos = p649
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, """]""", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p646
                If Not (ParseRange()) Then
                    .BufPos = p649
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            p656 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, """]""", 0, ctx.UserData)
            If .BufData(.BufPos) = 93 Then          ' "]"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, """]""", 2, ctx.UserData)
                Call ConsoleTrace(ctx.BufPos + 1, "ClassNoSpace", 2, ctx.UserData)
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call pvSetAdvance
                ParseClassNoSpace = True
                Exit Function
            End If
            If ConsoleTrace(ctx.BufPos + 1, """]""", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p661
            GoTo L45
        End If
        If ConsoleTrace(ctx.BufPos + 1, """[""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p661
L45:
        If ConsoleTrace(ctx.BufPos + 1, "ClassNoSpace", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseDOT() As Boolean
    Dim p959 As Long
    Dim p952 As Long

    With ctx
        p959 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "DOT", 1, ctx.UserData)
        p952 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """.""", 0, ctx.UserData)
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """.""", 2, ctx.UserData)
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "DOT", 2, ctx.UserData)
            Call pvSetAdvance
            ParseDOT = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """.""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p959
        If ConsoleTrace(ctx.BufPos + 1, "DOT", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseBEGIN() As Boolean
    Dim p1111 As Long
    Dim p1104 As Long

    With ctx
        p1111 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "BEGIN", 1, ctx.UserData)
        p1104 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """<""", 0, ctx.UserData)
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """<""", 2, ctx.UserData)
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "BEGIN", 2, ctx.UserData)
            Call pvSetAdvance
            ParseBEGIN = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """<""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p1111
        If ConsoleTrace(ctx.BufPos + 1, "BEGIN", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseEND() As Boolean
    Dim p1124 As Long
    Dim p1117 As Long

    With ctx
        p1124 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "END", 1, ctx.UserData)
        p1117 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """>""", 0, ctx.UserData)
        If .BufData(.BufPos) = 62 Then              ' ">"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """>""", 2, ctx.UserData)
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "END", 2, ctx.UserData)
            Call pvSetAdvance
            ParseEND = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """>""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p1124
        If ConsoleTrace(ctx.BufPos + 1, "END", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSettingName() As Boolean
    Dim p277 As Long
    Dim p271 As Long

    With ctx
        p277 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 1, ctx.UserData)
        If ParseAT() Then
            p271 = .BufPos
            If ParseIdentifier() Then
                Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p271
            End If
            If ParseLiteral() Then
                Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p271
            End If
            Call ParseEMPTY
            Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
            ParseSettingName = True
            Exit Function
            .BufPos = p277
        Else
            .BufPos = p277
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SettingName", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSettingValue() As Boolean
    Dim p289 As Long
    Dim p281 As Long
    Dim p283 As Long

    With ctx
        p289 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 1, ctx.UserData)
        p281 = .BufPos
        If ParseIdentifier() Then
            p283 = .BufPos
            If Not (ParseSEMI()) Then
                .BufPos = p283
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p281
        End If
        If ParseLiteral() Then
            p283 = .BufPos
            If Not (ParseSEMI()) Then
                .BufPos = p283
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p281
        End If
        If ParseCodeBlock() Then
            p283 = .BufPos
            If Not (ParseSEMI()) Then
                .BufPos = p283
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p281
        End If
        .BufPos = p289
        If ConsoleTrace(ctx.BufPos + 1, "SettingValue", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseAT() As Boolean
    Dim p1207 As Long
    Dim p1200 As Long

    With ctx
        p1207 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "AT", 1, ctx.UserData)
        p1200 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """@""", 0, ctx.UserData)
        If .BufData(.BufPos) = 64 Then              ' "@"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """@""", 2, ctx.UserData)
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "AT", 2, ctx.UserData)
            Call pvSetAdvance
            ParseAT = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """@""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p1207
        If ConsoleTrace(ctx.BufPos + 1, "AT", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseEndOfLine() As Boolean
    Dim p1091 As Long
    Dim p1074 As Long
    Dim p1080 As Long
    Dim p1086 As Long

    With ctx
        p1091 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 1, ctx.UserData)
        p1074 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\r\n""", 0, ctx.UserData)
        If .BufData(.BufPos) = 13 And .BufData(.BufPos + 1) = 10 Then ' "\r\n"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """\r\n""", 2, ctx.UserData)
            Call ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 2, ctx.UserData)
            Call pvSetAdvance
            ParseEndOfLine = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\r\n""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        p1080 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\n""", 0, ctx.UserData)
        If .BufData(.BufPos) = 10 Then              ' "\n"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\n""", 2, ctx.UserData)
            Call ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 2, ctx.UserData)
            Call pvSetAdvance
            ParseEndOfLine = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\n""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p1074
        p1086 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\r""", 0, ctx.UserData)
        If .BufData(.BufPos) = 13 Then              ' "\r"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\r""", 2, ctx.UserData)
            Call ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 2, ctx.UserData)
            Call pvSetAdvance
            ParseEndOfLine = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\r""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p1074
        .BufPos = p1091
        If ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCodeNC() As Boolean
    Dim p357 As Long
    Dim p336 As Long
    Dim p342 As Long
    Dim p351 As Long

    With ctx
        p357 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CodeNC", 1, ctx.UserData)
        p336 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """{""", 0, ctx.UserData)
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """{""", 2, ctx.UserData)
            Do
                p342 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^{}]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^{}]", 3, ctx.UserData) And False) Then
                        If Not (ParseCodeNC()) Then
                            .BufPos = p342
                            .BufPos = p342
                            Exit Do
                        End If
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[^{}]", 2, ctx.UserData)
                    Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[^{}]", 3, ctx.UserData) And False) Then
                            If Not (ParseCodeNC()) Then
                                .BufPos = p342
                                .BufPos = p342
                                Exit Do
                            End If
                        End If
                    End If
                End Select
            Loop
            p351 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, """}""", 0, ctx.UserData)
            If .BufData(.BufPos) = 125 Then         ' "}"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, """}""", 2, ctx.UserData)
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "CodeNC", 2, ctx.UserData)
                Call pvSetAdvance
                ParseCodeNC = True
                Exit Function
            End If
            If ConsoleTrace(ctx.BufPos + 1, """}""", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p357
            GoTo L51
        End If
        If ConsoleTrace(ctx.BufPos + 1, """{""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p357
L51:
        If ConsoleTrace(ctx.BufPos + 1, "CodeNC", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseIdentStart() As Boolean
    Dim p477 As Long
    Dim p472 As Long

    With ctx
        p477 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "IdentStart", 1, ctx.UserData)
        p472 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[a-zA-Z_\-]", 0, ctx.UserData)
        Select Case .BufData(.BufPos)
        Case 97 To 122, 65 To 90, 95, 92, 45        ' [a-zA-Z_\-]
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[a-zA-Z_\-]", 2, ctx.UserData)
            Call ConsoleTrace(ctx.BufPos + 1, "IdentStart", 2, ctx.UserData)
            Call pvSetAdvance
            ParseIdentStart = True
            Exit Function
        End Select
        If ConsoleTrace(ctx.BufPos + 1, "[a-zA-Z_\-]", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p477
        If ConsoleTrace(ctx.BufPos + 1, "IdentStart", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseIdentCont() As Boolean
    Dim p488 As Long
    Dim p483 As Long

    With ctx
        p488 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "IdentCont", 1, ctx.UserData)
        p483 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[a-zA-Z_0-9\-]", 0, ctx.UserData)
        Select Case .BufData(.BufPos)
        Case 97 To 122, 65 To 90, 95, 48 To 57, 92, 45 ' [a-zA-Z_0-9\-]
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[a-zA-Z_0-9\-]", 2, ctx.UserData)
            Call ConsoleTrace(ctx.BufPos + 1, "IdentCont", 2, ctx.UserData)
            Call pvSetAdvance
            ParseIdentCont = True
            Exit Function
        End Select
        If ConsoleTrace(ctx.BufPos + 1, "[a-zA-Z_0-9\-]", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p488
        If ConsoleTrace(ctx.BufPos + 1, "IdentCont", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseChar() As Boolean
    Dim p820 As Long
    Dim p724 As Long
    Dim p690 As Long
    Dim p696 As Long
    Dim p703 As Long
    Dim p709 As Long
    Dim p716 As Long
    Dim p722 As Long
    Dim p729 As Long
    Dim p735 As Long
    Dim p742 As Long
    Dim p749 As Long
    Dim p755 As Long
    Dim p762 As Long
    Dim p768 As Long
    Dim p774 As Long
    Dim p781 As Long
    Dim p787 As Long
    Dim p793 As Long
    Dim p799 As Long
    Dim p805 As Long
    Dim p814 As Long
    Dim p812 As Long

    With ctx
        p820 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Char", 1, ctx.UserData)
        p724 = .BufPos
        p690 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            p696 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[abefnrtv'""\[\]\\]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 97, 98, 101, 102, 110, 114, 116, 118, 39, 34, 91, 93, 92 ' [abefnrtv'"\[\]\\]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[abefnrtv'""\[\]\\]", 2, ctx.UserData)
                Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                Call pvSetAdvance
                ParseChar = True
                Exit Function
            End Select
            If ConsoleTrace(ctx.BufPos + 1, "[abefnrtv'""\[\]\\]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p724
            GoTo L54
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p724
L54:
        p703 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            p709 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[0-3]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 48 To 51                           ' [0-3]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[0-3]", 2, ctx.UserData)
                p716 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 2, ctx.UserData)
                    p722 = .BufPos
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 0, ctx.UserData)
                    Select Case .BufData(.BufPos)
                    Case 48 To 55                   ' [0-7]
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 2, ctx.UserData)
                        Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                        Call pvSetAdvance
                        ParseChar = True
                        Exit Function
                    End Select
                    If ConsoleTrace(ctx.BufPos + 1, "[0-7]", 3, ctx.UserData) And False Then
                        '--- do nothing
                    End If
                    .BufPos = p724
                    GoTo L60
                End Select
                If ConsoleTrace(ctx.BufPos + 1, "[0-7]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p724
L60:
                GoTo L61
            End Select
            If ConsoleTrace(ctx.BufPos + 1, "[0-3]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p724
L61:
            GoTo L59
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p724
L59:
        p729 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            p735 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 48 To 55                           ' [0-7]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 2, ctx.UserData)
                p742 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 2, ctx.UserData)
                    GoTo L68
                End Select
                If ConsoleTrace(ctx.BufPos + 1, "[0-7]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p742
L68:
                Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                Call pvSetAdvance
                ParseChar = True
                Exit Function
            End Select
            If ConsoleTrace(ctx.BufPos + 1, "[0-7]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p724
            GoTo L67
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p724
L67:
        p749 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\x""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 120 Then ' "\\x"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """\\x""", 2, ctx.UserData)
            p755 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                p762 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 48 To 57, 97 To 102, 65 To 70  ' [0-9a-fA-F]
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                    Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                    Call pvSetAdvance
                    ParseChar = True
                    Exit Function
                End Select
                If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p724
                GoTo L72
            End Select
            If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p724
L72:
            GoTo L71
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\x""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p724
L71:
        p768 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\u""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 117 Then ' "\\u"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """\\u""", 2, ctx.UserData)
            p774 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                p781 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 48 To 57, 97 To 102, 65 To 70  ' [0-9a-fA-F]
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                    p787 = .BufPos
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
                    Select Case .BufData(.BufPos)
                    Case 48 To 57, 97 To 102, 65 To 70 ' [0-9a-fA-F]
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                        p793 = .BufPos
                        Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
                        Select Case .BufData(.BufPos)
                        Case 48 To 57, 97 To 102, 65 To 70 ' [0-9a-fA-F]
                            .BufPos = .BufPos + 1
                            Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                            Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                            Call pvSetAdvance
                            ParseChar = True
                            Exit Function
                        End Select
                        If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                            '--- do nothing
                        End If
                        .BufPos = p724
                        GoTo L84
                    End Select
                    If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                        '--- do nothing
                    End If
                    .BufPos = p724
L84:
                    GoTo L85
                End Select
                If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p724
L85:
                GoTo L86
            End Select
            If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p724
L86:
            GoTo L83
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\u""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p724
L83:
        p799 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            p805 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, """-""", 0, ctx.UserData)
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, """-""", 2, ctx.UserData)
                Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                Call pvSetAdvance
                ParseChar = True
                Exit Function
            End If
            If ConsoleTrace(ctx.BufPos + 1, """-""", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p724
            GoTo L90
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p724
L90:
        p814 = .BufPos
        p812 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            .BufPos = p724
            GoTo L91
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p814
        If .BufPos < .BufSize Then
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
            Call pvSetAdvance
            ParseChar = True
            Exit Function
        Else
            .BufPos = p724
        End If
L91:
        .BufPos = p820
        If ConsoleTrace(ctx.BufPos + 1, "Char", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        Call pvSetExpected("char")
    End With
End Function

Private Function ParseClass() As Boolean
    Dim p630 As Long

    With ctx
        p630 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Class", 1, ctx.UserData)
        If ParseClassNoSpace() Then
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "Class", 2, ctx.UserData)
            ParseClass = True
            Exit Function
        Else
            .BufPos = p630
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Class", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseRange() As Boolean
    Dim p684 As Long
    Dim p680 As Long
    Dim p668 As Long
    Dim p677 As Long
    Dim p675 As Long

    With ctx
        p684 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Range", 1, ctx.UserData)
        p680 = .BufPos
        If ParseChar() Then
            p668 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, """-""", 0, ctx.UserData)
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, """-""", 2, ctx.UserData)
                p677 = .BufPos
                p675 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, """]""", 0, ctx.UserData)
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, """]""", 2, ctx.UserData)
                    .BufPos = p680
                    GoTo L94
                End If
                If ConsoleTrace(ctx.BufPos + 1, """]""", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p677
                If ParseChar() Then
                    Call ConsoleTrace(ctx.BufPos + 1, "Range", 2, ctx.UserData)
                    Call pvSetAdvance
                    ParseRange = True
                    Exit Function
                Else
                    .BufPos = p680
                End If
L94:
                GoTo L93
            End If
            If ConsoleTrace(ctx.BufPos + 1, """-""", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p680
L93:
        Else
            .BufPos = p680
        End If
        If ParseChar() Then
            Call ConsoleTrace(ctx.BufPos + 1, "Range", 2, ctx.UserData)
            Call pvSetAdvance
            ParseRange = True
            Exit Function
        Else
            .BufPos = p680
        End If
        .BufPos = p684
        If ConsoleTrace(ctx.BufPos + 1, "Range", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseLEFTARROW() As Boolean
    Dim p834 As Long
    Dim p827 As Long

    With ctx
        p834 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LEFTARROW", 1, ctx.UserData)
        p827 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """<-""", 0, ctx.UserData)
        If .BufData(.BufPos) = 60 And .BufData(.BufPos + 1) = 45 Then ' "<-"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """<-""", 2, ctx.UserData)
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "LEFTARROW", 2, ctx.UserData)
            Call pvSetAdvance
            ParseLEFTARROW = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """<-""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p834
        If ConsoleTrace(ctx.BufPos + 1, "LEFTARROW", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSLASH() As Boolean
    Dim p855 As Long
    Dim p841 As Long
    Dim p849 As Long
    Dim p847 As Long

    With ctx
        p855 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SLASH", 1, ctx.UserData)
        p841 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """/""", 0, ctx.UserData)
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """/""", 2, ctx.UserData)
            p849 = .BufPos
            p847 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, """/""", 0, ctx.UserData)
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, """/""", 2, ctx.UserData)
                .BufPos = p855
                GoTo L98
            End If
            If ConsoleTrace(ctx.BufPos + 1, """/""", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p849
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "SLASH", 2, ctx.UserData)
            Call pvSetAdvance
            ParseSLASH = True
            Exit Function
L98:
            GoTo L97
        End If
        If ConsoleTrace(ctx.BufPos + 1, """/""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p855
L97:
        If ConsoleTrace(ctx.BufPos + 1, "SLASH", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSpace() As Boolean
    Dim p1068 As Long
    Dim p1056 As Long
    Dim p1062 As Long

    With ctx
        p1068 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Space", 1, ctx.UserData)
        p1056 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """ """, 0, ctx.UserData)
        If .BufData(.BufPos) = 32 Then              ' " "
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """ """, 2, ctx.UserData)
            Call ConsoleTrace(ctx.BufPos + 1, "Space", 2, ctx.UserData)
            Call pvSetAdvance
            ParseSpace = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """ """, 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        p1062 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\t""", 0, ctx.UserData)
        If .BufData(.BufPos) = 9 Then               ' "\t"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\t""", 2, ctx.UserData)
            Call ConsoleTrace(ctx.BufPos + 1, "Space", 2, ctx.UserData)
            Call pvSetAdvance
            ParseSpace = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\t""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p1056
        If ParseEndOfLine() Then
            Call ConsoleTrace(ctx.BufPos + 1, "Space", 2, ctx.UserData)
            Call pvSetAdvance
            ParseSpace = True
            Exit Function
        Else
            .BufPos = p1056
        End If
        .BufPos = p1068
        If ConsoleTrace(ctx.BufPos + 1, "Space", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseLineComment() As Boolean
    Dim p1015 As Long
    Dim p993 As Long
    Dim p1005 As Long
    Dim p1011 As Long
    Dim p999 As Long

    With ctx
        p1015 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 1, ctx.UserData)
        p993 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """#""", 0, ctx.UserData)
        If .BufData(.BufPos) = 35 Then              ' "#"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """#""", 2, ctx.UserData)
            Do
                p1005 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 13, 10                         ' [\r\n]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 3, ctx.UserData) And False) Then
                        .BufPos = p1005
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 2, ctx.UserData)
                    Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 3, ctx.UserData) And False) Then
                            .BufPos = p1005
                            Exit Do
                        End If
                    End If
                End Select
            Loop
            p1011 = .BufPos
            If ParseEndOfLine() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p1011
            End If
            If ParseEndOfFile() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p1011
            End If
            .BufPos = p1015
            GoTo L101
        End If
        If ConsoleTrace(ctx.BufPos + 1, """#""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        p999 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """//""", 0, ctx.UserData)
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 47 Then ' "//"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """//""", 2, ctx.UserData)
            Do
                p1005 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 13, 10                         ' [\r\n]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 3, ctx.UserData) And False) Then
                        .BufPos = p1005
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 2, ctx.UserData)
                    Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 3, ctx.UserData) And False) Then
                            .BufPos = p1005
                            Exit Do
                        End If
                    End If
                End Select
            Loop
            p1011 = .BufPos
            If ParseEndOfLine() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p1011
            End If
            If ParseEndOfFile() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p1011
            End If
            .BufPos = p1015
            GoTo L101
        End If
        If ConsoleTrace(ctx.BufPos + 1, """//""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p993
        .BufPos = p1015
L101:
        If ConsoleTrace(ctx.BufPos + 1, "LineComment", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseBlockComment() As Boolean
    Dim p1050 As Long
    Dim p1021 As Long
    Dim p1027 As Long
    Dim p1036 As Long
    Dim p1034 As Long
    Dim p1045 As Long

    With ctx
        p1050 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "BlockComment", 1, ctx.UserData)
        p1021 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """/*""", 0, ctx.UserData)
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 42 Then ' "/*"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """/*""", 2, ctx.UserData)
            Do
                p1027 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^*/]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 42, 47                         ' [*/]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^*/]", 3, ctx.UserData) And False) Then
                        If Not (ParseBlockComment()) Then
                            .BufPos = p1027
                            p1036 = .BufPos
                            p1034 = .BufPos
                            Call ConsoleTrace(ctx.BufPos + 1, """*/""", 0, ctx.UserData)
                            If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                                .BufPos = .BufPos + 2
                                Call ConsoleTrace(ctx.BufPos + 1, """*/""", 2, ctx.UserData)
                                .BufPos = p1027
                                .BufPos = p1027
                                Exit Do
                            End If
                            If ConsoleTrace(ctx.BufPos + 1, """*/""", 3, ctx.UserData) And False Then
                                '--- do nothing
                            End If
                            .BufPos = p1036
                            If .BufPos < .BufSize Then
                                .BufPos = .BufPos + 1
                            Else
                                .BufPos = p1027
                                .BufPos = p1027
                                Exit Do
                            End If
                        End If
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[^*/]", 2, ctx.UserData)
                    Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[^*/]", 3, ctx.UserData) And False) Then
                            If Not (ParseBlockComment()) Then
                                .BufPos = p1027
                                p1036 = .BufPos
                                p1034 = .BufPos
                                Call ConsoleTrace(ctx.BufPos + 1, """*/""", 0, ctx.UserData)
                                If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                                    .BufPos = .BufPos + 2
                                    Call ConsoleTrace(ctx.BufPos + 1, """*/""", 2, ctx.UserData)
                                    .BufPos = p1027
                                    .BufPos = p1027
                                    Exit Do
                                End If
                                If ConsoleTrace(ctx.BufPos + 1, """*/""", 3, ctx.UserData) And False Then
                                    '--- do nothing
                                End If
                                .BufPos = p1036
                                If .BufPos < .BufSize Then
                                    .BufPos = .BufPos + 1
                                Else
                                    .BufPos = p1027
                                    .BufPos = p1027
                                    Exit Do
                                End If
                            End If
                        End If
                    End If
                End Select
            Loop
            p1045 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, """*/""", 0, ctx.UserData)
            If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                .BufPos = .BufPos + 2
                Call ConsoleTrace(ctx.BufPos + 1, """*/""", 2, ctx.UserData)
                Call ConsoleTrace(ctx.BufPos + 1, "BlockComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseBlockComment = True
                Exit Function
            End If
            If ConsoleTrace(ctx.BufPos + 1, """*/""", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p1050
            GoTo L108
        End If
        If ConsoleTrace(ctx.BufPos + 1, """/*""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p1050
L108:
        If ConsoleTrace(ctx.BufPos + 1, "BlockComment", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseEQUAL() As Boolean
    Dim p1155 As Long
    Dim p1148 As Long

    With ctx
        p1155 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "EQUAL", 1, ctx.UserData)
        p1148 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """=""", 0, ctx.UserData)
        If .BufData(.BufPos) = 61 Then              ' "="
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """=""", 2, ctx.UserData)
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "EQUAL", 2, ctx.UserData)
            Call pvSetAdvance
            ParseEQUAL = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """=""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p1155
        If ConsoleTrace(ctx.BufPos + 1, "EQUAL", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseBAR() As Boolean
    Dim p1168 As Long
    Dim p1161 As Long

    With ctx
        p1168 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "BAR", 1, ctx.UserData)
        p1161 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """|""", 0, ctx.UserData)
        If .BufData(.BufPos) = 124 Then             ' "|"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """|""", 2, ctx.UserData)
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "BAR", 2, ctx.UserData)
            Call pvSetAdvance
            ParseBAR = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, """|""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p1168
        If ConsoleTrace(ctx.BufPos + 1, "BAR", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParsePrepExpr() As Boolean
    Dim p1255 As Long
    Dim p1248 As Long

    With ctx
        p1255 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 1, ctx.UserData)
        p1248 = .BufPos
        If ParseIdentifier() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1248
        End If
        If ParseLiteral() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1248
        End If
        If ParseClass() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1248
        End If
        If ParseCodeNC() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1248
        End If
        If ParseParamBlock() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1248
        End If
        .BufPos = p1255
        If ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Sub pvImplAction(ByVal eAction As UcsParserActionsEnum, ByVal lOffset As Long, ByVal lSize As Long)
    Dim lNode           As Long
    
    With ctx.UserData
    Select Case eAction
    Case ucsAct_2_Setting
         .SetSettingValue .Pop, Mid$(ctx.Contents, lOffset, lSize)
    Case ucsAct_1_Setting
         .Push .MakeString(lOffset, Mid$(ctx.Contents, lOffset, lSize))
    Case ucsAct_5_Rule
         lNode = .Pop: .SetExpression .Pop, lNode
    Case ucsAct_4_Rule
         .RuleOption(.Top, Mid$(ctx.Contents, lOffset, lSize)) = True
    Case ucsAct_3_Rule
         .RuleDescription(.Top) = Mid$(ctx.Contents, lOffset, lSize)
    Case ucsAct_2_Rule
         .RuleParams(.Top) = Mid$(ctx.Contents, lOffset, lSize)
    Case ucsAct_1_Rule
           lNode = .BeginRule(.FindRule(lOffset, Mid$(ctx.Contents, lOffset, lSize)))
                                                                            If .RuleExpression(lNode) <> 0 Then
                                                                                AppendError "Rule '" & Mid$(ctx.Contents, lOffset, lSize) & "' redefined", lOffset
                                                                            End If
                                                                            .Push lNode
    Case ucsAct_1_ChoiceExpr
         lNode = .Pop: .Push .AppendChild(.MakeChoice(lOffset, .Pop), lNode)
    Case ucsAct_2_SequenceExpr
         .Push .MakePredicate(lOffset, "", AlwaysTrue:=True)
    Case ucsAct_1_SequenceExpr
         lNode = .Pop: .Push .AppendSequence(.MakeSequence(lOffset, .Pop), lNode)
    Case ucsAct_1_ErrorExpr
         .Push .AppendChild(.MakeChoice(lOffset, .Pop), .MakeErrorReference(lOffset, Mid$(ctx.Contents, lOffset, lSize)))
    Case ucsAct_2_ErrorExpr
         .Push .AppendChild(.MakeChoice(lOffset, .Pop), .MakeErrorString(lOffset, Mid$(ctx.Contents, lOffset, lSize)))
    Case ucsAct_2_VariableExpr
         .Push .SetVariable(lOffset, .Pop, .Pop)
    Case ucsAct_1_VariableExpr
         .Push .MakeVariable(lOffset, Mid$(ctx.Contents, lOffset, lSize))
    Case ucsAct_1_PrefixExpr
         .Push .MakePredicate(lOffset, Mid$(ctx.Contents, lOffset, lSize))
    Case ucsAct_2_PrefixExpr
         .Push .MakePeekFor(lOffset, .Pop)
    Case ucsAct_3_PrefixExpr
         .Push .MakePeekNot(lOffset, .Pop)
    Case ucsAct_1_SuffixExpr
         .Push .MakeQuery(lOffset, .Pop)
    Case ucsAct_2_SuffixExpr
         .Push .MakeStar(lOffset, .Pop)
    Case ucsAct_3_SuffixExpr
         .Push .MakePlus(lOffset, .Pop)
    Case ucsAct_2_PrimaryExpr
         .NodeParams(.Top) = Mid$(ctx.Contents, lOffset, lSize)
    Case ucsAct_1_PrimaryExpr
         .Push .MakeReference(lOffset, Mid$(ctx.Contents, lOffset, lSize))
    Case ucsAct_4_PrimaryExpr
         .Push .SetIgnoreCase(.Pop, Mid$(ctx.Contents, lOffset, lSize))
    Case ucsAct_3_PrimaryExpr
         .Push .MakeString(lOffset, Mid$(ctx.Contents, lOffset, lSize))
    Case ucsAct_6_PrimaryExpr
         .Push .SetIgnoreCase(.Pop, Mid$(ctx.Contents, lOffset, lSize))
    Case ucsAct_5_PrimaryExpr
         .Push .MakeClass(lOffset, Mid$(ctx.Contents, lOffset, lSize))
    Case ucsAct_7_PrimaryExpr
         .Push .MakeDot(lOffset)
    Case ucsAct_8_PrimaryExpr
         .Push .MakeAction(lOffset, Mid$(ctx.Contents, lOffset, lSize))
    Case ucsAct_9_PrimaryExpr
         .Push .MakePredicate(lOffset, "YY_BEGIN", AlwaysTrue:=True)
    Case ucsAct_10_PrimaryExpr
         .Push .MakePredicate(lOffset, "YY_END", AlwaysTrue:=True)
    End Select
    End With
End Sub

'= user-defined functions ================================================

Private Function ClearErrors(ByVal lOffset As Long) As Boolean
    Dim lIdx            As Long
    
    If Not ctx.ParseErrors Is Nothing Then
        For lIdx = ctx.ParseErrors.Count To 1 Step -1
            If LenB(ctx.ParseErrors.Item(lIdx)(2)) = 0 Then
                If ctx.ParseErrors.Item(lIdx)(1) <= lOffset Then
                    ctx.ParseErrors.Remove lIdx
                End If
            End If
        Next
        If ctx.ParseErrors.Count = 0 Then
            Set ctx.ParseErrors = Nothing
        End If
    End If
    ClearErrors = True
End Function

Private Function AppendError(sText As String, ByVal lOffset As Long, Optional sType As String) As Boolean
    If ctx.ParseErrors Is Nothing Then
        Set ctx.ParseErrors = New Collection
    End If
    ctx.ParseErrors.Add Array(sText, lOffset, sType)
End Function

Private Function SetUserDataTag(sText As String) As Boolean
    ctx.UserData.Tag = sText
    SetUserDataTag = True
End Function

Public Function GetParseErrors() As Collection
    Set GetParseErrors = ctx.ParseErrors
End Function

Private Function CheckSettingName(sName As String) As Boolean
    Const STR_ALL = "||private|public|modulename|userdatatype|vardatatype|members|start|ignorecase|prolog|epilog|trace|append|declares|"
    If InStr(1, STR_ALL, "|" & sName & "|", vbTextCompare) = 0 Then
        AppendError "Setting '" & sName & "' is unrecognized", ctx.CaptureBegin + 1, "warning"
    End If
    CheckSettingName = True
End Function

Public Sub AddParamRule(sRule As String)
    If ctx.ParamRules Is Nothing Then
        Set ctx.ParamRules = New Collection
    End If
    If Not SearchCollection(ctx.ParamRules, sRule) Then
        ctx.ParamRules.Add sRule, sRule
    End If
End Sub

Private Function IsParamRule(sRule As String) As Boolean
    IsParamRule = SearchCollection(ctx.ParamRules, sRule)
End Function

Private Function CheckRuleOptionName(sName As String) As Boolean
    Const STR_ALL = "|public|"
    If InStr(1, STR_ALL, "|" & sName & "|", vbTextCompare) = 0 Then
        AppendError "Rule option '" & sName & "' is unrecognized", ctx.CaptureBegin + 1, "warning"
    End If
    CheckRuleOptionName = True
End Function
