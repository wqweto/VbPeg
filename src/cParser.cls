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
' Auto-generated on 3.12.2018 1:29:09
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

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
    ParserVersion = "3.12.2018 1:29:09"
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
    Dim p25 As Long
    Dim q25 As Long
    Dim p6 As Long
    Dim i17 As Long
    Dim p16 As Long
    Dim q16 As Long
    Dim p21 As Long
    Dim q21 As Long

    With ctx
        p25 = .BufPos
        q25 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Grammar", 1, ctx.UserData)
        p6 = .BufPos
        If Not (ParsePreprocess()) Then
            .BufPos = p6
        End If
        '-- rewind position
        ctx.BufPos = 0
        Call ParseSpacing
        For i17 = 0 To LNG_MAXINT
            p16 = .BufPos
            q16 = .ThunkPos
            If Not (ParseSetting()) Then
                .BufPos = p16
                .ThunkPos = q16
                If Not (ParseRule()) Then
                    .BufPos = p16
                    .ThunkPos = q16
                    .BufPos = p16
                    .ThunkPos = q16
                    Exit For
                End If
            End If
        Next
        If i17 <> 0 Then
            p21 = .BufPos
            q21 = .ThunkPos
            If ParseEndOfFile() Then
                Call ConsoleTrace(ctx.BufPos + 1, "Grammar", 2, ctx.UserData)
                ParseGrammar = True
                Exit Function
            Else
                .BufPos = p21
                .ThunkPos = q21
            End If
            If AppendError("Expected rule or setting", ctx.BufPos + 1) Then
                Call ConsoleTrace(ctx.BufPos + 1, "Grammar", 2, ctx.UserData)
                ParseGrammar = True
                Exit Function
            End If
            .BufPos = p25
            .ThunkPos = q25
        Else
            .BufPos = p25
            .ThunkPos = q25
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Grammar", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParsePreprocess() As Boolean
    Dim p1234 As Long
    Dim i1229 As Long
    Dim p1225 As Long
    Dim p1217 As Long
    Dim p1219 As Long

    With ctx
        p1234 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Preprocess", 1, ctx.UserData)
        Call ParseSpacing
        For i1229 = 0 To LNG_MAXINT
            p1225 = .BufPos
            If ParseIdentifier() Then
                If ParseParamNC() Then
                    p1217 = .BufPos
                    If Not (ParseLiteralNC()) Then
                        .BufPos = p1217
                    End If
                    Do
                        p1219 = .BufPos
                        If Not (ParseRuleOptNC()) Then
                            .BufPos = p1219
                            Exit Do
                        End If
                    Loop
                    If ParseAssignLiteral() Then
                        '--- this might be premature
                        AddParamRule(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin))
                    Else
                        .BufPos = p1225
                        If Not (ParsePrepExpr()) Then
                            .BufPos = p1225
                            If .BufPos < .BufSize Then
                                .BufPos = .BufPos + 1
                                Call ParseSpacing
                            Else
                                .BufPos = p1225
                                Exit For
                            End If
                        End If
                    End If
                Else
                    .BufPos = p1225
                    If Not (ParsePrepExpr()) Then
                        .BufPos = p1225
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                            Call ParseSpacing
                        Else
                            .BufPos = p1225
                            Exit For
                        End If
                    End If
                End If
            Else
                .BufPos = p1225
                If Not (ParsePrepExpr()) Then
                    .BufPos = p1225
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        Call ParseSpacing
                    Else
                        .BufPos = p1225
                        Exit For
                    End If
                End If
            End If
        Next
        If i1229 <> 0 Then
            Call ConsoleTrace(ctx.BufPos + 1, "Preprocess", 2, ctx.UserData)
            ParsePreprocess = True
            Exit Function
        Else
            .BufPos = p1234
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Preprocess", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Sub ParseSpacing()
    Dim p971 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "Spacing", 1, ctx.UserData)
        Do
            p971 = .BufPos
            If Not (ParseSpace()) Then
                .BufPos = p971
                If Not (ParseLineComment()) Then
                    .BufPos = p971
                    If Not (ParseBlockComment()) Then
                        .BufPos = p971
                        .BufPos = p971
                        Exit Do
                    End If
                End If
            End If
        Loop
        Call ConsoleTrace(ctx.BufPos + 1, "Spacing", 2, ctx.UserData)
    End With
End Sub

Private Function ParseSetting() As Boolean
    Dim p266 As Long
    Dim q266 As Long
    Dim p262 As Long
    Dim q262 As Long

    With ctx
        p266 = .BufPos
        q266 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 1, ctx.UserData)
        If ParseSettingName() Then
            If CheckSettingName(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                pvPushThunk ucsAct_1_Setting, .CaptureBegin, .CaptureEnd
                If ClearErrors(ctx.BufPos + 1) And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                    p262 = .BufPos
                    q262 = .ThunkPos
                    If ParseSettingValue() Then
                        pvPushThunk ucsAct_2_Setting, .CaptureBegin, .CaptureEnd
                        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 2, ctx.UserData)
                        ParseSetting = True
                        Exit Function
                    Else
                        .BufPos = p262
                        .ThunkPos = q262
                    End If
                    If AppendError("Invalid value for setting '" & .UserData.Tag & "'", ctx.BufPos + 1) Then
                        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 2, ctx.UserData)
                        ParseSetting = True
                        Exit Function
                    End If
                    .BufPos = p266
                    .ThunkPos = q266
                Else
                    .BufPos = p266
                    .ThunkPos = q266
                End If
            Else
                .BufPos = p266
                .ThunkPos = q266
            End If
        Else
            .BufPos = p266
            .ThunkPos = q266
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Setting", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseRule() As Boolean
    Dim p62 As Long
    Dim q62 As Long
    Dim p35 As Long
    Dim q35 As Long
    Dim p40 As Long
    Dim q40 As Long
    Dim p45 As Long
    Dim q45 As Long
    Dim p57 As Long
    Dim q57 As Long

    With ctx
        p62 = .BufPos
        q62 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Rule", 1, ctx.UserData)
        If ParseIdentifier() Then
            pvPushThunk ucsAct_1_Rule, .CaptureBegin, .CaptureEnd
            If ClearErrors(ctx.BufPos + 1) And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                p35 = .BufPos
                q35 = .ThunkPos
                If ParseParamBlock() Then
                    pvPushThunk ucsAct_2_Rule, .CaptureBegin, .CaptureEnd
                Else
                    .BufPos = p35
                    .ThunkPos = q35
                End If
                p40 = .BufPos
                q40 = .ThunkPos
                If ParseLiteral() Then
                    pvPushThunk ucsAct_3_Rule, .CaptureBegin, .CaptureEnd
                Else
                    .BufPos = p40
                    .ThunkPos = q40
                End If
                Do
                    p45 = .BufPos
                    q45 = .ThunkPos
                    If Not (ParseRuleOption()) Then
                        .BufPos = p45
                        .ThunkPos = q45
                        Exit Do
                    End If
                    If Not (CheckRuleOptionName(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin))) Then
                        .BufPos = p45
                        .ThunkPos = q45
                        Exit Do
                    End If
                    pvPushThunk ucsAct_4_Rule, .CaptureBegin, .CaptureEnd
                Loop
                If ParseAssignLiteral() Then
                    Call ParseExpression
                    pvPushThunk ucsAct_5_Rule, .CaptureBegin, .CaptureEnd
                    p57 = .BufPos
                    q57 = .ThunkPos
                    If Not (ParseSEMI()) Then
                        .BufPos = p57
                        .ThunkPos = q57
                    End If
                    Call ConsoleTrace(ctx.BufPos + 1, "Rule", 2, ctx.UserData)
                    ParseRule = True
                    Exit Function
                Else
                    .BufPos = p62
                    .ThunkPos = q62
                End If
            Else
                .BufPos = p62
                .ThunkPos = q62
            End If
        Else
            .BufPos = p62
            .ThunkPos = q62
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Rule", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseEndOfFile() As Boolean
    Dim p1089 As Long
    Dim p1085 As Long

    With ctx
        p1089 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "EndOfFile", 1, ctx.UserData)
        p1085 = .BufPos
        If .BufPos < .BufSize Then
            .BufPos = p1089
        Else
            .BufPos = p1085
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
    Dim p457 As Long
    Dim lCaptureBegin As Long
    Dim p450 As Long
    Dim lCaptureEnd As Long

    With ctx
        p457 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Identifier", 1, ctx.UserData)
        lCaptureBegin = .BufPos
        If ParseIdentStart() Then
            Do
                p450 = .BufPos
                If Not (ParseIdentCont()) Then
                    .BufPos = p450
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
            .BufPos = p457
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Identifier", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseParamBlock() As Boolean
    Dim p387 As Long
    Dim p364 As Long
    Dim lCaptureBegin As Long
    Dim p372 As Long
    Dim lCaptureEnd As Long
    Dim p381 As Long

    With ctx
        p387 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ParamBlock", 1, ctx.UserData)
        p364 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """(""", 0, ctx.UserData)
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """(""", 2, ctx.UserData)
            lCaptureBegin = .BufPos
            Do
                p372 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^()]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 40, 41                         ' [()]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^()]", 3, ctx.UserData) And False) Then
                        If Not (ParseParamNC()) Then
                            .BufPos = p372
                            .BufPos = p372
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
                                .BufPos = p372
                                .BufPos = p372
                                Exit Do
                            End If
                        End If
                    End If
                End Select
            Loop
            lCaptureEnd = .BufPos
            p381 = .BufPos
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
            .BufPos = p387
            GoTo L11
        End If
        If ConsoleTrace(ctx.BufPos + 1, """(""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p387
L11:
        If ConsoleTrace(ctx.BufPos + 1, "ParamBlock", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseLiteral() As Boolean
    Dim p487 As Long

    With ctx
        p487 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Literal", 1, ctx.UserData)
        If ParseLiteralNoSpace() Then
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "Literal", 2, ctx.UserData)
            ParseLiteral = True
            Exit Function
        Else
            .BufPos = p487
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Literal", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseRuleOption() As Boolean
    Dim p427 As Long
    Dim p420 As Long

    With ctx
        p427 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "RuleOption", 1, ctx.UserData)
        p420 = .BufPos
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
                .BufPos = p427
            End If
            GoTo L14
        End If
        If ConsoleTrace(ctx.BufPos + 1, """-""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p427
L14:
        If ConsoleTrace(ctx.BufPos + 1, "RuleOption", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseAssignLiteral() As Boolean
    Dim p1124 As Long
    Dim p1120 As Long

    With ctx
        p1124 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 1, ctx.UserData)
        p1120 = .BufPos
        If ParseLEFTARROW() Then
            Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 2, ctx.UserData)
            ParseAssignLiteral = True
            Exit Function
        Else
            .BufPos = p1120
        End If
        If ParseEQUAL() Then
            Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 2, ctx.UserData)
            ParseAssignLiteral = True
            Exit Function
        Else
            .BufPos = p1120
        End If
        .BufPos = p1124
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
L15:
    End With
End Sub

Private Function ParseSEMI() As Boolean
    Dim p1185 As Long
    Dim p1178 As Long

    With ctx
        p1185 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SEMI", 1, ctx.UserData)
        p1178 = .BufPos
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
        .BufPos = p1185
        If ConsoleTrace(ctx.BufPos + 1, "SEMI", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Sub ParseChoiceExpr()
    Dim p76 As Long
    Dim q76 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceExpr", 1, ctx.UserData)
        Call ParseSequenceExpr
        Do
            p76 = .BufPos
            q76 = .ThunkPos
            If Not (ParseChoiceLiteral()) Then
                .BufPos = p76
                .ThunkPos = q76
                Exit Do
            End If
            Call ParseSequenceExpr
            pvPushThunk ucsAct_1_ChoiceExpr, .CaptureBegin, .CaptureEnd
        Loop
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceExpr", 2, ctx.UserData)
L16:
    End With
End Sub

Private Sub ParseSequenceExpr()
    Dim p93 As Long
    Dim q93 As Long
    Dim p89 As Long
    Dim q89 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "SequenceExpr", 1, ctx.UserData)
        p93 = .BufPos
        q93 = .ThunkPos
        If ParseErrorExpr() Then
            Do
                p89 = .BufPos
                q89 = .ThunkPos
                If Not (ParseErrorExpr()) Then
                    .BufPos = p89
                    .ThunkPos = q89
                    Exit Do
                End If
                pvPushThunk ucsAct_1_SequenceExpr, .CaptureBegin, .CaptureEnd
            Loop
        Else
            .BufPos = p93
            .ThunkPos = q93
            pvPushThunk ucsAct_2_SequenceExpr, .CaptureBegin, .CaptureEnd
        End If
        Call ConsoleTrace(ctx.BufPos + 1, "SequenceExpr", 2, ctx.UserData)
    End With
End Sub

Private Function ParseChoiceLiteral() As Boolean
    Dim p1133 As Long
    Dim p1129 As Long

    With ctx
        p1133 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 1, ctx.UserData)
        p1129 = .BufPos
        If ParseSLASH() Then
            Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 2, ctx.UserData)
            ParseChoiceLiteral = True
            Exit Function
        Else
            .BufPos = p1129
        End If
        If ParseBAR() Then
            Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 2, ctx.UserData)
            ParseChoiceLiteral = True
            Exit Function
        Else
            .BufPos = p1129
        End If
        .BufPos = p1133
        If ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseErrorExpr() As Boolean
    Dim p118 As Long
    Dim q118 As Long
    Dim p104 As Long
    Dim q104 As Long
    Dim p111 As Long
    Dim q111 As Long

    With ctx
        p118 = .BufPos
        q118 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 1, ctx.UserData)
        If ParseVariableExpr() Then
            p104 = .BufPos
            q104 = .ThunkPos
            If ParseCARET() Then
                If ClearErrors(ctx.BufPos + 1) Then
                    p111 = .BufPos
                    q111 = .ThunkPos
                    If ParseIdentifier() Then
                        pvPushThunk ucsAct_1_ErrorExpr, .CaptureBegin, .CaptureEnd
                        GoTo L19
                    Else
                        .BufPos = p111
                        .ThunkPos = q111
                    End If
                    If ParseLiteral() Then
                        pvPushThunk ucsAct_2_ErrorExpr, .CaptureBegin, .CaptureEnd
                        GoTo L19
                    Else
                        .BufPos = p111
                        .ThunkPos = q111
                    End If
                    If AppendError("Expected rule or literal for error label", ctx.BufPos + 1) Then
                        GoTo L19
                    End If
                    .BufPos = p104
                    .ThunkPos = q104
                Else
                    .BufPos = p104
                    .ThunkPos = q104
                End If
            Else
                .BufPos = p104
                .ThunkPos = q104
            End If
L19:
            Call ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 2, ctx.UserData)
            ParseErrorExpr = True
            Exit Function
        Else
            .BufPos = p118
            .ThunkPos = q118
        End If
        If ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseVariableExpr() As Boolean
    Dim p133 As Long
    Dim q133 As Long
    Dim p129 As Long
    Dim q129 As Long

    With ctx
        p133 = .BufPos
        q133 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 1, ctx.UserData)
        p129 = .BufPos
        q129 = .ThunkPos
        If ParseIdentifier() Then
            pvPushThunk ucsAct_1_VariableExpr, .CaptureBegin, .CaptureEnd
            If ParseCOLON() Then
                If ParsePrefixExpr() Then
                    pvPushThunk ucsAct_2_VariableExpr, .CaptureBegin, .CaptureEnd
                    Call ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 2, ctx.UserData)
                    ParseVariableExpr = True
                    Exit Function
                Else
                    .BufPos = p129
                    .ThunkPos = q129
                End If
            Else
                .BufPos = p129
                .ThunkPos = q129
            End If
        Else
            .BufPos = p129
            .ThunkPos = q129
        End If
        If ParsePrefixExpr() Then
            Call ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 2, ctx.UserData)
            ParseVariableExpr = True
            Exit Function
        Else
            .BufPos = p129
            .ThunkPos = q129
        End If
        .BufPos = p133
        .ThunkPos = q133
        If ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCARET() As Boolean
    Dim p1211 As Long
    Dim p1204 As Long

    With ctx
        p1211 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CARET", 1, ctx.UserData)
        p1204 = .BufPos
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
        .BufPos = p1211
        If ConsoleTrace(ctx.BufPos + 1, "CARET", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCOLON() As Boolean
    Dim p1172 As Long
    Dim p1165 As Long

    With ctx
        p1172 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "COLON", 1, ctx.UserData)
        p1165 = .BufPos
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
        .BufPos = p1172
        If ConsoleTrace(ctx.BufPos + 1, "COLON", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParsePrefixExpr() As Boolean
    Dim p156 As Long
    Dim q156 As Long
    Dim p146 As Long
    Dim q146 As Long

    With ctx
        p156 = .BufPos
        q156 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 1, ctx.UserData)
        p146 = .BufPos
        q146 = .ThunkPos
        If ParseAND() Then
            If ParseCodeBlock() Then
                pvPushThunk ucsAct_1_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p146
                .ThunkPos = q146
            End If
        Else
            .BufPos = p146
            .ThunkPos = q146
        End If
        If ParseAND() Then
            If ParseSuffixExpr() Then
                pvPushThunk ucsAct_2_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p146
                .ThunkPos = q146
            End If
        Else
            .BufPos = p146
            .ThunkPos = q146
        End If
        If ParseNOT() Then
            If ParseSuffixExpr() Then
                pvPushThunk ucsAct_3_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p146
                .ThunkPos = q146
            End If
        Else
            .BufPos = p146
            .ThunkPos = q146
        End If
        If ParseSuffixExpr() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
            ParsePrefixExpr = True
            Exit Function
        Else
            .BufPos = p146
            .ThunkPos = q146
        End If
        .BufPos = p156
        .ThunkPos = q156
        If ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseAND() As Boolean
    Dim p859 As Long
    Dim p852 As Long

    With ctx
        p859 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "AND", 1, ctx.UserData)
        p852 = .BufPos
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
        .BufPos = p859
        If ConsoleTrace(ctx.BufPos + 1, "AND", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCodeBlock() As Boolean
    Dim p331 As Long
    Dim p296 As Long
    Dim p307 As Long
    Dim p302 As Long
    Dim lCaptureBegin As Long
    Dim p315 As Long
    Dim lCaptureEnd As Long
    Dim p325 As Long

    With ctx
        p331 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CodeBlock", 1, ctx.UserData)
        p296 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """{""", 0, ctx.UserData)
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """{""", 2, ctx.UserData)
            Do
                p307 = .BufPos
                Do
                    p302 = .BufPos
                    Call ConsoleTrace(ctx.BufPos + 1, "[ \t]", 0, ctx.UserData)
                    Select Case .BufData(.BufPos)
                    Case 32, 9                      ' [ \t]
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[ \t]", 2, ctx.UserData)
                    Case Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[ \t]", 3, ctx.UserData) And False) Then
                            .BufPos = p302
                            Exit Do
                        End If
                    End Select
                Loop
                If Not (ParseEndOfLine()) Then
                    .BufPos = p307
                    Exit Do
                End If
            Loop
            lCaptureBegin = .BufPos
            Do
                p315 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^{}]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^{}]", 3, ctx.UserData) And False) Then
                        If Not (ParseCodeNC()) Then
                            .BufPos = p315
                            .BufPos = p315
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
                                .BufPos = p315
                                .BufPos = p315
                                Exit Do
                            End If
                        End If
                    End If
                End Select
            Loop
            lCaptureEnd = .BufPos
            p325 = .BufPos
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
            .BufPos = p331
            GoTo L22
        End If
        If ConsoleTrace(ctx.BufPos + 1, """{""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p331
L22:
        If ConsoleTrace(ctx.BufPos + 1, "CodeBlock", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSuffixExpr() As Boolean
    Dim p178 As Long
    Dim q178 As Long
    Dim p168 As Long
    Dim q168 As Long

    With ctx
        p178 = .BufPos
        q178 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 1, ctx.UserData)
        If ParsePrimaryExpr() Then
            p168 = .BufPos
            q168 = .ThunkPos
            If ParseQUESTION() Then
                pvPushThunk ucsAct_1_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L27
            Else
                .BufPos = p168
                .ThunkPos = q168
            End If
            If ParseSTAR() Then
                pvPushThunk ucsAct_2_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L27
            Else
                .BufPos = p168
                .ThunkPos = q168
            End If
            If ParsePLUS() Then
                pvPushThunk ucsAct_3_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L27
            Else
                .BufPos = p168
                .ThunkPos = q168
            End If
            .BufPos = p168
            .ThunkPos = q168
L27:
            Call ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 2, ctx.UserData)
            ParseSuffixExpr = True
            Exit Function
        Else
            .BufPos = p178
            .ThunkPos = q178
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseNOT() As Boolean
    Dim p872 As Long
    Dim p865 As Long

    With ctx
        p872 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "NOT", 1, ctx.UserData)
        p865 = .BufPos
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
        .BufPos = p872
        If ConsoleTrace(ctx.BufPos + 1, "NOT", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParsePrimaryExpr() As Boolean
    Dim p249 As Long
    Dim q249 As Long
    Dim p210 As Long
    Dim q210 As Long
    Dim p185 As Long
    Dim q185 As Long
    Dim p199 As Long
    Dim q199 As Long
    Dim p189 As Long
    Dim q189 As Long
    Dim p192 As Long
    Dim q192 As Long
    Dim p196 As Long
    Dim q196 As Long
    Dim p209 As Long
    Dim q209 As Long
    Dim p219 As Long
    Dim q219 As Long
    Dim p228 As Long
    Dim q228 As Long

    With ctx
        p249 = .BufPos
        q249 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 1, ctx.UserData)
        p210 = .BufPos
        q210 = .ThunkPos
        If ParseIdentifier() Then
            pvPushThunk ucsAct_1_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p185 = .BufPos
            q185 = .ThunkPos
            If IsParamRule(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                If ParseParamBlock() Then
                    pvPushThunk ucsAct_2_PrimaryExpr, .CaptureBegin, .CaptureEnd
                Else
                    .BufPos = p185
                    .ThunkPos = q185
                End If
            End If
            p199 = .BufPos
            q199 = .ThunkPos
            p189 = .BufPos
            q189 = .ThunkPos
            If Not (ParseParamNC()) Then
                .BufPos = p189
                .ThunkPos = q189
            End If
            p192 = .BufPos
            q192 = .ThunkPos
            If Not (ParseLiteralNC()) Then
                .BufPos = p192
                .ThunkPos = q192
            End If
            Do
                p196 = .BufPos
                q196 = .ThunkPos
                If Not (ParseRuleOptNC()) Then
                    .BufPos = p196
                    .ThunkPos = q196
                    Exit Do
                End If
            Loop
            If ParseAssignLiteral() Then
                .BufPos = p210
                .ThunkPos = q210
            Else
                .BufPos = p199
                .ThunkPos = q199
                Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                ParsePrimaryExpr = True
                Exit Function
            End If
        Else
            .BufPos = p210
            .ThunkPos = q210
        End If
        If ParseOPEN() Then
            If ClearErrors(ctx.BufPos + 1) Then
                p209 = .BufPos
                q209 = .ThunkPos
                Call ParseExpression
                If ParseCLOSE() Then
                    Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                    ParsePrimaryExpr = True
                    Exit Function
                Else
                    .BufPos = p209
                    .ThunkPos = q209
                End If
                If AppendError("Invalid expression", ctx.BufPos + 1) Then
                    Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                    ParsePrimaryExpr = True
                    Exit Function
                End If
                .BufPos = p210
                .ThunkPos = q210
            Else
                .BufPos = p210
                .ThunkPos = q210
            End If
        Else
            .BufPos = p210
            .ThunkPos = q210
        End If
        If ParseLiteralNoSpace() Then
            pvPushThunk ucsAct_3_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p219 = .BufPos
            q219 = .ThunkPos
            If ParseIgnoreCase() Then
                pvPushThunk ucsAct_4_PrimaryExpr, .CaptureBegin, .CaptureEnd
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                ParsePrimaryExpr = True
                Exit Function
            Else
                .BufPos = p219
                .ThunkPos = q219
            End If
            Call ParseEMPTY
            pvPushThunk ucsAct_4_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
            .BufPos = p210
            .ThunkPos = q210
        Else
            .BufPos = p210
            .ThunkPos = q210
        End If
        If ParseClassNoSpace() Then
            pvPushThunk ucsAct_5_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p228 = .BufPos
            q228 = .ThunkPos
            If ParseIgnoreCase() Then
                pvPushThunk ucsAct_6_PrimaryExpr, .CaptureBegin, .CaptureEnd
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                ParsePrimaryExpr = True
                Exit Function
            Else
                .BufPos = p228
                .ThunkPos = q228
            End If
            Call ParseEMPTY
            pvPushThunk ucsAct_6_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
            .BufPos = p210
            .ThunkPos = q210
        Else
            .BufPos = p210
            .ThunkPos = q210
        End If
        If ParseDOT() Then
            pvPushThunk ucsAct_7_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p210
            .ThunkPos = q210
        End If
        If ParseCodeBlock() Then
            pvPushThunk ucsAct_8_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p210
            .ThunkPos = q210
        End If
        If ParseBEGIN() Then
            pvPushThunk ucsAct_9_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p210
            .ThunkPos = q210
        End If
        If ParseEND() Then
            pvPushThunk ucsAct_10_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p210
            .ThunkPos = q210
        End If
        .BufPos = p249
        .ThunkPos = q249
        If ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseQUESTION() As Boolean
    Dim p885 As Long
    Dim p878 As Long

    With ctx
        p885 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "QUESTION", 1, ctx.UserData)
        p878 = .BufPos
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
        .BufPos = p885
        If ConsoleTrace(ctx.BufPos + 1, "QUESTION", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSTAR() As Boolean
    Dim p898 As Long
    Dim p891 As Long

    With ctx
        p898 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "STAR", 1, ctx.UserData)
        p891 = .BufPos
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
        .BufPos = p898
        If ConsoleTrace(ctx.BufPos + 1, "STAR", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParsePLUS() As Boolean
    Dim p911 As Long
    Dim p904 As Long

    With ctx
        p911 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "PLUS", 1, ctx.UserData)
        p904 = .BufPos
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
        .BufPos = p911
        If ConsoleTrace(ctx.BufPos + 1, "PLUS", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseParamNC() As Boolean
    Dim p414 As Long
    Dim p393 As Long
    Dim p399 As Long
    Dim p408 As Long

    With ctx
        p414 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ParamNC", 1, ctx.UserData)
        p393 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """(""", 0, ctx.UserData)
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """(""", 2, ctx.UserData)
            Do
                p399 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^()]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 40, 41                         ' [()]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^()]", 3, ctx.UserData) And False) Then
                        If Not (ParseParamNC()) Then
                            .BufPos = p399
                            .BufPos = p399
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
                                .BufPos = p399
                                .BufPos = p399
                                Exit Do
                            End If
                        End If
                    End If
                End Select
            Loop
            p408 = .BufPos
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
            .BufPos = p414
            GoTo L33
        End If
        If ConsoleTrace(ctx.BufPos + 1, """(""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p414
L33:
        If ConsoleTrace(ctx.BufPos + 1, "ParamNC", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseLiteralNC() As Boolean
    Dim p598 As Long
    Dim p594 As Long
    Dim p550 As Long
    Dim p560 As Long
    Dim p558 As Long
    Dim p556 As Long
    Dim p567 As Long
    Dim p574 As Long
    Dim p584 As Long
    Dim p582 As Long
    Dim p580 As Long
    Dim p591 As Long

    With ctx
        p598 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 1, ctx.UserData)
        p594 = .BufPos
        p550 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
            Do
                p560 = .BufPos
                p558 = .BufPos
                p556 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
                If .BufData(.BufPos) = 39 Then      ' "'"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
                    .BufPos = p560
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p558
                If Not (ParseChar()) Then
                    .BufPos = p560
                    Exit Do
                End If
            Loop
            p567 = .BufPos
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
            .BufPos = p594
            GoTo L36
        End If
        If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p594
L36:
        p574 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
            Do
                p584 = .BufPos
                p582 = .BufPos
                p580 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
                If .BufData(.BufPos) = 34 Then      ' """
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
                    .BufPos = p584
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p582
                If Not (ParseChar()) Then
                    .BufPos = p584
                    Exit Do
                End If
            Loop
            p591 = .BufPos
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
            .BufPos = p594
            GoTo L37
        End If
        If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p594
L37:
        .BufPos = p598
        If ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseRuleOptNC() As Boolean
    Dim p445 As Long
    Dim p433 As Long
    Dim p439 As Long

    With ctx
        p445 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "RuleOptNC", 1, ctx.UserData)
        p433 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """-""", 0, ctx.UserData)
        If .BufData(.BufPos) = 45 Then              ' "-"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """-""", 2, ctx.UserData)
            If ParseIdentStart() Then
                Do
                    p439 = .BufPos
                    If Not (ParseIdentCont()) Then
                        .BufPos = p439
                        Exit Do
                    End If
                Loop
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "RuleOptNC", 2, ctx.UserData)
                Call pvSetAdvance
                ParseRuleOptNC = True
                Exit Function
            Else
                .BufPos = p445
            End If
            GoTo L38
        End If
        If ConsoleTrace(ctx.BufPos + 1, """-""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p445
L38:
        If ConsoleTrace(ctx.BufPos + 1, "RuleOptNC", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseOPEN() As Boolean
    Dim p924 As Long
    Dim p917 As Long

    With ctx
        p924 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "OPEN", 1, ctx.UserData)
        p917 = .BufPos
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
        .BufPos = p924
        If ConsoleTrace(ctx.BufPos + 1, "OPEN", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCLOSE() As Boolean
    Dim p937 As Long
    Dim p930 As Long

    With ctx
        p937 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CLOSE", 1, ctx.UserData)
        p930 = .BufPos
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
        .BufPos = p937
        If ConsoleTrace(ctx.BufPos + 1, "CLOSE", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseLiteralNoSpace() As Boolean
    Dim p544 As Long
    Dim p540 As Long
    Dim p493 As Long
    Dim lCaptureBegin As Long
    Dim p506 As Long
    Dim p503 As Long
    Dim p501 As Long
    Dim lCaptureEnd As Long
    Dim p513 As Long
    Dim p519 As Long
    Dim p531 As Long
    Dim p529 As Long
    Dim p527 As Long
    Dim p538 As Long

    With ctx
        p544 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 1, ctx.UserData)
        p540 = .BufPos
        p493 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
            lCaptureBegin = .BufPos
            Do
                p506 = .BufPos
                p503 = .BufPos
                p501 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
                If .BufData(.BufPos) = 39 Then      ' "'"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
                    .BufPos = p506
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p503
                If Not (ParseChar()) Then
                    .BufPos = p506
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            p513 = .BufPos
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
            .BufPos = p540
            GoTo L39
        End If
        If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p540
L39:
        p519 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
            lCaptureBegin = .BufPos
            Do
                p531 = .BufPos
                p529 = .BufPos
                p527 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
                If .BufData(.BufPos) = 34 Then      ' """
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
                    .BufPos = p531
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p529
                If Not (ParseChar()) Then
                    .BufPos = p531
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            p538 = .BufPos
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
            .BufPos = p540
            GoTo L40
        End If
        If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p540
L40:
        .BufPos = p544
        If ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseIgnoreCase() As Boolean
    Dim p612 As Long
    Dim lCaptureBegin As Long
    Dim p605 As Long
    Dim lCaptureEnd As Long

    With ctx
        p612 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "IgnoreCase", 1, ctx.UserData)
        lCaptureBegin = .BufPos
        p605 = .BufPos
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
        .BufPos = p612
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
    Dim p652 As Long
    Dim p627 As Long
    Dim lCaptureBegin As Long
    Dim p640 As Long
    Dim p637 As Long
    Dim p635 As Long
    Dim lCaptureEnd As Long
    Dim p647 As Long

    With ctx
        p652 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ClassNoSpace", 1, ctx.UserData)
        p627 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """[""", 0, ctx.UserData)
        If .BufData(.BufPos) = 91 Then              ' "["
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """[""", 2, ctx.UserData)
            lCaptureBegin = .BufPos
            Do
                p640 = .BufPos
                p637 = .BufPos
                p635 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, """]""", 0, ctx.UserData)
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, """]""", 2, ctx.UserData)
                    .BufPos = p640
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, """]""", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p637
                If Not (ParseRange()) Then
                    .BufPos = p640
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            p647 = .BufPos
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
            .BufPos = p652
            GoTo L42
        End If
        If ConsoleTrace(ctx.BufPos + 1, """[""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p652
L42:
        If ConsoleTrace(ctx.BufPos + 1, "ClassNoSpace", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseDOT() As Boolean
    Dim p950 As Long
    Dim p943 As Long

    With ctx
        p950 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "DOT", 1, ctx.UserData)
        p943 = .BufPos
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
        .BufPos = p950
        If ConsoleTrace(ctx.BufPos + 1, "DOT", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseBEGIN() As Boolean
    Dim p1102 As Long
    Dim p1095 As Long

    With ctx
        p1102 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "BEGIN", 1, ctx.UserData)
        p1095 = .BufPos
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
        .BufPos = p1102
        If ConsoleTrace(ctx.BufPos + 1, "BEGIN", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseEND() As Boolean
    Dim p1115 As Long
    Dim p1108 As Long

    With ctx
        p1115 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "END", 1, ctx.UserData)
        p1108 = .BufPos
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
        .BufPos = p1115
        If ConsoleTrace(ctx.BufPos + 1, "END", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSettingName() As Boolean
    Dim p278 As Long
    Dim p272 As Long

    With ctx
        p278 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 1, ctx.UserData)
        If ParseAT() Then
            p272 = .BufPos
            If ParseIdentifier() Then
                Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p272
            End If
            If ParseLiteral() Then
                Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p272
            End If
            Call ParseEMPTY
            Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
            ParseSettingName = True
            Exit Function
            .BufPos = p278
        Else
            .BufPos = p278
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SettingName", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSettingValue() As Boolean
    Dim p290 As Long
    Dim p282 As Long
    Dim p284 As Long

    With ctx
        p290 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 1, ctx.UserData)
        p282 = .BufPos
        If ParseIdentifier() Then
            p284 = .BufPos
            If Not (ParseSEMI()) Then
                .BufPos = p284
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p282
        End If
        If ParseLiteral() Then
            p284 = .BufPos
            If Not (ParseSEMI()) Then
                .BufPos = p284
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p282
        End If
        If ParseCodeBlock() Then
            p284 = .BufPos
            If Not (ParseSEMI()) Then
                .BufPos = p284
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p282
        End If
        .BufPos = p290
        If ConsoleTrace(ctx.BufPos + 1, "SettingValue", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseAT() As Boolean
    Dim p1198 As Long
    Dim p1191 As Long

    With ctx
        p1198 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "AT", 1, ctx.UserData)
        p1191 = .BufPos
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
        .BufPos = p1198
        If ConsoleTrace(ctx.BufPos + 1, "AT", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseEndOfLine() As Boolean
    Dim p1082 As Long
    Dim p1065 As Long
    Dim p1071 As Long
    Dim p1077 As Long

    With ctx
        p1082 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 1, ctx.UserData)
        p1065 = .BufPos
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
        p1071 = .BufPos
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
        .BufPos = p1065
        p1077 = .BufPos
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
        .BufPos = p1065
        .BufPos = p1082
        If ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCodeNC() As Boolean
    Dim p358 As Long
    Dim p337 As Long
    Dim p343 As Long
    Dim p352 As Long

    With ctx
        p358 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CodeNC", 1, ctx.UserData)
        p337 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """{""", 0, ctx.UserData)
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """{""", 2, ctx.UserData)
            Do
                p343 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^{}]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^{}]", 3, ctx.UserData) And False) Then
                        If Not (ParseCodeNC()) Then
                            .BufPos = p343
                            .BufPos = p343
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
                                .BufPos = p343
                                .BufPos = p343
                                Exit Do
                            End If
                        End If
                    End If
                End Select
            Loop
            p352 = .BufPos
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
            .BufPos = p358
            GoTo L48
        End If
        If ConsoleTrace(ctx.BufPos + 1, """{""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p358
L48:
        If ConsoleTrace(ctx.BufPos + 1, "CodeNC", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseIdentStart() As Boolean
    Dim p468 As Long
    Dim p463 As Long

    With ctx
        p468 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "IdentStart", 1, ctx.UserData)
        p463 = .BufPos
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
        .BufPos = p468
        If ConsoleTrace(ctx.BufPos + 1, "IdentStart", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseIdentCont() As Boolean
    Dim p479 As Long
    Dim p474 As Long

    With ctx
        p479 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "IdentCont", 1, ctx.UserData)
        p474 = .BufPos
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
        .BufPos = p479
        If ConsoleTrace(ctx.BufPos + 1, "IdentCont", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseChar() As Boolean
    Dim p811 As Long
    Dim p715 As Long
    Dim p681 As Long
    Dim p687 As Long
    Dim p694 As Long
    Dim p700 As Long
    Dim p707 As Long
    Dim p713 As Long
    Dim p720 As Long
    Dim p726 As Long
    Dim p733 As Long
    Dim p740 As Long
    Dim p746 As Long
    Dim p753 As Long
    Dim p759 As Long
    Dim p765 As Long
    Dim p772 As Long
    Dim p778 As Long
    Dim p784 As Long
    Dim p790 As Long
    Dim p796 As Long
    Dim p805 As Long
    Dim p803 As Long

    With ctx
        p811 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Char", 1, ctx.UserData)
        p715 = .BufPos
        p681 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            p687 = .BufPos
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
            .BufPos = p715
            GoTo L51
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p715
L51:
        p694 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            p700 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[0-3]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 48 To 51                           ' [0-3]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[0-3]", 2, ctx.UserData)
                p707 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 2, ctx.UserData)
                    p713 = .BufPos
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
                    .BufPos = p715
                    GoTo L57
                End Select
                If ConsoleTrace(ctx.BufPos + 1, "[0-7]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p715
L57:
                GoTo L58
            End Select
            If ConsoleTrace(ctx.BufPos + 1, "[0-3]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p715
L58:
            GoTo L56
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p715
L56:
        p720 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            p726 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 48 To 55                           ' [0-7]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 2, ctx.UserData)
                p733 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 2, ctx.UserData)
                    GoTo L65
                End Select
                If ConsoleTrace(ctx.BufPos + 1, "[0-7]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p733
L65:
                Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                Call pvSetAdvance
                ParseChar = True
                Exit Function
            End Select
            If ConsoleTrace(ctx.BufPos + 1, "[0-7]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p715
            GoTo L64
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p715
L64:
        p740 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\x""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 120 Then ' "\\x"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """\\x""", 2, ctx.UserData)
            p746 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                p753 = .BufPos
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
                .BufPos = p715
                GoTo L69
            End Select
            If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p715
L69:
            GoTo L68
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\x""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p715
L68:
        p759 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\u""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 117 Then ' "\\u"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """\\u""", 2, ctx.UserData)
            p765 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                p772 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 48 To 57, 97 To 102, 65 To 70  ' [0-9a-fA-F]
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                    p778 = .BufPos
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
                    Select Case .BufData(.BufPos)
                    Case 48 To 57, 97 To 102, 65 To 70 ' [0-9a-fA-F]
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                        p784 = .BufPos
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
                        .BufPos = p715
                        GoTo L81
                    End Select
                    If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                        '--- do nothing
                    End If
                    .BufPos = p715
L81:
                    GoTo L82
                End Select
                If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p715
L82:
                GoTo L83
            End Select
            If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p715
L83:
            GoTo L80
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\u""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p715
L80:
        p790 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            p796 = .BufPos
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
            .BufPos = p715
            GoTo L87
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p715
L87:
        p805 = .BufPos
        p803 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            .BufPos = p715
            GoTo L88
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p805
        If .BufPos < .BufSize Then
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
            Call pvSetAdvance
            ParseChar = True
            Exit Function
        Else
            .BufPos = p715
        End If
L88:
        .BufPos = p811
        If ConsoleTrace(ctx.BufPos + 1, "Char", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        Call pvSetExpected("char")
    End With
End Function

Private Function ParseClass() As Boolean
    Dim p621 As Long

    With ctx
        p621 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Class", 1, ctx.UserData)
        If ParseClassNoSpace() Then
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "Class", 2, ctx.UserData)
            ParseClass = True
            Exit Function
        Else
            .BufPos = p621
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Class", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseRange() As Boolean
    Dim p675 As Long
    Dim p671 As Long
    Dim p659 As Long
    Dim p668 As Long
    Dim p666 As Long

    With ctx
        p675 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Range", 1, ctx.UserData)
        p671 = .BufPos
        If ParseChar() Then
            p659 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, """-""", 0, ctx.UserData)
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, """-""", 2, ctx.UserData)
                p668 = .BufPos
                p666 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, """]""", 0, ctx.UserData)
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, """]""", 2, ctx.UserData)
                    .BufPos = p671
                    GoTo L91
                End If
                If ConsoleTrace(ctx.BufPos + 1, """]""", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p668
                If ParseChar() Then
                    Call ConsoleTrace(ctx.BufPos + 1, "Range", 2, ctx.UserData)
                    Call pvSetAdvance
                    ParseRange = True
                    Exit Function
                Else
                    .BufPos = p671
                End If
L91:
                GoTo L90
            End If
            If ConsoleTrace(ctx.BufPos + 1, """-""", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p671
L90:
        Else
            .BufPos = p671
        End If
        If ParseChar() Then
            Call ConsoleTrace(ctx.BufPos + 1, "Range", 2, ctx.UserData)
            Call pvSetAdvance
            ParseRange = True
            Exit Function
        Else
            .BufPos = p671
        End If
        .BufPos = p675
        If ConsoleTrace(ctx.BufPos + 1, "Range", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseLEFTARROW() As Boolean
    Dim p825 As Long
    Dim p818 As Long

    With ctx
        p825 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LEFTARROW", 1, ctx.UserData)
        p818 = .BufPos
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
        .BufPos = p825
        If ConsoleTrace(ctx.BufPos + 1, "LEFTARROW", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSLASH() As Boolean
    Dim p846 As Long
    Dim p832 As Long
    Dim p840 As Long
    Dim p838 As Long

    With ctx
        p846 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SLASH", 1, ctx.UserData)
        p832 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """/""", 0, ctx.UserData)
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """/""", 2, ctx.UserData)
            p840 = .BufPos
            p838 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, """/""", 0, ctx.UserData)
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, """/""", 2, ctx.UserData)
                .BufPos = p846
                GoTo L95
            End If
            If ConsoleTrace(ctx.BufPos + 1, """/""", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p840
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "SLASH", 2, ctx.UserData)
            Call pvSetAdvance
            ParseSLASH = True
            Exit Function
L95:
            GoTo L94
        End If
        If ConsoleTrace(ctx.BufPos + 1, """/""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p846
L94:
        If ConsoleTrace(ctx.BufPos + 1, "SLASH", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSpace() As Boolean
    Dim p1059 As Long
    Dim p1047 As Long
    Dim p1053 As Long

    With ctx
        p1059 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Space", 1, ctx.UserData)
        p1047 = .BufPos
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
        p1053 = .BufPos
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
        .BufPos = p1047
        If ParseEndOfLine() Then
            Call ConsoleTrace(ctx.BufPos + 1, "Space", 2, ctx.UserData)
            Call pvSetAdvance
            ParseSpace = True
            Exit Function
        Else
            .BufPos = p1047
        End If
        .BufPos = p1059
        If ConsoleTrace(ctx.BufPos + 1, "Space", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseLineComment() As Boolean
    Dim p1006 As Long
    Dim p984 As Long
    Dim p996 As Long
    Dim p1002 As Long
    Dim p990 As Long

    With ctx
        p1006 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 1, ctx.UserData)
        p984 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """#""", 0, ctx.UserData)
        If .BufData(.BufPos) = 35 Then              ' "#"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """#""", 2, ctx.UserData)
            Do
                p996 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 13, 10                         ' [\r\n]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 3, ctx.UserData) And False) Then
                        .BufPos = p996
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 2, ctx.UserData)
                    Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 3, ctx.UserData) And False) Then
                            .BufPos = p996
                            Exit Do
                        End If
                    End If
                End Select
            Loop
            p1002 = .BufPos
            If ParseEndOfLine() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p1002
            End If
            If ParseEndOfFile() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p1002
            End If
            .BufPos = p1006
            GoTo L98
        End If
        If ConsoleTrace(ctx.BufPos + 1, """#""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        p990 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """//""", 0, ctx.UserData)
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 47 Then ' "//"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """//""", 2, ctx.UserData)
            Do
                p996 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 13, 10                         ' [\r\n]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 3, ctx.UserData) And False) Then
                        .BufPos = p996
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 2, ctx.UserData)
                    Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 3, ctx.UserData) And False) Then
                            .BufPos = p996
                            Exit Do
                        End If
                    End If
                End Select
            Loop
            p1002 = .BufPos
            If ParseEndOfLine() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p1002
            End If
            If ParseEndOfFile() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p1002
            End If
            .BufPos = p1006
            GoTo L98
        End If
        If ConsoleTrace(ctx.BufPos + 1, """//""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p984
        .BufPos = p1006
L98:
        If ConsoleTrace(ctx.BufPos + 1, "LineComment", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseBlockComment() As Boolean
    Dim p1041 As Long
    Dim p1012 As Long
    Dim p1018 As Long
    Dim p1027 As Long
    Dim p1025 As Long
    Dim p1036 As Long

    With ctx
        p1041 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "BlockComment", 1, ctx.UserData)
        p1012 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """/*""", 0, ctx.UserData)
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 42 Then ' "/*"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """/*""", 2, ctx.UserData)
            Do
                p1018 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^*/]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 42, 47                         ' [*/]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^*/]", 3, ctx.UserData) And False) Then
                        If Not (ParseBlockComment()) Then
                            .BufPos = p1018
                            p1027 = .BufPos
                            p1025 = .BufPos
                            Call ConsoleTrace(ctx.BufPos + 1, """*/""", 0, ctx.UserData)
                            If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                                .BufPos = .BufPos + 2
                                Call ConsoleTrace(ctx.BufPos + 1, """*/""", 2, ctx.UserData)
                                .BufPos = p1018
                                .BufPos = p1018
                                Exit Do
                            End If
                            If ConsoleTrace(ctx.BufPos + 1, """*/""", 3, ctx.UserData) And False Then
                                '--- do nothing
                            End If
                            .BufPos = p1027
                            If .BufPos < .BufSize Then
                                .BufPos = .BufPos + 1
                            Else
                                .BufPos = p1018
                                .BufPos = p1018
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
                                .BufPos = p1018
                                p1027 = .BufPos
                                p1025 = .BufPos
                                Call ConsoleTrace(ctx.BufPos + 1, """*/""", 0, ctx.UserData)
                                If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                                    .BufPos = .BufPos + 2
                                    Call ConsoleTrace(ctx.BufPos + 1, """*/""", 2, ctx.UserData)
                                    .BufPos = p1018
                                    .BufPos = p1018
                                    Exit Do
                                End If
                                If ConsoleTrace(ctx.BufPos + 1, """*/""", 3, ctx.UserData) And False Then
                                    '--- do nothing
                                End If
                                .BufPos = p1027
                                If .BufPos < .BufSize Then
                                    .BufPos = .BufPos + 1
                                Else
                                    .BufPos = p1018
                                    .BufPos = p1018
                                    Exit Do
                                End If
                            End If
                        End If
                    End If
                End Select
            Loop
            p1036 = .BufPos
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
            .BufPos = p1041
            GoTo L105
        End If
        If ConsoleTrace(ctx.BufPos + 1, """/*""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p1041
L105:
        If ConsoleTrace(ctx.BufPos + 1, "BlockComment", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseEQUAL() As Boolean
    Dim p1146 As Long
    Dim p1139 As Long

    With ctx
        p1146 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "EQUAL", 1, ctx.UserData)
        p1139 = .BufPos
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
        .BufPos = p1146
        If ConsoleTrace(ctx.BufPos + 1, "EQUAL", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseBAR() As Boolean
    Dim p1159 As Long
    Dim p1152 As Long

    With ctx
        p1159 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "BAR", 1, ctx.UserData)
        p1152 = .BufPos
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
        .BufPos = p1159
        If ConsoleTrace(ctx.BufPos + 1, "BAR", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParsePrepExpr() As Boolean
    Dim p1245 As Long
    Dim p1238 As Long

    With ctx
        p1245 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 1, ctx.UserData)
        p1238 = .BufPos
        If ParseIdentifier() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1238
        End If
        If ParseLiteral() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1238
        End If
        If ParseClass() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1238
        End If
        If ParseCodeNC() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1238
        End If
        If ParseParamBlock() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1238
        End If
        .BufPos = p1245
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
