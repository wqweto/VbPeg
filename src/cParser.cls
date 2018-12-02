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
' Auto-generated on 2.12.2018 23:55:58
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const LNG_MAXINT            As Long = 2 ^ 31 - 1

'= generated enum ========================================================

Private Enum UcsParserActionsEnum
    ucsAct_2_Setting
    ucsAct_1_Setting
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
    ParserVersion = "2.12.2018 23:55:58"
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

Public Function ParsePreprocess() As Boolean
    Dim p1194 As Long
    Dim i1189 As Long
    Dim p1185 As Long
    Dim p1179 As Long

    With ctx
        p1194 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Preprocess", 1, ctx.UserData)
        Call ParseSpacing
        For i1189 = 0 To LNG_MAXINT
            p1185 = .BufPos
            If ParseIdentifier() Then
                If ParseParamNC() Then
                    p1179 = .BufPos
                    If Not (ParseLiteral()) Then
                        .BufPos = p1179
                    End If
                    If ParseAssignLiteral() Then
                        '--- this might be premature
                        AddParamRule(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin))
                    Else
                        .BufPos = p1185
                        If Not (ParsePrepExpr()) Then
                            .BufPos = p1185
                            If .BufPos < .BufSize Then
                                .BufPos = .BufPos + 1
                                Call ParseSpacing
                            Else
                                .BufPos = p1185
                                Exit For
                            End If
                        End If
                    End If
                Else
                    .BufPos = p1185
                    If Not (ParsePrepExpr()) Then
                        .BufPos = p1185
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                            Call ParseSpacing
                        Else
                            .BufPos = p1185
                            Exit For
                        End If
                    End If
                End If
            Else
                .BufPos = p1185
                If Not (ParsePrepExpr()) Then
                    .BufPos = p1185
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        Call ParseSpacing
                    Else
                        .BufPos = p1185
                        Exit For
                    End If
                End If
            End If
        Next
        If i1189 <> 0 Then
            Call ConsoleTrace(ctx.BufPos + 1, "Preprocess", 2, ctx.UserData)
            ParsePreprocess = True
            Exit Function
        Else
            .BufPos = p1194
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Preprocess", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Sub ParseSpacing()
    Dim p809 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "Spacing", 1, ctx.UserData)
        Do
            p809 = .BufPos
            If Not (ParseSpace()) Then
                .BufPos = p809
                If Not (ParseLineComment()) Then
                    .BufPos = p809
                    If Not (ParseBlockComment()) Then
                        .BufPos = p809
                        .BufPos = p809
                        Exit Do
                    End If
                End If
            End If
        Loop
        Call ConsoleTrace(ctx.BufPos + 1, "Spacing", 2, ctx.UserData)
    End With
End Sub

Public Function ParseSetting() As Boolean
    Dim p257 As Long
    Dim q257 As Long
    Dim p253 As Long
    Dim q253 As Long

    With ctx
        p257 = .BufPos
        q257 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 1, ctx.UserData)
        If ParseSettingName() Then
            If CheckSettingName(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                pvPushThunk ucsAct_1_Setting, .CaptureBegin, .CaptureEnd
                If ClearErrors(ctx.BufPos + 1) And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                    p253 = .BufPos
                    q253 = .ThunkPos
                    If ParseSettingValue() Then
                        pvPushThunk ucsAct_2_Setting, .CaptureBegin, .CaptureEnd
                        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 2, ctx.UserData)
                        ParseSetting = True
                        Exit Function
                    Else
                        .BufPos = p253
                        .ThunkPos = q253
                    End If
                    If AppendError("Invalid value for setting '" & .UserData.Tag & "'", ctx.BufPos + 1) Then
                        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 2, ctx.UserData)
                        ParseSetting = True
                        Exit Function
                    End If
                    .BufPos = p257
                    .ThunkPos = q257
                Else
                    .BufPos = p257
                    .ThunkPos = q257
                End If
            Else
                .BufPos = p257
                .ThunkPos = q257
            End If
        Else
            .BufPos = p257
            .ThunkPos = q257
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Setting", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseRule() As Boolean
    Dim p56 As Long
    Dim q56 As Long
    Dim p35 As Long
    Dim q35 As Long
    Dim p40 As Long
    Dim q40 As Long
    Dim p51 As Long
    Dim q51 As Long

    With ctx
        p56 = .BufPos
        q56 = .ThunkPos
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
                If ParseAssignLiteral() Then
                    Call ParseExpression
                    pvPushThunk ucsAct_4_Rule, .CaptureBegin, .CaptureEnd
                    p51 = .BufPos
                    q51 = .ThunkPos
                    If Not (ParseSEMI()) Then
                        .BufPos = p51
                        .ThunkPos = q51
                    End If
                    Call ConsoleTrace(ctx.BufPos + 1, "Rule", 2, ctx.UserData)
                    ParseRule = True
                    Exit Function
                Else
                    .BufPos = p56
                    .ThunkPos = q56
                End If
            Else
                .BufPos = p56
                .ThunkPos = q56
            End If
        Else
            .BufPos = p56
            .ThunkPos = q56
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Rule", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseEndOfFile() As Boolean
    Dim p928 As Long
    Dim p924 As Long

    With ctx
        p928 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "EndOfFile", 1, ctx.UserData)
        p924 = .BufPos
        If .BufPos < .BufSize Then
            .BufPos = p928
        Else
            .BufPos = p924
            Call ConsoleTrace(ctx.BufPos + 1, "EndOfFile", 2, ctx.UserData)
            ParseEndOfFile = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "EndOfFile", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseIdentifier() As Boolean
    Dim p295 As Long
    Dim lCaptureBegin As Long
    Dim p288 As Long
    Dim lCaptureEnd As Long

    With ctx
        p295 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Identifier", 1, ctx.UserData)
        lCaptureBegin = .BufPos
        If ParseIdentStart() Then
            Do
                p288 = .BufPos
                If Not (ParseIdentCont()) Then
                    .BufPos = p288
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
            .BufPos = p295
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Identifier", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseParamBlock() As Boolean
    Dim p1024 As Long
    Dim p1001 As Long
    Dim lCaptureBegin As Long
    Dim p1009 As Long
    Dim lCaptureEnd As Long
    Dim p1018 As Long

    With ctx
        p1024 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ParamBlock", 1, ctx.UserData)
        p1001 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """(""", 0, ctx.UserData)
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """(""", 2, ctx.UserData)
            lCaptureBegin = .BufPos
            Do
                p1009 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^()]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 40, 41                         ' [()]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^()]", 3, ctx.UserData) And False) Then
                        If Not (ParseParamNC()) Then
                            .BufPos = p1009
                            .BufPos = p1009
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
                                .BufPos = p1009
                                .BufPos = p1009
                                Exit Do
                            End If
                        End If
                    End If
                End Select
            Loop
            lCaptureEnd = .BufPos
            p1018 = .BufPos
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
            .BufPos = p1024
            GoTo L11
        End If
        If ConsoleTrace(ctx.BufPos + 1, """(""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p1024
L11:
        If ConsoleTrace(ctx.BufPos + 1, "ParamBlock", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseLiteral() As Boolean
    Dim p325 As Long

    With ctx
        p325 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Literal", 1, ctx.UserData)
        If ParseLiteralNoSpace() Then
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "Literal", 2, ctx.UserData)
            ParseLiteral = True
            Exit Function
        Else
            .BufPos = p325
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Literal", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseAssignLiteral() As Boolean
    Dim p1086 As Long
    Dim p1082 As Long

    With ctx
        p1086 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 1, ctx.UserData)
        p1082 = .BufPos
        If ParseLEFTARROW() Then
            Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 2, ctx.UserData)
            ParseAssignLiteral = True
            Exit Function
        Else
            .BufPos = p1082
        End If
        If ParseEQUAL() Then
            Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 2, ctx.UserData)
            ParseAssignLiteral = True
            Exit Function
        Else
            .BufPos = p1082
        End If
        .BufPos = p1086
        If ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Sub ParseExpression()
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "Expression", 1, ctx.UserData)
        Call ParseChoiceExpr
        Call ConsoleTrace(ctx.BufPos + 1, "Expression", 2, ctx.UserData)
L14:
    End With
End Sub

Private Function ParseSEMI() As Boolean
    Dim p1147 As Long
    Dim p1140 As Long

    With ctx
        p1147 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SEMI", 1, ctx.UserData)
        p1140 = .BufPos
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
        .BufPos = p1147
        If ConsoleTrace(ctx.BufPos + 1, "SEMI", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Sub ParseChoiceExpr()
    Dim p70 As Long
    Dim q70 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceExpr", 1, ctx.UserData)
        Call ParseSequenceExpr
        Do
            p70 = .BufPos
            q70 = .ThunkPos
            If Not (ParseChoiceLiteral()) Then
                .BufPos = p70
                .ThunkPos = q70
                Exit Do
            End If
            Call ParseSequenceExpr
            pvPushThunk ucsAct_1_ChoiceExpr, .CaptureBegin, .CaptureEnd
        Loop
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceExpr", 2, ctx.UserData)
L15:
    End With
End Sub

Public Sub ParseSequenceExpr()
    Dim p87 As Long
    Dim q87 As Long
    Dim p83 As Long
    Dim q83 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "SequenceExpr", 1, ctx.UserData)
        p87 = .BufPos
        q87 = .ThunkPos
        If ParseErrorExpr() Then
            Do
                p83 = .BufPos
                q83 = .ThunkPos
                If Not (ParseErrorExpr()) Then
                    .BufPos = p83
                    .ThunkPos = q83
                    Exit Do
                End If
                pvPushThunk ucsAct_1_SequenceExpr, .CaptureBegin, .CaptureEnd
            Loop
        Else
            .BufPos = p87
            .ThunkPos = q87
            pvPushThunk ucsAct_2_SequenceExpr, .CaptureBegin, .CaptureEnd
        End If
        Call ConsoleTrace(ctx.BufPos + 1, "SequenceExpr", 2, ctx.UserData)
    End With
End Sub

Public Function ParseChoiceLiteral() As Boolean
    Dim p1095 As Long
    Dim p1091 As Long

    With ctx
        p1095 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 1, ctx.UserData)
        p1091 = .BufPos
        If ParseSLASH() Then
            Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 2, ctx.UserData)
            ParseChoiceLiteral = True
            Exit Function
        Else
            .BufPos = p1091
        End If
        If ParseBAR() Then
            Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 2, ctx.UserData)
            ParseChoiceLiteral = True
            Exit Function
        Else
            .BufPos = p1091
        End If
        .BufPos = p1095
        If ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseErrorExpr() As Boolean
    Dim p112 As Long
    Dim q112 As Long
    Dim p98 As Long
    Dim q98 As Long
    Dim p105 As Long
    Dim q105 As Long

    With ctx
        p112 = .BufPos
        q112 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 1, ctx.UserData)
        If ParseVariableExpr() Then
            p98 = .BufPos
            q98 = .ThunkPos
            If ParseCARET() Then
                If ClearErrors(ctx.BufPos + 1) Then
                    p105 = .BufPos
                    q105 = .ThunkPos
                    If ParseIdentifier() Then
                        pvPushThunk ucsAct_1_ErrorExpr, .CaptureBegin, .CaptureEnd
                        GoTo L18
                    Else
                        .BufPos = p105
                        .ThunkPos = q105
                    End If
                    If ParseLiteral() Then
                        pvPushThunk ucsAct_2_ErrorExpr, .CaptureBegin, .CaptureEnd
                        GoTo L18
                    Else
                        .BufPos = p105
                        .ThunkPos = q105
                    End If
                    If AppendError("Expected rule or literal for error label", ctx.BufPos + 1) Then
                        GoTo L18
                    End If
                    .BufPos = p98
                    .ThunkPos = q98
                Else
                    .BufPos = p98
                    .ThunkPos = q98
                End If
            Else
                .BufPos = p98
                .ThunkPos = q98
            End If
L18:
            Call ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 2, ctx.UserData)
            ParseErrorExpr = True
            Exit Function
        Else
            .BufPos = p112
            .ThunkPos = q112
        End If
        If ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseVariableExpr() As Boolean
    Dim p127 As Long
    Dim q127 As Long
    Dim p123 As Long
    Dim q123 As Long

    With ctx
        p127 = .BufPos
        q127 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 1, ctx.UserData)
        p123 = .BufPos
        q123 = .ThunkPos
        If ParseIdentifier() Then
            pvPushThunk ucsAct_1_VariableExpr, .CaptureBegin, .CaptureEnd
            If ParseCOLON() Then
                If ParsePrefixExpr() Then
                    pvPushThunk ucsAct_2_VariableExpr, .CaptureBegin, .CaptureEnd
                    Call ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 2, ctx.UserData)
                    ParseVariableExpr = True
                    Exit Function
                Else
                    .BufPos = p123
                    .ThunkPos = q123
                End If
            Else
                .BufPos = p123
                .ThunkPos = q123
            End If
        Else
            .BufPos = p123
            .ThunkPos = q123
        End If
        If ParsePrefixExpr() Then
            Call ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 2, ctx.UserData)
            ParseVariableExpr = True
            Exit Function
        Else
            .BufPos = p123
            .ThunkPos = q123
        End If
        .BufPos = p127
        .ThunkPos = q127
        If ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCARET() As Boolean
    Dim p1173 As Long
    Dim p1166 As Long

    With ctx
        p1173 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CARET", 1, ctx.UserData)
        p1166 = .BufPos
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
        .BufPos = p1173
        If ConsoleTrace(ctx.BufPos + 1, "CARET", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCOLON() As Boolean
    Dim p1134 As Long
    Dim p1127 As Long

    With ctx
        p1134 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "COLON", 1, ctx.UserData)
        p1127 = .BufPos
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
        .BufPos = p1134
        If ConsoleTrace(ctx.BufPos + 1, "COLON", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParsePrefixExpr() As Boolean
    Dim p150 As Long
    Dim q150 As Long
    Dim p140 As Long
    Dim q140 As Long

    With ctx
        p150 = .BufPos
        q150 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 1, ctx.UserData)
        p140 = .BufPos
        q140 = .ThunkPos
        If ParseAND() Then
            If ParseCodeBlock() Then
                pvPushThunk ucsAct_1_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p140
                .ThunkPos = q140
            End If
        Else
            .BufPos = p140
            .ThunkPos = q140
        End If
        If ParseAND() Then
            If ParseSuffixExpr() Then
                pvPushThunk ucsAct_2_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p140
                .ThunkPos = q140
            End If
        Else
            .BufPos = p140
            .ThunkPos = q140
        End If
        If ParseNOT() Then
            If ParseSuffixExpr() Then
                pvPushThunk ucsAct_3_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p140
                .ThunkPos = q140
            End If
        Else
            .BufPos = p140
            .ThunkPos = q140
        End If
        If ParseSuffixExpr() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
            ParsePrefixExpr = True
            Exit Function
        Else
            .BufPos = p140
            .ThunkPos = q140
        End If
        .BufPos = p150
        .ThunkPos = q150
        If ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseAND() As Boolean
    Dim p697 As Long
    Dim p690 As Long

    With ctx
        p697 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "AND", 1, ctx.UserData)
        p690 = .BufPos
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
        .BufPos = p697
        If ConsoleTrace(ctx.BufPos + 1, "AND", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseCodeBlock() As Boolean
    Dim p968 As Long
    Dim p934 As Long
    Dim p944 As Long
    Dim p940 As Long
    Dim lCaptureBegin As Long
    Dim p952 As Long
    Dim lCaptureEnd As Long
    Dim p962 As Long

    With ctx
        p968 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CodeBlock", 1, ctx.UserData)
        p934 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """{""", 0, ctx.UserData)
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """{""", 2, ctx.UserData)
            Do
                p944 = .BufPos
                Do
                    p940 = .BufPos
                    Call ConsoleTrace(ctx.BufPos + 1, "[ \t]", 0, ctx.UserData)
                    Select Case .BufData(.BufPos)
                    Case 32, 9                      ' [ \t]
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[ \t]", 2, ctx.UserData)
                    Case Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[ \t]", 3, ctx.UserData) And False) Then
                            .BufPos = p940
                            Exit Do
                        End If
                    End Select
                Loop
                If Not (ParseEndOfLine()) Then
                    .BufPos = p944
                    Exit Do
                End If
            Loop
            lCaptureBegin = .BufPos
            Do
                p952 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^{}]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^{}]", 3, ctx.UserData) And False) Then
                        If Not (ParseCodeNC()) Then
                            .BufPos = p952
                            .BufPos = p952
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
                                .BufPos = p952
                                .BufPos = p952
                                Exit Do
                            End If
                        End If
                    End If
                End Select
            Loop
            lCaptureEnd = .BufPos
            p962 = .BufPos
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
            .BufPos = p968
            GoTo L21
        End If
        If ConsoleTrace(ctx.BufPos + 1, """{""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p968
L21:
        If ConsoleTrace(ctx.BufPos + 1, "CodeBlock", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseSuffixExpr() As Boolean
    Dim p172 As Long
    Dim q172 As Long
    Dim p162 As Long
    Dim q162 As Long

    With ctx
        p172 = .BufPos
        q172 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 1, ctx.UserData)
        If ParsePrimaryExpr() Then
            p162 = .BufPos
            q162 = .ThunkPos
            If ParseQUESTION() Then
                pvPushThunk ucsAct_1_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L26
            Else
                .BufPos = p162
                .ThunkPos = q162
            End If
            If ParseSTAR() Then
                pvPushThunk ucsAct_2_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L26
            Else
                .BufPos = p162
                .ThunkPos = q162
            End If
            If ParsePLUS() Then
                pvPushThunk ucsAct_3_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L26
            Else
                .BufPos = p162
                .ThunkPos = q162
            End If
            .BufPos = p162
            .ThunkPos = q162
L26:
            Call ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 2, ctx.UserData)
            ParseSuffixExpr = True
            Exit Function
        Else
            .BufPos = p172
            .ThunkPos = q172
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseNOT() As Boolean
    Dim p710 As Long
    Dim p703 As Long

    With ctx
        p710 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "NOT", 1, ctx.UserData)
        p703 = .BufPos
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
        .BufPos = p710
        If ConsoleTrace(ctx.BufPos + 1, "NOT", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParsePrimaryExpr() As Boolean
    Dim p240 As Long
    Dim q240 As Long
    Dim p201 As Long
    Dim q201 As Long
    Dim p179 As Long
    Dim q179 As Long
    Dim p190 As Long
    Dim q190 As Long
    Dim p183 As Long
    Dim q183 As Long
    Dim p186 As Long
    Dim q186 As Long
    Dim p200 As Long
    Dim q200 As Long
    Dim p210 As Long
    Dim q210 As Long
    Dim p219 As Long
    Dim q219 As Long

    With ctx
        p240 = .BufPos
        q240 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 1, ctx.UserData)
        p201 = .BufPos
        q201 = .ThunkPos
        If ParseIdentifier() Then
            pvPushThunk ucsAct_1_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p179 = .BufPos
            q179 = .ThunkPos
            If IsParamRule(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                If ParseParamBlock() Then
                    pvPushThunk ucsAct_2_PrimaryExpr, .CaptureBegin, .CaptureEnd
                Else
                    .BufPos = p179
                    .ThunkPos = q179
                End If
            End If
            p190 = .BufPos
            q190 = .ThunkPos
            p183 = .BufPos
            q183 = .ThunkPos
            If Not (ParseParamNC()) Then
                .BufPos = p183
                .ThunkPos = q183
            End If
            p186 = .BufPos
            q186 = .ThunkPos
            If Not (ParseLiteralNC()) Then
                .BufPos = p186
                .ThunkPos = q186
            End If
            If ParseAssignLiteral() Then
                .BufPos = p201
                .ThunkPos = q201
            Else
                .BufPos = p190
                .ThunkPos = q190
                Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                ParsePrimaryExpr = True
                Exit Function
            End If
        Else
            .BufPos = p201
            .ThunkPos = q201
        End If
        If ParseOPEN() Then
            If ClearErrors(ctx.BufPos + 1) Then
                p200 = .BufPos
                q200 = .ThunkPos
                Call ParseExpression
                If ParseCLOSE() Then
                    Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                    ParsePrimaryExpr = True
                    Exit Function
                Else
                    .BufPos = p200
                    .ThunkPos = q200
                End If
                If AppendError("Invalid expression", ctx.BufPos + 1) Then
                    Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                    ParsePrimaryExpr = True
                    Exit Function
                End If
                .BufPos = p201
                .ThunkPos = q201
            Else
                .BufPos = p201
                .ThunkPos = q201
            End If
        Else
            .BufPos = p201
            .ThunkPos = q201
        End If
        If ParseLiteralNoSpace() Then
            pvPushThunk ucsAct_3_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p210 = .BufPos
            q210 = .ThunkPos
            If ParseIgnoreCase() Then
                pvPushThunk ucsAct_4_PrimaryExpr, .CaptureBegin, .CaptureEnd
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                ParsePrimaryExpr = True
                Exit Function
            Else
                .BufPos = p210
                .ThunkPos = q210
            End If
            Call ParseEMPTY
            pvPushThunk ucsAct_4_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
            .BufPos = p201
            .ThunkPos = q201
        Else
            .BufPos = p201
            .ThunkPos = q201
        End If
        If ParseClassNoSpace() Then
            pvPushThunk ucsAct_5_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p219 = .BufPos
            q219 = .ThunkPos
            If ParseIgnoreCase() Then
                pvPushThunk ucsAct_6_PrimaryExpr, .CaptureBegin, .CaptureEnd
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                ParsePrimaryExpr = True
                Exit Function
            Else
                .BufPos = p219
                .ThunkPos = q219
            End If
            Call ParseEMPTY
            pvPushThunk ucsAct_6_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
            .BufPos = p201
            .ThunkPos = q201
        Else
            .BufPos = p201
            .ThunkPos = q201
        End If
        If ParseDOT() Then
            pvPushThunk ucsAct_7_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p201
            .ThunkPos = q201
        End If
        If ParseCodeBlock() Then
            pvPushThunk ucsAct_8_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p201
            .ThunkPos = q201
        End If
        If ParseBEGIN() Then
            pvPushThunk ucsAct_9_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p201
            .ThunkPos = q201
        End If
        If ParseEND() Then
            pvPushThunk ucsAct_10_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p201
            .ThunkPos = q201
        End If
        .BufPos = p240
        .ThunkPos = q240
        If ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseQUESTION() As Boolean
    Dim p723 As Long
    Dim p716 As Long

    With ctx
        p723 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "QUESTION", 1, ctx.UserData)
        p716 = .BufPos
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
        .BufPos = p723
        If ConsoleTrace(ctx.BufPos + 1, "QUESTION", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSTAR() As Boolean
    Dim p736 As Long
    Dim p729 As Long

    With ctx
        p736 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "STAR", 1, ctx.UserData)
        p729 = .BufPos
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
        .BufPos = p736
        If ConsoleTrace(ctx.BufPos + 1, "STAR", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParsePLUS() As Boolean
    Dim p749 As Long
    Dim p742 As Long

    With ctx
        p749 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "PLUS", 1, ctx.UserData)
        p742 = .BufPos
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
        .BufPos = p749
        If ConsoleTrace(ctx.BufPos + 1, "PLUS", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseParamNC() As Boolean
    Dim p1051 As Long
    Dim p1030 As Long
    Dim p1036 As Long
    Dim p1045 As Long

    With ctx
        p1051 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ParamNC", 1, ctx.UserData)
        p1030 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """(""", 0, ctx.UserData)
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """(""", 2, ctx.UserData)
            Do
                p1036 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^()]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 40, 41                         ' [()]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^()]", 3, ctx.UserData) And False) Then
                        If Not (ParseParamNC()) Then
                            .BufPos = p1036
                            .BufPos = p1036
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
                                .BufPos = p1036
                                .BufPos = p1036
                                Exit Do
                            End If
                        End If
                    End If
                End Select
            Loop
            p1045 = .BufPos
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
            .BufPos = p1051
            GoTo L32
        End If
        If ConsoleTrace(ctx.BufPos + 1, """(""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p1051
L32:
        If ConsoleTrace(ctx.BufPos + 1, "ParamNC", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseLiteralNC() As Boolean
    Dim p436 As Long
    Dim p432 As Long
    Dim p388 As Long
    Dim p398 As Long
    Dim p396 As Long
    Dim p394 As Long
    Dim p405 As Long
    Dim p412 As Long
    Dim p422 As Long
    Dim p420 As Long
    Dim p418 As Long
    Dim p429 As Long

    With ctx
        p436 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 1, ctx.UserData)
        p432 = .BufPos
        p388 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
            Do
                p398 = .BufPos
                p396 = .BufPos
                p394 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
                If .BufData(.BufPos) = 39 Then      ' "'"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
                    .BufPos = p398
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p396
                If Not (ParseChar()) Then
                    .BufPos = p398
                    Exit Do
                End If
            Loop
            p405 = .BufPos
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
            .BufPos = p432
            GoTo L35
        End If
        If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p432
L35:
        p412 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
            Do
                p422 = .BufPos
                p420 = .BufPos
                p418 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
                If .BufData(.BufPos) = 34 Then      ' """
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
                    .BufPos = p422
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p420
                If Not (ParseChar()) Then
                    .BufPos = p422
                    Exit Do
                End If
            Loop
            p429 = .BufPos
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
            .BufPos = p432
            GoTo L36
        End If
        If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p432
L36:
        .BufPos = p436
        If ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseOPEN() As Boolean
    Dim p762 As Long
    Dim p755 As Long

    With ctx
        p762 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "OPEN", 1, ctx.UserData)
        p755 = .BufPos
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
        .BufPos = p762
        If ConsoleTrace(ctx.BufPos + 1, "OPEN", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCLOSE() As Boolean
    Dim p775 As Long
    Dim p768 As Long

    With ctx
        p775 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CLOSE", 1, ctx.UserData)
        p768 = .BufPos
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
        .BufPos = p775
        If ConsoleTrace(ctx.BufPos + 1, "CLOSE", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseLiteralNoSpace() As Boolean
    Dim p382 As Long
    Dim p378 As Long
    Dim p331 As Long
    Dim lCaptureBegin As Long
    Dim p344 As Long
    Dim p341 As Long
    Dim p339 As Long
    Dim lCaptureEnd As Long
    Dim p351 As Long
    Dim p357 As Long
    Dim p369 As Long
    Dim p367 As Long
    Dim p365 As Long
    Dim p376 As Long

    With ctx
        p382 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 1, ctx.UserData)
        p378 = .BufPos
        p331 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
            lCaptureBegin = .BufPos
            Do
                p344 = .BufPos
                p341 = .BufPos
                p339 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
                If .BufData(.BufPos) = 39 Then      ' "'"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
                    .BufPos = p344
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p341
                If Not (ParseChar()) Then
                    .BufPos = p344
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            p351 = .BufPos
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
            .BufPos = p378
            GoTo L37
        End If
        If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p378
L37:
        p357 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
            lCaptureBegin = .BufPos
            Do
                p369 = .BufPos
                p367 = .BufPos
                p365 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
                If .BufData(.BufPos) = 34 Then      ' """
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
                    .BufPos = p369
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p367
                If Not (ParseChar()) Then
                    .BufPos = p369
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            p376 = .BufPos
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
            .BufPos = p378
            GoTo L38
        End If
        If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p378
L38:
        .BufPos = p382
        If ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseIgnoreCase() As Boolean
    Dim p450 As Long
    Dim lCaptureBegin As Long
    Dim p443 As Long
    Dim lCaptureEnd As Long

    With ctx
        p450 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "IgnoreCase", 1, ctx.UserData)
        lCaptureBegin = .BufPos
        p443 = .BufPos
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
        .BufPos = p450
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

Public Function ParseClassNoSpace() As Boolean
    Dim p490 As Long
    Dim p465 As Long
    Dim lCaptureBegin As Long
    Dim p478 As Long
    Dim p475 As Long
    Dim p473 As Long
    Dim lCaptureEnd As Long
    Dim p485 As Long

    With ctx
        p490 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ClassNoSpace", 1, ctx.UserData)
        p465 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """[""", 0, ctx.UserData)
        If .BufData(.BufPos) = 91 Then              ' "["
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """[""", 2, ctx.UserData)
            lCaptureBegin = .BufPos
            Do
                p478 = .BufPos
                p475 = .BufPos
                p473 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, """]""", 0, ctx.UserData)
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, """]""", 2, ctx.UserData)
                    .BufPos = p478
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, """]""", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p475
                If Not (ParseRange()) Then
                    .BufPos = p478
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            p485 = .BufPos
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
            .BufPos = p490
            GoTo L40
        End If
        If ConsoleTrace(ctx.BufPos + 1, """[""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p490
L40:
        If ConsoleTrace(ctx.BufPos + 1, "ClassNoSpace", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseDOT() As Boolean
    Dim p788 As Long
    Dim p781 As Long

    With ctx
        p788 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "DOT", 1, ctx.UserData)
        p781 = .BufPos
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
        .BufPos = p788
        If ConsoleTrace(ctx.BufPos + 1, "DOT", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseBEGIN() As Boolean
    Dim p1064 As Long
    Dim p1057 As Long

    With ctx
        p1064 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "BEGIN", 1, ctx.UserData)
        p1057 = .BufPos
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
        .BufPos = p1064
        If ConsoleTrace(ctx.BufPos + 1, "BEGIN", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseEND() As Boolean
    Dim p1077 As Long
    Dim p1070 As Long

    With ctx
        p1077 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "END", 1, ctx.UserData)
        p1070 = .BufPos
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
        .BufPos = p1077
        If ConsoleTrace(ctx.BufPos + 1, "END", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseSettingName() As Boolean
    Dim p269 As Long
    Dim p263 As Long

    With ctx
        p269 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 1, ctx.UserData)
        If ParseAT() Then
            p263 = .BufPos
            If ParseIdentifier() Then
                Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p263
            End If
            If ParseLiteral() Then
                Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p263
            End If
            Call ParseEMPTY
            Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
            ParseSettingName = True
            Exit Function
            .BufPos = p269
        Else
            .BufPos = p269
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SettingName", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseSettingValue() As Boolean
    Dim p281 As Long
    Dim p273 As Long
    Dim p275 As Long

    With ctx
        p281 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 1, ctx.UserData)
        p273 = .BufPos
        If ParseIdentifier() Then
            p275 = .BufPos
            If Not (ParseSEMI()) Then
                .BufPos = p275
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p273
        End If
        If ParseLiteral() Then
            p275 = .BufPos
            If Not (ParseSEMI()) Then
                .BufPos = p275
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p273
        End If
        If ParseCodeBlock() Then
            p275 = .BufPos
            If Not (ParseSEMI()) Then
                .BufPos = p275
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p273
        End If
        .BufPos = p281
        If ConsoleTrace(ctx.BufPos + 1, "SettingValue", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseAT() As Boolean
    Dim p1160 As Long
    Dim p1153 As Long

    With ctx
        p1160 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "AT", 1, ctx.UserData)
        p1153 = .BufPos
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
        .BufPos = p1160
        If ConsoleTrace(ctx.BufPos + 1, "AT", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseIdentStart() As Boolean
    Dim p306 As Long
    Dim p301 As Long

    With ctx
        p306 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "IdentStart", 1, ctx.UserData)
        p301 = .BufPos
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
        .BufPos = p306
        If ConsoleTrace(ctx.BufPos + 1, "IdentStart", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseIdentCont() As Boolean
    Dim p317 As Long
    Dim p312 As Long

    With ctx
        p317 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "IdentCont", 1, ctx.UserData)
        p312 = .BufPos
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
        .BufPos = p317
        If ConsoleTrace(ctx.BufPos + 1, "IdentCont", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseChar() As Boolean
    Dim p649 As Long
    Dim p553 As Long
    Dim p519 As Long
    Dim p525 As Long
    Dim p532 As Long
    Dim p538 As Long
    Dim p545 As Long
    Dim p551 As Long
    Dim p558 As Long
    Dim p564 As Long
    Dim p571 As Long
    Dim p578 As Long
    Dim p584 As Long
    Dim p591 As Long
    Dim p597 As Long
    Dim p603 As Long
    Dim p610 As Long
    Dim p616 As Long
    Dim p622 As Long
    Dim p628 As Long
    Dim p634 As Long
    Dim p643 As Long
    Dim p641 As Long

    With ctx
        p649 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Char", 1, ctx.UserData)
        p553 = .BufPos
        p519 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            p525 = .BufPos
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
            .BufPos = p553
            GoTo L45
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p553
L45:
        p532 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            p538 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[0-3]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 48 To 51                           ' [0-3]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[0-3]", 2, ctx.UserData)
                p545 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 2, ctx.UserData)
                    p551 = .BufPos
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
                    .BufPos = p553
                    GoTo L51
                End Select
                If ConsoleTrace(ctx.BufPos + 1, "[0-7]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p553
L51:
                GoTo L52
            End Select
            If ConsoleTrace(ctx.BufPos + 1, "[0-3]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p553
L52:
            GoTo L50
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p553
L50:
        p558 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            p564 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 48 To 55                           ' [0-7]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 2, ctx.UserData)
                p571 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 2, ctx.UserData)
                    GoTo L59
                End Select
                If ConsoleTrace(ctx.BufPos + 1, "[0-7]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p571
L59:
                Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                Call pvSetAdvance
                ParseChar = True
                Exit Function
            End Select
            If ConsoleTrace(ctx.BufPos + 1, "[0-7]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p553
            GoTo L58
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p553
L58:
        p578 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\x""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 120 Then ' "\\x"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """\\x""", 2, ctx.UserData)
            p584 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                p591 = .BufPos
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
                .BufPos = p553
                GoTo L63
            End Select
            If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p553
L63:
            GoTo L62
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\x""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p553
L62:
        p597 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\u""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 117 Then ' "\\u"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """\\u""", 2, ctx.UserData)
            p603 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                p610 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 48 To 57, 97 To 102, 65 To 70  ' [0-9a-fA-F]
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                    p616 = .BufPos
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
                    Select Case .BufData(.BufPos)
                    Case 48 To 57, 97 To 102, 65 To 70 ' [0-9a-fA-F]
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                        p622 = .BufPos
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
                        .BufPos = p553
                        GoTo L75
                    End Select
                    If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                        '--- do nothing
                    End If
                    .BufPos = p553
L75:
                    GoTo L76
                End Select
                If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p553
L76:
                GoTo L77
            End Select
            If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p553
L77:
            GoTo L74
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\u""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p553
L74:
        p628 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            p634 = .BufPos
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
            .BufPos = p553
            GoTo L81
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p553
L81:
        p643 = .BufPos
        p641 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            .BufPos = p553
            GoTo L82
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p643
        If .BufPos < .BufSize Then
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
            Call pvSetAdvance
            ParseChar = True
            Exit Function
        Else
            .BufPos = p553
        End If
L82:
        .BufPos = p649
        If ConsoleTrace(ctx.BufPos + 1, "Char", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        Call pvSetExpected("char")
    End With
End Function

Public Function ParseClass() As Boolean
    Dim p459 As Long

    With ctx
        p459 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Class", 1, ctx.UserData)
        If ParseClassNoSpace() Then
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "Class", 2, ctx.UserData)
            ParseClass = True
            Exit Function
        Else
            .BufPos = p459
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Class", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseRange() As Boolean
    Dim p513 As Long
    Dim p509 As Long
    Dim p497 As Long
    Dim p506 As Long
    Dim p504 As Long

    With ctx
        p513 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Range", 1, ctx.UserData)
        p509 = .BufPos
        If ParseChar() Then
            p497 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, """-""", 0, ctx.UserData)
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, """-""", 2, ctx.UserData)
                p506 = .BufPos
                p504 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, """]""", 0, ctx.UserData)
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, """]""", 2, ctx.UserData)
                    .BufPos = p509
                    GoTo L85
                End If
                If ConsoleTrace(ctx.BufPos + 1, """]""", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p506
                If ParseChar() Then
                    Call ConsoleTrace(ctx.BufPos + 1, "Range", 2, ctx.UserData)
                    Call pvSetAdvance
                    ParseRange = True
                    Exit Function
                Else
                    .BufPos = p509
                End If
L85:
                GoTo L84
            End If
            If ConsoleTrace(ctx.BufPos + 1, """-""", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p509
L84:
        Else
            .BufPos = p509
        End If
        If ParseChar() Then
            Call ConsoleTrace(ctx.BufPos + 1, "Range", 2, ctx.UserData)
            Call pvSetAdvance
            ParseRange = True
            Exit Function
        Else
            .BufPos = p509
        End If
        .BufPos = p513
        If ConsoleTrace(ctx.BufPos + 1, "Range", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseLEFTARROW() As Boolean
    Dim p663 As Long
    Dim p656 As Long

    With ctx
        p663 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LEFTARROW", 1, ctx.UserData)
        p656 = .BufPos
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
        .BufPos = p663
        If ConsoleTrace(ctx.BufPos + 1, "LEFTARROW", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSLASH() As Boolean
    Dim p684 As Long
    Dim p670 As Long
    Dim p678 As Long
    Dim p676 As Long

    With ctx
        p684 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SLASH", 1, ctx.UserData)
        p670 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """/""", 0, ctx.UserData)
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """/""", 2, ctx.UserData)
            p678 = .BufPos
            p676 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, """/""", 0, ctx.UserData)
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, """/""", 2, ctx.UserData)
                .BufPos = p684
                GoTo L89
            End If
            If ConsoleTrace(ctx.BufPos + 1, """/""", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p678
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "SLASH", 2, ctx.UserData)
            Call pvSetAdvance
            ParseSLASH = True
            Exit Function
L89:
            GoTo L88
        End If
        If ConsoleTrace(ctx.BufPos + 1, """/""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p684
L88:
        If ConsoleTrace(ctx.BufPos + 1, "SLASH", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseSpace() As Boolean
    Dim p898 As Long
    Dim p886 As Long
    Dim p892 As Long

    With ctx
        p898 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Space", 1, ctx.UserData)
        p886 = .BufPos
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
        p892 = .BufPos
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
        .BufPos = p886
        If ParseEndOfLine() Then
            Call ConsoleTrace(ctx.BufPos + 1, "Space", 2, ctx.UserData)
            Call pvSetAdvance
            ParseSpace = True
            Exit Function
        Else
            .BufPos = p886
        End If
        .BufPos = p898
        If ConsoleTrace(ctx.BufPos + 1, "Space", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseLineComment() As Boolean
    Dim p845 As Long
    Dim p822 As Long
    Dim p834 As Long
    Dim p841 As Long
    Dim p828 As Long

    With ctx
        p845 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 1, ctx.UserData)
        p822 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """#""", 0, ctx.UserData)
        If .BufData(.BufPos) = 35 Then              ' "#"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """#""", 2, ctx.UserData)
            Do
                p834 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 13, 10                         ' [\r\n]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 3, ctx.UserData) And False) Then
                        .BufPos = p834
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 2, ctx.UserData)
                    Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 3, ctx.UserData) And False) Then
                            .BufPos = p834
                            Exit Do
                        End If
                    End If
                End Select
            Loop
            p841 = .BufPos
            If ParseEndOfLine() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p841
            End If
            If ParseEndOfFile() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p841
            End If
            .BufPos = p845
            GoTo L92
        End If
        If ConsoleTrace(ctx.BufPos + 1, """#""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        p828 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """//""", 0, ctx.UserData)
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 47 Then ' "//"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """//""", 2, ctx.UserData)
            Do
                p834 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 13, 10                         ' [\r\n]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 3, ctx.UserData) And False) Then
                        .BufPos = p834
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 2, ctx.UserData)
                    Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 3, ctx.UserData) And False) Then
                            .BufPos = p834
                            Exit Do
                        End If
                    End If
                End Select
            Loop
            p841 = .BufPos
            If ParseEndOfLine() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p841
            End If
            If ParseEndOfFile() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p841
            End If
            .BufPos = p845
            GoTo L92
        End If
        If ConsoleTrace(ctx.BufPos + 1, """//""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p822
        .BufPos = p845
L92:
        If ConsoleTrace(ctx.BufPos + 1, "LineComment", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseBlockComment() As Boolean
    Dim p880 As Long
    Dim p851 As Long
    Dim p857 As Long
    Dim p866 As Long
    Dim p864 As Long
    Dim p875 As Long

    With ctx
        p880 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "BlockComment", 1, ctx.UserData)
        p851 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """/*""", 0, ctx.UserData)
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 42 Then ' "/*"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """/*""", 2, ctx.UserData)
            Do
                p857 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^*/]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 42, 47                         ' [*/]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^*/]", 3, ctx.UserData) And False) Then
                        If Not (ParseBlockComment()) Then
                            .BufPos = p857
                            p866 = .BufPos
                            p864 = .BufPos
                            Call ConsoleTrace(ctx.BufPos + 1, """*/""", 0, ctx.UserData)
                            If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                                .BufPos = .BufPos + 2
                                Call ConsoleTrace(ctx.BufPos + 1, """*/""", 2, ctx.UserData)
                                .BufPos = p857
                                .BufPos = p857
                                Exit Do
                            End If
                            If ConsoleTrace(ctx.BufPos + 1, """*/""", 3, ctx.UserData) And False Then
                                '--- do nothing
                            End If
                            .BufPos = p866
                            If .BufPos < .BufSize Then
                                .BufPos = .BufPos + 1
                            Else
                                .BufPos = p857
                                .BufPos = p857
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
                                .BufPos = p857
                                p866 = .BufPos
                                p864 = .BufPos
                                Call ConsoleTrace(ctx.BufPos + 1, """*/""", 0, ctx.UserData)
                                If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                                    .BufPos = .BufPos + 2
                                    Call ConsoleTrace(ctx.BufPos + 1, """*/""", 2, ctx.UserData)
                                    .BufPos = p857
                                    .BufPos = p857
                                    Exit Do
                                End If
                                If ConsoleTrace(ctx.BufPos + 1, """*/""", 3, ctx.UserData) And False Then
                                    '--- do nothing
                                End If
                                .BufPos = p866
                                If .BufPos < .BufSize Then
                                    .BufPos = .BufPos + 1
                                Else
                                    .BufPos = p857
                                    .BufPos = p857
                                    Exit Do
                                End If
                            End If
                        End If
                    End If
                End Select
            Loop
            p875 = .BufPos
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
            .BufPos = p880
            GoTo L99
        End If
        If ConsoleTrace(ctx.BufPos + 1, """/*""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p880
L99:
        If ConsoleTrace(ctx.BufPos + 1, "BlockComment", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseEndOfLine() As Boolean
    Dim p921 As Long
    Dim p904 As Long
    Dim p910 As Long
    Dim p916 As Long

    With ctx
        p921 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 1, ctx.UserData)
        p904 = .BufPos
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
        p910 = .BufPos
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
        .BufPos = p904
        p916 = .BufPos
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
        .BufPos = p904
        .BufPos = p921
        If ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseCodeNC() As Boolean
    Dim p995 As Long
    Dim p974 As Long
    Dim p980 As Long
    Dim p989 As Long

    With ctx
        p995 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CodeNC", 1, ctx.UserData)
        p974 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """{""", 0, ctx.UserData)
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """{""", 2, ctx.UserData)
            Do
                p980 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^{}]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^{}]", 3, ctx.UserData) And False) Then
                        If Not (ParseCodeNC()) Then
                            .BufPos = p980
                            .BufPos = p980
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
                                .BufPos = p980
                                .BufPos = p980
                                Exit Do
                            End If
                        End If
                    End If
                End Select
            Loop
            p989 = .BufPos
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
            .BufPos = p995
            GoTo L103
        End If
        If ConsoleTrace(ctx.BufPos + 1, """{""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p995
L103:
        If ConsoleTrace(ctx.BufPos + 1, "CodeNC", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseEQUAL() As Boolean
    Dim p1108 As Long
    Dim p1101 As Long

    With ctx
        p1108 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "EQUAL", 1, ctx.UserData)
        p1101 = .BufPos
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
        .BufPos = p1108
        If ConsoleTrace(ctx.BufPos + 1, "EQUAL", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseBAR() As Boolean
    Dim p1121 As Long
    Dim p1114 As Long

    With ctx
        p1121 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "BAR", 1, ctx.UserData)
        p1114 = .BufPos
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
        .BufPos = p1121
        If ConsoleTrace(ctx.BufPos + 1, "BAR", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParsePrepExpr() As Boolean
    Dim p1205 As Long
    Dim p1198 As Long

    With ctx
        p1205 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 1, ctx.UserData)
        p1198 = .BufPos
        If ParseIdentifier() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1198
        End If
        If ParseLiteral() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1198
        End If
        If ParseClass() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1198
        End If
        If ParseCodeNC() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1198
        End If
        If ParseParamBlock() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1198
        End If
        .BufPos = p1205
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
    Case ucsAct_4_Rule
         lNode = .Pop: .SetExpression .Pop, lNode
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
