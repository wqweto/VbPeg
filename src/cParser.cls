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
' Auto-generated on 2.12.2018 19:47:55
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
    ParserVersion = "2.12.2018 19:47:55"
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
    Dim p1141 As Long
    Dim i1136 As Long
    Dim p1131 As Long
    Dim p1121 As Long
    Dim p1127 As Long

    With ctx
        p1141 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Preprocess", 1, ctx.UserData)
        Call ParseSpacing
        For i1136 = 0 To LNG_MAXINT
            p1131 = .BufPos
            If ParseIdentifier() Then
                If ParseCodeNC() Then
                    p1121 = .BufPos
                    If Not (ParseLiteral()) Then
                        .BufPos = p1121
                    End If
                    If ParseAssignLiteral() Then
                        '--- this might be premature
                        AddParamRule(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin))
                    Else
                        .BufPos = p1131
                        If ParseOPEN() Then
                            Do
                                p1127 = .BufPos
                                If Not (ParsePrepExpr()) Then
                                    .BufPos = p1127
                                    Exit Do
                                End If
                            Loop
                            If Not (ParseCLOSE()) Then
                                .BufPos = p1131
                                If Not (ParsePrepExpr()) Then
                                    .BufPos = p1131
                                    If .BufPos < .BufSize Then
                                        .BufPos = .BufPos + 1
                                        Call ParseSpacing
                                    Else
                                        .BufPos = p1131
                                        Exit For
                                    End If
                                End If
                            End If
                        Else
                            .BufPos = p1131
                            If Not (ParsePrepExpr()) Then
                                .BufPos = p1131
                                If .BufPos < .BufSize Then
                                    .BufPos = .BufPos + 1
                                    Call ParseSpacing
                                Else
                                    .BufPos = p1131
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Else
                    .BufPos = p1131
                    If ParseOPEN() Then
                        Do
                            p1127 = .BufPos
                            If Not (ParsePrepExpr()) Then
                                .BufPos = p1127
                                Exit Do
                            End If
                        Loop
                        If Not (ParseCLOSE()) Then
                            .BufPos = p1131
                            If Not (ParsePrepExpr()) Then
                                .BufPos = p1131
                                If .BufPos < .BufSize Then
                                    .BufPos = .BufPos + 1
                                    Call ParseSpacing
                                Else
                                    .BufPos = p1131
                                    Exit For
                                End If
                            End If
                        End If
                    Else
                        .BufPos = p1131
                        If Not (ParsePrepExpr()) Then
                            .BufPos = p1131
                            If .BufPos < .BufSize Then
                                .BufPos = .BufPos + 1
                                Call ParseSpacing
                            Else
                                .BufPos = p1131
                                Exit For
                            End If
                        End If
                    End If
                End If
            Else
                .BufPos = p1131
                If ParseOPEN() Then
                    Do
                        p1127 = .BufPos
                        If Not (ParsePrepExpr()) Then
                            .BufPos = p1127
                            Exit Do
                        End If
                    Loop
                    If Not (ParseCLOSE()) Then
                        .BufPos = p1131
                        If Not (ParsePrepExpr()) Then
                            .BufPos = p1131
                            If .BufPos < .BufSize Then
                                .BufPos = .BufPos + 1
                                Call ParseSpacing
                            Else
                                .BufPos = p1131
                                Exit For
                            End If
                        End If
                    End If
                Else
                    .BufPos = p1131
                    If Not (ParsePrepExpr()) Then
                        .BufPos = p1131
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                            Call ParseSpacing
                        Else
                            .BufPos = p1131
                            Exit For
                        End If
                    End If
                End If
            End If
        Next
        If i1136 <> 0 Then
            Call ConsoleTrace(ctx.BufPos + 1, "Preprocess", 2, ctx.UserData)
            ParsePreprocess = True
            Exit Function
        Else
            .BufPos = p1141
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Preprocess", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Sub ParseSpacing()
    Dim p808 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "Spacing", 1, ctx.UserData)
        Do
            p808 = .BufPos
            If Not (ParseSpace()) Then
                .BufPos = p808
                If Not (ParseLineComment()) Then
                    .BufPos = p808
                    If Not (ParseBlockComment()) Then
                        .BufPos = p808
                        .BufPos = p808
                        Exit Do
                    End If
                End If
            End If
        Loop
        Call ConsoleTrace(ctx.BufPos + 1, "Spacing", 2, ctx.UserData)
    End With
End Sub

Public Function ParseSetting() As Boolean
    Dim p256 As Long
    Dim q256 As Long
    Dim p252 As Long
    Dim q252 As Long

    With ctx
        p256 = .BufPos
        q256 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 1, ctx.UserData)
        If ParseSettingName() Then
            If CheckSettingName(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                pvPushThunk ucsAct_1_Setting, .CaptureBegin, .CaptureEnd
                If ClearErrors(ctx.BufPos + 1) And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                    p252 = .BufPos
                    q252 = .ThunkPos
                    If ParseSettingValue() Then
                        pvPushThunk ucsAct_2_Setting, .CaptureBegin, .CaptureEnd
                        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 2, ctx.UserData)
                        ParseSetting = True
                        Exit Function
                    Else
                        .BufPos = p252
                        .ThunkPos = q252
                    End If
                    If AppendError("Invalid value for setting '" & .UserData.Tag & "'", ctx.BufPos + 1) Then
                        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 2, ctx.UserData)
                        ParseSetting = True
                        Exit Function
                    End If
                    .BufPos = p256
                    .ThunkPos = q256
                Else
                    .BufPos = p256
                    .ThunkPos = q256
                End If
            Else
                .BufPos = p256
                .ThunkPos = q256
            End If
        Else
            .BufPos = p256
            .ThunkPos = q256
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
                If ParseCodeBlock() Then
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
    Dim p927 As Long
    Dim p923 As Long

    With ctx
        p927 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "EndOfFile", 1, ctx.UserData)
        p923 = .BufPos
        If .BufPos < .BufSize Then
            .BufPos = p927
        Else
            .BufPos = p923
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
    Dim p294 As Long
    Dim lCaptureBegin As Long
    Dim p287 As Long
    Dim lCaptureEnd As Long

    With ctx
        p294 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Identifier", 1, ctx.UserData)
        lCaptureBegin = .BufPos
        If ParseIdentStart() Then
            Do
                p287 = .BufPos
                If Not (ParseIdentCont()) Then
                    .BufPos = p287
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
            .BufPos = p294
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Identifier", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseCodeBlock() As Boolean
    Dim p966 As Long
    Dim p933 As Long
    Dim p943 As Long
    Dim p939 As Long
    Dim lCaptureBegin As Long
    Dim p951 As Long
    Dim lCaptureEnd As Long
    Dim p960 As Long

    With ctx
        p966 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CodeBlock", 1, ctx.UserData)
        p933 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """{""", 0, ctx.UserData)
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """{""", 2, ctx.UserData)
            Do
                p943 = .BufPos
                Do
                    p939 = .BufPos
                    Call ConsoleTrace(ctx.BufPos + 1, "[ \t]", 0, ctx.UserData)
                    Select Case .BufData(.BufPos)
                    Case 32, 9                      ' [ \t]
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[ \t]", 2, ctx.UserData)
                    Case Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[ \t]", 3, ctx.UserData) And False) Then
                            .BufPos = p939
                            Exit Do
                        End If
                    End Select
                Loop
                If Not (ParseEndOfLine()) Then
                    .BufPos = p943
                    Exit Do
                End If
            Loop
            lCaptureBegin = .BufPos
            Do
                p951 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^{}]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^{}]", 3, ctx.UserData) And False) Then
                        If Not (ParseCodeNC()) Then
                            .BufPos = p951
                            .BufPos = p951
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
                                .BufPos = p951
                                .BufPos = p951
                                Exit Do
                            End If
                        End If
                    End If
                End Select
            Loop
            lCaptureEnd = .BufPos
            p960 = .BufPos
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
            .BufPos = p966
            GoTo L12
        End If
        If ConsoleTrace(ctx.BufPos + 1, """{""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p966
L12:
        If ConsoleTrace(ctx.BufPos + 1, "CodeBlock", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseLiteral() As Boolean
    Dim p324 As Long

    With ctx
        p324 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Literal", 1, ctx.UserData)
        If ParseLiteralNoSpace() Then
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "Literal", 2, ctx.UserData)
            ParseLiteral = True
            Exit Function
        Else
            .BufPos = p324
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Literal", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseAssignLiteral() As Boolean
    Dim p1028 As Long
    Dim p1024 As Long

    With ctx
        p1028 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 1, ctx.UserData)
        p1024 = .BufPos
        If ParseLEFTARROW() Then
            Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 2, ctx.UserData)
            ParseAssignLiteral = True
            Exit Function
        Else
            .BufPos = p1024
        End If
        If ParseEQUAL() Then
            Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 2, ctx.UserData)
            ParseAssignLiteral = True
            Exit Function
        Else
            .BufPos = p1024
        End If
        .BufPos = p1028
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
L17:
    End With
End Sub

Private Function ParseSEMI() As Boolean
    Dim p1089 As Long
    Dim p1082 As Long

    With ctx
        p1089 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SEMI", 1, ctx.UserData)
        p1082 = .BufPos
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
        .BufPos = p1089
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
L18:
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
    Dim p1037 As Long
    Dim p1033 As Long

    With ctx
        p1037 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 1, ctx.UserData)
        p1033 = .BufPos
        If ParseSLASH() Then
            Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 2, ctx.UserData)
            ParseChoiceLiteral = True
            Exit Function
        Else
            .BufPos = p1033
        End If
        If ParseBAR() Then
            Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 2, ctx.UserData)
            ParseChoiceLiteral = True
            Exit Function
        Else
            .BufPos = p1033
        End If
        .BufPos = p1037
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
                        GoTo L21
                    Else
                        .BufPos = p105
                        .ThunkPos = q105
                    End If
                    If ParseLiteral() Then
                        pvPushThunk ucsAct_2_ErrorExpr, .CaptureBegin, .CaptureEnd
                        GoTo L21
                    Else
                        .BufPos = p105
                        .ThunkPos = q105
                    End If
                    If AppendError("Expected rule or literal for error label", ctx.BufPos + 1) Then
                        GoTo L21
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
L21:
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
    Dim p1115 As Long
    Dim p1108 As Long

    With ctx
        p1115 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CARET", 1, ctx.UserData)
        p1108 = .BufPos
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
        .BufPos = p1115
        If ConsoleTrace(ctx.BufPos + 1, "CARET", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCOLON() As Boolean
    Dim p1076 As Long
    Dim p1069 As Long

    With ctx
        p1076 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "COLON", 1, ctx.UserData)
        p1069 = .BufPos
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
        .BufPos = p1076
        If ConsoleTrace(ctx.BufPos + 1, "COLON", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParsePrefixExpr() As Boolean
    Dim p149 As Long
    Dim q149 As Long
    Dim p139 As Long
    Dim q139 As Long

    With ctx
        p149 = .BufPos
        q149 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 1, ctx.UserData)
        p139 = .BufPos
        q139 = .ThunkPos
        If ParseAND() Then
            If ParseCodeBlock() Then
                pvPushThunk ucsAct_1_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p139
                .ThunkPos = q139
            End If
        Else
            .BufPos = p139
            .ThunkPos = q139
        End If
        If ParseAND() Then
            If ParseSuffixExpr() Then
                pvPushThunk ucsAct_2_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p139
                .ThunkPos = q139
            End If
        Else
            .BufPos = p139
            .ThunkPos = q139
        End If
        If ParseNOT() Then
            If ParseSuffixExpr() Then
                pvPushThunk ucsAct_3_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p139
                .ThunkPos = q139
            End If
        Else
            .BufPos = p139
            .ThunkPos = q139
        End If
        If ParseSuffixExpr() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
            ParsePrefixExpr = True
            Exit Function
        Else
            .BufPos = p139
            .ThunkPos = q139
        End If
        .BufPos = p149
        .ThunkPos = q149
        If ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseAND() As Boolean
    Dim p696 As Long
    Dim p689 As Long

    With ctx
        p696 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "AND", 1, ctx.UserData)
        p689 = .BufPos
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
        .BufPos = p696
        If ConsoleTrace(ctx.BufPos + 1, "AND", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseSuffixExpr() As Boolean
    Dim p171 As Long
    Dim q171 As Long
    Dim p161 As Long
    Dim q161 As Long

    With ctx
        p171 = .BufPos
        q171 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 1, ctx.UserData)
        If ParsePrimaryExpr() Then
            p161 = .BufPos
            q161 = .ThunkPos
            If ParseQUESTION() Then
                pvPushThunk ucsAct_1_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L22
            Else
                .BufPos = p161
                .ThunkPos = q161
            End If
            If ParseSTAR() Then
                pvPushThunk ucsAct_2_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L22
            Else
                .BufPos = p161
                .ThunkPos = q161
            End If
            If ParsePLUS() Then
                pvPushThunk ucsAct_3_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L22
            Else
                .BufPos = p161
                .ThunkPos = q161
            End If
            .BufPos = p161
            .ThunkPos = q161
L22:
            Call ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 2, ctx.UserData)
            ParseSuffixExpr = True
            Exit Function
        Else
            .BufPos = p171
            .ThunkPos = q171
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseNOT() As Boolean
    Dim p709 As Long
    Dim p702 As Long

    With ctx
        p709 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "NOT", 1, ctx.UserData)
        p702 = .BufPos
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
        .BufPos = p709
        If ConsoleTrace(ctx.BufPos + 1, "NOT", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParsePrimaryExpr() As Boolean
    Dim p239 As Long
    Dim q239 As Long
    Dim p200 As Long
    Dim q200 As Long
    Dim p178 As Long
    Dim q178 As Long
    Dim p189 As Long
    Dim q189 As Long
    Dim p182 As Long
    Dim q182 As Long
    Dim p185 As Long
    Dim q185 As Long
    Dim p199 As Long
    Dim q199 As Long
    Dim p209 As Long
    Dim q209 As Long
    Dim p218 As Long
    Dim q218 As Long

    With ctx
        p239 = .BufPos
        q239 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 1, ctx.UserData)
        p200 = .BufPos
        q200 = .ThunkPos
        If ParseIdentifier() Then
            pvPushThunk ucsAct_1_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p178 = .BufPos
            q178 = .ThunkPos
            If IsParamRule(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                If ParseCodeBlock() Then
                    pvPushThunk ucsAct_2_PrimaryExpr, .CaptureBegin, .CaptureEnd
                Else
                    .BufPos = p178
                    .ThunkPos = q178
                End If
            End If
            p189 = .BufPos
            q189 = .ThunkPos
            p182 = .BufPos
            q182 = .ThunkPos
            If Not (ParseCodeNC()) Then
                .BufPos = p182
                .ThunkPos = q182
            End If
            p185 = .BufPos
            q185 = .ThunkPos
            If Not (ParseLiteralNC()) Then
                .BufPos = p185
                .ThunkPos = q185
            End If
            If ParseAssignLiteral() Then
                .BufPos = p200
                .ThunkPos = q200
            Else
                .BufPos = p189
                .ThunkPos = q189
                Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                ParsePrimaryExpr = True
                Exit Function
            End If
        Else
            .BufPos = p200
            .ThunkPos = q200
        End If
        If ParseOPEN() Then
            If ClearErrors(ctx.BufPos + 1) Then
                p199 = .BufPos
                q199 = .ThunkPos
                Call ParseExpression
                If ParseCLOSE() Then
                    Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                    ParsePrimaryExpr = True
                    Exit Function
                Else
                    .BufPos = p199
                    .ThunkPos = q199
                End If
                If AppendError("Invalid expression", ctx.BufPos + 1) Then
                    Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                    ParsePrimaryExpr = True
                    Exit Function
                End If
                .BufPos = p200
                .ThunkPos = q200
            Else
                .BufPos = p200
                .ThunkPos = q200
            End If
        Else
            .BufPos = p200
            .ThunkPos = q200
        End If
        If ParseLiteralNoSpace() Then
            pvPushThunk ucsAct_3_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p209 = .BufPos
            q209 = .ThunkPos
            If ParseIgnoreCase() Then
                pvPushThunk ucsAct_4_PrimaryExpr, .CaptureBegin, .CaptureEnd
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                ParsePrimaryExpr = True
                Exit Function
            Else
                .BufPos = p209
                .ThunkPos = q209
            End If
            Call ParseEMPTY
            pvPushThunk ucsAct_4_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
            .BufPos = p200
            .ThunkPos = q200
        Else
            .BufPos = p200
            .ThunkPos = q200
        End If
        If ParseClassNoSpace() Then
            pvPushThunk ucsAct_5_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p218 = .BufPos
            q218 = .ThunkPos
            If ParseIgnoreCase() Then
                pvPushThunk ucsAct_6_PrimaryExpr, .CaptureBegin, .CaptureEnd
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                ParsePrimaryExpr = True
                Exit Function
            Else
                .BufPos = p218
                .ThunkPos = q218
            End If
            Call ParseEMPTY
            pvPushThunk ucsAct_6_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
            .BufPos = p200
            .ThunkPos = q200
        Else
            .BufPos = p200
            .ThunkPos = q200
        End If
        If ParseDOT() Then
            pvPushThunk ucsAct_7_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p200
            .ThunkPos = q200
        End If
        If ParseCodeBlock() Then
            pvPushThunk ucsAct_8_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p200
            .ThunkPos = q200
        End If
        If ParseBEGIN() Then
            pvPushThunk ucsAct_9_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p200
            .ThunkPos = q200
        End If
        If ParseEND() Then
            pvPushThunk ucsAct_10_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p200
            .ThunkPos = q200
        End If
        .BufPos = p239
        .ThunkPos = q239
        If ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseQUESTION() As Boolean
    Dim p722 As Long
    Dim p715 As Long

    With ctx
        p722 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "QUESTION", 1, ctx.UserData)
        p715 = .BufPos
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
        .BufPos = p722
        If ConsoleTrace(ctx.BufPos + 1, "QUESTION", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSTAR() As Boolean
    Dim p735 As Long
    Dim p728 As Long

    With ctx
        p735 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "STAR", 1, ctx.UserData)
        p728 = .BufPos
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
        .BufPos = p735
        If ConsoleTrace(ctx.BufPos + 1, "STAR", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParsePLUS() As Boolean
    Dim p748 As Long
    Dim p741 As Long

    With ctx
        p748 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "PLUS", 1, ctx.UserData)
        p741 = .BufPos
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
        .BufPos = p748
        If ConsoleTrace(ctx.BufPos + 1, "PLUS", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseCodeNC() As Boolean
    Dim p993 As Long
    Dim p972 As Long
    Dim p978 As Long
    Dim p987 As Long

    With ctx
        p993 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CodeNC", 1, ctx.UserData)
        p972 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """{""", 0, ctx.UserData)
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """{""", 2, ctx.UserData)
            Do
                p978 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^{}]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^{}]", 3, ctx.UserData) And False) Then
                        If Not (ParseCodeNC()) Then
                            .BufPos = p978
                            .BufPos = p978
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
                                .BufPos = p978
                                .BufPos = p978
                                Exit Do
                            End If
                        End If
                    End If
                End Select
            Loop
            p987 = .BufPos
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
            .BufPos = p993
            GoTo L28
        End If
        If ConsoleTrace(ctx.BufPos + 1, """{""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p993
L28:
        If ConsoleTrace(ctx.BufPos + 1, "CodeNC", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseLiteralNC() As Boolean
    Dim p435 As Long
    Dim p431 As Long
    Dim p387 As Long
    Dim p397 As Long
    Dim p395 As Long
    Dim p393 As Long
    Dim p404 As Long
    Dim p411 As Long
    Dim p421 As Long
    Dim p419 As Long
    Dim p417 As Long
    Dim p428 As Long

    With ctx
        p435 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 1, ctx.UserData)
        p431 = .BufPos
        p387 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
            Do
                p397 = .BufPos
                p395 = .BufPos
                p393 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
                If .BufData(.BufPos) = 39 Then      ' "'"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
                    .BufPos = p397
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p395
                If Not (ParseChar()) Then
                    .BufPos = p397
                    Exit Do
                End If
            Loop
            p404 = .BufPos
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
            .BufPos = p431
            GoTo L31
        End If
        If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p431
L31:
        p411 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
            Do
                p421 = .BufPos
                p419 = .BufPos
                p417 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
                If .BufData(.BufPos) = 34 Then      ' """
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
                    .BufPos = p421
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p419
                If Not (ParseChar()) Then
                    .BufPos = p421
                    Exit Do
                End If
            Loop
            p428 = .BufPos
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
            .BufPos = p431
            GoTo L32
        End If
        If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p431
L32:
        .BufPos = p435
        If ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseOPEN() As Boolean
    Dim p761 As Long
    Dim p754 As Long

    With ctx
        p761 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "OPEN", 1, ctx.UserData)
        p754 = .BufPos
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
        .BufPos = p761
        If ConsoleTrace(ctx.BufPos + 1, "OPEN", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCLOSE() As Boolean
    Dim p774 As Long
    Dim p767 As Long

    With ctx
        p774 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CLOSE", 1, ctx.UserData)
        p767 = .BufPos
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
        .BufPos = p774
        If ConsoleTrace(ctx.BufPos + 1, "CLOSE", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseLiteralNoSpace() As Boolean
    Dim p381 As Long
    Dim p377 As Long
    Dim p330 As Long
    Dim lCaptureBegin As Long
    Dim p343 As Long
    Dim p340 As Long
    Dim p338 As Long
    Dim lCaptureEnd As Long
    Dim p350 As Long
    Dim p356 As Long
    Dim p368 As Long
    Dim p366 As Long
    Dim p364 As Long
    Dim p375 As Long

    With ctx
        p381 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 1, ctx.UserData)
        p377 = .BufPos
        p330 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
            lCaptureBegin = .BufPos
            Do
                p343 = .BufPos
                p340 = .BufPos
                p338 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[']", 0, ctx.UserData)
                If .BufData(.BufPos) = 39 Then      ' "'"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[']", 2, ctx.UserData)
                    .BufPos = p343
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p340
                If Not (ParseChar()) Then
                    .BufPos = p343
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            p350 = .BufPos
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
            .BufPos = p377
            GoTo L33
        End If
        If ConsoleTrace(ctx.BufPos + 1, "[']", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p377
L33:
        p356 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
            lCaptureBegin = .BufPos
            Do
                p368 = .BufPos
                p366 = .BufPos
                p364 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[""]", 0, ctx.UserData)
                If .BufData(.BufPos) = 34 Then      ' """
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[""]", 2, ctx.UserData)
                    .BufPos = p368
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p366
                If Not (ParseChar()) Then
                    .BufPos = p368
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            p375 = .BufPos
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
            .BufPos = p377
            GoTo L34
        End If
        If ConsoleTrace(ctx.BufPos + 1, "[""]", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p377
L34:
        .BufPos = p381
        If ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseIgnoreCase() As Boolean
    Dim p449 As Long
    Dim lCaptureBegin As Long
    Dim p442 As Long
    Dim lCaptureEnd As Long

    With ctx
        p449 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "IgnoreCase", 1, ctx.UserData)
        lCaptureBegin = .BufPos
        p442 = .BufPos
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
        .BufPos = p449
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
    Dim p489 As Long
    Dim p464 As Long
    Dim lCaptureBegin As Long
    Dim p477 As Long
    Dim p474 As Long
    Dim p472 As Long
    Dim lCaptureEnd As Long
    Dim p484 As Long

    With ctx
        p489 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ClassNoSpace", 1, ctx.UserData)
        p464 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """[""", 0, ctx.UserData)
        If .BufData(.BufPos) = 91 Then              ' "["
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """[""", 2, ctx.UserData)
            lCaptureBegin = .BufPos
            Do
                p477 = .BufPos
                p474 = .BufPos
                p472 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, """]""", 0, ctx.UserData)
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, """]""", 2, ctx.UserData)
                    .BufPos = p477
                    Exit Do
                End If
                If ConsoleTrace(ctx.BufPos + 1, """]""", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p474
                If Not (ParseRange()) Then
                    .BufPos = p477
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            p484 = .BufPos
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
            .BufPos = p489
            GoTo L36
        End If
        If ConsoleTrace(ctx.BufPos + 1, """[""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p489
L36:
        If ConsoleTrace(ctx.BufPos + 1, "ClassNoSpace", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseDOT() As Boolean
    Dim p787 As Long
    Dim p780 As Long

    With ctx
        p787 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "DOT", 1, ctx.UserData)
        p780 = .BufPos
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
        .BufPos = p787
        If ConsoleTrace(ctx.BufPos + 1, "DOT", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseBEGIN() As Boolean
    Dim p1006 As Long
    Dim p999 As Long

    With ctx
        p1006 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "BEGIN", 1, ctx.UserData)
        p999 = .BufPos
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
        .BufPos = p1006
        If ConsoleTrace(ctx.BufPos + 1, "BEGIN", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseEND() As Boolean
    Dim p1019 As Long
    Dim p1012 As Long

    With ctx
        p1019 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "END", 1, ctx.UserData)
        p1012 = .BufPos
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
        .BufPos = p1019
        If ConsoleTrace(ctx.BufPos + 1, "END", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseSettingName() As Boolean
    Dim p268 As Long
    Dim p262 As Long

    With ctx
        p268 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 1, ctx.UserData)
        If ParseAT() Then
            p262 = .BufPos
            If ParseIdentifier() Then
                Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p262
            End If
            If ParseLiteral() Then
                Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p262
            End If
            Call ParseEMPTY
            Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
            ParseSettingName = True
            Exit Function
            .BufPos = p268
        Else
            .BufPos = p268
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SettingName", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseSettingValue() As Boolean
    Dim p280 As Long
    Dim p272 As Long
    Dim p274 As Long

    With ctx
        p280 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 1, ctx.UserData)
        p272 = .BufPos
        If ParseIdentifier() Then
            p274 = .BufPos
            If Not (ParseSEMI()) Then
                .BufPos = p274
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p272
        End If
        If ParseLiteral() Then
            p274 = .BufPos
            If Not (ParseSEMI()) Then
                .BufPos = p274
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p272
        End If
        If ParseCodeBlock() Then
            p274 = .BufPos
            If Not (ParseSEMI()) Then
                .BufPos = p274
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p272
        End If
        .BufPos = p280
        If ConsoleTrace(ctx.BufPos + 1, "SettingValue", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseAT() As Boolean
    Dim p1102 As Long
    Dim p1095 As Long

    With ctx
        p1102 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "AT", 1, ctx.UserData)
        p1095 = .BufPos
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
        .BufPos = p1102
        If ConsoleTrace(ctx.BufPos + 1, "AT", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseIdentStart() As Boolean
    Dim p305 As Long
    Dim p300 As Long

    With ctx
        p305 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "IdentStart", 1, ctx.UserData)
        p300 = .BufPos
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
        .BufPos = p305
        If ConsoleTrace(ctx.BufPos + 1, "IdentStart", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseIdentCont() As Boolean
    Dim p316 As Long
    Dim p311 As Long

    With ctx
        p316 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "IdentCont", 1, ctx.UserData)
        p311 = .BufPos
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
        .BufPos = p316
        If ConsoleTrace(ctx.BufPos + 1, "IdentCont", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseChar() As Boolean
    Dim p648 As Long
    Dim p552 As Long
    Dim p518 As Long
    Dim p524 As Long
    Dim p531 As Long
    Dim p537 As Long
    Dim p544 As Long
    Dim p550 As Long
    Dim p557 As Long
    Dim p563 As Long
    Dim p570 As Long
    Dim p577 As Long
    Dim p583 As Long
    Dim p590 As Long
    Dim p596 As Long
    Dim p602 As Long
    Dim p609 As Long
    Dim p615 As Long
    Dim p621 As Long
    Dim p627 As Long
    Dim p633 As Long
    Dim p642 As Long
    Dim p640 As Long

    With ctx
        p648 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Char", 1, ctx.UserData)
        p552 = .BufPos
        p518 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            p524 = .BufPos
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
            .BufPos = p552
            GoTo L41
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p552
L41:
        p531 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            p537 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[0-3]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 48 To 51                           ' [0-3]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[0-3]", 2, ctx.UserData)
                p544 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 2, ctx.UserData)
                    p550 = .BufPos
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
                    .BufPos = p552
                    GoTo L47
                End Select
                If ConsoleTrace(ctx.BufPos + 1, "[0-7]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p552
L47:
                GoTo L48
            End Select
            If ConsoleTrace(ctx.BufPos + 1, "[0-3]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p552
L48:
            GoTo L46
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p552
L46:
        p557 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            p563 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 48 To 55                           ' [0-7]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 2, ctx.UserData)
                p570 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-7]", 2, ctx.UserData)
                    GoTo L55
                End Select
                If ConsoleTrace(ctx.BufPos + 1, "[0-7]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p570
L55:
                Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                Call pvSetAdvance
                ParseChar = True
                Exit Function
            End Select
            If ConsoleTrace(ctx.BufPos + 1, "[0-7]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p552
            GoTo L54
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p552
L54:
        p577 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\x""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 120 Then ' "\\x"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """\\x""", 2, ctx.UserData)
            p583 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                p590 = .BufPos
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
                .BufPos = p552
                GoTo L59
            End Select
            If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p552
L59:
            GoTo L58
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\x""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p552
L58:
        p596 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\u""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 117 Then ' "\\u"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """\\u""", 2, ctx.UserData)
            p602 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                p609 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 48 To 57, 97 To 102, 65 To 70  ' [0-9a-fA-F]
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                    p615 = .BufPos
                    Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 0, ctx.UserData)
                    Select Case .BufData(.BufPos)
                    Case 48 To 57, 97 To 102, 65 To 70 ' [0-9a-fA-F]
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 2, ctx.UserData)
                        p621 = .BufPos
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
                        .BufPos = p552
                        GoTo L71
                    End Select
                    If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                        '--- do nothing
                    End If
                    .BufPos = p552
L71:
                    GoTo L72
                End Select
                If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p552
L72:
                GoTo L73
            End Select
            If ConsoleTrace(ctx.BufPos + 1, "[0-9a-fA-F]", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p552
L73:
            GoTo L70
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\u""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p552
L70:
        p627 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            p633 = .BufPos
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
            .BufPos = p552
            GoTo L77
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p552
L77:
        p642 = .BufPos
        p640 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """\\""", 0, ctx.UserData)
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """\\""", 2, ctx.UserData)
            .BufPos = p552
            GoTo L78
        End If
        If ConsoleTrace(ctx.BufPos + 1, """\\""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p642
        If .BufPos < .BufSize Then
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
            Call pvSetAdvance
            ParseChar = True
            Exit Function
        Else
            .BufPos = p552
        End If
L78:
        .BufPos = p648
        If ConsoleTrace(ctx.BufPos + 1, "Char", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        Call pvSetExpected("char")
    End With
End Function

Public Function ParseClass() As Boolean
    Dim p458 As Long

    With ctx
        p458 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Class", 1, ctx.UserData)
        If ParseClassNoSpace() Then
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "Class", 2, ctx.UserData)
            ParseClass = True
            Exit Function
        Else
            .BufPos = p458
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Class", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseRange() As Boolean
    Dim p512 As Long
    Dim p508 As Long
    Dim p496 As Long
    Dim p505 As Long
    Dim p503 As Long

    With ctx
        p512 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Range", 1, ctx.UserData)
        p508 = .BufPos
        If ParseChar() Then
            p496 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, """-""", 0, ctx.UserData)
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, """-""", 2, ctx.UserData)
                p505 = .BufPos
                p503 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, """]""", 0, ctx.UserData)
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, """]""", 2, ctx.UserData)
                    .BufPos = p508
                    GoTo L81
                End If
                If ConsoleTrace(ctx.BufPos + 1, """]""", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p505
                If ParseChar() Then
                    Call ConsoleTrace(ctx.BufPos + 1, "Range", 2, ctx.UserData)
                    Call pvSetAdvance
                    ParseRange = True
                    Exit Function
                Else
                    .BufPos = p508
                End If
L81:
                GoTo L80
            End If
            If ConsoleTrace(ctx.BufPos + 1, """-""", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p508
L80:
        Else
            .BufPos = p508
        End If
        If ParseChar() Then
            Call ConsoleTrace(ctx.BufPos + 1, "Range", 2, ctx.UserData)
            Call pvSetAdvance
            ParseRange = True
            Exit Function
        Else
            .BufPos = p508
        End If
        .BufPos = p512
        If ConsoleTrace(ctx.BufPos + 1, "Range", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseLEFTARROW() As Boolean
    Dim p662 As Long
    Dim p655 As Long

    With ctx
        p662 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LEFTARROW", 1, ctx.UserData)
        p655 = .BufPos
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
        .BufPos = p662
        If ConsoleTrace(ctx.BufPos + 1, "LEFTARROW", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSLASH() As Boolean
    Dim p683 As Long
    Dim p669 As Long
    Dim p677 As Long
    Dim p675 As Long

    With ctx
        p683 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SLASH", 1, ctx.UserData)
        p669 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """/""", 0, ctx.UserData)
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """/""", 2, ctx.UserData)
            p677 = .BufPos
            p675 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, """/""", 0, ctx.UserData)
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, """/""", 2, ctx.UserData)
                .BufPos = p683
                GoTo L85
            End If
            If ConsoleTrace(ctx.BufPos + 1, """/""", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p677
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "SLASH", 2, ctx.UserData)
            Call pvSetAdvance
            ParseSLASH = True
            Exit Function
L85:
            GoTo L84
        End If
        If ConsoleTrace(ctx.BufPos + 1, """/""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p683
L84:
        If ConsoleTrace(ctx.BufPos + 1, "SLASH", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseSpace() As Boolean
    Dim p897 As Long
    Dim p885 As Long
    Dim p891 As Long

    With ctx
        p897 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Space", 1, ctx.UserData)
        p885 = .BufPos
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
        p891 = .BufPos
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
        .BufPos = p885
        If ParseEndOfLine() Then
            Call ConsoleTrace(ctx.BufPos + 1, "Space", 2, ctx.UserData)
            Call pvSetAdvance
            ParseSpace = True
            Exit Function
        Else
            .BufPos = p885
        End If
        .BufPos = p897
        If ConsoleTrace(ctx.BufPos + 1, "Space", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseLineComment() As Boolean
    Dim p844 As Long
    Dim p821 As Long
    Dim p833 As Long
    Dim p840 As Long
    Dim p827 As Long

    With ctx
        p844 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 1, ctx.UserData)
        p821 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """#""", 0, ctx.UserData)
        If .BufData(.BufPos) = 35 Then              ' "#"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """#""", 2, ctx.UserData)
            Do
                p833 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 13, 10                         ' [\r\n]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 3, ctx.UserData) And False) Then
                        .BufPos = p833
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 2, ctx.UserData)
                    Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 3, ctx.UserData) And False) Then
                            .BufPos = p833
                            Exit Do
                        End If
                    End If
                End Select
            Loop
            p840 = .BufPos
            If ParseEndOfLine() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p840
            End If
            If ParseEndOfFile() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p840
            End If
            .BufPos = p844
            GoTo L88
        End If
        If ConsoleTrace(ctx.BufPos + 1, """#""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        p827 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """//""", 0, ctx.UserData)
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 47 Then ' "//"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """//""", 2, ctx.UserData)
            Do
                p833 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 13, 10                         ' [\r\n]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 3, ctx.UserData) And False) Then
                        .BufPos = p833
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 2, ctx.UserData)
                    Else
                        If Not (ConsoleTrace(ctx.BufPos + 1, "[^\r\n]", 3, ctx.UserData) And False) Then
                            .BufPos = p833
                            Exit Do
                        End If
                    End If
                End Select
            Loop
            p840 = .BufPos
            If ParseEndOfLine() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p840
            End If
            If ParseEndOfFile() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p840
            End If
            .BufPos = p844
            GoTo L88
        End If
        If ConsoleTrace(ctx.BufPos + 1, """//""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p821
        .BufPos = p844
L88:
        If ConsoleTrace(ctx.BufPos + 1, "LineComment", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseBlockComment() As Boolean
    Dim p879 As Long
    Dim p850 As Long
    Dim p856 As Long
    Dim p865 As Long
    Dim p863 As Long
    Dim p874 As Long

    With ctx
        p879 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "BlockComment", 1, ctx.UserData)
        p850 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, """/*""", 0, ctx.UserData)
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 42 Then ' "/*"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, """/*""", 2, ctx.UserData)
            Do
                p856 = .BufPos
                Call ConsoleTrace(ctx.BufPos + 1, "[^*/]", 0, ctx.UserData)
                Select Case .BufData(.BufPos)
                Case 42, 47                         ' [*/]
                    If Not (ConsoleTrace(ctx.BufPos + 1, "[^*/]", 3, ctx.UserData) And False) Then
                        If Not (ParseBlockComment()) Then
                            .BufPos = p856
                            p865 = .BufPos
                            p863 = .BufPos
                            Call ConsoleTrace(ctx.BufPos + 1, """*/""", 0, ctx.UserData)
                            If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                                .BufPos = .BufPos + 2
                                Call ConsoleTrace(ctx.BufPos + 1, """*/""", 2, ctx.UserData)
                                .BufPos = p856
                                .BufPos = p856
                                Exit Do
                            End If
                            If ConsoleTrace(ctx.BufPos + 1, """*/""", 3, ctx.UserData) And False Then
                                '--- do nothing
                            End If
                            .BufPos = p865
                            If .BufPos < .BufSize Then
                                .BufPos = .BufPos + 1
                            Else
                                .BufPos = p856
                                .BufPos = p856
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
                                .BufPos = p856
                                p865 = .BufPos
                                p863 = .BufPos
                                Call ConsoleTrace(ctx.BufPos + 1, """*/""", 0, ctx.UserData)
                                If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                                    .BufPos = .BufPos + 2
                                    Call ConsoleTrace(ctx.BufPos + 1, """*/""", 2, ctx.UserData)
                                    .BufPos = p856
                                    .BufPos = p856
                                    Exit Do
                                End If
                                If ConsoleTrace(ctx.BufPos + 1, """*/""", 3, ctx.UserData) And False Then
                                    '--- do nothing
                                End If
                                .BufPos = p865
                                If .BufPos < .BufSize Then
                                    .BufPos = .BufPos + 1
                                Else
                                    .BufPos = p856
                                    .BufPos = p856
                                    Exit Do
                                End If
                            End If
                        End If
                    End If
                End Select
            Loop
            p874 = .BufPos
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
            .BufPos = p879
            GoTo L95
        End If
        If ConsoleTrace(ctx.BufPos + 1, """/*""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p879
L95:
        If ConsoleTrace(ctx.BufPos + 1, "BlockComment", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseEndOfLine() As Boolean
    Dim p920 As Long
    Dim p903 As Long
    Dim p909 As Long
    Dim p915 As Long

    With ctx
        p920 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 1, ctx.UserData)
        p903 = .BufPos
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
        p909 = .BufPos
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
        .BufPos = p903
        p915 = .BufPos
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
        .BufPos = p903
        .BufPos = p920
        If ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseEQUAL() As Boolean
    Dim p1050 As Long
    Dim p1043 As Long

    With ctx
        p1050 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "EQUAL", 1, ctx.UserData)
        p1043 = .BufPos
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
        .BufPos = p1050
        If ConsoleTrace(ctx.BufPos + 1, "EQUAL", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseBAR() As Boolean
    Dim p1063 As Long
    Dim p1056 As Long

    With ctx
        p1063 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "BAR", 1, ctx.UserData)
        p1056 = .BufPos
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
        .BufPos = p1063
        If ConsoleTrace(ctx.BufPos + 1, "BAR", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParsePrepExpr() As Boolean
    Dim p1159 As Long
    Dim p1153 As Long
    Dim p1150 As Long
    Dim p1144 As Long
    Dim p1146 As Long

    With ctx
        p1159 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 1, ctx.UserData)
        p1153 = .BufPos
        If ParseIdentifier() Then
            p1150 = .BufPos
            p1144 = .BufPos
            If Not (ParseCodeNC()) Then
                .BufPos = p1144
            End If
            p1146 = .BufPos
            If Not (ParseLiteralNC()) Then
                .BufPos = p1146
            End If
            If ParseAssignLiteral() Then
                .BufPos = p1153
            Else
                .BufPos = p1150
                Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
                ParsePrepExpr = True
                Exit Function
            End If
        Else
            .BufPos = p1153
        End If
        If ParseLiteral() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1153
        End If
        If ParseClass() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1153
        End If
        If ParseCodeNC() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrepExpr", 2, ctx.UserData)
            ParsePrepExpr = True
            Exit Function
        Else
            .BufPos = p1153
        End If
        .BufPos = p1159
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
