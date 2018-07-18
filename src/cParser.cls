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
' Auto-generated on 18.7.2018 17:33:20
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
    ucsAct_3_Rule
    ucsAct_2_Rule
    ucsAct_1_Rule
    ucsAct_1_ChoiceExpr
    ucsAct_1_SequenceExpr
    ucsAct_2_SequenceExpr
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
    ucsAct_1_PrimaryExpr
    ucsAct_3_PrimaryExpr
    ucsAct_2_PrimaryExpr
    ucsAct_5_PrimaryExpr
    ucsAct_4_PrimaryExpr
    ucsAct_6_PrimaryExpr
    ucsAct_7_PrimaryExpr
    ucsAct_8_PrimaryExpr
    ucsAct_9_PrimaryExpr
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
    LastError           As String
    LastBufPos          As Long
    UserData            As cTree
    VarResult           As Variant
    '-- from @members setting
    ParseErrors         As Collection
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
    ParserVersion = "18.7.2018 17:33:20"
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Match(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As cTree) As Long
    If BeginMatch(sSubject, StartPos, UserData) Then
        If ParseGrammar() Then
            Match = EndMatch()
        ElseIf LenB(ctx.LastError) = 0 Then
            ctx.LastError = "Fail"
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

Public Function EndMatch() As Long
    Dim lIdx            As Long
    
    With ctx
        For lIdx = 0 To .ThunkPos - 1
            With .ThunkData(lIdx)
                pvImplAction .Action, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin
            End With
        Next
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

'= generated functions ===================================================

Public Function ParseGrammar() As Boolean
    Dim p17 As Long
    Dim q17 As Long
    Dim i10 As Long
    Dim p9 As Long
    Dim q9 As Long

    With ctx
        p17 = .BufPos
        q17 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Grammar", 1, ctx.UserData)
        Call ParseSpacing
        For i10 = 0 To LNG_MAXINT
            p9 = .BufPos
            q9 = .ThunkPos
            If ParseSetting() Then
                GoTo L1
            Else
                .BufPos = p9
                .ThunkPos = q9
                If ParseRule() Then
                    GoTo L1
                Else
                    .BufPos = p9
                    .ThunkPos = q9
                    Exit For
                End If
            End If
L1:
        Next
        If i10 <> 0 Then
            If ParseEndOfFile() Then
                Call ConsoleTrace(ctx.BufPos + 1, "Grammar", 2, ctx.UserData)
                ParseGrammar = True
                Exit Function
            End If
            If AppendError("Expected rule or setting", ctx.BufPos + 1) Then
                Call ConsoleTrace(ctx.BufPos + 1, "Grammar", 2, ctx.UserData)
                ParseGrammar = True
                Exit Function
            End If
            .BufPos = p17
            .ThunkPos = q17
        Else
            .BufPos = p17
            .ThunkPos = q17
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Grammar", 3, ctx.UserData) And False Then
            ParseGrammar = True
        End If
    End With
End Function

Public Sub ParseSpacing()
    Call ConsoleTrace(ctx.BufPos + 1, "Spacing", 1, ctx.UserData)
    Do
        If ParseSpace() Then
            GoTo L3
        Else
            If ParseComment() Then
                GoTo L3
            Else
                Exit Do
            End If
        End If
L3:
    Loop
    Call ConsoleTrace(ctx.BufPos + 1, "Spacing", 2, ctx.UserData)
End Sub

Public Function ParseSetting() As Boolean
    Dim p234 As Long
    Dim q234 As Long
    Dim p230 As Long
    Dim q230 As Long

    With ctx
        p234 = .BufPos
        q234 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 1, ctx.UserData)
        If ParseSettingName() Then
            pvPushThunk ucsAct_1_Setting, .CaptureBegin, .CaptureEnd
            If ClearErrors(ctx.BufPos + 1) And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                p230 = .BufPos
                q230 = .ThunkPos
                If ParseSettingValue() Then
                    pvPushThunk ucsAct_2_Setting, .CaptureBegin, .CaptureEnd
                    Call ConsoleTrace(ctx.BufPos + 1, "Setting", 2, ctx.UserData)
                    ParseSetting = True
                    Exit Function
                Else
                    .BufPos = p230
                    .ThunkPos = q230
                End If
                If AppendError("Invalid value for setting '" & .UserData.Tag & "'", ctx.BufPos + 1) Then
                    Call ConsoleTrace(ctx.BufPos + 1, "Setting", 2, ctx.UserData)
                    ParseSetting = True
                    Exit Function
                End If
                .BufPos = p234
                .ThunkPos = q234
            Else
                .BufPos = p234
                .ThunkPos = q234
            End If
        Else
            .BufPos = p234
            .ThunkPos = q234
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Setting", 3, ctx.UserData) And False Then
            ParseSetting = True
        End If
    End With
End Function

Public Function ParseRule() As Boolean
    Dim p43 As Long
    Dim q43 As Long
    Dim p29 As Long
    Dim q29 As Long

    With ctx
        p43 = .BufPos
        q43 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Rule", 1, ctx.UserData)
        If ParseIdentifier() Then
            pvPushThunk ucsAct_1_Rule, .CaptureBegin, .CaptureEnd
            If ClearErrors(ctx.BufPos + 1) And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                p29 = .BufPos
                q29 = .ThunkPos
                If ParseLiteral() Then
                    pvPushThunk ucsAct_2_Rule, .CaptureBegin, .CaptureEnd
                    GoTo L6
                Else
                    .BufPos = p29
                    .ThunkPos = q29
                End If
L6:
                If ParseAssignLiteral() Then
                    Call ParseExpression
                    pvPushThunk ucsAct_3_Rule, .CaptureBegin, .CaptureEnd
                    Call ParseSEMI
                    Call ConsoleTrace(ctx.BufPos + 1, "Rule", 2, ctx.UserData)
                    ParseRule = True
                    Exit Function
                Else
                    .BufPos = p43
                    .ThunkPos = q43
                End If
            Else
                .BufPos = p43
                .ThunkPos = q43
            End If
        Else
            .BufPos = p43
            .ThunkPos = q43
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Rule", 3, ctx.UserData) And False Then
            ParseRule = True
        End If
    End With
End Function

Public Function ParseEndOfFile() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "EndOfFile", 1, ctx.UserData)
        If Not .BufPos < .BufSize Then
            Call ConsoleTrace(ctx.BufPos + 1, "EndOfFile", 2, ctx.UserData)
            ParseEndOfFile = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "EndOfFile", 3, ctx.UserData) And False Then
            ParseEndOfFile = True
        End If
    End With
End Function

Public Function ParseIdentifier() As Boolean
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "Identifier", 1, ctx.UserData)
        lCaptureBegin = .BufPos
        If ParseIdentStart() Then
            Do
                If Not ParseIdentCont() Then
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            Call ParseSpacing
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "Identifier", 2, ctx.UserData)
            ParseIdentifier = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Identifier", 3, ctx.UserData) And False Then
            ParseIdentifier = True
        End If
    End With
End Function

Public Function ParseLiteral() As Boolean
    Dim p294 As Long
    Dim q294 As Long

    With ctx
        p294 = .BufPos
        q294 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Literal", 1, ctx.UserData)
        If ParseLiteralNoSpace() Then
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "Literal", 2, ctx.UserData)
            ParseLiteral = True
            Exit Function
        Else
            .BufPos = p294
            .ThunkPos = q294
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Literal", 3, ctx.UserData) And False Then
            ParseLiteral = True
        End If
    End With
End Function

Public Function ParseAssignLiteral() As Boolean
    Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 1, ctx.UserData)
    If ParseLEFTARROW() Then
        Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 2, ctx.UserData)
        ParseAssignLiteral = True
        Exit Function
    End If
    If ParseEQUAL() Then
        Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 2, ctx.UserData)
        ParseAssignLiteral = True
        Exit Function
    End If
    If ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 3, ctx.UserData) And False Then
        ParseAssignLiteral = True
    End If
End Function

Public Sub ParseExpression()
    Call ConsoleTrace(ctx.BufPos + 1, "Expression", 1, ctx.UserData)
    Call ParseChoiceExpr
    Call ConsoleTrace(ctx.BufPos + 1, "Expression", 2, ctx.UserData)
End Sub

Private Function ParseSEMI() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "SEMI", 1, ctx.UserData)
        If .BufData(.BufPos) = 59 Then              ' ";"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "SEMI", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseSEMI = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SEMI", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseSEMI = True
        End If
    End With
End Function

Public Sub ParseChoiceExpr()
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceExpr", 1, ctx.UserData)
        Call ParseSequenceExpr
        Do
            If Not ParseChoiceLiteral() Then
                Exit Do
            End If
            Call ParseSequenceExpr
            pvPushThunk ucsAct_1_ChoiceExpr, .CaptureBegin, .CaptureEnd
        Loop
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceExpr", 2, ctx.UserData)
    End With
End Sub

Public Sub ParseSequenceExpr()
    Dim p74 As Long
    Dim q74 As Long
    Dim p72 As Long
    Dim q72 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "SequenceExpr", 1, ctx.UserData)
        p74 = .BufPos
        q74 = .ThunkPos
        If ParseErrorExpr() Then
            Do
                p72 = .BufPos
                q72 = .ThunkPos
                If Not ParseErrorExpr() Then
                    .BufPos = p72
                    .ThunkPos = q72
                    Exit Do
                End If
                pvPushThunk ucsAct_1_SequenceExpr, .CaptureBegin, .CaptureEnd
            Loop
            Call ConsoleTrace(ctx.BufPos + 1, "SequenceExpr", 2, ctx.UserData)
            GoTo L9
        Else
            .BufPos = p74
            .ThunkPos = q74
        End If
        pvPushThunk ucsAct_2_SequenceExpr, .CaptureBegin, .CaptureEnd
        Call ConsoleTrace(ctx.BufPos + 1, "SequenceExpr", 2, ctx.UserData)
        GoTo L9
L9:
    End With
End Sub

Public Function ParseChoiceLiteral() As Boolean
    Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 1, ctx.UserData)
    If ParseSLASH() Then
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 2, ctx.UserData)
        ParseChoiceLiteral = True
        Exit Function
    End If
    If ParseBAR() Then
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 2, ctx.UserData)
        ParseChoiceLiteral = True
        Exit Function
    End If
    If ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 3, ctx.UserData) And False Then
        ParseChoiceLiteral = True
    End If
End Function

Public Function ParseErrorExpr() As Boolean
    Dim p99 As Long
    Dim q99 As Long
    Dim p87 As Long
    Dim q87 As Long
    Dim p90 As Long
    Dim q90 As Long
    Dim p93 As Long
    Dim q93 As Long

    With ctx
        p99 = .BufPos
        q99 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 1, ctx.UserData)
        If ParseVariableExpr() Then
            p87 = .BufPos
            q87 = .ThunkPos
            If ParseCARET() Then
                If ClearErrors(ctx.BufPos + 1) Then
                    p90 = .BufPos
                    q90 = .ThunkPos
                    If ParseIdentifier() Then
                        pvPushThunk ucsAct_1_ErrorExpr, .CaptureBegin, .CaptureEnd
                        GoTo L10
                    Else
                        .BufPos = p90
                        .ThunkPos = q90
                    End If
                    p93 = .BufPos
                    q93 = .ThunkPos
                    If ParseLiteral() Then
                        pvPushThunk ucsAct_2_ErrorExpr, .CaptureBegin, .CaptureEnd
                        GoTo L10
                    Else
                        .BufPos = p93
                        .ThunkPos = q93
                    End If
                    If AppendError("Expected rule or literal for error label", ctx.BufPos + 1) Then
                        GoTo L10
                    End If
                    .BufPos = p87
                    .ThunkPos = q87
                Else
                    .BufPos = p87
                    .ThunkPos = q87
                End If
            End If
L10:
            Call ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 2, ctx.UserData)
            ParseErrorExpr = True
            Exit Function
        Else
            .BufPos = p99
            .ThunkPos = q99
        End If
        If ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 3, ctx.UserData) And False Then
            ParseErrorExpr = True
        End If
    End With
End Function

Public Function ParseVariableExpr() As Boolean
    Dim p112 As Long
    Dim q112 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 1, ctx.UserData)
        p112 = .BufPos
        q112 = .ThunkPos
        If ParseIdentifier() Then
            pvPushThunk ucsAct_1_VariableExpr, .CaptureBegin, .CaptureEnd
            If ParseCOLON() Then
                If ParsePrefixExpr() Then
                    pvPushThunk ucsAct_2_VariableExpr, .CaptureBegin, .CaptureEnd
                    Call ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 2, ctx.UserData)
                    ParseVariableExpr = True
                    Exit Function
                Else
                    .BufPos = p112
                    .ThunkPos = q112
                End If
            Else
                .BufPos = p112
                .ThunkPos = q112
            End If
        Else
            .BufPos = p112
            .ThunkPos = q112
        End If
        If ParsePrefixExpr() Then
            Call ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 2, ctx.UserData)
            ParseVariableExpr = True
            Exit Function
        Else
            .BufPos = p112
            .ThunkPos = q112
        End If
        If ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 3, ctx.UserData) And False Then
            ParseVariableExpr = True
        End If
    End With
End Function

Private Function ParseCARET() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "CARET", 1, ctx.UserData)
        If .BufData(.BufPos) = 94 Then              ' "^"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "CARET", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseCARET = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "CARET", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseCARET = True
        End If
    End With
End Function

Private Function ParseCOLON() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "COLON", 1, ctx.UserData)
        If .BufData(.BufPos) = 58 Then              ' ":"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "COLON", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseCOLON = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "COLON", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseCOLON = True
        End If
    End With
End Function

Public Function ParsePrefixExpr() As Boolean
    Dim p129 As Long
    Dim q129 As Long
    Dim p127 As Long
    Dim q127 As Long
    Dim p133 As Long
    Dim q133 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 1, ctx.UserData)
        p129 = .BufPos
        q129 = .ThunkPos
        If ParseAND() Then
            If ParseCodeBlock() Then
                pvPushThunk ucsAct_1_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p129
                .ThunkPos = q129
            End If
        End If
        p127 = .BufPos
        q127 = .ThunkPos
        If ParseAND() Then
            If ParseSuffixExpr() Then
                pvPushThunk ucsAct_2_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p127
                .ThunkPos = q127
            End If
        End If
        p133 = .BufPos
        q133 = .ThunkPos
        If ParseNOT() Then
            If ParseSuffixExpr() Then
                pvPushThunk ucsAct_3_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p133
                .ThunkPos = q133
            End If
        End If
        If ParseSuffixExpr() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
            ParsePrefixExpr = True
            Exit Function
        Else
            .BufPos = p129
            .ThunkPos = q129
        End If
        If ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 3, ctx.UserData) And False Then
            ParsePrefixExpr = True
        End If
    End With
End Function

Private Function ParseAND() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "AND", 1, ctx.UserData)
        If .BufData(.BufPos) = 38 Then              ' "&"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "AND", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseAND = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "AND", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseAND = True
        End If
    End With
End Function

Public Function ParseCodeBlock() As Boolean
    Dim p589 As Long
    Dim q589 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        p589 = .BufPos
        q589 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "CodeBlock", 1, ctx.UserData)
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Do
                If Not ParseEndOfLine() Then
                    Exit Do
                End If
            Loop
            lCaptureBegin = .BufPos
            Do
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If ParseCodeNC() Then
                        GoTo L11
                    Else
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        GoTo L11
                    Else
                        If ParseCodeNC() Then
                            GoTo L11
                        Else
                            Exit Do
                        End If
                    End If
                End Select
L11:
            Loop
            lCaptureEnd = .BufPos
            If .BufData(.BufPos) = 125 Then         ' "}"
                .BufPos = .BufPos + 1
                Call ParseSpacing
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "CodeBlock", 2, ctx.UserData)
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString: .LastBufPos = .BufPos
                End If
                ParseCodeBlock = True
                Exit Function
            Else
                .BufPos = p589
                .ThunkPos = q589
            End If
        End If
        If ConsoleTrace(ctx.BufPos + 1, "CodeBlock", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseCodeBlock = True
        End If
    End With
End Function

Public Function ParseSuffixExpr() As Boolean
    Dim p159 As Long
    Dim q159 As Long

    With ctx
        p159 = .BufPos
        q159 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 1, ctx.UserData)
        If ParsePrimaryExpr() Then
            If ParseQUESTION() Then
                pvPushThunk ucsAct_1_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L12
            End If
            If ParseSTAR() Then
                pvPushThunk ucsAct_2_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L12
            End If
            If ParsePLUS() Then
                pvPushThunk ucsAct_3_SuffixExpr, .CaptureBegin, .CaptureEnd
            End If
L12:
            Call ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 2, ctx.UserData)
            ParseSuffixExpr = True
            Exit Function
        Else
            .BufPos = p159
            .ThunkPos = q159
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 3, ctx.UserData) And False Then
            ParseSuffixExpr = True
        End If
    End With
End Function

Private Function ParseNOT() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "NOT", 1, ctx.UserData)
        If .BufData(.BufPos) = 33 Then              ' "!"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "NOT", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseNOT = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "NOT", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseNOT = True
        End If
    End With
End Function

Public Function ParsePrimaryExpr() As Boolean
    Dim p170 As Long
    Dim q170 As Long
    Dim p168 As Long
    Dim q168 As Long
    Dim p175 As Long
    Dim q175 As Long
    Dim p179 As Long
    Dim q179 As Long
    Dim p186 As Long
    Dim q186 As Long
    Dim p190 As Long
    Dim q190 As Long
    Dim p196 As Long
    Dim q196 As Long
    Dim p199 As Long
    Dim q199 As Long
    Dim p208 As Long
    Dim q208 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 1, ctx.UserData)
        p170 = .BufPos
        q170 = .ThunkPos
        If ParseIdentifier() Then
            p168 = .BufPos
            q168 = .ThunkPos
            Call ParseLiteralNC
            If ParseAssignLiteral() Then
                .BufPos = p170
                .ThunkPos = q170
                GoTo L13
            Else
                .BufPos = p168
                .ThunkPos = q168
                pvPushThunk ucsAct_1_PrimaryExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                ParsePrimaryExpr = True
                Exit Function
            End If
L13:
        Else
            .BufPos = p170
            .ThunkPos = q170
        End If
        p175 = .BufPos
        q175 = .ThunkPos
        If ParseOPEN() Then
            If ClearErrors(ctx.BufPos + 1) Then
                p179 = .BufPos
                q179 = .ThunkPos
                Call ParseExpression
                If ParseCLOSE() Then
                    Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                    ParsePrimaryExpr = True
                    Exit Function
                Else
                    .BufPos = p179
                    .ThunkPos = q179
                End If
                If AppendError("Invalid expression", ctx.BufPos + 1) Then
                    Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                    ParsePrimaryExpr = True
                    Exit Function
                End If
                .BufPos = p175
                .ThunkPos = q175
            Else
                .BufPos = p175
                .ThunkPos = q175
            End If
        End If
        p186 = .BufPos
        q186 = .ThunkPos
        If ParseLiteralNoSpace() Then
            pvPushThunk ucsAct_2_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p190 = .BufPos
            q190 = .ThunkPos
            If ParseIgnoreCase() Then
                pvPushThunk ucsAct_3_PrimaryExpr, .CaptureBegin, .CaptureEnd
                GoTo L15
            Else
                .BufPos = p190
                .ThunkPos = q190
            End If
L15:
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p186
            .ThunkPos = q186
        End If
        p196 = .BufPos
        q196 = .ThunkPos
        If ParseClassNoSpace() Then
            pvPushThunk ucsAct_4_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p199 = .BufPos
            q199 = .ThunkPos
            If ParseIgnoreCase() Then
                pvPushThunk ucsAct_5_PrimaryExpr, .CaptureBegin, .CaptureEnd
                GoTo L16
            Else
                .BufPos = p199
                .ThunkPos = q199
            End If
L16:
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p196
            .ThunkPos = q196
        End If
        If ParseDOT() Then
            pvPushThunk ucsAct_6_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        End If
        p208 = .BufPos
        q208 = .ThunkPos
        If ParseCodeBlock() Then
            pvPushThunk ucsAct_7_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p208
            .ThunkPos = q208
        End If
        If ParseBEGIN() Then
            pvPushThunk ucsAct_8_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        End If
        If ParseEND() Then
            pvPushThunk ucsAct_9_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 3, ctx.UserData) And False Then
            ParsePrimaryExpr = True
        End If
    End With
End Function

Private Function ParseQUESTION() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "QUESTION", 1, ctx.UserData)
        If .BufData(.BufPos) = 63 Then              ' "?"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "QUESTION", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseQUESTION = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "QUESTION", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseQUESTION = True
        End If
    End With
End Function

Private Function ParseSTAR() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "STAR", 1, ctx.UserData)
        If .BufData(.BufPos) = 42 Then              ' "*"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "STAR", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseSTAR = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "STAR", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseSTAR = True
        End If
    End With
End Function

Private Function ParsePLUS() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "PLUS", 1, ctx.UserData)
        If .BufData(.BufPos) = 43 Then              ' "+"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "PLUS", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParsePLUS = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "PLUS", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParsePLUS = True
        End If
    End With
End Function

Public Function ParseLiteralNC() As Boolean
    Dim p333 As Long
    Dim p342 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 1, ctx.UserData)
        p333 = .BufPos
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            Do
                If .BufData(.BufPos) = 39 Then      ' "'"
                    Exit Do
                End If
                If Not ParseChar() Then
                    Exit Do
                End If
            Loop
            If .BufData(.BufPos) = 39 Then          ' "'"
                .BufPos = .BufPos + 1
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 2, ctx.UserData)
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString: .LastBufPos = .BufPos
                End If
                ParseLiteralNC = True
                Exit Function
            Else
                .BufPos = p333
            End If
        End If
        p342 = .BufPos
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            Do
                If .BufData(.BufPos) = 34 Then      ' """
                    Exit Do
                End If
                If Not ParseChar() Then
                    Exit Do
                End If
            Loop
            If .BufData(.BufPos) = 34 Then          ' """
                .BufPos = .BufPos + 1
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 2, ctx.UserData)
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString: .LastBufPos = .BufPos
                End If
                ParseLiteralNC = True
                Exit Function
            Else
                .BufPos = p342
            End If
        End If
        If ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseLiteralNC = True
        End If
    End With
End Function

Private Function ParseOPEN() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "OPEN", 1, ctx.UserData)
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "OPEN", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseOPEN = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "OPEN", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseOPEN = True
        End If
    End With
End Function

Private Function ParseCLOSE() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "CLOSE", 1, ctx.UserData)
        If .BufData(.BufPos) = 41 Then              ' ")"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "CLOSE", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseCLOSE = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "CLOSE", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseCLOSE = True
        End If
    End With
End Function

Public Function ParseLiteralNoSpace() As Boolean
    Dim p300 As Long
    Dim q300 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long
    Dim p312 As Long
    Dim q312 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 1, ctx.UserData)
        p300 = .BufPos
        q300 = .ThunkPos
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            lCaptureBegin = .BufPos
            Do
                If .BufData(.BufPos) = 39 Then      ' "'"
                    Exit Do
                End If
                If Not ParseChar() Then
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            If .BufData(.BufPos) = 39 Then          ' "'"
                .BufPos = .BufPos + 1
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 2, ctx.UserData)
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString: .LastBufPos = .BufPos
                End If
                ParseLiteralNoSpace = True
                Exit Function
            Else
                .BufPos = p300
                .ThunkPos = q300
            End If
        End If
        p312 = .BufPos
        q312 = .ThunkPos
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            lCaptureBegin = .BufPos
            Do
                If .BufData(.BufPos) = 34 Then      ' """
                    Exit Do
                End If
                If Not ParseChar() Then
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            If .BufData(.BufPos) = 34 Then          ' """
                .BufPos = .BufPos + 1
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 2, ctx.UserData)
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString: .LastBufPos = .BufPos
                End If
                ParseLiteralNoSpace = True
                Exit Function
            Else
                .BufPos = p312
                .ThunkPos = q312
            End If
        End If
        If ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseLiteralNoSpace = True
        End If
    End With
End Function

Public Function ParseIgnoreCase() As Boolean
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "IgnoreCase", 1, ctx.UserData)
        lCaptureBegin = .BufPos
        Select Case .BufData(.BufPos)
        Case 115, 105                               ' [si]
            .BufPos = .BufPos + 1
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "IgnoreCase", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseIgnoreCase = True
            Exit Function
        End Select
        If ConsoleTrace(ctx.BufPos + 1, "IgnoreCase", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseIgnoreCase = True
        End If
    End With
End Function

Public Function ParseClassNoSpace() As Boolean
    Dim p374 As Long
    Dim q374 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        p374 = .BufPos
        q374 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "ClassNoSpace", 1, ctx.UserData)
        If .BufData(.BufPos) = 91 Then              ' "["
            .BufPos = .BufPos + 1
            lCaptureBegin = .BufPos
            Do
                If .BufData(.BufPos) = 93 Then      ' "]"
                    Exit Do
                End If
                If Not ParseRange() Then
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            If .BufData(.BufPos) = 93 Then          ' "]"
                .BufPos = .BufPos + 1
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "ClassNoSpace", 2, ctx.UserData)
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString: .LastBufPos = .BufPos
                End If
                ParseClassNoSpace = True
                Exit Function
            Else
                .BufPos = p374
                .ThunkPos = q374
            End If
        End If
        If ConsoleTrace(ctx.BufPos + 1, "ClassNoSpace", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseClassNoSpace = True
        End If
    End With
End Function

Private Function ParseDOT() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "DOT", 1, ctx.UserData)
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "DOT", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseDOT = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "DOT", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseDOT = True
        End If
    End With
End Function

Private Function ParseBEGIN() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "BEGIN", 1, ctx.UserData)
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "BEGIN", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseBEGIN = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "BEGIN", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseBEGIN = True
        End If
    End With
End Function

Private Function ParseEND() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "END", 1, ctx.UserData)
        If .BufData(.BufPos) = 62 Then              ' ">"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "END", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseEND = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "END", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseEND = True
        End If
    End With
End Function

Public Function ParseSettingName() As Boolean
    Dim p242 As Long
    Dim q242 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 1, ctx.UserData)
        If ParseAT() Then
            p242 = .BufPos
            q242 = .ThunkPos
            If ParseIdentifier() Then
                Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p242
                .ThunkPos = q242
            End If
            If ParseLiteral() Then
                Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p242
                .ThunkPos = q242
            End If
            Call ParseEMPTY
            Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
            ParseSettingName = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SettingName", 3, ctx.UserData) And False Then
            ParseSettingName = True
        End If
    End With
End Function

Public Function ParseSettingValue() As Boolean
    Dim p253 As Long
    Dim q253 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 1, ctx.UserData)
        p253 = .BufPos
        q253 = .ThunkPos
        If ParseIdentifier() Then
            Call ParseSEMI
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p253
            .ThunkPos = q253
        End If
        If ParseLiteral() Then
            Call ParseSEMI
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p253
            .ThunkPos = q253
        End If
        If ParseCodeBlock() Then
            Call ParseSEMI
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p253
            .ThunkPos = q253
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SettingValue", 3, ctx.UserData) And False Then
            ParseSettingValue = True
        End If
    End With
End Function

Private Function ParseAT() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "AT", 1, ctx.UserData)
        If .BufData(.BufPos) = 64 Then              ' "@"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "AT", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseAT = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "AT", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseAT = True
        End If
    End With
End Function

Private Sub ParseEMPTY()
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "EMPTY", 1, ctx.UserData)
        lCaptureBegin = .BufPos
        lCaptureEnd = .BufPos
        Call ParseSpacing
        .CaptureBegin = lCaptureBegin
        .CaptureEnd = lCaptureEnd
        Call ConsoleTrace(ctx.BufPos + 1, "EMPTY", 2, ctx.UserData)
        If .BufPos > .LastBufPos Then
            .LastError = vbNullString: .LastBufPos = .BufPos
        End If
    End With
End Sub

Public Function ParseIdentStart() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "IdentStart", 1, ctx.UserData)
        Select Case .BufData(.BufPos)
        Case 97 To 122, 65 To 90, 95, 92, 45        ' [a-zA-Z_\-]
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "IdentStart", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseIdentStart = True
            Exit Function
        End Select
        If ConsoleTrace(ctx.BufPos + 1, "IdentStart", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseIdentStart = True
        End If
    End With
End Function

Public Function ParseIdentCont() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "IdentCont", 1, ctx.UserData)
        Select Case .BufData(.BufPos)
        Case 97 To 122, 65 To 90, 95, 48 To 57, 92, 45 ' [a-zA-Z_0-9\-]
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "IdentCont", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseIdentCont = True
            Exit Function
        End Select
        If ConsoleTrace(ctx.BufPos + 1, "IdentCont", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseIdentCont = True
        End If
    End With
End Function

Public Function ParseChar() As Boolean
    Dim p393 As Long
    Dim p396 As Long
    Dim p402 As Long
    Dim p407 As Long
    Dim p411 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "Char", 1, ctx.UserData)
        p393 = .BufPos
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 97, 98, 101, 102, 110, 114, 116, 118, 39, 34, 91, 93, 92 ' [abefnrtv'"\[\]\\]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString: .LastBufPos = .BufPos
                End If
                ParseChar = True
                Exit Function
            Case Else
                .BufPos = p393
            End Select
        End If
        p396 = .BufPos
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 48 To 51                           ' [0-3]
                .BufPos = .BufPos + 1
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                    Select Case .BufData(.BufPos)
                    Case 48 To 55                   ' [0-7]
                        .BufPos = .BufPos + 1
                        Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                        If .BufPos > .LastBufPos Then
                            .LastError = vbNullString: .LastBufPos = .BufPos
                        End If
                        ParseChar = True
                        Exit Function
                    Case Else
                        .BufPos = p396
                    End Select
                Case Else
                    .BufPos = p396
                End Select
            Case Else
                .BufPos = p396
            End Select
        End If
        p402 = .BufPos
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 48 To 55                           ' [0-7]
                .BufPos = .BufPos + 1
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                End Select
                Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString: .LastBufPos = .BufPos
                End If
                ParseChar = True
                Exit Function
            Case Else
                .BufPos = p402
            End Select
        End If
        p407 = .BufPos
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 120 Then ' "\\x"
            .BufPos = .BufPos + 2
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
                Select Case .BufData(.BufPos)
                Case 48 To 57, 97 To 102, 65 To 70  ' [0-9a-fA-F]
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                    If .BufPos > .LastBufPos Then
                        .LastError = vbNullString: .LastBufPos = .BufPos
                    End If
                    ParseChar = True
                    Exit Function
                Case Else
                    .BufPos = p407
                End Select
            Case Else
                .BufPos = p407
            End Select
        End If
        p411 = .BufPos
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
                            Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                            If .BufPos > .LastBufPos Then
                                .LastError = vbNullString: .LastBufPos = .BufPos
                            End If
                            ParseChar = True
                            Exit Function
                        Case Else
                            .BufPos = p411
                        End Select
                    Case Else
                        .BufPos = p411
                    End Select
                Case Else
                    .BufPos = p411
                End Select
            Case Else
                .BufPos = p411
            End Select
        End If
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 45 Then ' "\\-"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseChar = True
            Exit Function
        End If
        If Not .BufData(.BufPos) = 92 Then          ' "\\"
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString: .LastBufPos = .BufPos
                End If
                ParseChar = True
                Exit Function
            End If
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Char", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseChar = True
            Exit Function
        End If
        If LenB(.LastError) = 0 Or .BufPos > .LastBufPos Then
            .LastError = "Expected char" : .LastBufPos = .BufPos
        End If
    End With
End Function

Public Function ParseRange() As Boolean
    Dim p380 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "Range", 1, ctx.UserData)
        p380 = .BufPos
        If ParseChar() Then
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = p380
                Else
                    If ParseChar() Then
                        Call ConsoleTrace(ctx.BufPos + 1, "Range", 2, ctx.UserData)
                        If .BufPos > .LastBufPos Then
                            .LastError = vbNullString: .LastBufPos = .BufPos
                        End If
                        ParseRange = True
                        Exit Function
                    Else
                        .BufPos = p380
                    End If
                End If
            Else
                .BufPos = p380
            End If
        End If
        If ParseChar() Then
            Call ConsoleTrace(ctx.BufPos + 1, "Range", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseRange = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Range", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseRange = True
        End If
    End With
End Function

Private Function ParseLEFTARROW() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "LEFTARROW", 1, ctx.UserData)
        If .BufData(.BufPos) = 60 And .BufData(.BufPos + 1) = 45 Then ' "<-"
            .BufPos = .BufPos + 2
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "LEFTARROW", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseLEFTARROW = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "LEFTARROW", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseLEFTARROW = True
        End If
    End With
End Function

Private Function ParseSLASH() As Boolean
    Dim p443 As Long

    With ctx
        p443 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SLASH", 1, ctx.UserData)
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = p443
            Else
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "SLASH", 2, ctx.UserData)
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString: .LastBufPos = .BufPos
                End If
                ParseSLASH = True
                Exit Function
            End If
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SLASH", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseSLASH = True
        End If
    End With
End Function

Public Function ParseSpace() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "Space", 1, ctx.UserData)
        If .BufData(.BufPos) = 32 Then              ' " "
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "Space", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseSpace = True
            Exit Function
        End If
        If .BufData(.BufPos) = 9 Then               ' "\t"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "Space", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseSpace = True
            Exit Function
        End If
        If ParseEndOfLine() Then
            Call ConsoleTrace(ctx.BufPos + 1, "Space", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseSpace = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Space", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseSpace = True
        End If
    End With
End Function

Public Function ParseComment() As Boolean
    Dim p545 As Long
    Dim p540 As Long

    With ctx
        p545 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Comment", 1, ctx.UserData)
        If .BufData(.BufPos) = 35 Then              ' "#"
            .BufPos = .BufPos + 1
            Do
                p540 = .BufPos
                If ParseEndOfLine() Then
                    .BufPos = p540
                    Exit Do
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p540
                    Exit Do
                End If
            Loop
            If ParseEndOfLine() Then
                Call ConsoleTrace(ctx.BufPos + 1, "Comment", 2, ctx.UserData)
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString: .LastBufPos = .BufPos
                End If
                ParseComment = True
                Exit Function
            Else
                .BufPos = p545
            End If
        End If
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 47 Then ' "//"
            .BufPos = .BufPos + 2
            Do
                p540 = .BufPos
                If ParseEndOfLine() Then
                    .BufPos = p540
                    Exit Do
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p540
                    Exit Do
                End If
            Loop
            If ParseEndOfLine() Then
                Call ConsoleTrace(ctx.BufPos + 1, "Comment", 2, ctx.UserData)
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString: .LastBufPos = .BufPos
                End If
                ParseComment = True
                Exit Function
            Else
                .BufPos = p545
            End If
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Comment", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseComment = True
        End If
    End With
End Function

Public Function ParseEndOfLine() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 1, ctx.UserData)
        If .BufData(.BufPos) = 13 And .BufData(.BufPos + 1) = 10 Then ' "\r\n"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseEndOfLine = True
            Exit Function
        End If
        If .BufData(.BufPos) = 10 Then              ' "\n"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseEndOfLine = True
            Exit Function
        End If
        If .BufData(.BufPos) = 13 Then              ' "\r"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseEndOfLine = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseEndOfLine = True
        End If
    End With
End Function

Public Function ParseCodeNC() As Boolean
    Dim p602 As Long

    With ctx
        p602 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CodeNC", 1, ctx.UserData)
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Do
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If ParseCodeNC() Then
                        GoTo L24
                    Else
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        GoTo L24
                    Else
                        If ParseCodeNC() Then
                            GoTo L24
                        Else
                            Exit Do
                        End If
                    End If
                End Select
L24:
            Loop
            If .BufData(.BufPos) = 125 Then         ' "}"
                .BufPos = .BufPos + 1
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "CodeNC", 2, ctx.UserData)
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString: .LastBufPos = .BufPos
                End If
                ParseCodeNC = True
                Exit Function
            Else
                .BufPos = p602
            End If
        End If
        If ConsoleTrace(ctx.BufPos + 1, "CodeNC", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseCodeNC = True
        End If
    End With
End Function

Private Function ParseEQUAL() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "EQUAL", 1, ctx.UserData)
        If .BufData(.BufPos) = 61 Then              ' "="
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "EQUAL", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseEQUAL = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "EQUAL", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseEQUAL = True
        End If
    End With
End Function

Private Function ParseBAR() As Boolean
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "BAR", 1, ctx.UserData)
        If .BufData(.BufPos) = 124 Then             ' "|"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "BAR", 2, ctx.UserData)
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseBAR = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "BAR", 3, ctx.UserData) And False Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            ParseBAR = True
        End If
    End With
End Function

Private Sub pvImplAction(ByVal eAction As UcsParserActionsEnum, ByVal lOffset As Long, ByVal lSize As Long)
    Dim lNode           As Long
    
    With ctx
    Select Case eAction
    Case ucsAct_2_Setting
         With .UserData: .SetSettingValue .Pop, Mid$(ctx.Contents, lOffset, lSize): End With
    Case ucsAct_1_Setting
         With .UserData: .Push .MakeString(lOffset, Mid$(ctx.Contents, lOffset, lSize)): End With
    Case ucsAct_3_Rule
         With .UserData: lNode = .Pop: .SetExpression .Pop, lNode: End With
    Case ucsAct_2_Rule
         With .UserData: .RuleDescription(.Top) = Mid$(ctx.Contents, lOffset, lSize): End With
    Case ucsAct_1_Rule
           With .UserData
                                                                                lNode = .BeginRule(.FindRule(lOffset, Mid$(ctx.Contents, lOffset, lSize)))
                                                                                If .RuleExpression(lNode) <> 0 Then
                                                                                    AppendError "Rule '" & Mid$(ctx.Contents, lOffset, lSize) & "' redefined", lOffset
                                                                                End If
                                                                                .Push lNode
                                                                            End With
    Case ucsAct_1_ChoiceExpr
         With .UserData: lNode = .Pop: .Push .AppendChild(.MakeChoice(lOffset, .Pop), lNode): End With
    Case ucsAct_1_SequenceExpr
         With .UserData: lNode = .Pop: .Push .AppendSequence(.MakeSequence(lOffset, .Pop), lNode): End With
    Case ucsAct_2_SequenceExpr
         With .UserData: .Push .MakePredicate(lOffset, "", AlwaysTrue:=True): End With
    Case ucsAct_1_ErrorExpr
         With .UserData: .Push .AppendChild(.MakeChoice(lOffset, .Pop), .MakeErrorReference(lOffset, Mid$(ctx.Contents, lOffset, lSize))): End With
    Case ucsAct_2_ErrorExpr
         With .UserData: .Push .AppendChild(.MakeChoice(lOffset, .Pop), .MakeErrorString(lOffset, Mid$(ctx.Contents, lOffset, lSize))): End With
    Case ucsAct_2_VariableExpr
         With .UserData: .Push .SetVariable(lOffset, .Pop, .Pop): End With
    Case ucsAct_1_VariableExpr
         With .UserData: .Push .MakeVariable(lOffset, Mid$(ctx.Contents, lOffset, lSize)): End With
    Case ucsAct_1_PrefixExpr
         With .UserData: .Push .MakePredicate(lOffset, Mid$(ctx.Contents, lOffset, lSize)): End With
    Case ucsAct_2_PrefixExpr
         With .UserData: .Push .MakePeekFor(lOffset, .Pop): End With
    Case ucsAct_3_PrefixExpr
         With .UserData: .Push .MakePeekNot(lOffset, .Pop): End With
    Case ucsAct_1_SuffixExpr
         With .UserData: .Push .MakeQuery(lOffset, .Pop): End With
    Case ucsAct_2_SuffixExpr
         With .UserData: .Push .MakeStar(lOffset, .Pop): End With
    Case ucsAct_3_SuffixExpr
         With .UserData: .Push .MakePlus(lOffset, .Pop): End With
    Case ucsAct_1_PrimaryExpr
         With .UserData: .Push .MakeReference(lOffset, Mid$(ctx.Contents, lOffset, lSize)): End With
    Case ucsAct_3_PrimaryExpr
         With .UserData: .SetIgnoreCase .Top, (Mid$(ctx.Contents, lOffset, lSize) = "i"): End With
    Case ucsAct_2_PrimaryExpr
         With .UserData: .Push .MakeString(lOffset, Mid$(ctx.Contents, lOffset, lSize)): End With
    Case ucsAct_5_PrimaryExpr
         With .UserData: .SetIgnoreCase .Top, (Mid$(ctx.Contents, lOffset, lSize) = "i"): End With
    Case ucsAct_4_PrimaryExpr
         With .UserData: .Push .MakeClass(lOffset, Mid$(ctx.Contents, lOffset, lSize)): End With
    Case ucsAct_6_PrimaryExpr
         With .UserData: .Push .MakeDot(lOffset): End With
    Case ucsAct_7_PrimaryExpr
         With .UserData: .Push .MakeAction(lOffset, Mid$(ctx.Contents, lOffset, lSize)): End With
    Case ucsAct_8_PrimaryExpr
         With .UserData: .Push .MakePredicate(lOffset, "YY_BEGIN", AlwaysTrue:=True): End With
    Case ucsAct_9_PrimaryExpr
         With .UserData: .Push .MakePredicate(lOffset, "YY_END", AlwaysTrue:=True): End With
    End Select
    End With
End Sub

'= user-defined functions ================================================

Private Function ClearErrors(ByVal lOffset As Long) As Boolean
    Dim lIdx            As Long
    
    If Not ctx.ParseErrors Is Nothing Then
        For lIdx = ctx.ParseErrors.Count To 1 Step -1
            If ctx.ParseErrors.Item(lIdx)(1) <= lOffset Then
                ctx.ParseErrors.Remove lIdx
            End If
        Next
        If ctx.ParseErrors.Count = 0 Then
            Set ctx.ParseErrors = Nothing
        End If
    End If
    ClearErrors = True
End Function

Private Function AppendError(sText As String, ByVal lOffset As Long) As Boolean
    If ctx.ParseErrors Is Nothing Then
        Set ctx.ParseErrors = New Collection
    End If
    ctx.ParseErrors.Add Array(sText, lOffset)
End Function

Private Function SetUserDataTag(sText As String) As Boolean
    ctx.UserData.Tag = sText
    SetUserDataTag = True
End Function

Public Function GetParseErrors() As Collection
    Set GetParseErrors = ctx.ParseErrors
End Function

Public Function GetContents() As String
    GetContents = ctx.Contents
End Function
