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
' Auto-generated on 23.7.2018 14:24:42
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
    LastExpected        As String
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
    ParserVersion = "23.7.2018 14:24:42"
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
    Dim p19 As Long
    Dim q19 As Long
    Dim i10 As Long
    Dim p9 As Long
    Dim q9 As Long
    Dim p15 As Long
    Dim q15 As Long

    With ctx
        p19 = .BufPos
        q19 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Grammar", 1, ctx.UserData)
        Call ParseSpacing
        For i10 = 0 To LNG_MAXINT
            p9 = .BufPos
            q9 = .ThunkPos
            If Not ParseSetting() Then
                .BufPos = p9
                .ThunkPos = q9
                If Not ParseRule() Then
                    .BufPos = p9
                    .ThunkPos = q9
                    .BufPos = p9
                    .ThunkPos = q9
                    Exit For
                End If
            End If
        Next
        If i10 <> 0 Then
            p15 = .BufPos
            q15 = .ThunkPos
            If ParseEndOfFile() Then
                Call ConsoleTrace(ctx.BufPos + 1, "Grammar", 2, ctx.UserData)
                ParseGrammar = True
                Exit Function
            Else
                .BufPos = p15
                .ThunkPos = q15
            End If
            If AppendError("Expected rule or setting", ctx.BufPos + 1) Then
                Call ConsoleTrace(ctx.BufPos + 1, "Grammar", 2, ctx.UserData)
                ParseGrammar = True
                Exit Function
            End If
            .BufPos = p19
            .ThunkPos = q19
        Else
            .BufPos = p19
            .ThunkPos = q19
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Grammar", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Sub ParseSpacing()
    Dim p521 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "Spacing", 1, ctx.UserData)
        Do
            p521 = .BufPos
            If Not ParseSpace() Then
                .BufPos = p521
                If Not ParseLineComment() Then
                    .BufPos = p521
                    If Not ParseBlockComment() Then
                        .BufPos = p521
                        .BufPos = p521
                        Exit Do
                    End If
                End If
            End If
        Loop
        Call ConsoleTrace(ctx.BufPos + 1, "Spacing", 2, ctx.UserData)
    End With
End Sub

Public Function ParseSetting() As Boolean
    Dim p237 As Long
    Dim q237 As Long
    Dim p233 As Long
    Dim q233 As Long

    With ctx
        p237 = .BufPos
        q237 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 1, ctx.UserData)
        If ParseSettingName() Then
            If CheckSettingName(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                pvPushThunk ucsAct_1_Setting, .CaptureBegin, .CaptureEnd
                If ClearErrors(ctx.BufPos + 1) And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                    p233 = .BufPos
                    q233 = .ThunkPos
                    If ParseSettingValue() Then
                        pvPushThunk ucsAct_2_Setting, .CaptureBegin, .CaptureEnd
                        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 2, ctx.UserData)
                        ParseSetting = True
                        Exit Function
                    Else
                        .BufPos = p233
                        .ThunkPos = q233
                    End If
                    If AppendError("Invalid value for setting '" & .UserData.Tag & "'", ctx.BufPos + 1) Then
                        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 2, ctx.UserData)
                        ParseSetting = True
                        Exit Function
                    End If
                    .BufPos = p237
                    .ThunkPos = q237
                Else
                    .BufPos = p237
                    .ThunkPos = q237
                End If
            Else
                .BufPos = p237
                .ThunkPos = q237
            End If
        Else
            .BufPos = p237
            .ThunkPos = q237
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Setting", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseRule() As Boolean
    Dim p45 As Long
    Dim q45 As Long
    Dim p29 As Long
    Dim q29 As Long
    Dim p40 As Long
    Dim q40 As Long

    With ctx
        p45 = .BufPos
        q45 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Rule", 1, ctx.UserData)
        If ParseIdentifier() Then
            pvPushThunk ucsAct_1_Rule, .CaptureBegin, .CaptureEnd
            If ClearErrors(ctx.BufPos + 1) And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                p29 = .BufPos
                q29 = .ThunkPos
                If ParseLiteral() Then
                    pvPushThunk ucsAct_2_Rule, .CaptureBegin, .CaptureEnd
                Else
                    .BufPos = p29
                    .ThunkPos = q29
                End If
                If ParseAssignLiteral() Then
                    Call ParseExpression
                    pvPushThunk ucsAct_3_Rule, .CaptureBegin, .CaptureEnd
                    p40 = .BufPos
                    q40 = .ThunkPos
                    If Not ParseSEMI() Then
                        .BufPos = p40
                        .ThunkPos = q40
                    End If
                    Call ConsoleTrace(ctx.BufPos + 1, "Rule", 2, ctx.UserData)
                    ParseRule = True
                    Exit Function
                Else
                    .BufPos = p45
                    .ThunkPos = q45
                End If
            Else
                .BufPos = p45
                .ThunkPos = q45
            End If
        Else
            .BufPos = p45
            .ThunkPos = q45
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Rule", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseEndOfFile() As Boolean
    Dim p584 As Long
    Dim p580 As Long

    With ctx
        p584 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "EndOfFile", 1, ctx.UserData)
        p580 = .BufPos
        If .BufPos < .BufSize Then
            .BufPos = p584
        Else
            .BufPos = p580
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
    Dim p276 As Long
    Dim lCaptureBegin As Long
    Dim p269 As Long
    Dim lCaptureEnd As Long

    With ctx
        p276 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Identifier", 1, ctx.UserData)
        lCaptureBegin = .BufPos
        If ParseIdentStart() Then
            Do
                p269 = .BufPos
                If Not ParseIdentCont() Then
                    .BufPos = p269
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
            .BufPos = p276
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Identifier", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseLiteral() As Boolean
    Dim p296 As Long

    With ctx
        p296 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Literal", 1, ctx.UserData)
        If ParseLiteralNoSpace() Then
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "Literal", 2, ctx.UserData)
            ParseLiteral = True
            Exit Function
        Else
            .BufPos = p296
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Literal", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseAssignLiteral() As Boolean
    Dim p640 As Long
    Dim p636 As Long

    With ctx
        p640 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 1, ctx.UserData)
        p636 = .BufPos
        If ParseLEFTARROW() Then
            Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 2, ctx.UserData)
            ParseAssignLiteral = True
            Exit Function
        Else
            .BufPos = p636
        End If
        If ParseEQUAL() Then
            Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 2, ctx.UserData)
            ParseAssignLiteral = True
            Exit Function
        Else
            .BufPos = p636
        End If
        .BufPos = p640
        If ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Sub ParseExpression()
    Call ConsoleTrace(ctx.BufPos + 1, "Expression", 1, ctx.UserData)
    Call ParseChoiceExpr
    Call ConsoleTrace(ctx.BufPos + 1, "Expression", 2, ctx.UserData)
L6:
End Sub

Private Function ParseSEMI() As Boolean
    Dim p681 As Long

    With ctx
        p681 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SEMI", 1, ctx.UserData)
        If .BufData(.BufPos) = 59 Then              ' ";"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "SEMI", 2, ctx.UserData)
            Call pvSetAdvance
            ParseSEMI = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SEMI", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Sub ParseChoiceExpr()
    Dim p59 As Long
    Dim q59 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceExpr", 1, ctx.UserData)
        Call ParseSequenceExpr
        Do
            p59 = .BufPos
            q59 = .ThunkPos
            If Not ParseChoiceLiteral() Then
                .BufPos = p59
                .ThunkPos = q59
                Exit Do
            End If
            Call ParseSequenceExpr
            pvPushThunk ucsAct_1_ChoiceExpr, .CaptureBegin, .CaptureEnd
        Loop
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceExpr", 2, ctx.UserData)
L7:
    End With
End Sub

Public Sub ParseSequenceExpr()
    Dim p76 As Long
    Dim q76 As Long
    Dim p72 As Long
    Dim q72 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "SequenceExpr", 1, ctx.UserData)
        p76 = .BufPos
        q76 = .ThunkPos
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
        Else
            .BufPos = p76
            .ThunkPos = q76
            pvPushThunk ucsAct_2_SequenceExpr, .CaptureBegin, .CaptureEnd
        End If
        Call ConsoleTrace(ctx.BufPos + 1, "SequenceExpr", 2, ctx.UserData)
    End With
End Sub

Public Function ParseChoiceLiteral() As Boolean
    Dim p649 As Long
    Dim p645 As Long

    With ctx
        p649 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 1, ctx.UserData)
        p645 = .BufPos
        If ParseSLASH() Then
            Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 2, ctx.UserData)
            ParseChoiceLiteral = True
            Exit Function
        Else
            .BufPos = p645
        End If
        If ParseBAR() Then
            Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 2, ctx.UserData)
            ParseChoiceLiteral = True
            Exit Function
        Else
            .BufPos = p645
        End If
        .BufPos = p649
        If ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseErrorExpr() As Boolean
    Dim p101 As Long
    Dim q101 As Long
    Dim p87 As Long
    Dim q87 As Long
    Dim p94 As Long
    Dim q94 As Long

    With ctx
        p101 = .BufPos
        q101 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 1, ctx.UserData)
        If ParseVariableExpr() Then
            p87 = .BufPos
            q87 = .ThunkPos
            If ParseCARET() Then
                If ClearErrors(ctx.BufPos + 1) Then
                    p94 = .BufPos
                    q94 = .ThunkPos
                    If ParseIdentifier() Then
                        pvPushThunk ucsAct_1_ErrorExpr, .CaptureBegin, .CaptureEnd
                        GoTo L10
                    Else
                        .BufPos = p94
                        .ThunkPos = q94
                    End If
                    If ParseLiteral() Then
                        pvPushThunk ucsAct_2_ErrorExpr, .CaptureBegin, .CaptureEnd
                        GoTo L10
                    Else
                        .BufPos = p94
                        .ThunkPos = q94
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
            Else
                .BufPos = p87
                .ThunkPos = q87
            End If
L10:
            Call ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 2, ctx.UserData)
            ParseErrorExpr = True
            Exit Function
        Else
            .BufPos = p101
            .ThunkPos = q101
        End If
        If ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseVariableExpr() As Boolean
    Dim p116 As Long
    Dim q116 As Long
    Dim p112 As Long
    Dim q112 As Long

    With ctx
        p116 = .BufPos
        q116 = .ThunkPos
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
        .BufPos = p116
        .ThunkPos = q116
        If ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCARET() As Boolean
    Dim p697 As Long

    With ctx
        p697 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CARET", 1, ctx.UserData)
        If .BufData(.BufPos) = 94 Then              ' "^"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "CARET", 2, ctx.UserData)
            Call pvSetAdvance
            ParseCARET = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "CARET", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCOLON() As Boolean
    Dim p673 As Long

    With ctx
        p673 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "COLON", 1, ctx.UserData)
        If .BufData(.BufPos) = 58 Then              ' ":"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "COLON", 2, ctx.UserData)
            Call pvSetAdvance
            ParseCOLON = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "COLON", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParsePrefixExpr() As Boolean
    Dim p139 As Long
    Dim q139 As Long
    Dim p129 As Long
    Dim q129 As Long

    With ctx
        p139 = .BufPos
        q139 = .ThunkPos
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
        Else
            .BufPos = p129
            .ThunkPos = q129
        End If
        If ParseAND() Then
            If ParseSuffixExpr() Then
                pvPushThunk ucsAct_2_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p129
                .ThunkPos = q129
            End If
        Else
            .BufPos = p129
            .ThunkPos = q129
        End If
        If ParseNOT() Then
            If ParseSuffixExpr() Then
                pvPushThunk ucsAct_3_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p129
                .ThunkPos = q129
            End If
        Else
            .BufPos = p129
            .ThunkPos = q129
        End If
        If ParseSuffixExpr() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
            ParsePrefixExpr = True
            Exit Function
        Else
            .BufPos = p129
            .ThunkPos = q129
        End If
        .BufPos = p139
        .ThunkPos = q139
        If ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseAND() As Boolean
    Dim p449 As Long

    With ctx
        p449 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "AND", 1, ctx.UserData)
        If .BufData(.BufPos) = 38 Then              ' "&"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "AND", 2, ctx.UserData)
            Call pvSetAdvance
            ParseAND = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "AND", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseCodeBlock() As Boolean
    Dim p602 As Long
    Dim p587 As Long
    Dim lCaptureBegin As Long
    Dim p594 As Long
    Dim lCaptureEnd As Long

    With ctx
        p602 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CodeBlock", 1, ctx.UserData)
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Do
                p587 = .BufPos
                If Not ParseEndOfLine() Then
                    .BufPos = p587
                    Exit Do
                End If
            Loop
            lCaptureBegin = .BufPos
            Do
                p594 = .BufPos
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If Not ParseCodeNC() Then
                        .BufPos = p594
                        .BufPos = p594
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                    Else
                        If Not ParseCodeNC() Then
                            .BufPos = p594
                            .BufPos = p594
                            Exit Do
                        End If
                    End If
                End Select
            Loop
            lCaptureEnd = .BufPos
            If .BufData(.BufPos) = 125 Then         ' "}"
                .BufPos = .BufPos + 1
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "CodeBlock", 2, ctx.UserData)
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call pvSetAdvance
                ParseCodeBlock = True
                Exit Function
            End If
        End If
        If ConsoleTrace(ctx.BufPos + 1, "CodeBlock", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseSuffixExpr() As Boolean
    Dim p161 As Long
    Dim q161 As Long
    Dim p151 As Long
    Dim q151 As Long

    With ctx
        p161 = .BufPos
        q161 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 1, ctx.UserData)
        If ParsePrimaryExpr() Then
            p151 = .BufPos
            q151 = .ThunkPos
            If ParseQUESTION() Then
                pvPushThunk ucsAct_1_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L12
            Else
                .BufPos = p151
                .ThunkPos = q151
            End If
            If ParseSTAR() Then
                pvPushThunk ucsAct_2_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L12
            Else
                .BufPos = p151
                .ThunkPos = q151
            End If
            If ParsePLUS() Then
                pvPushThunk ucsAct_3_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L12
            Else
                .BufPos = p151
                .ThunkPos = q151
            End If
            .BufPos = p151
            .ThunkPos = q151
L12:
            Call ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 2, ctx.UserData)
            ParseSuffixExpr = True
            Exit Function
        Else
            .BufPos = p161
            .ThunkPos = q161
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseNOT() As Boolean
    Dim p457 As Long

    With ctx
        p457 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "NOT", 1, ctx.UserData)
        If .BufData(.BufPos) = 33 Then              ' "!"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "NOT", 2, ctx.UserData)
            Call pvSetAdvance
            ParseNOT = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "NOT", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParsePrimaryExpr() As Boolean
    Dim p220 As Long
    Dim q220 As Long
    Dim p182 As Long
    Dim q182 As Long
    Dim p169 As Long
    Dim q169 As Long
    Dim p165 As Long
    Dim p181 As Long
    Dim q181 As Long
    Dim p190 As Long
    Dim q190 As Long
    Dim p199 As Long
    Dim q199 As Long

    With ctx
        p220 = .BufPos
        q220 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 1, ctx.UserData)
        p182 = .BufPos
        q182 = .ThunkPos
        If ParseIdentifier() Then
            p169 = .BufPos
            q169 = .ThunkPos
            p165 = .BufPos
            If Not ParseLiteralNC() Then
                .BufPos = p165
            End If
            If ParseAssignLiteral() Then
                .BufPos = p182
                .ThunkPos = q182
            Else
                .BufPos = p169
                .ThunkPos = q169
                pvPushThunk ucsAct_1_PrimaryExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                ParsePrimaryExpr = True
                Exit Function
            End If
        Else
            .BufPos = p182
            .ThunkPos = q182
        End If
        If ParseOPEN() Then
            If ClearErrors(ctx.BufPos + 1) Then
                p181 = .BufPos
                q181 = .ThunkPos
                Call ParseExpression
                If ParseCLOSE() Then
                    Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                    ParsePrimaryExpr = True
                    Exit Function
                Else
                    .BufPos = p181
                    .ThunkPos = q181
                End If
                If AppendError("Invalid expression", ctx.BufPos + 1) Then
                    Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                    ParsePrimaryExpr = True
                    Exit Function
                End If
                .BufPos = p182
                .ThunkPos = q182
            Else
                .BufPos = p182
                .ThunkPos = q182
            End If
        Else
            .BufPos = p182
            .ThunkPos = q182
        End If
        If ParseLiteralNoSpace() Then
            pvPushThunk ucsAct_2_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p190 = .BufPos
            q190 = .ThunkPos
            If ParseIgnoreCase() Then
                pvPushThunk ucsAct_3_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Else
                .BufPos = p190
                .ThunkPos = q190
            End If
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p182
            .ThunkPos = q182
        End If
        If ParseClassNoSpace() Then
            pvPushThunk ucsAct_4_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p199 = .BufPos
            q199 = .ThunkPos
            If ParseIgnoreCase() Then
                pvPushThunk ucsAct_5_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Else
                .BufPos = p199
                .ThunkPos = q199
            End If
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p182
            .ThunkPos = q182
        End If
        If ParseDOT() Then
            pvPushThunk ucsAct_6_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p182
            .ThunkPos = q182
        End If
        If ParseCodeBlock() Then
            pvPushThunk ucsAct_7_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p182
            .ThunkPos = q182
        End If
        If ParseBEGIN() Then
            pvPushThunk ucsAct_8_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p182
            .ThunkPos = q182
        End If
        If ParseEND() Then
            pvPushThunk ucsAct_9_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p182
            .ThunkPos = q182
        End If
        .BufPos = p220
        .ThunkPos = q220
        If ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseQUESTION() As Boolean
    Dim p465 As Long

    With ctx
        p465 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "QUESTION", 1, ctx.UserData)
        If .BufData(.BufPos) = 63 Then              ' "?"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "QUESTION", 2, ctx.UserData)
            Call pvSetAdvance
            ParseQUESTION = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "QUESTION", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSTAR() As Boolean
    Dim p473 As Long

    With ctx
        p473 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "STAR", 1, ctx.UserData)
        If .BufData(.BufPos) = 42 Then              ' "*"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "STAR", 2, ctx.UserData)
            Call pvSetAdvance
            ParseSTAR = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "STAR", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParsePLUS() As Boolean
    Dim p481 As Long

    With ctx
        p481 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "PLUS", 1, ctx.UserData)
        If .BufData(.BufPos) = 43 Then              ' "+"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "PLUS", 2, ctx.UserData)
            Call pvSetAdvance
            ParsePLUS = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "PLUS", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseLiteralNC() As Boolean
    Dim p347 As Long
    Dim p343 As Long
    Dim p329 As Long
    Dim p327 As Long
    Dim p338 As Long
    Dim p336 As Long

    With ctx
        p347 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 1, ctx.UserData)
        p343 = .BufPos
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            Do
                p329 = .BufPos
                p327 = .BufPos
                If .BufData(.BufPos) = 39 Then      ' "'"
                    .BufPos = p329
                    Exit Do
                Else
                    .BufPos = p327
                End If
                If Not ParseChar() Then
                    .BufPos = p329
                    Exit Do
                End If
            Loop
            If .BufData(.BufPos) = 39 Then          ' "'"
                .BufPos = .BufPos + 1
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLiteralNC = True
                Exit Function
            End If
        End If
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            Do
                p338 = .BufPos
                p336 = .BufPos
                If .BufData(.BufPos) = 34 Then      ' """
                    .BufPos = p338
                    Exit Do
                Else
                    .BufPos = p336
                End If
                If Not ParseChar() Then
                    .BufPos = p338
                    Exit Do
                End If
            Loop
            If .BufData(.BufPos) = 34 Then          ' """
                .BufPos = .BufPos + 1
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLiteralNC = True
                Exit Function
            End If
        End If
        .BufPos = p347
        If ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseOPEN() As Boolean
    Dim p489 As Long

    With ctx
        p489 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "OPEN", 1, ctx.UserData)
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "OPEN", 2, ctx.UserData)
            Call pvSetAdvance
            ParseOPEN = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "OPEN", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCLOSE() As Boolean
    Dim p497 As Long

    With ctx
        p497 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CLOSE", 1, ctx.UserData)
        If .BufData(.BufPos) = 41 Then              ' ")"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "CLOSE", 2, ctx.UserData)
            Call pvSetAdvance
            ParseCLOSE = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "CLOSE", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseLiteralNoSpace() As Boolean
    Dim p323 As Long
    Dim p319 As Long
    Dim lCaptureBegin As Long
    Dim p305 As Long
    Dim p302 As Long
    Dim lCaptureEnd As Long
    Dim p315 As Long
    Dim p313 As Long

    With ctx
        p323 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 1, ctx.UserData)
        p319 = .BufPos
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            lCaptureBegin = .BufPos
            Do
                p305 = .BufPos
                p302 = .BufPos
                If .BufData(.BufPos) = 39 Then      ' "'"
                    .BufPos = p305
                    Exit Do
                Else
                    .BufPos = p302
                End If
                If Not ParseChar() Then
                    .BufPos = p305
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            If .BufData(.BufPos) = 39 Then          ' "'"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 2, ctx.UserData)
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call pvSetAdvance
                ParseLiteralNoSpace = True
                Exit Function
            End If
        End If
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            lCaptureBegin = .BufPos
            Do
                p315 = .BufPos
                p313 = .BufPos
                If .BufData(.BufPos) = 34 Then      ' """
                    .BufPos = p315
                    Exit Do
                Else
                    .BufPos = p313
                End If
                If Not ParseChar() Then
                    .BufPos = p315
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            If .BufData(.BufPos) = 34 Then          ' """
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 2, ctx.UserData)
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call pvSetAdvance
                ParseLiteralNoSpace = True
                Exit Function
            End If
        End If
        .BufPos = p323
        If ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseIgnoreCase() As Boolean
    Dim p356 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        p356 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "IgnoreCase", 1, ctx.UserData)
        lCaptureBegin = .BufPos
        Select Case .BufData(.BufPos)
        Case 115, 105                               ' [si]
            .BufPos = .BufPos + 1
            lCaptureEnd = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "IgnoreCase", 2, ctx.UserData)
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            Call pvSetAdvance
            ParseIgnoreCase = True
            Exit Function
        End Select
        If ConsoleTrace(ctx.BufPos + 1, "IgnoreCase", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseClassNoSpace() As Boolean
    Dim p372 As Long
    Dim lCaptureBegin As Long
    Dim p365 As Long
    Dim p362 As Long
    Dim lCaptureEnd As Long

    With ctx
        p372 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ClassNoSpace", 1, ctx.UserData)
        If .BufData(.BufPos) = 91 Then              ' "["
            .BufPos = .BufPos + 1
            lCaptureBegin = .BufPos
            Do
                p365 = .BufPos
                p362 = .BufPos
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = p365
                    Exit Do
                Else
                    .BufPos = p362
                End If
                If Not ParseRange() Then
                    .BufPos = p365
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            If .BufData(.BufPos) = 93 Then          ' "]"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "ClassNoSpace", 2, ctx.UserData)
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call pvSetAdvance
                ParseClassNoSpace = True
                Exit Function
            End If
        End If
        If ConsoleTrace(ctx.BufPos + 1, "ClassNoSpace", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseDOT() As Boolean
    Dim p505 As Long

    With ctx
        p505 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "DOT", 1, ctx.UserData)
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "DOT", 2, ctx.UserData)
            Call pvSetAdvance
            ParseDOT = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "DOT", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseBEGIN() As Boolean
    Dim p623 As Long

    With ctx
        p623 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "BEGIN", 1, ctx.UserData)
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "BEGIN", 2, ctx.UserData)
            Call pvSetAdvance
            ParseBEGIN = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "BEGIN", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseEND() As Boolean
    Dim p631 As Long

    With ctx
        p631 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "END", 1, ctx.UserData)
        If .BufData(.BufPos) = 62 Then              ' ">"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "END", 2, ctx.UserData)
            Call pvSetAdvance
            ParseEND = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "END", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseSettingName() As Boolean
    Dim p250 As Long
    Dim p243 As Long

    With ctx
        p250 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 1, ctx.UserData)
        If ParseAT() Then
            p243 = .BufPos
            If ParseIdentifier() Then
                Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p243
            End If
            If ParseLiteral() Then
                Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p243
            End If
            Call ParseEMPTY
            Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
            ParseSettingName = True
            Exit Function
            .BufPos = p250
        Else
            .BufPos = p250
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SettingName", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseSettingValue() As Boolean
    Dim p262 As Long
    Dim p254 As Long
    Dim p256 As Long

    With ctx
        p262 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 1, ctx.UserData)
        p254 = .BufPos
        If ParseIdentifier() Then
            p256 = .BufPos
            If Not ParseSEMI() Then
                .BufPos = p256
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p254
        End If
        If ParseLiteral() Then
            p256 = .BufPos
            If Not ParseSEMI() Then
                .BufPos = p256
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p254
        End If
        If ParseCodeBlock() Then
            p256 = .BufPos
            If Not ParseSEMI() Then
                .BufPos = p256
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p254
        End If
        .BufPos = p262
        If ConsoleTrace(ctx.BufPos + 1, "SettingValue", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseAT() As Boolean
    Dim p689 As Long

    With ctx
        p689 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "AT", 1, ctx.UserData)
        If .BufData(.BufPos) = 64 Then              ' "@"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "AT", 2, ctx.UserData)
            Call pvSetAdvance
            ParseAT = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "AT", 3, ctx.UserData) And False Then
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
        lCaptureEnd = .BufPos
        Call ParseSpacing
        Call ConsoleTrace(ctx.BufPos + 1, "EMPTY", 2, ctx.UserData)
        .CaptureBegin = lCaptureBegin
        .CaptureEnd = lCaptureEnd
    End With
End Sub

Public Function ParseIdentStart() As Boolean
    Dim p282 As Long

    With ctx
        p282 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "IdentStart", 1, ctx.UserData)
        Select Case .BufData(.BufPos)
        Case 97 To 122, 65 To 90, 95, 92, 45        ' [a-zA-Z_\-]
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "IdentStart", 2, ctx.UserData)
            Call pvSetAdvance
            ParseIdentStart = True
            Exit Function
        End Select
        If ConsoleTrace(ctx.BufPos + 1, "IdentStart", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseIdentCont() As Boolean
    Dim p288 As Long

    With ctx
        p288 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "IdentCont", 1, ctx.UserData)
        Select Case .BufData(.BufPos)
        Case 97 To 122, 65 To 90, 95, 48 To 57, 92, 45 ' [a-zA-Z_0-9\-]
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "IdentCont", 2, ctx.UserData)
            Call pvSetAdvance
            ParseIdentCont = True
            Exit Function
        End Select
        If ConsoleTrace(ctx.BufPos + 1, "IdentCont", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseChar() As Boolean
    Dim p421 As Long
    Dim p395 As Long
    Dim p415 As Long

    With ctx
        p421 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Char", 1, ctx.UserData)
        p395 = .BufPos
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 97, 98, 101, 102, 110, 114, 116, 118, 39, 34, 91, 93, 92 ' [abefnrtv'"\[\]\\]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                Call pvSetAdvance
                ParseChar = True
                Exit Function
            End Select
        End If
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
                        Call pvSetAdvance
                        ParseChar = True
                        Exit Function
                    End Select
                End Select
            End Select
        End If
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
                Call pvSetAdvance
                ParseChar = True
                Exit Function
            End Select
        End If
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 120 Then ' "\\x"
            .BufPos = .BufPos + 2
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
                Select Case .BufData(.BufPos)
                Case 48 To 57, 97 To 102, 65 To 70  ' [0-9a-fA-F]
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                    Call pvSetAdvance
                    ParseChar = True
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
                            Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                            Call pvSetAdvance
                            ParseChar = True
                            Exit Function
                        End Select
                    End Select
                End Select
            End Select
        End If
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 45 Then ' "\\-"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
            Call pvSetAdvance
            ParseChar = True
            Exit Function
        End If
        p415 = .BufPos
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = p395
        Else
            .BufPos = p415
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                Call pvSetAdvance
                ParseChar = True
                Exit Function
            Else
                .BufPos = p395
            End If
        End If
        .BufPos = p421
        If ConsoleTrace(ctx.BufPos + 1, "Char", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        Call pvSetExpected("char")
    End With
End Function

Public Function ParseRange() As Boolean
    Dim p385 As Long
    Dim p381 As Long
    Dim p378 As Long

    With ctx
        p385 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Range", 1, ctx.UserData)
        p381 = .BufPos
        If ParseChar() Then
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                p378 = .BufPos
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = p381
                Else
                    .BufPos = p378
                    If ParseChar() Then
                        Call ConsoleTrace(ctx.BufPos + 1, "Range", 2, ctx.UserData)
                        Call pvSetAdvance
                        ParseRange = True
                        Exit Function
                    Else
                        .BufPos = p381
                    End If
                End If
            Else
                .BufPos = p381
            End If
        Else
            .BufPos = p381
        End If
        If ParseChar() Then
            Call ConsoleTrace(ctx.BufPos + 1, "Range", 2, ctx.UserData)
            Call pvSetAdvance
            ParseRange = True
            Exit Function
        Else
            .BufPos = p381
        End If
        .BufPos = p385
        If ConsoleTrace(ctx.BufPos + 1, "Range", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseLEFTARROW() As Boolean
    Dim p430 As Long

    With ctx
        p430 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LEFTARROW", 1, ctx.UserData)
        If .BufData(.BufPos) = 60 And .BufData(.BufPos + 1) = 45 Then ' "<-"
            .BufPos = .BufPos + 2
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "LEFTARROW", 2, ctx.UserData)
            Call pvSetAdvance
            ParseLEFTARROW = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "LEFTARROW", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseSLASH() As Boolean
    Dim p441 As Long
    Dim p435 As Long

    With ctx
        p441 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SLASH", 1, ctx.UserData)
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            p435 = .BufPos
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = p441
            Else
                .BufPos = p435
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "SLASH", 2, ctx.UserData)
                Call pvSetAdvance
                ParseSLASH = True
                Exit Function
            End If
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SLASH", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseSpace() As Boolean
    Dim p568 As Long
    Dim p563 As Long

    With ctx
        p568 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Space", 1, ctx.UserData)
        p563 = .BufPos
        If .BufData(.BufPos) = 32 Then              ' " "
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "Space", 2, ctx.UserData)
            Call pvSetAdvance
            ParseSpace = True
            Exit Function
        End If
        If .BufData(.BufPos) = 9 Then               ' "\t"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "Space", 2, ctx.UserData)
            Call pvSetAdvance
            ParseSpace = True
            Exit Function
        End If
        If ParseEndOfLine() Then
            Call ConsoleTrace(ctx.BufPos + 1, "Space", 2, ctx.UserData)
            Call pvSetAdvance
            ParseSpace = True
            Exit Function
        Else
            .BufPos = p563
        End If
        .BufPos = p568
        If ConsoleTrace(ctx.BufPos + 1, "Space", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseLineComment() As Boolean
    Dim p543 As Long
    Dim p539 As Long

    With ctx
        p543 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 1, ctx.UserData)
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
            p539 = .BufPos
            If ParseEndOfLine() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p539
            End If
            If ParseEndOfFile() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p539
            End If
            .BufPos = p543
            GoTo L23
        End If
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
            p539 = .BufPos
            If ParseEndOfLine() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p539
            End If
            If ParseEndOfFile() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p539
            End If
            .BufPos = p543
            GoTo L23
        End If
        .BufPos = p543
L23:
        If ConsoleTrace(ctx.BufPos + 1, "LineComment", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseBlockComment() As Boolean
    Dim p559 As Long
    Dim p548 As Long
    Dim p550 As Long

    With ctx
        p559 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "BlockComment", 1, ctx.UserData)
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 42 Then ' "/*"
            .BufPos = .BufPos + 2
            Do
                p548 = .BufPos
                Select Case .BufData(.BufPos)
                Case 42, 47                         ' [*/]
                    If Not ParseBlockComment() Then
                        .BufPos = p548
                        p550 = .BufPos
                        If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                            .BufPos = p548
                            .BufPos = p548
                            Exit Do
                        Else
                            .BufPos = p550
                            If .BufPos < .BufSize Then
                                .BufPos = .BufPos + 1
                            Else
                                .BufPos = p548
                                .BufPos = p548
                                Exit Do
                            End If
                        End If
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                    Else
                        If Not ParseBlockComment() Then
                            .BufPos = p548
                            p550 = .BufPos
                            If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                                .BufPos = p548
                                .BufPos = p548
                                Exit Do
                            Else
                                .BufPos = p550
                                If .BufPos < .BufSize Then
                                    .BufPos = .BufPos + 1
                                Else
                                    .BufPos = p548
                                    .BufPos = p548
                                    Exit Do
                                End If
                            End If
                        End If
                    End If
                End Select
            Loop
            If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                .BufPos = .BufPos + 2
                Call ConsoleTrace(ctx.BufPos + 1, "BlockComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseBlockComment = True
                Exit Function
            End If
        End If
        If ConsoleTrace(ctx.BufPos + 1, "BlockComment", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseEndOfLine() As Boolean
    Dim p577 As Long

    With ctx
        p577 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 1, ctx.UserData)
        If .BufData(.BufPos) = 13 And .BufData(.BufPos + 1) = 10 Then ' "\r\n"
            .BufPos = .BufPos + 2
            Call ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 2, ctx.UserData)
            Call pvSetAdvance
            ParseEndOfLine = True
            Exit Function
        End If
        If .BufData(.BufPos) = 10 Then              ' "\n"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 2, ctx.UserData)
            Call pvSetAdvance
            ParseEndOfLine = True
            Exit Function
        End If
        If .BufData(.BufPos) = 13 Then              ' "\r"
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 2, ctx.UserData)
            Call pvSetAdvance
            ParseEndOfLine = True
            Exit Function
        End If
        .BufPos = p577
        If ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseCodeNC() As Boolean
    Dim p615 As Long
    Dim p607 As Long

    With ctx
        p615 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CodeNC", 1, ctx.UserData)
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Do
                p607 = .BufPos
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If Not ParseCodeNC() Then
                        .BufPos = p607
                        .BufPos = p607
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                    Else
                        If Not ParseCodeNC() Then
                            .BufPos = p607
                            .BufPos = p607
                            Exit Do
                        End If
                    End If
                End Select
            Loop
            If .BufData(.BufPos) = 125 Then         ' "}"
                .BufPos = .BufPos + 1
                Call ParseSpacing
                Call ConsoleTrace(ctx.BufPos + 1, "CodeNC", 2, ctx.UserData)
                Call pvSetAdvance
                ParseCodeNC = True
                Exit Function
            End If
        End If
        If ConsoleTrace(ctx.BufPos + 1, "CodeNC", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseEQUAL() As Boolean
    Dim p657 As Long

    With ctx
        p657 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "EQUAL", 1, ctx.UserData)
        If .BufData(.BufPos) = 61 Then              ' "="
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "EQUAL", 2, ctx.UserData)
            Call pvSetAdvance
            ParseEQUAL = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "EQUAL", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseBAR() As Boolean
    Dim p665 As Long

    With ctx
        p665 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "BAR", 1, ctx.UserData)
        If .BufData(.BufPos) = 124 Then             ' "|"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "BAR", 2, ctx.UserData)
            Call pvSetAdvance
            ParseBAR = True
            Exit Function
        End If
        If ConsoleTrace(ctx.BufPos + 1, "BAR", 3, ctx.UserData) And False Then
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
    Case ucsAct_3_Rule
         lNode = .Pop: .SetExpression .Pop, lNode
    Case ucsAct_2_Rule
         .RuleDescription(.Top) = Mid$(ctx.Contents, lOffset, lSize)
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
    Case ucsAct_1_PrimaryExpr
         .Push .MakeReference(lOffset, Mid$(ctx.Contents, lOffset, lSize))
    Case ucsAct_3_PrimaryExpr
         .SetIgnoreCase .Top, (Mid$(ctx.Contents, lOffset, lSize) = "i")
    Case ucsAct_2_PrimaryExpr
         .Push .MakeString(lOffset, Mid$(ctx.Contents, lOffset, lSize))
    Case ucsAct_5_PrimaryExpr
         .SetIgnoreCase .Top, (Mid$(ctx.Contents, lOffset, lSize) = "i")
    Case ucsAct_4_PrimaryExpr
         .Push .MakeClass(lOffset, Mid$(ctx.Contents, lOffset, lSize))
    Case ucsAct_6_PrimaryExpr
         .Push .MakeDot(lOffset)
    Case ucsAct_7_PrimaryExpr
         .Push .MakeAction(lOffset, Mid$(ctx.Contents, lOffset, lSize))
    Case ucsAct_8_PrimaryExpr
         .Push .MakePredicate(lOffset, "YY_BEGIN", AlwaysTrue:=True)
    Case ucsAct_9_PrimaryExpr
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

Public Function VbPegGetParseErrors() As Collection
    Set VbPegGetParseErrors = ctx.ParseErrors
End Function

Public Function VbPegGetContents() As String
    VbPegGetContents = ctx.Contents
End Function

Private Function CheckSettingName(sName As String) As Boolean
    Const STR_ALL = "||private|public|modulename|userdatatype|vardatatype|members|start|ignorecase|prolog|epilog|trace|append|declares|"
    If InStr(1, STR_ALL, "|" & sName & "|", vbTextCompare) = 0 Then
        AppendError "Setting '" & sName & "' is unrecognized", ctx.CaptureBegin + 1, "warning"
    End If
    CheckSettingName = True
End Function
