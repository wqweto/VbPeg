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
' Auto-generated on 23.7.2018 18:29:35
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
    ParserVersion = "23.7.2018 18:29:35"
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
    Dim p21 As Long
    Dim q21 As Long
    Dim i12 As Long
    Dim p11 As Long
    Dim q11 As Long
    Dim p17 As Long
    Dim q17 As Long

    With ctx
        p21 = .BufPos
        q21 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Grammar", 1, ctx.UserData)
        Call ParseSpacing
        For i12 = 0 To LNG_MAXINT
            p11 = .BufPos
            q11 = .ThunkPos
            If Not ParseSetting() Then
                .BufPos = p11
                .ThunkPos = q11
                If Not ParseRule() Then
                    .BufPos = p11
                    .ThunkPos = q11
                    .BufPos = p11
                    .ThunkPos = q11
                    Exit For
                End If
            End If
        Next
        If i12 <> 0 Then
            p17 = .BufPos
            q17 = .ThunkPos
            If ParseEndOfFile() Then
                Call ConsoleTrace(ctx.BufPos + 1, "Grammar", 2, ctx.UserData)
                ParseGrammar = True
                Exit Function
            Else
                .BufPos = p17
                .ThunkPos = q17
            End If
            If AppendError("Expected rule or setting", ctx.BufPos + 1) Then
                Call ConsoleTrace(ctx.BufPos + 1, "Grammar", 2, ctx.UserData)
                ParseGrammar = True
                Exit Function
            End If
            .BufPos = p21
            .ThunkPos = q21
        Else
            .BufPos = p21
            .ThunkPos = q21
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Grammar", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Sub ParseSpacing()
    Dim p523 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "Spacing", 1, ctx.UserData)
        Do
            p523 = .BufPos
            If Not ParseSpace() Then
                .BufPos = p523
                If Not ParseLineComment() Then
                    .BufPos = p523
                    If Not ParseBlockComment() Then
                        .BufPos = p523
                        .BufPos = p523
                        Exit Do
                    End If
                End If
            End If
        Loop
        Call ConsoleTrace(ctx.BufPos + 1, "Spacing", 2, ctx.UserData)
    End With
End Sub

Public Function ParseSetting() As Boolean
    Dim p239 As Long
    Dim q239 As Long
    Dim p235 As Long
    Dim q235 As Long

    With ctx
        p239 = .BufPos
        q239 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 1, ctx.UserData)
        If ParseSettingName() Then
            If CheckSettingName(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                pvPushThunk ucsAct_1_Setting, .CaptureBegin, .CaptureEnd
                If ClearErrors(ctx.BufPos + 1) And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                    p235 = .BufPos
                    q235 = .ThunkPos
                    If ParseSettingValue() Then
                        pvPushThunk ucsAct_2_Setting, .CaptureBegin, .CaptureEnd
                        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 2, ctx.UserData)
                        ParseSetting = True
                        Exit Function
                    Else
                        .BufPos = p235
                        .ThunkPos = q235
                    End If
                    If AppendError("Invalid value for setting '" & .UserData.Tag & "'", ctx.BufPos + 1) Then
                        Call ConsoleTrace(ctx.BufPos + 1, "Setting", 2, ctx.UserData)
                        ParseSetting = True
                        Exit Function
                    End If
                    .BufPos = p239
                    .ThunkPos = q239
                Else
                    .BufPos = p239
                    .ThunkPos = q239
                End If
            Else
                .BufPos = p239
                .ThunkPos = q239
            End If
        Else
            .BufPos = p239
            .ThunkPos = q239
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Setting", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseRule() As Boolean
    Dim p47 As Long
    Dim q47 As Long
    Dim p31 As Long
    Dim q31 As Long
    Dim p42 As Long
    Dim q42 As Long

    With ctx
        p47 = .BufPos
        q47 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Rule", 1, ctx.UserData)
        If ParseIdentifier() Then
            pvPushThunk ucsAct_1_Rule, .CaptureBegin, .CaptureEnd
            If ClearErrors(ctx.BufPos + 1) And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                p31 = .BufPos
                q31 = .ThunkPos
                If ParseLiteral() Then
                    pvPushThunk ucsAct_2_Rule, .CaptureBegin, .CaptureEnd
                Else
                    .BufPos = p31
                    .ThunkPos = q31
                End If
                If ParseAssignLiteral() Then
                    Call ParseExpression
                    pvPushThunk ucsAct_3_Rule, .CaptureBegin, .CaptureEnd
                    p42 = .BufPos
                    q42 = .ThunkPos
                    If Not ParseSEMI() Then
                        .BufPos = p42
                        .ThunkPos = q42
                    End If
                    Call ConsoleTrace(ctx.BufPos + 1, "Rule", 2, ctx.UserData)
                    ParseRule = True
                    Exit Function
                Else
                    .BufPos = p47
                    .ThunkPos = q47
                End If
            Else
                .BufPos = p47
                .ThunkPos = q47
            End If
        Else
            .BufPos = p47
            .ThunkPos = q47
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Rule", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseEndOfFile() As Boolean
    Dim p586 As Long
    Dim p582 As Long

    With ctx
        p586 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "EndOfFile", 1, ctx.UserData)
        p582 = .BufPos
        If .BufPos < .BufSize Then
            .BufPos = p586
        Else
            .BufPos = p582
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
    Dim p278 As Long
    Dim lCaptureBegin As Long
    Dim p271 As Long
    Dim lCaptureEnd As Long

    With ctx
        p278 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Identifier", 1, ctx.UserData)
        lCaptureBegin = .BufPos
        If ParseIdentStart() Then
            Do
                p271 = .BufPos
                If Not ParseIdentCont() Then
                    .BufPos = p271
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
            .BufPos = p278
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Identifier", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseLiteral() As Boolean
    Dim p298 As Long

    With ctx
        p298 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Literal", 1, ctx.UserData)
        If ParseLiteralNoSpace() Then
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "Literal", 2, ctx.UserData)
            ParseLiteral = True
            Exit Function
        Else
            .BufPos = p298
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Literal", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseAssignLiteral() As Boolean
    Dim p642 As Long
    Dim p638 As Long

    With ctx
        p642 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 1, ctx.UserData)
        p638 = .BufPos
        If ParseLEFTARROW() Then
            Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 2, ctx.UserData)
            ParseAssignLiteral = True
            Exit Function
        Else
            .BufPos = p638
        End If
        If ParseEQUAL() Then
            Call ConsoleTrace(ctx.BufPos + 1, "AssignLiteral", 2, ctx.UserData)
            ParseAssignLiteral = True
            Exit Function
        Else
            .BufPos = p638
        End If
        .BufPos = p642
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
    Dim p683 As Long

    With ctx
        p683 = .BufPos
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
    Dim p61 As Long
    Dim q61 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceExpr", 1, ctx.UserData)
        Call ParseSequenceExpr
        Do
            p61 = .BufPos
            q61 = .ThunkPos
            If Not ParseChoiceLiteral() Then
                .BufPos = p61
                .ThunkPos = q61
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
    Dim p78 As Long
    Dim q78 As Long
    Dim p74 As Long
    Dim q74 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "SequenceExpr", 1, ctx.UserData)
        p78 = .BufPos
        q78 = .ThunkPos
        If ParseErrorExpr() Then
            Do
                p74 = .BufPos
                q74 = .ThunkPos
                If Not ParseErrorExpr() Then
                    .BufPos = p74
                    .ThunkPos = q74
                    Exit Do
                End If
                pvPushThunk ucsAct_1_SequenceExpr, .CaptureBegin, .CaptureEnd
            Loop
        Else
            .BufPos = p78
            .ThunkPos = q78
            pvPushThunk ucsAct_2_SequenceExpr, .CaptureBegin, .CaptureEnd
        End If
        Call ConsoleTrace(ctx.BufPos + 1, "SequenceExpr", 2, ctx.UserData)
    End With
End Sub

Public Function ParseChoiceLiteral() As Boolean
    Dim p651 As Long
    Dim p647 As Long

    With ctx
        p651 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 1, ctx.UserData)
        p647 = .BufPos
        If ParseSLASH() Then
            Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 2, ctx.UserData)
            ParseChoiceLiteral = True
            Exit Function
        Else
            .BufPos = p647
        End If
        If ParseBAR() Then
            Call ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 2, ctx.UserData)
            ParseChoiceLiteral = True
            Exit Function
        Else
            .BufPos = p647
        End If
        .BufPos = p651
        If ConsoleTrace(ctx.BufPos + 1, "ChoiceLiteral", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseErrorExpr() As Boolean
    Dim p103 As Long
    Dim q103 As Long
    Dim p89 As Long
    Dim q89 As Long
    Dim p96 As Long
    Dim q96 As Long

    With ctx
        p103 = .BufPos
        q103 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 1, ctx.UserData)
        If ParseVariableExpr() Then
            p89 = .BufPos
            q89 = .ThunkPos
            If ParseCARET() Then
                If ClearErrors(ctx.BufPos + 1) Then
                    p96 = .BufPos
                    q96 = .ThunkPos
                    If ParseIdentifier() Then
                        pvPushThunk ucsAct_1_ErrorExpr, .CaptureBegin, .CaptureEnd
                        GoTo L10
                    Else
                        .BufPos = p96
                        .ThunkPos = q96
                    End If
                    If ParseLiteral() Then
                        pvPushThunk ucsAct_2_ErrorExpr, .CaptureBegin, .CaptureEnd
                        GoTo L10
                    Else
                        .BufPos = p96
                        .ThunkPos = q96
                    End If
                    If AppendError("Expected rule or literal for error label", ctx.BufPos + 1) Then
                        GoTo L10
                    End If
                    .BufPos = p89
                    .ThunkPos = q89
                Else
                    .BufPos = p89
                    .ThunkPos = q89
                End If
            Else
                .BufPos = p89
                .ThunkPos = q89
            End If
L10:
            Call ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 2, ctx.UserData)
            ParseErrorExpr = True
            Exit Function
        Else
            .BufPos = p103
            .ThunkPos = q103
        End If
        If ConsoleTrace(ctx.BufPos + 1, "ErrorExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseVariableExpr() As Boolean
    Dim p118 As Long
    Dim q118 As Long
    Dim p114 As Long
    Dim q114 As Long

    With ctx
        p118 = .BufPos
        q118 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 1, ctx.UserData)
        p114 = .BufPos
        q114 = .ThunkPos
        If ParseIdentifier() Then
            pvPushThunk ucsAct_1_VariableExpr, .CaptureBegin, .CaptureEnd
            If ParseCOLON() Then
                If ParsePrefixExpr() Then
                    pvPushThunk ucsAct_2_VariableExpr, .CaptureBegin, .CaptureEnd
                    Call ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 2, ctx.UserData)
                    ParseVariableExpr = True
                    Exit Function
                Else
                    .BufPos = p114
                    .ThunkPos = q114
                End If
            Else
                .BufPos = p114
                .ThunkPos = q114
            End If
        Else
            .BufPos = p114
            .ThunkPos = q114
        End If
        If ParsePrefixExpr() Then
            Call ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 2, ctx.UserData)
            ParseVariableExpr = True
            Exit Function
        Else
            .BufPos = p114
            .ThunkPos = q114
        End If
        .BufPos = p118
        .ThunkPos = q118
        If ConsoleTrace(ctx.BufPos + 1, "VariableExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseCARET() As Boolean
    Dim p699 As Long

    With ctx
        p699 = .BufPos
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
    Dim p675 As Long

    With ctx
        p675 = .BufPos
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
    Dim p141 As Long
    Dim q141 As Long
    Dim p131 As Long
    Dim q131 As Long

    With ctx
        p141 = .BufPos
        q141 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 1, ctx.UserData)
        p131 = .BufPos
        q131 = .ThunkPos
        If ParseAND() Then
            If ParseCodeBlock() Then
                pvPushThunk ucsAct_1_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p131
                .ThunkPos = q131
            End If
        Else
            .BufPos = p131
            .ThunkPos = q131
        End If
        If ParseAND() Then
            If ParseSuffixExpr() Then
                pvPushThunk ucsAct_2_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p131
                .ThunkPos = q131
            End If
        Else
            .BufPos = p131
            .ThunkPos = q131
        End If
        If ParseNOT() Then
            If ParseSuffixExpr() Then
                pvPushThunk ucsAct_3_PrefixExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p131
                .ThunkPos = q131
            End If
        Else
            .BufPos = p131
            .ThunkPos = q131
        End If
        If ParseSuffixExpr() Then
            Call ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 2, ctx.UserData)
            ParsePrefixExpr = True
            Exit Function
        Else
            .BufPos = p131
            .ThunkPos = q131
        End If
        .BufPos = p141
        .ThunkPos = q141
        If ConsoleTrace(ctx.BufPos + 1, "PrefixExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseAND() As Boolean
    Dim p451 As Long

    With ctx
        p451 = .BufPos
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
    Dim p604 As Long
    Dim p589 As Long
    Dim lCaptureBegin As Long
    Dim p596 As Long
    Dim lCaptureEnd As Long

    With ctx
        p604 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CodeBlock", 1, ctx.UserData)
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Do
                p589 = .BufPos
                If Not ParseEndOfLine() Then
                    .BufPos = p589
                    Exit Do
                End If
            Loop
            lCaptureBegin = .BufPos
            Do
                p596 = .BufPos
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If Not ParseCodeNC() Then
                        .BufPos = p596
                        .BufPos = p596
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                    Else
                        If Not ParseCodeNC() Then
                            .BufPos = p596
                            .BufPos = p596
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
    Dim p163 As Long
    Dim q163 As Long
    Dim p153 As Long
    Dim q153 As Long

    With ctx
        p163 = .BufPos
        q163 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 1, ctx.UserData)
        If ParsePrimaryExpr() Then
            p153 = .BufPos
            q153 = .ThunkPos
            If ParseQUESTION() Then
                pvPushThunk ucsAct_1_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L12
            Else
                .BufPos = p153
                .ThunkPos = q153
            End If
            If ParseSTAR() Then
                pvPushThunk ucsAct_2_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L12
            Else
                .BufPos = p153
                .ThunkPos = q153
            End If
            If ParsePLUS() Then
                pvPushThunk ucsAct_3_SuffixExpr, .CaptureBegin, .CaptureEnd
                GoTo L12
            Else
                .BufPos = p153
                .ThunkPos = q153
            End If
            .BufPos = p153
            .ThunkPos = q153
L12:
            Call ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 2, ctx.UserData)
            ParseSuffixExpr = True
            Exit Function
        Else
            .BufPos = p163
            .ThunkPos = q163
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SuffixExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseNOT() As Boolean
    Dim p459 As Long

    With ctx
        p459 = .BufPos
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
    Dim p222 As Long
    Dim q222 As Long
    Dim p184 As Long
    Dim q184 As Long
    Dim p171 As Long
    Dim q171 As Long
    Dim p167 As Long
    Dim p183 As Long
    Dim q183 As Long
    Dim p192 As Long
    Dim q192 As Long
    Dim p201 As Long
    Dim q201 As Long

    With ctx
        p222 = .BufPos
        q222 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 1, ctx.UserData)
        p184 = .BufPos
        q184 = .ThunkPos
        If ParseIdentifier() Then
            p171 = .BufPos
            q171 = .ThunkPos
            p167 = .BufPos
            If Not ParseLiteralNC() Then
                .BufPos = p167
            End If
            If ParseAssignLiteral() Then
                .BufPos = p184
                .ThunkPos = q184
            Else
                .BufPos = p171
                .ThunkPos = q171
                pvPushThunk ucsAct_1_PrimaryExpr, .CaptureBegin, .CaptureEnd
                Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                ParsePrimaryExpr = True
                Exit Function
            End If
        Else
            .BufPos = p184
            .ThunkPos = q184
        End If
        If ParseOPEN() Then
            If ClearErrors(ctx.BufPos + 1) Then
                p183 = .BufPos
                q183 = .ThunkPos
                Call ParseExpression
                If ParseCLOSE() Then
                    Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                    ParsePrimaryExpr = True
                    Exit Function
                Else
                    .BufPos = p183
                    .ThunkPos = q183
                End If
                If AppendError("Invalid expression", ctx.BufPos + 1) Then
                    Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
                    ParsePrimaryExpr = True
                    Exit Function
                End If
                .BufPos = p184
                .ThunkPos = q184
            Else
                .BufPos = p184
                .ThunkPos = q184
            End If
        Else
            .BufPos = p184
            .ThunkPos = q184
        End If
        If ParseLiteralNoSpace() Then
            pvPushThunk ucsAct_2_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p192 = .BufPos
            q192 = .ThunkPos
            If ParseIgnoreCase() Then
                pvPushThunk ucsAct_3_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Else
                .BufPos = p192
                .ThunkPos = q192
            End If
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p184
            .ThunkPos = q184
        End If
        If ParseClassNoSpace() Then
            pvPushThunk ucsAct_4_PrimaryExpr, .CaptureBegin, .CaptureEnd
            p201 = .BufPos
            q201 = .ThunkPos
            If ParseIgnoreCase() Then
                pvPushThunk ucsAct_5_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Else
                .BufPos = p201
                .ThunkPos = q201
            End If
            Call ParseSpacing
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p184
            .ThunkPos = q184
        End If
        If ParseDOT() Then
            pvPushThunk ucsAct_6_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p184
            .ThunkPos = q184
        End If
        If ParseCodeBlock() Then
            pvPushThunk ucsAct_7_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p184
            .ThunkPos = q184
        End If
        If ParseBEGIN() Then
            pvPushThunk ucsAct_8_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p184
            .ThunkPos = q184
        End If
        If ParseEND() Then
            pvPushThunk ucsAct_9_PrimaryExpr, .CaptureBegin, .CaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 2, ctx.UserData)
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p184
            .ThunkPos = q184
        End If
        .BufPos = p222
        .ThunkPos = q222
        If ConsoleTrace(ctx.BufPos + 1, "PrimaryExpr", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseQUESTION() As Boolean
    Dim p467 As Long

    With ctx
        p467 = .BufPos
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
    Dim p475 As Long

    With ctx
        p475 = .BufPos
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
    Dim p483 As Long

    With ctx
        p483 = .BufPos
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
    Dim p349 As Long
    Dim p345 As Long
    Dim p331 As Long
    Dim p329 As Long
    Dim p340 As Long
    Dim p338 As Long

    With ctx
        p349 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 1, ctx.UserData)
        p345 = .BufPos
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            Do
                p331 = .BufPos
                p329 = .BufPos
                If .BufData(.BufPos) = 39 Then      ' "'"
                    .BufPos = p331
                    Exit Do
                Else
                    .BufPos = p329
                End If
                If Not ParseChar() Then
                    .BufPos = p331
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
                p340 = .BufPos
                p338 = .BufPos
                If .BufData(.BufPos) = 34 Then      ' """
                    .BufPos = p340
                    Exit Do
                Else
                    .BufPos = p338
                End If
                If Not ParseChar() Then
                    .BufPos = p340
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
        .BufPos = p349
        If ConsoleTrace(ctx.BufPos + 1, "LiteralNC", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseOPEN() As Boolean
    Dim p491 As Long

    With ctx
        p491 = .BufPos
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
    Dim p499 As Long

    With ctx
        p499 = .BufPos
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
    Dim p325 As Long
    Dim p321 As Long
    Dim lCaptureBegin As Long
    Dim p307 As Long
    Dim p304 As Long
    Dim lCaptureEnd As Long
    Dim p317 As Long
    Dim p315 As Long

    With ctx
        p325 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 1, ctx.UserData)
        p321 = .BufPos
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            lCaptureBegin = .BufPos
            Do
                p307 = .BufPos
                p304 = .BufPos
                If .BufData(.BufPos) = 39 Then      ' "'"
                    .BufPos = p307
                    Exit Do
                Else
                    .BufPos = p304
                End If
                If Not ParseChar() Then
                    .BufPos = p307
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
                p317 = .BufPos
                p315 = .BufPos
                If .BufData(.BufPos) = 34 Then      ' """
                    .BufPos = p317
                    Exit Do
                Else
                    .BufPos = p315
                End If
                If Not ParseChar() Then
                    .BufPos = p317
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
        .BufPos = p325
        If ConsoleTrace(ctx.BufPos + 1, "LiteralNoSpace", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseIgnoreCase() As Boolean
    Dim p358 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        p358 = .BufPos
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
    Dim p374 As Long
    Dim lCaptureBegin As Long
    Dim p367 As Long
    Dim p364 As Long
    Dim lCaptureEnd As Long

    With ctx
        p374 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "ClassNoSpace", 1, ctx.UserData)
        If .BufData(.BufPos) = 91 Then              ' "["
            .BufPos = .BufPos + 1
            lCaptureBegin = .BufPos
            Do
                p367 = .BufPos
                p364 = .BufPos
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = p367
                    Exit Do
                Else
                    .BufPos = p364
                End If
                If Not ParseRange() Then
                    .BufPos = p367
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
    Dim p507 As Long

    With ctx
        p507 = .BufPos
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
    Dim p625 As Long

    With ctx
        p625 = .BufPos
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
    Dim p633 As Long

    With ctx
        p633 = .BufPos
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
    Dim p252 As Long
    Dim p245 As Long

    With ctx
        p252 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 1, ctx.UserData)
        If ParseAT() Then
            p245 = .BufPos
            If ParseIdentifier() Then
                Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p245
            End If
            If ParseLiteral() Then
                Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p245
            End If
            Call ParseEMPTY
            Call ConsoleTrace(ctx.BufPos + 1, "SettingName", 2, ctx.UserData)
            ParseSettingName = True
            Exit Function
            .BufPos = p252
        Else
            .BufPos = p252
        End If
        If ConsoleTrace(ctx.BufPos + 1, "SettingName", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseSettingValue() As Boolean
    Dim p264 As Long
    Dim p256 As Long
    Dim p258 As Long

    With ctx
        p264 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 1, ctx.UserData)
        p256 = .BufPos
        If ParseIdentifier() Then
            p258 = .BufPos
            If Not ParseSEMI() Then
                .BufPos = p258
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p256
        End If
        If ParseLiteral() Then
            p258 = .BufPos
            If Not ParseSEMI() Then
                .BufPos = p258
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p256
        End If
        If ParseCodeBlock() Then
            p258 = .BufPos
            If Not ParseSEMI() Then
                .BufPos = p258
            End If
            Call ConsoleTrace(ctx.BufPos + 1, "SettingValue", 2, ctx.UserData)
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p256
        End If
        .BufPos = p264
        If ConsoleTrace(ctx.BufPos + 1, "SettingValue", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseAT() As Boolean
    Dim p691 As Long

    With ctx
        p691 = .BufPos
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
    Dim p284 As Long

    With ctx
        p284 = .BufPos
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
    Dim p290 As Long

    With ctx
        p290 = .BufPos
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
    Dim p423 As Long
    Dim p397 As Long
    Dim p417 As Long

    With ctx
        p423 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Char", 1, ctx.UserData)
        p397 = .BufPos
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
        p417 = .BufPos
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = p397
        Else
            .BufPos = p417
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "Char", 2, ctx.UserData)
                Call pvSetAdvance
                ParseChar = True
                Exit Function
            Else
                .BufPos = p397
            End If
        End If
        .BufPos = p423
        If ConsoleTrace(ctx.BufPos + 1, "Char", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        Call pvSetExpected("char")
    End With
End Function

Public Function ParseRange() As Boolean
    Dim p387 As Long
    Dim p383 As Long
    Dim p380 As Long

    With ctx
        p387 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Range", 1, ctx.UserData)
        p383 = .BufPos
        If ParseChar() Then
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                p380 = .BufPos
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = p383
                Else
                    .BufPos = p380
                    If ParseChar() Then
                        Call ConsoleTrace(ctx.BufPos + 1, "Range", 2, ctx.UserData)
                        Call pvSetAdvance
                        ParseRange = True
                        Exit Function
                    Else
                        .BufPos = p383
                    End If
                End If
            Else
                .BufPos = p383
            End If
        Else
            .BufPos = p383
        End If
        If ParseChar() Then
            Call ConsoleTrace(ctx.BufPos + 1, "Range", 2, ctx.UserData)
            Call pvSetAdvance
            ParseRange = True
            Exit Function
        Else
            .BufPos = p383
        End If
        .BufPos = p387
        If ConsoleTrace(ctx.BufPos + 1, "Range", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Function ParseLEFTARROW() As Boolean
    Dim p432 As Long

    With ctx
        p432 = .BufPos
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
    Dim p443 As Long
    Dim p437 As Long

    With ctx
        p443 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "SLASH", 1, ctx.UserData)
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            p437 = .BufPos
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = p443
            Else
                .BufPos = p437
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
    Dim p570 As Long
    Dim p565 As Long

    With ctx
        p570 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Space", 1, ctx.UserData)
        p565 = .BufPos
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
            .BufPos = p565
        End If
        .BufPos = p570
        If ConsoleTrace(ctx.BufPos + 1, "Space", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseLineComment() As Boolean
    Dim p545 As Long
    Dim p541 As Long

    With ctx
        p545 = .BufPos
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
            p541 = .BufPos
            If ParseEndOfLine() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p541
            End If
            If ParseEndOfFile() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p541
            End If
            .BufPos = p545
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
            p541 = .BufPos
            If ParseEndOfLine() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p541
            End If
            If ParseEndOfFile() Then
                Call ConsoleTrace(ctx.BufPos + 1, "LineComment", 2, ctx.UserData)
                Call pvSetAdvance
                ParseLineComment = True
                Exit Function
            Else
                .BufPos = p541
            End If
            .BufPos = p545
            GoTo L23
        End If
        .BufPos = p545
L23:
        If ConsoleTrace(ctx.BufPos + 1, "LineComment", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseBlockComment() As Boolean
    Dim p561 As Long
    Dim p550 As Long
    Dim p552 As Long

    With ctx
        p561 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "BlockComment", 1, ctx.UserData)
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 42 Then ' "/*"
            .BufPos = .BufPos + 2
            Do
                p550 = .BufPos
                Select Case .BufData(.BufPos)
                Case 42, 47                         ' [*/]
                    If Not ParseBlockComment() Then
                        .BufPos = p550
                        p552 = .BufPos
                        If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                            .BufPos = p550
                            .BufPos = p550
                            Exit Do
                        Else
                            .BufPos = p552
                            If .BufPos < .BufSize Then
                                .BufPos = .BufPos + 1
                            Else
                                .BufPos = p550
                                .BufPos = p550
                                Exit Do
                            End If
                        End If
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                    Else
                        If Not ParseBlockComment() Then
                            .BufPos = p550
                            p552 = .BufPos
                            If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                                .BufPos = p550
                                .BufPos = p550
                                Exit Do
                            Else
                                .BufPos = p552
                                If .BufPos < .BufSize Then
                                    .BufPos = .BufPos + 1
                                Else
                                    .BufPos = p550
                                    .BufPos = p550
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
    Dim p579 As Long

    With ctx
        p579 = .BufPos
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
        .BufPos = p579
        If ConsoleTrace(ctx.BufPos + 1, "EndOfLine", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseCodeNC() As Boolean
    Dim p617 As Long
    Dim p609 As Long

    With ctx
        p617 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "CodeNC", 1, ctx.UserData)
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Do
                p609 = .BufPos
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If Not ParseCodeNC() Then
                        .BufPos = p609
                        .BufPos = p609
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                    Else
                        If Not ParseCodeNC() Then
                            .BufPos = p609
                            .BufPos = p609
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
    Dim p659 As Long

    With ctx
        p659 = .BufPos
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
    Dim p667 As Long

    With ctx
        p667 = .BufPos
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
