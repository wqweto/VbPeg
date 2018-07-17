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
' Auto-generated on 17.7.2018 18:26:18
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Const LOCALE_USER_DEFAULT           As Long = &H400
Private Const NORM_IGNORECASE               As Long = 1
Private Const CSTR_EQUAL                    As Long = 2

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function RtlCompareMemory Lib "ntdll" (Source1 As Any, Source2 As Any, ByVal Length As Long) As Long
Private Declare Function CompareStringW Lib "kernel32" (ByVal Locale As Long, ByVal dwCmpFlags As Long, lpString1 As Any, ByVal cchCount1 As Long, lpString2 As Any, ByVal cchCount2 As Long) As Long

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
    LastError           As String
    LastBufPos          As Long
    UserData            As cTree
    '-- from @members setting
    Errors              As Collection
End Type

Private ctx                     As UcsParserType

'=========================================================================
' Properties
'=========================================================================

Property Get LastError() As String
    LastError = ctx.LastError
End Property

Property Get LastBufPos() As Long
    LastBufPos = ctx.LastBufPos
End Property

Property Get ParserVersion() As String
    ParserVersion = "17.7.2018 18:26:18"
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Match(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As cTree) As Long
    If BeginMatch(sSubject, StartPos, UserData) Then
        If ParseGrammar() Then
            Match = EndMatch(UserData)
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

Public Function EndMatch(Optional UserData As cTree) As Long
    Dim lIdx            As Long
    Dim uEmpty          As UcsParserType
    
    With ctx
        For lIdx = 0 To .ThunkPos - 1
            With .ThunkData(lIdx)
                pvImplAction .Action, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin
            End With
        Next
        If IsObject(.UserData) Then
            Set UserData = .UserData
        Else
            UserData = .UserData
        End If
        EndMatch = .BufPos + 1
    End With
    uEmpty.LastError = ctx.LastError
    uEmpty.LastBufPos = ctx.LastBufPos
    ctx = uEmpty
End Function

Private Sub pvPushAction(ByVal eAction As UcsParserActionsEnum)
    pvPushThunk eAction, ctx.CaptureBegin, ctx.CaptureEnd
End Sub

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

Private Function pvMatchString(sText As String) As Boolean
    With ctx
        If .BufPos + Len(sText) <= .BufSize Then
            pvMatchString = RtlCompareMemory(.BufData(.BufPos), ByVal StrPtr(sText), LenB(sText)) = LenB(sText)
        End If
    End With
End Function

Private Function pvMatchIgnoreCase(sText As String) As Boolean
    With ctx
        If .BufPos + Len(sText) <= .BufSize Then
            pvMatchIgnoreCase = CompareStringW(LOCALE_USER_DEFAULT, NORM_IGNORECASE, ByVal StrPtr(sText), Len(sText), .BufData(.BufPos), Len(sText)) = CSTR_EQUAL
        End If
    End With
End Function

'= generated functions ===================================================

Public Function ParseGrammar() As Boolean
    Dim p10 As Long
    Dim q10 As Long
    Dim i9 As Long
    Dim p8 As Long
    Dim q8 As Long

    With ctx
        p10 = .BufPos
        q10 = .ThunkPos
        Call ParseSpacing
        For i9 = 0 To LNG_MAXINT
            p8 = .BufPos
            q8 = .ThunkPos
            If ParseSetting() Then
                GoTo L1
            Else
                .BufPos = p8
                .ThunkPos = q8
                If ParseRule() Then
                    GoTo L1
                Else
                    .BufPos = p8
                    .ThunkPos = q8
                    Exit For
                End If
            End If
L1:
        Next
        If i9 <> 0 Then
            If ParseEndOfFile() Then
                ParseGrammar = True
                Exit Function
            End If
            If  AppendError("Expected rule or setting")  Then
                ParseGrammar = True
                Exit Function
            End If
            .BufPos = p10
            .ThunkPos = q10
        Else
            .BufPos = p10
            .ThunkPos = q10
        End If
    End With
End Function

Public Sub ParseSpacing()
    Do
        If ParseSpace() Then
            GoTo L2
        Else
            If ParseComment() Then
                GoTo L2
            Else
                Exit Do
            End If
        End If
L2:
    Loop
End Sub

Public Function ParseSetting() As Boolean
    Dim p174 As Long
    Dim q174 As Long
    Dim p179 As Long
    Dim q179 As Long

    With ctx
        p174 = .BufPos
        q174 = .ThunkPos
        If ParseSettingName() Then
            pvPushAction ucsAct_1_Setting
            If  ClearErrors() And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin))  Then
                p179 = .BufPos
                q179 = .ThunkPos
                If ParseSettingValue() Then
                    pvPushAction ucsAct_2_Setting
                    ParseSetting = True
                    Exit Function
                Else
                    .BufPos = p179
                    .ThunkPos = q179
                End If
                If  AppendError("Invalid value for setting '" & .UserData.Tag & "'")  Then
                    ParseSetting = True
                    Exit Function
                End If
                .BufPos = p174
                .ThunkPos = q174
            Else
                .BufPos = p174
                .ThunkPos = q174
            End If
        Else
            .BufPos = p174
            .ThunkPos = q174
        End If
    End With
End Function

Public Function ParseRule() As Boolean
    Dim p18 As Long
    Dim q18 As Long
    Dim p23 As Long
    Dim q23 As Long

    With ctx
        p18 = .BufPos
        q18 = .ThunkPos
        If ParseIdentifier() Then
            pvPushAction ucsAct_1_Rule
            If  ClearErrors() And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin))  Then
                p23 = .BufPos
                q23 = .ThunkPos
                If ParseLiteral() Then
                    pvPushAction ucsAct_2_Rule
                    GoTo L6
                Else
                    .BufPos = p23
                    .ThunkPos = q23
                End If
L6:
                If ParseAssignLiteral() Then
                    Call ParseChoiceExpr
                    pvPushAction ucsAct_3_Rule
                    Call ParseSEMI
                    ParseRule = True
                    Exit Function
                    If  AppendError("Invalid definition for rule '" & .UserData.Tag & "'")  Then
                        Call ParseSEMI
                        ParseRule = True
                        Exit Function
                    End If
                Else
                    .BufPos = p18
                    .ThunkPos = q18
                End If
            Else
                .BufPos = p18
                .ThunkPos = q18
            End If
        Else
            .BufPos = p18
            .ThunkPos = q18
        End If
    End With
End Function

Public Function ParseEndOfFile() As Boolean
    With ctx
        If Not .BufPos < .BufSize Then
            ParseEndOfFile = True
        End If
    End With
End Function

Public Function ParseIdentifier() As Boolean
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        Select Case .BufData(.BufPos)
        Case 97 To 122, 65 To 90, 95, 92, 45        ' [a-zA-Z_\-]
            .BufPos = .BufPos + 1
            Do
                Select Case .BufData(.BufPos)
                Case 97 To 122, 65 To 90, 95, 92 To 48, 45, 57 ' [a-zA-Z_\-0-9]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit Do
                End Select
            Loop
            lCaptureEnd = .BufPos
            Call ParseSpacing
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseIdentifier = True
        End Select
    End With
End Function

Public Function ParseLiteral() As Boolean
    Dim p213 As Long
    Dim q213 As Long

    With ctx
        p213 = .BufPos
        q213 = .ThunkPos
        If ParseLiteralNoSpace() Then
            Call ParseSpacing
            ParseLiteral = True
            Exit Function
        Else
            .BufPos = p213
            .ThunkPos = q213
        End If
    End With
End Function

Public Function ParseAssignLiteral() As Boolean
    If ParseLEFTARROW() Then
        ParseAssignLiteral = True
        Exit Function
    End If
    If ParseEQUAL() Then
        ParseAssignLiteral = True
    End If
End Function

Private Function ParseSEMI() As Boolean
    With ctx
        If .BufData(.BufPos) = 59 Then              ' ";"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseSEMI = True
        End If
    End With
End Function

Public Sub ParseChoiceExpr()
    Call ParseSequenceExpr
    Do
        If Not ParseChoiceLiteral() Then
            Exit Do
        End If
        Call ParseSequenceExpr
        pvPushAction ucsAct_1_ChoiceExpr
    Loop
End Sub

Public Sub ParseSequenceExpr()
    Dim p53 As Long
    Dim q53 As Long
    Dim p51 As Long
    Dim q51 As Long

    With ctx
        p53 = .BufPos
        q53 = .ThunkPos
        If ParseErrorExpr() Then
            Do
                p51 = .BufPos
                q51 = .ThunkPos
                If Not ParseErrorExpr() Then
                    .BufPos = p51
                    .ThunkPos = q51
                    Exit Do
                End If
                pvPushAction ucsAct_1_SequenceExpr
            Loop
            GoTo L8
        Else
            .BufPos = p53
            .ThunkPos = q53
            pvPushAction ucsAct_2_SequenceExpr
        End If
L8:
    End With
End Sub

Public Function ParseChoiceLiteral() As Boolean
    If ParseSLASH() Then
        ParseChoiceLiteral = True
        Exit Function
    End If
    If ParseBAR() Then
        ParseChoiceLiteral = True
    End If
End Function

Public Function ParseErrorExpr() As Boolean
    Dim p71 As Long
    Dim q71 As Long
    Dim p61 As Long
    Dim q61 As Long
    Dim p64 As Long
    Dim q64 As Long
    Dim p67 As Long
    Dim q67 As Long

    With ctx
        p71 = .BufPos
        q71 = .ThunkPos
        If ParseVariableExpr() Then
            p61 = .BufPos
            q61 = .ThunkPos
            If ParseCARET() Then
                If  ClearErrors()  Then
                    p64 = .BufPos
                    q64 = .ThunkPos
                    If ParseIdentifier() Then
                        pvPushAction ucsAct_1_ErrorExpr
                        GoTo L9
                    Else
                        .BufPos = p64
                        .ThunkPos = q64
                    End If
                    p67 = .BufPos
                    q67 = .ThunkPos
                    If ParseLiteral() Then
                        pvPushAction ucsAct_2_ErrorExpr
                        GoTo L9
                    Else
                        .BufPos = p67
                        .ThunkPos = q67
                    End If
                    If  AppendError("Expected rule or literal for error label")  Then
                        GoTo L9
                    End If
                    .BufPos = p61
                    .ThunkPos = q61
                Else
                    .BufPos = p61
                    .ThunkPos = q61
                End If
            End If
L9:
            ParseErrorExpr = True
            Exit Function
        Else
            .BufPos = p71
            .ThunkPos = q71
        End If
    End With
End Function

Public Function ParseVariableExpr() As Boolean
    Dim p81 As Long
    Dim q81 As Long

    With ctx
        p81 = .BufPos
        q81 = .ThunkPos
        If ParseIdentifier() Then
            pvPushAction ucsAct_1_VariableExpr
            If ParseCOLON() Then
                If ParsePrefixExpr() Then
                    pvPushAction ucsAct_2_VariableExpr
                    ParseVariableExpr = True
                    Exit Function
                Else
                    .BufPos = p81
                    .ThunkPos = q81
                End If
            Else
                .BufPos = p81
                .ThunkPos = q81
            End If
        Else
            .BufPos = p81
            .ThunkPos = q81
        End If
        If ParsePrefixExpr() Then
            ParseVariableExpr = True
            Exit Function
        Else
            .BufPos = p81
            .ThunkPos = q81
        End If
    End With
End Function

Private Function ParseCARET() As Boolean
    With ctx
        If .BufData(.BufPos) = 94 Then              ' "^"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseCARET = True
        End If
    End With
End Function

Private Function ParseCOLON() As Boolean
    With ctx
        If .BufData(.BufPos) = 58 Then              ' ":"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseCOLON = True
        End If
    End With
End Function

Public Function ParsePrefixExpr() As Boolean
    Dim p93 As Long
    Dim q93 As Long
    Dim p91 As Long
    Dim q91 As Long
    Dim p97 As Long
    Dim q97 As Long

    With ctx
        p93 = .BufPos
        q93 = .ThunkPos
        If ParseAND() Then
            If ParseCodeBlock() Then
                pvPushAction ucsAct_1_PrefixExpr
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p93
                .ThunkPos = q93
            End If
        End If
        p91 = .BufPos
        q91 = .ThunkPos
        If ParseAND() Then
            If ParseSuffixExpr() Then
                pvPushAction ucsAct_2_PrefixExpr
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p91
                .ThunkPos = q91
            End If
        End If
        p97 = .BufPos
        q97 = .ThunkPos
        If ParseNOT() Then
            If ParseSuffixExpr() Then
                pvPushAction ucsAct_3_PrefixExpr
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p97
                .ThunkPos = q97
            End If
        End If
        If ParseSuffixExpr() Then
            ParsePrefixExpr = True
            Exit Function
        Else
            .BufPos = p93
            .ThunkPos = q93
        End If
    End With
End Function

Private Function ParseAND() As Boolean
    With ctx
        If .BufData(.BufPos) = 38 Then              ' "&"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseAND = True
        End If
    End With
End Function

Public Function ParseCodeBlock() As Boolean
    Dim p383 As Long
    Dim q383 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        p383 = .BufPos
        q383 = .ThunkPos
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
                        GoTo L10
                    Else
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        GoTo L10
                    Else
                        If ParseCodeNC() Then
                            GoTo L10
                        Else
                            Exit Do
                        End If
                    End If
                End Select
L10:
            Loop
            lCaptureEnd = .BufPos
            If .BufData(.BufPos) = 125 Then         ' "}"
                .BufPos = .BufPos + 1
                Call ParseSpacing
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseCodeBlock = True
                Exit Function
            Else
                .BufPos = p383
                .ThunkPos = q383
            End If
        End If
    End With
End Function

Public Function ParseSuffixExpr() As Boolean
    Dim p116 As Long
    Dim q116 As Long

    With ctx
        p116 = .BufPos
        q116 = .ThunkPos
        If ParsePrimaryExpr() Then
            If ParseQUESTION() Then
                pvPushAction ucsAct_1_SuffixExpr
                GoTo L11
            End If
            If ParseSTAR() Then
                pvPushAction ucsAct_2_SuffixExpr
                GoTo L11
            End If
            If ParsePLUS() Then
                pvPushAction ucsAct_3_SuffixExpr
            End If
L11:
            ParseSuffixExpr = True
            Exit Function
        Else
            .BufPos = p116
            .ThunkPos = q116
        End If
    End With
End Function

Private Function ParseNOT() As Boolean
    With ctx
        If .BufData(.BufPos) = 33 Then              ' "!"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseNOT = True
        End If
    End With
End Function

Public Function ParsePrimaryExpr() As Boolean
    Dim p124 As Long
    Dim q124 As Long
    Dim p122 As Long
    Dim q122 As Long
    Dim p129 As Long
    Dim q129 As Long
    Dim p133 As Long
    Dim q133 As Long
    Dim p140 As Long
    Dim q140 As Long
    Dim p144 As Long
    Dim q144 As Long
    Dim p150 As Long
    Dim q150 As Long
    Dim p153 As Long
    Dim q153 As Long
    Dim p162 As Long
    Dim q162 As Long

    With ctx
        p124 = .BufPos
        q124 = .ThunkPos
        If ParseIdentifier() Then
            p122 = .BufPos
            q122 = .ThunkPos
            Call ParseLiteralNC
            If ParseAssignLiteral() Then
                .BufPos = p124
                .ThunkPos = q124
                GoTo L12
            Else
                .BufPos = p122
                .ThunkPos = q122
                pvPushAction ucsAct_1_PrimaryExpr
                ParsePrimaryExpr = True
                Exit Function
            End If
L12:
        Else
            .BufPos = p124
            .ThunkPos = q124
        End If
        p129 = .BufPos
        q129 = .ThunkPos
        If ParseOPEN() Then
            If  ClearErrors()  Then
                p133 = .BufPos
                q133 = .ThunkPos
                Call ParseChoiceExpr
                If ParseCLOSE() Then
                    ParsePrimaryExpr = True
                    Exit Function
                Else
                    .BufPos = p133
                    .ThunkPos = q133
                End If
                If  AppendError("Invalid expression")  Then
                    ParsePrimaryExpr = True
                    Exit Function
                End If
                .BufPos = p129
                .ThunkPos = q129
            Else
                .BufPos = p129
                .ThunkPos = q129
            End If
        End If
        p140 = .BufPos
        q140 = .ThunkPos
        If ParseLiteralNoSpace() Then
            pvPushAction ucsAct_2_PrimaryExpr
            p144 = .BufPos
            q144 = .ThunkPos
            If ParseIgnoreCase() Then
                pvPushAction ucsAct_3_PrimaryExpr
                GoTo L14
            Else
                .BufPos = p144
                .ThunkPos = q144
            End If
L14:
            Call ParseSpacing
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p140
            .ThunkPos = q140
        End If
        p150 = .BufPos
        q150 = .ThunkPos
        If ParseClassNoSpace() Then
            pvPushAction ucsAct_4_PrimaryExpr
            p153 = .BufPos
            q153 = .ThunkPos
            If ParseIgnoreCase() Then
                pvPushAction ucsAct_5_PrimaryExpr
                GoTo L15
            Else
                .BufPos = p153
                .ThunkPos = q153
            End If
L15:
            Call ParseSpacing
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p150
            .ThunkPos = q150
        End If
        If ParseDOT() Then
            pvPushAction ucsAct_6_PrimaryExpr
            ParsePrimaryExpr = True
            Exit Function
        End If
        p162 = .BufPos
        q162 = .ThunkPos
        If ParseCodeBlock() Then
            pvPushAction ucsAct_7_PrimaryExpr
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p162
            .ThunkPos = q162
        End If
        If ParseBEGIN() Then
            pvPushAction ucsAct_8_PrimaryExpr
            ParsePrimaryExpr = True
            Exit Function
        End If
        If ParseEND() Then
            pvPushAction ucsAct_9_PrimaryExpr
            ParsePrimaryExpr = True
        End If
    End With
End Function

Private Function ParseQUESTION() As Boolean
    With ctx
        If .BufData(.BufPos) = 63 Then              ' "?"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseQUESTION = True
        End If
    End With
End Function

Private Function ParseSTAR() As Boolean
    With ctx
        If .BufData(.BufPos) = 42 Then              ' "*"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseSTAR = True
        End If
    End With
End Function

Private Function ParsePLUS() As Boolean
    With ctx
        If .BufData(.BufPos) = 43 Then              ' "+"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParsePLUS = True
        End If
    End With
End Function

Public Function ParseLiteralNC() As Boolean
    Dim p244 As Long
    Dim p253 As Long

    With ctx
        p244 = .BufPos
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
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseLiteralNC = True
                Exit Function
            Else
                .BufPos = p244
            End If
        End If
        p253 = .BufPos
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
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseLiteralNC = True
                Exit Function
            Else
                .BufPos = p253
            End If
        End If
    End With
End Function

Private Function ParseOPEN() As Boolean
    With ctx
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseOPEN = True
        End If
    End With
End Function

Private Function ParseCLOSE() As Boolean
    With ctx
        If .BufData(.BufPos) = 41 Then              ' ")"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseCLOSE = True
        End If
    End With
End Function

Public Function ParseLiteralNoSpace() As Boolean
    Dim p216 As Long
    Dim q216 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long
    Dim p228 As Long
    Dim q228 As Long

    With ctx
        p216 = .BufPos
        q216 = .ThunkPos
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
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseLiteralNoSpace = True
                Exit Function
            Else
                .BufPos = p216
                .ThunkPos = q216
            End If
        End If
        p228 = .BufPos
        q228 = .ThunkPos
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
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseLiteralNoSpace = True
                Exit Function
            Else
                .BufPos = p228
                .ThunkPos = q228
            End If
        End If
    End With
End Function

Public Function ParseIgnoreCase() As Boolean
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        Select Case .BufData(.BufPos)
        Case 115, 105                               ' [si]
            .BufPos = .BufPos + 1
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseIgnoreCase = True
        End Select
    End With
End Function

Public Function ParseClassNoSpace() As Boolean
    Dim p264 As Long
    Dim q264 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        p264 = .BufPos
        q264 = .ThunkPos
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
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseClassNoSpace = True
                Exit Function
            Else
                .BufPos = p264
                .ThunkPos = q264
            End If
        End If
    End With
End Function

Private Function ParseDOT() As Boolean
    With ctx
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseDOT = True
        End If
    End With
End Function

Private Function ParseBEGIN() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseBEGIN = True
        End If
    End With
End Function

Private Function ParseEND() As Boolean
    With ctx
        If .BufData(.BufPos) = 62 Then              ' ">"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseEND = True
        End If
    End With
End Function

Public Function ParseSettingName() As Boolean
    Dim p186 As Long
    Dim q186 As Long

    With ctx
        If ParseAT() Then
            p186 = .BufPos
            q186 = .ThunkPos
            If ParseIdentifier() Then
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p186
                .ThunkPos = q186
            End If
            If ParseLiteral() Then
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p186
                .ThunkPos = q186
            End If
            Call ParseEMPTY
            ParseSettingName = True
        End If
    End With
End Function

Public Function ParseSettingValue() As Boolean
    Dim p192 As Long
    Dim q192 As Long

    With ctx
        p192 = .BufPos
        q192 = .ThunkPos
        If ParseIdentifier() Then
            Call ParseSEMI
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p192
            .ThunkPos = q192
        End If
        If ParseLiteral() Then
            Call ParseSEMI
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p192
            .ThunkPos = q192
        End If
        If ParseCodeBlock() Then
            Call ParseSEMI
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p192
            .ThunkPos = q192
        End If
    End With
End Function

Private Function ParseAT() As Boolean
    With ctx
        If .BufData(.BufPos) = 64 Then              ' "@"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseAT = True
        End If
    End With
End Function

Private Sub ParseEMPTY()
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        lCaptureEnd = .BufPos
        Call ParseSpacing
        .CaptureBegin = lCaptureBegin
        .CaptureEnd = lCaptureEnd
    End With
End Sub

Public Function ParseChar() As Boolean
    Dim p284 As Long
    Dim p287 As Long
    Dim p293 As Long
    Dim p298 As Long
    Dim p302 As Long

    With ctx
        p284 = .BufPos
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 97, 98, 101, 102, 110, 114, 116, 118, 39, 34, 91, 93, 92 ' [abefnrtv'"\[\]\\]
                .BufPos = .BufPos + 1
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseChar = True
                Exit Function
            Case Else
                .BufPos = p284
            End Select
        End If
        p287 = .BufPos
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
                        If .BufPos > .LastBufPos Then
                            .LastError = vbNullString : .LastBufPos = .BufPos
                        End If
                        ParseChar = True
                        Exit Function
                    Case Else
                        .BufPos = p287
                    End Select
                Case Else
                    .BufPos = p287
                End Select
            Case Else
                .BufPos = p287
            End Select
        End If
        p293 = .BufPos
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 48 To 55                           ' [0-7]
                .BufPos = .BufPos + 1
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                End Select
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseChar = True
                Exit Function
            Case Else
                .BufPos = p293
            End Select
        End If
        p298 = .BufPos
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 120 Then ' "\\x"
            .BufPos = .BufPos + 2
            Select Case .BufData(.BufPos)
            Case 48 To 57, 97 To 102, 65 To 70      ' [0-9a-fA-F]
                .BufPos = .BufPos + 1
                Select Case .BufData(.BufPos)
                Case 48 To 57, 97 To 102, 65 To 70  ' [0-9a-fA-F]
                    .BufPos = .BufPos + 1
                    If .BufPos > .LastBufPos Then
                        .LastError = vbNullString : .LastBufPos = .BufPos
                    End If
                    ParseChar = True
                    Exit Function
                Case Else
                    .BufPos = p298
                End Select
            Case Else
                .BufPos = p298
            End Select
        End If
        p302 = .BufPos
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
                            If .BufPos > .LastBufPos Then
                                .LastError = vbNullString : .LastBufPos = .BufPos
                            End If
                            ParseChar = True
                            Exit Function
                        Case Else
                            .BufPos = p302
                        End Select
                    Case Else
                        .BufPos = p302
                    End Select
                Case Else
                    .BufPos = p302
                End Select
            Case Else
                .BufPos = p302
            End Select
        End If
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 45 Then ' "\\-"
            .BufPos = .BufPos + 2
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseChar = True
            Exit Function
        End If
        If Not .BufData(.BufPos) = 92 Then          ' "\\"
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseChar = True
                Exit Function
            End If
        End If
        If LenB(.LastError) = 0 Or .BufPos > .LastBufPos Then
            .LastError = "Expected char" : .LastBufPos = .BufPos
        End If
    End With
End Function

Public Function ParseRange() As Boolean
    Dim p276 As Long

    With ctx
        p276 = .BufPos
        If ParseChar() Then
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = p276
                Else
                    If ParseChar() Then
                        If .BufPos > .LastBufPos Then
                            .LastError = vbNullString : .LastBufPos = .BufPos
                        End If
                        ParseRange = True
                        Exit Function
                    Else
                        .BufPos = p276
                    End If
                End If
            Else
                .BufPos = p276
            End If
        End If
        If ParseChar() Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseRange = True
        End If
    End With
End Function

Private Function ParseLEFTARROW() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 And .BufData(.BufPos + 1) = 45 Then ' "<-"
            .BufPos = .BufPos + 2
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseLEFTARROW = True
        End If
    End With
End Function

Private Function ParseSLASH() As Boolean
    Dim p321 As Long

    With ctx
        p321 = .BufPos
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = p321
            Else
                Call ParseSpacing
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseSLASH = True
            End If
        End If
    End With
End Function

Public Function ParseSpace() As Boolean
    With ctx
        If .BufData(.BufPos) = 32 Then              ' " "
            .BufPos = .BufPos + 1
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseSpace = True
            Exit Function
        End If
        If .BufData(.BufPos) = 9 Then               ' "\t"
            .BufPos = .BufPos + 1
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseSpace = True
            Exit Function
        End If
        If ParseEndOfLine() Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseSpace = True
        End If
    End With
End Function

Public Function ParseComment() As Boolean
    Dim p368 As Long
    Dim p366 As Long

    With ctx
        p368 = .BufPos
        If .BufData(.BufPos) = 35 Then              ' "#"
            .BufPos = .BufPos + 1
            Do
                p366 = .BufPos
                If ParseEndOfLine() Then
                    .BufPos = p366
                    Exit Do
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p366
                    Exit Do
                End If
            Loop
            If ParseEndOfLine() Then
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseComment = True
                Exit Function
            Else
                .BufPos = p368
            End If
        End If
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 47 Then ' "//"
            .BufPos = .BufPos + 2
            Do
                p366 = .BufPos
                If ParseEndOfLine() Then
                    .BufPos = p366
                    Exit Do
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p366
                    Exit Do
                End If
            Loop
            If ParseEndOfLine() Then
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseComment = True
                Exit Function
            Else
                .BufPos = p368
            End If
        End If
    End With
End Function

Public Function ParseEndOfLine() As Boolean
    With ctx
        If .BufData(.BufPos) = 13 And .BufData(.BufPos + 1) = 10 Then ' "\r\n"
            .BufPos = .BufPos + 2
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseEndOfLine = True
            Exit Function
        End If
        If .BufData(.BufPos) = 10 Then              ' "\n"
            .BufPos = .BufPos + 1
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseEndOfLine = True
            Exit Function
        End If
        If .BufData(.BufPos) = 13 Then              ' "\r"
            .BufPos = .BufPos + 1
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseEndOfLine = True
        End If
    End With
End Function

Public Function ParseCodeNC() As Boolean
    Dim p399 As Long

    With ctx
        p399 = .BufPos
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Do
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If ParseCodeNC() Then
                        GoTo L23
                    Else
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        GoTo L23
                    Else
                        If ParseCodeNC() Then
                            GoTo L23
                        Else
                            Exit Do
                        End If
                    End If
                End Select
L23:
            Loop
            If .BufData(.BufPos) = 125 Then         ' "}"
                .BufPos = .BufPos + 1
                Call ParseSpacing
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseCodeNC = True
                Exit Function
            Else
                .BufPos = p399
            End If
        End If
    End With
End Function

Private Function ParseEQUAL() As Boolean
    With ctx
        If .BufData(.BufPos) = 61 Then              ' "="
            .BufPos = .BufPos + 1
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseEQUAL = True
        End If
    End With
End Function

Private Function ParseBAR() As Boolean
    With ctx
        If .BufData(.BufPos) = 124 Then             ' "|"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
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
         With .UserData: .AppendSetting .Pop, Mid$(ctx.Contents, lOffset, lSize): End With
    Case ucsAct_1_Setting
         With .UserData: .Push .MakeString(lOffset, Mid$(ctx.Contents, lOffset, lSize)): End With
    Case ucsAct_3_Rule
         With .UserData: lNode = .Pop(): .RuleExpression(.Pop()) = lNode: End With
    Case ucsAct_2_Rule
         With .UserData: .RuleDescription(.Top) = Mid$(ctx.Contents, lOffset, lSize): End With
    Case ucsAct_1_Rule
           With .UserData
                                                                                lNode = .BeginRule(.FindRule(lOffset, Mid$(ctx.Contents, lOffset, lSize)))
                                                                                If .RuleExpression(lNode) <> 0 Then
                                                                                    ConsoleError "Rule '%1' redefined" & vbCrLf, Mid$(ctx.Contents, lOffset, lSize)
                                                                                End If
                                                                                .Push lNode
                                                                            End With
    Case ucsAct_1_ChoiceExpr
         With .UserData: lNode = .Pop(): .Push .AppendChild(.MakeAlternate(lOffset, .Pop), lNode): End With
    Case ucsAct_2_SequenceExpr
         With .UserData: .Push .MakePredicate(lOffset, "True"): End With
    Case ucsAct_1_SequenceExpr
         With .UserData: lNode = .Pop(): .Push .AppendSequence(.MakeSequence(lOffset, .Pop), lNode): End With
    Case ucsAct_1_ErrorExpr
         With .UserData: .Push .AppendChild(.MakeAlternate(lOffset, .Pop), .MakeErrorReference(lOffset, Mid$(ctx.Contents, lOffset, lSize))): End With
    Case ucsAct_2_ErrorExpr
         With .UserData: .Push .AppendChild(.MakeAlternate(lOffset, .Pop), .MakeErrorString(lOffset, Mid$(ctx.Contents, lOffset, lSize))): End With
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
         With .UserData: .Push .MakePredicate(lOffset, "YY_BEGIN"): End With
    Case ucsAct_9_PrimaryExpr
         With .UserData: .Push .MakePredicate(lOffset, "YY_END"): End With
    End Select
    End With
End Sub

'= user-defined functions ================================================

Private Function ClearErrors() As Boolean
    Dim lIdx            As Long
    
    If Not ctx.Errors Is Nothing Then
        For lIdx = ctx.Errors.Count To 1 Step -1
            If ctx.Errors.Item(lIdx)(1) <= ctx.BufPos Then
                ctx.Errors.Remove lIdx
            End If
        Next
        If ctx.Errors.Count = 0 Then
            Set ctx.Errors = Nothing
        End If
    End If
    ClearErrors = True
End Function

Private Function AppendError(sText As String) As Boolean
    If ctx.Errors Is Nothing Then
        Set ctx.Errors = New Collection
    End If
    ctx.Errors.Add Array(sText, ctx.BufPos)
End Function

Private Function SetUserDataTag(sText As String) As Boolean
    ctx.UserData.Tag = sText
    SetUserDataTag = True
End Function

Public Function GetParseErrors() As Collection
    Set GetParseErrors = ctx.Errors
End Function
