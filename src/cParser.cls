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
' Auto-generated on 17.7.2018 13:24:06
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
    ParserVersion = "17.7.2018 13:24:06"
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
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseGrammar = True
                Exit Function
            End If
            If  AppendError("Expected rule or setting")  Then
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
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
    Dim p177 As Long
    Dim q177 As Long
    Dim p182 As Long
    Dim q182 As Long

    With ctx
        p177 = .BufPos
        q177 = .ThunkPos
        If ParseSettingName() Then
            pvPushAction ucsAct_1_Setting
            If  ClearErrors() And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin))  Then
                p182 = .BufPos
                q182 = .ThunkPos
                If ParseSettingValue() Then
                    pvPushAction ucsAct_2_Setting
                    If .BufPos > .LastBufPos Then
                        .LastError = vbNullString : .LastBufPos = .BufPos
                    End If
                    ParseSetting = True
                    Exit Function
                Else
                    .BufPos = p182
                    .ThunkPos = q182
                End If
                If  AppendError("Invalid value for setting '" & .UserData.Tag & "'")  Then
                    If .BufPos > .LastBufPos Then
                        .LastError = vbNullString : .LastBufPos = .BufPos
                    End If
                    ParseSetting = True
                    Exit Function
                End If
                .BufPos = p177
                .ThunkPos = q177
            Else
                .BufPos = p177
                .ThunkPos = q177
            End If
        Else
            .BufPos = p177
            .ThunkPos = q177
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
                If ParseAssign() Then
                    Call ParseChoiceExpr
                    pvPushAction ucsAct_3_Rule
                    Call ParseSEMI
                    If .BufPos > .LastBufPos Then
                        .LastError = vbNullString : .LastBufPos = .BufPos
                    End If
                    ParseRule = True
                    Exit Function
                    If  AppendError("Invalid definition for rule '" & .UserData.Tag & "'")  Then
                        Call ParseSEMI
                        If .BufPos > .LastBufPos Then
                            .LastError = vbNullString : .LastBufPos = .BufPos
                        End If
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
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
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
    Dim p216 As Long
    Dim q216 As Long

    With ctx
        p216 = .BufPos
        q216 = .ThunkPos
        If ParseLiteralNoSpace() Then
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseLiteral = True
            Exit Function
        Else
            .BufPos = p216
            .ThunkPos = q216
        End If
    End With
End Function

Public Function ParseAssign() As Boolean
    With ctx
        If ParseLEFTARROW() Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseAssign = True
            Exit Function
        End If
        If ParseEQUAL() Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseAssign = True
        End If
    End With
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
        If Not ParseChoice() Then
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

Public Function ParseChoice() As Boolean
    With ctx
        If ParseSLASH() Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseChoice = True
            Exit Function
        End If
        If ParseBAR() Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseChoice = True
        End If
    End With
End Function

Public Function ParseErrorExpr() As Boolean
    Dim p60 As Long
    Dim q60 As Long
    Dim p64 As Long
    Dim q64 As Long
    Dim p67 As Long
    Dim q67 As Long
    Dim p73 As Long
    Dim q73 As Long

    With ctx
        p60 = .BufPos
        q60 = .ThunkPos
        If ParseVariableExpr() Then
            If ParseCARET() Then
                If  ClearErrors()  Then
                    p64 = .BufPos
                    q64 = .ThunkPos
                    If ParseIdentifier() Then
                        pvPushAction ucsAct_1_ErrorExpr
                        If .BufPos > .LastBufPos Then
                            .LastError = vbNullString : .LastBufPos = .BufPos
                        End If
                        ParseErrorExpr = True
                        Exit Function
                    Else
                        .BufPos = p64
                        .ThunkPos = q64
                    End If
                    p67 = .BufPos
                    q67 = .ThunkPos
                    If ParseLiteral() Then
                        pvPushAction ucsAct_2_ErrorExpr
                        If .BufPos > .LastBufPos Then
                            .LastError = vbNullString : .LastBufPos = .BufPos
                        End If
                        ParseErrorExpr = True
                        Exit Function
                    Else
                        .BufPos = p67
                        .ThunkPos = q67
                    End If
                    If  AppendError("Expected rule or literal for error label")  Then
                        If .BufPos > .LastBufPos Then
                            .LastError = vbNullString : .LastBufPos = .BufPos
                        End If
                        ParseErrorExpr = True
                        Exit Function
                    End If
                    .BufPos = p60
                    .ThunkPos = q60
                Else
                    .BufPos = p60
                    .ThunkPos = q60
                End If
            Else
                .BufPos = p60
                .ThunkPos = q60
            End If
        Else
            .BufPos = p60
            .ThunkPos = q60
        End If
        p73 = .BufPos
        q73 = .ThunkPos
        If ParseVariableExpr() Then
            If ParseCARET() Then
                .BufPos = p73
                .ThunkPos = q73
            Else
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseErrorExpr = True
                Exit Function
            End If
        Else
            .BufPos = p73
            .ThunkPos = q73
        End If
    End With
End Function

Public Function ParseVariableExpr() As Boolean
    Dim p84 As Long
    Dim q84 As Long

    With ctx
        p84 = .BufPos
        q84 = .ThunkPos
        If ParseIdentifier() Then
            pvPushAction ucsAct_1_VariableExpr
            If ParseCOLON() Then
                If ParsePrefixExpr() Then
                    pvPushAction ucsAct_2_VariableExpr
                    If .BufPos > .LastBufPos Then
                        .LastError = vbNullString : .LastBufPos = .BufPos
                    End If
                    ParseVariableExpr = True
                    Exit Function
                Else
                    .BufPos = p84
                    .ThunkPos = q84
                End If
            Else
                .BufPos = p84
                .ThunkPos = q84
            End If
        Else
            .BufPos = p84
            .ThunkPos = q84
        End If
        If ParsePrefixExpr() Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseVariableExpr = True
            Exit Function
        Else
            .BufPos = p84
            .ThunkPos = q84
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
    Dim p96 As Long
    Dim q96 As Long
    Dim p94 As Long
    Dim q94 As Long
    Dim p100 As Long
    Dim q100 As Long

    With ctx
        p96 = .BufPos
        q96 = .ThunkPos
        If ParseAND() Then
            If ParseAction() Then
                pvPushAction ucsAct_1_PrefixExpr
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p96
                .ThunkPos = q96
            End If
        End If
        p94 = .BufPos
        q94 = .ThunkPos
        If ParseAND() Then
            If ParseSuffixExpr() Then
                pvPushAction ucsAct_2_PrefixExpr
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p94
                .ThunkPos = q94
            End If
        End If
        p100 = .BufPos
        q100 = .ThunkPos
        If ParseNOT() Then
            If ParseSuffixExpr() Then
                pvPushAction ucsAct_3_PrefixExpr
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParsePrefixExpr = True
                Exit Function
            Else
                .BufPos = p100
                .ThunkPos = q100
            End If
        End If
        If ParseSuffixExpr() Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParsePrefixExpr = True
            Exit Function
        Else
            .BufPos = p96
            .ThunkPos = q96
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

Public Function ParseAction() As Boolean
    Dim p376 As Long
    Dim q376 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        p376 = .BufPos
        q376 = .ThunkPos
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
                    If ParseActionNC() Then
                        GoTo L9
                    Else
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        GoTo L9
                    Else
                        If ParseActionNC() Then
                            GoTo L9
                        Else
                            Exit Do
                        End If
                    End If
                End Select
L9:
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
                ParseAction = True
                Exit Function
            Else
                .BufPos = p376
                .ThunkPos = q376
            End If
        End If
    End With
End Function

Public Function ParseSuffixExpr() As Boolean
    Dim p119 As Long
    Dim q119 As Long

    With ctx
        p119 = .BufPos
        q119 = .ThunkPos
        If ParsePrimaryExpr() Then
            If ParseQUESTION() Then
                pvPushAction ucsAct_1_SuffixExpr
                GoTo L10
            End If
            If ParseSTAR() Then
                pvPushAction ucsAct_2_SuffixExpr
                GoTo L10
            End If
            If ParsePLUS() Then
                pvPushAction ucsAct_3_SuffixExpr
            End If
L10:
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseSuffixExpr = True
            Exit Function
        Else
            .BufPos = p119
            .ThunkPos = q119
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
    Dim p127 As Long
    Dim q127 As Long
    Dim p125 As Long
    Dim q125 As Long
    Dim p132 As Long
    Dim q132 As Long
    Dim p136 As Long
    Dim q136 As Long
    Dim p143 As Long
    Dim q143 As Long
    Dim p147 As Long
    Dim q147 As Long
    Dim p153 As Long
    Dim q153 As Long
    Dim p156 As Long
    Dim q156 As Long
    Dim p165 As Long
    Dim q165 As Long

    With ctx
        p127 = .BufPos
        q127 = .ThunkPos
        If ParseIdentifier() Then
            p125 = .BufPos
            q125 = .ThunkPos
            Call ParseLiteralNC
            If ParseAssign() Then
                .BufPos = p127
                .ThunkPos = q127
                GoTo L11
            Else
                .BufPos = p125
                .ThunkPos = q125
                pvPushAction ucsAct_1_PrimaryExpr
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParsePrimaryExpr = True
                Exit Function
            End If
L11:
        Else
            .BufPos = p127
            .ThunkPos = q127
        End If
        p132 = .BufPos
        q132 = .ThunkPos
        If ParseOPEN() Then
            If  ClearErrors()  Then
                p136 = .BufPos
                q136 = .ThunkPos
                Call ParseChoiceExpr
                If ParseCLOSE() Then
                    If .BufPos > .LastBufPos Then
                        .LastError = vbNullString : .LastBufPos = .BufPos
                    End If
                    ParsePrimaryExpr = True
                    Exit Function
                Else
                    .BufPos = p136
                    .ThunkPos = q136
                End If
                If  AppendError("Invalid expression")  Then
                    If .BufPos > .LastBufPos Then
                        .LastError = vbNullString : .LastBufPos = .BufPos
                    End If
                    ParsePrimaryExpr = True
                    Exit Function
                End If
                .BufPos = p132
                .ThunkPos = q132
            Else
                .BufPos = p132
                .ThunkPos = q132
            End If
        End If
        p143 = .BufPos
        q143 = .ThunkPos
        If ParseLiteralNoSpace() Then
            pvPushAction ucsAct_2_PrimaryExpr
            p147 = .BufPos
            q147 = .ThunkPos
            If ParseIgnoreCase() Then
                pvPushAction ucsAct_3_PrimaryExpr
                GoTo L13
            Else
                .BufPos = p147
                .ThunkPos = q147
            End If
L13:
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p143
            .ThunkPos = q143
        End If
        p153 = .BufPos
        q153 = .ThunkPos
        If ParseClassNoSpace() Then
            pvPushAction ucsAct_4_PrimaryExpr
            p156 = .BufPos
            q156 = .ThunkPos
            If ParseIgnoreCase() Then
                pvPushAction ucsAct_5_PrimaryExpr
                GoTo L14
            Else
                .BufPos = p156
                .ThunkPos = q156
            End If
L14:
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p153
            .ThunkPos = q153
        End If
        If ParseDOT() Then
            pvPushAction ucsAct_6_PrimaryExpr
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParsePrimaryExpr = True
            Exit Function
        End If
        p165 = .BufPos
        q165 = .ThunkPos
        If ParseAction() Then
            pvPushAction ucsAct_7_PrimaryExpr
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p165
            .ThunkPos = q165
        End If
        If ParseBEGIN() Then
            pvPushAction ucsAct_8_PrimaryExpr
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParsePrimaryExpr = True
            Exit Function
        End If
        If ParseEND() Then
            pvPushAction ucsAct_9_PrimaryExpr
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
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
    Dim p247 As Long
    Dim p256 As Long

    With ctx
        p247 = .BufPos
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
                .BufPos = p247
            End If
        End If
        p256 = .BufPos
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
                .BufPos = p256
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
    Dim p219 As Long
    Dim q219 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long
    Dim p231 As Long
    Dim q231 As Long

    With ctx
        p219 = .BufPos
        q219 = .ThunkPos
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
                .BufPos = p219
                .ThunkPos = q219
            End If
        End If
        p231 = .BufPos
        q231 = .ThunkPos
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
                .BufPos = p231
                .ThunkPos = q231
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
    Dim p267 As Long
    Dim q267 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        p267 = .BufPos
        q267 = .ThunkPos
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
                .BufPos = p267
                .ThunkPos = q267
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
    Dim p189 As Long
    Dim q189 As Long

    With ctx
        If ParseAT() Then
            p189 = .BufPos
            q189 = .ThunkPos
            If ParseIdentifier() Then
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p189
                .ThunkPos = q189
            End If
            If ParseLiteral() Then
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p189
                .ThunkPos = q189
            End If
            Call ParseEMPTY
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseSettingName = True
        End If
    End With
End Function

Public Function ParseSettingValue() As Boolean
    Dim p195 As Long
    Dim q195 As Long

    With ctx
        p195 = .BufPos
        q195 = .ThunkPos
        If ParseIdentifier() Then
            Call ParseSEMI
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p195
            .ThunkPos = q195
        End If
        If ParseLiteral() Then
            Call ParseSEMI
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p195
            .ThunkPos = q195
        End If
        If ParseAction() Then
            Call ParseSEMI
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p195
            .ThunkPos = q195
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
    Dim p287 As Long
    Dim p290 As Long
    Dim p296 As Long

    With ctx
        p287 = .BufPos
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
                .BufPos = p287
            End Select
        End If
        p290 = .BufPos
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
                        .BufPos = p290
                    End Select
                Case Else
                    .BufPos = p290
                End Select
            Case Else
                .BufPos = p290
            End Select
        End If
        p296 = .BufPos
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
                .BufPos = p296
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
            End If
        End If
    End With
End Function

Public Function ParseRange() As Boolean
    Dim p279 As Long

    With ctx
        p279 = .BufPos
        If ParseChar() Then
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = p279
                Else
                    If ParseChar() Then
                        If .BufPos > .LastBufPos Then
                            .LastError = vbNullString : .LastBufPos = .BufPos
                        End If
                        ParseRange = True
                        Exit Function
                    Else
                        .BufPos = p279
                    End If
                End If
            Else
                .BufPos = p279
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
    Dim p314 As Long

    With ctx
        p314 = .BufPos
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = p314
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
    Dim p361 As Long
    Dim p359 As Long

    With ctx
        p361 = .BufPos
        If .BufData(.BufPos) = 35 Then              ' "#"
            .BufPos = .BufPos + 1
            Do
                p359 = .BufPos
                If ParseEndOfLine() Then
                    .BufPos = p359
                    Exit Do
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p359
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
                .BufPos = p361
            End If
        End If
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 47 Then ' "//"
            .BufPos = .BufPos + 2
            Do
                p359 = .BufPos
                If ParseEndOfLine() Then
                    .BufPos = p359
                    Exit Do
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p359
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
                .BufPos = p361
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

Public Function ParseActionNC() As Boolean
    Dim p392 As Long

    With ctx
        p392 = .BufPos
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Do
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If ParseActionNC() Then
                        GoTo L22
                    Else
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        GoTo L22
                    Else
                        If ParseActionNC() Then
                            GoTo L22
                        Else
                            Exit Do
                        End If
                    End If
                End Select
L22:
            Loop
            If .BufData(.BufPos) = 125 Then         ' "}"
                .BufPos = .BufPos + 1
                Call ParseSpacing
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseActionNC = True
                Exit Function
            Else
                .BufPos = p392
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
         With .UserData: .Push .MakePredicate(lOffset, "1"): End With
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
