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
' Auto-generated on 16.7.2018 18:51:05
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
    ucsAct_3_Definition
    ucsAct_2_Definition
    ucsAct_1_Definition
    ucsAct_1_Expression
    ucsAct_2_Sequence
    ucsAct_1_Sequence
    ucsAct_1_Error
    ucsAct_2_Error
    ucsAct_1_Prefix
    ucsAct_2_Prefix
    ucsAct_3_Prefix
    ucsAct_2_Variable
    ucsAct_1_Variable
    ucsAct_1_Suffix
    ucsAct_2_Suffix
    ucsAct_3_Suffix
    ucsAct_1_Primary
    ucsAct_3_Primary
    ucsAct_2_Primary
    ucsAct_5_Primary
    ucsAct_4_Primary
    ucsAct_6_Primary
    ucsAct_7_Primary
    ucsAct_8_Primary
    ucsAct_9_Primary
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
    ParserVersion = "16.7.2018 18:51:05"
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
                If ParseDefinition() Then
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
    Dim p175 As Long
    Dim q175 As Long

    With ctx
        p175 = .BufPos
        q175 = .ThunkPos
        If ParseSettingName() Then
            pvPushAction ucsAct_1_Setting
            If  ClearErrors() And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin))  Then
                If ParseSettingValue() Then
                    pvPushAction ucsAct_2_Setting
                    If .BufPos > .LastBufPos Then
                        .LastError = vbNullString : .LastBufPos = .BufPos
                    End If
                    ParseSetting = True
                    Exit Function
                End If
                If  AppendError("Invalid value for setting '" & .UserData.Tag & "'")  Then
                    If .BufPos > .LastBufPos Then
                        .LastError = vbNullString : .LastBufPos = .BufPos
                    End If
                    ParseSetting = True
                    Exit Function
                End If
                .BufPos = p175
                .ThunkPos = q175
            Else
                .BufPos = p175
                .ThunkPos = q175
            End If
        End If
    End With
End Function

Public Function ParseDefinition() As Boolean
    Dim p18 As Long
    Dim q18 As Long

    With ctx
        p18 = .BufPos
        q18 = .ThunkPos
        If ParseIdentifier() Then
            pvPushAction ucsAct_1_Definition
            If  ClearErrors() And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin))  Then
                If ParseLiteral() Then
                    pvPushAction ucsAct_2_Definition
                End If
                If ParseAssign() Then
                    Call ParseExpression
                    pvPushAction ucsAct_3_Definition
                    Call ParseSEMI
                    If .BufPos > .LastBufPos Then
                        .LastError = vbNullString : .LastBufPos = .BufPos
                    End If
                    ParseDefinition = True
                    Exit Function
                    If  AppendError("Invalid definition for rule '" & .UserData.Tag & "'")  Then
                        Call ParseSEMI
                        If .BufPos > .LastBufPos Then
                            .LastError = vbNullString : .LastBufPos = .BufPos
                        End If
                        ParseDefinition = True
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
    With ctx
        If ParseLiteralNoSpace() Then
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseLiteral = True
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

Public Sub ParseExpression()
    Call ParseSequence
    Do
        If Not ParseChoice() Then
            Exit Do
        End If
        Call ParseSequence
        pvPushAction ucsAct_1_Expression
    Loop
End Sub

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

Public Sub ParseSequence()
    If ParseError() Then
        Do
            If Not ParseError() Then
                Exit Do
            End If
            pvPushAction ucsAct_1_Sequence
        Loop
        GoTo L8
    Else
        pvPushAction ucsAct_2_Sequence
    End If
L8:
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

Public Function ParseError() As Boolean
    Dim p58 As Long
    Dim q58 As Long
    Dim p71 As Long
    Dim q71 As Long

    With ctx
        p58 = .BufPos
        q58 = .ThunkPos
        If ParsePrefix() Then
            If ParseCARET() Then
                If  ClearErrors()  Then
                    If ParseIdentifier() Then
                        pvPushAction ucsAct_1_Error
                        If .BufPos > .LastBufPos Then
                            .LastError = vbNullString : .LastBufPos = .BufPos
                        End If
                        ParseError = True
                        Exit Function
                    End If
                    If ParseLiteral() Then
                        pvPushAction ucsAct_2_Error
                        If .BufPos > .LastBufPos Then
                            .LastError = vbNullString : .LastBufPos = .BufPos
                        End If
                        ParseError = True
                        Exit Function
                    End If
                    If  AppendError("Expected rule or literal for error label")  Then
                        If .BufPos > .LastBufPos Then
                            .LastError = vbNullString : .LastBufPos = .BufPos
                        End If
                        ParseError = True
                        Exit Function
                    End If
                    .BufPos = p58
                    .ThunkPos = q58
                Else
                    .BufPos = p58
                    .ThunkPos = q58
                End If
            Else
                .BufPos = p58
                .ThunkPos = q58
            End If
        End If
        p71 = .BufPos
        q71 = .ThunkPos
        If ParsePrefix() Then
            If ParseCARET() Then
                .BufPos = p71
                .ThunkPos = q71
            Else
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseError = True
            End If
        End If
    End With
End Function

Public Function ParsePrefix() As Boolean
    Dim p84 As Long
    Dim q84 As Long
    Dim p82 As Long
    Dim q82 As Long
    Dim p88 As Long
    Dim q88 As Long

    With ctx
        p84 = .BufPos
        q84 = .ThunkPos
        If ParseAND() Then
            If ParseAction() Then
                pvPushAction ucsAct_1_Prefix
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParsePrefix = True
                Exit Function
            Else
                .BufPos = p84
                .ThunkPos = q84
            End If
        End If
        p82 = .BufPos
        q82 = .ThunkPos
        If ParseAND() Then
            If ParseVariable() Then
                pvPushAction ucsAct_2_Prefix
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParsePrefix = True
                Exit Function
            Else
                .BufPos = p82
                .ThunkPos = q82
            End If
        End If
        p88 = .BufPos
        q88 = .ThunkPos
        If ParseNOT() Then
            If ParseVariable() Then
                pvPushAction ucsAct_3_Prefix
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParsePrefix = True
                Exit Function
            Else
                .BufPos = p88
                .ThunkPos = q88
            End If
        End If
        If ParseVariable() Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParsePrefix = True
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
    Dim p374 As Long
    Dim q374 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        p374 = .BufPos
        q374 = .ThunkPos
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
                .BufPos = p374
                .ThunkPos = q374
            End If
        End If
    End With
End Function

Public Function ParseVariable() As Boolean
    Dim p100 As Long
    Dim q100 As Long

    With ctx
        p100 = .BufPos
        q100 = .ThunkPos
        If ParseIdentifier() Then
            pvPushAction ucsAct_1_Variable
            If ParseCOLON() Then
                If ParseSuffix() Then
                    pvPushAction ucsAct_2_Variable
                    If .BufPos > .LastBufPos Then
                        .LastError = vbNullString : .LastBufPos = .BufPos
                    End If
                    ParseVariable = True
                    Exit Function
                Else
                    .BufPos = p100
                    .ThunkPos = q100
                End If
            Else
                .BufPos = p100
                .ThunkPos = q100
            End If
        End If
        If ParseSuffix() Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseVariable = True
            Exit Function
        Else
            .BufPos = p100
            .ThunkPos = q100
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

Public Function ParseSuffix() As Boolean
    With ctx
        If ParsePrimary() Then
            If ParseQUESTION() Then
                pvPushAction ucsAct_1_Suffix
                GoTo L10
            End If
            If ParseSTAR() Then
                pvPushAction ucsAct_2_Suffix
                GoTo L10
            End If
            If ParsePLUS() Then
                pvPushAction ucsAct_3_Suffix
            End If
L10:
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseSuffix = True
        End If
    End With
End Function

Public Function ParsePrimary() As Boolean
    Dim p125 As Long
    Dim q125 As Long
    Dim p123 As Long
    Dim q123 As Long
    Dim p130 As Long
    Dim q130 As Long
    Dim p134 As Long
    Dim q134 As Long

    With ctx
        p125 = .BufPos
        q125 = .ThunkPos
        If ParseIdentifier() Then
            p123 = .BufPos
            q123 = .ThunkPos
            Call ParseLiteralNC
            If ParseAssign() Then
                .BufPos = p125
                .ThunkPos = q125
                GoTo L11
            Else
                .BufPos = p123
                .ThunkPos = q123
                pvPushAction ucsAct_1_Primary
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParsePrimary = True
                Exit Function
            End If
L11:
        End If
        p130 = .BufPos
        q130 = .ThunkPos
        If ParseOPEN() Then
            If  ClearErrors()  Then
                p134 = .BufPos
                q134 = .ThunkPos
                Call ParseExpression
                If ParseCLOSE() Then
                    If .BufPos > .LastBufPos Then
                        .LastError = vbNullString : .LastBufPos = .BufPos
                    End If
                    ParsePrimary = True
                    Exit Function
                Else
                    .BufPos = p134
                    .ThunkPos = q134
                End If
                If  AppendError("Invalid expression")  Then
                    If .BufPos > .LastBufPos Then
                        .LastError = vbNullString : .LastBufPos = .BufPos
                    End If
                    ParsePrimary = True
                    Exit Function
                End If
                .BufPos = p130
                .ThunkPos = q130
            Else
                .BufPos = p130
                .ThunkPos = q130
            End If
        End If
        If ParseLiteralNoSpace() Then
            pvPushAction ucsAct_2_Primary
            If ParseIgnoreCase() Then
                pvPushAction ucsAct_3_Primary
            End If
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParsePrimary = True
            Exit Function
        End If
        If ParseClassNoSpace() Then
            pvPushAction ucsAct_4_Primary
            If ParseIgnoreCase() Then
                pvPushAction ucsAct_5_Primary
            End If
            Call ParseSpacing
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParsePrimary = True
            Exit Function
        End If
        If ParseDOT() Then
            pvPushAction ucsAct_6_Primary
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParsePrimary = True
            Exit Function
        End If
        If ParseAction() Then
            pvPushAction ucsAct_7_Primary
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParsePrimary = True
            Exit Function
        End If
        If ParseBEGIN() Then
            pvPushAction ucsAct_8_Primary
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParsePrimary = True
            Exit Function
        End If
        If ParseEND() Then
            pvPushAction ucsAct_9_Primary
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParsePrimary = True
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
    Dim p245 As Long
    Dim p254 As Long

    With ctx
        p245 = .BufPos
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
                .BufPos = p245
            End If
        End If
        p254 = .BufPos
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
                .BufPos = p254
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
    Dim p217 As Long
    Dim q217 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long
    Dim p229 As Long
    Dim q229 As Long

    With ctx
        p217 = .BufPos
        q217 = .ThunkPos
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
                .BufPos = p217
                .ThunkPos = q217
            End If
        End If
        p229 = .BufPos
        q229 = .ThunkPos
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
                .BufPos = p229
                .ThunkPos = q229
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
    Dim p265 As Long
    Dim q265 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        p265 = .BufPos
        q265 = .ThunkPos
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
                .BufPos = p265
                .ThunkPos = q265
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
    Dim p187 As Long
    Dim q187 As Long

    With ctx
        If ParseAT() Then
            p187 = .BufPos
            q187 = .ThunkPos
            If ParseIdentifier() Then
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p187
                .ThunkPos = q187
            End If
            If ParseLiteral() Then
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString : .LastBufPos = .BufPos
                End If
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p187
                .ThunkPos = q187
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
    Dim p193 As Long
    Dim q193 As Long

    With ctx
        p193 = .BufPos
        q193 = .ThunkPos
        If ParseIdentifier() Then
            Call ParseSEMI
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p193
            .ThunkPos = q193
        End If
        If ParseLiteral() Then
            Call ParseSEMI
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p193
            .ThunkPos = q193
        End If
        If ParseAction() Then
            Call ParseSEMI
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString : .LastBufPos = .BufPos
            End If
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p193
            .ThunkPos = q193
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
    Dim p285 As Long
    Dim p288 As Long
    Dim p294 As Long

    With ctx
        p285 = .BufPos
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
                .BufPos = p285
            End Select
        End If
        p288 = .BufPos
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
                        .BufPos = p288
                    End Select
                Case Else
                    .BufPos = p288
                End Select
            Case Else
                .BufPos = p288
            End Select
        End If
        p294 = .BufPos
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
                .BufPos = p294
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
    Dim p277 As Long

    With ctx
        p277 = .BufPos
        If ParseChar() Then
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = p277
                Else
                    If ParseChar() Then
                        If .BufPos > .LastBufPos Then
                            .LastError = vbNullString : .LastBufPos = .BufPos
                        End If
                        ParseRange = True
                        Exit Function
                    Else
                        .BufPos = p277
                    End If
                End If
            Else
                .BufPos = p277
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
    Dim p312 As Long

    With ctx
        p312 = .BufPos
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = p312
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
    Dim p359 As Long
    Dim p357 As Long

    With ctx
        p359 = .BufPos
        If .BufData(.BufPos) = 35 Then              ' "#"
            .BufPos = .BufPos + 1
            Do
                p357 = .BufPos
                If ParseEndOfLine() Then
                    .BufPos = p357
                    Exit Do
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p357
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
                .BufPos = p359
            End If
        End If
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 47 Then ' "//"
            .BufPos = .BufPos + 2
            Do
                p357 = .BufPos
                If ParseEndOfLine() Then
                    .BufPos = p357
                    Exit Do
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p357
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
                .BufPos = p359
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
    Dim p390 As Long

    With ctx
        p390 = .BufPos
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
                .BufPos = p390
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
    With ctx
        Select Case eAction
        Case ucsAct_2_Setting
             With .UserData: .AppendSetting .Pop, Mid$(ctx.Contents, lOffset, lSize): End With
        Case ucsAct_1_Setting
             With .UserData: .Push .MakeString(lOffset, Mid$(ctx.Contents, lOffset, lSize)): End With
        Case ucsAct_3_Definition
             With .UserData: Dim lNode As Long: lNode = .Pop(): .RuleExpression(.Pop()) = lNode: End With
        Case ucsAct_2_Definition
             With .UserData: .RuleDescription(.Top) = Mid$(ctx.Contents, lOffset, lSize): End With
        Case ucsAct_1_Definition
               With .UserData
                                                                                lNode = .BeginRule(.FindRule(lOffset, Mid$(ctx.Contents, lOffset, lSize)))
                                                                                If .RuleExpression(lNode) <> 0 Then
                                                                                    ConsoleError "Rule '%1' redefined" & vbCrLf, Mid$(ctx.Contents, lOffset, lSize)
                                                                                End If
                                                                                .Push lNode
                                                                            End With
        Case ucsAct_1_Expression
             With .UserData: lNode = .Pop(): .Push .AppendChild(.MakeAlternate(lOffset, .Pop), lNode): End With
        Case ucsAct_2_Sequence
             With .UserData: .Push .MakePredicate(lOffset, "1"): End With
        Case ucsAct_1_Sequence
             With .UserData: lNode = .Pop(): .Push .AppendSequence(.MakeSequence(lOffset, .Pop), lNode): End With
        Case ucsAct_1_Error
             With .UserData: .Push .AppendChild(.MakeAlternate(lOffset, .Pop), .MakeErrorReference(lOffset, Mid$(ctx.Contents, lOffset, lSize))): End With
        Case ucsAct_2_Error
             With .UserData: .Push .AppendChild(.MakeAlternate(lOffset, .Pop), .MakeErrorString(lOffset, Mid$(ctx.Contents, lOffset, lSize))): End With
        Case ucsAct_1_Prefix
             With .UserData: .Push .MakePredicate(lOffset, Mid$(ctx.Contents, lOffset, lSize)): End With
        Case ucsAct_2_Prefix
             With .UserData: .Push .MakePeekFor(lOffset, .Pop): End With
        Case ucsAct_3_Prefix
             With .UserData: .Push .MakePeekNot(lOffset, .Pop): End With
        Case ucsAct_2_Variable
             With .UserData: .Push .SetVariable(lOffset, .Pop, .Pop): End With
        Case ucsAct_1_Variable
             With .UserData: .Push .MakeVariable(lOffset, Mid$(ctx.Contents, lOffset, lSize)): End With
        Case ucsAct_1_Suffix
             With .UserData: .Push .MakeQuery(lOffset, .Pop): End With
        Case ucsAct_2_Suffix
             With .UserData: .Push .MakeStar(lOffset, .Pop): End With
        Case ucsAct_3_Suffix
             With .UserData: .Push .MakePlus(lOffset, .Pop): End With
        Case ucsAct_1_Primary
             With .UserData: .Push .MakeReference(lOffset, Mid$(ctx.Contents, lOffset, lSize)): End With
        Case ucsAct_3_Primary
             With .UserData: .SetIgnoreCase .Top, (Mid$(ctx.Contents, lOffset, lSize) = "i"): End With
        Case ucsAct_2_Primary
             With .UserData: .Push .MakeString(lOffset, Mid$(ctx.Contents, lOffset, lSize)): End With
        Case ucsAct_5_Primary
             With .UserData: .SetIgnoreCase .Top, (Mid$(ctx.Contents, lOffset, lSize) = "i"): End With
        Case ucsAct_4_Primary
             With .UserData: .Push .MakeClass(lOffset, Mid$(ctx.Contents, lOffset, lSize)): End With
        Case ucsAct_6_Primary
             With .UserData: .Push .MakeDot(lOffset): End With
        Case ucsAct_7_Primary
             With .UserData: .Push .MakeAction(lOffset, Mid$(ctx.Contents, lOffset, lSize)): End With
        Case ucsAct_8_Primary
             With .UserData: .Push .MakePredicate(lOffset, "YY_BEGIN"): End With
        Case ucsAct_9_Primary
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
