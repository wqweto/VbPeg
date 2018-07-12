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
' Auto-generated on 12.7.2018 13:15:54
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function RtlCompareMemory Lib "ntdll" (Source1 As Any, Source2 As Any, ByVal Length As Long) As Long

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
    ucsAct_1_Prefix
    ucsAct_2_Prefix
    ucsAct_3_Prefix
    ucsAct_2_Variable
    ucsAct_1_Variable
    ucsAct_1_Suffix
    ucsAct_2_Suffix
    ucsAct_3_Suffix
    ucsAct_1_Primary
    ucsAct_2_Primary
    ucsAct_3_Primary
    ucsAct_4_Primary
    ucsAct_5_Primary
    ucsAct_6_Primary
    ucsAct_7_Primary
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
    LastError           As Variant
    UserData            As cTree
End Type

Private ctx                     As UcsParserType

'=========================================================================
' Properties
'=========================================================================

Property Get LastError() As Variant
    If IsObject(ctx.LastError) Then
        Set LastError = ctx.LastError
    Else
        LastError = ctx.LastError
    End If
End Property

Property Get ParserVersion() As String
    ParserVersion = "12.7.2018 13:15:54"
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Match(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As cTree) As Long
    If BeginMatch(sSubject, StartPos, UserData) Then
        If ParseGrammar() Then
            Match = EndMatch(UserData)
        End If
    End If
End Function

Public Function BeginMatch(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As cTree) As Boolean
    With ctx
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
    If IsObject(ctx.LastError) Then
        Set uEmpty.LastError = ctx.LastError
    Else
        uEmpty.LastError = ctx.LastError
    End If
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
    Dim p144 As Long
    Dim q144 As Long

    With ctx
        p144 = .BufPos
        q144 = .ThunkPos
        If ParseSettingName() Then
            pvPushAction ucsAct_1_Setting
            If  ClearErrors() And SetUserDataTag(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin))  Then
                If ParseSettingValue() Then
                    pvPushAction ucsAct_2_Setting
                    ParseSetting = True
                    Exit Function
                End If
                If  AppendError("Invalid value for setting '" & .UserData.Tag & "'")  Then
                    ParseSetting = True
                    Exit Function
                End If
                .BufPos = p144
                .ThunkPos = q144
            Else
                .BufPos = p144
                .ThunkPos = q144
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
                    ParseDefinition = True
                    Exit Function
                    If  AppendError("Invalid definition for rule '" & .UserData.Tag & "'")  Then
                        Call ParseSEMI
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
            ParseIdentifier = True
        End Select
    End With
End Function

Public Function ParseLiteral() As Boolean
    Dim p182 As Long
    Dim q182 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long
    Dim p195 As Long
    Dim q195 As Long

    With ctx
        p182 = .BufPos
        q182 = .ThunkPos
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
                Call ParseSpacing
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                ParseLiteral = True
                Exit Function
            Else
                .BufPos = p182
                .ThunkPos = q182
            End If
        End If
        p195 = .BufPos
        q195 = .ThunkPos
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
                Call ParseSpacing
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                ParseLiteral = True
                Exit Function
            Else
                .BufPos = p195
                .ThunkPos = q195
            End If
        End If
    End With
End Function

Public Function ParseAssign() As Boolean
    If ParseLEFTARROW() Then
        ParseAssign = True
        Exit Function
    End If
    If ParseEQUAL() Then
        ParseAssign = True
    End If
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
            ParseSEMI = True
        End If
    End With
End Function

Public Sub ParseSequence()
    If ParsePrefix() Then
        Do
            If Not ParsePrefix() Then
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
    If ParseSLASH() Then
        ParseChoice = True
        Exit Function
    End If
    If ParseBAR() Then
        ParseChoice = True
    End If
End Function

Public Function ParsePrefix() As Boolean
    Dim p65 As Long
    Dim q65 As Long
    Dim p63 As Long
    Dim q63 As Long
    Dim p69 As Long
    Dim q69 As Long

    With ctx
        p65 = .BufPos
        q65 = .ThunkPos
        If ParseAND() Then
            If ParseAction() Then
                pvPushAction ucsAct_1_Prefix
                ParsePrefix = True
                Exit Function
            Else
                .BufPos = p65
                .ThunkPos = q65
            End If
        End If
        p63 = .BufPos
        q63 = .ThunkPos
        If ParseAND() Then
            If ParseVariable() Then
                pvPushAction ucsAct_2_Prefix
                ParsePrefix = True
                Exit Function
            Else
                .BufPos = p63
                .ThunkPos = q63
            End If
        End If
        p69 = .BufPos
        q69 = .ThunkPos
        If ParseNOT() Then
            If ParseVariable() Then
                pvPushAction ucsAct_3_Prefix
                ParsePrefix = True
                Exit Function
            Else
                .BufPos = p69
                .ThunkPos = q69
            End If
        End If
        If ParseVariable() Then
            ParsePrefix = True
            Exit Function
        Else
            .BufPos = p65
            .ThunkPos = q65
        End If
    End With
End Function

Private Function ParseAND() As Boolean
    With ctx
        If .BufData(.BufPos) = 38 Then              ' "&"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseAND = True
        End If
    End With
End Function

Public Function ParseAction() As Boolean
    Dim p337 As Long
    Dim q337 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        p337 = .BufPos
        q337 = .ThunkPos
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
                    If ParseActNonCapture() Then
                        GoTo L9
                    Else
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        GoTo L9
                    Else
                        If ParseActNonCapture() Then
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
                ParseAction = True
                Exit Function
            Else
                .BufPos = p337
                .ThunkPos = q337
            End If
        End If
    End With
End Function

Public Function ParseVariable() As Boolean
    Dim p81 As Long
    Dim q81 As Long

    With ctx
        p81 = .BufPos
        q81 = .ThunkPos
        If ParseIdentifier() Then
            pvPushAction ucsAct_1_Variable
            If ParseCOLON() Then
                If ParseSuffix() Then
                    pvPushAction ucsAct_2_Variable
                    ParseVariable = True
                    Exit Function
                Else
                    .BufPos = p81
                    .ThunkPos = q81
                End If
            Else
                .BufPos = p81
                .ThunkPos = q81
            End If
        End If
        If ParseSuffix() Then
            ParseVariable = True
            Exit Function
        Else
            .BufPos = p81
            .ThunkPos = q81
        End If
    End With
End Function

Private Function ParseNOT() As Boolean
    With ctx
        If .BufData(.BufPos) = 33 Then              ' "!"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseNOT = True
        End If
    End With
End Function

Private Function ParseCOLON() As Boolean
    With ctx
        If .BufData(.BufPos) = 58 Then              ' ":"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseCOLON = True
        End If
    End With
End Function

Public Function ParseSuffix() As Boolean
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
        ParseSuffix = True
    End If
End Function

Public Function ParsePrimary() As Boolean
    Dim p106 As Long
    Dim q106 As Long
    Dim p104 As Long
    Dim q104 As Long
    Dim p111 As Long
    Dim q111 As Long
    Dim p115 As Long
    Dim q115 As Long

    With ctx
        p106 = .BufPos
        q106 = .ThunkPos
        If ParseIdentifier() Then
            p104 = .BufPos
            q104 = .ThunkPos
            Call ParseLitNonCapture
            If ParseAssign() Then
                .BufPos = p106
                .ThunkPos = q106
                GoTo L11
            Else
                .BufPos = p104
                .ThunkPos = q104
                pvPushAction ucsAct_1_Primary
                ParsePrimary = True
                Exit Function
            End If
L11:
        End If
        p111 = .BufPos
        q111 = .ThunkPos
        If ParseOPEN() Then
            If  ClearErrors()  Then
                p115 = .BufPos
                q115 = .ThunkPos
                Call ParseExpression
                If ParseCLOSE() Then
                    ParsePrimary = True
                    Exit Function
                Else
                    .BufPos = p115
                    .ThunkPos = q115
                End If
                If  AppendError("Invalid expression")  Then
                    ParsePrimary = True
                    Exit Function
                End If
                .BufPos = p111
                .ThunkPos = q111
            Else
                .BufPos = p111
                .ThunkPos = q111
            End If
        End If
        If ParseLiteral() Then
            pvPushAction ucsAct_2_Primary
            ParsePrimary = True
            Exit Function
        End If
        If ParseClass() Then
            pvPushAction ucsAct_3_Primary
            ParsePrimary = True
            Exit Function
        End If
        If ParseDOT() Then
            pvPushAction ucsAct_4_Primary
            ParsePrimary = True
            Exit Function
        End If
        If ParseAction() Then
            pvPushAction ucsAct_5_Primary
            ParsePrimary = True
            Exit Function
        End If
        If ParseBEGIN() Then
            pvPushAction ucsAct_6_Primary
            ParsePrimary = True
            Exit Function
        End If
        If ParseEND() Then
            pvPushAction ucsAct_7_Primary
            ParsePrimary = True
        End If
    End With
End Function

Private Function ParseQUESTION() As Boolean
    With ctx
        If .BufData(.BufPos) = 63 Then              ' "?"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseQUESTION = True
        End If
    End With
End Function

Private Function ParseSTAR() As Boolean
    With ctx
        If .BufData(.BufPos) = 42 Then              ' "*"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseSTAR = True
        End If
    End With
End Function

Private Function ParsePLUS() As Boolean
    With ctx
        If .BufData(.BufPos) = 43 Then              ' "+"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParsePLUS = True
        End If
    End With
End Function

Public Function ParseLitNonCapture() As Boolean
    Dim p212 As Long
    Dim p221 As Long

    With ctx
        p212 = .BufPos
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
                ParseLitNonCapture = True
                Exit Function
            Else
                .BufPos = p212
            End If
        End If
        p221 = .BufPos
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
                ParseLitNonCapture = True
                Exit Function
            Else
                .BufPos = p221
            End If
        End If
    End With
End Function

Private Function ParseOPEN() As Boolean
    With ctx
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseOPEN = True
        End If
    End With
End Function

Private Function ParseCLOSE() As Boolean
    With ctx
        If .BufData(.BufPos) = 41 Then              ' ")"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseCLOSE = True
        End If
    End With
End Function

Public Function ParseClass() As Boolean
    Dim p227 As Long
    Dim q227 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        p227 = .BufPos
        q227 = .ThunkPos
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
                Call ParseSpacing
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                ParseClass = True
                Exit Function
            Else
                .BufPos = p227
                .ThunkPos = q227
            End If
        End If
    End With
End Function

Private Function ParseDOT() As Boolean
    With ctx
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseDOT = True
        End If
    End With
End Function

Private Function ParseBEGIN() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseBEGIN = True
        End If
    End With
End Function

Private Function ParseEND() As Boolean
    With ctx
        If .BufData(.BufPos) = 62 Then              ' ">"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseEND = True
        End If
    End With
End Function

Public Function ParseSettingName() As Boolean
    Dim p155 As Long
    Dim q155 As Long

    With ctx
        If .BufData(.BufPos) = 64 Then              ' "@"
            .BufPos = .BufPos + 1
            p155 = .BufPos
            q155 = .ThunkPos
            If ParseIdentifier() Then
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p155
                .ThunkPos = q155
            End If
            If ParseLiteral() Then
                ParseSettingName = True
                Exit Function
            Else
                .BufPos = p155
                .ThunkPos = q155
            End If
            Call ParseEMPTY
            ParseSettingName = True
        End If
    End With
End Function

Public Function ParseSettingValue() As Boolean
    Dim p161 As Long
    Dim q161 As Long

    With ctx
        p161 = .BufPos
        q161 = .ThunkPos
        If ParseIdentifier() Then
            Call ParseSEMI
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p161
            .ThunkPos = q161
        End If
        If ParseLiteral() Then
            Call ParseSEMI
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p161
            .ThunkPos = q161
        End If
        If ParseAction() Then
            Call ParseSEMI
            ParseSettingValue = True
            Exit Function
        Else
            .BufPos = p161
            .ThunkPos = q161
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
    Dim p248 As Long
    Dim p251 As Long
    Dim p257 As Long

    With ctx
        p248 = .BufPos
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 97, 98, 101, 102, 110, 114, 116, 118, 39, 34, 91, 93, 92 ' [abefnrtv'"\[\]\\]
                .BufPos = .BufPos + 1
                ParseChar = True
                Exit Function
            Case Else
                .BufPos = p248
            End Select
        End If
        p251 = .BufPos
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
                        ParseChar = True
                        Exit Function
                    Case Else
                        .BufPos = p251
                    End Select
                Case Else
                    .BufPos = p251
                End Select
            Case Else
                .BufPos = p251
            End Select
        End If
        p257 = .BufPos
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 48 To 55                           ' [0-7]
                .BufPos = .BufPos + 1
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                End Select
                ParseChar = True
                Exit Function
            Case Else
                .BufPos = p257
            End Select
        End If
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 45 Then ' "\\-"
            .BufPos = .BufPos + 2
            ParseChar = True
            Exit Function
        End If
        If Not .BufData(.BufPos) = 92 Then          ' "\\"
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
                ParseChar = True
            End If
        End If
    End With
End Function

Public Function ParseRange() As Boolean
    Dim p240 As Long

    With ctx
        p240 = .BufPos
        If ParseChar() Then
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = p240
                Else
                    If ParseChar() Then
                        ParseRange = True
                        Exit Function
                    Else
                        .BufPos = p240
                    End If
                End If
            Else
                .BufPos = p240
            End If
        End If
        If ParseChar() Then
            ParseRange = True
        End If
    End With
End Function

Private Function ParseLEFTARROW() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 And .BufData(.BufPos + 1) = 45 Then ' "<-"
            .BufPos = .BufPos + 2
            Call ParseSpacing
            ParseLEFTARROW = True
        End If
    End With
End Function

Private Function ParseSLASH() As Boolean
    Dim p275 As Long

    With ctx
        p275 = .BufPos
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = p275
            Else
                Call ParseSpacing
                ParseSLASH = True
            End If
        End If
    End With
End Function

Public Function ParseSpace() As Boolean
    With ctx
        If .BufData(.BufPos) = 32 Then              ' " "
            .BufPos = .BufPos + 1
            ParseSpace = True
            Exit Function
        End If
        If .BufData(.BufPos) = 9 Then               ' "\t"
            .BufPos = .BufPos + 1
            ParseSpace = True
            Exit Function
        End If
        If ParseEndOfLine() Then
            ParseSpace = True
        End If
    End With
End Function

Public Function ParseComment() As Boolean
    Dim p322 As Long
    Dim p320 As Long

    With ctx
        p322 = .BufPos
        If .BufData(.BufPos) = 35 Then              ' "#"
            .BufPos = .BufPos + 1
            Do
                p320 = .BufPos
                If ParseEndOfLine() Then
                    .BufPos = p320
                    Exit Do
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p320
                    Exit Do
                End If
            Loop
            If ParseEndOfLine() Then
                ParseComment = True
                Exit Function
            Else
                .BufPos = p322
            End If
        End If
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 47 Then ' "//"
            .BufPos = .BufPos + 2
            Do
                p320 = .BufPos
                If ParseEndOfLine() Then
                    .BufPos = p320
                    Exit Do
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p320
                    Exit Do
                End If
            Loop
            If ParseEndOfLine() Then
                ParseComment = True
                Exit Function
            Else
                .BufPos = p322
            End If
        End If
    End With
End Function

Public Function ParseEndOfLine() As Boolean
    With ctx
        If .BufData(.BufPos) = 13 And .BufData(.BufPos + 1) = 10 Then ' "\r\n"
            .BufPos = .BufPos + 2
            ParseEndOfLine = True
            Exit Function
        End If
        If .BufData(.BufPos) = 10 Then              ' "\n"
            .BufPos = .BufPos + 1
            ParseEndOfLine = True
            Exit Function
        End If
        If .BufData(.BufPos) = 13 Then              ' "\r"
            .BufPos = .BufPos + 1
            ParseEndOfLine = True
        End If
    End With
End Function

Public Function ParseActNonCapture() As Boolean
    Dim p353 As Long

    With ctx
        p353 = .BufPos
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Do
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If ParseActNonCapture() Then
                        GoTo L20
                    Else
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        GoTo L20
                    Else
                        If ParseActNonCapture() Then
                            GoTo L20
                        Else
                            Exit Do
                        End If
                    End If
                End Select
L20:
            Loop
            If .BufData(.BufPos) = 125 Then         ' "}"
                .BufPos = .BufPos + 1
                Call ParseSpacing
                ParseActNonCapture = True
                Exit Function
            Else
                .BufPos = p353
            End If
        End If
    End With
End Function

Private Function ParseEQUAL() As Boolean
    With ctx
        If .BufData(.BufPos) = 61 Then              ' "="
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseEQUAL = True
        End If
    End With
End Function

Private Function ParseBAR() As Boolean
    With ctx
        If .BufData(.BufPos) = 124 Then             ' "|"
            .BufPos = .BufPos + 1
            Call ParseSpacing
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
             With .UserData: .Push .MakeString(Mid$(ctx.Contents, lOffset, lSize)): End With
        Case ucsAct_3_Definition
             With .UserData: Dim lNode As Long: lNode = .Pop(): .RuleExpression(.Pop()) = lNode: End With
        Case ucsAct_2_Definition
             With .UserData: .RuleDescription(.Top) = Mid$(ctx.Contents, lOffset, lSize): End With
        Case ucsAct_1_Definition
               With .UserData
                                                                                lNode = .BeginRule(.FindRule(Mid$(ctx.Contents, lOffset, lSize)))
                                                                                If .RuleExpression(lNode) <> 0 Then
                                                                                    ConsoleError "Rule '%1' redefined" & vbCrLf, Mid$(ctx.Contents, lOffset, lSize)
                                                                                End If
                                                                                .Push lNode
                                                                            End With
        Case ucsAct_1_Expression
             With .UserData: lNode = .Pop(): .Push .AppendChild(.MakeAlternate(.Pop), lNode): End With
        Case ucsAct_2_Sequence
             With .UserData: .Push .MakePredicate("1"): End With
        Case ucsAct_1_Sequence
             With .UserData: lNode = .Pop(): .Push .AppendSequence(.MakeSequence(.Pop), lNode): End With
        Case ucsAct_1_Prefix
             With .UserData: .Push .MakePredicate(Mid$(ctx.Contents, lOffset, lSize)): End With
        Case ucsAct_2_Prefix
             With .UserData: .Push .MakePeekFor(.Pop): End With
        Case ucsAct_3_Prefix
             With .UserData: .Push .MakePeekNot(.Pop): End With
        Case ucsAct_2_Variable
             With .UserData: .Push .SetVariable(.Pop, .Pop): End With
        Case ucsAct_1_Variable
             With .UserData: .Push .MakeVariable(Mid$(ctx.Contents, lOffset, lSize)): End With
        Case ucsAct_1_Suffix
             With .UserData: .Push .MakeQuery(.Pop): End With
        Case ucsAct_2_Suffix
             With .UserData: .Push .MakeStar(.Pop): End With
        Case ucsAct_3_Suffix
             With .UserData: .Push .MakePlus(.Pop): End With
        Case ucsAct_1_Primary
             With .UserData: .Push .MakeReference(.FindRule(Mid$(ctx.Contents, lOffset, lSize))): End With
        Case ucsAct_2_Primary
             With .UserData: .Push .MakeString(Mid$(ctx.Contents, lOffset, lSize)): End With
        Case ucsAct_3_Primary
             With .UserData: .Push .MakeClass(Mid$(ctx.Contents, lOffset, lSize)): End With
        Case ucsAct_4_Primary
             With .UserData: .Push .MakeDot(): End With
        Case ucsAct_5_Primary
             With .UserData: .Push .MakeAction(Mid$(ctx.Contents, lOffset, lSize)): End With
        Case ucsAct_6_Primary
             With .UserData: .Push .MakePredicate("YY_BEGIN"): End With
        Case ucsAct_7_Primary
             With .UserData: .Push .MakePredicate("YY_END"): End With
        End Select
    End With
End Sub

'= user-defined functions ================================================

Private Function ClearErrors() As Boolean
    Dim lIdx            As Long
    
    If IsObject(ctx.LastError) Then
        For lIdx = ctx.LastError.Count To 1 Step -1
            If ctx.LastError.Item(lIdx)(1) <= ctx.BufPos Then
                ctx.LastError.Remove lIdx
            End If
        Next
        If ctx.LastError.Count = 0 Then
            ctx.LastError = Empty
        End If
    End If
    ClearErrors = True
End Function

Private Function AppendError(sText As String) As Boolean
    If Not IsObject(ctx.LastError) Then
        Set ctx.LastError = New Collection
    End If
    ctx.LastError.Add Array(sText, ctx.BufPos)
End Function

Private Function SetUserDataTag(sText As String) As Boolean
    ctx.UserData.Tag = sText
    SetUserDataTag = True
End Function
