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
' Auto-generated on 11.7.2018 19:44:42
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
    LastError           As String
    UserData            As cTree
End Type

Private ctx                     As UcsParserType

'=========================================================================
' Properties
'=========================================================================

Property Get LastError() As String
    LastError = ctx.LastError
End Property

Property Get ParserVersion() As String
    ParserVersion = "11.7.2018 19:44:42"
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
    uEmpty.LastError = ctx.LastError
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
    Dim p7 As Long
    Dim q7 As Long
    Dim i6 As Long

    With ctx
        p7 = .BufPos
        q7 = .ThunkPos
        Call ParseSpacing
        For i6 = 0 To LNG_MAXINT
            If Not ParseDefinition() Then
                Exit For
            End If
        Next
        If i6 <> 0 Then
            If ParseEndOfFile() Then
                ParseGrammar = True
                Exit Function
            Else
                .BufPos = p7
                .ThunkPos = q7
            End If
        Else
            .BufPos = p7
            .ThunkPos = q7
        End If
    End With
End Function

Public Sub ParseSpacing()
    Do
        If ParseSpace() Then
            GoTo L1
        Else
            If ParseComment() Then
                GoTo L1
            Else
                Exit Do
            End If
        End If
L1:
    Loop
End Sub

Public Function ParseDefinition() As Boolean
    Dim p13 As Long
    Dim q13 As Long

    With ctx
        p13 = .BufPos
        q13 = .ThunkPos
        If ParseIdentifier() Then
            pvPushAction ucsAct_1_Definition
            If ParseLiteral() Then
                pvPushAction ucsAct_2_Definition
            End If
            If ParseAssign() Then
                Call ParseExpression
                Call ParseSEMI
                pvPushAction ucsAct_3_Definition
                ParseDefinition = True
                Exit Function
            Else
                .BufPos = p13
                .ThunkPos = q13
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
    Dim p144 As Long
    Dim q144 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long
    Dim p157 As Long
    Dim q157 As Long

    With ctx
        p144 = .BufPos
        q144 = .ThunkPos
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
                .BufPos = p144
                .ThunkPos = q144
            End If
        End If
        p157 = .BufPos
        q157 = .ThunkPos
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
                .BufPos = p157
                .ThunkPos = q157
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
        GoTo L5
    Else
        pvPushAction ucsAct_2_Sequence
    End If
L5:
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
    Dim p56 As Long
    Dim q56 As Long
    Dim p54 As Long
    Dim q54 As Long
    Dim p60 As Long
    Dim q60 As Long

    With ctx
        p56 = .BufPos
        q56 = .ThunkPos
        If ParseAND() Then
            If ParseAction() Then
                pvPushAction ucsAct_1_Prefix
                ParsePrefix = True
                Exit Function
            Else
                .BufPos = p56
                .ThunkPos = q56
            End If
        End If
        p54 = .BufPos
        q54 = .ThunkPos
        If ParseAND() Then
            If ParseVariable() Then
                pvPushAction ucsAct_2_Prefix
                ParsePrefix = True
                Exit Function
            Else
                .BufPos = p54
                .ThunkPos = q54
            End If
        End If
        p60 = .BufPos
        q60 = .ThunkPos
        If ParseNOT() Then
            If ParseVariable() Then
                pvPushAction ucsAct_3_Prefix
                ParsePrefix = True
                Exit Function
            Else
                .BufPos = p60
                .ThunkPos = q60
            End If
        End If
        If ParseVariable() Then
            ParsePrefix = True
            Exit Function
        Else
            .BufPos = p56
            .ThunkPos = q56
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
    Dim p292 As Long
    Dim q292 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        p292 = .BufPos
        q292 = .ThunkPos
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            lCaptureBegin = .BufPos
            Do
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If ParseActNonCapture() Then
                        GoTo L6
                    Else
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        GoTo L6
                    Else
                        If ParseActNonCapture() Then
                            GoTo L6
                        Else
                            Exit Do
                        End If
                    End If
                End Select
L6:
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
                .BufPos = p292
                .ThunkPos = q292
            End If
        End If
    End With
End Function

Public Function ParseVariable() As Boolean
    Dim p72 As Long
    Dim q72 As Long

    With ctx
        p72 = .BufPos
        q72 = .ThunkPos
        If ParseIdentifier() Then
            pvPushAction ucsAct_1_Variable
            If ParseCOLON() Then
                If ParseSuffix() Then
                    pvPushAction ucsAct_2_Variable
                    ParseVariable = True
                    Exit Function
                Else
                    .BufPos = p72
                    .ThunkPos = q72
                End If
            Else
                .BufPos = p72
                .ThunkPos = q72
            End If
        End If
        If ParseSuffix() Then
            ParseVariable = True
            Exit Function
        Else
            .BufPos = p72
            .ThunkPos = q72
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
            GoTo L7
        End If
        If ParseSTAR() Then
            pvPushAction ucsAct_2_Suffix
            GoTo L7
        End If
        If ParsePLUS() Then
            pvPushAction ucsAct_3_Suffix
        End If
L7:
        ParseSuffix = True
    End If
End Function

Public Function ParsePrimary() As Boolean
    Dim p97 As Long
    Dim q97 As Long
    Dim p95 As Long
    Dim q95 As Long
    Dim p102 As Long
    Dim q102 As Long

    With ctx
        p97 = .BufPos
        q97 = .ThunkPos
        If ParseIdentifier() Then
            p95 = .BufPos
            q95 = .ThunkPos
            Call ParseLitNonCapture
            If ParseAssign() Then
                .BufPos = p97
                .ThunkPos = q97
                GoTo L8
            Else
                .BufPos = p95
                .ThunkPos = q95
                pvPushAction ucsAct_1_Primary
                ParsePrimary = True
                Exit Function
            End If
L8:
        End If
        p102 = .BufPos
        q102 = .ThunkPos
        If ParseOPEN() Then
            Call ParseExpression
            If ParseCLOSE() Then
                ParsePrimary = True
                Exit Function
            Else
                .BufPos = p102
                .ThunkPos = q102
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
    Dim p174 As Long
    Dim p183 As Long

    With ctx
        p174 = .BufPos
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
                .BufPos = p174
            End If
        End If
        p183 = .BufPos
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
                .BufPos = p183
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
    Dim p189 As Long
    Dim q189 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        p189 = .BufPos
        q189 = .ThunkPos
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
                .BufPos = p189
                .ThunkPos = q189
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

Public Function ParseChar() As Boolean
    Dim p210 As Long
    Dim p213 As Long
    Dim p219 As Long

    With ctx
        p210 = .BufPos
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 97, 98, 101, 102, 110, 114, 116, 118, 39, 34, 91, 93, 92 ' [abefnrtv'"\[\]\\]
                .BufPos = .BufPos + 1
                ParseChar = True
                Exit Function
            Case Else
                .BufPos = p210
            End Select
        End If
        p213 = .BufPos
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
                        .BufPos = p213
                    End Select
                Case Else
                    .BufPos = p213
                End Select
            Case Else
                .BufPos = p213
            End Select
        End If
        p219 = .BufPos
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
                .BufPos = p219
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
    Dim p202 As Long

    With ctx
        p202 = .BufPos
        If ParseChar() Then
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = p202
                Else
                    If ParseChar() Then
                        ParseRange = True
                        Exit Function
                    Else
                        .BufPos = p202
                    End If
                End If
            Else
                .BufPos = p202
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
    Dim p237 As Long

    With ctx
        p237 = .BufPos
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = p237
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
    Dim p278 As Long
    Dim p276 As Long

    With ctx
        p278 = .BufPos
        If .BufData(.BufPos) = 35 Then              ' "#"
            .BufPos = .BufPos + 1
            Do
                p276 = .BufPos
                If ParseEndOfLine() Then
                    .BufPos = p276
                    Exit Do
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p276
                    Exit Do
                End If
            Loop
            If ParseEndOfLine() Then
                ParseComment = True
                Exit Function
            Else
                .BufPos = p278
            End If
        End If
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 47 Then ' "//"
            .BufPos = .BufPos + 2
            Do
                p276 = .BufPos
                If ParseEndOfLine() Then
                    .BufPos = p276
                    Exit Do
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p276
                    Exit Do
                End If
            Loop
            If ParseEndOfLine() Then
                ParseComment = True
                Exit Function
            Else
                .BufPos = p278
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
    Dim p307 As Long

    With ctx
        p307 = .BufPos
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Do
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    If ParseActNonCapture() Then
                        GoTo L12
                    Else
                        Exit Do
                    End If
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        GoTo L12
                    Else
                        If ParseActNonCapture() Then
                            GoTo L12
                        Else
                            Exit Do
                        End If
                    End If
                End Select
L12:
            Loop
            If .BufData(.BufPos) = 125 Then         ' "}"
                .BufPos = .BufPos + 1
                ParseActNonCapture = True
                Exit Function
            Else
                .BufPos = p307
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
