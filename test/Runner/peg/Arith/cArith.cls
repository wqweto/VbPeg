VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArith"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Auto-generated on 25.7.2018 18:55:45
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
    ucsAct_2_Expression
    ucsAct_1_Expression
    ucsAct_2_Term
    ucsAct_1_Term
    ucsAct_1_Factor
    ucsAct_1_Integer
    ucsActVarAlloc = -1
    ucsActVarSet = -2
    ucsActResultClear = -3
    ucsActResultSet = -4
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
    UserData            As Variant
    VarResult           As Variant
    VarStack()          As Variant
    VarPos              As Long
    '-- from @members setting
    Indent              As Long
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
    ParserVersion = "25.7.2018 18:55:45"
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Match(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant, Optional Result As Variant) As Long
    If BeginMatch(sSubject, StartPos, UserData) Then
        If ParseExpression() Then
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

Public Function BeginMatch(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant) As Boolean
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
        ReDim .VarStack(0 To 1024) As Variant
        For lIdx = 0 To .ThunkPos - 1
            Select Case .ThunkData(lIdx).Action
            Case ucsActVarAlloc
                .VarPos = .VarPos + .ThunkData(lIdx).CaptureBegin
            Case ucsActVarSet
                If IsObject(.VarResult) Then
                    Set .VarStack(.VarPos - .ThunkData(lIdx).CaptureBegin) = .VarResult
                Else
                    .VarStack(.VarPos - .ThunkData(lIdx).CaptureBegin) = .VarResult
                End If
            Case ucsActResultClear
                .VarResult = Empty
            Case ucsActResultSet
                With .ThunkData(lIdx)
                    ctx.VarResult = Mid$(ctx.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)
                End With
            Case Else
                With .ThunkData(lIdx)
                    pvImplAction .Action, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin
                End With
            End Select
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

Public Function ParseExpression() As Boolean
    Dim p44 As Long
    Dim q44 As Long
    Dim p32 As Long
    Dim q32 As Long
    Dim lCaptureBegin As Long
    Dim p19 As Long
    Dim q19 As Long
    Dim p25 As Long
    Dim q25 As Long
    Dim lCaptureEnd As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 3
        p44 = .BufPos
        q44 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Expression", 1, ctx.UserData)
        pvPushThunk ucsActResultClear
        If ParseTerm() Then
            pvPushThunk ucsActVarSet, 1
            Do
                p32 = .BufPos
                q32 = .ThunkPos
                Call Parse_
                pvPushThunk ucsActResultClear
                lCaptureBegin = .BufPos
                p19 = .BufPos
                q19 = .ThunkPos
                Call ConsoleTrace(ctx.BufPos + 1, """+""", 0, ctx.UserData)
                If .BufData(.BufPos) = 43 Then      ' "+"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, """+""", 2, ctx.UserData)
                Else
                    If Not (ConsoleTrace(ctx.BufPos + 1, """+""", 3, ctx.UserData) And False) Then
                        p25 = .BufPos
                        q25 = .ThunkPos
                        Call ConsoleTrace(ctx.BufPos + 1, """-""", 0, ctx.UserData)
                        If .BufData(.BufPos) = 45 Then ' "-"
                            .BufPos = .BufPos + 1
                            Call ConsoleTrace(ctx.BufPos + 1, """-""", 2, ctx.UserData)
                            GoTo L1
                        End If
                        If ConsoleTrace(ctx.BufPos + 1, """-""", 3, ctx.UserData) And False Then
                            '--- do nothing
                        End If
                        .BufPos = p19
                        .ThunkPos = q19
                        .BufPos = p32
                        .ThunkPos = q32
                        Exit Do
                    End If
                End If
L1:
                lCaptureEnd = .BufPos
                pvPushThunk ucsActResultSet, lCaptureBegin, lCaptureEnd
                pvPushThunk ucsActVarSet, 2
                Call Parse_
                pvPushThunk ucsActResultClear
                If ParseTerm() Then
                    pvPushThunk ucsActVarSet, 3
                Else
                    .BufPos = p32
                    .ThunkPos = q32
                    Exit Do
                End If
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                pvPushThunk ucsAct_1_Expression, lCaptureBegin, lCaptureEnd
            Loop
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_2_Expression, lCaptureBegin, lCaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "Expression", 2, ctx.UserData)
            pvPushThunk ucsActVarAlloc, -3
            Call pvSetAdvance
            ParseExpression = True
            Exit Function
        Else
            .BufPos = p44
            .ThunkPos = q44
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Expression", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseTerm() As Boolean
    Dim p83 As Long
    Dim q83 As Long
    Dim p71 As Long
    Dim q71 As Long
    Dim lCaptureBegin As Long
    Dim p58 As Long
    Dim q58 As Long
    Dim p64 As Long
    Dim q64 As Long
    Dim lCaptureEnd As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 3
        p83 = .BufPos
        q83 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Term", 1, ctx.UserData)
        pvPushThunk ucsActResultClear
        If ParseFactor() Then
            pvPushThunk ucsActVarSet, 1
            Do
                p71 = .BufPos
                q71 = .ThunkPos
                Call Parse_
                pvPushThunk ucsActResultClear
                lCaptureBegin = .BufPos
                p58 = .BufPos
                q58 = .ThunkPos
                Call ConsoleTrace(ctx.BufPos + 1, """*""i", 0, ctx.UserData)
                If .BufData(.BufPos) = 42 Then      ' "*"i
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, """*""i", 2, ctx.UserData)
                Else
                    If Not (ConsoleTrace(ctx.BufPos + 1, """*""i", 3, ctx.UserData) And False) Then
                        p64 = .BufPos
                        q64 = .ThunkPos
                        Call ConsoleTrace(ctx.BufPos + 1, """/""", 0, ctx.UserData)
                        If .BufData(.BufPos) = 47 Then ' "/"
                            .BufPos = .BufPos + 1
                            Call ConsoleTrace(ctx.BufPos + 1, """/""", 2, ctx.UserData)
                            GoTo L2
                        End If
                        If ConsoleTrace(ctx.BufPos + 1, """/""", 3, ctx.UserData) And False Then
                            '--- do nothing
                        End If
                        .BufPos = p58
                        .ThunkPos = q58
                        .BufPos = p71
                        .ThunkPos = q71
                        Exit Do
                    End If
                End If
L2:
                lCaptureEnd = .BufPos
                pvPushThunk ucsActResultSet, lCaptureBegin, lCaptureEnd
                pvPushThunk ucsActVarSet, 2
                Call Parse_
                pvPushThunk ucsActResultClear
                If ParseFactor() Then
                    pvPushThunk ucsActVarSet, 3
                Else
                    .BufPos = p71
                    .ThunkPos = q71
                    Exit Do
                End If
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                pvPushThunk ucsAct_1_Term, lCaptureBegin, lCaptureEnd
            Loop
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_2_Term, lCaptureBegin, lCaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "Term", 2, ctx.UserData)
            pvPushThunk ucsActVarAlloc, -3
            Call pvSetAdvance
            ParseTerm = True
            Exit Function
        Else
            .BufPos = p83
            .ThunkPos = q83
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Term", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Sub Parse_()
    Dim p149 As Long

    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "_", 1, ctx.UserData)
        Do
            p149 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[ \t\n\r]i", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 32, 9, 10, 13                      ' [ \t\n\r]i
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[ \t\n\r]i", 2, ctx.UserData)
            Case Else
                If Not (ConsoleTrace(ctx.BufPos + 1, "[ \t\n\r]i", 3, ctx.UserData) And False) Then
                    .BufPos = p149
                    Exit Do
                End If
            End Select
        Loop
        Call ConsoleTrace(ctx.BufPos + 1, "_", 2, ctx.UserData)
        Call pvSetAdvance
    End With
End Sub

Public Function ParseFactor() As Boolean
    Dim p114 As Long
    Dim q114 As Long
    Dim p110 As Long
    Dim q110 As Long
    Dim p89 As Long
    Dim q89 As Long
    Dim p102 As Long
    Dim q102 As Long
    Dim p107 As Long
    Dim q107 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        p114 = .BufPos
        q114 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Factor", 1, ctx.UserData)
        p110 = .BufPos
        q110 = .ThunkPos
        p89 = .BufPos
        q89 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, """(""", 0, ctx.UserData)
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, """(""", 2, ctx.UserData)
            Call Parse_
            pvPushThunk ucsActResultClear
            If ParseExpression() Then
                pvPushThunk ucsActVarSet, 1
                Call Parse_
                p102 = .BufPos
                q102 = .ThunkPos
                Call ConsoleTrace(ctx.BufPos + 1, """)""", 0, ctx.UserData)
                If .BufData(.BufPos) = 41 Then      ' ")"
                    .BufPos = .BufPos + 1
                    Call ConsoleTrace(ctx.BufPos + 1, """)""", 2, ctx.UserData)
                    pvPushThunk ucsAct_1_Factor, .CaptureBegin, .CaptureEnd
                    Call ConsoleTrace(ctx.BufPos + 1, "Factor", 2, ctx.UserData)
                    pvPushThunk ucsActVarAlloc, -1
                    Call pvSetAdvance
                    ParseFactor = True
                    Exit Function
                End If
                If ConsoleTrace(ctx.BufPos + 1, """)""", 3, ctx.UserData) And False Then
                    '--- do nothing
                End If
                .BufPos = p110
                .ThunkPos = q110
            Else
                .BufPos = p110
                .ThunkPos = q110
            End If
            GoTo L5
        End If
        If ConsoleTrace(ctx.BufPos + 1, """(""", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p110
        .ThunkPos = q110
L5:
        p107 = .BufPos
        q107 = .ThunkPos
        If ParseInteger() Then
            Call ConsoleTrace(ctx.BufPos + 1, "Factor", 2, ctx.UserData)
            pvPushThunk ucsActVarAlloc, -1
            Call pvSetAdvance
            ParseFactor = True
            Exit Function
        Else
            .BufPos = p107
            .ThunkPos = q107
        End If
        If ParseErrInteger() Then
            Call ConsoleTrace(ctx.BufPos + 1, "Factor", 2, ctx.UserData)
            pvPushThunk ucsActVarAlloc, -1
            Call pvSetAdvance
            ParseFactor = True
            Exit Function
        Else
            .BufPos = p107
            .ThunkPos = q107
        End If
        .BufPos = p110
        .ThunkPos = q110
        .BufPos = p114
        .ThunkPos = q114
        If ConsoleTrace(ctx.BufPos + 1, "Factor", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseInteger() As Boolean
    Dim p131 As Long
    Dim q131 As Long
    Dim lCaptureBegin As Long
    Dim i125 As Long
    Dim p123 As Long
    Dim lCaptureEnd As Long

    With ctx
        p131 = .BufPos
        q131 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Integer", 1, ctx.UserData)
        Call Parse_
        lCaptureBegin = .BufPos
        For i125 = 0 To LNG_MAXINT
            p123 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[0-9]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[0-9]", 2, ctx.UserData)
            Case Else
                If Not (ConsoleTrace(ctx.BufPos + 1, "[0-9]", 3, ctx.UserData) And False) Then
                    .BufPos = p123
                    Exit For
                End If
            End Select
        Next
        If i125 <> 0 Then
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_Integer, lCaptureBegin, lCaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "Integer", 2, ctx.UserData)
            Call pvSetAdvance
            ParseInteger = True
            Exit Function
        Else
            .BufPos = p131
            .ThunkPos = q131
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Integer", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        Call pvSetExpected("integer")
    End With
End Function

Public Function ParseErrInteger() As Boolean
    Dim p143 As Long
    Dim q143 As Long
    Dim i137 As Long
    Dim p136 As Long
    Dim q136 As Long
    Dim p134 As Long
    Dim q134 As Long
    Dim e134 As String

    With ctx
        p143 = .BufPos
        q143 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "ErrInteger", 1, ctx.UserData)
        For i137 = 0 To LNG_MAXINT
            p136 = .BufPos
            q136 = .ThunkPos
            p134 = .BufPos
            q134 = .ThunkPos
            e134 = .LastExpected
            If ParseInteger() Then
                .BufPos = p136
                .ThunkPos = q136
                Exit For
            Else
                .BufPos = p134
                .ThunkPos = q134
                .LastExpected = e134
            End If
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
            Else
                .BufPos = p136
                .ThunkPos = q136
                Exit For
            End If
        Next
        If i137 <> 0 Then
            If ParseInteger() Then
                Call ConsoleTrace(ctx.BufPos + 1, "ErrInteger", 2, ctx.UserData)
                ParseErrInteger = True
                Exit Function
            Else
                .BufPos = p143
                .ThunkPos = q143
            End If
        Else
            .BufPos = p143
            .ThunkPos = q143
        End If
        If ConsoleTrace(ctx.BufPos + 1, "ErrInteger", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Sub pvImplAction(ByVal eAction As UcsParserActionsEnum, ByVal lOffset As Long, ByVal lSize As Long)
    Dim oJson As Object
    With ctx.UserData
    Select Case eAction
    Case ucsAct_2_Expression
         ctx.VarResult = ctx.VarStack(ctx.VarPos - 1)
    Case ucsAct_1_Expression
                If ctx.VarStack(ctx.VarPos - 2) = "+" Then
            ctx.VarStack(ctx.VarPos - 1) = ctx.VarStack(ctx.VarPos - 1) + ctx.VarStack(ctx.VarPos - 3)
        Else
            ctx.VarStack(ctx.VarPos - 1) = ctx.VarStack(ctx.VarPos - 1) - ctx.VarStack(ctx.VarPos - 3)
        End If

    Case ucsAct_2_Term
         ctx.VarResult = ctx.VarStack(ctx.VarPos - 1)
    Case ucsAct_1_Term
                If ctx.VarStack(ctx.VarPos - 2) = "*" Then
            ctx.VarStack(ctx.VarPos - 1) = ctx.VarStack(ctx.VarPos - 1) * ctx.VarStack(ctx.VarPos - 3)
        Else
            ctx.VarStack(ctx.VarPos - 1) = ctx.VarStack(ctx.VarPos - 1) / ctx.VarStack(ctx.VarPos - 3)
        End If

    Case ucsAct_1_Factor
         ctx.VarResult = ctx.VarStack(ctx.VarPos - 1)
    Case ucsAct_1_Integer
         ctx.VarResult = CLng(Mid$(ctx.Contents, lOffset, lSize))
    End Select
    End With
End Sub

'= user-defined functions ================================================

Private Sub Test()

End Sub
