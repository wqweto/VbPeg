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
' Auto-generated on 24.7.2018 14:11:34
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
    ParserVersion = "24.7.2018 14:11:34"
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
    Dim p35 As Long
    Dim q35 As Long
    Dim p23 As Long
    Dim q23 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 3
        p35 = .BufPos
        q35 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Expression", 1, ctx.UserData)
        pvPushThunk ucsActResultClear
        If ParseTerm() Then
            pvPushThunk ucsActVarSet, 1
            Do
                p23 = .BufPos
                q23 = .ThunkPos
                Call Parse_
                pvPushThunk ucsActResultClear
                lCaptureBegin = .BufPos
                If .BufData(.BufPos) = 43 Then      ' "+"
                    .BufPos = .BufPos + 1
                Else
                    If .BufData(.BufPos) = 45 Then  ' "-"
                        .BufPos = .BufPos + 1
                    Else
                        .BufPos = p23
                        .ThunkPos = q23
                        Exit Do
                    End If
                End If
                lCaptureEnd = .BufPos
                pvPushThunk ucsActResultSet, lCaptureBegin, lCaptureEnd
                pvPushThunk ucsActVarSet, 2
                Call Parse_
                pvPushThunk ucsActResultClear
                If ParseTerm() Then
                    pvPushThunk ucsActVarSet, 3
                Else
                    .BufPos = p23
                    .ThunkPos = q23
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
            .BufPos = p35
            .ThunkPos = q35
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Expression", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseTerm() As Boolean
    Dim p65 As Long
    Dim q65 As Long
    Dim p53 As Long
    Dim q53 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 3
        p65 = .BufPos
        q65 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Term", 1, ctx.UserData)
        pvPushThunk ucsActResultClear
        If ParseFactor() Then
            pvPushThunk ucsActVarSet, 1
            Do
                p53 = .BufPos
                q53 = .ThunkPos
                Call Parse_
                pvPushThunk ucsActResultClear
                lCaptureBegin = .BufPos
                If .BufData(.BufPos) = 42 Then      ' "*"i
                    .BufPos = .BufPos + 1
                Else
                    If .BufData(.BufPos) = 47 Then  ' "/"
                        .BufPos = .BufPos + 1
                    Else
                        .BufPos = p53
                        .ThunkPos = q53
                        Exit Do
                    End If
                End If
                lCaptureEnd = .BufPos
                pvPushThunk ucsActResultSet, lCaptureBegin, lCaptureEnd
                pvPushThunk ucsActVarSet, 2
                Call Parse_
                pvPushThunk ucsActResultClear
                If ParseFactor() Then
                    pvPushThunk ucsActVarSet, 3
                Else
                    .BufPos = p53
                    .ThunkPos = q53
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
            .BufPos = p65
            .ThunkPos = q65
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Term", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Private Sub Parse_()
    With ctx
        Call ConsoleTrace(ctx.BufPos + 1, "_", 1, ctx.UserData)
        Do
            Select Case .BufData(.BufPos)
            Case 32, 9, 10, 13                      ' [ \t\n\r]i
                .BufPos = .BufPos + 1
            Case Else
                Exit Do
            End Select
        Loop
        Call ConsoleTrace(ctx.BufPos + 1, "_", 2, ctx.UserData)
        Call pvSetAdvance
    End With
End Sub

Public Function ParseFactor() As Boolean
    Dim p86 As Long
    Dim q86 As Long
    Dim p82 As Long
    Dim q82 As Long
    Dim p79 As Long
    Dim q79 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        p86 = .BufPos
        q86 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Factor", 1, ctx.UserData)
        p82 = .BufPos
        q82 = .ThunkPos
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call Parse_
            pvPushThunk ucsActResultClear
            If ParseExpression() Then
                pvPushThunk ucsActVarSet, 1
                Call Parse_
                If .BufData(.BufPos) = 41 Then      ' ")"
                    .BufPos = .BufPos + 1
                    pvPushThunk ucsAct_1_Factor, .CaptureBegin, .CaptureEnd
                    Call ConsoleTrace(ctx.BufPos + 1, "Factor", 2, ctx.UserData)
                    pvPushThunk ucsActVarAlloc, -1
                    Call pvSetAdvance
                    ParseFactor = True
                    Exit Function
                Else
                    .BufPos = p82
                    .ThunkPos = q82
                End If
            Else
                .BufPos = p82
                .ThunkPos = q82
            End If
        End If
        p79 = .BufPos
        q79 = .ThunkPos
        If ParseInteger() Then
            Call ConsoleTrace(ctx.BufPos + 1, "Factor", 2, ctx.UserData)
            pvPushThunk ucsActVarAlloc, -1
            Call pvSetAdvance
            ParseFactor = True
            Exit Function
        Else
            .BufPos = p79
            .ThunkPos = q79
        End If
        If ParseErrInteger() Then
            Call ConsoleTrace(ctx.BufPos + 1, "Factor", 2, ctx.UserData)
            pvPushThunk ucsActVarAlloc, -1
            Call pvSetAdvance
            ParseFactor = True
            Exit Function
        Else
            .BufPos = p79
            .ThunkPos = q79
        End If
        .BufPos = p82
        .ThunkPos = q82
        .BufPos = p86
        .ThunkPos = q86
        If ConsoleTrace(ctx.BufPos + 1, "Factor", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseInteger() As Boolean
    Dim p98 As Long
    Dim q98 As Long
    Dim lCaptureBegin As Long
    Dim i92 As Long
    Dim lCaptureEnd As Long

    With ctx
        p98 = .BufPos
        q98 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "Integer", 1, ctx.UserData)
        Call Parse_
        lCaptureBegin = .BufPos
        For i92 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i92 <> 0 Then
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_Integer, lCaptureBegin, lCaptureEnd
            Call ConsoleTrace(ctx.BufPos + 1, "Integer", 2, ctx.UserData)
            Call pvSetAdvance
            ParseInteger = True
            Exit Function
        Else
            .BufPos = p98
            .ThunkPos = q98
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Integer", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        Call pvSetExpected("integer")
    End With
End Function

Public Function ParseErrInteger() As Boolean
    Dim p110 As Long
    Dim q110 As Long
    Dim i104 As Long
    Dim p103 As Long
    Dim q103 As Long
    Dim p101 As Long
    Dim q101 As Long
    Dim e101 As String

    With ctx
        p110 = .BufPos
        q110 = .ThunkPos
        Call ConsoleTrace(ctx.BufPos + 1, "ErrInteger", 1, ctx.UserData)
        For i104 = 0 To LNG_MAXINT
            p103 = .BufPos
            q103 = .ThunkPos
            p101 = .BufPos
            q101 = .ThunkPos
            e101 = .LastExpected
            If ParseInteger() Then
                .BufPos = p103
                .ThunkPos = q103
                Exit For
            Else
                .BufPos = p101
                .ThunkPos = q101
                .LastExpected = e101
            End If
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
            Else
                .BufPos = p103
                .ThunkPos = q103
                Exit For
            End If
        Next
        If i104 <> 0 Then
            If ParseInteger() Then
                Call ConsoleTrace(ctx.BufPos + 1, "ErrInteger", 2, ctx.UserData)
                ParseErrInteger = True
                Exit Function
            Else
                .BufPos = p110
                .ThunkPos = q110
            End If
        Else
            .BufPos = p110
            .ThunkPos = q110
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
