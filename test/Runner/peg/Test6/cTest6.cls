VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTest6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Auto-generated on 25.7.2018 18:55:46
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
    ParserVersion = "25.7.2018 18:55:46"
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Match(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant, Optional Result As Variant) As Long
    If BeginMatch(sSubject, StartPos, UserData) Then
        If Parsestart() Then
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
    With ctx
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

Public Function Parsestart() As Boolean
    Dim p25 As Long
    Dim i14 As Long
    Dim p13 As Long
    Dim p5 As Long
    Dim e5 As String
    Dim p11 As Long
    Dim p19 As Long

    With ctx
        p25 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "start", 1, ctx.UserData)
        For i14 = 0 To LNG_MAXINT
            p13 = .BufPos
            p5 = .BufPos
            e5 = .LastExpected
            If ParseEndline() Then
                .BufPos = p13
                If ParseEndline() Then
                    p11 = .BufPos
                    If ParseInline() Then
                        .BufPos = p11
                    Else
                        .BufPos = p13
                        .BufPos = p13
                        Exit For
                    End If
                Else
                    .BufPos = p13
                    .BufPos = p13
                    Exit For
                End If
            Else
                .BufPos = p5
                .LastExpected = e5
                If Not (ParseInline()) Then
                    .BufPos = p13
                    If ParseEndline() Then
                        p11 = .BufPos
                        If ParseInline() Then
                            .BufPos = p11
                        Else
                            .BufPos = p13
                            .BufPos = p13
                            Exit For
                        End If
                    Else
                        .BufPos = p13
                        .BufPos = p13
                        Exit For
                    End If
                End If
            End If
        Next
        If i14 <> 0 Then
            p19 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, """!""", 0, ctx.UserData)
            If .BufData(.BufPos) = 33 Then          ' "!"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, """!""", 2, ctx.UserData)
                Call ConsoleTrace(ctx.BufPos + 1, "start", 2, ctx.UserData)
                Call pvSetAdvance
                Parsestart = True
                Exit Function
            End If
            If ConsoleTrace(ctx.BufPos + 1, """!""", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p25
        Else
            .BufPos = p25
        End If
        If ConsoleTrace(ctx.BufPos + 1, "start", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
    End With
End Function

Public Function ParseEndline() As Boolean
    Dim p36 As Long
    Dim p31 As Long

    With ctx
        p36 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Endline", 1, ctx.UserData)
        p31 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "[ \r\n]", 0, ctx.UserData)
        Select Case .BufData(.BufPos)
        Case 32, 13, 10                             ' [ \r\n]
            .BufPos = .BufPos + 1
            Call ConsoleTrace(ctx.BufPos + 1, "[ \r\n]", 2, ctx.UserData)
            Call ConsoleTrace(ctx.BufPos + 1, "Endline", 2, ctx.UserData)
            Call pvSetAdvance
            ParseEndline = True
            Exit Function
        End Select
        If ConsoleTrace(ctx.BufPos + 1, "[ \r\n]", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        .BufPos = p36
        If ConsoleTrace(ctx.BufPos + 1, "Endline", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        Call pvSetExpected("EndLine")
    End With
End Function

Public Function ParseInline() As Boolean
    Dim p48 As Long
    Dim i44 As Long
    Dim p42 As Long

    With ctx
        p48 = .BufPos
        Call ConsoleTrace(ctx.BufPos + 1, "Inline", 1, ctx.UserData)
        For i44 = 0 To LNG_MAXINT
            p42 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, "[a-z]", 0, ctx.UserData)
            Select Case .BufData(.BufPos)
            Case 97 To 122                          ' [a-z]
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, "[a-z]", 2, ctx.UserData)
            Case Else
                If Not (ConsoleTrace(ctx.BufPos + 1, "[a-z]", 3, ctx.UserData) And False) Then
                    .BufPos = p42
                    Exit For
                End If
            End Select
        Next
        If i44 <> 0 Then
            Call ConsoleTrace(ctx.BufPos + 1, "Inline", 2, ctx.UserData)
            Call pvSetAdvance
            ParseInline = True
            Exit Function
        Else
            .BufPos = p48
        End If
        If ConsoleTrace(ctx.BufPos + 1, "Inline", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        Call pvSetExpected("Inline")
    End With
End Function

