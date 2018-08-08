VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTest3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Auto-generated on 27.7.2018 19:38:23
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
    ParserVersion = "27.7.2018 19:38:23"
End Property

Property Get Contents(Optional ByVal lOffset As Long = 1, Optional ByVal lSize As Long = LNG_MAXINT) As String
    Contents = Mid$(ctx.Contents, lOffset, lSize)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Match(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant, Optional Result As Variant) As Long
    If BeginMatch(sSubject, StartPos, UserData) Then
        If Parselist() Then
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

'= generated functions ===================================================

Public Function Parselist() As Boolean
    Dim p5 As Long

    With ctx
        If Parseopen() Then
            Do
                p5 = .BufPos
                If Not (Parseelem()) Then
                    .BufPos = p5
                    Exit Do
                End If
            Loop
            If Parseclose() Then
                Parselist = True
            End If
        End If
    End With
End Function

Public Function Parseopen() As Boolean
    Dim p39 As Long

    With ctx
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Do
                p39 = .BufPos
                If Not (Parsespace()) Then
                    .BufPos = p39
                    Exit Do
                End If
            Loop
            Call pvSetAdvance
            Parseopen = True
        End If
    End With
End Function

Public Function Parseelem() As Boolean
    Dim p13 As Long

    With ctx
        p13 = .BufPos
        If Parselist() Then
            Parseelem = True
            Exit Function
        Else
            .BufPos = p13
        End If
        If Parseatom() Then
            Parseelem = True
            Exit Function
        Else
            .BufPos = p13
        End If
        If Parsesstring() Then
            Parseelem = True
            Exit Function
        Else
            .BufPos = p13
        End If
        If Parsedstring() Then
            Parseelem = True
            Exit Function
        Else
            .BufPos = p13
        End If
    End With
End Function

Public Function Parseclose() As Boolean
    Dim p43 As Long

    With ctx
        If .BufData(.BufPos) = 41 Then              ' ")"
            .BufPos = .BufPos + 1
            Do
                p43 = .BufPos
                If Not (Parsespace()) Then
                    .BufPos = p43
                    Exit Do
                End If
            Loop
            Call pvSetAdvance
            Parseclose = True
        End If
    End With
End Function

Public Function Parseatom() As Boolean
    Dim i19 As Long
    Dim p21 As Long

    With ctx
        For i19 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 97 To 122, 48 To 57, 95            ' [a-z0-9_]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i19 <> 0 Then
            Do
                p21 = .BufPos
                If Not (Parsespace()) Then
                    .BufPos = p21
                    Exit Do
                End If
            Loop
            Call pvSetAdvance
            Parseatom = True
        End If
    End With
End Function

Public Function Parsesstring() As Boolean
    Dim p36 As Long

    With ctx
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            Do
                If .BufData(.BufPos) <> 39 And .BufPos < .BufSize Then ' "'"
                    .BufPos = .BufPos + 1
                Else
                    Exit Do
                End If
            Loop
            If .BufData(.BufPos) = 39 Then          ' "'"
                .BufPos = .BufPos + 1
                Do
                    p36 = .BufPos
                    If Not (Parsespace()) Then
                        .BufPos = p36
                        Exit Do
                    End If
                Loop
                Call pvSetAdvance
                Parsesstring = True
            End If
        End If
    End With
End Function

Public Function Parsedstring() As Boolean
    Dim p29 As Long

    With ctx
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            Do
                If .BufData(.BufPos) <> 34 And .BufPos < .BufSize Then ' """
                    .BufPos = .BufPos + 1
                Else
                    Exit Do
                End If
            Loop
            If .BufData(.BufPos) = 34 Then          ' """
                .BufPos = .BufPos + 1
                Do
                    p29 = .BufPos
                    If Not (Parsespace()) Then
                        .BufPos = p29
                        Exit Do
                    End If
                Loop
                Call pvSetAdvance
                Parsedstring = True
            End If
        End If
    End With
End Function

Public Function Parsespace() As Boolean
    Dim p48 As Long

    With ctx
        p48 = .BufPos
        If .BufData(.BufPos) = 32 Then              ' " "
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            Parsespace = True
            Exit Function
        End If
        If .BufData(.BufPos) = 9 Then               ' "\t"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            Parsespace = True
            Exit Function
        End If
        If Parseeol() Then
            Call pvSetAdvance
            Parsespace = True
            Exit Function
        Else
            .BufPos = p48
        End If
    End With
End Function

Public Function Parseeol() As Boolean
    With ctx
        If .BufData(.BufPos) = 13 And .BufData(.BufPos + 1) = 10 Then ' "\r\n"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            Parseeol = True
            Exit Function
        End If
        If .BufData(.BufPos) = 10 Then              ' "\n"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            Parseeol = True
            Exit Function
        End If
        If .BufData(.BufPos) = 13 Then              ' "\r"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            Parseeol = True
        End If
    End With
End Function

