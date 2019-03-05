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
' Auto-generated on 5.3.2019 16:44:41
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function ArrPtr Lib "vbe7" Alias "VarPtr" (Ptr() As Any) As LongPtr
Private Const NULL_PTR                  As LongPtr = 0
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Const NULL_PTR                  As Long = 0
#End If
#If Win64 Then
Private Const PTR_SIZE                  As Long = 8
#Else
Private Const PTR_SIZE                  As Long = 4
#End If

Private Type SAFEARRAY1D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
#If VBA7 Then
    pvData              As LongPtr
#Else
    pvData              As Long
#End If
    cElements           As Long
    lLbound             As Long
End Type

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
    BufArray            As SAFEARRAY1D
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
    ParserVersion = "5.3.2019 16:44:41"
End Property

Property Get Contents(Optional ByVal lOffset As Long = 1, Optional ByVal lSize As Long = LNG_MAXINT) As String
    Contents = Mid$(ctx.Contents, lOffset, lSize)
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
        Call CopyMemory(ByVal VarPtr(.Contents), ByVal VarPtr(sSubject), PTR_SIZE)
        With .BufArray
            .cDims = 1
            .fFeatures = 1 ' FADF_AUTO
            .cbElements = 2
            .pvData = StrPtr(sSubject)
            .cElements = Len(sSubject) + 2 '-- look-ahead chars
        End With
        Call CopyMemory(ByVal ArrPtr(.BufData), VarPtr(.BufArray), PTR_SIZE)
        .BufPos = StartPos
        .BufSize = Len(sSubject)
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
        Call CopyMemory(ByVal VarPtr(.Contents), NULL_PTR, PTR_SIZE)
        Call CopyMemory(ByVal ArrPtr(.BufData), NULL_PTR, PTR_SIZE)
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
                GoTo L7
            End If
            .BufPos = p5
            .LastExpected = e5
            If Not (ParseInline()) Then
                .BufPos = p13
                GoTo L7
            End If
            GoTo L5
L7:
            If Not (ParseEndline()) Then
                .BufPos = p13
                GoTo L6
            End If
            p11 = .BufPos
            If ParseInline() Then
                .BufPos = p11
            Else
                .BufPos = p13
                GoTo L6
            End If
            GoTo L5
L6:
            .BufPos = p13
            Exit For
L5:
        Next
        If i14 <> 0 Then
            p19 = .BufPos
            Call ConsoleTrace(ctx.BufPos + 1, """!""", 0, ctx.UserData)
            If .BufData(.BufPos) = 33 Then          ' "!"
                .BufPos = .BufPos + 1
                Call ConsoleTrace(ctx.BufPos + 1, """!""", 2, ctx.UserData)
                GoTo L3
            End If
            If ConsoleTrace(ctx.BufPos + 1, """!""", 3, ctx.UserData) And False Then
                '--- do nothing
            End If
            .BufPos = p25
            GoTo L2
L3:
            Call ConsoleTrace(ctx.BufPos + 1, "start", 2, ctx.UserData)
            GoTo L1
        End If
        .BufPos = p25
L2:
        If ConsoleTrace(ctx.BufPos + 1, "start", 3, ctx.UserData) And False Then
            '--- do nothing
        End If
        Exit Function
L1:
        Call pvSetAdvance
        Parsestart = True
    End With
End Function

Private Function ParseEndline() As Boolean
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
        Case Else
            GoTo L11
        End Select
        Call ConsoleTrace(ctx.BufPos + 1, "[ \r\n]", 2, ctx.UserData)
        GoTo L10
L11:
        If Not (ConsoleTrace(ctx.BufPos + 1, "[ \r\n]", 3, ctx.UserData) And False) Then
            .BufPos = p36
            GoTo L9
        End If
L10:
        Call ConsoleTrace(ctx.BufPos + 1, "Endline", 2, ctx.UserData)
        Call pvSetAdvance
        ParseEndline = True
        Exit Function
L9:
        If Not (ConsoleTrace(ctx.BufPos + 1, "Endline", 3, ctx.UserData) And False) Then
            Call pvSetExpected("EndLine")
        End If
    End With
End Function

Private Function ParseInline() As Boolean
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
            Case Else
                GoTo L15
            End Select
            Call ConsoleTrace(ctx.BufPos + 1, "[a-z]", 2, ctx.UserData)
            GoTo L14
L15:
            If Not (ConsoleTrace(ctx.BufPos + 1, "[a-z]", 3, ctx.UserData) And False) Then
                .BufPos = p42
                Exit For
            End If
L14:
        Next
        If i44 = 0 Then
            .BufPos = p48
            GoTo L13
        End If
        Call ConsoleTrace(ctx.BufPos + 1, "Inline", 2, ctx.UserData)
        Call pvSetAdvance
        ParseInline = True
        Exit Function
L13:
        If Not (ConsoleTrace(ctx.BufPos + 1, "Inline", 3, ctx.UserData) And False) Then
            Call pvSetExpected("Inline")
        End If
    End With
End Function

