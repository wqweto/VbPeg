VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTest5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Auto-generated on 3.12.2018 1:29:29
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

Private Type SAFEARRAY1D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As Long
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
    ParserVersion = "3.12.2018 1:29:29"
End Property

Property Get Contents(Optional ByVal lOffset As Long = 1, Optional ByVal lSize As Long = LNG_MAXINT) As String
    Contents = Mid$(ctx.Contents, lOffset, lSize)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Match(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant, Optional Result As Variant) As Long
    If BeginMatch(sSubject, StartPos, UserData) Then
        If ParseStart() Then
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
        Call CopyMemory(ByVal VarPtr(.Contents), ByVal VarPtr(sSubject), 4)
        With .BufArray
            .cDims = 1
            .cbElements = 2
            .fFeatures = 1 ' FADF_AUTO
            .pvData = StrPtr(sSubject)
            .cElements = Len(sSubject) + 2 '-- look-ahead chars
        End With
        Call CopyMemory(ByVal ArrPtr(.BufData), VarPtr(.BufArray), 4)
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
        Call CopyMemory(ByVal VarPtr(.Contents), 0&, 4)
        Call CopyMemory(ByVal ArrPtr(.BufData), 0&, 4)
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

Public Function ParseStart() As Boolean
    Dim i4 As Long
    Dim p3 As Long

    With ctx
        For i4 = 0 To LNG_MAXINT
            p3 = .BufPos
            If Not (ParseChar()) Then
                .BufPos = p3
                Exit For
            End If
        Next
        If i4 <> 0 Then
            If ParseEnd() Then
                ParseStart = True
            End If
        End If
    End With
End Function

Private Function ParseChar() As Boolean
    Dim p10 As Long
    Dim e10 As String

    With ctx
        p10 = .BufPos
        e10 = .LastExpected
        If ParseEnd() Then
            Call pvSetExpected("[a-z]")
        Else
            .BufPos = p10
            .LastExpected = e10
            Select Case .BufData(.BufPos)
            Case 97 To 122                          ' [a-z]
                .BufPos = .BufPos + 1
                Call pvSetAdvance
                ParseChar = True
                Exit Function
            Case Else
                Call pvSetExpected("[a-z]")
            End Select
        End If
    End With
End Function

Private Function ParseEnd() As Boolean
    With ctx
        If .BufData(.BufPos) = 101 Then             ' "e"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseEnd = True
            Exit Function
        Else
            Call pvSetExpected("'e'")
        End If
    End With
End Function

