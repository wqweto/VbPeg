VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCalc3"
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
    ucsAct_1_additive
    ucsAct_1_multiplicative
    ucsAct_1_primary
    ucsAct_1_integer
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

Public Function Parsestart() As Boolean
    If Parseadditive() Then
        Parsestart = True
    End If
End Function

Public Function Parseadditive() As Boolean
    Dim p16 As Long
    Dim q16 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        p16 = .BufPos
        q16 = .ThunkPos
        pvPushThunk ucsActResultClear
        If Parsemultiplicative() Then
            pvPushThunk ucsActVarSet, 1
            If .BufData(.BufPos) = 43 Then          ' "+"
                .BufPos = .BufPos + 1
                pvPushThunk ucsActResultClear
                If Parseadditive() Then
                    pvPushThunk ucsActVarSet, 2
                    pvPushThunk ucsAct_1_additive, .CaptureBegin, .CaptureEnd
                    pvPushThunk ucsActVarAlloc, -2
                    Call pvSetAdvance
                    Parseadditive = True
                    Exit Function
                Else
                    .BufPos = p16
                    .ThunkPos = q16
                End If
            Else
                .BufPos = p16
                .ThunkPos = q16
            End If
        Else
            .BufPos = p16
            .ThunkPos = q16
        End If
        If Parsemultiplicative() Then
            pvPushThunk ucsActVarAlloc, -2
            Call pvSetAdvance
            Parseadditive = True
            Exit Function
        Else
            .BufPos = p16
            .ThunkPos = q16
        End If
    End With
End Function

Public Function Parsemultiplicative() As Boolean
    Dim p29 As Long
    Dim q29 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        p29 = .BufPos
        q29 = .ThunkPos
        pvPushThunk ucsActResultClear
        If Parseprimary() Then
            pvPushThunk ucsActVarSet, 1
            If .BufData(.BufPos) = 42 Then          ' "*"
                .BufPos = .BufPos + 1
                pvPushThunk ucsActResultClear
                If Parsemultiplicative() Then
                    pvPushThunk ucsActVarSet, 2
                    pvPushThunk ucsAct_1_multiplicative, .CaptureBegin, .CaptureEnd
                    pvPushThunk ucsActVarAlloc, -2
                    Call pvSetAdvance
                    Parsemultiplicative = True
                    Exit Function
                Else
                    .BufPos = p29
                    .ThunkPos = q29
                End If
            Else
                .BufPos = p29
                .ThunkPos = q29
            End If
        Else
            .BufPos = p29
            .ThunkPos = q29
        End If
        If Parseprimary() Then
            pvPushThunk ucsActVarAlloc, -2
            Call pvSetAdvance
            Parsemultiplicative = True
            Exit Function
        Else
            .BufPos = p29
            .ThunkPos = q29
        End If
    End With
End Function

Public Function Parseprimary() As Boolean
    Dim p40 As Long
    Dim q40 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        p40 = .BufPos
        q40 = .ThunkPos
        If Parseinteger() Then
            pvPushThunk ucsActVarAlloc, -1
            Call pvSetAdvance
            Parseprimary = True
            Exit Function
        Else
            .BufPos = p40
            .ThunkPos = q40
        End If
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            pvPushThunk ucsActResultClear
            If Parseadditive() Then
                pvPushThunk ucsActVarSet, 1
                If .BufData(.BufPos) = 41 Then      ' ")"
                    .BufPos = .BufPos + 1
                    pvPushThunk ucsAct_1_primary, .CaptureBegin, .CaptureEnd
                    pvPushThunk ucsActVarAlloc, -1
                    Call pvSetAdvance
                    Parseprimary = True
                    Exit Function
                Else
                    .BufPos = p40
                    .ThunkPos = q40
                End If
            Else
                .BufPos = p40
                .ThunkPos = q40
            End If
        End If
    End With
End Function

Public Function Parseinteger() As Boolean
    Dim lCaptureBegin As Long
    Dim i44 As Long
    Dim lCaptureEnd As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        pvPushThunk ucsActResultClear
        lCaptureBegin = .BufPos
        For i44 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i44 <> 0 Then
            lCaptureEnd = .BufPos
            pvPushThunk ucsActResultSet, lCaptureBegin, lCaptureEnd
            pvPushThunk ucsActVarSet, 1
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_integer, lCaptureBegin, lCaptureEnd
            pvPushThunk ucsActVarAlloc, -1
            Call pvSetAdvance
            Parseinteger = True
            Exit Function
        Else
            Call pvSetExpected("integer")
        End If
    End With
End Function

Private Sub pvImplAction(ByVal eAction As UcsParserActionsEnum, ByVal lOffset As Long, ByVal lSize As Long)
    With ctx
        Select Case eAction
        Case ucsAct_1_additive
             .VarResult = .VarStack(.VarPos - 1) + .VarStack(.VarPos - 2)
        Case ucsAct_1_multiplicative
             .VarResult = .VarStack(.VarPos - 1) * .VarStack(.VarPos - 2)
        Case ucsAct_1_primary
             .VarResult = .VarStack(.VarPos - 1)
        Case ucsAct_1_integer
             .VarResult = CLng(.VarStack(.VarPos - 1))
        End Select
    End With
End Sub

