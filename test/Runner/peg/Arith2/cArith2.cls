VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArith2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Auto-generated on 5.3.2019 1:30:17
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

Private Enum UcsParserActionsEnum
    ucsAct_1_additive
    ucsAct_2_additive
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
    ParserVersion = "5.3.2019 1:30:17"
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

Private Function Parseadditive() As Boolean
    Dim p23 As Long
    Dim q23 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        p23 = .BufPos
        q23 = .ThunkPos
        pvPushThunk ucsActResultClear
        If Parsemultiplicative() Then
            pvPushThunk ucsActVarSet, 1
            If .BufData(.BufPos) = 43 Then          ' "+"
                .BufPos = .BufPos + 1
                pvPushThunk ucsActResultClear
                If Parseadditive() Then
                    pvPushThunk ucsActVarSet, 2
                    pvPushThunk ucsAct_1_additive, .CaptureBegin, .CaptureEnd
                    GoTo L1
                Else
                    .BufPos = p23
                    .ThunkPos = q23
                End If
            Else
                .BufPos = p23
                .ThunkPos = q23
            End If
        Else
            .BufPos = p23
            .ThunkPos = q23
        End If
        pvPushThunk ucsActResultClear
        If Parsemultiplicative() Then
            pvPushThunk ucsActVarSet, 1
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                pvPushThunk ucsActResultClear
                If Parseadditive() Then
                    pvPushThunk ucsActVarSet, 2
                    pvPushThunk ucsAct_2_additive, .CaptureBegin, .CaptureEnd
                    GoTo L1
                Else
                    .BufPos = p23
                    .ThunkPos = q23
                End If
            Else
                .BufPos = p23
                .ThunkPos = q23
            End If
        Else
            .BufPos = p23
            .ThunkPos = q23
        End If
        If Parsemultiplicative() Then
            GoTo L1
        Else
            .BufPos = p23
            .ThunkPos = q23
        End If
        If False Then
L1:
            pvPushThunk ucsActVarAlloc, -2
            Call pvSetAdvance
            Parseadditive = True
        End If
    End With
End Function

Private Function Parsemultiplicative() As Boolean
    Dim p37 As Long
    Dim q37 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        p37 = .BufPos
        q37 = .ThunkPos
        pvPushThunk ucsActResultClear
        If Parseprimary() Then
            pvPushThunk ucsActVarSet, 1
            If .BufData(.BufPos) = 42 Then          ' "*"
                .BufPos = .BufPos + 1
                pvPushThunk ucsActResultClear
                If Parsemultiplicative() Then
                    pvPushThunk ucsActVarSet, 2
                    pvPushThunk ucsAct_1_multiplicative, .CaptureBegin, .CaptureEnd
                    GoTo L4
                Else
                    .BufPos = p37
                    .ThunkPos = q37
                End If
            Else
                .BufPos = p37
                .ThunkPos = q37
            End If
        Else
            .BufPos = p37
            .ThunkPos = q37
        End If
        If Parseprimary() Then
            GoTo L4
        Else
            .BufPos = p37
            .ThunkPos = q37
        End If
        If False Then
L4:
            pvPushThunk ucsActVarAlloc, -2
            Call pvSetAdvance
            Parsemultiplicative = True
        End If
    End With
End Function

Private Function Parseprimary() As Boolean
    Dim p48 As Long
    Dim q48 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        p48 = .BufPos
        q48 = .ThunkPos
        If Parseinteger() Then
            pvPushThunk ucsActVarAlloc, -1
            Call pvSetAdvance
            Parseprimary = True
            Exit Function
        Else
            .BufPos = p48
            .ThunkPos = q48
            If .BufData(.BufPos) = 40 Then          ' "("
                .BufPos = .BufPos + 1
            Else
                GoTo L7
            End If
            pvPushThunk ucsActResultClear
            If Parseadditive() Then
                pvPushThunk ucsActVarSet, 1
            Else
                .BufPos = p48
                .ThunkPos = q48
                GoTo L7
            End If
            If .BufData(.BufPos) = 41 Then          ' ")"
                .BufPos = .BufPos + 1
            Else
                .BufPos = p48
                .ThunkPos = q48
                GoTo L7
            End If
            pvPushThunk ucsAct_1_primary, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -1
            Call pvSetAdvance
            Parseprimary = True
            Exit Function
L7:
            Call pvSetExpected("primary")
        End If
    End With
End Function

Private Function Parseinteger() As Boolean
    Dim lCaptureBegin As Long
    Dim i51 As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        For i51 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i51 <> 0 Then
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_integer, lCaptureBegin, lCaptureEnd
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
        Case ucsAct_2_additive
             .VarResult = .VarStack(.VarPos - 1) - .VarStack(.VarPos - 2)
        Case ucsAct_1_multiplicative
             .VarResult = .VarStack(.VarPos - 1) * .VarStack(.VarPos - 2)
        Case ucsAct_1_primary
             .VarResult = .VarStack(.VarPos - 1)
        Case ucsAct_1_integer
             .VarResult = CDbl(Mid$(.Contents, lOffset, lSize))
        End Select
    End With
End Sub

