VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTest10"
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

Private Enum UcsParserActionsEnum
    ucsAct_4_block
    ucsAct_3_block
    ucsAct_2_block
    ucsAct_1_block
    ucsAct_1_IDENT
    ucsAct_1_LINE
    ucsAct_1_EMPTY
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
        If Parseblock() Then
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

'= generated functions ===================================================

Public Function Parseblock(Optional ByVal Level As Long) As Boolean
    Dim i54 As Long
    Dim p53 As Long
    Dim q53 As Long
    Dim p42 As Long
    Dim q42 As Long
    Dim p41 As Long
    Dim q41 As Long
    Dim p31 As Long
    Dim q31 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 3
        pvPushThunk ucsActResultClear
        pvPushThunk ucsActVarSet, 1
        pvPushThunk ucsAct_1_block, .CaptureBegin, .CaptureEnd
        For i54 = 0 To LNG_MAXINT
            p53 = .BufPos
            q53 = .ThunkPos
            Call ParseSPACE
            If Not (.CaptureEnd - .CaptureBegin = Level) Then
                .BufPos = p53
                .ThunkPos = q53
                GoTo L4
            End If
            p42 = .BufPos
            q42 = .ThunkPos
            pvPushThunk ucsActResultClear
            If ParseIDENT() Then
                pvPushThunk ucsActVarSet, 2
            Else
                .BufPos = p42
                .ThunkPos = q42
                GoTo L10
            End If
            If Not (ParseEQ()) Then
                .BufPos = p42
                .ThunkPos = q42
                GoTo L10
            End If
            pvPushThunk ucsActResultClear
            If ParseLINE() Then
                pvPushThunk ucsActVarSet, 3
            Else
                .BufPos = p42
                .ThunkPos = q42
                GoTo L10
            End If
            Do
                Select Case .BufData(.BufPos)
                Case 13, 10                         ' [\r\n]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit Do
                End Select
            Loop
            GoTo L5
L10:
            pvPushThunk ucsActResultClear
            If ParseLINE() Then
                pvPushThunk ucsActVarSet, 2
            Else
                .BufPos = p42
                .ThunkPos = q42
                GoTo L6
            End If
            Do
                Select Case .BufData(.BufPos)
                Case 13, 10                         ' [\r\n]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit Do
                End Select
            Loop
            p41 = .BufPos
            q41 = .ThunkPos
            p31 = .BufPos
            q31 = .ThunkPos
            Call ParseSPACE
            .BufPos = p31
            .ThunkPos = q31
            If Not (.CaptureEnd - .CaptureBegin > Level) Then
                .BufPos = p41
                .ThunkPos = q41
                GoTo L9
            End If
            pvPushThunk ucsActResultClear
            If Parseblock(.CaptureEnd - .CaptureBegin) Then
                pvPushThunk ucsActVarSet, 3
            Else
                .BufPos = p41
                .ThunkPos = q41
                GoTo L9
            End If
            GoTo L7
L9:
            pvPushThunk ucsActResultClear
            Call ParseEMPTY
            pvPushThunk ucsActVarSet, 3
            GoTo L7
            .BufPos = p42
            .ThunkPos = q42
            GoTo L6
L7:
            GoTo L5
L6:
            .BufPos = p53
            .ThunkPos = q53
            GoTo L4
L5:
            pvPushThunk ucsAct_2_block, .CaptureBegin, .CaptureEnd
            GoTo L2
L4:
            If Not (.CaptureEnd - .CaptureBegin > Level) Then
                GoTo L3
            End If
            Call ParseSPACE
            pvPushThunk ucsActResultClear
            If ParseLINE() Then
                pvPushThunk ucsActVarSet, 2
            Else
                .BufPos = p53
                .ThunkPos = q53
                GoTo L3
            End If
            Do
                Select Case .BufData(.BufPos)
                Case 13, 10                         ' [\r\n]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit Do
                End Select
            Loop
            pvPushThunk ucsAct_3_block, .CaptureBegin, .CaptureEnd
            GoTo L2
L3:
            .BufPos = p53
            .ThunkPos = q53
            Exit For
L2:
        Next
        If i54 <> 0 Then
            pvPushThunk ucsAct_4_block, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -3
            Call pvSetAdvance
            Parseblock = True
        End If
    End With
End Function

Private Sub ParseSPACE()
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        Do
            If .BufData(.BufPos) = 32 Then          ' " "
                .BufPos = .BufPos + 1
            Else
                Exit Do
            End If
        Loop
        lCaptureEnd = .BufPos
        .CaptureBegin = lCaptureBegin
        .CaptureEnd = lCaptureEnd
        Call pvSetAdvance
    End With
End Sub

Private Function ParseIDENT() As Boolean
    Dim lCaptureBegin As Long
    Dim i74 As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        For i74 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90                ' [a-zA-Z]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i74 <> 0 Then
            lCaptureEnd = .BufPos
            Call Parse_
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_IDENT, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseIDENT = True
        End If
    End With
End Function

Private Function ParseEQ() As Boolean
    With ctx
        If .BufData(.BufPos) = 61 Then              ' "="
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseEQ = True
        End If
    End With
End Function

Private Function ParseLINE() As Boolean
    Dim lCaptureBegin As Long
    Dim i68 As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        For i68 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 13, 10                             ' [\r\n]
                Exit For
            Case Else
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    Exit For
                End If
            End Select
        Next
        If i68 <> 0 Then
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_LINE, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseLINE = True
        End If
    End With
End Function

Private Sub ParseEMPTY()
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        lCaptureEnd = .BufPos
        .CaptureBegin = lCaptureBegin
        .CaptureEnd = lCaptureEnd
        pvPushThunk ucsAct_1_EMPTY, lCaptureBegin, lCaptureEnd
    End With
End Sub

Private Sub Parse_()
    With ctx
        Do
            Select Case .BufData(.BufPos)
            Case 32, 9                              ' [ \t]
                .BufPos = .BufPos + 1
            Case Else
                Exit Do
            End Select
        Loop
        Call pvSetAdvance
    End With
End Sub

Private Sub pvImplAction(ByVal eAction As UcsParserActionsEnum, ByVal lOffset As Long, ByVal lSize As Long)
    Dim oJson As Object

    With ctx
    Select Case eAction
    Case ucsAct_4_block
         Set .VarResult = .VarStack(.VarPos - 1)
    Case ucsAct_3_block
         Set oJson = .VarStack(.VarPos - 1) : JsonItem(oJson, .VarStack(.VarPos - 2)) = Array("Error baddent")
    Case ucsAct_2_block
         Set oJson = .VarStack(.VarPos - 1) : JsonItem(oJson, .VarStack(.VarPos - 2)) = .VarStack(.VarPos - 3)
    Case ucsAct_1_block
         JsonItem(oJson, vbNullString) = Empty : Set .VarStack(.VarPos - 1) = oJson
    Case ucsAct_1_IDENT
         .VarResult = Mid$(.Contents, lOffset, lSize)
    Case ucsAct_1_LINE
         .VarResult = Mid$(.Contents, lOffset, lSize)
    Case ucsAct_1_EMPTY
         .VarResult = Mid$(.Contents, lOffset, lSize)
    End Select
    End With
End Sub

