Attribute VB_Name = "Module1"
' Auto-generated on 11.7.2018 18:49:51
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
    LastError           As String
    UserData            As Variant
    VarResult           As Variant
    VarStack()          As Variant
    VarPos              As Long
End Type

Private ctx                     As UcsParserType

'=========================================================================
' Properties
'=========================================================================

Property Get VbPegLastError() As String
    VbPegLastError = ctx.LastError
End Property

Property Get VbPegParserVersion() As String
    VbPegParserVersion = "11.7.2018 18:49:51"
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function VbPegMatch(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant, Optional Result As Variant) As Long
    If VbPegBeginMatch(sSubject, StartPos, UserData) Then
        If VbPegParsestart() Then
            VbPegMatch = VbPegEndMatch(UserData, Result)
        End If
    End If
End Function

Public Function VbPegBeginMatch(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant) As Boolean
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
    VbPegBeginMatch = True
End Function

Public Function VbPegEndMatch(Optional UserData As Variant, Optional Result As Variant) As Long
    Dim lIdx            As Long
    Dim uEmpty          As UcsParserType
    
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
        If IsObject(.UserData) Then
            Set UserData = .UserData
        Else
            UserData = .UserData
        End If
        VbPegEndMatch = .BufPos + 1
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

Public Function VbPegParsestart() As Boolean
    If VbPegParseadditive() Then
        VbPegParsestart = True
    End If
End Function

Public Function VbPegParseadditive() As Boolean
    Dim p16 As Long
    Dim q16 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        p16 = .BufPos
        q16 = .ThunkPos
        pvPushThunk ucsActResultClear
        If VbPegParsemultiplicative() Then
            pvPushThunk ucsActVarSet, 1
            If .BufData(.BufPos) = 43 Then          ' "+"
                .BufPos = .BufPos + 1
                pvPushThunk ucsActResultClear
                If VbPegParseadditive() Then
                    pvPushThunk ucsActVarSet, 2
                    pvPushAction ucsAct_1_additive
                    pvPushThunk ucsActVarAlloc, -2
                    VbPegParseadditive = True
                    Exit Function
                Else
                    .BufPos = p16
                    .ThunkPos = q16
                End If
            Else
                .BufPos = p16
                .ThunkPos = q16
            End If
        End If
        If VbPegParsemultiplicative() Then
            pvPushThunk ucsActVarAlloc, -2
            VbPegParseadditive = True
            Exit Function
        Else
            .BufPos = p16
            .ThunkPos = q16
        End If
    End With
End Function

Public Function VbPegParsemultiplicative() As Boolean
    Dim p29 As Long
    Dim q29 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        p29 = .BufPos
        q29 = .ThunkPos
        pvPushThunk ucsActResultClear
        If VbPegParseprimary() Then
            pvPushThunk ucsActVarSet, 1
            If .BufData(.BufPos) = 42 Then          ' "*"
                .BufPos = .BufPos + 1
                pvPushThunk ucsActResultClear
                If VbPegParsemultiplicative() Then
                    pvPushThunk ucsActVarSet, 2
                    pvPushAction ucsAct_1_multiplicative
                    pvPushThunk ucsActVarAlloc, -2
                    VbPegParsemultiplicative = True
                    Exit Function
                Else
                    .BufPos = p29
                    .ThunkPos = q29
                End If
            Else
                .BufPos = p29
                .ThunkPos = q29
            End If
        End If
        If VbPegParseprimary() Then
            pvPushThunk ucsActVarAlloc, -2
            VbPegParsemultiplicative = True
            Exit Function
        Else
            .BufPos = p29
            .ThunkPos = q29
        End If
    End With
End Function

Public Function VbPegParseprimary() As Boolean
    Dim p40 As Long
    Dim q40 As Long
    Dim p37 As Long
    Dim q37 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        p40 = .BufPos
        q40 = .ThunkPos
        If VbPegParseinteger() Then
            pvPushThunk ucsActVarAlloc, -1
            VbPegParseprimary = True
            Exit Function
        Else
            .BufPos = p40
            .ThunkPos = q40
        End If
        p37 = .BufPos
        q37 = .ThunkPos
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            pvPushThunk ucsActResultClear
            If VbPegParseadditive() Then
                pvPushThunk ucsActVarSet, 1
                If .BufData(.BufPos) = 41 Then      ' ")"
                    .BufPos = .BufPos + 1
                    pvPushAction ucsAct_1_primary
                    pvPushThunk ucsActVarAlloc, -1
                    VbPegParseprimary = True
                    Exit Function
                Else
                    .BufPos = p37
                    .ThunkPos = q37
                End If
            Else
                .BufPos = p37
                .ThunkPos = q37
            End If
        End If
    End With
End Function

Public Function VbPegParseinteger() As Boolean
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
            pvPushAction ucsAct_1_integer
            pvPushThunk ucsActVarAlloc, -1
            VbPegParseinteger = True
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

