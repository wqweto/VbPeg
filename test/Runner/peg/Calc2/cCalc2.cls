VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCalc2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Auto-generated on 5.3.2019 16:44:39
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
    ucsAct_1_Stmt
    ucsAct_2_Stmt
    ucsAct_3_Sum
    ucsAct_2_Sum
    ucsAct_1_Sum
    ucsAct_3_Product
    ucsAct_2_Product
    ucsAct_1_Product
    ucsAct_1_Value
    ucsAct_2_Value
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
    ParserVersion = "5.3.2019 16:44:39"
End Property

Property Get Contents(Optional ByVal lOffset As Long = 1, Optional ByVal lSize As Long = LNG_MAXINT) As String
    Contents = Mid$(ctx.Contents, lOffset, lSize)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Match(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant, Optional Result As Variant) As Long
    If BeginMatch(sSubject, StartPos, UserData) Then
        If ParseStmt() Then
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

Public Function ParseStmt() As Boolean
    Dim p24 As Long
    Dim q24 As Long
    Dim lCaptureBegin As Long
    Dim i19 As Long
    Dim p18 As Long
    Dim q18 As Long
    Dim p16 As Long
    Dim q16 As Long
    Dim e16 As String
    Dim lCaptureEnd As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        Call Parse_
        pvPushThunk ucsActResultClear
        If ParseSum() Then
            pvPushThunk ucsActVarSet, 1
            p24 = .BufPos
            q24 = .ThunkPos
            If ParseEOL() Then
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                pvPushThunk ucsAct_1_Stmt, lCaptureBegin, lCaptureEnd
                GoTo L2
            End If
            .BufPos = p24
            .ThunkPos = q24
            lCaptureBegin = .BufPos
            For i19 = 0 To LNG_MAXINT
                p18 = .BufPos
                q18 = .ThunkPos
                p16 = .BufPos
                q16 = .ThunkPos
                e16 = .LastExpected
                If ParseEOL() Then
                    .BufPos = p18
                    .ThunkPos = q18
                    Exit For
                End If
                .BufPos = p16
                .ThunkPos = q16
                .LastExpected = e16
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p18
                    .ThunkPos = q18
                    Exit For
                End If
            Next
            If i19 <> 0 Then
                lCaptureEnd = .BufPos
                If ParseEOL() Then
                    .CaptureBegin = lCaptureBegin
                    .CaptureEnd = lCaptureEnd
                    pvPushThunk ucsAct_2_Stmt, lCaptureBegin, lCaptureEnd
                    GoTo L2
                End If
                .BufPos = p24
                .ThunkPos = q24
            Else
                .BufPos = p24
                .ThunkPos = q24
            End If
            Exit Function
L2:
            pvPushThunk ucsActVarAlloc, -1
            ParseStmt = True
        End If
    End With
End Function

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

Private Function ParseSum() As Boolean
    Dim p45 As Long
    Dim q45 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        If ParseProduct() Then
            pvPushThunk ucsActVarSet, 1
            Do
                p45 = .BufPos
                q45 = .ThunkPos
                If Not (ParsePLUS()) Then
                    .BufPos = p45
                    .ThunkPos = q45
                    GoTo L8
                End If
                pvPushThunk ucsActResultClear
                If ParseProduct() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p45
                    .ThunkPos = q45
                    GoTo L8
                End If
                pvPushThunk ucsAct_1_Sum, .CaptureBegin, .CaptureEnd
                GoTo L6
L8:
                If Not (ParseMINUS()) Then
                    .BufPos = p45
                    .ThunkPos = q45
                    GoTo L7
                End If
                pvPushThunk ucsActResultClear
                If ParseProduct() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p45
                    .ThunkPos = q45
                    GoTo L7
                End If
                pvPushThunk ucsAct_2_Sum, .CaptureBegin, .CaptureEnd
                GoTo L6
L7:
                .BufPos = p45
                .ThunkPos = q45
                Exit Do
L6:
            Loop
            pvPushThunk ucsAct_3_Sum, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -2
            ParseSum = True
        End If
    End With
End Function

Private Function ParseEOL() As Boolean
    Dim p119 As Long

    With ctx
        p119 = .BufPos
        If Not (.BufPos < .BufSize) Then
            .BufPos = p119
            ParseEOL = True
        End If
    End With
End Function

Private Function ParseProduct() As Boolean
    Dim p68 As Long
    Dim q68 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        If ParseValue() Then
            pvPushThunk ucsActVarSet, 1
            Do
                p68 = .BufPos
                q68 = .ThunkPos
                If Not (ParseTIMES()) Then
                    .BufPos = p68
                    .ThunkPos = q68
                    GoTo L12
                End If
                pvPushThunk ucsActResultClear
                If ParseValue() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p68
                    .ThunkPos = q68
                    GoTo L12
                End If
                pvPushThunk ucsAct_1_Product, .CaptureBegin, .CaptureEnd
                GoTo L10
L12:
                If Not (ParseDIVIDE()) Then
                    .BufPos = p68
                    .ThunkPos = q68
                    GoTo L11
                End If
                pvPushThunk ucsActResultClear
                If ParseValue() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p68
                    .ThunkPos = q68
                    GoTo L11
                End If
                pvPushThunk ucsAct_2_Product, .CaptureBegin, .CaptureEnd
                GoTo L10
L11:
                .BufPos = p68
                .ThunkPos = q68
                Exit Do
L10:
            Loop
            pvPushThunk ucsAct_3_Product, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -2
            ParseProduct = True
        End If
    End With
End Function

Private Function ParsePLUS() As Boolean
    With ctx
        If .BufData(.BufPos) = 43 Then              ' "+"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParsePLUS = True
        End If
    End With
End Function

Private Function ParseMINUS() As Boolean
    With ctx
        If .BufData(.BufPos) = 45 Then              ' "-"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseMINUS = True
        End If
    End With
End Function

Private Function ParseValue() As Boolean
    Dim p85 As Long
    Dim q85 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        p85 = .BufPos
        q85 = .ThunkPos
        If ParseNUMBER() Then
            pvPushThunk ucsAct_1_Value, .CaptureBegin, .CaptureEnd
            GoTo L15
        End If
        .BufPos = p85
        .ThunkPos = q85
        If ParseOPEN() Then
            pvPushThunk ucsActResultClear
            If ParseSum() Then
                pvPushThunk ucsActVarSet, 1
                If ParseCLOSE() Then
                    pvPushThunk ucsAct_2_Value, .CaptureBegin, .CaptureEnd
                    GoTo L15
                End If
                .BufPos = p85
                .ThunkPos = q85
            Else
                .BufPos = p85
                .ThunkPos = q85
            End If
        Else
            .BufPos = p85
            .ThunkPos = q85
        End If
        Exit Function
L15:
        pvPushThunk ucsActVarAlloc, -1
        ParseValue = True
    End With
End Function

Private Function ParseTIMES() As Boolean
    With ctx
        If .BufData(.BufPos) = 42 Then              ' "*"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseTIMES = True
        End If
    End With
End Function

Private Function ParseDIVIDE() As Boolean
    With ctx
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseDIVIDE = True
        End If
    End With
End Function

Private Function ParseNUMBER() As Boolean
    Dim lCaptureBegin As Long
    Dim i88 As Long
    Dim p93 As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        For i88 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i88 <> 0 Then
            p93 = .BufPos
            If .BufData(.BufPos) = 46 Then          ' "."
                .BufPos = .BufPos + 1
                Do
                    Select Case .BufData(.BufPos)
                    Case 48 To 57                   ' [0-9]
                        .BufPos = .BufPos + 1
                    Case Else
                        Exit Do
                    End Select
                Loop
            End If
            lCaptureEnd = .BufPos
            Call Parse_
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            Call pvSetAdvance
            ParseNUMBER = True
        End If
    End With
End Function

Private Function ParseOPEN() As Boolean
    With ctx
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseOPEN = True
        End If
    End With
End Function

Private Function ParseCLOSE() As Boolean
    With ctx
        If .BufData(.BufPos) = 41 Then              ' ")"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseCLOSE = True
        End If
    End With
End Function

Private Sub pvImplAction(ByVal eAction As UcsParserActionsEnum, ByVal lOffset As Long, ByVal lSize As Long)
    With ctx
        Select Case eAction
        Case ucsAct_1_Stmt
             .VarResult = .VarStack(.VarPos - 1)
        Case ucsAct_2_Stmt
             .VarResult = .VarStack(.VarPos - 1): .LastError = "Extra characters: " & Mid$(.Contents, lOffset, lSize)
        Case ucsAct_3_Sum
             .VarResult = .VarStack(.VarPos - 1)
        Case ucsAct_2_Sum
             .VarStack(.VarPos - 1) = .VarStack(.VarPos - 1) - .VarStack(.VarPos - 2)
        Case ucsAct_1_Sum
             .VarStack(.VarPos - 1) = .VarStack(.VarPos - 1) + .VarStack(.VarPos - 2)
        Case ucsAct_3_Product
             .VarResult = .VarStack(.VarPos - 1)
        Case ucsAct_2_Product
             .VarStack(.VarPos - 1) = .VarStack(.VarPos - 1) / .VarStack(.VarPos - 2)
        Case ucsAct_1_Product
             .VarStack(.VarPos - 1) = .VarStack(.VarPos - 1) * .VarStack(.VarPos - 2)
        Case ucsAct_1_Value
             .VarResult = Val(Mid$(.Contents, lOffset, lSize))
        Case ucsAct_2_Value
             .VarResult = .VarStack(.VarPos - 1)
        End Select
    End With
End Sub

