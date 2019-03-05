VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTest8"
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
    ucsAct_3_start
    ucsAct_2_start
    ucsAct_1_start
    ucsAct_2_binexpr
    ucsAct_1_binexpr
    ucsAct_1_unexpr
    ucsAct_1_BINOP
    ucsAct_1_UNOP
    ucsAct_1_DOUBLE
    ucsAct_1_IDENT
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
    '-- from @members setting
    BinopInfo           As Dictionary
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

Public Function Parsestart() As Boolean
    Dim i15 As Long
    Dim p13 As Long
    Dim q13 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        Call Parse_
        pvPushThunk ucsActVarSet, 1
        pvPushThunk ucsAct_1_start, .CaptureBegin, .CaptureEnd
        For i15 = 0 To LNG_MAXINT
            p13 = .BufPos
            q13 = .ThunkPos
            pvPushThunk ucsActResultClear
            If Parsestmt() Then
                pvPushThunk ucsActVarSet, 2
            Else
                .BufPos = p13
                .ThunkPos = q13
                Exit For
            End If
            pvPushThunk ucsAct_2_start, .CaptureBegin, .CaptureEnd
        Next
        If i15 <> 0 Then
            pvPushThunk ucsAct_3_start, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -2
            Parsestart = True
        End If
    End With
End Function

Private Sub Parse_()
    Dim p122 As Long

    With ctx
        Do
            p122 = .BufPos
            If Not (ParseLineComment()) Then
                .BufPos = p122
                Select Case .BufData(.BufPos)
                Case 32, 9, 13, 10                  ' [ \t\r\n]
                    .BufPos = .BufPos + 1
                Case Else
                    .BufPos = p122
                    Exit Do
                End Select
            End If
        Loop
        Call pvSetAdvance
    End With
End Sub

Private Function Parsestmt() As Boolean
    Dim p20 As Long
    Dim q20 As Long

    With ctx
        If Parsebinexpr() Then
            p20 = .BufPos
            q20 = .ThunkPos
            If Not (ParseSEMI()) Then
                .BufPos = p20
                .ThunkPos = q20
            End If
            Parsestmt = True
        End If
    End With
End Function

Private Function ParseSEMI() As Boolean
    With ctx
        If .BufData(.BufPos) = 59 Then              ' ";"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseSEMI = True
        End If
    End With
End Function

Private Function Parsebinexpr(Optional ByVal lMinPrec As Long) As Boolean
    Dim p34 As Long
    Dim q34 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 3
        pvPushThunk ucsActResultClear
        If Parseunexpr() Then
            pvPushThunk ucsActVarSet, 1
            Do
                p34 = .BufPos
                q34 = .ThunkPos
                pvPushThunk ucsActResultClear
                If ParseBINOP() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p34
                    .ThunkPos = q34
                    Exit Do
                End If
                Dim vOpInfo         As Variant
                vOpInfo = GetBinopInfo(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin))
                If Not (vOpInfo(0) >= lMinPrec) Then
                    .BufPos = p34
                    .ThunkPos = q34
                    Exit Do
                End If
                pvPushThunk ucsActResultClear
                If Parsebinexpr(vOpInfo(0) + vOpInfo(1)) Then
                    pvPushThunk ucsActVarSet, 3
                Else
                    .BufPos = p34
                    .ThunkPos = q34
                    Exit Do
                End If
                pvPushThunk ucsAct_1_binexpr, .CaptureBegin, .CaptureEnd
            Loop
            pvPushThunk ucsAct_2_binexpr, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -3
            Parsebinexpr = True
        End If
    End With
End Function

Private Function Parseunexpr() As Boolean
    Dim p56 As Long
    Dim q56 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        p56 = .BufPos
        q56 = .ThunkPos
        pvPushThunk ucsActResultClear
        If ParseUNOP() Then
            pvPushThunk ucsActVarSet, 1
            pvPushThunk ucsActResultClear
            If Parseunexpr() Then
                pvPushThunk ucsActVarSet, 2
                pvPushThunk ucsAct_1_unexpr, .CaptureBegin, .CaptureEnd
                GoTo L7
            End If
            .BufPos = p56
            .ThunkPos = q56
        Else
            .BufPos = p56
            .ThunkPos = q56
        End If
        If Parseprimary() Then
            GoTo L7
        End If
        .BufPos = p56
        .ThunkPos = q56
        Exit Function
L7:
        pvPushThunk ucsActVarAlloc, -2
        Parseunexpr = True
    End With
End Function

Private Function ParseBINOP() As Boolean
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        Select Case .BufData(.BufPos)
        Case 45, 43, 42, 47, 60, 62, 61, 58, 124, 38, 94 ' [-+*/<>=:|&^]
            .BufPos = .BufPos + 1
            If .BufData(.BufPos) = 61 Then          ' "="
                .BufPos = .BufPos + 1
            End If
            lCaptureEnd = .BufPos
            Call Parse_
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_BINOP, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseBINOP = True
        End Select
    End With
End Function

Private Function ParseUNOP() As Boolean
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        Select Case .BufData(.BufPos)
        Case 45, 43, 33, 126                        ' [-+!~]
            .BufPos = .BufPos + 1
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_UNOP, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseUNOP = True
        End Select
    End With
End Function

Private Function Parseprimary() As Boolean
    Dim p65 As Long
    Dim q65 As Long

    With ctx
        p65 = .BufPos
        q65 = .ThunkPos
        If ParseDOUBLE() Then
            GoTo L12
        End If
        .BufPos = p65
        .ThunkPos = q65
        If ParseLPAREN() Then
            If Parsebinexpr() Then
                If ParseRPAREN() Then
                    GoTo L12
                End If
                .BufPos = p65
                .ThunkPos = q65
            Else
                .BufPos = p65
                .ThunkPos = q65
            End If
        Else
            .BufPos = p65
            .ThunkPos = q65
        End If
        If ParseIDENT() Then
            GoTo L12
        End If
        .BufPos = p65
        .ThunkPos = q65
        Exit Function
L12:
        Parseprimary = True
    End With
End Function

Private Function ParseDOUBLE() As Boolean
    Dim lCaptureBegin As Long
    Dim p91 As Long
    Dim i79 As Long
    Dim i90 As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        p91 = .BufPos
        For i79 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i79 <> 0 Then
            If .BufData(.BufPos) = 46 Then          ' "."
                .BufPos = .BufPos + 1
            End If
            Do
                Select Case .BufData(.BufPos)
                Case 48 To 57                       ' [0-9]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit Do
                End Select
            Loop
            GoTo L15
        End If
        .BufPos = p91
        Do
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit Do
            End Select
        Loop
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
            For i90 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 48 To 57                       ' [0-9]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If i90 <> 0 Then
                GoTo L15
            End If
            .BufPos = p91
        Else
            .BufPos = p91
        End If
        Exit Function
L15:
        lCaptureEnd = .BufPos
        Call Parse_
        .CaptureBegin = lCaptureBegin
        .CaptureEnd = lCaptureEnd
        pvPushThunk ucsAct_1_DOUBLE, lCaptureBegin, lCaptureEnd
        Call pvSetAdvance
        ParseDOUBLE = True
    End With
End Function

Private Function ParseLPAREN() As Boolean
    With ctx
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseLPAREN = True
        End If
    End With
End Function

Private Function ParseRPAREN() As Boolean
    With ctx
        If .BufData(.BufPos) = 41 Then              ' ")"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseRPAREN = True
        End If
    End With
End Function

Private Function ParseIDENT() As Boolean
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        Select Case .BufData(.BufPos)
        Case 97 To 122, 65 To 90, 95                ' [a-zA-Z_]
            .BufPos = .BufPos + 1
            Do
                Select Case .BufData(.BufPos)
                Case 97 To 122, 65 To 90, 95, 48 To 57 ' [a-zA-Z_0-9]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit Do
                End Select
            Loop
            lCaptureEnd = .BufPos
            Call Parse_
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_IDENT, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseIDENT = True
        End Select
    End With
End Function

Private Function ParseLineComment() As Boolean
    With ctx
        If .BufData(.BufPos) = 35 Then              ' "#"
            .BufPos = .BufPos + 1
            Do
                Select Case .BufData(.BufPos)
                Case 13, 10                         ' [\r\n]
                    Exit Do
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                    Else
                        Exit Do
                    End If
                End Select
            Loop
            Call pvSetAdvance
            ParseLineComment = True
        End If
    End With
End Function

Private Sub pvImplAction(ByVal eAction As UcsParserActionsEnum, ByVal lOffset As Long, ByVal lSize As Long)
    Dim oJson As Object

    With ctx
    Select Case eAction
    Case ucsAct_3_start
           Set .VarResult = .VarStack(.VarPos - 1)
    Case ucsAct_2_start
           Set oJson = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, -1) = .VarStack(.VarPos - 2)
    Case ucsAct_1_start
           JsonItem(oJson, -1) = Empty
                                                                            Set .VarStack(.VarPos - 1) = oJson
    Case ucsAct_2_binexpr
           AssignVariant .VarResult, .VarStack(.VarPos - 1)
    Case ucsAct_1_binexpr
           JsonItem(oJson, "Tag") = "BinOp"
                                                                            JsonItem(oJson, "Left") = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, "Oper") = .VarStack(.VarPos - 2)
                                                                            JsonItem(oJson, "Right") = .VarStack(.VarPos - 3)
                                                                            JsonItem(oJson, "BinopInfo") = GetBinopInfo(C_Str(.VarStack(.VarPos - 2)))
                                                                            Set .VarStack(.VarPos - 1) = oJson
    Case ucsAct_1_unexpr
           JsonItem(oJson, "Tag") = "UnOp"
                                                                            JsonItem(oJson, "Oper") = .VarStack(.VarPos - 1)
                                                                            JsonItem(oJson, "Expr") = .VarStack(.VarPos - 2)
                                                                            Set .VarResult = oJson
    Case ucsAct_1_BINOP
         .VarResult = Mid$(.Contents, lOffset, lSize)
    Case ucsAct_1_UNOP
         .VarResult = Mid$(.Contents, lOffset, lSize)
    Case ucsAct_1_DOUBLE
         .VarResult = C_Dbl(Mid$(.Contents, lOffset, lSize))
    Case ucsAct_1_IDENT
         .VarResult = Mid$(.Contents, lOffset, lSize)
    End Select
    End With
End Sub

'= user-defined functions ================================================

Public Function GetBinopInfo(sOp As String) As Variant
    If ctx.BinopInfo Is Nothing Then
        Set ctx.BinopInfo = New Scripting.Dictionary
        ctx.BinopInfo.Item("<") = Array(10, 1)
        ctx.BinopInfo.Item("+") = Array(20, 1)
        ctx.BinopInfo.Item("-") = Array(20, 1)
        ctx.BinopInfo.Item("*") = Array(40, 1)
        ctx.BinopInfo.Item("/") = Array(40, 1)
        ctx.BinopInfo.Item("^") = Array(50, 0)
    End If
    If ctx.BinopInfo.Exists(sOp) Then
        GetBinopInfo = ctx.BinopInfo.Item(sOp)
    Else
        GetBinopInfo = Array(0, 0)
    End If
End Function

Public Function SetBinopInfo(sOp As String, ByVal lPrec As Long,  ByVal lAssoc As Long) As Variant
    If ctx.BinopInfo Is Nothing Then
        GetBinopInfo vbNullString
    End If
    ctx.BinopInfo.Item(sOp) = Array(lPrec, lAssoc, "Override")
End Function
