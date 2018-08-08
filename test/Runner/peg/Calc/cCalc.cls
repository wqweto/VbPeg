VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Auto-generated on 27.7.2018 19:38:22
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
    ucsAct_1_Stmt
    ucsAct_2_Stmt
    ucsAct_1_Expr
    ucsAct_2_Expr
    ucsAct_1_ID
    ucsAct_3_Sum
    ucsAct_2_Sum
    ucsAct_1_Sum
    ucsAct_3_Product
    ucsAct_2_Product
    ucsAct_1_Product
    ucsAct_1_Value
    ucsAct_2_Value
    ucsAct_3_Value
    ucsAct_1_NUMBER
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
    '-- from @members setting
    LocalVars(0 to 255)     As Long
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
    ParserVersion = "27.7.2018 19:38:22"
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

'= generated functions ===================================================

Public Function ParseStmt() As Boolean
    Dim p22 As Long
    Dim q22 As Long
    Dim p17 As Long
    Dim q17 As Long
    Dim p15 As Long
    Dim q15 As Long
    Dim e15 As String

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        p22 = .BufPos
        q22 = .ThunkPos
        Call Parse_
        pvPushThunk ucsActResultClear
        If ParseExpr() Then
            pvPushThunk ucsActVarSet, 1
            If ParseEOL() Then
                pvPushThunk ucsAct_1_Stmt, .CaptureBegin, .CaptureEnd
                pvPushThunk ucsActVarAlloc, -1
                ParseStmt = True
                Exit Function
            Else
                .BufPos = p22
                .ThunkPos = q22
            End If
        Else
            .BufPos = p22
            .ThunkPos = q22
        End If
        Do
            p17 = .BufPos
            q17 = .ThunkPos
            p15 = .BufPos
            q15 = .ThunkPos
            e15 = .LastExpected
            If ParseEOL() Then
                .BufPos = p17
                .ThunkPos = q17
                Exit Do
            Else
                .BufPos = p15
                .ThunkPos = q15
                .LastExpected = e15
            End If
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
            Else
                .BufPos = p17
                .ThunkPos = q17
                Exit Do
            End If
        Loop
        If ParseEOL() Then
            pvPushThunk ucsAct_2_Stmt, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -1
            ParseStmt = True
            Exit Function
        Else
            .BufPos = p22
            .ThunkPos = q22
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

Public Function ParseExpr() As Boolean
    Dim p40 As Long
    Dim q40 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        p40 = .BufPos
        q40 = .ThunkPos
        pvPushThunk ucsActResultClear
        If ParseID() Then
            pvPushThunk ucsActVarSet, 1
            If ParseASSIGN() Then
                pvPushThunk ucsActResultClear
                If ParseExpr() Then
                    pvPushThunk ucsActVarSet, 2
                    pvPushThunk ucsAct_1_Expr, .CaptureBegin, .CaptureEnd
                    pvPushThunk ucsActVarAlloc, -2
                    ParseExpr = True
                    Exit Function
                Else
                    .BufPos = p40
                    .ThunkPos = q40
                End If
            Else
                .BufPos = p40
                .ThunkPos = q40
            End If
        Else
            .BufPos = p40
            .ThunkPos = q40
        End If
        pvPushThunk ucsActResultClear
        If ParseSum() Then
            pvPushThunk ucsActVarSet, 2
            pvPushThunk ucsAct_2_Expr, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -2
            ParseExpr = True
            Exit Function
        Else
            .BufPos = p40
            .ThunkPos = q40
        End If
    End With
End Function

Private Function ParseEOL() As Boolean
    With ctx
        If .BufData(.BufPos) = 10 Then              ' "\n"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseEOL = True
            Exit Function
        End If
        If .BufData(.BufPos) = 13 And .BufData(.BufPos + 1) = 10 Then ' "\r\n"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseEOL = True
            Exit Function
        End If
        If .BufData(.BufPos) = 13 Then              ' "\r"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseEOL = True
            Exit Function
        End If
        If .BufData(.BufPos) = 59 Then              ' ";"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseEOL = True
        End If
    End With
End Function

Private Function ParseID() As Boolean
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        Select Case .BufData(.BufPos)
        Case 97 To 122                              ' [a-z]
            .BufPos = .BufPos + 1
            lCaptureEnd = .BufPos
            Call Parse_
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_ID, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseID = True
        End Select
    End With
End Function

Private Function ParseASSIGN() As Boolean
    With ctx
        If .BufData(.BufPos) = 61 Then              ' "="
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseASSIGN = True
        End If
    End With
End Function

Public Function ParseSum() As Boolean
    Dim p61 As Long
    Dim q61 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        If ParseProduct() Then
            pvPushThunk ucsActVarSet, 1
            Do
                p61 = .BufPos
                q61 = .ThunkPos
                If ParsePLUS() Then
                    pvPushThunk ucsActResultClear
                    If ParseProduct() Then
                        pvPushThunk ucsActVarSet, 2
                        pvPushThunk ucsAct_1_Sum, .CaptureBegin, .CaptureEnd
                    Else
                        .BufPos = p61
                        .ThunkPos = q61
                        If ParseMINUS() Then
                            pvPushThunk ucsActResultClear
                            If ParseProduct() Then
                                pvPushThunk ucsActVarSet, 2
                                pvPushThunk ucsAct_2_Sum, .CaptureBegin, .CaptureEnd
                            Else
                                .BufPos = p61
                                .ThunkPos = q61
                                .BufPos = p61
                                .ThunkPos = q61
                                Exit Do
                            End If
                        Else
                            .BufPos = p61
                            .ThunkPos = q61
                            .BufPos = p61
                            .ThunkPos = q61
                            Exit Do
                        End If
                    End If
                Else
                    .BufPos = p61
                    .ThunkPos = q61
                    If ParseMINUS() Then
                        pvPushThunk ucsActResultClear
                        If ParseProduct() Then
                            pvPushThunk ucsActVarSet, 2
                            pvPushThunk ucsAct_2_Sum, .CaptureBegin, .CaptureEnd
                        Else
                            .BufPos = p61
                            .ThunkPos = q61
                            .BufPos = p61
                            .ThunkPos = q61
                            Exit Do
                        End If
                    Else
                        .BufPos = p61
                        .ThunkPos = q61
                        .BufPos = p61
                        .ThunkPos = q61
                        Exit Do
                    End If
                End If
            Loop
            pvPushThunk ucsAct_3_Sum, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -2
            ParseSum = True
        End If
    End With
End Function

Public Function ParseProduct() As Boolean
    Dim p84 As Long
    Dim q84 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        If ParseValue() Then
            pvPushThunk ucsActVarSet, 1
            Do
                p84 = .BufPos
                q84 = .ThunkPos
                If ParseTIMES() Then
                    pvPushThunk ucsActResultClear
                    If ParseValue() Then
                        pvPushThunk ucsActVarSet, 2
                        pvPushThunk ucsAct_1_Product, .CaptureBegin, .CaptureEnd
                    Else
                        .BufPos = p84
                        .ThunkPos = q84
                        If ParseDIVIDE() Then
                            pvPushThunk ucsActResultClear
                            If ParseValue() Then
                                pvPushThunk ucsActVarSet, 2
                                pvPushThunk ucsAct_2_Product, .CaptureBegin, .CaptureEnd
                            Else
                                .BufPos = p84
                                .ThunkPos = q84
                                .BufPos = p84
                                .ThunkPos = q84
                                Exit Do
                            End If
                        Else
                            .BufPos = p84
                            .ThunkPos = q84
                            .BufPos = p84
                            .ThunkPos = q84
                            Exit Do
                        End If
                    End If
                Else
                    .BufPos = p84
                    .ThunkPos = q84
                    If ParseDIVIDE() Then
                        pvPushThunk ucsActResultClear
                        If ParseValue() Then
                            pvPushThunk ucsActVarSet, 2
                            pvPushThunk ucsAct_2_Product, .CaptureBegin, .CaptureEnd
                        Else
                            .BufPos = p84
                            .ThunkPos = q84
                            .BufPos = p84
                            .ThunkPos = q84
                            Exit Do
                        End If
                    Else
                        .BufPos = p84
                        .ThunkPos = q84
                        .BufPos = p84
                        .ThunkPos = q84
                        Exit Do
                    End If
                End If
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

Public Function ParseValue() As Boolean
    Dim p99 As Long
    Dim q99 As Long
    Dim p97 As Long
    Dim q97 As Long
    Dim e97 As String

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        p99 = .BufPos
        q99 = .ThunkPos
        pvPushThunk ucsActResultClear
        If ParseNUMBER() Then
            pvPushThunk ucsActVarSet, 1
            pvPushThunk ucsAct_1_Value, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -1
            ParseValue = True
            Exit Function
        Else
            .BufPos = p99
            .ThunkPos = q99
        End If
        pvPushThunk ucsActResultClear
        If ParseID() Then
            pvPushThunk ucsActVarSet, 1
            p97 = .BufPos
            q97 = .ThunkPos
            e97 = .LastExpected
            If ParseASSIGN() Then
                .BufPos = p99
                .ThunkPos = q99
            Else
                .BufPos = p97
                .ThunkPos = q97
                .LastExpected = e97
                pvPushThunk ucsAct_2_Value, .CaptureBegin, .CaptureEnd
                pvPushThunk ucsActVarAlloc, -1
                ParseValue = True
                Exit Function
            End If
        Else
            .BufPos = p99
            .ThunkPos = q99
        End If
        If ParseOPEN() Then
            pvPushThunk ucsActResultClear
            If ParseExpr() Then
                pvPushThunk ucsActVarSet, 1
                If ParseCLOSE() Then
                    pvPushThunk ucsAct_3_Value, .CaptureBegin, .CaptureEnd
                    pvPushThunk ucsActVarAlloc, -1
                    ParseValue = True
                    Exit Function
                Else
                    .BufPos = p99
                    .ThunkPos = q99
                End If
            Else
                .BufPos = p99
                .ThunkPos = q99
            End If
        Else
            .BufPos = p99
            .ThunkPos = q99
        End If
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
    Dim i111 As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        For i111 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i111 <> 0 Then
            lCaptureEnd = .BufPos
            Call Parse_
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_NUMBER, lCaptureBegin, lCaptureEnd
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
             ConsolePrint .VarStack(.VarPos - 1) & vbCrLf
        Case ucsAct_2_Stmt
             ConsolePrint "error" & vbCrLf
        Case ucsAct_1_Expr
             .LocalVars(.VarStack(.VarPos - 1)) = .VarStack(.VarPos - 2) : .VarResult = .VarStack(.VarPos - 2)
        Case ucsAct_2_Expr
             .VarResult = .VarStack(.VarPos - 2)
        Case ucsAct_1_ID
             .VarResult = Asc(Mid$(.Contents, lOffset, lSize))
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
             .VarResult = CLng(Mid$(.Contents, lOffset, lSize))
        Case ucsAct_2_Value
             .VarResult = .LocalVars(.VarStack(.VarPos - 1))
        Case ucsAct_3_Value
             .VarResult = .VarStack(.VarPos - 1)
        Case ucsAct_1_NUMBER
             .VarResult = CLng(Mid$(.Contents, lOffset, lSize))
        End Select
    End With
End Sub

