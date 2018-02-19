Attribute VB_Name = "mdCalc"
' Auto-generated on 19.2.2018 23:25:29
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
    VbPegParserVersion = "19.2.2018 23:25:29"
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function VbPegMatch(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant, Optional Result As Variant) As Long
    If VbPegBeginMatch(sSubject, StartPos, UserData) Then
        If VbPegParseStmt() Then
            VbPegMatch = VbPegEndMatch(Result)
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

Public Function VbPegEndMatch(Optional Result As Variant) As Long
    Dim lIdx            As Long
    Dim uEmpty          As UcsParserType
    
    With ctx
        ReDim .VarStack(0 To 1024) As Variant
        For lIdx = 0 To .ThunkPos - 1
            Select Case .ThunkData(lIdx).Action
            Case ucsActVarAlloc
                .VarPos = .VarPos + .ThunkData(lIdx).CaptureBegin
            Case ucsActVarSet
                .VarStack(.VarPos - .ThunkData(lIdx).CaptureBegin) = .VarResult
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
        VbPegEndMatch = .BufPos
    End With
    ctx = uEmpty
End Function

Private Sub pvPushAction(ByVal eAction As UcsParserActionsEnum)
    pvPushThunk eAction, ctx.CaptureBegin, ctx.CaptureEnd
End Sub

Private Sub pvPushThunk(ByVal eAction As UcsParserActionsEnum, ByVal lBegin As Long, Optional ByVal lEnd As Long)
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

Public Function VbPegParseStmt() As Boolean
    Dim p19 As Long
    Dim q19 As Long
    Dim p14 As Long
    Dim q14 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        p19 = .BufPos
        q19 = .ThunkPos
        Call Parse_
        If VbPegParseExpr() Then
            pvPushThunk ucsActVarSet, 1
            If ParseEOL() Then
                pvPushAction ucsAct_1_Stmt
                pvPushThunk ucsActVarAlloc, -1
                VbPegParseStmt = True
                Exit Function
            Else
                .BufPos = p19
                .ThunkPos = q19
            End If
        Else
            .BufPos = p19
            .ThunkPos = q19
        End If
        Do
            p14 = .BufPos
            q14 = .ThunkPos
            If ParseEOL() Then
                .BufPos = p14
                .ThunkPos = q14
                Exit Do
            End If
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
            Else
                .BufPos = p14
                .ThunkPos = q14
                Exit Do
            End If
        Loop
        If ParseEOL() Then
            pvPushAction ucsAct_2_Stmt
            pvPushThunk ucsActVarAlloc, -1
            VbPegParseStmt = True
            Exit Function
        Else
            .BufPos = p19
            .ThunkPos = q19
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
    End With
End Sub

Public Function VbPegParseExpr() As Boolean
    Dim p33 As Long
    Dim q33 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        p33 = .BufPos
        q33 = .ThunkPos
        If ParseID() Then
            pvPushThunk ucsActVarSet, 1
            If ParseASSIGN() Then
                If VbPegParseExpr() Then
                    pvPushThunk ucsActVarSet, 2
                    pvPushAction ucsAct_1_Expr
                    pvPushThunk ucsActVarAlloc, -2
                    VbPegParseExpr = True
                    Exit Function
                Else
                    .BufPos = p33
                    .ThunkPos = q33
                End If
            Else
                .BufPos = p33
                .ThunkPos = q33
            End If
        End If
        If VbPegParseSum() Then
            pvPushThunk ucsActVarSet, 2
            pvPushAction ucsAct_2_Expr
            pvPushThunk ucsActVarAlloc, -2
            VbPegParseExpr = True
        End If
    End With
End Function

Private Function ParseEOL() As Boolean
    With ctx
        If .BufData(.BufPos) = 10 Then              ' "\n"
            .BufPos = .BufPos + 1
            ParseEOL = True
            Exit Function
        End If
        If .BufData(.BufPos) = 13 And .BufData(.BufPos + 1) = 10 Then ' "\r\n"
            .BufPos = .BufPos + 2
            ParseEOL = True
            Exit Function
        End If
        If .BufData(.BufPos) = 13 Then              ' "\r"
            .BufPos = .BufPos + 1
            ParseEOL = True
            Exit Function
        End If
        If .BufData(.BufPos) = 59 Then              ' ";"
            .BufPos = .BufPos + 1
            ParseEOL = True
        End If
    End With
End Function

Private Function ParseID() As Boolean
    With ctx
        .CaptureBegin = .BufPos
        Select Case .BufData(.BufPos)
        Case 97 To 122                              ' [a-z]
            .BufPos = .BufPos + 1
            .CaptureEnd = .BufPos
            Call Parse_
            pvPushAction ucsAct_1_ID
            ParseID = True
        End Select
    End With
End Function

Private Function ParseASSIGN() As Boolean
    With ctx
        If .BufData(.BufPos) = 61 Then              ' "="
            .BufPos = .BufPos + 1
            Call Parse_
            ParseASSIGN = True
        End If
    End With
End Function

Public Function VbPegParseSum() As Boolean
    Dim p48 As Long
    Dim q48 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        If VbPegParseProduct() Then
            pvPushThunk ucsActVarSet, 1
            Do
                p48 = .BufPos
                q48 = .ThunkPos
                If Not ParsePLUS() Then
                    If Not ParseMINUS() Then
                        Exit Do
                    End If
                    If VbPegParseProduct() Then
                        pvPushThunk ucsActVarSet, 2
                    Else
                        .BufPos = p48
                        .ThunkPos = q48
                        Exit Do
                    End If
                    pvPushAction ucsAct_2_Sum
                End If
                If VbPegParseProduct() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p48
                    .ThunkPos = q48
                    If Not ParseMINUS() Then
                        Exit Do
                    End If
                    If VbPegParseProduct() Then
                        pvPushThunk ucsActVarSet, 2
                    Else
                        .BufPos = p48
                        .ThunkPos = q48
                        Exit Do
                    End If
                    pvPushAction ucsAct_2_Sum
                End If
                pvPushAction ucsAct_1_Sum
            Loop
            pvPushAction ucsAct_3_Sum
            pvPushThunk ucsActVarAlloc, -2
            VbPegParseSum = True
        End If
    End With
End Function

Public Function VbPegParseProduct() As Boolean
    Dim p66 As Long
    Dim q66 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        If VbPegParseValue() Then
            pvPushThunk ucsActVarSet, 1
            Do
                p66 = .BufPos
                q66 = .ThunkPos
                If Not ParseTIMES() Then
                    If Not ParseDIVIDE() Then
                        Exit Do
                    End If
                    If VbPegParseValue() Then
                        pvPushThunk ucsActVarSet, 2
                    Else
                        .BufPos = p66
                        .ThunkPos = q66
                        Exit Do
                    End If
                    pvPushAction ucsAct_2_Product
                End If
                If VbPegParseValue() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p66
                    .ThunkPos = q66
                    If Not ParseDIVIDE() Then
                        Exit Do
                    End If
                    If VbPegParseValue() Then
                        pvPushThunk ucsActVarSet, 2
                    Else
                        .BufPos = p66
                        .ThunkPos = q66
                        Exit Do
                    End If
                    pvPushAction ucsAct_2_Product
                End If
                pvPushAction ucsAct_1_Product
            Loop
            pvPushAction ucsAct_3_Product
            pvPushThunk ucsActVarAlloc, -2
            VbPegParseProduct = True
        End If
    End With
End Function

Private Function ParsePLUS() As Boolean
    With ctx
        If .BufData(.BufPos) = 43 Then              ' "+"
            .BufPos = .BufPos + 1
            Call Parse_
            ParsePLUS = True
        End If
    End With
End Function

Private Function ParseMINUS() As Boolean
    With ctx
        If .BufData(.BufPos) = 45 Then              ' "-"
            .BufPos = .BufPos + 1
            Call Parse_
            ParseMINUS = True
        End If
    End With
End Function

Public Function VbPegParseValue() As Boolean
    Dim p80 As Long
    Dim q80 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        p80 = .BufPos
        q80 = .ThunkPos
        If ParseNUMBER() Then
            pvPushThunk ucsActVarSet, 1
            pvPushAction ucsAct_1_Value
            pvPushThunk ucsActVarAlloc, -1
            VbPegParseValue = True
            Exit Function
        End If
        If ParseID() Then
            pvPushThunk ucsActVarSet, 1
            If ParseASSIGN() Then
                .BufPos = p80
                .ThunkPos = q80
            Else
                pvPushAction ucsAct_2_Value
                pvPushThunk ucsActVarAlloc, -1
                VbPegParseValue = True
                Exit Function
            End If
        End If
        If ParseOPEN() Then
            If VbPegParseExpr() Then
                pvPushThunk ucsActVarSet, 1
                If ParseCLOSE() Then
                    pvPushAction ucsAct_3_Value
                    pvPushThunk ucsActVarAlloc, -1
                    VbPegParseValue = True
                    Exit Function
                Else
                    .BufPos = p80
                    .ThunkPos = q80
                End If
            Else
                .BufPos = p80
                .ThunkPos = q80
            End If
        End If
    End With
End Function

Private Function ParseTIMES() As Boolean
    With ctx
        If .BufData(.BufPos) = 42 Then              ' "*"
            .BufPos = .BufPos + 1
            Call Parse_
            ParseTIMES = True
        End If
    End With
End Function

Private Function ParseDIVIDE() As Boolean
    With ctx
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            Call Parse_
            ParseDIVIDE = True
        End If
    End With
End Function

Private Function ParseNUMBER() As Boolean
    Dim i90 As Long

    With ctx
        .CaptureBegin = .BufPos
        For i90 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i90 <> 0 Then
            .CaptureEnd = .BufPos
            Call Parse_
            pvPushAction ucsAct_1_NUMBER
            ParseNUMBER = True
        End If
    End With
End Function

Private Function ParseOPEN() As Boolean
    With ctx
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call Parse_
            ParseOPEN = True
        End If
    End With
End Function

Private Function ParseCLOSE() As Boolean
    With ctx
        If .BufData(.BufPos) = 41 Then              ' ")"
            .BufPos = .BufPos + 1
            Call Parse_
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
             .UserData(.VarStack(.VarPos - 1)) = .VarStack(.VarPos - 2): .VarResult = .VarStack(.VarPos - 2)
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
             .VarResult = .UserData(.VarStack(.VarPos - 1))
        Case ucsAct_3_Value
             .VarResult = .VarStack(.VarPos - 1)
        Case ucsAct_1_NUMBER
             .VarResult = CLng(Mid$(.Contents, lOffset, lSize))
        End Select
    End With
End Sub

