Attribute VB_Name = "mdCalc"
' Auto-generated on 17.3.2018 16:35:28
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
    VbPegParserVersion = "17.3.2018 16:35:28"
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function VbPegMatch(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant, Optional Result As Variant) As Long
    If VbPegBeginMatch(sSubject, StartPos, UserData) Then
        If VbPegParseStmt() Then
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
        If IsObject(.UserData) Then
            Set UserData = .UserData
        Else
            UserData = .UserData
        End If
        VbPegEndMatch = .BufPos
    End With
    uEmpty.LastError = ctx.LastError
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
    Dim p7 As Long
    Dim q7 As Long
    Dim p22 As Long
    Dim q22 As Long
    Dim i17 As Long
    Dim p16 As Long
    Dim q16 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        p7 = .BufPos
        q7 = .ThunkPos
        Call Parse_
        If VbPegParseSum() Then
            pvPushThunk ucsActVarSet, 1
            p22 = .BufPos
            q22 = .ThunkPos
            If ParseEOL() Then
                pvPushAction ucsAct_1_Stmt
                pvPushThunk ucsActVarAlloc, -1
                VbPegParseStmt = True
                Exit Function
            End If
            .CaptureBegin = .BufPos
            For i17 = 0 To LNG_MAXINT
                p16 = .BufPos
                q16 = .ThunkPos
                If ParseEOL() Then
                    .BufPos = p16
                    .ThunkPos = q16
                    Exit For
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p16
                    .ThunkPos = q16
                    Exit For
                End If
            Next
            If i17 <> 0 Then
                .CaptureEnd = .BufPos
                If ParseEOL() Then
                    pvPushAction ucsAct_2_Stmt
                    pvPushThunk ucsActVarAlloc, -1
                    VbPegParseStmt = True
                    Exit Function
                Else
                    .BufPos = p22
                    .ThunkPos = q22
                End If
            End If
            .BufPos = p7
            .ThunkPos = q7
        Else
            .BufPos = p7
            .ThunkPos = q7
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

Public Function VbPegParseSum() As Boolean
    Dim p37 As Long
    Dim q37 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        If VbPegParseProduct() Then
            pvPushThunk ucsActVarSet, 1
            Do
                p37 = .BufPos
                q37 = .ThunkPos
                If Not ParsePLUS() Then
                    If Not ParseMINUS() Then
                        Exit Do
                    End If
                    If VbPegParseProduct() Then
                        pvPushThunk ucsActVarSet, 2
                    Else
                        .BufPos = p37
                        .ThunkPos = q37
                        Exit Do
                    End If
                    pvPushAction ucsAct_2_Sum
                    GoTo L1
                End If
                If VbPegParseProduct() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p37
                    .ThunkPos = q37
                    If Not ParseMINUS() Then
                        Exit Do
                    End If
                    If VbPegParseProduct() Then
                        pvPushThunk ucsActVarSet, 2
                    Else
                        .BufPos = p37
                        .ThunkPos = q37
                        Exit Do
                    End If
                    pvPushAction ucsAct_2_Sum
                    GoTo L1
                End If
                pvPushAction ucsAct_1_Sum
L1:
            Loop
            pvPushAction ucsAct_3_Sum
            pvPushThunk ucsActVarAlloc, -2
            VbPegParseSum = True
        End If
    End With
End Function

Private Function ParseEOL() As Boolean
    With ctx
        If Not .BufPos < .BufSize Then
            ParseEOL = True
        End If
    End With
End Function

Public Function VbPegParseProduct() As Boolean
    Dim p55 As Long
    Dim q55 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        If VbPegParseValue() Then
            pvPushThunk ucsActVarSet, 1
            Do
                p55 = .BufPos
                q55 = .ThunkPos
                If Not ParseTIMES() Then
                    If Not ParseDIVIDE() Then
                        Exit Do
                    End If
                    If VbPegParseValue() Then
                        pvPushThunk ucsActVarSet, 2
                    Else
                        .BufPos = p55
                        .ThunkPos = q55
                        Exit Do
                    End If
                    pvPushAction ucsAct_2_Product
                    GoTo L2
                End If
                If VbPegParseValue() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p55
                    .ThunkPos = q55
                    If Not ParseDIVIDE() Then
                        Exit Do
                    End If
                    If VbPegParseValue() Then
                        pvPushThunk ucsActVarSet, 2
                    Else
                        .BufPos = p55
                        .ThunkPos = q55
                        Exit Do
                    End If
                    pvPushAction ucsAct_2_Product
                    GoTo L2
                End If
                pvPushAction ucsAct_1_Product
L2:
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
    Dim p71 As Long
    Dim q71 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        p71 = .BufPos
        q71 = .ThunkPos
        If ParseNUMBER() Then
            pvPushThunk ucsActVarSet, 1
            pvPushAction ucsAct_1_Value
            pvPushThunk ucsActVarAlloc, -1
            VbPegParseValue = True
            Exit Function
        End If
        If ParseOPEN() Then
            If VbPegParseSum() Then
                pvPushThunk ucsActVarSet, 1
                If ParseCLOSE() Then
                    pvPushAction ucsAct_2_Value
                    pvPushThunk ucsActVarAlloc, -1
                    VbPegParseValue = True
                    Exit Function
                Else
                    .BufPos = p71
                    .ThunkPos = q71
                End If
            Else
                .BufPos = p71
                .ThunkPos = q71
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
    Dim i74 As Long

    With ctx
        .CaptureBegin = .BufPos
        For i74 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i74 <> 0 Then
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
                GoTo L3
            End If
L3:
            .CaptureEnd = .BufPos
            Call Parse_
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
            
            On Error Resume Next
            .VarResult = CDbl(Mid$(.Contents, lOffset, lSize)) 
            If Err.Number <> 0 Then
                .VarResult = 0
                .LastError = Err.Description
            End If 

        Case ucsAct_2_Value
             .VarResult = .VarStack(.VarPos - 1)
        End Select
    End With
End Sub
