VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXpath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Auto-generated on 2.12.2018 13:41:50
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Const LOCALE_USER_DEFAULT           As Long = &H400
Private Const NORM_IGNORECASE               As Long = 1
Private Const CSTR_EQUAL                    As Long = 2

Private Declare Function CompareStringW Lib "kernel32" (ByVal Locale As Long, ByVal dwCmpFlags As Long, lpString1 As Any, ByVal cchCount1 As Long, lpString2 As Any, ByVal cchCount2 As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

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
    ParserVersion = "2.12.2018 13:41:50"
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
    With ctx
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

Private Function pvMatchString(sText As String, Optional ByVal CmpFlags As Long) As Boolean
    With ctx
        If .BufPos + Len(sText) <= .BufSize Then
            pvMatchString = CompareStringW(LOCALE_USER_DEFAULT, CmpFlags, ByVal StrPtr(sText), Len(sText), .BufData(.BufPos), Len(sText)) = CSTR_EQUAL
        End If
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

'= generated functions ===================================================

Public Function Parsestart() As Boolean
    Dim i6 As Long
    Dim p5 As Long

    With ctx
        For i6 = 0 To LNG_MAXINT
            p5 = .BufPos
            If Not (ParseLocationPath()) Then
                .BufPos = p5
                Exit For
            End If
        Next
        If i6 <> 0 Then
            Parsestart = True
        End If
    End With
End Function

Public Function ParseLocationPath() As Boolean
    Dim p15 As Long

    With ctx
        p15 = .BufPos
        If ParseRelativeLocationPath() Then
            ParseLocationPath = True
            Exit Function
        Else
            .BufPos = p15
        End If
        If ParseAbsoluteLocationPath() Then
            ParseLocationPath = True
            Exit Function
        Else
            .BufPos = p15
        End If
    End With
End Function

Private Sub Parse_()
    Dim p9 As Long

    With ctx
        Do
            p9 = .BufPos
            If Not (ParseS()) Then
                .BufPos = p9
                Exit Do
            End If
        Loop
    End With
End Sub

Public Function ParseRelativeLocationPath() As Boolean
    Dim p30 As Long

    With ctx
        If ParseStep() Then
            Do
                p30 = .BufPos
                Call Parse_
                If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 47 Then ' "//"
                    .BufPos = .BufPos + 2
                Else
                    If .BufData(.BufPos) = 47 Then  ' "/"
                        .BufPos = .BufPos + 1
                    Else
                        .BufPos = p30
                        Exit Do
                    End If
                End If
                Call Parse_
                If Not (ParseStep()) Then
                    .BufPos = p30
                    Exit Do
                End If
            Loop
            Call pvSetAdvance
            ParseRelativeLocationPath = True
        End If
    End With
End Function

Public Function ParseAbsoluteLocationPath() As Boolean
    Dim p23 As Long
    Dim p21 As Long

    With ctx
        p23 = .BufPos
        If ParseAbbreviatedAbsoluteLocationPath() Then
            Call pvSetAdvance
            ParseAbsoluteLocationPath = True
            Exit Function
        Else
            .BufPos = p23
        End If
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            Call Parse_
            p21 = .BufPos
            If Not (ParseRelativeLocationPath()) Then
                .BufPos = p21
            End If
            Call pvSetAdvance
            ParseAbsoluteLocationPath = True
        End If
    End With
End Function

Public Function ParseAbbreviatedAbsoluteLocationPath() As Boolean
    With ctx
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 47 Then ' "//"
            .BufPos = .BufPos + 2
            Call Parse_
            If ParseRelativeLocationPath() Then
                Call pvSetAdvance
                ParseAbbreviatedAbsoluteLocationPath = True
            End If
        End If
    End With
End Function

Public Function ParseStep() As Boolean
    Dim p48 As Long
    Dim p44 As Long

    With ctx
        p48 = .BufPos
        Call ParseAxisSpecifier
        Call Parse_
        If ParseNodeTest() Then
            Do
                p44 = .BufPos
                Call Parse_
                If Not (ParsePredicate()) Then
                    .BufPos = p44
                    Exit Do
                End If
            Loop
            ParseStep = True
            Exit Function
        Else
            .BufPos = p48
        End If
        If ParseAbbreviatedStep() Then
            ParseStep = True
            Exit Function
        Else
            .BufPos = p48
        End If
    End With
End Function

Public Sub ParseAxisSpecifier()
    Dim p56 As Long

    With ctx
        p56 = .BufPos
        If ParseAxisName() Then
            Call Parse_
            If .BufData(.BufPos) = 58 And .BufData(.BufPos + 1) = 58 Then ' "::"
                .BufPos = .BufPos + 2
                Call pvSetAdvance
            Else
                .BufPos = p56
            End If
        Else
            .BufPos = p56
        End If
        Call ParseAbbreviatedAxisSpecifier
        Call pvSetAdvance
    End With
End Sub

Public Function ParseNodeTest() As Boolean
    Dim p87 As Long

    With ctx
        p87 = .BufPos
        If ParseNodeType() Then
            Call Parse_
            If .BufData(.BufPos) = 40 Then          ' "("
                .BufPos = .BufPos + 1
                Call Parse_
                If .BufData(.BufPos) = 41 Then      ' ")"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseNodeTest = True
                    Exit Function
                Else
                    .BufPos = p87
                End If
            Else
                .BufPos = p87
            End If
        Else
            .BufPos = p87
        End If
        If pvMatchString("processing-instruction") Then ' "processing-instruction"
            .BufPos = .BufPos + 22
            Call Parse_
            If .BufData(.BufPos) = 40 Then          ' "("
                .BufPos = .BufPos + 1
                Call Parse_
                If ParseLiteral() Then
                    Call Parse_
                    If .BufData(.BufPos) = 41 Then  ' ")"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseNodeTest = True
                        Exit Function
                    Else
                        .BufPos = p87
                    End If
                Else
                    .BufPos = p87
                End If
            Else
                .BufPos = p87
            End If
        End If
        If ParseNameTest() Then
            Call pvSetAdvance
            ParseNodeTest = True
            Exit Function
        Else
            .BufPos = p87
        End If
    End With
End Function

Public Function ParsePredicate() As Boolean
    With ctx
        If .BufData(.BufPos) = 91 Then              ' "["
            .BufPos = .BufPos + 1
            Call Parse_
            If ParseOrExpr() Then
                Call Parse_
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParsePredicate = True
                End If
            End If
        End If
    End With
End Function

Public Function ParseAbbreviatedStep() As Boolean
    With ctx
        If .BufData(.BufPos) = 46 And .BufData(.BufPos + 1) = 46 Then ' ".."
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseAbbreviatedStep = True
            Exit Function
        End If
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAbbreviatedStep = True
        End If
    End With
End Function

Public Function ParseAxisName() As Boolean
    With ctx
        If pvMatchString("ancestor-or-self") Then   ' "ancestor-or-self"
            .BufPos = .BufPos + 16
            Call pvSetAdvance
            ParseAxisName = True
            Exit Function
        End If
        If pvMatchString("ancestor") Then           ' "ancestor"
            .BufPos = .BufPos + 8
            Call pvSetAdvance
            ParseAxisName = True
            Exit Function
        End If
        If pvMatchString("attribute") Then          ' "attribute"
            .BufPos = .BufPos + 9
            Call pvSetAdvance
            ParseAxisName = True
            Exit Function
        End If
        If pvMatchString("child") Then              ' "child"
            .BufPos = .BufPos + 5
            Call pvSetAdvance
            ParseAxisName = True
            Exit Function
        End If
        If pvMatchString("descendant-or-self") Then ' "descendant-or-self"
            .BufPos = .BufPos + 18
            Call pvSetAdvance
            ParseAxisName = True
            Exit Function
        End If
        If pvMatchString("descendant") Then         ' "descendant"
            .BufPos = .BufPos + 10
            Call pvSetAdvance
            ParseAxisName = True
            Exit Function
        End If
        If pvMatchString("following-sibling") Then  ' "following-sibling"
            .BufPos = .BufPos + 17
            Call pvSetAdvance
            ParseAxisName = True
            Exit Function
        End If
        If pvMatchString("following") Then          ' "following"
            .BufPos = .BufPos + 9
            Call pvSetAdvance
            ParseAxisName = True
            Exit Function
        End If
        If pvMatchString("namespace") Then          ' "namespace"
            .BufPos = .BufPos + 9
            Call pvSetAdvance
            ParseAxisName = True
            Exit Function
        End If
        If pvMatchString("parent") Then             ' "parent"
            .BufPos = .BufPos + 6
            Call pvSetAdvance
            ParseAxisName = True
            Exit Function
        End If
        If pvMatchString("preceding-sibling") Then  ' "preceding-sibling"
            .BufPos = .BufPos + 17
            Call pvSetAdvance
            ParseAxisName = True
            Exit Function
        End If
        If pvMatchString("preceding") Then          ' "preceding"
            .BufPos = .BufPos + 9
            Call pvSetAdvance
            ParseAxisName = True
            Exit Function
        End If
        If pvMatchString("self") Then               ' "self"
            .BufPos = .BufPos + 4
            Call pvSetAdvance
            ParseAxisName = True
        End If
    End With
End Function

Public Sub ParseAbbreviatedAxisSpecifier()
    With ctx
        If .BufData(.BufPos) = 64 Then              ' "@"
            .BufPos = .BufPos + 1
        End If
        Call pvSetAdvance
    End With
End Sub

Public Function ParseNodeType() As Boolean
    With ctx
        If pvMatchString("comment") Then            ' "comment"
            .BufPos = .BufPos + 7
            Call pvSetAdvance
            ParseNodeType = True
            Exit Function
        End If
        If pvMatchString("text") Then               ' "text"
            .BufPos = .BufPos + 4
            Call pvSetAdvance
            ParseNodeType = True
            Exit Function
        End If
        If pvMatchString("processing-instruction") Then ' "processing-instruction"
            .BufPos = .BufPos + 22
            Call pvSetAdvance
            ParseNodeType = True
            Exit Function
        End If
        If pvMatchString("node") Then               ' "node"
            .BufPos = .BufPos + 4
            Call pvSetAdvance
            ParseNodeType = True
        End If
    End With
End Function

Public Function ParseLiteral() As Boolean
    Dim p273 As Long

    With ctx
        p273 = .BufPos
        If .BufData(.BufPos) = 34 Then              ' "\""
            .BufPos = .BufPos + 1
            Do
                If .BufData(.BufPos) <> 34 And .BufPos < .BufSize Then ' """
                    .BufPos = .BufPos + 1
                Else
                    Exit Do
                End If
            Loop
            If .BufData(.BufPos) = 34 Then          ' "\""
                .BufPos = .BufPos + 1
                Call pvSetAdvance
                ParseLiteral = True
                Exit Function
            Else
                .BufPos = p273
            End If
        End If
        If .BufData(.BufPos) = 34 Then              ' "\""
            .BufPos = .BufPos + 1
            Do
                If .BufData(.BufPos) <> 34 And .BufPos < .BufSize Then ' """
                    .BufPos = .BufPos + 1
                Else
                    Exit Do
                End If
            Loop
            If .BufData(.BufPos) = 34 Then          ' "\""
                .BufPos = .BufPos + 1
                Call pvSetAdvance
                ParseLiteral = True
                Exit Function
            Else
                .BufPos = p273
            End If
        End If
    End With
End Function

Public Function ParseNameTest() As Boolean
    Dim p323 As Long

    With ctx
        p323 = .BufPos
        If .BufData(.BufPos) = 42 Then              ' "*"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameTest = True
            Exit Function
        End If
        If ParseName() Then
            If .BufData(.BufPos) = 58 Then          ' ":"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 42 Then      ' "*"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseNameTest = True
                    Exit Function
                Else
                    .BufPos = p323
                End If
            Else
                .BufPos = p323
            End If
        Else
            .BufPos = p323
        End If
        If ParseQName() Then
            Call pvSetAdvance
            ParseNameTest = True
            Exit Function
        Else
            .BufPos = p323
        End If
    End With
End Function

Public Function ParseOrExpr() As Boolean
    Dim p177 As Long

    With ctx
        If ParseAndExpr() Then
            Do
                p177 = .BufPos
                Call Parse_
                If .BufData(.BufPos) = 111 And .BufData(.BufPos + 1) = 114 Then ' "or"
                    .BufPos = .BufPos + 2
                Else
                    .BufPos = p177
                    Exit Do
                End If
                Call Parse_
                If Not (ParseAndExpr()) Then
                    .BufPos = p177
                    Exit Do
                End If
            Loop
            Call pvSetAdvance
            ParseOrExpr = True
        End If
    End With
End Function

Public Function ParsePrimaryExpr() As Boolean
    Dim p119 As Long

    With ctx
        p119 = .BufPos
        If ParseVariableReference() Then
            Call pvSetAdvance
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p119
        End If
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call Parse_
            If ParseOrExpr() Then
                Call Parse_
                If .BufData(.BufPos) = 41 Then      ' ")"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParsePrimaryExpr = True
                    Exit Function
                Else
                    .BufPos = p119
                End If
            Else
                .BufPos = p119
            End If
        End If
        If ParseLiteral() Then
            Call pvSetAdvance
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p119
        End If
        If ParseNumber() Then
            Call pvSetAdvance
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p119
        End If
        If ParseFunctionCall() Then
            Call pvSetAdvance
            ParsePrimaryExpr = True
            Exit Function
        Else
            .BufPos = p119
        End If
    End With
End Function

Public Function ParseVariableReference() As Boolean
    With ctx
        If .BufData(.BufPos) = 36 Then              ' "$"
            .BufPos = .BufPos + 1
            If ParseQName() Then
                Call pvSetAdvance
                ParseVariableReference = True
            End If
        End If
    End With
End Function

Public Function ParseNumber() As Boolean
    Dim p285 As Long
    Dim p279 As Long
    Dim p277 As Long

    With ctx
        p285 = .BufPos
        If ParseDigits() Then
            p279 = .BufPos
            If .BufData(.BufPos) = 46 Then          ' "."
                .BufPos = .BufPos + 1
                p277 = .BufPos
                If Not (ParseDigits()) Then
                    .BufPos = p277
                End If
            End If
            Call pvSetAdvance
            ParseNumber = True
            Exit Function
        Else
            .BufPos = p285
        End If
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
            If ParseDigits() Then
                Call pvSetAdvance
                ParseNumber = True
                Exit Function
            Else
                .BufPos = p285
            End If
        End If
    End With
End Function

Public Function ParseFunctionCall() As Boolean
    Dim p133 As Long
    Dim p136 As Long

    With ctx
        If ParseFunctionName() Then
            Call Parse_
            If .BufData(.BufPos) = 40 Then          ' "("
                .BufPos = .BufPos + 1
                p133 = .BufPos
                Call Parse_
                If ParseOrExpr() Then
                    Do
                        p136 = .BufPos
                        Call Parse_
                        If .BufData(.BufPos) = 44 Then ' ","
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p136
                            Exit Do
                        End If
                        Call Parse_
                        If Not (ParseOrExpr()) Then
                            .BufPos = p136
                            Exit Do
                        End If
                    Loop
                Else
                    .BufPos = p133
                End If
                Call Parse_
                If .BufData(.BufPos) = 41 Then      ' ")"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseFunctionCall = True
                End If
            End If
        End If
    End With
End Function

Public Function ParseFunctionName() As Boolean
    Dim p310 As Long
    Dim e310 As String

    With ctx
        p310 = .BufPos
        e310 = .LastExpected
        If Not (ParseNodeType()) Then
            .BufPos = p310
            .LastExpected = e310
            If ParseQName() Then
                ParseFunctionName = True
            End If
        End If
    End With
End Function

Public Function ParseUnionExpr() As Boolean
    Dim p149 As Long

    With ctx
        If ParsePathExpr() Then
            Do
                p149 = .BufPos
                Call Parse_
                If .BufData(.BufPos) = 124 Then     ' "|"
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p149
                    Exit Do
                End If
                Call Parse_
                If Not (ParsePathExpr()) Then
                    .BufPos = p149
                    Exit Do
                End If
            Loop
            Call pvSetAdvance
            ParseUnionExpr = True
        End If
    End With
End Function

Public Function ParsePathExpr() As Boolean
    Dim p166 As Long
    Dim p160 As Long

    With ctx
        p166 = .BufPos
        If ParseFilterExpr() Then
            p160 = .BufPos
            Call Parse_
            If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 47 Then ' "//"
                .BufPos = .BufPos + 2
                Call Parse_
                If ParseRelativeLocationPath() Then
                    GoTo L8
                Else
                    .BufPos = p160
                End If
                GoTo L9
            End If
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                Call Parse_
                If ParseRelativeLocationPath() Then
                    GoTo L8
                Else
                    .BufPos = p160
                End If
                GoTo L9
            End If
            .BufPos = p160
L9:
L8:
            Call pvSetAdvance
            ParsePathExpr = True
            Exit Function
        Else
            .BufPos = p166
        End If
        If ParseLocationPath() Then
            Call pvSetAdvance
            ParsePathExpr = True
            Exit Function
        Else
            .BufPos = p166
        End If
    End With
End Function

Public Function ParseFilterExpr() As Boolean
    Dim p170 As Long

    With ctx
        If ParsePrimaryExpr() Then
            Do
                p170 = .BufPos
                Call Parse_
                If Not (ParsePredicate()) Then
                    .BufPos = p170
                    Exit Do
                End If
            Loop
            ParseFilterExpr = True
        End If
    End With
End Function

Public Function ParseAndExpr() As Boolean
    Dim p186 As Long

    With ctx
        If ParseEqualityExpr() Then
            Do
                p186 = .BufPos
                Call Parse_
                If pvMatchString("and") Then        ' "and"
                    .BufPos = .BufPos + 3
                Else
                    .BufPos = p186
                    Exit Do
                End If
                Call Parse_
                If Not (ParseEqualityExpr()) Then
                    .BufPos = p186
                    Exit Do
                End If
            Loop
            Call pvSetAdvance
            ParseAndExpr = True
        End If
    End With
End Function

Public Function ParseEqualityExpr() As Boolean
    Dim p197 As Long

    With ctx
        If ParseRelationalExpr() Then
            Do
                p197 = .BufPos
                Call Parse_
                If .BufData(.BufPos) = 33 And .BufData(.BufPos + 1) = 61 Then ' "!="
                    .BufPos = .BufPos + 2
                Else
                    If .BufData(.BufPos) = 61 Then  ' "="
                        .BufPos = .BufPos + 1
                    Else
                        .BufPos = p197
                        Exit Do
                    End If
                End If
                Call Parse_
                If Not (ParseRelationalExpr()) Then
                    .BufPos = p197
                    Exit Do
                End If
            Loop
            Call pvSetAdvance
            ParseEqualityExpr = True
        End If
    End With
End Function

Public Function ParseRelationalExpr() As Boolean
    Dim p210 As Long

    With ctx
        If ParseAdditiveExpr() Then
            Do
                p210 = .BufPos
                Call Parse_
                If .BufData(.BufPos) = 60 And .BufData(.BufPos + 1) = 61 Then ' "<="
                    .BufPos = .BufPos + 2
                Else
                    If .BufData(.BufPos) = 60 Then  ' "<"
                        .BufPos = .BufPos + 1
                    Else
                        If .BufData(.BufPos) = 62 And .BufData(.BufPos + 1) = 61 Then ' ">="
                            .BufPos = .BufPos + 2
                        Else
                            If .BufData(.BufPos) = 62 Then ' ">"
                                .BufPos = .BufPos + 1
                            Else
                                .BufPos = p210
                                Exit Do
                            End If
                        End If
                    End If
                End If
                Call Parse_
                If Not (ParseAdditiveExpr()) Then
                    .BufPos = p210
                    Exit Do
                End If
            Loop
            Call pvSetAdvance
            ParseRelationalExpr = True
        End If
    End With
End Function

Public Function ParseAdditiveExpr() As Boolean
    Dim p221 As Long

    With ctx
        If ParseMultiplicativeExpr() Then
            Do
                p221 = .BufPos
                If .BufData(.BufPos) = 43 Then      ' "+"
                    .BufPos = .BufPos + 1
                Else
                    If .BufData(.BufPos) = 45 Then  ' "-"
                        .BufPos = .BufPos + 1
                    Else
                        .BufPos = p221
                        Exit Do
                    End If
                End If
                If Not (ParseMultiplicativeExpr()) Then
                    .BufPos = p221
                    Exit Do
                End If
            Loop
            Call pvSetAdvance
            ParseAdditiveExpr = True
        End If
    End With
End Function

Public Function ParseMultiplicativeExpr() As Boolean
    Dim p232 As Long

    With ctx
        If ParseUnaryExpr() Then
            Do
                p232 = .BufPos
                Call Parse_
                If .BufData(.BufPos) = 42 Then      ' "*"
                    .BufPos = .BufPos + 1
                Else
                    If pvMatchString("div") Then    ' "div"
                        .BufPos = .BufPos + 3
                    Else
                        If pvMatchString("mod") Then ' "mod"
                            .BufPos = .BufPos + 3
                        Else
                            .BufPos = p232
                            Exit Do
                        End If
                    End If
                End If
                Call Parse_
                If Not (ParseUnaryExpr()) Then
                    .BufPos = p232
                    Exit Do
                End If
            Loop
            Call pvSetAdvance
            ParseMultiplicativeExpr = True
        End If
    End With
End Function

Public Function ParseUnaryExpr() As Boolean
    Dim p242 As Long

    With ctx
        p242 = .BufPos
        If ParseUnionExpr() Then
            Call pvSetAdvance
            ParseUnaryExpr = True
            Exit Function
        Else
            .BufPos = p242
        End If
        If .BufData(.BufPos) = 45 Then              ' "-"
            .BufPos = .BufPos + 1
            Call Parse_
            If ParseUnaryExpr() Then
                Call pvSetAdvance
                ParseUnaryExpr = True
                Exit Function
            Else
                .BufPos = p242
            End If
        End If
    End With
End Function

Public Function ParseOperator() As Boolean
    Dim p291 As Long

    With ctx
        p291 = .BufPos
        If ParseOperatorName() Then
            Call pvSetAdvance
            ParseOperator = True
            Exit Function
        Else
            .BufPos = p291
        End If
        If .BufData(.BufPos) = 42 Then              ' "*"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseOperator = True
            Exit Function
        End If
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 47 Then ' "//"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseOperator = True
            Exit Function
        End If
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseOperator = True
            Exit Function
        End If
        If .BufData(.BufPos) = 124 Then             ' "|"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseOperator = True
            Exit Function
        End If
        If .BufData(.BufPos) = 43 Then              ' "+"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseOperator = True
            Exit Function
        End If
        If .BufData(.BufPos) = 45 Then              ' "-"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseOperator = True
            Exit Function
        End If
        If .BufData(.BufPos) = 61 Then              ' "="
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseOperator = True
            Exit Function
        End If
        If .BufData(.BufPos) = 33 And .BufData(.BufPos + 1) = 61 Then ' "!="
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseOperator = True
            Exit Function
        End If
        If .BufData(.BufPos) = 60 And .BufData(.BufPos + 1) = 61 Then ' "<="
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseOperator = True
            Exit Function
        End If
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseOperator = True
            Exit Function
        End If
        If .BufData(.BufPos) = 62 And .BufData(.BufPos + 1) = 61 Then ' ">="
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseOperator = True
            Exit Function
        End If
        If .BufData(.BufPos) = 62 Then              ' ">"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseOperator = True
        End If
    End With
End Function

Public Function ParseDigits() As Boolean
    Dim i287 As Long

    With ctx
        For i287 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i287 <> 0 Then
            Call pvSetAdvance
            ParseDigits = True
        End If
    End With
End Function

Public Function ParseOperatorName() As Boolean
    With ctx
        If pvMatchString("and") Then                ' "and"
            .BufPos = .BufPos + 3
            Call pvSetAdvance
            ParseOperatorName = True
            Exit Function
        End If
        If .BufData(.BufPos) = 111 And .BufData(.BufPos + 1) = 114 Then ' "or"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseOperatorName = True
            Exit Function
        End If
        If pvMatchString("mod") Then                ' "mod"
            .BufPos = .BufPos + 3
            Call pvSetAdvance
            ParseOperatorName = True
            Exit Function
        End If
        If pvMatchString("div") Then                ' "div"
            .BufPos = .BufPos + 3
            Call pvSetAdvance
            ParseOperatorName = True
        End If
    End With
End Function

Public Function ParseQName() As Boolean
    Dim p338 As Long

    With ctx
        p338 = .BufPos
        If ParsePrefixedName() Then
            ParseQName = True
            Exit Function
        Else
            .BufPos = p338
        End If
        If ParseName() Then
            ParseQName = True
            Exit Function
        Else
            .BufPos = p338
        End If
    End With
End Function

Private Function ParseS() As Boolean
    Dim i333 As Long

    With ctx
        For i333 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 32, 13, 10, 9                      ' [ \r\n\t]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i333 <> 0 Then
            Call pvSetAdvance
            ParseS = True
        End If
    End With
End Function

Public Function ParsePrefixedName() As Boolean
    With ctx
        If ParseName() Then
            If .BufData(.BufPos) = 58 Then          ' ":"
                .BufPos = .BufPos + 1
                If ParseName() Then
                    Call pvSetAdvance
                    ParsePrefixedName = True
                End If
            End If
        End If
    End With
End Function

Public Function ParseName() As Boolean
    Dim p353 As Long

    With ctx
        If ParseNameStartChar() Then
            Do
                p353 = .BufPos
                If Not (ParseNameChar()) Then
                    .BufPos = p353
                    Exit Do
                End If
            Loop
            ParseName = True
        End If
    End With
End Function

Public Function ParseNameStartChar() As Boolean
    With ctx
        Select Case .BufData(.BufPos)
        Case 65 To 90                               ' [A-Z]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameStartChar = True
            Exit Function
        End Select
        If .BufData(.BufPos) = 95 Then              ' "_"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameStartChar = True
            Exit Function
        End If
        Select Case .BufData(.BufPos)
        Case 97 To 122                              ' [a-z]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameStartChar = True
            Exit Function
        End Select
        Select Case .BufData(.BufPos)
        Case 65 To 79                               ' [\u00C0-\u00D6]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameStartChar = True
            Exit Function
        End Select
        Select Case .BufData(.BufPos)
        Case 79 To 111                              ' [\u00D8-\u00F6]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameStartChar = True
            Exit Function
        End Select
        Select Case .BufData(.BufPos)
        Case 111 To 63                              ' [\u00F8-\u02FF]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameStartChar = True
            Exit Function
        End Select
        Select Case .BufData(.BufPos)
        Case 63 To 63                               ' [\u0370-\u037D]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameStartChar = True
            Exit Function
        End Select
        Select Case .BufData(.BufPos)
        Case 63 To 63                               ' [\u037F-\u1FFF]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameStartChar = True
            Exit Function
        End Select
        Select Case .BufData(.BufPos)
        Case 63 To 63                               ' [\u200C-\u200D]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameStartChar = True
            Exit Function
        End Select
        Select Case .BufData(.BufPos)
        Case 63 To 63                               ' [\u2070-\u218F]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameStartChar = True
            Exit Function
        End Select
        Select Case .BufData(.BufPos)
        Case 63 To 63                               ' [\u2C00-\u2FEF]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameStartChar = True
            Exit Function
        End Select
        Select Case .BufData(.BufPos)
        Case 63 To 63                               ' [\u3001-\uD7FF]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameStartChar = True
            Exit Function
        End Select
        Select Case .BufData(.BufPos)
        Case 63 To 63                               ' [\uF900-\uFDCF]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameStartChar = True
            Exit Function
        End Select
        Select Case .BufData(.BufPos)
        Case 63 To 63                               ' [\uFDF0-\uFFFD]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameStartChar = True
        End Select
    End With
End Function

Public Function ParseNameChar() As Boolean
    Dim p373 As Long

    With ctx
        p373 = .BufPos
        If ParseNameStartChar() Then
            Call pvSetAdvance
            ParseNameChar = True
            Exit Function
        Else
            .BufPos = p373
        End If
        If .BufData(.BufPos) = 45 Then              ' "-"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameChar = True
            Exit Function
        End If
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameChar = True
            Exit Function
        End If
        Select Case .BufData(.BufPos)
        Case 48 To 57                               ' [0-9]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameChar = True
            Exit Function
        End Select
        Select Case .BufData(.BufPos)
        Case 183                                    ' [\u00B7]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameChar = True
            Exit Function
        End Select
        Select Case .BufData(.BufPos)
        Case 63 To 63                               ' [\u0300-\u036F]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameChar = True
            Exit Function
        End Select
        Select Case .BufData(.BufPos)
        Case 63 To 63                               ' [\u203F-\u2040]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNameChar = True
        End Select
    End With
End Function

