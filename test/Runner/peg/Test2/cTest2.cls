VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTest2"
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
        If ParsePreprocessor() Then
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

Public Function ParsePreprocessor() As Boolean
    Dim i15 As Long
    Dim p6 As Long

    With ctx
        For i15 = 0 To LNG_MAXINT
            p6 = .BufPos
            If Not (ParseLineComment()) Then
                .BufPos = p6
                If Not (ParseBlockComment()) Then
                    .BufPos = p6
                    If Not (Parsetext()) Then
                        .BufPos = p6
                        If Not (ParseScope()) Then
                            .BufPos = p6
                            If Not (Parsejunk()) Then
                                .BufPos = p6
                                If .BufPos < .BufSize Then
                                    .BufPos = .BufPos + 1
                                Else
                                    .BufPos = p6
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
        If i15 <> 0 Then
            ParsePreprocessor = True
        End If
    End With
End Function

Public Function ParseLineComment() As Boolean
    Dim p44 As Long
    Dim p42 As Long
    Dim e42 As String

    With ctx
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 47 Then ' "//"
            .BufPos = .BufPos + 2
            Do
                p44 = .BufPos
                p42 = .BufPos
                e42 = .LastExpected
                If ParseEndOfLine() Then
                    .BufPos = p44
                    Exit Do
                Else
                    .BufPos = p42
                    .LastExpected = e42
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p44
                    Exit Do
                End If
            Loop
            If ParseEndOfLine() Then
                Call pvSetAdvance
                ParseLineComment = True
            End If
        End If
    End With
End Function

Public Function ParseBlockComment() As Boolean
    Dim p52 As Long
    Dim p50 As Long

    With ctx
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 42 Then ' "/*"
            .BufPos = .BufPos + 2
            Do
                p52 = .BufPos
                p50 = .BufPos
                If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                    .BufPos = p52
                    Exit Do
                Else
                    .BufPos = p50
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p52
                    Exit Do
                End If
            Loop
            If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                .BufPos = .BufPos + 2
                Call pvSetAdvance
                ParseBlockComment = True
            End If
        End If
    End With
End Function

Public Function Parsetext() As Boolean
    If Parsequote() Then
        Call ParseText2
        If Parsequote() Then
            Parsetext = True
        End If
    End If
End Function

Public Function ParseScope() As Boolean
    Dim p74 As Long
    Dim i73 As Long
    Dim p72 As Long
    Dim p70 As Long

    With ctx
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            p74 = .BufPos
            If ParseScope() Then
                If .BufData(.BufPos) = 125 Then     ' "}"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseScope = True
                End If
            Else
                .BufPos = p74
            End If
            For i73 = 0 To LNG_MAXINT
                p72 = .BufPos
                p70 = .BufPos
                Select Case .BufData(.BufPos)
                Case 123, 125                       ' [{}]
                    .BufPos = p72
                    Exit For
                Case Else
                    .BufPos = p70
                End Select
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p72
                    Exit For
                End If
            Next
            If i73 <> 0 Then
                If .BufData(.BufPos) = 125 Then     ' "}"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseScope = True
                End If
            End If
        End If
    End With
End Function

Public Function Parsejunk() As Boolean
    Dim i134 As Long
    Dim p133 As Long
    Dim p131 As Long
    Dim p135 As Long

    With ctx
        For i134 = 0 To LNG_MAXINT
            p133 = .BufPos
            p131 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 123, 125, 34, 47, 42, 58, 40, 41 ' [a-zA-Z_{}"/*:()]
                .BufPos = p133
                Exit For
            Case Else
                .BufPos = p131
            End Select
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
            Else
                .BufPos = p133
                Exit For
            End If
        Next
        If i134 <> 0 Then
            Do
                p135 = .BufPos
                If Not (ParseSpacing()) Then
                    .BufPos = p135
                    Exit Do
                End If
            Loop
            Call pvSetAdvance
            Parsejunk = True
        End If
    End With
End Function

Public Function ParseFile() As Boolean
    Dim i21 As Long
    Dim p19 As Long

    With ctx
        For i21 = 0 To LNG_MAXINT
            p19 = .BufPos
            If Not (ParseScope()) Then
                .BufPos = p19
                If Not (Parsejunk()) Then
                    .BufPos = p19
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                    Else
                        .BufPos = p19
                        Exit For
                    End If
                End If
            End If
        Next
        If i21 <> 0 Then
            ParseFile = True
        End If
    End With
End Function

Public Function ParseStatement() As Boolean
    Dim p37 As Long

    With ctx
        If ParseSpacing() Then
            p37 = .BufPos
            If ParseExpression() Then
                If .BufData(.BufPos) = 59 Then      ' ";"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseStatement = True
                    Exit Function
                Else
                    .BufPos = p37
                End If
            Else
                .BufPos = p37
            End If
            If ParseFile() Then
                Call pvSetAdvance
                ParseStatement = True
                Exit Function
            Else
                .BufPos = p37
            End If
        End If
    End With
End Function

Public Function ParseEndOfFile() As Boolean
    Dim p140 As Long

    With ctx
        p140 = .BufPos
        If Not (.BufPos < .BufSize) Then
            .BufPos = p140
            ParseEndOfFile = True
        End If
    End With
End Function

Public Function ParseSpacing() As Boolean
    Dim i118 As Long
    Dim p115 As Long

    With ctx
        For i118 = 0 To LNG_MAXINT
            p115 = .BufPos
            If Not (ParseLineComment()) Then
                .BufPos = p115
                If Not (ParseBlockComment()) Then
                    .BufPos = p115
                    Select Case .BufData(.BufPos)
                    Case 32, 9, 10, 13              ' [ \t\n\r]
                        .BufPos = .BufPos + 1
                    Case Else
                        If Not (ParseEndOfLine()) Then
                            .BufPos = p115
                            .BufPos = p115
                            Exit For
                        End If
                    End Select
                End If
            End If
        Next
        If i118 <> 0 Then
            Call pvSetAdvance
            ParseSpacing = True
        End If
    End With
End Function

Public Function ParseExpression() As Boolean
    If ParseType() Then
        If ParseIdentifier() Then
            ParseExpression = True
        End If
    End If
End Function

Public Function ParseEndOfLine() As Boolean
    Dim p120 As Long

    With ctx
        p120 = .BufPos
        If Not (.BufData(.BufPos) = 92) Then        ' "\\"
            .BufPos = p120
            If .BufData(.BufPos) = 13 And .BufData(.BufPos + 1) = 10 Then ' "\r\n"
                .BufPos = .BufPos + 2
                Call pvSetAdvance
                ParseEndOfLine = True
                Exit Function
            End If
            If .BufData(.BufPos) = 10 Then          ' "\n"
                .BufPos = .BufPos + 1
                Call pvSetAdvance
                ParseEndOfLine = True
                Exit Function
            End If
            If .BufData(.BufPos) = 13 Then          ' "\r"
                .BufPos = .BufPos + 1
                Call pvSetAdvance
                ParseEndOfLine = True
            End If
        End If
    End With
End Function

Public Function Parsequote() As Boolean
    Dim p127 As Long

    With ctx
        p127 = .BufPos
        If Not (.BufData(.BufPos) = 92) Then        ' "\\"
            .BufPos = p127
            If .BufData(.BufPos) = 34 Then          ' """
                .BufPos = .BufPos + 1
                Call pvSetAdvance
                Parsequote = True
            End If
        End If
    End With
End Function

Public Sub ParseText2()
    Dim p65 As Long
    Dim p63 As Long
    Dim e63 As String

    With ctx
        Do
            p65 = .BufPos
            p63 = .BufPos
            e63 = .LastExpected
            If Parsequote() Then
                .BufPos = p65
                Exit Do
            Else
                .BufPos = p63
                .LastExpected = e63
            End If
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
            Else
                .BufPos = p65
                Exit Do
            End If
        Loop
    End With
End Sub

Public Function ParseType() As Boolean
    Dim i102 As Long
    Dim p88 As Long
    Dim p84 As Long
    Dim p91 As Long
    Dim i101 As Long
    Dim p95 As Long

    With ctx
        For i102 = 0 To LNG_MAXINT
            p88 = .BufPos
            p84 = .BufPos
            If ParseSpacing() Then
                If pvMatchString("const") Then      ' "const"
                    .BufPos = .BufPos + 5
                    If Not (ParseSpacing()) Then
                        .BufPos = p84
                    End If
                Else
                    .BufPos = p84
                End If
            Else
                .BufPos = p84
            End If
            If Not (ParseIdentifier()) Then
                .BufPos = p88
                Exit For
            End If
            p91 = .BufPos
            If .BufData(.BufPos) = 58 And .BufData(.BufPos + 1) = 58 Then ' "::"
                .BufPos = .BufPos + 2
                If Not (ParseIdentifier()) Then
                    .BufPos = p91
                End If
            End If
            For i101 = 0 To LNG_MAXINT
                p95 = .BufPos
                If Not (ParseSpacing()) Then
                    .BufPos = p95
                    If .BufData(.BufPos) = 42 Then  ' "*"
                        .BufPos = .BufPos + 1
                    Else
                        If .BufData(.BufPos) = 38 Then ' "&"
                            .BufPos = .BufPos + 1
                        Else
                            If ParseSpacing() Then
                                If pvMatchString("const") Then ' "const"
                                    .BufPos = .BufPos + 5
                                    If Not (ParseSpacing()) Then
                                        .BufPos = p95
                                        .BufPos = p95
                                        Exit For
                                    End If
                                Else
                                    .BufPos = p95
                                    .BufPos = p95
                                    Exit For
                                End If
                            Else
                                .BufPos = p95
                                .BufPos = p95
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next
            If i101 = 0 Then
                .BufPos = p88
                Exit For
            End If
        Next
        If i102 <> 0 Then
            Call pvSetAdvance
            ParseType = True
        End If
    End With
End Function

Public Function ParseIdentifier() As Boolean
    Dim p111 As Long

    With ctx
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
            Do
                p111 = .BufPos
                If Not (ParseSpacing()) Then
                    .BufPos = p111
                    Exit Do
                End If
            Loop
            Call pvSetAdvance
            ParseIdentifier = True
        End Select
    End With
End Function

