VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Auto-generated on 8.2.2018 21:26:36
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
    ucsAct_2_Definition
    ucsAct_1_Definition
    ucsAct_1_Expression
    ucsAct_2_Sequence
    ucsAct_1_Sequence
    ucsAct_1_Prefix
    ucsAct_2_Prefix
    ucsAct_3_Prefix
    ucsAct_1_Suffix
    ucsAct_2_Suffix
    ucsAct_3_Suffix
    ucsAct_2_Primary
    ucsAct_1_Primary
    ucsAct_3_Primary
    ucsAct_4_Primary
    ucsAct_5_Primary
    ucsAct_6_Primary
    ucsAct_7_Primary
    ucsAct_8_Primary
    ucsAct_9_Primary
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
    UserData            As cTree
End Type

Private ctx                     As UcsParserType

'=========================================================================
' Properties
'=========================================================================

Property Get LastError() As String
    LastError = ctx.LastError
End Property

Property Get ParserVersion() As String
    ParserVersion = "8.2.2018 21:26:36"
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Match(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As cTree) As Long
    If BeginMatch(sSubject, StartPos, UserData) Then
        If ParseGrammar() Then
            Match = EndMatch()
        End If
    End If
End Function

Public Function BeginMatch(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As cTree) As Boolean
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
    BeginMatch = True
End Function

Public Function EndMatch() As Long
    Dim lIdx            As Long
    Dim uEmpty          As UcsParserType
    
    With ctx
        For lIdx = 0 To .ThunkPos - 1
            With .ThunkData(lIdx)
                pvImplAction .Action, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin
            End With
        Next
        EndMatch = .BufPos
    End With
    ctx = uEmpty
End Function

Private Sub pvPushAction(ByVal eAction As UcsParserActionsEnum)
    With ctx
        If UBound(.ThunkData) < .ThunkPos Then
            ReDim Preserve .ThunkData(0 To 2 * UBound(.ThunkData)) As UcsParserThunkType
        End If
        With .ThunkData(.ThunkPos)
            .Action = eAction
            .CaptureBegin = ctx.CaptureBegin
            .CaptureEnd = ctx.CaptureEnd
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

Public Function ParseGrammar() As Boolean
    Dim p7 As Long
    Dim q7 As Long
    Dim i6 As Long

    With ctx
        p7 = .BufPos
        q7 = .ThunkPos
        Call ParseSpacing
        For i6 = 0 To LNG_MAXINT
            If Not ParseDefinition() Then
                Exit For
            End If
        Next
        If i6 <> 0 Then
            If ParseEndOfFile() Then
                ParseGrammar = True
                Exit Function
            Else
                .BufPos = p7
                .ThunkPos = q7
            End If
        Else
            .BufPos = p7
            .ThunkPos = q7
        End If
    End With
End Function

Public Sub ParseSpacing()
    Do
        If Not ParseSpace() Then
            If Not ParseComment() Then
                Exit Do
            End If
        End If
    Loop
End Sub

Public Function ParseDefinition() As Boolean
    Dim p13 As Long
    Dim q13 As Long

    With ctx
        p13 = .BufPos
        q13 = .ThunkPos
        If ParseIdentifier() Then
            pvPushAction ucsAct_1_Definition
            If ParseAssign() Then
                Call ParseExpression
                pvPushAction ucsAct_2_Definition
                ParseDefinition = True
                Exit Function
            Else
                .BufPos = p13
                .ThunkPos = q13
            End If
        End If
    End With
End Function

Public Function ParseEndOfFile() As Boolean
    With ctx
        If Not .BufPos < .BufSize Then
            ParseEndOfFile = True
        End If
    End With
End Function

Public Function ParseIdentifier() As Boolean
    With ctx
        .CaptureBegin = .BufPos
        Select Case .BufData(.BufPos)
        Case 97 To 122, 65 To 90, 95, 92, 45        ' [a-zA-Z_\-]
            .BufPos = .BufPos + 1
            Do
                Select Case .BufData(.BufPos)
                Case 97 To 122, 65 To 90, 95, 92 To 48, 45, 57 ' [a-zA-Z_\-0-9]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit Do
                End Select
            Loop
            .CaptureEnd = .BufPos
            Call ParseSpacing
            ParseIdentifier = True
        End Select
    End With
End Function

Public Function ParseAssign() As Boolean
    If ParseLEFTARROW() Then
        ParseAssign = True
        Exit Function
    End If
    If ParseEQUAL() Then
        ParseAssign = True
    End If
End Function

Public Sub ParseExpression()
    Call ParseSequence
    Do
        If Not ParseChoice() Then
            Exit Do
        End If
        Call ParseSequence
        pvPushAction ucsAct_1_Expression
    Loop
End Sub

Public Sub ParseSequence()
    If Not ParsePrefix() Then
        pvPushAction ucsAct_2_Sequence
    End If
    Do
        If Not ParsePrefix() Then
            Exit Do
        End If
        pvPushAction ucsAct_1_Sequence
    Loop
End Sub

Public Function ParseChoice() As Boolean
    If ParseSLASH() Then
        ParseChoice = True
        Exit Function
    End If
    If ParseBAR() Then
        ParseChoice = True
    End If
End Function

Public Function ParsePrefix() As Boolean
    Dim p48 As Long
    Dim q48 As Long

    With ctx
        p48 = .BufPos
        q48 = .ThunkPos
        If ParseAND() Then
            If ParseAction() Then
                pvPushAction ucsAct_1_Prefix
                ParsePrefix = True
                Exit Function
            Else
                .BufPos = p48
                .ThunkPos = q48
            End If
        End If
        If ParseAND() Then
            If ParseSuffix() Then
                pvPushAction ucsAct_2_Prefix
                ParsePrefix = True
                Exit Function
            Else
                .BufPos = p48
                .ThunkPos = q48
            End If
        End If
        If ParseNOT() Then
            If ParseSuffix() Then
                pvPushAction ucsAct_3_Prefix
                ParsePrefix = True
                Exit Function
            Else
                .BufPos = p48
                .ThunkPos = q48
            End If
        End If
        If ParseSuffix() Then
            ParsePrefix = True
        End If
    End With
End Function

Public Function ParseAND() As Boolean
    With ctx
        If .BufData(.BufPos) = 38 Then              ' "&"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseAND = True
        End If
    End With
End Function

Public Function ParseAction() As Boolean
    Dim p253 As Long

    With ctx
        p253 = .BufPos
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            .CaptureBegin = .BufPos
            Do
                If .BufData(.BufPos) <> 125 And .BufPos < .BufSize Then ' "}"
                    .BufPos = .BufPos + 1
                Else
                    Exit Do
                End If
            Loop
            .CaptureEnd = .BufPos
            If .BufData(.BufPos) = 125 Then         ' "}"
                .BufPos = .BufPos + 1
                Call ParseSpacing
                ParseAction = True
                Exit Function
            Else
                .BufPos = p253
            End If
        End If
    End With
End Function

Public Function ParseSuffix() As Boolean
    If ParsePrimary() Then
        If ParseQUESTION() Then
            pvPushAction ucsAct_1_Suffix
            GoTo L1
        End If
        If ParseSTAR() Then
            pvPushAction ucsAct_2_Suffix
            GoTo L1
        End If
        If ParsePLUS() Then
            pvPushAction ucsAct_3_Suffix
            GoTo L1
        End If
L1:
        ParseSuffix = True
    End If
End Function

Public Function ParseNOT() As Boolean
    With ctx
        If .BufData(.BufPos) = 33 Then              ' "!"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseNOT = True
        End If
    End With
End Function

Public Function ParsePrimary() As Boolean
    Dim p86 As Long
    Dim q86 As Long

    With ctx
        p86 = .BufPos
        q86 = .ThunkPos
        If ParseIdentifier() Then
            pvPushAction ucsAct_1_Primary
            If ParseCOLON() Then
                If ParseIdentifier() Then
                    If ParseAssign() Then
                        .BufPos = p86
                        .ThunkPos = q86
                    Else
                        pvPushAction ucsAct_2_Primary
                        ParsePrimary = True
                        Exit Function
                    End If
                Else
                    .BufPos = p86
                    .ThunkPos = q86
                End If
            Else
                .BufPos = p86
                .ThunkPos = q86
            End If
        End If
        If ParseIdentifier() Then
            If ParseAssign() Then
                .BufPos = p86
                .ThunkPos = q86
            Else
                pvPushAction ucsAct_3_Primary
                ParsePrimary = True
                Exit Function
            End If
        End If
        If ParseOPEN() Then
            Call ParseExpression
            If ParseCLOSE() Then
                ParsePrimary = True
                Exit Function
            Else
                .BufPos = p86
                .ThunkPos = q86
            End If
        End If
        If ParseLiteral() Then
            pvPushAction ucsAct_4_Primary
            ParsePrimary = True
            Exit Function
        End If
        If ParseClass() Then
            pvPushAction ucsAct_5_Primary
            ParsePrimary = True
            Exit Function
        End If
        If ParseDOT() Then
            pvPushAction ucsAct_6_Primary
            ParsePrimary = True
            Exit Function
        End If
        If ParseAction() Then
            pvPushAction ucsAct_7_Primary
            ParsePrimary = True
            Exit Function
        End If
        If ParseBEGIN() Then
            pvPushAction ucsAct_8_Primary
            ParsePrimary = True
            Exit Function
        End If
        If ParseEND() Then
            pvPushAction ucsAct_9_Primary
            ParsePrimary = True
        End If
    End With
End Function

Public Function ParseQUESTION() As Boolean
    With ctx
        If .BufData(.BufPos) = 63 Then              ' "?"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseQUESTION = True
        End If
    End With
End Function

Public Function ParseSTAR() As Boolean
    With ctx
        If .BufData(.BufPos) = 42 Then              ' "*"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseSTAR = True
        End If
    End With
End Function

Public Function ParsePLUS() As Boolean
    With ctx
        If .BufData(.BufPos) = 43 Then              ' "+"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParsePLUS = True
        End If
    End With
End Function

Public Function ParseCOLON() As Boolean
    With ctx
        If .BufData(.BufPos) = 58 Then              ' ":"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseCOLON = True
        End If
    End With
End Function

Public Function ParseOPEN() As Boolean
    With ctx
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseOPEN = True
        End If
    End With
End Function

Public Function ParseCLOSE() As Boolean
    With ctx
        If .BufData(.BufPos) = 41 Then              ' ")"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseCLOSE = True
        End If
    End With
End Function

Public Function ParseLiteral() As Boolean
    Dim p152 As Long

    With ctx
        p152 = .BufPos
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            .CaptureBegin = .BufPos
            Do
                If .BufData(.BufPos) = 39 Then      ' "'"
                    Exit Do
                End If
                If Not ParseChar() Then
                    Exit Do
                End If
            Loop
            .CaptureEnd = .BufPos
            If .BufData(.BufPos) = 39 Then          ' "'"
                .BufPos = .BufPos + 1
                Call ParseSpacing
                ParseLiteral = True
                Exit Function
            Else
                .BufPos = p152
            End If
        End If
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            .CaptureBegin = .BufPos
            Do
                If .BufData(.BufPos) = 34 Then      ' """
                    Exit Do
                End If
                If Not ParseChar() Then
                    Exit Do
                End If
            Loop
            .CaptureEnd = .BufPos
            If .BufData(.BufPos) = 34 Then          ' """
                .BufPos = .BufPos + 1
                Call ParseSpacing
                ParseLiteral = True
                Exit Function
            Else
                .BufPos = p152
            End If
        End If
    End With
End Function

Public Function ParseClass() As Boolean
    Dim p155 As Long

    With ctx
        p155 = .BufPos
        If .BufData(.BufPos) = 91 Then              ' "["
            .BufPos = .BufPos + 1
            .CaptureBegin = .BufPos
            Do
                If .BufData(.BufPos) = 93 Then      ' "]"
                    Exit Do
                End If
                If Not ParseRange() Then
                    Exit Do
                End If
            Loop
            .CaptureEnd = .BufPos
            If .BufData(.BufPos) = 93 Then          ' "]"
                .BufPos = .BufPos + 1
                Call ParseSpacing
                ParseClass = True
                Exit Function
            Else
                .BufPos = p155
            End If
        End If
    End With
End Function

Public Function ParseDOT() As Boolean
    With ctx
        If .BufData(.BufPos) = 46 Then              ' "."
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseDOT = True
        End If
    End With
End Function

Public Function ParseBEGIN() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseBEGIN = True
        End If
    End With
End Function

Public Function ParseEND() As Boolean
    With ctx
        If .BufData(.BufPos) = 62 Then              ' ">"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseEND = True
        End If
    End With
End Function

Public Function ParseChar() As Boolean
    Dim p181 As Long

    With ctx
        p181 = .BufPos
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 97, 98, 101, 102, 110, 114, 116, 118, 39, 34, 91, 93, 92 ' [abefnrtv'"\[\]\\]
                .BufPos = .BufPos + 1
                ParseChar = True
                Exit Function
            Case Else
                .BufPos = p181
            End Select
        End If
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 48 To 51                           ' [0-3]
                .BufPos = .BufPos + 1
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                    Select Case .BufData(.BufPos)
                    Case 48 To 55                   ' [0-7]
                        .BufPos = .BufPos + 1
                        ParseChar = True
                        Exit Function
                    Case Else
                        .BufPos = p181
                    End Select
                Case Else
                    .BufPos = p181
                End Select
            Case Else
                .BufPos = p181
            End Select
        End If
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Select Case .BufData(.BufPos)
            Case 48 To 55                           ' [0-7]
                .BufPos = .BufPos + 1
                Select Case .BufData(.BufPos)
                Case 48 To 55                       ' [0-7]
                    .BufPos = .BufPos + 1
                    GoTo L2
                End Select
L2:
                ParseChar = True
                Exit Function
            Case Else
                .BufPos = p181
            End Select
        End If
        If .BufData(.BufPos) = 92 And .BufData(.BufPos + 1) = 45 Then ' "\\-"
            .BufPos = .BufPos + 2
            ParseChar = True
            Exit Function
        End If
        If Not .BufData(.BufPos) = 92 Then          ' "\\"
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
                ParseChar = True
            End If
        End If
    End With
End Function

Public Function ParseRange() As Boolean
    Dim p172 As Long

    With ctx
        p172 = .BufPos
        If ParseChar() Then
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = p172
                Else
                    If ParseChar() Then
                        ParseRange = True
                        Exit Function
                    Else
                        .BufPos = p172
                    End If
                End If
            Else
                .BufPos = p172
            End If
        End If
        If ParseChar() Then
            ParseRange = True
        End If
    End With
End Function

Public Function ParseLEFTARROW() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 And .BufData(.BufPos + 1) = 45 Then ' "<-"
            .BufPos = .BufPos + 2
            Call ParseSpacing
            ParseLEFTARROW = True
        End If
    End With
End Function

Public Function ParseSLASH() As Boolean
    With ctx
        If .BufData(.BufPos) = 47 Then              ' "/"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseSLASH = True
        End If
    End With
End Function

Public Function ParseSpace() As Boolean
    With ctx
        If .BufData(.BufPos) = 32 Then              ' " "
            .BufPos = .BufPos + 1
            ParseSpace = True
            Exit Function
        End If
        If .BufData(.BufPos) = 9 Then               ' "\t"
            .BufPos = .BufPos + 1
            ParseSpace = True
            Exit Function
        End If
        If ParseEndOfLine() Then
            ParseSpace = True
        End If
    End With
End Function

Public Function ParseComment() As Boolean
    Dim p239 As Long
    Dim p237 As Long

    With ctx
        p239 = .BufPos
        If .BufData(.BufPos) = 35 Then              ' "#"
            .BufPos = .BufPos + 1
            Do
                p237 = .BufPos
                If ParseEndOfLine() Then
                    .BufPos = p237
                    Exit Do
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p237
                    Exit Do
                End If
            Loop
            If ParseEndOfLine() Then
                ParseComment = True
                Exit Function
            Else
                .BufPos = p239
            End If
        End If
    End With
End Function

Public Function ParseEndOfLine() As Boolean
    With ctx
        If .BufData(.BufPos) = 13 And .BufData(.BufPos + 1) = 10 Then ' "\r\n"
            .BufPos = .BufPos + 2
            ParseEndOfLine = True
            Exit Function
        End If
        If .BufData(.BufPos) = 10 Then              ' "\n"
            .BufPos = .BufPos + 1
            ParseEndOfLine = True
            Exit Function
        End If
        If .BufData(.BufPos) = 13 Then              ' "\r"
            .BufPos = .BufPos + 1
            ParseEndOfLine = True
        End If
    End With
End Function

Public Function ParseEQUAL() As Boolean
    With ctx
        If .BufData(.BufPos) = 61 Then              ' "="
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseEQUAL = True
        End If
    End With
End Function

Public Function ParseBAR() As Boolean
    With ctx
        If .BufData(.BufPos) = 124 Then             ' "|"
            .BufPos = .BufPos + 1
            Call ParseSpacing
            ParseBAR = True
        End If
    End With
End Function

Private Sub pvImplAction(ByVal eAction As UcsParserActionsEnum, ByVal lOffset As Long, ByVal lSize As Long)
    With ctx.UserData
        Select Case eAction
        Case ucsAct_2_Definition
             Dim lNode As Long: lNode = .Pop(): .RuleExpression(.Pop()) = lNode
        Case ucsAct_1_Definition
            
            lNode = .BeginRule(.FindRule(Mid$(ctx.Contents, lOffset, lSize)))
            If .RuleExpression(lNode) <> 0 Then
                ConsoleError "Rule '%1' redefined" & vbCrLf, Mid$(ctx.Contents, lOffset, lSize)
            End If
            .Push lNode

        Case ucsAct_1_Expression
             lNode = .Pop(): .Push .AppendChild(.MakeAlternate(.Pop), lNode)
        Case ucsAct_2_Sequence
             .Push .MakePredicate("1")
        Case ucsAct_1_Sequence
             lNode = .Pop(): .Push .AppendSequence(.MakeSequence(.Pop), lNode)
        Case ucsAct_1_Prefix
             .Push .MakePredicate(Mid$(ctx.Contents, lOffset, lSize))
        Case ucsAct_2_Prefix
             .Push .MakePeekFor(.Pop)
        Case ucsAct_3_Prefix
             .Push .MakePeekNot(.Pop)
        Case ucsAct_1_Suffix
             .Push .MakeQuery(.Pop)
        Case ucsAct_2_Suffix
             .Push .MakeStar(.Pop)
        Case ucsAct_3_Suffix
             .Push .MakePlus(.Pop)
        Case ucsAct_2_Primary
             .Push .MakeReference(.FindRule(Mid$(ctx.Contents, lOffset, lSize)), .Pop)
        Case ucsAct_1_Primary
             .Push .MakeVariable(Mid$(ctx.Contents, lOffset, lSize))
        Case ucsAct_3_Primary
             .Push .MakeReference(.FindRule(Mid$(ctx.Contents, lOffset, lSize)))
        Case ucsAct_4_Primary
             .Push .MakeString(Mid$(ctx.Contents, lOffset, lSize))
        Case ucsAct_5_Primary
             .Push .MakeClass(Mid$(ctx.Contents, lOffset, lSize))
        Case ucsAct_6_Primary
             .Push .MakeDot()
        Case ucsAct_7_Primary
             .Push .MakeAction(Mid$(ctx.Contents, lOffset, lSize))
        Case ucsAct_8_Primary
             .Push .MakePredicate("YY_BEGIN")
        Case ucsAct_9_Primary
             .Push .MakePredicate("YY_END")
        End Select
    End With
End Sub
