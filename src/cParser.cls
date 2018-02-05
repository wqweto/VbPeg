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
'=========================================================================
'
' VbPeg (c) 2018 by wqweto@gmail.com
'
' PEG parser generator for VB6
'
' cParser.cls - PEG grammar manual parser
'
'=========================================================================
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

Private Enum UcsParserActionsEnum
    ucsAct_1_Definition = 1
    ucsAct_2_Definition
    ucsAct_1_Expression
    ucsAct_1_Sequence
    ucsAct_2_Sequence
    ucsAct_1_Prefix
    ucsAct_2_Prefix
    ucsAct_3_Prefix
    ucsAct_1_Suffix
    ucsAct_2_Suffix
    ucsAct_3_Suffix
    ucsAct_1_Primary
    ucsAct_2_Primary
    ucsAct_3_Primary
    ucsAct_4_Primary
    ucsAct_5_Primary
    ucsAct_6_Primary
    ucsAct_7_Primary
End Enum

Private Type UcsParserThunkType
    Action              As UcsParserActionsEnum
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

Private Type UcsParserPosType
    BufPos              As Long
    ThunkPos            As Long
End Type

Private ctx                     As UcsParserType

'=========================================================================
' Properties
'=========================================================================

Property Get LastError() As String
    LastError = ctx.LastError
End Property

Property Get ParserVersion() As String
    ParserVersion = "4.2.2018"
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(sContents As String, oUserData As cTree) As Boolean
    Dim uEmpty          As UcsParserType
    
    ctx = uEmpty
    With ctx
        If LenB(sContents) = 0 Then
            .LastError = "Input has no contents"
            Exit Function
        End If
        .Contents = sContents
        ReDim .BufData(0 To Len(sContents) + 3) As Integer
        Call CopyMemory(.BufData(0), ByVal StrPtr(sContents), LenB(sContents))
        .BufSize = Len(sContents)
        .BufData(.BufSize) = -1
        ReDim .ThunkData(0 To 4) As UcsParserThunkType
        Set .UserData = oUserData
    End With
    '--- success
    Init = True
End Function

Public Function Flush() As Boolean
    Dim lIdx            As Long
    
    With ctx
        For lIdx = 0 To .ThunkPos - 1
            With .ThunkData(lIdx)
                pvImplAction .Action, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin
            End With
        Next
        .BufPos = 0
        ReDim .ThunkData(0 To 4) As UcsParserThunkType
        .ThunkPos = 0
        .CaptureBegin = 0
        .CaptureEnd = 0
        .LastError = vbNullString
    End With
    '--- success
    Flush = True
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

Private Sub pvSavePos(uPos As UcsParserPosType)
    uPos.BufPos = ctx.BufPos
    uPos.ThunkPos = ctx.ThunkPos
End Sub

Private Sub pvRestorePos(uPos As UcsParserPosType)
    ctx.BufPos = uPos.BufPos
    ctx.ThunkPos = uPos.ThunkPos
End Sub

Private Function pvMatchDot() As Boolean
    With ctx
        If .BufPos < .BufSize Then
            .BufPos = .BufPos + 1
            pvMatchDot = True
        End If
    End With
End Function

Private Function pvMatchChar(ByVal nChar As Integer) As Boolean
    With ctx
        If .BufData(.BufPos) = nChar Then
            .BufPos = .BufPos + 1
            pvMatchChar = True
        End If
    End With
End Function

Private Function pvMatchString(sText As String) As Boolean
    With ctx
        If .BufPos + Len(sText) <= .BufSize Then
            pvMatchString = RtlCompareMemory(.BufData(.BufPos), ByVal StrPtr(sText), LenB(sText)) = LenB(sText)
        End If
    End With
End Function

Private Function pvMatchClass1() As Boolean
    '--- [^}]
    With ctx
        Select Case .BufData(.BufPos)
        Case 125 '--- Asc("}")
        Case Else
            .BufPos = .BufPos + 1
            pvMatchClass1 = True
        End Select
    End With
End Function

Private Function pvMatchClass2() As Boolean
    '--- [abefnrtv'"\[\]\\]
    With ctx
        Select Case .BufData(.BufPos)
        '--- Asc("a"), Asc("b"), Asc("e"), Asc("f"), Asc("n"), Asc("r"), Asc("t"), Asc("v"),
        '---   Asc("'"), Asc(""""), Asc("["), Asc("]"), Asc("\")
        Case 97, 98, 101, 102, 110, 114, 116, 118, _
                39, 34, 91, 93, 92
            .BufPos = .BufPos + 1
            pvMatchClass2 = True
        End Select
    End With
End Function

Private Function pvMatchClass3() As Boolean
    '--- [0-3]
    With ctx
        Select Case .BufData(.BufPos)
        Case 48 To 51
            .BufPos = .BufPos + 1
            pvMatchClass3 = True
        End Select
    End With
End Function

Private Function pvMatchClass4() As Boolean
    '--- [0-7]
    With ctx
        Select Case .BufData(.BufPos)
        Case 48 To 55
            .BufPos = .BufPos + 1
            pvMatchClass4 = True
        End Select
    End With
End Function

Public Function ParseGrammar() As Boolean
    Dim p1 As UcsParserPosType: pvSavePos p1
    '--- Spacing Definition+ EndOfFile
    If ParseSpacing() Then
        If ParseDefinition() Then
            Do
                If Not ParseDefinition() Then
                    Exit Do
                End If
            Loop
            If ParseEndOfFile() Then
                GoTo L0
            End If
        End If
        pvRestorePos p1
    End If
    Exit Function
L0:
    '--- success
    ParseGrammar = True
End Function

Public Function ParseDefinition() As Boolean
    Dim p1 As UcsParserPosType: pvSavePos p1
    '--- Identifier LEFTARROW Expression
    If ParseIdentifier() Then
        pvPushAction ucsAct_1_Definition
        If ParseLEFTARROW() Then
            If ParseExpression() Then
                pvPushAction ucsAct_2_Definition
                GoTo L0
            End If
        End If
        pvRestorePos p1
    End If
    Exit Function
L0:
    '--- success
    ParseDefinition = True
End Function

Public Function ParseExpression() As Boolean
    '--- Sequence (SLASH Sequence)*
    If ParseSequence() Then
        Do
            Dim p1 As UcsParserPosType: pvSavePos p1
            If Not ParseSLASH() Then
                Exit Do
            End If
            If Not ParseSequence() Then
                pvRestorePos p1
                Exit Do
            End If
            pvPushAction ucsAct_1_Expression
        Loop
        GoTo L0
    End If
    Exit Function
L0:
    '--- success
    ParseExpression = True
End Function

Public Function ParseSequence() As Boolean
    '--- Prefix (Prefix)*
    If ParsePrefix() Then
        Do
            If Not ParsePrefix() Then
                Exit Do
            End If
            pvPushAction ucsAct_1_Sequence
        Loop
        GoTo L0
    Else
        pvPushAction ucsAct_2_Sequence
        GoTo L0
    End If
    Exit Function
L0:
    '--- success
    ParseSequence = True
End Function

Public Function ParsePrefix() As Boolean
    Dim p1 As UcsParserPosType: pvSavePos p1
    '--- AND Action / AND Suffix / NOT Suffix / Suffix
    If ParseAND() Then
        If ParseAction() Then
            pvPushAction ucsAct_1_Prefix
            GoTo L0
        End If
        pvRestorePos p1
    End If
    If ParseAND() Then
        If ParseSuffix() Then
            pvPushAction ucsAct_2_Prefix
            GoTo L0
        End If
        pvRestorePos p1
    End If
    If ParseNOT() Then
        If ParseSuffix() Then
            pvPushAction ucsAct_3_Prefix
            GoTo L0
        End If
        pvRestorePos p1
    End If
    If ParseSuffix() Then
        GoTo L0
    End If
    Exit Function
L0:
    '--- success
    ParsePrefix = True
End Function

Public Function ParseSuffix() As Boolean
    '--- Primary (QUESTION / STAR / PLUS)?
    If ParsePrimary() Then
        If ParseQUESTION() Then
            pvPushAction ucsAct_1_Suffix
            GoTo L0
        End If
        If ParseSTAR() Then
            pvPushAction ucsAct_2_Suffix
            GoTo L0
        End If
        If ParsePLUS() Then
            pvPushAction ucsAct_3_Suffix
            GoTo L0
        End If
        GoTo L0
    End If
    Exit Function
L0:
    '--- success
    ParseSuffix = True
End Function

Public Function ParsePrimary() As Boolean
    Dim p1 As UcsParserPosType: pvSavePos p1
    '--- Identifier !LEFTARROW
    If ParseIdentifier() Then
        If Not ParseLEFTARROW() Then
            pvPushAction ucsAct_1_Primary
            GoTo L0
        End If
        pvRestorePos p1
    End If
    '--- OPEN Expression CLOSE
    If ParseOPEN() Then
        If ParseExpression() Then
            If ParseCLOSE() Then
                GoTo L0
            End If
        End If
        pvRestorePos p1
    End If
    '--- Literal
    If ParseLiteral() Then
        pvPushAction ucsAct_2_Primary
        GoTo L0
    End If
    '--- Class
    If ParseClass() Then
        pvPushAction ucsAct_3_Primary
        GoTo L0
    End If
    '--- DOT
    If ParseDOT() Then
        pvPushAction ucsAct_4_Primary
        GoTo L0
    End If
    '--- Action
    If ParseAction() Then
        pvPushAction ucsAct_5_Primary
        GoTo L0
    End If
    '--- BEGIN
    If ParseBEGIN() Then
        pvPushAction ucsAct_6_Primary
        GoTo L0
    End If
    '--- END
    If ParseEND() Then
        pvPushAction ucsAct_7_Primary
        GoTo L0
    End If
    Exit Function
L0:
    '--- success
    ParsePrimary = True
End Function

Public Function ParseIdentifier() As Boolean
    '--- < IdentStart IdentCont* > Spacing
    ctx.CaptureBegin = ctx.BufPos
    If ParseIdentStart() Then
        Do
            If Not ParseIdentCont() Then
                Exit Do
            End If
        Loop
        ctx.CaptureEnd = ctx.BufPos
        If ParseSpacing() Then
            GoTo L0
        End If
    End If
    Exit Function
L0:
    '--- success
    ParseIdentifier = True
End Function

Public Function ParseIdentStart() As Boolean
    '--- [a-zA-Z_]
    With ctx
        Select Case .BufData(.BufPos)
        Case 97 To 122, 65 To 90, 95
            .BufPos = .BufPos + 1
            '--- success
            ParseIdentStart = True
        End Select
    End With
End Function

Public Function ParseIdentCont() As Boolean
    '--- [a-zA-Z_0-9]
    With ctx
        Select Case .BufData(.BufPos)
        Case 97 To 122, 65 To 90, 95, 48 To 57
            .BufPos = .BufPos + 1
            '--- success
            ParseIdentCont = True
        End Select
    End With
End Function

Public Function ParseLiteral() As Boolean
    Dim p1 As UcsParserPosType: pvSavePos p1
    '--- ['] < (!['] Char )* > ['] Spacing
    If pvMatchChar(39) Then             '--- Asc("'")
        ctx.CaptureBegin = ctx.BufPos
        Do
            Dim p2 As UcsParserPosType: pvSavePos p2
            If pvMatchChar(39) Then     '--- Asc("'")
                pvRestorePos p2
                Exit Do
            End If
            If Not ParseChar() Then
                Exit Do
            End If
        Loop
        ctx.CaptureEnd = ctx.BufPos
        If pvMatchChar(39) Then
            If ParseSpacing() Then
                GoTo L0
            End If
        End If
        pvRestorePos p1
    End If
    '--- ["] < (!["] Char)* > ["] Spacing
    If pvMatchChar(34) Then             '--- Asc("""")
        ctx.CaptureBegin = ctx.BufPos
        Do
            pvSavePos p2
            If pvMatchChar(34) Then     '--- Asc("""")
                pvRestorePos p2
                Exit Do
            End If
            If Not ParseChar() Then
                Exit Do
            End If
        Loop
        ctx.CaptureEnd = ctx.BufPos
        If pvMatchChar(34) Then
            If ParseSpacing() Then
                GoTo L0
            End If
        End If
        pvRestorePos p1
    End If
    Exit Function
L0:
    '--- success
    ParseLiteral = True
End Function

Public Function ParseClass() As Boolean
    Dim p1 As UcsParserPosType: pvSavePos p1
    '--- '[' < (!']' Range)* > ']' Spacing
    If pvMatchChar(91) Then             '--- Asc("[")
        ctx.CaptureBegin = ctx.BufPos
        Do
            Dim p2 As UcsParserPosType: pvSavePos p2
            If pvMatchChar(93) Then     '--- Asc("]")
                pvRestorePos p2
                Exit Do
            End If
            If Not ParseRange() Then
                Exit Do
            End If
        Loop
        ctx.CaptureEnd = ctx.BufPos
        If pvMatchChar(93) Then
            If ParseSpacing() Then
                GoTo L0
            End If
        End If
        pvRestorePos p1
    End If
    Exit Function
L0:
    '--- success
    ParseClass = True
End Function

Public Function ParseRange() As Boolean
    Dim p1 As UcsParserPosType: pvSavePos p1
    '--- Char '-' Char / Char
    If ParseChar() Then
        If pvMatchChar(45) Then         '--- Asc("-")
            If ParseChar() Then
                GoTo L0
            End If
        End If
        pvRestorePos p1
    End If
    If ParseChar() Then
        GoTo L0
    End If
    Exit Function
L0:
    '--- success
    ParseRange = True
End Function

Public Function ParseChar() As Boolean
    Dim p1 As UcsParserPosType: pvSavePos p1
    '--- '\\' [abefnrtv'"\[\]\\]
    If pvMatchChar(92) Then             '--- Asc("\")
        If pvMatchClass2() Then
            GoTo L0
        End If
        pvRestorePos p1
    End If
    '--- '\\' [0-3][0-7][0-7]
    If pvMatchChar(92) Then             '--- Asc("\")
        If pvMatchClass3() Then         '--- [0-3]
            If pvMatchClass4() Then     '--- [0-7]
                If pvMatchClass4() Then '--- [0-7]
                    GoTo L0
                End If
            End If
        End If
        pvRestorePos p1
    End If
    '--- '\\' [0-7][0-7]?
    If pvMatchChar(92) Then         '--- Asc("\")
        If pvMatchClass4() Then     '--- [0-7]
            Call pvMatchClass4      '--- [0-7]
            GoTo L0
        End If
        pvRestorePos p1
    End If
    '--- '\\' '-'
    If pvMatchChar(92) Then         '--- Asc("\")
        If pvMatchChar(45) Then     '--- Asc("-")
            GoTo L0
        End If
        pvRestorePos p1
    End If
    '--- !'\\' .
    If pvMatchChar(92) Then         '--- Asc("\")
        pvRestorePos p1
    Else
        If pvMatchDot() Then
            GoTo L0
        End If
    End If
    Exit Function
L0:
    '--- success
    ParseChar = True
End Function

Public Function ParseLEFTARROW() As Boolean
    If pvMatchString("<-") Then
        ctx.BufPos = ctx.BufPos + 2
        If ParseSpacing() Then
            '--- success
            ParseLEFTARROW = True
        End If
    End If
End Function

Public Function ParseSLASH() As Boolean
    If pvMatchChar(47) Then         '--- Asc("/")
        If ParseSpacing() Then
            '--- success
            ParseSLASH = True
        End If
    End If
End Function

Public Function ParseAND() As Boolean
    If pvMatchChar(38) Then         '--- Asc("&")
        If ParseSpacing() Then
            '--- success
            ParseAND = True
        End If
    End If
End Function

Public Function ParseNOT() As Boolean
    If pvMatchChar(33) Then         '--- Asc("!")
        If ParseSpacing() Then
            '--- success
            ParseNOT = True
        End If
    End If
End Function

Public Function ParseQUESTION() As Boolean
    If pvMatchChar(63) Then         '--- Asc("?")
        If ParseSpacing() Then
            '--- success
            ParseQUESTION = True
        End If
    End If
End Function

Public Function ParseSTAR() As Boolean
    If pvMatchChar(42) Then         '--- Asc("*")
        If ParseSpacing() Then
            '--- success
            ParseSTAR = True
        End If
    End If
End Function

Public Function ParsePLUS() As Boolean
    If pvMatchChar(43) Then         '--- Asc("+")
        If ParseSpacing() Then
            '--- success
            ParsePLUS = True
        End If
    End If
End Function

Public Function ParseOPEN() As Boolean
    If pvMatchChar(40) Then         '--- Asc("(")
        If ParseSpacing() Then
            '--- success
            ParseOPEN = True
        End If
    End If
End Function

Public Function ParseCLOSE() As Boolean
    If pvMatchChar(41) Then         '--- Asc(")")
        If ParseSpacing() Then
            '--- success
            ParseCLOSE = True
        End If
    End If
End Function

Public Function ParseDOT() As Boolean
    If pvMatchChar(46) Then         '--- Asc(".")
        If ParseSpacing() Then
            '--- success
            ParseDOT = True
        End If
    End If
End Function

Public Function ParseSpacing() As Boolean
    Do
        If Not ParseSpace() Then
            If Not ParseComment() Then
                Exit Do
            End If
        End If
    Loop
    '--- success
    ParseSpacing = True
End Function

Public Function ParseComment() As Boolean
    Dim p1 As UcsParserPosType: pvSavePos p1
    '--- '#' (!EndOfLine .)* EndOfLine
    If pvMatchChar(35) Then         '--- Asc("#")
        Do
            Dim p2 As UcsParserPosType: pvSavePos p2
            If ParseEndOfLine() Then
                pvRestorePos p2
                Exit Do
            End If
            If Not pvMatchDot() Then
                Exit Do
            End If
        Loop
        If ParseEndOfLine() Then
            GoTo L0
        End If
        pvRestorePos p1
    End If
    Exit Function
L0:
    '--- success
    ParseComment = True
End Function

Public Function ParseSpace() As Boolean
    '--- ' ' / '\t' / EndOfLine
    If pvMatchChar(32) Then         '--- Asc(" ")
        GoTo L0
    End If
    If pvMatchChar(9) Then          '--- Asc(vbTab)
        GoTo L0
    End If
    If ParseEndOfLine() Then
        GoTo L0
    End If
    Exit Function
L0:
    '--- success
    ParseSpace = True
End Function

Public Function ParseEndOfLine() As Boolean
    Dim p1 As UcsParserPosType: pvSavePos p1
    '--- '\r\n' / '\n' / '\r'
    If pvMatchChar(13) Then         '--- Asc(vbCr)
        If pvMatchChar(10) Then     '--- Asc(vbLf)
            GoTo L0
        End If
        pvRestorePos p1
    End If
    If pvMatchChar(10) Then         '--- Asc(vbLf)
        GoTo L0
    End If
    If pvMatchChar(13) Then         '--- Asc(vbCr)
        GoTo L0
    End If
    Exit Function
L0:
    '--- success
    ParseEndOfLine = True
End Function

Public Function ParseEndOfFile() As Boolean
    Dim p1 As UcsParserPosType: pvSavePos p1
    If pvMatchDot() Then
        pvRestorePos p1
        Exit Function
    End If
    ParseEndOfFile = True
End Function

Public Function ParseAction() As Boolean
    Dim p1 As UcsParserPosType: pvSavePos p1
    '--- '{' < [^}]* > '}' Spacing
    If pvMatchChar(123) Then        '--- Asc("{")
        ctx.CaptureBegin = ctx.BufPos
        Do
            If Not pvMatchClass1() Then
                Exit Do
            End If
        Loop
        ctx.CaptureEnd = ctx.BufPos
        If pvMatchChar(125) Then    '--- Asc("}")
            If ParseSpacing() Then
                GoTo L0
            End If
        End If
        pvRestorePos p1
    End If
    Exit Function
L0:
    '--- success
    ParseAction = True
End Function

Public Function ParseBEGIN() As Boolean
    If pvMatchChar(60) Then         '--- Asc("<")
        If ParseSpacing() Then
            '--- success
            ParseBEGIN = True
        End If
    End If
End Function

Public Function ParseEND() As Boolean
    If pvMatchChar(62) Then         '--- Asc(">")
        If ParseSpacing() Then
            '--- success
            ParseEND = True
        End If
    End If
End Function

Private Sub pvImplAction(eAction As UcsParserActionsEnum, ByVal lOffset As Long, ByVal lSize As Long)
    Dim lNode           As Long
    
    With ctx.UserData
        Select Case eAction
        Case ucsAct_1_Definition
            lNode = .BeginRule(.FindRule(Mid$(ctx.Contents, lOffset, lSize)))
            If .RuleExpression(lNode) <> 0 Then
                ConsoleError "Rule '%1' redefined" & vbCrLf, Mid$(ctx.Contents, lOffset, lSize)
            End If
            .Push lNode
        Case ucsAct_2_Definition
            lNode = .Pop()
            .RuleExpression(.Pop()) = lNode
        Case ucsAct_1_Expression
            lNode = .Pop()
            .Push .AppendChild(.MakeAlternate(.Pop), lNode)
        Case ucsAct_1_Sequence
            lNode = .Pop()
            .Push .AppendSequence(.MakeSequence(.Pop), lNode)
        Case ucsAct_2_Sequence
            .Push .MakePredicate("1")
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
        Case ucsAct_1_Primary
            .Push .MakeReference(.FindRule(Mid$(ctx.Contents, lOffset, lSize)))
        Case ucsAct_2_Primary
            .Push .MakeString(Mid$(ctx.Contents, lOffset, lSize))
        Case ucsAct_3_Primary
            .Push .MakeClass(Mid$(ctx.Contents, lOffset, lSize))
        Case ucsAct_4_Primary
            .Push .MakeDot()
        Case ucsAct_5_Primary
            .Push .MakeAction(Mid$(ctx.Contents, lOffset, lSize))
        Case ucsAct_6_Primary
            .Push .MakePredicate("YY_BEGIN")
        Case ucsAct_7_Primary
            .Push .MakePredicate("YY_END")
        Case Else
            ConsolePrint "Unknown action: " & eAction & vbCrLf
        End Select
    End With
End Sub
