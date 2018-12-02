VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMarkdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Auto-generated on 3.12.2018 1:29:29
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
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

Private Type SAFEARRAY1D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As Long
    cElements           As Long
    lLbound             As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const LNG_MAXINT            As Long = 2 ^ 31 - 1

'= generated enum ========================================================

Private Enum UcsParserActionsEnum
    ucsAct_2_Doc
    ucsAct_1_Doc
    ucsAct_1_StartList
    ucsAct_1_BlockQuote
    ucsAct_2_Verbatim
    ucsAct_1_Verbatim
    ucsAct_3_Note
    ucsAct_2_Note
    ucsAct_1_Note
    ucsAct_1_Reference
    ucsAct_1_HorizontalRule
    ucsAct_1_OrderedList
    ucsAct_1_BulletList
    ucsAct_1_HtmlBlock
    ucsAct_1_StyleBlock
    ucsAct_1_Para
    ucsAct_1_Plain
    ucsAct_3_Inlines
    ucsAct_2_Inlines
    ucsAct_1_Inlines
    ucsAct_1_AtxStart
    ucsAct_2_AtxHeading
    ucsAct_1_AtxHeading
    ucsAct_2_SetextHeading1
    ucsAct_1_SetextHeading1
    ucsAct_2_SetextHeading2
    ucsAct_1_SetextHeading2
    ucsAct_4_BlockQuoteRaw
    ucsAct_1_BlockQuoteRaw
    ucsAct_2_BlockQuoteRaw
    ucsAct_3_BlockQuoteRaw
    ucsAct_1_Line
    ucsAct_3_VerbatimChunk
    ucsAct_2_VerbatimChunk
    ucsAct_1_VerbatimChunk
    ucsAct_2_ListTight
    ucsAct_1_ListTight
    ucsAct_2_ListLoose
    ucsAct_1_ListLoose
    ucsAct_3_ListItemTight
    ucsAct_2_ListItemTight
    ucsAct_1_ListItemTight
    ucsAct_3_ListItem
    ucsAct_2_ListItem
    ucsAct_1_ListItem
    ucsAct_3_ListBlock
    ucsAct_2_ListBlock
    ucsAct_1_ListBlock
    ucsAct_3_ListContinuationBlock
    ucsAct_2_ListContinuationBlock
    ucsAct_1_ListContinuationBlock
    ucsAct_3_Str
    ucsAct_2_Str
    ucsAct_1_Str
    ucsAct_1_UlOrStarLine
    ucsAct_1_Space
    ucsAct_2_Strike
    ucsAct_1_Strike
    ucsAct_1_Image
    ucsAct_1_NoteReference
    ucsAct_2_InlineNote
    ucsAct_1_InlineNote
    ucsAct_1_Code
    ucsAct_1_RawHtml
    ucsAct_1_Entity
    ucsAct_1_EscapedChar
    ucsAct_1_Symbol
    ucsAct_1_StrChunk
    ucsAct_1_AposChunk
    ucsAct_1_LineBreak
    ucsAct_1_TerminalEndline
    ucsAct_1_NormalEndline
    ucsAct_3_EmphStar
    ucsAct_2_EmphStar
    ucsAct_1_EmphStar
    ucsAct_3_EmphUl
    ucsAct_2_EmphUl
    ucsAct_1_EmphUl
    ucsAct_2_StrongStar
    ucsAct_1_StrongStar
    ucsAct_2_StrongUl
    ucsAct_1_StrongUl
    ucsAct_1_ExplicitLink
    ucsAct_1_ReferenceLinkDouble
    ucsAct_1_ReferenceLinkSingle
    ucsAct_2_Label
    ucsAct_1_Label
    ucsAct_1_Source
    ucsAct_1_Title
    ucsAct_1_AutoLinkUrl
    ucsAct_1_AutoLinkEmail
    ucsAct_1_RefSrc
    ucsAct_1_RefTitle
    ucsAct_1_Ellipsis
    ucsAct_2_SingleQuoted
    ucsAct_1_SingleQuoted
    ucsAct_2_DoubleQuoted
    ucsAct_1_DoubleQuoted
    ucsAct_1_Apostrophe
    ucsAct_1_EmDash
    ucsAct_1_EnDash
    ucsAct_1_RawNoteReference
    ucsAct_3_RawNoteBlock
    ucsAct_2_RawNoteBlock
    ucsAct_1_RawNoteBlock
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

Private Enum UcsStateEnum
    EXT_NOTES
    EXT_STRIKE
    EXT_SMART
End Enum

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
    ParserVersion = "3.12.2018 1:29:29"
End Property

Property Get Contents(Optional ByVal lOffset As Long = 1, Optional ByVal lSize As Long = LNG_MAXINT) As String
    Contents = Mid$(ctx.Contents, lOffset, lSize)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Match(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant, Optional Result As Variant) As Long
    If BeginMatch(sSubject, StartPos, UserData) Then
        If ParseDoc() Then
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
        Call CopyMemory(ByVal VarPtr(.Contents), ByVal VarPtr(sSubject), 4)
        With .BufArray
            .cDims = 1
            .cbElements = 2
            .fFeatures = 1 ' FADF_AUTO
            .pvData = StrPtr(sSubject)
            .cElements = Len(sSubject) + 2 '-- look-ahead chars
        End With
        Call CopyMemory(ByVal ArrPtr(.BufData), VarPtr(.BufArray), 4)
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
        Call CopyMemory(ByVal VarPtr(.Contents), 0&, 4)
        Call CopyMemory(ByVal ArrPtr(.BufData), 0&, 4)
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

Public Function ParseDoc() As Boolean
    Dim p14 As Long
    Dim q14 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        If pvMatchString("﻿") Then                ' "\357\273\277"
            .BufPos = .BufPos + 3
        End If
        pvPushThunk ucsActResultClear
        If ParseStartList() Then
            pvPushThunk ucsActVarSet, 1
            Do
                p14 = .BufPos
                q14 = .ThunkPos
                If Not (ParseBlock()) Then
                    .BufPos = p14
                    .ThunkPos = q14
                    Exit Do
                End If
                pvPushThunk ucsAct_1_Doc, .CaptureBegin, .CaptureEnd
            Loop
            pvPushThunk ucsAct_2_Doc, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -1
            Call pvSetAdvance
            ParseDoc = True
        End If
    End With
End Function

Private Function ParseStartList() As Boolean
    Dim p2759 As Long

    With ctx
        p2759 = .BufPos
        If .BufPos < .BufSize Then
            .BufPos = p2759
            pvPushThunk ucsAct_1_StartList, .CaptureBegin, .CaptureEnd
            ParseStartList = True
        End If
    End With
End Function

Private Function ParseBlock() As Boolean
    Dim p18 As Long
    Dim p24 As Long
    Dim q24 As Long

    With ctx
        Do
            p18 = .BufPos
            If Not (ParseBlankLine()) Then
                .BufPos = p18
                Exit Do
            End If
        Loop
        p24 = .BufPos
        q24 = .ThunkPos
        If ParseBlockQuote() Then
            ParseBlock = True
            Exit Function
        Else
            .BufPos = p24
            .ThunkPos = q24
        End If
        If ParseVerbatim() Then
            ParseBlock = True
            Exit Function
        Else
            .BufPos = p24
            .ThunkPos = q24
        End If
        If ParseNote() Then
            ParseBlock = True
            Exit Function
        Else
            .BufPos = p24
            .ThunkPos = q24
        End If
        If ParseReference() Then
            ParseBlock = True
            Exit Function
        Else
            .BufPos = p24
            .ThunkPos = q24
        End If
        If ParseHorizontalRule() Then
            ParseBlock = True
            Exit Function
        Else
            .BufPos = p24
            .ThunkPos = q24
        End If
        If ParseHeading() Then
            ParseBlock = True
            Exit Function
        Else
            .BufPos = p24
            .ThunkPos = q24
        End If
        If ParseOrderedList() Then
            ParseBlock = True
            Exit Function
        Else
            .BufPos = p24
            .ThunkPos = q24
        End If
        If ParseBulletList() Then
            ParseBlock = True
            Exit Function
        Else
            .BufPos = p24
            .ThunkPos = q24
        End If
        If ParseHtmlBlock() Then
            ParseBlock = True
            Exit Function
        Else
            .BufPos = p24
            .ThunkPos = q24
        End If
        If ParseStyleBlock() Then
            ParseBlock = True
            Exit Function
        Else
            .BufPos = p24
            .ThunkPos = q24
        End If
        If ParsePara() Then
            ParseBlock = True
            Exit Function
        Else
            .BufPos = p24
            .ThunkPos = q24
        End If
        If ParsePlain() Then
            ParseBlock = True
            Exit Function
        Else
            .BufPos = p24
            .ThunkPos = q24
        End If
    End With
End Function

Private Function ParseBlankLine() As Boolean
    Call ParseSp
    If ParseNewline() Then
        ParseBlankLine = True
    End If
End Function

Private Function ParseBlockQuote() As Boolean
    With ctx
        pvPushThunk ucsActVarAlloc, 1
        pvPushThunk ucsActResultClear
        If ParseBlockQuoteRaw() Then
            pvPushThunk ucsActVarSet, 1
            pvPushThunk ucsAct_1_BlockQuote, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -1
            ParseBlockQuote = True
        End If
    End With
End Function

Private Function ParseVerbatim() As Boolean
    Dim i228 As Long
    Dim p227 As Long
    Dim q227 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        pvPushThunk ucsActResultClear
        If ParseStartList() Then
            pvPushThunk ucsActVarSet, 1
            For i228 = 0 To LNG_MAXINT
                p227 = .BufPos
                q227 = .ThunkPos
                If Not (ParseVerbatimChunk()) Then
                    .BufPos = p227
                    .ThunkPos = q227
                    Exit For
                End If
                pvPushThunk ucsAct_1_Verbatim, .CaptureBegin, .CaptureEnd
            Next
            If i228 <> 0 Then
                pvPushThunk ucsAct_2_Verbatim, .CaptureBegin, .CaptureEnd
                pvPushThunk ucsActVarAlloc, -1
                ParseVerbatim = True
            End If
        End If
    End With
End Function

Private Function ParseNote() As Boolean
    Dim p2940 As Long
    Dim q2940 As Long
    Dim p2938 As Long
    Dim q2938 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        If extension(EXT_NOTES) Then
            Call ParseNonindentSpace
            pvPushThunk ucsActResultClear
            If ParseRawNoteReference() Then
                pvPushThunk ucsActVarSet, 1
                If .BufData(.BufPos) = 58 Then      ' ":"
                    .BufPos = .BufPos + 1
                    Call ParseSp
                    pvPushThunk ucsActResultClear
                    If ParseStartList() Then
                        pvPushThunk ucsActVarSet, 2
                        If ParseRawNoteBlock() Then
                            pvPushThunk ucsAct_1_Note, .CaptureBegin, .CaptureEnd
                            Do
                                p2940 = .BufPos
                                q2940 = .ThunkPos
                                p2938 = .BufPos
                                q2938 = .ThunkPos
                                If ParseIndent() Then
                                    .BufPos = p2938
                                    .ThunkPos = q2938
                                Else
                                    .BufPos = p2940
                                    .ThunkPos = q2940
                                    Exit Do
                                End If
                                If Not (ParseRawNoteBlock()) Then
                                    .BufPos = p2940
                                    .ThunkPos = q2940
                                    Exit Do
                                End If
                                pvPushThunk ucsAct_2_Note, .CaptureBegin, .CaptureEnd
                            Loop
                            pvPushThunk ucsAct_3_Note, .CaptureBegin, .CaptureEnd
                            pvPushThunk ucsActVarAlloc, -2
                            Call pvSetAdvance
                            ParseNote = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseReference() As Boolean
    Dim p2157 As Long
    Dim i2176 As Long
    Dim p2175 As Long
    Dim q2175 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 3
        Call ParseNonindentSpace
        p2157 = .BufPos
        If Not (.BufData(.BufPos) = 91 And .BufData(.BufPos + 1) = 93) Then ' "[]"
            .BufPos = p2157
            pvPushThunk ucsActResultClear
            If ParseLabel() Then
                pvPushThunk ucsActVarSet, 1
                If .BufData(.BufPos) = 58 Then      ' ":"
                    .BufPos = .BufPos + 1
                    Call ParseSpnl
                    pvPushThunk ucsActResultClear
                    If ParseRefSrc() Then
                        pvPushThunk ucsActVarSet, 2
                        pvPushThunk ucsActResultClear
                        Call ParseRefTitle
                        pvPushThunk ucsActVarSet, 3
                        For i2176 = 0 To LNG_MAXINT
                            p2175 = .BufPos
                            q2175 = .ThunkPos
                            If Not (ParseBlankLine()) Then
                                .BufPos = p2175
                                .ThunkPos = q2175
                                Exit For
                            End If
                        Next
                        If i2176 <> 0 Then
                            pvPushThunk ucsAct_1_Reference, .CaptureBegin, .CaptureEnd
                            pvPushThunk ucsActVarAlloc, -3
                            Call pvSetAdvance
                            ParseReference = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHorizontalRule() As Boolean
    Dim p251 As Long
    Dim p239 As Long
    Dim i266 As Long
    Dim p265 As Long
    Dim p249 As Long
    Dim p260 As Long

    With ctx
        Call ParseNonindentSpace
        p251 = .BufPos
        If .BufData(.BufPos) = 42 Then              ' "*"
            .BufPos = .BufPos + 1
            Call ParseSp
            If .BufData(.BufPos) = 42 Then          ' "*"
                .BufPos = .BufPos + 1
                Call ParseSp
                If .BufData(.BufPos) = 42 Then      ' "*"
                    .BufPos = .BufPos + 1
                    Do
                        p239 = .BufPos
                        Call ParseSp
                        If .BufData(.BufPos) = 42 Then ' "*"
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p239
                            Exit Do
                        End If
                    Loop
                    Call ParseSp
                    If ParseNewline() Then
                        For i266 = 0 To LNG_MAXINT
                            p265 = .BufPos
                            If Not (ParseBlankLine()) Then
                                .BufPos = p265
                                Exit For
                            End If
                        Next
                        If i266 <> 0 Then
                            pvPushThunk ucsAct_1_HorizontalRule, .CaptureBegin, .CaptureEnd
                            Call pvSetAdvance
                            ParseHorizontalRule = True
                            Exit Function
                        End If
                    End If
                Else
                    .BufPos = p251
                End If
            Else
                .BufPos = p251
            End If
        End If
        If .BufData(.BufPos) = 45 Then              ' "-"
            .BufPos = .BufPos + 1
            Call ParseSp
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                Call ParseSp
                If .BufData(.BufPos) = 45 Then      ' "-"
                    .BufPos = .BufPos + 1
                    Do
                        p249 = .BufPos
                        Call ParseSp
                        If .BufData(.BufPos) = 45 Then ' "-"
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p249
                            Exit Do
                        End If
                    Loop
                    Call ParseSp
                    If ParseNewline() Then
                        For i266 = 0 To LNG_MAXINT
                            p265 = .BufPos
                            If Not (ParseBlankLine()) Then
                                .BufPos = p265
                                Exit For
                            End If
                        Next
                        If i266 <> 0 Then
                            pvPushThunk ucsAct_1_HorizontalRule, .CaptureBegin, .CaptureEnd
                            Call pvSetAdvance
                            ParseHorizontalRule = True
                            Exit Function
                        End If
                    End If
                Else
                    .BufPos = p251
                End If
            Else
                .BufPos = p251
            End If
        End If
        If .BufData(.BufPos) = 95 Then              ' "_"
            .BufPos = .BufPos + 1
            Call ParseSp
            If .BufData(.BufPos) = 95 Then          ' "_"
                .BufPos = .BufPos + 1
                Call ParseSp
                If .BufData(.BufPos) = 95 Then      ' "_"
                    .BufPos = .BufPos + 1
                    Do
                        p260 = .BufPos
                        Call ParseSp
                        If .BufData(.BufPos) = 95 Then ' "_"
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p260
                            Exit Do
                        End If
                    Loop
                    Call ParseSp
                    If ParseNewline() Then
                        For i266 = 0 To LNG_MAXINT
                            p265 = .BufPos
                            If Not (ParseBlankLine()) Then
                                .BufPos = p265
                                Exit For
                            End If
                        Next
                        If i266 <> 0 Then
                            pvPushThunk ucsAct_1_HorizontalRule, .CaptureBegin, .CaptureEnd
                            Call pvSetAdvance
                            ParseHorizontalRule = True
                            Exit Function
                        End If
                    End If
                Else
                    .BufPos = p251
                End If
            Else
                .BufPos = p251
            End If
        End If
    End With
End Function

Private Function ParseHeading() As Boolean
    Dim p169 As Long
    Dim q169 As Long

    With ctx
        p169 = .BufPos
        q169 = .ThunkPos
        If ParseSetextHeading() Then
            ParseHeading = True
            Exit Function
        Else
            .BufPos = p169
            .ThunkPos = q169
        End If
        If ParseAtxHeading() Then
            ParseHeading = True
            Exit Function
        Else
            .BufPos = p169
            .ThunkPos = q169
        End If
    End With
End Function

Private Function ParseOrderedList() As Boolean
    Dim p395 As Long
    Dim p398 As Long
    Dim q398 As Long

    With ctx
        p395 = .BufPos
        If ParseEnumerator() Then
            .BufPos = p395
            p398 = .BufPos
            q398 = .ThunkPos
            If ParseListTight() Then
                pvPushThunk ucsAct_1_OrderedList, .CaptureBegin, .CaptureEnd
                ParseOrderedList = True
                Exit Function
            Else
                .BufPos = p398
                .ThunkPos = q398
            End If
            If ParseListLoose() Then
                pvPushThunk ucsAct_1_OrderedList, .CaptureBegin, .CaptureEnd
                ParseOrderedList = True
                Exit Function
            Else
                .BufPos = p398
                .ThunkPos = q398
            End If
        End If
    End With
End Function

Private Function ParseBulletList() As Boolean
    Dim p281 As Long
    Dim q281 As Long
    Dim p286 As Long
    Dim q286 As Long

    With ctx
        p281 = .BufPos
        q281 = .ThunkPos
        If ParseBullet() Then
            .BufPos = p281
            .ThunkPos = q281
            p286 = .BufPos
            q286 = .ThunkPos
            If ParseListTight() Then
                pvPushThunk ucsAct_1_BulletList, .CaptureBegin, .CaptureEnd
                ParseBulletList = True
                Exit Function
            Else
                .BufPos = p286
                .ThunkPos = q286
            End If
            If ParseListLoose() Then
                pvPushThunk ucsAct_1_BulletList, .CaptureBegin, .CaptureEnd
                ParseBulletList = True
                Exit Function
            Else
                .BufPos = p286
                .ThunkPos = q286
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlock() As Boolean
    Dim lCaptureBegin As Long
    Dim p1541 As Long
    Dim lCaptureEnd As Long
    Dim i1547 As Long
    Dim p1546 As Long

    With ctx
        lCaptureBegin = .BufPos
        p1541 = .BufPos
        If ParseHtmlBlockInTags() Then
            lCaptureEnd = .BufPos
            For i1547 = 0 To LNG_MAXINT
                p1546 = .BufPos
                If Not (ParseBlankLine()) Then
                    .BufPos = p1546
                    Exit For
                End If
            Next
            If i1547 <> 0 Then
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                pvPushThunk ucsAct_1_HtmlBlock, lCaptureBegin, lCaptureEnd
                ParseHtmlBlock = True
                Exit Function
            End If
        Else
            .BufPos = p1541
        End If
        If ParseHtmlComment() Then
            lCaptureEnd = .BufPos
            For i1547 = 0 To LNG_MAXINT
                p1546 = .BufPos
                If Not (ParseBlankLine()) Then
                    .BufPos = p1546
                    Exit For
                End If
            Next
            If i1547 <> 0 Then
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                pvPushThunk ucsAct_1_HtmlBlock, lCaptureBegin, lCaptureEnd
                ParseHtmlBlock = True
                Exit Function
            End If
        Else
            .BufPos = p1541
        End If
        If ParseHtmlBlockSelfClosing() Then
            lCaptureEnd = .BufPos
            For i1547 = 0 To LNG_MAXINT
                p1546 = .BufPos
                If Not (ParseBlankLine()) Then
                    .BufPos = p1546
                    Exit For
                End If
            Next
            If i1547 <> 0 Then
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                pvPushThunk ucsAct_1_HtmlBlock, lCaptureBegin, lCaptureEnd
                ParseHtmlBlock = True
                Exit Function
            End If
        Else
            .BufPos = p1541
        End If
    End With
End Function

Private Function ParseStyleBlock() As Boolean
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long
    Dim p1665 As Long

    With ctx
        lCaptureBegin = .BufPos
        If ParseInStyleTags() Then
            lCaptureEnd = .BufPos
            Do
                p1665 = .BufPos
                If Not (ParseBlankLine()) Then
                    .BufPos = p1665
                    Exit Do
                End If
            Loop
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_StyleBlock, lCaptureBegin, lCaptureEnd
            ParseStyleBlock = True
        End If
    End With
End Function

Private Function ParsePara() As Boolean
    Dim i55 As Long
    Dim p54 As Long
    Dim q54 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        Call ParseNonindentSpace
        pvPushThunk ucsActResultClear
        If ParseInlines() Then
            pvPushThunk ucsActVarSet, 1
            For i55 = 0 To LNG_MAXINT
                p54 = .BufPos
                q54 = .ThunkPos
                If Not (ParseBlankLine()) Then
                    .BufPos = p54
                    .ThunkPos = q54
                    Exit For
                End If
            Next
            If i55 <> 0 Then
                pvPushThunk ucsAct_1_Para, .CaptureBegin, .CaptureEnd
                pvPushThunk ucsActVarAlloc, -1
                ParsePara = True
            End If
        End If
    End With
End Function

Private Function ParsePlain() As Boolean
    With ctx
        pvPushThunk ucsActVarAlloc, 1
        pvPushThunk ucsActResultClear
        If ParseInlines() Then
            pvPushThunk ucsActVarSet, 1
            pvPushThunk ucsAct_1_Plain, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -1
            ParsePlain = True
        End If
    End With
End Function

Private Sub ParseNonindentSpace()
    With ctx
        If pvMatchString("   ") Then                ' "   "
            .BufPos = .BufPos + 3
            Call pvSetAdvance
        End If
        If .BufData(.BufPos) = 32 And .BufData(.BufPos + 1) = 32 Then ' "  "
            .BufPos = .BufPos + 2
            Call pvSetAdvance
        End If
        If .BufData(.BufPos) = 32 Then              ' " "
            .BufPos = .BufPos + 1
            Call pvSetAdvance
        End If
        Call pvSetAdvance
    End With
End Sub

Private Function ParseInlines() As Boolean
    Dim i1685 As Long
    Dim p1684 As Long
    Dim q1684 As Long
    Dim p1673 As Long
    Dim q1673 As Long
    Dim e1673 As String
    Dim p1682 As Long
    Dim q1682 As Long
    Dim p1686 As Long
    Dim q1686 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        If ParseStartList() Then
            pvPushThunk ucsActVarSet, 1
            For i1685 = 0 To LNG_MAXINT
                p1684 = .BufPos
                q1684 = .ThunkPos
                p1673 = .BufPos
                q1673 = .ThunkPos
                e1673 = .LastExpected
                If ParseEndline() Then
                    .BufPos = p1684
                    .ThunkPos = q1684
                    pvPushThunk ucsActResultClear
                    If ParseEndline() Then
                        pvPushThunk ucsActVarSet, 2
                        p1682 = .BufPos
                        q1682 = .ThunkPos
                        If ParseInline() Then
                            .BufPos = p1682
                            .ThunkPos = q1682
                            pvPushThunk ucsAct_2_Inlines, .CaptureBegin, .CaptureEnd
                        Else
                            .BufPos = p1684
                            .ThunkPos = q1684
                            .BufPos = p1684
                            .ThunkPos = q1684
                            Exit For
                        End If
                    Else
                        .BufPos = p1684
                        .ThunkPos = q1684
                        .BufPos = p1684
                        .ThunkPos = q1684
                        Exit For
                    End If
                Else
                    .BufPos = p1673
                    .ThunkPos = q1673
                    .LastExpected = e1673
                    If ParseInline() Then
                        pvPushThunk ucsAct_1_Inlines, .CaptureBegin, .CaptureEnd
                    Else
                        .BufPos = p1684
                        .ThunkPos = q1684
                        pvPushThunk ucsActResultClear
                        If ParseEndline() Then
                            pvPushThunk ucsActVarSet, 2
                            p1682 = .BufPos
                            q1682 = .ThunkPos
                            If ParseInline() Then
                                .BufPos = p1682
                                .ThunkPos = q1682
                                pvPushThunk ucsAct_2_Inlines, .CaptureBegin, .CaptureEnd
                            Else
                                .BufPos = p1684
                                .ThunkPos = q1684
                                .BufPos = p1684
                                .ThunkPos = q1684
                                Exit For
                            End If
                        Else
                            .BufPos = p1684
                            .ThunkPos = q1684
                            .BufPos = p1684
                            .ThunkPos = q1684
                            Exit For
                        End If
                    End If
                End If
            Next
            If i1685 <> 0 Then
                p1686 = .BufPos
                q1686 = .ThunkPos
                If Not (ParseEndline()) Then
                    .BufPos = p1686
                    .ThunkPos = q1686
                End If
                pvPushThunk ucsAct_3_Inlines, .CaptureBegin, .CaptureEnd
                pvPushThunk ucsActVarAlloc, -2
                ParseInlines = True
            End If
        End If
    End With
End Function

Private Function ParseAtxInline() As Boolean
    Dim p65 As Long
    Dim e65 As String
    Dim p73 As Long

    With ctx
        p65 = .BufPos
        e65 = .LastExpected
        If Not (ParseNewline()) Then
            .BufPos = p65
            .LastExpected = e65
            p73 = .BufPos
            Call ParseSp
            Do
                If .BufData(.BufPos) = 35 Then      ' "#"
                    .BufPos = .BufPos + 1
                Else
                    Exit Do
                End If
            Loop
            Call ParseSp
            If Not (ParseNewline()) Then
                .BufPos = p73
                If ParseInline() Then
                    Call pvSetAdvance
                    ParseAtxInline = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseNewline() As Boolean
    Dim p2550 As Long

    With ctx
        p2550 = .BufPos
        If .BufData(.BufPos) = 10 Then              ' "\n"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseNewline = True
            Exit Function
        End If
        If .BufData(.BufPos) = 13 Then              ' "\r"
            .BufPos = .BufPos + 1
            If .BufData(.BufPos) = 10 Then          ' "\n"
                .BufPos = .BufPos + 1
            End If
            Call pvSetAdvance
            ParseNewline = True
        End If
    End With
End Function

Private Sub ParseSp()
    Dim p2551 As Long

    With ctx
        Do
            p2551 = .BufPos
            If Not (ParseSpacechar()) Then
                .BufPos = p2551
                Exit Do
            End If
        Loop
    End With
End Sub

Private Function ParseInline() As Boolean
    Dim p1692 As Long
    Dim q1692 As Long

    With ctx
        p1692 = .BufPos
        q1692 = .ThunkPos
        If ParseStr() Then
            ParseInline = True
            Exit Function
        Else
            .BufPos = p1692
            .ThunkPos = q1692
        End If
        If ParseEndline() Then
            ParseInline = True
            Exit Function
        Else
            .BufPos = p1692
            .ThunkPos = q1692
        End If
        If ParseUlOrStarLine() Then
            ParseInline = True
            Exit Function
        Else
            .BufPos = p1692
            .ThunkPos = q1692
        End If
        If ParseSpace() Then
            ParseInline = True
            Exit Function
        Else
            .BufPos = p1692
            .ThunkPos = q1692
        End If
        If ParseStrong() Then
            ParseInline = True
            Exit Function
        Else
            .BufPos = p1692
            .ThunkPos = q1692
        End If
        If ParseEmph() Then
            ParseInline = True
            Exit Function
        Else
            .BufPos = p1692
            .ThunkPos = q1692
        End If
        If ParseStrike() Then
            ParseInline = True
            Exit Function
        Else
            .BufPos = p1692
            .ThunkPos = q1692
        End If
        If ParseImage() Then
            ParseInline = True
            Exit Function
        Else
            .BufPos = p1692
            .ThunkPos = q1692
        End If
        If ParseLink() Then
            ParseInline = True
            Exit Function
        Else
            .BufPos = p1692
            .ThunkPos = q1692
        End If
        If ParseNoteReference() Then
            ParseInline = True
            Exit Function
        Else
            .BufPos = p1692
            .ThunkPos = q1692
        End If
        If ParseInlineNote() Then
            ParseInline = True
            Exit Function
        Else
            .BufPos = p1692
            .ThunkPos = q1692
        End If
        If ParseCode() Then
            ParseInline = True
            Exit Function
        Else
            .BufPos = p1692
            .ThunkPos = q1692
        End If
        If ParseRawHtml() Then
            ParseInline = True
            Exit Function
        Else
            .BufPos = p1692
            .ThunkPos = q1692
        End If
        If ParseEntity() Then
            ParseInline = True
            Exit Function
        Else
            .BufPos = p1692
            .ThunkPos = q1692
        End If
        If ParseEscapedChar() Then
            ParseInline = True
            Exit Function
        Else
            .BufPos = p1692
            .ThunkPos = q1692
        End If
        If ParseSmart() Then
            ParseInline = True
            Exit Function
        Else
            .BufPos = p1692
            .ThunkPos = q1692
        End If
        If ParseSymbol() Then
            ParseInline = True
            Exit Function
        Else
            .BufPos = p1692
            .ThunkPos = q1692
        End If
    End With
End Function

Private Function ParseAtxStart() As Boolean
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        If pvMatchString("######") Then             ' "######"
            .BufPos = .BufPos + 6
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_AtxStart, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseAtxStart = True
            Exit Function
        End If
        If pvMatchString("#####") Then              ' "#####"
            .BufPos = .BufPos + 5
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_AtxStart, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseAtxStart = True
            Exit Function
        End If
        If pvMatchString("####") Then               ' "####"
            .BufPos = .BufPos + 4
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_AtxStart, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseAtxStart = True
            Exit Function
        End If
        If pvMatchString("###") Then                ' "###"
            .BufPos = .BufPos + 3
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_AtxStart, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseAtxStart = True
            Exit Function
        End If
        If .BufData(.BufPos) = 35 And .BufData(.BufPos + 1) = 35 Then ' "##"
            .BufPos = .BufPos + 2
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_AtxStart, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseAtxStart = True
            Exit Function
        End If
        If .BufData(.BufPos) = 35 Then              ' "#"
            .BufPos = .BufPos + 1
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_AtxStart, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseAtxStart = True
        End If
    End With
End Function

Private Function ParseAtxHeading() As Boolean
    Dim i102 As Long
    Dim p101 As Long
    Dim q101 As Long
    Dim p106 As Long
    Dim q106 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        If ParseAtxStart() Then
            pvPushThunk ucsActVarSet, 1
            Call ParseSp
            pvPushThunk ucsActResultClear
            If ParseStartList() Then
                pvPushThunk ucsActVarSet, 2
                For i102 = 0 To LNG_MAXINT
                    p101 = .BufPos
                    q101 = .ThunkPos
                    If Not (ParseAtxInline()) Then
                        .BufPos = p101
                        .ThunkPos = q101
                        Exit For
                    End If
                    pvPushThunk ucsAct_1_AtxHeading, .CaptureBegin, .CaptureEnd
                Next
                If i102 <> 0 Then
                    p106 = .BufPos
                    q106 = .ThunkPos
                    Call ParseSp
                    Do
                        If .BufData(.BufPos) = 35 Then ' "#"
                            .BufPos = .BufPos + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    Call ParseSp
                    If ParseNewline() Then
                        pvPushThunk ucsAct_2_AtxHeading, .CaptureBegin, .CaptureEnd
                        pvPushThunk ucsActVarAlloc, -2
                        Call pvSetAdvance
                        ParseAtxHeading = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseSetextHeading() As Boolean
    Dim p116 As Long
    Dim q116 As Long

    With ctx
        p116 = .BufPos
        q116 = .ThunkPos
        If ParseSetextHeading1() Then
            ParseSetextHeading = True
            Exit Function
        Else
            .BufPos = p116
            .ThunkPos = q116
        End If
        If ParseSetextHeading2() Then
            ParseSetextHeading = True
            Exit Function
        Else
            .BufPos = p116
            .ThunkPos = q116
        End If
    End With
End Function

Private Function ParseSetextHeading1() As Boolean
    Dim p131 As Long
    Dim i143 As Long
    Dim p141 As Long
    Dim q141 As Long
    Dim p139 As Long
    Dim q139 As Long
    Dim e139 As String

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        p131 = .BufPos
        If ParseRawLine() Then
            If ParseSetextBottom1() Then
                .BufPos = p131
                pvPushThunk ucsActResultClear
                If ParseStartList() Then
                    pvPushThunk ucsActVarSet, 1
                    For i143 = 0 To LNG_MAXINT
                        p141 = .BufPos
                        q141 = .ThunkPos
                        p139 = .BufPos
                        q139 = .ThunkPos
                        e139 = .LastExpected
                        If ParseEndline() Then
                            .BufPos = p141
                            .ThunkPos = q141
                            Exit For
                        Else
                            .BufPos = p139
                            .ThunkPos = q139
                            .LastExpected = e139
                        End If
                        If Not (ParseInline()) Then
                            .BufPos = p141
                            .ThunkPos = q141
                            Exit For
                        End If
                        pvPushThunk ucsAct_1_SetextHeading1, .CaptureBegin, .CaptureEnd
                    Next
                    If i143 <> 0 Then
                        Call ParseSp
                        If ParseNewline() Then
                            If ParseSetextBottom1() Then
                                pvPushThunk ucsAct_2_SetextHeading1, .CaptureBegin, .CaptureEnd
                                pvPushThunk ucsActVarAlloc, -1
                                ParseSetextHeading1 = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseSetextHeading2() As Boolean
    Dim p151 As Long
    Dim i162 As Long
    Dim p160 As Long
    Dim q160 As Long
    Dim p158 As Long
    Dim q158 As Long
    Dim e158 As String

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        p151 = .BufPos
        If ParseRawLine() Then
            If ParseSetextBottom2() Then
                .BufPos = p151
                pvPushThunk ucsActResultClear
                If ParseStartList() Then
                    pvPushThunk ucsActVarSet, 1
                    For i162 = 0 To LNG_MAXINT
                        p160 = .BufPos
                        q160 = .ThunkPos
                        p158 = .BufPos
                        q158 = .ThunkPos
                        e158 = .LastExpected
                        If ParseEndline() Then
                            .BufPos = p160
                            .ThunkPos = q160
                            Exit For
                        Else
                            .BufPos = p158
                            .ThunkPos = q158
                            .LastExpected = e158
                        End If
                        If Not (ParseInline()) Then
                            .BufPos = p160
                            .ThunkPos = q160
                            Exit For
                        End If
                        pvPushThunk ucsAct_1_SetextHeading2, .CaptureBegin, .CaptureEnd
                    Next
                    If i162 <> 0 Then
                        Call ParseSp
                        If ParseNewline() Then
                            If ParseSetextBottom2() Then
                                pvPushThunk ucsAct_2_SetextHeading2, .CaptureBegin, .CaptureEnd
                                pvPushThunk ucsActVarAlloc, -1
                                ParseSetextHeading2 = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseSetextBottom1() As Boolean
    Dim i119 As Long

    With ctx
        For i119 = 0 To LNG_MAXINT
            If .BufData(.BufPos) = 61 Then          ' "="
                .BufPos = .BufPos + 1
            Else
                Exit For
            End If
        Next
        If i119 <> 0 Then
            If ParseNewline() Then
                Call pvSetAdvance
                ParseSetextBottom1 = True
            End If
        End If
    End With
End Function

Private Function ParseSetextBottom2() As Boolean
    Dim i124 As Long

    With ctx
        For i124 = 0 To LNG_MAXINT
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
            Else
                Exit For
            End If
        Next
        If i124 <> 0 Then
            If ParseNewline() Then
                Call pvSetAdvance
                ParseSetextBottom2 = True
            End If
        End If
    End With
End Function

Private Function ParseRawLine() As Boolean
    Dim p2782 As Long
    Dim lCaptureBegin As Long
    Dim p2770 As Long
    Dim p2767 As Long
    Dim p2769 As Long
    Dim lCaptureEnd As Long
    Dim i2778 As Long

    With ctx
        p2782 = .BufPos
        lCaptureBegin = .BufPos
        Do
            p2770 = .BufPos
            p2767 = .BufPos
            If .BufData(.BufPos) = 13 Then          ' "\r"
                .BufPos = p2770
                Exit Do
            Else
                .BufPos = p2767
            End If
            p2769 = .BufPos
            If .BufData(.BufPos) = 10 Then          ' "\n"
                .BufPos = p2770
                Exit Do
            Else
                .BufPos = p2769
            End If
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
            Else
                .BufPos = p2770
                Exit Do
            End If
        Loop
        If ParseNewline() Then
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            Call pvSetAdvance
            ParseRawLine = True
            Exit Function
        Else
            .BufPos = p2782
        End If
        lCaptureBegin = .BufPos
        For i2778 = 0 To LNG_MAXINT
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
            Else
                Exit For
            End If
        Next
        If i2778 <> 0 Then
            lCaptureEnd = .BufPos
            If ParseEof() Then
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call pvSetAdvance
                ParseRawLine = True
                Exit Function
            Else
                .BufPos = p2782
            End If
        Else
            .BufPos = p2782
        End If
    End With
End Function

Private Function ParseEndline() As Boolean
    Dim p1786 As Long
    Dim q1786 As Long

    With ctx
        p1786 = .BufPos
        q1786 = .ThunkPos
        If ParseLineBreak() Then
            ParseEndline = True
            Exit Function
        Else
            .BufPos = p1786
            .ThunkPos = q1786
        End If
        If ParseTerminalEndline() Then
            ParseEndline = True
            Exit Function
        Else
            .BufPos = p1786
            .ThunkPos = q1786
        End If
        If ParseNormalEndline() Then
            ParseEndline = True
            Exit Function
        Else
            .BufPos = p1786
            .ThunkPos = q1786
        End If
    End With
End Function

Private Function ParseBlockQuoteRaw() As Boolean
    Dim i199 As Long
    Dim p183 As Long
    Dim q183 As Long
    Dim p191 As Long
    Dim q191 As Long
    Dim p188 As Long
    Dim q188 As Long
    Dim p190 As Long
    Dim q190 As Long
    Dim e190 As String
    Dim p197 As Long
    Dim q197 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        pvPushThunk ucsActResultClear
        If ParseStartList() Then
            pvPushThunk ucsActVarSet, 1
            For i199 = 0 To LNG_MAXINT
                p183 = .BufPos
                q183 = .ThunkPos
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                Else
                    Exit For
                End If
                If .BufData(.BufPos) = 32 Then      ' " "
                    .BufPos = .BufPos + 1
                End If
                If Not (ParseLine()) Then
                    .BufPos = p183
                    .ThunkPos = q183
                    Exit For
                End If
                pvPushThunk ucsAct_1_BlockQuoteRaw, .CaptureBegin, .CaptureEnd
                Do
                    p191 = .BufPos
                    q191 = .ThunkPos
                    p188 = .BufPos
                    q188 = .ThunkPos
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = p191
                        .ThunkPos = q191
                        Exit Do
                    Else
                        .BufPos = p188
                        .ThunkPos = q188
                    End If
                    p190 = .BufPos
                    q190 = .ThunkPos
                    e190 = .LastExpected
                    If ParseBlankLine() Then
                        .BufPos = p191
                        .ThunkPos = q191
                        Exit Do
                    Else
                        .BufPos = p190
                        .ThunkPos = q190
                        .LastExpected = e190
                    End If
                    If Not (ParseLine()) Then
                        .BufPos = p191
                        .ThunkPos = q191
                        Exit Do
                    End If
                    pvPushThunk ucsAct_2_BlockQuoteRaw, .CaptureBegin, .CaptureEnd
                Loop
                Do
                    p197 = .BufPos
                    q197 = .ThunkPos
                    If Not (ParseBlankLine()) Then
                        .BufPos = p197
                        .ThunkPos = q197
                        Exit Do
                    End If
                    pvPushThunk ucsAct_3_BlockQuoteRaw, .CaptureBegin, .CaptureEnd
                Loop
            Next
            If i199 <> 0 Then
                pvPushThunk ucsAct_4_BlockQuoteRaw, .CaptureBegin, .CaptureEnd
                pvPushThunk ucsActVarAlloc, -1
                Call pvSetAdvance
                ParseBlockQuoteRaw = True
            End If
        End If
    End With
End Function

Private Function ParseLine() As Boolean
    With ctx
        If ParseRawLine() Then
            pvPushThunk ucsAct_1_Line, .CaptureBegin, .CaptureEnd
            ParseLine = True
        End If
    End With
End Function

Private Function ParseNonblankIndentedLine() As Boolean
    Dim p203 As Long
    Dim e203 As String

    With ctx
        p203 = .BufPos
        e203 = .LastExpected
        If Not (ParseBlankLine()) Then
            .BufPos = p203
            .LastExpected = e203
            If ParseIndentedLine() Then
                ParseNonblankIndentedLine = True
            End If
        End If
    End With
End Function

Private Function ParseIndentedLine() As Boolean
    If ParseIndent() Then
        If ParseLine() Then
            ParseIndentedLine = True
        End If
    End If
End Function

Private Function ParseVerbatimChunk() As Boolean
    Dim p214 As Long
    Dim q214 As Long
    Dim i219 As Long
    Dim p218 As Long
    Dim q218 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        pvPushThunk ucsActResultClear
        If ParseStartList() Then
            pvPushThunk ucsActVarSet, 1
            Do
                p214 = .BufPos
                q214 = .ThunkPos
                If Not (ParseBlankLine()) Then
                    .BufPos = p214
                    .ThunkPos = q214
                    Exit Do
                End If
                pvPushThunk ucsAct_1_VerbatimChunk, .CaptureBegin, .CaptureEnd
            Loop
            For i219 = 0 To LNG_MAXINT
                p218 = .BufPos
                q218 = .ThunkPos
                If Not (ParseNonblankIndentedLine()) Then
                    .BufPos = p218
                    .ThunkPos = q218
                    Exit For
                End If
                pvPushThunk ucsAct_2_VerbatimChunk, .CaptureBegin, .CaptureEnd
            Next
            If i219 <> 0 Then
                pvPushThunk ucsAct_3_VerbatimChunk, .CaptureBegin, .CaptureEnd
                pvPushThunk ucsActVarAlloc, -1
                ParseVerbatimChunk = True
            End If
        End If
    End With
End Function

Private Function ParseBullet() As Boolean
    Dim p270 As Long
    Dim q270 As Long
    Dim e270 As String
    Dim i279 As Long
    Dim p278 As Long
    Dim q278 As Long

    With ctx
        p270 = .BufPos
        q270 = .ThunkPos
        e270 = .LastExpected
        If Not (ParseHorizontalRule()) Then
            .BufPos = p270
            .ThunkPos = q270
            .LastExpected = e270
            Call ParseNonindentSpace
            If .BufData(.BufPos) = 43 Then          ' "+"
                .BufPos = .BufPos + 1
                For i279 = 0 To LNG_MAXINT
                    p278 = .BufPos
                    q278 = .ThunkPos
                    If Not (ParseSpacechar()) Then
                        .BufPos = p278
                        .ThunkPos = q278
                        Exit For
                    End If
                Next
                If i279 <> 0 Then
                    Call pvSetAdvance
                    ParseBullet = True
                End If
            End If
            If .BufData(.BufPos) = 42 Then          ' "*"
                .BufPos = .BufPos + 1
                For i279 = 0 To LNG_MAXINT
                    p278 = .BufPos
                    q278 = .ThunkPos
                    If Not (ParseSpacechar()) Then
                        .BufPos = p278
                        .ThunkPos = q278
                        Exit For
                    End If
                Next
                If i279 <> 0 Then
                    Call pvSetAdvance
                    ParseBullet = True
                End If
            End If
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                For i279 = 0 To LNG_MAXINT
                    p278 = .BufPos
                    q278 = .ThunkPos
                    If Not (ParseSpacechar()) Then
                        .BufPos = p278
                        .ThunkPos = q278
                        Exit For
                    End If
                Next
                If i279 <> 0 Then
                    Call pvSetAdvance
                    ParseBullet = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseSpacechar() As Boolean
    With ctx
        If .BufData(.BufPos) = 32 Then              ' " "
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseSpacechar = True
            Exit Function
        End If
        If .BufData(.BufPos) = 9 Then               ' "\t"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseSpacechar = True
        End If
    End With
End Function

Private Function ParseListTight() As Boolean
    Dim i297 As Long
    Dim p296 As Long
    Dim q296 As Long
    Dim p298 As Long
    Dim q298 As Long
    Dim p304 As Long
    Dim q304 As Long
    Dim p303 As Long
    Dim q303 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        pvPushThunk ucsActResultClear
        If ParseStartList() Then
            pvPushThunk ucsActVarSet, 1
            For i297 = 0 To LNG_MAXINT
                p296 = .BufPos
                q296 = .ThunkPos
                If Not (ParseListItemTight()) Then
                    .BufPos = p296
                    .ThunkPos = q296
                    Exit For
                End If
                pvPushThunk ucsAct_1_ListTight, .CaptureBegin, .CaptureEnd
            Next
            If i297 <> 0 Then
                Do
                    p298 = .BufPos
                    q298 = .ThunkPos
                    If Not (ParseBlankLine()) Then
                        .BufPos = p298
                        .ThunkPos = q298
                        Exit Do
                    End If
                Loop
                p304 = .BufPos
                q304 = .ThunkPos
                p303 = .BufPos
                q303 = .ThunkPos
                If Not (ParseBullet()) Then
                    .BufPos = p303
                    .ThunkPos = q303
                    If Not (ParseEnumerator()) Then
                        .BufPos = p303
                        .ThunkPos = q303
                        .BufPos = p304
                        .ThunkPos = q304
                        pvPushThunk ucsAct_2_ListTight, .CaptureBegin, .CaptureEnd
                        pvPushThunk ucsActVarAlloc, -1
                        ParseListTight = True
                        Exit Function
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseListLoose() As Boolean
    Dim i318 As Long
    Dim p314 As Long
    Dim q314 As Long
    Dim p315 As Long
    Dim q315 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        If ParseStartList() Then
            pvPushThunk ucsActVarSet, 1
            For i318 = 0 To LNG_MAXINT
                p314 = .BufPos
                q314 = .ThunkPos
                pvPushThunk ucsActResultClear
                If ParseListItem() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p314
                    .ThunkPos = q314
                    Exit For
                End If
                Do
                    p315 = .BufPos
                    q315 = .ThunkPos
                    If Not (ParseBlankLine()) Then
                        .BufPos = p315
                        .ThunkPos = q315
                        Exit Do
                    End If
                Loop
                pvPushThunk ucsAct_1_ListLoose, .CaptureBegin, .CaptureEnd
            Next
            If i318 <> 0 Then
                pvPushThunk ucsAct_2_ListLoose, .CaptureBegin, .CaptureEnd
                pvPushThunk ucsActVarAlloc, -2
                ParseListLoose = True
            End If
        End If
    End With
End Function

Private Function ParseListItemTight() As Boolean
    Dim p339 As Long
    Dim q339 As Long
    Dim p350 As Long
    Dim q350 As Long
    Dim p348 As Long
    Dim q348 As Long
    Dim e348 As String
    Dim p354 As Long
    Dim q354 As Long
    Dim e354 As String

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        p339 = .BufPos
        q339 = .ThunkPos
        If ParseBullet() Then
            pvPushThunk ucsActResultClear
            If ParseStartList() Then
                pvPushThunk ucsActVarSet, 1
                If ParseListBlock() Then
                    pvPushThunk ucsAct_1_ListItemTight, .CaptureBegin, .CaptureEnd
                    Do
                        p350 = .BufPos
                        q350 = .ThunkPos
                        p348 = .BufPos
                        q348 = .ThunkPos
                        e348 = .LastExpected
                        If ParseBlankLine() Then
                            .BufPos = p350
                            .ThunkPos = q350
                            Exit Do
                        Else
                            .BufPos = p348
                            .ThunkPos = q348
                            .LastExpected = e348
                        End If
                        If Not (ParseListContinuationBlock()) Then
                            .BufPos = p350
                            .ThunkPos = q350
                            Exit Do
                        End If
                        pvPushThunk ucsAct_2_ListItemTight, .CaptureBegin, .CaptureEnd
                    Loop
                    p354 = .BufPos
                    q354 = .ThunkPos
                    e354 = .LastExpected
                    If Not (ParseListContinuationBlock()) Then
                        .BufPos = p354
                        .ThunkPos = q354
                        .LastExpected = e354
                        pvPushThunk ucsAct_3_ListItemTight, .CaptureBegin, .CaptureEnd
                        pvPushThunk ucsActVarAlloc, -1
                        ParseListItemTight = True
                        Exit Function
                    End If
                End If
            End If
        Else
            .BufPos = p339
            .ThunkPos = q339
        End If
        If ParseEnumerator() Then
            pvPushThunk ucsActResultClear
            If ParseStartList() Then
                pvPushThunk ucsActVarSet, 1
                If ParseListBlock() Then
                    pvPushThunk ucsAct_1_ListItemTight, .CaptureBegin, .CaptureEnd
                    Do
                        p350 = .BufPos
                        q350 = .ThunkPos
                        p348 = .BufPos
                        q348 = .ThunkPos
                        e348 = .LastExpected
                        If ParseBlankLine() Then
                            .BufPos = p350
                            .ThunkPos = q350
                            Exit Do
                        Else
                            .BufPos = p348
                            .ThunkPos = q348
                            .LastExpected = e348
                        End If
                        If Not (ParseListContinuationBlock()) Then
                            .BufPos = p350
                            .ThunkPos = q350
                            Exit Do
                        End If
                        pvPushThunk ucsAct_2_ListItemTight, .CaptureBegin, .CaptureEnd
                    Loop
                    p354 = .BufPos
                    q354 = .ThunkPos
                    e354 = .LastExpected
                    If Not (ParseListContinuationBlock()) Then
                        .BufPos = p354
                        .ThunkPos = q354
                        .LastExpected = e354
                        pvPushThunk ucsAct_3_ListItemTight, .CaptureBegin, .CaptureEnd
                        pvPushThunk ucsActVarAlloc, -1
                        ParseListItemTight = True
                        Exit Function
                    End If
                End If
            End If
        Else
            .BufPos = p339
            .ThunkPos = q339
        End If
    End With
End Function

Private Function ParseEnumerator() As Boolean
    Dim i389 As Long
    Dim i393 As Long
    Dim p392 As Long

    With ctx
        Call ParseNonindentSpace
        For i389 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i389 <> 0 Then
            If .BufData(.BufPos) = 46 Then          ' "."
                .BufPos = .BufPos + 1
                For i393 = 0 To LNG_MAXINT
                    p392 = .BufPos
                    If Not (ParseSpacechar()) Then
                        .BufPos = p392
                        Exit For
                    End If
                Next
                If i393 <> 0 Then
                    Call pvSetAdvance
                    ParseEnumerator = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseListItem() As Boolean
    Dim p322 As Long
    Dim q322 As Long
    Dim p334 As Long
    Dim q334 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        p322 = .BufPos
        q322 = .ThunkPos
        If ParseBullet() Then
            pvPushThunk ucsActResultClear
            If ParseStartList() Then
                pvPushThunk ucsActVarSet, 1
                If ParseListBlock() Then
                    pvPushThunk ucsAct_1_ListItem, .CaptureBegin, .CaptureEnd
                    Do
                        p334 = .BufPos
                        q334 = .ThunkPos
                        If Not (ParseListContinuationBlock()) Then
                            .BufPos = p334
                            .ThunkPos = q334
                            Exit Do
                        End If
                        pvPushThunk ucsAct_2_ListItem, .CaptureBegin, .CaptureEnd
                    Loop
                    pvPushThunk ucsAct_3_ListItem, .CaptureBegin, .CaptureEnd
                    pvPushThunk ucsActVarAlloc, -1
                    ParseListItem = True
                    Exit Function
                End If
            End If
        Else
            .BufPos = p322
            .ThunkPos = q322
        End If
        If ParseEnumerator() Then
            pvPushThunk ucsActResultClear
            If ParseStartList() Then
                pvPushThunk ucsActVarSet, 1
                If ParseListBlock() Then
                    pvPushThunk ucsAct_1_ListItem, .CaptureBegin, .CaptureEnd
                    Do
                        p334 = .BufPos
                        q334 = .ThunkPos
                        If Not (ParseListContinuationBlock()) Then
                            .BufPos = p334
                            .ThunkPos = q334
                            Exit Do
                        End If
                        pvPushThunk ucsAct_2_ListItem, .CaptureBegin, .CaptureEnd
                    Loop
                    pvPushThunk ucsAct_3_ListItem, .CaptureBegin, .CaptureEnd
                    pvPushThunk ucsActVarAlloc, -1
                    ParseListItem = True
                    Exit Function
                End If
            End If
        Else
            .BufPos = p322
            .ThunkPos = q322
        End If
    End With
End Function

Private Function ParseListBlock() As Boolean
    Dim p361 As Long
    Dim q361 As Long
    Dim e361 As String
    Dim p367 As Long
    Dim q367 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        pvPushThunk ucsActResultClear
        If ParseStartList() Then
            pvPushThunk ucsActVarSet, 1
            p361 = .BufPos
            q361 = .ThunkPos
            e361 = .LastExpected
            If Not (ParseBlankLine()) Then
                .BufPos = p361
                .ThunkPos = q361
                .LastExpected = e361
                If ParseLine() Then
                    pvPushThunk ucsAct_1_ListBlock, .CaptureBegin, .CaptureEnd
                    Do
                        p367 = .BufPos
                        q367 = .ThunkPos
                        If Not (ParseListBlockLine()) Then
                            .BufPos = p367
                            .ThunkPos = q367
                            Exit Do
                        End If
                        pvPushThunk ucsAct_2_ListBlock, .CaptureBegin, .CaptureEnd
                    Loop
                    pvPushThunk ucsAct_3_ListBlock, .CaptureBegin, .CaptureEnd
                    pvPushThunk ucsActVarAlloc, -1
                    ParseListBlock = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseListContinuationBlock() As Boolean
    Dim lCaptureBegin As Long
    Dim p375 As Long
    Dim q375 As Long
    Dim lCaptureEnd As Long
    Dim i385 As Long
    Dim p383 As Long
    Dim q383 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        pvPushThunk ucsActResultClear
        If ParseStartList() Then
            pvPushThunk ucsActVarSet, 1
            lCaptureBegin = .BufPos
            Do
                p375 = .BufPos
                q375 = .ThunkPos
                If Not (ParseBlankLine()) Then
                    .BufPos = p375
                    .ThunkPos = q375
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_ListContinuationBlock, lCaptureBegin, lCaptureEnd
            For i385 = 0 To LNG_MAXINT
                p383 = .BufPos
                q383 = .ThunkPos
                If Not (ParseIndent()) Then
                    .BufPos = p383
                    .ThunkPos = q383
                    Exit For
                End If
                If Not (ParseListBlock()) Then
                    .BufPos = p383
                    .ThunkPos = q383
                    Exit For
                End If
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                pvPushThunk ucsAct_2_ListContinuationBlock, lCaptureBegin, lCaptureEnd
            Next
            If i385 <> 0 Then
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                pvPushThunk ucsAct_3_ListContinuationBlock, lCaptureBegin, lCaptureEnd
                pvPushThunk ucsActVarAlloc, -1
                ParseListContinuationBlock = True
            End If
        End If
    End With
End Function

Private Function ParseListBlockLine() As Boolean
    Dim p402 As Long
    Dim e402 As String
    Dim p409 As Long
    Dim q409 As Long
    Dim p403 As Long
    Dim p407 As Long
    Dim q407 As Long
    Dim p412 As Long
    Dim q412 As Long
    Dim e412 As String

    With ctx
        p402 = .BufPos
        e402 = .LastExpected
        If Not (ParseBlankLine()) Then
            .BufPos = p402
            .LastExpected = e402
            p409 = .BufPos
            q409 = .ThunkPos
            p403 = .BufPos
            If Not (ParseIndent()) Then
                .BufPos = p403
            End If
            p407 = .BufPos
            q407 = .ThunkPos
            If Not (ParseBullet()) Then
                .BufPos = p407
                .ThunkPos = q407
                If Not (ParseEnumerator()) Then
                    .BufPos = p407
                    .ThunkPos = q407
                    .BufPos = p409
                    .ThunkPos = q409
                    p412 = .BufPos
                    q412 = .ThunkPos
                    e412 = .LastExpected
                    If Not (ParseHorizontalRule()) Then
                        .BufPos = p412
                        .ThunkPos = q412
                        .LastExpected = e412
                        If ParseOptionallyIndentedLine() Then
                            ParseListBlockLine = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseIndent() As Boolean
    With ctx
        If .BufData(.BufPos) = 9 Then               ' "\t"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseIndent = True
            Exit Function
        End If
        If pvMatchString("    ") Then               ' "    "
            .BufPos = .BufPos + 4
            Call pvSetAdvance
            ParseIndent = True
        End If
    End With
End Function

Private Function ParseOptionallyIndentedLine() As Boolean
    Dim p2754 As Long

    With ctx
        p2754 = .BufPos
        If Not (ParseIndent()) Then
            .BufPos = p2754
        End If
        If ParseLine() Then
            ParseOptionallyIndentedLine = True
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenAddress() As Boolean
    Dim p425 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("address") Then        ' "address"
                .BufPos = .BufPos + 7
                Call ParseSpnl
                Do
                    p425 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p425
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenAddress = True
                End If
            End If
            If pvMatchString("ADDRESS") Then        ' "ADDRESS"
                .BufPos = .BufPos + 7
                Call ParseSpnl
                Do
                    p425 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p425
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenAddress = True
                End If
            End If
        End If
    End With
End Function

Private Sub ParseSpnl()
    Dim p2556 As Long

    With ctx
        Call ParseSp
        p2556 = .BufPos
        If ParseNewline() Then
            Call ParseSp
        Else
            .BufPos = p2556
        End If
    End With
End Sub

Private Function ParseHtmlAttribute() As Boolean
    Dim i2497 As Long
    Dim p2496 As Long
    Dim p2502 As Long
    Dim p2509 As Long
    Dim i2508 As Long
    Dim p2507 As Long
    Dim p2505 As Long

    With ctx
        For i2497 = 0 To LNG_MAXINT
            p2496 = .BufPos
            Select Case .BufData(.BufPos)
            Case 65 To 90, 97 To 122, 48 To 57      ' [A-Za-z0-9]
                .BufPos = .BufPos + 1
            Case Else
                If .BufData(.BufPos) = 45 Then      ' "-"
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p2496
                    Exit For
                End If
            End Select
        Next
        If i2497 <> 0 Then
            Call ParseSpnl
            p2502 = .BufPos
            If .BufData(.BufPos) = 61 Then          ' "="
                .BufPos = .BufPos + 1
                Call ParseSpnl
                p2509 = .BufPos
                If ParseQuoted() Then
                    GoTo L14
                Else
                    .BufPos = p2509
                End If
                For i2508 = 0 To LNG_MAXINT
                    p2507 = .BufPos
                    p2505 = .BufPos
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = p2507
                        Exit For
                    Else
                        .BufPos = p2505
                    End If
                    If Not (ParseNonspacechar()) Then
                        .BufPos = p2507
                        Exit For
                    End If
                Next
                If i2508 <> 0 Then
                    GoTo L14
                End If
                .BufPos = p2502
            End If
L14:
            Call ParseSpnl
            Call pvSetAdvance
            ParseHtmlAttribute = True
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseAddress() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("address") Then    ' "address"
                    .BufPos = .BufPos + 7
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseAddress = True
                    End If
                End If
                If pvMatchString("ADDRESS") Then    ' "ADDRESS"
                    .BufPos = .BufPos + 7
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseAddress = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockAddress() As Boolean
    Dim p445 As Long
    Dim p442 As Long
    Dim e442 As String

    With ctx
        If ParseHtmlBlockOpenAddress() Then
            Do
                p445 = .BufPos
                If Not (ParseHtmlBlockAddress()) Then
                    .BufPos = p445
                    p442 = .BufPos
                    e442 = .LastExpected
                    If ParseHtmlBlockCloseAddress() Then
                        .BufPos = p445
                        .BufPos = p445
                        Exit Do
                    Else
                        .BufPos = p442
                        .LastExpected = e442
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p445
                            .BufPos = p445
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseAddress() Then
                ParseHtmlBlockAddress = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenBlockquote() As Boolean
    Dim p457 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("blockquote") Then     ' "blockquote"
                .BufPos = .BufPos + 10
                Call ParseSpnl
                Do
                    p457 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p457
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenBlockquote = True
                End If
            End If
            If pvMatchString("BLOCKQUOTE") Then     ' "BLOCKQUOTE"
                .BufPos = .BufPos + 10
                Call ParseSpnl
                Do
                    p457 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p457
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenBlockquote = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseBlockquote() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("blockquote") Then ' "blockquote"
                    .BufPos = .BufPos + 10
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseBlockquote = True
                    End If
                End If
                If pvMatchString("BLOCKQUOTE") Then ' "BLOCKQUOTE"
                    .BufPos = .BufPos + 10
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseBlockquote = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockBlockquote() As Boolean
    Dim p477 As Long
    Dim p474 As Long
    Dim e474 As String

    With ctx
        If ParseHtmlBlockOpenBlockquote() Then
            Do
                p477 = .BufPos
                If Not (ParseHtmlBlockBlockquote()) Then
                    .BufPos = p477
                    p474 = .BufPos
                    e474 = .LastExpected
                    If ParseHtmlBlockCloseBlockquote() Then
                        .BufPos = p477
                        .BufPos = p477
                        Exit Do
                    Else
                        .BufPos = p474
                        .LastExpected = e474
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p477
                            .BufPos = p477
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseBlockquote() Then
                ParseHtmlBlockBlockquote = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenCenter() As Boolean
    Dim p489 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("center") Then         ' "center"
                .BufPos = .BufPos + 6
                Call ParseSpnl
                Do
                    p489 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p489
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenCenter = True
                End If
            End If
            If pvMatchString("CENTER") Then         ' "CENTER"
                .BufPos = .BufPos + 6
                Call ParseSpnl
                Do
                    p489 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p489
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenCenter = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseCenter() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("center") Then     ' "center"
                    .BufPos = .BufPos + 6
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseCenter = True
                    End If
                End If
                If pvMatchString("CENTER") Then     ' "CENTER"
                    .BufPos = .BufPos + 6
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseCenter = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCenter() As Boolean
    Dim p509 As Long
    Dim p506 As Long
    Dim e506 As String

    With ctx
        If ParseHtmlBlockOpenCenter() Then
            Do
                p509 = .BufPos
                If Not (ParseHtmlBlockCenter()) Then
                    .BufPos = p509
                    p506 = .BufPos
                    e506 = .LastExpected
                    If ParseHtmlBlockCloseCenter() Then
                        .BufPos = p509
                        .BufPos = p509
                        Exit Do
                    Else
                        .BufPos = p506
                        .LastExpected = e506
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p509
                            .BufPos = p509
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseCenter() Then
                ParseHtmlBlockCenter = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenDir() As Boolean
    Dim p521 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("dir") Then            ' "dir"
                .BufPos = .BufPos + 3
                Call ParseSpnl
                Do
                    p521 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p521
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenDir = True
                End If
            End If
            If pvMatchString("DIR") Then            ' "DIR"
                .BufPos = .BufPos + 3
                Call ParseSpnl
                Do
                    p521 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p521
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenDir = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseDir() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("dir") Then        ' "dir"
                    .BufPos = .BufPos + 3
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseDir = True
                    End If
                End If
                If pvMatchString("DIR") Then        ' "DIR"
                    .BufPos = .BufPos + 3
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseDir = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockDir() As Boolean
    Dim p541 As Long
    Dim p538 As Long
    Dim e538 As String

    With ctx
        If ParseHtmlBlockOpenDir() Then
            Do
                p541 = .BufPos
                If Not (ParseHtmlBlockDir()) Then
                    .BufPos = p541
                    p538 = .BufPos
                    e538 = .LastExpected
                    If ParseHtmlBlockCloseDir() Then
                        .BufPos = p541
                        .BufPos = p541
                        Exit Do
                    Else
                        .BufPos = p538
                        .LastExpected = e538
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p541
                            .BufPos = p541
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseDir() Then
                ParseHtmlBlockDir = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenDiv() As Boolean
    Dim p553 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("div") Then            ' "div"
                .BufPos = .BufPos + 3
                Call ParseSpnl
                Do
                    p553 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p553
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenDiv = True
                End If
            End If
            If pvMatchString("DIV") Then            ' "DIV"
                .BufPos = .BufPos + 3
                Call ParseSpnl
                Do
                    p553 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p553
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenDiv = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseDiv() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("div") Then        ' "div"
                    .BufPos = .BufPos + 3
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseDiv = True
                    End If
                End If
                If pvMatchString("DIV") Then        ' "DIV"
                    .BufPos = .BufPos + 3
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseDiv = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockDiv() As Boolean
    Dim p573 As Long
    Dim p570 As Long
    Dim e570 As String

    With ctx
        If ParseHtmlBlockOpenDiv() Then
            Do
                p573 = .BufPos
                If Not (ParseHtmlBlockDiv()) Then
                    .BufPos = p573
                    p570 = .BufPos
                    e570 = .LastExpected
                    If ParseHtmlBlockCloseDiv() Then
                        .BufPos = p573
                        .BufPos = p573
                        Exit Do
                    Else
                        .BufPos = p570
                        .LastExpected = e570
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p573
                            .BufPos = p573
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseDiv() Then
                ParseHtmlBlockDiv = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenDl() As Boolean
    Dim p585 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 100 And .BufData(.BufPos + 1) = 108 Then ' "dl"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p585 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p585
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenDl = True
                End If
            End If
            If .BufData(.BufPos) = 68 And .BufData(.BufPos + 1) = 76 Then ' "DL"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p585 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p585
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenDl = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseDl() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 100 And .BufData(.BufPos + 1) = 108 Then ' "dl"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseDl = True
                    End If
                End If
                If .BufData(.BufPos) = 68 And .BufData(.BufPos + 1) = 76 Then ' "DL"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseDl = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockDl() As Boolean
    Dim p605 As Long
    Dim p602 As Long
    Dim e602 As String

    With ctx
        If ParseHtmlBlockOpenDl() Then
            Do
                p605 = .BufPos
                If Not (ParseHtmlBlockDl()) Then
                    .BufPos = p605
                    p602 = .BufPos
                    e602 = .LastExpected
                    If ParseHtmlBlockCloseDl() Then
                        .BufPos = p605
                        .BufPos = p605
                        Exit Do
                    Else
                        .BufPos = p602
                        .LastExpected = e602
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p605
                            .BufPos = p605
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseDl() Then
                ParseHtmlBlockDl = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenFieldset() As Boolean
    Dim p617 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("fieldset") Then       ' "fieldset"
                .BufPos = .BufPos + 8
                Call ParseSpnl
                Do
                    p617 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p617
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenFieldset = True
                End If
            End If
            If pvMatchString("FIELDSET") Then       ' "FIELDSET"
                .BufPos = .BufPos + 8
                Call ParseSpnl
                Do
                    p617 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p617
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenFieldset = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseFieldset() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("fieldset") Then   ' "fieldset"
                    .BufPos = .BufPos + 8
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseFieldset = True
                    End If
                End If
                If pvMatchString("FIELDSET") Then   ' "FIELDSET"
                    .BufPos = .BufPos + 8
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseFieldset = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockFieldset() As Boolean
    Dim p637 As Long
    Dim p634 As Long
    Dim e634 As String

    With ctx
        If ParseHtmlBlockOpenFieldset() Then
            Do
                p637 = .BufPos
                If Not (ParseHtmlBlockFieldset()) Then
                    .BufPos = p637
                    p634 = .BufPos
                    e634 = .LastExpected
                    If ParseHtmlBlockCloseFieldset() Then
                        .BufPos = p637
                        .BufPos = p637
                        Exit Do
                    Else
                        .BufPos = p634
                        .LastExpected = e634
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p637
                            .BufPos = p637
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseFieldset() Then
                ParseHtmlBlockFieldset = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenForm() As Boolean
    Dim p649 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("form") Then           ' "form"
                .BufPos = .BufPos + 4
                Call ParseSpnl
                Do
                    p649 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p649
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenForm = True
                End If
            End If
            If pvMatchString("FORM") Then           ' "FORM"
                .BufPos = .BufPos + 4
                Call ParseSpnl
                Do
                    p649 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p649
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenForm = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseForm() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("form") Then       ' "form"
                    .BufPos = .BufPos + 4
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseForm = True
                    End If
                End If
                If pvMatchString("FORM") Then       ' "FORM"
                    .BufPos = .BufPos + 4
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseForm = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockForm() As Boolean
    Dim p669 As Long
    Dim p666 As Long
    Dim e666 As String

    With ctx
        If ParseHtmlBlockOpenForm() Then
            Do
                p669 = .BufPos
                If Not (ParseHtmlBlockForm()) Then
                    .BufPos = p669
                    p666 = .BufPos
                    e666 = .LastExpected
                    If ParseHtmlBlockCloseForm() Then
                        .BufPos = p669
                        .BufPos = p669
                        Exit Do
                    Else
                        .BufPos = p666
                        .LastExpected = e666
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p669
                            .BufPos = p669
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseForm() Then
                ParseHtmlBlockForm = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenH1() As Boolean
    Dim p681 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 49 Then ' "h1"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p681 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p681
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenH1 = True
                End If
            End If
            If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 49 Then ' "H1"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p681 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p681
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenH1 = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseH1() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 49 Then ' "h1"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseH1 = True
                    End If
                End If
                If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 49 Then ' "H1"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseH1 = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockH1() As Boolean
    Dim p701 As Long
    Dim p698 As Long
    Dim e698 As String

    With ctx
        If ParseHtmlBlockOpenH1() Then
            Do
                p701 = .BufPos
                If Not (ParseHtmlBlockH1()) Then
                    .BufPos = p701
                    p698 = .BufPos
                    e698 = .LastExpected
                    If ParseHtmlBlockCloseH1() Then
                        .BufPos = p701
                        .BufPos = p701
                        Exit Do
                    Else
                        .BufPos = p698
                        .LastExpected = e698
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p701
                            .BufPos = p701
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseH1() Then
                ParseHtmlBlockH1 = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenH2() As Boolean
    Dim p713 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 50 Then ' "h2"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p713 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p713
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenH2 = True
                End If
            End If
            If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 50 Then ' "H2"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p713 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p713
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenH2 = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseH2() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 50 Then ' "h2"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseH2 = True
                    End If
                End If
                If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 50 Then ' "H2"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseH2 = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockH2() As Boolean
    Dim p733 As Long
    Dim p730 As Long
    Dim e730 As String

    With ctx
        If ParseHtmlBlockOpenH2() Then
            Do
                p733 = .BufPos
                If Not (ParseHtmlBlockH2()) Then
                    .BufPos = p733
                    p730 = .BufPos
                    e730 = .LastExpected
                    If ParseHtmlBlockCloseH2() Then
                        .BufPos = p733
                        .BufPos = p733
                        Exit Do
                    Else
                        .BufPos = p730
                        .LastExpected = e730
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p733
                            .BufPos = p733
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseH2() Then
                ParseHtmlBlockH2 = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenH3() As Boolean
    Dim p745 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 51 Then ' "h3"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p745 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p745
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenH3 = True
                End If
            End If
            If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 51 Then ' "H3"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p745 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p745
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenH3 = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseH3() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 51 Then ' "h3"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseH3 = True
                    End If
                End If
                If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 51 Then ' "H3"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseH3 = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockH3() As Boolean
    Dim p765 As Long
    Dim p762 As Long
    Dim e762 As String

    With ctx
        If ParseHtmlBlockOpenH3() Then
            Do
                p765 = .BufPos
                If Not (ParseHtmlBlockH3()) Then
                    .BufPos = p765
                    p762 = .BufPos
                    e762 = .LastExpected
                    If ParseHtmlBlockCloseH3() Then
                        .BufPos = p765
                        .BufPos = p765
                        Exit Do
                    Else
                        .BufPos = p762
                        .LastExpected = e762
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p765
                            .BufPos = p765
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseH3() Then
                ParseHtmlBlockH3 = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenH4() As Boolean
    Dim p777 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 52 Then ' "h4"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p777 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p777
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenH4 = True
                End If
            End If
            If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 52 Then ' "H4"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p777 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p777
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenH4 = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseH4() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 52 Then ' "h4"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseH4 = True
                    End If
                End If
                If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 52 Then ' "H4"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseH4 = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockH4() As Boolean
    Dim p797 As Long
    Dim p794 As Long
    Dim e794 As String

    With ctx
        If ParseHtmlBlockOpenH4() Then
            Do
                p797 = .BufPos
                If Not (ParseHtmlBlockH4()) Then
                    .BufPos = p797
                    p794 = .BufPos
                    e794 = .LastExpected
                    If ParseHtmlBlockCloseH4() Then
                        .BufPos = p797
                        .BufPos = p797
                        Exit Do
                    Else
                        .BufPos = p794
                        .LastExpected = e794
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p797
                            .BufPos = p797
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseH4() Then
                ParseHtmlBlockH4 = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenH5() As Boolean
    Dim p809 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 53 Then ' "h5"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p809 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p809
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenH5 = True
                End If
            End If
            If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 53 Then ' "H5"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p809 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p809
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenH5 = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseH5() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 53 Then ' "h5"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseH5 = True
                    End If
                End If
                If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 53 Then ' "H5"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseH5 = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockH5() As Boolean
    Dim p829 As Long
    Dim p826 As Long
    Dim e826 As String

    With ctx
        If ParseHtmlBlockOpenH5() Then
            Do
                p829 = .BufPos
                If Not (ParseHtmlBlockH5()) Then
                    .BufPos = p829
                    p826 = .BufPos
                    e826 = .LastExpected
                    If ParseHtmlBlockCloseH5() Then
                        .BufPos = p829
                        .BufPos = p829
                        Exit Do
                    Else
                        .BufPos = p826
                        .LastExpected = e826
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p829
                            .BufPos = p829
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseH5() Then
                ParseHtmlBlockH5 = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenH6() As Boolean
    Dim p841 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 54 Then ' "h6"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p841 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p841
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenH6 = True
                End If
            End If
            If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 54 Then ' "H6"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p841 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p841
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenH6 = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseH6() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 54 Then ' "h6"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseH6 = True
                    End If
                End If
                If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 54 Then ' "H6"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseH6 = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockH6() As Boolean
    Dim p861 As Long
    Dim p858 As Long
    Dim e858 As String

    With ctx
        If ParseHtmlBlockOpenH6() Then
            Do
                p861 = .BufPos
                If Not (ParseHtmlBlockH6()) Then
                    .BufPos = p861
                    p858 = .BufPos
                    e858 = .LastExpected
                    If ParseHtmlBlockCloseH6() Then
                        .BufPos = p861
                        .BufPos = p861
                        Exit Do
                    Else
                        .BufPos = p858
                        .LastExpected = e858
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p861
                            .BufPos = p861
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseH6() Then
                ParseHtmlBlockH6 = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenMenu() As Boolean
    Dim p873 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("menu") Then           ' "menu"
                .BufPos = .BufPos + 4
                Call ParseSpnl
                Do
                    p873 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p873
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenMenu = True
                End If
            End If
            If pvMatchString("MENU") Then           ' "MENU"
                .BufPos = .BufPos + 4
                Call ParseSpnl
                Do
                    p873 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p873
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenMenu = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseMenu() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("menu") Then       ' "menu"
                    .BufPos = .BufPos + 4
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseMenu = True
                    End If
                End If
                If pvMatchString("MENU") Then       ' "MENU"
                    .BufPos = .BufPos + 4
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseMenu = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockMenu() As Boolean
    Dim p893 As Long
    Dim p890 As Long
    Dim e890 As String

    With ctx
        If ParseHtmlBlockOpenMenu() Then
            Do
                p893 = .BufPos
                If Not (ParseHtmlBlockMenu()) Then
                    .BufPos = p893
                    p890 = .BufPos
                    e890 = .LastExpected
                    If ParseHtmlBlockCloseMenu() Then
                        .BufPos = p893
                        .BufPos = p893
                        Exit Do
                    Else
                        .BufPos = p890
                        .LastExpected = e890
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p893
                            .BufPos = p893
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseMenu() Then
                ParseHtmlBlockMenu = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenNoframes() As Boolean
    Dim p905 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("noframes") Then       ' "noframes"
                .BufPos = .BufPos + 8
                Call ParseSpnl
                Do
                    p905 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p905
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenNoframes = True
                End If
            End If
            If pvMatchString("NOFRAMES") Then       ' "NOFRAMES"
                .BufPos = .BufPos + 8
                Call ParseSpnl
                Do
                    p905 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p905
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenNoframes = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseNoframes() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("noframes") Then   ' "noframes"
                    .BufPos = .BufPos + 8
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseNoframes = True
                    End If
                End If
                If pvMatchString("NOFRAMES") Then   ' "NOFRAMES"
                    .BufPos = .BufPos + 8
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseNoframes = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockNoframes() As Boolean
    Dim p925 As Long
    Dim p922 As Long
    Dim e922 As String

    With ctx
        If ParseHtmlBlockOpenNoframes() Then
            Do
                p925 = .BufPos
                If Not (ParseHtmlBlockNoframes()) Then
                    .BufPos = p925
                    p922 = .BufPos
                    e922 = .LastExpected
                    If ParseHtmlBlockCloseNoframes() Then
                        .BufPos = p925
                        .BufPos = p925
                        Exit Do
                    Else
                        .BufPos = p922
                        .LastExpected = e922
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p925
                            .BufPos = p925
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseNoframes() Then
                ParseHtmlBlockNoframes = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenNoscript() As Boolean
    Dim p937 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("noscript") Then       ' "noscript"
                .BufPos = .BufPos + 8
                Call ParseSpnl
                Do
                    p937 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p937
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenNoscript = True
                End If
            End If
            If pvMatchString("NOSCRIPT") Then       ' "NOSCRIPT"
                .BufPos = .BufPos + 8
                Call ParseSpnl
                Do
                    p937 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p937
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenNoscript = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseNoscript() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("noscript") Then   ' "noscript"
                    .BufPos = .BufPos + 8
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseNoscript = True
                    End If
                End If
                If pvMatchString("NOSCRIPT") Then   ' "NOSCRIPT"
                    .BufPos = .BufPos + 8
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseNoscript = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockNoscript() As Boolean
    Dim p957 As Long
    Dim p954 As Long
    Dim e954 As String

    With ctx
        If ParseHtmlBlockOpenNoscript() Then
            Do
                p957 = .BufPos
                If Not (ParseHtmlBlockNoscript()) Then
                    .BufPos = p957
                    p954 = .BufPos
                    e954 = .LastExpected
                    If ParseHtmlBlockCloseNoscript() Then
                        .BufPos = p957
                        .BufPos = p957
                        Exit Do
                    Else
                        .BufPos = p954
                        .LastExpected = e954
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p957
                            .BufPos = p957
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseNoscript() Then
                ParseHtmlBlockNoscript = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenOl() As Boolean
    Dim p969 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 111 And .BufData(.BufPos + 1) = 108 Then ' "ol"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p969 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p969
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenOl = True
                End If
            End If
            If .BufData(.BufPos) = 79 And .BufData(.BufPos + 1) = 76 Then ' "OL"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p969 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p969
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenOl = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseOl() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 111 And .BufData(.BufPos + 1) = 108 Then ' "ol"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseOl = True
                    End If
                End If
                If .BufData(.BufPos) = 79 And .BufData(.BufPos + 1) = 76 Then ' "OL"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseOl = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOl() As Boolean
    Dim p989 As Long
    Dim p986 As Long
    Dim e986 As String

    With ctx
        If ParseHtmlBlockOpenOl() Then
            Do
                p989 = .BufPos
                If Not (ParseHtmlBlockOl()) Then
                    .BufPos = p989
                    p986 = .BufPos
                    e986 = .LastExpected
                    If ParseHtmlBlockCloseOl() Then
                        .BufPos = p989
                        .BufPos = p989
                        Exit Do
                    Else
                        .BufPos = p986
                        .LastExpected = e986
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p989
                            .BufPos = p989
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseOl() Then
                ParseHtmlBlockOl = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenP() As Boolean
    Dim p1001 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 112 Then         ' "p"
                .BufPos = .BufPos + 1
                Call ParseSpnl
                Do
                    p1001 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1001
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenP = True
                End If
            End If
            If .BufData(.BufPos) = 80 Then          ' "P"
                .BufPos = .BufPos + 1
                Call ParseSpnl
                Do
                    p1001 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1001
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenP = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseP() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 112 Then     ' "p"
                    .BufPos = .BufPos + 1
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseP = True
                    End If
                End If
                If .BufData(.BufPos) = 80 Then      ' "P"
                    .BufPos = .BufPos + 1
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseP = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockP() As Boolean
    Dim p1021 As Long
    Dim p1018 As Long
    Dim e1018 As String

    With ctx
        If ParseHtmlBlockOpenP() Then
            Do
                p1021 = .BufPos
                If Not (ParseHtmlBlockP()) Then
                    .BufPos = p1021
                    p1018 = .BufPos
                    e1018 = .LastExpected
                    If ParseHtmlBlockCloseP() Then
                        .BufPos = p1021
                        .BufPos = p1021
                        Exit Do
                    Else
                        .BufPos = p1018
                        .LastExpected = e1018
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p1021
                            .BufPos = p1021
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseP() Then
                ParseHtmlBlockP = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenPre() As Boolean
    Dim p1033 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("pre") Then            ' "pre"
                .BufPos = .BufPos + 3
                Call ParseSpnl
                Do
                    p1033 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1033
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenPre = True
                End If
            End If
            If pvMatchString("PRE") Then            ' "PRE"
                .BufPos = .BufPos + 3
                Call ParseSpnl
                Do
                    p1033 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1033
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenPre = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockClosePre() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("pre") Then        ' "pre"
                    .BufPos = .BufPos + 3
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockClosePre = True
                    End If
                End If
                If pvMatchString("PRE") Then        ' "PRE"
                    .BufPos = .BufPos + 3
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockClosePre = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockPre() As Boolean
    Dim p1053 As Long
    Dim p1050 As Long
    Dim e1050 As String

    With ctx
        If ParseHtmlBlockOpenPre() Then
            Do
                p1053 = .BufPos
                If Not (ParseHtmlBlockPre()) Then
                    .BufPos = p1053
                    p1050 = .BufPos
                    e1050 = .LastExpected
                    If ParseHtmlBlockClosePre() Then
                        .BufPos = p1053
                        .BufPos = p1053
                        Exit Do
                    Else
                        .BufPos = p1050
                        .LastExpected = e1050
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p1053
                            .BufPos = p1053
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockClosePre() Then
                ParseHtmlBlockPre = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenTable() As Boolean
    Dim p1065 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("table") Then          ' "table"
                .BufPos = .BufPos + 5
                Call ParseSpnl
                Do
                    p1065 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1065
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenTable = True
                End If
            End If
            If pvMatchString("TABLE") Then          ' "TABLE"
                .BufPos = .BufPos + 5
                Call ParseSpnl
                Do
                    p1065 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1065
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenTable = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseTable() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("table") Then      ' "table"
                    .BufPos = .BufPos + 5
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseTable = True
                    End If
                End If
                If pvMatchString("TABLE") Then      ' "TABLE"
                    .BufPos = .BufPos + 5
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseTable = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockTable() As Boolean
    Dim p1085 As Long
    Dim p1082 As Long
    Dim e1082 As String

    With ctx
        If ParseHtmlBlockOpenTable() Then
            Do
                p1085 = .BufPos
                If Not (ParseHtmlBlockTable()) Then
                    .BufPos = p1085
                    p1082 = .BufPos
                    e1082 = .LastExpected
                    If ParseHtmlBlockCloseTable() Then
                        .BufPos = p1085
                        .BufPos = p1085
                        Exit Do
                    Else
                        .BufPos = p1082
                        .LastExpected = e1082
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p1085
                            .BufPos = p1085
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseTable() Then
                ParseHtmlBlockTable = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenUl() As Boolean
    Dim p1097 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 117 And .BufData(.BufPos + 1) = 108 Then ' "ul"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p1097 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1097
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenUl = True
                End If
            End If
            If .BufData(.BufPos) = 85 And .BufData(.BufPos + 1) = 76 Then ' "UL"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p1097 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1097
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenUl = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseUl() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 117 And .BufData(.BufPos + 1) = 108 Then ' "ul"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseUl = True
                    End If
                End If
                If .BufData(.BufPos) = 85 And .BufData(.BufPos + 1) = 76 Then ' "UL"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseUl = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockUl() As Boolean
    Dim p1117 As Long
    Dim p1114 As Long
    Dim e1114 As String

    With ctx
        If ParseHtmlBlockOpenUl() Then
            Do
                p1117 = .BufPos
                If Not (ParseHtmlBlockUl()) Then
                    .BufPos = p1117
                    p1114 = .BufPos
                    e1114 = .LastExpected
                    If ParseHtmlBlockCloseUl() Then
                        .BufPos = p1117
                        .BufPos = p1117
                        Exit Do
                    Else
                        .BufPos = p1114
                        .LastExpected = e1114
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p1117
                            .BufPos = p1117
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseUl() Then
                ParseHtmlBlockUl = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenDd() As Boolean
    Dim p1129 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 100 And .BufData(.BufPos + 1) = 100 Then ' "dd"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p1129 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1129
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenDd = True
                End If
            End If
            If .BufData(.BufPos) = 68 And .BufData(.BufPos + 1) = 68 Then ' "DD"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p1129 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1129
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenDd = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseDd() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 100 And .BufData(.BufPos + 1) = 100 Then ' "dd"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseDd = True
                    End If
                End If
                If .BufData(.BufPos) = 68 And .BufData(.BufPos + 1) = 68 Then ' "DD"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseDd = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockDd() As Boolean
    Dim p1149 As Long
    Dim p1146 As Long
    Dim e1146 As String

    With ctx
        If ParseHtmlBlockOpenDd() Then
            Do
                p1149 = .BufPos
                If Not (ParseHtmlBlockDd()) Then
                    .BufPos = p1149
                    p1146 = .BufPos
                    e1146 = .LastExpected
                    If ParseHtmlBlockCloseDd() Then
                        .BufPos = p1149
                        .BufPos = p1149
                        Exit Do
                    Else
                        .BufPos = p1146
                        .LastExpected = e1146
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p1149
                            .BufPos = p1149
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseDd() Then
                ParseHtmlBlockDd = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenDt() As Boolean
    Dim p1161 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 100 And .BufData(.BufPos + 1) = 116 Then ' "dt"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p1161 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1161
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenDt = True
                End If
            End If
            If .BufData(.BufPos) = 68 And .BufData(.BufPos + 1) = 84 Then ' "DT"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p1161 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1161
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenDt = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseDt() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 100 And .BufData(.BufPos + 1) = 116 Then ' "dt"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseDt = True
                    End If
                End If
                If .BufData(.BufPos) = 68 And .BufData(.BufPos + 1) = 84 Then ' "DT"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseDt = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockDt() As Boolean
    Dim p1181 As Long
    Dim p1178 As Long
    Dim e1178 As String

    With ctx
        If ParseHtmlBlockOpenDt() Then
            Do
                p1181 = .BufPos
                If Not (ParseHtmlBlockDt()) Then
                    .BufPos = p1181
                    p1178 = .BufPos
                    e1178 = .LastExpected
                    If ParseHtmlBlockCloseDt() Then
                        .BufPos = p1181
                        .BufPos = p1181
                        Exit Do
                    Else
                        .BufPos = p1178
                        .LastExpected = e1178
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p1181
                            .BufPos = p1181
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseDt() Then
                ParseHtmlBlockDt = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenFrameset() As Boolean
    Dim p1193 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("frameset") Then       ' "frameset"
                .BufPos = .BufPos + 8
                Call ParseSpnl
                Do
                    p1193 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1193
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenFrameset = True
                End If
            End If
            If pvMatchString("FRAMESET") Then       ' "FRAMESET"
                .BufPos = .BufPos + 8
                Call ParseSpnl
                Do
                    p1193 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1193
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenFrameset = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseFrameset() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("frameset") Then   ' "frameset"
                    .BufPos = .BufPos + 8
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseFrameset = True
                    End If
                End If
                If pvMatchString("FRAMESET") Then   ' "FRAMESET"
                    .BufPos = .BufPos + 8
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseFrameset = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockFrameset() As Boolean
    Dim p1213 As Long
    Dim p1210 As Long
    Dim e1210 As String

    With ctx
        If ParseHtmlBlockOpenFrameset() Then
            Do
                p1213 = .BufPos
                If Not (ParseHtmlBlockFrameset()) Then
                    .BufPos = p1213
                    p1210 = .BufPos
                    e1210 = .LastExpected
                    If ParseHtmlBlockCloseFrameset() Then
                        .BufPos = p1213
                        .BufPos = p1213
                        Exit Do
                    Else
                        .BufPos = p1210
                        .LastExpected = e1210
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p1213
                            .BufPos = p1213
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseFrameset() Then
                ParseHtmlBlockFrameset = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenLi() As Boolean
    Dim p1225 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 108 And .BufData(.BufPos + 1) = 105 Then ' "li"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p1225 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1225
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenLi = True
                End If
            End If
            If .BufData(.BufPos) = 76 And .BufData(.BufPos + 1) = 73 Then ' "LI"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p1225 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1225
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenLi = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseLi() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 108 And .BufData(.BufPos + 1) = 105 Then ' "li"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseLi = True
                    End If
                End If
                If .BufData(.BufPos) = 76 And .BufData(.BufPos + 1) = 73 Then ' "LI"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseLi = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockLi() As Boolean
    Dim p1245 As Long
    Dim p1242 As Long
    Dim e1242 As String

    With ctx
        If ParseHtmlBlockOpenLi() Then
            Do
                p1245 = .BufPos
                If Not (ParseHtmlBlockLi()) Then
                    .BufPos = p1245
                    p1242 = .BufPos
                    e1242 = .LastExpected
                    If ParseHtmlBlockCloseLi() Then
                        .BufPos = p1245
                        .BufPos = p1245
                        Exit Do
                    Else
                        .BufPos = p1242
                        .LastExpected = e1242
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p1245
                            .BufPos = p1245
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseLi() Then
                ParseHtmlBlockLi = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenTbody() As Boolean
    Dim p1257 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("tbody") Then          ' "tbody"
                .BufPos = .BufPos + 5
                Call ParseSpnl
                Do
                    p1257 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1257
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenTbody = True
                End If
            End If
            If pvMatchString("TBODY") Then          ' "TBODY"
                .BufPos = .BufPos + 5
                Call ParseSpnl
                Do
                    p1257 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1257
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenTbody = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseTbody() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("tbody") Then      ' "tbody"
                    .BufPos = .BufPos + 5
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseTbody = True
                    End If
                End If
                If pvMatchString("TBODY") Then      ' "TBODY"
                    .BufPos = .BufPos + 5
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseTbody = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockTbody() As Boolean
    Dim p1277 As Long
    Dim p1274 As Long
    Dim e1274 As String

    With ctx
        If ParseHtmlBlockOpenTbody() Then
            Do
                p1277 = .BufPos
                If Not (ParseHtmlBlockTbody()) Then
                    .BufPos = p1277
                    p1274 = .BufPos
                    e1274 = .LastExpected
                    If ParseHtmlBlockCloseTbody() Then
                        .BufPos = p1277
                        .BufPos = p1277
                        Exit Do
                    Else
                        .BufPos = p1274
                        .LastExpected = e1274
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p1277
                            .BufPos = p1277
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseTbody() Then
                ParseHtmlBlockTbody = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenTd() As Boolean
    Dim p1289 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 116 And .BufData(.BufPos + 1) = 100 Then ' "td"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p1289 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1289
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenTd = True
                End If
            End If
            If .BufData(.BufPos) = 84 And .BufData(.BufPos + 1) = 68 Then ' "TD"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p1289 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1289
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenTd = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseTd() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 116 And .BufData(.BufPos + 1) = 100 Then ' "td"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseTd = True
                    End If
                End If
                If .BufData(.BufPos) = 84 And .BufData(.BufPos + 1) = 68 Then ' "TD"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseTd = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockTd() As Boolean
    Dim p1309 As Long
    Dim p1306 As Long
    Dim e1306 As String

    With ctx
        If ParseHtmlBlockOpenTd() Then
            Do
                p1309 = .BufPos
                If Not (ParseHtmlBlockTd()) Then
                    .BufPos = p1309
                    p1306 = .BufPos
                    e1306 = .LastExpected
                    If ParseHtmlBlockCloseTd() Then
                        .BufPos = p1309
                        .BufPos = p1309
                        Exit Do
                    Else
                        .BufPos = p1306
                        .LastExpected = e1306
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p1309
                            .BufPos = p1309
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseTd() Then
                ParseHtmlBlockTd = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenTfoot() As Boolean
    Dim p1321 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("tfoot") Then          ' "tfoot"
                .BufPos = .BufPos + 5
                Call ParseSpnl
                Do
                    p1321 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1321
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenTfoot = True
                End If
            End If
            If pvMatchString("TFOOT") Then          ' "TFOOT"
                .BufPos = .BufPos + 5
                Call ParseSpnl
                Do
                    p1321 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1321
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenTfoot = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseTfoot() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("tfoot") Then      ' "tfoot"
                    .BufPos = .BufPos + 5
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseTfoot = True
                    End If
                End If
                If pvMatchString("TFOOT") Then      ' "TFOOT"
                    .BufPos = .BufPos + 5
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseTfoot = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockTfoot() As Boolean
    Dim p1341 As Long
    Dim p1338 As Long
    Dim e1338 As String

    With ctx
        If ParseHtmlBlockOpenTfoot() Then
            Do
                p1341 = .BufPos
                If Not (ParseHtmlBlockTfoot()) Then
                    .BufPos = p1341
                    p1338 = .BufPos
                    e1338 = .LastExpected
                    If ParseHtmlBlockCloseTfoot() Then
                        .BufPos = p1341
                        .BufPos = p1341
                        Exit Do
                    Else
                        .BufPos = p1338
                        .LastExpected = e1338
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p1341
                            .BufPos = p1341
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseTfoot() Then
                ParseHtmlBlockTfoot = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenTh() As Boolean
    Dim p1353 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 116 And .BufData(.BufPos + 1) = 104 Then ' "th"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p1353 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1353
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenTh = True
                End If
            End If
            If .BufData(.BufPos) = 84 And .BufData(.BufPos + 1) = 72 Then ' "TH"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p1353 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1353
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenTh = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseTh() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 116 And .BufData(.BufPos + 1) = 104 Then ' "th"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseTh = True
                    End If
                End If
                If .BufData(.BufPos) = 84 And .BufData(.BufPos + 1) = 72 Then ' "TH"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseTh = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockTh() As Boolean
    Dim p1373 As Long
    Dim p1370 As Long
    Dim e1370 As String

    With ctx
        If ParseHtmlBlockOpenTh() Then
            Do
                p1373 = .BufPos
                If Not (ParseHtmlBlockTh()) Then
                    .BufPos = p1373
                    p1370 = .BufPos
                    e1370 = .LastExpected
                    If ParseHtmlBlockCloseTh() Then
                        .BufPos = p1373
                        .BufPos = p1373
                        Exit Do
                    Else
                        .BufPos = p1370
                        .LastExpected = e1370
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p1373
                            .BufPos = p1373
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseTh() Then
                ParseHtmlBlockTh = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenThead() As Boolean
    Dim p1385 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("thead") Then          ' "thead"
                .BufPos = .BufPos + 5
                Call ParseSpnl
                Do
                    p1385 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1385
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenThead = True
                End If
            End If
            If pvMatchString("THEAD") Then          ' "THEAD"
                .BufPos = .BufPos + 5
                Call ParseSpnl
                Do
                    p1385 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1385
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenThead = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseThead() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("thead") Then      ' "thead"
                    .BufPos = .BufPos + 5
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseThead = True
                    End If
                End If
                If pvMatchString("THEAD") Then      ' "THEAD"
                    .BufPos = .BufPos + 5
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseThead = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockThead() As Boolean
    Dim p1405 As Long
    Dim p1402 As Long
    Dim e1402 As String

    With ctx
        If ParseHtmlBlockOpenThead() Then
            Do
                p1405 = .BufPos
                If Not (ParseHtmlBlockThead()) Then
                    .BufPos = p1405
                    p1402 = .BufPos
                    e1402 = .LastExpected
                    If ParseHtmlBlockCloseThead() Then
                        .BufPos = p1405
                        .BufPos = p1405
                        Exit Do
                    Else
                        .BufPos = p1402
                        .LastExpected = e1402
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p1405
                            .BufPos = p1405
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseThead() Then
                ParseHtmlBlockThead = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenTr() As Boolean
    Dim p1417 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 116 And .BufData(.BufPos + 1) = 114 Then ' "tr"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p1417 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1417
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenTr = True
                End If
            End If
            If .BufData(.BufPos) = 84 And .BufData(.BufPos + 1) = 82 Then ' "TR"
                .BufPos = .BufPos + 2
                Call ParseSpnl
                Do
                    p1417 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1417
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenTr = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseTr() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If .BufData(.BufPos) = 116 And .BufData(.BufPos + 1) = 114 Then ' "tr"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseTr = True
                    End If
                End If
                If .BufData(.BufPos) = 84 And .BufData(.BufPos + 1) = 82 Then ' "TR"
                    .BufPos = .BufPos + 2
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseTr = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockTr() As Boolean
    Dim p1437 As Long
    Dim p1434 As Long
    Dim e1434 As String

    With ctx
        If ParseHtmlBlockOpenTr() Then
            Do
                p1437 = .BufPos
                If Not (ParseHtmlBlockTr()) Then
                    .BufPos = p1437
                    p1434 = .BufPos
                    e1434 = .LastExpected
                    If ParseHtmlBlockCloseTr() Then
                        .BufPos = p1437
                        .BufPos = p1437
                        Exit Do
                    Else
                        .BufPos = p1434
                        .LastExpected = e1434
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p1437
                            .BufPos = p1437
                            Exit Do
                        End If
                    End If
                End If
            Loop
            If ParseHtmlBlockCloseTr() Then
                ParseHtmlBlockTr = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenScript() As Boolean
    Dim p1449 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("script") Then         ' "script"
                .BufPos = .BufPos + 6
                Call ParseSpnl
                Do
                    p1449 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1449
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenScript = True
                End If
            End If
            If pvMatchString("SCRIPT") Then         ' "SCRIPT"
                .BufPos = .BufPos + 6
                Call ParseSpnl
                Do
                    p1449 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1449
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenScript = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseScript() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("script") Then     ' "script"
                    .BufPos = .BufPos + 6
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseScript = True
                    End If
                End If
                If pvMatchString("SCRIPT") Then     ' "SCRIPT"
                    .BufPos = .BufPos + 6
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseScript = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockScript() As Boolean
    Dim p1467 As Long
    Dim p1465 As Long
    Dim e1465 As String

    With ctx
        If ParseHtmlBlockOpenScript() Then
            Do
                p1467 = .BufPos
                p1465 = .BufPos
                e1465 = .LastExpected
                If ParseHtmlBlockCloseScript() Then
                    .BufPos = p1467
                    Exit Do
                Else
                    .BufPos = p1465
                    .LastExpected = e1465
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p1467
                    Exit Do
                End If
            Loop
            If ParseHtmlBlockCloseScript() Then
                ParseHtmlBlockScript = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockOpenHead() As Boolean
    Dim p1479 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("head") Then           ' "head"
                .BufPos = .BufPos + 4
                Call ParseSpnl
                Do
                    p1479 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1479
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenHead = True
                End If
            End If
            If pvMatchString("HEAD") Then           ' "HEAD"
                .BufPos = .BufPos + 4
                Call ParseSpnl
                Do
                    p1479 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1479
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlBlockOpenHead = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockCloseHead() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("head") Then       ' "head"
                    .BufPos = .BufPos + 4
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseHead = True
                    End If
                End If
                If pvMatchString("HEAD") Then       ' "HEAD"
                    .BufPos = .BufPos + 4
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockCloseHead = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockHead() As Boolean
    Dim p1497 As Long
    Dim p1495 As Long
    Dim e1495 As String

    With ctx
        If ParseHtmlBlockOpenHead() Then
            Do
                p1497 = .BufPos
                p1495 = .BufPos
                e1495 = .LastExpected
                If ParseHtmlBlockCloseHead() Then
                    .BufPos = p1497
                    Exit Do
                Else
                    .BufPos = p1495
                    .LastExpected = e1495
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p1497
                    Exit Do
                End If
            Loop
            If ParseHtmlBlockCloseHead() Then
                ParseHtmlBlockHead = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockInTags() As Boolean
    Dim p1504 As Long

    With ctx
        p1504 = .BufPos
        If ParseHtmlBlockAddress() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockBlockquote() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockCenter() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockDir() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockDiv() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockDl() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockFieldset() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockForm() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockH1() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockH2() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockH3() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockH4() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockH5() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockH6() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockMenu() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockNoframes() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockNoscript() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockOl() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockP() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockPre() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockTable() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockUl() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockDd() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockDt() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockFrameset() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockLi() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockTbody() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockTd() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockTfoot() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockTh() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockThead() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockTr() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockScript() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
        If ParseHtmlBlockHead() Then
            ParseHtmlBlockInTags = True
            Exit Function
        Else
            .BufPos = p1504
        End If
    End With
End Function

Private Function ParseHtmlComment() As Boolean
    Dim p2516 As Long
    Dim p2514 As Long

    With ctx
        If pvMatchString("<!--") Then               ' "<!--"
            .BufPos = .BufPos + 4
            Do
                p2516 = .BufPos
                p2514 = .BufPos
                If pvMatchString("-->") Then        ' "-->"
                    .BufPos = p2516
                    Exit Do
                Else
                    .BufPos = p2514
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p2516
                    Exit Do
                End If
            Loop
            If pvMatchString("-->") Then            ' "-->"
                .BufPos = .BufPos + 3
                Call pvSetAdvance
                ParseHtmlComment = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockSelfClosing() As Boolean
    Dim p1555 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If ParseHtmlBlockType() Then
                Call ParseSpnl
                Do
                    p1555 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1555
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 47 Then      ' "/"
                    .BufPos = .BufPos + 1
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseHtmlBlockSelfClosing = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseHtmlBlockType() As Boolean
    With ctx
        If pvMatchString("address") Then            ' "address"
            .BufPos = .BufPos + 7
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("blockquote") Then         ' "blockquote"
            .BufPos = .BufPos + 10
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("center") Then             ' "center"
            .BufPos = .BufPos + 6
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("dir") Then                ' "dir"
            .BufPos = .BufPos + 3
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("div") Then                ' "div"
            .BufPos = .BufPos + 3
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 100 And .BufData(.BufPos + 1) = 108 Then ' "dl"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("fieldset") Then           ' "fieldset"
            .BufPos = .BufPos + 8
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("form") Then               ' "form"
            .BufPos = .BufPos + 4
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 49 Then ' "h1"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 50 Then ' "h2"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 51 Then ' "h3"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 52 Then ' "h4"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 53 Then ' "h5"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 54 Then ' "h6"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 104 And .BufData(.BufPos + 1) = 114 Then ' "hr"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("isindex") Then            ' "isindex"
            .BufPos = .BufPos + 7
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("menu") Then               ' "menu"
            .BufPos = .BufPos + 4
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("noframes") Then           ' "noframes"
            .BufPos = .BufPos + 8
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("noscript") Then           ' "noscript"
            .BufPos = .BufPos + 8
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 111 And .BufData(.BufPos + 1) = 108 Then ' "ol"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 112 Then             ' "p"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("pre") Then                ' "pre"
            .BufPos = .BufPos + 3
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("table") Then              ' "table"
            .BufPos = .BufPos + 5
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 117 And .BufData(.BufPos + 1) = 108 Then ' "ul"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 100 And .BufData(.BufPos + 1) = 100 Then ' "dd"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 100 And .BufData(.BufPos + 1) = 116 Then ' "dt"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("frameset") Then           ' "frameset"
            .BufPos = .BufPos + 8
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 108 And .BufData(.BufPos + 1) = 105 Then ' "li"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("tbody") Then              ' "tbody"
            .BufPos = .BufPos + 5
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 116 And .BufData(.BufPos + 1) = 100 Then ' "td"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("tfoot") Then              ' "tfoot"
            .BufPos = .BufPos + 5
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 116 And .BufData(.BufPos + 1) = 104 Then ' "th"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("thead") Then              ' "thead"
            .BufPos = .BufPos + 5
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 116 And .BufData(.BufPos + 1) = 114 Then ' "tr"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("script") Then             ' "script"
            .BufPos = .BufPos + 6
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("ADDRESS") Then            ' "ADDRESS"
            .BufPos = .BufPos + 7
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("BLOCKQUOTE") Then         ' "BLOCKQUOTE"
            .BufPos = .BufPos + 10
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("CENTER") Then             ' "CENTER"
            .BufPos = .BufPos + 6
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("DIR") Then                ' "DIR"
            .BufPos = .BufPos + 3
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("DIV") Then                ' "DIV"
            .BufPos = .BufPos + 3
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 68 And .BufData(.BufPos + 1) = 76 Then ' "DL"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("FIELDSET") Then           ' "FIELDSET"
            .BufPos = .BufPos + 8
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("FORM") Then               ' "FORM"
            .BufPos = .BufPos + 4
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 49 Then ' "H1"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 50 Then ' "H2"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 51 Then ' "H3"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 52 Then ' "H4"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 53 Then ' "H5"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 54 Then ' "H6"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 72 And .BufData(.BufPos + 1) = 82 Then ' "HR"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("ISINDEX") Then            ' "ISINDEX"
            .BufPos = .BufPos + 7
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("MENU") Then               ' "MENU"
            .BufPos = .BufPos + 4
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("NOFRAMES") Then           ' "NOFRAMES"
            .BufPos = .BufPos + 8
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("NOSCRIPT") Then           ' "NOSCRIPT"
            .BufPos = .BufPos + 8
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 79 And .BufData(.BufPos + 1) = 76 Then ' "OL"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 80 Then              ' "P"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("PRE") Then                ' "PRE"
            .BufPos = .BufPos + 3
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("TABLE") Then              ' "TABLE"
            .BufPos = .BufPos + 5
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 85 And .BufData(.BufPos + 1) = 76 Then ' "UL"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 68 And .BufData(.BufPos + 1) = 68 Then ' "DD"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 68 And .BufData(.BufPos + 1) = 84 Then ' "DT"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("FRAMESET") Then           ' "FRAMESET"
            .BufPos = .BufPos + 8
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 76 And .BufData(.BufPos + 1) = 73 Then ' "LI"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("TBODY") Then              ' "TBODY"
            .BufPos = .BufPos + 5
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 84 And .BufData(.BufPos + 1) = 68 Then ' "TD"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("TFOOT") Then              ' "TFOOT"
            .BufPos = .BufPos + 5
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 84 And .BufData(.BufPos + 1) = 72 Then ' "TH"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("THEAD") Then              ' "THEAD"
            .BufPos = .BufPos + 5
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If .BufData(.BufPos) = 84 And .BufData(.BufPos + 1) = 82 Then ' "TR"
            .BufPos = .BufPos + 2
            Call pvSetAdvance
            ParseHtmlBlockType = True
            Exit Function
        End If
        If pvMatchString("SCRIPT") Then             ' "SCRIPT"
            .BufPos = .BufPos + 6
            Call pvSetAdvance
            ParseHtmlBlockType = True
        End If
    End With
End Function

Private Function ParseStyleOpen() As Boolean
    Dim p1639 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If pvMatchString("style") Then          ' "style"
                .BufPos = .BufPos + 5
                Call ParseSpnl
                Do
                    p1639 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1639
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseStyleOpen = True
                End If
            End If
            If pvMatchString("STYLE") Then          ' "STYLE"
                .BufPos = .BufPos + 5
                Call ParseSpnl
                Do
                    p1639 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p1639
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseStyleOpen = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseStyleClose() As Boolean
    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
                If pvMatchString("style") Then      ' "style"
                    .BufPos = .BufPos + 5
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseStyleClose = True
                    End If
                End If
                If pvMatchString("STYLE") Then      ' "STYLE"
                    .BufPos = .BufPos + 5
                    Call ParseSpnl
                    If .BufData(.BufPos) = 62 Then  ' ">"
                        .BufPos = .BufPos + 1
                        Call pvSetAdvance
                        ParseStyleClose = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseInStyleTags() As Boolean
    Dim p1657 As Long
    Dim p1655 As Long
    Dim e1655 As String

    With ctx
        If ParseStyleOpen() Then
            Do
                p1657 = .BufPos
                p1655 = .BufPos
                e1655 = .LastExpected
                If ParseStyleClose() Then
                    .BufPos = p1657
                    Exit Do
                Else
                    .BufPos = p1655
                    .LastExpected = e1655
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p1657
                    Exit Do
                End If
            Loop
            If ParseStyleClose() Then
                ParseInStyleTags = True
            End If
        End If
    End With
End Function

Private Function ParseStr() As Boolean
    Dim lCaptureBegin As Long
    Dim i1734 As Long
    Dim p1733 As Long
    Dim q1733 As Long
    Dim lCaptureEnd As Long
    Dim p1740 As Long
    Dim q1740 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        pvPushThunk ucsActResultClear
        If ParseStartList() Then
            pvPushThunk ucsActVarSet, 1
            lCaptureBegin = .BufPos
            For i1734 = 0 To LNG_MAXINT
                p1733 = .BufPos
                q1733 = .ThunkPos
                If Not (ParseNormalChar()) Then
                    .BufPos = p1733
                    .ThunkPos = q1733
                    Exit For
                End If
            Next
            If i1734 <> 0 Then
                lCaptureEnd = .BufPos
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                pvPushThunk ucsAct_1_Str, lCaptureBegin, lCaptureEnd
                Do
                    p1740 = .BufPos
                    q1740 = .ThunkPos
                    If Not (ParseStrChunk()) Then
                        .BufPos = p1740
                        .ThunkPos = q1740
                        Exit Do
                    End If
                    .CaptureBegin = lCaptureBegin
                    .CaptureEnd = lCaptureEnd
                    pvPushThunk ucsAct_2_Str, lCaptureBegin, lCaptureEnd
                Loop
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                pvPushThunk ucsAct_3_Str, lCaptureBegin, lCaptureEnd
                pvPushThunk ucsActVarAlloc, -1
                ParseStr = True
            End If
        End If
    End With
End Function

Private Function ParseUlOrStarLine() As Boolean
    Dim p1828 As Long

    With ctx
        p1828 = .BufPos
        If ParseUlLine() Then
            pvPushThunk ucsAct_1_UlOrStarLine, .CaptureBegin, .CaptureEnd
            ParseUlOrStarLine = True
            Exit Function
        Else
            .BufPos = p1828
        End If
        If ParseStarLine() Then
            pvPushThunk ucsAct_1_UlOrStarLine, .CaptureBegin, .CaptureEnd
            ParseUlOrStarLine = True
            Exit Function
        Else
            .BufPos = p1828
        End If
    End With
End Function

Private Function ParseSpace() As Boolean
    Dim i1724 As Long
    Dim p1723 As Long

    With ctx
        For i1724 = 0 To LNG_MAXINT
            p1723 = .BufPos
            If Not (ParseSpacechar()) Then
                .BufPos = p1723
                Exit For
            End If
        Next
        If i1724 <> 0 Then
            pvPushThunk ucsAct_1_Space, .CaptureBegin, .CaptureEnd
            ParseSpace = True
        End If
    End With
End Function

Private Function ParseStrong() As Boolean
    Dim p1922 As Long
    Dim q1922 As Long

    With ctx
        p1922 = .BufPos
        q1922 = .ThunkPos
        If ParseStrongStar() Then
            ParseStrong = True
            Exit Function
        Else
            .BufPos = p1922
            .ThunkPos = q1922
        End If
        If ParseStrongUl() Then
            ParseStrong = True
            Exit Function
        Else
            .BufPos = p1922
            .ThunkPos = q1922
        End If
    End With
End Function

Private Function ParseEmph() As Boolean
    Dim p1865 As Long
    Dim q1865 As Long

    With ctx
        p1865 = .BufPos
        q1865 = .ThunkPos
        If ParseEmphStar() Then
            ParseEmph = True
            Exit Function
        Else
            .BufPos = p1865
            .ThunkPos = q1865
        End If
        If ParseEmphUl() Then
            ParseEmph = True
            Exit Function
        Else
            .BufPos = p1865
            .ThunkPos = q1865
        End If
    End With
End Function

Private Function ParseStrike() As Boolean
    Dim p1965 As Long
    Dim e1965 As String
    Dim i1978 As Long
    Dim p1976 As Long
    Dim q1976 As Long
    Dim p1971 As Long
    Dim q1971 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        If extension(EXT_STRIKE) Then
            If .BufData(.BufPos) = 126 And .BufData(.BufPos + 1) = 126 Then ' "~~"
                .BufPos = .BufPos + 2
                p1965 = .BufPos
                e1965 = .LastExpected
                If Not (ParseWhitespace()) Then
                    .BufPos = p1965
                    .LastExpected = e1965
                    pvPushThunk ucsActResultClear
                    If ParseStartList() Then
                        pvPushThunk ucsActVarSet, 1
                        For i1978 = 0 To LNG_MAXINT
                            p1976 = .BufPos
                            q1976 = .ThunkPos
                            p1971 = .BufPos
                            q1971 = .ThunkPos
                            If .BufData(.BufPos) = 126 And .BufData(.BufPos + 1) = 126 Then ' "~~"
                                .BufPos = p1976
                                .ThunkPos = q1976
                                Exit For
                            Else
                                .BufPos = p1971
                                .ThunkPos = q1971
                            End If
                            pvPushThunk ucsActResultClear
                            If ParseInline() Then
                                pvPushThunk ucsActVarSet, 2
                            Else
                                .BufPos = p1976
                                .ThunkPos = q1976
                                Exit For
                            End If
                            pvPushThunk ucsAct_1_Strike, .CaptureBegin, .CaptureEnd
                        Next
                        If i1978 <> 0 Then
                            If .BufData(.BufPos) = 126 And .BufData(.BufPos + 1) = 126 Then ' "~~"
                                .BufPos = .BufPos + 2
                                pvPushThunk ucsAct_2_Strike, .CaptureBegin, .CaptureEnd
                                pvPushThunk ucsActVarAlloc, -2
                                Call pvSetAdvance
                                ParseStrike = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseImage() As Boolean
    Dim p1986 As Long
    Dim q1986 As Long

    With ctx
        If .BufData(.BufPos) = 33 Then              ' "!"
            .BufPos = .BufPos + 1
            p1986 = .BufPos
            q1986 = .ThunkPos
            If ParseExplicitLink() Then
                pvPushThunk ucsAct_1_Image, .CaptureBegin, .CaptureEnd
                Call pvSetAdvance
                ParseImage = True
                Exit Function
            Else
                .BufPos = p1986
                .ThunkPos = q1986
            End If
            If ParseReferenceLink() Then
                pvPushThunk ucsAct_1_Image, .CaptureBegin, .CaptureEnd
                Call pvSetAdvance
                ParseImage = True
                Exit Function
            Else
                .BufPos = p1986
                .ThunkPos = q1986
            End If
        End If
    End With
End Function

Private Function ParseLink() As Boolean
    Dim p1991 As Long
    Dim q1991 As Long

    With ctx
        p1991 = .BufPos
        q1991 = .ThunkPos
        If ParseExplicitLink() Then
            ParseLink = True
            Exit Function
        Else
            .BufPos = p1991
            .ThunkPos = q1991
        End If
        If ParseReferenceLink() Then
            ParseLink = True
            Exit Function
        Else
            .BufPos = p1991
            .ThunkPos = q1991
        End If
        If ParseAutoLink() Then
            ParseLink = True
            Exit Function
        Else
            .BufPos = p1991
            .ThunkPos = q1991
        End If
    End With
End Function

Private Function ParseNoteReference() As Boolean
    With ctx
        pvPushThunk ucsActVarAlloc, 1
        If extension(EXT_NOTES) Then
            pvPushThunk ucsActResultClear
            If ParseRawNoteReference() Then
                pvPushThunk ucsActVarSet, 1
                pvPushThunk ucsAct_1_NoteReference, .CaptureBegin, .CaptureEnd
                pvPushThunk ucsActVarAlloc, -1
                ParseNoteReference = True
            End If
        End If
    End With
End Function

Private Function ParseInlineNote() As Boolean
    Dim i2956 As Long
    Dim p2954 As Long
    Dim q2954 As Long
    Dim p2952 As Long
    Dim q2952 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        If extension(EXT_NOTES) Then
            If .BufData(.BufPos) = 94 And .BufData(.BufPos + 1) = 91 Then ' "^["
                .BufPos = .BufPos + 2
                pvPushThunk ucsActResultClear
                If ParseStartList() Then
                    pvPushThunk ucsActVarSet, 1
                    For i2956 = 0 To LNG_MAXINT
                        p2954 = .BufPos
                        q2954 = .ThunkPos
                        p2952 = .BufPos
                        q2952 = .ThunkPos
                        If .BufData(.BufPos) = 93 Then ' "]"
                            .BufPos = p2954
                            .ThunkPos = q2954
                            Exit For
                        Else
                            .BufPos = p2952
                            .ThunkPos = q2952
                        End If
                        If Not (ParseInline()) Then
                            .BufPos = p2954
                            .ThunkPos = q2954
                            Exit For
                        End If
                        pvPushThunk ucsAct_1_InlineNote, .CaptureBegin, .CaptureEnd
                    Next
                    If i2956 <> 0 Then
                        If .BufData(.BufPos) = 93 Then ' "]"
                            .BufPos = .BufPos + 1
                            pvPushThunk ucsAct_2_InlineNote, .CaptureBegin, .CaptureEnd
                            pvPushThunk ucsActVarAlloc, -1
                            Call pvSetAdvance
                            ParseInlineNote = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseCode() As Boolean
    Dim p2370 As Long
    Dim lCaptureBegin As Long
    Dim i2336 As Long
    Dim p2324 As Long
    Dim i2318 As Long
    Dim p2317 As Long
    Dim p2315 As Long
    Dim p2320 As Long
    Dim e2320 As String
    Dim p2328 As Long
    Dim p2334 As Long
    Dim p2332 As Long
    Dim e2332 As String
    Dim i2322 As Long
    Dim lCaptureEnd As Long
    Dim i2366 As Long
    Dim p2354 As Long
    Dim i2348 As Long
    Dim p2347 As Long
    Dim p2345 As Long
    Dim p2350 As Long
    Dim e2350 As String
    Dim p2358 As Long
    Dim p2364 As Long
    Dim p2362 As Long
    Dim e2362 As String
    Dim i2352 As Long
    Dim i2397 As Long
    Dim p2385 As Long
    Dim i2379 As Long
    Dim p2378 As Long
    Dim p2376 As Long
    Dim p2381 As Long
    Dim e2381 As String
    Dim p2389 As Long
    Dim p2395 As Long
    Dim p2393 As Long
    Dim e2393 As String
    Dim i2383 As Long
    Dim i2427 As Long
    Dim p2415 As Long
    Dim i2409 As Long
    Dim p2408 As Long
    Dim p2406 As Long
    Dim p2411 As Long
    Dim e2411 As String
    Dim p2419 As Long
    Dim p2425 As Long
    Dim p2423 As Long
    Dim e2423 As String
    Dim i2413 As Long
    Dim i2457 As Long
    Dim p2445 As Long
    Dim i2439 As Long
    Dim p2438 As Long
    Dim p2436 As Long
    Dim p2441 As Long
    Dim e2441 As String
    Dim p2449 As Long
    Dim p2455 As Long
    Dim p2453 As Long
    Dim e2453 As String
    Dim i2443 As Long

    With ctx
        p2370 = .BufPos
        If ParseTicks1() Then
            Call ParseSp
            lCaptureBegin = .BufPos
            For i2336 = 0 To LNG_MAXINT
                p2324 = .BufPos
                For i2318 = 0 To LNG_MAXINT
                    p2317 = .BufPos
                    p2315 = .BufPos
                    If .BufData(.BufPos) = 96 Then  ' "`"
                        .BufPos = p2317
                        Exit For
                    Else
                        .BufPos = p2315
                    End If
                    If Not (ParseNonspacechar()) Then
                        .BufPos = p2317
                        Exit For
                    End If
                Next
                If i2318 <> 0 Then
                Else
                    p2320 = .BufPos
                    e2320 = .LastExpected
                    If ParseTicks1() Then
                        .BufPos = p2324
                        p2328 = .BufPos
                        Call ParseSp
                        If ParseTicks1() Then
                            .BufPos = p2324
                            .BufPos = p2324
                            Exit For
                        Else
                            .BufPos = p2328
                            p2334 = .BufPos
                            If ParseSpacechar() Then
                                GoTo L48
                            Else
                                .BufPos = p2334
                            End If
                            If ParseNewline() Then
                                p2332 = .BufPos
                                e2332 = .LastExpected
                                If ParseBlankLine() Then
                                    .BufPos = p2334
                                Else
                                    .BufPos = p2332
                                    .LastExpected = e2332
                                    GoTo L48
                                End If
                            Else
                                .BufPos = p2334
                            End If
                            .BufPos = p2324
                            .BufPos = p2324
                            Exit For
                        End If
                    Else
                        .BufPos = p2320
                        .LastExpected = e2320
                        For i2322 = 0 To LNG_MAXINT
                            If .BufData(.BufPos) = 96 Then ' "`"
                                .BufPos = .BufPos + 1
                            Else
                                Exit For
                            End If
                        Next
                        If i2322 <> 0 Then
                        Else
                            .BufPos = p2324
                            p2328 = .BufPos
                            Call ParseSp
                            If ParseTicks1() Then
                                .BufPos = p2324
                                .BufPos = p2324
                                Exit For
                            Else
                                .BufPos = p2328
                                p2334 = .BufPos
                                If ParseSpacechar() Then
                                    GoTo L48
                                Else
                                    .BufPos = p2334
                                End If
                                If ParseNewline() Then
                                    p2332 = .BufPos
                                    e2332 = .LastExpected
                                    If ParseBlankLine() Then
                                        .BufPos = p2334
                                    Else
                                        .BufPos = p2332
                                        .LastExpected = e2332
                                        GoTo L48
                                    End If
                                Else
                                    .BufPos = p2334
                                End If
                                .BufPos = p2324
                                .BufPos = p2324
                                Exit For
                            End If
                        End If
                    End If
                End If
L48:
            Next
            If i2336 <> 0 Then
                lCaptureEnd = .BufPos
                Call ParseSp
                If ParseTicks1() Then
                    .CaptureBegin = lCaptureBegin
                    .CaptureEnd = lCaptureEnd
                    pvPushThunk ucsAct_1_Code, lCaptureBegin, lCaptureEnd
                    Call pvSetAdvance
                    ParseCode = True
                    Exit Function
                Else
                    .BufPos = p2370
                End If
            Else
                .BufPos = p2370
            End If
        Else
            .BufPos = p2370
        End If
        If ParseTicks2() Then
            Call ParseSp
            lCaptureBegin = .BufPos
            For i2366 = 0 To LNG_MAXINT
                p2354 = .BufPos
                For i2348 = 0 To LNG_MAXINT
                    p2347 = .BufPos
                    p2345 = .BufPos
                    If .BufData(.BufPos) = 96 Then  ' "`"
                        .BufPos = p2347
                        Exit For
                    Else
                        .BufPos = p2345
                    End If
                    If Not (ParseNonspacechar()) Then
                        .BufPos = p2347
                        Exit For
                    End If
                Next
                If i2348 <> 0 Then
                Else
                    p2350 = .BufPos
                    e2350 = .LastExpected
                    If ParseTicks2() Then
                        .BufPos = p2354
                        p2358 = .BufPos
                        Call ParseSp
                        If ParseTicks2() Then
                            .BufPos = p2354
                            .BufPos = p2354
                            Exit For
                        Else
                            .BufPos = p2358
                            p2364 = .BufPos
                            If ParseSpacechar() Then
                                GoTo L49
                            Else
                                .BufPos = p2364
                            End If
                            If ParseNewline() Then
                                p2362 = .BufPos
                                e2362 = .LastExpected
                                If ParseBlankLine() Then
                                    .BufPos = p2364
                                Else
                                    .BufPos = p2362
                                    .LastExpected = e2362
                                    GoTo L49
                                End If
                            Else
                                .BufPos = p2364
                            End If
                            .BufPos = p2354
                            .BufPos = p2354
                            Exit For
                        End If
                    Else
                        .BufPos = p2350
                        .LastExpected = e2350
                        For i2352 = 0 To LNG_MAXINT
                            If .BufData(.BufPos) = 96 Then ' "`"
                                .BufPos = .BufPos + 1
                            Else
                                Exit For
                            End If
                        Next
                        If i2352 <> 0 Then
                        Else
                            .BufPos = p2354
                            p2358 = .BufPos
                            Call ParseSp
                            If ParseTicks2() Then
                                .BufPos = p2354
                                .BufPos = p2354
                                Exit For
                            Else
                                .BufPos = p2358
                                p2364 = .BufPos
                                If ParseSpacechar() Then
                                    GoTo L49
                                Else
                                    .BufPos = p2364
                                End If
                                If ParseNewline() Then
                                    p2362 = .BufPos
                                    e2362 = .LastExpected
                                    If ParseBlankLine() Then
                                        .BufPos = p2364
                                    Else
                                        .BufPos = p2362
                                        .LastExpected = e2362
                                        GoTo L49
                                    End If
                                Else
                                    .BufPos = p2364
                                End If
                                .BufPos = p2354
                                .BufPos = p2354
                                Exit For
                            End If
                        End If
                    End If
                End If
L49:
            Next
            If i2366 <> 0 Then
                lCaptureEnd = .BufPos
                Call ParseSp
                If ParseTicks2() Then
                    .CaptureBegin = lCaptureBegin
                    .CaptureEnd = lCaptureEnd
                    pvPushThunk ucsAct_1_Code, lCaptureBegin, lCaptureEnd
                    Call pvSetAdvance
                    ParseCode = True
                    Exit Function
                Else
                    .BufPos = p2370
                End If
            Else
                .BufPos = p2370
            End If
        Else
            .BufPos = p2370
        End If
        If ParseTicks3() Then
            Call ParseSp
            lCaptureBegin = .BufPos
            For i2397 = 0 To LNG_MAXINT
                p2385 = .BufPos
                For i2379 = 0 To LNG_MAXINT
                    p2378 = .BufPos
                    p2376 = .BufPos
                    If .BufData(.BufPos) = 96 Then  ' "`"
                        .BufPos = p2378
                        Exit For
                    Else
                        .BufPos = p2376
                    End If
                    If Not (ParseNonspacechar()) Then
                        .BufPos = p2378
                        Exit For
                    End If
                Next
                If i2379 <> 0 Then
                Else
                    p2381 = .BufPos
                    e2381 = .LastExpected
                    If ParseTicks3() Then
                        .BufPos = p2385
                        p2389 = .BufPos
                        Call ParseSp
                        If ParseTicks3() Then
                            .BufPos = p2385
                            .BufPos = p2385
                            Exit For
                        Else
                            .BufPos = p2389
                            p2395 = .BufPos
                            If ParseSpacechar() Then
                                GoTo L50
                            Else
                                .BufPos = p2395
                            End If
                            If ParseNewline() Then
                                p2393 = .BufPos
                                e2393 = .LastExpected
                                If ParseBlankLine() Then
                                    .BufPos = p2395
                                Else
                                    .BufPos = p2393
                                    .LastExpected = e2393
                                    GoTo L50
                                End If
                            Else
                                .BufPos = p2395
                            End If
                            .BufPos = p2385
                            .BufPos = p2385
                            Exit For
                        End If
                    Else
                        .BufPos = p2381
                        .LastExpected = e2381
                        For i2383 = 0 To LNG_MAXINT
                            If .BufData(.BufPos) = 96 Then ' "`"
                                .BufPos = .BufPos + 1
                            Else
                                Exit For
                            End If
                        Next
                        If i2383 <> 0 Then
                        Else
                            .BufPos = p2385
                            p2389 = .BufPos
                            Call ParseSp
                            If ParseTicks3() Then
                                .BufPos = p2385
                                .BufPos = p2385
                                Exit For
                            Else
                                .BufPos = p2389
                                p2395 = .BufPos
                                If ParseSpacechar() Then
                                    GoTo L50
                                Else
                                    .BufPos = p2395
                                End If
                                If ParseNewline() Then
                                    p2393 = .BufPos
                                    e2393 = .LastExpected
                                    If ParseBlankLine() Then
                                        .BufPos = p2395
                                    Else
                                        .BufPos = p2393
                                        .LastExpected = e2393
                                        GoTo L50
                                    End If
                                Else
                                    .BufPos = p2395
                                End If
                                .BufPos = p2385
                                .BufPos = p2385
                                Exit For
                            End If
                        End If
                    End If
                End If
L50:
            Next
            If i2397 <> 0 Then
                lCaptureEnd = .BufPos
                Call ParseSp
                If ParseTicks3() Then
                    .CaptureBegin = lCaptureBegin
                    .CaptureEnd = lCaptureEnd
                    pvPushThunk ucsAct_1_Code, lCaptureBegin, lCaptureEnd
                    Call pvSetAdvance
                    ParseCode = True
                    Exit Function
                Else
                    .BufPos = p2370
                End If
            Else
                .BufPos = p2370
            End If
        Else
            .BufPos = p2370
        End If
        If ParseTicks4() Then
            Call ParseSp
            lCaptureBegin = .BufPos
            For i2427 = 0 To LNG_MAXINT
                p2415 = .BufPos
                For i2409 = 0 To LNG_MAXINT
                    p2408 = .BufPos
                    p2406 = .BufPos
                    If .BufData(.BufPos) = 96 Then  ' "`"
                        .BufPos = p2408
                        Exit For
                    Else
                        .BufPos = p2406
                    End If
                    If Not (ParseNonspacechar()) Then
                        .BufPos = p2408
                        Exit For
                    End If
                Next
                If i2409 <> 0 Then
                Else
                    p2411 = .BufPos
                    e2411 = .LastExpected
                    If ParseTicks4() Then
                        .BufPos = p2415
                        p2419 = .BufPos
                        Call ParseSp
                        If ParseTicks4() Then
                            .BufPos = p2415
                            .BufPos = p2415
                            Exit For
                        Else
                            .BufPos = p2419
                            p2425 = .BufPos
                            If ParseSpacechar() Then
                                GoTo L51
                            Else
                                .BufPos = p2425
                            End If
                            If ParseNewline() Then
                                p2423 = .BufPos
                                e2423 = .LastExpected
                                If ParseBlankLine() Then
                                    .BufPos = p2425
                                Else
                                    .BufPos = p2423
                                    .LastExpected = e2423
                                    GoTo L51
                                End If
                            Else
                                .BufPos = p2425
                            End If
                            .BufPos = p2415
                            .BufPos = p2415
                            Exit For
                        End If
                    Else
                        .BufPos = p2411
                        .LastExpected = e2411
                        For i2413 = 0 To LNG_MAXINT
                            If .BufData(.BufPos) = 96 Then ' "`"
                                .BufPos = .BufPos + 1
                            Else
                                Exit For
                            End If
                        Next
                        If i2413 <> 0 Then
                        Else
                            .BufPos = p2415
                            p2419 = .BufPos
                            Call ParseSp
                            If ParseTicks4() Then
                                .BufPos = p2415
                                .BufPos = p2415
                                Exit For
                            Else
                                .BufPos = p2419
                                p2425 = .BufPos
                                If ParseSpacechar() Then
                                    GoTo L51
                                Else
                                    .BufPos = p2425
                                End If
                                If ParseNewline() Then
                                    p2423 = .BufPos
                                    e2423 = .LastExpected
                                    If ParseBlankLine() Then
                                        .BufPos = p2425
                                    Else
                                        .BufPos = p2423
                                        .LastExpected = e2423
                                        GoTo L51
                                    End If
                                Else
                                    .BufPos = p2425
                                End If
                                .BufPos = p2415
                                .BufPos = p2415
                                Exit For
                            End If
                        End If
                    End If
                End If
L51:
            Next
            If i2427 <> 0 Then
                lCaptureEnd = .BufPos
                Call ParseSp
                If ParseTicks4() Then
                    .CaptureBegin = lCaptureBegin
                    .CaptureEnd = lCaptureEnd
                    pvPushThunk ucsAct_1_Code, lCaptureBegin, lCaptureEnd
                    Call pvSetAdvance
                    ParseCode = True
                    Exit Function
                Else
                    .BufPos = p2370
                End If
            Else
                .BufPos = p2370
            End If
        Else
            .BufPos = p2370
        End If
        If ParseTicks5() Then
            Call ParseSp
            lCaptureBegin = .BufPos
            For i2457 = 0 To LNG_MAXINT
                p2445 = .BufPos
                For i2439 = 0 To LNG_MAXINT
                    p2438 = .BufPos
                    p2436 = .BufPos
                    If .BufData(.BufPos) = 96 Then  ' "`"
                        .BufPos = p2438
                        Exit For
                    Else
                        .BufPos = p2436
                    End If
                    If Not (ParseNonspacechar()) Then
                        .BufPos = p2438
                        Exit For
                    End If
                Next
                If i2439 <> 0 Then
                Else
                    p2441 = .BufPos
                    e2441 = .LastExpected
                    If ParseTicks5() Then
                        .BufPos = p2445
                        p2449 = .BufPos
                        Call ParseSp
                        If ParseTicks5() Then
                            .BufPos = p2445
                            .BufPos = p2445
                            Exit For
                        Else
                            .BufPos = p2449
                            p2455 = .BufPos
                            If ParseSpacechar() Then
                                GoTo L52
                            Else
                                .BufPos = p2455
                            End If
                            If ParseNewline() Then
                                p2453 = .BufPos
                                e2453 = .LastExpected
                                If ParseBlankLine() Then
                                    .BufPos = p2455
                                Else
                                    .BufPos = p2453
                                    .LastExpected = e2453
                                    GoTo L52
                                End If
                            Else
                                .BufPos = p2455
                            End If
                            .BufPos = p2445
                            .BufPos = p2445
                            Exit For
                        End If
                    Else
                        .BufPos = p2441
                        .LastExpected = e2441
                        For i2443 = 0 To LNG_MAXINT
                            If .BufData(.BufPos) = 96 Then ' "`"
                                .BufPos = .BufPos + 1
                            Else
                                Exit For
                            End If
                        Next
                        If i2443 <> 0 Then
                        Else
                            .BufPos = p2445
                            p2449 = .BufPos
                            Call ParseSp
                            If ParseTicks5() Then
                                .BufPos = p2445
                                .BufPos = p2445
                                Exit For
                            Else
                                .BufPos = p2449
                                p2455 = .BufPos
                                If ParseSpacechar() Then
                                    GoTo L52
                                Else
                                    .BufPos = p2455
                                End If
                                If ParseNewline() Then
                                    p2453 = .BufPos
                                    e2453 = .LastExpected
                                    If ParseBlankLine() Then
                                        .BufPos = p2455
                                    Else
                                        .BufPos = p2453
                                        .LastExpected = e2453
                                        GoTo L52
                                    End If
                                Else
                                    .BufPos = p2455
                                End If
                                .BufPos = p2445
                                .BufPos = p2445
                                Exit For
                            End If
                        End If
                    End If
                End If
L52:
            Next
            If i2457 <> 0 Then
                lCaptureEnd = .BufPos
                Call ParseSp
                If ParseTicks5() Then
                    .CaptureBegin = lCaptureBegin
                    .CaptureEnd = lCaptureEnd
                    pvPushThunk ucsAct_1_Code, lCaptureBegin, lCaptureEnd
                    Call pvSetAdvance
                    ParseCode = True
                    Exit Function
                Else
                    .BufPos = p2370
                End If
            Else
                .BufPos = p2370
            End If
        Else
            .BufPos = p2370
        End If
    End With
End Function

Private Function ParseRawHtml() As Boolean
    Dim lCaptureBegin As Long
    Dim p2466 As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        p2466 = .BufPos
        If ParseHtmlComment() Then
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_RawHtml, lCaptureBegin, lCaptureEnd
            ParseRawHtml = True
            Exit Function
        Else
            .BufPos = p2466
        End If
        If ParseHtmlBlockScript() Then
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_RawHtml, lCaptureBegin, lCaptureEnd
            ParseRawHtml = True
            Exit Function
        Else
            .BufPos = p2466
        End If
        If ParseHtmlTag() Then
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_RawHtml, lCaptureBegin, lCaptureEnd
            ParseRawHtml = True
            Exit Function
        Else
            .BufPos = p2466
        End If
    End With
End Function

Private Function ParseEntity() As Boolean
    Dim p1777 As Long

    With ctx
        p1777 = .BufPos
        If ParseHexEntity() Then
            pvPushThunk ucsAct_1_Entity, .CaptureBegin, .CaptureEnd
            ParseEntity = True
            Exit Function
        Else
            .BufPos = p1777
        End If
        If ParseDecEntity() Then
            pvPushThunk ucsAct_1_Entity, .CaptureBegin, .CaptureEnd
            ParseEntity = True
            Exit Function
        Else
            .BufPos = p1777
        End If
        If ParseCharEntity() Then
            pvPushThunk ucsAct_1_Entity, .CaptureBegin, .CaptureEnd
            ParseEntity = True
            Exit Function
        Else
            .BufPos = p1777
        End If
    End With
End Function

Private Function ParseEscapedChar() As Boolean
    Dim p1767 As Long
    Dim e1767 As String
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            p1767 = .BufPos
            e1767 = .LastExpected
            If Not (ParseNewline()) Then
                .BufPos = p1767
                .LastExpected = e1767
                lCaptureBegin = .BufPos
                Select Case .BufData(.BufPos)
                Case 45, 92, 96, 124, 42, 95, 123, 125, 91, 93, 40, 41, 35, 43, 46, 33, 62, 60 ' [-\\`|*_{}[\]()#+.!><]
                    .BufPos = .BufPos + 1
                    lCaptureEnd = .BufPos
                    .CaptureBegin = lCaptureBegin
                    .CaptureEnd = lCaptureEnd
                    pvPushThunk ucsAct_1_EscapedChar, lCaptureBegin, lCaptureEnd
                    Call pvSetAdvance
                    ParseEscapedChar = True
                End Select
            End If
        End If
    End With
End Function

Private Function ParseSmart() As Boolean
    Dim p2818 As Long
    Dim q2818 As Long

    With ctx
        If extension(EXT_SMART) Then
            p2818 = .BufPos
            q2818 = .ThunkPos
            If ParseEllipsis() Then
                ParseSmart = True
                Exit Function
            Else
                .BufPos = p2818
                .ThunkPos = q2818
            End If
            If ParseDash() Then
                ParseSmart = True
                Exit Function
            Else
                .BufPos = p2818
                .ThunkPos = q2818
            End If
            If ParseSingleQuoted() Then
                ParseSmart = True
                Exit Function
            Else
                .BufPos = p2818
                .ThunkPos = q2818
            End If
            If ParseDoubleQuoted() Then
                ParseSmart = True
                Exit Function
            Else
                .BufPos = p2818
                .ThunkPos = q2818
            End If
            If ParseApostrophe() Then
                ParseSmart = True
                Exit Function
            Else
                .BufPos = p2818
                .ThunkPos = q2818
            End If
        End If
    End With
End Function

Private Function ParseSymbol() As Boolean
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        If ParseSpecialChar() Then
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_Symbol, lCaptureBegin, lCaptureEnd
            ParseSymbol = True
        End If
    End With
End Function

Private Function ParseNormalChar() As Boolean
    Dim p2581 As Long
    Dim p2579 As Long

    With ctx
        p2581 = .BufPos
        p2579 = .BufPos
        If Not (ParseSpecialChar()) Then
            .BufPos = p2579
            If Not (ParseSpacechar()) Then
                .BufPos = p2579
                If Not (ParseNewline()) Then
                    .BufPos = p2579
                    .BufPos = p2581
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                        ParseNormalChar = True
                        Exit Function
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseStrChunk() As Boolean
    Dim p1758 As Long
    Dim q1758 As Long
    Dim lCaptureBegin As Long
    Dim i1752 As Long
    Dim p1751 As Long
    Dim i1746 As Long
    Dim p1749 As Long
    Dim lCaptureEnd As Long

    With ctx
        p1758 = .BufPos
        q1758 = .ThunkPos
        lCaptureBegin = .BufPos
        For i1752 = 0 To LNG_MAXINT
            p1751 = .BufPos
            If Not (ParseNormalChar()) Then
                .BufPos = p1751
                For i1746 = 0 To LNG_MAXINT
                    If .BufData(.BufPos) = 95 Then  ' "_"
                        .BufPos = .BufPos + 1
                    Else
                        Exit For
                    End If
                Next
                If i1746 <> 0 Then
                    p1749 = .BufPos
                    If ParseAlphanumeric() Then
                        .BufPos = p1749
                    Else
                        .BufPos = p1751
                        .BufPos = p1751
                        Exit For
                    End If
                Else
                    .BufPos = p1751
                    .BufPos = p1751
                    Exit For
                End If
            End If
        Next
        If i1752 <> 0 Then
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_StrChunk, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseStrChunk = True
            Exit Function
        Else
            .BufPos = p1758
            .ThunkPos = q1758
        End If
        If ParseAposChunk() Then
            Call pvSetAdvance
            ParseStrChunk = True
            Exit Function
        Else
            .BufPos = p1758
            .ThunkPos = q1758
        End If
    End With
End Function

Private Function ParseAlphanumeric() As Boolean
    With ctx
        Select Case .BufData(.BufPos)
        Case 48 To 57, 65 To 90, 97 To 122          ' [0-9A-Za-z]
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End Select
        If .BufData(.BufPos) = 1026 Then            ' "\200"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1027 Then            ' "\201"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8218 Then            ' "\202"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1107 Then            ' "\203"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8222 Then            ' "\204"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8230 Then            ' "\205"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8224 Then            ' "\206"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8225 Then            ' "\207"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8364 Then            ' "\210"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8240 Then            ' "\211"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1033 Then            ' "\212"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8249 Then            ' "\213"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1034 Then            ' "\214"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1036 Then            ' "\215"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1035 Then            ' "\216"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1039 Then            ' "\217"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1106 Then            ' "\220"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8216 Then            ' "\221"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8217 Then            ' "\222"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8220 Then            ' "\223"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8221 Then            ' "\224"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8226 Then            ' "\225"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8211 Then            ' "\226"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8212 Then            ' "\227"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 152 Then             ' "\230"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8482 Then            ' "\231"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1113 Then            ' "\232"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8250 Then            ' "\233"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1114 Then            ' "\234"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1116 Then            ' "\235"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1115 Then            ' "\236"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1119 Then            ' "\237"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 160 Then             ' "\240"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1038 Then            ' "\241"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1118 Then            ' "\242"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1032 Then            ' "\243"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 164 Then             ' "\244"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1168 Then            ' "\245"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 166 Then             ' "\246"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 167 Then             ' "\247"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1025 Then            ' "\250"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 169 Then             ' "\251"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1028 Then            ' "\252"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 171 Then             ' "\253"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 172 Then             ' "\254"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 173 Then             ' "\255"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 174 Then             ' "\256"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1031 Then            ' "\257"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 176 Then             ' "\260"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 177 Then             ' "\261"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1030 Then            ' "\262"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1110 Then            ' "\263"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1169 Then            ' "\264"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 181 Then             ' "\265"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 182 Then             ' "\266"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 183 Then             ' "\267"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1105 Then            ' "\270"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 8470 Then            ' "\271"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1108 Then            ' "\272"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 187 Then             ' "\273"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1112 Then            ' "\274"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1029 Then            ' "\275"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1109 Then            ' "\276"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1111 Then            ' "\277"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1040 Then            ' "\300"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1041 Then            ' "\301"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1042 Then            ' "\302"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1043 Then            ' "\303"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1044 Then            ' "\304"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1045 Then            ' "\305"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1046 Then            ' "\306"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1047 Then            ' "\307"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1048 Then            ' "\310"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1049 Then            ' "\311"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1050 Then            ' "\312"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1051 Then            ' "\313"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1052 Then            ' "\314"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1053 Then            ' "\315"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1054 Then            ' "\316"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1055 Then            ' "\317"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1056 Then            ' "\320"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1057 Then            ' "\321"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1058 Then            ' "\322"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1059 Then            ' "\323"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1060 Then            ' "\324"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1061 Then            ' "\325"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1062 Then            ' "\326"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1063 Then            ' "\327"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1064 Then            ' "\330"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1065 Then            ' "\331"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1066 Then            ' "\332"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1067 Then            ' "\333"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1068 Then            ' "\334"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1069 Then            ' "\335"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1070 Then            ' "\336"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1071 Then            ' "\337"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1072 Then            ' "\340"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1073 Then            ' "\341"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1074 Then            ' "\342"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1075 Then            ' "\343"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1076 Then            ' "\344"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1077 Then            ' "\345"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1078 Then            ' "\346"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1079 Then            ' "\347"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1080 Then            ' "\350"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1081 Then            ' "\351"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1082 Then            ' "\352"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1083 Then            ' "\353"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1084 Then            ' "\354"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1085 Then            ' "\355"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1086 Then            ' "\356"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1087 Then            ' "\357"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1088 Then            ' "\360"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1089 Then            ' "\361"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1090 Then            ' "\362"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1091 Then            ' "\363"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1092 Then            ' "\364"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1093 Then            ' "\365"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1094 Then            ' "\366"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1095 Then            ' "\367"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1096 Then            ' "\370"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1097 Then            ' "\371"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1098 Then            ' "\372"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1099 Then            ' "\373"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1100 Then            ' "\374"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1101 Then            ' "\375"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1102 Then            ' "\376"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
            Exit Function
        End If
        If .BufData(.BufPos) = 1103 Then            ' "\377"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseAlphanumeric = True
        End If
    End With
End Function

Private Function ParseAposChunk() As Boolean
    Dim p1763 As Long

    With ctx
        If extension(EXT_SMART) Then
            If .BufData(.BufPos) = 39 Then          ' "\'"
                .BufPos = .BufPos + 1
                p1763 = .BufPos
                If ParseAlphanumeric() Then
                    .BufPos = p1763
                    pvPushThunk ucsAct_1_AposChunk, .CaptureBegin, .CaptureEnd
                    Call pvSetAdvance
                    ParseAposChunk = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseHexEntity() As Boolean
    Dim lCaptureBegin As Long
    Dim i2724 As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        If .BufData(.BufPos) = 38 Then              ' "&"
            .BufPos = .BufPos + 1
            If .BufData(.BufPos) = 35 Then          ' "#"
                .BufPos = .BufPos + 1
                Select Case .BufData(.BufPos)
                Case 88, 120                        ' [Xx]
                    .BufPos = .BufPos + 1
                    For i2724 = 0 To LNG_MAXINT
                        Select Case .BufData(.BufPos)
                        Case 48 To 57, 97 To 102, 65 To 70 ' [0-9a-fA-F]
                            .BufPos = .BufPos + 1
                        Case Else
                            Exit For
                        End Select
                    Next
                    If i2724 <> 0 Then
                        If .BufData(.BufPos) = 59 Then ' ";"
                            .BufPos = .BufPos + 1
                            lCaptureEnd = .BufPos
                            .CaptureBegin = lCaptureBegin
                            .CaptureEnd = lCaptureEnd
                            Call pvSetAdvance
                            ParseHexEntity = True
                        End If
                    End If
                End Select
            End If
        End If
    End With
End Function

Private Function ParseDecEntity() As Boolean
    Dim lCaptureBegin As Long
    Dim i2732 As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        If .BufData(.BufPos) = 38 Then              ' "&"
            .BufPos = .BufPos + 1
            If .BufData(.BufPos) = 35 Then          ' "#"
                .BufPos = .BufPos + 1
                For i2732 = 0 To LNG_MAXINT
                    Select Case .BufData(.BufPos)
                    Case 48 To 57                   ' [0-9]
                        .BufPos = .BufPos + 1
                    Case Else
                        Exit For
                    End Select
                Next
                If i2732 <> 0 Then
                    lCaptureEnd = .BufPos
                    If .BufData(.BufPos) = 59 Then  ' ";"
                        .BufPos = .BufPos + 1
                        lCaptureEnd = .BufPos
                        .CaptureBegin = lCaptureBegin
                        .CaptureEnd = lCaptureEnd
                        Call pvSetAdvance
                        ParseDecEntity = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseCharEntity() As Boolean
    Dim lCaptureBegin As Long
    Dim i2740 As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        If .BufData(.BufPos) = 38 Then              ' "&"
            .BufPos = .BufPos + 1
            For i2740 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 65 To 90, 97 To 122, 48 To 57  ' [A-Za-z0-9]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If i2740 <> 0 Then
                If .BufData(.BufPos) = 59 Then      ' ";"
                    .BufPos = .BufPos + 1
                    lCaptureEnd = .BufPos
                    .CaptureBegin = lCaptureBegin
                    .CaptureEnd = lCaptureEnd
                    Call pvSetAdvance
                    ParseCharEntity = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseLineBreak() As Boolean
    With ctx
        If .BufData(.BufPos) = 32 And .BufData(.BufPos + 1) = 32 Then ' "  "
            .BufPos = .BufPos + 2
            If ParseNormalEndline() Then
                pvPushThunk ucsAct_1_LineBreak, .CaptureBegin, .CaptureEnd
                Call pvSetAdvance
                ParseLineBreak = True
            End If
        End If
    End With
End Function

Private Function ParseTerminalEndline() As Boolean
    With ctx
        Call ParseSp
        If ParseNewline() Then
            If ParseEof() Then
                pvPushThunk ucsAct_1_TerminalEndline, .CaptureBegin, .CaptureEnd
                ParseTerminalEndline = True
            End If
        End If
    End With
End Function

Private Function ParseNormalEndline() As Boolean
    Dim p1793 As Long
    Dim e1793 As String
    Dim p1795 As Long
    Dim p1797 As Long
    Dim q1797 As Long
    Dim e1797 As String
    Dim p1806 As Long
    Dim q1806 As Long
    Dim i1800 As Long
    Dim i1802 As Long

    With ctx
        Call ParseSp
        If ParseNewline() Then
            p1793 = .BufPos
            e1793 = .LastExpected
            If Not (ParseBlankLine()) Then
                .BufPos = p1793
                .LastExpected = e1793
                p1795 = .BufPos
                If Not (.BufData(.BufPos) = 62) Then ' ">"
                    .BufPos = p1795
                    p1797 = .BufPos
                    q1797 = .ThunkPos
                    e1797 = .LastExpected
                    If Not (ParseAtxStart()) Then
                        .BufPos = p1797
                        .ThunkPos = q1797
                        .LastExpected = e1797
                        p1806 = .BufPos
                        q1806 = .ThunkPos
                        If Not (ParseLine()) Then
                            .BufPos = p1806
                            .ThunkPos = q1806
                            pvPushThunk ucsAct_1_NormalEndline, .CaptureBegin, .CaptureEnd
                            Call pvSetAdvance
                            ParseNormalEndline = True
                            Exit Function
                        End If
                        For i1800 = 0 To LNG_MAXINT
                            If .BufData(.BufPos) = 61 Then ' "="
                                .BufPos = .BufPos + 1
                            Else
                                Exit For
                            End If
                        Next
                        If i1800 <> 0 Then
                        Else
                            For i1802 = 0 To LNG_MAXINT
                                If .BufData(.BufPos) = 45 Then ' "-"
                                    .BufPos = .BufPos + 1
                                Else
                                    Exit For
                                End If
                            Next
                            If i1802 <> 0 Then
                            Else
                                .BufPos = p1806
                                .ThunkPos = q1806
                                pvPushThunk ucsAct_1_NormalEndline, .CaptureBegin, .CaptureEnd
                                Call pvSetAdvance
                                ParseNormalEndline = True
                                Exit Function
                            End If
                        End If
                        If Not (ParseNewline()) Then
                            .BufPos = p1806
                            .ThunkPos = q1806
                            pvPushThunk ucsAct_1_NormalEndline, .CaptureBegin, .CaptureEnd
                            Call pvSetAdvance
                            ParseNormalEndline = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseEof() As Boolean
    Dim p2535 As Long

    With ctx
        p2535 = .BufPos
        If Not (.BufPos < .BufSize) Then
            .BufPos = p2535
            ParseEof = True
        End If
    End With
End Function

Private Function ParseSpecialChar() As Boolean
    Dim p2561 As Long

    With ctx
        p2561 = .BufPos
        If .BufData(.BufPos) = 126 Then             ' "~"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseSpecialChar = True
            Exit Function
        End If
        If .BufData(.BufPos) = 42 Then              ' "*"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseSpecialChar = True
            Exit Function
        End If
        If .BufData(.BufPos) = 95 Then              ' "_"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseSpecialChar = True
            Exit Function
        End If
        If .BufData(.BufPos) = 96 Then              ' "`"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseSpecialChar = True
            Exit Function
        End If
        If .BufData(.BufPos) = 38 Then              ' "&"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseSpecialChar = True
            Exit Function
        End If
        If .BufData(.BufPos) = 91 Then              ' "["
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseSpecialChar = True
            Exit Function
        End If
        If .BufData(.BufPos) = 93 Then              ' "]"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseSpecialChar = True
            Exit Function
        End If
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseSpecialChar = True
            Exit Function
        End If
        If .BufData(.BufPos) = 41 Then              ' ")"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseSpecialChar = True
            Exit Function
        End If
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseSpecialChar = True
            Exit Function
        End If
        If .BufData(.BufPos) = 33 Then              ' "!"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseSpecialChar = True
            Exit Function
        End If
        If .BufData(.BufPos) = 35 Then              ' "#"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseSpecialChar = True
            Exit Function
        End If
        If .BufData(.BufPos) = 92 Then              ' "\\"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseSpecialChar = True
            Exit Function
        End If
        If .BufData(.BufPos) = 39 Then              ' "\'"
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseSpecialChar = True
            Exit Function
        End If
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            Call pvSetAdvance
            ParseSpecialChar = True
            Exit Function
        End If
        If ParseExtendedSpecialChar() Then
            Call pvSetAdvance
            ParseSpecialChar = True
            Exit Function
        Else
            .BufPos = p2561
        End If
    End With
End Function

Private Function ParseUlLine() As Boolean
    Dim p1860 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long
    Dim i1856 As Long
    Dim p1858 As Long

    With ctx
        p1860 = .BufPos
        lCaptureBegin = .BufPos
        If pvMatchString("____") Then               ' "____"
            .BufPos = .BufPos + 4
            Do
                If .BufData(.BufPos) = 95 Then      ' "_"
                    .BufPos = .BufPos + 1
                Else
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            Call pvSetAdvance
            ParseUlLine = True
            Exit Function
        End If
        lCaptureBegin = .BufPos
        If ParseSpacechar() Then
            For i1856 = 0 To LNG_MAXINT
                If .BufData(.BufPos) = 95 Then      ' "_"
                    .BufPos = .BufPos + 1
                Else
                    Exit For
                End If
            Next
            If i1856 <> 0 Then
                p1858 = .BufPos
                If ParseSpacechar() Then
                    .BufPos = p1858
                    lCaptureEnd = .BufPos
                    .CaptureBegin = lCaptureBegin
                    .CaptureEnd = lCaptureEnd
                    Call pvSetAdvance
                    ParseUlLine = True
                    Exit Function
                Else
                    .BufPos = p1860
                End If
            Else
                .BufPos = p1860
            End If
        Else
            .BufPos = p1860
        End If
    End With
End Function

Private Function ParseStarLine() As Boolean
    Dim p1845 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long
    Dim i1841 As Long
    Dim p1843 As Long

    With ctx
        p1845 = .BufPos
        lCaptureBegin = .BufPos
        If pvMatchString("****") Then               ' "****"
            .BufPos = .BufPos + 4
            Do
                If .BufData(.BufPos) = 42 Then      ' "*"
                    .BufPos = .BufPos + 1
                Else
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            Call pvSetAdvance
            ParseStarLine = True
            Exit Function
        End If
        lCaptureBegin = .BufPos
        If ParseSpacechar() Then
            For i1841 = 0 To LNG_MAXINT
                If .BufData(.BufPos) = 42 Then      ' "*"
                    .BufPos = .BufPos + 1
                Else
                    Exit For
                End If
            Next
            If i1841 <> 0 Then
                p1843 = .BufPos
                If ParseSpacechar() Then
                    .BufPos = p1843
                    lCaptureEnd = .BufPos
                    .CaptureBegin = lCaptureBegin
                    .CaptureEnd = lCaptureEnd
                    Call pvSetAdvance
                    ParseStarLine = True
                    Exit Function
                Else
                    .BufPos = p1845
                End If
            Else
                .BufPos = p1845
            End If
        Else
            .BufPos = p1845
        End If
    End With
End Function

Private Function ParseEmphStar() As Boolean
    Dim p1872 As Long
    Dim e1872 As String
    Dim i1892 As Long
    Dim p1891 As Long
    Dim q1891 As Long
    Dim p1879 As Long
    Dim q1879 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        If .BufData(.BufPos) = 42 Then              ' "*"
            .BufPos = .BufPos + 1
            p1872 = .BufPos
            e1872 = .LastExpected
            If Not (ParseWhitespace()) Then
                .BufPos = p1872
                .LastExpected = e1872
                pvPushThunk ucsActResultClear
                If ParseStartList() Then
                    pvPushThunk ucsActVarSet, 1
                    For i1892 = 0 To LNG_MAXINT
                        p1891 = .BufPos
                        q1891 = .ThunkPos
                        p1879 = .BufPos
                        q1879 = .ThunkPos
                        If .BufData(.BufPos) = 42 Then ' "*"
                            .BufPos = p1891
                            .ThunkPos = q1891
                            pvPushThunk ucsActResultClear
                            If ParseStrongStar() Then
                                pvPushThunk ucsActVarSet, 2
                                pvPushThunk ucsAct_2_EmphStar, .CaptureBegin, .CaptureEnd
                            Else
                                .BufPos = p1891
                                .ThunkPos = q1891
                                .BufPos = p1891
                                .ThunkPos = q1891
                                Exit For
                            End If
                        Else
                            .BufPos = p1879
                            .ThunkPos = q1879
                            pvPushThunk ucsActResultClear
                            If ParseInline() Then
                                pvPushThunk ucsActVarSet, 2
                                pvPushThunk ucsAct_1_EmphStar, .CaptureBegin, .CaptureEnd
                            Else
                                .BufPos = p1891
                                .ThunkPos = q1891
                                pvPushThunk ucsActResultClear
                                If ParseStrongStar() Then
                                    pvPushThunk ucsActVarSet, 2
                                    pvPushThunk ucsAct_2_EmphStar, .CaptureBegin, .CaptureEnd
                                Else
                                    .BufPos = p1891
                                    .ThunkPos = q1891
                                    .BufPos = p1891
                                    .ThunkPos = q1891
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                    If i1892 <> 0 Then
                        If .BufData(.BufPos) = 42 Then ' "*"
                            .BufPos = .BufPos + 1
                            pvPushThunk ucsAct_3_EmphStar, .CaptureBegin, .CaptureEnd
                            pvPushThunk ucsActVarAlloc, -2
                            Call pvSetAdvance
                            ParseEmphStar = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseEmphUl() As Boolean
    Dim p1897 As Long
    Dim e1897 As String
    Dim i1917 As Long
    Dim p1916 As Long
    Dim q1916 As Long
    Dim p1904 As Long
    Dim q1904 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        If .BufData(.BufPos) = 95 Then              ' "_"
            .BufPos = .BufPos + 1
            p1897 = .BufPos
            e1897 = .LastExpected
            If Not (ParseWhitespace()) Then
                .BufPos = p1897
                .LastExpected = e1897
                pvPushThunk ucsActResultClear
                If ParseStartList() Then
                    pvPushThunk ucsActVarSet, 1
                    For i1917 = 0 To LNG_MAXINT
                        p1916 = .BufPos
                        q1916 = .ThunkPos
                        p1904 = .BufPos
                        q1904 = .ThunkPos
                        If .BufData(.BufPos) = 95 Then ' "_"
                            .BufPos = p1916
                            .ThunkPos = q1916
                            pvPushThunk ucsActResultClear
                            If ParseStrongUl() Then
                                pvPushThunk ucsActVarSet, 2
                                pvPushThunk ucsAct_2_EmphUl, .CaptureBegin, .CaptureEnd
                            Else
                                .BufPos = p1916
                                .ThunkPos = q1916
                                .BufPos = p1916
                                .ThunkPos = q1916
                                Exit For
                            End If
                        Else
                            .BufPos = p1904
                            .ThunkPos = q1904
                            pvPushThunk ucsActResultClear
                            If ParseInline() Then
                                pvPushThunk ucsActVarSet, 2
                                pvPushThunk ucsAct_1_EmphUl, .CaptureBegin, .CaptureEnd
                            Else
                                .BufPos = p1916
                                .ThunkPos = q1916
                                pvPushThunk ucsActResultClear
                                If ParseStrongUl() Then
                                    pvPushThunk ucsActVarSet, 2
                                    pvPushThunk ucsAct_2_EmphUl, .CaptureBegin, .CaptureEnd
                                Else
                                    .BufPos = p1916
                                    .ThunkPos = q1916
                                    .BufPos = p1916
                                    .ThunkPos = q1916
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                    If i1917 <> 0 Then
                        If .BufData(.BufPos) = 95 Then ' "_"
                            .BufPos = .BufPos + 1
                            pvPushThunk ucsAct_3_EmphUl, .CaptureBegin, .CaptureEnd
                            pvPushThunk ucsActVarAlloc, -2
                            Call pvSetAdvance
                            ParseEmphUl = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseWhitespace() As Boolean
    Dim p1869 As Long

    With ctx
        p1869 = .BufPos
        If ParseSpacechar() Then
            ParseWhitespace = True
            Exit Function
        Else
            .BufPos = p1869
        End If
        If ParseNewline() Then
            ParseWhitespace = True
            Exit Function
        Else
            .BufPos = p1869
        End If
    End With
End Function

Private Function ParseStrongStar() As Boolean
    Dim p1925 As Long
    Dim e1925 As String
    Dim i1939 As Long
    Dim p1937 As Long
    Dim q1937 As Long
    Dim p1932 As Long
    Dim q1932 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 42 Then ' "**"
            .BufPos = .BufPos + 2
            p1925 = .BufPos
            e1925 = .LastExpected
            If Not (ParseWhitespace()) Then
                .BufPos = p1925
                .LastExpected = e1925
                pvPushThunk ucsActResultClear
                If ParseStartList() Then
                    pvPushThunk ucsActVarSet, 1
                    For i1939 = 0 To LNG_MAXINT
                        p1937 = .BufPos
                        q1937 = .ThunkPos
                        p1932 = .BufPos
                        q1932 = .ThunkPos
                        If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 42 Then ' "**"
                            .BufPos = p1937
                            .ThunkPos = q1937
                            Exit For
                        Else
                            .BufPos = p1932
                            .ThunkPos = q1932
                        End If
                        pvPushThunk ucsActResultClear
                        If ParseInline() Then
                            pvPushThunk ucsActVarSet, 2
                        Else
                            .BufPos = p1937
                            .ThunkPos = q1937
                            Exit For
                        End If
                        pvPushThunk ucsAct_1_StrongStar, .CaptureBegin, .CaptureEnd
                    Next
                    If i1939 <> 0 Then
                        If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 42 Then ' "**"
                            .BufPos = .BufPos + 2
                            pvPushThunk ucsAct_2_StrongStar, .CaptureBegin, .CaptureEnd
                            pvPushThunk ucsActVarAlloc, -2
                            Call pvSetAdvance
                            ParseStrongStar = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseStrongUl() As Boolean
    Dim p1944 As Long
    Dim e1944 As String
    Dim i1958 As Long
    Dim p1956 As Long
    Dim q1956 As Long
    Dim p1951 As Long
    Dim q1951 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        If .BufData(.BufPos) = 95 And .BufData(.BufPos + 1) = 95 Then ' "__"
            .BufPos = .BufPos + 2
            p1944 = .BufPos
            e1944 = .LastExpected
            If Not (ParseWhitespace()) Then
                .BufPos = p1944
                .LastExpected = e1944
                pvPushThunk ucsActResultClear
                If ParseStartList() Then
                    pvPushThunk ucsActVarSet, 1
                    For i1958 = 0 To LNG_MAXINT
                        p1956 = .BufPos
                        q1956 = .ThunkPos
                        p1951 = .BufPos
                        q1951 = .ThunkPos
                        If .BufData(.BufPos) = 95 And .BufData(.BufPos + 1) = 95 Then ' "__"
                            .BufPos = p1956
                            .ThunkPos = q1956
                            Exit For
                        Else
                            .BufPos = p1951
                            .ThunkPos = q1951
                        End If
                        pvPushThunk ucsActResultClear
                        If ParseInline() Then
                            pvPushThunk ucsActVarSet, 2
                        Else
                            .BufPos = p1956
                            .ThunkPos = q1956
                            Exit For
                        End If
                        pvPushThunk ucsAct_1_StrongUl, .CaptureBegin, .CaptureEnd
                    Next
                    If i1958 <> 0 Then
                        If .BufData(.BufPos) = 95 And .BufData(.BufPos + 1) = 95 Then ' "__"
                            .BufPos = .BufPos + 2
                            pvPushThunk ucsAct_2_StrongUl, .CaptureBegin, .CaptureEnd
                            pvPushThunk ucsActVarAlloc, -2
                            Call pvSetAdvance
                            ParseStrongUl = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseExplicitLink() As Boolean
    With ctx
        pvPushThunk ucsActVarAlloc, 3
        pvPushThunk ucsActResultClear
        If ParseLabel() Then
            pvPushThunk ucsActVarSet, 1
            If .BufData(.BufPos) = 40 Then          ' "("
                .BufPos = .BufPos + 1
                Call ParseSp
                pvPushThunk ucsActResultClear
                Call ParseSource
                pvPushThunk ucsActVarSet, 2
                Call ParseSpnl
                pvPushThunk ucsActResultClear
                Call ParseTitle
                pvPushThunk ucsActVarSet, 3
                Call ParseSp
                If .BufData(.BufPos) = 41 Then      ' ")"
                    .BufPos = .BufPos + 1
                    pvPushThunk ucsAct_1_ExplicitLink, .CaptureBegin, .CaptureEnd
                    pvPushThunk ucsActVarAlloc, -3
                    Call pvSetAdvance
                    ParseExplicitLink = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseReferenceLink() As Boolean
    Dim p1998 As Long
    Dim q1998 As Long

    With ctx
        p1998 = .BufPos
        q1998 = .ThunkPos
        If ParseReferenceLinkDouble() Then
            ParseReferenceLink = True
            Exit Function
        Else
            .BufPos = p1998
            .ThunkPos = q1998
        End If
        If ParseReferenceLinkSingle() Then
            ParseReferenceLink = True
            Exit Function
        Else
            .BufPos = p1998
            .ThunkPos = q1998
        End If
    End With
End Function

Private Function ParseAutoLink() As Boolean
    Dim p2120 As Long
    Dim q2120 As Long

    With ctx
        p2120 = .BufPos
        q2120 = .ThunkPos
        If ParseAutoLinkUrl() Then
            ParseAutoLink = True
            Exit Function
        Else
            .BufPos = p2120
            .ThunkPos = q2120
        End If
        If ParseAutoLinkEmail() Then
            ParseAutoLink = True
            Exit Function
        Else
            .BufPos = p2120
            .ThunkPos = q2120
        End If
    End With
End Function

Private Function ParseReferenceLinkDouble() As Boolean
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long
    Dim p2008 As Long
    Dim q2008 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        If ParseLabel() Then
            pvPushThunk ucsActVarSet, 1
            lCaptureBegin = .BufPos
            Call ParseSpnl
            lCaptureEnd = .BufPos
            p2008 = .BufPos
            q2008 = .ThunkPos
            If Not (.BufData(.BufPos) = 91 And .BufData(.BufPos + 1) = 93) Then ' "[]"
                .BufPos = p2008
                .ThunkPos = q2008
                pvPushThunk ucsActResultClear
                If ParseLabel() Then
                    pvPushThunk ucsActVarSet, 2
                    .CaptureBegin = lCaptureBegin
                    .CaptureEnd = lCaptureEnd
                    pvPushThunk ucsAct_1_ReferenceLinkDouble, lCaptureBegin, lCaptureEnd
                    pvPushThunk ucsActVarAlloc, -2
                    Call pvSetAdvance
                    ParseReferenceLinkDouble = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseReferenceLinkSingle() As Boolean
    Dim lCaptureBegin As Long
    Dim p2021 As Long
    Dim q2021 As Long
    Dim lCaptureEnd As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        pvPushThunk ucsActResultClear
        If ParseLabel() Then
            pvPushThunk ucsActVarSet, 1
            lCaptureBegin = .BufPos
            p2021 = .BufPos
            q2021 = .ThunkPos
            Call ParseSpnl
            If .BufData(.BufPos) = 91 And .BufData(.BufPos + 1) = 93 Then ' "[]"
                .BufPos = .BufPos + 2
            Else
                .BufPos = p2021
                .ThunkPos = q2021
            End If
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_ReferenceLinkSingle, lCaptureBegin, lCaptureEnd
            pvPushThunk ucsActVarAlloc, -1
            Call pvSetAdvance
            ParseReferenceLinkSingle = True
        End If
    End With
End Function

Private Function ParseLabel() As Boolean
    Dim p2187 As Long
    Dim p2180 As Long
    Dim p2196 As Long
    Dim q2196 As Long
    Dim p2194 As Long
    Dim q2194 As Long
    Dim p2184 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        If .BufData(.BufPos) = 91 Then              ' "["
            .BufPos = .BufPos + 1
            p2187 = .BufPos
            p2180 = .BufPos
            If .BufData(.BufPos) = 94 Then          ' "^"
                .BufPos = p2187
            Else
                .BufPos = p2180
                If extension(EXT_NOTES) Then
                    pvPushThunk ucsActResultClear
                    If ParseStartList() Then
                        pvPushThunk ucsActVarSet, 1
                        Do
                            p2196 = .BufPos
                            q2196 = .ThunkPos
                            p2194 = .BufPos
                            q2194 = .ThunkPos
                            If .BufData(.BufPos) = 93 Then ' "]"
                                .BufPos = p2196
                                .ThunkPos = q2196
                                Exit Do
                            Else
                                .BufPos = p2194
                                .ThunkPos = q2194
                            End If
                            If Not (ParseInline()) Then
                                .BufPos = p2196
                                .ThunkPos = q2196
                                Exit Do
                            End If
                            pvPushThunk ucsAct_1_Label, .CaptureBegin, .CaptureEnd
                        Loop
                        If .BufData(.BufPos) = 93 Then ' "]"
                            .BufPos = .BufPos + 1
                            pvPushThunk ucsAct_2_Label, .CaptureBegin, .CaptureEnd
                            pvPushThunk ucsActVarAlloc, -1
                            Call pvSetAdvance
                            ParseLabel = True
                            Exit Function
                        End If
                    End If
                Else
                    .BufPos = p2187
                End If
            End If
            p2184 = .BufPos
            If .BufPos < .BufSize Then
                .BufPos = p2184
                If Not extension(EXT_NOTES) Then
                    pvPushThunk ucsActResultClear
                    If ParseStartList() Then
                        pvPushThunk ucsActVarSet, 1
                        Do
                            p2196 = .BufPos
                            q2196 = .ThunkPos
                            p2194 = .BufPos
                            q2194 = .ThunkPos
                            If .BufData(.BufPos) = 93 Then ' "]"
                                .BufPos = p2196
                                .ThunkPos = q2196
                                Exit Do
                            Else
                                .BufPos = p2194
                                .ThunkPos = q2194
                            End If
                            If Not (ParseInline()) Then
                                .BufPos = p2196
                                .ThunkPos = q2196
                                Exit Do
                            End If
                            pvPushThunk ucsAct_1_Label, .CaptureBegin, .CaptureEnd
                        Loop
                        If .BufData(.BufPos) = 93 Then ' "]"
                            .BufPos = .BufPos + 1
                            pvPushThunk ucsAct_2_Label, .CaptureBegin, .CaptureEnd
                            pvPushThunk ucsActVarAlloc, -1
                            Call pvSetAdvance
                            ParseLabel = True
                            Exit Function
                        End If
                    End If
                Else
                    .BufPos = p2187
                End If
            Else
                .BufPos = p2187
            End If
        End If
    End With
End Function

Private Sub ParseSource()
    Dim p2056 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        p2056 = .BufPos
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            lCaptureBegin = .BufPos
            Call ParseSourceContents
            lCaptureEnd = .BufPos
            If .BufData(.BufPos) = 62 Then          ' ">"
                .BufPos = .BufPos + 1
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                pvPushThunk ucsAct_1_Source, lCaptureBegin, lCaptureEnd
                Call pvSetAdvance
            Else
                .BufPos = p2056
            End If
        End If
        lCaptureBegin = .BufPos
        Call ParseSourceContents
        lCaptureEnd = .BufPos
        .CaptureBegin = lCaptureBegin
        .CaptureEnd = lCaptureEnd
        pvPushThunk ucsAct_1_Source, lCaptureBegin, lCaptureEnd
        Call pvSetAdvance
    End With
End Sub

Private Sub ParseTitle()
    Dim p2079 As Long
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        p2079 = .BufPos
        If Not (ParseTitleSingle()) Then
            .BufPos = p2079
            If Not (ParseTitleDouble()) Then
                .BufPos = p2079
                lCaptureBegin = .BufPos
                lCaptureEnd = .BufPos
            End If
        End If
        .CaptureBegin = lCaptureBegin
        .CaptureEnd = lCaptureEnd
        pvPushThunk ucsAct_1_Title, lCaptureBegin, lCaptureEnd
    End With
End Sub

Private Sub ParseSourceContents()
    Dim p2073 As Long
    Dim i2068 As Long
    Dim p2063 As Long
    Dim p2060 As Long
    Dim p2062 As Long
    Dim p2065 As Long

    With ctx
        Do
            p2073 = .BufPos
            For i2068 = 0 To LNG_MAXINT
                p2063 = .BufPos
                p2060 = .BufPos
                If .BufData(.BufPos) = 40 Then      ' "("
                    .BufPos = p2063
                    Exit For
                Else
                    .BufPos = p2060
                End If
                p2062 = .BufPos
                If .BufData(.BufPos) = 41 Then      ' ")"
                    .BufPos = p2063
                    Exit For
                Else
                    .BufPos = p2062
                End If
                p2065 = .BufPos
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = p2063
                    Exit For
                Else
                    .BufPos = p2065
                End If
                If Not (ParseNonspacechar()) Then
                    .BufPos = p2063
                    Exit For
                End If
            Next
            If i2068 <> 0 Then
            Else
                If .BufData(.BufPos) = 40 Then      ' "("
                    .BufPos = .BufPos + 1
                    Call ParseSourceContents
                    If .BufData(.BufPos) = 41 Then  ' ")"
                        .BufPos = .BufPos + 1
                    Else
                        .BufPos = p2073
                        .BufPos = p2073
                        Exit Do
                    End If
                Else
                    .BufPos = p2073
                    Exit Do
                End If
            End If
        Loop
        Call pvSetAdvance
    End With
End Sub

Private Function ParseNonspacechar() As Boolean
    Dim p2540 As Long
    Dim e2540 As String
    Dim p2542 As Long
    Dim e2542 As String

    With ctx
        p2540 = .BufPos
        e2540 = .LastExpected
        If Not (ParseSpacechar()) Then
            .BufPos = p2540
            .LastExpected = e2540
            p2542 = .BufPos
            e2542 = .LastExpected
            If Not (ParseNewline()) Then
                .BufPos = p2542
                .LastExpected = e2542
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                    ParseNonspacechar = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseTitleSingle() As Boolean
    Dim lCaptureBegin As Long
    Dim p2097 As Long
    Dim p2095 As Long
    Dim p2094 As Long
    Dim lCaptureEnd As Long

    With ctx
        If .BufData(.BufPos) = 39 Then              ' "\'"
            .BufPos = .BufPos + 1
            lCaptureBegin = .BufPos
            Do
                p2097 = .BufPos
                p2095 = .BufPos
                If .BufData(.BufPos) = 39 Then      ' "\'"
                    .BufPos = .BufPos + 1
                    Call ParseSp
                    p2094 = .BufPos
                    If .BufData(.BufPos) = 41 Then  ' ")"
                        .BufPos = .BufPos + 1
                        .BufPos = p2097
                        Exit Do
                    End If
                    If ParseNewline() Then
                        .BufPos = p2097
                        Exit Do
                    Else
                        .BufPos = p2094
                    End If
                    .BufPos = p2095
                Else
                    .BufPos = p2095
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p2097
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            If .BufData(.BufPos) = 39 Then          ' "\'"
                .BufPos = .BufPos + 1
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call pvSetAdvance
                ParseTitleSingle = True
            End If
        End If
    End With
End Function

Private Function ParseTitleDouble() As Boolean
    Dim lCaptureBegin As Long
    Dim p2112 As Long
    Dim p2110 As Long
    Dim p2109 As Long
    Dim lCaptureEnd As Long

    With ctx
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            lCaptureBegin = .BufPos
            Do
                p2112 = .BufPos
                p2110 = .BufPos
                If .BufData(.BufPos) = 34 Then      ' """
                    .BufPos = .BufPos + 1
                    Call ParseSp
                    p2109 = .BufPos
                    If .BufData(.BufPos) = 41 Then  ' ")"
                        .BufPos = .BufPos + 1
                        .BufPos = p2112
                        Exit Do
                    End If
                    If ParseNewline() Then
                        .BufPos = p2112
                        Exit Do
                    Else
                        .BufPos = p2109
                    End If
                    .BufPos = p2110
                Else
                    .BufPos = p2110
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p2112
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            If .BufData(.BufPos) = 34 Then          ' """
                .BufPos = .BufPos + 1
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call pvSetAdvance
                ParseTitleDouble = True
            End If
        End If
    End With
End Function

Private Function ParseAutoLinkUrl() As Boolean
    Dim lCaptureBegin As Long
    Dim i2125 As Long
    Dim i2133 As Long
    Dim p2131 As Long
    Dim p2128 As Long
    Dim e2128 As String
    Dim p2130 As Long
    Dim lCaptureEnd As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            lCaptureBegin = .BufPos
            For i2125 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 65 To 90, 97 To 122            ' [A-Za-z]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If i2125 <> 0 Then
                If pvMatchString("://") Then        ' "://"
                    .BufPos = .BufPos + 3
                    For i2133 = 0 To LNG_MAXINT
                        p2131 = .BufPos
                        p2128 = .BufPos
                        e2128 = .LastExpected
                        If ParseNewline() Then
                            .BufPos = p2131
                            Exit For
                        Else
                            .BufPos = p2128
                            .LastExpected = e2128
                        End If
                        p2130 = .BufPos
                        If .BufData(.BufPos) = 62 Then ' ">"
                            .BufPos = p2131
                            Exit For
                        Else
                            .BufPos = p2130
                        End If
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p2131
                            Exit For
                        End If
                    Next
                    If i2133 <> 0 Then
                        lCaptureEnd = .BufPos
                        If .BufData(.BufPos) = 62 Then ' ">"
                            .BufPos = .BufPos + 1
                            .CaptureBegin = lCaptureBegin
                            .CaptureEnd = lCaptureEnd
                            pvPushThunk ucsAct_1_AutoLinkUrl, lCaptureBegin, lCaptureEnd
                            Call pvSetAdvance
                            ParseAutoLinkUrl = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseAutoLinkEmail() As Boolean
    Dim lCaptureBegin As Long
    Dim i2143 As Long
    Dim i2151 As Long
    Dim p2149 As Long
    Dim p2146 As Long
    Dim e2146 As String
    Dim p2148 As Long
    Dim lCaptureEnd As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            If pvMatchString("mailto:") Then        ' "mailto:"
                .BufPos = .BufPos + 7
            End If
            lCaptureBegin = .BufPos
            For i2143 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 45, 65 To 90, 97 To 122, 48 To 57, 43, 95, 46, 47, 33, 37, 126, 36 ' [-A-Za-z0-9+_./!%~$]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If i2143 <> 0 Then
                If .BufData(.BufPos) = 64 Then      ' "@"
                    .BufPos = .BufPos + 1
                    For i2151 = 0 To LNG_MAXINT
                        p2149 = .BufPos
                        p2146 = .BufPos
                        e2146 = .LastExpected
                        If ParseNewline() Then
                            .BufPos = p2149
                            Exit For
                        Else
                            .BufPos = p2146
                            .LastExpected = e2146
                        End If
                        p2148 = .BufPos
                        If .BufData(.BufPos) = 62 Then ' ">"
                            .BufPos = p2149
                            Exit For
                        Else
                            .BufPos = p2148
                        End If
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p2149
                            Exit For
                        End If
                    Next
                    If i2151 <> 0 Then
                        lCaptureEnd = .BufPos
                        If .BufData(.BufPos) = 62 Then ' ">"
                            .BufPos = .BufPos + 1
                            .CaptureBegin = lCaptureBegin
                            .CaptureEnd = lCaptureEnd
                            pvPushThunk ucsAct_1_AutoLinkEmail, lCaptureBegin, lCaptureEnd
                            Call pvSetAdvance
                            ParseAutoLinkEmail = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseRefSrc() As Boolean
    Dim lCaptureBegin As Long
    Dim i2203 As Long
    Dim p2202 As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        For i2203 = 0 To LNG_MAXINT
            p2202 = .BufPos
            If Not (ParseNonspacechar()) Then
                .BufPos = p2202
                Exit For
            End If
        Next
        If i2203 <> 0 Then
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_RefSrc, lCaptureBegin, lCaptureEnd
            ParseRefSrc = True
        End If
    End With
End Function

Private Sub ParseRefTitle()
    Dim p2211 As Long

    With ctx
        p2211 = .BufPos
        If Not (ParseRefTitleSingle()) Then
            .BufPos = p2211
            If Not (ParseRefTitleDouble()) Then
                .BufPos = p2211
                If Not (ParseRefTitleParens()) Then
                    .BufPos = p2211
                    Call ParseEmptyTitle
                End If
            End If
        End If
        pvPushThunk ucsAct_1_RefTitle, .CaptureBegin, .CaptureEnd
    End With
End Sub

Private Function ParseRefTitleSingle() As Boolean
    Dim lCaptureBegin As Long
    Dim p2234 As Long
    Dim p2232 As Long
    Dim p2231 As Long
    Dim lCaptureEnd As Long

    With ctx
        Call ParseSpnl
        If .BufData(.BufPos) = 39 Then              ' "\'"
            .BufPos = .BufPos + 1
            lCaptureBegin = .BufPos
            Do
                p2234 = .BufPos
                p2232 = .BufPos
                p2231 = .BufPos
                If .BufData(.BufPos) = 39 Then      ' "\'"
                    .BufPos = .BufPos + 1
                    Call ParseSp
                    If ParseNewline() Then
                        .BufPos = p2234
                        Exit Do
                    Else
                        .BufPos = p2231
                    End If
                End If
                If ParseNewline() Then
                    .BufPos = p2234
                    Exit Do
                Else
                    .BufPos = p2231
                End If
                .BufPos = p2232
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p2234
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            If .BufData(.BufPos) = 39 Then          ' "\'"
                .BufPos = .BufPos + 1
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call pvSetAdvance
                ParseRefTitleSingle = True
            End If
        End If
    End With
End Function

Private Function ParseRefTitleDouble() As Boolean
    Dim lCaptureBegin As Long
    Dim p2250 As Long
    Dim p2248 As Long
    Dim p2247 As Long
    Dim lCaptureEnd As Long

    With ctx
        Call ParseSpnl
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            lCaptureBegin = .BufPos
            Do
                p2250 = .BufPos
                p2248 = .BufPos
                p2247 = .BufPos
                If .BufData(.BufPos) = 34 Then      ' """
                    .BufPos = .BufPos + 1
                    Call ParseSp
                    If ParseNewline() Then
                        .BufPos = p2250
                        Exit Do
                    Else
                        .BufPos = p2247
                    End If
                End If
                If ParseNewline() Then
                    .BufPos = p2250
                    Exit Do
                Else
                    .BufPos = p2247
                End If
                .BufPos = p2248
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p2250
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            If .BufData(.BufPos) = 34 Then          ' """
                .BufPos = .BufPos + 1
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call pvSetAdvance
                ParseRefTitleDouble = True
            End If
        End If
    End With
End Function

Private Function ParseRefTitleParens() As Boolean
    Dim lCaptureBegin As Long
    Dim p2266 As Long
    Dim p2264 As Long
    Dim p2263 As Long
    Dim lCaptureEnd As Long

    With ctx
        Call ParseSpnl
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            lCaptureBegin = .BufPos
            Do
                p2266 = .BufPos
                p2264 = .BufPos
                p2263 = .BufPos
                If .BufData(.BufPos) = 41 Then      ' ")"
                    .BufPos = .BufPos + 1
                    Call ParseSp
                    If ParseNewline() Then
                        .BufPos = p2266
                        Exit Do
                    Else
                        .BufPos = p2263
                    End If
                End If
                If ParseNewline() Then
                    .BufPos = p2266
                    Exit Do
                Else
                    .BufPos = p2263
                End If
                .BufPos = p2264
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p2266
                    Exit Do
                End If
            Loop
            lCaptureEnd = .BufPos
            If .BufData(.BufPos) = 41 Then          ' ")"
                .BufPos = .BufPos + 1
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                Call pvSetAdvance
                ParseRefTitleParens = True
            End If
        End If
    End With
End Function

Private Sub ParseEmptyTitle()
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        lCaptureEnd = .BufPos
        .CaptureBegin = lCaptureBegin
        .CaptureEnd = lCaptureEnd
    End With
End Sub

Private Function ParseSkipBlock() As Boolean
    Dim p2798 As Long
    Dim q2798 As Long
    Dim i2794 As Long
    Dim p2788 As Long
    Dim q2788 As Long
    Dim p2785 As Long
    Dim q2785 As Long
    Dim p2787 As Long
    Dim q2787 As Long
    Dim e2787 As String
    Dim p2790 As Long
    Dim q2790 As Long
    Dim e2790 As String
    Dim p2792 As Long
    Dim q2792 As Long
    Dim e2792 As String
    Dim p2795 As Long
    Dim q2795 As Long
    Dim i2800 As Long
    Dim p2799 As Long
    Dim q2799 As Long

    With ctx
        p2798 = .BufPos
        q2798 = .ThunkPos
        If ParseHtmlBlock() Then
            Call pvSetAdvance
            ParseSkipBlock = True
            Exit Function
        Else
            .BufPos = p2798
            .ThunkPos = q2798
        End If
        For i2794 = 0 To LNG_MAXINT
            p2788 = .BufPos
            q2788 = .ThunkPos
            p2785 = .BufPos
            q2785 = .ThunkPos
            If .BufData(.BufPos) = 35 Then          ' "#"
                .BufPos = p2788
                .ThunkPos = q2788
                Exit For
            Else
                .BufPos = p2785
                .ThunkPos = q2785
            End If
            p2787 = .BufPos
            q2787 = .ThunkPos
            e2787 = .LastExpected
            If ParseSetextBottom1() Then
                .BufPos = p2788
                .ThunkPos = q2788
                Exit For
            Else
                .BufPos = p2787
                .ThunkPos = q2787
                .LastExpected = e2787
            End If
            p2790 = .BufPos
            q2790 = .ThunkPos
            e2790 = .LastExpected
            If ParseSetextBottom2() Then
                .BufPos = p2788
                .ThunkPos = q2788
                Exit For
            Else
                .BufPos = p2790
                .ThunkPos = q2790
                .LastExpected = e2790
            End If
            p2792 = .BufPos
            q2792 = .ThunkPos
            e2792 = .LastExpected
            If ParseBlankLine() Then
                .BufPos = p2788
                .ThunkPos = q2788
                Exit For
            Else
                .BufPos = p2792
                .ThunkPos = q2792
                .LastExpected = e2792
            End If
            If Not (ParseRawLine()) Then
                .BufPos = p2788
                .ThunkPos = q2788
                Exit For
            End If
        Next
        If i2794 <> 0 Then
            Do
                p2795 = .BufPos
                q2795 = .ThunkPos
                If Not (ParseBlankLine()) Then
                    .BufPos = p2795
                    .ThunkPos = q2795
                    Exit Do
                End If
            Loop
            Call pvSetAdvance
            ParseSkipBlock = True
            Exit Function
        Else
            .BufPos = p2798
            .ThunkPos = q2798
        End If
        For i2800 = 0 To LNG_MAXINT
            p2799 = .BufPos
            q2799 = .ThunkPos
            If Not (ParseBlankLine()) Then
                .BufPos = p2799
                .ThunkPos = q2799
                Exit For
            End If
        Next
        If i2800 <> 0 Then
            Call pvSetAdvance
            ParseSkipBlock = True
            Exit Function
        End If
        If ParseRawLine() Then
            Call pvSetAdvance
            ParseSkipBlock = True
            Exit Function
        Else
            .BufPos = p2798
            .ThunkPos = q2798
        End If
    End With
End Function

Private Function ParseTicks1() As Boolean
    Dim p2288 As Long

    With ctx
        If .BufData(.BufPos) = 96 Then              ' "`"
            .BufPos = .BufPos + 1
            p2288 = .BufPos
            If Not (.BufData(.BufPos) = 96) Then    ' "`"
                .BufPos = p2288
                Call pvSetAdvance
                ParseTicks1 = True
            End If
        End If
    End With
End Function

Private Function ParseTicks2() As Boolean
    Dim p2293 As Long

    With ctx
        If .BufData(.BufPos) = 96 And .BufData(.BufPos + 1) = 96 Then ' "``"
            .BufPos = .BufPos + 2
            p2293 = .BufPos
            If Not (.BufData(.BufPos) = 96) Then    ' "`"
                .BufPos = p2293
                Call pvSetAdvance
                ParseTicks2 = True
            End If
        End If
    End With
End Function

Private Function ParseTicks3() As Boolean
    Dim p2298 As Long

    With ctx
        If pvMatchString("```") Then                ' "```"
            .BufPos = .BufPos + 3
            p2298 = .BufPos
            If Not (.BufData(.BufPos) = 96) Then    ' "`"
                .BufPos = p2298
                Call pvSetAdvance
                ParseTicks3 = True
            End If
        End If
    End With
End Function

Private Function ParseTicks4() As Boolean
    Dim p2303 As Long

    With ctx
        If pvMatchString("````") Then               ' "````"
            .BufPos = .BufPos + 4
            p2303 = .BufPos
            If Not (.BufData(.BufPos) = 96) Then    ' "`"
                .BufPos = p2303
                Call pvSetAdvance
                ParseTicks4 = True
            End If
        End If
    End With
End Function

Private Function ParseTicks5() As Boolean
    Dim p2308 As Long

    With ctx
        If pvMatchString("`````") Then              ' "`````"
            .BufPos = .BufPos + 5
            p2308 = .BufPos
            If Not (.BufData(.BufPos) = 96) Then    ' "`"
                .BufPos = p2308
                Call pvSetAdvance
                ParseTicks5 = True
            End If
        End If
    End With
End Function

Private Function ParseHtmlTag() As Boolean
    Dim i2526 As Long
    Dim p2528 As Long

    With ctx
        If .BufData(.BufPos) = 60 Then              ' "<"
            .BufPos = .BufPos + 1
            Call ParseSpnl
            If .BufData(.BufPos) = 47 Then          ' "/"
                .BufPos = .BufPos + 1
            End If
            For i2526 = 0 To LNG_MAXINT
                Select Case .BufData(.BufPos)
                Case 65 To 90, 97 To 122, 48 To 57  ' [A-Za-z0-9]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit For
                End Select
            Next
            If i2526 <> 0 Then
                Call ParseSpnl
                Do
                    p2528 = .BufPos
                    If Not (ParseHtmlAttribute()) Then
                        .BufPos = p2528
                        Exit Do
                    End If
                Loop
                If .BufData(.BufPos) = 47 Then      ' "/"
                    .BufPos = .BufPos + 1
                End If
                Call ParseSpnl
                If .BufData(.BufPos) = 62 Then      ' ">"
                    .BufPos = .BufPos + 1
                    Call pvSetAdvance
                    ParseHtmlTag = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseQuoted() As Boolean
    Dim p2492 As Long
    Dim p2480 As Long
    Dim p2478 As Long
    Dim p2488 As Long
    Dim p2486 As Long

    With ctx
        p2492 = .BufPos
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            Do
                p2480 = .BufPos
                p2478 = .BufPos
                If .BufData(.BufPos) = 34 Then      ' """
                    .BufPos = p2480
                    Exit Do
                Else
                    .BufPos = p2478
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p2480
                    Exit Do
                End If
            Loop
            If .BufData(.BufPos) = 34 Then          ' """
                .BufPos = .BufPos + 1
                Call pvSetAdvance
                ParseQuoted = True
                Exit Function
            Else
                .BufPos = p2492
            End If
        End If
        If .BufData(.BufPos) = 39 Then              ' "\'"
            .BufPos = .BufPos + 1
            Do
                p2488 = .BufPos
                p2486 = .BufPos
                If .BufData(.BufPos) = 39 Then      ' "\'"
                    .BufPos = p2488
                    Exit Do
                Else
                    .BufPos = p2486
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p2488
                    Exit Do
                End If
            Loop
            If .BufData(.BufPos) = 39 Then          ' "\'"
                .BufPos = .BufPos + 1
                Call pvSetAdvance
                ParseQuoted = True
                Exit Function
            Else
                .BufPos = p2492
            End If
        End If
    End With
End Function

Private Function ParseExtendedSpecialChar() As Boolean
    Dim p2812 As Long

    With ctx
        p2812 = .BufPos
        If extension(EXT_SMART) Then
            If .BufData(.BufPos) = 46 Then          ' "."
                .BufPos = .BufPos + 1
                Call pvSetAdvance
                ParseExtendedSpecialChar = True
                Exit Function
            End If
            If .BufData(.BufPos) = 45 Then          ' "-"
                .BufPos = .BufPos + 1
                Call pvSetAdvance
                ParseExtendedSpecialChar = True
                Exit Function
            End If
            If .BufData(.BufPos) = 39 Then          ' "\'"
                .BufPos = .BufPos + 1
                Call pvSetAdvance
                ParseExtendedSpecialChar = True
                Exit Function
            End If
            If .BufData(.BufPos) = 34 Then          ' """
                .BufPos = .BufPos + 1
                Call pvSetAdvance
                ParseExtendedSpecialChar = True
                Exit Function
            End If
            .BufPos = p2812
        End If
        If extension(EXT_NOTES) Then
            If .BufData(.BufPos) = 94 Then          ' "^"
                .BufPos = .BufPos + 1
                Call pvSetAdvance
                ParseExtendedSpecialChar = True
            End If
        End If
    End With
End Function

Private Function ParseEllipsis() As Boolean
    With ctx
        If pvMatchString("...") Then                ' "..."
            .BufPos = .BufPos + 3
            pvPushThunk ucsAct_1_Ellipsis, .CaptureBegin, .CaptureEnd
            Call pvSetAdvance
            ParseEllipsis = True
            Exit Function
        End If
        If pvMatchString(". . .") Then              ' ". . ."
            .BufPos = .BufPos + 5
            pvPushThunk ucsAct_1_Ellipsis, .CaptureBegin, .CaptureEnd
            Call pvSetAdvance
            ParseEllipsis = True
        End If
    End With
End Function

Private Function ParseDash() As Boolean
    Dim p2838 As Long
    Dim q2838 As Long

    With ctx
        p2838 = .BufPos
        q2838 = .ThunkPos
        If ParseEmDash() Then
            ParseDash = True
            Exit Function
        Else
            .BufPos = p2838
            .ThunkPos = q2838
        End If
        If ParseEnDash() Then
            ParseDash = True
            Exit Function
        Else
            .BufPos = p2838
            .ThunkPos = q2838
        End If
    End With
End Function

Private Function ParseSingleQuoted() As Boolean
    Dim i2875 As Long
    Dim p2873 As Long
    Dim q2873 As Long
    Dim p2868 As Long
    Dim q2868 As Long
    Dim e2868 As String

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        If ParseSingleQuoteStart() Then
            pvPushThunk ucsActResultClear
            If ParseStartList() Then
                pvPushThunk ucsActVarSet, 1
                For i2875 = 0 To LNG_MAXINT
                    p2873 = .BufPos
                    q2873 = .ThunkPos
                    p2868 = .BufPos
                    q2868 = .ThunkPos
                    e2868 = .LastExpected
                    If ParseSingleQuoteEnd() Then
                        .BufPos = p2873
                        .ThunkPos = q2873
                        Exit For
                    Else
                        .BufPos = p2868
                        .ThunkPos = q2868
                        .LastExpected = e2868
                    End If
                    pvPushThunk ucsActResultClear
                    If ParseInline() Then
                        pvPushThunk ucsActVarSet, 2
                    Else
                        .BufPos = p2873
                        .ThunkPos = q2873
                        Exit For
                    End If
                    pvPushThunk ucsAct_1_SingleQuoted, .CaptureBegin, .CaptureEnd
                Next
                If i2875 <> 0 Then
                    If ParseSingleQuoteEnd() Then
                        pvPushThunk ucsAct_2_SingleQuoted, .CaptureBegin, .CaptureEnd
                        pvPushThunk ucsActVarAlloc, -2
                        ParseSingleQuoted = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseDoubleQuoted() As Boolean
    Dim i2896 As Long
    Dim p2894 As Long
    Dim q2894 As Long
    Dim p2889 As Long
    Dim q2889 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            pvPushThunk ucsActResultClear
            If ParseStartList() Then
                pvPushThunk ucsActVarSet, 1
                For i2896 = 0 To LNG_MAXINT
                    p2894 = .BufPos
                    q2894 = .ThunkPos
                    p2889 = .BufPos
                    q2889 = .ThunkPos
                    If .BufData(.BufPos) = 34 Then  ' """
                        .BufPos = p2894
                        .ThunkPos = q2894
                        Exit For
                    Else
                        .BufPos = p2889
                        .ThunkPos = q2889
                    End If
                    pvPushThunk ucsActResultClear
                    If ParseInline() Then
                        pvPushThunk ucsActVarSet, 2
                    Else
                        .BufPos = p2894
                        .ThunkPos = q2894
                        Exit For
                    End If
                    pvPushThunk ucsAct_1_DoubleQuoted, .CaptureBegin, .CaptureEnd
                Next
                If i2896 <> 0 Then
                    If .BufData(.BufPos) = 34 Then  ' """
                        .BufPos = .BufPos + 1
                        pvPushThunk ucsAct_2_DoubleQuoted, .CaptureBegin, .CaptureEnd
                        pvPushThunk ucsActVarAlloc, -2
                        Call pvSetAdvance
                        ParseDoubleQuoted = True
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function ParseApostrophe() As Boolean
    With ctx
        If .BufData(.BufPos) = 39 Then              ' "\'"
            .BufPos = .BufPos + 1
            pvPushThunk ucsAct_1_Apostrophe, .CaptureBegin, .CaptureEnd
            Call pvSetAdvance
            ParseApostrophe = True
        End If
    End With
End Function

Private Function ParseEmDash() As Boolean
    With ctx
        If pvMatchString("---") Then                ' "---"
            .BufPos = .BufPos + 3
            pvPushThunk ucsAct_1_EmDash, .CaptureBegin, .CaptureEnd
            Call pvSetAdvance
            ParseEmDash = True
            Exit Function
        End If
        If .BufData(.BufPos) = 45 And .BufData(.BufPos + 1) = 45 Then ' "--"
            .BufPos = .BufPos + 2
            pvPushThunk ucsAct_1_EmDash, .CaptureBegin, .CaptureEnd
            Call pvSetAdvance
            ParseEmDash = True
        End If
    End With
End Function

Private Function ParseEnDash() As Boolean
    Dim p2841 As Long

    With ctx
        If .BufData(.BufPos) = 45 Then              ' "-"
            .BufPos = .BufPos + 1
            p2841 = .BufPos
            Select Case .BufData(.BufPos)
            Case 48 To 57                           ' [0-9]
                .BufPos = p2841
                pvPushThunk ucsAct_1_EnDash, .CaptureBegin, .CaptureEnd
                Call pvSetAdvance
                ParseEnDash = True
            End Select
        End If
    End With
End Function

Private Function ParseSingleQuoteStart() As Boolean
    Dim p2854 As Long
    Dim p2853 As Long

    With ctx
        If .BufData(.BufPos) = 39 Then              ' "\'"
            .BufPos = .BufPos + 1
            p2854 = .BufPos
            p2853 = .BufPos
            If Not (ParseSpacechar()) Then
                .BufPos = p2853
                If Not (ParseNewline()) Then
                    .BufPos = p2853
                    .BufPos = p2854
                    Call pvSetAdvance
                    ParseSingleQuoteStart = True
                    Exit Function
                End If
            End If
        End If
    End With
End Function

Private Function ParseSingleQuoteEnd() As Boolean
    Dim p2859 As Long
    Dim e2859 As String

    With ctx
        If .BufData(.BufPos) = 39 Then              ' "\'"
            .BufPos = .BufPos + 1
            p2859 = .BufPos
            e2859 = .LastExpected
            If Not (ParseAlphanumeric()) Then
                .BufPos = p2859
                .LastExpected = e2859
                Call pvSetAdvance
                ParseSingleQuoteEnd = True
            End If
        End If
    End With
End Function

Private Function ParseRawNoteReference() As Boolean
    Dim lCaptureBegin As Long
    Dim i2916 As Long
    Dim p2914 As Long
    Dim p2911 As Long
    Dim e2911 As String
    Dim p2913 As Long
    Dim lCaptureEnd As Long

    With ctx
        If .BufData(.BufPos) = 91 And .BufData(.BufPos + 1) = 94 Then ' "[^"
            .BufPos = .BufPos + 2
            lCaptureBegin = .BufPos
            For i2916 = 0 To LNG_MAXINT
                p2914 = .BufPos
                p2911 = .BufPos
                e2911 = .LastExpected
                If ParseNewline() Then
                    .BufPos = p2914
                    Exit For
                Else
                    .BufPos = p2911
                    .LastExpected = e2911
                End If
                p2913 = .BufPos
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = p2914
                    Exit For
                Else
                    .BufPos = p2913
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p2914
                    Exit For
                End If
            Next
            If i2916 <> 0 Then
                lCaptureEnd = .BufPos
                If .BufData(.BufPos) = 93 Then      ' "]"
                    .BufPos = .BufPos + 1
                    .CaptureBegin = lCaptureBegin
                    .CaptureEnd = lCaptureEnd
                    pvPushThunk ucsAct_1_RawNoteReference, lCaptureBegin, lCaptureEnd
                    Call pvSetAdvance
                    ParseRawNoteReference = True
                End If
            End If
        End If
    End With
End Function

Private Function ParseRawNoteBlock() As Boolean
    Dim i2982 As Long
    Dim p2980 As Long
    Dim q2980 As Long
    Dim p2978 As Long
    Dim q2978 As Long
    Dim e2978 As String
    Dim lCaptureBegin As Long
    Dim p2984 As Long
    Dim q2984 As Long
    Dim lCaptureEnd As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 1
        pvPushThunk ucsActResultClear
        If ParseStartList() Then
            pvPushThunk ucsActVarSet, 1
            For i2982 = 0 To LNG_MAXINT
                p2980 = .BufPos
                q2980 = .ThunkPos
                p2978 = .BufPos
                q2978 = .ThunkPos
                e2978 = .LastExpected
                If ParseBlankLine() Then
                    .BufPos = p2980
                    .ThunkPos = q2980
                    Exit For
                Else
                    .BufPos = p2978
                    .ThunkPos = q2978
                    .LastExpected = e2978
                End If
                If Not (ParseOptionallyIndentedLine()) Then
                    .BufPos = p2980
                    .ThunkPos = q2980
                    Exit For
                End If
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                pvPushThunk ucsAct_1_RawNoteBlock, lCaptureBegin, lCaptureEnd
            Next
            If i2982 <> 0 Then
                lCaptureBegin = .BufPos
                Do
                    p2984 = .BufPos
                    q2984 = .ThunkPos
                    If Not (ParseBlankLine()) Then
                        .BufPos = p2984
                        .ThunkPos = q2984
                        Exit Do
                    End If
                Loop
                lCaptureEnd = .BufPos
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                pvPushThunk ucsAct_2_RawNoteBlock, lCaptureBegin, lCaptureEnd
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                pvPushThunk ucsAct_3_RawNoteBlock, lCaptureBegin, lCaptureEnd
                pvPushThunk ucsActVarAlloc, -1
                ParseRawNoteBlock = True
            End If
        End If
    End With
End Function

Private Sub pvImplAction(ByVal eAction As UcsParserActionsEnum, ByVal lOffset As Long, ByVal lSize As Long)
    With ctx
        Select Case eAction
        Case ucsAct_2_Doc
             ' parse_result = reverse(a);
        Case ucsAct_1_Doc
             ' a = cons(.VarResult, a);
        Case ucsAct_1_StartList
             ' .VarResult = NULL;
        Case ucsAct_1_BlockQuote
             '  .VarResult = mk_element(BLOCKQUOTE);
               '  .VarResult->children = a;

        Case ucsAct_2_Verbatim
             ' .VarResult = mk_str_from_list(a, false);
                 ' .VarResult->key = VERBATIM;
        Case ucsAct_1_Verbatim
             ' a = cons(.VarResult, a);
        Case ucsAct_3_Note
             '   .VarResult = mk_list(NOTE, a);
                  '   .VarResult->contents.str = strdup(ref->contents.str);

        Case ucsAct_2_Note
             ' a = cons(.VarResult, a);
        Case ucsAct_1_Note
             ' a = cons(.VarResult, a);
        Case ucsAct_1_Reference
             ' .VarResult = mk_link(l->children, s->contents.str, t->contents.str);
              ' free_element(s);
              ' free_element(t);
              ' free(l);
              ' .VarResult->key = REFERENCE;
        Case ucsAct_1_HorizontalRule
             ' .VarResult = mk_element(HRULE);
        Case ucsAct_1_OrderedList
             ' .VarResult->key = ORDEREDLIST;
        Case ucsAct_1_BulletList
             ' .VarResult->key = BULLETLIST;
        Case ucsAct_1_HtmlBlock
             '   if (extension(EXT_FILTER_HTML)) { '
              '       .VarResult = mk_list(LIST, NULL);
              '   } else { '
              '       .VarResult = mk_str(yytext);
              '       .VarResult->key = HTMLBLOCK;
              '   }

        Case ucsAct_1_StyleBlock
             '   if (extension(EXT_FILTER_STYLES)) { '
                  '       .VarResult = mk_list(LIST, NULL);
                  '   } else { '
                  '       .VarResult = mk_str(yytext);
                  '       .VarResult->key = HTMLBLOCK;
                  '   }

        Case ucsAct_1_Para
             ' .VarResult = a; .VarResult->key = PARA;
        Case ucsAct_1_Plain
             ' .VarResult = a; .VarResult->key = PLAIN;
        Case ucsAct_3_Inlines
             ' .VarResult = mk_list(LIST, a);
        Case ucsAct_2_Inlines
             ' a = cons(c, a);
        Case ucsAct_1_Inlines
             ' a = cons(.VarResult, a);
        Case ucsAct_1_AtxStart
             ' .VarResult = mk_element(H1 + (strlen(yytext) - 1));
        Case ucsAct_2_AtxHeading
             ' .VarResult = mk_list(s->key, a);
              ' free(s);
        Case ucsAct_1_AtxHeading
             ' a = cons(.VarResult, a);
        Case ucsAct_2_SetextHeading1
             ' .VarResult = mk_list(H1, a);
        Case ucsAct_1_SetextHeading1
             ' a = cons(.VarResult, a);
        Case ucsAct_2_SetextHeading2
             ' .VarResult = mk_list(H2, a);
        Case ucsAct_1_SetextHeading2
             ' a = cons(.VarResult, a);
        Case ucsAct_4_BlockQuoteRaw
             '   .VarResult = mk_str_from_list(a, true);
                   '   .VarResult->key = RAW;

        Case ucsAct_1_BlockQuoteRaw
             ' a = cons(.VarResult, a);
        Case ucsAct_2_BlockQuoteRaw
             ' a = cons(.VarResult, a);
        Case ucsAct_3_BlockQuoteRaw
             ' a = cons(mk_str("\n"), a);
        Case ucsAct_1_Line
             ' .VarResult = mk_str(yytext);
        Case ucsAct_3_VerbatimChunk
             ' .VarResult = mk_str_from_list(a, false);
        Case ucsAct_2_VerbatimChunk
             ' a = cons(.VarResult, a);
        Case ucsAct_1_VerbatimChunk
             ' a = cons(mk_str("\n"), a);
        Case ucsAct_2_ListTight
             ' .VarResult = mk_list(LIST, a);
        Case ucsAct_1_ListTight
             ' a = cons(.VarResult, a);
        Case ucsAct_2_ListLoose
             ' .VarResult = mk_list(LIST, a);
        Case ucsAct_1_ListLoose
             '   element *li;
                '   li = b->children;
                '   li->contents.str = realloc(li->contents.str, strlen(li->contents.str) + 3);
                '   strcat(li->contents.str, "\n\n");  /* In loose list, \n\n added to end of each element */
                '   a = cons(b, a);

        Case ucsAct_3_ListItemTight
             '  element *raw;
              '  raw = mk_str_from_list(a, false);
              '  raw->key = RAW;
              '  .VarResult = mk_element(LISTITEM);
              '  .VarResult->children = raw;

        Case ucsAct_2_ListItemTight
             ' a = cons(.VarResult, a);
        Case ucsAct_1_ListItemTight
             ' a = cons(.VarResult, a);
        Case ucsAct_3_ListItem
             '  element *raw;
              '  raw = mk_str_from_list(a, false);
              '  raw->key = RAW;
              '  .VarResult = mk_element(LISTITEM);
              '  .VarResult->children = raw;

        Case ucsAct_2_ListItem
             ' a = cons(.VarResult, a);
        Case ucsAct_1_ListItem
             ' a = cons(.VarResult, a);
        Case ucsAct_3_ListBlock
             ' .VarResult = mk_str_from_list(a, false);
        Case ucsAct_2_ListBlock
             ' a = cons(.VarResult, a);
        Case ucsAct_1_ListBlock
             ' a = cons(.VarResult, a);
        Case ucsAct_3_ListContinuationBlock
             '  .VarResult = mk_str_from_list(a, false);
        Case ucsAct_2_ListContinuationBlock
             ' a = cons(.VarResult, a);
        Case ucsAct_1_ListContinuationBlock
             '   if (strlen(yytext) == 0)
                            '        a = cons(mk_str("\001"), a); /* block separator */
                            '   else
                            '        a = cons(mk_str(yytext), a);
        Case ucsAct_3_Str
             ' if (a->next == NULL) { .VarResult = a; } else { .VarResult = mk_list(LIST, a); }
        Case ucsAct_2_Str
             ' a = cons(.VarResult, a);
        Case ucsAct_1_Str
             ' a = cons(mk_str(yytext), a);
        Case ucsAct_1_UlOrStarLine
             ' .VarResult = mk_str(yytext);
        Case ucsAct_1_Space
             ' .VarResult = mk_str(" ");
          ' .VarResult->key = SPACE;
        Case ucsAct_2_Strike
             ' .VarResult = mk_list(STRIKE, a);
        Case ucsAct_1_Strike
             ' a = cons(b, a);
        Case ucsAct_1_Image
             ' if (.VarResult->key == LINK) { '
          '     .VarResult->key = IMAGE;
          ' } else { '
          '     element *result;
          '     result = .VarResult;
          '     .VarResult->children = cons(mk_str("!"), result->children);
          ' }
        Case ucsAct_1_NoteReference
             '   element *match;
                  '   if (find_note(&match, ref->contents.str)) { '
                  '       .VarResult = mk_element(NOTE);
                  '       assert(match->children != NULL);
                  '       .VarResult->children = match->children;
                  '       .VarResult->contents.str = 0;
                  '   } else { '
                  '       char *s;
                  '       s = malloc(strlen(ref->contents.str) + 4);
                  '       sprintf(s, "[^%s]", ref->contents.str);
                  '       .VarResult = mk_str(s);
                  '       free(s);
                  '   }

        Case ucsAct_2_InlineNote
             ' .VarResult = mk_list(NOTE, a);
                  ' .VarResult->contents.str = 0;
        Case ucsAct_1_InlineNote
             ' a = cons(.VarResult, a);
        Case ucsAct_1_Code
             ' .VarResult = mk_str(yytext); .VarResult->key = CODE;
        Case ucsAct_1_RawHtml
             '   if (extension(EXT_FILTER_HTML)) { '
              '       .VarResult = mk_list(LIST, NULL);
              '   } else { '
              '       .VarResult = mk_str(yytext);
              '       .VarResult->key = HTML;
              '   }

        Case ucsAct_1_Entity
             ' .VarResult = mk_str(yytext); .VarResult->key = HTML;
        Case ucsAct_1_EscapedChar
             ' .VarResult = mk_str(yytext);
        Case ucsAct_1_Symbol
             ' .VarResult = mk_str(yytext);
        Case ucsAct_1_StrChunk
             ' .VarResult = mk_str(yytext);
        Case ucsAct_1_AposChunk
             ' .VarResult = mk_element(APOSTROPHE);
        Case ucsAct_1_LineBreak
             ' .VarResult = mk_element(LINEBREAK);
        Case ucsAct_1_TerminalEndline
             ' .VarResult = NULL;
        Case ucsAct_1_NormalEndline
             ' .VarResult = mk_str("\n");
                    ' .VarResult->key = SPACE;
        Case ucsAct_3_EmphStar
             ' .VarResult = mk_list(EMPH, a);
        Case ucsAct_2_EmphStar
             ' a = cons(b, a);
        Case ucsAct_1_EmphStar
             ' a = cons(b, a);
        Case ucsAct_3_EmphUl
             ' .VarResult = mk_list(EMPH, a);
        Case ucsAct_2_EmphUl
             ' a = cons(b, a);
        Case ucsAct_1_EmphUl
             ' a = cons(b, a);
        Case ucsAct_2_StrongStar
             ' .VarResult = mk_list(STRONG, a);
        Case ucsAct_1_StrongStar
             ' a = cons(b, a);
        Case ucsAct_2_StrongUl
             ' .VarResult = mk_list(STRONG, a);
        Case ucsAct_1_StrongUl
             ' a = cons(b, a);
        Case ucsAct_1_ExplicitLink
             ' .VarResult = mk_link(l->children, s->contents.str, t->contents.str);
                  ' free_element(s);
                  ' free_element(t);
                  ' free(l);
        Case ucsAct_1_ReferenceLinkDouble
             '   link match;
                         '   if (find_reference(&match, b->children)) { '
                         '       .VarResult = mk_link(a->children, match.url, match.title);
                         '       free(a);
                         '       free_element_list(b);
                         '   } else { '
                         '       element *result;
                         '       result = mk_element(LIST);
                         '       result->children = cons(mk_str("["), cons(a, cons(mk_str("]"), cons(mk_str(yytext),
                         '                           cons(mk_str("["), cons(b, mk_str("]")))))));
                         '       .VarResult = result;
                         '   }

        Case ucsAct_1_ReferenceLinkSingle
             '   link match;
                         '   if (find_reference(&match, a->children)) { '
                         '       .VarResult = mk_link(a->children, match.url, match.title);
                         '       free(a);
                         '   }
                         '   else { '
                         '       element *result;
                         '       result = mk_element(LIST);
                         '       result->children = cons(mk_str("["), cons(a, cons(mk_str("]"), mk_str(yytext))));
                         '       .VarResult = result;
                         '   }

        Case ucsAct_2_Label
             ' .VarResult = mk_list(LIST, a);
        Case ucsAct_1_Label
             ' a = cons(.VarResult, a);
        Case ucsAct_1_Source
             ' .VarResult = mk_str(yytext);
        Case ucsAct_1_Title
             ' .VarResult = mk_str(yytext);
        Case ucsAct_1_AutoLinkUrl
             '   .VarResult = mk_link(mk_str(yytext), yytext, "");
        Case ucsAct_1_AutoLinkEmail
             '   char *mailto = malloc(strlen(yytext) + 8);
                  '   sprintf(mailto, "mailto:%s", yytext);
                  '   .VarResult = mk_link(mk_str(yytext), mailto, "");
                  '   free(mailto);

        Case ucsAct_1_RefSrc
             ' .VarResult = mk_str(yytext); 
           ' .VarResult->key = HTML;
        Case ucsAct_1_RefTitle
             ' .VarResult = mk_str(yytext);
        Case ucsAct_1_Ellipsis
             ' .VarResult = mk_element(ELLIPSIS);
        Case ucsAct_2_SingleQuoted
             ' .VarResult = mk_list(SINGLEQUOTED, a);
        Case ucsAct_1_SingleQuoted
             ' a = cons(b, a);
        Case ucsAct_2_DoubleQuoted
             ' .VarResult = mk_list(DOUBLEQUOTED, a);
        Case ucsAct_1_DoubleQuoted
             ' a = cons(b, a);
        Case ucsAct_1_Apostrophe
             ' .VarResult = mk_element(APOSTROPHE);
        Case ucsAct_1_EmDash
             ' .VarResult = mk_element(EMDASH);
        Case ucsAct_1_EnDash
             ' .VarResult = mk_element(ENDASH);
        Case ucsAct_1_RawNoteReference
             ' .VarResult = mk_str(yytext);
        Case ucsAct_3_RawNoteBlock
             '   .VarResult = mk_str_from_list(a, true);
                  '   .VarResult->key = RAW;

        Case ucsAct_2_RawNoteBlock
             ' a = cons(mk_str(yytext), a);
        Case ucsAct_1_RawNoteBlock
             ' a = cons(.VarResult, a);
        End Select
    End With
End Sub

'= user-defined functions ================================================

Private Function extension(ByVal eType As UcsStateEnum) As Boolean

End Function
