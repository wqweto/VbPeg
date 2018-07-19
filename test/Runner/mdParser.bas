Attribute VB_Name = "mdParser"
' Auto-generated on 19.7.2018 16:48:40
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
    LastBufPos          As Long
    UserData            As Variant
    VarResult           As Variant
End Type

Private ctx                     As UcsParserType

'=========================================================================
' Properties
'=========================================================================

Property Get VbPegLastError() As String
    VbPegLastError = ctx.LastError
End Property

Property Get VbPegLastOffset() As Long
    VbPegLastOffset = ctx.LastBufPos + 1
End Property

Property Get VbPegParserVersion() As String
    VbPegParserVersion = "19.7.2018 16:48:40"
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function VbPegMatch(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant, Optional Result As Variant) As Long
    If VbPegBeginMatch(sSubject, StartPos, UserData) Then
        If VbPegParselist() Then
            VbPegMatch = VbPegEndMatch()
        ElseIf LenB(ctx.LastError) = 0 Then
            ctx.LastError = "Fail"
        End If
    End If
End Function

Public Function VbPegBeginMatch(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant) As Boolean
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
    VbPegBeginMatch = True
End Function

Public Function VbPegEndMatch() As Long
    With ctx
        VbPegEndMatch = .BufPos + 1
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

'= generated functions ===================================================

Public Function VbPegParselist() As Boolean
    Dim p7 As Long

    With ctx
        p7 = .BufPos
        If VbPegParseopen() Then
            Do
                If Not VbPegParseelem() Then
                    Exit Do
                End If
            Loop
            If VbPegParseclose() Then
                VbPegParselist = True
                Exit Function
            Else
                .BufPos = p7
            End If
        End If
    End With
End Function

Public Function VbPegParseopen() As Boolean
    With ctx
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Do
                If Not VbPegParsespace() Then
                    Exit Do
                End If
            Loop
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            VbPegParseopen = True
        End If
    End With
End Function

Public Function VbPegParseelem() As Boolean
    If VbPegParselist() Then
        VbPegParseelem = True
        Exit Function
    End If
    If VbPegParseatom() Then
        VbPegParseelem = True
        Exit Function
    End If
    If VbPegParsesstring() Then
        VbPegParseelem = True
        Exit Function
    End If
    If VbPegParsedstring() Then
        VbPegParseelem = True
    End If
End Function

Public Function VbPegParseclose() As Boolean
    With ctx
        If .BufData(.BufPos) = 41 Then              ' ")"
            .BufPos = .BufPos + 1
            Do
                If Not VbPegParsespace() Then
                    Exit Do
                End If
            Loop
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            VbPegParseclose = True
        End If
    End With
End Function

Public Function VbPegParseatom() As Boolean
    Dim i19 As Long

    With ctx
        For i19 = 0 To LNG_MAXINT
            Select Case .BufData(.BufPos)
            Case 97 To 122, 48 To 57, 95            ' [a-z0-9_]
                .BufPos = .BufPos + 1
            Case Else
                Exit For
            End Select
        Next
        If i19 <> 0 Then
            Do
                If Not VbPegParsespace() Then
                    Exit Do
                End If
            Loop
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            VbPegParseatom = True
        End If
    End With
End Function

Public Function VbPegParsesstring() As Boolean
    Dim p34 As Long

    With ctx
        p34 = .BufPos
        If .BufData(.BufPos) = 39 Then              ' "'"
            .BufPos = .BufPos + 1
            Do
                If .BufData(.BufPos) <> 39 And .BufPos < .BufSize Then ' "'"
                    .BufPos = .BufPos + 1
                Else
                    Exit Do
                End If
            Loop
            If .BufData(.BufPos) = 39 Then          ' "'"
                .BufPos = .BufPos + 1
                Do
                    If Not VbPegParsespace() Then
                        Exit Do
                    End If
                Loop
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString: .LastBufPos = .BufPos
                End If
                VbPegParsesstring = True
                Exit Function
            Else
                .BufPos = p34
            End If
        End If
    End With
End Function

Public Function VbPegParsedstring() As Boolean
    Dim p27 As Long

    With ctx
        p27 = .BufPos
        If .BufData(.BufPos) = 34 Then              ' """
            .BufPos = .BufPos + 1
            Do
                If .BufData(.BufPos) <> 34 And .BufPos < .BufSize Then ' """
                    .BufPos = .BufPos + 1
                Else
                    Exit Do
                End If
            Loop
            If .BufData(.BufPos) = 34 Then          ' """
                .BufPos = .BufPos + 1
                Do
                    If Not VbPegParsespace() Then
                        Exit Do
                    End If
                Loop
                If .BufPos > .LastBufPos Then
                    .LastError = vbNullString: .LastBufPos = .BufPos
                End If
                VbPegParsedstring = True
                Exit Function
            Else
                .BufPos = p27
            End If
        End If
    End With
End Function

Public Function VbPegParsespace() As Boolean
    With ctx
        If .BufData(.BufPos) = 32 Then              ' " "
            .BufPos = .BufPos + 1
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            VbPegParsespace = True
            Exit Function
        End If
        If .BufData(.BufPos) = 9 Then               ' "\t"
            .BufPos = .BufPos + 1
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            VbPegParsespace = True
            Exit Function
        End If
        If VbPegParseeol() Then
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            VbPegParsespace = True
        End If
    End With
End Function

Public Function VbPegParseeol() As Boolean
    With ctx
        If .BufData(.BufPos) = 13 And .BufData(.BufPos + 1) = 10 Then ' "\r\n"
            .BufPos = .BufPos + 2
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            VbPegParseeol = True
            Exit Function
        End If
        If .BufData(.BufPos) = 10 Then              ' "\n"
            .BufPos = .BufPos + 1
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            VbPegParseeol = True
            Exit Function
        End If
        If .BufData(.BufPos) = 13 Then              ' "\r"
            .BufPos = .BufPos + 1
            If .BufPos > .LastBufPos Then
                .LastError = vbNullString: .LastBufPos = .BufPos
            End If
            VbPegParseeol = True
        End If
    End With
End Function

