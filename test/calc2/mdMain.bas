Attribute VB_Name = "mdMain"
Option Explicit

Private Const STD_OUTPUT_HANDLE             As Long = -11&
Private Const STD_ERROR_HANDLE              As Long = -12&

Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, lpszDst As Any, ByVal cchDstLength As Long) As Long

Sub Main()
    Dim sText       As String
    Dim vResult     As Double
    Dim vError      As Variant
    Dim lIdx        As Long
    Dim bResult     As Boolean
    Dim dblTimer    As Double
    
    sText = Command$
    dblTimer = Timer
'    Do While dblTimer = Timer
'        dblTimer = Timer
'    Loop
'    dblTimer = Timer
    For lIdx = 1 To 100000
        bResult = VbPegMatch(sText, Result:=vResult)
    Next
    If bResult Then
        ConsolePrint "Result: %1" & vbCrLf, vResult
        If LenB(VbPegLastError) <> 0 Then
            ConsolePrint "Error: %1" & vbCrLf, VbPegLastError
        End If
    Else
        ConsolePrint "Error: %1" & vbCrLf, VbPegLastError
    End If
    ConsolePrint "Elapsed: %1" & vbCrLf, Format$(Timer - dblTimer, "0.000")
End Sub

Sub Main1()
    Dim lIdx        As Long
    Dim dblTimer    As Double
    Dim vResult     As Variant
    
    dblTimer = Timer
    For lIdx = 1 To 1000000
        vResult = CDbl("0.75")
'        VbPegMatch "1*0.75", Result:=vResult
'        VbPegMatch "1*5", Result:=vResult
'        VbPegMatch "1*0.5+20", Result:=vResult
    Next
    Debug.Print Timer - dblTimer
End Sub

Public Function ConsolePrint(ByVal sText As String, ParamArray A() As Variant) As String
    ConsolePrint = pvConsoleOutput(GetStdHandle(STD_OUTPUT_HANDLE), sText, CVar(A))
End Function

Public Function ConsoleError(ByVal sText As String, ParamArray A() As Variant) As String
    ConsoleError = pvConsoleOutput(GetStdHandle(STD_ERROR_HANDLE), sText, CVar(A))
End Function

Private Function pvConsoleOutput(ByVal hOut As Long, ByVal sText As String, A As Variant) As String
    Const LNG_PRIVATE   As Long = &HE1B6 '-- U+E000 to U+F8FF - Private Use Area (PUA)
    Dim lIdx            As Long
    Dim sArg            As String
    Dim baBuffer()      As Byte
    Dim dwDummy         As Long

    If LenB(sText) = 0 Then
        Exit Function
    End If
    '--- format
    For lIdx = UBound(A) To LBound(A) Step -1
        sArg = Replace(A(lIdx), "%", ChrW$(LNG_PRIVATE))
        sText = Replace(sText, "%" & (lIdx - LBound(A) + 1), sArg)
    Next
    pvConsoleOutput = Replace(sText, ChrW$(LNG_PRIVATE), "%")
    '--- output
    If hOut = 0 Then
        Debug.Print pvConsoleOutput;
    Else
        ReDim baBuffer(0 To Len(pvConsoleOutput) - 1) As Byte
        If CharToOemBuff(pvConsoleOutput, baBuffer(0), UBound(baBuffer) + 1) Then
            Call WriteFile(hOut, baBuffer(0), UBound(baBuffer) + 1, dwDummy, ByVal 0&)
        End If
    End If
End Function
