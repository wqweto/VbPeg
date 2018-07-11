Attribute VB_Name = "Module2"
Option Explicit

Sub main()
    Dim sContents As String
    Dim vResult As Variant
    
    sContents = "1+2*3"
    If VbPegMatch(sContents, Result:=vResult) > 0 Then
        Debug.Print sContents, vResult
    End If
End Sub
