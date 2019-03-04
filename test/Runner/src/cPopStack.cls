VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPopStack"
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
' cPopStack.cls - pop from stack using RAII
'
'=========================================================================
Option Explicit

Public Stack        As VBA.Collection

Private Sub Class_Terminate()
    Stack.Remove 1
End Sub
