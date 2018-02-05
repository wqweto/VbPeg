VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIR"
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
' cIR.cls - VB codegen intermediate representation builder
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_NODE_TYPES        As String = "|Block|DeclFunc|DeclEnum|StmtCall|StmtRetVal|StmtExit|StmtDoLoop|StmtDoExit|StmtIf|StmtSelect|StmtCase|StmtSavePos|StmtRestorePos|StmtEmit|StmtWith|StmtCapture|ExprCall|ExprNot|ExprEmit"
Private Const NODE_ROOT             As Long = 0
Private Const IR_DECL               As Long = 0
Private Const IR_ROOT               As Long = 1
Private Const IR_CONTINUE           As Long = 0
Private Const LNG_COMMENT_COLUMN    As Long = 51
Private Const STR_FUNC_PREFIX       As String = "Parse"
Private Const STR_BUFPOS_PREFIX     As String = "p"
Private Const STR_THUNKPOS_PREFIX   As String = "q"
Private Const STR_CTX_VAR_NAME      As String = "ctx"
Private Const STR_ACTIONS_ENUM_NAME As String = "UcsParserActionsEnum"
Private Const STR_ACTIONS_ENUM_FMT  As String = "ucsAct_%1_%2"

Public Enum UcsIrNodeTypeEnum
    ucsIrtBlock = 1
    ucsIrtDeclFunc
    ucsIrtDeclEnum
    ucsIrtStmtCall
    ucsIrtStmtRetVal
    ucsIrtStmtExit
    ucsIrtStmtDoLoop
    ucsIrtStmtDoExit
    ucsIrtStmtIf
    ucsIrtStmtSelect
    ucsIrtStmtCase
    ucsIrtStmtSavePos
    ucsIrtStmtRestorePos
    ucsIrtStmtEmit
    ucsIrtStmtWith
    ucsIrtStmtCapture
    ucsIrtExprCall
    ucsIrtExprNot
    ucsIrtExprEmit
End Enum

Public Enum UcsIrNodeFlagsEnum
    ucsIrfPublic = 2 ^ 0
    ucsIrfVoid = 2 ^ 1
End Enum

Private Type UcsIrNodeType
    NodeType            As UcsIrNodeTypeEnum
    Text                As String
    Children()          As Long
    Count               As Long
    Flags               As UcsIrNodeFlagsEnum
    Comment             As String
End Type

Private Type UcsIrType
    Nodes()             As UcsIrNodeType
    Count               As Long
    CodeGenTree         As cTree
    IrActionsEnum       As Long
    IrActionsSelect     As Long
    OutModulePrefix     As String
    OutIndent           As Long
    OutBuilder          As Collection
    OutNodeTypes        As Variant
    OutFuncScope        As String
    OutFuncName         As String
    OutFuncType         As String
    LastError           As String
End Type

Private m_uIr                   As UcsIrType

'=========================================================================
' Properties
'=========================================================================

Property Get LastError() As String
    LastError = m_uIr.LastError
End Property

Property Get IrFlag(ByVal lNode As Long, ByVal eFlag As UcsIrNodeFlagsEnum) As Boolean
    IrFlag = (m_uIr.Nodes(lNode).Flags And eFlag) <> 0
End Property

Property Let IrFlag(ByVal lNode As Long, ByVal eFlag As UcsIrNodeFlagsEnum, ByVal bValue As Boolean)
    With m_uIr.Nodes(lNode)
        If bValue Then
            .Flags = .Flags Or eFlag
        Else
            .Flags = .Flags And Not eFlag
        End If
    End With
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function AddNode( _
            ByVal eNodeType As UcsIrNodeTypeEnum, _
            Optional Text As String, _
            Optional Comment As String) As Long
    With m_uIr
        If .Count = 0 Then
            ReDim .Nodes(0 To 4) As UcsIrNodeType
            '-- skip IR_DECL and IR_ROOT
            .Count = IR_ROOT + 1
        ElseIf .Count > UBound(.Nodes) Then
            ReDim Preserve .Nodes(0 To 2 * UBound(.Nodes)) As UcsIrNodeType
        End If
        AddNode = .Count
        With .Nodes(AddNode)
            .NodeType = eNodeType
            .Text = Text
            .Comment = Comment
        End With
        .Count = .Count + 1
        Select Case eNodeType
        Case ucsIrtDeclEnum
            AddChild IR_DECL, AddNode
        Case ucsIrtDeclFunc
            AddChild IR_ROOT, AddNode
        End Select
    End With
End Function

Public Function AddChild(ByVal lParent As Long, ByVal lNode As Long) As Long
    Dim lIdx            As Long
    
    If lNode <= IR_ROOT Then
        GoTo QH
    End If
    '--- if block -> try merge children in parent statement
    With m_uIr.Nodes(lNode)
        If .NodeType = ucsIrtBlock Then
            If m_uIr.Nodes(lParent).NodeType <> ucsIrtStmtIf Or .Count = 1 Then
                For lIdx = 0 To .Count - 1
                    AddChild lParent, .Children(lIdx)
                Next
                GoTo QH
            End If
        End If
    End With
    With m_uIr.Nodes(lParent)
        If .Count = 0 Then
            ReDim .Children(0 To 4) As Long
        ElseIf .Count > UBound(.Children) Then
            ReDim Preserve .Children(0 To 2 * UBound(.Children)) As Long
        End If
        .Children(.Count) = lNode
        .Count = .Count + 1
    End With
QH:
    AddChild = lParent
End Function

Public Function CodeGen(oTree As cTree, ByVal bAllRules As Boolean) As Boolean
    Dim vElem           As Variant
    Dim lBody           As Long
    Dim lIrNode         As Long
    Dim lSuccess        As Long
    Dim lFailure        As Long
    
    Set m_uIr.CodeGenTree = oTree
    With m_uIr.CodeGenTree
        For Each vElem In .NodeChildren(NODE_ROOT)
            If .NodeFlag(vElem, ucsNdfUsed) Or bAllRules Then
                lBody = AddNode(ucsIrtDeclFunc, STR_FUNC_PREFIX & .NodeText(vElem))
                IrFlag(lBody, ucsIrfPublic) = True
                If .NodeFlag(vElem, ucsNdfSavePos) Then
                    AddChild lBody, AddNode(ucsIrtStmtSavePos, STR_BUFPOS_PREFIX & vElem)
                    If .NodeFlag(vElem, ucsNdfHasActions) Then
                        AddChild lBody, AddNode(ucsIrtStmtSavePos, STR_THUNKPOS_PREFIX & vElem)
                    End If
                End If
                If .NodeFlag(vElem, ucsNdfAlwaysTrue) Then
                    IrFlag(lBody, ucsIrfVoid) = True
                    lSuccess = IR_CONTINUE
                    lFailure = IR_CONTINUE
                Else
                    lSuccess = AddChild(AddChild(AddNode(ucsIrtBlock), _
                        AddNode(ucsIrtStmtRetVal, "True")), _
                        AddNode(ucsIrtStmtExit))
                    lFailure = AddNode(ucsIrtStmtExit)
                End If
                lIrNode = pvCodeGenNode(.NodeChild(vElem, 0), lSuccess, lFailure)
                If lIrNode = 0 Then
                    GoTo QH
                End If
                pvCodeGenTrimLastExit lIrNode
                If pvCodeGenFindCtxUse(lIrNode) Then
                    AddChild lBody, AddChild(AddNode(ucsIrtStmtWith, STR_CTX_VAR_NAME), lIrNode)
                Else
                    AddChild lBody, lIrNode
                End If
            End If
        Next
    End With
    '--- success
    CodeGen = True
QH:
End Function

Private Function pvCodeGenNode( _
            ByVal lNode As Long, _
            ByVal lSuccess As Long, _
            ByVal lFailure As Long) As Long
    Dim lIdx            As Long
    Dim lChildNode      As Long
    Dim lIrNode         As Long
    Dim lIrFailure      As Long
    Dim sText           As String
    
    With m_uIr.CodeGenTree
        Select Case .NodeType(lNode)
        Case ucsNdtSequence
            pvCodeGenNode = AddNode(ucsIrtBlock)
            If .NodeFlag(lNode, ucsNdfSavePos) Then
                AddChild pvCodeGenNode, AddNode(ucsIrtStmtSavePos, STR_BUFPOS_PREFIX & lNode)
                If .NodeFlag(lNode, ucsNdfHasActions) Then
                    AddChild pvCodeGenNode, AddNode(ucsIrtStmtSavePos, STR_THUNKPOS_PREFIX & lNode)
                End If
            End If
            If lSuccess = IR_CONTINUE Then
                For lIdx = 0 To .NodeNumChildren(lNode) - 1
                    lChildNode = .NodeChild(lNode, lIdx)
                    If .NodeFlag(lChildNode, ucsNdfRestorePos) Then
                        lIrFailure = AddNode(ucsIrtBlock)
                        AddChild lIrFailure, AddNode(ucsIrtStmtRestorePos, STR_BUFPOS_PREFIX & .NodeRestoreFrom(lChildNode))
                        If .NodeFlag(lChildNode, ucsNdfHasActions) Then
                            AddChild lIrFailure, AddNode(ucsIrtStmtRestorePos, STR_THUNKPOS_PREFIX & .NodeRestoreFrom(lChildNode))
                        End If
                        AddChild lIrFailure, lFailure
                    Else
                        lIrFailure = lFailure
                    End If
                    lIrNode = pvCodeGenNode(lChildNode, IR_CONTINUE, lIrFailure)
                    If lIrNode = 0 Then
                        GoTo QH
                    End If
                    AddChild pvCodeGenNode, lIrNode
                Next
            Else
                lIrNode = lSuccess
                For lIdx = .NodeNumChildren(lNode) - 1 To 0 Step -1
                    lChildNode = .NodeChild(lNode, lIdx)
                    If .NodeFlag(lChildNode, ucsNdfRestorePos) Then
                        lIrFailure = AddNode(ucsIrtStmtRestorePos, STR_BUFPOS_PREFIX & .NodeRestoreFrom(lChildNode))
                        If .NodeFlag(lChildNode, ucsNdfHasActions) Then
                            lIrFailure = AddChild(AddChild(AddNode(ucsIrtBlock), _
                                lIrFailure), _
                                AddNode(ucsIrtStmtRestorePos, STR_THUNKPOS_PREFIX & .NodeRestoreFrom(lChildNode)))
                        End If
                    Else
                        lIrFailure = IR_CONTINUE
                    End If
                    lIrNode = pvCodeGenNode(lChildNode, lIrNode, lIrFailure)
                    If lIrNode = 0 Then
                        GoTo QH
                    End If
                Next
                AddChild pvCodeGenNode, lIrNode
                AddChild pvCodeGenNode, lFailure
            End If
        Case ucsNdtAlternate
            pvCodeGenNode = AddNode(ucsIrtBlock)
            If .NodeFlag(lNode, ucsNdfSavePos) Then
                AddChild pvCodeGenNode, AddNode(ucsIrtStmtSavePos, STR_BUFPOS_PREFIX & lNode)
                If .NodeFlag(lNode, ucsNdfHasActions) Then
                    AddChild pvCodeGenNode, AddNode(ucsIrtStmtSavePos, STR_THUNKPOS_PREFIX & lNode)
                End If
            End If
            If lSuccess = IR_CONTINUE Then
                lIrNode = lFailure
                For lIdx = .NodeNumChildren(lNode) - 1 To 0 Step -1
                    lChildNode = .NodeChild(lNode, lIdx)
                    If .NodeFlag(lChildNode, ucsNdfRestorePos) Then
                        lIrFailure = AddNode(ucsIrtBlock)
                        AddChild lIrFailure, AddNode(ucsIrtStmtRestorePos, STR_BUFPOS_PREFIX & .NodeRestoreFrom(lChildNode))
                        If .NodeFlag(lChildNode, ucsNdfHasActions) Then
                            AddChild lIrFailure, AddNode(ucsIrtStmtRestorePos, STR_THUNKPOS_PREFIX & .NodeRestoreFrom(lChildNode))
                        End If
                        AddChild lIrFailure, lIrNode
                    Else
                        lIrFailure = lIrNode
                    End If
                    lIrNode = pvCodeGenNode(lChildNode, IR_CONTINUE, lIrFailure)
                    If lIrNode = 0 Then
                        GoTo QH
                    End If
                Next
                AddChild pvCodeGenNode, lIrNode
            Else
                For lIdx = 0 To .NodeNumChildren(lNode) - 1
                    lChildNode = .NodeChild(lNode, lIdx)
                    If .NodeFlag(lChildNode, ucsNdfRestorePos) Then
                        lIrFailure = AddNode(ucsIrtStmtRestorePos, STR_BUFPOS_PREFIX & .NodeRestoreFrom(lChildNode))
                        If .NodeFlag(lChildNode, ucsNdfHasActions) Then
                            lIrFailure = AddChild(AddChild(AddNode(ucsIrtBlock), _
                                lIrFailure), _
                                AddNode(ucsIrtStmtRestorePos, STR_THUNKPOS_PREFIX & .NodeRestoreFrom(lChildNode)))
                        End If
                    Else
                        lIrFailure = IR_CONTINUE
                    End If
                    lIrNode = pvCodeGenNode(lChildNode, lSuccess, lIrFailure)
                    If lIrNode = 0 Then
                        GoTo QH
                    End If
                    AddChild pvCodeGenNode, lIrNode
                Next
                AddChild pvCodeGenNode, lFailure
            End If
        Case ucsNdtReference
            If .NodeFlag(lNode, ucsNdfAlwaysTrue) Or _
                    lSuccess = IR_CONTINUE And lFailure = IR_CONTINUE Then
                pvCodeGenNode = AddNode(ucsIrtBlock)
                AddChild pvCodeGenNode, AddNode(ucsIrtStmtCall, STR_FUNC_PREFIX & .NodeText(lNode))
                AddChild pvCodeGenNode, lSuccess
            ElseIf lSuccess = IR_CONTINUE Then
                pvCodeGenNode = AddNode(ucsIrtStmtIf)
                AddChild pvCodeGenNode, AddChild(AddNode(ucsIrtExprNot), AddNode(ucsIrtExprCall, STR_FUNC_PREFIX & .NodeText(lNode)))
                AddChild pvCodeGenNode, lFailure
            Else
                pvCodeGenNode = AddNode(ucsIrtStmtIf)
                AddChild pvCodeGenNode, AddNode(ucsIrtExprCall, STR_FUNC_PREFIX & .NodeText(lNode))
                AddChild pvCodeGenNode, lSuccess
                AddChild pvCodeGenNode, lFailure
            End If
        Case ucsNdtAction
            sText = Replace(Replace(STR_ACTIONS_ENUM_FMT, "%1", .NodeId(lNode)), "%2", .NodeText(.NodeRule(lNode)))
            pvCodeGenNode = AddChild(AddChild(AddNode(ucsIrtBlock), _
                AddNode(ucsIrtStmtEmit, "pvPushAction " & sText)), _
                lSuccess)
            If m_uIr.IrActionsEnum = 0 Then
                m_uIr.IrActionsEnum = AddNode(ucsIrtDeclEnum, STR_ACTIONS_ENUM_NAME)
            End If
            AddChild m_uIr.IrActionsEnum, AddNode(ucsIrtStmtEmit, sText)
            If m_uIr.IrActionsSelect = 0 Then
                m_uIr.IrActionsSelect = AddNode(ucsIrtStmtSelect, "eAction")
                lIrNode = AddChild( _
                    AddNode(ucsIrtDeclFunc, "pvImplAction(eAction As " & STR_ACTIONS_ENUM_NAME & ", ByVal lOffset As Long, ByVal lSize As Long)"), _
                        AddChild(AddNode(ucsIrtStmtWith, STR_CTX_VAR_NAME & ".UserData"), _
                            m_uIr.IrActionsSelect))
                IrFlag(lIrNode, ucsIrfVoid) = True
            End If
            AddChild m_uIr.IrActionsSelect, _
                AddChild(AddNode(ucsIrtStmtCase, sText), _
                    AddNode(ucsIrtStmtEmit, .NodeText(lNode)))
        Case ucsNdtStar, ucsNdtPlus
            lIrNode = pvCodeGenNode(.NodeChild(lNode, 0), IR_CONTINUE, AddNode(ucsIrtStmtDoExit))
            If lIrNode = 0 Then
                GoTo QH
            End If
            If lSuccess = IR_CONTINUE Then
                pvCodeGenNode = AddChild(AddNode(ucsIrtStmtDoLoop), lIrNode)
            Else
                pvCodeGenNode = AddChild(AddChild(AddNode(ucsIrtBlock), _
                    AddChild(AddNode(ucsIrtStmtDoLoop), lIrNode)), _
                    lSuccess)
            End If
            If .NodeType(lNode) = ucsNdtPlus Then
                lIrNode = pvCodeGenNode(.NodeChild(lNode, 0), IR_CONTINUE, lFailure)
                If lIrNode = 0 Then
                    GoTo QH
                End If
                pvCodeGenNode = AddChild(AddChild(AddNode(ucsIrtBlock), lIrNode), pvCodeGenNode)
            End If
        Case ucsNdtQuery
            pvCodeGenNode = AddNode(ucsIrtBlock)
            AddChild pvCodeGenNode, pvCodeGenNode(.NodeChild(lNode, 0), IR_CONTINUE, IR_CONTINUE)
            AddChild pvCodeGenNode, lSuccess
        Case ucsNdtPeekFor
            If Not .NodeFlag(lNode, ucsNdfIsNot) Then
                If .NodeFlag(lNode, ucsNdfRestorePos) Then
                    lIrNode = AddNode(ucsIrtBlock)
                    AddChild lIrNode, AddNode(ucsIrtStmtRestorePos, STR_BUFPOS_PREFIX & .NodeRestoreFrom(lNode))
                    If .NodeFlag(lNode, ucsNdfHasActions) Then
                        AddChild lIrNode, AddNode(ucsIrtStmtRestorePos, STR_THUNKPOS_PREFIX & .NodeRestoreFrom(lNode))
                    End If
                    AddChild lIrNode, lSuccess
                Else
                    lIrNode = lSuccess
                End If
                pvCodeGenNode = pvCodeGenNode(.NodeChild(lNode, 0), lIrNode, lFailure)
            Else
                pvCodeGenNode = pvCodeGenNode(.NodeChild(lNode, 0), lFailure, lSuccess)
            End If
        Case ucsNdtString
            sText = pvFromPegString(.NodeText(lNode))
            If Len(sText) = 1 Then
                If Not .NodeFlag(lNode, ucsNdfIsNot) Then
                    lIrNode = AddNode(ucsIrtExprEmit, ".BufData(.BufPos) = " & Asc(sText))
                Else
                    lIrNode = AddNode(ucsIrtExprEmit, ".BufData(.BufPos) <> " & Asc(sText))
                End If
            ElseIf Not .NodeFlag(lNode, ucsNdfIsNot) Then
                lIrNode = AddNode(ucsIrtExprEmit, "pvMatchString(" & pvToVbString(sText) & ")")
            Else
                lIrNode = AddNode(ucsIrtExprEmit, "Not pvMatchString(" & pvToVbString(sText) & ")")
            End If
            If Not .NodeFlag(lNode, ucsNdfSkipAdvance) Then
                lSuccess = AddChild(AddChild(AddNode(ucsIrtBlock), _
                    AddNode(ucsIrtStmtEmit, ".BufPos = .BufPos + " & Len(sText))), _
                    lSuccess)
            End If
EmitStmtIfOnIrNode:
            pvCodeGenNode = AddNode(ucsIrtStmtIf, Comment:=IIf(.NodeType(lNode) = ucsNdtString, """" & .NodeText(lNode) & """", vbNullString))
            If lSuccess = IR_CONTINUE Then
                AddChild pvCodeGenNode, AddChild(AddNode(ucsIrtExprNot), lIrNode)
                AddChild pvCodeGenNode, lFailure
            Else
                AddChild pvCodeGenNode, lIrNode
                AddChild pvCodeGenNode, lSuccess
                AddChild pvCodeGenNode, lFailure
            End If
        Case ucsNdtDot
            If Not .NodeFlag(lNode, ucsNdfIsNot) Then
                lIrNode = AddNode(ucsIrtExprEmit, ".BufPos < .BufSize")
            Else
                lIrNode = AddNode(ucsIrtExprEmit, ".BufPos >= .BufSize")
            End If
            If Not .NodeFlag(lNode, ucsNdfSkipAdvance) Then
                lSuccess = AddChild(AddChild(AddNode(ucsIrtBlock), _
                    AddNode(ucsIrtStmtEmit, ".BufPos = .BufPos + 1")), _
                    lSuccess)
            End If
            GoTo EmitStmtIfOnIrNode
        Case ucsNdtPredicate
            Select Case .NodeText(lNode)
            Case "YY_BEGIN"
                pvCodeGenNode = AddNode(ucsIrtBlock)
                AddChild pvCodeGenNode, AddNode(ucsIrtStmtCapture, ".CaptureBegin")
                AddChild pvCodeGenNode, lSuccess
            Case "YY_END"
                pvCodeGenNode = AddNode(ucsIrtBlock)
                AddChild pvCodeGenNode, AddNode(ucsIrtStmtCapture, ".CaptureEnd")
                AddChild pvCodeGenNode, lSuccess
            Case Else
                If Not .NodeFlag(lNode, ucsNdfIsNot) Then
                    lIrNode = AddNode(ucsIrtExprEmit, .NodeText(lNode))
                Else
                    lIrNode = AddNode(ucsIrtExprEmit, "Not (" & .NodeText(lNode) & ")")
                End If
                GoTo EmitStmtIfOnIrNode
            End Select
        Case ucsNdtClass
            pvCodeGenNode = AddNode(ucsIrtStmtSelect, ".BufData(.BufPos)")
            lIrNode = AddNode(ucsIrtStmtCase, pvToVbCaseRange(pvFromPegString(.NodeText(lNode))), Comment:="[" & .NodeText(lNode) & "]")
            If Not .NodeFlag(lNode, ucsNdfIsNot) Then
                AddChild lIrNode, AddNode(ucsIrtStmtEmit, ".BufPos = .BufPos + 1")
            End If
            AddChild lIrNode, lSuccess
            AddChild pvCodeGenNode, lIrNode
            If lFailure <> IR_CONTINUE Or .NodeFlag(lNode, ucsNdfIsNot) Then
                lIrNode = AddNode(ucsIrtStmtCase, "Else")
                If .NodeFlag(lNode, ucsNdfIsNot) Then
                    AddChild lIrNode, AddNode(ucsIrtStmtEmit, ".BufPos = .BufPos + 1")
                End If
                AddChild pvCodeGenNode, AddChild(lIrNode, lFailure)
            End If
        Case Else
            m_uIr.LastError = "Unsupported parse tree node type " & .NodeType(lNode)
            GoTo QH
        End Select
    End With
    Exit Function
QH:
End Function

Private Function pvCodeGenFindCtxUse(ByVal lNode As Long) As Boolean
    Dim lIdx            As Long
    
    With m_uIr.Nodes(lNode)
        '--- note: euristic for emitted VB6 code that uses context member var
        If Left$(.Text, 1) = "." Or .NodeType = ucsIrtStmtSavePos Or .NodeType = ucsIrtStmtRestorePos Then
            pvCodeGenFindCtxUse = True
            Exit Function
        End If
        For lIdx = 0 To .Count - 1
            If pvCodeGenFindCtxUse(.Children(lIdx)) Then
                pvCodeGenFindCtxUse = True
                Exit Function
            End If
        Next
    End With
End Function

Private Function pvCodeGenTrimLastExit(lNode As Long) As Boolean
    Dim lNewNode        As Long
    
    With m_uIr.Nodes(lNode)
        If .Count > 0 Then
            If m_uIr.Nodes(.Children(.Count - 1)).NodeType = ucsIrtStmtExit _
                    And LenB(m_uIr.Nodes(.Children(.Count - 1)).Text) = 0 Then
                lNewNode = AddNode(.NodeType)
                m_uIr.Nodes(lNewNode) = m_uIr.Nodes(lNode)
                m_uIr.Nodes(lNewNode).Count = .Count - 1
                lNode = lNewNode
                pvCodeGenTrimLastExit = True
            End If
        End If
    End With
    With m_uIr.Nodes(lNode)
        If .Count > 0 Then
            lNewNode = .Children(.Count - 1)
            If pvCodeGenTrimLastExit(lNewNode) Then
                .Children(.Count - 1) = lNewNode
'                pvCodeGenTrimLastExit = True
            End If
        End If
    End With
End Function

Private Function pvFromPegString(sText As String) As String
    Dim lIdx            As Long
    
    For lIdx = 1 To Len(sText)
        If Mid$(sText, lIdx, 1) = "\" And Len(sText) > lIdx Then
            Select Case Mid$(sText, lIdx + 1, 1)
            Case "a"
                pvFromPegString = pvFromPegString & Chr$(7)     ' alert
            Case "b"
                pvFromPegString = pvFromPegString & Chr$(8)     ' bell
            Case "e"
                pvFromPegString = pvFromPegString & Chr$(27)    ' escape
            Case "f"
                pvFromPegString = pvFromPegString & Chr$(12)    ' formfeed
            Case "n"
                pvFromPegString = pvFromPegString & vbLf
            Case "r"
                pvFromPegString = pvFromPegString & vbCr
            Case "t"
                pvFromPegString = pvFromPegString & vbTab
            Case "v"
                pvFromPegString = pvFromPegString & vbVerticalTab
            Case "'", """", "[", "]", "\"
                pvFromPegString = pvFromPegString & Mid$(sText, lIdx + 1, 1)
            Case Else
                pvFromPegString = pvFromPegString & "\"
                lIdx = lIdx - 1
            End Select
            lIdx = lIdx + 1
        Else
            pvFromPegString = pvFromPegString & Mid$(sText, lIdx, 1)
        End If
    Next
End Function

Private Function pvToVbString(sText As String) As String
    pvToVbString = """" & Replace(Replace(Replace(Replace(Replace( _
        sText, """", """"""), _
        vbCrLf, """ & vbCrLf & """), _
        vbTab, """ & vbTab & """), _
        vbCr, """ & vbCr & """), _
        vbLf, """ & vbLf & """) & """"
    If Left$(pvToVbString, 5) = """"" & " Then
        pvToVbString = Mid$(pvToVbString, 6)
    End If
    If Right$(pvToVbString, 5) = " & """"" Then
        pvToVbString = Left$(pvToVbString, Len(pvToVbString) - 5)
    End If
End Function

Private Function pvToVbCaseRange(sText As String) As String
    Dim lIdx            As Long
    
    For lIdx = 1 To Len(sText)
        pvToVbCaseRange = IIf(LenB(pvToVbCaseRange) <> 0, pvToVbCaseRange & ", ", vbNullString) & Asc(Mid$(sText, lIdx, 1))
        If Mid$(sText, lIdx + 1, 1) = "-" And Len(sText) > lIdx + 1 Then
            pvToVbCaseRange = pvToVbCaseRange & " To " & Asc(Mid$(sText, lIdx + 2, 1))
            lIdx = lIdx + 2
        End If
    Next
End Function

Public Function DumpIrTree() As String
    Dim lIdx            As Long
    
    With m_uIr
        Set .OutBuilder = New Collection
        .OutNodeTypes = Split(STR_NODE_TYPES, "|")
        With .Nodes(IR_DECL)
            For lIdx = 0 To .Count - 1
                pvOutputIrNode .Children(lIdx)
            Next
        End With
        With .Nodes(IR_ROOT)
            For lIdx = 0 To .Count - 1
                pvOutputIrNode .Children(lIdx)
            Next
        End With
        DumpIrTree = ConcatCollection(.OutBuilder, vbCrLf)
    End With
End Function

'--- note: used only for debugging codegen
Private Function DumpIrNode(ByVal lNode As Long) As String
    With m_uIr
        Set .OutBuilder = New Collection
        .OutNodeTypes = Split(STR_NODE_TYPES, "|")
        .OutIndent = 0
        pvOutputIrNode lNode
        DumpIrNode = ConcatCollection(.OutBuilder, vbCrLf)
    End With
End Function

Private Sub pvOutputIrNode(ByVal lNode As Long)
    Dim lIdx            As Long
    
    With m_uIr.Nodes(lNode)
        pvOutput "[" & lNode & "] " & At(m_uIr.OutNodeTypes, .NodeType) & _
            IIf(LenB(.Text) <> 0, ", Text=" & .Text, vbNullString)
        m_uIr.OutIndent = m_uIr.OutIndent + 1
        For lIdx = 0 To .Count - 1
            If .Children(lIdx) > IR_ROOT Then
                pvOutputIrNode .Children(lIdx)
            Else
                pvOutput "Missing"
            End If
        Next
        m_uIr.OutIndent = m_uIr.OutIndent - 1
    End With
End Sub

Public Function EmitCode( _
            cOutput As Collection, _
            ByVal bClass As String, _
            ByVal bPublic As Boolean, _
            sModuleName As String, _
            sUserData As String) As Boolean
    Dim lIdx            As Long
    
    If LenB(sModuleName) = 0 Then
        sModuleName = IIf(bClass, "cParser", "mdParser")
    End If
    If LenB(sUserData) = 0 Then
        sUserData = "Variant"
    End If
    With m_uIr
        Set .OutBuilder = New Collection
        .OutIndent = 0
        If bClass Then
            pvOutput _
"VERSION 1.0 CLASS" & vbCrLf & _
"BEGIN" & vbCrLf & _
"  MultiUse = -1  'True" & vbCrLf & _
"  Persistable = 0  'NotPersistable" & vbCrLf & _
"  DataBindingBehavior = 0  'vbNone" & vbCrLf & _
"  DataSourceBehavior  = 0  'vbNone" & vbCrLf & _
"  MTSTransactionMode  = 0  'NotAnMTSObject" & vbCrLf & _
"END" & vbCrLf & _
"Attribute VB_Name = """ & sModuleName & """" & vbCrLf & _
"Attribute VB_GlobalNameSpace = False" & vbCrLf & _
"Attribute VB_Creatable = True" & vbCrLf & _
"Attribute VB_PredeclaredId = False" & vbCrLf & _
"Attribute VB_Exposed = " & bPublic
        Else
            .OutModulePrefix = "VbPeg"
            pvOutput _
"Attribute VB_Name = """ & sModuleName & """"
        End If
        pvOutput _
"' Auto-generated on " & Now & vbCrLf & _
"Option Explicit" & vbCrLf & _
"DefObj A-Z" & vbCrLf
        pvOutput _
"'=========================================================================" & vbCrLf & _
"' API" & vbCrLf & _
"'=========================================================================" & vbCrLf
        pvOutput _
"Private Declare Sub CopyMemory Lib ""kernel32"" Alias ""RtlMoveMemory"" (Destination As Any, Source As Any, ByVal Length As Long)" & vbCrLf & _
"Private Declare Function RtlCompareMemory Lib ""ntdll"" (Source1 As Any, Source2 As Any, ByVal Length As Long) As Long" & vbCrLf
        pvOutput _
"'=========================================================================" & vbCrLf & _
"' Constants and member variables" & vbCrLf & _
"'=========================================================================" & vbCrLf
        pvOutput _
"'= generated enum ========================================================" & vbCrLf
        With .Nodes(IR_DECL)
            For lIdx = 0 To .Count - 1
                If Not pvEmitStmt(.Children(lIdx)) Then
                    GoTo QH
                End If
                pvOutput vbNullString
            Next
        End With
        pvOutput _
"Private Type UcsParserThunkType" & vbCrLf & _
"    Action              As " & STR_ACTIONS_ENUM_NAME & vbCrLf & _
"    CaptureBegin        As Long" & vbCrLf & _
"    CaptureEnd          As Long" & vbCrLf & _
"End Type" & vbCrLf
        pvOutput _
"Private Type UcsParserType" & vbCrLf & _
"    Contents            As String" & vbCrLf & _
"    BufData()           As Integer" & vbCrLf & _
"    BufPos              As Long" & vbCrLf & _
"    BufSize             As Long" & vbCrLf & _
"    ThunkData()         As UcsParserThunkType" & vbCrLf & _
"    ThunkPos            As Long" & vbCrLf & _
"    CaptureBegin        As Long" & vbCrLf & _
"    CaptureEnd          As Long" & vbCrLf & _
"    LastError           As String" & vbCrLf & _
"    UserData            As " & sUserData & vbCrLf & _
"End Type" & vbCrLf
        pvOutput _
"Private " & STR_CTX_VAR_NAME & "                     As UcsParserType" & vbCrLf
        pvOutput _
"'=========================================================================" & vbCrLf & _
"' Properties" & vbCrLf & _
"'=========================================================================" & vbCrLf
        pvOutput _
"Property Get " & m_uIr.OutModulePrefix & "LastError() As String" & vbCrLf & _
"    " & m_uIr.OutModulePrefix & "LastError = " & STR_CTX_VAR_NAME & ".LastError" & vbCrLf & _
"End Property" & vbCrLf
        pvOutput _
"Property Get " & m_uIr.OutModulePrefix & "ParserVersion() As String" & vbCrLf & _
"    " & m_uIr.OutModulePrefix & "ParserVersion = """ & Now & """" & vbCrLf & _
"End Property" & vbCrLf
        pvOutput _
"'=========================================================================" & vbCrLf & _
"' Methods" & vbCrLf & _
"'=========================================================================" & vbCrLf
        pvOutput _
"Public Function " & .OutModulePrefix & "Init(sContents As String, UserData As " & sUserData & ") As Boolean" & vbCrLf & _
"    Dim uEmpty          As UcsParserType" & vbCrLf & _
"    " & vbCrLf & _
"    " & STR_CTX_VAR_NAME & " = uEmpty" & vbCrLf & _
"    With " & STR_CTX_VAR_NAME & vbCrLf & _
"        If LenB(sContents) = 0 Then" & vbCrLf & _
"            .LastError = ""Input has no contents""" & vbCrLf & _
"            Exit Function" & vbCrLf & _
"        End If" & vbCrLf & _
"        .Contents = sContents" & vbCrLf & _
"        ReDim .BufData(0 To Len(sContents) + 3) As Integer" & vbCrLf & _
"        Call CopyMemory(.BufData(0), ByVal StrPtr(sContents), LenB(sContents))" & vbCrLf & _
"        .BufSize = Len(sContents)" & vbCrLf & _
"        .BufData(.BufSize) = -1" & vbCrLf & _
"        ReDim .ThunkData(0 To 4) As UcsParserThunkType" & vbCrLf & _
"        If IsObject(UserData) Then" & vbCrLf & _
"            Set .UserData = UserData" & vbCrLf & _
"        Else" & vbCrLf & _
"            .UserData = UserData" & vbCrLf & _
"        End If" & vbCrLf & _
"    End With" & vbCrLf & _
"    " & m_uIr.OutModulePrefix & "Init = True" & vbCrLf & _
"End Function" & vbCrLf
        pvOutput _
"Public Function " & .OutModulePrefix & "Flush() As Boolean" & vbCrLf & _
"    Dim lIdx            As Long" & vbCrLf & _
"    " & vbCrLf & _
"    With " & STR_CTX_VAR_NAME & vbCrLf & _
"        For lIdx = 0 To .ThunkPos - 1" & vbCrLf & _
"            With .ThunkData(lIdx)" & vbCrLf & _
"                pvImplAction .Action, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin" & vbCrLf & _
"            End With" & vbCrLf & _
"        Next" & vbCrLf & _
"        .BufPos = 0" & vbCrLf & _
"        ReDim .ThunkData(0 To 4) As UcsParserThunkType" & vbCrLf & _
"        .ThunkPos = 0" & vbCrLf & _
"        .CaptureBegin = 0" & vbCrLf & _
"        .CaptureEnd = 0" & vbCrLf & _
"        .LastError = vbNullString" & vbCrLf & _
"    End With" & vbCrLf & _
"    " & m_uIr.OutModulePrefix & "Flush = True" & vbCrLf & _
"End Function" & vbCrLf
        pvOutput _
"Private Sub pvPushAction(ByVal eAction As " & STR_ACTIONS_ENUM_NAME & ")" & vbCrLf & _
"    With " & STR_CTX_VAR_NAME & vbCrLf & _
"        If UBound(.ThunkData) < .ThunkPos Then" & vbCrLf & _
"            ReDim Preserve .ThunkData(0 To 2 * UBound(.ThunkData)) As UcsParserThunkType" & vbCrLf & _
"        End If" & vbCrLf & _
"        With .ThunkData(.ThunkPos)" & vbCrLf & _
"            .Action = eAction" & vbCrLf & _
"            .CaptureBegin = " & STR_CTX_VAR_NAME & ".CaptureBegin" & vbCrLf & _
"            .CaptureEnd = " & STR_CTX_VAR_NAME & ".CaptureEnd" & vbCrLf & _
"        End With" & vbCrLf & _
"        .ThunkPos = .ThunkPos + 1" & vbCrLf & _
"    End With" & vbCrLf & _
"End Sub" & vbCrLf
        pvOutput _
"Private Function pvMatchString(sText As String) As Boolean" & vbCrLf & _
"    With " & STR_CTX_VAR_NAME & vbCrLf & _
"        If .BufPos + Len(sText) <= .BufSize Then" & vbCrLf & _
"            pvMatchString = RtlCompareMemory(.BufData(.BufPos), ByVal StrPtr(sText), LenB(sText)) = LenB(sText)" & vbCrLf & _
"        End If" & vbCrLf & _
"    End With" & vbCrLf & _
"End Function" & vbCrLf
        pvOutput _
"'= generated functions ===================================================" & vbCrLf
        With .Nodes(IR_ROOT)
            For lIdx = 0 To .Count - 1
                If Not pvEmitStmt(.Children(lIdx)) Then
                    GoTo QH
                End If
                pvOutput vbNullString
            Next
        End With
        Set cOutput = .OutBuilder
        '--- success
        EmitCode = True
    End With
QH:
End Function

Private Function pvEmitStmt(ByVal lNode As Long) As Boolean
    Dim lIdx            As Long
    
    With m_uIr.Nodes(lNode)
        Select Case .NodeType
        Case ucsIrtBlock
            '--- don't indent
            For lIdx = 0 To .Count - 1
                If Not pvEmitStmt(.Children(lIdx)) Then
                    GoTo QH
                End If
            Next
        Case ucsIrtDeclFunc
            m_uIr.OutFuncScope = IIf((.Flags And ucsIrfPublic) <> 0, "Public", "Private")
            m_uIr.OutFuncName = .Text
            If (.Flags And ucsIrfVoid) <> 0 Then
                m_uIr.OutFuncType = "Sub"
                pvOutput m_uIr.OutFuncScope & " Sub " & .Text & IIf(Right$(.Text, 1) <> ")", "()", vbNullString)
            Else
                m_uIr.OutFuncType = "Function"
                pvOutput m_uIr.OutFuncScope & " Function " & .Text & "() As Boolean"
            End If
            m_uIr.OutIndent = m_uIr.OutIndent + 1
            For lIdx = 0 To .Count - 1
                If Not pvEmitStmt(.Children(lIdx)) Then
                    GoTo QH
                End If
            Next
            m_uIr.OutIndent = m_uIr.OutIndent - 1
            pvOutput "End " & m_uIr.OutFuncType
        Case ucsIrtDeclEnum
            pvOutput "Private Enum " & .Text
            m_uIr.OutIndent = m_uIr.OutIndent + 1
            For lIdx = 0 To .Count - 1
                If Not pvEmitStmt(.Children(lIdx)) Then
                    GoTo QH
                End If
            Next
            m_uIr.OutIndent = m_uIr.OutIndent - 1
            pvOutput "End Enum"
        Case ucsIrtStmtCall
            pvOutput "Call " & .Text
        Case ucsIrtStmtRetVal
            pvOutput m_uIr.OutFuncName & " = " & .Text
        Case ucsIrtStmtExit
            pvOutput "Exit " & m_uIr.OutFuncType
        Case ucsIrtStmtDoLoop
            pvOutput "Do"
            m_uIr.OutIndent = m_uIr.OutIndent + 1
            For lIdx = 0 To .Count - 1
                If Not pvEmitStmt(.Children(lIdx)) Then
                    GoTo QH
                End If
            Next
            m_uIr.OutIndent = m_uIr.OutIndent - 1
            pvOutput "Loop"
        Case ucsIrtStmtDoExit
            pvOutput "Exit Do"
        Case ucsIrtStmtIf
            pvOutput "If " & pvEmitExpr(.Children(0)) & " Then", Comment:=.Comment
            If .Count > 1 Then
                m_uIr.OutIndent = m_uIr.OutIndent + 1
                pvEmitStmt .Children(1)
                m_uIr.OutIndent = m_uIr.OutIndent - 1
            Else
                pvOutput "'--- do nothing"
            End If
            If .Count > 2 Then
                pvOutput "Else"
                m_uIr.OutIndent = m_uIr.OutIndent + 1
                pvEmitStmt .Children(2)
                m_uIr.OutIndent = m_uIr.OutIndent - 1
            End If
            pvOutput "End If"
        Case ucsIrtStmtSelect
            pvOutput "Select Case " & .Text
            For lIdx = 0 To .Count - 1
                If Not pvEmitStmt(.Children(lIdx)) Then
                    GoTo QH
                End If
            Next
            pvOutput "End Select"
        Case ucsIrtStmtCase
            pvOutput "Case " & .Text, Comment:=.Comment
            m_uIr.OutIndent = m_uIr.OutIndent + 1
            For lIdx = 0 To .Count - 1
                If Not pvEmitStmt(.Children(lIdx)) Then
                    GoTo QH
                End If
            Next
            m_uIr.OutIndent = m_uIr.OutIndent - 1
        Case ucsIrtStmtSavePos
            pvOutput "Dim " & .Text & " As Long"
            If Left$(.Text, Len(STR_BUFPOS_PREFIX)) = STR_BUFPOS_PREFIX Then
                pvOutput .Text & " = .BufPos"
            Else
                pvOutput .Text & " = .ThunkPos"
            End If
        Case ucsIrtStmtRestorePos
            If Left$(.Text, Len(STR_BUFPOS_PREFIX)) = STR_BUFPOS_PREFIX Then
                pvOutput ".BufPos = " & .Text
            Else
                pvOutput ".ThunkPos = " & .Text
            End If
        Case ucsIrtStmtEmit
            pvOutput .Text
        Case ucsIrtStmtWith
            pvOutput "With " & .Text
            m_uIr.OutIndent = m_uIr.OutIndent + 1
            For lIdx = 0 To .Count - 1
                If Not pvEmitStmt(.Children(lIdx)) Then
                    GoTo QH
                End If
            Next
            m_uIr.OutIndent = m_uIr.OutIndent - 1
            pvOutput "End With"
        Case ucsIrtStmtCapture
            pvOutput .Text & " = .BufPos"
        Case Else
            m_uIr.LastError = "Unsupported IR node type " & .NodeType
            GoTo QH
        End Select
    End With
    '--- success
    pvEmitStmt = True
QH:
End Function

Private Function pvEmitExpr(ByVal lNode As Long) As String
    With m_uIr.Nodes(lNode)
        Select Case .NodeType
        Case ucsIrtExprCall
            pvEmitExpr = .Text & "()"
        Case ucsIrtExprNot
            pvEmitExpr = "Not " & pvEmitExpr(.Children(0))
        Case ucsIrtExprEmit
            pvEmitExpr = .Text
        End Select
    End With
End Function

Private Sub pvOutput(ByVal sText As String, Optional Comment As String)
    sText = RTrim$(Space$(m_uIr.OutIndent * 4) & sText)
    If LenB(Comment) Then
        If Len(sText) < LNG_COMMENT_COLUMN Then
            sText = sText & Space$(LNG_COMMENT_COLUMN - Len(sText))
        End If
        sText = sText & " ' " & Comment
    End If
    m_uIr.OutBuilder.Add sText
End Sub

