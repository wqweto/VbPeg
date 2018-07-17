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

Private Const STR_NODE_TYPES        As String = "|Block|DeclFunc|DeclEnum|DeclLabel|StmtCall|StmtRetVal|StmtExit|StmtDoLoop|StmtDoExit|StmtIf|StmtSelect|StmtCase|StmtSavePos|StmtRestorePos|StmtIncrPos|StmtEmit|StmtWith|StmtCapture|StmtPushAction|StmtGoTo|StmtForLoop|StmtForExit|ExprCall|ExprNot|ExprEmit"
Private Const NODE_ROOT             As Long = 0
Private Const IR_CONTINUE           As Long = 0
Private Const IR_DECL               As Long = 1
Private Const IR_ROOT               As Long = 2
Private Const LNG_COMMENT_COLUMN    As Long = 51
Private Const STR_FUNC_PREFIX       As String = "Parse"
Private Const STR_BUFPOS_PREFIX     As String = "p"
Private Const STR_THUNKPOS_PREFIX   As String = "q"
Private Const STR_FOR_PREFIX        As String = "i"
Private Const STR_CTX_VAR_NAME      As String = "ctx"
Private Const STR_ACTIONS_ENUM_NAME As String = "UcsParserActionsEnum"
Private Const STR_ACTIONS_ENUM_FMT  As String = "ucsAct_%1_%2"
Private Const STR_ACTIONS_ENUM_VARS As String = "ucsActVarAlloc = -1|ucsActVarSet = -2|ucsActResultClear = -3|ucsActResultSet = -4"

Public Enum UcsIrNodeTypeEnum
    ucsIrtBlock = 1
    ucsIrtDeclFunc
    ucsIrtDeclEnum
    ucsIrtDeclLabel
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
    ucsIrtStmtIncrPos
    ucsIrtStmtEmit
    ucsIrtStmtWith
    ucsIrtStmtCapture
    ucsIrtStmtPushAction
    ucsIrtStmtGoTo
    ucsIrtStmtForLoop
    ucsIrtStmtForExit
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
    Parent              As Long
    Children()          As Long
    Count               As Long
    Flags               As UcsIrNodeFlagsEnum
    Comment             As String
End Type

Private Type UcsIrType
    Nodes()             As UcsIrNodeType
    Count               As Long
    CodeGenTree         As cTree
    CodeGenId           As Long
    CodeGenStartFunc    As String
    CodeGenVarStack     As Boolean
    IrActionsEnum       As Long
    IrActionsSelect     As Long
    OutModulePrefix     As String
    OutBuilder          As Collection
    OutNodeTypes        As Variant
    OutIndent           As Long
    OutDictionary       As Object
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

Property Get SettingValue(sName As String) As String
    SettingValue = m_uIr.CodeGenTree.SettingValue(sName)
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
                    m_uIr.Nodes(.Children(lIdx)).Parent = 0
                    AddChild lParent, .Children(lIdx)
                Next
                GoTo QH
            End If
        End If
    End With
    '--- note: outside With block as m_uIr.Nodes might get resized in AddNode
    If m_uIr.Nodes(lNode).Parent <> 0 Then
        lIdx = AddNode(ucsIrtBlock)
        m_uIr.Nodes(lIdx) = m_uIr.Nodes(lNode)
        lNode = lIdx
    End If
    Debug.Assert m_uIr.Nodes(lNode).Parent <> lParent
    m_uIr.Nodes(lNode).Parent = lParent
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
    
    On Error GoTo EH
    With m_uIr
        Set .CodeGenTree = oTree
        .CodeGenId = 0
        Set .OutDictionary = CreateObject("Scripting.Dictionary")
        .OutDictionary.CompareMode = vbTextCompare
        With .CodeGenTree
            For Each vElem In .NodeChildren(NODE_ROOT)
                If .NodeFlag(vElem, ucsNdfUsed) Or bAllRules Then
                    pvCodeGenUniqueName vElem
                    If .NodeNumVariables(vElem) > 0 Then
                        m_uIr.CodeGenVarStack = True
                    End If
                End If
            Next
            For Each vElem In .NodeChildren(NODE_ROOT)
                If .NodeFlag(vElem, ucsNdfUsed) Or bAllRules Then
                    lBody = AddNode(ucsIrtDeclFunc, STR_FUNC_PREFIX & .NodeText(vElem))
                    IrFlag(lBody, ucsIrfPublic) = (.NodeText(vElem) <> UCase$(.NodeText(vElem)))
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
                        If .NodeFlag(vElem, ucsNdfHasAdvance) Then
                            lIrNode = AddChild(AddChild( _
                                AddNode(ucsIrtStmtIf), _
                                    AddNode(ucsIrtExprEmit, ".BufPos > .LastBufPos")), _
                                    AddNode(ucsIrtStmtEmit, ".LastError = vbNullString : .LastBufPos = .BufPos"))
                        Else
                            lIrNode = IR_CONTINUE
                        End If
                        lSuccess = AddChild(AddChild(AddChild( _
                            AddNode(ucsIrtBlock), _
                                lIrNode), _
                                AddNode(ucsIrtStmtRetVal, "True")), _
                                AddNode(ucsIrtStmtExit))
                        lFailure = IR_CONTINUE
                    End If
                    If .NodeNumVariables(vElem) > 0 Then
                        lSuccess = AddChild(AddChild( _
                            AddNode(ucsIrtBlock), _
                                AddNode(ucsIrtStmtEmit, "pvPushThunk ucsActVarAlloc, -" & .NodeNumVariables(vElem))), _
                                lSuccess)
                    End If
                    lIrNode = pvCodeGenNode(.NodeChild(vElem, 0), lSuccess, lFailure)
                    If lIrNode = 0 Then
                        GoTo QH
                    End If
                    pvCodeGenTrimLastExit lIrNode
                    pvCodeGenTrimGotoNext lIrNode
                    If .NodeNumVariables(vElem) > 0 Then
                        lIrNode = AddChild(AddChild( _
                            AddNode(ucsIrtBlock), _
                                AddNode(ucsIrtStmtEmit, "pvPushThunk ucsActVarAlloc, " & .NodeNumVariables(vElem))), _
                                lIrNode)
                    End If
                    If pvCodeGenFindCtxUse(lIrNode) Then
                        AddChild lBody, AddChild(AddNode(ucsIrtStmtWith, STR_CTX_VAR_NAME), lIrNode)
                    Else
                        AddChild lBody, lIrNode
                    End If
                    If vElem = .StartRule Then
                        m_uIr.CodeGenStartFunc = m_uIr.Nodes(lBody).Text
                    End If
                End If
            Next
        End With
        If .IrActionsSelect <> 0 Then
            lIrNode = AddNode(ucsIrtDeclFunc, "pvImplAction(ByVal eAction As " & STR_ACTIONS_ENUM_NAME & _
                ", ByVal lOffset As Long, ByVal lSize As Long)")
            If LenB(SettingValue("prolog")) <> 0 Or LenB(SettingValue("epilog")) <> 0 Then
                AddChild lIrNode, AddChild(AddChild(AddChild( _
                    AddNode(ucsIrtBlock), _
                        AddNode(ucsIrtStmtEmit, LTrim$(pvTrimCrLf(SettingValue("prolog"))))), _
                        .IrActionsSelect), _
                        AddNode(ucsIrtStmtEmit, LTrim$(pvTrimCrLf(SettingValue("epilog")))))
            Else
                AddChild lIrNode, AddChild( _
                    AddNode(ucsIrtStmtWith, STR_CTX_VAR_NAME), _
                        .IrActionsSelect)
            End If
            IrFlag(lIrNode, ucsIrfVoid) = True
        End If
        If m_uIr.CodeGenVarStack Then
            If m_uIr.IrActionsEnum = 0 Then
                m_uIr.IrActionsEnum = AddNode(ucsIrtDeclEnum, STR_ACTIONS_ENUM_NAME)
            End If
            For Each vElem In Split(STR_ACTIONS_ENUM_VARS, "|")
                AddChild m_uIr.IrActionsEnum, _
                    AddNode(ucsIrtStmtEmit, CStr(vElem))
            Next
        End If
    End With
    '--- success
    CodeGen = True
QH:
    Exit Function
EH:
    m_uIr.LastError = Err.Description & " [cIR.CodeGen" & vbCrLf & Err.Source & "]"
End Function

Private Sub pvCodeGenUniqueName(ByVal lNode As Long)
    Dim lIdx            As Long
    
    With m_uIr.CodeGenTree
        .NodeText(lNode) = Replace(.NodeText(lNode), "-", "_")
        If m_uIr.OutDictionary.Exists(.NodeText(lNode)) Then
            For lIdx = 2 To 1000
                If Not m_uIr.OutDictionary.Exists(.NodeText(lNode) & lIdx) Then
                    .NodeText(lNode) = .NodeText(lNode) & lIdx
                    Exit For
                End If
            Next
        End If
        m_uIr.OutDictionary.Item(.NodeText(lNode)) = True
    End With
End Sub

Private Function pvCodeGenNode( _
            ByVal lNode As Long, _
            ByVal lSuccess As Long, _
            ByVal lFailure As Long) As Long
    Dim lIdx            As Long
    Dim lIrNode         As Long
    Dim sText           As String
    Dim bIgnoreCase     As Boolean
    
    On Error GoTo EH
    With m_uIr.CodeGenTree
        If .NodeFlag(lNode, ucsNdfRestorePos) Then
            If .NodeFlag(lNode, ucsNdfHasActions) Then
                lIrNode = AddNode(ucsIrtStmtRestorePos, STR_THUNKPOS_PREFIX & .NodeRestoreFrom(lNode))
            Else
                lIrNode = IR_CONTINUE
            End If
            lFailure = AddChild(AddChild(AddChild( _
                AddNode(ucsIrtBlock), _
                    AddNode(ucsIrtStmtRestorePos, STR_BUFPOS_PREFIX & .NodeRestoreFrom(lNode))), _
                    lIrNode), _
                    lFailure)
        End If
        If .NodeRefVar(lNode) <> 0 Then
            lSuccess = AddChild(AddChild( _
                AddNode(ucsIrtBlock), _
                    AddNode(ucsIrtStmtEmit, "pvPushThunk ucsActVarSet, " & .NodeId(.NodeRefVar(lNode)))), _
                    lSuccess)
        End If
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
                If Not pvCodeGenIsFinalJump(lFailure) Then
                    sText = "L" & pvCodeGenId
                    lFailure = AddChild(AddChild( _
                        AddNode(ucsIrtBlock), _
                            lFailure), _
                            AddNode(ucsIrtStmtGoTo, sText))
                End If
                For lIdx = 0 To .NodeNumChildren(lNode) - 1
                    lIrNode = pvCodeGenNode(.NodeChild(lNode, lIdx), IR_CONTINUE, pvCodeGenReplaceLabels(lFailure))
                    If lIrNode = 0 Then
                        GoTo QH
                    End If
                    AddChild pvCodeGenNode, lIrNode
                Next
                If LenB(sText) <> 0 Then
                    AddChild pvCodeGenNode, AddNode(ucsIrtDeclLabel, sText)
                End If
            Else
                If Not pvCodeGenIsFinalJump(lSuccess) Then
                    sText = "L" & pvCodeGenId
                    lSuccess = AddChild(AddChild( _
                        AddNode(ucsIrtBlock), _
                            lSuccess), _
                            AddNode(ucsIrtStmtGoTo, sText))
                End If
                lIrNode = lSuccess
                For lIdx = .NodeNumChildren(lNode) - 1 To 0 Step -1
                    lIrNode = pvCodeGenNode(.NodeChild(lNode, lIdx), lIrNode, pvCodeGenReplaceLabels(lFailure))
                    If lIrNode = 0 Then
                        GoTo QH
                    End If
                Next
                AddChild pvCodeGenNode, lIrNode
                If LenB(sText) <> 0 Then
                    AddChild pvCodeGenNode, AddNode(ucsIrtDeclLabel, sText)
                End If
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
                If Not pvCodeGenIsFinalJump(lFailure) Then
                    sText = "L" & pvCodeGenId
                    lFailure = AddChild(AddChild( _
                        AddNode(ucsIrtBlock), _
                            lFailure), _
                            AddNode(ucsIrtStmtGoTo, sText))
                End If
                If LenB(sText) = 0 Then
                    sText = "L" & pvCodeGenId
                End If
                lSuccess = AddNode(ucsIrtStmtGoTo, sText)
                lIrNode = lFailure
                For lIdx = .NodeNumChildren(lNode) - 1 To 0 Step -1
                    lIrNode = pvCodeGenNode(.NodeChild(lNode, lIdx), lSuccess, lIrNode)
                    If lIrNode = 0 Then
                        GoTo QH
                    End If
                Next
                AddChild pvCodeGenNode, lIrNode
                If LenB(sText) <> 0 Then
                    AddChild pvCodeGenNode, AddNode(ucsIrtDeclLabel, sText)
                End If
            Else
                If Not pvCodeGenIsFinalJump(lSuccess) Then
                    sText = "L" & pvCodeGenId
                    lSuccess = AddChild(AddChild( _
                        AddNode(ucsIrtBlock), _
                            lSuccess), _
                            AddNode(ucsIrtStmtGoTo, sText))
                End If
                For lIdx = 0 To .NodeNumChildren(lNode) - 1
                    lIrNode = pvCodeGenNode(.NodeChild(lNode, lIdx), pvCodeGenReplaceLabels(lSuccess), IR_CONTINUE)
                    If lIrNode = 0 Then
                        GoTo QH
                    End If
                    AddChild pvCodeGenNode, lIrNode
                Next
                AddChild pvCodeGenNode, lFailure
                If LenB(sText) <> 0 Then
                    AddChild pvCodeGenNode, AddNode(ucsIrtDeclLabel, sText)
                End If
            End If
        Case ucsNdtReference
            If .NodeFlag(lNode, ucsNdfAlwaysTrue) Or _
                    lSuccess = IR_CONTINUE And lFailure = IR_CONTINUE Then
                pvCodeGenNode = AddChild(AddChild( _
                    AddNode(ucsIrtBlock), _
                        AddNode(ucsIrtStmtCall, STR_FUNC_PREFIX & .NodeText(.NodeRefRule(lNode)))), _
                        lSuccess)
            ElseIf lSuccess = IR_CONTINUE Then
                pvCodeGenNode = AddChild(AddChild( _
                    AddNode(ucsIrtStmtIf), AddChild( _
                        AddNode(ucsIrtExprNot), _
                            AddNode(ucsIrtExprCall, STR_FUNC_PREFIX & .NodeText(.NodeRefRule(lNode))))), _
                        lFailure)
            Else
                pvCodeGenNode = AddChild(AddChild(AddChild( _
                    AddNode(ucsIrtStmtIf), _
                        AddNode(ucsIrtExprCall, STR_FUNC_PREFIX & .NodeText(.NodeRefRule(lNode)))), _
                        lSuccess), _
                        lFailure)
            End If
        Case ucsNdtError
            If .NodeFlag(lNode, ucsNdfAlwaysTrue) Or _
                    lSuccess = IR_CONTINUE And lFailure = IR_CONTINUE Then
                pvCodeGenNode = AddChild(AddChild( _
                    AddNode(ucsIrtBlock), _
                        AddNode(ucsIrtStmtCall, STR_FUNC_PREFIX & .NodeText(.NodeRefRule(lNode)))), _
                        lSuccess)
            Else
                lIrNode = AddNode(ucsIrtStmtEmit, ".LastError = " & pvToVbString(.NodeText(lNode)) & " : .LastBufPos = .BufPos")
                If lFailure = IR_CONTINUE Then
                    lFailure = lIrNode
                Else
                    lFailure = AddChild(AddChild( _
                        AddNode(ucsIrtBlock), _
                            lIrNode), _
                            lFailure)
                End If
                If .NodeRefRule(lNode) = 0 Then
                    pvCodeGenNode = lFailure
                Else
                    pvCodeGenNode = AddChild(AddChild(AddChild( _
                        AddNode(ucsIrtStmtIf), _
                            AddNode(ucsIrtExprCall, STR_FUNC_PREFIX & .NodeText(.NodeRefRule(lNode)))), _
                            lSuccess), _
                            lFailure)
                End If
            End If
        Case ucsNdtAction
            '-- always true -> discard lFailure
            If LenB(Trim$(.NodeText(lNode))) <> 0 Then
                sText = Replace(Replace(STR_ACTIONS_ENUM_FMT, "%1", .NodeId(lNode)), "%2", .NodeText(.NodeRefRule(lNode)))
            End If
            pvCodeGenNode = AddChild(AddChild( _
                AddNode(ucsIrtBlock), _
                    AddNode(ucsIrtStmtPushAction, sText)), _
                    lSuccess)
            If LenB(sText) <> 0 Then
                If m_uIr.IrActionsEnum = 0 Then
                    m_uIr.IrActionsEnum = AddNode(ucsIrtDeclEnum, STR_ACTIONS_ENUM_NAME)
                End If
                AddChild m_uIr.IrActionsEnum, _
                    AddNode(ucsIrtStmtEmit, sText)
                If m_uIr.IrActionsSelect = 0 Then
                    m_uIr.IrActionsSelect = AddNode(ucsIrtStmtSelect, "eAction")
                End If
                AddChild m_uIr.IrActionsSelect, AddChild( _
                    AddNode(ucsIrtStmtCase, sText), _
                        AddNode(ucsIrtStmtEmit, pvCodeGenAction(.NodeText(lNode), .NodeRefRule(lNode))))
            End If
        Case ucsNdtStar
            '-- always true -> discard lFailure
            lIrNode = pvCodeGenNode(.NodeChild(lNode, 0), IR_CONTINUE, AddNode(ucsIrtStmtDoExit))
            If lIrNode = 0 Then
                GoTo QH
            End If
            If lSuccess = IR_CONTINUE Then
                pvCodeGenNode = AddChild( _
                    AddNode(ucsIrtStmtDoLoop), _
                        lIrNode)
            Else
                pvCodeGenNode = AddChild(AddChild( _
                    AddNode(ucsIrtBlock), AddChild( _
                        AddNode(ucsIrtStmtDoLoop), _
                            lIrNode)), _
                        lSuccess)
            End If
        Case ucsNdtPlus
            lIrNode = pvCodeGenNode(.NodeChild(lNode, 0), IR_CONTINUE, AddNode(ucsIrtStmtForExit))
            If lIrNode = 0 Then
                GoTo QH
            End If
            If lSuccess = IR_CONTINUE Then
                pvCodeGenNode = AddChild(AddChild( _
                    AddNode(ucsIrtBlock), AddChild( _
                        AddNode(ucsIrtStmtForLoop, STR_FOR_PREFIX & lNode), _
                            lIrNode)), AddChild(AddChild( _
                        AddNode(ucsIrtStmtIf), _
                            AddNode(ucsIrtExprEmit, STR_FOR_PREFIX & lNode & " = 0")), _
                            lFailure))
            Else
                pvCodeGenNode = AddChild(AddChild( _
                    AddNode(ucsIrtBlock), AddChild( _
                        AddNode(ucsIrtStmtForLoop, STR_FOR_PREFIX & lNode), _
                            lIrNode)), AddChild(AddChild(AddChild( _
                        AddNode(ucsIrtStmtIf), _
                            AddNode(ucsIrtExprEmit, STR_FOR_PREFIX & lNode & " <> 0")), _
                            lSuccess), _
                            lFailure))
            End If
        Case ucsNdtQuery
            '-- always true -> discard lFailure
            sText = "L" & pvCodeGenId
            pvCodeGenNode = AddNode(ucsIrtBlock)
            AddChild pvCodeGenNode, pvCodeGenNode(.NodeChild(lNode, 0), AddNode(ucsIrtStmtGoTo, sText), IR_CONTINUE)
            AddChild pvCodeGenNode, AddNode(ucsIrtDeclLabel, sText)
            AddChild pvCodeGenNode, lSuccess
        Case ucsNdtPeekFor
            If Not .NodeFlag(lNode, ucsNdfIsNot) Then
                If .NodeFlag(lNode, ucsNdfRestorePos) Then
                    lIrNode = AddChild( _
                        AddNode(ucsIrtBlock), _
                            AddNode(ucsIrtStmtRestorePos, STR_BUFPOS_PREFIX & .NodeRestoreFrom(lNode)))
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
            bIgnoreCase = .NodeFlag(lNode, ucsNdfIgnoreCase) And LCase$(sText) <> UCase$(sText)
            If Len(sText) = 1 Then
                If Not .NodeFlag(lNode, ucsNdfIsNot) Then
                    If bIgnoreCase Then
                        lIrNode = AddNode(ucsIrtExprEmit, ".BufData(.BufPos) = " & AscW(LCase$(sText)) & " Or .BufData(.BufPos) = " & AscW(UCase$(sText)))
                    Else
                        lIrNode = AddNode(ucsIrtExprEmit, ".BufData(.BufPos) = " & AscW(sText))
                    End If
                Else
                    If bIgnoreCase Then
                        lIrNode = AddNode(ucsIrtExprEmit, ".BufData(.BufPos) <> " & AscW(LCase$(sText)) & " And .BufData(.BufPos) <> " & AscW(UCase$(sText)) & " And .BufPos < .BufSize")
                    Else
                        lIrNode = AddNode(ucsIrtExprEmit, ".BufData(.BufPos) <> " & AscW(sText) & " And .BufPos < .BufSize")
                    End If
                End If
            ElseIf Len(sText) = 2 And Not bIgnoreCase Then
                If Not .NodeFlag(lNode, ucsNdfIsNot) Then
                    lIrNode = AddNode(ucsIrtExprEmit, ".BufData(.BufPos) = " & AscW(sText) & " And .BufData(.BufPos + 1) = " & AscW(Mid$(sText, 2, 1)))
                Else
                    lIrNode = AddNode(ucsIrtExprEmit, "(.BufData(.BufPos) <> " & AscW(sText) & " Or .BufData(.BufPos + 1) <> " & AscW(Mid$(sText, 2, 1)) & ") And .BufPos + 1 < .BufSize")
                End If
            ElseIf Len(sText) = 0 Then
                If Not .NodeFlag(lNode, ucsNdfIsNot) Then
                    If lSuccess = IR_CONTINUE Then
                        pvCodeGenNode = AddNode(ucsIrtBlock)
                    Else
                        pvCodeGenNode = lSuccess
                    End If
                Else
                    If lFailure = IR_CONTINUE Then
                        pvCodeGenNode = AddNode(ucsIrtBlock)
                    Else
                        pvCodeGenNode = lFailure
                    End If
                End If
                GoTo QH
            ElseIf Not .NodeFlag(lNode, ucsNdfIsNot) Then
                lIrNode = AddNode(ucsIrtExprEmit, IIf(bIgnoreCase, "pvMatchIgnoreCase", "pvMatchString") & "(" & pvToVbString(sText) & ")")
            Else
                lIrNode = AddNode(ucsIrtExprEmit, "Not " & IIf(bIgnoreCase, "pvMatchIgnoreCase", "pvMatchString") & "(" & pvToVbString(sText) & ")")
            End If
            If Not .NodeFlag(lNode, ucsNdfSkipAdvance) Then
                lSuccess = AddChild(AddChild(AddNode(ucsIrtBlock), _
                    AddNode(ucsIrtStmtIncrPos, Len(sText))), _
                    lSuccess)
            End If
EmitStmtIfOnIrNode:
            pvCodeGenNode = AddNode(ucsIrtStmtIf, Comment:=IIf(.NodeType(lNode) = ucsNdtString, """" & .NodeText(lNode) & """" & _
                IIf(.NodeFlag(lNode, ucsNdfIgnoreCase), "i", vbNullString), vbNullString))
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
                    AddNode(ucsIrtStmtIncrPos, 1)), _
                    lSuccess)
            End If
            GoTo EmitStmtIfOnIrNode
        Case ucsNdtPredicate
            Select Case .NodeText(lNode)
            Case "YY_BEGIN"
                pvCodeGenNode = AddNode(ucsIrtBlock)
                AddChild pvCodeGenNode, AddNode(ucsIrtStmtCapture, "lCaptureBegin")
                AddChild pvCodeGenNode, lSuccess
            Case "YY_END"
                pvCodeGenNode = AddNode(ucsIrtBlock)
                AddChild pvCodeGenNode, AddNode(ucsIrtStmtCapture, "lCaptureEnd")
                AddChild pvCodeGenNode, lSuccess
            Case "YY_CLEAR"
                pvCodeGenNode = AddNode(ucsIrtBlock)
                AddChild pvCodeGenNode, AddNode(ucsIrtStmtEmit, "pvPushThunk ucsActResultClear")
                AddChild pvCodeGenNode, lSuccess
            Case "YY_SET"
                pvCodeGenNode = AddNode(ucsIrtBlock)
                AddChild pvCodeGenNode, AddNode(ucsIrtStmtEmit, "pvPushThunk ucsActResultSet, lCaptureBegin, lCaptureEnd")
                AddChild pvCodeGenNode, lSuccess
            Case Else
                sText = pvCodeGenAction(.NodeText(lNode), .NodeRefRule(lNode), IsImmediate:=True)
                If Not .NodeFlag(lNode, ucsNdfIsNot) Then
                    lIrNode = AddNode(ucsIrtExprEmit, sText)
                Else
                    lIrNode = AddNode(ucsIrtExprEmit, "Not (" & sText & ")")
                End If
                GoTo EmitStmtIfOnIrNode
            End Select
        Case ucsNdtClass
            pvCodeGenNode = AddNode(ucsIrtStmtSelect, ".BufData(.BufPos)")
            lIrNode = AddNode(ucsIrtStmtCase, pvToVbCaseRange(pvFromPegString(.NodeText(lNode)), .NodeFlag(lNode, ucsNdfIgnoreCase)), _
                Comment:="[" & .NodeText(lNode) & "]" & IIf(.NodeFlag(lNode, ucsNdfIgnoreCase), "i", vbNullString))
            AddChild pvCodeGenNode, lIrNode
            If Not .NodeFlag(lNode, ucsNdfIsNot) Then
                If Not .NodeFlag(lNode, ucsNdfSkipAdvance) Then
                    AddChild lIrNode, AddNode(ucsIrtStmtIncrPos, 1)
                End If
                AddChild lIrNode, lSuccess
                If lFailure <> IR_CONTINUE Then
                    AddChild pvCodeGenNode, AddChild( _
                        AddNode(ucsIrtStmtCase, "Else"), _
                            lFailure)
                End If
            Else
                AddChild lIrNode, lFailure
                AddChild pvCodeGenNode, AddChild( _
                    AddNode(ucsIrtStmtCase, "Else"), AddChild(AddChild(AddChild( _
                        AddNode(ucsIrtStmtIf), _
                            AddNode(ucsIrtExprEmit, ".BufPos < .BufSize")), AddChild(AddChild( _
                            AddNode(ucsIrtBlock), _
                                AddNode(ucsIrtStmtIncrPos, 1)), _
                                lSuccess)), _
                            lFailure))
            End If
        Case Else
            m_uIr.LastError = "Unsupported parse tree node type " & .NodeType(lNode)
            GoTo QH
        End Select
    End With
QH:
    Exit Function
EH:
    Err.Raise Err.Number, "cIR.pvCodeGenNode" & vbCrLf & Err.Source, Err.Description
    Resume
End Function

Private Function pvCodeGenReplaceLabels(ByVal lNode As Long) As Long
    Dim oCol            As Collection
    Dim cLab            As Collection
    Dim vElem           As Variant
    Dim lNewNode        As Long
    Dim cMatch          As Collection
    Dim cRev            As Collection
    Dim lIdx            As Long
    
    On Error GoTo EH
    pvCodeGenReplaceLabels = lNode
    If lNode <> IR_CONTINUE Then
        Set oCol = New Collection
        pvEnumChildNodes lNode, oCol
        Set cLab = New Collection
        For Each vElem In oCol
            With m_uIr.Nodes(vElem)
                If .NodeType = ucsIrtDeclLabel Then
                    cLab.Add "L" & pvCodeGenId, .Text
                End If
            End With
        Next
        If cLab.Count > 0 Then
            Set cMatch = New Collection
            Set cRev = New Collection
            For Each vElem In oCol
                If Not SearchCollection(cMatch, "#" & vElem) Then
                    lNewNode = AddNode(m_uIr.Nodes(vElem).NodeType)
                    cRev.Add vElem, "#" & lNewNode
                    cMatch.Add lNewNode, "#" & vElem
                End If
            Next
            For Each vElem In cMatch
                m_uIr.Nodes(vElem) = m_uIr.Nodes(cRev.Item("#" & vElem))
                With m_uIr.Nodes(vElem)
                    Select Case .NodeType
                    Case ucsIrtDeclLabel, ucsIrtStmtGoTo
                        If SearchCollection(cLab, .Text) Then
                            .Text = cLab.Item(.Text)
                        End If
                    End Select
                    For lIdx = 0 To .Count - 1
                        If SearchCollection(cMatch, "#" & .Children(lIdx)) Then
                            .Children(lIdx) = cMatch.Item("#" & .Children(lIdx))
                        End If
                    Next
                End With
            Next
            pvCodeGenReplaceLabels = cMatch.Item("#" & lNode)
        End If
    End If
    Exit Function
EH:
    Err.Raise Err.Number, "cIR.pvCodeGenReplaceLabels" & vbCrLf & Err.Source, Err.Description
    Resume
End Function

Private Function pvEnumChildNodes(ByVal lNode As Long, Optional oCol As Collection) As Long
    Dim lIdx            As Long
    
    If Not oCol Is Nothing Then
        If Not SearchCollection(oCol, "#" & lNode) Then
            oCol.Add lNode, "#" & lNode
        Else
            oCol.Add lNode
        End If
    End If
    pvEnumChildNodes = 1
    With m_uIr.Nodes(lNode)
        For lIdx = 0 To .Count - 1
            pvEnumChildNodes = pvEnumChildNodes + pvEnumChildNodes(.Children(lIdx), oCol)
        Next
    End With
End Function

Private Function pvCodeGenAction(sText As String, ByVal lRule As Long, Optional ByVal IsImmediate As Boolean) As String
    Dim lIdx            As Long
    Dim cOutput         As Collection
    Dim oMatch          As Object
    Dim lPos            As Long
    Dim oVars           As Object
    
    Set oVars = CreateObject("Scripting.Dictionary")
    oVars.CompareMode = vbTextCompare
    oVars.Item("$$") = ".VarResult"
    If IsImmediate Then
        oVars.Item("$$text") = "Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)"
    Else
        oVars.Item("$$text") = "Mid$(.Contents, lOffset, lSize)"
    End If
    With m_uIr.CodeGenTree
        For lIdx = 0 To .NodeNumVariables(lRule) - 1
            lPos = .NodeVariable(lRule, lIdx)
            oVars.Item("$" & .NodeText(lPos)) = ".VarStack(.VarPos - " & .NodeId(lPos) & ")"
        Next
    End With
    lPos = 1
    Set cOutput = New Collection
    For Each oMatch In pvCodeGenEnumVars(sText)
        cOutput.Add Mid$(sText, lPos, oMatch.FirstIndex - lPos + 1)
        If oVars.Exists(oMatch.Value) Then
            cOutput.Add oVars.Item(oMatch.Value)
        Else
            cOutput.Add oMatch.Value
        End If
        lPos = oMatch.FirstIndex + oMatch.Length + 1
    Next
    cOutput.Add Mid$(sText, lPos)
    pvCodeGenAction = ConcatCollection(cOutput)
End Function

Public Function pvCodeGenEnumVars(sAction As String) As Object
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "\$[\$\w_][\d\w_]*"
        Set pvCodeGenEnumVars = .Execute(sAction)
    End With
End Function

Private Function pvCodeGenId() As Long
    m_uIr.CodeGenId = m_uIr.CodeGenId + 1
    pvCodeGenId = m_uIr.CodeGenId
End Function

Private Function pvCodeGenFindCtxUse(ByVal lNode As Long) As Boolean
    Dim lIdx            As Long
    
    With m_uIr.Nodes(lNode)
        '--- note: euristic for emitted VB6 code that uses context member var
        If Left$(.Text, 1) = "." Or .NodeType = ucsIrtStmtSavePos Or .NodeType = ucsIrtStmtRestorePos Or .NodeType = ucsIrtStmtCapture Then
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

Private Function pvCodeGenTrimLastExit(ByVal lNode As Long) As Boolean
    With m_uIr.Nodes(lNode)
        If .Count > 0 Then
            If m_uIr.Nodes(.Children(.Count - 1)).NodeType = ucsIrtStmtExit _
                    And LenB(m_uIr.Nodes(.Children(.Count - 1)).Text) = 0 Then
                m_uIr.Nodes(lNode).Count = .Count - 1
                pvCodeGenTrimLastExit = True
            End If
        End If
    End With
    With m_uIr.Nodes(lNode)
        If .Count > 0 Then
            If pvCodeGenTrimLastExit(.Children(.Count - 1)) Then
                pvCodeGenTrimLastExit = True
            End If
        End If
    End With
End Function

Private Function pvCodeGenTrimGotoNext(ByVal lNode As Long) As Boolean
    Dim oCol            As Collection
    Dim lIdx            As Long
    Dim cLab            As Collection
    
    Set oCol = New Collection
    pvEnumChildNodes lNode, oCol
    Set cLab = New Collection
    For lIdx = 1 To oCol.Count
        With m_uIr.Nodes(oCol.Item(lIdx))
            If .NodeType = ucsIrtDeclLabel Then
                cLab.Add lIdx, .Text
            End If
        End With
    Next
    If cLab.Count > 0 Then
        For lIdx = 1 To oCol.Count
            With m_uIr.Nodes(oCol.Item(lIdx))
                If .NodeType = ucsIrtStmtGoTo Then
                    If cLab.Item(.Text) = lIdx + 1 Then
    '                    Debug.Print "[" & oCol.Item(lIdx) & "] Remove GoTo " & .Text
                        .NodeType = ucsIrtBlock
                        .Text = vbNullString
                    End If
                End If
            End With
        Next
        For lIdx = 1 To oCol.Count
            With m_uIr.Nodes(oCol.Item(lIdx))
                If .NodeType = ucsIrtStmtGoTo Then
                    If SearchCollection(cLab, .Text) Then
                        cLab.Remove .Text
                    End If
                End If
            End With
        Next
        For lIdx = 1 To cLab.Count
            With m_uIr.Nodes(oCol.Item(cLab.Item(lIdx)))
                .NodeType = ucsIrtBlock
                .Text = vbNullString
            End With
        Next
        For lIdx = 1 To oCol.Count
            With m_uIr.Nodes(oCol.Item(lIdx))
                If .NodeType = ucsIrtStmtIf Then
                    If .Count = 2 Then
                        If m_uIr.Nodes(.Children(0)).NodeType = ucsIrtExprCall Then
                            If pvEnumChildNodes(.Children(1)) = 1 Then
                                If m_uIr.Nodes(.Children(1)).NodeType = ucsIrtBlock Then
'                                    Debug.Print "[" & oCol.Item(lIdx) & "] Transform StmtIf to StmtCall " & m_uIr.Nodes(.Children(0)).Text
                                    .NodeType = ucsIrtStmtCall
                                    .Text = m_uIr.Nodes(.Children(0)).Text
                                    .Count = 0
                                    Erase .Children
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Next
    End If
End Function

Private Function pvCodeGenIsFinalJump(ByVal lNode As Long) As Boolean
    With m_uIr.Nodes(lNode)
        Select Case .NodeType
        Case ucsIrtStmtExit, ucsIrtStmtDoExit, ucsIrtStmtForExit, ucsIrtStmtGoTo
            pvCodeGenIsFinalJump = True
        Case Else
            If .Count > 0 Then
                pvCodeGenIsFinalJump = pvCodeGenIsFinalJump(.Children(.Count - 1))
            End If
        End Select
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

Private Function pvToVbCaseRange(ByVal sText As String, ByVal bIgnoreCase As Boolean) As String
    Dim lIdx            As Long
    Dim sUpper          As String
    
    If bIgnoreCase Then
        sText = LCase$(sText)
        sUpper = UCase$(sText)
    End If
    For lIdx = 1 To Len(sText)
        pvToVbCaseRange = IIf(LenB(pvToVbCaseRange) <> 0, pvToVbCaseRange & ", ", vbNullString) & Asc(Mid$(sText, lIdx, 1))
        If Mid$(sText, lIdx + 1, 1) = "-" And Len(sText) > lIdx + 1 Then
            pvToVbCaseRange = pvToVbCaseRange & " To " & Asc(Mid$(sText, lIdx + 2, 1))
            lIdx = lIdx + 2
        End If
    Next
    If bIgnoreCase Then
        For lIdx = 1 To Len(sText)
            If Mid$(sText, lIdx + 1, 1) = "-" And Len(sText) > lIdx + 1 Then
                If Mid$(sText, lIdx, 1) <> Mid$(sUpper, lIdx, 1) Or Mid$(sText, lIdx + 2, 1) <> Mid$(sUpper, lIdx + 2, 1) Then
                    pvToVbCaseRange = pvToVbCaseRange & ", " & Asc(Mid$(sUpper, lIdx, 1)) & " To " & Asc(Mid$(sUpper, lIdx + 2, 1))
                End If
                lIdx = lIdx + 2
            Else
                If Mid$(sText, lIdx, 1) <> Mid$(sUpper, lIdx, 1) Then
                    pvToVbCaseRange = pvToVbCaseRange & ", " & Asc(Mid$(sUpper, lIdx, 1))
                End If
            End If
        Next
    End If
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
            ByVal ePublic As VbTriState, _
            sModuleName As String, _
            sUserData As String, _
            sOutput As String) As Boolean
    Dim lIdx            As Long
    Dim sVarData        As String
    
    If C_Bool(SettingValue("public")) Or C_Bool(SettingValue("private")) Then
        ePublic = IIf(C_Bool(SettingValue("public")), vbTrue, vbFalse)
    End If
    sModuleName = Zn(Zn(SettingValue("modulename"), sModuleName), IIf(ePublic <> vbUseDefault, "cParser", "mdParser"))
    sUserData = Zn(Zn(SettingValue("userdatatype"), sUserData), "Variant")
    sVarData = Zn(SettingValue("vardatatype"), "Variant")
    With m_uIr
        Set .OutBuilder = New Collection
        .OutIndent = 0
        Set .OutDictionary = CreateObject("Scripting.Dictionary")
        .OutDictionary.CompareMode = vbTextCompare
        If ePublic <> vbUseDefault Then
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
"Attribute VB_Exposed = " & (ePublic = vbTrue)
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
"Private Const LOCALE_USER_DEFAULT           As Long = &H400" & vbCrLf & _
"Private Const NORM_IGNORECASE               As Long = 1" & vbCrLf & _
"Private Const CSTR_EQUAL                    As Long = 2" & vbCrLf
        pvOutput _
"Private Declare Sub CopyMemory Lib ""kernel32"" Alias ""RtlMoveMemory"" (Destination As Any, Source As Any, ByVal Length As Long)" & vbCrLf & _
"Private Declare Function RtlCompareMemory Lib ""ntdll"" (Source1 As Any, Source2 As Any, ByVal Length As Long) As Long" & vbCrLf & _
"Private Declare Function CompareStringW Lib ""kernel32"" (ByVal Locale As Long, ByVal dwCmpFlags As Long, lpString1 As Any, ByVal cchCount1 As Long, lpString2 As Any, ByVal cchCount2 As Long) As Long" & vbCrLf
        pvOutput _
"'=========================================================================" & vbCrLf & _
"' Constants and member variables" & vbCrLf & _
"'=========================================================================" & vbCrLf
        pvOutput _
"Private Const LNG_MAXINT            As Long = 2 ^ 31 - 1" & vbCrLf
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
"    Action              As Long" & vbCrLf & _
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
"    LastBufPos          As Long" & vbCrLf & _
"    UserData            As " & sUserData
        If .CodeGenVarStack Then
            pvOutput _
"    VarResult           As " & sVarData & vbCrLf & _
"    VarStack()          As " & sVarData & vbCrLf & _
"    VarPos              As Long"
        End If
        If LenB(SettingValue("members")) <> 0 Then
            pvOutput _
"    '-- from @members setting"
            pvOutput pvTrimCrLf(SettingValue("members"))
        End If
        pvOutput _
"End Type" & vbCrLf
        pvOutput _
"Private " & STR_CTX_VAR_NAME & "                     As UcsParserType" & vbCrLf
        pvOutput _
"'=========================================================================" & vbCrLf & _
"' Properties" & vbCrLf & _
"'=========================================================================" & vbCrLf
        pvOutput _
"Property Get " & .OutModulePrefix & "LastError() As String" & vbCrLf & _
"    " & .OutModulePrefix & "LastError = " & STR_CTX_VAR_NAME & ".LastError" & vbCrLf & _
"End Property" & vbCrLf
        pvOutput _
"Property Get " & .OutModulePrefix & "LastBufPos() As Long" & vbCrLf & _
"    " & .OutModulePrefix & "LastBufPos = " & STR_CTX_VAR_NAME & ".LastBufPos" & vbCrLf & _
"End Property" & vbCrLf
        pvOutput _
"Property Get " & .OutModulePrefix & "ParserVersion() As String" & vbCrLf & _
"    " & .OutModulePrefix & "ParserVersion = """ & Now & """" & vbCrLf & _
"End Property" & vbCrLf
        pvOutput _
"'=========================================================================" & vbCrLf & _
"' Methods" & vbCrLf & _
"'=========================================================================" & vbCrLf
        If LenB(.CodeGenStartFunc) <> 0 Then
            pvOutput _
"Public Function " & .OutModulePrefix & "Match(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As " & sUserData & IIf(.CodeGenVarStack, ", Optional Result As " & sVarData, vbNullString) & ") As Long" & vbCrLf & _
"    If " & .OutModulePrefix & "BeginMatch(sSubject, StartPos, UserData) Then" & vbCrLf & _
"        If " & .OutModulePrefix & .CodeGenStartFunc & "() Then" & vbCrLf & _
"            " & .OutModulePrefix & "Match = " & .OutModulePrefix & "EndMatch(UserData" & IIf(.CodeGenVarStack, ", Result", vbNullString) & ")" & vbCrLf & _
"        ElseIf LenB(" & STR_CTX_VAR_NAME & ".LastError) = 0 Then" & vbCrLf & _
"            " & STR_CTX_VAR_NAME & ".LastError = ""Fail""" & vbCrLf & _
"        End If" & vbCrLf & _
"    End If" & vbCrLf & _
"End Function" & vbCrLf
        End If
        pvOutput _
"Public Function " & .OutModulePrefix & "BeginMatch(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As " & sUserData & ") As Boolean" & vbCrLf & _
"    With " & STR_CTX_VAR_NAME & vbCrLf & _
"        .LastBufPos = 0" & vbCrLf & _
"        If LenB(sSubject) = 0 Then" & vbCrLf & _
"            .LastError = ""Cannot match empty input""" & vbCrLf & _
"            Exit Function" & vbCrLf & _
"        End If" & vbCrLf & _
"        .Contents = sSubject" & vbCrLf & _
"        ReDim .BufData(0 To Len(sSubject) + 3) As Integer" & vbCrLf & _
"        Call CopyMemory(.BufData(0), ByVal StrPtr(sSubject), LenB(sSubject))" & vbCrLf & _
"        .BufPos = StartPos" & vbCrLf & _
"        .BufSize = Len(sSubject)" & vbCrLf & _
"        .BufData(.BufSize) = -1 '-- EOF anchor" & vbCrLf & _
"        ReDim .ThunkData(0 To 4) As UcsParserThunkType" & vbCrLf & _
"        .ThunkPos = 0" & vbCrLf & _
"        .CaptureBegin = 0" & vbCrLf & _
"        .CaptureEnd = 0"
        pvOutput _
"        If IsObject(UserData) Then" & vbCrLf & _
"            Set .UserData = UserData" & vbCrLf & _
"        Else" & vbCrLf & _
"            .UserData = UserData" & vbCrLf & _
"        End If" & vbCrLf & _
"    End With" & vbCrLf & _
"    " & .OutModulePrefix & "BeginMatch = True" & vbCrLf & _
"End Function" & vbCrLf
        If .IrActionsEnum = 0 Then
            pvOutput _
"Public Function " & .OutModulePrefix & "EndMatch(Optional UserData As " & sUserData & ") As Long" & vbCrLf & _
"    Dim uEmpty          As UcsParserType" & vbCrLf & _
"    " & vbCrLf & _
"    " & .OutModulePrefix & "EndMatch = " & STR_CTX_VAR_NAME & ".BufPos" & vbCrLf & _
"    uEmpty.LastError = " & STR_CTX_VAR_NAME & ".LastError" & vbCrLf & _
"    uEmpty.LastBufPos = " & STR_CTX_VAR_NAME & ".LastBufPos" & vbCrLf & _
"    " & STR_CTX_VAR_NAME & " = uEmpty" & vbCrLf & _
"End Function" & vbCrLf
        Else
            pvOutput _
"Public Function " & .OutModulePrefix & "EndMatch(Optional UserData As " & sUserData & IIf(.CodeGenVarStack, ", Optional Result As " & sVarData, vbNullString) & ") As Long" & vbCrLf & _
"    Dim lIdx            As Long" & vbCrLf & _
"    Dim uEmpty          As UcsParserType" & vbCrLf & _
"    " & vbCrLf & _
"    With " & STR_CTX_VAR_NAME
            If .CodeGenVarStack Then
                pvOutput _
"        ReDim .VarStack(0 To 1024) As " & sVarData & vbCrLf & _
"        For lIdx = 0 To .ThunkPos - 1" & vbCrLf & _
"            Select Case .ThunkData(lIdx).Action" & vbCrLf & _
"            Case ucsActVarAlloc" & vbCrLf & _
"                .VarPos = .VarPos + .ThunkData(lIdx).CaptureBegin" & vbCrLf & _
"            Case ucsActVarSet" & vbCrLf & _
"                If IsObject(.VarResult) Then" & vbCrLf & _
"                    Set .VarStack(.VarPos - .ThunkData(lIdx).CaptureBegin) = .VarResult" & vbCrLf & _
"                Else" & vbCrLf & _
"                    .VarStack(.VarPos - .ThunkData(lIdx).CaptureBegin) = .VarResult" & vbCrLf & _
"                End If" & vbCrLf & _
"            Case ucsActResultClear" & vbCrLf & _
"                .VarResult = Empty" & vbCrLf & _
"            Case ucsActResultSet" & vbCrLf & _
"                With .ThunkData(lIdx)" & vbCrLf & _
"                    ctx.VarResult = Mid$(ctx.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)" & vbCrLf & _
"                End With" & vbCrLf & _
"            Case Else" & vbCrLf & _
"                With .ThunkData(lIdx)" & vbCrLf & _
"                    pvImplAction .Action, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin" & vbCrLf & _
"                End With" & vbCrLf & _
"            End Select" & vbCrLf & _
"        Next" & vbCrLf
                pvOutput _
"        If IsObject(.VarResult) Then" & vbCrLf & _
"            Set Result = .VarResult" & vbCrLf & _
"        Else" & vbCrLf & _
"            Result = .VarResult" & vbCrLf & _
"        End If"
            Else
                pvOutput _
"        For lIdx = 0 To .ThunkPos - 1" & vbCrLf & _
"            With .ThunkData(lIdx)" & vbCrLf & _
"                pvImplAction .Action, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin" & vbCrLf & _
"            End With" & vbCrLf & _
"        Next"
            End If
            pvOutput _
"        If IsObject(.UserData) Then" & vbCrLf & _
"            Set UserData = .UserData" & vbCrLf & _
"        Else" & vbCrLf & _
"            UserData = .UserData" & vbCrLf & _
"        End If" & vbCrLf & _
"        " & .OutModulePrefix & "EndMatch = .BufPos + 1" & vbCrLf & _
"    End With" & vbCrLf & _
"    uEmpty.LastError = " & STR_CTX_VAR_NAME & ".LastError" & vbCrLf & _
"    uEmpty.LastBufPos = " & STR_CTX_VAR_NAME & ".LastBufPos" & vbCrLf & _
"    " & STR_CTX_VAR_NAME & " = uEmpty" & vbCrLf & _
"End Function" & vbCrLf
            pvOutput _
"Private Sub pvPushAction(ByVal eAction As UcsParserActionsEnum)" & vbCrLf & _
"    pvPushThunk eAction, " & STR_CTX_VAR_NAME & ".CaptureBegin, " & STR_CTX_VAR_NAME & ".CaptureEnd" & vbCrLf & _
"End Sub" & vbCrLf
            pvOutput _
"Private Sub pvPushThunk(ByVal eAction As UcsParserActionsEnum, Optional ByVal lBegin As Long, Optional ByVal lEnd As Long)" & vbCrLf & _
"    With " & STR_CTX_VAR_NAME & vbCrLf & _
"        If UBound(.ThunkData) < .ThunkPos Then" & vbCrLf & _
"            ReDim Preserve .ThunkData(0 To 2 * UBound(.ThunkData)) As UcsParserThunkType" & vbCrLf & _
"        End If" & vbCrLf & _
"        With .ThunkData(.ThunkPos)" & vbCrLf & _
"            .Action = eAction" & vbCrLf & _
"            .CaptureBegin = lBegin" & vbCrLf & _
"            .CaptureEnd = lEnd" & vbCrLf & _
"        End With" & vbCrLf & _
"        .ThunkPos = .ThunkPos + 1" & vbCrLf & _
"    End With" & vbCrLf & _
"End Sub" & vbCrLf
        End If
        pvOutput _
"Private Function pvMatchString(sText As String) As Boolean" & vbCrLf & _
"    With " & STR_CTX_VAR_NAME & vbCrLf & _
"        If .BufPos + Len(sText) <= .BufSize Then" & vbCrLf & _
"            pvMatchString = RtlCompareMemory(.BufData(.BufPos), ByVal StrPtr(sText), LenB(sText)) = LenB(sText)" & vbCrLf & _
"        End If" & vbCrLf & _
"    End With" & vbCrLf & _
"End Function" & vbCrLf
        pvOutput _
"Private Function pvMatchIgnoreCase(sText As String) As Boolean" & vbCrLf & _
"    With ctx" & vbCrLf & _
"        If .BufPos + Len(sText) <= .BufSize Then" & vbCrLf & _
"            pvMatchIgnoreCase = CompareStringW(LOCALE_USER_DEFAULT, NORM_IGNORECASE, ByVal StrPtr(sText), Len(sText), .BufData(.BufPos), Len(sText)) = CSTR_EQUAL" & vbCrLf & _
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
        If LenB(SettingValue(vbNullString)) <> 0 Then
            pvOutput _
"'= user-defined functions ================================================" & vbCrLf
            pvOutput pvTrimCrLf(SettingValue(vbNullString))
        End If
        pvOutput vbNullString
        sOutput = ConcatCollection(.OutBuilder, vbCrLf)
        '--- success
        EmitCode = True
    End With
QH:
End Function

Private Function pvEmitStmt(ByVal lNode As Long) As Boolean
    Dim lIdx            As Long
    Dim vElem           As Variant
    Dim bDontIndent     As Boolean
    
    With m_uIr.Nodes(lNode)
        Select Case .NodeType
        Case ucsIrtBlock
            bDontIndent = True
            GoSub EmitChildren
        Case ucsIrtDeclFunc
            If (.Flags And ucsIrfPublic) <> 0 Then
                m_uIr.OutFuncScope = "Public"
                m_uIr.OutFuncName = m_uIr.OutModulePrefix & .Text
            Else
                m_uIr.OutFuncScope = "Private"
                m_uIr.OutFuncName = .Text
            End If
            If (.Flags And ucsIrfVoid) <> 0 Then
                m_uIr.OutFuncType = "Sub"
                pvOutput m_uIr.OutFuncScope & " Sub " & m_uIr.OutFuncName & IIf(Right$(.Text, 1) <> ")", "()", vbNullString)
            Else
                m_uIr.OutFuncType = "Function"
                pvOutput m_uIr.OutFuncScope & " Function " & m_uIr.OutFuncName & "() As Boolean"
            End If
            m_uIr.OutDictionary.RemoveAll
            pvEmitCollectDeclVar lNode
            m_uIr.OutIndent = m_uIr.OutIndent + 1
            For Each vElem In m_uIr.OutDictionary.Keys
                pvOutput "Dim " & vElem & " As " & m_uIr.OutDictionary.Item(vElem)
            Next
            If m_uIr.OutDictionary.Count > 0 Then
                pvOutput vbNullString
            End If
            m_uIr.OutIndent = m_uIr.OutIndent - 1
            GoSub EmitChildren
            pvOutput "End " & m_uIr.OutFuncType
        Case ucsIrtDeclEnum
            pvOutput "Private Enum " & .Text
            GoSub EmitChildren
            pvOutput "End Enum"
        Case ucsIrtDeclLabel
            '--- note: no indent
            m_uIr.OutBuilder.Add .Text & ":"
        Case ucsIrtStmtCall
            If pvEmitIsPublic(.Text) Then
                pvOutput "Call " & m_uIr.OutModulePrefix & .Text
            Else
                pvOutput "Call " & .Text
            End If
        Case ucsIrtStmtRetVal
            pvOutput m_uIr.OutFuncName & " = " & .Text
        Case ucsIrtStmtExit
            pvOutput "Exit " & m_uIr.OutFuncType
        Case ucsIrtStmtDoLoop
            pvOutput "Do"
            GoSub EmitChildren
            pvOutput "Loop"
        Case ucsIrtStmtDoExit
            pvOutput "Exit Do"
        Case ucsIrtStmtIf
            pvOutput "If " & pvEmitExpr(.Children(0)) & " Then", Comment:=.Comment
            m_uIr.OutIndent = m_uIr.OutIndent + 1
            If .Count > 1 Then
                If Not pvEmitStmt(.Children(1)) Then
                    GoTo QH
                End If
            Else
                pvOutput "'--- do nothing"
            End If
            m_uIr.OutIndent = m_uIr.OutIndent - 1
            If .Count > 2 Then
                pvOutput "Else"
                m_uIr.OutIndent = m_uIr.OutIndent + 1
                If Not pvEmitStmt(.Children(2)) Then
                    GoTo QH
                End If
                m_uIr.OutIndent = m_uIr.OutIndent - 1
            End If
            pvOutput "End If"
        Case ucsIrtStmtSelect
            pvOutput "Select Case " & .Text
            bDontIndent = True
            GoSub EmitChildren
            pvOutput "End Select"
        Case ucsIrtStmtCase
            pvOutput "Case " & .Text, Comment:=.Comment
            If .Count > 0 Then
                GoSub EmitChildren
            Else
                m_uIr.OutIndent = m_uIr.OutIndent + 1
                pvOutput "'--- do nothing"
                m_uIr.OutIndent = m_uIr.OutIndent - 1
            End If
        Case ucsIrtStmtSavePos
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
        Case ucsIrtStmtIncrPos
            If Left$(.Text, 1) = "-" Then
                pvOutput ".BufPos = .BufPos - " & Abs(.Text)
            Else
                pvOutput ".BufPos = .BufPos + " & .Text
            End If
        Case ucsIrtStmtEmit
            pvOutput .Text
        Case ucsIrtStmtWith
            pvOutput "With " & .Text
            GoSub EmitChildren
            pvOutput "End With"
        Case ucsIrtStmtCapture
            pvOutput .Text & " = .BufPos"
        Case ucsIrtStmtPushAction
            If m_uIr.OutDictionary.Exists("lCaptureBegin") Then
                pvOutput ".CaptureBegin = lCaptureBegin"
            End If
            If m_uIr.OutDictionary.Exists("lCaptureEnd") Then
                pvOutput ".CaptureEnd = lCaptureEnd"
            End If
            If LenB(.Text) <> 0 Then
                pvOutput "pvPushAction " & .Text
            End If
        Case ucsIrtStmtGoTo
            pvOutput "GoTo " & .Text
        Case ucsIrtStmtForLoop
            pvOutput "For " & .Text & " = 0 To LNG_MAXINT"
            GoSub EmitChildren
            pvOutput "Next"
        Case ucsIrtStmtForExit
            pvOutput "Exit For"
        Case Else
            m_uIr.LastError = "Unsupported IR node type " & .NodeType
            GoTo QH
        End Select
    End With
    '--- success
    pvEmitStmt = True
QH:
    Exit Function
EmitChildren:
    With m_uIr.Nodes(lNode)
        If Not bDontIndent Then
            m_uIr.OutIndent = m_uIr.OutIndent + 1
        End If
        For lIdx = 0 To .Count - 1
            If Not pvEmitStmt(.Children(lIdx)) Then
                GoTo QH
            End If
        Next
        If Not bDontIndent Then
            m_uIr.OutIndent = m_uIr.OutIndent - 1
        End If
    End With
    Return
End Function

Private Sub pvEmitCollectDeclVar(ByVal lNode As Long)
    Dim lIdx            As Long
    
    With m_uIr.Nodes(lNode)
        Select Case .NodeType
        Case ucsIrtStmtSavePos, ucsIrtStmtForLoop, ucsIrtStmtCapture
            m_uIr.OutDictionary.Item(.Text) = "Long"
        End Select
        For lIdx = 0 To .Count - 1
            pvEmitCollectDeclVar .Children(lIdx)
        Next
    End With
End Sub

Private Function pvEmitExpr(ByVal lNode As Long) As String
    With m_uIr.Nodes(lNode)
        Select Case .NodeType
        Case ucsIrtExprCall
            If pvEmitIsPublic(.Text) Then
                pvEmitExpr = m_uIr.OutModulePrefix & .Text & "()"
            Else
                pvEmitExpr = .Text & "()"
            End If
        Case ucsIrtExprNot
            pvEmitExpr = "Not " & pvEmitExpr(.Children(0))
        Case ucsIrtExprEmit
            pvEmitExpr = .Text
        End Select
    End With
End Function

Private Function pvEmitIsPublic(sText As String) As Boolean
    Dim lSize           As Long
    
    '--- ToDo: check func decl
    lSize = Len(STR_FUNC_PREFIX)
    If Left$(sText, lSize) = STR_FUNC_PREFIX And Mid$(sText, lSize + 1) <> UCase$(Mid$(sText, lSize + 1)) Then
        pvEmitIsPublic = True
    End If
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

Private Function pvTrimCrLf(ByVal sText As String) As String
    Do While Left$(sText, 2) = vbCrLf
        sText = Mid$(sText, 3)
    Loop
    Do While Right$(sText, 2) = vbCrLf
        sText = RTrim$(Left$(sText, Len(sText) - 2))
    Loop
    pvTrimCrLf = sText
End Function
