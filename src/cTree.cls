VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTree"
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
' cTree.cls - PEG grammar parse tree
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_TYPES             As String = "|Rule|Variable|Reference|Error|Dot|String|Class|Action|Predicate|Alternate|Sequence|PeekFor|PeekNot|Query|Star|Plus"
Private Const STR_FLAGS             As String = "Used|Reached|AlwaysTrue|MultiStep|SavePos|RestorePos|IsNot|CanInline|HasActions|SkipAdvance|IgnoreCase|HasAdvance"
Private Const NODE_ROOT             As Long = 0

Public Enum UcsTreeNodeTypeEnum
    ucsNdtRule = 1
    ucsNdtVariable
    ucsNdtReference
    ucsNdtError
    ucsNdtDot
    ucsNdtString
    ucsNdtClass
    ucsNdtAction
    ucsNdtPredicate
    ucsNdtAlternate
    ucsNdtSequence
    ucsNdtPeekFor
    ucsNdtPeekNot
    ucsNdtQuery
    ucsNdtStar
    ucsNdtPlus
End Enum

Public Enum UcsTreeNodeFlagsEnum
    ucsNdfUsed = 2 ^ 0
    ucsNdfReached = 2 ^ 1
    ucsNdfAlwaysTrue = 2 ^ 2
    ucsNdfMultiStep = 2 ^ 3
    ucsNdfSavePos = 2 ^ 4
    ucsNdfRestorePos = 2 ^ 5
    ucsNdfIsNot = 2 ^ 6
    ucsNdfCanInline = 2 ^ 7
    ucsNdfHasActions = 2 ^ 8
    ucsNdfSkipAdvance = 2 ^ 9
    ucsNdfIgnoreCase = 2 ^ 10
    ucsNdfHasAdvance = 2 ^ 11
End Enum

Private Type UcsTreeNodeType
    NodeType            As UcsTreeNodeTypeEnum
    Offset              As Long
    Text                As String
    Description         As String
    Children()          As Long
    Variables()         As Long
    Id                  As Long
    RefRule             As Long
    RefVar              As Long
    Flags               As UcsTreeNodeFlagsEnum
    RestoreFrom         As Long
End Type

Private Type UcsTreeType
    Nodes()             As UcsTreeNodeType
    Count               As Long
    ThisRule            As Long
    RulesCount          As Long
    ActionsCount        As Long
    StartRule           As Long
    StackData()         As Long
    StackIndex          As Long
    OutBuilder          As Collection
    OutNodeTypes        As Variant
    OutNodeFlags        As Variant
    LastError           As String
End Type

Private m_uTree                 As UcsTreeType
Private m_oSettings             As Object
Private m_vTag                  As Variant
Private m_sFileName             As String
Private m_laOffsets()           As Long

'=========================================================================
' Properties
'=========================================================================

Property Get LastError() As String
    LastError = m_uTree.LastError
End Property

Property Get RuleExpression(ByVal lRule As Long) As Long
    With m_uTree.Nodes(lRule)
        If .NodeType = ucsNdtRule Then
            RuleExpression = .Children(0)
        End If
    End With
End Property

Property Let RuleExpression(ByVal lRule As Long, ByVal lExpr As Long)
    With m_uTree
        Debug.Assert .Nodes(lRule).NodeType = ucsNdtRule
        .RulesCount = .RulesCount + 1
        .Nodes(lRule).Children(0) = lExpr
        .Nodes(lRule).Id = .RulesCount
        If .StartRule = 0 Or .Nodes(lRule).Text = SettingValue("start") Then
            .StartRule = lRule
        End If
    End With
End Property

Property Get RuleDescription(ByVal lRule As Long) As String
    With m_uTree.Nodes(lRule)
        If .NodeType = ucsNdtRule Then
            RuleDescription = .Description
        End If
    End With
End Property

Property Let RuleDescription(ByVal lRule As Long, sValue As String)
    With m_uTree
        Debug.Assert .Nodes(lRule).NodeType = ucsNdtRule
        .Nodes(lRule).Description = sValue
    End With
End Property

Property Get NodeType(ByVal lNode As Long) As UcsTreeNodeTypeEnum
    NodeType = m_uTree.Nodes(lNode).NodeType
End Property

Property Get NodeText(ByVal lNode As Long) As String
    NodeText = m_uTree.Nodes(lNode).Text
End Property

Property Let NodeText(ByVal lNode As Long, sValue As String)
    m_uTree.Nodes(lNode).Text = sValue
End Property

Property Get NodeChild(ByVal lNode As Long, ByVal lIdx As Long) As Long
    NodeChild = m_uTree.Nodes(lNode).Children(lIdx)
End Property

Property Get NodeChildren(ByVal lNode As Long) As Long()
    NodeChildren = m_uTree.Nodes(lNode).Children
End Property

Property Get NodeNumChildren(ByVal lNode As Long) As Long
    NodeNumChildren = UBound(m_uTree.Nodes(lNode).Children) + 1
End Property

Property Get NodeFlag(ByVal lNode As Long, ByVal eFlag As UcsTreeNodeFlagsEnum) As Boolean
    NodeFlag = (m_uTree.Nodes(lNode).Flags And eFlag) <> 0
End Property

Private Property Let NodeFlag(ByVal lNode As Long, ByVal eFlag As UcsTreeNodeFlagsEnum, ByVal bValue As Boolean)
    With m_uTree.Nodes(lNode)
        If bValue Then
            .Flags = .Flags Or eFlag
        Else
            .Flags = .Flags And Not eFlag
        End If
    End With
End Property

Property Get NodeRestoreFrom(ByVal lNode As Long) As Long
    NodeRestoreFrom = m_uTree.Nodes(lNode).RestoreFrom
End Property

Property Get NodeRefRule(ByVal lNode As Long) As Long
    NodeRefRule = m_uTree.Nodes(lNode).RefRule
End Property

Property Get NodeRefVar(ByVal lNode As Long) As Long
    NodeRefVar = m_uTree.Nodes(lNode).RefVar
End Property

Property Get NodeId(ByVal lNode As Long) As Long
    NodeId = m_uTree.Nodes(lNode).Id
End Property

Property Get NodeVariable(ByVal lNode As Long, ByVal lIdx As Long) As Long
    NodeVariable = m_uTree.Nodes(lNode).Variables(lIdx)
End Property

Property Get NodeNumVariables(ByVal lNode As Long) As Long
    NodeNumVariables = UBound(m_uTree.Nodes(lNode).Variables) + 1
End Property

Property Get StartRule() As Long
    StartRule = m_uTree.StartRule
End Property

Property Get SettingValue(sName As String) As String
    If Not m_oSettings Is Nothing Then
        If m_oSettings.Exists(sName) Then
            SettingValue = m_oSettings.Item(sName)
        End If
    End If
End Property

Property Let SettingValue(sName As String, sValue As String)
    If m_oSettings Is Nothing Then
        Set m_oSettings = CreateObject("Scripting.Dictionary")
    End If
    m_oSettings.Item(sName) = sValue
End Property

Property Get Tag() As Variant
    If IsObject(m_vTag) Then
        Set Tag = m_vTag
    Else
        Tag = m_vTag
    End If
End Property

Property Let Tag(vValue As Variant)
    m_vTag = vValue
End Property

Property Set Tag(vValue As Variant)
    Set m_vTag = vValue
End Property

'=========================================================================
' Methods
'=========================================================================

Private Function pvNewNode( _
            ByVal eNodeType As UcsTreeNodeTypeEnum, _
            ByVal lOffset As Long, _
            Optional Text As String, _
            Optional ByVal NumChildren As Long, _
            Optional RetVal As Long) As Long
    With m_uTree
        If .Count = 0 Then
            ReDim .Nodes(0 To 4) As UcsTreeNodeType
            '-- root node holds all the rules as children
            .Nodes(NODE_ROOT).Children = EmptyLongArray
            .Nodes(NODE_ROOT).Variables = EmptyLongArray
            .Count = 1
        ElseIf .Count > UBound(.Nodes) Then
            ReDim Preserve .Nodes(0 To 2 * UBound(.Nodes)) As UcsTreeNodeType
        End If
        With .Nodes(.Count)
            .NodeType = eNodeType
            .Offset = lOffset
            .Text = Text
            If NumChildren > 0 Then
                ReDim .Children(0 To NumChildren - 1) As Long
            Else
                .Children = EmptyLongArray
            End If
            .Variables = EmptyLongArray
        End With
        RetVal = .Count
        .Count = .Count + 1
    End With
    pvNewNode = RetVal
End Function

Public Function MakeRule(ByVal lOffset As Long, sName As String) As Long
    With m_uTree
        MakeRule = pvNewNode(ucsNdtRule, lOffset, Text:=sName, NumChildren:=1)
        AppendChild NODE_ROOT, MakeRule
    End With
End Function

Public Function FindRule(ByVal lOffset As Long, sName As String) As Long
    Dim vElem           As Variant
    
    With m_uTree
        If .Count > 0 Then
            For Each vElem In .Nodes(NODE_ROOT).Children
                FindRule = vElem
                If .Nodes(FindRule).Text = sName Then
                    Exit Function
                End If
            Next
        End If
    End With
    FindRule = MakeRule(lOffset, sName)
End Function

Public Function BeginRule(ByVal lRule As Long) As Long
    With m_uTree
        Debug.Assert .Nodes(lRule).NodeType = ucsNdtRule
        .ThisRule = lRule
        .ActionsCount = 0
        BeginRule = lRule
    End With
End Function

Public Sub SetExpression(ByVal lRule As Long, ByVal lNode As Long)
    RuleExpression(lRule) = lNode
End Sub

Public Function MakeVariable(ByVal lOffset As Long, sName As String) As Long
    Dim vElem           As Variant
    Dim lIdx            As Long
    
    With m_uTree
        Debug.Assert .ThisRule <> 0
        For Each vElem In .Nodes(.ThisRule).Variables
            If .Nodes(vElem).Text = sName Then
                MakeVariable = vElem
                GoTo QH
            End If
        Next
        MakeVariable = pvNewNode(ucsNdtVariable, lOffset, Text:=sName)
        .Nodes(MakeVariable).Id = UBound(.Nodes(.ThisRule).Variables) + 2
        With .Nodes(.ThisRule)
            lIdx = UBound(.Variables) + 1
            ReDim Preserve .Variables(0 To lIdx) As Long
            .Variables(lIdx) = MakeVariable
        End With
    End With
QH:
End Function

Public Function SetVariable(ByVal lOffset As Long, ByVal lNode As Long, ByVal lVariable As Long) As Long
    Dim lLastChild      As Long
    
    With m_uTree
        SetVariable = AppendSequence(MakeSequence(lOffset, MakePredicate(lOffset, "YY_CLEAR")), lNode)
        If NodeNumChildren(lNode) > 0 Then
            lLastChild = NodeChild(lNode, NodeNumChildren(lNode) - 1)
            If .Nodes(lLastChild).NodeType = ucsNdtPredicate And .Nodes(lLastChild).Text = "YY_END" Then
                lNode = MakePredicate(lOffset, "YY_SET")
                SetVariable = AppendSequence(SetVariable, lNode)
            End If
        End If
        .Nodes(lNode).RefVar = lVariable
    End With
End Function

Public Function MakeReference(ByVal lOffset As Long, sName As String) As Long
    MakeReference = pvNewRef(ucsNdtReference, lOffset, sName, FindRule(lOffset, sName))
End Function

Public Function MakeErrorReference(ByVal lOffset As Long, sName As String) As Long
    MakeErrorReference = pvNewRef(ucsNdtError, lOffset, sName, FindRule(lOffset, sName))
End Function

Public Function pvNewRef(ByVal eType As UcsTreeNodeTypeEnum, ByVal lOffset As Long, sName As String, ByVal lRule As Long) As Long
    With m_uTree
        pvNewRef = pvNewNode(eType, lOffset, Text:=sName)
        .Nodes(pvNewRef).RefRule = lRule
        .Nodes(lRule).Flags = .Nodes(lRule).Flags Or ucsNdfUsed
    End With
End Function

Public Function MakeDot(ByVal lOffset As Long) As Long
    MakeDot = pvNewNode(ucsNdtDot, lOffset)
End Function

Public Function MakeString(ByVal lOffset As Long, sText As String) As Long
    MakeString = pvNewNode(ucsNdtString, lOffset, Text:=sText)
    NodeFlag(MakeString, ucsNdfIgnoreCase) = C_Bool(SettingValue("ignorecase"))
End Function

Public Function MakeErrorString(ByVal lOffset As Long, sText As String) As Long
    MakeErrorString = pvNewNode(ucsNdtError, lOffset, Text:=sText)
End Function

Public Function MakeClass(ByVal lOffset As Long, sText As String) As Long
    MakeClass = pvNewNode(ucsNdtClass, lOffset, Text:=sText)
    NodeFlag(MakeClass, ucsNdfIgnoreCase) = C_Bool(SettingValue("ignorecase"))
End Function

Public Function SetIgnoreCase(ByVal lNode As Long, ByVal bValue As Boolean) As Long
    NodeFlag(lNode, ucsNdfIgnoreCase) = bValue
    SetIgnoreCase = lNode
End Function

Public Function MakeAction(ByVal lOffset As Long, sText As String) As Long
    With m_uTree
        Debug.Assert .ThisRule <> 0
        MakeAction = pvNewNode(ucsNdtAction, lOffset, Text:=sText)
        .ActionsCount = .ActionsCount + 1
        .Nodes(MakeAction).Id = .ActionsCount
        .Nodes(MakeAction).RefRule = .ThisRule
    End With
End Function

Public Function MakePredicate(ByVal lOffset As Long, sText As String) As Long
    MakePredicate = pvNewNode(ucsNdtPredicate, lOffset, Text:=sText)
End Function

Public Function MakeAlternate(ByVal lOffset As Long, ByVal lNode As Long) As Long
    With m_uTree
        If .Nodes(lNode).NodeType = ucsNdtAlternate Then
            MakeAlternate = lNode
        Else
            MakeAlternate = pvNewNode(ucsNdtAlternate, lOffset, NumChildren:=1)
            .Nodes(MakeAlternate).Children(0) = lNode
        End If
    End With
End Function

Public Function MakeSequence(ByVal lOffset As Long, ByVal lNode As Long) As Long
    With m_uTree
        If .Nodes(lNode).NodeType = ucsNdtSequence Then
            MakeSequence = lNode
        Else
            MakeSequence = pvNewNode(ucsNdtSequence, lOffset, NumChildren:=1)
            .Nodes(MakeSequence).Children(0) = lNode
        End If
    End With
End Function

Public Function AppendSequence(ByVal lParent As Long, ByVal lNode As Long) As Long
    Dim vElem           As Variant
    
    If m_uTree.Nodes(lNode).NodeType = ucsNdtSequence Then
        For Each vElem In m_uTree.Nodes(lNode).Children
            AppendSequence = AppendChild(lParent, vElem)
        Next
    Else
        AppendSequence = AppendChild(lParent, lNode)
    End If
End Function

Public Function AppendChild(ByVal lParent As Long, ByVal lNode As Long) As Long
    Dim lIdx            As Long
    
    With m_uTree.Nodes(lParent)
        lIdx = UBound(.Children) + 1
        ReDim Preserve .Children(0 To lIdx) As Long
        .Children(lIdx) = lNode
    End With
    AppendChild = lParent
End Function

Public Function MakePeekNot(ByVal lOffset As Long, ByVal lNode As Long) As Long
    With m_uTree
        MakePeekNot = pvNewNode(ucsNdtPeekNot, lOffset, NumChildren:=1)
        .Nodes(MakePeekNot).Children(0) = lNode
    End With
End Function

Public Function MakePeekFor(ByVal lOffset As Long, ByVal lNode As Long) As Long
    With m_uTree
        MakePeekFor = pvNewNode(ucsNdtPeekFor, lOffset, NumChildren:=1)
        .Nodes(MakePeekFor).Children(0) = lNode
    End With
End Function

Public Function MakeQuery(ByVal lOffset As Long, ByVal lNode As Long) As Long
    With m_uTree
        MakeQuery = pvNewNode(ucsNdtQuery, lOffset, NumChildren:=1)
        .Nodes(MakeQuery).Children(0) = lNode
    End With
End Function

Public Function MakeStar(ByVal lOffset As Long, ByVal lNode As Long) As Long
    With m_uTree
        MakeStar = pvNewNode(ucsNdtStar, lOffset, NumChildren:=1)
        .Nodes(MakeStar).Children(0) = lNode
    End With
End Function

Public Function MakePlus(ByVal lOffset As Long, ByVal lNode As Long) As Long
    With m_uTree
        MakePlus = pvNewNode(ucsNdtPlus, lOffset, NumChildren:=1)
        .Nodes(MakePlus).Children(0) = lNode
    End With
End Function

Public Sub SetSettingValue(ByVal lName As Long, sValue As String)
    Dim sName           As String
    
    sName = m_uTree.Nodes(lName).Text
    If LenB(sName) = 0 Then
        SettingValue(sName) = SettingValue(sName) & sValue
    Else
        SettingValue(sName) = sValue
    End If
End Sub

Public Function Push(ByVal lNode As Long) As Long
    With m_uTree
        If .StackIndex = 0 Then
            ReDim .StackData(0 To 1024) As Long
        End If
        .StackIndex = .StackIndex + 1
        .StackData(.StackIndex) = lNode
    End With
    Push = lNode
End Function

Public Function Pop() As Long
    With m_uTree
        Pop = .StackData(.StackIndex)
        .StackIndex = .StackIndex - 1
    End With
End Function

Public Function Top() As Long
    With m_uTree
        Top = .StackData(.StackIndex)
    End With
End Function

Public Function DumpParseTree(Optional ByVal StartNode As Long) As String
    With m_uTree
        Set .OutBuilder = New Collection
        .OutNodeTypes = Split(STR_TYPES, "|")
        .OutNodeFlags = Split(STR_FLAGS, "|")
        pvOutputNode StartNode, StartNode = NODE_ROOT
        DumpParseTree = ConcatCollection(.OutBuilder, vbCrLf)
    End With
End Function

Private Sub pvOutputNode(ByVal lNode As Long, ByVal lIndent As Long)
    Dim vElem           As Variant
    
    With m_uTree.Nodes(lNode)
        If .NodeType > 0 Then
            m_uTree.OutBuilder.Add Space$(4 * lIndent) & "[" & lNode & "] " & At(m_uTree.OutNodeTypes, .NodeType) & _
                IIf(.Offset <> 0, ", Offset=" & .Offset, vbNullString) & _
                IIf(LenB(.Text) <> 0, ", Text=" & .Text, vbNullString) & _
                IIf(LenB(.Description) <> 0, ", Description=" & .Description, vbNullString) & _
                IIf(.Id <> 0, ", Id=" & .Id, vbNullString) & _
                IIf(.RefRule <> 0, ", RefRule=[" & .RefRule & "]", vbNullString) & _
                IIf(.RefVar <> 0, ", RefVar=[" & .RefVar & "]", vbNullString) & _
                IIf(.Flags <> 0, ", Flags=(" & pvGetNodeFlags(.Flags) & ")", vbNullString) & _
                IIf(UBound(.Children) > 0, ", NumChildren=" & UBound(.Children) + 1, vbNullString) & _
                IIf(.RestoreFrom <> 0, ", RestoreFrom=[" & .RestoreFrom & "]", vbNullString)
        End If
        For Each vElem In .Children
            If vElem > NODE_ROOT Then
                pvOutputNode vElem, lIndent + 1
            Else
                m_uTree.OutBuilder.Add Space$(4 * (lIndent + 1)) & "Missing"
            End If
        Next
        For Each vElem In .Variables
            If vElem > NODE_ROOT Then
                pvOutputNode vElem, lIndent + 1
            Else
                m_uTree.OutBuilder.Add Space$(4 * (lIndent + 1)) & "Missing"
            End If
        Next
    End With
End Sub

Private Function pvGetNodeFlags(ByVal eFlags As UcsTreeNodeFlagsEnum) As String
    Dim vElem           As Variant
    Dim lIdx            As Long
    
    For Each vElem In m_uTree.OutNodeFlags
        If (eFlags And (2 ^ lIdx)) <> 0 Then
            pvGetNodeFlags = IIf(LenB(pvGetNodeFlags) <> 0, pvGetNodeFlags & "+", vbNullString) & vElem
        End If
        lIdx = lIdx + 1
    Next
End Function

Public Function CheckTree(cWarnings As Collection) As Boolean
    Dim vElem           As Variant
    Dim cResult         As Collection
    
    Set cResult = New Collection
    With m_uTree.Nodes(NODE_ROOT)
        For Each vElem In .Children
            With m_uTree.Nodes(vElem)
                If .Children(0) = 0 Then
                    m_uTree.LastError = Join(CalcLine(.Offset), ":") & ": error: Undefined rule '" & .Text & "'"
                    GoTo QH
                End If
            End With
            pvConsumeInput vElem, cResult
        Next
    End With
    If cResult.Count > 0 Then
        Set cWarnings = cResult
    End If
    '--- success
    CheckTree = True
QH:
End Function

Private Function pvConsumeInput(ByVal lNode As Long, cResult As Collection) As Boolean
    Dim lIdx            As Long
    
    Debug.Assert lNode <> 0
    With m_uTree.Nodes(lNode)
        Select Case .NodeType
        Case ucsNdtRule
            If (.Flags And ucsNdfReached) <> 0 Then
                If Not SearchCollection(cResult, "#" & lNode) Then
                    cResult.Add Array("Possible left recursion in rule '" & .Text & "'", .Offset), "#" & lNode
                End If
            Else
                .Flags = .Flags Or ucsNdfReached
                pvConsumeInput = pvConsumeInput(.Children(0), cResult)
                .Flags = .Flags And Not ucsNdfReached
            End If
        Case ucsNdtDot, ucsNdtClass
            pvConsumeInput = True
        Case ucsNdtString
            pvConsumeInput = LenB(.Text) <> 0
        Case ucsNdtPlus
            pvConsumeInput = pvConsumeInput(.Children(0), cResult)
        Case ucsNdtReference, ucsNdtError
            If .RefRule <> 0 Then
                pvConsumeInput = pvConsumeInput(.RefRule, cResult)
            End If
        Case ucsNdtAlternate
            For lIdx = 0 To UBound(.Children)
                If Not pvConsumeInput(.Children(lIdx), cResult) Then
                    GoTo QH
                End If
            Next
            pvConsumeInput = True
        Case ucsNdtSequence
            For lIdx = 0 To UBound(.Children)
                If pvConsumeInput(.Children(lIdx), cResult) Then
                    pvConsumeInput = True
                    Exit For
                End If
            Next
        End Select
    End With
QH:
End Function

Public Function OptimizeTree(cWarnings As Collection) As Boolean
    Dim vElem           As Variant
    Dim cResult         As Collection
    
    Do While pvUpdateHasActions(NODE_ROOT, False)
    Loop
    Do While pvOptimizePass1(NODE_ROOT)
    Loop
    Do While pvOptimizePass2(NODE_ROOT, False)
    Loop
    '--- for all rules re-adjust Used flag
    With m_uTree.Nodes(NODE_ROOT)
        For Each vElem In .Children
            With m_uTree.Nodes(vElem)
                .Flags = .Flags And Not ucsNdfUsed
            End With
        Next
    End With
    With m_uTree.Nodes(m_uTree.StartRule)
        .Flags = .Flags Or ucsNdfUsed
    End With
    pvUpdateUsed NODE_ROOT
    pvUpdateRestoreFrom NODE_ROOT, 0
    Do While pvUpdateHasActions(NODE_ROOT, False)
    Loop
    Set cResult = New Collection
    With m_uTree.Nodes(NODE_ROOT)
        For Each vElem In .Children
            With m_uTree.Nodes(vElem)
                If pvFindAdvance(vElem) Then
                    .Flags = .Flags Or ucsNdfHasAdvance
                End If
                If (.Flags And ucsNdfUsed) = 0 And (.Flags And ucsNdfCanInline) = 0 Then
                    cResult.Add Array("Unused rule '" & .Text & "'", .Offset)
                End If
            End With
        Next
    End With
    If cResult.Count > 0 Then
        Set cWarnings = cResult
    End If
    '--- success
    OptimizeTree = True
End Function

'--- sets AlwaysTrue and CanInline flags only
Private Function pvOptimizePass1(ByVal lNode As Long) As Boolean
    Dim vElem           As Variant
    Dim lFlags          As Long
    
    With m_uTree.Nodes(lNode)
        For Each vElem In .Children
            Debug.Assert vElem > NODE_ROOT
            If pvOptimizePass1(vElem) Then
                pvOptimizePass1 = True
            End If
        Next
        Select Case .NodeType
        Case ucsNdtStar, ucsNdtQuery, ucsNdtAction
            If (.Flags And ucsNdfAlwaysTrue) = 0 Then
                .Flags = .Flags Or ucsNdfAlwaysTrue
                pvOptimizePass1 = True
            End If
        Case ucsNdtString
            If LenB(.Text) = 0 Then
                If (.Flags And ucsNdfAlwaysTrue) = 0 Then
                    .Flags = .Flags Or ucsNdfAlwaysTrue
                    pvOptimizePass1 = True
                End If
            End If
        Case ucsNdtSequence, ucsNdtRule
            lFlags = .Flags Or ucsNdfAlwaysTrue
            For Each vElem In .Children
                If (m_uTree.Nodes(vElem).Flags And ucsNdfAlwaysTrue) = 0 Then
                    lFlags = lFlags And Not ucsNdfAlwaysTrue
                End If
            Next
            If .Flags <> lFlags Then
                .Flags = lFlags
                pvOptimizePass1 = True
            End If
            If .NodeType = ucsNdtRule Then
                If (.Flags And ucsNdfCanInline) = 0 Then
                    If UBound(m_uTree.Nodes(.Children(0)).Children) < 0 Then
                        .Flags = .Flags Or ucsNdfCanInline
                        pvOptimizePass1 = True
                    End If
                End If
            End If
        Case ucsNdtAlternate
            If (.Flags And ucsNdfAlwaysTrue) = 0 Then
                For Each vElem In .Children
                    If (m_uTree.Nodes(vElem).Flags And ucsNdfAlwaysTrue) <> 0 Then
                        .Flags = .Flags Or ucsNdfAlwaysTrue
                        pvOptimizePass1 = True
                        Exit For
                    End If
                Next
            End If
        Case ucsNdtReference, ucsNdtError
            If (.Flags And ucsNdfAlwaysTrue) = 0 And .RefRule <> 0 Then
                If (m_uTree.Nodes(.RefRule).Flags And ucsNdfAlwaysTrue) <> 0 Then
                    .Flags = .Flags Or ucsNdfAlwaysTrue
                    pvOptimizePass1 = True
                End If
            End If
        Case ucsNdtPredicate
            If (.Flags And ucsNdfAlwaysTrue) = 0 Then
                Select Case .Text
                Case "YY_BEGIN", "YY_END", "YY_CLEAR", "YY_SET"
                    .Flags = .Flags Or ucsNdfAlwaysTrue
                    pvOptimizePass1 = True
                End Select
            End If
        End Select
    End With
End Function

'--- manages Save/RestorePos flags
'--- inlines References to rules marked CanInline
'--- in Sequence nodes merges sequential String nodes
'--- in Alternate nodes merges sequential Class nodes
'--- transforms PeekNot -> Not PeekFor nodes
'--- transforms single char Class -> single char String
Private Function pvOptimizePass2(ByVal lNode As Long, ByVal bParentSavePos As Boolean) As Boolean
    Dim vElem           As Variant
    Dim lCount          As Long
    Dim lFlags          As Long
    Dim lRefVar         As Long
    Dim sText           As String
    Dim bChildSavePos   As Boolean
    Dim bNeedSavePos    As Boolean
    
    With m_uTree.Nodes(lNode)
        bChildSavePos = bParentSavePos Or (.Flags And ucsNdfSavePos) <> 0
        For Each vElem In .Children
            If pvOptimizePass2(vElem, bChildSavePos) Then
                pvOptimizePass2 = True
            End If
            bChildSavePos = False
        Next
        Select Case .NodeType
        Case ucsNdtSequence, ucsNdtRule
            '--- merge String nodes
            lFlags = m_uTree.Nodes(.Children(0)).Flags
            For Each vElem In .Children
                With m_uTree.Nodes(vElem)
                    If .NodeType <> ucsNdtString Or .Flags <> lFlags Then
                        sText = vbNullString
                        Exit For
                    End If
                    sText = sText & .Text
                End With
            Next
            If LenB(sText) <> 0 Then
                m_uTree.Nodes(lNode) = m_uTree.Nodes(.Children(0))
                .Text = sText
                pvOptimizePass2 = True
                Exit Function
            End If
            lFlags = .Flags
            For Each vElem In .Children
                If (m_uTree.Nodes(vElem).Flags And ucsNdfAlwaysTrue) = 0 Then
                    If (m_uTree.Nodes(vElem).Flags And ucsNdfHasActions) <> 0 Then
                        lCount = lCount + 1
                    Else
                        Select Case m_uTree.Nodes(vElem).NodeType
                        Case ucsNdtPeekFor, ucsNdtPeekNot
                            If (m_uTree.Nodes(vElem).Flags And ucsNdfSkipAdvance) = 0 Then
                                lCount = lCount + 1
                            End If
                        End Select
                    End If
                    If lCount > 0 Then
                        If bParentSavePos Then
                            lFlags = lFlags And Not ucsNdfMultiStep And Not ucsNdfSavePos
                        Else
                            lFlags = lFlags Or ucsNdfMultiStep Or ucsNdfSavePos
                        End If
                        With m_uTree.Nodes(vElem)
                            If (.Flags And ucsNdfRestorePos) = 0 Then
                                .Flags = .Flags Or ucsNdfRestorePos
                                pvOptimizePass2 = True
                            End If
                        End With
                    End If
                End If
                '--- buffer position non-moving nodes
                Select Case m_uTree.Nodes(vElem).NodeType
                Case ucsNdtPeekFor, ucsNdtPeekNot, ucsNdtPredicate
                Case Else
                    lCount = lCount + 1
                End Select
            Next
            If (.Flags And Not ucsNdfSavePos) <> (lFlags And Not ucsNdfSavePos) Then
                .Flags = lFlags
                pvOptimizePass2 = True
            End If
            If .NodeType = ucsNdtRule Then
                If (.Flags And ucsNdfCanInline) = 0 Then
                    If UBound(m_uTree.Nodes(.Children(0)).Children) < 0 Then
                        .Flags = .Flags Or ucsNdfCanInline
                        pvOptimizePass2 = True
                    End If
                End If
            End If
        Case ucsNdtAlternate
            For Each vElem In .Children
                If (m_uTree.Nodes(vElem).Flags And ucsNdfHasActions) <> 0 Then
                    If (m_uTree.Nodes(vElem).Flags And ucsNdfRestorePos) = 0 Then
                        With m_uTree.Nodes(vElem)
                            .Flags = .Flags Or ucsNdfRestorePos
                        End With
                        pvOptimizePass2 = True
                    End If
                End If
                lCount = lCount + 1
                If lCount > 1 Then
                    If (m_uTree.Nodes(vElem).Flags And ucsNdfRestorePos) <> 0 Then
                        bNeedSavePos = True
                    End If
                End If
            Next
            If bNeedSavePos Then
                If (.Flags And ucsNdfSavePos) = 0 Then
                    .Flags = .Flags Or ucsNdfSavePos Or ucsNdfMultiStep
                    pvOptimizePass2 = True
                End If
            ElseIf (.Flags And ucsNdfSavePos) = IIf(bParentSavePos, ucsNdfSavePos, 0) Then
                For Each vElem In .Children
                    If (m_uTree.Nodes(vElem).Flags And ucsNdfSavePos) <> 0 Then
                    If (m_uTree.Nodes(vElem).Flags And ucsNdfMultiStep) = 0 Then
                        With m_uTree.Nodes(vElem)
                            .Flags = .Flags And Not ucsNdfSavePos
                        End With
                        pvOptimizePass2 = True
                    End If
                    End If
                Next
            End If
            '--- merge Class nodes
            lFlags = m_uTree.Nodes(.Children(0)).Flags
            For Each vElem In .Children
                With m_uTree.Nodes(vElem)
                    If .NodeType <> ucsNdtClass Or .Flags <> lFlags Then
                        sText = vbNullString
                        Exit For
                    End If
                    sText = sText & .Text
                End With
            Next
            If LenB(sText) <> 0 Then
                m_uTree.Nodes(lNode) = m_uTree.Nodes(.Children(0))
                .Text = sText
                pvOptimizePass2 = True
            End If
        Case ucsNdtReference
            If (m_uTree.Nodes(.RefRule).Flags And ucsNdfCanInline) <> 0 Then
                lFlags = m_uTree.Nodes(lNode).Flags
                lRefVar = m_uTree.Nodes(lNode).RefVar
                m_uTree.Nodes(lNode) = m_uTree.Nodes(m_uTree.Nodes(.RefRule).Children(0))
                With m_uTree.Nodes(lNode)
                    .Flags = (.Flags And (ucsNdfIgnoreCase Or ucsNdfIsNot)) Or lFlags
                    .RefVar = lRefVar
                End With
                pvOptimizePass2 = True
            End If
        Case ucsNdtPeekNot
            .NodeType = ucsNdtPeekFor
            .Flags = .Flags Or ucsNdfIsNot
            pvOptimizePass2 = True
        Case ucsNdtClass
            If (.Flags And ucsNdfIsNot) = 0 Then
                If Left$(.Text, 1) = "^" Then
                    .Text = Mid$(.Text, 2)
                    .Flags = .Flags Or ucsNdfIsNot
                    pvOptimizePass2 = True
                End If
            End If
            If Len(.Text) = 1 Then
                .NodeType = ucsNdtString
                pvOptimizePass2 = True
            End If
        End Select
        If .NodeType = ucsNdtPeekFor And (.Flags And ucsNdfSkipAdvance) = 0 Then
            Select Case m_uTree.Nodes(.Children(0)).NodeType
            Case ucsNdtString, ucsNdtClass, ucsNdtDot
                .Flags = (.Flags Or ucsNdfSkipAdvance) And Not ucsNdfRestorePos
                With m_uTree.Nodes(.Children(0))
                    .Flags = .Flags Or ucsNdfSkipAdvance
                End With
                pvOptimizePass2 = True
            End Select
        End If
    End With
End Function

Private Function pvUpdateUsed(ByVal lNode As Long) As Boolean
    Dim vElem           As Variant

    With m_uTree.Nodes(lNode)
        For Each vElem In .Children
            If pvUpdateUsed(vElem) Then
                pvUpdateUsed = True
            End If
        Next
        Select Case .NodeType
        Case ucsNdtReference, ucsNdtError
            If .RefRule <> 0 Then
                If (m_uTree.Nodes(.RefRule).Flags And ucsNdfUsed) = 0 Then
                    m_uTree.Nodes(.RefRule).Flags = m_uTree.Nodes(.RefRule).Flags Or ucsNdfUsed
                    pvUpdateUsed = True
                End If
            End If
        End Select
    End With
End Function

Private Function pvUpdateRestoreFrom(ByVal lNode As Long, ByVal lRestoreFrom As Long) As Boolean
    Dim vElem           As Variant
    
    With m_uTree.Nodes(lNode)
        If (.Flags And ucsNdfRestorePos) <> 0 Then
            If .RestoreFrom <> lRestoreFrom Then
                .RestoreFrom = lRestoreFrom
                pvUpdateRestoreFrom = True
            End If
        End If
        If (.Flags And ucsNdfSavePos) <> 0 Then
            lRestoreFrom = lNode
        End If
        For Each vElem In .Children
            If pvUpdateRestoreFrom(vElem, lRestoreFrom) Then
                pvUpdateRestoreFrom = True
            End If
        Next
    End With
End Function

Private Function pvUpdateHasActions(ByVal lNode As Long, bHasActions As Boolean) As Boolean
    Dim vElem           As Variant

    With m_uTree.Nodes(lNode)
        Select Case .NodeType
        Case ucsNdtRule
            bHasActions = False
        End Select
        For Each vElem In .Children
            If pvUpdateHasActions(vElem, bHasActions) Then
                pvUpdateHasActions = True
            End If
        Next
        Select Case .NodeType
        Case ucsNdtAction
            bHasActions = True
        Case ucsNdtRule
            If bHasActions Then
                If (.Flags And ucsNdfHasActions) = 0 Then
                    .Flags = .Flags Or ucsNdfHasActions
                    pvUpdateHasActions = True
                End If
            End If
        Case ucsNdtReference, ucsNdtError
            If .RefRule <> 0 Then
                If (m_uTree.Nodes(.RefRule).Flags And ucsNdfHasActions) <> 0 Then
                    bHasActions = True
                    If (.Flags And ucsNdfHasActions) = 0 Then
                        .Flags = .Flags Or ucsNdfHasActions
                        pvUpdateHasActions = True
                    End If
                End If
            End If
        End Select
        If bHasActions And (.Flags And ucsNdfSavePos) <> 0 Then
            If (.Flags And ucsNdfHasActions) = 0 Then
                .Flags = .Flags Or ucsNdfHasActions
                pvUpdateHasActions = True
            End If
        End If
        If .RestoreFrom <> 0 Then
            If (m_uTree.Nodes(.RestoreFrom).Flags And ucsNdfHasActions) <> 0 Then
                If (.Flags And ucsNdfHasActions) = 0 Then
                    .Flags = .Flags Or ucsNdfHasActions
                    pvUpdateHasActions = True
                End If
            End If
        End If
    End With
End Function

Private Function pvFindAdvance(ByVal lNode As Long) As Boolean
    Dim vElem           As Variant
    
    With m_uTree.Nodes(lNode)
        Select Case .NodeType
        Case ucsNdtString, ucsNdtClass
            pvFindAdvance = True
            Exit Function
        End Select
        For Each vElem In .Children
            If pvFindAdvance(vElem) Then
                pvFindAdvance = True
                Exit Function
            End If
        Next
    End With
End Function

Public Function ReadFile(sFile As String) As String
    m_sFileName = Mid$(sFile, InStrRev(sFile, "\") + 1)
    ReadFile = ReadTextFile(sFile)
    pvBuildLineInfo ReadFile
End Function

Private Sub pvBuildLineInfo(sSubject As String)
    Dim lIdx            As Long
    
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "\r?\n"
        With .Execute(sSubject)
            ReDim m_laOffsets(0 To .Count) As Long
            For lIdx = 0 To .Count - 1
                With .Item(lIdx)
                    m_laOffsets(lIdx + 1) = .FirstIndex + .Length
                End With
            Next
        End With
    End With
End Sub

Public Function CalcLine(ByVal lOffset As Long) As Variant
    Dim lLower          As Long
    Dim lUpper          As Long
    Dim lMiddle         As Long
    
    lUpper = UBound(m_laOffsets)
    Do While lLower < lUpper
        lMiddle = (lLower + lUpper + 1) \ 2
        If m_laOffsets(lMiddle) < lOffset Then
            lLower = lMiddle
        Else
            lUpper = lMiddle - 1
        End If
    Loop
    CalcLine = Array(m_sFileName, lUpper + 1, lOffset - m_laOffsets(lUpper))
End Function
