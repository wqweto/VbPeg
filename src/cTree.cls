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

Private Const STR_TYPES             As String = "|Rule|Reference|Dot|String|Class|Action|Predicate|Alternate|Sequence|PeekFor|PeekNot|Query|Star|Plus"
Private Const NODE_ROOT             As Long = 0

Public Enum UcsTreeNodeTypeEnum
    ucsNdtRule = 1
'    ucsNdtVariable
    ucsNdtReference
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
End Enum

Private Type UcsTreeNodeType
    NodeType            As UcsTreeNodeTypeEnum
    Text                As String
    Children()          As Long
    Id                  As Long
    Rule                As Long
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
    LastError           As String
End Type

Private m_uTree                 As UcsTreeType

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
        If .StartRule = 0 Or .Nodes(lRule).Text = "start" Then
            .StartRule = lRule
        End If
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

Property Get NodeRestoreFrom(ByVal lNode As Long) As Long
    NodeRestoreFrom = m_uTree.Nodes(lNode).RestoreFrom
End Property

Property Get NodeRule(ByVal lNode As Long) As Long
    NodeRule = m_uTree.Nodes(lNode).Rule
End Property

Property Get NodeId(ByVal lNode As Long) As Long
    NodeId = m_uTree.Nodes(lNode).Id
End Property

Property Get StartRule() As Long
    StartRule = m_uTree.StartRule
End Property

'=========================================================================
' Methods
'=========================================================================

Private Function pvNewNode( _
            ByVal eNodeType As UcsTreeNodeTypeEnum, _
            Optional sText As String, _
            Optional ByVal NumChildren As Long, _
            Optional RetVal As Long) As Long
    With m_uTree
        If .Count = 0 Then
            ReDim .Nodes(0 To 4) As UcsTreeNodeType
            '-- root node holds all the rules as children
            .Nodes(NODE_ROOT).Children = EmptyLongArray
            .Count = 1
        ElseIf .Count > UBound(.Nodes) Then
            ReDim Preserve .Nodes(0 To 2 * UBound(.Nodes)) As UcsTreeNodeType
        End If
        With .Nodes(.Count)
            .NodeType = eNodeType
            .Text = sText
            If NumChildren > 0 Then
                ReDim .Children(0 To NumChildren - 1) As Long
            Else
                .Children = EmptyLongArray
            End If
        End With
        RetVal = .Count
        .Count = .Count + 1
    End With
    pvNewNode = RetVal
End Function

Public Function MakeRule(sName As String) As Long
    With m_uTree
        MakeRule = pvNewNode(ucsNdtRule, sName, NumChildren:=1)
        AppendChild NODE_ROOT, MakeRule
    End With
End Function

Public Function FindRule(sName As String) As Long
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
    FindRule = MakeRule(sName)
End Function

Public Function BeginRule(ByVal lRule As Long) As Long
    With m_uTree
        Debug.Assert .Nodes(lRule).NodeType = ucsNdtRule
        .ThisRule = lRule
        .ActionsCount = 0
        BeginRule = lRule
    End With
End Function

'Public Function MakeVariable(sName As String) As Long
'    With m_uTree
'        Debug.Assert .ThisRule <> 0
'        MakeVariable = pvNewNode(ucsNdtVariable, sName)
'        AppendChild .ThisRule, MakeVariable
'    End With
'End Function

Public Function MakeReference(ByVal lRule As Long) As Long
    Dim sName           As String
    With m_uTree
        sName = .Nodes(lRule).Text
        MakeReference = pvNewNode(ucsNdtReference, sName)
        .Nodes(MakeReference).Rule = lRule
        .Nodes(lRule).Flags = .Nodes(lRule).Flags Or ucsNdfUsed
    End With
End Function

Public Function MakeDot() As Long
    MakeDot = pvNewNode(ucsNdtDot)
End Function

Public Function MakeString(sText As String) As Long
    MakeString = pvNewNode(ucsNdtString, sText)
End Function

Public Function MakeClass(sText As String) As Long
    MakeClass = pvNewNode(ucsNdtClass, sText)
End Function

Public Function MakeAction(sText As String) As Long
    With m_uTree
        Debug.Assert .ThisRule <> 0
        MakeAction = pvNewNode(ucsNdtAction, sText)
        .ActionsCount = .ActionsCount + 1
        .Nodes(MakeAction).Id = .ActionsCount
        .Nodes(MakeAction).Rule = .ThisRule
    End With
End Function

Public Function MakePredicate(sText As String) As Long
    MakePredicate = pvNewNode(ucsNdtPredicate, sText)
End Function

Public Function MakeAlternate(ByVal lNode As Long) As Long
    With m_uTree
        If .Nodes(lNode).NodeType = ucsNdtAlternate Then
            MakeAlternate = lNode
        Else
            MakeAlternate = pvNewNode(ucsNdtAlternate, NumChildren:=1)
            .Nodes(MakeAlternate).Children(0) = lNode
        End If
    End With
End Function

Public Function MakeSequence(ByVal lNode As Long) As Long
    With m_uTree
        If .Nodes(lNode).NodeType = ucsNdtSequence Then
            MakeSequence = lNode
        Else
            MakeSequence = pvNewNode(ucsNdtSequence, NumChildren:=1)
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

Public Function MakePeekNot(ByVal lNode As Long) As Long
    With m_uTree
        MakePeekNot = pvNewNode(ucsNdtPeekNot, NumChildren:=1)
        .Nodes(MakePeekNot).Children(0) = lNode
    End With
End Function

Public Function MakePeekFor(ByVal lNode As Long) As Long
    With m_uTree
        MakePeekFor = pvNewNode(ucsNdtPeekFor, NumChildren:=1)
        .Nodes(MakePeekFor).Children(0) = lNode
    End With
End Function

Public Function MakeQuery(ByVal lNode As Long) As Long
    With m_uTree
        MakeQuery = pvNewNode(ucsNdtQuery, NumChildren:=1)
        .Nodes(MakeQuery).Children(0) = lNode
    End With
End Function

Public Function MakeStar(ByVal lNode As Long) As Long
    With m_uTree
        MakeStar = pvNewNode(ucsNdtStar, NumChildren:=1)
        .Nodes(MakeStar).Children(0) = lNode
    End With
End Function

Public Function MakePlus(ByVal lNode As Long) As Long
    With m_uTree
        MakePlus = pvNewNode(ucsNdtPlus, NumChildren:=1)
        .Nodes(MakePlus).Children(0) = lNode
    End With
End Function

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

Public Function DumpParseTree(Optional StartNode As Long) As String
    With m_uTree
        Set .OutBuilder = New Collection
        .OutNodeTypes = Split(STR_TYPES, "|")
        pvOutputNode StartNode, StartNode = NODE_ROOT
        DumpParseTree = ConcatCollection(.OutBuilder, vbCrLf)
    End With
End Function

Private Sub pvOutputNode(ByVal lNode As Long, ByVal lIndent As Long)
    Dim vElem           As Variant
    
    With m_uTree.Nodes(lNode)
        If .NodeType > 0 Then
            m_uTree.OutBuilder.Add Space$(4 * lIndent) & "[" & lNode & "] " & At(m_uTree.OutNodeTypes, .NodeType) & _
                IIf(LenB(.Text) <> 0, ", Text=" & .Text, vbNullString) & _
                IIf(.Id <> 0, ", Id=" & .Id, vbNullString) & _
                IIf(.Rule <> 0, ", Rule=[" & .Rule & "]", vbNullString) & _
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
    End With
End Sub

Private Function pvGetNodeFlags(ByVal eFlags As UcsTreeNodeFlagsEnum) As String
    If (eFlags And ucsNdfUsed) <> 0 Then
        pvGetNodeFlags = "Used"
    End If
    If (eFlags And ucsNdfReached) <> 0 Then
        pvGetNodeFlags = IIf(LenB(pvGetNodeFlags) <> 0, pvGetNodeFlags & "+", vbNullString) & "Reached"
    End If
    If (eFlags And ucsNdfAlwaysTrue) <> 0 Then
        pvGetNodeFlags = IIf(LenB(pvGetNodeFlags) <> 0, pvGetNodeFlags & "+", vbNullString) & "AlwaysTrue"
    End If
    If (eFlags And ucsNdfMultiStep) <> 0 Then
        pvGetNodeFlags = IIf(LenB(pvGetNodeFlags) <> 0, pvGetNodeFlags & "+", vbNullString) & "MultiStep"
    End If
    If (eFlags And ucsNdfSavePos) <> 0 Then
        pvGetNodeFlags = IIf(LenB(pvGetNodeFlags) <> 0, pvGetNodeFlags & "+", vbNullString) & "SavePos"
    End If
    If (eFlags And ucsNdfRestorePos) <> 0 Then
        pvGetNodeFlags = IIf(LenB(pvGetNodeFlags) <> 0, pvGetNodeFlags & "+", vbNullString) & "RestorePos"
    End If
    If (eFlags And ucsNdfIsNot) <> 0 Then
        pvGetNodeFlags = IIf(LenB(pvGetNodeFlags) <> 0, pvGetNodeFlags & "+", vbNullString) & "IsNot"
    End If
    If (eFlags And ucsNdfCanInline) <> 0 Then
        pvGetNodeFlags = IIf(LenB(pvGetNodeFlags) <> 0, pvGetNodeFlags & "+", vbNullString) & "CanInline"
    End If
    If (eFlags And ucsNdfHasActions) <> 0 Then
        pvGetNodeFlags = IIf(LenB(pvGetNodeFlags) <> 0, pvGetNodeFlags & "+", vbNullString) & "HasActions"
    End If
    If (eFlags And ucsNdfSkipAdvance) <> 0 Then
        pvGetNodeFlags = IIf(LenB(pvGetNodeFlags) <> 0, pvGetNodeFlags & "+", vbNullString) & "SkipAdvance"
    End If
End Function

Public Function OptimizeTree() As Boolean
    Dim vElem           As Variant
    
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
    '--- success
    OptimizeTree = True
End Function

'--- sets AlwaysTrue and CanInline flags only
Private Function pvOptimizePass1(ByVal lNode As Long) As Boolean
    Dim vElem           As Variant
    Dim lFlags          As Long
    
    With m_uTree.Nodes(lNode)
        For Each vElem In .Children
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
        Case ucsNdtReference
            If (.Flags And ucsNdfAlwaysTrue) = 0 Then
                If (m_uTree.Nodes(.Rule).Flags And ucsNdfAlwaysTrue) <> 0 Then
                    .Flags = .Flags Or ucsNdfAlwaysTrue
                    pvOptimizePass1 = True
                End If
            End If
        Case ucsNdtPredicate
            If (.Flags And ucsNdfAlwaysTrue) = 0 Then
                Select Case .Text
                Case "YY_BEGIN", "YY_END"
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
    Dim sText           As String
    Dim bChildSavePos   As Boolean
    
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
                    Select Case m_uTree.Nodes(vElem).NodeType
                    Case ucsNdtPeekFor, ucsNdtPeekNot
'                        lFlags = lFlags Or ucsNdfMultiStep Or ucsNdfSavePos
                        If (m_uTree.Nodes(vElem).Flags And ucsNdfSkipAdvance) = 0 Then
                            lCount = lCount + 1
                        End If
                    End Select
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
            If (.Flags And ucsNdfSavePos) = IIf(bParentSavePos, ucsNdfSavePos, 0) Then
                For Each vElem In .Children
                    If (m_uTree.Nodes(vElem).Flags And ucsNdfSavePos) <> 0 Then
                        If bParentSavePos Then
                            .Flags = .Flags And Not ucsNdfSavePos
                        Else
                            .Flags = .Flags Or ucsNdfSavePos
                        End If
                        With m_uTree.Nodes(vElem)
                            .Flags = .Flags And Not ucsNdfSavePos
                        End With
                        pvOptimizePass2 = True
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
            If (m_uTree.Nodes(.Rule).Flags And ucsNdfCanInline) <> 0 Then
                lFlags = m_uTree.Nodes(lNode).Flags
                m_uTree.Nodes(lNode) = m_uTree.Nodes(m_uTree.Nodes(.Rule).Children(0))
                m_uTree.Nodes(lNode).Flags = lFlags
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
        Case ucsNdtReference
            If (m_uTree.Nodes(.Rule).Flags And ucsNdfUsed) = 0 Then
                m_uTree.Nodes(.Rule).Flags = m_uTree.Nodes(.Rule).Flags Or ucsNdfUsed
                pvUpdateUsed = True
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
        Case ucsNdtReference
            If (m_uTree.Nodes(.Rule).Flags And ucsNdfHasActions) <> 0 Then
                bHasActions = True
                If (.Flags And ucsNdfHasActions) = 0 Then
                    .Flags = .Flags Or ucsNdfHasActions
                    pvUpdateHasActions = True
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
