VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FEModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class must be used in conjunction with the following classes and modules:
'Node2D.cls
'Member2D.cls
'Segment.cls
'Matrix.cls
'MatrixMath.bas
'StaticCondensation.bas
'FERFunctions.bas

'Enforce explicit variable declarations in this module
Option Explicit

Public Nodes As Object       'A dictionary of all the nodes in the model with node names for keys
Private NodesByID As Object  'A dictionary of all the nodes in the model with node ID's for keys
Public Members As Object     'A dictionary of all the members in the finite element model
Public LoadCombos As Object  'A dictionary of the model's load combinations (Combo, ComboName)
Private Solved As Boolean    'Flag indicating whether the model has been solved or not

'Enumerations for this class
'Types of forces
Public Enum ForceType
    FX
    FY
    MZ
End Enum

'Types of displacements
Public Enum DispType
    DX
    DY
    RZ
End Enum

'Directions for member distributed loads
Public Enum LoadDir
    Transverse
    Axial
End Enum

'Adds a node to the model
Public Sub AddNode(NodeName As String, XCoord As Double, YCoord As Double)

    'Create the new node
    Dim NewNode As New Node2D
    NewNode.Name = NodeName
    NewNode.XCoord = XCoord
    NewNode.YCoord = YCoord
    
    'Add the node to the dictionary of nodes
    Call Nodes.Add(NodeName, NewNode)
    
    'Flag the model as unsolved
    Solved = False
    
End Sub

'Adds a member to the model
Public Sub AddMember(MemberName As String, iNode As String, jNode As String, Elasticity As Double, Inertia As Double, Area As Double)
    
    'Create the new member
    Dim NewMember As New Member2D
    NewMember.Name = MemberName
    NewMember.Elasticity = Elasticity
    NewMember.Inertia = Inertia
    NewMember.Area = Area
    
    'Link the member's load combinations to the model's load combinations
    Set NewMember.LoadCombos = LoadCombos
    
    'Assign the iNode and jNode to the member
    Set NewMember.iNode = Nodes(iNode)
    Set NewMember.jNode = Nodes(jNode)
    
    'Add the new member to the dictionary of members
    Call Members.Add(MemberName, NewMember)
    
    'Flag the model as unsolved
    Solved = False
    
End Sub

'Removes a node from the model
Public Sub RemoveNode(NodeName As String)
    
    'Search through each item in the collection
    Call Nodes.Remove(NodeName)
                
    'Flag the model as unsolved
    Solved = False
    
End Sub

'Edits the support conditions at any given node
Public Sub EditSupport(NodeName As String, SupportDX As Boolean, SupportDY As Boolean, SupportRZ As Boolean)
            
    'Add the support
    Nodes(NodeName).SupportDX = SupportDX
    Nodes(NodeName).SupportDY = SupportDY
    Nodes(NodeName).SupportRZ = SupportRZ
            
    'Flag the model as unsolved
    Solved = False
    
End Sub

'Edits the end release configuration for any given member
Public Sub EditEndReleases(MemberName As String, iRelease As Boolean, jRelease As Boolean)
            
    'Release the appropriate degrees of freedom
    Call Members(MemberName).AddRelease(3, iRelease)
    Call Members(MemberName).AddRelease(6, jRelease)
            
    'Flag the model as unsolved
    Solved = False
    
End Sub

'Adds a point load to a member
Public Sub AddMemberPointLoad(MemberName As String, P As Double, x As Double, Direction As LoadDir, Optional LoadCase = "Case 1")
    
    'Step through each member
    Dim Member As Member2D
    Set Member = Members(MemberName)
            
    'Add the member load
    If Direction = Transverse Then
        Call Member.AddPtLoad(P, x, "Transverse")
    ElseIf Direction = Axial Then
        Call Member.AddPtLoad(P, x, "Axial")
    End If
    
    'Flag the model as unsolved
    Solved = False
    
    'Exit the subroutine
    Exit Sub
            
End Sub

'Adds a concentrated moment to the member
Public Sub AddMemberMoment(MemberName As String, Moment As Double, x As Double)
      
    'Get the member from the `Members` dictionary
    Set Member = Members(MemberName)
      
    'Add the member moment
    Call Member.AddMoment(Moment, x)
    
    'Flag the model as unsolved
    Solved = False
    
End Sub

'Adds a linear distributed load to a member
Public Sub AddMemberDistLoad(MemberName As String, w1 As Double, w2 As Double, Optional x1 As Double = -1, Optional x2 As Double = -1, Optional Direction As LoadDir = Transverse, Optional LoadCase As Variant = "Case 1")
    
    'Find the member in the collection
    Dim Member As Member2D
    Set Member = Members(MemberName)
            
    'Default to the ends of the member if no values have been specified
    If x1 = -1 Then
        x1 = 0
    End If
    
    If x2 = -1 Then
        x2 = Member.Length
    End If
    
    'Add the member load
    If Direction = Transverse Then
        Call Member.AddLinLoad(w1, w2, x1, x2, "Transverse", LoadCase)
    ElseIf Direction = Axial Then
        Call Member.AddLinLoad(w1, w2, x1, x2, "Axial", LoadCase)
    End If
    
    'Flag the model as unsolved
    Solved = False
    
    'Exit the subroutine
    Exit Sub
    
End Sub

'Adds a nodal load to a node
Public Sub AddNodeLoad(NodeName As String, Load As Double, Direction As ForceType, Optional LoadCase = "Case 1")
    
    'Get the node to add the load to
    Dim Node As Node2D
    Set Node = Nodes(NodeName)
    
    'Add the load
    Dim i As Long
    i = Node.NodeLoads.NumRows + 1                 'Determine how long the array will be with 1 more item
    Call Node.NodeLoads.Resize(i, 3, True)         'Resize the array to contain 1 more item
    Call Node.NodeLoads.SetValue(i, 1, Direction)  'Add the load direction
    Call Node.NodeLoads.SetValue(i, 2, Load)       'Add the load magnitude
    Call Node.NodeLoads.SetValue(i, 3, LoadCase)   'Add the load case
    
    'Flag the model as unsolved
    Solved = False
    
    'Exit the subroutine
    Exit Sub
    
End Sub

'Adds a load combination to the model
Public Sub AddLoadCombo(ComboName As String, LoadCases As Variant, Factors As Variant)

    'Create a new load combination object
    Dim NewCombo As New LoadCombo
    
    'Add load factors to the load combination
    Dim i As Integer, LoadCase As Variant
    i = 0
    For Each LoadCase In LoadCases
        Call NewCombo.Factors.Add(LoadCases(i), Factors(i))
        i = i + 1
    Next LoadCase
    
    'Add the load combination to the dictionary of load combinations
    Call LoadCombos.Add(ComboName, NewCombo)
    'LoadCombos.Add Key:=ComboName, Item:=NewCombo
    
End Sub

'Analyzes the model
Public Sub Analyze()
    
    'Validate the model
    Call ValidateModel
    
    'Give each node in the structure a unique ID
    Dim Node As Node2D, NodeName As Variant, i As Long, NumSupports As Integer
    NumSupports = 0
    i = 1
    For Each NodeName In Nodes.keys
    
        Set Node = Nodes(NodeName)
        
        'Assign the node ID
        Node.ID = i
        i = i + 1
        
        'Store the node by its ID in the `NodesByID` dictionary
        Call NodesByID.Add(Node.ID, Node)
        
        'Count the number of supported DOF's
        If Node.SupportDX = True Then
            NumSupports = NumSupports + 1
        End If
        If Node.SupportDY = True Then
            NumSupports = NumSupports + 1
        End If
        If Node.SupportRZ = True Then
            NumSupports = NumSupports + 1
        End If
        
    Next NodeName
    
    'Ensure there is at least 1 load combination to solve if the user didn't define any
    Dim DefaultCombo As LoadCombo
    If LoadCombos.Count = 0 Then
        
        'Create and add a default load combination to the dictionary of load combinations
        Call LoadCombos.Add("Combo 1", Array("Case 1"), Array(1))
        
    End If
    
    'Matrix inversion is not possible if every DOF is supported. Check for this condition.
    Dim ComboName As Variant
    If NumSupports = Nodes.Count * 3 Then
            
        'Set each displacement at each node to zero
        For Each NodeName In Nodes.keys
        
            Set Node = Nodes(NodeName)
            
            For Each ComboName In LoadCombos.keys
                Node.DX(ComboName) = 0
                Node.DY(ComboName) = 0
                Node.RZ(ComboName) = 0
            Next ComboName
            
        Next NodeName
    
    Else
        
        'Step through each load combination
        Dim StructDisp As Matrix
        For Each ComboName In LoadCombos.keys
        
            'Calculate the structure's global displacement matrix for this load combination
            Set StructDisp = MMultiply(MInvert(StructStiff), MSubtract(StructNodalForces(ComboName), StructFER(ComboName)))
        
            'Place these displacements back into each node's displacement dictionary
            'The dictionary keys will be the load combination names
            Dim NodeID As Long
            i = 1
            For NodeID = 1 To NodesByID.Count
                
                Set Node = NodesByID(NodeID)
                
                If Node.SupportDX = False Then
                    Node.DX(ComboName) = StructDisp.GetValue(i, 1)
                    i = i + 1
                Else
                    Node.DX(ComboName) = 0
                End If
            
                If Node.SupportDY = False Then
                    Node.DY(ComboName) = StructDisp.GetValue(i, 1)
                    i = i + 1
                Else
                    Node.DY(ComboName) = 0
                End If
            
                If Node.SupportRZ = False Then
                    Node.RZ(ComboName) = StructDisp.GetValue(i, 1)
                    i = i + 1
                Else
                    Node.RZ(ComboName) = 0
                End If
           
            Next NodeID

        Next ComboName
    
    End If
    
    'Calculate the reactions
    Call CalcReactions
    
    'Flag the model as solved
    Solved = True
    
End Sub

'Validates the model
Private Sub ValidateModel()

    'Make sure there is at least 1 node defined
    If Nodes.Count = 0 Then
        MsgBox ("No nodes defined")
        Stop
    End If

    'Make sure there is at least 1 member defined
    If Members.Count = 0 Then
        MsgBox ("No members defined")
        Stop
    End If

End Sub

'Assembles and returns the structure's stiffness matrix
Private Function StructStiff() As Matrix
    
    'Determine the number of degrees of freedom (DOF's) in the model
    Dim NumDOF As Long
    NumDOF = CLng(Nodes.Count) * 3
    
    'Size the stiffness matrix to hold all the terms
    Set StructStiff = New Matrix
    Call StructStiff.Resize(NumDOF, NumDOF)
    
    'Place terms from each member into the global stiffness matrix
    Dim Member As Member2D, MemberName As Variant, i As Long, j As Long, MemStiff As Matrix
    Dim m As Long, n As Long
    For Each MemberName In Members.keys
        
        Set Member = Members(MemberName)
        
        'Identify the first DOF at the member's iNode and jNode
        i = Member.iNode.ID * 3 - 2
        j = Member.jNode.ID * 3 - 2
        
        'Get the member's local stiffness matrix
        Set MemStiff = Member.GlobalStiff
        
        'Copy terms from the member's stiffness matrix into the structure's stiffness matrix
        For m = 1 To 3
            For n = 1 To 3
                With StructStiff
                    Call .SetValue(i - 1 + m, i - 1 + n, .GetValue(i - 1 + m, i - 1 + n) + MemStiff.GetValue(m, n))
                    Call .SetValue(i - 1 + m, j - 1 + n, .GetValue(i - 1 + m, j - 1 + n) + MemStiff.GetValue(m, 3 + n))
                    Call .SetValue(j - 1 + m, i - 1 + n, .GetValue(j - 1 + m, i - 1 + n) + MemStiff.GetValue(3 + m, n))
                    Call .SetValue(j - 1 + m, j - 1 + n, .GetValue(j - 1 + m, j - 1 + n) + MemStiff.GetValue(3 + m, 3 + n))
                End With
            Next n
        Next m
        
    Next MemberName
    
    'Remove all the terms associated with the supports
    'Working backwards through the DOF's is easier (otherwise the indices would be changing)
    Dim NodeID As Long, Node As Node2D
    NodeID = CLng(NodesByID.Count)
    While NodeID > 0
        
        'The `-1` term is because dictionary keys are indexed starting at zero
        Set Node = NodesByID(NodeID)
        
        If Node.SupportRZ = True Then
            Call StructStiff.RemoveRow((NodeID - 1) * 3 + 3)
            Call StructStiff.RemoveCol((NodeID - 1) * 3 + 3)
        End If
        
        If Node.SupportDY = True Then
            Call StructStiff.RemoveRow((NodeID - 1) * 3 + 2)
            Call StructStiff.RemoveCol((NodeID - 1) * 3 + 2)
        End If
        
        If Node.SupportDX = True Then
            Call StructStiff.RemoveRow((NodeID - 1) * 3 + 1)
            Call StructStiff.RemoveCol((NodeID - 1) * 3 + 1)
        End If
        
        NodeID = NodeID - 1
        
    Wend
    
End Function

'Assembles and returns the structure's fixed end reaction vector
Private Function StructFER(Optional ComboName As Variant = "Combo 1") As Matrix

    'Determine the number of degrees of freedom (DOF's) in the model
    Dim NumDOF As Long
    NumDOF = CLng(Nodes.Count) * 3
    
    'Size the fixed end reaction vector to hold all the terms
    Set StructFER = New Matrix
    Call StructFER.Resize(NumDOF, 1)
    
    'Variable declarations for the next code block
    Dim Member As Member2D, MemberName As Variant
    Dim MemberFER As Matrix
    Dim i As Long, j As Long
    Dim m As Long
    
    'Place terms from each member into the global fixed end reaction vector
    For Each MemberName In Members.keys
        
        'Get the member from the dictionary
        Set Member = Members(MemberName)
        
        'Identify the first DOF at the member's iNode and jNode
        i = Member.iNode.ID * 3 - 2
        j = Member.jNode.ID * 3 - 2
        
        'Get the member's local fixed end reaction vector
        Set MemberFER = Member.GlobalFER(ComboName)
        
        'Copy terms from the member's fixed end reaction vector into the structure's vector
        For m = 1 To 3
            With StructFER
                Call .SetValue(i - 1 + m, 1, .GetValue(i - 1 + m, 1) + MemberFER.GetValue(m, 1))
                Call .SetValue(j - 1 + m, 1, .GetValue(j - 1 + m, 1) + MemberFER.GetValue(3 + m, 1))
            End With
        Next m
        
    Next MemberName
    
    'Remove all the terms associated with the supports
    'Working backwards through the DOF's is easier (otherwise the indices would be changing)
    Dim NodeID As Long, Node As Node2D
    NodeID = CLng(Nodes.Count)
    While NodeID > 0
        
        Set Node = NodesByID(NodeID)
        
        If Node.SupportRZ = True Then
            Call StructFER.RemoveRow((NodeID - 1) * 3 + 3)
        End If
        
        If Node.SupportDY = True Then
            Call StructFER.RemoveRow((NodeID - 1) * 3 + 2)
        End If
        
        If Node.SupportDX = True Then
            Call StructFER.RemoveRow((NodeID - 1) * 3 + 1)
        End If
        
        NodeID = NodeID - 1
        
    Wend
    
End Function

'Assembles and returns the structure's nodal force vector
Private Function StructNodalForces(Optional ComboName As Variant = "Combo 1") As Matrix
    
    'Determine the number of degrees of freedom (DOF's) in the model
    Dim NumDOF As Long
    NumDOF = CLng(Nodes.Count) * 3
    
    'Size the nodal force vector to hold all the terms
    Set StructNodalForces = New Matrix
    Call StructNodalForces.Resize(NumDOF, 1)
    
    'Get the requested load combination
    Dim Combo As LoadCombo
    Set Combo = LoadCombos(ComboName)
    
    'Place terms from each node into the nodal force vector
    Dim NodeName As Variant, Node As Node2D, LoadCase As Variant, Factor As Double, i As Long
    For Each NodeName In Nodes.keys
        
        'Get the node from the `Nodes` dictionary
        Set Node = Nodes(NodeName)
        
        'Loop through each load case in the load combination
        For Each LoadCase In LoadCombos(ComboName).Factors.keys
            
            'Get the load factor for this load case
            Factor = LoadCombos(ComboName).Factors(LoadCase)
            
            'Add the node's loads to the global node load vector
            For i = 1 To Node.NodeLoads.NumRows
                
                If Node.NodeLoads.GetValue(i, 1) = "FX" And LoadCase = Node.NodeLoads.GetValue(i, 3) Then
                    Call StructNodalForces.SetValue((Node.ID - 1) * 3 + 1, 1, Factor * Node.NodeLoads.GetValue(i, 2))
                ElseIf Node.NodeLoads.GetValue(i, 1) = "FY" And LoadCase = Node.NodeLoads.GetValue(i, 3) Then
                    Call StructNodalForces.SetValue((Node.ID - 1) * 3 + 2, 1, Factor * Node.NodeLoads.GetValue(i, 2))
                ElseIf Node.NodeLoads.GetValue(i, 1) = "MZ" And LoadCase = Node.NodeLoads.GetValue(i, 3) Then
                    Call StructNodalForces.SetValue((Node.ID - 1) * 3 + 3, 1, Factor * Node.NodeLoads.GetValue(i, 2))
                End If
                
            Next i
        
        Next LoadCase
        
    Next NodeName
    
    'Remove all the terms associated with the supports
    'Working backwards through the DOF's is easier (otherwise the indices would be changing)
    Dim NodeID As Long
    NodeID = CLng(Nodes.Count)
    While NodeID > 0
        
        Set Node = NodesByID(NodeID)
        
        If Node.SupportRZ = True Then
            Call StructNodalForces.RemoveRow((NodeID - 1) * 3 + 3)
        End If
        
        If Node.SupportDY = True Then
            Call StructNodalForces.RemoveRow((NodeID - 1) * 3 + 2)
        End If
        
        If Node.SupportDX = True Then
            Call StructNodalForces.RemoveRow((NodeID - 1) * 3 + 1)
        End If
        
        NodeID = NodeID - 1
        
    Wend
    
End Function

'Calculates the reactions once the model is solved
Private Function CalcReactions()
    
    'Variable declarations for the code below
    Dim NodeName As Variant, Node As Node2D
    Dim MemberName As Variant, Member As Member2D
    Dim ComboName As Variant, LoadCase As Variant, Factor As Double
    Dim MemberF As Matrix, NodeLoad As EZArray
    Dim i As Long
    
    'Calculate the reactions node by node
    For Each NodeName In Nodes.keys
    
        Set Node = Nodes(NodeName)
        
        'Step through each load combination
        For Each ComboName In LoadCombos.keys
            
            'Initialize reactions for this node and load combination
            Node.ReactionX(ComboName) = 0
            Node.ReactionY(ComboName) = 0
            Node.ReactionMZ(ComboName) = 0
            
            'Determine if the node has any supports
            If Node.SupportDX = True Or Node.SupportDY = True Or Node.SupportRZ = True Then
                
                'Sum the member end forces at the node
                For Each MemberName In Members.keys
                
                    'Get the member from the dictionary
                    Set Member = Members(MemberName)

                    If Member.iNode.Name = Node.Name Then
                    
                        'Get the member's global force matrix
                        'Storing it as a local variable eliminates the need to rebuild it every time a term is needed
                        Set MemberF = Member.GlobalForces(ComboName)
                        
                        Node.ReactionX(ComboName) = Node.ReactionX(ComboName) + MemberF.GetValue(1, 1)
                        Node.ReactionY(ComboName) = Node.ReactionY(ComboName) + MemberF.GetValue(2, 1)
                        Node.ReactionMZ(ComboName) = Node.ReactionY(ComboName) + MemberF.GetValue(3, 1)
                                                
                    ElseIf Member.jNode.Name = Node.Name Then
                        
                        'Get the member's global force matrix
                        'Storing it as a local variable eliminates the need to rebuild it every time a term is needed
                        Set MemberF = Member.GlobalForces(ComboName)
                        
                        Node.ReactionX(ComboName) = Node.ReactionX(ComboName) + MemberF.GetValue(4, 1)
                        Node.ReactionY(ComboName) = Node.ReactionY(ComboName) + MemberF.GetValue(5, 1)
                        Node.ReactionMZ(ComboName) = Node.ReactionY(ComboName) + MemberF.GetValue(6, 1)
                        
                    End If
                    
                Next MemberName
                
                'Sum the joint loads applied to the node
                For i = 1 To Node.NodeLoads.NumRows
                    
                    'Step through each load case in the current load combination
                    For Each LoadCase In LoadCombos(ComboName).Factors.keys
                        
                        'Check to see if load `i` in the array belongs to `LoadCase`
                        If LoadCase = Node.NodeLoads.GetValue(i, 3) Then
                            
                            Factor = LoadCombos(ComboName).Factors(LoadCase)
                            
                            If Node.NodeLoads.GetValue(i, 1) = "FX" Then
                                Node.ReactionX(ComboName) = Node.ReactionX(ComboName) - Factor * Node.NodeLoads.GetValue(i, 2)
                            ElseIf Node.NodeLoads.GetValue(i, 1) = "FY" Then
                                Node.ReactionY(ComboName) = Node.ReactionY(ComboName) - Factor * Node.NodeLoads.GetValue(i, 2)
                            ElseIf Node.NodeLoads.GetValue(i, 1) = "MZ" Then
                                Node.ReactionMZ(ComboName) = Node.ReactionMZ(ComboName) - Factor * Node.NodeLoads.GetValue(i, 2)
                            End If
                        
                        End If
                    
                    Next LoadCase
                    
                Next i
                
            End If
            
        Next ComboName
        
    Next NodeName

End Function

'Returns the displacement at a node
Public Function GetDisp(NodeName As String, Direction As DispType, Optional ComboName As Variant = "Combo 1") As Double
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Get the node from the `Nodes` dictionary
    Dim Node As Node2D
    Set Node = Nodes(NodeName)
    
    'Return the requested displacement
    If Direction = DX Then
        GetDisp = Node.DX(ComboName)
    ElseIf Direction = DY Then
        GetDisp = Node.DY(ComboName)
    ElseIf Direction = RZ Then
        GetDisp = Node.RZ(ComboName)
    End If
    
End Function

Public Function GetMemberDisp(MemberName As String, x As Double, Optional ComboName As Variant = "Combo 1") As Double
    
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Get the member from the `Members` dictionary
    Dim Member As Member2D
    Set Member = Members(MemberName)
            
    'Return the requested displacement
    GetMemberDisp = Member.Deflection(x, ComboName)
    
End Function

'Returns the local displacement diagram for a member
Public Function GetDispDiagram(MemberName As String, Optional ComboName As Variant = "Combo 1", Optional NumPoints As Integer = 20) As EZArray
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Start a new array to hold the results
    Set GetDispDiagram = New EZArray
    Call GetDispDiagram.Resize(CLng(NumPoints), 2)
    
    'Get the member from the `Members` dictionary
    Dim Member As Member2D
    Set Member = Members(MemberName)
            
    'Get the displacement at `NumPoints` points
    Dim i As Long, x As Double
    For i = 1 To NumPoints
        
        'Calculate the position of the x-coordinate
        x = (i - 1) * Member.Length / (NumPoints - 1)
        
        'Get the diagram coordinates
        Call GetDispDiagram.SetValue(i, 1, x)
        Call GetDispDiagram.SetValue(i, 2, Member.Deflection(x, ComboName))
        
    Next i
    
End Function

'Returns the local displacement diagram for a member
'This function ensures locations of load discontinuities are included in the displacement diagram.
'The number of points will vary depending on the number of discontinuities in the member.
Public Function GetDispDiagram2(MemberName As String, Optional ComboName As Variant = "Combo 1") As EZArray
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Start a new array to hold the results
    Set GetDispDiagram2 = New EZArray
    
    'Get the member from the `Members` dictionary
    Dim Member As Member2D
    Set Member = Members(MemberName)
    
    Const P As Long = 12 'qty of intermediate partitions per segment
    Dim r As Long 'EZArray index
    Dim s As Long 'Segment counter
    Dim SegmentItem As Variant 'Segment
    
    'Resize EZarray
    Call GetDispDiagram2.Resize(Member.NumSegs * (P + 1), 2)

    For Each SegmentItem In Member.Segments
    
        With SegmentItem
        
            Dim i As Long, x As Double
            s = s + 1
    
            'Get diagram coordinates at intermediate points
            For i = 1 To P + 1
                
                'Calculate the position of the x-coordinate
                x = .SegStart + (i - 1) * (.SegEnd - .SegStart) / P
                
                'Determine EZArray index value
                r = i + (s - 1) * (P + 1)
                
                'Get the diagram coordinates
                Call GetDispDiagram2.SetValue(r, 1, x)
                Call GetDispDiagram2.SetValue(r, 2, Member.Deflection(x, ComboName))
                
            Next i
            
        End With
        
    Next SegmentItem
    
End Function

'Returns the maximum displacement in a given member
Public Function GetMaxDisp(MemberName As String, Optional ComboName As Variant = "Combo 1") As Double
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    GetMaxDisp = Members(MemberName).MaxDisplacement(ComboName)
    
End Function

'Returns the minimum displacement in a given member
Public Function GetMinDisp(MemberName As String, Optional ComboName As Variant = "Combo 1") As Double
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    GetMinDisp = Members(MemberName).MinDisplacement(ComboName)

End Function

'Returns the shear in a member at a given location
Public Function GetShear(MemberName As String, x As Double, Optional ComboName As Variant = "Combo 1") As Double
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Get the shear at 'x'
    GetShear = Members(MemberName).Shear(x, ComboName)

End Function

'Returns the shear diagram for a member
Public Function GetShearDiagram(MemberName As String, Optional ComboName As Variant = "Combo 1", Optional NumPoints As Integer = 20) As EZArray
       
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Start a new array to hold the results
    Set GetShearDiagram = New EZArray
    Call GetShearDiagram.Resize(CLng(NumPoints), 2)
    
    'Get the member from the `Members` dictionary
    Dim Member As Member2D
    Set Member = Members(MemberName)
            
    'Get the shear force at `NumPoints` points
    Dim i As Long, x As Double
    For i = 1 To NumPoints
        
        'Calculate the position of the x-coordinate
        x = (i - 1) * Member.Length / (NumPoints - 1)
        
        'Get the diagram coordinates
        Call GetShearDiagram.SetValue(i, 1, x)
        Call GetShearDiagram.SetValue(i, 2, Member.Shear(x, ComboName))
        
    Next i
    
End Function

'Returns the shear diagram for a member
'This function ensures locations of load discontinuities are included in the shear diagram. The
'number of points will vary depending on the number of discontinuities in the member.
Public Function GetShearDiagram2(MemberName As String, Optional ComboName As Variant = "Combo 1") As EZArray

    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If

    'Start a new array to hold the results
    Set GetShearDiagram2 = New EZArray

    'Get the member from the `Members` dictionary
    Dim Member As Member2D
    Set Member = Members(MemberName)
    
    Const P As Long = 12 'qty of intermediate partitions per segment
    Dim r As Long 'EZArray index
    Dim s As Long 'Segment counter
    Dim SegmentItem As Variant 'Segment

    'Resize EZarray
    Call GetShearDiagram2.Resize(Member.NumSegs * (P + 1), 2)

    For Each SegmentItem In Member.Segments
    
        With SegmentItem
        
            Dim i As Long, x As Double
            s = s + 1
            
            'Determine EZArray index value
            r = 1 + (s - 1) * (P + 1)
            
            'Get diagram coordinates at start of segment
            Call GetShearDiagram2.SetValue(r, 1, .SegStart)
            Call GetShearDiagram2.SetValue(r, 2, .V1)
            
            'Get diagram coordinates at intermediate points
            For i = 2 To P + 1

                'Calculate the position of the x-coordinate
                x = .SegStart + (i - 1) * (.SegEnd - .SegStart) / P
                
                'Determine EZArray index value
                r = i + (s - 1) * (P + 1)

                'Get the rest of the diagram coordinates
                Call GetShearDiagram2.SetValue(r, 1, x)
                Call GetShearDiagram2.SetValue(r, 2, Member.Shear(x, ComboName))

            Next i
            
        End With
        
    Next SegmentItem

End Function

'Returns the maximum shear in a member
Public Function GetMaxShear(MemberName As String, Optional ComboName As Variant = "Combo 1") As Double
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Get the maximum shear
    GetMaxShear = Members(MemberName).Vmax(ComboName)
    
End Function

'Returns the minimum shear in a member
Public Function GetMinShear(MemberName As String, Optional ComboName As Variant = "Combo 1") As Double
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Get the maximum shear
    GetMinShear = Members(MemberName).Vmin(ComboName)
    
End Function

'Returns the moment in a member at a given location
Public Function GetMoment(MemberName As String, x As Double, Optional ComboName As Variant = "Combo 1") As Double
       
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Get the moment at 'x'
    GetMoment = Members(MemberName).Moment(x, ComboName)
            
End Function

'Returns the moment diagram for a member
Public Function GetMomentDiagram(MemberName As String, Optional ComboName As Variant = "Combo 1", Optional NumPoints As Integer = 20) As EZArray
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Start a new array to hold the results
    Set GetMomentDiagram = New EZArray
    Call GetMomentDiagram.Resize(CLng(NumPoints), 2)
    
    'Get the member from the `Members` dictionary
    Dim Member As Member2D
    Set Member = Members(MemberName)
    
    'Get the moment at `NumPoints` points
    Dim i As Integer, x As Double
    For i = 1 To NumPoints
        
        'Calculate the position of the x-coordinate
        x = (i - 1) * Member.Length / (NumPoints - 1)
        
        'Get the diagram coordinates
        Call GetMomentDiagram.SetValue(CLng(i), 1, x)
        Call GetMomentDiagram.SetValue(CLng(i), 2, Member.Moment(x, ComboName))
        
    Next i
    
End Function

'Returns the moment diagram for a member
'This function ensures locations of load discontinuities are included in the moment diagram. The
'number of points will vary depending on the number of discontinuities in the member.
Public Function GetMomentDiagram2(MemberName As String, Optional ComboName As Variant = "Combo 1") As EZArray
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Start a new array to hold the results
    Set GetMomentDiagram2 = New EZArray
    
    'Get the member from the `Members` dictionary
    Dim Member As Member2D
    Set Member = Members(MemberName)
    
    Const P As Long = 12 'qty of intermediate partitions per segment
    Dim r As Long 'EZArray index
    Dim s As Long 'Segment counter
    Dim SegmentItem As Variant 'Segment

    'Resize EZarray
    Call GetMomentDiagram2.Resize(Member.NumSegs * (P + 1), 2)

    For Each SegmentItem In Member.Segments
    
        With SegmentItem
        
            Dim i As Long, x As Double
            s = s + 1
            
            'Determine EZArray index value
            r = 1 + (s - 1) * (P + 1)
            
            'Get diagram coordinates at start of segment
            Call GetMomentDiagram2.SetValue(r, 1, .SegStart)
            Call GetMomentDiagram2.SetValue(r, 2, .M1)
    
            'Get diagram coordinates at intermediate points
            For i = 2 To P + 1
                
                'Calculate the position of the x-coordinate
                x = .SegStart + (i - 1) * (.SegEnd - .SegStart) / P
                
                'Determine EZArray index value
                r = i + (s - 1) * (P + 1)
                
                'Get the diagram coordinates
                Call GetMomentDiagram2.SetValue(r, 1, x)
                Call GetMomentDiagram2.SetValue(r, 2, Member.Moment(x, ComboName))
                
            Next i
            
        End With
        
    Next SegmentItem
    
End Function

'Returns the maximum moment in a member
Public Function GetMaxMoment(MemberName As String, Optional ComboName As Variant = "Combo 1") As Double
    
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If

    'Get the maximum moment
    GetMaxMoment = Members(MemberName).Mmax(ComboName)
    
End Function

'Returns the minimum moment in a member
Public Function GetMinMoment(MemberName As String, Optional ComboName As Variant = "Combo 1") As Double
    
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Get the minimum moment
    GetMinMoment = Members(MemberName).Mmin(ComboName)
    
End Function

'Returns the axial force in a member
Public Function GetAxial(MemberName As String, x As Double, Optional ComboName As Variant = "Combo 1") As Double
    
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If

    'Get the axial force at 'x'
    GetAxial = Members(MemberName).Axial(x, ComboName)
    
End Function

'Returns the axial force diagram for a member
Public Function GetAxialDiagram(MemberName As String, Optional ComboName As Variant = "Combo 1", Optional NumPoints As Integer = 20) As EZArray
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Start a new array to hold the results
    Set GetAxialDiagram = New EZArray
    Call GetAxialDiagram.Resize(CLng(NumPoints), 2)
    
    'Get the member from the `Members` dictionary
    Dim Member As Member2D
    Set Member = Members(MemberName)
            
    'Get the axial force at `NumPoints` points
    Dim i As Long, x As Double
    For i = 1 To NumPoints
        
        'Calculate the position of the x-coordinate
        x = (i - 1) * Member.Length / (NumPoints - 1)
        
        'Get the diagram coordinates
        Call GetAxialDiagram.SetValue(i, 1, x)
        Call GetAxialDiagram.SetValue(i, 2, Member.Axial(x, ComboName))
        
    Next i
    
End Function

'Returns the maximum axial force in a member
Public Function GetMaxAxial(MemberName As String, Optional ComboName As Variant = "Combo 1") As Double
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Get the maximum axial force
    GetMaxAxial = Members(MemberName).Pmax(ComboName)
    
End Function

'Returns the minimum axial force in a member
Public Function GetMinAxial(MemberName As String, Optional ComboName As Variant = "Combo 1") As Double
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Get the minimum axial force
    GetMinAxial = Members(MemberName).Pmin(ComboName)

End Function

'Constructor
Private Sub Class_Initialize()
        
    'Initialize the dictionaries
    Set Nodes = CreateObject("Scripting.Dictionary")
    Set NodesByID = CreateObject("Scripting.Dictionary")
    Set Members = CreateObject("Scripting.Dictionary")
    Set LoadCombos = CreateObject("Scripting.Dictionary")
        
    'Flag the model as not having been solved
    Solved = False
    
End Sub
