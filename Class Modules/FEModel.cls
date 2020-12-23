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

Public Nodes As New Collection      'A collection of all the nodes in the finite element model
Public Members As New Collection    'A collection of all the members in the finite element model
Private Solved As Boolean           'Flag indicating whether the model has been solved or not

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

'Directions for member loads
Public Enum LoadDirection
    Transverse
    Axial
End Enum

'Adds a node to the model
Public Sub AddNode(NodeName As String, XCoord As Double, YCoord As Double)

    'Create the new node
    Dim NewNode As New Node2D
    
    'Assign properties to the node
    NewNode.Name = NodeName
    NewNode.XCoord = XCoord
    NewNode.YCoord = YCoord
    
    'Add the node to the collection
    Nodes.Add NewNode
    
    'Flag the model as unsolved
    Solved = False
    
End Sub

'Removes a node from the model
Public Sub RemoveNode(NodeName As String)
    
    'Search through each item in the collection
    Dim NodeItem As Node2D, i As Integer
    i = 1
    For Each NodeItem In Nodes
        
        'Compare the item's ID to 'NodeID'
        If NodeItem.Name = NodeName Then
            
            'Remove the item at 'i'
            Nodes.Remove i
                
            'Flag the model as unsolved
            Solved = False
    
            'Exit the subroutine
            Exit Sub
            
        End If
        
        'Move to the next index
        i = i + 1
        
    Next NodeItem
    
End Sub

'Adds a support to the model
Public Sub EditSupport(NodeName As String, SupportDX As Boolean, SupportDY As Boolean, SupportRZ As Boolean)
    
    'Search through each item in the collection
    Dim NodeItem As Node2D
    For Each NodeItem In Nodes
        
        'Find the node
        If NodeItem.Name = NodeName Then
            
            'Add the support
            NodeItem.SupportDX = SupportDX
            NodeItem.SupportDY = SupportDY
            NodeItem.SupportRZ = SupportRZ
            
            'Flag the model as unsolved
            Solved = False
            
            'Exit the subroutine
            Exit Sub
            
        End If
        
    Next NodeItem
    
End Sub

'Adds a member to the model
Public Sub AddMember(MemberName As String, iNode As String, jNode As String, Elasticity As Double, Inertia As Double, Area As Double)
    
    'Create the new member
    Dim NewMember As New Member2D
    
    'Assign properties to the member
    NewMember.Name = MemberName
    NewMember.Elasticity = Elasticity
    NewMember.Inertia = Inertia
    NewMember.Area = Area
    
    'Assign the iNode and jNode to the member
    Dim NodeItem As Node2D
    For Each NodeItem In Nodes
        If NodeItem.Name = iNode Then
            Set NewMember.iNode = NodeItem
        ElseIf NodeItem.Name = jNode Then
            Set NewMember.jNode = NodeItem
        End If
    Next NodeItem
    
    'Add the member to the collection
    Members.Add NewMember
    
    'Flag the model as unsolved
    Solved = False
    
End Sub

'Adds a member end release
Public Sub EditEndReleases(MemberName As String, iRelease As Boolean, jRelease As Boolean)
    
    'Step through each member in the collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members
        
        'Find the member
        If MemberItem.Name = MemberName Then
            
            'Release the appropriate ends
            Call MemberItem.AddRelease(3, iRelease)
            Call MemberItem.AddRelease(6, jRelease)
            
            'Flag the model as unsolved
            Solved = False
            
            'Exit the subroutine
            Exit Sub
            
        End If
        
    Next MemberItem
    
End Sub

'Adds a point load to a member
Public Sub AddMemberPointLoad(MemberName As String, P As Double, x As Double, Direction As LoadDirection)
    
    'Step through each member
    Dim MemberItem As Member2D
    For Each MemberItem In Members
        
        'Find the member
        If MemberItem.Name = MemberName Then
            
            'Add the member load
            If Direction = Transverse Then
                Call MemberItem.AddPtLoad(P, x, "Transverse")
            ElseIf Direction = Axial Then
                Call MemberItem.AddPtLoad(P, x, "Axial")
            End If
            
            'Flag the model as unsolved
            Solved = False
            
            'Exit the subroutine
            Exit Sub
            
        End If
        
    Next MemberItem
    
End Sub

'Adds a concentrated moment to the member
Public Sub AddMemberMoment(MemberName As String, M As Double, x As Double)
    
    'Step through each member
    Dim MemberItem As Member2D
    For Each MemberItem In Members
        
        'Find the member
        If MemberItem.Name = MemberName Then
            
            'Add the member moment
            Call MemberItem.AddMoment(M, x)
            
            'Flag the model as unsolved
            Solved = False
            
            'Exit the subroutine
            Exit Sub
            
        End If
        
    Next MemberItem
    
End Sub

'Adds a linear distributed load to a member
Public Sub AddMemberDistLoad(MemberName As String, w1 As Double, w2 As Double, Optional x1 As Double = -1, Optional x2 As Double = -1, Optional Direction As LoadDirection = Transverse)
    
    'Find the member in the collection
    Dim MemberItem As Member2D
    
    'Step through each member in the 'Members' collection
    For Each MemberItem In Members
        
        'Find the member in the 'Members' collection
        If MemberItem.Name = MemberName Then
            
            'Default to the ends of the member if no values have been specified
            If x1 = -1 Then
                x1 = 0
            End If
            
            If x2 = -1 Then
                x2 = MemberItem.Length
            End If
            
            'Add the member load
            If Direction = Transverse Then
                Call MemberItem.AddLinLoad(w1, w2, x1, x2, "Transverse")
            ElseIf Direction = Axial Then
                Call MemberItem.AddLinLoad(w1, w2, x1, x2, "Axial")
            End If
            
            'Flag the model as unsolved
            Solved = False
            
            'Exit the subroutine
            Exit Sub
            
        End If
        
    Next MemberItem
    
End Sub

'Adds a nodal load to a node
Public Sub AddNodeLoad(NodeName As String, FX As Double, FY As Double, MZ As Double)
    
    'Step through each node in the collection
    Dim NodeItem As Node2D
    For Each NodeItem In Nodes
        
        'Find the node
        If NodeItem.Name = NodeName Then
            
            'Add the load
            Call NodeItem.AddForceX(FX)
            Call NodeItem.AddForceY(FY)
            Call NodeItem.AddMomentZ(MZ)
            
            'Flag the model as unsolved
            Solved = False
            
            'Exit the subroutine
            Exit Sub
            
        End If
        
    Next NodeItem
    
End Sub

'Analyzes the model
Private Sub Analyze()
    
    'Validate the model
    Call ValidateModel
    
    'Give each node in the structure a unique ID
    Dim NodeItem As Node2D, i As Long, NumSupports As Integer
    NumSupports = 0
    i = 1
    For Each NodeItem In Nodes
        
        'Assign the node ID
        NodeItem.ID = i
        i = i + 1
        
        'Count the number of supported DOF's
        If NodeItem.SupportDX = True Then
            NumSupports = NumSupports + 1
        End If
        If NodeItem.SupportDY = True Then
            NumSupports = NumSupports + 1
        End If
        If NodeItem.SupportRZ = True Then
            NumSupports = NumSupports + 1
        End If
        
    Next NodeItem
    
    'Matrix inversion is not possible if every DOF is supported - check for this condition
    If NumSupports = Nodes.Count * 3 Then
        
        'Set each displacement at each node to zero
        For Each NodeItem In Nodes
            NodeItem.DX = 0
            NodeItem.DY = 0
            NodeItem.RZ = 0
        Next NodeItem
        
    Else
            
        'Calculate the structure's global displacement matrix
        Dim StructDisp As Matrix
        Set StructDisp = MMultiply(MInvert(StructStiff), MSubtract(StructNodalForces, StructFER))
        
        'Place these displacements back into each node
        i = 1
        For Each NodeItem In Nodes
        
            If NodeItem.SupportDX = False Then
                NodeItem.DX = StructDisp.GetValue(i, 1)
                i = i + 1
            Else
                NodeItem.DX = 0
            End If
            
            If NodeItem.SupportDY = False Then
                NodeItem.DY = StructDisp.GetValue(i, 1)
                i = i + 1
            Else
                NodeItem.DY = 0
            End If
            
            If NodeItem.SupportRZ = False Then
                NodeItem.RZ = StructDisp.GetValue(i, 1)
                i = i + 1
            Else
                NodeItem.RZ = 0
            End If
           
        Next NodeItem
    
    End If
    
    'Calculate internal member forces
    Dim MemberItem As Member2D
    For Each MemberItem In Members
        MemberItem.SegmentMember
    Next MemberItem
    
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
    Dim MemberItem As Member2D, i As Long, j As Long, MemStiff As Matrix
    Dim M As Long, n As Long
    For Each MemberItem In Members
        
        'Identify the first DOF at the member's iNode and jNode
        i = MemberItem.iNode.ID * 3 - 2
        j = MemberItem.jNode.ID * 3 - 2
        
        'Get the member's local stiffness matrix
        Set MemStiff = MemberItem.GlobalStiff
        
        'Copy terms from the member's stiffness matrix into the structure's stiffness matrix
        For M = 1 To 3
            For n = 1 To 3
                With StructStiff
                    Call .SetValue(i - 1 + M, i - 1 + n, .GetValue(i - 1 + M, i - 1 + n) + MemStiff.GetValue(M, n))
                    Call .SetValue(i - 1 + M, j - 1 + n, .GetValue(i - 1 + M, j - 1 + n) + MemStiff.GetValue(M, 3 + n))
                    Call .SetValue(j - 1 + M, i - 1 + n, .GetValue(j - 1 + M, i - 1 + n) + MemStiff.GetValue(3 + M, n))
                    Call .SetValue(j - 1 + M, j - 1 + n, .GetValue(j - 1 + M, j - 1 + n) + MemStiff.GetValue(3 + M, 3 + n))
                End With
            Next n
        Next M
        
    Next MemberItem
    
    'Remove all the terms associated with the supports
    'Working backwards through the DOF's is easier (otherwise the indices would be changing)
    Dim NodeID As Long
    NodeID = Nodes.Count
    While NodeID > 0
    
        If Nodes(NodeID).SupportRZ = True Then
            Call StructStiff.RemoveRow((NodeID - 1) * 3 + 3)
            Call StructStiff.RemoveCol((NodeID - 1) * 3 + 3)
        End If
        
        If Nodes(NodeID).SupportDY = True Then
            Call StructStiff.RemoveRow((NodeID - 1) * 3 + 2)
            Call StructStiff.RemoveCol((NodeID - 1) * 3 + 2)
        End If
        
        If Nodes(NodeID).SupportDX = True Then
            Call StructStiff.RemoveRow((NodeID - 1) * 3 + 1)
            Call StructStiff.RemoveCol((NodeID - 1) * 3 + 1)
        End If
        
        NodeID = NodeID - 1
        
    Wend
    
End Function

'Assembles and returns the structure's fixed end reaction vector
Private Function StructFER() As Matrix

    'Determine the number of degrees of freedom (DOF's) in the model
    Dim NumDOF As Long
    NumDOF = CLng(Nodes.Count) * 3
    
    'Size the fixed end reaction vector to hold all the terms
    Set StructFER = New Matrix
    Call StructFER.Resize(NumDOF, 1)
    
    'Place terms from each member into the global fixed end reaction vector
    Dim MemberItem As Member2D, i As Long, j As Long, MemFER As Matrix
    Dim M As Long, n As Long
    For Each MemberItem In Members
        
        'Identify the first DOF at the member's iNode and jNode
        i = MemberItem.iNode.ID * 3 - 2
        j = MemberItem.jNode.ID * 3 - 2
        
        'Get the member's local fixed end reaction vector
        Set MemFER = MemberItem.GlobalFER
        
        'Copy terms from the member's fixed end reaction vector into the structure's vector
        For M = 1 To 3
            With StructFER
                Call .SetValue(i - 1 + M, 1, .GetValue(i - 1 + M, 1) + MemFER.GetValue(M, 1))
                Call .SetValue(j - 1 + M, 1, .GetValue(j - 1 + M, 1) + MemFER.GetValue(3 + M, 1))
            End With
        Next M
        
    Next MemberItem
    
    'Remove all the terms associated with the supports
    'Working backwards through the DOF's is easier (otherwise the indices would be changing)
    Dim NodeID As Long
    NodeID = Nodes.Count
    While NodeID > 0
    
        If Nodes(NodeID).SupportRZ = True Then
            Call StructFER.RemoveRow((NodeID - 1) * 3 + 3)
        End If
        
        If Nodes(NodeID).SupportDY = True Then
            Call StructFER.RemoveRow((NodeID - 1) * 3 + 2)
        End If
        
        If Nodes(NodeID).SupportDX = True Then
            Call StructFER.RemoveRow((NodeID - 1) * 3 + 1)
        End If
        
        NodeID = NodeID - 1
        
    Wend
    
End Function

'Assembles and returns the structure's nodal force vector
Private Function StructNodalForces() As Matrix
    
    'Determine the number of degrees of freedom (DOF's) in the model
    Dim NumDOF As Long
    NumDOF = CLng(Nodes.Count) * 3
    
    'Size the nodal force vector to hold all the terms
    Set StructNodalForces = New Matrix
    Call StructNodalForces.Resize(NumDOF, 1)
    
    'Place terms from each node into the nodal force vector
    Dim NodeItem As Node2D
    For Each NodeItem In Nodes
        
        Call StructNodalForces.SetValue((NodeItem.ID - 1) * 3 + 1, 1, NodeItem.GetForceX)
        Call StructNodalForces.SetValue((NodeItem.ID - 1) * 3 + 2, 1, NodeItem.GetForceY)
        Call StructNodalForces.SetValue((NodeItem.ID - 1) * 3 + 3, 1, NodeItem.GetMomentZ)
        
    Next NodeItem
    
    'Remove all the terms associated with the supports
    'Working backwards through the DOF's is easier (otherwise the indices would be changing)
    Dim NodeID As Long
    NodeID = Nodes.Count
    While NodeID > 0
    
        If Nodes(NodeID).SupportRZ = True Then
            Call StructNodalForces.RemoveRow((NodeID - 1) * 3 + 3)
        End If
        
        If Nodes(NodeID).SupportDY = True Then
            Call StructNodalForces.RemoveRow((NodeID - 1) * 3 + 2)
        End If
        
        If Nodes(NodeID).SupportDX = True Then
            Call StructNodalForces.RemoveRow((NodeID - 1) * 3 + 1)
        End If
        
        NodeID = NodeID - 1
        
    Wend
    
End Function

'Returns the requested reaction
Public Function GetReaction(NodeName As String, RxnType As ForceType)
    
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    Dim NodeItem As Node2D, MemberItem As Member2D, Sum As Double
    
    'Find the node in the 'Nodes' collection
    For Each NodeItem In Nodes
    
        If NodeItem.Name = NodeName Then
            
            'Determine which direction is of interest
            If RxnType = FX Then
                
                'Sum the nodal force at the node
                Sum = NodeItem.GetForceX
                
                'Sum the forces from any members framing into the node
                For Each MemberItem In Members
                
                    If MemberItem.iNode.Name = NodeName Then
                        Sum = Sum - MemberItem.GlobalForces.GetValue(1, 1)
                    ElseIf MemberItem.jNode.Name = NodeName Then
                        Sum = Sum - MemberItem.GlobalForces.GetValue(4, 1)
                    End If
                    
                Next MemberItem
                
            ElseIf RxnType = FY Then
            
                'Sum the nodal force at the node
                Sum = NodeItem.GetForceY
                
                'Sum the forces from any members framing into the node
                For Each MemberItem In Members
                
                    If MemberItem.iNode.Name = NodeName Then
                        Sum = Sum - MemberItem.GlobalForces.GetValue(2, 1)
                    ElseIf MemberItem.jNode.Name = NodeName Then
                        Sum = Sum - MemberItem.GlobalForces.GetValue(5, 1)
                    End If
                    
                Next MemberItem
                
            ElseIf RxnType = MZ Then
            
                'Sum the nodal force at the node
                Sum = NodeItem.GetMomentZ
                
                'Sum the forces from any members framing into the node
                For Each MemberItem In Members
                
                    If MemberItem.iNode.Name = NodeName Then
                        Sum = Sum - MemberItem.GlobalForces.GetValue(3, 1)
                    ElseIf MemberItem.jNode.Name = NodeName Then
                        Sum = Sum - MemberItem.GlobalForces.GetValue(6, 1)
                    End If
                    
                Next MemberItem
                
            End If
            
        End If
        
    Next NodeItem
    
    'Return the result
    GetReaction = -Sum

End Function

'Returns the displacement at a node
Public Function GetDisp(NodeName As String, Direction As DispType) As Double
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Step through each node in the collection
    Dim NodeItem As Node2D
    For Each NodeItem In Nodes
        
        'Find the node
        If NodeItem.Name = NodeName Then
            
            'Return the requested displacement
            If Direction = DX Then
                GetDisp = NodeItem.DX
            ElseIf Direction = DY Then
                GetDisp = NodeItem.DY
            ElseIf Direction = RZ Then
                GetDisp = NodeItem.RZ
            End If
            
            'Exit the function
            Exit Function
            
        End If
        
    Next NodeItem
    
End Function

Public Function GetMemberDisp(MemberName As String, x As Double) As Double
    
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Step through each member in the collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members
        
        'Find the node
        If MemberItem.Name = MemberName Then
            
            'Return the requested displacement
            GetMemberDisp = MemberItem.Deflection(x)
            
            'Exit the function
            Exit Function
            
        End If
        
    Next MemberItem
    
End Function

'Returns the local displacement diagram for a member
Public Function GetDispDiagram(MemberName As String) As EZArray
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Start a new array to hold the results
    Set GetDispDiagram = New EZArray
    Call GetDispDiagram.Resize(20, 2)
    
    'Step through each member in the 'Members' collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members
    
        'Find the member in the 'Members' collection
        If MemberItem.Name = MemberName Then
            
            'Get the displacement at 20 points
            Dim i As Long, x As Double
            For i = 1 To 20
                
                'Calculate the position of the x-coordinate
                x = (i - 1) * MemberItem.Length / 19
                
                'Get the diagram coordinates
                Call GetDispDiagram.SetValue(i, 1, x)
                Call GetDispDiagram.SetValue(i, 2, MemberItem.Deflection(x))
                
            Next i
            
            Exit Function
            
        End If
        
    Next MemberItem
    
End Function

'Returns the local displacement diagram for a member, segment-by-segment
Public Function GetDispDiagram2(MemberName As String) As EZArray
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Start a new array to hold the results
    Set GetDispDiagram2 = New EZArray
    
    'Step through each member in the 'Members' collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members
    
        'Find the member in the 'Members' collection
        If MemberItem.Name = MemberName Then
            Const p As Long = 12 'qty of intermediate partitions per segment
            Dim r As Long 'EZArray index
            Dim s As Long 'Segment counter
            Dim SegmentItem As Variant 'Segment

            'Resize EZarray
            Call GetDispDiagram2.Resize(MemberItem.NumSegs * (p + 1), 2)

            For Each SegmentItem In MemberItem.Segments
                With SegmentItem
                
                    Dim i As Long, x As Double
                    s = s + 1
            
                    'Get diagram coordinates at intermediate points
                    For i = 1 To p + 1
                        
                        'Calculate the position of the x-coordinate
                        x = .SegStart + (i - 1) * (.SegEnd - .SegStart) / p
                        
                        'Determine EZArray index value
                        r = i + (s - 1) * (p + 1)
                        
                        'Get the diagram coordinates
                        Call GetDispDiagram2.SetValue(r, 1, x)
                        Call GetDispDiagram2.SetValue(r, 2, MemberItem.Deflection(x))
                        
                    Next i
                    
                End With 'SegmentItem
            Next SegmentItem
            
            Exit Function
            
        End If
        
    Next MemberItem
    
End Function

'Returns the maximum displacement in a given member
Public Function GetMaxDisp(MemberName As String) As Double
    
    'Declare local variables
    Dim MemberItem As Member2D
    
    'Step through each member in the collection
    For Each MemberItem In Members
    
        'Get the desired member's maximum displacement
        If MemberItem.Name = MemberName Then
            GetMaxDisp = MemberItem.MaxDisplacement
        End If
        
    Next MemberItem
    
End Function

'Returns the minimum displacement in a given member
Public Function GetMinDisp(MemberName As String) As Double
    
    'Declare local variables
    Dim MemberItem As Member2D
    
    'Step through each member in the collection
    For Each MemberItem In Members
    
        'Get the desired member's minimum displacement
        If MemberItem.Name = MemberName Then
            GetMinDisp = MemberItem.MinDisplacement
        End If
        
    Next MemberItem
    
End Function

'Returns the shear in a member at a given location
Public Function GetShear(MemberName As String, x As Double) As Double
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Step through each member in the collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members
        
        'Find the member
        If MemberItem.Name = MemberName Then
            
            'Get the shear at 'x'
            GetShear = MemberItem.Shear(x)
            Exit Function
            
        End If
        
    Next MemberItem

End Function

'Returns the shear diagram for a member
Public Function GetShearDiagram(MemberName As String) As EZArray
       
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Start a new array to hold the results
    Set GetShearDiagram = New EZArray
    Call GetShearDiagram.Resize(20, 2)
    
    'Step through each member in the 'Members' collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members
    
        'Find the member in the 'Members' collection
        If MemberItem.Name = MemberName Then
            
            'Get the shear force at 20 points
            Dim i As Long, x As Double
            For i = 1 To 20
                
                'Calculate the position of the x-coordinate
                x = (i - 1) * MemberItem.Length / 19
                
                'Get the diagram coordinates
                Call GetShearDiagram.SetValue(i, 1, x)
                Call GetShearDiagram.SetValue(i, 2, MemberItem.Shear(x))
                
            Next i
            
            Exit Function
            
        End If
        
    Next MemberItem
    
End Function

'Returns the shear diagram for a member, segment-by-segment
Public Function GetShearDiagram2(MemberName As String) As EZArray

    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If

    'Start a new array to hold the results
    Set GetShearDiagram2 = New EZArray

    'Step through each member in the 'Members' collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members

        'Find the member in the 'Members' collection
        If MemberItem.Name = MemberName Then
            Const p As Long = 12 'qty of intermediate partitions per segment
            Dim r As Long 'EZArray index
            Dim s As Long 'Segment counter
            Dim SegmentItem As Variant 'Segment

            'Resize EZarray
            Call GetShearDiagram2.Resize(MemberItem.NumSegs * (p + 1), 2)

            For Each SegmentItem In MemberItem.Segments
                With SegmentItem
                
                    Dim i As Long, x As Double
                    s = s + 1
                    
                    'Determine EZArray index value
                    r = 1 + (s - 1) * (p + 1)
                    
                    'Get diagram coordinates at start of segment
                    Call GetShearDiagram2.SetValue(r, 1, .SegStart)
                    Call GetShearDiagram2.SetValue(r, 2, .V1)
                    
                    'Get diagram coordinates at intermediate points
                    For i = 2 To p + 1
        
                        'Calculate the position of the x-coordinate
                        x = .SegStart + (i - 1) * (.SegEnd - .SegStart) / p
                        
                        'Determine EZArray index value
                        r = i + (s - 1) * (p + 1)
        
                        'Get the rest of the diagram coordinates
                        Call GetShearDiagram2.SetValue(r, 1, x)
                        Call GetShearDiagram2.SetValue(r, 2, MemberItem.Shear(x))
        
                    Next i
                    
                End With 'SegmentItem
            Next SegmentItem
            
            Exit Function

        End If

    Next MemberItem

End Function

'Returns the maximum shear in a member
Public Function GetMaxShear(MemberName As String) As Double
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Step through each member in the collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members
        
        'Find the member
        If MemberItem.Name = MemberName Then
            
            'Get the maximum shear
            GetMaxShear = MemberItem.Vmax
            Exit Function
            
        End If
        
    Next MemberItem
    
End Function

'Returns the minimum shear in a member
Public Function GetMinShear(MemberName As String) As Double
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Step through each member in the collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members
        
        'Find the member
        If MemberItem.Name = MemberName Then
            
            'Get the maximum shear
            GetMinShear = MemberItem.Vmin
            Exit Function
            
        End If
        
    Next MemberItem
    
End Function

'Returns the moment in a member at a given location
Public Function GetMoment(MemberName As String, x As Double) As Double
       
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Step through each member in the collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members
        
        'Find the member
        If MemberItem.Name = MemberName Then
            
            'Get the moment at 'x'
            GetMoment = MemberItem.Moment(x)
            Exit Function
            
        End If
        
    Next MemberItem
    
End Function

'Returns the moment diagram for a member
Public Function GetMomentDiagram(MemberName As String) As EZArray
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Start a new array to hold the results
    Set GetMomentDiagram = New EZArray
    Call GetMomentDiagram.Resize(20, 2)
    
    'Step through each member in the 'Members' collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members
    
        'Find the member in the 'Members' collection
        If MemberItem.Name = MemberName Then
            
            'Get the moment at 20 points
            Dim i As Long, x As Double
            For i = 1 To 20
                
                'Calculate the position of the x-coordinate
                x = (i - 1) * MemberItem.Length / 19
                
                'Get the diagram coordinates
                Call GetMomentDiagram.SetValue(i, 1, x)
                Call GetMomentDiagram.SetValue(i, 2, MemberItem.Moment(x))
                
            Next i
            
            Exit Function
            
        End If
        
    Next MemberItem
    
End Function

'Returns the moment diagram for a member, segment-by-segment
Public Function GetMomentDiagram2(MemberName As String) As EZArray
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Start a new array to hold the results
    Set GetMomentDiagram2 = New EZArray
    
    'Step through each member in the 'Members' collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members
    
        'Find the member in the 'Members' collection
        If MemberItem.Name = MemberName Then
            Const p As Long = 12 'qty of intermediate partitions per segment
            Dim r As Long 'EZArray index
            Dim s As Long 'Segment counter
            Dim SegmentItem As Variant 'Segment

            'Resize EZarray
            Call GetMomentDiagram2.Resize(MemberItem.NumSegs * (p + 1), 2)

            For Each SegmentItem In MemberItem.Segments
                With SegmentItem
                
                    Dim i As Long, x As Double
                    s = s + 1
                    
                    'Determine EZArray index value
                    r = 1 + (s - 1) * (p + 1)
                    
                    'Get diagram coordinates at start of segment
                    Call GetMomentDiagram2.SetValue(r, 1, .SegStart)
                    Call GetMomentDiagram2.SetValue(r, 2, .M1)
            
                    'Get diagram coordinates at intermediate points
                    For i = 2 To p + 1
                        
                        'Calculate the position of the x-coordinate
                        x = .SegStart + (i - 1) * (.SegEnd - .SegStart) / p
                        
                        'Determine EZArray index value
                        r = i + (s - 1) * (p + 1)
                        
                        'Get the diagram coordinates
                        Call GetMomentDiagram2.SetValue(r, 1, x)
                        Call GetMomentDiagram2.SetValue(r, 2, MemberItem.Moment(x))
                        
                    Next i
                    
                End With 'SegmentItem
            Next SegmentItem
            
            Exit Function
            
        End If
        
    Next MemberItem
    
End Function

'Returns the maximum moment in a member
Public Function GetMaxMoment(MemberName As String) As Double
    
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Step through each member in the collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members
        
        'Find the member
        If MemberItem.Name = MemberName Then
            
            'Get the maximum moment
            GetMaxMoment = MemberItem.Mmax
            Exit Function
            
        End If
        
    Next MemberItem
    
End Function

'Returns the minimum moment in a member
Public Function GetMinMoment(MemberName As String) As Double
    
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Step through each member in the collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members
        
        'Find the member
        If MemberItem.Name = MemberName Then
            
            'Get the minimum moment
            GetMinMoment = MemberItem.Mmin
            Exit Function
            
        End If
        
    Next MemberItem
    
End Function

'Returns the axial force in a member
Public Function GetAxial(MemberName As String, x As Double) As Double
    
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Step through each member in the collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members
        
        'Find the member
        If MemberItem.Name = MemberName Then
            
            'Get the axial force at 'x'
            GetAxial = MemberItem.Axial(x)
            Exit Function
            
        End If
        
    Next MemberItem
    
End Function

'Returns the axial force diagram for a member
Public Function GetAxialDiagram(MemberName As String) As EZArray
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Start a new array to hold the results
    Set GetAxialDiagram = New EZArray
    Call GetAxialDiagram.Resize(20, 2)
    
    'Step through each member in the 'Members' collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members
    
        'Find the member in the 'Members' collection
        If MemberItem.Name = MemberName Then
            
            'Get the axial force at 20 points
            Dim i As Long, x As Double
            For i = 1 To 20
                
                'Calculate the position of the x-coordinate
                x = (i - 1) * MemberItem.Length / 19
                
                'Get the diagram coordinates
                Call GetAxialDiagram.SetValue(i, 1, x)
                Call GetAxialDiagram.SetValue(i, 2, MemberItem.Axial(x))
                
            Next i
            
            Exit Function
            
        End If
        
    Next MemberItem
    
End Function

'Returns the maximum axial force in a member
Public Function GetMaxAxial(MemberName As String) As Double
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Step through each member in the collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members
        
        'Find the member
        If MemberItem.Name = MemberName Then
            
            'Get the maximum axial force
            GetMaxAxial = MemberItem.Pmax
            Exit Function
            
        End If
        
    Next MemberItem
    
End Function

'Returns the minimum axial force in a member
Public Function GetMinAxial(MemberName As String) As Double
        
    'Solve the model if necessary
    If Solved = False Then
        Call Analyze
    End If
    
    'Step through each member in the collection
    Dim MemberItem As Member2D
    For Each MemberItem In Members
        
        'Find the member
        If MemberItem.Name = MemberName Then
            
            'Get the minimum axial force
            GetMinAxial = MemberItem.Pmin
            Exit Function
            
        End If
        
    Next MemberItem
    
End Function

'Constructor
Private Sub Class_Initialize()
    
    'Flag the model as not having been solved
    Solved = False
    
End Sub
