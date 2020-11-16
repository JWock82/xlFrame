VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Member2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Represents a 2D frame element

'Dependencies:
'Segment.cls
'EZArray.cls
'Matrix.cls
'FERFunctions.bas

'Force explicit variable declaration in this class module
Option Explicit

'Class variables
'Member/node ID's
Private m_Name As String                    'Member name
Private m_MemberID As Integer               'Member ID that can be assigned in code-behind to help track the member
Private m_iNode As Node2D                    'i-node
Private m_jNode As Node2D                    'j-node

'Member properties
Private m_E As Double                       'Modulus of elasticity
Private m_I As Double                       'Moment of inertia
Private m_A As Double                       'Cross sectional area

'End releases
Private m_Releases(1 To 6) As Boolean       'Array representing member end releases
Private m_Condense As Boolean               'Flag indicating whether static condensation is required

'Loads
Private m_Loads As New EZArray              'Array representing member loads
Private m_FER As New Matrix                 'Member local fixed end force vector

'Internal forces
Private m_Segments() As New Segment         'Array representing mathematically continuous segments of the member
Private m_NumSegs As Integer                'Number of mathematically continuous segments in the member

'Class properties
Property Get NumSegs() As Integer
    NumSegs = m_NumSegs
End Property

Property Get Segments() As Segment()
    Segments = m_Segments
End Property
  
Property Let Name(Value As String)
    m_Name = Value
End Property

Property Get Name() As String
    Name = m_Name
End Property

Property Let MemberID(Value As Integer)
    m_MemberID = Value
End Property

Property Get MemberID() As Integer
    MemberID = m_MemberID
End Property

Property Set iNode(Value As Node2D)
    Set m_iNode = Value
End Property

Property Get iNode() As Node2D
    Set iNode = m_iNode
End Property

Property Set jNode(Value As Node2D)
    Set m_jNode = Value
End Property

Property Get jNode() As Node2D
    Set jNode = m_jNode
End Property

Property Let Elasticity(Value As Double)
    m_E = Value
End Property

Property Get Elasticity() As Double
    Elasticity = m_E
End Property

Property Let Inertia(Value As Double)
    m_I = Value
End Property

Property Get Inertia() As Double
    Inertia = m_I
End Property

Property Let Area(Value As Double)
    m_A = Value
End Property

Property Get Area() As Double
    Area = m_A
End Property

Property Get Length() As Double
    Length = Sqr((m_jNode.XCoord - m_iNode.XCoord) ^ 2 + (m_jNode.YCoord - m_iNode.YCoord) ^ 2)
End Property

'Initializes the class prior to use
Private Sub Class_Initialize()
    
    'Initialize the FER vector and global displacement vector
    Call ClearLoads
    
    'Initialize all the end releases to "False"
    Dim i As Integer
    For i = 1 To 6
        m_Releases(i) = False
    Next i
    
    'Initialize the static condensation flag
    m_Condense = False
    
End Sub

'Deletes all loads from the element
Public Sub ClearLoads()
    
    'Declare local variables
    Dim i As Integer
    
    'Initialize the segment count to zero
    m_NumSegs = 0
    
    'Set all terms in the FER vector to zero
    Call m_FER.Resize(6, 1, False)
    
    'Delete the old array of loads
    Set m_Loads = Nothing
    
    'Create a new array of loads
    Set m_Loads = New EZArray
    
End Sub

'Changes a member end release at the specified degree of freedom
Public Sub AddRelease(DOF As Integer, Released As Boolean)

    'Add the release
    m_Releases(DOF) = Released
    
    'Flag the member if it needs static condensation
    If Released = True Then
        m_Condense = True
    End If

End Sub

'Returns whether the specified degree of freedom is released or not
Public Function IsReleased(DOF As Integer) As Boolean
    
    'Validate the DOF
    If DOF >= 1 And DOF <= 6 Then
        'Return the requested release
        IsReleased = m_Releases(DOF)
    Else
        MsgBox ("Error: Unable to return end release. Invalid DOF.")
        Stop
    End If
    
End Function

'Returns the transformation matrix
Public Function TransMatrix() As Matrix

    'Find the direction cosines (member orientation)
    Dim DirCos As Double, DirSin As Double
    DirCos = (m_jNode.XCoord - m_iNode.XCoord) / Length
    DirSin = (m_jNode.YCoord - m_iNode.YCoord) / Length
    
    'Initialize the transformation matrix
    Set TransMatrix = New Matrix
    Call TransMatrix.Resize(6, 6, False)
    
    'Add nonzero terms to the matrix
    Call TransMatrix.SetValue(1, 1, DirCos)
    Call TransMatrix.SetValue(1, 2, DirSin)
    Call TransMatrix.SetValue(2, 1, -DirSin)
    Call TransMatrix.SetValue(2, 2, DirCos)
    Call TransMatrix.SetValue(3, 3, 1)
    Call TransMatrix.SetValue(4, 4, DirCos)
    Call TransMatrix.SetValue(4, 5, DirSin)
    Call TransMatrix.SetValue(5, 4, -DirSin)
    Call TransMatrix.SetValue(5, 5, DirCos)
    Call TransMatrix.SetValue(6, 6, 1)
    
End Function

'Returns the local stiffness matrix.
'If static condensation is required the expanded form of the local stiffness matrix is returned
Public Function LocalStiff(Optional Condensed As Boolean = True) As Matrix
    
    'Create the local stiffness matrix
    Set LocalStiff = New Matrix
    Call LocalStiff.Resize(6, 6)
    
    'Add nonzero terms to the matrix
    Dim Coeff As Double
    Coeff = Elasticity * Inertia / Length ^ 3
    Call LocalStiff.SetValue(1, 1, Coeff * Area * Length ^ 2 / Inertia)
    Call LocalStiff.SetValue(1, 4, -Coeff * Area * Length ^ 2 / Inertia)
    Call LocalStiff.SetValue(2, 2, Coeff * 12)
    Call LocalStiff.SetValue(2, 3, Coeff * 6 * Length)
    Call LocalStiff.SetValue(2, 5, -Coeff * 12)
    Call LocalStiff.SetValue(2, 6, Coeff * 6 * Length)
    Call LocalStiff.SetValue(3, 3, Coeff * 4 * Length ^ 2)
    Call LocalStiff.SetValue(3, 5, -Coeff * 6 * Length)
    Call LocalStiff.SetValue(3, 6, Coeff * 2 * Length ^ 2)
    Call LocalStiff.SetValue(4, 4, Coeff * Area * Length ^ 2 / Inertia)
    Call LocalStiff.SetValue(5, 5, Coeff * 12)
    Call LocalStiff.SetValue(5, 6, -Coeff * 6 * Length)
    Call LocalStiff.SetValue(6, 6, Coeff * 4 * Length ^ 2)
    
    'Apply symmetry to the matrix to get the remaining terms
    Dim i As Long, j As Long
    For i = 1 To 6
        For j = i To 6
            Call LocalStiff.SetValue(j, i, LocalStiff.GetValue(i, j))
        Next j
    Next i
    
    'Determine whether the matrix should be condensed
    If Condensed = True Then
        'Condense the matrix
        Set LocalStiff = k_Condense(LocalStiff, m_Releases)
    End If
    
End Function

'Returns the expanded global stiffness matrix of the member
Public Function GlobalStiff() As Variant
    
    'Declare local variables
    Dim M1 As Matrix, M2 As Matrix, M3 As Matrix
    
    'Transpose the transformation matrix
    Set M1 = MTranspose(TransMatrix)
    
    'Get the expanded local stiffness matrix (condensed if it applies)
    Set M2 = LocalStiff(True)
    
    'Multiply the local stiffness matrix by the transposed transformation matrix
    Set M3 = MMultiply(M1, M2)
    
    'Multiply the transformation matrix by the result from the previous operation
    Set GlobalStiff = MMultiply(M3, TransMatrix)
    
End Function

'Returns the local fixed end reaction vector
Public Function LocalFER(Optional Condensed As Boolean = True) As Matrix
    
    'Determine if the local fixed end reaction vector requires static condensation
    If Condensed = True Then
        
        'Condense the local fixed end reaction vector
        Set LocalFER = CondenseFER(m_FER, LocalStiff(False), m_Releases)
        
    Else
    
        'Get the local fixed end reaction vector
        Set LocalFER = m_FER
        
    End If
    
End Function

'Returns the global fixed end reaction vector
Public Function GlobalFER() As Matrix
    
    'Transpose the transformation matrix
    Dim M1 As Matrix
    Set M1 = MTranspose(TransMatrix)
    
    'Get the local fixed end reaction vector
    Dim M2 As Matrix
    Set M2 = LocalFER(True)
    
    'Multiply the local fixed end reaction vector by the transposed transformation matrix
    Set GlobalFER = MMultiply(M1, M2)
    
End Function

'Returns the local displacement vector
Public Function LocalDisp() As Matrix
    
    'Calculate the local displacement vector
    Set LocalDisp = MMultiply(TransMatrix, GlobalDisp)
    
End Function

'Returns the global displacement vector
'The nodal displacements must have been solved for prior to calling this function
Public Function GlobalDisp() As Matrix
    
    'Generate the global displacement vector
    Set GlobalDisp = New Matrix
    Call GlobalDisp.Resize(6, 1, False)
    Call GlobalDisp.SetValue(1, 1, iNode.DX)
    Call GlobalDisp.SetValue(2, 1, iNode.DY)
    Call GlobalDisp.SetValue(3, 1, iNode.RZ)
    Call GlobalDisp.SetValue(4, 1, jNode.DX)
    Call GlobalDisp.SetValue(5, 1, jNode.DY)
    Call GlobalDisp.SetValue(6, 1, jNode.RZ)
    
End Function

'Returns the local member end force vector
Public Function LocalForces() As Matrix
    
    'Declare local variables
    Dim M1 As Matrix, M2 As Matrix, M3 As Matrix, M4 As Matrix
    
    'Get the local stiffness matrix
    Set M1 = LocalStiff(True)
    
    'Get the local displacement vector
    Set M2 = LocalDisp
    
    'Get the local FER vector
    Set M3 = LocalFER(True)
    
    'Multiply the local displacement vector by the local stiffness matrix
    Set M4 = MMultiply(M1, M2)
    
    'Add the result of the previous operation to the local FER vector
    Set LocalForces = MAdd(M4, M3)
        
End Function

'Returns the global member end force vector
Public Function GlobalForces() As Matrix
    
    'Declare local variables
    Dim M1 As Matrix, M2 As Matrix
    
    'Transpose the transformation matrix
    Set M1 = MTranspose(TransMatrix)
    
    'Get the local member end force vector
    Set M2 = LocalForces
    
    'Multiply the local member end force vector by the transposed transformation matrix
    Set GlobalForces = MMultiply(M1, M2)
        
End Function

'Adds the fixed end reactions for a point load to the FER vector and places the load in the load array
Public Sub AddPtLoad(P As Double, x As Double, Direction As String)
    
    'Resize the "m_Loads" array
    Call m_Loads.Resize(m_Loads.NumRows + 1, 7, True)
    
    'Add the load to the "m_Loads" array
    Call m_Loads.SetValue(m_Loads.NumRows, 1, 0)
    Call m_Loads.SetValue(m_Loads.NumRows, 2, 0)
    Call m_Loads.SetValue(m_Loads.NumRows, 3, P)
    Call m_Loads.SetValue(m_Loads.NumRows, 4, 0)
    Call m_Loads.SetValue(m_Loads.NumRows, 5, x)
    Call m_Loads.SetValue(m_Loads.NumRows, 6, 0)
    Call m_Loads.SetValue(m_Loads.NumRows, 7, Direction)
    
    'Add the fixed end reactions to the FER vector
    If Direction = "Transverse" Then
        Set m_FER = MAdd(m_FER, FER_PtLoad(P, x, Length))
    ElseIf Direction = "Axial" Then
        Set m_FER = MAdd(m_FER, FER_AxialPtLoad(P, x, Length))
    End If
    
End Sub

'Adds the fixed end reactions for a distributed load to the FER vector and places the load in the load vector
Public Sub AddLinLoad(w1 As Double, w2 As Double, x1 As Double, x2 As Double, Direction As String)
    
    'Resize the "m_Loads" array
    Call m_Loads.Resize(m_Loads.NumRows + 1, 7)
    
    'Add the load to the "m_Loads" array
    Call m_Loads.SetValue(m_Loads.NumRows, 1, w1)
    Call m_Loads.SetValue(m_Loads.NumRows, 2, w2)
    Call m_Loads.SetValue(m_Loads.NumRows, 3, 0)
    Call m_Loads.SetValue(m_Loads.NumRows, 4, 0)
    Call m_Loads.SetValue(m_Loads.NumRows, 5, x1)
    Call m_Loads.SetValue(m_Loads.NumRows, 6, x2)
    Call m_Loads.SetValue(m_Loads.NumRows, 7, Direction)
    
    'Add the fixed end reactions to the FER vector
    If Direction = "Transverse" Then
        Set m_FER = MAdd(m_FER, FER_LinLoad(w1, w2, x1, x2, Length))
    Else
        Set m_FER = MAdd(m_FER, FER_AxialLinLoad(w1, w2, x1, x2, Length))
    End If
    
End Sub

'Adds the fixed end reactions for a moment to the FER vector and places the load in the load vector
Public Sub AddMoment(Moment As Double, Location As Double)
    
    'Declare variables
    Dim NewRxns() As Variant
    Dim Temp() As Variant
    Dim i As Integer, j As Integer
    
    'Resize the "m_Loads" array
    Call m_Loads.Resize(m_Loads.NumRows + 1, 7)
    
    'Add the load to the "m_Loads" array
    Call m_Loads.SetValue(m_Loads.NumRows, 1, 0)
    Call m_Loads.SetValue(m_Loads.NumRows, 2, 0)
    Call m_Loads.SetValue(m_Loads.NumRows, 3, 0)
    Call m_Loads.SetValue(m_Loads.NumRows, 4, Moment)
    Call m_Loads.SetValue(m_Loads.NumRows, 5, Location)
    Call m_Loads.SetValue(m_Loads.NumRows, 6, 0)
    Call m_Loads.SetValue(m_Loads.NumRows, 7, vbNullString)
    
    'Add the fixed end reactions to the FER vector
    Set m_FER = MAdd(m_FER, FER_Moment(Moment, Location, Length))
    
End Sub

'Adds a global displacement to the global displacement vector
Public Sub AddDisplacement(Index As Integer, Displacement As Double)
    
    'Store the displacement in the global displacement vector
    m_D(Index, 1) = Displacement
    
End Sub

'Segments the member (and loads) into continuous segments
Public Sub SegmentMember()
    
    'Declare local variables
    Dim i As Long, j As Long
    Dim Discont() As Variant
    Dim w1 As Double, w2 As Double, P As Double, M As Double
    Dim x As Double, x1 As Double, x2 As Double
    Dim Direction As String
    
    'Estimate the number of continuous segments in the beam
    m_NumSegs = 1 + m_Loads.NumRows * 2
    
    'Start an array of discontinuity locations
    ReDim Discont(1 To m_NumSegs + 1)
    Discont(1) = 0
    Discont(2) = Length
    
    'Step through each load on the member
    For i = 1 To m_Loads.NumRows
        'Store the 2 discontinuity locations in the array
        Discont(i * 2 + 1) = m_Loads.GetValue(i, 5)
        Discont(i * 2 + 2) = m_Loads.GetValue(i, 6)
    Next i
    
    'Sort the array and eliminate duplicate values
    Call RemoveDuplicates(Discont)
    Call SortAscending(Discont)
    
    'Create an array of segments
    m_NumSegs = UBound(Discont) - LBound(Discont)
    ReDim m_Segments(1 To m_NumSegs)
    
    'Get the local end forces
    Dim EndForces As Matrix
    Set EndForces = LocalForces
    
    'Calculate the rotation and deflection at the start of the member
    Dim M1 As Double, M2 As Double, FEM1 As Double, FEM2 As Double, Delta1 As Double, Delta2 As Double
    M1 = LocalForces.GetValue(3, 1)
    M2 = LocalForces.GetValue(6, 1)
    FEM1 = m_FER.GetValue(3, 1)
    FEM2 = m_FER.GetValue(6, 1)
    Delta1 = LocalDisp.GetValue(2, 1)
    Delta2 = LocalDisp.GetValue(5, 1)
    m_Segments(1).theta1 = 1 / 3 * ((M1 - FEM1) * Length / (m_E * m_I) - (M2 - FEM2) * Length / (2 * m_E * m_I) + 3 * (Delta2 - Delta1) / Length)
    m_Segments(1).Delta1 = Delta1
    
    'Step through each segment
    For i = 1 To m_NumSegs
    
        'Determine the starting and ending points of the segment
        m_Segments(i).SegStart = Discont(i)
        m_Segments(i).SegEnd = Discont(i + 1)
        
        'Determine the segment start point
        x = Round(m_Segments(i).SegStart, 8)    'Starting point of the segment
        
        'Initialize loads on the segment
        m_Segments(i).w1 = 0                'Initializing w1 to zero
        m_Segments(i).w2 = 0                'Initializing w2 to zero
        
        'Initialize the rotation and deflection at the start of the segment
        If i > 1 Then
            m_Segments(i).theta1 = m_Segments(i - 1).Rotation(m_Segments(i - 1).Length, m_E * m_I)
            m_Segments(i).Delta1 = m_Segments(i - 1).Deflection(m_Segments(i - 1).Length, m_E * m_I)
        End If
        
        'Add the beam end forces to the segment
        m_Segments(i).p1 = EndForces.GetValue(1, 1)                             'Axial force at the start of the member
        m_Segments(i).V1 = EndForces.GetValue(2, 1)                             'Shear at the start of the member
        m_Segments(i).M1 = EndForces.GetValue(3, 1) - m_Segments(i).V1 * x      'Moment due to shear and moment at the start of the member
        
        'Step through each load on the member
        For j = 1 To m_Loads.NumRows
            
            'Find the parameters for the load
            w1 = m_Loads.GetValue(j, 1)                  'Distributed load (transverse or axial) start magnitude
            w2 = m_Loads.GetValue(j, 2)                  'Distributed load (transverse or axial) end magnitude
            P = m_Loads.GetValue(j, 3)                   'Concentrated force
            M = m_Loads.GetValue(j, 4)                   'Concentrated moment
            x1 = Round(m_Loads.GetValue(j, 5), 8)        'Load start location
            x2 = Round(m_Loads.GetValue(j, 6), 8)        'Load end location
            Direction = m_Loads.GetValue(j, 7)           'Load direction ("Transverse" or "Axial")
            
            'Determine the load type
            'Also, determine whether it affects (i.e. comes before or at the start of) segment "i"
            
            'Distributed loads
            If (w1 <> 0 Or w2 <> 0) And Round(x, 8) >= Round(x1, 8) Then
                
                'Determine whether the load is actually on segment "i"
                If Round(x, 8) < Round(x2, 8) Then
                    
                    If Direction = "Transverse" Then
                        'Break up the load and place it on the segment
                        m_Segments(i).w1 = m_Segments(i).w1 + (w2 - w1) / (x2 - x1) * (x - x1) + w1
                        m_Segments(i).w2 = m_Segments(i).w2 + (w2 - w1) / (x2 - x1) * (m_Segments(i).SegEnd - x1) + w1
                    ElseIf Direction = "Axial" Then
                        m_Segments(i).w1_Axial = m_Segments(i).w1_Axial + (w2 - w1) / (x2 - x1) * (x - x1) + w1
                        m_Segments(i).w2_Axial = m_Segments(i).w2_Axial + (w2 - w1) / (x2 - x1) * (m_Segments(i).SegEnd - x1) + w1
                    End If
                    
                    'Prior to this next step "w2" and "x2" correspond to the end of the segment. We only want to sum the load effects
                    'up to the start of the segment, so these values need adjustment before we sum the load.
                    w2 = (w2 - w1) / (x2 - x1) * (x - x1) + w1
                    x2 = x
                    
                End If
                
                If Direction = "Transverse" Then
                    'Sum the load effects (due to the distributed load) up to the start of the segment
                    m_Segments(i).V1 = m_Segments(i).V1 + (x2 - x1) * (w1 + w2) / 2
                    m_Segments(i).M1 = m_Segments(i).M1 - (x1 - x2) * (2 * w1 * x1 + w1 * x2 + w2 * x1 - 3 * w1 * x + 2 * w2 * x2 - 3 * w2 * x) / 6
                ElseIf Direction = "Axial" Then
                    m_Segments(i).p1 = m_Segments(i).p1 + (x2 - x1) * (w1 + w2) / 2
                End If
                
            'Point loads
            ElseIf P <> 0 And Round(x, 8) >= Round(x1, 8) Then
                
                If Direction = "Transverse" Then
                    'Find the shear and moment (due to the point load) at the start of the segment
                    m_Segments(i).V1 = m_Segments(i).V1 + P
                    m_Segments(i).M1 = m_Segments(i).M1 - P * (x - x1)
                ElseIf Direction = "Axial" Then
                    m_Segments(i).p1 = m_Segments(i).p1 + P
                End If
            
            'Moments
            ElseIf M <> 0 And Round(x, 8) >= Round(x1, 8) Then
                
                'Find the moment (due to the applied moment) at the start of the segment
                m_Segments(i).M1 = m_Segments(i).M1 + M
                
            End If
            
        Next j
        
    Next i
    
End Sub

'Returns the maximum axial force in the member
Public Function Pmax() As Double
    
    Dim i As Integer
    Dim Max As Double
    
    Max = m_Segments(1).p1
    For i = 1 To m_NumSegs
        If m_Segments(i).MaxAxial > Max Then
            Max = m_Segments(i).MaxAxial
        End If
    Next i
    
    Pmax = WorksheetFunction.Max(Max, LocalForces.GetValue(1, 1), -LocalForces.GetValue(4, 1))
    
End Function

'Returns the minimum axial force in the member
Public Function Pmin() As Double
    
    Dim i As Integer
    Dim Min As Double
    
    Min = m_Segments(1).p1
    For i = 1 To m_NumSegs
        If m_Segments(i).MinAxial < Min Then
            Min = m_Segments(i).MinAxial
        End If
    Next i
    
    Pmin = WorksheetFunction.Min(Min, LocalForces.GetValue(1, 1), -LocalForces.GetValue(4, 1))
    
End Function

'Returns the maximum shear in the member
Public Function Vmax() As Double
    
    Dim i As Integer
    Dim Max As Double
    
    Max = m_Segments(1).V1
    For i = 1 To m_NumSegs
        If m_Segments(i).MaxShear > Max Then
            Max = m_Segments(i).MaxShear
        End If
    Next i
    
    Vmax = WorksheetFunction.Max(Max, LocalForces.GetValue(2, 1), -LocalForces.GetValue(5, 1))
    
End Function

'Returns the minimum shear in the member
Public Function Vmin() As Double
    
    Dim i As Integer
    Dim Min As Double
    
    Min = m_Segments(1).V1
    For i = 1 To m_NumSegs
        If m_Segments(i).MinShear < Min Then
            Min = m_Segments(i).MinShear
        End If
    Next i
    
    Vmin = WorksheetFunction.Min(Min, LocalForces.GetValue(2, 1), -LocalForces.GetValue(5, 1))
    
End Function

'Returns the maximum moment in the member
Public Function Mmax() As Double
    
    Dim i As Integer
    Dim Max As Double
    
    Max = m_Segments(1).M1
    For i = 1 To m_NumSegs
        If m_Segments(i).MaxMoment > Max Then
            Max = m_Segments(i).MaxMoment
        End If
    Next i
    
    Mmax = WorksheetFunction.Max(Max, -LocalForces.GetValue(3, 1), LocalForces.GetValue(6, 1))
    
End Function

'Returns the minimum moment in the member
Public Function Mmin() As Double
    
    Dim i As Integer
    Dim Min As Double
    
    Min = m_Segments(1).M1
    For i = 1 To m_NumSegs
        If m_Segments(i).MinMoment < Min Then
            Min = m_Segments(i).MinMoment
        End If
    Next i
    
    Mmin = WorksheetFunction.Min(Min, -LocalForces.GetValue(3, 1), LocalForces.GetValue(6, 1))
    
End Function

'Returns the maximum displacement in the member
Public Function MaxDisplacement() As Double
    
    'Declare local variables
    Dim i As Integer, Delta As Double
    
    'Initialize the maximum deflection to zero
    MaxDisplacement = 0
    
    'Step through each beam segment in the member
    For i = 1 To m_NumSegs
        
        'Get the maximum deflection in the segment
        Delta = m_Segments(i).MaxDeflection(m_E * m_I)
        
        'Save the largest value
        If Delta > MaxDisplacement Then
            MaxDisplacement = Delta
        End If
        
    Next i
    
End Function

'Returns the minimum displacement in the member
Public Function MinDisplacement() As Double
    
    'Declare local variables
    Dim i As Integer, Delta As Double
    
    'Initialize the maximum deflection to zero
    MinDisplacement = 0
    
    'Step through each beam segment in the member
    For i = 1 To m_NumSegs
        
        'Get the minimum deflection in the segment
        Delta = m_Segments(i).MinDeflection(m_E * m_I)
        
        'Save the largest value
        If Delta < MinDisplacement Then
            MinDisplacement = Delta
        End If
        
    Next i
    
End Function

'Returns the axial force at a location on the member
Public Function Axial(x As Double) As Double
    
    'Declare local variables
    Dim i As Integer
    
    If Round(x, 8) = 0 Then
        Axial = LocalForces.GetValue(1, 1)
    ElseIf Round(x, 8) = Round(Length, 8) Then
        Axial = -LocalForces.GetValue(4, 1)
    Else
        'Find the segment the location falls within
        i = 1
        While Round(x, 8) > Round(m_Segments(i).SegEnd, 8)
            i = i + 1
        Wend
    
        Axial = m_Segments(i).Axial(x - m_Segments(i).SegStart)
    End If
    
End Function

'Returns the shear force at a location on the member
Public Function Shear(x As Double) As Double
    
    'Declare local variables
    Dim i As Integer
    
    If Round(x, 8) = 0 Then
        Shear = LocalForces.GetValue(2, 1)
    ElseIf Round(x, 8) = Round(Length, 8) Then
        Shear = -LocalForces.GetValue(5, 1)
    Else
        'Find the segment the location falls within
        i = 1
        While Round(x, 8) > Round(m_Segments(i).SegEnd, 8)
            i = i + 1
        Wend
    
        Shear = m_Segments(i).Shear(Round(x - m_Segments(i).SegStart, 8))
    End If
    
End Function

'Returns the moment at a location on the member
Public Function Moment(x As Double) As Double
    
    'Declare local variables
    Dim i As Integer
    
    If Round(x, 8) = 0 Then
        Moment = -LocalForces.GetValue(3, 1)
    ElseIf Round(x, 8) = Round(Length, 8) Then
        Moment = LocalForces.GetValue(6, 1)
    Else
        'Find the segment that "x" falls on
        i = 1
        While Round(x, 8) > Round(m_Segments(i).SegEnd, 8)
            i = i + 1
        Wend
    
        Moment = m_Segments(i).Moment(x - m_Segments(i).SegStart)
    End If
    
End Function

'Returns the deflection at a location on the member
Public Function Deflection(x As Double) As Double
    
    'Find the segment 'x' falls on
    Dim i As Integer
    i = 1
    While Round(x, 8) > Round(m_Segments(i).SegEnd, 8)
        i = i + 1
    Wend
    
    Deflection = m_Segments(i).Deflection(x - m_Segments(i).SegStart, m_E * m_I)
    
End Function
