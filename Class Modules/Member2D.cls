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
Public Name As String                       'Member name
Public MemberID As Integer                  'Member ID that can be assigned in code-behind to help track the member
Public iNode As Node2D                      'i-node
Public jNode As Node2D                      'j-node

'Member properties
Public Elasticity As Double                 'Modulus of elasticity
Public Inertia As Double                    'Moment of inertia
Public Area As Double                       'Cross sectional area

'End releases
Private m_Releases(1 To 6) As Boolean       'Array representing member end releases
Private m_Condense As Boolean               'Flag indicating whether static condensation is required

'Loads
Private m_PtLoads As New EZArray            'An array of point loads and moments applied to the element (Direction, P, x, LoadCase="Case 1") or (Direction, M, x, LoadCase="Case 1")
Private m_DistLoads As New EZArray          'A list of linear distributed loads applied to the element (Direction, w1, w2, x1, x2, LoadCase='Case 1')

'Load combinations
Public LoadCombos As Object                 'A dictionary of load combinations in the model this member belongs to
Private m_SolvedCombo As String             'The name of the currently solve load combination

'Internal forces
Private m_Segments() As New Segment         'Array representing mathematically continuous segments of the member
Public NumSegs As Integer                   'Number of mathematically continuous segments in the member

Property Get Segments() As Segment()
    Segments = m_Segments
End Property

Property Get Length() As Double
    Length = Sqr((jNode.XCoord - iNode.XCoord) ^ 2 + (jNode.YCoord - iNode.YCoord) ^ 2)
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
        
    'Initialize the solved load combination value
    m_SolvedCombo = vbNullString
    
End Sub

'Deletes all loads from the element
Public Sub ClearLoads()
    
    'Declare local variables
    Dim i As Integer
    
    'Initialize the segment count to zero
    NumSegs = 0
    
    'Delete the old array of loads
    Set m_PtLoads = Nothing
    Set m_DistLoads = Nothing
    
    'Create a new array of loads
    Set m_PtLoads = New EZArray
    Set m_DistLoads = New EZArray
    
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
    DirCos = (jNode.XCoord - iNode.XCoord) / Length
    DirSin = (jNode.YCoord - iNode.YCoord) / Length
    
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

Private Function LocalFER_Unc(Optional ComboName As Variant = "Combo 1") As Matrix

    'Initialize the fixed end reaction vector
    Dim Result As New Matrix
    Call Result.Resize(6, 1)

    'Loop through each load case in the load combination
    Dim LoadCase As Variant, Factor As Double, i As Long
    Dim P As Double, w1 As Double, w2 As Double, x As Double, x1 As Double, x2 As Double, L As Double
    For Each LoadCase In LoadCombos(ComboName).Factors.Keys

        'Get the load factor for this load case
        Factor = LoadCombos(ComboName).Factors(LoadCase)

        'Sum the fixed end reactions for the point loads and moments
        For i = 1 To m_PtLoads.NumRows

            'Check if the current point load corresponds to the current load case
            If m_PtLoads.GetValue(i, 4) = LoadCase Then

                'Get the load parameters
                P = m_PtLoads.GetValue(i, 2) * Factor 'Note that 'P' could be a force or moment
                x = m_PtLoads.GetValue(i, 3)
                L = Length

                If m_PtLoads.GetValue(i, 1) = "Axial" Then
                    Set Result = MAdd(Result, FER_AxialPtLoad(P, x, L))
                ElseIf m_PtLoads.GetValue(i, 1) = "Transverse" Then
                    Set Result = MAdd(Result, FER_PtLoad(P, x, L))
                ElseIf m_PtLoads.GetValue(i, 1) = "Moment" Then
                    Set Result = MAdd(Result, FER_Moment(P, x, L))
                End If

            End If

        Next i

        'Sum the fixed end reactions for the distributed loads
        For i = 1 To m_DistLoads.NumRows

            'Check if the current distributed load corresponds to the current load case
            If m_DistLoads.GetValue(i, 6) = LoadCase Then

                'Get the load parameters
                w1 = m_DistLoads.GetValue(i, 2) * Factor 'Note that 'w1' could be a transverse or axial load
                w2 = m_DistLoads.GetValue(i, 3) * Factor 'Note that 'w2' could be a transverse or axial load
                x1 = m_DistLoads.GetValue(i, 4)
                x2 = m_DistLoads.GetValue(i, 5)
                L = Length
                
                If m_DistLoads.GetValue(i, 1) = "Axial" Then
                    Set Result = MAdd(Result, FER_AxialLinLoad(w1, w2, x1, x2, L))
                ElseIf m_DistLoads.GetValue(i, 1) = "Transverse" Then
                    Set Result = MAdd(Result, FER_LinLoad(w1, w2, x1, x2, L))
                End If

            End If

        Next i

    Next LoadCase
    
    Set LocalFER_Unc = Result

End Function

'Returns the local fixed end reaction vector
Public Function LocalFER(Optional ComboName As Variant = "Combo 1", Optional Condensed As Boolean = True) As Matrix
    
    'Determine if the local fixed end reaction vector requires static condensation
    If Condensed = True Then
        
        'Condense the local fixed end reaction vector
        Set LocalFER = CondenseFER(LocalFER_Unc(ComboName), LocalStiff(False), m_Releases)
        
    Else
    
        'Get the local fixed end reaction vector
        Set LocalFER = LocalFER_Unc(ComboName)
        
    End If
    
End Function

'Returns the global fixed end reaction vector
Public Function GlobalFER(Optional ComboName As Variant = "Combo 1") As Matrix
    
    'Transpose the transformation matrix
    Dim M1 As Matrix
    Set M1 = MTranspose(TransMatrix)
    
    'Get the local fixed end reaction vector
    Dim M2 As Matrix
    Set M2 = LocalFER(ComboName, True)
    
    'Multiply the local fixed end reaction vector by the transposed transformation matrix
    Set GlobalFER = MMultiply(M1, M2)
    
End Function

'Returns the local displacement vector
Public Function LocalDisp(Optional ComboName As Variant = "Combo 1") As Matrix
    
    'Calculate the local displacement vector
    Set LocalDisp = MMultiply(TransMatrix, GlobalDisp(ComboName))
    
End Function

'Returns the global displacement vector
'The nodal displacements must have been solved for prior to calling this function
Public Function GlobalDisp(Optional ComboName As Variant = "Combo 1") As Matrix
    
    'Generate the global displacement vector
    Set GlobalDisp = New Matrix
    Call GlobalDisp.Resize(6, 1, False)
    Call GlobalDisp.SetValue(1, 1, iNode.DX(ComboName))
    Call GlobalDisp.SetValue(2, 1, iNode.DY(ComboName))
    Call GlobalDisp.SetValue(3, 1, iNode.RZ(ComboName))
    Call GlobalDisp.SetValue(4, 1, jNode.DX(ComboName))
    Call GlobalDisp.SetValue(5, 1, jNode.DY(ComboName))
    Call GlobalDisp.SetValue(6, 1, jNode.RZ(ComboName))
    
End Function

'Returns the local member end force vector
Public Function LocalForces(Optional ComboName As Variant = "Combo 1") As Matrix
    
    'Declare local variables
    Dim M1 As Matrix, M2 As Matrix, M3 As Matrix, M4 As Matrix
    
    'Get the local stiffness matrix
    Set M1 = LocalStiff(True)
    
    'Get the local displacement vector
    Set M2 = LocalDisp(ComboName)
    
    'Get the local FER vector
    Set M3 = LocalFER(ComboName, True)
    
    'Multiply the local displacement vector by the local stiffness matrix
    Set M4 = MMultiply(M1, M2)
    
    'Add the result of the previous operation to the local FER vector
    Set LocalForces = MAdd(M4, M3)
        
End Function

'Returns the global member end force vector
Public Function GlobalForces(Optional ComboName As Variant = "Combo 1") As Matrix
    
    'Declare local variables
    Dim M1 As Matrix, M2 As Matrix
    
    'Transpose the transformation matrix
    Set M1 = MTranspose(TransMatrix)
    
    'Get the local member end force vector
    Set M2 = LocalForces(ComboName)
    
    'Multiply the local member end force vector by the transposed transformation matrix
    Set GlobalForces = MMultiply(M1, M2)
        
End Function

'Adds a point load to the member
Public Sub AddPtLoad(P As Double, x As Double, Direction As String, Optional LoadCase As Variant = "Case 1")
    
    'Resize the "m_PtLoads" EZArray
    Call m_PtLoads.Resize(m_PtLoads.NumRows + 1, 4, True)
    
    'Add the load to the "m_PtLoads" EZArray
    Call m_PtLoads.SetValue(m_PtLoads.NumRows, 1, Direction)
    Call m_PtLoads.SetValue(m_PtLoads.NumRows, 2, P)
    Call m_PtLoads.SetValue(m_PtLoads.NumRows, 3, x)
    Call m_PtLoads.SetValue(m_PtLoads.NumRows, 4, LoadCase)
    
End Sub

'Adds a distributed load to the member
Public Sub AddLinLoad(w1 As Double, w2 As Double, x1 As Double, x2 As Double, Direction As String, Optional LoadCase As Variant = "Case 1")
    
    'Resize the "m_DistLoads" EZArray
    Call m_DistLoads.Resize(m_DistLoads.NumRows + 1, 6, True)
    
    'Add the load to the "m_DistLoads" array
    Call m_DistLoads.SetValue(m_DistLoads.NumRows, 1, Direction)
    Call m_DistLoads.SetValue(m_DistLoads.NumRows, 2, w1)
    Call m_DistLoads.SetValue(m_DistLoads.NumRows, 3, w2)
    Call m_DistLoads.SetValue(m_DistLoads.NumRows, 4, x1)
    Call m_DistLoads.SetValue(m_DistLoads.NumRows, 5, x2)
    Call m_DistLoads.SetValue(m_DistLoads.NumRows, 6, LoadCase)
    
End Sub

'Adds the fixed end reactions for a moment to the FER vector and places the load in the load vector
Public Sub AddMoment(m As Double, x As Double, Optional LoadCase As Variant = "Case 1")
    
    'Resize the "m_PtLoads" EZArray
    Call m_PtLoads.Resize(m_PtLoads.NumRows + 1, 4, True)
    
    'Add the load to the "m_PtLoads" EZArray
    Call m_PtLoads.SetValue(m_PtLoads.NumRows, 1, "Moment")
    Call m_PtLoads.SetValue(m_PtLoads.NumRows, 2, m)
    Call m_PtLoads.SetValue(m_PtLoads.NumRows, 3, x)
    Call m_PtLoads.SetValue(m_PtLoads.NumRows, 4, LoadCase)
    
End Sub

'Adds a global displacement to the global displacement vector
Public Sub AddDisplacement(Index As Integer, Displacement As Double)
    
    'Store the displacement in the global displacement vector
    m_D(Index, 1) = Displacement
    
End Sub

'Divides the element up into mathematically continuous segments
Public Sub SegmentMember(Optional ComboName As Variant = "Combo 1")

    'Create an array of member load discontinuity locations
    Dim Discont() As Variant
    ReDim Discont(1 To 2 + m_PtLoads.NumRows + m_DistLoads.NumRows * 2)
    Discont(1) = 0
    Discont(2) = Length
    
    'Declare/initialize variables used in the next two `For` loops
    Dim i As Long
    Dim NumPtLoads As Integer
    NumPtLoads = m_PtLoads.NumRows

    'Step through each point load on the member
    For i = 1 To m_PtLoads.NumRows
        'Store the discontinuity location in the array
        Discont(2 + i) = m_PtLoads.GetValue(i, 3)
    Next i
    
    'Step through each distributed load on the member
    For i = 1 To m_DistLoads.NumRows
        'Store the discontinuity locations in the array
        Discont(2 + NumPtLoads + (2 * i - 1)) = m_DistLoads.GetValue(i, 4)
        Discont(2 + NumPtLoads + (2 * i)) = m_DistLoads.GetValue(i, 5)
    Next i

    'Eliminate duplicate values and then sort the array
    Call RemoveDuplicates(Discont)
    Call SortAscending(Discont)
    
    'Clear out old data from any previous analyses
    Erase m_Segments

    'Create an array of member segments, with one segment between each discontinuity
    NumSegs = UBound(Discont) - LBound(Discont)
    ReDim m_Segments(1 To NumSegs)
    
    'Get the member local end forces, local fixed end reactions, and local displacements
    Dim EndForces As Matrix, FE_Rxns As Matrix, EndDisp As Matrix
    Set EndForces = LocalForces(ComboName)
    Set FE_Rxns = LocalFER_Unc(ComboName)
    Set EndDisp = LocalDisp(ComboName)
    
    'Initialize the first segment by calculating and storing the slope and deflection at the start of the member
    Dim M1 As Double, M2 As Double, FEM1 As Double, FEM2 As Double, Delta1 As Double, Delta2 As Double
    M1 = EndForces.GetValue(3, 1)
    M2 = EndForces.GetValue(6, 1)
    FEM1 = FE_Rxns.GetValue(3, 1)
    FEM2 = FE_Rxns.GetValue(6, 1)
    Delta1 = EndDisp.GetValue(2, 1)
    Delta2 = EndDisp.GetValue(5, 1)
    m_Segments(1).Delta1 = Delta1
    m_Segments(1).theta1 = 1 / 3 * ((M1 - FEM1) * Length / (Elasticity * Inertia) - (M2 - FEM2) * Length / (2 * Elasticity * Inertia) + 3 * (Delta2 - Delta1) / Length)
    
    'TODO:
    'Once axial deflections have been added to the `Segment` class, uncomment the next line
    'm_Segments(1).Delta_x1 = EndDisp(1, 1)

    'Declare variables used in the loops below
    Dim j As Long
    Dim P As Double, w1 As Double, w2 As Double
    Dim x As Double, x1 As Double, x2 As Double
    Dim Direction As String
    Dim Factor As Double

    'Add loads to each segment
    Dim LoadCase As Variant
    For i = 1 To NumSegs
    
        'Determine the starting and ending points of the segment
        m_Segments(i).SegStart = Discont(i)
        m_Segments(i).SegEnd = Discont(i + 1)
        
        'Store the rounded starting point of the segment in the variable `x` to simplify the expressions below
        'It is easier to use `x` than `Round(m_Segments(i).SegStart, 8)` or `Round(Discont(i), 8)` every time
        x = Round(m_Segments(i).SegStart, 8)
        
        'Initialize the distirbuted loads on the segment to zero
        m_Segments(i).w1 = 0  'Initializing w1 to zero
        m_Segments(i).w2 = 0  'Initializing w2 to zero
        
        'Initialize the slope and displacement at the start of the segment
        If i > 1 Then  'The first segment was already initialized prior to entering this `For` loop
            m_Segments(i).theta1 = m_Segments(i - 1).Slope(m_Segments(i - 1).Length, Elasticity * Inertia)
            m_Segments(i).Delta1 = m_Segments(i - 1).Deflection(m_Segments(i - 1).Length, Elasticity * Inertia)
        End If
        
        'Add the effects of the beam end forces to the segment
        m_Segments(i).p1 = EndForces.GetValue(1, 1)                                 'Axial force at the start of the member
        m_Segments(i).V1 = EndForces.GetValue(2, 1)                                 'Shear at the start of the member
        m_Segments(i).M1 = EndForces.GetValue(3, 1) - EndForces.GetValue(2, 1) * x  'Moment due to shear and moment at the start of the member
        
        'Step through each load case in the specified load combination
        For Each LoadCase In LoadCombos(ComboName).Factors.Keys()

            'Get the load factor for this load case
            Factor = LoadCombos(ComboName).Factors(LoadCase)

            'Add the effects of point loads occuring prior to this segment
            For j = 1 To m_PtLoads.NumRows

                P = m_PtLoads.GetValue(j, 2) * Factor    'Note that P can either be a point load or a concentrated moment
                x1 = Round(m_PtLoads.GetValue(j, 3), 8)  'Load location relative to the start of the member
                Direction = m_PtLoads.GetValue(j, 1)     'Load direction ("Transverse", "Axial", or "Moment")

                'Determine if the load affects this segment
                If x1 <= x And LoadCase = m_PtLoads.GetValue(j, 4) Then
                    
                    If Direction = "Axial" Then
                        m_Segments(i).p1 = m_Segments(i).p1 + P
                    ElseIf Direction = "Transverse" Then
                        m_Segments(i).V1 = m_Segments(i).V1 + P
                        m_Segments(i).M1 = m_Segments(i).M1 - P * (x - x1)
                    ElseIf Direction = "Moment" Then
                        m_Segments(i).M1 = m_Segments(i).M1 + P
                    End If
                
                End If
            
            Next j

            'Add the effects of distributed loads starting prior to this segment
            For j = 1 To m_DistLoads.NumRows
            
                'Find the parameters for the load
                w1 = m_DistLoads.GetValue(j, 2) * Factor   'Distributed load (transverse or axial) start magnitude
                w2 = m_DistLoads.GetValue(j, 3) * Factor   'Distributed load (transverse or axial) end magnitude
                x1 = Round(m_DistLoads.GetValue(j, 4), 8)  'Load start location
                x2 = Round(m_DistLoads.GetValue(j, 5), 8)  'Load end location
                Direction = m_DistLoads.GetValue(j, 1)     'Load direction ("Transverse" or "Axial")
                
                'Determine if the load affects this segment
                If x1 <= x And LoadCase = m_DistLoads.GetValue(j, 6) Then

                    If Direction = "Axial" Then
                
                        'Determine whether the load ends after the start of the segment
                        If x2 > x Then

                            'Break up the load and place it on the segment
                            'Note that `w1` and `w2` are really axial loads here
                            m_Segments(i).w1_Axial = m_Segments(i).w1_Axial + (w2 - w1) / (x2 - x1) * (x - x1) + w1
                            m_Segments(i).w2_Axial = m_Segments(i).w2_Axial + (w2 - w1) / (x2 - x1) * (m_Segments(i).SegEnd - x1) + w1

                            'Replace `w2` and `x2` with the values for the load at the start of the segment
                            w2 = w1 + (w2 - w1) / (x2 - x1) * (x - x1)
                            x2 = x
                        
                        End If

                        'Sum the axial forces at the start of the segment
                        m_Segments(i).p1 = m_Segments(i).p1 + (w1 + w2) / 2 * (x2 - x1)

                    ElseIf Direction = "Transverse" Then

                        'Determine whether the load ends after the start of the segment
                        If x2 > Round(x, 8) Then

                            'Break up the load and place it on the segment
                            m_Segments(i).w1 = m_Segments(i).w1 + (w2 - w1) / (x2 - x1) * (m_Segments(i).SegStart - x1) + w1
                            m_Segments(i).w2 = m_Segments(i).w2 + (w2 - w1) / (x2 - x1) * (m_Segments(i).SegEnd - x1) + w1

                            'Replace `w2` and `x2` with the values for the load at the start of the segment
                            w2 = w1 + (w2 - w1) / (x2 - x1) * (x - x1)
                            x2 = x

                        End If
                        
                        'Sum the shears and moments at the start of the segment
                        m_Segments(i).V1 = m_Segments(i).V1 + (w1 + w2) / 2 * (x2 - x1)
                        m_Segments(i).M1 = m_Segments(i).M1 - (x1 - x2) * (2 * w1 * x1 - 3 * w1 * x + w1 * x2 + w2 * x1 - 3 * w2 * x + 2 * w2 * x2) / 6
                    
                    End If
                
                End If
            
            Next j
            
        Next LoadCase
            
    Next i
    
End Sub

'Returns the maximum axial force in the member
Public Function Pmax(Optional ComboName As Variant = "Combo 1") As Double
    
    Dim i As Integer
    Dim Max As Double
    
    'Segment the member if necessary
    If m_SolvedCombo <> ComboName Then
        Call SegmentMember(ComboName)
        m_SolvedCombo = ComboName
    End If
    
    Max = m_Segments(1).p1
    For i = 1 To NumSegs
        If m_Segments(i).MaxAxial > Max Then
            Max = m_Segments(i).MaxAxial
        End If
    Next i
        
    Dim LocalF As Matrix
    Set LocalF = LocalForces(ComboName)
    Pmax = WorksheetFunction.Max(Max, LocalF.GetValue(1, 1), -LocalF.GetValue(4, 1))
    
End Function

'Returns the minimum axial force in the member
Public Function Pmin(Optional ComboName As Variant = "Combo 1") As Double
    
    Dim i As Integer
    Dim Min As Double
    
    'Segment the member if necessary
    If m_SolvedCombo <> ComboName Then
        Call SegmentMember(ComboName)
        m_SolvedCombo = ComboName
    End If
    
    Min = m_Segments(1).p1
    For i = 1 To NumSegs
        If m_Segments(i).MinAxial < Min Then
            Min = m_Segments(i).MinAxial
        End If
    Next i
        
    Dim LocalF As Matrix
    Set LocalF = LocalForces(ComboName)
    Pmin = WorksheetFunction.Min(Min, LocalF.GetValue(1, 1), -LocalF.GetValue(4, 1))
    
End Function

'Returns the maximum shear in the member
Public Function Vmax(Optional ComboName As Variant = "Combo 1") As Double
    
    Dim i As Integer
    Dim Max As Double
    
    'Segment the member if necessary
    If m_SolvedCombo <> ComboName Then
        Call SegmentMember(ComboName)
        m_SolvedCombo = ComboName
    End If
    
    Max = m_Segments(1).V1
    For i = 1 To NumSegs
        If m_Segments(i).MaxShear > Max Then
            Max = m_Segments(i).MaxShear
        End If
    Next i
        
    Dim LocalF As Matrix
    Set LocalF = LocalForces(ComboName)
    Vmax = WorksheetFunction.Max(Max, LocalF.GetValue(2, 1), -LocalF.GetValue(5, 1))
    
End Function

'Returns the minimum shear in the member
Public Function Vmin(Optional ComboName As Variant = "Combo 1") As Double
    
    Dim i As Integer
    Dim Min As Double
    
    'Segment the member if necessary
    If m_SolvedCombo <> ComboName Then
        Call SegmentMember(ComboName)
        m_SolvedCombo = ComboName
    End If
    
    Min = m_Segments(1).V1
    For i = 1 To NumSegs
        If m_Segments(i).MinShear < Min Then
            Min = m_Segments(i).MinShear
        End If
    Next i
        
    Dim LocalF As Matrix
    Set LocalF = LocalForces(ComboName)
    Vmin = WorksheetFunction.Min(Min, LocalF.GetValue(2, 1), -LocalF.GetValue(5, 1))
    
End Function

'Returns the maximum moment in the member
Public Function Mmax(Optional ComboName As Variant = "Combo 1") As Double
    
    Dim i As Integer
    Dim Max As Double
    
    'Segment the member if necessary
    If m_SolvedCombo <> ComboName Then
        Call SegmentMember(ComboName)
        m_SolvedCombo = ComboName
    End If
    
    Max = m_Segments(1).M1
    For i = 1 To NumSegs
        If m_Segments(i).MaxMoment > Max Then
            Max = m_Segments(i).MaxMoment
        End If
    Next i
    
    Dim LocalF As Matrix
    Set LocalF = LocalForces(ComboName)
    Mmax = WorksheetFunction.Max(Max, LocalF.GetValue(3, 1), -LocalF.GetValue(6, 1))
    
End Function

'Returns the minimum moment in the member
Public Function Mmin(Optional ComboName As Variant = "Combo 1") As Double
    
    Dim i As Integer
    Dim Min As Double
    
    'Segment the member if necessary
    If m_SolvedCombo <> ComboName Then
        Call SegmentMember(ComboName)
        m_SolvedCombo = ComboName
    End If
    
    Min = m_Segments(1).M1
    For i = 1 To NumSegs
        If m_Segments(i).MinMoment < Min Then
            Min = m_Segments(i).MinMoment
        End If
    Next i
    
    Dim LocalF As Matrix
    Set LocalF = LocalForces(ComboName)
    Mmin = WorksheetFunction.Min(Min, LocalF.GetValue(3, 1), -LocalF.GetValue(6, 1))
    
End Function

'Returns the maximum displacement in the member
Public Function MaxDisplacement(Optional ComboName As Variant = "Combo 1") As Double
    
    'Declare local variables
    Dim i As Integer, Delta As Double
    
    'Segment the member if necessary
    If m_SolvedCombo <> ComboName Then
        Call SegmentMember(ComboName)
        m_SolvedCombo = ComboName
    End If
    
    'Initialize the maximum deflection to zero
    MaxDisplacement = 0
    
    'Step through each beam segment in the member
    For i = 1 To NumSegs
        
        'Get the maximum deflection in the segment
        Delta = m_Segments(i).MaxDeflection(Elasticity * Inertia)
        
        'Save the largest value
        If Delta > MaxDisplacement Then
            MaxDisplacement = Delta
        End If
        
    Next i
    
End Function

'Returns the minimum displacement in the member
Public Function MinDisplacement(Optional ComboName As Variant = "Combo 1") As Double
    
    'Declare local variables
    Dim i As Integer, Delta As Double
    
    'Segment the member if necessary
    If m_SolvedCombo <> ComboName Then
        Call SegmentMember(ComboName)
        m_SolvedCombo = ComboName
    End If
    
    'Initialize the maximum deflection to zero
    MinDisplacement = 0
    
    'Step through each beam segment in the member
    For i = 1 To NumSegs
        
        'Get the minimum deflection in the segment
        Delta = m_Segments(i).MinDeflection(Elasticity * Inertia)
        
        'Save the largest value
        If Delta < MinDisplacement Then
            MinDisplacement = Delta
        End If
        
    Next i
    
End Function

'Returns the axial force at a location on the member
Public Function Axial(x As Double, Optional ComboName As Variant = "Combo 1") As Double
    
    'Declare local variables
    Dim i As Integer
    
    'Segment the member if necessary
    If m_SolvedCombo <> ComboName Then
        Call SegmentMember(ComboName)
        m_SolvedCombo = ComboName
    End If
    
    If Round(x, 8) = 0 Then
        Axial = LocalForces(ComboName).GetValue(1, 1)
    ElseIf Round(x, 8) = Round(Length, 8) Then
        Axial = -LocalForces(ComboName).GetValue(4, 1)
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
Public Function Shear(x As Double, Optional ComboName As Variant = "Combo 1") As Double
    
    'Declare local variables
    Dim i As Integer
    
    'Segment the member if necessary
    If m_SolvedCombo <> ComboName Then
        Call SegmentMember(ComboName)
        m_SolvedCombo = ComboName
    End If
    
    If Round(x, 8) = 0 Then
        Shear = LocalForces(ComboName).GetValue(2, 1)
    ElseIf Round(x, 8) = Round(Length, 8) Then
        Shear = -LocalForces(ComboName).GetValue(5, 1)
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
Public Function Moment(x As Double, Optional ComboName As Variant = "Combo 1") As Double
    
    'Declare local variables
    Dim i As Integer
    
    'Segment the member if necessary
    If m_SolvedCombo <> ComboName Then
        Call SegmentMember(ComboName)
        m_SolvedCombo = ComboName
    End If
    
    If Round(x, 8) = 0 Then
        Moment = LocalForces(ComboName).GetValue(3, 1)
    ElseIf Round(x, 8) = Round(Length, 8) Then
        Moment = -LocalForces(ComboName).GetValue(6, 1)
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
Public Function Deflection(x As Double, Optional ComboName As Variant = "Combo 1") As Double
    
    'Segment the member if necessary
    If m_SolvedCombo <> ComboName Then
        Call SegmentMember(ComboName)
        m_SolvedCombo = ComboName
    End If
    
    'Find the segment 'x' falls on
    Dim i As Integer
    i = 1
    While Round(x, 8) > Round(m_Segments(i).SegEnd, 8)
        i = i + 1
    Wend
    
    Deflection = m_Segments(i).Deflection(x - m_Segments(i).SegStart, Elasticity * Inertia)
    
End Function
