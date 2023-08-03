VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Segment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Segment.cls
'Represents a mathematically continuous beam segment

'Class variables
Public SegStart As Double  'Segment start location relative to the start of the beam
Public SegEnd As Double    'Segment end location relative to the start of the beam
Public w1 As Double        'Distributed load at the start of the segment
Public w2 As Double        'Distributed load at the end of the segement
Public w1_Axial As Double  'Distributed axial load at the start of the segment
Public w2_Axial As Double  'Distributed axial load at the end of the segement
Public p1 As Double        'Axial load at the start of the segment
Public V1 As Double        'Shear at the start of the segment
Public M1 As Double        'Moment at the start of the segment
Public theta1 As Double    'Slope at the start of the segment
Public Delta1 As Double    'Deflection at the start of the segment

'Returns the length of the segment
Public Function Length() As Double
    Length = SegEnd - SegStart
End Function

'Returns the axial force at a location on the segment
Public Function Axial(x As Double) As Double
    
    'Return the axial force at the location
    Axial = p1 + w1_Axial * x + (w2_Axial - w1_Axial) / (2 * Length) * x ^ 2
    
End Function

'Returns the shear force at a location on the segment
Public Function Shear(x As Double) As Double
    
    'Calculate the shear at the location
    Shear = V1 + w1 * x + (w2 - w1) / (2 * Length) * x ^ 2
    
End Function

'Returns the moment at a location on the segment
Public Function Moment(x As Double) As Double
    
    'Return the moment at the location
    Moment = M1 - V1 * x - w1 * x ^ 2 / 2 - (w2 - w1) / (6 * Length) * x ^ 3
    
End Function

'Returns the rotation at a location on the segment
Public Function Slope(x As Double, EI As Double) As Double
    
    'Return the rotation at the location
    Slope = theta1 - (M1 * x - V1 * x ^ 2 / 2 - w1 * x ^ 3 / 6 - (w2 - w1) / (24 * Length) * x ^ 4) / EI
    
End Function

'Returns the deflection at a location on the segment
Public Function Deflection(x As Double, EI As Double) As Double

    'Return the deflection at the location
    Deflection = Delta1 + theta1 * x + (-M1 * x ^ 2 / 2 + V1 * x ^ 3 / 6 + w1 * x ^ 4 / 24 + (w2 - w1) / (120 * Length) * x ^ 5) / EI
    
End Function

'Returns the maximum axial force in the segment
Public Function MaxAxial() As Double
    
    'Declare local variables
    Dim x1 As Double, x2 As Double, x3 As Double
    Dim p1 As Double, p2 As Double, P3 As Double
    
    'Determine possible locations of maximum axial force
    If w1_Axial - w2_Axial <> 0 Then
        x1 = Length * w1_Axial / (w1_Axial - w2_Axial)
    Else
        x1 = 0
    End If
    
    If Round(x1, 8) < 0 Or Round(x1, 8) > Length Then
        x1 = 0
    End If
    
    x2 = 0
    x3 = Length
    
    'Find the axial force at each location of interest
    p1 = Axial(x1)
    p2 = Axial(x2)
    P3 = Axial(x3)
    
    'Return the maximum axial force
    MaxAxial = WorksheetFunction.Max(p1, p2, P3)
    
End Function

'Returns the minimum axial force in the segment
Public Function MinAxial() As Double
    
    'Declare local variables
    Dim x1 As Double, x2 As Double, x3 As Double
    Dim p1 As Double, p2 As Double, P3 As Double
    
    'Determine possible locations of minimum axial force
    If w1_Axial - w2_Axial <> 0 Then
        x1 = Length * w1_Axial / (w1_Axial - w2_Axial)
    Else
        x1 = 0
    End If
    
    If Round(x1, 8) < 0 Or Round(x1, 8) > Length Then
        x1 = 0
    End If
    
    x2 = 0
    x3 = Length
    
    'Find the axial force at each location of interest
    p1 = Axial(x1)
    p2 = Axial(x2)
    P3 = Axial(x3)
    
    'Return the minimum axial force
    MinAxial = WorksheetFunction.Min(p1, p2, P3)
    
End Function

'Returns the maximum shear in the segment
Public Function MaxShear() As Double
    
    'Declare local variables
    Dim x1 As Double, x2 As Double, x3 As Double
    Dim V1 As Double, V2 As Double, V3 As Double
    
    'Determine possible locations of maximum shear
    If w1 - w2 = 0 Then
        x1 = 0
    Else
        x1 = w1 * Length / (w1 - w2)
    End If
    
    If Round(x1, 8) < 0 Or Round(x1, 8) > Length Then
        x1 = 0
    End If
    
    x2 = 0
    x3 = Length
    
    'Find the shear at each location of interest
    V1 = Shear(x1)
    V2 = Shear(x2)
    V3 = Shear(x3)
    
    'Return the maximum shear
    MaxShear = WorksheetFunction.Max(V1, V2, V3)
    
End Function

'Returns the minimum shear in the segment
Public Function MinShear() As Double
    
    'Declare local variables
    Dim x1 As Double, x2 As Double, x3 As Double
    Dim V1 As Double, V2 As Double, V3 As Double
    
    'Determine possible locations of minimum shear
    If w1 - w2 = 0 Then
        x1 = 0
    Else
        x1 = w1 * Length / (w1 - w2)
    End If
    
    If Round(x1, 8) < 0 Or Round(x1, 8) > Length Then
        x1 = 0
    End If
    
    x2 = 0
    x3 = Length
    
    'Find the shear at each location of interest
    V1 = Shear(x1)
    V2 = Shear(x2)
    V3 = Shear(x3)
    
    'Return the minimum shear
    MinShear = WorksheetFunction.Min(V1, V2, V3)
    
End Function

'Returns the maximum moment in the segment
Public Function MaxMoment() As Double

    'Declare local variables
    Dim A As Double, b As Double, c As Double
    Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double
    Dim M1 As Double, M2 As Double, M3 As Double, M4 As Double
    
    'Find the quadratic equation parameters
    A = (w1 - w2) / (2 * Length)
    b = -w1
    c = -V1
    
    'Determine possible locations of maximum moment
    If A = 0 Then
        If b <> 0 Then
            x1 = -c / b
        Else
            x1 = 0
        End If
        x2 = 0
    ElseIf b ^ 2 - 4 * A * c < 0 Then
        x1 = 0
        x2 = 0
    Else
        x1 = (-b + (b ^ 2 - 4 * A * c) ^ 0.5) / (2 * A)
        x2 = (-b - (b ^ 2 - 4 * A * c) ^ 0.5) / (2 * A)
    End If
    
    x3 = 0
    x4 = Length
    
    If Round(x1, 8) < 0 Or Round(x1, 8) > Length Then
        x1 = 0
    End If
    
    If Round(x2, 8) < 0 Or Round(x2, 8) > Length Then
        x2 = 0
    End If
    
    'Find the moment at each location of interest
    M1 = Moment(x1)
    M2 = Moment(x2)
    M3 = Moment(x3)
    M4 = Moment(x4)
    
    'Return the maximum moment
    MaxMoment = WorksheetFunction.Max(M1, M2, M3, M4)
        
End Function

'Returns the minimum moment in the segment
Public Function MinMoment() As Double

    'Declare local variables
    Dim A As Double, b As Double, c As Double
    Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double
    Dim M1 As Double, M2 As Double, M3 As Double, M4 As Double
    
    'Find the quadratic equation parameters
    A = (w1 - w2) / (2 * Length)
    b = -w1
    c = -V1
    
    'Determine possible locations of minimum moment
    If A = 0 Then
        If b <> 0 Then
            x1 = -c / b
        Else
            x1 = 0
        End If
        x2 = 0
    ElseIf b ^ 2 - 4 * A * c < 0 Then
        x1 = 0
        x2 = 0
    Else
        x1 = (-b + (b ^ 2 - 4 * A * c) ^ 0.5) / (2 * A)
        x2 = (-b - (b ^ 2 - 4 * A * c) ^ 0.5) / (2 * A)
    End If
    
    x3 = 0
    x4 = Length
    
    If Round(x1, 8) < 0 Or Round(x1, 8) > Length Then
        x1 = 0
    End If
    
    If Round(x2, 8) < 0 Or Round(x2, 8) > Length Then
        x2 = 0
    End If
    
    'Find the moment at each location of interest
    M1 = Moment(x1)
    M2 = Moment(x2)
    M3 = Moment(x3)
    M4 = Moment(x4)
    
    'Return the minimum moment
    MinMoment = WorksheetFunction.Min(M1, M2, M3, M4)
    
End Function

'Returns the maximum deflection in the segment
Public Function MaxDeflection(EI As Double) As Double
    
    'Declare local variables
    Dim i As Integer, x As Double, Delta As Double
    
    'Initialize the maximum deflection to zero
    MaxDeflection = 0
    
    'Step through 100 points on the segment
    For i = 1 To 100
        
        'Calculate the location of the point
        x = (i - 1) * Length / 99
        
        'Calculate the deflection at this point
        Delta = Deflection(x, EI)
        
        'Determine if this is the maximum value so far
        If Delta > MaxDeflection Then
            MaxDeflection = Delta
        End If
        
    Next i
   
End Function

'Returns the minimum deflection in the segment
Public Function MinDeflection(EI As Double) As Double
    
    'Declare local variables
    Dim i As Integer, x As Double, Delta As Double
    
    'Initialize the minimum deflection to zero
    MinDeflection = 0
    
    'Step through 100 points on the segment
    For i = 1 To 100
        
        'Calculate the location of the point
        x = (i - 1) * Length / 99
        
        'Calculate the deflection at this point
        Delta = Deflection(x, EI)
        
        'Determine if this is the maximum value so far
        If Delta < MinDeflection Then
            MinDeflection = Delta
        End If
        
    Next i
   
End Function
