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
'Checked 1/10/2017
Private m_SegStart As Double        'Segment start location relative to the start of the beam
Private m_SegEnd As Double          'Segment end location relative to the start of the beam
Private m_w1 As Double              'Distributed load at the start of the segment
Private m_w2 As Double              'Distributed load at the end of the segement
Private m_w1_Axial As Double        'Distributed axial load at the start of the segment
Private m_w2_Axial As Double        'Distributed axial load at the end of the segement
Private m_P1 As Double              'Axial load at the start of the segment
Private m_V1 As Double              'Shear at the start of the segment
Private m_M1 As Double              'Moment at the start of the segment
Private m_theta1 As Double          'Rotation at the start of the segment
Private m_Delta1 As Double          'Deflection at the start of the segment

'Class properties
'Checked 1/10/2017
Property Let SegStart(Value As Double)
    m_SegStart = Value
End Property

Property Get SegStart() As Double
    SegStart = m_SegStart
End Property

Property Let SegEnd(Value As Double)
    m_SegEnd = Value
End Property

Property Get SegEnd() As Double
    SegEnd = m_SegEnd
End Property

Property Let w1(Value As Double)
    m_w1 = Value
End Property

Property Get w1() As Double
    w1 = m_w1
End Property

Property Let w2(Value As Double)
    m_w2 = Value
End Property

Property Get w2() As Double
    w2 = m_w2
End Property

Property Let w1_Axial(Value As Double)
    m_w1_Axial = Value
End Property

Property Get w1_Axial() As Double
    w1_Axial = m_w1_Axial
End Property

Property Let w2_Axial(Value As Double)
    m_w2_Axial = Value
End Property

Property Get w2_Axial() As Double
    w2_Axial = m_w2_Axial
End Property

Property Let p1(Value As Double)
    m_P1 = Value
End Property

Property Get p1() As Double
    p1 = m_P1
End Property

Property Let V1(Value As Double)
    m_V1 = Value
End Property

Property Get V1() As Double
    V1 = m_V1
End Property

Property Let M1(Value As Double)
    m_M1 = Value
End Property

Property Get M1() As Double
    M1 = m_M1
End Property

Property Let theta1(Value As Double)
    m_theta1 = Value
End Property

Property Let Delta1(Value As Double)
    m_Delta1 = Value
End Property

'Returns the length of the segment
'Checked 3/6/2017
Public Function Length() As Double
    Length = m_SegEnd - m_SegStart
End Function

'Returns the axial force at a location on the segment
'Checked 1/10/2017
Public Function Axial(x As Double) As Double
    
    'Return the axial force at the location
    Axial = m_P1 + (m_w2_Axial - m_w1_Axial) / (2 * Length) * x ^ 2 + m_w1_Axial * x
    
End Function

'Returns the shear force at a location on the segment
'Checked 1/10/2017
Public Function Shear(x As Double) As Double
    
    'Return the shear at the location
    Shear = m_V1 - (m_w2 - m_w1) / (2 * Length) * x ^ 2 - m_w1 * x
    
End Function

'Returns the moment at a location on the segment
'Checked 1/10/2017
Public Function Moment(x As Double) As Double
    
    'Return the moment at the location
    Moment = m_M1 + m_V1 * x - (m_w2 - m_w1) / (6 * Length) * x ^ 3 - m_w1 * x ^ 2 / 2
    
End Function

'Returns the rotation at a location on the segment
'Checked 1/10/2017
Public Function Rotation(x As Double, EI As Double) As Double
    
    'Return the rotation at the location
    Rotation = m_theta1 + 1 / EI * (m_M1 * x + m_V1 / 2 * x ^ 2 - (m_w2 - m_w1) / (24 * Length) * x ^ 4 - m_w1 / 6 * x ^ 3)
    
End Function

'Returns the deflection at a location on the segment
'Checked 1/10/2017
Public Function Deflection(x As Double, EI As Double) As Double

    'Return the deflection at the location
    Deflection = m_Delta1 + m_theta1 * x + 1 / EI * (m_M1 / 2 * x ^ 2 + m_V1 / 6 * x ^ 3 - (m_w2 - m_w1) / (120 * Length) * x ^ 5 - m_w1 / 24 * x ^ 4)
    
End Function

'Returns the maximum axial force in the segment
'Checked 3/6/2017
Public Function MaxAxial() As Double
    
    'Declare local variables
    Dim x1 As Double, x2 As Double, x3 As Double
    Dim p1 As Double, p2 As Double, P3 As Double
    
    'Determine possible locations of maximum axial force
    If m_w1_Axial - m_w2_Axial <> 0 Then
        x1 = Length * m_w1_Axial / (m_w1_Axial - m_w2_Axial)
    Else
        x1 = 0
    End If
    
    If x1 < 0 Or x1 > Length Then
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
'Checked 3/6/2017
Public Function MinAxial() As Double
    
    'Declare local variables
    Dim x1 As Double, x2 As Double, x3 As Double
    Dim p1 As Double, p2 As Double, P3 As Double
    
    'Determine possible locations of minimum axial force
    If m_w1_Axial - m_w2_Axial <> 0 Then
        x1 = Length * m_w1_Axial / (m_w1_Axial - m_w2_Axial)
    Else
        x1 = 0
    End If
    
    If x1 < 0 Or x1 > Length Then
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
'Checked 2/12/2017
Public Function MaxShear() As Double
    
    'Declare local variables
    Dim x1 As Double, x2 As Double, x3 As Double
    Dim V1 As Double, V2 As Double, V3 As Double
    
    'Determine possible locations of maximum shear
    If m_w1 - m_w2 = 0 Then
        x1 = 0
    Else
        x1 = m_w1 * Length / (m_w1 - m_w2)
    End If
    
    If x1 < 0 Or x1 > Length Then
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
'Checked 2/12/2017
Public Function MinShear() As Double
    
    'Declare local variables
    Dim x1 As Double, x2 As Double, x3 As Double
    Dim V1 As Double, V2 As Double, V3 As Double
    
    'Determine possible locations of minimum shear
    If m_w1 - m_w2 = 0 Then
        x1 = 0
    Else
        x1 = m_w1 * Length / (m_w1 - m_w2)
    End If
    
    If x1 < 0 Or x1 > Length Then
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
'Checked 2/12/2017
Public Function MaxMoment() As Double

    'Declare local variables
    Dim a As Double, b As Double, c As Double
    Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double
    Dim M1 As Double, M2 As Double, M3 As Double, M4 As Double
    
    'Find the quadratic equation parameters
    a = (m_w1 - m_w2) / (2 * Length)
    b = -m_w1
    c = m_V1
    
    'Determine possible locations of maximum moment
    If a = 0 Then
        If b <> 0 Then
            x1 = -c / b
        Else
            x1 = 0
        End If
        x2 = 0
    ElseIf b ^ 2 - 4 * a * c < 0 Then
        x1 = 0
        x2 = 0
    Else
        x1 = (-b + (b ^ 2 - 4 * a * c) ^ 0.5) / (2 * a)
        x2 = (-b - (b ^ 2 - 4 * a * c) ^ 0.5) / (2 * a)
    End If
    
    x3 = 0
    x4 = Length
    
    If x1 < 0 Or x1 > Length Then
        x1 = 0
    End If
    
    If x2 < 0 Or x2 > Length Then
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
'Checked 2/12/2017
Public Function MinMoment() As Double

    'Declare local variables
    Dim a As Double, b As Double, c As Double
    Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double
    Dim M1 As Double, M2 As Double, M3 As Double, M4 As Double
    
    'Find the quadratic equation parameters
    a = (m_w1 - m_w2) / (2 * Length)
    b = -m_w1
    c = m_V1
    
    'Determine possible locations of minimum moment
    If a = 0 Then
        If b <> 0 Then
            x1 = -c / b
        Else
            x1 = 0
        End If
        x2 = 0
    ElseIf b ^ 2 - 4 * a * c < 0 Then
        x1 = 0
        x2 = 0
    Else
        x1 = (-b + (b ^ 2 - 4 * a * c) ^ 0.5) / (2 * a)
        x2 = (-b - (b ^ 2 - 4 * a * c) ^ 0.5) / (2 * a)
    End If
    
    x3 = 0
    x4 = Length
    
    If x1 < 0 Or x1 > Length Then
        x1 = 0
    End If
    
    If x2 < 0 Or x2 > Length Then
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
