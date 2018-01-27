Attribute VB_Name = "FERFunctions"

'Returns the fixed end reaction vector for a point load
Public Function FER_PtLoad(P As Double, x As Double, L As Double) As Matrix
    
    'Define variables
    Dim b As Double
    b = L - x
    
    'Create the fixed end reaction vector
    Set FER_PtLoad = New Matrix
    Call FER_PtLoad.Resize(6, 1)
    
    'Populate the fixed end reaction vector
    Call FER_PtLoad.SetValue(1, 1, 0)
    Call FER_PtLoad.SetValue(2, 1, P * b ^ 2 * (L + 2 * x) / L ^ 3)
    Call FER_PtLoad.SetValue(3, 1, P * x * b ^ 2 / L ^ 2)
    Call FER_PtLoad.SetValue(4, 1, 0)
    Call FER_PtLoad.SetValue(5, 1, P * x ^ 2 * (L + 2 * b) / L ^ 3)
    Call FER_PtLoad.SetValue(6, 1, -P * x ^ 2 * b / L ^ 2)
    
End Function

'Returns the fixed end reaction vector for a moment
Public Function FER_Moment(M As Double, x As Double, L As Double) As Matrix
    
    'Define variables
    Dim b As Double
    b = L - x
    
    'Create the fixed end reaction vector
    Set FER_Moment = New Matrix
    Call FER_Moment.Resize(6, 1)
    
    'Populate the fixed end reaction vector
    Call FER_Moment.SetValue(1, 1, 0)
    Call FER_Moment.SetValue(2, 1, -M * (x ^ 2 + b ^ 2 - 4 * x * b - L ^ 2) / L ^ 3)
    Call FER_Moment.SetValue(3, 1, M * b * (2 * x - b) / L ^ 2)
    Call FER_Moment.SetValue(4, 1, 0)
    Call FER_Moment.SetValue(5, 1, M * (x ^ 2 + b ^ 2 - 4 * x * b - L ^ 2) / L ^ 3)
    Call FER_Moment.SetValue(6, 1, M * x * (2 * b - x) / L ^ 2)
    
End Function

'Returns the fixed end reaction vector for a linear distributed load
Public Function FER_LinLoad(w1 As Double, w2 As Double, x1 As Double, x2 As Double, L As Double) As Matrix
        
    'Create the fixed end reaction vector
    Set FER_LinLoad = New Matrix
    Call FER_LinLoad.Resize(6, 1)
    
    'Populate the fixed end reaction vector
    Call FER_LinLoad.SetValue(1, 1, 0)
    Call FER_LinLoad.SetValue(2, 1, -(x1 - x2) * (10 * L ^ 3 * w1 + 10 * L ^ 3 * w2 - 15 * L * w1 * x1 ^ 2 - 10 * L * w1 * x1 * x2 - 5 * L * w1 * x2 ^ 2 - 5 * L * w2 * x1 ^ 2 - 10 * L * w2 * x1 * x2 - 15 * L * w2 * x2 ^ 2 + 8 * w1 * x1 ^ 3 + 6 * w1 * x1 ^ 2 * x2 + 4 * w1 * x1 * x2 ^ 2 + 2 * w1 * x2 ^ 3 + 2 * w2 * x1 ^ 3 + 4 * w2 * x1 ^ 2 * x2 + 6 * w2 * x1 * x2 ^ 2 + 8 * w2 * x2 ^ 3) / (20 * L ^ 3))
    Call FER_LinLoad.SetValue(3, 1, -(x1 - x2) * (20 * L ^ 2 * w1 * x1 + 10 * L ^ 2 * w1 * x2 + 10 * L ^ 2 * w2 * x1 + 20 * L ^ 2 * w2 * x2 - 30 * L * w1 * x1 ^ 2 - 20 * L * w1 * x1 * x2 - 10 * L * w1 * x2 ^ 2 - 10 * L * w2 * x1 ^ 2 - 20 * L * w2 * x1 * x2 - 30 * L * w2 * x2 ^ 2 + 12 * w1 * x1 ^ 3 + 9 * w1 * x1 ^ 2 * x2 + 6 * w1 * x1 * x2 ^ 2 + 3 * w1 * x2 ^ 3 + 3 * w2 * x1 ^ 3 + 6 * w2 * x1 ^ 2 * x2 + 9 * w2 * x1 * x2 ^ 2 + 12 * w2 * x2 ^ 3) / (60 * L ^ 2))
    Call FER_LinLoad.SetValue(4, 1, 0)
    Call FER_LinLoad.SetValue(5, 1, (x1 - x2) * (-15 * L * w1 * x1 ^ 2 - 10 * L * w1 * x1 * x2 - 5 * L * w1 * x2 ^ 2 - 5 * L * w2 * x1 ^ 2 - 10 * L * w2 * x1 * x2 - 15 * L * w2 * x2 ^ 2 + 8 * w1 * x1 ^ 3 + 6 * w1 * x1 ^ 2 * x2 + 4 * w1 * x1 * x2 ^ 2 + 2 * w1 * x2 ^ 3 + 2 * w2 * x1 ^ 3 + 4 * w2 * x1 ^ 2 * x2 + 6 * w2 * x1 * x2 ^ 2 + 8 * w2 * x2 ^ 3) / (20 * L ^ 3))
    Call FER_LinLoad.SetValue(6, 1, -(x1 - x2) * (-15 * L * w1 * x1 ^ 2 - 10 * L * w1 * x1 * x2 - 5 * L * w1 * x2 ^ 2 - 5 * L * w2 * x1 ^ 2 - 10 * L * w2 * x1 * x2 - 15 * L * w2 * x2 ^ 2 + 12 * w1 * x1 ^ 3 + 9 * w1 * x1 ^ 2 * x2 + 6 * w1 * x1 * x2 ^ 2 + 3 * w1 * x2 ^ 3 + 3 * w2 * x1 ^ 3 + 6 * w2 * x1 ^ 2 * x2 + 9 * w2 * x1 * x2 ^ 2 + 12 * w2 * x2 ^ 3) / (60 * L ^ 2))
    
End Function

'Returns the fixed end reaction vector for an axial point load
Public Function FER_AxialPtLoad(P As Double, x As Double, L As Double) As Matrix
    
    'Create the fixed end reaction vector
    Set FER_AxialPtLoad = New Matrix
    Call FER_AxialPtLoad.Resize(6, 1)
    
    'Populate the fixed end reaction vector
    Call FER_AxialPtLoad.SetValue(1, 1, -P * (L - x) / L)
    Call FER_AxialPtLoad.SetValue(2, 1, 0)
    Call FER_AxialPtLoad.SetValue(3, 1, 0)
    Call FER_AxialPtLoad.SetValue(4, 1, -P * x / L)
    Call FER_AxialPtLoad.SetValue(5, 1, 0)
    Call FER_AxialPtLoad.SetValue(6, 1, 0)
    
End Function

'Returns the fixed end reaction vector for a distributed axial load
Public Function FER_AxialLinLoad(p1 As Double, p2 As Double, x1 As Double, x2 As Double, L As Double) As Matrix
    
    'Create the fixed end reaction vector
    Set FER_AxialLinLoad = New Matrix
    Call FER_AxialLinLoad.Resize(6, 1)
    
    'Populate the fixed end reaction vector
    Call FER_AxialLinLoad.SetValue(1, 1, 1 / (6 * L) * (x1 - x2) * (3 * L * p1 + 3 * L * p2 - 2 * p1 * x1 - p1 * x2 - p2 * x1 - 2 * p2 * x2))
    Call FER_AxialLinLoad.SetValue(2, 1, 0)
    Call FER_AxialLinLoad.SetValue(3, 1, 0)
    Call FER_AxialLinLoad.SetValue(4, 1, 1 / (6 * L) * (x1 - x2) * (2 * p1 * x1 + p1 * x2 + p2 * x1 + 2 * p2 * x2))
    Call FER_AxialLinLoad.SetValue(5, 1, 0)
    Call FER_AxialLinLoad.SetValue(6, 1, 0)
    
End Function
