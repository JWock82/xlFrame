Attribute VB_Name = "StaticCondensation"
Option Explicit

'Performs static condensation on a stiffness matrix and returns the condensed (expanded) matrix
Public Function k_Condense(StiffMatrix As Matrix, DOF() As Boolean) As Matrix

    'Count the number of DOF's to be condensed out of the matrix
    Dim NumDOF As Long, i As Long
    NumDOF = 0
    For i = 1 To 6
        If DOF(i) = True Then
            NumDOF = NumDOF + 1
        End If
    Next i
    
    'If no degrees of freedom are to be condensed out return the orginal matrix
    If NumDOF = 0 Then
        Set k_Condense = StiffMatrix
        Exit Function
    End If
    
    'Size each sub-matrix
    Dim k11 As New Matrix, k12 As New Matrix, k21 As New Matrix, k22 As New Matrix
    Call k11.Resize(StiffMatrix.NumRows - NumDOF, StiffMatrix.NumCols - NumDOF)
    Call k12.Resize(StiffMatrix.NumRows - NumDOF, NumDOF)
    Call k21.Resize(NumDOF, StiffMatrix.NumCols - NumDOF)
    Call k22.Resize(NumDOF, NumDOF)
    
    'Initialize row/column index variables for each sub-matrix
    Dim i11 As Long, i12 As Long, i21 As Long, i22 As Long
    Dim j11 As Long, j12 As Long, j21 As Long, j22 As Long
    i11 = 1
    i12 = 1
    i21 = 1
    i22 = 1
    j11 = 1
    j12 = 1
    j21 = 1
    j22 = 1
    
    'Partition the matrix into four sub-matrices
    Dim j As Long
    For i = 1 To StiffMatrix.NumRows
    
        For j = 1 To StiffMatrix.NumCols
        
            'Determine which sub-matrix term (i, j) belongs to
            If DOF(i) = True Then
                If DOF(j) = True Then
                    'Place the term in the "k22" sub-matrix
                    Call k22.SetValue(i22, j22, StiffMatrix.GetValue(i, j))
                    j22 = j22 + 1
                    If j22 > NumDOF Then
                        j22 = 1
                        i22 = i22 + 1
                    End If
                Else
                    'Place the term in the "k21" sub-matrix
                    Call k21.SetValue(i21, j21, StiffMatrix.GetValue(i, j))
                    j21 = j21 + 1
                    If j21 > 6 - NumDOF Then
                        j21 = 1
                        i21 = i21 + 1
                    End If
                End If
            Else
                If DOF(j) = True Then
                    'Place the term in the "k12" sub-matrix
                    Call k12.SetValue(i12, j12, StiffMatrix.GetValue(i, j))
                    j12 = j12 + 1
                    If j12 > NumDOF Then
                        j12 = 1
                        i12 = i12 + 1
                    End If
                Else
                    'Place the term in the "k11" sub-matrix
                    Call k11.SetValue(i11, j11, StiffMatrix.GetValue(i, j))
                    j11 = j11 + 1
                    If j11 > 6 - NumDOF Then
                        j11 = 1
                        i11 = i11 + 1
                    End If
                End If
            End If
            
        Next j
        
    Next i
    
    'Calculate the condensed matrix
    Dim kc As New Matrix, M1 As New Matrix, M2 As New Matrix, M3 As New Matrix
    Set M1 = MInvert(k22)
    Set M2 = MMultiply(M1, k21)
    Set M3 = MMultiply(k12, M2)
    Set kc = MSubtract(k11, M3)
    
    'Expand the condensed matrix
    For i = 1 To StiffMatrix.NumRows
        If DOF(i) = True Then
            Call kc.InsertRow(i)
            Call kc.InsertCol(i)
        End If
    Next i
    
    'Return the expanded condensed matrix
    Set k_Condense = kc
    
End Function

'Performs static condensation on a load vector
Public Function CondenseFER(FERVector As Matrix, StiffMatrix As Matrix, DOF() As Boolean) As Matrix
    
    'Count the number of DOF's to be condensed out of the matrix
    Dim NumDOF As Long, i As Long
    NumDOF = 0
    For i = 1 To FERVector.NumRows
        If DOF(i) = True Then
            NumDOF = NumDOF + 1
        End If
    Next i
    
    'Proceed only if the matrix needs to be condensed
    If NumDOF > 0 Then
    
        'Determine the size of "StiffMatrix"
        Dim NumRows As Long, NumCols As Long
        NumRows = StiffMatrix.NumRows
        NumCols = StiffMatrix.NumCols
        
        'Partition the stiffness matrix
        'Only two of the four submatrices, [k12] and [k22], are required to condense the load vector
        'The other two will not be calculated
        Dim k12 As New Matrix, k22 As New Matrix
        Call k12.Resize(NumRows - NumDOF, NumDOF)
        Call k22.Resize(NumDOF, NumDOF)
        
        'Initialize row/column index variables for each sub-matrix
        Dim i12 As Long, i22 As Long, j12 As Long, j22 As Long
        i12 = 1
        i22 = 1
        j12 = 1
        j22 = 1
        
        'Partition the matrix into sub-matrices
        Dim j As Long
        For i = 1 To NumRows
        
            For j = 1 To NumCols
            
                'Determine which sub-matrix term (i, j) belongs to
                If DOF(i) = True Then
                    If DOF(j) = True Then
                        'Place the term in the "k22" sub-matrix
                        Call k22.SetValue(i22, j22, StiffMatrix.GetValue(i, j))
                        j22 = j22 + 1
                        If j22 > NumDOF Then
                            j22 = 1
                            i22 = i22 + 1
                        End If
                    End If
                Else
                    If DOF(j) = True Then
                        'Place the term in the "k12" sub-matrix
                        Call k12.SetValue(i12, j12, StiffMatrix.GetValue(i, j))
                        j12 = j12 + 1
                        If j12 > NumDOF Then
                            j12 = 1
                            i12 = i12 + 1
                        End If
                    End If
                End If
                
            Next j
            
        Next i
        
        'Size each sub-vector
        Dim f1 As New Matrix, f2 As New Matrix
        Call f1.Resize(NumRows - NumDOF, 1)
        Call f2.Resize(NumDOF, 1)
        
        'Initialize row/column index variables for each sub-matrix
        Dim i1 As Long, i2 As Long
        i1 = 1
        i2 = 1
        
        'Partition the load vector into two sub-vectors
        For i = 1 To NumRows
            If DOF(i) = True Then
                Call f2.SetValue(i2, 1, FERVector.GetValue(i, 1))
                i2 = i2 + 1
            Else
                Call f1.SetValue(i1, 1, FERVector.GetValue(i, 1))
                i1 = i1 + 1
            End If
        Next i
        
        'Bring the fixed end reactions to the left side of the equation
        Set f1 = MScalarMult(f1, -1)
        Set f2 = MScalarMult(f2, -1)
        
        'Calculate the condensed vector
        Dim M1 As New Matrix, M2 As New Matrix, M3 As New Matrix
        Set CondenseFER = New Matrix
        Call CondenseFER.Resize(NumRows - NumDOF, 1)
        Set M1 = MInvert(k22)
        Set M2 = MMultiply(M1, f2)
        Set M3 = MMultiply(k12, M2)
        Set CondenseFER = MSubtract(f1, M3)
        
        'Bring the condensed forces back to the right side of the equation
        Set CondenseFER = MScalarMult(CondenseFER, -1)
        
        'Expand the condensed vector
        For i = 1 To NumRows
            If DOF(i) = True Then
                Call CondenseFER.InsertRow(i)
            End If
        Next i
    
    Else
        
        Set CondenseFER = FERVector
        
    End If
    
    
End Function
