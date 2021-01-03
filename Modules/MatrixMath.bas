Attribute VB_Name = "MatrixMath"

'Note: This module assumes all matrices are indexed starting at 1
Option Base 1

'Returns a transposed matrix
Public Function MTranspose(Matrix1 As Matrix) As Matrix

    'Create the new matrix
    Dim Result As New Matrix
    Call Result.Resize(Matrix1.NumCols, Matrix1.NumRows)
    
    'Transpose 'Matrix1'
    Dim i As Long, j As Long
    For i = 1 To Result.NumRows
        For j = 1 To Result.NumCols
            Call Result.SetValue(i, j, Matrix1.GetValue(j, i))
        Next j
    Next i
    
    'Return the transposed matrix
    Set MTranspose = Result
    
End Function

'Returns the sum of two matrices
Public Function MAdd(Matrix1 As Matrix, Matrix2 As Matrix) As Matrix

    'Create the new matrix
    Set MAdd = New Matrix
    Call MAdd.Resize(Matrix1.NumRows, Matrix1.NumCols, True)
    
    'Sum the two matrices
    Dim i As Long, j As Long
    For i = 1 To Matrix1.NumRows
        For j = 1 To Matrix1.NumCols
            Call MAdd.SetValue(i, j, Matrix1.GetValue(i, j) + Matrix2.GetValue(i, j))
        Next j
    Next i
    
End Function

'Returns the difference of two matrices
Public Function MSubtract(Matrix1 As Matrix, Matrix2 As Matrix) As Matrix

    'Create the new matrix
    Dim Result As New Matrix
    Call Result.Resize(Matrix1.NumRows, Matrix1.NumCols)
    
    'Calculate the difference
    Dim i As Long, j As Long
    For i = 1 To Matrix1.NumRows
        For j = 1 To Matrix1.NumCols
            Call Result.SetValue(i, j, Matrix1.GetValue(i, j) - Matrix2.GetValue(i, j))
        Next j
    Next i
    
    'Return the difference
    Set MSubtract = Result
    
End Function

'Returns the product of two matrices
Public Function MMultiply(Matrix1 As Matrix, Matrix2 As Matrix) As Matrix

On Error GoTo ErrorHandler:

    'Initialize the 'Result' matrix
    Dim Result As New Matrix
    Call Result.Resize(Matrix1.NumRows, Matrix2.NumCols)
    
    'Multiply the matrices
    Dim i As Long, j As Long, k As Long
    For i = 1 To Matrix1.NumRows
        For j = 1 To Matrix2.NumCols
            For k = 1 To Matrix1.NumCols
                Call Result.SetValue(i, j, Result.GetValue(i, j) + Matrix1.GetValue(i, k) * Matrix2.GetValue(k, j))
            Next k
        Next j
    Next i
    
    'Return the product
    Set MMultiply = Result
    
    'Exit the function
    Exit Function

ErrorHandler:

    MsgBox ("Error: Unable to multiply matrices.")
    Stop
    
End Function

'Returns an inverted matrix
'This function augments the identity matrix to the right hand side of a matrix
'and then reduces the total matrix to RREF format to obtain the inverted matrix.
'If the matrix is not invertable, a runtime error will occur.
Public Function MInvert(Matrix1 As Matrix) As Matrix

On Error GoTo ErrorHandler:

    'Create the 'Result' matrix, which will be used for Gauss Elimination
    'By default the 'Resize' method initializes the matrix to a zero matrix
    Dim Result As New Matrix
    Call Result.Resize(Matrix1.NumRows, 2 * Matrix1.NumCols)
    
    'Copy 'Matrix1' into the left hand side of 'Result'
    Dim i As Long, j As Long
    For i = 1 To Matrix1.NumRows
        For j = 1 To Matrix1.NumCols
            Call Result.SetValue(i, j, Matrix1.GetValue(i, j))
        Next j
    Next i
    
    'Turn the right hand side of 'Result' into the identity matrix by adding a diagonal of 1's
    j = Matrix1.NumCols + 1
    For i = 1 To Result.NumRows
        Call Result.SetValue(i, j, 1)
        j = j + 1
    Next i
    
    'Step through each row of the matrix
    Dim Multiplier As Double
    For i = 1 To Result.NumRows
    
        'Find the first nonzero term in the row
        'Precision out to 15 decimal places is checked by this code
        'This is the same level of precision that Excel uses
        j = 1
        While Round(Result.GetValue(i, j), 15) = 0
            j = j + 1
        Wend
        
        'Eliminate all terms above and below the first nonzero term
        Dim A As Long, b As Long
        For A = 1 To Result.NumRows
        
            If A <> i Then
            
                If Result.GetValue(A, j) <> 0 Then
                
                    Multiplier = -Result.GetValue(i, j) / Result.GetValue(A, j)
                    For b = 1 To Result.NumCols
                    
                        'This next "if" statement is used to eliminate precision errors by forcing an exact zero value.
                        'Mathematically the "Else" portion of this statement should do that, but it leaves a tiny
                        'precision error which will trigger the "If Result.GetValue(a, j) <> 0" statement above to equal 'True'.
                        If b = j Then
                            Call Result.SetValue(A, b, 0)
                        Else
                            Call Result.SetValue(A, b, Multiplier * Result.GetValue(A, b) + Result.GetValue(i, b))
                        End If
                        
                    Next b
                    
                End If
                
            End If
            
        Next A
        
    Next i
    
    'Put the leading terms in the upper left part of the matrix, going row by row
    'This loop moves diagonally down the matrix, so 'k' represents both a row and column at the same time
    Dim k As Long
    For k = 1 To Result.NumRows
        
        'Find the row with the nonzero term in column k
        i = k
        While Result.GetValue(i, k) = 0 And i <= Result.NumRows
            i = i + 1
        Wend
        
        'If there was not a leading nonzero value in any of the rows, the matrix is uninvertable
        If i > Result.NumRows Then
            MsgBox ("Matrix is uninvertable")
            Stop
        End If
        
        'Swap row 'i' with row 'k'
        Call Result.SwapRows(i, k)
        
    Next k
    
    'Divide each term in the matrix by the leading term to reduce it
    For k = 1 To Result.NumRows
        Multiplier = 1 / Result.GetValue(k, k)
        For j = k To Result.NumCols
            Call Result.SetValue(k, j, Multiplier * Result.GetValue(k, j))
        Next j
    Next k
    
    'Store the right hand side of the matrix in the left hand size of the matrix
    For i = 1 To Result.NumRows
        For j = 1 To Matrix1.NumCols
            Call Result.SetValue(i, j, Result.GetValue(i, Matrix1.NumCols + j))
        Next j
    Next i
    
    'Remove the right hand side from the matrix
    Call Result.Resize(Matrix1.NumRows, Matrix1.NumCols, True)
    
    'Return the remaining matrix which is the inverse
    Set MInvert = Result
    
    Exit Function

ErrorHandler:

    MsgBox ("Error: Unable to invert matrix.")
    Stop
    
End Function

'Multiplies a matrix by a scalar
Public Function MScalarMult(Matrix1 As Matrix, Scalar As Double) As Matrix
    
    Set MScalarMult = New Matrix
    Call MScalarMult.Resize(Matrix1.NumRows, Matrix1.NumCols)
    
    'Multiply 'Matrix1' by 'Scalar'
    Dim i As Long, j As Long
    For i = 1 To MScalarMult.NumRows
        For j = 1 To MScalarMult.NumCols
            Call MScalarMult.SetValue(i, j, Matrix1.GetValue(i, j) * Scalar)
        Next j
    Next i
    
End Function
