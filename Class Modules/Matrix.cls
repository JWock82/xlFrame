VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Matrix() As Double

'Resizes the matrix
Public Sub Resize(M As Long, n As Long, Optional SaveValues As Boolean = True)
    
    'Save the old matrix if necessary
    'Trying to save values in an empty matrix will generate an error
    'Stop the error from occuring
    Dim Temp() As Double
    If SaveValues = True And NumRows > 0 And NumCols > 0 Then
        Temp = m_Matrix
    Else
        SaveValues = False
    End If
    
    'Resize the matrix
    ReDim m_Matrix(1 To M, 1 To n) As Double
    
    'Initialize the matrix
    Dim i As Long, j As Long
    For i = 1 To UBound(m_Matrix, 1)
        For j = 1 To UBound(m_Matrix, 2)
            m_Matrix(i, j) = 0
        Next j
    Next i
    
    'Add values from the old matrix if necessary
    If SaveValues = True Then
    
        Dim a As Long, b As Long
        a = WorksheetFunction.Min(UBound(Temp, 1), M)   'The 'Min' function is necessary if the matrix is getting smaller
        b = WorksheetFunction.Min(UBound(Temp, 2), n)
        
        For i = 1 To a
            For j = 1 To b
                m_Matrix(i, j) = Temp(i, j)
            Next j
        Next i
        
    End If
    
End Sub

'Returns the number of rows in the matrix
Public Function NumRows() As Long
    
On Error GoTo ErrorHandler:
    
    'Calculate and return the number of rows
    NumRows = UBound(m_Matrix, 1) - LBound(m_Matrix, 2) + 1
    Exit Function
        
ErrorHandler:
    
    'If an error occured it's because there are no rows defined yet
    NumRows = 0
    
End Function

'Returns the number of columns in the matrix
Public Function NumCols() As Long

On Error GoTo ErrorHandler:
    
    'Calculate and return the number of columns
    NumCols = UBound(m_Matrix, 2) - LBound(m_Matrix, 2) + 1
    Exit Function
    
ErrorHandler:

    'If an error occured it's because there are no columns defined yet
    NumCols = 0

End Function

'Sets a value at a specified index
Public Sub SetValue(i As Long, j As Long, Value As Double)

    If i <= NumRows And j <= NumCols Then
        m_Matrix(i, j) = Value
    Else
        MsgBox ("Error: Specified matrix index does not exist")
        Stop
    End If
    
End Sub

'Returns the value at a specified index
Public Function GetValue(i As Long, j As Long) As Double

    If i <= NumRows And j <= NumCols Then
        GetValue = m_Matrix(i, j)
    Else
        MsgBox ("Error: Specified matrix index does not exist")
        Stop
    End If
    
End Function

'Inserts a new row of zeros at index 'i'
Public Sub InsertRow(i As Long)
    
On Error GoTo ErrorHandler:
    
        'Resize the matrix to hold one more row
        Call Resize(NumRows + 1, NumCols, True)
    
        'Shift everything at or below row 'i' down one row
        Dim a As Long, j As Long
        a = NumRows
        While a > i
        
            For j = 1 To NumCols
                m_Matrix(a, j) = m_Matrix(a - 1, j)
            Next j
            
            a = a - 1
            
        Wend
        
        'Initialize the new row 'i' to zero
        For j = 1 To NumCols
            m_Matrix(i, j) = 0
        Next j
        
        Exit Sub
        
ErrorHandler:
        
        'The user specified a bad index
        MsgBox ("Error: Cannot insert new row into matrix at specified index.")
        Stop
        
End Sub

'Inserts a new column of zeros at index 'i'
Public Sub InsertCol(j As Long)

On Error GoTo ErrorHandler:

    'Resize the matrix to hold one more column
    Call Resize(NumRows, NumCols + 1, True)
    
    'Shift everything at or to the right of column 'j' to the right one column
    Dim b As Long, i As Long
    b = NumCols
    While b > j
        
        For i = 1 To NumRows
            m_Matrix(i, b) = m_Matrix(i, b - 1)
        Next i
            
        b = b - 1
            
    Wend
        
    'Initialize the new column 'j' to zero
    For i = 1 To NumRows
        m_Matrix(i, j) = 0
    Next i
        
    Exit Sub

ErrorHandler:

    'The user specified a bad index
    MsgBox ("Error: Cannot insert new column into matrix at specified index.")
    Stop

End Sub

'Removes a row from the matrix
Public Sub RemoveRow(RowNum As Long)
    
    'Make sure a valid index has been specified
    If RowNum <= NumRows And RowNum > 0 Then
    
        'Remove the row from the matrix
        Dim i As Long, j As Long
        For i = RowNum To NumRows - 1
            For j = 1 To NumCols
                m_Matrix(i, j) = m_Matrix(i + 1, j)
            Next j
        Next i
    
        Call Resize(NumRows - 1, NumCols, True)
    
    Else
        
        MsgBox ("Error: No row to remove at specified index.")
        Stop
        
    End If
    
End Sub

'Removes a column from the matrix
Public Sub RemoveCol(ColNum As Long)

    'Make sure a valid index has been specified
    If ColNum <= NumCols And ColNum > 0 Then

        'Remove the column from the matrix
        Dim i As Long, j As Long
        For j = ColNum To NumCols - 1
            For i = 1 To NumRows
                m_Matrix(i, j) = m_Matrix(i, j + 1)
            Next i
        Next j
    
        Call Resize(NumRows, NumCols - 1, True)
    
    Else
        
        MsgBox ("Error: No column to remove at specified index.")
        Stop
        
    End If
    
End Sub

'Swaps 2 rows in the matrix with each other
Public Sub SwapRows(Row1 As Long, Row2 As Long)
    
On Error GoTo ErrorHandler:

    'Step through each column in the matrix
    Dim j As Integer, Value As Double
    For j = 1 To NumCols
    
        'Store the value from 'Row2' in 'Value'
        Value = m_Matrix(Row2, j)
        
        'Copy the value from 'Row1' into 'Row2'
        m_Matrix(Row2, j) = m_Matrix(Row1, j)
        
        'Copy 'Value' into 'Row1'
        m_Matrix(Row1, j) = Value
        
    Next j
    
    'Exit the subroutine
    Exit Sub
    
ErrorHandler:
    
    'Send an error to the user if an invalid row was specified
    MsgBox ("Unable to swap matrix rows. Invalid row specified.")
    Stop
    
End Sub

'Prints the matrix to a specified range in a workbook
Public Sub PrintMatrix(PrintRange As Range, Optional ClearExisting As Boolean = False)
    
    'Determine if existing values are to be cleared
    If ClearExisting = True Then
    
        'Identify the range the existing data occupies
        Dim FittedPrintRange As Range
        Set FittedPrintRange = PrintRange.Worksheet.Range(PrintRange.Cells(1, 1).Address, PrintRange.Cells(1, 1).End(xlToRight).End(xlDown).Address)
        
        'Clear the data from the range
        FittedPrintRange.ClearContents
        
    End If
    
    'Count the number of rows and columns in the matrix
    Dim M As Integer, n As Integer
    M = NumRows
    n = NumCols
    
    'Identify the range the matrix will be printed to
    Set FittedPrintRange = PrintRange.Worksheet.Range(PrintRange.Cells(1, 1).Address, PrintRange.Cells(M, n).Address)
    
    'Print the matrix to the range
    FittedPrintRange = m_Matrix
    
End Sub
