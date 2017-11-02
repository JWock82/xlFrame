Attribute VB_Name = "ArrayTools"

'Adds an item to the end of a 1D array
Public Sub AddItem(ByRef ArrayArg() As Variant, ItemToAdd As Variant)
    
    'Resize the array to hold 1 more item
    If CountRows(ArrayArg) = 0 Then
        ReDim ArrayArg(1 To 1)
    Else
        ReDim Preserve ArrayArg(LBound(ArrayArg) To UBound(ArrayArg) + 1)
    End If
    
    'Add the item
    ArrayArg(UBound(ArrayArg)) = ItemToAdd
    
End Sub

'Finds an item's index in an array
Public Function FindIndex(ArrayArg() As Variant, ItemToFind As Variant) As Integer
    
    'Find the first occurance of the item
    Dim i As Integer
    For i = LBound(ArrayArg) To UBound(ArrayArg)
        
        'Compare the 'ith' element to 'ItemToFind'
        If ArrayArg(i) = ItemToFind Then
        
            'Return the index and exit the function
            FindIndex = i
            Exit Function
            
        End If
        
    Next i
    
    'If code execution has reached this point the item is not in the array
    ItemToFind = -1
    
End Function

'Sorts a 1D array by ascending values
'Bubble sorting algorithm (inefficient for long lists, but simple to program)
Public Sub SortAscending(ByRef ArrayArg() As Variant)
    
    'Declare local variables
    Dim i As Integer
    Dim Temp As Variant
    Dim Sorted As Boolean
    
    'Assume the array is not sorted to begin with
    Sorted = False
    
    'Sort the array in ascending order
    While Sorted = False
    
        Sorted = True
        For i = LBound(ArrayArg) To UBound(ArrayArg) - 1
        
            If ArrayArg(i) > ArrayArg(i + 1) Then
                Temp = ArrayArg(i)
                ArrayArg(i) = ArrayArg(i + 1)
                ArrayArg(i + 1) = Temp
                Sorted = False
            End If
            
        Next i
        
    Wend
    
End Sub

'Sorts a 1D array by descending values
'Bubble sorting algorithm (inefficient for long lists, but simple to program)
Public Sub SortDescending(ByRef ArrayArg() As Variant)
    
    'Declare local variables
    Dim i As Integer
    Dim Temp As Variant
    Dim Sorted As Boolean
    
    'Assume the array is not sorted to begin with
    Sorted = False
    
    'Sort the array in descending order
    While Sorted = False
    
        Sorted = True
        For i = LBound(ArrayArg) To UBound(ArrayArg) - 1
        
            If ArrayArg(i) < ArrayArg(i + 1) Then
                Temp = ArrayArg(i)
                ArrayArg(i) = ArrayArg(i + 1)
                ArrayArg(i + 1) = Temp
                Sorted = False
            End If
            
        Next i
        
    Wend
    
End Sub

'Sorts a 2D array by ascending values in a given column
'Bubble sorting algorithm (inefficient for long lists, but simple to program)
Public Sub SortAscendingByColumn(ByRef ArrayArg() As Variant, Optional ByVal SortIndex As Integer = -1)
    
    Dim Sorted As Double
    Sorted = False
    
    'Default to the first column if the user has not specified which column to use
    If SortIndex = -1 Then
        SortIndex = LBound(ArrayArg, 2)
    End If
    
    Dim i As Integer, Temp As Variant
    While Sorted = False
        
        Sorted = True
        
        For i = LBound(ArrayArg, 1) To UBound(ArrayArg, 1) - 1
            
            If ArrayArg(i, SortIndex) > ArrayArg(i + 1, SortIndex) Then
                
                Temp = ArrayArg(i, SortIndex)
                ArrayArg(i, SortIndex) = ArrayArg(i + 1, SortIndex)
                ArrayArg(i + 1, SortIndex) = Temp
                Sorted = False
                
            End If
            
        Next i
        
    Wend
    
End Sub

'Removes duplicate values from a 1D array
'Note: The comparisons used in this subroutine do not account for precision errors
Public Sub RemoveDuplicates(ByRef ArrayArg() As Variant)
    
    'Step through each element in the array
    Dim i As Integer, j As Integer
    
    'Note: VBA 'For' loops do not re-evaluate the term after 'To' upon each iteration
    'A 'While' loop has been used instead to force re-evaluation
    i = LBound(ArrayArg)
    While i < UBound(ArrayArg)
        j = i + 1
        While j <= UBound(ArrayArg)
            If ArrayArg(i) = ArrayArg(j) Then
                Call RemoveValue(ArrayArg, j)
                j = j - 1
            End If
            j = j + 1
        Wend
        i = i + 1
    Wend
    
End Sub

'Removes empty values from an array
Public Sub RemoveEmpty(ByRef ArrayArg() As Variant)

    'Step through each element in the array
    Dim i As Integer, j As Integer

    'Note: VBA 'For' loops do not re-evaluate the term after 'To' upon each iteration
    'A 'While' loop has been used instead to force re-evaluation
    i = LBound(ArrayArg)
    While i <= UBound(ArrayArg)
        If ArrayArg(i) = Empty Then
            Call RemoveValue(ArrayArg, i)
            i = i - 1
        End If
        i = i + 1
    Wend
    
End Sub

'Removes a value from an array
Public Sub RemoveValue(ByRef ArrayArg() As Variant, Index As Integer)
    
    'Step through each element in the list after 'Index'
    Dim i As Integer
    For i = Index To UBound(ArrayArg) - 1
        ArrayArg(i) = ArrayArg(i + 1)
    Next i
    
    'Redimension the list
    ReDim Preserve ArrayArg(LBound(ArrayArg) To UBound(ArrayArg) - 1) As Variant
    
End Sub

'Counts the number of rows in an array
Public Function CountRows(List() As Variant) As Integer
    
On Error GoTo 1:
    CountRows = UBound(List(), 1) - LBound(List(), 1) + 1
    Exit Function
1:
    CountRows = 0

End Function

'Counts the number of columns in an array
Public Function CountColumns(List() As Variant) As Integer

On Error GoTo 1:
    CountColumns = UBound(List(), 2) - LBound(List(), 2) + 1
    Exit Function
1:
    CountColumns = 0
    
End Function

'Prints an array to a specified range in a workbook
Public Sub PrintArray(ArrayToPrint() As Variant, PrintRange As Range)
    
    'Count the number of rows and columns in the array
    Dim M As Integer, n As Integer
    M = CountRows(ArrayToPrint)
    n = CountColumns(ArrayToPrint)
    
    'Identify the range the array will be printed to
    Dim FittedPrintRange As Range
    Set FittedPrintRange = PrintRange.Worksheet.Range(PrintRange.Cells(1, 1).Address, PrintRange.Cells(M, n).Address)
    
    'Print the array to the range
    FittedPrintRange = ArrayToPrint
    
End Sub

'Clears an array that has previously been printed to a specified range in a workbook
Public Sub ClearArray(PrintRange As Range)

    'Identify the range the array occupies
    Dim FittedPrintRange As Range
    Set FittedPrintRange = PrintRange.Worksheet.Range(PrintRange.Cells(1, 1).Address, PrintRange.Cells(1, 1).End(xlToRight).End(xlDown).Address)
    
    'Clear the data from the range
    FittedPrintRange.ClearContents

End Sub

'Used for testing/debugging
Private Sub test()

    Dim myArray(1 To 3, 1 To 4) As Variant
    
    myArray(1, 1) = "A"
    myArray(1, 2) = "B"
    myArray(1, 3) = "C"
    myArray(1, 4) = "D"
    myArray(2, 1) = "E"
    myArray(2, 2) = "F"
    myArray(2, 3) = "G"
    myArray(2, 4) = "H"
    myArray(3, 1) = "I"
    myArray(3, 2) = "J"
    myArray(3, 3) = "K"
    myArray(3, 4) = "L"
    
    Call PrintArray(myArray, Sheet1.Range("W21"))
    
    Call ClearArray(Sheet1.Range("W21"))

End Sub
