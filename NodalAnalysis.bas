Attribute VB_Name = "Module1"
Option Explicit

Sub NodalAnalysis()
    Dim G(1 To 3, 1 To 3) As Double
    Dim I(1 To 3) As Double
    Dim V() As Double
    
    ' Example values for resistances (ohms)
    Dim R12 As Double: R12 = 2
    Dim R13 As Double: R13 = 3
    Dim R23 As Double: R23 = 4
    
    ' Conductance matrix G (1/R)
    G(1, 1) = 1 / R12 + 1 / R13
    G(1, 2) = -1 / R12
    G(1, 3) = -1 / R13
    
    G(2, 1) = -1 / R12
    G(2, 2) = 1 / R12 + 1 / R23
    G(2, 3) = -1 / R23
    
    G(3, 1) = -1 / R13
    G(3, 2) = -1 / R23
    G(3, 3) = 1 / R13 + 1 / R23
    
    ' Current source vector I
    ' Assuming no current sources (I = 0 for all nodes)
    I(1) = 0.001
    I(2) = 0
    I(3) = 0
    
    ' Solve G * V = I
    V = SolveLinearSystem(G, I)
    
    ' Output the node voltages to Sheet1
    Dim x As Integer
    For x = LBound(V) To UBound(V)
        Worksheets("Sheet1").Cells(x, 1).Value = "V" & x & " = " & V(x)
    Next x
End Sub

Function SolveLinearSystem(G As Variant, I As Variant) As Variant
    ' Solve the linear system G * V = I using Gaussian elimination
    Dim n As Integer
    n = UBound(G, 1)
    
    Dim A() As Double
    Dim B() As Double
    ReDim A(1 To n, 1 To n)
    ReDim B(1 To n)
    
    Dim row As Integer, col As Integer
    For row = 1 To n
        For col = 1 To n
            A(row, col) = G(row, col)
        Next col
        B(row) = I(row)
    Next row
    
    ' Perform Gaussian elimination
    Dim k As Integer, m As Double
    For k = 1 To n - 1
        For row = k + 1 To n
            m = A(row, k) / A(k, k)
            For col = k + 1 To n
                A(row, col) = A(row, col) - m * A(k, col)
            Next col
            B(row) = B(row) - m * B(k)
        Next row
    Next k
    
    ' Back substitution
    Dim V() As Double
    ReDim V(1 To n)
    V(n) = B(n) / A(n, n)
    For row = n - 1 To 1 Step -1
        V(row) = B(row)
        For col = row + 1 To n
            V(row) = V(row) - A(row, col) * V(col)
        Next col
        V(row) = V(row) / A(row, row)
    Next row
    
    SolveLinearSystem = V
End Function

