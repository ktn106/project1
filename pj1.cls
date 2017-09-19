VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "YounisBBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 1

Function addmatrix(a, b)
Dim av() As Variant, bv() As Variant
av = a.Value
bv = b.Value

Dim a1 As Double, a2 As Double, b1 As Double, b2 As Double
a1 = -LBound(av, 1) + UBound(av, 1) + 1
a2 = -LBound(av, 2) + UBound(av, 2) + 1
b1 = -LBound(bv, 1) + UBound(bv, 1) + 1
b2 = -LBound(bv, 2) + UBound(bv, 2) + 1

If a1 <> b1 Or a2 <> b2 Then
    addmatrix = "Matrixs must have same dimensions"
Else

Dim c() As Double, i As Double, j As Double
ReDim c(1 To a1, 1 To a2)
For i = 1 To a1
For j = 1 To a2
    c(i, j) = a(i, j) + b(i, j)
    Next j
        Next i

addmatrix = c
   End If
    
End Function


Function minusmatrix(a, b)
Dim av() As Variant, bv() As Variant
av = a.Value
bv = b.Value

Dim a1 As Double, a2 As Double, b1 As Double, b2 As Double
a1 = -LBound(av, 1) + UBound(av, 1) + 1
a2 = -LBound(av, 2) + UBound(av, 2) + 1
b1 = -LBound(bv, 1) + UBound(bv, 1) + 1
b2 = -LBound(bv, 2) + UBound(bv, 2) + 1

If a1 <> b1 Or a2 <> b2 Then
minusmatrix = "Matrixs must have same dimensions"
Else
Dim c() As Double, i As Double, j As Double
ReDim c(1 To a1, 1 To a2)
For i = 1 To a1
For j = 1 To a2
    c(i, j) = a(i, j) - b(i, j)
    Next j
    Next i
minusmatrix = c
   End If
    
End Function