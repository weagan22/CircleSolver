Public Class MainForm
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim points(,) As Double = {{6, 8}, {8, 4}, {3, 9}}

        Dim matrixOps As MatrixOps = New MatrixOps

        Dim aMatrix(,) As Double
        ReDim aMatrix(2, 2)


        For r1 = 0 To UBound(aMatrix, 1)
            For c1 = 0 To UBound(aMatrix, 2)
                If c1 = 2 Then
                    aMatrix(r1, c1) = -1
                Else
                    aMatrix(r1, c1) = points(r1, c1)
                End If
            Next
        Next

        Dim bMatrix() As Double
        ReDim bMatrix(2)
        For r1 = 0 To UBound(bMatrix)
            bMatrix(r1) = -(points(r1, 0) ^ 2 + points(r1, 1) ^ 2)
        Next



        Dim invMatrix(,) As Double = MatrixOps.matrix_Inverse(aMatrix)
        Dim solMatrix() As Double = MatrixOps.matrixMultSingle(invMatrix, bMatrix)


        MsgBox(MatrixOps.printMatrix(points) & vbNewLine & vbNewLine & MatrixOps.printMatrix(aMatrix) & vbNewLine & vbNewLine & MatrixOps.printSingleMatrix(bMatrix) & vbNewLine & vbNewLine & MatrixOps.printMatrix(invMatrix) & vbNewLine & vbNewLine & MatrixOps.printSingleMatrix(solMatrix))

        MsgBox("Center: (" & -solMatrix(0) / 2 & "," & -solMatrix(1) / 2 & ") & Radius: " & Math.Sqrt((solMatrix(0) / 2) ^ 2 + (solMatrix(1) / 2) ^ 2 + solMatrix(2)))

    End Sub
End Class
