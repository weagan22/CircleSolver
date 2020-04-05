Public Class MainForm
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Initialize points
        Dim points() As Point

        ReDim points(4)
        points(0) = New Point(2, 4)
        points(1) = New Point(6, 8)
        points(2) = New Point(8, 6)
        points(3) = New Point(2.7977, 7.7039)
        points(4) = New Point(7.4941, 2.1489)


        Dim fitCir = New BestFit
        Try
            fitCir.fitCircle(points)
        Catch ex As Exception
            MsgBox("Failed to fit circle to the points." & vbNewLine & vbNewLine & ex.Message, , "Error")
        End Try

        Dim test1 = avgResidualSqr(points, fitCir.centerPnt.X, fitCir.centerPnt.Y, fitCir.rHat)


    End Sub

    Sub calcCir(points() As Point)
        Dim matrixOps As MatrixOps = New MatrixOps

        'Move points into A matrix
        'x^2 + y^2 + ax + by - 1c = 0
        '             ^    ^   ^
        Dim aMatrix(,) As Double
        ReDim aMatrix(2, 2)

        For r1 = 0 To UBound(aMatrix, 1)
            aMatrix(r1, 0) = points(r1).X
            aMatrix(r1, 1) = points(r1).Y
            aMatrix(r1, 2) = -1
        Next

        'Calculate and add values to B matrix
        '  x^2 + y^2 + ax + by - 1c  = 0
        '-(x^2 + y^2)                  -(x^2 + y^2)   
        '              ax + by - 1c = -(x^2 + y^2)
        Dim bMatrix() As Double
        ReDim bMatrix(2)
        For r1 = 0 To UBound(bMatrix)
            bMatrix(r1) = -(points(r1).X ^ 2 + points(r1).Y ^ 2)
        Next


        'Calculate invA
        Dim invAmatrix(,) As Double = MatrixOps.matrix_Inverse(aMatrix)

        'Calcualte solution matrix AX=b => X=b*invA
        Dim solMatrix() As Double = MatrixOps.matrixMultSingle(invAmatrix, bMatrix)


        'Convert from general form to standard form x^2 + y^2 + ax + by = c  ==>  (x-a)^2 + (y-b)^2 = r^2
        'C=(α,β)=(−a/2,−b/2); r=sqrt(α^2+β^2+c)
        Dim alpha As Double = -solMatrix(0) / 2
        Dim beta As Double = -solMatrix(1) / 2
        Dim R As Double = Math.Sqrt(alpha ^ 2 + beta ^ 2 + solMatrix(2))

        'MsgBox(MatrixOps.printMatrix(aMatrix) & vbNewLine & vbNewLine & MatrixOps.printSingleMatrix(bMatrix) & vbNewLine & vbNewLine & MatrixOps.printMatrix(invAmatrix) & vbNewLine & vbNewLine & MatrixOps.printSingleMatrix(solMatrix))

        'MsgBox("Center: (" & FormatNumber(Math.Round(alpha, 4), 4) & "," & FormatNumber(Math.Round(beta, 4), 4) & ")" & vbNewLine & "Radius: " & FormatNumber(Math.Round(R, 4), 4))

        ReDim Preserve points(3)
        points(3) = New Point(4, 10)

        Dim test = avgResidualSqr(points, alpha, beta, R)
    End Sub

    Function avgResidualSqr(points() As Point, alpha As Double, beta As Double, R As Double) As Double

        Dim totRvar As Double = 0
        Dim pntCnt As Integer = points.Length

        'Calculate 
        Dim i As Integer
        For i = 0 To pntCnt - 1
            totRvar = totRvar + pointDeltaError(alpha, beta, R, points(i)) ^ 2
        Next

        Return totRvar / pntCnt
    End Function


    Function pointDeltaError(alpha As Double, beta As Double, R As Double, point As Point) As Double
        Return Math.Sqrt((point.X - alpha) ^ 2 + (point.Y - beta) ^ 2) - R
    End Function
End Class

Public Class Point
    Public X As Double
    Public Y As Double

    Public Sub New(xVal As Double, yVal As Double)
        X = xVal
        Y = yVal
    End Sub
End Class
