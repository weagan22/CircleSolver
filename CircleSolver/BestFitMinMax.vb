Public Class BestFitMinMax

    Public fitType As String = "Min_Sep"
    Public centerPnt As Point = New Point(0, 0)
    Public rHat As Double
    Public minRad As Double
    Public maxRad As Double
    Public completedIterations As Integer

    Sub fitCircle(points() As Point)
        Call initializeFit(points)
        completedIterations = MinMaxMinimize(points, 10000, 0.01, 0.0001)
    End Sub

    Function initializeFit(points() As Point) As Point
        'Algorithm 1 from paper
        'Returnes the average center of the centers from every set of 3 points

        Dim totX As Double = 0
        Dim totY As Double = 0
        Dim pointCnt As Integer = UBound(points)
        Dim calcCnt As Integer = 0

        For i = 0 To pointCnt - 2
            For j_iter = i + 1 To pointCnt - 1
                For k = j_iter + 1 To pointCnt
                    Dim Xi, Xj, Xk, Yi, Yj, Yk As Double
                    Xi = points(i).X
                    Xj = points(j_iter).X
                    Xk = points(k).X

                    Yi = points(i).Y
                    Yj = points(j_iter).Y
                    Yk = points(k).Y

                    'Equation (3)
                    Dim det As Double = (Xk - Xj) * (Yj - Yi) - (Xj - Xi) * (Yk - Yj)
                    If Math.Abs(det) > 0.0000000001 Then

                        'Equation (4)
                        Dim Cx As Double = ((Yk - Yj) * (Xi ^ 2 + Yi ^ 2) + (Yi - Yk) * (Xj ^ 2 + Yj ^ 2) + (Yj - Yi) * (Xk ^ 2 + Yk ^ 2)) / (2 * det)
                        Dim Cy As Double = -((Xk - Xj) * (Xi ^ 2 + Yi ^ 2) + (Xi - Xk) * (Xj ^ 2 + Yj ^ 2) + (Xj - Xi) * (Xk ^ 2 + Yk ^ 2)) / (2 * det)

                        totX = totX + Cx
                        totY = totY + Cy
                        calcCnt = calcCnt + 1
                    End If
                Next
            Next
        Next

        If calcCnt = 0 Then
            Throw New Exception("All of the points are aligned.")
        End If

        centerPnt.X = totX / calcCnt
        centerPnt.Y = totY / calcCnt

        optimumRad(points)

        Return centerPnt
    End Function

    Function optimumRad(points() As Point) As Double
        'Equation 5 from paper
        'Returns optimum radius for a given center point

        For i = 0 To UBound(points)
            If i = 0 Then
                minRad = distP2P(points(i), centerPnt)
                maxRad = distP2P(points(i), centerPnt)
            ElseIf distP2P(points(i), centerPnt) < minRad Then
                minRad = distP2P(points(i), centerPnt)
            ElseIf distP2P(points(i), centerPnt) > maxRad Then
                maxRad = distP2P(points(i), centerPnt)
            End If
        Next

        rHat = (minRad + maxRad) / 2
        Return rHat
    End Function

    Function radTest(points() As Point, testCP As Point) As Double
        'Equation 5 from paper
        'Returns optimum radius for a given center point

        Dim minRadTest As Double
        Dim maxRadTest As Double

        For i = 0 To UBound(points)
            Dim distCalc As Double = distP2P(points(i), testCP)

            If i = 0 Then
                minRadTest = distCalc
                maxRadTest = distCalc
            ElseIf distCalc < minRadTest Then
                minRadTest = distCalc
            ElseIf distCalc > maxRadTest Then
                maxRadTest = distCalc
            End If
        Next

        Dim separation = maxRadTest - minRadTest
        Return separation
    End Function

    Function MinMaxMinimize(points() As Point, maxIter As Integer, inThresh As Double, outThresh As Double) As Integer
        Dim prevSep As Double = maxRad - minRad
        Dim searchZone As Double = rHat / 2

        For inter = 0 To maxIter

            If inter <> 0 Then
                searchZone = searchZone / 2
            End If

            Dim bestPoint As New Point(0, 0)

            Dim innerSep As Double = prevSep

            For i = -10 To 10
                For j = -10 To 10

                    Dim testCP As New Point(centerPnt.X + (searchZone * (i / 10)), centerPnt.Y + (searchZone * (j / 10)))
                    Dim testSep As Double = radTest(points, testCP)

                    If testSep < innerSep Then
                        bestPoint.X = testCP.X
                        bestPoint.Y = testCP.Y
                        innerSep = testSep
                    End If
                Next
            Next

            If radTest(points, bestPoint) < maxRad - minRad Then
                centerPnt.X = bestPoint.X
                centerPnt.Y = bestPoint.Y
            End If

            optimumRad(points)
            Dim currentSep As Double = maxRad - minRad

            If currentSep < prevSep And (prevSep - currentSep) < outThresh Then
                Return inter
            End If

            prevSep = currentSep
        Next

        Throw New Exception("Max iterations were run prior to convergence.")
    End Function




    Function distP2P(point1 As Point, point2 As Point) As Double
        'Pythagorean theorem solving distance between two points
        'di from paper when applied relative to center

        Return Math.Sqrt((point1.X - point2.X) ^ 2 + (point1.Y - point2.Y) ^ 2)
    End Function
End Class

