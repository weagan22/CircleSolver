Class BestFit
    Public centerPnt As Point = New Point(0, 0)
    Public rHat As Double
    Dim J As Double
    Dim dJdx As Double
    Dim dJdy As Double
    Public completedIterations As Integer

    Sub fitCircle(points() As Point)
        Call initializeFit(points)
        completedIterations = minimize(points, 1000, 0.1, 0.000000000001)
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

        Dim totalRad As Double = 0
        For i = 0 To UBound(points)
            totalRad = totalRad + distP2P(points(i), centerPnt)
        Next

        rHat = totalRad / points.Length
        Return rHat
    End Function

    Function minimize(points() As Point, maxIter As Integer, inThresh As Double, outThresh As Double) As Integer
        Call computeCost(points)

        If J < 0.0000000001 Or Math.Sqrt(dJdx ^ 2 + dJdy ^ 2) < 0.0000000001 Then
            'Already at the local min
            Return 0
        End If

        Dim prevJ As Double = J
        Dim prevU As Double = 0
        Dim prevV As Double = 0
        Dim prevDJdx As Double = 0
        Dim prevDJdy As Double = 0

        For i = 0 To maxIter

            'Search direction
            Dim u As Double = -dJdx
            Dim v As Double = -dJdy

            If i <> 0 Then
                'Polak-Ribiere coefficient
                Dim beta As Double = (dJdx * (dJdx - prevDJdx) + dJdy * (dJdy - prevDJdy)) / (prevDJdx ^ 2 + prevDJdy ^ 2)
                u = u + beta * prevU
                v = v + beta * prevV
            End If

            prevDJdx = dJdx
            prevDJdy = dJdy
            prevU = u
            prevV = v

            'Rough minimization along the search direction
            Dim innerJ As Double = 0
            Do While i < maxIter And (Math.Abs(J - innerJ) / J) > inThresh
                innerJ = J

                Dim lambda As Double = newtonStep(u, v, points)

                centerPnt.X = centerPnt.X + (lambda * u)
                centerPnt.Y = centerPnt.Y + (lambda * v)

                Call optimumRad(points)
                Call computeCost(points)

                i = i + 1
            Loop


            'global convergence test
            If ((Math.Abs(J - prevJ) / J) < outThresh) Then
                Return i
            End If

            prevJ = J
        Next

        Throw New Exception("Max iterations were run prior to convergence.")
    End Function

    Sub computeCost(points() As Point)
        'Equation 6 from paper
        'Returns cost value

        J = 0
        dJdx = 0
        dJdy = 0

        For i = 0 To UBound(points)
            Dim di As Double = distP2P(points(i), centerPnt)

            If di < 0.0000000001 Then
                Throw New Exception("Cost sigularity, point at circle center")
            End If

            Dim dr As Double = di - rHat
            Dim ratio As Double = dr / di


            J = J + dr * (di + rHat)
            dJdx = dJdx + (points(i).X - centerPnt.X) * ratio
            dJdy = dJdy + (points(i).Y - centerPnt.Y) * ratio
        Next

        dJdx = dJdx * 2
        dJdy = dJdy * 2

    End Sub

    Function newtonStep(u As Double, v As Double, points() As Point) As Double
        'compute the first And second derivatives of the cost along the specified search direction

        Dim sum1 As Double = 0
        Dim sum2 As Double = 0
        Dim sumFac As Double = 0
        Dim sumFac2R As Double = 0

        For i = 0 To UBound(points)
            Dim dx As Double = centerPnt.X - points(i).X
            Dim dy As Double = centerPnt.Y - points(i).Y
            Dim di As Double = distP2P(points(i), centerPnt)

            Dim coeff1 As Double = (dx * u + dy * v) / di
            Dim coeff2 As Double = di - rHat
            sum1 += coeff1 * coeff2
            sum2 += coeff2 / di
            sumFac += coeff1
            sumFac2R += coeff1 * coeff1 / di
        Next

        'step length attempting to nullify the first derivative
        Return -sum1 / ((u * u + v * v) * sum2 - sumFac * sumFac / points.Length + rHat * sumFac2R)
    End Function

    Function distP2P(point1 As Point, point2 As Point) As Double
        'Pythagorean theorem solving distance between two points
        'di from paper when applied relative to center

        Return Math.Sqrt((point1.X - point2.X) ^ 2 + (point1.Y - point2.Y) ^ 2)
    End Function
End Class
