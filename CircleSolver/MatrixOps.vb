Public Class MatrixOps

    Public Shared Function matrix_Inverse(Operations_Matrix(,) As Double) As Double(,)

        If UBound(Operations_Matrix, 1) <> UBound(Operations_Matrix, 2) Then
            Err.Raise(2, "Calculate Inverse", "Matrix must be square to compute inverse")
        End If

        Dim matrixDim As Integer = UBound(Operations_Matrix, 1)

        ReDim Preserve Operations_Matrix(matrixDim, 2 * matrixDim + 1)

        'Assign values from singular matrix [I] at the right
        For C = 0 To matrixDim
            For R = 0 To matrixDim
                If R = C Then
                    Operations_Matrix(R, C + matrixDim + 1) = 1
                End If
            Next
        Next

        'Build the Singular matrix [I] at the left
        For k = 0 To matrixDim

            If Operations_Matrix(k, k) = 0 Then 'Bring a non-zero element first by changing lines if necessary

                Dim line_1 As Integer = 0

                For R = k To matrixDim
                    If Operations_Matrix(R, k) <> 0 Then line_1 = R : Exit For 'Finds line_1 with non-zero element
                Next
                'Change line k with line_1
                For C = k To (matrixDim * 2) + 1
                    Dim temporary_1 As Double = 0
                    temporary_1 = Operations_Matrix(k, C)
                    Operations_Matrix(k, C) = Operations_Matrix(line_1, C)
                    Operations_Matrix(line_1, C) = temporary_1
                Next
            End If

            Dim elem1 As Double = Operations_Matrix(k, k)
            For C = k To (matrixDim * 2) + 1
                Operations_Matrix(k, C) = Operations_Matrix(k, C) / elem1
            Next

            'For other lines, make a zero element by using:
            'Ai1=Aij-A11*(Aij/A11)
            'and change all the line using the same formula for other elements
            For R = 0 To matrixDim
                If R = k And R = matrixDim Then Exit For 'Finished

                If R = k And R < matrixDim Then Continue For 'Do not change that element (already equals to 1)

                If Operations_Matrix(R, k) <> 0 Then 'if it is zero, stays as it is
                    Dim multiplier_1 As Double = Operations_Matrix(R, k) / Operations_Matrix(k, k)

                    For m = k To (matrixDim * 2) + 1
                        Operations_Matrix(R, m) = Operations_Matrix(R, m) - Operations_Matrix(k, m) * multiplier_1
                    Next
                End If
            Next

        Next

        Dim Inverse_Matrix(,) As Double
        ReDim Inverse_Matrix(matrixDim, matrixDim)
        'Assign the right part to the Inverse_Matrix
        For R = 0 To matrixDim
            For C = 0 To matrixDim
                Inverse_Matrix(R, C) = Operations_Matrix(R, matrixDim + 1 + C)
            Next
        Next

        Return Inverse_Matrix
    End Function

    Public Shared Function matrixMult(matrix1(,) As Double, matrix2(,) As Double) As Double(,)
        If UBound(matrix1, 2) <> UBound(matrix2, 1) Then
            Err.Raise(1, "Matrix Multiply", "Check your array dimensions to make sure they can be multiplied")
        End If

        Dim returnArr(UBound(matrix1, 1), UBound(matrix2, 2)) As Double

        For r1 = 0 To UBound(matrix1, 1)
            For c2 = 0 To UBound(matrix2, 2)

                Dim dotProd As Double = 0

                For c1 = 0 To UBound(matrix1, 2)
                    Dim m1 = matrix1(r1, c1)
                    Dim m2 = matrix2(c1, c2)
                    dotProd = dotProd + matrix1(r1, c1) * matrix2(c1, c2)
                Next

                returnArr(r1, c2) = dotProd
            Next
        Next

        Return returnArr

    End Function

    Public Shared Function matrixMultSingle(matrix1(,) As Double, matrix2() As Double) As Double()
        If UBound(matrix1, 2) <> UBound(matrix2) Then
            Err.Raise(1, "Matrix Multiply", "Check your array dimensions to make sure they can be multiplied")
        End If

        Dim returnArr(UBound(matrix2)) As Double

        For r1 = 0 To UBound(matrix1, 1)
            Dim dotProd As Double = 0

            For c1 = 0 To UBound(matrix1, 2)
                Dim m1 = matrix1(r1, c1)
                Dim m2 = matrix2(c1)
                dotProd = dotProd + matrix1(r1, c1) * matrix2(c1)
            Next

            returnArr(r1) = dotProd
        Next

        Return returnArr

    End Function

    Public Shared Function matrixMultByConst(ByVal matrix(,) As Double, constantMult As Double) As Double(,)

        Dim retMatrix(,) As Double
        ReDim retMatrix(UBound(matrix, 1), UBound(matrix, 2))
        For r1 = 0 To UBound(matrix, 1)
            For c1 = 0 To UBound(matrix, 2)
                retMatrix(r1, c1) = matrix(r1, c1) * constantMult
            Next
        Next

        Return retMatrix

    End Function

    Public Shared Function matrixMultByConstSingle(matrix() As Double, constantMult As Double) As Double()

        Dim retMatrix() As Double
        ReDim retMatrix(UBound(matrix))
        For r = 0 To UBound(matrix)
            retMatrix(r) = matrix(r) * constantMult
        Next

        Return retMatrix

    End Function

    Public Shared Function matrixAdd(matrix1(,) As Double, matrix2(,) As Double) As Double(,)
        If UBound(matrix1, 1) <> UBound(matrix2, 1) Or UBound(matrix1, 2) <> UBound(matrix2, 2) Then
            Err.Raise(1, "Matrix Add", "Check your array dimensions to make sure they are equal")
        End If

        Dim returnArr(UBound(matrix1, 1), UBound(matrix1, 2)) As Double

        For r1 = 0 To UBound(matrix1, 1)
            For c1 = 0 To UBound(matrix1, 2)
                returnArr(r1, c1) = matrix1(r1, c1) + matrix2(r1, c1)
            Next
        Next

        Return returnArr
    End Function

    Public Shared Function matrixAddSingle(matrix1() As Double, matrix2() As Double) As Double()
        If UBound(matrix1, 1) <> UBound(matrix2, 1) Then
            Err.Raise(1, "Matrix Add", "Check your array dimensions to make sure they are equal")
        End If

        Dim returnArr(UBound(matrix1)) As Double

        For r = 0 To UBound(matrix1)
            returnArr(r) = matrix1(r) + matrix2(r)
        Next

        Return returnArr
    End Function

    Public Shared Function printMatrix(matrix(,) As Double, Optional showVal As Boolean = False, Optional roundNum As Boolean = True, Optional roundLength As Integer = 3) As String
        Dim message As String = ""
        For N = 0 To UBound(matrix, 1)
            For k = 0 To UBound(matrix, 2)
                If message = "" Then
                    If roundNum = True Then message = stringSpaceoutRet(Math.Round(matrix(N, k), roundLength)) & ","
                    If roundNum = False Then message = stringSpaceoutRet(matrix(N, k)) & ","
                ElseIf Strings.Right(message, 1) = vbLf Then
                    If roundNum = True Then message = message & stringSpaceoutRet(Math.Round(matrix(N, k), roundLength)) & ","
                    If roundNum = False Then message = message & stringSpaceoutRet(matrix(N, k)) & ","
                Else
                    If roundNum = True Then message = message & stringSpaceoutRet(Math.Round(matrix(N, k), roundLength)) & ","
                    If roundNum = False Then message = message & stringSpaceoutRet(matrix(N, k)) & ","
                End If
            Next
            message = Left(message, Len(message) - 1) & vbNewLine
        Next

        If showVal = True Then MsgBox(message,, "Matrix Output")
        Return message


    End Function

    Public Shared Function printSingleMatrix(matrix() As Double, Optional showVal As Boolean = False, Optional roundNum As Boolean = True, Optional roundLength As Integer = 3) As String
        Dim message As String = ""
        For N = 0 To UBound(matrix, 1)
            If message = "" Then
                If roundNum = True Then message = stringSpaceoutRet(Math.Round(matrix(N), roundLength)) & ","
                If roundNum = False Then message = stringSpaceoutRet(matrix(N)) & ","
            ElseIf Strings.Right(message, 1) = vbLf Then
                If roundNum = True Then message = message & stringSpaceoutRet(Math.Round(matrix(N), roundLength)) & ","
                If roundNum = False Then message = message & stringSpaceoutRet(matrix(N)) & ","
            Else
                If roundNum = True Then message = message & stringSpaceoutRet(Math.Round(matrix(N), roundLength)) & ","
                If roundNum = False Then message = message & stringSpaceoutRet(matrix(N)) & ","
            End If

            message = Left(message, Len(message) - 1) & vbNewLine
        Next

        If showVal = True Then MsgBox(message,, "Matrix Output")
        Return message

    End Function

    Public Shared Function stringSpaceoutRet(inStr As String) As String
        For i = 1 To 5 - Len(inStr)
            inStr = " " & inStr
        Next

        Return inStr
    End Function

    'Public Shared Function radians(degAngle As Double) As Double
    '    Return degAngle * 3.14159 / 180
    'End Function
    'Public Shared Function degrees(radAngle As Double) As Double
    '    Return radAngle * 180 / 3.14159
    'End Function
End Class
