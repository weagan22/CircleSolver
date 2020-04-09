Public Class MainForm

    Public Shared Property CATIA As INFITF.Application

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If RunCheck() Then Exit Sub

        Dim uSel As INFITF.Selection = MainForm.CATIA.ActiveDocument.Selection

        Dim TheSPAWorkbench As INFITF.Workbench = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")

        Dim projPlane As HybridShapeTypeLib.Plane = uSel.Item(1).Value

        Dim PlaneMeasure As SPATypeLib.Measurable
        PlaneMeasure = TheSPAWorkbench.GetMeasurable(projPlane)

        Dim oPlaneComps(8) As Object
        PlaneMeasure.GetPlane(oPlaneComps)

        Dim matrixOps As MatrixOps = New MatrixOps

        'Plane vector dir and origin
        Dim plnXYZ() As Double = {oPlaneComps(0), oPlaneComps(1), oPlaneComps(2)}
        Dim plnU() As Double = {oPlaneComps(3), oPlaneComps(4), oPlaneComps(5)}
        Dim plnV() As Double = {oPlaneComps(6), oPlaneComps(7), oPlaneComps(8)}

        'Plane normal vector
        Dim plnN() As Double = MatrixOps.crossProd(plnU, plnV)


        'Load in the selected points from CATIA
        Dim selPoints(uSel.Count - 2) As Object
        For i = 2 To uSel.Count
            selPoints(i - 2) = uSel.Item(i).Value
        Next


        'Initialize points as 'Points' type
        Dim points() As Point = getPoints(selPoints, plnXYZ, plnU, plnV, plnN)



        'Run fittng algorithm on the points
        Dim fitCir = New BestFitMinMax 'BestFitMinMax 'BestFit
        Try
            fitCir.fitCircle(points)
        Catch ex As Exception
            MsgBox("Failed to fit circle to the points." & vbNewLine & vbNewLine & ex.Message, , "Error")
        End Try


#Region "Excel test output to show plot of nearby solutions"


        '        Dim Excel As Object
        '        Excel = CreateObject("Excel.Application")
        '        Excel.Visible = True
        '        Excel.workbooks.Add

        '        Dim originalX As Double = fitCir.centerPnt.X
        '        Dim originalY As Double = fitCir.centerPnt.Y

        '        Dim minError As Double = fitCir.maxRad - fitCir.rHat

        '        Dim rowNum As Integer = 2
        '        For z = -10 To 10
        '            Dim colNum As Integer = 2
        '            For y = -10 To 10

        '                Dim testCP = New Point(originalX + ((z / 5) * fitCir.rHat / 10), originalY + ((y / 10) * fitCir.rHat / 5))

        '                Dim minRad As Double = 0
        '                Dim maxRad As Double = 0

        '                Dim BFforEq = New BestFit

        '                For i = 0 To UBound(points)
        '                    If i = 0 Then
        '                        minRad = BFforEq.distP2P(points(i), testCP)
        '                        maxRad = BFforEq.distP2P(points(i), testCP)
        '                    ElseIf BFforEq.distP2P(points(i), testCP) < minRad Then
        '                        minRad = BFforEq.distP2P(points(i), testCP)
        '                    ElseIf BFforEq.distP2P(points(i), testCP) > maxRad Then
        '                        maxRad = BFforEq.distP2P(points(i), testCP)
        '                    End If
        '                Next


        '                fitCir.optimumRad(points)

        '                Excel.cells(rowNum, 1) = testCP.X
        '                Excel.cells(1, colNum) = testCP.Y
        '                Excel.cells(rowNum, colNum) = maxRad - minRad

        '                If z = 0 And y = 0 Then
        '                    Excel.cells(rowNum, colNum).font.bold = True
        '                End If

        '                colNum = colNum + 1
        '            Next
        '            rowNum = rowNum + 1
        '        Next

#End Region



        MsgBox(minMaxResidual(points, fitCir.centerPnt.X, fitCir.centerPnt.Y, fitCir.rHat))

        'Add best fit circle geometry into CATIA
        Dim cirCenter() As Double = MatrixOps.matrixAddSingle(MatrixOps.matrixMultByConstSingle(plnU, fitCir.centerPnt.X), MatrixOps.matrixMultByConstSingle(plnV, fitCir.centerPnt.Y))

        Dim originProj() As Double = {0, 0, 0}
        originProj = MatrixOps.projPntToPln(originProj, plnXYZ, plnN)

        cirCenter = MatrixOps.projPntToPln(cirCenter, plnXYZ, plnN)

        cirCenter = MatrixOps.matrixAddSingle(MatrixOps.matrixSubtractSingle(plnXYZ, originProj), cirCenter)


        Dim uPart As MECMOD.Part = CATIA.ActiveDocument.Part

        Dim hybShpFac As HybridShapeTypeLib.HybridShapeFactory = uPart.HybridShapeFactory

        Dim newPoint As HybridShapeTypeLib.HybridShapePointCoord = hybShpFac.AddNewPointCoord(cirCenter(0), cirCenter(1), cirCenter(2))
        newPoint.Name = "BF Center"

        Dim BFcircle As HybridShapeTypeLib.HybridShapeCircleCtrRad = hybShpFac.AddNewCircleCtrRad(newPoint, projPlane, True, fitCir.rHat)
        BFcircle.Name = fitCir.fitType & " BF Circle"

        Dim BFCircleHybBody As MECMOD.HybridBody = Nothing

        For i = 1 To uPart.HybridBodies.Count
            If uPart.HybridBodies.Item(i).Name = "BF Circles" Then
                BFCircleHybBody = uPart.HybridBodies.Item(i)
            End If
        Next

        If BFCircleHybBody Is Nothing Then
            BFCircleHybBody = uPart.HybridBodies.Add()
            BFCircleHybBody.Name = "BF Circles"
        End If

        BFCircleHybBody.AppendHybridShape(BFcircle)

        If fitCir.fitType = "Min_Sep" Then
            BFcircle = hybShpFac.AddNewCircleCtrRad(newPoint, projPlane, True, fitCir.maxRad)
            BFcircle.Name = "Max Inscribed Circle"
            BFCircleHybBody.AppendHybridShape(BFcircle)

            BFcircle = hybShpFac.AddNewCircleCtrRad(newPoint, projPlane, True, fitCir.minRad)
            BFcircle.Name = "Min Circumscribed Circle"
            BFCircleHybBody.AppendHybridShape(BFcircle)
        End If

        uPart.Update()
        uPart.Update()

        Me.Close()
    End Sub

    Function getPoints(selPoints() As Object, plnXYZ() As Double, plnU() As Double, plnV() As Double, plnN() As Double) As Point()

        If plnXYZ.Length <> 3 Or plnN.Length <> 3 Or plnU.Length <> 3 Or plnV.Length <> 3 Then
            Throw New Exception("Get points only works for 3D vectors.")
        End If

        Dim TheSPAWorkbench As INFITF.Workbench = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")

        Dim retArr(UBound(selPoints)) As Point


        For i = 0 To UBound(selPoints)
            Dim pointMeasure As SPATypeLib.Measurable
            pointMeasure = TheSPAWorkbench.GetMeasurable(selPoints(i))

            'Measure point location from CATIA
            Dim oPoint(2) As Object
            pointMeasure.GetPoint(oPoint)

            'Input coords into array pntXYZ
            Dim pntXYZ() As Double = {oPoint(0), oPoint(1), oPoint(2)}

            'Project point onto the selected plane
            Dim projPnt() As Double = MatrixOps.projPntToPln(pntXYZ, plnXYZ, plnN)

            'Calculate coords with respect to the plane u and v unit vectors
            Dim dirToProjPnt() As Double = MatrixOps.matrixSubtractSingle(projPnt, plnXYZ)

            Dim xVal As Double = MatrixOps.dotProd(dirToProjPnt, plnU)
            Dim yVal As Double = MatrixOps.dotProd(dirToProjPnt, plnV)

            retArr(i) = New Point(xVal, yVal)
        Next

        Return retArr
    End Function

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

    Function minMaxResidual(points() As Point, alpha As Double, beta As Double, R As Double) As String

        Dim minRes As Double
        Dim maxRes As Double
        Dim pntCnt As Integer = points.Length

        'Calculate 
        Dim i As Integer
        For i = 0 To pntCnt - 1
            If i = 0 Then
                minRes = pointDeltaError(alpha, beta, R, points(i))
                maxRes = pointDeltaError(alpha, beta, R, points(i))
            ElseIf pointDeltaError(alpha, beta, R, points(i)) < minRes Then
                minRes = pointDeltaError(alpha, beta, R, points(i))
            ElseIf pointDeltaError(alpha, beta, R, points(i)) > maxRes Then
                maxRes = pointDeltaError(alpha, beta, R, points(i))
            End If
        Next

        Return "Max Error: " & CStr(maxRes) & " | Min Error: " & CStr(minRes)
    End Function


    Function pointDeltaError(alpha As Double, beta As Double, R As Double, point As Point) As Double
        Return Math.Sqrt((point.X - alpha) ^ 2 + (point.Y - beta) ^ 2) - R
    End Function

    Function GetCATIA() As Boolean
        Dim countTest As Integer = 0

        For Each p As Process In Process.GetProcesses()
            If p.ProcessName = "CNEXT" Then
                If countTest = 1 Then
                    MsgBox("More than one instance of CATIA is running please make sure only one instance is running prior to running this macro",, "Error")
                    Return False
                End If
                countTest = 1
            End If
        Next

        If countTest = 0 Then
            MsgBox("CATIA is not running.",, "Error")
            Return False
            End
        End If

        'Get CATIA object
        Try
            CATIA = GetObject(, "CATIA.application")
        Catch
            Return False
        End Try

        Return True
    End Function

    Private Function RunCheck(Optional requiredType As String = "") As Boolean
        If GetCATIA() = False Then Return True

        If requiredType = "Drawing" Then
            If CATIA.GetWorkbenchId <> "Drw" And CATIA.GetWorkbenchId <> "DrwBG" Then
                Me.Hide()
                MsgBox("Only for drawings.", , "Error")
                Me.Show()
                Return True
            Else
                CATIA.ActiveDocument.Sheets.Item(1).Views.Item(2).Name = "Background View"
            End If

        ElseIf requiredType = "Part" Then
            If TypeName(CATIA.ActiveDocument) <> "PartDocument" Then
                Me.Hide()
                MsgBox("Please make sure that your active document is a part.",, "Error")
                Me.Show()
                Return True
            End If

        ElseIf requiredType = "Composite" Then
            If CATIA.GetWorkbenchId <> "CompositesWorkbench" Then
                Me.Hide()
                MsgBox("Please make sure that you are on composite workbench.",, "Error")
                Me.Show()
                Return True
            End If
        End If

        Return False
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

