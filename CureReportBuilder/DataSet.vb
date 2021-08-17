Public Class DataSet
    Public Number As Integer = 0
    Public Type As String = ""
    Public values() As Double
    Public ramp() As Double

    Public Sub New(inNumber As Integer, inType As String)
        Number = inNumber
        Type = inType
    End Sub

    Public Sub calcRamp(stepRate As Integer,
                        dateArr() As DateTime)

        ReDim ramp(UBound(values))

        Dim i As Integer
        For i = 0 To UBound(values)

            Dim startVal As Integer = i - (stepRate \ 2)
            If startVal < 0 Then startVal = 0

            Dim endVal As Integer = i + (stepRate \ 2)
            If endVal > UBound(values) Then endVal = UBound(values)

            ramp(i) = LinReg(dateArr, values, startVal, endVal)
        Next
    End Sub

    Public Function Count() As Integer
        Return UBound(values)
    End Function

    Public Function Min(Optional indexStart As Integer = 0,
                        Optional indexEnd As Integer = 0) As Double

        Dim holder As Double

        If indexEnd = 0 Then indexEnd = Count()

        Dim i As Integer
        For i = indexStart To indexEnd
            If i = indexStart Then
                holder = values(i)
            ElseIf values(i) < holder Then
                holder = values(i)
            End If
        Next

        Return holder
    End Function

    Public Function Max(Optional indexStart As Integer = 0,
                        Optional indexEnd As Integer = 0) As Double

        Dim holder As Double

        If indexEnd = 0 Then indexEnd = Count()

        Dim i As Integer
        For i = indexStart To indexEnd
            If i = indexStart Then
                holder = values(i)
            ElseIf values(i) > holder Then
                holder = values(i)
            End If
        Next

        Return holder
    End Function

    Public Function Average(Optional indexStart As Integer = 0,
                            Optional indexEnd As Integer = 0) As Double

        Dim total As Double = 0
        Dim addCnt As Integer = 0

        If indexEnd = 0 Then indexEnd = Count()

        Dim i As Integer
        For i = indexStart To indexEnd
            total = total + values(i)
            addCnt = addCnt + 1
        Next

        Return total / addCnt
    End Function

    Public Function MinRamp(Optional indexStart As Integer = 0,
                            Optional indexEnd As Integer = 0,
                            Optional goal As Double = -1,
                            Optional greatThan As Boolean = True) As Double

        Dim holder As Double

        If indexEnd = 0 Then indexEnd = Count()

        Dim i As Integer
        For i = indexStart To indexEnd
            If goal = -1 Then
                If i = indexStart Then
                    holder = ramp(i)
                ElseIf ramp(i) < holder Then
                    holder = ramp(i)
                End If
            Else
                If greatThan Then
                    If values(i) < goal Then
                        If i = indexStart Then
                            holder = ramp(i)
                        ElseIf ramp(i) < holder Then
                            holder = ramp(i)
                        End If
                    End If
                Else
                    If values(i) > goal Then
                        If i = indexStart Then
                            holder = ramp(i)
                        ElseIf ramp(i) < holder Then
                            holder = ramp(i)
                        End If
                    End If
                End If
            End If
        Next

        Return holder
    End Function

    Public Function MaxRamp(Optional indexStart As Integer = 0,
                            Optional indexEnd As Integer = 0,
                            Optional goal As Double = -1,
                            Optional greatThan As Boolean = True) As Double

        Dim holder As Double

        If indexEnd = 0 Then indexEnd = Count()

        Dim i As Integer
        For i = indexStart To indexEnd
            If goal = -1 Then
                If i = indexStart Then
                    holder = ramp(i)
                ElseIf ramp(i) > holder Then
                    holder = ramp(i)
                End If
            Else
                If greatThan Then
                    If values(i) < goal Then
                        If i = indexStart Then
                            holder = ramp(i)
                        ElseIf ramp(i) > holder Then
                            holder = ramp(i)
                        End If
                    End If
                Else
                    If values(i) > goal Then
                        If i = indexStart Then
                            holder = ramp(i)
                        ElseIf ramp(i) > holder Then
                            holder = ramp(i)
                        End If
                    End If
                End If
            End If
        Next

        Return holder
    End Function

    Public Function AverageRamp(Optional indexStart As Integer = 0,
                                Optional indexEnd As Integer = 0,
                                Optional goal As Double = -1,
                                Optional rampRate As Double = 0,
                                Optional vals() As Double = Nothing) As Double

        Dim greatThan As Boolean = True

        If rampRate > 0 Then
            greatThan = True
        Else
            greatThan = False
        End If

        Dim total As Double = 0
        Dim addCnt As Integer = 0

        If indexEnd = 0 Then indexEnd = Count()

        Dim i As Integer
        For i = indexStart To indexEnd
            If goal = -1 Then
                total = total + ramp(i)
                addCnt = addCnt + 1
            Else
                If greatThan Then
                    If vals(i) < goal Then
                        total = total + ramp(i)
                        addCnt = addCnt + 1
                    End If
                Else
                    If vals(i) > goal Then
                        total = total + ramp(i)
                        addCnt = addCnt + 1
                    End If
                End If
            End If

        Next

        Return total / addCnt
    End Function


    Function LinReg(dataX() As DateTime,
                    dataY() As Double,
                    indexStart As Integer,
                    indexEnd As Integer) As Double

        '**Dim'd variables are 0 by default, no need to set them if this is all you need.**
        Dim count As Long
        Dim x As Double
        Dim y As Double
        Dim MeanX As Double
        Dim MeanY As Double
        Dim SumX As Double
        Dim SumY As Double
        Dim SumX2 As Double
        Dim SumY2 As Double
        Dim Sumy2_prime As Double
        Dim Sumx2_prime As Double
        Dim sumxy_prime As Double
        Dim Sx As Double
        Dim Sy As Double
        Dim r As Double

        '**Add error checking to make sure that data contains enough rows and cols
        If indexStart > indexEnd Then
            Throw New Exception("Index start must be less than index end.")
        End If

        If UBound(dataX) < indexEnd Then
            Throw New Exception("Data array must contain at least index end number of values")
        End If

        If UBound(dataX) <> UBound(dataY) Then
            Throw New Exception("Data x and y arrays must have the same number of values")
        End If

        Dim startTime As DateTime = dataX(indexStart)

        '**Get summations for mean**
        Dim i As Integer
        For i = indexStart To indexEnd
            x = (dataX(i) - startTime).TotalMinutes
            y = dataY(i)
            SumX = SumX + x
            SumY = SumY + y
            SumX2 = SumX2 + x ^ 2
            SumY2 = SumY2 + y ^ 2
            count = count + 1
        Next

        MeanX = SumX / count
        MeanY = SumY / count

        '**Get residuals**
        For i = indexStart To indexEnd
            x = (dataX(i) - startTime).TotalMinutes - MeanX
            y = dataY(i) - MeanY
            Sumx2_prime = Sumx2_prime + x ^ 2
            Sumy2_prime = Sumy2_prime + y ^ 2
            sumxy_prime = sumxy_prime + x * y
        Next

        '**Calculate linear regression**
        If Sumy2_prime = 0 Then
            LinReg = 0
        Else
            r = sumxy_prime / Math.Sqrt(Sumx2_prime * Sumy2_prime)
            Sx = Math.Sqrt(SumX2 - (SumX ^ 2 / count)) / count
            Sy = Math.Sqrt(SumY2 - (SumY ^ 2 / count)) / count

            LinReg = r * Sy / Sx
        End If

    End Function
End Class
