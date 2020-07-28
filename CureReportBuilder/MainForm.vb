Imports System.Runtime.CompilerServices

Public Class MainForm

    Public cureProfiles() As CureProfile
    Dim curePro As CureProfile = New CureProfile

    Public partValues As New Dictionary(Of String, String) From {{"JobNum", ""}, {"PONum", ""}, {"PartNum", ""}, {"PartRev", ""}, {"PartNom", ""}, {"ProgramNum", ""}, {"PartQty", ""}, {"DataPath", ""}}

    Public loadedDataSet(,) As String

    Dim dataCnt As Integer = 0
    Dim headerRow As Integer = 0
    Dim headerCount As Integer = 2

    Dim cureStart As Integer = 0
    Dim cureEnd As Integer = 0

    Dim dateValues As New Dictionary(Of String, DateTime) From {{"startTime", Nothing}, {"endTime", Nothing}}

    Public machType As String = "Unknown"

    Public dateArr() As DateTime
    Dim stepVal As Integer = 2
    Dim partTC_Arr() As DataSet
    Dim vac_Arr() As DataSet
    Dim vessel_TC As DataSet = New DataSet(0, "vessel_TC")
    Dim vesselPress As DataSet = New DataSet(0, "vessel_Press")

    Dim leadTC As DataSet = New DataSet(0, "leadTC")
    Dim lagTC As DataSet = New DataSet(0, "lagTC")

    Dim minVac As DataSet = New DataSet(0, "minVac")
    Dim maxVac As DataSet = New DataSet(0, "maxVac")

    Dim usrRunTC() As Integer = {1} ', 2}
    Dim usrRunVac() As Integer = {2, 3, 5}



    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call errorReset()

        Call loadCureProfiles("C:\Users\Will Eagan\Source\Repos\CureReportBuilder\CureReportBuilder\Sample Files") '\test.cprof")

        loadCSVin("C:\Users\Will Eagan\source\repos\CureReportBuilder\CureReportBuilder\Sample Files\DA-18-20.csv")
        'loadCSVin("C:\Users\Will Eagan\source\repos\CureReportBuilder\CureReportBuilder\Sample Files\BATCH 38 JOB 101573, 101574 1-23-20.CSV")
        'loadCSVin("C:\Users\Will Eagan\source\repos\CureReportBuilder\CureReportBuilder\Sample Files\Autoclave Simple.CSV")

        curePro = cureProfiles(0)

        Call loadCureData()






        Call runCalc()

        Call outputResults()

    End Sub

    Sub outputResults()
        Dim Excel As Object
        Excel = CreateObject("Excel.Application")
        Excel.Visible = True
        Excel.workbooks.Add

        Dim currentRow As Integer = 1

        Call outputExcel(Excel, "Cure Name:", "'" & curePro.Name, currentRow, 1)
        Call outputExcel(Excel, "Cure Doc:", curePro.cureDoc, currentRow, 1)
        Call outputExcel(Excel, "Cure Rev", curePro.cureDocRev, currentRow, 1)
        Call outputExcel(Excel, "", "", currentRow, 1)
        Call outputExcel(Excel, "", "", currentRow, 1)

        Dim i As Integer = 0
        For i = 0 To UBound(curePro.CureSteps)
            Dim currentStep As CureStep = curePro.CureSteps(i)

            Call outputExcel(Excel, "Cure Step: ", currentStep.stepName, currentRow, 1)
            Call outputExcel(Excel, "Step Pass: ", currentStep.stepPass, currentRow, 1)
            Call outputExcel(Excel, "", "", currentRow, 1)

            If curePro.checkTemp Then
                Call outputExcel(Excel, "Temp", "", currentRow, 1)
                Call outputExcel(Excel, "Temp Pass", currentStep.tempPass, currentRow, 1)
                Call outputExcel(Excel, "Temp Max", currentStep.tempResult("Max"), currentRow, 1)
                Call outputExcel(Excel, "Temp Min", currentStep.tempResult("Min"), currentRow, 1)
                Call outputExcel(Excel, "Temp Avg", currentStep.tempResult("Avg"), currentRow, 1)
                Call outputExcel(Excel, "Temp Ramp Max", currentStep.tempResult("MaxRamp"), currentRow, 1)
                Call outputExcel(Excel, "Temp Ramp Min", currentStep.tempResult("MinRamp"), currentRow, 1)
                Call outputExcel(Excel, "Temp Ramp Avg", currentStep.tempResult("AvgRamp"), currentRow, 1)
                Call outputExcel(Excel, "", "", currentRow, 1)
            End If

            If curePro.checkPressure Then
                Call outputExcel(Excel, "Pressure", "", currentRow, 1)
                Call outputExcel(Excel, "Pressure Pass", currentStep.pressurePass, currentRow, 1)
                Call outputExcel(Excel, "Pressure Max", currentStep.pressureResult("Max"), currentRow, 1)
                Call outputExcel(Excel, "Pressure Min", currentStep.pressureResult("Min"), currentRow, 1)
                Call outputExcel(Excel, "Pressure Avg", currentStep.pressureResult("Avg"), currentRow, 1)
                Call outputExcel(Excel, "Pressure Ramp Max", currentStep.pressureResult("MaxRamp"), currentRow, 1)
                Call outputExcel(Excel, "Pressure Ramp Min", currentStep.pressureResult("MinRamp"), currentRow, 1)
                Call outputExcel(Excel, "Pressure Ramp Avg", currentStep.pressureResult("AvgRamp"), currentRow, 1)
                Call outputExcel(Excel, "", "", currentRow, 1)
            End If

            If curePro.checkVac Then
                Call outputExcel(Excel, "vac", "", currentRow, 1)
                Call outputExcel(Excel, "vac Pass", currentStep.vacPass, currentRow, 1)
                Call outputExcel(Excel, "vac Max", currentStep.vacResult("Max"), currentRow, 1)
                Call outputExcel(Excel, "vac Min", currentStep.vacResult("Min"), currentRow, 1)
                Call outputExcel(Excel, "vac Avg", currentStep.vacResult("Avg"), currentRow, 1)
                Call outputExcel(Excel, "", "", currentRow, 1)
            End If


            Call outputExcel(Excel, "", "", currentRow, 1)
        Next

        currentRow = 6
        For i = 0 To UBound(curePro.CureSteps)
            Dim currentStep As CureStep = curePro.CureSteps(i)

            Call outputExcel(Excel, "", "", currentRow, 4)
            Call outputExcel(Excel, "", "", currentRow, 4)
            Call outputExcel(Excel, "", "", currentRow, 4)

            If curePro.checkTemp Then
                Call outputExcel(Excel, "", "", currentRow, 4)
                Call outputExcel(Excel, "", "", currentRow, 4)
                Call outputExcel(Excel, "Setpoint", currentStep.tempSet("SetPoint"), currentRow, 4)
                Call outputExcel(Excel, "PosTol", currentStep.tempSet("PosTol"), currentRow, 4)
                Call outputExcel(Excel, "NegTol", currentStep.tempSet("NegTol"), currentRow, 4)
                Call outputExcel(Excel, "RampSet", currentStep.tempSet("RampRate"), currentRow, 4)
                Call outputExcel(Excel, "RampPosTol", currentStep.tempSet("RampPosTol"), currentRow, 4)
                Call outputExcel(Excel, "RampNegTol", currentStep.tempSet("RampNegTol"), currentRow, 4)
                Call outputExcel(Excel, "", "", currentRow, 4)
            End If



            If curePro.checkPressure Then
                Call outputExcel(Excel, "", "", currentRow, 4)
                Call outputExcel(Excel, "Setpoint", currentStep.pressureSet("SetPoint"), currentRow, 4)
                Call outputExcel(Excel, "PosTol", currentStep.pressureSet("PosTol"), currentRow, 4)
                Call outputExcel(Excel, "NegTol", currentStep.pressureSet("NegTol"), currentRow, 4)
                Call outputExcel(Excel, "RampSet", currentStep.pressureSet("RampRate"), currentRow, 4)
                Call outputExcel(Excel, "RampPosTol", currentStep.pressureSet("RampPosTol"), currentRow, 4)
                Call outputExcel(Excel, "RampNegTol", currentStep.pressureSet("RampNegTol"), currentRow, 4)
                Call outputExcel(Excel, "", "", currentRow, 4)
                Call outputExcel(Excel, "", "", currentRow, 4)
            End If

            If curePro.checkVac Then

                Call outputExcel(Excel, "", "", currentRow, 4)
                Call outputExcel(Excel, "Setpoint", currentStep.vacSet("SetPoint"), currentRow, 4)
                Call outputExcel(Excel, "PosTol", currentStep.vacSet("PosTol"), currentRow, 4)
                Call outputExcel(Excel, "NegTol", currentStep.vacSet("NegTol"), currentRow, 4)
                Call outputExcel(Excel, "RampSet", currentStep.vacSet("RampRate"), currentRow, 4)
                Call outputExcel(Excel, "", "", currentRow, 4)
            End If

            Call outputExcel(Excel, "", "", currentRow, 4)
        Next
    End Sub

    Sub outputExcel(Excel As Object, desc As String, uVal As String, ByRef currentRow As Integer, inCol As Integer)
        Excel.cells(currentRow, inCol) = desc
        Excel.cells(currentRow, inCol + 1) = uVal
        currentRow = currentRow + 1
    End Sub

    Sub runCalc()
        Call leadlagTC()

        If curePro.checkVac Then
            Call leadlagVac()
        End If

        Call startEndTime()
        Call cureStepTest()
        Call cureStepResults()

    End Sub

    Sub cureStepResults()
        Dim i As Integer
        For i = 0 To UBound(curePro.CureSteps)
            If curePro.CureSteps(i).hardFail Then
                Exit For
            End If

            Dim firstStep As Boolean = False
            Dim lastStep As Boolean = False

            If i = 0 Then firstStep = True
            If i = UBound(curePro.CureSteps) Then lastStep = True

            Dim indexStart As Integer = curePro.CureSteps(i).stepStart
            Dim indexEnd As Integer = curePro.CureSteps(i).stepEnd

            Dim total As Double = 0
            Dim addCnt As Integer = 0

            'Calculate temp results for a given step
            If curePro.checkTemp = True Then

                ''Max temp
                If Not firstStep Then
                    If curePro.CureSteps(i - 1).tempSet("RampRate") < 0 Then
                        indexStart = indexStart + curePro.CureSteps(i - 1).tempSet("RampRate") \ (dateArr(1) - dateArr(0)).TotalMinutes
                    End If
                End If

                If Not lastStep Then
                    If curePro.CureSteps(i + 1).tempSet("RampRate") > 0 Then
                        indexEnd = indexEnd - curePro.CureSteps(i - 1).tempSet("RampRate") \ (dateArr(1) - dateArr(0)).TotalMinutes
                    End If
                End If

                curePro.CureSteps(i).tempResult("Max") = leadTC.Max(indexStart, indexEnd)


                ''Min temp
                If Not firstStep Then
                    If curePro.CureSteps(i - 1).tempSet("RampRate") > 0 Then
                        indexStart = indexStart + curePro.CureSteps(i - 1).tempSet("RampRate") \ (dateArr(1) - dateArr(0)).TotalMinutes
                    End If
                End If

                If Not lastStep Then
                    If curePro.CureSteps(i + 1).tempSet("RampRate") < 0 Then
                        indexEnd = indexEnd - curePro.CureSteps(i - 1).tempSet("RampRate") \ (dateArr(1) - dateArr(0)).TotalMinutes
                    End If
                End If

                curePro.CureSteps(i).tempResult("Min") = lagTC.Min(indexStart, indexEnd)


                ''Average temp
                total = 0
                addCnt = 0
                indexStart = curePro.CureSteps(i).stepStart
                indexEnd = curePro.CureSteps(i).stepEnd

                For z = 0 To UBound(partTC_Arr)
                    For y = 0 To UBound(usrRunTC)
                        If partTC_Arr(z).Number = usrRunTC(y) Then
                            total = total + partTC_Arr(z).Average(indexStart, indexEnd)
                            addCnt = addCnt + 1
                        End If
                    Next
                Next

                curePro.CureSteps(i).tempResult("Avg") = total / addCnt

                ''Max temp ramp
                Dim holder As Double = 0
                indexStart = curePro.CureSteps(i).stepStart
                indexEnd = curePro.CureSteps(i).stepEnd

                For z = 0 To UBound(partTC_Arr)
                    For y = 0 To UBound(usrRunTC)
                        If holder = 0 Then
                            holder = partTC_Arr(z).MaxRamp(indexStart, indexEnd)
                        ElseIf partTC_Arr(z).MaxRamp(indexStart, indexEnd) > holder Then
                            holder = partTC_Arr(z).MaxRamp(indexStart, indexEnd)
                        End If
                    Next
                Next

                curePro.CureSteps(i).tempResult("MaxRamp") = holder

                ''Min temp ramp
                holder = 0
                indexStart = curePro.CureSteps(i).stepStart
                indexEnd = curePro.CureSteps(i).stepEnd

                For z = 0 To UBound(partTC_Arr)
                    For y = 0 To UBound(usrRunTC)
                        If holder = 0 Then
                            holder = partTC_Arr(z).MinRamp(indexStart, indexEnd)
                        ElseIf partTC_Arr(z).MinRamp(indexStart, indexEnd) < holder Then
                            holder = partTC_Arr(z).MinRamp(indexStart, indexEnd)
                        End If
                    Next
                Next
                curePro.CureSteps(i).tempResult("MinRamp") = holder


                ''Average temp ramp
                total = 0
                addCnt = 0
                indexStart = curePro.CureSteps(i).stepStart
                indexEnd = curePro.CureSteps(i).stepEnd

                For z = 0 To UBound(partTC_Arr)
                    For y = 0 To UBound(usrRunTC)
                        If partTC_Arr(z).Number = usrRunTC(y) Then
                            total = total + partTC_Arr(z).AverageRamp(indexStart, indexEnd)
                            addCnt = addCnt + 1
                        End If
                    Next
                Next

                curePro.CureSteps(i).tempResult("AvgRamp") = total / addCnt

                'Check temp for passing
                If curePro.CureSteps(i).tempResult("Min") < curePro.CureSteps(i).tempSet("SetPoint") + curePro.CureSteps(i).tempSet("NegTol") Then curePro.CureSteps(i).tempPass = False
                If curePro.CureSteps(i).tempResult("Max") > curePro.CureSteps(i).tempSet("SetPoint") + curePro.CureSteps(i).tempSet("PosTol") Then curePro.CureSteps(i).tempPass = False
                If curePro.CureSteps(i).tempResult("MinRamp") < curePro.CureSteps(i).tempSet("RampRate") + curePro.CureSteps(i).tempSet("RampNegTol") Then curePro.CureSteps(i).tempPass = False
                If curePro.CureSteps(i).tempResult("MaxRamp") > curePro.CureSteps(i).tempSet("RampRate") + curePro.CureSteps(i).tempSet("RampPosTol") Then curePro.CureSteps(i).tempPass = False
            Else
                curePro.CureSteps(i).tempPass = True
            End If

            'Check autoclave only features

            'Calculate pressure results for a given step
            If curePro.checkPressure Then
                ''Max pressure
                If Not firstStep Then
                    If curePro.CureSteps(i - 1).pressureSet("RampRate") < 0 Then
                        indexStart = indexStart + curePro.CureSteps(i - 1).pressureSet("RampRate") \ (dateArr(1) - dateArr(0)).TotalMinutes
                    End If
                End If

                If Not lastStep Then
                    If curePro.CureSteps(i + 1).pressureSet("RampRate") > 0 Then
                        indexEnd = indexEnd - curePro.CureSteps(i - 1).pressureSet("RampRate") \ (dateArr(1) - dateArr(0)).TotalMinutes
                    End If
                End If
                curePro.CureSteps(i).pressureResult("Max") = vesselPress.Max(indexStart, indexEnd)

                ''Min pressure
                If Not firstStep Then
                    If curePro.CureSteps(i - 1).pressureSet("RampRate") > 0 Then
                        indexStart = indexStart + curePro.CureSteps(i - 1).pressureSet("RampRate") \ (dateArr(1) - dateArr(0)).TotalMinutes
                    End If
                End If

                If Not lastStep Then
                    If curePro.CureSteps(i + 1).pressureSet("RampRate") < 0 Then
                        indexEnd = indexEnd - curePro.CureSteps(i - 1).pressureSet("RampRate") \ (dateArr(1) - dateArr(0)).TotalMinutes
                    End If

                End If
                curePro.CureSteps(i).pressureResult("Min") = vesselPress.Min(indexStart, indexEnd)

                ''Average pressure
                indexStart = curePro.CureSteps(i).stepStart
                indexEnd = curePro.CureSteps(i).stepEnd
                curePro.CureSteps(i).pressureResult("Avg") = vesselPress.Average(indexStart, indexEnd)

                curePro.CureSteps(i).pressureResult("MaxRamp") = vesselPress.MaxRamp(indexStart, indexEnd)
                curePro.CureSteps(i).pressureResult("MinRamp") = vesselPress.MinRamp(indexStart, indexEnd)
                curePro.CureSteps(i).pressureResult("AvgRamp") = vesselPress.AverageRamp(indexStart, indexEnd)



                'Check pressure for passing
                If curePro.CureSteps(i).pressureResult("Min") < curePro.CureSteps(i).pressureSet("SetPoint") + curePro.CureSteps(i).pressureSet("NegTol") Then curePro.CureSteps(i).pressurePass = False
                If curePro.CureSteps(i).pressureResult("Max") > curePro.CureSteps(i).pressureSet("SetPoint") + curePro.CureSteps(i).pressureSet("PosTol") Then curePro.CureSteps(i).pressurePass = False
                If curePro.CureSteps(i).pressureResult("MinRamp") < curePro.CureSteps(i).pressureSet("RampRate") + curePro.CureSteps(i).pressureSet("RampNegTol") Then curePro.CureSteps(i).pressurePass = False
                If curePro.CureSteps(i).pressureResult("MaxRamp") > curePro.CureSteps(i).pressureSet("RampRate") + curePro.CureSteps(i).pressureSet("RampPosTol") Then curePro.CureSteps(i).pressurePass = False
            Else
                curePro.CureSteps(i).pressurePass = True
            End If

            'Calculate vac results for a given step
            If curePro.checkVac Then
                curePro.CureSteps(i).vacResult("Max") = maxVac.Max(indexStart, indexEnd)
                curePro.CureSteps(i).vacResult("Min") = minVac.Min(indexStart, indexEnd)

                ''Average vac
                total = 0
                addCnt = 0
                indexStart = curePro.CureSteps(i).stepStart
                indexEnd = curePro.CureSteps(i).stepEnd

                For z = 0 To UBound(vac_Arr)
                    For y = 0 To UBound(usrRunVac)
                        If vac_Arr(z).Number = usrRunVac(y) Then
                            total = total + vac_Arr(z).Average(indexStart, indexEnd)
                            addCnt = addCnt + 1
                        End If
                    Next
                Next

                curePro.CureSteps(i).vacResult("Avg") = total / addCnt

                'Check vacuum for passing
                If curePro.CureSteps(i).vacResult("Min") < curePro.CureSteps(i).vacSet("SetPoint") + curePro.CureSteps(i).vacSet("NegTol") Then curePro.CureSteps(i).vacPass = False
                If curePro.CureSteps(i).vacResult("Max") > curePro.CureSteps(i).vacSet("SetPoint") + curePro.CureSteps(i).vacSet("PosTol") Then curePro.CureSteps(i).vacPass = False
            Else
                curePro.CureSteps(i).vacPass = True
            End If



            'Check for all passing
            If curePro.CureSteps(i).vacPass And curePro.CureSteps(i).tempPass And curePro.CureSteps(i).pressurePass Then
                curePro.CureSteps(i).stepPass = True
            Else
                curePro.CureSteps(i).stepPass = False
                curePro.curePass = False
            End If

        Next
    End Sub

    Sub cureStepTest()
        Dim currentStep As Integer = 0
        Dim i As Integer
        For i = cureStart To cureEnd
            If curePro.CureSteps(currentStep).stepStart = 0 Then
                curePro.CureSteps(currentStep).stepStart = i
            End If

            If meetTerms(curePro.CureSteps(currentStep), i) Then
                curePro.CureSteps(currentStep).stepEnd = i
                If UBound(curePro.CureSteps) = currentStep Then
                    curePro.CureSteps(currentStep).stepEnd = cureEnd
                    Exit For
                End If

                currentStep = currentStep + 1
            End If
        Next

        If UBound(curePro.CureSteps) > currentStep Then
            For i = currentStep To UBound(curePro.CureSteps)
                curePro.CureSteps(i).hardFail = True
                curePro.CureSteps(i).pressurePass = False
                curePro.CureSteps(i).tempPass = False
                curePro.CureSteps(i).vacPass = False
                curePro.CureSteps(i).stepPass = False
                curePro.curePass = False
            Next
        End If
    End Sub

    Function meetTerms(cureStep As CureStep, currentStep As Integer) As Boolean

        Dim pass1 As Boolean = termTest(cureStep.termCond1, currentStep, cureStep)
        Dim pass2 As Boolean = termTest(cureStep.termCond2, currentStep, cureStep)

        If cureStep.termCondOper = "OR" Then
            If pass1 Or pass2 Then
                Return True
            End If

        ElseIf cureStep.termCondOper = "AND" Then
            If pass1 And pass2 Then
                Return True
            End If
        Else
            Throw New Exception("Terminating condition operator for step " & cureStep.stepName & " is not valid")
        End If

        Return False
    End Function

    Function termTest(termCond As Dictionary(Of String, Object), currentStep As Integer, cureStep As CureStep) As Boolean
        If termCond("Type") = "None" Then
            Return True

        ElseIf termCond("Type") = "Time" Then
            Dim stepDuration As TimeSpan = dateArr(currentStep) - dateArr(cureStep.stepStart)

            If termCond("Condition") = "GREATER" Then
                If stepDuration.TotalMinutes > termCond("Goal") Then
                    Return True
                End If
            Else
                Throw New Exception("Terminating condition Condition for step " & cureStep.stepName & " is not valid")
            End If

        ElseIf termCond("Type") = "Temp" Then

            If termCond("Condition") = "GREATER" Then
                If termCond("TCNum") = "Lag" Then
                    Dim test = lagTC.values(currentStep)
                    If lagTC.values(currentStep) > termCond("Goal") Then
                        Return True
                    End If
                ElseIf termCond("TCNum") = "Lead" Then
                    If lagTC.values(currentStep) > termCond("Goal") Then
                        Return True
                    End If
                ElseIf termCond("TCNum") = "Air" Then
                    If vessel_TC.values(currentStep) > termCond("Goal") Then
                        Return True
                    End If
                Else
                    Throw New Exception("Terminating condition TCNum for step " & cureStep.stepName & " is not valid")
                End If

            ElseIf termCond("Condition") = "LESS" Then
                If termCond("TCNum") = "Lag" Then
                    If lagTC.values(currentStep) < termCond("Goal") Then
                        Return True
                    End If
                ElseIf termCond("TCNum") = "Lead" Then
                    If lagTC.values(currentStep) < termCond("Goal") Then
                        Return True
                    End If
                ElseIf termCond("TCNum") = "Air" Then
                    If vessel_TC.values(currentStep) < termCond("Goal") Then
                        Return True
                    End If
                Else
                    Throw New Exception("Terminating condition TCNum for step " & cureStep.stepName & " is not valid")
                End If
            Else
                Throw New Exception("Terminating condition Condition for step " & cureStep.stepName & " is not valid")
            End If

        ElseIf termCond("Type") = "Press" Then
            If termCond("Condition") = "GREATER" Then
                If vesselPress.values(currentStep) > termCond("Goal") Then
                    Return True
                End If
            ElseIf termCond("Condition") = "LESS" Then

                If vesselPress.values(currentStep) < termCond("Goal") Then
                    Return True
                End If
            Else
                Throw New Exception("Terminating condition Condition for step " & cureStep.stepName & " is not valid")
            End If

        ElseIf termCond("Type") = "Vac" Then
            'Might need this???
        Else
            Throw New Exception("Terminating condition type for step " & cureStep.stepName & " is not valid")
        End If
    End Function


    Sub startEndTime()
        Dim start_end_temp As Double = 140

        Dim i As Integer
        For i = 0 To dataCnt
            If leadTC.values(i) > start_end_temp And dateValues("startTime") = Nothing Then
                dateValues("startTime") = dateArr(i)
                cureStart = i
                Exit For
            End If
        Next
        If cureStart = 0 Then
            dateValues("startTime") = dateArr(0)
        End If

        Dim runStart As Boolean = False
        For i = 0 To dataCnt
            If lagTC.values(i) > start_end_temp And runStart = False Then
                runStart = True
                i = i + 5
            End If

            If runStart = True Then
                If lagTC.values(i) < start_end_temp Then
                    dateValues("endTime") = dateArr(i)
                    cureEnd = i
                    Exit For
                End If
            End If
        Next

        If cureEnd = 0 Then
            dateValues("endTime") = dateArr(dataCnt)
            cureEnd = dataCnt
        End If
    End Sub

    Sub leadlagTC()
        Dim i As Integer
        For i = 0 To dataCnt 'Look at each step
            Dim z As Integer
            For z = 0 To UBound(partTC_Arr) 'Look in each part array
                Dim v As Integer
                For v = 0 To UBound(usrRunTC) 'Compare to each user defined

                    If usrRunTC(v) = partTC_Arr(z).Number Then
                        If leadTC.values Is Nothing OrElse UBound(leadTC.values) < i Then
                            leadTC.values.AddValArr(partTC_Arr(z).values(i))
                        Else
                            If partTC_Arr(z).values(i) > leadTC.values(i) Then
                                leadTC.values(i) = partTC_Arr(z).values(i)
                            End If
                        End If

                        If lagTC.values Is Nothing OrElse UBound(lagTC.values) < i Then
                            lagTC.values.AddValArr(partTC_Arr(z).values(i))
                        Else
                            If partTC_Arr(z).values(i) < lagTC.values(i) Then
                                lagTC.values(i) = partTC_Arr(z).values(i)
                            End If
                        End If
                    End If
                Next
            Next
        Next
    End Sub

    Sub leadlagVac()
        Dim i As Integer
        For i = 0 To dataCnt 'Look at each step
            Dim z As Integer
            For z = 0 To UBound(vac_Arr) 'Look in each part array
                Dim v As Integer
                For v = 0 To UBound(usrRunVac) 'Compare to each user defined

                    If usrRunVac(v) = vac_Arr(z).Number Then
                        If minVac.values Is Nothing OrElse UBound(minVac.values) < i Then
                            minVac.values.AddValArr(vac_Arr(z).values(i))
                        Else
                            If vac_Arr(z).values(i) > minVac.values(i) Then
                                minVac.values(i) = vac_Arr(z).values(i)
                            End If
                        End If

                        If maxVac.values Is Nothing OrElse UBound(maxVac.values) < i Then
                            maxVac.values.AddValArr(vac_Arr(z).values(i))
                        Else
                            If vac_Arr(z).values(i) < maxVac.values(i) Then
                                maxVac.values(i) = vac_Arr(z).values(i)
                            End If
                        End If
                    End If
                Next
            Next
        Next
    End Sub

    Sub loadCureData()
        Call getTime()

        dataCnt = UBound(dateArr)

        'Calculate ramp rates over a set period for smoothing, numerator sets the number of minutes to look at
        stepVal = 5 / ((dateArr(1) - dateArr(0)).TotalMinutes)
        If stepVal = 1 Then stepVal = 2

        If curePro.checkTemp = True Then
            Dim searchVal As String = ""
            If machType = "Omega" Then
                searchVal = "Channel "
            ElseIf machType = "Autoclave" Then
                searchVal = "{Part_"
            Else
                searchVal = "TC"
            End If

            Call getDataMulti(searchVal, partTC_Arr, "TC")
        End If

        If curePro.checkPressure Then
            Call getDataSearch("{Vessel_Pressure}", vesselPress)
        End If

        If curePro.checkVac Then
            Call getDataMulti("{VacGroup_", vac_Arr, "VAC")
        End If

        If machType = "Autoclave" Then
            Call getDataSearch("{Air_TC}", vessel_TC)
        End If
    End Sub

    Sub getDataSearch(searchVal As String, ByRef loadTo As DataSet)
        Dim i As Integer
        For i = 0 To UBound(loadedDataSet, 1)
            If InStr(loadedDataSet(i, headerRow), searchVal, 0) <> 0 Then
                getData(i, loadTo)
            End If
        Next

        If loadTo.values Is Nothing Then
            Throw New Exception("Failed to find data for " & loadTo.Type & " | " & loadTo.Number)
        End If
    End Sub

    Sub getData(inCol As Integer, ByRef loadTo As DataSet)
        Dim row As Integer
        For row = headerRow + headerCount To UBound(loadedDataSet, 2)
            If IsNumeric(loadedDataSet(inCol, row)) Then
                loadTo.values.AddValArr(loadedDataSet(inCol, row))
            Else
                loadTo.values.AddValArr(0)
            End If
        Next

        loadTo.calcRamp(stepVal)
    End Sub

    Sub getDataMulti(searchVal As String, ByRef inArr() As DataSet, type As String)
        Dim i As Integer
        For i = 0 To UBound(loadedDataSet, 1)
            If InStr(loadedDataSet(i, headerRow), searchVal, 0) <> 0 Then
                If dataUsed(i, headerRow + headerCount) Then
                    If inArr Is Nothing Then
                        ReDim inArr(0)
                        inArr(0) = New DataSet(0, "")
                    Else
                        ReDim Preserve inArr(UBound(inArr) + 1)
                        inArr(UBound(inArr)) = New DataSet(0, "")
                    End If


                    inArr(UBound(inArr)).Number = Integer.Parse(System.Text.RegularExpressions.Regex.Replace(loadedDataSet(i, headerRow), "[^\d]", ""))
                    inArr(UBound(inArr)).Type = type

                    getData(i, inArr(UBound(inArr)))
                End If
            End If
        Next
        If inArr Is Nothing Then
            Throw New Exception("Failed to find data for " & type)
        End If
    End Sub

    Function dataUsed(chkCol As Integer, chkRowStart As Integer) As Boolean
        Dim r As Integer
        For r = chkRowStart To UBound(loadedDataSet, 2)
            If IsNumeric(loadedDataSet(chkCol, r)) AndAlso loadedDataSet(chkCol, r) <> 0 Then
                Return True
            End If
        Next

        Return False
    End Function

    Sub getTime()
        'Use type to determine date column and load that column into dateArr
        Dim i As Integer
        If machType = "Omega" Then
            For i = headerRow + headerCount To UBound(loadedDataSet, 2)
                dateArr.AddValArr(Convert.ToDateTime(loadedDataSet(0, i)))
            Next
        ElseIf machType = "Autoclave" Then
            For i = headerRow + headerCount To UBound(loadedDataSet, 2)
                dateArr.AddValArr(Convert.ToDateTime(loadedDataSet(1, i) & " " & loadedDataSet(2, i) & "." & loadedDataSet(3, i)))
            Next
        ElseIf machType = "Unknown" Then
            Dim dateFnd As Boolean = False
            For z = 0 To UBound(loadedDataSet, 1)
                Dim fill As DateTime
                If DateTime.TryParse(loadedDataSet(z, headerRow + headerCount), fill) Then
                    dateFnd = True
                    For i = headerRow + headerCount To UBound(loadedDataSet, 2)
                        dateArr.AddValArr(Convert.ToDateTime(loadedDataSet(z, i)))
                    Next
                End If
            Next
            If dateFnd = False Then
                Throw New Exception("Failed to find dates in the specified cure document. Check formatting.")
            End If
        End If
    End Sub

    Sub errorReset()

        'Clear out the current loaded values in arrays if they exist
        loadedDataSet.clearArr()
        dateArr.clearArr()
        partTC_Arr.clearArr()
        vac_Arr.clearArr()

        vessel_TC = New DataSet(0, "vessel_TC")
        vesselPress = New DataSet(0, "vessel_Press")


        'Reset machType to null
        machType = ""

        'Reset partValues to nothing
        partValues("JobNum") = String.Empty
        partValues("PONum") = String.Empty
        partValues("PartNum") = String.Empty
        partValues("PartRev") = String.Empty
        partValues("PartNom") = String.Empty
        partValues("ProgramNum") = String.Empty
        partValues("PartQty") = String.Empty
        partValues("DataPath") = String.Empty

        'Reset dateValues to null values
        dateValues("startTime") = Nothing
        dateValues("endTime") = Nothing

    End Sub

    Sub loadCSVin(inFile As String)

        Call errorReset()

        If IO.File.Exists(inFile) AndAlso IO.Path.GetExtension(inFile).ToLower = ".csv" Then
            Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(inFile)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters({",", vbTab})
                Dim currentRow As String()

                While Not MyReader.EndOfData
                    Try
                        currentRow = MyReader.ReadFields()
                        loadedDataSet.AddArr(currentRow)

                        'Finds file type
                        If InStr(currentRow(0), "Omega", 0) <> 0 Then
                            machType = "Omega"
                            headerRow = 2
                            headerCount = 2

                        ElseIf currentRow(0) = "No." AndAlso InStr(currentRow(1), "Date", 0) <> 0 AndAlso InStr(currentRow(2), "Time", 0) <> 0 AndAlso InStr(currentRow(3), "Millitm", 0) <> 0 AndAlso InStr(currentRow(4), "{Air_TC}", 0) <> 0 Then
                            machType = "Autoclave"
                            headerRow = 0
                            headerCount = 2
                        End If

                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                    End Try
                End While
            End Using
        End If

    End Sub

    Sub outputCureProfiles(inPath As String)
        Dim outputWriter As IO.StreamWriter = New System.IO.StreamWriter(inPath)

        Dim i As Integer
        For i = 0 To UBound(cureProfiles)
            outputWriter.Write(cureProfiles(i).serializeCure)
            If i < UBound(cureProfiles) Then
                outputWriter.Write("~&&&~" & vbNewLine)
            End If
        Next

        outputWriter.Close()
    End Sub

    Sub loadCureProfiles(inPath As String)

        If IO.File.Exists(inPath) Then
            loadCureFile(inPath)
        ElseIf IO.Directory.Exists(inPath) Then
            For Each file In IO.Directory.GetFiles(inPath)
                loadCureFile(file)
            Next
        End If


    End Sub

    Sub loadCureFile(inPath As String)
        Dim cureDef() As String

        If IO.File.Exists(inPath) And IO.Path.GetExtension(inPath) = ".cprof" Then
            cureDef = Split(IO.File.ReadAllText(inPath), "~&&&~")

            For i = 0 To UBound(cureDef)
                If cureProfiles Is Nothing Then
                    ReDim cureProfiles(0)
                    cureProfiles(i) = New CureProfile()
                Else
                    ReDim Preserve cureProfiles(UBound(cureProfiles) + 1)
                    cureProfiles(i) = New CureProfile()
                End If

                cureProfiles(i).deserializeCure(cureDef(i))
            Next
        End If
    End Sub

End Class

Module ArrayExtensions

    ''' <summary>
    ''' Addes values of a 1D array to a 2D array. 1D array length must match 1st dimension of 2D array (if they do not match, values will not be addded).
    ''' </summary>

    <Extension()>
    Public Sub AddArr(Of T)(ByRef arr As T(,), addArr() As T)
        Dim i As Integer

        If arr IsNot Nothing Then
            If UBound(addArr) <= UBound(arr, 1) Then
                ReDim Preserve arr(UBound(arr, 1), UBound(arr, 2) + 1)

                For i = 0 To UBound(addArr)
                    arr(i, UBound(arr, 2)) = addArr(i)
                Next
            End If
        Else
            ReDim arr(UBound(addArr), 0)
            For i = 0 To UBound(arr, 1)
                arr(i, UBound(arr, 2)) = addArr(i)
            Next
        End If
    End Sub

    ''' <summary>
    ''' Extends the array as required and adds the values to the end of the array.
    ''' </summary>

    <Extension()>
    Public Sub AddValArr(Of T)(ByRef arr As T(), addVal As T)
        If arr IsNot Nothing Then
            ReDim Preserve arr(UBound(arr, 1) + 1)
            arr(UBound(arr, 1)) = addVal
        Else
            ReDim arr(0)
            arr(0) = addVal
        End If
    End Sub

    ''' <summary>
    ''' Completely clears an array and sets it back to its declared state
    ''' </summary>

    <Extension()>
    Public Sub clearArr(Of T)(ByRef arr As T())
        If arr IsNot Nothing Then
            Array.Clear(arr, 0, arr.Length)
            arr = Nothing
        End If
    End Sub

    <Extension()>
    Public Sub clearArr(Of T)(ByRef arr As T(,))
        If arr IsNot Nothing Then
            Array.Clear(arr, 0, arr.Length)
            arr = Nothing
        End If
    End Sub

End Module

