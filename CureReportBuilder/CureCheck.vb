Public Class CureCheck

    Public curePro As CureProfile = New CureProfile("null")

    'Public partValues As New Dictionary(Of String, String) From {{"JobNum", ""}, {"PONum", ""}, {"PartNum", ""}, {"PartRev", ""}, {"PartNom", ""}, {"ProgramNum", ""}, {"PartQty", ""}, {"DataPath", ""}}
    Public JobNum As String = ""
    Public PONum As String = ""
    Public PartNum As String = ""
    Public PartRev As String = ""
    Public PartNom As String = ""
    Public ProgramNum As String = ""
    Public PartQty As String = ""
    Public DataPath As String = ""
    Public completedBy As String = ""


    Public equipSerialNum As String = ""

    Public loadedDataSet(,) As String
    Public dataCnt As Integer = 0
    Public headerRow As Integer = 0
    Public headerCount As Integer = 2

    Public cureStart As Integer = 0
    Public cureEnd As Integer = 0

    'Public dateValues As New Dictionary(Of String, DateTime) From {{"startTime", Nothing}, {"endTime", Nothing}}
    Public cureStartTime As DateTime = Nothing
    Public cureEndTime As DateTime = Nothing


    Public machType As String = "Unknown"

    Public dateArr() As DateTime
    Public stepVal As Integer = 2
    Public partTC_Arr() As DataSet
    Public vac_Arr() As DataSet
    Public vessel_TC As DataSet = New DataSet(0, "vessel_TC")
    Public vesselPress As DataSet = New DataSet(0, "vessel_Press")

    Public leadTC As DataSet = New DataSet(0, "leadTC")
    Public lagTC As DataSet = New DataSet(0, "lagTC")

    Public minVac As DataSet = New DataSet(0, "minVac")
    Public maxVac As DataSet = New DataSet(0, "maxVac")

    Public usrRunTC() As Integer
    Public usrRunVac() As Integer

#Region "Data test against cure definition"
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
            Dim currentStep As CureStep = curePro.CureSteps(i)

            If currentStep.hardFail Then
                Exit For
            End If

            Dim firstStep As Boolean = False
            Dim lastStep As Boolean = False

            If i = 0 Then firstStep = True
            If i = UBound(curePro.CureSteps) Then lastStep = True

            Dim indexStart As Integer = currentStep.stepStart
            Dim indexEnd As Integer = currentStep.stepEnd
            Dim goal As Double = -1
            Dim greaterThanGoal As Boolean = True
            Dim holder As Double = 0

            Dim total As Double = 0
            Dim addCnt As Integer = 0

            Dim dataStepDuration As Double = (dateArr(UBound(dateArr)) - dateArr(0)).TotalMinutes / dataCnt

            'Check step time length
            If currentStep.stepDuration > 0 Then
                If (dateArr(currentStep.stepStart) - dateArr(currentStep.stepStart)).TotalMinutes > currentStep.stepDuration Then
                    currentStep.timeLimitPass = False
                End If
            End If



            'Calculate temp results for a given step
            If curePro.checkTemp = True And currentStep.tempSet.SetPoint <> -1 Then

                ''Max temp
                If Not firstStep Then
                    If curePro.CureSteps(i - 1).tempSet.RampRate < 0 Then
                        indexStart = indexStart + Math.Abs(Math.Round(curePro.CureSteps(i - 1).tempSet.RampRate / dataStepDuration, 0))
                    End If
                End If

                If Not lastStep Then
                    If curePro.CureSteps(i + 1).tempSet.RampRate > 0 Then
                        indexEnd = indexEnd - Math.Abs(Math.Round(curePro.CureSteps(i + 1).tempSet.RampRate / dataStepDuration, 0))
                    End If
                End If

                ''Handles a very short step: just look at max of whole step, no buffer
                If indexEnd <= indexStart Then
                    indexStart = currentStep.stepStart
                    indexEnd = currentStep.stepEnd
                End If

                currentStep.tempResult.Max = Math.Round(leadTC.Max(indexStart, indexEnd), 0)


                ''Min temp
                indexStart = currentStep.stepStart
                indexEnd = currentStep.stepEnd

                If Not firstStep Then
                    If curePro.CureSteps(i - 1).tempSet.RampRate > 0 Then
                        indexStart = indexStart + Math.Abs(Math.Round(curePro.CureSteps(i - 1).tempSet.RampRate / dataStepDuration, 0))
                    End If
                End If

                If Not lastStep Then
                    If curePro.CureSteps(i + 1).tempSet.RampRate < 0 Then
                        indexEnd = indexEnd - Math.Abs(Math.Round(curePro.CureSteps(i + 1).tempSet.RampRate / dataStepDuration, 0))
                    End If
                End If

                ''Handles a very short step: just look at min of whole step, no buffer
                If indexEnd <= indexStart Then
                    indexStart = currentStep.stepStart
                    indexEnd = currentStep.stepEnd
                End If

                currentStep.tempResult.Min = Math.Round(lagTC.Min(indexStart, indexEnd), 0)


                ''Average temp
                total = 0
                addCnt = 0
                indexStart = currentStep.stepStart
                indexEnd = currentStep.stepEnd

                For z = 0 To UBound(partTC_Arr)
                    For y = 0 To UBound(usrRunTC)
                        If partTC_Arr(z).Number = usrRunTC(y) Then
                            total = total + partTC_Arr(z).Average(indexStart, indexEnd)
                            addCnt = addCnt + 1
                        End If
                    Next
                Next

                currentStep.tempResult.Avg = Math.Round(total / addCnt, 0)




                ''##Temp Ramps##
                goal = -1
                greaterThanGoal = True

                If currentStep.tempSet.RampRate > 0 Then
                    goal = currentStep.tempSet.SetPoint
                    greaterThanGoal = True
                ElseIf currentStep.tempSet.RampRate < 0 Then
                    goal = currentStep.tempSet.SetPoint
                    greaterThanGoal = False
                End If

                ''Buffer start and end by half the linear regression step length so ramp is only from the step region
                indexStart = currentStep.stepStart + (stepVal / 2)
                indexEnd = currentStep.stepEnd - (stepVal / 2)

                ''For very short step (<linear reg step) just look at the middle point 
                If indexEnd < indexStart Then
                    Dim midPoint As Integer = indexStart + ((indexEnd - indexStart) / 2)
                    indexStart = midPoint
                    indexEnd = midPoint
                End If

                ''Max temp ramp
                holder = 0

                For z = 0 To UBound(partTC_Arr)
                    For y = 0 To UBound(usrRunTC)
                        If partTC_Arr(z).Number = usrRunTC(y) Then
                            If holder = 0 Then
                                holder = partTC_Arr(z).MaxRamp(indexStart, indexEnd, goal, greaterThanGoal)
                            ElseIf partTC_Arr(z).MaxRamp(indexStart, indexEnd, goal, greaterThanGoal) > holder Then
                                holder = partTC_Arr(z).MaxRamp(indexStart, indexEnd, goal, greaterThanGoal)
                            End If
                        End If
                    Next
                Next

                currentStep.tempResult.MaxRamp = Math.Round(holder, 1)


                ''Min temp ramp
                holder = 0

                For z = 0 To UBound(partTC_Arr)
                    For y = 0 To UBound(usrRunTC)
                        If partTC_Arr(z).Number = usrRunTC(y) Then
                            If holder = 0 Then
                                holder = partTC_Arr(z).MinRamp(indexStart, indexEnd, goal, greaterThanGoal)
                            ElseIf partTC_Arr(z).MinRamp(indexStart, indexEnd, goal, greaterThanGoal) < holder Then
                                holder = partTC_Arr(z).MinRamp(indexStart, indexEnd, goal, greaterThanGoal)
                            End If
                        End If
                    Next
                Next

                currentStep.tempResult.MinRamp = Math.Round(holder, 1)


                ''Average temp ramp
                total = 0
                addCnt = 0

                For z = 0 To UBound(partTC_Arr)
                    For y = 0 To UBound(usrRunTC)
                        If partTC_Arr(z).Number = usrRunTC(y) Then
                            total = total + partTC_Arr(z).AverageRamp(indexStart, indexEnd, currentStep.tempSet.SetPoint, currentStep.tempSet.RampRate, partTC_Arr(z).values)
                            addCnt = addCnt + 1
                        End If
                    Next
                Next

                currentStep.tempResult.AvgRamp = Math.Round(total / addCnt, 1)

                'Check temp for passing
                If Math.Abs(currentStep.tempSet.SetPoint) = Math.Abs(currentStep.tempSet.NegTol) Then
                    If currentStep.tempResult.Max > currentStep.tempSet.SetPoint + currentStep.tempSet.PosTol Then currentStep.tempPass = False
                ElseIf Math.Abs(currentStep.tempSet.SetPoint) = Math.Abs(currentStep.tempSet.PosTol) Then
                    If currentStep.tempResult.Min < currentStep.tempSet.SetPoint + currentStep.tempSet.NegTol Then currentStep.tempPass = False
                Else
                    If currentStep.tempResult.Min < currentStep.tempSet.SetPoint + currentStep.tempSet.NegTol Then currentStep.tempPass = False
                    If currentStep.tempResult.Max > currentStep.tempSet.SetPoint + currentStep.tempSet.PosTol Then currentStep.tempPass = False
                End If

                'Check temp ramp for passing if not equal to 0
                If currentStep.tempSet.RampRate <> 0 Then
                    If currentStep.tempResult.MinRamp < currentStep.tempSet.RampRate + currentStep.tempSet.RampNegTol Then currentStep.tempRampPass = False
                    If currentStep.tempResult.MaxRamp > currentStep.tempSet.RampRate + currentStep.tempSet.RampPosTol Then currentStep.tempRampPass = False
                End If
            Else
                currentStep.tempPass = True
                currentStep.tempRampPass = True
            End If



            'Check autoclave only features

            'Calculate pressure results for a given step
            If curePro.checkPressure And currentStep.pressureSet.SetPoint <> -1 Then

                ''Max pressure
                indexStart = currentStep.stepStart
                indexEnd = currentStep.stepEnd

                If Not firstStep Then
                    If curePro.CureSteps(i - 1).pressureSet.RampRate < 0 Then
                        indexStart = indexStart + Math.Abs(Math.Round(curePro.CureSteps(i - 1).pressureSet.RampRate / dataStepDuration, 0))
                    End If
                End If

                If Not lastStep Then
                    If curePro.CureSteps(i + 1).pressureSet.RampRate > 0 Then
                        indexEnd = indexEnd - Math.Abs(Math.Round(curePro.CureSteps(i + 1).pressureSet.RampRate / dataStepDuration, 0))
                    End If
                End If

                ''Handles a very short step: just look at max of whole step, no buffer
                If indexEnd <= indexStart Then
                    indexStart = currentStep.stepStart
                    indexEnd = currentStep.stepEnd
                End If

                currentStep.pressureResult.Max = Math.Round(vesselPress.Max(indexStart, indexEnd), 1)


                ''Min pressure
                indexStart = currentStep.stepStart
                indexEnd = currentStep.stepEnd
                If Not firstStep Then
                    If curePro.CureSteps(i - 1).pressureSet.RampRate > 0 Then
                        indexStart = indexStart + Math.Abs(Math.Round(curePro.CureSteps(i - 1).pressureSet.RampRate / dataStepDuration, 0))
                    End If
                End If

                If Not lastStep Then
                    If curePro.CureSteps(i + 1).pressureSet.RampRate < 0 Then
                        indexEnd = indexEnd - Math.Abs(Math.Round(curePro.CureSteps(i + 1).pressureSet.RampRate / dataStepDuration, 0))
                    End If

                End If

                ''Handles a very short step: just look at max of whole step, no buffer
                If indexEnd <= indexStart Then
                    indexStart = currentStep.stepStart
                    indexEnd = currentStep.stepEnd
                End If

                currentStep.pressureResult.Min = Math.Round(vesselPress.Min(indexStart, indexEnd), 1)


                ''Average pressure
                indexStart = currentStep.stepStart
                indexEnd = currentStep.stepEnd

                currentStep.pressureResult.Avg = Math.Round(vesselPress.Average(indexStart, indexEnd), 1)




                ''##Pressure ramps##
                goal = -1
                greaterThanGoal = True

                If currentStep.pressureSet.RampRate > 0 Then
                    goal = currentStep.pressureSet.SetPoint
                    greaterThanGoal = True
                ElseIf currentStep.pressureSet.RampRate < 0 Then
                    goal = currentStep.pressureSet.SetPoint
                    greaterThanGoal = False
                End If

                ''Buffer start and end by half the linear regression step length so ramp is only from the step region
                indexStart = currentStep.stepStart + (stepVal / 2)
                indexEnd = currentStep.stepEnd - (stepVal / 2)

                ''For very short step (<linear reg step) just look at the middle point 
                If indexEnd < indexStart Then
                    Dim midPoint As Integer = indexStart + ((indexEnd - indexStart) / 2)
                    indexStart = midPoint
                    indexEnd = midPoint
                End If

                ''Max pressure ramp
                currentStep.pressureResult.MaxRamp = Math.Round(vesselPress.MaxRamp(indexStart, indexEnd, goal, greaterThanGoal), 1)

                ''Min pressure ramp
                currentStep.pressureResult.MinRamp = Math.Round(vesselPress.MinRamp(indexStart, indexEnd, goal, greaterThanGoal), 1)

                ''Avg pressure ramp
                currentStep.pressureResult.AvgRamp = Math.Round(vesselPress.AverageRamp(indexStart, indexEnd, currentStep.pressureSet.SetPoint, currentStep.pressureSet.RampRate, vesselPress.values), 1)



                'Check pressure for passing
                If Math.Abs(currentStep.pressureSet.SetPoint) = Math.Abs(currentStep.pressureSet.NegTol) Then
                    If currentStep.pressureResult.Max > currentStep.pressureSet.SetPoint + currentStep.pressureSet.PosTol Then currentStep.pressurePass = False
                ElseIf Math.Abs(currentStep.pressureSet.SetPoint) = Math.Abs(currentStep.pressureSet.PosTol) Then
                    If currentStep.pressureResult.Min < currentStep.pressureSet.SetPoint + currentStep.pressureSet.NegTol Then currentStep.pressurePass = False
                Else
                    If currentStep.pressureResult.Min < currentStep.pressureSet.SetPoint + currentStep.pressureSet.NegTol Then currentStep.pressurePass = False
                    If currentStep.pressureResult.Max > currentStep.pressureSet.SetPoint + currentStep.pressureSet.PosTol Then currentStep.pressurePass = False
                End If

                'Check pressure ramp if not set to 0
                If currentStep.pressureSet.RampRate <> 0 Then
                    If currentStep.pressureResult.MinRamp < currentStep.pressureSet.RampRate + currentStep.pressureSet.RampNegTol Then currentStep.pressureRampPass = False
                    If currentStep.pressureResult.MaxRamp > currentStep.pressureSet.RampRate + currentStep.pressureSet.RampPosTol Then currentStep.pressureRampPass = False
                End If
            Else
                currentStep.pressurePass = True
                currentStep.pressureRampPass = True
            End If

            'Calculate vac results for a given step
            If curePro.checkVac And currentStep.vacSet.SetPoint <> -1 Then
                indexStart = currentStep.stepStart
                indexEnd = currentStep.stepEnd

                currentStep.vacResult.Max = Math.Round(maxVac.Min(indexStart, indexEnd), 1)
                currentStep.vacResult.Min = Math.Round(minVac.Max(indexStart, indexEnd), 1)

                ''Average vac
                total = 0
                addCnt = 0
                indexStart = currentStep.stepStart
                indexEnd = currentStep.stepEnd

                For z = 0 To UBound(vac_Arr)
                    For y = 0 To UBound(usrRunVac)
                        If vac_Arr(z).Number = usrRunVac(y) Then
                            total = total + vac_Arr(z).Average(indexStart, indexEnd)
                            addCnt = addCnt + 1
                        End If
                    Next
                Next

                currentStep.vacResult.Avg = Math.Round(total / addCnt, 1)

                'Check vacuum for passing
                If Math.Abs(currentStep.vacSet.SetPoint) = Math.Abs(currentStep.vacSet.NegTol) Then
                    If currentStep.vacResult.Min > currentStep.vacSet.SetPoint + currentStep.vacSet.PosTol Then currentStep.vacPass = False
                ElseIf Math.Abs(currentStep.vacSet.SetPoint) = Math.Abs(currentStep.vacSet.PosTol) Then
                    If currentStep.vacResult.Max < currentStep.vacSet.SetPoint + currentStep.vacSet.NegTol Then currentStep.vacPass = False
                Else
                    If currentStep.vacResult.Min > currentStep.vacSet.SetPoint + currentStep.vacSet.PosTol Then currentStep.vacPass = False
                    If currentStep.vacResult.Max < currentStep.vacSet.SetPoint + currentStep.vacSet.NegTol Then currentStep.vacPass = False
                End If

            Else
                currentStep.vacPass = True
            End If


            'Soak total time pass check
            If currentStep.pressureSet.RampRate = 0 And currentStep.tempSet.RampRate = 0 And (currentStep.termCond1Type = "Time" Or currentStep.termCond2Type = "Time") Then

                If Not currentStep.vacPass Or Not currentStep.tempPass Or Not currentStep.pressurePass Then
                    Dim totalTime As Double = 0

                    Dim z As Integer
                    For z = currentStep.stepStart To currentStep.stepEnd
                        If leadTC.values(z) < currentStep.tempSet.SetPoint + currentStep.tempSet.PosTol And lagTC.values(z) > currentStep.tempSet.SetPoint + currentStep.tempSet.NegTol Or currentStep.tempSet.SetPoint = -1 Then
                            If vesselPress.values(z) < currentStep.pressureSet.SetPoint + currentStep.pressureSet.PosTol And vesselPress.values(z) > currentStep.pressureSet.SetPoint + currentStep.pressureSet.NegTol Or currentStep.pressureSet.SetPoint = -1 Then
                                If maxVac.values(z) > currentStep.vacSet.SetPoint + currentStep.vacSet.NegTol And minVac.values(z) < currentStep.vacSet.SetPoint + currentStep.vacSet.PosTol Or currentStep.vacSet.SetPoint = -1 Then
                                    totalTime = totalTime + (dateArr(z) - dateArr(z - 1)).TotalMinutes
                                End If
                            ElseIf vesselPress.values(z) > currentStep.pressureSet.SetPoint + currentStep.pressureSet.PosTol Then
                                totalTime = 0
                                Exit For
                            End If
                        ElseIf leadTC.values(z) > currentStep.tempSet.SetPoint + currentStep.tempSet.PosTol Then
                            totalTime = 0
                            Exit For
                        End If
                    Next

                    'Check if the time in values is at least 90% of the total step time, requirment <10% out of tolerance
                    If (totalTime / currentStep.stepDuration) > 0.9 Then
                        If currentStep.termCond1Type = "Time" Then
                            If totalTime > currentStep.termCond1Goal Then
                                currentStep.soakStepPass = True
                                currentStep.vacPass = True
                                currentStep.tempPass = True
                                currentStep.pressurePass = True
                            End If
                        ElseIf currentStep.termCond2Type = "Time" Then
                            If totalTime > currentStep.termCond2Goal Then
                                currentStep.soakStepPass = True
                                currentStep.vacPass = True
                                currentStep.tempPass = True
                                currentStep.pressurePass = True
                            End If
                        End If
                    End If
                End If
            End If



            'Check for all passing
            If currentStep.vacPass And currentStep.tempPass And currentStep.tempRampPass And currentStep.pressurePass And currentStep.pressureRampPass And currentStep.stepTerminate And currentStep.timeLimitPass Then
                currentStep.stepPass = True
            Else
                currentStep.stepPass = False
                curePro.curePass = False
            End If

        Next
    End Sub

    Sub cureStepTest()
        Dim currentStep As Integer = 0
        Dim i As Integer
        For i = cureStart To dataCnt
            'Set the start of the step
            If curePro.CureSteps(currentStep).stepStart = -1 And currentStep <> 0 Then
                curePro.CureSteps(currentStep).stepStart = i - 1
            ElseIf curePro.CureSteps(currentStep).stepStart = -1 Then
                curePro.CureSteps(currentStep).stepStart = i
            End If

            'Check if you meet the terminating conditions to end the step
            If meetTerms(curePro.CureSteps(currentStep), i) Then
                curePro.CureSteps(currentStep).stepEnd = i - 1


                'If not the last step and the termination type is time check to see if the remaining time needs to be passed to the next step.
                If UBound(curePro.CureSteps) <> currentStep Then
                    If curePro.CureSteps(currentStep).termCond1Type = "Time" Then
                        If curePro.CureSteps(currentStep).termCond1Modifier = "Pass" Then
                            Dim timeToPass As Double = 0
                            timeToPass = curePro.CureSteps(currentStep).termCond1Goal - (dateArr(curePro.CureSteps(currentStep).stepEnd) - dateArr(curePro.CureSteps(currentStep).stepStart)).TotalMinutes
                            If timeToPass < 0 Then timeToPass = 0

                            If curePro.CureSteps(currentStep + 1).termCond1Type = "Time" Then
                                curePro.CureSteps(currentStep + 1).termCond1Goal = timeToPass
                            ElseIf curePro.CureSteps(currentStep + 1).termCond2Type = "Time" Then
                                curePro.CureSteps(currentStep + 1).termCond2Goal = timeToPass
                            Else
                                Throw New Exception("No 'Time' terminating condition to pass time to on step " & curePro.CureSteps(currentStep + 1).stepName)
                            End If
                        End If
                    End If
                    If curePro.CureSteps(currentStep).termCond2Type = "Time" Then
                        If curePro.CureSteps(currentStep).termCond2Modifier = "Pass" Then
                            Dim timeToPass As Double = 0
                            timeToPass = curePro.CureSteps(currentStep).termCond2Goal - (dateArr(curePro.CureSteps(currentStep).stepEnd) - dateArr(curePro.CureSteps(currentStep).stepStart)).TotalMinutes
                            If timeToPass < 0 Then timeToPass = 0

                            If curePro.CureSteps(currentStep + 1).termCond1Type = "Time" Then
                                curePro.CureSteps(currentStep + 1).termCond1Goal = timeToPass
                            ElseIf curePro.CureSteps(currentStep + 1).termCond2Type = "Time" Then
                            curePro.CureSteps(currentStep + 1).termCond2Goal = timeToPass
                            Else
                                Throw New Exception("No 'Time' terminating condition to pass time to on step " & curePro.CureSteps(currentStep + 1).stepName)
                        End If
                    End If
                End If
                End If

                'If the last step has completed then exit the loop and set cure end to the end of this step
                If UBound(curePro.CureSteps) = currentStep Then
                    If cureEnd < curePro.CureSteps(currentStep).stepEnd Then
                        cureEnd = curePro.CureSteps(currentStep).stepEnd
                        cureEndTime = dateArr(curePro.CureSteps(currentStep).stepEnd)
                        currentStep += 1
                        Exit For
                    End If
                End If

                currentStep += 1
            End If
        Next

        'Check to see if the cure was completed and fail any steps that were not reached.
        If UBound(curePro.CureSteps) > currentStep Then
            curePro.CureSteps(currentStep).stepTerminate = False
            curePro.CureSteps(currentStep).stepPass = False
            curePro.curePass = False

            For i = currentStep + 1 To UBound(curePro.CureSteps)
                curePro.CureSteps(i).hardFail = True
            Next
        ElseIf UBound(curePro.CureSteps) = currentStep Then
            curePro.CureSteps(currentStep).stepTerminate = False
            curePro.CureSteps(currentStep).stepPass = False
            curePro.curePass = False
        End If
    End Sub

    Function meetTerms(cureStep As CureStep, currentStep As Integer) As Boolean

        Dim pass1 As Boolean = termTest(cureStep.termCond1Type, cureStep.termCond1Condition, cureStep.termCond1Goal, cureStep.termCond1Modifier, currentStep, cureStep)
        Dim pass2 As Boolean = termTest(cureStep.termCond2Type, cureStep.termCond2Condition, cureStep.termCond2Goal, cureStep.termCond2Modifier, currentStep, cureStep)

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

    Function termTest(termCondType As String, termCondCondition As String, termCondGoal As Double, termCondModifier As Object, currentStep As Integer, cureStep As CureStep) As Boolean
        If termCondType = "None" Then
            Return True

        ElseIf termCondType = "Time" Then
            Dim stepDuration As TimeSpan = dateArr(currentStep) - dateArr(cureStep.stepStart)

            If termCondCondition = "GREATER" Then
                If stepDuration.TotalMinutes > termCondGoal Then
                    Return True
                End If
            Else
                Throw New Exception("Terminating condition Condition for step " & cureStep.stepName & " is not valid")
            End If

        ElseIf termCondType = "Temp" Then

            If termCondCondition = "GREATER" Then
                If termCondModifier = "Lag" Then
                    If lagTC.values(currentStep) > termCondGoal Then
                        Return True
                    End If
                ElseIf termCondModifier = "Lead" Then
                    If leadTC.values(currentStep) > termCondGoal Then
                        Return True
                    End If
                ElseIf termCondModifier = "Air" Then
                    If vessel_TC.values(currentStep) > termCondGoal Then
                        Return True
                    End If
                Else
                    Throw New Exception("Terminating condition Modifier For Step " & cureStep.stepName & " Is Not valid")
                End If

            ElseIf termCondCondition = "LESS" Then
                If termCondModifier = "Lag" Then
                    If lagTC.values(currentStep) < termCondGoal Then
                        Return True
                    End If
                ElseIf termCondModifier = "Lead" Then
                    If leadTC.values(currentStep) < termCondGoal Then
                        Return True
                    End If
                ElseIf termCondModifier = "Air" Then
                    If vessel_TC.values(currentStep) < termCondGoal Then
                        Return True
                    End If
                Else
                    Throw New Exception("Terminating condition Modifier For Step " & cureStep.stepName & " Is Not valid")
                End If
            Else
                Throw New Exception("Terminating condition Condition For Step " & cureStep.stepName & " Is Not valid")
            End If

        ElseIf termCondType = "Press" Then
            If termCondCondition = "GREATER" Then
                If vesselPress.values(currentStep) > termCondGoal Then
                    Return True
                End If
            ElseIf termCondCondition = "LESS" Then

                If vesselPress.values(currentStep) < termCondGoal Then
                    Return True
                End If
            Else
                Throw New Exception("Terminating condition Condition for step " & cureStep.stepName & " is not valid")
            End If

            'ElseIf termCondType") = "Vac" Then
            'Might need this???
        Else
            Throw New Exception("Terminating condition type for step " & cureStep.stepName & " is not valid")
        End If

        Return False
    End Function
#End Region

#Region "Derived arrays from data"
    Sub startEndTime()
        Dim start_end_temp As Double = 150
        Dim start_end_press As Double = 5

        'Get start time
        Dim i As Integer
        For i = 0 To dataCnt
            Dim currentPress As Double = 0

            If curePro.checkPressure Then
                currentPress = vesselPress.values(i)
            Else
                currentPress = 0
            End If

            If leadTC.values(i) > start_end_temp Or currentPress > start_end_press And cureStartTime = Nothing Then
                cureStartTime = dateArr(i)
                cureStart = i
                Exit For
            End If
        Next

        'If the start time wasn't triggered then the cure starts at 0
        If cureStart = 0 Then
            cureStartTime = dateArr(0)
        End If



        'Get end time
        Dim runStart As Boolean = False
        For i = 0 To dataCnt

            Dim currentPress As Double = 0

            If curePro.checkPressure Then
                currentPress = vesselPress.values(i)
            Else
                currentPress = start_end_press + 1
            End If

            If lagTC.values(i) > start_end_temp And currentPress > start_end_press And runStart = False Then
                runStart = True
                i += 5
            End If


            If curePro.checkPressure Then
                currentPress = vesselPress.values(i)
            Else
                currentPress = start_end_press - 1
            End If

            If runStart = True Then
                If leadTC.values(i) < start_end_temp And currentPress < start_end_press Then
                    cureEndTime = dateArr(i)
                    cureEnd = i
                    Exit For
                End If
            End If
        Next

        'If cureEnd didn't get set then the end of data is the end
        If cureEnd = 0 Then
            cureEndTime = dateArr(dataCnt)
            cureEnd = dataCnt
        End If

        If cureEnd <> dataCnt Then
            For i = cureEnd + 1 To dataCnt
                If leadTC.values(i) > start_end_temp Or vesselPress.values(i) > start_end_press Then
                    'MsgBox("Temperature rose above " & start_end_temp & "°F after the cure completed. Check data for deviation.", vbExclamation, "Temperature Error")
                    cureEnd = dataCnt
                    curePro.curePass = False
                    Exit For
                End If
            Next
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
#End Region

#Region "Load in data from csv to given arrays"
    Sub loadCureData()
        Call getTime()

        dataCnt = UBound(dateArr)

        'Calculate ramp rates over a set period for smoothing, numerator sets the number of minutes to look at
        stepVal = 10 / ((dateArr(UBound(dateArr)) - dateArr(0)).TotalMinutes / dataCnt)

        'Make sure step is even so it is evenly distributed around point of interest
        If stepVal Mod 2 = 1 Then
            stepVal -= 1
        End If

        'If stepVal was 0 or 1 then increase to 2 so that a linear regression can be performed
        If stepVal = 0 Then stepVal = 2

        If curePro.checkTemp = True Then
            partTC_Arr.clearArr()
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
            vesselPress.values.clearArr()
            Call getDataSearch("{Vessel_Pressure}", vesselPress)
        End If

        If curePro.checkVac Then
            vac_Arr.clearArr()
            Call getDataMulti("{VacGroup_", vac_Arr, "VAC")
        End If

        If machType = "Autoclave" Then
            vessel_TC.values.clearArr()
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

        loadTo.calcRamp(stepVal, dateArr)
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
        dateArr.clearArr

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

                If dateFnd = True Then Exit Sub
            Next
            If dateFnd = False Then
                Throw New Exception("Failed to find dates in the specified cure document. Check formatting.")
            End If
        End If
    End Sub
#End Region

    Sub errorReset()

        'Clear out the current loaded values in arrays if they exist
        loadedDataSet.clearArr()
        dateArr.clearArr()
        partTC_Arr.clearArr()
        vac_Arr.clearArr()

        vessel_TC = New DataSet(0, "vessel_TC")
        vesselPress = New DataSet(0, "vessel_Press")

        leadTC = New DataSet(0, "leadTC")
        lagTC = New DataSet(0, "lagTC")

        minVac = New DataSet(0, "minVac")
        maxVac = New DataSet(0, "maxVac")

        usrRunTC.clearArr()
        usrRunVac.clearArr()

        'Reset machType to null
        machType = "Unknown"

        equipSerialNum = ""

        dataCnt = 0
        headerRow = 0
        headerCount = 2

        cureStart = 0
        cureEnd = 0

        stepVal = 2

        'Reset partValues to nothing
        JobNum = String.Empty
        PONum = String.Empty
        PartNum = String.Empty
        PartRev = String.Empty
        PartNom = String.Empty
        ProgramNum = String.Empty
        PartQty = String.Empty
        DataPath = String.Empty

        'Reset dateValues to null values
        cureStartTime = Nothing
        cureEndTime = Nothing

        'curePro = Nothing
        'curePro = New CureProfile()


    End Sub

    Sub loadCSVin(inFile As String)

        Call errorReset()

        If IO.File.Exists(inFile) AndAlso IO.Path.GetExtension(inFile).ToLower = ".csv" Then

            Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(inFile)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters({",", vbTab})
                Dim currentRow As String()
                Dim headerSet As Boolean = False

                While Not MyReader.EndOfData
                    Try
                        currentRow = MyReader.ReadFields()

                        'Only check rows with greater than 4 columns
                        If UBound(currentRow) > 3 Then

                            'Stop input for stop line symbol in Omega files
                            If currentRow(3) = "ooOoo" Then
                                Call setHeaderRow()
                                Exit Sub
                            End If

                            'Stop input if first or second item isn't a date/time after 6 lines
                            If Not loadedDataSet Is Nothing AndAlso headerSet = True Then
                                If Not IsDate(currentRow(0)) And Not IsDate(currentRow(1)) Then
                                    Call setHeaderRow()
                                    Exit Sub
                                End If
                            End If

                            loadedDataSet.AddArr(currentRow)


                            If headerSet = False Then
                                If IsDate(currentRow(0)) Then
                                    headerRow = UBound(loadedDataSet, 2)
                                    headerSet = True
                                ElseIf currentRow(0) <> "Date:" Then
                                    If IsDate(currentRow(1)) Then
                                        headerRow = UBound(loadedDataSet, 2)
                                        headerSet = True
                                    End If
                                End If
                            End If

                        End If





                        'Finds file type
                        If InStr(currentRow(0), "Omega", 0) <> 0 Then
                            machType = "Omega"
                            headerCount = 2
                            If curePro.Name = "null" Then
                                curePro.checkPressure = False
                                curePro.checkVac = False
                            ElseIf curePro.checkPressure <> False And curePro.checkVac <> False Then
                                Throw New Exception("Omega TC reader data cannot be used to check this cure profile as is contains either pressure or vacuum requirements.")
                            End If

                        ElseIf currentRow(0) = "No." AndAlso InStr(currentRow(1), "Date", 0) <> 0 AndAlso InStr(currentRow(2), "Time", 0) <> 0 AndAlso InStr(currentRow(3), "Millitm", 0) <> 0 AndAlso InStr(currentRow(4), "{Air_TC}", 0) <> 0 Then
                            machType = "Autoclave"
                            headerCount = 2
                        End If

                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                    End Try
                End While
            End Using

            Call setHeaderRow()
        Else
            Throw New Exception("Input file is not a .csv, please load a new file.")
        End If

    End Sub

    Sub setHeaderRow()
        headerRow = headerRow - headerCount

        If headerRow < 0 Then
            Throw New Exception("Header row cannot be less than 0. Please check to make sure your file has the correct amount of header lines for it's type.")
        End If
    End Sub

End Class
