Imports System.Runtime.CompilerServices

Public Class MainForm

    Public cureProfiles() As CureProfile
    Dim curePro As CureProfile = New CureProfile

    Public partValues As New Dictionary(Of String, String) From {{"JobNum", ""}, {"PONum", ""}, {"PartNum", ""}, {"PartRev", ""}, {"PartNom", ""}, {"ProgramNum", ""}, {"PartQty", ""}, {"DataPath", ""}}

    Public loadedDataSet(,) As String

    Dim dataCnt As Integer = 0
    Dim cureStart As Integer = 0
    Dim cureEnd As Integer = 0

    Dim dateValues As New Dictionary(Of String, DateTime) From {{"startTime", Nothing}, {"endTime", Nothing}}

    Public machType As String = ""

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

        Call loadCureData()


        curePro = cureProfiles(0)

        Call runCalc()

        Call outputResults()

    End Sub

    Sub outputResults()
        Dim Excel As Object
        Excel = CreateObject("Excel.Application")
        Excel.Visible = True
        Excel.workbooks.Add

        Dim currentRow As Integer = 1

        Call outputExcel(Excel, "Cure Name:", "'" & curePro.Name, currentRow)
        Call outputExcel(Excel, "Cure Doc:", curePro.cureDoc, currentRow)
        Call outputExcel(Excel, "Cure Rev", curePro.cureDocRev, currentRow)
        Call outputExcel(Excel, "", "", currentRow)
        Call outputExcel(Excel, "", "", currentRow)

        Dim i As Integer = 0
        For i = 0 To UBound(curePro.CureSteps)
            Dim currentStep As CureStep = curePro.CureSteps(i)

            Call outputExcel(Excel, "Cure Step: ", currentStep.stepName, currentRow)
            Call outputExcel(Excel, "Step Pass: ", currentStep.stepPass, currentRow)
            Call outputExcel(Excel, "", "", currentRow)

            Call outputExcel(Excel, "Temp", "", currentRow)
            Call outputExcel(Excel, "Temp Pass", currentStep.tempPass, currentRow)
            Call outputExcel(Excel, "Temp Max", currentStep.tempResult("Max"), currentRow)
            Call outputExcel(Excel, "Temp Min", currentStep.tempResult("Min"), currentRow)
            Call outputExcel(Excel, "Temp Avg", currentStep.tempResult("Avg"), currentRow)
            Call outputExcel(Excel, "Temp Ramp Max", currentStep.tempResult("MaxRamp"), currentRow)
            Call outputExcel(Excel, "Temp Ramp Min", currentStep.tempResult("MinRamp"), currentRow)
            Call outputExcel(Excel, "Temp Ramp Avg", currentStep.tempResult("AvgRamp"), currentRow)
            Call outputExcel(Excel, "", "", currentRow)

            If machType = "Autoclave" Then

                Call outputExcel(Excel, "Pressure", "", currentRow)
                Call outputExcel(Excel, "Pressure Pass", currentStep.pressurePass, currentRow)
                Call outputExcel(Excel, "Pressure Max", currentStep.pressureResult("Max"), currentRow)
                Call outputExcel(Excel, "Pressure Min", currentStep.pressureResult("Min"), currentRow)
                Call outputExcel(Excel, "Pressure Avg", currentStep.pressureResult("Avg"), currentRow)
                Call outputExcel(Excel, "Pressure Ramp Max", currentStep.pressureResult("MaxRamp"), currentRow)
                Call outputExcel(Excel, "Pressure Ramp Min", currentStep.pressureResult("MinRamp"), currentRow)
                Call outputExcel(Excel, "Pressure Ramp Avg", currentStep.pressureResult("AvgRamp"), currentRow)
                Call outputExcel(Excel, "", "", currentRow)

                Call outputExcel(Excel, "vac", "", currentRow)
                Call outputExcel(Excel, "vac Pass", currentStep.vacPass, currentRow)
                Call outputExcel(Excel, "vac Max", currentStep.vacResult("Max"), currentRow)
                Call outputExcel(Excel, "vac Min", currentStep.vacResult("Min"), currentRow)
                Call outputExcel(Excel, "vac Avg", currentStep.vacResult("Avg"), currentRow)
                Call outputExcel(Excel, "", "", currentRow)

            End If


            Call outputExcel(Excel, "", "", currentRow)
        Next

        currentRow = 6
        For i = 0 To UBound(curePro.CureSteps)
            Dim currentStep As CureStep = curePro.CureSteps(i)

            Call outputExcelC2(Excel, "", "", currentRow)
            Call outputExcelC2(Excel, "", "", currentRow)
            Call outputExcelC2(Excel, "", "", currentRow)

            Call outputExcelC2(Excel, "", "", currentRow)
            Call outputExcelC2(Excel, "", "", currentRow)
            Call outputExcelC2(Excel, "Setpoint", currentStep.tempSet("SetPoint"), currentRow)
            Call outputExcelC2(Excel, "PosTol", currentStep.tempSet("PosTol"), currentRow)
            Call outputExcelC2(Excel, "NegTol", currentStep.tempSet("NegTol"), currentRow)
            Call outputExcelC2(Excel, "RampSet", currentStep.tempSet("RampRate"), currentRow)
            Call outputExcelC2(Excel, "RampPosTol", currentStep.tempSet("RampPosTol"), currentRow)
            Call outputExcelC2(Excel, "RampNegTol", currentStep.tempSet("RampNegTol"), currentRow)
            Call outputExcelC2(Excel, "", "", currentRow)

            If machType = "Autoclave" Then

                Call outputExcelC2(Excel, "", "", currentRow)

                Call outputExcelC2(Excel, "Setpoint", currentStep.pressureSet("SetPoint"), currentRow)
                Call outputExcelC2(Excel, "PosTol", currentStep.pressureSet("PosTol"), currentRow)
                Call outputExcelC2(Excel, "NegTol", currentStep.pressureSet("NegTol"), currentRow)
                Call outputExcelC2(Excel, "RampSet", currentStep.pressureSet("RampRate"), currentRow)
                Call outputExcelC2(Excel, "RampPosTol", currentStep.pressureSet("RampPosTol"), currentRow)
                Call outputExcelC2(Excel, "RampNegTol", currentStep.pressureSet("RampNegTol"), currentRow)
                Call outputExcelC2(Excel, "", "", currentRow)
                Call outputExcelC2(Excel, "", "", currentRow)

                Call outputExcelC2(Excel, "", "", currentRow)
                Call outputExcelC2(Excel, "", "", currentRow)
                Call outputExcelC2(Excel, "Setpoint", currentStep.vacSet("SetPoint"), currentRow)
                Call outputExcelC2(Excel, "PosTol", currentStep.vacSet("PosTol"), currentRow)
                Call outputExcelC2(Excel, "NegTol", currentStep.vacSet("NegTol"), currentRow)
                Call outputExcelC2(Excel, "RampSet", currentStep.vacSet("RampRate"), currentRow)
                Call outputExcelC2(Excel, "", "", currentRow)

            End If


            Call outputExcelC2(Excel, "", "", currentRow)
        Next
    End Sub

    Sub outputExcel(Excel As Object, desc As String, uVal As String, ByRef currentRow As Integer)
        Excel.cells(currentRow, 1) = desc
        Excel.cells(currentRow, 2) = uVal
        currentRow = currentRow + 1
    End Sub

    Sub outputExcelC2(Excel As Object, desc As String, uVal As String, ByRef currentRow As Integer)
        Excel.cells(currentRow, 4) = desc
        Excel.cells(currentRow, 5) = uVal
        currentRow = currentRow + 1
    End Sub

    Sub runCalc()
        Call leadlagTC()

        If machType = "Autoclave" Then
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

            'Calculate temp results for a given step

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
            Dim total As Double = 0
            Dim addCnt As Integer = 0
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


            'Check autoclave only features
            If machType = "Autoclave" Then
                'Calculate pressure results for a given step

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



                'Calculate vac results for a given step
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

                'Check for all passing
                If curePro.CureSteps(i).vacPass And curePro.CureSteps(i).tempPass And curePro.CureSteps(i).pressurePass Then
                    curePro.CureSteps(i).stepPass = True
                Else
                    curePro.CureSteps(i).stepPass = False
                    curePro.curePass = False
                End If

            ElseIf machType = "Omega" Then
                If curePro.CureSteps(i).tempPass Then
                    curePro.CureSteps(i).stepPass = True
                Else
                    curePro.CureSteps(i).stepPass = False
                    curePro.curePass = False
                End If
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

        Dim pass1 As Boolean = False
        Dim pass2 As Boolean = False


        'Test term cond #1
        If cureStep.termCond1("Type") = "None" Then
            pass1 = True

        ElseIf cureStep.termCond1("Type") = "Time" Then
            Dim stepDuration As TimeSpan = dateArr(currentStep) - dateArr(cureStep.stepStart)

            If cureStep.termCond1("Condition") = "GREATER" Then
                If stepDuration.TotalMinutes > cureStep.termCond1("Goal") Then
                    pass1 = True
                End If
            End If

        ElseIf cureStep.termCond1("Type") = "Temp" Then

            If cureStep.termCond1("Condition") = "GREATER" Then
                If cureStep.termCond1("TCNum") = "Lag" Then
                    Dim test = lagTC.values(currentStep)
                    If lagTC.values(currentStep) > cureStep.termCond1("Goal") Then
                        pass1 = True
                    End If
                End If

                If cureStep.termCond1("TCNum") = "Lead" Then
                    If lagTC.values(currentStep) > cureStep.termCond1("Goal") Then
                        pass1 = True
                    End If
                End If

                If cureStep.termCond1("TCNum") = "Air" Then
                    If vessel_TC.values(currentStep) > cureStep.termCond1("Goal") Then
                        pass1 = True
                    End If
                End If

            ElseIf cureStep.termCond1("Condition") = "LESS" Then
                If cureStep.termCond1("TCNum") = "Lag" Then
                    If lagTC.values(currentStep) < cureStep.termCond1("Goal") Then
                        pass1 = True
                    End If
                End If

                If cureStep.termCond1("TCNum") = "Lead" Then
                    If lagTC.values(currentStep) < cureStep.termCond1("Goal") Then
                        pass1 = True
                    End If
                End If

                If cureStep.termCond1("TCNum") = "Air" Then
                    If vessel_TC.values(currentStep) < cureStep.termCond1("Goal") Then
                        pass1 = True
                    End If
                End If
            End If

        ElseIf cureStep.termCond1("Type") = "Press" Then
            If cureStep.termCond1("Condition") = "GREATER" Then
                If vesselPress.values(currentStep) > cureStep.termCond1("Goal") Then
                    pass1 = True
                End If
            ElseIf cureStep.termCond1("Condition") = "LESS" Then

                If vesselPress.values(currentStep) < cureStep.termCond1("Goal") Then
                    pass1 = True
                End If
            End If

        ElseIf cureStep.termCond1("Type") = "Vac" Then
            'Might need this???
        End If





        'Test term cond #2
        If cureStep.termCond2("Type") = "None" Then
            pass2 = True

        ElseIf cureStep.termCond2("Type") = "Time" Then
            Dim stepDuration As TimeSpan = dateArr(currentStep) - dateArr(cureStep.stepStart)

            If cureStep.termCond2("Condition") = "GREATER" Then
                If stepDuration.TotalMinutes > cureStep.termCond2("Goal") Then
                    pass2 = True
                End If
            End If

        ElseIf cureStep.termCond2("Type") = "Temp" Then

            If cureStep.termCond2("Condition") = "GREATER" Then
                If cureStep.termCond2("TCNum") = "Lag" Then
                    If lagTC.values(currentStep) > cureStep.termCond2("Goal") Then
                        pass2 = True
                    End If
                End If

                If cureStep.termCond2("TCNum") = "Lead" Then
                    If lagTC.values(currentStep) > cureStep.termCond2("Goal") Then
                        pass2 = True
                    End If
                End If

                If cureStep.termCond2("TCNum") = "Air" Then
                    If vessel_TC.values(currentStep) > cureStep.termCond2("Goal") Then
                        pass2 = True
                    End If
                End If

            ElseIf cureStep.termCond2("Condition") = "LESS" Then
                If cureStep.termCond2("TCNum") = "Lag" Then
                    If lagTC.values(currentStep) < cureStep.termCond2("Goal") Then
                        pass2 = True
                    End If
                End If

                If cureStep.termCond2("TCNum") = "Lead" Then
                    If lagTC.values(currentStep) < cureStep.termCond2("Goal") Then
                        pass2 = True
                    End If
                End If

                If cureStep.termCond2("TCNum") = "Air" Then
                    If vessel_TC.values(currentStep) < cureStep.termCond2("Goal") Then
                        pass2 = True
                    End If
                End If
            End If

        ElseIf cureStep.termCond2("Type") = "Press" Then
            If cureStep.termCond2("Condition") = "GREATER" Then
                If vesselPress.values(currentStep) > cureStep.termCond2("Goal") Then
                    pass2 = True
                End If
            ElseIf cureStep.termCond2("Condition") = "LESS" Then

                If vesselPress.values(currentStep) < cureStep.termCond2("Goal") Then
                    pass2 = True
                End If
            End If

        ElseIf cureStep.termCond2("Type") = "Vac" Then
            'Might need this???
        End If


        If cureStep.termCondOper = "OR" Then
            If pass1 Or pass2 Then
                Return True
            End If

        ElseIf cureStep.termCondOper = "AND" Then
            If pass1 And pass2 Then
                Return True
            End If
        End If

        Return False
    End Function


    Sub startEndTime()
        Dim i As Integer
        For i = 0 To dataCnt
            If leadTC.values(i) > 140 And dateValues("startTime") = Nothing Then
                dateValues("startTime") = dateArr(i)
                cureStart = i
                Exit For
            End If
        Next

        Dim runStart As Boolean = False
        For i = 0 To dataCnt
            If lagTC.values(i) > 140 And runStart = False Then
                runStart = True
                i = i + 5
            End If

            If runStart = True Then
                If lagTC.values(i) < 140 Then
                    dateValues("endTime") = dateArr(i)
                    cureEnd = i
                    Exit For
                End If
            End If
        Next
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

        Call getTC(stepVal)

        If machType = "Autoclave" Then
            Call getVac()
            Call getPress(stepVal)
            Call getvessTC(stepVal)
        End If
    End Sub

    Sub getPress(stepRate As Integer)
        For i = 0 To UBound(loadedDataSet, 1)
            Dim searchVal As String = "{Vessel_Pressure}"
            Dim searchRow As Integer = 0

            If InStr(loadedDataSet(i, searchRow), searchVal, 0) <> 0 Then

                Dim row As Integer
                For row = searchRow + 2 To UBound(loadedDataSet, 2)
                    If IsNumeric(loadedDataSet(i, row)) Then
                        vesselPress.values.AddValArr(loadedDataSet(i, row))
                    Else
                        vesselPress.values.AddValArr(0)
                    End If
                Next

                vesselPress.calcRamp(stepRate)
            End If
        Next
    End Sub

    Sub getvessTC(stepRate As Integer)
        For i = 0 To UBound(loadedDataSet, 1)
            Dim searchVal As String = "{Air_TC}"
            Dim searchRow As Integer = 0

            If InStr(loadedDataSet(i, searchRow), searchVal, 0) <> 0 Then

                Dim row As Integer
                For row = searchRow + 2 To UBound(loadedDataSet, 2)
                    If IsNumeric(loadedDataSet(i, row)) Then
                        vessel_TC.values.AddValArr(loadedDataSet(i, row))
                    Else
                        vessel_TC.values.AddValArr(0)
                    End If
                Next

                vessel_TC.calcRamp(stepRate)
            End If
        Next
    End Sub

    Sub getTC(stepRate As Integer)
        Dim i As Integer
        Dim searchVal As String = ""
        Dim searchRow As Integer = 0

        If machType = "Omega" Then
            searchVal = "Channel "
            searchRow = 2
        ElseIf machType = "Autoclave" Then
            searchVal = "{Part_"
            searchRow = 0
        End If

        For i = 0 To UBound(loadedDataSet, 1)
            If InStr(loadedDataSet(i, searchRow), searchVal, 0) <> 0 Then

                If TC_Used(i, searchRow + 2) Then
                    If partTC_Arr Is Nothing Then
                        ReDim partTC_Arr(0)
                        partTC_Arr(0) = New DataSet(0, "")
                    Else
                        ReDim Preserve partTC_Arr(UBound(partTC_Arr) + 1)
                        partTC_Arr(UBound(partTC_Arr)) = New DataSet(0, "")
                    End If


                    partTC_Arr(UBound(partTC_Arr)).Number = Integer.Parse(System.Text.RegularExpressions.Regex.Replace(loadedDataSet(i, searchRow), "[^\d]", ""))
                    partTC_Arr(UBound(partTC_Arr)).Type = "TC"

                    Dim row As Integer
                    For row = searchRow + 2 To UBound(loadedDataSet, 2)
                        If IsNumeric(loadedDataSet(i, row)) Then
                            partTC_Arr(UBound(partTC_Arr)).values.AddValArr(loadedDataSet(i, row))
                        Else
                            partTC_Arr(UBound(partTC_Arr)).values.AddValArr(0)
                        End If
                    Next

                    partTC_Arr(UBound(partTC_Arr)).calcRamp(stepRate)
                End If

            End If
        Next
    End Sub

    Function TC_Used(chkCol As Integer, chkRowStart As Integer) As Boolean
        Dim r As Integer
        For r = chkRowStart To UBound(loadedDataSet, 2)
            If IsNumeric(loadedDataSet(chkCol, r)) AndAlso loadedDataSet(chkCol, r) <> 0 Then
                Return True
            End If
        Next

        Return False
    End Function

    Sub getVac()
        Dim i As Integer
        Dim searchVal As String = "{VacGroup_"
        Dim searchRow As Integer = 0


        For i = 0 To UBound(loadedDataSet, 1)
            If InStr(loadedDataSet(i, searchRow), searchVal, 0) <> 0 Then

                If vac_Used(i, searchRow + 2) Then
                    If vac_Arr Is Nothing Then
                        ReDim vac_Arr(0)
                        vac_Arr(0) = New DataSet(0, "")
                    Else
                        ReDim Preserve vac_Arr(UBound(vac_Arr) + 1)
                        vac_Arr(UBound(vac_Arr)) = New DataSet(0, "")
                    End If


                    vac_Arr(UBound(vac_Arr)).Number = Integer.Parse(System.Text.RegularExpressions.Regex.Replace(loadedDataSet(i, searchRow), "[^\d]", ""))
                    vac_Arr(UBound(vac_Arr)).Type = "VAC"

                    Dim row As Integer
                    For row = searchRow + 2 To UBound(loadedDataSet, 2)
                        If IsNumeric(loadedDataSet(i, row)) Then
                            vac_Arr(UBound(vac_Arr)).values.AddValArr(loadedDataSet(i, row))
                        Else
                            vac_Arr(UBound(vac_Arr)).values.AddValArr(0)
                        End If

                    Next

                End If

            End If
        Next
    End Sub

    Function vac_Used(chkCol As Integer, chkRowStart As Integer) As Boolean
        Dim r As Integer
        For r = chkRowStart To UBound(loadedDataSet, 2)
            Dim test1 = loadedDataSet(chkCol, r)

            If IsNumeric(loadedDataSet(chkCol, r)) AndAlso loadedDataSet(chkCol, r) <> 0 Then
                Return True
            End If
        Next

        Return False

    End Function

    Sub getTime()
        Dim i As Integer

        If machType = "Omega" Then
            For i = 4 To UBound(loadedDataSet, 2)
                dateArr.AddValArr(Convert.ToDateTime(loadedDataSet(0, i)))
            Next

        ElseIf machType = "Autoclave" Then
            For i = 2 To UBound(loadedDataSet, 2)
                dateArr.AddValArr(Convert.ToDateTime(loadedDataSet(1, i) & " " & loadedDataSet(2, i) & "." & loadedDataSet(3, i)))
            Next
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
                        ElseIf currentRow(0) = "No." AndAlso InStr(currentRow(1), "Date", 0) <> 0 AndAlso InStr(currentRow(2), "Time", 0) <> 0 AndAlso InStr(currentRow(3), "Millitm", 0) <> 0 AndAlso InStr(currentRow(4), "{Air_TC}", 0) <> 0 Then
                            machType = "Autoclave"
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

