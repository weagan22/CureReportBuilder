Option Explicit On

Imports System.Runtime.CompilerServices
Imports Microsoft.Office.Interop

Public Class MainForm

    Public cureProfiles() As CureProfile
    Dim curePro As CureProfile = New CureProfile("null")

    Public partValues As New Dictionary(Of String, String) From {{"JobNum", ""}, {"PONum", ""}, {"PartNum", ""}, {"PartRev", ""}, {"PartNom", ""}, {"ProgramNum", ""}, {"PartQty", ""}, {"DataPath", ""}}
    Dim equipSerialNum As String = ""

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

    Dim usrRunTC() As Integer
    Dim usrRunVac() As Integer
    Public Property settings As Object

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Call errorReset()



        Txt_Technician.Text = Replace(Environment.UserName, ".", " ")
        Txt_CureProfilesPath.Text = My.Settings.CureProfilePath
        Txt_TemplatePath.Text = My.Settings.ReportTemplatePath

        Try
            Call loadCureProfiles(Txt_CureProfilesPath.Text)
        Catch ex As Exception
            If MsgBox(ex.Message, vbOKCancel, "Cure Profiles: Path Error") = vbCancel Then Me.Close()
        End Try

        Me.Show()


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Btn_Run.Click
        testRun()
    End Sub

    Sub testRun()

        partValues("JobNum") = Txt_JobNumber.Text
        partValues("PONum") = ""
        partValues("PartNum") = Txt_PartNumber.Text
        partValues("PartRev") = Txt_Revision.Text
        partValues("PartNom") = Txt_PartDesc.Text
        partValues("ProgramNum") = Txt_ProgramNumber.Text
        partValues("PartQty") = Txt_Qty.Text
        equipSerialNum = Txt_DataRecorder.Text


        'Get used TC lines
        If curePro.checkTemp Then
            usrRunTC.clearArr()

            For i = 0 To Data_TC.Rows.Count - 1
                If Data_TC.Item(0, i).Value Then
                    usrRunTC.AddValArr(Data_TC.Item(1, i).Value)
                End If
            Next

            If usrRunTC Is Nothing Then
                If MsgBox("No TC's were selected, are you sure you would like to run with all of them?", vbOKCancel, "TC Check") = vbCancel Then Exit Sub
                For i = 0 To Data_TC.Rows.Count - 1
                    usrRunTC.AddValArr(Data_TC.Item(1, i).Value)
                Next
            End If
        End If


        'Get used vac ports
        If curePro.checkVac Then
            usrRunVac.clearArr()

            For i = 0 To Data_Vac.Rows.Count - 1
                If Data_Vac.Item(0, i).Value Then
                    usrRunVac.AddValArr(Data_Vac.Item(1, i).Value)
                End If
            Next

            If usrRunVac Is Nothing Then
                If MsgBox("No Vac ports were selected, are you sure you would like to run with all of them?", vbOKCancel, "Vac Check") = vbCancel Then Exit Sub
                For i = 0 To Data_Vac.Rows.Count - 1
                    usrRunVac.AddValArr(Data_Vac.Item(1, i).Value)
                Next
            End If
        End If


        Call runCalc()

        Call outputResults()
    End Sub
#Region "Output"
    Sub formatFont(inCell As Object, checkStr As String, Optional fontSize As Double = 11, Optional bold As Boolean = False, Optional italic As Boolean = False, Optional underline As Boolean = False, Optional color As Color = Nothing, Optional checkExists As Boolean = True)
        'Exit sub to skip formatting
        'Exit Sub

        If checkStr = "" Then Exit Sub

        If Strings.Right(checkStr, 1) = vbLf Then
            checkStr = Strings.Left(checkStr, Len(checkStr) - 1)
        End If

        Dim startPos As Integer = InStr(inCell.value, checkStr)

        If checkExists Then
            If startPos = 0 Then
                Throw New Exception("Format font failed to locate the desired string within the specified cell.")
            End If
        ElseIf startPos = 0 Then
            Exit Sub
        End If

        If fontSize <> 11 Then
            inCell.Characters(Start:=startPos, Length:=Len(checkStr)).Font.Size = fontSize
        End If

        If bold Then
            inCell.Characters(Start:=startPos, Length:=Len(checkStr)).Font.Bold = bold
        End If

        If italic Then
            inCell.Characters(Start:=startPos, Length:=Len(checkStr)).Font.Italic = italic
        End If

        If underline Then
            inCell.Characters(Start:=startPos, Length:=Len(checkStr)).Font.Underline = underline
        End If

        If color <> Nothing Then
            inCell.Characters(Start:=startPos, Length:=Len(checkStr)).Font.Color = color
        End If

    End Sub

    Sub SwitchOff(app As Excel.Application, bSwitchOff As Boolean)
        Dim ws As Excel.Worksheet

        With app
            If bSwitchOff Then
                ' OFF 
                .Calculation = -4135
                .ScreenUpdating = False
                .EnableAnimations = False

                For Each ws In app.ActiveWorkbook.Worksheets
                    ws.DisplayPageBreaks = False
                Next ws
            Else
                ' ON
                .Calculation = -4105
                .ScreenUpdating = True
                .EnableAnimations = True

                For Each ws In app.ActiveWorkbook.Worksheets
                    ws.DisplayPageBreaks = True
                Next ws
            End If
        End With
    End Sub

    Sub outputResults()
        Dim Excel As Object
        Excel = CreateObject("Excel.Application")
        Excel.Visible = False

        Try
            Excel.Workbooks.Open(Txt_TemplatePath.Text)
        Catch ex As Exception
            MsgBox("Failed to open Excel Report Template.")
            Exit Sub
        End Try

        SwitchOff(Excel, True)

        Dim mainSheet As Excel.Worksheet = Excel.Sheets.Item(1)
        Dim dataSheet As Excel.Worksheet = Excel.Sheets.Item(3)

        mainSheet.Cells(2, 1) = "Job" & vbNewLine & partValues("JobNum") & vbNewLine & "Program" & vbNewLine & partValues("ProgramNum")
        formatFont(mainSheet.Cells(2, 1), "Job", 14, True, False, True)
        formatFont(mainSheet.Cells(2, 1), "Program", 14, True, False, True)

        mainSheet.Cells(2, 4) = "Part" & vbNewLine & partValues("PartNum") & vbNewLine & "Rev. " & partValues("PartRev") & vbNewLine & partValues("PartNom") & vbNewLine & "Qty: " & partValues("PartQty")
        formatFont(mainSheet.Cells(2, 4), "Part", 14, True, False, True)


        Dim hours As Integer = Math.Round((dateArr(cureEnd) - dateArr(cureStart)).TotalMinutes, 1) \ 60
        Dim minutes As Integer = Math.Round((dateArr(cureEnd) - dateArr(cureStart)).TotalMinutes, 1) - (hours * 60)

        mainSheet.Cells(2, 8) = "Cure" & vbNewLine & "Start: " & dateArr(cureStart).ToString("dd-MMM-yyyy | h:mm tt") & vbNewLine & "Finish: " & dateArr(cureEnd).ToString("dd-MMM-yyyy | h:mm tt") & vbNewLine & "Duration: " & hours & ":" & minutes 'Math.Round((dateArr(cureEnd) - dateArr(cureStart)).TotalMinutes, 1) & " min"
        formatFont(mainSheet.Cells(2, 8), "Cure", 14, True, False, True)

        mainSheet.Cells(5, 1) = "Equipment" & vbNewLine & machType & " | " & equipSerialNum
        formatFont(mainSheet.Cells(5, 1), "Equipment", 14, True, False, True)

        If curePro.curePass Then
            mainSheet.Cells(6, 4) = "PASS"
            formatFont(mainSheet.Cells(6, 4), "PASS", 30, True, False, False, Color.Green)
        Else
            mainSheet.Cells(6, 4) = "FAIL"
            formatFont(mainSheet.Cells(6, 4), "FAIL", 30, True, False, True, Color.Red)
        End If

        mainSheet.Cells(5, 8) = "Cure Document" & vbNewLine & curePro.cureDoc & vbNewLine & "Rev. " & curePro.cureDocRev & vbNewLine & "Profile: " & curePro.Name
        formatFont(mainSheet.Cells(5, 8), "Cure Document", 14, True, False, True)

        mainSheet.Cells(34, 3) = Txt_Technician.Text
        mainSheet.Cells(34, 9) = Today.Date().ToString("MM-dd-yyyy")


        Dim curRow As Integer = 40
        For i = 0 To UBound(curePro.CureSteps)

            Dim currentStep As CureStep = curePro.CureSteps(i)

            'Merge cells and center text
            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 1), mainSheet.Cells(curRow, 2)).Merge
            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 3), mainSheet.Cells(curRow, 6)).Merge
            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 7), mainSheet.Cells(curRow, 9)).Merge

            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 1), mainSheet.Cells(curRow, 2)).HorizontalAlignment = 3
            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 3), mainSheet.Cells(curRow, 6)).HorizontalAlignment = 3
            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 7), mainSheet.Cells(curRow, 9)).HorizontalAlignment = 3

            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 1), mainSheet.Cells(curRow, 9)).VerticalAlignment = -4108

            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 1), mainSheet.Cells(curRow, 10)).Borders(9).LineStyle = 1

            'Fill out data in cell 1
            If currentStep.hardFail Then
                mainSheet.Cells(curRow, 1) = currentStep.stepName
                formatFont(mainSheet.Cells(curRow, 1), currentStep.stepName, 14, True, False, False)
            Else
                Dim stepPass As String = "FAIL"
                If currentStep.stepPass Then stepPass = "PASS"

                Dim startTime As Integer = (dateArr(currentStep.stepStart) - dateArr(cureStart)).TotalMinutes
                Dim endTime As Integer = (dateArr(currentStep.stepEnd) - dateArr(cureStart)).TotalMinutes

                Dim endStr As String = ""
                If currentStep.stepTerminate = False Then
                    endStr = "Failed to Terminate"
                Else
                    endStr = "Finish: " & endTime & " min"
                End If


                mainSheet.Cells(curRow, 1) = currentStep.stepName & vbNewLine & "(" & stepPass & ")" & vbNewLine & "Start: " & startTime & " min" & vbNewLine & endStr

                formatFont(mainSheet.Cells(curRow, 1), currentStep.stepName, 14, True, False, False)

                If currentStep.stepPass = True Then
                    formatFont(mainSheet.Cells(curRow, 1), "(" & stepPass & ")", 11, False, False, False)
                Else
                    formatFont(mainSheet.Cells(curRow, 1), "(" & stepPass & ")", 11, False, False, False, Color.Red)
                End If

                formatFont(mainSheet.Cells(curRow, 1), "Start: " & startTime & " min", 9, False, False, False)
                formatFont(mainSheet.Cells(curRow, 1), endStr, 9, False, False, False)

                If currentStep.stepTerminate = False Then
                    formatFont(mainSheet.Cells(curRow, 1), endStr, 9, False, False, False, Color.Red)
                End If
            End If

            'Fill out cell 2
            Dim cell2 As String = ""

            Dim tempStr As String = ""
            Dim tempRmpStr As String = ""
            Dim pressStr As String = ""
            Dim pressRmpStr As String = ""
            Dim vacStr As String = ""

            Dim optionNeg As String = ""

            If curePro.checkTemp Then
                If Math.Abs(currentStep.tempSet("SetPoint")) = Math.Abs(currentStep.tempSet("NegTol")) Then
                    tempStr = "Temperature: Max " & currentStep.tempSet("SetPoint") + currentStep.tempSet("PosTol") & "°F"
                ElseIf Math.Abs(currentStep.tempSet("SetPoint")) = Math.Abs(currentStep.tempSet("PosTol")) Then
                    tempStr = "Temperature: Min " & currentStep.tempSet("SetPoint") + currentStep.tempSet("NegTol") & "°F"
                Else
                    optionNegTest(currentStep.tempSet("NegTol"), optionNeg)
                    tempStr = "Temperature: " & currentStep.tempSet("SetPoint") & "°F +" & currentStep.tempSet("PosTol") & "/" & optionNeg & currentStep.tempSet("NegTol") & "°F"
                End If
                tempStr = tempStr & vbNewLine

                If currentStep.tempSet("RampRate") <> 0 Then
                    optionNegTest(currentStep.tempSet("RampNegTol"), optionNeg)
                    tempRmpStr = "Temp. Ramp: " & currentStep.tempSet("RampRate") & "°F/min +" & currentStep.tempSet("RampPosTol") & "/" & optionNeg & currentStep.tempSet("RampNegTol") & "°F/min"
                    tempRmpStr = tempRmpStr & vbNewLine
                End If
            End If

            If curePro.checkPressure Then
                If Math.Abs(currentStep.pressureSet("SetPoint")) = Math.Abs(currentStep.pressureSet("NegTol")) Then
                    pressStr = "Pressure: Max " & currentStep.pressureSet("SetPoint") + currentStep.pressureSet("PosTol") & " psi"
                ElseIf Math.Abs(currentStep.pressureSet("SetPoint")) = Math.Abs(currentStep.pressureSet("PosTol")) Then
                    pressStr = "Pressure Min " & currentStep.pressureSet("SetPoint") + currentStep.pressureSet("NegTol") & " psi"
                Else
                    optionNegTest(currentStep.pressureSet("NegTol"), optionNeg)
                    pressStr = pressStr & "Pressure: " & currentStep.pressureSet("SetPoint") & " psi +" & currentStep.pressureSet("PosTol") & "/" & optionNeg & currentStep.pressureSet("NegTol") & " psi"
                End If
                pressStr = pressStr & vbNewLine

                If currentStep.pressureSet("RampRate") <> 0 Then
                    optionNegTest(currentStep.pressureSet("RampNegTol"), optionNeg)
                    pressRmpStr = pressRmpStr & "Pressure Ramp: " & currentStep.pressureSet("RampRate") & " psi/min +" & currentStep.pressureSet("RampPosTol") & "/" & optionNeg & currentStep.pressureSet("RampNegTol") & " psi/min"
                    pressRmpStr = pressRmpStr & vbNewLine
                End If
            End If

            If curePro.checkVac Then
                If Math.Abs(currentStep.vacSet("SetPoint")) = Math.Abs(currentStep.vacSet("NegTol")) Then
                    vacStr = "Pressure: Max " & currentStep.vacSet("SetPoint") + currentStep.vacSet("PosTol") & " inHg"
                ElseIf Math.Abs(currentStep.vacSet("SetPoint")) = Math.Abs(currentStep.vacSet("PosTol")) Then
                    vacStr = "Pressure: Min " & currentStep.vacSet("SetPoint") + currentStep.vacSet("NegTol") & " inHg"
                Else
                    optionNegTest(currentStep.vacSet("NegTol"), optionNeg)
                    vacStr = vacStr & "Vacuum: " & currentStep.vacSet("SetPoint") & " inHg +" & currentStep.vacSet("PosTol") & "/" & optionNeg & currentStep.vacSet("NegTol") & " inHg"
                End If
                vacStr = vacStr & vbNewLine
            End If

            cell2 = cell2 & tempStr & tempRmpStr & pressStr & pressRmpStr & vacStr

            cell2 = cell2 & "Terminate"
            cell2 = cell2 & vbNewLine



            Dim term1Str As String = termToStr(currentStep.termCond1, currentStep)
            Dim term2Str As String = ""

            If currentStep.termCond2("Type") <> "None" Then
                term2Str = term2Str & currentStep.termCondOper
                term2Str = term2Str & vbNewLine
                term2Str = term2Str & termToStr(currentStep.termCond2, currentStep)
            End If

            cell2 = cell2 & term1Str
            cell2 = cell2 & term2Str

            If Strings.Right(cell2, 1) = vbLf Then
                cell2 = Strings.Left(cell2, Len(cell2) - 1)
            End If

            mainSheet.Cells(curRow, 3) = cell2

            'Format cell 2

            formatFont(mainSheet.Cells(curRow, 3), tempStr, 9,, True)
            formatFont(mainSheet.Cells(curRow, 3), tempRmpStr, 9,, True)
            formatFont(mainSheet.Cells(curRow, 3), pressStr, 9,, True)
            formatFont(mainSheet.Cells(curRow, 3), pressRmpStr, 9,, True)
            formatFont(mainSheet.Cells(curRow, 3), vacStr, 9,, True)

            If Strings.Right(term1Str, 1) = vbLf Then
                term1Str = Strings.Left(term1Str, Len(term1Str) - 1)
            End If
            If Strings.Right(term2Str, 1) = vbLf Then
                term2Str = Strings.Left(term2Str, Len(term2Str) - 1)
            End If

            formatFont(mainSheet.Cells(curRow, 3), "Terminate", 9, True, True, True)
            formatFont(mainSheet.Cells(curRow, 3), term1Str, 9,, True)
            formatFont(mainSheet.Cells(curRow, 3), term2Str, 9,, True)


            'Fill out cell 3
            If currentStep.hardFail Then
                mainSheet.Cells(curRow, 7) = "Failed to Reach Step"
                formatFont(mainSheet.Cells(curRow, 7), "Failed to Reach Step", 11, True,, True, Color.Red)
            Else
                Dim cell3 As String = ""

                tempStr = ""
                tempRmpStr = ""
                pressStr = ""
                pressRmpStr = ""
                vacStr = ""

                If curePro.checkTemp Then
                    tempStr = "Temperature (°F)" & vbNewLine

                    If Math.Abs(currentStep.tempSet("SetPoint")) = Math.Abs(currentStep.tempSet("NegTol")) Then
                        tempStr = tempStr & "Max: " & String.Format("{0}", currentStep.tempResult("Max")) & vbNewLine
                    ElseIf Math.Abs(currentStep.tempSet("SetPoint")) = Math.Abs(currentStep.tempSet("PosTol")) Then
                        tempStr = tempStr & "Min: " & String.Format("{0}", currentStep.tempResult("Min")) & vbNewLine
                    Else
                        tempStr = tempStr & "Max: " & String.Format("{0}", currentStep.tempResult("Max")) & " | Min: " & String.Format("{0}", currentStep.tempResult("Min")) & "  | Avg: " & String.Format("{0}", currentStep.tempResult("Avg")) & vbNewLine
                    End If

                    If currentStep.tempSet("RampRate") <> 0 Then
                        tempRmpStr = "Temp. Ramp (°F/min)" & vbNewLine
                        tempRmpStr = tempRmpStr & "Max: " & String.Format("{0:0.0}", currentStep.tempResult("MaxRamp")) & " | Min: " & String.Format("{0:0.0}", currentStep.tempResult("MinRamp")) & "  | Avg: " & String.Format("{0:0.0}", currentStep.tempResult("AvgRamp")) & vbNewLine
                    End If
                End If

                If curePro.checkPressure Then
                    pressStr = "Pressure (psi)" & vbNewLine

                    If Math.Abs(currentStep.pressureSet("SetPoint")) = Math.Abs(currentStep.pressureSet("NegTol")) Then
                        pressStr = pressStr & "Max: " & String.Format("{0:0.0}", currentStep.pressureResult("Max")) & vbNewLine
                    ElseIf Math.Abs(currentStep.pressureSet("SetPoint")) = Math.Abs(currentStep.pressureSet("PosTol")) Then
                        pressStr = "Min: " & String.Format("{0:0.0}", currentStep.pressureResult("Min")) & vbNewLine
                    Else
                        pressStr = pressStr & "Max: " & String.Format("{0:0.0}", currentStep.pressureResult("Max")) & " | Min: " & String.Format("{0:0.0}", currentStep.pressureResult("Min")) & " | Avg: " & String.Format("{0:0.0}", currentStep.pressureResult("Avg")) & vbNewLine
                    End If

                    If currentStep.pressureSet("RampRate") <> 0 Then
                        pressRmpStr = "Pressure Ramp (psi/min)" & vbNewLine
                        pressRmpStr = pressRmpStr & "Max: " & String.Format("{0:0.0}", currentStep.pressureResult("MaxRamp")) & " | Min: " & String.Format("{0:0.0}", currentStep.pressureResult("MinRamp")) & "  | Avg: " & String.Format("{0:0.0}", currentStep.pressureResult("AvgRamp")) & vbNewLine
                    End If
                End If

                If curePro.checkVac Then
                    vacStr = "Vacuum (inHg)" & vbNewLine

                    If Math.Abs(currentStep.vacSet("SetPoint")) = Math.Abs(currentStep.vacSet("NegTol")) Then
                        vacStr = vacStr & "Max: " & String.Format("{0:0.0}", currentStep.vacResult("Max")) & vbNewLine
                    ElseIf Math.Abs(currentStep.vacSet("SetPoint")) = Math.Abs(currentStep.vacSet("PosTol")) Then
                        vacStr = vacStr & "Min: " & String.Format("{0:0.0}", currentStep.vacResult("Min")) & vbNewLine
                    Else
                        vacStr = vacStr & "Max: " & String.Format("{0:0.0}", currentStep.vacResult("Max")) & " | Min: " & String.Format("{0:0.0}", currentStep.vacResult("Min")) & " | Avg: " & String.Format("{0:0.0}", currentStep.vacResult("Avg")) & vbNewLine
                    End If

                End If

                cell3 = cell3 & tempStr & tempRmpStr & pressStr & pressRmpStr & vacStr


                If Strings.Right(cell3, 1) = vbLf Then
                    cell3 = Strings.Left(cell3, Len(cell3) - 1)
                End If

                mainSheet.Cells(curRow, 7) = cell3

                'Format cell 3

                If currentStep.tempPass Then
                    formatFont(mainSheet.Cells(curRow, 7), tempStr, 9)
                    formatFont(mainSheet.Cells(curRow, 7), "Temperature (°F)", 11, True,,,, False)
                Else
                    formatFont(mainSheet.Cells(curRow, 7), tempStr, 9,,, True, Color.Red)
                    formatFont(mainSheet.Cells(curRow, 7), "Temperature (°F)", 11, True,, True, Color.Red, False)
                End If

                If currentStep.tempSet("RampRate") <> 0 Then
                    If currentStep.tempRampPass Then
                        formatFont(mainSheet.Cells(curRow, 7), tempRmpStr, 9)
                        formatFont(mainSheet.Cells(curRow, 7), "Temp. Ramp (°F/min)", 11, True,,,, False)
                    Else
                        formatFont(mainSheet.Cells(curRow, 7), tempRmpStr, 9,,, True, Color.Red)
                        formatFont(mainSheet.Cells(curRow, 7), "Temp. Ramp (°F/min)", 11, True,, True, Color.Red, False)
                    End If
                End If

                If currentStep.pressurePass Then
                    formatFont(mainSheet.Cells(curRow, 7), pressStr, 9)
                    formatFont(mainSheet.Cells(curRow, 7), "Pressure (psi)", 11, True,,,, False)
                Else
                    formatFont(mainSheet.Cells(curRow, 7), pressStr, 9,,, True, Color.Red)
                    formatFont(mainSheet.Cells(curRow, 7), "Pressure (psi)", 11, True,, True, Color.Red, False)
                End If

                If currentStep.pressureSet("RampRate") <> 0 Then
                    If currentStep.pressureRampPass Then
                        formatFont(mainSheet.Cells(curRow, 7), pressRmpStr, 9)
                        formatFont(mainSheet.Cells(curRow, 7), "Pressure Ramp (psi/min)", 11, True,,,, False)
                    Else
                        formatFont(mainSheet.Cells(curRow, 7), pressRmpStr, 9,,, True, Color.Red)
                        formatFont(mainSheet.Cells(curRow, 7), "Pressure Ramp (psi/min)", 11, True,, True, Color.Red, False)
                    End If
                End If

                If currentStep.vacPass Then
                    formatFont(mainSheet.Cells(curRow, 7), vacStr, 9)
                    formatFont(mainSheet.Cells(curRow, 7), "Vacuum (inHg)", 11, True,,,, False)
                Else
                    formatFont(mainSheet.Cells(curRow, 7), vacStr, 9,,, True, Color.Red)
                    formatFont(mainSheet.Cells(curRow, 7), "Vacuum (inHg)", 11, True,, True, Color.Red, False)
                End If
            End If

            'Set row height
            Dim rowCount As Integer = 0

            If mainSheet.Cells(curRow, 1).Value.ToString.Split(vbNewLine).Length > rowCount Then rowCount = mainSheet.Cells(curRow, 1).Value.ToString.Split(vbNewLine).Length
            If mainSheet.Cells(curRow, 3).Value.ToString.Split(vbNewLine).Length > rowCount Then rowCount = mainSheet.Cells(curRow, 3).Value.ToString.Split(vbNewLine).Length
            If mainSheet.Cells(curRow, 7).Value.ToString.Split(vbNewLine).Length > rowCount Then rowCount = mainSheet.Cells(curRow, 7).Value.ToString.Split(vbNewLine).Length

            mainSheet.Rows(curRow).RowHeight = rowCount * 15.75

            curRow = curRow + 1
        Next





        'Fill out datasheet
        dataSheet.Activate()
        clearChart(mainSheet)


        Dim curCol As Integer = 1
        outputDataToColumn(dataSheet, curCol, "Time " & vbNewLine & "Stamp" & vbNewLine & "(DateTime)", arr2Str(dateArr), cureStart, cureEnd)
        outputDataToColumn(dataSheet, curCol, "Cure Time" & vbNewLine & "(min)", arr2Str(dateArr,, True), cureStart, cureEnd)
        setChartX(mainSheet)

        If curePro.checkTemp Then
            If machType = "Autoclave" Then
                outputDataToColumn(dataSheet, curCol, "Air TC" & vbNewLine & "(°F)", arr2Str(vessel_TC.values, 1), cureStart, cureEnd)
                Call addToChart(mainSheet, dataSheet, curCol - 1, "Air TC (°F)", vessel_TC, 1)
            End If

            If UBound(usrRunTC) > 0 Then
                outputDataToColumn(dataSheet, curCol, "Lead TC" & vbNewLine & "(°F)", arr2Str(leadTC.values, 1), cureStart, cureEnd)
                Call addToChart(mainSheet, dataSheet, curCol - 1, "Lead TC (°F)", leadTC, 1)
                outputDataToColumn(dataSheet, curCol, "Lag TC" & vbNewLine & "(°F)", arr2Str(lagTC.values, 1), cureStart, cureEnd)
                Call addToChart(mainSheet, dataSheet, curCol - 1, "Lag TC (°F)", lagTC, 1)
            End If

            For i = 0 To UBound(partTC_Arr)
                Dim z As Integer = 0
                For z = 0 To UBound(usrRunTC)
                    If partTC_Arr(i).Number = usrRunTC(z) Then
                        outputDataToColumn(dataSheet, curCol, "TC " & partTC_Arr(i).Number & vbNewLine & "(°F)", arr2Str(partTC_Arr(i).values, 1), cureStart, cureEnd)
                        outputDataToColumn(dataSheet, curCol, "TC " & partTC_Arr(i).Number & " - Ramp" & vbNewLine & "(°F/min)", arr2Str(partTC_Arr(i).ramp, 2), cureStart, cureEnd)
                        If UBound(usrRunTC) = 0 Then
                            Call addToChart(mainSheet, dataSheet, curCol - 2, "TC -" & partTC_Arr(i).Number & "- (°F)", partTC_Arr(i), 1)
                        End If
                    End If
                Next
            Next
        End If

        If curePro.checkPressure Then
            outputDataToColumn(dataSheet, curCol, "Pressure" & vbNewLine & "(psi)", arr2Str(vesselPress.values, 1), cureStart, cureEnd)
            Call addToChart(mainSheet, dataSheet, curCol - 1, "Pressure (psi)", vesselPress, 2)
            outputDataToColumn(dataSheet, curCol, "Pressure Ramp" & vbNewLine & "(psi/min)", arr2Str(vesselPress.ramp, 2), cureStart, cureEnd)
        End If

        If curePro.checkVac Then
            If UBound(usrRunVac) > 0 Then
                outputDataToColumn(dataSheet, curCol, "Max Vac" & vbNewLine & "(inHg)", arr2Str(maxVac.values, 1), cureStart, cureEnd)
                Call addToChart(mainSheet, dataSheet, curCol - 1, "Max Vacuum (inHg)", maxVac, 2)
                outputDataToColumn(dataSheet, curCol, "Min Vac" & vbNewLine & "(inHg)", arr2Str(minVac.values, 1), cureStart, cureEnd)
                Call addToChart(mainSheet, dataSheet, curCol - 1, "Min Vacuum (inHg)", minVac, 2)
            End If

            For i = 0 To UBound(vac_Arr)
                Dim z As Integer = 0
                For z = 0 To UBound(usrRunVac)
                    If vac_Arr(i).Number = usrRunVac(z) Then
                        outputDataToColumn(dataSheet, curCol, "Vac Port " & vac_Arr(i).Number & vbNewLine & "(inHg)", arr2Str(vac_Arr(i).values, 1), cureStart, cureEnd)
                        If UBound(usrRunVac) = 0 Then
                            Call addToChart(mainSheet, dataSheet, curCol - 1, "Vac Port -" & vac_Arr(i).Number & "- (inHg)", vac_Arr(i), 1)
                        End If
                    End If
                Next

            Next
        End If

        dataSheet.Range(dataSheet.Cells(2, 2), dataSheet.Cells(dataCnt, curCol)).NumberFormat = "0.0"
        dataSheet.UsedRange.Columns.AutoFit()
        dataSheet.UsedRange.Rows.AutoFit()
        dataSheet.UsedRange.HorizontalAlignment = 3

        For i = 0 To UBound(curePro.CureSteps) - 1
            Dim stepTime As Double = (dateArr(curePro.CureSteps(i).stepEnd) - dateArr(cureStart)).TotalMinutes
            addStepChart(mainSheet, stepTime)
        Next


        mainSheet.Activate()


        SwitchOff(Excel, False)

        Excel.Visible = True
    End Sub
    Sub clearChart(mainSheet As Excel.Worksheet)
        Dim mainChart As Excel.Chart = mainSheet.ChartObjects.Item(1).Chart
        Dim mainChartSeriesCollect As Excel.SeriesCollection = mainChart.SeriesCollection

        For i = mainChartSeriesCollect.Count To 1 Step -1
            mainChartSeriesCollect.Item(i).Delete()
        Next

        mainChart.Axes(2, 1).MaximumScale = 140
        mainChart.Axes(2, 1).MinimumScale = 100

        'mainChart.Axes(2, 2).MaximumScale = 0
        'mainChart.Axes(2, 2).MinimumScale = -30
    End Sub

    Sub setChartX(mainSheet As Excel.Worksheet)
        Dim mainChart As Excel.Chart = mainSheet.ChartObjects.Item(1).Chart
        Dim mainChartSeriesCollect As Excel.SeriesCollection = mainChart.SeriesCollection

        mainChart.Axes(1).MinimumScale = 0
        mainChart.Axes(1).MaximumScale = Math.Round((dateArr(cureEnd) - dateArr(cureStart)).TotalMinutes, 0)
    End Sub

    Sub addToChart(mainSheet As Excel.Worksheet, dataSheet As Excel.Worksheet, cNum As Integer, seriesName As String, startDataSet As DataSet, Optional axisGroup As Integer = 1)
        Dim mainChart As Excel.Chart = mainSheet.ChartObjects.Item(1).Chart
        Dim mainChartSeriesCollect As Excel.SeriesCollection = mainChart.SeriesCollection

        Dim cureTimesRng As Excel.Range = dataSheet.Range(dataSheet.Cells(2, 2), dataSheet.Cells(cureEnd - cureStart + 2, 2))
        Dim cureValRng As Excel.Range = dataSheet.Range(dataSheet.Cells(2, cNum), dataSheet.Cells(cureEnd - cureStart + 2, cNum))

        Dim ser As Excel.Series
        ser = mainChartSeriesCollect.NewSeries

        ser.Values = cureValRng
        ser.XValues = cureTimesRng
        ser.Name = seriesName

        If axisGroup <> 1 And axisGroup <> 2 Then
            Throw New Exception("Axis group can only be applied to 1 or 2.")
        End If

        ser.AxisGroup = axisGroup

        ser.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone
        ser.Format.Line.Visible = True
        ser.Format.Line.Weight = 1

        Dim valMax As Double = startDataSet.Max()

        If valMax > mainChart.Axes(2, axisGroup).MaximumScale Then
            mainChart.Axes(2, axisGroup).MaximumScale = valMax
        End If

        If axisGroup = 2 Then
            mainChart.Axes(2, axisGroup).HasTitle = True
            mainChart.Axes(2, axisGroup).AxisTitle.Text = "Pressure (psi) | Vacuum (inHg)"
        End If


    End Sub

    Sub addStepChart(mainSheet As Excel.Worksheet, intime As Double)
        Dim mainChart As Excel.Chart = mainSheet.ChartObjects.Item(1).Chart
        Dim mainChartSeriesCollect As Excel.SeriesCollection = mainChart.SeriesCollection

        Dim ser As Excel.Series
        ser = mainChartSeriesCollect.NewSeries

        Dim avgVal As Double = (mainChart.Axes(1).MinimumScale + mainChart.Axes(1).MaximumScale) / 2

        ser.Values = avgVal
        ser.XValues = intime
        ser.Name = ""


        ser.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone
        ser.Format.Line.Visible = False

        ser.HasErrorBars = True
        ser.ErrorBars.EndStyle = Excel.XlEndStyleCap.xlNoCap

        ser.ErrorBar(Excel.XlErrorBarDirection.xlY, Excel.XlErrorBarInclude.xlErrorBarIncludeBoth, Excel.XlErrorBarType.xlErrorBarTypePercent, 100)
        ser.ErrorBar(Excel.XlErrorBarDirection.xlX, Excel.XlErrorBarInclude.xlErrorBarIncludeNone, Excel.XlErrorBarType.xlErrorBarTypeFixedValue)
    End Sub


    Function arr2Str(Of T)(inArr() As T, Optional rndDecimals As Integer = 1, Optional calcDuration As Boolean = False) As String()
        Dim retArr() As String

        Dim i As Integer = 0
        For i = 0 To UBound(inArr)

            Dim inValue As String = ""

            If TypeName(inArr) = "Date()" And calcDuration Then
                Dim startDate As Object = inArr(cureStart)
                Dim checkDate As Object = inArr(i)
                inValue = Math.Round((checkDate - startDate).totalminutes, rndDecimals).ToString()
            ElseIf TypeName(inArr) = "Date()" Then
                inValue = inArr(i).ToString()
            ElseIf TypeName(inArr) = "Double()" Then
                Dim holder As Object = inArr(i)
                inValue = Math.Round(holder, rndDecimals).ToString()
            Else
                inValue = inArr(0).ToString
            End If

            retArr.AddValArr(inValue)
        Next

        Return retArr
    End Function

    Sub outputDataToColumn(dataSheet As Excel.Worksheet, ByRef cNum As Integer, title As String, outputDataSet() As String, Optional startVal As Integer = 0, Optional endVal As Integer = -1, Optional startRow As Integer = 1)
        If endVal = -1 Then endVal = UBound(outputDataSet)

        dataSheet.Cells(startRow, cNum) = title
        startRow += 1

        Dim i As Integer
        i = startVal
        Do While i <= endVal
            dataSheet.Cells(startRow, cNum) = outputDataSet(i)
            startRow += 1
            i += 1
        Loop

        cNum += 1
    End Sub


    Sub optionNegTest(negTolValue As Integer, ByRef optionNeg As String)
        If negTolValue = 0 Then
            optionNeg = "-"
        Else
            optionNeg = ""
        End If
    End Sub

    Function termToStr(termCond As Dictionary(Of String, Object), currentStep As CureStep) As String

        Dim currentStr As String = ""

        If termCond("Type") = "None" Then
            Return ""


        ElseIf termCond("Type") = "Time" Then
            currentStr = currentStr & "After " & termCond("Goal") & " min"
            currentStr = currentStr & vbNewLine


        ElseIf termCond("Type") = "Temp" Then
            currentStr = currentStr & termCond("TCNum")

            If termCond("Condition") = "GREATER" Then
                currentStr = currentStr & " > "
            ElseIf termCond("Condition") = "LESS" Then
                currentStr = currentStr & " < "
            Else
                Throw New Exception("Unknown termination condition for temp on " & currentStep.stepName)
            End If

            currentStr = currentStr & termCond("Goal") & "°F"
            currentStr = currentStr & vbNewLine


        ElseIf termCond("Type") = "Press" Then
            If termCond("Condition") = "GREATER" Then
                currentStr = currentStr & " > "
            ElseIf termCond("Condition") = "LESS" Then
                currentStr = currentStr & " < "
            Else
                Throw New Exception("Unknown termination condition for temp on " & currentStep.stepName)
            End If

            currentStr = currentStr & termCond("Goal") & " psi"
            currentStr = currentStr & vbNewLine
        End If

        Return currentStr
    End Function



#End Region

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
                        indexStart = indexStart + Math.Round(curePro.CureSteps(i - 1).tempSet("RampRate") / (dateArr(1) - dateArr(0)).TotalMinutes, 0)
                    End If
                End If

                If Not lastStep Then
                    If curePro.CureSteps(i + 1).tempSet("RampRate") > 0 Then
                        indexEnd = indexEnd - Math.Round(curePro.CureSteps(i - 1).tempSet("RampRate") / (dateArr(1) - dateArr(0)).TotalMinutes, 0)
                    End If
                End If

                curePro.CureSteps(i).tempResult("Max") = Math.Round(leadTC.Max(indexStart, indexEnd), 0)


                ''Min temp
                If Not firstStep Then
                    If curePro.CureSteps(i - 1).tempSet("RampRate") > 0 Then
                        indexStart = indexStart + Math.Round(curePro.CureSteps(i - 1).tempSet("RampRate") / (dateArr(1) - dateArr(0)).TotalMinutes, 0)
                    End If
                End If

                If Not lastStep Then
                    If curePro.CureSteps(i + 1).tempSet("RampRate") < 0 Then
                        indexEnd = indexEnd - Math.Round(curePro.CureSteps(i - 1).tempSet("RampRate") / (dateArr(1) - dateArr(0)).TotalMinutes, 0)
                    End If
                End If

                curePro.CureSteps(i).tempResult("Min") = Math.Round(lagTC.Min(indexStart, indexEnd), 0)


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

                curePro.CureSteps(i).tempResult("Avg") = Math.Round(total / addCnt, 0)

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

                curePro.CureSteps(i).tempResult("MaxRamp") = Math.Round(holder, 1)

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
                curePro.CureSteps(i).tempResult("MinRamp") = Math.Round(holder, 1)


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

                curePro.CureSteps(i).tempResult("AvgRamp") = Math.Round(total / addCnt, 1)

                'Check temp for passing
                If curePro.CureSteps(i).tempResult("Min") < curePro.CureSteps(i).tempSet("SetPoint") + curePro.CureSteps(i).tempSet("NegTol") Then curePro.CureSteps(i).tempPass = False
                If curePro.CureSteps(i).tempResult("Max") > curePro.CureSteps(i).tempSet("SetPoint") + curePro.CureSteps(i).tempSet("PosTol") Then curePro.CureSteps(i).tempPass = False

                'Check temp ramp for passing if not equal to 0
                If curePro.CureSteps(i).tempSet("RampRate") <> 0 Then
                    If curePro.CureSteps(i).tempResult("MinRamp") < curePro.CureSteps(i).tempSet("RampRate") + curePro.CureSteps(i).tempSet("RampNegTol") Then curePro.CureSteps(i).tempRampPass = False
                    If curePro.CureSteps(i).tempResult("MaxRamp") > curePro.CureSteps(i).tempSet("RampRate") + curePro.CureSteps(i).tempSet("RampPosTol") Then curePro.CureSteps(i).tempRampPass = False
                End If
            Else
                curePro.CureSteps(i).tempPass = True
                curePro.CureSteps(i).tempRampPass = True
            End If

            'Check autoclave only features

            'Calculate pressure results for a given step
            If curePro.checkPressure Then
                ''Max pressure
                If Not firstStep Then
                    If curePro.CureSteps(i - 1).pressureSet("RampRate") < 0 Then
                        indexStart = indexStart + Math.Round(curePro.CureSteps(i - 1).pressureSet("RampRate") / (dateArr(1) - dateArr(0)).TotalMinutes, 0)
                    End If
                End If

                If Not lastStep Then
                    If curePro.CureSteps(i + 1).pressureSet("RampRate") > 0 Then
                        indexEnd = indexEnd - Math.Round(curePro.CureSteps(i - 1).pressureSet("RampRate") / (dateArr(1) - dateArr(0)).TotalMinutes, 0)
                    End If
                End If
                curePro.CureSteps(i).pressureResult("Max") = Math.Round(vesselPress.Max(indexStart, indexEnd), 1)

                ''Min pressure
                If Not firstStep Then
                    If curePro.CureSteps(i - 1).pressureSet("RampRate") > 0 Then
                        indexStart = indexStart + Math.Round(curePro.CureSteps(i - 1).pressureSet("RampRate") / (dateArr(1) - dateArr(0)).TotalMinutes, 0)
                    End If
                End If

                If Not lastStep Then
                    If curePro.CureSteps(i + 1).pressureSet("RampRate") < 0 Then
                        indexEnd = indexEnd - Math.Round(curePro.CureSteps(i - 1).pressureSet("RampRate") / (dateArr(1) - dateArr(0)).TotalMinutes, 0)
                    End If

                End If
                curePro.CureSteps(i).pressureResult("Min") = Math.Round(vesselPress.Min(indexStart, indexEnd), 1)

                ''Average pressure
                indexStart = curePro.CureSteps(i).stepStart
                indexEnd = curePro.CureSteps(i).stepEnd
                curePro.CureSteps(i).pressureResult("Avg") = Math.Round(vesselPress.Average(indexStart, indexEnd), 1)

                curePro.CureSteps(i).pressureResult("MaxRamp") = Math.Round(vesselPress.MaxRamp(indexStart, indexEnd), 1)
                curePro.CureSteps(i).pressureResult("MinRamp") = Math.Round(vesselPress.MinRamp(indexStart, indexEnd), 1)
                curePro.CureSteps(i).pressureResult("AvgRamp") = Math.Round(vesselPress.AverageRamp(indexStart, indexEnd), 1)



                'Check pressure for passing
                If curePro.CureSteps(i).pressureResult("Min") < curePro.CureSteps(i).pressureSet("SetPoint") + curePro.CureSteps(i).pressureSet("NegTol") Then curePro.CureSteps(i).pressurePass = False
                If curePro.CureSteps(i).pressureResult("Max") > curePro.CureSteps(i).pressureSet("SetPoint") + curePro.CureSteps(i).pressureSet("PosTol") Then curePro.CureSteps(i).pressurePass = False

                'Check pressure ramp if not set to 0
                If curePro.CureSteps(i).pressureSet("RampRate") <> 0 Then
                    If curePro.CureSteps(i).pressureResult("MinRamp") < curePro.CureSteps(i).pressureSet("RampRate") + curePro.CureSteps(i).pressureSet("RampNegTol") Then curePro.CureSteps(i).pressureRampPass = False
                    If curePro.CureSteps(i).pressureResult("MaxRamp") > curePro.CureSteps(i).pressureSet("RampRate") + curePro.CureSteps(i).pressureSet("RampPosTol") Then curePro.CureSteps(i).pressureRampPass = False
                End If
            Else
                curePro.CureSteps(i).pressurePass = True
                curePro.CureSteps(i).pressureRampPass = True
            End If

            'Calculate vac results for a given step
            If curePro.checkVac Then
                curePro.CureSteps(i).vacResult("Max") = Math.Round(maxVac.Max(indexStart, indexEnd), 1)
                curePro.CureSteps(i).vacResult("Min") = Math.Round(minVac.Min(indexStart, indexEnd), 1)

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

                curePro.CureSteps(i).vacResult("Avg") = Math.Round(total / addCnt, 1)

                'Check vacuum for passing
                If curePro.CureSteps(i).vacResult("Min") < curePro.CureSteps(i).vacSet("SetPoint") + curePro.CureSteps(i).vacSet("NegTol") Then curePro.CureSteps(i).vacPass = False
                If curePro.CureSteps(i).vacResult("Max") > curePro.CureSteps(i).vacSet("SetPoint") + curePro.CureSteps(i).vacSet("PosTol") Then curePro.CureSteps(i).vacPass = False
            Else
                curePro.CureSteps(i).vacPass = True
            End If



            'Check for all passing
            If curePro.CureSteps(i).vacPass And curePro.CureSteps(i).tempPass And curePro.CureSteps(i).tempRampPass And curePro.CureSteps(i).pressurePass And curePro.CureSteps(i).pressureRampPass And curePro.CureSteps(i).stepTerminate Then
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
                    cureEnd = curePro.CureSteps(currentStep).stepEnd
                    dateValues("endTime") = dateArr(curePro.CureSteps(currentStep).stepEnd)
                    currentStep += 1
                    Exit For
                End If

                currentStep += 1
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
        ElseIf UBound(curePro.CureSteps) = currentStep Then
            curePro.CureSteps(currentStep).stepTerminate = False
            curePro.CureSteps(currentStep).stepPass = False
            curePro.curePass = False
            MsgBox("Failed to reach terminating conditions.")
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
                    If lagTC.values(currentStep) > termCond("Goal") Then
                        Return True
                    End If
                ElseIf termCond("TCNum") = "Lead" Then
                    If leadTC.values(currentStep) > termCond("Goal") Then
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
                    If leadTC.values(currentStep) < termCond("Goal") Then
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
#End Region

#Region "Derived arrays from data"
    Sub startEndTime()
        Dim start_end_temp As Double = 140

        'Get start time
        Dim i As Integer
        For i = 0 To dataCnt
            If leadTC.values(i) > start_end_temp And dateValues("startTime") = Nothing Then
                dateValues("startTime") = dateArr(i)
                cureStart = i
                Exit For
            End If
        Next

        'If the start time wasn't triggered then the cure starts at 0
        If cureStart = 0 Then
            dateValues("startTime") = dateArr(0)
        End If

        'Get end time
        Dim runStart As Boolean = False
        For i = 0 To dataCnt
            If lagTC.values(i) > start_end_temp And runStart = False Then
                runStart = True
                i += 5
            End If

            If runStart = True Then
                If leadTC.values(i) < start_end_temp Then
                    dateValues("endTime") = dateArr(i)
                    cureEnd = i
                    Exit For
                End If
            End If
        Next

        'If cureEnd didn't get set then the end of data is the end
        If cureEnd = 0 Then
            dateValues("endTime") = dateArr(dataCnt)
            cureEnd = dataCnt
        End If

        If cureEnd <> dataCnt Then
            For i = cureEnd + 1 To dataCnt
                If leadTC.values(i) > start_end_temp Then
                    MsgBox("Temperature rose above " & start_end_temp & "°F after the cure completed. Check data for deviation.", vbExclamation, "Temperature Error")
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
        stepVal = 5 / ((dateArr(1) - dateArr(0)).TotalMinutes)
        If stepVal = 1 Then stepVal = 2

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
            Call addToList(partTC_Arr, Data_TC, Box_TC)
        End If

        If curePro.checkPressure Then
            vesselPress.values.clearArr()
            Call getDataSearch("{Vessel_Pressure}", vesselPress)
        End If

        If curePro.checkVac Then
            vac_Arr.clearArr()
            Call getDataMulti("{VacGroup_", vac_Arr, "VAC")
            Call addToList(vac_Arr, Data_Vac, Box_Vac)
        End If

        If machType = "Autoclave" Then
            vessel_TC.values.clearArr()
            Call getDataSearch("{Air_TC}", vessel_TC)
        End If
    End Sub

    Sub addToList(inArr() As DataSet, inGrid As DataGridView, inContainer As GroupBox)
        inContainer.Visible = True
        inGrid.Rows.Clear()

        For i = 0 To UBound(inArr)
            inGrid.Rows.Add()
            inGrid.Item(1, i).Value = inArr(i).Number
            'inGrid.Item(0, i).Value = True
        Next

        inGrid.Sort(inGrid.Columns(1), System.ComponentModel.ListSortDirection.Ascending)

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


        'Reset machType to null
        machType = "Unknown"

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

            If System.Text.RegularExpressions.Regex.Match(inFile, "\d{6}").Value <> "" Then
                Txt_JobNumber.Text = System.Text.RegularExpressions.Regex.Match(inFile, "\d{6}").Value
            End If

            If System.Text.RegularExpressions.Regex.Match(inFile, "S(\d{4})").Value <> "" Then
                Txt_DataRecorder.Text = System.Text.RegularExpressions.Regex.Match(inFile, "S(\d{4})").Value
            End If

            If System.Text.RegularExpressions.Regex.Match(inFile, "D(\d{5})").Value <> "" Then
                Txt_ProgramNumber.Text = System.Text.RegularExpressions.Regex.Match(inFile, "D(\d{5})").Value
            End If




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

                                Box_DataRecorder.Visible = True

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

        If headerRow <0 Then
            Throw New Exception("Header row cannot be less than 0. Please check to make sure your file has the correct amount of header lines for it's type.")
        End If
    End Sub

#Region "Cure profile inport/export"
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

        Combo_CureProfile.Items.Clear()
        cureProfiles.clearArr()

        If IO.File.Exists(inPath) Then
            loadCureFile(inPath)
        ElseIf IO.Directory.Exists(inPath) Then
            For Each file In IO.Directory.GetFiles(inPath)
                loadCureFile(file)
            Next
        End If

        If cureProfiles Is Nothing Then
            Throw New Exception("No cure profiles were found at your specified path.")
        Else
            Combo_CureProfile.SelectedIndex = 0
        End If


    End Sub

    Sub loadCureFile(inPath As String)
        Dim cureDef() As String

        If IO.File.Exists(inPath) And IO.Path.GetExtension(inPath) = ".cprof" Then
            cureDef = Split(IO.File.ReadAllText(inPath), "~&&&~")

            For i = 0 To UBound(cureDef)
                If cureProfiles Is Nothing Then
                    ReDim cureProfiles(0)
                    cureProfiles(0) = New CureProfile()
                Else
                    ReDim Preserve cureProfiles(UBound(cureProfiles) + 1)
                    cureProfiles(UBound(cureProfiles)) = New CureProfile()
                End If

                cureProfiles(UBound(cureProfiles)).deserializeCure(cureDef(i))
                Combo_CureProfile.Items.Add(cureProfiles(UBound(cureProfiles)).Name)
            Next
        End If
    End Sub
#End Region

    Private Sub Btn_OpenFile_Click(sender As Object, e As EventArgs) Handles Btn_OpenFile.Click
        If IO.File.Exists(Txt_FilePath.Text) Then
            Call openCureFile()
        ElseIf OpenCSVFileDialog.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            Txt_FilePath.Text = OpenCSVFileDialog.FileName
            Call openCureFile()
        End If
    End Sub

    Sub openCureFile()
        Box_DataRecorder.Visible = False
        Box_TC.Visible = False
        Box_Vac.Visible = False
        Try
            Call loadCSVin(Txt_FilePath.Text)
            Call loadCureData()
            Box_RunParams.Enabled = True
            Box_RunLine.Enabled = True
        Catch ex As Exception
            If MsgBox(ex.Message, vbOKCancel, "Error") = vbCancel Then Me.Close()
            Box_RunLine.Enabled = False
        End Try

    End Sub

    Private Sub Txt_FilePath_TextChanged(sender As Object, e As EventArgs) Handles Txt_FilePath.TextChanged
        If IO.File.Exists(Txt_FilePath.Text) Then
            Txt_FilePath.BackColor = SystemColors.Window
        Else
            Txt_FilePath.BackColor = Color.PeachPuff
            Box_RunLine.Enabled = False
        End If
    End Sub

    Private Sub Btn_LoadProfileFiles_Click(sender As Object, e As EventArgs) Handles Btn_LoadProfileFiles.Click


        If IO.File.Exists(Txt_CureProfilesPath.Text) Then
            Call loadCureProfiles(Txt_CureProfilesPath.Text)
        ElseIf IO.Directory.Exists(Txt_CureProfilesPath.Text) Then
            Call loadCureProfiles(Txt_CureProfilesPath.Text)
        ElseIf OpenCureProfileFileDialog.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            Txt_CureProfilesPath.Text = OpenCureProfileFileDialog.FileName
            Call loadCureProfiles(Txt_CureProfilesPath.Text)
        End If

        My.Settings.CureProfilePath = Txt_CureProfilesPath.Text
    End Sub

    Private Sub Txt_CureProfiles_TextChanged(sender As Object, e As EventArgs) Handles Txt_CureProfilesPath.TextChanged
        If IO.File.Exists(Txt_CureProfilesPath.Text) Then
            Txt_CureProfilesPath.BackColor = SystemColors.Window
        ElseIf IO.Directory.Exists(Txt_CureProfilesPath.Text) Then
            Txt_CureProfilesPath.BackColor = SystemColors.Window
        Else
            Txt_CureProfilesPath.BackColor = Color.PeachPuff
        End If
    End Sub

    Private Sub Combo_CureProfile_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_CureProfile.SelectedIndexChanged

        If cureProfiles(Combo_CureProfile.SelectedIndex).Name <> Combo_CureProfile.Text Then
            Throw New Exception("Cure profile does not line up with loaded array, please reload cure profiles.")
        End If

        curePro = cureProfiles(Combo_CureProfile.SelectedIndex)

        Txt_CureDoc.Text = curePro.cureDoc
        Txt_DocRev.Text = curePro.cureDocRev

        Txt_FilePath.Text = "File path..."
        Box_CureDataLocation.Enabled = True
    End Sub

    Private Sub Txt_TemplatePath_TextChanged(sender As Object, e As EventArgs) Handles Txt_TemplatePath.TextChanged
        If IO.File.Exists(Txt_TemplatePath.Text) Then
            My.Settings.ReportTemplatePath = Txt_TemplatePath.Text
            Txt_TemplatePath.BackColor = SystemColors.Window
        Else
            Txt_TemplatePath.BackColor = Color.PeachPuff
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



