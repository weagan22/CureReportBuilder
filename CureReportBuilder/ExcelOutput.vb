Option Explicit On

Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices.Marshal

Public Class ExcelOutput
    Private Declare Function GetWindowThreadProcessId Lib "user32.dll" Alias "GetWindowThreadProcessId" (ByVal hwnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer

    Dim excelProcID As Integer

    'Dim checkToOutput As CureCheck = New CureCheck

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

    Sub outputResults(templatePath As String, checkToOutput As CureCheck, Optional outputPath As String = "")
        Dim Excel As Object
        Excel = CreateObject("Excel.Application")
        Call GetWindowThreadProcessId(IntPtr.op_Explicit(Excel.hwnd), excelProcID)
        Excel.Visible = False

        Try
            Excel.Workbooks.Open(templatePath)
        Catch ex As Exception
            killExcel(excelProcID)
            Throw New Exception("Failed to open Excel Report Template.")
            Exit Sub
        End Try

        SwitchOff(Excel, True)

        Dim mainSheet As Excel.Worksheet = Excel.Sheets.Item(1)
        Dim userSheet As Excel.Worksheet = Excel.Sheets.Item(2)
        Dim dataSheet As Excel.Worksheet = Excel.Sheets.Item(3)

        runInfoOutput(userSheet, checkToOutput)

        mainSheet.Cells(2, 1) = "Job" & vbNewLine & checkToOutput.JobNum & vbNewLine & "Program" & vbNewLine & checkToOutput.ProgramNum
        formatFont(mainSheet.Cells(2, 1), "Job", 14, True, False, True)
        formatFont(mainSheet.Cells(2, 1), "Program", 14, True, False, True)

        mainSheet.Cells(2, 4) = "Part" & vbNewLine & checkToOutput.PartNum & vbNewLine & "Rev. " & checkToOutput.PartRev & vbNewLine & checkToOutput.PartNom & vbNewLine & "Qty: " & checkToOutput.PartQty
        formatFont(mainSheet.Cells(2, 4), "Part", 14, True, False, True)


        Dim hours As Integer = Math.Round((checkToOutput.dateArr(checkToOutput.cureEnd) - checkToOutput.dateArr(checkToOutput.cureStart)).TotalMinutes, 1) \ 60
        Dim minutes As Integer = Math.Round((checkToOutput.dateArr(checkToOutput.cureEnd) - checkToOutput.dateArr(checkToOutput.cureStart)).TotalMinutes, 1) - (hours * 60)

        mainSheet.Cells(2, 8) = "Cure" & vbNewLine & "Start: " & checkToOutput.dateArr(checkToOutput.cureStart).ToString("dd-MMM-yyyy | h:mm tt") & vbNewLine & "Finish: " & checkToOutput.dateArr(checkToOutput.cureEnd).ToString("dd-MMM-yyyy | h:mm tt") & vbNewLine & "Duration: " & hours & "h:" & minutes & "m" 'Math.Round((checkToOutput.dateArr(checkToOutput.cureEnd) - checkToOutput.dateArr(checkToOutput.cureStart)).TotalMinutes, 1) & " min"
        formatFont(mainSheet.Cells(2, 8), "Cure", 14, True, False, True)

        mainSheet.Cells(5, 1) = "Equipment" & vbNewLine & checkToOutput.machType & " | " & checkToOutput.equipSerialNum
        formatFont(mainSheet.Cells(5, 1), "Equipment", 14, True, False, True)

        If checkToOutput.curePro.curePass Then
            mainSheet.Cells(6, 4) = "PASS"
            formatFont(mainSheet.Cells(6, 4), "PASS", 30, True, False, False, Color.Green)
        Else
            mainSheet.Cells(6, 4) = "FAIL"
            formatFont(mainSheet.Cells(6, 4), "FAIL", 30, True, False, True, Color.Red)
        End If

        mainSheet.Cells(5, 8) = "Cure Document" & vbNewLine & checkToOutput.curePro.cureDoc & vbNewLine & "Rev. " & checkToOutput.curePro.cureDocRev & vbNewLine & "Profile: " & checkToOutput.curePro.Name
        formatFont(mainSheet.Cells(5, 8), "Cure Document", 14, True, False, True)

        mainSheet.Cells(34, 3) = checkToOutput.completedBy
        mainSheet.Cells(34, 9) = Today.Date().ToString("MM-dd-yyyy")


        Dim curRow As Integer = 40
        For i = 0 To UBound(checkToOutput.curePro.CureSteps)

            Dim currentStep As CureStep = checkToOutput.curePro.CureSteps(i)

            'Merge cells and center text
            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 1), mainSheet.Cells(curRow, 2)).Merge
            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 3), mainSheet.Cells(curRow, 6)).Merge
            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 7), mainSheet.Cells(curRow, 9)).Merge

            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 1), mainSheet.Cells(curRow, 10)).HorizontalAlignment = 3

            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 1), mainSheet.Cells(curRow, 10)).VerticalAlignment = -4108

            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 1), mainSheet.Cells(curRow, 10)).WrapText = True

            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 1), mainSheet.Cells(curRow, 10)).Borders(9).LineStyle = 1

            'Fill out data in cell 1
            If currentStep.hardFail Then
                mainSheet.Cells(curRow, 1) = currentStep.stepName
                formatFont(mainSheet.Cells(curRow, 1), currentStep.stepName, 14, True, False, False)
            Else
                Dim stepPass As String = "FAIL"
                If currentStep.stepPass Then stepPass = "PASS"

                Dim startTime As Double = (checkToOutput.dateArr(currentStep.stepStart) - checkToOutput.dateArr(checkToOutput.cureStart)).TotalMinutes
                Dim endTime As Double = (checkToOutput.dateArr(currentStep.stepEnd) - checkToOutput.dateArr(checkToOutput.cureStart)).TotalMinutes

                Dim endStr As String = ""
                Dim durationStr As String = ""
                If currentStep.stepTerminate = False Then
                    endStr = "Failed to Terminate"
                    durationStr = ""
                Else
                    endStr = "Finish: " & endTime.ToString("0.0") & " min"
                    durationStr = vbNewLine & "Duration: " & String.Format("{0:0.0}", Math.Round(endTime - startTime, 1)) & " min"
                End If


                mainSheet.Cells(curRow, 1) = currentStep.stepName & vbNewLine & "(" & stepPass & ")" & vbNewLine & "Start: " & startTime.ToString("0.0") & " min" & vbNewLine & endStr & durationStr

                formatFont(mainSheet.Cells(curRow, 1), currentStep.stepName, 14, True, False, False)

                If currentStep.stepPass = True Then
                    formatFont(mainSheet.Cells(curRow, 1), "(" & stepPass & ")", 11, False, False, False)
                Else
                    formatFont(mainSheet.Cells(curRow, 1), "(" & stepPass & ")", 11, False, False, False, Color.Red)
                End If

                formatFont(mainSheet.Cells(curRow, 1), "Start: " & startTime.ToString("0.0") & " min", 9, False, False, False)


                If currentStep.stepTerminate = False Then
                    formatFont(mainSheet.Cells(curRow, 1), endStr, 9, False, False, False, Color.Red)
                Else
                    formatFont(mainSheet.Cells(curRow, 1), endStr, 9, False, False, False)

                    If currentStep.timeLimitPass Then
                        formatFont(mainSheet.Cells(curRow, 1), durationStr, 9, False, False, False)
                    Else
                        formatFont(mainSheet.Cells(curRow, 1), durationStr, 9, False, False, False, Color.Red)
                    End If

                End If
            End If

            'Fill out cell 2
            Dim cell2 As String = ""

            Dim tempStr As String = ""
            Dim tempRmpStr As String = ""
            Dim pressStr As String = ""
            Dim pressRmpStr As String = ""
            Dim vacStr As String = ""



            If checkToOutput.curePro.checkTemp And currentStep.tempSet.SetPoint <> -1 Then
                If Math.Abs(currentStep.tempSet.SetPoint) = Math.Abs(currentStep.tempSet.NegTol) Then
                    tempStr = "Temperature Max " & currentStep.tempSet.SetPoint + currentStep.tempSet.PosTol & "°F"
                ElseIf Math.Abs(currentStep.tempSet.SetPoint) = Math.Abs(currentStep.tempSet.PosTol) Then
                    tempStr = "Temperature Min " & currentStep.tempSet.SetPoint + currentStep.tempSet.NegTol & "°F"
                Else

                    tempStr = "Temperature " & currentStep.tempSet.SetPoint & "°F " & plusMinusVal(currentStep.tempSet.PosTol, currentStep.tempSet.NegTol) & "°F"
                End If
                tempStr = tempStr & vbNewLine

                If currentStep.tempSet.RampRate <> 0 Then
                    tempRmpStr = "Temp. Ramp " & currentStep.tempSet.RampRate & "°F/min " & plusMinusVal(currentStep.tempSet.RampPosTol, currentStep.tempSet.RampNegTol) & "°F/min"
                    tempRmpStr = tempRmpStr & vbNewLine
                End If
            End If

            If checkToOutput.curePro.checkPressure And currentStep.pressureSet.SetPoint <> -1 Then
                If Math.Abs(currentStep.pressureSet.SetPoint) = Math.Abs(currentStep.pressureSet.NegTol) Then
                    pressStr = "Pressure: Max " & currentStep.pressureSet.SetPoint + currentStep.pressureSet.PosTol & " psi"
                ElseIf Math.Abs(currentStep.pressureSet.SetPoint) = Math.Abs(currentStep.pressureSet.PosTol) Then
                    pressStr = "Pressure Min " & currentStep.pressureSet.SetPoint + currentStep.pressureSet.NegTol & " psi"
                Else
                    pressStr = pressStr & "Pressure:  " & currentStep.pressureSet.SetPoint & " psi " & plusMinusVal(currentStep.pressureSet.PosTol, currentStep.pressureSet.NegTol) & " psi"
                End If
                pressStr = pressStr & vbNewLine

                If currentStep.pressureSet.RampRate <> 0 Then
                    pressRmpStr = pressRmpStr & "Pressure Ramp " & currentStep.pressureSet.RampRate & " psi/min " & plusMinusVal(currentStep.pressureSet.RampPosTol, currentStep.pressureSet.RampNegTol) & " psi/min"
                    pressRmpStr = pressRmpStr & vbNewLine
                End If
            End If

            If checkToOutput.curePro.checkVac And currentStep.vacSet.SetPoint <> -1 Then
                If Math.Abs(currentStep.vacSet.SetPoint) = Math.Abs(currentStep.vacSet.NegTol) Then
                    vacStr = "Vacuum Min " & currentStep.vacSet.SetPoint + currentStep.vacSet.PosTol & " inHg"
                ElseIf Math.Abs(currentStep.vacSet.SetPoint) = Math.Abs(currentStep.vacSet.PosTol) Then
                    vacStr = "Vacuum: Max " & currentStep.vacSet.SetPoint + currentStep.vacSet.NegTol & " inHg"
                Else
                    vacStr = vacStr & "Vacuum:  " & currentStep.vacSet.SetPoint & " inHg " & plusMinusVal(currentStep.vacSet.PosTol, currentStep.vacSet.NegTol) & " inHg"
                End If
                vacStr = vacStr & vbNewLine
            End If

            cell2 = cell2 & tempStr & tempRmpStr & pressStr & pressRmpStr & vacStr

            cell2 = cell2 & "Terminate"
            cell2 = cell2 & vbNewLine



            Dim term1Str As String = termToStr(currentStep.termCond1Type, currentStep.termCond1Condition, currentStep.termCond1Goal, currentStep.termCond1Modifier, currentStep)
            Dim term2Str As String = ""

            If currentStep.termCond2Type <> "None" Then
                term2Str = term2Str & currentStep.termCondOper
                term2Str = term2Str & vbNewLine
                term2Str = term2Str & termToStr(currentStep.termCond2Type, currentStep.termCond2Condition, currentStep.termCond2Goal, currentStep.termCond2Modifier, currentStep)
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
                mainSheet.Cells(curRow, 7) = "Failed To Reach Step"
                formatFont(mainSheet.Cells(curRow, 7), "Failed To Reach Step", 11, True,, True, Color.Red)
            Else
                Dim cell3 As String = ""

                tempStr = ""
                tempRmpStr = ""
                pressStr = ""
                pressRmpStr = ""
                vacStr = ""

                If checkToOutput.curePro.checkTemp And currentStep.tempSet.SetPoint <> -1 Then
                    tempStr = "Temperature (°F)" & vbNewLine






                    If Math.Abs(currentStep.tempSet.SetPoint) = Math.Abs(currentStep.tempSet.NegTol) Then
                        tempStr = tempStr & "Max " & String.Format("{0}", Math.Round(currentStep.tempResult.Max, 0)) & vbNewLine
                    ElseIf Math.Abs(currentStep.tempSet.SetPoint) = Math.Abs(currentStep.tempSet.PosTol) Then
                        tempStr = tempStr & "Min " & String.Format("{0}", Math.Round(currentStep.tempResult.Min, 0)) & vbNewLine
                    Else
                        tempStr = tempStr & "Max " & String.Format("{0}", Math.Round(currentStep.tempResult.Max, 0))
                        tempStr = tempStr & " | Min " & String.Format("{0}", Math.Round(currentStep.tempResult.Min, 0))
                        tempStr = tempStr & "  | Avg " & String.Format("{0}", Math.Round(currentStep.tempResult.Avg, 0)) & vbNewLine
                    End If

                    If currentStep.tempSet.RampRate <> 0 Then
                        tempRmpStr = "Temp. Ramp (°F/min)" & vbNewLine
                        tempRmpStr = tempRmpStr & "Max " & String.Format("{0:0.0}", Math.Round(currentStep.tempResult.MaxRamp, 1))
                        tempRmpStr = tempRmpStr & " | Min " & String.Format("{0:0.0}", Math.Round(currentStep.tempResult.MinRamp, 1)) & vbNewLine
                        'tempRmpStr = tempRmpStr & "  | Avg " & String.Format("{0:  0.0}", Math.Round(currentStep.tempResult.AvgRamp, 1)) & vbNewLine
                    End If






                End If

                If checkToOutput.curePro.checkPressure And currentStep.pressureSet.SetPoint <> -1 Then
                    pressStr = "Pressure (psi)" & vbNewLine

                    If Math.Abs(currentStep.pressureSet.SetPoint) = Math.Abs(currentStep.pressureSet.NegTol) Then
                        pressStr = pressStr & "Max " & String.Format("{0:0.0}", Math.Round(currentStep.pressureResult.Max, 1)) & vbNewLine
                    ElseIf Math.Abs(currentStep.pressureSet.SetPoint) = Math.Abs(currentStep.pressureSet.PosTol) Then
                        pressStr = "Min " & String.Format("{0:0.0}", Math.Round(currentStep.pressureResult.Min, 1)) & vbNewLine
                    Else
                        pressStr = pressStr & "Max " & String.Format("{0:0.0}", Math.Round(currentStep.pressureResult.Max, 1))
                        pressStr = pressStr & " | Min " & String.Format("{0:0.0}", Math.Round(currentStep.pressureResult.Min, 1))
                        pressStr = pressStr & " | Avg " & String.Format("{0:0.0}", Math.Round(currentStep.pressureResult.Avg, 1)) & vbNewLine
                    End If

                    If currentStep.pressureSet.RampRate <> 0 Then
                        pressRmpStr = "Pressure Ramp (psi/min)" & vbNewLine
                        pressRmpStr = pressRmpStr & "Max " & String.Format("{0:0.0}", Math.Round(currentStep.pressureResult.MaxRamp, 1))
                        pressRmpStr = pressRmpStr & " | Min " & String.Format("{0:0.0}", Math.Round(currentStep.pressureResult.MinRamp, 1)) & vbNewLine
                        'pressRmpStr = pressRmpStr & "  | Avg " & String.Format("{0:0.0}", Math.Round(currentStep.pressureResult.AvgRamp, 1)) & vbNewLine
                    End If
                End If

                If checkToOutput.curePro.checkVac And currentStep.vacSet.SetPoint <> -1 Then
                    vacStr = "Vacuum (inHg)" & vbNewLine

                    If Math.Abs(currentStep.vacSet.SetPoint) = Math.Abs(currentStep.vacSet.NegTol) Then
                        vacStr = vacStr & "Min " & String.Format("{0:0.0}", Math.Round(currentStep.vacResult.Min, 1)) & vbNewLine
                    ElseIf Math.Abs(currentStep.vacSet.SetPoint) = Math.Abs(currentStep.vacSet.PosTol) Then
                        vacStr = vacStr & "Max " & String.Format("{0:0.0}", Math.Round(currentStep.vacResult.Max, 1)) & vbNewLine
                    Else
                        vacStr = vacStr & "Max " & String.Format("{0:0.0}", Math.Round(currentStep.vacResult.Min, 1))
                        vacStr = vacStr & " | Min " & String.Format("{0:0.0}", Math.Round(currentStep.vacResult.Max, 1))
                        vacStr = vacStr & " | Avg " & String.Format("{0:0.0}", Math.Round(currentStep.vacResult.Avg, 1)) & vbNewLine
                    End If

                End If

                cell3 = cell3 & tempStr & tempRmpStr & pressStr & pressRmpStr & vacStr

                Dim strTotalSoak As String = ""
                If currentStep.soakStepPass Then

                    'Dim timeReq As Double = 0

                    'If currentStep.termCond1Type = "Time" Then
                    '    timeReq = currentStep.termCond1Goal
                    'ElseIf currentStep.termCond2Type = "Time" Then
                    '    If currentStep.termCond1Goal > timeReq Then
                    '        timeReq = currentStep.termCond2Goal
                    '    End If

                    'End If

                    strTotalSoak = "Cumulative Time In Tolerance" & vbNewLine & "Exceeds Requirement"
                    cell3 = cell3 & strTotalSoak & vbNewLine
                End If

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

                If currentStep.tempSet.RampRate <> 0 Then
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

                If currentStep.pressureSet.RampRate <> 0 Then
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

                If currentStep.soakStepPass Then
                    formatFont(mainSheet.Cells(curRow, 7), strTotalSoak, 9,, True, True)
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
        outputDataToColumn(dataSheet, curCol, "Time " & vbNewLine & "Stamp" & vbNewLine & "(DateTime)", arr2Str(checkToOutput.dateArr, checkToOutput), checkToOutput.cureStart, checkToOutput.cureEnd)
        outputDataToColumn(dataSheet, curCol, "Cure Time" & vbNewLine & "(min)", arr2Str(checkToOutput.dateArr, checkToOutput,, True), checkToOutput.cureStart, checkToOutput.cureEnd)
        setChartX(mainSheet, checkToOutput)

        If checkToOutput.curePro.checkTemp Then
            If checkToOutput.machType = "Autoclave" Then
                outputDataToColumn(dataSheet, curCol, "Air TC" & vbNewLine & "(°F)", arr2Str(checkToOutput.vessel_TC.values, checkToOutput, 1), checkToOutput.cureStart, checkToOutput.cureEnd)
                'Adds air Tc to the chart, seems cluttered. Add back in ny uncommenting
                'Call addToChart(mainSheet, dataSheet, curCol - 1, "Air TC (°F)", checkToOutput.vessel_TC, 1)
            End If

            If UBound(checkToOutput.usrRunTC) > 0 Then
                outputDataToColumn(dataSheet, curCol, "Lead TC" & vbNewLine & "(°F)", arr2Str(checkToOutput.leadTC.values, checkToOutput, 1), checkToOutput.cureStart, checkToOutput.cureEnd)
                Call addToChart(checkToOutput, mainSheet, dataSheet, curCol - 1, "Lead TC (°F)", checkToOutput.leadTC, 1,, Color.Red)
                outputDataToColumn(dataSheet, curCol, "Lag TC" & vbNewLine & "(°F)", arr2Str(checkToOutput.lagTC.values, checkToOutput, 1), checkToOutput.cureStart, checkToOutput.cureEnd)
                Call addToChart(checkToOutput, mainSheet, dataSheet, curCol - 1, "Lag TC (°F)", checkToOutput.lagTC, 1,, Color.Blue)
            End If

            For i = 0 To UBound(checkToOutput.partTC_Arr)
                Dim z As Integer = 0
                For z = 0 To UBound(checkToOutput.usrRunTC)
                    If checkToOutput.partTC_Arr(i).Number = checkToOutput.usrRunTC(z) Then
                        outputDataToColumn(dataSheet, curCol, "TC " & checkToOutput.partTC_Arr(i).Number & vbNewLine & "(°F)", arr2Str(checkToOutput.partTC_Arr(i).values, checkToOutput, 1), checkToOutput.cureStart, checkToOutput.cureEnd)
                        outputDataToColumn(dataSheet, curCol, "TC " & checkToOutput.partTC_Arr(i).Number & " - Ramp" & vbNewLine & "(°F/min)", arr2Str(checkToOutput.partTC_Arr(i).ramp, checkToOutput, 2), checkToOutput.cureStart, checkToOutput.cureEnd)
                        If UBound(checkToOutput.usrRunTC) = 0 Then
                            Call addToChart(checkToOutput, mainSheet, dataSheet, curCol - 2, "TC -" & checkToOutput.partTC_Arr(i).Number & "- (°F)", checkToOutput.partTC_Arr(i), 1,, Color.Blue)
                        End If
                    End If
                Next
            Next
        End If

        If checkToOutput.curePro.checkPressure Then
            outputDataToColumn(dataSheet, curCol, "Pressure" & vbNewLine & "(psi)", arr2Str(checkToOutput.vesselPress.values, checkToOutput, 1), checkToOutput.cureStart, checkToOutput.cureEnd)
            Call addToChart(checkToOutput, mainSheet, dataSheet, curCol - 1, "Pressure (psi)", checkToOutput.vesselPress, 2, True, Color.Gray)
            outputDataToColumn(dataSheet, curCol, "Pressure Ramp" & vbNewLine & "(psi/min)", arr2Str(checkToOutput.vesselPress.ramp, checkToOutput, 2), checkToOutput.cureStart, checkToOutput.cureEnd)
        End If

        If checkToOutput.curePro.checkVac Then
            If UBound(checkToOutput.usrRunVac) > 0 Then
                outputDataToColumn(dataSheet, curCol, "Max Vac" & vbNewLine & "(inHg)", arr2Str(checkToOutput.maxVac.values, checkToOutput, 1), checkToOutput.cureStart, checkToOutput.cureEnd)
                'Adds max vac to the chart, never relavent
                'Call addToChart(mainSheet, dataSheet, curCol - 1, "Max Vacuum (inHg)", checkToOutput.maxVac, 2)
                outputDataToColumn(dataSheet, curCol, "Min Vac" & vbNewLine & "(inHg)", arr2Str(checkToOutput.minVac.values, checkToOutput, 1), checkToOutput.cureStart, checkToOutput.cureEnd)
                Call addToChart(checkToOutput, mainSheet, dataSheet, curCol - 1, "Min Vacuum (inHg)", checkToOutput.minVac, 2,, Color.Green)
            End If

            For i = 0 To UBound(checkToOutput.vac_Arr)
                Dim z As Integer = 0
                For z = 0 To UBound(checkToOutput.usrRunVac)
                    If checkToOutput.vac_Arr(i).Number = checkToOutput.usrRunVac(z) Then
                        outputDataToColumn(dataSheet, curCol, "Vac Port " & checkToOutput.vac_Arr(i).Number & vbNewLine & "(inHg)", arr2Str(checkToOutput.vac_Arr(i).values, checkToOutput, 1), checkToOutput.cureStart, checkToOutput.cureEnd)
                        If UBound(checkToOutput.usrRunVac) = 0 Then
                            Call addToChart(checkToOutput, mainSheet, dataSheet, curCol - 1, "Vac Port -" & checkToOutput.vac_Arr(i).Number & "- (inHg)", checkToOutput.vac_Arr(i), 2,, Color.Green)
                        End If
                    End If
                Next

            Next
        End If

        dataSheet.Range(dataSheet.Cells(2, 2), dataSheet.Cells(checkToOutput.dataCnt, curCol)).NumberFormat = "0.0"
        dataSheet.UsedRange.Columns.AutoFit()
        dataSheet.UsedRange.Rows.AutoFit()
        dataSheet.UsedRange.HorizontalAlignment = 3

        For i = 0 To UBound(checkToOutput.curePro.CureSteps) - 1
            Dim stepTime As Double = (checkToOutput.dateArr(checkToOutput.curePro.CureSteps(i).stepEnd) - checkToOutput.dateArr(checkToOutput.cureStart)).TotalMinutes
            addStepChart(mainSheet, stepTime)
        Next


        mainSheet.Activate()

        SwitchOff(Excel, False)

        'Excel.Visible = True

        If Not My.Computer.FileSystem.DirectoryExists(outputPath) Then
            outputPath = My.Computer.FileSystem.SpecialDirectories.Desktop
        End If

        If Not My.Computer.FileSystem.DirectoryExists(outputPath & "\CureReports") Then
            My.Computer.FileSystem.CreateDirectory(outputPath & "\CureReports")
        End If

        If Not My.Computer.FileSystem.DirectoryExists(outputPath & "\CureReports\ExcelReports") Then
            My.Computer.FileSystem.CreateDirectory(outputPath & "\CureReports\ExcelReports")
        End If

        'Dim postNameChange As String = ""
        'If InStr(checkToOutput.curePro.Name, "POST") Then
        '    postNameChange = "Post"
        'End If

        Dim cureAcro As String = strToAcronym(checkToOutput.curePro.Name)

        Dim mismatchStr As String = ""

        If My.Computer.FileSystem.FileExists(outputPath & "\CureReports\ExcelReports\CureReport_" & cureAcro & "_" & checkToOutput.JobNum & ".xlsx") Or My.Computer.FileSystem.FileExists(outputPath & "\CureReports\CureReport_" & cureAcro & "_" & checkToOutput.JobNum & ".pdf") Then
            Dim idNum As Integer = 1
            Do While mismatchStr = ""
                If Not My.Computer.FileSystem.FileExists(outputPath & "\CureReports\ExcelReports\CureReport_" & cureAcro & "_" & checkToOutput.JobNum & "(" & idNum & ")" & ".xlsx") And Not My.Computer.FileSystem.FileExists(outputPath & "\CureReports\CureReport_" & cureAcro & "_" & checkToOutput.JobNum & "(" & idNum & ")" & ".pdf") Then
                    mismatchStr = "(" & idNum & ")"
                End If
                idNum += 1
            Loop
        End If

        Excel.DisplayAlerts = False
        Excel.ActiveWorkbook.SaveAs(outputPath & "\CureReports\ExcelReports\CureReport_" & cureAcro & "_" & checkToOutput.JobNum & mismatchStr & ".xlsx", 51,,, True,,,,)
        Excel.DisplayAlerts = True

        If checkToOutput.curePro.curePass Then
            mainSheet.ExportAsFixedFormat(0, outputPath & "\CureReports\CureReport_" & cureAcro & "_" & checkToOutput.JobNum & mismatchStr, 0,,,,, False,)
        Else
            mainSheet.ExportAsFixedFormat(0, outputPath & "\CureReports\CureReport_" & cureAcro & "_" & checkToOutput.JobNum & mismatchStr, 0,,,,, True,)
        End If

        GC.Collect()
        GC.WaitForPendingFinalizers()

        Excel.Quit()

        FinalReleaseComObject(mainSheet)
        FinalReleaseComObject(userSheet)
        FinalReleaseComObject(dataSheet)
        FinalReleaseComObject(Excel)

        killExcel(excelProcID)
    End Sub

    Sub killExcel(procIdToKill As Integer)
        Dim procs() = Process.GetProcessesByName("EXCEL")

        For Each runExcel In procs
            If runExcel.Id = procIdToKill Then
                runExcel.Kill()
            End If
        Next

        excelProcID = 0
    End Sub

    Function strToAcronym(inStr As String) As String
        Dim retAcro As String = ""

        inStr = inStr.Trim()

        Dim strVals() As String = Split(inStr, " ")

        For i = 0 To UBound(strVals)
            If System.Text.RegularExpressions.Regex.IsMatch(strVals(i), "^[a-zA-Z]+$") Then
                retAcro = retAcro & Strings.UCase(Strings.Left(strVals(i), 1))
            Else
                'If i <> 0 Then
                '    retAcro = retAcro & "-"
                'End If

                retAcro = retAcro & "(" & strVals(i) & ")"

                'If i <> UBound(strVals) Then
                '    retAcro = retAcro & "-"
                'End If
            End If

        Next

        Return retAcro

    End Function

    Sub runInfoOutput(infoSht As Excel.Worksheet, checkToOutput As CureCheck)

        Dim cRow As Integer = 1

        outputExcelVal(infoSht, "RunDate", Now, cRow)
        outputExcelVal(infoSht, "comp", Environment.MachineName, cRow)
        outputExcelVal(infoSht, "compUser", Environment.UserName, cRow)

        outputExcelVal(infoSht, "JobNum", checkToOutput.JobNum, cRow)
        outputExcelVal(infoSht, "PONum", checkToOutput.PONum, cRow)
        outputExcelVal(infoSht, "PartNum", checkToOutput.PartNum, cRow)
        outputExcelVal(infoSht, "PartRev", checkToOutput.PartRev, cRow)
        outputExcelVal(infoSht, "PartNom", checkToOutput.PartNom, cRow)
        outputExcelVal(infoSht, "ProgramNum", checkToOutput.ProgramNum, cRow)
        outputExcelVal(infoSht, "PartQty", checkToOutput.PartQty, cRow)
        outputExcelVal(infoSht, "DataPath", checkToOutput.DataPath, cRow)

        outputExcelVal(infoSht, "equipSerialNum", checkToOutput.equipSerialNum, cRow)
        outputExcelVal(infoSht, "dataCnt", checkToOutput.dataCnt, cRow)
        outputExcelVal(infoSht, "headerRow", checkToOutput.headerRow, cRow)
        outputExcelVal(infoSht, "headerCount", checkToOutput.headerCount, cRow)
        outputExcelVal(infoSht, "cureStart", checkToOutput.cureStart, cRow)
        outputExcelVal(infoSht, "cureEnd", checkToOutput.cureEnd, cRow)
        outputExcelVal(infoSht, "startTime", checkToOutput.cureStartTime, cRow)
        outputExcelVal(infoSht, "endTime", checkToOutput.cureEndTime, cRow)
        outputExcelVal(infoSht, "machType", checkToOutput.machType, cRow)

        If checkToOutput.curePro.checkTemp Then
            For i = 0 To UBound(checkToOutput.usrRunTC)
                outputExcelVal(infoSht, "checkToOutput.usrRunTC", checkToOutput.usrRunTC(i), cRow)
            Next
        End If

        If checkToOutput.curePro.checkVac Then
            For i = 0 To UBound(checkToOutput.usrRunVac)
                outputExcelVal(infoSht, "checkToOutput.usrRunVac", checkToOutput.usrRunVac(i), cRow)
            Next
        End If


        outputExcelVal(infoSht, "cureProfileDate", checkToOutput.curePro.fileEditDate, cRow)

        outputExcelVal(infoSht, "checkToOutput.curePro", checkToOutput.curePro.SerializeCure, cRow)



    End Sub

    Sub outputExcelVal(sheet As Excel.Worksheet, inDesc As String, inVal As String, ByRef cRow As Integer)

        sheet.Cells(cRow, 1) = inDesc
        sheet.Cells(cRow, 2) = inVal
        cRow += 1
    End Sub

    Function plusMinusVal(plusVal As Double, minusVal As Double) As String
        Dim optionNeg As String = ""

        If minusVal = 0 Then
            optionNeg = "-"
        Else
            optionNeg = ""
        End If

        If Math.Abs(plusVal) = Math.Abs(minusVal) Then
            Return "±" & Math.Abs(plusVal)
        Else
            Return "+" & plusVal & "/" & optionNeg & minusVal
        End If

    End Function

    Sub clearChart(mainSheet As Excel.Worksheet)
        Dim mainChart As Excel.Chart = mainSheet.ChartObjects.Item("DataChart").Chart
        Dim mainChartSeriesCollect As Excel.SeriesCollection = mainChart.SeriesCollection

        For i = mainChartSeriesCollect.Count To 1 Step -1
            mainChartSeriesCollect.Item(i).Delete()
        Next

        mainChart.Axes(2, 1).MaximumScale = 140
        mainChart.Axes(2, 1).MinimumScale = 100

    End Sub

    Sub setChartX(mainSheet As Excel.Worksheet, checkToOutput As CureCheck)
        Dim mainChart As Excel.Chart = mainSheet.ChartObjects.Item("DataChart").Chart
        Dim mainChartSeriesCollect As Excel.SeriesCollection = mainChart.SeriesCollection

        mainChart.Axes(1).MinimumScale = 0
        mainChart.Axes(1).MaximumScale = Math.Round((checkToOutput.dateArr(checkToOutput.cureEnd) - checkToOutput.dateArr(checkToOutput.cureStart)).TotalMinutes, 0)
    End Sub

    Sub addToChart(checkToOutput As CureCheck, mainSheet As Excel.Worksheet, dataSheet As Excel.Worksheet, cNum As Integer, seriesName As String, startDataSet As DataSet, Optional axisGroup As Integer = 1, Optional lineDashed As Boolean = False, Optional lineColor As Color = Nothing)
        Dim mainChart As Excel.Chart = mainSheet.ChartObjects.Item("DataChart").Chart
        Dim mainChartSeriesCollect As Excel.SeriesCollection = mainChart.SeriesCollection

        Dim cureTimesRng As Excel.Range = dataSheet.Range(dataSheet.Cells(2, 2), dataSheet.Cells(checkToOutput.cureEnd - checkToOutput.cureStart + 2, 2))
        Dim cureValRng As Excel.Range = dataSheet.Range(dataSheet.Cells(2, cNum), dataSheet.Cells(checkToOutput.cureEnd - checkToOutput.cureStart + 2, cNum))

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

        If lineDashed Then
            ser.Format.Line.DashStyle = 7 'Microsoft.Office.Core.MsoLineDashStyle.msoLineLongDash
        End If

        If Not lineColor = Nothing Then
            ser.Format.Line.ForeColor.RGB = RGB(lineColor.R, lineColor.G, lineColor.B)
        End If

        Dim valMax As Double = startDataSet.Max()

        If valMax > mainChart.Axes(2, axisGroup).MaximumScale Then
            mainChart.Axes(2, axisGroup).MaximumScale = valMax + ((valMax - 100) * 0.05)
        End If

        If axisGroup = 2 Then
            mainChart.Axes(2, axisGroup).HasTitle = True
            mainChart.Axes(2, axisGroup).AxisTitle.Text = "Pressure (psi) | Vacuum (inHg)"
            mainChart.Axes(2, axisGroup).MinimumScale = -30
            mainChart.Axes(2, axisGroup).MaximumScale = 120
            mainChart.Axes(2, axisGroup).MajorUnit = 10
        End If


    End Sub

    Sub addStepChart(mainSheet As Excel.Worksheet, intime As Double)
        Dim mainChart As Excel.Chart = mainSheet.ChartObjects.Item("StepChart").Chart
        Dim mainChartSeriesCollect As Excel.SeriesCollection = mainChart.SeriesCollection

        Dim dataChart As Excel.Chart = mainSheet.ChartObjects.Item("dataChart").Chart

        mainChart.Axes(1).MinimumScale = dataChart.Axes(1).MinimumScale
        mainChart.Axes(1).MaximumScale = dataChart.Axes(1).MaximumScale

        mainChart.Axes(2).MinimumScale = dataChart.Axes(2).MinimumScale
        mainChart.Axes(2).MaximumScale = dataChart.Axes(2).MaximumScale

        Dim ser As Excel.Series
        ser = mainChartSeriesCollect.NewSeries

        Dim avgVal As Double = (mainChart.Axes(2, 1).MinimumScale + mainChart.Axes(2, 1).MaximumScale) / 2

        ser.Values = avgVal
        ser.XValues = intime
        ser.Name = ""


        ser.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone
        ser.Format.Line.Visible = False

        ser.HasErrorBars = True
        ser.ErrorBars.EndStyle = Excel.XlEndStyleCap.xlNoCap

        ser.ErrorBar(Excel.XlErrorBarDirection.xlY, Excel.XlErrorBarInclude.xlErrorBarIncludeBoth, Excel.XlErrorBarType.xlErrorBarTypePercent, 110)
        ser.ErrorBar(Excel.XlErrorBarDirection.xlX, Excel.XlErrorBarInclude.xlErrorBarIncludeNone, Excel.XlErrorBarType.xlErrorBarTypeFixedValue)

        mainSheet.Shapes.Item("StepChart").ZOrder(1) 'Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack)

    End Sub


    Function arr2Str(Of T)(inArr() As T, checkToOutput As CureCheck, Optional rndDecimals As Integer = 1, Optional calcDuration As Boolean = False) As String()
        Dim retArr() As String = Nothing

        Dim i As Integer = 0
        For i = 0 To UBound(inArr)

            Dim inValue As String = ""

            If TypeName(inArr) = "Date()" And calcDuration Then
                Dim startDate As Object = inArr(checkToOutput.cureStart)
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

    Function termToStr(termCondType As String, termCondCondition As String, termCondGoal As Double, termCondModifier As Object, currentStep As CureStep) As String

        Dim currentStr As String = ""

        If termCondType = "None" Then
            Return ""


        ElseIf termCondType = "Time" Then
            If termCondModifier = "Recieve" Then
                currentStr = currentStr & "Time continued from previous step."
            Else
                currentStr = currentStr & "After " & termCondGoal & " min"
            End If

            currentStr = currentStr & vbNewLine


        ElseIf termCondType = "Temp" Then
            currentStr = currentStr & termCondModifier

            If termCondCondition = "GREATER" Then
                currentStr = currentStr & " > "
            ElseIf termCondCondition = "LESS" Then
                currentStr = currentStr & " < "
            Else
                Throw New Exception("Unknown termination condition for temp on " & currentStep.stepName)
            End If

            currentStr = currentStr & termCondGoal & "°F"
            currentStr = currentStr & vbNewLine


        ElseIf termCondType = "Press" Then
            If termCondCondition = "GREATER" Then
                currentStr = currentStr & " > "
            ElseIf termCondCondition = "LESS" Then
                currentStr = currentStr & " < "
            Else
                Throw New Exception("Unknown termination condition for temp on " & currentStep.stepName)
            End If

            currentStr = currentStr & termCondGoal & " psi"
            currentStr = currentStr & vbNewLine
        End If

        Return currentStr
    End Function

End Class
