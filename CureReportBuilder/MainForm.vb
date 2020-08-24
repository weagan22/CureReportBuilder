Option Explicit On

Imports System.Runtime.CompilerServices
Imports Microsoft.Office.Interop

Public Class MainForm

    Public cureProfiles() As CureProfile
    Dim curePro As CureProfile = New CureProfile("null")

    Public autoclaveSNum As New Dictionary(Of String, Integer) From {{"Controller", 2601}, {"Air TC1", 2600}, {"Air TC2", 2603}, {"PT", 2599}, {"VT1", 2593}, {"VT2", 2598}, {"VT3", 2595}, {"VT4", 2594}, {"VT5", 2597}, {"VT6", 2596}}

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

    Dim STOP_Txt_FilePath_TextChanged As Boolean = False


    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call errorReset()

        TabControl1.TabPages.Remove(TabPage3)


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
        Box_RunLine.Enabled = False
    End Sub

    Sub testRun()

        partValues("JobNum") = Txt_JobNumber.Text
        partValues("PONum") = ""
        partValues("PartNum") = Txt_PartNumber.Text
        partValues("PartRev") = Txt_Revision.Text
        partValues("PartNom") = Txt_PartDesc.Text
        partValues("ProgramNum") = Txt_ProgramNumber.Text
        partValues("PartQty") = Txt_Qty.Text
        partValues("DataPath") = Txt_FilePath.Text

        If machType = "Omega" Then
            equipSerialNum = Txt_DataRecorder.Text
        End If



        Dim rowCnt As Integer = 1

        If machType = "Autoclave" Then
            addToEquip(equipSerialNum, "Controller", rowCnt)
            addToEquip(equipSerialNum, "Air TC1", rowCnt)
            addToEquip(equipSerialNum, "Air TC2", rowCnt)
            addToEquip(equipSerialNum, "PT", rowCnt)
        End If

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

            For i = 0 To UBound(usrRunVac)
                addToEquip(equipSerialNum, "VT" & usrRunVac(i), rowCnt)
            Next
        End If

        If machType = "Autoclave" Then
            If Strings.Right(equipSerialNum, 3) = " | " Then
                equipSerialNum = Strings.Left(equipSerialNum, Len(equipSerialNum) - 3)
            ElseIf Strings.Right(equipSerialNum, 1) = vbCrLf Then
                equipSerialNum = Strings.Left(equipSerialNum, Len(equipSerialNum) - 1)
            End If
        End If

        Call runCalc()

        Call outputResults()
    End Sub

    Sub addToEquip(ByRef equipSerialNum As String, name As String, ByRef rowCnt As Integer)

        rowCnt += 1

        If rowCnt = 2 Then
            rowCnt = 0
            equipSerialNum = equipSerialNum & name & ": S" & autoclaveSNum(name) & vbNewLine
        Else
            equipSerialNum = equipSerialNum & name & ": S" & autoclaveSNum(name) & " | "
        End If

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
        Excel.Visible = True

        Try
            Excel.Workbooks.Open(Txt_TemplatePath.Text)
        Catch ex As Exception
            MsgBox("Failed to open Excel Report Template.")
            Exit Sub
        End Try

        'SwitchOff(Excel, True)

        Dim mainSheet As Excel.Worksheet = Excel.Sheets.Item(1)
        Dim userSheet As Excel.Worksheet = Excel.Sheets.Item(2)
        Dim dataSheet As Excel.Worksheet = Excel.Sheets.Item(3)

        runInfoOutput(userSheet)

        mainSheet.Cells(2, 1) = "Job" & vbNewLine & partValues("JobNum") & vbNewLine & "Program" & vbNewLine & partValues("ProgramNum")
        formatFont(mainSheet.Cells(2, 1), "Job", 14, True, False, True)
        formatFont(mainSheet.Cells(2, 1), "Program", 14, True, False, True)

        mainSheet.Cells(2, 4) = "Part" & vbNewLine & partValues("PartNum") & vbNewLine & "Rev. " & partValues("PartRev") & vbNewLine & partValues("PartNom") & vbNewLine & "Qty: " & partValues("PartQty")
        formatFont(mainSheet.Cells(2, 4), "Part", 14, True, False, True)


        Dim hours As Integer = Math.Round((dateArr(cureEnd) - dateArr(cureStart)).TotalMinutes, 1) \ 60
        Dim minutes As Integer = Math.Round((dateArr(cureEnd) - dateArr(cureStart)).TotalMinutes, 1) - (hours * 60)

        mainSheet.Cells(2, 8) = "Cure" & vbNewLine & "Start: " & dateArr(cureStart).ToString("dd-MMM-yyyy | h:mm tt") & vbNewLine & "Finish: " & dateArr(cureEnd).ToString("dd-MMM-yyyy | h:mm tt") & vbNewLine & "Duration: " & hours & "h:" & minutes & "m" 'Math.Round((dateArr(cureEnd) - dateArr(cureStart)).TotalMinutes, 1) & " min"
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

            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 1), mainSheet.Cells(curRow, 10)).HorizontalAlignment = 3

            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 1), mainSheet.Cells(curRow, 10)).VerticalAlignment = -4108

            Excel.ActiveSheet.Range(mainSheet.Cells(curRow, 1), mainSheet.Cells(curRow, 10)).Borders(9).LineStyle = 1

            'Fill out data in cell 1
            If currentStep.hardFail Then
                mainSheet.Cells(curRow, 1) = currentStep.stepName
                formatFont(mainSheet.Cells(curRow, 1), currentStep.stepName, 14, True, False, False)
            Else
                Dim stepPass As String = "FAIL"
                If currentStep.stepPass Then stepPass = "PASS"

                Dim startTime As Double = (dateArr(currentStep.stepStart) - dateArr(cureStart)).TotalMinutes
                Dim endTime As Double = (dateArr(currentStep.stepEnd) - dateArr(cureStart)).TotalMinutes

                Dim endStr As String = ""
                If currentStep.stepTerminate = False Then
                    endStr = "Failed to Terminate"
                Else
                    endStr = "Finish: " & endTime.ToString("0.0") & " min"
                End If


                mainSheet.Cells(curRow, 1) = currentStep.stepName & vbNewLine & "(" & stepPass & ")" & vbNewLine & "Start: " & startTime.ToString("0.0") & " min" & vbNewLine & endStr

                formatFont(mainSheet.Cells(curRow, 1), currentStep.stepName, 14, True, False, False)

                If currentStep.stepPass = True Then
                    formatFont(mainSheet.Cells(curRow, 1), "(" & stepPass & ")", 11, False, False, False)
                Else
                    formatFont(mainSheet.Cells(curRow, 1), "(" & stepPass & ")", 11, False, False, False, Color.Red)
                End If

                formatFont(mainSheet.Cells(curRow, 1), "Start: " & startTime.ToString("0.0") & " min", 9, False, False, False)
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



            If curePro.checkTemp And currentStep.tempSet("SetPoint") <> -1 Then
                If Math.Abs(currentStep.tempSet("SetPoint")) = Math.Abs(currentStep.tempSet("NegTol")) Then
                    tempStr = "Temperature: Max " & currentStep.tempSet("SetPoint") + currentStep.tempSet("PosTol") & "°F"
                ElseIf Math.Abs(currentStep.tempSet("SetPoint")) = Math.Abs(currentStep.tempSet("PosTol")) Then
                    tempStr = "Temperature: Min " & currentStep.tempSet("SetPoint") + currentStep.tempSet("NegTol") & "°F"
                Else

                    tempStr = "Temperature: " & currentStep.tempSet("SetPoint") & "°F " & plusMinusVal(currentStep.tempSet("PosTol"), currentStep.tempSet("NegTol")) & "°F"
                End If
                tempStr = tempStr & vbNewLine

                If currentStep.tempSet("RampRate") <> 0 Then
                    tempRmpStr = "Temp. Ramp: " & currentStep.tempSet("RampRate") & "°F/min " & plusMinusVal(currentStep.tempSet("RampPosTol"), currentStep.tempSet("RampNegTol")) & "°F/min"
                    tempRmpStr = tempRmpStr & vbNewLine
                End If
            End If

            If curePro.checkPressure And currentStep.pressureSet("SetPoint") <> -1 Then
                If Math.Abs(currentStep.pressureSet("SetPoint")) = Math.Abs(currentStep.pressureSet("NegTol")) Then
                    pressStr = "Pressure: Max " & currentStep.pressureSet("SetPoint") + currentStep.pressureSet("PosTol") & " psi"
                ElseIf Math.Abs(currentStep.pressureSet("SetPoint")) = Math.Abs(currentStep.pressureSet("PosTol")) Then
                    pressStr = "Pressure Min " & currentStep.pressureSet("SetPoint") + currentStep.pressureSet("NegTol") & " psi"
                Else
                    pressStr = pressStr & "Pressure: " & currentStep.pressureSet("SetPoint") & " psi " & plusMinusVal(currentStep.pressureSet("PosTol"), currentStep.pressureSet("NegTol")) & " psi"
                End If
                pressStr = pressStr & vbNewLine

                If currentStep.pressureSet("RampRate") <> 0 Then
                    pressRmpStr = pressRmpStr & "Pressure Ramp: " & currentStep.pressureSet("RampRate") & " psi/min " & plusMinusVal(currentStep.pressureSet("RampPosTol"), currentStep.pressureSet("RampNegTol")) & " psi/min"
                    pressRmpStr = pressRmpStr & vbNewLine
                End If
            End If

            If curePro.checkVac And currentStep.vacSet("SetPoint") <> -1 Then
                If Math.Abs(currentStep.vacSet("SetPoint")) = Math.Abs(currentStep.vacSet("NegTol")) Then
                    vacStr = "Vacuum: Min " & currentStep.vacSet("SetPoint") + currentStep.vacSet("PosTol") & " inHg"
                ElseIf Math.Abs(currentStep.vacSet("SetPoint")) = Math.Abs(currentStep.vacSet("PosTol")) Then
                    vacStr = "Vacuum: Max " & currentStep.vacSet("SetPoint") + currentStep.vacSet("NegTol") & " inHg"
                Else
                    vacStr = vacStr & "Vacuum: " & currentStep.vacSet("SetPoint") & " inHg " & plusMinusVal(currentStep.vacSet("PosTol"), currentStep.vacSet("NegTol")) & " inHg"
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

                If curePro.checkTemp And currentStep.tempSet("SetPoint") <> -1 Then
                    tempStr = "Temperature (°F)" & vbNewLine






                    If Math.Abs(currentStep.tempSet("SetPoint")) = Math.Abs(currentStep.tempSet("NegTol")) Then
                        tempStr = tempStr & "Max: " & String.Format("{0}", Math.Round(currentStep.tempResult("Max"), 0)) & vbNewLine
                    ElseIf Math.Abs(currentStep.tempSet("SetPoint")) = Math.Abs(currentStep.tempSet("PosTol")) Then
                        tempStr = tempStr & "Min: " & String.Format("{0}", Math.Round(currentStep.tempResult("Min"), 0)) & vbNewLine
                    Else
                        tempStr = tempStr & "Max: " & String.Format("{0}", Math.Round(currentStep.tempResult("Max"), 0))
                        tempStr = tempStr & " | Min: " & String.Format("{0}", Math.Round(currentStep.tempResult("Min"), 0))
                        tempStr = tempStr & "  | Avg: " & String.Format("{0}", Math.Round(currentStep.tempResult("Avg"), 0)) & vbNewLine
                    End If

                    If currentStep.tempSet("RampRate") <> 0 Then
                        tempRmpStr = "Temp. Ramp (°F/min)" & vbNewLine
                        tempRmpStr = tempRmpStr & "Max: " & String.Format("{0:0.0}", Math.Round(currentStep.tempResult("MaxRamp"), 1))
                        tempRmpStr = tempRmpStr & " | Min: " & String.Format("{0:0.0}", Math.Round(currentStep.tempResult("MinRamp"), 1))
                        tempRmpStr = tempRmpStr & "  | Avg: " & String.Format("{0:0.0}", Math.Round(currentStep.tempResult("AvgRamp"), 1)) & vbNewLine
                    End If






                End If

                If curePro.checkPressure And currentStep.pressureSet("SetPoint") <> -1 Then
                    pressStr = "Pressure (psi)" & vbNewLine

                    If Math.Abs(currentStep.pressureSet("SetPoint")) = Math.Abs(currentStep.pressureSet("NegTol")) Then
                        pressStr = pressStr & "Max: " & String.Format("{0:0.0}", Math.Round(currentStep.pressureResult("Max"), 1)) & vbNewLine
                    ElseIf Math.Abs(currentStep.pressureSet("SetPoint")) = Math.Abs(currentStep.pressureSet("PosTol")) Then
                        pressStr = "Min: " & String.Format("{0:0.0}", Math.Round(currentStep.pressureResult("Min"), 1)) & vbNewLine
                    Else
                        pressStr = pressStr & "Max: " & String.Format("{0:0.0}", Math.Round(currentStep.pressureResult("Max"), 1))
                        pressStr = pressStr & " | Min: " & String.Format("{0:0.0}", Math.Round(currentStep.pressureResult("Min"), 1))
                        pressStr = pressStr & " | Avg: " & String.Format("{0:0.0}", Math.Round(currentStep.pressureResult("Avg"), 1)) & vbNewLine
                    End If

                    If currentStep.pressureSet("RampRate") <> 0 Then
                        pressRmpStr = "Pressure Ramp (psi/min)" & vbNewLine
                        pressRmpStr = pressRmpStr & "Max: " & String.Format("{0:0.0}", Math.Round(currentStep.pressureResult("MaxRamp"), 1))
                        pressRmpStr = pressRmpStr & " | Min: " & String.Format("{0:0.0}", Math.Round(currentStep.pressureResult("MinRamp"), 1))
                        pressRmpStr = pressRmpStr & "  | Avg: " & String.Format("{0:0.0}", Math.Round(currentStep.pressureResult("AvgRamp"), 1)) & vbNewLine
                    End If
                End If

                If curePro.checkVac And currentStep.vacSet("SetPoint") <> -1 Then
                    vacStr = "Vacuum (inHg)" & vbNewLine

                    If Math.Abs(currentStep.vacSet("SetPoint")) = Math.Abs(currentStep.vacSet("NegTol")) Then
                        vacStr = vacStr & "Min: " & String.Format("{0:0.0}", Math.Round(currentStep.vacResult("Min"), 1)) & vbNewLine
                    ElseIf Math.Abs(currentStep.vacSet("SetPoint")) = Math.Abs(currentStep.vacSet("PosTol")) Then
                        vacStr = vacStr & "Max: " & String.Format("{0:0.0}", Math.Round(currentStep.vacResult("Max"), 1)) & vbNewLine
                    Else
                        vacStr = vacStr & "Max: " & String.Format("{0:0.0}", Math.Round(currentStep.vacResult("Min"), 1))
                        vacStr = vacStr & " | Min: " & String.Format("{0:0.0}", Math.Round(currentStep.vacResult("Max"), 1))
                        vacStr = vacStr & " | Avg: " & String.Format("{0:0.0}", Math.Round(currentStep.vacResult("Avg"), 1)) & vbNewLine
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
                'Adds air Tc to the chart, seems cluttered. Add back in ny uncommenting
                'Call addToChart(mainSheet, dataSheet, curCol - 1, "Air TC (°F)", vessel_TC, 1)
            End If

            If UBound(usrRunTC) > 0 Then
                outputDataToColumn(dataSheet, curCol, "Lead TC" & vbNewLine & "(°F)", arr2Str(leadTC.values, 1), cureStart, cureEnd)
                Call addToChart(mainSheet, dataSheet, curCol - 1, "Lead TC (°F)", leadTC, 1,, Color.Red)
                outputDataToColumn(dataSheet, curCol, "Lag TC" & vbNewLine & "(°F)", arr2Str(lagTC.values, 1), cureStart, cureEnd)
                Call addToChart(mainSheet, dataSheet, curCol - 1, "Lag TC (°F)", lagTC, 1,, Color.Blue)
            End If

            For i = 0 To UBound(partTC_Arr)
                Dim z As Integer = 0
                For z = 0 To UBound(usrRunTC)
                    If partTC_Arr(i).Number = usrRunTC(z) Then
                        outputDataToColumn(dataSheet, curCol, "TC " & partTC_Arr(i).Number & vbNewLine & "(°F)", arr2Str(partTC_Arr(i).values, 1), cureStart, cureEnd)
                        outputDataToColumn(dataSheet, curCol, "TC " & partTC_Arr(i).Number & " - Ramp" & vbNewLine & "(°F/min)", arr2Str(partTC_Arr(i).ramp, 2), cureStart, cureEnd)
                        If UBound(usrRunTC) = 0 Then
                            Call addToChart(mainSheet, dataSheet, curCol - 2, "TC -" & partTC_Arr(i).Number & "- (°F)", partTC_Arr(i), 1,, Color.Blue)
                        End If
                    End If
                Next
            Next
        End If

        If curePro.checkPressure Then
            outputDataToColumn(dataSheet, curCol, "Pressure" & vbNewLine & "(psi)", arr2Str(vesselPress.values, 1), cureStart, cureEnd)
            Call addToChart(mainSheet, dataSheet, curCol - 1, "Pressure (psi)", vesselPress, 2, True, Color.Gray)
            outputDataToColumn(dataSheet, curCol, "Pressure Ramp" & vbNewLine & "(psi/min)", arr2Str(vesselPress.ramp, 2), cureStart, cureEnd)
        End If

        If curePro.checkVac Then
            If UBound(usrRunVac) > 0 Then
                outputDataToColumn(dataSheet, curCol, "Max Vac" & vbNewLine & "(inHg)", arr2Str(maxVac.values, 1), cureStart, cureEnd)
                'Adds max vac to the chart, never relavent
                'Call addToChart(mainSheet, dataSheet, curCol - 1, "Max Vacuum (inHg)", maxVac, 2)
                outputDataToColumn(dataSheet, curCol, "Min Vac" & vbNewLine & "(inHg)", arr2Str(minVac.values, 1), cureStart, cureEnd)
                Call addToChart(mainSheet, dataSheet, curCol - 1, "Min Vacuum (inHg)", minVac, 2,, Color.Green)
            End If

            For i = 0 To UBound(vac_Arr)
                Dim z As Integer = 0
                For z = 0 To UBound(usrRunVac)
                    If vac_Arr(i).Number = usrRunVac(z) Then
                        outputDataToColumn(dataSheet, curCol, "Vac Port " & vac_Arr(i).Number & vbNewLine & "(inHg)", arr2Str(vac_Arr(i).values, 1), cureStart, cureEnd)
                        If UBound(usrRunVac) = 0 Then
                            Call addToChart(mainSheet, dataSheet, curCol - 1, "Vac Port -" & vac_Arr(i).Number & "- (inHg)", vac_Arr(i), 2,, Color.Green)
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

    Sub runInfoOutput(infoSht As Excel.Worksheet)

        Dim cRow As Integer = 1

        outputExcelVal(infoSht, "RunDate", Now, cRow)
        outputExcelVal(infoSht, "comp", Environment.MachineName, cRow)
        outputExcelVal(infoSht, "compUser", Environment.UserName, cRow)

        outputExcelVal(infoSht, "JobNum", partValues("JobNum"), cRow)
        outputExcelVal(infoSht, "PONum", partValues("PONum"), cRow)
        outputExcelVal(infoSht, "PartNum", partValues("PartNum"), cRow)
        outputExcelVal(infoSht, "PartRev", partValues("PartRev"), cRow)
        outputExcelVal(infoSht, "PartNom", partValues("PartNom"), cRow)
        outputExcelVal(infoSht, "ProgramNum", partValues("ProgramNum"), cRow)
        outputExcelVal(infoSht, "PartQty", partValues("PartQty"), cRow)
        outputExcelVal(infoSht, "DataPath", partValues("DataPath"), cRow)

        outputExcelVal(infoSht, "equipSerialNum", equipSerialNum, cRow)
        outputExcelVal(infoSht, "dataCnt", dataCnt, cRow)
        outputExcelVal(infoSht, "headerRow", headerRow, cRow)
        outputExcelVal(infoSht, "headerCount", headerCount, cRow)
        outputExcelVal(infoSht, "cureStart", cureStart, cRow)
        outputExcelVal(infoSht, "cureEnd", cureEnd, cRow)
        outputExcelVal(infoSht, "startTime", dateValues("startTime"), cRow)
        outputExcelVal(infoSht, "endTime", dateValues("endTime"), cRow)
        outputExcelVal(infoSht, "machType", machType, cRow)

        If curePro.checkTemp Then
            For i = 0 To UBound(usrRunTC)
                outputExcelVal(infoSht, "usrRunTC", usrRunTC(i), cRow)
            Next
        End If

        If curePro.checkVac Then
            For i = 0 To UBound(usrRunVac)
                outputExcelVal(infoSht, "usrRunVac", usrRunVac(i), cRow)
            Next
        End If


        outputExcelVal(infoSht, "cureProFileDate", cureProfiles(Combo_CureProfile.SelectedIndex).fileEditDate, cRow)

        outputExcelVal(infoSht, "curePro", curePro.serializeCure, cRow)



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

    Sub setChartX(mainSheet As Excel.Worksheet)
        Dim mainChart As Excel.Chart = mainSheet.ChartObjects.Item("DataChart").Chart
        Dim mainChartSeriesCollect As Excel.SeriesCollection = mainChart.SeriesCollection

        mainChart.Axes(1).MinimumScale = 0
        mainChart.Axes(1).MaximumScale = Math.Round((dateArr(cureEnd) - dateArr(cureStart)).TotalMinutes, 0)
    End Sub

    Sub addToChart(mainSheet As Excel.Worksheet, dataSheet As Excel.Worksheet, cNum As Integer, seriesName As String, startDataSet As DataSet, Optional axisGroup As Integer = 1, Optional lineDashed As Boolean = False, Optional lineColor As Color = Nothing)
        Dim mainChart As Excel.Chart = mainSheet.ChartObjects.Item("DataChart").Chart
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
            Dim goal As Double = -1
            Dim greaterThanGoal As Boolean = True
            Dim holder As Double = 0

            Dim total As Double = 0
            Dim addCnt As Integer = 0

            Dim stepDuration As Double = (dateArr(UBound(dateArr)) - dateArr(0)).TotalMinutes / dataCnt

            'Calculate temp results for a given step
            If curePro.checkTemp = True And curePro.CureSteps(i).tempSet("SetPoint") <> -1 Then

                ''Max temp
                If Not firstStep Then
                    If curePro.CureSteps(i - 1).tempSet("RampRate") < 0 Then
                        indexStart = indexStart + Math.Abs(Math.Round(curePro.CureSteps(i - 1).tempSet("RampRate") / stepDuration, 0))
                    End If
                End If

                If Not lastStep Then
                    If curePro.CureSteps(i + 1).tempSet("RampRate") > 0 Then
                        indexEnd = indexEnd - Math.Abs(Math.Round(curePro.CureSteps(i + 1).tempSet("RampRate") / stepDuration, 0))
                    End If
                End If

                ''Handles a very short step: just look at max of whole step, no buffer
                If indexEnd <= indexStart Then
                    indexStart = curePro.CureSteps(i).stepStart
                    indexEnd = curePro.CureSteps(i).stepEnd
                End If

                curePro.CureSteps(i).tempResult("Max") = Math.Round(leadTC.Max(indexStart, indexEnd), 0)


                ''Min temp
                indexStart = curePro.CureSteps(i).stepStart
                indexEnd = curePro.CureSteps(i).stepEnd

                If Not firstStep Then
                    If curePro.CureSteps(i - 1).tempSet("RampRate") > 0 Then
                        indexStart = indexStart + Math.Abs(Math.Round(curePro.CureSteps(i - 1).tempSet("RampRate") / stepDuration, 0))
                    End If
                End If

                If Not lastStep Then
                    If curePro.CureSteps(i + 1).tempSet("RampRate") < 0 Then
                        indexEnd = indexEnd - Math.Abs(Math.Round(curePro.CureSteps(i + 1).tempSet("RampRate") / stepDuration, 0))
                    End If
                End If

                ''Handles a very short step: just look at min of whole step, no buffer
                If indexEnd <= indexStart Then
                    indexStart = curePro.CureSteps(i).stepStart
                    indexEnd = curePro.CureSteps(i).stepEnd
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




                ''##Temp Ramps##
                goal = -1
                greaterThanGoal = True

                If curePro.CureSteps(i).tempSet("RampRate") > 0 Then
                    goal = curePro.CureSteps(i).tempSet("SetPoint")
                    greaterThanGoal = True
                ElseIf curePro.CureSteps(i).tempSet("RampRate") < 0 Then
                    goal = curePro.CureSteps(i).tempSet("SetPoint")
                    greaterThanGoal = False
                End If

                ''Buffer start and end by half the linear regression step length so ramp is only from the step region
                indexStart = curePro.CureSteps(i).stepStart + (stepVal / 2)
                indexEnd = curePro.CureSteps(i).stepEnd - (stepVal / 2)

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

                curePro.CureSteps(i).tempResult("MaxRamp") = Math.Round(holder, 1)


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

                curePro.CureSteps(i).tempResult("MinRamp") = Math.Round(holder, 1)


                ''Average temp ramp
                total = 0
                addCnt = 0

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
                If Math.Abs(curePro.CureSteps(i).tempSet("SetPoint")) = Math.Abs(curePro.CureSteps(i).tempSet("NegTol")) Then
                    If curePro.CureSteps(i).tempResult("Max") > curePro.CureSteps(i).tempSet("SetPoint") + curePro.CureSteps(i).tempSet("PosTol") Then curePro.CureSteps(i).tempPass = False
                ElseIf Math.Abs(curePro.CureSteps(i).tempSet("SetPoint")) = Math.Abs(curePro.CureSteps(i).tempSet("PosTol")) Then
                    If curePro.CureSteps(i).tempResult("Min") < curePro.CureSteps(i).tempSet("SetPoint") + curePro.CureSteps(i).tempSet("NegTol") Then curePro.CureSteps(i).tempPass = False
                Else
                    If curePro.CureSteps(i).tempResult("Min") < curePro.CureSteps(i).tempSet("SetPoint") + curePro.CureSteps(i).tempSet("NegTol") Then curePro.CureSteps(i).tempPass = False
                    If curePro.CureSteps(i).tempResult("Max") > curePro.CureSteps(i).tempSet("SetPoint") + curePro.CureSteps(i).tempSet("PosTol") Then curePro.CureSteps(i).tempPass = False
                End If

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
            If curePro.checkPressure And curePro.CureSteps(i).pressureSet("SetPoint") <> -1 Then

                ''Max pressure
                indexStart = curePro.CureSteps(i).stepStart
                indexEnd = curePro.CureSteps(i).stepEnd

                If Not firstStep Then
                    If curePro.CureSteps(i - 1).pressureSet("RampRate") < 0 Then
                        indexStart = indexStart + Math.Abs(Math.Round(curePro.CureSteps(i - 1).pressureSet("RampRate") / stepDuration, 0))
                    End If
                End If

                If Not lastStep Then
                    If curePro.CureSteps(i + 1).pressureSet("RampRate") > 0 Then
                        indexEnd = indexEnd - Math.Abs(Math.Round(curePro.CureSteps(i + 1).pressureSet("RampRate") / stepDuration, 0))
                    End If
                End If

                ''Handles a very short step: just look at max of whole step, no buffer
                If indexEnd <= indexStart Then
                    indexStart = curePro.CureSteps(i).stepStart
                    indexEnd = curePro.CureSteps(i).stepEnd
                End If

                curePro.CureSteps(i).pressureResult("Max") = Math.Round(vesselPress.Max(indexStart, indexEnd), 1)


                ''Min pressure
                indexStart = curePro.CureSteps(i).stepStart
                indexEnd = curePro.CureSteps(i).stepEnd
                If Not firstStep Then
                    If curePro.CureSteps(i - 1).pressureSet("RampRate") > 0 Then
                        indexStart = indexStart + Math.Abs(Math.Round(curePro.CureSteps(i - 1).pressureSet("RampRate") / stepDuration, 0))
                    End If
                End If

                If Not lastStep Then
                    If curePro.CureSteps(i + 1).pressureSet("RampRate") < 0 Then
                        indexEnd = indexEnd - Math.Abs(Math.Round(curePro.CureSteps(i + 1).pressureSet("RampRate") / stepDuration, 0))
                    End If

                End If

                ''Handles a very short step: just look at max of whole step, no buffer
                If indexEnd <= indexStart Then
                    indexStart = curePro.CureSteps(i).stepStart
                    indexEnd = curePro.CureSteps(i).stepEnd
                End If

                curePro.CureSteps(i).pressureResult("Min") = Math.Round(vesselPress.Min(indexStart, indexEnd), 1)


                ''Average pressure
                indexStart = curePro.CureSteps(i).stepStart
                indexEnd = curePro.CureSteps(i).stepEnd

                curePro.CureSteps(i).pressureResult("Avg") = Math.Round(vesselPress.Average(indexStart, indexEnd), 1)




                ''##Pressure ramps##
                goal = -1
                greaterThanGoal = True

                If curePro.CureSteps(i).pressureSet("RampRate") > 0 Then
                    goal = curePro.CureSteps(i).pressureSet("SetPoint")
                    greaterThanGoal = True
                ElseIf curePro.CureSteps(i).pressureSet("RampRate") < 0 Then
                    goal = curePro.CureSteps(i).pressureSet("SetPoint")
                    greaterThanGoal = False
                End If

                ''Buffer start and end by half the linear regression step length so ramp is only from the step region
                indexStart = curePro.CureSteps(i).stepStart + (stepVal / 2)
                indexEnd = curePro.CureSteps(i).stepEnd - (stepVal / 2)

                ''For very short step (<linear reg step) just look at the middle point 
                If indexEnd < indexStart Then
                    Dim midPoint As Integer = indexStart + ((indexEnd - indexStart) / 2)
                    indexStart = midPoint
                    indexEnd = midPoint
                End If

                ''Max pressure ramp
                curePro.CureSteps(i).pressureResult("MaxRamp") = Math.Round(vesselPress.MaxRamp(indexStart, indexEnd, goal, greaterThanGoal), 1)

                ''Min pressure ramp
                curePro.CureSteps(i).pressureResult("MinRamp") = Math.Round(vesselPress.MinRamp(indexStart, indexEnd, goal, greaterThanGoal), 1)

                ''Avg pressure ramp
                curePro.CureSteps(i).pressureResult("AvgRamp") = Math.Round(vesselPress.AverageRamp(indexStart, indexEnd), 1)



                'Check pressure for passing
                If Math.Abs(curePro.CureSteps(i).pressureSet("SetPoint")) = Math.Abs(curePro.CureSteps(i).pressureSet("NegTol")) Then
                    If curePro.CureSteps(i).pressureResult("Max") > curePro.CureSteps(i).pressureSet("SetPoint") + curePro.CureSteps(i).pressureSet("PosTol") Then curePro.CureSteps(i).pressurePass = False
                ElseIf Math.Abs(curePro.CureSteps(i).pressureSet("SetPoint")) = Math.Abs(curePro.CureSteps(i).pressureSet("PosTol")) Then
                    If curePro.CureSteps(i).pressureResult("Min") < curePro.CureSteps(i).pressureSet("SetPoint") + curePro.CureSteps(i).pressureSet("NegTol") Then curePro.CureSteps(i).pressurePass = False
                Else
                    If curePro.CureSteps(i).pressureResult("Min") < curePro.CureSteps(i).pressureSet("SetPoint") + curePro.CureSteps(i).pressureSet("NegTol") Then curePro.CureSteps(i).pressurePass = False
                    If curePro.CureSteps(i).pressureResult("Max") > curePro.CureSteps(i).pressureSet("SetPoint") + curePro.CureSteps(i).pressureSet("PosTol") Then curePro.CureSteps(i).pressurePass = False
                End If

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
            If curePro.checkVac And curePro.CureSteps(i).vacSet("SetPoint") <> -1 Then
                indexStart = curePro.CureSteps(i).stepStart
                indexEnd = curePro.CureSteps(i).stepEnd

                curePro.CureSteps(i).vacResult("Max") = Math.Round(maxVac.Min(indexStart, indexEnd), 1)
                curePro.CureSteps(i).vacResult("Min") = Math.Round(minVac.Max(indexStart, indexEnd), 1)

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
                If Math.Abs(curePro.CureSteps(i).vacSet("SetPoint")) = Math.Abs(curePro.CureSteps(i).vacSet("NegTol")) Then
                    If curePro.CureSteps(i).vacResult("Min") > curePro.CureSteps(i).vacSet("SetPoint") + curePro.CureSteps(i).vacSet("PosTol") Then curePro.CureSteps(i).vacPass = False
                ElseIf Math.Abs(curePro.CureSteps(i).vacSet("SetPoint")) = Math.Abs(curePro.CureSteps(i).vacSet("PosTol")) Then
                    If curePro.CureSteps(i).vacResult("Max") < curePro.CureSteps(i).vacSet("SetPoint") + curePro.CureSteps(i).vacSet("NegTol") Then curePro.CureSteps(i).vacPass = False
                Else
                    If curePro.CureSteps(i).vacResult("Min") > curePro.CureSteps(i).vacSet("SetPoint") + curePro.CureSteps(i).vacSet("PosTol") Then curePro.CureSteps(i).vacPass = False
                    If curePro.CureSteps(i).vacResult("Max") < curePro.CureSteps(i).vacSet("SetPoint") + curePro.CureSteps(i).vacSet("NegTol") Then curePro.CureSteps(i).vacPass = False
                End If

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
        For i = cureStart To dataCnt
            If curePro.CureSteps(currentStep).stepStart = -1 And currentStep <> 0 Then
                curePro.CureSteps(currentStep).stepStart = i - 1
            ElseIf curePro.CureSteps(currentStep).stepStart = -1 Then
                curePro.CureSteps(currentStep).stepStart = i
            End If

            'Step start and end do not line up

            If meetTerms(curePro.CureSteps(currentStep), i) Then
                curePro.CureSteps(currentStep).stepEnd = i - 1
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

        If Not cureProfiles Is Nothing Then
            curePro = Nothing
            curePro = New CureProfile()
            curePro.deserializeCure(cureProfiles(Combo_CureProfile.SelectedIndex).serializeCure)
        End If


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
        Combo_CureProfileEdit.Items.Clear()
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
            Combo_CureProfileEdit.SelectedIndex = 0
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
                cureProfiles(UBound(cureProfiles)).fileEditDate = IO.File.GetLastWriteTime(inPath)
                Combo_CureProfile.Items.Add(cureProfiles(UBound(cureProfiles)).Name)
                Combo_CureProfileEdit.Items.Add(cureProfiles(UBound(cureProfiles)).Name)
            Next
        End If
    End Sub
#End Region

    Private Sub Btn_OpenFile_Click(sender As Object, e As EventArgs) Handles Btn_OpenFile.Click
        If OpenCSVFileDialog.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            Txt_FilePath.Text = OpenCSVFileDialog.FileName
            Call openCureFile()
        End If
    End Sub

    Sub openCureFile()
        Box_DataRecorder.Visible = False
        Box_TC.Visible = False
        Box_Vac.Visible = False

        If Txt_FilePath.Text = "File path..." Then
            Exit Sub
        End If

        Try
            Call loadCSVin(Txt_FilePath.Text)
            Call loadCureData()
            Box_RunParams.Enabled = True
            Box_RunLine.Enabled = True
        Catch ex As Exception
            If MsgBox(ex.Message, vbOKCancel, "Error") = vbCancel Then
                Box_RunParams.Enabled = False
                Box_RunLine.Enabled = False
                errorReset()
                Txt_FilePath.Text = "File path..."
            End If

            Box_RunLine.Enabled = False
        End Try

    End Sub

    Private Sub Txt_FilePath_TextChanged(sender As Object, e As EventArgs) Handles Txt_FilePath.TextChanged

        If STOP_Txt_FilePath_TextChanged Then Exit Sub

        If IO.File.Exists(Txt_FilePath.Text.Trim("""")) Then
            STOP_Txt_FilePath_TextChanged = True
            Txt_FilePath.Text = Txt_FilePath.Text.Trim("""")
            STOP_Txt_FilePath_TextChanged = False
        End If

        If IO.File.Exists(Txt_FilePath.Text) Then

            Txt_FilePath.BackColor = SystemColors.Window
            Call openCureFile()
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

        Combo_CureProfileEdit.SelectedIndex = Combo_CureProfile.SelectedIndex

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

    Private Sub Btn_ClearCells_Click(sender As Object, e As EventArgs) Handles Btn_ClearCells.Click
        Txt_JobNumber.Text = ""
        Txt_ProgramNumber.Text = ""
        Txt_DataRecorder.Text = "S"
        Txt_PartNumber.Text = ""
        Txt_Revision.Text = ""
        Txt_Qty.Text = ""
        Txt_PartDesc.Text = ""

        Box_RunLine.Enabled = False
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub


    Private Sub Combo_CureProfileEdit_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_CureProfileEdit.SelectedIndexChanged

        If cureProfiles(Combo_CureProfileEdit.SelectedIndex).Name <> Combo_CureProfileEdit.Text Then
            Throw New Exception("Cure profile does not line up with loaded array, please reload cure profiles.")
        End If

        Combo_CureProfile.SelectedIndex = Combo_CureProfileEdit.SelectedIndex
    End Sub



    Function numericCheck(sender As Object)
        If Not IsNumeric(sender.Text) Then
            If sender.Text = "-" Then
                Return False
            End If
            sender.Text = "0"
        End If

        If sender.Text = "-0" Then
            sender.Text = "-"
            sender.SelectionStart = 1
            Return False
        End If

        Return True
    End Function

#Region "Temp step intelligence"
    Private Sub Txt_tempSetEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_TempSetEdit.TextChanged
        If numericCheck(Txt_TempSetEdit) Then
            Call Txt_tempPosTolEdit_TextChanged(sender, e)
            Call Txt_tempNegTolEdit_TextChanged(sender, e)
        End If

        Check_tempMaxEdit_CheckedChanged(sender, e)
        Check_tempMinEdit_CheckedChanged(sender, e)
    End Sub

    Private Sub Txt_tempPosTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_TempPosTolEdit.TextChanged

        If numericCheck(Txt_TempPosTolEdit) Then
            If Math.Abs(CDbl(Txt_TempPosTolEdit.Text)) = Math.Abs(CDbl(Txt_TempSetEdit.Text)) Then
                Txt_TempPosTolEdit.BackColor = Color.PeachPuff
            Else
                Txt_TempPosTolEdit.BackColor = SystemColors.Window
            End If
        End If

    End Sub

    Private Sub Txt_tempNegTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_TempNegTolEdit.TextChanged
        If numericCheck(Txt_TempNegTolEdit) Then
            If Math.Abs(CDbl(Txt_TempNegTolEdit.Text)) = Math.Abs(CDbl(Txt_TempSetEdit.Text)) Then
                Txt_TempNegTolEdit.BackColor = Color.PeachPuff
            Else
                Txt_TempNegTolEdit.BackColor = SystemColors.Window
            End If
        End If

    End Sub

    Private Sub Txt_TempRampEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_TempRampEdit.TextChanged
        numericCheck(Txt_TempRampEdit)


        If Txt_TempRampEdit.Text = "0" Then
            Txt_TempRampPosTolEdit.Text = 0
            Txt_TempRampNegTolEdit.Text = 0
            Box_TempRampPosTolEdit.Enabled = False
            Box_TempRampNegTolEdit.Enabled = False
        Else
            Box_TempRampPosTolEdit.Enabled = True
            Box_TempRampNegTolEdit.Enabled = True
        End If
    End Sub

    Private Sub Txt_TempRampPosTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_TempRampPosTolEdit.TextChanged
        numericCheck(sender)
    End Sub

    Private Sub Txt_TempRampNegTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_TempRampNegTolEdit.TextChanged
        numericCheck(sender)
    End Sub

    Private Sub Check_tempStepEdit_CheckedChanged(sender As Object, e As EventArgs) Handles Check_tempStepEdit.CheckedChanged
        If Check_tempStepEdit.Checked Then
            Box_TempSetEdit.Enabled = True
            Box_TempPosTolEdit.Enabled = True
            Box_TempNegTolEdit.Enabled = True
            Box_TempRampEdit.Enabled = True
            Box_TempRampPosTolEdit.Enabled = True
            Box_TempRampNegTolEdit.Enabled = True
            Check_tempMaxEdit.Enabled = True
            Check_tempMinEdit.Enabled = True

            Txt_TempSetEdit.Text = 250
            Txt_TempPosTolEdit.Text = 20
            Txt_TempNegTolEdit.Text = -10
            Txt_TempRampEdit.Text = 3
            Txt_TempRampPosTolEdit.Text = 0
            Txt_TempRampNegTolEdit.Text = -3
            Check_tempMaxEdit.Checked = False
            Check_tempMinEdit.Checked = False
        Else
            Txt_TempSetEdit.Text = -1
            Txt_TempPosTolEdit.Text = 0
            Txt_TempNegTolEdit.Text = 0
            Txt_TempRampEdit.Text = 0
            Txt_TempRampPosTolEdit.Text = 0
            Txt_TempRampNegTolEdit.Text = 0
            Check_tempMaxEdit.Checked = False
            Check_tempMinEdit.Checked = False

            Check_tempMaxEdit.Enabled = False
            Check_tempMinEdit.Enabled = False
            Box_TempSetEdit.Enabled = False
            Box_TempPosTolEdit.Enabled = False
            Box_TempNegTolEdit.Enabled = False
            Box_TempRampEdit.Enabled = False
            Box_TempRampPosTolEdit.Enabled = False
            Box_TempRampNegTolEdit.Enabled = False
        End If
    End Sub

    Private Sub Check_tempMaxEdit_CheckedChanged(sender As Object, e As EventArgs) Handles Check_tempMaxEdit.CheckedChanged
        If Check_tempMaxEdit.Checked Then
            Check_tempMinEdit.Checked = False

            Box_TempNegTolEdit.Enabled = False
            Txt_TempNegTolEdit.Text = "-" & Txt_TempSetEdit.Text
        Else
            Box_TempNegTolEdit.Enabled = True
        End If
    End Sub

    Private Sub Check_tempMinEdit_CheckedChanged(sender As Object, e As EventArgs) Handles Check_tempMinEdit.CheckedChanged
        If Check_tempMinEdit.Checked Then
            Check_tempMaxEdit.Checked = False

            Box_TempPosTolEdit.Enabled = False
            Txt_TempPosTolEdit.Text = Txt_TempSetEdit.Text
        Else
            Box_TempPosTolEdit.Enabled = True
        End If
    End Sub
#End Region

#Region "Pressure step intelligence"
    Private Sub Txt_pressureSetEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_pressureSetEdit.TextChanged
        If numericCheck(Txt_pressureSetEdit) Then
            Call Txt_pressurePosTolEdit_TextChanged(sender, e)
            Call Txt_pressureNegTolEdit_TextChanged(sender, e)
        End If

        Check_pressureMaxEdit_CheckedChanged(sender, e)
        Check_pressureMinEdit_CheckedChanged(sender, e)
    End Sub

    Private Sub Txt_pressurePosTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_pressurePosTolEdit.TextChanged
        If numericCheck(Txt_pressurePosTolEdit) Then
            If Math.Abs(CDbl(Txt_pressurePosTolEdit.Text)) = Math.Abs(CDbl(Txt_pressureSetEdit.Text)) Then
                Txt_pressurePosTolEdit.BackColor = Color.PeachPuff
            Else
                Txt_pressurePosTolEdit.BackColor = SystemColors.Window
            End If
        End If

    End Sub

    Private Sub Txt_pressureNegTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_pressureNegTolEdit.TextChanged
        If numericCheck(Txt_pressureNegTolEdit) Then
            If Math.Abs(CDbl(Txt_pressureNegTolEdit.Text)) = Math.Abs(CDbl(Txt_pressureSetEdit.Text)) Then
                Txt_pressureNegTolEdit.BackColor = Color.PeachPuff
            Else
                Txt_pressureNegTolEdit.BackColor = SystemColors.Window
            End If
        End If
    End Sub

    Private Sub Txt_pressureRampEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_pressureRampEdit.TextChanged
        numericCheck(sender)

        If Txt_pressureRampEdit.Text = "0" Then
            Txt_pressureRampPosTolEdit.Text = 0
            Txt_pressureRampNegTolEdit.Text = 0
            Box_pressureRampPosTolEdit.Enabled = False
            Box_pressureRampNegTolEdit.Enabled = False
        Else
            Box_pressureRampPosTolEdit.Enabled = True
            Box_pressureRampNegTolEdit.Enabled = True
        End If
    End Sub

    Private Sub Txt_pressureRampPosTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_pressureRampPosTolEdit.TextChanged
        numericCheck(sender)
    End Sub

    Private Sub Txt_pressureRampNegTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_pressureRampNegTolEdit.TextChanged
        numericCheck(sender)
    End Sub

    Private Sub Check_pressureStepEdit_CheckedChanged(sender As Object, e As EventArgs) Handles Check_pressureStepEdit.CheckedChanged
        If Check_pressureStepEdit.Checked Then
            Box_pressureSetEdit.Enabled = True
            Box_pressurePosTolEdit.Enabled = True
            Box_pressureNegTolEdit.Enabled = True
            Box_pressureRampEdit.Enabled = True
            Box_pressureRampPosTolEdit.Enabled = True
            Box_pressureRampNegTolEdit.Enabled = True
            Check_pressureMaxEdit.Enabled = True
            Check_pressureMinEdit.Enabled = True

            Txt_pressureSetEdit.Text = 80
            Txt_pressurePosTolEdit.Text = 20
            Txt_pressureNegTolEdit.Text = -10
            Txt_pressureRampEdit.Text = 3
            Txt_pressureRampPosTolEdit.Text = 0
            Txt_pressureRampNegTolEdit.Text = -3
            Check_pressureMaxEdit.Checked = False
            Check_pressureMinEdit.Checked = False
        Else
            Txt_pressureSetEdit.Text = -1
            Txt_pressurePosTolEdit.Text = 0
            Txt_pressureNegTolEdit.Text = 0
            Txt_pressureRampEdit.Text = 0
            Txt_pressureRampPosTolEdit.Text = 0
            Txt_pressureRampNegTolEdit.Text = 0
            Check_pressureMaxEdit.Checked = False
            Check_pressureMinEdit.Checked = False

            Check_pressureMaxEdit.Enabled = False
            Check_pressureMinEdit.Enabled = False
            Box_pressureSetEdit.Enabled = False
            Box_pressurePosTolEdit.Enabled = False
            Box_pressureNegTolEdit.Enabled = False
            Box_pressureRampEdit.Enabled = False
            Box_pressureRampPosTolEdit.Enabled = False
            Box_pressureRampNegTolEdit.Enabled = False
        End If
    End Sub

    Private Sub Check_pressureMaxEdit_CheckedChanged(sender As Object, e As EventArgs) Handles Check_pressureMaxEdit.CheckedChanged
        If Check_pressureMaxEdit.Checked Then
            Check_pressureMinEdit.Checked = False

            Box_pressureNegTolEdit.Enabled = False
            Txt_pressureNegTolEdit.Text = "-" & Txt_pressureSetEdit.Text
        Else
            Box_pressureNegTolEdit.Enabled = True
        End If
    End Sub

    Private Sub Check_pressureMinEdit_CheckedChanged(sender As Object, e As EventArgs) Handles Check_pressureMinEdit.CheckedChanged
        If Check_pressureMinEdit.Checked Then
            Check_pressureMaxEdit.Checked = False

            Box_pressurePosTolEdit.Enabled = False
            Txt_pressurePosTolEdit.Text = Txt_pressureSetEdit.Text
        Else
            Box_pressurePosTolEdit.Enabled = True
        End If
    End Sub
#End Region

#Region "Vacuum step intelligence"
    Private Sub Txt_vacSetEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_vacSetEdit.TextChanged
        If numericCheck(Txt_vacSetEdit) Then
            Call Txt_vacPosTolEdit_TextChanged(sender, e)
            Call Txt_vacNegTolEdit_TextChanged(sender, e)
        End If

        Check_vacMaxEdit_CheckedChanged(sender, e)
        Check_vacMinEdit_CheckedChanged(sender, e)
    End Sub

    Private Sub Txt_vacPosTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_vacPosTolEdit.TextChanged

        If numericCheck(Txt_vacPosTolEdit) Then
            If Math.Abs(CDbl(Txt_vacPosTolEdit.Text)) = Math.Abs(CDbl(Txt_vacSetEdit.Text)) Then
                Txt_vacPosTolEdit.BackColor = Color.PeachPuff
            Else
                Txt_vacPosTolEdit.BackColor = SystemColors.Window
            End If
        End If

    End Sub

    Private Sub Txt_vacNegTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_vacNegTolEdit.TextChanged
        If numericCheck(Txt_vacNegTolEdit) Then
            If Math.Abs(CDbl(Txt_vacNegTolEdit.Text)) = Math.Abs(CDbl(Txt_vacSetEdit.Text)) Then
                Txt_vacNegTolEdit.BackColor = Color.PeachPuff
            Else
                Txt_vacNegTolEdit.BackColor = SystemColors.Window
            End If
        End If

    End Sub


    Private Sub Check_vacStepEdit_CheckedChanged(sender As Object, e As EventArgs) Handles Check_vacStepEdit.CheckedChanged
        If Check_vacStepEdit.Checked Then
            Box_vacSetEdit.Enabled = True
            Box_vacPosTolEdit.Enabled = True
            Box_vacNegTolEdit.Enabled = True
            Check_vacMaxEdit.Enabled = True
            Check_vacMinEdit.Enabled = True

            Txt_vacSetEdit.Text = -15
            Txt_vacPosTolEdit.Text = 2
            Txt_vacNegTolEdit.Text = -10
            Check_vacMaxEdit.Checked = False
            Check_vacMinEdit.Checked = False
        Else
            Txt_vacSetEdit.Text = -1
            Txt_vacPosTolEdit.Text = 0
            Txt_vacNegTolEdit.Text = 0
            Check_vacMaxEdit.Checked = False
            Check_vacMinEdit.Checked = False

            Check_vacMaxEdit.Enabled = False
            Check_vacMinEdit.Enabled = False
            Box_vacSetEdit.Enabled = False
            Box_vacPosTolEdit.Enabled = False
            Box_vacNegTolEdit.Enabled = False
        End If
    End Sub

    Private Sub Check_vacMaxEdit_CheckedChanged(sender As Object, e As EventArgs) Handles Check_vacMaxEdit.CheckedChanged
        If Check_vacMaxEdit.Checked Then
            Check_vacMinEdit.Checked = False

            Box_vacPosTolEdit.Enabled = False
            Box_vacNegTolEdit.Enabled = False
            Txt_vacPosTolEdit.Text = Txt_vacSetEdit.Text
            Txt_vacNegTolEdit.Text = 0
        ElseIf Not Check_vacMinEdit.Checked Then
            Box_vacPosTolEdit.Enabled = True
            Box_vacNegTolEdit.Enabled = True
        End If
    End Sub

    Private Sub Check_vacMinEdit_CheckedChanged(sender As Object, e As EventArgs) Handles Check_vacMinEdit.CheckedChanged
        If Check_vacMinEdit.Checked Then
            Check_vacMaxEdit.Checked = False

            Box_vacPosTolEdit.Enabled = False
            Box_vacNegTolEdit.Enabled = False

            Txt_vacPosTolEdit.Text = 0
            Txt_vacNegTolEdit.Text = Txt_vacSetEdit.Text
        ElseIf Not Check_vacMaxEdit.Checked Then
            Box_vacPosTolEdit.Enabled = True
            Box_vacNegTolEdit.Enabled = True
        End If
    End Sub

    Private Sub Check_TempEdit_CheckedChanged(sender As Object, e As EventArgs) Handles Check_TempEdit.CheckedChanged
        If Not Check_TempEdit.Checked Then
            Check_tempStepEdit.Checked = False
            Box_tempStepEdit.Enabled = False
        Else
            Check_tempStepEdit.Checked = True
            Box_tempStepEdit.Enabled = True
        End If
    End Sub
#End Region


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



