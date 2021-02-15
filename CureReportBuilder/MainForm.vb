Option Explicit On

Imports System.Runtime.CompilerServices
Imports System.Data.SqlClient
Imports System.Threading
Imports System.Text.RegularExpressions

Public Class MainForm

    Public cureProfiles() As CureProfile

    Dim mainCureCheck As CureCheck = New CureCheck()

    Public autoclaveSNum As New Dictionary(Of String, Integer) From {{"Controller", 2601}, {"Air TC1", 2600}, {"Air TC2", 2603}, {"PT", 2599}, {"VT1", 2593}, {"VT2", 2598}, {"VT3", 2595}, {"VT4", 2594}, {"VT5", 2597}, {"VT6", 2596}}


    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim vers As Version = My.Application.Info.Version
        Me.Text = "Cure Report Builder " & vers.Major & "." & vers.Minor & "." & vers.Build

        Call errorReset()

        'TabControl1.TabPages.Remove(TabPage3)

        Txt_Technician.Text = Replace(Environment.UserName, ".", " ")
        Txt_CureProfilesPath.Text = My.Settings.CureProfilePath
        Txt_TemplatePath.Text = My.Settings.ReportTemplatePath
        Txt_OutputPath.Text = My.Settings.OutputPath
        Txt_CureDataPath.Text = My.Settings.CureDataPath
        Txt_CureParamPath.Text = My.Settings.CureParamPath
        Txt_RunInterval.Text = My.Settings.RunInterval

        'If My.Settings.RunInterval > 0 Then
        '    Me.Timer1.Interval = My.Settings.RunInterval * 1000000
        '    Me.Timer1.Enabled = True
        'End If

        Try
            Call loadCureProfiles(Txt_CureProfilesPath.Text)
        Catch ex As Exception
            If MsgBox(ex.Message, vbOKCancel, "Cure Profiles: Path Error") = vbCancel Then Me.Close()
        End Try

        Me.Show()

        'Call batchRun()
    End Sub

#Region "Tool Strip Menu Commands"
    Private Sub LoadCureProfilesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadCureProfilesToolStripMenuItem.Click
        Btn_LoadProfileFiles.PerformClick()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
#End Region

#Region "Batch/Automated Run"
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim toRunList() As String
        Dim cureDataPath As String = ""
        Dim cureParamPath As String = ""

        If _
           System.IO.Directory.Exists(Txt_CureParamPath.Text) _
           AndAlso System.IO.File.Exists(Txt_CureParamPath.Text & "\ToRunList.log") _
           And System.IO.Directory.Exists(Txt_CureDataPath.Text) Then

            cureDataPath = Txt_CureDataPath.Text
            cureParamPath = Txt_CureParamPath.Text

            Dim logReader As IO.StreamReader = New IO.StreamReader(cureParamPath & "\" & "ToRunList.log")
            toRunList = Strings.Split(logReader.ReadToEnd, ",")
            logReader.Close()
        Else
            Throw New Exception("Failed to reach all file paths to run.")
        End If

        Dim threadArr() As Thread = Nothing

        For i = 0 To UBound(toRunList)
            Dim currentVal As String = toRunList(i)
            threadArr.AddValArr(New System.Threading.Thread(Sub() runCureCheckFromParamFile(currentVal)))
            threadArr(i).Start()
        Next

        For i = 0 To UBound(threadArr)
            threadArr(i).Join()
        Next

        'Get current newToRunList

        'Compare with runList and toRunList
        'Add any new items to the toRunList
        'Output toRunList as it currently stands



    End Sub

    Sub runCureCheckFromParamFile(runParamNam As String)
        Dim cureCheck As CureCheck = New CureCheck

        'Check if cure param file exists

        'Load in cure parameter file

        'Set cure



        'Find data file

        'Load in data file
        'If multiple are found try both and see which one has better results



        'If Epicor was loaded then Input cure parameters
        'Else try to get Epicor values

        'Input serial numbers

        'Input TC's used

        'Input vac used


        Call cureCheck.runCalc()

        Dim excelOutput As ExcelOutput = New ExcelOutput
        Call excelOutput.outputResults(Txt_TemplatePath.Text, cureCheck, Txt_OutputPath.Text)

        'Add cure param name to runList
        'Remove cure param name from toRunList
    End Sub

    'Sub batchRun()
    '    Dim inPath As String = "C:\Users\Will.Eagan\Desktop\CuresToRun.txt"

    '    ''Required input for batch run:
    '    '<curePro>CURE A1
    '    '<DataPath>S:\Manufacturing\Autoclave Cure Reports\Programs\D11675 - ARRW Shroud\BATCH RAW DATA\BATCH 155, JOB 102673, 102675, 5-26-20.CSV
    '    '<CompletedBy>Evan Sjostedt
    '    '<JobNum>102675
    '    '<PartNum>48777-100
    '    '<PartRev>E
    '    '<PartNom>ARRW COMPOSITE SHROUD (MOLDED)
    '    '<ProgramNum>D11675
    '    '<PartQty>1
    '    '<equipSerialNum>
    '    '<usrRunTC>2
    '    '<usrRunVac>5
    '    '~#####~


    '    Dim curesToRun() As String

    '    If IO.File.Exists(inPath) Then
    '        curesToRun = Split(IO.File.ReadAllText(inPath), "~#####~")


    '        For i = 0 To UBound(curesToRun)

    '            If curesToRun(i) = "" Then Exit Sub

    '            Txt_FilePath.Text = "File path..."

    '            Call errorReset()

    '            Dim cureParameters() As String
    '            cureParameters = Split(curesToRun(i), vbCrLf)

    '            Dim paramIndex As Integer = 0

    '            If cureParameters(0) = "" Then
    '                paramIndex = 1
    '            End If

    '            Dim cureProStr() As String = Split(cureParameters(paramIndex), ">")
    '            If cureProStr(0) <> "<curePro" Then Throw New Exception()
    '            Combo_CureProfile.SelectedIndex = Combo_CureProfile.Items.IndexOf(cureProStr(1))
    '            paramIndex += 1


    '            cureProStr = Split(cureParameters(paramIndex), ">")
    '            If cureProStr(0) <> "<DataPath" Then Throw New Exception()
    '            Txt_FilePath.Text = cureProStr(1)
    '            partValues("DataPath") = Txt_FilePath.Text
    '            paramIndex += 1

    '            cureProStr = Split(cureParameters(paramIndex), ">")
    '            If cureProStr(0) <> "<CompletedBy" Then Throw New Exception()
    '            Txt_Technician.Text = cureProStr(1)
    '            paramIndex += 1

    '            cureProStr = Split(cureParameters(paramIndex), ">")
    '            If cureProStr(0) <> "<JobNum" Then Throw New Exception()
    '            partValues("JobNum") = cureProStr(1)
    '            paramIndex += 1

    '            cureProStr = Split(cureParameters(paramIndex), ">")
    '            If cureProStr(0) <> "<PartNum" Then Throw New Exception()
    '            partValues("PartNum") = cureProStr(1)
    '            paramIndex += 1

    '            cureProStr = Split(cureParameters(paramIndex), ">")
    '            If cureProStr(0) <> "<PartRev" Then Throw New Exception()
    '            partValues("PartRev") = cureProStr(1)
    '            paramIndex += 1

    '            cureProStr = Split(cureParameters(paramIndex), ">")
    '            If cureProStr(0) <> "<PartNom" Then Throw New Exception()
    '            partValues("PartNom") = cureProStr(1)
    '            paramIndex += 1

    '            cureProStr = Split(cureParameters(paramIndex), ">")
    '            If cureProStr(0) <> "<ProgramNum" Then Throw New Exception()
    '            partValues("ProgramNum") = cureProStr(1)
    '            paramIndex += 1

    '            cureProStr = Split(cureParameters(paramIndex), ">")
    '            If cureProStr(0) <> "<PartQty" Then Throw New Exception()
    '            partValues("PartQty") = cureProStr(1)
    '            paramIndex += 1


    '            If machType = "Omega" Then
    '                cureProStr = Split(cureParameters(paramIndex), ">")
    '                If cureProStr(0) <> "<equipSerialNum" Then Throw New Exception()
    '                equipSerialNum = cureProStr(1)
    '            End If
    '            paramIndex += 1



    '            Dim rowCnt As Integer = 1

    '            If machType = "Autoclave" Then
    '                addToEquip(equipSerialNum, "Controller", rowCnt)
    '                addToEquip(equipSerialNum, "Air TC1", rowCnt)
    '                addToEquip(equipSerialNum, "Air TC2", rowCnt)
    '                addToEquip(equipSerialNum, "PT", rowCnt)
    '            End If

    '            'Get used TC lines
    '            If curePro.checkTemp Then
    '                usrRunTC.clearArr()

    '                cureProStr = Split(cureParameters(paramIndex), ">")
    '                If cureProStr(0) <> "<usrRunTC" Then Throw New Exception()
    '                Dim holder() As String = Split(cureProStr(1), ",")
    '                For j = 0 To UBound(holder)
    '                    If holder(j) <> "" Then
    '                        usrRunTC.AddValArr(CInt(holder(j)))
    '                    End If
    '                Next
    '                paramIndex += 1

    '                If usrRunTC Is Nothing Then
    '                    For j = 0 To Data_TC.Rows.Count - 1
    '                        usrRunTC.AddValArr(Data_TC.Item(1, j).Value)
    '                    Next
    '                End If
    '            End If


    '            'Get used vac ports
    '            If curePro.checkVac Then
    '                usrRunVac.clearArr()

    '                cureProStr = Split(cureParameters(paramIndex), ">")
    '                If cureProStr(0) <> "<usrRunVac" Then Throw New Exception()
    '                Dim holder() As String = Split(cureProStr(1), ",")
    '                For j = 0 To UBound(holder)
    '                    If holder(j) <> "" Then
    '                        usrRunVac.AddValArr(CInt(holder(j)))
    '                    End If
    '                Next
    '                paramIndex += 1

    '                If usrRunVac Is Nothing Then
    '                    For j = 0 To Data_Vac.Rows.Count - 1
    '                        usrRunVac.AddValArr(Data_Vac.Item(1, j).Value)
    '                    Next
    '                End If

    '                For j = 0 To UBound(usrRunVac)
    '                    addToEquip(equipSerialNum, "VT" & usrRunVac(j), rowCnt)
    '                Next
    '            End If

    '            If machType = "Autoclave" Then
    '                If Strings.Right(equipSerialNum, 3) = " | " Then
    '                    equipSerialNum = Strings.Left(equipSerialNum, Len(equipSerialNum) - 3)
    '                ElseIf Strings.Right(equipSerialNum, 1) = vbCrLf Then
    '                    equipSerialNum = Strings.Left(equipSerialNum, Len(equipSerialNum) - 1)
    '                End If
    '            End If

    '            Call runCalc()

    '            Call outputResults()


    '        Next
    '    End If
    'End Sub
#End Region

#Region "User Check - Input"
    Private Sub Combo_CureProfile_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_CureProfile.SelectedIndexChanged

        If cureProfiles(Combo_CureProfile.SelectedIndex).Name <> Combo_CureProfile.Text Then
            Throw New Exception("Cure profile does not line up with loaded array, please reload cure profiles.")
        End If



        Call errorReset()

        Txt_CureDoc.Text = mainCureCheck.curePro.cureDoc
        Txt_DocRev.Text = mainCureCheck.curePro.cureDocRev

        Txt_FilePath.Text = "File path..."
        Box_CureDataLocation.Enabled = True
    End Sub

    Private Sub Btn_OpenFile_Click(sender As Object, e As EventArgs) Handles Btn_OpenFile.Click
        If OpenCSVFileDialog.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            Txt_FilePath.Text = OpenCSVFileDialog.FileName
            Call openCureFile()
        End If
    End Sub

    Private Sub Txt_FilePath_TextChanged(sender As Object, e As EventArgs) Handles Txt_FilePath.TextChanged

        If IO.File.Exists(Txt_FilePath.Text.Trim("""")) Then
            RemoveHandler Txt_FilePath.TextChanged, AddressOf Txt_FilePath_TextChanged
            Txt_FilePath.Text = Txt_FilePath.Text.Trim("""")
            AddHandler Txt_FilePath.TextChanged, AddressOf Txt_FilePath_TextChanged
        End If

        If IO.File.Exists(Txt_FilePath.Text) Then
            Txt_FilePath.BackColor = SystemColors.Window
            Call openCureFile()
        Else
            Txt_FilePath.BackColor = Color.PeachPuff
            Box_RunLine.Enabled = False
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
            Call errorReset()
            Call mainCureCheck.loadCSVin(Txt_FilePath.Text)
            If System.Text.RegularExpressions.Regex.Match(Txt_FilePath.Text, "\d{6}").Value <> "" Then
                Txt_JobNumber.Text = System.Text.RegularExpressions.Regex.Match(Txt_FilePath.Text, "\d{6}").Value
            End If

            If System.Text.RegularExpressions.Regex.Match(Txt_FilePath.Text, "S(\d{4})").Value <> "" Then
                Txt_DataRecorder.Text = System.Text.RegularExpressions.Regex.Match(Txt_FilePath.Text, "S(\d{4})").Value
            End If

            If mainCureCheck.machType = "Omega" Then
                Box_DataRecorder.Visible = True
            End If

            Call mainCureCheck.loadCureData()

            If mainCureCheck.curePro.checkTemp = True Then
                Call addToList(mainCureCheck.partTC_Arr, Data_TC, Box_TC)
            End If

            If mainCureCheck.curePro.checkVac Then
                Call addToList(mainCureCheck.vac_Arr, Data_Vac, Box_Vac)
            End If


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

    Sub addToList(inArr() As DataSet, inGrid As DataGridView, inContainer As GroupBox)
        inContainer.Visible = True
        inGrid.Rows.Clear()

        For i = 0 To UBound(inArr)
            inGrid.Rows.Add()
            inGrid.Item(1, i).Value = inArr(i).Number
        Next

        inGrid.Sort(inGrid.Columns(1), System.ComponentModel.ListSortDirection.Ascending)

    End Sub

    Private Sub Txt_JobNumber_TextChanged(sender As Object, e As EventArgs) Handles Txt_JobNumber.TextChanged
        If Len(Txt_JobNumber.Text) = 6 And System.Text.RegularExpressions.Regex.IsMatch(Txt_JobNumber.Text, "^[0-9]+$") Then
            Try
                Dim epicorData As DataRow = getEpicorData(Txt_JobNumber.Text)
                Txt_ProgramNumber.Text = epicorData.Item("PhaseID")
                Txt_PartNumber.Text = epicorData.Item("PartNum")
                Txt_Revision.Text = epicorData.Item("RevisionNum")
                Txt_Qty.Text = Math.Round(epicorData.Item("QtyPer"), 0)
                Txt_PartDesc.Text = epicorData.Item("Description")

            Catch ex As Exception
                Txt_ProgramNumber.Text = ""
                Txt_PartNumber.Text = ""
                Txt_Revision.Text = ""
                Txt_Qty.Text = ""
                Txt_PartDesc.Text = ""
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Function getEpicorData(jobNum As Integer) As DataRow

        If Len(jobNum.ToString) <> 6 And System.Text.RegularExpressions.Regex.IsMatch(jobNum.ToString, "^[0-9]+$") Then
            Throw New Exception("Job number for Epicor must be a 6 digit number.")
        End If

        Dim queryResult As DataTable = New DataTable()

        Dim db As SqlConnection = New SqlConnection("Data Source = MAUI;
                                                     Initial Catalog=EpicorERP;
                                                     User ID = Reporting;
                                                     Password=$ystima1;
                                                     Integrated Security=false;
                                                     trusted_connection=false;")

        db.Open()
        Dim adapter As SqlDataAdapter = New SqlDataAdapter(
            "SELECT 
                Job.ProjectID, 
                JOB.PhaseID, 
                JOB.JobNum, 
                ASM.PartNum, 
                ASM.AssemblySeq, 
                OP.OprSeq, 
                ASM.RevisionNum, 
                ASM.Description, 
                ASM.QtyPer 
            FROM Erp.JobAsmbl ASM
                LEFT JOIN Erp.JobHead JOB ON JOB.JobNum = ASM.JobNum
                INNER JOIN Erp.JobOper OP ON OP.JobNum = ASM.JobNum AND ASM.AssemblySeq = OP.AssemblySeq AND OP.OpCode = 'LAYUP'
            WHERE JOB.JobNum = '" + jobNum.ToString + "'", db)

        adapter.Fill(queryResult)
        db.Close()

        If queryResult.Rows.Count = 0 Then
            Throw New Exception("Job number not found in Epicor")
        End If

        Return queryResult.Rows.Item(0)

    End Function

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
#End Region

#Region "User Check - Run"
    Private Sub Btn_Run_Click(sender As Object, e As EventArgs) Handles Btn_Run.Click
        RunUserCheck()
        Box_RunLine.Enabled = False
    End Sub

    Sub RunUserCheck()

        mainCureCheck.JobNum = Txt_JobNumber.Text
        mainCureCheck.PONum = ""
        mainCureCheck.PartNum = Txt_PartNumber.Text
        mainCureCheck.PartRev = Txt_Revision.Text
        mainCureCheck.PartNom = Txt_PartDesc.Text
        mainCureCheck.ProgramNum = Txt_ProgramNumber.Text
        mainCureCheck.PartQty = Txt_Qty.Text
        mainCureCheck.DataPath = Txt_FilePath.Text
        mainCureCheck.completedBy = Txt_Technician.Text

        If mainCureCheck.machType = "Omega" Then
            mainCureCheck.equipSerialNum = Txt_DataRecorder.Text
        End If


        Dim rowCnt As Integer = 1

        If mainCureCheck.machType = "Autoclave" Then
            addToEquip(mainCureCheck.equipSerialNum, "Controller", rowCnt)
            addToEquip(mainCureCheck.equipSerialNum, "Air TC1", rowCnt)
            addToEquip(mainCureCheck.equipSerialNum, "Air TC2", rowCnt)
            addToEquip(mainCureCheck.equipSerialNum, "PT", rowCnt)
        End If

        'Get used TC lines
        If mainCureCheck.curePro.checkTemp Then
            mainCureCheck.usrRunTC.clearArr()

            For i = 0 To Data_TC.Rows.Count - 1
                If Data_TC.Item(0, i).Value Then
                    mainCureCheck.usrRunTC.AddValArr(Data_TC.Item(1, i).Value)
                End If
            Next

            If mainCureCheck.usrRunTC Is Nothing Then
                If MsgBox("No TC's were selected, are you sure you would like to run with all of them?", vbOKCancel, "TC Check") = vbCancel Then Exit Sub
                For i = 0 To Data_TC.Rows.Count - 1
                    mainCureCheck.usrRunTC.AddValArr(Data_TC.Item(1, i).Value)
                Next
            End If
        End If


        'Get used vac ports
        If mainCureCheck.curePro.checkVac Then
            mainCureCheck.usrRunVac.clearArr()

            For i = 0 To Data_Vac.Rows.Count - 1
                If Data_Vac.Item(0, i).Value Then
                    mainCureCheck.usrRunVac.AddValArr(Data_Vac.Item(1, i).Value)
                End If
            Next

            If mainCureCheck.usrRunVac Is Nothing Then
                If MsgBox("No Vac ports were selected, are you sure you would like to run with all of them?", vbOKCancel, "Vac Check") = vbCancel Then Exit Sub
                For i = 0 To Data_Vac.Rows.Count - 1
                    mainCureCheck.usrRunVac.AddValArr(Data_Vac.Item(1, i).Value)
                Next
            End If

            For i = 0 To UBound(mainCureCheck.usrRunVac)
                addToEquip(mainCureCheck.equipSerialNum, "VT" & mainCureCheck.usrRunVac(i), rowCnt)
            Next
        End If

        If mainCureCheck.machType = "Autoclave" Then
            If Strings.Right(mainCureCheck.equipSerialNum, 3) = " | " Then
                mainCureCheck.equipSerialNum = Strings.Left(mainCureCheck.equipSerialNum, Len(mainCureCheck.equipSerialNum) - 3)
            ElseIf Strings.Right(mainCureCheck.equipSerialNum, 1) = vbCrLf Then
                mainCureCheck.equipSerialNum = Strings.Left(mainCureCheck.equipSerialNum, Len(mainCureCheck.equipSerialNum) - 1)
            End If
        End If

        Call mainCureCheck.runCalc()

        Dim excelOutput As ExcelOutput = New ExcelOutput
        Call excelOutput.outputResults(Txt_TemplatePath.Text, mainCureCheck, Txt_OutputPath.Text)

        Call errorReset()
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



    Sub errorReset()

        mainCureCheck = Nothing
        mainCureCheck = New CureCheck

        If Not cureProfiles Is Nothing Then
            mainCureCheck.curePro.DeserializeCure(cureProfiles(Combo_CureProfile.SelectedIndex).SerializeCure)
        End If


    End Sub
#End Region

#Region "Cure Profile Import/Export"

    Private Sub Txt_CureProfiles_TextChanged(sender As Object, e As EventArgs) Handles Txt_CureProfilesPath.TextChanged
        If IO.File.Exists(Txt_CureProfilesPath.Text) Then
            Txt_CureProfilesPath.BackColor = SystemColors.Window
        ElseIf IO.Directory.Exists(Txt_CureProfilesPath.Text) Then
            Txt_CureProfilesPath.BackColor = SystemColors.Window
        Else
            Txt_CureProfilesPath.BackColor = Color.PeachPuff
        End If
    End Sub

    Private Sub Btn_LoadProfileFiles_Click(sender As Object, e As EventArgs) Handles Btn_LoadProfileFiles.Click
        Dim curPath As String = Txt_CureProfilesPath.Text

        If IO.File.Exists(curPath) Or IO.Directory.Exists(curPath) Then
            Call loadCureProfiles(curPath)
        ElseIf OpenCureProfileFileDialog.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            Txt_CureProfilesPath.Text = OpenCureProfileFileDialog.FileName
            Call loadCureProfiles(Txt_CureProfilesPath.Text)
        End If

        My.Settings.CureProfilePath = Txt_CureProfilesPath.Text
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
            Try
                Combo_CureProfile.SelectedIndex = Combo_CureProfile.Items.IndexOf("CURE A1")
            Catch ex As Exception
                Combo_CureProfile.SelectedIndex = 0
            End Try

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

                cureProfiles(UBound(cureProfiles)).DeserializeCure(cureDef(i))
                cureProfiles(UBound(cureProfiles)).fileEditDate = IO.File.GetLastWriteTime(inPath)
                Combo_CureProfile.Items.Add(cureProfiles(UBound(cureProfiles)).Name)
            Next
        End If
    End Sub

    Sub outputCureProfiles(inPath As String)
        Dim i As Integer
        For i = 0 To UBound(cureProfiles)
            Dim outputWriter As IO.StreamWriter = New System.IO.StreamWriter(inPath & "/" & cureProfiles(i).Name & ".cprof")
            outputWriter.Write(cureProfiles(i).SerializeCure)
            outputWriter.Close()
        Next
    End Sub
#End Region

#Region "Settings Change"
    Private Sub Txt_TemplatePath_TextChanged(sender As Object, e As EventArgs) Handles Txt_TemplatePath.TextChanged
        If IO.File.Exists(Txt_TemplatePath.Text) Then
            My.Settings.ReportTemplatePath = Txt_TemplatePath.Text
            Txt_TemplatePath.BackColor = SystemColors.Window
        Else
            Txt_TemplatePath.BackColor = Color.PeachPuff
        End If

    End Sub

    Private Sub Txt_OutputPath_TextChanged(sender As Object, e As EventArgs) Handles Txt_OutputPath.TextChanged
        If IO.Directory.Exists(Txt_OutputPath.Text) Or Txt_OutputPath.Text = "" Then
            My.Settings.OutputPath = Txt_OutputPath.Text
            Txt_OutputPath.BackColor = SystemColors.Window
        Else
            Txt_OutputPath.BackColor = Color.PeachPuff
        End If
    End Sub

    Private Sub Txt_CureDataPath_TextChanged(sender As Object, e As EventArgs) Handles Txt_CureDataPath.TextChanged
        If IO.Directory.Exists(Txt_CureDataPath.Text) Then
            My.Settings.CureDataPath = Txt_CureDataPath.Text
            Txt_CureDataPath.BackColor = SystemColors.Window
        Else
            Txt_CureDataPath.BackColor = Color.PeachPuff
        End If
    End Sub

    Private Sub Txt_CureParamPath_TextChanged(sender As Object, e As EventArgs) Handles Txt_CureParamPath.TextChanged
        If IO.Directory.Exists(Txt_CureParamPath.Text) Then
            My.Settings.CureParamPath = Txt_CureParamPath.Text
            Txt_CureParamPath.BackColor = SystemColors.Window
        Else
            Txt_CureParamPath.BackColor = Color.PeachPuff
        End If
    End Sub

    Private Sub Txt_RunInterval_TextChanged(sender As Object, e As EventArgs) Handles Txt_RunInterval.TextChanged
        If IsNumeric(Txt_RunInterval.Text) AndAlso Txt_RunInterval.Text > 5 Then
            My.Settings.RunInterval = Txt_RunInterval.Text
            Me.Timer1.Interval = Txt_RunInterval.Text * 1000000
            Me.Timer1.Enabled = True
            Txt_RunInterval.BackColor = SystemColors.Window
        ElseIf IsNumeric(Txt_RunInterval.Text) Then
            My.Settings.RunInterval = Txt_RunInterval.Text
            Me.Timer1.Enabled = False
            Txt_RunInterval.BackColor = SystemColors.Window
        Else
            Me.Timer1.Enabled = False
            Txt_RunInterval.BackColor = Color.PeachPuff
        End If

    End Sub
#End Region

#Region "Cure Editor"
    Public profileEditFS As IO.FileStream = Nothing
    Public loadedEditProfile As CureProfile = Nothing
    Public Property Txt_CureDocrevEdit As Object
    Public loadedStepIndex As Integer = -1

#Region "Load/Save/Close Profile"
    Private Sub Txt_CurePathEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_CurePathEdit.TextChanged
        If IO.File.Exists(Txt_CurePathEdit.Text) Then
            Txt_CurePathEdit.BackColor = SystemColors.Window
        Else
            Txt_CurePathEdit.BackColor = Color.PeachPuff
        End If
    End Sub

    Private Sub Btn_CureLoadEdit_Click(sender As Object, e As EventArgs) Handles Btn_CureLoadEdit.Click
        If IO.File.Exists(Txt_CurePathEdit.Text) Then
            Call loadCureFileEdit(Txt_CurePathEdit.Text)
        ElseIf OpenCureProfileFileDialog.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            Txt_CurePathEdit.Text = OpenCureProfileFileDialog.FileName
            Call loadCureFileEdit(Txt_CurePathEdit.Text)
        End If
    End Sub

    Sub loadCureFileEdit(inPath As String)
        Call EditorReset()

        Dim cureDef() As String

        If IO.File.Exists(inPath) And IO.Path.GetExtension(inPath) = ".cprof" Then

            profileEditFS = New IO.FileStream(inPath, IO.FileMode.Open, IO.FileAccess.ReadWrite, IO.FileShare.None)

            Dim reader As IO.StreamReader = New IO.StreamReader(profileEditFS)
            cureDef = Split(reader.ReadToEnd, "~&&&~")

            loadedEditProfile = New CureProfile
            loadedEditProfile.DeserializeCure(cureDef(0))

            Txt_CureNameEdit.Text = loadedEditProfile.Name
            Txt_CureDocEdit.Text = loadedEditProfile.cureDoc
            Txt_DocRevEdit.Text = loadedEditProfile.cureDocRev
            Check_TempEdit.Checked = loadedEditProfile.checkTemp
            Check_PressureEdit.Checked = loadedEditProfile.checkPressure
            Check_VacEdit.Checked = loadedEditProfile.checkVac

            List_StepsEdit.Items.Clear()
            For i = 0 To UBound(loadedEditProfile.CureSteps)
                List_StepsEdit.Items.Add(loadedEditProfile.CureSteps(i).stepName)
            Next

            GrpBox_CureDefEdit.Enabled = True
            Btn_CloseEdit.Enabled = True
            Btn_SaveEdit.Enabled = True
            List_StepsEdit.SetSelected(0, True)
        End If
    End Sub

    Private Sub Btn_SaveEdit_Click(sender As Object, e As EventArgs) Handles Btn_SaveEdit.Click
        Call SetStepValues()

        Dim reader As IO.StreamReader = New IO.StreamReader(profileEditFS)

        Dim loadedCureStr As String = loadedEditProfile.SerializeCure

        If reader.ReadToEnd <> loadedCureStr Then
            profileEditFS.SetLength(0)

            Dim writer As IO.StreamWriter = New IO.StreamWriter(profileEditFS)
            writer.Write(loadedCureStr)
            writer.Flush()

        End If

    End Sub

    Private Sub Btn_CloseEdit_Click(sender As Object, e As EventArgs) Handles Btn_CloseEdit.Click
        Call EditorReset()
    End Sub

    Sub EditorReset()
        If Not profileEditFS Is Nothing And Not loadedEditProfile Is Nothing Then
            profileEditFS.Position = 0

            Dim reader As IO.StreamReader = New IO.StreamReader(profileEditFS)

            Dim loadedCureStr As String = loadedEditProfile.SerializeCure
            Dim fileCurStr As String = reader.ReadToEnd

            If fileCurStr <> loadedCureStr Then
                Dim usrResponse As Long = MsgBox("Would you like to save changes to the loaded cure profile before closing?", vbYesNoCancel, "Save Check")
                If usrResponse = vbYes Then
                    Btn_SaveEdit.PerformClick()
                ElseIf usrResponse = vbCancel Then
                    Exit Sub
                End If
            End If

            profileEditFS.Close()
        End If

        profileEditFS = Nothing
        loadedEditProfile = Nothing
        loadedStepIndex = -1
        Btn_CloseEdit.Enabled = False
        Btn_SaveEdit.Enabled = False
        GrpBox_CureDefEdit.Enabled = False
    End Sub
#End Region

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

    Function NumericTxtBoxCheck(inTxtBox As TextBox) As Boolean
        If IsNumeric(inTxtBox.Text) Then
            inTxtBox.BackColor = SystemColors.Window
            Return True
        Else
            inTxtBox.BackColor = Color.PeachPuff
        End If

        Return False
    End Function

    Private Sub Txt_CureNameEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_CureNameEdit.TextChanged
        loadedEditProfile.Name = Txt_CureNameEdit.Text
    End Sub

    Private Sub Txt_CureDocEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_CureDocEdit.TextChanged
        loadedEditProfile.cureDoc = Txt_CureDocEdit.Text
    End Sub

    Private Sub Txt_DocRevEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_DocRevEdit.TextChanged
        loadedEditProfile.cureDocRev = Txt_DocRevEdit.Text
    End Sub

    Private Sub Check_TempEdit_CheckedChanged(sender As Object, e As EventArgs) Handles Check_TempEdit.CheckedChanged
        'Needs more work
        If Not Check_TempEdit.Checked Then
            Check_tempStepEdit.Checked = False
            Box_tempStepEdit.Enabled = False
        Else
            Check_tempStepEdit.Checked = True
            Box_tempStepEdit.Enabled = True
        End If
    End Sub

    Private Sub Check_PressureEdit_CheckedChanged(sender As Object, e As EventArgs) Handles Check_PressureEdit.CheckedChanged
        'Needs more work
        If Not Check_PressureEdit.Checked Then
            Check_pressureStepEdit.Checked = False
            Box_pressureStepEdit.Enabled = False
        Else
            Check_pressureStepEdit.Checked = True
            Box_pressureStepEdit.Enabled = True
        End If
    End Sub

    Private Sub Check_VacEdit_CheckedChanged(sender As Object, e As EventArgs) Handles Check_VacEdit.CheckedChanged
        'Needs more work
        If Not Check_VacEdit.Checked Then
            Check_vacStepEdit.Checked = False
            Box_vacStepEdit.Enabled = False
        Else
            Check_vacStepEdit.Checked = True
            Box_vacStepEdit.Enabled = True
        End If
    End Sub

#Region "Steps List Commands"
    Private Sub Btn_MoveUpEdit_Click(sender As Object, e As EventArgs) Handles Btn_MoveStepUpEdit.Click
        Dim holder As String
        Dim selIndex As Integer = List_StepsEdit.SelectedIndex
        loadedStepIndex = -1

        loadedEditProfile.MoveStepUp(selIndex)

        RemoveHandler List_StepsEdit.SelectedIndexChanged, AddressOf List_StepsEdit_SelectedIndexChanged

        If selIndex <> 0 Then
            holder = List_StepsEdit.Items.Item(selIndex - 1)
            List_StepsEdit.Items.Item(selIndex - 1) = List_StepsEdit.Items.Item(selIndex)
            List_StepsEdit.Items.Item(selIndex) = holder

            AddHandler List_StepsEdit.SelectedIndexChanged, AddressOf List_StepsEdit_SelectedIndexChanged
            List_StepsEdit.SetSelected(selIndex - 1, True)
        Else
            holder = List_StepsEdit.Items.Item(0)

            For i = 0 To List_StepsEdit.Items.Count - 2
                List_StepsEdit.Items.Item(i) = List_StepsEdit.Items.Item(i + 1)
            Next

            List_StepsEdit.Items.Item(List_StepsEdit.Items.Count - 1) = holder

            AddHandler List_StepsEdit.SelectedIndexChanged, AddressOf List_StepsEdit_SelectedIndexChanged
            List_StepsEdit.SetSelected(List_StepsEdit.Items.Count - 1, True)
        End If



    End Sub

    Private Sub Btn_MoveDownEdit_Click(sender As Object, e As EventArgs) Handles Btn_MoveStepDownEdit.Click
        Dim holder As String
        Dim selIndex As Integer = List_StepsEdit.SelectedIndex
        loadedStepIndex = -1

        loadedEditProfile.MoveStepDown(selIndex)

        RemoveHandler List_StepsEdit.SelectedIndexChanged, AddressOf List_StepsEdit_SelectedIndexChanged

        If selIndex <> List_StepsEdit.Items.Count - 1 Then
            holder = List_StepsEdit.Items.Item(selIndex + 1)
            List_StepsEdit.Items.Item(selIndex + 1) = List_StepsEdit.Items.Item(selIndex)
            List_StepsEdit.Items.Item(selIndex) = holder

            AddHandler List_StepsEdit.SelectedIndexChanged, AddressOf List_StepsEdit_SelectedIndexChanged
            List_StepsEdit.SetSelected(selIndex + 1, True)
        Else
            holder = List_StepsEdit.Items.Item(List_StepsEdit.Items.Count - 1)

            For i = List_StepsEdit.Items.Count - 1 To 1 Step -1
                List_StepsEdit.Items.Item(i) = List_StepsEdit.Items.Item(i - 1)
            Next

            List_StepsEdit.Items.Item(0) = holder

            AddHandler List_StepsEdit.SelectedIndexChanged, AddressOf List_StepsEdit_SelectedIndexChanged
            List_StepsEdit.SetSelected(0, True)
        End If


    End Sub

    Private Sub Btn_AddStepEdit_Click(sender As Object, e As EventArgs) Handles Btn_AddStepEdit.Click
        RemoveHandler List_StepsEdit.SelectedIndexChanged, AddressOf List_StepsEdit_SelectedIndexChanged
        List_StepsEdit.Items.Add("New Step")
        AddHandler List_StepsEdit.SelectedIndexChanged, AddressOf List_StepsEdit_SelectedIndexChanged

        loadedEditProfile.addCureStep("New Step")

        List_StepsEdit.SetSelected(List_StepsEdit.Items.Count - 1, True)
    End Sub

    Private Sub Btn_RemoveStepEdit_Click(sender As Object, e As EventArgs) Handles Btn_RemoveStepEdit.Click
        Dim selIndex As Integer = List_StepsEdit.SelectedIndex
        If selIndex = -1 Then Exit Sub

        RemoveHandler List_StepsEdit.SelectedIndexChanged, AddressOf List_StepsEdit_SelectedIndexChanged
        List_StepsEdit.Items.RemoveAt(selIndex)
        loadedStepIndex = -1
        AddHandler List_StepsEdit.SelectedIndexChanged, AddressOf List_StepsEdit_SelectedIndexChanged

        loadedEditProfile.RemoveCureStep(selIndex)

        If List_StepsEdit.Items.Count > 0 Then
            If selIndex > List_StepsEdit.Items.Count - 1 Then
                List_StepsEdit.SetSelected(List_StepsEdit.Items.Count - 1, True)
            Else
                List_StepsEdit.SetSelected(selIndex, True)
            End If
        End If



    End Sub

    Private Sub List_StepsEdit_SelectedIndexChanged(sender As Object, e As EventArgs) Handles List_StepsEdit.SelectedIndexChanged
        Call SetStepValues()
        Call GetStepValues()
    End Sub


    Sub GetStepValues()
        Dim selIndex As Integer = List_StepsEdit.SelectedIndex
        Dim selStep As CureStep = loadedEditProfile.CureSteps(selIndex)

        loadedStepIndex = selIndex

        Call RemoveTempHandlers()
        Call RemoveVacHandlers()
        Call RemovePressureHandlers()
        Call RemoveDurationHandlers()
        Call RemoveTermHandlers()

        RemoveHandler Txt_StepNameEdit.TextChanged, AddressOf Txt_StepNameEdit_TextChanged
        Txt_StepNameEdit.Text = selStep.stepName
        AddHandler Txt_StepNameEdit.TextChanged, AddressOf Txt_StepNameEdit_TextChanged

        Txt_DurationEdit.Text = selStep.stepDuration

        Txt_TempSetEdit.Text = selStep.tempSet.SetPoint
        Txt_TempPosTolEdit.Text = selStep.tempSet.PosTol
        Txt_TempNegTolEdit.Text = selStep.tempSet.NegTol
        Txt_TempRampEdit.Text = selStep.tempSet.RampRate
        Txt_TempRampPosTolEdit.Text = selStep.tempSet.RampPosTol
        Txt_TempRampNegTolEdit.Text = selStep.tempSet.RampNegTol

        Txt_pressureSetEdit.Text = selStep.pressureSet.SetPoint
        Txt_pressurePosTolEdit.Text = selStep.pressureSet.PosTol
        Txt_pressureNegTolEdit.Text = selStep.pressureSet.NegTol
        Txt_pressureRampEdit.Text = selStep.pressureSet.RampRate
        Txt_pressureRampPosTolEdit.Text = selStep.pressureSet.RampPosTol
        Txt_pressureRampNegTolEdit.Text = selStep.pressureSet.RampNegTol

        Txt_vacSetEdit.Text = selStep.vacSet.SetPoint
        Txt_vacPosTolEdit.Text = selStep.vacSet.PosTol
        Txt_vacNegTolEdit.Text = selStep.vacSet.NegTol

        Combo_Term1Type.SelectedItem = selStep.termCond1Type
        Combo_Term2Type.SelectedItem = selStep.termCond2Type
        Call TermComboBoxOptionsUpdate()

        Combo_Term1Cond.SelectedItem = selStep.termCond1Condition
        Txt_Term1Goal.Text = selStep.termCond1Goal
        Combo_Term1Mod.SelectedItem = selStep.termCond1Modifier

        Combo_Term2Cond.SelectedItem = selStep.termCond2Condition
        Txt_Term2Goal.Text = selStep.termCond2Goal
        Combo_Term2Mod.SelectedItem = selStep.termCond2Modifier

        Combo_TermOper.SelectedItem = selStep.termCondOper

        Call RemoveTempHandlers(True)
        Call RemoveVacHandlers(True)
        Call RemovePressureHandlers(True)
        Call RemoveDurationHandlers(True)
        Call RemoveTermHandlers(True)

        TempStateCheck()
        VacStateCheck()
        PressureStateCheck()
        Me.Refresh()
        DurationStateCheck()
    End Sub

    Sub SetStepValues()
        If loadedStepIndex <> -1 Then

            Dim selStep As CureStep = loadedEditProfile.CureSteps(loadedStepIndex)

            selStep.stepName = Txt_StepNameEdit.Text

            selStep.stepDuration = Txt_DurationEdit.Text

            selStep.tempSet.SetPoint = Txt_TempSetEdit.Text
            selStep.tempSet.PosTol = Txt_TempPosTolEdit.Text
            selStep.tempSet.NegTol = Txt_TempNegTolEdit.Text
            selStep.tempSet.RampRate = Txt_TempRampEdit.Text
            selStep.tempSet.RampPosTol = Txt_TempRampPosTolEdit.Text
            selStep.tempSet.RampNegTol = Txt_TempRampNegTolEdit.Text

            selStep.pressureSet.SetPoint = Txt_pressureSetEdit.Text
            selStep.pressureSet.PosTol = Txt_pressurePosTolEdit.Text
            selStep.pressureSet.NegTol = Txt_pressureNegTolEdit.Text
            selStep.pressureSet.RampRate = Txt_pressureRampEdit.Text
            selStep.pressureSet.RampPosTol = Txt_pressureRampPosTolEdit.Text
            selStep.pressureSet.RampNegTol = Txt_pressureRampNegTolEdit.Text

            selStep.vacSet.SetPoint = Txt_vacSetEdit.Text
            selStep.vacSet.PosTol = Txt_vacPosTolEdit.Text
            selStep.vacSet.NegTol = Txt_vacNegTolEdit.Text

            selStep.termCond1Type = Combo_Term1Type.SelectedItem
            selStep.termCond1Condition = Combo_Term1Cond.SelectedItem
            selStep.termCond1Goal = Txt_Term1Goal.Text
            selStep.termCond1Modifier = Combo_Term1Mod.SelectedItem

            selStep.termCond2Type = Combo_Term2Type.SelectedItem
            selStep.termCond2Condition = Combo_Term2Cond.SelectedItem
            selStep.termCond2Goal = Txt_Term2Goal.Text
            selStep.termCond2Modifier = Combo_Term2Mod.SelectedItem

            selStep.termCondOper = Combo_TermOper.SelectedItem
        End If
    End Sub

    Private Sub Txt_StepNameEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_StepNameEdit.TextChanged
        RemoveHandler List_StepsEdit.SelectedIndexChanged, AddressOf List_StepsEdit_SelectedIndexChanged
        List_StepsEdit.Items.Item(loadedStepIndex) = Txt_StepNameEdit.Text
        AddHandler List_StepsEdit.SelectedIndexChanged, AddressOf List_StepsEdit_SelectedIndexChanged
        Me.Refresh()
    End Sub
#End Region

#Region "Temp step intelligence"
    Private Sub Check_tempStepEdit_CheckedChanged(sender As Object, e As EventArgs) Handles Check_tempStepEdit.CheckedChanged
        Call RemoveTempHandlers()

        If Not Check_tempStepEdit.Checked Then
            Call DisableStepTemp()
        Else
            Txt_TempSetEdit.Text = 0
            Txt_TempPosTolEdit.Text = -1
            Txt_TempNegTolEdit.Text = -1
        End If

        Call RemoveTempHandlers(True)

        Call TempStateCheck()
    End Sub

    Private Sub Txt_TempSetEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_TempSetEdit.TextChanged
        Call TempStateCheck()
    End Sub

    Private Sub Txt_TempPosTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_TempPosTolEdit.TextChanged
        Call TempStateCheck()
    End Sub

    Private Sub Txt_TempNegTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_TempNegTolEdit.TextChanged
        Call TempStateCheck()
    End Sub

    Private Sub Txt_TempRampEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_TempRampEdit.TextChanged
        Call TempStateCheck()
    End Sub

    Private Sub Txt_TempRampPosTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_TempRampPosTolEdit.TextChanged
        Call TempStateCheck()
    End Sub

    Private Sub Txt_TempRampNegTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_TempRampNegTolEdit.TextChanged
        Call TempStateCheck()
    End Sub

    Function TempStateCheck() As Boolean
        'Check temp set value
        If Txt_TempSetEdit.Text = "-1" Then
            Call DisableStepTemp()
            Return True
        Else
            Check_tempStepEdit.Checked = True
            Txt_TempSetEdit.Enabled = True
            Txt_TempPosTolEdit.Enabled = True
            Txt_TempNegTolEdit.Enabled = True
            Txt_TempRampEdit.Enabled = True
            Txt_TempRampPosTolEdit.Enabled = True
            Txt_TempRampNegTolEdit.Enabled = True
        End If

        'Check to make sure entries are numeric
        If Not NumericTxtBoxCheck(Txt_TempSetEdit) Or
            Not NumericTxtBoxCheck(Txt_TempPosTolEdit) Or
            Not NumericTxtBoxCheck(Txt_TempNegTolEdit) Or
            Not NumericTxtBoxCheck(Txt_TempRampEdit) Or
            Not NumericTxtBoxCheck(Txt_TempRampPosTolEdit) Or
            Not NumericTxtBoxCheck(Txt_TempRampNegTolEdit) Then

            Return False
        End If

        'Check neg tol value to see if max only checked
        If Math.Abs(CDbl(Txt_TempNegTolEdit.Text)) = Math.Abs(CDbl(Txt_TempSetEdit.Text)) Then
            Txt_TempNegTolEdit.BackColor = Color.Plum
            Label_TempMaxOnlyEdit.Visible = True
        Else
            Txt_TempNegTolEdit.BackColor = SystemColors.Window
            Label_TempMaxOnlyEdit.Visible = False
        End If

        'Check pos tol value to see if min only checked
        If Math.Abs(CDbl(Txt_TempPosTolEdit.Text)) = Math.Abs(CDbl(Txt_TempSetEdit.Text)) Then
            Txt_TempPosTolEdit.BackColor = Color.Plum
            Label_TempMinOnlyEdit.Visible = True
        Else
            Txt_TempPosTolEdit.BackColor = SystemColors.Window
            Label_TempMinOnlyEdit.Visible = False
        End If

        'If ramp is 0 tolerance is 0
        If Txt_TempRampEdit.Text = 0 Then
            Txt_TempRampPosTolEdit.Text = 0
            Txt_TempRampNegTolEdit.Text = 0
            Txt_TempRampPosTolEdit.Enabled = False
            Txt_TempRampNegTolEdit.Enabled = False
        Else
            Txt_TempRampPosTolEdit.Enabled = True
            Txt_TempRampNegTolEdit.Enabled = True
        End If


        Return True
    End Function

    Sub DisableStepTemp()
        Call RemoveTempHandlers()

        Check_tempStepEdit.Checked = False
        Txt_TempSetEdit.Text = -1
        Txt_TempPosTolEdit.Text = 0
        Txt_TempNegTolEdit.Text = 0
        Txt_TempRampEdit.Text = 0
        Txt_TempRampPosTolEdit.Text = 0
        Txt_TempRampNegTolEdit.Text = 0

        Txt_TempSetEdit.Enabled = False
        Txt_TempPosTolEdit.Enabled = False
        Txt_TempNegTolEdit.Enabled = False
        Txt_TempRampEdit.Enabled = False
        Txt_TempRampPosTolEdit.Enabled = False
        Txt_TempRampNegTolEdit.Enabled = False

        NumericTxtBoxCheck(Txt_TempSetEdit)
        NumericTxtBoxCheck(Txt_TempPosTolEdit)
        NumericTxtBoxCheck(Txt_TempNegTolEdit)
        NumericTxtBoxCheck(Txt_TempRampEdit)
        NumericTxtBoxCheck(Txt_TempRampPosTolEdit)
        NumericTxtBoxCheck(Txt_TempRampNegTolEdit)

        Label_TempMaxOnlyEdit.Visible = False
        Label_TempMinOnlyEdit.Visible = False

        Call RemoveTempHandlers(True)
    End Sub


    Sub RemoveTempHandlers(Optional Invert As Boolean = False)
        If Not Invert Then
            RemoveHandler Txt_TempSetEdit.TextChanged, AddressOf Txt_TempSetEdit_TextChanged
            RemoveHandler Txt_TempPosTolEdit.TextChanged, AddressOf Txt_TempPosTolEdit_TextChanged
            RemoveHandler Txt_TempNegTolEdit.TextChanged, AddressOf Txt_TempNegTolEdit_TextChanged
            RemoveHandler Txt_TempRampEdit.TextChanged, AddressOf Txt_TempRampEdit_TextChanged
        Else
            AddHandler Txt_TempSetEdit.TextChanged, AddressOf Txt_TempSetEdit_TextChanged
            AddHandler Txt_TempPosTolEdit.TextChanged, AddressOf Txt_TempPosTolEdit_TextChanged
            AddHandler Txt_TempNegTolEdit.TextChanged, AddressOf Txt_TempNegTolEdit_TextChanged
            AddHandler Txt_TempRampEdit.TextChanged, AddressOf Txt_TempRampEdit_TextChanged
        End If

    End Sub
#End Region

#Region "Pressure step intelligence"
    Private Sub Check_pressureStepEdit_CheckedChanged(sender As Object, e As EventArgs) Handles Check_pressureStepEdit.CheckedChanged
        Call RemovePressureHandlers()

        If Not Check_pressureStepEdit.Checked Then
            Call DisableStepPressure()
        Else
            Txt_pressureSetEdit.Text = 0
            Txt_pressurePosTolEdit.Text = -1
            Txt_pressureNegTolEdit.Text = -1
        End If

        Call RemovePressureHandlers(True)

        Call PressureStateCheck()
    End Sub

    Private Sub Txt_PressureSetEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_pressureSetEdit.TextChanged
        Call PressureStateCheck()
    End Sub

    Private Sub Txt_PressurePosTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_pressurePosTolEdit.TextChanged
        Call PressureStateCheck()
    End Sub

    Private Sub Txt_PressureNegTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_pressureNegTolEdit.TextChanged
        Call PressureStateCheck()
    End Sub

    Private Sub Txt_PressureRampEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_pressureRampEdit.TextChanged
        Call PressureStateCheck()
    End Sub

    Private Sub Txt_PressureRampPosTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_pressureRampPosTolEdit.TextChanged
        Call PressureStateCheck()
    End Sub

    Private Sub Txt_PressureRampNegTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_pressureRampNegTolEdit.TextChanged
        Call PressureStateCheck()
    End Sub

    Function PressureStateCheck() As Boolean

        'Check pressure set value
        If Txt_pressureSetEdit.Text = "-1" Then
            Call DisableStepPressure()
            Return True
        Else
            Check_pressureStepEdit.Checked = True
            Txt_pressureSetEdit.Enabled = True
            Txt_pressurePosTolEdit.Enabled = True
            Txt_pressureNegTolEdit.Enabled = True
            Txt_pressureRampEdit.Enabled = True
            Txt_pressureRampPosTolEdit.Enabled = True
            Txt_pressureRampNegTolEdit.Enabled = True
        End If

        'Check to make sure entries are numeric
        If Not NumericTxtBoxCheck(Txt_pressureSetEdit) Or
            Not NumericTxtBoxCheck(Txt_pressurePosTolEdit) Or
            Not NumericTxtBoxCheck(Txt_pressureNegTolEdit) Or
            Not NumericTxtBoxCheck(Txt_pressureRampEdit) Or
            Not NumericTxtBoxCheck(Txt_pressureRampPosTolEdit) Or
            Not NumericTxtBoxCheck(Txt_pressureRampNegTolEdit) Then

            Return False
        End If

        'Check neg tol value to see if max only checked
        If Math.Abs(CDbl(Txt_pressureNegTolEdit.Text)) = Math.Abs(CDbl(Txt_pressureSetEdit.Text)) Then
            Txt_pressureNegTolEdit.BackColor = Color.Plum
            Label_PressureMaxOnlyEdit.Visible = True
        Else
            Txt_pressureNegTolEdit.BackColor = SystemColors.Window
            Label_PressureMaxOnlyEdit.Visible = False
        End If

        'Check pos tol value to see if min only checked
        If Math.Abs(CDbl(Txt_pressurePosTolEdit.Text)) = Math.Abs(CDbl(Txt_pressureSetEdit.Text)) Then
            Txt_pressurePosTolEdit.BackColor = Color.Plum
            Label_PressureMinOnlyEdit.Visible = True
        Else
            Txt_pressurePosTolEdit.BackColor = SystemColors.Window
            Label_PressureMinOnlyEdit.Visible = False
        End If

        'If ramp is 0 tolerance is 0
        If Txt_pressureRampEdit.Text = 0 Then
            Txt_pressureRampPosTolEdit.Text = 0
            Txt_pressureRampNegTolEdit.Text = 0
            Txt_pressureRampPosTolEdit.Enabled = False
            Txt_pressureRampNegTolEdit.Enabled = False
        Else
            Txt_pressureRampPosTolEdit.Enabled = True
            Txt_pressureRampNegTolEdit.Enabled = True
        End If


        Return True
    End Function

    Sub DisableStepPressure()
        Call RemovePressureHandlers()
        Check_pressureStepEdit.Checked = False
        Txt_pressureSetEdit.Text = -1
        Txt_pressurePosTolEdit.Text = 0
        Txt_pressureNegTolEdit.Text = 0
        Txt_pressureRampEdit.Text = 0
        Txt_pressureRampPosTolEdit.Text = 0
        Txt_pressureRampNegTolEdit.Text = 0

        Txt_pressureSetEdit.Enabled = False
        Txt_pressurePosTolEdit.Enabled = False
        Txt_pressureNegTolEdit.Enabled = False
        Txt_pressureRampEdit.Enabled = False
        Txt_pressureRampPosTolEdit.Enabled = False
        Txt_pressureRampNegTolEdit.Enabled = False

        Txt_pressurePosTolEdit.BackColor = SystemColors.Window
        Txt_pressureNegTolEdit.BackColor = SystemColors.Window

        Label_PressureMaxOnlyEdit.Visible = False
        Label_PressureMinOnlyEdit.Visible = False

        Call RemovePressureHandlers(True)
    End Sub


    Sub RemovePressureHandlers(Optional Invert As Boolean = False)
        If Not Invert Then
            RemoveHandler Txt_pressureSetEdit.TextChanged, AddressOf Txt_PressureSetEdit_TextChanged
            RemoveHandler Txt_pressurePosTolEdit.TextChanged, AddressOf Txt_PressurePosTolEdit_TextChanged
            RemoveHandler Txt_pressureNegTolEdit.TextChanged, AddressOf Txt_PressureNegTolEdit_TextChanged
            RemoveHandler Txt_pressureRampEdit.TextChanged, AddressOf Txt_PressureRampEdit_TextChanged
        Else
            AddHandler Txt_pressureSetEdit.TextChanged, AddressOf Txt_PressureSetEdit_TextChanged
            AddHandler Txt_pressurePosTolEdit.TextChanged, AddressOf Txt_PressurePosTolEdit_TextChanged
            AddHandler Txt_pressureNegTolEdit.TextChanged, AddressOf Txt_PressureNegTolEdit_TextChanged
            AddHandler Txt_pressureRampEdit.TextChanged, AddressOf Txt_PressureRampEdit_TextChanged
        End If

    End Sub
#End Region

#Region "Vacuum step intelligence"
    Private Sub Check_VacStepEdit_CheckedChanged(sender As Object, e As EventArgs) Handles Check_vacStepEdit.CheckedChanged
        Call RemoveVacHandlers()

        If Not Check_vacStepEdit.Checked Then
            Call DisableStepVac()
        Else
            Txt_vacSetEdit.Text = 0
            Txt_vacPosTolEdit.Text = -1
            Txt_vacNegTolEdit.Text = -1
        End If

        Call RemoveVacHandlers(True)

        Call VacStateCheck()
    End Sub

    Private Sub Txt_VacSetEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_vacSetEdit.TextChanged
        Call VacStateCheck()
    End Sub

    Private Sub Txt_VacPosTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_vacPosTolEdit.TextChanged
        Call VacStateCheck()
    End Sub

    Private Sub Txt_VacNegTolEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_vacNegTolEdit.TextChanged
        Call VacStateCheck()
    End Sub


    Function VacStateCheck() As Boolean

        'Check vac set value
        If Txt_vacSetEdit.Text = "-1" Then
            Call DisableStepVac()
            Return True
        Else
            Check_vacStepEdit.Checked = True
            Txt_vacSetEdit.Enabled = True
            Txt_vacPosTolEdit.Enabled = True
            Txt_vacNegTolEdit.Enabled = True
        End If

        'Check to make sure entries are numeric
        If Not NumericTxtBoxCheck(Txt_vacSetEdit) Or
            Not NumericTxtBoxCheck(Txt_vacPosTolEdit) Or
            Not NumericTxtBoxCheck(Txt_vacNegTolEdit) Then

            Return False
        End If

        'Check neg tol value to see if max only checked
        If Math.Abs(CDbl(Txt_vacNegTolEdit.Text)) = Math.Abs(CDbl(Txt_vacSetEdit.Text)) Then
            Txt_vacNegTolEdit.BackColor = Color.Plum
            Label_VacMinOnlyEdit.Visible = True
        Else
            Txt_vacNegTolEdit.BackColor = SystemColors.Window
            Label_VacMinOnlyEdit.Visible = False
        End If

        'Check pos tol value to see if min only checked
        If Math.Abs(CDbl(Txt_vacPosTolEdit.Text)) = Math.Abs(CDbl(Txt_vacSetEdit.Text)) Then
            Txt_vacPosTolEdit.BackColor = Color.Plum
            Label_VacMaxOnlyEdit.Visible = True
        Else
            Txt_vacPosTolEdit.BackColor = SystemColors.Window
            Label_VacMaxOnlyEdit.Visible = False
        End If


        Return True
    End Function

    Sub DisableStepVac()
        Call RemoveVacHandlers()

        Check_vacStepEdit.Checked = False
        Txt_vacSetEdit.Text = -1
        Txt_vacPosTolEdit.Text = 0
        Txt_vacNegTolEdit.Text = 0

        Txt_vacSetEdit.Enabled = False
        Txt_vacPosTolEdit.Enabled = False
        Txt_vacNegTolEdit.Enabled = False


        Txt_vacPosTolEdit.BackColor = SystemColors.Window
        Txt_vacNegTolEdit.BackColor = SystemColors.Window

        Label_VacMinOnlyEdit.Visible = False
        Label_VacMaxOnlyEdit.Visible = False

        Call RemoveVacHandlers(True)
    End Sub


    Sub RemoveVacHandlers(Optional Invert As Boolean = False)
        If Not Invert Then
            RemoveHandler Txt_vacSetEdit.TextChanged, AddressOf Txt_VacSetEdit_TextChanged
            RemoveHandler Txt_vacPosTolEdit.TextChanged, AddressOf Txt_VacPosTolEdit_TextChanged
            RemoveHandler Txt_vacNegTolEdit.TextChanged, AddressOf Txt_VacNegTolEdit_TextChanged
        Else
            AddHandler Txt_vacSetEdit.TextChanged, AddressOf Txt_VacSetEdit_TextChanged
            AddHandler Txt_vacPosTolEdit.TextChanged, AddressOf Txt_VacPosTolEdit_TextChanged
            AddHandler Txt_vacNegTolEdit.TextChanged, AddressOf Txt_VacNegTolEdit_TextChanged
        End If

    End Sub
#End Region

#Region "Duration step intelligence"
    Private Sub Chk_DurationDisabled_CheckedChanged(sender As Object, e As EventArgs) Handles Chk_DurationDisabled.CheckedChanged
        Call RemoveDurationHandlers()

        If Not Chk_DurationDisabled.Checked Then
            Txt_DurationEdit.Text = 0
            Txt_DurationEdit.Enabled = True
        Else
            Txt_DurationEdit.Text = -1
            Txt_DurationEdit.Enabled = False
        End If

        Call RemoveDurationHandlers(True)
    End Sub

    Private Sub Txt_DurationEdit_TextChanged(sender As Object, e As EventArgs) Handles Txt_DurationEdit.TextChanged
        Call RemoveDurationHandlers()

        If NumericTxtBoxCheck(Txt_DurationEdit) Then
            If CDbl(Txt_DurationEdit.Text) = -1 Then
                Chk_DurationDisabled.Checked = True
            ElseIf CDbl(Txt_DurationEdit.Text) < 0 Then
                Txt_DurationEdit.Text = 0
            End If
        End If

        Call RemoveDurationHandlers(True)
    End Sub

    Sub DurationStateCheck()
        If Txt_DurationEdit.Text = "-1" Then
            Txt_DurationEdit.Enabled = False
            Chk_DurationDisabled.Checked = True
        End If
    End Sub

    Sub RemoveDurationHandlers(Optional Invert As Boolean = False)
        If Not Invert Then
            RemoveHandler Chk_DurationDisabled.CheckedChanged, AddressOf Chk_DurationDisabled_CheckedChanged
            RemoveHandler Txt_DurationEdit.TextChanged, AddressOf Txt_DurationEdit_TextChanged

        Else
            AddHandler Chk_DurationDisabled.CheckedChanged, AddressOf Chk_DurationDisabled_CheckedChanged
            AddHandler Txt_DurationEdit.TextChanged, AddressOf Txt_DurationEdit_TextChanged
        End If

    End Sub

#End Region

#Region "Terminating step intelligence"

    Private Sub Combo_Term1Type_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_Term1Type.SelectedIndexChanged
        Call TermComboBoxOptionsUpdate()

        Call RemoveTermHandlers()

        If Combo_Term1Type.SelectedItem = "Time" Then
            Combo_Term1Cond.SelectedValue = "GREATER"
            Txt_Term1Goal.Text = "0"
            Combo_Term1Mod.SelectedIndex = 0

        ElseIf Combo_Term1Type.SelectedItem = "Temp" Then
            Combo_Term1Mod.SelectedIndex = 0

        ElseIf Combo_Term1Type.SelectedItem = "Press" Then
            Combo_Term1Mod.SelectedIndex = 0
        Else
            Throw New Exception("Combo_Term1Type set to unknown type")
            Call RemoveTermHandlers(True)
        End If

        Call RemoveTermHandlers(True)
    End Sub

    Private Sub Txt_Term1Goal_TextChanged(sender As Object, e As EventArgs) Handles Txt_Term1Goal.TextChanged
        Call RemoveTermHandlers()

        If NumericTxtBoxCheck(Txt_Term1Goal) Then
            If Combo_Term1Type.SelectedItem = "Time" And CDbl(Txt_Term1Goal.Text) < 0 Then
                Txt_Term1Goal.Text = 0
            End If
        End If

        Call RemoveTermHandlers(True)
    End Sub

    Private Sub Combo_Term2Type_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_Term2Type.SelectedIndexChanged
        Call TermComboBoxOptionsUpdate()

        Call RemoveTermHandlers()

        If Combo_Term2Type.SelectedItem = "None" Then
            Combo_Term2Cond.SelectedValue = "0"
            Txt_Term2Goal.Text = "0"
            Combo_Term2Mod.SelectedValue = "0"
        ElseIf Combo_Term2Type.SelectedItem = "Time" Then
            Combo_Term2Cond.SelectedValue = "GREATER"
            Txt_Term2Goal.Text = "0"
            Combo_Term2Mod.SelectedIndex = 0

        ElseIf Combo_Term2Type.SelectedItem = "Temp" Then
            Combo_Term2Mod.SelectedIndex = 0

        ElseIf Combo_Term2Type.SelectedItem = "Press" Then
            Combo_Term2Mod.SelectedIndex = 0
        Else
            Throw New Exception("Combo_Term2Type set to unknown type")
            Call RemoveTermHandlers(True)
        End If

        Call RemoveTermHandlers(True)
    End Sub

    Private Sub Txt_Term2Goal_TextChanged(sender As Object, e As EventArgs) Handles Txt_Term2Goal.TextChanged
        Call RemoveTermHandlers()

        If NumericTxtBoxCheck(Txt_Term2Goal) Then
            If Combo_Term2Type.SelectedItem = "Time" And CDbl(Txt_Term2Goal.Text) < 0 Then
                Txt_Term2Goal.Text = 0
            End If
        End If

        Call RemoveTermHandlers(True)
    End Sub

    Sub TermComboBoxOptionsUpdate()

        'Terminating conditions #1
        If Combo_Term1Type.SelectedItem = "Time" Then
            Combo_Term1Cond.Enabled = False

            Combo_Term1Mod.Items.Clear()
            Combo_Term1Mod.Items.Add("None")
            Combo_Term1Mod.Items.Add("Pass")

        ElseIf Combo_Term1Type.SelectedItem = "Temp" Then
            Combo_Term1Cond.Enabled = True

            Combo_Term1Mod.Items.Clear()
            Combo_Term1Mod.Items.Add("Lag")
            Combo_Term1Mod.Items.Add("Lead")
            Combo_Term1Mod.Items.Add("Air")

        ElseIf Combo_Term1Type.SelectedItem = "Press" Then
            Combo_Term1Cond.Enabled = True

            Combo_Term1Mod.Items.Clear()
            Combo_Term1Mod.Items.Add("None")
        Else
            Throw New Exception("TermComboBoxOptions unknown termination type")
        End If

        'Terminating conditions #2
        If Combo_Term2Type.SelectedItem = "None" Then
            Combo_Term2Cond.Enabled = False
            Txt_Term2Goal.Enabled = False
            Combo_Term2Mod.Enabled = False

        ElseIf Combo_Term2Type.SelectedItem = "Time" Then
            Combo_Term2Cond.Enabled = False

            Txt_Term2Goal.Enabled = True

            Combo_Term2Mod.Enabled = True
            Combo_Term2Mod.Items.Clear()
            Combo_Term2Mod.Items.Add("None")
            Combo_Term2Mod.Items.Add("Pass")

        ElseIf Combo_Term2Type.SelectedItem = "Temp" Then
            Combo_Term2Cond.Enabled = True

            Txt_Term2Goal.Enabled = True

            Combo_Term2Mod.Enabled = True
            Combo_Term2Mod.Items.Clear()
            Combo_Term2Mod.Items.Add("Lag")
            Combo_Term2Mod.Items.Add("Lead")
            Combo_Term2Mod.Items.Add("Air")

        ElseIf Combo_Term2Type.SelectedItem = "Press" Then
            Combo_Term2Cond.Enabled = True

            Txt_Term2Goal.Enabled = True

            Combo_Term2Mod.Enabled = True
            Combo_Term2Mod.Items.Clear()
            Combo_Term2Mod.Items.Add("None")
        Else
            Throw New Exception("TermComboBoxOptions unknown termination type")
        End If

    End Sub

    Sub RemoveTermHandlers(Optional Invert As Boolean = False)
        If Not Invert Then
            RemoveHandler Combo_Term1Type.SelectedIndexChanged, AddressOf Combo_Term1Type_SelectedIndexChanged
            RemoveHandler Txt_Term1Goal.TextChanged, AddressOf Txt_Term1Goal_TextChanged
            RemoveHandler Combo_Term2Type.SelectedIndexChanged, AddressOf Combo_Term2Type_SelectedIndexChanged
            RemoveHandler Txt_Term2Goal.TextChanged, AddressOf Txt_Term2Goal_TextChanged

        Else
            AddHandler Combo_Term1Type.SelectedIndexChanged, AddressOf Combo_Term1Type_SelectedIndexChanged
            AddHandler Txt_Term1Goal.TextChanged, AddressOf Txt_Term1Goal_TextChanged
            AddHandler Combo_Term2Type.SelectedIndexChanged, AddressOf Combo_Term2Type_SelectedIndexChanged
            AddHandler Txt_Term2Goal.TextChanged, AddressOf Txt_Term2Goal_TextChanged
        End If

    End Sub

#End Region

#End Region
End Class


Module ArrayExtensions

    ''' <summary>
    ''' Adds values of a 1D array to a 2D array. 1D array length must match 1st dimension of 2D array (if they do not match, values will not be added).
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



