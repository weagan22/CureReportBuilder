Option Explicit On

Imports System.Runtime.CompilerServices
Imports System.Data.SqlClient
Imports System.Threading
Imports System.Text.RegularExpressions

Public Class MainForm

    Public cureProfiles() As CureProfile

    Dim mainCureCheck As CureCheck = New CureCheck()

    Public autoclaveSNum As New Dictionary(Of String, Integer) From {{"Controller", 2601}, {"Air TC1", 2600}, {"Air TC2", 2603}, {"PT", 2599}, {"VT1", 2593}, {"VT2", 2598}, {"VT3", 2595}, {"VT4", 2594}, {"VT5", 2597}, {"VT6", 2596}}

    Dim STOP_Txt_FilePath_TextChanged As Boolean = False


    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim vers As Version = My.Application.Info.Version
        Me.Text = "Cure Report Builder " & vers.Major & "." & vers.Minor & "." & vers.Build

        Call errorReset()

        TabControl1.TabPages.Remove(TabPage3)

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


    Function getEpicorData(jobNum As Integer) As DataRow

        If Len(jobNum.ToString) <> 6 And System.Text.RegularExpressions.Regex.IsMatch(jobNum.ToString, "^[0-9]+$") Then
            Throw New Exception("Job number for Epicor must be a 6 digit number.")
        End If

        Dim queryResult As DataTable = New DataTable()

        Dim db As SqlConnection = New SqlConnection("Data Source = MAUI; Initial Catalog=EpicorERP; User ID = Reporting; Password=$ystima1; Integrated Security=false; trusted_connection=false;")

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

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim toRunList() As String
        Dim cureDataPath As String = ""
        Dim cureParamPath As String = ""

        If System.IO.Directory.Exists(Txt_CureParamPath.Text) AndAlso System.IO.File.Exists(Txt_CureParamPath.Text & "\ToRunList.log") And System.IO.Directory.Exists(Txt_CureDataPath.Text) Then
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

        'Compare with runList and and toRunList
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Btn_Run.Click
        testRun()
        Box_RunLine.Enabled = False
    End Sub

    Sub testRun()

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














#Region "Cure profile inport/export"
    Sub outputCureProfiles(inPath As String)


        Dim i As Integer
        For i = 0 To UBound(cureProfiles)
            Dim outputWriter As IO.StreamWriter = New System.IO.StreamWriter(inPath & "/" & cureProfiles(i).Name & ".cprof")
            outputWriter.Write(cureProfiles(i).SerializeCure)
            outputWriter.Close()
        Next


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
            Try
                Combo_CureProfile.SelectedIndex = Combo_CureProfile.Items.IndexOf("CURE A1")
                Combo_CureProfileEdit.SelectedIndex = Combo_CureProfile.Items.IndexOf("CURE A1")
            Catch ex As Exception
                Combo_CureProfile.SelectedIndex = 0
                Combo_CureProfileEdit.SelectedIndex = 0
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

        Call errorReset()

        Txt_CureDoc.Text = mainCureCheck.curePro.cureDoc
        Txt_DocRev.Text = mainCureCheck.curePro.cureDocRev

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

    Private Sub Txt_OutputPath_TextChanged(sender As Object, e As EventArgs) Handles Txt_OutputPath.TextChanged
        If IO.Directory.Exists(Txt_OutputPath.Text) Then
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



