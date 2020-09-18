<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TableLayoutPanel7 = New System.Windows.Forms.TableLayoutPanel()
        Me.Box_CureDataLocation = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Txt_FilePath = New System.Windows.Forms.TextBox()
        Me.Btn_OpenFile = New System.Windows.Forms.Button()
        Me.Box_RunParams = New System.Windows.Forms.GroupBox()
        Me.Box_PartInfo = New System.Windows.Forms.GroupBox()
        Me.Btn_ClearCells = New System.Windows.Forms.Button()
        Me.Box_PartNumber = New System.Windows.Forms.GroupBox()
        Me.Txt_PartNumber = New System.Windows.Forms.TextBox()
        Me.Box_Revision = New System.Windows.Forms.GroupBox()
        Me.Txt_Revision = New System.Windows.Forms.TextBox()
        Me.Box_PartDesc = New System.Windows.Forms.GroupBox()
        Me.Txt_PartDesc = New System.Windows.Forms.TextBox()
        Me.Box_Qty = New System.Windows.Forms.GroupBox()
        Me.Txt_Qty = New System.Windows.Forms.TextBox()
        Me.Box_Vac = New System.Windows.Forms.GroupBox()
        Me.Data_Vac = New System.Windows.Forms.DataGridView()
        Me.DataGridViewCheckBoxColumn1 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Box_TC = New System.Windows.Forms.GroupBox()
        Me.Data_TC = New System.Windows.Forms.DataGridView()
        Me.Used = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Modifier = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Box_JobNumber = New System.Windows.Forms.GroupBox()
        Me.Txt_JobNumber = New System.Windows.Forms.TextBox()
        Me.Box_ProgramNumber = New System.Windows.Forms.GroupBox()
        Me.Txt_ProgramNumber = New System.Windows.Forms.TextBox()
        Me.Box_DataRecorder = New System.Windows.Forms.GroupBox()
        Me.Txt_DataRecorder = New System.Windows.Forms.TextBox()
        Me.Box_RunLine = New System.Windows.Forms.TableLayoutPanel()
        Me.Btn_Run = New System.Windows.Forms.Button()
        Me.Box_Technician = New System.Windows.Forms.GroupBox()
        Me.Txt_Technician = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel6 = New System.Windows.Forms.TableLayoutPanel()
        Me.Box_CureProfChoice = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.Box_CureProfile = New System.Windows.Forms.GroupBox()
        Me.Combo_CureProfile = New System.Windows.Forms.ComboBox()
        Me.TableLayoutPanel4 = New System.Windows.Forms.TableLayoutPanel()
        Me.Box_DocRev = New System.Windows.Forms.GroupBox()
        Me.Txt_DocRev = New System.Windows.Forms.TextBox()
        Me.Box_CureDoc = New System.Windows.Forms.GroupBox()
        Me.Txt_CureDoc = New System.Windows.Forms.TextBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Box_TemplatePath = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel5 = New System.Windows.Forms.TableLayoutPanel()
        Me.Txt_TemplatePath = New System.Windows.Forms.TextBox()
        Me.Box_CureProfiles = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.Txt_CureProfilesPath = New System.Windows.Forms.TextBox()
        Me.Btn_LoadProfileFiles = New System.Windows.Forms.Button()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Box_vacStepEdit = New System.Windows.Forms.GroupBox()
        Me.Check_vacMinEdit = New System.Windows.Forms.CheckBox()
        Me.Check_vacMaxEdit = New System.Windows.Forms.CheckBox()
        Me.Check_vacStepEdit = New System.Windows.Forms.CheckBox()
        Me.Box_vacNegTolEdit = New System.Windows.Forms.GroupBox()
        Me.Txt_vacNegTolEdit = New System.Windows.Forms.TextBox()
        Me.Box_vacPosTolEdit = New System.Windows.Forms.GroupBox()
        Me.Txt_vacPosTolEdit = New System.Windows.Forms.TextBox()
        Me.Box_vacSetEdit = New System.Windows.Forms.GroupBox()
        Me.Txt_vacSetEdit = New System.Windows.Forms.TextBox()
        Me.Box_pressureStepEdit = New System.Windows.Forms.GroupBox()
        Me.Check_pressureMinEdit = New System.Windows.Forms.CheckBox()
        Me.Check_pressureMaxEdit = New System.Windows.Forms.CheckBox()
        Me.Check_pressureStepEdit = New System.Windows.Forms.CheckBox()
        Me.Box_pressureRampNegTolEdit = New System.Windows.Forms.GroupBox()
        Me.Txt_pressureRampNegTolEdit = New System.Windows.Forms.TextBox()
        Me.Box_pressureRampPosTolEdit = New System.Windows.Forms.GroupBox()
        Me.Txt_pressureRampPosTolEdit = New System.Windows.Forms.TextBox()
        Me.Box_pressureRampEdit = New System.Windows.Forms.GroupBox()
        Me.Txt_pressureRampEdit = New System.Windows.Forms.TextBox()
        Me.Box_pressureNegTolEdit = New System.Windows.Forms.GroupBox()
        Me.Txt_pressureNegTolEdit = New System.Windows.Forms.TextBox()
        Me.Box_pressurePosTolEdit = New System.Windows.Forms.GroupBox()
        Me.Txt_pressurePosTolEdit = New System.Windows.Forms.TextBox()
        Me.Box_pressureSetEdit = New System.Windows.Forms.GroupBox()
        Me.Txt_pressureSetEdit = New System.Windows.Forms.TextBox()
        Me.Box_tempStepEdit = New System.Windows.Forms.GroupBox()
        Me.Check_tempMinEdit = New System.Windows.Forms.CheckBox()
        Me.Check_tempMaxEdit = New System.Windows.Forms.CheckBox()
        Me.Check_tempStepEdit = New System.Windows.Forms.CheckBox()
        Me.Box_TempRampNegTolEdit = New System.Windows.Forms.GroupBox()
        Me.Txt_TempRampNegTolEdit = New System.Windows.Forms.TextBox()
        Me.Box_TempRampPosTolEdit = New System.Windows.Forms.GroupBox()
        Me.Txt_TempRampPosTolEdit = New System.Windows.Forms.TextBox()
        Me.Box_TempRampEdit = New System.Windows.Forms.GroupBox()
        Me.Txt_TempRampEdit = New System.Windows.Forms.TextBox()
        Me.Box_TempNegTolEdit = New System.Windows.Forms.GroupBox()
        Me.Txt_TempNegTolEdit = New System.Windows.Forms.TextBox()
        Me.Box_TempPosTolEdit = New System.Windows.Forms.GroupBox()
        Me.Txt_TempPosTolEdit = New System.Windows.Forms.TextBox()
        Me.Box_TempSetEdit = New System.Windows.Forms.GroupBox()
        Me.Txt_TempSetEdit = New System.Windows.Forms.TextBox()
        Me.Box_StepNameEdit = New System.Windows.Forms.GroupBox()
        Me.Txt_StepNameEdit = New System.Windows.Forms.TextBox()
        Me.Box_CheckEdit = New System.Windows.Forms.GroupBox()
        Me.Check_Vac = New System.Windows.Forms.CheckBox()
        Me.Check_TempEdit = New System.Windows.Forms.CheckBox()
        Me.Check_PressureEdit = New System.Windows.Forms.CheckBox()
        Me.Box_CureNameEdit = New System.Windows.Forms.GroupBox()
        Me.Txt_CureNameEdit = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Txt_DocRevEdit = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Txt_CureDocEdit = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Combo_CureProfileEdit = New System.Windows.Forms.ComboBox()
        Me.OpenCSVFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.OpenCureProfileFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.MenuStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TableLayoutPanel7.SuspendLayout()
        Me.Box_CureDataLocation.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Box_RunParams.SuspendLayout()
        Me.Box_PartInfo.SuspendLayout()
        Me.Box_PartNumber.SuspendLayout()
        Me.Box_Revision.SuspendLayout()
        Me.Box_PartDesc.SuspendLayout()
        Me.Box_Qty.SuspendLayout()
        Me.Box_Vac.SuspendLayout()
        CType(Me.Data_Vac, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Box_TC.SuspendLayout()
        CType(Me.Data_TC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Box_JobNumber.SuspendLayout()
        Me.Box_ProgramNumber.SuspendLayout()
        Me.Box_DataRecorder.SuspendLayout()
        Me.Box_RunLine.SuspendLayout()
        Me.Box_Technician.SuspendLayout()
        Me.TableLayoutPanel6.SuspendLayout()
        Me.Box_CureProfChoice.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.Box_CureProfile.SuspendLayout()
        Me.TableLayoutPanel4.SuspendLayout()
        Me.Box_DocRev.SuspendLayout()
        Me.Box_CureDoc.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        Me.Box_TemplatePath.SuspendLayout()
        Me.TableLayoutPanel5.SuspendLayout()
        Me.Box_CureProfiles.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.Box_vacStepEdit.SuspendLayout()
        Me.Box_vacNegTolEdit.SuspendLayout()
        Me.Box_vacPosTolEdit.SuspendLayout()
        Me.Box_vacSetEdit.SuspendLayout()
        Me.Box_pressureStepEdit.SuspendLayout()
        Me.Box_pressureRampNegTolEdit.SuspendLayout()
        Me.Box_pressureRampPosTolEdit.SuspendLayout()
        Me.Box_pressureRampEdit.SuspendLayout()
        Me.Box_pressureNegTolEdit.SuspendLayout()
        Me.Box_pressurePosTolEdit.SuspendLayout()
        Me.Box_pressureSetEdit.SuspendLayout()
        Me.Box_tempStepEdit.SuspendLayout()
        Me.Box_TempRampNegTolEdit.SuspendLayout()
        Me.Box_TempRampPosTolEdit.SuspendLayout()
        Me.Box_TempRampEdit.SuspendLayout()
        Me.Box_TempNegTolEdit.SuspendLayout()
        Me.Box_TempPosTolEdit.SuspendLayout()
        Me.Box_TempSetEdit.SuspendLayout()
        Me.Box_StepNameEdit.SuspendLayout()
        Me.Box_CheckEdit.SuspendLayout()
        Me.Box_CureNameEdit.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(800, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(92, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 526)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(800, 22)
        Me.StatusStrip1.TabIndex = 1
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(48, 17)
        Me.ToolStripStatusLabel1.Text = "Status..."
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 24)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(800, 502)
        Me.TabControl1.TabIndex = 2
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.TableLayoutPanel7)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(792, 476)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Run Cure"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel7
        '
        Me.TableLayoutPanel7.ColumnCount = 2
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel7.Controls.Add(Me.Box_CureDataLocation, 1, 1)
        Me.TableLayoutPanel7.Controls.Add(Me.Box_RunParams, 1, 2)
        Me.TableLayoutPanel7.Controls.Add(Me.Box_RunLine, 1, 3)
        Me.TableLayoutPanel7.Controls.Add(Me.TableLayoutPanel6, 1, 0)
        Me.TableLayoutPanel7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel7.Location = New System.Drawing.Point(3, 3)
        Me.TableLayoutPanel7.Name = "TableLayoutPanel7"
        Me.TableLayoutPanel7.RowCount = 4
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 150.0!))
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 60.0!))
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 60.0!))
        Me.TableLayoutPanel7.Size = New System.Drawing.Size(786, 470)
        Me.TableLayoutPanel7.TabIndex = 13
        '
        'Box_CureDataLocation
        '
        Me.Box_CureDataLocation.Controls.Add(Me.TableLayoutPanel1)
        Me.Box_CureDataLocation.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Box_CureDataLocation.Enabled = False
        Me.Box_CureDataLocation.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Box_CureDataLocation.Location = New System.Drawing.Point(23, 153)
        Me.Box_CureDataLocation.Name = "Box_CureDataLocation"
        Me.Box_CureDataLocation.Size = New System.Drawing.Size(760, 54)
        Me.Box_CureDataLocation.TabIndex = 0
        Me.Box_CureDataLocation.TabStop = False
        Me.Box_CureDataLocation.Text = "Cure Data Location"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 75.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Txt_FilePath, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Btn_OpenFile, 1, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(3, 16)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(754, 35)
        Me.TableLayoutPanel1.TabIndex = 2
        '
        'Txt_FilePath
        '
        Me.Txt_FilePath.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_FilePath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Txt_FilePath.Location = New System.Drawing.Point(3, 7)
        Me.Txt_FilePath.Name = "Txt_FilePath"
        Me.Txt_FilePath.Size = New System.Drawing.Size(673, 20)
        Me.Txt_FilePath.TabIndex = 2
        Me.Txt_FilePath.Text = "File path..."
        '
        'Btn_OpenFile
        '
        Me.Btn_OpenFile.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Btn_OpenFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_OpenFile.Location = New System.Drawing.Point(682, 3)
        Me.Btn_OpenFile.Name = "Btn_OpenFile"
        Me.Btn_OpenFile.Size = New System.Drawing.Size(69, 29)
        Me.Btn_OpenFile.TabIndex = 3
        Me.Btn_OpenFile.Text = "Open File"
        Me.Btn_OpenFile.UseVisualStyleBackColor = True
        '
        'Box_RunParams
        '
        Me.Box_RunParams.Controls.Add(Me.Box_PartInfo)
        Me.Box_RunParams.Controls.Add(Me.Box_Vac)
        Me.Box_RunParams.Controls.Add(Me.Box_TC)
        Me.Box_RunParams.Controls.Add(Me.Box_JobNumber)
        Me.Box_RunParams.Controls.Add(Me.Box_ProgramNumber)
        Me.Box_RunParams.Controls.Add(Me.Box_DataRecorder)
        Me.Box_RunParams.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Box_RunParams.Location = New System.Drawing.Point(23, 213)
        Me.Box_RunParams.Name = "Box_RunParams"
        Me.Box_RunParams.Size = New System.Drawing.Size(760, 194)
        Me.Box_RunParams.TabIndex = 13
        Me.Box_RunParams.TabStop = False
        '
        'Box_PartInfo
        '
        Me.Box_PartInfo.Controls.Add(Me.Btn_ClearCells)
        Me.Box_PartInfo.Controls.Add(Me.Box_PartNumber)
        Me.Box_PartInfo.Controls.Add(Me.Box_Revision)
        Me.Box_PartInfo.Controls.Add(Me.Box_PartDesc)
        Me.Box_PartInfo.Controls.Add(Me.Box_Qty)
        Me.Box_PartInfo.Location = New System.Drawing.Point(6, 69)
        Me.Box_PartInfo.Name = "Box_PartInfo"
        Me.Box_PartInfo.Size = New System.Drawing.Size(382, 120)
        Me.Box_PartInfo.TabIndex = 13
        Me.Box_PartInfo.TabStop = False
        Me.Box_PartInfo.Text = "Part Info"
        '
        'Btn_ClearCells
        '
        Me.Btn_ClearCells.Location = New System.Drawing.Point(264, 24)
        Me.Btn_ClearCells.Name = "Btn_ClearCells"
        Me.Btn_ClearCells.Size = New System.Drawing.Size(104, 38)
        Me.Btn_ClearCells.TabIndex = 4
        Me.Btn_ClearCells.TabStop = False
        Me.Btn_ClearCells.Text = "Clear Cells"
        Me.Btn_ClearCells.UseVisualStyleBackColor = True
        '
        'Box_PartNumber
        '
        Me.Box_PartNumber.Controls.Add(Me.Txt_PartNumber)
        Me.Box_PartNumber.Location = New System.Drawing.Point(6, 19)
        Me.Box_PartNumber.Name = "Box_PartNumber"
        Me.Box_PartNumber.Size = New System.Drawing.Size(114, 44)
        Me.Box_PartNumber.TabIndex = 2
        Me.Box_PartNumber.TabStop = False
        Me.Box_PartNumber.Text = "Part Number"
        '
        'Txt_PartNumber
        '
        Me.Txt_PartNumber.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_PartNumber.Location = New System.Drawing.Point(3, 16)
        Me.Txt_PartNumber.Name = "Txt_PartNumber"
        Me.Txt_PartNumber.Size = New System.Drawing.Size(108, 20)
        Me.Txt_PartNumber.TabIndex = 7
        Me.Txt_PartNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Box_Revision
        '
        Me.Box_Revision.Controls.Add(Me.Txt_Revision)
        Me.Box_Revision.Location = New System.Drawing.Point(129, 19)
        Me.Box_Revision.Name = "Box_Revision"
        Me.Box_Revision.Size = New System.Drawing.Size(63, 44)
        Me.Box_Revision.TabIndex = 3
        Me.Box_Revision.TabStop = False
        Me.Box_Revision.Text = "Rev"
        '
        'Txt_Revision
        '
        Me.Txt_Revision.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_Revision.Location = New System.Drawing.Point(3, 16)
        Me.Txt_Revision.Name = "Txt_Revision"
        Me.Txt_Revision.Size = New System.Drawing.Size(57, 20)
        Me.Txt_Revision.TabIndex = 8
        Me.Txt_Revision.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Box_PartDesc
        '
        Me.Box_PartDesc.Controls.Add(Me.Txt_PartDesc)
        Me.Box_PartDesc.Location = New System.Drawing.Point(6, 69)
        Me.Box_PartDesc.Name = "Box_PartDesc"
        Me.Box_PartDesc.Size = New System.Drawing.Size(351, 44)
        Me.Box_PartDesc.TabIndex = 5
        Me.Box_PartDesc.TabStop = False
        Me.Box_PartDesc.Text = "Part Nomenclature"
        '
        'Txt_PartDesc
        '
        Me.Txt_PartDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_PartDesc.Location = New System.Drawing.Point(3, 16)
        Me.Txt_PartDesc.Name = "Txt_PartDesc"
        Me.Txt_PartDesc.Size = New System.Drawing.Size(345, 20)
        Me.Txt_PartDesc.TabIndex = 10
        Me.Txt_PartDesc.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Box_Qty
        '
        Me.Box_Qty.Controls.Add(Me.Txt_Qty)
        Me.Box_Qty.Location = New System.Drawing.Point(198, 20)
        Me.Box_Qty.Name = "Box_Qty"
        Me.Box_Qty.Size = New System.Drawing.Size(51, 44)
        Me.Box_Qty.TabIndex = 4
        Me.Box_Qty.TabStop = False
        Me.Box_Qty.Text = "Qty"
        '
        'Txt_Qty
        '
        Me.Txt_Qty.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_Qty.Location = New System.Drawing.Point(3, 16)
        Me.Txt_Qty.Name = "Txt_Qty"
        Me.Txt_Qty.Size = New System.Drawing.Size(45, 20)
        Me.Txt_Qty.TabIndex = 9
        Me.Txt_Qty.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Box_Vac
        '
        Me.Box_Vac.AutoSize = True
        Me.Box_Vac.Controls.Add(Me.Data_Vac)
        Me.Box_Vac.Location = New System.Drawing.Point(557, 19)
        Me.Box_Vac.Name = "Box_Vac"
        Me.Box_Vac.Size = New System.Drawing.Size(160, 169)
        Me.Box_Vac.TabIndex = 9
        Me.Box_Vac.TabStop = False
        Me.Box_Vac.Text = "Vac Monitors"
        Me.Box_Vac.Visible = False
        '
        'Data_Vac
        '
        Me.Data_Vac.AllowUserToAddRows = False
        Me.Data_Vac.AllowUserToDeleteRows = False
        Me.Data_Vac.AllowUserToResizeColumns = False
        Me.Data_Vac.AllowUserToResizeRows = False
        Me.Data_Vac.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.Data_Vac.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.Data_Vac.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Data_Vac.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewCheckBoxColumn1, Me.DataGridViewTextBoxColumn1})
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Data_Vac.DefaultCellStyle = DataGridViewCellStyle1
        Me.Data_Vac.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Data_Vac.Location = New System.Drawing.Point(3, 16)
        Me.Data_Vac.Name = "Data_Vac"
        Me.Data_Vac.RowHeadersVisible = False
        Me.Data_Vac.Size = New System.Drawing.Size(154, 150)
        Me.Data_Vac.TabIndex = 5
        '
        'DataGridViewCheckBoxColumn1
        '
        Me.DataGridViewCheckBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.DataGridViewCheckBoxColumn1.HeaderText = "Used"
        Me.DataGridViewCheckBoxColumn1.Name = "DataGridViewCheckBoxColumn1"
        Me.DataGridViewCheckBoxColumn1.Width = 38
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.DataGridViewTextBoxColumn1.HeaderText = "Vac Number"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        Me.DataGridViewTextBoxColumn1.Width = 91
        '
        'Box_TC
        '
        Me.Box_TC.AutoSize = True
        Me.Box_TC.Controls.Add(Me.Data_TC)
        Me.Box_TC.Location = New System.Drawing.Point(391, 19)
        Me.Box_TC.Name = "Box_TC"
        Me.Box_TC.Size = New System.Drawing.Size(160, 169)
        Me.Box_TC.TabIndex = 8
        Me.Box_TC.TabStop = False
        Me.Box_TC.Text = "Thermocouples"
        Me.Box_TC.Visible = False
        '
        'Data_TC
        '
        Me.Data_TC.AllowUserToAddRows = False
        Me.Data_TC.AllowUserToDeleteRows = False
        Me.Data_TC.AllowUserToResizeColumns = False
        Me.Data_TC.AllowUserToResizeRows = False
        Me.Data_TC.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.Data_TC.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.Data_TC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Data_TC.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Used, Me.Modifier})
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Data_TC.DefaultCellStyle = DataGridViewCellStyle2
        Me.Data_TC.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Data_TC.Location = New System.Drawing.Point(3, 16)
        Me.Data_TC.Name = "Data_TC"
        Me.Data_TC.RowHeadersVisible = False
        Me.Data_TC.Size = New System.Drawing.Size(154, 150)
        Me.Data_TC.TabIndex = 5
        '
        'Used
        '
        Me.Used.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.Used.HeaderText = "Used"
        Me.Used.Name = "Used"
        Me.Used.Width = 38
        '
        'Modifier
        '
        Me.Modifier.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.Modifier.HeaderText = "TC Number"
        Me.Modifier.Name = "Modifier"
        Me.Modifier.ReadOnly = True
        Me.Modifier.Width = 86
        '
        'Box_JobNumber
        '
        Me.Box_JobNumber.Controls.Add(Me.Txt_JobNumber)
        Me.Box_JobNumber.Location = New System.Drawing.Point(6, 19)
        Me.Box_JobNumber.Name = "Box_JobNumber"
        Me.Box_JobNumber.Size = New System.Drawing.Size(114, 44)
        Me.Box_JobNumber.TabIndex = 3
        Me.Box_JobNumber.TabStop = False
        Me.Box_JobNumber.Text = "Job Number"
        '
        'Txt_JobNumber
        '
        Me.Txt_JobNumber.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_JobNumber.Location = New System.Drawing.Point(3, 16)
        Me.Txt_JobNumber.Name = "Txt_JobNumber"
        Me.Txt_JobNumber.Size = New System.Drawing.Size(108, 20)
        Me.Txt_JobNumber.TabIndex = 4
        Me.Txt_JobNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Box_ProgramNumber
        '
        Me.Box_ProgramNumber.Controls.Add(Me.Txt_ProgramNumber)
        Me.Box_ProgramNumber.Location = New System.Drawing.Point(126, 19)
        Me.Box_ProgramNumber.Name = "Box_ProgramNumber"
        Me.Box_ProgramNumber.Size = New System.Drawing.Size(114, 44)
        Me.Box_ProgramNumber.TabIndex = 4
        Me.Box_ProgramNumber.TabStop = False
        Me.Box_ProgramNumber.Text = "Program Number"
        '
        'Txt_ProgramNumber
        '
        Me.Txt_ProgramNumber.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_ProgramNumber.Location = New System.Drawing.Point(3, 16)
        Me.Txt_ProgramNumber.Name = "Txt_ProgramNumber"
        Me.Txt_ProgramNumber.Size = New System.Drawing.Size(108, 20)
        Me.Txt_ProgramNumber.TabIndex = 5
        Me.Txt_ProgramNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Box_DataRecorder
        '
        Me.Box_DataRecorder.Controls.Add(Me.Txt_DataRecorder)
        Me.Box_DataRecorder.Location = New System.Drawing.Point(246, 19)
        Me.Box_DataRecorder.Name = "Box_DataRecorder"
        Me.Box_DataRecorder.Size = New System.Drawing.Size(142, 44)
        Me.Box_DataRecorder.TabIndex = 5
        Me.Box_DataRecorder.TabStop = False
        Me.Box_DataRecorder.Text = "Data Recorder ID (S#)"
        Me.Box_DataRecorder.Visible = False
        '
        'Txt_DataRecorder
        '
        Me.Txt_DataRecorder.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_DataRecorder.Location = New System.Drawing.Point(3, 16)
        Me.Txt_DataRecorder.Name = "Txt_DataRecorder"
        Me.Txt_DataRecorder.Size = New System.Drawing.Size(136, 20)
        Me.Txt_DataRecorder.TabIndex = 6
        Me.Txt_DataRecorder.Text = "S"
        Me.Txt_DataRecorder.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Box_RunLine
        '
        Me.Box_RunLine.ColumnCount = 2
        Me.Box_RunLine.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 77.04918!))
        Me.Box_RunLine.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 22.95082!))
        Me.Box_RunLine.Controls.Add(Me.Btn_Run, 1, 0)
        Me.Box_RunLine.Controls.Add(Me.Box_Technician, 0, 0)
        Me.Box_RunLine.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Box_RunLine.Enabled = False
        Me.Box_RunLine.Location = New System.Drawing.Point(23, 413)
        Me.Box_RunLine.Name = "Box_RunLine"
        Me.Box_RunLine.RowCount = 1
        Me.Box_RunLine.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.Box_RunLine.Size = New System.Drawing.Size(760, 54)
        Me.Box_RunLine.TabIndex = 14
        '
        'Btn_Run
        '
        Me.Btn_Run.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Btn_Run.FlatAppearance.BorderSize = 2
        Me.Btn_Run.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Btn_Run.Location = New System.Drawing.Point(588, 3)
        Me.Btn_Run.Name = "Btn_Run"
        Me.Btn_Run.Size = New System.Drawing.Size(169, 48)
        Me.Btn_Run.TabIndex = 12
        Me.Btn_Run.Text = "Run Output"
        Me.Btn_Run.UseVisualStyleBackColor = True
        '
        'Box_Technician
        '
        Me.Box_Technician.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Box_Technician.Controls.Add(Me.Txt_Technician)
        Me.Box_Technician.Location = New System.Drawing.Point(421, 5)
        Me.Box_Technician.Name = "Box_Technician"
        Me.Box_Technician.Size = New System.Drawing.Size(161, 44)
        Me.Box_Technician.TabIndex = 10
        Me.Box_Technician.TabStop = False
        Me.Box_Technician.Text = "Completed By:"
        '
        'Txt_Technician
        '
        Me.Txt_Technician.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_Technician.Location = New System.Drawing.Point(3, 16)
        Me.Txt_Technician.Name = "Txt_Technician"
        Me.Txt_Technician.Size = New System.Drawing.Size(155, 20)
        Me.Txt_Technician.TabIndex = 11
        Me.Txt_Technician.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TableLayoutPanel6
        '
        Me.TableLayoutPanel6.ColumnCount = 4
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 300.0!))
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel6.Controls.Add(Me.Box_CureProfChoice, 0, 0)
        Me.TableLayoutPanel6.Controls.Add(Me.PictureBox1, 2, 0)
        Me.TableLayoutPanel6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel6.Location = New System.Drawing.Point(23, 3)
        Me.TableLayoutPanel6.Name = "TableLayoutPanel6"
        Me.TableLayoutPanel6.RowCount = 1
        Me.TableLayoutPanel6.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel6.Size = New System.Drawing.Size(760, 144)
        Me.TableLayoutPanel6.TabIndex = 15
        '
        'Box_CureProfChoice
        '
        Me.Box_CureProfChoice.Controls.Add(Me.TableLayoutPanel3)
        Me.Box_CureProfChoice.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Box_CureProfChoice.Location = New System.Drawing.Point(3, 3)
        Me.Box_CureProfChoice.Name = "Box_CureProfChoice"
        Me.Box_CureProfChoice.Size = New System.Drawing.Size(414, 138)
        Me.Box_CureProfChoice.TabIndex = 11
        Me.Box_CureProfChoice.TabStop = False
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.ColumnCount = 1
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel3.Controls.Add(Me.Box_CureProfile, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.TableLayoutPanel4, 0, 1)
        Me.TableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(3, 16)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 2
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(408, 119)
        Me.TableLayoutPanel3.TabIndex = 12
        '
        'Box_CureProfile
        '
        Me.Box_CureProfile.Controls.Add(Me.Combo_CureProfile)
        Me.Box_CureProfile.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Box_CureProfile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.5!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Box_CureProfile.Location = New System.Drawing.Point(3, 3)
        Me.Box_CureProfile.Name = "Box_CureProfile"
        Me.Box_CureProfile.Size = New System.Drawing.Size(402, 53)
        Me.Box_CureProfile.TabIndex = 7
        Me.Box_CureProfile.TabStop = False
        Me.Box_CureProfile.Text = "Cure Profile"
        '
        'Combo_CureProfile
        '
        Me.Combo_CureProfile.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Combo_CureProfile.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_CureProfile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_CureProfile.FormattingEnabled = True
        Me.Combo_CureProfile.Location = New System.Drawing.Point(6, 16)
        Me.Combo_CureProfile.Name = "Combo_CureProfile"
        Me.Combo_CureProfile.Size = New System.Drawing.Size(390, 21)
        Me.Combo_CureProfile.TabIndex = 1
        '
        'TableLayoutPanel4
        '
        Me.TableLayoutPanel4.ColumnCount = 2
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 60.0!))
        Me.TableLayoutPanel4.Controls.Add(Me.Box_DocRev, 1, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.Box_CureDoc, 0, 0)
        Me.TableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel4.Location = New System.Drawing.Point(3, 62)
        Me.TableLayoutPanel4.Name = "TableLayoutPanel4"
        Me.TableLayoutPanel4.RowCount = 1
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel4.Size = New System.Drawing.Size(402, 54)
        Me.TableLayoutPanel4.TabIndex = 8
        '
        'Box_DocRev
        '
        Me.Box_DocRev.Controls.Add(Me.Txt_DocRev)
        Me.Box_DocRev.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Box_DocRev.Location = New System.Drawing.Point(345, 3)
        Me.Box_DocRev.Name = "Box_DocRev"
        Me.Box_DocRev.Size = New System.Drawing.Size(54, 48)
        Me.Box_DocRev.TabIndex = 3
        Me.Box_DocRev.TabStop = False
        Me.Box_DocRev.Text = "Rev"
        '
        'Txt_DocRev
        '
        Me.Txt_DocRev.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_DocRev.Enabled = False
        Me.Txt_DocRev.Location = New System.Drawing.Point(3, 16)
        Me.Txt_DocRev.Name = "Txt_DocRev"
        Me.Txt_DocRev.Size = New System.Drawing.Size(48, 20)
        Me.Txt_DocRev.TabIndex = 1
        Me.Txt_DocRev.TabStop = False
        Me.Txt_DocRev.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Box_CureDoc
        '
        Me.Box_CureDoc.Controls.Add(Me.Txt_CureDoc)
        Me.Box_CureDoc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Box_CureDoc.Location = New System.Drawing.Point(3, 3)
        Me.Box_CureDoc.Name = "Box_CureDoc"
        Me.Box_CureDoc.Size = New System.Drawing.Size(336, 48)
        Me.Box_CureDoc.TabIndex = 3
        Me.Box_CureDoc.TabStop = False
        Me.Box_CureDoc.Text = "Document"
        '
        'Txt_CureDoc
        '
        Me.Txt_CureDoc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_CureDoc.Enabled = False
        Me.Txt_CureDoc.Location = New System.Drawing.Point(3, 16)
        Me.Txt_CureDoc.Name = "Txt_CureDoc"
        Me.Txt_CureDoc.Size = New System.Drawing.Size(330, 20)
        Me.Txt_CureDoc.TabIndex = 1
        Me.Txt_CureDoc.TabStop = False
        Me.Txt_CureDoc.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.CureReportBuilder.My.Resources.Resources.Systima_Composites_2018
        Me.PictureBox1.Location = New System.Drawing.Point(443, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(294, 138)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 12
        Me.PictureBox1.TabStop = False
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Box_TemplatePath)
        Me.TabPage2.Controls.Add(Me.Box_CureProfiles)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(792, 476)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Settings"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Box_TemplatePath
        '
        Me.Box_TemplatePath.AutoSize = True
        Me.Box_TemplatePath.Controls.Add(Me.TableLayoutPanel5)
        Me.Box_TemplatePath.Location = New System.Drawing.Point(3, 116)
        Me.Box_TemplatePath.Name = "Box_TemplatePath"
        Me.Box_TemplatePath.Size = New System.Drawing.Size(646, 57)
        Me.Box_TemplatePath.TabIndex = 2
        Me.Box_TemplatePath.TabStop = False
        Me.Box_TemplatePath.Text = "Cure Report Template"
        '
        'TableLayoutPanel5
        '
        Me.TableLayoutPanel5.ColumnCount = 2
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 102.0!))
        Me.TableLayoutPanel5.Controls.Add(Me.Txt_TemplatePath, 0, 0)
        Me.TableLayoutPanel5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel5.Location = New System.Drawing.Point(3, 16)
        Me.TableLayoutPanel5.Name = "TableLayoutPanel5"
        Me.TableLayoutPanel5.RowCount = 1
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel5.Size = New System.Drawing.Size(640, 38)
        Me.TableLayoutPanel5.TabIndex = 2
        '
        'Txt_TemplatePath
        '
        Me.Txt_TemplatePath.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_TemplatePath.Location = New System.Drawing.Point(3, 9)
        Me.Txt_TemplatePath.Name = "Txt_TemplatePath"
        Me.Txt_TemplatePath.Size = New System.Drawing.Size(532, 20)
        Me.Txt_TemplatePath.TabIndex = 0
        Me.Txt_TemplatePath.Text = "S:\Engineering\Functional Groups\Composites\Macros\CureReportFiles\Cure Profiles\" &
    "Cure Report_Template.xlsx"
        '
        'Box_CureProfiles
        '
        Me.Box_CureProfiles.AutoSize = True
        Me.Box_CureProfiles.Controls.Add(Me.TableLayoutPanel2)
        Me.Box_CureProfiles.Location = New System.Drawing.Point(6, 6)
        Me.Box_CureProfiles.Name = "Box_CureProfiles"
        Me.Box_CureProfiles.Size = New System.Drawing.Size(646, 57)
        Me.Box_CureProfiles.TabIndex = 1
        Me.Box_CureProfiles.TabStop = False
        Me.Box_CureProfiles.Text = "Cure Profiles"
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 2
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 102.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.Txt_CureProfilesPath, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Btn_LoadProfileFiles, 1, 0)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(3, 16)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(640, 38)
        Me.TableLayoutPanel2.TabIndex = 2
        '
        'Txt_CureProfilesPath
        '
        Me.Txt_CureProfilesPath.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_CureProfilesPath.Location = New System.Drawing.Point(3, 9)
        Me.Txt_CureProfilesPath.Name = "Txt_CureProfilesPath"
        Me.Txt_CureProfilesPath.Size = New System.Drawing.Size(532, 20)
        Me.Txt_CureProfilesPath.TabIndex = 0
        Me.Txt_CureProfilesPath.Text = "S:\Engineering\Functional Groups\Composites\Macros\CureReportFiles\Cure Profiles"
        '
        'Btn_LoadProfileFiles
        '
        Me.Btn_LoadProfileFiles.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Btn_LoadProfileFiles.Location = New System.Drawing.Point(541, 3)
        Me.Btn_LoadProfileFiles.Name = "Btn_LoadProfileFiles"
        Me.Btn_LoadProfileFiles.Size = New System.Drawing.Size(96, 32)
        Me.Btn_LoadProfileFiles.TabIndex = 1
        Me.Btn_LoadProfileFiles.Text = "Load Profiles"
        Me.Btn_LoadProfileFiles.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.GroupBox4)
        Me.TabPage3.Controls.Add(Me.Box_CheckEdit)
        Me.TabPage3.Controls.Add(Me.Box_CureNameEdit)
        Me.TabPage3.Controls.Add(Me.GroupBox2)
        Me.TabPage3.Controls.Add(Me.GroupBox3)
        Me.TabPage3.Controls.Add(Me.GroupBox1)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(792, 476)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "CureProfileEdit"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Box_vacStepEdit)
        Me.GroupBox4.Controls.Add(Me.Box_pressureStepEdit)
        Me.GroupBox4.Controls.Add(Me.Box_tempStepEdit)
        Me.GroupBox4.Controls.Add(Me.Box_StepNameEdit)
        Me.GroupBox4.Location = New System.Drawing.Point(8, 113)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(776, 360)
        Me.GroupBox4.TabIndex = 18
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "GroupBox4"
        '
        'Box_vacStepEdit
        '
        Me.Box_vacStepEdit.Controls.Add(Me.Check_vacMinEdit)
        Me.Box_vacStepEdit.Controls.Add(Me.Check_vacMaxEdit)
        Me.Box_vacStepEdit.Controls.Add(Me.Check_vacStepEdit)
        Me.Box_vacStepEdit.Controls.Add(Me.Box_vacNegTolEdit)
        Me.Box_vacStepEdit.Controls.Add(Me.Box_vacPosTolEdit)
        Me.Box_vacStepEdit.Controls.Add(Me.Box_vacSetEdit)
        Me.Box_vacStepEdit.Location = New System.Drawing.Point(278, 121)
        Me.Box_vacStepEdit.Name = "Box_vacStepEdit"
        Me.Box_vacStepEdit.Size = New System.Drawing.Size(225, 213)
        Me.Box_vacStepEdit.TabIndex = 18
        Me.Box_vacStepEdit.TabStop = False
        Me.Box_vacStepEdit.Text = "Vacuum Step"
        '
        'Check_vacMinEdit
        '
        Me.Check_vacMinEdit.AutoSize = True
        Me.Check_vacMinEdit.Location = New System.Drawing.Point(7, 68)
        Me.Check_vacMinEdit.Name = "Check_vacMinEdit"
        Me.Check_vacMinEdit.Size = New System.Drawing.Size(67, 17)
        Me.Check_vacMinEdit.TabIndex = 24
        Me.Check_vacMinEdit.Text = "Min Only"
        Me.Check_vacMinEdit.UseVisualStyleBackColor = True
        '
        'Check_vacMaxEdit
        '
        Me.Check_vacMaxEdit.AutoSize = True
        Me.Check_vacMaxEdit.Location = New System.Drawing.Point(7, 45)
        Me.Check_vacMaxEdit.Name = "Check_vacMaxEdit"
        Me.Check_vacMaxEdit.Size = New System.Drawing.Size(70, 17)
        Me.Check_vacMaxEdit.TabIndex = 23
        Me.Check_vacMaxEdit.Text = "Max Only"
        Me.Check_vacMaxEdit.UseVisualStyleBackColor = True
        '
        'Check_vacStepEdit
        '
        Me.Check_vacStepEdit.AutoSize = True
        Me.Check_vacStepEdit.Checked = True
        Me.Check_vacStepEdit.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Check_vacStepEdit.Location = New System.Drawing.Point(7, 20)
        Me.Check_vacStepEdit.Name = "Check_vacStepEdit"
        Me.Check_vacStepEdit.Size = New System.Drawing.Size(57, 17)
        Me.Check_vacStepEdit.TabIndex = 22
        Me.Check_vacStepEdit.Text = "Check"
        Me.Check_vacStepEdit.UseVisualStyleBackColor = True
        '
        'Box_vacNegTolEdit
        '
        Me.Box_vacNegTolEdit.Controls.Add(Me.Txt_vacNegTolEdit)
        Me.Box_vacNegTolEdit.Location = New System.Drawing.Point(108, 110)
        Me.Box_vacNegTolEdit.Name = "Box_vacNegTolEdit"
        Me.Box_vacNegTolEdit.Size = New System.Drawing.Size(129, 43)
        Me.Box_vacNegTolEdit.TabIndex = 19
        Me.Box_vacNegTolEdit.TabStop = False
        Me.Box_vacNegTolEdit.Text = "NegTol (inHg)"
        '
        'Txt_vacNegTolEdit
        '
        Me.Txt_vacNegTolEdit.Location = New System.Drawing.Point(6, 19)
        Me.Txt_vacNegTolEdit.Name = "Txt_vacNegTolEdit"
        Me.Txt_vacNegTolEdit.Size = New System.Drawing.Size(100, 20)
        Me.Txt_vacNegTolEdit.TabIndex = 0
        Me.Txt_vacNegTolEdit.Text = "-50"
        '
        'Box_vacPosTolEdit
        '
        Me.Box_vacPosTolEdit.Controls.Add(Me.Txt_vacPosTolEdit)
        Me.Box_vacPosTolEdit.Location = New System.Drawing.Point(108, 65)
        Me.Box_vacPosTolEdit.Name = "Box_vacPosTolEdit"
        Me.Box_vacPosTolEdit.Size = New System.Drawing.Size(129, 43)
        Me.Box_vacPosTolEdit.TabIndex = 18
        Me.Box_vacPosTolEdit.TabStop = False
        Me.Box_vacPosTolEdit.Text = "PosTol (inHg)"
        '
        'Txt_vacPosTolEdit
        '
        Me.Txt_vacPosTolEdit.Location = New System.Drawing.Point(6, 19)
        Me.Txt_vacPosTolEdit.Name = "Txt_vacPosTolEdit"
        Me.Txt_vacPosTolEdit.Size = New System.Drawing.Size(100, 20)
        Me.Txt_vacPosTolEdit.TabIndex = 0
        Me.Txt_vacPosTolEdit.Text = "50"
        '
        'Box_vacSetEdit
        '
        Me.Box_vacSetEdit.Controls.Add(Me.Txt_vacSetEdit)
        Me.Box_vacSetEdit.Location = New System.Drawing.Point(97, 20)
        Me.Box_vacSetEdit.Name = "Box_vacSetEdit"
        Me.Box_vacSetEdit.Size = New System.Drawing.Size(129, 43)
        Me.Box_vacSetEdit.TabIndex = 17
        Me.Box_vacSetEdit.TabStop = False
        Me.Box_vacSetEdit.Text = "Set Point (inHg)"
        '
        'Txt_vacSetEdit
        '
        Me.Txt_vacSetEdit.Location = New System.Drawing.Point(6, 19)
        Me.Txt_vacSetEdit.Name = "Txt_vacSetEdit"
        Me.Txt_vacSetEdit.Size = New System.Drawing.Size(100, 20)
        Me.Txt_vacSetEdit.TabIndex = 0
        Me.Txt_vacSetEdit.Text = "500"
        '
        'Box_pressureStepEdit
        '
        Me.Box_pressureStepEdit.Controls.Add(Me.Check_pressureMinEdit)
        Me.Box_pressureStepEdit.Controls.Add(Me.Check_pressureMaxEdit)
        Me.Box_pressureStepEdit.Controls.Add(Me.Check_pressureStepEdit)
        Me.Box_pressureStepEdit.Controls.Add(Me.Box_pressureRampNegTolEdit)
        Me.Box_pressureStepEdit.Controls.Add(Me.Box_pressureRampPosTolEdit)
        Me.Box_pressureStepEdit.Controls.Add(Me.Box_pressureRampEdit)
        Me.Box_pressureStepEdit.Controls.Add(Me.Box_pressureNegTolEdit)
        Me.Box_pressureStepEdit.Controls.Add(Me.Box_pressurePosTolEdit)
        Me.Box_pressureStepEdit.Controls.Add(Me.Box_pressureSetEdit)
        Me.Box_pressureStepEdit.Location = New System.Drawing.Point(509, 29)
        Me.Box_pressureStepEdit.Name = "Box_pressureStepEdit"
        Me.Box_pressureStepEdit.Size = New System.Drawing.Size(254, 305)
        Me.Box_pressureStepEdit.TabIndex = 17
        Me.Box_pressureStepEdit.TabStop = False
        Me.Box_pressureStepEdit.Text = "Pressure Step"
        '
        'Check_pressureMinEdit
        '
        Me.Check_pressureMinEdit.AutoSize = True
        Me.Check_pressureMinEdit.Location = New System.Drawing.Point(7, 68)
        Me.Check_pressureMinEdit.Name = "Check_pressureMinEdit"
        Me.Check_pressureMinEdit.Size = New System.Drawing.Size(67, 17)
        Me.Check_pressureMinEdit.TabIndex = 24
        Me.Check_pressureMinEdit.Text = "Min Only"
        Me.Check_pressureMinEdit.UseVisualStyleBackColor = True
        '
        'Check_pressureMaxEdit
        '
        Me.Check_pressureMaxEdit.AutoSize = True
        Me.Check_pressureMaxEdit.Location = New System.Drawing.Point(7, 45)
        Me.Check_pressureMaxEdit.Name = "Check_pressureMaxEdit"
        Me.Check_pressureMaxEdit.Size = New System.Drawing.Size(70, 17)
        Me.Check_pressureMaxEdit.TabIndex = 23
        Me.Check_pressureMaxEdit.Text = "Max Only"
        Me.Check_pressureMaxEdit.UseVisualStyleBackColor = True
        '
        'Check_pressureStepEdit
        '
        Me.Check_pressureStepEdit.AutoSize = True
        Me.Check_pressureStepEdit.Checked = True
        Me.Check_pressureStepEdit.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Check_pressureStepEdit.Location = New System.Drawing.Point(7, 20)
        Me.Check_pressureStepEdit.Name = "Check_pressureStepEdit"
        Me.Check_pressureStepEdit.Size = New System.Drawing.Size(57, 17)
        Me.Check_pressureStepEdit.TabIndex = 22
        Me.Check_pressureStepEdit.Text = "Check"
        Me.Check_pressureStepEdit.UseVisualStyleBackColor = True
        '
        'Box_pressureRampNegTolEdit
        '
        Me.Box_pressureRampNegTolEdit.Controls.Add(Me.Txt_pressureRampNegTolEdit)
        Me.Box_pressureRampNegTolEdit.Location = New System.Drawing.Point(108, 255)
        Me.Box_pressureRampNegTolEdit.Name = "Box_pressureRampNegTolEdit"
        Me.Box_pressureRampNegTolEdit.Size = New System.Drawing.Size(141, 43)
        Me.Box_pressureRampNegTolEdit.TabIndex = 21
        Me.Box_pressureRampNegTolEdit.TabStop = False
        Me.Box_pressureRampNegTolEdit.Text = "Ramp NegTol (psi/min)"
        '
        'Txt_pressureRampNegTolEdit
        '
        Me.Txt_pressureRampNegTolEdit.Location = New System.Drawing.Point(6, 19)
        Me.Txt_pressureRampNegTolEdit.Name = "Txt_pressureRampNegTolEdit"
        Me.Txt_pressureRampNegTolEdit.Size = New System.Drawing.Size(100, 20)
        Me.Txt_pressureRampNegTolEdit.TabIndex = 0
        Me.Txt_pressureRampNegTolEdit.Text = "0"
        '
        'Box_pressureRampPosTolEdit
        '
        Me.Box_pressureRampPosTolEdit.Controls.Add(Me.Txt_pressureRampPosTolEdit)
        Me.Box_pressureRampPosTolEdit.Location = New System.Drawing.Point(108, 206)
        Me.Box_pressureRampPosTolEdit.Name = "Box_pressureRampPosTolEdit"
        Me.Box_pressureRampPosTolEdit.Size = New System.Drawing.Size(141, 43)
        Me.Box_pressureRampPosTolEdit.TabIndex = 19
        Me.Box_pressureRampPosTolEdit.TabStop = False
        Me.Box_pressureRampPosTolEdit.Text = "Ramp PosTol (psi/min)"
        '
        'Txt_pressureRampPosTolEdit
        '
        Me.Txt_pressureRampPosTolEdit.Location = New System.Drawing.Point(6, 19)
        Me.Txt_pressureRampPosTolEdit.Name = "Txt_pressureRampPosTolEdit"
        Me.Txt_pressureRampPosTolEdit.Size = New System.Drawing.Size(100, 20)
        Me.Txt_pressureRampPosTolEdit.TabIndex = 0
        Me.Txt_pressureRampPosTolEdit.Text = "0"
        '
        'Box_pressureRampEdit
        '
        Me.Box_pressureRampEdit.Controls.Add(Me.Txt_pressureRampEdit)
        Me.Box_pressureRampEdit.Location = New System.Drawing.Point(97, 161)
        Me.Box_pressureRampEdit.Name = "Box_pressureRampEdit"
        Me.Box_pressureRampEdit.Size = New System.Drawing.Size(129, 43)
        Me.Box_pressureRampEdit.TabIndex = 20
        Me.Box_pressureRampEdit.TabStop = False
        Me.Box_pressureRampEdit.Text = "Ramp Rate (psi/min)"
        '
        'Txt_pressureRampEdit
        '
        Me.Txt_pressureRampEdit.Location = New System.Drawing.Point(6, 19)
        Me.Txt_pressureRampEdit.Name = "Txt_pressureRampEdit"
        Me.Txt_pressureRampEdit.Size = New System.Drawing.Size(100, 20)
        Me.Txt_pressureRampEdit.TabIndex = 0
        Me.Txt_pressureRampEdit.Text = "0"
        '
        'Box_pressureNegTolEdit
        '
        Me.Box_pressureNegTolEdit.Controls.Add(Me.Txt_pressureNegTolEdit)
        Me.Box_pressureNegTolEdit.Location = New System.Drawing.Point(108, 110)
        Me.Box_pressureNegTolEdit.Name = "Box_pressureNegTolEdit"
        Me.Box_pressureNegTolEdit.Size = New System.Drawing.Size(129, 43)
        Me.Box_pressureNegTolEdit.TabIndex = 19
        Me.Box_pressureNegTolEdit.TabStop = False
        Me.Box_pressureNegTolEdit.Text = "NegTol (psi)"
        '
        'Txt_pressureNegTolEdit
        '
        Me.Txt_pressureNegTolEdit.Location = New System.Drawing.Point(6, 19)
        Me.Txt_pressureNegTolEdit.Name = "Txt_pressureNegTolEdit"
        Me.Txt_pressureNegTolEdit.Size = New System.Drawing.Size(100, 20)
        Me.Txt_pressureNegTolEdit.TabIndex = 0
        Me.Txt_pressureNegTolEdit.Text = "-50"
        '
        'Box_pressurePosTolEdit
        '
        Me.Box_pressurePosTolEdit.Controls.Add(Me.Txt_pressurePosTolEdit)
        Me.Box_pressurePosTolEdit.Location = New System.Drawing.Point(108, 65)
        Me.Box_pressurePosTolEdit.Name = "Box_pressurePosTolEdit"
        Me.Box_pressurePosTolEdit.Size = New System.Drawing.Size(129, 43)
        Me.Box_pressurePosTolEdit.TabIndex = 18
        Me.Box_pressurePosTolEdit.TabStop = False
        Me.Box_pressurePosTolEdit.Text = "PosTol (psi)"
        '
        'Txt_pressurePosTolEdit
        '
        Me.Txt_pressurePosTolEdit.Location = New System.Drawing.Point(6, 19)
        Me.Txt_pressurePosTolEdit.Name = "Txt_pressurePosTolEdit"
        Me.Txt_pressurePosTolEdit.Size = New System.Drawing.Size(100, 20)
        Me.Txt_pressurePosTolEdit.TabIndex = 0
        Me.Txt_pressurePosTolEdit.Text = "50"
        '
        'Box_pressureSetEdit
        '
        Me.Box_pressureSetEdit.Controls.Add(Me.Txt_pressureSetEdit)
        Me.Box_pressureSetEdit.Location = New System.Drawing.Point(97, 20)
        Me.Box_pressureSetEdit.Name = "Box_pressureSetEdit"
        Me.Box_pressureSetEdit.Size = New System.Drawing.Size(129, 43)
        Me.Box_pressureSetEdit.TabIndex = 17
        Me.Box_pressureSetEdit.TabStop = False
        Me.Box_pressureSetEdit.Text = "Set Point (psi)"
        '
        'Txt_pressureSetEdit
        '
        Me.Txt_pressureSetEdit.Location = New System.Drawing.Point(6, 19)
        Me.Txt_pressureSetEdit.Name = "Txt_pressureSetEdit"
        Me.Txt_pressureSetEdit.Size = New System.Drawing.Size(100, 20)
        Me.Txt_pressureSetEdit.TabIndex = 0
        Me.Txt_pressureSetEdit.Text = "500"
        '
        'Box_tempStepEdit
        '
        Me.Box_tempStepEdit.Controls.Add(Me.Check_tempMinEdit)
        Me.Box_tempStepEdit.Controls.Add(Me.Check_tempMaxEdit)
        Me.Box_tempStepEdit.Controls.Add(Me.Check_tempStepEdit)
        Me.Box_tempStepEdit.Controls.Add(Me.Box_TempRampNegTolEdit)
        Me.Box_tempStepEdit.Controls.Add(Me.Box_TempRampPosTolEdit)
        Me.Box_tempStepEdit.Controls.Add(Me.Box_TempRampEdit)
        Me.Box_tempStepEdit.Controls.Add(Me.Box_TempNegTolEdit)
        Me.Box_tempStepEdit.Controls.Add(Me.Box_TempPosTolEdit)
        Me.Box_tempStepEdit.Controls.Add(Me.Box_TempSetEdit)
        Me.Box_tempStepEdit.Location = New System.Drawing.Point(6, 29)
        Me.Box_tempStepEdit.Name = "Box_tempStepEdit"
        Me.Box_tempStepEdit.Size = New System.Drawing.Size(254, 305)
        Me.Box_tempStepEdit.TabIndex = 16
        Me.Box_tempStepEdit.TabStop = False
        Me.Box_tempStepEdit.Text = "Temperature Step"
        '
        'Check_tempMinEdit
        '
        Me.Check_tempMinEdit.AutoSize = True
        Me.Check_tempMinEdit.Location = New System.Drawing.Point(7, 68)
        Me.Check_tempMinEdit.Name = "Check_tempMinEdit"
        Me.Check_tempMinEdit.Size = New System.Drawing.Size(67, 17)
        Me.Check_tempMinEdit.TabIndex = 24
        Me.Check_tempMinEdit.Text = "Min Only"
        Me.Check_tempMinEdit.UseVisualStyleBackColor = True
        '
        'Check_tempMaxEdit
        '
        Me.Check_tempMaxEdit.AutoSize = True
        Me.Check_tempMaxEdit.Location = New System.Drawing.Point(7, 45)
        Me.Check_tempMaxEdit.Name = "Check_tempMaxEdit"
        Me.Check_tempMaxEdit.Size = New System.Drawing.Size(70, 17)
        Me.Check_tempMaxEdit.TabIndex = 23
        Me.Check_tempMaxEdit.Text = "Max Only"
        Me.Check_tempMaxEdit.UseVisualStyleBackColor = True
        '
        'Check_tempStepEdit
        '
        Me.Check_tempStepEdit.AutoSize = True
        Me.Check_tempStepEdit.Checked = True
        Me.Check_tempStepEdit.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Check_tempStepEdit.Location = New System.Drawing.Point(7, 20)
        Me.Check_tempStepEdit.Name = "Check_tempStepEdit"
        Me.Check_tempStepEdit.Size = New System.Drawing.Size(87, 17)
        Me.Check_tempStepEdit.TabIndex = 22
        Me.Check_tempStepEdit.Text = "Check Temp"
        Me.Check_tempStepEdit.UseVisualStyleBackColor = True
        '
        'Box_TempRampNegTolEdit
        '
        Me.Box_TempRampNegTolEdit.Controls.Add(Me.Txt_TempRampNegTolEdit)
        Me.Box_TempRampNegTolEdit.Location = New System.Drawing.Point(120, 255)
        Me.Box_TempRampNegTolEdit.Name = "Box_TempRampNegTolEdit"
        Me.Box_TempRampNegTolEdit.Size = New System.Drawing.Size(129, 43)
        Me.Box_TempRampNegTolEdit.TabIndex = 21
        Me.Box_TempRampNegTolEdit.TabStop = False
        Me.Box_TempRampNegTolEdit.Text = "Ramp NegTol (°F/min)"
        '
        'Txt_TempRampNegTolEdit
        '
        Me.Txt_TempRampNegTolEdit.Location = New System.Drawing.Point(6, 19)
        Me.Txt_TempRampNegTolEdit.Name = "Txt_TempRampNegTolEdit"
        Me.Txt_TempRampNegTolEdit.Size = New System.Drawing.Size(100, 20)
        Me.Txt_TempRampNegTolEdit.TabIndex = 0
        Me.Txt_TempRampNegTolEdit.Text = "0"
        '
        'Box_TempRampPosTolEdit
        '
        Me.Box_TempRampPosTolEdit.Controls.Add(Me.Txt_TempRampPosTolEdit)
        Me.Box_TempRampPosTolEdit.Location = New System.Drawing.Point(120, 206)
        Me.Box_TempRampPosTolEdit.Name = "Box_TempRampPosTolEdit"
        Me.Box_TempRampPosTolEdit.Size = New System.Drawing.Size(129, 43)
        Me.Box_TempRampPosTolEdit.TabIndex = 19
        Me.Box_TempRampPosTolEdit.TabStop = False
        Me.Box_TempRampPosTolEdit.Text = "Ramp PosTol (°F/min)"
        '
        'Txt_TempRampPosTolEdit
        '
        Me.Txt_TempRampPosTolEdit.Location = New System.Drawing.Point(6, 19)
        Me.Txt_TempRampPosTolEdit.Name = "Txt_TempRampPosTolEdit"
        Me.Txt_TempRampPosTolEdit.Size = New System.Drawing.Size(100, 20)
        Me.Txt_TempRampPosTolEdit.TabIndex = 0
        Me.Txt_TempRampPosTolEdit.Text = "0"
        '
        'Box_TempRampEdit
        '
        Me.Box_TempRampEdit.Controls.Add(Me.Txt_TempRampEdit)
        Me.Box_TempRampEdit.Location = New System.Drawing.Point(97, 161)
        Me.Box_TempRampEdit.Name = "Box_TempRampEdit"
        Me.Box_TempRampEdit.Size = New System.Drawing.Size(129, 43)
        Me.Box_TempRampEdit.TabIndex = 20
        Me.Box_TempRampEdit.TabStop = False
        Me.Box_TempRampEdit.Text = "Ramp Rate (°F/min)"
        '
        'Txt_TempRampEdit
        '
        Me.Txt_TempRampEdit.Location = New System.Drawing.Point(6, 19)
        Me.Txt_TempRampEdit.Name = "Txt_TempRampEdit"
        Me.Txt_TempRampEdit.Size = New System.Drawing.Size(100, 20)
        Me.Txt_TempRampEdit.TabIndex = 0
        Me.Txt_TempRampEdit.Text = "0"
        '
        'Box_TempNegTolEdit
        '
        Me.Box_TempNegTolEdit.Controls.Add(Me.Txt_TempNegTolEdit)
        Me.Box_TempNegTolEdit.Location = New System.Drawing.Point(120, 109)
        Me.Box_TempNegTolEdit.Name = "Box_TempNegTolEdit"
        Me.Box_TempNegTolEdit.Size = New System.Drawing.Size(129, 43)
        Me.Box_TempNegTolEdit.TabIndex = 19
        Me.Box_TempNegTolEdit.TabStop = False
        Me.Box_TempNegTolEdit.Text = "NegTol (°F)"
        '
        'Txt_TempNegTolEdit
        '
        Me.Txt_TempNegTolEdit.Location = New System.Drawing.Point(6, 19)
        Me.Txt_TempNegTolEdit.Name = "Txt_TempNegTolEdit"
        Me.Txt_TempNegTolEdit.Size = New System.Drawing.Size(100, 20)
        Me.Txt_TempNegTolEdit.TabIndex = 0
        Me.Txt_TempNegTolEdit.Text = "-50"
        '
        'Box_TempPosTolEdit
        '
        Me.Box_TempPosTolEdit.Controls.Add(Me.Txt_TempPosTolEdit)
        Me.Box_TempPosTolEdit.Location = New System.Drawing.Point(120, 64)
        Me.Box_TempPosTolEdit.Name = "Box_TempPosTolEdit"
        Me.Box_TempPosTolEdit.Size = New System.Drawing.Size(129, 43)
        Me.Box_TempPosTolEdit.TabIndex = 18
        Me.Box_TempPosTolEdit.TabStop = False
        Me.Box_TempPosTolEdit.Text = "PosTol (°F)"
        '
        'Txt_TempPosTolEdit
        '
        Me.Txt_TempPosTolEdit.Location = New System.Drawing.Point(6, 19)
        Me.Txt_TempPosTolEdit.Name = "Txt_TempPosTolEdit"
        Me.Txt_TempPosTolEdit.Size = New System.Drawing.Size(100, 20)
        Me.Txt_TempPosTolEdit.TabIndex = 0
        Me.Txt_TempPosTolEdit.Text = "50"
        '
        'Box_TempSetEdit
        '
        Me.Box_TempSetEdit.Controls.Add(Me.Txt_TempSetEdit)
        Me.Box_TempSetEdit.Location = New System.Drawing.Point(97, 19)
        Me.Box_TempSetEdit.Name = "Box_TempSetEdit"
        Me.Box_TempSetEdit.Size = New System.Drawing.Size(129, 43)
        Me.Box_TempSetEdit.TabIndex = 17
        Me.Box_TempSetEdit.TabStop = False
        Me.Box_TempSetEdit.Text = "Set Point (°F)"
        '
        'Txt_TempSetEdit
        '
        Me.Txt_TempSetEdit.Location = New System.Drawing.Point(6, 19)
        Me.Txt_TempSetEdit.Name = "Txt_TempSetEdit"
        Me.Txt_TempSetEdit.Size = New System.Drawing.Size(100, 20)
        Me.Txt_TempSetEdit.TabIndex = 0
        Me.Txt_TempSetEdit.Text = "500"
        '
        'Box_StepNameEdit
        '
        Me.Box_StepNameEdit.Controls.Add(Me.Txt_StepNameEdit)
        Me.Box_StepNameEdit.Location = New System.Drawing.Point(291, 40)
        Me.Box_StepNameEdit.Name = "Box_StepNameEdit"
        Me.Box_StepNameEdit.Size = New System.Drawing.Size(151, 46)
        Me.Box_StepNameEdit.TabIndex = 12
        Me.Box_StepNameEdit.TabStop = False
        Me.Box_StepNameEdit.Text = "Step Name"
        '
        'Txt_StepNameEdit
        '
        Me.Txt_StepNameEdit.Location = New System.Drawing.Point(6, 19)
        Me.Txt_StepNameEdit.Name = "Txt_StepNameEdit"
        Me.Txt_StepNameEdit.Size = New System.Drawing.Size(139, 20)
        Me.Txt_StepNameEdit.TabIndex = 11
        '
        'Box_CheckEdit
        '
        Me.Box_CheckEdit.Controls.Add(Me.Check_Vac)
        Me.Box_CheckEdit.Controls.Add(Me.Check_TempEdit)
        Me.Box_CheckEdit.Controls.Add(Me.Check_PressureEdit)
        Me.Box_CheckEdit.Location = New System.Drawing.Point(584, 9)
        Me.Box_CheckEdit.Name = "Box_CheckEdit"
        Me.Box_CheckEdit.Size = New System.Drawing.Size(118, 100)
        Me.Box_CheckEdit.TabIndex = 15
        Me.Box_CheckEdit.TabStop = False
        Me.Box_CheckEdit.Text = "Check"
        '
        'Check_Vac
        '
        Me.Check_Vac.AutoSize = True
        Me.Check_Vac.Location = New System.Drawing.Point(20, 65)
        Me.Check_Vac.Name = "Check_Vac"
        Me.Check_Vac.Size = New System.Drawing.Size(65, 17)
        Me.Check_Vac.TabIndex = 16
        Me.Check_Vac.Text = "Vacuum"
        Me.Check_Vac.UseVisualStyleBackColor = True
        '
        'Check_TempEdit
        '
        Me.Check_TempEdit.AutoSize = True
        Me.Check_TempEdit.Location = New System.Drawing.Point(20, 19)
        Me.Check_TempEdit.Name = "Check_TempEdit"
        Me.Check_TempEdit.Size = New System.Drawing.Size(86, 17)
        Me.Check_TempEdit.TabIndex = 15
        Me.Check_TempEdit.Text = "Temperature"
        Me.Check_TempEdit.UseVisualStyleBackColor = True
        '
        'Check_PressureEdit
        '
        Me.Check_PressureEdit.AutoSize = True
        Me.Check_PressureEdit.Location = New System.Drawing.Point(20, 42)
        Me.Check_PressureEdit.Name = "Check_PressureEdit"
        Me.Check_PressureEdit.Size = New System.Drawing.Size(67, 17)
        Me.Check_PressureEdit.TabIndex = 14
        Me.Check_PressureEdit.Text = "Pressure"
        Me.Check_PressureEdit.UseVisualStyleBackColor = True
        '
        'Box_CureNameEdit
        '
        Me.Box_CureNameEdit.Controls.Add(Me.Txt_CureNameEdit)
        Me.Box_CureNameEdit.Location = New System.Drawing.Point(385, 10)
        Me.Box_CureNameEdit.Name = "Box_CureNameEdit"
        Me.Box_CureNameEdit.Size = New System.Drawing.Size(213, 58)
        Me.Box_CureNameEdit.TabIndex = 13
        Me.Box_CureNameEdit.TabStop = False
        Me.Box_CureNameEdit.Text = "Cure Name"
        '
        'Txt_CureNameEdit
        '
        Me.Txt_CureNameEdit.Location = New System.Drawing.Point(6, 19)
        Me.Txt_CureNameEdit.Name = "Txt_CureNameEdit"
        Me.Txt_CureNameEdit.Size = New System.Drawing.Size(183, 20)
        Me.Txt_CureNameEdit.TabIndex = 11
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Txt_DocRevEdit)
        Me.GroupBox2.Location = New System.Drawing.Point(317, 59)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(54, 48)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Rev"
        '
        'Txt_DocRevEdit
        '
        Me.Txt_DocRevEdit.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_DocRevEdit.Location = New System.Drawing.Point(3, 16)
        Me.Txt_DocRevEdit.Name = "Txt_DocRevEdit"
        Me.Txt_DocRevEdit.Size = New System.Drawing.Size(48, 20)
        Me.Txt_DocRevEdit.TabIndex = 1
        Me.Txt_DocRevEdit.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Txt_CureDocEdit)
        Me.GroupBox3.Location = New System.Drawing.Point(8, 59)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(297, 48)
        Me.GroupBox3.TabIndex = 10
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Document"
        '
        'Txt_CureDocEdit
        '
        Me.Txt_CureDocEdit.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_CureDocEdit.Location = New System.Drawing.Point(3, 16)
        Me.Txt_CureDocEdit.Name = "Txt_CureDocEdit"
        Me.Txt_CureDocEdit.Size = New System.Drawing.Size(291, 20)
        Me.Txt_CureDocEdit.TabIndex = 1
        Me.Txt_CureDocEdit.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Combo_CureProfileEdit)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.5!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(363, 53)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Cure Profile"
        '
        'Combo_CureProfileEdit
        '
        Me.Combo_CureProfileEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Combo_CureProfileEdit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_CureProfileEdit.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_CureProfileEdit.FormattingEnabled = True
        Me.Combo_CureProfileEdit.Location = New System.Drawing.Point(6, 16)
        Me.Combo_CureProfileEdit.Name = "Combo_CureProfileEdit"
        Me.Combo_CureProfileEdit.Size = New System.Drawing.Size(351, 21)
        Me.Combo_CureProfileEdit.TabIndex = 6
        '
        'OpenCSVFileDialog
        '
        Me.OpenCSVFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        Me.OpenCSVFileDialog.Title = "Open Path"
        '
        'OpenCureProfileFileDialog
        '
        Me.OpenCureProfileFileDialog.CheckFileExists = False
        Me.OpenCureProfileFileDialog.CheckPathExists = False
        Me.OpenCureProfileFileDialog.Filter = "Cure Profile Files (*.cprof)|*.cprof|All Files (*.*)|*.*"
        Me.OpenCureProfileFileDialog.ValidateNames = False
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 548)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "MainForm"
        Me.Text = "Cure Report Builder 0.3.5"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TableLayoutPanel7.ResumeLayout(False)
        Me.Box_CureDataLocation.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.Box_RunParams.ResumeLayout(False)
        Me.Box_RunParams.PerformLayout()
        Me.Box_PartInfo.ResumeLayout(False)
        Me.Box_PartNumber.ResumeLayout(False)
        Me.Box_PartNumber.PerformLayout()
        Me.Box_Revision.ResumeLayout(False)
        Me.Box_Revision.PerformLayout()
        Me.Box_PartDesc.ResumeLayout(False)
        Me.Box_PartDesc.PerformLayout()
        Me.Box_Qty.ResumeLayout(False)
        Me.Box_Qty.PerformLayout()
        Me.Box_Vac.ResumeLayout(False)
        CType(Me.Data_Vac, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Box_TC.ResumeLayout(False)
        CType(Me.Data_TC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Box_JobNumber.ResumeLayout(False)
        Me.Box_JobNumber.PerformLayout()
        Me.Box_ProgramNumber.ResumeLayout(False)
        Me.Box_ProgramNumber.PerformLayout()
        Me.Box_DataRecorder.ResumeLayout(False)
        Me.Box_DataRecorder.PerformLayout()
        Me.Box_RunLine.ResumeLayout(False)
        Me.Box_Technician.ResumeLayout(False)
        Me.Box_Technician.PerformLayout()
        Me.TableLayoutPanel6.ResumeLayout(False)
        Me.Box_CureProfChoice.ResumeLayout(False)
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.Box_CureProfile.ResumeLayout(False)
        Me.TableLayoutPanel4.ResumeLayout(False)
        Me.Box_DocRev.ResumeLayout(False)
        Me.Box_DocRev.PerformLayout()
        Me.Box_CureDoc.ResumeLayout(False)
        Me.Box_CureDoc.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.Box_TemplatePath.ResumeLayout(False)
        Me.TableLayoutPanel5.ResumeLayout(False)
        Me.TableLayoutPanel5.PerformLayout()
        Me.Box_CureProfiles.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.Box_vacStepEdit.ResumeLayout(False)
        Me.Box_vacStepEdit.PerformLayout()
        Me.Box_vacNegTolEdit.ResumeLayout(False)
        Me.Box_vacNegTolEdit.PerformLayout()
        Me.Box_vacPosTolEdit.ResumeLayout(False)
        Me.Box_vacPosTolEdit.PerformLayout()
        Me.Box_vacSetEdit.ResumeLayout(False)
        Me.Box_vacSetEdit.PerformLayout()
        Me.Box_pressureStepEdit.ResumeLayout(False)
        Me.Box_pressureStepEdit.PerformLayout()
        Me.Box_pressureRampNegTolEdit.ResumeLayout(False)
        Me.Box_pressureRampNegTolEdit.PerformLayout()
        Me.Box_pressureRampPosTolEdit.ResumeLayout(False)
        Me.Box_pressureRampPosTolEdit.PerformLayout()
        Me.Box_pressureRampEdit.ResumeLayout(False)
        Me.Box_pressureRampEdit.PerformLayout()
        Me.Box_pressureNegTolEdit.ResumeLayout(False)
        Me.Box_pressureNegTolEdit.PerformLayout()
        Me.Box_pressurePosTolEdit.ResumeLayout(False)
        Me.Box_pressurePosTolEdit.PerformLayout()
        Me.Box_pressureSetEdit.ResumeLayout(False)
        Me.Box_pressureSetEdit.PerformLayout()
        Me.Box_tempStepEdit.ResumeLayout(False)
        Me.Box_tempStepEdit.PerformLayout()
        Me.Box_TempRampNegTolEdit.ResumeLayout(False)
        Me.Box_TempRampNegTolEdit.PerformLayout()
        Me.Box_TempRampPosTolEdit.ResumeLayout(False)
        Me.Box_TempRampPosTolEdit.PerformLayout()
        Me.Box_TempRampEdit.ResumeLayout(False)
        Me.Box_TempRampEdit.PerformLayout()
        Me.Box_TempNegTolEdit.ResumeLayout(False)
        Me.Box_TempNegTolEdit.PerformLayout()
        Me.Box_TempPosTolEdit.ResumeLayout(False)
        Me.Box_TempPosTolEdit.PerformLayout()
        Me.Box_TempSetEdit.ResumeLayout(False)
        Me.Box_TempSetEdit.PerformLayout()
        Me.Box_StepNameEdit.ResumeLayout(False)
        Me.Box_StepNameEdit.PerformLayout()
        Me.Box_CheckEdit.ResumeLayout(False)
        Me.Box_CheckEdit.PerformLayout()
        Me.Box_CureNameEdit.ResumeLayout(False)
        Me.Box_CureNameEdit.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As ToolStripStatusLabel
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents Box_CureDataLocation As GroupBox
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents OpenCSVFileDialog As OpenFileDialog
    Friend WithEvents Btn_OpenFile As Button
    Friend WithEvents Txt_FilePath As TextBox
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents Box_DataRecorder As GroupBox
    Friend WithEvents Txt_DataRecorder As TextBox
    Friend WithEvents Box_DocRev As GroupBox
    Friend WithEvents Txt_DocRev As TextBox
    Friend WithEvents Box_CureDoc As GroupBox
    Friend WithEvents Txt_CureDoc As TextBox
    Friend WithEvents Box_Qty As GroupBox
    Friend WithEvents Txt_Qty As TextBox
    Friend WithEvents Box_JobNumber As GroupBox
    Friend WithEvents Txt_JobNumber As TextBox
    Friend WithEvents Box_PartDesc As GroupBox
    Friend WithEvents Txt_PartDesc As TextBox
    Friend WithEvents Box_ProgramNumber As GroupBox
    Friend WithEvents Txt_ProgramNumber As TextBox
    Friend WithEvents Box_Revision As GroupBox
    Friend WithEvents Txt_Revision As TextBox
    Friend WithEvents Box_PartNumber As GroupBox
    Friend WithEvents Txt_PartNumber As TextBox
    Friend WithEvents Data_TC As DataGridView
    Friend WithEvents Box_CureProfile As GroupBox
    Friend WithEvents Combo_CureProfile As ComboBox
    Friend WithEvents Box_TC As GroupBox
    Friend WithEvents Used As DataGridViewCheckBoxColumn
    Friend WithEvents Modifier As DataGridViewTextBoxColumn
    Friend WithEvents Box_Vac As GroupBox
    Friend WithEvents Data_Vac As DataGridView
    Friend WithEvents DataGridViewCheckBoxColumn1 As DataGridViewCheckBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn1 As DataGridViewTextBoxColumn
    Friend WithEvents Box_Technician As GroupBox
    Friend WithEvents Txt_Technician As TextBox
    Friend WithEvents Box_CureProfiles As GroupBox
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents Txt_CureProfilesPath As TextBox
    Friend WithEvents Btn_LoadProfileFiles As Button
    Friend WithEvents OpenCureProfileFileDialog As OpenFileDialog
    Friend WithEvents Box_CureProfChoice As GroupBox
    Friend WithEvents TableLayoutPanel3 As TableLayoutPanel
    Friend WithEvents TableLayoutPanel4 As TableLayoutPanel
    Friend WithEvents Box_TemplatePath As GroupBox
    Friend WithEvents TableLayoutPanel5 As TableLayoutPanel
    Friend WithEvents Txt_TemplatePath As TextBox
    Friend WithEvents Btn_Run As Button
    Friend WithEvents Box_PartInfo As GroupBox
    Friend WithEvents TableLayoutPanel7 As TableLayoutPanel
    Friend WithEvents Box_RunParams As GroupBox
    Friend WithEvents Box_RunLine As TableLayoutPanel
    Friend WithEvents TableLayoutPanel6 As TableLayoutPanel
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Btn_ClearCells As Button
    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Txt_DocRevEdit As TextBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents Txt_CureDocEdit As TextBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Combo_CureProfileEdit As ComboBox
    Friend WithEvents Box_tempStepEdit As GroupBox
    Friend WithEvents Box_TempRampNegTolEdit As GroupBox
    Friend WithEvents Txt_TempRampNegTolEdit As TextBox
    Friend WithEvents Box_TempRampPosTolEdit As GroupBox
    Friend WithEvents Txt_TempRampPosTolEdit As TextBox
    Friend WithEvents Box_TempRampEdit As GroupBox
    Friend WithEvents Txt_TempRampEdit As TextBox
    Friend WithEvents Box_TempNegTolEdit As GroupBox
    Friend WithEvents Txt_TempNegTolEdit As TextBox
    Friend WithEvents Box_TempPosTolEdit As GroupBox
    Friend WithEvents Txt_TempPosTolEdit As TextBox
    Friend WithEvents Box_TempSetEdit As GroupBox
    Friend WithEvents Txt_TempSetEdit As TextBox
    Friend WithEvents Box_CheckEdit As GroupBox
    Friend WithEvents Check_Vac As CheckBox
    Friend WithEvents Check_TempEdit As CheckBox
    Friend WithEvents Check_PressureEdit As CheckBox
    Friend WithEvents Box_CureNameEdit As GroupBox
    Friend WithEvents Txt_CureNameEdit As TextBox
    Friend WithEvents Box_StepNameEdit As GroupBox
    Friend WithEvents Txt_StepNameEdit As TextBox
    Friend WithEvents Check_tempStepEdit As CheckBox
    Friend WithEvents Check_tempMinEdit As CheckBox
    Friend WithEvents Check_tempMaxEdit As CheckBox
    Friend WithEvents Box_pressureStepEdit As GroupBox
    Friend WithEvents Check_pressureMinEdit As CheckBox
    Friend WithEvents Check_pressureMaxEdit As CheckBox
    Friend WithEvents Check_pressureStepEdit As CheckBox
    Friend WithEvents Box_pressureRampNegTolEdit As GroupBox
    Friend WithEvents Txt_pressureRampNegTolEdit As TextBox
    Friend WithEvents Box_pressureRampPosTolEdit As GroupBox
    Friend WithEvents Txt_pressureRampPosTolEdit As TextBox
    Friend WithEvents Box_pressureRampEdit As GroupBox
    Friend WithEvents Txt_pressureRampEdit As TextBox
    Friend WithEvents Box_pressureNegTolEdit As GroupBox
    Friend WithEvents Txt_pressureNegTolEdit As TextBox
    Friend WithEvents Box_pressurePosTolEdit As GroupBox
    Friend WithEvents Txt_pressurePosTolEdit As TextBox
    Friend WithEvents Box_pressureSetEdit As GroupBox
    Friend WithEvents Txt_pressureSetEdit As TextBox
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents Box_vacStepEdit As GroupBox
    Friend WithEvents Check_vacMinEdit As CheckBox
    Friend WithEvents Check_vacMaxEdit As CheckBox
    Friend WithEvents Check_vacStepEdit As CheckBox
    Friend WithEvents Box_vacNegTolEdit As GroupBox
    Friend WithEvents Txt_vacNegTolEdit As TextBox
    Friend WithEvents Box_vacPosTolEdit As GroupBox
    Friend WithEvents Txt_vacPosTolEdit As TextBox
    Friend WithEvents Box_vacSetEdit As GroupBox
    Friend WithEvents Txt_vacSetEdit As TextBox
End Class
