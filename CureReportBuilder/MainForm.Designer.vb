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
        Me.Box_CureProfChoice = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.Box_CureProfile = New System.Windows.Forms.GroupBox()
        Me.Combo_CureProfile = New System.Windows.Forms.ComboBox()
        Me.TableLayoutPanel4 = New System.Windows.Forms.TableLayoutPanel()
        Me.Box_DocRev = New System.Windows.Forms.GroupBox()
        Me.Txt_DocRev = New System.Windows.Forms.TextBox()
        Me.Box_CureDoc = New System.Windows.Forms.GroupBox()
        Me.Txt_CureDoc = New System.Windows.Forms.TextBox()
        Me.Box_RunParams = New System.Windows.Forms.GroupBox()
        Me.Box_PartInfo = New System.Windows.Forms.GroupBox()
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
        Me.TCNum = New System.Windows.Forms.DataGridViewTextBoxColumn()
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
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Box_TemplatePath = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel5 = New System.Windows.Forms.TableLayoutPanel()
        Me.Txt_TemplatePath = New System.Windows.Forms.TextBox()
        Me.Box_CureProfiles = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.Txt_CureProfilesPath = New System.Windows.Forms.TextBox()
        Me.Btn_LoadProfileFiles = New System.Windows.Forms.Button()
        Me.OpenCSVFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.OpenCureProfileFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.TableLayoutPanel6 = New System.Windows.Forms.TableLayoutPanel()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Btn_ClearCells = New System.Windows.Forms.Button()
        Me.MenuStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TableLayoutPanel7.SuspendLayout()
        Me.Box_CureDataLocation.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Box_CureProfChoice.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.Box_CureProfile.SuspendLayout()
        Me.TableLayoutPanel4.SuspendLayout()
        Me.Box_DocRev.SuspendLayout()
        Me.Box_CureDoc.SuspendLayout()
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
        Me.TabPage2.SuspendLayout()
        Me.Box_TemplatePath.SuspendLayout()
        Me.TableLayoutPanel5.SuspendLayout()
        Me.Box_CureProfiles.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.TableLayoutPanel6.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(761, 24)
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
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 526)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(761, 22)
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
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 24)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(761, 502)
        Me.TabControl1.TabIndex = 2
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.TableLayoutPanel7)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(753, 476)
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
        Me.TableLayoutPanel7.Size = New System.Drawing.Size(747, 470)
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
        Me.Box_CureDataLocation.Size = New System.Drawing.Size(721, 54)
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
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(715, 35)
        Me.TableLayoutPanel1.TabIndex = 2
        '
        'Txt_FilePath
        '
        Me.Txt_FilePath.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_FilePath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Txt_FilePath.Location = New System.Drawing.Point(3, 7)
        Me.Txt_FilePath.Name = "Txt_FilePath"
        Me.Txt_FilePath.Size = New System.Drawing.Size(634, 20)
        Me.Txt_FilePath.TabIndex = 0
        Me.Txt_FilePath.Text = "File path..."
        '
        'Btn_OpenFile
        '
        Me.Btn_OpenFile.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Btn_OpenFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_OpenFile.Location = New System.Drawing.Point(643, 3)
        Me.Btn_OpenFile.Name = "Btn_OpenFile"
        Me.Btn_OpenFile.Size = New System.Drawing.Size(69, 29)
        Me.Btn_OpenFile.TabIndex = 1
        Me.Btn_OpenFile.Text = "Open File"
        Me.Btn_OpenFile.UseVisualStyleBackColor = True
        '
        'Box_CureProfChoice
        '
        Me.Box_CureProfChoice.Controls.Add(Me.TableLayoutPanel3)
        Me.Box_CureProfChoice.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Box_CureProfChoice.Location = New System.Drawing.Point(3, 3)
        Me.Box_CureProfChoice.Name = "Box_CureProfChoice"
        Me.Box_CureProfChoice.Size = New System.Drawing.Size(375, 138)
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
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(369, 119)
        Me.TableLayoutPanel3.TabIndex = 12
        '
        'Box_CureProfile
        '
        Me.Box_CureProfile.Controls.Add(Me.Combo_CureProfile)
        Me.Box_CureProfile.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Box_CureProfile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.5!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Box_CureProfile.Location = New System.Drawing.Point(3, 3)
        Me.Box_CureProfile.Name = "Box_CureProfile"
        Me.Box_CureProfile.Size = New System.Drawing.Size(363, 53)
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
        Me.Combo_CureProfile.Size = New System.Drawing.Size(351, 21)
        Me.Combo_CureProfile.TabIndex = 6
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
        Me.TableLayoutPanel4.Size = New System.Drawing.Size(363, 54)
        Me.TableLayoutPanel4.TabIndex = 8
        '
        'Box_DocRev
        '
        Me.Box_DocRev.Controls.Add(Me.Txt_DocRev)
        Me.Box_DocRev.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Box_DocRev.Location = New System.Drawing.Point(306, 3)
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
        Me.Txt_DocRev.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Box_CureDoc
        '
        Me.Box_CureDoc.Controls.Add(Me.Txt_CureDoc)
        Me.Box_CureDoc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Box_CureDoc.Location = New System.Drawing.Point(3, 3)
        Me.Box_CureDoc.Name = "Box_CureDoc"
        Me.Box_CureDoc.Size = New System.Drawing.Size(297, 48)
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
        Me.Txt_CureDoc.Size = New System.Drawing.Size(291, 20)
        Me.Txt_CureDoc.TabIndex = 1
        Me.Txt_CureDoc.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
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
        Me.Box_RunParams.Size = New System.Drawing.Size(721, 194)
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
        Me.Txt_PartNumber.TabIndex = 1
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
        Me.Txt_Revision.TabIndex = 1
        Me.Txt_Revision.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Box_PartDesc
        '
        Me.Box_PartDesc.Controls.Add(Me.Txt_PartDesc)
        Me.Box_PartDesc.Location = New System.Drawing.Point(6, 69)
        Me.Box_PartDesc.Name = "Box_PartDesc"
        Me.Box_PartDesc.Size = New System.Drawing.Size(351, 44)
        Me.Box_PartDesc.TabIndex = 3
        Me.Box_PartDesc.TabStop = False
        Me.Box_PartDesc.Text = "Part Nomenclature"
        '
        'Txt_PartDesc
        '
        Me.Txt_PartDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_PartDesc.Location = New System.Drawing.Point(3, 16)
        Me.Txt_PartDesc.Name = "Txt_PartDesc"
        Me.Txt_PartDesc.Size = New System.Drawing.Size(345, 20)
        Me.Txt_PartDesc.TabIndex = 1
        Me.Txt_PartDesc.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Box_Qty
        '
        Me.Box_Qty.Controls.Add(Me.Txt_Qty)
        Me.Box_Qty.Location = New System.Drawing.Point(198, 20)
        Me.Box_Qty.Name = "Box_Qty"
        Me.Box_Qty.Size = New System.Drawing.Size(51, 44)
        Me.Box_Qty.TabIndex = 3
        Me.Box_Qty.TabStop = False
        Me.Box_Qty.Text = "Qty"
        '
        'Txt_Qty
        '
        Me.Txt_Qty.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_Qty.Location = New System.Drawing.Point(3, 16)
        Me.Txt_Qty.Name = "Txt_Qty"
        Me.Txt_Qty.Size = New System.Drawing.Size(45, 20)
        Me.Txt_Qty.TabIndex = 1
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
        Me.Data_TC.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Used, Me.TCNum})
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
        'TCNum
        '
        Me.TCNum.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.TCNum.HeaderText = "TC Number"
        Me.TCNum.Name = "TCNum"
        Me.TCNum.ReadOnly = True
        Me.TCNum.Width = 86
        '
        'Box_JobNumber
        '
        Me.Box_JobNumber.Controls.Add(Me.Txt_JobNumber)
        Me.Box_JobNumber.Location = New System.Drawing.Point(6, 19)
        Me.Box_JobNumber.Name = "Box_JobNumber"
        Me.Box_JobNumber.Size = New System.Drawing.Size(114, 44)
        Me.Box_JobNumber.TabIndex = 4
        Me.Box_JobNumber.TabStop = False
        Me.Box_JobNumber.Text = "Job Number"
        '
        'Txt_JobNumber
        '
        Me.Txt_JobNumber.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_JobNumber.Location = New System.Drawing.Point(3, 16)
        Me.Txt_JobNumber.Name = "Txt_JobNumber"
        Me.Txt_JobNumber.Size = New System.Drawing.Size(108, 20)
        Me.Txt_JobNumber.TabIndex = 1
        Me.Txt_JobNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Box_ProgramNumber
        '
        Me.Box_ProgramNumber.Controls.Add(Me.Txt_ProgramNumber)
        Me.Box_ProgramNumber.Location = New System.Drawing.Point(126, 19)
        Me.Box_ProgramNumber.Name = "Box_ProgramNumber"
        Me.Box_ProgramNumber.Size = New System.Drawing.Size(114, 44)
        Me.Box_ProgramNumber.TabIndex = 3
        Me.Box_ProgramNumber.TabStop = False
        Me.Box_ProgramNumber.Text = "Program Number"
        '
        'Txt_ProgramNumber
        '
        Me.Txt_ProgramNumber.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_ProgramNumber.Location = New System.Drawing.Point(3, 16)
        Me.Txt_ProgramNumber.Name = "Txt_ProgramNumber"
        Me.Txt_ProgramNumber.Size = New System.Drawing.Size(108, 20)
        Me.Txt_ProgramNumber.TabIndex = 1
        Me.Txt_ProgramNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Box_DataRecorder
        '
        Me.Box_DataRecorder.Controls.Add(Me.Txt_DataRecorder)
        Me.Box_DataRecorder.Location = New System.Drawing.Point(246, 19)
        Me.Box_DataRecorder.Name = "Box_DataRecorder"
        Me.Box_DataRecorder.Size = New System.Drawing.Size(142, 44)
        Me.Box_DataRecorder.TabIndex = 3
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
        Me.Txt_DataRecorder.TabIndex = 1
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
        Me.Box_RunLine.Size = New System.Drawing.Size(721, 54)
        Me.Box_RunLine.TabIndex = 14
        '
        'Btn_Run
        '
        Me.Btn_Run.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Btn_Run.FlatAppearance.BorderSize = 2
        Me.Btn_Run.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Btn_Run.Location = New System.Drawing.Point(558, 3)
        Me.Btn_Run.Name = "Btn_Run"
        Me.Btn_Run.Size = New System.Drawing.Size(160, 48)
        Me.Btn_Run.TabIndex = 12
        Me.Btn_Run.Text = "Run Output"
        Me.Btn_Run.UseVisualStyleBackColor = True
        '
        'Box_Technician
        '
        Me.Box_Technician.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Box_Technician.Controls.Add(Me.Txt_Technician)
        Me.Box_Technician.Location = New System.Drawing.Point(391, 5)
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
        Me.Txt_Technician.TabIndex = 1
        Me.Txt_Technician.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Box_TemplatePath)
        Me.TabPage2.Controls.Add(Me.Box_CureProfiles)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(753, 476)
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
        Me.Txt_TemplatePath.Text = "C:\Users\Will Eagan\Source\repos\CureReportBuilder\CureReportBuilder\Sample Files" &
    "\Cure Report_Template.xlsx"
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
        Me.Txt_CureProfilesPath.Text = "C:\Users\Will Eagan\Source\repos\CureReportBuilder\CureReportBuilder\Sample Files" &
    "\Cure Profiles"
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
        Me.TableLayoutPanel6.Size = New System.Drawing.Size(721, 144)
        Me.TableLayoutPanel6.TabIndex = 15
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.CureReportBuilder.My.Resources.Resources.Systima_Composites_2018
        Me.PictureBox1.Location = New System.Drawing.Point(404, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(294, 138)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 12
        Me.PictureBox1.TabStop = False
        '
        'Btn_ClearCells
        '
        Me.Btn_ClearCells.Location = New System.Drawing.Point(264, 24)
        Me.Btn_ClearCells.Name = "Btn_ClearCells"
        Me.Btn_ClearCells.Size = New System.Drawing.Size(104, 38)
        Me.Btn_ClearCells.TabIndex = 4
        Me.Btn_ClearCells.Text = "Clear Cells"
        Me.Btn_ClearCells.UseVisualStyleBackColor = True
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(761, 548)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "MainForm"
        Me.Text = "Cure Report Builder"
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
        Me.Box_CureProfChoice.ResumeLayout(False)
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.Box_CureProfile.ResumeLayout(False)
        Me.TableLayoutPanel4.ResumeLayout(False)
        Me.Box_DocRev.ResumeLayout(False)
        Me.Box_DocRev.PerformLayout()
        Me.Box_CureDoc.ResumeLayout(False)
        Me.Box_CureDoc.PerformLayout()
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
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.Box_TemplatePath.ResumeLayout(False)
        Me.TableLayoutPanel5.ResumeLayout(False)
        Me.TableLayoutPanel5.PerformLayout()
        Me.Box_CureProfiles.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.TableLayoutPanel6.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents TCNum As DataGridViewTextBoxColumn
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
End Class
