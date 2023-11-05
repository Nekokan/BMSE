<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmWindowFind
#Region "Windows フォーム デザイナによって生成されたコード "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'この呼び出しは、Windows フォーム デザイナで必要です。
		InitializeComponent()
	End Sub
	'Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows フォーム デザイナで必要です。
	Private components As System.ComponentModel.IContainer
    Public WithEvents cmdClose As System.Windows.Forms.Button
    Public WithEvents cmdDecide As System.Windows.Forms.Button
	Public WithEvents txtReplace As System.Windows.Forms.TextBox
	Public WithEvents optProcessReplace As System.Windows.Forms.RadioButton
	Public WithEvents optProcessDelete As System.Windows.Forms.RadioButton
	Public WithEvents optProcessSelect As System.Windows.Forms.RadioButton
	Public WithEvents fraProcess As System.Windows.Forms.GroupBox
	Public WithEvents cmdInvert As System.Windows.Forms.Button
	Public WithEvents cmdReset As System.Windows.Forms.Button
	Public WithEvents cmdSelect As System.Windows.Forms.Button
	Public WithEvents _lstGrid_2 As System.Windows.Forms.CheckedListBox
	Public WithEvents _lstGrid_1 As System.Windows.Forms.CheckedListBox
	Public WithEvents _lstGrid_0 As System.Windows.Forms.CheckedListBox
	Public WithEvents _lstGrid_3 As System.Windows.Forms.CheckedListBox
	Public WithEvents lblBGM As System.Windows.Forms.Label
	Public WithEvents lblPlayer2 As System.Windows.Forms.Label
	Public WithEvents lblPlayer1 As System.Windows.Forms.Label
	Public WithEvents lblEtc As System.Windows.Forms.Label
	Public WithEvents fraSearchGrid As System.Windows.Forms.GroupBox
	Public WithEvents txtNumMax As System.Windows.Forms.TextBox
	Public WithEvents txtNumMin As System.Windows.Forms.TextBox
	Public WithEvents lblNotice As System.Windows.Forms.Label
	Public WithEvents lblNum As System.Windows.Forms.Label
	Public WithEvents fraSearchNum As System.Windows.Forms.GroupBox
	Public WithEvents txtMeasureMax As System.Windows.Forms.TextBox
	Public WithEvents txtMeasureMin As System.Windows.Forms.TextBox
	Public WithEvents lblMeasure As System.Windows.Forms.Label
	Public WithEvents fraSearchMeasure As System.Windows.Forms.GroupBox
	Public WithEvents optSearchSelect As System.Windows.Forms.RadioButton
	Public WithEvents optSearchAll As System.Windows.Forms.RadioButton
	Public WithEvents fraSearchObject As System.Windows.Forms.GroupBox
    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使って変更できます。
    'コード エディタを使用して、変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.cmdDecide = New System.Windows.Forms.Button()
        Me.fraProcess = New System.Windows.Forms.GroupBox()
        Me.txtReplace = New System.Windows.Forms.TextBox()
        Me.optProcessReplace = New System.Windows.Forms.RadioButton()
        Me.optProcessDelete = New System.Windows.Forms.RadioButton()
        Me.optProcessSelect = New System.Windows.Forms.RadioButton()
        Me.fraSearchGrid = New System.Windows.Forms.GroupBox()
        Me.cmdInvert = New System.Windows.Forms.Button()
        Me.cmdReset = New System.Windows.Forms.Button()
        Me.cmdSelect = New System.Windows.Forms.Button()
        Me._lstGrid_2 = New System.Windows.Forms.CheckedListBox()
        Me._lstGrid_1 = New System.Windows.Forms.CheckedListBox()
        Me._lstGrid_0 = New System.Windows.Forms.CheckedListBox()
        Me._lstGrid_3 = New System.Windows.Forms.CheckedListBox()
        Me.lblBGM = New System.Windows.Forms.Label()
        Me.lblPlayer2 = New System.Windows.Forms.Label()
        Me.lblPlayer1 = New System.Windows.Forms.Label()
        Me.lblEtc = New System.Windows.Forms.Label()
        Me.fraSearchNum = New System.Windows.Forms.GroupBox()
        Me.txtNumMax = New System.Windows.Forms.TextBox()
        Me.txtNumMin = New System.Windows.Forms.TextBox()
        Me.lblNotice = New System.Windows.Forms.Label()
        Me.lblNum = New System.Windows.Forms.Label()
        Me.fraSearchMeasure = New System.Windows.Forms.GroupBox()
        Me.txtMeasureMax = New System.Windows.Forms.TextBox()
        Me.txtMeasureMin = New System.Windows.Forms.TextBox()
        Me.lblMeasure = New System.Windows.Forms.Label()
        Me.fraSearchObject = New System.Windows.Forms.GroupBox()
        Me.optSearchSelect = New System.Windows.Forms.RadioButton()
        Me.optSearchAll = New System.Windows.Forms.RadioButton()
        Me.fraProcess.SuspendLayout()
        Me.fraSearchGrid.SuspendLayout()
        Me.fraSearchNum.SuspendLayout()
        Me.fraSearchMeasure.SuspendLayout()
        Me.fraSearchObject.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdClose
        '
        Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
        Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdClose.Location = New System.Drawing.Point(436, 136)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdClose.Size = New System.Drawing.Size(125, 25)
        Me.cmdClose.TabIndex = 29
        Me.cmdClose.Text = "閉じる"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'cmdDecide
        '
        Me.cmdDecide.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDecide.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDecide.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDecide.Location = New System.Drawing.Point(436, 164)
        Me.cmdDecide.Name = "cmdDecide"
        Me.cmdDecide.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDecide.Size = New System.Drawing.Size(125, 25)
        Me.cmdDecide.TabIndex = 30
        Me.cmdDecide.Text = "実行"
        Me.cmdDecide.UseVisualStyleBackColor = False
        '
        'fraProcess
        '
        Me.fraProcess.BackColor = System.Drawing.SystemColors.Control
        Me.fraProcess.Controls.Add(Me.txtReplace)
        Me.fraProcess.Controls.Add(Me.optProcessReplace)
        Me.fraProcess.Controls.Add(Me.optProcessDelete)
        Me.fraProcess.Controls.Add(Me.optProcessSelect)
        Me.fraProcess.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraProcess.Location = New System.Drawing.Point(436, 4)
        Me.fraProcess.Name = "fraProcess"
        Me.fraProcess.Padding = New System.Windows.Forms.Padding(0)
        Me.fraProcess.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraProcess.Size = New System.Drawing.Size(125, 81)
        Me.fraProcess.TabIndex = 24
        Me.fraProcess.TabStop = False
        Me.fraProcess.Text = "処理"
        '
        'txtReplace
        '
        Me.txtReplace.AcceptsReturn = True
        Me.txtReplace.BackColor = System.Drawing.SystemColors.Window
        Me.txtReplace.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtReplace.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtReplace.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtReplace.Location = New System.Drawing.Point(92, 56)
        Me.txtReplace.MaxLength = 2
        Me.txtReplace.Name = "txtReplace"
        Me.txtReplace.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtReplace.Size = New System.Drawing.Size(25, 19)
        Me.txtReplace.TabIndex = 28
        Me.txtReplace.Text = "01"
        '
        'optProcessReplace
        '
        Me.optProcessReplace.BackColor = System.Drawing.SystemColors.Control
        Me.optProcessReplace.Cursor = System.Windows.Forms.Cursors.Default
        Me.optProcessReplace.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optProcessReplace.Location = New System.Drawing.Point(8, 56)
        Me.optProcessReplace.Name = "optProcessReplace"
        Me.optProcessReplace.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optProcessReplace.Size = New System.Drawing.Size(109, 21)
        Me.optProcessReplace.TabIndex = 27
        Me.optProcessReplace.TabStop = True
        Me.optProcessReplace.Text = "置換"
        Me.optProcessReplace.UseVisualStyleBackColor = False
        '
        'optProcessDelete
        '
        Me.optProcessDelete.BackColor = System.Drawing.SystemColors.Control
        Me.optProcessDelete.Cursor = System.Windows.Forms.Cursors.Default
        Me.optProcessDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optProcessDelete.Location = New System.Drawing.Point(8, 36)
        Me.optProcessDelete.Name = "optProcessDelete"
        Me.optProcessDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optProcessDelete.Size = New System.Drawing.Size(109, 21)
        Me.optProcessDelete.TabIndex = 26
        Me.optProcessDelete.TabStop = True
        Me.optProcessDelete.Text = "削除"
        Me.optProcessDelete.UseVisualStyleBackColor = False
        '
        'optProcessSelect
        '
        Me.optProcessSelect.BackColor = System.Drawing.SystemColors.Control
        Me.optProcessSelect.Checked = True
        Me.optProcessSelect.Cursor = System.Windows.Forms.Cursors.Default
        Me.optProcessSelect.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optProcessSelect.Location = New System.Drawing.Point(8, 16)
        Me.optProcessSelect.Name = "optProcessSelect"
        Me.optProcessSelect.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optProcessSelect.Size = New System.Drawing.Size(109, 21)
        Me.optProcessSelect.TabIndex = 25
        Me.optProcessSelect.TabStop = True
        Me.optProcessSelect.Text = "選択"
        Me.optProcessSelect.UseVisualStyleBackColor = False
        '
        'fraSearchGrid
        '
        Me.fraSearchGrid.BackColor = System.Drawing.SystemColors.Control
        Me.fraSearchGrid.Controls.Add(Me.cmdInvert)
        Me.fraSearchGrid.Controls.Add(Me.cmdReset)
        Me.fraSearchGrid.Controls.Add(Me.cmdSelect)
        Me.fraSearchGrid.Controls.Add(Me._lstGrid_2)
        Me.fraSearchGrid.Controls.Add(Me._lstGrid_1)
        Me.fraSearchGrid.Controls.Add(Me._lstGrid_0)
        Me.fraSearchGrid.Controls.Add(Me._lstGrid_3)
        Me.fraSearchGrid.Controls.Add(Me.lblBGM)
        Me.fraSearchGrid.Controls.Add(Me.lblPlayer2)
        Me.fraSearchGrid.Controls.Add(Me.lblPlayer1)
        Me.fraSearchGrid.Controls.Add(Me.lblEtc)
        Me.fraSearchGrid.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraSearchGrid.Location = New System.Drawing.Point(192, 4)
        Me.fraSearchGrid.Name = "fraSearchGrid"
        Me.fraSearchGrid.Padding = New System.Windows.Forms.Padding(0)
        Me.fraSearchGrid.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraSearchGrid.Size = New System.Drawing.Size(237, 185)
        Me.fraSearchGrid.TabIndex = 12
        Me.fraSearchGrid.TabStop = False
        Me.fraSearchGrid.Text = "列の指定"
        '
        'cmdInvert
        '
        Me.cmdInvert.BackColor = System.Drawing.SystemColors.Control
        Me.cmdInvert.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdInvert.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdInvert.Location = New System.Drawing.Point(20, 156)
        Me.cmdInvert.Name = "cmdInvert"
        Me.cmdInvert.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdInvert.Size = New System.Drawing.Size(61, 21)
        Me.cmdInvert.TabIndex = 21
        Me.cmdInvert.Text = "反転"
        Me.cmdInvert.UseVisualStyleBackColor = False
        '
        'cmdReset
        '
        Me.cmdReset.BackColor = System.Drawing.SystemColors.Control
        Me.cmdReset.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdReset.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdReset.Location = New System.Drawing.Point(84, 156)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdReset.Size = New System.Drawing.Size(61, 21)
        Me.cmdReset.TabIndex = 22
        Me.cmdReset.Text = "全解除"
        Me.cmdReset.UseVisualStyleBackColor = False
        '
        'cmdSelect
        '
        Me.cmdSelect.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSelect.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSelect.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSelect.Location = New System.Drawing.Point(148, 156)
        Me.cmdSelect.Name = "cmdSelect"
        Me.cmdSelect.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSelect.Size = New System.Drawing.Size(81, 21)
        Me.cmdSelect.TabIndex = 23
        Me.cmdSelect.Text = "全選択"
        Me.cmdSelect.UseVisualStyleBackColor = False
        '
        '_lstGrid_2
        '
        Me._lstGrid_2.BackColor = System.Drawing.SystemColors.Window
        Me._lstGrid_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lstGrid_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lstGrid_2.Items.AddRange(New Object() {"001"})
        Me._lstGrid_2.Location = New System.Drawing.Point(110, 36)
        Me._lstGrid_2.Name = "_lstGrid_2"
        Me._lstGrid_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lstGrid_2.Size = New System.Drawing.Size(63, 116)
        Me._lstGrid_2.TabIndex = 18
        '
        '_lstGrid_1
        '
        Me._lstGrid_1.BackColor = System.Drawing.SystemColors.Window
        Me._lstGrid_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lstGrid_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lstGrid_1.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "SC"})
        Me._lstGrid_1.Location = New System.Drawing.Point(59, 36)
        Me._lstGrid_1.Name = "_lstGrid_1"
        Me._lstGrid_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lstGrid_1.Size = New System.Drawing.Size(48, 116)
        Me._lstGrid_1.TabIndex = 16
        '
        '_lstGrid_0
        '
        Me._lstGrid_0.BackColor = System.Drawing.SystemColors.Window
        Me._lstGrid_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lstGrid_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lstGrid_0.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "SC"})
        Me._lstGrid_0.Location = New System.Drawing.Point(8, 36)
        Me._lstGrid_0.Name = "_lstGrid_0"
        Me._lstGrid_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lstGrid_0.Size = New System.Drawing.Size(48, 116)
        Me._lstGrid_0.TabIndex = 14
        '
        '_lstGrid_3
        '
        Me._lstGrid_3.BackColor = System.Drawing.SystemColors.Window
        Me._lstGrid_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lstGrid_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me._lstGrid_3.Items.AddRange(New Object() {"BPM", "STOP", "BGA", "Layer", "Poor"})
        Me._lstGrid_3.Location = New System.Drawing.Point(176, 36)
        Me._lstGrid_3.Name = "_lstGrid_3"
        Me._lstGrid_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lstGrid_3.Size = New System.Drawing.Size(53, 116)
        Me._lstGrid_3.TabIndex = 20
        '
        'lblBGM
        '
        Me.lblBGM.AutoSize = True
        Me.lblBGM.BackColor = System.Drawing.SystemColors.Control
        Me.lblBGM.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblBGM.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBGM.Location = New System.Drawing.Point(110, 20)
        Me.lblBGM.Name = "lblBGM"
        Me.lblBGM.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblBGM.Size = New System.Drawing.Size(23, 12)
        Me.lblBGM.TabIndex = 17
        Me.lblBGM.Text = "BGM"
        '
        'lblPlayer2
        '
        Me.lblPlayer2.AutoSize = True
        Me.lblPlayer2.BackColor = System.Drawing.SystemColors.Control
        Me.lblPlayer2.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPlayer2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPlayer2.Location = New System.Drawing.Point(59, 20)
        Me.lblPlayer2.Name = "lblPlayer2"
        Me.lblPlayer2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPlayer2.Size = New System.Drawing.Size(53, 12)
        Me.lblPlayer2.TabIndex = 15
        Me.lblPlayer2.Text = "2P"
        '
        'lblPlayer1
        '
        Me.lblPlayer1.AutoSize = True
        Me.lblPlayer1.BackColor = System.Drawing.SystemColors.Control
        Me.lblPlayer1.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPlayer1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPlayer1.Location = New System.Drawing.Point(8, 20)
        Me.lblPlayer1.Name = "lblPlayer1"
        Me.lblPlayer1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPlayer1.Size = New System.Drawing.Size(53, 12)
        Me.lblPlayer1.TabIndex = 13
        Me.lblPlayer1.Text = "1P"
        '
        'lblEtc
        '
        Me.lblEtc.AutoSize = True
        Me.lblEtc.BackColor = System.Drawing.SystemColors.Control
        Me.lblEtc.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblEtc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblEtc.Location = New System.Drawing.Point(176, 20)
        Me.lblEtc.Name = "lblEtc"
        Me.lblEtc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEtc.Size = New System.Drawing.Size(23, 12)
        Me.lblEtc.TabIndex = 19
        Me.lblEtc.Text = "Etc"
        '
        'fraSearchNum
        '
        Me.fraSearchNum.BackColor = System.Drawing.SystemColors.Control
        Me.fraSearchNum.Controls.Add(Me.txtNumMax)
        Me.fraSearchNum.Controls.Add(Me.txtNumMin)
        Me.fraSearchNum.Controls.Add(Me.lblNotice)
        Me.fraSearchNum.Controls.Add(Me.lblNum)
        Me.fraSearchNum.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraSearchNum.Location = New System.Drawing.Point(8, 112)
        Me.fraSearchNum.Name = "fraSearchNum"
        Me.fraSearchNum.Padding = New System.Windows.Forms.Padding(0)
        Me.fraSearchNum.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraSearchNum.Size = New System.Drawing.Size(177, 78)
        Me.fraSearchNum.TabIndex = 7
        Me.fraSearchNum.TabStop = False
        Me.fraSearchNum.Text = "オブジェ番号の指定"
        '
        'txtNumMax
        '
        Me.txtNumMax.AcceptsReturn = True
        Me.txtNumMax.BackColor = System.Drawing.SystemColors.Window
        Me.txtNumMax.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNumMax.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNumMax.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtNumMax.Location = New System.Drawing.Point(92, 16)
        Me.txtNumMax.MaxLength = 2
        Me.txtNumMax.Name = "txtNumMax"
        Me.txtNumMax.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNumMax.Size = New System.Drawing.Size(37, 19)
        Me.txtNumMax.TabIndex = 10
        Me.txtNumMax.Text = "ZZ"
        '
        'txtNumMin
        '
        Me.txtNumMin.AcceptsReturn = True
        Me.txtNumMin.BackColor = System.Drawing.SystemColors.Window
        Me.txtNumMin.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNumMin.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNumMin.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtNumMin.Location = New System.Drawing.Point(28, 16)
        Me.txtNumMin.MaxLength = 2
        Me.txtNumMin.Name = "txtNumMin"
        Me.txtNumMin.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNumMin.Size = New System.Drawing.Size(37, 19)
        Me.txtNumMin.TabIndex = 8
        Me.txtNumMin.Text = "01"
        '
        'lblNotice
        '
        Me.lblNotice.BackColor = System.Drawing.SystemColors.Control
        Me.lblNotice.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNotice.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNotice.Location = New System.Drawing.Point(8, 39)
        Me.lblNotice.Name = "lblNotice"
        Me.lblNotice.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNotice.Size = New System.Drawing.Size(161, 28)
        Me.lblNotice.TabIndex = 11
        Me.lblNotice.Text = "This item doesn't influence BPM/STOP object"
        '
        'lblNum
        '
        Me.lblNum.AutoSize = True
        Me.lblNum.BackColor = System.Drawing.SystemColors.Control
        Me.lblNum.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNum.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNum.Location = New System.Drawing.Point(72, 20)
        Me.lblNum.Name = "lblNum"
        Me.lblNum.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNum.Size = New System.Drawing.Size(17, 12)
        Me.lblNum.TabIndex = 9
        Me.lblNum.Text = "～"
        Me.lblNum.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'fraSearchMeasure
        '
        Me.fraSearchMeasure.BackColor = System.Drawing.SystemColors.Control
        Me.fraSearchMeasure.Controls.Add(Me.txtMeasureMax)
        Me.fraSearchMeasure.Controls.Add(Me.txtMeasureMin)
        Me.fraSearchMeasure.Controls.Add(Me.lblMeasure)
        Me.fraSearchMeasure.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraSearchMeasure.Location = New System.Drawing.Point(8, 68)
        Me.fraSearchMeasure.Name = "fraSearchMeasure"
        Me.fraSearchMeasure.Padding = New System.Windows.Forms.Padding(0)
        Me.fraSearchMeasure.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraSearchMeasure.Size = New System.Drawing.Size(177, 41)
        Me.fraSearchMeasure.TabIndex = 3
        Me.fraSearchMeasure.TabStop = False
        Me.fraSearchMeasure.Text = "小節範囲の指定"
        '
        'txtMeasureMax
        '
        Me.txtMeasureMax.AcceptsReturn = True
        Me.txtMeasureMax.BackColor = System.Drawing.SystemColors.Window
        Me.txtMeasureMax.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMeasureMax.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMeasureMax.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtMeasureMax.Location = New System.Drawing.Point(92, 16)
        Me.txtMeasureMax.MaxLength = 3
        Me.txtMeasureMax.Name = "txtMeasureMax"
        Me.txtMeasureMax.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMeasureMax.Size = New System.Drawing.Size(37, 19)
        Me.txtMeasureMax.TabIndex = 6
        Me.txtMeasureMax.Text = "999"
        '
        'txtMeasureMin
        '
        Me.txtMeasureMin.AcceptsReturn = True
        Me.txtMeasureMin.BackColor = System.Drawing.SystemColors.Window
        Me.txtMeasureMin.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMeasureMin.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMeasureMin.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtMeasureMin.Location = New System.Drawing.Point(28, 16)
        Me.txtMeasureMin.MaxLength = 3
        Me.txtMeasureMin.Name = "txtMeasureMin"
        Me.txtMeasureMin.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMeasureMin.Size = New System.Drawing.Size(37, 19)
        Me.txtMeasureMin.TabIndex = 4
        Me.txtMeasureMin.Text = "0"
        '
        'lblMeasure
        '
        Me.lblMeasure.AutoSize = True
        Me.lblMeasure.BackColor = System.Drawing.SystemColors.Control
        Me.lblMeasure.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMeasure.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMeasure.Location = New System.Drawing.Point(72, 20)
        Me.lblMeasure.Name = "lblMeasure"
        Me.lblMeasure.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMeasure.Size = New System.Drawing.Size(17, 12)
        Me.lblMeasure.TabIndex = 5
        Me.lblMeasure.Text = "～"
        Me.lblMeasure.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'fraSearchObject
        '
        Me.fraSearchObject.BackColor = System.Drawing.SystemColors.Control
        Me.fraSearchObject.Controls.Add(Me.optSearchSelect)
        Me.fraSearchObject.Controls.Add(Me.optSearchAll)
        Me.fraSearchObject.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraSearchObject.Location = New System.Drawing.Point(8, 4)
        Me.fraSearchObject.Name = "fraSearchObject"
        Me.fraSearchObject.Padding = New System.Windows.Forms.Padding(0)
        Me.fraSearchObject.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraSearchObject.Size = New System.Drawing.Size(177, 61)
        Me.fraSearchObject.TabIndex = 0
        Me.fraSearchObject.TabStop = False
        Me.fraSearchObject.Text = "選択対象"
        '
        'optSearchSelect
        '
        Me.optSearchSelect.BackColor = System.Drawing.SystemColors.Control
        Me.optSearchSelect.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSearchSelect.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSearchSelect.Location = New System.Drawing.Point(8, 36)
        Me.optSearchSelect.Name = "optSearchSelect"
        Me.optSearchSelect.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSearchSelect.Size = New System.Drawing.Size(161, 17)
        Me.optSearchSelect.TabIndex = 2
        Me.optSearchSelect.TabStop = True
        Me.optSearchSelect.Text = "選択されているオブジェ"
        Me.optSearchSelect.UseVisualStyleBackColor = False
        '
        'optSearchAll
        '
        Me.optSearchAll.BackColor = System.Drawing.SystemColors.Control
        Me.optSearchAll.Checked = True
        Me.optSearchAll.Cursor = System.Windows.Forms.Cursors.Default
        Me.optSearchAll.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optSearchAll.Location = New System.Drawing.Point(8, 16)
        Me.optSearchAll.Name = "optSearchAll"
        Me.optSearchAll.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optSearchAll.Size = New System.Drawing.Size(161, 17)
        Me.optSearchAll.TabIndex = 1
        Me.optSearchAll.TabStop = True
        Me.optSearchAll.Text = "全てのオブジェ"
        Me.optSearchAll.UseVisualStyleBackColor = False
        '
        'frmWindowFind
        '
        Me.AcceptButton = Me.cmdDecide
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(566, 193)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdDecide)
        Me.Controls.Add(Me.fraProcess)
        Me.Controls.Add(Me.fraSearchGrid)
        Me.Controls.Add(Me.fraSearchNum)
        Me.Controls.Add(Me.fraSearchMeasure)
        Me.Controls.Add(Me.fraSearchObject)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Location = New System.Drawing.Point(3, 19)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmWindowFind"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.Text = "検索・削除・置換"
        Me.fraProcess.ResumeLayout(False)
        Me.fraProcess.PerformLayout()
        Me.fraSearchGrid.ResumeLayout(False)
        Me.fraSearchGrid.PerformLayout()
        Me.fraSearchNum.ResumeLayout(False)
        Me.fraSearchNum.PerformLayout()
        Me.fraSearchMeasure.ResumeLayout(False)
        Me.fraSearchMeasure.PerformLayout()
        Me.fraSearchObject.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class