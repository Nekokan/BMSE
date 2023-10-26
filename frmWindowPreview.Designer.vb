<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmWindowPreview
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
    Public WithEvents cmdPreviewBack As System.Windows.Forms.Button
    Public WithEvents cmdPreviewNext As System.Windows.Forms.Button
	Public WithEvents cmdPreviewEnd As System.Windows.Forms.Button
	Public WithEvents cmdPreviewHome As System.Windows.Forms.Button
    Public WithEvents fraBGACmd As System.Windows.Forms.Panel
    Public WithEvents cmdCopy As System.Windows.Forms.Button
	Public WithEvents chkLock As System.Windows.Forms.CheckBox
	Public WithEvents chkBGLine As System.Windows.Forms.CheckBox
	Public WithEvents _txtBGAPara_0 As System.Windows.Forms.TextBox
	Public WithEvents _txtBGAPara_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtBGAPara_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtBGAPara_3 As System.Windows.Forms.TextBox
	Public WithEvents _txtBGAPara_4 As System.Windows.Forms.TextBox
	Public WithEvents _txtBGAPara_5 As System.Windows.Forms.TextBox
	Public WithEvents _txtBGAPara_6 As System.Windows.Forms.TextBox
	Public WithEvents _lblBGAPara_0 As System.Windows.Forms.Label
	Public WithEvents _lblBGAPara_1 As System.Windows.Forms.Label
	Public WithEvents _lblBGAPara_2 As System.Windows.Forms.Label
	Public WithEvents _lblBGAPara_3 As System.Windows.Forms.Label
	Public WithEvents _lblBGAPara_4 As System.Windows.Forms.Label
	Public WithEvents _lblBGAPara_5 As System.Windows.Forms.Label
	Public WithEvents _lblBGAPara_6 As System.Windows.Forms.Label
    Public WithEvents fraBGAPara As System.Windows.Forms.Panel
    Public WithEvents picBackBuffer As System.Windows.Forms.PictureBox
	Public WithEvents picPreview As System.Windows.Forms.PictureBox
    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使って変更できます。
    'コード エディタを使用して、変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.fraBGACmd = New System.Windows.Forms.Panel()
        Me.cmdPreviewBack = New System.Windows.Forms.Button()
        Me.cmdPreviewNext = New System.Windows.Forms.Button()
        Me.cmdPreviewEnd = New System.Windows.Forms.Button()
        Me.cmdPreviewHome = New System.Windows.Forms.Button()
        Me.fraBGAPara = New System.Windows.Forms.Panel()
        Me.cmdCopy = New System.Windows.Forms.Button()
        Me.chkLock = New System.Windows.Forms.CheckBox()
        Me.chkBGLine = New System.Windows.Forms.CheckBox()
        Me._txtBGAPara_0 = New System.Windows.Forms.TextBox()
        Me._txtBGAPara_1 = New System.Windows.Forms.TextBox()
        Me._txtBGAPara_2 = New System.Windows.Forms.TextBox()
        Me._txtBGAPara_3 = New System.Windows.Forms.TextBox()
        Me._txtBGAPara_4 = New System.Windows.Forms.TextBox()
        Me._txtBGAPara_5 = New System.Windows.Forms.TextBox()
        Me._txtBGAPara_6 = New System.Windows.Forms.TextBox()
        Me._lblBGAPara_0 = New System.Windows.Forms.Label()
        Me._lblBGAPara_1 = New System.Windows.Forms.Label()
        Me._lblBGAPara_2 = New System.Windows.Forms.Label()
        Me._lblBGAPara_3 = New System.Windows.Forms.Label()
        Me._lblBGAPara_4 = New System.Windows.Forms.Label()
        Me._lblBGAPara_5 = New System.Windows.Forms.Label()
        Me._lblBGAPara_6 = New System.Windows.Forms.Label()
        Me.picBackBuffer = New System.Windows.Forms.PictureBox()
        Me.picPreview = New System.Windows.Forms.PictureBox()
        Me.fraBGACmd.SuspendLayout()
        Me.fraBGAPara.SuspendLayout()
        CType(Me.picBackBuffer, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picPreview, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'fraBGACmd
        '
        Me.fraBGACmd.BackColor = System.Drawing.SystemColors.Control
        Me.fraBGACmd.Controls.Add(Me.cmdPreviewBack)
        Me.fraBGACmd.Controls.Add(Me.cmdPreviewNext)
        Me.fraBGACmd.Controls.Add(Me.cmdPreviewEnd)
        Me.fraBGACmd.Controls.Add(Me.cmdPreviewHome)
        Me.fraBGACmd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraBGACmd.Location = New System.Drawing.Point(260, 234)
        Me.fraBGACmd.Name = "fraBGACmd"
        Me.fraBGACmd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraBGACmd.Size = New System.Drawing.Size(152, 21)
        Me.fraBGACmd.TabIndex = 20
        '
        'cmdPreviewBack
        '
        Me.cmdPreviewBack.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPreviewBack.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPreviewBack.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPreviewBack.Location = New System.Drawing.Point(32, 0)
        Me.cmdPreviewBack.Name = "cmdPreviewBack"
        Me.cmdPreviewBack.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPreviewBack.Size = New System.Drawing.Size(25, 21)
        Me.cmdPreviewBack.TabIndex = 22
        Me.cmdPreviewBack.Text = "<"
        Me.cmdPreviewBack.UseVisualStyleBackColor = False
        '
        'cmdPreviewNext
        '
        Me.cmdPreviewNext.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPreviewNext.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPreviewNext.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPreviewNext.Location = New System.Drawing.Point(64, 0)
        Me.cmdPreviewNext.Name = "cmdPreviewNext"
        Me.cmdPreviewNext.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPreviewNext.Size = New System.Drawing.Size(25, 21)
        Me.cmdPreviewNext.TabIndex = 23
        Me.cmdPreviewNext.Text = ">"
        Me.cmdPreviewNext.UseVisualStyleBackColor = False
        '
        'cmdPreviewEnd
        '
        Me.cmdPreviewEnd.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPreviewEnd.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPreviewEnd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPreviewEnd.Location = New System.Drawing.Point(96, 0)
        Me.cmdPreviewEnd.Name = "cmdPreviewEnd"
        Me.cmdPreviewEnd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPreviewEnd.Size = New System.Drawing.Size(25, 21)
        Me.cmdPreviewEnd.TabIndex = 24
        Me.cmdPreviewEnd.Text = "ZZ"
        Me.cmdPreviewEnd.UseVisualStyleBackColor = False
        '
        'cmdPreviewHome
        '
        Me.cmdPreviewHome.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPreviewHome.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPreviewHome.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPreviewHome.Location = New System.Drawing.Point(0, 0)
        Me.cmdPreviewHome.Name = "cmdPreviewHome"
        Me.cmdPreviewHome.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPreviewHome.Size = New System.Drawing.Size(25, 21)
        Me.cmdPreviewHome.TabIndex = 21
        Me.cmdPreviewHome.Text = "01"
        Me.cmdPreviewHome.UseVisualStyleBackColor = False
        '
        'fraBGAPara
        '
        Me.fraBGAPara.BackColor = System.Drawing.SystemColors.Control
        Me.fraBGAPara.Controls.Add(Me.cmdCopy)
        Me.fraBGAPara.Controls.Add(Me.chkLock)
        Me.fraBGAPara.Controls.Add(Me.chkBGLine)
        Me.fraBGAPara.Controls.Add(Me._txtBGAPara_0)
        Me.fraBGAPara.Controls.Add(Me._txtBGAPara_1)
        Me.fraBGAPara.Controls.Add(Me._txtBGAPara_2)
        Me.fraBGAPara.Controls.Add(Me._txtBGAPara_3)
        Me.fraBGAPara.Controls.Add(Me._txtBGAPara_4)
        Me.fraBGAPara.Controls.Add(Me._txtBGAPara_5)
        Me.fraBGAPara.Controls.Add(Me._txtBGAPara_6)
        Me.fraBGAPara.Controls.Add(Me._lblBGAPara_0)
        Me.fraBGAPara.Controls.Add(Me._lblBGAPara_1)
        Me.fraBGAPara.Controls.Add(Me._lblBGAPara_2)
        Me.fraBGAPara.Controls.Add(Me._lblBGAPara_3)
        Me.fraBGAPara.Controls.Add(Me._lblBGAPara_4)
        Me.fraBGAPara.Controls.Add(Me._lblBGAPara_5)
        Me.fraBGAPara.Controls.Add(Me._lblBGAPara_6)
        Me.fraBGAPara.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraBGAPara.Location = New System.Drawing.Point(260, 0)
        Me.fraBGAPara.Name = "fraBGAPara"
        Me.fraBGAPara.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraBGAPara.Size = New System.Drawing.Size(153, 157)
        Me.fraBGAPara.TabIndex = 2
        '
        'cmdCopy
        '
        Me.cmdCopy.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCopy.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCopy.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCopy.Location = New System.Drawing.Point(5, 90)
        Me.cmdCopy.Name = "cmdCopy"
        Me.cmdCopy.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCopy.Size = New System.Drawing.Size(67, 21)
        Me.cmdCopy.TabIndex = 17
        Me.cmdCopy.Text = "Copy"
        Me.cmdCopy.UseVisualStyleBackColor = False
        '
        'chkLock
        '
        Me.chkLock.BackColor = System.Drawing.SystemColors.Control
        Me.chkLock.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLock.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLock.Location = New System.Drawing.Point(4, 135)
        Me.chkLock.Name = "chkLock"
        Me.chkLock.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLock.Size = New System.Drawing.Size(141, 17)
        Me.chkLock.TabIndex = 19
        Me.chkLock.Text = "Lock"
        Me.chkLock.UseVisualStyleBackColor = False
        '
        'chkBGLine
        '
        Me.chkBGLine.BackColor = System.Drawing.SystemColors.Control
        Me.chkBGLine.Checked = True
        Me.chkBGLine.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkBGLine.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkBGLine.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkBGLine.Location = New System.Drawing.Point(4, 117)
        Me.chkBGLine.Name = "chkBGLine"
        Me.chkBGLine.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkBGLine.Size = New System.Drawing.Size(141, 17)
        Me.chkBGLine.TabIndex = 18
        Me.chkBGLine.Text = "BG-Line"
        Me.chkBGLine.UseVisualStyleBackColor = False
        '
        '_txtBGAPara_0
        '
        Me._txtBGAPara_0.AcceptsReturn = True
        Me._txtBGAPara_0.BackColor = System.Drawing.SystemColors.Window
        Me._txtBGAPara_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtBGAPara_0.Enabled = False
        Me._txtBGAPara_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtBGAPara_0.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me._txtBGAPara_0.Location = New System.Drawing.Point(110, 0)
        Me._txtBGAPara_0.MaxLength = 8
        Me._txtBGAPara_0.Name = "_txtBGAPara_0"
        Me._txtBGAPara_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtBGAPara_0.Size = New System.Drawing.Size(36, 19)
        Me._txtBGAPara_0.TabIndex = 4
        Me._txtBGAPara_0.Text = "01"
        '
        '_txtBGAPara_1
        '
        Me._txtBGAPara_1.AcceptsReturn = True
        Me._txtBGAPara_1.BackColor = System.Drawing.SystemColors.Window
        Me._txtBGAPara_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtBGAPara_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtBGAPara_1.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me._txtBGAPara_1.Location = New System.Drawing.Point(110, 24)
        Me._txtBGAPara_1.MaxLength = 8
        Me._txtBGAPara_1.Name = "_txtBGAPara_1"
        Me._txtBGAPara_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtBGAPara_1.Size = New System.Drawing.Size(36, 19)
        Me._txtBGAPara_1.TabIndex = 6
        Me._txtBGAPara_1.Text = "0"
        '
        '_txtBGAPara_2
        '
        Me._txtBGAPara_2.AcceptsReturn = True
        Me._txtBGAPara_2.BackColor = System.Drawing.SystemColors.Window
        Me._txtBGAPara_2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtBGAPara_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtBGAPara_2.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me._txtBGAPara_2.Location = New System.Drawing.Point(172, 24)
        Me._txtBGAPara_2.MaxLength = 8
        Me._txtBGAPara_2.Name = "_txtBGAPara_2"
        Me._txtBGAPara_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtBGAPara_2.Size = New System.Drawing.Size(36, 19)
        Me._txtBGAPara_2.TabIndex = 8
        Me._txtBGAPara_2.Text = "0"
        '
        '_txtBGAPara_3
        '
        Me._txtBGAPara_3.AcceptsReturn = True
        Me._txtBGAPara_3.BackColor = System.Drawing.SystemColors.Window
        Me._txtBGAPara_3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtBGAPara_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtBGAPara_3.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me._txtBGAPara_3.Location = New System.Drawing.Point(110, 44)
        Me._txtBGAPara_3.MaxLength = 8
        Me._txtBGAPara_3.Name = "_txtBGAPara_3"
        Me._txtBGAPara_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtBGAPara_3.Size = New System.Drawing.Size(36, 19)
        Me._txtBGAPara_3.TabIndex = 10
        Me._txtBGAPara_3.Text = "0"
        '
        '_txtBGAPara_4
        '
        Me._txtBGAPara_4.AcceptsReturn = True
        Me._txtBGAPara_4.BackColor = System.Drawing.SystemColors.Window
        Me._txtBGAPara_4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtBGAPara_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtBGAPara_4.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me._txtBGAPara_4.Location = New System.Drawing.Point(172, 44)
        Me._txtBGAPara_4.MaxLength = 8
        Me._txtBGAPara_4.Name = "_txtBGAPara_4"
        Me._txtBGAPara_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtBGAPara_4.Size = New System.Drawing.Size(36, 19)
        Me._txtBGAPara_4.TabIndex = 12
        Me._txtBGAPara_4.Text = "0"
        '
        '_txtBGAPara_5
        '
        Me._txtBGAPara_5.AcceptsReturn = True
        Me._txtBGAPara_5.BackColor = System.Drawing.SystemColors.Window
        Me._txtBGAPara_5.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtBGAPara_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtBGAPara_5.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me._txtBGAPara_5.Location = New System.Drawing.Point(110, 65)
        Me._txtBGAPara_5.MaxLength = 8
        Me._txtBGAPara_5.Name = "_txtBGAPara_5"
        Me._txtBGAPara_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtBGAPara_5.Size = New System.Drawing.Size(36, 19)
        Me._txtBGAPara_5.TabIndex = 14
        Me._txtBGAPara_5.Text = "0"
        '
        '_txtBGAPara_6
        '
        Me._txtBGAPara_6.AcceptsReturn = True
        Me._txtBGAPara_6.BackColor = System.Drawing.SystemColors.Window
        Me._txtBGAPara_6.Cursor = System.Windows.Forms.Cursors.IBeam
        Me._txtBGAPara_6.ForeColor = System.Drawing.SystemColors.WindowText
        Me._txtBGAPara_6.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me._txtBGAPara_6.Location = New System.Drawing.Point(172, 65)
        Me._txtBGAPara_6.MaxLength = 8
        Me._txtBGAPara_6.Name = "_txtBGAPara_6"
        Me._txtBGAPara_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._txtBGAPara_6.Size = New System.Drawing.Size(36, 19)
        Me._txtBGAPara_6.TabIndex = 16
        Me._txtBGAPara_6.Text = "0"
        '
        '_lblBGAPara_0
        '
        Me._lblBGAPara_0.AutoSize = True
        Me._lblBGAPara_0.BackColor = System.Drawing.SystemColors.Control
        Me._lblBGAPara_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblBGAPara_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblBGAPara_0.Location = New System.Drawing.Point(4, 4)
        Me._lblBGAPara_0.Name = "_lblBGAPara_0"
        Me._lblBGAPara_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblBGAPara_0.Size = New System.Drawing.Size(23, 12)
        Me._lblBGAPara_0.TabIndex = 3
        Me._lblBGAPara_0.Text = "Num"
        '
        '_lblBGAPara_1
        '
        Me._lblBGAPara_1.AutoSize = True
        Me._lblBGAPara_1.BackColor = System.Drawing.SystemColors.Control
        Me._lblBGAPara_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblBGAPara_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblBGAPara_1.Location = New System.Drawing.Point(4, 26)
        Me._lblBGAPara_1.Name = "_lblBGAPara_1"
        Me._lblBGAPara_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblBGAPara_1.Size = New System.Drawing.Size(17, 12)
        Me._lblBGAPara_1.TabIndex = 5
        Me._lblBGAPara_1.Text = "X1"
        '
        '_lblBGAPara_2
        '
        Me._lblBGAPara_2.AutoSize = True
        Me._lblBGAPara_2.BackColor = System.Drawing.SystemColors.Control
        Me._lblBGAPara_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblBGAPara_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblBGAPara_2.Location = New System.Drawing.Point(82, 26)
        Me._lblBGAPara_2.Name = "_lblBGAPara_2"
        Me._lblBGAPara_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblBGAPara_2.Size = New System.Drawing.Size(17, 12)
        Me._lblBGAPara_2.TabIndex = 7
        Me._lblBGAPara_2.Text = "Y1"
        '
        '_lblBGAPara_3
        '
        Me._lblBGAPara_3.AutoSize = True
        Me._lblBGAPara_3.BackColor = System.Drawing.SystemColors.Control
        Me._lblBGAPara_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblBGAPara_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblBGAPara_3.Location = New System.Drawing.Point(4, 48)
        Me._lblBGAPara_3.Name = "_lblBGAPara_3"
        Me._lblBGAPara_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblBGAPara_3.Size = New System.Drawing.Size(17, 12)
        Me._lblBGAPara_3.TabIndex = 9
        Me._lblBGAPara_3.Text = "X2"
        '
        '_lblBGAPara_4
        '
        Me._lblBGAPara_4.AutoSize = True
        Me._lblBGAPara_4.BackColor = System.Drawing.SystemColors.Control
        Me._lblBGAPara_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblBGAPara_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblBGAPara_4.Location = New System.Drawing.Point(82, 48)
        Me._lblBGAPara_4.Name = "_lblBGAPara_4"
        Me._lblBGAPara_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblBGAPara_4.Size = New System.Drawing.Size(17, 12)
        Me._lblBGAPara_4.TabIndex = 11
        Me._lblBGAPara_4.Text = "Y2"
        '
        '_lblBGAPara_5
        '
        Me._lblBGAPara_5.AutoSize = True
        Me._lblBGAPara_5.BackColor = System.Drawing.SystemColors.Control
        Me._lblBGAPara_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblBGAPara_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblBGAPara_5.Location = New System.Drawing.Point(4, 69)
        Me._lblBGAPara_5.Name = "_lblBGAPara_5"
        Me._lblBGAPara_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblBGAPara_5.Size = New System.Drawing.Size(17, 12)
        Me._lblBGAPara_5.TabIndex = 13
        Me._lblBGAPara_5.Text = "dX"
        '
        '_lblBGAPara_6
        '
        Me._lblBGAPara_6.AutoSize = True
        Me._lblBGAPara_6.BackColor = System.Drawing.SystemColors.Control
        Me._lblBGAPara_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblBGAPara_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me._lblBGAPara_6.Location = New System.Drawing.Point(82, 69)
        Me._lblBGAPara_6.Name = "_lblBGAPara_6"
        Me._lblBGAPara_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblBGAPara_6.Size = New System.Drawing.Size(17, 12)
        Me._lblBGAPara_6.TabIndex = 15
        Me._lblBGAPara_6.Text = "dY"
        '
        'picBackBuffer
        '
        Me.picBackBuffer.BackColor = System.Drawing.Color.Black
        Me.picBackBuffer.Cursor = System.Windows.Forms.Cursors.Default
        Me.picBackBuffer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.picBackBuffer.Location = New System.Drawing.Point(0, 0)
        Me.picBackBuffer.Name = "picBackBuffer"
        Me.picBackBuffer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.picBackBuffer.Size = New System.Drawing.Size(61, 36)
        Me.picBackBuffer.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.picBackBuffer.TabIndex = 0
        Me.picBackBuffer.TabStop = False
        Me.picBackBuffer.Visible = False
        '
        'picPreview
        '
        Me.picPreview.AllowDrop = True
        Me.picPreview.BackColor = System.Drawing.Color.Black
        Me.picPreview.Cursor = System.Windows.Forms.Cursors.Default
        Me.picPreview.ForeColor = System.Drawing.Color.Green
        Me.picPreview.Location = New System.Drawing.Point(0, 40)
        Me.picPreview.Name = "picPreview"
        Me.picPreview.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.picPreview.Size = New System.Drawing.Size(61, 37)
        Me.picPreview.TabIndex = 1
        Me.picPreview.TabStop = False
        '
        'frmWindowPreview
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(414, 256)
        Me.Controls.Add(Me.fraBGACmd)
        Me.Controls.Add(Me.fraBGAPara)
        Me.Controls.Add(Me.picBackBuffer)
        Me.Controls.Add(Me.picPreview)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Location = New System.Drawing.Point(4, 20)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmWindowPreview"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.Text = "Preview Window"
        Me.fraBGACmd.ResumeLayout(False)
        Me.fraBGAPara.ResumeLayout(False)
        Me.fraBGAPara.PerformLayout()
        CType(Me.picBackBuffer, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picPreview, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class