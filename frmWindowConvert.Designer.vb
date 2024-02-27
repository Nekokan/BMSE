<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmWindowConvert
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
    Public WithEvents chkSortByName As System.Windows.Forms.CheckBox
    Public WithEvents chkFileRecycle As System.Windows.Forms.CheckBox
	Public WithEvents chkDeleteFile As System.Windows.Forms.CheckBox
	Public WithEvents txtExtension As System.Windows.Forms.TextBox
	Public WithEvents chkFileNameConvert As System.Windows.Forms.CheckBox
	Public WithEvents chkUseOldFormat As System.Windows.Forms.CheckBox
	Public WithEvents chkListAlign As System.Windows.Forms.CheckBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdDecide As System.Windows.Forms.Button
	Public WithEvents chkDeleteUnusedFile As System.Windows.Forms.CheckBox
	Public WithEvents lblNotice As System.Windows.Forms.Label
	Public WithEvents lblExtension As System.Windows.Forms.Label
	'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.chkSortByName = New System.Windows.Forms.CheckBox()
        Me.chkFileRecycle = New System.Windows.Forms.CheckBox()
        Me.chkDeleteFile = New System.Windows.Forms.CheckBox()
        Me.txtExtension = New System.Windows.Forms.TextBox()
        Me.chkFileNameConvert = New System.Windows.Forms.CheckBox()
        Me.chkUseOldFormat = New System.Windows.Forms.CheckBox()
        Me.chkListAlign = New System.Windows.Forms.CheckBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdDecide = New System.Windows.Forms.Button()
        Me.chkDeleteUnusedFile = New System.Windows.Forms.CheckBox()
        Me.lblNotice = New System.Windows.Forms.Label()
        Me.lblExtension = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'chkSortByName
        '
        Me.chkSortByName.BackColor = System.Drawing.SystemColors.Control
        Me.chkSortByName.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSortByName.Enabled = False
        Me.chkSortByName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSortByName.Location = New System.Drawing.Point(24, 194)
        Me.chkSortByName.Name = "chkSortByName"
        Me.chkSortByName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSortByName.Size = New System.Drawing.Size(289, 25)
        Me.chkSortByName.TabIndex = 11
        Me.chkSortByName.Text = "ファイル名順でソートする"
        Me.chkSortByName.UseVisualStyleBackColor = False
        '
        'chkFileRecycle
        '
        Me.chkFileRecycle.BackColor = System.Drawing.SystemColors.Control
        Me.chkFileRecycle.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkFileRecycle.Enabled = False
        Me.chkFileRecycle.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkFileRecycle.Location = New System.Drawing.Point(24, 96)
        Me.chkFileRecycle.Name = "chkFileRecycle"
        Me.chkFileRecycle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkFileRecycle.Size = New System.Drawing.Size(289, 25)
        Me.chkFileRecycle.TabIndex = 6
        Me.chkFileRecycle.Text = "ごみ箱に移動しないですぐに削除する"
        Me.chkFileRecycle.UseVisualStyleBackColor = False
        '
        'chkDeleteFile
        '
        Me.chkDeleteFile.BackColor = System.Drawing.SystemColors.Control
        Me.chkDeleteFile.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDeleteFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDeleteFile.Location = New System.Drawing.Point(8, 48)
        Me.chkDeleteFile.Name = "chkDeleteFile"
        Me.chkDeleteFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDeleteFile.Size = New System.Drawing.Size(297, 25)
        Me.chkDeleteFile.TabIndex = 3
        Me.chkDeleteFile.Text = "フォルダ内の使用していないファイルを削除 (*)"
        Me.chkDeleteFile.UseVisualStyleBackColor = False
        '
        'txtExtension
        '
        Me.txtExtension.AcceptsReturn = True
        Me.txtExtension.BackColor = System.Drawing.SystemColors.Window
        Me.txtExtension.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtExtension.Enabled = False
        Me.txtExtension.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtExtension.Location = New System.Drawing.Point(112, 76)
        Me.txtExtension.MaxLength = 0
        Me.txtExtension.Name = "txtExtension"
        Me.txtExtension.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtExtension.Size = New System.Drawing.Size(193, 19)
        Me.txtExtension.TabIndex = 5
        Me.txtExtension.Text = "wav,mp3,bmp,jpg,gif"
        '
        'chkFileNameConvert
        '
        Me.chkFileNameConvert.BackColor = System.Drawing.SystemColors.Control
        Me.chkFileNameConvert.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkFileNameConvert.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkFileNameConvert.Location = New System.Drawing.Point(8, 224)
        Me.chkFileNameConvert.Name = "chkFileNameConvert"
        Me.chkFileNameConvert.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkFileNameConvert.Size = New System.Drawing.Size(297, 25)
        Me.chkFileNameConvert.TabIndex = 9
        Me.chkFileNameConvert.Text = "ファイル名を連番 (01 - ZZ) に変換 (*)"
        Me.chkFileNameConvert.UseVisualStyleBackColor = False
        '
        'chkUseOldFormat
        '
        Me.chkUseOldFormat.BackColor = System.Drawing.SystemColors.Control
        Me.chkUseOldFormat.Checked = True
        Me.chkUseOldFormat.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkUseOldFormat.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkUseOldFormat.Enabled = False
        Me.chkUseOldFormat.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkUseOldFormat.Location = New System.Drawing.Point(24, 160)
        Me.chkUseOldFormat.Name = "chkUseOldFormat"
        Me.chkUseOldFormat.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkUseOldFormat.Size = New System.Drawing.Size(289, 35)
        Me.chkUseOldFormat.TabIndex = 8
        Me.chkUseOldFormat.Text = "可能な限り古いフォーマット (01 - zz -> ZZ -> FF) を使う"
        Me.chkUseOldFormat.UseVisualStyleBackColor = False
        '
        'chkListAlign
        '
        Me.chkListAlign.BackColor = System.Drawing.SystemColors.Control
        Me.chkListAlign.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkListAlign.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkListAlign.Location = New System.Drawing.Point(8, 136)
        Me.chkListAlign.Name = "chkListAlign"
        Me.chkListAlign.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkListAlign.Size = New System.Drawing.Size(297, 25)
        Me.chkListAlign.TabIndex = 7
        Me.chkListAlign.Text = "定義リストの整列"
        Me.chkListAlign.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(216, 296)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(89, 25)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "キャンセル"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdDecide
        '
        Me.cmdDecide.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDecide.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDecide.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDecide.Location = New System.Drawing.Point(88, 296)
        Me.cmdDecide.Name = "cmdDecide"
        Me.cmdDecide.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDecide.Size = New System.Drawing.Size(121, 25)
        Me.cmdDecide.TabIndex = 0
        Me.cmdDecide.Text = "実行"
        Me.cmdDecide.UseVisualStyleBackColor = False
        '
        'chkDeleteUnusedFile
        '
        Me.chkDeleteUnusedFile.BackColor = System.Drawing.SystemColors.Control
        Me.chkDeleteUnusedFile.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDeleteUnusedFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkDeleteUnusedFile.Location = New System.Drawing.Point(8, 16)
        Me.chkDeleteUnusedFile.Name = "chkDeleteUnusedFile"
        Me.chkDeleteUnusedFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDeleteUnusedFile.Size = New System.Drawing.Size(297, 17)
        Me.chkDeleteUnusedFile.TabIndex = 2
        Me.chkDeleteUnusedFile.Text = "使用していない #WAV・#BMP・#BGA の定義を消去"
        Me.chkDeleteUnusedFile.UseVisualStyleBackColor = False
        '
        'lblNotice
        '
        Me.lblNotice.BackColor = System.Drawing.SystemColors.Control
        Me.lblNotice.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNotice.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblNotice.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNotice.Location = New System.Drawing.Point(8, 264)
        Me.lblNotice.Name = "lblNotice"
        Me.lblNotice.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNotice.Size = New System.Drawing.Size(297, 17)
        Me.lblNotice.TabIndex = 10
        Me.lblNotice.Text = "(*)・・・この操作はやり直しができません"
        '
        'lblExtension
        '
        Me.lblExtension.AutoSize = True
        Me.lblExtension.BackColor = System.Drawing.SystemColors.Control
        Me.lblExtension.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblExtension.Enabled = False
        Me.lblExtension.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblExtension.Location = New System.Drawing.Point(26, 80)
        Me.lblExtension.Name = "lblExtension"
        Me.lblExtension.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblExtension.Size = New System.Drawing.Size(86, 12)
        Me.lblExtension.TabIndex = 4
        Me.lblExtension.Text = "検索する拡張子:"
        '
        'frmWindowConvert
        '
        Me.AcceptButton = Me.cmdDecide
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(312, 330)
        Me.Controls.Add(Me.chkSortByName)
        Me.Controls.Add(Me.chkFileRecycle)
        Me.Controls.Add(Me.chkDeleteFile)
        Me.Controls.Add(Me.txtExtension)
        Me.Controls.Add(Me.chkFileNameConvert)
        Me.Controls.Add(Me.chkUseOldFormat)
        Me.Controls.Add(Me.chkListAlign)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdDecide)
        Me.Controls.Add(Me.chkDeleteUnusedFile)
        Me.Controls.Add(Me.lblNotice)
        Me.Controls.Add(Me.lblExtension)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Location = New System.Drawing.Point(3, 19)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmWindowConvert"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.Text = "変換ウィザード"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
End Class