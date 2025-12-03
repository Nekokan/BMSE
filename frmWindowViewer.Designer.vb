<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmWindowViewer
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
    Public WithEvents cmdDelete As System.Windows.Forms.Button
    Public WithEvents cmdAdd As System.Windows.Forms.Button
	Public WithEvents txtStop As System.Windows.Forms.TextBox
	Public WithEvents txtPlay As System.Windows.Forms.TextBox
	Public WithEvents txtPlayAll As System.Windows.Forms.TextBox
	Public WithEvents cmdViewerPath As System.Windows.Forms.Button
	Public WithEvents txtViewerPath As System.Windows.Forms.TextBox
	Public WithEvents txtViewerName As System.Windows.Forms.TextBox
	Public WithEvents lblNotice As System.Windows.Forms.Label
	Public WithEvents lblStop As System.Windows.Forms.Label
	Public WithEvents lblPlay As System.Windows.Forms.Label
	Public WithEvents lblPlayAll As System.Windows.Forms.Label
	Public WithEvents lblViewerPath As System.Windows.Forms.Label
	Public WithEvents lblViewerName As System.Windows.Forms.Label
	Public WithEvents fraViewer As System.Windows.Forms.GroupBox
    Public WithEvents lstViewer As System.Windows.Forms.ListBox
    Public WithEvents cmdExcUp As System.Windows.Forms.Button
    Public WithEvents cmdExcDown As System.Windows.Forms.Button
    Public WithEvents cmdCancel As System.Windows.Forms.Button
    Public WithEvents cmdOK As System.Windows.Forms.Button
    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使って変更できます。
    'コード エディタを使用して、変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.fraViewer = New System.Windows.Forms.GroupBox()
        Me.txtStop = New System.Windows.Forms.TextBox()
        Me.txtPlay = New System.Windows.Forms.TextBox()
        Me.txtPlayAll = New System.Windows.Forms.TextBox()
        Me.cmdViewerPath = New System.Windows.Forms.Button()
        Me.txtViewerPath = New System.Windows.Forms.TextBox()
        Me.txtViewerName = New System.Windows.Forms.TextBox()
        Me.lblNotice = New System.Windows.Forms.Label()
        Me.lblStop = New System.Windows.Forms.Label()
        Me.lblPlay = New System.Windows.Forms.Label()
        Me.lblPlayAll = New System.Windows.Forms.Label()
        Me.lblViewerPath = New System.Windows.Forms.Label()
        Me.lblViewerName = New System.Windows.Forms.Label()
        Me.lstViewer = New System.Windows.Forms.ListBox()
        Me.cmdExcUp = New System.Windows.Forms.Button()
        Me.cmdExcDown = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOK = New System.Windows.Forms.Button()
        Me.fraViewer.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdDelete
        '
        Me.cmdDelete.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDelete.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDelete.Location = New System.Drawing.Point(36, 264)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDelete.Size = New System.Drawing.Size(53, 21)
        Me.cmdDelete.TabIndex = 1
        Me.cmdDelete.Text = "削除"
        Me.cmdDelete.UseVisualStyleBackColor = False
        '
        'cmdAdd
        '
        Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAdd.Location = New System.Drawing.Point(96, 264)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAdd.Size = New System.Drawing.Size(53, 21)
        Me.cmdAdd.TabIndex = 2
        Me.cmdAdd.Text = "追加"
        Me.cmdAdd.UseVisualStyleBackColor = False
        '
        'fraViewer
        '
        Me.fraViewer.BackColor = System.Drawing.SystemColors.Control
        Me.fraViewer.Controls.Add(Me.txtStop)
        Me.fraViewer.Controls.Add(Me.txtPlay)
        Me.fraViewer.Controls.Add(Me.txtPlayAll)
        Me.fraViewer.Controls.Add(Me.cmdViewerPath)
        Me.fraViewer.Controls.Add(Me.txtViewerPath)
        Me.fraViewer.Controls.Add(Me.txtViewerName)
        Me.fraViewer.Controls.Add(Me.lblNotice)
        Me.fraViewer.Controls.Add(Me.lblStop)
        Me.fraViewer.Controls.Add(Me.lblPlay)
        Me.fraViewer.Controls.Add(Me.lblPlayAll)
        Me.fraViewer.Controls.Add(Me.lblViewerPath)
        Me.fraViewer.Controls.Add(Me.lblViewerName)
        Me.fraViewer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraViewer.Location = New System.Drawing.Point(156, 4)
        Me.fraViewer.Name = "fraViewer"
        Me.fraViewer.Padding = New System.Windows.Forms.Padding(0)
        Me.fraViewer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraViewer.Size = New System.Drawing.Size(237, 281)
        Me.fraViewer.TabIndex = 3
        Me.fraViewer.TabStop = False
        '
        'txtStop
        '
        Me.txtStop.AcceptsReturn = True
        Me.txtStop.BackColor = System.Drawing.SystemColors.Window
        Me.txtStop.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStop.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStop.Location = New System.Drawing.Point(8, 204)
        Me.txtStop.MaxLength = 0
        Me.txtStop.Name = "txtStop"
        Me.txtStop.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStop.Size = New System.Drawing.Size(221, 19)
        Me.txtStop.TabIndex = 14
        '
        'txtPlay
        '
        Me.txtPlay.AcceptsReturn = True
        Me.txtPlay.BackColor = System.Drawing.SystemColors.Window
        Me.txtPlay.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPlay.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPlay.Location = New System.Drawing.Point(8, 160)
        Me.txtPlay.MaxLength = 0
        Me.txtPlay.Name = "txtPlay"
        Me.txtPlay.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPlay.Size = New System.Drawing.Size(221, 19)
        Me.txtPlay.TabIndex = 12
        '
        'txtPlayAll
        '
        Me.txtPlayAll.AcceptsReturn = True
        Me.txtPlayAll.BackColor = System.Drawing.SystemColors.Window
        Me.txtPlayAll.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPlayAll.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPlayAll.Location = New System.Drawing.Point(8, 116)
        Me.txtPlayAll.MaxLength = 0
        Me.txtPlayAll.Name = "txtPlayAll"
        Me.txtPlayAll.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPlayAll.Size = New System.Drawing.Size(221, 19)
        Me.txtPlayAll.TabIndex = 10
        '
        'cmdViewerPath
        '
        Me.cmdViewerPath.BackColor = System.Drawing.SystemColors.Control
        Me.cmdViewerPath.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdViewerPath.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdViewerPath.Location = New System.Drawing.Point(192, 72)
        Me.cmdViewerPath.Name = "cmdViewerPath"
        Me.cmdViewerPath.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdViewerPath.Size = New System.Drawing.Size(37, 17)
        Me.cmdViewerPath.TabIndex = 8
        Me.cmdViewerPath.Text = "参照"
        Me.cmdViewerPath.UseVisualStyleBackColor = False
        '
        'txtViewerPath
        '
        Me.txtViewerPath.AcceptsReturn = True
        Me.txtViewerPath.BackColor = System.Drawing.SystemColors.Window
        Me.txtViewerPath.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtViewerPath.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtViewerPath.Location = New System.Drawing.Point(8, 72)
        Me.txtViewerPath.MaxLength = 0
        Me.txtViewerPath.Name = "txtViewerPath"
        Me.txtViewerPath.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtViewerPath.Size = New System.Drawing.Size(181, 19)
        Me.txtViewerPath.TabIndex = 7
        '
        'txtViewerName
        '
        Me.txtViewerName.AcceptsReturn = True
        Me.txtViewerName.BackColor = System.Drawing.SystemColors.Window
        Me.txtViewerName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtViewerName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtViewerName.Location = New System.Drawing.Point(8, 28)
        Me.txtViewerName.MaxLength = 0
        Me.txtViewerName.Name = "txtViewerName"
        Me.txtViewerName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtViewerName.Size = New System.Drawing.Size(221, 19)
        Me.txtViewerName.TabIndex = 5
        '
        'lblNotice
        '
        Me.lblNotice.BackColor = System.Drawing.SystemColors.Control
        Me.lblNotice.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNotice.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNotice.Location = New System.Drawing.Point(8, 232)
        Me.lblNotice.Name = "lblNotice"
        Me.lblNotice.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNotice.Size = New System.Drawing.Size(221, 41)
        Me.lblNotice.TabIndex = 15
        Me.lblNotice.Text = "lblNotice"
        '
        'lblStop
        '
        Me.lblStop.BackColor = System.Drawing.SystemColors.Control
        Me.lblStop.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblStop.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblStop.Location = New System.Drawing.Point(12, 188)
        Me.lblStop.Name = "lblStop"
        Me.lblStop.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblStop.Size = New System.Drawing.Size(213, 13)
        Me.lblStop.TabIndex = 13
        Me.lblStop.Text = "「停止」の引数"
        '
        'lblPlay
        '
        Me.lblPlay.BackColor = System.Drawing.SystemColors.Control
        Me.lblPlay.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPlay.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPlay.Location = New System.Drawing.Point(12, 144)
        Me.lblPlay.Name = "lblPlay"
        Me.lblPlay.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPlay.Size = New System.Drawing.Size(213, 13)
        Me.lblPlay.TabIndex = 11
        Me.lblPlay.Text = "「現在位置から再生」の引数"
        '
        'lblPlayAll
        '
        Me.lblPlayAll.BackColor = System.Drawing.SystemColors.Control
        Me.lblPlayAll.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPlayAll.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPlayAll.Location = New System.Drawing.Point(12, 100)
        Me.lblPlayAll.Name = "lblPlayAll"
        Me.lblPlayAll.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPlayAll.Size = New System.Drawing.Size(213, 13)
        Me.lblPlayAll.TabIndex = 9
        Me.lblPlayAll.Text = "「最初から再生」の引数"
        '
        'lblViewerPath
        '
        Me.lblViewerPath.BackColor = System.Drawing.SystemColors.Control
        Me.lblViewerPath.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblViewerPath.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblViewerPath.Location = New System.Drawing.Point(12, 56)
        Me.lblViewerPath.Name = "lblViewerPath"
        Me.lblViewerPath.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblViewerPath.Size = New System.Drawing.Size(213, 13)
        Me.lblViewerPath.TabIndex = 6
        Me.lblViewerPath.Text = "実行ファイルのパス"
        '
        'lblViewerName
        '
        Me.lblViewerName.BackColor = System.Drawing.SystemColors.Control
        Me.lblViewerName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblViewerName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblViewerName.Location = New System.Drawing.Point(12, 12)
        Me.lblViewerName.Name = "lblViewerName"
        Me.lblViewerName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblViewerName.Size = New System.Drawing.Size(213, 13)
        Me.lblViewerName.TabIndex = 4
        Me.lblViewerName.Text = "表示する名前"
        '
        'lstViewer
        '
        Me.lstViewer.BackColor = System.Drawing.SystemColors.Window
        Me.lstViewer.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstViewer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstViewer.ItemHeight = 12
        Me.lstViewer.Location = New System.Drawing.Point(8, 12)
        Me.lstViewer.Name = "lstViewer"
        Me.lstViewer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstViewer.Size = New System.Drawing.Size(141, 244)
        Me.lstViewer.TabIndex = 0
        '
        'cmdExcUp
        '
        Me.cmdExcUp.BackColor = System.Drawing.SystemColors.Control
        Me.cmdExcUp.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdExcUp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdExcUp.Location = New System.Drawing.Point(8, 292)
        Me.cmdExcUp.Name = "cmdExcUp"
        Me.cmdExcUp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdExcUp.Size = New System.Drawing.Size(21, 21)
        Me.cmdExcUp.TabIndex = 18
        Me.cmdExcUp.Text = "▲"
        Me.cmdExcUp.UseVisualStyleBackColor = False
        '
        'cmdExcDown
        '
        Me.cmdExcDown.BackColor = System.Drawing.SystemColors.Control
        Me.cmdExcDown.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdExcDown.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdExcDown.Location = New System.Drawing.Point(33, 292)
        Me.cmdExcDown.Name = "cmdExcDown"
        Me.cmdExcDown.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdExcDown.Size = New System.Drawing.Size(21, 21)
        Me.cmdExcDown.TabIndex = 19
        Me.cmdExcDown.Text = "▼"
        Me.cmdExcDown.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(310, 292)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(83, 25)
        Me.cmdCancel.TabIndex = 17
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(206, 292)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(97, 25)
        Me.cmdOK.TabIndex = 16
        Me.cmdOK.Text = "OK"
        Me.cmdOK.UseVisualStyleBackColor = False
        '
        'frmWindowViewer
        '
        Me.AllowDrop = True
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(400, 323)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.fraViewer)
        Me.Controls.Add(Me.lstViewer)
        Me.Controls.Add(Me.cmdExcUp)
        Me.Controls.Add(Me.cmdExcDown)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Location = New System.Drawing.Point(3, 19)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmWindowViewer"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.Text = "Viewer Config"
        Me.fraViewer.ResumeLayout(False)
        Me.fraViewer.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
#End Region
End Class