Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.IO
Imports System.Text

Friend Class frmWindowViewer
	Inherits System.Windows.Forms.Form

	Dim m_LocalViewer() As g_udtViewer
	Dim m_lngViewerNum As Integer

	Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click

		If Len(Trim(txtViewerName.Text)) = 0 Then Exit Sub
		If Len(Trim(txtViewerPath.Text)) = 0 Then Exit Sub

		ReDim Preserve m_LocalViewer(UBound(m_LocalViewer) + 1)

		With m_LocalViewer(UBound(m_LocalViewer))

			.strAppName = txtViewerName.Text
			.strAppPath = txtViewerPath.Text
			.strArgAll = txtPlayAll.Text
			.strArgPlay = txtPlay.Text
			.strArgStop = txtStop.Text

			Call lstViewer.Items.Add(.strAppName)
			lstViewer.SelectedIndex = UBound(m_LocalViewer) - 1

		End With

	End Sub

	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click

		Call Me.Close()

	End Sub

	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click

		Dim i As Integer
		Dim lngTemp As Integer

		ReDim g_Viewer(UBound(m_LocalViewer))

		If lstViewer.SelectedIndex <> -1 Then

			With m_LocalViewer(lstViewer.SelectedIndex + 1)

				.strAppName = txtViewerName.Text
				.strAppPath = txtViewerPath.Text
				.strArgAll = txtPlayAll.Text
				.strArgPlay = txtPlay.Text
				.strArgStop = txtStop.Text

			End With

			m_lngViewerNum = 0

		End If

		lngTemp = frmMain.cboViewer.SelectedIndex
		If lngTemp < 0 Then lngTemp = 0

		Call frmMain.cboViewer.Items.Clear()

		Using writer As New StreamWriter(g_strAppDir & "bmse_viewer.ini", False, Encoding.UTF8, 64)

			For i = 1 To UBound(m_LocalViewer)

				With m_LocalViewer(i)

					writer.WriteLine(.strAppName)
					writer.WriteLine(.strAppPath)
					writer.WriteLine(.strArgAll)
					writer.WriteLine(.strArgPlay)
					writer.WriteLine(.strArgStop)
					writer.WriteLine("")

					Call frmMain.cboViewer.Items.Add(.strAppName)

					g_Viewer(i).strAppName = .strAppName
					g_Viewer(i).strAppPath = .strAppPath
					g_Viewer(i).strArgAll = .strArgAll
					g_Viewer(i).strArgPlay = .strArgPlay
					g_Viewer(i).strArgStop = .strArgStop

				End With

			Next i

		End Using

		With frmMain

			If .cboViewer.Items.Count = 0 Then

				.tlbMenu.Items.Item("PlayAll").Enabled = False
				.tlbMenu.Items.Item("Play").Enabled = False
				.tlbMenu.Items.Item("_Stop").Enabled = False
				.mnuToolsPlayAll.Enabled = False
				.mnuToolsPlay.Enabled = False
				.mnuToolsPlayStop.Enabled = False
				.cboViewer.Enabled = False

			Else

				.tlbMenu.Items.Item("PlayAll").Enabled = True
				.tlbMenu.Items.Item("Play").Enabled = True
				.tlbMenu.Items.Item("_Stop").Enabled = True
				.mnuToolsPlayAll.Enabled = True
				.mnuToolsPlay.Enabled = True
				.mnuToolsPlayStop.Enabled = True
				.cboViewer.Enabled = True

				If frmMain.cboViewer.Items.Count > lngTemp Then

					frmMain.cboViewer.SelectedIndex = lngTemp

				Else

					frmMain.cboViewer.SelectedIndex = 0

				End If

			End If

		End With

		Call Me.Close()

	End Sub

	Private Sub cmdDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDelete.Click

		With lstViewer

			If .SelectedIndex < 0 Then Exit Sub

			Call ViewerDelete(.SelectedIndex + 1)

			Call lstViewer.Items.RemoveAt(.Items.Count - 1)

			ReDim Preserve m_LocalViewer(UBound(m_LocalViewer) - 1)

		End With

	End Sub

	Private Sub ViewerDelete(ByVal Num As Integer)

		If Num < UBound(m_LocalViewer) Then

			With m_LocalViewer(Num + 1)

				m_LocalViewer(Num).strAppName = .strAppName
				m_LocalViewer(Num).strAppPath = .strAppPath
				m_LocalViewer(Num).strArgAll = .strArgAll
				m_LocalViewer(Num).strArgPlay = .strArgPlay
				m_LocalViewer(Num).strArgStop = .strArgStop

				modMain.SetItemString(lstViewer, Num - 1, modMain.GetItemString(lstViewer, Num))

			End With

			Call ViewerDelete(Num + 1)

		End If

	End Sub

	Private Sub cmdViewerPath_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdViewerPath.Click
		On Error GoTo Err_Renamed

		Dim strArray() As String

		With frmMain.dlgMainOpen

			'UPGRADE_WARNING: Filter に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
			.Filter = "EXE files (*.exe)|*.exe|All files (*.*)|*.*"
			.FileName = txtViewerPath.Text

			If .ShowDialog() <> DialogResult.OK Then
				Exit Sub
			End If

			txtViewerPath.Text = .FileName
			strArray = Split(.FileName, "\")
			frmMain.dlgMainOpen.InitialDirectory = VB.Left(.FileName, Len(.FileName) - Len(strArray(UBound(strArray))))
			frmMain.dlgMainSave.InitialDirectory = VB.Left(.FileName, Len(.FileName) - Len(strArray(UBound(strArray))))

		End With

		Exit Sub

Err_Renamed:

	End Sub

	'UPGRADE_WARNING: Form イベント frmWindowViewer.Activate には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
	Private Sub frmWindowViewer_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

		Dim i As Integer

		Me.Left = frmMain.Left + frmMain.Width \ 2 - Me.Width \ 2
		Me.Top = frmMain.Top + frmMain.Height \ 2 - Me.Height \ 2

		m_lngViewerNum = 0

		ReDim m_LocalViewer(UBound(g_Viewer))

		Call lstViewer.Items.Clear()

		For i = 1 To UBound(g_Viewer)

			With g_Viewer(i)

				Call lstViewer.Items.Add(.strAppName)
				m_LocalViewer(i).strAppName = .strAppName
				m_LocalViewer(i).strAppPath = .strAppPath
				m_LocalViewer(i).strArgAll = .strArgAll
				m_LocalViewer(i).strArgPlay = .strArgPlay
				m_LocalViewer(i).strArgStop = .strArgStop

			End With

		Next i

		If lstViewer.Items.Count Then

			With m_LocalViewer(1)

				txtViewerName.Text = .strAppName
				txtViewerPath.Text = .strAppPath
				txtPlayAll.Text = .strArgAll
				txtPlay.Text = .strArgPlay
				txtStop.Text = .strArgStop

			End With

			lstViewer.SelectedIndex = 0

		End If

		Call cmdOK.Focus()

	End Sub

	'UPGRADE_WARNING: イベント lstViewer.SelectedIndexChanged は、フォームが初期化されたときに発生します。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"' をクリックしてください。
	Private Sub lstViewer_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstViewer.SelectedIndexChanged

		With m_LocalViewer(lstViewer.SelectedIndex + 1)

			txtViewerName.Text = .strAppName
			txtViewerPath.Text = .strAppPath
			txtPlayAll.Text = .strArgAll
			txtPlay.Text = .strArgPlay
			txtStop.Text = .strArgStop

		End With

		m_lngViewerNum = lstViewer.SelectedIndex

	End Sub

	Private Sub lstViewer_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lstViewer.MouseDown

		Call lstViewer_SelectedIndexChanged(lstViewer, New System.EventArgs())

	End Sub

	Private Sub txtPlay_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPlay.Enter

		txtPlay.SelectionStart = 0
		txtPlay.SelectionLength = Len(txtPlay.Text)

	End Sub

	Private Sub txtPlayAll_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPlayAll.Enter

		txtPlayAll.SelectionStart = 0
		txtPlayAll.SelectionLength = Len(txtPlayAll.Text)

	End Sub

	Private Sub txtStop_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtStop.Enter

		txtStop.SelectionStart = 0
		txtStop.SelectionLength = Len(txtStop.Text)

	End Sub

	Private Sub txtViewerName_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtViewerName.Enter

		txtViewerName.SelectionStart = 0
		txtViewerName.SelectionLength = Len(txtViewerName.Text)

	End Sub

	Private Sub txtViewerPath_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtViewerPath.Enter

		txtViewerPath.SelectionStart = 0
		txtViewerPath.SelectionLength = Len(txtViewerPath.Text)

	End Sub

	Private Sub frmWindowViewer_FormClosed(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
		e.Cancel = True

		Erase m_LocalViewer

		Call Me.Hide()

		Call frmMain.picMain.Focus()

	End Sub

	Private Sub txtField_KeyUp(eventSender As Object, eventArgs As EventArgs) Handles txtViewerName.KeyUp, txtViewerPath.KeyUp, txtPlayAll.KeyUp, txtPlay.KeyUp, txtStop.KeyUp
		'KeyDownだと反映タイミングと内容が１動作分遅れる -> KeyUp時に実行することで対処

		With m_LocalViewer(m_lngViewerNum + 1)

			.strAppName = txtViewerName.Text
			.strAppPath = txtViewerPath.Text
			.strArgAll = txtPlayAll.Text
			.strArgPlay = txtPlay.Text
			.strArgStop = txtStop.Text

			Call modMain.SetItemString(lstViewer, m_lngViewerNum, txtViewerName.Text)

		End With

	End Sub

	Private Sub frmWindoViewer_DragEnter(sender As Object, eventArgs As DragEventArgs) Handles MyBase.DragEnter
		If (eventArgs.Data.GetDataPresent(DataFormats.FileDrop)) Then
			eventArgs.Effect = DragDropEffects.Copy
		Else
			eventArgs.Effect = DragDropEffects.None
		End If
	End Sub

	Private Sub frmWindoViewer_DragDrop(eventSender As Object, eventArgs As DragEventArgs) Handles MyBase.DragDrop

		Dim pathArray As String() = CType(eventArgs.Data.GetData(DataFormats.FileDrop), String())
		Dim filename As String
		Dim filepath As String

		For Each filepath In pathArray

			filename = System.IO.Path.GetFileNameWithoutExtension(filepath)

			ReDim Preserve m_LocalViewer(UBound(m_LocalViewer) + 1)

			With m_LocalViewer(UBound(m_LocalViewer))

				.strAppName = filename
				.strAppPath = filepath
				.strArgAll = ""
				.strArgPlay = ""
				.strArgStop = ""

				Call lstViewer.Items.Add(.strAppName)
				lstViewer.SelectedIndex = UBound(m_LocalViewer) - 1

			End With

		Next

	End Sub

	Private Sub cmdExcUp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExcUp.Click
		With lstViewer
			If .SelectedIndex > 0 Then

				SwapList(lstViewer, .SelectedIndex - 1, .SelectedIndex)
				SwapViewer(m_LocalViewer, .SelectedIndex + 1, .SelectedIndex)

				.SelectedIndex = .SelectedIndex - 1

			End If
		End With
	End Sub

	Private Sub cmdExcDown_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExcDown.Click
		With lstViewer
			If .SelectedIndex < .Items.Count - 1 Then

				SwapList(lstViewer, .SelectedIndex + 1, .SelectedIndex)

				Dim viewer As g_udtViewer
				viewer = m_LocalViewer(.SelectedIndex + 1)
				m_LocalViewer(.SelectedIndex + 1) = m_LocalViewer(.SelectedIndex + 2)
				m_LocalViewer(.SelectedIndex + 2) = viewer

				.SelectedIndex = .SelectedIndex + 1

			End If
		End With
	End Sub

	Private Sub SwapList(ByRef list As ListBox, ByVal i As Integer, ByVal j As Integer)
		Dim strTemp = modMain.GetItemString(list, i)
		modMain.SetItemString(lstViewer, i, modMain.GetItemString(lstViewer, j))
		modMain.SetItemString(lstViewer, j, strTemp)
	End Sub

	Private Sub SwapViewer(LocalViewer() As g_udtViewer, ByVal i As Integer, ByVal j As Integer)

		Dim viewer As g_udtViewer
		viewer = LocalViewer(i)
		LocalViewer(i) = LocalViewer(j)
		LocalViewer(j) = viewer

	End Sub

End Class