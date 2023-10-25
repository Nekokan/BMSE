Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmWindowPreview
	Inherits System.Windows.Forms.Form

    Private Const GWL_STYLE As Integer = (-16)
    Private Const GWL_EXSTYLE As Integer = -20

    Public Sub SetWindowSize()
		
		Dim rectTemp As RECT
		
		Call frmWindowPreview_Resize(Me, New System.EventArgs())
		
		With rectTemp
			
			.left_Renamed = 0
			.Top = 0
            .right_Renamed = 256 + fraBGAPara.Width
            .Bottom = 256
			
		End With
		
		With Me

            If IntPtr.Size = 4 Then
                Call AdjustWindowRectEx(rectTemp, GetWindowLong(.Handle, GWL_STYLE), False, GetWindowLong(.Handle, GWL_EXSTYLE))
            Else
                Call AdjustWindowRectEx(rectTemp, GetWindowLongPtr(.Handle, GWL_STYLE), False, GetWindowLongPtr(.Handle, GWL_EXSTYLE))
            End If

            Call .SetBounds(.Left, .Top, rectTemp.right_Renamed - rectTemp.left_Renamed, rectTemp.Bottom - rectTemp.Top)

        End With
		
	End Sub
	
	'UPGRADE_WARNING: イベント chkBGLine.CheckStateChanged は、フォームが初期化されたときに発生します。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"' をクリックしてください。
	Private Sub chkBGLine_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkBGLine.CheckStateChanged

        picPreview.Refresh()

    End Sub

    Private Sub cmdCopy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCopy.Click
        Dim strArray(6) As String

        strArray(0) = _txtBGAPara_0.Text
        If Len(strArray(0)) = 0 Then strArray(0) = "0"
        strArray(1) = _txtBGAPara_1.Text
        If Len(strArray(1)) = 0 Then strArray(1) = "0"
        strArray(2) = _txtBGAPara_2.Text
        If Len(strArray(2)) = 0 Then strArray(2) = "0"
        strArray(3) = _txtBGAPara_3.Text
        If Len(strArray(3)) = 0 Then strArray(3) = "0"
        strArray(4) = _txtBGAPara_4.Text
        If Len(strArray(4)) = 0 Then strArray(4) = "0"
        strArray(5) = _txtBGAPara_5.Text
        If Len(strArray(5)) = 0 Then strArray(5) = "0"
        strArray(6) = _txtBGAPara_6.Text
        If Len(strArray(6)) = 0 Then strArray(6) = "0"

        Call My.Computer.Clipboard.Clear()
		Call My.Computer.Clipboard.SetText(Join(strArray, " "))
		
	End Sub
	
	Private Sub cmdPreviewBack_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPreviewBack.Click
		
		Dim i As Integer
		
		With frmMain

            If ._optChangeBottom_2.Checked Then

                i = .lstBGA.SelectedIndex - 1

                Do While i = 0

                    If Len(modMain.GetItemString(.lstBGA, i)) > 8 Then

                        .lstBGA.SelectedIndex = i

                        Exit Sub

                    End If

                    i = i - 1

                Loop

            Else

                i = .lstBMP.SelectedIndex - 1

                Do While i >= 0

                    If Len(modMain.GetItemString(.lstBMP, i)) > 8 Then

                        .lstBMP.SelectedIndex = i

                        Exit Sub

                    End If

                    i = i - 1


                Loop

            End If

        End With

    End Sub

    Private Sub cmdPreviewEnd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPreviewEnd.Click

        With frmMain

            If ._optChangeBottom_2.Checked Then

                .lstBGA.SelectedIndex = .lstBGA.Items.Count - 1

            Else

                .lstBMP.SelectedIndex = .lstBMP.Items.Count - 1

            End If

        End With

    End Sub

    Private Sub cmdPreviewHome_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPreviewHome.Click

        With frmMain

            If ._optChangeBottom_2.Checked Then

                .lstBGA.SelectedIndex = 0

            Else

                .lstBMP.SelectedIndex = 0

            End If

        End With

    End Sub

    Private Sub cmdPreviewNext_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPreviewNext.Click

        Dim i As Integer

        With frmMain

            If ._optChangeBottom_2.Checked Then

                i = .lstBGA.SelectedIndex + 1

                Do While i < .lstBGA.Items.Count

                    If Len(modMain.GetItemString(.lstBGA, i)) > 8 Then

                        .lstBGA.SelectedIndex = i

                        Exit Sub

                    End If

                    i = i + 1

                Loop

            Else

                i = .lstBMP.SelectedIndex + 1

                Do While i < .lstBMP.Items.Count

                    If Len(modMain.GetItemString(.lstBMP, i)) > 8 Then

                        .lstBMP.SelectedIndex = i

                        Exit Sub

                    End If

                    i = i + 1
					
				Loop 
				
			End If
			
		End With
		
	End Sub

    Private Sub frmWindowPreview_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        With Me

            Call .picPreview.SetBounds(0, 0, 256, 256)
            Call .picBackBuffer.SetBounds(0, 0, 256, 256)

        End With

    End Sub

    'UPGRADE_WARNING: イベント frmWindowPreview.Resize は、フォームが初期化されたときに発生します。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"' をクリックしてください。
    Private Sub frmWindowPreview_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
        On Error Resume Next

        Dim lngTemp As Integer

        With Me

            lngTemp = 120

            ._lblBGAPara_0.Left = lngTemp
            ._lblBGAPara_1.Left = lngTemp
            ._lblBGAPara_3.Left = lngTemp
            ._lblBGAPara_5.Left = lngTemp
            .cmdCopy.Left = lngTemp
            .chkBGLine.Left = lngTemp
            .chkLock.Left = lngTemp

            .cmdPreviewHome.Left = lngTemp
            .cmdPreviewBack.Left = .cmdPreviewHome.Left + .cmdPreviewHome.Width + 4
            .cmdPreviewNext.Left = .cmdPreviewBack.Left + .cmdPreviewBack.Width + 4
            .cmdPreviewEnd.Left = .cmdPreviewNext.Left + .cmdPreviewNext.Width + 4
            .fraBGACmd.Width = .cmdPreviewEnd.Left + .cmdPreviewEnd.Width + 4
            .fraBGACmd.Height = .cmdPreviewEnd.Height

            lngTemp = lngTemp + ._lblBGAPara_0.Width + 4

            ._lblBGAPara_0.Left = lngTemp
            ._lblBGAPara_1.Left = lngTemp
            ._lblBGAPara_3.Left = lngTemp
            ._lblBGAPara_5.Left = lngTemp

            lngTemp = lngTemp + ._lblBGAPara_0.Width + 12

            ._lblBGAPara_2.Left = lngTemp
            ._lblBGAPara_4.Left = lngTemp
            ._lblBGAPara_6.Left = lngTemp

            lngTemp = lngTemp + ._lblBGAPara_0.Width + 4

            ._lblBGAPara_2.Left = lngTemp
            ._lblBGAPara_4.Left = lngTemp
            ._lblBGAPara_6.Left = lngTemp

            lngTemp = lngTemp + ._lblBGAPara_0.Width

            .chkBGLine.Width = lngTemp - 8
            .chkLock.Width = lngTemp - 8

            .fraBGAPara.Width = lngTemp + 4

            Call .picPreview.SetBounds(0, 0, .ClientRectangle.Width - fraBGAPara.Width, .ClientRectangle.Height)
            Call .fraBGAPara.SetBounds(.ClientRectangle.Width - fraBGAPara.Width, 4, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
            Call .fraBGACmd.SetBounds(.ClientRectangle.Width - fraBGAPara.Width, .ClientRectangle.Height - fraBGACmd.Height - 4, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)

            Call .picPreview.Refresh()

        End With

    End Sub

    Private Sub frmWindowPreview_FormClosed(ByVal eventSender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True

        Call Me.Hide()

        Call frmMain.picMain.Focus()
    End Sub

    Private Sub picPreview_DragEnter(sender As Object, e As DragEventArgs) Handles picPreview.DragEnter
        If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub picPreview_DragDrop(sender As Object, e As DragEventArgs) Handles picPreview.DragDrop
        On Error GoTo Err_Renamed

        Dim i As Integer
        Dim strTemp As String

        For i = 0 To CType(e.Data.GetData(DataFormats.FileDrop), String()).Length - 1

            'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
            If Dir(CType(e.Data.GetData(DataFormats.FileDrop), String())(i), FileAttribute.Normal) <> vbNullString Then

                strTemp = CType(e.Data.GetData(DataFormats.FileDrop), String())(i)

                Call frmMain.PreviewBMP(strTemp)

                Me.Text = VB.Right(Me.Text, Len(Me.Text) - 3)

            End If

        Next i

Err_Renamed:
    End Sub

    Public Sub picPreview_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles picPreview.Paint
        Dim i As Integer

        With picPreview
            Call eventArgs.Graphics.Clear(.BackColor)

            Dim hDC As IntPtr = eventArgs.Graphics.GetHdc()

            If picBackBuffer.Image IsNot Nothing Then
                Using picBackBuffer_BitMap As New Bitmap(picBackBuffer.Image)
                    Dim hBitMap As IntPtr = picBackBuffer_BitMap.GetHbitmap
                    Dim hMDC As IntPtr = CreateCompatibleDC(hDC)

                    SelectObject(hMDC, hBitMap)
                    Call BitBlt(hDC, (.ClientRectangle.Width \ 2 - 128) + Val(_txtBGAPara_5.Text) - Val(_txtBGAPara_1.Text), (.ClientRectangle.Height \ 2 - 128) + Val(_txtBGAPara_6.Text) - Val(_txtBGAPara_2.Text), picBackBuffer.ClientRectangle.Width, picBackBuffer.ClientRectangle.Height, hMDC, 0, 0, SRCCOPY)

                    Dim hPen As IntPtr
                    Dim hOldPen As IntPtr

                    If chkBGLine.CheckState Then

                        hPen = CreatePen(PS_SOLID, 1, 0)
                        hOldPen = SelectObject(hDC, hPen)

                        For i = 4 To .ClientRectangle.Height Step 8

                            Call MoveToEx(hDC, 0, i, 0)
                            Call LineTo(hDC, .ClientRectangle.Width, i)

                        Next i

                        SelectObject(hDC, hOldPen)
                        DeleteObject(hPen)

                    End If

                    hPen = CreatePen(PS_SOLID, 2, 0)
                    hOldPen = SelectObject(hDC, hPen)

                    Call Rectangle(hDC, .ClientRectangle.Width \ 2 - 129, .ClientRectangle.Height \ 2 - 129, .ClientRectangle.Width \ 2 + 130, .ClientRectangle.Height \ 2 + 130)

                    SelectObject(hDC, hOldPen)
                    DeleteObject(hPen)

                    Call BitBlt(hDC, (.ClientRectangle.Width \ 2 - 128) + Val(_txtBGAPara_5.Text), (.ClientRectangle.Height \ 2 - 128) + Val(_txtBGAPara_6.Text), lngNumField(Val(_txtBGAPara_3.Text) - Val(_txtBGAPara_1.Text), 0, 256), lngNumField(Val(_txtBGAPara_4.Text) - Val(_txtBGAPara_2.Text), 0, 256), hMDC, Val(_txtBGAPara_1.Text), Val(_txtBGAPara_2.Text), SRCCOPY)

                    DeleteDC(hMDC)
                    DeleteObject(hBitMap)
                End Using
            End If

            eventArgs.Graphics.ReleaseHdc()

        End With

    End Sub

    'UPGRADE_WARNING: イベント txtBGAPara.TextChanged は、フォームが初期化されたときに発生します。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"' をクリックしてください。
    Private Sub txtBGAPara_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _txtBGAPara_6.TextChanged, _txtBGAPara_5.TextChanged, _txtBGAPara_4.TextChanged, _txtBGAPara_3.TextChanged, _txtBGAPara_2.TextChanged, _txtBGAPara_1.TextChanged, _txtBGAPara_0.TextChanged
        Select Case DirectCast(eventSender, TextBox).Name
            Case _txtBGAPara_0.Name
                If Val(_txtBGAPara_0.Text) < 0 Then _txtBGAPara_0.Text = CStr(0)
            Case _txtBGAPara_1.Name
                If Val(_txtBGAPara_1.Text) < 0 Then _txtBGAPara_1.Text = CStr(0)
            Case _txtBGAPara_2.Name
                If Val(_txtBGAPara_2.Text) < 0 Then _txtBGAPara_2.Text = CStr(0)
            Case _txtBGAPara_3.Name
                If Val(_txtBGAPara_3.Text) < 0 Then _txtBGAPara_3.Text = CStr(0)
            Case _txtBGAPara_4.Name
                If Val(_txtBGAPara_4.Text) < 0 Then _txtBGAPara_4.Text = CStr(0)
            Case _txtBGAPara_5.Name
                If Val(_txtBGAPara_5.Text) < 0 Then _txtBGAPara_5.Text = CStr(0)
            Case _txtBGAPara_6.Name
                If Val(_txtBGAPara_6.Text) < 0 Then _txtBGAPara_6.Text = CStr(0)
        End Select

        picPreview.Refresh()

    End Sub

    Private Sub txtBGAPara_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _txtBGAPara_6.Enter, _txtBGAPara_5.Enter, _txtBGAPara_4.Enter, _txtBGAPara_3.Enter, _txtBGAPara_2.Enter, _txtBGAPara_1.Enter, _txtBGAPara_0.Enter

        Select Case DirectCast(eventSender, TextBox).Name
            Case _txtBGAPara_0.Name
                _txtBGAPara_0.SelectionStart = 0
                _txtBGAPara_0.SelectionLength = Len(_txtBGAPara_0.Text)
            Case _txtBGAPara_1.Name
                _txtBGAPara_1.SelectionStart = 0
                _txtBGAPara_1.SelectionLength = Len(_txtBGAPara_1.Text)
            Case _txtBGAPara_2.Name
                _txtBGAPara_2.SelectionStart = 0
                _txtBGAPara_2.SelectionLength = Len(_txtBGAPara_2.Text)
            Case _txtBGAPara_3.Name
                _txtBGAPara_3.SelectionStart = 0
                _txtBGAPara_3.SelectionLength = Len(_txtBGAPara_3.Text)
            Case _txtBGAPara_4.Name
                _txtBGAPara_4.SelectionStart = 0
                _txtBGAPara_4.SelectionLength = Len(_txtBGAPara_4.Text)
            Case _txtBGAPara_5.Name
                _txtBGAPara_5.SelectionStart = 0
                _txtBGAPara_5.SelectionLength = Len(_txtBGAPara_5.Text)
            Case _txtBGAPara_6.Name
                _txtBGAPara_6.SelectionStart = 0
                _txtBGAPara_6.SelectionLength = Len(_txtBGAPara_6.Text)
        End Select

    End Sub

    Private Function lngNumField(ByVal lngNum As Integer, ByVal lngMin As Integer, ByVal lngMax As Integer) As Integer

        If lngNum < lngMin Then lngNum = lngMin

        If lngNum > lngMax Then lngNum = lngMax

        lngNumField = lngNum

    End Function

End Class