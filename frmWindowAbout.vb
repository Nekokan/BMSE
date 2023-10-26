Option Strict Off
Option Explicit On
Friend Class frmWindowAbout
	Inherits System.Windows.Forms.Form

    Private m_sngRaster() As Single
    Private m_lngCounter As Integer
    Public stringFont As Font

    Private Sub PrintText(ByVal hDC As IntPtr, ByRef Text_Renamed As String, ByVal X As Integer, ByVal Y As Integer)

        Dim intTemp As Integer

        intTemp = Text_Renamed.Length

        With Me
            Call SetTextColor(hDC, 0)
            Call TextOut(hDC, X - 1, Y - 1, Text_Renamed, intTemp)
            Call TextOut(hDC, X, Y - 1, Text_Renamed, intTemp)
            Call TextOut(hDC, X + 1, Y - 1, Text_Renamed, intTemp)
            Call TextOut(hDC, X - 1, Y, Text_Renamed, intTemp)
            Call TextOut(hDC, X + 1, Y, Text_Renamed, intTemp)
            Call TextOut(hDC, X - 1, Y + 1, Text_Renamed, intTemp)
            Call TextOut(hDC, X, Y + 1, Text_Renamed, intTemp)
            Call TextOut(hDC, X + 1, Y + 1, Text_Renamed, intTemp)
            Call SetTextColor(hDC, 16777215)
            Call TextOut(hDC, X, Y, Text_Renamed, intTemp)
        End With

    End Sub

    'UPGRADE_WARNING: Form イベント frmWindowAbout.Activate には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
    Private Sub frmWindowAbout_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		Dim i As Integer

        ReDim m_sngRaster(Me.ClientRectangle.Height - 1)

        For i = 0 To UBound(m_sngRaster)
			
			m_sngRaster(i) = 0
			
		Next i

        m_lngCounter = 0

        Refresh()
    End Sub
	
	Private Sub frmWindowAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Click
		
		Call Me.Close()
		
	End Sub
	
	'UPGRADE_WARNING: Form イベント frmWindowAbout.Deactivate には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
	Private Sub frmWindowAbout_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
		
		Call Me.Close()
		
	End Sub

    Private Sub frmWindowAbout_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown

        Select Case eventArgs.KeyCode

            Case System.Windows.Forms.Keys.M

                tmrMain.Enabled = True

        End Select

    End Sub

    Private Sub frmWindowAbout_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        With Me

            '2016/01/11 所: .width = .widthになる、結局何をやっているかわからない
            '.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(.Width) * VB6.TwipsPerPixelX) / 15)
            '.Height = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(.Height) * VB6.TwipsPerPixelY) / 15)
            .Text = "About: " & g_strAppTitle & " (" & RELEASEDATE & " ver.)"

        End With

        Dim rectTemp As RECT

        With rectTemp

            .left_Renamed = 0
            .Top = 0
            .right_Renamed = 500
            .Bottom = 215

        End With

        With Me

            If IntPtr.Size = 4 Then
                Call AdjustWindowRectEx(rectTemp, GetWindowLong(.Handle, GWL_STYLE), False, GetWindowLong(.Handle, GWL_EXSTYLE))
            Else
                Call AdjustWindowRectEx(rectTemp, GetWindowLongPtr(.Handle, GWL_STYLE), False, GetWindowLongPtr(.Handle, GWL_EXSTYLE))
            End If

            Call .SetBounds(.Left, .Top, rectTemp.right_Renamed - rectTemp.left_Renamed, rectTemp.Bottom - rectTemp.Top)

        End With

        'Activate
        Dim i As Integer

        ReDim m_sngRaster(Me.ClientRectangle.Height - 1)

        For i = 0 To UBound(m_sngRaster)

            m_sngRaster(i) = 0

        Next i

        m_lngCounter = 0

    End Sub

    Private Sub frmWindowAbout_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint

        Dim i As Integer
        Dim strTemp As String
        Dim lngTemp As Integer
        Dim sngTemp As Single

        With Me
            Call eventArgs.Graphics.Clear(.BackColor)

            Dim hDC As IntPtr = eventArgs.Graphics.GetHdc()

            sngTemp = m_lngCounter / 10
            If sngTemp > 8 Then sngTemp = 8

            Using picMain_BitMap As New Bitmap(picMain.BackgroundImage)
                Dim hBitMap As IntPtr = picMain_BitMap.GetHbitmap
                Dim hMDC As IntPtr = CreateCompatibleDC(hDC)
                SelectObject(hMDC, hBitMap)

                For i = 0 To .ClientRectangle.Height - 1

                    'm_sngRaster(i) = m_sngRaster(i) + Sin((i + m_lngCounter) * RAD * 8)
                    'm_sngRaster(i) = m_sngRaster(i) + g_sngSin(((i + m_lngCounter) * 8) And 255)
                    m_sngRaster(i) = g_sngSin(((i + m_lngCounter) * 8) And 255) * sngTemp
                    'm_sngRaster(i) = (m_sngRaster(i) + i) And .ScaleWidth

                    If tmrMain.Enabled Then lngTemp = m_sngRaster(i)

                    'Call StretchBlt(.hdc, lngTemp - .ScaleWidth, i, .ScaleWidth, 1, picMain.hdc, 0, i, .ScaleWidth, 1, SRCCOPY)
                    'Call StretchBlt(.hdc, lngTemp, i, .ScaleWidth, 1, picMain.hdc, 0, i, .ScaleWidth, 1, SRCCOPY)

                    Call BitBlt(hDC, lngTemp, i, .ClientRectangle.Width, 1, hMDC, 0, i, SRCCOPY)

                Next i

                DeleteDC(hMDC)
                DeleteObject(hBitMap)
            End Using

            'Call BitBlt(.hWnd, 0, 0, .ScaleWidth, .ScaleHeight, picMain.hWnd, 0, 0, SRCCOPY)

            lngTemp = 0

            Dim oldFont As Font = stringFont

            Dim newstyle As FontStyle
            If newstyle And FontStyle.Underline Then
                newstyle = newstyle Xor FontStyle.Underline
            End If
            .stringFont = New Font(.stringFont.FontFamily, 11, newstyle Or FontStyle.Bold, .stringFont.Unit, .stringFont.GdiCharSet, .stringFont.GdiVerticalFont)

            oldFont.Dispose()

            lngTemp = g_InputLog.GetBufferSize

            Select Case lngTemp

                Case Is < 1024

                    strTemp = lngTemp & " Byte"

                Case Is < 1048576

                    strTemp = System.Math.Round(lngTemp / 1024, 2) & " KB"

                Case Else

                    strTemp = System.Math.Round(lngTemp / 1048576, 2) & " MB"

            End Select

            Dim hFont As IntPtr = stringFont.ToHfont()
            Dim hOldFont As IntPtr = SelectObject(hDC, hFont)

            SetBkMode(hDC, TRANSPARENT)

            strTemp = "Undo Buffer Size: " & strTemp
            Call PrintText(hDC, strTemp, 1, 1)

            Call PrintText(hDC, "Undo Counter: " & g_InputLog.GetPos & " / " & g_InputLog.Max, 1, 18)

            SelectObject(hDC, hOldFont)
            DeleteObject(hFont)

            '.Font.SIZE = 12 * (Screen.TwipsPerPixelX / 15)
            '.Font.Underline = True

            'strTemp = App.Major & "." & App.Minor & "." & App.Revision
            'lngTemp = LenB(StrConv(strTemp, vbFromUnicode))

            'Call PrintText(strTemp, 251, 174)

            eventArgs.Graphics.ReleaseHdc()
        End With

    End Sub

    Private Sub picMain_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles picMain.Click

        Call Me.Close()

    End Sub

    Private Sub tmrMain_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrMain.Tick

        m_lngCounter = m_lngCounter + 1
        Refresh()

    End Sub

    Private Sub frmWindowAbout_FormClosed(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        e.Cancel = True

        tmrMain.Enabled = False

        Erase m_sngRaster

        Call Me.Hide()

        Call frmMain.picMain.Focus()
    End Sub
End Class