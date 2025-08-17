Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmWindowFind
	Inherits System.Windows.Forms.Form
	
	Private m_strArray() As String
	
	Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
		
		Call Me.Close()
		
	End Sub
	
	Private Sub cmdDecide_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDecide.Click
        Dim i As Integer

        ReDim m_strArray(0)
		m_strArray(0) = ""
		
		For i = 0 To UBound(g_Obj) - 1
			
			With g_Obj(i)

                If (optSearchAll.Checked = True Or .intSelect = modMain.OBJ_SELECT.Selected) And (Val(txtMeasureMin.Text) <= .intMeasure And Val(txtMeasureMax.Text) >= .intMeasure) And ((modInput.strToNum((txtNumMin.Text)) <= .sngValue And modInput.strToNum((txtNumMax.Text)) >= .sngValue) Or .intCh = 8 Or .intCh = 9) Then

                    .intSelect = modMain.OBJ_SELECT.NON_SELECT

                    Select Case .intCh

                        Case OBJ_CH.CH_EXBPM 'BPM

                            If _lstGrid_3.GetItemChecked(0) Then Call SearchProcess(i)

                        Case OBJ_CH.CH_STOP 'STOP

                            If _lstGrid_3.GetItemChecked(1) Then Call SearchProcess(i)

                        Case OBJ_CH.CH_1P_KEY1 To OBJ_CH.CH_1P_KEY5, OBJ_CH.CH_1P_INV_KEY1 To OBJ_CH.CH_1P_INV_KEY5, OBJ_CH.CH_1P_LN_KEY1 To OBJ_CH.CH_1P_LN_KEY5 '1P 1-5

                            If _lstGrid_0.GetItemChecked(.intCh Mod 36 - 1) Then Call SearchProcess(i)

                        Case OBJ_CH.CH_1P_KEY6, OBJ_CH.CH_1P_KEY7, OBJ_CH.CH_1P_INV_KEY6, OBJ_CH.CH_1P_INV_KEY7, OBJ_CH.CH_1P_LN_KEY6, OBJ_CH.CH_1P_LN_KEY7 '1P 6-7

                            If _lstGrid_0.GetItemChecked(.intCh Mod 36 - 3) Then Call SearchProcess(i)

                        Case OBJ_CH.CH_1P_SC, OBJ_CH.CH_1P_INV_SC, OBJ_CH.CH_1P_LN_SC '1P SC

                            If _lstGrid_0.GetItemChecked(7) Then Call SearchProcess(i)

                        Case OBJ_CH.CH_2P_KEY1 To OBJ_CH.CH_2P_KEY5, OBJ_CH.CH_2P_INV_KEY1 To OBJ_CH.CH_2P_INV_KEY5, OBJ_CH.CH_2P_LN_KEY1 To OBJ_CH.CH_2P_LN_KEY5 '2P 1-5

                            If _lstGrid_1.GetItemChecked(.intCh Mod 10 - 1) Then Call SearchProcess(i)

                        Case OBJ_CH.CH_2P_KEY6, OBJ_CH.CH_2P_KEY7, OBJ_CH.CH_2P_INV_KEY6, OBJ_CH.CH_2P_INV_KEY7, OBJ_CH.CH_2P_LN_KEY6, OBJ_CH.CH_2P_LN_KEY7 '2P 6-7

                            If _lstGrid_1.GetItemChecked(.intCh Mod 10 - 3) Then Call SearchProcess(i)

                        Case OBJ_CH.CH_2P_SC, OBJ_CH.CH_2P_INV_SC, OBJ_CH.CH_2P_LN_SC '2P SC

                            If _lstGrid_1.GetItemChecked(7) Then Call SearchProcess(i)

                        Case OBJ_CH.CH_BGA 'BGA

                            If _lstGrid_3.GetItemChecked(2) Then Call SearchProcess(i)

                        Case OBJ_CH.CH_LAYER 'Layer

                            If _lstGrid_3.GetItemChecked(3) Then Call SearchProcess(i)

                        Case OBJ_CH.CH_POOR 'Poor

                            If _lstGrid_3.GetItemChecked(4) Then Call SearchProcess(i)

                        Case Is > OBJ_CH.CH_BGM_LANE_OFFSET 'BGM

                            If _lstGrid_2.GetItemChecked(.intCh - (OBJ_CH.CH_BGM_LANE_OFFSET + 1)) Then Call SearchProcess(i)

                    End Select

                Else

                    .intSelect = modMain.OBJ_SELECT.NON_SELECT

                End If
				
			End With
			
		Next i
		
		If optProcessDelete.Checked Then
			
			Call frmMain.mnuEditDelete_Click(Nothing, New System.EventArgs())
			
		ElseIf optProcessSelect.Checked Then 
			
			Call modDraw.MoveSelectedObj()
			
		ElseIf optProcessReplace.Checked Then 
			
			If UBound(m_strArray) Then
				
				Call g_InputLog.AddData(Join(m_strArray, modLog.getSeparator))
				
			End If
			
		End If

        frmMain.picMain.Refresh()

    End Sub
	
	Private Sub cmdReset_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdReset.Click

        Dim j As Integer

        With _lstGrid_0

            .Visible = False

            For j = 0 To .Items.Count - 1

                .SetItemChecked(j, False)

            Next j

            .SelectedIndex = 0
            .Visible = True

        End With


        With _lstGrid_1

            .Visible = False

            For j = 0 To .Items.Count - 1

                .SetItemChecked(j, False)

            Next j

            .SelectedIndex = 0
            .Visible = True

        End With


        With _lstGrid_2

            .Visible = False

            For j = 0 To .Items.Count - 1

                .SetItemChecked(j, False)

            Next j

            .SelectedIndex = 0
            .Visible = True

        End With


        With _lstGrid_3

            .Visible = False

            For j = 0 To .Items.Count - 1

                .SetItemChecked(j, False)

            Next j

            .SelectedIndex = 0
            .Visible = True

        End With
    End Sub
	
	Private Sub cmdInvert_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdInvert.Click

        Dim j As Integer

        With _lstGrid_0

            .Visible = False

            For j = 0 To .Items.Count - 1

                .SetItemChecked(j, Not .GetItemChecked(j))

            Next j

            .SelectedIndex = 0
            .Visible = True

        End With

        With _lstGrid_1

            .Visible = False

            For j = 0 To .Items.Count - 1

                .SetItemChecked(j, Not .GetItemChecked(j))

            Next j

            .SelectedIndex = 0
            .Visible = True

        End With

        With _lstGrid_2

            .Visible = False

            For j = 0 To .Items.Count - 1

                .SetItemChecked(j, Not .GetItemChecked(j))

            Next j

            .SelectedIndex = 0
            .Visible = True

        End With

        With _lstGrid_3

            .Visible = False

            For j = 0 To .Items.Count - 1

                .SetItemChecked(j, Not .GetItemChecked(j))

            Next j

            .SelectedIndex = 0
            .Visible = True

        End With

    End Sub
	
	Private Sub cmdSelect_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSelect.Click

        Dim j As Integer

        With _lstGrid_0

            .Visible = False

            For j = 0 To .Items.Count - 1

                .SetItemChecked(j, True)

            Next j

            .SelectedIndex = 0
            .Visible = True

        End With

        With _lstGrid_0

            .Visible = False

            For j = 0 To .Items.Count - 1

                .SetItemChecked(j, True)

            Next j

            .SelectedIndex = 0
            .Visible = True

        End With

        With _lstGrid_2

            .Visible = False

            For j = 0 To .Items.Count - 1

                .SetItemChecked(j, True)

            Next j

            .SelectedIndex = 0
            .Visible = True

        End With

        With _lstGrid_3

            .Visible = False

            For j = 0 To .Items.Count - 1

                .SetItemChecked(j, True)

            Next j

            .SelectedIndex = 0
            .Visible = True

        End With
    End Sub
	
	'UPGRADE_WARNING: Form イベント frmWindowFind.Activate には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
	Private Sub frmWindowFind_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		Call cmdDecide.Focus()
		
	End Sub

    Private Sub frmWindowFind_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim i As Integer

        For i = 2 To 128

            Call _lstGrid_2.Items.Add(Format(i, "000"))

        Next i

        Call cmdSelect_Click(cmdSelect, New System.EventArgs())

        _lstGrid_0.SelectedIndex = 0
        _lstGrid_1.SelectedIndex = 0
        _lstGrid_2.SelectedIndex = 0
        _lstGrid_3.SelectedIndex = 0
    End Sub

    Private Sub txtMeasureMax_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMeasureMax.Enter

        txtMeasureMax.SelectionStart = 0
        txtMeasureMax.SelectionLength = Len(txtMeasureMax.Text)

    End Sub

    Private Sub txtMeasureMin_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMeasureMin.Enter

        txtMeasureMin.SelectionStart = 0
        txtMeasureMin.SelectionLength = Len(txtMeasureMin.Text)

    End Sub

    Private Sub txtNumMax_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumMax.Enter

        txtNumMax.SelectionStart = 0
        txtNumMax.SelectionLength = Len(txtNumMax.Text)

    End Sub

    Private Sub txtNumMin_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumMin.Enter

        txtNumMin.SelectionStart = 0
        txtNumMin.SelectionLength = Len(txtNumMin.Text)

    End Sub

    Private Sub txtReplace_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReplace.Enter

        optProcessReplace.Checked = True
        txtReplace.SelectionStart = 0
        txtReplace.SelectionLength = Len(txtReplace.Text)

    End Sub

    Private Sub SearchProcess(ByVal num As Integer)
        With g_Obj(num)

            If optProcessReplace.Checked Then

                If .intCh <> 8 And .intCh <> 9 Then

                    m_strArray(UBound(m_strArray)) = modInput.strFromNum(modMain.CMD_LOG.OBJ_CHANGE) & modInput.strFromNum(.lngID, 4) & modInput.strFromNum(.sngValue) & VB.Right("0" & txtReplace.Text, 2)
                    ReDim Preserve m_strArray(UBound(m_strArray) + 1)
                    .sngValue = modInput.strToNum((txtReplace.Text))

                End If

            End If

            .intSelect = modMain.OBJ_SELECT.Selected

        End With

    End Sub

    Private Sub frmWindowFind_FormClosed(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        e.Cancel = True

        Call Me.Hide()

        Call frmMain.picMain.Focus()
    End Sub
End Class