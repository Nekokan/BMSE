Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Text

Module modOutput

    Public Sub CreateBMS(ByRef strOutputPath As String, Optional ByVal Flag As Integer = 0)
        On Error GoTo Err_Renamed

        Dim strObjData(,) As String
        Dim blnObjData(,) As Boolean
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim lngFFile As Integer
        Dim intBPMNum As Integer
        Dim intSTOPNum As Integer
        Dim intSCROLLNum As Integer
        Dim intSPEEDNum As Integer
        Dim lngMaxMeasure As Integer
        Dim lngTemp As Integer
        Dim strTemp As String
        Dim intArray() As Integer
        Dim sngSTOP(MATERIAL_MAX) As Single
        Dim sngBPM(MATERIAL_MAX) As Single
        Dim sngSCROLL(MATERIAL_MAX) As Single
        Dim sngSPEED(MATERIAL_MAX) As Single

        If Flag = 0 Then frmMain.Text = g_strAppTitle & " - Now Saving..."

        frmMain.Enabled = False

        For i = 0 To MATERIAL_MAX

            sngBPM(i) = 0
            sngSTOP(i) = 0
            sngSCROLL(i) = 0
            sngSPEED(i) = 0

        Next i

        'オブジェ整理
        For i = 0 To UBound(g_Obj) - 1

            With g_Obj(i)

                If .intCh Then

                    If lngMaxMeasure < .intMeasure Then

                        lngMaxMeasure = .intMeasure

                    End If

                    Select Case .intCh

                        Case modInput.OBJ_CH.CH_EXBPM

                            If .sngValue > 0 And .sngValue < 256 And .sngValue = CInt(.sngValue) Then

                                .intCh = modInput.OBJ_CH.CH_BPM

                            Else

                                If intBPMNum > MATERIAL_MAX Then

                                    Call MsgBox(g_Message(modMain.Message.ERR_OVERFLOW_BPM) & vbCrLf & g_Message(modMain.Message.ERR_SAVE_CANCEL), MsgBoxStyle.Critical, g_strAppTitle)

                                    lngTemp = i - 1

                                    GoTo Init

                                End If

                                If Array.IndexOf(sngBPM, .sngValue) = -1 Then
                                    intBPMNum = intBPMNum + 1
                                    sngBPM(intBPMNum) = .sngValue
                                    .sngValue = intBPMNum
                                Else
                                    .sngValue = Array.IndexOf(sngBPM, .sngValue)
                                End If

                            End If

                        Case modInput.OBJ_CH.CH_STOP

                            If intSTOPNum > MATERIAL_MAX Then

                                Call MsgBox(g_Message(modMain.Message.ERR_OVERFLOW_STOP) & vbCrLf & g_Message(modMain.Message.ERR_SAVE_CANCEL), MsgBoxStyle.Critical, g_strAppTitle)

                                lngTemp = i - 1

                                GoTo Init

                            End If

                            If Array.IndexOf(sngSTOP, .sngValue) = -1 Then
                                intSTOPNum = intSTOPNum + 1
                                sngSTOP(intSTOPNum) = .sngValue
                                .sngValue = intSTOPNum
                            Else
                                .sngValue = Array.IndexOf(sngSTOP, .sngValue)
                            End If

                        Case modInput.OBJ_CH.CH_SCROLL

                            If intSCROLLNum > MATERIAL_MAX Then

                                Call MsgBox(g_Message(modMain.Message.ERR_OVERFLOW_SCROLL) & vbCrLf & g_Message(modMain.Message.ERR_SAVE_CANCEL), MsgBoxStyle.Critical, g_strAppTitle)

                                lngTemp = i - 1

                                GoTo Init

                            End If

                            If .sngValue = 0 Then .sngValue = 1.0E-45 '強引な=0対応、精度的な意味で差は出ないのだ

                            If Array.IndexOf(sngSCROLL, .sngValue) = -1 Then
                                intSCROLLNum = intSCROLLNum + 1
                                sngSCROLL(intSCROLLNum) = .sngValue
                                .sngValue = intSCROLLNum
                            Else
                                .sngValue = Array.IndexOf(sngSCROLL, .sngValue)
                            End If

                        Case modInput.OBJ_CH.CH_SPEED

                            If intSPEEDNum > MATERIAL_MAX Then

                                Call MsgBox(g_Message(modMain.Message.ERR_OVERFLOW_SPEED) & vbCrLf & g_Message(modMain.Message.ERR_SAVE_CANCEL), MsgBoxStyle.Critical, g_strAppTitle)

                                lngTemp = i - 1

                                GoTo Init

                            End If

                            If .sngValue = 0 Then .sngValue = 1.0E-45 '強引な=0対応、精度的な意味で差は出ないのだ

                            If Array.IndexOf(sngSPEED, .sngValue) = -1 Then
                                intSPEEDNum = intSPEEDNum + 1
                                sngSPEED(intSPEEDNum) = .sngValue
                                .sngValue = intSPEEDNum
                            Else
                                .sngValue = Array.IndexOf(sngSPEED, .sngValue)
                            End If

                        Case OBJ_CH.CH_KEY_MIN To OBJ_CH.CH_KEY_MAX

                            If .intAtt = modMain.OBJ_ATT.OBJ_INVISIBLE Then

                                .intCh = .intCh + OBJ_CH.CH_INV

                            ElseIf .intAtt = modMain.OBJ_ATT.OBJ_LONGNOTE Then

                                .intCh = .intCh + OBJ_CH.CH_LN

                            ElseIf .intAtt = modMain.OBJ_ATT.OBJ_MINE Then

                                .intCh = .intCh + OBJ_CH.CH_MINE

                            End If

                    End Select

                End If

            End With

        Next i

        ReDim strObjData(OBJ_CH.CH_BGM_LANE_OFFSET + modInput.BGM_LANE, lngMaxMeasure)
        ReDim blnObjData(OBJ_CH.CH_BGM_LANE_OFFSET + modInput.BGM_LANE, lngMaxMeasure)

        'Measure と intCh ごとのObjの位置の必要分割数 ReqDev(intCh, intMeasure) を得る
        Dim ReqDev(,) As Integer = GetReqDevision(g_Obj)
        'ここまで

        For i = 0 To lngMaxMeasure

            For j = LBound(strObjData, 1) To UBound(strObjData, 1)

                If ReqDev(j, i) = 0 Then
                    strObjData(j, i) = New String("0", 2)
                Else
                    strObjData(j, i) = New String("0", ReqDev(j, i) * 2)
                End If

            Next j

        Next i

        'オブジェからラインデータに変換
        For i = 0 To UBound(g_Obj) - 1

            With g_Obj(i)

                '分数Position
                Dim temp() As Integer = GetFraction(.lngPosition)
                '分数Positionを x/必要分割数 の形に
                Dim Numerator As Integer = temp(0) \ intGCD(temp(0), g_Measure(.intMeasure).intLen)
                Dim Denominator As Integer = temp(1) * g_Measure(.intMeasure).intLen \ intGCD(temp(0), g_Measure(.intMeasure).intLen)
                '通分
                Numerator = Numerator * ReqDev(.intCh, .intMeasure) \ Denominator
                'Denominator = ReqDev(.intCh, .intMeasure)

                Select Case .intCh

                    Case Is < 0

                    Case Is > 10000

                    Case Is > OBJ_CH.CH_BGM_LANE_OFFSET

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), Numerator * 2) & Right("0" & IIf(frmMain._mnuOptionsBase62.Checked, modInput.strFromNum62ZZ(.sngValue), IIf(frmMain._mnuOptionsBase16.Checked, strFromNumFF(.sngValue), modInput.strFromNumZZ(.sngValue))), 2) & Mid(strObjData(.intCh, .intMeasure), Numerator * 2 + 3)

                        For j = OBJ_CH.CH_BGM_LANE_OFFSET + 1 To .intCh - 1

                            blnObjData(j, .intMeasure) = True

                        Next j

                    Case modInput.OBJ_CH.CH_BPM

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), Numerator * 2) & Right("0" & Hex(.sngValue), 2) & Mid(strObjData(.intCh, .intMeasure), Numerator * 2 + 3)

                    Case modInput.OBJ_CH.CH_EXBPM

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), Numerator * 2) & Right("0" & IIf(intBPMNum > 1295, modInput.strFromNum62ZZ(.sngValue), modInput.strFromNumZZ(.sngValue)), 2) & Mid(strObjData(.intCh, .intMeasure), Numerator * 2 + 3)

                    Case modInput.OBJ_CH.CH_STOP

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), Numerator * 2) & Right("0" & IIf(intSTOPNum > 1295, modInput.strFromNum62ZZ(.sngValue), modInput.strFromNumZZ(.sngValue)), 2) & Mid(strObjData(.intCh, .intMeasure), Numerator * 2 + 3)

                    Case modInput.OBJ_CH.CH_SCROLL

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), Numerator * 2) & Right("0" & IIf(intSCROLLNum > 1295, modInput.strFromNum62ZZ(.sngValue), modInput.strFromNumZZ(.sngValue)), 2) & Mid(strObjData(.intCh, .intMeasure), Numerator * 2 + 3)

                    Case modInput.OBJ_CH.CH_SPEED

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), Numerator * 2) & Right("0" & IIf(intSPEEDNum > 1295, modInput.strFromNum62ZZ(.sngValue), modInput.strFromNumZZ(.sngValue)), 2) & Mid(strObjData(.intCh, .intMeasure), Numerator * 2 + 3)

                    Case modInput.OBJ_CH.CH_KEY_MINE_MIN To modInput.OBJ_CH.CH_KEY_MINE_MAX ' 地雷だけは36進数（でなければいけないはず）

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), Numerator * 2) & modInput.strFromNumZZ(.sngValue) & Mid(strObjData(.intCh, .intMeasure), Numerator * 2 + 3)

                    Case Else

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), Numerator * 2) & Right("0" & IIf(frmMain._mnuOptionsBase62.Checked, modInput.strFromNum62ZZ(.sngValue), IIf(frmMain._mnuOptionsBase16.Checked, strFromNumFF(.sngValue), modInput.strFromNumZZ(.sngValue))), 2) & Mid(strObjData(.intCh, .intMeasure), Numerator * 2 + 3)

                End Select

                blnObjData(.intCh, .intMeasure) = True

            End With

        Next i

        '先に分割数を計算するようにしたので以下のブロック丸々不要のはず（？）

        'For i = LBound(strObjData, 2) To UBound(strObjData, 2)

        '    For j = LBound(strObjData, 1) To UBound(strObjData, 1)

        '        If blnObjData(j, i) Then

        '            If strObjData(j, i) <> "00" Then

        '                ReDim intArray(g_Measure(i).intLen + 1)

        '                intArray(0) = g_Measure(i).intLen
        '                strTemp = ""
        '                lngTemp = 1

        '                For k = 1 To Len(strObjData(j, i)) \ 2

        '                    If Mid(strObjData(j, i), k * 2 - 1, 2) = "00" Then

        '                        strTemp = strTemp & "0"

        '                    Else

        '                        intArray(lngTemp) = Len(strTemp)
        '                        lngTemp = lngTemp + 1
        '                        strTemp = "1"

        '                    End If

        '                Next k

        '                ReDim Preserve intArray(lngTemp)

        '                intArray(lngTemp) = Len(strTemp)

        '                lngTemp = intGetMaxDev(intArray)

        '                If lngTemp Then

        '                    strTemp = ""

        '                    For k = 1 To Len(strObjData(j, i)) \ 2 Step lngTemp

        '                        strTemp = strTemp & Mid(strObjData(j, i), k * 2 - 1, 2)

        '                    Next k

        '                    strObjData(j, i) = strTemp

        '                End If

        '            End If

        '        End If

        '    Next j

        'Next i

        Using writer As New StreamWriter(strOutputPath, False, Encoding.Default, 7 + 192 * 10000 * 2) 'Position最小単位=1/10000 -> 最大分割数=192*10000

            With frmMain

                writer.WriteLine()
                writer.WriteLine("*---------------------- HEADER FIELD")
                writer.WriteLine()
                'If Flag Then Print #lngFFile, "#PATH_WAV " & g_BMS.strDir

                If .cboPlayer.SelectedIndex > 1 Then

                    writer.WriteLine("#PLAYER 3")

                Else

                    writer.WriteLine("#PLAYER " & .cboPlayer.SelectedIndex + 1)

                End If

                writer.WriteLine("#GENRE " & Trim(.txtGenre.Text))
                writer.WriteLine("#TITLE " & Trim(.txtTitle.Text))
                writer.WriteLine("#ARTIST " & Trim(.txtArtist.Text))
                writer.WriteLine("#BPM " & Trim(.txtBPM.Text))
                writer.WriteLine("#PLAYLEVEL " & Trim(.cboPlayLevel.Text))
                writer.WriteLine("#RANK " & .cboPlayRank.SelectedIndex)

                If Val(.txtTotal.Text) Then writer.WriteLine("#TOTAL " & .txtTotal.Text)
                If Val(.txtVolume.Text) Then writer.WriteLine("#VOLWAV " & .txtVolume.Text)

                If frmMain._mnuOptionsBase62.Checked AndAlso
                    (IsRequireBase62(g_strWAV) OrElse IsRequireBase62(g_strBMP) OrElse
                    IsRequireBase62(sngBPM) OrElse IsRequireBase62(sngSTOP) OrElse
                    IsRequireBase62(sngSCROLL) OrElse IsRequireBase62(sngSPEED)) Then

                    writer.WriteLine("#BASE 62")
                Else
                    writer.WriteLine("#BASE 36")
                End If

                writer.WriteLine()

                If Trim(.txtSubTitle.Text) <> "" Then writer.WriteLine("#SUBTITLE " & Trim(.txtSubTitle.Text))
                If Trim(.txtSubArtist.Text) <> "" Then writer.WriteLine("#SUBARTIST " & Trim(.txtSubArtist.Text))
                If .cboDifficulty.SelectedIndex > 0 Then writer.WriteLine("#DIFFICULTY " & .cboDifficulty.SelectedIndex)
                If Trim(.txtStageFile.Text) <> "" Then writer.WriteLine("#STAGEFILE " & Trim(.txtStageFile.Text))
                If Trim(.txtPreview.Text) <> "" Then writer.WriteLine("#PREVIEW " & Trim(.txtPreview.Text))
                If Trim(.txtBanner.Text) <> "" Then writer.WriteLine("#BANNER " & Trim(.txtBanner.Text))
                If Trim(.txtBackBmp.Text) <> "" Then writer.WriteLine("#BACKBMP " & Trim(.txtBackBmp.Text))
                If Trim(.txtDefExRank.Text) <> "" Then writer.WriteLine("#DEFEXRANK " & Trim(.txtDefExRank.Text))
                If .cboLNMode.SelectedIndex > 0 Then writer.WriteLine("#LNMODE " & .cboLNMode.SelectedIndex)

                If .cboLNObj.SelectedIndex > 0 Then
                    writer.WriteLine("#LNOBJ " & strFromNum(.cboLNObj.SelectedIndex))
                Else
                    writer.WriteLine("#LNTYPE 1")
                End If

                If Trim(.txtComment.Text) <> "" Then
                    '#COMMENTはダブルクオーテーション Chr(34) 必須のための処理 
                    If Left(Trim(.txtComment.Text), 1) = Chr(34) And Right(Trim(.txtComment.Text), 1) = Chr(34) Then
                        strTemp = Trim(.txtComment.Text)
                    Else
                        strTemp = Chr(34) & Trim(.txtComment.Text) & Chr(34)
                    End If
                    writer.WriteLine("#COMMENT " & strTemp)
                End If

                writer.WriteLine()

                If Len(Trim(.txtLandmineWAV.Text)) Then

                    writer.WriteLine("#WAV00 " & .txtLandmineWAV.Text)

                End If

                For i = 1 To MATERIAL_MAX

                    If Len(g_strWAV(i)) Then

                        If frmMain._mnuOptionsBase62.Checked Then
                            writer.WriteLine("#WAV" & modInput.strFromNum62ZZ(i) & " " & g_strWAV(i))
                        ElseIf frmMain._mnuOptionsBase16.Checked Then
                            writer.WriteLine("#WAV" & modInput.strFromNumFF(i) & " " & g_strWAV(i))
                        Else
                            writer.WriteLine("#WAV" & modInput.strFromNumZZ(i) & " " & g_strWAV(i))
                        End If

                    End If

                Next i

                writer.WriteLine()

                If Len(Trim(.txtMissBMP.Text)) Then

                    writer.WriteLine("#BMP00 " & .txtMissBMP.Text)

                End If

                For i = 1 To MATERIAL_MAX

                    If Len(g_strBMP(i)) Then

                        If frmMain._mnuOptionsBase62.Checked Then
                            writer.WriteLine("#BMP" & modInput.strFromNum62ZZ(i) & " " & g_strBMP(i))
                        ElseIf frmMain._mnuOptionsBase16.Checked Then
                            writer.WriteLine("#BMP" & modInput.strFromNumFF(i) & " " & g_strBMP(i))
                        Else
                            writer.WriteLine("#BMP" & modInput.strFromNumZZ(i) & " " & g_strBMP(i))
                        End If

                    End If

                Next i

                writer.WriteLine()

                For i = 1 To MATERIAL_MAX

                    If Len(g_strBGA(i)) Then

                        If frmMain._mnuOptionsBase62.Checked Then
                            writer.WriteLine("#BGA" & modInput.strFromNum62ZZ(i) & " " & g_strBGA(i))
                        ElseIf frmMain._mnuOptionsBase16.Checked Then
                            writer.WriteLine("#BGA" & modInput.strFromNumFF(i) & " " & g_strBGA(i))
                        Else
                            writer.WriteLine("#BGA" & modInput.strFromNumZZ(i) & " " & g_strBGA(i))
                        End If

                    End If

                Next i

                writer.WriteLine()

                If intBPMNum > 1295 Then

                    For i = 1 To MATERIAL_MAX

                        If sngBPM(i) Then

                            writer.WriteLine("#BPM" & Right("0" & modInput.strFromNum62ZZ(i), 2) & " " & CDec(sngBPM(i)))

                        End If

                    Next i

                ElseIf intBPMNum Then

                    For i = 1 To 1295

                        If sngBPM(i) Then

                            writer.WriteLine("#BPM" & Right("0" & modInput.strFromNumZZ(i), 2) & " " & CDec(sngBPM(i)))

                        End If

                    Next i

                End If

                writer.WriteLine()

                If intSTOPNum > 1295 Then

                    For i = 1 To MATERIAL_MAX

                        If sngSTOP(i) Then

                            writer.WriteLine("#STOP" & Right("0" & modInput.strFromNum62ZZ(i), 2) & " " & sngSTOP(i))

                        End If

                    Next i

                ElseIf intSTOPNum Then

                    For i = 1 To 1295

                        If sngSTOP(i) Then

                            writer.WriteLine("#STOP" & Right("0" & modInput.strFromNumZZ(i), 2) & " " & sngSTOP(i))

                        End If

                    Next i

                End If

                If intSCROLLNum Then

                    For i = 1 To MATERIAL_MAX

                        If sngSCROLL(i) Then

                            writer.WriteLine("#SCROLL" & Right("0" & modInput.strFromNum62ZZ(i), 2) & " " & sngSCROLL(i))

                        End If

                    Next i

                End If

                If intSPEEDNum Then

                    For i = 1 To MATERIAL_MAX

                        If sngSPEED(i) Then

                            writer.WriteLine("#SPEED" & Right("0" & modInput.strFromNum62ZZ(i), 2) & " " & sngSPEED(i))

                        End If

                    Next i

                End If

                writer.WriteLine()

                writer.WriteLine(.txtExInfo.Text)

                writer.WriteLine()

            End With

            writer.WriteLine()
            writer.WriteLine("*---------------------- MAIN DATA FIELD")
            writer.WriteLine()

            For i = 0 To UBound(blnObjData, 2)

                For j = OBJ_CH.CH_BGM_LANE_OFFSET + 1 To OBJ_CH.CH_BGM_LANE_OFFSET + modInput.BGM_LANE

                    If blnObjData(j, i) Then

                        writer.WriteLine("#" & Format(i, "000") & "01" & ":" & strObjData(j, i))

                    End If

                Next j

                With g_Measure(i)

                    If .intLen <> MEASURE_LENGTH Then

                        writer.WriteLine("#" & Format(i, "000") & "02:" & .intLen / MEASURE_LENGTH)

                    End If

                End With

                For j = 3 To OBJ_CH.CH_BGM_LANE_OFFSET - 1

                    If blnObjData(j, i) Then

                        writer.WriteLine("#" & Format(i, "000") & strFromNumZZ(j) & ":" & strObjData(j, i))

                    End If

                Next j

                writer.WriteLine()

            Next i

            lngTemp = UBound(blnObjData, 2) + 1

            For i = lngTemp To 999

                With g_Measure(i)

                    If .intLen <> MEASURE_LENGTH Then

                        writer.WriteLine("#" & Format(i, "000") & "02:" & .intLen / MEASURE_LENGTH)

                    End If

                End With

            Next i

            lngTemp = UBound(g_Obj) - 1

            With g_BMS

                .intPlayerType = frmMain.cboPlayer.SelectedIndex + 1
                .strGenre = frmMain.txtGenre.Text
                .strTitle = frmMain.txtTitle.Text
                .strArtist = frmMain.txtArtist.Text
                .lngPlayLevel = Val(frmMain.cboPlayLevel.Text)
                .sngBPM = Val(frmMain.txtBPM.Text)

                .intPlayRank = frmMain.cboPlayRank.SelectedIndex
                .sngTotal = Val(frmMain.txtTotal.Text)
                .intVolume = Val(frmMain.txtVolume.Text)
                .strStageFile = frmMain.txtStageFile.Text

                .strSubTitle = frmMain.txtSubTitle.Text
                .strSubArtist = frmMain.txtSubArtist.Text
                .intDifficulty = frmMain.cboDifficulty.SelectedIndex
                .strPreview = frmMain.txtPreview.Text
                .strBanner = frmMain.txtBanner.Text
                .intLNObj = frmMain.cboLNObj.SelectedIndex
                .intLNMode = frmMain.cboLNMode.SelectedIndex
                .intDefExRank = CInt(Val(frmMain.txtDefExRank.Text))
                .strBackBMP = frmMain.txtBackBmp.Text
                .strComment = frmMain.txtComment.Text

            End With

        End Using

Init:

        'FileClose(lngFFile)

        For i = 0 To lngTemp

            With g_Obj(i)

                Select Case .intCh

                    Case modInput.OBJ_CH.CH_BPM

                        .intCh = modInput.OBJ_CH.CH_EXBPM

                    Case modInput.OBJ_CH.CH_EXBPM

                        .sngValue = sngBPM(.sngValue)

                    Case modInput.OBJ_CH.CH_STOP

                        .sngValue = sngSTOP(.sngValue)

                    Case modInput.OBJ_CH.CH_SCROLL

                        .sngValue = sngSCROLL(.sngValue)

                    Case modInput.OBJ_CH.CH_SPEED

                        .sngValue = sngSPEED(.sngValue)

                    Case OBJ_CH.CH_KEY_INV_MIN To OBJ_CH.CH_KEY_INV_MAX

                        .intCh = .intCh - OBJ_CH.CH_INV

                    Case OBJ_CH.CH_KEY_LN_MIN To OBJ_CH.CH_KEY_LN_MAX

                        .intCh = .intCh - OBJ_CH.CH_LN

                    Case OBJ_CH.CH_KEY_MINE_MIN To OBJ_CH.CH_KEY_MINE_MAX

                        .intCh = .intCh - OBJ_CH.CH_MINE

                End Select

            End With

        Next i

        frmMain.Enabled = True

        If Flag = 0 Then

            g_BMS.blnSaveFlag = True

            If Len(g_BMS.strDir) Then

                If frmMain._mnuOptionsItem_1.Checked Then

                    frmMain.Text = g_strAppTitle & " - " & g_BMS.strFileName

                Else

                    frmMain.Text = g_strAppTitle & " - " & g_BMS.strDir & g_BMS.strFileName

                End If

            End If

        End If

        Exit Sub

Err_Renamed:
        Call MsgBox(g_Message(modMain.Message.ERR_SAVE_ERROR) & vbCrLf & g_Message(modMain.Message.ERR_SAVE_CANCEL) & vbCrLf & "Error No." & Err.Number & " " & Err.Description, MsgBoxStyle.Critical, g_strAppTitle)
        frmMain.Enabled = True
        frmMain.Text = g_strAppTitle & " - " & g_BMS.strDir & g_BMS.strFileName
    End Sub

    Private Function intGetMaxDev(ByRef BaseValue() As Integer) As Integer

        Dim Count As Integer '配列の最大インデックス
        Dim i As Integer 'カウンタ
        Dim a, b As Integer '最大公約数を求める2つの要素

        Count = UBound(BaseValue)
        a = BaseValue(0)

        '繰り返す回数は、(配列の数－1)回
        For i = 1 To Count

            b = BaseValue(i)

            If b Then

                Do While a <> b

                    If a > b Then

                        a = a - b

                    Else

                        b = b - a

                    End If

                Loop

                '1で等しい場合、最大公約数はない
                If a = 1 Then intGetMaxDev = 0 : Exit Function

            End If

        Next i

        '最大公約数を返す
        intGetMaxDev = a

    End Function

    '最小公倍数
    Function intLCM(a As Integer, b As Integer) As Integer
        If a = 0 OrElse b = 0 Then Return 0
        Return Math.Abs(a * b / intGCD(a, b))
    End Function

    '小数を分数に変換
    '戻り値 = {分子, 分母}
    Function GetFraction(dbl As Double) As Integer()

        Const MaxDenominator As Integer = 10000
        Const Tolerance As Double = 1 / (MaxDenominator ^ 2)
        Dim i As Integer

        If dbl < 0 Then Return {CInt(dbl * MaxDenominator）, MaxDenominator}
        If dbl = 0 Then Return {0, 1} '0 = 0/1 とする

        For i = 1 To MaxDenominator

            Dim Nume1 As Integer = Int(dbl * i)
            Dim Nume2 As Integer = Nume1 + 1

            Dim d1 As Double = Nume1 / i
            Dim d2 As Double = Nume2 / i

            If dbl - d1 < Tolerance Then　'dbl > 0 -> dbl > d1

                Return {Nume1, i}
                Exit Function

            ElseIf d2 - dbl < Tolerance Then　'dbl > 0 -> d2 > dbl

                Return {Nume2, i}
                Exit Function

            End If

        Next i

        '最後まで見つからなかった
        Return {CInt(dbl * MaxDenominator）, MaxDenominator}

    End Function

    'ChとMeasureごとに必要な分割数を取得
    '(Ch, Measure)
    Function GetReqDevision(g_Obj() As g_udtObj) As Integer(,)
        Dim intArray(,) As Integer
        ReDim intArray(OBJ_CH.CH_BGM_LANE_OFFSET + modInput.BGM_LANE, MEASURE_MAX)
        Dim i As Integer
        Dim g_ObjClone() As g_udtObj
        g_ObjClone = g_Obj.Clone()

        For i = 0 To UBound(g_ObjClone) - 1

            'Positionを分数(DevPosition)化して小節内位置(PosAtMeasure(0)/PosAtMeasure(1))を求める
            Dim DevPosition As Integer() = GetFraction(g_ObjClone(i).lngPosition)
            Dim PosAtMeasure As Integer() = {DevPosition(0), DevPosition(1) * g_Measure(g_ObjClone(i).intMeasure).intLen / intGCD(g_Measure(g_ObjClone(i).intMeasure).intLen, DevPosition(0))}

            'Positionを分数化したときの分母の最小公倍数
            If intArray(g_ObjClone(i).intCh, g_ObjClone(i).intMeasure) = 0 Then
                intArray(g_ObjClone(i).intCh, g_ObjClone(i).intMeasure) = PosAtMeasure(1)
            Else
                intArray(g_ObjClone(i).intCh, g_ObjClone(i).intMeasure) = intLCM(PosAtMeasure(1), intArray(g_ObjClone(i).intCh, g_ObjClone(i).intMeasure))
            End If

        Next

        ArrangeObj()

        Return intArray

    End Function

    'strArray: WAV,BMPのファイル名
    Public Function IsRequireBase62(strArray As String()) As Boolean
        Dim i As Integer

        For i = 0 To UBound(strArray)
            Select Case i
                Case strToNum62ZZ("0a") To strToNum62ZZ("0z"),
                    strToNum62ZZ("1a") To strToNum62ZZ("1z"),
                    strToNum62ZZ("2a") To strToNum62ZZ("2z"),
                    strToNum62ZZ("3a") To strToNum62ZZ("3z"),
                    strToNum62ZZ("4a") To strToNum62ZZ("4z"),
                    strToNum62ZZ("5a") To strToNum62ZZ("5z"),
                    strToNum62ZZ("6a") To strToNum62ZZ("6z"),
                    strToNum62ZZ("7a") To strToNum62ZZ("7z"),
                    strToNum62ZZ("8a") To strToNum62ZZ("8z"),
                    Is >= strToNum62ZZ("9a")

                    If Len(strArray(i)) > 0 Then
                        Return True
                    End If
            End Select
        Next

        Return False
    End Function

    'sngArray:BPM,STOP,SCROLL,SPEED
    Public Function IsRequireBase62(sngArray As Single()) As Boolean
        Dim i As Integer

        For i = 0 To UBound(sngArray)
            If i >= 1296 Then
                Return True
            Else
                If sngArray(i) = 0 Then
                    Return False
                End If
            End If
        Next

        Return True
    End Function

End Module