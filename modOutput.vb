Option Strict Off
Option Explicit On
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

                        Case 1 * 36 + 1 To 2 * 36 + 9

                            If .intAtt = modMain.OBJ_ATT.OBJ_INVISIBLE Then

                                .intCh = .intCh + 2 * 36 + 0

                            ElseIf .intAtt = modMain.OBJ_ATT.OBJ_LONGNOTE Then

                                .intCh = .intCh + 4 * 36 + 0

                            ElseIf .intAtt = modMain.OBJ_ATT.OBJ_MINE Then

                                .intCh = .intCh + 12 * 36 + 0

                            End If

                    End Select

                End If

            End With

        Next i

        ReDim strObjData(36 ^ 2 + modInput.BGM_LANE, lngMaxMeasure)
        ReDim blnObjData(36 ^ 2 + modInput.BGM_LANE, lngMaxMeasure)

        For i = 0 To lngMaxMeasure

            For j = LBound(strObjData, 1) To UBound(strObjData, 1)

                strObjData(j, i) = New String("0", g_Measure(i).intLen * 2)

            Next j

        Next i

        'オブジェからラインデータに変換
        For i = 0 To UBound(g_Obj) - 1

            With g_Obj(i)

                Select Case .intCh

                    Case Is < 0

                    Case Is > 10000

                    Case Is > 36 ^ 2

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), .lngPosition * 2) & Right("0" & IIf(frmMain._mnuOptionsBase62.Checked, modInput.strFromNum62ZZ(.sngValue), IIf(frmMain._mnuOptionsBase16.Checked, strFromNumFF(.sngValue), modInput.strFromNumZZ(.sngValue))), 2) & Mid(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)

                        For j = 36 ^ 2 + 1 To .intCh - 1

                            blnObjData(j, .intMeasure) = True

                        Next j

                    Case modInput.OBJ_CH.CH_BPM

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), .lngPosition * 2) & Right("0" & Hex(.sngValue), 2) & Mid(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)

                    Case modInput.OBJ_CH.CH_EXBPM

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), .lngPosition * 2) & Right("0" & IIf(intBPMNum > 1295, modInput.strFromNum62ZZ(.sngValue), modInput.strFromNumZZ(.sngValue)), 2) & Mid(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)

                    Case modInput.OBJ_CH.CH_STOP

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), .lngPosition * 2) & Right("0" & IIf(intSTOPNum > 1295, modInput.strFromNum62ZZ(.sngValue), modInput.strFromNumZZ(.sngValue)), 2) & Mid(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)

                    Case modInput.OBJ_CH.CH_SCROLL

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), .lngPosition * 2) & Right("0" & IIf(intSCROLLNum > 1295, modInput.strFromNum62ZZ(.sngValue), modInput.strFromNumZZ(.sngValue)), 2) & Mid(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)

                    Case modInput.OBJ_CH.CH_SPEED

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), .lngPosition * 2) & Right("0" & IIf(intSPEEDNum > 1295, modInput.strFromNum62ZZ(.sngValue), modInput.strFromNumZZ(.sngValue)), 2) & Mid(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)

                    Case modInput.OBJ_CH.CH_KEY_MINE_MIN To modInput.OBJ_CH.CH_KEY_MINE_MAX ' 地雷だけは36進数（でなければいけないはず）

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), .lngPosition * 2) & modInput.strFromNumZZ(.sngValue) & Mid(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)

                    Case Else

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), .lngPosition * 2) & Right("0" & IIf(frmMain._mnuOptionsBase62.Checked, modInput.strFromNum62ZZ(.sngValue), IIf(frmMain._mnuOptionsBase16.Checked, strFromNumFF(.sngValue), modInput.strFromNumZZ(.sngValue))), 2) & Mid(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)

                End Select

                blnObjData(.intCh, .intMeasure) = True

            End With

        Next i

        For i = LBound(strObjData, 2) To UBound(strObjData, 2)

            For j = LBound(strObjData, 1) To UBound(strObjData, 1)

                If blnObjData(j, i) Then

                    If strObjData(j, i) <> "00" Then

                        ReDim intArray(g_Measure(i).intLen + 1)

                        intArray(0) = g_Measure(i).intLen
                        strTemp = ""
                        lngTemp = 1

                        For k = 1 To Len(strObjData(j, i)) \ 2

                            If Mid(strObjData(j, i), k * 2 - 1, 2) = "00" Then

                                strTemp = strTemp & "0"

                            Else

                                intArray(lngTemp) = Len(strTemp)
                                lngTemp = lngTemp + 1
                                strTemp = "1"

                            End If

                        Next k

                        ReDim Preserve intArray(lngTemp)

                        intArray(lngTemp) = Len(strTemp)

                        lngTemp = intGetMaxDev(intArray)

                        If lngTemp Then

                            strTemp = ""

                            For k = 1 To Len(strObjData(j, i)) \ 2 Step lngTemp

                                strTemp = strTemp & Mid(strObjData(j, i), k * 2 - 1, 2)

                            Next k

                            strObjData(j, i) = strTemp

                        End If

                    End If

                End If

            Next j

        Next i

        lngFFile = FreeFile()

        '出力開始
        FileOpen(lngFFile, strOutputPath, OpenMode.Output)

        With frmMain

            PrintLine(lngFFile)
            PrintLine(lngFFile, "*---------------------- HEADER FIELD")
            PrintLine(lngFFile)
            'If Flag Then Print #lngFFile, "#PATH_WAV " & g_BMS.strDir

            If .cboPlayer.SelectedIndex > 1 Then

                PrintLine(lngFFile, "#PLAYER 3")

            Else

                PrintLine(lngFFile, "#PLAYER " & .cboPlayer.SelectedIndex + 1)

            End If

            PrintLine(lngFFile, "#GENRE " & Trim(.txtGenre.Text))
            PrintLine(lngFFile, "#TITLE " & Trim(.txtTitle.Text))
            PrintLine(lngFFile, "#ARTIST " & Trim(.txtArtist.Text))
            PrintLine(lngFFile, "#BPM " & Trim(.txtBPM.Text))
            PrintLine(lngFFile, "#PLAYLEVEL " & Trim(.cboPlayLevel.Text))
            PrintLine(lngFFile, "#RANK " & .cboPlayRank.SelectedIndex)

            If Val(.txtTotal.Text) Then PrintLine(lngFFile, "#TOTAL " & .txtTotal.Text)
            If Val(.txtVolume.Text) Then PrintLine(lngFFile, "#VOLWAV " & .txtVolume.Text)

            If frmMain._mnuOptionsBase62.Checked Then
                PrintLine(lngFFile, "#BASE 62")
            Else
                PrintLine(lngFFile, "#BASE 36")
            End If

            PrintLine(lngFFile)

            If Trim(.txtSubTitle.Text) <> "" Then PrintLine(lngFFile, "#SUBTITLE " & Trim(.txtSubTitle.Text))
            If Trim(.txtSubArtist.Text) <> "" Then PrintLine(lngFFile, "#SUBARTIST " & Trim(.txtSubArtist.Text))
            If .cboDifficulty.SelectedIndex > 0 Then PrintLine(lngFFile, "#DIFFICULTY " & .cboDifficulty.SelectedIndex)
            If Trim(.txtStageFile.Text) <> "" Then PrintLine(lngFFile, "#STAGEFILE " & Trim(.txtStageFile.Text))
            If Trim(.txtPreview.Text) <> "" Then PrintLine(lngFFile, "#PREVIEW " & Trim(.txtPreview.Text))
            If Trim(.txtBanner.Text) <> "" Then PrintLine(lngFFile, "#BANNER " & Trim(.txtBanner.Text))
            If Trim(.txtBackBmp.Text) <> "" Then PrintLine(lngFFile, "#BACKBMP " & Trim(.txtBackBmp.Text))
            If Trim(.txtDefExRank.Text) <> "" Then PrintLine(lngFFile, "#DEFEXRANK " & Trim(.txtDefExRank.Text))
            If .cboLNMode.SelectedIndex > 0 Then PrintLine(lngFFile, "#LNMODE " & .cboLNMode.SelectedIndex)

            If .cboLNObj.SelectedIndex > 0 Then
                PrintLine(lngFFile, "#LNOBJ " & strFromNum(.cboLNObj.SelectedIndex))
            Else
                PrintLine(lngFFile, "#LNTYPE 1")
            End If

            If Trim(.txtComment.Text) <> "" Then
                '#COMMENTはダブルクオーテーション Chr(34) 必須のための処理 
                If Left(Trim(.txtComment.Text), 1) = Chr(34) And Right(Trim(.txtComment.Text), 1) = Chr(34) Then
                    strTemp = Trim(.txtComment.Text)
                Else
                    strTemp = Chr(34) & Trim(.txtComment.Text) & Chr(34)
                End If
                PrintLine(lngFFile, "#COMMENT " & strTemp)
            End If

            PrintLine(lngFFile)

            For i = 1 To MATERIAL_MAX

                If Len(g_strWAV(i)) Then

                    If frmMain._mnuOptionsBase62.Checked Then
                        PrintLine(lngFFile, "#WAV" & modInput.strFromNum62ZZ(i) & " " & g_strWAV(i))
                    ElseIf frmMain._mnuOptionsBase16.Checked Then
                        PrintLine(lngFFile, "#WAV" & modInput.strFromNumFF(i) & " " & g_strWAV(i))
                    Else
                        PrintLine(lngFFile, "#WAV" & modInput.strFromNumZZ(i) & " " & g_strWAV(i))
                    End If

                End If

            Next i

            PrintLine(lngFFile)

            If Len(Trim(.txtMissBMP.Text)) Then

                PrintLine(lngFFile, "#BMP00 " & .txtMissBMP.Text)

            End If

            If Len(Trim(.txtLandmineWAV.Text)) Then

                PrintLine(lngFFile, "#WAV00 " & .txtLandmineWAV.Text)

            End If

            For i = 1 To MATERIAL_MAX

                If Len(g_strBMP(i)) Then

                    If frmMain._mnuOptionsBase62.Checked Then
                        PrintLine(lngFFile, "#BMP" & modInput.strFromNum62ZZ(i) & " " & g_strBMP(i))
                    ElseIf frmMain._mnuOptionsBase16.Checked Then
                        PrintLine(lngFFile, "#BMP" & modInput.strFromNumFF(i) & " " & g_strBMP(i))
                    Else
                        PrintLine(lngFFile, "#BMP" & modInput.strFromNumZZ(i) & " " & g_strBMP(i))
                    End If

                End If

            Next i

            PrintLine(lngFFile)

            For i = 1 To MATERIAL_MAX

                If Len(g_strBGA(i)) Then

                    If frmMain._mnuOptionsBase62.Checked Then
                        PrintLine(lngFFile, "#BGA" & modInput.strFromNum62ZZ(i) & " " & g_strBGA(i))
                    ElseIf frmMain._mnuOptionsBase16.Checked Then
                        PrintLine(lngFFile, "#BGA" & modInput.strFromNumFF(i) & " " & g_strBGA(i))
                    Else
                        PrintLine(lngFFile, "#BGA" & modInput.strFromNumZZ(i) & " " & g_strBGA(i))
                    End If

                End If

            Next i

            PrintLine(lngFFile)

            If intBPMNum > 1295 Then

                For i = 1 To MATERIAL_MAX

                    If sngBPM(i) Then

                        PrintLine(lngFFile, "#BPM" & Right("0" & modInput.strFromNum62ZZ(i), 2) & " " & CDec(sngBPM(i)))

                    End If

                Next i

            ElseIf intBPMNum Then

                For i = 1 To 1295

                    If sngBPM(i) Then

                        PrintLine(lngFFile, "#BPM" & Right("0" & modInput.strFromNumZZ(i), 2) & " " & CDec(sngBPM(i)))

                    End If

                Next i

            End If

            PrintLine(lngFFile)

            If intSTOPNum > 1295 Then

                For i = 1 To MATERIAL_MAX

                    If sngSTOP(i) Then

                        PrintLine(lngFFile, "#STOP" & Right("0" & modInput.strFromNum62ZZ(i), 2) & " " & sngSTOP(i))

                    End If

                Next i

            ElseIf intSTOPNum Then

                For i = 1 To 1295

                    If sngSTOP(i) Then

                        PrintLine(lngFFile, "#STOP" & Right("0" & modInput.strFromNumZZ(i), 2) & " " & sngSTOP(i))

                    End If

                Next i

            End If

            If intSCROLLNum Then

                For i = 1 To MATERIAL_MAX

                    If sngSCROLL(i) Then

                        PrintLine(lngFFile, "#SCROLL" & Right("0" & modInput.strFromNum62ZZ(i), 2) & " " & sngSCROLL(i))

                    End If

                Next i

            End If

            If intSPEEDNum Then

                For i = 1 To MATERIAL_MAX

                    If sngSPEED(i) Then

                        PrintLine(lngFFile, "#SPEED" & Right("0" & modInput.strFromNum62ZZ(i), 2) & " " & sngSPEED(i))

                    End If

                Next i

            End If

            PrintLine(lngFFile)

            PrintLine(lngFFile, .txtExInfo.Text)

            PrintLine(lngFFile)

        End With

        PrintLine(lngFFile)
        PrintLine(lngFFile, "*---------------------- MAIN DATA FIELD")
        PrintLine(lngFFile)

        For i = 0 To UBound(blnObjData, 2)

            For j = 36 ^ 2 + 1 To 36 ^ 2 + 1 + modInput.BGM_LANE - 1

                If blnObjData(j, i) Then

                    PrintLine(lngFFile, "#" & Format(i, "000") & "01" & ":" & strObjData(j, i))

                End If

            Next j

            With g_Measure(i)

                If .intLen <> MEASURE_LENGTH Then

                    PrintLine(lngFFile, "#" & Format(i, "000") & "02:" & .intLen / MEASURE_LENGTH)

                End If

            End With

            For j = 3 To 36 ^ 2 - 1

                If blnObjData(j, i) Then

                    PrintLine(lngFFile, "#" & Format(i, "000") & strFromNumZZ(j) & ":" & strObjData(j, i))

                End If

            Next j

            PrintLine(lngFFile)

        Next i

        lngTemp = UBound(blnObjData, 2) + 1

        For i = lngTemp To 999

            With g_Measure(i)

                If .intLen <> MEASURE_LENGTH Then

                    PrintLine(lngFFile, "#" & Format(i, "000") & "02:" & .intLen / MEASURE_LENGTH)

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

        End With

Init:

        FileClose(lngFFile)

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

                    Case 3 * 36 + 1 To 4 * 36 + 9

                        .intCh = .intCh - (2 * 36 + 0)

                    Case 5 * 36 + 1 To 6 * 36 + 9

                        .intCh = .intCh - (4 * 36 + 0)

                    Case 13 * 36 + 1 To 14 * 36 + 9

                        .intCh = .intCh - (12 * 36 + 0)

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
End Module