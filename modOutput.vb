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
        Dim lngMaxMeasure As Integer
        Dim lngTemp As Integer
        Dim strTemp As String
        Dim intArray() As Integer
        Dim lngStop(MATERIAL_MAX) As Integer
        Dim sngBPM(MATERIAL_MAX) As Single

        If Flag = 0 Then frmMain.Text = g_strAppTitle & " - Now Saving..."

        frmMain.Enabled = False

        For i = 0 To MATERIAL_MAX

            sngBPM(i) = 0
            lngStop(i) = 0

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

                                intBPMNum = intBPMNum + 1
                                sngBPM(intBPMNum) = .sngValue
                                .sngValue = intBPMNum

                            End If

                        Case modInput.OBJ_CH.CH_STOP

                            If intSTOPNum > MATERIAL_MAX Then

                                Call MsgBox(g_Message(modMain.Message.ERR_OVERFLOW_STOP) & vbCrLf & g_Message(modMain.Message.ERR_SAVE_CANCEL), MsgBoxStyle.Critical, g_strAppTitle)

                                lngTemp = i - 1

                                GoTo Init

                            End If

                            intSTOPNum = intSTOPNum + 1
                            lngStop(intSTOPNum) = .sngValue
                            .sngValue = intSTOPNum

                        Case 11 To 29

                            If .intAtt = modMain.OBJ_ATT.OBJ_INVISIBLE Then

                                .intCh = .intCh + 20

                            ElseIf .intAtt = modMain.OBJ_ATT.OBJ_LONGNOTE Then

                                .intCh = .intCh + 40

                            End If

                    End Select

                End If

            End With

        Next i

        ReDim strObjData(100 + modInput.BGM_LANE, lngMaxMeasure)
        ReDim blnObjData(100 + modInput.BGM_LANE, lngMaxMeasure)

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

                    Case Is > 1000

                    Case Is > 100

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), .lngPosition * 2) & modInput.strFromNum(.sngValue) & Mid(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)

                        For j = 101 To .intCh - 1

                            blnObjData(j, .intMeasure) = True

                        Next j

                    Case modInput.OBJ_CH.CH_BPM

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), .lngPosition * 2) & Right("0" & Hex(.sngValue), 2) & Mid(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)

                    Case modInput.OBJ_CH.CH_EXBPM

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), .lngPosition * 2) & Right("0" & IIf(intBPMNum > 255, modInput.strFromNum(.sngValue), Hex(.sngValue)), 2) & Mid(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)

                    Case modInput.OBJ_CH.CH_STOP

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), .lngPosition * 2) & Right("0" & IIf(intSTOPNum > 255, modInput.strFromNum(.sngValue), Hex(.sngValue)), 2) & Mid(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)

                    Case Else

                        strObjData(.intCh, .intMeasure) = Left(strObjData(.intCh, .intMeasure), .lngPosition * 2) & modInput.strFromNum(.sngValue) & Mid(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)

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

            PrintLine(lngFFile, "#STAGEFILE " & Trim(.txtStageFile.Text))
            PrintLine(lngFFile)

            For i = 1 To 1295

                If Len(g_strWAV(i)) Then

                    PrintLine(lngFFile, "#WAV" & modInput.strFromNum(i) & " " & g_strWAV(i))

                End If

            Next i

            PrintLine(lngFFile)

            If Len(Trim(.txtMissBMP.Text)) Then

                PrintLine(lngFFile, "#BMP00 " & .txtMissBMP.Text)

            End If

            For i = 1 To 1295

                If Len(g_strBMP(i)) Then

                    PrintLine(lngFFile, "#BMP" & modInput.strFromNum(i) & " " & g_strBMP(i))

                End If

            Next i

            PrintLine(lngFFile)

            For i = 1 To 1295

                If Len(g_strBGA(i)) Then

                    PrintLine(lngFFile, "#BGA" & modInput.strFromNum(i) & " " & g_strBGA(i))

                End If

            Next i

            PrintLine(lngFFile)

            If intBPMNum > 255 Then

                For i = 1 To 1295

                    If sngBPM(i) Then

                        PrintLine(lngFFile, "#BPM" & Right("0" & modInput.strFromNum(i), 2) & " " & sngBPM(i))

                    End If

                Next i

            ElseIf intBPMNum Then

                For i = 1 To 255

                    If sngBPM(i) Then

                        PrintLine(lngFFile, "#BPM" & Right("0" & Hex(i), 2) & " " & sngBPM(i))

                    End If

                Next i

            End If

            PrintLine(lngFFile)

            If intSTOPNum > 255 Then

                For i = 1 To MATERIAL_MAX

                    If lngStop(i) Then

                        PrintLine(lngFFile, "#STOP" & Right("0" & modInput.strFromNum(i), 2) & " " & lngStop(i))

                    End If

                Next i

            ElseIf intSTOPNum Then

                For i = 1 To 255

                    If lngStop(i) Then

                        PrintLine(lngFFile, "#STOP" & Right("0" & Hex(i), 2) & " " & lngStop(i))

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

            For j = 101 To 101 + modInput.BGM_LANE - 1

                If blnObjData(j, i) Then

                    PrintLine(lngFFile, "#" & Format(i, "000") & "01" & ":" & strObjData(j, i))

                End If

            Next j

            With g_Measure(i)

                If .intLen <> MEASURE_LENGTH Then

                    PrintLine(lngFFile, "#" & Format(i, "000") & "02:" & .intLen / MEASURE_LENGTH)

                End If

            End With

            For j = 3 To 99

                If blnObjData(j, i) Then

                    PrintLine(lngFFile, "#" & Format(i, "000") & Format(j, "00") & ":" & strObjData(j, i))

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

                        .sngValue = lngStop(.sngValue)

                    Case 31 To 49

                        .intCh = .intCh - 20

                    Case 51 To 69

                        .intCh = .intCh - 40

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