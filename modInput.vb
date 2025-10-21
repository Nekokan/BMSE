Option Strict Off
Option Explicit On
Module modInput

    Public Enum OBJ_CH
        'BMSのchannelって実は36進数
        CH_NONE = 0
        CH_SPEED = 1033 'SP
        CH_SCROLL = 1020 'SC
        CH_BGM = 1
        CH_MEASURE_LENGTH = 2
        CH_BPM = 3
        CH_BGA = 4
        CH_EXTCHR = 5
        CH_POOR = 6
        CH_LAYER = 7
        CH_EXBPM = 8
        CH_STOP = 9
        CH_BGM_LANE_OFFSET = 36 ^ 2 '36進数の101以上で左から順に各BGM列の仮チャンネル
        CH_BGM_LANE_MAX = CH_BGM_LANE_OFFSET + BGM_LANE
        'keysound type offset
        CH_NM = 0
        CH_INV = 2 * 36
        CH_LN = 4 * 36
        CH_MINE = 12 * 36
        'play side
        CH_1P = 1 * 36
        CH_2P = 2 * 36
        'lane
        CH_KEY1 = 1
        CH_KEY2 = 2
        CH_KEY3 = 3
        CH_KEY4 = 4
        CH_KEY5 = 5
        CH_KEY6 = 8
        CH_KEY7 = 9
        CH_SC = 6
        CH_FZ = 7 'free zone
        ' channel number = type + side + lane
        CH_1P_KEY1 = CH_NM + CH_1P + CH_KEY1
        CH_1P_KEY2 = CH_NM + CH_1P + CH_KEY2
        CH_1P_KEY3 = CH_NM + CH_1P + CH_KEY3
        CH_1P_KEY4 = CH_NM + CH_1P + CH_KEY4
        CH_1P_KEY5 = CH_NM + CH_1P + CH_KEY5
        CH_1P_KEY6 = CH_NM + CH_1P + CH_KEY6
        CH_1P_KEY7 = CH_NM + CH_1P + CH_KEY7
        CH_1P_SC = CH_NM + CH_1P + CH_SC
        CH_1P_FZ = CH_NM + CH_1P + CH_FZ
        CH_2P_KEY1 = CH_NM + CH_2P + CH_KEY1
        CH_2P_KEY2 = CH_NM + CH_2P + CH_KEY2
        CH_2P_KEY3 = CH_NM + CH_2P + CH_KEY3
        CH_2P_KEY4 = CH_NM + CH_2P + CH_KEY4
        CH_2P_KEY5 = CH_NM + CH_2P + CH_KEY5
        CH_2P_KEY6 = CH_NM + CH_2P + CH_KEY6
        CH_2P_KEY7 = CH_NM + CH_2P + CH_KEY7
        CH_2P_SC = CH_NM + CH_2P + CH_SC
        CH_2P_FZ = CH_NM + CH_2P + CH_FZ
        CH_1P_INV_KEY1 = CH_INV + CH_1P + CH_KEY1
        CH_1P_INV_KEY5 = CH_INV + CH_1P + CH_KEY5
        CH_1P_INV_KEY6 = CH_INV + CH_1P + CH_KEY6
        CH_1P_INV_KEY7 = CH_INV + CH_1P + CH_KEY7
        CH_1P_INV_SC = CH_INV + CH_1P + CH_SC
        CH_1P_LN_KEY1 = CH_LN + CH_1P + CH_KEY1
        CH_1P_LN_KEY5 = CH_LN + CH_1P + CH_KEY5
        CH_1P_LN_KEY6 = CH_LN + CH_1P + CH_KEY6
        CH_1P_LN_KEY7 = CH_LN + CH_1P + CH_KEY7
        CH_1P_LN_SC = CH_LN + CH_1P + CH_SC
        CH_1P_MINE_KEY1 = CH_MINE + CH_1P + CH_KEY1
        CH_1P_MINE_KEY5 = CH_MINE + CH_1P + CH_KEY5
        CH_1P_MINE_KEY6 = CH_MINE + CH_1P + CH_KEY6
        CH_1P_MINE_KEY7 = CH_MINE + CH_1P + CH_KEY7
        CH_1P_MINE_SC = CH_MINE + CH_1P + CH_SC
        CH_2P_INV_KEY1 = CH_INV + CH_2P + CH_KEY1
        CH_2P_INV_KEY5 = CH_INV + CH_2P + CH_KEY5
        CH_2P_INV_KEY6 = CH_INV + CH_2P + CH_KEY6
        CH_2P_INV_KEY7 = CH_INV + CH_2P + CH_KEY7
        CH_2P_INV_SC = CH_INV + CH_2P + CH_SC
        CH_2P_LN_KEY1 = CH_LN + CH_2P + CH_KEY1
        CH_2P_LN_KEY5 = CH_LN + CH_2P + CH_KEY5
        CH_2P_LN_KEY6 = CH_LN + CH_2P + CH_KEY6
        CH_2P_LN_KEY7 = CH_LN + CH_2P + CH_KEY7
        CH_2P_LN_SC = CH_LN + CH_2P + CH_SC
        CH_2P_MINE_KEY1 = CH_MINE + CH_2P + CH_KEY1
        CH_2P_MINE_KEY5 = CH_MINE + CH_2P + CH_KEY5
        CH_2P_MINE_KEY6 = CH_MINE + CH_2P + CH_KEY6
        CH_2P_MINE_KEY7 = CH_MINE + CH_2P + CH_KEY7
        CH_2P_MINE_SC = CH_MINE + CH_2P + CH_SC
        CH_KEY_MIN = CH_1P_KEY1
        CH_KEY_MAX = CH_2P_KEY7
        CH_KEY_INV_MIN = CH_INV + CH_KEY_MIN
        CH_KEY_INV_MAX = CH_INV + CH_KEY_MAX
        CH_KEY_LN_MIN = CH_LN + CH_KEY_MIN
        CH_KEY_LN_MAX = CH_LN + CH_KEY_MAX
        CH_KEY_MINE_MIN = CH_MINE + CH_KEY_MIN
        CH_KEY_MINE_MAX = CH_MINE + CH_KEY_MAX
    End Enum

    Public Enum PLAYER_TYPE
        PLAYER_1P = 1
        PLAYER_2P = 2
        PLAYER_DP = 3
        PLAYER_PMS = 4
        PLAYER_OCT = 5
    End Enum

    '判定ランク
    Public Enum PLAY_RANK
        RANK_VERYHARD = 0
        RANK_HARD = 1
        RANK_NORMAL = 2
        RANK_EASY = 3
        RANK_MIN = PLAY_RANK.RANK_VERYHARD
        RANK_MAX = PLAY_RANK.RANK_EASY
    End Enum

    'DIFFICULTY
    Public Enum DIFFICULTY
        NONE = 0
        BEGINNER = 1
        NORMAL = 2
        HYPER = 3
        ANOTHER = 4
        INSANE = 5
        MIN = DIFFICULTY.NONE
        MAX = DIFFICULTY.INSANE
    End Enum

    'LNMODE
    Public Enum LNMODE
        NONE = 0
        LN = 1
        CN = 2
        HCN = 3
        MIN = LNMODE.NONE
        MAX = LNMODE.HCN
    End Enum

    Public Const MATERIAL_MAX As Integer = 3843
    Public Const MEASURE_MAX As Integer = 999
    Public Const MEASURE_LENGTH As Integer = 192 '絶対変えない

    Public Const BGM_LANE As Integer = 128

    Private Const DEFAULT_BPM As Integer = 130
    Private Const DEFAULT_VOLUME As Integer = 1

    Private m_blnUnreadFlag As Boolean
    Private m_strEXInfo As String

    Private m_blnBGM(BGM_LANE * (MEASURE_MAX + 1) - 1) As Boolean

    Private blnSepaDiff As Boolean = False

    Public Structure m_udtMeasure
        Dim intLen As Integer
        Dim lngY As Integer
    End Structure

    Public g_Measure(MEASURE_MAX) As m_udtMeasure

    Public g_strWAV(MATERIAL_MAX) As String
    Public g_strBMP(MATERIAL_MAX) As String
    Public g_strBGA(MATERIAL_MAX) As String

    Private m_sngStop(MATERIAL_MAX) As Single
    Private m_sngBPM(MATERIAL_MAX) As Single
    Private m_sngSCROLL(MATERIAL_MAX) As Single
    Private m_sngSPEED(MATERIAL_MAX) As Single

    Public Sub LoadBMS()
        On Error GoTo Err_Renamed

        'ファイルの存在チェック
        'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
        If Dir(g_BMS.strDir & g_BMS.strFileName, FileAttribute.Normal) = vbNullString Then

            Call MsgBox(g_Message(modMain.Message.ERR_FILE_NOT_FOUND) & vbCrLf & g_Message(modMain.Message.ERR_LOAD_CANCEL), MsgBoxStyle.Critical, g_strAppTitle)

            Exit Sub

        End If

        frmMain.Text = g_strAppTitle & " - Now Loading"

        Call LoadBMSStart()

        Call LoadBMSData()

        Call LoadBMSEnd()

        Exit Sub

Err_Renamed:
        Call modMain.CleanUp(Err.Number, Err.Description, "LoadBMS")
    End Sub

    Public Sub LoadBMSStart()
        On Error GoTo Err_Renamed

        Dim i As Integer

        With frmMain

            For i = 0 To MATERIAL_MAX

                g_strWAV(i) = ""
                g_strBMP(i) = ""
                g_strBGA(i) = ""
                m_sngBPM(i) = 0
                m_sngStop(i) = 0
                m_sngSCROLL(i) = 0

            Next i

            .cboPlayer.SelectedIndex = 0
            .txtGenre.Text = ""
            .txtTitle.Text = ""
            .txtArtist.Text = ""
            .cboPlayLevel.Text = CStr(1)
            .txtBPM.Text = CStr(DEFAULT_BPM)
            .cboPlayRank.SelectedIndex = PLAY_RANK.RANK_EASY
            .txtTotal.Text = ""
            .txtVolume.Text = ""
            .txtStageFile.Text = ""
            .txtMissBMP.Text = ""
            .txtSubTitle.Text = ""
            .txtSubArtist.Text = ""
            .cboDifficulty.SelectedIndex = 0
            .txtPreview.Text = ""
            .txtBanner.Text = ""
            .txtBackBmp.Text = ""
            .cboLNMode.SelectedIndex = 0
            .cboLNObj.SelectedIndex = 0
            .txtDefExRank.Text = ""
            .txtComment.Text = ""
            .lstWAV.SelectedIndex = 0
            .lstBMP.SelectedIndex = 0
            .lstBGA.SelectedIndex = 0
            .lstMeasureLen.SelectedIndex = 0
            .lstMeasureLen.Visible = False
            .txtExInfo.Text = ""
            .Enabled = False

            '.vsbMain.Value = .vsbMain.Maximum - frmMain.vsbMain.LargeChange + 1
            '.hsbMain.Value = 0
            .cboVScroll.SelectedIndex = .cboVScroll.Items.Count - 2

            For i = 0 To MEASURE_MAX

                g_Measure(i).intLen = MEASURE_LENGTH
                modMain.SetItemString(.lstMeasureLen, i, "#" & Format(i, "000") & ":4/4")

            Next i

        End With

        With g_BMS

            .intPlayerType = PLAYER_TYPE.PLAYER_1P
            .strGenre = ""
            .strTitle = ""
            .strArtist = ""
            .sngBPM = DEFAULT_BPM
            .lngPlayLevel = 1
            .intPlayRank = PLAY_RANK.RANK_EASY
            .sngTotal = 0
            .intVolume = 0
            .strStageFile = ""
            .strSubTitle = ""
            .strSubArtist = ""
            .intDifficulty = 0
            .strPreview = ""
            .strBanner = ""
            .strBackBMP = ""
            .intLNMode = 0
            .intLNObj = 0
            .intDefExRank = 0
            .strComment = ""
        End With

        blnSepaDiff = False
        g_disp.intMaxMeasure = 0
        Call modDraw.lngChangeMaxMeasure(15)
        Call modDraw.ChangeResolution()

        Call g_InputLog.Clear()

        ReDim g_Obj(0)
        ReDim g_lngObjID(0)
        g_lngIDNum = 0

        m_blnUnreadFlag = False
        m_strEXInfo = ""

        ReDim m_blnBGM(BGM_LANE * (MEASURE_MAX + 1) - 1)

        For i = 0 To UBound(m_blnBGM)

            m_blnBGM(i) = False

        Next i

        Exit Sub

Err_Renamed:
        Call modMain.CleanUp(Err.Number, Err.Description, "LoadBMSStart")
    End Sub

    Public Sub LoadBMSEnd()
        On Error GoTo Err_Renamed

        If blnSepaDiff Then Call MsgBox(g_Message(modMain.Message.ERR_POSITION_ROUNDED), MsgBoxStyle.Exclamation, g_strAppTitle)

        With frmMain

            Call modEasterEgg.LoadEffect()

            Call frmMain.RefreshList()

            .lstMeasureLen.Visible = True

            Call modDraw.ChangeResolution()

            .Enabled = True

            If UCase(Right(g_BMS.strFileName, 3)) = "PMS" Then

                .cboPlayer.SelectedIndex = 3
                g_BMS.intPlayerType = 4

            End If

            m_blnUnreadFlag = False
            .txtExInfo.Text = m_strEXInfo
            m_strEXInfo = ""

        End With

        g_BMS.blnSaveFlag = True

        Call modDraw.InitVerticalLine()

        With frmMain

            If Len(g_BMS.strDir) Then

                If ._mnuOptionsItem_1.Checked Then

                    .Text = g_strAppTitle & " - " & g_BMS.strFileName

                Else

                    .Text = g_strAppTitle & " - " & g_BMS.strDir & g_BMS.strFileName

                End If

            Else

                .Text = g_strAppTitle

            End If

            .vsbMain.Value = .vsbMain.Maximum - frmMain.vsbMain.LargeChange + 1
            .hsbMain.Value = 0

            Call .Show()

            Call .picMain.Focus()

        End With

        Exit Sub

Err_Renamed:
        Call modMain.CleanUp(Err.Number, Err.Description, "LoadBMSEnd")
    End Sub

    Private Sub LoadBMSData()
        On Error GoTo Err_Renamed

        Dim i As Integer
        Dim strArray() As String
        Dim strTemp As String
        Dim lngFFile As Integer

        lngFFile = FreeFile()

        FileOpen(lngFFile, g_BMS.strDir & g_BMS.strFileName, OpenMode.Input)

        Do While Not EOF(lngFFile)

            System.Windows.Forms.Application.DoEvents()

            strTemp = LineInput(lngFFile)

            strArray = Split(Replace(Replace(strTemp, vbCr, vbCrLf), vbLf, vbCrLf), vbCrLf)

            For i = 0 To UBound(strArray)

                If Left(strArray(i), 1) = "#" Then Call LoadBMSLine(strArray(i))

            Next i

        Loop

        FileClose(lngFFile)

        ReDim Preserve g_Obj(UBound(g_Obj))

        For i = 0 To UBound(g_Obj) - 1

            With g_Obj(i)

                .lngPosition = (g_Measure(.intMeasure).intLen / .lngHeight) * .lngPosition

                If .intCh = OBJ_CH.CH_BPM Then 'BPM

                    .intCh = OBJ_CH.CH_EXBPM

                ElseIf .intCh = OBJ_CH.CH_EXBPM Then  '拡張BPM

                    If m_sngBPM(.sngValue) = 0 Then

                        .intCh = OBJ_CH.CH_NONE

                    Else

                        .sngValue = m_sngBPM(.sngValue)

                    End If

                ElseIf .intCh = OBJ_CH.CH_STOP Then  'ストップシーケンス

                    .sngValue = m_sngStop(.sngValue)

                ElseIf .intCh = OBJ_CH.CH_SCROLL Then  'SCROLL

                    .sngValue = m_sngSCROLL(.sngValue)

                ElseIf .intCh = OBJ_CH.CH_SPEED Then  'SPEED

                    .sngValue = m_sngSPEED(.sngValue)

                End If

            End With

        Next i

        'Call QuickSort(0, UBound(g_Obj))

        Exit Sub

Err_Renamed:
        Call modMain.CleanUp(Err.Number, Err.Description, "LoadBMSData")
    End Sub

    Public Sub LoadBMSLine(ByRef strLineData As String, Optional ByVal blnDirectInput As Boolean = False)
        On Error GoTo Err_Renamed

        Dim strArray() As String
        Dim strFunc As String
        Dim strFuncUC As String
        Dim strParam As String

        strArray = Split(Replace(strLineData, " ", ":", 1, 1), ":")

        If UBound(strArray) > 0 Then

            strFunc = strArray(0)
            strFuncUC = UCase(strFunc)
            strParam = Mid(strLineData, Len(strFunc) + 2)

            Select Case strFuncUC

                Case "#IF", "#RANDOM", "#RONDAM"

                    If blnDirectInput = False Then

                        m_blnUnreadFlag = True

                        m_strEXInfo = m_strEXInfo & strLineData & vbCrLf

                    End If

                Case Else

                    If m_blnUnreadFlag = False Then

                        If LoadBMSHeader(strFunc, strParam, blnDirectInput) = False Then

                            If LoadBMSObject(strFunc, strParam) = False Then

                                m_strEXInfo = m_strEXInfo & strLineData & vbCrLf

                            End If

                        End If

                    Else

                        m_strEXInfo = m_strEXInfo & strLineData & vbCrLf

                    End If

            End Select

        ElseIf UCase(strArray(0)) = "#ENDIF" Then

            m_blnUnreadFlag = False

            m_strEXInfo = m_strEXInfo & strLineData & vbCrLf

        Else

            m_strEXInfo = m_strEXInfo & strLineData & vbCrLf

        End If

        Exit Sub

Err_Renamed:
        Call modMain.CleanUp(Err.Number, Err.Description, "LoadBMSLine")
    End Sub

    Private Function LoadBMSHeader(ByRef strFunc As String, ByRef strParam As String, Optional ByVal blnDirectInput As Boolean = False) As Boolean
        On Error GoTo Err_Renamed

        Dim lngNum As Integer
        Dim strFuncUC = UCase(strFunc) 'コマンドを大文字に統一して判定させる

        modMain.bln62AutoSwiched = False

        With frmMain

            Select Case strFuncUC

                'Case "#PATH_WAV"

                'g_BMS.strDir = strParam

                Case "#PLAYER"

                    g_BMS.intPlayerType = Val(strParam)
                    .cboPlayer.SelectedIndex = Val(strParam) - 1

                Case "#GENRE", "#GENLE"

                    g_BMS.strGenre = strParam
                    .txtGenre.Text = strParam

                Case "#TITLE"
                    g_BMS.strTitle = strParam
                    .txtTitle.Text = strParam

                Case "#ARTIST"

                    g_BMS.strArtist = strParam
                    .txtArtist.Text = strParam

                Case "#BPM"

                    g_BMS.sngBPM = Val(strParam)
                    .txtBPM.Text = CStr(Val(strParam))

                Case "#PLAYLEVEL"

                    g_BMS.lngPlayLevel = Val(strParam)
                    .cboPlayLevel.Text = CStr(Val(strParam))

                Case "#RANK"

                    g_BMS.intPlayRank = Val(strParam)

                    If g_BMS.intPlayRank < PLAY_RANK.RANK_MIN Then g_BMS.intPlayRank = PLAY_RANK.RANK_MIN

                    If g_BMS.intPlayRank > PLAY_RANK.RANK_MAX Then g_BMS.intPlayRank = PLAY_RANK.RANK_MAX

                    .cboPlayRank.SelectedIndex = g_BMS.intPlayRank

                Case "#TOTAL"

                    g_BMS.sngTotal = Val(strParam)
                    .txtTotal.Text = CStr(Val(strParam))

                Case "#VOLWAV"

                    g_BMS.intVolume = Val(strParam)
                    .txtVolume.Text = CStr(Val(strParam))

                Case "#BASE"

                    If CInt(strParam) = 62 Then
                        frmMain._mnuOptionsBase16.Checked = False
                        frmMain._mnuOptionsBase36.Checked = False
                        frmMain._mnuOptionsBase62.Checked = True
                        modMain.bln62AutoSwiched = True
                    Else
                        'frmMain._mnuOptionsBase62.Checked = False
                    End If

                Case "#STAGEFILE"

                    g_BMS.strStageFile = strParam
                    .txtStageFile.Text = strParam

                Case "#SUBTITLE"

                    g_BMS.strSubTitle = strParam
                    .txtSubTitle.Text = strParam

                Case "#SUBARTIST"

                    g_BMS.strSubArtist = strParam
                    .txtSubArtist.Text = strParam

                Case "#DIFFICULTY"

                    g_BMS.intDifficulty = Val(strParam)

                    If g_BMS.intDifficulty < DIFFICULTY.MIN Then g_BMS.intDifficulty = DIFFICULTY.MIN
                    If g_BMS.intDifficulty > DIFFICULTY.MAX Then g_BMS.intDifficulty = DIFFICULTY.MAX

                    .cboDifficulty.SelectedIndex = g_BMS.intDifficulty

                Case "#PREVIEW"

                    g_BMS.strPreview = strParam
                    .txtPreview.Text = strParam

                Case "#BANNER"

                    g_BMS.strBanner = strParam
                    .txtBanner.Text = strParam

                Case "#BACKBMP"

                    g_BMS.strBackBMP = strParam
                    .txtBackBmp.Text = strParam

                Case "#LNMODE"

                    g_BMS.intLNMode = Val(strParam)

                    If g_BMS.intLNMode < LNMODE.MIN Then g_BMS.intLNMode = LNMODE.MIN
                    If g_BMS.intLNMode > LNMODE.MAX Then g_BMS.intLNMode = LNMODE.MAX

                    .cboLNMode.SelectedIndex = g_BMS.intLNMode

                Case "#LNOBJ"

                    g_BMS.intLNObj = strToNum(strParam)

                Case "#DEFEXRANK"

                    g_BMS.intDefExRank = Val(strParam)
                    .txtDefExRank.Text = CStr(Val(strParam))

                Case "#COMMENT"

                    '#COMMENTはダブルクオーテーション Chr(34) 必須のための処理
                    strParam = IIf(Left(strParam, 1) = Chr(34) And Right(strParam, 1) = Chr(34), strParam, Chr(34) & strParam & Chr(34))

                    g_BMS.strComment = strParam
                    .txtComment.Text = strParam

                Case Else

                    '定義番号部：strFuncUCは大文字化して62進数でなくなっているのでstrFuncの方を使う
                    lngNum = IIf(frmMain._mnuOptionsBase62.Checked, strToNum62ZZ(Right(strFunc, 2)), IIf(frmMain._mnuOptionsBase16.Checked, strToNumFF(Right(strFunc, 2)), strToNumZZ(Right(strFunc, 2))))

                    Select Case Left(strFuncUC, Len(strFuncUC) - 2)

                        Case "#WAV"

                            If blnDirectInput = False Then

                                If lngNum <> 0 Then

                                    g_strWAV(lngNum) = strParam

                                    '                            If Asc(left$(strTemp, 1)) > Asc("F") Or Asc(right$(strTemp, 1)) > Asc("F") Then
                                    '
                                    '                                .mnuOptionsItem(USE_OLD_FORMAT).Checked = False
                                    '
                                    '                            End If

                                Else

                                    .txtLandmineWAV.Text = strParam

                                End If

                            End If


                        Case "#BMP"

                            If blnDirectInput = False Then

                                If lngNum <> 0 Then

                                    g_strBMP(lngNum) = strParam

                                    '                                If Asc(left$(strTemp, 1)) > Asc("F") Or Asc(right$(strTemp, 1)) > Asc("F") Then
                                    '
                                    '                                    .mnuOptionsItem(USE_OLD_FORMAT).Checked = False
                                    '
                                    '                                End If

                                Else

                                    .txtMissBMP.Text = strParam

                                End If

                            End If

                        Case "#BGA"

                            If lngNum <> 0 And blnDirectInput = False Then

                                g_strBGA(lngNum) = strParam

                                '                            If Asc(left$(strTemp, 1)) > Asc("F") Or Asc(right$(strTemp, 1)) > Asc("F") Then
                                '
                                '                                .mnuOptionsItem(USE_OLD_FORMAT).Checked = False
                                '
                                '                            End If

                            End If

                        Case "#BPM"

                            If lngNum <> 0 And blnDirectInput = False Then

                                m_sngBPM(lngNum) = CSng(strParam)

                            End If

                        Case "#STOP"

                            If lngNum <> 0 And blnDirectInput = False Then

                                m_sngStop(lngNum) = CSng(strParam)

                            End If

                        Case "#SCROLL"

                            If lngNum <> 0 And blnDirectInput = False Then

                                m_sngSCROLL(lngNum) = CSng(strParam)

                            End If

                        Case "#SPEED"

                            If lngNum <> 0 And blnDirectInput = False Then

                                m_sngSPEED(lngNum) = CSng(strParam)

                            End If

                        Case Else

                            LoadBMSHeader = True

                    End Select

            End Select

        End With

        LoadBMSHeader = Not LoadBMSHeader

        Exit Function

Err_Renamed:
        Call modMain.CleanUp(Err.Number, Err.Description, "LoadBMSHeader")
    End Function

    Private Function LoadBMSObject(ByRef strFunc As String, ByRef strParam As String) As Boolean
        On Error GoTo Err_Renamed

        Dim i As Integer
        Dim j As Integer
        Dim intTemp As Integer
        Dim intMeasure As Integer
        Dim intCh As Integer
        Dim lngSepaNum As Integer
        Dim Value As String = Space(2)

        intMeasure = Val(Mid(strFunc, 2, 3))
        intCh = strToNumZZ(Mid(strFunc, 5, 2))

        lngSepaNum = Len(strParam) \ 2

        If intCh = OBJ_CH.CH_MEASURE_LENGTH Then

            If Val(strParam) = 0 Then Exit Function

            '以下小節長を分数に変換する処理、微妙に怪しい。
            '例えば0.00520833333333333 は 1/192 ではなく 2/384 と変換される。もちろんそれらは等しいのだけど。
            '念のため対応表
            '	小節長	->	192*小節長		->	最大公約数
            '	strParam		intLen		intTemp
            '	32				12288		384->96
            '	16				6144		384->96 (*旧仕様)
            '	2(8/4)			768			384->96
            '	6/4				576			192->96
            '	4/4				384			384->96
            '	1/3(0.333..)	128			128->96
            '	1/4				96			96
            '	1/64			6				6
            '	1/96			4				4
            '	1/192			2				2

            intTemp = intGCD(Int(MEASURE_LENGTH * Val(strParam)), MEASURE_LENGTH) '192*小節長と192の最大公約数

            If intTemp <= 1 Then intTemp = 1

            If intTemp >= 48 Then intTemp = 48

            With g_Measure(intMeasure)

                .intLen = CInt(MEASURE_LENGTH * Val(strParam))

                If .intLen < 1 Then .intLen = 1 '小節長1/192未満は小節長1/192へ

                Do While .intLen \ intTemp > 4 * 128 '最大小節長128

                    If intTemp >= 48 Then '(intLen > 4 * 128 * 48 のとき)

                        .intLen = 4 * 128 * 48

                        Exit Do

                    End If

                    intTemp = intTemp * 2

                Loop

                modMain.SetItemString(frmMain.lstMeasureLen, intMeasure, "#" & Format(intMeasure, "000") & ":" & (.intLen \ intTemp) & "/" & (MEASURE_LENGTH \ intTemp))

            End With

        Else

            If intCh = OBJ_CH.CH_BGM Then

                For j = 0 To BGM_LANE - 1

                    If m_blnBGM(intMeasure * BGM_LANE + j) = False Then

                        m_blnBGM(intMeasure * BGM_LANE + j) = True
                        intTemp = (36 ^ 2 + 1) + j '101+j

                        Exit For

                    End If

                Next j

            End If

            For i = 1 To lngSepaNum

                Value = Mid(strParam, i * 2 - 1, 2)

                If Value <> "00" Then

                    With g_Obj(UBound(g_Obj))

                        .lngID = g_lngIDNum
                        g_lngObjID(g_lngIDNum) = g_lngIDNum
                        .lngPosition = i - 1
                        .lngHeight = lngSepaNum
                        .intMeasure = intMeasure
                        .intCh = intCh

                        Call modDraw.lngChangeMaxMeasure(.intMeasure)

                        Select Case intCh

                            Case OBJ_CH.CH_BGM 'BGM

                                .sngValue = IIf(frmMain._mnuOptionsBase62.Checked, modInput.strToNum62ZZ(Value), IIf(frmMain._mnuOptionsBase16.Checked, modInput.strToNumFF(Value), modInput.strToNumZZ(Value)))
                                .intCh = intTemp

                            Case OBJ_CH.CH_BGA, OBJ_CH.CH_POOR, OBJ_CH.CH_LAYER, OBJ_CH.CH_EXBPM, OBJ_CH.CH_STOP, OBJ_CH.CH_SCROLL, OBJ_CH.CH_SPEED 'BGA,Poor,Layer,拡張BPM,ストップシーケンス,SCROLL,SPEED

                                .sngValue = IIf(frmMain._mnuOptionsBase62.Checked, modInput.strToNum62ZZ(Value), IIf(frmMain._mnuOptionsBase16.Checked, modInput.strToNumFF(Value), modInput.strToNumZZ(Value)))

                            Case OBJ_CH.CH_BPM 'BPM

                                .sngValue = Val("&H" & Value)

                            Case OBJ_CH.CH_1P_KEY1 To OBJ_CH.CH_1P_SC, OBJ_CH.CH_1P_KEY6, OBJ_CH.CH_1P_KEY7, OBJ_CH.CH_2P_KEY1 To OBJ_CH.CH_2P_SC, OBJ_CH.CH_2P_KEY6, OBJ_CH.CH_2P_KEY7 'キー音

                                .sngValue = IIf(frmMain._mnuOptionsBase62.Checked, modInput.strToNum62ZZ(Value), IIf(frmMain._mnuOptionsBase16.Checked, modInput.strToNumFF(Value), modInput.strToNumZZ(Value)))

                            Case OBJ_CH.CH_1P_INV_KEY1 To OBJ_CH.CH_1P_INV_SC, OBJ_CH.CH_1P_INV_KEY6, OBJ_CH.CH_1P_INV_KEY7, OBJ_CH.CH_2P_INV_KEY1 To OBJ_CH.CH_2P_INV_SC, OBJ_CH.CH_2P_INV_KEY6, OBJ_CH.CH_2P_INV_KEY7 'キー音 (INV)

                                .sngValue = IIf(frmMain._mnuOptionsBase62.Checked, modInput.strToNum62ZZ(Value), IIf(frmMain._mnuOptionsBase16.Checked, modInput.strToNumFF(Value), modInput.strToNumZZ(Value)))
                                .intCh = .intCh - OBJ_CH.CH_INV
                                .intAtt = modMain.OBJ_ATT.OBJ_INVISIBLE

                            Case OBJ_CH.CH_1P_INV_KEY1 To OBJ_CH.CH_1P_LN_SC, OBJ_CH.CH_1P_LN_KEY6, OBJ_CH.CH_1P_LN_KEY7, OBJ_CH.CH_2P_LN_KEY1 To OBJ_CH.CH_2P_LN_SC, OBJ_CH.CH_2P_LN_KEY6, OBJ_CH.CH_2P_LN_KEY7  'キー音 (LN)

                                .sngValue = IIf(frmMain._mnuOptionsBase62.Checked, modInput.strToNum62ZZ(Value), IIf(frmMain._mnuOptionsBase16.Checked, modInput.strToNumFF(Value), modInput.strToNumZZ(Value)))
                                .intCh = .intCh - OBJ_CH.CH_LN
                                .intAtt = modMain.OBJ_ATT.OBJ_LONGNOTE

                            Case OBJ_CH.CH_1P_MINE_KEY1 To OBJ_CH.CH_1P_MINE_SC, OBJ_CH.CH_1P_MINE_KEY6, OBJ_CH.CH_1P_MINE_KEY7, OBJ_CH.CH_2P_MINE_KEY1 To OBJ_CH.CH_2P_MINE_SC, OBJ_CH.CH_2P_MINE_KEY6, OBJ_CH.CH_2P_MINE_KEY7  'キー音 (地雷)

                                .sngValue = strToNumZZ(Value) ' 地雷は36進数（でなければいけないはず；なぜならZZを最大としているため）
                                If .sngValue > 1295 Then .sngValue = 1295
                                .intCh = .intCh - OBJ_CH.CH_MINE
                                .intAtt = modMain.OBJ_ATT.OBJ_MINE

                            Case Else

                                Exit Function

                        End Select

                    End With

                    ReDim Preserve g_Obj(UBound(g_Obj) + 1)

                    g_lngIDNum = g_lngIDNum + 1
                    ReDim Preserve g_lngObjID(g_lngIDNum)

                End If

            Next i

        End If

        If lngSepaNum <> 0 AndAlso g_Measure(intMeasure).intLen Mod lngSepaNum <> 0 Then blnSepaDiff = True

        LoadBMSObject = True

        Exit Function

Err_Renamed:
        Call modMain.CleanUp(Err.Number, Err.Description, "LoadBMSObject")
    End Function

    Public Sub QuickSort(ByVal l As Integer, ByVal r As Integer)

        Dim i As Integer
        Dim j As Integer
        Dim p As Single

        p = g_Obj((l + r) \ 2).lngPosition
        i = l
        j = r

        Do

            Do While g_Obj(i).lngPosition < p

                i = i + 1

            Loop

            Do While g_Obj(j).lngPosition > p

                j = j - 1

            Loop

            If i >= j Then Exit Do

            Call SwapObj(i, j)

            i = i + 1
            j = j - 1

        Loop

        If (l < i - 1) Then Call QuickSort(l, i - 1)
        If (r > j + 1) Then Call QuickSort(j + 1, r)

    End Sub

    Public Sub SwapObj(ByVal Obj1Num As Integer, ByVal Obj2Num As Integer)

        Dim dummy As g_udtObj

        g_lngObjID(g_Obj(Obj1Num).lngID) = Obj2Num
        g_lngObjID(g_Obj(Obj2Num).lngID) = Obj1Num

        dummy = g_Obj(Obj1Num)
        g_Obj(Obj1Num) = g_Obj(Obj2Num)
        g_Obj(Obj2Num) = dummy

    End Sub

    Public Sub SwapObj(Obj() As g_udtObj, ByVal Obj1Num As Integer, ByVal Obj2Num As Integer)

        Dim dummy As g_udtObj

        g_lngObjID(Obj(Obj1Num).lngID) = Obj2Num
        g_lngObjID(Obj(Obj2Num).lngID) = Obj1Num

        dummy = Obj(Obj1Num)
        Obj(Obj1Num) = Obj(Obj2Num)
        Obj(Obj2Num) = dummy

    End Sub

    Public Function strToNum(ByRef strNum As String) As Integer

        If frmMain._mnuOptionsBase62.Checked Then

            strToNum = strToNum62ZZ(strNum)

        ElseIf frmMain._mnuOptionsBase16.Checked Then

            strToNum = strToNumFF(strNum)

        Else

            strToNum = strToNumZZ(strNum)

        End If

    End Function

    Public Function strToNumZZ(ByRef strNum As String) As Integer

        Dim i As Integer
        Dim ret As Integer

        For i = 1 To Len(strNum)

            ret = ret + subStrToNumZZ(Mid(strNum, i, 1)) * (36 ^ (Len(strNum) - i))

        Next i

        strToNumZZ = ret

    End Function

    Public Function subStrToNumZZ(ByRef b As String) As Integer

        Dim r As Integer

        r = System.Math.Abs(Asc(UCase(b)))

        If r >= 65 And r <= 90 Then 'A-Z

            subStrToNumZZ = r - 55

        Else

            subStrToNumZZ = (r - 48) Mod 36

        End If

    End Function

    Public Function strToNum62ZZ(ByRef strNum As String) As Integer

        Dim i As Integer
        Dim ret As Integer

        For i = 1 To Len(strNum)

            ret = ret + subStrToNum62ZZ(Mid(strNum, i, 1)) * (62 ^ (Len(strNum) - i))

        Next i

        strToNum62ZZ = ret

    End Function

    Public Function subStrToNum62ZZ(ByRef b As String) As Integer

        Dim r As Integer

        r = System.Math.Abs(Asc((b)))

        If r >= 65 And r <= 90 Then 'A-Z

            subStrToNum62ZZ = r - 55 ' 10-35

        ElseIf r >= 97 And r <= 122 Then 'a-z

            subStrToNum62ZZ = r - 61 ' 36-61

        ElseIf r >= 48 And r <= 57 Then ' 0-9

            subStrToNum62ZZ = (r - 48) Mod 62 ' 0-9

        Else

            subStrToNum62ZZ = 0

        End If

    End Function

    Public Function strToNumFF(ByRef strNum As String) As Integer

        Dim ret As Integer
        Dim l As String '右2文字以外
        Dim m As String '右から2文字目
        Dim r As String '右端1文字

        r = Right(strNum, 1)

        If Len(strNum) > 2 Then

            l = Left(strNum, Len(strNum) - 2)

        Else

            l = "0"

        End If

        If Len(strNum) > 1 Then

            m = Mid(strNum, Len(strNum) - 1, 1)

        Else

            m = "0"

        End If

        If l = 0 Then

            If Asc(m) <= 70 Then 'F 以下

                If Asc(r) <= 70 Then 'FF 以下

                    ret = Val("&H" & m & r)

                Else 'FZ 以下

                    ret = Val("&H" & m) * 20 + subStrToNumFF(r)

                End If

            ElseIf Asc(m) >= 65 And Asc(m) <= 90 Then  'G0-ZZ

                ret = strToNumZZ(m & r)

            Else

                ret = 0

            End If

        Else

            ret = strToNumZZ(l & m & r)

        End If

        strToNumFF = ret

    End Function

    Private Function subStrToNumFF(ByRef b As String) As Integer

        Dim r As Integer

        r = Asc(b)

        If r >= 65 And r <= 70 Then 'A-F

            subStrToNumFF = r - 55 ' 10-16

        ElseIf r >= 71 And r <= 90 Then 'G-Z

            subStrToNumFF = r + 185 ' 256-275

        ElseIf 48 <= r And r <= 58 Then ' 0-9

            subStrToNumFF = (r - 48) Mod 36 ' 0-9

        Else

            subStrToNumFF = 0

        End If

    End Function

    Public Function strFromNum(ByVal lngNum As Integer, Optional ByVal Length As Integer = 2) As String

        If frmMain._mnuOptionsBase62.Checked Then

            strFromNum = strFromNum62ZZ(lngNum, Length)

        ElseIf frmMain._mnuOptionsBase16.Checked Then

            strFromNum = strFromNumFF(lngNum, Length)

        Else

            strFromNum = strFromNumZZ(lngNum, Length)

        End If

    End Function

    Public Function strFromNumZZ(ByVal lngNum As Integer, Optional ByVal Length As Integer = 2) As String

        Dim strTemp As String = ""

        Do While lngNum

            strTemp = subStrFromNumZZ(lngNum Mod 36) & strTemp
            lngNum = lngNum \ 36

        Loop

        Do While Len(strTemp) < Length

            strTemp = "0" & strTemp

        Loop

        strFromNumZZ = Right(strTemp, Length)

    End Function

    Public Function subStrFromNumZZ(ByVal b As Integer) As String

        Select Case b

            Case 0 To 9

                subStrFromNumZZ = CStr(b)

            Case Else

                subStrFromNumZZ = Chr(b + 55)

        End Select

    End Function


    Public Function strFromNum62ZZ(ByVal lngNum As Integer, Optional ByVal Length As Integer = 2) As String

        Dim strTemp As String = ""

        Do While lngNum

            strTemp = subStrFromNum62ZZ(lngNum Mod 62) & strTemp
            lngNum = lngNum \ 62

        Loop

        Do While Len(strTemp) < Length

            strTemp = "0" & strTemp

        Loop

        strFromNum62ZZ = Right(strTemp, Length)

    End Function

    Public Function subStrFromNum62ZZ(ByVal b As Integer) As String

        Select Case b

            Case 0 To 9

                subStrFromNum62ZZ = CStr(b)

            Case 10 To 35

                subStrFromNum62ZZ = Chr(b + 55) ' 65-90 = A-Z

            Case 36 To 61

                subStrFromNum62ZZ = Chr(b + 61) ' 97-122 = a-z

            Case Else

                Return ""

        End Select

    End Function

    Public Function strFromNumFF(ByVal lngNum As Integer, Optional ByVal Length As Integer = 2) As String

        If frmMain._mnuOptionsBase16.Checked Then

            Select Case lngNum

                Case Is < 256 '～FF

                    strFromNumFF = Right(New String("0", Length) & Hex(lngNum), Length)

                Case Is < 576 '～0G-FZ

                    lngNum = lngNum - 256
                    strFromNumFF = Hex(lngNum \ 20) & subStrFromNumZZ((lngNum Mod 20) + 16)
                    strFromNumFF = Right(New String("0", Length) & strFromNumFF, Length)

                Case Is < 1296 '～ZZ

                    lngNum = lngNum - 576
                    strFromNumFF = subStrFromNumZZ(lngNum \ 36 + 16) & subStrFromNumZZ(lngNum Mod 36)
                    strFromNumFF = Right(New String("0", Length) & strFromNumFF, Length)

                Case Else

                    strFromNumFF = strFromNumZZ(lngNum, Length)

            End Select

        ElseIf frmMain._mnuOptionsBase36.Checked Then

            strFromNumFF = strFromNumZZ(lngNum, Length)

        Else 'frmMain._mnuOptionsBase62.Checked Then

            strFromNumFF = strFromNum62ZZ(lngNum, Length)

        End If

    End Function


    Public Function intGCD(ByVal m As Integer, ByVal n As Integer) As Integer

        If m <= 0 Or n <= 0 Then
            intGCD = 1
            Exit Function
        End If

        If m Mod n = 0 Then

            intGCD = n

        Else

            intGCD = intGCD(n, m Mod n)

        End If

    End Function
End Module
