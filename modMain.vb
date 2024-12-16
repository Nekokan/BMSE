Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Text
Imports System.Runtime.InteropServices

Module modMain

#Const MODE_DEBUG = True

#If MODE_DEBUG = True Then

    Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Integer) As Integer
    Public Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Integer) As Integer

#End If

    Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringW" (<MarshalAs(UnmanagedType.LPWStr)> ByVal lpstrCommand As String, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpstrTempurnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As IntPtr) As Integer
    Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringW" (ByVal dwError As Integer, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpstrBuffer As String, ByVal uLength As Integer) As Integer

    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" (<MarshalAs(UnmanagedType.LPWStr)> ByVal lpApplicationName As String, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpKeyName As String, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpDefault As String, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpReturnedString As StringBuilder, ByVal nSize As UInt32, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpFileName As String) As UInt32
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringW" (<MarshalAs(UnmanagedType.LPWStr)> ByVal lpApplicationName As String, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpKeyName As String, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpString As String, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpFileName As String) As Integer

    Public Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As IntPtr, <Out()> ByRef lpwndpl As WINDOWPLACEMENT) As Integer
    Public Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As IntPtr, <[In]()> ByRef lpwndpl As WINDOWPLACEMENT) As Integer

    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hwnd As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpOperation As String, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpFile As String, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpParameters As String, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpDirectory As String, ByVal nShowCmd As Integer) As IntPtr

    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hwnd As IntPtr, ByVal nIndex As Integer) As Integer
    Public Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrW" (ByVal hwnd As IntPtr, ByVal nIndex As Integer) As IntPtr

    Public Declare Function AdjustWindowRectEx Lib "user32" (<[In]()> ByRef lpRect As RECT, ByVal dsStyle As Integer, ByVal bMenu As Integer, ByVal dwEsStyle As Integer) As Integer

    Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Integer) As IntPtr
    Private Declare Function GetObject_Renamed Lib "gdi32" Alias "GetObjectW" (ByVal hObject As IntPtr, ByVal nCount As Integer, <Out()> ByRef lpObject As LOGFONT) As Integer

    'Get/SetWindowPlacement・ShellExecute 関連の定数
    Public Const SW_HIDE As Integer = 0
    Public Const SW_MAXIMIZE As Integer = 3
    Public Const SW_MINIMIZE As Integer = 6
    Public Const SW_RESTORE As Integer = 9
    Public Const SW_SHOW As Integer = 5
    Public Const SW_SHOWDEFAULT As Integer = 10
    Public Const SW_SHOWMAXIMIZED As Integer = 3
    Public Const SW_SHOWMINIMIZED As Integer = 2
    Public Const SW_SHOWMINNOACTIVE As Integer = 7
    Public Const SW_SHOWNA As Integer = 8
    Public Const SW_SHOWNOACTIVATE As Integer = 4
    Public Const SW_SHOWNORMAL As Integer = 1

    'GetWindowLong 関連の定数
    Public Const GWL_STYLE As Integer = -16
    Public Const GWL_EXSTYLE As Integer = -20

    'GetStockObject 関連の定数
    Private Const OEM_FIXED_FONT As Integer = 10
    Private Const ANSI_FIXED_FONT As Integer = 11
    Private Const ANSI_VAR_FONT As Integer = 12
    Private Const SYSTEM_FONT As Integer = 13
    Private Const SYSTEM_FIXED_FONT As Integer = 16
    Private Const DEFAULT_GUI_FONT As Integer = 17

    'LOGFONT 関連の定数
    Private Const DEFAULT_CHARSET As Byte = 1
    Private Const LF_FACESIZE As Integer = 32

    'SystemParametersInfo 関連の定数
    Private Const SPI_GETICONTITLELOGFONT As Integer = 31
    Private Const SPI_GETNONCLIENTMETRICS As Integer = 41

    Public Structure ItemWithData
        Public ItemString As String
        Public ItemData As Integer

        'ComboBoxには、ToString()メソッドの内容が表示される
        Public Overrides Function ToString() As String
            Return ItemString
        End Function

        '登録を簡単にするため、引数付きのコンストラクタを定義
        Public Sub New(ByVal S As String, ByVal I As Integer)
            ItemString = S
            ItemData = I
        End Sub

        Public Sub SetItemString(ByVal S As String)
            ItemString = S
        End Sub

        Public Sub SetItemData(ByVal I As Integer)
            ItemData = I
        End Sub
    End Structure

    Public Sub SetItemString(obj As ComboBox, index As Integer, itemstring As String)
        obj.Items.Insert(index, itemstring)
        If obj.Items.Count > index + 1 Then
            If obj.SelectedIndex = index + 1 Then
                obj.SelectedIndex = index
            End If
            obj.Items.RemoveAt(index + 1)
        End If
    End Sub

    Public Sub SetItemString(obj As ListBox, index As Integer, itemstring As String)
        obj.Items.Insert(index, itemstring)
        If obj.Items.Count > index + 1 Then
            If obj.SelectedIndex = index + 1 Then
                obj.SelectedIndex = index
            End If
            obj.Items.RemoveAt(index + 1)
        End If
    End Sub

    Public Function GetItemString(obj As ComboBox, index As Integer) As String
        GetItemString = obj.Items.Item(index).ToString()
    End Function

    Public Function GetItemString(obj As ListBox, index As Integer) As String
        GetItemString = obj.Items.Item(index).ToString()
    End Function

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto, Pack:=1)> Public Structure LOGFONT
        Public lfHeight As Int32
        Public lfWidth As Int32
        Public lfEscapement As Int32
        Public lfOrientation As Int32
        Public lfWeight As Int32
        Public lfItalic As Byte
        Public lfUnderline As Byte
        Public lfStrikeOut As Byte
        Public lfCharSet As Byte
        Public lfOutPrecision As Byte
        Public lfClipPrecision As Byte
        Public lfQuality As Byte
        Public lfPitchAndFamily As Byte
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=LF_FACESIZE)> Public lfFaceName As String
    End Structure

    Private Structure NONCLIENTMETRICS
        Dim cbSize As Integer
        Dim iBorderWidth As Integer
        Dim iScrollWidth As Integer
        Dim iScrollHeight As Integer
        Dim iCaptionWidth As Integer
        Dim iCaptionHeight As Integer
        Dim lfCaptionFont As LOGFONT
        Dim iSMCaptionWidth As Integer
        Dim iSMCaptionHeight As Integer
        Dim lfSMCaptionFont As LOGFONT
        Dim iMenuWidth As Integer
        Dim iMenuHeight As Integer
        Dim lfMenuFont As LOGFONT
        Dim lfStatusFont As LOGFONT
        Dim lfMessageFont As LOGFONT
    End Structure

    Public Structure POINTAPI
        Dim X As Integer
        Dim Y As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)> Public Structure RECT
        Dim left_Renamed As Integer
        Dim Top As Integer
        Dim right_Renamed As Integer
        Dim Bottom As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)> Public Structure WINDOWPLACEMENT
        Dim Length As Integer
        Dim flags As Integer
        Dim showCmd As Integer
        Dim ptMinPosition As POINTAPI
        Dim ptMaxPosition As POINTAPI
        Dim rcNormalPosition As RECT
    End Structure

    Public Const PI As Single = 3.14159265358979
    Public Const RAD As Single = PI / 180

    Public Enum OBJ_SELECT
        NON_SELECT '未選択
        Selected '選択
        EDIT_RECT '白枠 (編集モード)
        DELETE_RECT '赤枠 (消去モード)
        SELECTAREA_IN '選択範囲内にあるオブジェ、選択中
        SELECTAREA_OUT '選択範囲を展開した時に既に選択状態にあったオブジェ、選択中
        SELECTAREA_SELECTED '↑かつ選択範囲内、つまり選択状態でなくなったオブジェ
    End Enum

    Public Enum OBJ_ATT
        OBJ_NORMAL
        OBJ_INVISIBLE
        OBJ_LONGNOTE
        OBJ_MINE
    End Enum

    Public Enum BGA_PARA
        BGA_NUM
        BGA_X1
        BGA_Y1
        BGA_X2
        BGA_Y2
        BGA_dX
        BGA_dY
    End Enum

    Public Enum CMD_LOG
        NONE
        OBJ_ADD
        OBJ_DEL
        OBJ_MOVE
        OBJ_CHANGE
        MSR_ADD
        MSR_DEL
        MSR_CHANGE
        WAV_CHANGE
        BMP_CHANGE
        LIST_ALIGN
        LIST_DEL
    End Enum

    Public g_strAppTitle As String

    Public Structure m_udtMouse
        Dim X As Integer
        Dim Y As Integer
        Dim Shift As Keys
        Dim Button As MouseButtons
        Dim measure As Integer
    End Structure

    Public g_Mouse As m_udtMouse

    Public Structure m_udtDisplay
        Dim X As Integer
        Dim Y As Integer
        Dim Width As Single
        Dim Height As Single
        Dim lngMaxX As Integer
        Dim lngMaxY As Integer
        Dim intStartMeasure As Integer
        Dim intEndMeasure As Integer
        Dim lngStartPos As Integer
        Dim lngEndPos As Integer
        Dim intMaxMeasure As Integer '最大表示小節
        Dim intResolution As Integer '分解能
        Dim intEffect As Integer '画面効果
    End Structure

    Public g_disp As m_udtDisplay

    Public Structure m_udtBMS
        Dim strDir As String 'ディレクトリ
        Dim strFileName As String 'BMSファイル名
        Dim intPlayerType As Integer '#PLAYER
        Dim strGenre As String '#GENRE
        Dim strTitle As String '#TITLE
        Dim strArtist As String '#ARTIST
        Dim sngBPM As Single '#BPM
        Dim lngPlayLevel As Integer '#PLAYLEVEL
        Dim intPlayRank As Integer '#RANK
        Dim sngTotal As Single '#TOTAL
        Dim intVolume As Integer '#VOLWAV
        Dim strStageFile As String '#STAGEFILE
        Dim strSubTitle As String '#SUBTITLE
        Dim strSubArtist As String '#SUBARTIST
        Dim intDifficulty As Integer '#DIFFICULTY
        Dim strPreview As String '#PREVIEW
        Dim strBanner As String '#BANNER
        Dim intLNObj As Integer '#LNOBJ
        Dim intLNMode As Integer '#LNMODE
        Dim intDefExRank As Integer '#DEFEXRANK
        Dim strBackBMP As String '#BACKBMP
        Dim strComment As String '#COMMENT
        Dim blnSaveFlag As Boolean
    End Structure

    Public g_BMS As m_udtBMS

    Public Structure m_udtVerticalLine
        Dim blnVisible As Boolean
        Dim intCh As Integer
        Dim strText As String
        Dim intWidth As Integer
        Dim lngLeft As Integer
        Dim lngObjLeft As Integer
        Dim lngBackColor As Integer
        Dim intLightNum As Integer
        Dim intShadowNum As Integer
        Dim intBrushNum As Integer
        Dim blnDraw As Boolean
    End Structure

    Public g_VGrid(31 + modInput.BGM_LANE) As m_udtVerticalLine

    Public g_intVGridNum(36 ^ 2 + 4 * 36 + modInput.BGM_LANE) As Integer

    Public Structure g_udtObj
        Dim lngID As Integer
        Dim intCh As Integer
        Dim intAtt As OBJ_ATT
        Dim intMeasure As Integer
        Dim lngHeight As Integer
        Dim lngPosition As Integer
        Dim lngTail As Long
        Dim sngValue As Single
        Dim intSelect As OBJ_SELECT
        '0・・・未選択
        '1・・・選択
        '2・・・白枠 (編集モード)
        '3・・・赤枠 (消去モード)
        '4・・・選択範囲内にあるオブジェ、選択中
        '5・・・選択範囲を展開した時に既に選択状態にあったオブジェ、選択中
        '6・・・5番かつ選択範囲内、つまり選択状態でなくなったオブジェ
    End Structure

    Public g_Obj() As g_udtObj

    Public g_lngObjID() As Integer
    Public g_lngIDNum As Integer

    Public Structure m_udtSelectArea
        Dim blnFlag As Boolean
        Dim X1 As Single
        Dim Y1 As Single
        Dim X2 As Single
        Dim Y2 As Single
    End Structure

    Public g_SelectArea As m_udtSelectArea

    Public g_strLangFileName() As String
    Public g_strThemeFileName() As String
    Public g_strStatusBar(25) As String

    Public g_blnIgnoreInput As Boolean
    Public g_strAppDir As String
    Public g_strHelpFilename As String
    Public g_strFiler As String
    Public g_strRecentFiles(4) As String

    Public g_InputLog As New clsLog

    Public Structure g_udtViewer
        Dim strAppName As String
        Dim strAppPath As String
        Dim strArgAll As String
        Dim strArgPlay As String
        Dim strArgStop As String
    End Structure

    Public g_Viewer() As g_udtViewer

    Public Enum Message
        ERR_01
        ERR_02
        ERR_FILE_NOT_FOUND
        ERR_LOAD_CANCEL
        ERR_SAVE_ERROR
        ERR_SAVE_CANCEL
        ERR_OVERFLOW_LARGE
        ERR_OVERFLOW_SMALL
        ERR_OVERFLOW_BPM
        ERR_OVERFLOW_STOP
        ERR_OVERFLOW_SCROLL
        ERR_OVERFLOW_SPEED
        ERR_APP_NOT_FOUND
        ERR_FILE_ALREADY_EXIST
        MSG_CONFIRM
        MSG_FILE_CHANGED
        MSG_INI_CHANGED
        MSG_ALIGN_LIST
        MSG_DELETE_FILE
        INPUT_SCROLL
        INPUT_BPM
        INPUT_STOP
        INPUT_SPEED
        INPUT_RENAME
        INPUT_SIZE
        Max
    End Enum

    Public g_Message(Message.Max - 1) As String

    Public Function LenB(ByVal stTarget As String) As Integer
        If stTarget <> Nothing Then
            Return System.Text.Encoding.Unicode.GetByteCount(stTarget)
        Else
            Return 0
        End If
    End Function

    Public Sub StartUp()
        Dim i As Integer
        Dim strTemp As String
        Dim intTemp As Integer
        Dim lngFFile As Integer

        If Right(My.Application.Info.DirectoryPath, 1) = "\" Then

            g_strAppDir = My.Application.Info.DirectoryPath

        Else

            g_strAppDir = My.Application.Info.DirectoryPath & "\"

        End If

        g_strAppTitle = "BMx Sequence Editor " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build
        g_strAppTitle = g_strAppTitle & ""

#If MODE_DEBUG = False Then
		
		Call modMessage.SubClass(frmMain.hwnd)

#End If

#If MODE_SPEEDTEST Then

        Call timeBeginPeriod(1)

#End If

        ReDim g_strLangFileName(0)

        Call g_InputLog.Clear()

        ReDim g_Viewer(1)

        'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
        If Dir(g_strAppDir & "bmse_viewer.ini", FileAttribute.Normal) = vbNullString Then

            lngFFile = FreeFile()

            FileOpen(lngFFile, g_strAppDir & "bmse_viewer.ini", OpenMode.Output)

            PrintLine(lngFFile, "mBMplay")
            PrintLine(lngFFile, "..\mBMplay\mBMplay.exe")
            PrintLine(lngFFile, "<filename>")
            PrintLine(lngFFile, "-s <measure> <filename>")
            PrintLine(lngFFile, "-t")
            PrintLine(lngFFile)
            PrintLine(lngFFile, "uBMplay")
            PrintLine(lngFFile, "..\uBMplay\uBMplay.exe")
            PrintLine(lngFFile, "-P -N0 <filename>")
            PrintLine(lngFFile, "-P -N<measure> <filename>")
            PrintLine(lngFFile, "-S")
            PrintLine(lngFFile)
            PrintLine(lngFFile, "beatoraja")
            PrintLine(lngFFile, "..\beatoraja-jre\beatoraja.exe")
            PrintLine(lngFFile, "-a <filename>")
            PrintLine(lngFFile, "-p <filename>")
            PrintLine(lngFFile, "")
            PrintLine(lngFFile)
            PrintLine(lngFFile, "LR2")
            PrintLine(lngFFile, "..\LR2\LR2body.exe")
            PrintLine(lngFFile, "-a <filename>")
            PrintLine(lngFFile, "-a <filename>")
            PrintLine(lngFFile, "")

            FileClose(lngFFile)

        End If

        i = 0
        lngFFile = FreeFile()

        FileOpen(lngFFile, g_strAppDir & "bmse_viewer.ini", OpenMode.Input)

        Do While Not EOF(lngFFile)

            strTemp = LineInput(lngFFile)

            Select Case i Mod 6

                Case 0

                    If Len(strTemp) = 0 Then Exit Do
                    g_Viewer(UBound(g_Viewer)).strAppName = strTemp

                Case 1

                    If Len(strTemp) = 0 Then Exit Do
                    g_Viewer(UBound(g_Viewer)).strAppPath = strTemp

                Case 2

                    g_Viewer(UBound(g_Viewer)).strArgAll = strTemp

                Case 3

                    g_Viewer(UBound(g_Viewer)).strArgPlay = strTemp

                Case 4

                    g_Viewer(UBound(g_Viewer)).strArgStop = strTemp

                    Call frmMain.cboViewer.Items.Add(g_Viewer(UBound(g_Viewer)).strAppName)
                    ReDim Preserve g_Viewer(UBound(g_Viewer) + 1)

            End Select

            i = i + 1

        Loop

        FileClose(lngFFile)

        ReDim Preserve g_Viewer(frmMain.cboViewer.Items.Count)

        'ランゲージファイル読み込み
        'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
        ReDim Preserve g_strLangFileName(2)

        strTemp = Dir(g_strAppDir & "lang\*.ini")
        intTemp = 0

        Do While strTemp <> ""
            If strGet_ini("Main", "Key", "", "lang\" & strTemp) = "BMSE" Then
                g_strLangFileName(intTemp) = strTemp

                Select Case intTemp
                    Case 0
                        With frmMain._mnuLanguage_0
                            .Text = "&" & strGet_ini("Main", "Language", strTemp, "lang\" & strTemp)
                            If .Text = "&" Then .Text = "&" & strTemp
                            .Visible = True
                        End With
                        intTemp = intTemp + 1

                    Case 1
                        With frmMain._mnuLanguage_1
                            .Text = "&" & strGet_ini("Main", "Language", strTemp, "lang\" & strTemp)
                            If .Text = "&" Then .Text = "&" & strTemp
                            .Visible = True
                        End With
                        intTemp = intTemp + 1

                    Case 2
                        With frmMain._mnuLanguage_2
                            .Text = "&" & strGet_ini("Main", "Language", strTemp, "lang\" & strTemp)
                            If .Text = "&" Then .Text = "&" & strTemp
                            .Visible = True
                        End With
                        intTemp = intTemp + 1
                End Select
            End If

            strTemp = Dir()
        Loop

        If intTemp Then
        Else
            frmMain.mnuLanguageParent.Enabled = False
        End If

        'テーマファイル読み込み
        'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
        ReDim Preserve g_strThemeFileName(2)

        strTemp = Dir(g_strAppDir & "theme\*.ini")
        intTemp = 0

        Do While strTemp <> ""
            If strGet_ini("Main", "Key", "", "theme\" & strTemp) = "BMSE" Then


                g_strThemeFileName(intTemp) = strTemp

                Select Case intTemp
                    Case 0
                        With frmMain._mnuTheme_0
                            .Text = "&" & strGet_ini("Main", "Name", strTemp, "theme\" & strTemp)
                            If .Text = "&" Then .Text = "&" & strTemp
                            .Visible = True
                        End With
                        intTemp = intTemp + 1

                    Case 1
                        With frmMain._mnuTheme_1
                            .Text = "&" & strGet_ini("Main", "Name", strTemp, "theme\" & strTemp)
                            If .Text = "&" Then .Text = "&" & strTemp
                            .Visible = True
                        End With
                        intTemp = intTemp + 1

                    Case 2
                        With frmMain._mnuTheme_2
                            .Text = "&" & strGet_ini("Main", "Name", strTemp, "theme\" & strTemp)
                            If .Text = "&" Then .Text = "&" & strTemp
                            .Visible = True
                        End With
                        intTemp = intTemp + 1
                End Select
            End If

            strTemp = Dir()
        Loop

        If intTemp Then
        Else
            frmMain.mnuThemeParent.Enabled = False
        End If

        '初期化
        With g_BMS

            .intPlayerType = modInput.PLAYER_TYPE.PLAYER_1P
            .strGenre = ""
            .strTitle = ""
            .strArtist = ""
            .sngBPM = CSng(Val(frmMain.txtBPM.Text))
            .lngPlayLevel = 1
            .intPlayRank = modInput.PLAY_RANK.RANK_EASY
            .sngTotal = 0
            .intVolume = 0
            .blnSaveFlag = True

        End With

        ReDim g_Obj(0)
        ReDim g_lngObjID(0)
        g_lngIDNum = 0

        For i = 0 To 256 + 64

            g_sngSin(i) = System.Math.Sin(i * PI / 128)

        Next i

        For i = 0 To UBound(g_VGrid)

            With g_VGrid(i)

                .intCh = Choose(i + 1, 0, 1033, 1020, 8, 9, 0,
                                2 * 36 + 1, 1 * 36 + 6, 1 * 36 + 1, 1 * 36 + 2, 1 * 36 + 3, 1 * 36 + 4, 1 * 36 + 5, 1 * 36 + 8, 1 * 36 + 9, 1 * 36 + 6, 0,
                                2 * 36 + 6, 2 * 36 + 1, 2 * 36 + 2, 2 * 36 + 3, 2 * 36 + 4, 2 * 36 + 5, 2 * 36 + 8, 2 * 36 + 9, 2 * 36 + 6, 0,
                                4, 7, 6, 0,
                                1 * 36 ^ 2 + 1, 1 * 36 ^ 2 + 2, 1 * 36 ^ 2 + 3, 1 * 36 ^ 2 + 4, 1 * 36 ^ 2 + 5, 1 * 36 ^ 2 + 6, 1 * 36 ^ 2 + 7, 1 * 36 ^ 2 + 8, 1 * 36 ^ 2 + 9, 1 * 36 ^ 2 + 10,
                                1 * 36 ^ 2 + 11, 1 * 36 ^ 2 + 12, 1 * 36 ^ 2 + 13, 1 * 36 ^ 2 + 14, 1 * 36 ^ 2 + 15, 1 * 36 ^ 2 + 16, 1 * 36 ^ 2 + 17, 1 * 36 ^ 2 + 18, 1 * 36 ^ 2 + 19, 1 * 36 ^ 2 + 20,
                                1 * 36 ^ 2 + 21, 1 * 36 ^ 2 + 22, 1 * 36 ^ 2 + 23, 1 * 36 ^ 2 + 24, 1 * 36 ^ 2 + 25, 1 * 36 ^ 2 + 26, 1 * 36 ^ 2 + 27, 1 * 36 ^ 2 + 28, 1 * 36 ^ 2 + 29, 1 * 36 ^ 2 + 30,
                                1 * 36 ^ 2 + 31, 1 * 36 ^ 2 + 32, 1 * 36 ^ 2 + 33, 1 * 36 ^ 2 + 34, 1 * 36 ^ 2 + 35, 1 * 36 ^ 2 + 36, 1 * 36 ^ 2 + 37, 1 * 36 ^ 2 + 38, 1 * 36 ^ 2 + 39, 1 * 36 ^ 2 + 40,
                                1 * 36 ^ 2 + 41, 1 * 36 ^ 2 + 42, 1 * 36 ^ 2 + 43, 1 * 36 ^ 2 + 44, 1 * 36 ^ 2 + 45, 1 * 36 ^ 2 + 46, 1 * 36 ^ 2 + 47, 1 * 36 ^ 2 + 48, 1 * 36 ^ 2 + 49, 1 * 36 ^ 2 + 50,
                                1 * 36 ^ 2 + 51, 1 * 36 ^ 2 + 52, 1 * 36 ^ 2 + 53, 1 * 36 ^ 2 + 54, 1 * 36 ^ 2 + 55, 1 * 36 ^ 2 + 56, 1 * 36 ^ 2 + 57, 1 * 36 ^ 2 + 58, 1 * 36 ^ 2 + 59, 1 * 36 ^ 2 + 60,
                                1 * 36 ^ 2 + 61, 1 * 36 ^ 2 + 62, 1 * 36 ^ 2 + 63, 1 * 36 ^ 2 + 64, 1 * 36 ^ 2 + 65, 1 * 36 ^ 2 + 66, 1 * 36 ^ 2 + 67, 1 * 36 ^ 2 + 68, 1 * 36 ^ 2 + 69, 1 * 36 ^ 2 + 70,
                                1 * 36 ^ 2 + 71, 1 * 36 ^ 2 + 72, 1 * 36 ^ 2 + 73, 1 * 36 ^ 2 + 74, 1 * 36 ^ 2 + 75, 1 * 36 ^ 2 + 76, 1 * 36 ^ 2 + 77, 1 * 36 ^ 2 + 78, 1 * 36 ^ 2 + 79, 1 * 36 ^ 2 + 80,
                                1 * 36 ^ 2 + 81, 1 * 36 ^ 2 + 82, 1 * 36 ^ 2 + 83, 1 * 36 ^ 2 + 84, 1 * 36 ^ 2 + 85, 1 * 36 ^ 2 + 86, 1 * 36 ^ 2 + 87, 1 * 36 ^ 2 + 88, 1 * 36 ^ 2 + 89, 1 * 36 ^ 2 + 90,
                                1 * 36 ^ 2 + 91, 1 * 36 ^ 2 + 92, 1 * 36 ^ 2 + 93, 1 * 36 ^ 2 + 94, 1 * 36 ^ 2 + 95, 1 * 36 ^ 2 + 96, 1 * 36 ^ 2 + 97, 1 * 36 ^ 2 + 98, 1 * 36 ^ 2 + 99, 1 * 36 ^ 2 + 100,
                                1 * 36 ^ 2 + 101, 1 * 36 ^ 2 + 102, 1 * 36 ^ 2 + 103, 1 * 36 ^ 2 + 104, 1 * 36 ^ 2 + 105, 1 * 36 ^ 2 + 106, 1 * 36 ^ 2 + 107, 1 * 36 ^ 2 + 108, 1 * 36 ^ 2 + 109, 1 * 36 ^ 2 + 110,
                                1 * 36 ^ 2 + 111, 1 * 36 ^ 2 + 112, 1 * 36 ^ 2 + 113, 1 * 36 ^ 2 + 114, 1 * 36 ^ 2 + 115, 1 * 36 ^ 2 + 116, 1 * 36 ^ 2 + 117, 1 * 36 ^ 2 + 118, 1 * 36 ^ 2 + 119, 1 * 36 ^ 2 + 120,
                                1 * 36 ^ 2 + 121, 1 * 36 ^ 2 + 122, 1 * 36 ^ 2 + 123, 1 * 36 ^ 2 + 124, 1 * 36 ^ 2 + 125, 1 * 36 ^ 2 + 126, 1 * 36 ^ 2 + 127, 1 * 36 ^ 2 + 128, 0) 'これ直書きするの？マジで！？ 1
                'If .intCh Then g_intVGridNum(.intCh) = i
                .blnVisible = True

                Select Case .intCh

                    Case modInput.OBJ_CH.CH_BPM, modInput.OBJ_CH.CH_EXBPM, modInput.OBJ_CH.CH_STOP, modInput.OBJ_CH.CH_SCROLL, modInput.OBJ_CH.CH_SPEED 'BPM/STOP/SCROLL/SPEED

                        .intLightNum = modDraw.PEN_NUM.BPM_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.BPM_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.BPM

                    Case modInput.OBJ_CH.CH_BGA, modInput.OBJ_CH.CH_LAYER, modInput.OBJ_CH.CH_POOR 'BGA/Layer/Poor

                        .intLightNum = modDraw.PEN_NUM.BGA_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.BGA_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.BGA

                    Case 1 * 36 + 1

                        .intLightNum = modDraw.PEN_NUM.KEY01_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.KEY01_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.KEY01

                    Case 1 * 36 + 2

                        .intLightNum = modDraw.PEN_NUM.KEY02_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.KEY02_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.KEY02

                    Case 1 * 36 + 3

                        .intLightNum = modDraw.PEN_NUM.KEY03_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.KEY03_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.KEY03

                    Case 1 * 36 + 4

                        .intLightNum = modDraw.PEN_NUM.KEY04_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.KEY04_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.KEY04

                    Case 1 * 36 + 5

                        .intLightNum = modDraw.PEN_NUM.KEY05_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.KEY05_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.KEY05

                    Case 1 * 36 + 8

                        .intLightNum = modDraw.PEN_NUM.KEY06_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.KEY06_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.KEY06

                    Case 1 * 36 + 9

                        .intLightNum = modDraw.PEN_NUM.KEY07_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.KEY07_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.KEY07

                    Case 1 * 36 + 6

                        .intLightNum = modDraw.PEN_NUM.KEY08_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.KEY08_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.KEY08

                    Case 2 * 36 + 1

                        .intLightNum = modDraw.PEN_NUM.KEY11_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.KEY11_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.KEY11

                    Case 2 * 36 + 2

                        .intLightNum = modDraw.PEN_NUM.KEY12_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.KEY12_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.KEY12

                    Case 2 * 36 + 3

                        .intLightNum = modDraw.PEN_NUM.KEY13_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.KEY13_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.KEY13

                    Case 2 * 36 + 4

                        .intLightNum = modDraw.PEN_NUM.KEY14_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.KEY14_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.KEY14

                    Case 2 * 36 + 5

                        .intLightNum = modDraw.PEN_NUM.KEY15_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.KEY15_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.KEY15

                    Case 2 * 36 + 8

                        .intLightNum = modDraw.PEN_NUM.KEY16_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.KEY16_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.KEY16

                    Case 2 * 36 + 9

                        .intLightNum = modDraw.PEN_NUM.KEY17_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.KEY17_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.KEY17

                    Case 2 * 36 + 6

                        .intLightNum = modDraw.PEN_NUM.KEY18_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.KEY18_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.KEY18

                    Case Is > 36 ^ 2 'BGM

                        .intLightNum = modDraw.PEN_NUM.BGM_LIGHT
                        .intShadowNum = modDraw.PEN_NUM.BGM_SHADOW
                        .intBrushNum = modDraw.BRUSH_NUM.BGM

                End Select

                If .intCh Then

                    .intWidth = GRID_WIDTH

                Else

                    .intWidth = SPACE_WIDTH

                End If

            End With

        Next i

        'g_Disp.intMaxMeasure = 31
        g_disp.intMaxMeasure = 0
        Call modDraw.lngChangeMaxMeasure(15)
        Call modDraw.ChangeResolution()

    End Sub

    Public Sub CleanUp(Optional ByVal lngErrNum As Integer = 0, Optional ByRef strErrDescription As String = "", Optional ByRef strErrProcedure As String = "")
        On Error Resume Next

        Dim i As Integer

#If MODE_DEBUG = False Then
		
		Call modMessage.UnSubClass(frmMain.hwnd)
		
#End If

#If MODE_SPEEDTEST Then

        Call timeEndPeriod(1)

#End If

        g_InputLog = Nothing

        Call modEasterEgg.EndEffect()

        Call SaveConfig()

        Call mciSendString("close PREVIEW", vbNullString, 0, 0)

        Call lngDeleteFile(g_BMS.strDir & "___bmse_temp.bms")
        Call lngDeleteFile(g_strAppDir & "___bmse_temp.bms")

        If lngErrNum <> 0 And strErrDescription <> "" And strErrProcedure <> "" Then

            If Len(g_BMS.strDir) = 0 Then g_BMS.strDir = g_strAppDir

            For i = 0 To 9999

                g_BMS.strFileName = "temp" & Format(i, "0000") & ".bms"

                If i = 9999 Then

                    Call CreateBMS(g_BMS.strDir & g_BMS.strFileName)

                    'UPGRADE_WARNING: Dir に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
                ElseIf Dir(g_BMS.strDir & g_BMS.strFileName) = vbNullString Then

                    Call CreateBMS(g_BMS.strDir & g_BMS.strFileName)
                    Exit For

                End If

            Next i

            Call DebugOutput(lngErrNum, strErrDescription, strErrProcedure, True)

        End If

        End

    End Sub

    Public Sub DebugOutput(ByVal lngErrNum As Integer, ByRef strErrDescription As String, ByRef strErrProcedure As String, Optional ByVal blnCleanUp As Boolean = False)
        Dim lngFFile As Integer
        Dim strError As String = ""

        lngFFile = FreeFile()

        FileOpen(lngFFile, g_strAppDir & "error.txt", OpenMode.Append)

        PrintLine(lngFFile, Today & TimeOfDay & "ErrorNo." & lngErrNum & " " & strErrDescription & "@" & strErrProcedure & "/BMSE_" & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build)

        FileClose(lngFFile)

        strError = strError & "ErrorNo." & lngErrNum & " " & strErrDescription & "@" & strErrProcedure

        If blnCleanUp Then

            strError = g_Message(modMain.Message.ERR_01) & vbCrLf & strError & vbCrLf
            strError = strError & g_Message(modMain.Message.ERR_02) & vbCrLf
            strError = strError & g_BMS.strDir & g_BMS.strFileName

        End If

        Call frmMain.Show()
        Call MsgBox(strError, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, g_strAppTitle)

    End Sub

    Public Function lngDeleteFile(ByRef FileName As String) As Integer
        On Error GoTo Err_Renamed

        Kill(FileName)

        Exit Function

Err_Renamed:
        lngDeleteFile = 1
    End Function

    Public Function intSaveCheck() As Integer
        On Error GoTo Err_Renamed

        Dim lngTemp As Integer
        Dim strArray() As String

        With frmMain

            If .cboPlayer.SelectedIndex + 1 <> g_BMS.intPlayerType Then g_BMS.blnSaveFlag = False
            If .txtGenre.Text <> g_BMS.strGenre Then g_BMS.blnSaveFlag = False
            If .txtTitle.Text <> g_BMS.strTitle Then g_BMS.blnSaveFlag = False
            If .txtArtist.Text <> g_BMS.strArtist Then g_BMS.blnSaveFlag = False
            If CLng(Val(.cboPlayLevel.Text)) <> g_BMS.lngPlayLevel Then g_BMS.blnSaveFlag = False
            If CSng(Val(.txtBPM.Text)) <> g_BMS.sngBPM Then g_BMS.blnSaveFlag = False

            If .cboPlayRank.SelectedIndex <> g_BMS.intPlayRank Then g_BMS.blnSaveFlag = False
            If CSng(Val(.txtTotal.Text)) <> g_BMS.sngTotal Then g_BMS.blnSaveFlag = False
            If CInt(Val(.txtVolume.Text)) <> g_BMS.intVolume Then g_BMS.blnSaveFlag = False
            If .txtStageFile.Text <> g_BMS.strStageFile Then g_BMS.blnSaveFlag = False
            'If .txtMissBMP.Text <> g_strBMP(0) Then g_BMS.blnSaveFlag = False

        End With

        If g_BMS.blnSaveFlag Then

            intSaveCheck = 0

            Exit Function

        End If

        Call frmMain.Show()

        lngTemp = MsgBox(g_Message(modMain.Message.MSG_FILE_CHANGED), MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNoCancel, g_strAppTitle)

        Select Case lngTemp

            Case MsgBoxResult.Yes

                If g_BMS.strDir <> "" And g_BMS.strFileName <> "" Then

                    Call CreateBMS(g_BMS.strDir & g_BMS.strFileName)

                Else

                    With frmMain.dlgMainSave

                        'UPGRADE_WARNING: Filter に新しい動作が指定されています。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' をクリックしてください。
                        .Filter = "BMS files (*.bms,*.bme,*.bml,*.pms)|*.bms;*.bme;*.bml;*.pms|All files (*.*)|*.*"

                        .FileName = g_BMS.strFileName

                        If .ShowDialog() <> DialogResult.OK Then
                            intSaveCheck = 1
                            Exit Function
                        End If

                        strArray = Split(.FileName, "\")
                        g_BMS.strDir = Left(.FileName, Len(.FileName) - Len(strArray(UBound(strArray))))
                        g_BMS.strFileName = strArray(UBound(strArray))

                        Call CreateBMS(g_BMS.strDir & g_BMS.strFileName)

                        Call RecentFilesRotation(g_BMS.strDir & g_BMS.strFileName)

                        frmMain.dlgMainOpen.InitialDirectory = g_BMS.strDir
                        frmMain.dlgMainSave.InitialDirectory = g_BMS.strDir

                    End With

                End If

            Case MsgBoxResult.No

                intSaveCheck = 0

            Case MsgBoxResult.Cancel

                intSaveCheck = 1

        End Select

        Exit Function

Err_Renamed:

        intSaveCheck = 1

    End Function

    Public Sub RecentFilesRotation(ByRef strFilePath As String)
        Dim i As Integer
        Dim intTemp As Integer

        For i = 0 To UBound(g_strRecentFiles)

            If g_strRecentFiles(i) = strFilePath Then

                Call SubRotate(0, i, strFilePath)

                intTemp = 1

                Exit For

            End If

        Next i

        If intTemp = 0 Then Call SubRotate(0, UBound(g_strRecentFiles), strFilePath)

        frmMain.mnuLineRecent.Visible = True

    End Sub

    Private Sub SubRotate(ByVal intIndex As Integer, ByVal intEnd As Integer, ByRef strFilePath As String)
        If intIndex <> intEnd And g_strRecentFiles(intIndex) <> "" And intIndex <= UBound(g_strRecentFiles) Then

            Call SubRotate(intIndex + 1, intEnd, g_strRecentFiles(intIndex))

        End If

        g_strRecentFiles(intIndex) = strFilePath

        Select Case intIndex
            Case 0
                With frmMain._mnuRecentFiles_0
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

                With frmMain.ToolStripMenuItem0
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

            Case 1
                With frmMain._mnuRecentFiles_1
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

                With frmMain.ToolStripMenuItem1
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

            Case 2
                With frmMain._mnuRecentFiles_2
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

                With frmMain.ToolStripMenuItem2
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

            Case 3
                With frmMain._mnuRecentFiles_3
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

                With frmMain.ToolStripMenuItem3
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

            Case 4
                With frmMain._mnuRecentFiles_4
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

                With frmMain.ToolStripMenuItem4
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

            Case 5
                With frmMain._mnuRecentFiles_5
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

                With frmMain.ToolStripMenuItem5
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

            Case 6
                With frmMain._mnuRecentFiles_6
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

                With frmMain.ToolStripMenuItem6
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

            Case 7
                With frmMain._mnuRecentFiles_7
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

                With frmMain.ToolStripMenuItem7
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

            Case 8
                With frmMain._mnuRecentFiles_8
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

                With frmMain.ToolStripMenuItem8
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

            Case 9
                With frmMain._mnuRecentFiles_9
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With

                With frmMain.ToolStripMenuItem9
                    .Text = "&" & intIndex + 1 & ":" & strFilePath
                    .Enabled = True
                    .Visible = True
                End With
        End Select
    End Sub

    Public Sub GetCmdLine()
        On Error GoTo Err_Renamed

        Dim i As Integer
        Dim strTemp As String
        Dim strCmdArray() As String
        Dim strArray() As String
        Dim ReadLock As Boolean
        Dim blnReadFlag As Boolean

        strTemp = Trim(VB.Command())

        If strTemp = "" Then Exit Sub

        ReDim strCmdArray(0)

        For i = 1 To Len(strTemp)

            Select Case Asc(Mid(strTemp, i, 1))

                Case 32 'スペース

                    If ReadLock = False Then

                        ReDim Preserve strCmdArray(UBound(strCmdArray) + 1)

                    Else

                        strCmdArray(UBound(strCmdArray)) = strCmdArray(UBound(strCmdArray)) & " "

                    End If

                Case 34 'ダブルクオーテーション

                    ReadLock = Not ReadLock

                Case Else

                    strCmdArray(UBound(strCmdArray)) = strCmdArray(UBound(strCmdArray)) & Mid(strTemp, i, 1)

            End Select

        Next i

        For i = 0 To UBound(strCmdArray)

            If strCmdArray(i) <> "" Then

                If InStr(1, strCmdArray(i), ":\") <> 0 And (UCase(Right(strCmdArray(i), 4)) = ".BMS" Or UCase(Right(strCmdArray(i), 4)) = ".BME" Or UCase(Right(strCmdArray(i), 4)) = ".BML" Or UCase(Right(strCmdArray(i), 4)) = ".PMS") Then

                    If blnReadFlag Then

                        'UPGRADE_WARNING: App プロパティ App.EXEName には新しい動作が含まれます。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' をクリックしてください。
                        Call ShellExecute(0, "open", Chr(34) & g_strAppDir & My.Application.Info.AssemblyName & Chr(34), Chr(34) & strCmdArray(i) & Chr(34), "", SW_SHOWNORMAL)

                    Else

                        strArray = Split(strCmdArray(i), "\")
                        g_BMS.strFileName = Right(strCmdArray(i), Len(strArray(UBound(strArray))))
                        g_BMS.strDir = Left(strCmdArray(i), Len(strCmdArray(i)) - Len(strArray(UBound(strArray))))
                        frmMain.dlgMainOpen.InitialDirectory = g_BMS.strDir
                        frmMain.dlgMainSave.InitialDirectory = g_BMS.strDir
                        blnReadFlag = True

                        Call modInput.LoadBMS()
                        Call RecentFilesRotation(g_BMS.strDir & g_BMS.strFileName)

                    End If

                End If

            End If

        Next i

        Exit Sub

Err_Renamed:
        Call modMain.CleanUp(Err.Number, Err.Description, "GetCmdLine")
    End Sub

    Public Sub LoadThemeFile(ByRef strFileName As String)
        Dim strArray() As String
        Dim i As Integer
        Dim j As Integer
        Dim Color As Integer
        Dim lngTemp As Integer

        frmMain.picMain.BackColor = System.Drawing.ColorTranslator.FromOle(GetColor("Main", "Background", "0,0,0", strFileName))

        g_lngSystemColor(modDraw.COLOR_NUM.MEASURE_NUM) = GetColor("Main", "MeasureNum", "64,64,64", strFileName)
        g_lngSystemColor(modDraw.COLOR_NUM.MEASURE_LINE) = GetColor("Main", "MeasureLine", "255,255,255", strFileName)
        g_lngSystemColor(modDraw.COLOR_NUM.GRID_MAIN) = GetColor("Main", "GridMain", "96,96,96", strFileName)
        g_lngSystemColor(modDraw.COLOR_NUM.GRID_SUB) = GetColor("Main", "GridSub", "192,192,192", strFileName)
        g_lngSystemColor(modDraw.COLOR_NUM.VERTICAL_MAIN) = GetColor("Main", "VerticalMain", "255,255,255", strFileName)
        g_lngSystemColor(modDraw.COLOR_NUM.VERTICAL_SUB) = GetColor("Main", "VerticalSub", "128,128,128", strFileName)
        g_lngSystemColor(modDraw.COLOR_NUM.INFO) = GetColor("Main", "Info", "0,255,0", strFileName)


        For i = 0 To modDraw.BRUSH_NUM.Max - 1

            Select Case i

                Case modDraw.BRUSH_NUM.BGM

                    Color = GetColor("BGM", "Background", "48,0,0", strFileName)

                    strArray = Split(strGet_ini("BGM", "Text", "B001,B002,B003,B004,B005,B006,B007,B008,B009,B010,B011,B012,B013,B014,B015,B016,B017,B018,B019,B020,B021,B022,B023,B024,B025,B026,B027,B028,B029,B030,B031,B032,B033,B034,B035,B036,B037,B038,B039,B040,B041,B042,B043,B044,B045,B046,B047,B048,B049,B050,B051,B052,B053,B054,B055,B056,B057,B058,B059,B060,B061,B062,B063,B064,B065,B066,B067,B068,B069,B070,B071,B072,B073,B074,B075,B076,B077,B078,B079,B080,B081,B082,B083,B084,B085,B086,B087,B088,B089,B090,B091,B092,B093,B094,B095,B096,B097,B098,B099,B100,B101,B102,B103,B104,B105,B106,B107,B108,B109,B110,B111,B112,B113,B114,B115,B116,B117,B118,B119,B120,B121,B122,B123,B124,B125,B126,B127,B128", strFileName), ",") 'これ直書きするの？マジで！？ 2

                    For j = 0 To modInput.BGM_LANE - 1

                        g_VGrid(modDraw.GRID.NUM_BGM + j).strText = strArray(j)
                        g_VGrid(modDraw.GRID.NUM_BGM + j).lngBackColor = Color

                    Next j

                    g_lngPenColor(modDraw.PEN_NUM.BGM_LIGHT) = GetColor("BGM", "ObjectLight", "255,0,0", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.BGM_SHADOW) = GetColor("BGM", "ObjectShadow", "96,0,0", strFileName)
                    g_lngBrushColor(modDraw.BRUSH_NUM.BGM) = GetColor("BGM", "ObjectColor", "128,0,0", strFileName)

                Case modDraw.BRUSH_NUM.BPM

                    strArray = Split(strGet_ini("BPM", "Text", "SPEED,SCROLL,BPM,STOP", strFileName), ",")
                    g_VGrid(modDraw.GRID.NUM_SPEED).strText = strArray(0)
                    g_VGrid(modDraw.GRID.NUM_SCROLL).strText = strArray(1)
                    g_VGrid(modDraw.GRID.NUM_BPM).strText = strArray(2)
                    g_VGrid(modDraw.GRID.NUM_STOP).strText = strArray(3)

                    Color = GetColor("BPM", "Background", "48,48,48", strFileName)
                    g_VGrid(modDraw.GRID.NUM_SPEED).lngBackColor = Color
                    g_VGrid(modDraw.GRID.NUM_SCROLL).lngBackColor = Color
                    g_VGrid(modDraw.GRID.NUM_BPM).lngBackColor = Color
                    g_VGrid(modDraw.GRID.NUM_STOP).lngBackColor = Color

                    g_lngPenColor(modDraw.PEN_NUM.BPM_LIGHT) = GetColor("BPM", "ObjectLight", "192,192,0", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.BPM_SHADOW) = GetColor("BPM", "ObjectShadow", "128,128,0", strFileName)
                    g_lngBrushColor(modDraw.BRUSH_NUM.BPM) = GetColor("BPM", "ObjectColor", "160,160,0", strFileName)

                Case modDraw.BRUSH_NUM.BGA

                    strArray = Split(strGet_ini("BGA", "Text", "BGA,LAYER,POOR", strFileName), ",")
                    g_VGrid(modDraw.GRID.NUM_BGA).strText = strArray(0)
                    g_VGrid(modDraw.GRID.NUM_LAYER).strText = strArray(1)
                    g_VGrid(modDraw.GRID.NUM_POOR).strText = strArray(2)

                    Color = GetColor("BGA", "Background", "0,24,0", strFileName)
                    g_VGrid(modDraw.GRID.NUM_BGA).lngBackColor = Color
                    g_VGrid(modDraw.GRID.NUM_LAYER).lngBackColor = Color
                    g_VGrid(modDraw.GRID.NUM_POOR).lngBackColor = Color

                    g_lngPenColor(modDraw.PEN_NUM.BGA_LIGHT) = GetColor("BGA", "ObjectLight", "0,255,0", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.BGA_SHADOW) = GetColor("BGA", "ObjectShadow", "0,96,0", strFileName)
                    g_lngBrushColor(modDraw.BRUSH_NUM.BGA) = GetColor("BGA", "ObjectColor", "0,128,0", strFileName)

                Case modDraw.BRUSH_NUM.KEY01, modDraw.BRUSH_NUM.KEY03, modDraw.BRUSH_NUM.KEY05, modDraw.BRUSH_NUM.KEY07

                    lngTemp = (i - modDraw.BRUSH_NUM.KEY01) + 1

                    g_VGrid(modDraw.GRID.NUM_1P_1KEY + lngTemp - 1).strText = strGet_ini("KEY_1P_0" & lngTemp, "Text", lngTemp, strFileName)

                    g_VGrid(modDraw.GRID.NUM_1P_1KEY + lngTemp - 1).lngBackColor = GetColor("KEY_1P_0" & lngTemp, "Background", "32,32,32", strFileName)

                    Color = GetColor("KEY_1P_0" & lngTemp, "ObjectLight", "192,192,192", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.KEY01_LIGHT + lngTemp - 1) = Color
                    g_lngPenColor(modDraw.PEN_NUM.INV_KEY01_LIGHT + lngTemp - 1) = HalfColor(Color)
                    Color = GetColor("KEY_1P_0" & lngTemp, "ObjectShadow", "96,96,96", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.KEY01_SHADOW + lngTemp - 1) = Color
                    g_lngPenColor(modDraw.PEN_NUM.INV_KEY01_SHADOW + lngTemp - 1) = HalfColor(Color)
                    Color = GetColor("KEY_1P_0" & lngTemp, "ObjectColor", "128,128,128", strFileName)
                    g_lngBrushColor(modDraw.BRUSH_NUM.KEY01 + lngTemp - 1) = Color
                    g_lngBrushColor(modDraw.BRUSH_NUM.INV_KEY01 + lngTemp - 1) = HalfColor(Color)

                Case modDraw.BRUSH_NUM.KEY02, modDraw.BRUSH_NUM.KEY04, modDraw.BRUSH_NUM.KEY06

                    lngTemp = (i - modDraw.BRUSH_NUM.KEY01) + 1
                    g_VGrid(modDraw.GRID.NUM_1P_1KEY + lngTemp - 1).strText = strGet_ini("KEY_1P_0" & lngTemp, "Text", lngTemp, strFileName)

                    g_VGrid(modDraw.GRID.NUM_1P_1KEY + lngTemp - 1).lngBackColor = GetColor("KEY_1P_0" & lngTemp, "Background", "0,0,40", strFileName)

                    Color = GetColor("KEY_1P_0" & lngTemp, "ObjectLight", "96,96,255", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.KEY01_LIGHT + lngTemp - 1) = Color
                    g_lngPenColor(modDraw.PEN_NUM.INV_KEY01_LIGHT + lngTemp - 1) = HalfColor(Color)
                    Color = GetColor("KEY_1P_0" & lngTemp, "ObjectShadow", "0,0,128", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.KEY01_SHADOW + lngTemp - 1) = Color
                    g_lngPenColor(modDraw.PEN_NUM.INV_KEY01_SHADOW + lngTemp - 1) = HalfColor(Color)
                    Color = GetColor("KEY_1P_0" & lngTemp, "ObjectColor", "0,0,255", strFileName)
                    g_lngBrushColor(modDraw.BRUSH_NUM.KEY01 + lngTemp - 1) = Color
                    g_lngBrushColor(modDraw.BRUSH_NUM.INV_KEY01 + lngTemp - 1) = HalfColor(Color)

                Case modDraw.BRUSH_NUM.KEY08

                    g_VGrid(modDraw.GRID.NUM_1P_SC_L).strText = strGet_ini("KEY_1P_SC", "Text", "SC", strFileName)
                    g_VGrid(modDraw.GRID.NUM_1P_SC_R).strText = strGet_ini("KEY_1P_SC", "Text", "SC", strFileName)

                    Color = GetColor("KEY_1P_SC", "Background", "48,0,0", strFileName)
                    g_VGrid(modDraw.GRID.NUM_1P_SC_L).lngBackColor = Color
                    g_VGrid(modDraw.GRID.NUM_1P_SC_R).lngBackColor = Color

                    Color = GetColor("KEY_1P_SC", "ObjectLight", "255,96,96", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.KEY08_LIGHT) = Color
                    g_lngPenColor(modDraw.PEN_NUM.INV_KEY08_LIGHT) = HalfColor(Color)
                    Color = GetColor("KEY_1P_SC", "ObjectShadow", "128,0,0", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.KEY08_SHADOW) = Color
                    g_lngPenColor(modDraw.PEN_NUM.INV_KEY08_SHADOW) = HalfColor(Color)
                    Color = GetColor("KEY_1P_SC", "ObjectColor", "255,0,0", strFileName)
                    g_lngBrushColor(modDraw.BRUSH_NUM.KEY08) = Color
                    g_lngBrushColor(modDraw.BRUSH_NUM.INV_KEY08) = HalfColor(Color)

                Case modDraw.BRUSH_NUM.KEY11, modDraw.BRUSH_NUM.KEY13, modDraw.BRUSH_NUM.KEY15, modDraw.BRUSH_NUM.KEY17

                    lngTemp = (i - modDraw.BRUSH_NUM.KEY11) + 1
                    g_VGrid(modDraw.GRID.NUM_2P_1KEY + lngTemp - 1).strText = strGet_ini("KEY_2P_0" & lngTemp, "Text", lngTemp, strFileName)

                    g_VGrid(modDraw.GRID.NUM_2P_1KEY + lngTemp - 1).lngBackColor = GetColor("KEY_2P_0" & lngTemp, "Background", "32,32,32", strFileName)

                    Color = GetColor("KEY_2P_0" & lngTemp, "ObjectLight", "192,192,192", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.KEY11_LIGHT + lngTemp - 1) = Color
                    g_lngPenColor(modDraw.PEN_NUM.INV_KEY11_LIGHT + lngTemp - 1) = HalfColor(Color)
                    Color = GetColor("KEY_2P_0" & lngTemp, "ObjectShadow", "96,96,96", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.KEY11_SHADOW + lngTemp - 1) = Color
                    g_lngPenColor(modDraw.PEN_NUM.INV_KEY11_SHADOW + lngTemp - 1) = HalfColor(Color)
                    Color = GetColor("KEY_2P_0" & lngTemp, "ObjectColor", "128,128,128", strFileName)
                    g_lngBrushColor(modDraw.BRUSH_NUM.KEY11 + lngTemp - 1) = Color
                    g_lngBrushColor(modDraw.BRUSH_NUM.INV_KEY11 + lngTemp - 1) = HalfColor(Color)

                    If i = modDraw.BRUSH_NUM.KEY11 Then

                        g_VGrid(modDraw.GRID.NUM_FOOTPEDAL).strText = strGet_ini("KEY_2P_01", "Text", lngTemp, strFileName)
                        g_VGrid(modDraw.GRID.NUM_FOOTPEDAL).lngBackColor = GetColor("KEY_2P_01", "Background", "32,32,32", strFileName)
                        'color = GetColor("KEY_2P_0" & lngTemp, "ObjectLight", "192,192,192", strFileName)

                    End If

                Case modDraw.BRUSH_NUM.KEY12, modDraw.BRUSH_NUM.KEY14, modDraw.BRUSH_NUM.KEY16

                    lngTemp = (i - modDraw.BRUSH_NUM.KEY11) + 1
                    g_VGrid(modDraw.GRID.NUM_2P_1KEY + lngTemp - 1).strText = strGet_ini("KEY_2P_0" & lngTemp, "Text", lngTemp, strFileName)

                    g_VGrid(modDraw.GRID.NUM_2P_1KEY + lngTemp - 1).lngBackColor = GetColor("KEY_2P_0" & lngTemp, "Background", "0,0,40", strFileName)

                    Color = GetColor("KEY_2P_0" & lngTemp, "ObjectLight", "96,96,255", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.KEY11_LIGHT + lngTemp - 1) = Color
                    g_lngPenColor(modDraw.PEN_NUM.INV_KEY11_LIGHT + lngTemp - 1) = HalfColor(Color)
                    Color = GetColor("KEY_2P_0" & lngTemp, "ObjectShadow", "0,0,128", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.KEY11_SHADOW + lngTemp - 1) = Color
                    g_lngPenColor(modDraw.PEN_NUM.INV_KEY11_SHADOW + lngTemp - 1) = HalfColor(Color)
                    Color = GetColor("KEY_2P_0" & lngTemp, "ObjectColor", "0,0,255", strFileName)
                    g_lngBrushColor(modDraw.BRUSH_NUM.KEY11 + lngTemp - 1) = Color
                    g_lngBrushColor(modDraw.BRUSH_NUM.INV_KEY11 + lngTemp - 1) = HalfColor(Color)

                Case modDraw.BRUSH_NUM.KEY18

                    g_VGrid(modDraw.GRID.NUM_2P_SC_L).strText = strGet_ini("KEY_2P_SC", "Text", "SC", strFileName)
                    g_VGrid(modDraw.GRID.NUM_2P_SC_R).strText = strGet_ini("KEY_2P_SC", "Text", "SC", strFileName)

                    Color = GetColor("KEY_2P_SC", "Background", "48,0,0", strFileName)
                    g_VGrid(modDraw.GRID.NUM_2P_SC_L).lngBackColor = Color
                    g_VGrid(modDraw.GRID.NUM_2P_SC_R).lngBackColor = Color

                    Color = GetColor("KEY_2P_SC", "ObjectLight", "255,96,96", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.KEY18_LIGHT) = Color
                    g_lngPenColor(modDraw.PEN_NUM.INV_KEY18_LIGHT) = HalfColor(Color)
                    Color = GetColor("KEY_2P_SC", "ObjectShadow", "128,0,0", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.KEY18_SHADOW) = Color
                    g_lngPenColor(modDraw.PEN_NUM.INV_KEY18_SHADOW) = HalfColor(Color)
                    Color = GetColor("KEY_2P_SC", "ObjectColor", "255,0,0", strFileName)
                    g_lngBrushColor(modDraw.BRUSH_NUM.KEY18) = Color
                    g_lngBrushColor(modDraw.BRUSH_NUM.INV_KEY18) = HalfColor(Color)

                Case modDraw.BRUSH_NUM.LONGNOTE

                    g_lngPenColor(modDraw.PEN_NUM.LONGNOTE_LIGHT) = GetColor("KEY_LONGNOTE", "ObjectLight", "0,128,0", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.LONGNOTE_SHADOW) = GetColor("KEY_LONGNOTE", "ObjectShadow", "0,32,0", strFileName)
                    g_lngBrushColor(modDraw.BRUSH_NUM.LONGNOTE) = GetColor("KEY_LONGNOTE", "ObjectColor", "0,64,0", strFileName)

                Case modDraw.BRUSH_NUM.MINE

                    g_lngPenColor(modDraw.PEN_NUM.MINE_LIGHT) = GetColor("KEY_MINE", "ObjectLight", "64,0,64", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.MINE_SHADOW) = GetColor("KEY_MINE", "ObjectShadow", "16,0,16", strFileName)
                    g_lngBrushColor(modDraw.BRUSH_NUM.MINE) = GetColor("KEY_MINE", "ObjectColor", "32,0,32", strFileName)

                Case modDraw.BRUSH_NUM.SELECT_OBJ

                    g_lngPenColor(modDraw.PEN_NUM.SELECT_OBJ_LIGHT) = GetColor("SELECT", "ObjectLight", "255,255,255", strFileName)
                    g_lngPenColor(modDraw.PEN_NUM.SELECT_OBJ_SHADOW) = GetColor("SELECT", "ObjectShadow", "128,128,128", strFileName)
                    g_lngBrushColor(modDraw.BRUSH_NUM.SELECT_OBJ) = GetColor("SELECT", "ObjectColor", "0,255,255", strFileName)

                Case modDraw.BRUSH_NUM.EDIT_FRAME

                    g_lngPenColor(modDraw.PEN_NUM.EDIT_FRAME) = GetColor("SELECT", "EditFrame", "255,255,255", strFileName)

                Case modDraw.BRUSH_NUM.DELETE_FRAME

                    g_lngPenColor(modDraw.PEN_NUM.DELETE_FRAME) = GetColor("SELECT", "DeleteFrame", "255,255,255", strFileName)

            End Select

        Next i

    End Sub

    Public Sub LoadLanguageFile(ByRef strFileName As String)
        g_strStatusBar(1) = strGet_ini("StatusBar", "CH_01", "BGM", strFileName)
        g_strStatusBar(4) = strGet_ini("StatusBar", "CH_04", "BGA", strFileName)
        g_strStatusBar(6) = strGet_ini("StatusBar", "CH_06", "BGA Poor", strFileName)
        g_strStatusBar(7) = strGet_ini("StatusBar", "CH_07", "BGA Layer", strFileName)
        g_strStatusBar(8) = strGet_ini("StatusBar", "CH_08", "BPM Change", strFileName)
        g_strStatusBar(9) = strGet_ini("StatusBar", "CH_09", "Stop Sequence", strFileName)
        g_strStatusBar(11) = strGet_ini("StatusBar", "CH_KEY_1P", "1P Key", strFileName)
        g_strStatusBar(12) = strGet_ini("StatusBar", "CH_KEY_2P", "2P Key", strFileName)
        g_strStatusBar(13) = strGet_ini("StatusBar", "CH_SCRATCH_1P", "1P Scratch", strFileName)
        g_strStatusBar(14) = strGet_ini("StatusBar", "CH_SCRATCH_2P", "2P Scratch", strFileName)
        g_strStatusBar(15) = strGet_ini("StatusBar", "CH_INVISIBLE", "(Invisible)", strFileName)
        g_strStatusBar(16) = strGet_ini("StatusBar", "CH_LONGNOTE", "(LongNote)", strFileName)
        g_strStatusBar(17) = strGet_ini("StatusBar", "CH_MINE", "(LandMine)", strFileName)
        g_strStatusBar(20) = strGet_ini("StatusBar", "MODE_EDIT", "Edit Mode", strFileName)
        g_strStatusBar(21) = strGet_ini("StatusBar", "MODE_WRITE", "Write Mode", strFileName)
        g_strStatusBar(22) = strGet_ini("StatusBar", "MODE_DELETE", "Delete Mode", strFileName)
        g_strStatusBar(23) = strGet_ini("StatusBar", "MEASURE", "Measure", strFileName)
        g_strStatusBar(24) = strGet_ini("StatusBar", "CH_SCROLL", "SCROLL", strFileName)
        g_strStatusBar(25) = strGet_ini("StatusBar", "CH_SPEED", "SPEED", strFileName)

        With frmMain

            .mnuFile.Text = strGet_ini("Menu", "FILE", "&File", strFileName)
            .mnuFileNew.Text = strGet_ini("Menu", "FILE_NEW", "&New", strFileName)
            .mnuFileOpen.Text = strGet_ini("Menu", "FILE_OPEN", "&Open", strFileName)
            .mnuFileSave.Text = strGet_ini("Menu", "FILE_SAVE", "&Save", strFileName)
            .mnuFileSaveAs.Text = strGet_ini("Menu", "FILE_SAVE_AS", "Save &As", strFileName)
            .mnuFileOpenDirectory.Text = strGet_ini("Menu", "FILE_OPEN_DIRECTORY", "Open &Directory", strFileName)
            '.mnuFileDeleteUnusedFile.Caption = strGet_ini("Menu", "FILE_DELETE_UNUSED_FILE", "&Delete Unused File(s)", strFileName)
            '.mnuFileNameConvert.Caption = strGet_ini("Menu", "FILE_CONVERT_FILENAME", "&Convert Filenames to [01-ZZ]", strFileName)
            '.mnuFileListAlign.Caption = strGet_ini("Menu", "FILE_ALIGN_LIST", "Rewrite &List into old format [01-FF]", strFileName)
            .mnuFileConvertWizard.Text = strGet_ini("Menu", "FILE_CONVERT_WIZARD", "Show &Conversion Wizard", strFileName)
            .mnuFileExit.Text = strGet_ini("Menu", "FILE_EXIT", "&Exit", strFileName)

            .mnuEdit.Text = strGet_ini("Menu", "EDIT", "&Edit", strFileName)
            .mnuEditUndo.Text = strGet_ini("Menu", "EDIT_UNDO", "&Undo", strFileName)
            .mnuEditRedo.Text = strGet_ini("Menu", "EDIT_REDO", "&Redo", strFileName)
            .mnuEditCut.Text = strGet_ini("Menu", "EDIT_CUT", "Cu&t", strFileName)
            .mnuEditCopy.Text = strGet_ini("Menu", "EDIT_COPY", "&Copy", strFileName)
            .mnuEditPaste.Text = strGet_ini("Menu", "EDIT_PASTE", "&Paste", strFileName)
            .mnuEditDelete.Text = strGet_ini("Menu", "EDIT_DELETE", "&Delete", strFileName)
            .mnuEditSelectAll.Text = strGet_ini("Menu", "EDIT_SELECT_ALL", "&Find/Replace/Delete", strFileName)
            .mnuEditFind.Text = strGet_ini("Menu", "EDIT_FIND", "&Select All", strFileName)
            ._mnuEditMode_0.Text = strGet_ini("Menu", "EDIT_MODE_EDIT", "Edit &Mode", strFileName)
            ._mnuEditMode_1.Text = strGet_ini("Menu", "EDIT_MODE_WRITE", "Write &Mode", strFileName)
            ._mnuEditMode_2.Text = strGet_ini("Menu", "EDIT_MODE_DELETE", "Delete &Mode", strFileName)

            .mnuView.Text = strGet_ini("Menu", "VIEW", "&View", strFileName)
            ._mnuViewItem_0.Text = strGet_ini("Menu", "VIEW_TOOL_BAR", "&Tool Bar", strFileName)
            ._mnuViewItem_1.Text = strGet_ini("Menu", "VIEW_DIRECT_INPUT", "&Direct Input", strFileName)
            ._mnuViewItem_2.Text = strGet_ini("Menu", "VIEW_STATUS_BAR", "&Status Bar", strFileName)

            ._mnuViewItem_0_New.Text = strGet_ini("Menu", "VIEW_TOOL_BAR_NEW", "New", strFileName)
            ._mnuViewItem_0_Open.Text = strGet_ini("Menu", "VIEW_TOOL_BAR_OPEN", "Open", strFileName)
            ._mnuViewItem_0_Reload.Text = strGet_ini("Menu", "VIEW_TOOL_BAR_RELOAD", "Reload", strFileName)
            ._mnuViewItem_0_Save.Text = strGet_ini("Menu", "VIEW_TOOL_BAR_SAVE", "Save", strFileName)
            ._mnuViewItem_0_SaveAs.Text = strGet_ini("Menu", "VIEW_TOOL_BAR_SAVEAS", "SaveAs", strFileName)
            ._mnuViewItem_0_Mode.Text = strGet_ini("Menu", "VIEW_TOOL_BAR_MODE", "Mode", strFileName)
            ._mnuViewItem_0_Preview.Text = strGet_ini("Menu", "VIEW_TOOL_BAR_PREVIEW", "Preview", strFileName)
            ._mnuViewItem_0_Grid.Text = strGet_ini("Menu", "VIEW_TOOL_BAR_GRID", "Grid", strFileName)
            ._mnuViewItem_0_Size.Text = strGet_ini("Menu", "VIEW_TOOL_BAR_SIZE", "Size", strFileName)
            ._mnuViewItem_0_Resolution.Text = strGet_ini("Menu", "VIEW_TOOL_BAR_RESOLUTION", "Resolution", strFileName)

            .mnuOptions.Text = strGet_ini("Menu", "OPTIONS", "&Options", strFileName)
            ._mnuOptionsItem_0.Text = strGet_ini("Menu", "OPTIONS_IGNORE_ACTIVE", "&Control Unavailable When Active", strFileName)
            ._mnuOptionsItem_1.Text = strGet_ini("Menu", "OPTIONS_FILE_NAME_ONLY", "Display &File Name Only", strFileName)
            ._mnuOptionsItem_2.Text = strGet_ini("Menu", "OPTIONS_VERTICAL", "&Vertical Grid Info", strFileName)
            ._mnuOptionsItem_3.Text = strGet_ini("Menu", "OPTIONS_LANE_BG", "&Background Color", strFileName)
            '.mnuOptionsItem(SELECT_PREVIEW).Caption = strGet_ini("Menu", "OPTIONS_SINGLE_SELECT_SOUND", "&Sound Upon Object Selection", strFileName)
            ._mnuOptionsItem_4.Text = strGet_ini("Menu", "OPTIONS_SINGLE_SELECT_PREVIEW", "&Preview Upon Object Selection", strFileName)
            ._mnuOptionsItem_6.Text = strGet_ini("Menu", "OPTIONS_OBJECT_FILE_NAME", "Show &Objects' File Names", strFileName)
            ._mnuOptionsItem_5.Text = strGet_ini("Menu", "OPTIONS_MOVE_ON_GRID", "Restrict Objects' &Movement Onto Grid", strFileName)
            ._mnuOptionsItem_7.Text = strGet_ini("Menu", "OPTIONS_USE_NEW_FORMAT", "&Use New Base62 Format (01-ZZ-zz)", strFileName)
            ._mnuOptionsItem_8.Text = strGet_ini("Menu", "OPTIONS_Y_AXIS_FIXED", "Y-Axis Fixed", strFileName)
            ._mnuOptionsItem_9.Text = strGet_ini("Menu", "OPTIONS_ENABLE_TOOLTIP", "Enable Tooltip of Object", strFileName)
            '.mnuOptionsItem(RCLICK_DELETE).Caption = strGet_ini("Menu", "OPTIONS_RIGHT_CLICK_DELETE", "&Right Click To Delete Objects", strFileName)
            ._mnuOptionsBaseCaution.Text = strGet_ini("Menu", "OPTIONS_BASE_CAUTION", "CAUTION: Don't change During edit.", strFileName)
            ._mnuOptionsBase16.Text = strGet_ini("Menu", "OPTIONS_BASE16", "Prefer Base16 (FF Definition)", strFileName)
            ._mnuOptionsBase36.Text = strGet_ini("Menu", "OPTIONS_BASE36", "Use Base36 (ZZ Definition)", strFileName)
            ._mnuOptionsBase62.Text = strGet_ini("Menu", "OPTIONS_BASE62", "Use Base62 (zz Case Sensitive)", strFileName)

            .mnuTools.Text = strGet_ini("Menu", "TOOLS", "&Tools", strFileName)
            .mnuToolsPlayAll.Text = strGet_ini("Menu", "TOOLS_PLAY_FIRST", "Play &All", strFileName)
            .mnuToolsPlay.Text = strGet_ini("Menu", "TOOLS_PLAY", "&Play From Current Position", strFileName)
            .mnuToolsPlayStop.Text = strGet_ini("Menu", "TOOLS_STOP", "&Stop", strFileName)
            .mnuToolsSetting.Text = strGet_ini("Menu", "TOOLS_SETTING", "&Viewer Setting", strFileName)

            .mnuHelp.Text = strGet_ini("Menu", "HELP", "&Help", strFileName)
            .mnuHelpOpen.Text = strGet_ini("Menu", "HELP_OPEN", "&Help", strFileName)
            .mnuHelpWeb.Text = strGet_ini("Menu", "HELP_WEB", "Open &Website", strFileName)
            .mnuHelpAbout.Text = strGet_ini("Menu", "HELP_ABOUT", "&About BMSE", strFileName)

            .mnuContext.Visible = False
            .mnuContextPlayAll.Text = strGet_ini("Menu", "TOOLS_PLAY_FIRST", "Play &All", strFileName)
            .mnuContextPlay.Text = strGet_ini("Menu", "TOOLS_PLAY", "&Play From Current Position", strFileName)
            .mnuContextInsertMeasure.Text = strGet_ini("Menu", "CONTEXT_MEASURE_INSERT", "&Insert Measure", strFileName)
            .mnuContextDeleteMeasure.Text = strGet_ini("Menu", "CONTEXT_MEASURE_DELETE", "Delete &Measure", strFileName)
            .mnuContextEditCut.Text = strGet_ini("Menu", "EDIT_CUT", "Cu&t", strFileName)
            .mnuContextEditCopy.Text = strGet_ini("Menu", "EDIT_COPY", "&Copy", strFileName)
            .mnuContextEditPaste.Text = strGet_ini("Menu", "EDIT_PASTE", "&Paste", strFileName)
            .mnuContextEditDelete.Text = strGet_ini("Menu", "EDIT_DELETE", "&Delete", strFileName)

            .mnuContextList.Visible = False
            .mnuContextListLoad.Text = strGet_ini("Menu", "CONTEXT_LIST_LOAD", "&Load", strFileName)
            .mnuContextListDelete.Text = strGet_ini("Menu", "CONTEXT_LIST_DELETE", "&Delete", strFileName)
            .mnuContextListRename.Text = strGet_ini("Menu", "CONTEXT_LIST_RENAME", "&Rename", strFileName)

            ._optChangeTop_0.Text = strGet_ini("Header", "TAB_BASIC", "Basic", strFileName)
            ._optChangeTop_1.Text = strGet_ini("Header", "TAB_EXPAND", "Expand", strFileName)
            ._optChangeTop_2.Text = strGet_ini("Header", "TAB_CONFIG", "Config", strFileName)

            .lblPlayMode.Text = strGet_ini("Header", "BASIC_PLAYER", "#PLAYER", strFileName)
            .cboPlayer.Items.Item(0) = strGet_ini("Header", "BASIC_PLAYER_1P", "1 Player", strFileName)
            .cboPlayer.Items.Item(1) = strGet_ini("Header", "BASIC_PLAYER_2P", "2 Player", strFileName)
            .cboPlayer.Items.Item(2) = strGet_ini("Header", "BASIC_PLAYER_DP", "Double Play", strFileName)
            .cboPlayer.Items.Item(3) = strGet_ini("Header", "BASIC_PLAYER_PMS", "9 Keys (PMS)", strFileName)
            .cboPlayer.Items.Item(4) = strGet_ini("Header", "BASIC_PLAYER_OCT", "13 Keys (Oct)", strFileName)
            .lblGenre.Text = strGet_ini("Header", "BASIC_GENRE", "#GENRE", strFileName)
            .lblTitle.Text = strGet_ini("Header", "BASIC_TITLE", "#TITLE", strFileName)
            .lblArtist.Text = strGet_ini("Header", "BASIC_ARTIST", "#ARTIST", strFileName)
            .lblPlayLevel.Text = strGet_ini("Header", "BASIC_PLAYLEVEL", "#PLAYLEVEL", strFileName)
            .lblBPM.Text = strGet_ini("Header", "BASIC_BPM", "#BPM", strFileName)

            .lblPlayRank.Text = strGet_ini("Header", "EXPAND_RANK", "#RANK", strFileName)
            .cboPlayRank.Items.Item(0) = strGet_ini("Header", "EXPAND_RANK_VERY_HARD", "Very Hard", strFileName)
            .cboPlayRank.Items.Item(1) = strGet_ini("Header", "EXPAND_RANK_HARD", "Hard", strFileName)
            .cboPlayRank.Items.Item(2) = strGet_ini("Header", "EXPAND_RANK_NORMAL", "Normal", strFileName)
            .cboPlayRank.Items.Item(3) = strGet_ini("Header", "EXPAND_RANK_EASY", "Easy", strFileName)
            .lblTotal.Text = strGet_ini("Header", "EXPAND_TOTAL", "#TOTAL", strFileName)
            .lblVolume.Text = strGet_ini("Header", "EXPAND_VOLWAV", "#VOLWAV", strFileName)
            .lblStageFile.Text = strGet_ini("Header", "EXPAND_STAGEFILE", "#STAGEFILE", strFileName)
            .lblMissBMP.Text = strGet_ini("Header", "EXPAND_BPM_MISS", "#BMP00", strFileName)
            .cmdLoadMissBMP.Text = strGet_ini("Header", "EXPAND_SET_FILE", "...", strFileName)
            .cmdLoadStageFile.Text = strGet_ini("Header", "EXPAND_SET_FILE", "...", strFileName)

            .ToolTip1.SetToolTip(.lblPlayMode, strGet_ini("Header", "TOOLTIP_PLAYMODE", "", strFileName))
            .ToolTip1.SetToolTip(.cboPlayer, strGet_ini("Header", "TOOLTIP_PLAYMODE", "", strFileName))
            .ToolTip1.SetToolTip(.lblGenre, strGet_ini("Header", "TOOLTIP_GENRE", "", strFileName))
            .ToolTip1.SetToolTip(.txtGenre, strGet_ini("Header", "TOOLTIP_GENRE", "", strFileName))
            .ToolTip1.SetToolTip(.lblTitle, strGet_ini("Header", "TOOLTIP_TITLE", "can't omit.", strFileName))
            .ToolTip1.SetToolTip(.txtTitle, strGet_ini("Header", "TOOLTIP_TITLE", "can't omit.", strFileName))
            .ToolTip1.SetToolTip(.lblArtist, strGet_ini("Header", "TOOLTIP_ARTIST", "", strFileName))
            .ToolTip1.SetToolTip(.txtArtist, strGet_ini("Header", "TOOLTIP_ARTIST", "", strFileName))
            .ToolTip1.SetToolTip(.lblPlayLevel, strGet_ini("Header", "TOOLTIP_PLAYLEVEL", "must be positive integer.", strFileName))
            .ToolTip1.SetToolTip(.cboPlayLevel, strGet_ini("Header", "TOOLTIP_PLAYLEVEL", "must be positive integer.", strFileName))
            .ToolTip1.SetToolTip(.lblBPM, strGet_ini("Header", "TOOLTIP_BPM", "=Beat Per Minute. can't omit.(>0)", strFileName))
            .ToolTip1.SetToolTip(.txtBPM, strGet_ini("Header", "TOOLTIP_BPM", "=Beat Per Minute. can't omit.(>0)", strFileName))

            .ToolTip1.SetToolTip(.lblPlayRank, strGet_ini("Header", "TOOLTIP_RANK", "Strictness of judgment time.", strFileName))
            .ToolTip1.SetToolTip(.cboPlayRank, strGet_ini("Header", "TOOLTIP_RANK", "Strictness of judgment time.", strFileName))
            .ToolTip1.SetToolTip(.lblTotal, strGet_ini("Header", "TOOLTIP_TOTAL", "Maximum gauge increase [%].", strFileName))
            .ToolTip1.SetToolTip(.txtTotal, strGet_ini("Header", "TOOLTIP_TOTAL", "Maximum gauge increase [%].", strFileName))
            .ToolTip1.SetToolTip(.lblVolume, strGet_ini("Header", "TOOLTIP_VOLUME", "", strFileName))
            .ToolTip1.SetToolTip(.txtVolume, strGet_ini("Header", "TOOLTIP_VOLUME", "", strFileName))
            .ToolTip1.SetToolTip(.lblStageFile, strGet_ini("Header", "TOOLTIP_STAGEFILE", "Image shown while loading.", strFileName))
            .ToolTip1.SetToolTip(.txtStageFile, strGet_ini("Header", "TOOLTIP_STAGEFILE", "Image shown while loading.", strFileName))
            .ToolTip1.SetToolTip(.lblMissBMP, strGet_ini("Header", "TOOLTIP_BMP_MISS", "Default Image shown when player misses.", strFileName))
            .ToolTip1.SetToolTip(.txtMissBMP, strGet_ini("Header", "TOOLTIP_BMP_MISS", "Default Image shown when player misses.", strFileName))

            .lblDispFrame.Text = strGet_ini("Header", "CONFIG_KEY_FRAME", "Key Frame", strFileName)
            .cboDispFrame.Items.Item(0) = strGet_ini("Header", "CONFIG_KEY_HALF", "Half", strFileName)
            .cboDispFrame.Items.Item(1) = strGet_ini("Header", "CONFIG_KEY_SEPARATE", "Separate", strFileName)
            .lblDispKey.Text = strGet_ini("Header", "CONFIG_KEY_POSITION", "Key Position", strFileName)
            .cboDispKey.Items.Item(0) = strGet_ini("Header", "CONFIG_KEY_5KEYS", "5Keys/10Keys", strFileName)
            .cboDispKey.Items.Item(1) = strGet_ini("Header", "CONFIG_KEY_7KEYS", "7Keys/14Keys", strFileName)
            .lblDispSC1P.Text = strGet_ini("Header", "CONFIG_SCRATCH_1P", "Scratch 1P", strFileName)
            .cboDispSC1P.Items.Item(0) = strGet_ini("Header", "CONFIG_SCRATCH_LEFT", "L", strFileName)
            .cboDispSC1P.Items.Item(1) = strGet_ini("Header", "CONFIG_SCRATCH_RIGHT", "R", strFileName)
            .lblDispSC2P.Text = strGet_ini("Header", "CONFIG_SCRATCH_2P", "2P", strFileName)
            .cboDispSC2P.Items.Item(0) = strGet_ini("Header", "CONFIG_SCRATCH_LEFT", "L", strFileName)
            .cboDispSC2P.Items.Item(1) = strGet_ini("Header", "CONFIG_SCRATCH_RIGHT", "R", strFileName)

            ._optChangeBottom_0.Text = strGet_ini("Material", "TAB_WAV", "#WAV", strFileName)
            ._optChangeBottom_1.Text = strGet_ini("Material", "TAB_BMP", "#BMP", strFileName)
            ._optChangeBottom_2.Text = strGet_ini("Material", "TAB_BGA", "#BGA", strFileName)
            ._optChangeBottom_3.Text = strGet_ini("Material", "TAB_BEAT", "Beat", strFileName)
            ._optChangeBottom_4.Text = strGet_ini("Material", "TAB_EXPAND", "Expand", strFileName)

            .cmdSoundStop.Text = strGet_ini("Material", "MATERIAL_STOP", "Stop", strFileName)
            .cmdSoundExcUp.Text = strGet_ini("Material", "MATERIAL_EXCHANGE_UP", "<", strFileName)
            .cmdSoundExcDown.Text = strGet_ini("Material", "MATERIAL_EXCHANGE_DOWN", ">", strFileName)
            .cmdSoundDelete.Text = strGet_ini("Material", "MATERIAL_DELETE", "Del", strFileName)
            .cmdSoundLoad.Text = strGet_ini("Material", "MATERIAL_SET_FILE", "...", strFileName)

            .cmdBMPPreview.Text = strGet_ini("Material", "MATERIAL_PREVIEW", "Preview", strFileName)
            .cmdBMPExcUp.Text = strGet_ini("Material", "MATERIAL_EXCHANGE_UP", "<", strFileName)
            .cmdBMPExcDown.Text = strGet_ini("Material", "MATERIAL_EXCHANGE_DOWN", ">", strFileName)
            .cmdBMPDelete.Text = strGet_ini("Material", "MATERIAL_DELETE", "Del", strFileName)
            .cmdBMPLoad.Text = strGet_ini("Material", "MATERIAL_SET_FILE", "...", strFileName)

            .cmdBGAPreview.Text = strGet_ini("Material", "MATERIAL_PREVIEW", "Preview", strFileName)
            .cmdBGAExcUp.Text = strGet_ini("Material", "MATERIAL_EXCHANGE_UP", "<", strFileName)
            .cmdBGAExcDown.Text = strGet_ini("Material", "MATERIAL_EXCHANGE_DOWN", ">", strFileName)
            .cmdBGADelete.Text = strGet_ini("Material", "MATERIAL_DELETE", "Del", strFileName)
            .cmdBGASet.Text = strGet_ini("Material", "MATERIAL_INPUT", "Input", strFileName)

            .cmdMeasureSelectAll.Text = strGet_ini("Material", "MATERIAL_SELECT_ALL", "All", strFileName)

            .cmdInputMeasureLen.Text = strGet_ini("Material", "MATERIAL_INPUT", "Input", strFileName)

            .lblSubTitle.Text = strGet_ini("Material", "EX_SUBTITLE", "#SUBTITLE", strFileName)
            .lblSubArtist.Text = strGet_ini("Material", "EX_SUBARTIST", "#SUBARTIST", strFileName)
            .cboDifficulty.Items.Item(0) = strGet_ini("Material", "EX_DIFFICULTY_0", "(None)", strFileName)
            .cboDifficulty.Items.Item(1) = strGet_ini("Material", "EX_DIFFICULTY_1", "BEGINNER / EASY", strFileName)
            .cboDifficulty.Items.Item(2) = strGet_ini("Material", "EX_DIFFICULTY_2", "NORMAL", strFileName)
            .cboDifficulty.Items.Item(3) = strGet_ini("Material", "EX_DIFFICULTY_3", "HYPER / HARD", strFileName)
            .cboDifficulty.Items.Item(4) = strGet_ini("Material", "EX_DIFFICULTY_4", "ANOTHER / EX", strFileName)
            .cboDifficulty.Items.Item(5) = strGet_ini("Material", "EX_DIFFICULTY_5", "INSANE / OTHER", strFileName)
            .lblPreview.Text = strGet_ini("Material", "EX_PREVIEW", "#PREVIEW", strFileName)
            .cmdLoadPreview.Text = strGet_ini("Material", "EX_SET_FILE", "...", strFileName)
            .lblBanner.Text = strGet_ini("Material", "EX_BANNER", "#BANNER", strFileName)
            .cmdLoadBanner.Text = strGet_ini("Material", "EX_SET_FILE", "...", strFileName)
            .lblBackBmp.Text = strGet_ini("Material", "EX_BACKBMP", "#BACKBMP", strFileName)
            .cmdLoadBackBmp.Text = strGet_ini("Material", "EX_SET_FILE", "...", strFileName)
            .lblLandmineWAV.Text = strGet_ini("Material", "EX_LANDMINEWAV", "#WAV00", strFileName)
            .cmdLoadLandmineWAV.Text = strGet_ini("Material", "EX_SET_FILE", "...", strFileName)
            .lblLNMode.Text = strGet_ini("Material", "EX_LNMODE", "#LNMODE", strFileName)
            .cboLNMode.Items.Item(0) = strGet_ini("Material", "EX_LNMODE_0", "(Selectable)", strFileName)
            .cboLNMode.Items.Item(1) = strGet_ini("Material", "EX_LNMODE_1", "LN Only", strFileName)
            .cboLNMode.Items.Item(2) = strGet_ini("Material", "EX_LNMODE_2", "CN Only", strFileName)
            .cboLNMode.Items.Item(3) = strGet_ini("Material", "EX_LNMODE_3", "HCN Only", strFileName)
            .lblLNObj.Text = strGet_ini("Material", "EX_LNOBJ", "#LNOBJ", strFileName)
            .cboLNObj.Items.Item(0) = strGet_ini("Material", "EX_LNOBJ_0", "(#LNTYPE 1)", strFileName)
            .lblDefExRank.Text = strGet_ini("Material", "EX_DEFEXRANK", "#DEFEXRANK", strFileName)
            .lblComment.Text = strGet_ini("Material", "EX_COMMENT", "#COMMENT", strFileName)
            .lblExInfo.Text = strGet_ini("Material", "EX_EXINFO", "EXTRA INFORMATION", strFileName)

            .ToolTip1.SetToolTip(.lblSubTitle, strGet_ini("Material", "TOOLTIP_SUBTITLE", "", strFileName))
            .ToolTip1.SetToolTip(.txtSubTitle, strGet_ini("Material", "TOOLTIP_SUBTITLE", "", strFileName))
            .ToolTip1.SetToolTip(.lblSubArtist, strGet_ini("Material", "TOOLTIP_SUBARTIST", "", strFileName))
            .ToolTip1.SetToolTip(.txtSubArtist, strGet_ini("Material", "TOOLTIP_SUBARTIST", "", strFileName))
            .ToolTip1.SetToolTip(.lblDifficulty, strGet_ini("Material", "TOOLTIP_DIFFICULTY", "Indicates what type of chart this is. It is recommended to set.", strFileName))
            .ToolTip1.SetToolTip(.cboDifficulty, strGet_ini("Material", "TOOLTIP_DIFFICULTY", "Indicates what type of chart this is. It is recommended to set.", strFileName))
            .ToolTip1.SetToolTip(.lblPreview, strGet_ini("Material", "TOOLTIP_PREVIEW", "Sound played on the music selection scene.", strFileName))
            .ToolTip1.SetToolTip(.txtPreview, strGet_ini("Material", "TOOLTIP_PREVIEW", "Sound played on the music selection scene.", strFileName))
            .ToolTip1.SetToolTip(.lblBanner, strGet_ini("Material", "TOOLTIP_BANNER", "", strFileName))
            .ToolTip1.SetToolTip(.txtBanner, strGet_ini("Material", "TOOLTIP_BANNER", "", strFileName))
            .ToolTip1.SetToolTip(.lblBackBmp, strGet_ini("Material", "TOOLTIP_BACKBMP", "", strFileName))
            .ToolTip1.SetToolTip(.txtBackBmp, strGet_ini("Material", "TOOLTIP_BACKBMP", "", strFileName))
            .ToolTip1.SetToolTip(.lblLandmineWAV, strGet_ini("Material", "TOOLTIP_LANDMINEWAV", "Landmine sound", strFileName))
            .ToolTip1.SetToolTip(.txtLandmineWAV, strGet_ini("Material", "TOOLTIP_LANDMINEWAV", "Landmine sound", strFileName))
            .ToolTip1.SetToolTip(.lblLNMode, strGet_ini("Material", "TOOLTIP_LNMODE", "LR2 does not support this command." & vbCrLf & "LN:KeyUp timing is NOT judged." & vbCrLf & "CN:KeyUp timing is judged." & vbCrLf & "HCN:Constantly deplete the gauge when not pressed.", strFileName))
            .ToolTip1.SetToolTip(.cboLNMode, strGet_ini("Material", "TOOLTIP_LNMODE", "LR2 does not support this command." & vbCrLf & "LN:KeyUp timing is NOT judged." & vbCrLf & "CN:KeyUp timing is judged." & vbCrLf & "HCN:Constantly deplete the gauge when not pressed.", strFileName))
            .ToolTip1.SetToolTip(.lblLNObj, strGet_ini("Material", "TOOLTIP_LNOBJ", "", strFileName))
            .ToolTip1.SetToolTip(.cboLNObj, strGet_ini("Material", "TOOLTIP_LNOBJ", "", strFileName))
            .ToolTip1.SetToolTip(.lblDefExRank, strGet_ini("Material", "TOOLTIP_DEFEXRANK", "Judge timing rate, 100 is same as #RANK:Normal.", strFileName))
            .ToolTip1.SetToolTip(.txtDefExRank, strGet_ini("Material", "TOOLTIP_DEFEXRANK", "Judge timing rate, 100 is same as #RANK:Normal.", strFileName))
            .ToolTip1.SetToolTip(.txtComment, strGet_ini("Material", "TOOLTIP_COMMENT", Chr(34) & "#COMMENT Must be wrapped in double quotes marks." & Chr(34), strFileName))
            .ToolTip1.SetToolTip(.lblComment, strGet_ini("Material", "TOOLTIP_COMMENT", Chr(34) & "#COMMENT Must be wrapped in double quotes marks." & Chr(34), strFileName))

            .lblGridMain.Text = strGet_ini("ToolBar", "GRID_MAIN", "Grid", strFileName)
            .lblGridSub.Text = strGet_ini("ToolBar", "GRID_SUB", "Sub", strFileName)
            .lblDispHeight.Text = strGet_ini("ToolBar", "DISP_HEIGHT", "Height", strFileName)
            .lblDispWidth.Text = strGet_ini("ToolBar", "DISP_WIDTH", "Width", strFileName)
            DirectCast(.cboDispHeight.Items.Item(.cboDispHeight.Items.Count - 1), modMain.ItemWithData).SetItemString(strGet_ini("ToolBar", "DISP_VALUE_OTHER", "Other", strFileName))
            DirectCast(.cboDispWidth.Items.Item(.cboDispWidth.Items.Count - 1), modMain.ItemWithData).SetItemString(strGet_ini("ToolBar", "DISP_VALUE_OTHER", "Other", strFileName))
            .lblVScroll.Text = strGet_ini("ToolBar", "VSCROLL", "VScroll", strFileName)

            If DirectCast(.tlbMenu.Items.Item("Edit"), ToolStripButton).Checked = True Then

                .staMain.Items.Item("Mode").Text = g_strStatusBar(20)

            ElseIf DirectCast(.tlbMenu.Items.Item("Write"), ToolStripButton).Checked = True Then

                .staMain.Items.Item("Mode").Text = g_strStatusBar(21)

            Else

                .staMain.Items.Item("Mode").Text = g_strStatusBar(22)

            End If

        End With

        With frmMain.tlbMenu

            .Items.Item("_New").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_NEW", "New", strFileName)
            .Items.Item("_New").AccessibleDescription = strGet_ini("ToolBar", "TOOLTIP_NEW", "New", strFileName)
            .Items.Item("Open").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_OPEN", "Open", strFileName)
            .Items.Item("Open").AccessibleDescription = strGet_ini("ToolBar", "TOOLTIP_OPEN", "Open", strFileName)
            .Items.Item("Reload").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_RELOAD", "Reload", strFileName)
            .Items.Item("Reload").AccessibleDescription = strGet_ini("ToolBar", "TOOLTIP_RELOAD", "Reload", strFileName)
            .Items.Item("Save").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_SAVE", "Save", strFileName)
            .Items.Item("Save").AccessibleDescription = strGet_ini("ToolBar", "TOOLTIP_SAVE", "Save", strFileName)
            .Items.Item("SaveAs").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_SAVE_AS", "Save As", strFileName)
            .Items.Item("SaveAs").AccessibleDescription = strGet_ini("ToolBar", "TOOLTIP_SAVE_AS", "Save As", strFileName)

            .Items.Item("Edit").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_MODE_EDIT", "Edit Mode", strFileName)
            .Items.Item("Edit").AccessibleDescription = strGet_ini("ToolBar", "TOOLTIP_MODE_EDIT", "Edit Mode", strFileName)
            .Items.Item("Write").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_MODE_WRITE", "Write Mode", strFileName)
            .Items.Item("Write").AccessibleDescription = strGet_ini("ToolBar", "TOOLTIP_MODE_WRITE", "Write Mode", strFileName)
            .Items.Item("Delete").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_MODE_DELETE", "Delete Mode", strFileName)
            .Items.Item("Delete").AccessibleDescription = strGet_ini("ToolBar", "TOOLTIP_MODE_DELETE", "Delete Mode", strFileName)

            .Items.Item("PlayAll").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_PLAY_FIRST", "Play All", strFileName)
            .Items.Item("PlayAll").AccessibleDescription = strGet_ini("ToolBar", "TOOLTIP_PLAY_FIRST", "Play All", strFileName)
            .Items.Item("Play").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_PLAY", "Play From Current Position", strFileName)
            .Items.Item("Play").AccessibleDescription = strGet_ini("ToolBar", "TOOLTIP_PLAY", "Play From Current Position", strFileName)
            .Items.Item("_Stop").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_STOP", "Stop", strFileName)
            .Items.Item("_Stop").AccessibleDescription = strGet_ini("ToolBar", "TOOLTIP_STOP", "Stop", strFileName)

        End With

        With frmWindowFind

            .Text = strGet_ini("Find", "TITLE", "Find/Delete/Replace", strFileName)

            .fraSearchObject.Text = strGet_ini("Find", "FRAME_SEARCH", "Range", strFileName)
            .fraSearchMeasure.Text = strGet_ini("Find", "FRAME_MEASURE", "Range of measure", strFileName)
            .fraSearchNum.Text = strGet_ini("Find", "FRAME_OBJ_NUM", "Range of object number", strFileName)
            .fraSearchGrid.Text = strGet_ini("Find", "FRAME_GRID", "Lane", strFileName)
            .fraProcess.Text = strGet_ini("Find", "FRAME_PROCESS", "Method", strFileName)

            .optSearchAll.Text = strGet_ini("Find", "OPT_OBJ_ALL", "All object", strFileName)
            .optSearchSelect.Text = strGet_ini("Find", "OPT_OBJ_SELECT", "Selected object", strFileName)
            .optProcessSelect.Text = strGet_ini("Find", "OPT_PROCESS_SELECT", "Select", strFileName)
            .optProcessDelete.Text = strGet_ini("Find", "OPT_PROCESS_DELETE", "Delete", strFileName)
            .optProcessReplace.Text = strGet_ini("Find", "OPT_PROCESS_REPLACE", "Replace to", strFileName)

            .cmdInvert.Text = strGet_ini("Find", "CMD_INVERT", "Invert", strFileName)
            .cmdReset.Text = strGet_ini("Find", "CMD_RESET", "Reset", strFileName)
            .cmdSelect.Text = strGet_ini("Find", "CMD_SELECT", "Select All", strFileName)
            .cmdClose.Text = strGet_ini("Find", "CMD_CLOSE", "Close", strFileName)
            .cmdDecide.Text = strGet_ini("Find", "CMD_DECIDE", "Run", strFileName)

            .lblNotice.Text = strGet_ini("Find", "LBL_NOTICE", "This item doesn't influence BPM/STOP object", strFileName)
            .lblMeasure.Text = strGet_ini("Find", "LBL_DASH", "to", strFileName)
            .lblNum.Text = strGet_ini("Find", "LBL_DASH", "to", strFileName)

        End With

        With frmWindowInput

            .Text = strGet_ini("Input", "TITLE", "Input Form", strFileName)

        End With

        With frmWindowViewer

            .Text = strGet_ini("Viewer", "TITLE", "Viewer Setting", strFileName)

            .cmdViewerPath.Text = strGet_ini("Viewer", "CMD_SET", "...", strFileName)
            .cmdAdd.Text = strGet_ini("Viewer", "CMD_ADD", "Add", strFileName)
            .cmdDelete.Text = strGet_ini("Viewer", "CMD_DELETE", "Delete", strFileName)
            .cmdOK.Text = strGet_ini("Viewer", "CMD_OK", "OK", strFileName)
            .cmdCancel.Text = strGet_ini("Viewer", "CMD_CANCEL", "Cancel", strFileName)

            .lblViewerName.Text = strGet_ini("Viewer", "LBL_APP_NAME", "Player name", strFileName)
            .lblViewerPath.Text = strGet_ini("Viewer", "LBL_APP_PATH", "Path", strFileName)
            .lblPlayAll.Text = strGet_ini("Viewer", "LBL_ARG_PLAY_ALL", "Argument of ""Play All""", strFileName)
            .lblPlay.Text = strGet_ini("Viewer", "LBL_ARG_PLAY", "Argument of ""Play""", strFileName)
            .lblStop.Text = strGet_ini("Viewer", "LBL_ARG_STOP", "Argument of ""Stop""", strFileName)
            .lblNotice.Text = Replace(strGet_ini("Viewer", "LBL_ARG_INFO", "Syntax reference:\n<filename> File name\n<measure> Current measure", strFileName), "\n", vbCrLf)

        End With

        With frmWindowConvert

            .Text = strGet_ini("Convert", "TITLE", "Conversion Wizard", strFileName)

            .chkDeleteUnusedFile.Text = strGet_ini("Convert", "CHK_DELETE_LIST", "Clear unused definition from a list", strFileName)

            .chkDeleteFile.Text = strGet_ini("Convert", "CHK_DELETE_FILE", "Delete unused files in this BMS folder (*)", strFileName)
            .lblExtension.Text = strGet_ini("Convert", "LBL_EXTENSION", "Search extensions:", strFileName)
            .chkFileRecycle.Text = strGet_ini("Convert", "CHK_RECYCLE", "Delete soon with no through recycled", strFileName)

            .chkListAlign.Text = strGet_ini("Convert", "CHK_ALIGN_LIST", "Sort definition list", strFileName)
            .chkUseOldFormat.Text = strGet_ini("Convert", "CHK_USE_OLD_FORMAT", "Use old Format [01 - FF] if possible", strFileName)
            .chkSortByName.Text = strGet_ini("Convert", "CHK_SORT_BY_NAME", "Sorting by filename", strFileName)

            .chkFileNameConvert.Text = strGet_ini("Convert", "CHK_CONVERT_FILENAME", "Change filename to list number [01 - ZZ] (*)", strFileName)

            .lblNotice.Text = strGet_ini("Convert", "LBL_NOTICE", "(*) Cannot undo this command", strFileName)

            .cmdDecide.Text = strGet_ini("Convert", "CMD_DECIDE", "Run", strFileName)
            .cmdCancel.Text = strGet_ini("Convert", "CMD_CANCEL", "Cancel", strFileName)

        End With

        g_Message(modMain.Message.ERR_01) = Replace(strGet_ini("Message", "ERROR_MESSAGE_01", "The unexpected error occurred. Program will shut down.\nRefer to the ""error.txt"" for the details of an error.", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.ERR_02) = Replace(strGet_ini("Message", "ERROR_MESSAGE_02", "Temporary file is saved to...", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.ERR_FILE_NOT_FOUND) = Replace(strGet_ini("Message", "ERROR_FILE_NOT_FOUND", "File not found.", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.ERR_LOAD_CANCEL) = Replace(strGet_ini("Message", "ERROR_LOAD_CANCEL", "Loading will be aborted.", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.ERR_SAVE_ERROR) = Replace(strGet_ini("Message", "ERROR_SAVE_ERROR", "Error occured while saving.", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.ERR_SAVE_CANCEL) = Replace(strGet_ini("Message", "ERROR_SAVE_CANCEL", "Saving will be aborted.", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.ERR_OVERFLOW_LARGE) = Replace(strGet_ini("Message", "ERROR_OVERFLOW_LARGE", "Error:\nValue is too large.", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.ERR_OVERFLOW_SMALL) = Replace(strGet_ini("Message", "ERROR_OVERFLOW_SMALL", "Error:\nValue is too small.", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.ERR_OVERFLOW_BPM) = Replace(strGet_ini("Message", "ERROR_OVERFLOW_BPM", "You have used more than kinds of 3843 BPM change command.\nNumber of kinds should be 3843 or less.", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.ERR_OVERFLOW_STOP) = Replace(strGet_ini("Message", "ERROR_OVERFLOW_STOP", "You have used more than kinds of 3843 STOP command.\nNumber of kinds should be 3843 or less.", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.ERR_OVERFLOW_SCROLL) = Replace(strGet_ini("Message", "ERROR_OVERFLOW_SCROLL", "You have used more than 3843 kinds of SCROLL command.\nNumber of kinds should be 3843 or less.", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.ERR_OVERFLOW_SPEED) = Replace(strGet_ini("Message", "ERROR_OVERFLOW_SPEED", "You have used more than 3843 kinds of SPEED command.\nNumber of kinds should be 3843 or less.", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.ERR_APP_NOT_FOUND) = Replace(strGet_ini("Message", "ERROR_APP_NOT_FOUND", " is not found.", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.ERR_FILE_ALREADY_EXIST) = Replace(strGet_ini("Message", "ERROR_FILE_ALREADY_EXIST", "File already exist.", strFileName), "\n", vbCrLf)

        g_Message(modMain.Message.MSG_CONFIRM) = Replace(strGet_ini("Message", "INFO_CONFIRM", "This command cannot be undone, OK?", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.MSG_FILE_CHANGED) = Replace(strGet_ini("Message", "INFO_FILE_CHANGED", "Do you want to save changes?", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.MSG_INI_CHANGED) = Replace(strGet_ini("Message", "INFO_INI_CHANGED", "ini format has changed.\n(All setting will reset)", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.MSG_ALIGN_LIST) = Replace(strGet_ini("Message", "INFO_ALIGN_LIST", "Do you want the filelist to be rewrited into the old format [01 - FF]?\n(Attention: Some programs are compatible only with old format files.)", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.MSG_DELETE_FILE) = Replace(strGet_ini("Message", "INFO_DELETE_FILE", "They have been deleted:", strFileName), "\n", vbCrLf)

        g_Message(modMain.Message.INPUT_BPM) = Replace(strGet_ini("Input", "INPUT_BPM", "Enter the BPM you wish to change to.\n(Decimal number can be used. Enter 0 to cancel)", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.INPUT_STOP) = Replace(strGet_ini("Input", "INPUT_STOP", "Enter the length of stoppage 1 corresponds to 1/192 of the measure.\n(Enter under 0 to cancel)", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.INPUT_SCROLL) = Replace(strGet_ini("Input", "INPUT_SCROLL", "Enter the ratio of scroll.\n1 corresponds to same scroll mount.\n(Decimal, zero, and negative number can be used.)", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.INPUT_SPEED) = Replace(strGet_ini("Input", "INPUT_SPEED", "Enter the ratio of Hi-Speed.\nIt multiplies to the player's Hi-Speed.\n(Decimal, zero, and negative number can be used.)", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.INPUT_RENAME) = Replace(strGet_ini("Input", "INPUT_RENAME", "Please enter new filename.", strFileName), "\n", vbCrLf)
        g_Message(modMain.Message.INPUT_SIZE) = Replace(strGet_ini("Input", "INPUT_SIZE", "Type your display magnification.\n(Maximum 16.00. Enter under 0 to cancel)", strFileName), "\n", vbCrLf)

        Dim DefaultFont As String
        Dim SystemFont As String
        Dim FixedFont As String

        'Dim ncm As NONCLIENTMETRICS
        'ncm.cbSize = LenB(ncm)
        'Call SystemParametersInfo(SPI_GETNONCLIENTMETRICS, LenB(ncm), ncm, 0)
        'DefaultFont = StrConv(ncm.lfSMCaptionFont.lfFaceName, vbUnicode)

        Dim lf As LOGFONT = New LOGFONT
        Call GetObject_Renamed(GetStockObject(DEFAULT_GUI_FONT), Runtime.InteropServices.Marshal.SizeOf(lf), lf)
        DefaultFont = Trim(lf.lfFaceName)

        SystemFont = strGet_ini("Main", "Font", DefaultFont, strFileName)
        FixedFont = strGet_ini("Main", "FixedFont", DefaultFont, strFileName)

        'フォント強制変更
        SystemFont = strGet_ini("Main", "Font", SystemFont, "bmse.ini")
        FixedFont = strGet_ini("Main", "FixedFont", FixedFont, "bmse.ini")

        Call LoadFont(SystemFont, FixedFont, strGet_ini("Main", "Charset", DEFAULT_CHARSET, strFileName))

        Call frmMain.frmMain_Resize(Nothing, New System.EventArgs())

    End Sub

    Private Sub LoadFont(ByRef MainFont As String, ByRef FixedFont As String, ByVal Charset As Integer)
        On Error GoTo Err_Renamed

        Dim i As Integer
        Dim objCtl As Object

        For i = 0 To Application.OpenForms.Count - 1

            Application.OpenForms.Item(i).Font = New Font(MainFont, Application.OpenForms.Item(i).Font.Size, Application.OpenForms.Item(i).Font.Style, Application.OpenForms.Item(i).Font.Unit, Charset, Application.OpenForms.Item(i).Font.GdiVerticalFont)

            For Each objCtl In Application.OpenForms.Item(i).Controls

                If TypeOf objCtl Is System.Windows.Forms.Label Or TypeOf objCtl Is System.Windows.Forms.TextBox Or TypeOf objCtl Is System.Windows.Forms.ComboBox Or TypeOf objCtl Is System.Windows.Forms.Button Or TypeOf objCtl Is System.Windows.Forms.RadioButton Or TypeOf objCtl Is System.Windows.Forms.CheckBox Or TypeOf objCtl Is System.Windows.Forms.GroupBox Or TypeOf objCtl Is System.Windows.Forms.Panel Then
                    objCtl.Font = New Font(MainFont, objCtl.Font.Size, objCtl.Font.Style, objCtl.Font.Unit, Charset)
                ElseIf TypeOf objCtl Is System.Windows.Forms.PictureBox Or TypeOf objCtl Is System.Windows.Forms.ListBox Then
                    objCtl.Font = New Font(FixedFont, objCtl.Font.Size, objCtl.Font.Style, objCtl.Font.Unit, Charset)
                End If

            Next objCtl

        Next i

        Exit Sub

Err_Renamed:
        Call modMain.CleanUp(Err.Number, Err.Description, "LoadFont")
    End Sub

    Public Sub LoadConfig()
        On Error GoTo InitConfig

        Dim i As Integer
        Dim wp As WINDOWPLACEMENT
        Dim strTemp As String
        Dim lngTemp As Integer
        Static lngCount As Integer

        If strGet_ini("Main", "Key", "", "bmse.ini") <> "BMSE" Then GoTo InitConfig

        With frmMain

            strTemp = strGet_ini("Main", "Language", "english.ini", "bmse.ini")

            If strTemp = g_strLangFileName(0) Then
                ._mnuLanguage_0.Checked = True
            End If
            If strTemp = g_strLangFileName(1) Then
                ._mnuLanguage_1.Checked = True
            End If
            If strTemp = g_strLangFileName(2) Then
                ._mnuLanguage_2.Checked = True
            End If

            'frmWindowAbout.Show()
            'frmWindowFind.Show()
            'frmWindowInput.Show()
            'frmWindowPreview.Show()
            'frmWindowTips.Show()
            'frmWindowViewer.Show()
            'frmWindowConvert.Show()

            Call LoadLanguageFile("lang\" & strTemp)

            Call frmWindowPreview.SetWindowSize()

            If strGet_ini("Main", "ini", "", "bmse.ini") <> modVersion.INI_VERSION Then

                Call MsgBox(g_Message(Message.MSG_INI_CHANGED), vbInformation, g_strAppTitle)

                GoTo InitConfig

            End If

            wp.Length = 44
            Call GetWindowPlacement(.Handle, wp)

            With wp

                .showCmd = SW_HIDE

                With .rcNormalPosition

                    .right_Renamed = strGet_ini("Main", "Width", 1280, "bmse.ini")
                    .Bottom = strGet_ini("Main", "Height", 720, "bmse.ini")
                    '.Left = strGet_ini("Main", "X", (Screen.Width \ Screen.TwipsPerPixelX - .Right) \ 2, "bmse.ini")
                    '.Top = strGet_ini("Main", "Y", (Screen.Height \ Screen.TwipsPerPixelY - .Bottom) \ 2, "bmse.ini")
                    .left_Renamed = strGet_ini("Main", "X", 0, "bmse.ini")
                    .Top = strGet_ini("Main", "Y", 0, "bmse.ini")
                    .right_Renamed = .left_Renamed + .right_Renamed
                    .Bottom = .Top + .Bottom

                    If .right_Renamed > System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width Then

                        .left_Renamed = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - .right_Renamed - .left_Renamed
                        .right_Renamed = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width

                    End If

                    If .left_Renamed < 0 Then .left_Renamed = 0

                    If .Bottom > System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height Then

                        .Top = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height - (.Bottom - .Top)
                        .Bottom = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height

                    End If

                    If .Top < 0 Then .Top = 0

                End With

            End With

            strTemp = strGet_ini("Main", "Theme", "default.ini", "bmse.ini")

            If strTemp = g_strThemeFileName(0) Then
                ._mnuTheme_0.Checked = True
            End If
            If strTemp = g_strThemeFileName(1) Then
                ._mnuTheme_1.Checked = True
            End If
            If strTemp = g_strThemeFileName(2) Then
                ._mnuTheme_2.Checked = True
            End If

            Call LoadThemeFile("theme\" & strTemp)

            g_strHelpFilename = strGet_ini("Main", "Help", "", "bmse.ini")
            g_strFiler = strGet_ini("Main", "Filer", "", "bmse.ini")

            If g_strHelpFilename <> "" Then

                .mnuHelpOpen.Enabled = True

            End If

            '.hsbDispWidth.Value = strGet_ini("View", "Width", 100, "bmse.ini")
            '.hsbDispHeight.Value = strGet_ini("View", "Height", 100, "bmse.ini")

            With .cboDispWidth

                lngTemp = strGet_ini("View", "Width", 100, "bmse.ini")

                For i = 0 To .Items.Count - 1

                    If DirectCast(frmMain.cboDispWidth.Items.Item(i), modMain.ItemWithData).ItemData = lngTemp Then

                        .SelectedIndex = i

                        Exit For

                    ElseIf DirectCast(frmMain.cboDispWidth.Items.Item(i), modMain.ItemWithData).ItemData > lngTemp Then

                        Call .Items.Insert(i, New modMain.ItemWithData("x" & Format(lngTemp / 100, "#0.00"), lngTemp))
                        .SelectedIndex = i

                        Exit For

                    End If

                Next i

            End With

            With .cboDispHeight

                lngTemp = strGet_ini("View", "Height", 50, "bmse.ini")

                For i = 0 To .Items.Count - 1

                    If DirectCast(frmMain.cboDispHeight.Items.Item(i), modMain.ItemWithData).ItemData = lngTemp Then

                        .SelectedIndex = i

                        Exit For

                    ElseIf DirectCast(frmMain.cboDispHeight.Items.Item(i), modMain.ItemWithData).ItemData > lngTemp Then

                        Call .Items.Insert(i, New modMain.ItemWithData("x" & Format(lngTemp / 100, "#0.00"), lngTemp))
                        .SelectedIndex = i

                        Exit For

                    End If

                Next i

            End With

            .cboDispGridMain.SelectedIndex = strGet_ini("View", "VGridMain", 1, "bmse.ini")
            .cboDispGridSub.SelectedIndex = strGet_ini("View", "VGridSub", 2, "bmse.ini")
            .cboDispFrame.SelectedIndex = strGet_ini("View", "Frame", 1, "bmse.ini")
            .cboVScroll.SelectedIndex = strGet_ini("View", "VScroll", 4, "bmse.ini")
            .cboDispKey.SelectedIndex = strGet_ini("View", "Key", 1, "bmse.ini")
            .cboDispSC1P.SelectedIndex = strGet_ini("View", "SC_1P", 0, "bmse.ini")
            .cboDispSC2P.SelectedIndex = strGet_ini("View", "SC_2P", 1, "bmse.ini")

            ._mnuViewItem_0.Checked = strGet_ini("View", "ToolBar", True, "bmse.ini")
            ._mnuViewItem_1.Checked = strGet_ini("View", "DirectInput", True, "bmse.ini")
            ._mnuViewItem_2.Checked = strGet_ini("View", "StatusBar", True, "bmse.ini")

            ._mnuViewItem_0_New.Checked = strGet_ini("ToolBar", "New", True, "bmse.ini")
            ._mnuViewItem_0_Open.Checked = strGet_ini("ToolBar", "Open", True, "bmse.ini")
            ._mnuViewItem_0_Reload.Checked = strGet_ini("ToolBar", "Reload", True, "bmse.ini")
            ._mnuViewItem_0_Save.Checked = strGet_ini("ToolBar", "Save", True, "bmse.ini")
            ._mnuViewItem_0_SaveAs.Checked = strGet_ini("ToolBar", "SaveAs", True, "bmse.ini")
            ._mnuViewItem_0_Mode.Checked = strGet_ini("ToolBar", "Mode", True, "bmse.ini")
            ._mnuViewItem_0_Preview.Checked = strGet_ini("ToolBar", "Preview", True, "bmse.ini")
            ._mnuViewItem_0_Grid.Checked = strGet_ini("ToolBar", "Grid", True, "bmse.ini")
            ._mnuViewItem_0_Size.Checked = strGet_ini("ToolBar", "Size", True, "bmse.ini")
            ._mnuViewItem_0_Resolution.Checked = strGet_ini("ToolBar", "Resolution", True, "bmse.ini")

            If .cboViewer.Items.Count Then

                If .cboViewer.Items.Count > strGet_ini("View", "ViewerNum", 0, "bmse.ini") Then

                    .cboViewer.SelectedIndex = strGet_ini("View", "ViewerNum", 0, "bmse.ini")

                Else

                    .cboViewer.SelectedIndex = 0

                End If

            End If

            ._mnuOptionsItem_0.Checked = strGet_ini("Options", "Active", True, "bmse.ini")
            ._mnuOptionsItem_1.Checked = strGet_ini("Options", "FileNameOnly", False, "bmse.ini")
            ._mnuOptionsItem_2.Checked = strGet_ini("Options", "VerticalWriting", True, "bmse.ini")
            ._mnuOptionsItem_3.Checked = strGet_ini("Options", "LaneBG", True, "bmse.ini")
            ._mnuOptionsItem_4.Checked = strGet_ini("Options", "SelectSound", True, "bmse.ini")
            ._mnuOptionsItem_5.Checked = strGet_ini("Options", "MoveOnGrid", True, "bmse.ini")
            ._mnuOptionsItem_6.Checked = strGet_ini("Options", "ObjectFileName", False, "bmse.ini")
            ._mnuOptionsItem_7.Checked = strGet_ini("Options", "UseNewFormat", False, "bmse.ini")
            ._mnuOptionsItem_8.Checked = strGet_ini("Options", "YAxisFixed", False, "bmse.ini")
            ._mnuOptionsItem_9.Checked = strGet_ini("Options", "EnableTooltip", False, "bmse.ini")
            '.mnuOptionsItem(RCLICK_DELETE).Checked = strGet_ini("Options", "RightClickDelete", False, "bmse.ini")

            strTemp = strGet_ini("Options", "BaseNumber", "36", "bmse.ini")
            If strTemp = "16" Then
                ._mnuOptionsBase16.Checked = True
                ._mnuOptionsBase36.Checked = False
                ._mnuOptionsBase62.Checked = False
            ElseIf strTemp = "62" Then
                ._mnuOptionsBase16.Checked = False
                ._mnuOptionsBase36.Checked = False
                ._mnuOptionsBase62.Checked = True
            Else
                ._mnuOptionsBase16.Checked = False
                ._mnuOptionsBase36.Checked = True
                ._mnuOptionsBase62.Checked = False
            End If

            .tlbMenu.Items.Item("_New").Visible = strGet_ini("ToolBar", "New", True, "bmse.ini")
            .tlbMenu.Items.Item("Open").Visible = strGet_ini("ToolBar", "Open", True, "bmse.ini")
            .tlbMenu.Items.Item("Reload").Visible = strGet_ini("ToolBar", "Reload", False, "bmse.ini")
            .tlbMenu.Items.Item("Save").Visible = strGet_ini("ToolBar", "Save", True, "bmse.ini")
            .tlbMenu.Items.Item("SaveAs").Visible = strGet_ini("ToolBar", "SaveAs", True, "bmse.ini")

            .tlbMenu.Items.Item("SepMode").Visible = strGet_ini("ToolBar", "Mode", True, "bmse.ini")
            .tlbMenu.Items.Item("Edit").Visible = strGet_ini("ToolBar", "Mode", True, "bmse.ini")
            .tlbMenu.Items.Item("Write").Visible = strGet_ini("ToolBar", "Mode", True, "bmse.ini")
            .tlbMenu.Items.Item("Delete").Visible = strGet_ini("ToolBar", "Mode", True, "bmse.ini")

            .tlbMenu.Items.Item("SepViewer").Visible = strGet_ini("ToolBar", "Preview", True, "bmse.ini")
            .tlbMenu.Items.Item("PlayAll").Visible = strGet_ini("ToolBar", "Preview", True, "bmse.ini")
            .tlbMenu.Items.Item("Play").Visible = strGet_ini("ToolBar", "Preview", True, "bmse.ini")
            .tlbMenu.Items.Item("_Stop").Visible = strGet_ini("ToolBar", "Preview", True, "bmse.ini")

            .tlbMenu.Items.Item("SepGrid").Visible = strGet_ini("ToolBar", "Grid", True, "bmse.ini")
            .tlbMenu.Items.Item("ChangeGrid").Visible = strGet_ini("ToolBar", "Grid", True, "bmse.ini")

            .tlbMenu.Items.Item("SepSize").Visible = strGet_ini("ToolBar", "Size", True, "bmse.ini")
            .tlbMenu.Items.Item("DispSize").Visible = strGet_ini("ToolBar", "Size", True, "bmse.ini")

            .tlbMenu.Items.Item("SepResolution").Visible = strGet_ini("ToolBar", "Resolution", False, "bmse.ini")
            .tlbMenu.Items.Item("Resolution").Visible = strGet_ini("ToolBar", "Resolution", False, "bmse.ini")

            For i = 0 To UBound(g_strRecentFiles)

                g_strRecentFiles(i) = strGet_ini("RecentFiles", i, "", "bmse.ini")

                If Len(g_strRecentFiles(i)) Then
                    Select Case i
                        Case 0
                            With ._mnuRecentFiles_0

                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True

                            End With

                            With .ToolStripMenuItem0
                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True
                            End With

                        Case 1
                            With ._mnuRecentFiles_1

                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True

                            End With

                            With .ToolStripMenuItem1
                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True
                            End With

                        Case 2
                            With ._mnuRecentFiles_2

                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True

                            End With

                            With .ToolStripMenuItem2
                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True
                            End With

                        Case 3
                            With ._mnuRecentFiles_3

                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True

                            End With

                            With .ToolStripMenuItem3
                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True
                            End With

                        Case 4
                            With ._mnuRecentFiles_4

                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True

                            End With

                            With .ToolStripMenuItem4
                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True
                            End With

                        Case 5
                            With ._mnuRecentFiles_5

                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True

                            End With

                            With .ToolStripMenuItem5
                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True
                            End With

                        Case 6
                            With ._mnuRecentFiles_6

                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True

                            End With

                            With .ToolStripMenuItem6
                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True
                            End With

                        Case 7
                            With ._mnuRecentFiles_7

                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True

                            End With

                            With .ToolStripMenuItem7
                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True
                            End With

                        Case 8
                            With ._mnuRecentFiles_8

                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True

                            End With

                            With .ToolStripMenuItem8
                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True
                            End With

                        Case 9
                            With ._mnuRecentFiles_9

                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True

                            End With

                            With .ToolStripMenuItem9
                                .Text = "&" & Right(CStr(i + 1), 1) & ":" & g_strRecentFiles(i)
                                .Enabled = True
                                .Visible = True
                            End With
                    End Select

                    .mnuLineRecent.Visible = True

                End If

            Next i

            Call SetWindowPlacement(.Handle, wp)

        End With

        Call modEasterEgg.InitEffect()

        With frmWindowPreview

            .Left = strGet_ini("Preview", "X", (frmMain.Left + frmMain.Width \ 2) - .Width \ 2, "bmse.ini")
            If .Left < 0 Then .Left = 0
            If .Left > System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width Then .Left = 0

            .Top = strGet_ini("Preview", "Y", (frmMain.Top + frmMain.Height \ 2) - .Height \ 2, "bmse.ini")
            If .Top < 0 Then .Top = 0
            If .Top > System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height Then .Top = 0

        End With

        Exit Sub

InitConfig:

        lngCount = lngCount + 1

        If lngCount > 5 Then

            Call modMain.CleanUp(Err.Number, Err.Description, "LoadConfig")

        Else

            Call CreateConfig()

        End If

    End Sub

    Private Sub CreateConfig()
        Call lngSet_ini("Main", "Key", Chr(34) & "BMSE" & Chr(34))
        Call lngSet_ini("Main", "ini", modVersion.INI_VERSION)
        'Call lngSet_ini("Main", "X", (Screen.Width \ Screen.TwipsPerPixelX - 800) \ 2)
        'Call lngSet_ini("Main", "Y", (Screen.Height \ Screen.TwipsPerPixelY - 600) \ 2)
        Call lngSet_ini("Main", "X", 0)
        Call lngSet_ini("Main", "Y", 0)
        Call lngSet_ini("Main", "Width", "1280")
        Call lngSet_ini("Main", "Height", "720")
        Call lngSet_ini("Main", "State", SW_SHOWNORMAL)
        Call lngSet_ini("Main", "Language", Chr(34) & "english.ini" & Chr(34))
        Call lngSet_ini("Main", "Theme", Chr(34) & "default.ini" & Chr(34))
        Call lngSet_ini("Main", "Help", Chr(34) & Chr(34))

        Call lngSet_ini("View", "Width", 100)
        Call lngSet_ini("View", "Height", 50)
        Call lngSet_ini("View", "VGridMain", 1)
        Call lngSet_ini("View", "VGridSub", 2)
        Call lngSet_ini("View", "VScroll", 4)
        Call lngSet_ini("View", "Frame", 1)
        Call lngSet_ini("View", "Key", 1)
        Call lngSet_ini("View", "SC_1P", 0)
        Call lngSet_ini("View", "SC_2P", 1)

        Call lngSet_ini("View", "ToolBar", True)
        Call lngSet_ini("View", "DirectInput", True)
        Call lngSet_ini("View", "StatusBar", True)

        Call lngSet_ini("View", "ViewerNum", 0)

        Call lngSet_ini("ToolBar", "New", True)
        Call lngSet_ini("ToolBar", "Open", True)
        Call lngSet_ini("ToolBar", "Reload", False)
        Call lngSet_ini("ToolBar", "Save", True)
        Call lngSet_ini("ToolBar", "SaveAs", True)
        Call lngSet_ini("ToolBar", "Mode", True)
        Call lngSet_ini("ToolBar", "Preview", True)
        Call lngSet_ini("ToolBar", "Grid", True)
        Call lngSet_ini("ToolBar", "Size", True)
        Call lngSet_ini("ToolBar", "Resolution", False)

        Call lngSet_ini("Options", "Active", True)
        Call lngSet_ini("Options", "FileNameOnly", False)
        Call lngSet_ini("Options", "VerticalWriting", True)
        Call lngSet_ini("Options", "LaneBG", True)
        Call lngSet_ini("Options", "SelectSound", True)
        Call lngSet_ini("Options", "MoveOnGrid", True)
        Call lngSet_ini("Options", "ObjectFileName", False)
        Call lngSet_ini("Options", "UseNewFormat", False)
        Call lngSet_ini("Options", "RightClickDelete", False)
        Call lngSet_ini("Options", "YAxisFixed", False)
        Call lngSet_ini("Options", "BaseNumber", "36")

        Call lngSet_ini("Preview", "X", 0)
        Call lngSet_ini("Preview", "Y", 0)

        Call LoadConfig()

    End Sub

    Public Sub SaveConfig()
        Dim i As Integer
        Dim wp As WINDOWPLACEMENT

        Call lngSet_ini("Main", "Key", Chr(34) & "BMSE" & Chr(34))

        wp.Length = 44
        Call GetWindowPlacement(frmMain.Handle, wp)

        With wp

            If wp.showCmd <> SW_SHOWMINIMIZED Then

                Call lngSet_ini("Main", "State", wp.showCmd)

            Else

                Call lngSet_ini("Main", "State", SW_SHOWNORMAL)

            End If

            With .rcNormalPosition

                Call lngSet_ini("Main", "X", .left_Renamed)
                Call lngSet_ini("Main", "Y", .Top)
                Call lngSet_ini("Main", "Width", .right_Renamed - .left_Renamed)
                Call lngSet_ini("Main", "Height", .Bottom - .Top)

            End With

        End With

        With frmMain

            If ._mnuLanguage_0.Checked = True Then
                Call lngSet_ini("Main", "Language", Chr(34) & g_strLangFileName(0) & Chr(34))
            End If
            If ._mnuLanguage_1.Checked = True Then
                Call lngSet_ini("Main", "Language", Chr(34) & g_strLangFileName(1) & Chr(34))
            End If
            If ._mnuLanguage_2.Checked = True Then
                Call lngSet_ini("Main", "Language", Chr(34) & g_strLangFileName(2) & Chr(34))
            End If

            If ._mnuTheme_0.Checked = True Then
                Call lngSet_ini("Main", "Theme", Chr(34) & g_strThemeFileName(0) & Chr(34))
            End If
            If ._mnuTheme_1.Checked = True Then
                Call lngSet_ini("Main", "Theme", Chr(34) & g_strThemeFileName(1) & Chr(34))
            End If
            If ._mnuTheme_2.Checked = True Then
                Call lngSet_ini("Main", "Theme", Chr(34) & g_strThemeFileName(2) & Chr(34))
            End If

            Call lngSet_ini("View", "Width", DirectCast(.cboDispWidth.SelectedItem, modMain.ItemWithData).ItemData)
            Call lngSet_ini("View", "Height", DirectCast(.cboDispHeight.SelectedItem, modMain.ItemWithData).ItemData)
            Call lngSet_ini("View", "VGridMain", .cboDispGridMain.SelectedIndex)
            Call lngSet_ini("View", "VGridSub", .cboDispGridSub.SelectedIndex)
            Call lngSet_ini("View", "VScroll", .cboVScroll.SelectedIndex)
            Call lngSet_ini("View", "Frame", .cboDispFrame.SelectedIndex)
            Call lngSet_ini("View", "Key", .cboDispKey.SelectedIndex)
            Call lngSet_ini("View", "SC_1P", .cboDispSC1P.SelectedIndex)
            Call lngSet_ini("View", "SC_2P", .cboDispSC2P.SelectedIndex)

            Call lngSet_ini("View", "ToolBar", ._mnuViewItem_0.Checked)
            Call lngSet_ini("View", "DirectInput", ._mnuViewItem_1.Checked)
            Call lngSet_ini("View", "StatusBar", ._mnuViewItem_2.Checked)

            If .cboViewer.Items.Count Then

                Call lngSet_ini("View", "ViewerNum", .cboViewer.SelectedIndex)

            End If

            Call lngSet_ini("Options", "Active", ._mnuOptionsItem_0.Checked)
            Call lngSet_ini("Options", "FileNameOnly", ._mnuOptionsItem_1.Checked)
            Call lngSet_ini("Options", "VerticalWriting", ._mnuOptionsItem_2.Checked)
            Call lngSet_ini("Options", "LaneBG", ._mnuOptionsItem_3.Checked)
            Call lngSet_ini("Options", "SelectSound", ._mnuOptionsItem_4.Checked)
            Call lngSet_ini("Options", "MoveOnGrid", ._mnuOptionsItem_5.Checked)
            Call lngSet_ini("Options", "ObjectFileName", ._mnuOptionsItem_6.Checked)
            Call lngSet_ini("Options", "UseNewFormat", ._mnuOptionsItem_7.Checked)
            Call lngSet_ini("Options", "YAxisFixed", ._mnuOptionsItem_8.Checked)
            Call lngSet_ini("Options", "EnableTooltip", ._mnuOptionsItem_9.Checked)
            'Call lngSet_ini("Options", "RightClickDelete", .mnuOptionsItem(RCLICK_DELETE).Checked)
            If ._mnuOptionsBase16.Checked Then
                Call lngSet_ini("Options", "BaseNumber", "16")
            ElseIf ._mnuOptionsBase36.Checked Then
                Call lngSet_ini("Options", "BaseNumber", "36")
            ElseIf ._mnuOptionsBase62.Checked Then
                Call lngSet_ini("Options", "BaseNumber", "62")
            End If

            Call lngSet_ini("ToolBar", "New", ._mnuViewItem_0_New.Checked)
            Call lngSet_ini("ToolBar", "Open", ._mnuViewItem_0_Open.Checked)
            Call lngSet_ini("ToolBar", "Reload", ._mnuViewItem_0_Reload.Checked)
            Call lngSet_ini("ToolBar", "Save", ._mnuViewItem_0_Save.Checked)
            Call lngSet_ini("ToolBar", "SaveAs", ._mnuViewItem_0_SaveAs.Checked)
            Call lngSet_ini("ToolBar", "Mode", ._mnuViewItem_0_Mode.Checked)
            Call lngSet_ini("ToolBar", "Preview", ._mnuViewItem_0_Preview.Checked)
            Call lngSet_ini("ToolBar", "Grid", ._mnuViewItem_0_Grid.Checked)
            Call lngSet_ini("ToolBar", "Size", ._mnuViewItem_0_Size.Checked)
            Call lngSet_ini("ToolBar", "Resolution", ._mnuViewItem_0_Resolution.Checked)

            For i = 0 To UBound(g_strRecentFiles)

                Call lngSet_ini("RecentFiles", i, Chr(34) & g_strRecentFiles(i) & Chr(34))

            Next i

        End With

        With frmWindowPreview

            Call lngSet_ini("Preview", "X", .Left)
            Call lngSet_ini("Preview", "Y", .Top)

        End With

    End Sub

    Public Function lngSet_ini(ByRef strSection As String, ByVal strKey As String, ByVal strSet As String) As Integer
        Dim lngTemp As Integer

        'API呼び出し＆変数を返す
        lngTemp = WritePrivateProfileString(strSection & Chr(0), strKey, strSet, g_strAppDir & "bmse.ini" & Chr(0))

        lngSet_ini = lngTemp

    End Function

    Public Function strGet_ini(ByRef strSection As String, ByVal strKey As String, ByVal strDefault As String, ByRef strFileName As String) As String
        'バッファの初期化（256もあれば良いよね。）<- ダメ！ (nksv-1.5.1)
        Dim strGetBuf As StringBuilder = New StringBuilder(modInput.BGM_LANE * 5) '収容するstringのバッファ, "B"+数字3桁+"," + ... +"B"+数字3桁+Chr(0) == BGM_LANE * 5
        Dim LeftLength As Integer

        'API呼び出し
        GetPrivateProfileString(strSection & Chr(0), strKey, strDefault & Chr(0), strGetBuf, strGetBuf.Capacity, g_strAppDir & strFileName & Chr(0))

        '文字列を返す
        LeftLength = InStr(strGetBuf.ToString(), Chr(0)) - 1
        If LeftLength > 0 Then
            strGet_ini = Trim(Left(strGetBuf.ToString(), LeftLength))
        Else
            strGet_ini = Trim(strGetBuf.ToString())
        End If

        If Val(strGet_ini) < 0 Then strGet_ini = CStr(0)

    End Function

    Private Function GetColor(ByRef strSection As String, ByRef strKey As String, ByRef strDefault As String, ByRef strFileName As String) As Integer
        Dim strArray() As String

        strArray = Split(strGet_ini(strSection, strKey, strDefault, strFileName), ",")

        If UBound(strArray) < 2 Then
            GetColor = RGB(0, 0, 0)
            Exit Function
        End If

        GetColor = RGB(CInt(strArray(0)), CInt(strArray(1)), CInt(strArray(2)))

    End Function

    Private Function HalfColor(ByVal Color As Integer) As Integer

        Dim r As Byte
        Dim g As Byte
        Dim b As Byte

        r = Color And &HFF
        g = (Color \ &HFF) And &HFF
        b = (Color \ &HFFFF) And &HFF

        HalfColor = RGB(r \ 2, g \ 2, b \ 2)

    End Function
End Module
