Option Strict Off
Option Explicit On
Public Module modBMS

    'BMSのchannelって実は36進数
    Public Enum OBJ_CH
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
        CH_INV = 2 * 36 + 0
        CH_LN = 4 * 36 + 0
        CH_MINE = 12 * 36 + 0
        CH_KEY_MIN = 1 * 36 + 0
        CH_1P_KEY1 = OBJ_CH.CH_KEY_MIN + 1
        CH_1P_KEY2 = OBJ_CH.CH_KEY_MIN + 2
        CH_1P_KEY3 = OBJ_CH.CH_KEY_MIN + 3
        CH_1P_KEY4 = OBJ_CH.CH_KEY_MIN + 4
        CH_1P_KEY5 = OBJ_CH.CH_KEY_MIN + 5
        CH_1P_KEY6 = OBJ_CH.CH_KEY_MIN + 8
        CH_1P_KEY7 = OBJ_CH.CH_KEY_MIN + 9
        CH_1P_SC = OBJ_CH.CH_KEY_MIN + 6
        CH_2P_KEY1 = OBJ_CH.CH_1P_KEY1 + 1 * 36 + 0
        CH_2P_KEY2 = OBJ_CH.CH_1P_KEY2 + 1 * 36 + 0
        CH_2P_KEY3 = OBJ_CH.CH_1P_KEY3 + 1 * 36 + 0
        CH_2P_KEY4 = OBJ_CH.CH_1P_KEY4 + 1 * 36 + 0
        CH_2P_KEY5 = OBJ_CH.CH_1P_KEY5 + 1 * 36 + 0
        CH_2P_KEY6 = OBJ_CH.CH_1P_KEY6 + 1 * 36 + 0
        CH_2P_KEY7 = OBJ_CH.CH_1P_KEY7 + 1 * 36 + 0
        CH_2P_SC = OBJ_CH.CH_1P_SC + 1 * 36 + 0
        CH_KEY_MAX = OBJ_CH.CH_KEY_MIN + 2 * 36 + 0
        CH_KEY_INV_MIN = OBJ_CH.CH_KEY_MIN + OBJ_CH.CH_INV
        CH_KEY_INV_MAX = OBJ_CH.CH_KEY_MAX + OBJ_CH.CH_INV
        CH_KEY_LN_MIN = OBJ_CH.CH_KEY_MIN + OBJ_CH.CH_LN
        CH_KEY_LN_MAX = OBJ_CH.CH_KEY_MAX + OBJ_CH.CH_LN
        CH_KEY_MINE_MIN = OBJ_CH.CH_KEY_MIN + OBJ_CH.CH_MINE
        CH_KEY_MINE_MAX = OBJ_CH.CH_KEY_MAX + OBJ_CH.CH_MINE
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

    Public Enum BMS_CONSTANT '絶対変えない
        MATERIAL_MAX = 3843
        MEASURE_MAX = 999
        MEASURE_LENGTH = 192
    End Enum

    Public Structure m_udtMeasure
        Dim intLen As Integer '4拍=BMS_CONSTANT.MEASURE_LENGTH
        Dim lngY As Integer
    End Structure

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

End Module
