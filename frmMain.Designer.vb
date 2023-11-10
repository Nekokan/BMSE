<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmMain
#Region "Windows フォーム デザイナによって生成されたコード "
    <System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
        MyBase.New()
        'この呼び出しは、Windows フォーム デザイナで必要です。
        InitializeComponent()
        stringFont = New Font(Font.FontFamily, Font.Size, Font.Style, Font.Unit, Font.GdiCharSet, Font.GdiVerticalFont)
    End Sub
    'Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        stringFont.Dispose()
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    'Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents mnuFileNew As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFileOpen As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFileSave As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFileSaveAs As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFileOpenDirectory As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLineFile As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuRecentFiles_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuRecentFiles_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuRecentFiles_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuRecentFiles_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuRecentFiles_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuRecentFiles_5 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuRecentFiles_6 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuRecentFiles_7 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuRecentFiles_8 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuRecentFiles_9 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLineRecent As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuFileConvertWizard As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLineExit As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuFileExit As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuEditUndo As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuEditRedo As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLineEdit1 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuEditCut As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuEditCopy As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuEditPaste As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuEditDelete As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLineEdit2 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuEditSelectAll As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLineEdit3 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuEditFind As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLineEdit4 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuEditMode_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuEditMode_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuEditMode_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuEdit As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuViewItem_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuViewItem_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuViewItem_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuView As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuOptionsItem_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuOptionsItem_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuOptionsItem_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuOptionsItem_3 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuOptionsItem_4 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuOptionsItem_5 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuOptionsItem_6 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuOptionsItem_7 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuOptionsItem_8 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuOptionsItem_9 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLineOptions As System.Windows.Forms.ToolStripSeparator
    Public WithEvents _mnuLanguage_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuLanguage_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuLanguage_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLanguageParent As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuTheme_0 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuTheme_1 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents _mnuTheme_2 As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuThemeParent As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuOptions As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuToolsPlayAll As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuToolsPlay As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuToolsPlayStop As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLineTools As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuToolsSetting As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuTools As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelpOpen As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuLineHelp As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuHelpWeb As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelpAbout As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuContextPlayAll As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuContextPlay As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuContextBar1 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuContextInsertMeasure As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuContextDeleteMeasure As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuContextBar2 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents mnuContextEditCut As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuContextEditCopy As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuContextEditPaste As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuContextEditDelete As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuContext As System.Windows.Forms.ContextMenuStrip
    Public WithEvents mnuContextListLoad As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuContextListDelete As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuContextListRename As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents mnuContextList As System.Windows.Forms.ContextMenuStrip
    Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
    Public WithEvents Mode As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents Position As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents WAV As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents BMP As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents Measure As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents _staMain_Panel6 As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents staMain As System.Windows.Forms.StatusStrip
    Public WithEvents picSiromaru As System.Windows.Forms.PictureBox
    Public WithEvents tmrEffect As System.Windows.Forms.Timer
    Public WithEvents cboVScroll As System.Windows.Forms.ComboBox
    Public WithEvents lblVScroll As System.Windows.Forms.Label
    Public WithEvents fraResolution As System.Windows.Forms.Panel
    Public WithEvents cboDispHeight As System.Windows.Forms.ComboBox
    Public WithEvents cboDispWidth As System.Windows.Forms.ComboBox
    Public WithEvents lblDispWidth As System.Windows.Forms.Label
    Public WithEvents lblDispHeight As System.Windows.Forms.Label
    Public WithEvents fraDispSize As System.Windows.Forms.Panel
    Public WithEvents cboViewer As System.Windows.Forms.ComboBox
    Public WithEvents fraViewer As System.Windows.Forms.Panel
    Public WithEvents tmrMain As System.Windows.Forms.Timer
    Public WithEvents ilsMenu As System.Windows.Forms.ImageList
    Public WithEvents cboDirectInput As System.Windows.Forms.ComboBox
    Public WithEvents cmdDirectInput As System.Windows.Forms.Button
    Public WithEvents cboDispGridMain As System.Windows.Forms.ComboBox
    Public WithEvents cboDispGridSub As System.Windows.Forms.ComboBox
    Public WithEvents lblGridSub As System.Windows.Forms.Label
    Public WithEvents lblGridMain As System.Windows.Forms.Label
    Public WithEvents fraGrid As System.Windows.Forms.Panel
    Public WithEvents cboPlayer As System.Windows.Forms.ComboBox
    Public WithEvents txtGenre As System.Windows.Forms.TextBox
    Public WithEvents txtTitle As System.Windows.Forms.TextBox
    Public WithEvents txtArtist As System.Windows.Forms.TextBox
    Public WithEvents cboPlayLevel As System.Windows.Forms.ComboBox
    Public WithEvents txtBPM As System.Windows.Forms.TextBox
    Public WithEvents lblPlayMode As System.Windows.Forms.Label
    Public WithEvents lblGenre As System.Windows.Forms.Label
    Public WithEvents lblTitle As System.Windows.Forms.Label
    Public WithEvents lblArtist As System.Windows.Forms.Label
    Public WithEvents lblPlayLevel As System.Windows.Forms.Label
    Public WithEvents lblBPM As System.Windows.Forms.Label
    Public WithEvents _fraTop_0 As System.Windows.Forms.Panel
    Public WithEvents cboPlayRank As System.Windows.Forms.ComboBox
    Public WithEvents txtTotal As System.Windows.Forms.TextBox
    Public WithEvents txtVolume As System.Windows.Forms.TextBox
    Public WithEvents txtStageFile As System.Windows.Forms.TextBox
    Public WithEvents cmdLoadStageFile As System.Windows.Forms.Button
    Public WithEvents cmdLoadMissBMP As System.Windows.Forms.Button
    Public WithEvents txtMissBMP As System.Windows.Forms.TextBox
    Public WithEvents lblPlayRank As System.Windows.Forms.Label
    Public WithEvents lblTotal As System.Windows.Forms.Label
    Public WithEvents lblVolume As System.Windows.Forms.Label
    Public WithEvents lblStageFile As System.Windows.Forms.Label
    Public WithEvents lblMissBMP As System.Windows.Forms.Label
    Public WithEvents _fraTop_1 As System.Windows.Forms.Panel
    Public WithEvents cboDispFrame As System.Windows.Forms.ComboBox
    Public WithEvents cboDispSC2P As System.Windows.Forms.ComboBox
    Public WithEvents cboDispSC1P As System.Windows.Forms.ComboBox
    Public WithEvents cboDispKey As System.Windows.Forms.ComboBox
    Public WithEvents lblDispFrame As System.Windows.Forms.Label
    Public WithEvents lblDispSC2P As System.Windows.Forms.Label
    Public WithEvents lblDispSC1P As System.Windows.Forms.Label
    Public WithEvents lblDispKey As System.Windows.Forms.Label
    Public WithEvents _fraTop_2 As System.Windows.Forms.Panel
    Public WithEvents _optChangeTop_0 As System.Windows.Forms.RadioButton
    Public WithEvents _optChangeTop_2 As System.Windows.Forms.RadioButton
    Public WithEvents _optChangeTop_1 As System.Windows.Forms.RadioButton
    Public WithEvents fraHeader As System.Windows.Forms.Panel
    Public WithEvents _optChangeBottom_0 As System.Windows.Forms.RadioButton
    Public WithEvents _optChangeBottom_1 As System.Windows.Forms.RadioButton
    Public WithEvents _optChangeBottom_2 As System.Windows.Forms.RadioButton
    Public WithEvents _optChangeBottom_3 As System.Windows.Forms.RadioButton
    Public WithEvents _optChangeBottom_4 As System.Windows.Forms.RadioButton
    Public WithEvents txtExInfo As System.Windows.Forms.TextBox
    Public WithEvents _fraBottom_4 As System.Windows.Forms.Panel
    Public WithEvents cmdSoundExcUp As System.Windows.Forms.Button
    Public WithEvents cmdSoundExcDown As System.Windows.Forms.Button
    Public WithEvents cmdSoundDelete As System.Windows.Forms.Button
    Public WithEvents cmdSoundLoad As System.Windows.Forms.Button
    Public WithEvents cmdSoundStop As System.Windows.Forms.Button
    Public WithEvents lstWAV As System.Windows.Forms.ListBox
    Public WithEvents _fraBottom_0 As System.Windows.Forms.Panel
    Public WithEvents cmdBMPExcDown As System.Windows.Forms.Button
    Public WithEvents cmdBMPExcUp As System.Windows.Forms.Button
    Public WithEvents lstBMP As System.Windows.Forms.ListBox
    Public WithEvents cmdBMPDelete As System.Windows.Forms.Button
    Public WithEvents cmdBMPLoad As System.Windows.Forms.Button
    Public WithEvents cmdBMPPreview As System.Windows.Forms.Button
    Public WithEvents _fraBottom_1 As System.Windows.Forms.Panel
    Public WithEvents cmdBGAExcDown As System.Windows.Forms.Button
    Public WithEvents cmdBGAExcUp As System.Windows.Forms.Button
    Public WithEvents txtBGAInput As System.Windows.Forms.TextBox
    Public WithEvents cmdBGAPreview As System.Windows.Forms.Button
    Public WithEvents cmdBGASet As System.Windows.Forms.Button
    Public WithEvents cmdBGADelete As System.Windows.Forms.Button
    Public WithEvents lstBGA As System.Windows.Forms.ListBox
    Public WithEvents _fraBottom_2 As System.Windows.Forms.Panel
    Public WithEvents cboNumerator As System.Windows.Forms.ComboBox
    Public WithEvents cboDenominator As System.Windows.Forms.ComboBox
    Public WithEvents cmdMeasureSelectAll As System.Windows.Forms.Button
    Public WithEvents cmdInputMeasureLen As System.Windows.Forms.Button
    Public WithEvents lstMeasureLen As System.Windows.Forms.ListBox
    Public WithEvents lblFraction As System.Windows.Forms.Label
    Public WithEvents _fraBottom_3 As System.Windows.Forms.Panel
    Public WithEvents fraMaterial As System.Windows.Forms.Panel
    Public WithEvents picMain As System.Windows.Forms.PictureBox
    Public WithEvents hsbMain As System.Windows.Forms.HScrollBar
    Public WithEvents vsbMain As System.Windows.Forms.VScrollBar
    Public dlgMainOpen As System.Windows.Forms.OpenFileDialog
    Public dlgMainSave As System.Windows.Forms.SaveFileDialog
    Public WithEvents _New As System.Windows.Forms.ToolStripButton
    Public WithEvents Open As System.Windows.Forms.ToolStripSplitButton
    Public WithEvents Reload As System.Windows.Forms.ToolStripButton
    Public WithEvents Save As System.Windows.Forms.ToolStripButton
    Public WithEvents SaveAs As System.Windows.Forms.ToolStripButton
    Public WithEvents SepMode As System.Windows.Forms.ToolStripSeparator
    Public WithEvents Edit As System.Windows.Forms.ToolStripButton
    Public WithEvents Write As System.Windows.Forms.ToolStripButton
    Public WithEvents Delete As System.Windows.Forms.ToolStripButton
    Public WithEvents SepViewer As System.Windows.Forms.ToolStripSeparator
    Public WithEvents Viewer As System.Windows.Forms.ToolStripSeparator
    Public WithEvents PlayAll As System.Windows.Forms.ToolStripButton
    Public WithEvents Play As System.Windows.Forms.ToolStripButton
    Public WithEvents _Stop As System.Windows.Forms.ToolStripButton
    Public WithEvents SepGrid As System.Windows.Forms.ToolStripSeparator
    Public WithEvents ChangeGrid As System.Windows.Forms.ToolStripSeparator
    Public WithEvents SepSize As System.Windows.Forms.ToolStripSeparator
    Public WithEvents DispSize As System.Windows.Forms.ToolStripSeparator
    Public WithEvents SepResolution As System.Windows.Forms.ToolStripSeparator
    Public WithEvents Resolution As System.Windows.Forms.ToolStripSeparator
    Public WithEvents tlbMenu As System.Windows.Forms.ToolStrip
    Public WithEvents lblDirectInput As System.Windows.Forms.Label
    Public WithEvents txtKeys As System.Windows.Forms.TextBox
    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使って変更できます。
    'コード エディタを使用して、変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.MainMenu1 = New System.Windows.Forms.MenuStrip()
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileNew = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileOpen = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileSave = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileSaveAs = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileOpenDirectory = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLineFile = New System.Windows.Forms.ToolStripSeparator()
        Me._mnuRecentFiles_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuRecentFiles_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuRecentFiles_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuRecentFiles_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuRecentFiles_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuRecentFiles_5 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuRecentFiles_6 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuRecentFiles_7 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuRecentFiles_8 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuRecentFiles_9 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLineRecent = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuFileConvertWizard = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLineExit = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuFileExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEdit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditUndo = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditRedo = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLineEdit1 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuEditCut = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditCopy = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditPaste = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLineEdit2 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuEditSelectAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLineEdit3 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuEditFind = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLineEdit4 = New System.Windows.Forms.ToolStripSeparator()
        Me._mnuEditMode_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuEditMode_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuEditMode_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuView = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuViewItem_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuViewItem_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuViewItem_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOptions = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuOptionsItem_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuOptionsItem_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuOptionsItem_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuOptionsItem_3 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuOptionsItem_4 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuOptionsItem_5 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuOptionsItem_6 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuOptionsItem_7 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuOptionsItem_8 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuOptionsItem_9 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLineOptions = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuLanguageParent = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuLanguage_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuLanguage_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuLanguage_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuThemeParent = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuTheme_0 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuTheme_1 = New System.Windows.Forms.ToolStripMenuItem()
        Me._mnuTheme_2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTools = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuToolsPlayAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuToolsPlay = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuToolsPlayStop = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLineTools = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuToolsSetting = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelpOpen = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLineHelp = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuHelpWeb = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelpAbout = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuContext = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mnuContextPlayAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuContextPlay = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuContextBar1 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuContextInsertMeasure = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuContextDeleteMeasure = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuContextBar2 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuContextEditCut = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuContextEditCopy = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuContextEditPaste = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuContextEditDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuContextList = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mnuContextListLoad = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuContextListDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuContextListRename = New System.Windows.Forms.ToolStripMenuItem()
        Me.staMain = New System.Windows.Forms.StatusStrip()
        Me.Mode = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Position = New System.Windows.Forms.ToolStripStatusLabel()
        Me.WAV = New System.Windows.Forms.ToolStripStatusLabel()
        Me.BMP = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Measure = New System.Windows.Forms.ToolStripStatusLabel()
        Me._staMain_Panel6 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.picSiromaru = New System.Windows.Forms.PictureBox()
        Me.tmrEffect = New System.Windows.Forms.Timer(Me.components)
        Me.fraResolution = New System.Windows.Forms.Panel()
        Me.cboVScroll = New System.Windows.Forms.ComboBox()
        Me.lblVScroll = New System.Windows.Forms.Label()
        Me.fraDispSize = New System.Windows.Forms.Panel()
        Me.cboDispHeight = New System.Windows.Forms.ComboBox()
        Me.cboDispWidth = New System.Windows.Forms.ComboBox()
        Me.lblDispWidth = New System.Windows.Forms.Label()
        Me.lblDispHeight = New System.Windows.Forms.Label()
        Me.fraViewer = New System.Windows.Forms.Panel()
        Me.cboViewer = New System.Windows.Forms.ComboBox()
        Me.tmrMain = New System.Windows.Forms.Timer(Me.components)
        Me.ilsMenu = New System.Windows.Forms.ImageList(Me.components)
        Me.cboDirectInput = New System.Windows.Forms.ComboBox()
        Me.cmdDirectInput = New System.Windows.Forms.Button()
        Me.fraGrid = New System.Windows.Forms.Panel()
        Me.cboDispGridMain = New System.Windows.Forms.ComboBox()
        Me.cboDispGridSub = New System.Windows.Forms.ComboBox()
        Me.lblGridSub = New System.Windows.Forms.Label()
        Me.lblGridMain = New System.Windows.Forms.Label()
        Me.fraHeader = New System.Windows.Forms.Panel()
        Me._optChangeTop_2 = New System.Windows.Forms.RadioButton()
        Me._optChangeTop_1 = New System.Windows.Forms.RadioButton()
        Me._optChangeTop_0 = New System.Windows.Forms.RadioButton()
        Me._fraTop_0 = New System.Windows.Forms.Panel()
        Me.cboPlayer = New System.Windows.Forms.ComboBox()
        Me.txtGenre = New System.Windows.Forms.TextBox()
        Me.txtTitle = New System.Windows.Forms.TextBox()
        Me.txtArtist = New System.Windows.Forms.TextBox()
        Me.cboPlayLevel = New System.Windows.Forms.ComboBox()
        Me.txtBPM = New System.Windows.Forms.TextBox()
        Me.lblPlayMode = New System.Windows.Forms.Label()
        Me.lblGenre = New System.Windows.Forms.Label()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lblArtist = New System.Windows.Forms.Label()
        Me.lblPlayLevel = New System.Windows.Forms.Label()
        Me.lblBPM = New System.Windows.Forms.Label()
        Me._fraTop_1 = New System.Windows.Forms.Panel()
        Me.cboPlayRank = New System.Windows.Forms.ComboBox()
        Me.txtTotal = New System.Windows.Forms.TextBox()
        Me.txtVolume = New System.Windows.Forms.TextBox()
        Me.txtStageFile = New System.Windows.Forms.TextBox()
        Me.cmdLoadStageFile = New System.Windows.Forms.Button()
        Me.cmdLoadMissBMP = New System.Windows.Forms.Button()
        Me.txtMissBMP = New System.Windows.Forms.TextBox()
        Me.lblPlayRank = New System.Windows.Forms.Label()
        Me.lblTotal = New System.Windows.Forms.Label()
        Me.lblVolume = New System.Windows.Forms.Label()
        Me.lblStageFile = New System.Windows.Forms.Label()
        Me.lblMissBMP = New System.Windows.Forms.Label()
        Me._fraTop_2 = New System.Windows.Forms.Panel()
        Me.cboDispFrame = New System.Windows.Forms.ComboBox()
        Me.cboDispSC2P = New System.Windows.Forms.ComboBox()
        Me.cboDispSC1P = New System.Windows.Forms.ComboBox()
        Me.cboDispKey = New System.Windows.Forms.ComboBox()
        Me.lblDispFrame = New System.Windows.Forms.Label()
        Me.lblDispSC2P = New System.Windows.Forms.Label()
        Me.lblDispSC1P = New System.Windows.Forms.Label()
        Me.lblDispKey = New System.Windows.Forms.Label()
        Me.fraMaterial = New System.Windows.Forms.Panel()
        Me._optChangeBottom_0 = New System.Windows.Forms.RadioButton()
        Me._optChangeBottom_1 = New System.Windows.Forms.RadioButton()
        Me._optChangeBottom_2 = New System.Windows.Forms.RadioButton()
        Me._optChangeBottom_3 = New System.Windows.Forms.RadioButton()
        Me._optChangeBottom_4 = New System.Windows.Forms.RadioButton()
        Me._fraBottom_4 = New System.Windows.Forms.Panel()
        Me.txtExInfo = New System.Windows.Forms.TextBox()
        Me._fraBottom_0 = New System.Windows.Forms.Panel()
        Me.cmdSoundExcUp = New System.Windows.Forms.Button()
        Me.cmdSoundExcDown = New System.Windows.Forms.Button()
        Me.cmdSoundDelete = New System.Windows.Forms.Button()
        Me.cmdSoundLoad = New System.Windows.Forms.Button()
        Me.cmdSoundStop = New System.Windows.Forms.Button()
        Me.lstWAV = New System.Windows.Forms.ListBox()
        Me._fraBottom_1 = New System.Windows.Forms.Panel()
        Me.cmdBMPExcDown = New System.Windows.Forms.Button()
        Me.cmdBMPExcUp = New System.Windows.Forms.Button()
        Me.lstBMP = New System.Windows.Forms.ListBox()
        Me.cmdBMPDelete = New System.Windows.Forms.Button()
        Me.cmdBMPLoad = New System.Windows.Forms.Button()
        Me.cmdBMPPreview = New System.Windows.Forms.Button()
        Me._fraBottom_2 = New System.Windows.Forms.Panel()
        Me.cmdBGAExcDown = New System.Windows.Forms.Button()
        Me.cmdBGAExcUp = New System.Windows.Forms.Button()
        Me.txtBGAInput = New System.Windows.Forms.TextBox()
        Me.cmdBGAPreview = New System.Windows.Forms.Button()
        Me.cmdBGASet = New System.Windows.Forms.Button()
        Me.cmdBGADelete = New System.Windows.Forms.Button()
        Me.lstBGA = New System.Windows.Forms.ListBox()
        Me._fraBottom_3 = New System.Windows.Forms.Panel()
        Me.cboNumerator = New System.Windows.Forms.ComboBox()
        Me.cboDenominator = New System.Windows.Forms.ComboBox()
        Me.cmdMeasureSelectAll = New System.Windows.Forms.Button()
        Me.cmdInputMeasureLen = New System.Windows.Forms.Button()
        Me.lstMeasureLen = New System.Windows.Forms.ListBox()
        Me.lblFraction = New System.Windows.Forms.Label()
        Me.picMain = New System.Windows.Forms.PictureBox()
        Me.hsbMain = New System.Windows.Forms.HScrollBar()
        Me.vsbMain = New System.Windows.Forms.VScrollBar()
        Me.dlgMainOpen = New System.Windows.Forms.OpenFileDialog()
        Me.dlgMainSave = New System.Windows.Forms.SaveFileDialog()
        Me.tlbMenu = New System.Windows.Forms.ToolStrip()
        Me._New = New System.Windows.Forms.ToolStripButton()
        Me.Open = New System.Windows.Forms.ToolStripSplitButton()
        Me.ToolStripFileOpen = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripLineFile = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripMenuItem0 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem4 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem5 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem6 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem7 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem8 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem9 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Reload = New System.Windows.Forms.ToolStripButton()
        Me.Save = New System.Windows.Forms.ToolStripButton()
        Me.SaveAs = New System.Windows.Forms.ToolStripButton()
        Me.SepMode = New System.Windows.Forms.ToolStripSeparator()
        Me.Edit = New System.Windows.Forms.ToolStripButton()
        Me.Write = New System.Windows.Forms.ToolStripButton()
        Me.Delete = New System.Windows.Forms.ToolStripButton()
        Me.SepViewer = New System.Windows.Forms.ToolStripSeparator()
        Me.Viewer = New System.Windows.Forms.ToolStripSeparator()
        Me.PlayAll = New System.Windows.Forms.ToolStripButton()
        Me.Play = New System.Windows.Forms.ToolStripButton()
        Me._Stop = New System.Windows.Forms.ToolStripButton()
        Me.SepGrid = New System.Windows.Forms.ToolStripSeparator()
        Me.ChangeGrid = New System.Windows.Forms.ToolStripSeparator()
        Me.SepSize = New System.Windows.Forms.ToolStripSeparator()
        Me.DispSize = New System.Windows.Forms.ToolStripSeparator()
        Me.SepResolution = New System.Windows.Forms.ToolStripSeparator()
        Me.Resolution = New System.Windows.Forms.ToolStripSeparator()
        Me.lblDirectInput = New System.Windows.Forms.Label()
        Me.MainMenu1.SuspendLayout()
        Me.mnuContext.SuspendLayout()
        Me.mnuContextList.SuspendLayout()
        Me.staMain.SuspendLayout()
        CType(Me.picSiromaru, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraResolution.SuspendLayout()
        Me.fraDispSize.SuspendLayout()
        Me.fraViewer.SuspendLayout()
        Me.fraGrid.SuspendLayout()
        Me.fraHeader.SuspendLayout()
        Me._fraTop_0.SuspendLayout()
        Me._fraTop_1.SuspendLayout()
        Me._fraTop_2.SuspendLayout()
        Me.fraMaterial.SuspendLayout()
        Me._fraBottom_4.SuspendLayout()
        Me._fraBottom_0.SuspendLayout()
        Me._fraBottom_1.SuspendLayout()
        Me._fraBottom_2.SuspendLayout()
        Me._fraBottom_3.SuspendLayout()
        CType(Me.picMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tlbMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuEdit, Me.mnuView, Me.mnuOptions, Me.mnuTools, Me.mnuHelp})
        Me.MainMenu1.Location = New System.Drawing.Point(0, 0)
        Me.MainMenu1.Name = "MainMenu1"
        Me.MainMenu1.Size = New System.Drawing.Size(1192, 24)
        Me.MainMenu1.TabIndex = 101
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFileNew, Me.mnuFileOpen, Me.mnuFileSave, Me.mnuFileSaveAs, Me.mnuFileOpenDirectory, Me.mnuLineFile, Me._mnuRecentFiles_0, Me._mnuRecentFiles_1, Me._mnuRecentFiles_2, Me._mnuRecentFiles_3, Me._mnuRecentFiles_4, Me._mnuRecentFiles_5, Me._mnuRecentFiles_6, Me._mnuRecentFiles_7, Me._mnuRecentFiles_8, Me._mnuRecentFiles_9, Me.mnuLineRecent, Me.mnuFileConvertWizard, Me.mnuLineExit, Me.mnuFileExit})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(61, 20)
        Me.mnuFile.Text = "mnuFile"
        '
        'mnuFileNew
        '
        Me.mnuFileNew.Name = "mnuFileNew"
        Me.mnuFileNew.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
        Me.mnuFileNew.Size = New System.Drawing.Size(234, 22)
        Me.mnuFileNew.Text = "mnuFileNew"
        '
        'mnuFileOpen
        '
        Me.mnuFileOpen.Name = "mnuFileOpen"
        Me.mnuFileOpen.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.mnuFileOpen.Size = New System.Drawing.Size(234, 22)
        Me.mnuFileOpen.Text = "mnuFileOpen"
        '
        'mnuFileSave
        '
        Me.mnuFileSave.Name = "mnuFileSave"
        Me.mnuFileSave.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.mnuFileSave.Size = New System.Drawing.Size(234, 22)
        Me.mnuFileSave.Text = "mnuFileSave"
        '
        'mnuFileSaveAs
        '
        Me.mnuFileSaveAs.Name = "mnuFileSaveAs"
        Me.mnuFileSaveAs.Size = New System.Drawing.Size(234, 22)
        Me.mnuFileSaveAs.Text = "mnuFileSaveAs"
        '
        'mnuFileOpenDirectory
        '
        Me.mnuFileOpenDirectory.Name = "mnuFileOpenDirectory"
        Me.mnuFileOpenDirectory.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.D), System.Windows.Forms.Keys)
        Me.mnuFileOpenDirectory.Size = New System.Drawing.Size(234, 22)
        Me.mnuFileOpenDirectory.Text = "mnuFileOpenDirectory"
        '
        'mnuLineFile
        '
        Me.mnuLineFile.Name = "mnuLineFile"
        Me.mnuLineFile.Size = New System.Drawing.Size(231, 6)
        '
        '_mnuRecentFiles_0
        '
        Me._mnuRecentFiles_0.Name = "_mnuRecentFiles_0"
        Me._mnuRecentFiles_0.Size = New System.Drawing.Size(234, 22)
        Me._mnuRecentFiles_0.Text = "&1:"
        '
        '_mnuRecentFiles_1
        '
        Me._mnuRecentFiles_1.Name = "_mnuRecentFiles_1"
        Me._mnuRecentFiles_1.Size = New System.Drawing.Size(234, 22)
        Me._mnuRecentFiles_1.Text = "&2:"
        '
        '_mnuRecentFiles_2
        '
        Me._mnuRecentFiles_2.Name = "_mnuRecentFiles_2"
        Me._mnuRecentFiles_2.Size = New System.Drawing.Size(234, 22)
        Me._mnuRecentFiles_2.Text = "&3:"
        '
        '_mnuRecentFiles_3
        '
        Me._mnuRecentFiles_3.Name = "_mnuRecentFiles_3"
        Me._mnuRecentFiles_3.Size = New System.Drawing.Size(234, 22)
        Me._mnuRecentFiles_3.Text = "&4:"
        '
        '_mnuRecentFiles_4
        '
        Me._mnuRecentFiles_4.Name = "_mnuRecentFiles_4"
        Me._mnuRecentFiles_4.Size = New System.Drawing.Size(234, 22)
        Me._mnuRecentFiles_4.Text = "&5:"
        '
        '_mnuRecentFiles_5
        '
        Me._mnuRecentFiles_5.Name = "_mnuRecentFiles_5"
        Me._mnuRecentFiles_5.Size = New System.Drawing.Size(234, 22)
        Me._mnuRecentFiles_5.Text = "&6:"
        Me._mnuRecentFiles_5.Visible = False
        '
        '_mnuRecentFiles_6
        '
        Me._mnuRecentFiles_6.Name = "_mnuRecentFiles_6"
        Me._mnuRecentFiles_6.Size = New System.Drawing.Size(234, 22)
        Me._mnuRecentFiles_6.Text = "&7:"
        Me._mnuRecentFiles_6.Visible = False
        '
        '_mnuRecentFiles_7
        '
        Me._mnuRecentFiles_7.Name = "_mnuRecentFiles_7"
        Me._mnuRecentFiles_7.Size = New System.Drawing.Size(234, 22)
        Me._mnuRecentFiles_7.Text = "&8:"
        Me._mnuRecentFiles_7.Visible = False
        '
        '_mnuRecentFiles_8
        '
        Me._mnuRecentFiles_8.Name = "_mnuRecentFiles_8"
        Me._mnuRecentFiles_8.Size = New System.Drawing.Size(234, 22)
        Me._mnuRecentFiles_8.Text = "&9:"
        Me._mnuRecentFiles_8.Visible = False
        '
        '_mnuRecentFiles_9
        '
        Me._mnuRecentFiles_9.Name = "_mnuRecentFiles_9"
        Me._mnuRecentFiles_9.Size = New System.Drawing.Size(234, 22)
        Me._mnuRecentFiles_9.Text = "&0:"
        Me._mnuRecentFiles_9.Visible = False
        '
        'mnuLineRecent
        '
        Me.mnuLineRecent.Name = "mnuLineRecent"
        Me.mnuLineRecent.Size = New System.Drawing.Size(231, 6)
        '
        'mnuFileConvertWizard
        '
        Me.mnuFileConvertWizard.Name = "mnuFileConvertWizard"
        Me.mnuFileConvertWizard.Size = New System.Drawing.Size(234, 22)
        Me.mnuFileConvertWizard.Text = "mnuFileConvertWizard"
        '
        'mnuLineExit
        '
        Me.mnuLineExit.Name = "mnuLineExit"
        Me.mnuLineExit.Size = New System.Drawing.Size(231, 6)
        '
        'mnuFileExit
        '
        Me.mnuFileExit.Name = "mnuFileExit"
        Me.mnuFileExit.Size = New System.Drawing.Size(234, 22)
        Me.mnuFileExit.Text = "mnuFileExit"
        '
        'mnuEdit
        '
        Me.mnuEdit.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuEditUndo, Me.mnuEditRedo, Me.mnuLineEdit1, Me.mnuEditCut, Me.mnuEditCopy, Me.mnuEditPaste, Me.mnuEditDelete, Me.mnuLineEdit2, Me.mnuEditSelectAll, Me.mnuLineEdit3, Me.mnuEditFind, Me.mnuLineEdit4, Me._mnuEditMode_0, Me._mnuEditMode_1, Me._mnuEditMode_2})
        Me.mnuEdit.Name = "mnuEdit"
        Me.mnuEdit.Size = New System.Drawing.Size(63, 20)
        Me.mnuEdit.Text = "mnuEdit"
        '
        'mnuEditUndo
        '
        Me.mnuEditUndo.Name = "mnuEditUndo"
        Me.mnuEditUndo.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Z), System.Windows.Forms.Keys)
        Me.mnuEditUndo.Size = New System.Drawing.Size(204, 22)
        Me.mnuEditUndo.Text = "mnuEditUndo"
        '
        'mnuEditRedo
        '
        Me.mnuEditRedo.Name = "mnuEditRedo"
        Me.mnuEditRedo.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Y), System.Windows.Forms.Keys)
        Me.mnuEditRedo.Size = New System.Drawing.Size(204, 22)
        Me.mnuEditRedo.Text = "mnuEditRedo"
        '
        'mnuLineEdit1
        '
        Me.mnuLineEdit1.Name = "mnuLineEdit1"
        Me.mnuLineEdit1.Size = New System.Drawing.Size(201, 6)
        '
        'mnuEditCut
        '
        Me.mnuEditCut.Name = "mnuEditCut"
        Me.mnuEditCut.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
        Me.mnuEditCut.Size = New System.Drawing.Size(204, 22)
        Me.mnuEditCut.Text = "mnuEditCut"
        '
        'mnuEditCopy
        '
        Me.mnuEditCopy.Name = "mnuEditCopy"
        Me.mnuEditCopy.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.mnuEditCopy.Size = New System.Drawing.Size(204, 22)
        Me.mnuEditCopy.Text = "mnuEditCopy"
        '
        'mnuEditPaste
        '
        Me.mnuEditPaste.Name = "mnuEditPaste"
        Me.mnuEditPaste.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.V), System.Windows.Forms.Keys)
        Me.mnuEditPaste.Size = New System.Drawing.Size(204, 22)
        Me.mnuEditPaste.Text = "mnuEditPaste"
        '
        'mnuEditDelete
        '
        Me.mnuEditDelete.Name = "mnuEditDelete"
        Me.mnuEditDelete.ShortcutKeys = System.Windows.Forms.Keys.Delete
        Me.mnuEditDelete.Size = New System.Drawing.Size(204, 22)
        Me.mnuEditDelete.Text = "mnuEditDelete"
        '
        'mnuLineEdit2
        '
        Me.mnuLineEdit2.Name = "mnuLineEdit2"
        Me.mnuLineEdit2.Size = New System.Drawing.Size(201, 6)
        '
        'mnuEditSelectAll
        '
        Me.mnuEditSelectAll.Name = "mnuEditSelectAll"
        Me.mnuEditSelectAll.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
        Me.mnuEditSelectAll.Size = New System.Drawing.Size(204, 22)
        Me.mnuEditSelectAll.Text = "mnuEditSelectAll"
        '
        'mnuLineEdit3
        '
        Me.mnuLineEdit3.Name = "mnuLineEdit3"
        Me.mnuLineEdit3.Size = New System.Drawing.Size(201, 6)
        '
        'mnuEditFind
        '
        Me.mnuEditFind.Name = "mnuEditFind"
        Me.mnuEditFind.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)
        Me.mnuEditFind.Size = New System.Drawing.Size(204, 22)
        Me.mnuEditFind.Text = "mnuEditFind"
        '
        'mnuLineEdit4
        '
        Me.mnuLineEdit4.Name = "mnuLineEdit4"
        Me.mnuLineEdit4.Size = New System.Drawing.Size(201, 6)
        '
        '_mnuEditMode_0
        '
        Me._mnuEditMode_0.Name = "_mnuEditMode_0"
        Me._mnuEditMode_0.ShortcutKeys = System.Windows.Forms.Keys.F2
        Me._mnuEditMode_0.Size = New System.Drawing.Size(204, 22)
        Me._mnuEditMode_0.Text = "mnuEditMode(0)"
        '
        '_mnuEditMode_1
        '
        Me._mnuEditMode_1.Name = "_mnuEditMode_1"
        Me._mnuEditMode_1.ShortcutKeys = System.Windows.Forms.Keys.F3
        Me._mnuEditMode_1.Size = New System.Drawing.Size(204, 22)
        Me._mnuEditMode_1.Text = "mnuEditMode(1)"
        '
        '_mnuEditMode_2
        '
        Me._mnuEditMode_2.Name = "_mnuEditMode_2"
        Me._mnuEditMode_2.ShortcutKeys = System.Windows.Forms.Keys.F4
        Me._mnuEditMode_2.Size = New System.Drawing.Size(204, 22)
        Me._mnuEditMode_2.Text = "mnuEditMode(2)"
        '
        'mnuView
        '
        Me.mnuView.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuViewItem_0, Me._mnuViewItem_1, Me._mnuViewItem_2})
        Me.mnuView.Name = "mnuView"
        Me.mnuView.Size = New System.Drawing.Size(68, 20)
        Me.mnuView.Text = "mnuView"
        '
        '_mnuViewItem_0
        '
        Me._mnuViewItem_0.Name = "_mnuViewItem_0"
        Me._mnuViewItem_0.Size = New System.Drawing.Size(160, 22)
        Me._mnuViewItem_0.Text = "mnuViewItem(0)"
        '
        '_mnuViewItem_1
        '
        Me._mnuViewItem_1.Name = "_mnuViewItem_1"
        Me._mnuViewItem_1.Size = New System.Drawing.Size(160, 22)
        Me._mnuViewItem_1.Text = "mnuViewItem(1)"
        '
        '_mnuViewItem_2
        '
        Me._mnuViewItem_2.Name = "_mnuViewItem_2"
        Me._mnuViewItem_2.Size = New System.Drawing.Size(160, 22)
        Me._mnuViewItem_2.Text = "mnuViewItem(2)"
        '
        'mnuOptions
        '
        Me.mnuOptions.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuOptionsItem_0, Me._mnuOptionsItem_1, Me._mnuOptionsItem_2, Me._mnuOptionsItem_3, Me._mnuOptionsItem_4, Me._mnuOptionsItem_5, Me._mnuOptionsItem_6, Me._mnuOptionsItem_7, Me._mnuOptionsItem_8, Me._mnuOptionsItem_9, Me.mnuLineOptions, Me.mnuLanguageParent, Me.mnuThemeParent})
        Me.mnuOptions.Name = "mnuOptions"
        Me.mnuOptions.Size = New System.Drawing.Size(85, 20)
        Me.mnuOptions.Text = "mnuOptions"
        '
        '_mnuOptionsItem_0
        '
        Me._mnuOptionsItem_0.Checked = True
        Me._mnuOptionsItem_0.CheckState = System.Windows.Forms.CheckState.Checked
        Me._mnuOptionsItem_0.Name = "_mnuOptionsItem_0"
        Me._mnuOptionsItem_0.Size = New System.Drawing.Size(177, 22)
        Me._mnuOptionsItem_0.Text = "mnuOptionsItem(0)"
        '
        '_mnuOptionsItem_1
        '
        Me._mnuOptionsItem_1.Checked = True
        Me._mnuOptionsItem_1.CheckState = System.Windows.Forms.CheckState.Checked
        Me._mnuOptionsItem_1.Name = "_mnuOptionsItem_1"
        Me._mnuOptionsItem_1.Size = New System.Drawing.Size(177, 22)
        Me._mnuOptionsItem_1.Text = "mnuOptionsItem(1)"
        '
        '_mnuOptionsItem_2
        '
        Me._mnuOptionsItem_2.Checked = True
        Me._mnuOptionsItem_2.CheckState = System.Windows.Forms.CheckState.Checked
        Me._mnuOptionsItem_2.Name = "_mnuOptionsItem_2"
        Me._mnuOptionsItem_2.Size = New System.Drawing.Size(177, 22)
        Me._mnuOptionsItem_2.Text = "mnuOptionsItem(2)"
        '
        '_mnuOptionsItem_3
        '
        Me._mnuOptionsItem_3.Checked = True
        Me._mnuOptionsItem_3.CheckState = System.Windows.Forms.CheckState.Checked
        Me._mnuOptionsItem_3.Name = "_mnuOptionsItem_3"
        Me._mnuOptionsItem_3.Size = New System.Drawing.Size(177, 22)
        Me._mnuOptionsItem_3.Text = "mnuOptionsItem(3)"
        '
        '_mnuOptionsItem_4
        '
        Me._mnuOptionsItem_4.Checked = True
        Me._mnuOptionsItem_4.CheckState = System.Windows.Forms.CheckState.Checked
        Me._mnuOptionsItem_4.Name = "_mnuOptionsItem_4"
        Me._mnuOptionsItem_4.Size = New System.Drawing.Size(177, 22)
        Me._mnuOptionsItem_4.Text = "mnuOptionsItem(4)"
        '
        '_mnuOptionsItem_5
        '
        Me._mnuOptionsItem_5.Checked = True
        Me._mnuOptionsItem_5.CheckState = System.Windows.Forms.CheckState.Checked
        Me._mnuOptionsItem_5.Name = "_mnuOptionsItem_5"
        Me._mnuOptionsItem_5.Size = New System.Drawing.Size(177, 22)
        Me._mnuOptionsItem_5.Text = "mnuOptionsItem(5)"
        '
        '_mnuOptionsItem_6
        '
        Me._mnuOptionsItem_6.Checked = True
        Me._mnuOptionsItem_6.CheckState = System.Windows.Forms.CheckState.Checked
        Me._mnuOptionsItem_6.Name = "_mnuOptionsItem_6"
        Me._mnuOptionsItem_6.Size = New System.Drawing.Size(177, 22)
        Me._mnuOptionsItem_6.Text = "mnuOptionsItem(6)"
        '
        '_mnuOptionsItem_7
        '
        Me._mnuOptionsItem_7.Checked = True
        Me._mnuOptionsItem_7.CheckState = System.Windows.Forms.CheckState.Checked
        Me._mnuOptionsItem_7.Name = "_mnuOptionsItem_7"
        Me._mnuOptionsItem_7.Size = New System.Drawing.Size(177, 22)
        Me._mnuOptionsItem_7.Text = "mnuOptionsItem(7)"
        '
        '_mnuOptionsItem_8
        '
        Me._mnuOptionsItem_8.Checked = True
        Me._mnuOptionsItem_8.CheckState = System.Windows.Forms.CheckState.Checked
        Me._mnuOptionsItem_8.Name = "_mnuOptionsItem_8"
        Me._mnuOptionsItem_8.Size = New System.Drawing.Size(177, 22)
        Me._mnuOptionsItem_8.Text = "mnuOptionsItem(8)"
        '
        '_mnuOptionsItem_9
        '
        Me._mnuOptionsItem_9.Checked = True
        Me._mnuOptionsItem_9.CheckState = System.Windows.Forms.CheckState.Checked
        Me._mnuOptionsItem_9.Name = "_mnuOptionsItem_9"
        Me._mnuOptionsItem_9.Size = New System.Drawing.Size(177, 22)
        Me._mnuOptionsItem_9.Text = "mnuOptionsItem(9)"
        '
        'mnuLineOptions
        '
        Me.mnuLineOptions.Name = "mnuLineOptions"
        Me.mnuLineOptions.Size = New System.Drawing.Size(174, 6)
        '
        'mnuLanguageParent
        '
        Me.mnuLanguageParent.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuLanguage_0, Me._mnuLanguage_1, Me._mnuLanguage_2})
        Me.mnuLanguageParent.Name = "mnuLanguageParent"
        Me.mnuLanguageParent.Size = New System.Drawing.Size(177, 22)
        Me.mnuLanguageParent.Text = "Select &Language"
        '
        '_mnuLanguage_0
        '
        Me._mnuLanguage_0.Name = "_mnuLanguage_0"
        Me._mnuLanguage_0.Size = New System.Drawing.Size(164, 22)
        Me._mnuLanguage_0.Text = "mnuLanguage(0)"
        '
        '_mnuLanguage_1
        '
        Me._mnuLanguage_1.Name = "_mnuLanguage_1"
        Me._mnuLanguage_1.Size = New System.Drawing.Size(164, 22)
        Me._mnuLanguage_1.Text = "mnuLanguage(1)"
        '
        '_mnuLanguage_2
        '
        Me._mnuLanguage_2.Name = "_mnuLanguage_2"
        Me._mnuLanguage_2.Size = New System.Drawing.Size(164, 22)
        Me._mnuLanguage_2.Text = "mnuLanguage(2)"
        '
        'mnuThemeParent
        '
        Me.mnuThemeParent.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me._mnuTheme_0, Me._mnuTheme_1, Me._mnuTheme_2})
        Me.mnuThemeParent.Name = "mnuThemeParent"
        Me.mnuThemeParent.Size = New System.Drawing.Size(177, 22)
        Me.mnuThemeParent.Text = "Select &Theme"
        '
        '_mnuTheme_0
        '
        Me._mnuTheme_0.Name = "_mnuTheme_0"
        Me._mnuTheme_0.Size = New System.Drawing.Size(147, 22)
        Me._mnuTheme_0.Text = "mnuTheme(0)"
        '
        '_mnuTheme_1
        '
        Me._mnuTheme_1.Name = "_mnuTheme_1"
        Me._mnuTheme_1.Size = New System.Drawing.Size(147, 22)
        Me._mnuTheme_1.Text = "mnuTheme(1)"
        '
        '_mnuTheme_2
        '
        Me._mnuTheme_2.Name = "_mnuTheme_2"
        Me._mnuTheme_2.Size = New System.Drawing.Size(147, 22)
        Me._mnuTheme_2.Text = "mnuTheme(2)"
        '
        'mnuTools
        '
        Me.mnuTools.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuToolsPlayAll, Me.mnuToolsPlay, Me.mnuToolsPlayStop, Me.mnuLineTools, Me.mnuToolsSetting})
        Me.mnuTools.Name = "mnuTools"
        Me.mnuTools.Size = New System.Drawing.Size(70, 20)
        Me.mnuTools.Text = "mnuTools"
        '
        'mnuToolsPlayAll
        '
        Me.mnuToolsPlayAll.Name = "mnuToolsPlayAll"
        Me.mnuToolsPlayAll.ShortcutKeys = System.Windows.Forms.Keys.F6
        Me.mnuToolsPlayAll.Size = New System.Drawing.Size(190, 22)
        Me.mnuToolsPlayAll.Text = "mnuToolsPlayAll"
        '
        'mnuToolsPlay
        '
        Me.mnuToolsPlay.Name = "mnuToolsPlay"
        Me.mnuToolsPlay.ShortcutKeys = System.Windows.Forms.Keys.F7
        Me.mnuToolsPlay.Size = New System.Drawing.Size(190, 22)
        Me.mnuToolsPlay.Text = "mnuToolsPlay"
        '
        'mnuToolsPlayStop
        '
        Me.mnuToolsPlayStop.Name = "mnuToolsPlayStop"
        Me.mnuToolsPlayStop.ShortcutKeys = System.Windows.Forms.Keys.F8
        Me.mnuToolsPlayStop.Size = New System.Drawing.Size(190, 22)
        Me.mnuToolsPlayStop.Text = "mnuToolsPlayStop"
        '
        'mnuLineTools
        '
        Me.mnuLineTools.Name = "mnuLineTools"
        Me.mnuLineTools.Size = New System.Drawing.Size(187, 6)
        '
        'mnuToolsSetting
        '
        Me.mnuToolsSetting.Name = "mnuToolsSetting"
        Me.mnuToolsSetting.Size = New System.Drawing.Size(190, 22)
        Me.mnuToolsSetting.Text = "mnuToolsSetting"
        '
        'mnuHelp
        '
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuHelpOpen, Me.mnuLineHelp, Me.mnuHelpWeb, Me.mnuHelpAbout})
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(68, 20)
        Me.mnuHelp.Text = "mnuHelp"
        '
        'mnuHelpOpen
        '
        Me.mnuHelpOpen.Name = "mnuHelpOpen"
        Me.mnuHelpOpen.ShortcutKeys = System.Windows.Forms.Keys.F1
        Me.mnuHelpOpen.Size = New System.Drawing.Size(171, 22)
        Me.mnuHelpOpen.Text = "mnuHelpOpen"
        '
        'mnuLineHelp
        '
        Me.mnuLineHelp.Name = "mnuLineHelp"
        Me.mnuLineHelp.Size = New System.Drawing.Size(168, 6)
        '
        'mnuHelpWeb
        '
        Me.mnuHelpWeb.Name = "mnuHelpWeb"
        Me.mnuHelpWeb.Size = New System.Drawing.Size(171, 22)
        Me.mnuHelpWeb.Text = "mnuHelpWeb"
        '
        'mnuHelpAbout
        '
        Me.mnuHelpAbout.Name = "mnuHelpAbout"
        Me.mnuHelpAbout.Size = New System.Drawing.Size(171, 22)
        Me.mnuHelpAbout.Text = "mnuHelpAbout"
        '
        'mnuContext
        '
        Me.mnuContext.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.mnuContext.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuContextPlayAll, Me.mnuContextPlay, Me.mnuContextBar1, Me.mnuContextInsertMeasure, Me.mnuContextDeleteMeasure, Me.mnuContextBar2, Me.mnuContextEditCut, Me.mnuContextEditCopy, Me.mnuContextEditPaste, Me.mnuContextEditDelete})
        Me.mnuContext.Name = "mnuContext"
        Me.mnuContext.Size = New System.Drawing.Size(218, 192)
        Me.mnuContext.Text = "mnuContext"
        '
        'mnuContextPlayAll
        '
        Me.mnuContextPlayAll.Name = "mnuContextPlayAll"
        Me.mnuContextPlayAll.Size = New System.Drawing.Size(217, 22)
        Me.mnuContextPlayAll.Text = "mnuContextPlayAll"
        '
        'mnuContextPlay
        '
        Me.mnuContextPlay.Name = "mnuContextPlay"
        Me.mnuContextPlay.Size = New System.Drawing.Size(217, 22)
        Me.mnuContextPlay.Text = "mnuContextPlay"
        '
        'mnuContextBar1
        '
        Me.mnuContextBar1.Name = "mnuContextBar1"
        Me.mnuContextBar1.Size = New System.Drawing.Size(214, 6)
        '
        'mnuContextInsertMeasure
        '
        Me.mnuContextInsertMeasure.Name = "mnuContextInsertMeasure"
        Me.mnuContextInsertMeasure.Size = New System.Drawing.Size(217, 22)
        Me.mnuContextInsertMeasure.Text = "mnuContextInsertMeasure"
        '
        'mnuContextDeleteMeasure
        '
        Me.mnuContextDeleteMeasure.Name = "mnuContextDeleteMeasure"
        Me.mnuContextDeleteMeasure.Size = New System.Drawing.Size(217, 22)
        Me.mnuContextDeleteMeasure.Text = "mnuContextDeleteMeasure"
        '
        'mnuContextBar2
        '
        Me.mnuContextBar2.Name = "mnuContextBar2"
        Me.mnuContextBar2.Size = New System.Drawing.Size(214, 6)
        '
        'mnuContextEditCut
        '
        Me.mnuContextEditCut.Name = "mnuContextEditCut"
        Me.mnuContextEditCut.Size = New System.Drawing.Size(217, 22)
        Me.mnuContextEditCut.Text = "mnuContextEditCut"
        '
        'mnuContextEditCopy
        '
        Me.mnuContextEditCopy.Name = "mnuContextEditCopy"
        Me.mnuContextEditCopy.Size = New System.Drawing.Size(217, 22)
        Me.mnuContextEditCopy.Text = "mnuContextEditCopy"
        '
        'mnuContextEditPaste
        '
        Me.mnuContextEditPaste.Name = "mnuContextEditPaste"
        Me.mnuContextEditPaste.Size = New System.Drawing.Size(217, 22)
        Me.mnuContextEditPaste.Text = "mnuContextEditPaste"
        '
        'mnuContextEditDelete
        '
        Me.mnuContextEditDelete.Name = "mnuContextEditDelete"
        Me.mnuContextEditDelete.Size = New System.Drawing.Size(217, 22)
        Me.mnuContextEditDelete.Text = "mnuContextEditDelete"
        '
        'mnuContextList
        '
        Me.mnuContextList.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.mnuContextList.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuContextListLoad, Me.mnuContextListDelete, Me.mnuContextListRename})
        Me.mnuContextList.Name = "mnuContextList"
        Me.mnuContextList.Size = New System.Drawing.Size(200, 70)
        Me.mnuContextList.Text = "mnuContextList"
        '
        'mnuContextListLoad
        '
        Me.mnuContextListLoad.Name = "mnuContextListLoad"
        Me.mnuContextListLoad.Size = New System.Drawing.Size(199, 22)
        Me.mnuContextListLoad.Text = "mnuContextListLoad"
        '
        'mnuContextListDelete
        '
        Me.mnuContextListDelete.Name = "mnuContextListDelete"
        Me.mnuContextListDelete.Size = New System.Drawing.Size(199, 22)
        Me.mnuContextListDelete.Text = "mnuContextListDelete"
        '
        'mnuContextListRename
        '
        Me.mnuContextListRename.Name = "mnuContextListRename"
        Me.mnuContextListRename.Size = New System.Drawing.Size(199, 22)
        Me.mnuContextListRename.Text = "mnuContextListRename"
        '
        'staMain
        '
        Me.staMain.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.staMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Mode, Me.Position, Me.WAV, Me.BMP, Me.Measure, Me._staMain_Panel6})
        Me.staMain.Location = New System.Drawing.Point(0, 508)
        Me.staMain.Name = "staMain"
        Me.staMain.Size = New System.Drawing.Size(1192, 22)
        Me.staMain.TabIndex = 98
        '
        'Mode
        '
        Me.Mode.AutoSize = False
        Me.Mode.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.Mode.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me.Mode.Margin = New System.Windows.Forms.Padding(0)
        Me.Mode.Name = "Mode"
        Me.Mode.Size = New System.Drawing.Size(81, 22)
        Me.Mode.Text = "Edit Mode"
        '
        'Position
        '
        Me.Position.AutoSize = False
        Me.Position.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.Position.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me.Position.Margin = New System.Windows.Forms.Padding(0)
        Me.Position.Name = "Position"
        Me.Position.Size = New System.Drawing.Size(906, 22)
        Me.Position.Spring = True
        Me.Position.Text = "Position:"
        Me.Position.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'WAV
        '
        Me.WAV.AutoSize = False
        Me.WAV.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.WAV.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me.WAV.Margin = New System.Windows.Forms.Padding(0)
        Me.WAV.Name = "WAV"
        Me.WAV.Size = New System.Drawing.Size(68, 22)
        Me.WAV.Text = "#WAV 01"
        '
        'BMP
        '
        Me.BMP.AutoSize = False
        Me.BMP.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.BMP.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me.BMP.Margin = New System.Windows.Forms.Padding(0)
        Me.BMP.Name = "BMP"
        Me.BMP.Size = New System.Drawing.Size(68, 22)
        Me.BMP.Text = "#BMP 01"
        '
        'Measure
        '
        Me.Measure.AutoSize = False
        Me.Measure.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.Measure.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me.Measure.Margin = New System.Windows.Forms.Padding(0)
        Me.Measure.Name = "Measure"
        Me.Measure.Size = New System.Drawing.Size(54, 22)
        Me.Measure.Text = "4/4"
        '
        '_staMain_Panel6
        '
        Me._staMain_Panel6.AutoSize = False
        Me._staMain_Panel6.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me._staMain_Panel6.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me._staMain_Panel6.Margin = New System.Windows.Forms.Padding(0)
        Me._staMain_Panel6.Name = "_staMain_Panel6"
        Me._staMain_Panel6.Size = New System.Drawing.Size(54, 21)
        Me._staMain_Panel6.Text = "0ms"
        Me._staMain_Panel6.Visible = False
        '
        'picSiromaru
        '
        Me.picSiromaru.BackColor = System.Drawing.Color.Black
        Me.picSiromaru.Cursor = System.Windows.Forms.Cursors.Default
        Me.picSiromaru.ForeColor = System.Drawing.SystemColors.ControlText
        Me.picSiromaru.Image = CType(resources.GetObject("picSiromaru.Image"), System.Drawing.Image)
        Me.picSiromaru.Location = New System.Drawing.Point(1184, 56)
        Me.picSiromaru.Name = "picSiromaru"
        Me.picSiromaru.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.picSiromaru.Size = New System.Drawing.Size(64, 256)
        Me.picSiromaru.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.picSiromaru.TabIndex = 99
        Me.picSiromaru.TabStop = False
        Me.picSiromaru.Visible = False
        '
        'tmrEffect
        '
        Me.tmrEffect.Interval = 10
        '
        'fraResolution
        '
        Me.fraResolution.BackColor = System.Drawing.SystemColors.Control
        Me.fraResolution.Controls.Add(Me.cboVScroll)
        Me.fraResolution.Controls.Add(Me.lblVScroll)
        Me.fraResolution.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraResolution.Location = New System.Drawing.Point(1008, 200)
        Me.fraResolution.Name = "fraResolution"
        Me.fraResolution.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraResolution.Size = New System.Drawing.Size(117, 21)
        Me.fraResolution.TabIndex = 16
        '
        'cboVScroll
        '
        Me.cboVScroll.BackColor = System.Drawing.SystemColors.Window
        Me.cboVScroll.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboVScroll.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboVScroll.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboVScroll.Location = New System.Drawing.Point(48, 0)
        Me.cboVScroll.Name = "cboVScroll"
        Me.cboVScroll.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboVScroll.Size = New System.Drawing.Size(57, 20)
        Me.cboVScroll.TabIndex = 18
        '
        'lblVScroll
        '
        Me.lblVScroll.AutoSize = True
        Me.lblVScroll.BackColor = System.Drawing.SystemColors.Control
        Me.lblVScroll.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVScroll.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVScroll.Location = New System.Drawing.Point(0, 4)
        Me.lblVScroll.Name = "lblVScroll"
        Me.lblVScroll.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVScroll.Size = New System.Drawing.Size(47, 12)
        Me.lblVScroll.TabIndex = 17
        Me.lblVScroll.Text = "VScroll"
        '
        'fraDispSize
        '
        Me.fraDispSize.BackColor = System.Drawing.SystemColors.Control
        Me.fraDispSize.Controls.Add(Me.cboDispHeight)
        Me.fraDispSize.Controls.Add(Me.cboDispWidth)
        Me.fraDispSize.Controls.Add(Me.lblDispWidth)
        Me.fraDispSize.Controls.Add(Me.lblDispHeight)
        Me.fraDispSize.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraDispSize.Location = New System.Drawing.Point(808, 200)
        Me.fraDispSize.Name = "fraDispSize"
        Me.fraDispSize.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraDispSize.Size = New System.Drawing.Size(197, 21)
        Me.fraDispSize.TabIndex = 11
        '
        'cboDispHeight
        '
        Me.cboDispHeight.BackColor = System.Drawing.SystemColors.Window
        Me.cboDispHeight.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDispHeight.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDispHeight.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDispHeight.Location = New System.Drawing.Point(32, 0)
        Me.cboDispHeight.Name = "cboDispHeight"
        Me.cboDispHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDispHeight.Size = New System.Drawing.Size(57, 20)
        Me.cboDispHeight.TabIndex = 13
        '
        'cboDispWidth
        '
        Me.cboDispWidth.BackColor = System.Drawing.SystemColors.Window
        Me.cboDispWidth.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDispWidth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDispWidth.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDispWidth.Location = New System.Drawing.Point(116, 0)
        Me.cboDispWidth.Name = "cboDispWidth"
        Me.cboDispWidth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDispWidth.Size = New System.Drawing.Size(57, 20)
        Me.cboDispWidth.TabIndex = 15
        '
        'lblDispWidth
        '
        Me.lblDispWidth.AutoSize = True
        Me.lblDispWidth.BackColor = System.Drawing.SystemColors.Control
        Me.lblDispWidth.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDispWidth.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDispWidth.Location = New System.Drawing.Point(96, 4)
        Me.lblDispWidth.Name = "lblDispWidth"
        Me.lblDispWidth.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDispWidth.Size = New System.Drawing.Size(17, 12)
        Me.lblDispWidth.TabIndex = 14
        Me.lblDispWidth.Text = "幅"
        '
        'lblDispHeight
        '
        Me.lblDispHeight.AutoSize = True
        Me.lblDispHeight.BackColor = System.Drawing.SystemColors.Control
        Me.lblDispHeight.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDispHeight.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDispHeight.Location = New System.Drawing.Point(0, 4)
        Me.lblDispHeight.Name = "lblDispHeight"
        Me.lblDispHeight.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDispHeight.Size = New System.Drawing.Size(29, 12)
        Me.lblDispHeight.TabIndex = 12
        Me.lblDispHeight.Text = "高さ"
        '
        'fraViewer
        '
        Me.fraViewer.BackColor = System.Drawing.SystemColors.Control
        Me.fraViewer.Controls.Add(Me.cboViewer)
        Me.fraViewer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraViewer.Location = New System.Drawing.Point(808, 136)
        Me.fraViewer.Name = "fraViewer"
        Me.fraViewer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraViewer.Size = New System.Drawing.Size(97, 21)
        Me.fraViewer.TabIndex = 4
        '
        'cboViewer
        '
        Me.cboViewer.BackColor = System.Drawing.SystemColors.Window
        Me.cboViewer.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboViewer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboViewer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboViewer.Location = New System.Drawing.Point(0, 0)
        Me.cboViewer.Name = "cboViewer"
        Me.cboViewer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboViewer.Size = New System.Drawing.Size(89, 20)
        Me.cboViewer.TabIndex = 5
        '
        'tmrMain
        '
        '
        'ilsMenu
        '
        Me.ilsMenu.ImageStream = CType(resources.GetObject("ilsMenu.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ilsMenu.TransparentColor = System.Drawing.Color.White
        Me.ilsMenu.Images.SetKeyName(0, "")
        Me.ilsMenu.Images.SetKeyName(1, "")
        Me.ilsMenu.Images.SetKeyName(2, "")
        Me.ilsMenu.Images.SetKeyName(3, "")
        Me.ilsMenu.Images.SetKeyName(4, "")
        Me.ilsMenu.Images.SetKeyName(5, "")
        Me.ilsMenu.Images.SetKeyName(6, "")
        Me.ilsMenu.Images.SetKeyName(7, "")
        Me.ilsMenu.Images.SetKeyName(8, "")
        Me.ilsMenu.Images.SetKeyName(9, "")
        Me.ilsMenu.Images.SetKeyName(10, "")
        '
        'cboDirectInput
        '
        Me.cboDirectInput.BackColor = System.Drawing.SystemColors.Window
        Me.cboDirectInput.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDirectInput.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDirectInput.Location = New System.Drawing.Point(40, 480)
        Me.cboDirectInput.Name = "cboDirectInput"
        Me.cboDirectInput.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDirectInput.Size = New System.Drawing.Size(29, 20)
        Me.cboDirectInput.TabIndex = 96
        '
        'cmdDirectInput
        '
        Me.cmdDirectInput.BackColor = System.Drawing.SystemColors.Control
        Me.cmdDirectInput.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdDirectInput.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdDirectInput.Location = New System.Drawing.Point(72, 480)
        Me.cmdDirectInput.Name = "cmdDirectInput"
        Me.cmdDirectInput.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdDirectInput.Size = New System.Drawing.Size(61, 21)
        Me.cmdDirectInput.TabIndex = 97
        Me.cmdDirectInput.Text = "Input"
        Me.cmdDirectInput.UseVisualStyleBackColor = False
        '
        'fraGrid
        '
        Me.fraGrid.BackColor = System.Drawing.SystemColors.Control
        Me.fraGrid.Controls.Add(Me.cboDispGridMain)
        Me.fraGrid.Controls.Add(Me.cboDispGridSub)
        Me.fraGrid.Controls.Add(Me.lblGridSub)
        Me.fraGrid.Controls.Add(Me.lblGridMain)
        Me.fraGrid.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraGrid.Location = New System.Drawing.Point(808, 168)
        Me.fraGrid.Name = "fraGrid"
        Me.fraGrid.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraGrid.Size = New System.Drawing.Size(197, 21)
        Me.fraGrid.TabIndex = 6
        '
        'cboDispGridMain
        '
        Me.cboDispGridMain.BackColor = System.Drawing.SystemColors.Window
        Me.cboDispGridMain.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDispGridMain.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDispGridMain.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDispGridMain.Location = New System.Drawing.Point(128, 0)
        Me.cboDispGridMain.Name = "cboDispGridMain"
        Me.cboDispGridMain.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDispGridMain.Size = New System.Drawing.Size(57, 20)
        Me.cboDispGridMain.TabIndex = 10
        '
        'cboDispGridSub
        '
        Me.cboDispGridSub.BackColor = System.Drawing.SystemColors.Window
        Me.cboDispGridSub.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDispGridSub.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDispGridSub.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDispGridSub.Location = New System.Drawing.Point(32, 0)
        Me.cboDispGridSub.Name = "cboDispGridSub"
        Me.cboDispGridSub.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDispGridSub.Size = New System.Drawing.Size(57, 20)
        Me.cboDispGridSub.TabIndex = 8
        '
        'lblGridSub
        '
        Me.lblGridSub.AutoSize = True
        Me.lblGridSub.BackColor = System.Drawing.SystemColors.Control
        Me.lblGridSub.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblGridSub.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblGridSub.Location = New System.Drawing.Point(96, 4)
        Me.lblGridSub.Name = "lblGridSub"
        Me.lblGridSub.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblGridSub.Size = New System.Drawing.Size(29, 12)
        Me.lblGridSub.TabIndex = 9
        Me.lblGridSub.Text = "補助"
        '
        'lblGridMain
        '
        Me.lblGridMain.AutoSize = True
        Me.lblGridMain.BackColor = System.Drawing.SystemColors.Control
        Me.lblGridMain.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblGridMain.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblGridMain.Location = New System.Drawing.Point(0, 4)
        Me.lblGridMain.Name = "lblGridMain"
        Me.lblGridMain.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblGridMain.Size = New System.Drawing.Size(29, 12)
        Me.lblGridMain.TabIndex = 7
        Me.lblGridMain.Text = "Grid"
        '
        'fraHeader
        '
        Me.fraHeader.BackColor = System.Drawing.SystemColors.Control
        Me.fraHeader.Controls.Add(Me._optChangeTop_2)
        Me.fraHeader.Controls.Add(Me._optChangeTop_1)
        Me.fraHeader.Controls.Add(Me._optChangeTop_0)
        Me.fraHeader.Controls.Add(Me._fraTop_0)
        Me.fraHeader.Controls.Add(Me._fraTop_1)
        Me.fraHeader.Controls.Add(Me._fraTop_2)
        Me.fraHeader.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraHeader.Location = New System.Drawing.Point(140, 56)
        Me.fraHeader.Name = "fraHeader"
        Me.fraHeader.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraHeader.Size = New System.Drawing.Size(666, 165)
        Me.fraHeader.TabIndex = 19
        '
        '_optChangeTop_2
        '
        Me._optChangeTop_2.Appearance = System.Windows.Forms.Appearance.Button
        Me._optChangeTop_2.BackColor = System.Drawing.SystemColors.Control
        Me._optChangeTop_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._optChangeTop_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optChangeTop_2.Location = New System.Drawing.Point(130, 0)
        Me._optChangeTop_2.Name = "_optChangeTop_2"
        Me._optChangeTop_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optChangeTop_2.Size = New System.Drawing.Size(61, 21)
        Me._optChangeTop_2.TabIndex = 22
        Me._optChangeTop_2.TabStop = True
        Me._optChangeTop_2.Text = "環境"
        Me._optChangeTop_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me._optChangeTop_2.UseVisualStyleBackColor = False
        '
        '_optChangeTop_1
        '
        Me._optChangeTop_1.Appearance = System.Windows.Forms.Appearance.Button
        Me._optChangeTop_1.BackColor = System.Drawing.SystemColors.Control
        Me._optChangeTop_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optChangeTop_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optChangeTop_1.Location = New System.Drawing.Point(65, 0)
        Me._optChangeTop_1.Name = "_optChangeTop_1"
        Me._optChangeTop_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optChangeTop_1.Size = New System.Drawing.Size(61, 21)
        Me._optChangeTop_1.TabIndex = 21
        Me._optChangeTop_1.TabStop = True
        Me._optChangeTop_1.Text = "拡張"
        Me._optChangeTop_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me._optChangeTop_1.UseVisualStyleBackColor = False
        '
        '_optChangeTop_0
        '
        Me._optChangeTop_0.Appearance = System.Windows.Forms.Appearance.Button
        Me._optChangeTop_0.BackColor = System.Drawing.SystemColors.Control
        Me._optChangeTop_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optChangeTop_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optChangeTop_0.Location = New System.Drawing.Point(0, 0)
        Me._optChangeTop_0.Name = "_optChangeTop_0"
        Me._optChangeTop_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optChangeTop_0.Size = New System.Drawing.Size(61, 21)
        Me._optChangeTop_0.TabIndex = 20
        Me._optChangeTop_0.TabStop = True
        Me._optChangeTop_0.Text = "基本"
        Me._optChangeTop_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me._optChangeTop_0.UseVisualStyleBackColor = False
        '
        '_fraTop_0
        '
        Me._fraTop_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraTop_0.Controls.Add(Me.cboPlayer)
        Me._fraTop_0.Controls.Add(Me.txtGenre)
        Me._fraTop_0.Controls.Add(Me.txtTitle)
        Me._fraTop_0.Controls.Add(Me.txtArtist)
        Me._fraTop_0.Controls.Add(Me.cboPlayLevel)
        Me._fraTop_0.Controls.Add(Me.txtBPM)
        Me._fraTop_0.Controls.Add(Me.lblPlayMode)
        Me._fraTop_0.Controls.Add(Me.lblGenre)
        Me._fraTop_0.Controls.Add(Me.lblTitle)
        Me._fraTop_0.Controls.Add(Me.lblArtist)
        Me._fraTop_0.Controls.Add(Me.lblPlayLevel)
        Me._fraTop_0.Controls.Add(Me.lblBPM)
        Me._fraTop_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraTop_0.Location = New System.Drawing.Point(0, 32)
        Me._fraTop_0.Name = "_fraTop_0"
        Me._fraTop_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraTop_0.Size = New System.Drawing.Size(217, 125)
        Me._fraTop_0.TabIndex = 23
        Me._fraTop_0.Visible = False
        '
        'cboPlayer
        '
        Me.cboPlayer.BackColor = System.Drawing.SystemColors.Window
        Me.cboPlayer.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboPlayer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPlayer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboPlayer.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.cboPlayer.Location = New System.Drawing.Point(80, 8)
        Me.cboPlayer.Name = "cboPlayer"
        Me.cboPlayer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboPlayer.Size = New System.Drawing.Size(133, 20)
        Me.cboPlayer.TabIndex = 25
        '
        'txtGenre
        '
        Me.txtGenre.AcceptsReturn = True
        Me.txtGenre.BackColor = System.Drawing.SystemColors.Window
        Me.txtGenre.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtGenre.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtGenre.Location = New System.Drawing.Point(80, 32)
        Me.txtGenre.MaxLength = 0
        Me.txtGenre.Name = "txtGenre"
        Me.txtGenre.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtGenre.Size = New System.Drawing.Size(133, 19)
        Me.txtGenre.TabIndex = 27
        '
        'txtTitle
        '
        Me.txtTitle.AcceptsReturn = True
        Me.txtTitle.BackColor = System.Drawing.SystemColors.Window
        Me.txtTitle.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTitle.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTitle.Location = New System.Drawing.Point(80, 56)
        Me.txtTitle.MaxLength = 0
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTitle.Size = New System.Drawing.Size(133, 19)
        Me.txtTitle.TabIndex = 29
        '
        'txtArtist
        '
        Me.txtArtist.AcceptsReturn = True
        Me.txtArtist.BackColor = System.Drawing.SystemColors.Window
        Me.txtArtist.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtArtist.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtArtist.Location = New System.Drawing.Point(80, 80)
        Me.txtArtist.MaxLength = 0
        Me.txtArtist.Name = "txtArtist"
        Me.txtArtist.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtArtist.Size = New System.Drawing.Size(133, 19)
        Me.txtArtist.TabIndex = 31
        '
        'cboPlayLevel
        '
        Me.cboPlayLevel.BackColor = System.Drawing.SystemColors.Window
        Me.cboPlayLevel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboPlayLevel.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboPlayLevel.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.cboPlayLevel.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "0"})
        Me.cboPlayLevel.Location = New System.Drawing.Point(80, 104)
        Me.cboPlayLevel.Name = "cboPlayLevel"
        Me.cboPlayLevel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboPlayLevel.Size = New System.Drawing.Size(49, 20)
        Me.cboPlayLevel.TabIndex = 33
        '
        'txtBPM
        '
        Me.txtBPM.AcceptsReturn = True
        Me.txtBPM.BackColor = System.Drawing.SystemColors.Window
        Me.txtBPM.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBPM.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBPM.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtBPM.Location = New System.Drawing.Point(172, 104)
        Me.txtBPM.MaxLength = 0
        Me.txtBPM.Name = "txtBPM"
        Me.txtBPM.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBPM.Size = New System.Drawing.Size(41, 19)
        Me.txtBPM.TabIndex = 35
        Me.txtBPM.Text = "130"
        '
        'lblPlayMode
        '
        Me.lblPlayMode.BackColor = System.Drawing.SystemColors.Control
        Me.lblPlayMode.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPlayMode.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPlayMode.Location = New System.Drawing.Point(4, 12)
        Me.lblPlayMode.Name = "lblPlayMode"
        Me.lblPlayMode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPlayMode.Size = New System.Drawing.Size(72, 12)
        Me.lblPlayMode.TabIndex = 24
        Me.lblPlayMode.Text = "プレイモード"
        Me.lblPlayMode.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblGenre
        '
        Me.lblGenre.BackColor = System.Drawing.SystemColors.Control
        Me.lblGenre.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblGenre.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblGenre.Location = New System.Drawing.Point(4, 36)
        Me.lblGenre.Name = "lblGenre"
        Me.lblGenre.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblGenre.Size = New System.Drawing.Size(72, 12)
        Me.lblGenre.TabIndex = 26
        Me.lblGenre.Text = "ジャンル"
        Me.lblGenre.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTitle
        '
        Me.lblTitle.BackColor = System.Drawing.SystemColors.Control
        Me.lblTitle.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTitle.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTitle.Location = New System.Drawing.Point(4, 60)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTitle.Size = New System.Drawing.Size(72, 12)
        Me.lblTitle.TabIndex = 28
        Me.lblTitle.Text = "タイトル"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblArtist
        '
        Me.lblArtist.BackColor = System.Drawing.SystemColors.Control
        Me.lblArtist.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblArtist.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblArtist.Location = New System.Drawing.Point(4, 84)
        Me.lblArtist.Name = "lblArtist"
        Me.lblArtist.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblArtist.Size = New System.Drawing.Size(72, 12)
        Me.lblArtist.TabIndex = 30
        Me.lblArtist.Text = "アーティスト"
        Me.lblArtist.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblPlayLevel
        '
        Me.lblPlayLevel.BackColor = System.Drawing.SystemColors.Control
        Me.lblPlayLevel.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPlayLevel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPlayLevel.Location = New System.Drawing.Point(4, 108)
        Me.lblPlayLevel.Name = "lblPlayLevel"
        Me.lblPlayLevel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPlayLevel.Size = New System.Drawing.Size(72, 12)
        Me.lblPlayLevel.TabIndex = 32
        Me.lblPlayLevel.Text = "難易度表示"
        Me.lblPlayLevel.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblBPM
        '
        Me.lblBPM.BackColor = System.Drawing.SystemColors.Control
        Me.lblBPM.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblBPM.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBPM.Location = New System.Drawing.Point(133, 108)
        Me.lblBPM.Name = "lblBPM"
        Me.lblBPM.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblBPM.Size = New System.Drawing.Size(35, 12)
        Me.lblBPM.TabIndex = 34
        Me.lblBPM.Text = "BPM"
        Me.lblBPM.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_fraTop_1
        '
        Me._fraTop_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraTop_1.Controls.Add(Me.cboPlayRank)
        Me._fraTop_1.Controls.Add(Me.txtTotal)
        Me._fraTop_1.Controls.Add(Me.txtVolume)
        Me._fraTop_1.Controls.Add(Me.txtStageFile)
        Me._fraTop_1.Controls.Add(Me.cmdLoadStageFile)
        Me._fraTop_1.Controls.Add(Me.cmdLoadMissBMP)
        Me._fraTop_1.Controls.Add(Me.txtMissBMP)
        Me._fraTop_1.Controls.Add(Me.lblPlayRank)
        Me._fraTop_1.Controls.Add(Me.lblTotal)
        Me._fraTop_1.Controls.Add(Me.lblVolume)
        Me._fraTop_1.Controls.Add(Me.lblStageFile)
        Me._fraTop_1.Controls.Add(Me.lblMissBMP)
        Me._fraTop_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraTop_1.Location = New System.Drawing.Point(220, 32)
        Me._fraTop_1.Name = "_fraTop_1"
        Me._fraTop_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraTop_1.Size = New System.Drawing.Size(217, 125)
        Me._fraTop_1.TabIndex = 36
        Me._fraTop_1.Visible = False
        '
        'cboPlayRank
        '
        Me.cboPlayRank.BackColor = System.Drawing.SystemColors.Window
        Me.cboPlayRank.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboPlayRank.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPlayRank.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboPlayRank.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.cboPlayRank.Items.AddRange(New Object() {"Very Hard", "Hard", "Normal", "Easy"})
        Me.cboPlayRank.Location = New System.Drawing.Point(80, 8)
        Me.cboPlayRank.Name = "cboPlayRank"
        Me.cboPlayRank.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboPlayRank.Size = New System.Drawing.Size(133, 20)
        Me.cboPlayRank.TabIndex = 38
        '
        'txtTotal
        '
        Me.txtTotal.AcceptsReturn = True
        Me.txtTotal.BackColor = System.Drawing.SystemColors.Window
        Me.txtTotal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTotal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTotal.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtTotal.Location = New System.Drawing.Point(80, 32)
        Me.txtTotal.MaxLength = 0
        Me.txtTotal.Name = "txtTotal"
        Me.txtTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTotal.Size = New System.Drawing.Size(133, 19)
        Me.txtTotal.TabIndex = 40
        '
        'txtVolume
        '
        Me.txtVolume.AcceptsReturn = True
        Me.txtVolume.BackColor = System.Drawing.SystemColors.Window
        Me.txtVolume.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtVolume.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVolume.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtVolume.Location = New System.Drawing.Point(80, 56)
        Me.txtVolume.MaxLength = 0
        Me.txtVolume.Name = "txtVolume"
        Me.txtVolume.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtVolume.Size = New System.Drawing.Size(133, 19)
        Me.txtVolume.TabIndex = 42
        '
        'txtStageFile
        '
        Me.txtStageFile.AcceptsReturn = True
        Me.txtStageFile.BackColor = System.Drawing.SystemColors.Window
        Me.txtStageFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtStageFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtStageFile.Location = New System.Drawing.Point(80, 80)
        Me.txtStageFile.MaxLength = 0
        Me.txtStageFile.Name = "txtStageFile"
        Me.txtStageFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtStageFile.Size = New System.Drawing.Size(93, 19)
        Me.txtStageFile.TabIndex = 44
        '
        'cmdLoadStageFile
        '
        Me.cmdLoadStageFile.BackColor = System.Drawing.SystemColors.Control
        Me.cmdLoadStageFile.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdLoadStageFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdLoadStageFile.Location = New System.Drawing.Point(176, 80)
        Me.cmdLoadStageFile.Name = "cmdLoadStageFile"
        Me.cmdLoadStageFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdLoadStageFile.Size = New System.Drawing.Size(37, 17)
        Me.cmdLoadStageFile.TabIndex = 45
        Me.cmdLoadStageFile.Text = "参照"
        Me.cmdLoadStageFile.UseVisualStyleBackColor = False
        '
        'cmdLoadMissBMP
        '
        Me.cmdLoadMissBMP.BackColor = System.Drawing.SystemColors.Control
        Me.cmdLoadMissBMP.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdLoadMissBMP.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdLoadMissBMP.Location = New System.Drawing.Point(176, 104)
        Me.cmdLoadMissBMP.Name = "cmdLoadMissBMP"
        Me.cmdLoadMissBMP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdLoadMissBMP.Size = New System.Drawing.Size(37, 17)
        Me.cmdLoadMissBMP.TabIndex = 48
        Me.cmdLoadMissBMP.Text = "参照"
        Me.cmdLoadMissBMP.UseVisualStyleBackColor = False
        '
        'txtMissBMP
        '
        Me.txtMissBMP.AcceptsReturn = True
        Me.txtMissBMP.BackColor = System.Drawing.SystemColors.Window
        Me.txtMissBMP.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMissBMP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMissBMP.Location = New System.Drawing.Point(80, 104)
        Me.txtMissBMP.MaxLength = 0
        Me.txtMissBMP.Name = "txtMissBMP"
        Me.txtMissBMP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMissBMP.Size = New System.Drawing.Size(93, 19)
        Me.txtMissBMP.TabIndex = 47
        '
        'lblPlayRank
        '
        Me.lblPlayRank.BackColor = System.Drawing.SystemColors.Control
        Me.lblPlayRank.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPlayRank.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPlayRank.Location = New System.Drawing.Point(4, 12)
        Me.lblPlayRank.Name = "lblPlayRank"
        Me.lblPlayRank.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPlayRank.Size = New System.Drawing.Size(72, 12)
        Me.lblPlayRank.TabIndex = 37
        Me.lblPlayRank.Text = "#RANK"
        Me.lblPlayRank.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotal
        '
        Me.lblTotal.BackColor = System.Drawing.SystemColors.Control
        Me.lblTotal.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTotal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTotal.Location = New System.Drawing.Point(4, 36)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTotal.Size = New System.Drawing.Size(72, 12)
        Me.lblTotal.TabIndex = 39
        Me.lblTotal.Text = "#TOTAL"
        Me.lblTotal.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblVolume
        '
        Me.lblVolume.BackColor = System.Drawing.SystemColors.Control
        Me.lblVolume.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVolume.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVolume.Location = New System.Drawing.Point(4, 60)
        Me.lblVolume.Name = "lblVolume"
        Me.lblVolume.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVolume.Size = New System.Drawing.Size(72, 12)
        Me.lblVolume.TabIndex = 41
        Me.lblVolume.Text = "#VOLWAV"
        Me.lblVolume.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblStageFile
        '
        Me.lblStageFile.BackColor = System.Drawing.SystemColors.Control
        Me.lblStageFile.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblStageFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblStageFile.Location = New System.Drawing.Point(4, 84)
        Me.lblStageFile.Name = "lblStageFile"
        Me.lblStageFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblStageFile.Size = New System.Drawing.Size(72, 12)
        Me.lblStageFile.TabIndex = 43
        Me.lblStageFile.Text = "#STAGEFILE"
        Me.lblStageFile.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblMissBMP
        '
        Me.lblMissBMP.BackColor = System.Drawing.SystemColors.Control
        Me.lblMissBMP.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMissBMP.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMissBMP.Location = New System.Drawing.Point(4, 108)
        Me.lblMissBMP.Name = "lblMissBMP"
        Me.lblMissBMP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMissBMP.Size = New System.Drawing.Size(72, 12)
        Me.lblMissBMP.TabIndex = 46
        Me.lblMissBMP.Text = "#BMP00"
        Me.lblMissBMP.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_fraTop_2
        '
        Me._fraTop_2.BackColor = System.Drawing.SystemColors.Control
        Me._fraTop_2.Controls.Add(Me.cboDispFrame)
        Me._fraTop_2.Controls.Add(Me.cboDispSC2P)
        Me._fraTop_2.Controls.Add(Me.cboDispSC1P)
        Me._fraTop_2.Controls.Add(Me.cboDispKey)
        Me._fraTop_2.Controls.Add(Me.lblDispFrame)
        Me._fraTop_2.Controls.Add(Me.lblDispSC2P)
        Me._fraTop_2.Controls.Add(Me.lblDispSC1P)
        Me._fraTop_2.Controls.Add(Me.lblDispKey)
        Me._fraTop_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraTop_2.Location = New System.Drawing.Point(440, 32)
        Me._fraTop_2.Name = "_fraTop_2"
        Me._fraTop_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraTop_2.Size = New System.Drawing.Size(217, 125)
        Me._fraTop_2.TabIndex = 49
        Me._fraTop_2.Visible = False
        '
        'cboDispFrame
        '
        Me.cboDispFrame.BackColor = System.Drawing.SystemColors.Window
        Me.cboDispFrame.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDispFrame.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDispFrame.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDispFrame.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.cboDispFrame.Items.AddRange(New Object() {"ハーフ", "セパレート"})
        Me.cboDispFrame.Location = New System.Drawing.Point(80, 8)
        Me.cboDispFrame.Name = "cboDispFrame"
        Me.cboDispFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDispFrame.Size = New System.Drawing.Size(133, 20)
        Me.cboDispFrame.TabIndex = 51
        '
        'cboDispSC2P
        '
        Me.cboDispSC2P.BackColor = System.Drawing.SystemColors.Window
        Me.cboDispSC2P.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDispSC2P.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDispSC2P.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDispSC2P.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.cboDispSC2P.Items.AddRange(New Object() {"左", "右"})
        Me.cboDispSC2P.Location = New System.Drawing.Point(164, 56)
        Me.cboDispSC2P.Name = "cboDispSC2P"
        Me.cboDispSC2P.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDispSC2P.Size = New System.Drawing.Size(49, 20)
        Me.cboDispSC2P.TabIndex = 57
        '
        'cboDispSC1P
        '
        Me.cboDispSC1P.BackColor = System.Drawing.SystemColors.Window
        Me.cboDispSC1P.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDispSC1P.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDispSC1P.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDispSC1P.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.cboDispSC1P.Items.AddRange(New Object() {"左", "右"})
        Me.cboDispSC1P.Location = New System.Drawing.Point(80, 56)
        Me.cboDispSC1P.Name = "cboDispSC1P"
        Me.cboDispSC1P.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDispSC1P.Size = New System.Drawing.Size(49, 20)
        Me.cboDispSC1P.TabIndex = 55
        '
        'cboDispKey
        '
        Me.cboDispKey.BackColor = System.Drawing.SystemColors.Window
        Me.cboDispKey.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDispKey.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDispKey.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDispKey.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.cboDispKey.Items.AddRange(New Object() {"5Keys/10Keys", "7Keys/14Keys"})
        Me.cboDispKey.Location = New System.Drawing.Point(80, 32)
        Me.cboDispKey.Name = "cboDispKey"
        Me.cboDispKey.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDispKey.Size = New System.Drawing.Size(133, 20)
        Me.cboDispKey.TabIndex = 53
        '
        'lblDispFrame
        '
        Me.lblDispFrame.BackColor = System.Drawing.SystemColors.Control
        Me.lblDispFrame.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDispFrame.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDispFrame.Location = New System.Drawing.Point(4, 12)
        Me.lblDispFrame.Name = "lblDispFrame"
        Me.lblDispFrame.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDispFrame.Size = New System.Drawing.Size(72, 12)
        Me.lblDispFrame.TabIndex = 50
        Me.lblDispFrame.Text = "キー表示"
        Me.lblDispFrame.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblDispSC2P
        '
        Me.lblDispSC2P.BackColor = System.Drawing.SystemColors.Control
        Me.lblDispSC2P.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDispSC2P.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDispSC2P.Location = New System.Drawing.Point(133, 60)
        Me.lblDispSC2P.Name = "lblDispSC2P"
        Me.lblDispSC2P.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDispSC2P.Size = New System.Drawing.Size(27, 12)
        Me.lblDispSC2P.TabIndex = 56
        Me.lblDispSC2P.Text = "2P"
        Me.lblDispSC2P.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblDispSC1P
        '
        Me.lblDispSC1P.BackColor = System.Drawing.SystemColors.Control
        Me.lblDispSC1P.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDispSC1P.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDispSC1P.Location = New System.Drawing.Point(4, 60)
        Me.lblDispSC1P.Name = "lblDispSC1P"
        Me.lblDispSC1P.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDispSC1P.Size = New System.Drawing.Size(72, 12)
        Me.lblDispSC1P.TabIndex = 54
        Me.lblDispSC1P.Text = "スクラッチ1P"
        Me.lblDispSC1P.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblDispKey
        '
        Me.lblDispKey.BackColor = System.Drawing.SystemColors.Control
        Me.lblDispKey.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDispKey.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDispKey.Location = New System.Drawing.Point(4, 36)
        Me.lblDispKey.Name = "lblDispKey"
        Me.lblDispKey.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDispKey.Size = New System.Drawing.Size(72, 12)
        Me.lblDispKey.TabIndex = 52
        Me.lblDispKey.Text = "キー配置"
        Me.lblDispKey.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'fraMaterial
        '
        Me.fraMaterial.BackColor = System.Drawing.SystemColors.Control
        Me.fraMaterial.Controls.Add(Me._optChangeBottom_0)
        Me.fraMaterial.Controls.Add(Me._optChangeBottom_1)
        Me.fraMaterial.Controls.Add(Me._optChangeBottom_2)
        Me.fraMaterial.Controls.Add(Me._optChangeBottom_3)
        Me.fraMaterial.Controls.Add(Me._optChangeBottom_4)
        Me.fraMaterial.Controls.Add(Me._fraBottom_4)
        Me.fraMaterial.Controls.Add(Me._fraBottom_0)
        Me.fraMaterial.Controls.Add(Me._fraBottom_1)
        Me.fraMaterial.Controls.Add(Me._fraBottom_2)
        Me.fraMaterial.Controls.Add(Me._fraBottom_3)
        Me.fraMaterial.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraMaterial.Location = New System.Drawing.Point(140, 228)
        Me.fraMaterial.Name = "fraMaterial"
        Me.fraMaterial.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraMaterial.Size = New System.Drawing.Size(1105, 269)
        Me.fraMaterial.TabIndex = 58
        '
        '_optChangeBottom_0
        '
        Me._optChangeBottom_0.Appearance = System.Windows.Forms.Appearance.Button
        Me._optChangeBottom_0.BackColor = System.Drawing.SystemColors.Control
        Me._optChangeBottom_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optChangeBottom_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optChangeBottom_0.Location = New System.Drawing.Point(0, 0)
        Me._optChangeBottom_0.Name = "_optChangeBottom_0"
        Me._optChangeBottom_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optChangeBottom_0.Size = New System.Drawing.Size(61, 21)
        Me._optChangeBottom_0.TabIndex = 59
        Me._optChangeBottom_0.TabStop = True
        Me._optChangeBottom_0.Text = "#WAV"
        Me._optChangeBottom_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me._optChangeBottom_0.UseVisualStyleBackColor = False
        '
        '_optChangeBottom_1
        '
        Me._optChangeBottom_1.Appearance = System.Windows.Forms.Appearance.Button
        Me._optChangeBottom_1.BackColor = System.Drawing.SystemColors.Control
        Me._optChangeBottom_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optChangeBottom_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optChangeBottom_1.Location = New System.Drawing.Point(65, 0)
        Me._optChangeBottom_1.Name = "_optChangeBottom_1"
        Me._optChangeBottom_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optChangeBottom_1.Size = New System.Drawing.Size(61, 21)
        Me._optChangeBottom_1.TabIndex = 60
        Me._optChangeBottom_1.TabStop = True
        Me._optChangeBottom_1.Text = "#BMP"
        Me._optChangeBottom_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me._optChangeBottom_1.UseVisualStyleBackColor = False
        '
        '_optChangeBottom_2
        '
        Me._optChangeBottom_2.Appearance = System.Windows.Forms.Appearance.Button
        Me._optChangeBottom_2.BackColor = System.Drawing.SystemColors.Control
        Me._optChangeBottom_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._optChangeBottom_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optChangeBottom_2.Location = New System.Drawing.Point(130, 0)
        Me._optChangeBottom_2.Name = "_optChangeBottom_2"
        Me._optChangeBottom_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optChangeBottom_2.Size = New System.Drawing.Size(61, 21)
        Me._optChangeBottom_2.TabIndex = 61
        Me._optChangeBottom_2.TabStop = True
        Me._optChangeBottom_2.Text = "#BGA"
        Me._optChangeBottom_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me._optChangeBottom_2.UseVisualStyleBackColor = False
        '
        '_optChangeBottom_3
        '
        Me._optChangeBottom_3.Appearance = System.Windows.Forms.Appearance.Button
        Me._optChangeBottom_3.BackColor = System.Drawing.SystemColors.Control
        Me._optChangeBottom_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._optChangeBottom_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optChangeBottom_3.Location = New System.Drawing.Point(0, 25)
        Me._optChangeBottom_3.Name = "_optChangeBottom_3"
        Me._optChangeBottom_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optChangeBottom_3.Size = New System.Drawing.Size(61, 21)
        Me._optChangeBottom_3.TabIndex = 62
        Me._optChangeBottom_3.TabStop = True
        Me._optChangeBottom_3.Text = "拍子"
        Me._optChangeBottom_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me._optChangeBottom_3.UseVisualStyleBackColor = False
        '
        '_optChangeBottom_4
        '
        Me._optChangeBottom_4.Appearance = System.Windows.Forms.Appearance.Button
        Me._optChangeBottom_4.BackColor = System.Drawing.SystemColors.Control
        Me._optChangeBottom_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._optChangeBottom_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._optChangeBottom_4.Location = New System.Drawing.Point(65, 25)
        Me._optChangeBottom_4.Name = "_optChangeBottom_4"
        Me._optChangeBottom_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optChangeBottom_4.Size = New System.Drawing.Size(61, 21)
        Me._optChangeBottom_4.TabIndex = 63
        Me._optChangeBottom_4.TabStop = True
        Me._optChangeBottom_4.Text = "拡張命令"
        Me._optChangeBottom_4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me._optChangeBottom_4.UseVisualStyleBackColor = False
        '
        '_fraBottom_4
        '
        Me._fraBottom_4.BackColor = System.Drawing.SystemColors.Control
        Me._fraBottom_4.Controls.Add(Me.txtExInfo)
        Me._fraBottom_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraBottom_4.Location = New System.Drawing.Point(884, 32)
        Me._fraBottom_4.Name = "_fraBottom_4"
        Me._fraBottom_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraBottom_4.Size = New System.Drawing.Size(217, 233)
        Me._fraBottom_4.TabIndex = 93
        Me._fraBottom_4.Visible = False
        '
        'txtExInfo
        '
        Me.txtExInfo.AcceptsReturn = True
        Me.txtExInfo.BackColor = System.Drawing.SystemColors.Window
        Me.txtExInfo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtExInfo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtExInfo.Location = New System.Drawing.Point(0, 8)
        Me.txtExInfo.MaxLength = 0
        Me.txtExInfo.Multiline = True
        Me.txtExInfo.Name = "txtExInfo"
        Me.txtExInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtExInfo.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtExInfo.Size = New System.Drawing.Size(213, 217)
        Me.txtExInfo.TabIndex = 94
        Me.txtExInfo.WordWrap = False
        '
        '_fraBottom_0
        '
        Me._fraBottom_0.BackColor = System.Drawing.SystemColors.Control
        Me._fraBottom_0.Controls.Add(Me.cmdSoundExcUp)
        Me._fraBottom_0.Controls.Add(Me.cmdSoundExcDown)
        Me._fraBottom_0.Controls.Add(Me.cmdSoundDelete)
        Me._fraBottom_0.Controls.Add(Me.cmdSoundLoad)
        Me._fraBottom_0.Controls.Add(Me.cmdSoundStop)
        Me._fraBottom_0.Controls.Add(Me.lstWAV)
        Me._fraBottom_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraBottom_0.Location = New System.Drawing.Point(4, 56)
        Me._fraBottom_0.Name = "_fraBottom_0"
        Me._fraBottom_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraBottom_0.Size = New System.Drawing.Size(217, 209)
        Me._fraBottom_0.TabIndex = 64
        Me._fraBottom_0.Visible = False
        '
        'cmdSoundExcUp
        '
        Me.cmdSoundExcUp.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSoundExcUp.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSoundExcUp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSoundExcUp.Location = New System.Drawing.Point(61, 180)
        Me.cmdSoundExcUp.Name = "cmdSoundExcUp"
        Me.cmdSoundExcUp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSoundExcUp.Size = New System.Drawing.Size(21, 21)
        Me.cmdSoundExcUp.TabIndex = 67
        Me.cmdSoundExcUp.Text = "▲"
        Me.cmdSoundExcUp.UseVisualStyleBackColor = False
        '
        'cmdSoundExcDown
        '
        Me.cmdSoundExcDown.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSoundExcDown.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSoundExcDown.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSoundExcDown.Location = New System.Drawing.Point(86, 180)
        Me.cmdSoundExcDown.Name = "cmdSoundExcDown"
        Me.cmdSoundExcDown.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSoundExcDown.Size = New System.Drawing.Size(21, 21)
        Me.cmdSoundExcDown.TabIndex = 68
        Me.cmdSoundExcDown.Text = "▼"
        Me.cmdSoundExcDown.UseVisualStyleBackColor = False
        '
        'cmdSoundDelete
        '
        Me.cmdSoundDelete.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSoundDelete.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSoundDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSoundDelete.Location = New System.Drawing.Point(126, 180)
        Me.cmdSoundDelete.Name = "cmdSoundDelete"
        Me.cmdSoundDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSoundDelete.Size = New System.Drawing.Size(41, 21)
        Me.cmdSoundDelete.TabIndex = 69
        Me.cmdSoundDelete.Text = "消去"
        Me.cmdSoundDelete.UseVisualStyleBackColor = False
        '
        'cmdSoundLoad
        '
        Me.cmdSoundLoad.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSoundLoad.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSoundLoad.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSoundLoad.Location = New System.Drawing.Point(172, 180)
        Me.cmdSoundLoad.Name = "cmdSoundLoad"
        Me.cmdSoundLoad.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSoundLoad.Size = New System.Drawing.Size(41, 21)
        Me.cmdSoundLoad.TabIndex = 70
        Me.cmdSoundLoad.Text = "指定"
        Me.cmdSoundLoad.UseVisualStyleBackColor = False
        '
        'cmdSoundStop
        '
        Me.cmdSoundStop.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSoundStop.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSoundStop.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSoundStop.Location = New System.Drawing.Point(0, 180)
        Me.cmdSoundStop.Name = "cmdSoundStop"
        Me.cmdSoundStop.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSoundStop.Size = New System.Drawing.Size(53, 21)
        Me.cmdSoundStop.TabIndex = 66
        Me.cmdSoundStop.Text = "停止"
        Me.cmdSoundStop.UseVisualStyleBackColor = False
        '
        'lstWAV
        '
        Me.lstWAV.AllowDrop = True
        Me.lstWAV.BackColor = System.Drawing.SystemColors.Window
        Me.lstWAV.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstWAV.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstWAV.ItemHeight = 12
        Me.lstWAV.Location = New System.Drawing.Point(0, 8)
        Me.lstWAV.Name = "lstWAV"
        Me.lstWAV.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstWAV.Size = New System.Drawing.Size(213, 148)
        Me.lstWAV.TabIndex = 65
        '
        '_fraBottom_1
        '
        Me._fraBottom_1.BackColor = System.Drawing.SystemColors.Control
        Me._fraBottom_1.Controls.Add(Me.cmdBMPExcDown)
        Me._fraBottom_1.Controls.Add(Me.cmdBMPExcUp)
        Me._fraBottom_1.Controls.Add(Me.lstBMP)
        Me._fraBottom_1.Controls.Add(Me.cmdBMPDelete)
        Me._fraBottom_1.Controls.Add(Me.cmdBMPLoad)
        Me._fraBottom_1.Controls.Add(Me.cmdBMPPreview)
        Me._fraBottom_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraBottom_1.Location = New System.Drawing.Point(224, 32)
        Me._fraBottom_1.Name = "_fraBottom_1"
        Me._fraBottom_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraBottom_1.Size = New System.Drawing.Size(217, 233)
        Me._fraBottom_1.TabIndex = 71
        Me._fraBottom_1.Visible = False
        '
        'cmdBMPExcDown
        '
        Me.cmdBMPExcDown.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBMPExcDown.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBMPExcDown.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBMPExcDown.Location = New System.Drawing.Point(86, 204)
        Me.cmdBMPExcDown.Name = "cmdBMPExcDown"
        Me.cmdBMPExcDown.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBMPExcDown.Size = New System.Drawing.Size(21, 21)
        Me.cmdBMPExcDown.TabIndex = 75
        Me.cmdBMPExcDown.Text = "▼"
        Me.cmdBMPExcDown.UseVisualStyleBackColor = False
        '
        'cmdBMPExcUp
        '
        Me.cmdBMPExcUp.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBMPExcUp.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBMPExcUp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBMPExcUp.Location = New System.Drawing.Point(61, 204)
        Me.cmdBMPExcUp.Name = "cmdBMPExcUp"
        Me.cmdBMPExcUp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBMPExcUp.Size = New System.Drawing.Size(21, 21)
        Me.cmdBMPExcUp.TabIndex = 74
        Me.cmdBMPExcUp.Text = "▲"
        Me.cmdBMPExcUp.UseVisualStyleBackColor = False
        '
        'lstBMP
        '
        Me.lstBMP.AllowDrop = True
        Me.lstBMP.BackColor = System.Drawing.SystemColors.Window
        Me.lstBMP.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstBMP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstBMP.ItemHeight = 12
        Me.lstBMP.Location = New System.Drawing.Point(0, 8)
        Me.lstBMP.Name = "lstBMP"
        Me.lstBMP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstBMP.Size = New System.Drawing.Size(213, 148)
        Me.lstBMP.TabIndex = 72
        '
        'cmdBMPDelete
        '
        Me.cmdBMPDelete.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBMPDelete.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBMPDelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBMPDelete.Location = New System.Drawing.Point(126, 204)
        Me.cmdBMPDelete.Name = "cmdBMPDelete"
        Me.cmdBMPDelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBMPDelete.Size = New System.Drawing.Size(41, 21)
        Me.cmdBMPDelete.TabIndex = 76
        Me.cmdBMPDelete.Text = "消去"
        Me.cmdBMPDelete.UseVisualStyleBackColor = False
        '
        'cmdBMPLoad
        '
        Me.cmdBMPLoad.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBMPLoad.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBMPLoad.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBMPLoad.Location = New System.Drawing.Point(172, 204)
        Me.cmdBMPLoad.Name = "cmdBMPLoad"
        Me.cmdBMPLoad.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBMPLoad.Size = New System.Drawing.Size(41, 21)
        Me.cmdBMPLoad.TabIndex = 77
        Me.cmdBMPLoad.Text = "指定"
        Me.cmdBMPLoad.UseVisualStyleBackColor = False
        '
        'cmdBMPPreview
        '
        Me.cmdBMPPreview.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBMPPreview.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBMPPreview.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBMPPreview.Location = New System.Drawing.Point(0, 204)
        Me.cmdBMPPreview.Name = "cmdBMPPreview"
        Me.cmdBMPPreview.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBMPPreview.Size = New System.Drawing.Size(53, 21)
        Me.cmdBMPPreview.TabIndex = 73
        Me.cmdBMPPreview.Text = "表示"
        Me.cmdBMPPreview.UseVisualStyleBackColor = False
        '
        '_fraBottom_2
        '
        Me._fraBottom_2.BackColor = System.Drawing.SystemColors.Control
        Me._fraBottom_2.Controls.Add(Me.cmdBGAExcDown)
        Me._fraBottom_2.Controls.Add(Me.cmdBGAExcUp)
        Me._fraBottom_2.Controls.Add(Me.txtBGAInput)
        Me._fraBottom_2.Controls.Add(Me.cmdBGAPreview)
        Me._fraBottom_2.Controls.Add(Me.cmdBGASet)
        Me._fraBottom_2.Controls.Add(Me.cmdBGADelete)
        Me._fraBottom_2.Controls.Add(Me.lstBGA)
        Me._fraBottom_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraBottom_2.Location = New System.Drawing.Point(444, 32)
        Me._fraBottom_2.Name = "_fraBottom_2"
        Me._fraBottom_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraBottom_2.Size = New System.Drawing.Size(217, 233)
        Me._fraBottom_2.TabIndex = 78
        Me._fraBottom_2.Visible = False
        '
        'cmdBGAExcDown
        '
        Me.cmdBGAExcDown.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBGAExcDown.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBGAExcDown.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBGAExcDown.Location = New System.Drawing.Point(86, 204)
        Me.cmdBGAExcDown.Name = "cmdBGAExcDown"
        Me.cmdBGAExcDown.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBGAExcDown.Size = New System.Drawing.Size(21, 21)
        Me.cmdBGAExcDown.TabIndex = 83
        Me.cmdBGAExcDown.Text = "▼"
        Me.cmdBGAExcDown.UseVisualStyleBackColor = False
        '
        'cmdBGAExcUp
        '
        Me.cmdBGAExcUp.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBGAExcUp.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBGAExcUp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBGAExcUp.Location = New System.Drawing.Point(61, 204)
        Me.cmdBGAExcUp.Name = "cmdBGAExcUp"
        Me.cmdBGAExcUp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBGAExcUp.Size = New System.Drawing.Size(21, 21)
        Me.cmdBGAExcUp.TabIndex = 82
        Me.cmdBGAExcUp.Text = "▲"
        Me.cmdBGAExcUp.UseVisualStyleBackColor = False
        '
        'txtBGAInput
        '
        Me.txtBGAInput.AcceptsReturn = True
        Me.txtBGAInput.BackColor = System.Drawing.SystemColors.Window
        Me.txtBGAInput.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBGAInput.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBGAInput.Location = New System.Drawing.Point(0, 176)
        Me.txtBGAInput.MaxLength = 0
        Me.txtBGAInput.Name = "txtBGAInput"
        Me.txtBGAInput.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBGAInput.Size = New System.Drawing.Size(213, 19)
        Me.txtBGAInput.TabIndex = 80
        '
        'cmdBGAPreview
        '
        Me.cmdBGAPreview.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBGAPreview.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBGAPreview.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBGAPreview.Location = New System.Drawing.Point(0, 205)
        Me.cmdBGAPreview.Name = "cmdBGAPreview"
        Me.cmdBGAPreview.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBGAPreview.Size = New System.Drawing.Size(53, 21)
        Me.cmdBGAPreview.TabIndex = 81
        Me.cmdBGAPreview.Text = "表示"
        Me.cmdBGAPreview.UseVisualStyleBackColor = False
        '
        'cmdBGASet
        '
        Me.cmdBGASet.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBGASet.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBGASet.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBGASet.Location = New System.Drawing.Point(172, 204)
        Me.cmdBGASet.Name = "cmdBGASet"
        Me.cmdBGASet.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBGASet.Size = New System.Drawing.Size(41, 21)
        Me.cmdBGASet.TabIndex = 85
        Me.cmdBGASet.Text = "入力"
        Me.cmdBGASet.UseVisualStyleBackColor = False
        '
        'cmdBGADelete
        '
        Me.cmdBGADelete.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBGADelete.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBGADelete.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBGADelete.Location = New System.Drawing.Point(126, 204)
        Me.cmdBGADelete.Name = "cmdBGADelete"
        Me.cmdBGADelete.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBGADelete.Size = New System.Drawing.Size(41, 21)
        Me.cmdBGADelete.TabIndex = 84
        Me.cmdBGADelete.Text = "消去"
        Me.cmdBGADelete.UseVisualStyleBackColor = False
        '
        'lstBGA
        '
        Me.lstBGA.BackColor = System.Drawing.SystemColors.Window
        Me.lstBGA.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstBGA.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstBGA.ItemHeight = 12
        Me.lstBGA.Location = New System.Drawing.Point(0, 8)
        Me.lstBGA.Name = "lstBGA"
        Me.lstBGA.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstBGA.Size = New System.Drawing.Size(213, 148)
        Me.lstBGA.TabIndex = 79
        '
        '_fraBottom_3
        '
        Me._fraBottom_3.BackColor = System.Drawing.SystemColors.Control
        Me._fraBottom_3.Controls.Add(Me.cboNumerator)
        Me._fraBottom_3.Controls.Add(Me.cboDenominator)
        Me._fraBottom_3.Controls.Add(Me.cmdMeasureSelectAll)
        Me._fraBottom_3.Controls.Add(Me.cmdInputMeasureLen)
        Me._fraBottom_3.Controls.Add(Me.lstMeasureLen)
        Me._fraBottom_3.Controls.Add(Me.lblFraction)
        Me._fraBottom_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._fraBottom_3.Location = New System.Drawing.Point(664, 32)
        Me._fraBottom_3.Name = "_fraBottom_3"
        Me._fraBottom_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._fraBottom_3.Size = New System.Drawing.Size(217, 233)
        Me._fraBottom_3.TabIndex = 86
        Me._fraBottom_3.Visible = False
        '
        'cboNumerator
        '
        Me.cboNumerator.BackColor = System.Drawing.SystemColors.Window
        Me.cboNumerator.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboNumerator.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboNumerator.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboNumerator.Location = New System.Drawing.Point(75, 204)
        Me.cboNumerator.Name = "cboNumerator"
        Me.cboNumerator.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboNumerator.Size = New System.Drawing.Size(41, 20)
        Me.cboNumerator.TabIndex = 89
        '
        'cboDenominator
        '
        Me.cboDenominator.BackColor = System.Drawing.SystemColors.Window
        Me.cboDenominator.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboDenominator.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDenominator.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboDenominator.Items.AddRange(New Object() {"4", "6", "8", "12", "16", "24", "32", "48", "64", "96", "192", "384"})
        Me.cboDenominator.Location = New System.Drawing.Point(127, 204)
        Me.cboDenominator.Name = "cboDenominator"
        Me.cboDenominator.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboDenominator.Size = New System.Drawing.Size(41, 20)
        Me.cboDenominator.TabIndex = 91
        '
        'cmdMeasureSelectAll
        '
        Me.cmdMeasureSelectAll.BackColor = System.Drawing.SystemColors.Control
        Me.cmdMeasureSelectAll.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdMeasureSelectAll.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdMeasureSelectAll.Location = New System.Drawing.Point(0, 204)
        Me.cmdMeasureSelectAll.Name = "cmdMeasureSelectAll"
        Me.cmdMeasureSelectAll.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdMeasureSelectAll.Size = New System.Drawing.Size(53, 21)
        Me.cmdMeasureSelectAll.TabIndex = 88
        Me.cmdMeasureSelectAll.Text = "全選択"
        Me.cmdMeasureSelectAll.UseVisualStyleBackColor = False
        '
        'cmdInputMeasureLen
        '
        Me.cmdInputMeasureLen.BackColor = System.Drawing.SystemColors.Control
        Me.cmdInputMeasureLen.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdInputMeasureLen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdInputMeasureLen.Location = New System.Drawing.Point(172, 204)
        Me.cmdInputMeasureLen.Name = "cmdInputMeasureLen"
        Me.cmdInputMeasureLen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdInputMeasureLen.Size = New System.Drawing.Size(41, 21)
        Me.cmdInputMeasureLen.TabIndex = 92
        Me.cmdInputMeasureLen.Text = "入力"
        Me.cmdInputMeasureLen.UseVisualStyleBackColor = False
        '
        'lstMeasureLen
        '
        Me.lstMeasureLen.BackColor = System.Drawing.SystemColors.Window
        Me.lstMeasureLen.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstMeasureLen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstMeasureLen.ItemHeight = 12
        Me.lstMeasureLen.Location = New System.Drawing.Point(0, 8)
        Me.lstMeasureLen.Name = "lstMeasureLen"
        Me.lstMeasureLen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstMeasureLen.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lstMeasureLen.Size = New System.Drawing.Size(213, 148)
        Me.lstMeasureLen.TabIndex = 87
        '
        'lblFraction
        '
        Me.lblFraction.BackColor = System.Drawing.SystemColors.Control
        Me.lblFraction.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFraction.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFraction.Location = New System.Drawing.Point(119, 208)
        Me.lblFraction.Name = "lblFraction"
        Me.lblFraction.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFraction.Size = New System.Drawing.Size(6, 12)
        Me.lblFraction.TabIndex = 90
        Me.lblFraction.Text = "/"
        '
        'picMain
        '
        Me.picMain.AllowDrop = True
        Me.picMain.BackColor = System.Drawing.Color.Black
        Me.picMain.Cursor = System.Windows.Forms.Cursors.Default
        Me.picMain.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.picMain.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.picMain.Location = New System.Drawing.Point(0, 56)
        Me.picMain.Name = "picMain"
        Me.picMain.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.picMain.Size = New System.Drawing.Size(57, 33)
        Me.picMain.TabIndex = 0
        Me.picMain.TabStop = False
        '
        'hsbMain
        '
        Me.hsbMain.Cursor = System.Windows.Forms.Cursors.Default
        Me.hsbMain.LargeChange = 1
        Me.hsbMain.Location = New System.Drawing.Point(0, 456)
        Me.hsbMain.Maximum = 0
        Me.hsbMain.Name = "hsbMain"
        Me.hsbMain.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.hsbMain.Size = New System.Drawing.Size(113, 17)
        Me.hsbMain.TabIndex = 2
        '
        'vsbMain
        '
        Me.vsbMain.Cursor = System.Windows.Forms.Cursors.Default
        Me.vsbMain.LargeChange = 8
        Me.vsbMain.Location = New System.Drawing.Point(112, 60)
        Me.vsbMain.Maximum = 71
        Me.vsbMain.Name = "vsbMain"
        Me.vsbMain.Size = New System.Drawing.Size(17, 397)
        Me.vsbMain.TabIndex = 1
        Me.vsbMain.Value = 64
        '
        'dlgMainOpen
        '
        Me.dlgMainOpen.Filter = "BMS files (*.bms,*.bme,*.bml,*.pms)|*.bms;*.bme;*.bml;*.pms|All files (*.*)|*.*"
        '
        'dlgMainSave
        '
        Me.dlgMainSave.Filter = "BMS files (*.bms,*.bme,*.bml,*.pms)|*.bms;*.bme;*.bml;*.pms|All files (*.*)|*.*"
        '
        'tlbMenu
        '
        Me.tlbMenu.CanOverflow = False
        Me.tlbMenu.ImageList = Me.ilsMenu
        Me.tlbMenu.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.tlbMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me._New, Me.Open, Me.Reload, Me.Save, Me.SaveAs, Me.SepMode, Me.Edit, Me.Write, Me.Delete, Me.SepViewer, Me.Viewer, Me.PlayAll, Me.Play, Me._Stop, Me.SepGrid, Me.ChangeGrid, Me.SepSize, Me.DispSize, Me.SepResolution, Me.Resolution})
        Me.tlbMenu.Location = New System.Drawing.Point(0, 24)
        Me.tlbMenu.Name = "tlbMenu"
        Me.tlbMenu.Size = New System.Drawing.Size(1192, 25)
        Me.tlbMenu.TabIndex = 3
        '
        '_New
        '
        Me._New.AutoSize = False
        Me._New.ImageIndex = 0
        Me._New.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me._New.Name = "_New"
        Me._New.Size = New System.Drawing.Size(24, 22)
        Me._New.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me._New.ToolTipText = "新規作成"
        '
        'Open
        '
        Me.Open.AutoSize = False
        Me.Open.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem0, Me.ToolStripMenuItem1, Me.ToolStripMenuItem2, Me.ToolStripMenuItem3, Me.ToolStripMenuItem4, Me.ToolStripMenuItem5, Me.ToolStripMenuItem6, Me.ToolStripMenuItem7, Me.ToolStripMenuItem8, Me.ToolStripMenuItem9})
        Me.Open.ImageIndex = 1
        Me.Open.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.Open.Name = "Open"
        Me.Open.Size = New System.Drawing.Size(37, 22)
        Me.Open.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.Open.ToolTipText = "開く"
        '
        'ToolStripFileOpen
        '
        Me.ToolStripFileOpen.Name = "ToolStripFileOpen"
        Me.ToolStripFileOpen.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.ToolStripFileOpen.Size = New System.Drawing.Size(89, 22)
        Me.ToolStripFileOpen.Text = "Open ..."
        '
        'ToolStripMenuItem0
        '
        Me.ToolStripMenuItem0.Enabled = False
        Me.ToolStripMenuItem0.Name = "ToolStripMenuItem0"
        Me.ToolStripMenuItem0.Size = New System.Drawing.Size(89, 22)
        Me.ToolStripMenuItem0.Text = "&1:"
        Me.ToolStripMenuItem0.Visible = False
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Enabled = False
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(89, 22)
        Me.ToolStripMenuItem1.Text = "&2:"
        Me.ToolStripMenuItem1.Visible = False
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Enabled = False
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(89, 22)
        Me.ToolStripMenuItem2.Text = "&3:"
        Me.ToolStripMenuItem2.Visible = False
        '
        'ToolStripMenuItem3
        '
        Me.ToolStripMenuItem3.Enabled = False
        Me.ToolStripMenuItem3.Name = "ToolStripMenuItem3"
        Me.ToolStripMenuItem3.Size = New System.Drawing.Size(89, 22)
        Me.ToolStripMenuItem3.Text = "&4:"
        Me.ToolStripMenuItem3.Visible = False
        '
        'ToolStripMenuItem4
        '
        Me.ToolStripMenuItem4.Enabled = False
        Me.ToolStripMenuItem4.Name = "ToolStripMenuItem4"
        Me.ToolStripMenuItem4.Size = New System.Drawing.Size(89, 22)
        Me.ToolStripMenuItem4.Text = "&5:"
        Me.ToolStripMenuItem4.Visible = False
        '
        'ToolStripMenuItem5
        '
        Me.ToolStripMenuItem5.Enabled = False
        Me.ToolStripMenuItem5.Name = "ToolStripMenuItem5"
        Me.ToolStripMenuItem5.Size = New System.Drawing.Size(89, 22)
        Me.ToolStripMenuItem5.Text = "&6:"
        Me.ToolStripMenuItem5.Visible = False
        '
        'ToolStripMenuItem6
        '
        Me.ToolStripMenuItem6.Enabled = False
        Me.ToolStripMenuItem6.Name = "ToolStripMenuItem6"
        Me.ToolStripMenuItem6.Size = New System.Drawing.Size(89, 22)
        Me.ToolStripMenuItem6.Text = "&7:"
        Me.ToolStripMenuItem6.Visible = False
        '
        'ToolStripMenuItem7
        '
        Me.ToolStripMenuItem7.Enabled = False
        Me.ToolStripMenuItem7.Name = "ToolStripMenuItem7"
        Me.ToolStripMenuItem7.Size = New System.Drawing.Size(89, 22)
        Me.ToolStripMenuItem7.Text = "&8:"
        Me.ToolStripMenuItem7.Visible = False
        '
        'ToolStripMenuItem8
        '
        Me.ToolStripMenuItem8.Enabled = False
        Me.ToolStripMenuItem8.Name = "ToolStripMenuItem8"
        Me.ToolStripMenuItem8.Size = New System.Drawing.Size(89, 22)
        Me.ToolStripMenuItem8.Text = "&9:"
        Me.ToolStripMenuItem8.Visible = False
        '
        'ToolStripMenuItem9
        '
        Me.ToolStripMenuItem9.Enabled = False
        Me.ToolStripMenuItem9.Name = "ToolStripMenuItem9"
        Me.ToolStripMenuItem9.Size = New System.Drawing.Size(89, 22)
        Me.ToolStripMenuItem9.Text = "&10:"
        Me.ToolStripMenuItem9.Visible = False
        '
        'Reload
        '
        Me.Reload.AutoSize = False
        Me.Reload.ImageIndex = 2
        Me.Reload.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.Reload.Name = "Reload"
        Me.Reload.Size = New System.Drawing.Size(24, 22)
        Me.Reload.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.Reload.ToolTipText = "再読み込み"
        '
        'Save
        '
        Me.Save.AutoSize = False
        Me.Save.ImageIndex = 3
        Me.Save.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.Save.Name = "Save"
        Me.Save.Size = New System.Drawing.Size(24, 22)
        Me.Save.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.Save.ToolTipText = "上書き保存"
        '
        'SaveAs
        '
        Me.SaveAs.AutoSize = False
        Me.SaveAs.ImageIndex = 4
        Me.SaveAs.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.SaveAs.Name = "SaveAs"
        Me.SaveAs.Size = New System.Drawing.Size(24, 22)
        Me.SaveAs.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.SaveAs.ToolTipText = "名前を付けて保存"
        '
        'SepMode
        '
        Me.SepMode.AutoSize = False
        Me.SepMode.Name = "SepMode"
        Me.SepMode.Size = New System.Drawing.Size(24, 22)
        '
        'Edit
        '
        Me.Edit.AutoSize = False
        Me.Edit.Checked = True
        Me.Edit.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Edit.ImageIndex = 5
        Me.Edit.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.Edit.Name = "Edit"
        Me.Edit.Size = New System.Drawing.Size(24, 22)
        Me.Edit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.Edit.ToolTipText = "編集"
        '
        'Write
        '
        Me.Write.AutoSize = False
        Me.Write.ImageIndex = 6
        Me.Write.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.Write.Name = "Write"
        Me.Write.Size = New System.Drawing.Size(24, 22)
        Me.Write.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.Write.ToolTipText = "書込"
        '
        'Delete
        '
        Me.Delete.AutoSize = False
        Me.Delete.ImageIndex = 7
        Me.Delete.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.Delete.Name = "Delete"
        Me.Delete.Size = New System.Drawing.Size(24, 22)
        Me.Delete.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.Delete.ToolTipText = "消去"
        '
        'SepViewer
        '
        Me.SepViewer.AutoSize = False
        Me.SepViewer.Name = "SepViewer"
        Me.SepViewer.Size = New System.Drawing.Size(24, 22)
        '
        'Viewer
        '
        Me.Viewer.AutoSize = False
        Me.Viewer.Name = "Viewer"
        Me.Viewer.Size = New System.Drawing.Size(93, 22)
        '
        'PlayAll
        '
        Me.PlayAll.AutoSize = False
        Me.PlayAll.ImageIndex = 8
        Me.PlayAll.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.PlayAll.Name = "PlayAll"
        Me.PlayAll.Size = New System.Drawing.Size(24, 22)
        Me.PlayAll.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.PlayAll.ToolTipText = "最初から再生"
        '
        'Play
        '
        Me.Play.AutoSize = False
        Me.Play.ImageIndex = 9
        Me.Play.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.Play.Name = "Play"
        Me.Play.Size = New System.Drawing.Size(24, 22)
        Me.Play.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.Play.ToolTipText = "現在位置から再生"
        '
        '_Stop
        '
        Me._Stop.AutoSize = False
        Me._Stop.ImageIndex = 10
        Me._Stop.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me._Stop.Name = "_Stop"
        Me._Stop.Size = New System.Drawing.Size(24, 22)
        Me._Stop.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me._Stop.ToolTipText = "停止"
        '
        'SepGrid
        '
        Me.SepGrid.AutoSize = False
        Me.SepGrid.Name = "SepGrid"
        Me.SepGrid.Size = New System.Drawing.Size(24, 22)
        '
        'ChangeGrid
        '
        Me.ChangeGrid.AutoSize = False
        Me.ChangeGrid.Name = "ChangeGrid"
        Me.ChangeGrid.Size = New System.Drawing.Size(197, 22)
        '
        'SepSize
        '
        Me.SepSize.AutoSize = False
        Me.SepSize.Name = "SepSize"
        Me.SepSize.Size = New System.Drawing.Size(24, 22)
        '
        'DispSize
        '
        Me.DispSize.AutoSize = False
        Me.DispSize.Name = "DispSize"
        Me.DispSize.Size = New System.Drawing.Size(197, 22)
        '
        'SepResolution
        '
        Me.SepResolution.AutoSize = False
        Me.SepResolution.Name = "SepResolution"
        Me.SepResolution.Size = New System.Drawing.Size(24, 22)
        '
        'Resolution
        '
        Me.Resolution.AutoSize = False
        Me.Resolution.Name = "Resolution"
        Me.Resolution.Size = New System.Drawing.Size(137, 22)
        '
        'lblDirectInput
        '
        Me.lblDirectInput.AutoSize = True
        Me.lblDirectInput.BackColor = System.Drawing.SystemColors.Control
        Me.lblDirectInput.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDirectInput.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDirectInput.Location = New System.Drawing.Point(4, 482)
        Me.lblDirectInput.Name = "lblDirectInput"
        Me.lblDirectInput.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDirectInput.Size = New System.Drawing.Size(41, 12)
        Me.lblDirectInput.TabIndex = 95
        Me.lblDirectInput.Text = "Direct"
        '
        'frmMain
        '
        Me.AllowDrop = True
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1192, 530)
        Me.Controls.Add(Me.staMain)
        Me.Controls.Add(Me.picSiromaru)
        Me.Controls.Add(Me.fraResolution)
        Me.Controls.Add(Me.fraDispSize)
        Me.Controls.Add(Me.fraViewer)
        Me.Controls.Add(Me.cboDirectInput)
        Me.Controls.Add(Me.cmdDirectInput)
        Me.Controls.Add(Me.fraGrid)
        Me.Controls.Add(Me.fraHeader)
        Me.Controls.Add(Me.fraMaterial)
        Me.Controls.Add(Me.picMain)
        Me.Controls.Add(Me.hsbMain)
        Me.Controls.Add(Me.vsbMain)
        Me.Controls.Add(Me.tlbMenu)
        Me.Controls.Add(Me.lblDirectInput)
        Me.Controls.Add(Me.MainMenu1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(11, 49)
        Me.Name = "frmMain"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "BMx Sequence Editor"
        Me.MainMenu1.ResumeLayout(False)
        Me.MainMenu1.PerformLayout()
        Me.mnuContext.ResumeLayout(False)
        Me.mnuContextList.ResumeLayout(False)
        Me.staMain.ResumeLayout(False)
        Me.staMain.PerformLayout()
        CType(Me.picSiromaru, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraResolution.ResumeLayout(False)
        Me.fraResolution.PerformLayout()
        Me.fraDispSize.ResumeLayout(False)
        Me.fraDispSize.PerformLayout()
        Me.fraViewer.ResumeLayout(False)
        Me.fraGrid.ResumeLayout(False)
        Me.fraGrid.PerformLayout()
        Me.fraHeader.ResumeLayout(False)
        Me._fraTop_0.ResumeLayout(False)
        Me._fraTop_0.PerformLayout()
        Me._fraTop_1.ResumeLayout(False)
        Me._fraTop_1.PerformLayout()
        Me._fraTop_2.ResumeLayout(False)
        Me._fraTop_2.PerformLayout()
        Me.fraMaterial.ResumeLayout(False)
        Me._fraBottom_4.ResumeLayout(False)
        Me._fraBottom_4.PerformLayout()
        Me._fraBottom_0.ResumeLayout(False)
        Me._fraBottom_1.ResumeLayout(False)
        Me._fraBottom_2.ResumeLayout(False)
        Me._fraBottom_2.PerformLayout()
        Me._fraBottom_3.ResumeLayout(False)
        CType(Me.picMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tlbMenu.ResumeLayout(False)
        Me.tlbMenu.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ToolStripFileOpen As ToolStripMenuItem
    Friend WithEvents ToolStripLineFile As ToolStripSeparator
    Friend WithEvents ToolStripMenuItem0 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem3 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem4 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem5 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem6 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem7 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem8 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem9 As ToolStripMenuItem
#End Region
End Class
