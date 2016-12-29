<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmZaan
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmZaan))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.tcCubes = New System.Windows.Forms.TabControl()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.pnlCube = New System.Windows.Forms.Panel()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.trvW = New System.Windows.Forms.TreeView()
        Me.cmsTrvW = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tsmiTrvWAdd = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiTrvWAddExternal = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssTrvW1 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiTrvWRefresh = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssTrvW2 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiTrvWRename = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiTrvWDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssTrvW3 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiTrvWImport = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiTrvWExport = New System.Windows.Forms.ToolStripMenuItem()
        Me.imgFileTypes = New System.Windows.Forms.ImageList(Me.components)
        Me.lvBookmark = New System.Windows.Forms.ListView()
        Me.cmsBookmark = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tsmiSelTabsAdd = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiSelTabsRefresh = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiSelTabsDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssSelTabs1 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiSelTabsDefault = New System.Windows.Forms.ToolStripMenuItem()
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer()
        Me.pnlList = New System.Windows.Forms.Panel()
        Me.pgbZaan = New System.Windows.Forms.ProgressBar()
        Me.lvIn = New System.Windows.Forms.ListView()
        Me.cmsLvIn = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tsmiLvInDocPerPage = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInDpp10 = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInDpp25 = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInDpp50 = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInDpp100 = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInDpp250 = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInDpp500 = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInDpp1000 = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInImageMode = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssLvIn1 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvInRefresh = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInSelAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInUndoMove = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssLvIn2 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvInCut = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInCopy = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInPaste = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssLvIn3 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvInCopyPath = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInOpenFolder = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssLvIn4 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvInAutoRename = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInRename = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssLvIn5 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvInNew = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInNewNote = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInNewText = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInNewPres = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInNewSpSh = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssLvIn6 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvInCopyToZC = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInMoveOut = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssLvIn7 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvInExport = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInExpNameTable = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvInExpRefTable = New System.Windows.Forms.ToolStripMenuItem()
        Me.imgLargeIcons = New System.Windows.Forms.ImageList(Me.components)
        Me.lvMatrix = New System.Windows.Forms.ListView()
        Me.lblDocName = New System.Windows.Forms.Label()
        Me.lstPage = New System.Windows.Forms.ListBox()
        Me.lstInDir = New System.Windows.Forms.ListBox()
        Me.pctLvIn = New System.Windows.Forms.PictureBox()
        Me.pnlZoomVideo = New System.Windows.Forms.Panel()
        Me.pnlVideoControl = New System.Windows.Forms.Panel()
        Me.pgbVideo = New System.Windows.Forms.ProgressBar()
        Me.btnVideoBegin = New System.Windows.Forms.Button()
        Me.btnVideoEnd = New System.Windows.Forms.Button()
        Me.btnVideoPlayPause = New System.Windows.Forms.Button()
        Me.btnVideoStop = New System.Windows.Forms.Button()
        Me.wbDoc = New System.Windows.Forms.WebBrowser()
        Me.pnlZoom = New System.Windows.Forms.Panel()
        Me.pctZoom = New System.Windows.Forms.PictureBox()
        Me.cmsPctZoom = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tsmiPctZoomSlideShow = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssPctZoom1 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiPctZoomInterval = New System.Windows.Forms.ToolStripMenuItem()
        Me.tstbPctZoom = New System.Windows.Forms.ToolStripTextBox()
        Me.pnlSlideControl = New System.Windows.Forms.Panel()
        Me.lblZoomName = New System.Windows.Forms.Label()
        Me.btnSlideRotateLeft = New System.Windows.Forms.Button()
        Me.btnSlideRotateRight = New System.Windows.Forms.Button()
        Me.btnSlideDelete = New System.Windows.Forms.Button()
        Me.lblSlideNb = New System.Windows.Forms.Label()
        Me.btnSlidePrevious = New System.Windows.Forms.Button()
        Me.btnSlideNext = New System.Windows.Forms.Button()
        Me.btnSlidePlayPause = New System.Windows.Forms.Button()
        Me.btnSlideStop = New System.Windows.Forms.Button()
        Me.pctZoom2 = New System.Windows.Forms.PictureBox()
        Me.pnlListTop = New System.Windows.Forms.Panel()
        Me.pnlSelectorNav = New System.Windows.Forms.Panel()
        Me.pnlSelectorHeaderBottom = New System.Windows.Forms.Panel()
        Me.btnWho2Axis = New System.Windows.Forms.Button()
        Me.btnWhat2Axis = New System.Windows.Forms.Button()
        Me.btnWhereAxis = New System.Windows.Forms.Button()
        Me.btnWhatAxis = New System.Windows.Forms.Button()
        Me.btnPanelView = New System.Windows.Forms.Button()
        Me.btnWhoAxis = New System.Windows.Forms.Button()
        Me.btnWhenAxis = New System.Windows.Forms.Button()
        Me.btnToday = New System.Windows.Forms.Button()
        Me.btnDataAccessBlank = New System.Windows.Forms.Button()
        Me.btnDataAccessAxis = New System.Windows.Forms.Button()
        Me.pnlSelectorHeaderTop = New System.Windows.Forms.Panel()
        Me.btnCubeTube = New System.Windows.Forms.Button()
        Me.btnPanelImport = New System.Windows.Forms.Button()
        Me.btnWho2 = New System.Windows.Forms.Button()
        Me.btnWho2Root = New System.Windows.Forms.Button()
        Me.btnWhat2 = New System.Windows.Forms.Button()
        Me.btnWhat2Root = New System.Windows.Forms.Button()
        Me.btnWhere = New System.Windows.Forms.Button()
        Me.btnWhereRoot = New System.Windows.Forms.Button()
        Me.btnWhat = New System.Windows.Forms.Button()
        Me.btnWhatRoot = New System.Windows.Forms.Button()
        Me.btnWho = New System.Windows.Forms.Button()
        Me.btnWhoRoot = New System.Windows.Forms.Button()
        Me.btnWhen = New System.Windows.Forms.Button()
        Me.btnWhenRoot = New System.Windows.Forms.Button()
        Me.btnDataAccess = New System.Windows.Forms.Button()
        Me.btnDataAccessRoot = New System.Windows.Forms.Button()
        Me.pnlSelectorHeader = New System.Windows.Forms.Panel()
        Me.lblEmpty = New System.Windows.Forms.Label()
        Me.lblAboutZaan = New System.Windows.Forms.Label()
        Me.lblWho2 = New System.Windows.Forms.Label()
        Me.lblWhat2 = New System.Windows.Forms.Label()
        Me.lblWhere = New System.Windows.Forms.Label()
        Me.lblWhat = New System.Windows.Forms.Label()
        Me.lblWho = New System.Windows.Forms.Label()
        Me.lblWhen = New System.Windows.Forms.Label()
        Me.lblDataAccess = New System.Windows.Forms.Label()
        Me.lvSelector = New System.Windows.Forms.ListView()
        Me.pnlSelectorSearch = New System.Windows.Forms.Panel()
        Me.pnlSelectorTop = New System.Windows.Forms.Panel()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.btnPrev = New System.Windows.Forms.Button()
        Me.pnlSearch = New System.Windows.Forms.Panel()
        Me.tbSearch = New System.Windows.Forms.TextBox()
        Me.lblSearchDoc = New System.Windows.Forms.Label()
        Me.lblSearchFolder = New System.Windows.Forms.Label()
        Me.lblSelectorReset = New System.Windows.Forms.Label()
        Me.lblBookmark = New System.Windows.Forms.Label()
        Me.pnlSelectorBottom = New System.Windows.Forms.Panel()
        Me.btnPanelCube = New System.Windows.Forms.Button()
        Me.lblSelPagePrev = New System.Windows.Forms.Label()
        Me.lblSelPage = New System.Windows.Forms.Label()
        Me.lblSelPageNext = New System.Windows.Forms.Label()
        Me.btnPanelEmpty = New System.Windows.Forms.Button()
        Me.btnPanelTree = New System.Windows.Forms.Button()
        Me.btnMatrix = New System.Windows.Forms.Button()
        Me.pnlDatabase = New System.Windows.Forms.Panel()
        Me.pctZaanLogo = New System.Windows.Forms.PictureBox()
        Me.cmsSelector = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tsmiSelectorTreeView = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiSelectorViewer = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiSelectorImport = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssSelector0 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiSelectorTreeLock = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssSelector1 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiSelectorNightMode = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssSelector2 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiSelectorChangeDb = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiSelectorChangeDbImage = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssSelector3 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiSelectorCheckDb = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssSelector4 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiSelectorExportDBwin = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiSelectorImportDBwin = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiSelectorAutoImport = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssSelector5 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiSelectorCloseDb = New System.Windows.Forms.ToolStripMenuItem()
        Me.tcFolders = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.lvOut = New System.Windows.Forms.ListView()
        Me.cmsLvOut = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tsmiLvOutChangeFolder = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvOutFoldersVisible = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssLvOut1 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvOutRefresh = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvOutSelAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvOutUndoMove = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssLvOut2 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvOutCut = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvOutCopy = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvOutPaste = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssLvOut3 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvOutRename = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvOutDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssLvOut4 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvOutMoveIn = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvOutAutoFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssLvOut5 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvOutImport = New System.Windows.Forms.ToolStripMenuItem()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer()
        Me.trvInput = New System.Windows.Forms.TreeView()
        Me.cmsTrvInput = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tsmiTrvInputChangeRoot = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssTrvInput0 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiTrvInputRefresh = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssTrvInput1 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiTrvInputDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.lvTemp = New System.Windows.Forms.ListView()
        Me.cmsLvTemp = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tsmiLvTempRefresh = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvTempSelAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvTempUndoMove = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssLvTemp1 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvTempCut = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvTempCopy = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvTempPaste = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssLvTemp2 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvTempResizePicture = New System.Windows.Forms.ToolStripMenuItem()
        Me.tstLvTempResizePercentage = New System.Windows.Forms.ToolStripTextBox()
        Me.tssLvTemp3 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvTempRename = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsmiLvTempDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssLvTemp4 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvTempSend = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssLvTemp5 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvTempImport = New System.Windows.Forms.ToolStripMenuItem()
        Me.imgPanels = New System.Windows.Forms.ImageList(Me.components)
        Me.tlTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.fbdZaanPath = New System.Windows.Forms.FolderBrowserDialog()
        Me.tmrLvTempLostFocus = New System.Windows.Forms.Timer(Me.components)
        Me.tmrLicenseCheck = New System.Windows.Forms.Timer(Me.components)
        Me.fswData = New System.IO.FileSystemWatcher()
        Me.tmrLvInDisplay = New System.Windows.Forms.Timer(Me.components)
        Me.tmrTrvWDisplay = New System.Windows.Forms.Timer(Me.components)
        Me.fswXin = New System.IO.FileSystemWatcher()
        Me.fswTree = New System.IO.FileSystemWatcher()
        Me.tmrCubeImport = New System.Windows.Forms.Timer(Me.components)
        Me.fswInput = New System.IO.FileSystemWatcher()
        Me.tmrSlideShow = New System.Windows.Forms.Timer(Me.components)
        Me.tmrSlideShowEffect = New System.Windows.Forms.Timer(Me.components)
        Me.tmrVideoProgress = New System.Windows.Forms.Timer(Me.components)
        Me.fswZaan = New System.IO.FileSystemWatcher()
        Me.cmsLvInWhos = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tsmiLvInWhosMultiple = New System.Windows.Forms.ToolStripMenuItem()
        Me.tssLvInWhos0 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsmiLvInWhosDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.tscbLvInWhosList = New System.Windows.Forms.ToolStripComboBox()
        Me.tmrSelectorList = New System.Windows.Forms.Timer(Me.components)
        Me.ofdTxtImport = New System.Windows.Forms.OpenFileDialog()
        Me.ofdZaanFile = New System.Windows.Forms.OpenFileDialog()
        Me.cmsSelChildren = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Child31ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Child32ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tmrSelChildren = New System.Windows.Forms.Timer(Me.components)
        Me.cmsSelBookmarks = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.tmrDbExport = New System.Windows.Forms.Timer(Me.components)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.tcCubes.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.pnlCube.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.cmsTrvW.SuspendLayout()
        Me.cmsBookmark.SuspendLayout()
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.pnlList.SuspendLayout()
        Me.cmsLvIn.SuspendLayout()
        CType(Me.pctLvIn, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlZoomVideo.SuspendLayout()
        Me.pnlVideoControl.SuspendLayout()
        Me.pnlZoom.SuspendLayout()
        CType(Me.pctZoom, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cmsPctZoom.SuspendLayout()
        Me.pnlSlideControl.SuspendLayout()
        CType(Me.pctZoom2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlListTop.SuspendLayout()
        Me.pnlSelectorNav.SuspendLayout()
        Me.pnlSelectorHeaderBottom.SuspendLayout()
        Me.pnlSelectorHeaderTop.SuspendLayout()
        Me.pnlSelectorHeader.SuspendLayout()
        Me.pnlSelectorSearch.SuspendLayout()
        Me.pnlSelectorTop.SuspendLayout()
        Me.pnlSearch.SuspendLayout()
        Me.pnlSelectorBottom.SuspendLayout()
        Me.pnlDatabase.SuspendLayout()
        CType(Me.pctZaanLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cmsSelector.SuspendLayout()
        Me.tcFolders.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.cmsLvOut.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.cmsTrvInput.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.cmsLvTemp.SuspendLayout()
        CType(Me.fswData, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fswXin, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fswTree, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fswInput, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fswZaan, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.cmsLvInWhos.SuspendLayout()
        Me.cmsSelChildren.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.BackColor = System.Drawing.SystemColors.ControlDark
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(10, 10)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(6)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.tcCubes)
        Me.SplitContainer1.Panel1MinSize = 300
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.tcFolders)
        Me.SplitContainer1.Size = New System.Drawing.Size(2332, 1022)
        Me.SplitContainer1.SplitterDistance = 387
        Me.SplitContainer1.SplitterWidth = 8
        Me.SplitContainer1.TabIndex = 0
        '
        'tcCubes
        '
        Me.tcCubes.Controls.Add(Me.TabPage4)
        Me.tcCubes.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tcCubes.Font = New System.Drawing.Font("Franklin Gothic Book", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tcCubes.ImageList = Me.imgFileTypes
        Me.tcCubes.Location = New System.Drawing.Point(0, 0)
        Me.tcCubes.Margin = New System.Windows.Forms.Padding(6)
        Me.tcCubes.Name = "tcCubes"
        Me.tcCubes.SelectedIndex = 0
        Me.tcCubes.ShowToolTips = True
        Me.tcCubes.Size = New System.Drawing.Size(2332, 387)
        Me.tcCubes.TabIndex = 19
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.pnlCube)
        Me.TabPage4.Location = New System.Drawing.Point(8, 52)
        Me.TabPage4.Margin = New System.Windows.Forms.Padding(6)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(6)
        Me.TabPage4.Size = New System.Drawing.Size(2316, 327)
        Me.TabPage4.TabIndex = 0
        Me.TabPage4.Text = "..."
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'pnlCube
        '
        Me.pnlCube.Controls.Add(Me.SplitContainer2)
        Me.pnlCube.Controls.Add(Me.pnlListTop)
        Me.pnlCube.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlCube.Location = New System.Drawing.Point(6, 6)
        Me.pnlCube.Margin = New System.Windows.Forms.Padding(6)
        Me.pnlCube.Name = "pnlCube"
        Me.pnlCube.Size = New System.Drawing.Size(2304, 315)
        Me.pnlCube.TabIndex = 20
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 137)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(6)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.trvW)
        Me.SplitContainer2.Panel1.Controls.Add(Me.lvBookmark)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer3)
        Me.SplitContainer2.Size = New System.Drawing.Size(2304, 178)
        Me.SplitContainer2.SplitterDistance = 274
        Me.SplitContainer2.SplitterWidth = 8
        Me.SplitContainer2.TabIndex = 1
        '
        'trvW
        '
        Me.trvW.AllowDrop = True
        Me.trvW.BackColor = System.Drawing.SystemColors.Control
        Me.trvW.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.trvW.ContextMenuStrip = Me.cmsTrvW
        Me.trvW.Dock = System.Windows.Forms.DockStyle.Fill
        Me.trvW.Font = New System.Drawing.Font("Arial Rounded MT Bold", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.trvW.ForeColor = System.Drawing.Color.MidnightBlue
        Me.trvW.HotTracking = True
        Me.trvW.ImageIndex = 0
        Me.trvW.ImageList = Me.imgFileTypes
        Me.trvW.Indent = 16
        Me.trvW.ItemHeight = 18
        Me.trvW.LabelEdit = True
        Me.trvW.Location = New System.Drawing.Point(0, 0)
        Me.trvW.Margin = New System.Windows.Forms.Padding(6)
        Me.trvW.Name = "trvW"
        Me.trvW.SelectedImageIndex = 0
        Me.trvW.ShowNodeToolTips = True
        Me.trvW.Size = New System.Drawing.Size(274, 178)
        Me.trvW.TabIndex = 6
        '
        'cmsTrvW
        '
        Me.cmsTrvW.BackColor = System.Drawing.SystemColors.Menu
        Me.cmsTrvW.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.cmsTrvW.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiTrvWAdd, Me.tsmiTrvWAddExternal, Me.tssTrvW1, Me.tsmiTrvWRefresh, Me.tssTrvW2, Me.tsmiTrvWRename, Me.tsmiTrvWDelete, Me.tssTrvW3, Me.tsmiTrvWImport, Me.tsmiTrvWExport})
        Me.cmsTrvW.Name = "cmsTrvW"
        Me.cmsTrvW.Size = New System.Drawing.Size(341, 288)
        '
        'tsmiTrvWAdd
        '
        Me.tsmiTrvWAdd.Image = Global.WindowsApplication1.My.Resources.Resources.add_lg
        Me.tsmiTrvWAdd.Name = "tsmiTrvWAdd"
        Me.tsmiTrvWAdd.Size = New System.Drawing.Size(340, 38)
        Me.tsmiTrvWAdd.Text = "Add"
        '
        'tsmiTrvWAddExternal
        '
        Me.tsmiTrvWAddExternal.Image = Global.WindowsApplication1.My.Resources.Resources.add_lg
        Me.tsmiTrvWAddExternal.Name = "tsmiTrvWAddExternal"
        Me.tsmiTrvWAddExternal.Size = New System.Drawing.Size(340, 38)
        Me.tsmiTrvWAddExternal.Text = "Add as external folder"
        '
        'tssTrvW1
        '
        Me.tssTrvW1.Name = "tssTrvW1"
        Me.tssTrvW1.Size = New System.Drawing.Size(337, 6)
        '
        'tsmiTrvWRefresh
        '
        Me.tsmiTrvWRefresh.Name = "tsmiTrvWRefresh"
        Me.tsmiTrvWRefresh.Size = New System.Drawing.Size(340, 38)
        Me.tsmiTrvWRefresh.Text = "Refresh"
        '
        'tssTrvW2
        '
        Me.tssTrvW2.Name = "tssTrvW2"
        Me.tssTrvW2.Size = New System.Drawing.Size(337, 6)
        '
        'tsmiTrvWRename
        '
        Me.tsmiTrvWRename.Name = "tsmiTrvWRename"
        Me.tsmiTrvWRename.Size = New System.Drawing.Size(340, 38)
        Me.tsmiTrvWRename.Text = "Rename"
        '
        'tsmiTrvWDelete
        '
        Me.tsmiTrvWDelete.Enabled = False
        Me.tsmiTrvWDelete.Image = Global.WindowsApplication1.My.Resources.Resources.delete_lg
        Me.tsmiTrvWDelete.Name = "tsmiTrvWDelete"
        Me.tsmiTrvWDelete.Size = New System.Drawing.Size(340, 38)
        Me.tsmiTrvWDelete.Text = "Delete"
        '
        'tssTrvW3
        '
        Me.tssTrvW3.Name = "tssTrvW3"
        Me.tssTrvW3.Size = New System.Drawing.Size(337, 6)
        '
        'tsmiTrvWImport
        '
        Me.tsmiTrvWImport.Name = "tsmiTrvWImport"
        Me.tsmiTrvWImport.Size = New System.Drawing.Size(340, 38)
        Me.tsmiTrvWImport.Text = "Import tree view..."
        '
        'tsmiTrvWExport
        '
        Me.tsmiTrvWExport.Name = "tsmiTrvWExport"
        Me.tsmiTrvWExport.Size = New System.Drawing.Size(340, 38)
        Me.tsmiTrvWExport.Text = "Export tree view"
        '
        'imgFileTypes
        '
        Me.imgFileTypes.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit
        Me.imgFileTypes.ImageSize = New System.Drawing.Size(16, 16)
        Me.imgFileTypes.TransparentColor = System.Drawing.Color.Transparent
        '
        'lvBookmark
        '
        Me.lvBookmark.AllowDrop = True
        Me.lvBookmark.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lvBookmark.ContextMenuStrip = Me.cmsBookmark
        Me.lvBookmark.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lvBookmark.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvBookmark.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvBookmark.ForeColor = System.Drawing.Color.MidnightBlue
        Me.lvBookmark.Location = New System.Drawing.Point(0, 0)
        Me.lvBookmark.Margin = New System.Windows.Forms.Padding(6)
        Me.lvBookmark.MultiSelect = False
        Me.lvBookmark.Name = "lvBookmark"
        Me.lvBookmark.Scrollable = False
        Me.lvBookmark.Size = New System.Drawing.Size(274, 178)
        Me.lvBookmark.SmallImageList = Me.imgFileTypes
        Me.lvBookmark.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lvBookmark.TabIndex = 0
        Me.lvBookmark.UseCompatibleStateImageBehavior = False
        Me.lvBookmark.View = System.Windows.Forms.View.Details
        Me.lvBookmark.Visible = False
        '
        'cmsBookmark
        '
        Me.cmsBookmark.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.cmsBookmark.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiSelTabsAdd, Me.tsmiSelTabsRefresh, Me.tsmiSelTabsDelete, Me.tssSelTabs1, Me.tsmiSelTabsDefault})
        Me.cmsBookmark.Name = "cmsSelTabs"
        Me.cmsBookmark.Size = New System.Drawing.Size(309, 162)
        '
        'tsmiSelTabsAdd
        '
        Me.tsmiSelTabsAdd.Image = Global.WindowsApplication1.My.Resources.Resources.add_lg
        Me.tsmiSelTabsAdd.Name = "tsmiSelTabsAdd"
        Me.tsmiSelTabsAdd.Size = New System.Drawing.Size(308, 38)
        Me.tsmiSelTabsAdd.Text = "Add bookmark"
        '
        'tsmiSelTabsRefresh
        '
        Me.tsmiSelTabsRefresh.Enabled = False
        Me.tsmiSelTabsRefresh.Name = "tsmiSelTabsRefresh"
        Me.tsmiSelTabsRefresh.Size = New System.Drawing.Size(308, 38)
        Me.tsmiSelTabsRefresh.Text = "Refresh bookmark"
        '
        'tsmiSelTabsDelete
        '
        Me.tsmiSelTabsDelete.Enabled = False
        Me.tsmiSelTabsDelete.Image = Global.WindowsApplication1.My.Resources.Resources.delete_lg
        Me.tsmiSelTabsDelete.Name = "tsmiSelTabsDelete"
        Me.tsmiSelTabsDelete.Size = New System.Drawing.Size(308, 38)
        Me.tsmiSelTabsDelete.Text = "Delete bookmark"
        '
        'tssSelTabs1
        '
        Me.tssSelTabs1.Name = "tssSelTabs1"
        Me.tssSelTabs1.Size = New System.Drawing.Size(305, 6)
        '
        'tsmiSelTabsDefault
        '
        Me.tsmiSelTabsDefault.Name = "tsmiSelTabsDefault"
        Me.tsmiSelTabsDefault.Size = New System.Drawing.Size(308, 38)
        Me.tsmiSelTabsDefault.Text = "Default bookmark"
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Margin = New System.Windows.Forms.Padding(6)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.pnlList)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.lblDocName)
        Me.SplitContainer3.Panel2.Controls.Add(Me.lstPage)
        Me.SplitContainer3.Panel2.Controls.Add(Me.lstInDir)
        Me.SplitContainer3.Panel2.Controls.Add(Me.pctLvIn)
        Me.SplitContainer3.Panel2.Controls.Add(Me.pnlZoomVideo)
        Me.SplitContainer3.Panel2.Controls.Add(Me.wbDoc)
        Me.SplitContainer3.Panel2.Controls.Add(Me.pnlZoom)
        Me.SplitContainer3.Size = New System.Drawing.Size(2022, 178)
        Me.SplitContainer3.SplitterDistance = 265
        Me.SplitContainer3.SplitterWidth = 8
        Me.SplitContainer3.TabIndex = 0
        '
        'pnlList
        '
        Me.pnlList.Controls.Add(Me.pgbZaan)
        Me.pnlList.Controls.Add(Me.lvIn)
        Me.pnlList.Controls.Add(Me.lvMatrix)
        Me.pnlList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlList.Location = New System.Drawing.Point(0, 0)
        Me.pnlList.Margin = New System.Windows.Forms.Padding(6)
        Me.pnlList.Name = "pnlList"
        Me.pnlList.Size = New System.Drawing.Size(265, 178)
        Me.pnlList.TabIndex = 1
        '
        'pgbZaan
        '
        Me.pgbZaan.Dock = System.Windows.Forms.DockStyle.Top
        Me.pgbZaan.Location = New System.Drawing.Point(0, 0)
        Me.pgbZaan.Margin = New System.Windows.Forms.Padding(6)
        Me.pgbZaan.Name = "pgbZaan"
        Me.pgbZaan.Size = New System.Drawing.Size(265, 38)
        Me.pgbZaan.TabIndex = 11
        Me.pgbZaan.Visible = False
        '
        'lvIn
        '
        Me.lvIn.AllowColumnReorder = True
        Me.lvIn.AllowDrop = True
        Me.lvIn.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lvIn.ContextMenuStrip = Me.cmsLvIn
        Me.lvIn.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvIn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvIn.ForeColor = System.Drawing.Color.Black
        Me.lvIn.FullRowSelect = True
        Me.lvIn.LabelEdit = True
        Me.lvIn.LargeImageList = Me.imgLargeIcons
        Me.lvIn.Location = New System.Drawing.Point(0, 0)
        Me.lvIn.Margin = New System.Windows.Forms.Padding(6)
        Me.lvIn.Name = "lvIn"
        Me.lvIn.Size = New System.Drawing.Size(265, 178)
        Me.lvIn.SmallImageList = Me.imgFileTypes
        Me.lvIn.TabIndex = 8
        Me.lvIn.UseCompatibleStateImageBehavior = False
        Me.lvIn.View = System.Windows.Forms.View.Details
        '
        'cmsLvIn
        '
        Me.cmsLvIn.BackColor = System.Drawing.SystemColors.Menu
        Me.cmsLvIn.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.cmsLvIn.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiLvInDocPerPage, Me.tsmiLvInImageMode, Me.tssLvIn1, Me.tsmiLvInRefresh, Me.tsmiLvInSelAll, Me.tsmiLvInUndoMove, Me.tssLvIn2, Me.tsmiLvInCut, Me.tsmiLvInCopy, Me.tsmiLvInPaste, Me.tssLvIn3, Me.tsmiLvInCopyPath, Me.tsmiLvInOpenFolder, Me.tssLvIn4, Me.tsmiLvInAutoRename, Me.tsmiLvInRename, Me.tsmiLvInDelete, Me.tssLvIn5, Me.tsmiLvInNew, Me.tssLvIn6, Me.tsmiLvInCopyToZC, Me.tsmiLvInMoveOut, Me.tssLvIn7, Me.tsmiLvInExport, Me.tsmiLvInExpNameTable, Me.tsmiLvInExpRefTable})
        Me.cmsLvIn.Name = "cmsLvIn"
        Me.cmsLvIn.Size = New System.Drawing.Size(396, 768)
        '
        'tsmiLvInDocPerPage
        '
        Me.tsmiLvInDocPerPage.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiLvInDpp10, Me.tsmiLvInDpp25, Me.tsmiLvInDpp50, Me.tsmiLvInDpp100, Me.tsmiLvInDpp250, Me.tsmiLvInDpp500, Me.tsmiLvInDpp1000})
        Me.tsmiLvInDocPerPage.Name = "tsmiLvInDocPerPage"
        Me.tsmiLvInDocPerPage.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInDocPerPage.Text = "1000 documents per page"
        '
        'tsmiLvInDpp10
        '
        Me.tsmiLvInDpp10.Name = "tsmiLvInDpp10"
        Me.tsmiLvInDpp10.Size = New System.Drawing.Size(167, 38)
        Me.tsmiLvInDpp10.Text = "10"
        '
        'tsmiLvInDpp25
        '
        Me.tsmiLvInDpp25.Name = "tsmiLvInDpp25"
        Me.tsmiLvInDpp25.Size = New System.Drawing.Size(167, 38)
        Me.tsmiLvInDpp25.Text = "25"
        '
        'tsmiLvInDpp50
        '
        Me.tsmiLvInDpp50.Name = "tsmiLvInDpp50"
        Me.tsmiLvInDpp50.Size = New System.Drawing.Size(167, 38)
        Me.tsmiLvInDpp50.Text = "50"
        '
        'tsmiLvInDpp100
        '
        Me.tsmiLvInDpp100.Name = "tsmiLvInDpp100"
        Me.tsmiLvInDpp100.Size = New System.Drawing.Size(167, 38)
        Me.tsmiLvInDpp100.Text = "100"
        '
        'tsmiLvInDpp250
        '
        Me.tsmiLvInDpp250.Name = "tsmiLvInDpp250"
        Me.tsmiLvInDpp250.Size = New System.Drawing.Size(167, 38)
        Me.tsmiLvInDpp250.Text = "250"
        '
        'tsmiLvInDpp500
        '
        Me.tsmiLvInDpp500.Name = "tsmiLvInDpp500"
        Me.tsmiLvInDpp500.Size = New System.Drawing.Size(167, 38)
        Me.tsmiLvInDpp500.Text = "500"
        '
        'tsmiLvInDpp1000
        '
        Me.tsmiLvInDpp1000.Name = "tsmiLvInDpp1000"
        Me.tsmiLvInDpp1000.Size = New System.Drawing.Size(167, 38)
        Me.tsmiLvInDpp1000.Text = "1000"
        '
        'tsmiLvInImageMode
        '
        Me.tsmiLvInImageMode.Name = "tsmiLvInImageMode"
        Me.tsmiLvInImageMode.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInImageMode.Text = "Image mode"
        '
        'tssLvIn1
        '
        Me.tssLvIn1.Name = "tssLvIn1"
        Me.tssLvIn1.Size = New System.Drawing.Size(392, 6)
        '
        'tsmiLvInRefresh
        '
        Me.tsmiLvInRefresh.Name = "tsmiLvInRefresh"
        Me.tsmiLvInRefresh.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInRefresh.Text = "Refresh"
        '
        'tsmiLvInSelAll
        '
        Me.tsmiLvInSelAll.Name = "tsmiLvInSelAll"
        Me.tsmiLvInSelAll.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
        Me.tsmiLvInSelAll.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInSelAll.Text = "Select all"
        '
        'tsmiLvInUndoMove
        '
        Me.tsmiLvInUndoMove.Enabled = False
        Me.tsmiLvInUndoMove.Name = "tsmiLvInUndoMove"
        Me.tsmiLvInUndoMove.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Z), System.Windows.Forms.Keys)
        Me.tsmiLvInUndoMove.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInUndoMove.Text = "Undo move"
        '
        'tssLvIn2
        '
        Me.tssLvIn2.Name = "tssLvIn2"
        Me.tssLvIn2.Size = New System.Drawing.Size(392, 6)
        '
        'tsmiLvInCut
        '
        Me.tsmiLvInCut.Enabled = False
        Me.tsmiLvInCut.Name = "tsmiLvInCut"
        Me.tsmiLvInCut.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
        Me.tsmiLvInCut.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInCut.Text = "Cut"
        '
        'tsmiLvInCopy
        '
        Me.tsmiLvInCopy.Enabled = False
        Me.tsmiLvInCopy.Name = "tsmiLvInCopy"
        Me.tsmiLvInCopy.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.tsmiLvInCopy.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInCopy.Text = "Copy"
        '
        'tsmiLvInPaste
        '
        Me.tsmiLvInPaste.Enabled = False
        Me.tsmiLvInPaste.Name = "tsmiLvInPaste"
        Me.tsmiLvInPaste.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.V), System.Windows.Forms.Keys)
        Me.tsmiLvInPaste.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInPaste.Text = "Paste"
        '
        'tssLvIn3
        '
        Me.tssLvIn3.Name = "tssLvIn3"
        Me.tssLvIn3.Size = New System.Drawing.Size(392, 6)
        '
        'tsmiLvInCopyPath
        '
        Me.tsmiLvInCopyPath.Enabled = False
        Me.tsmiLvInCopyPath.Name = "tsmiLvInCopyPath"
        Me.tsmiLvInCopyPath.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInCopyPath.Text = "Copy location"
        '
        'tsmiLvInOpenFolder
        '
        Me.tsmiLvInOpenFolder.Enabled = False
        Me.tsmiLvInOpenFolder.Name = "tsmiLvInOpenFolder"
        Me.tsmiLvInOpenFolder.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInOpenFolder.Text = "Open folder"
        '
        'tssLvIn4
        '
        Me.tssLvIn4.Name = "tssLvIn4"
        Me.tssLvIn4.Size = New System.Drawing.Size(392, 6)
        '
        'tsmiLvInAutoRename
        '
        Me.tsmiLvInAutoRename.Enabled = False
        Me.tsmiLvInAutoRename.Name = "tsmiLvInAutoRename"
        Me.tsmiLvInAutoRename.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInAutoRename.Text = "Rename automatically"
        '
        'tsmiLvInRename
        '
        Me.tsmiLvInRename.Enabled = False
        Me.tsmiLvInRename.Name = "tsmiLvInRename"
        Me.tsmiLvInRename.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInRename.Text = "Rename"
        '
        'tsmiLvInDelete
        '
        Me.tsmiLvInDelete.Enabled = False
        Me.tsmiLvInDelete.Image = Global.WindowsApplication1.My.Resources.Resources.delete_lg
        Me.tsmiLvInDelete.Name = "tsmiLvInDelete"
        Me.tsmiLvInDelete.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInDelete.Text = "Delete"
        '
        'tssLvIn5
        '
        Me.tssLvIn5.Name = "tssLvIn5"
        Me.tssLvIn5.Size = New System.Drawing.Size(392, 6)
        Me.tssLvIn5.Visible = False
        '
        'tsmiLvInNew
        '
        Me.tsmiLvInNew.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiLvInNewNote, Me.tsmiLvInNewText, Me.tsmiLvInNewPres, Me.tsmiLvInNewSpSh})
        Me.tsmiLvInNew.Name = "tsmiLvInNew"
        Me.tsmiLvInNew.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInNew.Text = "New document"
        Me.tsmiLvInNew.Visible = False
        '
        'tsmiLvInNewNote
        '
        Me.tsmiLvInNewNote.Name = "tsmiLvInNewNote"
        Me.tsmiLvInNewNote.Size = New System.Drawing.Size(248, 38)
        Me.tsmiLvInNewNote.Text = "Note"
        '
        'tsmiLvInNewText
        '
        Me.tsmiLvInNewText.Name = "tsmiLvInNewText"
        Me.tsmiLvInNewText.Size = New System.Drawing.Size(248, 38)
        Me.tsmiLvInNewText.Text = "Text"
        '
        'tsmiLvInNewPres
        '
        Me.tsmiLvInNewPres.Name = "tsmiLvInNewPres"
        Me.tsmiLvInNewPres.Size = New System.Drawing.Size(248, 38)
        Me.tsmiLvInNewPres.Text = "Presentation"
        '
        'tsmiLvInNewSpSh
        '
        Me.tsmiLvInNewSpSh.Name = "tsmiLvInNewSpSh"
        Me.tsmiLvInNewSpSh.Size = New System.Drawing.Size(248, 38)
        Me.tsmiLvInNewSpSh.Text = "Spreadsheet"
        '
        'tssLvIn6
        '
        Me.tssLvIn6.Name = "tssLvIn6"
        Me.tssLvIn6.Size = New System.Drawing.Size(392, 6)
        '
        'tsmiLvInCopyToZC
        '
        Me.tsmiLvInCopyToZC.Enabled = False
        Me.tsmiLvInCopyToZC.Image = Global.WindowsApplication1.My.Resources.Resources.copy_lg
        Me.tsmiLvInCopyToZC.Name = "tsmiLvInCopyToZC"
        Me.tsmiLvInCopyToZC.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInCopyToZC.Text = "Copy to ZAAN\copy"
        '
        'tsmiLvInMoveOut
        '
        Me.tsmiLvInMoveOut.Enabled = False
        Me.tsmiLvInMoveOut.Image = Global.WindowsApplication1.My.Resources.Resources.down_lg
        Me.tsmiLvInMoveOut.Name = "tsmiLvInMoveOut"
        Me.tsmiLvInMoveOut.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInMoveOut.Text = "Move out"
        '
        'tssLvIn7
        '
        Me.tssLvIn7.Name = "tssLvIn7"
        Me.tssLvIn7.Size = New System.Drawing.Size(392, 6)
        '
        'tsmiLvInExport
        '
        Me.tsmiLvInExport.Enabled = False
        Me.tsmiLvInExport.Name = "tsmiLvInExport"
        Me.tsmiLvInExport.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInExport.Text = "Export ZAAN cube"
        '
        'tsmiLvInExpNameTable
        '
        Me.tsmiLvInExpNameTable.Enabled = False
        Me.tsmiLvInExpNameTable.Name = "tsmiLvInExpNameTable"
        Me.tsmiLvInExpNameTable.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInExpNameTable.Text = "Export name table"
        '
        'tsmiLvInExpRefTable
        '
        Me.tsmiLvInExpRefTable.Enabled = False
        Me.tsmiLvInExpRefTable.Name = "tsmiLvInExpRefTable"
        Me.tsmiLvInExpRefTable.Size = New System.Drawing.Size(395, 38)
        Me.tsmiLvInExpRefTable.Text = "Export reference table"
        '
        'imgLargeIcons
        '
        Me.imgLargeIcons.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit
        Me.imgLargeIcons.ImageSize = New System.Drawing.Size(90, 90)
        Me.imgLargeIcons.TransparentColor = System.Drawing.Color.Transparent
        '
        'lvMatrix
        '
        Me.lvMatrix.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lvMatrix.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvMatrix.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvMatrix.GridLines = True
        Me.lvMatrix.Location = New System.Drawing.Point(0, 0)
        Me.lvMatrix.Margin = New System.Windows.Forms.Padding(6)
        Me.lvMatrix.Name = "lvMatrix"
        Me.lvMatrix.Size = New System.Drawing.Size(265, 178)
        Me.lvMatrix.SmallImageList = Me.imgFileTypes
        Me.lvMatrix.TabIndex = 12
        Me.lvMatrix.UseCompatibleStateImageBehavior = False
        Me.lvMatrix.View = System.Windows.Forms.View.Details
        Me.lvMatrix.Visible = False
        '
        'lblDocName
        '
        Me.lblDocName.BackColor = System.Drawing.SystemColors.Control
        Me.lblDocName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDocName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDocName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDocName.Location = New System.Drawing.Point(132, 8)
        Me.lblDocName.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.lblDocName.Name = "lblDocName"
        Me.lblDocName.Size = New System.Drawing.Size(216, 31)
        Me.lblDocName.TabIndex = 13
        Me.lblDocName.Tag = ""
        Me.lblDocName.Text = "lblDocName"
        Me.lblDocName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblDocName.Visible = False
        '
        'lstPage
        '
        Me.lstPage.BackColor = System.Drawing.SystemColors.Control
        Me.lstPage.FormattingEnabled = True
        Me.lstPage.ItemHeight = 38
        Me.lstPage.Location = New System.Drawing.Point(250, 390)
        Me.lstPage.Margin = New System.Windows.Forms.Padding(6)
        Me.lstPage.Name = "lstPage"
        Me.lstPage.Size = New System.Drawing.Size(112, 80)
        Me.lstPage.TabIndex = 16
        Me.lstPage.Visible = False
        '
        'lstInDir
        '
        Me.lstInDir.BackColor = System.Drawing.SystemColors.Control
        Me.lstInDir.FormattingEnabled = True
        Me.lstInDir.ItemHeight = 38
        Me.lstInDir.Location = New System.Drawing.Point(122, 388)
        Me.lstInDir.Margin = New System.Windows.Forms.Padding(6)
        Me.lstInDir.Name = "lstInDir"
        Me.lstInDir.Size = New System.Drawing.Size(112, 80)
        Me.lstInDir.TabIndex = 15
        Me.lstInDir.Visible = False
        '
        'pctLvIn
        '
        Me.pctLvIn.BackColor = System.Drawing.SystemColors.Control
        Me.pctLvIn.Location = New System.Drawing.Point(514, 300)
        Me.pctLvIn.Margin = New System.Windows.Forms.Padding(6)
        Me.pctLvIn.Name = "pctLvIn"
        Me.pctLvIn.Size = New System.Drawing.Size(180, 173)
        Me.pctLvIn.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pctLvIn.TabIndex = 10
        Me.pctLvIn.TabStop = False
        Me.pctLvIn.Visible = False
        '
        'pnlZoomVideo
        '
        Me.pnlZoomVideo.BackColor = System.Drawing.SystemColors.ControlLight
        Me.pnlZoomVideo.Controls.Add(Me.pnlVideoControl)
        Me.pnlZoomVideo.Location = New System.Drawing.Point(892, 63)
        Me.pnlZoomVideo.Margin = New System.Windows.Forms.Padding(6)
        Me.pnlZoomVideo.Name = "pnlZoomVideo"
        Me.pnlZoomVideo.Size = New System.Drawing.Size(214, 171)
        Me.pnlZoomVideo.TabIndex = 14
        Me.pnlZoomVideo.Visible = False
        '
        'pnlVideoControl
        '
        Me.pnlVideoControl.BackColor = System.Drawing.Color.Black
        Me.pnlVideoControl.Controls.Add(Me.pgbVideo)
        Me.pnlVideoControl.Controls.Add(Me.btnVideoBegin)
        Me.pnlVideoControl.Controls.Add(Me.btnVideoEnd)
        Me.pnlVideoControl.Controls.Add(Me.btnVideoPlayPause)
        Me.pnlVideoControl.Controls.Add(Me.btnVideoStop)
        Me.pnlVideoControl.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlVideoControl.Location = New System.Drawing.Point(0, 0)
        Me.pnlVideoControl.Margin = New System.Windows.Forms.Padding(6)
        Me.pnlVideoControl.Name = "pnlVideoControl"
        Me.pnlVideoControl.Size = New System.Drawing.Size(214, 50)
        Me.pnlVideoControl.TabIndex = 0
        '
        'pgbVideo
        '
        Me.pgbVideo.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pgbVideo.Location = New System.Drawing.Point(0, 35)
        Me.pgbVideo.Margin = New System.Windows.Forms.Padding(6)
        Me.pgbVideo.Name = "pgbVideo"
        Me.pgbVideo.Size = New System.Drawing.Size(38, 15)
        Me.pgbVideo.TabIndex = 17
        '
        'btnVideoBegin
        '
        Me.btnVideoBegin.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnVideoBegin.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnVideoBegin.Image = Global.WindowsApplication1.My.Resources.Resources.previous
        Me.btnVideoBegin.Location = New System.Drawing.Point(38, 0)
        Me.btnVideoBegin.Margin = New System.Windows.Forms.Padding(6)
        Me.btnVideoBegin.Name = "btnVideoBegin"
        Me.btnVideoBegin.Size = New System.Drawing.Size(44, 50)
        Me.btnVideoBegin.TabIndex = 13
        Me.btnVideoBegin.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnVideoBegin.UseVisualStyleBackColor = True
        '
        'btnVideoEnd
        '
        Me.btnVideoEnd.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnVideoEnd.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnVideoEnd.Image = Global.WindowsApplication1.My.Resources.Resources.next_z
        Me.btnVideoEnd.Location = New System.Drawing.Point(82, 0)
        Me.btnVideoEnd.Margin = New System.Windows.Forms.Padding(6)
        Me.btnVideoEnd.Name = "btnVideoEnd"
        Me.btnVideoEnd.Size = New System.Drawing.Size(44, 50)
        Me.btnVideoEnd.TabIndex = 15
        Me.btnVideoEnd.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnVideoEnd.UseVisualStyleBackColor = True
        '
        'btnVideoPlayPause
        '
        Me.btnVideoPlayPause.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnVideoPlayPause.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnVideoPlayPause.Image = Global.WindowsApplication1.My.Resources.Resources.pause
        Me.btnVideoPlayPause.Location = New System.Drawing.Point(126, 0)
        Me.btnVideoPlayPause.Margin = New System.Windows.Forms.Padding(6)
        Me.btnVideoPlayPause.Name = "btnVideoPlayPause"
        Me.btnVideoPlayPause.Size = New System.Drawing.Size(44, 50)
        Me.btnVideoPlayPause.TabIndex = 14
        Me.btnVideoPlayPause.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnVideoPlayPause.UseVisualStyleBackColor = True
        '
        'btnVideoStop
        '
        Me.btnVideoStop.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnVideoStop.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnVideoStop.Image = Global.WindowsApplication1.My.Resources.Resources.stop_z
        Me.btnVideoStop.Location = New System.Drawing.Point(170, 0)
        Me.btnVideoStop.Margin = New System.Windows.Forms.Padding(6)
        Me.btnVideoStop.Name = "btnVideoStop"
        Me.btnVideoStop.Size = New System.Drawing.Size(44, 50)
        Me.btnVideoStop.TabIndex = 16
        Me.btnVideoStop.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnVideoStop.UseVisualStyleBackColor = True
        '
        'wbDoc
        '
        Me.wbDoc.Dock = System.Windows.Forms.DockStyle.Left
        Me.wbDoc.Location = New System.Drawing.Point(0, 0)
        Me.wbDoc.Margin = New System.Windows.Forms.Padding(6)
        Me.wbDoc.MinimumSize = New System.Drawing.Size(40, 38)
        Me.wbDoc.Name = "wbDoc"
        Me.wbDoc.Size = New System.Drawing.Size(108, 178)
        Me.wbDoc.TabIndex = 8
        Me.wbDoc.Visible = False
        '
        'pnlZoom
        '
        Me.pnlZoom.BackColor = System.Drawing.SystemColors.ControlLight
        Me.pnlZoom.Controls.Add(Me.pctZoom)
        Me.pnlZoom.Controls.Add(Me.pnlSlideControl)
        Me.pnlZoom.Controls.Add(Me.pctZoom2)
        Me.pnlZoom.Location = New System.Drawing.Point(120, 63)
        Me.pnlZoom.Margin = New System.Windows.Forms.Padding(6)
        Me.pnlZoom.Name = "pnlZoom"
        Me.pnlZoom.Size = New System.Drawing.Size(732, 175)
        Me.pnlZoom.TabIndex = 12
        Me.pnlZoom.Visible = False
        '
        'pctZoom
        '
        Me.pctZoom.BackColor = System.Drawing.Color.Black
        Me.pctZoom.ContextMenuStrip = Me.cmsPctZoom
        Me.pctZoom.Cursor = System.Windows.Forms.Cursors.Hand
        Me.pctZoom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pctZoom.Location = New System.Drawing.Point(0, 50)
        Me.pctZoom.Margin = New System.Windows.Forms.Padding(6)
        Me.pctZoom.Name = "pctZoom"
        Me.pctZoom.Size = New System.Drawing.Size(732, 125)
        Me.pctZoom.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pctZoom.TabIndex = 9
        Me.pctZoom.TabStop = False
        '
        'cmsPctZoom
        '
        Me.cmsPctZoom.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.cmsPctZoom.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiPctZoomSlideShow, Me.tssPctZoom1, Me.tsmiPctZoomInterval, Me.tstbPctZoom})
        Me.cmsPctZoom.Name = "cmsPctZoom"
        Me.cmsPctZoom.Size = New System.Drawing.Size(311, 127)
        '
        'tsmiPctZoomSlideShow
        '
        Me.tsmiPctZoomSlideShow.CheckOnClick = True
        Me.tsmiPctZoomSlideShow.Name = "tsmiPctZoomSlideShow"
        Me.tsmiPctZoomSlideShow.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.tsmiPctZoomSlideShow.Size = New System.Drawing.Size(310, 38)
        Me.tsmiPctZoomSlideShow.Text = "Slide show"
        '
        'tssPctZoom1
        '
        Me.tssPctZoom1.Name = "tssPctZoom1"
        Me.tssPctZoom1.Size = New System.Drawing.Size(307, 6)
        '
        'tsmiPctZoomInterval
        '
        Me.tsmiPctZoomInterval.Name = "tsmiPctZoomInterval"
        Me.tsmiPctZoomInterval.Size = New System.Drawing.Size(310, 38)
        Me.tsmiPctZoomInterval.Text = "Interval (sec.) :"
        '
        'tstbPctZoom
        '
        Me.tstbPctZoom.Name = "tstbPctZoom"
        Me.tstbPctZoom.Size = New System.Drawing.Size(100, 39)
        Me.tstbPctZoom.Text = "3"
        '
        'pnlSlideControl
        '
        Me.pnlSlideControl.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(20, Byte), Integer))
        Me.pnlSlideControl.Controls.Add(Me.lblZoomName)
        Me.pnlSlideControl.Controls.Add(Me.btnSlideRotateLeft)
        Me.pnlSlideControl.Controls.Add(Me.btnSlideRotateRight)
        Me.pnlSlideControl.Controls.Add(Me.btnSlideDelete)
        Me.pnlSlideControl.Controls.Add(Me.lblSlideNb)
        Me.pnlSlideControl.Controls.Add(Me.btnSlidePrevious)
        Me.pnlSlideControl.Controls.Add(Me.btnSlideNext)
        Me.pnlSlideControl.Controls.Add(Me.btnSlidePlayPause)
        Me.pnlSlideControl.Controls.Add(Me.btnSlideStop)
        Me.pnlSlideControl.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlSlideControl.Location = New System.Drawing.Point(0, 0)
        Me.pnlSlideControl.Margin = New System.Windows.Forms.Padding(6)
        Me.pnlSlideControl.Name = "pnlSlideControl"
        Me.pnlSlideControl.Padding = New System.Windows.Forms.Padding(4)
        Me.pnlSlideControl.Size = New System.Drawing.Size(732, 50)
        Me.pnlSlideControl.TabIndex = 11
        Me.pnlSlideControl.Visible = False
        '
        'lblZoomName
        '
        Me.lblZoomName.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblZoomName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblZoomName.ForeColor = System.Drawing.Color.RoyalBlue
        Me.lblZoomName.Location = New System.Drawing.Point(4, 4)
        Me.lblZoomName.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.lblZoomName.Name = "lblZoomName"
        Me.lblZoomName.Size = New System.Drawing.Size(136, 42)
        Me.lblZoomName.TabIndex = 13
        Me.lblZoomName.Text = "Name"
        Me.lblZoomName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnSlideRotateLeft
        '
        Me.btnSlideRotateLeft.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSlideRotateLeft.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnSlideRotateLeft.Image = Global.WindowsApplication1.My.Resources.Resources.turn_left
        Me.btnSlideRotateLeft.Location = New System.Drawing.Point(140, 4)
        Me.btnSlideRotateLeft.Margin = New System.Windows.Forms.Padding(6)
        Me.btnSlideRotateLeft.Name = "btnSlideRotateLeft"
        Me.btnSlideRotateLeft.Size = New System.Drawing.Size(44, 42)
        Me.btnSlideRotateLeft.TabIndex = 16
        Me.btnSlideRotateLeft.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnSlideRotateLeft.UseVisualStyleBackColor = True
        '
        'btnSlideRotateRight
        '
        Me.btnSlideRotateRight.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSlideRotateRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnSlideRotateRight.Image = Global.WindowsApplication1.My.Resources.Resources.turn_right
        Me.btnSlideRotateRight.Location = New System.Drawing.Point(184, 4)
        Me.btnSlideRotateRight.Margin = New System.Windows.Forms.Padding(6)
        Me.btnSlideRotateRight.Name = "btnSlideRotateRight"
        Me.btnSlideRotateRight.Size = New System.Drawing.Size(44, 42)
        Me.btnSlideRotateRight.TabIndex = 15
        Me.btnSlideRotateRight.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnSlideRotateRight.UseVisualStyleBackColor = True
        '
        'btnSlideDelete
        '
        Me.btnSlideDelete.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSlideDelete.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnSlideDelete.Image = Global.WindowsApplication1.My.Resources.Resources.delete
        Me.btnSlideDelete.Location = New System.Drawing.Point(228, 4)
        Me.btnSlideDelete.Margin = New System.Windows.Forms.Padding(6)
        Me.btnSlideDelete.Name = "btnSlideDelete"
        Me.btnSlideDelete.Size = New System.Drawing.Size(44, 42)
        Me.btnSlideDelete.TabIndex = 14
        Me.btnSlideDelete.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnSlideDelete.UseVisualStyleBackColor = True
        '
        'lblSlideNb
        '
        Me.lblSlideNb.BackColor = System.Drawing.Color.Transparent
        Me.lblSlideNb.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSlideNb.Dock = System.Windows.Forms.DockStyle.Right
        Me.lblSlideNb.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSlideNb.ForeColor = System.Drawing.Color.RoyalBlue
        Me.lblSlideNb.Location = New System.Drawing.Point(272, 4)
        Me.lblSlideNb.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.lblSlideNb.Name = "lblSlideNb"
        Me.lblSlideNb.Size = New System.Drawing.Size(280, 42)
        Me.lblSlideNb.TabIndex = 11
        Me.lblSlideNb.Text = "100 000 / 100 000"
        Me.lblSlideNb.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnSlidePrevious
        '
        Me.btnSlidePrevious.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSlidePrevious.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnSlidePrevious.Image = Global.WindowsApplication1.My.Resources.Resources.previous
        Me.btnSlidePrevious.Location = New System.Drawing.Point(552, 4)
        Me.btnSlidePrevious.Margin = New System.Windows.Forms.Padding(6)
        Me.btnSlidePrevious.Name = "btnSlidePrevious"
        Me.btnSlidePrevious.Size = New System.Drawing.Size(44, 42)
        Me.btnSlidePrevious.TabIndex = 8
        Me.btnSlidePrevious.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnSlidePrevious.UseVisualStyleBackColor = True
        '
        'btnSlideNext
        '
        Me.btnSlideNext.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSlideNext.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnSlideNext.Image = Global.WindowsApplication1.My.Resources.Resources.next_z
        Me.btnSlideNext.Location = New System.Drawing.Point(596, 4)
        Me.btnSlideNext.Margin = New System.Windows.Forms.Padding(6)
        Me.btnSlideNext.Name = "btnSlideNext"
        Me.btnSlideNext.Size = New System.Drawing.Size(44, 42)
        Me.btnSlideNext.TabIndex = 10
        Me.btnSlideNext.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnSlideNext.UseVisualStyleBackColor = True
        '
        'btnSlidePlayPause
        '
        Me.btnSlidePlayPause.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSlidePlayPause.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnSlidePlayPause.Image = Global.WindowsApplication1.My.Resources.Resources.pause
        Me.btnSlidePlayPause.Location = New System.Drawing.Point(640, 4)
        Me.btnSlidePlayPause.Margin = New System.Windows.Forms.Padding(6)
        Me.btnSlidePlayPause.Name = "btnSlidePlayPause"
        Me.btnSlidePlayPause.Size = New System.Drawing.Size(44, 42)
        Me.btnSlidePlayPause.TabIndex = 9
        Me.btnSlidePlayPause.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnSlidePlayPause.UseVisualStyleBackColor = True
        '
        'btnSlideStop
        '
        Me.btnSlideStop.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSlideStop.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnSlideStop.Image = Global.WindowsApplication1.My.Resources.Resources.stop_z
        Me.btnSlideStop.Location = New System.Drawing.Point(684, 4)
        Me.btnSlideStop.Margin = New System.Windows.Forms.Padding(6)
        Me.btnSlideStop.Name = "btnSlideStop"
        Me.btnSlideStop.Size = New System.Drawing.Size(44, 42)
        Me.btnSlideStop.TabIndex = 12
        Me.btnSlideStop.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnSlideStop.UseVisualStyleBackColor = True
        '
        'pctZoom2
        '
        Me.pctZoom2.BackColor = System.Drawing.Color.Black
        Me.pctZoom2.Location = New System.Drawing.Point(296, 100)
        Me.pctZoom2.Margin = New System.Windows.Forms.Padding(6)
        Me.pctZoom2.Name = "pctZoom2"
        Me.pctZoom2.Size = New System.Drawing.Size(62, 56)
        Me.pctZoom2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pctZoom2.TabIndex = 10
        Me.pctZoom2.TabStop = False
        Me.pctZoom2.Visible = False
        '
        'pnlListTop
        '
        Me.pnlListTop.BackColor = System.Drawing.SystemColors.Control
        Me.pnlListTop.Controls.Add(Me.pnlSelectorNav)
        Me.pnlListTop.Controls.Add(Me.pnlSelectorSearch)
        Me.pnlListTop.Controls.Add(Me.pnlDatabase)
        Me.pnlListTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlListTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlListTop.Margin = New System.Windows.Forms.Padding(6)
        Me.pnlListTop.Name = "pnlListTop"
        Me.pnlListTop.Size = New System.Drawing.Size(2304, 137)
        Me.pnlListTop.TabIndex = 0
        '
        'pnlSelectorNav
        '
        Me.pnlSelectorNav.Controls.Add(Me.pnlSelectorHeaderBottom)
        Me.pnlSelectorNav.Controls.Add(Me.pnlSelectorHeaderTop)
        Me.pnlSelectorNav.Controls.Add(Me.pnlSelectorHeader)
        Me.pnlSelectorNav.Controls.Add(Me.lvSelector)
        Me.pnlSelectorNav.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlSelectorNav.Location = New System.Drawing.Point(698, 0)
        Me.pnlSelectorNav.Margin = New System.Windows.Forms.Padding(6)
        Me.pnlSelectorNav.Name = "pnlSelectorNav"
        Me.pnlSelectorNav.Padding = New System.Windows.Forms.Padding(0, 0, 10, 0)
        Me.pnlSelectorNav.Size = New System.Drawing.Size(1606, 137)
        Me.pnlSelectorNav.TabIndex = 17
        '
        'pnlSelectorHeaderBottom
        '
        Me.pnlSelectorHeaderBottom.Controls.Add(Me.btnWho2Axis)
        Me.pnlSelectorHeaderBottom.Controls.Add(Me.btnWhat2Axis)
        Me.pnlSelectorHeaderBottom.Controls.Add(Me.btnWhereAxis)
        Me.pnlSelectorHeaderBottom.Controls.Add(Me.btnWhatAxis)
        Me.pnlSelectorHeaderBottom.Controls.Add(Me.btnPanelView)
        Me.pnlSelectorHeaderBottom.Controls.Add(Me.btnWhoAxis)
        Me.pnlSelectorHeaderBottom.Controls.Add(Me.btnWhenAxis)
        Me.pnlSelectorHeaderBottom.Controls.Add(Me.btnToday)
        Me.pnlSelectorHeaderBottom.Controls.Add(Me.btnDataAccessBlank)
        Me.pnlSelectorHeaderBottom.Controls.Add(Me.btnDataAccessAxis)
        Me.pnlSelectorHeaderBottom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlSelectorHeaderBottom.Location = New System.Drawing.Point(0, 90)
        Me.pnlSelectorHeaderBottom.Margin = New System.Windows.Forms.Padding(6)
        Me.pnlSelectorHeaderBottom.Name = "pnlSelectorHeaderBottom"
        Me.pnlSelectorHeaderBottom.Padding = New System.Windows.Forms.Padding(0, 0, 2, 0)
        Me.pnlSelectorHeaderBottom.Size = New System.Drawing.Size(1596, 47)
        Me.pnlSelectorHeaderBottom.TabIndex = 20
        '
        'btnWho2Axis
        '
        Me.btnWho2Axis.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWho2Axis.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWho2Axis.FlatAppearance.BorderSize = 0
        Me.btnWho2Axis.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWho2Axis.Font = New System.Drawing.Font("Lucida Console", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWho2Axis.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWho2Axis.Location = New System.Drawing.Point(1120, 0)
        Me.btnWho2Axis.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWho2Axis.Name = "btnWho2Axis"
        Me.btnWho2Axis.Size = New System.Drawing.Size(140, 47)
        Me.btnWho2Axis.TabIndex = 40
        Me.btnWho2Axis.Text = "+"
        Me.btnWho2Axis.UseVisualStyleBackColor = True
        '
        'btnWhat2Axis
        '
        Me.btnWhat2Axis.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWhat2Axis.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWhat2Axis.FlatAppearance.BorderSize = 0
        Me.btnWhat2Axis.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWhat2Axis.Font = New System.Drawing.Font("Lucida Console", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWhat2Axis.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhat2Axis.Location = New System.Drawing.Point(980, 0)
        Me.btnWhat2Axis.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWhat2Axis.Name = "btnWhat2Axis"
        Me.btnWhat2Axis.Size = New System.Drawing.Size(140, 47)
        Me.btnWhat2Axis.TabIndex = 39
        Me.btnWhat2Axis.Text = "+"
        Me.btnWhat2Axis.UseVisualStyleBackColor = True
        '
        'btnWhereAxis
        '
        Me.btnWhereAxis.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWhereAxis.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWhereAxis.FlatAppearance.BorderSize = 0
        Me.btnWhereAxis.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWhereAxis.Font = New System.Drawing.Font("Lucida Console", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWhereAxis.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhereAxis.Location = New System.Drawing.Point(840, 0)
        Me.btnWhereAxis.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWhereAxis.Name = "btnWhereAxis"
        Me.btnWhereAxis.Size = New System.Drawing.Size(140, 47)
        Me.btnWhereAxis.TabIndex = 38
        Me.btnWhereAxis.Text = "+"
        Me.btnWhereAxis.UseVisualStyleBackColor = True
        '
        'btnWhatAxis
        '
        Me.btnWhatAxis.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWhatAxis.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWhatAxis.FlatAppearance.BorderSize = 0
        Me.btnWhatAxis.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWhatAxis.Font = New System.Drawing.Font("Lucida Console", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWhatAxis.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhatAxis.Location = New System.Drawing.Point(700, 0)
        Me.btnWhatAxis.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWhatAxis.Name = "btnWhatAxis"
        Me.btnWhatAxis.Size = New System.Drawing.Size(140, 47)
        Me.btnWhatAxis.TabIndex = 37
        Me.btnWhatAxis.Text = "+"
        Me.btnWhatAxis.UseVisualStyleBackColor = True
        '
        'btnPanelView
        '
        Me.btnPanelView.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnPanelView.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnPanelView.FlatAppearance.BorderSize = 0
        Me.btnPanelView.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnPanelView.Image = Global.WindowsApplication1.My.Resources.Resources.pnl_view_1
        Me.btnPanelView.Location = New System.Drawing.Point(1546, 0)
        Me.btnPanelView.Margin = New System.Windows.Forms.Padding(6)
        Me.btnPanelView.Name = "btnPanelView"
        Me.btnPanelView.Padding = New System.Windows.Forms.Padding(0, 0, 0, 2)
        Me.btnPanelView.Size = New System.Drawing.Size(48, 47)
        Me.btnPanelView.TabIndex = 28
        Me.btnPanelView.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnPanelView.UseVisualStyleBackColor = False
        '
        'btnWhoAxis
        '
        Me.btnWhoAxis.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWhoAxis.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWhoAxis.FlatAppearance.BorderSize = 0
        Me.btnWhoAxis.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWhoAxis.Font = New System.Drawing.Font("Lucida Console", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWhoAxis.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhoAxis.Location = New System.Drawing.Point(560, 0)
        Me.btnWhoAxis.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWhoAxis.Name = "btnWhoAxis"
        Me.btnWhoAxis.Size = New System.Drawing.Size(140, 47)
        Me.btnWhoAxis.TabIndex = 36
        Me.btnWhoAxis.Text = "+"
        Me.btnWhoAxis.UseVisualStyleBackColor = True
        '
        'btnWhenAxis
        '
        Me.btnWhenAxis.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWhenAxis.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWhenAxis.FlatAppearance.BorderSize = 0
        Me.btnWhenAxis.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWhenAxis.Font = New System.Drawing.Font("Lucida Console", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWhenAxis.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhenAxis.Location = New System.Drawing.Point(420, 0)
        Me.btnWhenAxis.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWhenAxis.Name = "btnWhenAxis"
        Me.btnWhenAxis.Size = New System.Drawing.Size(140, 47)
        Me.btnWhenAxis.TabIndex = 35
        Me.btnWhenAxis.Text = "+"
        Me.btnWhenAxis.UseVisualStyleBackColor = True
        '
        'btnToday
        '
        Me.btnToday.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnToday.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnToday.FlatAppearance.BorderSize = 0
        Me.btnToday.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnToday.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnToday.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnToday.Location = New System.Drawing.Point(280, 0)
        Me.btnToday.Margin = New System.Windows.Forms.Padding(6)
        Me.btnToday.Name = "btnToday"
        Me.btnToday.Size = New System.Drawing.Size(140, 47)
        Me.btnToday.TabIndex = 41
        Me.btnToday.Text = "2010-09-01"
        Me.btnToday.UseVisualStyleBackColor = True
        '
        'btnDataAccessBlank
        '
        Me.btnDataAccessBlank.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnDataAccessBlank.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnDataAccessBlank.FlatAppearance.BorderSize = 0
        Me.btnDataAccessBlank.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDataAccessBlank.Font = New System.Drawing.Font("Lucida Console", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDataAccessBlank.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDataAccessBlank.Location = New System.Drawing.Point(140, 0)
        Me.btnDataAccessBlank.Margin = New System.Windows.Forms.Padding(6)
        Me.btnDataAccessBlank.Name = "btnDataAccessBlank"
        Me.btnDataAccessBlank.Size = New System.Drawing.Size(140, 47)
        Me.btnDataAccessBlank.TabIndex = 42
        Me.btnDataAccessBlank.UseVisualStyleBackColor = True
        '
        'btnDataAccessAxis
        '
        Me.btnDataAccessAxis.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnDataAccessAxis.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnDataAccessAxis.FlatAppearance.BorderSize = 0
        Me.btnDataAccessAxis.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDataAccessAxis.Font = New System.Drawing.Font("Lucida Console", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDataAccessAxis.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDataAccessAxis.Location = New System.Drawing.Point(0, 0)
        Me.btnDataAccessAxis.Margin = New System.Windows.Forms.Padding(6)
        Me.btnDataAccessAxis.Name = "btnDataAccessAxis"
        Me.btnDataAccessAxis.Size = New System.Drawing.Size(140, 47)
        Me.btnDataAccessAxis.TabIndex = 34
        Me.btnDataAccessAxis.Text = "+"
        Me.btnDataAccessAxis.UseVisualStyleBackColor = True
        '
        'pnlSelectorHeaderTop
        '
        Me.pnlSelectorHeaderTop.Controls.Add(Me.btnCubeTube)
        Me.pnlSelectorHeaderTop.Controls.Add(Me.btnPanelImport)
        Me.pnlSelectorHeaderTop.Controls.Add(Me.btnWho2)
        Me.pnlSelectorHeaderTop.Controls.Add(Me.btnWho2Root)
        Me.pnlSelectorHeaderTop.Controls.Add(Me.btnWhat2)
        Me.pnlSelectorHeaderTop.Controls.Add(Me.btnWhat2Root)
        Me.pnlSelectorHeaderTop.Controls.Add(Me.btnWhere)
        Me.pnlSelectorHeaderTop.Controls.Add(Me.btnWhereRoot)
        Me.pnlSelectorHeaderTop.Controls.Add(Me.btnWhat)
        Me.pnlSelectorHeaderTop.Controls.Add(Me.btnWhatRoot)
        Me.pnlSelectorHeaderTop.Controls.Add(Me.btnWho)
        Me.pnlSelectorHeaderTop.Controls.Add(Me.btnWhoRoot)
        Me.pnlSelectorHeaderTop.Controls.Add(Me.btnWhen)
        Me.pnlSelectorHeaderTop.Controls.Add(Me.btnWhenRoot)
        Me.pnlSelectorHeaderTop.Controls.Add(Me.btnDataAccess)
        Me.pnlSelectorHeaderTop.Controls.Add(Me.btnDataAccessRoot)
        Me.pnlSelectorHeaderTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlSelectorHeaderTop.Margin = New System.Windows.Forms.Padding(6)
        Me.pnlSelectorHeaderTop.Name = "pnlSelectorHeaderTop"
        Me.pnlSelectorHeaderTop.Size = New System.Drawing.Size(1502, 46)
        Me.pnlSelectorHeaderTop.TabIndex = 19
        '
        'btnCubeTube
        '
        Me.btnCubeTube.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnCubeTube.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnCubeTube.FlatAppearance.BorderSize = 0
        Me.btnCubeTube.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCubeTube.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCubeTube.ForeColor = System.Drawing.Color.MidnightBlue
        Me.btnCubeTube.Image = Global.WindowsApplication1.My.Resources.Resources.view_cube_1
        Me.btnCubeTube.Location = New System.Drawing.Point(1406, 0)
        Me.btnCubeTube.Margin = New System.Windows.Forms.Padding(6)
        Me.btnCubeTube.Name = "btnCubeTube"
        Me.btnCubeTube.Size = New System.Drawing.Size(48, 46)
        Me.btnCubeTube.TabIndex = 24
        Me.btnCubeTube.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnCubeTube.UseVisualStyleBackColor = True
        Me.btnCubeTube.Visible = False
        '
        'btnPanelImport
        '
        Me.btnPanelImport.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnPanelImport.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnPanelImport.FlatAppearance.BorderSize = 0
        Me.btnPanelImport.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnPanelImport.Image = Global.WindowsApplication1.My.Resources.Resources.pnl_import_1
        Me.btnPanelImport.Location = New System.Drawing.Point(1454, 0)
        Me.btnPanelImport.Margin = New System.Windows.Forms.Padding(6)
        Me.btnPanelImport.Name = "btnPanelImport"
        Me.btnPanelImport.Padding = New System.Windows.Forms.Padding(0, 0, 2, 2)
        Me.btnPanelImport.Size = New System.Drawing.Size(48, 46)
        Me.btnPanelImport.TabIndex = 41
        Me.btnPanelImport.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnPanelImport.UseVisualStyleBackColor = False
        '
        'btnWho2
        '
        Me.btnWho2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWho2.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWho2.FlatAppearance.BorderSize = 0
        Me.btnWho2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWho2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWho2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWho2.ImageList = Me.imgFileTypes
        Me.btnWho2.Location = New System.Drawing.Point(880, 0)
        Me.btnWho2.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWho2.Name = "btnWho2"
        Me.btnWho2.Size = New System.Drawing.Size(100, 46)
        Me.btnWho2.TabIndex = 33
        Me.btnWho2.Text = "Wo2"
        Me.btnWho2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWho2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnWho2.UseVisualStyleBackColor = True
        '
        'btnWho2Root
        '
        Me.btnWho2Root.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWho2Root.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWho2Root.FlatAppearance.BorderSize = 0
        Me.btnWho2Root.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWho2Root.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWho2Root.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWho2Root.ImageList = Me.imgFileTypes
        Me.btnWho2Root.Location = New System.Drawing.Point(840, 0)
        Me.btnWho2Root.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWho2Root.Name = "btnWho2Root"
        Me.btnWho2Root.Size = New System.Drawing.Size(40, 46)
        Me.btnWho2Root.TabIndex = 40
        Me.btnWho2Root.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWho2Root.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnWho2Root.UseVisualStyleBackColor = True
        '
        'btnWhat2
        '
        Me.btnWhat2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWhat2.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWhat2.FlatAppearance.BorderSize = 0
        Me.btnWhat2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWhat2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWhat2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhat2.ImageList = Me.imgFileTypes
        Me.btnWhat2.Location = New System.Drawing.Point(740, 0)
        Me.btnWhat2.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWhat2.Name = "btnWhat2"
        Me.btnWhat2.Size = New System.Drawing.Size(100, 46)
        Me.btnWhat2.TabIndex = 32
        Me.btnWhat2.Text = "Wa2"
        Me.btnWhat2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhat2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnWhat2.UseVisualStyleBackColor = True
        '
        'btnWhat2Root
        '
        Me.btnWhat2Root.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWhat2Root.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWhat2Root.FlatAppearance.BorderSize = 0
        Me.btnWhat2Root.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWhat2Root.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWhat2Root.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhat2Root.ImageList = Me.imgFileTypes
        Me.btnWhat2Root.Location = New System.Drawing.Point(700, 0)
        Me.btnWhat2Root.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWhat2Root.Name = "btnWhat2Root"
        Me.btnWhat2Root.Size = New System.Drawing.Size(40, 46)
        Me.btnWhat2Root.TabIndex = 39
        Me.btnWhat2Root.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhat2Root.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnWhat2Root.UseVisualStyleBackColor = True
        '
        'btnWhere
        '
        Me.btnWhere.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWhere.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWhere.FlatAppearance.BorderSize = 0
        Me.btnWhere.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWhere.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWhere.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhere.ImageList = Me.imgFileTypes
        Me.btnWhere.Location = New System.Drawing.Point(600, 0)
        Me.btnWhere.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWhere.Name = "btnWhere"
        Me.btnWhere.Size = New System.Drawing.Size(100, 46)
        Me.btnWhere.TabIndex = 31
        Me.btnWhere.Text = "Wr"
        Me.btnWhere.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhere.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnWhere.UseVisualStyleBackColor = True
        '
        'btnWhereRoot
        '
        Me.btnWhereRoot.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWhereRoot.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWhereRoot.FlatAppearance.BorderSize = 0
        Me.btnWhereRoot.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWhereRoot.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWhereRoot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhereRoot.ImageList = Me.imgFileTypes
        Me.btnWhereRoot.Location = New System.Drawing.Point(560, 0)
        Me.btnWhereRoot.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWhereRoot.Name = "btnWhereRoot"
        Me.btnWhereRoot.Size = New System.Drawing.Size(40, 46)
        Me.btnWhereRoot.TabIndex = 38
        Me.btnWhereRoot.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhereRoot.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnWhereRoot.UseVisualStyleBackColor = True
        '
        'btnWhat
        '
        Me.btnWhat.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWhat.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWhat.FlatAppearance.BorderSize = 0
        Me.btnWhat.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWhat.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWhat.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhat.ImageList = Me.imgFileTypes
        Me.btnWhat.Location = New System.Drawing.Point(460, 0)
        Me.btnWhat.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWhat.Name = "btnWhat"
        Me.btnWhat.Size = New System.Drawing.Size(100, 46)
        Me.btnWhat.TabIndex = 30
        Me.btnWhat.Text = "Wa"
        Me.btnWhat.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhat.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnWhat.UseVisualStyleBackColor = True
        '
        'btnWhatRoot
        '
        Me.btnWhatRoot.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWhatRoot.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWhatRoot.FlatAppearance.BorderSize = 0
        Me.btnWhatRoot.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWhatRoot.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWhatRoot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhatRoot.ImageList = Me.imgFileTypes
        Me.btnWhatRoot.Location = New System.Drawing.Point(420, 0)
        Me.btnWhatRoot.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWhatRoot.Name = "btnWhatRoot"
        Me.btnWhatRoot.Size = New System.Drawing.Size(40, 46)
        Me.btnWhatRoot.TabIndex = 37
        Me.btnWhatRoot.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhatRoot.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnWhatRoot.UseVisualStyleBackColor = True
        '
        'btnWho
        '
        Me.btnWho.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWho.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWho.FlatAppearance.BorderSize = 0
        Me.btnWho.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWho.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWho.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWho.ImageList = Me.imgFileTypes
        Me.btnWho.Location = New System.Drawing.Point(320, 0)
        Me.btnWho.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWho.Name = "btnWho"
        Me.btnWho.Size = New System.Drawing.Size(100, 46)
        Me.btnWho.TabIndex = 29
        Me.btnWho.Text = "Wo"
        Me.btnWho.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWho.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnWho.UseVisualStyleBackColor = True
        '
        'btnWhoRoot
        '
        Me.btnWhoRoot.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWhoRoot.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWhoRoot.FlatAppearance.BorderSize = 0
        Me.btnWhoRoot.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWhoRoot.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWhoRoot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhoRoot.ImageList = Me.imgFileTypes
        Me.btnWhoRoot.Location = New System.Drawing.Point(280, 0)
        Me.btnWhoRoot.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWhoRoot.Name = "btnWhoRoot"
        Me.btnWhoRoot.Size = New System.Drawing.Size(40, 46)
        Me.btnWhoRoot.TabIndex = 36
        Me.btnWhoRoot.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhoRoot.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnWhoRoot.UseVisualStyleBackColor = True
        '
        'btnWhen
        '
        Me.btnWhen.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWhen.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWhen.FlatAppearance.BorderSize = 0
        Me.btnWhen.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWhen.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWhen.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhen.ImageList = Me.imgFileTypes
        Me.btnWhen.Location = New System.Drawing.Point(180, 0)
        Me.btnWhen.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWhen.Name = "btnWhen"
        Me.btnWhen.Size = New System.Drawing.Size(100, 46)
        Me.btnWhen.TabIndex = 28
        Me.btnWhen.Tag = ""
        Me.btnWhen.Text = "Wn"
        Me.btnWhen.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhen.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnWhen.UseVisualStyleBackColor = True
        '
        'btnWhenRoot
        '
        Me.btnWhenRoot.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnWhenRoot.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnWhenRoot.FlatAppearance.BorderSize = 0
        Me.btnWhenRoot.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnWhenRoot.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnWhenRoot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhenRoot.ImageList = Me.imgFileTypes
        Me.btnWhenRoot.Location = New System.Drawing.Point(140, 0)
        Me.btnWhenRoot.Margin = New System.Windows.Forms.Padding(6)
        Me.btnWhenRoot.Name = "btnWhenRoot"
        Me.btnWhenRoot.Size = New System.Drawing.Size(40, 46)
        Me.btnWhenRoot.TabIndex = 35
        Me.btnWhenRoot.Tag = ""
        Me.btnWhenRoot.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnWhenRoot.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnWhenRoot.UseVisualStyleBackColor = True
        '
        'btnDataAccess
        '
        Me.btnDataAccess.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnDataAccess.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnDataAccess.FlatAppearance.BorderSize = 0
        Me.btnDataAccess.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDataAccess.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDataAccess.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDataAccess.ImageList = Me.imgFileTypes
        Me.btnDataAccess.Location = New System.Drawing.Point(40, 0)
        Me.btnDataAccess.Margin = New System.Windows.Forms.Padding(6)
        Me.btnDataAccess.Name = "btnDataAccess"
        Me.btnDataAccess.Size = New System.Drawing.Size(100, 46)
        Me.btnDataAccess.TabIndex = 27
        Me.btnDataAccess.Text = "DA"
        Me.btnDataAccess.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDataAccess.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnDataAccess.UseVisualStyleBackColor = True
        '
        'btnDataAccessRoot
        '
        Me.btnDataAccessRoot.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnDataAccessRoot.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnDataAccessRoot.FlatAppearance.BorderSize = 0
        Me.btnDataAccessRoot.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDataAccessRoot.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDataAccessRoot.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDataAccessRoot.ImageList = Me.imgFileTypes
        Me.btnDataAccessRoot.Location = New System.Drawing.Point(0, 0)
        Me.btnDataAccessRoot.Margin = New System.Windows.Forms.Padding(6)
        Me.btnDataAccessRoot.Name = "btnDataAccessRoot"
        Me.btnDataAccessRoot.Size = New System.Drawing.Size(40, 46)
        Me.btnDataAccessRoot.TabIndex = 34
        Me.btnDataAccessRoot.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDataAccessRoot.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnDataAccessRoot.UseVisualStyleBackColor = True
        '
        'pnlSelectorHeader
        '
        Me.pnlSelectorHeader.BackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.pnlSelectorHeader.Controls.Add(Me.lblEmpty)
        Me.pnlSelectorHeader.Controls.Add(Me.lblAboutZaan)
        Me.pnlSelectorHeader.Controls.Add(Me.lblWho2)
        Me.pnlSelectorHeader.Controls.Add(Me.lblWhat2)
        Me.pnlSelectorHeader.Controls.Add(Me.lblWhere)
        Me.pnlSelectorHeader.Controls.Add(Me.lblWhat)
        Me.pnlSelectorHeader.Controls.Add(Me.lblWho)
        Me.pnlSelectorHeader.Controls.Add(Me.lblWhen)
        Me.pnlSelectorHeader.Controls.Add(Me.lblDataAccess)
        Me.pnlSelectorHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlSelectorHeader.Location = New System.Drawing.Point(0, 46)
        Me.pnlSelectorHeader.Margin = New System.Windows.Forms.Padding(6)
        Me.pnlSelectorHeader.Name = "pnlSelectorHeader"
        Me.pnlSelectorHeader.Size = New System.Drawing.Size(1596, 44)
        Me.pnlSelectorHeader.TabIndex = 18
        '
        'lblEmpty
        '
        Me.lblEmpty.BackColor = System.Drawing.SystemColors.Window
        Me.lblEmpty.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblEmpty.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblEmpty.Font = New System.Drawing.Font("Arial Rounded MT Bold", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmpty.ForeColor = System.Drawing.Color.White
        Me.lblEmpty.Image = Global.WindowsApplication1.My.Resources.Resources.grey_empty800
        Me.lblEmpty.Location = New System.Drawing.Point(980, 0)
        Me.lblEmpty.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.lblEmpty.Name = "lblEmpty"
        Me.lblEmpty.Size = New System.Drawing.Size(568, 44)
        Me.lblEmpty.TabIndex = 22
        Me.lblEmpty.Tag = ""
        Me.lblEmpty.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblAboutZaan
        '
        Me.lblAboutZaan.BackColor = System.Drawing.SystemColors.Window
        Me.lblAboutZaan.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblAboutZaan.Dock = System.Windows.Forms.DockStyle.Right
        Me.lblAboutZaan.Font = New System.Drawing.Font("Arial Rounded MT Bold", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAboutZaan.ForeColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.lblAboutZaan.Image = Global.WindowsApplication1.My.Resources.Resources.question
        Me.lblAboutZaan.Location = New System.Drawing.Point(1548, 0)
        Me.lblAboutZaan.Margin = New System.Windows.Forms.Padding(0)
        Me.lblAboutZaan.Name = "lblAboutZaan"
        Me.lblAboutZaan.Size = New System.Drawing.Size(48, 44)
        Me.lblAboutZaan.TabIndex = 27
        Me.lblAboutZaan.Tag = ""
        Me.lblAboutZaan.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblWho2
        '
        Me.lblWho2.BackColor = System.Drawing.SystemColors.Window
        Me.lblWho2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblWho2.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblWho2.Font = New System.Drawing.Font("Arial Rounded MT Bold", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWho2.ForeColor = System.Drawing.Color.White
        Me.lblWho2.Image = Global.WindowsApplication1.My.Resources.Resources._c_5w400
        Me.lblWho2.Location = New System.Drawing.Point(840, 0)
        Me.lblWho2.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.lblWho2.Name = "lblWho2"
        Me.lblWho2.Size = New System.Drawing.Size(140, 44)
        Me.lblWho2.TabIndex = 18
        Me.lblWho2.Tag = ""
        Me.lblWho2.Text = "Wo2"
        Me.lblWho2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblWhat2
        '
        Me.lblWhat2.BackColor = System.Drawing.SystemColors.Window
        Me.lblWhat2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblWhat2.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblWhat2.Font = New System.Drawing.Font("Arial Rounded MT Bold", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWhat2.ForeColor = System.Drawing.Color.White
        Me.lblWhat2.Image = Global.WindowsApplication1.My.Resources.Resources._b_5w400
        Me.lblWhat2.Location = New System.Drawing.Point(700, 0)
        Me.lblWhat2.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.lblWhat2.Name = "lblWhat2"
        Me.lblWhat2.Size = New System.Drawing.Size(140, 44)
        Me.lblWhat2.TabIndex = 15
        Me.lblWhat2.Tag = ""
        Me.lblWhat2.Text = "Wa2"
        Me.lblWhat2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblWhere
        '
        Me.lblWhere.BackColor = System.Drawing.SystemColors.Window
        Me.lblWhere.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblWhere.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblWhere.Font = New System.Drawing.Font("Arial Rounded MT Bold", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWhere.ForeColor = System.Drawing.Color.White
        Me.lblWhere.Image = Global.WindowsApplication1.My.Resources.Resources._e_5w400
        Me.lblWhere.Location = New System.Drawing.Point(560, 0)
        Me.lblWhere.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.lblWhere.Name = "lblWhere"
        Me.lblWhere.Size = New System.Drawing.Size(140, 44)
        Me.lblWhere.TabIndex = 16
        Me.lblWhere.Tag = ""
        Me.lblWhere.Text = "Wr"
        Me.lblWhere.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblWhat
        '
        Me.lblWhat.BackColor = System.Drawing.SystemColors.Window
        Me.lblWhat.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblWhat.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblWhat.Font = New System.Drawing.Font("Arial Rounded MT Bold", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWhat.ForeColor = System.Drawing.Color.White
        Me.lblWhat.Image = Global.WindowsApplication1.My.Resources.Resources._a_5w400
        Me.lblWhat.Location = New System.Drawing.Point(420, 0)
        Me.lblWhat.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.lblWhat.Name = "lblWhat"
        Me.lblWhat.Size = New System.Drawing.Size(140, 44)
        Me.lblWhat.TabIndex = 14
        Me.lblWhat.Tag = ""
        Me.lblWhat.Text = "Wa"
        Me.lblWhat.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblWho
        '
        Me.lblWho.BackColor = System.Drawing.SystemColors.Window
        Me.lblWho.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblWho.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblWho.Font = New System.Drawing.Font("Arial Rounded MT Bold", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWho.ForeColor = System.Drawing.Color.White
        Me.lblWho.Image = Global.WindowsApplication1.My.Resources.Resources._o_5w400
        Me.lblWho.Location = New System.Drawing.Point(280, 0)
        Me.lblWho.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.lblWho.Name = "lblWho"
        Me.lblWho.Size = New System.Drawing.Size(140, 44)
        Me.lblWho.TabIndex = 13
        Me.lblWho.Tag = ""
        Me.lblWho.Text = "Wo"
        Me.lblWho.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblWhen
        '
        Me.lblWhen.BackColor = System.Drawing.SystemColors.Window
        Me.lblWhen.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblWhen.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblWhen.Font = New System.Drawing.Font("Arial Rounded MT Bold", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWhen.ForeColor = System.Drawing.Color.White
        Me.lblWhen.Image = Global.WindowsApplication1.My.Resources.Resources._t_5w400
        Me.lblWhen.Location = New System.Drawing.Point(140, 0)
        Me.lblWhen.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.lblWhen.Name = "lblWhen"
        Me.lblWhen.Size = New System.Drawing.Size(140, 44)
        Me.lblWhen.TabIndex = 19
        Me.lblWhen.Tag = ""
        Me.lblWhen.Text = "Wn"
        Me.lblWhen.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblDataAccess
        '
        Me.lblDataAccess.BackColor = System.Drawing.SystemColors.Window
        Me.lblDataAccess.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblDataAccess.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblDataAccess.Font = New System.Drawing.Font("Arial Rounded MT Bold", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDataAccess.ForeColor = System.Drawing.Color.White
        Me.lblDataAccess.Image = Global.WindowsApplication1.My.Resources.Resources._u_5w400
        Me.lblDataAccess.Location = New System.Drawing.Point(0, 0)
        Me.lblDataAccess.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.lblDataAccess.Name = "lblDataAccess"
        Me.lblDataAccess.Size = New System.Drawing.Size(140, 44)
        Me.lblDataAccess.TabIndex = 20
        Me.lblDataAccess.Tag = ""
        Me.lblDataAccess.Text = "DA"
        Me.lblDataAccess.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lvSelector
        '
        Me.lvSelector.Activation = System.Windows.Forms.ItemActivation.OneClick
        Me.lvSelector.BackColor = System.Drawing.SystemColors.Control
        Me.lvSelector.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lvSelector.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lvSelector.Dock = System.Windows.Forms.DockStyle.Top
        Me.lvSelector.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvSelector.ForeColor = System.Drawing.Color.MidnightBlue
        Me.lvSelector.Location = New System.Drawing.Point(0, 0)
        Me.lvSelector.Margin = New System.Windows.Forms.Padding(6)
        Me.lvSelector.Name = "lvSelector"
        Me.lvSelector.Scrollable = False
        Me.lvSelector.Size = New System.Drawing.Size(1596, 46)
        Me.lvSelector.SmallImageList = Me.imgFileTypes
        Me.lvSelector.TabIndex = 12
        Me.lvSelector.UseCompatibleStateImageBehavior = False
        Me.lvSelector.View = System.Windows.Forms.View.Details
        '
        'pnlSelectorSearch
        '
        Me.pnlSelectorSearch.Controls.Add(Me.pnlSelectorTop)
        Me.pnlSelectorSearch.Controls.Add(Me.pnlSearch)
        Me.pnlSelectorSearch.Controls.Add(Me.pnlSelectorBottom)
        Me.pnlSelectorSearch.Dock = System.Windows.Forms.DockStyle.Left
        Me.pnlSelectorSearch.Location = New System.Drawing.Point(280, 0)
        Me.pnlSelectorSearch.Margin = New System.Windows.Forms.Padding(6)
        Me.pnlSelectorSearch.Name = "pnlSelectorSearch"
        Me.pnlSelectorSearch.Size = New System.Drawing.Size(418, 137)
        Me.pnlSelectorSearch.TabIndex = 19
        '
        'pnlSelectorTop
        '
        Me.pnlSelectorTop.Controls.Add(Me.btnNext)
        Me.pnlSelectorTop.Controls.Add(Me.btnPrev)
        Me.pnlSelectorTop.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlSelectorTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlSelectorTop.Margin = New System.Windows.Forms.Padding(6)
        Me.pnlSelectorTop.Name = "pnlSelectorTop"
        Me.pnlSelectorTop.Padding = New System.Windows.Forms.Padding(14, 2, 2, 6)
        Me.pnlSelectorTop.Size = New System.Drawing.Size(418, 47)
        Me.pnlSelectorTop.TabIndex = 20
        '
        'btnNext
        '
        Me.btnNext.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnNext.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnNext.FlatAppearance.BorderSize = 0
        Me.btnNext.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnNext.ImageList = Me.imgFileTypes
        Me.btnNext.Location = New System.Drawing.Point(62, 2)
        Me.btnNext.Margin = New System.Windows.Forms.Padding(6)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(48, 39)
        Me.btnNext.TabIndex = 26
        Me.btnNext.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnPrev
        '
        Me.btnPrev.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnPrev.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnPrev.FlatAppearance.BorderSize = 0
        Me.btnPrev.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnPrev.ImageList = Me.imgFileTypes
        Me.btnPrev.Location = New System.Drawing.Point(14, 2)
        Me.btnPrev.Margin = New System.Windows.Forms.Padding(6)
        Me.btnPrev.Name = "btnPrev"
        Me.btnPrev.Size = New System.Drawing.Size(48, 39)
        Me.btnPrev.TabIndex = 25
        Me.btnPrev.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnPrev.UseVisualStyleBackColor = True
        '
        'pnlSearch
        '
        Me.pnlSearch.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.pnlSearch.Controls.Add(Me.tbSearch)
        Me.pnlSearch.Controls.Add(Me.lblSearchDoc)
        Me.pnlSearch.Controls.Add(Me.lblSearchFolder)
        Me.pnlSearch.Controls.Add(Me.lblSelectorReset)
        Me.pnlSearch.Controls.Add(Me.lblBookmark)
        Me.pnlSearch.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlSearch.Location = New System.Drawing.Point(0, 47)
        Me.pnlSearch.Margin = New System.Windows.Forms.Padding(6)
        Me.pnlSearch.Name = "pnlSearch"
        Me.pnlSearch.Size = New System.Drawing.Size(418, 44)
        Me.pnlSearch.TabIndex = 10
        '
        'tbSearch
        '
        Me.tbSearch.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.tbSearch.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tbSearch.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbSearch.ForeColor = System.Drawing.Color.White
        Me.tbSearch.Location = New System.Drawing.Point(112, 0)
        Me.tbSearch.Margin = New System.Windows.Forms.Padding(6)
        Me.tbSearch.Name = "tbSearch"
        Me.tbSearch.Size = New System.Drawing.Size(208, 41)
        Me.tbSearch.TabIndex = 23
        Me.tbSearch.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblSearchDoc
        '
        Me.lblSearchDoc.BackColor = System.Drawing.SystemColors.Window
        Me.lblSearchDoc.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblSearchDoc.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblSearchDoc.Font = New System.Drawing.Font("Arial Rounded MT Bold", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSearchDoc.ForeColor = System.Drawing.Color.White
        Me.lblSearchDoc.Image = Global.WindowsApplication1.My.Resources.Resources.search_down
        Me.lblSearchDoc.Location = New System.Drawing.Point(64, 0)
        Me.lblSearchDoc.Margin = New System.Windows.Forms.Padding(0)
        Me.lblSearchDoc.Name = "lblSearchDoc"
        Me.lblSearchDoc.Size = New System.Drawing.Size(48, 44)
        Me.lblSearchDoc.TabIndex = 25
        Me.lblSearchDoc.Tag = ""
        Me.lblSearchDoc.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSearchFolder
        '
        Me.lblSearchFolder.BackColor = System.Drawing.SystemColors.Window
        Me.lblSearchFolder.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblSearchFolder.Dock = System.Windows.Forms.DockStyle.Right
        Me.lblSearchFolder.Font = New System.Drawing.Font("Arial Rounded MT Bold", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSearchFolder.ForeColor = System.Drawing.Color.White
        Me.lblSearchFolder.Image = Global.WindowsApplication1.My.Resources.Resources.search_right
        Me.lblSearchFolder.Location = New System.Drawing.Point(320, 0)
        Me.lblSearchFolder.Margin = New System.Windows.Forms.Padding(0)
        Me.lblSearchFolder.Name = "lblSearchFolder"
        Me.lblSearchFolder.Size = New System.Drawing.Size(48, 44)
        Me.lblSearchFolder.TabIndex = 24
        Me.lblSearchFolder.Tag = ""
        Me.lblSearchFolder.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSelectorReset
        '
        Me.lblSelectorReset.BackColor = System.Drawing.SystemColors.Window
        Me.lblSelectorReset.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblSelectorReset.Dock = System.Windows.Forms.DockStyle.Right
        Me.lblSelectorReset.Font = New System.Drawing.Font("Arial Rounded MT Bold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelectorReset.ForeColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.lblSelectorReset.Image = Global.WindowsApplication1.My.Resources.Resources.reset
        Me.lblSelectorReset.Location = New System.Drawing.Point(368, 0)
        Me.lblSelectorReset.Margin = New System.Windows.Forms.Padding(0)
        Me.lblSelectorReset.Name = "lblSelectorReset"
        Me.lblSelectorReset.Size = New System.Drawing.Size(50, 44)
        Me.lblSelectorReset.TabIndex = 22
        Me.lblSelectorReset.Tag = ""
        Me.lblSelectorReset.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblBookmark
        '
        Me.lblBookmark.BackColor = System.Drawing.SystemColors.Window
        Me.lblBookmark.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblBookmark.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblBookmark.Font = New System.Drawing.Font("Arial Rounded MT Bold", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBookmark.ForeColor = System.Drawing.Color.White
        Me.lblBookmark.Image = Global.WindowsApplication1.My.Resources.Resources.cube_left_end_1b
        Me.lblBookmark.Location = New System.Drawing.Point(0, 0)
        Me.lblBookmark.Margin = New System.Windows.Forms.Padding(0)
        Me.lblBookmark.Name = "lblBookmark"
        Me.lblBookmark.Size = New System.Drawing.Size(64, 44)
        Me.lblBookmark.TabIndex = 26
        Me.lblBookmark.Tag = ""
        Me.lblBookmark.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlSelectorBottom
        '
        Me.pnlSelectorBottom.Controls.Add(Me.btnPanelCube)
        Me.pnlSelectorBottom.Controls.Add(Me.lblSelPagePrev)
        Me.pnlSelectorBottom.Controls.Add(Me.lblSelPage)
        Me.pnlSelectorBottom.Controls.Add(Me.lblSelPageNext)
        Me.pnlSelectorBottom.Controls.Add(Me.btnPanelEmpty)
        Me.pnlSelectorBottom.Controls.Add(Me.btnPanelTree)
        Me.pnlSelectorBottom.Controls.Add(Me.btnMatrix)
        Me.pnlSelectorBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlSelectorBottom.Location = New System.Drawing.Point(0, 91)
        Me.pnlSelectorBottom.Margin = New System.Windows.Forms.Padding(6)
        Me.pnlSelectorBottom.Name = "pnlSelectorBottom"
        Me.pnlSelectorBottom.Padding = New System.Windows.Forms.Padding(16, 4, 4, 4)
        Me.pnlSelectorBottom.Size = New System.Drawing.Size(418, 46)
        Me.pnlSelectorBottom.TabIndex = 19
        '
        'btnPanelCube
        '
        Me.btnPanelCube.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnPanelCube.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnPanelCube.FlatAppearance.BorderSize = 0
        Me.btnPanelCube.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnPanelCube.Image = Global.WindowsApplication1.My.Resources.Resources.pnl_cube_1
        Me.btnPanelCube.Location = New System.Drawing.Point(72, 4)
        Me.btnPanelCube.Margin = New System.Windows.Forms.Padding(6)
        Me.btnPanelCube.Name = "btnPanelCube"
        Me.btnPanelCube.Size = New System.Drawing.Size(44, 38)
        Me.btnPanelCube.TabIndex = 30
        Me.btnPanelCube.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnPanelCube.UseVisualStyleBackColor = False
        '
        'lblSelPagePrev
        '
        Me.lblSelPagePrev.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblSelPagePrev.Dock = System.Windows.Forms.DockStyle.Right
        Me.lblSelPagePrev.Enabled = False
        Me.lblSelPagePrev.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelPagePrev.ForeColor = System.Drawing.Color.Black
        Me.lblSelPagePrev.Location = New System.Drawing.Point(212, 4)
        Me.lblSelPagePrev.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.lblSelPagePrev.Name = "lblSelPagePrev"
        Me.lblSelPagePrev.Size = New System.Drawing.Size(48, 38)
        Me.lblSelPagePrev.TabIndex = 20
        Me.lblSelPagePrev.Text = "<"
        Me.lblSelPagePrev.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblSelPage
        '
        Me.lblSelPage.AutoSize = True
        Me.lblSelPage.Dock = System.Windows.Forms.DockStyle.Right
        Me.lblSelPage.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelPage.ForeColor = System.Drawing.Color.Black
        Me.lblSelPage.Location = New System.Drawing.Point(260, 4)
        Me.lblSelPage.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.lblSelPage.Name = "lblSelPage"
        Me.lblSelPage.Padding = New System.Windows.Forms.Padding(0, 4, 0, 0)
        Me.lblSelPage.Size = New System.Drawing.Size(58, 33)
        Me.lblSelPage.TabIndex = 0
        Me.lblSelPage.Text = "1 / 1"
        Me.lblSelPage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblSelPageNext
        '
        Me.lblSelPageNext.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblSelPageNext.Dock = System.Windows.Forms.DockStyle.Right
        Me.lblSelPageNext.Enabled = False
        Me.lblSelPageNext.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelPageNext.ForeColor = System.Drawing.Color.Black
        Me.lblSelPageNext.Location = New System.Drawing.Point(318, 4)
        Me.lblSelPageNext.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.lblSelPageNext.Name = "lblSelPageNext"
        Me.lblSelPageNext.Size = New System.Drawing.Size(48, 38)
        Me.lblSelPageNext.TabIndex = 19
        Me.lblSelPageNext.Text = ">"
        Me.lblSelPageNext.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnPanelEmpty
        '
        Me.btnPanelEmpty.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnPanelEmpty.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnPanelEmpty.Enabled = False
        Me.btnPanelEmpty.FlatAppearance.BorderSize = 0
        Me.btnPanelEmpty.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnPanelEmpty.Location = New System.Drawing.Point(60, 4)
        Me.btnPanelEmpty.Margin = New System.Windows.Forms.Padding(6)
        Me.btnPanelEmpty.Name = "btnPanelEmpty"
        Me.btnPanelEmpty.Size = New System.Drawing.Size(12, 38)
        Me.btnPanelEmpty.TabIndex = 27
        Me.btnPanelEmpty.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnPanelEmpty.UseVisualStyleBackColor = True
        '
        'btnPanelTree
        '
        Me.btnPanelTree.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnPanelTree.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnPanelTree.FlatAppearance.BorderSize = 0
        Me.btnPanelTree.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnPanelTree.Image = Global.WindowsApplication1.My.Resources.Resources.pnl_tree_1
        Me.btnPanelTree.Location = New System.Drawing.Point(16, 4)
        Me.btnPanelTree.Margin = New System.Windows.Forms.Padding(6)
        Me.btnPanelTree.Name = "btnPanelTree"
        Me.btnPanelTree.Size = New System.Drawing.Size(44, 38)
        Me.btnPanelTree.TabIndex = 26
        Me.btnPanelTree.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnPanelTree.UseVisualStyleBackColor = False
        '
        'btnMatrix
        '
        Me.btnMatrix.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnMatrix.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnMatrix.FlatAppearance.BorderSize = 0
        Me.btnMatrix.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnMatrix.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnMatrix.ForeColor = System.Drawing.Color.MidnightBlue
        Me.btnMatrix.Image = Global.WindowsApplication1.My.Resources.Resources.matrix_1
        Me.btnMatrix.Location = New System.Drawing.Point(366, 4)
        Me.btnMatrix.Margin = New System.Windows.Forms.Padding(6)
        Me.btnMatrix.Name = "btnMatrix"
        Me.btnMatrix.Size = New System.Drawing.Size(48, 38)
        Me.btnMatrix.TabIndex = 29
        Me.btnMatrix.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnMatrix.UseVisualStyleBackColor = True
        '
        'pnlDatabase
        '
        Me.pnlDatabase.Controls.Add(Me.pctZaanLogo)
        Me.pnlDatabase.Dock = System.Windows.Forms.DockStyle.Left
        Me.pnlDatabase.Location = New System.Drawing.Point(0, 0)
        Me.pnlDatabase.Margin = New System.Windows.Forms.Padding(6)
        Me.pnlDatabase.Name = "pnlDatabase"
        Me.pnlDatabase.Size = New System.Drawing.Size(280, 137)
        Me.pnlDatabase.TabIndex = 20
        '
        'pctZaanLogo
        '
        Me.pctZaanLogo.ContextMenuStrip = Me.cmsSelector
        Me.pctZaanLogo.Cursor = System.Windows.Forms.Cursors.Default
        Me.pctZaanLogo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pctZaanLogo.Image = Global.WindowsApplication1.My.Resources.Resources.zaan_cube_50i
        Me.pctZaanLogo.InitialImage = Global.WindowsApplication1.My.Resources.Resources.zaan_cube_50i
        Me.pctZaanLogo.Location = New System.Drawing.Point(0, 0)
        Me.pctZaanLogo.Margin = New System.Windows.Forms.Padding(0)
        Me.pctZaanLogo.Name = "pctZaanLogo"
        Me.pctZaanLogo.Size = New System.Drawing.Size(280, 137)
        Me.pctZaanLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.pctZaanLogo.TabIndex = 0
        Me.pctZaanLogo.TabStop = False
        '
        'cmsSelector
        '
        Me.cmsSelector.BackColor = System.Drawing.SystemColors.Menu
        Me.cmsSelector.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.cmsSelector.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiSelectorTreeView, Me.tsmiSelectorViewer, Me.tsmiSelectorImport, Me.tssSelector0, Me.tsmiSelectorTreeLock, Me.tssSelector1, Me.tsmiSelectorNightMode, Me.tssSelector2, Me.tsmiSelectorChangeDb, Me.tsmiSelectorChangeDbImage, Me.tssSelector3, Me.tsmiSelectorCheckDb, Me.tssSelector4, Me.tsmiSelectorExportDBwin, Me.tsmiSelectorImportDBwin, Me.tsmiSelectorAutoImport, Me.tssSelector5, Me.tsmiSelectorCloseDb})
        Me.cmsSelector.Name = "cmsSelector"
        Me.cmsSelector.Size = New System.Drawing.Size(466, 496)
        '
        'tsmiSelectorTreeView
        '
        Me.tsmiSelectorTreeView.Name = "tsmiSelectorTreeView"
        Me.tsmiSelectorTreeView.Size = New System.Drawing.Size(465, 38)
        Me.tsmiSelectorTreeView.Text = "Tree view panel visible"
        Me.tsmiSelectorTreeView.Visible = False
        '
        'tsmiSelectorViewer
        '
        Me.tsmiSelectorViewer.Name = "tsmiSelectorViewer"
        Me.tsmiSelectorViewer.Size = New System.Drawing.Size(465, 38)
        Me.tsmiSelectorViewer.Text = "Viewer panel visible"
        Me.tsmiSelectorViewer.Visible = False
        '
        'tsmiSelectorImport
        '
        Me.tsmiSelectorImport.Name = "tsmiSelectorImport"
        Me.tsmiSelectorImport.Size = New System.Drawing.Size(465, 38)
        Me.tsmiSelectorImport.Text = "Import panel visible"
        Me.tsmiSelectorImport.Visible = False
        '
        'tssSelector0
        '
        Me.tssSelector0.Name = "tssSelector0"
        Me.tssSelector0.Size = New System.Drawing.Size(462, 6)
        Me.tssSelector0.Visible = False
        '
        'tsmiSelectorTreeLock
        '
        Me.tsmiSelectorTreeLock.Name = "tsmiSelectorTreeLock"
        Me.tsmiSelectorTreeLock.Size = New System.Drawing.Size(465, 38)
        Me.tsmiSelectorTreeLock.Text = "Tree view locked"
        '
        'tssSelector1
        '
        Me.tssSelector1.Name = "tssSelector1"
        Me.tssSelector1.Size = New System.Drawing.Size(462, 6)
        '
        'tsmiSelectorNightMode
        '
        Me.tsmiSelectorNightMode.Name = "tsmiSelectorNightMode"
        Me.tsmiSelectorNightMode.Size = New System.Drawing.Size(465, 38)
        Me.tsmiSelectorNightMode.Text = "Night mode"
        '
        'tssSelector2
        '
        Me.tssSelector2.Name = "tssSelector2"
        Me.tssSelector2.Size = New System.Drawing.Size(462, 6)
        '
        'tsmiSelectorChangeDb
        '
        Me.tsmiSelectorChangeDb.Name = "tsmiSelectorChangeDb"
        Me.tsmiSelectorChangeDb.Size = New System.Drawing.Size(465, 38)
        Me.tsmiSelectorChangeDb.Text = "Change database"
        '
        'tsmiSelectorChangeDbImage
        '
        Me.tsmiSelectorChangeDbImage.Name = "tsmiSelectorChangeDbImage"
        Me.tsmiSelectorChangeDbImage.Size = New System.Drawing.Size(465, 38)
        Me.tsmiSelectorChangeDbImage.Text = "Change database image"
        '
        'tssSelector3
        '
        Me.tssSelector3.Name = "tssSelector3"
        Me.tssSelector3.Size = New System.Drawing.Size(462, 6)
        '
        'tsmiSelectorCheckDb
        '
        Me.tsmiSelectorCheckDb.Name = "tsmiSelectorCheckDb"
        Me.tsmiSelectorCheckDb.Size = New System.Drawing.Size(465, 38)
        Me.tsmiSelectorCheckDb.Text = "Check database"
        '
        'tssSelector4
        '
        Me.tssSelector4.Name = "tssSelector4"
        Me.tssSelector4.Size = New System.Drawing.Size(462, 6)
        '
        'tsmiSelectorExportDBwin
        '
        Me.tsmiSelectorExportDBwin.Name = "tsmiSelectorExportDBwin"
        Me.tsmiSelectorExportDBwin.Size = New System.Drawing.Size(465, 38)
        Me.tsmiSelectorExportDBwin.Text = "Export database to Windows"
        '
        'tsmiSelectorImportDBwin
        '
        Me.tsmiSelectorImportDBwin.Enabled = False
        Me.tsmiSelectorImportDBwin.Name = "tsmiSelectorImportDBwin"
        Me.tsmiSelectorImportDBwin.Size = New System.Drawing.Size(465, 38)
        Me.tsmiSelectorImportDBwin.Text = "Import database from Windows..."
        Me.tsmiSelectorImportDBwin.Visible = False
        '
        'tsmiSelectorAutoImport
        '
        Me.tsmiSelectorAutoImport.Enabled = False
        Me.tsmiSelectorAutoImport.Name = "tsmiSelectorAutoImport"
        Me.tsmiSelectorAutoImport.Size = New System.Drawing.Size(465, 38)
        Me.tsmiSelectorAutoImport.Text = "ZAAN cubes automatic import"
        Me.tsmiSelectorAutoImport.Visible = False
        '
        'tssSelector5
        '
        Me.tssSelector5.Name = "tssSelector5"
        Me.tssSelector5.Size = New System.Drawing.Size(462, 6)
        '
        'tsmiSelectorCloseDb
        '
        Me.tsmiSelectorCloseDb.Name = "tsmiSelectorCloseDb"
        Me.tsmiSelectorCloseDb.Size = New System.Drawing.Size(465, 38)
        Me.tsmiSelectorCloseDb.Text = "Close database"
        '
        'tcFolders
        '
        Me.tcFolders.Controls.Add(Me.TabPage1)
        Me.tcFolders.Controls.Add(Me.TabPage2)
        Me.tcFolders.Controls.Add(Me.TabPage3)
        Me.tcFolders.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tcFolders.Font = New System.Drawing.Font("Franklin Gothic Book", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tcFolders.ImageList = Me.imgFileTypes
        Me.tcFolders.Location = New System.Drawing.Point(0, 0)
        Me.tcFolders.Margin = New System.Windows.Forms.Padding(6)
        Me.tcFolders.Name = "tcFolders"
        Me.tcFolders.SelectedIndex = 0
        Me.tcFolders.ShowToolTips = True
        Me.tcFolders.Size = New System.Drawing.Size(2332, 627)
        Me.tcFolders.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.lvOut)
        Me.TabPage1.ForeColor = System.Drawing.Color.MidnightBlue
        Me.TabPage1.Location = New System.Drawing.Point(8, 52)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(6)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(2316, 567)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Import"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'lvOut
        '
        Me.lvOut.AllowDrop = True
        Me.lvOut.BackColor = System.Drawing.SystemColors.Window
        Me.lvOut.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lvOut.ContextMenuStrip = Me.cmsLvOut
        Me.lvOut.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvOut.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvOut.ForeColor = System.Drawing.Color.MidnightBlue
        Me.lvOut.FullRowSelect = True
        Me.lvOut.LabelEdit = True
        Me.lvOut.Location = New System.Drawing.Point(0, 0)
        Me.lvOut.Margin = New System.Windows.Forms.Padding(6)
        Me.lvOut.Name = "lvOut"
        Me.lvOut.Size = New System.Drawing.Size(2316, 567)
        Me.lvOut.SmallImageList = Me.imgFileTypes
        Me.lvOut.TabIndex = 7
        Me.lvOut.UseCompatibleStateImageBehavior = False
        Me.lvOut.View = System.Windows.Forms.View.Details
        '
        'cmsLvOut
        '
        Me.cmsLvOut.BackColor = System.Drawing.SystemColors.Menu
        Me.cmsLvOut.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.cmsLvOut.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiLvOutChangeFolder, Me.tsmiLvOutFoldersVisible, Me.tssLvOut1, Me.tsmiLvOutRefresh, Me.tsmiLvOutSelAll, Me.tsmiLvOutUndoMove, Me.tssLvOut2, Me.tsmiLvOutCut, Me.tsmiLvOutCopy, Me.tsmiLvOutPaste, Me.tssLvOut3, Me.tsmiLvOutRename, Me.tsmiLvOutDelete, Me.tssLvOut4, Me.tsmiLvOutMoveIn, Me.tsmiLvOutAutoFile, Me.tssLvOut5, Me.tsmiLvOutImport})
        Me.cmsLvOut.Name = "cmsLvOut"
        Me.cmsLvOut.Size = New System.Drawing.Size(339, 528)
        '
        'tsmiLvOutChangeFolder
        '
        Me.tsmiLvOutChangeFolder.Enabled = False
        Me.tsmiLvOutChangeFolder.Name = "tsmiLvOutChangeFolder"
        Me.tsmiLvOutChangeFolder.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvOutChangeFolder.Text = "Change folder"
        '
        'tsmiLvOutFoldersVisible
        '
        Me.tsmiLvOutFoldersVisible.Enabled = False
        Me.tsmiLvOutFoldersVisible.Name = "tsmiLvOutFoldersVisible"
        Me.tsmiLvOutFoldersVisible.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvOutFoldersVisible.Text = "Folders tree visible"
        '
        'tssLvOut1
        '
        Me.tssLvOut1.Name = "tssLvOut1"
        Me.tssLvOut1.Size = New System.Drawing.Size(335, 6)
        '
        'tsmiLvOutRefresh
        '
        Me.tsmiLvOutRefresh.Name = "tsmiLvOutRefresh"
        Me.tsmiLvOutRefresh.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvOutRefresh.Text = "Refresh"
        '
        'tsmiLvOutSelAll
        '
        Me.tsmiLvOutSelAll.Name = "tsmiLvOutSelAll"
        Me.tsmiLvOutSelAll.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
        Me.tsmiLvOutSelAll.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvOutSelAll.Text = "Select all"
        '
        'tsmiLvOutUndoMove
        '
        Me.tsmiLvOutUndoMove.Enabled = False
        Me.tsmiLvOutUndoMove.Name = "tsmiLvOutUndoMove"
        Me.tsmiLvOutUndoMove.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Z), System.Windows.Forms.Keys)
        Me.tsmiLvOutUndoMove.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvOutUndoMove.Text = "Undo move"
        '
        'tssLvOut2
        '
        Me.tssLvOut2.Name = "tssLvOut2"
        Me.tssLvOut2.Size = New System.Drawing.Size(335, 6)
        '
        'tsmiLvOutCut
        '
        Me.tsmiLvOutCut.Enabled = False
        Me.tsmiLvOutCut.Name = "tsmiLvOutCut"
        Me.tsmiLvOutCut.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
        Me.tsmiLvOutCut.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvOutCut.Text = "Cut"
        '
        'tsmiLvOutCopy
        '
        Me.tsmiLvOutCopy.Enabled = False
        Me.tsmiLvOutCopy.Name = "tsmiLvOutCopy"
        Me.tsmiLvOutCopy.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.tsmiLvOutCopy.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvOutCopy.Text = "Copy"
        '
        'tsmiLvOutPaste
        '
        Me.tsmiLvOutPaste.Enabled = False
        Me.tsmiLvOutPaste.Name = "tsmiLvOutPaste"
        Me.tsmiLvOutPaste.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.V), System.Windows.Forms.Keys)
        Me.tsmiLvOutPaste.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvOutPaste.Text = "Paste"
        '
        'tssLvOut3
        '
        Me.tssLvOut3.Name = "tssLvOut3"
        Me.tssLvOut3.Size = New System.Drawing.Size(335, 6)
        '
        'tsmiLvOutRename
        '
        Me.tsmiLvOutRename.Enabled = False
        Me.tsmiLvOutRename.Name = "tsmiLvOutRename"
        Me.tsmiLvOutRename.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvOutRename.Text = "Rename"
        '
        'tsmiLvOutDelete
        '
        Me.tsmiLvOutDelete.Enabled = False
        Me.tsmiLvOutDelete.Image = Global.WindowsApplication1.My.Resources.Resources.delete_lg
        Me.tsmiLvOutDelete.Name = "tsmiLvOutDelete"
        Me.tsmiLvOutDelete.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvOutDelete.Text = "Delete"
        '
        'tssLvOut4
        '
        Me.tssLvOut4.Name = "tssLvOut4"
        Me.tssLvOut4.Size = New System.Drawing.Size(335, 6)
        '
        'tsmiLvOutMoveIn
        '
        Me.tsmiLvOutMoveIn.Enabled = False
        Me.tsmiLvOutMoveIn.Image = Global.WindowsApplication1.My.Resources.Resources.up_lg
        Me.tsmiLvOutMoveIn.Name = "tsmiLvOutMoveIn"
        Me.tsmiLvOutMoveIn.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvOutMoveIn.Text = "Move in"
        '
        'tsmiLvOutAutoFile
        '
        Me.tsmiLvOutAutoFile.Enabled = False
        Me.tsmiLvOutAutoFile.Name = "tsmiLvOutAutoFile"
        Me.tsmiLvOutAutoFile.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvOutAutoFile.Text = "File automatically"
        '
        'tssLvOut5
        '
        Me.tssLvOut5.Name = "tssLvOut5"
        Me.tssLvOut5.Size = New System.Drawing.Size(335, 6)
        '
        'tsmiLvOutImport
        '
        Me.tsmiLvOutImport.Enabled = False
        Me.tsmiLvOutImport.Name = "tsmiLvOutImport"
        Me.tsmiLvOutImport.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvOutImport.Text = "Import ZAAN cube(s)"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.SplitContainer4)
        Me.TabPage2.ForeColor = System.Drawing.Color.MidnightBlue
        Me.TabPage2.Location = New System.Drawing.Point(8, 52)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(6)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(2316, 567)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Scan"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer4.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer4.Margin = New System.Windows.Forms.Padding(6)
        Me.SplitContainer4.Name = "SplitContainer4"
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.trvInput)
        Me.SplitContainer4.Size = New System.Drawing.Size(2316, 567)
        Me.SplitContainer4.SplitterDistance = 460
        Me.SplitContainer4.SplitterWidth = 8
        Me.SplitContainer4.TabIndex = 0
        '
        'trvInput
        '
        Me.trvInput.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.trvInput.ContextMenuStrip = Me.cmsTrvInput
        Me.trvInput.Dock = System.Windows.Forms.DockStyle.Fill
        Me.trvInput.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.trvInput.FullRowSelect = True
        Me.trvInput.HotTracking = True
        Me.trvInput.ImageIndex = 0
        Me.trvInput.ImageList = Me.imgFileTypes
        Me.trvInput.Indent = 16
        Me.trvInput.ItemHeight = 18
        Me.trvInput.Location = New System.Drawing.Point(0, 0)
        Me.trvInput.Margin = New System.Windows.Forms.Padding(6)
        Me.trvInput.Name = "trvInput"
        Me.trvInput.SelectedImageIndex = 0
        Me.trvInput.Size = New System.Drawing.Size(460, 567)
        Me.trvInput.TabIndex = 2
        '
        'cmsTrvInput
        '
        Me.cmsTrvInput.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.cmsTrvInput.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiTrvInputChangeRoot, Me.tssTrvInput0, Me.tsmiTrvInputRefresh, Me.tssTrvInput1, Me.tsmiTrvInputDelete})
        Me.cmsTrvInput.Name = "cmsTrvInput"
        Me.cmsTrvInput.Size = New System.Drawing.Size(319, 130)
        '
        'tsmiTrvInputChangeRoot
        '
        Me.tsmiTrvInputChangeRoot.Name = "tsmiTrvInputChangeRoot"
        Me.tsmiTrvInputChangeRoot.Size = New System.Drawing.Size(318, 38)
        Me.tsmiTrvInputChangeRoot.Text = "Change root folder"
        '
        'tssTrvInput0
        '
        Me.tssTrvInput0.Name = "tssTrvInput0"
        Me.tssTrvInput0.Size = New System.Drawing.Size(315, 6)
        '
        'tsmiTrvInputRefresh
        '
        Me.tsmiTrvInputRefresh.Name = "tsmiTrvInputRefresh"
        Me.tsmiTrvInputRefresh.Size = New System.Drawing.Size(318, 38)
        Me.tsmiTrvInputRefresh.Text = "Refresh"
        '
        'tssTrvInput1
        '
        Me.tssTrvInput1.Name = "tssTrvInput1"
        Me.tssTrvInput1.Size = New System.Drawing.Size(315, 6)
        '
        'tsmiTrvInputDelete
        '
        Me.tsmiTrvInputDelete.Enabled = False
        Me.tsmiTrvInputDelete.Image = Global.WindowsApplication1.My.Resources.Resources.delete_lg
        Me.tsmiTrvInputDelete.Name = "tsmiTrvInputDelete"
        Me.tsmiTrvInputDelete.ShortcutKeys = System.Windows.Forms.Keys.Delete
        Me.tsmiTrvInputDelete.Size = New System.Drawing.Size(318, 38)
        Me.tsmiTrvInputDelete.Text = "Delete"
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.lvTemp)
        Me.TabPage3.ForeColor = System.Drawing.Color.ForestGreen
        Me.TabPage3.Location = New System.Drawing.Point(8, 52)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(6)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(2316, 567)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Copy"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'lvTemp
        '
        Me.lvTemp.AllowDrop = True
        Me.lvTemp.BackColor = System.Drawing.SystemColors.Window
        Me.lvTemp.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lvTemp.ContextMenuStrip = Me.cmsLvTemp
        Me.lvTemp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvTemp.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvTemp.ForeColor = System.Drawing.Color.ForestGreen
        Me.lvTemp.FullRowSelect = True
        Me.lvTemp.LabelEdit = True
        Me.lvTemp.Location = New System.Drawing.Point(0, 0)
        Me.lvTemp.Margin = New System.Windows.Forms.Padding(6)
        Me.lvTemp.Name = "lvTemp"
        Me.lvTemp.Size = New System.Drawing.Size(2316, 567)
        Me.lvTemp.SmallImageList = Me.imgFileTypes
        Me.lvTemp.TabIndex = 0
        Me.lvTemp.UseCompatibleStateImageBehavior = False
        Me.lvTemp.View = System.Windows.Forms.View.Details
        '
        'cmsLvTemp
        '
        Me.cmsLvTemp.BackColor = System.Drawing.SystemColors.Menu
        Me.cmsLvTemp.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.cmsLvTemp.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiLvTempRefresh, Me.tsmiLvTempSelAll, Me.tsmiLvTempUndoMove, Me.tssLvTemp1, Me.tsmiLvTempCut, Me.tsmiLvTempCopy, Me.tsmiLvTempPaste, Me.tssLvTemp2, Me.tsmiLvTempResizePicture, Me.tstLvTempResizePercentage, Me.tssLvTemp3, Me.tsmiLvTempRename, Me.tsmiLvTempDelete, Me.tssLvTemp4, Me.tsmiLvTempSend, Me.tssLvTemp5, Me.tsmiLvTempImport})
        Me.cmsLvTemp.Name = "cmsLvTemp"
        Me.cmsLvTemp.Size = New System.Drawing.Size(339, 493)
        '
        'tsmiLvTempRefresh
        '
        Me.tsmiLvTempRefresh.Name = "tsmiLvTempRefresh"
        Me.tsmiLvTempRefresh.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvTempRefresh.Text = "Refresh"
        '
        'tsmiLvTempSelAll
        '
        Me.tsmiLvTempSelAll.Name = "tsmiLvTempSelAll"
        Me.tsmiLvTempSelAll.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
        Me.tsmiLvTempSelAll.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvTempSelAll.Text = "Select all"
        '
        'tsmiLvTempUndoMove
        '
        Me.tsmiLvTempUndoMove.Enabled = False
        Me.tsmiLvTempUndoMove.Name = "tsmiLvTempUndoMove"
        Me.tsmiLvTempUndoMove.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Z), System.Windows.Forms.Keys)
        Me.tsmiLvTempUndoMove.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvTempUndoMove.Text = "Undo move"
        '
        'tssLvTemp1
        '
        Me.tssLvTemp1.Name = "tssLvTemp1"
        Me.tssLvTemp1.Size = New System.Drawing.Size(335, 6)
        '
        'tsmiLvTempCut
        '
        Me.tsmiLvTempCut.Enabled = False
        Me.tsmiLvTempCut.Name = "tsmiLvTempCut"
        Me.tsmiLvTempCut.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
        Me.tsmiLvTempCut.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvTempCut.Text = "Cut"
        '
        'tsmiLvTempCopy
        '
        Me.tsmiLvTempCopy.Enabled = False
        Me.tsmiLvTempCopy.Name = "tsmiLvTempCopy"
        Me.tsmiLvTempCopy.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.tsmiLvTempCopy.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvTempCopy.Text = "Copy"
        '
        'tsmiLvTempPaste
        '
        Me.tsmiLvTempPaste.Enabled = False
        Me.tsmiLvTempPaste.Name = "tsmiLvTempPaste"
        Me.tsmiLvTempPaste.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.V), System.Windows.Forms.Keys)
        Me.tsmiLvTempPaste.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvTempPaste.Text = "Paste"
        '
        'tssLvTemp2
        '
        Me.tssLvTemp2.Name = "tssLvTemp2"
        Me.tssLvTemp2.Size = New System.Drawing.Size(335, 6)
        '
        'tsmiLvTempResizePicture
        '
        Me.tsmiLvTempResizePicture.Enabled = False
        Me.tsmiLvTempResizePicture.Name = "tsmiLvTempResizePicture"
        Me.tsmiLvTempResizePicture.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvTempResizePicture.Text = "Resize pictures (%)"
        '
        'tstLvTempResizePercentage
        '
        Me.tstLvTempResizePercentage.Name = "tstLvTempResizePercentage"
        Me.tstLvTempResizePercentage.Size = New System.Drawing.Size(100, 39)
        Me.tstLvTempResizePercentage.Text = "20"
        Me.tstLvTempResizePercentage.TextBoxTextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tssLvTemp3
        '
        Me.tssLvTemp3.Name = "tssLvTemp3"
        Me.tssLvTemp3.Size = New System.Drawing.Size(335, 6)
        '
        'tsmiLvTempRename
        '
        Me.tsmiLvTempRename.Enabled = False
        Me.tsmiLvTempRename.Name = "tsmiLvTempRename"
        Me.tsmiLvTempRename.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvTempRename.Text = "Rename"
        '
        'tsmiLvTempDelete
        '
        Me.tsmiLvTempDelete.Enabled = False
        Me.tsmiLvTempDelete.Image = Global.WindowsApplication1.My.Resources.Resources.delete_lg
        Me.tsmiLvTempDelete.Name = "tsmiLvTempDelete"
        Me.tsmiLvTempDelete.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvTempDelete.Text = "Delete"
        '
        'tssLvTemp4
        '
        Me.tssLvTemp4.Name = "tssLvTemp4"
        Me.tssLvTemp4.Size = New System.Drawing.Size(335, 6)
        '
        'tsmiLvTempSend
        '
        Me.tsmiLvTempSend.Enabled = False
        Me.tsmiLvTempSend.Image = Global.WindowsApplication1.My.Resources.Resources.mail_lg
        Me.tsmiLvTempSend.Name = "tsmiLvTempSend"
        Me.tsmiLvTempSend.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvTempSend.Text = "Open mail"
        '
        'tssLvTemp5
        '
        Me.tssLvTemp5.Name = "tssLvTemp5"
        Me.tssLvTemp5.Size = New System.Drawing.Size(335, 6)
        '
        'tsmiLvTempImport
        '
        Me.tsmiLvTempImport.Enabled = False
        Me.tsmiLvTempImport.Name = "tsmiLvTempImport"
        Me.tsmiLvTempImport.Size = New System.Drawing.Size(338, 38)
        Me.tsmiLvTempImport.Text = "Import ZAAN cube(s)"
        '
        'imgPanels
        '
        Me.imgPanels.ImageStream = CType(resources.GetObject("imgPanels.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgPanels.TransparentColor = System.Drawing.Color.Transparent
        Me.imgPanels.Images.SetKeyName(0, "left")
        Me.imgPanels.Images.SetKeyName(1, "right")
        '
        'tmrLvTempLostFocus
        '
        Me.tmrLvTempLostFocus.Interval = 1000
        '
        'tmrLicenseCheck
        '
        Me.tmrLicenseCheck.Interval = 1000
        '
        'fswData
        '
        Me.fswData.EnableRaisingEvents = True
        Me.fswData.IncludeSubdirectories = True
        Me.fswData.SynchronizingObject = Me
        '
        'tmrLvInDisplay
        '
        Me.tmrLvInDisplay.Interval = 1000
        '
        'tmrTrvWDisplay
        '
        Me.tmrTrvWDisplay.Interval = 1000
        '
        'fswXin
        '
        Me.fswXin.EnableRaisingEvents = True
        Me.fswXin.NotifyFilter = CType((System.IO.NotifyFilters.FileName Or System.IO.NotifyFilters.LastWrite), System.IO.NotifyFilters)
        Me.fswXin.SynchronizingObject = Me
        '
        'fswTree
        '
        Me.fswTree.EnableRaisingEvents = True
        Me.fswTree.NotifyFilter = CType((System.IO.NotifyFilters.FileName Or System.IO.NotifyFilters.LastWrite), System.IO.NotifyFilters)
        Me.fswTree.SynchronizingObject = Me
        '
        'tmrCubeImport
        '
        Me.tmrCubeImport.Interval = 1000
        '
        'fswInput
        '
        Me.fswInput.EnableRaisingEvents = True
        Me.fswInput.NotifyFilter = CType((System.IO.NotifyFilters.FileName Or System.IO.NotifyFilters.LastWrite), System.IO.NotifyFilters)
        Me.fswInput.SynchronizingObject = Me
        '
        'tmrSlideShow
        '
        '
        'tmrSlideShowEffect
        '
        '
        'tmrVideoProgress
        '
        '
        'fswZaan
        '
        Me.fswZaan.EnableRaisingEvents = True
        Me.fswZaan.IncludeSubdirectories = True
        Me.fswZaan.SynchronizingObject = Me
        '
        'cmsLvInWhos
        '
        Me.cmsLvInWhos.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.cmsLvInWhos.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsmiLvInWhosMultiple, Me.tssLvInWhos0, Me.tsmiLvInWhosDelete, Me.tscbLvInWhosList})
        Me.cmsLvInWhos.Name = "cmsLvInWhos"
        Me.cmsLvInWhos.Size = New System.Drawing.Size(286, 130)
        '
        'tsmiLvInWhosMultiple
        '
        Me.tsmiLvInWhosMultiple.Name = "tsmiLvInWhosMultiple"
        Me.tsmiLvInWhosMultiple.Size = New System.Drawing.Size(285, 38)
        Me.tsmiLvInWhosMultiple.Text = "Multiple Who(s)"
        '
        'tssLvInWhos0
        '
        Me.tssLvInWhos0.Name = "tssLvInWhos0"
        Me.tssLvInWhos0.Size = New System.Drawing.Size(282, 6)
        '
        'tsmiLvInWhosDelete
        '
        Me.tsmiLvInWhosDelete.Enabled = False
        Me.tsmiLvInWhosDelete.Name = "tsmiLvInWhosDelete"
        Me.tsmiLvInWhosDelete.Size = New System.Drawing.Size(285, 38)
        Me.tsmiLvInWhosDelete.Text = "Delete Who :"
        '
        'tscbLvInWhosList
        '
        Me.tscbLvInWhosList.Enabled = False
        Me.tscbLvInWhosList.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.tscbLvInWhosList.Name = "tscbLvInWhosList"
        Me.tscbLvInWhosList.Size = New System.Drawing.Size(121, 40)
        '
        'tmrSelectorList
        '
        Me.tmrSelectorList.Interval = 3000
        '
        'ofdZaanFile
        '
        Me.ofdZaanFile.FileName = "OpenFileDialog1"
        '
        'cmsSelChildren
        '
        Me.cmsSelChildren.AutoSize = False
        Me.cmsSelChildren.BackColor = System.Drawing.SystemColors.Control
        Me.cmsSelChildren.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.cmsSelChildren.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1, Me.ToolStripMenuItem2, Me.ToolStripMenuItem3})
        Me.cmsSelChildren.Name = "cmsSelChildren"
        Me.cmsSelChildren.Size = New System.Drawing.Size(153, 92)
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(190, 38)
        Me.ToolStripMenuItem1.Text = "Child 1"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(190, 38)
        Me.ToolStripMenuItem2.Text = "Child 2"
        '
        'ToolStripMenuItem3
        '
        Me.ToolStripMenuItem3.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Child31ToolStripMenuItem, Me.Child32ToolStripMenuItem})
        Me.ToolStripMenuItem3.Name = "ToolStripMenuItem3"
        Me.ToolStripMenuItem3.Size = New System.Drawing.Size(190, 38)
        Me.ToolStripMenuItem3.Text = "Child 3"
        '
        'Child31ToolStripMenuItem
        '
        Me.Child31ToolStripMenuItem.Name = "Child31ToolStripMenuItem"
        Me.Child31ToolStripMenuItem.Size = New System.Drawing.Size(203, 38)
        Me.Child31ToolStripMenuItem.Text = "Child 31"
        '
        'Child32ToolStripMenuItem
        '
        Me.Child32ToolStripMenuItem.Name = "Child32ToolStripMenuItem"
        Me.Child32ToolStripMenuItem.Size = New System.Drawing.Size(203, 38)
        Me.Child32ToolStripMenuItem.Text = "Child 32"
        '
        'tmrSelChildren
        '
        Me.tmrSelChildren.Interval = 500
        '
        'cmsSelBookmarks
        '
        Me.cmsSelBookmarks.BackColor = System.Drawing.SystemColors.Control
        Me.cmsSelBookmarks.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.cmsSelBookmarks.Name = "cmsSelBookmarks"
        Me.cmsSelBookmarks.Size = New System.Drawing.Size(86, 4)
        '
        'tmrDbExport
        '
        Me.tmrDbExport.Interval = 60000
        '
        'frmZaan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlDark
        Me.ClientSize = New System.Drawing.Size(2352, 1042)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(6)
        Me.MinimumSize = New System.Drawing.Size(674, 415)
        Me.Name = "frmZaan"
        Me.Padding = New System.Windows.Forms.Padding(10)
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Zaan"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.tcCubes.ResumeLayout(False)
        Me.TabPage4.ResumeLayout(False)
        Me.pnlCube.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.cmsTrvW.ResumeLayout(False)
        Me.cmsBookmark.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer3.ResumeLayout(False)
        Me.pnlList.ResumeLayout(False)
        Me.cmsLvIn.ResumeLayout(False)
        CType(Me.pctLvIn, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlZoomVideo.ResumeLayout(False)
        Me.pnlVideoControl.ResumeLayout(False)
        Me.pnlZoom.ResumeLayout(False)
        CType(Me.pctZoom, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cmsPctZoom.ResumeLayout(False)
        Me.cmsPctZoom.PerformLayout()
        Me.pnlSlideControl.ResumeLayout(False)
        CType(Me.pctZoom2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlListTop.ResumeLayout(False)
        Me.pnlSelectorNav.ResumeLayout(False)
        Me.pnlSelectorHeaderBottom.ResumeLayout(False)
        Me.pnlSelectorHeaderTop.ResumeLayout(False)
        Me.pnlSelectorHeader.ResumeLayout(False)
        Me.pnlSelectorSearch.ResumeLayout(False)
        Me.pnlSelectorTop.ResumeLayout(False)
        Me.pnlSearch.ResumeLayout(False)
        Me.pnlSearch.PerformLayout()
        Me.pnlSelectorBottom.ResumeLayout(False)
        Me.pnlSelectorBottom.PerformLayout()
        Me.pnlDatabase.ResumeLayout(False)
        CType(Me.pctZaanLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cmsSelector.ResumeLayout(False)
        Me.tcFolders.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.cmsLvOut.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer4.ResumeLayout(False)
        Me.cmsTrvInput.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.cmsLvTemp.ResumeLayout(False)
        Me.cmsLvTemp.PerformLayout()
        CType(Me.fswData, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fswXin, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fswTree, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fswInput, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fswZaan, System.ComponentModel.ISupportInitialize).EndInit()
        Me.cmsLvInWhos.ResumeLayout(False)
        Me.cmsSelChildren.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents pctZaanLogo As System.Windows.Forms.PictureBox
    Friend WithEvents trvW As System.Windows.Forms.TreeView
    Friend WithEvents lvIn As System.Windows.Forms.ListView
    Friend WithEvents lvOut As System.Windows.Forms.ListView
    Friend WithEvents wbDoc As System.Windows.Forms.WebBrowser
    Friend WithEvents tlTip As System.Windows.Forms.ToolTip
    Friend WithEvents fbdZaanPath As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents imgFileTypes As System.Windows.Forms.ImageList
    Friend WithEvents imgLargeIcons As System.Windows.Forms.ImageList
    Friend WithEvents pctZoom As System.Windows.Forms.PictureBox
    Friend WithEvents lvTemp As System.Windows.Forms.ListView
    Friend WithEvents pctLvIn As System.Windows.Forms.PictureBox
    Friend WithEvents tmrLvTempLostFocus As System.Windows.Forms.Timer
    Friend WithEvents tmrLicenseCheck As System.Windows.Forms.Timer
    Friend WithEvents fswData As System.IO.FileSystemWatcher
    Friend WithEvents tmrLvInDisplay As System.Windows.Forms.Timer
    Friend WithEvents imgPanels As System.Windows.Forms.ImageList
    Friend WithEvents cmsLvIn As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tsmiLvInSelAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInCut As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInCopy As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInPaste As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInRename As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInDelete As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssLvIn2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents lblDocName As System.Windows.Forms.Label
    Friend WithEvents cmsLvOut As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tsmiLvOutSelAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssLvOut2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiLvOutCut As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvOutCopy As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvOutPaste As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssLvOut3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiLvOutRename As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvOutDelete As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmsLvTemp As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tsmiLvTempSelAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssLvTemp1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiLvTempCut As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvTempCopy As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvTempPaste As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssLvTemp2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiLvTempRename As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvTempDelete As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssLvTemp3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiLvTempSend As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmsTrvW As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tsmiTrvWAdd As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiTrvWRename As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiTrvWDelete As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssLvIn3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents cmsSelector As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tssLvIn4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiLvInCopyToZC As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInMoveOut As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssLvOut4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiLvOutMoveIn As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents pnlList As System.Windows.Forms.Panel
    Friend WithEvents tssLvOut1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents lblWhat As System.Windows.Forms.Label
    Friend WithEvents lblWho As System.Windows.Forms.Label
    Friend WithEvents lblWhere As System.Windows.Forms.Label
    Friend WithEvents lblWhat2 As System.Windows.Forms.Label
    Friend WithEvents lblWhen As System.Windows.Forms.Label
    Friend WithEvents lblWho2 As System.Windows.Forms.Label
    Friend WithEvents cmsBookmark As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tsmiSelTabsAdd As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiSelTabsDelete As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInAutoRename As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssTrvW1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tmrTrvWDisplay As System.Windows.Forms.Timer
    Friend WithEvents tsmiLvTempResizePicture As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tstLvTempResizePercentage As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents tssLvTemp4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiLvInExport As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents fswXin As System.IO.FileSystemWatcher
    Friend WithEvents fswTree As System.IO.FileSystemWatcher
    Friend WithEvents tsmiSelectorExportDBwin As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents pgbZaan As System.Windows.Forms.ProgressBar
    Friend WithEvents tsmiSelectorAutoImport As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tmrCubeImport As System.Windows.Forms.Timer
    Friend WithEvents trvInput As System.Windows.Forms.TreeView
    Friend WithEvents tsmiLvOutFoldersVisible As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmsTrvInput As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tsmiTrvInputChangeRoot As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiTrvInputDelete As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssTrvInput0 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents fswInput As System.IO.FileSystemWatcher
    Friend WithEvents lblDataAccess As System.Windows.Forms.Label
    Friend WithEvents tsmiLvInUndoMove As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmsPctZoom As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tsmiPctZoomSlideShow As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssPctZoom1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiPctZoomInterval As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tstbPctZoom As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents tmrSlideShow As System.Windows.Forms.Timer
    Friend WithEvents pctZoom2 As System.Windows.Forms.PictureBox
    Friend WithEvents tmrSlideShowEffect As System.Windows.Forms.Timer
    Friend WithEvents pnlSlideControl As System.Windows.Forms.Panel
    Friend WithEvents btnSlidePrevious As System.Windows.Forms.Button
    Friend WithEvents btnSlidePlayPause As System.Windows.Forms.Button
    Friend WithEvents btnSlideNext As System.Windows.Forms.Button
    Friend WithEvents lblSlideNb As System.Windows.Forms.Label
    Friend WithEvents btnSlideStop As System.Windows.Forms.Button
    Friend WithEvents pnlZoom As System.Windows.Forms.Panel
    Friend WithEvents lblZoomName As System.Windows.Forms.Label
    Friend WithEvents btnSlideDelete As System.Windows.Forms.Button
    Friend WithEvents btnSlideRotateRight As System.Windows.Forms.Button
    Friend WithEvents btnSlideRotateLeft As System.Windows.Forms.Button
    Friend WithEvents pnlZoomVideo As System.Windows.Forms.Panel
    Friend WithEvents pnlVideoControl As System.Windows.Forms.Panel
    Friend WithEvents btnVideoBegin As System.Windows.Forms.Button
    Friend WithEvents btnVideoEnd As System.Windows.Forms.Button
    Friend WithEvents btnVideoPlayPause As System.Windows.Forms.Button
    Friend WithEvents btnVideoStop As System.Windows.Forms.Button
    Friend WithEvents pgbVideo As System.Windows.Forms.ProgressBar
    Friend WithEvents tmrVideoProgress As System.Windows.Forms.Timer
    Friend WithEvents fswZaan As System.IO.FileSystemWatcher
    Friend WithEvents tsmiLvOutAutoFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvOutUndoMove As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvTempUndoMove As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInRefresh As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvOutRefresh As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvTempRefresh As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmsLvInWhos As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tsmiLvInWhosMultiple As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInWhosDelete As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tscbLvInWhosList As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents tssLvInWhos0 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiTrvWRefresh As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssTrvW2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tbSearch As System.Windows.Forms.TextBox
    Friend WithEvents tmrSelectorList As System.Windows.Forms.Timer
    Friend WithEvents lstInDir As System.Windows.Forms.ListBox
    Friend WithEvents lvSelector As System.Windows.Forms.ListView
    Friend WithEvents pnlSelectorHeader As System.Windows.Forms.Panel
    Friend WithEvents pnlSelectorNav As System.Windows.Forms.Panel
    Friend WithEvents lvBookmark As System.Windows.Forms.ListView
    Friend WithEvents tsmiSelTabsRefresh As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInExpRefTable As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssTrvW3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiTrvWImport As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ofdTxtImport As System.Windows.Forms.OpenFileDialog
    Friend WithEvents tsmiTrvWExport As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssSelector1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiSelectorImportDBwin As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssSelector2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiSelectorTreeLock As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents pnlSearch As System.Windows.Forms.Panel
    Friend WithEvents btnPrev As System.Windows.Forms.Button
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents lstPage As System.Windows.Forms.ListBox
    Friend WithEvents tsmiLvInCopyPath As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssLvIn5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiLvInOpenFolder As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiSelectorCheckDb As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssLvIn6 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiLvInExpNameTable As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssLvTemp5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiLvTempImport As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssLvOut5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiLvOutImport As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiSelectorChangeDbImage As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssSelector3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiSelectorChangeDb As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ofdZaanFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents pnlSelectorBottom As System.Windows.Forms.Panel
    Friend WithEvents lblSelPagePrev As System.Windows.Forms.Label
    Friend WithEvents lblSelPage As System.Windows.Forms.Label
    Friend WithEvents lblSelPageNext As System.Windows.Forms.Label
    Friend WithEvents tsmiLvInDocPerPage As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssLvIn1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiLvInDpp10 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInDpp25 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInDpp50 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInDpp100 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInDpp250 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInDpp500 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInDpp1000 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tcFolders As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents tsmiLvOutChangeFolder As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiTrvInputRefresh As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssTrvInput1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents pnlSelectorSearch As System.Windows.Forms.Panel
    Friend WithEvents btnCubeTube As System.Windows.Forms.Button
    Friend WithEvents tcCubes As System.Windows.Forms.TabControl
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents pnlListTop As System.Windows.Forms.Panel
    Friend WithEvents pnlCube As System.Windows.Forms.Panel
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents pnlDatabase As System.Windows.Forms.Panel
    Friend WithEvents tsmiSelectorTreeView As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiSelectorViewer As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiSelectorImport As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lblSelectorReset As System.Windows.Forms.Label
    Friend WithEvents lblSearchDoc As System.Windows.Forms.Label
    Friend WithEvents lblSearchFolder As System.Windows.Forms.Label
    Friend WithEvents lblBookmark As System.Windows.Forms.Label
    Friend WithEvents lblEmpty As System.Windows.Forms.Label
    Friend WithEvents lblAboutZaan As System.Windows.Forms.Label
    Friend WithEvents tsmiLvInImageMode As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmsSelChildren As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem3 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tmrSelChildren As System.Windows.Forms.Timer
    Friend WithEvents Child31ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Child32ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmsSelBookmarks As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents tssSelector0 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents pnlSelectorHeaderTop As System.Windows.Forms.Panel
    Friend WithEvents btnWho2 As System.Windows.Forms.Button
    Friend WithEvents btnWhat2 As System.Windows.Forms.Button
    Friend WithEvents btnWhere As System.Windows.Forms.Button
    Friend WithEvents btnWhat As System.Windows.Forms.Button
    Friend WithEvents btnWho As System.Windows.Forms.Button
    Friend WithEvents btnWhen As System.Windows.Forms.Button
    Friend WithEvents btnDataAccess As System.Windows.Forms.Button
    Friend WithEvents tsmiSelectorNightMode As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssSelector4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents pnlSelectorTop As System.Windows.Forms.Panel
    Friend WithEvents btnPanelTree As System.Windows.Forms.Button
    Friend WithEvents btnPanelEmpty As System.Windows.Forms.Button
    Friend WithEvents btnPanelView As System.Windows.Forms.Button
    Friend WithEvents btnWho2Root As System.Windows.Forms.Button
    Friend WithEvents btnWhat2Root As System.Windows.Forms.Button
    Friend WithEvents btnWhereRoot As System.Windows.Forms.Button
    Friend WithEvents btnWhatRoot As System.Windows.Forms.Button
    Friend WithEvents btnWhoRoot As System.Windows.Forms.Button
    Friend WithEvents btnWhenRoot As System.Windows.Forms.Button
    Friend WithEvents btnDataAccessRoot As System.Windows.Forms.Button
    Friend WithEvents pnlSelectorHeaderBottom As System.Windows.Forms.Panel
    Friend WithEvents lvMatrix As System.Windows.Forms.ListView
    Friend WithEvents btnWho2Axis As System.Windows.Forms.Button
    Friend WithEvents btnWhat2Axis As System.Windows.Forms.Button
    Friend WithEvents btnWhereAxis As System.Windows.Forms.Button
    Friend WithEvents btnWhatAxis As System.Windows.Forms.Button
    Friend WithEvents btnWhoAxis As System.Windows.Forms.Button
    Friend WithEvents btnWhenAxis As System.Windows.Forms.Button
    Friend WithEvents btnDataAccessAxis As System.Windows.Forms.Button
    Friend WithEvents btnMatrix As System.Windows.Forms.Button
    Friend WithEvents tsmiLvInNew As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInNewNote As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInNewText As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInNewPres As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmiLvInNewSpSh As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssLvIn7 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiTrvWAddExternal As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnPanelImport As System.Windows.Forms.Button
    Friend WithEvents tmrDbExport As System.Windows.Forms.Timer
    Friend WithEvents btnToday As System.Windows.Forms.Button
    Friend WithEvents btnDataAccessBlank As System.Windows.Forms.Button
    Friend WithEvents tssSelTabs1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiSelTabsDefault As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssSelector5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents tsmiSelectorCloseDb As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnPanelCube As System.Windows.Forms.Button
End Class
