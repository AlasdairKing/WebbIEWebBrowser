<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
    <System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub
    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
	Protected Overrides Sub Dispose(ByVal Disposing As Boolean)
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
    Public WithEvents tmrRefreshIfNotChange As System.Windows.Forms.Timer
    Public WithEvents tmrSetFocus As System.Windows.Forms.Timer
    Public WithEvents tmrProcessAfterLoad As System.Windows.Forms.Timer
    Public cdlgOpen As System.Windows.Forms.OpenFileDialog
    Public WithEvents tmrBusyAnimation As System.Windows.Forms.Timer
    Public WithEvents _staMain_Panel1 As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents _staMain_Panel2 As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents staMain As System.Windows.Forms.StatusStrip
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.tmrRefreshIfNotChange = New System.Windows.Forms.Timer(Me.components)
        Me.tmrSetFocus = New System.Windows.Forms.Timer(Me.components)
        Me.tmrProcessAfterLoad = New System.Windows.Forms.Timer(Me.components)
        Me.cdlgOpen = New System.Windows.Forms.OpenFileDialog()
        Me.tmrBusyAnimation = New System.Windows.Forms.Timer(Me.components)
        Me.staMain = New System.Windows.Forms.StatusStrip()
        Me._staMain_Panel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me._staMain_Panel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.mnuBar = New System.Windows.Forms.ToolStripSeparator()
        Me.MenuStripMain = New System.Windows.Forms.MenuStrip()
        Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileOpen = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileSave = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFilePrint = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFileExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEdit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditCopy = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditPaste = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditSelectall = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditFindtext = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditFindnext = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuEditWebsearch = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuView = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuViewCroppage = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuViewRSSNewsFeed = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuViewIE = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuViewSource = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFavorites = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFavoritesAdd = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFavoritesShowfavorites = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFavoritesGotofavorites = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuFavoritesOrganise = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuNavigate = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNavigateBack = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNavigateStop = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNavigateHome = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNavigateRefresh = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNavigateForward = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNavigateGotoheadline = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuNavigateGotoform = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLinks = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLinksNextlink = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLinksPreviouslink = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLinksSkipDown = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLinksSkipup = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLinksViewlinks = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLinksDownloadlinktarget = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLinksFollowlinkaddress = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOptions = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOptionsFont = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOptionsColour = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOptionsSetHomepage = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuOptionsOptions = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelpManual = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelpTeamviewer = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelpWebbIEOrg = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelpAbout = New System.Windows.Forms.ToolStripMenuItem()
        Me.fdMain = New System.Windows.Forms.FontDialog()
        Me.MainToolStrip = New System.Windows.Forms.ToolStrip()
        Me.btnBack = New System.Windows.Forms.ToolStripButton()
        Me.btnStop = New System.Windows.Forms.ToolStripButton()
        Me.btnRefresh = New System.Windows.Forms.ToolStripButton()
        Me.btnHome = New System.Windows.Forms.ToolStripButton()
        Me.btnSearch = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnCrop = New System.Windows.Forms.ToolStripButton()
        Me.btnHeading = New System.Windows.Forms.ToolStripButton()
        Me.picBusy = New System.Windows.Forms.ToolStripLabel()
        Me.btnSkiplinks = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnRSS = New System.Windows.Forms.ToolStripButton()
        Me.pnlMain = New System.Windows.Forms.Panel()
        Me.txtText = New System.Windows.Forms.RichTextBox()
        Me.cboAddress = New System.Windows.Forms.ComboBox()
        Me.lblAltDForAddress = New System.Windows.Forms.Label()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.tmrDelayLoadBookmarks = New System.Windows.Forms.Timer(Me.components)
        Me.staMain.SuspendLayout()
        Me.MenuStripMain.SuspendLayout()
        Me.MainToolStrip.SuspendLayout()
        Me.pnlMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'tmrRefreshIfNotChange
        '
        Me.tmrRefreshIfNotChange.Interval = 2000
        '
        'tmrSetFocus
        '
        Me.tmrSetFocus.Interval = 50
        '
        'tmrProcessAfterLoad
        '
        Me.tmrProcessAfterLoad.Interval = 50
        '
        'tmrBusyAnimation
        '
        Me.tmrBusyAnimation.Enabled = True
        Me.tmrBusyAnimation.Interval = 300
        '
        'staMain
        '
        Me.staMain.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.staMain.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.staMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me._staMain_Panel1, Me._staMain_Panel2})
        Me.staMain.Location = New System.Drawing.Point(0, 713)
        Me.staMain.Name = "staMain"
        Me.staMain.Padding = New System.Windows.Forms.Padding(2, 0, 24, 0)
        Me.staMain.Size = New System.Drawing.Size(1362, 28)
        Me.staMain.TabIndex = 2
        Me.staMain.Tag = "frmMain.staMain"
        '
        '_staMain_Panel1
        '
        Me._staMain_Panel1.AutoSize = False
        Me._staMain_Panel1.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me._staMain_Panel1.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me._staMain_Panel1.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._staMain_Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me._staMain_Panel1.Name = "_staMain_Panel1"
        Me._staMain_Panel1.Size = New System.Drawing.Size(140, 28)
        Me._staMain_Panel1.Tag = "frmMain.staMain.txtBusy"
        Me._staMain_Panel1.Text = "Idle"
        Me._staMain_Panel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._staMain_Panel1.ToolTipText = " What WebbIE is doing "
        '
        '_staMain_Panel2
        '
        Me._staMain_Panel2.AutoSize = False
        Me._staMain_Panel2.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me._staMain_Panel2.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
        Me._staMain_Panel2.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._staMain_Panel2.Margin = New System.Windows.Forms.Padding(0)
        Me._staMain_Panel2.Name = "_staMain_Panel2"
        Me._staMain_Panel2.Size = New System.Drawing.Size(1196, 28)
        Me._staMain_Panel2.Spring = True
        Me._staMain_Panel2.Tag = "frmMain.staMain.txtTitle"
        Me._staMain_Panel2.Text = "Blank"
        Me._staMain_Panel2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me._staMain_Panel2.ToolTipText = " Title of this web page "
        '
        'mnuBar
        '
        Me.mnuBar.Name = "mnuBar"
        Me.mnuBar.Size = New System.Drawing.Size(166, 6)
        '
        'MenuStripMain
        '
        Me.MenuStripMain.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MenuStripMain.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.MenuStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuEdit, Me.mnuView, Me.mnuFavorites, Me.mnuNavigate, Me.mnuLinks, Me.mnuOptions, Me.mnuHelp})
        Me.MenuStripMain.Location = New System.Drawing.Point(0, 0)
        Me.MenuStripMain.Name = "MenuStripMain"
        Me.MenuStripMain.Padding = New System.Windows.Forms.Padding(10, 4, 0, 4)
        Me.MenuStripMain.Size = New System.Drawing.Size(1362, 63)
        Me.MenuStripMain.TabIndex = 0
        Me.MenuStripMain.Text = "MenuStripMain"
        '
        'mnuFile
        '
        Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFileOpen, Me.mnuFileSave, Me.mnuFilePrint, Me.mnuFileExit})
        Me.mnuFile.Name = "mnuFile"
        Me.mnuFile.Size = New System.Drawing.Size(91, 55)
        Me.mnuFile.Text = "&File"
        '
        'mnuFileOpen
        '
        Me.mnuFileOpen.Name = "mnuFileOpen"
        Me.mnuFileOpen.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.mnuFileOpen.Size = New System.Drawing.Size(358, 56)
        Me.mnuFileOpen.Text = "&Open"
        '
        'mnuFileSave
        '
        Me.mnuFileSave.Name = "mnuFileSave"
        Me.mnuFileSave.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.mnuFileSave.Size = New System.Drawing.Size(358, 56)
        Me.mnuFileSave.Text = "&Save"
        '
        'mnuFilePrint
        '
        Me.mnuFilePrint.Name = "mnuFilePrint"
        Me.mnuFilePrint.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P), System.Windows.Forms.Keys)
        Me.mnuFilePrint.Size = New System.Drawing.Size(358, 56)
        Me.mnuFilePrint.Text = "&Print"
        '
        'mnuFileExit
        '
        Me.mnuFileExit.Name = "mnuFileExit"
        Me.mnuFileExit.Size = New System.Drawing.Size(358, 56)
        Me.mnuFileExit.Text = "E&xit"
        '
        'mnuEdit
        '
        Me.mnuEdit.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuEditCopy, Me.mnuEditPaste, Me.mnuEditSelectall, Me.mnuEditFindtext, Me.mnuEditFindnext, Me.mnuEditWebsearch})
        Me.mnuEdit.Name = "mnuEdit"
        Me.mnuEdit.Size = New System.Drawing.Size(97, 55)
        Me.mnuEdit.Text = "&Edit"
        '
        'mnuEditCopy
        '
        Me.mnuEditCopy.Name = "mnuEditCopy"
        Me.mnuEditCopy.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.mnuEditCopy.Size = New System.Drawing.Size(566, 56)
        Me.mnuEditCopy.Text = "&Copy"
        '
        'mnuEditPaste
        '
        Me.mnuEditPaste.Name = "mnuEditPaste"
        Me.mnuEditPaste.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.V), System.Windows.Forms.Keys)
        Me.mnuEditPaste.Size = New System.Drawing.Size(566, 56)
        Me.mnuEditPaste.Text = "&Paste"
        '
        'mnuEditSelectall
        '
        Me.mnuEditSelectall.Name = "mnuEditSelectall"
        Me.mnuEditSelectall.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)
        Me.mnuEditSelectall.Size = New System.Drawing.Size(566, 56)
        Me.mnuEditSelectall.Text = "Select &all"
        '
        'mnuEditFindtext
        '
        Me.mnuEditFindtext.Name = "mnuEditFindtext"
        Me.mnuEditFindtext.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)
        Me.mnuEditFindtext.Size = New System.Drawing.Size(566, 56)
        Me.mnuEditFindtext.Text = "&Find (on this page)"
        '
        'mnuEditFindnext
        '
        Me.mnuEditFindnext.Name = "mnuEditFindnext"
        Me.mnuEditFindnext.ShortcutKeys = System.Windows.Forms.Keys.F3
        Me.mnuEditFindnext.Size = New System.Drawing.Size(566, 56)
        Me.mnuEditFindnext.Text = "Find &next"
        '
        'mnuEditWebsearch
        '
        Me.mnuEditWebsearch.Name = "mnuEditWebsearch"
        Me.mnuEditWebsearch.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.W), System.Windows.Forms.Keys)
        Me.mnuEditWebsearch.Size = New System.Drawing.Size(566, 56)
        Me.mnuEditWebsearch.Text = "&Websearch"
        '
        'mnuView
        '
        Me.mnuView.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuViewCroppage, Me.mnuViewRSSNewsFeed, Me.mnuViewIE, Me.mnuViewSource})
        Me.mnuView.Name = "mnuView"
        Me.mnuView.Size = New System.Drawing.Size(114, 55)
        Me.mnuView.Text = "&View"
        '
        'mnuViewCroppage
        '
        Me.mnuViewCroppage.Name = "mnuViewCroppage"
        Me.mnuViewCroppage.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.K), System.Windows.Forms.Keys)
        Me.mnuViewCroppage.Size = New System.Drawing.Size(587, 56)
        Me.mnuViewCroppage.Text = "&Crop page"
        '
        'mnuViewRSSNewsFeed
        '
        Me.mnuViewRSSNewsFeed.Name = "mnuViewRSSNewsFeed"
        Me.mnuViewRSSNewsFeed.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.E), System.Windows.Forms.Keys)
        Me.mnuViewRSSNewsFeed.Size = New System.Drawing.Size(587, 56)
        Me.mnuViewRSSNewsFeed.Text = "View &RSS news feed"
        '
        'mnuViewIE
        '
        Me.mnuViewIE.Name = "mnuViewIE"
        Me.mnuViewIE.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.I), System.Windows.Forms.Keys)
        Me.mnuViewIE.Size = New System.Drawing.Size(587, 56)
        Me.mnuViewIE.Text = "Webpage"
        '
        'mnuViewSource
        '
        Me.mnuViewSource.Name = "mnuViewSource"
        Me.mnuViewSource.Size = New System.Drawing.Size(587, 56)
        Me.mnuViewSource.Text = "S&ource"
        '
        'mnuFavorites
        '
        Me.mnuFavorites.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFavoritesAdd, Me.mnuFavoritesShowfavorites, Me.mnuFavoritesGotofavorites, Me.mnuFavoritesOrganise, Me.ToolStripMenuItem1})
        Me.mnuFavorites.Name = "mnuFavorites"
        Me.mnuFavorites.Size = New System.Drawing.Size(182, 55)
        Me.mnuFavorites.Text = "F&avorites"
        '
        'mnuFavoritesAdd
        '
        Me.mnuFavoritesAdd.Name = "mnuFavoritesAdd"
        Me.mnuFavoritesAdd.Size = New System.Drawing.Size(501, 56)
        Me.mnuFavoritesAdd.Text = "&Add to favorites"
        '
        'mnuFavoritesShowfavorites
        '
        Me.mnuFavoritesShowfavorites.Name = "mnuFavoritesShowfavorites"
        Me.mnuFavoritesShowfavorites.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.B), System.Windows.Forms.Keys)
        Me.mnuFavoritesShowfavorites.Size = New System.Drawing.Size(501, 56)
        Me.mnuFavoritesShowfavorites.Text = "Show favorites"
        '
        'mnuFavoritesGotofavorites
        '
        Me.mnuFavoritesGotofavorites.Name = "mnuFavoritesGotofavorites"
        Me.mnuFavoritesGotofavorites.Size = New System.Drawing.Size(501, 56)
        Me.mnuFavoritesGotofavorites.Text = "&Go to favorites"
        '
        'mnuFavoritesOrganise
        '
        Me.mnuFavoritesOrganise.Name = "mnuFavoritesOrganise"
        Me.mnuFavoritesOrganise.Size = New System.Drawing.Size(501, 56)
        Me.mnuFavoritesOrganise.Text = "&Organise favorites"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(498, 6)
        '
        'mnuNavigate
        '
        Me.mnuNavigate.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuNavigateBack, Me.mnuNavigateStop, Me.mnuNavigateHome, Me.mnuNavigateRefresh, Me.mnuNavigateForward, Me.mnuNavigateGotoheadline, Me.mnuNavigateGotoform})
        Me.mnuNavigate.Name = "mnuNavigate"
        Me.mnuNavigate.Size = New System.Drawing.Size(182, 55)
        Me.mnuNavigate.Text = "&Navigate"
        '
        'mnuNavigateBack
        '
        Me.mnuNavigateBack.Name = "mnuNavigateBack"
        Me.mnuNavigateBack.ShortcutKeyDisplayString = "Backspace"
        Me.mnuNavigateBack.Size = New System.Drawing.Size(499, 56)
        Me.mnuNavigateBack.Text = "&Back"
        '
        'mnuNavigateStop
        '
        Me.mnuNavigateStop.Name = "mnuNavigateStop"
        Me.mnuNavigateStop.ShortcutKeyDisplayString = "Escape"
        Me.mnuNavigateStop.Size = New System.Drawing.Size(499, 56)
        Me.mnuNavigateStop.Text = "&Stop"
        '
        'mnuNavigateHome
        '
        Me.mnuNavigateHome.Name = "mnuNavigateHome"
        Me.mnuNavigateHome.ShortcutKeyDisplayString = "Alt+Home"
        Me.mnuNavigateHome.Size = New System.Drawing.Size(499, 56)
        Me.mnuNavigateHome.Text = "&Home"
        '
        'mnuNavigateRefresh
        '
        Me.mnuNavigateRefresh.Name = "mnuNavigateRefresh"
        Me.mnuNavigateRefresh.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
        Me.mnuNavigateRefresh.Size = New System.Drawing.Size(499, 56)
        Me.mnuNavigateRefresh.Text = "&Refresh"
        '
        'mnuNavigateForward
        '
        Me.mnuNavigateForward.Name = "mnuNavigateForward"
        Me.mnuNavigateForward.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.Right), System.Windows.Forms.Keys)
        Me.mnuNavigateForward.Size = New System.Drawing.Size(499, 56)
        Me.mnuNavigateForward.Text = "&Forward"
        '
        'mnuNavigateGotoheadline
        '
        Me.mnuNavigateGotoheadline.Name = "mnuNavigateGotoheadline"
        Me.mnuNavigateGotoheadline.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.H), System.Windows.Forms.Keys)
        Me.mnuNavigateGotoheadline.Size = New System.Drawing.Size(499, 56)
        Me.mnuNavigateGotoheadline.Text = "Goto &headline"
        '
        'mnuNavigateGotoform
        '
        Me.mnuNavigateGotoform.Name = "mnuNavigateGotoform"
        Me.mnuNavigateGotoform.ShortcutKeys = System.Windows.Forms.Keys.F6
        Me.mnuNavigateGotoform.Size = New System.Drawing.Size(499, 56)
        Me.mnuNavigateGotoform.Text = "Goto &form"
        '
        'mnuLinks
        '
        Me.mnuLinks.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuLinksNextlink, Me.mnuLinksPreviouslink, Me.mnuLinksSkipDown, Me.mnuLinksSkipup, Me.mnuLinksViewlinks, Me.mnuLinksDownloadlinktarget, Me.mnuLinksFollowlinkaddress})
        Me.mnuLinks.Name = "mnuLinks"
        Me.mnuLinks.Size = New System.Drawing.Size(118, 55)
        Me.mnuLinks.Text = "&Links"
        '
        'mnuLinksNextlink
        '
        Me.mnuLinksNextlink.Name = "mnuLinksNextlink"
        Me.mnuLinksNextlink.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Tab), System.Windows.Forms.Keys)
        Me.mnuLinksNextlink.Size = New System.Drawing.Size(883, 56)
        Me.mnuLinksNextlink.Text = "Next link"
        '
        'mnuLinksPreviouslink
        '
        Me.mnuLinksPreviouslink.Name = "mnuLinksPreviouslink"
        Me.mnuLinksPreviouslink.ShortcutKeys = CType(((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Shift) _
            Or System.Windows.Forms.Keys.Tab), System.Windows.Forms.Keys)
        Me.mnuLinksPreviouslink.Size = New System.Drawing.Size(883, 56)
        Me.mnuLinksPreviouslink.Text = "Previous link"
        '
        'mnuLinksSkipDown
        '
        Me.mnuLinksSkipDown.Name = "mnuLinksSkipDown"
        Me.mnuLinksSkipDown.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Down), System.Windows.Forms.Keys)
        Me.mnuLinksSkipDown.Size = New System.Drawing.Size(883, 56)
        Me.mnuLinksSkipDown.Text = "Skip links down"
        '
        'mnuLinksSkipup
        '
        Me.mnuLinksSkipup.Name = "mnuLinksSkipup"
        Me.mnuLinksSkipup.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Up), System.Windows.Forms.Keys)
        Me.mnuLinksSkipup.Size = New System.Drawing.Size(883, 56)
        Me.mnuLinksSkipup.Text = "Skip links up"
        '
        'mnuLinksViewlinks
        '
        Me.mnuLinksViewlinks.Name = "mnuLinksViewlinks"
        Me.mnuLinksViewlinks.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.L), System.Windows.Forms.Keys)
        Me.mnuLinksViewlinks.Size = New System.Drawing.Size(883, 56)
        Me.mnuLinksViewlinks.Text = "View links"
        '
        'mnuLinksDownloadlinktarget
        '
        Me.mnuLinksDownloadlinktarget.Name = "mnuLinksDownloadlinktarget"
        Me.mnuLinksDownloadlinktarget.ShortcutKeyDisplayString = "Shift+Enter"
        Me.mnuLinksDownloadlinktarget.Size = New System.Drawing.Size(883, 56)
        Me.mnuLinksDownloadlinktarget.Text = "Download link target to Desktop"
        '
        'mnuLinksFollowlinkaddress
        '
        Me.mnuLinksFollowlinkaddress.Name = "mnuLinksFollowlinkaddress"
        Me.mnuLinksFollowlinkaddress.ShortcutKeyDisplayString = "Ctrl+Enter"
        Me.mnuLinksFollowlinkaddress.Size = New System.Drawing.Size(883, 56)
        Me.mnuLinksFollowlinkaddress.Text = "Follow link address (not click)"
        '
        'mnuOptions
        '
        Me.mnuOptions.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuOptionsFont, Me.mnuOptionsColour, Me.mnuOptionsSetHomepage, Me.mnuOptionsOptions})
        Me.mnuOptions.Name = "mnuOptions"
        Me.mnuOptions.Size = New System.Drawing.Size(167, 55)
        Me.mnuOptions.Text = "&Options"
        '
        'mnuOptionsFont
        '
        Me.mnuOptionsFont.Name = "mnuOptionsFont"
        Me.mnuOptionsFont.Size = New System.Drawing.Size(382, 56)
        Me.mnuOptionsFont.Text = "Font..."
        '
        'mnuOptionsColour
        '
        Me.mnuOptionsColour.Name = "mnuOptionsColour"
        Me.mnuOptionsColour.Size = New System.Drawing.Size(382, 56)
        Me.mnuOptionsColour.Text = "&Colour..."
        '
        'mnuOptionsSetHomepage
        '
        Me.mnuOptionsSetHomepage.Name = "mnuOptionsSetHomepage"
        Me.mnuOptionsSetHomepage.Size = New System.Drawing.Size(382, 56)
        Me.mnuOptionsSetHomepage.Text = "Set home page"
        '
        'mnuOptionsOptions
        '
        Me.mnuOptionsOptions.Name = "mnuOptionsOptions"
        Me.mnuOptionsOptions.Size = New System.Drawing.Size(382, 56)
        Me.mnuOptionsOptions.Text = "&Options..."
        '
        'mnuHelp
        '
        Me.mnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuHelpManual, Me.mnuHelpTeamviewer, Me.mnuHelpWebbIEOrg, Me.mnuHelpAbout})
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(112, 55)
        Me.mnuHelp.Text = "&Help"
        '
        'mnuHelpManual
        '
        Me.mnuHelpManual.Name = "mnuHelpManual"
        Me.mnuHelpManual.ShortcutKeys = System.Windows.Forms.Keys.F1
        Me.mnuHelpManual.Size = New System.Drawing.Size(452, 56)
        Me.mnuHelpManual.Text = "Manual"
        '
        'mnuHelpTeamviewer
        '
        Me.mnuHelpTeamviewer.Name = "mnuHelpTeamviewer"
        Me.mnuHelpTeamviewer.Size = New System.Drawing.Size(452, 56)
        Me.mnuHelpTeamviewer.Text = "TeamViewer"
        '
        'mnuHelpWebbIEOrg
        '
        Me.mnuHelpWebbIEOrg.Name = "mnuHelpWebbIEOrg"
        Me.mnuHelpWebbIEOrg.Size = New System.Drawing.Size(452, 56)
        Me.mnuHelpWebbIEOrg.Text = "www.webbie.org.uk"
        '
        'mnuHelpAbout
        '
        Me.mnuHelpAbout.Name = "mnuHelpAbout"
        Me.mnuHelpAbout.Size = New System.Drawing.Size(452, 56)
        Me.mnuHelpAbout.Text = "About"
        '
        'fdMain
        '
        Me.fdMain.Color = System.Drawing.SystemColors.ControlText
        '
        'MainToolStrip
        '
        Me.MainToolStrip.AutoSize = False
        Me.MainToolStrip.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.MainToolStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnBack, Me.btnStop, Me.btnRefresh, Me.btnHome, Me.btnSearch, Me.ToolStripSeparator1, Me.btnCrop, Me.btnHeading, Me.picBusy, Me.btnSkiplinks, Me.ToolStripSeparator2, Me.btnRSS})
        Me.MainToolStrip.Location = New System.Drawing.Point(0, 63)
        Me.MainToolStrip.Name = "MainToolStrip"
        Me.MainToolStrip.Size = New System.Drawing.Size(1362, 90)
        Me.MainToolStrip.TabIndex = 3
        Me.MainToolStrip.Text = "ToolStrip1"
        '
        'btnBack
        '
        Me.btnBack.Enabled = False
        Me.btnBack.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBack.Image = CType(resources.GetObject("btnBack.Image"), System.Drawing.Image)
        Me.btnBack.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btnBack.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(114, 87)
        Me.btnBack.Text = "Back"
        Me.btnBack.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnBack.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'btnStop
        '
        Me.btnStop.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStop.Image = Global.WebbIE4.My.Resources.Resources.stop_image
        Me.btnStop.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btnStop.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(114, 87)
        Me.btnStop.Text = "Stop"
        Me.btnStop.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnStop.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'btnRefresh
        '
        Me.btnRefresh.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRefresh.Image = Global.WebbIE4.My.Resources.Resources.refresh
        Me.btnRefresh.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btnRefresh.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(167, 87)
        Me.btnRefresh.Text = "Refresh"
        Me.btnRefresh.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnRefresh.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'btnHome
        '
        Me.btnHome.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnHome.Image = Global.WebbIE4.My.Resources.Resources.home1
        Me.btnHome.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btnHome.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnHome.Name = "btnHome"
        Me.btnHome.Size = New System.Drawing.Size(142, 87)
        Me.btnHome.Text = "Home"
        Me.btnHome.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnHome.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'btnSearch
        '
        Me.btnSearch.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearch.Image = Global.WebbIE4.My.Resources.Resources.search
        Me.btnSearch.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btnSearch.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(152, 87)
        Me.btnSearch.Text = "Search"
        Me.btnSearch.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 90)
        '
        'btnCrop
        '
        Me.btnCrop.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCrop.Image = Global.WebbIE4.My.Resources.Resources.crop1
        Me.btnCrop.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btnCrop.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnCrop.Name = "btnCrop"
        Me.btnCrop.Size = New System.Drawing.Size(120, 87)
        Me.btnCrop.Text = "Crop"
        Me.btnCrop.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnCrop.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'btnHeading
        '
        Me.btnHeading.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnHeading.Image = Global.WebbIE4.My.Resources.Resources.heading
        Me.btnHeading.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btnHeading.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnHeading.Name = "btnHeading"
        Me.btnHeading.Size = New System.Drawing.Size(186, 87)
        Me.btnHeading.Text = "Heading"
        Me.btnHeading.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnHeading.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'picBusy
        '
        Me.picBusy.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.picBusy.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.picBusy.Image = Global.WebbIE4.My.Resources.Resources.timer_done_big
        Me.picBusy.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.picBusy.Name = "picBusy"
        Me.picBusy.Size = New System.Drawing.Size(81, 87)
        '
        'btnSkiplinks
        '
        Me.btnSkiplinks.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSkiplinks.Image = Global.WebbIE4.My.Resources.Resources.skiplinks
        Me.btnSkiplinks.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btnSkiplinks.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnSkiplinks.Name = "btnSkiplinks"
        Me.btnSkiplinks.Size = New System.Drawing.Size(212, 87)
        Me.btnSkiplinks.Text = "Skip Links"
        Me.btnSkiplinks.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnSkiplinks.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 90)
        '
        'btnRSS
        '
        Me.btnRSS.Enabled = False
        Me.btnRSS.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRSS.Image = Global.WebbIE4.My.Resources.Resources.rss
        Me.btnRSS.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.btnRSS.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnRSS.Name = "btnRSS"
        Me.btnRSS.Size = New System.Drawing.Size(98, 109)
        Me.btnRSS.Text = "RSS"
        Me.btnRSS.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnRSS.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        '
        'pnlMain
        '
        Me.pnlMain.Controls.Add(Me.txtText)
        Me.pnlMain.Controls.Add(Me.cboAddress)
        Me.pnlMain.Controls.Add(Me.lblAltDForAddress)
        Me.pnlMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlMain.Location = New System.Drawing.Point(0, 153)
        Me.pnlMain.Name = "pnlMain"
        Me.pnlMain.Size = New System.Drawing.Size(1362, 560)
        Me.pnlMain.TabIndex = 4
        '
        'txtText
        '
        Me.txtText.AccessibleName = "&Text"
        Me.txtText.DetectUrls = False
        Me.txtText.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtText.Location = New System.Drawing.Point(0, 65)
        Me.txtText.Name = "txtText"
        Me.txtText.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.txtText.Size = New System.Drawing.Size(1362, 495)
        Me.txtText.TabIndex = 1
        Me.txtText.Text = ""
        '
        'cboAddress
        '
        Me.cboAddress.AccessibleName = "A&ddress"
        Me.cboAddress.Dock = System.Windows.Forms.DockStyle.Top
        Me.cboAddress.FormattingEnabled = True
        Me.cboAddress.Location = New System.Drawing.Point(0, 0)
        Me.cboAddress.Name = "cboAddress"
        Me.cboAddress.Size = New System.Drawing.Size(1362, 65)
        Me.cboAddress.TabIndex = 1
        '
        'lblAltDForAddress
        '
        Me.lblAltDForAddress.AutoSize = True
        Me.lblAltDForAddress.Location = New System.Drawing.Point(76, 5)
        Me.lblAltDForAddress.Name = "lblAltDForAddress"
        Me.lblAltDForAddress.Size = New System.Drawing.Size(50, 57)
        Me.lblAltDForAddress.TabIndex = 0
        Me.lblAltDForAddress.Text = "&d"
        '
        'tmrDelayLoadBookmarks
        '
        Me.tmrDelayLoadBookmarks.Interval = 200
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(23.0!, 57.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1362, 741)
        Me.Controls.Add(Me.pnlMain)
        Me.Controls.Add(Me.MainToolStrip)
        Me.Controls.Add(Me.MenuStripMain)
        Me.Controls.Add(Me.staMain)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(15, 57)
        Me.Margin = New System.Windows.Forms.Padding(5, 6, 5, 6)
        Me.Name = "frmMain"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Tag = "frmMain"
        Me.Text = "WebbIE"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.staMain.ResumeLayout(False)
        Me.staMain.PerformLayout()
        Me.MenuStripMain.ResumeLayout(False)
        Me.MenuStripMain.PerformLayout()
        Me.MainToolStrip.ResumeLayout(False)
        Me.MainToolStrip.PerformLayout()
        Me.pnlMain.ResumeLayout(False)
        Me.pnlMain.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents mnuBar As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents MenuStripMain As System.Windows.Forms.MenuStrip
    Friend WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFileOpen As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFileSave As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFilePrint As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFileExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuEdit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuEditCopy As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuEditPaste As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuEditSelectall As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuEditFindtext As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuEditFindnext As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuEditWebsearch As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuView As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuViewCroppage As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuViewRSSNewsFeed As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuViewSource As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFavorites As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFavoritesAdd As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFavoritesGotofavorites As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFavoritesOrganise As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNavigate As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNavigateBack As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNavigateStop As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNavigateHome As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNavigateRefresh As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNavigateForward As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNavigateGotoheadline As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuNavigateGotoform As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuLinks As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuLinksNextlink As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuLinksPreviouslink As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuLinksSkipDown As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuLinksSkipup As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuLinksViewlinks As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuLinksDownloadlinktarget As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuLinksFollowlinkaddress As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuOptions As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuOptionsFont As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuOptionsOptions As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuOptionsSetHomepage As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuHelpManual As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuHelpWebbIEOrg As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuHelpAbout As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents fdMain As System.Windows.Forms.FontDialog
    Friend WithEvents MainToolStrip As System.Windows.Forms.ToolStrip
    Friend WithEvents btnBack As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnHeading As System.Windows.Forms.ToolStripButton
    Friend WithEvents pnlMain As System.Windows.Forms.Panel
    Friend WithEvents txtText As System.Windows.Forms.RichTextBox
    Friend WithEvents btnStop As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnRefresh As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnHome As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnRSS As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnSkiplinks As System.Windows.Forms.ToolStripButton
    Friend WithEvents picBusy As System.Windows.Forms.ToolStripLabel
    Friend WithEvents PrintDocument1 As System.Drawing.Printing.PrintDocument
    Friend WithEvents btnSearch As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnCrop As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Private WithEvents cboAddress As System.Windows.Forms.ComboBox
    Friend WithEvents lblAltDForAddress As System.Windows.Forms.Label
    Friend WithEvents mnuViewIE As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuFavoritesShowfavorites As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tmrDelayLoadBookmarks As System.Windows.Forms.Timer
    Friend WithEvents mnuOptionsColour As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuHelpTeamviewer As System.Windows.Forms.ToolStripMenuItem

#End Region
End Class