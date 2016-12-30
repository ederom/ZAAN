'ZAAN : Multidimensional Document Management software
'Copyright (C) 2015 Emmanuel DEROME
'
'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program. If not, see <http://www.gnu.org/licenses/>.

Public Class frmZaan
    Private mFoundTreeNodeKeys, mFoundListItemKeys As String
    Private mlvInOrder(10), mlvOutOrder(3), mlvTempOrder(3) As Integer
    Private mTreeClicked, mZaanAutoClose, mInputTreeClicked, mSplitter3MouseUp, mSplitter5MouseUp As Boolean
    Private mStartDragLvInItem As Boolean = False
    Private mSelectorMouseDown As Boolean = False
    Private mBookmarkListRightClick As Boolean = False
    Private mSelectorToUpdate As Boolean = False
    Private mSplitterBottomDistance As Integer
    Private mTreeViewWasClosed As Integer = False

    Private mMoveEffectData As Byte() = New Byte(3) {2, 0, 0, 0}     'used for file(s) moving in Cut/paste operations with Clipboard
    Private mStreamMove As Stream = New MemoryStream(4)              'will be set in InitFileMoveCopyStreams() with mMoveEffectData() value
    Private mCopyEffectData As Byte() = New Byte(3) {5, 0, 0, 0}     'used for file(s) copying in Copy/paste operations with Clipboard
    Private mStreamCopy As Stream = New MemoryStream(4)              'will be set in InitFileMoveCopyStreams() with mCopyEffectData() value

    Private DatabaseFontRegular As New System.Drawing.Font("Arial Rounded MT Bold", 14, Drawing.FontStyle.Regular)
    Private DatabaseFontUnderline As New System.Drawing.Font("Arial Rounded MT Bold", 14, Drawing.FontStyle.Underline)
    Private SelectorFontRegular As New System.Drawing.Font("Arial Rounded MT Bold", 12, Drawing.FontStyle.Regular)
    Private SelectorFontUnderline As New System.Drawing.Font("Arial Rounded MT Bold", 12, Drawing.FontStyle.Underline)

    Private SelectionFontDetail As New System.Drawing.Font("Microsoft Sans Serif", 10)
    Private SelectionFontImage As New System.Drawing.Font("Microsoft Sans Serif", 9)
    Private FileFontRegular As New System.Drawing.Font("Microsoft Sans Serif", 10, Drawing.FontStyle.Regular)
    Private FileFontUnderline As New System.Drawing.Font("Microsoft Sans Serif", 10, Drawing.FontStyle.Underline)
    Private DirFontRegular As New System.Drawing.Font("Microsoft Sans Serif", 9, Drawing.FontStyle.Regular)
    Private DirFontUnderline As New System.Drawing.Font("Microsoft Sans Serif", 9, Drawing.FontStyle.Underline)

    Private Sub SetCubeTubeButton()
        'Sets cube/tube button image depending on current lvIn display and ZAAN image style

        If mIsCubeView Then                                     '=> sets tube view mode
            If mImageStyle < 10 Then                             'case of a white background
                btnCubeTube.Image = My.Resources.Resources.view_tube_1
            Else                                                 'case of a dark blue background
                btnCubeTube.Image = My.Resources.Resources.view_tube_0
            End If
        Else                                                     '=> sets cube view mode
            If mImageStyle < 10 Then                             'case of a white background
                btnCubeTube.Image = My.Resources.Resources.view_cube_1
            Else                                                 'case of a dark blue background
                btnCubeTube.Image = My.Resources.Resources.view_cube_0
            End If
        End If
    End Sub

    Private Sub SetDetailsImagesButton()
        'Sets details/images button image depending on current lvIn display and ZAAN image style

        'If mIsImageView Then                                     '=> sets images view mode
        '  If mImageStyle < 10 Then                             'case of a white background
        '    btnDetailsImages.Image = My.Resources.Resources.view_details_1
        '  Else                                                 'case of a dark blue background
        '    btnDetailsImages.Image = My.Resources.Resources.view_details_0
        '  End If
        'Else                                                     '=> sets details view mode
        '  If mImageStyle < 10 Then                             'case of a white background
        '    btnDetailsImages.Image = My.Resources.Resources.view_icons_1
        '  Else                                                 'case of a dark blue background
        '    btnDetailsImages.Image = My.Resources.Resources.view_icons_0
        '  End If
        'End If
        tsmiLvInImageMode.Checked = mIsImageView
    End Sub

    Private Sub UpdatesListsHeadersVisibility()
        'Sets lvSelector, lvIn, lvOut and lvTemp headers visibility depending on mListsHeadersVisible state
        If mListsHeadersVisible Then                                 'makes lists headers visible
            lvIn.HeaderStyle = ColumnHeaderStyle.Clickable
            lvOut.HeaderStyle = ColumnHeaderStyle.Clickable
            lvTemp.HeaderStyle = ColumnHeaderStyle.Clickable
        Else                                                         'hides lists headers
            lvIn.HeaderStyle = ColumnHeaderStyle.None
            lvOut.HeaderStyle = ColumnHeaderStyle.None
            lvTemp.HeaderStyle = ColumnHeaderStyle.None
        End If
    End Sub

    Private Sub SetLeftPanelButton()
        'Sets left panel button display depending on splitter and background color
        tsmiSelectorTreeView.Checked = (Not SplitContainer2.Panel1Collapsed) And trvW.Visible
    End Sub

    Private Sub SetRightPanelButton()
        'Sets right panel button image depending on splitter and background color
        tsmiSelectorViewer.Checked = Not SplitContainer3.Panel2Collapsed
    End Sub

    Private Sub SetBottomPanelButton()
        'Sets bottom panel button display depending on splitter and background color
        tsmiSelectorImport.Checked = Not SplitContainer1.Panel2Collapsed
    End Sub

    Private Sub SetZaanControlsColors()
        'Sets background and fore color of all ZAAN controls to given mImageStyle and inializes file type images
        Dim fColorDB, fColorIn, fColorImport, fColorCopy As Color

        'Debug.Print("SetZaanControlsColors:  mImageStyle = " & mImageStyle)   'TEST/DEBUG

        If mImageStyle < 10 Then                                     'case of a white background
            mBackColorHeader = pctLvIn.BackColor                     'get standard control color by default
            mBackColorHeaderSel = Color.White
            mBackColorContent = Color.White
            mListsHeadersVisible = True
        Else                                                         'case of a dark blue background
            mBackColorHeader = Color.FromArgb(20, 20, 40)            'set light dark blue color
            mBackColorHeaderSel = Color.FromArgb(40, 40, 80)
            mBackColorContent = Color.FromArgb(0, 0, 20)             'set dark blue color
            mListsHeadersVisible = False
        End If
        Call UpdatesListsHeadersVisibility()                         'sets lvSelector, lvIn, lvOut and lvTemp headers visibility depending on mListsHeadersVisible state

        mBackColorPicture = Color.Black

        'Debug.Print("bg header brightness = " & mBackColorHeader.GetBrightness * 10)    'TEST/DEBUG
        If (mBackColorHeader.GetBrightness * 10 < 5) Then            'case of a dark header background
            fColorDB = Color.Silver
            mForeColorHeader = Color.SteelBlue
            mForeColorHeaderSel = Color.White
        Else                                                         'case of bright header background
            fColorDB = Color.DimGray
            mForeColorHeader = Color.MidnightBlue
            mForeColorHeaderSel = Color.MidnightBlue
        End If

        'Debug.Print("bg content brightness = " & mBackColorContent.GetBrightness * 10)    'TEST/DEBUG
        If (mBackColorContent.GetBrightness * 10 < 5) Then           'case of a dark content background
            fColorIn = Color.Silver
            fColorImport = Color.SteelBlue
            fColorCopy = Color.LimeGreen
            lvIn.GridLines = False
        Else                                                         'case of bright content background
            fColorIn = Color.Black
            fColorImport = Color.MidnightBlue
            fColorCopy = Color.ForestGreen
            lvIn.GridLines = True
        End If

        pnlListTop.BackColor = mBackColorHeader
        lvSelector.BackColor = mBackColorHeader
        lvSelector.ForeColor = fColorIn

        'pnlBookmarkLeft.BackColor = mBackColorHeader
        lvBookmark.BackColor = mBackColorHeader
        lvBookmark.ForeColor = fColorIn

        'tbSearch.BackColor = mBackColorContent
        'tbSearch.ForeColor = fColorDB

        'trvW.BackColor = mBackColorContent
        trvW.BackColor = mBackColorHeader
        trvW.ForeColor = fColorIn

        If lvIn.View = View.LargeIcon Then
            lvIn.BackColor = mBackColorPicture
            lvIn.ForeColor = Color.SteelBlue
        Else
            lvIn.BackColor = mBackColorContent
            lvIn.ForeColor = fColorIn
        End If

        lvOut.BackColor = mBackColorContent
        lvOut.ForeColor = fColorImport
        trvInput.BackColor = mBackColorContent
        trvInput.ForeColor = fColorImport

        lvTemp.BackColor = mBackColorContent
        lvTemp.ForeColor = fColorCopy

        'btnSearchFolder.ForeColor = fColorImport
        'btnSearchDoc.ForeColor = fColorImport

        'btnAboutZaan.ForeColor = fColorIn

        lblSelPagePrev.ForeColor = fColorIn
        lblSelPage.ForeColor = fColorIn
        lblSelPageNext.ForeColor = fColorIn

        btnDataAccess.ForeColor = fColorIn
        btnWhen.ForeColor = fColorIn
        btnWho.ForeColor = fColorIn
        btnWhat.ForeColor = fColorIn
        btnWhere.ForeColor = fColorIn
        btnWhat2.ForeColor = fColorIn
        btnWho2.ForeColor = fColorIn

        lblDataAccess.ForeColor = Color.White
        lblWhen.ForeColor = Color.White
        lblWho.ForeColor = Color.White
        lblWhat.ForeColor = Color.White
        lblWhere.ForeColor = Color.White
        lblWhat2.ForeColor = Color.White
        lblWho2.ForeColor = Color.White

        lblDataAccess.TextAlign = ContentAlignment.MiddleCenter
        lblWhen.TextAlign = ContentAlignment.MiddleCenter
        lblWho.TextAlign = ContentAlignment.MiddleCenter
        lblWhat.TextAlign = ContentAlignment.MiddleCenter
        lblWhere.TextAlign = ContentAlignment.MiddleCenter
        lblWhat2.TextAlign = ContentAlignment.MiddleCenter
        lblWho2.TextAlign = ContentAlignment.MiddleCenter

        SplitContainer2.BackColor = mBackColorHeader
        SplitContainer3.BackColor = mBackColorHeader

        If mImageStyle < 10 Then                                 'case of a white background
            tsmiSelectorNightMode.Checked = False
            'btnPrev.Image = My.Resources.Resources.prev_1
            'btnNext.Image = My.Resources.Resources.next_1
            lblBookmark.Image = My.Resources.Resources.cube_left_end_1b
            btnPanelTree.Image = My.Resources.Resources.pnl_tree_1
            btnPanelCube.Image = My.Resources.Resources.pnl_cube_1
            btnPanelView.Image = My.Resources.Resources.pnl_view_1
            btnPanelImport.Image = My.Resources.Resources.pnl_import_1
        Else                                                     'case of a dark blue background
            tsmiSelectorNightMode.Checked = True
            'btnPrev.Image = My.Resources.Resources.prev_0
            'btnNext.Image = My.Resources.Resources.next_0
            lblBookmark.Image = My.Resources.Resources.cube_left_end_0b
            btnPanelTree.Image = My.Resources.Resources.pnl_tree_0
            btnPanelCube.Image = My.Resources.Resources.pnl_cube_0
            btnPanelView.Image = My.Resources.Resources.pnl_view_0
            btnPanelImport.Image = My.Resources.Resources.pnl_import_0
        End If

        Call SetLeftPanelButton()                                    'sets left panel button image depending on splitter and background color
        Call SetRightPanelButton()                                   'sets right panel button image depending on splitter and background color
        Call SetBottomPanelButton()                                  'sets bottom panel button display depending on splitter and background color
        Call SetDetailsImagesButton()                                'sets details/images button image depending on current lvIn display and ZAAN image style
        'Call SetCubeTubeButton()                                     'sets cube/tube button image depending on current lvIn display and ZAAN image style
    End Sub

    Private Sub UpdateChildNodesImages(ByVal ParentNode As TreeNode)
        'Updates child nodes images of given parent node with current mImageStyle
        Dim NodeX As TreeNode

        For Each NodeX In ParentNode.Nodes                           'updates tree nodes images
            NodeX.ImageKey = Microsoft.VisualBasic.Left(NodeX.ImageKey, 3) & mImageStyle
            Call UpdateChildNodesImages(NodeX)                       'updates child nodes images with current mImageStyle
        Next
    End Sub

    Private Sub ChangeImageStyle(ByVal Style As Integer)
        'Sets mImageStyle to given style if not zero, updates related image background color and sets colors of all ZAAN controls
        Dim NodeX As TreeNode
        Dim itmX As ListViewItem

        'Debug.Print("ChangeImageStyle...")    'TEST/DEBUG
        If Style > 0 Then
            mImageStyle = Style                                      'stores selected style in mImageStyle

            btnDataAccessRoot.ImageKey = "_uh" & mImageStyle
            btnWhenRoot.ImageKey = "_th" & mImageStyle
            btnWhoRoot.ImageKey = "_oh" & mImageStyle
            btnWhatRoot.ImageKey = "_ah" & mImageStyle
            btnWhereRoot.ImageKey = "_eh" & mImageStyle
            btnWhat2Root.ImageKey = "_bh" & mImageStyle
            btnWho2Root.ImageKey = "_ch" & mImageStyle

            btnNext.ImageKey = "next_f" & mImageStyle
            btnPrev.ImageKey = "prev_f" & mImageStyle

            lvSelector.Columns(0).ImageKey = "_uh" & mImageStyle     'updates Access control column image
            lvSelector.Columns(1).ImageKey = "_th" & mImageStyle     'updates When column image
            lvSelector.Columns(2).ImageKey = "_oh" & mImageStyle     'updates Who column image
            lvSelector.Columns(3).ImageKey = "_ah" & mImageStyle     'updates What column image
            lvSelector.Columns(4).ImageKey = "_eh" & mImageStyle     'updates Where column image
            lvSelector.Columns(5).ImageKey = "_bh" & mImageStyle     'updates What else column image
            lvSelector.Columns(6).ImageKey = "_ch" & mImageStyle     'updates Other column image

            lvIn.Columns(1).ImageKey = "_u_" & mImageStyle      'updates Access control column image
            lvIn.Columns(2).ImageKey = "_t_" & mImageStyle      'updates When column image
            lvIn.Columns(3).ImageKey = "_o_" & mImageStyle      'updates Who column image
            lvIn.Columns(4).ImageKey = "_a_" & mImageStyle      'updates What column image
            lvIn.Columns(5).ImageKey = "_e_" & mImageStyle      'updates Where column image
            lvIn.Columns(6).ImageKey = "_b_" & mImageStyle      'updates What else column image
            lvIn.Columns(7).ImageKey = "_c_" & mImageStyle      'updates Other column image

            For Each NodeX In trvW.Nodes                        'updates tree nodes images
                NodeX.ImageKey = Microsoft.VisualBasic.Left(NodeX.ImageKey, 3) & mImageStyle
                Call UpdateChildNodesImages(NodeX)                'updates child nodes images with current mImageStyle
            Next
            If (Not trvW.SelectedImageKey = Nothing) And (Not trvW.SelectedNode Is Nothing) Then
                trvW.SelectedImageKey = Microsoft.VisualBasic.Left(trvW.SelectedNode.ImageKey, 3) & mImageStyle  'update current selection
            End If
            For Each itmX In lvBookmark.Items
                itmX.ImageKey = "_x_" & mImageStyle              'updates bookmarks images in bookmark list
            Next
        End If
        Call SetZaanControlsColors()                                 'sets colors of all ZAAN controls
    End Sub

    Private Sub EncryptFile(ByVal sInputFilename As String, ByVal sOutputFilename As String, ByVal sKey As String)
        'Encrypt an input file into an output file using a DES aglorithm with a given 64 bits secret key

        Dim fsInput As New FileStream(sInputFilename, FileMode.Open, FileAccess.Read)
        Dim fsEncrypted As New FileStream(sOutputFilename, FileMode.Create, FileAccess.Write)
        Dim DES As New DESCryptoServiceProvider()
        DES.Key = ASCIIEncoding.ASCII.GetBytes(sKey)                 'secret key of DES algorithm using a 64 bits (8 char.) secret key
        DES.IV = ASCIIEncoding.ASCII.GetBytes(sKey)                  'initiaization vector of DES algorithm

        Dim desencrypt As ICryptoTransform = DES.CreateEncryptor()
        Dim cryptostream As New CryptoStream(fsEncrypted, desencrypt, CryptoStreamMode.Write)

        Dim bytearrayinput(fsInput.Length - 1) As Byte               'reads file text into a byte array
        fsInput.Read(bytearrayinput, 0, bytearrayinput.Length)
        cryptostream.Write(bytearrayinput, 0, bytearrayinput.Length) 'writes byte array into encrypted file
        cryptostream.Close()
    End Sub

    Private Sub CreateFileFromZaanResourceString(ByVal FilePathName As String, ByVal FileContent As String)
        'Copy given Zaan resource/file content to given file name for creating it (ex: \help and \ZAAN-Demo1 directories initial loading)
        Try
            My.Computer.FileSystem.WriteAllText(FilePathName, FileContent, False)
        Catch ex As Exception
            Debug.Print(ex.Message)                                  'ZAAN\info directory not found !
        End Try
    End Sub

    Private Sub CreateFileFromZaanResourceByte(ByVal FilePathName As String, ByVal FileContent As Byte())
        'Copy given Zaan resource/file content to given file name for creating it (ex: \help and \ZAAN-Demo1 directories initial loading)
        Try
            My.Computer.FileSystem.WriteAllBytes(FilePathName, FileContent, False)
        Catch ex As Exception
            Debug.Print(ex.Message)                                  'ZAAN\info directory not found !
        End Try
    End Sub

    Private Sub CreateFileFromZaanResourceImage(ByVal FilePathName As String, ByVal ImageContent As Image)
        'Copy given Zaan resource/file content to given file name for creating it (ex: \help and \ZAAN-Demo1 directories initial loading)
        Try
            ImageContent.Save(FilePathName)
        Catch ex As Exception
            Debug.Print(ex.Message)                                  'ZAAN\info directory not found !
        End Try
    End Sub

    Private Sub CreateOnceZaanDirectories()
        'Creates once, if not exists, "ZAAN" and "ZAAN_demos" directories and related sub-directories "import", "copy" and "info"
        Dim DirPathName, DirPathNameZD, OldDirPathName, NewDirPathName As String

        DirPathName = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ZAAN"
        If CreateDirIfNotExistsOK(DirPathName) Then                       'tries to create a "ZAAN" directory if not exists in MyDocuments folder
            mZaanAppliPath = DirPathName & "\"                            'sets mZaanAppliPath
            mMyZaanImportPath = mZaanAppliPath & "import\"
            mZaanImportPath = mMyZaanImportPath
            mZaanCopyPath = mZaanAppliPath & "copy\"

            If CreateDirIfNotExistsOK(mMyZaanImportPath) Then             'tries to create a "ZAAN\import" directory if not exists in MyDocuments folder
            End If
            If CreateDirIfNotExistsOK(mZaanCopyPath) Then                 'tries to create a "ZAAN\copy" directory if not exists in MyDocuments folder
            End If
            If CreateDirIfNotExistsStatus(mZaanAppliPath & "info") = 1 Then       'case of "ZAAN\info" directory just created in MyDocuments folder
            End If
        Else
            MsgBox(mMessage(68) & vbCrLf & vbCrLf & mMessage(60), MsgBoxStyle.Critical)       'creation of ZAAN directory in (My)Documents failed ! ZAAN application will be closed !
            mZaanAutoClose = True                                         'enables ZAAN application closing without user confirmation
            Me.Close()                                                    'closes ZAAN application
        End If

        DirPathNameZD = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\ZAAN_demos"
        If CreateDirIfNotExistsOK(DirPathNameZD) Then
            mZaanDemoPath = DirPathNameZD & "\"                           'sets mZaanDemoPath

            OldDirPathName = mZaanAppliPath & "ZAAN-Demo1"                      'old demo path that may have been created by former ZAAN versions (2.3.166 and earlier)
            NewDirPathName = mZaanDemoPath & "ZAAN-Demo1"
            If My.Computer.FileSystem.DirectoryExists(OldDirPathName) Then     'searches for an eventual old ZAAN-Demo1 database
                Try
                    My.Computer.FileSystem.MoveDirectory(OldDirPathName, NewDirPathName)      'moves existing \ZAAN\ZAAN-Demo1 directory to \ZAAN_demos\
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation)           'cannot move existing \ZAAN\ZAAN-Demo1 directory !
                End Try
            End If
        Else
            MsgBox(mMessage(148) & vbCrLf & vbCrLf & mMessage(60), MsgBoxStyle.Critical)      'creation of ZAAN_demos directory in (My)Documents failed ! ZAAN application will be closed !
            mZaanAutoClose = True                                         'enables ZAAN application closing without user confirmation
            Me.Close()                                                    'closes ZAAN application
        End If
    End Sub

    Private Sub CreateDemoTreeFile(ByVal FileNameStart As String, ByVal TitleEN As String, Optional ByVal TitleFR As String = "")
        'Creates a Tree file with given file name in "ZAAN-Demo1" directory
        Dim FileName As String

        If TitleFR = "" Then TitleFR = TitleEN
        If mLanguage = "EN" Then
            FileName = FileNameStart & "__" & TitleEN
        Else
            FileName = FileNameStart & "__" & TitleFR
        End If
        Call CreateFileFromZaanResourceString(mZaanDemoPath & "ZAAN-Demo1\tree\" & FileName & ".txt", My.Resources._z000000000000__new)
    End Sub

    Private Sub CreateZaanPathTreeFile(ByVal FileName As String)
        'Creates a Tree file with given file name in current ZAAN database root directory set by mZaanDbPath
        Call CreateFileFromZaanResourceString(mZaanDbPath & "tree\" & FileName & ".txt", My.Resources._z000000000000__new)
    End Sub

    Private Sub CreateDemoDataFile(ByVal DestDataDir As String, ByVal DestFileName As String, ByVal StringContent As String, ByVal ImageContent As Image, ByVal ByteContent As Byte())
        'Creates a data file with given directory and file name in "ZAAN-Demo1" directory using given resource name
        Dim DataDirName, UCFileType As String

        DataDirName = mZaanDemoPath & "ZAAN-Demo1\data\" & DestDataDir
        If Not My.Computer.FileSystem.DirectoryExists(DataDirName) Then   'case of directory doesn't exist
            'Debug.Print("CreateDemoDataFile:  create directory : " & DestDataDir)       'TEST/DEBUG
            My.Computer.FileSystem.CreateDirectory(DataDirName)           'creates given directory
        End If
        UCFileType = UCase(Microsoft.VisualBasic.Right(DestFileName, 4))
        'Debug.Print("FileExt of : " & DestFileName & "  is : " & UCFileType)   'TEST/DEBUG
        Select Case UCFileType
            Case ".JPG", ".GIF", ".BMP"
                Call CreateFileFromZaanResourceImage(DataDirName & "\" & DestFileName, ImageContent)
            Case ".TXT"
                Call CreateFileFromZaanResourceString(DataDirName & "\" & DestFileName, StringContent)
            Case Else
                Call CreateFileFromZaanResourceByte(DataDirName & "\" & DestFileName, ByteContent)
        End Select
    End Sub

    Private Sub CreateZaanLogoIfNoTopLeftFile(ByVal DirPathName As String)
        Dim dirSel As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim TopLeftLogoExists As Boolean

        TopLeftLogoExists = False
        dirSel = My.Computer.FileSystem.GetDirectoryInfo(DirPathName)
        For Each fi In dirSel.GetFiles("top_left_logo.*")            'searches for existing "top_left_logo" file
            TopLeftLogoExists = True
            Exit For
        Next
        If Not TopLeftLogoExists Then
            Call CreateFileFromZaanResourceImage(DirPathName & "\top_left_logo.gif", My.Resources.zaan_cube_50i)
        End If
    End Sub

    Private Sub LoadResourceTreeZaanDemo(ByVal DemoNumber As Integer)
        'Loads/imports ZAAN tree files in current demo database from given resource demo tree file number
        Dim FileContent, FileLines(), LineCells(), NodeKey, ParentKey, NodeText, TreeFileName As String
        Dim i, j, LangIndex As Integer

        'Debug.Print("LoadResourceTreeZaanDemo :  mLanguage = " & mLanguage & "  DemoNumber = " & DemoNumber)     'TEST/DEBUG
        LangIndex = 2                                                     'sets English index (col. 2) by default
        Try
            Select Case DemoNumber
                Case 1
                    FileContent = My.Resources.demo_tree_1                'get demo_tree_1.txt file from ZAAN resources
                Case Else
                    FileContent = My.Resources.demo_tree_2                'get demo_tree_2.txt file from ZAAN resources
            End Select
            FileLines = Split(FileContent, vbCrLf)
            For i = 0 To FileLines.Length - 2
                If Mid(FileLines(i), 1, 1) <> "=" Then                    'line is not a comment line (starting with "=")
                    'Debug.Print("line " & i & " : " & FileLines(i))
                    LineCells = Split(FileLines(i), vbTab)
                    If i = 0 Then                                         'first line is like "NodeKey	ParentKey	EN	FR"
                        For j = 2 To LineCells.Length - 1
                            If LineCells(j) = mLanguage Then
                                LangIndex = j                             'mLanguage found : stores related column index
                                'Debug.Print("LangIndex = " & LangIndex)   'TEST/DEBUG
                                Exit For
                            End If
                        Next j
                    Else                                                  'reads message text of LangIndex column
                        If LineCells(0) = "" Then Exit For 'terminates at first empty line
                        NodeKey = LineCells(0)                            'get node key
                        ParentKey = LineCells(1)                          'get parent node key
                        If LangIndex > LineCells.Length - 1 Then
                            NodeText = LineCells(2)                       'get node text in English by default
                        Else
                            NodeText = LineCells(LangIndex)               'get node text in current language (mLanguage)
                        End If
                        TreeFileName = mZaanDemoPath & "ZAAN-Demo1\tree\_" & NodeKey & "_" & ParentKey & "__" & NodeText & ".txt"
                        Call CreateFileFromZaanResourceString(TreeFileName, My.Resources._z000000000000__new)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information)              'file not found !
            'Debug.Print(ex.Message)                                 'file not found !
        End Try
    End Sub

    Public Sub CreateTreeFilesRootsAndYear(ByVal CurDate As DateTime)
        'Creates in current ZAAN database (v3 format) 6 tree root files and 1 year/12 months at given year
        Call CreateZaanPathTreeFile("_o" & mTreeRootKey & "_o" & mTopRootKey & "__1 - Who")
        Call CreateZaanPathTreeFile("_a" & mTreeRootKey & "_a" & mTopRootKey & "__2 - What")
        Call CreateZaanPathTreeFile("_t" & mTreeRootKey & "_t" & mTopRootKey & "__3 - When")
        Call CreateZaanPathTreeFile("_e" & mTreeRootKey & "_e" & mTopRootKey & "__4 - Where")
        Call CreateZaanPathTreeFile("_b" & mTreeRootKey & "_b" & mTopRootKey & "__5 - Status")
        Call CreateZaanPathTreeFile("_c" & mTreeRootKey & "_c" & mTopRootKey & "__6 - Action for")

        Call CreateTreeFileYear(CurDate)                             'creates 1 year tree file including given date
    End Sub

    Private Sub LoadOnceZaanDemoDatabase()
        'Loads once, if not exists, "ZAAN_demos" directory and related "ZAAN-Demo1" sub-directory in v3+ database format (with 4Z When keys)
        Dim DirPathName, InfoDirPathName, TreeDirPathName, DataDirPathName, CurDataDirName As String
        Dim DemoNumber As Integer = 2                                     'initialized to business demo... with travel pictures !
        'Dim Response As MsgBoxResult

        Dim CurYear As String = Format(Now, "yyyy")
        Dim CurYearCode As String = GetWhenV3KeyFromDateText(CurYear)
        'Debug.Print("LoadOnceZaanDemoDatabase :  CurYearCode = " & CurYearCode)    'TEST/DEBUG

        'Debug.Print("LoadOnceZaanDemoDatabase...")       'TEST/DEBUG

        DirPathName = mZaanDemoPath & "ZAAN-Demo1"
        If CreateDirIfNotExistsStatus(DirPathName) = 1 Then               'case of "ZAAN-Demo1" directory just created
            'Debug.Print("LoadOnceZaanDemoDatabase :  ZAAN-Demo1 directory just created !")     'TEST/DEBUG
            Call UpdateZaanDbRootPathName(DirPathName & "\")              'updates current database selection with mZaanDbRoot, mZaanDbPath and mZaanDbName

            'Response = MsgBox(mMessage(141), MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Exclamation)   'Do you want to start with a business example (picture example if not) ?
            'If Response = MsgBoxResult.Yes Then                           'business demo
            '  DemoNumber = 2
            'Else                                                          'picture demo
            '  DemoNumber = 1
            'End If
            InfoDirPathName = DirPathName & "\info"
            TreeDirPathName = DirPathName & "\tree"
            DataDirPathName = DirPathName & "\data"

            If CreateDirIfNotExistsStatus(InfoDirPathName) = 1 Then       'case of "ZAAN-Demo1\info" directory just created
                Call CreateZaanLogoIfNoTopLeftFile(InfoDirPathName)
                mFileFilter = "*_t" & CurYearCode & "*_ozzzr*_azzzp*_ezzy6"         'initializes ZAAN-Demo1 FileFilter
                'If DemoNumber = 1 Then                                    'picture demo
                '    mFileFilter = "*_trgzz*_ozzzr*_azzzp*_ezzy6"          'TO BE UPDATED initializes ZAAN-Demo1 FileFilter
                'Else                                                      'business demo
                '    mFileFilter = "*_trgzz*_ozzzr*_azzzp*_ezzy6"          'initializes ZAAN-Demo1 FileFilter
                'End If
                Call SaveFileFilterIni()                                  'saves mFileFilter of ZAAN-Demo1 ZAAN database in ZAAN-Demo1\info\zaan.ini
            End If

            If CreateDirIfNotExistsStatus(TreeDirPathName) = 1 Then       'case of "ZAAN-Demo1\tree" directory just created
                Call CreateTreeFilesRootsAndYear(Now)                     'creates in current ZAAN database (v3 format) 6 tree root files and 1 year/12 months at given year
                Call LoadResourceTreeZaanDemo(DemoNumber)                 'loads ZAAN tree files in current demo database from given resource demo tree file number
            End If

            If CreateDirIfNotExistsStatus(DataDirPathName) = 1 Then       'case of "ZAAN-Demo1\data" directory just created
                If DemoNumber = 1 Then                                    'picture demo
                    CurDataDirName = "_t" & CurYearCode & "_ozzzszzzr_azzzszzzp_ezzymzzy6zzy5zzy4zzy3"
                    Call CreateDemoDataFile(CurDataDirName, "Port Navalo lighthouse top.jpg", "", My.Resources.EDE_0017, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "Fishing boats at Port Navalo.jpg", "", My.Resources.EDE_0047, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "Traditional home at Port Navalo.jpg", "", My.Resources.EDE_0048, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "New Jersey sailboat at Port Navalo.jpg", "", My.Resources.EDE_0082, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "Bay shuttle at Port Navalo.jpg", "", My.Resources.EDE_0085, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "Traditional sailboat at Port Navalo.jpg", "", My.Resources.EDE_0150, Nothing)

                    CurDataDirName = "_t" & CurYearCode & "_ozzzszzzr_azzzszzzp_ezzymzzy6zzy2zzy1zzy0"
                    Call CreateDemoDataFile(CurDataDirName, "Esplanade of La Defense.jpg", "", My.Resources.EDE_0530, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "SFR tower at La Defense.jpg", "", My.Resources.EDE_0531, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "Le Dome mall at La Defense.jpg", "", My.Resources.EDE_0541, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "EDF tower at La Defense.jpg", "", My.Resources.EDE_0545, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "La Grande Arche of La Defense.jpg", "", My.Resources.EDE_0546, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "The Red Spider of Alexander Calder.jpg", "", My.Resources.EDE_0547, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "EDF tower entry at La Defense.jpg", "", My.Resources.EDE_0578, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "Arc de Triomphe view from La Defense.jpg", "", My.Resources.EDE_0583, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "Cafe terrasse at La Defense.jpg", "", My.Resources.EDE_0587, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "Coeur de Defense building.jpg", "", My.Resources.EDE_0591, Nothing)
                Else                                                      'business demo
                    CurDataDirName = "_t" & CurYearCode & "_ozzzszzzr_azzzszzzp_ezzymzzy6zzy5zzy4zzy3"
                    Call CreateDemoDataFile(CurDataDirName, "Port Navalo lighthouse top.jpg", "", My.Resources.EDE_0017, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "Fishing boats at Port Navalo.jpg", "", My.Resources.EDE_0047, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "Traditional home at Port Navalo.jpg", "", My.Resources.EDE_0048, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "New Jersey sailboat at Port Navalo.jpg", "", My.Resources.EDE_0082, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "Bay shuttle at Port Navalo.jpg", "", My.Resources.EDE_0085, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "Traditional sailboat at Port Navalo.jpg", "", My.Resources.EDE_0150, Nothing)

                    CurDataDirName = "_t" & CurYearCode & "_ozzzszzzr_azzzszzzp_ezzymzzy6zzy2zzy1zzy0"
                    Call CreateDemoDataFile(CurDataDirName, "Esplanade of La Defense.jpg", "", My.Resources.EDE_0530, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "SFR tower at La Defense.jpg", "", My.Resources.EDE_0531, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "Le Dome mall at La Defense.jpg", "", My.Resources.EDE_0541, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "EDF tower at La Defense.jpg", "", My.Resources.EDE_0545, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "La Grande Arche of La Defense.jpg", "", My.Resources.EDE_0546, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "The Red Spider of Alexander Calder.jpg", "", My.Resources.EDE_0547, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "EDF tower entry at La Defense.jpg", "", My.Resources.EDE_0578, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "Arc de Triomphe view from La Defense.jpg", "", My.Resources.EDE_0583, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "Cafe terrasse at La Defense.jpg", "", My.Resources.EDE_0587, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "Coeur de Defense building.jpg", "", My.Resources.EDE_0591, Nothing)

                    CurDataDirName = "_t" & CurYearCode & "_ozzzszzzr_azzzlzzzkzzzizzzh_ezzymzzy6"
                    Call CreateDemoDataFile(CurDataDirName, "ZAAN license GNU-GPLv3.txt", My.Resources.gnu_gpl_v3, Nothing, Nothing)
                    Call CreateDemoDataFile(CurDataDirName, "ZAAN pitch FR 2015-02.pdf", "", Nothing, My.Resources.ZAAN_pitch_FR_2015_02)
                    Call CreateDemoDataFile(CurDataDirName, "ZAAN pitch EN 2015-02.pdf", "", Nothing, My.Resources.ZAAN_pitch_EN_2015_02)
                End If
                Call CreateFileFromZaanResourceString(mZaanAppliPath & "import\" & "zaan_test_1.txt", My.Resources.zaan_test_1)
                Call CreateFileFromZaanResourceString(mZaanAppliPath & "import\" & "zaan_test_2.txt", My.Resources.zaan_test_2)
            End If

            If CreateDirIfNotExistsOK(DirPathName & "\xin") Then          'tries to create a db "xin" directory if it doesn't exist
            End If
        End If
    End Sub

    Private Sub LoadSelectorIcons()
        'Loads in imgFileTypes all o/a/t/e/c/b/x icons to be displayed in Selector (and in Tree and in Lists), with 4 available styles
        '(WARNING : loading icons at design time causes colors variations !???)

        imgFileTypes.Images.Clear()
        imgFileTypes.Images.Add(".zaan", My.Resources.zaan)
        imgFileTypes.Images.Add(".zqm", My.Resources._zqm)

        imgFileTypes.Images.Add("_a_3", My.Resources._a_3)
        imgFileTypes.Images.Add("_a_13", My.Resources._a_13)
        imgFileTypes.Images.Add("_ah3", My.Resources._ah3)
        imgFileTypes.Images.Add("_ah13", My.Resources._ah13)

        imgFileTypes.Images.Add("_b_3", My.Resources._b_3)
        imgFileTypes.Images.Add("_b_13", My.Resources._b_13)
        imgFileTypes.Images.Add("_bh3", My.Resources._bh3)
        imgFileTypes.Images.Add("_bh13", My.Resources._bh13)

        imgFileTypes.Images.Add("_c_3", My.Resources._c_3)
        imgFileTypes.Images.Add("_c_13", My.Resources._c_13)
        imgFileTypes.Images.Add("_ch3", My.Resources._ch3)
        imgFileTypes.Images.Add("_ch13", My.Resources._ch13)

        imgFileTypes.Images.Add("_e_3", My.Resources._e_3)
        imgFileTypes.Images.Add("_e_13", My.Resources._e_13)
        imgFileTypes.Images.Add("_eh3", My.Resources._eh3)
        imgFileTypes.Images.Add("_eh13", My.Resources._eh13)

        imgFileTypes.Images.Add("_o_3", My.Resources._o_3)
        imgFileTypes.Images.Add("_o_13", My.Resources._o_13)
        imgFileTypes.Images.Add("_oh3", My.Resources._oh3)
        imgFileTypes.Images.Add("_oh13", My.Resources._oh13)

        imgFileTypes.Images.Add("_t_3", My.Resources._t_3)
        imgFileTypes.Images.Add("_t_13", My.Resources._t_13)
        imgFileTypes.Images.Add("_th3", My.Resources._th3)
        imgFileTypes.Images.Add("_th13", My.Resources._th13)

        imgFileTypes.Images.Add("_u_3", My.Resources._u_3)
        imgFileTypes.Images.Add("_u_13", My.Resources._u_13)
        imgFileTypes.Images.Add("_uh3", My.Resources._uh3)
        imgFileTypes.Images.Add("_uh13", My.Resources._uh13)

        imgFileTypes.Images.Add("_x_3", My.Resources._x_3)
        imgFileTypes.Images.Add("_x_13", My.Resources._x_13)
        imgFileTypes.Images.Add("_x_3d", My.Resources._x_3d)
        imgFileTypes.Images.Add("_x_13d", My.Resources._x_13d)

        imgFileTypes.Images.Add("next_f3", My.Resources.next_f3)
        imgFileTypes.Images.Add("next_f13", My.Resources.next_f13)
        imgFileTypes.Images.Add("prev_f3", My.Resources.prev_f3)
        imgFileTypes.Images.Add("prev_f13", My.Resources.prev_f13)

        imgFileTypes.Images.Add("add", My.Resources.add)
    End Sub

    Private Sub LoadTreeTenYears(ByVal ParentNode As TreeNode)
        'Loads in tree view 10 years and related months as children of given TenYearsCode (ex : "799" for "2000-2009") in v3 database format
        Dim TenYearsCode As String = Mid(ParentNode.Tag, 3, 3)
        Dim TenYears As String = GetDateTextFromWhenV3Key(TenYearsCode)
        Dim NodeY, NodeM As TreeNode
        Dim KeyY, KeyM As String
        Dim i, j, n, m As Integer

        'Debug.Print("LoadTreeTenYears : " & TenYearsCode & " => " & TenYears)     'TEST/DEBUG
        For i = 0 To 9
            n = 9 - i
            KeyY = "t" & TenYearsCode & n
            NodeY = ParentNode.Nodes.Add(KeyY, TenYears & i, mTreeCodeImageStyle)   'adds year node with related image
            NodeY.Tag = "_" & KeyY
            For j = 1 To 12
                m = 99 - j
                KeyM = KeyY & m
                NodeM = NodeY.Nodes.Add(KeyM, NodeY.Text & "-" & Format(j, "00"), mTreeCodeImageStyle)  'adds month node with related image
                NodeM.Tag = "_" & KeyM
            Next
        Next
    End Sub

    Private Sub LoadTreeYearMonths(ByVal ParentNode As TreeNode)
        'Loads in tree view 12 months of given year node in v3/v3+ database format
        Dim KeyY, KeyM, kM As String
        Dim m As Integer
        Dim NodeM As TreeNode

        'Debug.Print("LoadTreeYearMonths :  ParentNode = " & ParentNode.Text)    'TEST/DEBUG
        KeyY = Mid(ParentNode.Tag, 2, 3)
        For m = 1 To 12                                              'adds related 12 months
            kM = GetZ4KeyFromIndex(m, 1) & "z"                       'builds 1Z month code with appended "z" code of undefined day
            KeyM = KeyY & kM
            NodeM = ParentNode.Nodes.Add(KeyM, ParentNode.Text & "-" & Format(m, "00"), mTreeCodeImageStyle)  'adds month node with related image
            NodeM.Tag = "_" & KeyM
        Next
    End Sub

    Private Sub LoadTreeYearMonthsDays(ByVal ParentNode As TreeNode)
        'Loads in tree view 12 months and related days of given year node in v3/v3+ database format
        Dim KeyY, KeyM, kM, KeyD, kD As String
        Dim m, d, y, daysOfMonth As Integer
        Dim NodeM, NodeD As TreeNode
        Dim Date1, Date2 As Date

        'Debug.Print("LoadTreeYearMonths :  ParentNode = " & ParentNode.Text)    'TEST/DEBUG
        KeyY = Mid(ParentNode.Tag, 2, 3)
        y = ParentNode.Text
        For m = 1 To 12                                              'adds related 12 months
            kM = GetZ4KeyFromIndex(m, 1)                             'gets related 1Z month code... 
            KeyM = KeyY & kM & "z"                                   '... and appends it with "z" code of undefined day
            NodeM = ParentNode.Nodes.Add(KeyM, ParentNode.Text & "-" & Format(m, "00"), mTreeCodeImageStyle)  'adds month node with related image
            NodeM.Tag = "_" & KeyM

            Date1 = DateSerial(y, m, 1)
            Date2 = DateSerial(y, m + 1, 1)
            daysOfMonth = DateDiff(DateInterval.Day, Date1, Date2)   'get days count of current year/month
            For d = 1 To daysOfMonth
                kD = GetZ4KeyFromIndex(d, 1)                         'gets related 1Z day code... 
                KeyD = KeyY & kM & kD                                '... and appends it to get full key
                NodeD = NodeM.Nodes.Add(KeyD, ParentNode.Text & "-" & Format(m, "00") & "-" & Format(d, "00"), mTreeCodeImageStyle)  'adds day node with related image
                NodeD.Tag = "_" & KeyD
            Next
        Next
    End Sub

    Private Sub AddLineLogFileContent(ByVal LogLine As String)
        'Adds to mLogFileContent given line and initializes log content with current date/time

        If mLogFileContent = "" Then
            mLogFileContent = Format(Now, "yyyy-MM-dd HH:mm:ss") & " - " & mMessage(189) & vbCrLf   '<date/time> - ZAAN database check/correction report :
        End If
        mLogFileContent = mLogFileContent & LogLine & vbCrLf         'adds given log line
    End Sub

    Private Function GetDataDirectoryTitles(ByVal DataDirName As String) As String
        'Returns titles expression corresponding to given data directory name in current ZAAN database
        Dim Titles As String = ""
        Dim Keys(), Title As String
        Dim i As Integer

        If DataDirName <> "" Then
            Call ReplaceHierarchicalKeysBySimpleKeys(DataDirName)     'replaces in given data directory name hierarchical keys by related simple keys
            Keys = Split(DataDirName, "_")
            If Keys.Length > 0 Then
                For i = 1 To Keys.Length - 1                              'scans keys list to build related titles series (separated by " ")
                    Title = GetTreeNodeText(Keys(i))                      'get tree node title related to given key
                    If Title <> "" Then                                   'tree node title found => appends it to titles series
                        If Titles = "" Then
                            Titles = Title
                        Else
                            Titles = Titles & " " & Title
                        End If
                    End If
                Next
            End If
        End If
        GetDataDirectoryTitles = Titles
    End Function

    Private Sub RenameDataDirectoryOrMoveFiles(ByVal SceDirPathName As String, ByVal SceDirName As String, ByVal DestDirName As String, Optional ByVal ToLog As Boolean = False)
        'Renames given data directory of current database into given destination name if new, else move source files into destination and deletes source directory
        Dim dirInfo As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim DestFullDirName As String

        If DestDirName = "" Then
            DestDirName = "_t" & GetWhenV3KeyFromDateText(Format(Now, "yyyy"))                          'sets current year by default
        End If

        If SceDirName <> DestDirName Then
            If ToLog Then
                'AddLineLogFileContent("> " & mMessage(193) & SceDirName & " => " & DestDirName & vbCrLf & "(" & GetDataDirectoryTitles(DestDirName) & ")")  '> Move : ... => ... (titles)
                AddLineLogFileContent("> " & mMessage(193) & SceDirPathName & " => " & DestDirName & vbCrLf & "(" & GetDataDirectoryTitles(DestDirName) & ")")  '> Move : ... => ... (titles)
            End If
            DestFullDirName = mZaanDbPath & "data\" & DestDirName

            If My.Computer.FileSystem.DirectoryExists(DestFullDirName) Then                             'destination directory exists
                'dirInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "data\" & SceDirName)
                dirInfo = My.Computer.FileSystem.GetDirectoryInfo(SceDirPathName)
                For Each fi In dirInfo.GetFiles("*.*")                                                  'scans all source files to move then into destination directory
                    My.Computer.FileSystem.MoveFile(fi.FullName, DestFullDirName & "\" & fi.Name, FileIO.UIOption.AllDialogs, FileIO.UICancelOption.DoNothing)
                Next
            Else                                                                                        'destination directory is new
                'My.Computer.FileSystem.RenameDirectory(mZaanDbPath & "data\" & SceDirName, DestDirName)
                My.Computer.FileSystem.RenameDirectory(SceDirPathName, DestDirName)
            End If
            Call InitDisplaySelectedFiles()                          'initializes display of all selected files, starting at first page
        End If
    End Sub

    Private Sub CheckMissingTreeYearsDataPath(ByVal DataPath As String)
        'Checks missing years in given data path
        Dim dirInfo, dirSel As System.IO.DirectoryInfo
        Dim TreeCode, Keys(), KeyY, KeyY1, KeyY2, YearText, FileName, DestDirName As String
        Dim yMin, yMax, y, i As Integer

        'Debug.Print("CheckMissingTreeYearsDataPath :  DataPath = " & DataPath)   'TEST/DEBUG

        dirInfo = My.Computer.FileSystem.GetDirectoryInfo(DataPath)
        KeyY = ""
        KeyY1 = ""
        KeyY2 = ""
        yMin = Format(Now, "yyyy")                                   'initializes yMin and yMax to current year
        yMax = yMin
        For Each dirSel In dirInfo.GetDirectories("_*")        'scans all data directories to get used period
            Keys = Split(dirSel.Name, "_")
            If Keys.Length > 1 Then
                TreeCode = Mid(Keys(1), 1, 1)                        'in v3/v3+ dtabase format, When key, if any, is always at first position
                If TreeCode = "t" Then
                    KeyY = Mid(Keys(1), 2)                           'get When key code after "t"
                    If KeyY.Length < 4 Then                          'case of a too short When key
                        If KeyY.Length < 2 Then                      'invalid When key
                            KeyY = ""
                            DestDirName = ""
                        Else
                            KeyY = Mid(KeyY & "zz", 1, 4)            'appends undefined missing month/day "z" to uncomplete v3+ When key
                            DestDirName = "_t" & KeyY
                        End If
                        If Keys.Length > 2 Then
                            For i = 2 To Keys.Length - 1
                                If Keys(i) <> "" Then
                                    DestDirName = DestDirName & "_" & Keys(i)
                                End If
                            Next
                        End If
                        Call RenameDataDirectoryOrMoveFiles(dirSel.FullName, dirSel.Name, DestDirName)    'renames source directory into destination name if new, else move source files and deletes source dir.
                    End If
                    If (KeyY2 = "") And (KeyY <> "") Then
                        KeyY2 = KeyY                                           'stores end year code (related to most recent year)
                        y = Mid(GetDateTextFromWhenV3Key(KeyY2), 1, 4)         'sets yMax to most recent year
                        If y > yMax Then yMax = y
                    End If
                End If
            End If
        Next
        If KeyY <> "" Then
            If KeyY1 = "" Then
                KeyY1 = KeyY                                         'stores start year code (related to oldest year)
                y = Mid(GetDateTextFromWhenV3Key(KeyY1), 1, 4)       'sets yMin to oldest year
                If y < yMin Then yMin = y
            End If
            For y = yMin To yMax                                     'scans years in data period and checks eventually missing year in tree directory
                YearText = Format(y, "0000")
                FileName = "_t" & GetWhenV3KeyFromDateText(YearText) & "_t" & mTreeRootKey & "__" & YearText
                If Not My.Computer.FileSystem.FileExists(mZaanDbPath & "tree\" & FileName & ".txt") Then
                    'Debug.Print(" => CheckMissingTreeYears :  creates tree file : " & FileName)    'TEST/DEBUG
                    Call CreateZaanPathTreeFile(FileName)            'adds year tree file
                End If
            Next
        End If
    End Sub

    Private Sub CheckMissingTreeYears()
        'Checks from referenced years in data directory eventual missing years in tree directory and creates them (v3 database format)
        'and eventually appends missing undefined month/day "z" to v3+ When keys
        Dim NodeX As TreeNode

        Call CheckMissingTreeYearsDataPath(mZaanDbPath & "data")               'checks missing years in main data path
        If mLicTypeCode >= 30 Then
            For Each NodeX In trvW.Nodes(0).Nodes                              'scans all child nodes (groups) of "Access control" root node (dim 0)
                Call CheckMissingTreeYearsDataPath(Mid(NodeX.Tag, 3))          'checks missing years in Access control data path
            Next
        End If
    End Sub

    Public Function ParentNodeExistsCreatesTreeNodeFile(ByVal ParentKey As String, ByVal NodeKey As String, ByVal Title As String) As Boolean
        'Returns true if given parent node exists and creates in this case given tree node and related file
        Dim ParentExists As Boolean = False
        Dim NodeX(), NewNode As TreeNode
        Dim TreeCode, ImageKey, FileName As String

        NodeX = trvW.Nodes.Find(ParentKey, True)                          'searches for related tree node
        If NodeX.Length > 0 Then                                          'parent node found !
            ParentExists = True
            TreeCode = Mid(ParentKey, 1, 1)
            ImageKey = "_" & TreeCode & "_" & mImageStyle
            NewNode = NodeX(0).Nodes.Add(NodeKey, Title, ImageKey)        'adds restored node in tree view
            NewNode.SelectedImageKey = ImageKey
            FileName = "_" & NodeKey & "_" & ParentKey & "__" & Title
            NewNode.Tag = FileName & ".txt"
            Call CreateZaanPathTreeFile(FileName)                         'adds restored node file in tree directory
            AddLineLogFileContent("> " & mMessage(192) & Title)           '> Folder restored in tree view :
        End If
        ParentNodeExistsCreatesTreeNodeFile = ParentExists
    End Function

    Public Sub RestoreMissingTreeFile(ByRef BackupContent As String, ByVal ItemIndex As Integer)
        'Restores missing tree file from given backup file content at given item/node index position, including missing parents
        Dim BackupLine, LineCells(), TreeCode, ParentKey As String
        Dim p As Integer

        BackupLine = ""
        p = InStr(ItemIndex + 2, BackupContent, vbCrLf)
        If p > 0 Then
            BackupLine = Mid(BackupContent, ItemIndex + 2, p - ItemIndex - 2)
        Else
            BackupLine = Mid(BackupContent, ItemIndex + 2)
        End If

        If BackupLine <> "" Then
            LineCells = Split(BackupLine, vbTab)
            If LineCells.Length > 2 Then
                ParentKey = LineCells(1)
                TreeCode = Mid(ParentKey, 1, 1)
                If Not ParentNodeExistsCreatesTreeNodeFile(ParentKey, LineCells(0), LineCells(2)) Then
                    p = InStr(BackupContent, vbCrLf & ParentKey)               'searches for lost parent node in backup file
                    If p > 0 Then                                              'lost parent node has been found in backup
                        Call RestoreMissingTreeFile(BackupContent, p)          'restores related missing tree file from given backup file content
                    Else                                                       'lost tree key does not exist in backup !
                        ParentKey = TreeCode & mTreeRootKey                    'lost parent : replaces it by related tree root node
                    End If
                    If ParentNodeExistsCreatesTreeNodeFile(ParentKey, LineCells(0), LineCells(2)) Then
                    End If
                End If
            End If
        End If
    End Sub

    Private Function CheckRestoreNodeHierarchicalKey(ByVal HierarchicalKey As String, ByRef BackupContent As String) As String
        'Checks if related simple key (v3+ database format) exists in tree node : restores tree related file(s) if exists only in backup file else returns parent hierarchical key
        Dim TreeCode, SimpleKey As String
        Dim NodeX() As TreeNode
        Dim ReturnedKey, HierarchicalKey4, NewKey As String
        Dim p As Integer

        ReturnedKey = HierarchicalKey
        TreeCode = Mid(HierarchicalKey, 1, 1)                        'is o/a/e/b/c key code
        If HierarchicalKey.Length > 5 Then
            SimpleKey = TreeCode & Microsoft.VisualBasic.Right(HierarchicalKey, 4)     'get simple key (5 char.)
        Else
            SimpleKey = HierarchicalKey
        End If
        HierarchicalKey4 = Mid(SimpleKey, 2, 4)                      'initializes key(4) to simple key

        NodeX = trvW.Nodes.Find(SimpleKey, True)                     'searches for related tree node
        If NodeX.Length > 0 Then                                     'given key exists in tree view => rebuilds related hierarchical key
            Call ReplaceSimpleKey4ByHierarchicalKey4(TreeCode, HierarchicalKey4)   'replaces given simple key4 by related hierarchical key4 (including parent key4s)
            NewKey = TreeCode & HierarchicalKey4                     'rebuilds hierarchical key
            If NewKey <> HierarchicalKey Then                        'hierarchical key needs to be updated
                AddLineLogFileContent(vbCrLf & "! " & mMessage(194) & HierarchicalKey)   '! Defective data index : ...
            End If
            ReturnedKey = NewKey
        Else                                                         'given key is not defined in tree view => searches in backup file
            AddLineLogFileContent(vbCrLf & "! " & mMessage(190) & SimpleKey)   '! Undefined folder index : ...
            p = InStr(BackupContent, vbCrLf & SimpleKey)             'searches for lost tree key in backup file
            If p > 0 Then                                            'lost tree key has been found in backup
                Call RestoreMissingTreeFile(BackupContent, p)        'restores missing tree file from given backup file content at given item/node index position, including missing parents
                Call ReplaceSimpleKey4ByHierarchicalKey4(TreeCode, HierarchicalKey4)   'replaces given simple key4 by related hierarchical key4 (including parent key4s)
                ReturnedKey = TreeCode & HierarchicalKey4            'rebuilds hierarchical key
            Else                                                     'lost tree key does not exist in backup !
                AddLineLogFileContent("! " & mMessage(191))          '! Index not found in tree view backup file
                If HierarchicalKey.Length > 8 Then
                    HierarchicalKey = Mid(HierarchicalKey, 1, Len(HierarchicalKey) - 4)      'truncates last 4 char. of hierarchical key
                    ReturnedKey = CheckRestoreNodeHierarchicalKey(HierarchicalKey, BackupContent)       'recursive call using to check if new key is valid...
                Else
                    ReturnedKey = ""                                 'no valid key !
                End If
            End If
        End If
        'Debug.Print(" => CheckRestoreNodeHierarchicalKey : returned key = " & ReturnedKey)   'TEST/DEBUG
        CheckRestoreNodeHierarchicalKey = ReturnedKey
    End Function

    Private Sub CheckDataDirIndexes(ByVal DataPath As String, ByRef BackupContent As String)
        'Checks given data directory having broken or undefined o/a/e/b/c keys for trying to restore them, using eventually tree backup file
        Dim dirInfo, dirSel As System.IO.DirectoryInfo
        Dim TreeCode, Keys(), DestDirName As String
        Dim i, FilNb, k As Integer

        'Debug.Print("CheckDataDirIndexes :  DataPath = " & DataPath)    'TEST/DEBUG

        dirInfo = My.Computer.FileSystem.GetDirectoryInfo(DataPath)
        FilNb = dirInfo.GetDirectories.Length
        k = 0
        For Each dirSel In dirInfo.GetDirectories("_*", SearchOption.TopDirectoryOnly)   'search for ZAAN cube directories
            k = k + 1
            pgbZaan.Value = 100 * k / FilNb
            Keys = Split(dirSel.Name, "_")
            DestDirName = ""
            For i = 1 To Keys.Length - 1                                         'scans all keys of current data directory
                TreeCode = Mid(Keys(i), 1, 1)
                If TreeCode <> "t" Then                                          'case of o/a/e/b/c key
                    Keys(i) = CheckRestoreNodeHierarchicalKey(Keys(i), BackupContent)     'checks if related simple key exists in tree node : restores tree related file(s) if exists only in backup file else returns parent hierarchical key
                End If
                If Keys(i) <> "" Then
                    DestDirName = DestDirName & "_" & Keys(i)
                End If
            Next
            Call RenameDataDirectoryOrMoveFiles(dirSel.FullName, dirSel.Name, DestDirName, True)  'renames source directory into destination name if new, else move source files and deletes source dir.
        Next
    End Sub

    Private Sub CheckDatabaseIndexes()
        'Checks data directories having broken or undefined o/a/e/b/c keys for trying to restore them, using eventually tree backup file (v3+ format only !)
        Dim Response As MsgBoxResult
        Dim BackupPathFileName, BackupContent, LogPathFileName As String
        Dim NodeX As TreeNode

        Response = MsgBox(mMessage(187), MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation)    'Do you confirm current database check (database administration operation) ?
        If Response = MsgBoxResult.Yes Then
            BackupPathFileName = mZaanDbPath & "info\" & mBackupTreeFileName             'get tree view backup file
            BackupContent = ""
            If My.Computer.FileSystem.FileExists(BackupPathFileName) Then                'tree backup file exists
                Try
                    BackupContent = My.Computer.FileSystem.ReadAllText(BackupPathFileName)
                Catch ex As Exception
                    Debug.Print(ex.Message)                                              'backup file not found !
                End Try
            End If

            BackupContent = vbCrLf & BackupContent                                       'makes sure that even first key will start after a CR/LF char. series

            If BackupContent <> "" Then
                mLogFileContent = ""                                                     'clears new log file content
                pgbZaan.Value = 0
                pgbZaan.Visible = True                                                   'shows progress bar

                Call CheckDataDirIndexes(mZaanDbPath & "data", BackupContent)            'checks main data directory
                If mLicTypeCode >= 30 Then
                    For Each NodeX In trvW.Nodes(0).Nodes                                'scans all child nodes (groups) of "Access control" root node (dim 0)
                        Call CheckDataDirIndexes(Mid(NodeX.Tag, 3), BackupContent)       'checks given Access control data directory
                    Next
                End If

                pgbZaan.Value = pgbZaan.Maximum                                          'makes sure that progress bar is set to maximum
                If mLogFileContent = "" Then                                             'no issue/correction logged
                    MsgBox(mMessage(188), MsgBoxStyle.Information)                       'ZAAN database check is terminated (no issue detected) !
                Else                                                                     'some issue(s)/correction(s) logged
                    MsgBox(mLogFileContent, MsgBoxStyle.Exclamation)                     '<date/time> - ZAAN database check report : ...
                    Try
                        LogPathFileName = mZaanDbPath & "info\" & mAdminLogFileName      'sets admin log file name
                        My.Computer.FileSystem.WriteAllText(LogPathFileName, mLogFileContent & vbCrLf, True)   'appends new log file content to existing file content
                    Catch ex As System.IO.DirectoryNotFoundException
                        Debug.Print(ex.Message)                                  'TEST/DEBUG
                    End Try
                    mLogFileContent = ""                                                 'clears admin log file content
                End If
                pgbZaan.Visible = False                                                  'hides progress bar
            End If
        End If
    End Sub

    Private Sub LoadChildNodes(ByVal ParentNode As TreeNode)
        'Loads child nodes (if any) of given parent node
        Dim dirInfo As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim ParentKey, FileFilter, Key, Title, TreeCode As String
        Dim NodeX As TreeNode

        If mTreeKeyLength = 0 Then                                   'no node keys in v4 database format (use directly Windows path names)
            dirInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "tree\" & ParentNode.FullPath)
            For Each dirSel In dirInfo.GetDirectories()
                'Debug.Print("=> LoadChildNodes : " & dirSel.Name)     'TEST/DEBUG
                Key = ParentNode.Tag & "\" & dirSel.Name
                NodeX = ParentNode.Nodes.Add(Key, dirSel.Name, mTreeCodeImageStyle)   'adds the tree root node with related image
                'NodeX = ParentNode.Nodes.Add(dirSel.Name)            'adds the tree root node with related image
                'NodeX.ImageKey = mTreeCodeImageStyle
                mTreeNodesCount = mTreeNodesCount + 1
                NodeX.Tag = Key
                LoadChildNodes(NodeX)                                'loads child nodes (children) if any
            Next
        Else
            dirInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "tree\")
            ParentKey = Mid(ParentNode.Tag, 2, mTreeKeyLength)
            FileFilter = "_*_" & ParentKey & "*.txt"
            For Each fi In dirInfo.GetFiles(FileFilter)              'scans all tree files pointing to selected parent key
                Key = Mid(fi.Name, 2, mTreeKeyLength)
                TreeCode = Microsoft.VisualBasic.Left(Key, 1)
                Title = Mid(fi.Name, 5 + 2 * mTreeKeyLength)
                If Microsoft.VisualBasic.Right(Title, 4) = ".txt" Then
                    Title = Microsoft.VisualBasic.Left(Title, Len(Title) - 4)
                End If
                If TreeCode = "t" Then                               'in v3 database format, dates are implicitly defined by related key code
                    If (Mid(Key, 5, 1) = "x") And (Len(Title) = 9) Then        'is a 10 years tree file (ex: "2000-2009" not more used since v3.0.235 release)
                        Call DeleteTreeFile(fi.Name)                 'deletes 10 years tree file
                    Else                                             'in v3 database format (since v3.0.235 release), years are directly listed as When root children
                        NodeX = ParentNode.Nodes.Add(Key, Title, "_" & TreeCode & "_" & mImageStyle)    'adds a child node
                        NodeX.Tag = "_" & Key
                        mTreeNodesCount = mTreeNodesCount + 1
                        Call LoadTreeYearMonthsDays(NodeX)           'loads in tree view year (related to given node) and related months/days
                    End If
                Else
                    NodeX = ParentNode.Nodes.Add(Key, Title, "_" & TreeCode & "_" & mImageStyle)    'adds a child node
                    NodeX.Tag = fi.Name
                    mTreeNodesCount = mTreeNodesCount + 1
                    LoadChildNodes(NodeX)                            'loads child nodes (children) if any
                End If
            Next
        End If
    End Sub

    Private Sub LoadTreesFromWin()
        'Loads Who/What/When/Where/What2/Who2 trees in stored order of related Windows tree directories in (v4 database format)
        Dim TreeCodes() As String = Split(mTreeCodeSeries, " ")
        Dim dirSel, dirInfo As System.IO.DirectoryInfo
        Dim SelPattern As String = "? - *"
        Dim TreePos, ColIndex As Integer
        Dim TreeCode, LocalTitle, Key As String
        Dim NodeX As TreeNode

        'Debug.Print("LoadTreesFromWin...")    'TEST/DEBUG

        dirInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "tree")
        For Each dirSel In dirInfo.GetDirectories(SelPattern)        'searches for root folders (like "1 - What" for ex.)
            LocalTitle = Mid(dirSel.Name, 5)
            TreePos = Mid(dirSel.Name, 1, 1)
            TreeCode = ""
            If (TreePos > 0) And (TreePos < 7) And (TreePos <= TreeCodes.Length) Then
                TreeCode = TreeCodes(TreePos - 1)
                If ((TreeCode = "b") Or (TreeCode = "c")) And (mLicTypeCode < 30) Then TreeCode = "" 'cancels "b" and "c" tree code if not ZAAN-Pro license
            End If
            If TreeCode <> "" Then
                mTreeCodeImageStyle = "_" & TreeCode & "_" & mImageStyle     'set image key relative to current tree code
                Key = "z" & TreePos
                NodeX = trvW.Nodes.Add(Key, dirSel.Name, mTreeCodeImageStyle)   'adds the tree root node with related image
                'NodeX = trvW.Nodes.Add(dirSel.Name)                  'adds the tree root node with related image
                'NodeX.ImageKey = mTreeCodeImageStyle
                lvSelector.Columns(TreePos).Text = LocalTitle        'updates related lvSelector column header
                lvSelector.Columns(TreePos).ImageKey = mTreeCodeImageStyle
                ColIndex = TreePos + 1
                lvIn.Columns(ColIndex).Text = LocalTitle             'updates related lvIn column header
                lvIn.Columns(ColIndex).ImageKey = mTreeCodeImageStyle

                mTreeNodesCount = mTreeNodesCount + 1
                NodeX.Tag = Key
                LoadChildNodes(NodeX)                                'loads child nodes (children) if any
            End If
        Next
    End Sub

    Private Sub LoadTreeW(ByVal TreeCode As String, ByVal TreePos As Integer)
        'Loads tree root node corresponding to given tree code ("o", "a", "t", "e", "b" and "c") and all related child nodes
        Dim PathFile, PathFileName, RootFileName, Key, LocalTitle As String
        Dim NodeX As TreeNode

        mTreeCodeImageStyle = "_" & TreeCode & "_" & mImageStyle     'set image key relative to current tree code

        LocalTitle = TreePos & " - " & mMessage(TreePos)             'get local text of tree root node
        'Debug.Print("LoadTreeW :  TreeCode = " & TreeCode & "  LocalTitle = " & LocalTitle)     'TEST/DEBUG

        RootFileName = "_" & TreeCode & mTreeRootKey & "_" & TreeCode & mTopRootKey & "__" & LocalTitle & ".txt"
        Key = TreeCode & mTreeRootKey
        PathFile = mZaanDbPath & "tree\_" & TreeCode & mTreeRootKey & "_" & TreeCode & mTopRootKey & "__*.*"

        PathFileName = Dir(PathFile)                                 'searches for tree root file
        If PathFileName <> "" Then                                   'such a tree root file exists (maybe with a different title)
            'If PathFileName <> RootFileName Then                     'title is not a local title
            '  Try
            '    My.Computer.FileSystem.RenameFile(mZaanDbPath & "tree\" & PathFileName, RootFileName)      'renames tree root file with local title
            '  Catch ex As Exception
            '    Debug.Print("** LoadTreeW: " & ex.Message)       'file renaming error
            '  End Try
            'End If
        Else                                                         'such a tree root file does not exist (may be a new dimension)
            'Debug.Print("=> LoadTreeW :  create file : " & RootFileName)     'TEST/DEBUG
            Call CreateFileFromZaanResourceString(mZaanDbPath & "tree\" & RootFileName, My.Resources._z000000000000__new)   'creates tree root file
        End If

        NodeX = trvW.Nodes.Add(Key, LocalTitle, "_" & TreeCode & "_" & mImageStyle)   'adds the tree root node with related image
        mTreeNodesCount = mTreeNodesCount + 1
        NodeX.Tag = RootFileName
        LoadChildNodes(NodeX)                                        'loads child nodes (children) if any
    End Sub

    Private Sub LoadTreeDataAccessFolders(ByRef Node0 As TreeNode, ByVal DirPath As String, ByVal DirFilter As String, ByVal StartPos As Integer, Optional ByVal AddWatcher As Boolean = False)
        'Loads tree data access (internal/external) folders at given path/filter
        Dim dirInfo, dirSel, dirTest As System.IO.DirectoryInfo
        Dim Key, GroupName, testDirPathName As String
        Dim NodeX As TreeNode

        'Debug.Print("LoadTreeDataAccessFolders :  DirPath = " & DirPath & "  DirFilter = " & DirFilter)   'TEST/DEBUG

        dirInfo = My.Computer.FileSystem.GetDirectoryInfo(DirPath)
        For Each dirSel In dirInfo.GetDirectories(DirFilter)
            Try
                testDirPathName = ""
                For Each dirTest In dirSel.GetDirectories()          'dirSel object can be accessed only if current Windows user has related access rights
                    testDirPathName = dirTest.Name
                    Exit For
                Next
                GroupName = Mid(dirSel.Name, StartPos)
                'Key = "u" & GroupName
                Key = "u" & dirSel.FullName                          'SINCE 2012-03-30 : uses full path of related data directory

                NodeX = Node0.Nodes.Add(Key, GroupName, "_u_" & mImageStyle)   'adds child node with related image
                mTreeNodesCount = mTreeNodesCount + 1
                NodeX.Tag = "_" & Key                                'saves full path of related data directory
                NodeX.ToolTipText = dirSel.FullName
                'Debug.Print("LoadTreeDataAccessFolders :  NodeX.Tag = " & NodeX.Tag)   'TEST/DEBUG

                If AddWatcher Then
                    'Debug.Print("LoadTreeDataAccessFolders :  add watcher at path = " & dirSel.FullName)   'TEST/DEBUG
                    'TO BE DONE !!!

                End If
            Catch ex As Exception
                Debug.Print("LoadTreeDataAccessFolders : " & ex.Message)       'user access to dirSel directory is not allowed !
            End Try
        Next
    End Sub

    Private Sub LoadTreeDataAccess()
        'Loads any existing "data\data_*" directories matching with user access right ("UnauthorizedAccessException" is catched if user access is denied by Windows !)
        Dim dirInfo, dirSel, dirTest As System.IO.DirectoryInfo
        Dim LocalTitle, Key, GroupName, testDirPathName As String
        Dim Node0, NodeX As TreeNode

        LocalTitle = "0 - " & mMessage(7)                                      'get local text of "Access control" tree root node

        If mTreeKeyLength = 0 Then                                             'no node keys in v4 (use directly Windows path names)
            Key = "z0"
            Node0 = trvW.Nodes.Add(Key, LocalTitle, "_u_" & mImageStyle)       'adds the tree root node with related image
            mTreeNodesCount = mTreeNodesCount + 1
            Node0.Tag = Key
            dirInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath)
            For Each dirSel In dirInfo.GetDirectories("data_*")
                Try
                    testDirPathName = ""
                    For Each dirTest In dirSel.GetDirectories()                'dirSel object can be accessed only if current Windows user has related access rights
                        testDirPathName = dirTest.Name
                        Exit For
                    Next
                    GroupName = Mid(dirSel.Name, 6)
                    Key = "z0\" & GroupName
                    NodeX = Node0.Nodes.Add(Key, GroupName, "_u_" & mImageStyle)   'adds child node with related image
                    mTreeNodesCount = mTreeNodesCount + 1
                    NodeX.Tag = Key
                Catch ex As Exception
                    Debug.Print(ex.Message)                                    'user access to dirSel directory is not allowed !
                End Try
            Next
        Else
            Key = "u" & mTreeRootKey
            Node0 = trvW.Nodes.Add(Key, LocalTitle, "_u_" & mImageStyle)       'adds the tree root node with related image
            mTreeNodesCount = mTreeNodesCount + 1
            Node0.Tag = "_" & Key                                              'no "tree\*.txt" file associated to this (virtual) Windows directory
            Node0.ToolTipText = mZaanDbPath & "data"

            Call LoadTreeDataAccessFolders(Node0, mZaanDbPath & "data", "data_*", 6)                               'loads tree data access (internal) folders at given path/filter
            If mIsLocalDatabase Then
                'Call LoadTreeDataAccessFolders(Node0, mZaanDbRoot, mZaanDbName & "_*", Len(mZaanDbName) + 2, True)          'loads tree data access (external) folders at given path/filter
                Call LoadTreeDataAccessFolders(Node0, mZaanDbRoot, mZaanDbName & "_data_*", Len(mZaanDbName) + 7, True)     'loads tree data access (external) folders at given path/filter
            End If
        End If
    End Sub

    Private Sub LoadTrees()
        'Loads all trees (code) : Who (o), What (a), When (t), Where (e) and What else (b) and Other (c) in ZAAN-Pro mode

        'Debug.Print("=> LoadTrees...")
        Me.Cursor = Cursors.WaitCursor                     'sets wait cursor
        fswTree.EnableRaisingEvents = False                'locks fswTree related events

        trvW.Nodes.Clear()
        mTreeNodesCount = 0                                'resets tree nodes counter

        If mLicTypeCode >= 30 Then                         'case of "Pro" license (access control enabled)
            LoadTreeDataAccess()                           'loads any existing "data\data_*" directories matching with user access right
        End If

        If mTreeKeyLength = 0 Then                         'no node keys in v4 (use directly Windows path names)
            Call LoadTreesFromWin()                        'loads 4 Who/What/When/Where trees in stored order of related Windows tree directories
        Else
            LoadTreeW("t", 1)                              'loads When tree
            LoadTreeW("o", 2)                              'loads Who tree
            LoadTreeW("a", 3)                              'loads What tree
            LoadTreeW("e", 4)                              'loads Where tree
            If mLicTypeCode >= 30 Then                     'case of "Pro" license (6 dimensions) :
                LoadTreeW("b", 5)                          'loads What2/Status tree
                LoadTreeW("c", 6)                          'loads Who2/Action for tree
            End If
        End If

        trvW.Sort()
        fswTree.EnableRaisingEvents = True                 'unlocks fswTree related events
        Me.Cursor = Cursors.Default                        'resets default cursor
    End Sub

    Private Sub LoadExportedChildNodes(ByVal TreeCode As String, ByVal dirInfo As System.IO.DirectoryInfo, ByVal ParentNode As TreeNode)
        'Loads child nodes (if any) of given parent node of exported tree
        Dim infos As FileSystemInfo() = dirInfo.GetFileSystemInfos
        Dim info As FileSystemInfo
        Dim NodeX As TreeNode

        For Each info In infos
            If info.Attributes = FileAttributes.Directory Then       'is a directory
                Try
                    'Debug.Print("dir found : " & info.Name & "  " & info.FullName)     'TEST/DEBUG
                    NodeX = ParentNode.Nodes.Add(info.FullName, info.Name, "_" & TreeCode & "_" & mImageStyle)    'adds found directory as a child node
                    NodeX.Tag = info.FullName
                    mTreeNodesCount = mTreeNodesCount + 1
                    LoadExportedChildNodes(TreeCode, info, NodeX)    'loads child nodes if any
                Catch ex As Exception
                    Debug.Print(ex.Message)                          'error message !
                End Try
            End If
        Next
    End Sub

    Private Sub LoadExportedTreeW(ByVal TreeCode As String, ByVal TreePos As Integer)
        'Loads exported tree root node corresponding to given tree code ("o", "a", "t", "e", "b" and "c") and all related child nodes
        Dim dirInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath)
        Dim dirSel As System.IO.DirectoryInfo
        Dim DirName, LocalName, FullLocalName As String
        Dim NodeX As TreeNode

        LocalName = TreePos & " - " & mMessage(TreePos)              'get local text of tree root node
        Debug.Print("LoadExportedTreeW :  TreeCode = " & TreeCode & "  LocalName = " & LocalName)     'TEST/DEBUG

        FullLocalName = mZaanDbPath & LocalName
        DirName = ""
        For Each dirSel In dirInfo.GetDirectories(TreePos & "*")
            DirName = dirSel.Name
            If DirName <> LocalName Then
                Try
                    My.Computer.FileSystem.RenameDirectory(mZaanDbPath & DirName, LocalName)    'renames tree root file with local name
                    'Debug.Print("=> LoadTreeW :  directory renamed : " & mZaanDbPath & "docs\" & DirName & " in " & LocalName)    'TEST/DEBUG
                Catch ex As Exception
                    Debug.Print("** LoadTreeW: " & ex.Message)       'directory renaming error
                End Try
            End If
            NodeX = trvW.Nodes.Add(FullLocalName, LocalName, "_" & TreeCode & "_" & mImageStyle)    'adds the tree root node with related image
            NodeX.Tag = FullLocalName
            mTreeNodesCount = mTreeNodesCount + 1
            LoadExportedChildNodes(TreeCode, dirSel, NodeX)          'loads child nodes (children) if any
            Exit For
        Next
        If DirName = "" Then                                         'such a tree root does not exist (may be a new dimension)
            My.Computer.FileSystem.CreateDirectory(FullLocalName)              'creates related tree root directory
            dirSel = My.Computer.FileSystem.GetDirectoryInfo(FullLocalName)
            NodeX = trvW.Nodes.Add(FullLocalName, LocalName, "_" & TreeCode & "_" & mImageStyle)    'adds the tree root node with related image
            NodeX.Tag = FullLocalName
            mTreeNodesCount = mTreeNodesCount + 1
            LoadExportedChildNodes(TreeCode, dirSel, NodeX)          'loads child nodes (children) if any
        End If
    End Sub

    Private Sub LoadExportedTrees()
        'Loads exported trees (into Windows directories) : Who (o), What (a), When (t), Where (e) and What else (b) and Other (c) in ZAAN-Pro mode

        'Debug.Print("=> LoadExportedTrees...")
        Me.Cursor = Cursors.WaitCursor                     'sets wait cursor
        fswTree.EnableRaisingEvents = False                'locks fswTree related events

        trvW.Nodes.Clear()
        mTreeNodesCount = 0                                'resets tree nodes counter

        LoadExportedTreeW("o", 1)                          'loads Who tree, starting at root
        LoadExportedTreeW("a", 2)                          'loads What tree, starting at root
        LoadExportedTreeW("t", 3)                          'loads When tree, starting at root
        LoadExportedTreeW("e", 4)                          'loads Where tree, starting at root

        If mLicTypeCode >= 30 Then                         'case of "Pro" license (6 dimensions) :
            LoadExportedTreeW("b", 5)                      'loads What else tree, starting at root
            LoadExportedTreeW("c", 6)                      'loads Other tree, starting at root
        End If
        trvW.Sort()

        fswTree.EnableRaisingEvents = True                 'unlocks fswTree related events
        Me.Cursor = Cursors.Default                        'resets default cursor
    End Sub

    Private Sub LoadInputChildNodes(ByVal dirInfo As System.IO.DirectoryInfo, ByVal ParentNode As TreeNode, ByVal SubLevel As Integer)
        'Loads child nodes (if any) of given parent node
        Dim dirSels As System.IO.DirectoryInfo()
        Dim dirSel As System.IO.DirectoryInfo
        Dim NodeX As TreeNode
        Dim NodeTab As TreeNode()
        Dim testInfoName As String

        'Debug.Print("LoadInputChildNodes :  SubLevel = " & SubLevel & "  ParentNode = " & ParentNode.Name)     'TEST/DEBUG
        If SubLevel < 2 Then                                         'limits recursivity to 2 levels
            SubLevel = SubLevel + 1
            dirSels = dirInfo.GetDirectories
            For Each dirSel In dirSels
                If (dirSel.Attributes And FileAttributes.Hidden) <> FileAttributes.Hidden Then
                    NodeX = Nothing
                    Try
                        'Debug.Print("dir found : " & dirSel.Name & "  " & dirSel.FullName)     'TEST/DEBUG
                        testInfoName = dirSel.Name                       'dirSel object can be accessed only if current Windows user has related access rights
                        NodeTab = trvInput.Nodes.Find(dirSel.FullName, True)
                        If NodeTab.Length > 0 Then
                            NodeX = NodeTab(0)
                        Else
                            NodeX = ParentNode.Nodes.Add(dirSel.FullName, dirSel.Name, "folder")    'adds found directory as a child node
                            NodeX.Tag = dirSel.FullName
                        End If
                        LoadInputChildNodes(dirSel, NodeX, SubLevel)     'loads child nodes if any
                    Catch ex As Exception
                        'Debug.Print(dirSel.Name & " : " & ex.Message)   'error message !
                        If Not NodeX Is Nothing Then
                            NodeX.Remove()                               'remove directory node that cannot be accessed (ex: Music, Pictures, Videos, etc. )
                        End If
                    End Try
                End If
            Next
        End If
    End Sub

    Private Sub LoadInputTree()
        'Loads input tree with Windows directory structure starting at selected root
        Dim dirInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(mInputPath)
        Dim NodeX As TreeNode
        Dim FolderFileType As String = GetFileTypeAndImage("folder")

        'Debug.Print("LoadInputTree...")   'TEST/DEBUG
        Me.Cursor = Cursors.WaitCursor                     'sets wait cursor
        trvInput.Nodes.Clear()

        NodeX = trvInput.Nodes.Add(dirInfo.FullName, dirInfo.Name, "folder")   'adds the tree root node with related document count
        NodeX.Tag = dirInfo.FullName
        LoadInputChildNodes(dirInfo, NodeX, 0)             'loads child nodes if any

        trvInput.Sort()
        trvInput.SelectedImageKey = "folder"
        NodeX.Expand()
        Me.Cursor = Cursors.Default                        'resets default cursor
    End Sub

    Private Function GetExportDimTreeNodePath(ByVal FileNameKey As String, ByVal TreeCode As String, ByVal DimNb As Integer) As String
        'Returns Unix type tree node path (with "/") of the selected tree (c) of given key if existing, empty string if not
        Dim NodeKey, WhenKey, NodePath, DimNodePath As String
        Dim p, q As Integer
        Dim NodeX(10) As TreeNode

        'Debug.Print("GetExportDimTreeNodePath (" & FileNameKey & ", " & TreeCode & ")")
        DimNodePath = ""
        p = 1
        Do
            p = InStr(p, FileNameKey, "_" & TreeCode, vbTextCompare)           'searches for TreeCode occurences in key
            If p > 0 Then
                If TreeCode = "t" Then
                    q = InStr(p + 2, FileNameKey, "_", vbTextCompare)          'searches for next tree code separator
                    If q = 0 Then
                        WhenKey = Mid(FileNameKey, p + 2)
                    Else
                        WhenKey = Mid(FileNameKey, p + 2, q - p - 2)
                    End If
                    'NodePath = GetDateTextFromWhenV3Key(WhenKey)              'returns date text (YYYY-MM-DD at max) from given WhenKey in v3 database format
                    'DimNodePath = DimNodePath & "    <dim>" & NodePath & "</dim>" & vbCrLf
                    NodeKey = "t" & WhenKey
                    NodeX = trvW.Nodes.Find(NodeKey, True)
                    If NodeX.Length > 0 Then                                   'node found
                        NodePath = Replace(NodeX(0).FullPath, "\", "/")        'replaces Windows "\" separators by Unix "/" separators
                        DimNodePath = DimNodePath & "    <dim>" & NodePath & "</dim>" & vbCrLf
                    End If
                    Exit Do
                Else
                    NodeKey = Mid(FileNameKey, p + 1, mTreeKeyLength)
                    NodeX = trvW.Nodes.Find(NodeKey, True)
                    If NodeX.Length > 0 Then                                   'node found
                        NodePath = Replace(NodeX(0).FullPath, "\", "/")        'replaces Windows "\" separators by Unix "/" separators
                        DimNodePath = DimNodePath & "    <dim>" & NodePath & "</dim>" & vbCrLf
                    End If
                End If
                'Debug.Print("GetExportDimTreeNodePath:  " & NodeKey & " => " & NodePath)     'TEST/DEBUG
                p = p + 2
            End If
        Loop Until p = 0
        GetExportDimTreeNodePath = DimNodePath
    End Function

    Private Function GetTreeNodeKey(ByVal NodeX As TreeNode) As String
        'Returns the node key of given tree node (empty if tree root node)
        Dim TreeCode As String = Mid(NodeX.Tag, 2, 1)
        Dim NodeKey As String = ""

        If TreeCode = "u" Then                                       'is an access control node
            NodeKey = Mid(NodeX.Tag, 2)
        Else
            NodeKey = Mid(NodeX.Tag, 2, mTreeKeyLength)
        End If
        If Mid(NodeKey, 2, mTreeKeyLength - 1) = mTreeRootKey Then   'case of a tree root destination node
            NodeKey = ""                                             'erases destination node key
        End If
        'Debug.Print("GetTreeNodeKey :  NodeX.Text = " & NodeX.Text & "  NodeX.Tag = " & NodeX.Tag & " => NodeKey = " & NodeKey)   'TEST/DEBUG
        GetTreeNodeKey = NodeKey
    End Function

    Private Function GetTreeNodeText(ByVal NodeKey As String) As String
        'Returns the node text of given NodeKey if existing, empty text if not
        Dim TreeCode As String = Mid(NodeKey, 1, 1)
        Dim NodeText As String
        Dim NodeX() As TreeNode

        If TreeCode = "t" Then                                       'in v3 database format, dates (YYYY-MM-DD) are extractable from related NodeKey
            NodeText = GetDateTextFromWhenV3Key(Mid(NodeKey, 2))     'returns date text (YYYY-MM-DD at max) from given WhenKey in v3 database format
        Else
            NodeText = ""
            NodeX = trvW.Nodes.Find(NodeKey, True)                   'related node found
            If NodeX.Length > 0 Then
                NodeText = NodeX(0).Text                             'get node text
            End If
        End If
        'Debug.Print("GetTreeNodeText :  NodeKey = " & NodeKey & "  => NodeText = " & NodeText)   'TEST/DEBUG
        GetTreeNodeText = NodeText
    End Function

    Private Function GetFileTreeNodeText(ByRef FileNameKey As String, ByVal TreeCode As String) As String
        'Returns node text (list for multiple Who/What) of tree node(s) of given tree code found in given filename key if existing, empty text if not
        Dim NodeKey, WhenKey, NodeText As String
        Dim p, q As Integer
        Dim NodeX() As TreeNode

        'Debug.Print("GetFileTreeNodeText (" & FileNameKey & ", " & TreeCode & ")")
        NodeText = ""
        p = 1
        Do
            p = InStr(p, FileNameKey, "_" & TreeCode, vbTextCompare)           'searches for TreeCode occurences in given key
            If p > 0 Then
                If TreeCode = "t" Then                                         'in v3 database format, dates (YYYY-MM-DD) are extractable from related key
                    q = InStr(p + 2, FileNameKey, "_", vbTextCompare)          'searches for next tree code separator
                    If q = 0 Then
                        WhenKey = Mid(FileNameKey, p + 2)
                    Else
                        WhenKey = Mid(FileNameKey, p + 2, q - p - 2)
                    End If
                    NodeText = GetDateTextFromWhenV3Key(WhenKey)               'returns date text (YYYY-MM-DD at max) from given WhenKey in v3 database format
                    Exit Do
                Else
                    NodeKey = Mid(FileNameKey, p + 1, mTreeKeyLength)
                    NodeX = trvW.Nodes.Find(NodeKey, True)
                    If NodeX.Length > 0 Then                                   'node found
                        If NodeText = "" Then
                            NodeText = NodeX(0).Text                           'initializes output with new text
                        Else
                            NodeText = NodeText & ", " & NodeX(0).Text         'adds new text to output (list for multiple Who/What)
                        End If
                    End If
                End If
                p = p + 2
            End If
        Loop Until p = 0
        GetFileTreeNodeText = NodeText
    End Function

    Private Function GetFileTreeNodeKeys(ByRef FileNameKey As String, ByVal TreeCode As String) As String
        'Returns node key(s) of given tree code found in given filename keys (empty string if not found) in v3 database format
        Dim NodeKey, Key As String
        Dim p, q As Integer

        'Debug.Print("GetFileTreeNodeKeys (" & FileNameKey & ", " & TreeCode & ")")    'TEST/DEBUG
        NodeKey = ""
        p = 1
        Do
            p = InStr(p, FileNameKey, "_" & TreeCode, vbTextCompare)           'searches for TreeCode occurences in given key
            If p > 0 Then
                q = InStr(p + 2, FileNameKey, "_", vbTextCompare)              'searches for next tree code separator (variable key length in v3)
                If q = 0 Then
                    Key = Mid(FileNameKey, p + 1)
                Else
                    Key = Mid(FileNameKey, p + 1, q - p - 1)
                End If
                If TreeCode = "t" Then                                         'no multiple When keys possible
                    NodeKey = Key
                    Exit Do
                Else
                    If NodeKey = "" Then
                        NodeKey = Key
                    Else
                        NodeKey = NodeKey & "_" & Key                          'builds multiple keys series
                    End If
                End If
                p = p + 2
            End If
        Loop Until p = 0
        GetFileTreeNodeKeys = NodeKey
    End Function

    Private Function GetFileTreeNodeKey(ByRef FileNameKey As String, ByVal TreeCode As String) As String
        'Returns first node key of given tree code found in given filename keys (empty string if not found)
        Dim NodeKey As String
        Dim p, q As Integer

        'Debug.Print("GetFileTreeNodeKey (" & FileNameKey & ", " & TreeCode & ")")    'TEST/DEBUG
        NodeKey = ""
        p = InStr(1, FileNameKey, "_" & TreeCode, vbTextCompare)          'searches for TreeCode occurences in given key
        If p > 0 Then
            If TreeCode = "t" Then                                        'in v3 database format, dates (YYYY-MM-DD) are extractible from related key
                q = InStr(p + 2, FileNameKey, "_", vbTextCompare)         'searches for next tree code separator
                If q = 0 Then
                    NodeKey = Mid(FileNameKey, p + 1)
                Else
                    NodeKey = Mid(FileNameKey, p + 1, q - p - 1)
                End If
            Else
                NodeKey = Mid(FileNameKey, p + 1, mTreeKeyLength)
            End If
        End If
        GetFileTreeNodeKey = NodeKey
    End Function

    Private Function CheckBasicDocPath(ByVal DirPathName As String) As String
        'Returns validated directory name depending on current ZAAN license

        'Debug.Print("CheckBasicDocPath :  DirPathName = " & DirPathName)   'TEST/DEBUG
        If Not My.Computer.FileSystem.DirectoryExists(DirPathName) Then   'case of given directory doesn't exist
            DirPathName = mZaanDemoPath & "ZAAN-Demo1\"                    'resets demo directory by default
            mFileFilter = ""
        End If
        CheckBasicDocPath = DirPathName
    End Function

    Private Function IsOldWhenKeyFormat()
        'Returns true if old when key (v3-/9 based) format is used in current database (9 complements per digit, ex: t7988 is key of 2011 year)
        Dim is9base As Boolean = False
        Dim dirInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "tree\")
        Dim fi As System.IO.FileInfo
        Dim FileFilter, Title3, Key3 As String

        FileFilter = "_t*_t" & mTreeRootKey & "*.txt"
        For Each fi In dirInfo.GetFiles(FileFilter)
            Title3 = Mid(fi.Name, 5 + 2 * mTreeKeyLength, 3)
            Key3 = Mid(fi.Name, 3, 3)
            If Key3 = GetWhenV3KeyFromDateText(Title3, False) Then        'is a base 9 code in old v3 database format
                is9base = True
                Exit For
            End If
        Next
        'Debug.Print("IsOldWhenKeyFormat :  is9base = " & is9base)   'TEST/DEBUG
        IsOldWhenKeyFormat = is9base
    End Function

    Private Sub GetFileFilterIni()
        'Initializes mFileFilter of current ZAAN database (set by mZaanDbPath) from db\info\zaan_computer.ini
        Dim FilePathName, FilePathName2, FileContent, FileLines(), LineCells() As String
        'Dim Response As MsgBoxResult
        Dim i As Integer

        'Debug.Print("GetFileFilterIni...")   'TEST/DEBUG

        Call CheckTreeKeyLength(mZaanDbName)                              'checks tree key length of currently open ZAAN database (is 16 since 28/9/2010 and was 13 before) and updates it in mTreeKeyLength

        mFileFilter = ""
        Call InitLvBookmark()                                             'initializes bookmark list

        'tsmiSelectorBookmarkVisible.CheckState = CheckState.Unchecked
        mColDisplayIndexes = ""
        mColDisplayWidths = ""
        mColSelectorWidths = ""
        mZaanExportDest = ""
        mXmode = ""
        mIsImageView = False                                              'sets image view to false (thus detail view) by default
        mTreeCodeSeriesIndex = 1                                          'initializes index to 13 for "t o a e b c" tree code series
        mTreeCodeSeries = "t o a e b c"                                   'initializes tree code series to "t o a e b c" 
        mExportDBwinRepeatIndex = 0                                       'initializes db export repeat to no repeat
        mExportDBwinTime = ""                                             'initializes db export time
        mExportDBwinDest = ""                                             'initializes db export destination path
        mTreeViewLocked = False                                           'flag indicating if tree view edition is locked (false by default)

        FileContent = ""

        FilePathName = mZaanDbPath & "info\zaan_" & mUserEmail & ".ini"   'NEW SINCE 2012-02-20 : searches for user (only) related *.ini file
        If My.Computer.FileSystem.FileExists(FilePathName) Then           'case of *.ini file exists
            Try
                FileContent = My.Computer.FileSystem.ReadAllText(FilePathName)
            Catch ex As Exception
                'MsgBox(ex.Message, MsgBoxStyle.Exclamation)               'file not found !
                Debug.Print(ex.Message)                                   'file not found !
            End Try
        Else
            FilePathName2 = mZaanDbPath & "info\zaan_" & LCase(My.Computer.Name) & ".ini"     'searches for user/computer related *.ini file
            If My.Computer.FileSystem.FileExists(FilePathName2) Then      'case of *.ini file exists
                Try
                    FileContent = My.Computer.FileSystem.ReadAllText(FilePathName2)
                Catch ex As Exception
                    'MsgBox(ex.Message, MsgBoxStyle.Exclamation)          'file not found !
                    Debug.Print(ex.Message)                              'file not found !
                End Try
            End If
        End If

        If FileContent <> "" Then
            FileLines = Split(FileContent, vbCrLf)
            For i = 0 To FileLines.Length - 1
                'Debug.Print("line " & i & " : " & FileLines(i))
                LineCells = Split(FileLines(i), vbTab)
                Select Case LineCells(0)
                    Case "Zdbf"                                           'ZAAN database filter
                        mFileFilter = LineCells(1)
                    Case "Cdi"                                            'LvIn columns display index order
                        mColDisplayIndexes = LineCells(1)
                    Case "Cdw"                                            'LvIn columns display width
                        mColDisplayWidths = LineCells(1)
                    Case "Csw"                                            'LvSelector columns display width
                        mColSelectorWidths = LineCells(1)
                    Case "Exp"
                        mZaanExportDest = LineCells(1)                    'sets ZAAN cube export destination path
                    Case "Xmode"
                        mXmode = LineCells(1)
                    Case "Imgv"
                        mIsImageView = LineCells(1)                       'sets current image view/detail view for displaying selected files
                    Case "Tcsi"
                        mTreeCodeSeriesIndex = LineCells(1)               'sets current tree codes series index (ex : 1 for "t o a e b c")
                    Case "Tcs"
                        mTreeCodeSeries = LineCells(1)                    'sets current tree codes series (ex : "t o a e b c")
                    Case "Expdbri"
                        mExportDBwinRepeatIndex = LineCells(1)            'sets current db export repeat index
                    Case "Expdbt"
                        mExportDBwinTime = LineCells(1)                   'sets current db export time
                    Case "Expdbd"
                        mExportDBwinDest = LineCells(1)                   'sets current db export destination path
                        If mExportDBwinDest = "" Then
                            Dim dirDbInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath)
                            mExportDBwinDest = dirDbInfo.Parent.FullName  'sets by default current db path name as db export destination path
                        End If
                    Case "Tvl"
                        mTreeViewLocked = LineCells(1)                    'sets flag indicating if tree view is locked
                    Case "Zbmd"
                        Call AddBookmark(LineCells(1), True)              'adds default bookmark to lvBookmark list
                        Dim FileFilter() As String = Split(LineCells(1), " ")
                        mFileFilter = FileFilter(0)                       'resets selector to default bookmark
                    Case Else
                        If Mid(LineCells(0), 1, 3) = "Zbm" Then           'is a bookmark
                            Call AddBookmark(LineCells(1))                'adds bookmark to lvBookmark list
                        End If
                End Select
            Next
        End If

        'Debug.Print("GetFileFilterIni :  mFileFilter = " & mFileFilter)    'TEST/DEBUG
        'Debug.Print("GetFileFilterIni :  mIsImageView = " & mIsImageView)    'TEST/DEBUG

        If mXmode = "auto" Then                                           'HIDDEN Automatic cube import function !
            tsmiSelectorAutoImport.CheckState = CheckState.Checked
        Else
            tsmiSelectorAutoImport.CheckState = CheckState.Unchecked
        End If
        Call SetLvInDisplayMode(mIsImageView)                        '=> forces stored view mode for displaying selected documents
        Call SetTreeViewLockMode()                                   'updates tree view lock control state in Selector's menu and related tree edition capabilities

        If mExportDBwinTime = "" Then                                'case of manual db export time set
            tmrDbExport.Enabled = False                              'stops db export timer
            mZaanTitleOption = ""                                    'clears form title display option
            tsmiSelectorExportDBwin.Checked = False                  'unchecks export control in selector menu
        Else                                                         'case of automatic db export time set
            tmrDbExport.Enabled = True                               'starts db export timer
            mZaanTitleOption = " (auto)"                             'sets form title display option with auto (export) indicator
            tsmiSelectorExportDBwin.Checked = True                   'checks export control in selector menu
        End If
        If mExportDBwinDest = "" Then
            mExportDBwinDest = "C:\"
        End If
    End Sub

    Private Sub SetSelectionPanelDisplay()
        'Sets selection panel display with cube tabs if import panel visible, without cube tabs else (simple cube display)

        If SplitContainer1.Panel2Collapsed Then                      'import panel is hidden
            pnlCube.Parent = SplitContainer1.Panel1                  'sets simple cube display without cube tabs
            tcCubes.Visible = False
            Me.Padding = New System.Windows.Forms.Padding(0)
        Else                                                         'import panel is visible
            pnlCube.Parent = tcCubes.TabPages(tcCubes.SelectedIndex) 'sets multiple cubes tabs display
            tcCubes.Visible = True
            Me.Padding = New System.Windows.Forms.Padding(5)
        End If
    End Sub

    Private Sub SetFormAndPanelsSizes()
        'Sets main form position and size and panels size
        Dim PanelsWith(), FormLocSize() As String
        Dim LeftPanelWidth As Integer = 0
        Dim RightPanelWidth As Integer = 0
        Dim BottomPanelHeight As Integer = 0
        Dim BottomPanelWidth As Integer = 0
        Dim minSize As Integer = 400
        Dim i As Integer

        If mPanelWidths = "" Then
            SplitContainer3.Panel2Collapsed = True                   'hides document viewer panel
            SplitContainer4.Panel1Collapsed = True                   'hides input folders tree panel
        Else
            PanelsWith = Split(mPanelWidths, " ")
            For i = 0 To PanelsWith.Length - 1
                Select Case i
                    Case 0
                        LeftPanelWidth = PanelsWith(i)
                    Case 1
                        RightPanelWidth = PanelsWith(i)
                    Case 2
                        BottomPanelHeight = PanelsWith(i)
                    Case 3
                        BottomPanelWidth = PanelsWith(i)
                End Select
            Next

            With SplitContainer2
                If LeftPanelWidth = 0 Then                           'left panel (tree view) is hidden
                    .Panel1Collapsed = True
                Else
                    .Panel1Collapsed = False
                    If (.Panel1MinSize <= LeftPanelWidth) And (LeftPanelWidth <= .Width - .Panel2MinSize) Then
                        .SplitterDistance = LeftPanelWidth
                    End If
                End If
            End With
            Call SetLeftPanelButton()                                'sets left panel button display depending on splitter and background color

            With SplitContainer3
                If RightPanelWidth = 0 Then                          'right panel (document viewer) is hidden
                    .Panel2Collapsed = True
                Else
                    .Panel2Collapsed = False
                    If (.Panel1MinSize <= RightPanelWidth) And (RightPanelWidth <= .Width - .Panel2MinSize) Then
                        .SplitterDistance = RightPanelWidth
                    End If
                End If
            End With
            Call SetRightPanelButton()                               'sets right panel button image depending on splitter and background color

            With SplitContainer1
                If BottomPanelHeight = 0 Then                        'bottom panel (import/copy) is hidden
                    .Panel2Collapsed = True
                Else
                    .Panel2Collapsed = False
                    If (.Panel1MinSize <= BottomPanelHeight) And (BottomPanelHeight <= .Width - .Panel2MinSize) Then
                        .SplitterDistance = BottomPanelHeight
                    End If
                End If
            End With
            Call SetBottomPanelButton()                              'sets bottom panel button image depending on splitter and background color
            Call SetSelectionPanelDisplay()                          'sets selection panel display with cube tabs if import panel visible, without cube tabs else

            With SplitContainer4
                If BottomPanelWidth = 0 Then                         'bottom left (Windows folder tree view) is hidden
                    .Panel1Collapsed = True
                    tsmiLvOutFoldersVisible.CheckState = CheckState.Unchecked
                Else
                    .Panel1Collapsed = False
                    If (.Panel1MinSize <= BottomPanelWidth) And (BottomPanelWidth <= .Width - .Panel2MinSize) Then
                        .SplitterDistance = BottomPanelWidth
                    End If
                    tsmiLvOutFoldersVisible.CheckState = CheckState.Checked
                End If
            End With
        End If
        'lvSelector.Scrollable = SplitContainer3.Panel2Collapsed

        If mFormLocSize = "" Then
            Me.Width = 1001
        Else
            FormLocSize = Split(mFormLocSize, " ")
            If FormLocSize.Length >= 4 Then
                If FormLocSize(0) < 0 Then
                    FormLocSize(0) = 0                               'makes sure than X is always >=0
                End If
                If FormLocSize(1) < 0 Then
                    FormLocSize(1) = 0                               'makes sure than Y is always >=0
                End If
                If FormLocSize(2) < minSize Then
                    FormLocSize(2) = minSize                         'makes sure than Width is always >=100
                End If
                If FormLocSize(3) < minSize Then
                    FormLocSize(3) = minSize                         'makes sure than Height is always >=100
                End If
                Me.SetBounds(FormLocSize(0), FormLocSize(1), FormLocSize(2), FormLocSize(3))       'sets X, Y, Width, Height of main ZAAN form
            End If
        End If
        'Debug.Print("SetFormAndPanelsSizes :  SplitContainer4.SplitterDistance = " & SplitContainer4.SplitterDistance)      'TEST/DEBUG
    End Sub

    Private Sub SaveCopyZaanDatabase()
        'Saves a copy of current ZAAN database (directory) near it, with a time stamp
        Dim SourceDirPathName, DestDirPathName As String

        If Microsoft.VisualBasic.Right(mZaanDbPath, 1) = "\" Then
            SourceDirPathName = Microsoft.VisualBasic.Left(mZaanDbPath, Len(mZaanDbPath) - 1)
        Else
            SourceDirPathName = mZaanDbPath
        End If
        DestDirPathName = SourceDirPathName & "_archive_" & Format(Now, "yyyy-MM-dd_HH-mm-ss")

        Me.Cursor = Cursors.WaitCursor                               'sets wait cursor
        Try
            My.Computer.FileSystem.CopyDirectory(SourceDirPathName, DestDirPathName, Microsoft.VisualBasic.FileIO.UIOption.AllDialogs)     'copies ZAAN database directory
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)              'displays error message
        End Try
        Me.Cursor = Cursors.Default                                  'resets default cursor
    End Sub

    Private Sub CheckTreeKeyLength(ByVal ZaanDbName As String)
        'Checks tree key length of currently open ZAAN database (is 16 since 28/9/2010 and was 13 before) and updates it in mTreeKeyLength
        Dim dbFormatOK As Boolean = True
        Dim PathFileName, FileName, FileNameSplit(), NodeKey As String
        Dim Response As MsgBoxResult
        Dim V3convToBeDone As Boolean = False
        Dim CurDate As DateTime = Now

        'Debug.Print("CheckTreeKeyLength :  ZaanDbName = " & ZaanDbName)   'TEST/DEBUG

        PathFileName = Dir(mZaanDbPath & "tree\_a*__*.txt")          'searches for first existing What tree file
        If PathFileName = "" Then
            mTreeKeyLength = 0                                       'no node keys in v4 (use directly Windows path names)
            'Debug.Print("CheckTreeKeyLength :  mTreeKeyLength = " & mTreeKeyLength)   'TEST/DEBUG
            dbFormatOK = False                                       'database format is not valid (TEST NOTE : O key length uses too long Windows paths and cubes searching is too slow !)
        Else
            mTreeKeyLength = 5                                       'stores current key length in global variable mTreeKeyLength
            mTopRootKey = "zzzz"
            mTreeRootKey = "zzzy"

            FileName = System.IO.Path.GetFileName(PathFileName)
            'Debug.Print("CheckTreeKeyLength :  FileName = " & FileName)   'TEST/DEBUG
            FileNameSplit = Split(FileName, "_a")
            If FileNameSplit.Length > 1 Then
                NodeKey = "a" & FileNameSplit(1)                               'get first What key from tree nodes list
                'Debug.Print("CheckTreeKeyLength :  NodeKey.Length = " & NodeKey.Length)   'TEST/DEBUG
                If NodeKey.Length <> 5 Or IsOldWhenKeyFormat() Then
                    dbFormatOK = False                                         'database format is not valid (NOTE : old v1/v2/v3- formats used before 16/03/2011 have been retired by 5/01/2012)
                End If
            End If
        End If

        If Not dbFormatOK Then
            Response = MsgBox(mMessage(133), MsgBoxStyle.Exclamation)          'This ZAAN database format is not valid !
            Call DeleteZaanDbFromTabs()                                        'deletes currently selected ZAAN database from tabs and selects first db from list or ZAAN-Demo1 if empty
            Call UpdateZaanDbRootPathName(mZaanDemoPath & "ZAAN-Demo1\")        'updates current database selection with mZaanDbRoot, mZaanDbPath and mZaanDbName

            Call AddZaanPathInTabs()                                           'adds mZaanDbPath directory in database/cube tabs if not already listed and updates its selection

            Call GetFileFilterIni()                                            'initializes mFileFilter of current ZAAN database from db\info\zaan.ini
            Call InitFileDataWatcher()                                         'initializes "data","tree" and "xin" file watchers to detect directory access/change in current ZAAN database
            Call InitTreeInOutFiles(False)                                     'displays selected Zaan database name, logo, tree, selector and selected files views, but not "imput", "out" and "temp" views
        End If
        'Debug.Print("=> CheckTreeKeyLength :  mTreeKeyLength = " & mTreeKeyLength)   'TEST/DEBUG
    End Sub

    Private Sub InitZaanDbPath()
        'Initializes mZaanDbPath with eventually launched .zaan file content (empty if direct ZAAN application launch)
        Dim FileRef, FilePathName, FileContent, FileLines(), LineCells() As String
        Dim i As Integer

        mZaanDbRoot = ""
        mZaanDbPath = ""                                             'resets main ZAAN database path
        mZaanDbName = ""

        FileRef = ""
        Try
            If Not AppDomain.CurrentDomain.SetupInformation.ActivationArguments Is Nothing Then
                If Not AppDomain.CurrentDomain.SetupInformation.ActivationArguments.ActivationData Is Nothing Then
                    FileRef = AppDomain.CurrentDomain.SetupInformation.ActivationArguments.ActivationData(0)   'gets file reference used to launch ZAAN application
                End If
            End If
        Catch ex As Exception
            Debug.Print(ex.Message)                                  'TEST/DEBUG
        End Try

        'FileRef = "file:///C:/Users/Emmanuel/Desktop/cube%20test%201.zaan"     'FORCED VALUE FOR TEST/DEBUG ONLY !!!

        If FileRef <> "" Then
            Dim FileUri As New Uri(FileRef)
            FilePathName = FileUri.LocalPath                         'gets local operating-system path file name
            'Debug.Print("InitZaanDbPath :  FilePathName = " & FilePathName)    'TEST/DEBUG
            'MsgBox("InitZaanDbPath : FilePathName = " & FilePathName, vbInformation)   'TEST/DEBUG

            FileContent = ""
            If My.Computer.FileSystem.FileExists(FilePathName) Then  'case of *.zaan file exists
                Try
                    FileContent = My.Computer.FileSystem.ReadAllText(FilePathName)
                Catch ex As Exception
                    Debug.Print(ex.Message)                          'TEST/DEBUG
                End Try
            End If

            If FileContent <> "" Then
                FileLines = Split(FileContent, vbCrLf)
                For i = 0 To FileLines.Length - 1
                    'Debug.Print("line " & i & " : " & FileLines(i))
                    LineCells = Split(FileLines(i), vbTab)
                    Select Case LineCells(0)
                        Case "Zdb"
                            Call UpdateZaanDbRootPathName(LineCells(1))        'updates current database selection with mZaanDbRoot, mZaanDbPath and mZaanDbName
                            'Debug.Print("InitZaanDbPath :  mZaanDbPath = " & mZaanDbPath)    'TEST/DEBUG
                            'MsgBox(" => InitZaanDbPath : mZaanDbPath = " & mZaanDbPath, vbInformation)   'TEST/DEBUG
                    End Select
                Next
            End If
        End If
        'Debug.Print("InitZaanDbPath :  mZaanDbPath = " & mZaanDbPath)     'TEST/DEBUG
    End Sub

    Private Sub AddDatabaseCubeTab(ByVal TabCubeIndex As Integer, ByVal DbPathName As String)
        'Adds or updates first cube tab at given tab index with given db pathname
        Dim DbName As String = GetLastDirName(DbPathName)

        If TabCubeIndex = 0 Then                                     'first cube tab (index = 0) already exists (set at design time)
            tcCubes.TabPages(0).Text = DbName
        Else                                                         'adds a new cube tab
            tcCubes.TabPages.Add(DbName)
        End If
        tcCubes.TabPages(TabCubeIndex).ToolTipText = DbPathName
        tcCubes.TabPages(TabCubeIndex).ImageKey = "_x_" & mImageStyle
    End Sub

    Private Sub GetZaanIni()
        'Initializes mLanguage, mLicenseAccepted, mOpacity, mImageStyle, mUserEmail, mLicenseKey, mLicTypeText, mZaanDbPath and mFileFilter from "zaan.ini" files if existing
        Dim LanguageCode2, FilePathName, FileContent, FileLines(50), LineCells(3) As String
        Dim n, i, TabCubeIndex As Integer

        'Debug.Print("GetZaanIni :  My.Application.Culture.Name = " & My.Application.Culture.Name)   'TEST/DEBUG
        LanguageCode2 = UCase(Mid(My.Application.Culture.Name, 1, 2))     'gets default system language code : WARNING : MAY CONFLICT WITH ctfmon.exe PROCESS/Office 2003 !
        'LanguageCode2 = "FR"                                         'BYPASS EVENTUAL ctfmon.exe CONFLICT !?

        If LanguageCode2 = "FR" Then
            mLanguage = "FR"                                         'sets French language by default
        Else
            mLanguage = "EN"                                         'sets English language by default
        End If
        'Debug.Print("mLanguage = " & mLanguage)   'TEST/DEBUG

        mOpacity = 100                                               'sets opacity to 100% by default
        mImageStyle = 3                                              'sets image style to 3 by default (colors icons and white background)

        mUserEmail = ""
        mLicenseKey = ""
        mUserLic2Line = ""                                           'available for storing a second computer Key line, if any existing
        mLicTypeText = ""                                            'sets ZAAN license type to none (<=> no valid license key) by default

        mInputPath = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\"
        mImportPath = mInputPath

        mLicTryNb = "0"                                              'sets number of licence (key) entry trials to 0 (TestUserLicence() will limit it to 3 trials maxi)
        mLicLastCheck = ""                                           'sets last check date of licence key to nothing
        mPanelWidths = "0 0 389 0"
        mFormLocSize = "0 0 1001 580"
        mColDisplayIndexes = ""
        mColDisplayWidths = ""
        mColSelectorWidths = ""
        mColImportWidths = ""
        mColCopyWidths = ""
        mListsHeadersVisible = True
        mIsCubeView = False                                          'classic tube view is set by default

        mInDocPerPage = mInDocPerPageMax
        mLicAccepted = "0"                                           'flag to be set to 1 at user's acceptation of ZAAN license (GNU GPLv3)
        n = 0
        TabCubeIndex = 0
        FileContent = ""

        FilePathName = mZaanAppliPath & "info\zaan.ini"
        If My.Computer.FileSystem.FileExists(FilePathName) Then      'case of "zaan.ini" file exists
            Try
                FileContent = My.Computer.FileSystem.ReadAllText(FilePathName)
            Catch ex As Exception
                'MsgBox(ex.Message, MsgBoxStyle.Exclamation)          'file not found !
                Debug.Print(ex.Message)                              'file not found !
            End Try
        End If

        If FileContent <> "" Then
            FileLines = Split(FileContent, vbCrLf)
            For i = 0 To FileLines.Length - 1
                'Debug.Print("line " & i & " : " & FileLines(i))
                LineCells = Split(FileLines(i), vbTab)
                Select Case RTrim(LineCells(0))                      'eliminates eventual trailing spaces
                    Case "Lang"
                        mLanguage = LineCells(1)
                    Case "Opa"
                        'mOpacity = LineCells(1)                     'IS NOW FORCED TO 100 !
                    Case "Sty"
                        mImageStyle = LineCells(1)
                        If mImageStyle < 10 Then
                            mImageStyle = 3                          'FORCES NEW IMAGE STYLE WITH COLORS !
                        Else
                            mImageStyle = 13                         'FORCES NEW IMAGE STYLE WITH COLORS !
                        End If
                    Case "Mail"
                        mUserEmail = LineCells(1)
                    Case "Key"
                        mLicenseKey = LineCells(1)
                    Case "Key_" & LCase(My.Computer.Name)            'NEW SINCE 2012-02-20 : Key depends on computer-name
                        mLicenseKey = LineCells(1)
                    Case "Lty"
                        n = Len(LineCells(0)) - Len("Lty")           'number of licence (key) entry trials is coded as number of trailing spaces of "Lty" tag
                        mLicTryNb = n
                        mLicTypeText = LineCells(1)
                        If (mLicTypeText <> "") And (n = 0) Then     'case of valid ZAAN license type ("Basic"/"First"/"Pro") and no ongoing key try
                            mAboutZaanAutoClose = True               'enables automatic closing of "About ZAAN" window after 3 sec. (timer set in "About ZAAN" window)
                        End If
                    Case "Llc"
                        mLicLastCheck = LineCells(1)                 'get last check date of licence key
                    Case "Pnl"                                       '(4) panels width or height
                        mPanelWidths = LineCells(1)
                    Case "Fls"                                       '(main) form location (x,y) and size (width, height)
                        mFormLocSize = LineCells(1)
                    Case "Cdi"                                       'LvIn columns display index order
                        mColDisplayIndexes = LineCells(1)
                    Case "Cdw"                                       'LvIn columns display width
                        mColDisplayWidths = LineCells(1)
                    Case "Csw"                                       'LvSelector columns display width
                        mColSelectorWidths = LineCells(1)
                    Case "Ciw"                                       'LvOut columns display width
                        mColImportWidths = LineCells(1)
                    Case "Ccw"                                       'LvTemp columns display width
                        mColCopyWidths = LineCells(1)
                    Case "Zdb"                                       'is last open ZAAN database directory
                        Call AddDatabaseCubeTab(TabCubeIndex, LineCells(1))    'adds or updates first cube tab at given tab index with given db pathname
                        If mZaanDbPath = "" Then                               'case of ZAAN application not launched through a .zaan db link file
                            Call UpdateZaanDbRootPathName(LineCells(1))        'updates current database selection with mZaanDbRoot, mZaanDbPath and mZaanDbName
                            tcCubes.SelectedIndex = TabCubeIndex               'activates related tab page
                        End If
                        TabCubeIndex = TabCubeIndex + 1
                    Case "Zdi"                                       'is a ZAAN database directory that has been previously opened
                        Call AddDatabaseCubeTab(TabCubeIndex, LineCells(1))    'adds or updates first cube tab at given tab index with given db pathname
                        TabCubeIndex = TabCubeIndex + 1
                    Case "Ind"
                        mInputPath = LineCells(1)                    'sets current input directory path selection
                    Case "Imp"
                        mImportPath = LineCells(1)                   'sets current import directory selection
                    Case "Lhv"
                        mListsHeadersVisible = LineCells(1)          'sets flag related to lvIn/lvOut/lvTemp headers visibility
                    Case "Dpp"
                        mInDocPerPage = LineCells(1)                 'sets number of documents per page to be displayed in selection list (lvIn)
                    Case "LAcc"
                        mLicAccepted = LineCells(1)                  'flag indicating if ZAAN license has been accepted
                    Case Else
                        If Mid(LineCells(0), 1, 4) = "Key_" Then     'second computer Key exists
                            mUserLic2Line = FileLines(i)             'saves related second computer Key definition line
                        End If
                End Select
            Next
        End If
        If mZaanDbPath = "" Then
            Call UpdateZaanDbRootPathName(mZaanDemoPath & "ZAAN-Demo1\")   'updates current database selection with mZaanDbRoot, mZaanDbPath and mZaanDbName
        End If

        Call SetZaanControlsColors()                                 'sets background color of all ZAAN controls to given mImageStyle
        Call UpdatesListsHeadersVisibility()                         'sets lvSelector, lvIn, lvOut and lvTemp headers visibility depending on mListsHeadersVisible state

        If mLanguage <> mLanguageIni Then
            Call InitMessages()                                      'initializes mMessage table with given language (or in English by default)
            mLanguageIni = mLanguage                                 'saves initial language selection
        End If

        mLicTypeText = "Pro"                                         'SINCE v3.7 : forces ZAAN-Pro license (GNU GPLv3)
        Call SetLicTypeCode()                                        'updates mLicTypeCode from mLicTypeText value ("Basic"/"First"/"Pro")
        mLicTypeTextIni = mLicTypeText                               'saves initial license type selection

        Call UpdateZaanProButtons()                                  'updates visibility of buttons reserved to ZAAN-Pro mode

        Call UpdateZaanDbRootPathName(CheckBasicDocPath(mZaanDbPath)) 'updates current database selection with mZaanDbRoot, mZaanDbPath and mZaanDbName
        'Debug.Print("GetZaanIni:  mZaanDbPath = " & mZaanDbPath)     'TEST/DEBUG

        Call LoadOnceZaanDemoDatabase()                              'loads once, if not exists, "ZAAN-Demo1" directory
        Call AddZaanPathInTabs()                                     'adds mZaanDbPath directory in database/cube tabs if not already listed and updates its selection

        Call GetFileFilterIni()                                      'initializes mFileFilter of current ZAAN database from db\info\zaan.ini
        Call InitFileDataWatcher()                                   'initializes "data","tree" and "xin" file watchers to detect directory access/change in current ZAAN database
        Call InitFileInputWatcher()                                  'initializes input file watcher to detect updates of related directories

        Call InitFileZaanWatcher()                                   'initializes ZAAN file watcher to detect ZAAN\copy and ZAAN\import file updates
    End Sub

    Private Sub SaveFavoriteFolderShortCut(ByVal BookmarkName As String, ByVal FileFilter As String, ByRef FavPathName As String)
        'Creates and save a favorite folder short cut (to be used from Windows) corresponding to given file filter/bookmark
        Dim DirPathName, FileName, FilePathName, BookmarkPathName As String
        Dim MyItem As ListViewItem

        'Debug.Print("SaveFavoriteFolderShortCut :  BookmarkName : " & BookmarkName & "  FileFilter = " & FileFilter)    'TEST/DEBUG
        If FileFilter = "" Then Exit Sub

        BookmarkPathName = FavPathName & "\" & BookmarkName
        If CreateDirIfNotExistsOK(BookmarkPathName) Then                       'creates _fav/bookmark directory if not exists and returns true if exists or creation succeeded
            mFileFilter = FileFilter                                           'get stored file filter in selected bookmark
            Call DisplaySelector()                                             'displays selector buttons using mFileFilter selections
            Call InitDisplaySelectedFiles()                                    'initializes display of all selected files, starting at first page
            For Each MyItem In lvIn.Items                                      'scans all displayed files related to current selection (all LvIn files)
                'DirPathName = mZaanDbPath & "data\" & CStr(MyItem.Tag)         'get dir name of listed file
                DirPathName = MyItem.Tag                                       'get dir name of listed file
                FileName = MyItem.Text & MyItem.SubItems(8).Text
                FilePathName = DirPathName & "\" & FileName
                Try
                    'Call CreateWinShortcut(mZaanDbPath, FilePathName, MyItem.Text, BookmarkPathName)    'creates a directory shortcut into Windows folder pointing to ZAAN database target position
                    Call CreateWinShortcut(mZaanDbPath, FilePathName, FileName, BookmarkPathName)    'creates a directory shortcut into Windows folder pointing to ZAAN database target position
                Catch ex As Exception
                    Debug.Print("SaveFavoriteFolderShortCut : " & ex.Message)      'TEST/DEBUG
                End Try
            Next
        End If
    End Sub

    Private Sub CleanFavoriteFolder(ByRef FavPathName As String, ByVal CreateFolder As Boolean)
        'Cleans given favorite folder of current database (deletes all existing folder short cuts)
        Dim dirSel, dirInfo As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo

        If Not My.Computer.FileSystem.DirectoryExists(FavPathName) And Not CreateFolder Then Exit Sub

        If CreateDirIfNotExistsOK(FavPathName) Then                  'creates given directory if not exists and returns true if exists or creation succeeded
            dirInfo = My.Computer.FileSystem.GetDirectoryInfo(FavPathName)
            For Each fi In dirInfo.GetFiles("*")                     'deletes all files from user favorite directory
                'Debug.Print("CleanFavoriteFolder :  file = " & fi.Name)      'TEST/DEBUG
                Try
                    fi.Delete()                                      'deletes file defintively
                Catch ex As Exception
                    Debug.Print("CleanFavoriteFolder error : " & ex.Message)    'TEST/DEBUG
                End Try
            Next
            For Each dirSel In dirInfo.GetDirectories("*")           'deletes all directories from user favorite directory
                'Debug.Print("CleanFavoriteFolder :  folder = " & dirSel.Name)      'TEST/DEBUG
                For Each fi In dirSel.GetFiles("*")                  '... starts to delete related files
                    'Debug.Print("=> CFF Deletes file = " & fi.Name)      'TEST/DEBUG
                    Try
                        fi.Delete()                                      'deletes file defintively
                    Catch ex As Exception
                        Debug.Print("=> CFF Deletes file error : " & ex.Message)    'TEST/DEBUG
                    End Try
                Next
                Try
                    dirSel.Delete()                                  '... then deletes empty directory
                Catch ex As Exception
                    Debug.Print("CleanFavoriteFolder error : " & ex.Message)    'TEST/DEBUG
                End Try
            Next
        End If
    End Sub

    Private Sub SaveFileFilterIni()
        'Saves mFileFilter and other parameters of current ZAAN database in db\info\zaan.ini file
        Dim FileContent, FilePathName, OldFavPathName, FavPathName, ComputerName As String
        Dim itmX As ListViewItem
        Dim i As Integer

        'Debug.Print("SaveFileFilterIni...")    'TEST/DEBUG

        ComputerName = LCase(My.Computer.Name)                                           'get computer name

        OldFavPathName = mZaanDbPath & "_fav\" & CStr(GetUserEmailStart()) & "_" & ComputerName   'sets favorite folder path name of current user
        Call CleanFavoriteFolder(OldFavPathName, False)                                  'cleans old favorite folder if still exists
        Call DeleteDirIfEmpty(OldFavPathName)                                            'deletes it if still exists

        FavPathName = mZaanDbPath & "_fav\" & mUserEmail                                 'NEW SINCE 2012-02-20 : sets favorite folder path name of current user
        Call CleanFavoriteFolder(FavPathName, True)                                      'cleans given favorite folder from current database and creates it if necessary

        'FilePathName = mZaanDbPath & "info\zaan_" & ComputerName & ".ini"
        Call DeleteFileIfExists(mZaanDbPath & "info\zaan_" & ComputerName & ".ini")      'deletes old computer related .ini file if still exists
        FilePathName = mZaanDbPath & "info\zaan_" & mUserEmail & ".ini"                  'NEW SINCE 2012-02-20 : sets user (only) .ini file name

        FileContent = "Zdbf" & vbTab & mFileFilter & vbCrLf                              'saves current selection (mFileFilter)

        i = 0
        For Each itmX In lvBookmark.Items                                                'save bookmarked cube list
            i = i + 1
            If mTreeKeyLength = 0 Then                                                   'no node keys in v4 (use directly Windows path names)
                FileContent = FileContent & "Zbm" & i & vbTab & CStr(itmX.Tag) & vbCrLf  'only stores file filter (in textual form)
            Else
                If Microsoft.VisualBasic.Right(itmX.ImageKey, 1) = "d" Then
                    FileContent = FileContent & "Zbmd" & vbTab & CStr(itmX.Tag) & " " & itmX.Text & vbCrLf        'stores defaut cube bookmark
                Else
                    FileContent = FileContent & "Zbm" & i & vbTab & CStr(itmX.Tag) & " " & itmX.Text & vbCrLf     'stores regular cube bookmark
                End If
                'BYPASSED ED 2013/07/03 : Call SaveFavoriteFolderShortCut(itmX.Text, CStr(itmX.Tag), FavPathName)  'creates and save a favorite folder short cut corresponding to given file filter/bookmark (WARNING : modifies mFileFilter !)
            End If
        Next

        mColDisplayIndexes = ""
        mColDisplayWidths = ""
        For i = 0 To lvIn.Columns.Count - 1
            mColDisplayIndexes = mColDisplayIndexes & lvIn.Columns(i).DisplayIndex & " "
            mColDisplayWidths = mColDisplayWidths & lvIn.Columns(i).Width & " "
        Next

        'mColSelectorWidths = ""
        mColSelectorWidths = pnlSelectorSearch.Width & " "
        For i = 0 To lvSelector.Columns.Count - 1
            mColSelectorWidths = mColSelectorWidths & lvSelector.Columns(i).Width & " "
        Next
        FileContent = FileContent & "Cdi" & vbTab & mColDisplayIndexes & vbCrLf
        FileContent = FileContent & "Cdw" & vbTab & mColDisplayWidths & vbCrLf
        FileContent = FileContent & "Csw" & vbTab & mColSelectorWidths & vbCrLf

        FileContent = FileContent & "Exp" & vbTab & mZaanExportDest & vbCrLf
        FileContent = FileContent & "Xmode" & vbTab & mXmode & vbCrLf
        FileContent = FileContent & "Imgv" & vbTab & mIsImageView & vbCrLf
        FileContent = FileContent & "Tcsi" & vbTab & mTreeCodeSeriesIndex & vbCrLf
        FileContent = FileContent & "Tcs" & vbTab & mTreeCodeSeries & vbCrLf
        FileContent = FileContent & "Expdbri" & vbTab & mExportDBwinRepeatIndex & vbCrLf
        FileContent = FileContent & "Expdbt" & vbTab & mExportDBwinTime & vbCrLf
        FileContent = FileContent & "Expdbd" & vbTab & mExportDBwinDest & vbCrLf

        FileContent = FileContent & "Tvl" & vbTab & mTreeViewLocked & vbCrLf

        'Debug.Print("*** SaveFileFilterIni:  FileContent = " & FileContent)    'TEST/DEBUG
        If CreateDirIfNotExistsOK(mZaanDbPath & "info") Then         'tries to create a db "info" directory if it doesn't exist
        End If
        Try
            My.Computer.FileSystem.WriteAllText(FilePathName, FileContent, False)   'saves ZAAN\info\zaan_*.ini file
        Catch ex As System.IO.DirectoryNotFoundException
            Debug.Print(ex.Message)                                  'TEST/DEBUG
        End Try
    End Sub

    Private Sub SaveZaanIni(Optional ByVal SaveFilterIni As Boolean = True)
        'Saves mLanguage, mOpacity, mImageStyle, mUserEmail, mLicenseKey, mLicTypeText, mZaanDbPath and mFileFilter in zaan.ini files
        Dim FileContent, FilePathName, TrailingSpaces As String
        Dim n, i, k, LeftPanelWidth, RightPanelWidth, BottomPanelHeight, BottomPanelWidth As Integer

        'Debug.Print("SaveZaanIni :  SaveFilterIni = " & SaveFilterIni)    'TEST/DEBUG

        n = CInt(mLicTryNb)                                          'get number of licence (key) entry trials
        TrailingSpaces = ""
        Do While n > 0
            TrailingSpaces = TrailingSpaces & " "                    'number of trailing spaces (of "Lty" tag) will code the number of licence (key) entry trials
            n = n - 1
        Loop
        FilePathName = mZaanAppliPath & "info\zaan.ini"              'is ZAAN\info\zaan.ini file

        FileContent = "Lang" & vbTab & mLanguage & vbCrLf
        FileContent = FileContent & "Opa" & vbTab & mOpacity & vbCrLf
        FileContent = FileContent & "Sty" & vbTab & mImageStyle & vbCrLf
        FileContent = FileContent & "Mail" & vbTab & mUserEmail & vbCrLf
        'FileContent = FileContent & "Key" & vbTab & mLicenseKey & vbCrLf
        FileContent = FileContent & "Key_" & LCase(My.Computer.Name) & vbTab & mLicenseKey & vbCrLf
        If mUserLic2Line <> "" Then
            FileContent = FileContent & mUserLic2Line & vbCrLf       'save second computer key definition line
        End If

        FileContent = FileContent & "Lty" & TrailingSpaces & vbTab & mLicTypeText & vbCrLf
        FileContent = FileContent & "Llc" & vbTab & mLicLastCheck & vbCrLf

        LeftPanelWidth = SplitContainer2.SplitterDistance
        RightPanelWidth = SplitContainer3.SplitterDistance
        BottomPanelHeight = SplitContainer1.SplitterDistance
        BottomPanelWidth = SplitContainer4.SplitterDistance

        If SplitContainer2.Panel1Collapsed Then LeftPanelWidth = 0
        If SplitContainer3.Panel2Collapsed Then RightPanelWidth = 0
        If SplitContainer1.Panel2Collapsed Then BottomPanelHeight = 0
        If SplitContainer4.Panel1Collapsed Then BottomPanelWidth = 0

        FileContent = FileContent & "Pnl" & vbTab & LeftPanelWidth & " " & RightPanelWidth & " " & BottomPanelHeight & " " & BottomPanelWidth & vbCrLf
        FileContent = FileContent & "Fls" & vbTab & Location.X & " " & Location.Y & " " & Size.Width & " " & Size.Height & vbCrLf

        mColDisplayIndexes = ""
        mColDisplayWidths = ""
        For i = 0 To lvIn.Columns.Count - 1
            mColDisplayIndexes = mColDisplayIndexes & lvIn.Columns(i).DisplayIndex & " "
            mColDisplayWidths = mColDisplayWidths & lvIn.Columns(i).Width & " "
        Next
        mColSelectorWidths = ""
        For i = 0 To lvSelector.Columns.Count - 1
            mColSelectorWidths = mColSelectorWidths & lvSelector.Columns(i).Width & " "
        Next
        mColImportWidths = ""
        For i = 0 To lvOut.Columns.Count - 1
            mColImportWidths = mColImportWidths & lvOut.Columns(i).Width & " "
        Next
        mColCopyWidths = ""
        For i = 0 To lvTemp.Columns.Count - 1
            mColCopyWidths = mColCopyWidths & lvTemp.Columns(i).Width & " "
        Next
        FileContent = FileContent & "Cdi" & vbTab & mColDisplayIndexes & vbCrLf
        FileContent = FileContent & "Cdw" & vbTab & mColDisplayWidths & vbCrLf
        FileContent = FileContent & "Csw" & vbTab & mColSelectorWidths & vbCrLf
        FileContent = FileContent & "Ciw" & vbTab & mColImportWidths & vbCrLf
        FileContent = FileContent & "Ccw" & vbTab & mColCopyWidths & vbCrLf

        For k = 0 To tcCubes.TabPages.Count - 1
            If CStr(tcCubes.TabPages(k).ToolTipText) <> "" Then                   'case of non empty database => save it !
                If CStr(tcCubes.TabPages(k).ToolTipText) = mZaanDbPath Then
                    FileContent = FileContent & "Zdb" & vbTab & mZaanDbPath & vbCrLf
                Else
                    FileContent = FileContent & "Zdi" & vbTab & CStr(tcCubes.TabPages(k).ToolTipText) & vbCrLf
                End If
            End If
        Next

        FileContent = FileContent & "Ind" & vbTab & mInputPath & vbCrLf
        FileContent = FileContent & "Imp" & vbTab & mImportPath & vbCrLf
        FileContent = FileContent & "Lhv" & vbTab & mListsHeadersVisible & vbCrLf
        FileContent = FileContent & "Dpp" & vbTab & mInDocPerPage & vbCrLf
        FileContent = FileContent & "LAcc" & vbTab & mLicAccepted & vbCrLf

        If CreateDirIfNotExistsOK(mZaanAppliPath & "info") Then      'tries to create a ZAAN "info" directory if it doesn't exist
        End If
        Try
            My.Computer.FileSystem.WriteAllText(FilePathName, FileContent, False)
        Catch ex As System.IO.DirectoryNotFoundException
            Debug.Print(ex.Message)                                  'TEST/DEBUG
        End Try

        If SaveFilterIni Then
            Call SaveFileFilterIni()                                 'saves mFileFilter of current ZAAN database in db\info\zaan.ini
        End If
    End Sub

    Private Sub CreateNewDocumentFile(ByVal ShortFileType As String)
        'Creates given new document in current selection directory
        Dim SelectedDirPath As String = GetSelectedDirPathName()               'gets current data directory pathname corresponding to current mFileFilter
        Dim RootPathName As String = SelectedDirPath & "\" & mMessage(74)
        Dim FilePathName As String
        Dim n As Integer

        FilePathName = RootPathName & " 0001" & "." & ShortFileType                 'adds an 4 digits index at the end of the file name
        n = 1
        Do While My.Computer.FileSystem.FileExists(FilePathName) And (n < 9999)     'tests if the file name is not already existing, else increments its index
            n = n + 1
            FilePathName = RootPathName & " " & Format(n, "0000") & "." & ShortFileType
        Loop

        If Not My.Computer.FileSystem.FileExists(FilePathName) Then
            Try
                'Debug.Print("CreateNewDocumentFile : " & FilePathName)     'TEST/DEBUG
                Select Case ShortFileType
                    Case "txt"
                        My.Computer.FileSystem.WriteAllText(FilePathName, "", False)
                    Case "doc"
                        My.Computer.FileSystem.WriteAllBytes(FilePathName, My.Resources.new_doc, False)
                    Case "docx"
                        My.Computer.FileSystem.WriteAllBytes(FilePathName, My.Resources.new_docx, False)
                    Case "ppt"
                        My.Computer.FileSystem.WriteAllBytes(FilePathName, My.Resources.new_ppt, False)
                    Case "pptx"
                        My.Computer.FileSystem.WriteAllBytes(FilePathName, My.Resources.new_pptx, False)
                    Case "xls"
                        My.Computer.FileSystem.WriteAllBytes(FilePathName, My.Resources.new_xls, False)
                    Case "xlsx"
                        My.Computer.FileSystem.WriteAllBytes(FilePathName, My.Resources.new_xlsx, False)
                End Select

            Catch ex As Exception
                Debug.Print(ex.Message)                                  'TEST/DEBUG
            End Try
        End If
    End Sub

    Private Sub UpdateZaanProButtons()
        'Updates visibility of buttons depending on ZAAN license type ("Basic"/"First"/"Pro")

        'Debug.Print("UpdateZaanProButtons...")    'TEST/DEBUG
        If mLicTypeCode >= 30 Then                         'case of "Pro" licenses (6 dimensions)
            tsmiSelectorExportDBwin.Enabled = True         'enables export database to Windows button
            'tsmiSelectorImportDBwin.Enabled = True         'enables database import from Windows button
            'tsmiSelectorAutoImport.Enabled = True          'enables automatic import mode
            lblWhat2.Visible = True                        'shows 5th dimension in Selector
            lblWho2.Visible = True                         'shows 6th dimension in Selector
            lblDataAccess.Visible = True                   'shows Access control in Selector
            'btnDataAccessNoSel.Visible = True
            mTreeWhoRootIndex = 1                          'set trvW.Nodes() index relative to Who root node - index 0 is reserved here to Access control
        Else                                               'case of "First" and "Basic" licenses (4 dimensions)
            tsmiSelectorExportDBwin.Enabled = True         'enables export database to Windows button
            'tsmiSelectorImportDBwin.Enabled = True         'enables database import from Windows button
            'tsmiSelectorAutoImport.Enabled = True          'enables automatic import mode
            lblWhat2.Visible = False                       'hides 5th dimension in Selector
            lblWho2.Visible = False                        'hides 6th dimension in Selector
            lblDataAccess.Visible = False                  'hides Access control in Selector
            'btnDataAccessNoSel.Visible = False
            mTreeWhoRootIndex = 0                          'set trvW.Nodes() index relative to Who root node
        End If
    End Sub

    Private Sub UpdatePagePileButtons()
        'Updates page pile (lstPage) if mFileFilter is new in pile and enables/desables accordingly Prev/Next navigation buttons

        'Debug.Print(" => UpdatePagePileButtons :  mFileFilter = " & mFileFilter)    'TEST/DEBUG

        If mFileFilter <> "" Then
            If lstPage.Items.Count > 0 Then
                If lstPage.SelectedItems(0) <> mFileFilter Then
                    lstPage.Items.Add(mFileFilter)                        'adds directory in list/pile
                    lstPage.SelectedIndex = lstPage.Items.Count - 1
                End If
            Else
                lstPage.Items.Add(mFileFilter)                            'adds directory in list/pile
                lstPage.SelectedIndex = 0
            End If
        End If

        If lstPage.SelectedIndex > 0 Then
            btnPrev.Enabled = True                                   'enables Prev navigation button
            'btnPrev.Visible = True                                   'shows Prev navigation button
        Else
            btnPrev.Enabled = False                                  'disables Prev navigation button
            'btnPrev.Visible = False                                  'hides Prev navigation button
        End If
        If lstPage.SelectedIndex < lstPage.Items.Count - 1 Then
            btnNext.Enabled = True                                   'enables Next navigation button
            'btnNext.Visible = True                                   'shows Next navigation button
        Else
            btnNext.Enabled = False                                  'disables Next navigation button
            'btnNext.Visible = False                                  'hides Next navigation button
        End If
    End Sub

    Private Sub DisplaySelector()
        'Displays selector buttons using mFileFilter selections
        Dim TreeCode, Key, FileCodeTable(), Title As String
        Dim i As Integer

        'Debug.Print("DisplaySelector/begin : mFileFilter = " & mFileFilter)   'TEST/DEBUG
        'Call UpdatePagePileButtons()                       'updates page pile (lstPage) if mFileFilter is new in pile and enables/desables accordingly Prev/Next navigation buttons

        For i = 0 To 6
            lvSelector.Columns(i).Tag = ""
        Next
        lblDataAccess.Text = ""
        lblDataAccess.Tag = ""
        lblWho.Text = ""
        lblWho.Tag = ""
        lblWhat.Text = ""
        lblWhat.Tag = ""
        lblWhen.Text = ""
        lblWhen.Tag = ""
        lblWhere.Text = ""
        lblWhere.Tag = ""
        lblWhat2.Text = ""
        lblWhat2.Tag = ""
        lblWho2.Text = ""
        lblWho2.Tag = ""

        If mTreeKeyLength = 0 Then                                             'no node keys in v4 (use directly Windows path names)
            If mFileFilter <> "" Then
                FileCodeTable = Split(mFileFilter, "*")
                For i = 1 To FileCodeTable.Length - 1
                    If FileCodeTable(i) <> "" Then
                        Call UpdateSelectorTexts(-1, "", FileCodeTable(i), False)  'updates selector's current label relative to given index (in v4 database format)
                    End If
                Next
            End If
        Else
            If mFileFilter <> "" Then
                FileCodeTable = Split(mFileFilter, "*_")
                For i = 1 To FileCodeTable.Length - 1                          'scans filter node keys
                    If FileCodeTable(i) <> "" Then
                        TreeCode = Mid(FileCodeTable(i), 1, 1)
                        Key = "*_" & FileCodeTable(i)                          'sets button key for active filtering of files
                        Title = GetTreeNodeText(FileCodeTable(i))              'get node text of given key if existing, empty text if not
                        If Title = "" Then                                     'case of key not found in current tree
                            Key = ""
                            FileCodeTable(i) = ""                              'erases non valid Filter key
                        End If
                        Call UpdateSelectorLabels(TreeCode, Title, Key)        'updates selector's current label relative to given tree code with given text and key
                    End If
                Next
            End If
            mFileFilter = lblDataAccess.Tag & lblWhen.Tag & lblWho.Tag & lblWhat.Tag & lblWhere.Tag & lblWhat2.Tag & lblWho2.Tag     'sets whole filter
            Call UpdateBookmarkListHeader()                                    'if bookmark list is visible, updates its header with current selector position and shows it if new

            'Debug.Print("> DisplaySelector/end : mFileFilter = " & mFileFilter)   'TEST/DEBUG
        End If
    End Sub

    Private Sub UpdatesInDocPerPageMenu(ByVal DocNb As Integer, Optional ByVal InitDisplay As Boolean = True)
        'Updates number of lvIn documents per page in menu and related selection menu items
        mInDocPerPage = DocNb                                        'updates number of documents per page
        tsmiLvInDocPerPage.Text = DocNb & " " & mMessage(164)        'XXX documents per page

        tsmiLvInDpp10.Checked = False                                'unchecks all menu items of document per page selection
        tsmiLvInDpp25.Checked = False
        tsmiLvInDpp50.Checked = False
        tsmiLvInDpp100.Checked = False
        tsmiLvInDpp250.Checked = False
        tsmiLvInDpp500.Checked = False
        tsmiLvInDpp1000.Checked = False
        Select Case DocNb                                            'checks given selection
            Case 10
                tsmiLvInDpp10.Checked = True
            Case 25
                tsmiLvInDpp25.Checked = True
            Case 50
                tsmiLvInDpp50.Checked = True
            Case 100
                tsmiLvInDpp100.Checked = True
            Case 250
                tsmiLvInDpp250.Checked = True
            Case 500
                tsmiLvInDpp500.Checked = True
            Case 1000
                tsmiLvInDpp1000.Checked = True
        End Select

        If InitDisplay Then
            Call InitDisplaySelectedFiles()                          'initializes display of all selected files, starting at first page
        End If
    End Sub

    Private Sub InitContextMenus()
        'Selector menu :
        tsmiSelectorTreeView.Text = mMessage(32)           'Tree view panel visible
        tsmiSelectorViewer.Text = mMessage(33)             'Viewer panel visible
        tsmiSelectorImport.Text = mMessage(34)             'Import panel visible
        tsmiSelectorTreeLock.Text = mMessage(181)          'Tree view locked
        tsmiSelectorNightMode.Text = mMessage(88)          'Night mode

        tsmiSelectorChangeDb.Text = mMessage(200)          'Change database
        tsmiSelectorChangeDbImage.Text = mMessage(199)     'Change database image
        tsmiSelectorExportDBwin.Text = mMessage(100)       'Export database to Windows
        tsmiSelectorImportDBwin.Text = mMessage(173)       'Import database from Windows...
        tsmiSelectorCheckDb.Text = mMessage(186)           'Check database
        tsmiSelectorAutoImport.Text = mMessage(116)        'ZAAN cubes automatic import
        tsmiSelectorCloseDb.Text = mMessage(222)           'Close database

        'Tree menu :
        tsmiTrvWAdd.Text = mMessage(20)                    'Add
        tsmiTrvWAddExternal.Text = mMessage(36)            'Add as external folder
        tsmiTrvWRefresh.Text = mMessage(66)                'Refresh
        tsmiTrvWRename.Text = mMessage(27)                 'Rename
        tsmiTrvWDelete.Text = mMessage(21)                 'Delete
        tsmiTrvWImport.Text = mMessage(169)                'Import tree view...
        tsmiTrvWExport.Text = mMessage(171)                'Export tree view...

        'LvIn menu :
        Call UpdatesInDocPerPageMenu(mInDocPerPage, False)      'updates tsmiLvInDocPerPage.Text with "XXX documents per page" and related selection menu items
        tsmiLvInImageMode.Text = mMessage(87)              'Image mode
        tsmiLvInRefresh.Text = mMessage(66)                'Refresh
        tsmiLvInSelAll.Text = mMessage(26)                 'Select all
        tsmiLvInUndoMove.Text = mMessage(134)              'Undo move
        tsmiLvInCut.Text = mMessage(76)                    'Cut
        tsmiLvInCopy.Text = mMessage(77)                   'Copy
        tsmiLvInPaste.Text = mMessage(78)                  'Paste
        tsmiLvInCopyPath.Text = mMessage(184)              'Copy location
        tsmiLvInOpenFolder.Text = mMessage(185)            'Open folder
        tsmiLvInAutoRename.Text = mMessage(93)             'Rename automatically
        tsmiLvInRename.Text = mMessage(27)                 'Rename
        tsmiLvInDelete.Text = mMessage(21)                 'Delete
        tsmiLvInNew.Text = mMessage(12)                    'New document

        tsmiLvInCopyToZC.Text = mMessage(25)               'Copy to ZAAN\copy
        tsmiLvInMoveOut.Text = mMessage(23)                'Move out documents
        tsmiLvInExport.Text = mMessage(64)                 'Export ZAAN cube
        tsmiLvInExpNameTable.Text = mMessage(195)          'Export name table
        tsmiLvInExpRefTable.Text = mMessage(168)           'Export reference table

        'LvInWhos menu :
        tsmiLvInWhosMultiple.Text = mMessage(142)          'Multiple references
        tsmiLvInWhosDelete.Text = mMessage(143)            'Delete reference :

        'LvOut menu :
        tsmiLvOutChangeFolder.Text = mMessage(71)          'Change folder
        tsmiLvOutFoldersVisible.Text = mMessage(124)       'Folders tree visible
        tsmiLvOutRefresh.Text = mMessage(66)               'Refresh
        tsmiLvOutSelAll.Text = mMessage(26)                'Select all
        tsmiLvOutUndoMove.Text = mMessage(134)             'Undo move
        tsmiLvOutCut.Text = mMessage(76)                   'Cut
        tsmiLvOutCopy.Text = mMessage(77)                  'Copy
        tsmiLvOutPaste.Text = mMessage(78)                 'Paste
        tsmiLvOutRename.Text = mMessage(27)                'Rename
        tsmiLvOutDelete.Text = mMessage(21)                'Delete
        tsmiLvOutMoveIn.Text = mMessage(22)                'Move in documents
        tsmiLvOutAutoFile.Text = mMessage(125)             'File automatically
        tsmiLvOutImport.Text = mMessage(196)               'Import ZAAN cubes

        'LvTemp menu :
        tsmiLvTempRefresh.Text = mMessage(66)              'Refresh
        tsmiLvTempSelAll.Text = mMessage(26)               'Select all
        tsmiLvTempUndoMove.Text = mMessage(134)            'Undo move
        tsmiLvTempCut.Text = mMessage(76)                  'Cut
        tsmiLvTempCopy.Text = mMessage(77)                 'Copy
        tsmiLvTempPaste.Text = mMessage(78)                'Paste
        tsmiLvTempResizePicture.Text = mMessage(97)        'Resize pictures (.jpg)
        tsmiLvTempRename.Text = mMessage(27)               'Rename
        tsmiLvTempDelete.Text = mMessage(21)               'Delete
        tsmiLvTempSend.Text = mMessage(28)                 'New mail
        tsmiLvTempImport.Text = mMessage(196)              'Import ZAAN cubes

        'Bookmark menu :
        tsmiSelTabsAdd.Text = mMessage(89)                 'Add this cube to favorites
        tsmiSelTabsRefresh.Text = mMessage(166)            'Refresh this favorite
        tsmiSelTabsDelete.Text = mMessage(90)              'Delete this favorite
        tsmiSelTabsDefault.Text = mMessage(221)            'Default cube

        'TrvInput menu :
        tsmiTrvInputChangeRoot.Text = mMessage(127)        'Change root folder
        tsmiTrvInputRefresh.Text = mMessage(66)            'Refresh
        tsmiTrvInputDelete.Text = mMessage(21)             'Delete

        'PctZoom menu :
        tsmiPctZoomSlideShow.Text = mMessage(135)          'Slide show
        tsmiPctZoomInterval.Text = mMessage(136)           'Interval (sec.) :
    End Sub

    Private Sub InitToolTips()
        'tlTip.SetToolTip(pctZaanLogo, mMessage(200))       ' Change database

        tlTip.SetToolTip(lblAboutZaan, mMessage(132))      ' About ZAAN / Help
        tlTip.SetToolTip(btnPrev, mMessage(18))            ' Previous selection
        tlTip.SetToolTip(btnNext, mMessage(19))            ' Next selection

        tlTip.SetToolTip(lblSearchFolder, mMessage(10))    ' Search folder
        tlTip.SetToolTip(lblSearchDoc, mMessage(11))       ' Search document
        tlTip.SetToolTip(lblSelectorReset, mMessage(35))   ' Reset selector

        tlTip.SetToolTip(lblBookmark, mMessage(115))       ' Select a favorite cube
        tlTip.SetToolTip(btnPanelCube, mMessage(86))       ' Show/hide favorite cubes

        tlTip.SetToolTip(btnPanelTree, mMessage(85))       ' Show/hide tree view
        tlTip.SetToolTip(btnPanelView, mMessage(209))      ' Show/hide viewer panel
        tlTip.SetToolTip(btnPanelImport, mMessage(210))    ' Show/hide import panel

        tlTip.SetToolTip(lblDataAccess, mMessage(9))       ' Select in tree view
        tlTip.SetToolTip(lblWhen, mMessage(9))             ' Select in tree view
        tlTip.SetToolTip(lblWho, mMessage(9))              ' Select in tree view
        tlTip.SetToolTip(lblWhat, mMessage(9))             ' Select in tree view
        tlTip.SetToolTip(lblWhere, mMessage(9))            ' Select in tree view
        tlTip.SetToolTip(lblWho2, mMessage(9))             ' Select in tree view
        tlTip.SetToolTip(lblWhat2, mMessage(9))            ' Select in tree view

        'tlTip.SetToolTip(lvSelector, mMessage(163))        ' Extend selection

        tlTip.SetToolTip(btnDataAccessRoot, mMessage(178)) ' Cancel selection
        tlTip.SetToolTip(btnWhenRoot, mMessage(178))       ' Cancel selection
        tlTip.SetToolTip(btnWhoRoot, mMessage(178))        ' Cancel selection
        tlTip.SetToolTip(btnWhatRoot, mMessage(178))       ' Cancel selection
        tlTip.SetToolTip(btnWhereRoot, mMessage(178))      ' Cancel selection
        tlTip.SetToolTip(btnWhat2Root, mMessage(178))      ' Cancel selection
        tlTip.SetToolTip(btnWho2Root, mMessage(178))       ' Cancel selection

        tlTip.SetToolTip(btnDataAccess, mMessage(163))     ' Extend selection
        tlTip.SetToolTip(btnWhen, mMessage(163))           ' Extend selection
        tlTip.SetToolTip(btnWho, mMessage(163))            ' Extend selection
        tlTip.SetToolTip(btnWhat, mMessage(163))           ' Extend selection
        tlTip.SetToolTip(btnWhere, mMessage(163))          ' Extend selection
        tlTip.SetToolTip(btnWhat2, mMessage(163))          ' Extend selection
        tlTip.SetToolTip(btnWho2, mMessage(163))           ' Extend selection

        'tlTip.SetToolTip(btnLeftPanel, mMessage(24))       ' Open/close tree view panel
        'tlTip.SetToolTip(btnRightPanel, mMessage(31))      ' Open/close viewer panel
        'tlTip.SetToolTip(btnBottomPanel, mMessage(80))     ' Open/close import panel

        tlTip.SetToolTip(btnToday, mMessage(59))            ' Select today's date

        tlTip.SetToolTip(lblSelPagePrev, mMessage(205))       ' Previous page
        tlTip.SetToolTip(lblSelPageNext, mMessage(206))       ' Next page

        tlTip.SetToolTip(pctZoom, mMessage(208))           ' Starts/stops slide show
    End Sub

    Private Sub InitlvInOrder(ByVal ColIndex As Integer, ByVal AscendingOrder As Boolean)
        'Initializes mlvInOrder() table to given index column in given order (True for ASC, False for DESC)
        Dim i As Integer

        For i = 0 To mlvInOrder.Length - 1
            mlvInOrder(i) = 0                              'next columns flags set to no order
        Next
        If ColIndex >= 0 And ColIndex < mlvInOrder.Length Then
            If AscendingOrder Then
                'Debug.Print("InitlvInOrder : " & ColIndex & " ASC")   'TEST/DEBUG
                mlvInOrder(ColIndex) = 1                   'ColIndex column flag set in ASCENDING order
            Else
                'Debug.Print("InitlvInOrder : " & ColIndex & " DESC")   'TEST/DEBUG
                mlvInOrder(ColIndex) = -1                  'ColIndex column flag set in DESCENDING order
            End If
        End If
    End Sub

    Private Sub InitLvInColWidth(ByRef ColWiths As String)
        'Initializes lvIn columns width using given widths string
        Dim ColWidthTab() As String
        Dim i As Integer

        If ColWiths <> "" Then                                       'case of existing column display width list (providing col. width)
            ColWidthTab = Split(ColWiths, " ")
            If ColWidthTab.Count > 1 Then
                For i = 0 To ColWidthTab.Count - 2
                    If i < lvIn.Columns.Count Then
                        If (mLicTypeCode < 30) And ((i = 1) Or (i = 6) Or (i = 7)) Then  'case of "Basic" and "First" licenses
                            lvIn.Columns(i).Width = 0                                    'hides Access control/Status/Action for columns
                        Else
                            lvIn.Columns(i).Width = ColWidthTab(i)
                        End If
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub InitListColWidth(ByRef CurList As ListView, ByRef ColWiths As String)
        'Initializes columns width of given list view, using given widths string
        Dim ColWidthTab() As String
        Dim i As Integer

        If ColWiths <> "" Then                             'case of existing column display width list (providing col. width)
            ColWidthTab = Split(ColWiths, " ")
            If ColWidthTab.Count > 1 Then
                For i = 0 To ColWidthTab.Count - 2
                    If i < CurList.Columns.Count Then
                        CurList.Columns(i).Width = ColWidthTab(i)
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub InitSelectorColWidth(ByRef ColWiths As String)
        'Initializes selector columns width using given widths string
        Dim ColWidthTab() As String
        Dim i, w, w1, w2 As Integer
        Dim wr As Integer = 20

        If ColWiths <> "" Then                             'case of existing column display width list (providing col. width)
            ColWidthTab = Split(ColWiths, " ")
            If ColWidthTab.Count > 2 Then
                pnlSelectorSearch.Width = ColWidthTab(0)

                For i = 1 To ColWidthTab.Count - 2
                    If i < lvSelector.Columns.Count Then
                        w = ColWidthTab(i)
                        lvSelector.Columns(i - 1).Width = w
                        If w > wr Then
                            w1 = wr
                            w2 = w - wr
                        Else
                            w1 = w
                            w2 = 0
                        End If
                        If (mLicTypeCode < 30) And ((i = 1) Or (i = 6) Or (i = 7)) Then  'case of "Basic" and "First" licenses
                            w = 0
                            w1 = 0
                            w2 = 0
                        End If

                        Select Case i
                            Case 1
                                lblDataAccess.Width = w
                                btnDataAccessRoot.Width = w1
                                btnDataAccess.Width = w2
                                btnDataAccessBlank.Width = w
                            Case 2
                                lblWhen.Width = w
                                btnWhenRoot.Width = w1
                                btnWhen.Width = w2
                                btnToday.Width = w
                            Case 3
                                lblWho.Width = w
                                btnWhoRoot.Width = w1
                                btnWho.Width = w2
                            Case 4
                                lblWhat.Width = w
                                btnWhatRoot.Width = w1
                                btnWhat.Width = w2
                            Case 5
                                lblWhere.Width = w
                                btnWhereRoot.Width = w1
                                btnWhere.Width = w2
                            Case 6
                                lblWhat2.Width = w
                                btnWhat2Root.Width = w1
                                btnWhat2.Width = w2
                            Case 7
                                lblWho2.Width = w
                                btnWho2Root.Width = w1
                                btnWho2.Width = w2
                        End Select
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub InitListsColOrderWidth()
        'Initializes LvIn columns order/width and other Lists column widths
        Dim ColOrderTab() As String
        Dim i As Integer

        If mColDisplayIndexes <> "" Then                             'case of existing column display index list (providing col. order)
            ColOrderTab = Split(mColDisplayIndexes, " ")
            If ColOrderTab.Count > 1 Then
                For i = 0 To ColOrderTab.Count - 2
                    If i < lvIn.Columns.Count Then
                        lvIn.Columns(i).DisplayIndex = ColOrderTab(i)
                    End If
                Next
            End If
        End If

        Call InitLvInColWidth(mColDisplayWidths)                     'initializes lvIn (and lvSelector) columns width using mColDisplayWidths
        If mIsImageView Then
            Call InitSelectorColWidth(mColSelectorWidths)            'initializes selector columns width using given widths string
        End If

        Call InitListColWidth(lvOut, mColImportWidths)               'initializes lvOut columns width using mColImportWidths
        Call InitListColWidth(lvTemp, mColCopyWidths)                'initializes lvTemp columns width using mColCopyWidths
    End Sub

    Private Sub InitLvBookmark()
        'Initializes bookmark list

        'Debug.Print("InitLvBookmark...")    'TEST/DEBUG
        lvBookmark.Items.Clear()
        lvBookmark.Columns.Clear()
        lvBookmark.Columns.Add("add", "", 500, Forms.HorizontalAlignment.Left, "add")          'sets column header with "add" icon
    End Sub

    Private Sub InitlvSelector()
        'Initializes Selector lists
        Dim w As Integer = 130   '77

        lvSelector.Items.Clear()
        lvSelector.Columns.Clear()

        If mLicTypeCode >= 30 Then                         'case of ZAAN-Pro license (access control)
            lvSelector.Columns.Add("_u_", mMessage(7), w, Forms.HorizontalAlignment.Left, "_uh" & mImageStyle)     'Access control
        Else
            lvSelector.Columns.Add("_u_", mMessage(7), 0, Forms.HorizontalAlignment.Left, "_uh" & mImageStyle)     'Access control hidden in ZAAN-Basic and ZAAN-First mode
        End If

        lvSelector.Columns.Add("_t_", mMessage(1), w, Forms.HorizontalAlignment.Left, "_th" & mImageStyle)     'When
        lvSelector.Columns.Add("_o_", mMessage(2), w, Forms.HorizontalAlignment.Left, "_oh" & mImageStyle)     'Who
        lvSelector.Columns.Add("_a_", mMessage(3), w, Forms.HorizontalAlignment.Left, "_ah" & mImageStyle)     'What
        lvSelector.Columns.Add("_e_", mMessage(4), w, Forms.HorizontalAlignment.Left, "_eh" & mImageStyle)     'Where

        If mLicTypeCode >= 30 Then                         'case of ZAAN-Pro license (6 dimensions)
            lvSelector.Columns.Add("_b_", mMessage(5), w, Forms.HorizontalAlignment.Left, "_bh" & mImageStyle)     'What else
            lvSelector.Columns.Add("_c_", mMessage(6), w, Forms.HorizontalAlignment.Left, "_ch" & mImageStyle)     'Other
        Else                                               'case of ZAAN-First and ZAAN-Basic licenses (4 dimensions)
            lvSelector.Columns.Add("_b_", mMessage(5), 0, Forms.HorizontalAlignment.Left, "_bh" & mImageStyle)     'What else (hidden in ZAAN-Basic mode)
            lvSelector.Columns.Add("_c_", mMessage(6), 0, Forms.HorizontalAlignment.Left, "_ch" & mImageStyle)     'Other (hidden in ZAAN-Basic mode)
        End If

        btnDataAccessRoot.ImageKey = "_uh" & mImageStyle
        btnWhenRoot.ImageKey = "_th" & mImageStyle
        btnWhoRoot.ImageKey = "_oh" & mImageStyle
        btnWhatRoot.ImageKey = "_ah" & mImageStyle
        btnWhereRoot.ImageKey = "_eh" & mImageStyle
        btnWhat2Root.ImageKey = "_bh" & mImageStyle
        btnWho2Root.ImageKey = "_ch" & mImageStyle

        btnNext.ImageKey = "next_f" & mImageStyle
        btnPrev.ImageKey = "prev_f" & mImageStyle
    End Sub

    Private Sub InitlvSelectorList(ByRef SelList As ListView)
        'Initializes given selector list

        SelList.Items.Clear()
        SelList.Columns.Clear()
        SelList.Columns.Add("", 60, Forms.HorizontalAlignment.Left)
    End Sub

    Private Sub InitlvIn()
        'Initializes selected files display with related columns, using given order if any

        'Debug.Print("InitlvIn...")   'TEST/DEBUG
        Call InitlvSelector()                              'initializes Selector list

        lvIn.Items.Clear()
        lvIn.Columns.Clear()

        lvIn.Columns.Add(mMessage(44), 300, Forms.HorizontalAlignment.Left)                                   'Document (name)

        If mLicTypeCode >= 30 Then                         'case of ZAAN-Pro license (access control)
            lvIn.Columns.Add("_u_", mMessage(7), 100, Forms.HorizontalAlignment.Left, "_u_" & mImageStyle)    'Access control
        Else
            lvIn.Columns.Add("_u_", mMessage(7), 0, Forms.HorizontalAlignment.Left, "_u_" & mImageStyle)      'Access control hidden in ZAAN-Basic and ZAAN-First mode
        End If

        lvIn.Columns.Add("_t_", mMessage(1), 100, Forms.HorizontalAlignment.Left, "_t_" & mImageStyle)        'When
        lvIn.Columns.Add("_o_", mMessage(2), 135, Forms.HorizontalAlignment.Left, "_o_" & mImageStyle)        'Who
        lvIn.Columns.Add("_a_", mMessage(3), 135, Forms.HorizontalAlignment.Left, "_a_" & mImageStyle)        'What
        lvIn.Columns.Add("_e_", mMessage(4), 125, Forms.HorizontalAlignment.Left, "_e_" & mImageStyle)        'Where

        If mLicTypeCode >= 30 Then                         'case of ZAAN-Pro license (6 dimensions)
            lvIn.Columns.Add("_b_", mMessage(5), 100, Forms.HorizontalAlignment.Left, "_b_" & mImageStyle)    'What else
            lvIn.Columns.Add("_c_", mMessage(6), 100, Forms.HorizontalAlignment.Left, "_c_" & mImageStyle)    'Other
        Else                                               'case of ZAAN-First and ZAAN-Basic licenses (4 dimensions)
            lvIn.Columns.Add("_b_", mMessage(5), 0, Forms.HorizontalAlignment.Left, "_b_" & mImageStyle)      'What else (hidden in ZAAN-Basic mode)
            lvIn.Columns.Add("_c_", mMessage(6), 0, Forms.HorizontalAlignment.Left, "_c_" & mImageStyle)      'Other (hidden in ZAAN-Basic mode)
        End If

        lvIn.Columns.Add(mMessage(50), 45, Forms.HorizontalAlignment.Left)                     'Type
        lvIn.Columns.Add(mMessage(45), 75, Forms.HorizontalAlignment.Right)                    'Size (Kb)
        lvIn.Columns.Add(mMessage(46), 140, Forms.HorizontalAlignment.Right)                   'Modification date
    End Sub

    Private Sub ShowHideExtraInCol()
        'Depending on ZAAN-Pro/-Basic mode, shows or hides "What else" and "Other" columns of "in" list
        Dim w As Integer = 100

        If mLicTypeCode >= 30 Then                         'case of "Pro" license (Access control + 6 dimensions)
            lvIn.Columns(1).Width = w                      'Access control
            lvIn.Columns(6).Width = w                      'Status (What2)
            lvIn.Columns(7).Width = w                      'Action (Who2)
        Else
            lvIn.Columns(1).Width = 0                      'Access control
            lvIn.Columns(6).Width = 0                      'Status (What2)
            lvIn.Columns(7).Width = 0                      'Action (Who2)
        End If
    End Sub

    Private Sub InitlvOut()
        Dim i As Integer

        lvOut.Items.Clear()
        lvOut.Columns.Clear()
        lvOut.Columns.Add(mMessage(82), 300, Forms.HorizontalAlignment.Left)                   'documents
        lvOut.Columns.Add(mMessage(50), 45, Forms.HorizontalAlignment.Left)                    'Type
        lvOut.Columns.Add(mMessage(45), 75, Forms.HorizontalAlignment.Right)                   'Size (Kb)
        lvOut.Columns.Add(mMessage(46), 140, Forms.HorizontalAlignment.Right)                  'Modification date

        'lvOut.View = View.Details
        mlvOutOrder(0) = 1                                 '1st column flag set in ascending order
        For i = 1 To mlvOutOrder.Length - 1
            mlvOutOrder(i) = 0                             'next columns flags set to no order
        Next
    End Sub

    Private Sub InitlvTemp()
        Dim i As Integer

        lvTemp.Items.Clear()
        lvTemp.Columns.Clear()
        lvTemp.Columns.Add(mMessage(82), 300, Forms.HorizontalAlignment.Left)                  'documents
        lvTemp.Columns.Add(mMessage(50), 45, Forms.HorizontalAlignment.Left)                   'Type
        lvTemp.Columns.Add(mMessage(45), 75, Forms.HorizontalAlignment.Right)                  'Size (Kb)
        lvTemp.Columns.Add(mMessage(46), 140, Forms.HorizontalAlignment.Right)                 'Modification Date

        'lvTemp.View = View.Details
        mlvTempOrder(0) = 1                                '1st column flag set in ascending order
        For i = 1 To mlvTempOrder.Length - 1
            mlvTempOrder(i) = 0                            'next columns flags set to no order
        Next
    End Sub

    Private Sub RenameFileImage(ByVal FilePath As String, ByVal OldFileName As String, ByVal NewFileName As String, ByVal FileType As String)
        'Renames thumbnail image file corresponding to given (old) file that may have been generated by ZAAN
        Dim dirSel As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim ThumbFileName, ThumbPrefix As String
        Dim p As Integer

        'Debug.Print("=> RenameFileImage:  " & FilePath & " : " & OldFileName & "  => " & NewFileName)    'TEST/DEBUG
        ThumbFileName = ""
        dirSel = My.Computer.FileSystem.GetDirectoryInfo(FilePath)
        For Each fi In dirSel.GetFiles("zzi*" & OldFileName & FileType)   'searches for corresponding thumbnail image file...
            ThumbFileName = fi.Name
            Exit For
        Next
        'Debug.Print("RenameFileImage :  ThumbFileName = " & ThumbFileName)    'TEST/DEBUG
        If ThumbFileName <> "" Then                                       'renames found thumbnail image file
            ThumbPrefix = ""
            p = InStr(ThumbFileName, OldFileName)                    'WARNING : do not use FileType with caps varies !!
            If p > 0 Then
                ThumbPrefix = Microsoft.VisualBasic.Left(ThumbFileName, p - 1)
            End If
            If ThumbPrefix <> "" Then
                Try
                    If UCase(NewFileName) = UCase(OldFileName) Then
                        My.Computer.FileSystem.RenameFile(FilePath & "\" & ThumbPrefix & OldFileName & FileType, ThumbPrefix & "ZAAN_TEMP" & FileType)
                        My.Computer.FileSystem.RenameFile(FilePath & "\" & ThumbPrefix & "ZAAN_TEMP" & FileType, ThumbPrefix & NewFileName & FileType)
                    Else
                        My.Computer.FileSystem.RenameFile(FilePath & "\" & ThumbPrefix & OldFileName & FileType, ThumbPrefix & NewFileName & FileType)
                    End If
                    'Debug.Print("=> FileImage renamed :  " & FilePath & " : " & ThumbPrefix & OldFileName & "  => " & ThumbPrefix & NewFileName)    'TEST/DEBUG
                Catch ex As Exception
                    Debug.Print("RenameFileImage error : " & ex.Message)    'TEST/DEBUG
                End Try
            End If
        End If
    End Sub

    Private Sub DeleteFileImage(ByVal DirPathName As String, ByVal FileName As String)
        'Deletes thumbnail image file corresponding to given file that may have been generated by ZAAN
        Dim dirSel As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim ThumbFileName As String

        'Debug.Print("DeleteFileImage:  DirPathName = " & DirPathName & "  FileName = " & FileName)    'TEST/DEBUG
        dirSel = My.Computer.FileSystem.GetDirectoryInfo(DirPathName)
        ThumbFileName = ""
        For Each fi In dirSel.GetFiles("zzi*" & FileName)                 'searches for corresponding thumbnail image file...
            ThumbFileName = fi.Name
            Exit For
        Next
        If ThumbFileName <> "" Then                                       'deletes found thumbnail image file
            Try
                'My.Computer.FileSystem.DeleteFile(DirPathName & "\" & ThumbFileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                My.Computer.FileSystem.DeleteFile(DirPathName & "\" & ThumbFileName)
            Catch ex As Exception
                Debug.Print("DeleteImage error : " & ex.Message)    'TEST/DEBUG
            End Try
        End If
    End Sub

    Private Function GetLvInItemImageAndSize(ByVal dirSel As System.IO.DirectoryInfo, ByVal DirPathName As String, ByVal FileName As String, ByVal FileTime As DateTime, ByVal FileType As String) As String
        'Loads pctLvIn with given file related thumbnail image to be eventually generated
        'if doesn't exist (which date is not older than source file date) and returns source image size.
        Dim SceImageSize, ThumbFileName As String
        Dim g As Graphics
        Dim destRect, srcRect As Rectangle
        Dim fi As System.IO.FileInfo

        If dirSel Is Nothing Then
            dirSel = My.Computer.FileSystem.GetDirectoryInfo(DirPathName)
        End If
        SceImageSize = ""
        ThumbFileName = ""
        For Each fi In dirSel.GetFiles("zzi*" & FileName)                 'searches for corresponding thumbnail image file...
            If fi.LastWriteTime >= FileTime Then                          '...that is not older than source file date/time
                ThumbFileName = fi.Name
                SceImageSize = Mid(ThumbFileName, 4, Len(ThumbFileName) - Len(FileName) - 4)
                Exit For
            End If
        Next

        If ThumbFileName <> "" Then                                       'thumbnail image file exists
            pctLvIn.Load(DirPathName & ThumbFileName)
        Else                                                              'thumbnail image file doesn't exist => creates it !
            Dim image As New Bitmap(DirPathName & FileName)               'loads source image (full size)
            Dim bm As New Bitmap(pctLvIn.Width, pctLvIn.Height)

            SceImageSize = image.Width & "x" & image.Height
            ThumbFileName = "zzi" & SceImageSize & "_" & FileName
            g = Graphics.FromImage(pctLvIn.Image)
            srcRect.X = 0
            srcRect.Y = 0
            srcRect.Width = image.Width
            srcRect.Height = image.Height
            If image.Height > image.Width Then
                destRect.Width = pctLvIn.Height * image.Width / image.Height
                destRect.Height = pctLvIn.Height
                destRect.X = (pctLvIn.Width - destRect.Width) / 2
                destRect.Y = 0
            Else
                destRect.Width = pctLvIn.Width
                destRect.Height = pctLvIn.Width * image.Height / image.Width
                destRect.X = 0
                destRect.Y = (pctLvIn.Height - destRect.Height) / 2
            End If
            g.DrawImage(image, destRect, srcRect, GraphicsUnit.Pixel)     'draw image at given position and size
            pctLvIn.Image.Save(DirPathName & ThumbFileName)               'saves generated thumbnail image as a zzi file

            image.Dispose()
            bm.Dispose()
        End If

        Dim fiNew As FileInfo = New FileInfo(DirPathName & ThumbFileName)
        'Debug.Print("fiNew = " & fiNew.FullName)   'TEST/DEBUG
        If (fiNew.Attributes And FileAttributes.Hidden) <> FileAttributes.Hidden Then         'thumnail image file (existing or new) is not hidden
            Try
                fiNew.Attributes = fiNew.Attributes Or FileAttributes.Hidden                  'hides thumbnail image file in Windows file system
            Catch ex As Exception
                Debug.Print(ex.Message)                                   'TEST/DEBUG error
            End Try
        End If

        GetLvInItemImageAndSize = SceImageSize                            'returns source image size string (Width x Height)
    End Function

    Private Function GetFileTypeAndImage(ByVal FileName As String) As String
        'Returns file type and get related system icon if not already loaded in imgFileTypes list of images
        Dim FileType As String
        Dim FileTypeImage As Image

        If FileName = "folder" Then
            FileType = "folder"
            FileName = mMyZaanImportPath                             'NEW SINCE 2012-03-09 : use an existing path name in GetDefaultIcon()
        Else
            FileType = LCase(System.IO.Path.GetExtension(FileName))  'get file extension/type
        End If

        If imgFileTypes.Images.IndexOfKey(FileType) = -1 Then        'file type is new
            'Debug.Print("GetFileTypeAndImage: new file type = " & FileType)     'TEST/DEBUG

            'FileTypeIcon = clsIcon.GetDefaultIcon(FileName)          'NEW SINCE 2012-03-09 get system application (small) icon related to FILE NAME (instead of file type)
            FileTypeImage = GetShellIconAsImage(FileName)          'NEW SINCE 2012-03-09 get system application (small) icon related to FILE NAME (instead of file type)
            If FileTypeImage Is Nothing Then
                imgFileTypes.Images.Add(FileType, imgFileTypes.Images(".zqm"))  'loads .zqm icon associated to unknown application
            Else
                imgFileTypes.Images.Add(FileType, FileTypeImage)      'loads system icon of related application
            End If
        End If
        GetFileTypeAndImage = FileType
    End Function

    Private Sub SetLvInDisplayMode(ByVal SetImageView As Boolean)
        'Sets "in" list display mode depending on given SetImageView, related colors and local menu display
        mIsImageView = SetImageView                                  'stores new image view/detail view mode
        Call SetDetailsImagesButton()                                'sets details/images button image depending on current lvIn display and ZAAN image style
        If mIsImageView Then                                         '=> sets images view mode
            lvIn.View = View.LargeIcon
            lvIn.Font = SelectionFontImage
            lvIn.BackColor = mBackColorPicture
            lvIn.ForeColor = Color.SteelBlue
        Else                                                         '=> sets details view mode
            lvIn.View = View.Details
            lvIn.Font = SelectionFontDetail
            lvIn.BackColor = mBackColorContent
            If (mBackColorContent.GetBrightness * 10 < 5) Then       'case of a dark content background
                lvIn.ForeColor = Color.Silver
            Else
                lvIn.ForeColor = Color.Black
            End If
        End If
    End Sub

    Private Sub SetTreeViewLockMode()
        'Updates tree view lock control state in Selector's menu and related tree edition capabilities

        If mTreeViewLocked Then
            tsmiSelectorTreeLock.CheckState = CheckState.Checked     'updates tree view lock control state to locked (in Selector's menu)
            'trvW.ContextMenuStrip = Nothing                          '=> disables tree local menu
            trvW.LabelEdit = False                                   '=> disables tree node edition
            trvW.AllowDrop = True                                    '=> enables data dragging on tree
        Else
            tsmiSelectorTreeLock.CheckState = CheckState.Unchecked   'updates tree view lock control state to unlocked (in Selector's menu)
            'trvW.ContextMenuStrip = cmsTrvW                          '=> enables tree local menu
            trvW.LabelEdit = True                                    '=> ensables tree node edition
            trvW.AllowDrop = True                                    '=> enables data dragging on tree
        End If
    End Sub

    Private Sub ToggleViewerPanel()
        'Toggles viewer panel display

        SplitContainer3.Panel2Collapsed = Not SplitContainer3.Panel2Collapsed
        Call SetRightPanelButton()                         'sets right panel button image depending on splitter and background color
        If SplitContainer3.Panel2Collapsed Then            'case of invisible zoom panel
            lblDocName.Text = ""                           'cancels current browser display, if any
            mSlideShowPathFileName = ""                    'sets file name to no image/slide by default
            If tmrVideoProgress.Enabled Then               'case of a playing video
                Call StopVideoPlayer()                     'stops VLC video player (which activity should be indicated by tmrVideoProgress activity)
            End If
        End If
        lvIn.Focus()
    End Sub

    Private Sub ToggleImportPanel()
        'Toggles import panel display

        SplitContainer1.Panel2Collapsed = Not SplitContainer1.Panel2Collapsed
        Call SetBottomPanelButton()                                  'sets bottom panel button display depending on splitter
        Call SetSelectionPanelDisplay()                              'sets selection panel display with cube tabs if import panel visible, without cube tabs else
        lvIn.Focus()
    End Sub

    Private Function IsNewPiledInDir(ByVal InDirPathName As String, Optional ByVal AddToPile As Boolean = True) As Boolean
        'Returns true if given in directory is new in lstInDir pile of current selection
        Dim i As Integer
        Dim newItem As Boolean = True

        'Debug.Print("IsNewPiledInDir : " & InDirPathName & " ?")    'TEST/DEBUG
        For i = 0 To lstInDir.Items.Count - 1                        'scans existing directories in list/pile
            'Debug.Print(" => IsNewPID :  item " & i & " = " & lstInDir.Items(i))    'TEST/DEBUG
            If lstInDir.Items(i) = InDirPathName Then
                newItem = False                                      'given directory exists in list/pile
                Exit For
            End If
        Next
        If newItem And AddToPile Then
            'Debug.Print(" => new InDPN = " & InDirPathName)    'TEST/DEBUG
            lstInDir.Items.Add(InDirPathName)                        'adds directory in list/pile
        End If
        'Debug.Print("IsNewPiledInDir : " & InDirPathName & "  AddToPile = " & AddToPile & "  newItem = " & newItem)    'TEST/DEBUG
        IsNewPiledInDir = newItem
    End Function

    Private Sub UpdateDimPathTabDataPath(ByVal DirPathName As String, ByRef DimPathTab() As String)
        'Updates dimension path table with given data path name splitted by dimension (v4 database format)
        Dim SubPath, DirName As String
        Dim i, j, p As Integer

        'Debug.Print("UpdateDimPathTabDataPath :  DirPathName = " & DirPathName)    'TEST/DEBUG
        For i = 1 To DimPathTab.Length - 1                           'clears all 6 dimension references in current lvIn item table
            DimPathTab(i) = ""
        Next
        p = 1
        i = 1
        SubPath = DirPathName

        Do While (p > 0) And (i < DimPathTab.Length)                 'extracts dimension references (that may be multiple) from path file name
            p = InStr(p, SubPath, "\z" & i & "\")                    'searches for i dimension tag
            If p > 0 Then
                SubPath = Mid(SubPath, p + 4)
                For j = i To 6
                    p = InStr(SubPath, "\z" & j & "\")               'get next tag position, if any
                    If p > 0 Then
                        DirName = Mid(SubPath, 1, p - 1)             'get dimension i sub path
                        If DimPathTab(i) = "" Then
                            DimPathTab(i) = DirName
                        Else                                         'case of multiple dimension references
                            DimPathTab(i) = DimPathTab(i) & "|" & DirName
                        End If
                        i = j
                        Exit For
                    End If
                Next
                If (p = 0) And (SubPath <> "") Then                  'case of last sub path with no tag after
                    DirName = SubPath                                'get dimension i sub path
                    If DimPathTab(i) = "" Then
                        DimPathTab(i) = DirName
                    Else                                             'case of multiple dimension references
                        DimPathTab(i) = DimPathTab(i) & "|" & DirName
                    End If
                End If
            End If
        Loop
    End Sub

    Private Sub UpdateDimRefInFileTab(ByVal DirPathName As String, ByRef itmXTab() As String, ByVal LastDirOnly As Boolean)
        'Updates lvIn dimension references table of given directory path name (v4 database format)
        Dim SubPath, DirName As String
        Dim i, j, p, q As Integer

        For i = 1 To itmXTab.Length - 1                              'clears all 6 dimension references in current lvIn item table
            itmXTab(i) = ""
        Next
        p = 1
        i = 1
        SubPath = DirPathName
        Do While (p > 0) And (i < itmXTab.Length)                    'extracts dimension references (that may be multiple) from path file name
            p = InStr(p, SubPath, "\z" & i & "\")                    'searches for i dimension tag
            If p > 0 Then
                SubPath = Mid(SubPath, p + 4)
                For j = i To 6
                    p = InStr(SubPath, "\z" & j & "\")               'get next tag position, if any
                    If p > 0 Then
                        DirName = Mid(SubPath, 1, p - 1)             'get dimension i sub path
                        q = InStrRev(DirName, "\")
                        If q > 0 Then
                            DirName = Mid(DirName, q + 1)            'get last directory name
                        End If
                        If itmXTab(i) = "" Then
                            itmXTab(i) = DirName
                        Else                                         'case of multiple dimension references
                            itmXTab(i) = itmXTab(i) & ", " & DirName
                        End If
                        i = j
                        Exit For
                    End If
                Next
                If (p = 0) And (SubPath <> "") Then                  'case of last sub path with no tag after
                    DirName = SubPath                                'get dimension i sub path
                    q = InStrRev(DirName, "\")
                    If q > 0 Then
                        DirName = Mid(DirName, q + 1)                'get last directory name
                    End If
                    If itmXTab(i) = "" Then
                        itmXTab(i) = DirName
                    Else                                             'case of multiple dimension references
                        itmXTab(i) = itmXTab(i) & ", " & DirName
                    End If
                End If
            End If
        Loop
    End Sub

    Private Sub LoadLvInItemsDimReferences()
        'Loads dimension references of all existing lvIn items (in v4 database format)
        Dim itmX As ListViewItem
        Dim itmXTab(11) As String
        Dim i As Integer

        For Each itmX In lvIn.Items                                  'scans all list items
            If IsNewPiledInDir(itmX.Tag) Then                        'eventually piles up new directory found
            End If
            Call UpdateDimRefInFileTab(itmX.Tag, itmXTab, True)      'updates lvIn dimension references table of given directory path name
            For i = 1 To 6
                itmX.SubItems(i + 1).Text = itmXTab(i)               'updates item dimension references
            Next
        Next
    End Sub

    Private Sub LoadDataFilteredPathFiles(ByRef dirSel As System.IO.DirectoryInfo, ByVal AccessGroup As String)
        'Loads in lvIn files from given filtered data directory (Windows exported v4 database format)
        Dim fi As System.IO.FileInfo
        Dim Prefix, DataPathName, FullDirPathName, DirPathName, FileName, FileNameKey, FileType As String
        Dim itmX As ListViewItem
        Dim itmXTab(11), groupName As String
        Dim i, FirstCount As Integer
        Dim isImage, isHiddenFile As Boolean

        'Debug.Print("  > LoadDataFPF : dirSel = " & dirSel.FullName & "  AccessGroup = " & AccessGroup)    'TEST/DEBUG
        'Debug.Print("  > LoadDataFPF : dirSel = " & dirSel.FullName)    'TEST/DEBUG

        FirstCount = 20                                              'number of items to be first displayed
        FileNameKey = dirSel.Name
        DataPathName = mZaanDbPath & "data" & AccessGroup
        If AccessGroup <> "" Then
            groupName = Mid(AccessGroup, 2)                          'access control group (skips "_")
        Else
            groupName = ""
        End If
        For i = 1 To 6                                               'clears all 6 dimension references in current lvIn item table (to be later updated)
            itmXTab(i) = ""
        Next

        Try
            For Each fi In dirSel.GetFiles("*.*", SearchOption.AllDirectories)
                'Debug.Print("=> " & fi.FullName)    'TEST/DEBUG
                FullDirPathName = System.IO.Path.GetDirectoryName(fi.FullName)
                DirPathName = Mid(FullDirPathName, DataPathName.Length + 2)

                FileName = fi.Name
                Prefix = Microsoft.VisualBasic.Left(FileName, 3)
                isHiddenFile = False
                If (fi.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then isHiddenFile = True
                If Microsoft.VisualBasic.Left(FileName, 1) = "~" Then isHiddenFile = True

                If FileName <> "." And FileName <> ".." And Prefix <> "zzi" And UCase(FileName) <> "THUMBS.DB" And Not isHiddenFile Then    'ignores current and encompassing directories, ZAAN and Windows XP thumbnail images and hidden files
                    'FileType = GetFileTypeAndImage(FileName)                             'returns file type and get image from system if not already loaded in imgFileTypes
                    FileType = GetFileTypeAndImage(fi.FullName)                          'returns file type and get image from system if not already loaded in imgFileTypes

                    isImage = (FileType = ".jpg") Or (FileType = ".gif") Or (FileType = ".png") Or (FileType = ".bmp")

                    If (lvIn.View = View.Details) Or isImage Then                        'displays all docs in details view and only the images in icon view
                        'Debug.Print(" > File : " & FileName)  'TEST/DEBUG
                        mInDocCount = mInDocCount + 1
                        itmX = lvIn.Items.Add(System.IO.Path.GetFileNameWithoutExtension(FileName))  'adds new item (file name) in list
                        itmX.Tag = DirPathName                                           'stores directory path name in item's tag
                        itmXTab(0) = groupName                                           'set access control (dimension 0)

                        'If IsNewPiledInDir(DirPathName) Then                             'eventually piles up new directory found
                        'End If
                        Call UpdateDimRefInFileTab("\" & DirPathName, itmXTab, True)     'updates lvIn dimension references table of given directory path name

                        itmXTab(7) = FileType                                            'file type
                        itmXTab(8) = Format(fi.Length / 1000, "### ### ###")             'file size
                        itmXTab(9) = Format(fi.LastWriteTime, "d MMM yyyy  HH:mm")       'modification date
                        itmXTab(10) = ""                                                 'available for picture dimensions

                        If lvIn.View = View.LargeIcon Then                               'displays item picture (as "large icon")
                            pctLvIn.Image = My.Resources.picture_black_90x90             'initializes image with empty picture
                            If isImage Then
                                itmXTab(10) = GetLvInItemImageAndSize(Nothing, FullDirPathName & "\", FileName, fi.LastWriteTime, FileType) & " pixels"   'get image size and loads related thumbnail image in pctLvIn
                            End If
                            imgLargeIcons.Images.Add(pctLvIn.Image)
                            itmX.ImageIndex = mInDocCount - 1
                        Else                                                             'displays item "details"
                            itmX.ImageKey = FileType
                        End If
                        For i = 0 To 9
                            itmX.SubItems.Add(itmXTab(i))                                'fills access control, who, what, when, where, what else, other, type, size and date columns
                        Next
                        itmX.ToolTipText = FileType & "  "
                        itmX.ToolTipText = itmX.ToolTipText & itmXTab(8) & " Kb" & vbCrLf & itmXTab(9) & vbCrLf & itmXTab(10) & vbCrLf
                        itmX.ToolTipText = itmX.ToolTipText & itmXTab(0) & vbCrLf & itmXTab(1) & vbCrLf & itmXTab(2) & vbCrLf & itmXTab(3) & vbCrLf & itmXTab(4) & vbCrLf & itmXTab(5) & vbCrLf & itmXTab(6)

                        If mInDocCount Mod FirstCount = 0 Then
                            lvIn.RedrawItems(mInDocPrevCount, mInDocCount - 1, False)    'forces partial display of (new) items within given index range
                            mInDocPrevCount = mInDocCount
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            Debug.Print(ex.Message)                        'TEST/DEBUG : error !
        End Try
    End Sub

    Private Function GetDirInfoDataPathFilter(ByVal DirPathName As String, ByVal FileFilter As String, ByVal AccessGroup As String) As System.IO.DirectoryInfo
        'Returns lowest data child directory corresponding to given filter series (pasted directory names preceded by "*") and loads related files in lvIn
        Dim dirInfo, dirSel, dirSelZ, dirInfoOutput As System.IO.DirectoryInfo
        Dim SelPatterns() As String = Split(FileFilter, "*")
        Dim SelPattNames(), SelPattPath As String
        Dim i, j As Integer
        Dim FilterIsEmpty As Boolean = True

        'Debug.Print("GetDirInfoDPF :  DirPathName = " & DirPathName & "  FileFilter = " & FileFilter & "  AccessGroup = " & AccessGroup)   'TEST/DEBUG
        'Debug.Print("GetDirInfoDPF :  DirPathName = " & DirPathName & "  FileFilter = " & FileFilter)   'TEST/DEBUG

        dirInfoOutput = Nothing
        dirInfo = My.Computer.FileSystem.GetDirectoryInfo(DirPathName)
        i = 1
        Do While i < SelPatterns.Length                                             'scans all pattern dimensions (after Access control)
            If SelPatterns(i) <> "" Then                                            'get 1st filter (that is not empty)
                If Mid(SelPatterns(i), 1, 2) <> "z0" Then                           'skips Access control filter (already used, if any, in DirPathName)
                    FilterIsEmpty = False
                    FileFilter = ""
                    If SelPatterns.Length > i + 1 Then
                        For j = i + 1 To SelPatterns.Length - 1
                            If SelPatterns(j) <> "" Then
                                FileFilter = FileFilter & "*" & SelPatterns(j)          'rebuilds FileFilter with remaining filters (1st one removed)
                            End If
                        Next
                    End If
                    SelPattNames = Split(SelPatterns(i), "\")                           'extracts directory names
                    SelPattPath = Mid(SelPatterns(i), Len(SelPattNames(0)) + 2)         'get directory path after z#\ path root
                    If SelPattPath <> "" Then
                        Try
                            For Each dirSelZ In dirInfo.GetDirectories(SelPattNames(0), SearchOption.AllDirectories)     'scans z# directories matching with current index
                                If My.Computer.FileSystem.DirectoryExists(dirSelZ.FullName & "\" & SelPattPath) Then
                                    For Each dirSel In dirSelZ.GetDirectories(SelPattPath, SearchOption.AllDirectories)      'scans SelPattPath sub-directories of found z# directories
                                        If FileFilter = "" Then                                     'end of recursive function call
                                            dirInfoOutput = dirSel
                                            Call LoadDataFilteredPathFiles(dirSel, AccessGroup)     'loads in lvIn files from given filtered data directory
                                        Else
                                            dirInfoOutput = GetDirInfoDataPathFilter(dirSel.FullName, FileFilter, AccessGroup)  'recursive function call until FileFilter is empty
                                        End If
                                    Next
                                End If
                            Next
                        Catch ex As Exception
                            Debug.Print(ex.Message)                                     'TEST/DEBUG : error !
                        End Try
                    End If
                    Exit Do                                                             'stops loop after 1st filter found
                End If
            End If
            i = i + 1
        Loop
        If FilterIsEmpty Then
            Call LoadDataFilteredPathFiles(dirInfo, AccessGroup)                    'loads in lvIn files from given data directory
        End If
        GetDirInfoDataPathFilter = dirInfoOutput
    End Function

    Private Sub LoadFilterFiles(ByVal AccessGroup As String)
        'Loads data files matching with given filter, in "data" directory or in given sub-directory DataAccessDirName, using v4 database format
        Dim dirInfo As System.IO.DirectoryInfo
        Dim FileFilter, FullDataPathName As String

        If mTreeKeyLength > 0 Then Exit Sub 'database v4 format don't use tree keys anymore (since 15 Dec 2010)

        FileFilter = mFileFilter
        'Debug.Print("  LoadFilterFiles START :  FileFilter = " & FileFilter & "  AccessGroup = " & AccessGroup)    'TEST/DEBUG

        If FileFilter = "BIDON" And AccessGroup = "" Then            'TEST/DEBUG "BIDON" Selector reseted => intialize it to current year (ex: Where = 2010)
            mFileFilter = "*" & Format(Now, "yyyy")                  'updates mFileFilter to current year
            'Debug.Print("  LFF => force mFileFilter to : " & mFileFilter)    'TEST/DEBUG
            Call DisplaySelector()                                   'displays selector buttons using mFileFilter selections
            Call InitDisplaySelectedFiles()                          'initializes display of all selected files, starting at first page
            Exit Sub
        End If

        FullDataPathName = mZaanDbPath & "data" & AccessGroup        'in v4 database format, AccessGroup directory is like "\data_group\", at same level as "\data\"
        If Not My.Computer.FileSystem.DirectoryExists(FullDataPathName) Then
            Debug.Print("=> LoadFilterFiles :  not existing directory : " & FullDataPathName)    'TEST/DEBUG
            Exit Sub
        End If

        dirInfo = GetDirInfoDataPathFilter(FullDataPathName, FileFilter, AccessGroup)   'get lowest data child directories corresponding to given filter series and loads related files in lvIn
        'Debug.Print("  LoadFilterFiles END :  last directory loaded = " & dirInfo.FullName)    'TEST/DEBUG
    End Sub

    Private Sub ChangeKeysInDataDirectory(ByVal DataPath As String, ByVal SceKey As String, ByVal DestKey As String)
        'Changes given source key by destination key in all data directories of current v3 database (as a result of tree node move, data hierarchical keys have to be updated)
        Dim dirInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(DataPath)
        Dim dirSel As System.IO.DirectoryInfo
        Dim DestDirName As String

        'Debug.Print("=> ChangeKeysInDataDirectory : DataPath = " & DataPath & "  SceKey = " & SceKey & "  DestKey = " & DestKey)   'TEST/DEBUG
        For Each dirSel In dirInfo.GetDirectories("*" & SceKey & "*", SearchOption.TopDirectoryOnly)    'scans all data directories using SceKey reference
            DestDirName = Replace(dirSel.Name, SceKey, DestKey)                                         'replaces SceKey by DestKey in local data directory name
            My.Computer.FileSystem.RenameDirectory(dirSel.FullName, DestDirName)                        'renames data directory with DesKey reference
        Next
    End Sub

    Private Sub ChangeKeysInDataDirectories(ByVal SceKey As String, ByVal DestKey As String)
        'Changes given source key by destination key in all data directories of current v3 database (as a result of tree node move, data hierarchical keys have to be updated)
        Dim NodeX As TreeNode

        Call ChangeKeysInDataDirectory(mZaanDbPath & "data", SceKey, DestKey)       'changes given source key by destination key in main data directory
        If mLicTypeCode >= 30 Then
            For Each NodeX In trvW.Nodes(0).Nodes                                   'scans all access control group child nodes
                Call ChangeKeysInDataDirectory(Mid(NodeX.Tag, 3), SceKey, DestKey)  'changes given source key by destination key in main data director
            Next
        End If
    End Sub

    Private Sub ReplaceHierarchicalKeysBySimpleKeys(ByRef DataDirName As String)
        'Replaces in given data directory name hierarchical keys by related simple keys (5 Char. in v3 database format)
        Dim FileNameKey, TreeCode, FileNameRefs() As String
        Dim i As Integer

        FileNameKey = ""
        If DataDirName <> "" Then
            FileNameRefs = Split(DataDirName, "_")
            If FileNameRefs.Length > 1 Then
                For i = 1 To FileNameRefs.Length - 1
                    If FileNameRefs(i) <> "" Then
                        TreeCode = Mid(FileNameRefs(i), 1, 1)
                        If (Len(FileNameRefs(i)) > 5) And (TreeCode <> "t") And (TreeCode <> "u") Then
                            FileNameRefs(i) = TreeCode & Microsoft.VisualBasic.Right(FileNameRefs(i), 4)     'get simple key (5 char.)
                        End If
                        FileNameKey = FileNameKey & "_" & FileNameRefs(i)      'rebuilds FileNameKey with simple keys
                    End If
                Next
            End If
        End If
        DataDirName = FileNameKey                                              'updates source directory name
    End Sub

    Private Sub ReplaceSimpleKeysByHierarchicalKeys(ByRef DataDirName As String)
        'Replaces in given data directory name simple keys (5 Char. in v3 database format, except for When/Date keys) by related hierarchical keys 
        Dim FileNameKey, TreeCode, FileNameRefs(), HierarchicalKey4 As String
        Dim i As Integer

        FileNameKey = ""
        If DataDirName <> "" Then
            FileNameRefs = Split(DataDirName, "_")
            If FileNameRefs.Length > 1 Then
                For i = 1 To FileNameRefs.Length - 1
                    If FileNameRefs(i) <> "" Then
                        TreeCode = Mid(FileNameRefs(i), 1, 1)
                        HierarchicalKey4 = Mid(FileNameRefs(i), 2, 4)          'initializes key to simple key
                        If (TreeCode <> "t") And (TreeCode <> "u") Then
                            Call ReplaceSimpleKey4ByHierarchicalKey4(TreeCode, HierarchicalKey4)   'replaces given simple key4 by related hierarchical key4 (including parent key4s)
                            FileNameKey = FileNameKey & "_" & TreeCode & HierarchicalKey4          'rebuilds FileNameKey with simple keys
                        Else
                            FileNameKey = FileNameKey & "_" & FileNameRefs(i)                      'rebuilds FileNameKey with initial keys
                        End If
                    End If
                Next
            End If
        End If
        DataDirName = FileNameKey                                              'updates source directory name
    End Sub

    Private Sub CountFilterFilesKEY(ByVal HierarchicalFileFilter As String, ByVal DataPath As String, ByVal GroupName As String)
        'Updates count of data files matching with given filter, in "data" directory or in given sub-directory DataAccessDirName
        Dim dirInfo, dirSel As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim FullDataPathName, Prefix, SelPattern, FileName, FileNameKey As String
        Dim isHiddenFile As Boolean

        If mTreeKeyLength = 0 Then Exit Sub '(Windows export) database format with no tree keys is not managed here

        'Debug.Print("  CountFilterFilesKEY :  HierarchicalFileFilter = " & HierarchicalFileFilter & "  DataPath = " & DataPath & "  GroupName = " & GroupName)    'TEST/DEBUG

        If HierarchicalFileFilter = "" And GroupName = "" Then Exit Sub

        'FullDataPathName = mZaanDbPath & "data" & DataAccessDirName
        FullDataPathName = DataPath
        If Not My.Computer.FileSystem.DirectoryExists(FullDataPathName) Then
            Debug.Print("=> CountFilterFilesKEY :  not existing directory : " & FullDataPathName)    'TEST/DEBUG
            Exit Sub
        End If

        dirInfo = My.Computer.FileSystem.GetDirectoryInfo(FullDataPathName)
        SelPattern = HierarchicalFileFilter & "*"

        For Each dirSel In dirInfo.GetDirectories(SelPattern, SearchOption.TopDirectoryOnly)
            'Debug.Print("=> CountFilterFilesKEY : dirSel.FullName = " & dirSel.FullName)  'TEST/DEBUG

            FileNameKey = dirSel.Name
            Call ReplaceHierarchicalKeysBySimpleKeys(FileNameKey)              'replaces in given data directory name hierarchical keys by related simple keys

            Try
                For Each fi In dirSel.GetFiles("*.*")
                    FileName = fi.Name
                    Prefix = Microsoft.VisualBasic.Left(FileName, 3)
                    isHiddenFile = False
                    If (fi.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then isHiddenFile = True
                    If Microsoft.VisualBasic.Left(FileName, 1) = "~" Then isHiddenFile = True

                    If FileName <> "." And FileName <> ".." And Prefix <> "zzi" And UCase(FileName) <> "THUMBS.DB" And Not isHiddenFile Then    'ignores current and encompassing directories, ZAAN and Windows XP thumbnail images and hidden files
                        mMatrixDocCount = mMatrixDocCount + 1                  'increments document counter
                    End If
                Next
            Catch ex As Exception
                Debug.Print("! CountFilterFilesKEY error : " & ex.Message)     'TEST/DEBUG : info directory doesn't exist anymore !
            End Try
        Next
    End Sub

    Private Sub LoadFilterFilesKEY(ByVal HierarchicalFileFilter As String, ByVal DataPath As String, ByVal GroupName As String)
        'Loads data files of directories matching with given hierarchical filter in given DataPath directory if non empty, else in default data directory
        Dim dirInfo, dirSel As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim FullDataPathName, Prefix, SelPattern, FileName, FileNameKey, FileType, DirPathName, WhenText, WhenKey As String
        Dim itmX As ListViewItem
        Dim itmXTab(11), whoName, whatName, whenName, whereName, what2Name, who2Name As String
        Dim i, FirstCount As Integer
        Dim isImage, isHiddenFile As Boolean

        If mTreeKeyLength = 0 Then Exit Sub '(Windows export) database format with no tree keys is not managed here

        'Debug.Print("  LoadFilterFilesKEY :  HierarchicalFileFilter = " & HierarchicalFileFilter & "  DataPath = " & DataPath & "  GroupName = " & GroupName)    'TEST/DEBUG

        If HierarchicalFileFilter = "" And GroupName = "" Then                 'TEST/DEBUG "BIDON" Selector reseted => intialize it to current year (ex: Where = 2010)
            WhenText = Format(Now, "yyyy")                                     'get current year
            WhenKey = "t" & GetWhenV3KeyFromDateText(WhenText)                 'returns When key in current v3+ database format
            mFileFilter = "*_" & WhenKey                                       'updates mFileFilter with current year only
            'Debug.Print("  LFF => force mFileFilter to : " & mFileFilter)     'TEST/DEBUG
            Call DisplaySelector()                                             'displays selector buttons using mFileFilter selections
            Call InitDisplaySelectedFiles()                                    'initializes display of all selected files, starting at first page
            Exit Sub
        End If

        'FullDataPathName = mZaanDbPath & "data" & DataAccessDirName
        FullDataPathName = DataPath
        If Not My.Computer.FileSystem.DirectoryExists(FullDataPathName) Then
            Debug.Print("=> LoadFilterFilesKEY :  not existing directory : " & FullDataPathName)    'TEST/DEBUG
            Exit Sub
        End If

        dirInfo = My.Computer.FileSystem.GetDirectoryInfo(FullDataPathName)
        SelPattern = HierarchicalFileFilter & "*"

        FirstCount = 20                                                        'number of items to be first displayed

        For Each dirSel In dirInfo.GetDirectories(SelPattern, SearchOption.TopDirectoryOnly)
            DirPathName = dirSel.FullName & "\"
            'Debug.Print("=> LoadFilterFilesKEY : DirPathName = " & DirPathName)  'TEST/DEBUG

            FileNameKey = dirSel.Name
            Call ReplaceHierarchicalKeysBySimpleKeys(FileNameKey)              'replaces in given data directory name hierarchical keys by related simple keys

            'If IsNewPiledInDir(GroupName & "\" & FileNameKey) Then
            If IsNewPiledInDir(dirSel.FullName) Then                           'case of given in directory is new in lstInDir pile of current selection
                whenName = GetFileTreeNodeText(FileNameKey, "t")                         'when
                whoName = GetFileTreeNodeText(FileNameKey, "o")                          'who
                whatName = GetFileTreeNodeText(FileNameKey, "a")                         'what
                whereName = GetFileTreeNodeText(FileNameKey, "e")                        'where
                what2Name = GetFileTreeNodeText(FileNameKey, "b")                        'what else
                who2Name = GetFileTreeNodeText(FileNameKey, "c")                         'other
                'Debug.Print(" => LoadFilterFilesKEY : " & dirSel.Name & " = " & groupName & "|" & whoName & "|" & whatName & "|" & whenName & "|" & whereName & "|" & what2Name & "|" & who2Name)  'TEST/DEBUG

                Try
                    For Each fi In dirSel.GetFiles("*.*")
                        FileName = fi.Name
                        Prefix = Microsoft.VisualBasic.Left(FileName, 3)
                        isHiddenFile = False
                        If (fi.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then isHiddenFile = True
                        If Microsoft.VisualBasic.Left(FileName, 1) = "~" Then isHiddenFile = True

                        If FileName <> "." And FileName <> ".." And Prefix <> "zzi" And UCase(FileName) <> "THUMBS.DB" And Not isHiddenFile Then    'ignores current and encompassing directories, ZAAN and Windows XP thumbnail images and hidden files
                            'FileType = GetFileTypeAndImage(FileName)                     'returns file type and get image from system if not already loaded in imgFileTypes
                            FileType = GetFileTypeAndImage(fi.FullName)                  'returns file type and get image from system if not already loaded in imgFileTypes

                            isImage = (FileType = ".jpg") Or (FileType = ".gif") Or (FileType = ".png") Or (FileType = ".bmp")
                            If isImage Then
                                mIsImageInLvIn = True                                    'sets flag indicating that at least one image is displayed in lvIn list
                            End If

                            If (lvIn.View = View.Details) Or isImage Then                'displays all docs in details view and only the images in icon view
                                'Debug.Print(" > File : " & FileName)  'TEST/DEBUG
                                mInDocCount = mInDocCount + 1                            'increments document counter
                                If (mInDocCount > (mInDocPageNb - 1) * mInDocPerPage) And (mInDocCount <= mInDocPageNb * mInDocPerPage) Then         'limits display list to mInDocPerPage of documents at current page
                                    itmX = lvIn.Items.Add(System.IO.Path.GetFileNameWithoutExtension(FileName))  'adds new item (file name) in list
                                    'If DataAccessDirName = "" Then
                                    '  itmX.Tag = dirSel.Name
                                    'Else
                                    '  itmX.Tag = Mid(DataAccessDirName, 2) & "\" & dirSel.Name
                                    'End If
                                    itmX.Tag = dirSel.FullName                           'SINCE 2012-03-30 : stores absolute path to file

                                    itmXTab(0) = GroupName                                         'access control

                                    itmXTab(1) = whenName                                          'when
                                    itmXTab(2) = whoName                                           'who
                                    itmXTab(3) = whatName                                          'what
                                    itmXTab(4) = whereName                                         'where

                                    itmXTab(5) = what2Name                                         'what else
                                    itmXTab(6) = who2Name                                          'other

                                    itmXTab(7) = FileType                                          'file type
                                    itmXTab(8) = Format(fi.Length / 1000, "### ### ###")           'file size
                                    itmXTab(9) = Format(fi.LastWriteTime, "d MMM yyyy  HH:mm")     'modification date
                                    itmXTab(10) = ""                                               'available for picture dimensions

                                    If lvIn.View = View.LargeIcon Then                             'displays item picture (as "large icon")
                                        pctLvIn.Image = My.Resources.picture_black_90x90           'initializes image with empty picture
                                        If isImage Then
                                            itmXTab(10) = GetLvInItemImageAndSize(dirSel, DirPathName, FileName, fi.LastWriteTime, FileType) & " pixels"   'get image size and loads related thumbnail image in pctLvIn
                                        End If
                                        imgLargeIcons.Images.Add(pctLvIn.Image)                    'loads thumbnail image in large image list
                                        'itmX.ImageIndex = mInDocCount - 1
                                        itmX.ImageIndex = mInDocCount - 1 - (mInDocPageNb - 1) * mInDocPerPage
                                    Else                                                           'displays item "details"
                                        itmX.ImageKey = FileType
                                    End If
                                    For i = 0 To 9
                                        itmX.SubItems.Add(itmXTab(i))                              'fills access control, who, what, when, where, what else, other, type, size and date columns
                                    Next
                                    itmX.ToolTipText = FileType & "  "
                                    itmX.ToolTipText = itmX.ToolTipText & itmXTab(8) & " Kb" & vbCrLf & itmXTab(9) & vbCrLf & itmXTab(10) & vbCrLf
                                    itmX.ToolTipText = itmX.ToolTipText & itmXTab(0) & vbCrLf & itmXTab(1) & vbCrLf & itmXTab(2) & vbCrLf & itmXTab(3) & vbCrLf & itmXTab(4) & vbCrLf & itmXTab(5) & vbCrLf & itmXTab(6)

                                    'If mInDocCount Mod FirstCount = 0 Then
                                    '  lvIn.RedrawItems(mInDocPrevCount, mInDocCount - 1, False) 'forces partial display of (new) items within given index range
                                    '  mInDocPrevCount = mInDocCount
                                    'End If
                                End If
                            End If
                        End If
                    Next
                Catch ex As Exception
                    Debug.Print("! LoadFilterFilesKEY error : " & ex.Message)                        'TEST/DEBUG : info directory doesn't exist anymore !
                End Try
            Else
                'Debug.Print("! LFF => existing directory in pile : " & groupName & "\" & FileNameKey)    'TEST/DEBUG
            End If
        Next
        'Debug.Print("End of LoadFilterFilesKEY :  mIsImageInLvIn = " & mIsImageInLvIn)    'TEST/DEBUG
    End Sub

    Private Sub DisplayLvInDocCountAndElapsedTime(ByVal startTime As DateTime)
        'Displays lvIn document count (mInDocCount) and elapsed time (in seconds) since given start time
        Dim stopTime As DateTime
        Dim diffTime As TimeSpan
        Dim diffTimeText As String
        Dim DisplayedDocNb As Integer

        stopTime = Now
        'Debug.Print("=> stopTime : " & Format(stopTime, "H:mm:ss:fff"))   'TEST/DEBUG
        diffTime = stopTime.Subtract(startTime)
        'diffTimeText = mMessage(8) & diffTime.Minutes & "mn " & diffTime.Seconds & "s " & diffTime.Milliseconds & "ms"   'Found in : ...
        diffTimeText = mMessage(8) & Format(diffTime.TotalSeconds, "###0.00") & " s"   'Found in : 1.38 s (ex.)
        'Debug.Print("=> elapsed time : " & diffTimeText)   'TEST/DEBUG

        'If mInDocCount > mInDocPerPage Then                                                      'maximum document display reached
        '  If lvIn.View = View.Details Then
        '    lvIn.Columns(0).Text = mInDocPerPage & "/" & mInDocCount & " " & mMessage(82) & " (" & diffTimeText & ")"    'displays document number followed by (searched time)
        '  End If
        '  MsgBox(mMessage(202) & " " & mInDocPerPage & " " & mMessage(203), MsgBoxStyle.Exclamation)    'Only the first XX documents have been displayed !
        'Else
        '  If lvIn.View = View.Details Then
        '    lvIn.Columns(0).Text = mInDocCount & " " & mMessage(82) & " (" & diffTimeText & ")"    'displays document number followed by (searched time)
        '  End If
        'End If

        mInDocPageMax = 1 + ((mInDocCount - 1) \ mInDocPerPage)                'updates max number of pages
        lblSelPage.Text = mInDocPageNb & " / " & mInDocPageMax                 'updates display of current page number
        If mInDocPageNb < mInDocPageMax Then
            lblSelPageNext.Enabled = True                                      'enables next page ">" label only if page max not reached
        Else
            lblSelPageNext.Enabled = False
        End If
        If mInDocPageNb > 1 Then
            lblSelPagePrev.Enabled = True                                      'enables previous page "<" label only if page 1 not reached
        Else
            lblSelPagePrev.Enabled = False
        End If

        If lvIn.View = View.Details Then
            If mInDocCount > mInDocPerPage Then                                'documents number exceeds number of documents per page
                If mInDocPageNb = mInDocPageMax Then
                    DisplayedDocNb = mInDocCount - (mInDocPageNb - 1) * mInDocPerPage
                Else
                    DisplayedDocNb = mInDocPerPage
                End If
                lvIn.Columns(0).Text = DisplayedDocNb & " / " & mInDocCount & " " & mMessage(82) & " (" & diffTimeText & ")"     'displays document number followed by (searched time)
            Else
                lvIn.Columns(0).Text = mInDocCount & " " & mMessage(82) & " (" & diffTimeText & ")"     'displays document count followed by (searched time)
            End If
        End If
    End Sub

    Private Function GetAccessGroupFileFilter() As String
        'Returns Access group from mFileFilter if any existing, else an empty string
        Dim FileFilters() As String = Split(mFileFilter, "*")
        Dim AccessGroup As String = ""
        Dim i As Integer

        For i = 1 To FileFilters.Length - 1
            If Mid(FileFilters(i), 1, 2) = "z0" Then                 'Access group filter exists
                AccessGroup = Mid(FileFilters(i), 4)
                Exit For
            End If
        Next
        GetAccessGroupFileFilter = AccessGroup
    End Function

    Private Sub InitDisplaySelectedFiles()
        'Initializes display of all selected files using current mFileFilter, starting at first page
        mInDocPageNb = 1
        mInDocPageMax = 1
        Call UpdatePagePileButtons()                       'updates page pile (lstPage) if mFileFilter is new in pile and enables/desables accordingly Prev/Next navigation buttons

        lvIn.Visible = True                                'shows selected files display
        Call DisplaySelectedFiles()                        'refreshes display of "in" files selected with current filter
    End Sub

    Private Function CountOfSelectedFiles(ByVal FileFilter As String) As Integer
        'Returns count of selected files matching with given FileFilter
        Dim DataPath As String
        Dim NodeX As TreeNode

        mMatrixDocCount = 0                                                    'initializes file/document counter (updated in CountFilterFiles() sub)

        DataPath = mZaanDbPath & "data"                                        'sets default data directory
        If mLicTypeCode < 30 Then                                              'case of ZAAN-First and ZAAN-Basic licenses
            Call CountDataSelectedFiles(DataPath, "", FileFilter)                        'displays selected files related to "data" root directory
        Else                                                                   'case of ZAAN-Pro license (access groups possible)
            If lblDataAccess.Tag = "" Then                                     'no "Access control" group selected
                Call CountDataSelectedFiles(DataPath, "", FileFilter)                    'displays selected files related to "data" root directory
                For Each NodeX In trvW.Nodes(0).Nodes                          'scans all child nodes (groups) of "Access control" root node (dim 0)
                    Call CountDataSelectedFiles(Mid(NodeX.Tag, 3), NodeX.Text, FileFilter)    'displays selected files related to this access control group only
                Next
            ElseIf lblDataAccess.Tag.Length > 3 Then                           'Access control group is selected (in dim 0 of Selector)
                Call CountDataSelectedFiles(Mid(lblDataAccess.Tag, 4), lblDataAccess.Text, FileFilter)  'displays selected files related to this access control group only
            End If
        End If
        CountOfSelectedFiles = mMatrixDocCount
    End Function

    Private Sub DisplaySelectedFiles()
        'Displays all selected files including access controlled (dim 0) data if user has access to related directories (listed Access control groups)
        Dim AccessGroup, DataPath As String
        'Dim DataAccessDirName As String
        Dim NodeX As TreeNode
        Dim startTime As DateTime

        'Debug.Print("=> DisplaySelectedFiles...")    'TEST/DEBUG
        startTime = Now                                              'starts search stopwatch (used for performance test)
        'Debug.Print("DisplaySelectedFiles startTime : " & Format(startTime, "H:mm:ss:fff"))   'PERFORMANCE TEST : result in thousands of sec.

        Me.Cursor = Cursors.WaitCursor                               'sets wait cursor
        fswData.EnableRaisingEvents = False                          'locks fswData related events (that eventual thumbnail image files creation would trigger)
        lvIn.ListViewItemSorter = Nothing                            'IMPORTANT : resets the listview item sorter property for saving time (if formerly allocated to a ListViewItemComparer or ListViewItemRevComparer object) 

        lvIn.Items.Clear()                                           'clears lvIn items list
        lstInDir.Items.Clear()                                       'clears lstInDir in directories list
        imgLargeIcons.Images.Clear()                                 'clears large icons list

        mInDocCount = 0                                              'initializes file/document counter (updated in LoadFilterFiles() sub)
        mInDocPrevCount = 0
        mIsImageInLvIn = False                                       'initializes flag that will indicate if (at least) one image file is displayed in list

        'If btnDataAccessNoSel.Visible Then                           'data access selector set at tree loading
        '  If lvIn.Columns(1).Width = 0 Then lvIn.Columns(1).Width = 85 'shows data access column
        'Else
        '  If lvIn.Columns(1).Width <> 0 Then lvIn.Columns(1).Width = 0 'hides data access column
        'End If

        If mTreeKeyLength = 0 Then                                             'no node keys in v4 database format (use directly Windows path names)
            AccessGroup = GetAccessGroupFileFilter()
            If mLicTypeCode < 30 Then                                          'case of ZAAN-First and ZAAN-Basic licenses
                Call LoadFilterFiles("")                                       'loads data files matching with given file filter
            Else                                                               'case of ZAAN-Pro license (access groups possible)
                If AccessGroup = "" Then                                       'no "Access control" group selected
                    Call LoadFilterFiles("")                                   'loads data files matching with given file filter
                    For Each NodeX In trvW.Nodes(0).Nodes                      'scans all child nodes (groups) of "Access control" root node (dim 0)
                        Call LoadFilterFiles("_" & NodeX.Text)                 'loads given group data files matching with given file filter
                    Next
                Else                                                           'Access control group is selected (in dim 0 of Selector)
                    Call LoadFilterFiles("_" & AccessGroup)                    'loads data files matching with given file filter
                End If
            End If
            Call DisplayLvInDocCountAndElapsedTime(startTime)                  'displays lvIn document count (mInDocCount) and elapsed time (in seconds) since given start time
            'Call LoadLvInItemsDimReferences()                                  'loads dimension references of all existing lvIn items (in v4 database format)
        Else
            DataPath = mZaanDbPath & "data"                                    'sets default data directory
            If mLicTypeCode < 30 Then                                          'case of ZAAN-First and ZAAN-Basic licenses
                Call DisplayDataSelectedFiles(DataPath, "")                    'displays selected files related to "data" root directory
            Else                                                               'case of ZAAN-Pro license (access groups possible)
                If lblDataAccess.Tag = "" Then                                 'no "Access control" group selected
                    Call DisplayDataSelectedFiles(DataPath, "")                'displays selected files related to "data" root directory
                    For Each NodeX In trvW.Nodes(0).Nodes                      'scans all child nodes (groups) of "Access control" root node (dim 0)
                        'DataAccessDirName = "\data_" & NodeX.Text
                        'Call DisplayDataSelectedFiles(DataAccessDirName)       'displays selected files related to this access control group only
                        Call DisplayDataSelectedFiles(Mid(NodeX.Tag, 3), NodeX.Text)               'displays selected files related to this access control group only
                    Next
                ElseIf lblDataAccess.Tag.Length > 3 Then                       'Access control group is selected (in dim 0 of Selector)
                    'DataAccessDirName = "\data_" & Mid(lblDataAccess.Tag, 4)
                    'Call DisplayDataSelectedFiles(DataAccessDirName)           'displays selected files related to this access control group only
                    Call DisplayDataSelectedFiles(Mid(lblDataAccess.Tag, 4), lblDataAccess.Text)   'displays selected files related to this access control group only
                End If
            End If
            Call DisplayLvInDocCountAndElapsedTime(startTime)                  'displays lvIn document count (mInDocCount) and elapsed time (in seconds) since given start time
        End If

        If Not mLockSortLvIn Then
            Call SortLvInColumn()                                    'refreshs "in" files display with column sorting set in mlvInOrder() or first one if no details view
        End If

        If lvIn.View = View.Details Then
            lvIn.ShowItemToolTips = False                            'hides item tooltips in "details" view
        Else
            lvIn.ShowItemToolTips = True                             'shows item tooltips in "large icon" view
        End If

        fswData.EnableRaisingEvents = True                           'unlocks fswData related events
        Me.Cursor = Cursors.Default                                  'resets default cursor

        Call LoadSelectorParentList()                                'loads Selector parent list with possible parent node selections of current selector positions
    End Sub

    Private Sub ReplaceSimpleKey4ByHierarchicalKey4(ByVal TreeCode As String, ByRef HierarchicalKey4 As String)
        'Replaces given simple key4 (4 char. in v3 database format) by related hierarchical key4 (including parent key4s)
        Dim ParentKey4 As String
        Dim NodeX(), ParentNode As TreeNode

        If (TreeCode = "t") Or (TreeCode = "u") Then Exit Sub

        NodeX = trvW.Nodes.Find(TreeCode & HierarchicalKey4, True)        'searches for tree node with given simple key4
        If NodeX.Length > 0 Then                                          'node key found in tree nodes => get related hierarchical key
            ParentNode = NodeX(0).Parent
            Do While Not ParentNode Is Nothing                            'if parent node exists, then adds its key4 to top of hierarchical key
                ParentKey4 = Mid(ParentNode.Tag, 3, 4)
                If ParentKey4 = mTreeRootKey Then                         'case of  tree root reached (like "1 - Who" with "zzzy" Key4)
                    Exit Do
                Else
                    HierarchicalKey4 = ParentKey4 & HierarchicalKey4      'appends parent key4 before current hierarchical key being built
                    ParentNode = ParentNode.Parent
                End If
            Loop
        End If
    End Sub

    Private Function GetFiltered4ZWhenKey(ByVal WhenKey As String) As String
        'Returns filtered 4Z Whenkey, skipping undefined "z" month and "z" day in v3+ database format (used in data directory searches)
        Dim FilteredKey As String = WhenKey

        Do While Len(FilteredKey) > 3                                     'scans characters after "tYY" (ex : "trh" coding 2011 year), that is 1Z char. for month and 1Z char. for day
            If Microsoft.VisualBasic.Right(FilteredKey, 1) = "z" Then
                FilteredKey = Mid(FilteredKey, 1, Len(FilteredKey) - 1)   'eliminates undefined month or day "z" code
            Else
                Exit Do
            End If
        Loop
        GetFiltered4ZWhenKey = FilteredKey
    End Function

    Private Function GetHierarchicalKey(ByVal SimpleKey As String, Optional ByVal Filter4ZWhenKey As Boolean = False) As String
        'Returns hierarchical key corresponding to given simple key if v3 database format, same simple key else
        Dim HierarchicalKey As String = SimpleKey
        Dim TreeCode As String = Mid(SimpleKey, 1, 1)
        Dim Key4 As String = Mid(SimpleKey, 2, 4)

        If (TreeCode <> "t") And (TreeCode <> "u") Then                        'case of o/a/e/b/c tree codes
            Call ReplaceSimpleKey4ByHierarchicalKey4(TreeCode, Key4)           'replaces given simple key4 by related hierarchical key4 (including parent key4s)
            HierarchicalKey = TreeCode & Key4
        Else
            If (TreeCode = "t") And Filter4ZWhenKey Then                       'case of v3+ When key with a 4Z When key filter request
                HierarchicalKey = GetFiltered4ZWhenKey(SimpleKey)              'rebuilds v3+ When key, skipping undefined "z" month and "z" day
            End If
        End If
        GetHierarchicalKey = HierarchicalKey
    End Function

    Private Function GetHierarchicalFilter(ByVal SceFileFilter As String) As String
        'Returns current file filter with v3/v3+ keys (5 char.) replaced with related hierarchical keys (5+ char. including parent key4s) and without "u" key
        Dim FileCodeTable() As String
        Dim FileFilter As String
        Dim TreeCode, HierarchicalKey4 As String
        Dim i As Integer

        'Debug.Print("GetHierarchicalFilter :  SceFileFilter = " & SceFileFilter)   'TEST/DEBUG
        FileFilter = ""
        If SceFileFilter <> "" Then
            FileCodeTable = Split(SceFileFilter, "*_")
            If FileCodeTable.Length > 1 Then
                For i = 1 To FileCodeTable.Length - 1                          'scans all filter keys for replacing them with related hierarchical keys
                    If FileCodeTable(i) <> "" Then
                        TreeCode = Mid(FileCodeTable(i), 1, 1)
                        If TreeCode <> "u" Then                                'skips "u" key
                            HierarchicalKey4 = Mid(FileCodeTable(i), 2, 4)     'initializes key to simple key
                            If TreeCode <> "t" Then
                                Call ReplaceSimpleKey4ByHierarchicalKey4(TreeCode, HierarchicalKey4)         'replaces given simple key4 by related hierarchical key4 (including parent key4s)
                                FileFilter = FileFilter & "*_" & TreeCode & HierarchicalKey4                 'rebuilds FileFilter with hierarchical key
                            Else
                                FileFilter = FileFilter & "*_" & GetFiltered4ZWhenKey(FileCodeTable(i))      'rebuilds FileFilter with v3+ When key, skipping undefined "z" month and "z" day
                            End If
                        End If
                    End If
                Next
            End If
        End If
        'Debug.Print(" => GetHierarchicalFilter :  FileFilter = " & FileFilter)   'TEST/DEBUG
        GetHierarchicalFilter = FileFilter
    End Function

    Private Sub CountDataSelectedFiles(ByVal DataPath As String, ByVal GroupName As String, ByVal FileFilter As String)
        'Displays selected files in given data directory

        Call CountFilterFilesKEY(GetHierarchicalFilter(FileFilter), DataPath, GroupName)      'loads data files matching with given file filter
    End Sub

    Private Sub DisplayDataSelectedFiles(ByVal DataPath As String, ByVal GroupName As String)
        'Displays selected files in given data directory

        Call LoadFilterFilesKEY(GetHierarchicalFilter(mFileFilter), DataPath, GroupName)      'loads data files matching with current file filter
    End Sub

    Private Sub DisplaySlideShowPicture()
        'Displays given mSlideShowIndex picture from current selection in zoom view
        Dim SelPathFile, FileName, SceFileType, FileType As String
        Dim isImage As Boolean = False

        If mSlideShowIndex > lvIn.Items.Count - 1 Then Exit Sub

        SceFileType = lvIn.Items(mSlideShowIndex).SubItems(8).Text
        FileType = UCase(SceFileType)
        isImage = (FileType = ".JPG") Or (FileType = ".GIF") Or (FileType = ".PNG") Or (FileType = ".BMP")
        Do While Not isImage                                                        'searches fo next image document, if any
            mSlideShowIndex = mSlideShowIndex + 1
            If mSlideShowIndex > lvIn.Items.Count - 1 Then Exit Do
            FileType = lvIn.Items(mSlideShowIndex).SubItems(8).Text
            isImage = (FileType = ".JPG") Or (FileType = ".GIF") Or (FileType = ".PNG") Or (FileType = ".BMP")
        Loop
        mSlideShowPathFileName = ""                                                 'sets file name to no image/slide by default
        If isImage Then
            'Debug.Print("DisplaySlideShowPicture :  mSlideShowIndex = " & mSlideShowIndex)   'TEST/DEBUG
            'SelPathFile = mZaanDbPath & "data\" & lvIn.Items(mSlideShowIndex).Tag     'gets selected document "in" file
            SelPathFile = lvIn.Items(mSlideShowIndex).Tag                           'gets selected document "in" file

            FileName = lvIn.Items(mSlideShowIndex).Text & SceFileType
            mSlideShowPathFileName = SelPathFile & "\" & FileName                   'stores current valid image/slide file name
            lblZoomName.Text = lvIn.Items(mSlideShowIndex).Text
            lblSlideNb.Text = mSlideShowIndex + 1 & " / " & lvIn.Items.Count        'updates current slide number display
            If mIsSlideShowEffect Then                                              'progressive zooming of next image from center
                Dim image As New Bitmap(mSlideShowPathFileName)                     'loads source image (full size)
                mSlideShowImageRatio = image.Width / image.Height
                'Debug.Print("image size = " & image.Width & "x" & image.Height & "  mSlideShowImageRatio = " & mSlideShowImageRatio)    'TEST/DEBUG
                If mSlideShowImageRatio > 1 Then                                    'is a horizontal picture
                    pctZoom2.Width = pctZoom.Width / mSlideShowEffectMax
                    pctZoom2.Height = pctZoom2.Width / mSlideShowImageRatio
                Else                                                                'is a vertical picture
                    pctZoom2.Height = pctZoom.Height / mSlideShowEffectMax
                    pctZoom2.Width = pctZoom2.Height * mSlideShowImageRatio
                End If
                pctZoom2.Top = pctZoom.Top + pctZoom.Height / 2 - pctZoom2.Height / 2
                pctZoom2.Left = pctZoom.Left + pctZoom.Width / 2 - pctZoom2.Width / 2
                pctZoom2.Load(SelPathFile & "\" & FileName)                         'loads picture from related file
                pctZoom2.Visible = True
                mSlideShowEffectIndex = 0
                tmrSlideShowEffect.Enabled = True                                   'starts slide show effect timer
                image.Dispose()
            Else                                                                    'direct display of next image in "full screen"
                pctZoom2.Top = pctZoom.Top
                pctZoom2.Left = pctZoom.Left
                pctZoom2.Width = pctZoom.Width
                pctZoom2.Height = pctZoom2.Width
                pctZoom2.Visible = False
                pctZoom2.Load(mSlideShowPathFileName)                               'loads new picture from related file
                pctZoom.Image = pctZoom2.Image                                      'updates current display with new picture
            End If
        End If
    End Sub

    Private Function GetEncoder(ByVal format As ImageFormat) As ImageCodecInfo
        'Get available image encoder
        Dim codecs As ImageCodecInfo() = ImageCodecInfo.GetImageDecoders()
        Dim codec As ImageCodecInfo
        For Each codec In codecs
            If codec.FormatID = format.Guid Then
                Return codec
            End If
        Next codec
        Return Nothing
    End Function

    Private Sub RotateSlideShowPicture(ByVal RotFlip As RotateFlipType)
        'Rotates currently selected picture to left, in the specified number of quadrants (90° angle steps)
        Dim Response As MsgBoxResult
        Dim DirPathName, FileName As String

        'Debug.Print("RotateSlideShowPicture : RotFlip = " & RotFlip & "  file = " & mSlideShowPathFileName)   'TEST/DEBUG
        If mSlideShowPathFileName Is Nothing Then Exit Sub

        mIsSlideShowPaused = True                                    'makes sure that play/pause status is on pause
        Call PlayPauseSlideShow()                                    'plays or stops activated slide show

        pctZoom.Image.RotateFlip(RotFlip)                            'rotate displayed picture using given RotFlip value
        pctZoom.Refresh()                                            'forces picture re-painting

        Response = MsgBox(mMessage(137), MsgBoxStyle.YesNo + MsgBoxStyle.Information)    'Do you confirm this image rotation ?
        If Response = MsgBoxResult.Yes Then
            Me.Cursor = Cursors.WaitCursor                           'sets wait cursor
            DirPathName = System.IO.Path.GetDirectoryName(mSlideShowPathFileName)
            FileName = System.IO.Path.GetFileName(mSlideShowPathFileName)
            Call DeleteFileImage(DirPathName, FileName)              'deletes corresponding thumbnail image file (zzi*) that may have been generated by ZAAN

            Dim FileType As String = LCase(System.IO.Path.GetExtension(mSlideShowPathFileName))
            Dim image As New Bitmap(mSlideShowPathFileName)          'loads source image (full size)
            Dim imageFmt As ImageFormat = image.RawFormat            'get image format (jpg, gif, etc.)

            image.RotateFlip(RotFlip)                                'rotate image using given RotFlip value

            If FileType = ".jpg" Then                                'case of jpeg : save rotated image with optimized compression rate
                Dim jgpEncoder As ImageCodecInfo = GetEncoder(ImageFormat.Jpeg)
                Dim myEncoder As System.Drawing.Imaging.Encoder = System.Drawing.Imaging.Encoder.Quality
                Dim myEncoderParameters As New EncoderParameters(1)
                Dim myEncoderParameter As New EncoderParameter(myEncoder, 100&)          'between 0 for highest compression and 100 for lowest compression (85 is usually OK)
                myEncoderParameters.Param(0) = myEncoderParameter
                image.Save(mSlideShowPathFileName, jgpEncoder, myEncoderParameters)      'saves rotated image in jpeg format with given compression parameter
            Else
                image.Save(mSlideShowPathFileName, imageFmt)         'save rotated image in its source format
            End If

            image.Dispose()
            Me.Cursor = Cursors.Default                              'resets default cursor
        Else
            pctZoom.Load(mSlideShowPathFileName)                     'loads rotated image from related file
        End If
    End Sub

    Private Sub PauseSlideShow()
        'Pauses slide show and resets current ZAAN navigation view

        tmrSlideShow.Enabled = False                                 'stops slide show transition timer
        pctZoom2.Visible = False                                     'hides transition picture
        tmrSlideShowEffect.Enabled = False                           'stops slide show effect timer
    End Sub

    Private Sub StartStopSlideShow()
        'Starts or stops slide show depending on checked control using current selection and given Interval

        'Debug.Print("StartStopSlideShow :  Interval (sec.) = " & tstbPctZoom.Text)    'TEST/DEBUG
        If lvIn.Items.Count = 0 Then
            tsmiPctZoomSlideShow.Checked = False                          'stops/closes slideshow if picture list is empty
        End If

        mIsSlideShowEffect = False                                        'no special effect by default
        'mIsSlideShowEffect = True                                         'progressive zooming of next image from center

        If tsmiPctZoomSlideShow.Checked Then                              'STARTS slide show
            Me.WindowState = FormWindowState.Maximized                    'sets ZAAN window to maximized display
            Me.FormBorderStyle = Forms.FormBorderStyle.None

            'mIsSlideShowLeftPanelHidden = SplitContainer2.Panel1Collapsed      'stores current display status relative to left panel
            'SplitContainer2.Panel1Collapsed = True                        'makes sure that left panel is collapsed
            'SplitContainer3.Panel1Collapsed = True                        'hides current ZAAN selection view (lvIn)
            'pnlListTop.Visible = False                                    'hides selector panel
            mIsSlideShowBottomPanelHidden = SplitContainer1.Panel2Collapsed
            SplitContainer1.Panel2Collapsed = True                        'hides import/copy panel

            pnlZoom.Parent = SplitContainer1.Panel1
            tcCubes.Visible = False

            lblZoomName.Text = lvIn.Items(mSlideShowIndex).Text                'updates slide name with related filename without extension
            lblSlideNb.Text = mSlideShowIndex + 1 & " / " & lvIn.Items.Count   'updates current slide number display
            pnlSlideControl.Visible = True                                'shows slide show control panel

        Else                                                              'STOPS slide show
            Me.WindowState = FormWindowState.Normal                       'resets ZAAN window to normal display
            Me.FormBorderStyle = Forms.FormBorderStyle.Sizable

            'SplitContainer2.Panel1Collapsed = mIsSlideShowLeftPanelHidden      'resets initial display status relative to left panel
            'SplitContainer3.Panel1Collapsed = False                       're-displays current ZAAN selection view (lvIn)
            'pnlListTop.Visible = True                                     'shows selector panel
            SplitContainer1.Panel2Collapsed = mIsSlideShowBottomPanelHidden    'resets initial display status relative to bottom panel

            pnlZoom.Parent = SplitContainer3.Panel2
            tcCubes.Visible = True

            If mIsImageViewBeforeSlideShow <> mIsImageView Then
                Call SetLvInDisplayMode(mIsImageViewBeforeSlideShow)      'resets initial lvIn view mode
                Call InitDisplaySelectedFiles()                           'displays all "in" files (empty filter)
            End If

            pnlSlideControl.Visible = False                               'hides slide show control panel
        End If
        Call SetLeftPanelButton()                                         'sets left panel button display depending on splitter and background color
        Call SetBottomPanelButton()                                       'sets bottom panel button display depending on splitter and background color

        mIsSlideShowPaused = Not tsmiPctZoomSlideShow.Checked
        Call PlayPauseSlideShow()                                         'plays or stops activated slide show
    End Sub

    Private Sub PlayPauseSlideShow()
        'Plays or stops activated slide show

        If mIsSlideShowPaused Then
            btnSlidePlayPause.Image = My.Resources.play              'updates play/pause button image to play
        Else
            btnSlidePlayPause.Image = My.Resources.pause             'updates play/pause button image to pause
        End If

        If mIsSlideShowPaused Then                                   '=> pauses slide show
            Call PauseSlideShow()                                    'pauses slide show and resets current ZAAN navigation view
        Else                                                         '=> plays slide show
            tmrSlideShow.Enabled = False
            tmrSlideShow.Interval = 1000 * tstbPctZoom.Text
            If mSlideShowIndex = -1 Then                             'no valid slide show index
                If lvIn.SelectedItems.Count = 0 Then
                    'mSlideShowIndex = 0                              'starts slide show at first document by default
                Else
                    mSlideShowIndex = lvIn.SelectedItems(0).Index    'starts slide show at current selection
                End If
            End If
            If mIsSlideShowEffect Then                               'case of progressive zooming of next image from center
                mSlideShowIndex = mSlideShowIndex + 1                'moves index to next picture
                Call DisplaySlideShowPicture()                       'displays given mSlideShowIndex picture from current selection in zoom view
            Else                                                     'case of direct display of next image in "full screen"
                tmrSlideShow.Enabled = True                          'starts slide show transition timer
            End If
        End If
    End Sub

    Private Sub GotoPreviousSlide()
        'Moves current slide cursor to previous slide and displays related slide
        If mSlideShowIndex > 0 Then
            mSlideShowIndex = mSlideShowIndex - 1                    'moves index to previous picture
        Else
            mSlideShowIndex = lvIn.Items.Count - 1                   'moves index to last picture
        End If
        Call DisplaySlideShowPicture()                               'displays given mSlideShowIndex picture from current selection in zoom view
    End Sub

    Private Sub GotoNextSlide()
        'Moves current slide cursor to next slide and displays related slide
        If mSlideShowIndex < lvIn.Items.Count - 1 Then
            mSlideShowIndex = mSlideShowIndex + 1                    'moves index to next picture
        Else
            mSlideShowIndex = 0                                      'moves index to first picture
        End If
        Call DisplaySlideShowPicture()                               'displays given mSlideShowIndex picture from current selection in zoom view
    End Sub

    Private Sub SortLvInColumn()
        'Refreshs "in" files display with last column sorting or first one if no details view
        Dim i As Integer

        'Debug.Print("SortLvInColumn...")     'TEST/DEBUG
        lvIn.Sorting = SortOrder.None                                'cancels automatic ordering mode (VERY IMPORTANT FOR ENABLING COMPARER TO WORK !)

        If lvIn.View = View.Details Then                             'case of details view : multiple columns are available
            For i = 0 To mlvInOrder.Length - 1
                If mlvInOrder(i) <> 0 Then                           'this i column has a sort flag set
                    'Debug.Print("SortLvInColumn :  mlvInOrder(" & i & ") = " & mlvInOrder(i))   'TEST/DEBUG
                    If mlvInOrder(i) = 1 Then
                        lvIn.ListViewItemSorter = New ListViewItemComparer(i)      'set ascending order
                    Else
                        lvIn.ListViewItemSorter = New ListViewItemRevComparer(i)   'set descending order
                    End If
                    Exit For
                End If
            Next
        Else                                                         'if no details view, multiple columns are not available
            Call InitlvInOrder(0, True)                              'initializes mlvInOrder() table to first index column in ascending order
            lvIn.ListViewItemSorter = New ListViewItemComparer(0)    'sorts documents in ascending (text) order
        End If
    End Sub

    Private Sub SortLvOutColumn()
        'Refreshs "out" files display with last column sorting
        Dim i As Integer

        'Debug.Print("SortLvOutColumn...")   'TEST/DEBUG
        lvOut.Sorting = SortOrder.None                               'cancels automatic ordering mode (VERY IMPORTANT FOR ENABLING COMPARER TO WORK !)
        For i = 0 To mlvOutOrder.Length - 1
            If mlvOutOrder(i) <> 0 Then
                'Debug.Print("mlvOutOrder(" & i & ")=" & mlvOutOrder(i))             'TEST/DEBUG
                If mlvOutOrder(i) = 1 Then
                    lvOut.ListViewItemSorter = New ListView2ItemComparer(i)         'set ascending order
                Else
                    lvOut.ListViewItemSorter = New ListView2ItemRevComparer(i)      'set descending order
                End If
                Exit For
            End If
        Next
    End Sub

    Private Sub SortLvTempColumn()
        'Refreshs "temp" files display with last column sorting
        Dim i As Integer

        'Debug.Print("SortLvTempColumn...")   'TEST/DEBUG
        lvTemp.Sorting = SortOrder.None                              'cancels automatic ordering mode (VERY IMPORTANT FOR ENABLING COMPARER TO WORK !)
        For i = 0 To mlvTempOrder.Length - 1
            If mlvTempOrder(i) <> 0 Then
                'Debug.Print("mlvTempOrder(" & i & ")=" & mlvTempOrder(i))           'TEST/DEBUG
                If mlvTempOrder(i) = 1 Then
                    lvTemp.ListViewItemSorter = New ListView2ItemComparer(i)        'set ascending order
                Else
                    lvTemp.ListViewItemSorter = New ListView2ItemRevComparer(i)     'set descending order
                End If
                Exit For
            End If
        Next
    End Sub

    Private Sub DisplayOutFiles(Optional ByVal ShowContent As Boolean = True, Optional ByVal TabIndex As Integer = -1)
        'Displays "out" files stored in "ZAAN/import" directory, which is the import/export area of ZAAN.
        Dim fi As System.IO.FileInfo
        Dim FileName, FileType As String
        Dim i As Long
        Dim itmX As ListViewItem
        Dim isHiddenFile As Boolean
        Dim CurTabIndex As Integer = TabIndex

        'Debug.Print("1 - DisplayOutFiles :  ShowContent = " & ShowContent & "  TabIndex = " & TabIndex & "  mZaanImportPath =" & mZaanImportPath)    'TEST/DEBUG
        Select Case TabIndex
            Case 0
                mZaanImportPath = mMyZaanImportPath
            Case 1
                mZaanImportPath = mImportPath
            Case Else
                If mZaanImportPath = mMyZaanImportPath Then
                    CurTabIndex = 0
                Else
                    CurTabIndex = 1
                End If
        End Select
        'Debug.Print(" 2 - DisplayOutFiles :  ShowContent = " & ShowContent & "  CurTabIndex = " & CurTabIndex & "  mZaanImportPath =" & mZaanImportPath)    'TEST/DEBUG
        Dim dirInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanImportPath)

        i = 0                                                        'initializes file/document counter
        If ShowContent Then
            lvTemp.ListViewItemSorter = Nothing                      'IMPORTANT : ListView2ItemComparer or ListView2ItemRevComparer may have been used for LvTemp ! 
            lvOut.ListViewItemSorter = Nothing                       'IMPORTANT : resets the listview item sorter property for saving time (if formerly allocated to a ListView2ItemComparer or ListView2ItemRevComparer object) 
            lvOut.Items.Clear()
        End If

        For Each fi In dirInfo.GetFiles("*.*")
            'Debug.Print("=> DisplayOutFiles :  file =" & fi.Name & " => " & fi.Attributes.ToString)    'TEST/DEBUG
            FileName = fi.Name
            isHiddenFile = False
            If (fi.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then isHiddenFile = True
            If Microsoft.VisualBasic.Left(FileName, 1) = "~" Then isHiddenFile = True

            If FileName <> "." And FileName <> ".." And Not isHiddenFile Then  'ignores current and encompassing directories and hidden files
                i = i + 1                                                      'increments file counter
                If ShowContent Then
                    'FileType = GetFileTypeAndImage(FileName)                            'returns file type and get image from system if not already loaded in imgFileTypes
                    FileType = GetFileTypeAndImage(fi.FullName)                          'returns file type and get image from system if not already loaded in imgFileTypes

                    itmX = lvOut.Items.Add(System.IO.Path.GetFileNameWithoutExtension(FileName))
                    itmX.ImageKey = FileType
                    itmX.SubItems.Add(FileType)                                         'file type
                    itmX.SubItems.Add(Format(fi.Length / 1000, "### ###"))              'file size
                    itmX.SubItems.Add(Format(fi.LastWriteTime, "d MMM yyyy  HH:mm"))    'modification date
                End If
            End If
        Next
        lvOut.Columns(0).Text = i & " " & mMessage(82)                         'displays file/document counter update

        tcFolders.TabPages(CurTabIndex).Text = SetUCaseAtFirst(dirInfo.Name) & " (" & i & " " & mMessage(72) & ")"   '... (n doc. to file...
        tcFolders.TabPages(CurTabIndex).ToolTipText = mZaanImportPath

        If ShowContent Then
            Call SortLvOutColumn()                                             'refreshs "out" files display with last column sorting
        End If
    End Sub

    Private Sub DisplayTempFiles()
        'Displays "temp" files to be used outside Zaan data "in" file list.
        Dim dirInfo As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim PathFile, FilePattern, FileName, FileType As String
        Dim i As Long
        Dim itmX As ListViewItem
        Dim isHiddenFile As Boolean

        'Debug.Print("=> DisplayTempFiles...")       'TEST/DEBUG
        PathFile = mZaanCopyPath
        FilePattern = "*.*"
        dirInfo = My.Computer.FileSystem.GetDirectoryInfo(PathFile)
        i = 0                                                        'initializes file/document counter

        lvOut.ListViewItemSorter = Nothing                           'IMPORTANT : ListView2ItemComparer or ListView2ItemRevComparer may have been used for LvTemp ! 
        lvTemp.ListViewItemSorter = Nothing                          'IMPORTANT : resets the listview item sorter property for saving time (if formerly allocated to a ListView2ItemComparer or ListView2ItemRevComparer object) 
        lvTemp.Items.Clear()

        For Each fi In dirInfo.GetFiles(FilePattern)
            FileName = fi.Name
            isHiddenFile = False
            If (fi.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then isHiddenFile = True
            If Microsoft.VisualBasic.Left(FileName, 1) = "~" Then isHiddenFile = True

            If FileName <> "." And FileName <> ".." And Not isHiddenFile Then       'ignores current and encompassing directories and hidden files
                i = i + 1                                            'increments file counter
                'FileType = GetFileTypeAndImage(FileName)             'returns file type and get image from system if not already loaded in imgFileTypes
                FileType = GetFileTypeAndImage(fi.FullName)          'returns file type and get image from system if not already loaded in imgFileTypes

                itmX = lvTemp.Items.Add(System.IO.Path.GetFileNameWithoutExtension(FileName))
                itmX.ImageKey = FileType
                itmX.SubItems.Add(FileType)                                         'file type
                itmX.SubItems.Add(Format(fi.Length / 1000, "### ###"))              'file size
                itmX.SubItems.Add(Format(fi.LastWriteTime, "d MMM yyyy  HH:mm"))    'modification date
            End If
        Next
        lvTemp.Columns(0).Text = i & " " & mMessage(82)                             'displays file/document counter update
        tcFolders.TabPages(2).Text = SetUCaseAtFirst(dirInfo.Name) & " (" & i & " " & mMessage(73) & ")"     '... (n doc. to delete)
        'tsmiLvTempSend.Enabled = False
        'If i > 0 Then tsmiLvTempSend.Enabled = True
        Call SortLvTempColumn()                                      'refreshs "temp" files display with last column sorting
    End Sub

    Private Sub InitFileInputWatcher()
        'Initializes input file watcher to detect updates of related directories.

        'Debug.Print("InitFileInputWatcher :  mInputPath = " & mInputPath)   'TEST/DEBUG
        If Not My.Computer.FileSystem.DirectoryExists(mInputPath) Then
            mInputPath = My.Computer.FileSystem.SpecialDirectories.MyDocuments      'goes back to initial setting
            mImportPath = mInputPath
        End If
        fswInput.EnableRaisingEvents = False
        'fswInput.Path = mInputPath
        fswInput.Path = mImportPath

        fswInput.EnableRaisingEvents = True
    End Sub

    Private Sub InitFileZaanWatcher()
        'Initializes ZAAN file watcher to detect ZAAN\copy and ZAAN\import file updates

        'Debug.Print("InitFileZaanWatcher :  mZaanAppliPath = " & mZaanAppliPath)   'TEST/DEBUG
        fswZaan.EnableRaisingEvents = False
        fswZaan.Path = mZaanAppliPath
        fswZaan.EnableRaisingEvents = True
    End Sub

    Private Sub InitFileDataWatcher()
        'Initializes "data","tree" and "xin" file watchers to detect directory access/change in current ZAAN database

        'Debug.Print("InitFileDataWatcher:  mZaanDbPath = " & mZaanDbPath)   'TEST/DEBUG
        fswData.EnableRaisingEvents = False
        fswData.Path = mZaanDbPath & "data\"
        fswData.InternalBufferSize = 65536                 'sets buffer size to maxi allowed : 64 kb (is 8 kb by default)
        fswData.EnableRaisingEvents = True

        fswTree.EnableRaisingEvents = False
        fswTree.Path = mZaanDbPath & "tree\"
        fswTree.EnableRaisingEvents = True

        If mXmode = "auto" Then
            fswXin.EnableRaisingEvents = False
            fswXin.Path = mZaanDbPath & "xin\"
            'fswXin.EnableRaisingEvents = True             'RETIRED AUTOMATIC CUBE IMPORT (SINCE 4 JULY 2011) !
        End If
    End Sub

    Private Sub LoadZaanDbImage()
        'Displays selected Zaan database image if ZAAN-PRO mode selected and if any "top_left_logo" image available
        Dim dirInfo As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim PathFile, FilePattern, FileName, FileType As String

        pctZaanLogo.Image = pctZaanLogo.InitialImage       'resets Zaan logo image

        PathFile = mZaanDbPath & "info\"
        'Debug.Print("LoadZaanDbImage : PathFile = " & PathFile)      'TEST/DEBUG
        FilePattern = "top_left_logo.*"
        Try
            dirInfo = My.Computer.FileSystem.GetDirectoryInfo(PathFile)
            For Each fi In dirInfo.GetFiles(FilePattern)
                FileName = fi.Name
                If FileName <> "." And FileName <> ".." Then              'ignores current and encompassing directories
                    FileType = System.IO.Path.GetExtension(FileName)      'get file extension/type
                    If (FileType = ".bmp") Or (FileType = ".gif") Or (FileType = ".png") Or (FileType = ".jpg") Then
                        pctZaanLogo.Load(PathFile & FileName)             'updates top left "home" image
                        Exit For
                    End If
                End If
            Next
        Catch ex As Exception
            Debug.Print(ex.Message)              'TEST/DEBUG : info directory doesn't exist !
        End Try
    End Sub

    Private Function GetLastDirName(ByVal zDbPath As String) As String
        'Extracts last directory name from given directory path name
        Dim DirTab(), ZaanDir As String

        ZaanDir = ""
        DirTab = Split(zDbPath, "\")
        If DirTab.Length > 1 Then
            ZaanDir = DirTab(DirTab.Length - 2)
        End If
        GetLastDirName = ZaanDir
    End Function

    Private Sub UpdateZaanDbRootPathName(ByVal DbPath As String)
        'Updates current database selection with mZaanDbRoot, mZaanDbPath and mZaanDbName using given DbPath

        If Microsoft.VisualBasic.Right(DbPath, 1) = "\" Then         'case of DbPath ends up with a "\" (is normal folder path notation)
            mZaanDbPath = DbPath
        Else                                                         'case od DbPath does not end up with a "\"
            mZaanDbPath = DbPath & "\"
        End If
        mZaanDbName = GetLastDirName(mZaanDbPath)                    'updates database name
        mZaanDbRoot = Mid(mZaanDbPath, 1, Len(mZaanDbPath) - Len(mZaanDbName) - 2)
        If UCase(Mid(DbPath, 1, 2)) = "C:" Then
            mIsLocalDatabase = True
        Else
            mIsLocalDatabase = False
        End If
        'Debug.Print("UpdateZaanDbRootPathName :  mZaanDbPath = " & mZaanDbPath)    'TEST/DEBUG
        'Debug.Print("1 => UpdateZaanDbRootPathName :  mZaanDbName = " & mZaanDbName)    'TEST/DEBUG
        'Debug.Print("2 => UpdateZaanDbRootPathName :  mZaanDbRoot = " & mZaanDbRoot)    'TEST/DEBUG
        'Debug.Print("3 => UpdateZaanDbRootPathName :  mIsLocalDatabase = " & mIsLocalDatabase)    'TEST/DEBUG
    End Sub

    Private Sub InitBrowserPanel()
        'Initializes browser display items (zoom view)

        'Debug.Print("InitBrowserPanel...")   'TEST/DEBUG

        lblDocName.Text = ""
        mIsSlideShowPaused = True
        pnlSlideControl.Visible = False
        mSlideShowIndex = -1                               'disables slide show
        pctZoom.Image = Nothing                            'clears image
        'VLC2Zoom.playlist.items.clear()                    'clears VLC playlist (and stops any running video)
    End Sub

    Public Sub InitZaanFormTitle(Optional ByVal TitleInfo As String = "")
        'Initializes ZAAN main form title

        'Me.Text = "ZAAN " & mLicTypeText & " - " & mZaanDbName & mZaanTitleOption   'updates main ZAAN form title with database name and eventual option
        Me.Text = mZaanDbName & mZaanTitleOption           'updates main ZAAN form title with database name and eventual option
        If TitleInfo <> "" Then
            Me.Text = Me.Text & " : " & TitleInfo          'appends title info to ZAAN form title
        End If
    End Sub

    Private Sub InitTreeInOutFiles(Optional ByVal OutTempToDisplay As Boolean = True)
        'Displays selected Zaan database name, logo, tree, selector and selected files and, by default, "imput", "out" and "temp" lists

        'Debug.Print("InitTreeInOutFiles...")    'TEST/DEBUG
        lstPage.Items.Clear()                              'clears page navigation pile

        Call InitZaanFormTitle()                           'initializes ZAAN main form title
        tbSearch.Text = ""                                 'resets node search string
        Call InitlvInOrder(10, False)                      'initializes mlvInOrder() table to date index column in descending order (details view by default)

        Call LoadZaanDbImage()                             'if ZAAN-PRO mode selected, displays Zaan db "top_left_logo" image if any
        Call LoadTrees()                                   'loads all trees of current ZAAN database

        Call InitBookmarkSelection()                       'initializes bookmark list with current (mFileFilter) selection and displays selector and selected files

        Call InitListsColOrderWidth()                      'initializes LvIn columns order/width and other Lists column widths

        If OutTempToDisplay Then
            Call LoadInputTree()                           'loads input tree with Windows directory structure starting at selected root

            Call DisplayOutFiles(True, 1)                  'displays tab 1 title and related "out" files
            Call DisplayOutFiles(True, 0)                  'displays tab 0 title and related "out" files

            Call DisplayTempFiles()                        'displays all "temp" files (to be used outside Zaan data "in" files)
        End If

        If Not SplitContainer3.Panel2Collapsed Then        'case of visible viewer panel
            Call InitBrowserPanel()                        'initializes browser display items
        End If
    End Sub

    Private Sub AddZaanPathInTabs()
        'Adds mZaanDbPath directory in database/cube tabs if not already listed and updates its selection
        Dim i, TabCubeIndex As Integer

        TabCubeIndex = -1
        For i = 0 To tcCubes.TabPages.Count - 1                      'searches if mZaanDbPath is already listed in tabs
            If tcCubes.TabPages(i).ToolTipText = mZaanDbPath Then    'directory is already listed
                TabCubeIndex = i                                     'get related tab index
                Exit For
            End If
        Next
        If TabCubeIndex < 0 Then                                     'related tab does not exist
            tcCubes.TabPages.Add(mZaanDbName)                        'adds directory as a new tab page
            TabCubeIndex = tcCubes.TabPages.Count - 1                'sets tab index to last position
            tcCubes.TabPages(TabCubeIndex).ToolTipText = mZaanDbPath
            tcCubes.TabPages(TabCubeIndex).ImageKey = "_x_" & mImageStyle
        End If
        tcCubes.SelectedIndex = TabCubeIndex                         'activates related tab page

        For i = 0 To tcCubes.TabPages.Count - 1                      'searches for eventual "..." tab page to be deleted (set at design time !)
            If tcCubes.TabPages(i).Text = "..." Then
                tcCubes.TabPages.RemoveAt(i)                         'removes initial tab page
                Exit For
            End If
        Next
    End Sub

    Private Sub DeleteZaanDbFromTabs()
        'Deletes currently selected ZAAN database from database/cube tabs and selects first db from list or ZAAN-Demo1 if empty
        tcCubes.TabPages.RemoveAt(tcCubes.SelectedIndex)             'removes selected item
        If tcCubes.TabPages.Count > 0 Then
            tcCubes.SelectedIndex = 0                                'resets selection to first item if any
        End If
    End Sub

    Private Sub ChangeZaanDb(ByVal DirPathName As String)
        'Changes ZAAN database to currently selected path name

        If mZaanDbPath <> DirPathName Then
            'Debug.Print("ChangeZaanDb :  DirPathName = " & DirPathName)    'TEST/DEBUG

            If My.Computer.FileSystem.DirectoryExists(DirPathName) Then      'checks than directory still exists or is still acessible
                Me.Cursor = Cursors.WaitCursor                       'sets wait cursor
                Call BackupTreeAndCleanData()                        'backups tree into a data\zz_backup_tree_*.txt file and removes eventual remaining unused thumbnail images and empty directories
                Call SaveFileFilterIni()                             'saves mFileFilter of current ZAAN database in db\info\zaan.ini
                Me.Cursor = Cursors.Default                          'resets default cursor

                Call UpdateZaanDbRootPathName(DirPathName)           'updates current database selection with mZaanDbRoot, mZaanDbPath and mZaanDbName
                Call AddZaanPathInTabs()                             'adds mZaanDbPath directory in database/cube tabs if not already listed and updates its selection

                Call GetFileFilterIni()                              'initializes mFileFilter of current ZAAN database from db\info\zaan.ini
                Call InitFileDataWatcher()                           'initializes "data","tree" and "xin" file watchers to detect directory access/change in current ZAAN database
                Call InitTreeInOutFiles(False)                       'displays selected Zaan database name, logo, tree, selector and selected files views, but not "imput", "out" and "temp" views
            Else
                MsgBox(mMessage(122), MsgBoxStyle.Exclamation)       'Sorry, this ZAAN database is not more accessible !
                Call DeleteZaanDbFromTabs()                          'deletes currently selected ZAAN database from tabs and selects first db from list or ZAAN-Demo1 if empty
            End If
        End If
    End Sub

    Private Sub UpdateImportPath(ByVal DirPathName As String)
        'If given directory exists, updates mZaanImportPath and related "out" files display.

        'Debug.Print("UpdateImportPath :  mZaanImportPath = " & mZaanImportPath & "  DirPathName = " & DirPathName)   'TEST/DEBUG
        If My.Computer.FileSystem.DirectoryExists(DirPathName) Then
            mZaanImportPath = DirPathName
            Call DisplayOutFiles()                                       'displays all "out" files
        End If
    End Sub

    Private Sub SelectInputDir()
        'Selects input directory to be used for importing included files

        fbdZaanPath.SelectedPath = mInputPath
        fbdZaanPath.Description = mMessage(126)                           'Select a folder
        fbdZaanPath.ShowNewFolderButton = False                           'disables new folder button display
        If fbdZaanPath.ShowDialog = Windows.Forms.DialogResult.OK Then
            mInputPath = fbdZaanPath.SelectedPath & "\"

            mImportPath = mInputPath
            Call DisplayOutFiles(False, 1)                                'displays tab 1 title only 

            Call InitFileInputWatcher()                                   'initializes input file watcher to detect updates of related directories
            Call LoadInputTree()                                          'loads input tree with Windows directory structure starting at selected root
            Call UpdateImportPath(mImportPath)                            'if given directory exists, updates mZaanImportPath and related tooltip, file watcher and "out" files display
        End If
    End Sub

    Private Function CreateTreeFilesMissingMonths(ByVal YearKey As String, ByVal FullYearText As String) As Integer
        'Creates in current ZAAN database tree files corresponding to missing months of given year and returns the number of months created
        Dim i, n As Integer
        Dim ShortYearText, MonthKey, MonthText, TreeText As String

        ShortYearText = Mid(FullYearText, 3, 2)                                     'get short year text like "10" (for 2010)
        n = 0
        For i = 1 To 12                                                             'adds related 12 months (like "2010-01", "2010-02"...) as children of given year
            MonthText = Format(i, "00")
            If mTreeKeyLength = 13 Then                                             'case of 13 char. node key length 
                MonthKey = "_t" & ShortYearText & MonthText & "00000000"            '=> creates a YYMM00 year-month key with no valid day
            Else                                                                    'case of 16 char. node key length 
                MonthKey = "_t" & FullYearText & MonthText & "000000000"            '=> creates a YYMM00 year-month key with no valid day
            End If

            TreeText = FullYearText & "-" & MonthText
            If Not TreeFileTextExists("t", TreeText) Then                           'case of not existing When tree file with same TreeText
                n = n + 1                                                           'increments counter of created months
                Call CreateZaanPathTreeFile(MonthKey & YearKey & "__" & TreeText)   'adds month tree file 
            End If
        Next
        CreateTreeFilesMissingMonths = n                                            'returns the number of months created
    End Function

    Private Function CreateTreeFileYear(ByVal CurDate As DateTime) As String
        'Creates in current ZAAN database (v3 format) one year tree file including given date
        Dim CurYear As String = Format(CurDate, "yyyy")
        Dim CurYearCode As String = GetWhenV3KeyFromDateText(CurYear)
        Dim FileName As String = "_t" & CurYearCode & "_t" & mTreeRootKey & "__" & CurYear

        Debug.Print("CreateTreeFileYear :  FileName = " & FileName)    'TEST/DEBUG
        Call CreateZaanPathTreeFile(FileName)                        'adds year tree file
        CreateTreeFileYear = CurYear
    End Function

    Public Sub CleanDataDir(ByVal DataPath As String)
        'Removes eventual remaining unused thumbnail images and empty directories from given data path
        Dim dirInfo, dirSel As System.IO.DirectoryInfo

        'Debug.Print("CleanDataDir : " & DataPath)    'TEST/DEBUG
        If Not My.Computer.FileSystem.DirectoryExists(DataPath) Then Exit Sub 'case of directory doesn't exist

        dirInfo = My.Computer.FileSystem.GetDirectoryInfo(DataPath)
        For Each dirSel In dirInfo.GetDirectories("_*")              'scans all existing data directories
            Call CleanDataOldFileImages(dirSel.FullName)             'deletes any remaining old thumbnail file images no more associated to source image files
            Call DeleteDirIfEmpty(dirSel.FullName)                   'deletes given directory if empty
        Next
    End Sub

    Private Sub BackupTreeAndCleanData()
        'Backups tree into one text file in db/info folder and removes eventual remaining unused thumbnail images and empty directories from current database
        'Dim dirInfo, dirSel As System.IO.DirectoryInfo
        'Dim SelPattern As String
        Dim NodeX As TreeNode

        fswData.EnableRaisingEvents = False                          'locks fswData related events

        Call CleanDataDir(mZaanDbPath & "data")                      'removes eventual remaining unused thumbnail images and empty directories from given data path
        If mLicTypeCode >= 30 Then
            For Each NodeX In trvW.Nodes(0).Nodes
                Call CleanDataDir(Mid(NodeX.Tag, 3))                 'removes eventual remaining unused thumbnail images and empty directories from given data path
            Next
        End If

        If mTreeKeyLength > 0 Then
            Call BackupTree()                                        'saves tree source files into one text file in info folder of current database
        End If
        fswData.EnableRaisingEvents = True                           'unlocks fswData related events
    End Sub

    Private Sub DisplayExportedDatabase()
        'TEST for loading/reading an exported ZAAN database

        Debug.Print("DisplayExportedDatabase : " & mZaanDbPath)      'TEST/DEBUG
        Call LoadExportedTrees()                                     'Loads exported trees (into Windows directories) : Who (o), What (a), When (t), Where (e) and What else (b) and Other (c) in ZAAN-Pro mode

    End Sub

    Private Sub ResizeAndSaveImageFile(ByVal SceFilePathName As String, ByVal DestFilePathName As String, ByVal NewWidth As Integer, ByVal NewHeight As Integer)
        'Resizes and saves given image file into given destination file name within given new frame limits (NewWidth x NewHeight)
        Dim FileType As String
        Dim SceSizeRatio, DestSizeRatio As Single
        Dim srcRect, destRect As Rectangle
        Dim image As New Bitmap(SceFilePathName)                     'loads source image (full size)
        Dim g As Graphics

        srcRect.Width = image.Width                                  'sets size and location of source rectangle (source image)
        srcRect.Height = image.Height
        srcRect.X = 0
        srcRect.Y = 0

        SceSizeRatio = image.Width / image.Height
        DestSizeRatio = NewWidth / NewHeight

        If DestSizeRatio > SceSizeRatio Then                         'sets maximum size and location of destination rectangle, within NewWidth x NewHeight frame
            destRect.Width = NewHeight * image.Width / image.Height
            destRect.Height = NewHeight
            'destRect.X = (NewWidth - destRect.Width) / 2
            destRect.X = 0
            destRect.Y = 0
        Else
            destRect.Width = NewWidth
            destRect.Height = NewWidth * image.Height / image.Width
            destRect.X = 0
            'destRect.Y = (NewHeight - destRect.Height) / 2
            destRect.Y = 0
        End If
        Dim bm As New Bitmap(destRect.Width, destRect.Height)        'sets new frame bitmap

        g = Graphics.FromImage(bm)
        g.DrawImage(image, destRect, srcRect, GraphicsUnit.Pixel)    'draws source image at given new size and location

        FileType = LCase(System.IO.Path.GetExtension(SceFilePathName))
        Select Case FileType
            Case ".bmp"
                bm.Save(DestFilePathName, System.Drawing.Imaging.ImageFormat.Bmp)        'saves resized image in bmp format
            Case ".gif"
                bm.Save(DestFilePathName, System.Drawing.Imaging.ImageFormat.Gif)        'saves resized image in gif format
            Case ".jpg"
                bm.Save(DestFilePathName, System.Drawing.Imaging.ImageFormat.Jpeg)       'saves resized image in jpeg format
            Case ".png"
                bm.Save(DestFilePathName, System.Drawing.Imaging.ImageFormat.Png)        'saves resized image in png format
            Case Else
                bm.Save(DestFilePathName)                                                'saves resized image in source format
        End Select

        image.Dispose()                                              'releases image resources
        bm.Dispose()                                                 'releases bitmap resources
    End Sub

    Private Sub SelectZaanDbImage()
        'Selects an image file (.bmp, .gif, .jpg or .png) for associating it to current ZAAN database
        Dim SceFilePathName, DestFilePathName As String
        Dim OldFileName As String

        ofdZaanFile.InitialDirectory = mZaanDbPath & "info\"
        ofdZaanFile.FileName = ""
        ofdZaanFile.Filter = "Images (*.bmp;*.gif;*.jpg;*.png)|*.bmp;*.gif;*.jpg;*.png"
        ofdZaanFile.Title = mMessage(201)                                      'ZAAN database : select an image

        If ofdZaanFile.ShowDialog = Windows.Forms.DialogResult.OK Then         'OpenFileDialog => a file image has been selected
            SceFilePathName = ofdZaanFile.FileName
            'Debug.Print("SelectZaanDbImage :  file selected = " & SceFilePathName)    'TEST/DEBUG

            DestFilePathName = mZaanDbPath & "info\top_left_logo" & System.IO.Path.GetExtension(SceFilePathName)

            Dim dirSel As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "info\")
            Dim fi As System.IO.FileInfo
            For Each fi In dirSel.GetFiles("top_left_logo*.*")                  'searches for existing "top_left_logo*" file(s)
                Try
                    'OldFileName = "top_left_logo_OLD" & System.IO.Path.GetExtension(fi.Name)
                    OldFileName = System.IO.Path.GetFileNameWithoutExtension(fi.Name) & "I" & System.IO.Path.GetExtension(fi.Name)
                    My.Computer.FileSystem.RenameFile(fi.FullName, OldFileName)   'renames existing file(s) as "top_left_logo_OLD"
                Catch ex As Exception
                    Debug.Print(ex.Message)            'TEST/DEBUG
                End Try
            Next
            'Call CopySelectedFile(SceFilePathName, "", DestFilePathName)       'copies source image file at current db\info\ directory and renames it to "top_left_logo"
            Call ResizeAndSaveImageFile(SceFilePathName, DestFilePathName, 150, 50)      'resizes and saves source image file into destination file within given frame

            Call LoadZaanDbImage()                                             'displays selected Zaan database image if ZAAN-PRO mode selected and if any "top_left_logo" image available
        End If
    End Sub

    Private Sub SelectZaanDb()
        'Selects a ZAAN database root directory and checks that info, tree and data sub-directories are existing.
        'If empty directory, proposes to user to create a new ZAAN database directory with info, tree and data sub-directories.
        'If not empty Windows directory selected (with no info, tree and data sub-directories) proposes a directory audit for new ZAAN db preparation.
        Dim DirPathName As String
        Dim DirInfoExists, DirTreeExists, DirDataExists As Boolean
        Dim DirExpWhoExists, DirExpWhatExists, DirExpWhenExists, DirExpWhereExists As Boolean
        Dim dirInfo, dirSel As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim DirIsEmpty As Boolean
        Dim Response As MsgBoxResult
        Dim FileName As String

        Me.Cursor = Cursors.WaitCursor                                    'sets wait cursor
        Call BackupTreeAndCleanData()                                     'backups tree into one text file in db/info folder and removes eventual remaining unused thumbnail images and empty directories from current database
        Call SaveFileFilterIni()                                          'saves mFileFilter of current ZAAN database in db\info\zaan.ini
        Me.Cursor = Cursors.Default                                       'resets default cursor

        fbdZaanPath.SelectedPath = mZaanDbPath
        fbdZaanPath.Description = mMessage(63) & vbCrLf & mMessage(224)   'Select or create a ZAAN database (Folder like : ZAAN-MyBase)
        fbdZaanPath.ShowNewFolderButton = True                            'enables new folder button display
        If fbdZaanPath.ShowDialog = Windows.Forms.DialogResult.OK Then
            DirPathName = fbdZaanPath.SelectedPath & "\"
            DirPathName = CheckBasicDocPath(DirPathName)                  'checks if given directory is within MyDocuments in ZAAN Basic mode and exists, else resets it to "ZAAN-Demo1" directory

            DirInfoExists = My.Computer.FileSystem.DirectoryExists(DirPathName & "info")
            DirTreeExists = My.Computer.FileSystem.DirectoryExists(DirPathName & "tree")
            DirDataExists = My.Computer.FileSystem.DirectoryExists(DirPathName & "data")
            If DirInfoExists And DirTreeExists And DirDataExists Then     'case of selected directory is a ZAAN database
                Call UpdateZaanDbRootPathName(DirPathName)                'updates current database selection with mZaanDbRoot, mZaanDbPath and mZaanDbName
                Call AddZaanPathInTabs()                                  'adds mZaanDbPath directory in database/cube tabs if not already listed and updates its selection

                Call GetFileFilterIni()                                   'initializes mFileFilter of current ZAAN database from db\info\zaan.ini
                Call InitFileDataWatcher()                                'initializes "data","tree" and "xin" file watchers to detect directory access/change in current ZAAN database
                Call InitTreeInOutFiles(False)                            'displays selected Zaan database name, logo, tree, selector and selected files views, but not "imput", "out" and "temp" views
            Else                                                          'case of selected directory is not a ZAAN database
                DirExpWhoExists = My.Computer.FileSystem.DirectoryExists(DirPathName & "1 - " & mMessage(1))
                DirExpWhatExists = My.Computer.FileSystem.DirectoryExists(DirPathName & "2 - " & mMessage(2))
                DirExpWhenExists = My.Computer.FileSystem.DirectoryExists(DirPathName & "3 - " & mMessage(3))
                DirExpWhereExists = My.Computer.FileSystem.DirectoryExists(DirPathName & "4 - " & mMessage(4))
                If DirExpWhoExists And DirExpWhatExists And DirExpWhenExists And DirExpWhereExists Then   'TEST/DEBUG
                    MsgBox(mMessage(162), MsgBoxStyle.Exclamation)                  'Sorry, this directory is an exported ZAAN database !
                Else
                    DirIsEmpty = True
                    dirInfo = My.Computer.FileSystem.GetDirectoryInfo(Microsoft.VisualBasic.Left(DirPathName, Len(DirPathName) - 1))
                    For Each fi In dirInfo.GetFiles("*.*")
                        FileName = fi.Name
                        If FileName <> "." And FileName <> ".." Then
                            DirIsEmpty = False                                      'source directory is not empty
                            Exit For
                        End If
                    Next
                    For Each dirSel In dirInfo.GetDirectories
                        If dirSel.Name <> "" Then
                            DirIsEmpty = False                                      'source directory is not empty
                            Exit For
                        End If
                    Next
                    If DirIsEmpty Then
                        Response = MsgBox(mMessage(62), MsgBoxStyle.YesNo + MsgBoxStyle.Question)  'this directory is empty : do you want to create a new ZAAN database ?
                        If Response = MsgBoxResult.Yes Then
                            DirInfoExists = CreateDirIfNotExistsOK(DirPathName & "info")      'creates info directory if not exists and returns true if exists or creation succeeded
                            DirTreeExists = CreateDirIfNotExistsOK(DirPathName & "tree")      'creates tree directory if not exists and returns true if exists or creation succeeded
                            DirDataExists = CreateDirIfNotExistsOK(DirPathName & "data")      'creates data directory if not exists and returns true if exists or creation succeeded
                            If DirInfoExists And DirTreeExists And DirDataExists Then         'if sub-directories creation succeeded
                                Call UpdateZaanDbRootPathName(DirPathName)          'updates current database selection with mZaanDbRoot, mZaanDbPath and mZaanDbName
                                Call AddZaanPathInTabs()                            'adds mZaanDbPath directory in database/cube tabs if not already listed and updates its selection

                                mFileFilter = ""
                                Call InitLvBookmark()                               'initializes bookmark list
                                mZaanExportDest = ""
                                mXmode = ""
                                tsmiSelectorAutoImport.CheckState = CheckState.Unchecked
                                Call InitFileDataWatcher()                          'initializes "data","tree" and "xin" file watchers to detect directory access/change in current ZAAN database
                                Call CreateZaanLogoIfNoTopLeftFile(DirPathName & "info")
                                Call CreateTreeFilesRootsAndYear(Now)               'creates in current ZAAN database (v3 format) 6 tree root files and 1 year/12 months at given year
                                Call InitTreeInOutFiles(False)                      'displays selected Zaan database name, logo, tree, selector and selected files views, but not "imput", "out" and "temp" views
                            End If
                            If CreateDirIfNotExistsOK(DirPathName & "xin") Then     'tries to create a db "xin" directory if it doesn't exist
                            End If
                        End If
                    Else
                        MsgBox(mMessage(67), MsgBoxStyle.Exclamation)               'selected directory is not a ZAAN database and is not empty for creating one !
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ResetFileMovePile()
        'Resets file move pile (used for "Undo move") and disables related local menu control
        mFileMovePile = ""                                 'resets mFileMovePile that is used for storing file moves and enabling eventual undo (moving back)
        tsmiLvInUndoMove.Enabled = False                   'disables related local menu control
        tsmiLvOutUndoMove.Enabled = False                  'disables related local menu control
        tsmiLvTempUndoMove.Enabled = False                 'disables related local menu control
    End Sub

    Private Sub AddFileMovePile(ByVal SourceFilePathName As String, ByVal DestFilePathName As String)
        'Adds to file move pile given source and destination file path names and enables 'Undo move" local menu control
        mFileMovePile = mFileMovePile & SourceFilePathName & ">" & DestFilePathName & vbCrLf
        tsmiLvInUndoMove.Enabled = True                    'enables related local menu control
        tsmiLvOutUndoMove.Enabled = True                   'enables related local menu control
        tsmiLvTempUndoMove.Enabled = True                  'enables related local menu control
        'Debug.Print(" => Pile = " & mFileMovePile)    'TEST/DEBUG
    End Sub

    Private Sub UndoFileMove()
        'Undo file moves stored in file move pile and resets it and disables related local menu control
        Dim MoveList As String() = Split(mFileMovePile, vbCrLf)
        Dim FileList As String()
        Dim MyFiles(MoveList.Length - 1) As String
        Dim LastDirPathName As String
        Dim i As Integer

        'Debug.Print("UndoFileMove :  Pile = " & mFileMovePile)    'TEST/DEBUG
        LastDirPathName = ""
        Call LocksFswDataInputZaan()                                           'sets wait cursor and locks fswData, fswInput and fswZaan (import and copy) related events
        For i = 0 To MoveList.Length - 1
            If MoveList(i) <> "" Then
                FileList = Split(MoveList(i), ">")                             'get stored SourceFilePathName and DestFilePathName
                If FileList.Length > 1 Then
                    'Debug.Print("UndoFileMove :  source = " & FileList(0) & "  destination = " & FileList(1))    'TEST/DEBUG
                    Call MoveSelectedFile(FileList(1), "", FileList(0), False, False)      'moves back file (and deletes related thumbnail images if any)
                End If
                MyFiles(i) = FileList(1)                                       'stores moved file in MyFile() table for later highlighting
                LastDirPathName = System.IO.Path.GetDirectoryName(FileList(0)) & "\"     'stores last destination directory
            End If
        Next
        Call ResetFileMovePile()                                               'resets file move pile (used for "Undo move") and disables related local menu control
        Call UnlocksFswDataInputZaan()                                         'unlocks fswZaan (import and copy), fswInput and fswData related events and resets default cursor
        Select Case LastDirPathName
            Case mZaanImportPath
                Call DisplayListsAndHighlightItems(MyFiles, lvOut)             'displays "in", "out" and "temp" lists and highlights moved files/items in destination list
            Case mZaanCopyPath
                Call DisplayListsAndHighlightItems(MyFiles, lvTemp)            'displays "in", "out" and "temp" lists and highlights moved files/items in destination list
            Case Else
                If mTreeKeyLength = 0 Then                                     'no node keys in v4 database format (use directly Windows path names)
                    'TO BE DONE
                Else
                    Dim DirItems() As String = Split(LastDirPathName, "\")
                    If DirItems.Length > 3 Then
                        Dim DataDir As String = DirItems(DirItems.Length - 2)
                        Call ReplaceHierarchicalKeysBySimpleKeys(DataDir)  'replaces in given data directory name hierarchical keys by simple keys
                        'Debug.Print("UndoFileMove :  DataDir = " & DataDir)    'TEST/DEBUG
                        mFileFilter = Replace(DataDir, "_", "*_")              'updates mFileFilter with last filter set in AutoFile mode
                        Call DisplaySelector()                                 'displays selector buttons using mFileFilter selections
                    End If
                    Call DisplayListsAndHighlightItems(MyFiles, lvIn)          'displays "in", "out" and "temp" lists and highlights moved files/items in destination list
                End If
        End Select
    End Sub

    Private Sub MoveSelectedFile(ByVal FileName As String, ByVal SourceDirPathName As String, ByVal DestFilePathName As String, ByVal MoveImage As Boolean, Optional ByVal UndoEnabled As Boolean = True)
        'Moves selected file from source directory to destination dir./file and move related thumbnail image file if any and if MoveImage is true
        Dim dirSel As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim SourceFilePathName, SourceDirectory, DestDirectory, SourceFileName, ThumbFileName As String

        If SourceDirPathName = "" Then
            SourceFilePathName = FileName
        Else
            SourceFilePathName = SourceDirPathName & "\" & FileName
        End If
        'Debug.Print("MoveSelectedFile : " & SourceFilePathName & " TO " & DestFilePathName)   'TEST/DEBUG

        If SourceFilePathName = DestFilePathName Then Exit Sub 'moving a file to itself is not possible !

        SourceDirectory = System.IO.Path.GetDirectoryName(SourceFilePathName)
        SourceFileName = System.IO.Path.GetFileName(SourceFilePathName)
        DestDirectory = System.IO.Path.GetDirectoryName(DestFilePathName)

        ThumbFileName = ""
        dirSel = My.Computer.FileSystem.GetDirectoryInfo(SourceDirectory)
        For Each fi In dirSel.GetFiles("zzi*" & SourceFileName)  'searches for corresponding thumbnail image file...
            ThumbFileName = fi.Name
            Exit For
        Next

        'Debug.Print("MoveSelectedFile : " & SourceFilePathName & " TO " & DestFilePathName)   'TEST/DEBUG
        Try
            My.Computer.FileSystem.MoveFile(SourceFilePathName, DestFilePathName, FileIO.UIOption.AllDialogs, FileIO.UICancelOption.DoNothing)
            If UndoEnabled Then
                Call AddFileMovePile(SourceFilePathName, DestFilePathName)     'adds to file move pile given source and destination file path names and enables 'Undo move" local menu control
            End If
            If ThumbFileName <> "" Then                              'case of an existing thumbnail image file related to FileName
                If MoveImage Then
                    My.Computer.FileSystem.MoveFile(SourceDirectory & "\" & ThumbFileName, DestDirectory & "\" & ThumbFileName, FileIO.UIOption.AllDialogs, FileIO.UICancelOption.DoNothing)
                Else
                    My.Computer.FileSystem.DeleteFile(SourceDirectory & "\" & ThumbFileName)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)              'displays error message
        End Try
    End Sub

    Private Sub CopySelectedFile(ByVal FileName As String, ByVal SourceDirPathName As String, ByVal DestFilePathName As String)
        'Copies selected file from source directory to destination dir./file
        Dim SourceFilePathName As String

        If SourceDirPathName = "" Then
            SourceFilePathName = FileName
        Else
            SourceFilePathName = SourceDirPathName & "\" & FileName
        End If

        If SourceFilePathName = DestFilePathName Then Exit Sub 'copy of a file on itself is not possible !

        'Debug.Print("CopySelectedFile : " & SourceFilePathName & " IN " & DestFilePathName)   'TEST/DEBUG
        Try
            My.Computer.FileSystem.CopyFile(SourceFilePathName, DestFilePathName, FileIO.UIOption.AllDialogs, FileIO.UICancelOption.DoNothing)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)              'displays error message
        End Try
    End Sub

    Private Sub DropSelectedFile(ByVal SourceFilePathName As String, ByVal DestFilePathName As String, ByVal DropEffect As Windows.DragDropEffects, Optional ByVal UndoEnabled As Boolean = True)
        'Drops given source file to destination depending on copy/move drop effect, else on disk units (equal => move, else copy)

        'Debug.Print("DropSelectedFile : " & SourceFilePathName & " ON " & DestFilePathName & "  DropEffect = " & DropEffect.ToString)   'TEST/DEBUG
        If SourceFilePathName = DestFilePathName Then Exit Sub 'copy of a file on itself is not possible !

        If DropEffect = Windows.DragDropEffects.Copy Then
            Call CopySelectedFile(SourceFilePathName, "", DestFilePathName)              'copies file to current selection
        ElseIf DropEffect = Windows.DragDropEffects.Move Then
            Call MoveSelectedFile(SourceFilePathName, "", DestFilePathName, False, UndoEnabled)       'moves file to current selection
        Else
            If Mid(SourceFilePathName, 1, 2) = Mid(DestFilePathName, 1, 2) Then          'case of same disk unit => move file
                Call MoveSelectedFile(SourceFilePathName, "", DestFilePathName, False, UndoEnabled)   'moves file to current selection
            Else
                Call CopySelectedFile(SourceFilePathName, "", DestFilePathName)          'copies file to current selection
            End If
        End If
    End Sub

    Private Function InTextWithNull(ByVal StartPos As Integer, ByRef SourceText As String, ByVal SearchedText As String) As Integer
        'Returns start index of given searched text, with inserted null characters after each character, in given source text
        Dim Z As Char = Chr(0)
        Dim SearchedTextWithNull As String = ""
        Dim i, p As Integer

        p = 0
        If SearchedText <> "" Then
            For i = 1 To Len(SearchedText)
                SearchedTextWithNull = SearchedTextWithNull & Mid(SearchedText, i, 1) & Z
            Next
            p = InStr(StartPos, SourceText, SearchedTextWithNull, CompareMethod.Text)
        End If
        InTextWithNull = p
    End Function

    Private Function TextWhithoutSubStgNorNullChars(ByVal EntryText As String) As String
        'Returns given entry text without substg (128 bytes blocs) nor null characters and no heading and trailing spaces
        Dim OutputText As String = ""
        Dim EntryStartText As String = ""
        Dim Z As Char = Chr(0)
        Dim C As Char
        Dim p, q, i As Integer

        p = 1
        Do                                                           'scans EntryText for eliminating eventual "__substg" blocs of 128 bytes length
            q = InTextWithNull(p, EntryText, "__substg")             'returns start index of given searched text, with inserted null characters after each character, in FileContent string
            If q > 0 Then
                If p = 1 Then
                    EntryStartText = Mid(EntryText, 1, q - 1)        'get string before first "__substg" bloc
                End If
                p = q + 128
            End If
        Loop Until q = 0
        EntryText = EntryStartText & Mid(EntryText, p)               'rebuilds EntryText without any "__substg" bloc

        For i = 1 To Len(EntryText)
            C = Mid(EntryText, i, 1)
            If C <> Z Then                                           'skips null characters
                OutputText = OutputText & C
            End If
        Next
        TextWhithoutSubStgNorNullChars = Trim(OutputText)
    End Function

    Private Sub UpdateFileLastWriteDate(ByVal FilePathName As String, ByVal LastWriteDate As Date)
        'Updates given file with given last write date
        Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FilePathName)

        'Debug.Print("UpdateFileLastWriteDate :  name = " & fi.Name & "  date = " & fi.LastWriteTime)   'TEST/DEBUG
        fi.LastWriteTime = LastWriteDate
    End Sub

    Private Sub DropOutlookFiles(ByRef eData, ByVal DestDirName, ByRef DestList)
        'Drops given Outlook files contained OutlookDataObject eData at destination directory
        Dim i, p, q, r, n As Integer
        Dim MsgFrom, MsgDate, FileNameText, FileType, DestFilePathName As String
        'Dim SearchedText As String
        Dim Z As Char = Chr(0)
        Dim MsgDateDate As Date

        If Microsoft.VisualBasic.Right(DestDirName, 1) <> "\" Then
            DestDirName = DestDirName & "\"
        End If
        'Debug.Print("DropOutlookFiles :  DestDirName = " & DestDirName)        'TEST/DEBUG

        Call LocksFswDataInputZaan()                                           'sets wait cursor and locks fswData, fswInput and fswZaan (import and copy) related events
        Call ResetFileMovePile()                                               'resets file move pile (used for "Undo move") and disables related local menu control

        Dim TempPath As String = Path.GetTempPath()                            'gets related temp directory path
        Dim dataObject As New OutlookDataObject(eData)                         'wraps standard IDataObject in OutlookDataObject
        Dim filenames As String() = DirectCast(dataObject.GetData("FileGroupDescriptor"), String())      'get the names and data streams of the files dropped
        Dim filestreams As MemoryStream() = DirectCast(dataObject.GetData("FileContents"), MemoryStream())

        For i = 0 To filenames.Length - 1                                      'scans names of files to be dropped
            Dim filestream As MemoryStream = filestreams(i)                    'gets related data stream
            Dim fullfilename As String = TempPath & filenames(i)
            Dim outputStream As FileStream = File.Create(fullfilename)
            filestream.WriteTo(outputStream)                                   'saves related file stream in temp directory
            outputStream.Close()                                               'and closes file stream

            MsgFrom = ""                                                       'initializes sender email to empty string
            MsgDate = ""                                                       'initializes email date to empty string
            Try
                Dim FileContent As String = My.Computer.FileSystem.ReadAllText(fullfilename)      'reads related temp file as a string content

                p = InStr(FileContent, "SMTP:", CompareMethod.Text)            'searches for message sender email address
                If p > 0 Then
                    MsgFrom = Mid(FileContent, p + 5, 59)                      'NOTE : 64 bytes max used for SMTP: address
                    q = InStr(MsgFrom, Z, CompareMethod.Binary)                'searches for first null byte/char
                    If q > 0 Then
                        MsgFrom = Mid(MsgFrom, 1, q - 1)                       'eliminates trailling null bytes
                    End If
                    MsgFrom = LCase(MsgFrom)
                End If

                'p = InTextWithNull(1, FileContent, "Return-Path:")             'returns start index of given searched text, with inserted null characters after each character, in FileContent string
                p = InTextWithNull(1, FileContent, "Received:")                'returns start index of given searched text, with inserted null characters after each character, in FileContent string
                If p > 0 Then
                    'q = InTextWithNull(p + 24, FileContent, "SMTP;")           'returns start index of given searched text, with inserted null characters after each character, in FileContent string
                    q = InTextWithNull(p + 18, FileContent, ";")               'returns start index of given searched text, with inserted null characters after each character, in FileContent string
                    If q > 0 Then
                        r = InTextWithNull(q + 2, FileContent, vbCrLf)         'returns start index of given searched text, with inserted null characters after each character, in FileContent string
                        If r > 0 Then
                            MsgDate = Mid(FileContent, q + 2, r - q - 2)
                            MsgDate = TextWhithoutSubStgNorNullChars(MsgDate)  'returns MsgDate without substg (128 bytes blocs) nor null characters and no heading and trailing spaces
                            If IsDate(MsgDate) Then
                                MsgDateDate = CDate(MsgDate)                   'converts date string to date format
                                MsgDate = Format(MsgDateDate, "yyyy-MM-dd HH:mm")
                            Else
                                MsgDate = ""
                            End If
                        End If
                    End If
                End If

            Catch ex As Exception
                'MsgBox(ex.Message, MsgBoxStyle.Exclamation)         'read file error !
                Debug.Print(ex.Message)                              'TEST/DEBUG error
            End Try
            'Debug.Print("DropOutlookFiles :  MsgFrom =>" & MsgFrom & "<=")    'TEST/DEBUG
            'Debug.Print("DropOutlookFiles :  MsgDate =>" & MsgDate & "<=")    'TEST/DEBUG

            FileNameText = System.IO.Path.GetFileNameWithoutExtension(filenames(i))
            FileType = System.IO.Path.GetExtension(filenames(i))
            If MsgFrom <> "" Then
                FileNameText = MsgFrom & " - " & FileNameText
            End If

            If MsgDate = "" Then
                DestFilePathName = DestDirName & FileNameText & " 0001" & FileType            'adds an 4 digits index at the end of the message file name
                n = 1
                Do While My.Computer.FileSystem.FileExists(DestFilePathName) And (n < 9999)   'tests if the file name is not already existing, else increments its index
                    n = n + 1
                    DestFilePathName = DestDirName & FileNameText & " " & Format(n, "0000") & FileType
                Loop
            Else
                DestFilePathName = DestDirName & Replace(MsgDate, ":", "h") & "_" & FileNameText & FileType      'adds date/time at the beginning of the message file name
            End If

            Call MoveSelectedFile(fullfilename, "", DestFilePathName, False, False)           'moves created file to destination directory

            If MsgDate <> "" Then
                Call UpdateFileLastWriteDate(DestFilePathName, MsgDateDate)    'updates last write date with message file (received) date
            End If
        Next
        Call UnlocksFswDataInputZaan()                                         'unlocks fswZaan (import and copy), fswInput and fswData related events and resets default cursor
        Call DisplayListsAndHighlightItems(filenames, DestList)                'displays "in", "out" and "temp" lists and highlights moved files/items in destination list
    End Sub

    Private Function GetSelectedDirPathName() As String
        'Returns current data directory pathname corresponding to current mFileFilter
        Dim TreeCode, DirName, DataAccessDirName, DataPath, OutputPathName As String
        Dim DirCodeTable() As String = Split(mFileFilter, "*")
        Dim i, TreePos As Integer

        'Debug.Print("GetSelectedDirPathName :  mFileFilter = " & mFileFilter)   'TEST/DEBUG
        DataAccessDirName = ""
        DirName = ""

        If mTreeKeyLength = 0 Then                                        'no node keys in v4 (use directly Windows path names)
            For i = 1 To DirCodeTable.Length - 1
                If Mid(DirCodeTable(i), 1, 1) = "z" Then
                    TreePos = Mid(DirCodeTable(i), 2, 1)
                    If TreePos = 0 Then                                   'is an Access group
                        DataAccessDirName = "_" & Mid(DirCodeTable(i), 3)
                    Else
                        DirName = DirName & "\" & DirCodeTable(i)
                    End If
                End If
            Next
            OutputPathName = mZaanDbPath & "data" & DataAccessDirName & DirName
        Else
            DataPath = mZaanDbPath & "data"                               'sets default data path
            For i = 1 To DirCodeTable.Length - 1
                TreeCode = Mid(DirCodeTable(i), 2, 1)
                If TreeCode = "u" Then
                    'DataAccessDirName = "data_" & Mid(DirCodeTable(i), 3) & "\"
                    DataPath = Mid(DirCodeTable(i), 3)                    'SINCE 2012-03-30 : full data path of Access Group is stored in mFileFilter
                Else
                    DirName = DirName & DirCodeTable(i)
                End If
            Next
            Dim p As Integer = InStr(DirName, "_t")                       'searches for When code
            If p = 0 Then                                                 'no When code found !
                Dim ThisYear As String = Format(Now, "yyyy")              'gets current year
                Dim ThisYearWhenKey As String = "t" & GetWhenV3KeyFromDateText(ThisYear)
                DirName = "_" & ThisYearWhenKey & DirName                 'appends current year code at the beginning of DirName
            End If
            Call ReplaceSimpleKeysByHierarchicalKeys(DirName)             'replaces in given data directory name simple keys by related hierarchical keys
            'OutputPathName = mZaanDbPath & "data\" & DataAccessDirName & DirName
            OutputPathName = DataPath & "\" & DirName
        End If
        'Debug.Print(" => GetSelectedDirPathName = " & OutputPathName)     'TEST/DEBUG
        GetSelectedDirPathName = OutputPathName
    End Function

    Private Function GetFromTextDimensionNodeKey(ByVal InputText As String, ByVal TreeCode As String) As String
        'Returns existing tree node key of first related dimension reference found in given input string
        Dim NodeKey, SelPattern As String
        Dim SubInputs As String() = Split(InputText, " - ")
        Dim SubSubInputs As String()
        Dim dirInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "tree")
        Dim fi As System.IO.FileInfo
        Dim i, j As Integer

        'Debug.Print("GetFromTextDNK : InputText = " & InputText & "  TreeCode = " & TreeCode) 'TEST/DEBUG
        NodeKey = ""
        For i = 0 To SubInputs.Length - 1                                                'scans all input words
            If SubInputs(i).Length > 1 Then                                              'skips separated characters
                'Debug.Print("=> SubInputs(" & i & ")= " & SubInputs(i)) 'TEST/DEBUG
                SubSubInputs = Split(Trim(SubInputs(i)), "_")                            'splits words containing "_" separator(s) into sub-words
                For j = 0 To SubSubInputs.Length - 1
                    If SubSubInputs(j).Length > 1 Then                                   'skips separated characters
                        'Debug.Print(" => SubSubInputs(" & j & ")= " & SubSubInputs(j))  'TEST/DEBUG
                        SelPattern = "_" & TreeCode & "*__" & Trim(SubSubInputs(j)) & ".txt"
                        For Each fi In dirInfo.GetFiles(SelPattern)                      'searches for given word in given dimension (TreeCode)
                            'Debug.Print("GetFromTextDimensionNodeKey : Tree file founded =>" & fi.Name & "<")        'TEST/DEBUG
                            NodeKey = Mid(fi.Name, 2, mTreeKeyLength)                    'get related full node key
                            Exit For
                        Next
                        If NodeKey <> "" Then Exit For
                    End If
                Next
                If NodeKey <> "" Then Exit For
            End If
        Next
        GetFromTextDimensionNodeKey = NodeKey
    End Function

    Private Function GetDataDirectoryFromInputFile(ByVal InputFilePathName As String, ByRef CurFileFilter As String) As String
        'Returns data directory related to first 4 dimension references found in given input file date/name and eventually update file date
        Dim DataDir As String = ""

        If My.Computer.FileSystem.FileExists(InputFilePathName) Then
            Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(InputFilePathName)
            Dim WhenText As String = Format(fi.LastWriteTime, "yyyy-MM-dd")    'get file modification date (such as 2010-09-01)
            Dim FileNameText As String = System.IO.Path.GetFileNameWithoutExtension(fi.Name)  'get file name
            Dim SubInputs As String() = Split(FileNameText, " - ")             'splits file name with " - " delimiter
            Dim SubSubInputs As String()
            Dim SearchText As String
            Dim i, j As Integer
            Dim WhenKey As String = ""
            Dim WhoKey As String = ""
            Dim WhatKey As String = ""
            Dim WhereKey As String = ""

            For i = 0 To SubInputs.Length - 1                                  'scans all input substrings
                'Debug.Print("=> SubInputs(" & i & ")= " & SubInputs(i))        'TEST/DEBUG
                SubSubInputs = Split(Trim(SubInputs(i)), "_")                  'splits substring with "_" delimiter
                For j = 0 To SubSubInputs.Length - 1
                    'Debug.Print(" => SubSubInputs(" & j & ")= " & SubSubInputs(j))   'TEST/DEBUG
                    SearchText = Trim(SubSubInputs(j))
                    If (i = 0) And (j = 0) And IsDate(Mid(SearchText, 1, 10)) Then     'checks if first substring starts with date like "2010-09-01"
                        WhenText = Mid(SearchText, 1, 10)                      'updates When text with first substring of file name
                        SearchText = Replace(SearchText, "h", ":")             'resets hour format if any time provided
                        If IsDate(SearchText) Then
                            fi.LastWriteTime = CDate(SearchText)               'updates date/time of current file
                        End If
                    ElseIf WhoKey = "" Then
                        WhoKey = GetNodeKeyOfTreeFileText("o", SearchText)
                    ElseIf WhatKey = "" Then
                        WhatKey = GetNodeKeyOfTreeFileText("a", SearchText)
                    ElseIf WhereKey = "" Then
                        WhereKey = GetNodeKeyOfTreeFileText("e", SearchText)
                    Else
                        Exit For
                    End If
                Next
            Next
            WhenKey = "t" & GetWhenV3KeyFromDateText(WhenText)

            DataDir = "_" & WhenKey & "_" & GetHierarchicalKey(WhoKey) & "_" & GetHierarchicalKey(WhatKey) & "_" & GetHierarchicalKey(WhereKey)
            CurFileFilter = "*_" & WhenKey & "*_" & WhoKey & "*_" & WhatKey & "*_" & WhereKey
        End If
        'Debug.Print("GetDataDirectoryFromInputFile :  InputFilePathName = " & InputFilePathName & " => DataDir = " & DataDir)   'TEST/DEBUG
        GetDataDirectoryFromInputFile = DataDir
    End Function

    Private Function IsZaanCubeFileName(ByVal FileName As String) As Boolean
        'Returns true if given filename is like "_zc_*.zip" or "_zc_*.xml"
        Dim FilePrefix As String = Mid(FileName, 1, 4)
        Dim FileType As String = LCase(System.IO.Path.GetExtension(FileName))
        Dim IsZaanCube As Boolean = False

        If FilePrefix = "_zc_" And ((FileType = ".zip") Or (FileType = ".xml")) Then
            IsZaanCube = True
        End If
        IsZaanCubeFileName = IsZaanCube
    End Function

    Private Sub ImportCubesFromListSelection(ByRef ImportList As ListView)
        'Imports selected cubes from given list view (lvOut or lvTemp) in current Zaan database (using its \xin directory)
        Dim FileName, FileNameText, InputFilePathName, DestFilePathName As String
        Dim MyFiles(ImportList.SelectedItems.Count - 1) As String
        Dim i As Integer
        Dim CurFileFilter, Keys(3) As String
        Dim Response As MsgBoxResult

        Response = MsgBox(mMessage(197), MsgBoxStyle.YesNo + MsgBoxStyle.Question)    'Do you confirm import of selected ZAAN cube(s) ?
        If Response = MsgBoxResult.Yes Then
            CurFileFilter = ""
            Call LocksFswDataInputZaan()                                       'sets wait cursor and locks fswData, fswInput and fswZaan (import and copy) related events
            Call ResetFileMovePile()                                           'resets file move pile (used for "Undo move") and disables related local menu control
            For i = 0 To ImportList.SelectedItems.Count - 1
                FileName = ImportList.SelectedItems(i).Text & ImportList.SelectedItems(i).SubItems(1).Text  'get document name & type
                MyFiles(i) = FileName                                          'stores file name in MyFiles() table for later highlighting
                If ImportList.Name = "lvTemp" Then                             'case of import from ZAAN\copy list
                    InputFilePathName = mZaanCopyPath & FileName
                Else                                                           'case of import from ZAAN\import list
                    InputFilePathName = mZaanImportPath & FileName
                End If
                FileNameText = System.IO.Path.GetFileNameWithoutExtension(FileName)

                If IsZaanCubeFileName(FileName) Then                           'case of a ZAAN cube file name (to be imported in database)
                    DestFilePathName = mZaanDbPath & "xin\" & FileName
                    Call MoveSelectedFile(InputFilePathName, "", DestFilePathName, False)     'moves file to "xin" folder for enabling cube import
                    tmrCubeImport.Enabled = True                               'starts timer to wait for 1 sec. before executing cube import
                End If
            Next
            If CurFileFilter <> "" Then
                mFileFilter = CurFileFilter                                    'updates mFileFilter with last filter set in AutoFile mode
                Call DisplaySelector()                                         'displays selector buttons using mFileFilter selections
                'Call DisplaySelectedFiles()                                    'displays all "in" files using mFileFilterTab table
            End If
            Call UnlocksFswDataInputZaan()                                     'unlocks fswZaan (import and copy), fswInput and fswData related events and resets default cursor
            Call DisplayListsAndHighlightItems(MyFiles, lvIn)                  'displays "in", "out" and "temp" lists and highlights moved files/items in destination list
        End If
    End Sub

    Private Sub MoveInAndFileOutSelection(ByVal AutoFile As Boolean)
        'Moves selected "out" files in current "in" list of current data directory if not AutoFile activated, else in related Who/When dimensions founded
        Dim FileName, FileNameText, InputFilePathName, DestFilePathName As String
        Dim MyFiles(lvOut.SelectedItems.Count - 1) As String
        Dim i As Integer
        'Dim fi As System.IO.FileInfo
        'Dim WhenText As String
        Dim DataDirKey, CurFileFilter, Keys(3) As String

        CurFileFilter = ""
        Call LocksFswDataInputZaan()                                           'sets wait cursor and locks fswData, fswInput and fswZaan (import and copy) related events
        Call ResetFileMovePile()                                               'resets file move pile (used for "Undo move") and disables related local menu control
        For i = 0 To lvOut.SelectedItems.Count - 1
            FileName = lvOut.SelectedItems(i).Text & lvOut.SelectedItems(i).SubItems(1).Text  'get document name & type
            MyFiles(i) = FileName                                              'stores file name in MyFiles() table for later highlighting
            InputFilePathName = mZaanImportPath & FileName
            FileNameText = System.IO.Path.GetFileNameWithoutExtension(FileName)

            If AutoFile Then
                'fi = My.Computer.FileSystem.GetFileInfo(InputFilePathName)
                'WhenText = Format(fi.LastWriteTime, "yyyy-MM-dd")              'get modification date such as 2010-09-01
                'Keys(0) = "t" & GetWhenV3KeyFromDateText(WhenText)             'returns When key in current v3/v3+ database format
                'Keys(1) = GetFromTextDimensionNodeKey(FileNameText, "o")       'returns existing tree node key of first Who dimension reference found in file name
                'Keys(2) = GetFromTextDimensionNodeKey(FileNameText, "a")       'returns existing tree node key of first What dimension reference found in file name
                'Keys(3) = GetFromTextDimensionNodeKey(FileNameText, "e")       'returns existing tree node key of first Where dimension reference found in file name
                'DataDirKey = ""
                'For j = 0 To Keys.Length - 1                                   'scans 4 keys table
                '  If Keys(j) <> "" Then
                '    DataDirKey = DataDirKey & "_" & GetHierarchicalKey(Keys(j))
                '  End If
                'Next
                DataDirKey = GetDataDirectoryFromInputFile(InputFilePathName, CurFileFilter)  'get data directory related to first 4 dimension references found in file date/name

                If DataDirKey = "" Then
                    MsgBox(mMessage(149), MsgBoxStyle.Information)             'No reference found in existing dimensions !
                Else
                    DestFilePathName = mZaanDbPath & "data\" & DataDirKey & "\" & FileName
                    'Debug.Print("AutoFile : DestFilePathName = " & DestFilePathName)    'TEST/DEBUG
                    Call MoveSelectedFile(InputFilePathName, "", DestFilePathName, False)    'moves file to current selection

                    'CurFileFilter = ""
                    'For j = 0 To Keys.Length - 1
                    '  If Keys(j) <> "" Then
                    '    CurFileFilter = CurFileFilter & "*_" & Keys(j)     'updates file filter (last one to be used for updating mFileFilter)
                    '  If
                    'Next
                End If
            Else
                DestFilePathName = GetSelectedDirPathName() & "\" & FileName
                Call MoveSelectedFile(InputFilePathName, "", DestFilePathName, False)    'moves file to current selection
            End If

        Next
        If CurFileFilter <> "" Then
            mFileFilter = CurFileFilter                                        'updates mFileFilter with last filter set in AutoFile mode
            Call DisplaySelector()                                             'displays selector buttons using mFileFilter selections
            'Call DisplaySelectedFiles()                                        'displays all "in" files using mFileFilterTab table
        End If
        Call UnlocksFswDataInputZaan()                                         'unlocks fswZaan (import and copy), fswInput and fswData related events and resets default cursor
        Call DisplayListsAndHighlightItems(MyFiles, lvIn)                      'displays "in", "out" and "temp" lists and highlights moved files/items in destination list
    End Sub

    Private Sub CheckFswInputChanges()
        'Checks Input/mImportPath changes for refreshing related display
        'Debug.Print("CheckFswInputChanges...")      'TEST/DEBUG

        If mZaanImportPath = mMyZaanImportPath Then                  'case of current import display is tab 0
            Call DisplayOutFiles(False, 1)                           'displays tab 1 title only 
            mZaanImportPath = mMyZaanImportPath                      'resets current import display to tab 0 (has been changed in DisplayOutFiles())
        Else                                                         'case of current import display is tab 1
            Call DisplayOutFiles(True, 1)                            'displays tab 1 title and related "out" files
        End If
    End Sub

    Private Sub CheckFswZaanChanges(ByVal Input As String)
        'Checks ZAAN\copy and ZAAN\import changes for refreshing related displays
        'Debug.Print("CheckFswZaanChanges :  Input = " & Input)     'TEST/DEBUG

        Input = LCase(Mid(Input, 1, 4))                              'get changed directory prefix of 4 characters
        Select Case Input
            Case "impo"
                If mZaanImportPath = mMyZaanImportPath Then          'case of current import display is tab 0
                    Call DisplayOutFiles(True, 0)                    'displays tab 0 title and related "out" files
                Else                                                 'case of current import display is tab 1
                    Call DisplayOutFiles(False, 0)                   'displays tab 0 title only
                    mZaanImportPath = mImportPath                    'resets current import display to tab 1 (has been changed in DisplayOutFiles())
                End If
            Case "copy"
                Call DisplayTempFiles()                              'refresh display of "temp"/"copy" files list
        End Select
    End Sub

    Private Sub InitTcFolders()
        'Initializes Windows folders tab control
        Dim FolderFileType As String = GetFileTypeAndImage("folder")

        tcFolders.TabPages(0).ImageKey = "folder"          'set folder image to Import folder
        tcFolders.TabPages(0).ToolTipText = mMyZaanImportPath
        tcFolders.TabPages(1).ImageKey = "folder"          'set folder image to Scan folder
        tcFolders.TabPages(1).ToolTipText = mImportPath
        tcFolders.TabPages(2).ImageKey = "folder"          'set folder image to Copy folder
        tcFolders.TabPages(2).ToolTipText = mZaanCopyPath
    End Sub

    Private Sub InitTodayButton()
        'Initializes today's date in today's button

        btnToday.Text = Format(Now, "yyyy-MM-dd")                         'initializes today date
        btnToday.Tag = "*_t" & GetWhenV3KeyFromDateText(btnToday.Text)    'sets related tag/When key
    End Sub

    Private Sub InitLocalDisplay()
        'Initializes display of labels, tooltips and list views in selected local language
        'Debug.Print("InitLocalDisplay...")    'TEST/DEBUG
        Call InitlvIn()                                    'initializes "in" list

        Call InitTcFolders()                               'initializes Windows folders tab control
        Call InitlvOut()                                   'initializes "out" list
        Call InitlvTemp()                                  'initializes "temp" list

        Call InitContextMenus()                            'initializes context menus in given language
        Call InitToolTips()                                'loads control tool tips in given language
        Call InitTreeInOutFiles()                          'displays selected Zaan database name, logo, tree, selector and selected files and, by default, "imput", "out" and "temp" views

        Call InitTodayButton()                             'initializes today's date in today's button
    End Sub

    Private Sub ExportSubTreesToWin(ByVal ParentNode As TreeNode, ByVal ParentDirPathName As String)
        'Exports given sub-trees of ZAAN database to a Windows folders tree
        Dim NodeX As TreeNode
        Dim DirPathName As String

        'Debug.Print("ExportDbToWinTree :  ParentNode = " & ParentNode.Text)          'TEST/DEBUG
        For Each NodeX In ParentNode.Nodes
            DirPathName = ParentDirPathName & "\" & NodeX.Text
            If CreateDirIfNotExistsOK(DirPathName) Then              'tries to create a sub-directory if not exists
                Call ExportSubTreesToWin(NodeX, DirPathName)         'recursive call to child node
            End If
        Next
    End Sub

    Private Sub ExportTreesToWindows(ByVal ExportPathName As String, ByVal GetGroups As Boolean)
        'Exports trees of current ZAAN database to a Windows directory tree organization at given export path name
        Dim TreeCodes() As String = Split(mTreeCodeSeries, " ")
        Dim NodeX, NodeY As TreeNode
        Dim ExpDirPathName, TreeCode, Title As String
        Dim NodesNb, n, i As Integer

        'Debug.Print("ExportTreesToWindows at : " & ExportPathName)   'TEST/DEBUG
        If CreateDirIfNotExistsOK(ExportPathName) Then                         'tries to create a "...exp*" directory if it doesn't exist
            NodesNb = trvW.Nodes.Count
            n = 0
            For Each NodeX In trvW.Nodes
                n = n + 1
                pgbZaan.Value = 10 * n / NodesNb
                TreeCode = Mid(NodeX.Tag, 2, 1)
                If TreeCode = "u" Then                                         'case of an Access control root node
                    If GetGroups Then
                        For Each NodeY In NodeX.Nodes                          'scans all existing groups (user should all directories access right to see them all)
                            ExpDirPathName = ExportPathName & "\data_" & NodeY.Text
                            If CreateDirIfNotExistsOK(ExpDirPathName) Then     'tries to create a "...exp*\_<group>" directory if not exists
                                If mTreeCodeSeries = "" Then
                                    Call ExportTreesToWindows(ExpDirPathName, False)                    'loads whole directory tree in current group directory
                                    Call ExportDocCopiesToWindows(ExpDirPathName, Mid(NodeY.Tag, 3))    'exports document copies of current ZAAN database to directories related to their first dimension available
                                Else
                                    Call ExportDataToWindows(ExpDirPathName, Mid(NodeY.Tag, 3))         'exports/copies files from given data path into a Windows mixted tree organisation depending on given tree codes series
                                End If
                            End If
                        Next
                    End If
                Else                                                           'case of a Who/What/When/Where... root node
                    Title = NodeX.Text
                    For i = 0 To TreeCodes.Length - 1
                        If TreeCodes(i) = TreeCode Then
                            Title = i + 1 & " - " & Mid(NodeX.Text, 5)         'sets requested order index to current node
                            Exit For
                        End If
                    Next
                    If mTreeCodeSeries = "" Then
                        ExpDirPathName = ExportPathName & "\" & Title
                    Else
                        ExpDirPathName = ExportPathName & "\tree\" & Title
                    End If
                    If CreateDirIfNotExistsOK(ExpDirPathName) Then             'tries to create a "...exp*\<Title>" directory if not exists
                        Call ExportSubTreesToWin(NodeX, ExpDirPathName)        'exports given sub-trees of ZAAN database to a Windows folders tree
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub ExportDocCopiesToWindows(ByVal ExportPathName As String, ByVal SourceDataPath As String)
        'Exports document/file copies of current ZAAN database to directories related to their first dimension available,
        'at given Windows export path and creates document/file shortcuts for remaining dimensions
        Dim dirInfo, dirSel As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim DirPathName, FileNameKey, FileName, Prefix As String
        Dim NodeKey, NodePath(5) As String
        Dim i, p, DirNb, FilNb, n, k As Integer
        Dim NodeX(10) As TreeNode

        'Debug.Print("ExportDocCopiesToWindows at : " & ExportPathName & "  SourceDataPath = " & SourceDataPath)     'TEST/DEBUG
        dirInfo = My.Computer.FileSystem.GetDirectoryInfo(SourceDataPath)

        DirNb = dirInfo.GetDirectories.Length
        n = 0
        For Each dirSel In dirInfo.GetDirectories("*.*")
            n = n + 1
            'pgbZaan.Value = 10 + 90 * n / DirNb
            DirPathName = dirInfo.FullName & "\" & dirSel.Name & "\"
            'Debug.Print("=> ExportDCTW : DirPathName = " & DirPathName)    'TEST/DEBUG

            FileNameKey = dirSel.Name                                          'get multidimensional source directory code
            Call ReplaceHierarchicalKeysBySimpleKeys(FileNameKey)          'replaces in given data directory name hierarchical keys by simple keys
            For i = 0 To 5
                NodePath(i) = ""                                               'RAZ NodePath() table
            Next
            p = 0
            For i = 0 To 5                                                     'build the NodePath() table
                p = InStr(p + 1, FileNameKey, "_", vbTextCompare)
                If p <= 0 Then Exit For
                NodeKey = Mid(FileNameKey, p + 1, mTreeKeyLength)
                NodeX = trvW.Nodes.Find(NodeKey, True)
                If NodeX.Length > 0 Then                                       'node key found in tree nodes
                    NodePath(i) = NodeX(0).FullPath
                End If
            Next
            'Debug.Print("ExportDCTW :  FileNameKey = " & FileNameKey & " => " & NodePath(0) & "|" & NodePath(1) & "|" & NodePath(2) & "|" & NodePath(3))   'TEST/DEBUG
            FilNb = dirSel.GetFiles.Length
            k = 0
            For Each fi In dirSel.GetFiles("*.*")
                k = k + 1
                pgbZaan.Value = 10 + 90 * ((n - 1) + (k / FilNb)) / DirNb      'updates progress bar
                FileName = fi.Name
                Prefix = Microsoft.VisualBasic.Left(FileName, 3)
                If FileName <> "." And FileName <> ".." And Prefix <> "zzi" And UCase(FileName) <> "THUMBS.DB" Then         'ignores current and encompassing directories, ZAAN and Windows XP thumbnail images
                    'Debug.Print("--> ExportDCTW : FileName = " & FileName)    'TEST/DEBUG
                    Call CopySelectedFile(DirPathName & FileName, "", ExportPathName & "\" & NodePath(0) & "\" & FileName)  'copies file
                    For i = 1 To 5                                             'creates shortcut for remaining dimensions
                        If NodePath(i) <> "" Then
                            Call CreateWinShortcut(ExportPathName, ExportPathName & "\" & NodePath(0) & "\" & FileName, FileName, ExportPathName & "\" & NodePath(i))   'creates shortcut
                        End If
                    Next
                End If
            Next
            'MsgBox("ExportDbProgress :  n = " & n, MsgBoxStyle.Information)   'TEST/DEBUG
        Next
    End Sub

    Private Sub ExportNameTableFromLvInSelection()
        'Exports name table of selected documents to be used for comments and/or printing
        Dim TempFilePathName, FileContent, CellText, EmailStart As String
        Dim i, j, p, DispIndexes(lvIn.Columns.Count) As Integer

        'Debug.Print("ExportNameTableFromLvInSelection...")   'TEST/DEBUG
        Me.Cursor = Cursors.WaitCursor                               'sets wait cursor
        fswZaan.EnableRaisingEvents = False                          'locks fswZaan (import and copy) related events

        For j = 0 To lvIn.Columns.Count - 1                          'builds table of columns display indexes
            DispIndexes(lvIn.Columns(j).DisplayIndex) = j
        Next
        FileContent = mMessage(83) & vbTab                           'Documents
        For j = 1 To lvIn.Columns.Count - 1                          'builds table headers line
            CellText = lvIn.Columns(DispIndexes(j)).Text
            FileContent = FileContent & CellText & vbTab
        Next
        FileContent = FileContent & vbCrLf                           'creates a new line

        For i = 0 To lvIn.SelectedItems.Count - 1                    'scans list of lvIn selected documents
            For j = 0 To lvIn.Columns.Count - 1                      'scans lvIn columns in display order
                CellText = lvIn.SelectedItems(i).SubItems(DispIndexes(j)).Text
                FileContent = FileContent & CellText & vbTab
            Next
            FileContent = FileContent & vbCrLf                       'creates a new line
        Next
        EmailStart = mUserEmail
        p = InStr(EmailStart, "@")
        If p > 0 Then
            EmailStart = Microsoft.VisualBasic.Left(EmailStart, p - 1)     'get first part of email (before @...)
        End If
        TempFilePathName = mZaanCopyPath & "_zn_" & Format(Now, "yyyy-MM-dd_HH-mm-ss") & "__" & EmailStart & "__" & UriEncode(mZaanDbName) & ".txt"
        Try
            My.Computer.FileSystem.WriteAllText(TempFilePathName, FileContent, False)    'creates
        Catch ex As System.IO.DirectoryNotFoundException
            Debug.Print(ex.Message)                                  'TEST/DEBUG
        End Try

        fswZaan.EnableRaisingEvents = True                           'unlocks import and copy related events
        Me.Cursor = Cursors.Default                                  'resets default cursor
        Call DisplayTempFiles()                                      'refreshes the display of all "temp"/"copy" files
        Call DisplayOutFiles()                                       'refresh the display of all "out" files (in case of !)
    End Sub

    Private Sub ExportRefTableFromLvInSelection()
        'Exports reference table of selected documents to be used for publishing (with Word, Publisher, etc.) in "copy" directory
        Dim DirName, FileName, FilePathName, FilePathName0, TempFilePathName, FileContent, EmailStart, CellText As String
        Dim CellTexts(lvIn.Columns.Count) As String
        Dim i, j, j0, jm, k, n, DispIndexes(lvIn.Columns.Count), p As Integer

        'Debug.Print("ExportRefTableFromLvInSelection...")   'TEST/DEBUG
        Me.Cursor = Cursors.WaitCursor                               'sets wait cursor
        fswZaan.EnableRaisingEvents = False                          'locks fswZaan (import and copy) related events

        FileContent = ""
        For j = 0 To lvIn.Columns.Count - 1                          'builds table of columns display indexes
            DispIndexes(lvIn.Columns(j).DisplayIndex) = j
            CellTexts(j) = ""
        Next
        FilePathName0 = ""
        j0 = -1                                                      'resets first displayed who/what/when/where column index
        jm = -1                                                      'resets multiple column index marker
        For i = 0 To lvIn.SelectedItems.Count - 1                    'scans list of lvIn selected documents 
            DirName = lvIn.SelectedItems(i).Tag                      'get dir name of selected file
            FileName = lvIn.SelectedItems(i).Text & lvIn.SelectedItems(i).SubItems(8).Text
            'FilePathName = mZaanDbPath & "data\" & DirName & "\" & FileName
            FilePathName = DirName & "\" & FileName

            If i = 0 Then
                FilePathName0 = FilePathName                         'stores first line file name
            End If
            For j = 0 To lvIn.Columns.Count - 1                      'scans lvIn columns in display order
                k = DispIndexes(j)                                   'get column index
                If (k > 1) And (k < 6) Then                          'is who, what, when or where column
                    If j0 = -1 Then
                        j0 = j                                       'stores first displayed who/what/when/where column index
                    End If
                    CellText = lvIn.SelectedItems(i).SubItems(k).Text
                    If CellText <> "" Then
                        If jm = -1 Then                              'multiple column has not been identified
                            If CellTexts(j) = "" Then
                                CellTexts(j) = CellText              'initializes cells before multiple column has been set
                            Else
                                If CellTexts(j) <> CellText Then     'multiple column has been found
                                    jm = j
                                    For n = j0 To jm - 1
                                        FileContent = FileContent & CellTexts(n) & vbTab
                                    Next
                                    FileContent = FileContent & FilePathName0 & vbTab & FilePathName & vbTab
                                    CellTexts(j) = CellText
                                End If
                            End If
                        Else
                            If j < jm Then                           'cell is before multiple column
                                If CellTexts(j) <> CellText Then     'cell is different than previous line => new output line
                                    FileContent = FileContent & vbCrLf                        'creates a new line
                                    For n = j0 To jm - 1
                                        FileContent = FileContent & CellTexts(n) & vbTab      'appends all first cells until this one
                                    Next
                                    CellTexts(j) = CellText
                                End If
                            Else
                                If j = jm Then                       'cell is at multiple column
                                    FileContent = FileContent & FilePathName & vbTab
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        Next

        EmailStart = mUserEmail
        p = InStr(EmailStart, "@")
        If p > 0 Then
            EmailStart = Microsoft.VisualBasic.Left(EmailStart, p - 1)     'get first part of email (before @...)
        End If
        TempFilePathName = mZaanCopyPath & "_zr_" & Format(Now, "yyyy-MM-dd_HH-mm-ss") & "__" & EmailStart & "__" & UriEncode(mZaanDbName) & ".txt"
        Try
            My.Computer.FileSystem.WriteAllText(TempFilePathName, FileContent, False)    'creates
        Catch ex As System.IO.DirectoryNotFoundException
            Debug.Print(ex.Message)                                  'TEST/DEBUG
        End Try

        fswZaan.EnableRaisingEvents = True                           'unlocks import and copy related events
        Me.Cursor = Cursors.Default                                  'resets default cursor
        Call DisplayTempFiles()                                      'refreshes the display of all "temp"/"copy" files
        Call DisplayOutFiles()                                       'refresh the display of all "out" files (in case of !)
    End Sub

    Private Function GetUserEmailStart()
        'Returns current user email start (string before @)
        Dim EmailStart As String = mUserEmail
        Dim p As Integer

        p = InStr(EmailStart, "@")
        If p > 0 Then
            EmailStart = Microsoft.VisualBasic.Left(EmailStart, p - 1)         'get first part of email (before @...)
        End If
        GetUserEmailStart = EmailStart
    End Function

    Private Sub ExportCubeFromLvInSelection()
        'Exports a document cube from current LvIn selection with related ZAAN dimensions to current destination (mZaanExportDest)
        Dim EmailStart, ZipCubeName, ZipCubeFilePathName, XmlCubeFilePathName As String
        Dim ExportDirPathName, DestFilePathName, LogFileName, LocFilePathName As String
        Dim FileContent, hm, cubePath(5), docRef, DirName, FileNameKey, PrevDirName, FileName, FileType, StoredFilePathName As String
        Dim InItem As ListViewItem
        Dim RightNow As Date = Now
        Dim i, n, m As Integer
        Dim CubeStarted As Boolean

        'If mZaanExportDest = "" Then
        '  fbdZaanPath.SelectedPath = mZaanCopyPath
        'Else
        '  fbdZaanPath.SelectedPath = mZaanExportDest
        'End If
        'fbdZaanPath.Description = mMessage(70)                                 'select a destination directory
        'fbdZaanPath.ShowNewFolderButton = False                                'disables new folder button display
        'If fbdZaanPath.ShowDialog = Windows.Forms.DialogResult.OK Then
        '  mZaanExportDest = fbdZaanPath.SelectedPath & "\"
        'End If
        mZaanExportDest = mZaanCopyPath                                        'sets ZAAN\copy as export destination directory (skips dialog box for selecting it)

        Me.Cursor = Cursors.WaitCursor                                         'sets wait cursor
        If Microsoft.VisualBasic.Right(mZaanExportDest, 1) = "\" Then
            mZaanExportDest = Microsoft.VisualBasic.Left(mZaanExportDest, Len(mZaanExportDest) - 1)
        End If
        EmailStart = GetUserEmailStart()                                       'get current user email start (string before @)

        ZipCubeName = "_zc_" & Format(RightNow, "yyyy-MM-dd_HH-mm-ss") & "__" & EmailStart & "__" & UriEncode(mZaanDbName)
        LogFileName = Mid(ZipCubeName, 5, 19) & "__export_cube_from" & Mid(ZipCubeName, 24)  'log file name to be created in "logs" directory

        ExportDirPathName = mZaanDbPath & "export"
        If CreateDirIfNotExistsOK(ExportDirPathName) Then                      'creates a db "export" directory if it doesn't exist
        End If
        DestFilePathName = mZaanExportDest & "\" & ZipCubeName & ".zip"
        ZipCubeFilePathName = ExportDirPathName & "\" & ZipCubeName & ".zip"

        Dim zip As Package = ZipPackage.Open(ZipCubeFilePathName, IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite)    'creates a new zip file

        FileContent = "  <issuer lang='" & mLanguage & "' user='" & My.User.Name & "'>" & vbCrLf
        FileContent = FileContent & "    <email>" & mUserEmail & "</email>" & vbCrLf
        FileContent = FileContent & "    <date>" & Format(RightNow, "yyyy-MM-dd HH:mm:ss") & "</date>" & vbCrLf
        FileContent = FileContent & "    <zdb>" & mZaanDbPath & "</zdb>" & vbCrLf
        FileContent = FileContent & "  </issuer>" & vbCrLf
        PrevDirName = ""
        CubeStarted = False
        n = 0                                                                  'resets selected file index
        m = 0                                                                  'resets zip internal file index
        For Each InItem In lvIn.SelectedItems                                  'scans all selected documents
            n = n + 1
            'DirName = InItem.Tag                                               'get dir name of selected doc/file
            DirName = GetLastDirName(InItem.Tag & "\")                         'get dir name of selected doc/file

            FileType = InItem.SubItems(8).Text
            FileName = InItem.Text & FileType
            'StoredFilePathName = mZaanDbPath & "data\" & DirName & "\" & FileName
            StoredFilePathName = InItem.Tag & "\" & FileName

            If FileType = ".url" Then
                'Debug.Print("Export cube : file type is : " & FileType)       'TEST/DEBUG
                docRef = GetWindowsUrlShortcut(StoredFilePathName)             'get the TargetPath value of the given Windows/Url (*.lnk/*.url) shortcut file name
            Else
                m = m + 1
                docRef = "/file" & m & FileType                                'creates a zip internal file name
                Call AddToZipFileNameUri(zip, StoredFilePathName, docRef)      'adds selected file to zip
            End If
            LocFilePathName = DirName & "\" & FileName
            Call RecordLogDataTreeUpdate(LogFileName, "data", "exp", LocFilePathName) 'records operation elements in given log file name of current database "logs" directory

            If DirName <> PrevDirName Then                                     'build a new cube with 1 to N documents having same dimensions
                PrevDirName = DirName
                FileNameKey = DirName
                Call ReplaceHierarchicalKeysBySimpleKeys(FileNameKey)          'replaces in given data directory name hierarchical keys by simple keys
                cubePath(0) = GetExportDimTreeNodePath(FileNameKey, "t", 3)
                cubePath(1) = GetExportDimTreeNodePath(FileNameKey, "o", 1)
                cubePath(2) = GetExportDimTreeNodePath(FileNameKey, "a", 2)
                cubePath(3) = GetExportDimTreeNodePath(FileNameKey, "e", 4)

                cubePath(4) = GetExportDimTreeNodePath(FileNameKey, "b", 5)
                cubePath(5) = GetExportDimTreeNodePath(FileNameKey, "c", 6)

                If CubeStarted Then                                            'close currently open cube
                    CubeStarted = False
                    FileContent = FileContent & "  </cube>" & vbCrLf
                End If
                CubeStarted = True
                FileContent = FileContent & "  <cube>" & vbCrLf                'starts cube building
                For i = 0 To 5
                    If cubePath(i) <> "" Then
                        FileContent = FileContent & cubePath(i)
                    End If
                Next
            End If
            FileContent = FileContent & "    <doc ref='" & docRef & "'>" & InItem.Text & "</doc>" & vbCrLf    'file name or url
        Next
        If CubeStarted Then                                                    'close last open cube
            CubeStarted = False
            FileContent = FileContent & "  </cube>" & vbCrLf
        End If
        hm = GetHMAC(FileContent)                                              'returns HMAC value of given message
        FileContent = "<?xml version='1.0' encoding='utf-8'?>" & vbCrLf & "<zaan_cube hm='" & hm & "'>" & vbCrLf & FileContent & "</zaan_cube>" & vbCrLf

        XmlCubeFilePathName = ExportDirPathName & "\" & ZipCubeName & ".xml"   'temporary xml file on db export directory

        Try
            My.Computer.FileSystem.WriteAllText(XmlCubeFilePathName, FileContent, False)     'creates xml file
            Call AddToZipFileNameUri(zip, XmlCubeFilePathName, "")             'adds xml file to zip
        Catch ex As System.IO.DirectoryNotFoundException
            Debug.Print("Export cube : " & ex.Message)                         'TEST/DEBUG
        End Try
        zip.Close()                                                            'closes the zip file

        Try
            My.Computer.FileSystem.DeleteFile(XmlCubeFilePathName)             'deletes temporary xml file
        Catch ex As Exception
            Debug.Print("Export cube : " & ex.Message)                         'TEST/DEBUG
        End Try

        Try                                                                    'moves zip cube to destination folder
            My.Computer.FileSystem.MoveFile(ZipCubeFilePathName, DestFilePathName, FileIO.UIOption.AllDialogs, FileIO.UICancelOption.DoNothing)
            'If mZaanExportDest & "\" = mZaanCopyPath Then DisplayTempFiles() 'refreshes the display of all "temp"/"copy" files
            DisplayTempFiles()                                                 'refreshes the display of all "temp"/"copy" files
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)                        'displays error message
        End Try
        Me.Cursor = Cursors.Default                                            'resets default cursor
    End Sub

    Private Sub RecordLogDataTreeUpdate(ByVal LogFileName As String, ByVal DirName As String, ByVal Operation As String, ByVal FileName As String, Optional ByVal Comment As String = "")
        'Records operation elements in given log file name of current database "logs" directory
        Dim LogDirPathName, FilePathName, FileContent, FileLine As String

        LogDirPathName = mZaanDbPath & "logs"
        If CreateDirIfNotExistsOK(LogDirPathName) Then               'creates a db "logs" directory if it doesn't exist
        End If

        FilePathName = LogDirPathName & "\" & LogFileName & ".txt"
        FileContent = ""
        If My.Computer.FileSystem.FileExists(FilePathName) Then      'case of log file already exists
            Try
                FileContent = My.Computer.FileSystem.ReadAllText(FilePathName)
            Catch ex As Exception
                'MsgBox(ex.Message, MsgBoxStyle.Exclamation)         'read file error !
                Debug.Print(ex.Message)                              'TEST/DEBUG error
            End Try
        End If
        FileLine = Format(Now, "yyyy-MM-dd HH:mm:ss") & vbTab & DirName & vbTab & Operation & vbTab & FileName & vbTab & Comment
        FileContent = FileContent & FileLine & vbCrLf
        Try
            My.Computer.FileSystem.WriteAllText(FilePathName, FileContent, False)
            'Debug.Print("Records in logs/" & LogFileName & " : " & FileLine)   'TEST/DEBUG
        Catch ex As System.IO.DirectoryNotFoundException
            'MsgBox(ex.Message, MsgBoxStyle.Exclamation)             'write file error !
            Debug.Print(ex.Message)                                  'TEST/DEBUG error
        End Try
    End Sub

    Private Function GetNextAvailableZ4Key(ByVal TreeCode As String) As String
        'Get next available tree code key for given tree code in v3 (4+ key length) database format
        Dim Key, PathFileName As String
        Dim KeyIndex, KeyMax, i, n As Integer

        PathFileName = Dir(mZaanDbPath & "tree\" & "_" & TreeCode & "*.txt")   'returns first tree node file of given TreeCode type (with smallest index)
        Key = mTreeRootKey                                           'initializes key to tree root key "zzzy"
        If PathFileName <> "" Then                                   'is last created tree node file of TreeCode type
            Key = Mid(PathFileName, 3, 4)                            'get its Key 
            KeyIndex = 0
            For i = 1 To 4
                n = GetIndexZ4KeyChar(Mid(Key, i, 1))
                KeyIndex = KeyIndex + n * 36 ^ (4 - i)               'builds Key index in base 36
            Next
            KeyMax = (36 ^ 4) - 1                                    '= 1 679 615
            If KeyIndex < KeyMax Then
                Key = GetZ4KeyFromIndex(KeyIndex + 1)                'get Z4 key corresponding to previous key index
            Else
                MsgBox(mMessage(120), MsgBoxStyle.Exclamation)       'No more index available for automatic folder creation !
                Key = ""                                             'invalid value to be returned
            End If
        End If
        GetNextAvailableZ4Key = Key
    End Function

    Private Function GetNextAvailableTreeAutoKey(ByVal TreeCode As String) As String
        'Get next available tree code key (starting with "_" if not empty) for given tree code (type) for automatic tree node creating (ex : import cube)
        Dim Key As String = ""

        Key = GetNextAvailableZ4Key(TreeCode)                    'get next available tree code key
        If Key <> "" Then
            Key = "_" & TreeCode & Key
        End If
        'Debug.Print("GetNextAvailableTreeAutoKey :  Key = " & Key)   'TEST/DEBUG
        GetNextAvailableTreeAutoKey = Key
    End Function

    Private Function GetDirNameCubePath(ByVal LogFileName As String, ByVal Issuer As String, ByVal cubePaths As String) As String
        'Get ZAAN data directory name corresponding to given cube paths if exists or if just created after user acceptance
        Dim CurNode, NodeX As TreeNode
        Dim PathItem(), NodeKey, DestDirName, RootPath, CurNodeKey, cubePath() As String
        Dim NewPath, TreeCode, Key, FileName, FilePathName, ParentKey As String
        Dim i, j, itemIndex, k, n As Integer
        Dim NodePathExists, TreeNodesAdded As Boolean
        Dim Response As MsgBoxResult

        'Debug.Print("=> 1-GetDirNameCubePath at : " & cubePaths)     'TEST/DEBUG
        cubePath = Split(cubePaths, "|")

        DestDirName = ""
        TreeNodesAdded = False

        For i = 1 To cubePath.Length - 1                             'searches for each given dimension if related tree node path exists
            If cubePath(i) <> "" Then                                'given cube path exists
                NodeKey = ""
                n = Mid(cubePath(i), 1, 1)                           'get related dimension index
                If mLicTypeCode < 30 Then                            'case of ZAAN-First and ZAAN-Basic licences : 4 dimensions only => Node 0 is 1 - Who
                    n = n - 1
                End If
                NodePathExists = True
                CurNode = trvW.Nodes(n)                              'get related tree root node
                RootPath = CurNode.Text
                CurNodeKey = Mid(CurNode.Tag, 1, mTreeKeyLength + 1) 'get related node key
                PathItem = Split(cubePath(i), "\")                   'splits cube path
                itemIndex = 1                                        'initializes index of first different item after root
                If PathItem.Length > 1 Then
                    For j = 1 To PathItem.Length - 1                 'starts to test each path item after root (that may be in a different language)
                        NodePathExists = False
                        For Each NodeX In CurNode.Nodes              'scans child nodes of current node
                            If NodeX.Text = PathItem(j) Then         'child node matches with given path item
                                NodePathExists = True
                                CurNode = NodeX
                                RootPath = RootPath & "\" & CurNode.Text
                                'CurNodeKey = Mid(CurNode.Tag, 1, mTreeKeyLength + 1)     'stores last found node key
                                CurNodeKey = "_" & GetTreeNodeKey(CurNode)     'returns the node key of given tree node (empty if tree root node)
                                Exit For
                            End If
                        Next
                        If Not NodePathExists Then                   'node matching child node found => stops sub-tree scanning
                            itemIndex = j                            'stores index of first different item
                            Exit For
                        End If
                    Next
                End If

                If NodePathExists Then                               'destination tree path exists
                    NodeKey = CurNodeKey
                Else                                                 'destination path does not exist
                    NewPath = RootPath
                    For k = itemIndex To PathItem.Length - 1         'scans current and remaining path items for adding them to last found root node, as children of each other
                        NewPath = NewPath & "\" & PathItem(k)        'rebuild new path with local name of target root
                    Next
                    Response = MsgBox(mMessage(116) & " (< " & Issuer & ")" & vbCrLf & mMessage(117) & " :" & vbCrLf & vbCrLf & NewPath, MsgBoxStyle.YesNo + MsgBoxStyle.Question)
                    If Response = MsgBoxResult.Yes Then              '=> create new path !
                        ParentKey = CurNodeKey                       'starts at given parent node key
                        TreeCode = Mid(CurNodeKey, 2, 1)
                        Key = ""
                        For k = itemIndex To PathItem.Length - 1     'scans current and remaining path items for adding them to last found root node, as children of each other
                            Key = GetNextAvailableTreeAutoKey(TreeCode)        'get next available tree code key (starting with "_")
                            If Key = "" Then Exit For 'no more key index available !
                            FileName = Key & ParentKey & "__" & PathItem(k) & ".txt"
                            FilePathName = mZaanDbPath & "tree\" & FileName

                            Call CreateFileFromZaanResourceString(FilePathName, My.Resources._z000000000000__new)      'adds related tree node file
                            CurNode = CurNode.Nodes.Add(Key, PathItem(k), "_" & TreeCode & "_" & mImageStyle)          'adds related node to current tree view (before complete reloading at end of ImportCubeFromXin() Sub.)
                            CurNode.Tag = FileName
                            TreeNodesAdded = True

                            Call RecordLogDataTreeUpdate(LogFileName, "tree", "add", FileName)   'records operation elements in given log file name of current database "logs" directory
                            ParentKey = Key                          'updates parent key to just created node key
                        Next
                        NodeKey = Key                                'stores last created child node key
                    End If
                End If
                If Mid(NodeKey, 2, 1) = "t" Then                     'set When key at first position
                    DestDirName = NodeKey & DestDirName
                Else
                    DestDirName = DestDirName & NodeKey
                End If
            End If
        Next
        If DestDirName = "" Then
            DestDirName = "_"
        Else
            If TreeNodesAdded Then                                   'if new paths/tree nodes created, reload tree for enabling node key search in ReplaceSimpleKey4ByHierarchicalKey4()
                Call LoadTrees()                                     '(re)loads all trees of current ZAAN database
            End If
            Call ReplaceSimpleKeysByHierarchicalKeys(DestDirName)    'replaces in given data directory name simple keys by related hierarchical keys
        End If
        'Debug.Print("==> 2-GetDirNameCubePath :  DirName = " & DestDirName)    'TEST/DEBUG
        GetDirNameCubePath = DestDirName
    End Function

    Private Sub ImportDocUrlOrFile(ByVal LogFileName As String, ByVal docUrl As String, ByVal docFile As String, ByVal DestDirName As String, ByVal FileName As String)
        'Creates .url shortcut (if not empty) or moves docFile (if not empty) to destination file pathname
        Dim OkToWrite As Boolean
        Dim Response As MsgBoxResult
        Dim DestDirPathName, DestFilePathName, FileType, LocFilePathName As String
        Dim dirInfo As System.IO.DirectoryInfo

        If docUrl = "" And docFile = "" Then Exit Sub

        DestDirPathName = mZaanDbPath & "data\" & DestDirName
        DestFilePathName = DestDirPathName & "\" & FileName
        FileType = LCase(System.IO.Path.GetExtension(FileName))
        OkToWrite = True

        If My.Computer.FileSystem.FileExists(DestFilePathName) Then  'destination file exists
            OkToWrite = False
            Response = MsgBox(mMessage(118) & vbCrLf & FileName & " " & mMessage(119), MsgBoxStyle.YesNo + MsgBoxStyle.Question)    '<filename> already exists ! Do you want to overwrite it ?
            If Response = MsgBoxResult.Yes Then
                OkToWrite = True                                     'OK to overwrite it
                If FileType = ".url" Then
                    Try
                        My.Computer.FileSystem.DeleteFile(DestFilePathName)
                    Catch ex As Exception
                        Debug.Print("Import doc : delete error : " & ex.Message)    'TEST/DEBUG
                    End Try
                End If
            End If
        ElseIf CreateDirIfNotExistsStatus(DestDirPathName) = 1 Then  'case of DestDirPathName directory just created
            'Debug.Print("Directory created : " & DestDirPathName)    'TEST/DEBUG
        End If

        If OkToWrite Then                                            'OK to write a new file or overwrite an existing one
            LocFilePathName = DestDirName & "\" & FileName
            If docFile = "" Then                                     'is an url
                If FileName = "_copy_all_files.url" Then             'case of a _copy_all_files command from docUrl source directory
                    'Debug.Print("COPY ALL FILES FROM : " & docUrl & " TO: " & DestDirPathName)    'TEST/DEBUG
                    dirInfo = My.Computer.FileSystem.GetDirectoryInfo(docUrl)
                    For Each fi In dirInfo.GetFiles("*.*")           'scans all source files to be copied in destination ZAAN data directory
                        Call CopySelectedFile(fi.FullName, "", DestDirPathName & "\" & fi.Name)    'copies file to destination directory
                    Next
                Else
                    'Debug.Print("ImportDocUrlOrFile : CREATE SHORTCUT: " & docUrl & " TO: " & DestFilePathName)   'TEST/DEBUG
                    Call CreateUrlOrWinShortcut(docUrl, DestFilePathName)     'creates shortcut and eventually overwrites existing one
                    Call RecordLogDataTreeUpdate(LogFileName, "data", "add", LocFilePathName)     'records operation elements in given log file name of current database "logs" directory
                End If
            Else                                                     'is a file
                'Debug.Print("ImportDocUrlOrFile : COPY FILE: " & docFile & " TO: " & DestFilePathName)   'TEST/DEBUG
                Try
                    My.Computer.FileSystem.MoveFile(docFile, DestFilePathName, True)      'moves file and eventually overwrites existing one
                    Call RecordLogDataTreeUpdate(LogFileName, "data", "add", LocFilePathName) 'records operation elements in given log file name of current database "logs" directory
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation)      'displays error message
                End Try
            End If
        End If
    End Sub

    Private Sub ExportSubTreeTable(ByRef FileContent As String, ByVal ParentNode As TreeNode, ByVal IndentString As String)
        'Updates given string FileContent with tabulated lines containing children titles of given node
        Dim NodeX As TreeNode

        'Debug.Print("ExportSubTreeTable :  ParentNode = " & ParentNode.Text)   'TEST/DEBUG
        IndentString = IndentString & vbTab
        For Each NodeX In ParentNode.Nodes
            'Debug.Print("ExportSubTreeTable : " & IndentString & NodeX.Text)   'TEST/DEBUG
            FileContent = FileContent & IndentString & NodeX.Text & vbCrLf     'indent child node title
            Call ExportSubTreeTable(FileContent, NodeX, IndentString)          'updates given string FileContent with tabulated lines containing children titles of given node
        Next
    End Sub

    Private Sub ExportTreeTable()
        'Exports tree structure with named Who/What/Where and What2/Who2 roots to a tabulated text file
        Dim FilePathName As String = mZaanCopyPath & mZaanDbName & "_export_structure_" & Format(Now, "yyyy-MM-dd_HH-mm-ss") & ".txt"
        Dim FileContent As String = ""
        Dim IndentString As String = ""
        Dim TreeCode As String
        Dim i, i0 As Integer

        'Debug.Print("ExportTreeTable...")   'TEST/DEBUG
        If mLicTypeCode < 30 Then                                                   'case of ZAAN-First and ZAAN-Basic licenses (4 dimensions)
            i0 = 0
        Else                                                                        'case of ZAAN-Pro license (Access control + 6 dimensions)
            i0 = 1                                                                  '=> skips Access control
        End If
        For i = i0 To trvW.Nodes.Count - 1                                          'scans root nodes (after Access control in ZAAN-Pro)
            TreeCode = Mid(trvW.Nodes(i).Tag, 2, 1)                                 'get root node tree code
            If TreeCode <> "t" Then                                                 'skips When tree (implicit in v3 database format)
                FileContent = FileContent & Mid(trvW.Nodes(i).Text, 5) & vbCrLf     'keeps only root title
                Call ExportSubTreeTable(FileContent, trvW.Nodes(i), IndentString)   'updates given string FileContent with tabulated lines containing children titles of given node
            End If
        Next
        Try
            My.Computer.FileSystem.WriteAllText(FilePathName, FileContent, False)   'saves ZAAN\info\zaan_*.ini file
            MsgBox(mMessage(180), MsgBoxStyle.Information)                          'Tree view table successfully created !
        Catch ex As System.IO.DirectoryNotFoundException
            Debug.Print(ex.Message)                                  'TEST/DEBUG
        End Try
    End Sub

    Private Sub ImportTreeTable()
        'Imports tree structure with named Who/What/Where roots from a tabulated text file
        Dim ImportFilePathName, FileContent, FileLines(), LineCells(3), TreeCode, ParentKeys(16), Key, CellText, ExistingKey As String
        Dim FilePathName, FileName, ImportedNodeTitles As String
        Dim i, j, k, n, m As Integer

        ofdTxtImport.InitialDirectory = mZaanImportPath
        ofdTxtImport.Title = mMessage(170)                            'ZAAN import : select a tree view table
        ofdTxtImport.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"

        If ofdTxtImport.ShowDialog = Windows.Forms.DialogResult.OK Then
            ImportFilePathName = ofdTxtImport.FileName
            'Debug.Print("ImportTreeTable : " & ImportFilePathName)   'TEST/DEBUG
            m = 0                                                    'initializes counter of imported tree nodes
            ImportedNodeTitles = ""                                  'initializes list of imported nodes titles
            TreeCode = ""
            For n = 0 To ParentKeys.Length - 1
                ParentKeys(n) = ""
            Next
            Try
                FileContent = My.Computer.FileSystem.ReadAllText(ImportFilePathName)     'read text file in UTF-8 format by default
                FileLines = Split(FileContent, vbCrLf)
                For i = 0 To FileLines.Length - 1                    'scans each line
                    LineCells = Split(FileLines(i), vbTab)
                    For j = 0 To LineCells.Length - 1                'scans each cell
                        CellText = Trim(LineCells(j))                'removes eventual leading and trailing spaces
                        If CellText <> "" Then
                            If j = 0 Then                            'test non empty first column cells (at root level)
                                TreeCode = ""
                                For n = 0 To ParentKeys.Length - 1
                                    ParentKeys(n) = ""
                                Next
                                For k = 2 To 6                                      'scans Who/What/Where/Status/Action for/ possible root name
                                    If UCase(mMessage(k)) = UCase(CellText) Then    'matching root name has been found
                                        Select Case k
                                            Case 2               'Who
                                                TreeCode = "o"
                                            Case 3               'What
                                                TreeCode = "a"
                                            Case 4               'Where
                                                TreeCode = "e"
                                            Case 5               'Status
                                                TreeCode = "b"
                                            Case 6               'Action for
                                                TreeCode = "c"
                                        End Select
                                        ParentKeys(j) = "_" & TreeCode & mTreeRootKey
                                        For n = j + 1 To ParentKeys.Length - 1      'clears all parent keys starting at current node level
                                            ParentKeys(n) = ""
                                        Next
                                        Exit For
                                    End If
                                Next
                            Else                                     'is not a tree root level
                                If TreeCode <> "" Then
                                    ExistingKey = GetNodeKeyIfTreeFileTextExists(TreeCode & "*" & ParentKeys(j - 1), CellText)
                                    If ExistingKey <> "" Then        'case of CellText exists with same parent
                                        ParentKeys(j) = ExistingKey
                                        For n = j + 1 To ParentKeys.Length - 1      'clears all parent keys starting at current node level
                                            ParentKeys(n) = ""
                                        Next
                                    Else                             'case of CellText does not exist with same parent
                                        Key = GetNextAvailableTreeAutoKey(TreeCode)      'gets next available tree code key (starting with "_")
                                        If Key = "" Then Exit For 'no more key index available ! (error message already displayed in GetNextAvailableZ4Key() function)

                                        m = m + 1                    'increments counter of imported tree nodes
                                        If ImportedNodeTitles = "" Then
                                            ImportedNodeTitles = CellText
                                        Else
                                            ImportedNodeTitles = ImportedNodeTitles & ", " & CellText
                                        End If
                                        FileName = Key & ParentKeys(j - 1) & "__" & CellText & ".txt"
                                        ParentKeys(j) = Key
                                        For n = j + 1 To ParentKeys.Length - 1           'clears all parent keys starting at current node level
                                            ParentKeys(n) = ""
                                        Next
                                        FilePathName = mZaanDbPath & "tree\" & FileName
                                        'Debug.Print("create tree file => " & FilePathName)     'TEST/DEBUG
                                        Call CreateFileFromZaanResourceString(FilePathName, My.Resources._z000000000000__new)      'adds related tree node file
                                    End If
                                End If
                            End If
                        End If
                    Next
                Next
                MsgBox(m & " " & mMessage(179) & vbCrLf & ImportedNodeTitles, MsgBoxStyle.Information)       '<m> new item(s) imported from tree view table : ...
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation)          'error message !
            End Try
        End If
    End Sub

    Private Function ImportCubeXmlFileIsOK(ByVal LogFileName As String, ByVal TempDirPathName As String, ByVal FileName As String) As Boolean
        'Imports given ZAAN cube .xml file for importing related documents, creating related .url shortcuts or copying files from given paths
        Dim FilePathName, FileContent, SubFileContent, hm, cellValue As String
        Dim curPath, cubePath(0), DestDirName(0), docRef, docFile, docType, docFileName, issuerEmail As String
        Dim p0, p, q, r, n, i As Integer
        Dim table As DataTable
        Dim row As DataRow
        Dim column As DataColumn
        Dim ImportOK As Boolean = True
        Dim Response As MsgBoxResult

        FilePathName = TempDirPathName & "\" & FileName                   'xml file path name
        FileContent = ""
        SubFileContent = ""
        Try
            FileContent = My.Computer.FileSystem.ReadAllText(FilePathName)
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Exclamation)                   'error on file reading !
            Debug.Print(ex.Message)                                       'error on file reading !
        End Try

        If FileContent <> "" Then                                         'xml file reading has been successful
            p0 = InStr(1, FileContent, "<zaan_cube", CompareMethod.Text)
            If p0 > 0 Then
                p = InStr(1, FileContent, "<zaan_cube hm='", CompareMethod.Text)
                If p > 0 Then
                    q = InStr(p + 15, FileContent, "'>" & vbCrLf, CompareMethod.Text)
                    If q > 0 Then
                        hm = Mid(FileContent, p + 15, q - p - 15)         'get ZAAN hm signature of xml file content
                        n = Len("'>" & vbCrLf)
                        r = InStr(q + n, FileContent, "</zaan_cube>", CompareMethod.Text)
                        If r > 0 Then
                            SubFileContent = Mid(FileContent, q + n, r - q - n)
                            If hm = GetHMAC(SubFileContent) Then          'xml file content integrity is OK !
                                'Debug.Print("===> hm is OK !")            'TEST/DEBUG
                            Else
                                SubFileContent = ""                       'cancels xml file content due to integrity issue
                            End If
                        End If
                    End If
                Else                                                      'no "hm" parameter
                    q = InStr(p0 + 9, FileContent, ">" & vbCrLf, CompareMethod.Text)
                    If q > 0 Then
                        n = Len("'>" & vbCrLf)
                        r = InStr(q + n, FileContent, "</zaan_cube>", CompareMethod.Text)
                        If r > 0 Then
                            Response = MsgBox(mMessage(207), MsgBoxStyle.YesNo + MsgBoxStyle.Question)  'This cube has not been exported by ZAAN : do you confirm its import ?
                            If Response = MsgBoxResult.Yes Then
                                SubFileContent = Mid(FileContent, q + n, r - q - n)
                                'Debug.Print("SubFileContent = " & SubFileContent)    'TEST/DEBUG
                            End If
                        End If
                    End If
                End If
            End If
        End If

        If SubFileContent <> "" Then                                      'xml file content is correct
            Dim newDataSet As New DataSet("New DataSet")
            newDataSet.ReadXml(FilePathName)                              'reads and parses xml file by table/row/column
            issuerEmail = ""
            curPath = ""
            docRef = ""
            docFile = ""
            docType = ""
            docFileName = ""
            For Each table In newDataSet.Tables                           'scans each imported dataset table
                Debug.Print("TableName: " & table.TableName)              'TEST/DEBUG
                For Each row In table.Rows                                'scans each row
                    For Each column In table.Columns                      'scans each column
                        cellValue = row(column).ToString()
                        'Debug.Print(" " & column.ToString & ":" & cellValue)    'TEST/DEBUG
                        Select Case table.TableName
                            Case "issuer"
                                Select Case column.ToString
                                    Case "email"
                                        issuerEmail = cellValue
                                End Select
                            Case "cube"
                                Select Case column.ToString
                                    Case "cube_Id"
                                        i = cellValue
                                        If i > DestDirName.Length - 1 Then ReDim Preserve DestDirName(i + 1)
                                        DestDirName(i) = ""
                                        If i > cubePath.Length - 1 Then ReDim Preserve cubePath(i + 1)
                                        cubePath(i) = ""
                                End Select
                            Case "dim"
                                Select Case column.ToString
                                    Case "dim_Text"
                                        curPath = Replace(cellValue, "/", "\")           'stores path with Unix "/" separators replaced by Windows "\" separators
                                    Case "cube_Id"
                                        i = cellValue                                    'get parent cube index
                                        cubePath(i) = cubePath(i) & "|" & curPath        'adds new path to current cube path list
                                End Select
                            Case "doc"
                                Select Case column.ToString
                                    Case "ref"
                                        docRef = ""
                                        docFile = ""
                                        docType = ""
                                        If Microsoft.VisualBasic.Left(cellValue, 1) = "/" Then   'is file name uri
                                            docFile = TempDirPathName & "\" & Mid(cellValue, 2)
                                            docType = IO.Path.GetExtension(docFile)
                                        Else                              'is http url
                                            docRef = cellValue
                                            docType = ".url"
                                        End If
                                    Case "doc_Text"
                                        docFileName = cellValue           'is the Windows source file name
                                    Case "cube_Id"
                                        i = cellValue                     'get parent cube index
                                        If DestDirName(i) = "" Then
                                            DestDirName(i) = GetDirNameCubePath(LogFileName, issuerEmail, cubePath(i))  'get ZAAN data directory name corresponding to given cube paths
                                        End If
                                        'Debug.Print("import in cube " & i & "  doc : " & docFileName & docType & "  at : " & DestDirName(i))  'TEST/DEBUG
                                        ImportDocUrlOrFile(LogFileName, docRef, docFile, DestDirName(i), docFileName & docType)    'creates .url shortcut (if not empty) or moves docFile (if not empty)
                                End Select
                        End Select
                    Next
                Next
            Next
            newDataSet.Dispose()
            'Debug.Print("ZAAN cube imported : " & ZipFilePathName)        'TEST/DEBUG
        Else
            ImportOK = False
        End If
        ImportCubeXmlFileIsOK = ImportOK
    End Function

    Private Sub ImportCubeFromXin()
        'Imports any existing ZAAN cube from xin folder of current ZAAN database
        Dim DirPathName, TempDirPathName, CubeName, CubeFilePathName, FileName, FileType, FilePathName As String
        Dim LogFileName, FileList As String
        Dim dirInfo, dirTempInfo As System.IO.DirectoryInfo
        Dim fi, fiXml As System.IO.FileInfo
        Dim pkgPart As PackagePart
        Dim FileStream As Stream
        Dim returnValue As Integer
        Dim XmlFileIsOK As Boolean

        'Debug.Print("* ImportCubeFromXin...")    'TEST/DEBUG
        fswXin.EnableRaisingEvents = False                                     'locks fswXin related events
        fswTree.EnableRaisingEvents = False                                    'locks fswTree related events
        Me.Cursor = Cursors.WaitCursor                                         'sets wait cursor
        DirPathName = mZaanDbPath & "xin"
        dirInfo = My.Computer.FileSystem.GetDirectoryInfo(DirPathName)
        FileList = ""

        For Each fi In dirInfo.GetFiles("_zc_*.*")                             'searches for any existing ZAAN cube file
            FileType = System.IO.Path.GetExtension(fi.Name)
            CubeFilePathName = DirPathName & "\" & fi.Name                     'cube file path name containing ZAAN cube xml and related documents files
            'Debug.Print("** Import cube from zip : " & CubeFilePathName)       'TEST/DEBUG
            CubeName = System.IO.Path.GetFileNameWithoutExtension(fi.Name)
            LogFileName = Mid(CubeName, 5, 19) & "__import_cube_from" & Mid(CubeName, 24)  'log file name to be created in "logs" directory
            XmlFileIsOK = False

            If FileType = ".zip" Then                                          'case of a .zip ZAAN cube file (most probably exported with ZAAN)
                TempDirPathName = DirPathName & "\" & CubeName                 'temporary directory to be used for extracting files from .zip cube
                If CreateDirIfNotExistsOK(TempDirPathName) Then                'creates a temporary directory for extracting files from ZAAN cube zip file
                End If
                Dim zip As Package = ZipPackage.Open(CubeFilePathName, IO.FileMode.Open, IO.FileAccess.Read) 'open zip file
                For Each pkgPart In zip.GetParts                               'extracts files from current ZAAN cube zip file
                    FileName = pkgPart.Uri.OriginalString
                    If Microsoft.VisualBasic.Left(FileName, 1) = "/" And FileName.Length > 1 Then
                        FileName = Mid(FileName, 2)                            'eliminates leading "/"
                    End If
                    FileStream = pkgPart.GetStream()                           'get stream data of current part/zipped file
                    Dim FileBytes(FileStream.Length - 1) As Byte
                    returnValue = FileStream.Read(FileBytes, 0, FileStream.Length) 'loads bytes() array with current file stream bytes
                    FilePathName = TempDirPathName & "\" & FileName
                    File.WriteAllBytes(FilePathName, FileBytes)                'extracts zip file (package part) to ZAAN\copy directory
                    FileStream.Close()                                         'closes stream and releases all related resources
                    FileBytes = Nothing
                Next
                zip.Close()                                                    'closes the zip file

                dirTempInfo = My.Computer.FileSystem.GetDirectoryInfo(TempDirPathName)
                For Each fiXml In dirTempInfo.GetFiles("_zc_*.xml")            'get cube xml file (should be only one)
                    If ImportCubeXmlFileIsOK(LogFileName, TempDirPathName, fiXml.Name) Then  'imports given ZAAN cube .xml file for importing/creating related documents/.url shortcuts or copying files from given path
                        XmlFileIsOK = True                                     'sets flag indicating that cube import is OK and that related source file should be removed
                    End If
                    Exit For                                                   'makes sure that only one xml file of .xip cube file is handled !
                Next
                Try                                                            'tries to delete temporary directory used for unzipping source cube
                    My.Computer.FileSystem.DeleteDirectory(TempDirPathName, FileIO.DeleteDirectoryOption.DeleteAllContents)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation)                'displays error message
                End Try

            ElseIf FileType = ".xml" Then                                      'case of a .xml ZAAN cube file (most probably built with a third party application)
                If ImportCubeXmlFileIsOK(LogFileName, DirPathName, fi.Name) Then  'imports given ZAAN cube .xml file for importing/creating related documents/.url shortcuts or copying files from given path
                    XmlFileIsOK = True                                         'sets flag indicating that cube import is OK and that related source file should be removed
                End If
            End If

            If XmlFileIsOK Then                                                'case of ZAAN cube import success
                FileList = FileList & fi.Name & vbCrLf                         'adds successfully imported ZAAN cube to list for displaying final message
                Try                                                            'tries to delete unzipped and source zip directories
                    My.Computer.FileSystem.DeleteFile(CubeFilePathName)        'deletes source cube file
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation)                'displays error message
                End Try
            Else                                                               'case of ZAAN cube import failure
                MsgBox(mMessage(161), MsgBoxStyle.Exclamation)                 'Sorry, importing this ZAAN cube is not possible, its XML descriptor beeing invalid !
                Call UndoFileMove()                                            'undo file moves stored in file move pile and resets it and disables related local menu control
            End If
        Next
        Me.Cursor = Cursors.Default                                            'resets default cursor
        fswTree.EnableRaisingEvents = True                                     'unlocks fswTree related events
        tmrTrvWDisplay.Enabled = True                                          'locks during 1s. any eventual re-display of tree view
        'If mXmode = "auto" Then fswXin.EnableRaisingEvents = True  'RETIRED unlocks fswXin related events if auto Xchange mode selected
        If FileList <> "" Then
            MsgBox(FileList & mMessage(146), MsgBoxStyle.Information)          '<file list> successfully imported !
        End If
    End Sub

    Private Sub LoadLvInMenuMultipleReferences()
        'Loads in local LvIn menu Who or What references (depending on mLvInColRightClick) of selected item, if any

        If lvIn.SelectedItems.Count = 0 Then Exit Sub

        tscbLvInWhosList.Items.Clear()
        Dim Refs As String() = Split(lvIn.SelectedItems(0).SubItems(mLvInColRightClick).Text, ", ")
        Dim i As Integer
        For i = 0 To Refs.Length - 1
            tscbLvInWhosList.Items.Add(i + 1 & "-" & Refs(i))      'loads each Who/What reference found of selected document
        Next
        tscbLvInWhosList.SelectedIndex = 0                         'selects first item by default
    End Sub

    Private Sub DeleteLvInSelItemDimReference(ByVal RefText As String, ByVal TreeCode As String)
        'Deletes given dimension reference of LvIn selected item, if any selected
        Dim FileName, DirPathName, DirName, DestDirName, LocalDataDirName, DestFilePathName, NodeKey, SelPattern As String
        Dim DestDirNameSplit, TreeCodeKeysSplit As String()
        Dim dirInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "tree")
        Dim fi As System.IO.FileInfo
        Dim n, p, q As Integer
        Dim Response As MsgBoxResult

        If lvIn.SelectedItems.Count = 0 Then Exit Sub

        'Debug.Print("DeleteLvInSelItemDimReference : " & lvIn.SelectedItems(0).Text & "  RefText = " & RefText & "  TreeCode : " & TreeCode)   'TEST/DEBUG
        FileName = lvIn.SelectedItems(0).Text & lvIn.SelectedItems(0).SubItems(8).Text
        'DirPathName = mZaanDbPath & "data\" & lvIn.SelectedItems(0).Tag
        DirPathName = lvIn.SelectedItems(0).Tag
        DirName = GetLastDirName(DirPathName & "\")

        NodeKey = ""
        n = 0
        p = InStr(RefText, "-")
        If p > 0 Then
            n = Mid(RefText, 1, p - 1)
        End If
        'TreeCodeKeysSplit = Split(lvIn.SelectedItems(0).Tag, "_" & TreeCode)
        TreeCodeKeysSplit = Split(DirName, "_" & TreeCode)
        If n > 0 Then                                                          'Whos reference list index found
            NodeKey = "_" & TreeCode & Mid(TreeCodeKeysSplit(n), 1)            'get related node key (and evetual othey keys to be deleted below) with "_" before
            q = InStr(2, NodeKey, "_")
            If q > 0 Then
                NodeKey = Mid(NodeKey, 1, q - 1)                               'extracts leading node key with "_" before
            End If
        Else                                                                   'Whos reference list index not found
            SelPattern = "_" & TreeCode & "*__*" & RefText & "*.txt"           '=> searches reference text in tree files with same TreeCode
            For Each fi In dirInfo.GetFiles(SelPattern)
                NodeKey = Mid(fi.Name, 1, mTreeKeyLength + 1)                  'get related full node key with "_" before
                Exit For
            Next
        End If

        If NodeKey <> "" Then
            'DestDirName = Replace(lvIn.SelectedItems(0).Tag, NodeKey, "")      'deletes found key from item tag to build new data directory name
            DestDirName = Replace(DirName, NodeKey, "")                        'deletes found key from item tag to build new data directory name
            If DestDirName <> "" Then
                DestDirNameSplit = Split(DestDirName, "\")
                LocalDataDirName = ""
                If DestDirNameSplit.Length > 1 Then
                    LocalDataDirName = DestDirNameSplit(1)
                Else
                    If DestDirNameSplit.Length > 0 Then
                        LocalDataDirName = DestDirNameSplit(0)
                    End If
                End If
                If LocalDataDirName = "" Then
                    DestDirName = ""
                End If
            End If
            If DestDirName = "" Then
                Response = MsgBox(mMessage(150), MsgBoxStyle.YesNo + MsgBoxStyle.Question)         'Do you want to delete this last dimension ?
                If Response = MsgBoxResult.Yes Then
                    DestFilePathName = mZaanDbPath & "data\_\" & FileName                          'no more filter set
                    Call MoveSelectedFile(FileName, DirPathName, DestFilePathName, True)           'moves file
                End If
            Else
                DestFilePathName = mZaanDbPath & "data\" & DestDirName & "\" & FileName
                Call MoveSelectedFile(FileName, DirPathName, DestFilePathName, True)               'moves file
            End If
        End If
    End Sub

    Private Sub LocksFswDataInputZaan()
        'Sets wait cursor and locks fswData, fswInput and fswZaan (import and copy) related events

        Me.Cursor = Cursors.WaitCursor                               'sets wait cursor
        fswData.EnableRaisingEvents = False                          'locks fswData related events
        fswInput.EnableRaisingEvents = False                         'locks fswInput related events
        fswZaan.EnableRaisingEvents = False                          'locks fswZaan (import and copy) related events
    End Sub

    Private Sub UnlocksFswDataInputZaan()
        'Unlocks fswZaan (import and copy), fswInput and fswData related events and resets default cursor

        fswZaan.EnableRaisingEvents = True                           'unlocks import and copy related events
        fswInput.EnableRaisingEvents = True                          'unlocks input related events
        fswData.EnableRaisingEvents = True                           'unlocks fswData related events
        Me.Cursor = Cursors.Default                                  'resets default cursor
    End Sub

    Private Sub HighlightDestListItems(ByRef MyFiles() As String, ByRef DestList As ListView)
        'Highlights given moved files/items in given destination list view
        Dim FileName, DestFileName As String
        Dim MyItem As ListViewItem
        Dim i As Integer

        For i = 0 To MyFiles.Length - 1                                        'scans all given files
            FileName = System.IO.Path.GetFileName(MyFiles(i))
            For Each MyItem In DestList.Items                                  'searches related item in destination list
                If DestList Is lvIn Then                                       'case of destination is lvIn
                    DestFileName = MyItem.Text & MyItem.SubItems(8).Text
                Else                                                           'case of destination is lvOut or lvTemp
                    DestFileName = MyItem.Text & MyItem.SubItems(1).Text
                End If
                If DestFileName = FileName Then
                    'Debug.Print("HighlightDestListItems :  MyItem = " & MyItem.Text)   'TEST/DEBUG
                    MyItem.Selected = True                                     'selects/highlights moved item
                    Exit For
                End If
            Next
        Next
        DestList.Focus()                                                       'sets focus on destination list
    End Sub

    Private Sub DisplayListsAndHighlightItems(ByRef MyFiles() As String, ByRef DestList As ListView)
        'Displays "in", "out" and "temp" lists and highlights given moved files/items in given destination list view

        Call InitDisplaySelectedFiles()                                        'initializes display of all selected files, starting at first page
        Call DisplayOutFiles()                                                 'refreshes the display of all "out" files
        Call DisplayTempFiles()                                                'refreshes the display of all "temp"/"copy" files (in case of !)
        Call HighlightDestListItems(MyFiles, DestList)                         'highlights moved files/items in destination list
    End Sub

    Private Sub UpdateAfterTreeNodeSelection()
        'Updates mFileFilter and related selector/selection and makes sure that lvIn is visible
        Call UpdateFileFilterAndDisplay()                  'updates tree buttons and mFileFilter with last tree selection and displays selected files if new filter

        If mTreeViewWasClosed Then
            mTreeViewWasClosed = False                     'resets flag
            SplitContainer2.Panel1Collapsed = True         'hides again tree view
        End If
    End Sub

    Private Sub ToggleSelectionView()
        'Toggles Selection (lvIn) display between detail and image views
        If lvIn.View = View.Details Then
            Call SetLvInDisplayMode(True)                  '=> sets images view mode
        Else
            Call SetLvInDisplayMode(False)                 '=> sets details view mode
        End If
        Call DisplaySelectedFiles()                        'displays all "in" files (empty filter)
    End Sub

    Private Sub DisplaySelectorParentNode(ByVal NodeKey As String, ByVal CodeNb As Integer)
        'Displays Selector previous labels related to parent node of given node key
        Dim NodeX() As TreeNode
        Dim Title As String = ""
        Dim Key As String = ""

        'Debug.Print("DisplaySelectorParentNode :  NodeKey = " & NodeKey & "  CodeNb = " & CodeNb)    'TEST/DEBUG

        If NodeKey <> "" Then                                        'given node is not a tree root node
            If mTreeKeyLength = 0 Then                                             'no node keys in v4 database format (use directly Windows path names)
            Else
                If (Mid(NodeKey, 1, 2) = "*_") And (NodeKey.Length > 2) Then
                    NodeKey = Mid(NodeKey, 3)
                End If
            End If
            NodeX = trvW.Nodes.Find(NodeKey, True)
            If NodeX.Length > 0 Then
                If Not NodeX(0).Parent Is Nothing Then                             'parent node exists
                    Call GetNodeSelectorTitleKey(NodeX(0).Parent, Title, Key)      'gets from parent node title and formated key (both empty if root node)
                End If
            End If
            lvSelector.Items(0).SubItems(CodeNb).Text = Title            'updates related cell of first line of selector list with parent node title/key
            lvSelector.Items(0).SubItems(CodeNb).Tag = Key
        End If

        Select Case CodeNb                                           'updates related selector parent button
            Case 0
                btnDataAccess.Text = Title
                btnDataAccess.Tag = Key
                btnDataAccess.Enabled = Not (NodeKey = "")           'disables related parent button if selected node is a root node
                btnDataAccessRoot.Enabled = btnDataAccess.Enabled
            Case 1
                btnWhen.Text = Title
                btnWhen.Tag = Key
                btnWhen.Enabled = Not (NodeKey = "")                 'disables related parent button if selected node is a root node
                btnWhenRoot.Enabled = btnWhen.Enabled
            Case 2
                btnWho.Text = Title
                btnWho.Tag = Key
                btnWho.Enabled = Not (NodeKey = "")                  'disables related parent button if selected node is a root node
                btnWhoRoot.Enabled = btnWho.Enabled
            Case 3
                btnWhat.Text = Title
                btnWhat.Tag = Key
                btnWhat.Enabled = Not (NodeKey = "")                 'disables related parent button if selected node is a root node
                btnWhatRoot.Enabled = btnWhat.Enabled
            Case 4
                btnWhere.Text = Title
                btnWhere.Tag = Key
                btnWhere.Enabled = Not (NodeKey = "")                'disables related parent button if selected node is a root node
                btnWhereRoot.Enabled = btnWhere.Enabled
            Case 5
                btnWhat2.Text = Title
                btnWhat2.Tag = Key
                btnWhat2.Enabled = Not (NodeKey = "")                'disables related parent button if selected node is a root node
                btnWhat2Root.Enabled = btnWhat2.Enabled
            Case 6
                btnWho2.Text = Title
                btnWho2.Tag = Key
                btnWho2.Enabled = Not (NodeKey = "")                 'disables related parent button if selected node is a root node
                btnWho2Root.Enabled = btnWho2.Enabled
        End Select
    End Sub

    Private Sub InitBookmarkSelection()
        'Initializes bookmark list with current (mFileFilter) selection and displays selector and selected files
        Dim itmX As ListViewItem

        'Debug.Print("InitBookmarkSelection :  Nb bookmarks = " & lvBookmark.Items.Count)    'TEST/DEBUG
        For Each itmX In lvBookmark.Items                            'scans bookmark list
            If itmX.Tag = mFileFilter Then
                itmX.Selected = True
                Exit For
            End If
        Next
        Call DisplaySelector()                                       'displays selector buttons using mFileFilter selections
        Call InitDisplaySelectedFiles()                              'initializes display of all selected files, starting at first page
    End Sub

    Private Function GetLvInCellReference(ByRef ItemX As ListViewItem, ByVal ColIndex As Integer) As String
        'Returns path reference of given selection item at given column (v4 database format)
        Dim CellRef As String = ""
        Dim p, q, i As Integer

        p = InStr(ItemX.Tag, "z" & ColIndex & "\")                   'searches in item directory path for related z#\ tag
        If p > 0 Then
            CellRef = Mid(ItemX.Tag, p)
            i = ColIndex
            Do While i < lvSelector.Columns.Count - 1                'searches for next z#\ tag if any
                i = i + 1
                q = InStr(4, CellRef, "\z" & i & "\")
                If q > 0 Then
                    CellRef = Mid(CellRef, 1, q - 1)
                    Exit Do
                End If
            Loop
        End If
        GetLvInCellReference = CellRef
    End Function

    Private Sub LoadSelectorParentList()
        'Loads Selector parent list with possible parent node selections of current selector positions
        Dim i As Integer

        lvSelector.Items.Clear()                                               'clears selector list
        lvSelector.Items.Add("")                                               'adds a first line and fill it with exisiting parent nodes
        For i = 0 To lvSelector.Columns.Count - 1
            lvSelector.Items(0).SubItems.Add("")
            Call DisplaySelectorParentNode(lvSelector.Columns(i).Tag, i)
        Next
    End Sub

    Private Sub ResetSelectorLabels()
        'Resets all Selector's labels

        lblDataAccess.Text = ""
        lblDataAccess.Tag = ""
        lblWhen.Text = ""
        lblWhen.Tag = ""
        lblWho.Text = ""
        lblWho.Tag = ""
        lblWhat.Text = ""
        lblWhat.Tag = ""
        lblWhere.Text = ""
        lblWhere.Tag = ""
        lblWhat2.Text = ""
        lblWhat2.Tag = ""
        lblWho2.Text = ""
        lblWho2.Tag = ""
    End Sub

    Private Sub ResetSelector()
        'Resets selector to blank dimensions (current year will be then forced in LoadFileFilter() sub)

        Call ResetSelectorLabels()                              'resets all Selector's labels
        Call ResetFolderAndDocSearch()                          'resets folder and document search strings
        trvW.SelectedNode = trvW.Nodes(mTreeWhoRootIndex)       'updates tree node selection to first tree root => will trigger trvW_AfterSelect() event
        trvW.SelectedImageKey = trvW.SelectedNode.ImageKey
        Call UpdateAfterTreeNodeSelection()                     'updates mFileFilter and related selector/selection and makes sure that lvIn is visible
    End Sub

    Private Sub ResetSelectorDim(ByVal SelIndex As Integer)
        'Resets given dimension of selector
        'Debug.Print("ResetSelectorDim :  SelIndex = " & SelIndex)    'TEST/DEBUG

        If SelIndex = 0 Then                                         'case of access control dimension
            Call SelectExpandFocusTreeNode(trvW.Nodes(0))            'selects Access control tree root node, expand tree at node position and make sure that selected node is visible
        Else                                                         'case of who/what/when/where/what else/other dimensions
            Call SelectExpandFocusTreeNode(trvW.Nodes(mTreeWhoRootIndex + SelIndex - 1))      'selects Who tree root node, expand tree at node position and make sure that selected node is visible
        End If
        Call ResetFolderAndDocSearch()                               'resets folder and document search strings
    End Sub

    Private Sub ResizeSelectorLabel(ByVal ColIndex As Integer)
        'Resizes selector label of given column index corresponding to lvIn list

        'Debug.Print("ResizeSelectorLabel :  ColIndex = " & ColIndex & "  mIsImageView = " & mIsImageView)   'TEST/DEBUG
        'If mIsImageView Then Exit Sub

        'Debug.Print("ResizeSelectorLabel :  ColIndex = " & ColIndex)   'TEST/DEBUG

        Dim w As Integer = lvIn.Columns(ColIndex).Width
        Dim wr As Integer = 20
        Dim w1, w2 As Integer
        If w > wr Then
            w1 = wr
            w2 = w - wr
        Else
            w1 = w
            w2 = 0
        End If
        If (mLicTypeCode < 30) And ((ColIndex = 1) Or (ColIndex = 6) Or (ColIndex = 7)) Then  'case of "Basic" and "First" licenses
            w = 0
            w1 = 0
            w2 = 0
        End If

        Select Case ColIndex
            Case 0
                'pnlSelectorSearch.Width = w + pnlBookmarkLeft.Width - pnlDatabase.Width
                pnlSelectorSearch.Width = w - pnlDatabase.Width
            Case 1
                lblDataAccess.Width = w
                btnDataAccessRoot.Width = w1
                btnDataAccess.Width = w2
                btnDataAccessBlank.Width = w
            Case 2
                lblWhen.Width = w
                btnWhenRoot.Width = w1
                btnWhen.Width = w2
                btnToday.Width = w
            Case 3
                lblWho.Width = w
                btnWhoRoot.Width = w1
                btnWho.Width = w2
            Case 4
                lblWhat.Width = w
                btnWhatRoot.Width = w1
                btnWhat.Width = w2
            Case 5
                lblWhere.Width = w
                btnWhereRoot.Width = w1
                btnWhere.Width = w2
            Case 6
                lblWhat2.Width = w
                btnWhat2Root.Width = w1
                btnWhat2.Width = w2
            Case 7
                lblWho2.Width = w
                btnWho2Root.Width = w1
                btnWho2.Width = w2
        End Select
        If (ColIndex > 0) And (ColIndex < 8) Then
            lvSelector.Columns(ColIndex - 1).Width = w
        End If
    End Sub

    Private Sub AddBookmark(ByVal FileFilter As String, Optional ByVal isDefault As Boolean = False)
        'Adds given FileFilter to bookmark list if not empty, or current Selector position (mFileFilter) to bookmark list
        Dim FilterKeys, FilterTitle, FilterTab() As String
        Dim itmX As ListViewItem
        Dim p, i As Integer
        Dim exists As Boolean = False

        If FileFilter = "" Then
            FilterKeys = mFileFilter
            FilterTitle = GetBookmark()                              'returns a bookmark string of current Selector position including related dimension labels
        Else
            FilterKeys = FileFilter
            FilterTitle = "TEST"                'TEST/DEBUG : bookmark text needs to be generated
            If mTreeKeyLength > 0 Then
                p = InStr(FileFilter, " ")                           'eventually separates filter keys from saved bookmark text
                If p > 0 Then
                    FilterKeys = Mid(FileFilter, 1, p - 1)
                    FilterTitle = Mid(FileFilter, p + 1)
                End If
            End If
        End If
        For Each itmX In lvBookmark.Items                            'scans existing bookmarks
            If itmX.Tag = FilterKeys Then
                exists = True
                itmX.Selected = True                                 'highlights existing bookmark
                'MsgBox(mMessage(165), MsgBoxStyle.Information)       'This bookmark already exists !
                Exit For
            End If
        Next
        If Not exists Then                                           'case of a new bookmark
            If mTreeKeyLength = 0 Then                               'no node keys in v4 (use directly Windows path names)
                FilterTitle = ""
                FilterTab = Split(FilterKeys, "*")
                If FilterTab.Length > 1 Then
                    For i = 1 To FilterTab.Length - 1
                        p = InStrRev(FilterTab(i), "\")
                        If p > 0 Then
                            FilterTab(i) = Mid(FilterTab(i), p + 1)  'extracts last directory name from path
                        End If
                        If FilterTitle = "" Then
                            FilterTitle = FilterTitle & FilterTab(i)
                        Else
                            FilterTitle = FilterTitle & " " & FilterTab(i)
                        End If
                    Next
                End If
            End If
            itmX = lvBookmark.Items.Add(FilterTitle)                 'adds new bookmark to list
            itmX.Tag = FilterKeys
            itmX.ImageKey = "_x_" & mImageStyle
            If isDefault Then
                itmX.ImageKey = "_x_" & mImageStyle & "d"            'sets defaut cube icon ("x" marked)
            Else
                itmX.ImageKey = "_x_" & mImageStyle                  'sets regular cube icon
            End If
            If FileFilter = "" Then                                  'case of a manual bookmark addition
                itmX.Selected = True                                 'highlights new bookmark
            End If
        End If
    End Sub

    Private Sub DeleteBookmark()
        'Deletes currently selected bookmark from list
        If lvBookmark.SelectedItems.Count > 0 Then
            'Debug.Print("DeleteBookmark : " & lvBookmark.SelectedItems(0).Text)   'TEST/DEBUG
            lvBookmark.SelectedItems(0).Remove()                     'removes selected item from list
        End If
    End Sub

    Private Sub frmZaan_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        'Debug.Print("*** frmZaan_Activated...  mLicTypeCodeIni=" & mLicTypeCodeIni & "  mLicTypeCode=" & mLicTypeCode)      'TEST/DEBUG

        If mLanguage <> mLanguageIni Then                  'case of language changed in "About ZAAN" form
            'Debug.Print("new language selected = " & mLanguage)   'TEST/DEBUG
            mLanguageIni = mLanguage                       'saves initial language selection
            Call InitLocalDisplay()                        'initializes display in selected local language
        End If

        If mLicTypeText <> mLicTypeTextIni Then            'case of license change (from frmAboutZaan/TestUserLicence)
            mLicTypeTextIni = mLicTypeText
            Call UpdateZaanProButtons()                    'updates visibility of buttons reserved to ZAAN-PRO mode
            Call UpdateZaanDbRootPathName(CheckBasicDocPath(mZaanDbPath))      'updates current database selection with mZaanDbRoot, mZaanDbPath and mZaanDbName

            Call GetFileFilterIni()                        'initializes mFileFilter of current ZAAN database from db\info\zaan.ini
            Call InitFileDataWatcher()                     'initializes "data","tree" and "xin" file watchers to detect directory access/change in current ZAAN database
            Call InitTreeInOutFiles()                      'displays selected Zaan database name, logo, tree, selector and selected files and, by default, "imput", "out" and "temp" views
            Call ShowHideExtraInCol()                      'depending on ZAAN-Pro/-Basic mode, shows or hides "What else" and "Other" columns of "in" list
        End If

        If mZaanFormLoaded Then                            'case of first frmZaan load
            'Debug.Print("** ** ** mZaanFormLoaded = " & mZaanFormLoaded)  'TEST/DEBUG
            mZaanFormLoaded = False
            Call LoadTrees()                               '(re)loads all trees of current ZAAN database
            trvW.SelectedNode = trvW.Nodes(mTreeWhoRootIndex)                  'updates tree node selection to first node (=> will trigger trvW_AfterSelect() event)
            trvW.SelectedImageKey = trvW.Nodes(mTreeWhoRootIndex).ImageKey     'forces correct image display of selected node (=> WARNING : this will call again trvW_AfterSelect() event !)
            SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
        End If

        If mUserLicEndDate = "." Then                      'maxium licence entry trials reached or license expired (set in "About ZAAN" form)
            mUserLicEndDate = ".."                         'prepares ZAAN closing made below
            'MsgBox(mMessage(49) & vbCrLf & vbCrLf & mMessage(60), MsgBoxStyle.Critical)    'Sorry, ZAAN application will be closed !
            MsgBox(mMessage(223), MsgBoxStyle.Critical)    'Sorry, ZAAN application will be closed !
        End If

        If mUserLicEndDate = ".." Then
            mZaanAutoClose = True                          'enables ZAAN application closing without user confirmation
            Me.Close()                                     'closes ZAAN application
        End If
    End Sub

    Private Sub frmZaan_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Response As MsgBoxResult

        Call SaveZaanIni(False)                            'saves mLanguage, mUserEmail, mLicenseKey, mLicTypeText, mZaanDbPath and mFileFilter in zaan.ini files

        If Not mZaanAutoClose Then                         'NO AUTOMATIC closing : case of init. or license error
            Response = MsgBox(mMessage(96), CType(MsgBoxStyle.YesNo + MsgBoxStyle.Information, MsgBoxStyle))     'Do you confirm ZAAN application closing ?
            If Response = MsgBoxResult.No Then
                e.Cancel = True                            'cancels closure of main ZAAN form
            Else
                Me.Cursor = Cursors.WaitCursor             'sets wait cursor
                Call BackupTreeAndCleanData()              'backups tree into one text file in db/info folder and removes eventual remaining unused thumbnail images and empty directories from current database
                Call SaveFileFilterIni()                   'saves mFileFilter of current ZAAN database in db\info\zaan.ini
            End If
        End If
    End Sub

    Private Sub frmZaan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Start of ZAAN application
        Debug.Print("================= Start of ZAAN =================")

        mLanguage = "FR"                                   'sets French language by default
        mLanguageIni = mLanguage
        Call InitMessages()                                'initializes mMessage table in English by default (for displaying initialization errors)
        Call LoadSelectorIcons()                           'loads in imgFileTypes all o/a/t/e/c/b/x icons from ZAAN resources (WARNING : loading icons at design time causes colors variations !???)

        mZaanAutoClose = False                             'resets flag used for automatic ZAAN closing (ex: dir. init. error or license error)
        mUserLicEndDate = "..."                            'resets licence validity date to non valid date
        'mAboutZaanAutoClose = False                        'resets flag (to be set in GetZaanIni()) used for automatic closing of "About ZAAN" form
        mAboutBoxAutoClose = True                          'sets flag to be used for automatic closing of About Box (case of no licence control)
        mLockSortLvIn = True                               'flag used in DisplaySelectedFiles() to lock SortLvInColumn() at form/lvIn loading

        Call InitZaanDbPath()                              'initializes mZaanDbPath with eventually launched .zaan file content (empty if direct ZAAN application launch)
        Call CreateOnceZaanDirectories()                   'creates once, if not existing, "ZAAN" and "ZAAN_demos" directories and related sub-directories in MyDocuments
        Call GetZaanIni()                                  'initializes mLanguage, mOpacity, mImageStyle, mUserEmail, mLicenseKey, mLicTypeText, mZaanDbPath and mFileFilter from zaan.ini files

        Call SetFormAndPanelsSizes()                       'sets main form position and size and panels size
        'pnlBookmarkLeft.Width = 0                          'hides bookmarks list/panel

        Call InitLocalDisplay()                            'initializes display in selected local language
        'Call UpdateNewSubMenuItems()                       'updates lvIn "New" sub menu items depending on available office applications

        tmrLicenseCheck.Enabled = True                     'starts timer for later opening the "About ZAAN" window and checking User Licence validity

        mTreeClicked = False                               'initializes flag used in trvW_AfterSelect() event
        mInputTreeClicked = False
        mSplitter3MouseUp = False
        mSplitter5MouseUp = False

        Call InitFileMoveCopyStreams()                     'initializes mStreamMove and mStreamCopy stream parameters used in cut and paste operations
        mFileMovePile = ""                                 'initializes mFileMovePile that is used for storing file moves and enabling eventual undo (moving back)

        pnlZoom.Dock = DockStyle.Fill                      'extends image control (initially invisible) in zoom panel
        pnlZoomVideo.Dock = DockStyle.Fill                 'extends video control (initially invisible) in zoom panel
        'wbDoc.Dock = DockStyle.Fill                        'extends Web browser control (initially invisible) in zoom panel
        'aapDoc.Dock = DockStyle.Fill                       'extends Adobe Acrobat Reader control (initially invisible) in zoom panel

        mCmsSelChildrenVisibleTreeCode = ""                'initializes flag indicating if cmsSelChildren menu is visible

        mZaanFormLoaded = True                             'set once at form loading to be checked at its first activation for tree (re)displaying
    End Sub

    Private Sub SelectExpandFocusTreeNode(ByVal SelNode As TreeNode)
        'Selects given tree node, expands tree at node position and makes sure that selected node is visible

        'Debug.Print("SelectExpandFocusTreeNode: SelNode=" & SelNode.Text & " Tag=" & SelNode.Tag & " ImageKey=" & SelNode.ImageKey)      'TEST/DEBUG

        'mTreeClicked = True                                'flag set for enabling UpdateFileFilterAndDisplay() call in trvW_AfterSelect() event
        trvW.SelectedNode = SelNode                        'updates tree node selection => will trigger trvW_AfterSelect() event
        trvW.SelectedImageKey = SelNode.ImageKey           'forces correct image display of selected node => WARNING : this will call again trvW_AfterSelect() event !!!
        Call UpdateFileFilterAndDisplay()                  'updates selector buttons and mFileFilter with last tree selection and displays selected files if new filter

        trvW.SelectedNode.Expand()                         'expands selected node
        trvW.SelectedNode.EnsureVisible()                  'makes sure that selected node is visible
        trvW.Focus()
    End Sub

    Private Function GetNodeKeyFromButtonTag(ByVal TreeCode As String, ByVal ButtonTag As String) As String
        'Returns node key corresponding to given selector button tag
        Dim TreeCodes() As String = Split(mTreeCodeSeries, " ")
        Dim NodeKey As String
        Dim i As Integer

        If mTreeKeyLength = 0 Then                               'no node keys in v4 database format (use directly Windows path names)
            NodeKey = ButtonTag
            If ButtonTag = "" Then                               'case of a tree root node
                'NodeKey = TreeCode & mTreeRootKey
                For i = 0 To TreeCodes.Length - 1
                    If TreeCodes(i) = TreeCode Then
                        NodeKey = "z" & i + 1
                        Exit For
                    End If
                Next
            End If
        Else
            If ButtonTag = "" Then                               'case of a tree root node
                NodeKey = TreeCode & mTreeRootKey
            Else                                                 'case of a child node
                NodeKey = Mid(ButtonTag, 3)
            End If
        End If
        GetNodeKeyFromButtonTag = NodeKey
    End Function

    Private Sub ShowBookmarkMenu(ByRef ResImage As Image, ByRef sender As Object)
        'Shows bookmark menu below related selector button
        Dim control As Control = CType(sender, Control)
        Dim StartPoint As System.Drawing.Point = control.PointToScreen(New System.Drawing.Point(6, lblBookmark.Height))
        Dim itmX As ListViewItem
        Dim item

        tmrSelChildren.Enabled = False                               'stops timer used for clearing mCmsSelChildrenVisibleTreeCode flag
        cmsSelBookmarks.Items.Clear()
        For Each itmX In lvBookmark.Items                            'scans bookmark list
            item = cmsSelBookmarks.Items.Add(itmX.Text)
            item.Tag = itmX.Tag
            item.image = ResImage
        Next
        cmsSelBookmarks.Show(StartPoint)
    End Sub

    Private Sub ShowChildrenMenu(ByVal TreeCode As String, ByRef SelButton As Label, ByRef ResImage As Image, ByRef sender As Object)
        'Shows children menu below selected parent label in selector
        Dim control As Control = CType(sender, Control)
        Dim StartPoint As System.Drawing.Point = control.PointToScreen(New System.Drawing.Point(0, pnlSelectorHeader.Height))
        Dim NodeKey As String = GetNodeKeyFromButtonTag(TreeCode, SelButton.Tag)
        Dim NodeX(), ChildNode As TreeNode
        Dim item
        Dim i, h, hmax As Integer

        'Debug.Print("ShowChildrenMenu :  TreeCode = " & TreeCode & "  SelButton.Text = " & SelButton.Text & "  SelButton.Tag = " & SelButton.Tag & " => NodeKey = " & NodeKey)    'TEST/DEBUG
        tmrSelChildren.Enabled = False                               'stops timer used for clearing mCmsSelChildrenVisibleTreeCode flag
        cmsSelChildren.Items.Clear()

        NodeX = trvW.Nodes.Find(NodeKey, True)
        If NodeX.Length > 0 Then                                     'case of node found
            i = 0
            For Each ChildNode In NodeX(0).Nodes                     'scans related child nodes
                i = i + 1
                item = cmsSelChildren.Items.Add(ChildNode.Text)
                item.Tag = ChildNode.Tag
                item.image = ResImage
            Next
            cmsSelChildren.Width = SelButton.Width
            h = 22 * i + 6
            hmax = Me.Height - 114
            If h > hmax Then                                         'limits menu height/length until reaching ZAAN form bottom
                h = hmax
            End If
            cmsSelChildren.Height = h
            cmsSelChildren.Show(StartPoint)
        End If
    End Sub

    Private Sub ExpandMenuOrTree(ByVal TreeCode As String, ByRef SelButton As Label, ByRef ResImage As Image, ByRef sender As Object)
        'Expands selectors menu if tree view panel is hidden, tree view else

        If SplitContainer2.Panel1Collapsed Then                      'tree view panel is hidden
            'Debug.Print("ExpandsMenuOrTree :  TreeCode = " & TreeCode & "  mCmsSelChildrenVisibleTreeCode = " & mCmsSelChildrenVisibleTreeCode)    'TEST/DEBUG
            If mCmsSelChildrenVisibleTreeCode = TreeCode Then
                mCmsSelChildrenVisibleTreeCode = ""
                cmsSelChildren.Hide()
            Else
                mCmsSelChildrenVisibleTreeCode = TreeCode
                If TreeCode = "x" Then
                    Call ShowBookmarkMenu(My.Resources._x_3b, sender)               'shows bookmark menu below related selector button
                Else
                    Call ShowChildrenMenu(TreeCode, SelButton, ResImage, sender)    'shows children menu below selected parent label in selector
                End If
            End If
            'Debug.Print(" => ExpandsMenuOrTree :  TreeCode = " & TreeCode & "  mCmsSelChildrenVisibleTreeCode = " & mCmsSelChildrenVisibleTreeCode)    'TEST/DEBUG
        Else
            If TreeCode = "x" Then
                lvBookmark.Visible = True
                trvW.Visible = False
            Else
                ExpandTree(TreeCode, SelButton.Tag)
                trvW.Visible = True
                lvBookmark.Visible = False
            End If
        End If
    End Sub

    Private Sub ExpandTree(ByVal TreeCode As String, ByVal ButtonTag As String, Optional ByVal ShowPanelTree As Boolean = False)
        'Selects and expands tree at given tree code root if given button tag is empty, or at related node if not
        Dim NodeKey As String = GetNodeKeyFromButtonTag(TreeCode, ButtonTag)
        Dim NodeX() As TreeNode

        'Debug.Print("ExpandTree :  TreeCode = " & TreeCode & "  ButtonTag = " & ButtonTag & " => NodeKey = " & NodeKey)    'TEST/DEBUG

        If ShowPanelTree Then
            mTreeViewWasClosed = SplitContainer2.Panel1Collapsed         'set a flag for later enabling automatic hidding of tree view after a node selection
            trvW.Visible = True
            lvBookmark.Visible = False
            Call SetLeftPanelButton()                                'sets left panel button display depending on splitter and background color
            SplitContainer2.Panel1Collapsed = False                  'makes sure that tree view panel is visible
        Else
            mTreeViewWasClosed = False
        End If
        trvW.CollapseAll()

        NodeX = trvW.Nodes.Find(NodeKey, True)
        If NodeX.Length > 0 Then                                     'case of node found
            Call SelectExpandFocusTreeNode(NodeX(0))                 'selects found tree node, expand tree at node position and make sure that selected node is visible
        End If
    End Sub

    Private Sub ExpandTreeNode(ByVal TreeCode As String, ByVal NodeTag As String)
        'Selects and expands tree at given tree code root if given node tag is empty, or at related node if not
        Dim ButtonTag As String

        'Debug.Print("ExpandTreeNode :  TreeCode = " & TreeCode & "  NodeTag = " & NodeTag)    'TEST/DEBUG

        If (TreeCode = "t") Or (TreeCode = "u") Then
            ButtonTag = "*" & NodeTag
        Else
            ButtonTag = "*_" & Mid(NodeTag, 2, mTreeKeyLength)
        End If
        Call ExpandTree(TreeCode, ButtonTag)                         'selects and expands tree at given tree code root if given button tag is empty, or at related node if not
    End Sub

    Private Sub CheckNodeAddStatus()
        'Checks child node (many) addition possibility that should be not enabled for access control group node (children of trvW.Nodes(0))
        Dim TreeCode As String

        tsmiTrvWAdd.Enabled = False
        tsmiTrvWAddExternal.Enabled = False
        tsmiTrvWRename.Enabled = False

        If Not trvW.SelectedNode Is Nothing Then                               'checks if a tree node has been selected
            TreeCode = Mid(trvW.SelectedNode.Tag, 2, 1)
            If TreeCode = "t" Then                                             'in v3 database format, only years may be added to When root
                If Mid(trvW.SelectedNode.Tag, 3, 4) = mTreeRootKey Then        'case of When root node
                    tsmiTrvWAdd.Enabled = True
                End If
            Else
                If TreeCode = "u" Then
                    If trvW.SelectedNode.Tag = trvW.Nodes(0).Tag Then          'case of Access control root node
                        tsmiTrvWAdd.Enabled = True
                        If mIsLocalDatabase Then                               'case of a local database (path starts with "C:" root)
                            tsmiTrvWAddExternal.Enabled = True                 '=> allows external access control folders
                        End If
                    End If
                Else
                    tsmiTrvWAdd.Enabled = True
                End If
                tsmiTrvWRename.Enabled = True
            End If
        End If
    End Sub

    Private Function IsEmptyFolderData(ByVal DataPath As String, ByVal DataFilter As String) As Boolean
        'Returns true if data folder with given filter has no data files attached to it
        Dim dirInfo, dirSel As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim IsEmpty As Boolean = True

        'Debug.Print("IsEmptyFolderData :  DataPath = " & DataPath & "  DataFilter = " & DataFilter)    'TEST/DEBUG

        dirInfo = My.Computer.FileSystem.GetDirectoryInfo(DataPath)
        For Each dirSel In dirInfo.GetDirectories(DataFilter, SearchOption.TopDirectoryOnly)  'searches for ZAAN cube directories
            For Each fi In dirSel.GetFiles("*.*")
                IsEmpty = False                                      'selected folder is not empty (contains at least one file) !
                Exit For
            Next
            If Not IsEmpty Then Exit For
        Next
        'Debug.Print(" => IsEmptyFolderData :  DataPath = " & DataPath & "  DataFilter = " & DataFilter & "  IsEmpty = " & IsEmpty)    'TEST/DEBUG
        IsEmptyFolderData = IsEmpty
    End Function

    Private Function IsEmptyFolder(ByVal NodeKey As String) As Boolean
        'Returns true if folder with given key has no data files attached to it
        Dim SelPattern As String
        Dim NodeX As TreeNode
        Dim IsEmpty As Boolean = True

        SelPattern = "*" & GetHierarchicalKey(NodeKey, True) & "*"
        IsEmpty = IsEmptyFolderData(mZaanDbPath & "data", SelPattern)          'returns true if data folder with given filter has no data files attached to it
        If IsEmpty And mLicTypeCode >= 30 Then
            For Each NodeX In trvW.Nodes(0).Nodes                              'scans all access control group directories
                IsEmpty = IsEmptyFolderData(Mid(NodeX.Tag, 3), SelPattern)     'returns true if data folder with given filter has no data files attached to it
                If Not IsEmpty Then Exit For
            Next
        End If
        IsEmptyFolder = IsEmpty
    End Function

    Private Sub CheckNodeDelStatus()
        'Checks deletion possibility of selected node and updates Del button enable status
        Dim NodeCount As Integer
        Dim TreeCode, NodeKey, DataDirPathName As String
        'Dim GroupName As String
        Dim dirInfo, dirSel As System.IO.DirectoryInfo
        Dim DirIsEmpty As Boolean = True

        tsmiTrvWDelete.Enabled = False                               'node deletion possibility disabled by default

        If Not trvW.SelectedNode Is Nothing Then                     'checks if a tree node has been selected
            NodeKey = Mid(trvW.SelectedNode.Tag, 2, mTreeKeyLength)
            If Mid(NodeKey, 2) = mTreeRootKey Then Exit Sub 'case of a tree root node : cannot delete it !

            TreeCode = Mid(trvW.SelectedNode.Tag, 2, 1)
            If TreeCode = "u" Then                                   'case of an Access control node
                'GroupName = trvW.SelectedNode.Text
                'dirInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "data\data_" & GroupName)
                DataDirPathName = Mid(trvW.SelectedNode.Tag, 3)
                dirInfo = My.Computer.FileSystem.GetDirectoryInfo(DataDirPathName)

                For Each dirSel In dirInfo.GetDirectories()
                    DirIsEmpty = False                               'group data directory is not empty !
                    Exit For
                Next
                If DirIsEmpty Then tsmiTrvWDelete.Enabled = True
            Else                                                     'case of a regular Who/What/When/Where... node
                If TreeCode = "t" Then                               'years with empty months may be deleted, but not month(s) separately
                    If Len(trvW.SelectedNode.Text) = 4 Then          'is year : checks if data are related to this year
                        If IsEmptyFolder(NodeKey) Then tsmiTrvWDelete.Enabled = True
                    End If
                Else
                    NodeCount = trvW.SelectedNode.GetNodeCount(True)     'get number of children/sub-children/etc. nodes, if any
                    If NodeCount > 0 Then Exit Sub
                    If IsEmptyFolder(NodeKey) Then tsmiTrvWDelete.Enabled = True
                End If
            End If
        End If
    End Sub

    Private Sub SelectTreeNode(ByVal NodeKey As String)
        'Changes selection to given tree node with given node key
        Dim NodeX() As TreeNode

        'Debug.Print("SelectTreeNode :  NodeKey = " & NodeKey)    'TEST/DEBUG
        If NodeKey = "" Then Exit Sub 'given node is a tree root node

        If mTreeKeyLength = 0 Then                                   'no node keys in v4 database format (use directly Windows path names)
        Else
            If (Mid(NodeKey, 1, 2) = "*_") And (NodeKey.Length > 2) Then
                NodeKey = Mid(NodeKey, 3)
            End If
            If (Mid(NodeKey, 1, 1) = "_") And (NodeKey.Length > 1) Then
                NodeKey = Mid(NodeKey, 2)
            End If
        End If
        NodeX = trvW.Nodes.Find(NodeKey, True)
        If NodeX.Length > 0 Then
            Call SelectExpandFocusTreeNode(NodeX(0))                 'selects node, expand tree at node position and make sure that selected node is visible
            Call ResetFolderAndDocSearch()                           'resets folder and document search strings
        End If
    End Sub

    Private Function SearchTreeNodeKeyInDir(ByVal DirPath As String, ByVal DirFilter As String) As String
        'Searches for a new tree node in given data access folder directory path/filter and returns related node key else empty string
        Dim dirInfo, dirSel, dirTest As System.IO.DirectoryInfo
        Dim testDirPathName As String
        Dim NodeKey As String = ""
        Dim p As Integer

        'Debug.Print("SearchTreeNodeKeyInDir :  DirPath = " & DirPath & "  DirFilter = " & DirFilter)    'TEST/DEBUG

        dirInfo = My.Computer.FileSystem.GetDirectoryInfo(DirPath)
        For Each dirSel In dirInfo.GetDirectories(DirFilter)
            Try
                testDirPathName = ""
                For Each dirTest In dirSel.GetDirectories()          'dirSel object can be accessed only if current Windows user has related access rights
                    testDirPathName = dirTest.Name
                    Exit For
                Next
                NodeKey = "u" & dirSel.FullName                      'SINCE 2012-03-30 : uses full path of related data directory

                p = InStr(1, mFoundTreeNodeKeys, "|" & NodeKey, CompareMethod.Text)      'checks if this node has not been already founded
                If p > 0 Then                                                            'case of node already founded !
                    NodeKey = ""
                Else                                                                     'case of new tree node found !
                    Exit For
                End If
            Catch ex As Exception
                Debug.Print("SearchTreeNodeKeyInDir : " & ex.Message)                    'user access to dirSel directory is not allowed !
            End Try
        Next
        SearchTreeNodeKeyInDir = NodeKey
    End Function

    Private Sub SearchTreeNode(ByVal NodeFilter As String)
        'Searches for a new tree node with a dir name filter including txtTreeSearch
        Dim SelPattern, NodeKey As String
        Dim dirInfo, dirSel As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim NodeX() As TreeNode
        Dim i, p, q As Integer
        Dim NodeFound As Boolean

        NodeFound = False
        NodeKey = ""
        Call ResetSelectorLabels()                              'resets all Selector's labels

        'Debug.Print("SearchTreeNode :  NodeFilter =>" & NodeFilter & "<")    'TEST/DEBUG
        If NodeFilter <> "" Then
            If mTreeKeyLength = 0 Then                                                   'no node keys in v4 database format (use directly Windows path names)
                dirInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath)           'searches in Access control tree nodes
                SelPattern = "data_*" & NodeFilter & "*"
                For Each dirSel In dirInfo.GetDirectories(SelPattern)
                    NodeKey = "z0\" & Mid(dirSel.Name, 6)
                    i = InStr(1, mFoundTreeNodeKeys, "|" & NodeKey, CompareMethod.Text)        'checks if this node has not been already founded
                    If i = 0 Then                                                        '=> new node founded !
                        NodeFound = True
                        Exit For
                    End If
                Next
                If Not NodeFound Then
                    dirInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "tree")  'continues search in regular Who/What/When/Where tree nodes
                    SelPattern = "*" & NodeFilter & "*"
                    For Each dirSel In dirInfo.GetDirectories(SelPattern, SearchOption.AllDirectories)
                        Debug.Print("Dir founded =>" & dirSel.Name & "<")    'TEST/DEBUG
                        NodeKey = dirSel.FullName
                        p = InStr(NodeKey, mZaanDbPath & "tree")
                        If p > 0 Then
                            NodeKey = Mid(NodeKey, Len(mZaanDbPath & "tree") + 2)
                        End If
                        q = InStr(NodeKey, "\")
                        If q > 0 Then
                            NodeKey = "z" & Mid(NodeKey, 1, 1) & Mid(NodeKey, q)
                        End If
                        i = InStr(1, mFoundTreeNodeKeys, "|" & NodeKey, CompareMethod.Text)    'checks if this node has not been already founded
                        If i = 0 Then                                                    '=> new node founded !
                            NodeFound = True
                            Exit For
                        End If
                    Next
                End If
            Else
                NodeKey = SearchTreeNodeKeyInDir(mZaanDbPath & "data", "data_*" & NodeFilter & "*")          'searches in internal access group directories
                If (NodeKey = "") And mIsLocalDatabase Then
                    'NodeKey = SearchTreeNodeKeyInDir(mZaanDbRoot, mZaanDbName & "_*" & NodeFilter & "*")     'searches in external access group directories
                    NodeKey = SearchTreeNodeKeyInDir(mZaanDbRoot, mZaanDbName & "_data_*" & NodeFilter & "*")     'searches in external access group directories
                End If

                If NodeKey = "" Then
                    dirInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "tree")   'continues search in regular Who/What/When/Where tree nodes
                    SelPattern = "*__*" & NodeFilter & "*.txt"
                    For Each fi In dirInfo.GetFiles(SelPattern)
                        NodeKey = Mid(fi.Name, 2, mTreeKeyLength)
                        i = InStr(1, mFoundTreeNodeKeys, "|" & NodeKey, CompareMethod.Text)   'checks if this node has not been already founded
                        If i = 0 Then                                                    '=> new node founded !
                            NodeFound = True
                            Exit For
                        End If
                    Next
                Else
                    NodeFound = True
                End If
            End If
        End If

        If NodeFound Then                                            'tree node has been found !
            mFoundTreeNodeKeys = mFoundTreeNodeKeys & "|" & NodeKey
            'Debug.Print("mFoundTreeNodeKeys = " & mFoundTreeNodeKeys)    'TEST/DEBUG
            NodeX = trvW.Nodes.Find(NodeKey, True)
            If NodeX.Length > 0 Then                                 'node key found in tree nodes
                Call SelectExpandFocusTreeNode(NodeX(0))             'selects found tree node, expand tree at node position and make sure that selected node is visible
            End If
        Else
            If mFoundTreeNodeKeys = "" Then
                MsgBox(mMessage(37), MsgBoxStyle.Exclamation)        'No folder found !
            Else
                MsgBox(mMessage(38), MsgBoxStyle.Information)        'End of folder search !
            End If
        End If
    End Sub

    Private Sub SearchFileInSelection(ByVal FileNameFilter As String)
        'Searches for a file within selected/displayed files/documents
        Dim itmX As ListViewItem
        Dim FileFound As Boolean
        Dim i, j As Integer

        FileFound = False
        lvIn.SelectedItems.Clear()

        'Debug.Print("SearchFileInSelection:  FileNameFilter =>" & FileNameFilter & "<")    'TEST/DEBUG
        If FileNameFilter <> "" Then
            For Each itmX In lvIn.Items
                i = InStr(1, itmX.Text, FileNameFilter, CompareMethod.Text)
                If i > 0 Then
                    j = InStr(1, mFoundListItemKeys, itmX.Index, CompareMethod.Text)
                    If j = 0 Then
                        FileFound = True
                        mFoundListItemKeys = mFoundListItemKeys & "*" & itmX.Index
                        itmX.Selected = True
                        itmX.EnsureVisible()
                        lvIn.Focus()
                        Exit For
                    End If
                End If
            Next
        End If
        If Not FileFound Then
            If mFoundListItemKeys = "" Then
                MsgBox(mMessage(39), MsgBoxStyle.Exclamation)   ' No document found !
            Else
                MsgBox(mMessage(40), MsgBoxStyle.Information)   ' End of document search !
            End If
            'lvIn.Focus()
        End If
    End Sub

    Private Sub ResetFolderAndDocSearch()
        'Resets folder and document search strings
        mFoundTreeNodeKeys = ""
        mFoundListItemKeys = ""
        tbSearch.Text = ""
    End Sub

    Private Sub UpdateSelectorTexts(ByVal TreePos As Integer, ByVal Title As String, ByVal SelPath As String, ByVal FileFilterToUpdate As Boolean)
        'Updates selector's current label relative to given index (in v4 database format) and updates mFileFilter is requested
        Dim p, i As Integer

        If TreePos = -1 Then
            If Mid(SelPath, 1, 1) = "z" Then
                TreePos = Mid(SelPath, 2, 1)
            End If
        End If
        If Title = "" Then
            Title = SelPath
            p = InStrRev(SelPath, "\")
            If p > 0 Then
                Title = Mid(SelPath, p + 1)                          'extracts last directory name from path
            End If
        End If
        'Debug.Print("UpdateSelectorTexts :  TreePos = " & TreePos & "  Title = " & Title & "  SelPath = " & SelPath)    'TEST/DEBUG

        Select Case TreePos
            Case 0                                                   'access control
                lblDataAccess.Text = Title
                lblDataAccess.Tag = SelPath
                lvSelector.Columns(0).Tag = SelPath
            Case 1                                                   '1st who/what/when/where
                lblWho.Text = Title
                lblWho.Tag = SelPath
                lvSelector.Columns(1).Tag = SelPath
            Case 2
                lblWhat.Text = Title
                lblWhat.Tag = SelPath
                lvSelector.Columns(2).Tag = SelPath
            Case 3
                lblWhen.Text = Title
                lblWhen.Tag = SelPath
                lvSelector.Columns(3).Tag = SelPath
            Case 4
                lblWhere.Text = Title
                lblWhere.Tag = SelPath
                lvSelector.Columns(4).Tag = SelPath
            Case 5
                lblWhat2.Text = Title
                lblWhat2.Tag = SelPath
                lvSelector.Columns(5).Tag = SelPath
            Case 6
                lblWho2.Text = Title
                lblWho2.Tag = SelPath
                lvSelector.Columns(6).Tag = SelPath
        End Select
        If FileFilterToUpdate Then                                   'updates mFileFilter in given order
            mFileFilter = ""
            For i = 0 To lvSelector.Columns.Count - 1
                If lvSelector.Columns(i).Tag <> "" Then
                    mFileFilter = mFileFilter & "*" & lvSelector.Columns(i).Tag
                End If
            Next
            Call UpdateBookmarkListHeader()                          'if bookmark list is visible, updates its header with current selector position and shows it if new
            'Debug.Print("=> UpdateSelectorTexts :  mFileFilter = " & mFileFilter)    'TEST/DEBUG
        End If
    End Sub

    Private Sub UpdateSelectorLabels(ByVal TreeCode As String, ByVal Title As String, ByVal Key As String)
        'Updates selector's current label relative to given tree code with given text and key

        Select Case TreeCode
            Case "u"
                lblDataAccess.Text = Title
                lblDataAccess.Tag = Key
                lvSelector.Columns(0).Tag = Key
            Case "t"
                lblWhen.Text = Title
                lblWhen.Tag = Key
                lvSelector.Columns(1).Tag = Key
            Case "o"
                lblWho.Text = Title
                lblWho.Tag = Key
                lvSelector.Columns(2).Tag = Key
            Case "a"
                lblWhat.Text = Title
                lblWhat.Tag = Key
                lvSelector.Columns(3).Tag = Key
            Case "e"
                lblWhere.Text = Title
                lblWhere.Tag = Key
                lvSelector.Columns(4).Tag = Key
            Case "b"
                lblWhat2.Text = Title
                lblWhat2.Tag = Key
                lvSelector.Columns(5).Tag = Key
            Case "c"
                lblWho2.Text = Title
                lblWho2.Tag = Key
                lvSelector.Columns(6).Tag = Key
        End Select
    End Sub

    Private Sub UpdateBookmarkListHeader()
        'If bookmark list is visible, updates its header with current selector position and shows it if new
        Dim itmX As ListViewItem
        Dim exists As Boolean = False

        If Not lvBookmark.Visible Then Exit Sub

        lvBookmark.Columns(0).Text = GetBookmark()                   'updates bookmark list header with current selector position
        lvBookmark.HeaderStyle = ColumnHeaderStyle.Clickable         'shows header
        For Each itmX In lvBookmark.Items                            'scans existing bookmarks
            If itmX.Tag = mFileFilter Then                           'case of existing bookmark
                exists = True
                lvBookmark.HeaderStyle = ColumnHeaderStyle.None      'hides header
                itmX.Selected = True                                 'highlights existing bookmark
                Exit For
            End If
        Next
        'Debug.Print("=> UpdateBookmarkListHeader :  bm = " & lvBookmark.Columns(0).Text & "  exists = " & exists)    'TEST/DEBUG
        lvBookmark.Focus()
    End Sub

    Private Function IsTreeRootNode(ByVal NodeX As TreeNode) As Boolean
        'Returns true if given node key belongs to a tree root
        Dim isRootNode As Boolean = False
        Dim NodeKey As String
        Dim p As Integer

        If mTreeKeyLength = 0 Then                                   'no node keys in v4 database format (use directly Windows path names)
            p = InStr(NodeX.Tag, "\")
            If p = 0 Then                                            'no sub-directory found => is a root path !
                isRootNode = True
            End If
        Else
            NodeKey = GetTreeNodeKey(NodeX)                          'get node key of given tree node (empty if tree root node)
            If NodeKey = "" Then                                     'case of a tree root node
                isRootNode = True
            End If
        End If
        IsTreeRootNode = isRootNode
    End Function

    Private Sub GetNodeSelectorTitleKey(ByVal NodeX As TreeNode, ByRef Title As String, ByRef Key As String)
        'Gets from given tree node title and formated key (both empty if root node) for later update of Selector's label
        Dim NodeKey As String

        If mTreeKeyLength = 0 Then                                   'no node keys in v4 database format (use directly Windows path names)
            If IsTreeRootNode(NodeX) Then                            'case of a tree root node :
                Title = ""                                           '=> sets a blank selection
                Key = ""                                             '=> sets a passthrough file filter
            Else
                Title = NodeX.Text                                   'sets node text
                Key = NodeX.Tag                                      'sets button key for active filtering of files
            End If
        Else
            NodeKey = GetTreeNodeKey(NodeX)                          'get node key of given tree node (empty if tree root node)
            If NodeKey = "" Then                                     'case of a tree root node
                Title = ""                                           '=> sets a blank selection
                Key = ""                                             '=> sets a passthrough file filter
            Else
                Title = NodeX.Text                                   'sets node text
                Key = "*_" & NodeKey                                 'sets button key for active filtering of files
            End If
        End If
        'Debug.Print(" => GetNodeSelectorTitleKey :  NodeX = " & NodeX.Text & "  => " & Title & "|" & Key)   'TEST/DEBUG
    End Sub

    Private Sub UpdateFileFilterAndDisplay()
        'Updates selector buttons and mFileFilter with last tree node selection and displays selected files if new filter
        Dim TreeCode, CurFileFilter As String
        Dim Title As String = ""
        Dim Key As String = ""
        Dim TreePos As Integer = -1

        CurFileFilter = mFileFilter
        'Debug.Print("UpdateFileFilterAndDisplay:  CurFileFilter = " & CurFileFilter)    'TEST/DEBUG

        Call GetNodeSelectorTitleKey(trvW.SelectedNode, Title, Key)  'gets from given tree node title and formated key (both empty if root node)
        If mTreeKeyLength = 0 Then
            If Mid(trvW.SelectedNode.Tag, 1, 1) = "z" Then
                TreePos = Mid(trvW.SelectedNode.Tag, 2, 1)           'gets tree code from tree node filename (stored in node tag)
                Call UpdateSelectorTexts(TreePos, Title, Key, True)  'updates selector's current label relative to given index (in v4 database format) and mFileFilter
            End If
        Else
            TreeCode = Mid(trvW.SelectedNode.Tag, 2, 1)              'gets tree code from tree node filename (stored in node tag)
            Call UpdateSelectorLabels(TreeCode, Title, Key)          'updates selector's current and previous labels relative to given tree code with given text and key
            mFileFilter = lblDataAccess.Tag & lblWhen.Tag & lblWho.Tag & lblWhat.Tag & lblWhere.Tag & lblWhat2.Tag & lblWho2.Tag     'sets whole filter
            'Call UpdateBookmarkListHeader()                          'if bookmark list is visible, updates its header with current selector position and shows it if new
        End If

        'Debug.Print("=> UpdateFileFilterAndDisplay:   mFileFilter = " & mFileFilter)    'TEST/DEBUG

        If mFileFilter <> CurFileFilter Then
            If lvBookmark.SelectedItems.Count > 0 Then               'case of an active bookmark selection
                lvBookmark.SelectedItems(0).Selected = False         'desactivates bookmark selection
            End If
            Call InitDisplaySelectedFiles()                          'initializes display of all selected files, starting at first page
            Call UpdateBookmarkListHeader()                          'if bookmark list is visible, updates its header with current selector position and shows it if new
        End If
    End Sub

    Private Function GetLvInItemFilterFullText(ByVal MyItem As ListViewItem) As String
        'Returns a string of given LvIn item filter using each dimension labels separated by "_"
        Dim ResultText As String = ""
        Dim i As Integer
        Dim DispIndexes(lvIn.Columns.Count) As Integer

        For j = 0 To lvIn.Columns.Count - 1                          'builds table of columns display indexes
            DispIndexes(lvIn.Columns(j).DisplayIndex) = j
        Next
        For i = 0 To lvIn.Columns.Count - 1
            If (DispIndexes(i) > 1) And (DispIndexes(i) < 8) Then
                If MyItem.SubItems(DispIndexes(i)).Text <> "" Then
                    ResultText = ResultText & MyItem.SubItems(DispIndexes(i)).Text & "_"
                End If
            End If
        Next
        GetLvInItemFilterFullText = ResultText
    End Function

    Private Function GetBMitem(ByVal FolderName As String) As String
        'Return given folder name followed by " " (space) if name is not empty, else empty string
        Dim BMitem As String = ""

        If FolderName <> "" Then
            'BMitem = Mid(FolderName, 1, 4) & " "
            BMitem = FolderName & " "
        End If
        GetBMitem = BMitem
    End Function

    Private Function GetBookmark() As String
        'Returns a bookmark string of current Selector position including related dimension labels
        Dim BMtext As String = ""

        'BMtext = GetBMitem(lblDataAccess.Text) & GetBMitem(lblWho.Text) & GetBMitem(lblWhat.Text) & GetBMitem(lblWhen.Text) & GetBMitem(lblWhere.Text) & GetBMitem(lblWhat2.Text) & GetBMitem(lblWho2.Text)
        BMtext = GetBMitem(lblDataAccess.Text) & GetBMitem(lblWhen.Text) & GetBMitem(lblWho.Text) & GetBMitem(lblWhat.Text) & GetBMitem(lblWhere.Text) & GetBMitem(lblWhat2.Text) & GetBMitem(lblWho2.Text)
        GetBookmark = RTrim(BMtext)                                  'eliminates trailing space
    End Function

    Private Function IsParentNode(ByVal node1 As TreeNode, ByVal node2 As TreeNode) As Boolean
        'Returns true if node1 is a parent of node2
        If node2.Parent Is Nothing Then
            Return False
        End If
        If node2.Parent.Equals(node1) Then
            Return True
        End If
        Return IsParentNode(node1, node2.Parent)                     'recursive call of this function with parent node2 as node2
    End Function

    Private Function GetCodeKeyFromDirFilter(ByVal TreeCode As String, ByVal DirFilter As String) As String
        'Get code key (or Keys for multiple Whos) from given dir filter using given tree code
        Dim CodeKey As String
        Dim p As Integer

        CodeKey = ""
        p = 1
        Do
            p = InStr(p, DirFilter, "_" & TreeCode)                  'searches for code type occurences in dir filter
            If p > 0 Then
                If CodeKey = "" Then
                    CodeKey = Mid(DirFilter, p + 1, mTreeKeyLength)      'if exists, appends corresponding code key
                Else
                    CodeKey = CodeKey & "_" & Mid(DirFilter, p + 1, mTreeKeyLength)      'if exists, appends corresponding code key with "_" between
                End If
                p = p + 2
            End If
        Loop Until p = 0
        Return CodeKey
    End Function

    Private Function UpdateMultipleRefKeys(ByVal SourceRefKeys As String, ByVal DestNodeKey As String, ByVal FileName As String) As String
        'Returns updated multiple reference keys with given destination node key, depending on user preferences
        Dim UpdatedRefKeys As String = SourceRefKeys
        Dim q As Integer

        If DestNodeKey = "" Then                                               'case of a root node selected as destination
            UpdatedRefKeys = ""                                                'clears current reference(s)
        Else
            If tsmiLvInWhosMultiple.Checked Then                               'multiple references checked
                q = InStr(SourceRefKeys, DestNodeKey)
                If q = 0 Then                                                  'DestNodeKey doesn't exists in filter
                    UpdatedRefKeys = SourceRefKeys & "_" & DestNodeKey         'adds new reference key separated by "_"
                Else
                    MsgBox(mMessage(145), MsgBoxStyle.Information)             'This reference is already defined !
                End If
            Else                                                               'multiple references unchecked
                UpdatedRefKeys = DestNodeKey                                   'sets new Who reference
            End If
        End If
        UpdateMultipleRefKeys = UpdatedRefKeys                                 'returns updated multiple reference keys
    End Function

    Private Function ChangeNodeInFileDir(ByVal DestNode As TreeNode, ByVal SceFileDir As String, ByVal FileName As String) As String
        'Returns updated file directory with destination node key inserted at right position, eventually replacing previous one of same type
        Dim SceFileDirFilter, SceFileRoot, DestFileRoot, DestFileDirFilter, DestFileDir, DestNodeKey, TreeCode As String
        Dim DataAccessDirName, FullCodeTable(6), DimPathTab(6), DestDimPath As String
        Dim i, TreePos As Integer

        'Debug.Print("ChangeNodeInFileDir :  SceFileDir = " & SceFileDir & "  DestNode = " & DestNode.Text)     'TEST/DEBUG
        DataAccessDirName = ""
        DestFileDirFilter = ""
        SceFileDirFilter = GetLastDirName(SceFileDir & "\")
        SceFileRoot = Mid(SceFileDir, 1, Len(SceFileDir) - Len(SceFileDirFilter))
        DestFileRoot = SceFileRoot                                             'keeps source data path by default

        If mTreeKeyLength = 0 Then                                             'no node keys in v4 (use directly Windows path names)
            For i = 0 To DimPathTab.Length - 1
                DimPathTab(i) = ""
            Next
            Call UpdateDimPathTabDataPath("\" & SceFileDirFilter, DimPathTab)  'updates dimension path table with given data path name
            If Mid(DestNode.Tag, 1, 1) = "z" Then
                TreePos = Mid(DestNode.Tag, 2, 1)
                DestDimPath = Mid(DestNode.Tag, 4)                             'get path starting after z#\
                If TreePos >= 0 And TreePos < 7 Then
                    TreeCode = Mid(lvSelector.Columns(TreePos).ImageKey, 2, 1)
                    Select Case TreeCode                                       'overwrites destination node key in full code table
                        Case "u"                                               'access control
                            If DestNode.Tag = "z0" Then                        '=> root node
                                DataAccessDirName = ""
                            Else                                               '=> group node
                                DataAccessDirName = "_" & DestNode.Text
                            End If
                        Case "o", "a"                                          'multiple references possible
                            DimPathTab(TreePos) = DestDimPath                  'TEST/DEBUG sets new reference
                            'update multiple references..

                        Case Else
                            DimPathTab(TreePos) = DestDimPath                  'sets new reference
                    End Select
                End If
            End If
            For i = 1 To DimPathTab.Length - 1
                If DimPathTab(i) <> "" Then
                    DestFileDirFilter = DestFileDirFilter & "\z" & i & "\" & DimPathTab(i)
                End If
            Next
            DestFileDir = mZaanDbPath & "data" & DataAccessDirName & DestFileDirFilter
        Else
            TreeCode = Mid(DestNode.Tag, 2, 1)
            DestNodeKey = GetTreeNodeKey(DestNode)                             'get the node key of given tree node (empty if tree root node)

            DestNodeKey = GetHierarchicalKey(DestNodeKey)                      'get corresponding hierarchical key
            FullCodeTable(0) = GetFileTreeNodeKeys(SceFileDirFilter, "t")      'get Who code key (or Keys for multiple Whos)
            FullCodeTable(1) = GetFileTreeNodeKeys(SceFileDirFilter, "o")      'get What code key (or Keys for multiple Whats)
            FullCodeTable(2) = GetFileTreeNodeKeys(SceFileDirFilter, "a")      'get When code key
            FullCodeTable(3) = GetFileTreeNodeKeys(SceFileDirFilter, "e")      'get Where code key
            FullCodeTable(4) = GetFileTreeNodeKeys(SceFileDirFilter, "b")      'get What2/Status code key
            FullCodeTable(5) = GetFileTreeNodeKeys(SceFileDirFilter, "c")      'get Who2 code key
            Select Case TreeCode                                               'overwrites destination node key in full code table
                Case "u"                                                       'access control
                    If DestNodeKey = "" Then
                        DestFileRoot = mZaanDbPath & "data\"                   'sets main data path
                    Else
                        DestFileRoot = Mid(DestNode.Tag, 3) & "\"              'sets path of target Access control group name
                    End If
                Case "t"
                    FullCodeTable(0) = DestNodeKey                             'sets new When reference
                Case "o"
                    FullCodeTable(1) = UpdateMultipleRefKeys(FullCodeTable(1), DestNodeKey, FileName)   'updates single/multiple reference keys with new Who reference
                Case "a"
                    FullCodeTable(2) = UpdateMultipleRefKeys(FullCodeTable(2), DestNodeKey, FileName)   'updates single/multiple reference keys with new What reference
                Case "e"
                    FullCodeTable(3) = UpdateMultipleRefKeys(FullCodeTable(3), DestNodeKey, FileName)   'sets new Where reference
                Case "b"
                    FullCodeTable(4) = UpdateMultipleRefKeys(FullCodeTable(4), DestNodeKey, FileName)   'sets new What2/Status reference
                Case "c"
                    FullCodeTable(5) = UpdateMultipleRefKeys(FullCodeTable(5), DestNodeKey, FileName)   'sets new Who2/Action for reference
            End Select

            For i = 0 To 5                                                     'builds destination file dir filter
                If FullCodeTable(i) <> "" Then
                    DestFileDirFilter = DestFileDirFilter & "_" & FullCodeTable(i)
                End If
            Next
            If DestFileDirFilter = "" Then
                DestFileDirFilter = "_"                                        'resets file dir filter to "no filter" directory
            End If
            DestFileDir = DestFileRoot & DestFileDirFilter
        End If
        'Debug.Print(" => ChangeNodeInFileDir :  DestFileDir = " & DestFileDir)   'TEST/DEBUG
        ChangeNodeInFileDir = DestFileDir
    End Function

    Private Sub InitFileMoveCopyStreams()
        'Initializes mStreamMove and mStreamCopy stream parameters used in cut and paste operations between lists and system file manager
        mStreamMove.Write(mMoveEffectData, 0, mMoveEffectData.Length)     'stream value set for file(s) Cut/paste operation with Clipboard
        mStreamCopy.Write(mCopyEffectData, 0, mCopyEffectData.Length)     'stream value set for file(s) Copy/paste operation with Clipboard
    End Sub

    Private Function RenameFileNameOK(ByVal FilePath As String, ByVal OldFileName As String, ByVal NewFileName As String, ByVal FileType As String) As Boolean
        'Returns true if successfully renamed given old file name with new file name and related thumbnail image file (if any), using given file path and type
        Dim RenameOK As Boolean = True

        'Debug.Print("RenameFileName:  FilePath = " & FilePath & "  OldFileName = " & OldFileName & "  NewFileName = " & NewFileName)    'TEST/DEBUG
        'Debug.Print("RenameFileName:  " & OldFileName & "  => " & NewFileName)    'TEST/DEBUG
        Try
            If UCase(NewFileName) = UCase(OldFileName) Then
                My.Computer.FileSystem.RenameFile(FilePath & "\" & OldFileName & FileType, "ZAAN_TEMP" & FileType)
                My.Computer.FileSystem.RenameFile(FilePath & "\" & "ZAAN_TEMP" & FileType, NewFileName & FileType)
            Else
                My.Computer.FileSystem.RenameFile(FilePath & "\" & OldFileName & FileType, NewFileName & FileType)
            End If
            Call RenameFileImage(FilePath, OldFileName, NewFileName, FileType)      'renames thumbnail image file corresponding to given (old) file that may have been generated by ZAAN
        Catch ex As Exception
            RenameOK = False
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)              'displays error message
        End Try
        RenameFileNameOK = RenameOK
    End Function

    Private Function YearTreeNodeExists(ByVal CurDate As DateTime) As Boolean
        'Returns true i year of given date already exists as a tree view node (v3 database format)
        Dim CurYear As String = Format(CurDate, "yyyy")
        Dim CurKey As String = "t" & GetWhenV3KeyFromDateText(CurYear)
        Dim NodeX() As TreeNode
        Dim YearExists As Boolean = False

        NodeX = trvW.Nodes.Find(CurKey, True)                        'searches for tree node with given year key
        If NodeX.Length > 0 Then                                     'node key found in tree nodes
            YearExists = True                                        'year exists !
        End If
        YearTreeNodeExists = YearExists
    End Function

    Private Function TenYearTreeFileExists(ByVal CurDate As DateTime) As Boolean
        'Returns true if year of given date already exists as a tree file title (v3 database format)
        Dim CurYear As String = Format(CurDate, "yyyy")
        Dim CurTenYearCode As String = Mid(GetWhenV3KeyFromDateText(CurYear), 1, 3)
        Dim dirInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "tree")
        Dim fi As System.IO.FileInfo
        Dim FileExists As Boolean = False

        For Each fi In dirInfo.GetFiles("_t" & CurTenYearCode & "x_tzzzy__*.txt")       'searches for given year text in TreeCode tree files
            FileExists = True                                                  'file exists !
            Exit For
        Next
        TenYearTreeFileExists = FileExists                                     'returns true if given ten year tree file exists
    End Function

    Private Function GetNodeKeyIfTreeFileTextExists(ByVal TreeCode As String, ByVal TreeText As String) As String
        'Returns node key if given TreeText already exists as a tree file title with same TreeCode, empty string else
        Dim NodeKey, SelPattern As String
        Dim dirInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "tree")
        Dim fi As System.IO.FileInfo

        NodeKey = ""
        SelPattern = "_" & TreeCode & "*__" & TreeText & ".txt"
        For Each fi In dirInfo.GetFiles(SelPattern)                  'searches for given year text in TreeCode tree files
            NodeKey = Mid(fi.Name, 1, 6)                             'get related node key (ie: "_ozzzx")
            Exit For
        Next
        'Debug.Print("GetNodeKeyIfTreeFileTextExists :  TreeCode = " & TreeCode & "  TreeText = " & TreeText & "  => NodeKey = " & NodeKey)   'TEST/DEBUG
        GetNodeKeyIfTreeFileTextExists = NodeKey
    End Function

    Private Function GetNodeKeyOfTreeFileText(ByVal TreeCode As String, ByVal TreeText As String) As String
        'Returns node key if given TreeText already exists as a tree file title with same TreeCode, empty string else
        Dim NodeKey, SelPattern As String
        Dim dirInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "tree")
        Dim fi As System.IO.FileInfo

        NodeKey = ""
        If (TreeCode <> "") And (TreeText <> "") Then
            SelPattern = "_" & TreeCode & "*__" & TreeText & ".txt"
            For Each fi In dirInfo.GetFiles(SelPattern)                  'searches for given year text in TreeCode tree files
                NodeKey = Mid(fi.Name, 2, mTreeKeyLength)                'get related node key (ie: "ozzzx")
                Exit For
            Next
        End If
        'Debug.Print("GetNodeKeyOfTreeFileText :  TreeCode = " & TreeCode & "  TreeText = " & TreeText & "  => NodeKey = " & NodeKey)   'TEST/DEBUG
        GetNodeKeyOfTreeFileText = NodeKey
    End Function

    Private Function TreeFileTextExists(ByVal TreeCode As String, ByVal TreeText As String) As Boolean
        'Returns true if given TreeText already exists as a tree file title with same TreeCode
        Dim TextExists As Boolean = False
        Dim SelPattern As String
        Dim dirInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "tree")
        Dim fi As System.IO.FileInfo

        SelPattern = "_" & TreeCode & "*__" & TreeText & ".txt"
        For Each fi In dirInfo.GetFiles(SelPattern)                  'searches for given year text in TreeCode tree files
            TextExists = True                                        'text exists !
            Exit For
        Next
        'Debug.Print("TreeFileTextExists :  TreeCode = " & TreeCode & "  TreeText = " & TreeText & "  => TextExists = " & TextExists)   'TEST/DEBUG
        TreeFileTextExists = TextExists
    End Function

    Private Function YearTreeFileExists(ByVal CurDate As DateTime) As Boolean
        'Returns true if year of given date already exists as a tree file title
        YearTreeFileExists = TreeFileTextExists("t", Format(CurDate, "yyyy"))  'returns true if given TreeText already exists as a When tree file title
    End Function

    Private Sub AddTrvWChildNode(Optional ByVal ExternalFolder As Boolean = False)
        'Adds a "new" child node to selected tree node
        Dim ParentKey, TreeCode, ImageKey, FileName, FilePathName, Key, Title, DataDirPathName, CreatedYear As String
        Dim ParentNode, NodeX As TreeNode
        Dim Response As MsgBoxResult
        Dim CurDate As DateTime
        Dim YearIncr As Integer

        'Debug.Print("AddTrvWChildNode to : " & trvW.SelectedNode.Text & " - " & trvW.SelectedNode.Tag)   'TEST/DEBUG
        fswTree.EnableRaisingEvents = False                               'locks fswTree related events (will be unlocked after node label editing)
        ParentNode = trvW.SelectedNode
        ParentKey = Mid(trvW.SelectedNode.Tag, 2, mTreeKeyLength)
        TreeCode = Mid(trvW.SelectedNode.Tag, 2, 1)

        Title = mMessage(74)                                              'new (default name of a new tree node)
        ImageKey = "_" & TreeCode & "_" & mImageStyle

        Select Case TreeCode
            Case "u"                                                      'case of an Access control node => creates a sub-data/external folder
                If ExternalFolder Then                                    'case of an external folder
                    'DataDirPathName = mZaanDbRoot & "\" & mZaanDbName & "_" & Title
                    DataDirPathName = mZaanDbRoot & "\" & mZaanDbName & "_data_" & Title
                Else                                                      'case of an internal sub-data folder
                    DataDirPathName = mZaanDbPath & "data\data_" & Title
                End If
                'Key = "u" & Title
                Key = "u" & DataDirPathName
                Try
                    If Not My.Computer.FileSystem.DirectoryExists(DataDirPathName) Then      'case of directory doesn't exist
                        My.Computer.FileSystem.CreateDirectory(DataDirPathName)              'creates given directory
                    End If
                    NodeX = ParentNode.Nodes.Add(Key, Title, ImageKey)    'adds a root node with image
                    NodeX.SelectedImageKey = ImageKey
                    NodeX.Tag = "_" & Key                                 'no "tree\*.txt" file associated to this Windows directory
                    trvW.SelectedNode = NodeX
                    'btnDataAccessNoSel.Visible = True                     'makes sure that "Access control" selector's icon is visible
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation)           'displays error message
                    'Debug.Print(ex.Message)              'TEST/DEBUG
                End Try
                trvW.SelectedNode.BeginEdit()                             'forces editing of "new" node label

            Case "t"                                                      'case of a When node
                If Mid(ParentKey, 2) = mTreeRootKey Then                  'selected parent node is root
                    CurDate = Now
                    If YearTreeNodeExists(CurDate) Then                   'case of current year already exists as a tree file
                        Response = MsgBox(mMessage(174), MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation)    'Do you want to create a year to come (past year if not) ?
                        If Response = MsgBoxResult.Yes Then               '=> creates first available year to come
                            YearIncr = 1
                        Else                                              '=> creates first available past year
                            YearIncr = -1
                        End If
                        Do
                            CurDate = CurDate.AddYears(YearIncr)          'moves to next (YearIncr = 1) or previous (YearIncr = -1) year of current year
                        Loop Until Not YearTreeNodeExists(CurDate)
                    End If

                    CreatedYear = CreateTreeFileYear(CurDate)             'creates in current ZAAN database (v3 format) current year tree file
                    If CreatedYear <> "" Then
                        Call LoadTrees()                                  'reloads all trees of current ZAAN database including this new When node
                        MsgBox(CreatedYear & " " & mMessage(175), MsgBoxStyle.Information)    '<year> has been created !
                    End If
                    fswTree.EnableRaisingEvents = True                    'unlocks fswTree related events (no edition mode => cannot be unlocked after node label editing)
                End If

            Case Else                                                     'case of a Who, What, Where, What else or Other node
                Key = TreeCode & GetNextAvailableZ4Key(TreeCode)          'get next available tree code key 
                NodeX = ParentNode.Nodes.Add(Key, Title, ImageKey)        'adds a root node with image
                NodeX.SelectedImageKey = ImageKey

                FileName = "_" & Key & "_" & ParentKey & "__" & Title & ".txt"
                NodeX.Tag = FileName
                trvW.SelectedNode = NodeX

                FilePathName = mZaanDbPath & "tree\" & FileName
                'Debug.Print("add tree file : " & FilePathName)
                Call CreateFileFromZaanResourceString(FilePathName, My.Resources._z000000000000__new)
                trvW.SelectedNode.BeginEdit()                                'forces editing of "new" node label
        End Select
    End Sub

    Private Function GetLvInItemNextNewFileIndex(ByVal MyItem As ListViewItem, ByVal FilePrefix As String, ByVal StartIndex As String) As String
        'Returns next new file index available for generating a new file name in given LvIn item source directory, using given file prefix
        Dim DirPathName, FileType, FullFileName, FileIndexFormat, NextIndex, StrIndex As String
        Dim i, nMax, j As Integer

        'DirPathName = mZaanDbPath & "data\" & MyItem.Tag             'get directory name of given LvIn item
        DirPathName = MyItem.Tag                                     'get directory name of given LvIn item
        FileType = MyItem.SubItems(8).Text

        FileIndexFormat = "0000"
        nMax = lvIn.SelectedItems.Count
        i = StartIndex
        i = i + 1                                                    'moves to next index
        If i > nMax Then nMax = i
        StrIndex = nMax
        If Len(StrIndex) > Len(FileIndexFormat) Then                 'extends format with additional "0" if necessary
            For j = 1 To Len(StrIndex) - Len(FileIndexFormat)
                FileIndexFormat = FileIndexFormat & "0"
            Next
        End If
        NextIndex = Format(i, FileIndexFormat)

        FullFileName = DirPathName & "\" & FilePrefix & NextIndex & FileType
        'Debug.Print("GetLvInItemNextNewFileIndex :  FullFileName = " & FullFileName)       'TEST/DEBUG

        If My.Computer.FileSystem.FileExists(FullFileName) Then                     'case of existing file => try next index
            NextIndex = GetLvInItemNextNewFileIndex(MyItem, FilePrefix, NextIndex)  'returns next new file index available for generating a new file name in given item source directory, using given file prefix
        End If
        GetLvInItemNextNewFileIndex = NextIndex
    End Function

    Private Sub AutoRenameLvInSelection()
        'Automatically renames selected files of "in" list using current selector's folder names
        Dim DirPathName, FileType, NewFilePrefix, NewFileName, NewFileIndex, NewFileNameList As String
        Dim MyItem As ListViewItem
        Dim Response As MsgBoxResult
        Dim ListFiles As String = ""
        Dim i, p, q, n As Integer

        'Debug.Print("AutoRenameLvInSelection...")      'TEST/DEBUG
        i = 0
        For Each MyItem In lvIn.SelectedItems                        'stores only the 3 first document names
            If i = 3 Then
                ListFiles = ListFiles & vbCrLf & "- ..."
            Else
                ListFiles = ListFiles & vbCrLf & "- " & MyItem.Text
            End If
            i = i + 1
            If i > 3 Then Exit For
        Next
        Response = MsgBox(mMessage(92) & vbCrLf & ListFiles, MsgBoxStyle.YesNo + MsgBoxStyle.Information)    'do you confirm the automatic renaming of selected documents... ?
        If Response = MsgBoxResult.Yes Then
            fswData.EnableRaisingEvents = False                      'locks fswData_Changed() event (that file renaming would trigger)
            'NewFileIndex = "0000"                                    'initializes file index
            NewFileNameList = ""
            For Each MyItem In lvIn.SelectedItems
                'DirPathName = mZaanDbPath & "data\" & MyItem.Tag     'get dir name of selected file
                DirPathName = MyItem.Tag                             'get dir name of selected file

                FileType = MyItem.SubItems(8).Text
                NewFilePrefix = GetLvInItemFilterFullText(MyItem)    'returns a string of given item filter using each dimension labels separated by "_"

                p = InStr(NewFileNameList, NewFilePrefix & "|")
                If p > 0 Then                                        'case of stored file prefix (with related last new file index)
                    n = Len(NewFilePrefix) + 1
                    q = InStr(p + n, NewFileNameList, "|_")          'searches for last file index end
                    If q > 0 Then
                        NewFileIndex = Mid(NewFileNameList, p + n, q - p - n)      'get stored index
                        NewFileIndex = GetLvInItemNextNewFileIndex(MyItem, NewFilePrefix, NewFileIndex)      'returns next new file index available for generating a new file name in given item source directory, using given file prefix
                        NewFileNameList = Microsoft.VisualBasic.Left(NewFileNameList, p + n - 1) & NewFileIndex & Mid(NewFileNameList, q)      'updates new file index
                    Else
                        NewFileIndex = "0000"                        'initializes file index
                        NewFileIndex = GetLvInItemNextNewFileIndex(MyItem, NewFilePrefix, NewFileIndex)      'returns next new file index available for generating a new file name in given item source directory, using given file prefix
                        NewFileNameList = NewFileNameList & NewFilePrefix & "|" & NewFileIndex & "|_"   'stores new file prefix and related last new index
                    End If
                Else
                    NewFileIndex = "0000"                            'initializes file index
                    NewFileIndex = GetLvInItemNextNewFileIndex(MyItem, NewFilePrefix, NewFileIndex)     'returns next new file index available for generating a new file name in given item source directory, using given file prefix
                    NewFileNameList = NewFileNameList & NewFilePrefix & "|" & NewFileIndex & "|_"       'stores new file prefix and related last new index
                End If
                NewFileName = NewFilePrefix & NewFileIndex

                'Debug.Print("Auto rename file : " & MyItem.Text & " => " & NewFileName)      'TEST/DEBUG
                If RenameFileNameOK(DirPathName, MyItem.Text, NewFileName, FileType) Then     'returns true if successfully renamed given old file name with new file name and related thumbnail image file (if any), using given file path and type
                End If
            Next
            fswData.EnableRaisingEvents = True                       'unlocks fswData_Changed() event
        End If
        lvIn.Focus()
        Call InitDisplaySelectedFiles()                              'initializes display of all selected files, starting at first page
    End Sub

    Private Sub DeleteLvInSelection()
        'Deletes selected files of "in" list
        Dim DirPathName, FileName, FullFileName, FileType, DestFilePathName As String
        'Dim TargetDirPath, TargetDirNew, TargetDirPathNew, HiddenFileName As string
        Dim MyItem As ListViewItem
        Dim Response As MsgBoxResult
        Dim ListFiles As String = ""
        Dim i As Integer = 0
        Dim isDirLink As Boolean = False

        'Debug.Print("DeleteLvInSelection...")      'TEST/DEBUG
        For Each MyItem In lvIn.SelectedItems                        'stores only the 3 first document names
            If i = 3 Then
                ListFiles = ListFiles & vbCrLf & "- ..."
            Else
                ListFiles = ListFiles & vbCrLf & "- " & MyItem.Text
            End If
            i = i + 1
            If i > 3 Then Exit For
        Next
        Response = MsgBox(mMessage(29) & vbCrLf & ListFiles, MsgBoxStyle.YesNo + MsgBoxStyle.Information)    'Do you confirm the deletion of selected documents from ZAAN database ?
        If Response = MsgBoxResult.Yes Then
            Call LocksFswDataInputZaan()                             'sets wait cursor and locks fswData, fswInput and fswZaan (import and copy) related events
            For Each MyItem In lvIn.SelectedItems
                'DirPathName = mZaanDbPath & "data\" & MyItem.Tag     'get dir name of selected file
                DirPathName = MyItem.Tag                             'get dir name of selected file

                FileType = MyItem.SubItems(8).Text                    'get file extension/type
                FileName = MyItem.Text & FileType
                FullFileName = DirPathName & "\" & FileName
                'Debug.Print("DeleteLvInSelection :  FullFileName = " & FullFileName)      'TEST/DEBUG

                'If Microsoft.VisualBasic.Right(MyItem.Text, 4) = " (z)" And FileType = ".lnk" Then      'is a directory link
                '  TargetDirPath = GetWinShortcutTargetPath(FullFileName)                              'returns TargetPath of given Windows shortcut file
                '  HiddenFileName = TargetDirPath & "\" & "_zaan_link.lnk"                             'is a hidden shortcut file providing reciprocal link to ZAAN data directory
                '  isDirLink = True
                '  If Microsoft.VisualBasic.Right(TargetDirPath, 4) = " (z)" Then                      'target folder names ends with (z) mark
                '    TargetDirPathNew = Microsoft.VisualBasic.Left(TargetDirPath, Len(TargetDirPath) - 4)
                '    TargetDirNew = System.IO.Path.GetFileName(TargetDirPathNew)
                '    Try
                '      My.Computer.FileSystem.DeleteFile(HiddenFileName)                           'deletes hidden shortcut file
                '      My.Computer.FileSystem.RenameDirectory(TargetDirPath, TargetDirNew)         'deletes (z) mark of target folder name
                '    Catch ex As Exception
                '      MsgBox(ex.Message, MsgBoxStyle.Exclamation)        'displays error message
                '    End Try
                '    If mZaanImportPath = TargetDirPath & "\" Then
                '      Call UpdateImportPath(TargetDirPathNew & "\")      'updates mZaanImportPath and related tooltip and "out" files display
                '    End If
                '  End If
                'End If

                Call DeleteFileImage(DirPathName, FileName)          'deletes corresponding thumbnail image file (zzi*) that may have been generated by ZAAN
                Try
                    If Mid(FullFileName, 1, 2) = Mid(mZaanAppliPath, 1, 2) Then                    'file to be deleted is on local disk
                        My.Computer.FileSystem.DeleteFile(FullFileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                    Else
                        Call ResetFileMovePile()                                                   'resets file move pile (used for "Undo move") and disables related local menu control
                        DestFilePathName = mZaanAppliPath & "import\" & FileName
                        Call MoveSelectedFile(FullFileName, "", DestFilePathName, False)           'moves file to current import path
                        My.Computer.FileSystem.DeleteFile(DestFilePathName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                    End If

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation)      'displays error message
                End Try
                Call DeleteDirIfEmpty(DirPathName)                   'deletes changed directory (if empty) of the pile
            Next
            Call UnlocksFswDataInputZaan()                           'unlocks fswZaan (import and copy), fswInput and fswData related events and resets default cursor
            Call InitDisplaySelectedFiles()                          'initializes display of all selected files, starting at first page
            If isDirLink Then
                Call LoadInputTree()                                 'loads input tree with Windows directory structure starting at selected root
            End If
        End If
        lvIn.Focus()
    End Sub

    Private Sub DeleteSlideShowPicture()
        'Deletes, if user confirms, current picture displayed in slide show
        Dim Response As MsgBoxResult
        Dim DirPathName, FileName As String

        'Debug.Print("DeleteSlideShowPicture : " & mSlideShowPathFileName)   'TEST/DEBUG
        If mSlideShowPathFileName Is Nothing Then Exit Sub

        mIsSlideShowPaused = True                                    'makes sure that play/pause status is on pause
        Call PlayPauseSlideShow()                                    'plays or stops activated slide show

        Response = MsgBox(mMessage(138), MsgBoxStyle.YesNo + MsgBoxStyle.Information)    'Do you confirm the deletion of this picture ?
        If Response = MsgBoxResult.Yes Then
            Me.Cursor = Cursors.WaitCursor                           'sets wait cursor
            fswData.EnableRaisingEvents = False                      'locks fswData related events
            fswInput.EnableRaisingEvents = False                     'locks input related events
            fswZaan.EnableRaisingEvents = False                      'locks import and copy related events

            DirPathName = System.IO.Path.GetDirectoryName(mSlideShowPathFileName)
            FileName = System.IO.Path.GetFileName(mSlideShowPathFileName)
            Call DeleteFileImage(DirPathName, FileName)              'deletes corresponding thumbnail image file (zzi*) that may have been generated by ZAAN
            Try
                My.Computer.FileSystem.DeleteFile(mSlideShowPathFileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation)          'displays error message
            End Try
            Call DeleteDirIfEmpty(DirPathName)                       'deletes changed directory (if empty) of the pile

            fswZaan.EnableRaisingEvents = True                       'unlocks import and copy related events
            fswInput.EnableRaisingEvents = True                      'unlocks input related events
            fswData.EnableRaisingEvents = True                       'unlocks fswData related events
            Me.Cursor = Cursors.Default                              'resets default cursor
            Call InitDisplaySelectedFiles()                          'initializes display of all selected files, starting at first page
            If lvIn.Items.Count = 0 Then                             'case of picture list is empty
                tsmiPctZoomSlideShow.Checked = False
                Call StartStopSlideShow()                            'stops slide show
                Call InitBrowserPanel()                              'initializes browser display items
            Else
                Call GotoNextSlide()                                 'moves current slide cursor to next slide and displays related slide
            End If
        End If
    End Sub

    Private Sub DeleteLvOutSelection()
        'Deletes selected files of "out" list
        Dim FileName As String
        Dim MyItem As ListViewItem
        Dim Response As MsgBoxResult
        Dim ListFiles As String = ""
        Dim i As Integer = 0

        'Debug.Print("DeleteLvOutSelection...")          'TEST/DEBUG
        For Each MyItem In lvOut.SelectedItems                       'stores only the 3 first document names
            If i = 3 Then
                ListFiles = ListFiles & vbCrLf & "- ..."
            Else
                ListFiles = ListFiles & vbCrLf & "- " & MyItem.Text
            End If
            i = i + 1
            If i > 3 Then Exit For
        Next
        Response = MsgBox(mMessage(41) & vbCrLf & ListFiles, MsgBoxStyle.YesNo + MsgBoxStyle.Information)    'Do you confirm the deletion of selected documents from ZAAN\import directory ?
        If Response = MsgBoxResult.Yes Then
            Me.Cursor = Cursors.WaitCursor                           'sets wait cursor
            fswZaan.EnableRaisingEvents = False
            For Each MyItem In lvOut.SelectedItems
                FileName = mZaanImportPath & MyItem.Text & MyItem.SubItems(1).Text     'get document name & type
                Try
                    My.Computer.FileSystem.DeleteFile(FileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation)      'displays error message
                End Try
            Next
            fswZaan.EnableRaisingEvents = True
            Me.Cursor = Cursors.Default                              'resets default cursor
            Call DisplayOutFiles()                                   'refreshes the display of all "out" files
        End If
        lvOut.Focus()
    End Sub

    Private Sub ResizePicturesLvTempSelection()
        'Resizes selected .jpg files of "temp"/"copy" list after user confirmation
        Dim FileName, ResizedFileName As String
        Dim MyItem As ListViewItem
        Dim Response As MsgBoxResult
        Dim ListFiles As String = ""
        Dim i As Integer = 0
        Dim NewWidth, NewHeight As Integer
        Dim g As Graphics
        Dim destRect As Rectangle

        'Debug.Print("ResizePicturesLvTempSelection...")          'TEST/DEBUG
        For Each MyItem In lvTemp.SelectedItems                      'stores only the 3 first document names
            If LCase(MyItem.SubItems(1).Text) = ".jpg" Then
                If i = 3 Then
                    ListFiles = ListFiles & vbCrLf & "- ..."
                Else
                    ListFiles = ListFiles & vbCrLf & "- " & MyItem.Text
                End If
                i = i + 1
                If i > 3 Then Exit For
            End If
        Next
        Response = MsgBox(mMessage(98) & vbCrLf & ListFiles, MsgBoxStyle.YesNo + MsgBoxStyle.Information)    'Do you confirm the resizing of selected pictures in ZAAN\copy directory ?
        If Response = MsgBoxResult.Yes Then
            Me.Cursor = Cursors.WaitCursor                           'sets wait cursor
            fswZaan.EnableRaisingEvents = False
            For Each MyItem In lvTemp.SelectedItems
                If LCase(MyItem.SubItems(1).Text) = ".jpg" Then
                    FileName = mZaanCopyPath & MyItem.Text & MyItem.SubItems(1).Text          'get source document path name & type
                    ResizedFileName = mZaanCopyPath & MyItem.Text & "_" & tstLvTempResizePercentage.Text & "pc" & MyItem.SubItems(1).Text
                    Dim image As New Bitmap(FileName)                'loads source image (full size)
                    NewWidth = image.Width * tstLvTempResizePercentage.Text / 100
                    NewHeight = image.Height * tstLvTempResizePercentage.Text / 100
                    'Debug.Print("Picture : " & MyItem.Text & "  NewWidth = " & NewWidth & "  NewHeight = " & NewHeight)   'TEST/DEBUG

                    Dim bm As New Bitmap(NewWidth, NewHeight)
                    destRect.X = 0
                    destRect.Y = 0
                    destRect.Width = NewWidth
                    destRect.Height = NewHeight
                    g = Graphics.FromImage(bm)
                    g.DrawImage(image, destRect)                     'draws source image at specified location and with the specified size
                    bm.Save(ResizedFileName, System.Drawing.Imaging.ImageFormat.Jpeg)    'saves resized image in jpeg format

                    image.Dispose()
                    bm.Dispose()

                    Try                                              'try to delete source picture file from ZAAN\copy
                        'My.Computer.FileSystem.DeleteFile(FileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                        My.Computer.FileSystem.DeleteFile(FileName)
                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.Exclamation)  'displays error message
                    End Try
                End If
            Next
            fswZaan.EnableRaisingEvents = True
            Me.Cursor = Cursors.Default                              'resets default cursor
            Call DisplayTempFiles()                                  'refreshes the display of all "temp"/"copy" files
        End If
        lvTemp.Focus()
    End Sub

    Private Sub DeleteLvTempSelection()
        'Deletes selected files of "temp"/"copy" list
        Dim FileName As String
        Dim MyItem As ListViewItem
        Dim Response As MsgBoxResult
        Dim ListFiles As String = ""
        Dim i As Integer = 0

        'Debug.Print("DeleteLvTempSelection...")          'TEST/DEBUG
        For Each MyItem In lvTemp.SelectedItems                       'stores only the 3 first document names
            If i = 3 Then
                ListFiles = ListFiles & vbCrLf & "- ..."
            Else
                ListFiles = ListFiles & vbCrLf & "- " & MyItem.Text
            End If
            i = i + 1
            If i > 3 Then Exit For
        Next
        Response = MsgBox(mMessage(42) & vbCrLf & ListFiles, MsgBoxStyle.YesNo + MsgBoxStyle.Information)    'Do you confirm the deletion of selected documents from ZAAN\copy directory ?
        If Response = MsgBoxResult.Yes Then
            Me.Cursor = Cursors.WaitCursor                           'sets wait cursor
            fswZaan.EnableRaisingEvents = False
            For Each MyItem In lvTemp.SelectedItems
                FileName = mZaanCopyPath & MyItem.Text & MyItem.SubItems(1).Text     'get document name & type
                Try
                    My.Computer.FileSystem.DeleteFile(FileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation)      'displays error message
                End Try
            Next
            fswZaan.EnableRaisingEvents = True
            Me.Cursor = Cursors.Default                              'resets default cursor
            Call DisplayTempFiles()                                  'refreshes the display of all "temp"/"copy" files
        End If
        lvTemp.Focus()
    End Sub

    Private Sub DeleteTreeFile(ByVal FileName As String)
        'Deletes tree file of given name
        Dim dirInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "tree\")
        Dim fi As System.IO.FileInfo

        For Each fi In dirInfo.GetFiles(FileName)
            'Debug.Print("DeleteTreeFile : " & fi.Name)     'TEST/DEBUG
            fi.Delete()                                    'deletes tree file
        Next
    End Sub

    Private Sub DeleteTrvWSelNode()
        'Deletes selected tree node (if enabled)
        Dim dirInfo As System.IO.DirectoryInfo
        Dim TreeCode, ParentNodeTag, DataDirPathName As String

        'Debug.Print("DeleteTrvWSelNode...")      'TEST/DEBUG
        TreeCode = Mid(trvW.SelectedNode.Tag, 2, 1)
        If TreeCode = "u" Then                                                 'case of an Access control node => creates a "data\data_new" directory
            fswData.EnableRaisingEvents = False                                'locks fswData related events
            'dirInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "data\data_" & trvW.SelectedNode.Text)
            DataDirPathName = Mid(trvW.SelectedNode.Tag, 3)
            dirInfo = My.Computer.FileSystem.GetDirectoryInfo(DataDirPathName)
            dirInfo.Delete()
            fswData.EnableRaisingEvents = True                                 'unlocks fswData related events
        Else
            fswTree.EnableRaisingEvents = False                                'locks fswTree related events
            If (TreeCode = "t") And (Len(trvW.SelectedNode.Text) = 4) Then     'in v3 database format, a year tree file can be deleted
                Call DeleteTreeFile(trvW.SelectedNode.Tag & "_t" & mTreeRootKey & "__" & trvW.SelectedNode.Text & ".txt")   'deletes year tree file
            Else
                Call DeleteTreeFile(trvW.SelectedNode.Tag)                     'deletes tree file of given name
            End If
            fswTree.EnableRaisingEvents = True                                 'unlocks fswTree related events
        End If
        ParentNodeTag = trvW.SelectedNode.Parent.Tag                           'saves parent node tag (for becoming next selected node)
        Call LoadTrees()                                                       'loads all trees of current ZAAN database
        ExpandTreeNode("", ParentNodeTag)                                      'selects parent node and expands tree view (that has been re-loaded) at this node
    End Sub

    Private Sub OpenMailCopyFiles()
        'Open mail application to select copy file(s) to be sent
        Dim MailSubject0, MailSubject As String
        'Dim AttachedFile As String
        Dim itmX As ListViewItem

        MailSubject0 = "Doc : "
        MailSubject = MailSubject0
        For Each itmX In lvTemp.SelectedItems
            If MailSubject <> MailSubject0 Then MailSubject = MailSubject & ", "
            MailSubject = MailSubject & itmX.Text & itmX.SubItems(1).Text
        Next
        OpenDocument("mailto:?SUBJECT=" & MailSubject)
        'OpenDocument("mailto:?SUBJECT=" & MailSubject & "&BODY=<please select attachments>")
        'AttachedFile = mZaanCopyPath & lvTemp.SelectedItems(0).Text & lvTemp.SelectedItems(0).SubItems(1).Text  'TEST/DEBUG
        'Debug.Print("OpenMailCopyFiles:  AttachedFile = " & AttachedFile)                                'TEST/DEBUG
        'OpenDocument("mailto:?SUBJECT=" & MailSubject & "&ATTACHMENT=" & AttachedFile)                   'TEST/DEBUG
    End Sub

    Private Sub SendMailWithAttachments()
        'Call mail system for sending selected documents
        Dim MailSubject, MailAttachment, DirName, FileName As String
        Dim i As Integer
        'Dim item As Attachment

        MailSubject = lvIn.SelectedItems(0).Text
        If lvIn.SelectedItems.Count > 1 Then MailSubject = MailSubject & "..."

        MailAttachment = ""
        For i = 0 To lvIn.SelectedItems.Count - 1
            DirName = lvIn.SelectedItems(i).Tag            'get dir name of selected file
            FileName = lvIn.SelectedItems(i).Text & lvIn.SelectedItems(i).SubItems(8).Text
            MailAttachment = MailAttachment & mZaanDbPath & "data\" & DirName & "\" & FileName '& ";"
        Next
        Debug.Print("MailAttachment =>" & MailAttachment & "<")  'TEST/DEBUG

        'message.Attachments.Add(New System.Net.Mail.Attachment(MailAttachment, System.Net.Mime.TransferEncoding.Base64))
        'message.Attachments.Add(New Attachment(MailAttachment,"")
        'Dim mailClient As New SmtpClient

        Dim message As New MailMessage("emmanuel.derome@zaan.com", "emmanuel.derome@zaan.com", MailSubject, "Test Zaan")   'TEST/DEBUG
        Dim mailClient As New SmtpClient
        Dim attach As New Attachment(MailAttachment)    'TEST/DEBUG with 1 selected attachment only
        Try
            message.IsBodyHtml = True
            mailClient.Host = "smtp.zaan.com"    'TEST/DEBUG
            mailClient.Port = 25                 'TEST/DEBUG
            mailClient.UseDefaultCredentials = True
            'mailClient.DeliveryMethod = SmtpDeliveryMethod.PickupDirectoryFromIis
            mailClient.DeliveryMethod = SmtpDeliveryMethod.Network
            'message.Attachments.Add(attach)
            mailClient.Send(message)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)              'displays error message
            'Debug.Print(ex.Message)              'TEST/DEBUG
        End Try
    End Sub

    Private Sub UpdateCutCopyPasteMenusItems()
        'Updates the enabling status of cut/copy/paste items of lists context menus
        If Forms.Clipboard.ContainsFileDropList Then             'clipboard contains a file drop list
            tsmiLvInPaste.Enabled = True
            tsmiLvOutPaste.Enabled = True
            tsmiLvTempPaste.Enabled = True
        Else
            tsmiLvInPaste.Enabled = False
            tsmiLvOutPaste.Enabled = False
            tsmiLvTempPaste.Enabled = False
        End If
    End Sub

    Private Sub SetMenuItemIfFileTypeExists(ByRef MenuItem As ToolStripItem, ByVal FileType1 As String, Optional ByVal FileType2 As String = "")
        'Sets given menu item visibility and tag with related file type if exists and updates related icon in imgFileTypes
        Dim FileName As String = mMessage(74)
        MenuItem.Enabled = False                           'disables menu item by default

        If FileType1 <> "" Then
            Dim FileTypeOut1 As String = GetFileTypeAndImage(FileName & FileType1)
            If FileTypeOut1 = ".zqm" Then                  'case of an unknown file type 1
                If FileType2 <> "" Then
                    Dim FileTypeOut2 As String = GetFileTypeAndImage(FileName & FileType2)
                    If FileTypeOut2 <> ".zqm" Then         'case of a known file type 2 : related application exists
                        MenuItem.Enabled = True                                'enables menu item
                        MenuItem.Tag = LCase(Mid(FileTypeOut2, 2))             'stores application file type without leading "."
                        MenuItem.Image = imgFileTypes.Images(FileTypeOut2)     'stores application icon
                    End If
                End If
            Else                                           'case of a known file type 1 : related application exists
                MenuItem.Enabled = True                                        'enables menu item
                MenuItem.Tag = LCase(Mid(FileTypeOut1, 2))                     'stores application file type without leading "."
                MenuItem.Image = imgFileTypes.Images(FileTypeOut1)             'stores application icon
            End If
        End If
    End Sub

    Private Sub UpdateNewSubMenuItems()
        'Updates lvIn "New" sub menu items depending on available MS Office applications
        'Debug.Print("UpdateNewSubMenuItems...")    'TEST/DEBUG

        tsmiLvInNewNote.Text = mMessage(13)                'Note
        tsmiLvInNewText.Text = mMessage(14)                'Text
        tsmiLvInNewPres.Text = mMessage(15)                'Presentation
        tsmiLvInNewSpSh.Text = mMessage(16)                'Spreadsheet

        Call SetMenuItemIfFileTypeExists(tsmiLvInNewNote, ".txt")              'sets given menu item visibility and tag with related file type if exists and updates related icon in imgFileTypes
        Call SetMenuItemIfFileTypeExists(tsmiLvInNewText, ".docx", ".doc")     'sets given menu item visibility and tag with related file type if exists and updates related icon in imgFileTypes
        Call SetMenuItemIfFileTypeExists(tsmiLvInNewPres, ".pptx", ".ppt")     'sets given menu item visibility and tag with related file type if exists and updates related icon in imgFileTypes
        Call SetMenuItemIfFileTypeExists(tsmiLvInNewSpSh, ".xlsx", ".xls")     'sets given menu item visibility and tag with related file type if exists and updates related icon in imgFileTypes
    End Sub

    Private Sub CutCopyLvInSelection(ByVal IsCut As Boolean)
        'Depending on IsCut, Cut or Copy selected files of "in" list view into Clipboard and updates lists context menus
        Dim DirPathName, FileName As String
        Dim MyItem As ListViewItem
        Dim MyFiles(lvIn.SelectedItems.Count - 1) As String
        Dim i As Integer = 0
        Dim ClipData As New Forms.DataObject()

        For Each MyItem In lvIn.SelectedItems
            'DirPathName = mZaanDbPath & "data\" & MyItem.Tag      'get dir name of selected file
            DirPathName = MyItem.Tag                            'get dir name of selected file
            FileName = MyItem.Text & MyItem.SubItems(8).Text
            Call DeleteFileImage(DirPathName, FileName)         'deletes corresponding thumbnail image file (zzi*) that may have been generated by ZAAN
            MyFiles(i) = DirPathName & "\" & FileName           'adds the file to file list
            i = i + 1
        Next
        Forms.Clipboard.Clear()                                       'removes all data from Clipboard
        ClipData.SetData("FileDrop", True, MyFiles)
        If IsCut Then
            ClipData.SetData("Preferred DropEffect", mStreamMove)   'sets drop effect to Move (compatible with Windows File Explorer)
        Else
            ClipData.SetData("Preferred DropEffect", mStreamCopy)   'sets drop effect to Copy (compatible with Windows File Explorer)
        End If
        Forms.Clipboard.SetDataObject(ClipData)                       'writes the special ClipData object in Clipboard
    End Sub

    Private Sub CutCopyListViewSelection(ByVal IsCut As Boolean, ByVal IsOutList As Boolean)
        'Depending on IsCut, Cut or Copy selected files of given "out" or "temp" list view into Clipboard and updates lists context menus
        Dim FileName As String
        Dim MyItem As ListViewItem
        Dim MyFiles() As String
        Dim i As Integer = 0
        Dim ClipData As New Forms.DataObject()

        If IsOutList Then                                       'case of (import/export) "out" list
            ReDim MyFiles(lvOut.SelectedItems.Count - 1)
            For Each MyItem In lvOut.SelectedItems
                FileName = MyItem.Text & MyItem.SubItems(1).Text     'get document name & type
                MyFiles(i) = mZaanImportPath & FileName              'adds the ListViewItem to the array of ListViewItems
                i = i + 1
            Next
        Else                                                    'case of (copy) "temp" list
            ReDim MyFiles(lvTemp.SelectedItems.Count - 1)
            For Each MyItem In lvTemp.SelectedItems
                FileName = MyItem.Text & MyItem.SubItems(1).Text     'get document name & type
                MyFiles(i) = mZaanCopyPath & FileName                'adds the ListViewItem to the array of ListViewItems
                i = i + 1
            Next
        End If

        Forms.Clipboard.Clear()                                       'removes all data from Clipboard
        ClipData.SetData("FileDrop", True, MyFiles)
        If IsCut Then
            ClipData.SetData("Preferred DropEffect", mStreamMove)   'sets drop effect to Move (compatible with Windows File Explorer)
        Else
            ClipData.SetData("Preferred DropEffect", mStreamCopy)   'sets drop effect to Copy (compatible with Windows File Explorer)
        End If
        Forms.Clipboard.SetDataObject(ClipData)                       'writes the special ClipData object in Clipboard
    End Sub

    Private Function IsMoveEffect() As Boolean
        'Returns true if first byte of "Preferred DropEffect" data in Clipboard is set to move effect value

        Dim ClipData As Forms.DataObject = Forms.Clipboard.GetDataObject()
        Dim CurStream As Stream = ClipData.GetData("Preferred DropEffect")
        Dim DropEffectData(3) As Byte
        Dim isMove As Boolean = False

        If Not CurStream Is Nothing Then
            CurStream.Read(DropEffectData, 0, DropEffectData.Length)
            If DropEffectData(0) = mMoveEffectData(0) Then      'compares first byte data array containing Move or Copy DropEffect
                isMove = True                                   'drop effect has been set to move
            End If
        End If
        'Debug.Print("isMoveEffect = " & isMove)    'TEST/DEBUG
        IsMoveEffect = isMove
    End Function

    Private Sub PasteFilesToFilePath(ByVal FilePath As String)
        'Move or copy Clipboard files at given (list) file path depending on embedded DropEffect, clears Clipboard and updates lists context menus
        Dim i As Integer
        Dim FileName As String

        If Not CheckMoveInLimitIsOk() Then Exit Sub 'case of ZAAN Basic licence with the limit of 1000 documents stored in current ZAAN database reached

        If Forms.Clipboard.ContainsFileDropList Then
            Dim ClipData As Forms.DataObject = Forms.Clipboard.GetDataObject()

            If Forms.Clipboard.ContainsData(Forms.DataFormats.FileDrop) Then
                Dim MyFiles() As String = ClipData.GetData(Forms.DataFormats.FileDrop)
                Dim isMove As Boolean = IsMoveEffect()                       'true if first byte of "Preferred DropEffect" data in Clipboard is set to move effect value

                If Not MyFiles Is Nothing Then
                    If isMove Then Call ResetFileMovePile() 'resets file move pile (used for "Undo move") and disables related local menu control
                    For i = 0 To MyFiles.Length - 1
                        FileName = System.IO.Path.GetFileName(MyFiles(i))
                        If isMove Then
                            'Debug.Print("tsmiLvOutPaste_Click:  MOVE file = " & FileName)    'TEST/DEBUG
                            Call MoveSelectedFile(MyFiles(i), "", FilePath & FileName, False)    'moves file
                        Else
                            'Debug.Print("tsmiLvOutPaste_Click:  COPY file = " & FileName)    'TEST/DEBUG
                            Call CopySelectedFile(MyFiles(i), "", FilePath & FileName)    'copies file
                        End If
                    Next
                End If

            ElseIf Windows.Clipboard.ContainsData("FileGroupDescriptor") Then              'case of an Outlook message to be dropped
                Call DropOutlookFiles(ClipData, FilePath, lvIn)     'drops given Outlook files contained OutlookDataObject eData at destination directory
            End If
        End If
        Forms.Clipboard.Clear()                                            'removes all data from Clipboard
    End Sub

    Private Sub OpenZaanDataDirectory(ByVal ZaanDataDirPathName As String)
        'Opens given ZAAN data directory, updates mFileFilter and displays Selector and related Selected files
        Dim NewZaanDb, NewDataDir, NewFileFilter As String
        Dim ZaanDataTab As String() = Split(ZaanDataDirPathName, "\data\")

        'Debug.Print("1 - OpenZaanDataDirectory :  ZaanDataDirPathName = " & ZaanDataDirPathName & "  old mFileFilter = " & mFileFilter)    'TEST/DEBUG

        If ZaanDataTab.Length > 1 Then
            NewZaanDb = ZaanDataTab(0) & "\"                         'gets ZAAN database path name
            NewDataDir = ZaanDataTab(1)                              'gets data directory
            NewFileFilter = NewDataDir.Replace("_", "*_")            'builds corresponding mFileFilter
            'Debug.Print("2 - OpenZaanDataDirectory :  NewZaanDb = " & NewZaanDb & "  NewDataDir = " & NewDataDir & " => NewFileFilter =" & NewFileFilter)    'TEST/DEBUG

            Call UpdateZaanDbRootPathName(NewZaanDb)                 'updates current database selection with mZaanDbRoot, mZaanDbPath and mZaanDbName

            mFileFilter = NewFileFilter                              'sets mFileFilter to given position
            Call InitFileDataWatcher()                               'initializes "data","tree" and "xin" file watchers to detect directory access/change in current ZAAN database
            Call InitTreeInOutFiles(False)                           'displays selected Zaan database name, logo, tree, selector and selected files views , but not "imput", "out" and "temp" views
        Else
            Debug.Print("Cannot open ZAAN data directory corresponding to : " & ZaanDataDirPathName)   'TEST/DEBUG
        End If
    End Sub

    Private Sub StopVideoPlayer()
        'Stops VLC video player (which activity should be indicated by tmrVideoProgress activity)

        'VLC2Zoom.playlist.stop()                                     'stops VLC playlist

        btnVideoPlayPause.Image = My.Resources.play                  'updates play/pause button image to play
        tmrVideoProgress.Enabled = False                             'stops video progress timer
        pgbVideo.Value = 0
    End Sub

    Private Sub CheckZoomDisplay(ByVal SelPathFile As String, ByVal RootFileName As String, ByVal FileType As String)
        'Resets zoom displays and if zoom view activated, selects display control relative to given file type and shows it
        Dim FileName, UCFileType As String

        If Microsoft.VisualBasic.Right(SelPathFile, 1) = "\" Then
            SelPathFile = Microsoft.VisualBasic.Left(SelPathFile, Len(SelPathFile) - 1)
        End If
        FileName = RootFileName & FileType
        'Debug.Print("CheckZoomDisplay :  SelPathFile = " & SelPathFile & "  FileName = " & FileName)    'TEST/DEBUG

        UCFileType = UCase(FileType)

        mSlideShowPathFileName = ""                                  'sets file name to no image/slide by default
        If tmrVideoProgress.Enabled Then                             'case of a playing video
            Call StopVideoPlayer()                                   'stops VLC video player (which activity should be indicated by tmrVideoProgress activity)
        End If
        If Not SplitContainer3.Panel2Collapsed Then                  'case of visible viewer panel : select appropriate display control
            Select Case UCFileType
                Case ".JPG", ".GIF", ".PNG", ".BMP"                  'displays image
                    pnlZoom.Visible = True
                    wbDoc.Visible = False
                    'aapDoc.Visible = False
                    pnlZoomVideo.Visible = False
                    mSlideShowPathFileName = SelPathFile & "\" & FileName  'stores current valid image/slide file name
                    pctZoom.Load(mSlideShowPathFileName)                   'loads picture from related file

                    'Case ".MPG"                                          'displays video
                    '  pnlZoomVideo.Visible = True
                    '  wbDoc.Visible = False
                    '  aapDoc.Visible = False
                    '  pnlZoom.Visible = False
                    '  VLC2Zoom.playlist.items.clear()                  '(asynchronously) clears VLC playlist (and stops any running video)
                    '  mVideoPlayListClearing = True                    'sets flag indicating that video playlist is clearing
                    '  mVideoPathFileName = SelPathFile & "\" & FileName      'stores current video file name
                    '  tmrVideoProgress.Enabled = True                  'starts video progress timer

                    'Case ".PDF"
                    '  aapDoc.Visible = True                            'displays Adobe Acrobat PDF reader
                    '  wbDoc.Visible = False
                    '  pnlZoom.Visible = False
                    '  pnlZoomVideo.Visible = False
                    '  aapDoc.src = SelPathFile & "\" & FileName        'sets Adobe Acrobat Reader source to document file name
                    '  If aapDoc.Width < 50 Then
                    '    SplitContainer3.SplitterDistance = SplitContainer3.SplitterDistance + 1    'forces resizing of child viewer that is "docked" into (DockStyle.Fill)
                    '    SplitContainer3.SplitterDistance = SplitContainer3.SplitterDistance - 1    'moves back splitter to initial position
                    '  End If
                Case Else                                            'displays Web browser (for pdf, txt... files)
                    wbDoc.Dock = DockStyle.Fill                      'extends Web browser control (initially invisible) in zoom panel
                    wbDoc.Visible = True
                    'aapDoc.Visible = False
                    pnlZoom.Visible = False
                    pnlZoomVideo.Visible = False
                    'lblDocName.Text = SelPathFile & "\" & FileName   'sets browser url to document file name
                    Try
                        wbDoc.Navigate(SelPathFile & "\" & FileName, False)    'open document in same browser (NewWindow = False)
                    Catch ex As Exception
                        Debug.Print(ex.Message)                        'file not found !
                    End Try

            End Select
        End If
    End Sub

    Private Function MaxDocCountIsReached(ByVal DirPath As String, ByRef DocCount As Integer) As Boolean
        'Returns true if maximum limit of stored document in given data dir path/filter has been reached
        Dim dirInfo, dirSel As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim FileName, Prefix As String
        Dim isHiddenFile As Boolean
        Dim MaxIsReached As Boolean = False

        dirInfo = My.Computer.FileSystem.GetDirectoryInfo(DirPath)
        For Each dirSel In dirInfo.GetDirectories("_*", SearchOption.TopDirectoryOnly)   'searches for ZAAN cube directories
            For Each fi In dirSel.GetFiles("*.*")
                FileName = fi.Name
                Prefix = Microsoft.VisualBasic.Left(FileName, 3)
                isHiddenFile = False
                If (fi.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then isHiddenFile = True
                If Microsoft.VisualBasic.Left(FileName, 1) = "~" Then isHiddenFile = True
                If FileName <> "." And FileName <> ".." And Prefix <> "zzi" And UCase(FileName) <> "THUMBS.DB" And Not isHiddenFile Then   'ignores current and encompassing directories, ZAAN and Windows XP thumbnail images and hidden files
                    DocCount = DocCount + 1                                               'increments file counter
                    'Debug.Print("CheckMoveInLimitIsOk : " & DocCount & " : " & fi.FullName)  'TEST/DEBUG
                    If DocCount > mMaxBasicDocToFile - 1 Then                            'checks doc count against max ZAAN Basic doc to file
                        MaxIsReached = True
                        MsgBox(mMessage(75), MsgBoxStyle.Exclamation)                    'Sorry, you cannot store more than 1 000 documents with ZAAN Basic (free limited license) !
                        Exit For
                    End If
                End If
            Next
            If MaxIsReached Then Exit For
        Next
        MaxDocCountIsReached = MaxIsReached
    End Function

    Private Function CheckMoveInLimitIsOk() As Boolean
        'Returns false if, with ZAAN Basic licence, the limit of 1000 documents stored in current ZAAN database has been reached
        Dim DocCount As Integer = 0
        Dim LimitIsOK As Boolean = True
        Dim DataPath As String

        DataPath = mZaanDbPath & "data"                              'sets default data directory
        If mLicTypeCode < 20 Then                                    'case of (free) limited ZAAN-Basic license
            LimitIsOK = False
            If Not MaxDocCountIsReached(DataPath, DocCount) Then     'checks doc count in main data directory 
                'NOTE : DOES NOT NEED TO CHECK ACCESS CONTROL/GROUP DIRECTORIES, AS THERE ARE NOT VISIBLE FROM ZAAN-BASIC !
                LimitIsOK = True                                     'maximum number of documents has not been reached
            End If
        End If
        CheckMoveInLimitIsOk = LimitIsOK
    End Function

    Public Sub ExportDatabaseToWindows(Optional ByVal Auto As Boolean = False)
        'Exports current ZAAN database into a mono-dimensional Windows tree organization and document copies to directories related to their first dimension available
        Dim Response As MsgBoxResult = MsgBoxResult.Yes
        Dim ExportPathName, InfoDirPathName As String

        If Not Auto Then                                                  'case of Manual mode : ask then for confirmation
            Response = MsgBox(mMessage(113), MsgBoxStyle.YesNo + MsgBoxStyle.Information)    'Do you confirm the export of current ZAAN database to a simple Windows tree organization and the copy of all documents ?
        End If
        If Response = MsgBoxResult.Yes Then
            If Microsoft.VisualBasic.Right(mExportDBwinDest, 1) = "\" Then
                ExportPathName = mExportDBwinDest & mZaanDbName & "_exp" & Format(Now, "yyMMdd-HHmm")
            Else
                ExportPathName = mExportDBwinDest & "\" & mZaanDbName & "_exp" & Format(Now, "yyMMdd-HHmm")
            End If

            Me.Cursor = Cursors.WaitCursor                                'sets wait cursor
            fswTree.EnableRaisingEvents = False                           'locks fswTree related events
            fswData.EnableRaisingEvents = False                           'locks fswData related events
            fswXin.EnableRaisingEvents = False                            'locks fswXin related events

            pgbZaan.Value = 0                                             'resets progress bar
            pgbZaan.Visible = True                                        'shows progress bar

            Call ExportTreesToWindows(ExportPathName, True)               'exports trees of current ZAAN database to a mono-dimensional Windows tree organization
            If mTreeCodeSeries = "" Then
                Call ExportDocCopiesToWindows(ExportPathName, mZaanDbPath & "data")      'exports document copies of current ZAAN database to directories related to their first dimension available
            Else
                Call ExportDataToWindows(ExportPathName & "\data", mZaanDbPath & "data") 'exports/copies files from given data path into a Windows mixted tree organisation depending on given tree codes series
                InfoDirPathName = ExportPathName & "\info"
                If CreateDirIfNotExistsStatus(InfoDirPathName) = 1 Then   'info directory just created : will enable exported database to be used as a ZAAN database without keys (mTreeKeyLength =0)
                    Call CreateZaanLogoIfNoTopLeftFile(InfoDirPathName)
                End If
                Call ExportInfoIniFiles(InfoDirPathName)                  'exports info\zaan_*.ini files of current ZAAN database into v4 format (no more keys)
            End If

            pgbZaan.Value = pgbZaan.Maximum                               'makes sure that progress bar is set to maximum
            If Not Auto Then                                              'case of Manual mode : inform user of export termination
                MsgBox(mMessage(114) & ExportPathName, MsgBoxStyle.Information)    'Current ZAAN database has been successfully exported to : <export name>
            End If

            'If mXmode = "auto" Then fswXin.EnableRaisingEvents = True   'RETIRED unlocks fswXin related events if auto Xchange mode selected
            fswData.EnableRaisingEvents = True                            'unlocks fswData related events
            fswTree.EnableRaisingEvents = True                            'unlocks fswTree related events
            Me.Cursor = Cursors.Default                                   'resets default cursor
            pgbZaan.Visible = False                                       'hides progress bar
        End If
    End Sub

    Private Sub trvW_AfterLabelEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.NodeLabelEditEventArgs) Handles trvW.AfterLabelEdit
        Dim OldFileName, NewFileName, NewTitle, FilePath, FileType, TreeCode, OldDataDirPathName, OldDataDirName, OldDataDirRoot, NewDataDirName As String
        Dim isCancelled As Boolean = False
        Dim isNewChildNode As Boolean = True
        Dim ParentNode, NodeTest As TreeNode
        Dim p As Integer

        'Debug.Print("trvW_AfterLabelEdit...")   'TEST/DEBUG
        fswTree.EnableRaisingEvents = False                          'locks fswTree related events

        If e.Label Is Nothing Then
            fswTree.EnableRaisingEvents = True                       'unlocks fswTree related events
            Exit Sub
        End If

        ParentNode = e.Node.Parent
        'Debug.Print("trvW_AfterLabelEdit :  e.Label = " & e.Label & "  e.Node.Text = " & e.Node.Text)   'TEST/DEBUG
        NewTitle = e.Label
        If NewTitle.Length = 0 Then                                  'entered label is empty
            isCancelled = True
        Else
            If NewTitle.IndexOfAny(New Char() {"<"c, ">"c, ":"c, """"c, "/"c, "\"c, "|"c, "?"c, "*"c, "&"c}) = -1 Then   'entered characters are valid (for directory name)
                For Each NodeTest In ParentNode.Nodes                'checks if child node does not exist with same title
                    If NodeTest.Text = NewTitle Then
                        isNewChildNode = False
                        isCancelled = True
                        Exit For
                    End If
                Next
            Else
                isCancelled = True                                   'entered label has no valid characters
            End If
        End If

        If isNewChildNode And Not isCancelled Then                   'entered label is new and has valid characters
            TreeCode = Mid(e.Node.Tag, 2, 1)
            If TreeCode = "u" Then                                   'case of an Access control node => creates a "data\data_new" directory
                'OldDataDirPathName = mZaanDbPath & "data\data_" & e.Node.Text
                OldDataDirPathName = Mid(e.Node.Tag, 3)
                OldDataDirName = GetLastDirName(OldDataDirPathName & "\")
                OldDataDirRoot = Mid(OldDataDirPathName, 1, Len(OldDataDirPathName) - Len(OldDataDirName))

                'p = InStr(OldDataDirName, "_")
                p = InStr(OldDataDirName, "data_")
                If p > 0 Then
                    'NewDataDirName = Mid(OldDataDirName, 1, p) & NewTitle
                    NewDataDirName = Mid(OldDataDirName, 1, p + 4) & NewTitle
                    fswData.EnableRaisingEvents = False                  'locks fswData related events
                    Try
                        If UCase(NewTitle) = UCase(e.Node.Text) Then
                            My.Computer.FileSystem.RenameDirectory(OldDataDirPathName, "ZAAN_TEMP")
                            'My.Computer.FileSystem.RenameDirectory(mZaanDbPath & "data\" & "ZAAN_TEMP", NewDataDirName)
                            My.Computer.FileSystem.RenameDirectory(OldDataDirRoot & "ZAAN_TEMP", NewDataDirName)
                        Else
                            My.Computer.FileSystem.RenameDirectory(OldDataDirPathName, NewDataDirName)
                        End If
                        trvW.SelectedNode.Text = NewTitle
                        'trvW.SelectedNode.Tag = "_u" & NewTitle
                        trvW.SelectedNode.Tag = "_u" & OldDataDirRoot & NewDataDirName

                        Call UpdateFileFilterAndDisplay()                'updates tree buttons and mFileFilter with last tree selection and display selected files if new filter
                        Call InitDisplaySelectedFiles()                  'displays files selected with filter (NECESSARY because UpdateFileFilterAndDisplay() has seen no mFileFilter change !!!)

                        'e.CancelEdit = True                              'cancels node label edition
                        'Call LoadTrees()                                 '(re)loads all trees of current ZAAN database
                        'ExpandTree("u", "*_u" & NewTitle)                'selects tree node at given key, expands tree at node position, updates selector/mFileFilter and related selection if new filter

                    Catch ex As Exception
                        e.CancelEdit = True                              'cancels node label edition
                        MsgBox(ex.Message, MsgBoxStyle.Exclamation)      'displays error message
                    End Try
                    fswData.EnableRaisingEvents = True                   'unlocks fswData related events
                End If
            Else                                                     'case of a Who, What, When, Where, What else or Other node
                If TreeCode <> "t" Then
                    OldFileName = System.IO.Path.GetFileNameWithoutExtension(e.Node.Tag)
                    NewFileName = Microsoft.VisualBasic.Left(OldFileName, 4 + 2 * mTreeKeyLength) & NewTitle
                    FilePath = mZaanDbPath & "tree\"
                    FileType = ".txt"
                    Try
                        If UCase(NewFileName) = UCase(OldFileName) Then
                            My.Computer.FileSystem.RenameFile(FilePath & OldFileName & FileType, "ZAAN_TEMP" & FileType)
                            My.Computer.FileSystem.RenameFile(FilePath & "ZAAN_TEMP" & FileType, NewFileName & FileType)
                        Else
                            My.Computer.FileSystem.RenameFile(FilePath & OldFileName & FileType, NewFileName & FileType)
                        End If
                        trvW.SelectedNode.Text = NewTitle
                        trvW.SelectedNode.Tag = NewFileName & ".txt"
                        Call UpdateFileFilterAndDisplay()            'updates tree buttons and mFileFilter with last tree selection and display selected files if new filter
                        Call InitDisplaySelectedFiles()              'displays files selected with filter (NECESSARY because UpdateFileFilterAndDisplay() has seen no mFileFilter change !!!)
                    Catch ex As Exception
                        e.CancelEdit = True                          'cancels node label edition
                        MsgBox(ex.Message, MsgBoxStyle.Exclamation)  'displays error message
                    End Try
                End If
            End If
        End If

        If isCancelled Then
            e.CancelEdit = True                                      'cancels node label edition
            If isNewChildNode Then
                MsgBox(mMessage(79), MsgBoxStyle.Exclamation)        'Entry not allowed !
            Else
                MsgBox(mMessage(157), MsgBoxStyle.Exclamation)       'This folder name already exists !
            End If
            e.Node.BeginEdit()                                       'goes back to node label edition mode
        End If
        fswTree.EnableRaisingEvents = True                           'unlocks fswTree related events
    End Sub

    Private Sub trvW_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles trvW.AfterSelect
        'Debug.Print("trvW_AfterSelect-1: e.Node Text/Tag/ImageKey=" & e.Node.Text & "/" & e.Node.Tag & "/" & e.Node.ImageKey)      'TEST/DEBUG

        tsmiTrvWDelete.Enabled = False                     'CheckNodeDelStatus() is called when opening local menu
        If mTreeClicked Then
            'Debug.Print("trvW_AfterSelect-2: e.Node Text/Tag/ImageKey=" & e.Node.Text & "/" & e.Node.Tag & "/" & e.Node.ImageKey)      'TEST/DEBUG
            mTreeClicked = False                           'resets flag for avoiding calling twice UpdateFileFilterAndDisplay() done below
            trvW.SelectedImageKey = e.Node.ImageKey        'forces correct image display of selected node => WARNING : this will call again trvW_AfterSelect() event !!!
            Call UpdateAfterTreeNodeSelection()            'updates mFileFilter and related selector/selection and makes sure that lvIn is visible
        End If
    End Sub

    Private Sub lvIn_AfterLabelEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.LabelEditEventArgs) Handles lvIn.AfterLabelEdit
        'Renames first selected "in" document and related image file (if any)
        Dim OldFileName, NewFileName, FilePath, FileType As String
        Dim isCancelled As Boolean = False

        If e.Label Is Nothing Then Exit Sub
        If e.Label.Length > 0 Then                                   'node label is not empty
            NewFileName = e.Label
            OldFileName = lvIn.SelectedItems(0).Text
            If NewFileName = OldFileName Then Exit Sub

            If e.Label.IndexOfAny(New Char() {"<"c, ">"c, ":"c, """"c, "/"c, "\"c, "|"c, "?"c, "*"c}) = -1 Then   'entered characters are valid (for file name)
                'FilePath = mZaanDbPath & "data\" & lvIn.SelectedItems(0).Tag
                FilePath = lvIn.SelectedItems(0).Tag                 'gets absolute path of selected file
                FileType = lvIn.SelectedItems(0).SubItems(8).Text
                If RenameFileNameOK(FilePath, OldFileName, NewFileName, FileType) Then   'returns true if successfully renamed given old file name with new file name and related thumbnail image file (if any), using given file path and type
                    lvIn.SelectedItems(0).Text = NewFileName         'updates selected item
                Else
                    e.CancelEdit = True                              'cancels item label edition
                End If
            Else
                isCancelled = True
            End If
        Else
            isCancelled = True
        End If
        If isCancelled Then
            e.CancelEdit = True                                      'cancels item label edition
            MsgBox(mMessage(79), MsgBoxStyle.Exclamation)            'Entry not allowed !
            If lvIn.SelectedItems.Count > 0 Then
                lvIn.SelectedItems(0).BeginEdit()                    'goes back to label edition mode at first selected item
            End If
        End If
    End Sub

    Private Sub lvIn_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvIn.ColumnClick
        'Changes "in" files list order (ascending/descending) depending on selected column
        Dim i As Integer

        'Debug.Print("lvIn_ColumnClick : col = " & e.Column & "  mlvInOrder : " & mlvInOrder(e.Column))    'TEST/DEBUG
        lvIn.Sorting = SortOrder.None                      'cancels automatic ordering mode (VERY IMPORTANT FOR ENABLING COMPARER TO WORK !)
        If mlvInOrder(e.Column) <= 0 Then
            mlvInOrder(e.Column) = 1                       'set ascending order flag
            lvIn.ListViewItemSorter = New ListViewItemComparer(e.Column)
        Else
            mlvInOrder(e.Column) = -1                      'set descending order flag
            lvIn.ListViewItemSorter = New ListViewItemRevComparer(e.Column)
        End If
        For i = 0 To mlvInOrder.Length - 1
            If i <> e.Column Then
                mlvInOrder(i) = 0                          'set other columns flags to no order
            End If
        Next
    End Sub

    Private Sub lvIn_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnWidthChangedEventArgs) Handles lvIn.ColumnWidthChanged
        Dim ColIndex As Integer = e.ColumnIndex

        If ColIndex = 0 Then
            'Debug.Print("lvIn_ColumnWidthChanged : check col.width = " & lvIn.Columns(ColIndex).Width)   'TEST/DEBUG
            If lvIn.Columns(0).Width < 300 Then            'makes sure that first column is wide enough for correct display of related search panel
                lvIn.Columns(0).Width = 300
            End If
        End If
        Call ResizeSelectorLabel(ColIndex)                 'resizes selector label of given column index corresponding to selector list

        If (mLicTypeCode < 30) And ((ColIndex = 1) Or (ColIndex = 6) Or (ColIndex = 7)) Then   'case of "Basic" and "First" licenses
            If lvIn.Columns(ColIndex).Width > 0 Then
                lvIn.Columns(ColIndex).Width = 0           'makes sure that Access control, Status and Action for columns are hidden
                'Debug.Print("lvIn_ColumnWidthChanged : hides col = " & ColIndex)   'TEST/DEBUG
            End If
        End If
    End Sub

    Private Sub lvIn_ColumnWidthChanging(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnWidthChangingEventArgs) Handles lvIn.ColumnWidthChanging
        Call ResizeSelectorLabel(e.ColumnIndex)            'resizes selector label of given column index corresponding to selector list
    End Sub

    Private Sub lvIn_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvIn.DragDrop
        'Terminates drag-drop operation of files to be moved in current "in" selection
        Dim i As Integer
        Dim FileName, DestFilePathName, SelectedDirPath As String
        Dim Response As MsgBoxResult
        Dim CubeImport As Boolean

        trvW.HideSelection = True                                              'disable tree node selection highlight even if tree has not the focus

        If e.Data.GetDataPresent(Forms.DataFormats.FileDrop) Then                    'case of file(s) to be dropped
            If mStartDragLvInItem Then                                         'item drag started from lvIn => cancels drop operation
                mStartDragLvInItem = False
                'clears clipboard TBD...
                Exit Sub                                                       'it is not possible to drop document(s) to source list (NOTE : risk is to drop at a different cube by mistake !)
            End If
            mStartDragLvInItem = False                                         'makes sure that lvIn item drag flag is reset
            If Not CheckMoveInLimitIsOk() Then Exit Sub 'case of ZAAN Basic licence with the limit of 1000 documents stored in current ZAAN database reached

            Call LocksFswDataInputZaan()                                       'sets wait cursor and locks fswData, fswInput and fswZaan (import and copy) related events
            Call ResetFileMovePile()                                           'resets file move pile (used for "Undo move") and disables related local menu control
            SelectedDirPath = GetSelectedDirPathName()                         'get current data directory pathname corresponding to current mFileFilter

            Dim MyFiles() As String = e.Data.GetData(Forms.DataFormats.FileDrop)
            For i = 0 To MyFiles.Length - 1
                FileName = System.IO.Path.GetFileName(MyFiles(i))
                DestFilePathName = ""
                CubeImport = False
                If IsZaanCubeFileName(FileName) Then                           'case of a ZAAN cube file name (to be imported in database)
                    Response = MsgBox(mMessage(197), MsgBoxStyle.YesNo + MsgBoxStyle.Question)     'Do you confirm import of selected ZAAN cube(s) ?
                    If Response = MsgBoxResult.Yes Then
                        DestFilePathName = mZaanDbPath & "xin\" & FileName
                        CubeImport = True
                    End If
                Else
                    DestFilePathName = SelectedDirPath & "\" & FileName
                End If
                If DestFilePathName <> "" Then
                    Call DropSelectedFile(MyFiles(i), DestFilePathName, e.AllowedEffect)     'drops given source file to destination depending on copy/move drop effect, else on disk units
                    If CubeImport Then
                        tmrCubeImport.Enabled = True                           'starts timer to wait for 1 sec. before executing cube import
                    End If
                End If
            Next
            Call UnlocksFswDataInputZaan()                                     'unlocks fswZaan (import and copy), fswInput and fswData related events and resets default cursor
            Call DisplayListsAndHighlightItems(MyFiles, lvIn)                  'displays "in", "out" and "temp" lists and highlights moved files/items in destination list

        ElseIf e.Data.GetDataPresent("FileGroupDescriptor") Then               'case of an Outlook message to be dropped
            If Not CheckMoveInLimitIsOk() Then Exit Sub 'case of ZAAN Basic licence with the limit of 1000 documents stored in current ZAAN database reached

            SelectedDirPath = GetSelectedDirPathName()                         'get current data directory pathname corresponding to current mFileFilter
            Call DropOutlookFiles(e.Data, SelectedDirPath, lvIn)               'drops given Outlook files contained OutlookDataObject eData at destination directory
        End If
    End Sub

    Private Sub lvIn_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvIn.DragEnter
        'Debug.Print("lvIn_DragEnter...")              'TEST/DEBUG
        'If e.Data.GetDataPresent(DataFormats.FileDrop) Or e.Data.GetDataPresent(GetType(TreeNode)) Or e.Data.GetDataPresent("FileGroupDescriptor") Then
        If e.Data.GetDataPresent(Forms.DataFormats.FileDrop) Or e.Data.GetDataPresent("FileGroupDescriptor") Then
            e.Effect = Forms.DragDropEffects.All
        Else
            e.Effect = Forms.DragDropEffects.None
        End If
    End Sub

    Private Sub lvIn_ItemDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles lvIn.ItemDrag
        'Starts to drag selected documents for moving them
        Dim DirPathName, FileName As String
        Dim MyItem As ListViewItem
        Dim MyFiles(sender.SelectedItems.Count - 1) As String
        Dim i As Integer = 0

        'Debug.Print("lvIn_ItemDrag...")                         'TEST/DEBUG
        mStartDragLvInItem = True                               'flag set to avoid file drop from/to lvIn (may creates an unwanted copy at Selector position)

        For Each MyItem In sender.SelectedItems
            'Debug.Print("lvIn_ItemDrag :  MyItem.Text = " & MyItem.Text & "  MyItem.Tag = " & MyItem.Tag)      'TEST/DEBUG
            'DirPathName = mZaanDbPath & "data\" & MyItem.Tag    'gets dir name of selected file
            DirPathName = MyItem.Tag                            'gets absolute path of selected file
            FileName = MyItem.Text & MyItem.SubItems(8).Text
            'Call DeleteFileImage(DirPathName, FileName)         'deletes corresponding thumbnail image file (zzi*) that may have been generated by ZAAN
            MyFiles(i) = DirPathName & "\" & FileName           'adds the file to file list
            'Debug.Print("lvIn_ItemDrag :  " & i & " = " & MyFiles(i))    'TEST/DEBUG
            i = i + 1
        Next
        'sender.DoDragDrop(New DataObject(DataFormats.FileDrop, MyFiles), DragDropEffects.All)      'creates a DataObject containg the array of ListViewItems
        sender.DoDragDrop(New Forms.DataObject(Forms.DataFormats.FileDrop, MyFiles), Forms.DragDropEffects.Copy)      'creates a DataObject containg the array of ListViewItems
    End Sub

    Private Sub lvIn_ItemSelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles lvIn.ItemSelectionChanged
        If e.IsSelected Then                                         'item selected
            'Debug.Print("lvIn_ISC: selected : " & e.Item.Text)
            If lvIn.View = View.Details Then
                mSlideShowIndex = -1                                 'disables slide show with a unvalid index
            Else
                mSlideShowIndex = e.Item.Index                       'enables slide show by setting slide show index at new selection
            End If
            'Call CheckZoomDisplay(mZaanDbPath & "data\" & e.Item.Tag, e.Item.Text, e.Item.SubItems(8).Text)       'resets zoom displays and if zoom view activated, selects display control and shows it
            Call CheckZoomDisplay(e.Item.Tag, e.Item.Text, e.Item.SubItems(8).Text)      'resets zoom displays and if zoom view activated, selects display control and shows it
        End If
    End Sub

    Private Sub lvIn_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvIn.MouseDoubleClick
        'Opens selected file from "in" file directory
        Dim SelPathFile, FileName As String

        If Not lvIn.SelectedItems(0) Is Nothing Then
            'SelPathFile = mZaanDbPath & "data\" & lvIn.SelectedItems(0).Tag   'gets first selected "in" file
            SelPathFile = lvIn.SelectedItems(0).Tag                  'gets first selected "in" file
            FileName = lvIn.SelectedItems(0).Text & lvIn.SelectedItems(0).SubItems(8).Text
            OpenDocument(SelPathFile & "\" & FileName)               'open document in a new window with corresp. application
        End If
    End Sub

    Private Sub lvOut_AfterLabelEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.LabelEditEventArgs) Handles lvOut.AfterLabelEdit
        'Renames first selected "out" document
        Dim OldFileName, NewFileName, FileType, OldFilePathName As String
        Dim isCancelled As Boolean = False

        If e.Label Is Nothing Then Exit Sub
        If e.Label.Length > 0 Then                                   'node label is not empty
            NewFileName = e.Label
            OldFileName = lvOut.SelectedItems(0).Text
            If NewFileName = OldFileName Then Exit Sub

            If e.Label.IndexOfAny(New Char() {"<"c, ">"c, ":"c, """"c, "/"c, "\"c, "|"c, "?"c, "*"c}) = -1 Then   'entered characters are valid (for file name)
                FileType = lvOut.SelectedItems(0).SubItems(1).Text
                Try
                    OldFilePathName = mZaanImportPath & OldFileName & FileType
                    If UCase(NewFileName) = UCase(OldFileName) Then
                        My.Computer.FileSystem.RenameFile(OldFilePathName, "ZAAN_TEMP" & FileType)
                        My.Computer.FileSystem.RenameFile(mZaanImportPath & "ZAAN_TEMP" & FileType, NewFileName & FileType)
                    Else
                        My.Computer.FileSystem.RenameFile(OldFilePathName, NewFileName & FileType)
                    End If
                    lvOut.SelectedItems(0).Text = NewFileName
                Catch ex As Exception
                    e.CancelEdit = True                              'cancels item label edition
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation)      'displays error message
                End Try
            Else
                isCancelled = True
            End If
        Else
            isCancelled = True
        End If
        If isCancelled Then
            e.CancelEdit = True                                      'cancels item label edition
            MsgBox(mMessage(79), MsgBoxStyle.Exclamation)            'Entry not allowed !
            If lvOut.SelectedItems.Count > 0 Then
                lvOut.SelectedItems(0).BeginEdit()                   'goes back to label edition mode at first selected item
            End If
        End If
    End Sub

    Private Sub lvOut_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvOut.ColumnClick
        Dim i As Integer

        lvOut.Sorting = SortOrder.None                     'cancels automatic ordering mode
        If mlvOutOrder(e.Column) <= 0 Then
            mlvOutOrder(e.Column) = 1                      'set ascending order flag
            lvOut.ListViewItemSorter = New ListView2ItemComparer(e.Column)
        Else
            mlvOutOrder(e.Column) = -1                     'set descending order flag
            lvOut.ListViewItemSorter = New ListView2ItemRevComparer(e.Column)
        End If
        For i = 0 To mlvOutOrder.Length - 1
            If i <> e.Column Then
                mlvOutOrder(i) = 0                         'set other columns flags to no order
            End If
        Next
    End Sub

    Private Sub lvOut_ItemSelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles lvOut.ItemSelectionChanged
        If e.IsSelected Then                                         'item selected
            'Debug.Print("lvOut_ISC: selected : " & e.Item.Text)
            mSlideShowIndex = -1                                     'disables slide show with a unvalid index
            Call CheckZoomDisplay(mZaanImportPath, e.Item.Text, e.Item.SubItems(1).Text)      'resets zoom displays and if zoom view activated, selects display control and shows it
        End If
    End Sub

    Private Sub lvOut_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvOut.MouseDoubleClick
        'Opens selected file from "out" file directory
        Dim FileName As String

        FileName = lvOut.SelectedItems(0).Text & lvOut.SelectedItems(0).SubItems(1).Text      'get document name & type
        OpenDocument(mZaanImportPath & FileName)           'open document in a new window with corresp. application
    End Sub

    Private Sub trvW_BeforeLabelEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.NodeLabelEditEventArgs) Handles trvW.BeforeLabelEdit
        'Debug.Print("trvW_BeforeLabelEdit :  Tag = " & trvW.SelectedNode.Tag)   'TEST/DEBUG
        fswTree.EnableRaisingEvents = False                          'locks fswTree related events

        If IsTreeRootNode(trvW.SelectedNode) Then                    'case of a tree root node => cannot be modified !
            MsgBox(mMessage(61), MsgBoxStyle.Exclamation)            'It is not possible to modify a tree root name !
            e.Node.EndEdit(True)                                     'cancels tree node editing
        End If
    End Sub

    Private Sub trvW_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles trvW.Click
        'Debug.Print("trvW_Click: ")      'TEST/DEBUG
        mTreeClicked = True                                'flag set for enabling UpdateFileFilterAndDisplay() call in trvW_AfterSelect() event
    End Sub

    Private Sub trvW_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles trvW.DragDrop
        'Terminates drag-drop operation of files to be moved in selected directory
        Dim NodeKey, ParentKey, NodeTreeCode, FileName, DirPathName, DestDirPathName, DestFilePathName As String
        'Dim TargetDirPath As String
        Dim SceHierarchicalKey, DestHierarchicalKey As String
        Dim MyFiles() As String = e.Data.GetData(Forms.DataFormats.FileDrop)
        Dim i As Integer
        Dim draggedNode As TreeNode = CType(e.Data.GetData(GetType(TreeNode)), TreeNode)
        Dim Response As MsgBoxResult

        Dim targetPoint As System.Drawing.Point = trvW.PointToClient(New System.Drawing.Point(e.X, e.Y))     'retrieves the client coordinates of the mouse/drop location
        Dim targetNode As TreeNode = trvW.GetNodeAt(targetPoint)               'retrieve the node at the drop location

        If targetNode Is Nothing Then Exit Sub 'makes sure that a target node has been reached !

        trvW.HideSelection = True                                              'disable tree node selection highlight even if tree has not the focus

        If e.Data.GetDataPresent(Forms.DataFormats.FileDrop) Then                    'case of file(s) to be dropped
            mStartDragLvInItem = False                                         'makes sure that lvIn item drag flag is reset

            Me.Cursor = Cursors.WaitCursor                                     'sets wait cursor
            fswData.EnableRaisingEvents = False                                'locks fswData related events
            Call ResetFileMovePile()                                           'resets file move pile (used for "Undo move") and disables related local menu control
            For i = 0 To MyFiles.Length - 1
                'Debug.Print("=> item to drop : " & MyFiles(i))                 'TEST/DEBUG
                FileName = System.IO.Path.GetFileName(MyFiles(i))
                DirPathName = System.IO.Path.GetDirectoryName(MyFiles(i))
                DestDirPathName = ChangeNodeInFileDir(targetNode, DirPathName, FileName)  'returns updated file directory with destination node key inserted at right position
                DestFilePathName = DestDirPathName & "\" & FileName            'updates file directory with target node
                Call MoveSelectedFile(MyFiles(i), "", DestFilePathName, True)  'moves file
            Next
            fswData.EnableRaisingEvents = True                                 'unlocks fswData related events
            Me.Cursor = Cursors.Default                                        'resets default cursor

            'Call UpdateFileFilterAndDisplay()                                  'updates tree buttons and mFileFilter with last tree selection and display selected files if new filter
            trvW.SelectedImageKey = targetNode.ImageKey    'forces correct image display of selected node => WARNING : this will call again trvW_AfterSelect() event !!!
            Call UpdateAfterTreeNodeSelection()            'updates mFileFilter and related selector/selection and makes sure that lvIn is visible

            Call HighlightDestListItems(MyFiles, lvIn)                         'highlights moved files/items in destination list

        ElseIf e.Data.GetDataPresent(GetType(TreeNode)) Then                        'case of tree node to be dropped
            If draggedNode.TreeView Is trvW Then
                If Not draggedNode.Equals(targetNode) Then
                    NodeTreeCode = Mid(draggedNode.Tag, 2, 1)
                    If IsParentNode(draggedNode, targetNode) Or (NodeTreeCode <> Mid(targetNode.Tag, 2, 1)) Or (NodeTreeCode = "t") Or (NodeTreeCode = "u") Then
                        MsgBox(mMessage(51), MsgBoxStyle.Exclamation)                    'This folder move is not allowed !
                    Else
                        Response = MsgBox(mMessage(176), MsgBoxStyle.YesNo + MsgBoxStyle.Information)   'Do you confirm this folder move (database administration operation) ?
                        If Response = MsgBoxResult.Yes Then
                            'Debug.Print("drop node : " & draggedNode.Text & " as a child of : " & targetNode.Text)    'TEST/DEBUG
                            NodeKey = Mid(draggedNode.Tag, 2, mTreeKeyLength)            'OK in v3 database format for o/a/e/b/c nodes
                            ParentKey = Mid(targetNode.Tag, 2, mTreeKeyLength)           'OK in v3 database format for o/a/e/b/c nodes
                            FileName = "_" & NodeKey & "_" & ParentKey & Mid(draggedNode.Tag, 3 + 2 * mTreeKeyLength)     'changes parent key of dragged node to target node key
                            'Debug.Print("trvW_DragDrop => change filename : " & draggedNode.Tag & " into : " & FileName)    'TEST/DEBUG
                            SceHierarchicalKey = GetHierarchicalKey(NodeKey)             'returns related hierarchical key if v3 database format, same simple key else
                            fswTree.EnableRaisingEvents = False                          'locks fswTree related events
                            Try
                                My.Computer.FileSystem.RenameFile(mZaanDbPath & "tree\" & draggedNode.Tag, FileName)   'renames file name of dragged file
                                Call LoadTrees()                                         'loads all trees of current ZAAN database
                                DestHierarchicalKey = GetHierarchicalKey(NodeKey)        'returns new related hierarchical key if v3 database format, same simple key else
                                ChangeKeysInDataDirectories(SceHierarchicalKey, DestHierarchicalKey)     'changes source key by destination key in all data directories
                                Call DisplaySelector()                                   'displays selector buttons using mFileFilter selections
                                Call InitDisplaySelectedFiles()                          'redisplays all "in" files (file tags to be updated with new hierarchical paths)
                                ExpandTreeNode(NodeTreeCode, draggedNode.Tag)            'selects dropped node and expands tree view (that has been re-loaded) at this node
                            Catch ex As Exception
                                MsgBox(ex.Message, MsgBoxStyle.Exclamation)              'displays error message
                            End Try
                            fswTree.EnableRaisingEvents = True                           'unlocks fswTree related events
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub trvW_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles trvW.DragEnter
        'Debug.Print("trvW_DragEnter :  mTreeClicked = " & mTreeClicked)         'TEST/DEBUG
        If e.Data.GetDataPresent(Forms.DataFormats.FileDrop) Or (e.Data.GetDataPresent(GetType(TreeNode)) And (Not mTreeViewLocked)) Then
            e.Effect = Forms.DragDropEffects.All
        Else
            e.Effect = Forms.DragDropEffects.None
        End If
    End Sub

    Private Sub trvW_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles trvW.DragOver
        Dim targetPoint As System.Drawing.Point = trvW.PointToClient(New System.Drawing.Point(e.X, e.Y))     'retrieves the client coordinates of the mouse position
        Dim targetNode As TreeNode = trvW.GetNodeAt(targetPoint)               'retrieve the node at the drop location

        If targetNode Is Nothing Then Exit Sub
        'Debug.Print("trvW_DragOver :  over node = " & targetNode.Text)   'TEST/DEBUG

        trvW.HideSelection = False                                   'enable tree node selection highlight even if tree has not the focus
        trvW.SelectedNode = targetNode                               'selects the node at the mouse position => will trigger trvW_AfterSelect() event
    End Sub

    Private Sub trvW_ItemDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles trvW.ItemDrag
        'Starts to drag selected tree node for moving it, if possible
        DoDragDrop(e.Item, Forms.DragDropEffects.Move)
    End Sub

    Private Sub lvTemp_AfterLabelEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.LabelEditEventArgs) Handles lvTemp.AfterLabelEdit
        'Renames first selected "temp"/"copy" document
        Dim OldFileName, NewFileName, FileType, OldFilePathName As String
        Dim isCancelled As Boolean = False

        If e.Label Is Nothing Then Exit Sub
        If e.Label.Length > 0 Then                                   'node label is not empty
            NewFileName = e.Label
            OldFileName = lvTemp.SelectedItems(0).Text
            If NewFileName = OldFileName Then Exit Sub

            If e.Label.IndexOfAny(New Char() {"<"c, ">"c, ":"c, """"c, "/"c, "\"c, "|"c, "?"c, "*"c}) = -1 Then   'entered characters are valid (for file name)
                FileType = lvTemp.SelectedItems(0).SubItems(1).Text
                Try
                    OldFilePathName = mZaanCopyPath & OldFileName & FileType
                    If UCase(NewFileName) = UCase(OldFileName) Then
                        My.Computer.FileSystem.RenameFile(OldFilePathName, "ZAAN_TEMP" & FileType)
                        My.Computer.FileSystem.RenameFile(mZaanCopyPath & "ZAAN_TEMP" & FileType, NewFileName & FileType)
                    Else
                        My.Computer.FileSystem.RenameFile(OldFilePathName, NewFileName & FileType)
                    End If
                    lvTemp.SelectedItems(0).Text = NewFileName
                Catch ex As Exception
                    e.CancelEdit = True                              'cancels item label edition
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation)      'displays error message
                End Try
            Else
                isCancelled = True
            End If
        Else
            isCancelled = True
        End If
        If isCancelled Then
            e.CancelEdit = True                                      'cancels item label edition
            MsgBox(mMessage(79), MsgBoxStyle.Exclamation)            'Entry not allowed !
            If lvTemp.SelectedItems.Count > 0 Then
                lvTemp.SelectedItems(0).BeginEdit()                   'goes back to label edition mode at first selected item
            End If
        End If
    End Sub

    Private Sub lvTemp_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvTemp.ColumnClick
        Dim i As Integer

        lvTemp.Sorting = SortOrder.None                    'cancels automatic ordering mode
        If mlvTempOrder(e.Column) <= 0 Then
            mlvTempOrder(e.Column) = 1                     'set ascending order flag
            lvTemp.ListViewItemSorter = New ListView2ItemComparer(e.Column)
        Else
            mlvTempOrder(e.Column) = -1                    'set descending order flag
            lvTemp.ListViewItemSorter = New ListView2ItemRevComparer(e.Column)
        End If
        For i = 0 To mlvTempOrder.Length - 1
            If i <> e.Column Then
                mlvTempOrder(i) = 0                        'set other columns flags to no order
            End If
        Next
    End Sub

    Private Sub lvTemp_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvTemp.DragDrop
        'Terminates drag-drop operation of "in" files to be copied in "copy" directory
        Dim MyFiles() As String = e.Data.GetData(Forms.DataFormats.FileDrop)
        Dim i As Integer
        Dim FileName As String

        'Debug.Print("lvTemp_DragDrop : ")   'TEST/DEBUG
        trvW.HideSelection = True                                    'disable tree node selection highlight even if tree has not the focus
        If e.Data.GetDataPresent(Forms.DataFormats.FileDrop) Then
            Me.Cursor = Cursors.WaitCursor                           'sets wait cursor
            fswZaan.EnableRaisingEvents = False                      'locks fswZaan (import and copy) related events
            For i = 0 To MyFiles.Length - 1
                FileName = System.IO.Path.GetFileName(MyFiles(i))
                Call CopySelectedFile(MyFiles(i), "", mZaanCopyPath & FileName)   'copies file to "ZAAN\copy" directory
            Next
            fswZaan.EnableRaisingEvents = True                       'unlocks import and copy related events
            Me.Cursor = Cursors.Default                              'resets default cursor
            Call DisplayTempFiles()                                  'refreshes the display of all "temp"/"copy" files
            Call DisplayOutFiles()                                   'refresh the display of all "out" files (in case of !)

        ElseIf e.Data.GetDataPresent("FileGroupDescriptor") Then     'case of an Outlook message to be dropped
            Call DropOutlookFiles(e.Data, mZaanCopyPath, lvTemp)     'drops given Outlook files contained OutlookDataObject eData at destination directory
        End If
    End Sub

    Private Sub lvTemp_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvTemp.DragEnter
        If e.Data.GetDataPresent(Forms.DataFormats.FileDrop) Or e.Data.GetDataPresent("FileGroupDescriptor") Then
            e.Effect = Forms.DragDropEffects.All
        Else
            e.Effect = Forms.DragDropEffects.None
        End If
    End Sub

    Private Sub lvTemp_ItemDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles lvTemp.ItemDrag
        'Starts to drag selected temp files
        Dim FileName As String
        Dim MyItem As ListViewItem
        Dim MyFiles(sender.SelectedItems.Count - 1) As String
        Dim i As Integer = 0

        'Debug.Print("lvTemp_ItemDrag...")                       'TEST/DEBUG
        mStartDragLvInItem = False                              'makes sure that lvIn item drag flag is reset

        For Each MyItem In sender.SelectedItems
            'Debug.Print("=> : " & MyItem.Text)                  'TEST/DEBUG
            FileName = MyItem.Text & MyItem.SubItems(1).Text    'get document name & type
            MyFiles(i) = mZaanCopyPath & FileName               'adds the ListViewItem to the array of ListViewItems
            i = i + 1
        Next
        sender.DoDragDrop(New Forms.DataObject(Forms.DataFormats.FileDrop, MyFiles), Forms.DragDropEffects.All)      'creates a DataObject containg the array of ListViewItems
    End Sub

    Private Sub lvTemp_ItemSelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles lvTemp.ItemSelectionChanged
        If e.IsSelected Then                               'item selected
            'Debug.Print("lvOut_ISC: selected : " & e.Item.Text)
            mSlideShowIndex = -1                                     'disables slide show with a unvalid index
            Call CheckZoomDisplay(mZaanCopyPath, e.Item.Text, e.Item.SubItems(1).Text)      'resets zoom displays and if zoom view activated, selects display control and shows it
        End If
    End Sub

    Private Sub lvTemp_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvTemp.MouseDoubleClick
        'Opens selected file from "temp" file directory
        Dim FileName As String

        FileName = lvTemp.SelectedItems(0).Text & lvTemp.SelectedItems(0).SubItems(1).Text    'get document name & type
        OpenDocument(mZaanCopyPath & FileName)       'open document in a new window with corresp. application
    End Sub

    Private Sub lvOut_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvOut.DragEnter
        'Debug.Print("lvOut_DragEnter...")         'TEST/DEBUG
        If e.Data.GetDataPresent(Forms.DataFormats.FileDrop) Or e.Data.GetDataPresent("FileGroupDescriptor") Then
            e.Effect = Forms.DragDropEffects.All
        Else
            e.Effect = Forms.DragDropEffects.None
        End If
    End Sub

    Private Sub lvOut_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvOut.DragDrop
        'Terminates drag-drop operation of files to be moved in "out" directory
        Dim MyFiles() As String = e.Data.GetData(Forms.DataFormats.FileDrop)
        Dim i As Integer
        Dim FileName, DestFilePathName As String

        'Debug.Print("lvOut_DragDrop : ")   'TEST/DEBUG
        trvW.HideSelection = True                                              'disable tree node selection highlight even if tree has not the focus
        If e.Data.GetDataPresent(Forms.DataFormats.FileDrop) Then                    'case of file(s) to be dropped
            Call LocksFswDataInputZaan()                                       'sets wait cursor and locks fswData, fswInput and fswZaan (import and copy) related events
            Call ResetFileMovePile()                                           'resets file move pile (used for "Undo move") and disables related local menu control
            For i = 0 To MyFiles.Length - 1
                FileName = System.IO.Path.GetFileName(MyFiles(i))
                DestFilePathName = mZaanImportPath & FileName
                Call DropSelectedFile(MyFiles(i), DestFilePathName, e.AllowedEffect)     'drops given source file to destination depending on copy/move drop effect, else on disk units
            Next
            Call UnlocksFswDataInputZaan()                                     'unlocks fswZaan (import and copy), fswInput and fswData related events and resets default cursor
            Call DisplayListsAndHighlightItems(MyFiles, lvOut)                 'displays "in", "out" and "temp" lists and highlights moved files/items in destination list

        ElseIf e.Data.GetDataPresent("FileGroupDescriptor") Then               'case of an Outlook message to be dropped
            Call DropOutlookFiles(e.Data, mZaanImportPath, lvOut)              'drops given Outlook files contained OutlookDataObject eData at destination directory
        End If
    End Sub

    Private Sub lvOut_ItemDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles lvOut.ItemDrag
        'Starts to drag selected out files
        Dim FileName As String
        Dim MyItem As ListViewItem
        Dim MyFiles(sender.SelectedItems.Count - 1) As String
        Dim i As Integer = 0

        'Debug.Print("lvOut_ItemDrag...")                        'TEST/DEBUG
        mStartDragLvInItem = False                              'makes sure that lvIn item drag flag is reset

        For Each MyItem In sender.SelectedItems
            'Debug.Print("=> : " & MyItem.Text)                  'TEST/DEBUG
            FileName = MyItem.Text & MyItem.SubItems(1).Text    'get document name & type
            MyFiles(i) = mZaanImportPath & FileName             'adds the ListViewItem to the array of ListViewItems
            i = i + 1
        Next
        sender.DoDragDrop(New Forms.DataObject(Forms.DataFormats.FileDrop, MyFiles), Forms.DragDropEffects.Move)     'creates a DataObject containg the array of ListViewItems and flags move effect
    End Sub

    Private Sub tmrLicenseCheck_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrLicenseCheck.Tick
        tmrLicenseCheck.Enabled = False
        'frmAboutZaan.Show()                                'opens "About ZAAN" window
        frmAboutBox.Show()                                 'opens "About ZAAN" window

        'tmrLicenseCheck.Interval = 1000 * 3600 * 12        'sets next user licence check in 12 hours (usefull when user PC is not shut down by the end of the day)
        'tmrLicenseCheck.Enabled = True

        mLockSortLvIn = False                              'unlocks SortLvInColumn() in DisplaySelectedFiles()
        Call SortLvInColumn()                              'refreshs "in" files display with last column sorting set in mlvInOrder() or first one if no details view
    End Sub

    Private Sub tmrLvInDisplay_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrLvInDisplay.Tick
        tmrLvInDisplay.Enabled = False
        Call InitDisplaySelectedFiles()                    'initializes display of all selected files, starting at first page
    End Sub

    Private Sub lblDocName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblDocName.TextChanged
        'Debug.Print("lblDocName_TextChanged :  lblDocName.Text =>" & lblDocName.Text & "<")   'TEST/DEBUG
        'Try
        '  wbDoc.Navigate(lblDocName.Text, False)         'open document in same browser (NewWindow = False)
        '  If lblDocName.Text = "" Then
        '    wbDoc.Navigate("about:blank", False)       'open <blank page> in same browser (NewWindow = False)
        '  Else
        '    wbDoc.Url = New Uri(lblDocName.Text)       'open document in same browser (NewWindow = False)
        '  End If
        'Catch ex As Exception
        '  Debug.Print(ex.Message)                        'file not found !
        'End Try
    End Sub

    Private Sub tsmiLvInSelAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInSelAll.Click
        'Debug.Print("LvIn Select all...")     'TEST/DEBUG
        Dim itmX As ListViewItem
        For Each itmX In lvIn.Items
            itmX.Selected = True
        Next
    End Sub

    Private Sub tsmiLvInCut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInCut.Click
        Call CutCopyLvInSelection(True)                    'Cut selected files of "in" list into Clipboard and updates lists context menus
    End Sub

    Private Sub tsmiLvInCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInCopy.Click
        Call CutCopyLvInSelection(False)                   'Copy selected files of "in" list into Clipboard and updates lists context menus
    End Sub

    Private Sub tsmiLvInPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInPaste.Click
        Call PasteFilesToFilePath(GetSelectedDirPathName() & "\")    'move and copy Clipboard files at given file path, clears Clipboard and updates lists context menus
    End Sub

    Private Sub tsmiLvInRename_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInRename.Click
        If lvIn.SelectedItems.Count > 0 Then
            lvIn.SelectedItems(0).BeginEdit()              'starts label edition mode at first selected item
        End If
    End Sub

    Private Sub tsmiLvInDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInDelete.Click
        Call DeleteLvInSelection()                         'deletes selected files of "in" list
    End Sub

    Private Sub tsmiLvOutSelAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvOutSelAll.Click
        'Debug.Print("LvOut Select all...")     'TEST/DEBUG
        Dim itmX As ListViewItem
        For Each itmX In lvOut.Items
            itmX.Selected = True
        Next
    End Sub

    Private Sub tsmiLvOutCut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvOutCut.Click
        Call CutCopyListViewSelection(True, True)          'Cut selected files of "out" list view into Clipboard and updates lists context menus
    End Sub

    Private Sub tsmiLvOutCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvOutCopy.Click
        Call CutCopyListViewSelection(False, True)         'Copy selected files of "out" list view into Clipboard and updates lists context menus
    End Sub

    Private Sub tsmiLvOutPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvOutPaste.Click
        Call PasteFilesToFilePath(mZaanImportPath)         'move and copy Clipboard files at given file path, clears Clipboard and updates lists context menus
        Call DisplayOutFiles()                             'refresh the display of all "out" files to be stored in Zaan data "in" files
    End Sub

    Private Sub tsmiLvOutRename_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvOutRename.Click
        If lvOut.SelectedItems.Count > 0 Then
            lvOut.SelectedItems(0).BeginEdit()             'starts label edition mode at first selected item
        End If
    End Sub

    Private Sub tsmiLvOutDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvOutDelete.Click
        Call DeleteLvOutSelection()                        'deletes selected files of "out" list
    End Sub

    Private Sub tsmiLvTempSelAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvTempSelAll.Click
        'Debug.Print("LvTem Select all...")     'TEST/DEBUG
        Dim itmX As ListViewItem
        For Each itmX In lvTemp.Items
            itmX.Selected = True
        Next
    End Sub

    Private Sub tsmiLvTempCut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvTempCut.Click
        Call CutCopyListViewSelection(True, False)         'Cut selected files of "temp" list view into Clipboard and updates lists context menus
    End Sub

    Private Sub tsmiLvTempCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvTempCopy.Click
        Call CutCopyListViewSelection(False, False)        'Copy selected files of "temp" list view into Clipboard and updates lists context menus
    End Sub

    Private Sub tsmiLvTempPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvTempPaste.Click
        Call PasteFilesToFilePath(mZaanCopyPath)           'move and copy Clipboard files at given file path, clears Clipboard and updates lists context menus
        Call DisplayTempFiles()                            'refresh the display of all "temp"/"copy" files
    End Sub

    Private Sub tsmiLvTempRename_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvTempRename.Click
        If lvTemp.SelectedItems.Count > 0 Then
            lvTemp.SelectedItems(0).BeginEdit()             'starts label edition mode at first selected item
        End If
    End Sub

    Private Sub tsmiLvTempDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvTempDelete.Click
        Call DeleteLvTempSelection()                       'deletes selected files of "temp"/"copy" list
    End Sub

    Private Sub tsmiLvTempSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvTempSend.Click
        'Call SendMailWithAttachments()                     'call mail system for sending selected documents
        Call OpenMailCopyFiles()                           'open mail application to select copy file(s) to be sent
    End Sub

    Private Sub tsmiTrvWAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiTrvWAdd.Click
        Call AddTrvWChildNode()                            'adds a "new" child node to selected tree node
    End Sub

    Private Sub tsmiTrvWRename_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiTrvWRename.Click
        trvW.SelectedNode.BeginEdit()
    End Sub

    Private Sub tsmiTrvWDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiTrvWDelete.Click
        Call DeleteTrvWSelNode()                           'deletes selected tree node (if enabled)
    End Sub

    Private Sub cmsLvIn_Opened(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmsLvIn.Opened
        If lvIn.SelectedItems.Count > 0 Then                         'at least one item has been selected
            tsmiLvInCut.Enabled = True
            tsmiLvInCopy.Enabled = True
            tsmiLvInCopyPath.Enabled = True
            tsmiLvInOpenFolder.Enabled = True
            tsmiLvInAutoRename.Enabled = True
            tsmiLvInRename.Enabled = True
            tsmiLvInDelete.Enabled = True
            tsmiLvInCopyToZC.Enabled = True
            tsmiLvInMoveOut.Enabled = True
            tsmiLvInExport.Enabled = True
            tsmiLvInExpNameTable.Enabled = True
            tsmiLvInExpRefTable.Enabled = True
        Else                                                         'no item selected
            tsmiLvInCut.Enabled = False
            tsmiLvInCopy.Enabled = False
            tsmiLvInCopyPath.Enabled = False
            tsmiLvInOpenFolder.Enabled = False
            tsmiLvInAutoRename.Enabled = False
            tsmiLvInRename.Enabled = False
            tsmiLvInDelete.Enabled = False
            tsmiLvInCopyToZC.Enabled = False
            tsmiLvInMoveOut.Enabled = False
            tsmiLvInExport.Enabled = False
            tsmiLvInExpNameTable.Enabled = False
            tsmiLvInExpRefTable.Enabled = False
        End If
    End Sub

    Private Sub tsmiLvInCopyToZC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInCopyToZC.Click
        'Copies selected "in" files in "copy" directory
        Dim DirName, FileName, StoredFilePathName, TempFilePathName As String
        Dim i As Integer

        Me.Cursor = Cursors.WaitCursor                               'sets wait cursor
        fswZaan.EnableRaisingEvents = False                          'locks fswZaan (import and copy) related events
        For i = 0 To lvIn.SelectedItems.Count - 1
            DirName = lvIn.SelectedItems(i).Tag                      'get dir name of selected file
            FileName = lvIn.SelectedItems(i).Text & lvIn.SelectedItems(i).SubItems(8).Text
            'StoredFilePathName = mZaanDbPath & "data\" & DirName & "\" & FileName
            StoredFilePathName = DirName & "\" & FileName
            TempFilePathName = mZaanCopyPath & FileName
            Call CopySelectedFile(StoredFilePathName, "", TempFilePathName)    'copies file
        Next
        fswZaan.EnableRaisingEvents = True                           'unlocks import and copy related events
        Me.Cursor = Cursors.Default                                  'resets default cursor
        Call DisplayTempFiles()                                      'refreshes the display of all "temp"/"copy" files
        Call DisplayOutFiles()                                       'refresh the display of all "out" files (in case of !)
    End Sub

    Private Sub tsmiLvInMoveOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInMoveOut.Click
        'Moves selected "in" files in "out" directory
        Dim DirPathName, FileName, OutputFilePathName As String
        Dim MyFiles(lvIn.SelectedItems.Count - 1) As String
        Dim i As Integer

        Call LocksFswDataInputZaan()                                           'sets wait cursor and locks fswData, fswInput and fswZaan (import and copy) related events
        Call ResetFileMovePile()                                               'resets file move pile (used for "Undo move") and disables related local menu control
        For i = 0 To lvIn.SelectedItems.Count - 1
            'DirPathName = mZaanDbPath & "data\" & lvIn.SelectedItems(i).Tag    'gets dir name of selected file
            DirPathName = lvIn.SelectedItems(i).Tag                            'gets absolute path of selected file
            FileName = lvIn.SelectedItems(i).Text & lvIn.SelectedItems(i).SubItems(8).Text
            MyFiles(i) = FileName                                              'stores file name in MyFiles() table for later highlighting
            OutputFilePathName = mZaanImportPath & FileName
            Call MoveSelectedFile(FileName, DirPathName, OutputFilePathName, False)   'moves file
        Next
        Call UnlocksFswDataInputZaan()                                         'unlocks fswZaan (import and copy), fswInput and fswData related events and resets default cursor
        Call DisplayListsAndHighlightItems(MyFiles, lvOut)                     'displays "in", "out" and "temp" lists and highlights moved files/items in destination list
    End Sub

    Private Sub tsmiLvOutMoveIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvOutMoveIn.Click
        If CheckMoveInLimitIsOk() Then                               'checks in case of ZAAN Basic licence that the limit of 1000 documents stored in current ZAAN database has not yet been reached !
            Call MoveInAndFileOutSelection(False)                    'moves selected "out" files in current "in" list of Zaan data directory
        End If
    End Sub

    Private Sub cmsLvIn_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cmsLvIn.Opening
        Call UpdateCutCopyPasteMenusItems()                          'updates the enabling status of Cut/Copy/Paste items of lists context menus
    End Sub

    Private Sub cmsLvOut_Opened(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmsLvOut.Opened
        Dim FileName As String

        If tcFolders.SelectedIndex = 1 Then                          'case of import tab 1 selected
            tsmiLvOutChangeFolder.Enabled = True                     'enables "Change folder" control
            tsmiLvOutFoldersVisible.Enabled = True                   'enables "Folders tree visible"
        Else
            tsmiLvOutChangeFolder.Enabled = False                    'disables "Change folder" control
            tsmiLvOutFoldersVisible.Enabled = False                  'disables "Folders tree visible"
        End If

        tsmiLvOutCut.Enabled = False
        tsmiLvOutCopy.Enabled = False
        tsmiLvOutRename.Enabled = False
        tsmiLvOutDelete.Enabled = False
        tsmiLvOutMoveIn.Enabled = False
        tsmiLvOutAutoFile.Enabled = False
        tsmiLvOutImport.Enabled = False

        If lvOut.SelectedItems.Count > 0 Then                        'at least one item has been selected
            FileName = lvOut.SelectedItems(0).Text & lvOut.SelectedItems(0).SubItems(1).Text

            tsmiLvOutCut.Enabled = True
            tsmiLvOutCopy.Enabled = True
            tsmiLvOutRename.Enabled = True
            tsmiLvOutDelete.Enabled = True

            If IsZaanCubeFileName(FileName) Then                     'case of a ZAAN cube file name (to be imported in database)
                tsmiLvOutImport.Enabled = True
            Else
                tsmiLvOutMoveIn.Enabled = True
                tsmiLvOutAutoFile.Enabled = True
            End If
        End If
    End Sub

    Private Sub cmsLvOut_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cmsLvOut.Opening
        Call UpdateCutCopyPasteMenusItems()                'updates the enabling status of Cut/Copy/Paste items of lists context menus
    End Sub

    Private Sub cmsLvTemp_Opened(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmsLvTemp.Opened
        Dim FileType, FileName As String

        tsmiLvTempCut.Enabled = False
        tsmiLvTempCopy.Enabled = False
        tsmiLvTempResizePicture.Enabled = False
        tstLvTempResizePercentage.Enabled = False
        tsmiLvTempRename.Enabled = False
        tsmiLvTempDelete.Enabled = False
        tsmiLvTempSend.Enabled = False
        tsmiLvTempImport.Enabled = False

        If lvTemp.SelectedItems.Count > 0 Then                       'at least one item has been selected
            FileType = lvTemp.SelectedItems(0).SubItems(1).Text
            FileName = lvTemp.SelectedItems(0).Text & FileType

            tsmiLvTempCut.Enabled = True
            tsmiLvTempCopy.Enabled = True
            tsmiLvTempRename.Enabled = True
            tsmiLvTempDelete.Enabled = True
            tsmiLvTempSend.Enabled = True
            If (FileType = ".jpg") Or (FileType = ".gif") Or (FileType = ".png") Or (FileType = ".bmp") Then
                tsmiLvTempResizePicture.Enabled = True
                tstLvTempResizePercentage.Enabled = True
            End If
            If IsZaanCubeFileName(FileName) Then                     'case of a ZAAN cube file name (to be imported in database)
                tsmiLvTempImport.Enabled = True
            End If
        End If
    End Sub

    Private Sub cmsLvTemp_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cmsLvTemp.Opening
        Call UpdateCutCopyPasteMenusItems()                     'updates the enabling status of Cut/Copy/Paste items of lists context menus
    End Sub

    Private Sub trvW_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles trvW.LostFocus
        If mTreeClicked Then
            'Debug.Print("trvW_LostFocus :  mTreeClicked = " & mTreeClicked)    'TEST/DEBUG
            mTreeClicked = False                                'resets flag for avoiding calling twice UpdateFileFilterAndDisplay() done below
        End If
    End Sub

    Private Sub trvW_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles trvW.MouseDown
        Dim selNode As TreeNode = trvW.GetNodeAt(e.X, e.Y)      'retrieve the node at the mouse location
        Dim EditOngoing As Boolean = False

        'Debug.Print("trvW_MouseDown...")   'TEST/DEBUG
        If Not selNode Is Nothing Then
            If e.Button = Forms.MouseButtons.Right Then         'right button has been selected
                'Debug.Print("trvW_MouseDown/RIGHT click:  selNode = " & selNode.Text)   'TEST/DEBUG
                If Not trvW.SelectedNode Is Nothing Then
                    If trvW.SelectedNode.IsEditing Then
                        'Debug.Print("=> trvW_MouseDown/RIGHT click : current node edition at : " & trvW.SelectedNode.Text)   'TEST/DEBUG
                        EditOngoing = True
                    End If
                End If
                If Not EditOngoing Then
                    mTreeClicked = True                         'flag set for enabling UpdateFileFilterAndDisplay() call in trvW_AfterSelect() event
                    trvW.SelectedNode = selNode                 'selects the node at the mouse/drop position => will trigger trvW_AfterSelect() event
                    'trvW.SelectedImageKey = selNode.ImageKey    'forces correct image display of selected node => WARNING : this will call again trvW_AfterSelect() event !!!
                End If
            End If
        End If
        If SplitContainer2.Panel1Collapsed Then                 'left panel is hidden 
            trvW.ContextMenuStrip = Nothing                          '=> disables tree local menu
            trvW.LabelEdit = False                                   '=> disables tree node edition
            trvW.AllowDrop = False                                   '=> disables data dragging on tree
        Else                                                    'left panel is visible
            trvW.ContextMenuStrip = cmsTrvW                          '=> enables tree local menu
            If mTreeViewLocked Then                             'tree view edition is locked
                'trvW.ContextMenuStrip = Nothing                      '=> disables tree local menu
                trvW.LabelEdit = False                               '=> disables tree node edition
                trvW.AllowDrop = True                                '=> enables data dragging on tree
            Else                                                'tree view edition is unlocked
                'trvW.ContextMenuStrip = cmsTrvW                      '=> enables tree local menu
                trvW.LabelEdit = True                                '=> ensables tree node edition
                trvW.AllowDrop = True                                '=> enables data dragging on tree
            End If
        End If
    End Sub

    Private Sub cmsTrvW_Opened(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmsTrvW.Opened
        If mTreeViewLocked Then
            tsmiTrvWAdd.Enabled = False
            tsmiTrvWAddExternal.Enabled = False
            tsmiTrvWRename.Enabled = False
            tsmiTrvWDelete.Enabled = False
            tsmiTrvWImport.Enabled = False
            tsmiTrvWExport.Enabled = False
        Else
            Me.Cursor = Cursors.WaitCursor                          'sets wait cursor
            fswTree.EnableRaisingEvents = False                     'locks fswTree related events
            fswData.EnableRaisingEvents = False                     'locks fswData related events

            Call CheckNodeAddStatus()                               'checks child node (many) addition possibility
            Call CheckNodeDelStatus()                               'checks selected node deletion status and updates Del button enable status
            tsmiTrvWImport.Enabled = True
            tsmiTrvWExport.Enabled = True

            fswData.EnableRaisingEvents = True                      'unlocks fswData related events
            fswTree.EnableRaisingEvents = True                      'unlocks fswTree related events
            Me.Cursor = Cursors.Default                             'resets default cursor
        End If
    End Sub

    Private Sub cmsTrvW_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cmsTrvW.Opening
        'Me.Cursor = Cursors.WaitCursor                          'sets wait cursor
        'Call CheckNodeDelStatus()                               'checks selected node deletion status and updates Del button enable status
        'Call CheckNodeAddStatus()                               'checks child node (many) addition possibility
        'Me.Cursor = Cursors.Default                             'resets default cursor
    End Sub

    Private Sub tsmiSelTabsAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiSelTabsAdd.Click
        Call AddBookmark("")                               'adds current Selector position (set by mFileFilter) to bookmark list
    End Sub

    Private Sub tsmiSelTabsDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiSelTabsDelete.Click
        Call DeleteBookmark()                              'deletes currently selected bookmark from list
    End Sub

    Private Sub tsmiLvInAutoRename_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInAutoRename.Click
        'Debug.Print("tsmiLvInAutoRename_Click...")        'TEST/DEBUG
        Call AutoRenameLvInSelection()                     'automatically renames selected files of "in" list using current selector's folder names
    End Sub

    Private Sub tsmiLvTempResizePicture_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvTempResizePicture.Click
        Call ResizePicturesLvTempSelection()               'resizes selected .jpg files of "temp"/"copy" list after user confirmation
    End Sub

    Private Sub tsmiLvInExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInExport.Click
        Call ExportCubeFromLvInSelection()                 'exports a document cube from current LvIn selection with related ZAAN dimensions
    End Sub

    Private Sub fswTree_Changed(ByVal sender As System.Object, ByVal e As System.IO.FileSystemEventArgs) Handles fswTree.Changed
        'Debug.Print("fswTree_Changed : " & e.Name)         'TEST/DEBUG
        tmrTrvWDisplay.Enabled = True                      'locks during 1s. any eventual re-display of tree view
        Call CheckIfWatchedTreeNodeIsInSelector(e.Name)    'checks if given watched tree node file name matches with a current Selector position and sets mSelectorToUpdate if true
    End Sub

    Private Sub fswTree_Deleted(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs) Handles fswTree.Deleted
        'Debug.Print("fswTree_Deleted : " & e.Name)         'TEST/DEBUG
        tmrTrvWDisplay.Enabled = True                      'locks during 1s. any eventual re-display of tree view
        Call CheckIfWatchedTreeNodeIsInSelector(e.Name)    'checks if given watched tree node file name matches with a current Selector position and sets mSelectorToUpdate if true
    End Sub

    Private Sub fswTree_Renamed(ByVal sender As Object, ByVal e As System.IO.RenamedEventArgs) Handles fswTree.Renamed
        'Debug.Print("fswTree_Renamed : " & e.Name)         'TEST/DEBUG
        tmrTrvWDisplay.Enabled = True                      'locks during 1s. any eventual re-display of tree view
        Call CheckIfWatchedTreeNodeIsInSelector(e.Name)    'checks if given watched tree node file name matches with a current Selector position and sets mSelectorToUpdate if true
    End Sub

    Private Sub tmrTrvWDisplay_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrTrvWDisplay.Tick
        Dim CurNodeTag As String

        tmrTrvWDisplay.Enabled = False                     'disables timer
        If trvW.SelectedNode Is Nothing Then
            CurNodeTag = ""
        Else
            CurNodeTag = trvW.SelectedNode.Tag             'saves current node selection
        End If
        Call LoadTrees()                                   'loads all trees of current ZAAN database
        If CurNodeTag <> "" Then
            ExpandTreeNode("", CurNodeTag)                 'restores current node selection and expands tree view (that has been re-loaded) at this node
        End If

        If mSelectorToUpdate Then                          'flag set in sub CheckIfsWatchedTreeNodeIsInSelector()
            mSelectorToUpdate = False                      'resets flag
            Call DisplaySelector()                         'displays selector buttons using mFileFilter selections
            Call InitDisplaySelectedFiles()                'initializes display of all selected files, starting at first page
        End If
    End Sub

    Private Sub fswData_Changed(ByVal sender As System.Object, ByVal e As System.IO.FileSystemEventArgs) Handles fswData.Changed
        'Debug.Print("fswData_Changed :  path = " & e.FullPath)       'TEST/DEBUG
        If Not IsNewPiledInDir(e.FullPath, False) Then               'tests if given watched data directory is in current Selection
            tmrLvInDisplay.Enabled = True                            'locks during 1s. any eventual re-display of "in" files
            Call DeleteDirIfEmpty(fswData.Path & e.Name)             'deletes changed directory (if empty) of the pile
        End If
    End Sub

    Private Sub fswData_Deleted(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs) Handles fswData.Deleted
        'Debug.Print("fswData_Deleted :  path = " & e.FullPath)       'TEST/DEBUG
        If Not IsNewPiledInDir(e.FullPath, False) Then               'tests if given watched data directory is in current Selection
            tmrLvInDisplay.Enabled = True                            'locks during 1s. any eventual re-display of "in" files
            Call DeleteDirIfEmpty(fswData.Path & e.Name)             'deletes changed directory (if empty) of the pile
        End If
    End Sub

    Private Sub CheckIfWatchedTreeNodeIsInSelector(ByVal WatchedNodeFileName As String)
        'Checks if given watched tree node file name matches with a current Selector position and sets mSelectorToUpdate if true
        Dim NodeKey As String = Mid(WatchedNodeFileName, 1, 6)
        Dim NodeText As String = Mid(WatchedNodeFileName, 15)
        Dim p As Integer

        'Debug.Print("CheckIfWatchedTreeNodeIsInSelector :  NodeKey = " & NodeKey & "  NodeText = " & NodeText & "  mFileFilter = " & mFileFilter)    'TEST/DEBUG
        p = InStr(mFileFilter, NodeKey)
        If p > 0 Then                                                'case of existing NodeKey in mFileFilter => this tree node matches with a current Selector position
            mSelectorToUpdate = True                                 'sets flag to be used in Sub tmrTrvWDisplay_Tick() for updating Selector display
        End If
        'Debug.Print(" => CheckIfWatchedTreeNodeIsInSelector :  mSelectorToUpdate = " & mSelectorToUpdate)    'TEST/DEBUG
    End Sub

    Private Sub ImportDataFromWindows()
        'Imports a ZAAN database (data) from Windows in v4 format
        Debug.Print("ImportDataFromWindows...")   'TEST/DEBUG

    End Sub

    Private Sub ExportDataToWindows(ByVal ExportPathName As String, ByVal SourceDataPath As String)
        'Exports/copies files from given data path into a Windows mixted tree organisation depending on given tree codes series
        Dim TreeCodes() As String = Split(mTreeCodeSeries, " ")
        Dim dirInfo, dirSel As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim DirPathName, FileNameKey, FileName, Prefix As String
        Dim TreeCode, NodeKeys(), NodePaths(6), NodePath, NodeFullPath As String
        Dim i, j, DirNb, FilNb, n, k, p As Integer
        Dim NodeX() As TreeNode

        'Debug.Print("ExportDataToWindows : " & mTreeCodeSeries & "  SourceDataPath = " & SourceDataPath)   'TEST/DEBUG
        dirInfo = My.Computer.FileSystem.GetDirectoryInfo(SourceDataPath)
        DirNb = dirInfo.GetDirectories.Length
        n = 0
        For Each dirSel In dirInfo.GetDirectories("_*.*")
            n = n + 1
            'pgbZaan.Value = 10 + 90 * n / DirNb
            DirPathName = dirInfo.FullName & "\" & dirSel.Name & "\"

            FileNameKey = dirSel.Name                                          'get multidimensional source directory code
            Call ReplaceHierarchicalKeysBySimpleKeys(FileNameKey)              'replaces in given data directory name hierarchical keys by simple keys

            NodeKeys = Split(FileNameKey, "_")                                 'initializes NodeKeys() table with tree nodes keys
            'ReDim NodePaths(NodeKeys.Length - 1)                              'sets node paths table length to node keys table length
            For i = 1 To NodePaths.Length - 1
                NodePaths(i) = ""
            Next
            For i = 1 To NodeKeys.Length - 1
                TreeCode = Mid(NodeKeys(i), 1, 1)
                NodeX = trvW.Nodes.Find(NodeKeys(i), True)
                If NodeX.Length > 0 Then                                       'node key found in tree nodes
                    For j = 0 To TreeCodes.Length - 1                          'searches for requested node key order
                        If TreeCodes(j) = TreeCode Then
                            NodeFullPath = NodeX(0).FullPath
                            p = InStr(NodeFullPath, "\", vbTextCompare)
                            If p > 0 Then
                                NodeFullPath = Mid(NodeFullPath, p + 1)        'eliminates root name of current path
                            End If
                            NodeFullPath = "z" & j + 1 & "\" & NodeFullPath    'appends z# for marking ZAAN dimension #

                            If NodePaths(j + 1) = "" Then
                                NodePaths(j + 1) = NodeFullPath                'stores in NodePaths() full path name at requested position
                            Else
                                NodePaths(j + 1) = NodePaths(j + 1) & "\" & NodeFullPath      'stores in NodePaths() full path name at requested position
                            End If
                            Exit For
                        End If
                    Next
                End If
            Next
            NodePath = ""
            For i = 1 To NodePaths.Length - 1                                  'builds full node path
                If NodePaths(i) <> "" Then
                    NodePath = NodePath & NodePaths(i) & "\"
                End If
            Next
            'Debug.Print("--> ExportDbToWin : NodePath = " & NodePath)    'TEST/DEBUG

            If NodePath = "" Then
                pgbZaan.Value = 10 + 90 * n / DirNb                            'updates progress bar
            Else
                FilNb = dirSel.GetFiles.Length
                k = 0
                For Each fi In dirSel.GetFiles("*.*")
                    k = k + 1
                    pgbZaan.Value = 10 + 90 * ((n - 1) + (k / FilNb)) / DirNb  'updates progress bar
                    FileName = fi.Name
                    Prefix = Microsoft.VisualBasic.Left(FileName, 3)
                    If FileName <> "." And FileName <> ".." And Prefix <> "zzi" And UCase(FileName) <> "THUMBS.DB" Then     'ignores current and encompassing directories, ZAAN and Windows XP thumbnail images
                        'Debug.Print("--> ExportDbToWin : FileName = " & FileName)    'TEST/DEBUG
                        Call CopySelectedFile(DirPathName & FileName, "", ExportPathName & "\" & NodePath & FileName)       'copies file
                    End If
                Next
            End If
        Next
    End Sub

    Private Sub ExportInfoIniFiles(ByVal InfoDirPathName As String)
        'Exports info\zaan_*.ini files of current ZAAN database into v4 format (no more keys)
        Dim dirInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "info")
        Dim TreeCodes() As String = Split("u " & mTreeCodeSeries, " ")
        Dim TreeCode, FileContent, FileLines(), LineCells(), FileFilter, FilterKeys(), FilterTitles(6) As String
        Dim fi As System.IO.FileInfo
        Dim i, j, k, p, q As Integer
        Dim NodeX() As TreeNode

        'Debug.Print("ExportInfoIniFiles...")    'TEST/DEBUG

        For Each fi In dirInfo.GetFiles("zaan_*.ini")                     'scans all users *.ini files
            Try
                FileContent = My.Computer.FileSystem.ReadAllText(fi.FullName)
                FileLines = Split(FileContent, vbCrLf)
                FileContent = ""                                          'resets file content (to be rebuilt)
                For i = 0 To FileLines.Length - 1
                    LineCells = Split(FileLines(i), vbTab)
                    If (LineCells(0) = "Zdbf") Or (Mid(LineCells(0), 1, 3) = "Zbm") Then
                        FileFilter = LineCells(1)
                        p = InStr(FileFilter, " ")                        'eventually separates filter keys from saved bookmark text
                        If p = 0 Then
                        Else
                            FileFilter = Mid(FileFilter, 1, p - 1)
                        End If
                        If FileFilter <> "" Then
                            FilterKeys = Split(FileFilter, "*_")
                            FileFilter = ""                                              'resets file filter
                            For k = 0 To FilterTitles.Length - 1                         'resets table that will store ordered filter titles
                                FilterTitles(k) = ""
                            Next
                            If FilterKeys.Length > 1 Then                                'scans filter keys for converting them
                                For j = 1 To FilterKeys.Length - 1
                                    NodeX = trvW.Nodes.Find(FilterKeys(j), True)
                                    If NodeX.Length > 0 Then                             'node found with given key
                                        TreeCode = Mid(FilterKeys(j), 1, 1)
                                        For k = 0 To TreeCodes.Length - 1                'searches for tree code position in tree codes table (ex : "u t o a e b c")
                                            If TreeCode = TreeCodes(k) Then
                                                FilterTitles(k) = NodeX(0).FullPath
                                                q = InStr(FilterTitles(k), "\")
                                                If q > 0 Then                            'replaces path root directory by z# with current k index
                                                    FilterTitles(k) = "z" & k & Mid(FilterTitles(k), q)
                                                End If
                                                Exit For
                                            End If
                                        Next
                                    End If
                                Next
                            End If
                            FileFilter = FilterTitles(0)                                 'starts to build ordered file filter with access control
                            For k = 1 To FilterTitles.Length - 1                         'continues to build ordered file filter with 6 dimensions (may be empty)
                                If FilterTitles(k) <> "" Then
                                    FileFilter = FileFilter & "*" & FilterTitles(k)
                                End If
                            Next
                            FileLines(i) = LineCells(0) & vbTab & FileFilter             'replaces Zdbf or Zbm? line by its converted content
                        End If
                    End If
                    FileContent = FileContent & FileLines(i) & vbCrLf     'rebuilds file content
                Next
                My.Computer.FileSystem.WriteAllText(InfoDirPathName & "\" & fi.Name, FileContent, False)     'saves ZAAN\info\zaan_*.ini file
            Catch ex As Exception
                'MsgBox(ex.Message, MsgBoxStyle.Exclamation)               'file not found !
                Debug.Print(ex.Message)                                   'file not found !
            End Try
        Next
    End Sub

    Private Sub tsmiSelectorExportDBwin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiSelectorExportDBwin.Click
        frmExport.Show()                                             'opens "About ZAAN" window
    End Sub

    Private Sub tsmiSelectorAutoImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiSelectorAutoImport.Click
        If tsmiSelectorAutoImport.CheckState = CheckState.Checked Then
            tsmiSelectorAutoImport.CheckState = CheckState.Unchecked
            mXmode = ""                                              'cancels automatic import mode
        Else
            tsmiSelectorAutoImport.CheckState = CheckState.Checked
            mXmode = "auto"                                          'sets automatic import mode
        End If
    End Sub

    Private Sub tmrCubeImport_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrCubeImport.Tick
        tmrCubeImport.Enabled = False
        Call ImportCubeFromXin()                                     'imports any existing document cube from xin folder of current ZAAN database
    End Sub

    Private Sub fswXin_Created(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs) Handles fswXin.Created
        'Debug.Print("*** fswXin_Created : " & e.Name)      'TEST/DEBUG
        'tmrCubeImport.Enabled = True                       'RETIRED AUTOMATIC CUBE IMPORT (starts timer to wait for 1 sec. before executing cube import)
    End Sub

    Private Sub tsmiLvOutFoldersVisible_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvOutFoldersVisible.Click
        If tsmiLvOutFoldersVisible.CheckState = CheckState.Checked Then
            tsmiLvOutFoldersVisible.CheckState = CheckState.Unchecked
            SplitContainer4.Panel1Collapsed = True
            'Call UpdateImportPath(mZaanAppliPath & "import\")        'resets mZaanImportPath to ZAAN\import and updates "out" files display
        Else
            tsmiLvOutFoldersVisible.CheckState = CheckState.Checked
            SplitContainer4.Panel1Collapsed = False
        End If
    End Sub

    Private Sub trvInput_AfterExpand(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles trvInput.AfterExpand
        Dim DirPathName As String = e.Node.Tag

        'Debug.Print("trvInput_AfterExpand :  node = " & DirPathName & "  mInputTreeClicked = " & mInputTreeClicked & "  children = " & e.Node.Nodes.Count)     'TEST/DEBUG
        If mInputTreeClicked Then
            mInputTreeClicked = False
            Dim dirInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(DirPathName)
            LoadInputChildNodes(dirInfo, e.Node, 0)                  'loads child nodes if any
        End If
    End Sub

    Private Sub trvInput_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles trvInput.AfterSelect
        'Dim ShortCutPathName, TargetDirPath As String

        'Debug.Print("1- trvInput_AfterSelect : " & e.Node.Text & " => Tag = " & e.Node.Tag & " => Name = " & e.Node.Name)    'TEST/DEBUG

        'ShortCutPathName = e.Node.Tag & "\" & "_zaan_link.lnk"
        'If My.Computer.FileSystem.FileExists(ShortCutPathName) Then       'case of an existing shortcut link to a ZAAN data (multi-dimensional) directory
        '  TargetDirPath = GetWinShortcutTargetPath(ShortCutPathName)    'returns TargetPath of given Windows shortcut file
        '  Call OpenZaanDataDirectory(TargetDirPath)                     'open given ZAAN data directory, updates mFileFilter and displays Selector and related Selected files
        'End If

        tsmiTrvInputDelete.Enabled = False
        mImportPath = e.Node.Tag                                          'updates mImportPath
        tcFolders.TabPages(1).ToolTipText = mImportPath
        If Microsoft.VisualBasic.Right(mImportPath, 1) <> "\" Then
            mImportPath = mImportPath & "\"
        End If
        Call UpdateImportPath(mImportPath)                                'if given directory exists, updates mZaanImportPath and related tooltip and "out" files display
    End Sub

    Private Sub trvInput_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles trvInput.Click
        'Debug.Print("trvInput_Click...")    'TEST/DEBUG
        mInputTreeClicked = True
    End Sub

    Private Sub tsmiTrvInputChangeRoot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiTrvInputChangeRoot.Click
        Call SelectInputDir()                                        'selects input directory to be used for importing included files
    End Sub

    Private Sub tsmiTrvInputDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiTrvInputDelete.Click
        Dim dirInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(trvInput.SelectedNode.Tag)
        Try
            'Debug.Print("tsmiTrvInputDelete_Click : deletes dir = " & trvInput.SelectedNode.Text)     'TEST/DEBUG
            dirInfo.Delete()                                         'deletes directory if empty
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)              'error message !
            'Debug.Print(ex.Message)                                  'error message !
        End Try
        Call LoadInputTree()                                         'loads input tree with Windows directory structure starting at selected root
    End Sub

    Private Sub cmsTrvInput_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cmsTrvInput.Opening
        Dim dirInfo As System.IO.DirectoryInfo
        'Dim info As FileSystemInfo

        tsmiTrvInputDelete.Enabled = False
        If Not trvInput.SelectedNode Is Nothing Then
            'Debug.Print("input directory : " & trvInput.SelectedNode.Text)   'TEST/DEBUG
            Try
                dirInfo = My.Computer.FileSystem.GetDirectoryInfo(trvInput.SelectedNode.Tag)
                Dim infos As FileSystemInfo() = dirInfo.GetFileSystemInfos
                If infos.Count = 0 Then
                    'Debug.Print("=> input directory is empty : " & dirInfo.Name)    'TEST/DEBUG
                    tsmiTrvInputDelete.Enabled = True
                End If
            Catch ex As Exception
                MsgBox(mMessage(94), vbExclamation)                 'This folder does not exist anymore !
                Call LoadInputTree()                                '(re)loads input tree with Windows directory structure starting at selected root
            End Try
        End If
    End Sub

    Private Sub fswInput_Changed(ByVal sender As System.Object, ByVal e As System.IO.FileSystemEventArgs) Handles fswInput.Changed
        'Debug.Print("fswInput_Changed : " & e.Name)    'TEST/DEBUG
        'Call LoadInputTree()                               'loads input tree with Windows directory structure starting at selected root
        Call CheckFswInputChanges()                        'checks Input/mImportPath changes for refreshing related display
    End Sub

    Private Sub fswInput_Deleted(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs) Handles fswInput.Deleted
        'Debug.Print("fswInput_Deleted : " & e.Name)    'TEST/DEBUG
        'Call LoadInputTree()                               'loads input tree with Windows directory structure starting at selected root
        Call CheckFswInputChanges()                        'checks Input/mImportPath changes for refreshing related display
    End Sub

    Private Sub fswInput_Renamed(ByVal sender As Object, ByVal e As System.IO.RenamedEventArgs) Handles fswInput.Renamed
        'Debug.Print("fswInput_Renamed : " & e.Name)    'TEST/DEBUG
        'Call LoadInputTree()                               'loads input tree with Windows directory structure starting at selected root
        Call CheckFswInputChanges()                        'checks Input/mImportPath changes for refreshing related display
    End Sub

    Private Sub tsmiLvInUndoMove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInUndoMove.Click
        Call UndoFileMove()                                'undo file moves stored in file move pile and resets it and disables related local menu control
    End Sub

    Private Sub tsmiPctZoomSlideShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiPctZoomSlideShow.Click
        'Debug.Print("tsmiPctZoomSlideShow_Click :  checked = " & tsmiPctZoomSlideShow.Checked & "  mSlideShowIndex = " & mSlideShowIndex)    'TEST/DEBUG
        Call StartStopSlideShow()                                    'starts or stops slide show
    End Sub

    Private Sub tmrSlideShow_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrSlideShow.Tick
        'Enables next slide/picture of current selection to be displayed

        'Debug.Print("tmrSlideShow_Tick...")    'TEST/DEBUG
        If mIsSlideShowEffect Then                                   'progressive zooming of next image from center
            tmrSlideShow.Enabled = False                             'stops slide show transition timer that will be re-started at end of zoom effect timer
        End If
        Call GotoNextSlide()                                         'moves current slide cursor to next slide and displays related slide
    End Sub

    Private Sub cmsPctZoom_Opened(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmsPctZoom.Opened
        If (lvIn.Items.Count > 0) And (mSlideShowIndex >= 0) Then    'selection list non empty and slide show is enabled (source list is LvIn in image view mode)
            tsmiPctZoomSlideShow.Enabled = True
            tsmiPctZoomInterval.Enabled = True
            tstbPctZoom.Enabled = True
        Else
            tsmiPctZoomSlideShow.Enabled = False
            tsmiPctZoomInterval.Enabled = False
            tstbPctZoom.Enabled = False
        End If
    End Sub

    Private Sub tmrSlideShowEffect_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrSlideShowEffect.Tick
        'Displays next zoom effect step of next image, until full size reached

        If mSlideShowEffectIndex < mSlideShowEffectMax - 1 Then
            mSlideShowEffectIndex = mSlideShowEffectIndex + 1
            If mSlideShowImageRatio > 1 Then                                        'is a horizontal picture
                pctZoom2.Width = pctZoom.Width / mSlideShowEffectMax * mSlideShowEffectIndex
                pctZoom2.Height = pctZoom2.Width / mSlideShowImageRatio
            Else                                                                    'is a vertical picture
                pctZoom2.Height = pctZoom.Height / mSlideShowEffectMax * mSlideShowEffectIndex
                pctZoom2.Width = pctZoom2.Height * mSlideShowImageRatio
            End If

            pctZoom2.Top = pctZoom.Top + pctZoom.Height / 2 - pctZoom2.Height / 2
            pctZoom2.Left = pctZoom.Left + pctZoom.Width / 2 - pctZoom2.Width / 2
        Else
            pctZoom.Image = pctZoom2.Image
            pctZoom2.Visible = False
            tmrSlideShowEffect.Enabled = False                       'stops effect timer
            tmrSlideShow.Enabled = True                              'starts slide show transition timer to next slide/picture
        End If
    End Sub

    Private Sub pctZoom_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pctZoom.Click
        'Launches slide show if enabled (source list is LvIn in image view mode) or ask user to launch it if LvIn images available
        Dim Response As MsgBoxResult

        'Debug.Print("pctZoom_Click :  mSlideShowIndex = " & mSlideShowIndex)     'TEST/DEBUG

        If mSlideShowIndex >= 0 Then                                 'slide show is enabled (source list is LvIn in image view mode)
            tsmiPctZoomSlideShow.Checked = Not tsmiPctZoomSlideShow.Checked
            Call StartStopSlideShow()                                'starts or stops slide show
        Else
            Response = MsgBox(mMessage(139), MsgBoxStyle.YesNo + MsgBoxStyle.Information)    'Do you want to launch slide show of current selection, if any image available ?
            If Response = MsgBoxResult.Yes Then
                If mIsImageInLvIn Then                               'selection list contains one image at least
                    mIsImageViewBeforeSlideShow = mIsImageView
                    Call SetLvInDisplayMode(True)                    '=> sets images view mode
                    Call InitDisplaySelectedFiles()                  'displays all "in" files (empty filter)

                    mSlideShowIndex = 0
                    'pctZoom.Load(mZaanDbPath & "data\" & lvIn.Items(0).Tag & "\" & lvIn.Items(0).Text & lvIn.Items(0).SubItems(8).Text) 'loads first image
                    pctZoom.Load(lvIn.Items(0).Tag & "\" & lvIn.Items(0).Text & lvIn.Items(0).SubItems(8).Text) 'loads first image
                    tsmiPctZoomSlideShow.Checked = True
                    Call StartStopSlideShow()                        'starts slide show
                Else                                                 'no images listed in selection
                    MsgBox(mMessage(140), MsgBoxStyle.Information)   'No image available in current selection !
                    Call SetLvInDisplayMode(False)                   '=> sets details view mode
                    Call InitDisplaySelectedFiles()                  'displays all "in" files (empty filter)
                End If
            End If
        End If
    End Sub

    Private Sub btnSlidePrevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSlidePrevious.Click
        Call GotoPreviousSlide()                                     'moves current slide cursor to previous slide and displays related slide
    End Sub

    Private Sub btnSlideNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSlideNext.Click
        Call GotoNextSlide()                                         'moves current slide cursor to next slide and displays related slide
    End Sub

    Private Sub btnSlidePlayPause_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSlidePlayPause.Click
        mIsSlideShowPaused = Not mIsSlideShowPaused                  'switches play/pause status
        Call PlayPauseSlideShow()                                    'plays or stops activated slide show
    End Sub

    Private Sub btnSlideStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSlideStop.Click
        tsmiPctZoomSlideShow.Checked = False
        Call StartStopSlideShow()                                    'starts or stops slide show
    End Sub

    Private Sub btnSlideRotateLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSlideRotateLeft.Click
        Call RotateSlideShowPicture(RotateFlipType.Rotate270FlipNone)      'rotate current image to the left (90 degrees)
    End Sub

    Private Sub btnSlideRotateRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSlideRotateRight.Click
        Call RotateSlideShowPicture(RotateFlipType.Rotate90FlipNone)      'rotate current image to the right (90 degrees)
    End Sub

    Private Sub btnSlideDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSlideDelete.Click
        Call DeleteSlideShowPicture()                                     'deletes, if user confirms, current picture displayed in slide show
    End Sub

    Private Sub btnVideoBegin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVideoBegin.Click
        Dim timeIncr As Integer = 3000

        'If VLC2Zoom.input.Time - timeIncr > 0 Then
        '  VLC2Zoom.input.Time = VLC2Zoom.input.Time - timeIncr
        'Else
        '  VLC2Zoom.input.Time = 0
        'End If

    End Sub

    Private Sub btnVideoEnd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVideoEnd.Click
        Dim timeIncr As Integer = 3000

        'If VLC2Zoom.input.Time + timeIncr < VLC2Zoom.input.Length Then
        '  VLC2Zoom.input.Time = VLC2Zoom.input.Time + timeIncr
        'End If

    End Sub

    Private Sub btnVideoPlayPause_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVideoPlayPause.Click
        'If VLC2Zoom.playlist.isPlaying Then
        '  VLC2Zoom.playlist.togglePause()                          'toggle the pause state of VLC playlist
        '  btnVideoPlayPause.Image = My.Resources.play              'updates play/pause button image to play
        'Else
        '  VLC2Zoom.playlist.play()                                 'plays VLC playlist
        '  btnVideoPlayPause.Image = My.Resources.pause             'updates play/pause button image to pause
        'End If

    End Sub

    Private Sub btnVideoStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVideoStop.Click
        Call StopVideoPlayer()                                       'stops VLC video player (which activity should be indicated by tmrVideoProgress activity)
    End Sub

    Private Sub tmrVideoProgress_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrVideoProgress.Tick
        'Dim itemNb As Integer

        'If mVideoPlayListClearing Then
        '  If VLC2Zoom.playlist.items.count = 0 Then                'case of video playlist clearing operation is terminated
        '    mVideoPlayListClearing = False
        '    itemNb = VLC2Zoom.playlist.add(mVideoPathFileName)   'adds selected video file to VLC playlist
        '    VLC2Zoom.playlist.playItem(itemNb)                   'plays VLC playlist item
        '    btnVideoPlayPause.Image = My.Resources.pause         'updates play/pause button image to pause
        '  End If
        'Else
        '  pgbVideo.Value = VLC2Zoom.input.Position * 100
        'End If
    End Sub

    Private Sub fswZaan_Changed(ByVal sender As System.Object, ByVal e As System.IO.FileSystemEventArgs) Handles fswZaan.Changed
        'Debug.Print("fswZaan_Changed : " & e.Name)         'TEST/DEBUG
        Call CheckFswZaanChanges(e.Name)                             'checks ZAAN\copy and ZAAN\import changes for refreshing related displays
    End Sub

    Private Sub fswZaan_Deleted(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs) Handles fswZaan.Deleted
        'Debug.Print("fswZaan_Deleted : " & e.Name)         'TEST/DEBUG
        'Call CheckFswZaanChanges(e.Name)                             'checks ZAAN\copy and ZAAN\import changes for refreshing related displays
    End Sub

    Private Sub fswZaan_Renamed(ByVal sender As Object, ByVal e As System.IO.RenamedEventArgs) Handles fswZaan.Renamed
        'Debug.Print("fswZaan_Renamed : " & e.Name)         'TEST/DEBUG
        'Call CheckFswZaanChanges(e.Name)                             'checks ZAAN\copy and ZAAN\import changes for refreshing related displays
    End Sub

    Private Sub tsmiLvOutAutoFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvOutAutoFile.Click
        If CheckMoveInLimitIsOk() Then                               'checks in case of ZAAN Basic licence that the limit of 1000 documents stored in current ZAAN database has not yet been reached !
            Call MoveInAndFileOutSelection(True)                     'moves selected "out" files in current Zaan data directory using related Who/When dimensions founded
        End If
    End Sub

    Private Sub tsmiLvOutUndoMove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvOutUndoMove.Click
        Call UndoFileMove()                                'undo file moves stored in file move pile and resets it and disables related local menu control
    End Sub

    Private Sub tsmiLvTempUndoMove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvTempUndoMove.Click
        Call UndoFileMove()                                'undo file moves stored in file move pile and resets it and disables related local menu control
    End Sub

    Private Sub tsmiLvInRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInRefresh.Click
        Call InitDisplaySelectedFiles()                    'initializes display of all selected files, starting at first page
    End Sub

    Private Sub tsmiLvOutRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvOutRefresh.Click
        Call DisplayOutFiles()                             'refresh the display of all "out" files to be stored in Zaan data "in" files
    End Sub

    Private Sub tsmiLvTempRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvTempRefresh.Click
        Call DisplayTempFiles()                            'refresh display of "temp"/"copy" files list
    End Sub

    Private Sub lvIn_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvIn.MouseDown
        Dim info As ListViewHitTestInfo = lvIn.HitTest(e.X, e.Y)          'get item at mouse pointer
        Dim subItem As ListViewItem.ListViewSubItem = Nothing
        Dim ColIndex As Integer = -1
        Dim TreeCode, NodeKey, FileNameKey As String
        Dim p As Integer

        If info Is Nothing Then Exit Sub

        If (info.Item IsNot Nothing) Then
            subItem = info.Item.GetSubItemAt(e.X, e.Y)                    'get subitem at mouse pointer, if any
            If (subItem IsNot Nothing) Then
                ColIndex = info.Item.SubItems.IndexOf(subItem)            'get related column index pointed
                'Debug.Print("lvIn_MouseDown :  ColIndex = " & ColIndex)   'TEST/DEBUG
            End If
        End If

        If ColIndex = 0 Then                                              'document name selected
            lvIn.LabelEdit = True                                         'allows label edition (for document name modification)
        Else
            lvIn.LabelEdit = False                                        'cancels label edition
        End If

        If e.Button = Forms.MouseButtons.Left Then                        'left button has been used
            If (ColIndex > 0) And (ColIndex < 8) Then                     'one of the 7 dimensions ("Access control" included) has been selected
                NodeKey = ""
                If mTreeKeyLength = 0 Then                                     'no node keys in v4 database format (use directly Windows path names)
                    NodeKey = GetLvInCellReference(info.Item, ColIndex - 1)    'get path reference of given selection item at given column
                Else
                    TreeCode = Mid(lvIn.Columns(ColIndex).ImageKey, 2, 1)      'get related tree code
                    FileNameKey = info.Item.Tag
                    If TreeCode = "u" Then
                        p = InStrRev(FileNameKey, "\")
                        If p > 0 Then
                            NodeKey = "u" & Mid(FileNameKey, 1, p - 1)
                        End If
                    Else
                        Call ReplaceHierarchicalKeysBySimpleKeys(FileNameKey)       'replaces in given data directory name hierarchical keys by simple keys
                        NodeKey = GetFileTreeNodeKey(FileNameKey, TreeCode)         'get first node key of tree node of given tree code found in given filename keys
                    End If
                End If
                Call SelectTreeNode(NodeKey)                              'changes selection to given tree node with given node key
            End If
        Else                                                              'right (or middle) button has been used
            If (subItem IsNot Nothing) Then
                'Debug.Print("Subitem found : " & subItem.Text & "  of item : " & info.Item.Text & " => Who : " & info.Item.SubItems(2).Text)     'TEST/DEBUG
                Select Case ColIndex
                    Case 0                                                'document name selected
                        lvIn.ContextMenuStrip = cmsLvIn
                    Case 3, 4, 5, 6, 7                                    'document Who/What/Where/Status/Action for reference selected
                        mLvInColRightClick = ColIndex                     'stores selected column for multiple Who/What references menu initialization
                        lvIn.ContextMenuStrip = cmsLvInWhos
                    Case Else
                        lvIn.ContextMenuStrip = Nothing
                End Select
            Else
                'Debug.Print("No subitem found !")     'TEST/DEBUG
                If info.Location = ListViewHitTestLocations.Image Then    'image clicked (in image view mode)
                    mLvInColRightClick = 3   '2                           'stores selected column for multiple Who refernces menu initialization
                    lvIn.ContextMenuStrip = cmsLvInWhos
                Else
                    lvIn.ContextMenuStrip = cmsLvIn
                End If
            End If
        End If
    End Sub

    Private Sub tsmiLvInWhosMultiple_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInWhosMultiple.Click
        tsmiLvInWhosMultiple.Checked = Not tsmiLvInWhosMultiple.Checked
        If tsmiLvInWhosMultiple.Checked Then
            MsgBox(mMessage(147), MsgBoxStyle.Information)           'you may add now multiple Who/What references by dropping selected documents on appropriate tree folders
            lvIn.Columns(3).Text = mMessage(2) & " (*)"              'updates Who column header text
            lvIn.Columns(4).Text = mMessage(3) & " (*)"              'updates What column header text
            lvIn.Columns(5).Text = mMessage(4) & " (*)"              'updates Where column header text
            lvIn.Columns(6).Text = mMessage(5) & " (*)"              'updates Status column header text
            lvIn.Columns(7).Text = mMessage(6) & " (*)"              'updates Action for column header text
        Else
            lvIn.Columns(3).Text = mMessage(2)                       'resets Who column header text
            lvIn.Columns(4).Text = mMessage(3)                       'resets What column header text
            lvIn.Columns(5).Text = mMessage(4)                       'updates Where column header text
            lvIn.Columns(6).Text = mMessage(5)                       'updates Status column header text
            lvIn.Columns(7).Text = mMessage(6)                       'updates Action for column header text
        End If
    End Sub

    Private Sub tsmiLvInWhosDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInWhosDelete.Click
        'Debug.Print("tsmiLvInWhosDelete_Click : " & lvIn.SelectedItems(0).Text & "  Who to delete : " & tscbLvInWhosList.Text)   'TEST/DEBUG
        Dim TreeCode As String = ""

        Select Case mLvInColRightClick
            Case 3
                TreeCode = "o"                                       'Who column selected
            Case 4
                TreeCode = "a"                                       'What column selected
            Case 5
                TreeCode = "e"                                       'Where column selected
            Case 6
                TreeCode = "b"                                       'Status column selected
            Case 7
                TreeCode = "c"                                       'Action for column selected
        End Select
        If TreeCode <> "" Then
            Call DeleteLvInSelItemDimReference(tscbLvInWhosList.Text, TreeCode)     'deletes selected reference of LvIn selected item
        End If
    End Sub

    Private Sub cmsLvInWhos_Opened(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmsLvInWhos.Opened
        If tsmiLvInWhosMultiple.Checked Then
            tsmiLvInWhosDelete.Enabled = True                        'enables "Delete Who/What" control
            tscbLvInWhosList.Enabled = True                          '... and related deletion list
            Call LoadLvInMenuMultipleReferences()                    'loads in local LvIn menu Who references of selected item, if any
        Else
            tsmiLvInWhosDelete.Enabled = False                       'disables "Delete Who/What" control
            tscbLvInWhosList.Enabled = False                         '... and related deletion list
        End If
    End Sub

    Private Sub trvW_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles trvW.MouseUp
        Dim selNode As TreeNode = trvW.GetNodeAt(e.X, e.Y)           'retrieve the node at the mouse location
        Dim EditOngoing As Boolean = False

        If Not selNode Is Nothing Then
            If e.Button = Forms.MouseButtons.Left Then               'left button has been selected
                'Debug.Print("trvW_MouseUp/LEFT click:  selNode = " & selNode.Text)   'TEST/DEBUG
                If Not trvW.SelectedNode Is Nothing Then
                    'Debug.Print("=> trvW_MouseUp/LEFT click : selected node : " & trvW.SelectedNode.Text)   'TEST/DEBUG
                    If trvW.SelectedNode.Text = selNode.Text Then    'user has clicked on same tree selection
                        mTreeClicked = False                         'resets flag for avoiding calling twice UpdateFileFilterAndDisplay() done below

                        'If SplitContainer2.Panel1Collapsed Then      'tree view panel is collapsed => go back to LvIn display with no Selection change
                        '  lvIn.Visible = True                      'makes sure that "in" list is visible
                        '  trvW.Visible = False                     'hides tree
                        'End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub tsmiTrvWRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiTrvWRefresh.Click
        Call LoadTrees()                                   'reloads all trees of current ZAAN database
    End Sub

    Private Sub tbSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tbSearch.KeyDown
        'Debug.Print("tbSearch_KeyDown : e.KeyCode = " & e.KeyCode)    'TEST/DEBUG
        If e.KeyCode = 13 Then
            e.SuppressKeyPress = True                      'cancels <enter> key entry (for avoiding "forbidden input" sound !)
            Call SearchTreeNode(tbSearch.Text)             'searches tree node
        End If
    End Sub

    Private Sub tbSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbSearch.TextChanged
        mFoundTreeNodeKeys = ""
        mFoundListItemKeys = ""
    End Sub

    Private Sub lvSelector_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvSelector.ColumnClick
        'Debug.Print("lvSelector_ColumnClick :  col = " & e.Column)   'TEST/DEBUG
        Call ResetSelectorDim(e.Column)                              'resets given dimension of selector

        'Dim ColIndex As Integer = e.Column
        'Dim Title, TreeCode, NodeKey As String
        'With lvSelector.Columns(ColIndex)
        '  Title = .Text
        '  NodeKey = .Tag
        '  TreeCode = Mid(.ImageKey, 2, 1)
        'End With
        'Debug.Print("lvSelector_ColumnClick : => " & TreeCode & "|" & Title & "|" & NodeKey)   'TEST/DEBUG
        'ExpandTree(TreeCode, NodeKey)                               'changes tree node selection and updates selector
    End Sub

    Private Sub lvSelector_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvSelector.MouseDown
        Dim info As ListViewHitTestInfo = lvSelector.HitTest(e.X, e.Y)    'get item at mouse pointer
        Dim subItem As ListViewItem.ListViewSubItem = Nothing
        Dim ColIndex As Integer = -1
        Dim Title, TreeCode, NodeKey As String

        If info Is Nothing Then Exit Sub

        If (info.Item IsNot Nothing) Then
            subItem = info.Item.GetSubItemAt(e.X, e.Y)               'get subitem at mouse pointer, if any
            If (subItem IsNot Nothing) Then
                ColIndex = info.Item.SubItems.IndexOf(subItem)       'get related column index pointed
                'Debug.Print("lvSelector_MouseDown :  ColIndex = " & ColIndex)   'TEST/DEBUG
            End If
        End If
        If ColIndex = -1 Then Exit Sub

        If e.Button = Forms.MouseButtons.Left Then                   'left button has been used
            If ColIndex = 0 Then
                Title = info.Item.Text
                NodeKey = info.Item.Tag
            Else
                Title = info.Item.SubItems(ColIndex).Text
                NodeKey = info.Item.SubItems(ColIndex).Tag
            End If
            TreeCode = Mid(lvSelector.Columns(ColIndex).ImageKey, 2, 1)
            'Debug.Print("lvSelector_MouseDown : => " & TreeCode & "|" & Title & "|" & NodeKey)   'TEST/DEBUG

            ExpandTree(TreeCode, NodeKey)                            'changes tree node selection and updates selector
        Else                                                         'right (or middle) button has been used
            'available...
        End If
    End Sub

    Private Sub lvBookmark_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvBookmark.ColumnClick
        'Debug.Print("lvBookmark_ColumnClick : col = " & e.Column)    'TEST/DEBUG
        Call AddBookmark("")                                         'adds current Selector position (set by mFileFilter) to bookmark list
        Call UpdateBookmarkListHeader()                              'if bookmark list is visible, updates its header with current selector position and shows it if new
    End Sub

    Private Sub lvBookmark_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvBookmark.MouseDown
        Dim info As ListViewHitTestInfo = lvBookmark.HitTest(e.X, e.Y)    'get item at mouse pointer
        Dim Title, Key As String

        If e.Button = Forms.MouseButtons.Left Then                   'left button clicked
        Else                                                         'right button clicked
            mBookmarkListRightClick = True                           'flag used to maintain bookmark panel expanded when related menu pops up
        End If

        If info Is Nothing Then Exit Sub
        If info.Item Is Nothing Then Exit Sub
        Title = info.Item.Text
        Key = info.Item.Tag
        'Debug.Print("lvBookmark_MouseDown : " & Title & "|" & Key)   'TEST/DEBUG

        If Key <> "" Then
            mFileFilter = Key                                    'get stored file filter in selected bookmark
            Call DisplaySelector()                               'displays selector buttons using mFileFilter selections
            Call InitDisplaySelectedFiles()                      'initializes display of all selected files, starting at first page
        End If
    End Sub

    Private Sub lblDataAccess_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblDataAccess.MouseDown
        If e.Button = Forms.MouseButtons.Left Then                             'left button has been used
            ExpandMenuOrTree("u", lblDataAccess, My.Resources._u_3b, sender)   'expands selectors menu if tree view panel is hidden, tree view else
        End If
    End Sub

    Private Sub lblWho_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblWho.MouseDown
        If e.Button = Forms.MouseButtons.Left Then                        'left button has been used
            ExpandMenuOrTree("o", lblWho, My.Resources._o_3b, sender)     'expands selectors menu if tree view panel is hidden, tree view else
        End If
    End Sub

    Private Sub lblWhat_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblWhat.MouseDown
        If e.Button = Forms.MouseButtons.Left Then                        'left button has been used
            ExpandMenuOrTree("a", lblWhat, My.Resources._a_3b, sender)    'expands selectors menu if tree view panel is hidden, tree view else
        End If
    End Sub

    Private Sub lblWhen_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblWhen.MouseDown
        If e.Button = Forms.MouseButtons.Left Then                        'left button has been used
            ExpandMenuOrTree("t", lblWhen, My.Resources._t_3b, sender)    'expands selectors menu if tree view panel is hidden, tree view else
        End If
    End Sub

    Private Sub lblWhere_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblWhere.MouseDown
        If e.Button = Forms.MouseButtons.Left Then                        'left button has been used
            ExpandMenuOrTree("e", lblWhere, My.Resources._e_3b, sender)   'expands selectors menu if tree view panel is hidden, tree view else
        End If
    End Sub

    Private Sub lblWhat2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblWhat2.MouseDown
        If e.Button = Forms.MouseButtons.Left Then                        'left button has been used
            ExpandMenuOrTree("b", lblWhat2, My.Resources._b_3b, sender)   'expands selectors menu if tree view panel is hidden, tree view else
        End If
    End Sub

    Private Sub lblWho2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblWho2.MouseDown
        If e.Button = Forms.MouseButtons.Left Then                        'left button has been used
            ExpandMenuOrTree("c", lblWho2, My.Resources._c_3b, sender)    'expands selectors menu if tree view panel is hidden, tree view else
        End If
    End Sub

    Private Sub cmsBookmark_Opened(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmsBookmark.Opened
        If lvBookmark.SelectedItems.Count > 0 Then                   'one item has been selected
            tsmiSelTabsRefresh.Enabled = True
            tsmiSelTabsDelete.Enabled = True
        Else
            tsmiSelTabsRefresh.Enabled = False
            tsmiSelTabsDelete.Enabled = False
        End If
        If Microsoft.VisualBasic.Right(lvBookmark.SelectedItems(0).ImageKey, 1) = "d" Then
            tsmiSelTabsDefault.Checked = True
        Else
            tsmiSelTabsDefault.Checked = False
        End If
    End Sub

    Private Sub tsmiSelTabsRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiSelTabsRefresh.Click
        'Refreshs selected bookmark with text extracts of current selector buttons
        If mTreeKeyLength = 0 Then Exit Sub 'no need to refresh bookmark in v4 database format (no keys)

        If lvBookmark.SelectedItems.Count > 0 Then
            'Debug.Print("tsmiSelTabsRefresh_Click : " & lvBookmark.SelectedItems(0).Text)    'TEST/DEBUG
            lvBookmark.SelectedItems(0).Tag = mFileFilter
            lvBookmark.SelectedItems(0).Text = GetBookmark()         'returns a bookmark string of current Selector position including related dimension labels
        End If
    End Sub

    Private Sub tsmiTrvWImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiTrvWImport.Click
        Call ImportTreeTable()                                       'imports tree structure with named Who/What/Where roots from a tabulated text file
    End Sub

    Private Sub tsmiTrvWExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiTrvWExport.Click
        Call ExportTreeTable()                                       'exports tree structure with named Who/What/Where roots to a tabulated text file
    End Sub

    Private Sub cmsSelector_Opened(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmsSelector.Opened
        If mTreeKeyLength = 0 Then
            tsmiSelectorExportDBwin.Enabled = False                  'disables export database to Windows button in v4 database format (no keys)
        End If
        tsmiSelectorCheckDb.Enabled = True
    End Sub

    Private Sub tsmiSelectorImportDBwin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiSelectorImportDBwin.Click
        Call ImportDataFromWindows()                                 'imports a ZAAN database (data) from Windows in v4 format
    End Sub

    Private Sub tsmiSelectorTreeLock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiSelectorTreeLock.Click
        If tsmiSelectorTreeLock.CheckState = CheckState.Checked Then
            tsmiSelectorTreeLock.CheckState = CheckState.Unchecked
            mTreeViewLocked = False                                  'global flag (saved in ZAAN\info\zaan.ini) indicating if tree view edition is locked (false by default)
        Else
            tsmiSelectorTreeLock.CheckState = CheckState.Checked
            mTreeViewLocked = True
        End If
    End Sub

    Private Sub btnPrev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrev.Click
        'Debug.Print("btnPrev_Click :  lstPage.Items.Count = " & lstPage.Items.Count)    'TEST/DEBUG
        If lstPage.Items.Count = 0 Then Exit Sub
        If lstPage.SelectedIndex > 0 Then
            lstPage.SelectedIndex = lstPage.SelectedIndex - 1
            mFileFilter = lstPage.SelectedItems(0)
            Call DisplaySelector()                                   'displays selector buttons using mFileFilter selections
            Call InitDisplaySelectedFiles()                          'initializes display of all selected files, starting at first page
        End If
    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        'Debug.Print("btnNext_Click :  lstPage.Items.Count = " & lstPage.Items.Count)    'TEST/DEBUG
        If lstPage.Items.Count = 0 Then Exit Sub
        If lstPage.SelectedIndex < lstPage.Items.Count - 1 Then
            lstPage.SelectedIndex = lstPage.SelectedIndex + 1
            mFileFilter = lstPage.SelectedItems(0)
            Call DisplaySelector()                                   'displays selector buttons using mFileFilter selections
            Call InitDisplaySelectedFiles()                          'initializes display of all selected files, starting at first page
        End If
    End Sub

    Private Sub tsmiLvInCopyPath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInCopyPath.Click
        'Copies into Windows clipboard path of first selected file
        Dim DirPathName, FileName, FullPath As String

        'Debug.Print("tsmiLvInCopyPath_Click...")    'TEST/DEBUG
        If lvIn.SelectedItems.Count > 0 Then                                   'at least one item has been selected
            'DirPathName = mZaanDbPath & "data\" & lvIn.SelectedItems(0).Tag    'get dir name of first selected file
            DirPathName = lvIn.SelectedItems(0).Tag                            'get dir name of first selected file
            FileName = lvIn.SelectedItems(0).Text & lvIn.SelectedItems(0).SubItems(8).Text
            FullPath = DirPathName & "\" & FileName                            'build full path to be copied
            Forms.Clipboard.Clear()                                            'removes all data from Clipboard
            Forms.Clipboard.SetDataObject(FullPath)                            'writes the special ClipData object in Clipboard
        End If
    End Sub

    Private Sub tsmiLvInOpenFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInOpenFolder.Click
        'Opens Windows folder related to first selected document
        Dim DirPathName As String

        'Debug.Print("tsmiLvInOpenFolder_Click...")    'TEST/DEBUG
        If lvIn.SelectedItems.Count > 0 Then                                   'at least one item has been selected
            'DirPathName = mZaanDbPath & "data\" & lvIn.SelectedItems(0).Tag    'get dir name of first selected file
            DirPathName = lvIn.SelectedItems(0).Tag                            'get dir name of first selected file
            OpenDocument(DirPathName)                                          'open related system folder in a new window
        End If
    End Sub

    Private Sub tsmiSelectorCheckDb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiSelectorCheckDb.Click
        'Checks current ZAAN database indexes and repairs them if corrupted
        Call CheckDatabaseIndexes()                                  'checks tree files against backup file and checks data directories having uncomplete t keys or broken o/a/e/b/c keys for trying to restore them
        Call CheckMissingTreeYears()                                 'checks in data directory eventual missing years in tree directory and creates them
    End Sub

    Private Sub tsmiLvInExpNameTable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInExpNameTable.Click
        Call ExportNameTableFromLvInSelection()                      'exports name table of selected documents to be used for comments and/or printing
    End Sub

    Private Sub tsmiLvInExpRefTable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInExpRefTable.Click
        Call ExportRefTableFromLvInSelection()                       'exports reference table of selected documents to be used for publishing (with Word, Publisher, etc.)
    End Sub

    Private Sub tsmiLvTempImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvTempImport.Click
        If CheckMoveInLimitIsOk() Then                               'checks in case of ZAAN Basic licence that the limit of 1000 documents stored in current ZAAN database has not yet been reached !
            Call ImportCubesFromListSelection(lvTemp)                'imports selected cubes from given list view (lvOut or lvTemp) in current Zaan database (using its \xin directory)
        End If
    End Sub

    Private Sub tsmiLvOutImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvOutImport.Click
        If CheckMoveInLimitIsOk() Then                               'checks in case of ZAAN Basic licence that the limit of 1000 documents stored in current ZAAN database has not yet been reached !
            Call ImportCubesFromListSelection(lvOut)                 'imports selected cubes from given list view (lvOut or lvTemp) in current Zaan database (using its \xin directory)
        End If
    End Sub

    Private Sub tsmiSelectorChangeDbImage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiSelectorChangeDbImage.Click
        Call SelectZaanDbImage()                                     'selects an image file (.bmp, .gif, .jpg or .png) for associating it to current ZAAN database
    End Sub

    Private Sub tsmiSelectorChangeDb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiSelectorChangeDb.Click
        Call SelectZaanDb()                                          'selects Zaan database root directory
    End Sub

    Private Sub lblSelPagePrev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblSelPagePrev.Click
        If mInDocPageNb > 1 Then
            mInDocPageNb = mInDocPageNb - 1
        End If
        Call DisplaySelectedFiles()                                  'refreshes display of "in" files selected with current filter
    End Sub

    Private Sub lblSelPageNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblSelPageNext.Click
        If mInDocPageNb < mInDocPageMax Then
            mInDocPageNb = mInDocPageNb + 1
        End If
        Call DisplaySelectedFiles()                                  'refreshes display of "in" files selected with current filter
    End Sub

    Private Sub tsmiLvInDpp10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInDpp10.Click
        Call UpdatesInDocPerPageMenu(10)                             'updates number of lvIn documents per page in menu and related selection menu items
    End Sub

    Private Sub tsmiLvInDpp25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInDpp25.Click
        Call UpdatesInDocPerPageMenu(25)                             'updates number of lvIn documents per page in menu and related selection menu items
    End Sub

    Private Sub tsmiLvInDpp50_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInDpp50.Click
        Call UpdatesInDocPerPageMenu(50)                             'updates number of lvIn documents per page in menu and related selection menu items
    End Sub

    Private Sub tsmiLvInDpp100_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInDpp100.Click
        Call UpdatesInDocPerPageMenu(100)                            'updates number of lvIn documents per page in menu and related selection menu itemss
    End Sub

    Private Sub tsmiLvInDpp250_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInDpp250.Click
        Call UpdatesInDocPerPageMenu(250)                            'updates number of lvIn documents per page in menu and related selection menu items
    End Sub

    Private Sub tsmiLvInDpp500_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInDpp500.Click
        Call UpdatesInDocPerPageMenu(500)                            'updates number of lvIn documents per page in menu and related selection menu items
    End Sub

    Private Sub tsmiLvInDpp1000_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInDpp1000.Click
        Call UpdatesInDocPerPageMenu(1000)                           'updates number of lvIn documents per page in menu and related selection menu items
    End Sub

    Private Sub tcFolders_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tcFolders.SelectedIndexChanged
        'Debug.Print("tcFolders_SelectedIndexChanged :  index = " & tcFolders.SelectedIndex)    'TEST/DEBUG

        Select Case tcFolders.SelectedIndex
            Case 0                                                   'tab "1 - To file..." selected
                lvOut.Parent = tcFolders.TabPages(0)
                Call UpdateImportPath(mMyZaanImportPath)             'updates mZaanImportPath to ZAAN\import and displays related "out" files
            Case 1                                                   'tab "2 - To file..." selected
                lvOut.Parent = SplitContainer4.Panel2
                Call UpdateImportPath(mImportPath)                   'updates mZaanImportPath to user defined import path and displays related "out" files
        End Select
    End Sub

    Private Sub tsmiLvOutChangeFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvOutChangeFolder.Click
        'Debug.Print("tsmiLvOutChangeFolder...")   'TEST/DEBUG
        Call SelectInputDir()                                        'selects input directory to be used for importing included files
    End Sub

    Private Sub tsmiTrvInputRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiTrvInputRefresh.Click
        Call LoadInputTree()                                         'loads input tree with Windows directory structure starting at selected root
    End Sub

    Private Sub btnCubeTube_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCubeTube.Click
        'Toggles Cube (2D matrix view) / Tube (classic view) views of Section

        'Debug.Print("btnCubeTube_Click...")     'TEST/DEBUG
    End Sub

    Private Sub tcCubes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tcCubes.SelectedIndexChanged
        Dim TabIndex As String = tcCubes.SelectedIndex
        Dim TabPathName As String = tcCubes.SelectedTab.ToolTipText

        'Debug.Print("tcCubes_SelectedIndexChanged :  TabIndex = " & TabIndex & "  TabPathName = " & TabPathName)    'TEST/DEBUG

        pnlCube.Parent = tcCubes.SelectedTab               'moves related ZAAN cube display to selected Tab
        Call ChangeZaanDb(TabPathName)                     'changes current ZAAN database from list selection
    End Sub

    Private Sub tsmiSelectorTreeView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiSelectorTreeView.Click
        SplitContainer2.Panel1Collapsed = Not SplitContainer2.Panel1Collapsed
        Call SetLeftPanelButton()                          'sets left panel button display depending on splitter and background color
        lvIn.Focus()
    End Sub

    Private Sub tsmiSelectorViewer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiSelectorViewer.Click
        Call ToggleViewerPanel()                           'toggles viewer panel display
    End Sub

    Private Sub tsmiSelectorImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiSelectorImport.Click
        Call ToggleImportPanel()                           'toggles import panel display
    End Sub

    Private Sub lblSelectorReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblSelectorReset.Click
        Call ResetSelector()                               'resets selector to blank dimensions (current year will be then forced in LoadFileFilter() sub)
    End Sub

    Private Sub lblSearchFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblSearchFolder.Click
        Call SearchTreeNode(tbSearch.Text)                 'searches tree node
    End Sub

    Private Sub lblSearchDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblSearchDoc.Click
        Call SearchFileInSelection(tbSearch.Text)          'searches for a file within current selection
    End Sub

    Private Sub lblAboutZaan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblAboutZaan.Click
        frmAboutBox.Show()                                'opens "About ZAAN" window
    End Sub

    Private Sub tsmiLvInImageMode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInImageMode.Click
        Call ToggleSelectionView()                         'toogles Selection (lvIn) display between detail and image views
    End Sub

    Private Sub cmsSelChildren_Closed(sender As Object, e As ToolStripDropDownClosedEventArgs) Handles cmsSelChildren.Closed
        'Debug.Print("cmsSelChildren_Closed...")    'TEST/DEBUG
        tmrSelChildren.Enabled = True                      'sets timer to clear mCmsSelChildrenVisibleTreeCode menu flag
    End Sub

    Private Sub cmsSelChildren_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles cmsSelChildren.ItemClicked
        'Debug.Print("cmsSelChildren_ItemClicked :  Text | Tag = " & e.ClickedItem.Text & " | " & e.ClickedItem.Tag)   'TEST/DEBUG
        Dim TreeCode As String = Mid(e.ClickedItem.Tag, 2, 1)
        Dim ButtonTag As String

        If TreeCode = "u" Then
            ButtonTag = "*" & e.ClickedItem.Tag                           'keeps full Access Control path
        Else
            ButtonTag = "*_" & Mid(e.ClickedItem.Tag, 2, mTreeKeyLength)  'keeps only first characters related to node key
        End If

        ExpandTree(TreeCode, ButtonTag)
    End Sub

    Private Sub tmrSelChildren_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrSelChildren.Tick
        'Debug.Print("=> tmrSelChildren_Tick :  cmsSelChildren.Visible = " & cmsSelChildren.Visible)    'TEST/DEBUG
        tmrSelChildren.Enabled = False                               'stops timer used for clearing mCmsSelChildrenVisibleTreeCode flag
        mCmsSelChildrenVisibleTreeCode = ""
    End Sub

    Private Sub lblBookmark_MouseDown(sender As Object, e As MouseEventArgs) Handles lblBookmark.MouseDown
        If e.Button = Forms.MouseButtons.Left Then                             'left button has been used
            ExpandMenuOrTree("x", lblBookmark, My.Resources._x_3b, sender)     'expands selectors menu if tree view panel is hidden, tree view else
        End If
    End Sub

    Private Sub cmsSelBookmarks_Closed(sender As Object, e As ToolStripDropDownClosedEventArgs) Handles cmsSelBookmarks.Closed
        'Debug.Print("cmsSelBookmarks_Closed...")    'TEST/DEBUG
        tmrSelChildren.Enabled = True                      'sets timer to clear mCmsSelChildrenVisibleTreeCode menu flag
    End Sub

    Private Sub cmsSelBookmarks_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles cmsSelBookmarks.ItemClicked
        'Debug.Print("cmsSelBookmarks_ItemClicked :  Text | Tag = " & e.ClickedItem.Text & " | " & e.ClickedItem.Tag)   'TEST/DEBUG
        Dim Key As String = e.ClickedItem.Tag

        If Key <> "" Then
            mFileFilter = Key                                    'get stored file filter in selected bookmark
            Call DisplaySelector()                               'displays selector buttons using mFileFilter selections
            Call InitDisplaySelectedFiles()                      'initializes display of all selected files, starting at first page
        End If
    End Sub

    Private Sub btnDataAccess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDataAccess.Click
        ExpandTree("u", btnDataAccess.Tag)                 'changes tree node selection and updates selector
    End Sub

    Private Sub btnWhen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWhen.Click
        ExpandTree("t", btnWhen.Tag)                       'changes tree node selection and updates selector
    End Sub

    Private Sub btnWho_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWho.Click
        ExpandTree("o", btnWho.Tag)                        'changes tree node selection and updates selector
    End Sub

    Private Sub btnWhat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWhat.Click
        ExpandTree("a", btnWhat.Tag)                       'changes tree node selection and updates selector
    End Sub

    Private Sub btnWhere_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWhere.Click
        ExpandTree("e", btnWhere.Tag)                      'changes tree node selection and updates selector
    End Sub

    Private Sub btnWhat2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWhat2.Click
        ExpandTree("b", btnWhat2.Tag)                      'changes tree node selection and updates selector
    End Sub

    Private Sub btnWho2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWho2.Click
        ExpandTree("c", btnWho2.Tag)                       'changes tree node selection and updates selector
    End Sub

    Private Sub pnlSelectorNav_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlSelectorNav.Resize
        pnlSelectorHeaderTop.Width = pnlSelectorNav.Width - 5
    End Sub

    Private Sub tsmiSelectorNightMode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiSelectorNightMode.Click
        If mImageStyle < 10 Then                                     'case of a white background
            Call ChangeImageStyle(13)                                'sets mImageStyle to given style with dark blue background color and sets colors of all ZAAN controls
        Else
            Call ChangeImageStyle(3)                                 'sets mImageStyle to given style with lightgray background color and sets colors of all ZAAN controls
        End If
    End Sub

    Private Sub btnPanelTree_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPanelTree.Click
        If trvW.Visible Then
            SplitContainer2.Panel1Collapsed = Not SplitContainer2.Panel1Collapsed
        Else
            SplitContainer2.Panel1Collapsed = False
        End If
        trvW.Visible = True
        lvBookmark.Visible = False
        Call SetLeftPanelButton()                                    'sets left panel button display depending on splitter and background color
        lvIn.Focus()
    End Sub

    Private Sub btnDataAccessRoot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDataAccessRoot.Click
        Call ResetSelectorDim(0)                                     'resets given dimension of selector
    End Sub

    Private Sub btnWhenRoot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWhenRoot.Click
        Call ResetSelectorDim(1)                                     'resets given dimension of selector
    End Sub

    Private Sub btnWhoRoot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWhoRoot.Click
        Call ResetSelectorDim(2)                                     'resets given dimension of selector
    End Sub

    Private Sub btnWhatRoot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWhatRoot.Click
        Call ResetSelectorDim(3)                                     'resets given dimension of selector
    End Sub

    Private Sub btnWhereRoot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWhereRoot.Click
        Call ResetSelectorDim(4)                                     'resets given dimension of selector
    End Sub

    Private Sub btnWhat2Root_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWhat2Root.Click
        Call ResetSelectorDim(5)                                     'resets given dimension of selector
    End Sub

    Private Sub btnWho2Root_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWho2Root.Click
        Call ResetSelectorDim(6)                                     'resets given dimension of selector
    End Sub

    Private Sub tsmiLvInNewNote_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInNewNote.Click
        'Debug.Print("tsmiLvInNewNote_Click :  Tag = " & tsmiLvInNewNote.Tag)    'TEST/DEBUG
        Call CreateNewDocumentFile(tsmiLvInNewNote.Tag)              'creates given new document in current selection directory
    End Sub

    Private Sub tsmiLvInNewText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInNewText.Click
        'Debug.Print("tsmiLvInNewText_Click :  Tag = " & tsmiLvInNewText.Tag)    'TEST/DEBUG
        Call CreateNewDocumentFile(tsmiLvInNewText.Tag)              'creates given new document in current selection directory
    End Sub

    Private Sub tsmiLvInNewPres_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInNewPres.Click
        'Debug.Print("tsmiLvInNewPres_Click :  Tag = " & tsmiLvInNewPres.Tag)    'TEST/DEBUG
        Call CreateNewDocumentFile(tsmiLvInNewPres.Tag)              'creates given new document in current selection directory
    End Sub

    Private Sub tsmiLvInNewSpSh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiLvInNewSpSh.Click
        'Debug.Print("tsmiLvInNewSpSh_Click :  Tag = " & tsmiLvInNewSpSh.Tag)    'TEST/DEBUG
        Call CreateNewDocumentFile(tsmiLvInNewSpSh.Tag)              'creates given new document in current selection directory
    End Sub

    Private Sub tsmiTrvWAddExternal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiTrvWAddExternal.Click
        'Debug.Print("tsmiTrvWAddExternal_Click...")    'TEST/DEBUG
        Call AddTrvWChildNode(True)                                  'adds a "new" child node to selected tree node as an external folder (only for Access Control)
    End Sub

    Private Sub wbDoc_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles wbDoc.DocumentCompleted
        'Debug.Print("wbDoc_DocumentCompleted :  url = " & wbDoc.Url.ToString)    'TEST/DEBUG
    End Sub

    Private Sub pctZaanLogo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pctZaanLogo.Click
        'Call SelectZaanDb()                                          'selects Zaan database root directory
    End Sub

    Private Sub btnPanelView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPanelView.Click
        'Debug.Print("btnPanelView_Click...")    'TEST/DEBUG
        Call ToggleViewerPanel()                           'shows or hides viewer panel
    End Sub

    Private Sub btnPanelImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPanelImport.Click
        'Debug.Print("btnPanelImport_Click...")    'TEST/DEBUG
        Call ToggleImportPanel()                           'toggles import panel display
    End Sub

    Private Sub tmrDbExport_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrDbExport.Tick
        'Debug.Print("tmrDbExport_Tick :  mExportDBwinTime = " & mExportDBwinTime)    'TEST/DEBUG
        If mExportDBwinTime = "" Then                                     'no (automatic) db export time set
            tmrDbExport.Enabled = False                                   'stops db export timer
            Call ExportDatabaseToWindows()                                'exports current ZAAN database into a mono-dimensional Windows tree organization
        Else
            Dim ExportTime As Date = CDate(mExportDBwinTime)              'get stored db export start date/time
            If DateDiff(DateInterval.Second, ExportTime, Now) > 0 Then    'db export time has been passed over 1 second
                If mExportDBwinRepeatIndex = 0 Then                       'only once db export requested
                    tmrDbExport.Enabled = False                           'stops db export timer
                    mExportDBwinTime = ""                                 'clears (automatic) db export time
                    mZaanTitleOption = ""                                 'clears form title display option
                    Call InitZaanFormTitle()                              'updates ZAAN main form title
                    tsmiSelectorExportDBwin.Checked = False               'unchecks export control in selector menu
                Else
                    Select Case mExportDBwinRepeatIndex                   'repeated db export requested
                        Case 1                                            'Every day index
                            ExportTime = DateAdd(DateInterval.Day, 1, ExportTime)        'adds a day to (next) export time
                        Case 2                                            'Every week index
                            ExportTime = DateAdd(DateInterval.WeekOfYear, 1, ExportTime) 'adds a week to (next) export time
                        Case 3                                            'Every month index
                            ExportTime = DateAdd(DateInterval.Month, 1, ExportTime)      'adds a month to (next) export time
                    End Select
                    mExportDBwinTime = CStr(ExportTime)                   'stores next db export start time
                End If
                'Debug.Print("Automatic db export at : " & CDate(Now) & "  Next export time : " & mExportDBwinTime)   'TEST/DEBUG
                Call ExportDatabaseToWindows(True)                        'exports current ZAAN database into a mono-dimensional Windows tree organization
            End If
        End If
    End Sub

    Private Sub btnToday_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToday.Click
        'Debug.Print("Today's date selected !")    'TEST/DEBUG
        Dim CurDate As DateTime = Now
        Dim CreatedYear As String

        If Not YearTreeNodeExists(CurDate) Then                      'selected year does not exist
            fswTree.EnableRaisingEvents = False                      'locks fswTree related events (will be unlocked after node label editing)
            CreatedYear = CreateTreeFileYear(CurDate)                'creates in current ZAAN database (v3 format) current year tree file
            If CreatedYear <> "" Then
                Call LoadTrees()                                     'reloads all trees of current ZAAN database including this new When node
                MsgBox(CreatedYear & " " & mMessage(175), MsgBoxStyle.Information)    '<year> has been created !
            End If
            fswTree.EnableRaisingEvents = True                       'unlocks fswTree related events (no edition mode => cannot be unlocked after node label editing)
        End If

        Call ExpandTree("t", btnToday.Tag)                           'selects and expands tree at today's date node
    End Sub

    Private Sub tsmiSelTabsDefault_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiSelTabsDefault.Click
        Dim itmX As ListViewItem
        tsmiSelTabsDefault.Checked = Not tsmiSelTabsDefault.Checked            'toggles check mark
        For Each itmX In lvBookmark.Items
            itmX.ImageKey = "_x_" & mImageStyle                                'sets regular cube icon for each bookmark
        Next
        If tsmiSelTabsDefault.Checked Then
            lvBookmark.SelectedItems(0).ImageKey = "_x_" & mImageStyle & "d"   'sets defaut cube icon ("x" marked) on selected bookmark
        End If
    End Sub

    Private Sub tsmiSelectorCloseDb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmiSelectorCloseDb.Click
        Call DeleteZaanDbFromTabs()                        'deletes currently selected ZAAN database from tabs and selects first db from list or ZAAN-Demo1 if empty
    End Sub

    Private Sub btnPanelCube_Click(sender As Object, e As EventArgs) Handles btnPanelCube.Click
        If lvBookmark.Visible Then
            SplitContainer2.Panel1Collapsed = Not SplitContainer2.Panel1Collapsed
        Else
            SplitContainer2.Panel1Collapsed = False
        End If
        trvW.Visible = False
        lvBookmark.Visible = True
        Call UpdateBookmarkListHeader()                              'if bookmark list is visible, updates its header with current selector position and shows it if new
        Call SetLeftPanelButton()                                    'sets left panel button display depending on splitter and background color
    End Sub

End Class
