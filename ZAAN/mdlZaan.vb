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

Module mdlZaan
    Public mZaanAppliPath, mZaanDemoPath, mZaanDbRoot, mZaanDbPath, mZaanDbName, mMyZaanImportPath, mZaanImportPath, mZaanCopyPath, mLanguage, mLanguageIni, mInputPath, mImportPath As String
    Public mZaanExportDest, mXmode, mColSelectorWidths, mColDisplayWidths, mColDisplayIndexes, mColImportWidths, mColCopyWidths, mPanelWidths, mFormLocSize As String
    Public mFileFilter, mSelectorCode As String
    Public mTreeWhoRootIndex As Integer
    Public mMACaddress, mUserEmail, mLicenseKey, mUserLic2Line, mUserLicEndDate, mLicTryNb, mLicAccepted, mUserManualMainVer, mUserManualMinorVer, mUserManualFileName, mLicLastCheck As String
    Public mMessage(250) As String
    Public mZaanFormLoaded, mAboutBoxAutoClose, mAboutZaanAutoClose, mIsImageView, mIsImageInLvIn, mIsImageViewBeforeSlideShow, mIsCubeView, mLockSortLvIn As Boolean
    Public mIsSlideShowPaused, mIsSlideShowEffect, mIsSlideShowLeftPanelHidden, mIsSlideShowBottomPanelHidden, mVideoPlayListClearing As Boolean
    Public mLicTypeText, mLicTypeTextIni, mSlideShowPathFileName, mVideoPathFileName As String
    Public mLicTypeCode, mLicTypeCodeIni, mOpacity, mImageStyle, mSlideShowIndex, mSlideShowEffectIndex, mLvInColRightClick As Integer
    Public mSlideShowEffectMax As Integer = 10
    Public mSlideShowImageRatio As Single
    Public mTreeNodesCount As Long
    Public mBackColorHeader, mBackColorHeaderSel, mForeColorHeader, mForeColorHeaderSel, mBackColorContent, mBackColorPicture As Color
    Public mFileMovePile, mTreeCodeSeries, mTreeCodeImageStyle As String
    Public mTreeCodeSeriesIndex As Integer                 '1 for "t o a e b c", etc.
    Public mTreeKeyLength As Integer = 5                   'new key length for v3+ db format since 16/03/2011 and checked in Sub CheckTreeKeyLength()
    Public mTopRootKey As String = "zzzz"
    Public mTreeRootKey As String = "zzzy"
    Public mListsHeadersVisible, mTreeViewLocked As Boolean
    Public mTreeCodesV3() As String = Split("t o a e b c", " ")
    Public m4ZWhenYearMin As Integer = 1705                'WARNING W0 (see above) is year 0 in v3+ database format coding years with 2Z characters (36^2-1 = 1295 years from year 1705 to year 3000)
    Public m4ZWhenYearMax As Integer = 3000                'is year max in v3+ database format coding years with 2Z characters (36^2-1 = 1295 years from year 1705 to year 3000)
    Public mBackupTreeFileName As String = "zaan_backup_tree.txt"
    Public mAdminLogFileName As String = "zaan_admin_log.txt"
    Public mLogFileContent As String
    Public mInDocCount, mInDocPrevCount, mInDocPageNb, mInDocPageMax, mMatrixDocCount As Integer
    Public mInDocPerPage As Integer                        'number of documents per page to be displayed in selection list (lvIn)
    Public mInDocPerPageMax As Integer = 1000              'maximum number of documents per page to be displayed in selection list (lvIn)
    Public mMaxBasicDocToFile As Integer = 1000            'default 1000 = maximum number of documents that can be filed in current ZAAN database with ZAAN-Basic (free limited) license !
    Public mIsLocalDatabase As Boolean                     'flag set in UpdateZaanDbRootPathName() to indicate if current db path is local (root is "C:")
    Public mExportDBwinRepeatIndex As Integer = 0          'index indicating eventual repetition of ZAAN db export to Windows
    Public mExportDBwinTime As String = ""                 'date/time string indicating db export date/time
    Public mExportDBwinDest As String                      'destination path of ZAAN db export to Windows
    Public mZaanTitleOption As String = ""                 'optional string displayed in main form title, after current database name
    Public mCmsSelChildrenVisibleTreeCode As String = ""   'flag set in ExpandMenuOrTree() for toggling cmsSelChildren menu display

    Public Function GetZ4CharFromIndex(ByVal Index As Integer) As Char
        'Returns ZAAN character ("z" to "a" or "9" to "0") matching to given index (0 to 35)
        Dim AscVal As Integer

        If Index > 25 Then                                 'is a "9" to "0" index
            AscVal = 48 + (35 - Index)
        Else                                               'is a "z" to "a" index
            AscVal = 97 + (25 - Index)
        End If
        GetZ4CharFromIndex = Chr(AscVal)
    End Function

    Public Function GetZ4KeyFromIndex(ByVal NodeIndex As Integer, Optional ByVal KeyLength As Integer = 4) As String
        'Returns ZAAN 4 characters key calculated from given tree node index (0:zzzz, 1:zzzy, 2:zzzx, etc.)
        Dim zBase As Integer = 36                          '= 26 letters + 10 digits
        Dim MaxIndex As Integer = zBase ^ KeyLength - 1    '= 1 679 615 (= maximum number of children nodes of a parent node) if KeyLength = 4 (by default)
        Dim i, p, r, qTable(4) As Integer
        Dim Z4Key As String

        If NodeIndex < MaxIndex Then                       'given index is below max value
            r = NodeIndex
            For i = KeyLength - 1 To 1 Step -1
                p = zBase ^ i
                qTable(i) = r \ p                          'get the integer quotient
                r = r Mod p                                'get the remainder of the integer division
            Next
            qTable(0) = r
            'Debug.Print("  NodeIndex : " & NodeIndex & "  qTable() = " & qTable(3) & " " & qTable(2) & " " & qTable(1) & " " & qTable(0))   'TEST/DEBUG
            Z4Key = ""
            For i = KeyLength - 1 To 0 Step -1
                Z4Key = Z4Key & GetZ4CharFromIndex(qTable(i))   'appends ZAAN character ("z" to "a" or "9" to "0") matching to given index (0 to 35)
            Next
        Else                                               'given index reaches max value
            Z4Key = Mid("0000", 1, KeyLength)
        End If
        GetZ4KeyFromIndex = Z4Key
    End Function

    Public Function GetWhenV3KeyFromDateText(ByVal WhenTitle As String, Optional ByVal Is4ZWhenKey As Boolean = True) As String
        'Returns When key in current v3 (9 based)/v3+ (Z4) database format from given date text (ex: "YYYY", "YYYY-MM" or "YYYY-MM-DD").
        Dim ZTKey As String = ""
        Dim DateItems() As String = Split(WhenTitle, "-")
        Dim i, j, n As Integer
        Dim c As Char

        If Is4ZWhenKey Then                                          'is the new v3+ database format (with 4Z keys)
            If DateItems.Length > 0 Then                             'year text is given
                Dim YearValue As Integer = DateItems(0)
                If (YearValue < m4ZWhenYearMin) Or (YearValue > m4ZWhenYearMax) Then     'case of given year is out of range (1705-3000)
                    MsgBox(DateItems(0) & " " & mMessage(183) & " (" & m4ZWhenYearMin & "-" & m4ZWhenYearMax & ") !", MsgBoxStyle.Exclamation)    '<y> year is out of range !
                Else
                    Dim YearIndex As Integer = YearValue - m4ZWhenYearMin           'year index is 0 at year 1705 (thus max year index 1295 = 36^2-1 corresponds to year 3000)
                    ZTKey = GetZ4KeyFromIndex(YearIndex, 2)                         'get year Z4 key (limited to 2 digits) corresponding to year index
                    If DateItems.Length > 1 Then                                    'month text is given
                        ZTKey = ZTKey & GetZ4KeyFromIndex(DateItems(1), 1)          'appends month Z4 key (limited to 1 digit) corresponding to month index
                        If DateItems.Length > 2 Then                                'day text is given
                            ZTKey = ZTKey & GetZ4KeyFromIndex(DateItems(2), 1)      'appends day Z4 key (limited to 1 digit) corresponding to day index
                        Else
                            ZTKey = ZTKey & "z"
                        End If
                    Else
                        ZTKey = ZTKey & "zz"
                    End If
                End If
            End If
        Else                                                         'is the old v3 database format (with 9 based keys)
            For i = 0 To DateItems.Length - 1                        'scans YYYY, MM and DD strings
                For j = 1 To DateItems(i).Length                     'scans every character of YYYY, MM or DD string
                    c = Mid(DateItems(i), j, 1)
                    n = Asc(c)
                    If (n > 47) And (n < 58) Then                    'is a digit between 0 and 9
                        n = 57 - n                                   'get n complement to 9 = 9 - (n - 48)
                        ZTKey = ZTKey & n                            'appends converted digit to converted time key code
                    End If
                Next
            Next
        End If
        GetWhenV3KeyFromDateText = ZTKey
    End Function

    Public Function GetIndexZ4KeyChar(ByVal InputChar As Char) As Integer
        'Returns index value of given Z4 key character (0 to 25 from "z" to "a", and 26 to 35 from "9" to "0")
        Dim p, n As Integer

        n = -1
        p = Asc(InputChar)                                 'get Ascii (dec) value of InputChar
        If (p > 47) And (p < 58) Then                      'is a "0" to "9"
            n = 83 - p                                     '= 35 - (p - 48)
        Else
            If (p > 96) And (p < 123) Then                 'is a "a" to "z"
                n = 122 - p                                '= 25 - (p - 97)
            End If
        End If
        GetIndexZ4KeyChar = n
    End Function

    Public Function GetDateTextFromWhenV3Key(ByVal WhenKey As String, Optional ByVal Is4ZWhenKey As Boolean = True) As String
        'Returns date text (ex: "YYYY", "YYYY-MM" or "YYYY-MM-DD") from given When key (code after "t") in v3/v3+ database format
        Dim DateText As String = ""
        Dim c As String
        Dim i, n As Integer

        If Is4ZWhenKey Then                                          'case of v3+ database format (with 4Z When keys)
            Dim YearValue As Integer
            Dim YearIndex As Integer = 0
            Dim MonthIndex As Integer = GetIndexZ4KeyChar(Mid(WhenKey, 3, 1))
            Dim DayIndex As Integer = GetIndexZ4KeyChar(Mid(WhenKey, 4, 1))

            For i = 1 To 2                                           'scans first 2Z coding year
                n = GetIndexZ4KeyChar(Mid(WhenKey, i, 1))
                YearIndex = YearIndex + n * 36 ^ (2 - i)             'builds Key index in base 36 (varies from 0 to 1295)
            Next
            YearValue = m4ZWhenYearMin + YearIndex                   'get year value (varies from 3000 to 1705)
            DateText = Format(YearValue, "0000")
            If MonthIndex <> 0 Then
                DateText = DateText & "-" & Format(MonthIndex, "00")
                If DayIndex <> 0 Then
                    DateText = DateText & "-" & Format(DayIndex, "00")
                End If
            End If
        Else                                                         'case of old v3 database format (with 9 base When keys)
            For i = 1 To WhenKey.Length                              'scans all when key digits
                c = Mid(WhenKey, i, 1)
                If c = "x" Then
                    DateText = DateText & "0-" & DateText & "9"
                    Exit For
                Else
                    n = c
                    n = 9 - n
                    If i = 5 Or i = 7 Then                           'is month or day index start
                        DateText = DateText & "-" & n                '=> insert "-" separator and appends converted digit
                    Else
                        DateText = DateText & n                      'appends converted digit
                    End If
                End If
            Next
        End If
        GetDateTextFromWhenV3Key = DateText
    End Function

    Public Function SetUCaseAtFirst(ByVal GivenName As String) As String
        'Returns given name with upper-case set at first letter
        Dim UCaseAtFirst As String = GivenName
        If Len(GivenName) > 1 Then
            UCaseAtFirst = UCase(Mid(GivenName, 1, 1)) & Mid(GivenName, 2)
        End If
        SetUCaseAtFirst = UCaseAtFirst
    End Function

    Public Declare Auto Function SHGetFileInfo Lib "shell32.dll" (ByVal pszPath As String, ByVal dwFileAttributes As Integer, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Integer, ByVal uFlags As Integer) As IntPtr

    Public Const SHGFI_ICON As Integer = &H100
    Public Const SHGFI_SMALLICON As Integer = &H1
    Public Const SHGFI_LARGEICON As Integer = &H0
    Public Const SHGFI_OPENICON As Integer = &H2

    Structure SHFILEINFO
        Public hIcon As IntPtr
        Public iIcon As Integer
        Public dwAttributes As Integer
        <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=260)> _
        Public szDisplayName As String
        <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=80)> _
        Public szTypeName As String
    End Structure

    Function GetShellIconAsImage(ByVal argPath As String) As Image
        Dim mShellFileInfo As New SHFILEINFO
        mShellFileInfo.szDisplayName = New String(Chr(0), 260)
        mShellFileInfo.szTypeName = New String(Chr(0), 80)
        SHGetFileInfo(argPath, 0, mShellFileInfo, System.Runtime.InteropServices.Marshal.SizeOf(mShellFileInfo), SHGFI_ICON Or SHGFI_SMALLICON)
        ' attempt to create a System.Drawing.Icon from the icon handle that was returned in the SHFILEINFO structure
        Dim mIcon As System.Drawing.Icon
        Dim mImage As System.Drawing.Image
        Try
            mIcon = System.Drawing.Icon.FromHandle(mShellFileInfo.hIcon)  'creates a System.Drawing.Icon from the icon handle that was returned in the SHFILEINFO structure
            mImage = mIcon.ToBitmap
        Catch ex As Exception
            mImage = New System.Drawing.Bitmap(16, 16)                    'creates a blank System.Drawing.Image
        End Try
        GetShellIconAsImage = mImage
    End Function

    Public Function XOREncryption(ByVal input As String, ByVal key As String) As String
        'Returns XOR encryption of given input string with given key.
        Dim n, p, r, a1, a2 As Integer
        Dim output, c As String

        output = ""
        n = key.Length
        For i = 1 To input.Length
            p = 1 + (i - 1) Mod n                                    'get key character position (key may be shorter than input)
            a1 = Asc(Mid(input, i, 1))
            a2 = Asc(Mid(key, p, 1))
            r = a1 Xor a2
            c = Chr(r)                                               'get character
            output = output & c                                      'replaces character
        Next i
        XOREncryption = output
    End Function

    Public Function getMd5Hash(ByVal input As String) As String
        'Returns the hexadecimal md5 code of given input string
        Dim md5Hasher As MD5 = MD5.Create()                'creates a new instance of the MD5 object
        Dim data As Byte() = md5Hasher.ComputeHash(Encoding.Default.GetBytes(input))     'converts input string to a byte array and compute the hash
        Dim sBuilder As New StringBuilder()                'creates a new Stringbuilder to collect the bytes and create a string
        Dim i As Integer

        For i = 0 To data.Length - 1                       'loops through each byte of the hashed data and format each one as a hexadecimal string
            sBuilder.Append(data(i).ToString("x2"))
        Next i
        Return sBuilder.ToString()                         'returns the hexadecimal string
    End Function

    Public Function GetHMAC(ByVal mess As String) As String
        'Returns HMAC value of given message - Used to check message integrity (md5 hash code) and authenticity (keyZ)
        Dim keyZ As String = "f1183cb6dd9e816720ea519423f36123"
        Dim ipad As String = "36363636363636363636363636363636"
        Dim opad As String = "5c5c5c5c5c5c5c5c5c5c5c5c5c5c5c5c"

        ipad = XOREncryption(ipad, keyZ)
        opad = XOREncryption(opad, keyZ)
        GetHMAC = getMd5Hash(opad & getMd5Hash(ipad & mess))
    End Function

    Public Sub SetLicTypeCode()
        'Updates mLicTypeCode from mLicTypeText value ("Basic"/"First"/"Pro")
        Select Case mLicTypeText                           'sets license type code
            Case "Basic"
                mLicTypeCode = 10
            Case "First"
                mLicTypeCode = 20
            Case "Pro"
                mLicTypeCode = 30
            Case Else
                mLicTypeCode = 0                           'no valid license type code
        End Select
    End Sub

    Public Function CreateDirIfNotExistsOK(ByVal DirPathName As String) As Boolean
        'Creates given directory if not exists and returns true if exists or creation succeeded
        Dim OK As Boolean = True

        If Not My.Computer.FileSystem.DirectoryExists(DirPathName) Then   'case of directory doesn't exist
            Try
                My.Computer.FileSystem.CreateDirectory(DirPathName)       'creates directory
                'Debug.Print("CreateDirIfNotExistsOK:  directory created : " & DirPathName)    'TEST/DEBUG
            Catch ex As Exception
                OK = False
                MsgBox(ex.Message, MsgBoxStyle.Exclamation)               'cannot create directory !
            End Try
        End If
        Return OK
    End Function

    Public Function CreateDirIfNotExistsStatus(ByVal DirPathName As String) As Integer
        'If given directory exists return 2, else try to create it and returns 1 creation succeeded or 0 if creation failed
        Dim Status As Integer = 2

        If Not My.Computer.FileSystem.DirectoryExists(DirPathName) Then   'case of directory doesn't exist
            Try
                My.Computer.FileSystem.CreateDirectory(DirPathName)       'creates directory
                'Debug.Print("CreateDirINES:  directory created : " & DirPathName)    'TEST/DEBUG
                Status = 1
            Catch ex As Exception
                Status = 0
                MsgBox(ex.Message, MsgBoxStyle.Exclamation)               'cannot create directory !
            End Try
        End If
        'Debug.Print("CreateDirINES :  DirPathName = " & DirPathName & "  Status = " & Status)   'TEST/DEBUG
        Return Status
    End Function

    Public Sub CleanDataOldFileImages(ByVal FilePath As String)
        'Deletes any remaining old thumbnail file images no more associated to source image files in a given directory
        Dim dirSel As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim FileName, FullFileName As String
        Dim p As Integer

        'Debug.Print("CleanDataOldFileImages: " & FilePath)          'TEST/DEBUG
        dirSel = My.Computer.FileSystem.GetDirectoryInfo(FilePath)

        For Each fi In dirSel.GetFiles("zzi*")                       'searches for corresponding thumbnail image file...
            p = InStr(fi.Name, "_")
            If p > 0 Then
                FileName = Mid(fi.Name, p + 1)
                FullFileName = FilePath & "\" & FileName
                If Not My.Computer.FileSystem.FileExists(FullFileName) Then   'case of no related source image file exists
                    Try
                        fi.Delete()                                  'deletes image file defintively
                        'Debug.Print("=> old Thumbnail image deleted : " & FilePath & "\" & fi.Name)    'TEST/DEBUG
                    Catch ex As Exception
                        Debug.Print("CleanDataOldFileImages error : " & ex.Message)    'TEST/DEBUG
                    End Try
                End If
            End If
        Next
    End Sub

    Public Sub DeleteFileIfExists(ByVal FilePathName As String)
        'Deletes given file if exists

        If My.Computer.FileSystem.FileExists(FilePathName) Then      'checks if file exists
            My.Computer.FileSystem.DeleteFile(FilePathName)          'deletes it
        End If
    End Sub

    Public Sub DeleteDirIfEmpty(ByVal DirPathName As String)
        'Deletes given directory if empty.
        Dim dirInfo As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim DirIsEmpty As Boolean
        Dim FileName As String

        If (DirPathName = "") Or (DirPathName = "_") Then Exit Sub
        If Not My.Computer.FileSystem.DirectoryExists(DirPathName) Then Exit Sub

        'Debug.Print("=> DeleteDirIfEmpty : DirPathName = " & DirPathName)   'TEST/DEBUG
        Try
            DirIsEmpty = True
            dirInfo = My.Computer.FileSystem.GetDirectoryInfo(DirPathName)
            For Each fi In dirInfo.GetFiles("*.*")
                FileName = fi.Name
                If FileName <> "." And FileName <> ".." Then
                    If UCase(FileName) = "THUMBS.DB" Then
                        'Debug.Print("=> DELETES thumbs.db : " & DirPathName & "\" & FileName)   'TEST/DEBUG
                        My.Computer.FileSystem.DeleteFile(DirPathName & "\" & FileName)
                    Else
                        DirIsEmpty = False                           'source directory is not empty
                    End If
                    Exit For
                End If
            Next
            If DirIsEmpty Then
                'Debug.Print(" => TO BE DELETED : DirPathName = " & DirPathName)   'TEST/DEBUG
                dirInfo.Delete()                                     'deletes directory if empty
            End If
        Catch ex As Exception
            Debug.Print("DeleteDirIfEmpty error : " & ex.Message)    'TEST/DEBUG
        End Try
    End Sub

    Public Sub BackupTree()
        'Saves tree source files into one text file in info folder of current database and deletes previous backups
        Dim dirInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(mZaanDbPath & "tree\")
        Dim fi As System.IO.FileInfo
        Dim FileName, FileContent, FileFilter, Key, ParentKey, Title As String
        Dim OldBackupPathFileName, BackupPathFileName As String

        FileContent = ""
        FileFilter = "_*_*__*.txt"
        For Each fi In dirInfo.GetFiles(FileFilter)                            'scans all tree files
            FileName = System.IO.Path.GetFileNameWithoutExtension(fi.Name)     'get file name without extension
            Key = Mid(FileName, 2, mTreeKeyLength)
            ParentKey = Mid(FileName, 3 + mTreeKeyLength, mTreeKeyLength)
            Title = Mid(FileName, 5 + 2 * mTreeKeyLength)
            FileContent = FileContent & Key & vbTab & ParentKey & vbTab & Title & vbCrLf
        Next

        OldBackupPathFileName = mZaanDbPath & "data\zz_backup_tree_zzzz.txt"   'sets old tree view backup file path name
        BackupPathFileName = mZaanDbPath & "info\" & mBackupTreeFileName       'sets tree view backup file path name (since v3 database format)
        Try
            My.Computer.FileSystem.WriteAllText(BackupPathFileName, FileContent, False)       'saves backup file
            If My.Computer.FileSystem.FileExists(OldBackupPathFileName) Then   'old backup file exists
                My.Computer.FileSystem.DeleteFile(OldBackupPathFileName)       '=> deletes it !
            End If
        Catch ex As System.IO.DirectoryNotFoundException
            'Debug.Print(ex.Message)                                            'TEST/DEBUG
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)                        'displays error message
        End Try
    End Sub

    Public Function UriEncode(ByVal s As String) As String
        'Returns an Uri compatible string from a given string (respecting the open uri convention)
        Dim result As String = ""
        Dim c As String

        For Each c In s
            Select Case c
                Case " "
                    result &= "_"
                Case "à", "â", "ä"
                    result &= "a"
                Case "é", "è", "ê", "ë"
                    result &= "e"
                Case "î", "ï"
                    result &= "i"
                Case "ô", "ö"
                    result &= "o"
                Case "ù", "ü"
                    result &= "u"
                Case "ç"
                    result &= "c"
                Case Else
                    result &= c
            End Select
        Next
        UriEncode = result
    End Function

    Public Sub AddToZipFileNameUri(ByVal zip As Package, ByVal FilePathName As String, ByVal LocalFileUri As String)
        'Adds given file to given zip package with given filename uri (starting with a "/" and respecting uri convention)
        If LocalFileUri = "" Then
            LocalFileUri = "/" & UriEncode(IO.Path.GetFileName(FilePathName))  'adds a slash and builds an Uri compatible string from file name
        End If
        Dim partUri As New Uri(LocalFileUri, UriKind.Relative)       'creates a new relative Uri object
        Dim contentType As String = Net.Mime.MediaTypeNames.Application.Zip
        Dim pkgPart As PackagePart = zip.CreatePart(partUri, contentType, CompressionOption.Normal)
        'Dim pkgPart As PackagePart = zip.CreatePart(partUri, contentType, CompressionOption.Maximum)
        Dim bites As Byte() = File.ReadAllBytes(FilePathName)        'reads all of the bytes from the file to add to the zip file

        pkgPart.GetStream().Write(bites, 0, bites.Length)            'compress and write the bytes to the zip file
    End Sub

    Public Function GetWindowsUrlShortcut(ByVal FilePathName As String) As String
        'Returns the TargetPath value of the given Windows/Url (*.lnk/*.url) shortcut file name
        Dim wsh As Object
        Dim Shortcut As Object

        wsh = CreateObject("wscript.shell")
        Shortcut = wsh.CreateShortcut(FilePathName)
        GetWindowsUrlShortcut = Shortcut.TargetPath
    End Function

    Public Sub CreateUrlOrWinShortcut(ByVal strTarget As String, ByVal ShortCutFilePathName As String)
        'Creates an Url shortcut file (.url) or a Windows shortcut file (.lnk) if given file name extension is one of these two types
        Dim wsh As Object
        Dim Shortcut As Object
        Dim FileType As String

        FileType = LCase(IO.Path.GetExtension(ShortCutFilePathName))
        If FileType = ".url" Or FileType = ".lnk" Then               'manages only .url or .lnk file
            wsh = CreateObject("wscript.shell")
            Shortcut = wsh.CreateShortcut(ShortCutFilePathName)      'creates shortcut object
            Shortcut.TargetPath = strTarget
            Shortcut.Save()                                          'saves it, while eventually overwriting any existing one with same file name
        End If
    End Sub

    Public Sub CreateWinShortcut(ByVal RelPath As String, ByVal strTarget As String, ByVal strShortCutName As String, ByVal strStartIn As String)
        'Creates a Windows shortcut file (.lnk format)
        Dim wsh As Object
        Dim Shortcut As Object

        wsh = CreateObject("wscript.shell")
        Shortcut = wsh.CreateShortcut(strStartIn & "\" & strShortCutName & ".lnk")
        Shortcut.TargetPath = strTarget
        Shortcut.RelativePath = RelPath
        Shortcut.Save()
    End Sub

    Public Function GetWinShortcutTargetPath(ByVal ShortCutFilePathName As String) As String
        'Returns TargetPath of given Windows shortcut file
        Dim strTarget As String = ""
        Dim wsh As Object
        Dim Shortcut As Object
        Dim FileType As String

        FileType = LCase(IO.Path.GetExtension(ShortCutFilePathName))
        If FileType = ".url" Or FileType = ".lnk" Then               'manages only .url or .lnk file
            wsh = CreateObject("wscript.shell")
            Shortcut = wsh.CreateShortcut(ShortCutFilePathName)      'creates shortcut object
            strTarget = Shortcut.TargetPath
        End If
        GetWinShortcutTargetPath = strTarget
    End Function

    Public Sub CleanBackupFiles(ByVal FilePattern As String)
        'Cleans/deletes previously backup files with provided pattern in data folder of current database
        Dim dirSel As System.IO.DirectoryInfo
        Dim fi As System.IO.FileInfo
        Dim FilePath As String

        'Debug.Print("CleanBackupFiles: " & FilePattern)         'TEST/DEBUG
        FilePath = mZaanDbPath & "data"
        dirSel = My.Computer.FileSystem.GetDirectoryInfo(FilePath)
        For Each fi In dirSel.GetFiles(FilePattern)
            'Debug.Print("CleanBackupFiles : " & fi.Name & " modified by " & fi.LastWriteTime)      'TEST/DEBUG
            Try
                fi.Delete()                                          'deletes file defintively
            Catch ex As Exception
                Debug.Print("CleanBackupFiles error : " & ex.Message)    'TEST/DEBUG
            End Try
        Next
    End Sub

    Public Sub OpenDocument(ByVal DocPathFile As String)
        'Opens given document with corresponding application from program files

        'Debug.Print("OpenDocument =>" & DocPathFile & "<")      'TEST/DEBUG
        Try
            Dim myProcess As Process = System.Diagnostics.Process.Start(DocPathFile)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information)
        End Try
        'MessageBox.Show(myProcess.ProcessName)
    End Sub

    Public Function GetWebPageContent(ByVal PageUrl As String) As String
        'Gets page content corresponding to given Web page url
        Dim address As New Uri(PageUrl)
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse = Nothing
        Dim reader As StreamReader
        Dim result As String = ""

        Try
            request = DirectCast(WebRequest.Create(address), HttpWebRequest)   'creates Web request
            response = DirectCast(request.GetResponse(), HttpWebResponse)      'gets response
            reader = New StreamReader(response.GetResponseStream())            'gets response stream
            result = reader.ReadToEnd()                                        'gets content as a string
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Information)
            Debug.Print("! GetWebPageContent error : " & ex.Message)                   'TEST/DEBUG
        Finally
            If Not response Is Nothing Then response.Close()
        End Try
        GetWebPageContent = result
    End Function

    Public Sub InitMessages()
        'Initializes mMessage table in given language (mLanguage) from message resource file
        Dim FileContent, FileLines(), LineCells(5) As String
        Dim i, j, k, LangIndex As Integer

        'Debug.Print("InitMessages :  mLanguage = " & mLanguage)     'TEST/DEBUG
        For i = 0 To mMessage.Length - 1                   'clears mMessage table
            mMessage(i) = ""
        Next
        LangIndex = 1                                      'sets English index (col. 1) by default
        Try
            FileContent = My.Resources.messages            'get message.txt file from ZAAN resources
            FileLines = Split(FileContent, vbCrLf)
            For i = 0 To FileLines.Length - 2
                'Debug.Print("line " & i & " : " & FileLines(i))
                LineCells = Split(FileLines(i), vbTab)
                If i = 0 Then                              'first line holds Index and language Names (English/Français...)
                    For j = 1 To LineCells.Length - 1
                        If LineCells(j) = mLanguage Then
                            LangIndex = j                  'mLanguage found : stores related column index
                            'Debug.Print("LangIndex = " & LangIndex)   'TEST/DEBUG
                            Exit For
                        End If
                    Next j
                Else                                       'reads message text of LangIndex column
                    If LineCells(0) = "" Then Exit For 'terminates at first empty line
                    k = LineCells(0)
                    'Debug.Print("i=" & i & " k=" & k & " message=" & LineCells(LangIndex))   'DEBUG.PRINT
                    mMessage(k) = LineCells(LangIndex)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information)    'file info\message.txt not found !
            'Debug.Print(ex.Message)                        'file info\message.txt not found !
        End Try
    End Sub

    Public Function GetLocalMonthName(ByVal MonthNb As Integer) As String
        'Returns local name of given month number
        Dim MonthName As String = ""
        Dim n As Integer

        n = MonthNb + 100
        'Debug.Print("GetLocalMonthName :  MonthNb = " & MonthNb & "  n = " & n)     'TEST/DEBUG
        MonthName = mMessage(n)
        Return MonthName
    End Function

    Class ListViewItemComparer
        Implements IComparer

        Private col As Integer

        Public Sub New()
            col = 0
        End Sub

        Public Sub New(ByVal column As Integer)
            col = column
        End Sub

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
            Dim Result As Integer

            Select Case col                      'test listview's column
                Case 9                           'size
                    Result = Val(CType(x, ListViewItem).SubItems(col).Text).CompareTo(Val(CType(y, ListViewItem).SubItems(col).Text))
                Case 10                          'modification date
                    Result = [Date].Compare(CType(x, ListViewItem).SubItems(col).Text, CType(y, ListViewItem).SubItems(col).Text)
                Case Else                        'document name and other texts
                    Result = [String].Compare(CType(x, ListViewItem).SubItems(col).Text, CType(y, ListViewItem).SubItems(col).Text)
            End Select
            Return Result
        End Function
    End Class

    Class ListViewItemRevComparer
        Implements IComparer

        Private col As Integer

        Public Sub New()
            col = 0
        End Sub

        Public Sub New(ByVal column As Integer)
            col = column
        End Sub

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
            Dim Result As Integer

            Select Case col                      'test listview's column
                Case 9                           'size
                    Result = Val(CType(y, ListViewItem).SubItems(col).Text).CompareTo(Val(CType(x, ListViewItem).SubItems(col).Text))
                Case 10                          'modification date
                    Result = [Date].Compare(CType(y, ListViewItem).SubItems(col).Text, CType(x, ListViewItem).SubItems(col).Text)
                Case Else                        'document name and other texts
                    Result = [String].Compare(CType(y, ListViewItem).SubItems(col).Text, CType(x, ListViewItem).SubItems(col).Text)
            End Select
            Return Result
        End Function
    End Class

    Class ListView2ItemComparer
        Implements IComparer

        Private col As Integer

        Public Sub New()
            col = 0
        End Sub

        Public Sub New(ByVal column As Integer)
            col = column
        End Sub

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
            Dim Result As Integer

            Select Case col                      'test listview's column
                Case 2                           'size
                    Result = Val(CType(x, ListViewItem).SubItems(col).Text).CompareTo(Val(CType(y, ListViewItem).SubItems(col).Text))
                Case 3                           'modification date
                    Result = [Date].Compare(CType(x, ListViewItem).SubItems(col).Text, CType(y, ListViewItem).SubItems(col).Text)
                Case Else                        'document name and other texts
                    Result = [String].Compare(CType(x, ListViewItem).SubItems(col).Text, CType(y, ListViewItem).SubItems(col).Text)
            End Select
            Return Result
        End Function
    End Class

    Class ListView2ItemRevComparer
        Implements IComparer

        Private col As Integer

        Public Sub New()
            col = 0
        End Sub

        Public Sub New(ByVal column As Integer)
            col = column
        End Sub

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
            Dim Result As Integer

            Select Case col                      'test listview's column
                Case 2                           'size
                    Result = Val(CType(y, ListViewItem).SubItems(col).Text).CompareTo(Val(CType(x, ListViewItem).SubItems(col).Text))
                Case 3                           'modification date
                    Result = [Date].Compare(CType(y, ListViewItem).SubItems(col).Text, CType(x, ListViewItem).SubItems(col).Text)
                Case Else                        'document name and other texts
                    Result = [String].Compare(CType(y, ListViewItem).SubItems(col).Text, CType(x, ListViewItem).SubItems(col).Text)
            End Select
            Return Result
        End Function
    End Class

    Public Class OutlookDataObject
        Implements System.Windows.Forms.IDataObject

#Region "NativeMethods"

        Private Class NativeMethods
            <DllImport("kernel32.dll")> _
            Private Shared Function GlobalLock(ByVal hMem As IntPtr) As IntPtr
            End Function

            <DllImport("ole32.dll", PreserveSig:=False)> _
            Public Shared Function CreateILockBytesOnHGlobal(ByVal hGlobal As IntPtr, ByVal fDeleteOnRelease As Boolean) As ILockBytes
            End Function

            <DllImport("OLE32.DLL", CharSet:=CharSet.Auto, PreserveSig:=False)> _
            Public Shared Function GetHGlobalFromILockBytes(ByVal pLockBytes As ILockBytes) As IntPtr
            End Function

            <DllImport("OLE32.DLL", CharSet:=CharSet.Unicode, PreserveSig:=False)> _
            Public Shared Function StgCreateDocfileOnILockBytes(ByVal plkbyt As ILockBytes, ByVal grfMode As UInteger, ByVal reserved As UInteger) As IStorage
            End Function

            <ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000B-0000-0000-C000-000000000046")> _
            Public Interface IStorage
                Function CreateStream(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal pwcsName As String, <[In](), MarshalAs(UnmanagedType.U4)> ByVal grfMode As Integer, <[In](), MarshalAs(UnmanagedType.U4)> ByVal reserved1 As Integer, <[In](), MarshalAs(UnmanagedType.U4)> ByVal reserved2 As Integer) As <MarshalAs(UnmanagedType.[Interface])> IStream
                Function OpenStream(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal pwcsName As String, ByVal reserved1 As IntPtr, <[In](), MarshalAs(UnmanagedType.U4)> ByVal grfMode As Integer, <[In](), MarshalAs(UnmanagedType.U4)> ByVal reserved2 As Integer) As <MarshalAs(UnmanagedType.[Interface])> IStream
                Function CreateStorage(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal pwcsName As String, <[In](), MarshalAs(UnmanagedType.U4)> ByVal grfMode As Integer, <[In](), MarshalAs(UnmanagedType.U4)> ByVal reserved1 As Integer, <[In](), MarshalAs(UnmanagedType.U4)> ByVal reserved2 As Integer) As <MarshalAs(UnmanagedType.[Interface])> IStorage
                Function OpenStorage(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal pwcsName As String, ByVal pstgPriority As IntPtr, <[In](), MarshalAs(UnmanagedType.U4)> ByVal grfMode As Integer, ByVal snbExclude As IntPtr, <[In](), MarshalAs(UnmanagedType.U4)> ByVal reserved As Integer) As <MarshalAs(UnmanagedType.[Interface])> IStorage
                Sub CopyTo(ByVal ciidExclude As Integer, <[In](), MarshalAs(UnmanagedType.LPArray)> ByVal pIIDExclude As Guid(), ByVal snbExclude As IntPtr, <[In](), MarshalAs(UnmanagedType.[Interface])> ByVal stgDest As IStorage)
                Sub MoveElementTo(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal pwcsName As String, <[In](), MarshalAs(UnmanagedType.[Interface])> ByVal stgDest As IStorage, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal pwcsNewName As String, <[In](), MarshalAs(UnmanagedType.U4)> ByVal grfFlags As Integer)
                Sub Commit(ByVal grfCommitFlags As Integer)
                Sub Revert()
                Sub EnumElements(<[In](), MarshalAs(UnmanagedType.U4)> ByVal reserved1 As Integer, ByVal reserved2 As IntPtr, <[In](), MarshalAs(UnmanagedType.U4)> ByVal reserved3 As Integer, <MarshalAs(UnmanagedType.[Interface])> ByRef ppVal As Object)
                Sub DestroyElement(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal pwcsName As String)
                Sub RenameElement(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal pwcsOldName As String, <[In](), MarshalAs(UnmanagedType.BStr)> ByVal pwcsNewName As String)
                Sub SetElementTimes(<[In](), MarshalAs(UnmanagedType.BStr)> ByVal pwcsName As String, <[In]()> ByVal pctime As System.Runtime.InteropServices.ComTypes.FILETIME, <[In]()> ByVal patime As System.Runtime.InteropServices.ComTypes.FILETIME, <[In]()> ByVal pmtime As System.Runtime.InteropServices.ComTypes.FILETIME)
                Sub SetClass(<[In]()> ByRef clsid As Guid)
                Sub SetStateBits(ByVal grfStateBits As Integer, ByVal grfMask As Integer)
                Sub Stat(<Out()> ByRef pStatStg As System.Runtime.InteropServices.ComTypes.STATSTG, ByVal grfStatFlag As Integer)
            End Interface

            <ComImport(), Guid("0000000A-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
            Public Interface ILockBytes
                Sub ReadAt(<[In](), MarshalAs(UnmanagedType.U8)> ByVal ulOffset As Long, <Out(), MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> ByVal pv As Byte(), <[In](), MarshalAs(UnmanagedType.U4)> ByVal cb As Integer, <Out(), MarshalAs(UnmanagedType.LPArray)> ByVal pcbRead As Integer())
                Sub WriteAt(<[In](), MarshalAs(UnmanagedType.U8)> ByVal ulOffset As Long, ByVal pv As IntPtr, <[In](), MarshalAs(UnmanagedType.U4)> ByVal cb As Integer, <Out(), MarshalAs(UnmanagedType.LPArray)> ByVal pcbWritten As Integer())
                Sub Flush()
                Sub SetSize(<[In](), MarshalAs(UnmanagedType.U8)> ByVal cb As Long)
                Sub LockRegion(<[In](), MarshalAs(UnmanagedType.U8)> ByVal libOffset As Long, <[In](), MarshalAs(UnmanagedType.U8)> ByVal cb As Long, <[In](), MarshalAs(UnmanagedType.U4)> ByVal dwLockType As Integer)
                Sub UnlockRegion(<[In](), MarshalAs(UnmanagedType.U8)> ByVal libOffset As Long, <[In](), MarshalAs(UnmanagedType.U8)> ByVal cb As Long, <[In](), MarshalAs(UnmanagedType.U4)> ByVal dwLockType As Integer)
                Sub Stat(<Out()> ByRef pstatstg As System.Runtime.InteropServices.ComTypes.STATSTG, <[In](), MarshalAs(UnmanagedType.U4)> ByVal grfStatFlag As Integer)
            End Interface

            <StructLayout(LayoutKind.Sequential)> _
            Public NotInheritable Class POINTL
                Public x As Integer
                Public y As Integer
            End Class

            <StructLayout(LayoutKind.Sequential)> _
            Public NotInheritable Class SIZEL
                Public cx As Integer
                Public cy As Integer
            End Class

            <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)> _
            Public NotInheritable Class FILEGROUPDESCRIPTORA
                Public cItems As UInteger
                Public fgd As FILEDESCRIPTORA()
            End Class

            <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)> _
            Public NotInheritable Class FILEDESCRIPTORA
                Public dwFlags As UInteger
                Public clsid As Guid
                Public sizel As SIZEL
                Public pointl As POINTL
                Public dwFileAttributes As UInteger
                Public ftCreationTime As System.Runtime.InteropServices.ComTypes.FILETIME
                Public ftLastAccessTime As System.Runtime.InteropServices.ComTypes.FILETIME
                Public ftLastWriteTime As System.Runtime.InteropServices.ComTypes.FILETIME
                Public nFileSizeHigh As UInteger
                Public nFileSizeLow As UInteger
                <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> _
                Public cFileName As String
            End Class

            <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
            Public NotInheritable Class FILEGROUPDESCRIPTORW
                Public cItems As UInteger
                Public fgd As FILEDESCRIPTORW()
            End Class

            <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
            Public NotInheritable Class FILEDESCRIPTORW
                Public dwFlags As UInteger
                Public clsid As Guid
                Public sizel As SIZEL
                Public pointl As POINTL
                Public dwFileAttributes As UInteger
                Public ftCreationTime As System.Runtime.InteropServices.ComTypes.FILETIME
                Public ftLastAccessTime As System.Runtime.InteropServices.ComTypes.FILETIME
                Public ftLastWriteTime As System.Runtime.InteropServices.ComTypes.FILETIME
                Public nFileSizeHigh As UInteger
                Public nFileSizeLow As UInteger
                <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> _
                Public cFileName As String
            End Class
        End Class

#End Region

#Region "Property(s)"

        Private underlyingDataObject As System.Windows.Forms.IDataObject
        ' Holds the "System.Windows.Forms.IDataObject" that this class is wrapping

        Private comUnderlyingDataObject As System.Runtime.InteropServices.ComTypes.IDataObject
        ' Holds the "System.Runtime.InteropServices.ComTypes.IDataObject" interface to the "System.Windows.Forms.IDataObject" that this class is wrapping

        Private oleUnderlyingDataObject As System.Windows.Forms.IDataObject
        ' Holds the internal ole "System.Windows.Forms.IDataObject" to the "System.Windows.Forms.IDataObject" that this class is wrapping

        Private getDataFromHGLOBLALMethod As MethodInfo
        ' Holds the "MethodInfo" of the "GetDataFromHGLOBLAL" method of the internal ole

#End Region

#Region "Constructor(s)"

        Public Sub New(ByVal underlyingDataObject As System.Windows.Forms.IDataObject)
            'get the underlying dataobject and its ComType IDataObject interface to it
            Me.underlyingDataObject = underlyingDataObject
            Me.comUnderlyingDataObject = DirectCast(Me.underlyingDataObject, System.Runtime.InteropServices.ComTypes.IDataObject)

            'get the internal ole dataobject and its GetDataFromHGLOBLAL so it can be called later
            Dim innerDataField As FieldInfo = Me.underlyingDataObject.[GetType]().GetField("innerData", BindingFlags.NonPublic Or BindingFlags.Instance)
            Me.oleUnderlyingDataObject = DirectCast(innerDataField.GetValue(Me.underlyingDataObject), System.Windows.Forms.IDataObject)
            Me.getDataFromHGLOBLALMethod = Me.oleUnderlyingDataObject.[GetType]().GetMethod("GetDataFromHGLOBLAL", BindingFlags.NonPublic Or BindingFlags.Instance)
        End Sub

#End Region

#Region "IDataObject Members"

        Public Function GetData(ByVal format As Type) As Object Implements System.Windows.Forms.IDataObject.GetData
            Return Me.GetData(format.FullName)
        End Function

        Public Function GetData(ByVal format As String) As Object Implements System.Windows.Forms.IDataObject.GetData
            Return Me.GetData(format, True)
        End Function

        Public Function GetData(ByVal format As String, ByVal autoConvert As Boolean) As Object Implements System.Windows.Forms.IDataObject.GetData
            'handle the "FileGroupDescriptor" and "FileContents" format request in this class otherwise pass through to underlying IDataObject 
            Select Case format
                Case "FileGroupDescriptor"
                    'override the default handling of FileGroupDescriptor which returns a MemoryStream and instead return a string array of file names
                    Dim fileGroupDescriptorAPointer As IntPtr = IntPtr.Zero
                    Try
                        'use the underlying IDataObject to get the FileGroupDescriptor as a MemoryStream
                        Dim fileGroupDescriptorStream As MemoryStream = DirectCast(Me.underlyingDataObject.GetData("FileGroupDescriptor", autoConvert), MemoryStream)
                        Dim fileGroupDescriptorBytes As Byte() = New Byte(fileGroupDescriptorStream.Length - 1) {}
                        fileGroupDescriptorStream.Read(fileGroupDescriptorBytes, 0, fileGroupDescriptorBytes.Length)
                        fileGroupDescriptorStream.Close()

                        'copy the file group descriptor into unmanaged memory 
                        fileGroupDescriptorAPointer = Marshal.AllocHGlobal(fileGroupDescriptorBytes.Length)
                        Marshal.Copy(fileGroupDescriptorBytes, 0, fileGroupDescriptorAPointer, fileGroupDescriptorBytes.Length)

                        'marshal the unmanaged memory to to FILEGROUPDESCRIPTORA struct
                        Dim fileGroupDescriptorObject As Object = Marshal.PtrToStructure(fileGroupDescriptorAPointer, GetType(NativeMethods.FILEGROUPDESCRIPTORA))
                        Dim fileGroupDescriptor As NativeMethods.FILEGROUPDESCRIPTORA = DirectCast(fileGroupDescriptorObject, NativeMethods.FILEGROUPDESCRIPTORA)

                        'create a new array to store file names in of the number of items in the file group descriptor
                        Dim fileNames As String() = New String(fileGroupDescriptor.cItems - 1) {}

                        'get the pointer to the first file descriptor
                        Dim fileDescriptorPointer As IntPtr = CType(CInt(fileGroupDescriptorAPointer) + Marshal.SizeOf(fileGroupDescriptorAPointer), IntPtr)

                        'loop for the number of files acording to the file group descriptor
                        For fileDescriptorIndex As Integer = 0 To fileGroupDescriptor.cItems - 1
                            'marshal the pointer top the file descriptor as a FILEDESCRIPTORA struct and get the file name
                            Dim fileDescriptor As NativeMethods.FILEDESCRIPTORA = DirectCast(Marshal.PtrToStructure(fileDescriptorPointer, GetType(NativeMethods.FILEDESCRIPTORA)), NativeMethods.FILEDESCRIPTORA)
                            fileNames(fileDescriptorIndex) = fileDescriptor.cFileName

                            'move the file descriptor pointer to the next file descriptor
                            fileDescriptorPointer = CType(CInt(fileDescriptorPointer) + Marshal.SizeOf(fileDescriptor), IntPtr)
                        Next

                        'return the array of filenames
                        Return fileNames
                    Finally
                        'free unmanaged memory pointer
                        Marshal.FreeHGlobal(fileGroupDescriptorAPointer)
                    End Try

                Case "FileGroupDescriptorW"
                    'override the default handling of FileGroupDescriptorW which returns a MemoryStream and instead return a string array of file names
                    Dim fileGroupDescriptorWPointer As IntPtr = IntPtr.Zero
                    Try
                        'use the underlying IDataObject to get the FileGroupDescriptorW as a MemoryStream
                        Dim fileGroupDescriptorStream As MemoryStream = DirectCast(Me.underlyingDataObject.GetData("FileGroupDescriptorW"), MemoryStream)
                        Dim fileGroupDescriptorBytes As Byte() = New Byte(fileGroupDescriptorStream.Length - 1) {}
                        fileGroupDescriptorStream.Read(fileGroupDescriptorBytes, 0, fileGroupDescriptorBytes.Length)
                        fileGroupDescriptorStream.Close()

                        'copy the file group descriptor into unmanaged memory
                        fileGroupDescriptorWPointer = Marshal.AllocHGlobal(fileGroupDescriptorBytes.Length)
                        Marshal.Copy(fileGroupDescriptorBytes, 0, fileGroupDescriptorWPointer, fileGroupDescriptorBytes.Length)

                        'marshal the unmanaged memory to to FILEGROUPDESCRIPTORW struct
                        Dim fileGroupDescriptorObject As Object = Marshal.PtrToStructure(fileGroupDescriptorWPointer, GetType(NativeMethods.FILEGROUPDESCRIPTORW))
                        Dim fileGroupDescriptor As NativeMethods.FILEGROUPDESCRIPTORW = DirectCast(fileGroupDescriptorObject, NativeMethods.FILEGROUPDESCRIPTORW)

                        'create a new array to store file names in of the number of items in the file group descriptor
                        Dim fileNames As String() = New String(fileGroupDescriptor.cItems - 1) {}

                        'get the pointer to the first file descriptor
                        Dim fileDescriptorPointer As IntPtr = CType(CInt(fileGroupDescriptorWPointer) + Marshal.SizeOf(fileGroupDescriptorWPointer), IntPtr)

                        'loop for the number of files acording to the file group descriptor
                        For fileDescriptorIndex As Integer = 0 To fileGroupDescriptor.cItems - 1
                            'marshal the pointer top the file descriptor as a FILEDESCRIPTORW struct and get the file name
                            Dim fileDescriptor As NativeMethods.FILEDESCRIPTORW = DirectCast(Marshal.PtrToStructure(fileDescriptorPointer, GetType(NativeMethods.FILEDESCRIPTORW)), NativeMethods.FILEDESCRIPTORW)
                            fileNames(fileDescriptorIndex) = fileDescriptor.cFileName

                            'move the file descriptor pointer to the next file descriptor
                            fileDescriptorPointer = CType(CInt(fileDescriptorPointer) + Marshal.SizeOf(fileDescriptor), IntPtr)
                        Next

                        'return the array of filenames
                        Return fileNames
                    Finally
                        'free unmanaged memory pointer
                        Marshal.FreeHGlobal(fileGroupDescriptorWPointer)
                    End Try

                Case "FileContents"
                    'override the default handling of FileContents which returns the contents of the first file as a memory stream
                    'and instead return an array of MemoryStreams containing the data to each file dropped

                    'get the array of filenames which lets us know how many file contents exist
                    Dim fileContentNames As String() = DirectCast(Me.GetData("FileGroupDescriptor"), String())

                    'create a MemoryStream array to store the file contents
                    Dim fileContents As MemoryStream() = New MemoryStream(fileContentNames.Length - 1) {}

                    'loop for the number of files acording to the file names
                    For fileIndex As Integer = 0 To fileContentNames.Length - 1
                        'get the data at the file index and store in array
                        fileContents(fileIndex) = Me.GetData(format, fileIndex)
                    Next

                    'return array of MemoryStreams containing file contents
                    Return fileContents
            End Select

            'use underlying IDataObject to handle getting of data
            Return Me.underlyingDataObject.GetData(format, autoConvert)
        End Function

        Public Function GetData(ByVal format As String, ByVal index As Integer) As MemoryStream
            'create a FORMATETC struct to request the data with
            Dim formatetc As New FORMATETC()
            'formatetc.cfFormat = CShort(DataFormats.GetFormat(format).Id)

            Dim myFormatID As Integer = Forms.DataFormats.GetFormat(format).Id
            If myFormatID > 32767 Then                                         'WARNING : CORRECTS OVERFLOW ISSUE IN VB (DOES NOT HAPPEN IN C# !!!)
                myFormatID = myFormatID - 65536
            End If
            formatetc.cfFormat = CShort(myFormatID)

            formatetc.dwAspect = DVASPECT.DVASPECT_CONTENT
            formatetc.lindex = index
            formatetc.ptd = New IntPtr(0)
            formatetc.tymed = TYMED.TYMED_ISTREAM Or TYMED.TYMED_ISTORAGE Or TYMED.TYMED_HGLOBAL

            'create STGMEDIUM to output request results into
            Dim medium As New STGMEDIUM()

            'using the Com IDataObject interface get the data using the defined FORMATETC
            Me.comUnderlyingDataObject.GetData(formatetc, medium)

            'retrieve the data depending on the returned store type
            Select Case medium.tymed
                Case TYMED.TYMED_ISTORAGE
                    'to handle a IStorage it needs to be written into a second unmanaged memory mapped storage
                    'and then the data can be read from memory into a managed byte and returned as a MemoryStream

                    Dim iStorage As NativeMethods.IStorage = Nothing
                    Dim iStorage2 As NativeMethods.IStorage = Nothing
                    Dim iLockBytes As NativeMethods.ILockBytes = Nothing
                    Dim iLockBytesStat As System.Runtime.InteropServices.ComTypes.STATSTG
                    Try
                        'marshal the returned pointer to a IStorage object
                        iStorage = DirectCast(Marshal.GetObjectForIUnknown(medium.unionmember), NativeMethods.IStorage)
                        Marshal.Release(medium.unionmember)

                        'create a ILockBytes (unmanaged byte array) and then create a IStorage using the byte array as a backing store
                        iLockBytes = NativeMethods.CreateILockBytesOnHGlobal(IntPtr.Zero, True)
                        iStorage2 = NativeMethods.StgCreateDocfileOnILockBytes(iLockBytes, &H1012, 0)

                        'copy the returned IStorage into the new IStorage
                        iStorage.CopyTo(0, Nothing, IntPtr.Zero, iStorage2)
                        iLockBytes.Flush()
                        iStorage2.Commit(0)

                        'get the STATSTG of the ILockBytes to determine how many bytes were written to it
                        iLockBytesStat = New System.Runtime.InteropServices.ComTypes.STATSTG()
                        iLockBytes.Stat(iLockBytesStat, 1)
                        Dim iLockBytesSize As Integer = CInt(iLockBytesStat.cbSize)

                        'read the data from the ILockBytes (unmanaged byte array) into a managed byte array
                        Dim iLockBytesContent As Byte() = New Byte(iLockBytesSize - 1) {}
                        iLockBytes.ReadAt(0, iLockBytesContent, iLockBytesContent.Length, Nothing)

                        'wrapped the managed byte array into a memory stream and return it
                        Return New MemoryStream(iLockBytesContent)
                    Finally
                        'release all unmanaged objects
                        Marshal.ReleaseComObject(iStorage2)
                        Marshal.ReleaseComObject(iLockBytes)
                        Marshal.ReleaseComObject(iStorage)
                    End Try

                Case TYMED.TYMED_ISTREAM
                    'to handle a IStream it needs to be read into a managed byte and returned as a MemoryStream

                    Dim iStream As IStream = Nothing
                    Dim iStreamStat As System.Runtime.InteropServices.ComTypes.STATSTG
                    Try
                        'marshal the returned pointer to a IStream object
                        iStream = DirectCast(Marshal.GetObjectForIUnknown(medium.unionmember), IStream)
                        Marshal.Release(medium.unionmember)

                        'get the STATSTG of the IStream to determine how many bytes are in it
                        iStreamStat = New System.Runtime.InteropServices.ComTypes.STATSTG()
                        iStream.Stat(iStreamStat, 0)
                        Dim iStreamSize As Integer = CInt(iStreamStat.cbSize)

                        'read the data from the IStream into a managed byte array
                        Dim iStreamContent As Byte() = New Byte(iStreamSize - 1) {}
                        iStream.Read(iStreamContent, iStreamContent.Length, IntPtr.Zero)

                        'wrapped the managed byte array into a memory stream and return it
                        Return New MemoryStream(iStreamContent)
                    Finally
                        'release all unmanaged objects
                        Marshal.ReleaseComObject(iStream)
                    End Try

                Case TYMED.TYMED_HGLOBAL
                    'to handle a HGlobal the exisitng "GetDataFromHGLOBLAL" method is invoked via reflection

                    Return DirectCast(Me.getDataFromHGLOBLALMethod.Invoke(Me.oleUnderlyingDataObject, New Object() {Forms.DataFormats.GetFormat(CShort(formatetc.cfFormat)).Name, medium.unionmember}), MemoryStream)
            End Select

            Return Nothing
        End Function

        Public Function GetDataPresent(ByVal format As Type) As Boolean Implements System.Windows.Forms.IDataObject.GetDataPresent
            Return Me.underlyingDataObject.GetDataPresent(format)
        End Function

        Public Function GetDataPresent(ByVal format As String) As Boolean Implements System.Windows.Forms.IDataObject.GetDataPresent
            Return Me.underlyingDataObject.GetDataPresent(format)
        End Function

        Public Function GetDataPresent(ByVal format As String, ByVal autoConvert As Boolean) As Boolean Implements System.Windows.Forms.IDataObject.GetDataPresent
            Return Me.underlyingDataObject.GetDataPresent(format, autoConvert)
        End Function

        Public Function GetFormats() As String() Implements System.Windows.Forms.IDataObject.GetFormats
            Return Me.underlyingDataObject.GetFormats()
        End Function

        Public Function GetFormats(ByVal autoConvert As Boolean) As String() Implements System.Windows.Forms.IDataObject.GetFormats
            Return Me.underlyingDataObject.GetFormats(autoConvert)
        End Function

        Public Sub SetData(ByVal data As Object) Implements System.Windows.Forms.IDataObject.SetData
            Me.underlyingDataObject.SetData(data)
        End Sub

        Public Sub SetData(ByVal format As Type, ByVal data As Object) Implements System.Windows.Forms.IDataObject.SetData
            Me.underlyingDataObject.SetData(format, data)
        End Sub

        Public Sub SetData(ByVal format As String, ByVal data As Object) Implements System.Windows.Forms.IDataObject.SetData
            Me.underlyingDataObject.SetData(format, data)
        End Sub

        Public Sub SetData(ByVal format As String, ByVal autoConvert As Boolean, ByVal data As Object) Implements System.Windows.Forms.IDataObject.SetData
            Me.underlyingDataObject.SetData(format, autoConvert, data)
        End Sub

#End Region
    End Class

End Module
