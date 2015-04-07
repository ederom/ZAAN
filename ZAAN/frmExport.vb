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

Public Class frmExport

    Private Sub LoadDimOrderList()
        'Loads dimensions order list with all possible combinations of When/Who/What/Where
        Dim i, j, k, l As Integer

        cmbDimOrder.Items.Clear()
        cmbDimOrder.Items.Add(mMessage(172))                         '<separated tree views>
        For i = 1 To 4
            For j = 1 To 4
                If j <> i Then
                    For k = 1 To 4
                        If (k <> i) And (k <> j) Then
                            For l = 1 To 4
                                If (l <> i) And (l <> j) And (l <> k) Then
                                    cmbDimOrder.Items.Add(mMessage(i) & " \ " & mMessage(j) & " \ " & mMessage(k) & " \ " & mMessage(l))
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        Next
        cmbDimOrder.SelectedIndex = mTreeCodeSeriesIndex             'sets stored dimensions order index
    End Sub

    Private Sub SetExportDBwinTreeCodeSeries()
        'Sets mTreeCodeSeries (like "t o a e b c") to be used by ExportDataToWindows() sub
        Dim TreeCodeSeries As String = ""
        Dim TreeRoots() As String
        Dim i As Integer

        mTreeCodeSeries = ""
        If Mid(cmbDimOrder.Text, 1, 1) <> "<" Then       'selected option is not the first one <separated tree views>
            TreeRoots = Split(cmbDimOrder.Text, " \ ")
            For i = 0 To TreeRoots.Length - 1
                Select Case TreeRoots(i)
                    Case mMessage(1)                                 'When
                        TreeCodeSeries = TreeCodeSeries & "t" & " "
                    Case mMessage(2)                                 'Who
                        TreeCodeSeries = TreeCodeSeries & "o" & " "
                    Case mMessage(3)                                 'What
                        TreeCodeSeries = TreeCodeSeries & "a" & " "
                    Case mMessage(4)                                 'Where
                        TreeCodeSeries = TreeCodeSeries & "e" & " "
                End Select
            Next
            If TreeCodeSeries <> "" Then
                mTreeCodeSeries = TreeCodeSeries & "b c"       'appends What2 and Who2
            End If
        End If
        'Debug.Print("SetExportDBwinTreeCodeSeries :  mTreeCodeSeries = " & mTreeCodeSeries)   'TEST/DEBUG
    End Sub

    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        If rbNow.Checked Then
            mExportDBwinTime = ""                                    'clears db export start time
            frmZaan.tmrDbExport.Interval = 2000                      'sets 2 s delay before launching ExportDatabaseToWindows() function
            frmZaan.tmrDbExport.Enabled = True                       'starts db export timer
        End If
        Me.Close()                                                   'exits Export form
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        rbNow.Checked = True                                         'sets now (=> will clear db export start time)
        Me.Close()                                                   'exits Export form
    End Sub

    Private Sub cmbDimOrder_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbDimOrder.SelectedIndexChanged
        'Debug.Print("cmbDimOrder_SelectedIndexChanged")    'TEST/DEBUG
        mTreeCodeSeriesIndex = cmbDimOrder.SelectedIndex             'stores selected dimensions order index
        Call SetExportDBwinTreeCodeSeries()                          'sets mTreeCodeSeries (like "o t a e b c")
        'Debug.Print("=> cmbDimOrder_SelectedIndexChanged :  mTreeCodeSeriesIndex = " & mTreeCodeSeriesIndex & "  mTreeCodeSeries = " & mTreeCodeSeries)   'TEST/DEBUG
    End Sub

    Private Sub dtpStartTime_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpStartTime.ValueChanged
        'Debug.Print("> dtpStartTime_ValueChanged : " & dtpStartTime.Value)    'TEST/DEBUG
        rbNow.Checked = Not dtpStartTime.Checked
        'If dtpStartTime.Checked Then
        mExportDBwinTime = CStr(dtpStartTime.Value)              'stores db export start time
        'End If
    End Sub

    Private Sub rbNow_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNow.CheckedChanged
        'Debug.Print("rbNow_CheckedChanged")    'TEST/DEBUG
        If rbNow.Checked Then
            dtpStartTime.Checked = False                             'disables export start time
            rbOnce.Checked = True                                    'cancels repeat mode
        End If
    End Sub

    Private Sub txtDestPath_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDestPath.Click
        fbdDestPath.SelectedPath = txtDestPath.Text
        fbdDestPath.Description = mMessage(70)                       'Select a destination folder
        fbdDestPath.ShowNewFolderButton = True                       'enables new folder button display
        If fbdDestPath.ShowDialog = Windows.Forms.DialogResult.OK Then
            txtDestPath.Text = fbdDestPath.SelectedPath              'updates selected path
            mExportDBwinDest = txtDestPath.Text
        End If
    End Sub

    Private Sub frmExport_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        txtDestPath.Text = mExportDBwinDest                          'sets destination path
        Call LoadDimOrderList()                                      'loads dimensions order list with all possible combinations of When/Who/What/Wher

        If mExportDBwinTime = "" Then
            dtpStartTime.Value = Now                                 'sets db export start date/time to now
            rbNow.Checked = True
        Else
            dtpStartTime.Value = CDate(mExportDBwinTime)             'sets stored db export start date/time
        End If

        Select Case mExportDBwinRepeatIndex
            Case 0
                rbOnce.Checked = True
            Case 1
                rbEveryDay.Checked = True
            Case 2
                rbEveryWeek.Checked = True
            Case 3
                rbEveryMonth.Checked = True
        End Select
    End Sub

    Private Sub frmExport_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If rbNow.Checked Then
            mExportDBwinTime = ""                                    'clears db export start time
        End If
        If mExportDBwinTime = "" Then
            'frmZaan.tmrDbExport.Enabled = False                      'makes sure that db export timer is stopped
            mZaanTitleOption = ""                                    'clears form title display option
            Call frmZaan.InitZaanFormTitle()                         'updates ZAAN main form title
            frmZaan.tsmiSelectorExportDBwin.Checked = False          'unchecks export control in selector menu
        Else
            frmZaan.tmrDbExport.Interval = 60000                     'sets 60 s delay before launching ExportDatabaseToWindows() function
            frmZaan.tmrDbExport.Enabled = True                       'starts db export timer
            mZaanTitleOption = " (auto)"                             'sets form title display option with auto (export) indicator
            Call frmZaan.InitZaanFormTitle()                         'updates ZAAN main form title
            frmZaan.tsmiSelectorExportDBwin.Checked = True           'checks export control in selector menu
            MsgBox(mMessage(220) & vbCrLf & mExportDBwinTime, MsgBoxStyle.Information)   'Next database export is scheduled at...
        End If
    End Sub

    Private Sub frmExport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        frmZaan.tmrDbExport.Enabled = False                          'makes sure that db export timer is stopped
        Me.Text = mMessage(100)                                      'Export database to Windows
        lblDestPath.Text = mMessage(215)                             'Destination folder :
        lblDimOrder.Text = mMessage(216)                             'Dimensions order :
        rbNow.Text = mMessage(217)                                   'Now

        rbOnce.Text = mMessage(211)                                  'Only once
        rbEveryDay.Text = mMessage(212)                              'Every day
        rbEveryWeek.Text = mMessage(213)                             'Every week
        rbEveryMonth.Text = mMessage(214)                            'Every month

        btnExport.Text = mMessage(218)                               'Export
        btnCancel.Text = mMessage(219)                               'Cancel
    End Sub

    Private Sub rbOnce_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbOnce.CheckedChanged
        If rbOnce.Checked Then
            mExportDBwinRepeatIndex = 0                              'updates export repeat index
        End If
    End Sub

    Private Sub rbEveryDay_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbEveryDay.CheckedChanged
        If rbEveryDay.Checked Then
            mExportDBwinRepeatIndex = 1                              'updates export repeat index
            dtpStartTime.Checked = True                              'enables export start time
            rbNow.Checked = False                                    'disables now
        End If
    End Sub

    Private Sub rbEveryWeek_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbEveryWeek.CheckedChanged
        If rbEveryWeek.Checked Then
            mExportDBwinRepeatIndex = 2                              'updates export repeat index
            dtpStartTime.Checked = True                              'enables export start time
            rbNow.Checked = False                                    'disables now
        End If
    End Sub

    Private Sub rbEveryMonth_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbEveryMonth.CheckedChanged
        If rbEveryMonth.Checked Then
            mExportDBwinRepeatIndex = 3                              'updates export repeat index
            dtpStartTime.Checked = True                              'enables export start time
            rbNow.Checked = False                                    'disables now
        End If
    End Sub
End Class