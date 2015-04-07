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

Public NotInheritable Class frmAboutBox

    Private mAboutZaanIni As Boolean

    Private Sub InitLanguageDisplay()
        'Displays current language selection
        cmbLanguage.Text = mMessage(58)                              'EN : English
    End Sub

    Private Sub InitLocalLabels()
        'Sets labels and tooltips in current local language
        Me.Text = mMessage(47)                                       'About ZAAN
        tlTip.SetToolTip(LogoPictureBox, "www.zaan.com")             'www.zaan.com
    End Sub

    Private Sub frmAboutBox_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If mLicAccepted = "0" Then
            mUserLicEndDate = "."                                    'this flag will indicate to ZAAN main form to close the application after a warning message !
        End If
    End Sub

    Private Sub frmAboutBox_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Start of "About ZAAN" form
        mAboutZaanIni = True                                         'set AboutZaanIni flag to true for avoiding cmbLanguage change event to reload mMessage table
        Call InitLanguageDisplay()                                   'displays current language selection
        Call InitLocalLabels()                                       'sets labels and tooltips in current local language

        If My.Application.IsNetworkDeployed Then                     'case of ZAAN application deployed with ClickOnce automatic update from www.zaan.com
            Me.LabelProductName.Text = My.Application.Info.ProductName & " v" & My.Application.Deployment.CurrentVersion.ToString   'ONLY FOR DEPLOYEMENT !!!
            mUserManualMainVer = Mid(My.Application.Deployment.CurrentVersion.ToString, 1, 1)
            mUserManualMinorVer = Mid(My.Application.Deployment.CurrentVersion.ToString, 3, 1)
        Else                                                         'case of ZAAN application is run/debugged locally 
            Me.LabelProductName.Text = My.Application.Info.ProductName & " v" & My.Application.Info.Version.ToString
            mUserManualMainVer = My.Application.Info.Version.Major
            mUserManualMinorVer = My.Application.Info.Version.Minor
        End If

        Me.LabelProductDescription.Text = My.Application.Info.Description
        Me.LabelCopyright.Text = My.Application.Info.Copyright
        Me.LabelCompanyName.Text = My.Application.Info.CompanyName
        Me.TextBoxDescription.Text = "This program is free software: you can redistribute it and/or modify " & _
            "it under the terms of the GNU General Public License as published by " & _
            "the Free Software Foundation, either version 3 of the License, or " & _
            "(at your option) any later version." & vbCrLf & _
            "This program is distributed in the hope that it will be useful, " & _
            "but WITHOUT ANY WARRANTY; without even the implied warranty of " & _
            "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the " & _
            "GNU General Public License for more details." & vbCrLf & _
            "You should have received a copy of the GNU General Public License " & _
            "along with this program. If not, see <http://www.gnu.org/licenses/>."

        If mAboutBoxAutoClose Then
            mAboutBoxAutoClose = False
            If mLicAccepted = "1" Then                               'case of ZAAN license already accepted : enables auto-exit
                tmrAboutZaanExit.Enabled = True                      'starts exit timer
            End If
        End If
    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
        mLicAccepted = "1"                                           'save user's OK for ZAAN license acceptation
        Me.Close()
    End Sub

    Private Sub LogoPictureBox_Click(sender As Object, e As EventArgs) Handles LogoPictureBox.Click
        OpenDocument("http://www.zaan.com")                          'opens home page of ZAAN.COM Web site
    End Sub

    Private Sub cmbLanguage_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbLanguage.SelectedIndexChanged
        'Updates mLanguage ("EN" or "FR"...) and initializes mMessage table and local labels
        If mAboutZaanIni Then
            mAboutZaanIni = False                                    'enables later cmbLanguage change event to reload mMessages (see below)
        Else
            mLanguage = Microsoft.VisualBasic.Left(cmbLanguage.Text, 2)
            Call InitMessages()                                      'initializes mMessage table in English by default (for displaying initialization errors)
            Call InitLocalLabels()                                   'sets ZAAN application description
        End If
    End Sub

    Private Sub tmrAboutZaanExit_Tick(sender As Object, e As EventArgs) Handles tmrAboutZaanExit.Tick
        tmrAboutZaanExit.Enabled = False                             'stops timer
        Me.Close()                                                   'exits About ZAAN form
    End Sub

End Class
