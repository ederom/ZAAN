<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmExport
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmExport))
        Me.lblDestPath = New System.Windows.Forms.Label()
        Me.txtDestPath = New System.Windows.Forms.TextBox()
        Me.lblDimOrder = New System.Windows.Forms.Label()
        Me.cmbDimOrder = New System.Windows.Forms.ComboBox()
        Me.rbNow = New System.Windows.Forms.RadioButton()
        Me.dtpStartTime = New System.Windows.Forms.DateTimePicker()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.rbEveryMonth = New System.Windows.Forms.RadioButton()
        Me.rbEveryWeek = New System.Windows.Forms.RadioButton()
        Me.rbEveryDay = New System.Windows.Forms.RadioButton()
        Me.rbOnce = New System.Windows.Forms.RadioButton()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.fbdDestPath = New System.Windows.Forms.FolderBrowserDialog()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblDestPath
        '
        Me.lblDestPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDestPath.Location = New System.Drawing.Point(9, 19)
        Me.lblDestPath.Name = "lblDestPath"
        Me.lblDestPath.Size = New System.Drawing.Size(163, 21)
        Me.lblDestPath.TabIndex = 1
        Me.lblDestPath.Text = "Destination folder :"
        Me.lblDestPath.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDestPath
        '
        Me.txtDestPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDestPath.Location = New System.Drawing.Point(171, 19)
        Me.txtDestPath.Name = "txtDestPath"
        Me.txtDestPath.ReadOnly = True
        Me.txtDestPath.Size = New System.Drawing.Size(333, 21)
        Me.txtDestPath.TabIndex = 2
        '
        'lblDimOrder
        '
        Me.lblDimOrder.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDimOrder.Location = New System.Drawing.Point(9, 55)
        Me.lblDimOrder.Name = "lblDimOrder"
        Me.lblDimOrder.Size = New System.Drawing.Size(163, 21)
        Me.lblDimOrder.TabIndex = 3
        Me.lblDimOrder.Text = "Dimensions order :"
        Me.lblDimOrder.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbDimOrder
        '
        Me.cmbDimOrder.FormattingEnabled = True
        Me.cmbDimOrder.Location = New System.Drawing.Point(171, 57)
        Me.cmbDimOrder.Name = "cmbDimOrder"
        Me.cmbDimOrder.Size = New System.Drawing.Size(333, 21)
        Me.cmbDimOrder.TabIndex = 4
        '
        'rbNow
        '
        Me.rbNow.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbNow.Location = New System.Drawing.Point(7, 11)
        Me.rbNow.Name = "rbNow"
        Me.rbNow.Size = New System.Drawing.Size(90, 20)
        Me.rbNow.TabIndex = 5
        Me.rbNow.Text = "Now"
        Me.rbNow.UseVisualStyleBackColor = True
        '
        'dtpStartTime
        '
        Me.dtpStartTime.Checked = False
        Me.dtpStartTime.CustomFormat = "dddd dd MMMM yyyy  -  H:mm:ss"
        Me.dtpStartTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpStartTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStartTime.Location = New System.Drawing.Point(122, 10)
        Me.dtpStartTime.Name = "dtpStartTime"
        Me.dtpStartTime.ShowCheckBox = True
        Me.dtpStartTime.ShowUpDown = True
        Me.dtpStartTime.Size = New System.Drawing.Size(333, 21)
        Me.dtpStartTime.TabIndex = 6
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbNow)
        Me.GroupBox1.Controls.Add(Me.dtpStartTime)
        Me.GroupBox1.Location = New System.Drawing.Point(49, 91)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(455, 37)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.rbEveryMonth)
        Me.GroupBox2.Controls.Add(Me.rbEveryWeek)
        Me.GroupBox2.Controls.Add(Me.rbEveryDay)
        Me.GroupBox2.Controls.Add(Me.rbOnce)
        Me.GroupBox2.Location = New System.Drawing.Point(49, 141)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(456, 37)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        '
        'rbEveryMonth
        '
        Me.rbEveryMonth.AutoSize = True
        Me.rbEveryMonth.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbEveryMonth.Location = New System.Drawing.Point(355, 12)
        Me.rbEveryMonth.Name = "rbEveryMonth"
        Me.rbEveryMonth.Size = New System.Drawing.Size(92, 19)
        Me.rbEveryMonth.TabIndex = 3
        Me.rbEveryMonth.Text = "Every month"
        Me.rbEveryMonth.UseVisualStyleBackColor = True
        '
        'rbEveryWeek
        '
        Me.rbEveryWeek.AutoSize = True
        Me.rbEveryWeek.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbEveryWeek.Location = New System.Drawing.Point(237, 12)
        Me.rbEveryWeek.Name = "rbEveryWeek"
        Me.rbEveryWeek.Size = New System.Drawing.Size(86, 19)
        Me.rbEveryWeek.TabIndex = 2
        Me.rbEveryWeek.Text = "Every week"
        Me.rbEveryWeek.UseVisualStyleBackColor = True
        '
        'rbEveryDay
        '
        Me.rbEveryDay.AutoSize = True
        Me.rbEveryDay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbEveryDay.Location = New System.Drawing.Point(122, 12)
        Me.rbEveryDay.Name = "rbEveryDay"
        Me.rbEveryDay.Size = New System.Drawing.Size(76, 19)
        Me.rbEveryDay.TabIndex = 1
        Me.rbEveryDay.Text = "Every day"
        Me.rbEveryDay.UseVisualStyleBackColor = True
        '
        'rbOnce
        '
        Me.rbOnce.AutoSize = True
        Me.rbOnce.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbOnce.Location = New System.Drawing.Point(7, 12)
        Me.rbOnce.Name = "rbOnce"
        Me.rbOnce.Size = New System.Drawing.Size(79, 19)
        Me.rbOnce.TabIndex = 0
        Me.rbOnce.Text = "Only once"
        Me.rbOnce.UseVisualStyleBackColor = True
        '
        'btnExport
        '
        Me.btnExport.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExport.Location = New System.Drawing.Point(49, 194)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(123, 29)
        Me.btnExport.TabIndex = 10
        Me.btnExport.Text = "OK"
        Me.btnExport.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(381, 194)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(123, 29)
        Me.btnCancel.TabIndex = 11
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'frmExport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(564, 239)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnExport)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmbDimOrder)
        Me.Controls.Add(Me.lblDimOrder)
        Me.Controls.Add(Me.txtDestPath)
        Me.Controls.Add(Me.lblDestPath)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmExport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Export ZAAN database to Windows"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblDestPath As System.Windows.Forms.Label
    Friend WithEvents txtDestPath As System.Windows.Forms.TextBox
    Friend WithEvents lblDimOrder As System.Windows.Forms.Label
    Friend WithEvents cmbDimOrder As System.Windows.Forms.ComboBox
    Friend WithEvents rbNow As System.Windows.Forms.RadioButton
    Friend WithEvents dtpStartTime As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rbEveryMonth As System.Windows.Forms.RadioButton
    Friend WithEvents rbEveryWeek As System.Windows.Forms.RadioButton
    Friend WithEvents rbEveryDay As System.Windows.Forms.RadioButton
    Friend WithEvents rbOnce As System.Windows.Forms.RadioButton
    Friend WithEvents btnExport As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents fbdDestPath As System.Windows.Forms.FolderBrowserDialog
End Class
