'<CompilerServices.DesignerGenerated()>

Partial Class PrincipalReports
    Inherits System.Windows.Forms.Form

    Private objLog As New LogObject
    Private objFolderBrowserDialogKPMS As FolderBrowserDialog
    Private objOpenFileDialogKPMS As OpenFileDialog
    Private openFileNameKPMS As String, folderNameKPMS As String

    Private objFolderBrowserDialogRemedy As FolderBrowserDialog
    Private objOpenFileDialogRemedy As OpenFileDialog
    Private openFileNameRemedy As String, folderNameRemedy As String

    Private objExcelProcess As ExcelOperator

    Private fileOpened As Boolean = False

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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btRemedyfileimport = New System.Windows.Forms.Button()
        Me.btKPMSfileImport = New System.Windows.Forms.Button()
        Me.lblRemedyReport = New System.Windows.Forms.Label()
        Me.lblSelectionKPMS = New System.Windows.Forms.Label()
        Me.Actions = New System.Windows.Forms.GroupBox()
        Me.cmbReportFrom = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cmbYear = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btGenerate = New System.Windows.Forms.Button()
        Me.cmbWeek = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmbMonth = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbReportType = New System.Windows.Forms.ComboBox()
        Me.gpLOG = New System.Windows.Forms.GroupBox()
        Me.txtLOG = New System.Windows.Forms.TextBox()
        Me.objOpenFileDialogKPMS = New System.Windows.Forms.OpenFileDialog()
        Me.objFolderBrowserDialogKPMS = New System.Windows.Forms.FolderBrowserDialog()
        Me.objOpenFileDialogRemedy = New System.Windows.Forms.OpenFileDialog()
        Me.objFolderBrowserDialogRemedy = New System.Windows.Forms.FolderBrowserDialog()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.DT_DailyReport = New System.Windows.Forms.DateTimePicker()
        Me.GroupBox1.SuspendLayout()
        Me.Actions.SuspendLayout()
        Me.gpLOG.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btRemedyfileimport)
        Me.GroupBox1.Controls.Add(Me.btKPMSfileImport)
        Me.GroupBox1.Controls.Add(Me.lblRemedyReport)
        Me.GroupBox1.Controls.Add(Me.lblSelectionKPMS)
        Me.GroupBox1.Location = New System.Drawing.Point(17, 16)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Size = New System.Drawing.Size(611, 85)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Report Selection"
        '
        'btRemedyfileimport
        '
        Me.btRemedyfileimport.Location = New System.Drawing.Point(413, 51)
        Me.btRemedyfileimport.Margin = New System.Windows.Forms.Padding(2)
        Me.btRemedyfileimport.Name = "btRemedyfileimport"
        Me.btRemedyfileimport.Size = New System.Drawing.Size(181, 23)
        Me.btRemedyfileimport.TabIndex = 3
        Me.btRemedyfileimport.Text = "Remedy"
        Me.btRemedyfileimport.UseVisualStyleBackColor = True
        '
        'btKPMSfileImport
        '
        Me.btKPMSfileImport.Location = New System.Drawing.Point(413, 9)
        Me.btKPMSfileImport.Margin = New System.Windows.Forms.Padding(2)
        Me.btKPMSfileImport.Name = "btKPMSfileImport"
        Me.btKPMSfileImport.Size = New System.Drawing.Size(181, 25)
        Me.btKPMSfileImport.TabIndex = 2
        Me.btKPMSfileImport.Text = "KPMS"
        Me.btKPMSfileImport.UseVisualStyleBackColor = True
        '
        'lblRemedyReport
        '
        Me.lblRemedyReport.AutoSize = True
        Me.lblRemedyReport.Location = New System.Drawing.Point(4, 51)
        Me.lblRemedyReport.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRemedyReport.Name = "lblRemedyReport"
        Me.lblRemedyReport.Size = New System.Drawing.Size(188, 13)
        Me.lblRemedyReport.TabIndex = 1
        Me.lblRemedyReport.Text = "Select the Remedy Report in Excel file"
        '
        'lblSelectionKPMS
        '
        Me.lblSelectionKPMS.AutoSize = True
        Me.lblSelectionKPMS.Location = New System.Drawing.Point(4, 21)
        Me.lblSelectionKPMS.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblSelectionKPMS.Name = "lblSelectionKPMS"
        Me.lblSelectionKPMS.Size = New System.Drawing.Size(179, 13)
        Me.lblSelectionKPMS.TabIndex = 0
        Me.lblSelectionKPMS.Text = "Select the KPMS Report in Excel file"
        '
        'Actions
        '
        Me.Actions.Controls.Add(Me.DT_DailyReport)
        Me.Actions.Controls.Add(Me.Label6)
        Me.Actions.Controls.Add(Me.cmbReportFrom)
        Me.Actions.Controls.Add(Me.Label5)
        Me.Actions.Controls.Add(Me.cmbYear)
        Me.Actions.Controls.Add(Me.Label4)
        Me.Actions.Controls.Add(Me.btGenerate)
        Me.Actions.Controls.Add(Me.cmbWeek)
        Me.Actions.Controls.Add(Me.Label3)
        Me.Actions.Controls.Add(Me.cmbMonth)
        Me.Actions.Controls.Add(Me.Label2)
        Me.Actions.Controls.Add(Me.Label1)
        Me.Actions.Controls.Add(Me.cmbReportType)
        Me.Actions.Controls.Add(Me.gpLOG)
        Me.Actions.Location = New System.Drawing.Point(17, 116)
        Me.Actions.Margin = New System.Windows.Forms.Padding(2)
        Me.Actions.Name = "Actions"
        Me.Actions.Padding = New System.Windows.Forms.Padding(2)
        Me.Actions.Size = New System.Drawing.Size(611, 283)
        Me.Actions.TabIndex = 1
        Me.Actions.TabStop = False
        Me.Actions.Text = "Actions"
        '
        'cmbReportFrom
        '
        Me.cmbReportFrom.FormattingEnabled = True
        Me.cmbReportFrom.Items.AddRange(New Object() {"KPMS", "REMEDY", "BOTH", "MERGE"})
        Me.cmbReportFrom.Location = New System.Drawing.Point(4, 32)
        Me.cmbReportFrom.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbReportFrom.Name = "cmbReportFrom"
        Me.cmbReportFrom.Size = New System.Drawing.Size(200, 21)
        Me.cmbReportFrom.TabIndex = 11
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(7, 16)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 13)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Report From"
        '
        'cmbYear
        '
        Me.cmbYear.FormattingEnabled = True
        Me.cmbYear.Location = New System.Drawing.Point(4, 107)
        Me.cmbYear.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbYear.Name = "cmbYear"
        Me.cmbYear.Size = New System.Drawing.Size(200, 21)
        Me.cmbYear.TabIndex = 9
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(5, 92)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(29, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Year"
        '
        'btGenerate
        '
        Me.btGenerate.Location = New System.Drawing.Point(93, 256)
        Me.btGenerate.Margin = New System.Windows.Forms.Padding(2)
        Me.btGenerate.Name = "btGenerate"
        Me.btGenerate.Size = New System.Drawing.Size(111, 23)
        Me.btGenerate.TabIndex = 7
        Me.btGenerate.Text = "Generate"
        Me.btGenerate.UseVisualStyleBackColor = True
        '
        'cmbWeek
        '
        Me.cmbWeek.FormattingEnabled = True
        Me.cmbWeek.Items.AddRange(New Object() {"WK1", "WK2", "WK3", "WK4", "WK5"})
        Me.cmbWeek.Location = New System.Drawing.Point(4, 186)
        Me.cmbWeek.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbWeek.Name = "cmbWeek"
        Me.cmbWeek.Size = New System.Drawing.Size(200, 21)
        Me.cmbWeek.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 171)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(36, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Week"
        '
        'cmbMonth
        '
        Me.cmbMonth.FormattingEnabled = True
        Me.cmbMonth.Items.AddRange(New Object() {"1 - January", "2 - February", "3 - March", "4 - April", "5 - May", "6 - June", "7 - July", "8 - August", "9 - Sepetember", "10 - October", "11 - November", "12 - December"})
        Me.cmbMonth.Location = New System.Drawing.Point(4, 146)
        Me.cmbMonth.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbMonth.Name = "cmbMonth"
        Me.cmbMonth.Size = New System.Drawing.Size(200, 21)
        Me.cmbMonth.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 131)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(37, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Month"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 52)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Type of Report"
        '
        'cmbReportType
        '
        Me.cmbReportType.FormattingEnabled = True
        Me.cmbReportType.Items.AddRange(New Object() {"Daily", "Weekly", "Monthly", "Annualy"})
        Me.cmbReportType.Location = New System.Drawing.Point(4, 67)
        Me.cmbReportType.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbReportType.Name = "cmbReportType"
        Me.cmbReportType.Size = New System.Drawing.Size(200, 21)
        Me.cmbReportType.TabIndex = 1
        '
        'gpLOG
        '
        Me.gpLOG.Controls.Add(Me.txtLOG)
        Me.gpLOG.Location = New System.Drawing.Point(226, 16)
        Me.gpLOG.Margin = New System.Windows.Forms.Padding(2)
        Me.gpLOG.Name = "gpLOG"
        Me.gpLOG.Padding = New System.Windows.Forms.Padding(2)
        Me.gpLOG.Size = New System.Drawing.Size(381, 263)
        Me.gpLOG.TabIndex = 0
        Me.gpLOG.TabStop = False
        Me.gpLOG.Text = "LOG"
        '
        'txtLOG
        '
        Me.txtLOG.Enabled = False
        Me.txtLOG.Location = New System.Drawing.Point(9, 17)
        Me.txtLOG.Margin = New System.Windows.Forms.Padding(2)
        Me.txtLOG.Multiline = True
        Me.txtLOG.Name = "txtLOG"
        Me.txtLOG.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtLOG.Size = New System.Drawing.Size(359, 234)
        Me.txtLOG.TabIndex = 0
        '
        'objOpenFileDialogKPMS
        '
        Me.objOpenFileDialogKPMS.DefaultExt = "xlsx"
        Me.objOpenFileDialogKPMS.Filter = "xlsx file (*.xlsx) | *.xlsx"
        '
        'objFolderBrowserDialogKPMS
        '
        Me.objFolderBrowserDialogKPMS.Description = "Select the directory that you want to use As the default."
        Me.objFolderBrowserDialogKPMS.RootFolder = System.Environment.SpecialFolder.MyDocuments
        Me.objFolderBrowserDialogKPMS.ShowNewFolderButton = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(1, 209)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(30, 13)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "Date"
        '
        'DT_DailyReport
        '
        Me.DT_DailyReport.Location = New System.Drawing.Point(4, 225)
        Me.DT_DailyReport.Name = "DT_DailyReport"
        Me.DT_DailyReport.Size = New System.Drawing.Size(200, 20)
        Me.DT_DailyReport.TabIndex = 2
        '
        'PrincipalReports
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(636, 410)
        Me.Controls.Add(Me.Actions)
        Me.Controls.Add(Me.GroupBox1)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "PrincipalReports"
        Me.Text = "Reports Generation"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Actions.ResumeLayout(False)
        Me.Actions.PerformLayout()
        Me.gpLOG.ResumeLayout(False)
        Me.gpLOG.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents lblSelectionKPMS As Label
    Friend WithEvents btRemedyfileimport As Button
    Friend WithEvents btKPMSfileImport As Button
    Friend WithEvents lblRemedyReport As Label
    Friend WithEvents Actions As GroupBox
    Friend WithEvents gpLOG As GroupBox
    Friend WithEvents cmbWeek As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents cmbMonth As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents cmbReportType As ComboBox
    Friend WithEvents btGenerate As Button
    Friend WithEvents txtLOG As TextBox

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        InitializeYearField()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub InitializeYearField()
        Dim countYears As Integer
        Dim NumYears As Integer
        Dim YearsSelection() As String

        NumYears = Today.Year

        ReDim YearsSelection(0)
        For countYears = 0 To 8
            If YearsSelection(0) <> "" Then
                ReDim Preserve YearsSelection(UBound(YearsSelection) + 1)
            End If
            YearsSelection(UBound(YearsSelection)) = NumYears - countYears
        Next

        Me.cmbYear.Items.AddRange(YearsSelection)
    End Sub

    Friend WithEvents cmbYear As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents cmbReportFrom As ComboBox
    Friend WithEvents Label5 As Label
    Friend WithEvents DT_DailyReport As DateTimePicker
    Friend WithEvents Label6 As Label
End Class
