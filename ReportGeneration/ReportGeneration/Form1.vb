Imports Microsoft.Office.Interop.Excel

Public Class PrincipalReports

    Private Sub btKPMSfileImport_Click(sender As Object, e As EventArgs) Handles btKPMSfileImport.Click

        Dim returnDialog As DialogResult = Me.objOpenFileDialogKPMS.ShowDialog()

        Me.txtLOG = objLog.insertLineinLog(Me.txtLOG, "Inicio do Processamento")

        If returnDialog = DialogResult.OK Then

            Me.txtLOG = objLog.insertLineinLog(Me.txtLOG, "Abrindo Arquivo - " + CStr(Me.objOpenFileDialogKPMS.FileName))
            MessageBox.Show(Me.objOpenFileDialogKPMS.FileName)

        End If

    End Sub

    Private Sub cmbYear_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbYear.SelectedIndexChanged

    End Sub

    Private Sub btGenerate_Click(sender As Object, e As EventArgs) Handles btGenerate.Click
        If VerificarPreenchimentoPathFile(cmbReportFrom.Text) Then

            If objExcelProcess Is Nothing Then objExcelProcess = New ExcelOperator

            Select Case cmbReportFrom.Text
                Case "KPMS"
                    Call objExcelProcess.openWorkBook(Me.objOpenFileDialogKPMS.FileName)
                    Call objExcelProcess.PrepareDataToReportAnalises(cmbReportFrom.Text)
                Case "REMEDY"
                    Call objExcelProcess.openWorkBook(Me.objOpenFileDialogRemedy.FileName)
                    Call objExcelProcess.PrepareDataToReportAnalises(cmbReportFrom.Text)
            End Select

        End If
    End Sub

    Function VerificarPreenchimentoPathFile(strFilter As String) As Double

        Select Case strFilter
            Case "KPMS"
                If Me.objOpenFileDialogKPMS.FileName <> "" Then Return True
            Case "Remedy"
                If Me.objOpenFileDialogRemedy.FileName <> "" Then Return True
        End Select


    End Function

    Private Sub cmbReportType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbReportType.SelectedIndexChanged
        Stop
        Select Case cmbReportType.Text
            Case "Annualy"
                Me.cmbMonth.Enabled = False
                Me.cmbWeek.Enabled = False
                Me.DT_DailyReport.Enabled = False
            Case "Monthly"
                Me.cmbMonth.Enabled = True
                Me.cmbWeek.Enabled = False
                Me.DT_DailyReport.Enabled = False
            Case "Weekly"
                Me.cmbMonth.Enabled = True
                Me.cmbWeek.Enabled = True
                Me.DT_DailyReport.Enabled = False
            Case "Daily"
                Me.cmbMonth.Enabled = True
                Me.cmbWeek.Enabled = True
                Me.DT_DailyReport.Enabled = True
            Case Else
                MessageBox.Show("Item selecionado é invalido.")
        End Select
    End Sub

    Private Sub cmbReportFrom_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbReportFrom.SelectedIndexChanged
        Select Case UCase(cmbReportFrom.Text)
            Case "KPMS"
                If Not VerificarPreenchimentoPathFile(cmbReportFrom.Text) Then MessageBox.Show("The file path of KPMS report is Manadatory, please fill it on the button KPMS")
            Case "REMEDY"
                If Not VerificarPreenchimentoPathFile(cmbReportFrom.Text) Then MessageBox.Show("The file path of Remedy report is Manadatory, please fill it on the button Remedy")
            Case "MERGE"
                MessageBox.Show("Opção ainda inativa")
            Case "BOTH"
                If Not VerificarPreenchimentoPathFile("KMPS") Or Not VerificarPreenchimentoPathFile("Remedy") Then MessageBox.Show("The file path of Remedy report is Manadatory, please fill it on the button Remedy")
        End Select
    End Sub
End Class
