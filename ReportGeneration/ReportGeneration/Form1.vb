Imports Microsoft.Office.Interop.Excel

Public Class PrincipalReports

    Private Sub btKPMSfileImport_Click(sender As Object, e As EventArgs) Handles btKPMSfileImport.Click
        '     Dim objOpenFileSelection As New FileSelection
        ' Dim strRetorno As String
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
        If VerificarPreenchimentoPathFile() Then

            If objExcelProcess Is Nothing Then objExcelProcess = New ExcelOperator

            Call objExcelProcess.openWorkBook(Me.objOpenFileDialogKPMS.FileName)

        End If
    End Sub

    Function VerificarPreenchimentoPathFile() As Double

        If Me.objOpenFileDialogKPMS.FileName <> "" Then Return True

    End Function



    Function IdentifyOperation()



    End Function

    Private Sub cmbReportType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbReportType.SelectedIndexChanged
        Select Case cmbReportType.GetItemText
            Case "Yearly"
            Case "Monthly"
            Case "Weekly"
            Case "Daily"
            Case Else
                MessageBox.Show("Item selecionado é invalido.")
        End Select
    End Sub
End Class
