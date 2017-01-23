Imports Microsoft.Office.Interop.Excel

Public Class ExcelOperator
    Const CauseType_Filter = "Model/Version(R)"
    Const ProblemType_Filter = "Product Categarization Version"
    Const ProductCategorization_Filter = "Product Categarization Version"
    Const Resolution_Filter = "Resolution"

    Dim objectEXL As Object
    Dim objExcelApp As Microsoft.Office.Interop.Excel.Application
    Dim objWorkBooks As Microsoft.Office.Interop.Excel.Workbooks
    Dim objWorkbook As Microsoft.Office.Interop.Excel.Workbook
    Dim objWorkSheet As Microsoft.Office.Interop.Excel.Worksheet

    Sub New()
        If objExcelApp Is Nothing Then objExcelApp = New Microsoft.Office.Interop.Excel.Application
        ' If objectEXL Is Nothing Then objectEXL = New Object("Excel")
        Me.objExcelApp.Visible = True
    End Sub

    Sub openWorkBook(ByVal strPathFile As String)
        Try
            ' Me.objWorkBook = Me.objExcelApp.Workbooks.Open(Session(strPathFile, , , , , , , , , True, , , , , ))
            Me.objWorkBooks = Me.objExcelApp.Workbooks
            Me.objWorkbook = Me.objWorkBooks.Open(strPathFile)
            Me.objWorkSheet = Me.objWorkbook.Sheets(1)

        Catch exc As Exception
            Console.WriteLine("ocorreu um erro " & exc.GetBaseException.ToString)
            Console.WriteLine("ocorreu um erro " & exc.Message)
        End Try
    End Sub

    Sub PrepareDataInWorkbook()


    End Sub

    Private Sub PrepareDataToReportAnalises(objFilter As Object, strFilter As String)

        Select Case strFilter
            Case "KPMS"
                Call PrepareDataKPMS(objFilter)
            Case "Remedy"
                Call PrepareDataRemedy(objFilter)
            Case "TimeLog"
                Call PrepareDataDaily(objFilter)
        End Select

    End Sub

    Private Sub PrepareDataKPMS(objFilter As Object)

        Me.objWorkbook.Activate()

    End Sub

    Private Sub PrepareDataRemedy(objFilter As Object)

    End Sub

    Private Sub PrepareDataDaily(objFilter As Object)

    End Sub

End Class
