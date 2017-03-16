Imports Microsoft.Office.Interop.Excel

Public Class ExcelOperator
    Const CauseType_Filter = "Model/Version(R)"
    Const ProblemType_Filter = "Product Categarization Version"
    Const ProductCategorization_Filter = "Product Categarization Version"
    Const Resolution_Filter = "Resolution"

    Public objectEXL As Object
    Public objExcelApp As Microsoft.Office.Interop.Excel.Application
    Public objWorkBooks As Microsoft.Office.Interop.Excel.Workbooks
    Public objWorkbook As Microsoft.Office.Interop.Excel.Workbook
    Public objWorkSheet As Microsoft.Office.Interop.Excel.Worksheet

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

    Sub PrepareDataToReportAnalises(strFilter As String)

        Select Case strFilter
            Case "KPMS"
                Call PrepareDataKPMS()
            Case "Remedy"
                Call PrepareDataRemedy()
            Case "TimeLog"
                Call PrepareDataDaily()
        End Select

    End Sub

    Private Sub PrepareDataKPMS()

        Dim objRange As Range
        Dim objSelection As Object

        objRange = objWorkSheet.Range("A8")
        objSelection = objRange.Select()

    End Sub

    Private Sub PrepareDataRemedy()

    End Sub

    Private Sub PrepareDataDaily()

    End Sub

End Class
