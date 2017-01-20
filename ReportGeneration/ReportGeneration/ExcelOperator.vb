Imports Microsoft.Office.Interop.Excel

Public Class ExcelOperator

    Dim objectEXL As Object
    Dim objExcelApp As Microsoft.Office.Interop.Excel.Application
    Dim objWorkBooks As Microsoft.Office.Interop.Excel.Workbooks
    Dim objWorkbook As Microsoft.Office.Interop.Excel.Workbook

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
        Catch exc As Exception
            Console.WriteLine("ocorreu um erro " & exc.GetBaseException.ToString)
            Console.WriteLine("ocorreu um erro " & exc.Message)
        End Try
    End Sub

    Sub PrepareDataInWorkbook()


    End Sub

End Class
