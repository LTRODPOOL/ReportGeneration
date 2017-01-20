Public Class LogObject
    Dim PreviousLog(0) As String
    Sub Main()
        Console.WriteLine("Iniciando Processamento de LOG")
    End Sub

    Public Function insertLineinLog(ByVal objFieldLog As Windows.Forms.TextBox, ByVal txtLog As String) As Windows.Forms.TextBox
        Try

            If Me.PreviousLog(0) <> "" Then ReDim Preserve Me.PreviousLog(UBound(Me.PreviousLog) + 1)

            Me.PreviousLog(UBound(Me.PreviousLog)) = Format(Now(), "dd/mm/yyyy hh:mm:ss") & " - " & txtLog

            'objFieldLog.Lines = Me.PreviousLog
            objFieldLog.Text = Me.PreviousLog(0)
            objFieldLog.Lines = Me.PreviousLog
            insertLineinLog = objFieldLog
            Return objFieldLog

        Catch exc As Exception
            Console.WriteLine("ocorreu um erro " & exc.GetBaseException.ToString)
            Console.WriteLine("ocorreu um erro " & exc.Message)
        End Try
        Return objFieldLog
    End Function
End Class
