Public Class WindowHelp
    Private Sub CmdClose_Click(sender As Object, e As RoutedEventArgs)
        On Error Resume Next
        Me.Close()
        If Err.Number <> 0 Then Err.Clear()
    End Sub
End Class
