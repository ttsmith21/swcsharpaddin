Public Class WindowAbout
    Public Sub New()
        On Error Resume Next
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        txtVersion.Text = GetProductVersion()   'System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString()
        txtProductName.Text = strProductName
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Sub CmdClose_Click(sender As Object, e As RoutedEventArgs)
        On Error Resume Next
        Me.Close()
        If Err.Number <> 0 Then Err.Clear()
    End Sub
End Class
