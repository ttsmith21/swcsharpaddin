Option Explicit On
Public Class PageSTEPFileData
    Private Sub PageSTEPFileData_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        dgvSTEPFilesInfo.DataContext = oSTEPFiles
        Me.DataContext = oSTEPFiles
    End Sub
    Private Sub CheckBox_Checked(sender As System.Object, e As System.Windows.RoutedEventArgs)
        On Error Resume Next
        If Not oSTEPFiles Is Nothing Then
            oSTEPFiles.UpdateCheckStatus(True)
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Sub CheckBox_Unchecked(sender As System.Object, e As System.Windows.RoutedEventArgs)
        On Error Resume Next
        If Not oSTEPFiles Is Nothing Then
            oSTEPFiles.UpdateCheckStatus(False)
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub RefreshDGV()
        On Error Resume Next
        If Not dgvSTEPFilesInfo Is Nothing Then
            dgvSTEPFilesInfo.Items.Refresh()
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Sub CMDShowEdgesForCutLength_Click(sender As Object, e As RoutedEventArgs)
        On Error Resume Next
        Dim b As Button = DirectCast(sender, Button)
        If Not b Is Nothing Then
            Dim oPC As CStepFile = DirectCast(b.CommandParameter, CStepFile)
            If Not oPC Is Nothing Then
                oPC.SelectAllEdgesForCutLength()
            End If
            oPC = Nothing
        End If
        b = Nothing
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Sub CMDShowEdgesForMaterialLength_Click(sender As Object, e As RoutedEventArgs)
        On Error Resume Next
        'Dim b As Button = DirectCast(sender, Button)
        'If Not b Is Nothing Then
        '    Dim oPC As CStepFile = DirectCast(b.CommandParameter, CStepFile)
        '    If Not oPC Is Nothing Then
        '        oPC.SelectAllEdgesForCutLength()
        '    End If
        '    oPC = Nothing
        'End If
        'b = Nothing
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Sub CMDShowEdgesForHoles_Click(sender As Object, e As RoutedEventArgs)
        On Error Resume Next
        Dim b As Button = DirectCast(sender, Button)
        If Not b Is Nothing Then
            Dim oPC As CStepFile = DirectCast(b.CommandParameter, CStepFile)
            If Not oPC Is Nothing Then
                oPC.SelectAllEdgesForHoles()
            End If
            oPC = Nothing
        End If
        b = Nothing
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Sub CMDActivateDocument_Click(sender As Object, e As RoutedEventArgs)
        On Error Resume Next
        Dim b As Button = DirectCast(sender, Button)
        If Not b Is Nothing Then
            Dim oPC As CStepFile = DirectCast(b.CommandParameter, CStepFile)
            If Not oPC Is Nothing Then
                oPC.ActivateMe()
                oPC.ClearAllSelection()
            End If
            oPC = Nothing
        End If
        b = Nothing
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Sub CMDCalloutForMaterialLength_Click(sender As Object, e As RoutedEventArgs)
        On Error Resume Next
        Dim b As Button = DirectCast(sender, Button)
        If Not b Is Nothing Then
            Dim oPC As CStepFile = DirectCast(b.CommandParameter, CStepFile)
            If Not oPC Is Nothing Then
                oPC.ActivateMe()
                oPC.CreateCallout()
            End If
            oPC = Nothing
        End If
        b = Nothing
        If Err.Number <> 0 Then Err.Clear()
    End Sub
End Class
