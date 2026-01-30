Option Explicit On
Imports SolidWorks.Interop
Public Class UCUI
    Implements IUCUI
    Private oPageDGV As PageSTEPFileData
    Public Sub New()
        On Error Resume Next
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        oPageDGV = New PageSTEPFileData
        oSTEPFiles = New CStepFileCollection
        pbProcessStatus.DataContext = CProgress.GetInstance()
        lblProcessStatus.DataContext = CProgress.GetInstance()
        CmdClearData.DataContext = oSTEPFiles
        CmdProcessData.DataContext = oSTEPFiles
        CmdSaveData.DataContext = oSTEPFiles
        ToggleShadedWithEdges.DataContext = oSTEPFiles
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub SetSWAddinObject(ByRef oSWApp As sldworks.SldWorks, ByRef sProductVersion As String)
        On Error Resume Next
        If Not oSWApp Is Nothing Then
            SWApp = oSWApp
            strProductVersion = sProductVersion
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Sub UCUI_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        On Error Resume Next
        frameMainUI.Navigate(oPageDGV)
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub CmdSelectSTEPFileFolder_Click(sender As Object, e As RoutedEventArgs)
        On Error Resume Next
        Dim swModel As sldworks.ModelDoc2 = SWApp.ActiveDoc
        If Not swModel Is Nothing Then
            My.Settings.profileFolder = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\"))
            My.Settings.Save()
            RefreshDGV(swModel)
            Call CmdProcessData_Click(sender, e)
            CmdSaveData_Click(sender, e)
        Else
            Dim oExcelApplicationObject As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
            Dim iExcelProcessID As IntPtr = 0
            If Not oExcelApplicationObject Is Nothing Then
                GetWindowThreadProcessId(oExcelApplicationObject.Hwnd, iExcelProcessID)
                oExcelApplicationObject.Visible = False
                oExcelApplicationObject.UserControl = False
                oExcelApplicationObject.DisplayAlerts = False

                Dim oFileDialog As Microsoft.Office.Core.FileDialog = oExcelApplicationObject.FileDialog(Microsoft.Office.Core.MsoFileDialogType.msoFileDialogFolderPicker)



                If Not oFileDialog Is Nothing Then
                    oFileDialog.AllowMultiSelect = False
                    oFileDialog.Title = "Select input folder for STEP files"
                    If Not String.IsNullOrEmpty(My.Settings.profileFolder) Then
                        oFileDialog.InitialFileName = My.Settings.profileFolder
                    End If

                    If oFileDialog.Show() = -1 Then
                        My.Settings.profileFolder = oFileDialog.SelectedItems.Item(1).ToString()
                    End If
                    My.Settings.Save()
                    RefreshDGV(swModel)

                End If
            End If

            'Close Excel and kill process..
            If iExcelProcessID = 0 Then
                GetWindowThreadProcessId(oExcelApplicationObject.Hwnd, iExcelProcessID)
            End If
            oExcelApplicationObject.Workbooks.Close()
            oExcelApplicationObject.DisplayAlerts = True
            oExcelApplicationObject.Quit()
            oExcelApplicationObject = Nothing

            Dim oExcelProcess As Process = Process.GetProcessById(iExcelProcessID.ToInt32)
            oExcelProcess.Kill()
            oExcelProcess = Nothing
            oExcelApplicationObject = Nothing
            iExcelProcessID = 0

        End If

        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Sub CmdRefreshData_Click(sender As Object, e As RoutedEventArgs)
        Dim swmodel As sldworks.ModelDoc2 = SWApp.ActiveDoc
        RefreshDGV(swmodel)
    End Sub
    Private Sub CmdClearData_Click(sender As Object, e As RoutedEventArgs)
        On Error Resume Next
        If Not oSTEPFiles Is Nothing Then
            oSTEPFiles.Clear()
        End If
        System.Windows.Forms.Application.DoEvents()
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Sub RefreshDGV(swModel As sldworks.ModelDoc2)
        On Error Resume Next
        If Not String.IsNullOrEmpty(My.Settings.profileFolder) Then
            If IO.Directory.Exists(My.Settings.profileFolder) Then
                If Not oSTEPFiles Is Nothing Then
                    oSTEPFiles.Clear()
                    Dim oFiles() As String = IO.Directory.GetFiles(My.Settings.profileFolder)
                    If Not swModel Is Nothing Then
                        oFiles(0) = swModel.GetPathName
                    End If

                    If Not oFiles Is Nothing Then
                        CProgress.GetInstance.Minimum = 0
                        CProgress.GetInstance.Maximum = UBound(oFiles)
                        CProgress.GetInstance.ProgressText = "Populating Files Data..."
                        CProgress.GetInstance.InitialiseProgressLevel()
                        CProgress.GetInstance.ProgressBarVisibility = Visibility.Visible
                        System.Windows.Forms.Application.DoEvents()
                        Dim iFiles As Integer = 0
                        If Not swModel Is Nothing Then
                            oSTEPFiles.Add(New CStepFile(oFiles(0)))
                        Else
                            For iFiles = LBound(oFiles) To UBound(oFiles)
                                If UCase(Trim(IO.Path.GetExtension(oFiles(iFiles)))) = ".STEP" Or UCase(Trim(IO.Path.GetExtension(oFiles(iFiles)))) = ".iges" Or UCase(Trim(IO.Path.GetExtension(oFiles(iFiles)))) = ".sldprt" Then
                                    oSTEPFiles.Add(New CStepFile(oFiles(iFiles)))
                                End If
                                CProgress.GetInstance.ProgressText = "Populating Files Data (" & iFiles + 1 & "/" & CProgress.GetInstance().Maximum & ")..."
                                CProgress.GetInstance.UpdateValue()
                            Next iFiles
                        End If
                        CProgress.GetInstance.Minimum = 0
                        CProgress.GetInstance.Maximum = 100
                        CProgress.GetInstance.ProgressText = ""
                        CProgress.GetInstance.InitialiseProgressLevel()
                        CProgress.GetInstance.ProgressBarVisibility = Visibility.Collapsed
                        System.Windows.Forms.Application.DoEvents()

                        If Not oPageDGV Is Nothing Then
                            oPageDGV.RefreshDGV()
                        End If
                    End If
                End If
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Sub CmdHelp_Click(sender As Object, e As RoutedEventArgs)
        On Error Resume Next
        Dim oWindow As New WindowHelp
        oWindow.ShowDialog()
        oWindow = Nothing
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Sub CmdAbout_Click(sender As Object, e As RoutedEventArgs)
        On Error Resume Next
        Dim oWindow As New WindowAbout
        oWindow.ShowDialog()
        oWindow = Nothing
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Sub CmdProcessData_Click(sender As Object, e As RoutedEventArgs)
        On Error Resume Next
        If Not oSTEPFiles Is Nothing Then
            oSTEPFiles.ExtractData()
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Sub CmdSaveData_Click(sender As Object, e As RoutedEventArgs)
        On Error Resume Next
        If Not oSTEPFiles Is Nothing Then
            oSTEPFiles.SaveDataAndFile()
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Sub ToggleShadedWithEdges_Click(sender As Object, e As RoutedEventArgs)
        On Error Resume Next
        If Not oSTEPFiles Is Nothing Then
            oSTEPFiles.UpdateView()
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Property ViewModeAsShadedWithEdges As Boolean
        Get
            Return My.Settings.ShadedWithEdges
        End Get
        Set(value As Boolean)
            On Error Resume Next
            If Not oSTEPFiles Is Nothing Then
                oSTEPFiles.IsShadedWithEdges = value
            End If
            If Err.Number <> 0 Then Err.Clear()
        End Set
    End Property
    Public ReadOnly Property IsDocumentLoading As Boolean
        Get
            Return My.Settings.IsDocumentLoading
        End Get
    End Property
    Public Sub RunForAutoPilot(swModel As sldworks.ModelDoc2) Implements IUCUI.RunForAutoPilot '(sender As Object, e As RoutedEventArgs) 
        Dim sender As Object = Nothing
        Dim e As RoutedEventArgs = Nothing
        If Not swModel Is Nothing Then
            My.Settings.profileFolder = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\"))
            My.Settings.Save()
            RefreshDGV(swModel)
            Call CmdProcessData_Click(sender, e)
            CmdSaveData_Click(sender, e)
        Else
            On Error Resume Next
            Dim oExcelApplicationObject As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
            Dim iExcelProcessID As IntPtr = 0
            If Not oExcelApplicationObject Is Nothing Then
                GetWindowThreadProcessId(oExcelApplicationObject.Hwnd, iExcelProcessID)
                oExcelApplicationObject.Visible = False
                oExcelApplicationObject.UserControl = False
                oExcelApplicationObject.DisplayAlerts = False

                Dim oFileDialog As Microsoft.Office.Core.FileDialog = oExcelApplicationObject.FileDialog(Microsoft.Office.Core.MsoFileDialogType.msoFileDialogFolderPicker)



                If Not oFileDialog Is Nothing Then
                    oFileDialog.AllowMultiSelect = False
                    oFileDialog.Title = "Select input folder for STEP files"
                    If Not String.IsNullOrEmpty(My.Settings.profileFolder) Then
                        oFileDialog.InitialFileName = My.Settings.profileFolder
                    End If

                    If oFileDialog.Show() = -1 Then
                        My.Settings.profileFolder = oFileDialog.SelectedItems.Item(1).ToString()
                    End If
                    My.Settings.Save()
                    RefreshDGV(swModel)

                End If
            End If

            'Close Excel and kill process..
            If iExcelProcessID = 0 Then
                GetWindowThreadProcessId(oExcelApplicationObject.Hwnd, iExcelProcessID)
            End If
            oExcelApplicationObject.Workbooks.Close()
            oExcelApplicationObject.DisplayAlerts = True
            oExcelApplicationObject.Quit()
            oExcelApplicationObject = Nothing


            Dim oExcelProcess As Process = Process.GetProcessById(iExcelProcessID.ToInt32)
            oExcelProcess.Kill()
            oExcelProcess = Nothing
            oExcelApplicationObject = Nothing
            iExcelProcessID = 0

        End If

        If Err.Number <> 0 Then Err.Clear()
    End Sub
End Class
