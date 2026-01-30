Option Explicit On

Imports System.Reflection
Imports System.Runtime.InteropServices

Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports SolidWorksTools
Imports SolidWorksTools.File

<Guid("C8A18D5B-73A8-4007-A489-C9670FC41443")>
<ComVisible(True)>
<SwAddin(
        Description:="Tool to extract data from selected profiles",
        Title:="Extract Profile Data",
        LoadAtStartup:=True
        )>
Public Class SwAddin
    Implements SolidWorks.Interop.swpublished.SwAddin

#Region "Local Variables"
    Dim WithEvents SWApp As SldWorks
    Dim addinID As Integer
    Dim openDocs As Hashtable
    Dim SwEventPtr As SldWorks
    Private oTaskPaneView As SolidWorks.Interop.sldworks.TaskpaneView
    Private otaskPaneObj As UCUIAddin
    Private strProductNameForUITop As String = strProductName & " (" & GetProductVersion() & ")"
    ' Public Properties
    ReadOnly Property SWApplication() As SldWorks
        Get
            Return SWApp
        End Get
    End Property
    ReadOnly Property OpenDocumentsTable() As Hashtable
        Get
            Return openDocs
        End Get
    End Property
    Public ReadOnly Property UCUIAddinObject As UCUIAddin
        Get
            Return otaskPaneObj
        End Get
    End Property
#End Region
#Region "SolidWorks Registration"
    <ComRegisterFunction()> Public Shared Sub RegisterFunction(ByVal t As Type)
        ' Get Custom Attribute: SwAddinAttribute
        Dim attributes() As Object
        Dim SWattr As SwAddinAttribute = Nothing
        attributes = System.Attribute.GetCustomAttributes(GetType(SwAddin), GetType(SwAddinAttribute))
        If attributes.Length > 0 Then
            SWattr = DirectCast(attributes(0), SwAddinAttribute)
        End If
        Try
            Dim hklm As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine
            Dim hkcu As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser
            Dim keyname As String = "SOFTWARE\SolidWorks\Addins\{" + t.GUID.ToString() + "}"
            Dim addinkey As Microsoft.Win32.RegistryKey = hklm.CreateSubKey(keyname)
            addinkey.SetValue(Nothing, 0)
            addinkey.SetValue("Description", SWattr.Description)
            addinkey.SetValue("Title", SWattr.Title)

            'Extract icon during registration 
            Dim thisAssembly As Assembly = System.Reflection.Assembly.GetExecutingAssembly()
            Dim iBmp As New BitmapHandler
            Dim strAddinIconPath As String = iBmp.CreateFileFromResourceBitmap("SWExtractDataAddin.AddinIcon.bmp", thisAssembly)
            'Copy the bitmap to a suitable permanent location with a meaningful filename 
            Dim addInPath As String = System.IO.Path.GetDirectoryName(thisAssembly.Location)
            Dim iconPath As String = System.IO.Path.Combine(addInPath, "AddinIcon.bmp")
            System.IO.File.Copy(strAddinIconPath, iconPath, True)
            'Register the icon location 
            addinkey.SetValue("Icon Path", iconPath)
            iBmp = Nothing
            thisAssembly = Nothing

            keyname = "Software\SolidWorks\AddInsStartup\{" + t.GUID.ToString() + "}"
            addinkey = hkcu.CreateSubKey(keyname)
            addinkey.SetValue(Nothing, SWattr.LoadAtStartup, Microsoft.Win32.RegistryValueKind.DWord)
        Catch nl As System.NullReferenceException
            Console.WriteLine("There was a problem registering this dll: SWattr is null.\n " & nl.Message)
            System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n" & nl.Message)
        Catch e As System.Exception
            Console.WriteLine("There was a problem registering this dll: " & e.Message)
            System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: " & e.Message)
        End Try
    End Sub
    <ComUnregisterFunction()> Public Shared Sub UnregisterFunction(ByVal t As Type)
        Try
            Dim hklm As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine
            Dim hkcu As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser
            Dim keyname As String = "SOFTWARE\SolidWorks\Addins\{" + t.GUID.ToString() + "}"
            hklm.DeleteSubKey(keyname)
            keyname = "Software\SolidWorks\AddInsStartup\{" + t.GUID.ToString() + "}"
            hkcu.DeleteSubKey(keyname)
        Catch nl As System.NullReferenceException
            Console.WriteLine("There was a problem unregistering this dll: SWattr is null.\n " & nl.Message)
            System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: SWattr is null.\n" & nl.Message)
        Catch e As System.Exception
            Console.WriteLine("There was a problem unregistering this dll: " & e.Message)
            System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: " & e.Message)
        End Try
    End Sub
#End Region
#Region "ISwAddin Implementation"
    Function ConnectToSW(ByVal ThisSW As Object, ByVal Cookie As Integer) As Boolean Implements SolidWorks.Interop.swpublished.SwAddin.ConnectToSW
        On Error Resume Next
        SWApp = ThisSW
        addinID = Cookie

        ' Setup callbacks
        SWApp.SetAddinCallbackInfo(0, Me, addinID)

        'Setup the Event Handlers
        SwEventPtr = SWApp
        openDocs = New Hashtable
        AttachEventHandlers()

        If Not oTaskPaneView Is Nothing Then
            oTaskPaneView.DeleteView()
            oTaskPaneView = Nothing
        End If
        Dim thisAssembly As Assembly = System.Reflection.Assembly.GetAssembly(Me.GetType())
        Dim iBmp As New BitmapHandler
        oTaskPaneView = SWApp.CreateTaskpaneView2(iBmp.CreateFileFromResourceBitmap("SWExtractDataAddin.ToolbarLarge.bmp", thisAssembly), strProductNameForUITop)
        iBmp = Nothing
        If Not oTaskPaneView Is Nothing Then
            If oTaskPaneView.GetControl() Is Nothing Then
                otaskPaneObj = oTaskPaneView.AddControl("SWExtractDataAddin.UCUIAddin", "")
            Else
                otaskPaneObj = oTaskPaneView.GetControl()
            End If
            If Not otaskPaneObj Is Nothing Then
                otaskPaneObj.SetSWAddinObject(SWApplication)
            End If
            oTaskPaneView.ShowView()
        End If
        ConnectToSW = True
        If Err.Number <> 0 Then Err.Clear()
    End Function
    Function DisconnectFromSW() As Boolean Implements SolidWorks.Interop.swpublished.SwAddin.DisconnectFromSW
        On Error Resume Next
        DetachEventHandlers()
        If Not oTaskPaneView Is Nothing Then
            oTaskPaneView.DeleteView()
            oTaskPaneView = Nothing
        End If
        If Not otaskPaneObj Is Nothing Then
            otaskPaneObj = Nothing
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(SWApp)
        SWApp = Nothing
        'The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.WaitForPendingFinalizers()
        DisconnectFromSW = True
        If Err.Number <> 0 Then Err.Clear()
    End Function
#End Region
#Region "Event Methods"
    Sub AttachEventHandlers()
        AttachSWEvents()
        'Listen for events on all currently open docs
        AttachEventsToAllDocuments()
    End Sub
    Sub DetachEventHandlers()
        DetachSWEvents()
        'Close events on all currently open docs
        Dim docHandler As DocumentEventHandler
        Dim key As ModelDoc2
        Dim numKeys As Integer
        numKeys = openDocs.Count
        If numKeys > 0 Then
            Dim keys() As Object = New Object(numKeys - 1) {}
            'Remove all document event handlers
            openDocs.Keys.CopyTo(keys, 0)
            For Each key In keys
                docHandler = openDocs.Item(key)
                docHandler.DetachEventHandlers() 'This also removes the pair from the hash
                docHandler = Nothing
                key = Nothing
            Next
        End If
    End Sub
    Sub AttachSWEvents()
        Try
            AddHandler SWApp.ActiveDocChangeNotify, AddressOf Me.SldWorks_ActiveDocChangeNotify
            AddHandler SWApp.DocumentLoadNotify2, AddressOf Me.SldWorks_DocumentLoadNotify2
            AddHandler SWApp.FileNewNotify2, AddressOf Me.SldWorks_FileNewNotify2
            AddHandler SWApp.ActiveModelDocChangeNotify, AddressOf Me.SldWorks_ActiveModelDocChangeNotify
            AddHandler SWApp.FileOpenPostNotify, AddressOf Me.SldWorks_FileOpenPostNotify
        Catch e As Exception
            Console.WriteLine(e.Message)
        End Try
    End Sub
    Sub DetachSWEvents()
        Try
            RemoveHandler SWApp.ActiveDocChangeNotify, AddressOf Me.SldWorks_ActiveDocChangeNotify
            RemoveHandler SWApp.DocumentLoadNotify2, AddressOf Me.SldWorks_DocumentLoadNotify2
            RemoveHandler SWApp.FileNewNotify2, AddressOf Me.SldWorks_FileNewNotify2
            RemoveHandler SWApp.ActiveModelDocChangeNotify, AddressOf Me.SldWorks_ActiveModelDocChangeNotify
            RemoveHandler SWApp.FileOpenPostNotify, AddressOf Me.SldWorks_FileOpenPostNotify
        Catch e As Exception
            Console.WriteLine(e.Message)
        End Try
    End Sub
    Sub AttachEventsToAllDocuments()
        Dim modDoc As ModelDoc2
        modDoc = SWApp.GetFirstDocument()
        While Not modDoc Is Nothing
            If Not openDocs.Contains(modDoc) Then
                AttachModelDocEventHandler(modDoc)
            Else
                Dim docHandler As DocumentEventHandler = openDocs(modDoc)
                If Not docHandler Is Nothing Then
                    docHandler.ConnectModelViews()
                End If
            End If
            modDoc = modDoc.GetNext()
        End While
    End Sub
    Function AttachModelDocEventHandler(ByVal modDoc As ModelDoc2) As Boolean
        On Error Resume Next
        If modDoc Is Nothing Then
            Return False
        End If
        Dim docHandler As DocumentEventHandler = Nothing
        If Not openDocs.Contains(modDoc) Then
            Select Case modDoc.GetType
                Case swDocumentTypes_e.swDocPART
                    docHandler = New PartEventHandler()
                Case swDocumentTypes_e.swDocASSEMBLY
                    docHandler = New AssemblyEventHandler()
                Case swDocumentTypes_e.swDocDRAWING
                    docHandler = New DrawingEventHandler()
            End Select
            docHandler.Init(SWApp, Me, modDoc)
            docHandler.AttachEventHandlers()
            openDocs.Add(modDoc, docHandler)
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return True
    End Function
    Sub DetachModelEventHandler(ByVal modDoc As ModelDoc2)
        On Error Resume Next
        Dim docHandler As DocumentEventHandler = openDocs.Item(modDoc)
        openDocs.Remove(modDoc)
        modDoc = Nothing
        docHandler = Nothing
        If Err.Number <> 0 Then Err.Clear()
    End Sub
#End Region
#Region "Event Handlers"
    Function SldWorks_ActiveDocChangeNotify() As Integer
        'TODO: Add your implementation here
        Return 0
    End Function
    Function SldWorks_DocumentLoadNotify2(ByVal docTitle As String, ByVal docPath As String) As Integer
        On Error Resume Next
        AttachEventsToAllDocuments()
        If Err.Number <> 0 Then Err.Clear()
        Return 0
    End Function
    Function SldWorks_FileNewNotify2(ByVal newDoc As Object, ByVal doctype As Integer, ByVal templateName As String) As Integer
        On Error Resume Next
        AttachEventsToAllDocuments()
        If Err.Number <> 0 Then Err.Clear()
        Return 0
    End Function
    Function SldWorks_ActiveModelDocChangeNotify() As Integer
        'TODO: Add your implementation here
        Return 0
    End Function
    Function SldWorks_FileOpenPostNotify(ByVal FileName As String) As Integer
        On Error Resume Next
        AttachEventsToAllDocuments()
        If Err.Number <> 0 Then Err.Clear()
        Return 0
    End Function
#End Region
End Class

