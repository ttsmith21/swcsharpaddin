Option Explicit On
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
'Base class for model evnt handlers
Public Class DocumentEventHandler
    Protected openModelViews As New Hashtable()
    Protected userAddin As SwAddin
    Protected iDocument As ModelDoc2
    Protected SWApp As SldWorks
    Overridable Function Init(ByVal sw As SldWorks, ByVal addin As SwAddin, ByVal model As ModelDoc2) As Boolean
        Return True
    End Function
    Overridable Function AttachEventHandlers() As Boolean
        Return True
    End Function
    Overridable Function DetachEventHandlers() As Boolean
        Return True
    End Function
    Function ConnectModelViews() As Boolean
        On Error Resume Next
        Dim iModelView As ModelView
        iModelView = iDocument.GetFirstModelView()

        While (Not iModelView Is Nothing)
            If Not openModelViews.Contains(iModelView) Then
                Dim mView As New DocView()
                mView.Init(userAddin, iModelView, Me)
                mView.AttachEventHandlers()
                openModelViews.Add(iModelView, mView)
            End If
            iModelView = iModelView.GetNext
        End While
        If Err.Number <> 0 Then Err.Clear()
        Return True
    End Function
    Function DisconnectModelViews() As Boolean
        On Error Resume Next
        'Close events on all currently open docs
        Dim mView As DocView = Nothing
        Dim key As ModelView = Nothing
        Dim numKeys As Integer = openModelViews.Count
        Dim keys() As Object = New Object(numKeys - 1) {}

        'Remove all ModelView event handlers
        openModelViews.Keys.CopyTo(keys, 0)
        For Each key In keys
            mView = openModelViews.Item(key)
            mView.DetachEventHandlers()
            openModelViews.Remove(key)
            mView = Nothing
            key = Nothing
        Next
        If Err.Number <> 0 Then Err.Clear()
        Return True
    End Function
    Sub DetachModelViewEventHandler(ByVal mView As ModelView)
        On Error Resume Next
        Dim docView As DocView = Nothing
        If openModelViews.Contains(mView) Then
            docView = openModelViews.Item(mView)
            openModelViews.Remove(mView)
            mView = Nothing
            docView = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
End Class
'Class to listen for Part Events
Public Class PartEventHandler
    Inherits DocumentEventHandler
    Dim WithEvents SWPart As PartDoc
    Overrides Function Init(ByVal sw As SldWorks, ByVal addin As SwAddin, ByVal model As ModelDoc2) As Boolean
        On Error Resume Next
        userAddin = addin
        SWPart = model
        iDocument = SWPart
        SWApp = sw
        If Err.Number <> 0 Then Err.Clear()
        Return True
    End Function
    Overrides Function AttachEventHandlers() As Boolean
        On Error Resume Next
        AddHandler SWPart.DestroyNotify, AddressOf Me.PartDoc_DestroyNotify
        AddHandler SWPart.NewSelectionNotify, AddressOf Me.PartDoc_NewSelectionNotify
        ConnectModelViews()
        If Err.Number <> 0 Then Err.Clear()
        Return True
    End Function
    Overrides Function DetachEventHandlers() As Boolean
        On Error Resume Next
        RemoveHandler SWPart.DestroyNotify, AddressOf Me.PartDoc_DestroyNotify
        RemoveHandler SWPart.NewSelectionNotify, AddressOf Me.PartDoc_NewSelectionNotify
        DisconnectModelViews()
        userAddin.DetachModelEventHandler(iDocument)
        If Err.Number <> 0 Then Err.Clear()
        Return True
    End Function
    Function PartDoc_DestroyNotify() As Integer
        On Error Resume Next
        DetachEventHandlers()
        If Err.Number <> 0 Then Err.Clear()
        Return 0
    End Function
    Function PartDoc_NewSelectionNotify() As Integer
        Return 0
    End Function
End Class
'Class to listen for Assembly Events
Public Class AssemblyEventHandler
    Inherits DocumentEventHandler

    Dim WithEvents SWAssembly As AssemblyDoc
    Dim swAddin As SwAddin

    Overrides Function Init(ByVal sw As SldWorks, ByVal addin As SwAddin, ByVal model As ModelDoc2) As Boolean
        On Error Resume Next
        userAddin = addin
        SWAssembly = model
        iDocument = SWAssembly
        SWApp = sw
        swAddin = addin
        If Err.Number <> 0 Then Err.Clear()
        Return True
    End Function
    Overrides Function AttachEventHandlers() As Boolean
        On Error Resume Next
        AddHandler SWAssembly.DestroyNotify, AddressOf Me.AssemblyDoc_DestroyNotify
        AddHandler SWAssembly.NewSelectionNotify, AddressOf Me.AssemblyDoc_NewSelectionNotify
        AddHandler SWAssembly.ComponentStateChangeNotify, AddressOf Me.AssemblyDoc_ComponentStateChangeNotify
        AddHandler SWAssembly.ComponentStateChangeNotify2, AddressOf Me.AssemblyDoc_ComponentStateChangeNotify2
        AddHandler SWAssembly.ComponentVisualPropertiesChangeNotify, AddressOf Me.AssemblyDoc_ComponentVisiblePropertiesChangeNotify
        AddHandler SWAssembly.ComponentDisplayStateChangeNotify, AddressOf Me.AssemblyDoc_ComponentDisplayStateChangeNotify
        ConnectModelViews()
        If Err.Number <> 0 Then Err.Clear()
        Return True
    End Function
    Overrides Function DetachEventHandlers() As Boolean
        On Error Resume Next
        RemoveHandler SWAssembly.DestroyNotify, AddressOf Me.AssemblyDoc_DestroyNotify
        RemoveHandler SWAssembly.NewSelectionNotify, AddressOf Me.AssemblyDoc_NewSelectionNotify
        RemoveHandler SWAssembly.ComponentStateChangeNotify, AddressOf Me.AssemblyDoc_ComponentStateChangeNotify
        RemoveHandler SWAssembly.ComponentStateChangeNotify2, AddressOf Me.AssemblyDoc_ComponentStateChangeNotify2
        RemoveHandler SWAssembly.ComponentVisualPropertiesChangeNotify, AddressOf Me.AssemblyDoc_ComponentVisiblePropertiesChangeNotify
        RemoveHandler SWAssembly.ComponentDisplayStateChangeNotify, AddressOf Me.AssemblyDoc_ComponentDisplayStateChangeNotify
        DisconnectModelViews()
        userAddin.DetachModelEventHandler(iDocument)
        If Err.Number <> 0 Then Err.Clear()
        Return True
    End Function
    Function AssemblyDoc_DestroyNotify() As Integer
        On Error Resume Next
        DetachEventHandlers()
        If Err.Number <> 0 Then Err.Clear()
        Return 0
    End Function
    Function AssemblyDoc_NewSelectionNotify() As Integer
        Return 0
    End Function
    Protected Function ComponentStateChange(ByVal componentModel As Object, Optional ByVal newCompState As Short = swComponentSuppressionState_e.swComponentResolved) As Integer
        On Error Resume Next
        Dim modDoc As ModelDoc2 = componentModel
        Dim newState As swComponentSuppressionState_e = newCompState
        Select Case newState
            Case swComponentSuppressionState_e.swComponentFullyResolved, swComponentSuppressionState_e.swComponentResolved
                If ((Not modDoc Is Nothing) AndAlso Not Me.swAddin.OpenDocumentsTable.Contains(modDoc)) Then
                    Me.swAddin.AttachModelDocEventHandler(modDoc)
                End If
                Exit Select
        End Select
        If Err.Number <> 0 Then Err.Clear()
        Return 0
    End Function
    'attach events to a component if it becomes resolved
    Public Function AssemblyDoc_ComponentStateChangeNotify(ByVal componentModel As Object, ByVal oldCompState As Short, ByVal newCompState As Short) As Integer
        Return ComponentStateChange(componentModel, newCompState)
    End Function
    'attach events to a component if it becomes resolved
    Public Function AssemblyDoc_ComponentStateChangeNotify2(ByVal componentModel As Object, ByVal CompName As String, ByVal oldCompState As Short, ByVal newCompState As Short) As Integer
        Return ComponentStateChange(componentModel, newCompState)
    End Function
    Public Function AssemblyDoc_ComponentVisiblePropertiesChangeNotify(ByVal swObject As Object) As Integer
        On Error Resume Next
        Dim component As Component2 = swObject
        Dim modDoc As ModelDoc2 = component.GetModelDoc2()
        If Err.Number <> 0 Then Err.Clear()
        Return ComponentStateChange(modDoc)
    End Function
    Public Function AssemblyDoc_ComponentDisplayStateChangeNotify(ByVal swObject As Object) As Integer
        On Error Resume Next
        Dim component As Component2 = swObject
        Dim modDoc As ModelDoc2 = component.GetModelDoc2()
        If Err.Number <> 0 Then Err.Clear()
        Return ComponentStateChange(modDoc)
    End Function
End Class
'Class to listen for Drawing Events
Public Class DrawingEventHandler
    Inherits DocumentEventHandler
    Dim WithEvents SWDrawing As DrawingDoc
    Overrides Function Init(ByVal sw As SldWorks, ByVal addin As SwAddin, ByVal model As ModelDoc2) As Boolean
        On Error Resume Next
        userAddin = addin
        SWDrawing = model
        iDocument = SWDrawing
        SWApp = sw
        If Err.Number <> 0 Then Err.Clear()
        Return True
    End Function
    Overrides Function AttachEventHandlers() As Boolean
        On Error Resume Next
        AddHandler SWDrawing.DestroyNotify, AddressOf Me.DrawingDoc_DestroyNotify
        AddHandler SWDrawing.NewSelectionNotify, AddressOf Me.DrawingDoc_NewSelectionNotify
        Dim bReturn As Boolean = ConnectModelViews()
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Overrides Function DetachEventHandlers() As Boolean
        On Error Resume Next
        RemoveHandler SWDrawing.DestroyNotify, AddressOf Me.DrawingDoc_DestroyNotify
        RemoveHandler SWDrawing.NewSelectionNotify, AddressOf Me.DrawingDoc_NewSelectionNotify
        Dim bReturn As Boolean = DisconnectModelViews()
        userAddin.DetachModelEventHandler(iDocument)
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Function DrawingDoc_DestroyNotify() As Integer
        On Error Resume Next
        DetachEventHandlers()
        If Err.Number <> 0 Then Err.Clear()
        Return 0
    End Function
    Function DrawingDoc_NewSelectionNotify() As Integer
        Return 0
    End Function
End Class
'Class for handling ModelView events
Public Class DocView
    Dim WithEvents SWModelView As ModelView
    Dim userAddin As SwAddin
    Dim parentDoc As DocumentEventHandler
    Function Init(ByVal addin As SwAddin, ByVal mView As ModelView, ByVal parent As DocumentEventHandler) As Boolean
        On Error Resume Next
        userAddin = addin
        SWModelView = mView
        parentDoc = parent
        If Err.Number <> 0 Then Err.Clear()
        Return True
    End Function
    Function AttachEventHandlers() As Boolean
        On Error Resume Next
        AddHandler SWModelView.DestroyNotify2, AddressOf Me.ModelView_DestroyNotify2
        AddHandler SWModelView.RepaintNotify, AddressOf Me.ModelView_RepaintNotify
        AddHandler SWModelView.DisplayModeChangePostNotify, AddressOf Me.ModelView_DisplayModeChangePostNotify
        If Err.Number <> 0 Then Err.Clear()
        Return True
    End Function
    Function DetachEventHandlers() As Boolean
        On Error Resume Next
        RemoveHandler SWModelView.DestroyNotify2, AddressOf Me.ModelView_DestroyNotify2
        RemoveHandler SWModelView.RepaintNotify, AddressOf Me.ModelView_RepaintNotify
        RemoveHandler SWModelView.DisplayModeChangePostNotify, AddressOf Me.ModelView_DisplayModeChangePostNotify
        parentDoc.DetachModelViewEventHandler(SWModelView)
        If Err.Number <> 0 Then Err.Clear()
        Return True
    End Function
    Function ModelView_DestroyNotify2(ByVal destroyTYpe As Integer) As Integer
        On Error Resume Next
        DetachEventHandlers()
        If Err.Number <> 0 Then Err.Clear()
        Return 0
    End Function
    Function ModelView_RepaintNotify(ByVal repaintTYpe As Integer) As Integer
        Return 0
    End Function
    Function ModelView_DisplayModeChangePostNotify() As Integer
        On Error Resume Next
        If userAddin.UCUIAddinObject.oUCUI.IsDocumentLoading = False Then
            If Not SWModelView Is Nothing Then
                If SWModelView.DisplayMode = SolidWorks.Interop.swconst.swViewDisplayMode_e.swViewDisplayMode_ShadedWithEdges Then
                    If Not userAddin Is Nothing Then
                        userAddin.UCUIAddinObject.oUCUI.ViewModeAsShadedWithEdges = True
                    End If
                Else
                    If Not userAddin Is Nothing Then
                        userAddin.UCUIAddinObject.oUCUI.ViewModeAsShadedWithEdges = False
                    End If
                End If
            End If
        End If
        Return 0
    End Function
End Class
