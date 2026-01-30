Option Explicit On
Imports System.ComponentModel
Imports SolidWorks.Interop
Imports System.Collections.ObjectModel
Friend Class CStepFile
    Implements INotifyPropertyChanged

    Private bIsSelected As Boolean = True
    Private strFullFileName As String = ""
    Private strShape As String = ""
    Private dWallThickness As Double = 0
    Private dCutLength As Double = 0
    Private dMaterialLength As Double = 0
    Private iNumberOfHoles As Integer = 0
    Private strCrossSection As String = ""

    Private isModelPopulated As Boolean = False

    Private swModel As sldworks.ModelDoc2 = Nothing
    Private swSelectData As sldworks.SelectData = Nothing
    Private swSelectionManager As sldworks.SelectionMgr = Nothing
    Private swImportStepFileData As sldworks.ImportStepData = Nothing
    Private swUserUnit As sldworks.UserUnit = Nothing

    Private oFaces As CFaceCollection = Nothing
    Private oPossibleShapesCollection As ObservableCollection(Of String)

    Private oEdgesOfHoles As List(Of sldworks.Edge)
    Private oEdgesForOnlyCutLengthAndNotHoles As List(Of sldworks.Edge)
    Private dStartPointOfMaterialLength(2) As Double
    Private dEndPointOfMaterialLength(2) As Double
    Private oCalloutForMaterialLength As sldworks.Callout = Nothing
    Private oCalloutHandler As New CCalloutHandler()
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Public Sub New(ByRef sFullFileName As String)
        If Not String.IsNullOrEmpty(sFullFileName) Then
            strFullFileName = sFullFileName
            bIsSelected = True
            oPossibleShapesCollection = New ObservableCollection(Of String) From {
                strNoShape,
                strRoundShape,
                strRectangleShape,
                strSquareShape,
                strAngleShape,
                strChannelShape
            }
        End If
    End Sub
    Public Property IsPopulatedInSW As Boolean
        Get
            Return isModelPopulated
        End Get
        Set(value As Boolean)
            isModelPopulated = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(""))
        End Set
    End Property
    Public ReadOnly Property PossibleShapes() As ObservableCollection(Of String)
        Get
            Return oPossibleShapesCollection
        End Get
    End Property
    Public Property IsSelected As Boolean
        Get
            Return bIsSelected
        End Get
        Set(value As Boolean)
            bIsSelected = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(""))
        End Set
    End Property
    Public ReadOnly Property FileName As String
        Get
            Return FileIO.FileSystem.GetName(strFullFileName)
        End Get
    End Property
    Public ReadOnly Property FullFileName As String
        Get
            Return strFullFileName
        End Get
    End Property
    Public Property Shape As String
        Get
            Return strShape
        End Get
        Set(value As String)
            strShape = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(""))
        End Set
    End Property
    Public Property CrossSection As String
        Get
            Return strCrossSection
        End Get
        Set(value As String)
            strCrossSection = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(""))
        End Set
    End Property
    Public Property WallThickness As Double
        Get
            Return dWallThickness
        End Get
        Set(value As Double)
            dWallThickness = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(""))
        End Set
    End Property
    Public Property CutLength As Double
        Get
            Return dCutLength
        End Get
        Set(value As Double)
            dCutLength = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(""))
        End Set
    End Property
    Public Property NumberOfHoles As Integer
        Get
            Return iNumberOfHoles
        End Get
        Set(value As Integer)
            iNumberOfHoles = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(""))
        End Set
    End Property
    Public Property MaterialLength As Double
        Get
            Return dMaterialLength
        End Get
        Set(value As Double)
            dMaterialLength = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(""))
        End Set
    End Property
    Public Function ReadFileAndPopulateModel() As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not strFullFileName Is Nothing Then
            swImportStepFileData = SWApp.GetImportFileData(strFullFileName)
            If Not swImportStepFileData Is Nothing Then
                swImportStepFileData.MapConfigurationData = False
                Dim lError As Integer = 0
                My.Settings.IsDocumentLoading = True
                My.Settings.Save()
                swModel = SWApp.LoadFile4(strFullFileName, "r", swImportStepFileData, lError)
                My.Settings.IsDocumentLoading = False
                My.Settings.Save()
                If Not swModel Is Nothing Then
                    UpdateView()
                    swSelectionManager = swModel.SelectionManager
                    swSelectData = swSelectionManager.CreateSelectData()
                    swUserUnit = swModel.GetUserUnit(swconst.swUserUnitsType_e.swLengthUnit)
                End If
            Else
                swModel = SWApp.ActiveDoc
                If Not swModel Is Nothing Then
                    UpdateView()
                    swSelectionManager = swModel.SelectionManager
                    swSelectData = swSelectionManager.CreateSelectData()
                    swUserUnit = swModel.GetUserUnit(swconst.swUserUnitsType_e.swLengthUnit)
                End If
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public ReadOnly Property Bodies() As Object
        Get
            Dim oReturn As Object = Nothing
            If Not swModel Is Nothing Then
                Dim swPart As sldworks.PartDoc = swModel
                If Not swPart Is Nothing Then
                    oReturn = swPart.GetBodies2(swconst.swBodyType_e.swSolidBody, False)
                End If
            End If
            Return oReturn
        End Get
    End Property
    Private Function GetAllEdgesForCutLength() As List(Of sldworks.Edge)
        On Error Resume Next
        Dim oReturn As New List(Of sldworks.Edge)
        Dim i As Integer = 0
        If Not oEdgesForOnlyCutLengthAndNotHoles Is Nothing Then
            For i = 0 To oEdgesForOnlyCutLengthAndNotHoles.Count - 1
                oReturn.Add(oEdgesForOnlyCutLengthAndNotHoles.Item(i))
            Next i
        End If
        If Not oEdgesOfHoles Is Nothing Then
            For i = 0 To oEdgesOfHoles.Count - 1
                oReturn.Add(oEdgesOfHoles.Item(i))
            Next i
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return oReturn
    End Function
    Private Function GetEdgesAsEntity(ByRef oEdges As List(Of sldworks.Edge)) As List(Of sldworks.Entity)
        On Error Resume Next
        Dim oReturn As New List(Of sldworks.Entity)
        If Not oEdges Is Nothing Then
            For j As Integer = 0 To oEdges.Count - 1
                oReturn.Add(oEdges(j))
            Next j
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return oReturn
    End Function
    Public Function ExtractData() As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If IsSelected Then
            If Not String.IsNullOrEmpty(strFullFileName) Then
                ReadFileAndPopulateModel()
                If Not swModel Is Nothing Then
                    IsPopulatedInSW = True
                    Dim oBodies As Object = Bodies
                    If Not oBodies Is Nothing Then
                        oFaces = New CFaceCollection
                        If Not oFaces Is Nothing Then
                            Dim i As Integer = 0
                            For i = LBound(oBodies) To UBound(oBodies)
                                oFaces.AddFacesOfThisBody(oBodies(i))
                            Next i
                            Dim eShape As EnumShape = EnumShape.none
                            Dim dWallTck As Double = 0
                            Dim dMatL As Double = 0
                            oFaces.SearchMaxAreaFaceAndSaveIndex(swModel, swSelectData, eShape, dWallTck, dMatL, oEdgesOfHoles, oEdgesForOnlyCutLengthAndNotHoles, strCrossSection)
                            Select Case eShape
                                Case EnumShape.round
                                    Shape = strRoundShape
                                Case EnumShape.rectangle
                                    Shape = strRectangleShape
                                Case EnumShape.square
                                    Shape = strSquareShape
                                Case EnumShape.angle
                                    Shape = strAngleShape
                                Case EnumShape.channel
                                    Shape = strChannelShape
                                Case EnumShape.none
                                    Shape = strNoShape
                            End Select
                            If Shape = strRoundShape Then
                                If Not String.IsNullOrEmpty(strCrossSection) Then
                                    CrossSection = "Ø" & swUserUnit.ConvertToUserUnit(CDbl(strCrossSection), False, False)
                                End If
                            Else
                                If Not String.IsNullOrEmpty(strCrossSection) Then
                                    If strCrossSection.Contains("x") Then
                                        Dim obj As Object = Split(strCrossSection, "x")
                                        If Not obj Is Nothing Then
                                            CrossSection = swUserUnit.ConvertToUserUnit(CDbl(Trim(obj(0))), False, False) & " x " & swUserUnit.ConvertToUserUnit(CDbl(Trim(obj(1))), False, False)
                                        End If
                                        obj = Nothing
                                    End If
                                End If
                            End If
                            WallThickness = swUserUnit.ConvertToUserUnit(dWallTck, False, False)
                            If Not oEdgesOfHoles Is Nothing And Not oEdgesForOnlyCutLengthAndNotHoles Is Nothing Then
                                CutLength = swUserUnit.ConvertToUserUnit(GetCutLengthFromEdgeCollection(GetEdgesAsEntity(GetAllEdgesForCutLength()), swModel, swSelectData), False, False)
                            End If
                            MaterialLength = swUserUnit.ConvertToUserUnit(dMatL, False, False)
                            NumberOfHoles = oFaces.NumberOfHoles
                            dStartPointOfMaterialLength(0) = oFaces.StartPointOfMaterialLength(0)
                            dStartPointOfMaterialLength(1) = oFaces.StartPointOfMaterialLength(1)
                            dStartPointOfMaterialLength(2) = oFaces.StartPointOfMaterialLength(2)

                            dEndPointOfMaterialLength(0) = oFaces.EndPointOfMaterialLength(0)
                            dEndPointOfMaterialLength(1) = oFaces.EndPointOfMaterialLength(1)
                            dEndPointOfMaterialLength(2) = oFaces.EndPointOfMaterialLength(2)
                        End If
                    End If
                    oBodies = Nothing
                End If
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Sub SaveDataAsCustomProperties()
        On Error Resume Next
        If Not swModel Is Nothing Then
            Dim swModelDocExtension As sldworks.ModelDocExtension = swModel.Extension
            If Not swModelDocExtension Is Nothing Then
                Dim swCustomPropertyManager As sldworks.CustomPropertyManager = swModelDocExtension.CustomPropertyManager("")
                If Not swCustomPropertyManager Is Nothing Then
                    swCustomPropertyManager.Add3("Shape", swconst.swCustomInfoType_e.swCustomInfoText, Shape, swconst.swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
                    swCustomPropertyManager.Add3("CrossSection", swconst.swCustomInfoType_e.swCustomInfoText, CrossSection, swconst.swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
                    swCustomPropertyManager.Add3("Material Length", swconst.swCustomInfoType_e.swCustomInfoText, MaterialLength, swconst.swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
                    swCustomPropertyManager.Add3("Wall Thickness", swconst.swCustomInfoType_e.swCustomInfoText, WallThickness, swconst.swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
                    swCustomPropertyManager.Add3("Cut Length", swconst.swCustomInfoType_e.swCustomInfoText, CutLength, swconst.swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
                    swCustomPropertyManager.Add3("Number of Holes", swconst.swCustomInfoType_e.swCustomInfoText, NumberOfHoles, swconst.swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd)
                End If
                swCustomPropertyManager = Nothing
            End If
            swModelDocExtension = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub SaveModel()
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not swModel Is Nothing Then
            Dim strDocPath As String = GetParentFolderName(strFullFileName)
            strDocPath = FileIO.FileSystem.CombinePath(strDocPath, strOutputfolderName)
            If Not IO.Directory.Exists(strDocPath) Then
                FileIO.FileSystem.CreateDirectory(strDocPath)
            End If
            If IO.Directory.Exists(strDocPath) Then
                Dim strFileNameOfSWFile As String = FileIO.FileSystem.CombinePath(strDocPath, GetFileNameWithoutExtensionFromFilePathName(strFullFileName) & ".sldprt")
                Dim lError As Integer = 0
                Dim lWarning As Integer = 0
                bReturn = swModel.Extension.SaveAs(strFileNameOfSWFile,
                                        swconst.swSaveAsVersion_e.swSaveAsCurrentVersion,
                                        swconst.swSaveAsOptions_e.swSaveAsOptions_Silent,
                                        Nothing, lError, lWarning)
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub ClearAllSelection()
        On Error Resume Next
        If Not swModel Is Nothing Then
            swModel.ClearSelection2(True)
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub SelectAllEdgesForCutLength()
        On Error Resume Next
        If Not swModel Is Nothing Then
            ActivateMe()
            ClearAllSelection()
            If Not oEdgesForOnlyCutLengthAndNotHoles Is Nothing And Not oEdgesOfHoles Is Nothing Then
                For i As Integer = 0 To oEdgesForOnlyCutLengthAndNotHoles.Count - 1
                    SelectEdge(oEdgesForOnlyCutLengthAndNotHoles.Item(i), swSelectData)
                Next i
                For i As Integer = 0 To oEdgesOfHoles.Count - 1
                    SelectEdge(oEdgesOfHoles.Item(i), swSelectData)
                Next i
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub SelectAllEdgesForHoles()
        On Error Resume Next
        If Not swModel Is Nothing Then
            ActivateMe()
            ClearAllSelection()
            If Not oEdgesOfHoles Is Nothing Then
                For i As Integer = 0 To oEdgesOfHoles.Count - 1
                    SelectEdge(oEdgesOfHoles.Item(i), swSelectData)
                Next i
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Protected Overrides Sub Finalize()
        On Error Resume Next
        MyBase.Finalize()
        If Not oPossibleShapesCollection Is Nothing Then
            oPossibleShapesCollection.Clear()
        End If
        oPossibleShapesCollection = Nothing

        swModel = Nothing
        swSelectData = Nothing
        swSelectionManager = Nothing
        swImportStepFileData = Nothing
        swUserUnit = Nothing

        oFaces = Nothing
        oEdgesOfHoles = Nothing
        oEdgesForOnlyCutLengthAndNotHoles = Nothing
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub ActivateMe()
        On Error Resume Next
        If Not SWApp Is Nothing And Not String.IsNullOrEmpty(strFullFileName) And Not swModel Is Nothing Then
            Dim lErr As swconst.swActivateDocError_e = 0
            Dim strFileToActivate As String = ""
            If Not swModel Is Nothing Then
                strFileToActivate = swModel.GetTitle()
            End If
            If String.IsNullOrEmpty(strFileToActivate) Then
                strFileToActivate = GetFileNameWithoutExtensionFromFilePathName(strFullFileName)
            End If
            Dim oModel As sldworks.ModelDoc2 = SWApp.ActivateDoc3(strFileToActivate, True, swconst.swRebuildOnActivation_e.swDontRebuildActiveDoc, lErr)
            If Not oModel Is Nothing Then
                swModel = SWApp.ActiveDoc()
                UpdateView()
            End If
            oModel = Nothing
            ShowHideCallout(False)
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub UpdateView()
        On Error Resume Next
        If Not swModel Is Nothing Then
            If My.Settings.ShadedWithEdges Then
                Dim oModelView As sldworks.ModelView = swModel.ActiveView
                If Not oModelView Is Nothing Then
                    oModelView.DisplayMode = swconst.swViewDisplayMode_e.swViewDisplayMode_ShadedWithEdges
                End If
                oModelView = Nothing
            Else
                Dim oModelView As sldworks.ModelView = swModel.ActiveView
                If Not oModelView Is Nothing Then
                    oModelView.DisplayMode = swconst.swViewDisplayMode_e.swViewDisplayMode_Shaded
                End If
                oModelView = Nothing
                'swModel.ViewDisplayShaded()
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Function ShowHideCallout(ByVal bShow As Boolean) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oCalloutForMaterialLength Is Nothing Then
            bReturn = oCalloutForMaterialLength.Display(bShow)
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Function CreateCallout() As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oCalloutForMaterialLength Is Nothing Then
            bReturn = ShowHideCallout(False)
            oCalloutForMaterialLength = Nothing
        End If
        oCalloutForMaterialLength = swModel.Extension.CreateCallout(2, oCalloutHandler)
        If Not oCalloutForMaterialLength Is Nothing Then
            'oCalloutForMaterialLength.Label2(0) = "Material Length"
            'oCalloutForMaterialLength.Value(0) = dMaterialLength
            'oCalloutForMaterialLength.IgnoreValue(0) = False
            'oCalloutForMaterialLength.ValueInactive(0) = True

            oCalloutForMaterialLength.Label2(0) = "Start Point"
            oCalloutForMaterialLength.Value(0) = "(" & dStartPointOfMaterialLength(0) & "," & dStartPointOfMaterialLength(1) & "," & dStartPointOfMaterialLength(2) & ")"
            oCalloutForMaterialLength.IgnoreValue(0) = False
            oCalloutForMaterialLength.SetTargetPoint(0, dStartPointOfMaterialLength(0), dStartPointOfMaterialLength(1), dStartPointOfMaterialLength(2))
            oCalloutForMaterialLength.ValueInactive(0) = True

            oCalloutForMaterialLength.Label2(1) = "End Point"
            oCalloutForMaterialLength.Value(1) = "(" & dEndPointOfMaterialLength(0) & "," & dEndPointOfMaterialLength(1) & "," & dEndPointOfMaterialLength(2) & ")"
            oCalloutForMaterialLength.IgnoreValue(1) = False
            oCalloutForMaterialLength.SetTargetPoint(1, dEndPointOfMaterialLength(0), dEndPointOfMaterialLength(1), dEndPointOfMaterialLength(2))
            oCalloutForMaterialLength.ValueInactive(1) = True

            oCalloutForMaterialLength.SetLeader(True, True)
            oCalloutForMaterialLength.Position = CreateMathPointAt((dStartPointOfMaterialLength(0) + dEndPointOfMaterialLength(0)) / 2, (dStartPointOfMaterialLength(1) + dEndPointOfMaterialLength(1)) / 2, (dStartPointOfMaterialLength(2) + dEndPointOfMaterialLength(2)) / 2)
            bReturn = ShowHideCallout(True)
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
End Class
