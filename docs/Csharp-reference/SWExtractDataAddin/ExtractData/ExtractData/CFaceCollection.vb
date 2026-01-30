Option Explicit On
Imports SolidWorks.Interop
Friend Class CFaceCollection
    Private oFaceCollection As List(Of CFace)
    Private oMaxAreaFaceIndexCollection As List(Of Integer)
    Private oHoles As CLoopCollection
    Private iIndexOfRoundFace As Integer = -1
    Private dStartPointOfMaterialLength(2) As Double
    Private dEndPointOfMaterialLength(2) As Double
    Public Sub New()
        On Error Resume Next
        oFaceCollection = New List(Of CFace)
        oMaxAreaFaceIndexCollection = New List(Of Integer)
        oHoles = New CLoopCollection()
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Protected Overrides Sub Finalize()
        On Error Resume Next
        MyBase.Finalize()
        oFaceCollection = Nothing
        oMaxAreaFaceIndexCollection = Nothing
        oHoles = Nothing
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub Clear()
        On Error Resume Next
        If Not oFaceCollection Is Nothing Then
            oFaceCollection.Clear()
        End If
        If Not oMaxAreaFaceIndexCollection Is Nothing Then
            oMaxAreaFaceIndexCollection.Clear()
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Function DoesThisFaceExist(ByRef oCFace As CFace) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oCFace Is Nothing And Not oFaceCollection Is Nothing Then
            Dim i As Integer = 0
            While i < oFaceCollection.Count And bReturn = False
                If oFaceCollection.Item(i).IsSame(oCFace) Then
                    bReturn = True
                End If
                i = i + 1
            End While
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Function DoesThisFaceExist(ByRef oFace As sldworks.Face2) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oFace Is Nothing And Not oFaceCollection Is Nothing Then
            bReturn = DoesThisFaceExist(New CFace(oFace))
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public ReadOnly Property Count As Integer
        Get
            On Error Resume Next
            If Not oFaceCollection Is Nothing Then
                Return oFaceCollection.Count
            End If
            If Err.Number <> 0 Then Err.Clear()
            Return 0
        End Get
    End Property
    Public ReadOnly Property NumberOfHoles() As Integer
        Get
            Dim iReturn As Integer = 0
            If Not oHoles Is Nothing Then
                If oHoles.Count > 0 Then
                    iReturn = oHoles.Count
                Else
                    If iIndexOfRoundFace > 0 Then
                        If Not oFaceCollection Is Nothing Then
                            oHoles = oFaceCollection.Item(iIndexOfRoundFace).GetHoles()
                            iReturn = oHoles.Count
                        End If
                    End If
                End If
            End If
            Return iReturn
        End Get
    End Property
    Public ReadOnly Property StartPointOfMaterialLength As Object
        Get
            Return dStartPointOfMaterialLength
        End Get
    End Property
    Public ReadOnly Property EndPointOfMaterialLength As Object
        Get
            Return dEndPointOfMaterialLength
        End Get
    End Property
    Private Sub Add(ByRef oFace As sldworks.Face2)
        On Error Resume Next
        If Not oFace Is Nothing And Not oFaceCollection Is Nothing Then
            oFaceCollection.Add(New CFace(oFace))
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub AddFacesOfThisBody(ByRef oBody As sldworks.Body2)
        On Error Resume Next
        If Not oBody Is Nothing Then
            Dim oFaces As Object = oBody.GetFaces()
            If Not oFaces Is Nothing Then
                Dim i As Integer = 0
                For i = LBound(oFaces) To UBound(oFaces)
                    Add(oFaces(i))
                Next i
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub SearchMaxAreaFaceAndSaveIndex(ByRef swModel As sldworks.ModelDoc2, ByRef swSelectData As sldworks.SelectData, ByRef eShape As EnumShape, ByRef dWallThickness As Double, ByRef dMaterialLength As Double, ByRef oEdgesForHoles As List(Of sldworks.Edge), ByRef oEdgesForOnlyCutLengthAndNotHoles As List(Of sldworks.Edge), ByRef strCrossSection As String)
        On Error Resume Next
        If Not oFaceCollection Is Nothing Then
            Dim dMaxArea As Double = 0
            Dim i As Integer = 0
            For i = 0 To oFaceCollection.Count - 1
                dMaxArea = Math.Max(dMaxArea, oFaceCollection.Item(i).Area)
            Next i
            For i = 0 To oFaceCollection.Count - 1
                If IsTendsToZero(oFaceCollection.Item(i).Area - dMaxArea) Then
                    oMaxAreaFaceIndexCollection.Add(i)
                End If
            Next i
            'compute Shape
            eShape = ComputeShape(swModel, swSelectData, dWallThickness, dMaterialLength, oEdgesForHoles, oEdgesForOnlyCutLengthAndNotHoles, strCrossSection)
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Function ComputeShape(ByRef swModel As sldworks.ModelDoc2, ByRef swSelectData As sldworks.SelectData, ByRef dWallThickness As Double, ByRef dMaterialLength As Double, ByRef oEdgesForHoles As List(Of sldworks.Edge), ByRef oEdgesForOnlyCutLengthAndNotHoles As List(Of sldworks.Edge), ByRef strCrossSection As String) As EnumShape
        On Error Resume Next
        Dim eShape As EnumShape = EnumShape.none
        If Not oMaxAreaFaceIndexCollection Is Nothing Then
            Dim i As Integer = 0
            'Check if any face is round..
            Dim bRound As Boolean = False
            While i < oMaxAreaFaceIndexCollection.Count And bRound = False
                If oFaceCollection.Item(oMaxAreaFaceIndexCollection.Item(i)).IsRound Then
                    bRound = True
                    iIndexOfRoundFace = oMaxAreaFaceIndexCollection.Item(i)
                End If
                i = i + 1
            End While
            If bRound Then
                eShape = EnumShape.round
                dWallThickness = GetWallThicknessForRoundProfile(strCrossSection)
                Dim iHoles As Integer = NumberOfHoles
                oEdgesForHoles = oHoles.GetEdges()
                oEdgesForOnlyCutLengthAndNotHoles = GetEdgesForCutLengthOtherThanHolesAndComputeMaterialLengthForRoundProfile(dMaterialLength)
            Else
                'get any planar face with max area..
                'we do not need max area face collection now..
                Dim oFaceIndexesCollectionForCutLength As New List(Of Integer)
                If Not oFaceIndexesCollectionForCutLength Is Nothing Then
                    Dim dMaxArea As Double = 0
                    Dim indexOfLargestPlanarFace As Integer = -1
                    For i = 0 To oFaceCollection.Count - 1
                        If oFaceCollection.Item(i).IsPlanar Then
                            dMaxArea = Math.Max(dMaxArea, oFaceCollection.Item(i).Area)
                            If IsTendsToZero(dMaxArea - oFaceCollection.Item(i).Area) Then
                                indexOfLargestPlanarFace = i
                            End If
                        End If
                    Next i
                    If indexOfLargestPlanarFace <> -1 Then
                        'get largest linear edge to define axis.
                        Dim oPrimaryEdgeDirection As Object = oFaceCollection.Item(indexOfLargestPlanarFace).GetDirectionOfLargestLinearEdge(dStartPointOfMaterialLength, dEndPointOfMaterialLength, dMaterialLength)
                        If Not oPrimaryEdgeDirection Is Nothing Then
                            Dim dLength As Double = 0
                            Dim dWidth As Double = 0
                            Dim dHeight As Double = 0
                            Dim oFaceIndexesParallelToPrimaryFace As New List(Of Integer)
                            Dim oFirstExtremeFaceIndexesCollectionForPrimaryFace As New List(Of Integer)
                            Dim oSecondExtremeFaceIndexesCollectionForPrimaryFace As New List(Of Integer)
                            Dim oDistancesCollectionForPrimaryFace As New List(Of Double)
                            Call GetAllPlanarFacesWhichAreNormalToThisDirection(oFaceCollection.Item(indexOfLargestPlanarFace).Normal,
                                                                                oFaceIndexesParallelToPrimaryFace)
                            dWallThickness = -1
                            Call GetMinAndMaxDistanceFromTheseParallelFaces(swModel, swSelectData,
                                                                            dWallThickness, dHeight,
                                                                            oFaceIndexesParallelToPrimaryFace,
                                                                            oFirstExtremeFaceIndexesCollectionForPrimaryFace,
                                                                            oSecondExtremeFaceIndexesCollectionForPrimaryFace,
                                                                            oDistancesCollectionForPrimaryFace)

                            CopyAllFaceIndexesFromThisCollectionToThisCollection(oFirstExtremeFaceIndexesCollectionForPrimaryFace, oFaceIndexesCollectionForCutLength)
                            CopyAllFaceIndexesFromThisCollectionToThisCollection(oSecondExtremeFaceIndexesCollectionForPrimaryFace, oFaceIndexesCollectionForCutLength)

                            Dim oFaceIndexesNormalToPrimaryEdge As New List(Of Integer)
                            Dim oFirstExtremeFaceIndexesCollectionForPrimaryEdge As New List(Of Integer)
                            Dim oSecondExtremeFaceIndexesCollectionForPrimaryEdge As New List(Of Integer)
                            Dim oDistancesCollectionForPrimaryEdge As New List(Of Double)
                            Call GetAllPlanarFacesWhichAreNormalToThisDirection(oPrimaryEdgeDirection,
                                                                                oFaceIndexesNormalToPrimaryEdge)
                            Call GetMinAndMaxDistanceFromTheseParallelFaces(swModel, swSelectData,
                                                                            dWallThickness, dLength,
                                                                            oFaceIndexesNormalToPrimaryEdge,
                                                                            oFirstExtremeFaceIndexesCollectionForPrimaryEdge,
                                                                            oSecondExtremeFaceIndexesCollectionForPrimaryEdge,
                                                                            oDistancesCollectionForPrimaryEdge)

                            Dim oFaceIndexesNormalToSecondaryEdge As New List(Of Integer)
                            Dim oFirstExtremeFaceIndexesCollectionForSecondaryEdge As New List(Of Integer)
                            Dim oSecondExtremeFaceIndexesCollectionForSecondaryEdge As New List(Of Integer)
                            Dim oDistancesCollectionForSecondaryEdge As New List(Of Double)
                            Dim oSecondaryEdgeNormal As Object = GetCrossProduct(oFaceCollection.Item(indexOfLargestPlanarFace).Normal,
                                                                                   oPrimaryEdgeDirection)

                            Call GetAllPlanarFacesWhichAreNormalToThisDirection(oSecondaryEdgeNormal,
                                                                                oFaceIndexesNormalToSecondaryEdge)
                            Call GetMinAndMaxDistanceFromTheseParallelFaces(swModel, swSelectData,
                                                                            dWallThickness, dWidth,
                                                                            oFaceIndexesNormalToSecondaryEdge,
                                                                            oFirstExtremeFaceIndexesCollectionForSecondaryEdge,
                                                                            oSecondExtremeFaceIndexesCollectionForSecondaryEdge,
                                                                            oDistancesCollectionForSecondaryEdge)
                            CopyAllFaceIndexesFromThisCollectionToThisCollection(oFirstExtremeFaceIndexesCollectionForSecondaryEdge, oFaceIndexesCollectionForCutLength)
                            CopyAllFaceIndexesFromThisCollectionToThisCollection(oSecondExtremeFaceIndexesCollectionForSecondaryEdge, oFaceIndexesCollectionForCutLength)

                            If IsTendsToZero(dHeight) Or IsTendsToZero(dWidth) Then
                                If IsTendsToZero(dHeight) And IsTendsToZero(dWidth) Then
                                    eShape = EnumShape.angle
                                Else
                                    eShape = EnumShape.channel
                                End If
                                If IsTendsToZero(dHeight) Then
                                    If oFaceCollection.Item(oFirstExtremeFaceIndexesCollectionForPrimaryFace.Item(0)).Area > oFaceCollection.Item(oSecondExtremeFaceIndexesCollectionForPrimaryFace.Item(0)).Area Then
                                        For i = 0 To oSecondExtremeFaceIndexesCollectionForPrimaryFace.Count - 1
                                            RemoveThisValueFromThisList(oFaceIndexesCollectionForCutLength, oSecondExtremeFaceIndexesCollectionForPrimaryFace.Item(i))
                                        Next i
                                    Else
                                        For i = 0 To oFirstExtremeFaceIndexesCollectionForPrimaryFace.Count - 1
                                            RemoveThisValueFromThisList(oFaceIndexesCollectionForCutLength, oFirstExtremeFaceIndexesCollectionForPrimaryFace.Item(i))
                                        Next i
                                    End If
                                End If
                                If IsTendsToZero(dWidth) Then
                                    If oFaceCollection.Item(oFirstExtremeFaceIndexesCollectionForSecondaryEdge.Item(0)).Area > oFaceCollection.Item(oSecondExtremeFaceIndexesCollectionForSecondaryEdge.Item(0)).Area Then
                                        For i = 0 To oSecondExtremeFaceIndexesCollectionForSecondaryEdge.Count - 1
                                            RemoveThisValueFromThisList(oFaceIndexesCollectionForCutLength, oSecondExtremeFaceIndexesCollectionForSecondaryEdge.Item(i))
                                        Next i
                                    Else
                                        For i = 0 To oFirstExtremeFaceIndexesCollectionForSecondaryEdge.Count - 1
                                            RemoveThisValueFromThisList(oFaceIndexesCollectionForCutLength, oFirstExtremeFaceIndexesCollectionForSecondaryEdge.Item(i))
                                        Next i
                                    End If
                                End If
                            Else
                                Dim iResult As Integer = 0
                                Dim iRemPrimary As Integer = oDistancesCollectionForPrimaryFace.Count Mod 2
                                Dim iRemSecondary As Integer = oDistancesCollectionForSecondaryEdge.Count Mod 2
                                If iRemPrimary <> 0 Or iRemSecondary <> 0 Then
                                    If iRemPrimary <> 0 And iRemSecondary <> 0 Then
                                        eShape = EnumShape.angle
                                        If oFaceCollection.Item(oFirstExtremeFaceIndexesCollectionForPrimaryFace.Item(0)).Area > oFaceCollection.Item(oSecondExtremeFaceIndexesCollectionForPrimaryFace.Item(0)).Area Then
                                            For i = 0 To oSecondExtremeFaceIndexesCollectionForPrimaryFace.Count - 1
                                                RemoveThisValueFromThisList(oFaceIndexesCollectionForCutLength, oSecondExtremeFaceIndexesCollectionForPrimaryFace.Item(i))
                                            Next i
                                        Else
                                            For i = 0 To oFirstExtremeFaceIndexesCollectionForPrimaryFace.Count - 1
                                                RemoveThisValueFromThisList(oFaceIndexesCollectionForCutLength, oFirstExtremeFaceIndexesCollectionForPrimaryFace.Item(i))
                                            Next i
                                        End If
                                        If oFaceCollection.Item(oFirstExtremeFaceIndexesCollectionForSecondaryEdge.Item(0)).Area > oFaceCollection.Item(oSecondExtremeFaceIndexesCollectionForSecondaryEdge.Item(0)).Area Then
                                            For i = 0 To oSecondExtremeFaceIndexesCollectionForSecondaryEdge.Count - 1
                                                RemoveThisValueFromThisList(oFaceIndexesCollectionForCutLength, oSecondExtremeFaceIndexesCollectionForSecondaryEdge.Item(i))
                                            Next i
                                        Else
                                            For i = 0 To oFirstExtremeFaceIndexesCollectionForSecondaryEdge.Count - 1
                                                RemoveThisValueFromThisList(oFaceIndexesCollectionForCutLength, oFirstExtremeFaceIndexesCollectionForSecondaryEdge.Item(i))
                                            Next i
                                        End If
                                    Else
                                        eShape = EnumShape.channel
                                        If iRemPrimary <> 0 Then
                                            If oFaceCollection.Item(oFirstExtremeFaceIndexesCollectionForPrimaryFace.Item(0)).Area > oFaceCollection.Item(oSecondExtremeFaceIndexesCollectionForPrimaryFace.Item(0)).Area Then
                                                For i = 0 To oSecondExtremeFaceIndexesCollectionForPrimaryFace.Count - 1
                                                    RemoveThisValueFromThisList(oFaceIndexesCollectionForCutLength, oSecondExtremeFaceIndexesCollectionForPrimaryFace.Item(i))
                                                Next i
                                            Else
                                                For i = 0 To oFirstExtremeFaceIndexesCollectionForPrimaryFace.Count - 1
                                                    RemoveThisValueFromThisList(oFaceIndexesCollectionForCutLength, oFirstExtremeFaceIndexesCollectionForPrimaryFace.Item(i))
                                                Next i
                                            End If
                                        Else
                                            If oFaceCollection.Item(oFirstExtremeFaceIndexesCollectionForSecondaryEdge.Item(0)).Area > oFaceCollection.Item(oSecondExtremeFaceIndexesCollectionForSecondaryEdge.Item(0)).Area Then
                                                For i = 0 To oSecondExtremeFaceIndexesCollectionForSecondaryEdge.Count - 1
                                                    RemoveThisValueFromThisList(oFaceIndexesCollectionForCutLength, oSecondExtremeFaceIndexesCollectionForSecondaryEdge.Item(i))
                                                Next i
                                            Else
                                                For i = 0 To oFirstExtremeFaceIndexesCollectionForSecondaryEdge.Count - 1
                                                    RemoveThisValueFromThisList(oFaceIndexesCollectionForCutLength, oFirstExtremeFaceIndexesCollectionForSecondaryEdge.Item(i))
                                                Next i
                                            End If
                                        End If
                                    End If
                                Else
                                    If IsTendsToZero(dHeight - dWidth) Then
                                        eShape = EnumShape.square
                                    Else
                                        eShape = EnumShape.rectangle
                                    End If
                                End If
                            End If
                            If dHeight > dWidth Then
                                strCrossSection = dHeight & " x " & dWidth
                            Else
                                strCrossSection = dWidth & " x " & dHeight
                            End If
                            dMaterialLength = Math.Max(dMaterialLength, dLength)
                            If dMaterialLength = 0 Then
                                'both the ends do not have planar faces or faces are not parallel..
                                'so compute using face..
                                For i = 0 To oFaceIndexesCollectionForCutLength.Count - 1
                                    dMaterialLength = Math.Max(dMaterialLength, oFaceCollection.Item(oFaceIndexesCollectionForCutLength.Item(i)).GetMaterialLength(dStartPointOfMaterialLength, dEndPointOfMaterialLength, oPrimaryEdgeDirection))
                                Next i
                            End If

                            DefineHolesForNonRoundProfiles(oFaceIndexesCollectionForCutLength)
                            'Dim iHoles As Integer = NumberOfHoles
                            oEdgesForHoles = oHoles.GetEdges()
                            oEdgesForOnlyCutLengthAndNotHoles = GetEdgesForCutLengthOtherThanHolesAndComputeMaterialLengthForOtherProfiles(oFaceIndexesCollectionForCutLength, eShape)
                            RemoveAllEdgesWhichAreParallelToThisDirection(oEdgesForOnlyCutLengthAndNotHoles, oPrimaryEdgeDirection)

                            oFaceIndexesParallelToPrimaryFace = Nothing
                            oFirstExtremeFaceIndexesCollectionForPrimaryFace = Nothing
                            oSecondExtremeFaceIndexesCollectionForPrimaryFace = Nothing
                            oDistancesCollectionForPrimaryFace = Nothing
                            oFaceIndexesNormalToPrimaryEdge = Nothing
                            oFirstExtremeFaceIndexesCollectionForPrimaryEdge = Nothing
                            oSecondExtremeFaceIndexesCollectionForPrimaryEdge = Nothing
                            oDistancesCollectionForPrimaryEdge = Nothing
                            oFaceIndexesNormalToSecondaryEdge = Nothing
                            oFirstExtremeFaceIndexesCollectionForSecondaryEdge = Nothing
                            oSecondExtremeFaceIndexesCollectionForSecondaryEdge = Nothing
                            oDistancesCollectionForSecondaryEdge = Nothing
                            oSecondaryEdgeNormal = Nothing
                        End If
                    End If
                End If
                oFaceIndexesCollectionForCutLength = Nothing
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return eShape
    End Function
    Private Sub CopyAllFaceIndexesFromThisCollectionToThisCollection(ByRef oCollectionToCopyFrom As List(Of Integer),
                                                                     ByRef oCollectionToCopyTo As List(Of Integer))
        On Error Resume Next
        If Not oCollectionToCopyFrom Is Nothing And Not oCollectionToCopyTo Is Nothing Then
            Dim i As Integer = 0
            For i = 0 To oCollectionToCopyFrom.Count - 1
                If Not oCollectionToCopyTo.Contains(oCollectionToCopyFrom.Item(i)) Then
                    oCollectionToCopyTo.Add(oCollectionToCopyFrom.Item(i))
                End If
            Next
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Function GetCrossProduct(ByVal varVec1() As Double, ByVal varVec2() As Double) As Double()
        Dim dblCross(2) As Double

        dblCross(0) = varVec1(1) * varVec2(2) - varVec1(2) * varVec2(1)
        dblCross(1) = varVec1(2) * varVec2(0) - varVec1(0) * varVec2(2)
        dblCross(2) = varVec1(0) * varVec2(1) - varVec1(1) * varVec2(0)
        GetCrossProduct = dblCross

    End Function
    Private Sub GetMinAndMaxDistanceFromTheseParallelFaces(ByRef swModel As sldworks.ModelDoc2, ByRef swSelectData As sldworks.SelectData,
                                                           ByRef dMinDistance As Double, ByRef dMaxDistance As Double,
                                                           ByRef oParallelFaceIndexes As List(Of Integer),
                                                           ByRef oFirstExtremeFaceIndexes As List(Of Integer),
                                                           ByRef oSecondExtremeFaceIndexes As List(Of Integer),
                                                           ByRef oDistancesCollection As List(Of Double))
        On Error Resume Next
        If Not swModel Is Nothing And Not swSelectData Is Nothing And
            Not oParallelFaceIndexes Is Nothing And Not oFirstExtremeFaceIndexes Is Nothing And
            Not oSecondExtremeFaceIndexes Is Nothing And Not oDistancesCollection Is Nothing Then
            If oParallelFaceIndexes.Count > 0 Then
                'dMinDistance = -1
                dMaxDistance = -1
                Dim i As Integer = 0
                Dim oMeasure As sldworks.Measure = swModel.Extension.CreateMeasure()
                If Not oMeasure Is Nothing Then
                    oMeasure.ArcOption = 0
                    oMeasure.AngleDecimalPlaces = 8
                    oMeasure.LengthDecimalPlaces = 8
                    Dim dDistance As Double = 0
                    Dim iFirstExtremeIndex As Integer = -1
                    Dim iSecondExtremeIndex As Integer = -1
                    If oParallelFaceIndexes.Count > 1 Then
                        For i = 0 To oParallelFaceIndexes.Count - 2
                            For j = i + 1 To oParallelFaceIndexes.Count - 1
                                dDistance = -1
                                oFaceCollection.Item(oParallelFaceIndexes.Item(i)).SelectFace(swSelectData, False)
                                oFaceCollection.Item(oParallelFaceIndexes.Item(j)).SelectFace(swSelectData, True)
                                If oMeasure.Calculate(Nothing) Then
                                    dDistance = oMeasure.NormalDistance
                                    If dDistance > 0 Then
                                        If dMinDistance = -1 Then
                                            dMinDistance = dDistance
                                        Else
                                            dMinDistance = Math.Min(dMinDistance, dDistance)
                                        End If
                                        If dMaxDistance = -1 Then
                                            dMaxDistance = dDistance
                                            iFirstExtremeIndex = oParallelFaceIndexes.Item(i)
                                            iSecondExtremeIndex = oParallelFaceIndexes.Item(j)
                                        Else
                                            dMaxDistance = Math.Max(dMaxDistance, dDistance)
                                            If IsTendsToZero(dMaxDistance - dDistance) Then
                                                iFirstExtremeIndex = oParallelFaceIndexes.Item(i)
                                                iSecondExtremeIndex = oParallelFaceIndexes.Item(j)
                                            End If
                                        End If
                                    End If
                                End If
                            Next j
                        Next i
                        'Now compute extremes
                        oFirstExtremeFaceIndexes.Add(iFirstExtremeIndex)
                        oSecondExtremeFaceIndexes.Add(iSecondExtremeIndex)
                        oDistancesCollection.Add(0)
                        For i = 0 To oParallelFaceIndexes.Count - 1
                            If oParallelFaceIndexes.Item(i) <> iFirstExtremeIndex Then
                                oFaceCollection.Item(iFirstExtremeIndex).SelectFace(swSelectData, False)
                                oFaceCollection.Item(oParallelFaceIndexes.Item(i)).SelectFace(swSelectData, True)
                                If oMeasure.Calculate(Nothing) Then
                                    dDistance = oMeasure.NormalDistance
                                    If dDistance > 0 Then
                                        If IsTendsToZero(dDistance) Then
                                            If Not oFirstExtremeFaceIndexes.Contains(oParallelFaceIndexes.Item(i)) Then
                                                oFirstExtremeFaceIndexes.Add(oParallelFaceIndexes.Item(i))
                                            End If
                                        End If
                                        If IsTendsToZero(dDistance - dMaxDistance) Then
                                            If Not oSecondExtremeFaceIndexes.Contains(oParallelFaceIndexes.Item(i)) Then
                                                oSecondExtremeFaceIndexes.Add(oParallelFaceIndexes.Item(i))
                                            End If
                                        End If
                                        If Not DoesThisDistanceExistInThisDistanceCollection(oDistancesCollection, dDistance) Then
                                            oDistancesCollection.Add(dDistance)
                                        End If
                                    End If
                                End If
                            End If
                        Next i
                    End If
                End If
                oMeasure = Nothing
            End If
        End If
    End Sub
    Private Function DoesThisDistanceExistInThisDistanceCollection(ByRef oCollection As List(Of Double), ByRef dDistance As Double) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oCollection Is Nothing Then
            Dim i As Integer = 0
            While i < oCollection.Count And bReturn = False
                If IsTendsToZero(oCollection.Item(i) - dDistance) Then
                    bReturn = True
                End If
                i = i + 1
            End While
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Private Sub GetAllPlanarFacesWhichAreNormalToThisDirection(ByRef oNormal As Object, ByRef oNormalFaceIndexesCollection As List(Of Integer))
        On Error Resume Next
        If Not oNormal Is Nothing And Not oNormalFaceIndexesCollection Is Nothing Then
            Dim i As Integer = 0
            For i = 0 To oFaceCollection.Count - 1
                If oFaceCollection.Item(i).IsFaceNormalParallelToThisNormal(oNormal) Then
                    oNormalFaceIndexesCollection.Add(i)
                End If
            Next i
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Function GetWallThicknessForRoundProfile(ByRef strCrossSection As String) As Double
        On Error Resume Next
        Dim dReturn As Double = 0
        If iIndexOfRoundFace <> -1 Then
            'this means that the outer round face has been identified..
            strCrossSection = 2 * oFaceCollection.Item(iIndexOfRoundFace).Radius
            Dim i As Integer = 0
            For i = 0 To oFaceCollection.Count - 1
                If i <> iIndexOfRoundFace Then
                    If oFaceCollection.Item(i).IsRound Then
                        If oFaceCollection.Item(iIndexOfRoundFace).IsFaceAxisParallelToThisFace(oFaceCollection.Item(i)) Then
                            Dim dWThick As Double = Math.Abs(oFaceCollection.Item(iIndexOfRoundFace).Radius - oFaceCollection.Item(i).Radius)
                            If dReturn <> 0 Then
                                dReturn = Math.Min(dReturn, dWThick)
                            Else
                                dReturn = dWThick
                            End If
                        End If
                    End If
                End If
            Next i
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return dReturn
    End Function
    Private Sub DefineHolesForNonRoundProfiles(ByRef oFaceIndexesToBeConsidered As List(Of Integer))
        On Error Resume Next
        If Not oFaceCollection Is Nothing And Not oFaceIndexesToBeConsidered Is Nothing And Not oHoles Is Nothing Then
            For i As Integer = 0 To oFaceIndexesToBeConsidered.Count - 1
                oHoles.AddAllLoopsOfThisLoopCollection(oFaceCollection.Item(oFaceIndexesToBeConsidered.Item(i)).GetHoles())
            Next i
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Function GetEdgesForCutLengthOtherThanHolesAndComputeMaterialLengthForRoundProfile(ByRef dMaterialLength As Double) As List(Of sldworks.Edge)
        On Error Resume Next
        Dim oReturn As New List(Of sldworks.Edge)
        If iIndexOfRoundFace <> -1 Then
            oReturn = oFaceCollection.Item(iIndexOfRoundFace).GetEdgesForCutLengthOtherThanHolesForRoundProfile()
            dMaterialLength = oFaceCollection.Item(iIndexOfRoundFace).GetMaterialLength(dStartPointOfMaterialLength, dEndPointOfMaterialLength)
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return oReturn
    End Function
    Private Function GetEdgesForCutLengthOtherThanHolesAndComputeMaterialLengthForOtherProfiles(ByRef oFaceIndexesToBeConsidered As List(Of Integer), ByRef eShape As EnumShape) As List(Of sldworks.Edge)
        On Error Resume Next
        Dim oReturn As New List(Of sldworks.Edge)
        If Not oFaceIndexesToBeConsidered Is Nothing And Not oFaceCollection Is Nothing Then
            Dim oFacesForCutLength As New CFaceCollection
            For i As Integer = 0 To oFaceIndexesToBeConsidered.Count - 1
                oFacesForCutLength.Add(oFaceCollection.Item(oFaceIndexesToBeConsidered.Item(i)).GetFace())
            Next i
            For i As Integer = 0 To oFaceIndexesToBeConsidered.Count - 1
                Dim oEdges As List(Of sldworks.Edge) = oFaceCollection.Item(oFaceIndexesToBeConsidered.Item(i)).GetEdgesForCutLengthOtherThanHolesForOtherThanRoundProfile(oFacesForCutLength, eShape)
                If Not oEdges Is Nothing Then
                    For j As Integer = 0 To oEdges.Count - 1
                        If DoesThisEdgeExistInThisList(oReturn, oEdges.Item(j)) = False Then
                            oReturn.Add(oEdges.Item(j))
                        End If
                    Next j
                End If
                oEdges = Nothing
            Next i
            oFacesForCutLength = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return oReturn
    End Function
End Class
