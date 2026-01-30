Option Explicit On
Imports SolidWorks.Interop
Imports System.Collections.Generic
Friend Module MObjects
    Public SWApp As sldworks.SldWorks = Nothing
    Public oSTEPFiles As CStepFileCollection = Nothing
    Public Function GetCutLengthFromEdgeCollection(ByRef oEdgeCollection As List(Of sldworks.Entity), ByRef swModel As sldworks.ModelDoc2, ByRef swSelectData As sldworks.SelectData) As Double
        On Error Resume Next
        Dim dCutLength As Double = 0
        If Not oEdgeCollection Is Nothing And Not swModel Is Nothing And Not swSelectData Is Nothing Then
            Dim i As Integer = 0
            swModel.ClearSelection2(True)
            For i = 0 To oEdgeCollection.Count - 1
                oEdgeCollection.Item(i).Select4(True, swSelectData)
            Next
            Dim oMeasure As sldworks.Measure = swModel.Extension.CreateMeasure()
            If Not oMeasure Is Nothing Then
                oMeasure.ArcOption = 0
                oMeasure.AngleDecimalPlaces = 8
                oMeasure.LengthDecimalPlaces = 8
                If oMeasure.Calculate(Nothing) Then
                    dCutLength = oMeasure.TotalLength()
                End If
            End If
            oMeasure = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return dCutLength
    End Function
    Public Function DoesThisEdgeExistInThisList(ByRef oEdges As List(Of sldworks.Edge), ByRef oEdge As sldworks.Edge) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oEdges Is Nothing And Not oEdge Is Nothing Then
            Dim i As Integer = 0
            While i < oEdges.Count And bReturn = False
                If SWApp.IsSame(oEdges(i), oEdge) = swconst.swSameAs_Status_e.swSameAs_Same Then
                    bReturn = True
                End If
                i = i + 1
            End While
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Function DoesAnyOfTheseFacesExistInThisList(ByRef oFaces As CFaceCollection, ByRef oFacesToCheck As Object) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oFaces Is Nothing And Not oFacesToCheck Is Nothing Then
            Dim i As Integer = 0
            While i <= UBound(oFacesToCheck) And bReturn = False
                If oFaces.DoesThisFaceExist(oFacesToCheck(i)) Then
                    bReturn = True
                End If
                i = i + 1
            End While
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Sub RemoveThisValueFromThisList(ByRef oListOfIndexes As List(Of Integer), ByRef iValue As Integer)
        On Error Resume Next
        If Not oListOfIndexes Is Nothing Then
            Dim i As Integer = 0
            For i = oListOfIndexes.Count - 1 To 0 Step -1
                If oListOfIndexes.Item(i) = iValue Then
                    oListOfIndexes.RemoveAt(i)
                End If
            Next i
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Function DoesStartAndEndPointsOfThisEdgeLieOnFacesFromThisFaceCollection(ByRef oEdge As sldworks.Edge, ByRef oFaceCollection As CFaceCollection) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oEdge Is Nothing And Not oFaceCollection Is Nothing Then
            Dim oStartVertex As sldworks.Vertex = oEdge.GetStartVertex()
            Dim oEndVertex As sldworks.Vertex = oEdge.GetEndVertex()
            If Not oStartVertex Is Nothing And Not oEndVertex Is Nothing Then
                Dim oAdjacentFacesStartPoint As Object = oStartVertex.GetAdjacentFaces()
                Dim oAdjacentFacesEndPoint As Object = oEndVertex.GetAdjacentFaces()
                If DoesAnyOfTheseFacesExistInThisList(oFaceCollection, oAdjacentFacesStartPoint) And DoesAnyOfTheseFacesExistInThisList(oFaceCollection, oAdjacentFacesEndPoint) Then
                    bReturn = True
                End If
                oAdjacentFacesStartPoint = Nothing
                oAdjacentFacesEndPoint = Nothing
            End If
            oStartVertex = Nothing
            oEndVertex = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Sub RemoveAllEdgesWhichAreParallelToThisDirection(ByRef oEdgeList As List(Of sldworks.Edge), ByRef oDirection As Object)
        On Error Resume Next
        If Not oEdgeList Is Nothing And Not oDirection Is Nothing Then
            Dim i As Integer = 0
            For i = oEdgeList.Count - 1 To 0 Step -1
                If AreTheseVectorsParallel(oDirection, GetDirectionOfThisEdge(oEdgeList.Item(i))) Then
                    oEdgeList.RemoveAt(i)
                End If
            Next i
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Function GetDirectionOfThisEdge(ByRef oEdge As sldworks.Edge) As Object
        On Error Resume Next
        Dim oReturn(2) As Double
        If Not oEdge Is Nothing Then
            Dim oCurve As sldworks.Curve = oEdge.GetCurve()
            If Not oCurve Is Nothing Then
                If oCurve.IsLine Then
                    Dim dStart As Double = 0
                    Dim dEnd As Double = 0
                    Dim isClosed As Boolean = False
                    Dim isPeriodic As Boolean = False
                    If oCurve.GetEndParams(dStart, dEnd, isClosed, isPeriodic) Then
                        Dim obj As Object = oCurve.LineParams()
                        oReturn(0) = obj(3)
                        oReturn(1) = obj(4)
                        oReturn(2) = obj(5)
                    End If
                End If
            End If
            oCurve = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return oReturn
    End Function
    Public Function AreTheseVectorsParallel(ByRef oVector1 As Object, ByRef oVector2 As Object) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oVector1 Is Nothing And Not oVector2 Is Nothing Then
            If (IsTendsToZero(oVector1(0) - oVector2(0)) And IsTendsToZero(oVector1(1) - oVector2(1)) And IsTendsToZero(oVector1(2) - oVector2(2))) Or
                (IsTendsToZero(oVector1(0) + oVector2(0)) And IsTendsToZero(oVector1(1) + oVector2(1)) And IsTendsToZero(oVector1(2) + oVector2(2))) Then
                bReturn = True
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Function SelectEdge(ByRef oEdge As sldworks.Edge, ByRef swSelectData As sldworks.SelectData) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oEdge Is Nothing And Not swSelectData Is Nothing Then
            Dim oEntity As sldworks.Entity = oEdge
            If Not oEntity Is Nothing Then
                bReturn = oEntity.Select4(True, swSelectData)
            End If
            oEntity = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Function CreateMathPointAt(ByRef X As Double, ByRef Y As Double, ByRef Z As Double) As sldworks.MathPoint
        On Error Resume Next
        Dim oReturn As sldworks.MathPoint = Nothing
        If Not SWApp Is Nothing Then
            Dim swMathUtility As sldworks.MathUtility = SWApp.GetMathUtility()
            If Not swMathUtility Is Nothing Then
                Dim dArr(2) As Double
                dArr(0) = X
                dArr(1) = Y
                dArr(2) = Z
                oReturn = swMathUtility.CreatePoint((dArr))
            End If
            swMathUtility = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return oReturn
    End Function
End Module
