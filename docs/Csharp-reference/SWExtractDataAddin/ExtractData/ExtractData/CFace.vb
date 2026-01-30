Option Explicit On
Imports SolidWorks.Interop
Friend Class CFace
    Private swFace As sldworks.Face2 = Nothing
    Private oSurface As sldworks.Surface = Nothing

    Private dArea As Double = 0
    Private bPlanar As Boolean = False
    Private oNormal As sldworks.MathVector = Nothing
    Private dNormal(2) As Double

    Private bRound As Boolean = False
    Private oAxis(2) As Double
    Private oOrigin(2) As Double
    Private dRadius As Double = 0

    Private oLoopsOnMe As CLoopCollection
    Private oHolesOnMe As CLoopCollection

    Public ReadOnly Property GetFace As sldworks.Face2
        Get
            Return swFace
        End Get
    End Property
    Public ReadOnly Property Area As Double
        Get
            Return dArea
        End Get
    End Property
    Public ReadOnly Property IsPlanar As Boolean
        Get
            Return bPlanar
        End Get
    End Property
    Public ReadOnly Property IsRound As Boolean
        Get
            Return bRound
        End Get
    End Property
    Public ReadOnly Property Axis As Object
        Get
            Return oAxis
        End Get
    End Property
    Public ReadOnly Property Normal As Object
        Get
            Return dNormal
        End Get
    End Property
    Public ReadOnly Property Edges As Object
        Get
            Return swFace.GetEdges()
        End Get
    End Property
    Public ReadOnly Property Radius As Double
        Get
            Return dRadius
        End Get
    End Property
    Public Sub New(ByRef oFace As sldworks.Face2)
        On Error Resume Next
        swFace = oFace
        Populate()
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Protected Overrides Sub Finalize()
        On Error Resume Next
        MyBase.Finalize()
        swFace = Nothing
        oSurface = Nothing
        oNormal = Nothing
        oLoopsOnMe = Nothing
        oHolesOnMe = Nothing
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Private Sub Populate()
        On Error Resume Next
        If Not swFace Is Nothing Then
            dArea = swFace.GetArea()
            oSurface = swFace.GetSurface()
            If Not oSurface Is Nothing Then
                If oSurface.IsPlane Then
                    bPlanar = True
                    Dim oNor As Object = swFace.Normal
                    If Not oNor Is Nothing Then
                        dNormal(0) = oNor(0)
                        dNormal(1) = oNor(1)
                        dNormal(2) = oNor(2)
                    End If
                    oNor = Nothing
                ElseIf oSurface.IsCylinder Then
                    bRound = True
                    Dim oParam As Object = oSurface.CylinderParams()
                    If Not oParam Is Nothing Then
                        oOrigin(0) = oParam(0)
                        oOrigin(1) = oParam(1)
                        oOrigin(2) = oParam(2)
                        oAxis(0) = oParam(3)
                        oAxis(1) = oParam(4)
                        oAxis(2) = oParam(5)
                        dRadius = oParam(6)
                    End If
                    oParam = Nothing
                End If
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Function IsSame(ByRef oFace As CFace) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not swFace Is Nothing And Not oFace Is Nothing Then
            bReturn = IsSame(oFace.swFace)
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Function IsSame(ByRef oFace As sldworks.Face2) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not swFace Is Nothing And Not oFace Is Nothing Then
            If SWApp.IsSame(swFace, oFace) = swconst.swSameAs_Status_e.swSameAs_Same Then
                bReturn = True
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Private Function IsFaceAxisParallelToThis(ByRef dAxis As Object) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not dAxis Is Nothing Then
            If (IsTendsToZero(dAxis(0) - oAxis(0)) And IsTendsToZero(dAxis(1) - oAxis(1)) And IsTendsToZero(dAxis(2) - oAxis(2))) Or
                (IsTendsToZero(dAxis(0) + oAxis(0)) And IsTendsToZero(dAxis(1) + oAxis(1)) And IsTendsToZero(dAxis(2) + oAxis(2))) Then
                bReturn = True
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Function IsFaceAxisParallelToThisFace(ByRef oFaceToTest As CFace) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oFaceToTest Is Nothing Then
            If oFaceToTest.IsRound And IsRound Then
                bReturn = IsFaceAxisParallelToThis(oFaceToTest.Axis)
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Function IsFaceNormalParallelToThisFace(ByRef oFaceToTest As CFace) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oFaceToTest Is Nothing Then
            If IsPlanar And oFaceToTest.IsPlanar Then
                bReturn = IsFaceNormalParallelToThisNormal(oFaceToTest.Normal)
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Function IsFaceNormalParallelToThisNormal(ByRef oFaceNormal As Object) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oFaceNormal Is Nothing Then
            If (IsTendsToZero(dNormal(0) - oFaceNormal(0)) And IsTendsToZero(dNormal(1) - oFaceNormal(1)) And IsTendsToZero(dNormal(2) - oFaceNormal(2))) Or
                (IsTendsToZero(dNormal(0) + oFaceNormal(0)) And IsTendsToZero(dNormal(1) + oFaceNormal(1)) And IsTendsToZero(dNormal(2) + oFaceNormal(2))) Then
                bReturn = True
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Function SelectFace(ByRef swSelectData As sldworks.SelectData, ByVal isAppend As Boolean) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not swFace Is Nothing Then
            Dim oEntity As sldworks.Entity = swFace
            bReturn = oEntity.Select4(isAppend, swSelectData)
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Function GetLoops() As CLoopCollection
        On Error Resume Next
        If oLoopsOnMe Is Nothing Then
            oLoopsOnMe = New CLoopCollection(swFace)
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return oLoopsOnMe
    End Function
    Public Function GetHoles() As CLoopCollection
        On Error Resume Next
        If oHolesOnMe Is Nothing Then
            oHolesOnMe = New CLoopCollection()
            oLoopsOnMe = GetLoops()
            oLoopsOnMe.GetAllHoles(oHolesOnMe)
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return oHolesOnMe
    End Function
    Public Function GetAllEdgesOfAllOuterLoops() As List(Of sldworks.Edge)
        On Error Resume Next
        Dim oEdgesOfOuterLoops As New List(Of sldworks.Edge)
        If Not oEdgesOfOuterLoops Is Nothing Then
            oLoopsOnMe = GetLoops()
            Dim oOuterLoops As New CLoopCollection()
            oLoopsOnMe.GetAllOuterLoops(oOuterLoops)
            oEdgesOfOuterLoops = oOuterLoops.GetEdges()
            oOuterLoops = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return oEdgesOfOuterLoops
    End Function
    Public Function GetDirectionOfLargestLinearEdge(ByRef dStartPoint As Object, ByRef dEndPoint As Object, ByRef dMaterialLength As Double) As Object
        On Error Resume Next
        Dim oReturn(2) As Double
        If Not swFace Is Nothing Then
            Dim oEdges As Object = Edges
            If Not oEdges Is Nothing Then
                Dim dMaxLength As Double = 0
                Dim i As Integer = 0
                Dim oEdge As sldworks.Edge = Nothing
                Dim oCurve As sldworks.Curve = Nothing
                Dim dStart As Double = 0
                Dim dEnd As Double = 0
                Dim isClosed As Boolean = False
                Dim isPeriodic As Boolean = False
                Dim dLength As Double = 0
                For i = LBound(oEdges) To UBound(oEdges)
                    oEdge = oEdges(i)
                    If Not oEdge Is Nothing Then
                        oCurve = oEdge.GetCurve()
                        If Not oCurve Is Nothing Then
                            If oCurve.IsLine Then
                                If oCurve.GetEndParams(dStart, dEnd, isClosed, isPeriodic) Then
                                    'This can be measured.
                                    dLength = oCurve.GetLength3(dStart, dEnd)
                                    dMaxLength = Math.Max(dMaxLength, dLength)
                                    If IsTendsToZero(dMaxLength - dLength) Then
                                        Dim obj As Object = oCurve.LineParams()
                                        oReturn(0) = obj(3)
                                        oReturn(1) = obj(4)
                                        oReturn(2) = obj(5)
                                        Dim iNumber As Integer = 0
                                        Dim oStart As Object = oCurve.Evaluate2(dStart, iNumber)
                                        Dim oEnd As Object = oCurve.Evaluate2(dEnd, iNumber)
                                        If Not oStart Is Nothing And Not oEnd Is Nothing Then
                                            dStartPoint(0) = oStart(0)
                                            dStartPoint(1) = oStart(1)
                                            dStartPoint(2) = oStart(2)
                                            dEndPoint(0) = oEnd(0)
                                            dEndPoint(1) = oEnd(1)
                                            dEndPoint(2) = oEnd(2)
                                        End If
                                        oStart = Nothing
                                        oEnd = Nothing
                                    End If
                                End If
                            End If
                        End If
                        oCurve = Nothing
                    End If
                Next i
                dMaterialLength = dMaxLength
            End If
            oEdges = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return oReturn
    End Function
    Public Function GetEdgesForCutLengthOtherThanHolesForRoundProfile() As List(Of sldworks.Edge)
        On Error Resume Next
        Dim oReturn As New List(Of sldworks.Edge)
        If Not swFace Is Nothing Then
            If IsRound Then
                oReturn = GetAllEdgesOfAllOuterLoops()
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return oReturn
    End Function
    Public Function GetEdgesForCutLengthOtherThanHolesForOtherThanRoundProfile(ByRef oFacesForCutLength As CFaceCollection, ByRef eShape As EnumShape) As List(Of sldworks.Edge)
        On Error Resume Next
        Dim oReturn As New List(Of sldworks.Edge)
        If Not swFace Is Nothing And Not oFacesForCutLength Is Nothing Then
            If IsRound = False Then
                Dim oOuterEdges As List(Of sldworks.Edge) = GetAllEdgesOfAllOuterLoops()
                oReturn = GetAllEdgesWhoseAdjacentFacesArePlanarAndWhoseOtherAdjacentFaceIsNotInThisCollection(oOuterEdges, oFacesForCutLength, eShape)
                oOuterEdges = Nothing
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return oReturn
    End Function
    Private Function GetAllEdgesWhoseAdjacentFacesArePlanarAndWhoseOtherAdjacentFaceIsNotInThisCollection(ByRef oEdgesToEvaluate As List(Of sldworks.Edge), ByRef oFaces As CFaceCollection, ByRef eShape As EnumShape) As List(Of sldworks.Edge)
        On Error Resume Next
        Dim oReturn As New List(Of sldworks.Edge)
        If Not oReturn Is Nothing And Not oEdgesToEvaluate Is Nothing And Not oFaces Is Nothing Then
            Dim oEdge As sldworks.Edge = Nothing
            Dim oFace As sldworks.Face2 = Nothing
            For i As Integer = 0 To oEdgesToEvaluate.Count - 1
                oEdge = oEdgesToEvaluate(i)
                If Not oEdge Is Nothing Then
                    oFace = GetAdjacentFaceOfThisEdgeOtherThanMe(oEdge)
                    If Not oFace Is Nothing Then
                        Dim oCFace As New CFace(oFace)
                        If Not oCFace Is Nothing Then
                            If oCFace.IsPlanar Then
                                If oFaces.DoesThisFaceExist(oCFace) = False Then
                                    Dim oEdgeCollection As New List(Of sldworks.Edge)
                                    oCFace.GetEdgesWhichAreInLoopWithThisEdge(oEdge, oEdgeCollection)
                                    If Not oEdgeCollection Is Nothing Then
                                        If eShape = EnumShape.rectangle Or eShape = EnumShape.square Then
                                            For j As Integer = 0 To oEdgeCollection.Count - 1
                                                If DoesThisEdgeExistInThisList(oReturn, oEdgeCollection.Item(j)) = False Then
                                                    oReturn.Add(oEdgeCollection.Item(j))
                                                End If
                                            Next j
                                        Else
                                            For j As Integer = 0 To oEdgeCollection.Count - 1
                                                'Check if the edge's start and end point lie on faces which are in the list...
                                                If DoesStartAndEndPointsOfThisEdgeLieOnFacesFromThisFaceCollection(oEdgeCollection.Item(j), oFaces) Then
                                                    If DoesThisEdgeExistInThisList(oReturn, oEdgeCollection.Item(j)) = False Then
                                                        oReturn.Add(oEdgeCollection.Item(j))
                                                    End If
                                                End If
                                            Next j
                                        End If
                                    End If
                                    oEdgeCollection = Nothing
                                End If
                            End If
                        End If
                        oCFace = Nothing
                    End If
                    oFace = Nothing
                End If
            Next i
            oFace = Nothing
            oEdge = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return oReturn
    End Function
    Private Function GetAdjacentFaceOfThisEdgeOtherThanMe(ByRef oEdge As sldworks.Edge) As sldworks.Face2
        On Error Resume Next
        Dim oReturn As sldworks.Face2 = Nothing
        If Not oEdge Is Nothing Then
            Dim oFaces As Object = oEdge.GetTwoAdjacentFaces2()
            If Not oFaces Is Nothing Then
                If Not oFaces(0) Is Nothing And Not oFaces(1) Is Nothing Then
                    If SWApp.IsSame(oFaces(0), swFace) = swconst.swSameAs_Status_e.swSameAs_Same Then
                        oReturn = oFaces(1)
                    Else
                        oReturn = oFaces(0)
                    End If
                End If
            End If
            oFaces = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return oReturn
    End Function
    Public Sub GetEdgesWhichAreInLoopWithThisEdge(ByRef oEdge As sldworks.Edge, ByRef oEdgesCollection As List(Of sldworks.Edge))
        On Error Resume Next
        If Not oEdge Is Nothing And Not oEdgesCollection Is Nothing And Not swFace Is Nothing Then
            Dim oLoops As Object = swFace.GetLoops()
            If Not oLoops Is Nothing Then
                Dim i As Integer = 0
                Dim j As Integer = 0
                Dim oLoop As sldworks.Loop2 = Nothing
                Dim oEdgesColl As Object = Nothing
                Dim isFound As Boolean = False
                While i <= UBound(oLoops) And isFound = False
                    oLoop = oLoops(i)
                    If Not oLoop Is Nothing Then
                        oEdgesColl = oLoop.GetEdges()
                        If Not oEdgesColl Is Nothing Then
                            j = 0
                            While j <= UBound(oEdgesColl) And isFound = False
                                If SWApp.IsSame(oEdgesColl(j), oEdge) = swconst.swSameAs_Status_e.swSameAs_Same Then
                                    isFound = True
                                End If
                                j = j + 1
                            End While
                            If isFound Then
                                For j = 0 To UBound(oEdgesColl)
                                    oEdgesCollection.Add(oEdgesColl(j))
                                Next j
                            End If
                        End If
                    End If
                    i = i + 1
                End While
            End If
            oLoops = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Function GetMaterialLength(ByRef dStartPoint As Object, ByRef dEndPoint As Object, Optional ByRef oAxis As Object = Nothing) As Double
        On Error Resume Next
        Dim dReturn As Double = 0
        If Not swFace Is Nothing Then
            Dim oTessTriangles As Object = swFace.GetTessTriangles(True)
            Dim XMin As Double = MaxDouble
            Dim XMax As Double = MinDouble
            Dim YMin As Double = MaxDouble
            Dim YMax As Double = MinDouble
            Dim ZMin As Double = MaxDouble
            Dim ZMax As Double = MinDouble
            ProcessTessTriangles(oTessTriangles, XMax, XMin, YMax, YMin, ZMax, ZMin)
            If IsRound Then
                If IsTendsToZero(Math.Abs(Axis(0)) - 1) Then
                    dReturn = XMax - XMin
                ElseIf IsTendsToZero(Math.Abs(Axis(1)) - 1) Then
                    dReturn = YMax - YMin
                ElseIf IsTendsToZero(Math.Abs(Axis(2)) - 1) Then
                    dReturn = ZMax - ZMin
                Else
                    dReturn = Math.Abs(Axis(0)) * (XMax - XMin) + Math.Abs(Axis(1)) * (YMax - YMin) + Math.Abs(Axis(2)) * (ZMax - ZMin)
                End If
                dStartPoint(0) = XMin
                dStartPoint(1) = YMin
                dStartPoint(2) = ZMin
                dEndPoint(0) = XMax
                dEndPoint(1) = YMax
                dEndPoint(2) = ZMax
            Else
                If Not oAxis Is Nothing Then
                    If IsTendsToZero(Math.Abs(oAxis(0)) - 1) Then
                        dReturn = XMax - XMin
                    ElseIf IsTendsToZero(Math.Abs(oAxis(1)) - 1) Then
                        dReturn = YMax - YMin
                    ElseIf IsTendsToZero(Math.Abs(oAxis(2)) - 1) Then
                        dReturn = ZMax - ZMin
                    Else
                        dReturn = (Math.Abs(oAxis(0)) * (XMax - XMin) + Math.Abs(oAxis(1)) * (YMax - YMin) + Math.Abs(oAxis(2)) * (ZMax - ZMin)) / (Math.Pow(Math.Pow(oAxis(0), 2) + Math.Pow(oAxis(1), 2) + Math.Pow(oAxis(2), 2), 0.5))
                    End If
                    dStartPoint(0) = XMin
                    dStartPoint(1) = YMin
                    dStartPoint(2) = ZMin
                    dEndPoint(0) = XMax
                    dEndPoint(1) = YMax
                    dEndPoint(2) = ZMax
                End If
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return dReturn
    End Function
    Private Function GetMax(ByRef Val1 As Double, ByRef Val2 As Double, ByRef Val3 As Double, ByRef Val4 As Double) As Double
        Dim dReturn As Double = Val1
        dReturn = Math.Max(dReturn, Val2)
        dReturn = Math.Max(dReturn, Val3)
        dReturn = Math.Max(dReturn, Val4)
        Return dReturn
    End Function
    Private Function GetMin(ByRef Val1 As Double, ByRef Val2 As Double, ByRef Val3 As Double, ByRef Val4 As Double) As Double
        Dim dReturn As Double = Val1
        dReturn = Math.Min(dReturn, Val2)
        dReturn = Math.Min(dReturn, Val3)
        dReturn = Math.Min(dReturn, Val4)
        Return dReturn
    End Function
    Private Sub ProcessTessTriangles(ByRef vTessTriangles As Object, ByRef X_max As Double, ByRef X_min As Double, ByRef Y_max As Double, ByRef Y_min As Double, ByRef Z_max As Double, ByRef Z_min As Double)
        If Not vTessTriangles Is Nothing Then
            Dim i As Long = 0
            For i = 0 To UBound(vTessTriangles) / (1 * 9) - 1
                X_max = GetMax((vTessTriangles(9 * i + 0)), (vTessTriangles(9 * i + 3)), (vTessTriangles(9 * i + 6)), X_max)
                X_min = GetMin((vTessTriangles(9 * i + 0)), (vTessTriangles(9 * i + 3)), (vTessTriangles(9 * i + 6)), X_min)
                Y_max = GetMax((vTessTriangles(9 * i + 1)), (vTessTriangles(9 * i + 4)), (vTessTriangles(9 * i + 7)), Y_max)
                Y_min = GetMin((vTessTriangles(9 * i + 1)), (vTessTriangles(9 * i + 4)), (vTessTriangles(9 * i + 7)), Y_min)
                Z_max = GetMax((vTessTriangles(9 * i + 2)), (vTessTriangles(9 * i + 5)), (vTessTriangles(9 * i + 8)), Z_max)
                Z_min = GetMin((vTessTriangles(9 * i + 2)), (vTessTriangles(9 * i + 5)), (vTessTriangles(9 * i + 8)), Z_min)
            Next i
        End If
    End Sub
End Class
