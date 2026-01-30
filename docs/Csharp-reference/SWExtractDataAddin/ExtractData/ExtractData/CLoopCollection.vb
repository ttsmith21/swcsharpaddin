Option Explicit On
Imports SolidWorks.Interop
Friend Class CLoopCollection
    Private oCollection As List(Of CLoop) = Nothing
    Public Sub New()
        On Error Resume Next
        oCollection = New List(Of CLoop)
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub New(ByRef oFace As sldworks.Face2)
        On Error Resume Next
        oCollection = New List(Of CLoop)
        If Not oFace Is Nothing Then
            Dim oLoops As Object = oFace.GetLoops()
            If Not oLoops Is Nothing Then
                Dim oLoop As sldworks.Loop2 = Nothing
                For i As Integer = LBound(oLoops) To UBound(oLoops)
                    oLoop = oLoops(i)
                    oCollection.Add(New CLoop(oLoop))
                Next i
            End If
            oLoops = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public ReadOnly Property Count As Integer
        Get
            On Error Resume Next
            If Not oCollection Is Nothing Then
                Return oCollection.Count
            End If
            If Err.Number <> 0 Then Err.Clear()
            Return 0
        End Get
    End Property
    Public ReadOnly Property Item(ByVal index As Integer) As CLoop
        Get
            If index >= 0 And index < Count Then
                Return oCollection.Item(index)
            End If
            Return Nothing
        End Get
    End Property
    Private Sub Add(ByRef oLoop As sldworks.Loop2)
        On Error Resume Next
        If Not oLoop Is Nothing And Not oCollection Is Nothing Then
            oCollection.Add(New CLoop(oLoop))
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub AddAllLoopsOfThisLoopCollection(ByRef oLoopCollection As CLoopCollection)
        On Error Resume Next
        If Not oCollection Is Nothing And Not oLoopCollection Is Nothing Then
            Dim i As Integer = 0
            For i = 0 To oLoopCollection.Count - 1
                Add(oLoopCollection.Item(i).MyLoop())
            Next i
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Function IsAnyLoopOuterLoop() As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oCollection Is Nothing Then
            Dim i As Integer = 0
            While i < oCollection.Count And bReturn = False
                If Not oCollection.Item(i).IsOuter() Then
                    bReturn = True
                End If
                i = i + 1
            End While
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Function GetAllOuterLoops(ByRef oOuterLoops As CLoopCollection) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oOuterLoops Is Nothing And Not oCollection Is Nothing Then
            For i As Integer = 0 To oCollection.Count - 1
                If oCollection.Item(i).IsOuter() = True Then
                    oOuterLoops.Add(oCollection.Item(i).MyLoop())
                End If
            Next i
            bReturn = True
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Function GetAllHoles(ByRef oInnerLoops As CLoopCollection) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oInnerLoops Is Nothing And Not oCollection Is Nothing Then
            For i As Integer = 0 To oCollection.Count - 1
                If oCollection.Item(i).IsOuter() = False Then
                    oInnerLoops.Add(oCollection.Item(i).MyLoop())
                End If
            Next i
            bReturn = True
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Function SelectAllLoops(ByRef swSelectData As sldworks.SelectData) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not oCollection Is Nothing And Not swSelectData Is Nothing Then
            For i As Integer = 0 To oCollection.Count - 1
                bReturn = oCollection.Item(i).SelectEdgesOfLoop(swSelectData)
            Next i
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Sub Clear()
        On Error Resume Next
        If Not oCollection Is Nothing Then
            oCollection.Clear()
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Function GetEdges() As List(Of sldworks.Edge)
        On Error Resume Next
        Dim oReturn As New List(Of sldworks.Edge)
        If Not oCollection Is Nothing Then
            For i As Integer = 0 To oCollection.Count - 1
                Dim oEdges As Object = oCollection.Item(i).GetEdges()
                If Not oEdges Is Nothing Then
                    For j As Integer = LBound(oEdges) To UBound(oEdges)
                        oReturn.Add(oEdges(j))
                    Next j
                End If
                oEdges = Nothing
            Next i
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return oReturn
    End Function
    Protected Overrides Sub Finalize()
        On Error Resume Next
        MyBase.Finalize()
        Clear()
        oCollection = Nothing
        If Err.Number <> 0 Then Err.Clear()
    End Sub
End Class
