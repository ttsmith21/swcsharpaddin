Option Explicit On
Imports SolidWorks.Interop
Friend Class CLoop
    Private swLoop As sldworks.Loop2 = Nothing
    Public Sub New(ByRef oLoop As sldworks.Loop2)
        On Error Resume Next
        swLoop = oLoop
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Sub New(ByRef oLoop As CLoop)
        On Error Resume Next
        If Not oLoop Is Nothing Then
            swLoop = oLoop.swLoop
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Protected Overrides Sub Finalize()
        On Error Resume Next
        MyBase.Finalize()
        swLoop = Nothing
        If Err.Number <> 0 Then Err.Clear()
    End Sub
    Public Function IsSame(ByRef oLoop As CLoop) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not swLoop Is Nothing And Not oLoop Is Nothing Then
            bReturn = IsSame(oLoop.swLoop)
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public Function IsSame(ByRef oLoop As sldworks.Loop2) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not swLoop Is Nothing And Not oLoop Is Nothing Then
            If SWApp.IsSame(swLoop, oLoop) = swconst.swSameAs_Status_e.swSameAs_Same Then
                bReturn = True
            End If
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
    Public ReadOnly Property MyLoop As sldworks.Loop2
        Get
            Return swLoop
        End Get
    End Property
    Public Function GetEdges() As Object
        If Not swLoop Is Nothing Then
            Return swLoop.GetEdges()
        End If
        Return Nothing
    End Function
    Public ReadOnly Property IsOuter As Boolean
        Get
            If Not swLoop Is Nothing Then
                Return swLoop.IsOuter()
            End If
            Return False
        End Get
    End Property
    Public Function SelectEdgesOfLoop(ByRef swSelectData As sldworks.SelectData) As Boolean
        On Error Resume Next
        Dim bReturn As Boolean = False
        If Not swLoop Is Nothing Then
            Dim oEdges As Object = swLoop.GetEdges()
            Dim oEdge As sldworks.Edge = Nothing
            Dim oEntity As sldworks.Entity = Nothing
            If Not oEdges Is Nothing Then
                For i As Integer = LBound(oEdges) To UBound(oEdges)
                    oEdge = oEdges(i)
                    If Not oEdge Is Nothing Then
                        oEntity = oEdge
                        If Not oEntity Is Nothing Then
                            bReturn = oEntity.Select4(True, swSelectData)
                        End If
                        oEntity = Nothing
                    End If
                    oEdge = Nothing
                Next i
            End If
            oEdges = Nothing
        End If
        If Err.Number <> 0 Then Err.Clear()
        Return bReturn
    End Function
End Class
