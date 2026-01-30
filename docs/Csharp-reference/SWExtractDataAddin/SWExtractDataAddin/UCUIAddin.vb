Option Explicit On
Imports SolidWorks.Interop
Public Class UCUIAddin
    Friend oUCUI As ExtractData.UCUI
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        oUCUI = New ExtractData.UCUI
        ElementHost1.Child = oUCUI
    End Sub
    Public Sub SetSWAddinObject(ByRef oSWApp As sldworks.SldWorks)
        On Error Resume Next
        If Not oSWApp Is Nothing And Not oUCUI Is Nothing Then
            oUCUI.SetSWAddinObject(oSWApp, GetProductVersion())
        End If
        If Err.Number <> 0 Then Err.Clear()
    End Sub
End Class
