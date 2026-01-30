Imports SolidWorks.Interop.sldworks

Public Class ExternalStart
    Implements IExternalStart
    Sub ExternalStart(swModel As ModelDoc2) Implements IExternalStart.ExternalStart
        Dim swLib As ExtractData.IUCUI = New ExtractData.UCUI
        swLib.RunForAutoPilot(swModel)
    End Sub

    Sub ExportGeo(GeoName As String) Implements IExternalStart.ExportGeo
        Dim TopsWorks As New DPS.ToPsWorks.SolidWorksAddIn
        Dim strModel As String
        Dim intLength As Integer
        Dim intPosition As Integer

        TopsWorks.ExportToGeo(GeoName & ".geo")
    End Sub
End Class
