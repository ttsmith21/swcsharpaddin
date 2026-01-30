Option Explicit On
Imports SolidWorks.Interop

Public Class CCalloutHandler
    Implements swpublished.SwCalloutHandler
    Private Function OnStringValueChanged(ByVal pManipulator As Object, ByVal RowID As Integer, ByVal Text As String) As Boolean Implements swpublished.ISwCalloutHandler.OnStringValueChanged
        Debug.Print("Text: " & Text)
        Debug.Print("Row: " & RowID)
        Return True
    End Function
End Class
