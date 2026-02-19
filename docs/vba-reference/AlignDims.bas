Attribute VB_Name = "AlignDims"
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swDraw As SldWorks.DrawingDoc
Dim swSheet As SldWorks.Sheet
Dim boolstatus As Boolean
Dim strError As String
Dim dimCollection As New Collection

Sub Align(swModel As SldWorks.ModelDoc2)
    Dim swViewAnnot As Variant
    Dim currAnn As SldWorks.Annotation
    Dim swViews As Variant
    
    On Error GoTo ErrorHandler
    strError = "Generic Error (macro)"
    
    Set swApp = Application.SldWorks
    If (swModel.GetType = SwConst.swDocumentTypes_e.swDocDRAWING) Then
        Set swDraw = swModel
        Set swSheet = swDraw.GetCurrentSheet
        swViews = swSheet.GetViews
        If UBound(swViews) > 0 Then
            For Each vView In swViews
                swViewAnnot = vView.GetAnnotations
                If Not IsEmpty(swViewAnnot) Then
                    For Each obj In swViewAnnot
                        Set currAnn = obj
                        dimCollection.Add currAnn
                        currAnn.Select3 True, Nothing
                    Next obj
                    boolstatus = swModel.Extension.AlignDimensions(swAlignDimensionType_e.swAlignDimensionType_AutoArrange, 0.001)
                End If
            Next vView
            swModel.ClearSelection2 True
            swModel.GraphicsRedraw2
        Else
            strError = "No views were found for this drawing sheet."
            GoTo ErrorHandler
        End If
        If (dimCollection.Count = 0) Then
            strError = "No dimensions were found on this drawing sheet."
            GoTo ErrorHandler
        End If
    Else
        strError = "Document type is not supported; macro works only with drawings"
        GoTo ErrorHandler
    End If
    GoTo swEnd
ErrorHandler:
swEnd:
End Sub
    
