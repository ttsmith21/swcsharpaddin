Attribute VB_Name = "DimensionDrawing"
Private Type ProcessedElement
    obj As Object
    X1 As Double
    X2 As Double
    Y1 As Double
    Y2 As Double
    Angle As Double
End Type

Private Type BendElement
    obj As SldWorks.SketchSegment
    Position As Double
    P2 As Double
    Angle As Double
End Type

Private Type EdgeElement
    obj As Object
    X As Double
    Y As Double
    Angle As Double
    Type As String
End Type

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Sub manualDim()
    Dim swApp As SldWorks.SldWorks
    Dim swDraw As SldWorks.DrawingDoc
    Dim swView As SldWorks.View

    Set swApp = Application.SldWorks
    Set swDraw = swApp.ActiveDoc
    Set swView = swDraw.GetFirstView
    Set swView = swView.GetNextView

    DimensionFlat swApp, swDraw, swView

End Sub

Sub DimensionFlat(swApp As SldWorks.SldWorks, swDraw As SldWorks.DrawingDoc, swView As SldWorks.View)

    Dim swModel As SldWorks.ModelDoc2
    Dim swModelDocExt As SldWorks.ModelDocExtension
    Dim CenterLoc As Double
    Dim swDimension As SldWorks.Dimension
    Dim X As Integer
    Dim Y As Integer
    Dim HorzBends() As BendElement
    Dim VertBends() As BendElement
    Dim TopPos As EdgeElement
    Dim BottomPos As EdgeElement
    Dim LeftPos As EdgeElement
    Dim RightPos As EdgeElement
    Dim RightView As SldWorks.View
    Dim TopView As SldWorks.View
    Dim vViewLines As Variant
    Dim TimeOut As Integer

    X = 0
    Y = 0

    Set swModel = swDraw

    origVXf = swView.GetXform

    FindBendLines swApp, swDraw, swView, HorzBends, VertBends

    TimeOut = 0
    While LeftPos.obj Is Nothing And TimeOut <> 25
        FindLeftPosLine swView, swDraw, swApp, LeftPos
        TimeOut = TimeOut + 1
    Wend

    TimeOut = 0
    While RightPos.obj Is Nothing And TimeOut <> 25
        FindRightPosLine swView, swDraw, swApp, RightPos
        TimeOut = TimeOut + 1
    Wend

    TimeOut = 0
    While TopPos.obj Is Nothing And TimeOut <> 25
    FindTopPosLine swView, swDraw, swApp, TopPos
        TimeOut = TimeOut + 1
    Wend

    TimeOut = 0
    While BottomPos.obj Is Nothing And TimeOut <> 25
        FindBottomPosLine swView, swDraw, swApp, BottomPos
        TimeOut = TimeOut + 1
    Wend

    If LeftPos.obj Is Nothing Or RightPos.obj Is Nothing Or TopPos.obj Is Nothing Or BottomPos.obj Is Nothing Then
        Position = swView.Position
        Position(0) = 0.1397
        Position(1) = 0.1215
        swView.Position = Position
        TimeOut = 0
        While swView.GetOutline(2) < 0.2476 And swView.GetOutline(3) < 0.1841 And TimeOut <> 25
            SheetProperties = swView.Sheet.GetProperties2()
            SheetProperties(2) = SheetProperties(2) * 1.05
            swView.Sheet.SetProperties2 SheetProperties(0), SheetProperties(1), SheetProperties(2), SheetProperties(3), SheetProperties(4), SheetProperties(5), SheetProperties(6), SheetProperties(7)
            TimeOut = TimeOut + 1
        Wend
        DimensionTube swApp, swDraw, swView
        Exit Sub
    End If

    swDraw.ClearSelection2 True

    swDraw.ActivateView swView.Name

    LeftPos.obj.Select4 True, Nothing
    RightPos.obj.Select4 True, Nothing
    swModel.AddHorizontalDimension2 (((RightPos.X - LeftPos.X) / 2) + LeftPos.X) * origVXf(2) / 39.3700787401575, (((BottomPos.Y * origVXf(2)) - 0.25) / 39.3700787401575), 0

    swDraw.ClearSelection2 True

    swDraw.ActivateView swView.Name
    TopPos.obj.Select4 True, Nothing
    BottomPos.obj.Select4 True, Nothing
    swModel.AddVerticalDimension2 (((LeftPos.X * origVXf(2)) - 0.25) / 39.3700787401575), (((TopPos.Y - BottomPos.Y) / 2) + BottomPos.Y) * origVXf(2) / 39.3700787401575, 0

    If Not VertBends(1).obj Is Nothing Then
        swDraw.ClearSelection2 True
        For i = LBound(VertBends) To UBound(VertBends) + 1
            If i = LBound(VertBends) Then
                LeftPos.obj.Select4 True, Nothing
                tempName = VertBends(i).obj.GetName & "@" & VertBends(i).obj.GetSketch.Name & "@" & swView.RootDrawingComponent.Name & "@" & swView.Name
                swModel.Extension.SelectByID2 tempName, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0
                CenterLoc = (((VertBends(i).Position - LeftPos.X) / 2) + LeftPos.X) * origVXf(2) / 39.3700787401575
            ElseIf i = UBound(VertBends) + 1 Then
                On Error Resume Next
                If VertBends(i - 1).Position <> VertBends(i - 2).Position Then
                    On Error GoTo 0
                    tempName = VertBends(i - 1).obj.GetName & "@" & VertBends(i - 1).obj.GetSketch.Name & "@" & swView.RootDrawingComponent.Name & "@" & swView.Name
                    swModel.Extension.SelectByID2 tempName, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0
                    RightPos.obj.Select4 True, Nothing
                    CenterLoc = (((RightPos.X - VertBends(i - 1).Position) / 2) + VertBends(i - 1).Position) * origVXf(2) / 39.3700787401575
                Else
                    tempPos = VertBends(i - 2).Position
                    tempP2 = VertBends(i - 2).P2
                    For k = LBound(VertBends) To UBound(VertBends)
                        If VertBends(k).Position = tempPos And VertBends(k).P2 <= tempP2 Then
                            swDraw.ClearSelection2 True
                            tempP2 = VertBends(k).P2
                            tempName = VertBends(k).obj.GetName & "@" & VertBends(k).obj.GetSketch.Name & "@" & swView.RootDrawingComponent.Name & "@" & swView.Name
                            swModel.Extension.SelectByID2 tempName, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0
                            RightPos.obj.Select4 True, Nothing
                            CenterLoc = (((RightPos.X - VertBends(k).Position) / 2) + VertBends(k).Position) * origVXf(2) / 39.3700787401575
                        End If
                    Next k
                End If
            Else
                If VertBends(i).Position <> VertBends(i - 1).Position And i <> (UBound(VertBends) / 2) + 1 Then
                    On Error Resume Next
                    If VertBends(i - 1).Position <> VertBends(i - 2).Position Then
                        On Error GoTo 0
                        tempName = VertBends(i - 1).obj.GetName & "@" & VertBends(i - 1).obj.GetSketch.Name & "@" & swView.RootDrawingComponent.Name & "@" & swView.Name
                        swModel.Extension.SelectByID2 tempName, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0
                        tempName = VertBends(i).obj.GetName & "@" & VertBends(i).obj.GetSketch.Name & "@" & swView.RootDrawingComponent.Name & "@" & swView.Name
                        swModel.Extension.SelectByID2 tempName, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0
                        CenterLoc = (((VertBends(i).Position - VertBends(i - 1).Position) / 2) + VertBends(i - 1).Position) * origVXf(2) / 39.3700787401575
                    Else
                        tempPos = VertBends(i - 2).Position
                        tempP2 = VertBends(i - 2).P2
                        For k = LBound(VertBends) To UBound(VertBends)
                            If VertBends(k).Position = tempPos And VertBends(k).P2 <= tempP2 Then
                                swDraw.ClearSelection2 True
                                tempP2 = VertBends(k).P2
                                tempName = VertBends(k).obj.GetName & "@" & VertBends(k).obj.GetSketch.Name & "@" & swView.RootDrawingComponent.Name & "@" & swView.Name
                                swModel.Extension.SelectByID2 tempName, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0
                                tempName = VertBends(i).obj.GetName & "@" & VertBends(i).obj.GetSketch.Name & "@" & swView.RootDrawingComponent.Name & "@" & swView.Name
                                swModel.Extension.SelectByID2 tempName, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0
                                CenterLoc = (((VertBends(i).Position - VertBends(k).Position) / 2) + VertBends(k).Position) * origVXf(2) / 39.3700787401575
                            End If
                        Next k
                    End If
                    On Error GoTo 0
                End If
            End If

            swModel.AddHorizontalDimension2 CenterLoc, (((BottomPos.Y * origVXf(2)) - 0.125) / 39.3700787401575), 0
            swModel.EditDimensionProperties2 0, 0, 0, "", "", True, 9, 2, True, 12, 12, "", " BL", True, "", "", False
            swDraw.ClearSelection2 True
            swDraw.EditRebuild
        Next i

        swDraw.ActivateView (swView.Name)
        swDraw.ViewZoomtofit2
        swDraw.Extension.SelectByID2 swView.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
        Set TopView = Nothing
        Set TopView = swDraw.CreateUnfoldedViewAt3(swView.Position(0), 0.3, 0, False)
        TopView.SetDisplayMode3 False, swHIDDEN, False, True
        swDraw.ClearSelection2 True
        swDraw.ActivateView (TopView.Name)
        swDraw.Extension.SelectByID2 TopView.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
        TopView.ReferencedConfiguration = "Default"
        TopView.AlignWithView swAlignViewVerticalCenter, swView
        swDraw.ForceRebuild
        DimensionOther swApp, swDraw, TopView

        If TopView.GetOutline(3) > 0.21 Then
            Position = TopView.Position
            Position(1) = Position(1) - (TopView.GetOutline(3) - 0.21)
            TopView.Position = Position
        End If

        If swView.GetOutline(3) > TopView.GetOutline(1) - 0.00635 Then
            Position = swView.Position
            ProjPosition = TopView.Position
            Position(1) = Position(1) - (swView.GetOutline(3) - (TopView.GetOutline(1) - 0.00635))
            swView.Position = Position
            TopView.Position = ProjPosition
        End If

        TopView.SetDisplayMode3 False, swHIDDEN_GREYED, False, True
    End If

    If Not HorzBends(1).obj Is Nothing Then
        swDraw.ClearSelection2 True
        For i = LBound(HorzBends) To UBound(HorzBends) + 1
            If i = LBound(HorzBends) Then
                BottomPos.obj.Select4 True, Nothing
                tempName = HorzBends(i).obj.GetName & "@" & HorzBends(i).obj.GetSketch.Name & "@" & swView.RootDrawingComponent.Name & "@" & swView.Name
                swModel.Extension.SelectByID2 tempName, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0
                CenterLoc = (((HorzBends(i).Position - BottomPos.Y) / 2) + BottomPos.Y) * origVXf(2) / 39.3700787401575
            ElseIf i = UBound(HorzBends) + 1 Then
                On Error Resume Next
                If HorzBends(i - 1).Position <> HorzBends(i - 2).Position Then
                    On Error GoTo 0
                    tempName = HorzBends(i - 1).obj.GetName & "@" & HorzBends(i - 1).obj.GetSketch.Name & "@" & swView.RootDrawingComponent.Name & "@" & swView.Name
                    swModel.Extension.SelectByID2 tempName, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0
                    TopPos.obj.Select4 True, Nothing
                    CenterLoc = (((TopPos.Y - HorzBends(i - 1).Position) / 2) + HorzBends(i - 1).Position) * origVXf(2) / 39.3700787401575
                Else
                    tempPos = HorzBends(i - 2).Position
                    tempP2 = HorzBends(i - 2).P2
                    For k = LBound(HorzBends) To UBound(HorzBends)
                        If HorzBends(k).Position = tempPos And HorzBends(k).P2 <= tempP2 Then
                            swDraw.ClearSelection2 True
                            tempP2 = HorzBends(k).P2
                            tempName = HorzBends(k).obj.GetName & "@" & HorzBends(k).obj.GetSketch.Name & "@" & swView.RootDrawingComponent.Name & "@" & swView.Name
                            swModel.Extension.SelectByID2 tempName, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0
                            TopPos.obj.Select4 True, Nothing
                            CenterLoc = (((TopPos.Y - HorzBends(k).Position) / 2) + HorzBends(k).Position) * origVXf(2) / 39.3700787401575
                        End If
                    Next k
                End If
            Else
                If HorzBends(i).Position <> HorzBends(i - 1).Position And i <> (UBound(HorzBends) / 2) + 1 Then
                    On Error Resume Next
                    If HorzBends(i - 1).Position <> HorzBends(i - 2).Position Then
                        On Error GoTo 0
                        tempName = HorzBends(i - 1).obj.GetName & "@" & HorzBends(i - 1).obj.GetSketch.Name & "@" & swView.RootDrawingComponent.Name & "@" & swView.Name
                        swModel.Extension.SelectByID2 tempName, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0
                        tempName = HorzBends(i).obj.GetName & "@" & HorzBends(i).obj.GetSketch.Name & "@" & swView.RootDrawingComponent.Name & "@" & swView.Name
                        swModel.Extension.SelectByID2 tempName, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0
                        CenterLoc = (((HorzBends(i).Position - HorzBends(i - 1).Position) / 2) + HorzBends(i - 1).Position) * origVXf(2) / 39.3700787401575
                    Else
                        tempPos = HorzBends(i - 2).Position
                        tempP2 = HorzBends(i - 2).P2
                        For k = LBound(HorzBends) To UBound(HorzBends)
                            If HorzBends(k).Position = tempPos And HorzBends(k).P2 <= tempP2 Then
                                swDraw.ClearSelection2 True
                                tempP2 = HorzBends(k).P2
                                tempName = HorzBends(k).obj.GetName & "@" & HorzBends(k).obj.GetSketch.Name & "@" & swView.RootDrawingComponent.Name & "@" & swView.Name
                                swModel.Extension.SelectByID2 tempName, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0
                                tempName = HorzBends(i).obj.GetName & "@" & HorzBends(i).obj.GetSketch.Name & "@" & swView.RootDrawingComponent.Name & "@" & swView.Name
                                swModel.Extension.SelectByID2 tempName, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0
                                CenterLoc = (((HorzBends(i).Position - HorzBends(k).Position) / 2) + HorzBends(k).Position) * origVXf(2) / 39.3700787401575
                            End If
                        Next k
                    End If
                    On Error GoTo 0
                End If
            End If

            swModel.AddVerticalDimension2 (((LeftPos.X * origVXf(2)) - 0.125) / 39.3700787401575), CenterLoc, 0
            swModel.EditDimensionProperties2 0, 0, 0, "", "", True, 9, 2, True, 12, 12, "", " BL", True, "", "", False
            swDraw.ClearSelection2 True
            swDraw.EditRebuild
        Next i

        swDraw.ActivateView (swView.Name)
        swDraw.Extension.SelectByID2 swView.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
        Set RightView = Nothing
        Set RightView = swDraw.CreateUnfoldedViewAt3(0.3, swView.Position(1), 0, False)
        RightView.SetDisplayMode3 False, swHIDDEN, False, True
        swDraw.ClearSelection2 True
        swDraw.ActivateView (RightView.Name)
        swDraw.Extension.SelectByID2 RightView.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
        RightView.ReferencedConfiguration = "Default"
        RightView.AlignWithView swAlignViewHorizontalCenter, swView
        swDraw.ForceRebuild
        swModel.GraphicsRedraw2
        DimensionOther swApp, swDraw, RightView

        If RightView.GetOutline(2) > 0.268 Then
            Position = RightView.Position
            Position(0) = Position(0) - (RightView.GetOutline(2) - 0.268)
            RightView.Position = Position
        End If

        If swView.GetOutline(2) > RightView.GetOutline(0) - 0.00635 Then
            Position = swView.Position
            ProjPosition = RightView.Position
            Position(0) = Position(0) - (swView.GetOutline(2) - (RightView.GetOutline(0) - 0.00635))
            swView.Position = Position
            RightView.Position = ProjPosition
        End If

        RightView.SetDisplayMode3 False, swHIDDEN_GREYED, False, True

    End If

    On Error GoTo 0

    If HorzBends(1).obj Is Nothing And VertBends(1).obj Is Nothing Then
        Position = swView.Position
        Position(0) = 0.1397
        Position(1) = 0.1215
        swView.Position = Position
        TimeOut = 0
        While swView.GetOutline(2) < 0.2476 And swView.GetOutline(3) < 0.1841 And TimeOut <> 25
            SheetProperties = swView.Sheet.GetProperties2()
            SheetProperties(2) = SheetProperties(2) * 1.05
            swView.Sheet.SetProperties2 SheetProperties(0), SheetProperties(1), SheetProperties(2), SheetProperties(3), SheetProperties(4), SheetProperties(5), SheetProperties(6), SheetProperties(7)
            TimeOut = TimeOut + 1
        Wend
    End If

    If HorzBends(1).obj Is Nothing And Not VertBends(1).obj Is Nothing Then
        Position = swView.Position
        Position(0) = 0.1397
        swView.Position = Position
        swDraw.EditRebuild
        TimeOut = 0
        While TopView.GetOutline(1) - swView.GetOutline(3) > 0.0254 And swView.GetOutline(2) < 0.2476 And swView.GetOutline(3) < 0.1841 And TimeOut <> 25
            SheetProperties = swView.Sheet.GetProperties2()
            SheetProperties(2) = SheetProperties(2) * 1.05
            swView.Sheet.SetProperties2 SheetProperties(0), SheetProperties(1), SheetProperties(2), SheetProperties(3), SheetProperties(4), SheetProperties(5), SheetProperties(6), SheetProperties(7)

            Position = swView.Position
            Position(1) = 0
            swView.Position = Position

            Position = swView.Position
            Position(1) = 0.0603 - swView.GetOutline(1)
            swView.Position = Position

            Position = TopView.Position
            Position(1) = Position(1) - (TopView.GetOutline(3) - 0.21)
            TopView.Position = Position

            TimeOut = TimeOut + 1
        Wend
    End If

    If Not HorzBends(1).obj Is Nothing And VertBends(1).obj Is Nothing Then
        Position = swView.Position
        Position(1) = 0.1215
        swView.Position = Position
        TimeOut = 0
        While RightView.GetOutline(0) - swView.GetOutline(2) > 0.0254 And swView.GetOutline(2) < 0.2476 And swView.GetOutline(3) < 0.1841 And TimeOut <> 25
            SheetProperties = swView.Sheet.GetProperties2()
            SheetProperties(2) = SheetProperties(2) * 1.05
            swView.Sheet.SetProperties2 SheetProperties(0), SheetProperties(1), SheetProperties(2), SheetProperties(3), SheetProperties(4), SheetProperties(5), SheetProperties(6), SheetProperties(7)

            Position = swView.Position
            Position(0) = 0
            swView.Position = Position

            Position = swView.Position
            Position(0) = 0.0445 - swView.GetOutline(0)
            swView.Position = Position

            Position = RightView.Position
            Position(0) = Position(0) - (RightView.GetOutline(2) - 0.268)
            RightView.Position = Position

            TimeOut = TimeOut + 1
        Wend
    End If

    If Not HorzBends(1).obj Is Nothing And Not VertBends(1).obj Is Nothing Then
        TimeOut = 0
        While TopView.GetOutline(1) - swView.GetOutline(3) > 0.0254 And RightView.GetOutline(0) - swView.GetOutline(2) > 0.0254 And TimeOut <> 25
            SheetProperties = swView.Sheet.GetProperties2()
            SheetProperties(2) = SheetProperties(2) * 1.05
            swView.Sheet.SetProperties2 SheetProperties(0), SheetProperties(1), SheetProperties(2), SheetProperties(3), SheetProperties(4), SheetProperties(5), SheetProperties(6), SheetProperties(7)

            Position = swView.Position
            Position(0) = 0
            Position(1) = 0
            swView.Position = Position

            Position = swView.Position
            Position(0) = 0.0445 - swView.GetOutline(0)
            Position(1) = 0.0603 - swView.GetOutline(1)
            swView.Position = Position

            Position = TopView.Position
            Position(1) = Position(1) - (TopView.GetOutline(3) - 0.21)
            TopView.Position = Position

            Position = RightView.Position
            Position(0) = Position(0) - (RightView.GetOutline(2) - 0.268)
            RightView.Position = Position

            TimeOut = TimeOut + 1
        Wend
    End If

    AlignDims.Align swModel
    swDraw.ClearSelection2 True
    swDraw.ForceRebuild

End Sub
Sub DimensionOther(swApp As SldWorks.SldWorks, swDraw As SldWorks.DrawingDoc, swView As SldWorks.View)

    Dim CenterLoc As Double
    Dim swDimension As SldWorks.Dimension
    Dim TopPos As EdgeElement
    Dim BottomPos As EdgeElement
    Dim LeftPos As EdgeElement
    Dim RightPos As EdgeElement
    Dim X As Integer
    Dim Y As Integer
    Dim swModel As ModelDoc2

    X = 0
    Y = 0

    Set swModel = swDraw

    origVXf = swView.GetXform

    FindLeftPosLine swView, swDraw, swApp, LeftPos
    FindRightPosLine swView, swDraw, swApp, RightPos
    FindTopPosLine swView, swDraw, swApp, TopPos
    FindBottomPosLine swView, swDraw, swApp, BottomPos

    swDraw.ClearSelection2 True
    On Error Resume Next
    LeftPos.obj.Select True
    RightPos.obj.Select True
    swModel.AddHorizontalDimension2 (((RightPos.X - LeftPos.X) / 2) + LeftPos.X) * origVXf(2) / 39.3700787401575, (((BottomPos.Y * origVXf(2)) - 0.5) / 39.3700787401575), 0

    swDraw.ClearSelection2 True

    TopPos.obj.Select True
    BottomPos.obj.Select True
    swModel.AddVerticalDimension2 (((LeftPos.X * origVXf(2)) - 0.5) / 39.3700787401575), (((TopPos.Y - BottomPos.Y) / 2) + BottomPos.Y) * origVXf(2) / 39.3700787401575, 0

    swDraw.ClearSelection2 True
    On Error GoTo 0

End Sub
Sub DimensionTube(swApp As SldWorks.SldWorks, swDraw As SldWorks.DrawingDoc, swView As SldWorks.View)

    Dim CenterLoc As Double
    Dim swDimension As SldWorks.DisplayDimension
    Dim TopPos As EdgeElement
    Dim BottomPos As EdgeElement
    Dim LeftPos As EdgeElement
    Dim RightPos As EdgeElement
    Dim X As Integer
    Dim Y As Integer
    Dim swModel As ModelDoc2

    X = 0
    Y = 0

    Set swModel = swDraw

    origVXf = swView.GetXform

    Outline = swView.GetOutline

    swDraw.ActivateView swView.Name
    LeftPos.X = Outline(0) * 39.3700787401575 / swView.GetXform(2)
    LeftPos.Y = (((Outline(3) - Outline(1)) / 2) + Outline(1)) * 39.3700787401575 / swView.GetXform(2)
    RightPos.X = Outline(2) * 39.3700787401575 / swView.GetXform(2)
    RightPos.Y = (((Outline(3) - Outline(1)) / 2) + Outline(1)) * 39.3700787401575 / swView.GetXform(2)
    TopPos.X = (((Outline(3) - Outline(0)) / 2) + Outline(0)) * 39.3700787401575 / swView.GetXform(2)
    TopPos.Y = Outline(3) * 39.3700787401575 / swView.GetXform(2)
    BottomPos.X = (((Outline(3) - Outline(0)) / 2) + Outline(0)) * 39.3700787401575 / swView.GetXform(2)
    BottomPos.Y = Outline(1) * 39.3700787401575 / swView.GetXform(2)

    swDraw.ClearSelection2 True
    b = swModel.Extension.SelectByRay(((LeftPos.X * 25.4) / 1000) * swView.GetXform(2), ((LeftPos.Y * 25.4) / 1000) * swView.GetXform(2), -500, 0, 0, -1, 0.005, 1, True, 0, 0)

    If Not swModel.SelectionManager.GetSelectedObjectCount2(-1) = 0 Then
        swDraw.SketchManager.SketchUseEdge3 False, False
        swApp.SetSelectionFilter swSelSKETCHSEGS, True
        swApp.SetApplySelectionFilter True
        swDraw.ActivateView swView.Name
        swDraw.Extension.SelectAll
        swApp.SetApplySelectionFilter False

        On Error Resume Next
        If swModel.SelectionManager.GetSelectedObject6(1, -1).GetType = swSketchARC Then
            swModel.EditUndo2 1
            swDraw.ClearSelection2 True
            swModel.Extension.SelectByRay ((LeftPos.X * 25.4) / 1000) * swView.GetXform(2), ((LeftPos.Y * 25.4) / 1000) * swView.GetXform(2), -500, 0, 0, -1, 0.005, 1, True, 0, 0
            swModel.AddDiameterDimension2 Outline(0), Outline(1), 0
            swDraw.ClearSelection2 True
        Else
            swModel.EditUndo2 1
            swDraw.ClearSelection2 True
            b = False
        End If
        On Error GoTo 0

    End If

    swDraw.ClearSelection2 True

    If b = False Then
        If RectProfile(LeftPos, RightPos, TopPos, BottomPos, swView, swModel, swApp) = False Then
            swModel.ViewZoomTo2 ((BottomPos.X * 25.4) / 1000) * swView.GetXform(2) + 0.001, ((BottomPos.Y * 25.4) / 1000) * swView.GetXform(2) + 0.001, 0, ((BottomPos.X * 25.4) / 1000) * swView.GetXform(2) - 0.001, ((BottomPos.Y * 25.4) / 1000) * swView.GetXform(2) - 0.001, 0
            b = swModel.Extension.SelectByRay(((BottomPos.X * 25.4) / 1000) * swView.GetXform(2), ((BottomPos.Y * 25.4) / 1000) * swView.GetXform(2), -500, 0, 0, -1, 0.005, 1, True, 0, 0)
            If b = True Then swModel.AddHorizontalDimension2 BottomPos.X, BottomPos.Y, 0
            swModel.ViewZoomtofit2
        End If
    End If
End Sub

Function RectProfile(LeftPos As EdgeElement, RightPos As EdgeElement, TopPos As EdgeElement, BottomPos As EdgeElement, swView As SldWorks.View, swModel As SldWorks.ModelDoc2, swApp As SldWorks.SldWorks) As Boolean

    Dim Dimension As Object
    RectProfile = False
    swModel.ViewZoomTo2 ((LeftPos.X * 25.4) / 1000) * swView.GetXform(2) + 0.001, ((LeftPos.Y * 25.4) / 1000) * swView.GetXform(2) + 0.001, 0, ((LeftPos.X * 25.4) / 1000) * swView.GetXform(2) - 0.001, ((LeftPos.Y * 25.4) / 1000) * swView.GetXform(2) - 0.001, 0
    b = swModel.Extension.SelectByRay(((LeftPos.X * 25.4) / 1000) * swView.GetXform(2), ((LeftPos.Y * 25.4) / 1000) * swView.GetXform(2), -500, 0, 0, -1, 0.005, 1, True, 0, 0)
    If b = False Then Exit Function
    b = swModel.Extension.SelectByRay(((RightPos.X * 25.4) / 1000) * swView.GetXform(2), ((RightPos.Y * 25.4) / 1000) * swView.GetXform(2), -500, 0, 0, -1, 0.005, 1, True, 0, 0)
    If b = False Then Exit Function
    Set Dimension = swModel.AddHorizontalDimension2((((RightPos.X - LeftPos.X) / 2) + LeftPos.X) * swView.GetXform(2) / 39.3700787401575, (((BottomPos.Y * swView.GetXform(2)) - 0.25) / 39.3700787401575), 0)
    If Dimension Is Nothing Then b = False
    If b = False Then Exit Function

    b = swModel.Extension.SelectByRay(((TopPos.X * 25.4) / 1000) * swView.GetXform(2), ((TopPos.Y * 25.4) / 1000) * swView.GetXform(2), -500, 0, 0, -1, 0.005, 1, True, 0, 0)
    If b = False Then Exit Function
    b = swModel.Extension.SelectByRay(((BottomPos.X * 25.4) / 1000) * swView.GetXform(2), ((BottomPos.Y * 25.4) / 1000) * swView.GetXform(2), -500, 0, 0, -1, 0.005, 1, True, 0, 0)
    If b = False Then Exit Function
    Set Dimension = swModel.AddVerticalDimension2((((LeftPos.X * swView.GetXform(2)) - 0.25) / 39.3700787401575), (((TopPos.Y - BottomPos.Y) / 2) + BottomPos.Y) * swView.GetXform(2) / 39.3700787401575, 0)
    If Dimension Is Nothing Then b = False
    If b = False Then Exit Function
    swModel.ViewZoomtofit2
    RectProfile = True
End Function
Sub FindBendLines(swApp As SldWorks.SldWorks, swDraw As SldWorks.DrawingDoc, swView As SldWorks.View, HorzBends() As BendElement, VertBends() As BendElement)

    Dim swBendLine As SldWorks.SketchLine
    Dim ProcessedLine As ProcessedElement
    Dim vBendLines As Variant
    ReDim HorzBends(1 To 1)
    ReDim VertBends(1 To 1)
    vBendLines = swView.GetBendLines

    If Not IsEmpty(vBendLines) Then

        For i = 0 To UBound(vBendLines)
            Set swBendLine = vBendLines(i)
            ProcessedLine = ProcessLine(swBendLine, swView, swDraw, swApp, True)
            If ProcessedLine.Angle = 0 Then
                X = X + 1
                ReDim Preserve HorzBends(1 To X)
                With HorzBends(X)
                    Set .obj = ProcessedLine.obj
                    .Position = ProcessedLine.Y1
                    If ProcessedLine.X1 < ProcessedLine.X2 Then
                        .P2 = ProcessedLine.X1
                    Else
                        .P2 = ProcessedLine.X2
                    End If
                End With
            ElseIf ProcessedLine.Angle = 90 Then
                Y = Y + 1
                ReDim Preserve VertBends(1 To Y)
                With VertBends(Y)
                    Set .obj = ProcessedLine.obj
                    .Position = ProcessedLine.X1
                    If ProcessedLine.Y1 < ProcessedLine.Y2 Then
                        .P2 = ProcessedLine.Y1
                    Else
                        .P2 = ProcessedLine.Y2
                    End If
                End With
            End If
        Next i

        If Not HorzBends(1).obj Is Nothing Then SortInfo HorzBends
        If Not VertBends(1).obj Is Nothing Then SortInfo VertBends

    End If
End Sub
Sub FindLeftPosLine(swView As SldWorks.View, swDraw As SldWorks.DrawingDoc, swApp As SldWorks.SldWorks, ByRef LeftPos As EdgeElement)
    Dim ProcessedLine As ProcessedElement
    Dim StartVertex As SldWorks.Vertex
    Dim EndVertex As SldWorks.Vertex
    Dim swLine As Object
    Dim vViewLines As Variant
    vViewLines = swView.GetPolylines7(1, Null)
    LeftPos.X = 9999
    For i = 0 To UBound(vViewLines)
        Set StartVertex = Nothing
        Set EndVertex = Nothing
        Set swLine = vViewLines(i)
        boolswLine = CBool(Not swLine Is Nothing)
        If boolswLine = True Then boolswLine = CBool(TypeOf swLine Is SldWorks.Edge Or TypeOf swLine Is SldWorks.SilhouetteEdge)
        If boolswLine = True Then boolswLine = CBool(Not swLine.GetCurve Is Nothing)
        If boolswLine = True Then
            If Not swLine.GetCurve.IsCircle = True And Not swLine.GetCurve.IsEllipse = True And Not swLine.GetCurve.IsBcurve = True Then
                ProcessedLine = ProcessLine(swLine, swView, swDraw, swApp, False, StartVertex, EndVertex)
                If ProcessedLine.X1 < LeftPos.X Then
                    If ProcessedLine.Angle = 90 Then
                        Set LeftPos.obj = ProcessedLine.obj
                        LeftPos.Type = "Line"
                        If ProcessedLine.Y1 < ProcessedLine.Y2 Then
                            LeftPos.Y = ProcessedLine.Y1
                        Else
                            LeftPos.Y = ProcessedLine.Y2
                        End If
                    Else
                        Set LeftPos.obj = StartVertex
                        LeftPos.Type = "Point"
                        LeftPos.Y = ProcessedLine.Y1
                    End If
                    LeftPos.X = ProcessedLine.X1

                ElseIf ProcessedLine.X1 = LeftPos.X And ProcessedLine.Angle = 90 Then
                    If ProcessedLine.Y1 < LeftPos.Y Or ProcessedLine.Y2 < LeftPos.Y Or LeftPos.Type <> "Line" Then
                        Set LeftPos.obj = ProcessedLine.obj
                        LeftPos.Type = "Line"
                        LeftPos.X = ProcessedLine.X1
                        If ProcessedLine.Y1 < ProcessedLine.Y2 Then
                            LeftPos.Y = ProcessedLine.Y1
                        Else
                            LeftPos.Y = ProcessedLine.Y2
                        End If
                    End If
                End If
                If ProcessedLine.X2 < LeftPos.X Then
                    Set LeftPos.obj = EndVertex
                    LeftPos.Type = "Point"
                    LeftPos.X = ProcessedLine.X2
                    LeftPos.Y = ProcessedLine.Y2
                End If
            End If
        End If
    Next i
End Sub
Sub FindRightPosLine(swView As SldWorks.View, swDraw As SldWorks.DrawingDoc, swApp As SldWorks.SldWorks, ByRef RightPos As EdgeElement)
    Dim ProcessedLine As ProcessedElement
    Dim StartVertex As SldWorks.Vertex
    Dim EndVertex As SldWorks.Vertex
    Dim swLine As Object
    Dim vViewLines As Variant
    vViewLines = swView.GetPolylines7(1, Null)
    RightPos.X = 0
    For i = 0 To UBound(vViewLines)
        Set StartVertex = Nothing
        Set EndVertex = Nothing
        Set swLine = vViewLines(i)
        boolswLine = CBool(Not swLine Is Nothing)
        If boolswLine = True Then boolswLine = CBool(TypeOf swLine Is SldWorks.Edge Or TypeOf swLine Is SldWorks.SilhouetteEdge)
        If boolswLine = True Then boolswLine = CBool(Not swLine.GetCurve Is Nothing)
        If boolswLine = True Then
            If Not swLine.GetCurve.IsCircle = True And Not swLine.GetCurve.IsEllipse = True And Not swLine.GetCurve.IsBcurve = True Then
                ProcessedLine = ProcessLine(swLine, swView, swDraw, swApp, False, StartVertex, EndVertex)
                If ProcessedLine.X1 > RightPos.X Then
                    If ProcessedLine.Angle = 90 Then
                        Set RightPos.obj = ProcessedLine.obj
                        RightPos.Type = "Line"
                        If ProcessedLine.Y1 < ProcessedLine.Y2 Then
                            RightPos.Y = ProcessedLine.Y1
                        Else
                            RightPos.Y = ProcessedLine.Y2
                        End If
                    Else
                        Set RightPos.obj = StartVertex
                        RightPos.Type = "Point"
                        RightPos.Y = ProcessedLine.Y1
                    End If
                    RightPos.X = ProcessedLine.X1
                ElseIf ProcessedLine.X1 = RightPos.X And ProcessedLine.Angle = 90 Then
                    If ProcessedLine.Y1 < RightPos.Y Or ProcessedLine.Y2 < RightPos.Y Or RightPos.Type <> "Line" Then
                        Set RightPos.obj = ProcessedLine.obj
                        RightPos.Type = "Line"
                        RightPos.X = ProcessedLine.X1
                        If ProcessedLine.Y1 < ProcessedLine.Y2 Then
                            RightPos.Y = ProcessedLine.Y1
                        Else
                            RightPos.Y = ProcessedLine.Y2
                        End If
                    End If
                End If
                If ProcessedLine.X2 > RightPos.X Then
                    Set RightPos.obj = EndVertex
                    RightPos.Type = "Point"
                    RightPos.X = ProcessedLine.X2
                    RightPos.Y = ProcessedLine.Y2
                End If
            End If
        End If
    Next i
End Sub
Sub FindTopPosLine(swView As SldWorks.View, swDraw As SldWorks.DrawingDoc, swApp As SldWorks.SldWorks, ByRef TopPos As EdgeElement)
    Dim ProcessedLine As ProcessedElement
    Dim StartVertex As SldWorks.Vertex
    Dim EndVertex As SldWorks.Vertex
    Dim swLine As Object
    Dim vViewLines As Variant
    vViewLines = swView.GetPolylines7(1, Null)
    TopPos.Y = 0
    For i = 0 To UBound(vViewLines)
        Set StartVertex = Nothing
        Set EndVertex = Nothing
        Set swLine = vViewLines(i)
        boolswLine = CBool(Not swLine Is Nothing)
        If boolswLine = True Then boolswLine = CBool(TypeOf swLine Is SldWorks.Edge Or TypeOf swLine Is SldWorks.SilhouetteEdge)
        If boolswLine = True Then boolswLine = CBool(Not swLine.GetCurve Is Nothing)
        If boolswLine = True Then
            If Not swLine.GetCurve.IsCircle = True And Not swLine.GetCurve.IsEllipse = True And Not swLine.GetCurve.IsBcurve = True Then
                ProcessedLine = ProcessLine(swLine, swView, swDraw, swApp, False, StartVertex, EndVertex)
                If ProcessedLine.Y1 > TopPos.Y Then
                    If ProcessedLine.Angle = 0 Then
                        Set TopPos.obj = ProcessedLine.obj
                        TopPos.Type = "Line"
                        If ProcessedLine.X1 < ProcessedLine.X2 Then
                            TopPos.X = ProcessedLine.X1
                        Else
                            TopPos.X = ProcessedLine.X2
                        End If
                    Else
                        Set TopPos.obj = StartVertex
                        TopPos.Type = "Point"
                        TopPos.X = ProcessedLine.X1
                    End If
                    TopPos.Y = ProcessedLine.Y1

                ElseIf ProcessedLine.Y1 = TopPos.Y And ProcessedLine.Angle = 0 Then
                    If ProcessedLine.X1 < TopPos.X Or ProcessedLine.X2 < TopPos.X Or TopPos.Type <> "Line" Then
                        Set TopPos.obj = ProcessedLine.obj
                        TopPos.Type = "Line"
                        TopPos.Y = ProcessedLine.Y1
                        If ProcessedLine.X1 < ProcessedLine.X2 Then
                            TopPos.X = ProcessedLine.X1
                        Else
                            TopPos.X = ProcessedLine.X2
                        End If
                    End If
                End If
                If ProcessedLine.Y2 > TopPos.Y Then
                    Set TopPos.obj = EndVertex
                    TopPos.Type = "Point"
                    TopPos.Y = ProcessedLine.Y2
                    TopPos.X = ProcessedLine.X2
                End If
            End If
        End If
    Next i
End Sub
Sub FindBottomPosLine(swView As SldWorks.View, swDraw As SldWorks.DrawingDoc, swApp As SldWorks.SldWorks, ByRef BottomPos As EdgeElement)
    Dim ProcessedLine As ProcessedElement
    Dim StartVertex As SldWorks.Vertex
    Dim EndVertex As SldWorks.Vertex
    Dim swLine As Object
    Dim vViewLines As Variant
    vViewLines = swView.GetPolylines7(1, Null)
    BottomPos.Y = 9999
    For i = 0 To UBound(vViewLines)
        Set StartVertex = Nothing
        Set EndVertex = Nothing
        Set swLine = vViewLines(i)
        boolswLine = CBool(Not swLine Is Nothing)
        If boolswLine = True Then boolswLine = CBool(TypeOf swLine Is SldWorks.Edge Or TypeOf swLine Is SldWorks.SilhouetteEdge)
        If boolswLine = True Then boolswLine = CBool(Not swLine.GetCurve Is Nothing)
        If boolswLine = True Then
            If Not swLine.GetCurve.IsCircle = True And Not swLine.GetCurve.IsEllipse = True And Not swLine.GetCurve.IsBcurve = True Then
                ProcessedLine = ProcessLine(swLine, swView, swDraw, swApp, False, StartVertex, EndVertex)
                If ProcessedLine.Y1 < BottomPos.Y Then
                    If ProcessedLine.Angle = 0 Then
                        Set BottomPos.obj = ProcessedLine.obj
                        BottomPos.Type = "Line"
                        If ProcessedLine.X1 < ProcessedLine.X2 Then
                            BottomPos.X = ProcessedLine.X1
                        Else
                            BottomPos.X = ProcessedLine.X2
                        End If
                    Else
                        Set BottomPos.obj = StartVertex
                        BottomPos.Type = "Point"
                        BottomPos.X = ProcessedLine.X1
                    End If
                    BottomPos.Y = ProcessedLine.Y1

                ElseIf ProcessedLine.Y1 = BottomPos.Y And ProcessedLine.Angle = 0 Then
                    If ProcessedLine.X1 < BottomPos.X Or ProcessedLine.X2 < BottomPos.X Or BottomPos.Type <> "Line" Then
                        Set BottomPos.obj = ProcessedLine.obj
                        BottomPos.Type = "Line"
                        BottomPos.Y = ProcessedLine.Y1
                        If ProcessedLine.X1 < ProcessedLine.X2 Then
                            BottomPos.X = ProcessedLine.X1
                        Else
                            BottomPos.X = ProcessedLine.X2
                        End If
                    End If

                End If
                If ProcessedLine.Y2 < BottomPos.Y Then
                    Set BottomPos.obj = EndVertex
                    BottomPos.Type = "Point"
                    BottomPos.Y = ProcessedLine.Y2
                    BottomPos.X = ProcessedLine.X2
                End If
            End If
        End If
    Next i
End Sub

Function ProcessLine(objLine As Object, swView As SldWorks.View, swDraw As SldWorks.DrawingDoc, swApp As SldWorks.SldWorks, Optional ByVal BendLines As Boolean, Optional ByRef StartVertex As SldWorks.Vertex, Optional ByRef EndVertex As SldWorks.Vertex) As ProcessedElement
    Dim swModelStartPt As SldWorks.MathPoint
    Dim swModelEndPt As SldWorks.MathPoint
    Dim swSketch As SldWorks.Sketch
    Dim swViewStartPt As SldWorks.MathPoint
    Dim swViewEndPt As SldWorks.MathPoint
    Dim pStart As SldWorks.SketchPoint
    Dim pEnd As SldWorks.SketchPoint

    On Error Resume Next
    Set swSketch = objLine.GetSketch
    On Error GoTo 0
    origVXf = swView.GetXform

    If BendLines = True Then
        Set pStart = objLine.GetStartPoint2
        Set pEnd = objLine.GetEndPoint2
        Set swModelStartPt = TransformSketchPointToModelSpace(swApp, swDraw, swSketch, pStart)
        Set swModelEndPt = TransformSketchPointToModelSpace(swApp, swDraw, swSketch, pEnd)
        Set swViewXform = swView.ModelToViewTransform
        Set swViewStartPt = swModelStartPt.MultiplyTransform(swViewXform)
        Set swViewEndPt = swModelEndPt.MultiplyTransform(swViewXform)
    ElseIf TypeOf objLine Is SldWorks.SilhouetteEdge Then
        Set swModelStartPt = objLine.GetStartPoint
        Set swModelEndPt = objLine.GetEndPoint
        Set swViewXform = swView.ModelToViewTransform
        Set swViewStartPt = swModelStartPt.MultiplyTransform(swViewXform)
        Set swViewEndPt = swModelEndPt.MultiplyTransform(swViewXform)
    ElseIf TypeOf objLine Is SldWorks.Edge Then
        Set swMathUtil = swApp.GetMathUtility
        Set StartVertex = objLine.GetStartVertex
        Set EndVertex = objLine.GetEndVertex
        Set swModelStartPt = swMathUtil.CreatePoint(StartVertex.GetPoint)
        Set swModelEndPt = swMathUtil.CreatePoint(EndVertex.GetPoint)
        Set swViewXform = swView.ModelToViewTransform
        Set swViewStartPt = swModelStartPt.MultiplyTransform(swViewXform)
        Set swViewEndPt = swModelEndPt.MultiplyTransform(swViewXform)
    Else
        If Not objLine.GetType = 3 Then
            Set pStart = objLine.GetStartPoint2
            Set pEnd = objLine.GetEndPoint2
            Set swViewStartPt = TransformSketchPointToModelSpace(swApp, swDraw, swSketch, pStart)
            Set swViewEndPt = TransformSketchPointToModelSpace(swApp, swDraw, swSketch, pEnd)
        Else
            Points = objLine.GetPoints2
            Set pStart = Points(LBound(Points))
            Set pEnd = Points(UBound(Points))
            Set swViewStartPt = TransformSketchPointToModelSpace(swApp, swDraw, swSketch, pStart)
            Set swViewEndPt = TransformSketchPointToModelSpace(swApp, swDraw, swSketch, pEnd)
        End If
    End If

    If Round((swViewEndPt.ArrayData(1) * 39.3700787401575) / origVXf(2), 3) - Round((swViewStartPt.ArrayData(1) * 39.3700787401575) / origVXf(2), 3) = 0 Then
        Angle = 0
    ElseIf Round((swViewEndPt.ArrayData(0) * 39.3700787401575) / origVXf(2), 3) - Round((swViewStartPt.ArrayData(0) * 39.3700787401575) / origVXf(2), 3) = 0 Then
        Angle = 90
    Else
        Angle = Round(Atn((swViewEndPt.ArrayData(1) - swViewStartPt.ArrayData(1)) / (swViewEndPt.ArrayData(0) - swViewStartPt.ArrayData(0))) * 180 / 3.14159265358979, 3)
    End If

    With ProcessLine
        Set .obj = objLine
        .Angle = Angle
        .X1 = Round((swViewStartPt.ArrayData(0) * 39.3700787401575) / origVXf(2), 3)
        .X2 = Round((swViewEndPt.ArrayData(0) * 39.3700787401575) / origVXf(2), 3)
        .Y1 = Round((swViewStartPt.ArrayData(1) * 39.3700787401575) / origVXf(2), 3)
        .Y2 = Round((swViewEndPt.ArrayData(1) * 39.3700787401575) / origVXf(2), 3)
    End With

End Function
Function TransformSketchPointToModelSpace(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swSketch As SldWorks.Sketch, swSkPt As SldWorks.SketchPoint) As SldWorks.MathPoint

    Dim swMathUtil              As SldWorks.MathUtility
    Dim swXform                 As SldWorks.MathTransform
    Dim nPt(2)                  As Double
    Dim vPt                     As Variant
    Dim swMathPt                As SldWorks.MathPoint

    nPt(0) = swSkPt.X:      nPt(1) = swSkPt.Y:      nPt(2) = swSkPt.Z
    vPt = nPt

    Set swMathUtil = swApp.GetMathUtility
    Set swXform = swSketch.ModelToSketchTransform
    Set swXform = swXform.Inverse
    Set swMathPt = swMathUtil.CreatePoint((vPt))
    Set swMathPt = swMathPt.MultiplyTransform(swXform)
    Set TransformSketchPointToModelSpace = swMathPt

End Function
Sub SortInfo(ByRef BendsToSort() As BendElement)
    Dim i As Long, j As Long

    For i = LBound(BendsToSort) To UBound(BendsToSort) - 1
        For j = i To UBound(BendsToSort)
            If BendsToSort(i).Position > BendsToSort(j).Position Then
                SwapInfo BendsToSort, i, j
            ElseIf BendsToSort(i).Position = BendsToSort(j).Position And BendsToSort(i).P2 > BendsToSort(j).P2 Then
                SwapInfo BendsToSort, i, j
            End If
        Next j
    Next i

End Sub
Sub SwapInfo(ByRef BendsToSort() As BendElement, ByVal lOne As Long, ByVal lTwo As Long)

    Dim tTemp As BendElement

    tTemp = BendsToSort(lOne)
    BendsToSort(lOne) = BendsToSort(lTwo)
    BendsToSort(lTwo) = tTemp

End Sub
