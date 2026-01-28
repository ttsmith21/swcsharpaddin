Attribute VB_Name = "sheetmetal1"
'----------------------------------------------

'

' Preconditions: Sheet metal part is open.

'

' Postconditions: None

'

'----------------------------------------------

 

Option Explicit

 

Sub Process_CustomBendAllowance(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swFeat As SldWorks.Feature, swCustBend As SldWorks.CustomBendAllowance, kFacErrFlag As Boolean)

    If swCustBend.Type = swBendAllowanceDirect Then
        Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
        Print #2, FileNameWithoutExtension(swModel.GetTitle) & "       " & swFeat.Name & " [" & swFeat.GetTypeName & "]" & "      --> BendAllowance    = " & Format(swCustBend.BendAllowance * 1000 / 25.41, "##.###") & " in"
        boolLogError = True
    ElseIf swCustBend.Type = swBendAllowanceDeduction Then
        Print #2, FileNameWithoutExtension(swModel.GetTitle) & "       " & swFeat.Name & " [" & swFeat.GetTypeName & "]" & "      --> BendDeduction    = " & Format(swCustBend.BendDeduction * 1000 / 25.41, "##.###") & " in"
        boolLogError = True
    ElseIf swCustBend.Type = swBendAllowanceBendTable Then
        'Debug.Print "      BendTableFile    = " & swCustBend.BendTableFile
    ElseIf swCustBend.Type = swBendAllowanceKFactor Then
        If kFacErrFlag Then
            Print #2, FileNameWithoutExtension(swModel.GetTitle) & "       " & swFeat.Name & " [" & swFeat.GetTypeName & "]" & "      --> KFactor          = " & swCustBend.KFactor
            boolLogError = True
        End If
    Else
        'Debug.Print "      Type             = " & swCustBend.Type
    End If
    

End Sub

 

Sub Process_SMBaseFlange(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swFeat As SldWorks.Feature, kFacErrFlag As Boolean)

    'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"

    Dim swBaseFlange                As SldWorks.BaseFlangeFeatureData

    

    Set swBaseFlange = swFeat.GetDefinition
    
    

    'Debug.Print "    BendRadius = " & swBaseFlange.BendRadius * 1000# & " mm"

End Sub

 

Sub Process_SheetMetal(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swFeat As SldWorks.Feature, kFacErrFlag As Boolean)

    'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"

    Dim swCustBend As SldWorks.CustomBendAllowance

    Dim swSheetMetal  As SldWorks.SheetMetalFeatureData
    Dim value As CustomBendAllowance
    
    
    Set swSheetMetal = swFeat.GetDefinition

    Set swCustBend = swSheetMetal.GetCustomBendAllowance
    Set value = swSheetMetal.GetCustomBendAllowance()
    'Debug.Print "    BendRadius = " & swSheetMetal.BendRadius * 1000# & " mm"
    
    If swSheetMetal.AutoReliefType <> swSheetMetalReliefObround And swSheetMetal.ReliefRatio <> 0.5 And swSheetMetal.UseAutoRelief <> False Then
        If swSheetMetal.AutoReliefType = swSheetMetalReliefRectangular Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = Rectangular"
        ElseIf swSheetMetal.AutoReliefType = swSheetMetalReliefTear Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = Tear"
        ElseIf swSheetMetal.AutoReliefType = swSheetMetalReliefNone Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = None"
        ElseIf swSheetMetal.AutoReliefType = swSheetMetalReliefTearBend Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = Tear Bend"
        End If
            
        If swSheetMetal.ReliefRatio <> "0.5" Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Relief Ratio = " & swSheetMetal.ReliefRatio
        End If
        If swSheetMetal.UseAutoRelief <> False Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Use Custom Relief = " & swSheetMetal.UseAutoRelief
        End If
    End If
    
    Process_CustomBendAllowance swApp, swModel, swFeat, swCustBend, kFacErrFlag

End Sub

 

Sub Process_SM3dBend(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swFeat As SldWorks.Feature, kFacErrFlag As Boolean)

    'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"

    Dim swSketchBend As SldWorks.SketchedBendFeatureData

    Dim swCustBend As SldWorks.CustomBendAllowance

    If swSketchBend Is Nothing Then    '**** ???? ****
    Exit Sub
    End If

    Set swSketchBend = swFeat.GetDefinition

    Set swCustBend = swSketchBend.GetCustomBendAllowance

    'Debug.Print "    UseDefaultBendAllowance = " & swSketchBend.UseDefaultBendAllowance

    'Debug.Print "    UseDefaultBendRadius = " & swSketchBend.UseDefaultBendRadius

    'Debug.Print "    BendRadius = " & swSketchBend.BendRadius * 1000# & " mm"

    'If swSketchBend.AutoReliefType <> swSheetMetalReliefObround And swSketchBend.ReliefRatio <> 0.5 And swSketchBend.UseAutoRelief <> False Then
        'If swSketchBend.AutoReliefType = swSheetMetalReliefRectangular Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = Rectangular"
        'ElseIf swSketchBend.AutoReliefType = swSheetMetalReliefTear Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = Tear"
        'ElseIf swSketchBend.AutoReliefType = swSheetMetalReliefNone Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = None"
        'ElseIf swSketchBend.AutoReliefType = swSheetMetalReliefTearBend Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = Tear Bend"
        'End If
        '
        'If swSketchBend.ReliefRatio <> "0.5" Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Relief Ratio = " & swSketchBend.ReliefRatio
        'End If
        'If swSketchBend.UseAutoRelief <> False Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Use Custom Relief = " & swSketchBend.UseAutoRelief
        'End If
    'End If

    Process_CustomBendAllowance swApp, swModel, swFeat, swCustBend, kFacErrFlag
End Sub

 

Sub Process_SMMiteredFlange(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swFeat As SldWorks.Feature)

    'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"

    Dim swMiterFlange As SldWorks.MiterFlangeFeatureData

    

    Set swMiterFlange = swFeat.GetDefinition

    If swMiterFlange.AutoReliefType <> swSheetMetalReliefObround And swMiterFlange.ReliefRatio <> 0.5 And swMiterFlange.UseAutoRelief <> False Then
        If swMiterFlange.AutoReliefType = swSheetMetalReliefRectangular Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = Rectangular"
        ElseIf swMiterFlange.AutoReliefType = swSheetMetalReliefTear Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = Tear"
        ElseIf swMiterFlange.AutoReliefType = swSheetMetalReliefNone Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = None"
        ElseIf swMiterFlange.AutoReliefType = swSheetMetalReliefTearBend Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = Tear Bend"
        End If
            
        If swMiterFlange.ReliefRatio <> "0.5" Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Relief Ratio = " & swSheetMetal.ReliefRatio
        End If
        If swMiterFlange.UseAutoRelief <> False Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Use Custom Relief = " & swSheetMetal.UseAutoRelief
        End If
    End If

    'Debug.Print "    UseDefaultBendAllowance = " & swMiterFlange.UseDefaultBendAllowance

    'Debug.Print "    UseDefaultBendRadius = " & swMiterFlange.UseDefaultBendRadius

    'Debug.Print "    BendRadius = " & swMiterFlange.BendRadius * 1000# & " mm"

End Sub

 

Sub Process_Bends(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swBends As SldWorks.BendsFeatureData, kFacErrFlag As Boolean)

    Dim swCustBend  As SldWorks.CustomBendAllowance

    

    Set swCustBend = swBends.GetCustomBendAllowance

    

    'Debug.Print "    BendRadius                 = " & swBends.BendRadius * 1000# & " mm"

    'Debug.Print "    UseDefaultBendAllowance    = " & swBends.UseDefaultBendAllowance

    'Debug.Print "    UseDefaultBendRadius       = " & swBends.UseDefaultBendRadius


    If swBends.AutoReliefType <> swSheetMetalReliefObround And swBends.ReliefRatio <> 0.5 And swBends.UseAutoRelief <> False Then
        If swBends.AutoReliefType = swSheetMetalReliefRectangular Then
           ' Print #2, "    ***Auto Relief Type = Rectangular"
        ElseIf swBends.AutoReliefType = swSheetMetalReliefTear Then
            'Print #2, "    ***Auto Relief Type = Tear"
        ElseIf swBends.AutoReliefType = swSheetMetalReliefNone Then
            'Print #2, "    ***Auto Relief Type = None"
        ElseIf swBends.AutoReliefType = swSheetMetalReliefTearBend Then
           ' Print #2, "    ***Auto Relief Type = Tear Bend"
        End If
            
        If swBends.ReliefRatio <> "0.5" Then
            'Print #2, "    ***Relief Ratio = " & swSheetMetal.ReliefRatio
        End If
        If swBends.UseAutoRelief <> False Then
            'Print #2, "    ***Use Custom Relief = " & swSheetMetal.UseAutoRelief
        End If
    End If

    Process_CustomBendAllowance swApp, swModel, swFeat, swCustBend, kFacErrFlag

End Sub

 

Sub Process_ProcessBends(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swFeat As SldWorks.Feature, kFacErrFlag As Boolean)

    'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"

    Dim swBends                     As SldWorks.BendsFeatureData

    

    Set swBends = swFeat.GetDefinition

    If swBends.AutoReliefType <> swSheetMetalReliefObround And swBends.ReliefRatio <> 0.5 And swBends.UseAutoRelief <> False Then
        If swBends.AutoReliefType = swSheetMetalReliefRectangular Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = Rectangular"
        ElseIf swBends.AutoReliefType = swSheetMetalReliefTear Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = Tear"
        ElseIf swBends.AutoReliefType = swSheetMetalReliefNone Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = None"
        ElseIf swBends.AutoReliefType = swSheetMetalReliefTearBend Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = Tear Bend"
        End If
            
        If swBends.ReliefRatio <> "0.5" Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Relief Ratio = " & swBends.ReliefRatio
        End If
        If swBends.UseAutoRelief <> False Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Use Custom Relief = " & swBends.UseAutoRelief
        End If
    End If

    Process_Bends swApp, swModel, swBends, kFacErrFlag

End Sub

 

Sub Process_FlattenBends(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swFeat As SldWorks.Feature, kFacErrFlag As Boolean)

    'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"

    

    Dim swBends                     As SldWorks.BendsFeatureData

    

    Set swBends = swFeat.GetDefinition

    If swBends.AutoReliefType <> swSheetMetalReliefObround And swBends.ReliefRatio <> 0.5 And swBends.UseAutoRelief <> False Then
        If swBends.AutoReliefType = swSheetMetalReliefRectangular Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = Rectangular"
        ElseIf swBends.AutoReliefType = swSheetMetalReliefTear Then
           ' Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = Tear"
        ElseIf swBends.AutoReliefType = swSheetMetalReliefNone Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = None"
        ElseIf swBends.AutoReliefType = swSheetMetalReliefTearBend Then
           ' Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
           ' Print #2, "    ***Auto Relief Type = Tear Bend"
        End If
            
        If swBends.ReliefRatio <> "0.5" Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Relief Ratio = " & swBends.ReliefRatio
        End If
        If swBends.UseAutoRelief <> False Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Use Custom Relief = " & swBends.UseAutoRelief
        End If
    End If

    Process_Bends swApp, swModel, swBends, kFacErrFlag

End Sub

 

Sub Process_EdgeFlange(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swFeat As SldWorks.Feature)

    'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"

    Dim swEdgeFlange                As SldWorks.EdgeFlangeFeatureData

    

    Set swEdgeFlange = swFeat.GetDefinition


    If swEdgeFlange.AutoReliefType <> swSheetMetalReliefObround And swEdgeFlange.ReliefRatio <> 0.5 Then 'And swEdgeFlange.UseAutoRelief <> False Then
        If swEdgeFlange.AutoReliefType = swSheetMetalReliefRectangular Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = Rectangular"
        ElseIf swEdgeFlange.AutoReliefType = swSheetMetalReliefTear Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = Tear"
        ElseIf swEdgeFlange.AutoReliefType = swSheetMetalReliefNone Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = None"
        ElseIf swEdgeFlange.AutoReliefType = swSheetMetalReliefTearBend Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Auto Relief Type = Tear Bend"
        End If
            
        If swEdgeFlange.ReliefRatio <> "0.5" Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    ***Relief Ratio = " & swEdgeFlange.ReliefRatio
        End If
        'If swEdgeFlange.UseAutoRelief <> False Then
            'Print #2, "    ***Use Custom Relief = " & swEdgeFlange.UseAutoRelief
        'End If
    End If
    'Debug.Print "    UseDefaultBendRadius = " & swEdgeFlange.UseDefaultBendRadius

    'Debug.Print "    BendRadius = " & swEdgeFlange.BendRadius * 1000# & " mm"

End Sub

 

Sub Process_FlatPattern(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swFeat As SldWorks.Feature)

    

    Dim swFlatPatt                  As SldWorks.FlatPatternFeatureData

    Set swFlatPatt = swFeat.GetDefinition

    If swFlatPatt.SimplifyBends <> False Then
        'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
        'Print #2, "    Fixed Simplify Bends was " & swFlatPatt.SimplifyBends
        swFlatPatt.AccessSelections swModel, Nothing
        swFlatPatt.SimplifyBends = False
        swFlatPatt.ReleaseSelectionAccess
    End If
    If swFlatPatt.MergeFace <> True Then
        'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
        'Print #2, "    Fixed MergeFaces was " & swFlatPatt.MergeFace
        swFlatPatt.AccessSelections swModel, Nothing
        swFlatPatt.MergeFace = True
        swFlatPatt.ReleaseSelectionAccess
    End If
    If swFlatPatt.CornerTreatment <> False Then
        'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
        'Print #2, "    Fixed Corner Treatment was " & swFlatPatt.CornerTreatment

        swFlatPatt.AccessSelections swModel, Nothing
        swFlatPatt.CornerTreatment = False
        swFlatPatt.ReleaseSelectionAccess
    End If
  

End Sub

 

Sub Process_Hem(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swFeat As SldWorks.Feature, kFacErrFlag As Boolean)

    'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"

    

    Dim swHem                       As SldWorks.HemFeatureData

    Dim swCustBend                  As SldWorks.CustomBendAllowance

    

    Set swHem = swFeat.GetDefinition

    Set swCustBend = swHem.GetCustomBendAllowance

    'Debug.Print "    UseDefaultBendAllowance = " & swHem.UseDefaultBendAllowance

    'Debug.Print "    Radius = " & swHem.Radius * 1000# & " mm"

    

    Process_CustomBendAllowance swApp, swModel, swFeat, swCustBend, kFacErrFlag

End Sub

 

Sub Process_Jog(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swFeat As SldWorks.Feature, kFacErrFlag As Boolean)

    'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"

    Dim swJog                       As SldWorks.JogFeatureData

    Dim swCustBend                  As SldWorks.CustomBendAllowance

    

    Set swJog = swFeat.GetDefinition

    Set swCustBend = swJog.GetCustomBendAllowance

    'Debug.Print "    UseDefaultBendAllowance = " & swJog.UseDefaultBendAllowance

    'Debug.Print "    UseDefaultBendRadius = " & swJog.UseDefaultBendRadius

    'Debug.Print "    BendRadius = " & swJog.BendRadius * 1000# & " mm"

    

    Process_CustomBendAllowance swApp, swModel, swFeat, swCustBend, kFacErrFlag

End Sub

 

Sub Process_LoftedBend(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swFeat As SldWorks.Feature)

    'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"

    

    Dim swLoftBend                  As SldWorks.LoftedBendsFeatureData

    

    Set swLoftBend = swFeat.GetDefinition

    

    'Debug.Print "    To do..."

End Sub

 

Sub Process_Rip(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swFeat As SldWorks.Feature)

    'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"

    

    Dim swRip                       As SldWorks.RipFeatureData

    

    Set swRip = swFeat.GetDefinition

    

    'Debug.Print "    To do..."

End Sub

 

Sub Process_CornerFeat(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swFeat As SldWorks.Feature)

    'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"

    

    Dim swCloseCorner               As SldWorks.ClosedCornerFeatureData

    

    Set swCloseCorner = swFeat.GetDefinition

    

    'Debug.Print "    To do..."

End Sub

 

Sub Process_OneBend(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swFeat As SldWorks.Feature, kFacErrFlag As Boolean)

    'Print #2, "    +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"

    

    Dim swOneBend  As SldWorks.OneBendFeatureData

    Dim swCustBend As SldWorks.CustomBendAllowance

    

    Set swOneBend = swFeat.GetDefinition

    Set swCustBend = swOneBend.GetCustomBendAllowance

    'Debug.Print "      UseDefaultBendAllowance = " & swOneBend.UseDefaultBendAllowance

    'Debug.Print "      UseDefaultBendRadius = " & swOneBend.UseDefaultBendRadius
    'If swOneBend.AutoReliefType <> swSheetMetalReliefObround And swOneBend.ReliefRatio <> 0.5 And swOneBend.UseAutoRelief <> False Then
        'If swOneBend.AutoReliefType = swSheetMetalReliefRectangular Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    Auto Relief Type = Rectangular"
        'ElseIf swOneBend.AutoReliefType = swSheetMetalReliefTear Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    Auto Relief Type = Tear"
        'ElseIf swOneBend.AutoReliefType = swSheetMetalReliefNone Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    Auto Relief Type = None"
        'ElseIf swOneBend.AutoReliefType = swSheetMetalReliefTearBend Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
           ' Print #2, "    Auto Relief Type = Tear Bend"
        'End If
            
        'If swOneBend.ReliefRatio <> "0.5" Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    Relief Ratio = " & swOneBend.ReliefRatio
        'End If
        'If swOneBend.UseAutoRelief <> False Then
            'Print #2, "  +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"
            'Print #2, "    Use Custom Relief = " & swOneBend.UseAutoRelief
        'End If
    'End If
    

    Process_CustomBendAllowance swApp, swModel, swFeat, swCustBend, kFacErrFlag

End Sub

 Sub Process_SketchBend(swApp As SldWorks.SldWorks, swModel As SldWorks.ModelDoc2, swFeat As SldWorks.Feature, kFacErrFlag As Boolean)

   ' Debug.Print "    +" & swFeat.Name & " [" & swFeat.GetTypeName & "]"

    Dim swSketchBend As SldWorks.OneBendFeatureData
    Dim swCustBend As SldWorks.CustomBendAllowance
    Set swSketchBend = swFeat.GetDefinition
    Set swCustBend = swSketchBend.GetCustomBendAllowance

    Process_CustomBendAllowance swApp, swModel, swFeat, swCustBend, kFacErrFlag

End Sub

Sub CheckBends(swModel As ModelDoc2, kFacErrFlag As Boolean)

    Dim swApp                       As SldWorks.SldWorks

    Dim swSelMgr                    As SldWorks.SelectionMgr

    Dim swFeat                      As SldWorks.Feature

    Dim swSubFeat                   As SldWorks.Feature

    Dim bRet                        As Boolean

    Set swSelMgr = swModel.SelectionManager

    Set swFeat = swModel.FirstFeature

    

    'Print #2, "File = " & swModel.GetPathName

    Do While Not swFeat Is Nothing

        ' Process top-level sheet metal features

        Select Case swFeat.GetTypeName

            Case "SMBaseFlange"

                Process_SMBaseFlange swApp, swModel, swFeat, kFacErrFlag

                

            Case "SheetMetal"

                Process_SheetMetal swApp, swModel, swFeat, kFacErrFlag

                

            Case "SM3dBend"

                Process_SM3dBend swApp, swModel, swFeat, kFacErrFlag

                

            Case "SMMiteredFlange"

                Process_SMMiteredFlange swApp, swModel, swFeat

                

            Case "ProcessBends"

                'Process_ProcessBends swApp, swModel, swFeat, kFacErrFlag

                

            Case "FlattenBends"

                'Process_FlattenBends swApp, swModel, swFeat, kFacErrFlag

                

            Case "EdgeFlange"

                Process_EdgeFlange swApp, swModel, swFeat

                

            Case "FlatPattern"

                Process_FlatPattern swApp, swModel, swFeat

                

            Case "Hem"

                Process_Hem swApp, swModel, swFeat, kFacErrFlag

                

            Case "Jog"

                Process_Jog swApp, swModel, swFeat, kFacErrFlag

                

            Case "LoftedBend"

                Process_LoftedBend swApp, swModel, swFeat

                

            Case "Rip"

                Process_Rip swApp, swModel, swFeat

                

            Case "CornerFeat"

                Process_CornerFeat swApp, swModel, swFeat

            

            Case Else

                ' Probably not a sheet metal feature

        End Select

        Debug.Print swFeat.GetTypeName

        ' process sheet metal sub-features

        Set swSubFeat = swFeat.GetFirstSubFeature

        Do While Not swSubFeat Is Nothing

            Select Case swSubFeat.GetTypeName

                Case "OneBend"

                    Process_OneBend swApp, swModel, swSubFeat, kFacErrFlag

                Case "SketchBend"

                    Process_SketchBend swApp, swModel, swSubFeat, kFacErrFlag

                Case Else
                    Debug.Print swSubFeat.GetTypeName
                    ' Probably not a sheet metal feature

            End Select

            Set swSubFeat = swSubFeat.GetNextSubFeature()

        Loop

        Set swFeat = swFeat.GetNextFeature

    Loop

End Sub
