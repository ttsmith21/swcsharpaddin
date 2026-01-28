Attribute VB_Name = "modMaterialCost"
Option Explicit
'********************************************************************************************
'  PROCEDURE:        MaterialUpdate
'
'  ABSTRACT:         Collection of macros for Northern Manufacturing to aid in material calcs and data export
'
'  NOTES:
'
'  SYSTEM CONCERNS:  None
'
'  INPUTS:           Varies
'
'  RETURNS:          Varies
'
'  PROGRAM GLOBALS:  Global constants listed below
'
'  HISTORY:
'
'  REVISION:      DATE               AUTHOR              INITIAL REASON
'     0.0      08 October 2009     3D Vision             Create date
'     0.1      09 October 2009     3D Vision/JSweeney    Updated from Tyson's beta testing
'     0.2      13 October 2009     3D Vision/JSweeney    Updated from Tyson's beta testing, routine name changed
'     0.3      21 October 2009     3D Vision/JSweeney    Added export module, fixed hanging Excel connection, a few other setting changes
'     0.4      23 October 2009     3D Vision/JSweeney    Added support for single files
'     0.5      13 November 2009    3D Vision/JSweeney    Removed the need to preselect the BOM in BOMExport. Fixed issue when instance is
'                                                        more than one character
'
'********************************************************************************************

'CalculateBendInfo constants
Public Const cdblRate1 As Double = 10 'seconds
Public Const cdblRate2 As Double = 30 'seconds
Public Const cdblRate3 As Double = 45 'seconds
Public Const cdblRate4 As Double = 200 'seconds
Public Const cdblRate5 As Double = 400 'seconds
Public Const cdblSetupRate As Double = 1.25 'minutes per foot for break setup
Public Const cdblBreakSetup As Double = 10 'brake setup constant
Public Const cdblRate3Weight As Double = 100 'max weight for rate 3
Public Const cdblRate2Weight As Double = 40 'max weight for rate 2
Public Const cdblRate1Weight As Double = 5 'max weight for rate 1
Public Const cdblRate1Length As Double = 12 'max length for rate 1
Public Const cdblRate2Length As Double = 60 'max length for rate 2
Public Const cdblLaserSetupRate As Double = 5 'laser setup time per sheet in minutes
Public Const cdblLaserSetupTime As Double = 0.5 'laser setup fixed time in minutes
Public Const cdblWaterJetSetupTime As Double = 15 'Waterjet setup fixed time in minutes
Public Const cdblWaterJetSetupRate As Double = 30 'Water Jet setup time per sheet load, in min
Public Const cdblStandardSheetWidth As Double = 60 'standard sheet size, in inches
Public Const cdblStandardSheetLength As Double = 120 'standard sheet size, in inches

'Standard Costs $/hr
Public Const cdblF115cost As Double = 120
Public Const cdblF300cost As Double = 44
Public Const cdblF210cost As Double = 42
Public Const cdblF140cost As Double = 80
Public Const cdblF145cost As Double = 175
Public Const cdblF155cost As Double = 120
Public Const cdblF325cost As Double = 65
Public Const cdblF400cost As Double = 48
Public Const cdblF385cost As Double = 37
Public Const cdblF500cost As Double = 48
Public Const cdblF525cost As Double = 47
Public Const cdblENGcost As Double = 50
Public Const cdblMaterialMarkup = 1.05  '50%
Public Const cdblTightPercent = 1.15 '15%
Public Const cdblNormalPercent = 1  '20%
Public Const cdblLoosePercent = 0.95 '20%

Public dblF115Hours As Double
Public dblF300Hours As Double
Public dblF210Hours As Double
Public dblF140Hours As Double
Public dblF155Hours As Double
Public dblF325Hours As Double
Public dblOtherHours As Double

Public dblF115Price As Double
Public dblF300Price As Double
Public dblF210Price As Double
Public dblF140Price As Double
Public dblF155Price As Double
Public dblF325Price As Double
Public dblOtherPrice As Double

Global value As Integer 'for returns
Global bReturn As Boolean 'for returns
Global swSelMgr As SelectionMgr

Public gstrConfigName As String

'CalculateCutInfo constants
Public Const cstrExcelLookupFile As String = "Laser.xls"
Public Const cdblPierceConstant As Double = 2 'constant added to calculated pierce total
Public Const consTabSpacing As Integer = 30

Sub CalcWeight(objPart As ModelDoc2, ByRef strMessage As String, ByRef swConfigMgr As ConfigurationManager, ByRef swConfig As Configuration)
    'updates "RawWeight" file property via an estimate of efficiency or by actual entered data by user
    Dim strWeightCalc As String
    Dim dblRawWeight As Double
    Dim dblSheetPercent As Double
    dblRawWeight = 0
    dblSheetPercent = 0
    'strWeightCalc = objPart.GetCustomInfoValue(gstrConfigName, "WeightCalc") 'see which mode we are working with
    '***MODIFIED 10/14/2013 TB ****
    strWeightCalc = objPart.GetCustomInfoValue(gstrConfigName, "rbWeightCalc")
    If strWeightCalc = "0" Then 'use efficiency
    
        Select Case GetThickness(objPart)
            Case Is > 1.2
                dblRawWeight = GetMass(objPart) / objPart.GetCustomInfoValue(gstrConfigName, "NestEfficiency") * 100 * 1.015
            Case Is > 0.9
                dblRawWeight = GetMass(objPart) / objPart.GetCustomInfoValue(gstrConfigName, "NestEfficiency") * 100 * 1.022
            Case Is > 0.82
                dblRawWeight = GetMass(objPart) / objPart.GetCustomInfoValue(gstrConfigName, "NestEfficiency") * 100 * 1.022
            Case Is > 0.7
                dblRawWeight = GetMass(objPart) / objPart.GetCustomInfoValue(gstrConfigName, "NestEfficiency") * 100 * 1.026
            Case Is > 0.59
                dblRawWeight = GetMass(objPart) / objPart.GetCustomInfoValue(gstrConfigName, "NestEfficiency") * 100 * 1.028
            Case Is > 0.43
                dblRawWeight = GetMass(objPart) / objPart.GetCustomInfoValue(gstrConfigName, "NestEfficiency") * 100 * 1.037
            Case Is > 0.34
                dblRawWeight = GetMass(objPart) / objPart.GetCustomInfoValue(gstrConfigName, "NestEfficiency") * 100 * 1.054
            Case Is > 0.29
                dblRawWeight = GetMass(objPart) / objPart.GetCustomInfoValue(gstrConfigName, "NestEfficiency") * 100 * 1.054
            Case Is > 0.23
                dblRawWeight = GetMass(objPart) / objPart.GetCustomInfoValue(gstrConfigName, "NestEfficiency") * 100 * 1.069
            Case Is > 0.1875
                dblRawWeight = GetMass(objPart) / objPart.GetCustomInfoValue(gstrConfigName, "NestEfficiency") * 100 * 1.096
            Case Else
                dblRawWeight = GetMass(objPart) / objPart.GetCustomInfoValue(gstrConfigName, "NestEfficiency") * 100
        End Select
    ElseIf objPart.GetCustomInfoValue(gstrConfigName, "Length") <> "" And objPart.GetCustomInfoValue(gstrConfigName, "Width") <> "" Then 'use numbers entered via user
       Select Case GetThickness(objPart)
            Case Is > 1.2
                dblRawWeight = GetThickness(objPart) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Length")) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Width")) * GetDensity(objPart) * 1.015
            Case Is > 0.9
                dblRawWeight = GetThickness(objPart) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Length")) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Width")) * GetDensity(objPart) * 1.022
            Case Is > 0.82
                dblRawWeight = GetThickness(objPart) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Length")) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Width")) * GetDensity(objPart) * 1.022
            Case Is > 0.7
                dblRawWeight = GetThickness(objPart) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Length")) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Width")) * GetDensity(objPart) * 1.026
            Case Is > 0.59
                dblRawWeight = GetThickness(objPart) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Length")) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Width")) * GetDensity(objPart) * 1.028
            Case Is > 0.43
                dblRawWeight = GetThickness(objPart) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Length")) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Width")) * GetDensity(objPart) * 1.037
            Case Is > 0.34
                dblRawWeight = GetThickness(objPart) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Length")) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Width")) * GetDensity(objPart) * 1.054
            Case Is > 0.29
                dblRawWeight = GetThickness(objPart) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Length")) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Width")) * GetDensity(objPart) * 1.054
            Case Is > 0.23
                dblRawWeight = GetThickness(objPart) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Length")) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Width")) * GetDensity(objPart) * 1.069
            Case Is > 0.1875
                dblRawWeight = GetThickness(objPart) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Length")) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Width")) * GetDensity(objPart) * 1.096
            Case Else
                 dblRawWeight = GetThickness(objPart) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Length")) * CDbl(objPart.GetCustomInfoValue(gstrConfigName, "Width")) * GetDensity(objPart)
        End Select
    End If
    If dblRawWeight > 0 Then
        objPart.DeleteCustomInfo2 gstrConfigName, "RawWeight"
        objPart.AddCustomInfo3 gstrConfigName, "RawWeight", swCustomInfoText, Format(dblRawWeight, ".####")
        
        dblSheetPercent = dblRawWeight / (GetThickness(objPart) * 60 * 120 * GetDensity(objPart))
        objPart.DeleteCustomInfo2 gstrConfigName, "SheetPercent"
        objPart.AddCustomInfo3 gstrConfigName, "SheetPercent", swCustomInfoText, Format(dblSheetPercent, ".####")
        strMessage = strMessage & "60x120%: " & Format(dblSheetPercent, ".####") & vbCrLf
    Else
        MsgBox "Raw Weight Cannot be Zero.  Fix Material Calculations"
        
    End If
    
End Sub

Sub CalculateBendInfo(swApp As SldWorks.SldWorks, objPart As ModelDoc2, objFace As Face2, ByRef strMessage As String, ByRef swConfigMgr As ConfigurationManager, ByRef swConfig As Configuration)

    'get longest bend line and number of bends
    Dim intBendCount As Integer
    Dim dblLongestBend As Double
    CountBends objPart, intBendCount, dblLongestBend, swApp
    If intBendCount = 0 Then
        MsgBox "No bend lines found", vbOKOnly + vbCritical, "No bends"
        GoTo ExitHere
    ElseIf objPart.GetCustomInfoValue(gstrConfigName, "PressBrake") <> "Checked" And objPart.GetCustomInfoValue(gstrConfigName, "F325") = 0 Then
        'objPart.DeleteCustomInfo2 "", "PressBrake"
        'objPart.AddCustomInfo3 "", "PressBrake", swCustomInfoText, "Checked"
        MsgBox "F140 Operation Not Checked and F325 Not Checked, but bends exist", vbOKOnly + vbCritical, "Go back and select one."
    End If
   
    'does the part need flipped?
    Dim blnFlip As Boolean
    blnFlip = FlipPart
   
    'find the longest length of the part
    Dim dblLength1 As Double
    Dim dblLength2 As Double
    LengthWidth objFace, dblLength1, dblLength2 'dblLength1 is really all we need here, it is the longest dim
   
   
    'run rate
    Dim dblRunRate As Double
    Dim dblWeight As Double
    dblWeight = GetMass(objPart)
    dblRunRate = FindRate(dblWeight, dblLength1)
   
    Dim dblF140_S As Double
    Dim dblF140_R As Double
   
    'calculate final setup values
    dblF140_S = dblLongestBend * cdblSetupRate / 12 + cdblBreakSetup
   
    If blnFlip = True Then
        dblF140_R = dblRunRate * (intBendCount + 1)
    Else
        dblF140_R = dblRunRate * (intBendCount)
    End If
    'update custom properties
    objPart.DeleteCustomInfo2 gstrConfigName, "F140_S"
    objPart.DeleteCustomInfo2 gstrConfigName, "F140_R"
    objPart.AddCustomInfo3 gstrConfigName, "F140_S", swCustomInfoText, Format(dblF140_S / 60, ".##")
    objPart.AddCustomInfo3 gstrConfigName, "F140_R", swCustomInfoText, Format(dblF140_R / 3600, ".####")
   
    objPart.DeleteCustomInfo2 gstrConfigName, "F140_S_Cost"                                                     'default cost to dblF140_S
    objPart.AddCustomInfo3 gstrConfigName, "F140_S_Cost", swCustomInfoText, Format(dblF140_S / 60, ".##")       '2/22/11 Todd
   
    dblF140Hours = (dblF140_S / 60 + dblF140_R / 3600 * frmMaterialUpdate.tbQty.value)
    dblF140Price = dblF140Hours * cdblF140cost
    strMessage = "Part weight: " & Format(dblWeight, ".##") & "lbs" & vbCrLf
    strMessage = strMessage & "Total bends: " & intBendCount & vbCrLf
    strMessage = strMessage & "Max bend length: " & Format(dblLongestBend, ".#") & "in" & vbCrLf
    strMessage = strMessage & "Longest length: " & Format(dblLength1, ".#") & "in" & vbCrLf
    strMessage = strMessage & "Calculated run rate: " & dblRunRate & "s/bend"
   
   
ExitHere:
   Dim part As Object
   Dim boolstatus As Boolean
   Set part = swApp.ActiveDoc
   boolstatus = part.EditRebuild3()
   boolstatus = part.EditRebuild3()
End Sub

Sub CheckBendTonnage(objPart As ModelDoc2, ByRef dblLongestLine As Double, swApp As SldWorks.SldWorks)
    Dim strBendExcelFile As String
    Dim intMaterial As Integer
    Dim strMaxLengthPossible As Double
    Dim dblThickness As Double
    Dim strTemp As String
    Dim intPosition As Integer
    Dim objWorksheet As Excel.Worksheet
    Dim objTempBook As Excel.Workbook
    Dim objWorkbook As Excel.Workbook
    'Dim blnTempExcel As Boolean
    Dim objExcel As Excel.Application
    'Dim blnTempWorkBook As Boolean
    Dim strSheetName As String
    
    'some Excel worksheet constants:
    Dim intRow As Integer
    Dim intSSNu As Integer
    Dim intSSN1 As Integer
    Dim intCS As Integer
    Dim intAL As Integer
    Dim intThicknessColumn As Integer
    
    'specify rows and colums by machine choice
    intRow = 5
    intSSNu = 4
    intSSN1 = 5
    intCS = 6
    intAL = 7
    intThicknessColumn = 3
    
    strBendExcelFile = "Laser.xls"
    strSheetName = "Bend"
    
    If frmMaterialUpdate.obSS.value = True Then
        intMaterial = intSSN1
    ElseIf frmMaterialUpdate.obCS.value = True Then
        intMaterial = intCS
    ElseIf frmMaterialUpdate.obAL.value = True Then
        intMaterial = intAL
    Else
        intMaterial = intSSN1
    End If
    
    strTemp = swApp.GetCurrentMacroPathFolder
    strTemp = swApp.GetCurrentMacroPathName
    
    intPosition = InStrRev(strTemp, "\")
    strBendExcelFile = Left$(swApp.GetCurrentMacroPathName, intPosition) & strBendExcelFile
    
    dblThickness = GetThickness(objPart)
     If Dir(strBendExcelFile) = "" Then
        MsgBox "Could not find " & strBendExcelFile, vbOKOnly + vbCritical, "Missing Excel File"
        'GoTo ExitHere
    End If

    Set objExcel = GlobalExcel
    objExcel.Visible = False
    Set objWorkbook = objExcel.Workbooks.Open(strBendExcelFile)
   
    'look trough the list see if we can find a match
    Dim strExcelValue As String
    On Error Resume Next
    Err.Clear
    strExcelValue = objWorkbook.Worksheets(strSheetName).Cells(intRow, intThicknessColumn).value
    If Err.Number = 9 Then 'couldn't find the proper tab in Excel
        MsgBox "Could not find material tab in Excel tables", vbOKOnly + vbCritical, "Material tab not found"
        'GetMaterialConstants = False
        GoTo ExitHere
    End If
    On Error GoTo 0
    Do While strExcelValue <> "" And strExcelValue >= (dblThickness - 0.005)
        'If Fuzz(CDbl(strExcelValue), dblThickness) Then
            strExcelValue = objWorkbook.Worksheets(strSheetName).Cells(intRow, intThicknessColumn).value
            strMaxLengthPossible = objWorkbook.Worksheets(strSheetName).Cells(intRow, intMaterial).value

            'GetMaterialConstants = True
            'Exit Do 'found what we are looking for jump out
        'End If
        intRow = intRow + 1
        strExcelValue = objWorkbook.Worksheets(strSheetName).Cells(intRow, intThicknessColumn).value
    Loop
    
    If strMaxLengthPossible < dblLongestLine Then
         MsgBox "Bend exceeds max length for standard tooling.  Check requirements"
    End If

ExitHere:
    objWorkbook.Close False
    Set objWorkbook = Nothing
    Set objExcel = Nothing
End Sub

Sub CalculateCutInfo(swApp As SldWorks.SldWorks, objPart As ModelDoc2, objFace As Face2, ByRef strMessage As String, ByRef swConfigMgr As ConfigurationManager, ByRef swConfig As Configuration)
    Dim strExcelLookupFile As String
    Dim dblDensity As Double
    dblDensity = GetDensity(objPart)
    strExcelLookupFile = ReturnExcelFile(swApp.GetCurrentMacroPathName)
    If Dir(strExcelLookupFile) = "" Then
        MsgBox "Could not find " & strExcelLookupFile, vbOKOnly + vbCritical, "Missing Excel File"
        GoTo ExitHere
    End If
    Dim strLaserType As String
    Dim strCuttingType As String
    strLaserType = objPart.GetCustomInfoValue(gstrConfigName, "OP20") 'find laser type from custom properties
    If strLaserType <> "" Then
        strLaserType = Left(strLaserType, 4)
        strLaserType = Right(strLaserType, 3)
    End If
  '  strCuttingType = objPart.GetCustomInfoValue(gstrConfigName, "CuttingType") 'find Cutting type from custom properties
    Select Case strLaserType
        Case Is = "105"
        
        Case Is = "115"
        
        Case Is = "120"
        
        Case Is = "125"
        
        Case Is = "135"
        
        Case Is = "155"
        
        Case Else
            strLaserType = ""
    End Select


    If strLaserType = "" Then
        MsgBox "Cost macro only works on 105, 115, 120, 125, 135, and 155 work centers."
        GoTo ExitHere
    End If
   
    '# of loops for face equals total number of pierces
    Dim intPierces As Integer
    intPierces = objFace.GetLoopCount + cdblPierceConstant
   
    ' Get list of edges and edge count for loop.
    Dim dblEdgeCount As Double
    Dim varEdgeList As Variant
    dblEdgeCount = objFace.GetEdgeCount
    varEdgeList = objFace.GetEdges()
   
    ' Set perimeter/cut length to zero
    Dim dblLength As Double 'total cut length
    dblLength = 0
   
    ' Looping through all edges on selected face.
    Dim objCurve As Curve
    Dim objEdge As Edge
    Dim intCount As Integer
    Dim curveParams As Variant
    For intCount = 0 To (dblEdgeCount - 1)
        ' Getting the edge object and curve parameters.
        Set objEdge = varEdgeList(intCount)
        Set objCurve = objEdge.GetCurve
        curveParams = objEdge.GetCurveParams2()
        ' Add length of curve to sum
        dblLength = dblLength + objCurve.GetLength2(curveParams(6), curveParams(7))
    Next intCount
    dblLength = dblLength * 39.36996 'convert to inches
    Set objCurve = Nothing
    
    'now time for each pierce and cut
    Dim blnSuccess As Boolean
    Dim dblCutTime As Double
    Dim dblSpeed As Double
    Dim dblThickness As Double
    Dim dblPierceTime As Double
    Dim strMaterial As String
   
    Dim dblAccLaserSetupTime As Double '**NEW**
    Dim dblSheetWeight As Double '**NEW*
    'Dim dblDensity As Double '**NEW**
    Dim dblRawWeight As Double '**NEW**
    dblDensity = GetDensity(objPart) '**NEW**



    dblSheetWeight = 0
    dblRawWeight = objPart.GetCustomInfoValue(gstrConfigName, "RawWeight")
 
    If frmMaterialUpdate.obSS.value = True Then
        strMaterial = "Stainless Steel" 'for right now assume everything is stainless steel
        dblDensity = GetDensity(objPart)
        If dblDensity < 0.28 Then
            MsgBox "Density doesn't match stainless steel"
        End If
    ElseIf frmMaterialUpdate.obCS.value = True Then
        strMaterial = "Carbon Steel"
        If dblDensity < 0.28 Then
            MsgBox "Density doesn't match carbon steel"
        End If
    ElseIf frmMaterialUpdate.obAL.value = True Then
        strMaterial = "Aluminum"
        If dblDensity < 0.095 Then
            MsgBox "Density doesn't match stainless steel"
        ElseIf dblDensity > 0.12 Then
            MsgBox "Density doesn't match Aluminum"
        End If
    Else
        strMaterial = "None"
    End If
    
    dblThickness = GetThickness(objPart) 'need to know how thick the part is

    dblSheetWeight = dblThickness * cdblStandardSheetWidth * cdblStandardSheetLength * dblDensity
   
    'dblAccLaserSetupTime = (dblRawWeight / dblSheetWeight) * cdblLaserSetupRate + cdblLaserSetupTime
    If strLaserType = "155" Then
        dblAccLaserSetupTime = cdblWaterJetSetupTime
    Else
        dblAccLaserSetupTime = cdblLaserSetupTime
    End If

    If dblThickness > 0 Then 'we found the material thickness
        'find the speed of cut and pierce time based on material & thickness from Excel table
        blnSuccess = GetMaterialConstants(dblThickness, strMaterial, dblSpeed, dblPierceTime, strExcelLookupFile, objPart)
    Else
        MsgBox "Error finding material thickness, is this a sheet metal part?", vbCritical + vbOKOnly, "Cannot find thickness"
    End If
    If blnSuccess = True Then 'found the values from Excel
        'total time to make all piercings
        dblPierceTime = dblPierceTime * (intPierces + dblLength / consTabSpacing)
        'total time to make the cut
        dblCutTime = dblLength / dblSpeed
        
        If strLaserType = "155" Then
            dblCutTime = dblCutTime + (dblRawWeight / dblSheetWeight) * cdblWaterJetSetupRate
            
        Else
            dblCutTime = dblCutTime + (dblRawWeight / dblSheetWeight) * cdblLaserSetupRate
        End If
        
        strMessage = "Number of holes: " & Format(intPierces - cdblPierceConstant - 1) & vbCrLf
        strMessage = strMessage & "Estimated number of tabs: " & Format(dblLength / consTabSpacing + cdblPierceConstant, "#") & vbCrLf
        strMessage = strMessage & "Total length of cut: " & Round(dblLength, 1) & " in." & vbCrLf
        strMessage = strMessage & "Total cut time: " & Format(dblCutTime / 60#, ".####") & " hours" & vbCrLf
        strMessage = strMessage & "Total pierce time: " & Format(dblPierceTime / 3600#, ".####") & " hours" & vbCrLf
       
        'update file properties
        objPart.DeleteCustomInfo2 gstrConfigName, "OP20_S"
        objPart.DeleteCustomInfo2 gstrConfigName, "OP20_R"



 
        dblAccLaserSetupTime = dblAccLaserSetupTime / 60
        If dblAccLaserSetupTime < 0.01 Then
            dblAccLaserSetupTime = 0.01
        End If
        objPart.AddCustomInfo3 gstrConfigName, "OP20_S", swCustomInfoText, Format(dblAccLaserSetupTime, ".##")
        dblF115Hours = (dblAccLaserSetupTime + (dblPierceTime / 3600 + dblCutTime / 60) * frmMaterialUpdate.tbQty.value)
        dblF115Price = dblF115Hours * cdblF115cost
       


        objPart.AddCustomInfo3 gstrConfigName, "OP20_R", swCustomInfoText, Format((dblPierceTime / 3600 + dblCutTime / 60), ".####")
    End If
   

ExitHere:

End Sub
Function GetThickness(objModel As ModelDoc2) As Double
    'returns the thickness of a sheet metal part in inches
    Dim objFeature As SldWorks.Feature
    Dim dblThickness As Double
    Dim objSheetMetal  As SldWorks.SheetMetalFeatureData

    Set objFeature = objModel.FirstFeature
    Do While Not objFeature Is Nothing
        If objFeature.GetTypeName2 = "SheetMetal" Then
            Set objSheetMetal = objFeature.GetDefinition
            dblThickness = objSheetMetal.Thickness / 0.0254 'thickness in inches
            Set objSheetMetal = Nothing
            Exit Do
        End If
        Set objFeature = objFeature.GetNextFeature
    Loop
    If dblThickness > 0 Then
        GetThickness = dblThickness
    Else
        GetThickness = -1
    End If
    Set objFeature = Nothing
End Function
Function GetSelectedFace(objPart As ModelDoc2) As Face2
'returns the selected face and has some checks to ensure sheet metal part
    Set GetSelectedFace = Nothing
    If objPart.GetType <> swDocPART Then
        MsgBox "This routine may only be run on a sheet metal part", vbOKOnly + vbCritical, "This routine runs on parts only"
        Exit Function
    End If
    Dim objSelManager As SelectionMgr
    Dim objSelectedEntity As Object
   
    Set objSelManager = objPart.SelectionManager
    If objSelManager.GetSelectedObjectCount <> 1 Then
        'MsgBox "Please select only the flat face", vbOKOnly + vbCritical, "Invalid selection"
        Exit Function
    End If
    Set objSelectedEntity = objSelManager.GetSelectedObject3(1)
    If objSelectedEntity Is Nothing Then
       ' MsgBox "Please select only the flat face", vbOKOnly + vbCritical, "Invalid selection"
        Exit Function
    End If
    If objSelectedEntity.GetType = swSelectType_e.swSelFACES Then
        Set GetSelectedFace = objSelManager.GetSelectedObject3(1)
    Else
        'MsgBox "Please select only the flat face", vbOKOnly + vbCritical, "Invalid selection"
    End If
    Set objSelectedEntity = Nothing
    Set objSelManager = Nothing
End Function
Function FindRate(dblWeight As Double, dblLength As Double) As Double
    If dblWeight > cdblRate3Weight Then
        FindRate = cdblRate5
    ElseIf dblWeight > cdblRate2Weight Then
        FindRate = cdblRate4
    ElseIf dblWeight > cdblRate1Weight Or dblLength > cdblRate2Length Then
        FindRate = cdblRate3
    ElseIf dblWeight > cdblRate1Weight Or dblLength > cdblRate1Length Then
        FindRate = cdblRate2
    Else
        FindRate = cdblRate1
    End If
End Function
Function FlipPart() As Boolean
'today just asks the user if part needs flipped
    Dim Response As VbMsgBoxResult
    FlipPart = MsgBox("Does this part need flipped?", vbYesNo + vbQuestion, "Flip part?")
    If Response = vbYes Then
        FlipPart = True
    Else
        FlipPart = False
    End If
End Function
Sub CountBends(objPart As SldWorks.ModelDoc2, ByRef intFeatureCount As Integer, ByRef dblLongestLine As Double, swApp As SldWorks.SldWorks)
    ' Get the 1st feature in part
    Dim Feature As SldWorks.Feature
    Dim FeatureName As String
    Dim FeatureType As String
    Dim SubFeat As SldWorks.Feature
    Set Feature = objPart.FirstFeature
    intFeatureCount = 0
    ' While we have a valid feature
    While Not Feature Is Nothing
        ' Get the name of the feature
        FeatureType = Feature.GetTypeName2
        If FeatureType = "FlatPattern" Then
            Set SubFeat = Feature.GetFirstSubFeature
            ' While we have a valid Sub-feature
            While Not SubFeat Is Nothing
                ' Get the type of the Sub-feature
                'If SubFeat.GetTypeName2 = "UiBend" Or SubFeat.GetTypeName2 = "UiFreeformBend" Then
                '    intFeatureCount = intFeatureCount + 1
                'End If
                'MsgBox SubFeat.Name
                If SubFeat.GetTypeName2 = "ProfileFeature" And Left(SubFeat.Name, 10) = "Bend-Lines" Then 'it is the bend line sketch
                    'what is the longest bendline?
                    Dim swSketch As Sketch
                    Dim vSketchSeg As Variant
                    Dim swSketchSeg As SketchSegment
                    Dim intCounter As Integer
                    Set swSketch = SubFeat.GetSpecificFeature2
                    vSketchSeg = swSketch.GetSketchSegments
                    On Error Resume Next
                    'If vSketchSeg <> Empty Then
                    If Not IsEmpty(vSketchSeg) Then  '  changed to isEmpty  Todd 2/17/11
                        dblLongestLine = 0  ' reset to 0 for hidden parts to count last loop through - Todd 2/17/11
                        For intCounter = 0 To UBound(vSketchSeg)
                            Set swSketchSeg = vSketchSeg(intCounter)
                            If swSketchSeg.GetType = swSketchSegments_e.swSketchLINE Then
                                If swSketchSeg.GetLength > dblLongestLine Then
                                    dblLongestLine = swSketchSeg.GetLength
                                End If
                            End If
                        Next intCounter
                        On Error GoTo 0
                        Set swSketchSeg = Nothing
                        Set swSketch = Nothing
                        'dblLongestLine = dblLongestLine * 39.36996 'convert to inches  moved outside loop - Todd 2/17/11
                        intFeatureCount = 0   ' reset to 0 for hidden parts to only count last loop through  - Todd 2/17/11
                        intFeatureCount = intFeatureCount + UBound(vSketchSeg) + 1
                    End If
                End If
                'Sketch::GetSketchSegments
                Set SubFeat = SubFeat.GetNextSubFeature
            ' Continue until the last Sub-feature is done
            Wend
            ' Continue until the last feature is done
        End If
        ' Get the next feature
        Set Feature = Feature.GetNextFeature()
    Wend
    dblLongestLine = dblLongestLine * 39.36996 'convert to inches  - Todd 2/17/11
    CheckBendTonnage objPart, dblLongestLine, swApp

End Sub
Function LengthWidth(objFace As Face2, ByRef dblLength1 As Double, ByRef dblLength2 As Double) As Boolean
    'Returns the 2 lengths of the plainar face in inches
    'dblLength1 is the largest dim, dblLength2 is the second largest
    On Error Resume Next
    Dim varBoundaries As Variant
    Dim dblX, dblY, dblZ As Double
    varBoundaries = objFace.GetBox
    dblX = Abs(varBoundaries(0) - varBoundaries(3)) * 39.36996
    dblY = Abs(varBoundaries(1) - varBoundaries(4)) * 39.36996
    dblZ = Abs(varBoundaries(2) - varBoundaries(5)) * 39.36996
    If dblX >= dblY And dblX >= dblZ Then
        dblLength1 = dblX
        If dblY > dblZ Then
            dblLength2 = dblY
        Else
            dblLength2 = dblZ
        End If
    ElseIf dblY > dblX And dblY > dblZ Then
        dblLength1 = dblY
        If dblX > dblZ Then
            dblLength2 = dblX
        Else
            dblLength2 = dblZ
        End If
    Else
        dblLength1 = dblZ
        If dblX > dblY Then
            dblLength2 = dblX
        Else
            dblLength2 = dblY
        End If
    End If
    If Err.Number = 0 Then
        LengthWidth = True
    Else
        LengthWidth = False
    End If
    On Error GoTo 0
End Function
Function GetMass(objModelDoc As ModelDoc2) As Double
    'returns the weight of the model in lbs
    GetMass = objModelDoc.Extension.CreateMassProperty.Mass * 2.20462262
End Function
Function GetDensity(objModelDoc As ModelDoc2) As Double
    'returns the density of the model in lbs/in^3
    GetDensity = objModelDoc.Extension.CreateMassProperty.Density * 0.000036127298
End Function
Function GetMaterialConstants(dblThickness As Double, strMaterial As String, ByRef dblSpeed As Double, ByRef dblPierce As Double, strFileName As String, objPart As ModelDoc2) As Boolean
    Dim objWorksheet As Excel.Worksheet
    Dim objTempBook As Excel.Workbook
    Dim objWorkbook As Excel.Workbook
    'Dim blnTempExcel As Boolean
    Dim objExcel As Excel.Application
    'Dim blnTempWorkBook As Boolean
    Dim strCuttingType As String
   
    'some Excel worksheet constants:
    Dim intRow As Integer
    Dim intPierceColumn As Integer
    Dim intSpeedColumn As Integer
    Dim intThicknessColumn As Integer
    
    'specify rows and colums by machine choice
    intRow = 5
    'strCuttingType = objPart.GetCustomInfoValue(gstrConfigName, "CuttingType")
    strCuttingType = objPart.GetCustomInfoValue(gstrConfigName, "OP20")
    If strCuttingType <> "" Then
        strCuttingType = Left(strCuttingType, 4)
        strCuttingType = Right(strCuttingType, 3)
    End If
    If strCuttingType = "120" Then
        intPierceColumn = 14
        intSpeedColumn = 13
        intThicknessColumn = 3
    ElseIf strCuttingType = "125" Then
        intPierceColumn = 16
        intSpeedColumn = 15
        intThicknessColumn = 3
    ElseIf strCuttingType = "155" Then
        intPierceColumn = 7
        intSpeedColumn = 6
        intThicknessColumn = 3
    Else '115, 135
        intPierceColumn = 5
        intSpeedColumn = 4
        intThicknessColumn = 3
    End If

    Set objExcel = GlobalExcel
    objExcel.Visible = False
    Set objWorkbook = objExcel.Workbooks.Open(strFileName)
 
    'look trough the list see if we can find a match
    Dim strExcelValue As String
    On Error Resume Next
    Err.Clear
    strExcelValue = objWorkbook.Worksheets(strMaterial).Cells(intRow, intThicknessColumn).value
    If Err.Number = 9 Then 'couldn't find the proper tab in Excel
        MsgBox "Could not find " & strMaterial & " tab in Excel tables", vbOKOnly + vbCritical, "Material tab not found"
        GetMaterialConstants = False
        GoTo ExitHere
    End If
    On Error GoTo 0
    Do While strExcelValue <> "" And strExcelValue >= (dblThickness - 0.005)
        'If Fuzz(CDbl(strExcelValue), dblThickness) Then
            dblSpeed = objWorkbook.Worksheets(strMaterial).Cells(intRow, intSpeedColumn).value
            dblPierce = objWorkbook.Worksheets(strMaterial).Cells(intRow, intPierceColumn).value
            GetMaterialConstants = True
            'Exit Do 'found what we are looking for jump out
        'End If
        intRow = intRow + 1
        strExcelValue = objWorkbook.Worksheets(strMaterial).Cells(intRow, intThicknessColumn).value
    Loop

ExitHere:
    objWorkbook.Close False
    Set objWorkbook = Nothing
    Set objExcel = Nothing
End Function
Function Fuzz(dblNumber1 As Double, dblNumber2 As Double) As Boolean
'returns true if the numbers are within .005 of each other
    If Abs(dblNumber1 - dblNumber2) < 0.005 Then
        Fuzz = True
    Else
        Fuzz = False
    End If
End Function
Public Function ReturnExcelFile(strFullFile As String) As String
'strips off the path and returns the Excel file
    Dim intPosition As Integer
    intPosition = InStrRev(strFullFile, "\")
    ReturnExcelFile = Left$(strFullFile, intPosition) & cstrExcelLookupFile
End Function
Sub BendData(swApp As SldWorks.SldWorks, objPart As ModelDoc2)

Dim swSheetMetal As SldWorks.SheetMetalFeatureData
Dim swFeat As SldWorks.Feature
Dim lngBendAllowanceType As Long
Dim dblBendAllowance As Double

Set swFeat = objPart.FirstFeature
While Not swFeat Is Nothing
'Debug.Print swfeat.Name & " - " & swfeat.GetTypeName

If swFeat.GetTypeName = "SheetMetal" Then
Set swSheetMetal = swFeat.GetDefinition
lngBendAllowanceType = swSheetMetal.BendAllowanceType
dblBendAllowance = swSheetMetal.BendAllowance

End If

Set swFeat = swFeat.GetNextFeature

Wend
End Sub

Public Sub MaterialCost()
    Dim swApp               As SldWorks.SldWorks
    Dim objPart             As ModelDoc2
    Dim swConfigMgr         As SldWorks.ConfigurationManager
    Dim swConfig            As SldWorks.Configuration
    Dim swCustPropMgr       As SldWorks.CustomPropertyManager
    Dim strCutMessage       As String
    Dim strBendMessage      As String
    Dim strWeightMessage      As String
    Dim strOP20 As String
    Dim strOptiMaterial As String
    Dim strFileName As String
    blnTempExcel = False
    blnTempWorkBook = False
    
    Set swApp = Application.SldWorks
    Set objPart = swApp.ActiveDoc
    Set swConfigMgr = objPart.ConfigurationManager
    Set swConfig = swConfigMgr.ActiveConfiguration
    Debug.Print "Name of this configuration:                             " & swConfig.Name
    If swConfig.Name = "Default" Then
        gstrConfigName = ""
    Else
        gstrConfigName = swConfig.Name
    End If
    
    
    strFileName = modMaterialUpdate.ReturnExcelFile(swApp.GetCurrentMacroPathName)
    
    If objPart Is Nothing Then
        MsgBox "This routine may only be run with a file open", vbOKOnly + vbCritical, "Please open file first"
        GoTo ExitHere
    End If
    
    frmMaterialUpdate.Show
    
    If frmMaterialUpdate.obOther.value = True Then
        MsgBox "Material Unknown, Exiting Program."
        Exit Sub
    End If
    
    If objPart.GetCustomInfoValue("", "OP20") = "F110 - TUBE LASER" Then
        Dim PathName As String
        Set SP.swModel = objPart
        PathName = objPart.GetPathName
        Set SP.swApp = swApp
        Call SP.ExtractTubeData
        Call SP.TubeCustomProperties
        Call SP.CustomProperties
        'Call SP.CreateDrawing
        Call SP.SaveCurrentModel
        Exit Sub
        
    End If
    
    Dim objFace As Face2
    Set objFace = GetSelectedFace(objPart)
    
    If objFace Is Nothing Then
        Call SelectFlatPattern(objPart)
        Set objFace = GetFixedFace(objPart)
        If objFace Is Nothing Then
        MsgBox "Must Select Flat Face for this Part, Exiting Program."
        GoTo ExitHere
        End If
        Call FlattenPart(objPart)
    End If
    
'   Check to see that this part is a laser/waterjet part
    strOP20 = ""
    strOP20 = objPart.GetCustomInfoValue(gstrConfigName, "OP20")
    If strOP20 <> "" Then
        strOP20 = Left(strOP20, 4)
        strOP20 = Right(strOP20, 3)
    End If
    If strOP20 = "105" Or strOP20 = "115" Or strOP20 = "120" Or strOP20 = "125" Or strOP20 = "135" Or strOP20 = "155" Then
        Call CalcWeight(objPart, strWeightMessage, swConfigMgr, swConfig) 'updates weight information
        Call CalculateCutInfo(swApp, objPart, objFace, strCutMessage, swConfigMgr, swConfig) 'updates laser cutting info
        strOptiMaterial = objPart.GetCustomInfoValue(gstrConfigName, "OptiMaterial")
        
        If Not ThicknessCheck.CheckThickness(objPart, strFileName, strOptiMaterial) Then
            MsgBox "Check model thickness vs Opti Material"
        End If
        strWeightMessage = "Weight information:" & vbCrLf & strWeightMessage & vbCrLf

    
        If strCutMessage <> "" Then
            strCutMessage = "-----" & vbCrLf & "Cut information:" & vbCrLf & strCutMessage & vbCrLf
        End If
    ElseIf strOP20 <> "110" Then
        MsgBox "No OP20 Selected"
    End If
    
    'Check for Tapped Holes
    Call TappedHoles(objPart)
    
    'Check to see if part gets bent
    'If objPart.GetCustomInfoValue(gstrConfigName, "PressBrake") = "Checked" Then
        Call CalculateBendInfo(swApp, objPart, objFace, strBendMessage, swConfigMgr, swConfig) 'updates press break info
        If strBendMessage <> "" Then
            strBendMessage = "-----" & vbCrLf & "Bend information:" & vbCrLf & strBendMessage
        End If
        
    
    'End If

    If strBendMessage <> "" Or strCutMessage <> "" Then
        MsgBox strWeightMessage & strCutMessage & strBendMessage, vbInformation + vbOKOnly, "Complete"
    End If
    
    If gstrConfigName <> "Default" Then
        objPart.DeleteCustomInfo2 gstrConfigName, "confDrawing"
        objPart.AddCustomInfo3 gstrConfigName, "confDrawing", swCustomInfoText, gstrConfigName
    End If
    
    
' If quoting enabled then...
    If frmMaterialUpdate.cbQuote.value = True Then
        objPart.DeleteCustomInfo2 gstrConfigName, "TotalPrice"
        objPart.AddCustomInfo3 gstrConfigName, "TotalPrice", swCustomInfoText, Format(TotalCost(objPart), ".##")
    End If
ExitHere:
    'little cleanup
    Call UnFlattenPart(objPart)
    Set objFace = Nothing
    Set objPart = Nothing
    Set swApp = Nothing
End Sub


Sub FlattenPart(model As ModelDoc2)

If model.GetBendState <> swSMBendStateFlattened Then
    value = model.SetBendState(swSMBendStateFlattened)
    'Debug.Print "SetBendState = " & value
    bReturn = model.ForceRebuild3(False)
    bReturn = model.ForceRebuild3(False)
End If


End Sub

Sub UnFlattenPart(model As ModelDoc2)
If model.GetBendState = swSMBendStateFlattened Then
    value = model.SetBendState(swSMBendStateFolded)
    'Debug.Print "SetBendState = " & value
    bReturn = model.ForceRebuild3(False)
    bReturn = model.ForceRebuild3(False)
End If

End Sub

Sub GetFlatFeatures(model As ModelDoc2)
Dim swFeat As SldWorks.Feature
Dim swFlattPatt As SldWorks.FlatPatternFeatureData
Dim swSheetMetal As SldWorks.SheetMetalFeatureData

swSelMgr = model.SelectionManager

SelectFlatPattern model
If swSelMgr.GetSelectedObjectCount2(0) > 0 Then
    Set swFlatPatt = swSelMgr.GetSelectedObject6(1, 0)
End If

SelectSheetMetal model
If swSelMgr.GetSelectedObjectCount2(0) > 0 Then
End If

End Sub

Function GetFixedFace(swModel As ModelDoc2) As Face2
    Dim swSelMgr                As SldWorks.SelectionMgr
    Dim swFeat                  As SldWorks.Feature
    Dim swFlatPatt              As SldWorks.FlatPatternFeatureData
    Dim swFixedFace             As SldWorks.Face2
    Dim bRet                    As Boolean
    
    Set swSelMgr = swModel.SelectionManager
    Set swFeat = swSelMgr.GetSelectedObject5(1)
    Set swFlatPatt = swFeat.GetDefinition


  '   Roll back part because flat pattern will absorb faces
    bRet = swFlatPatt.AccessSelections(swModel, Nothing)
    Set swFixedFace = swFlatPatt.FixedFace
    bRet = swFixedFace.Select(False): Debug.Assert bRet
    Set GetFixedFace = swFixedFace
    
    ' Cancel any changes made
    swFlatPatt.ReleaseSelectionAccess
End Function

Sub SelectFlatPattern(model As ModelDoc2)

model.ClearSelection2 (True)

Dim swFeatMgr As FeatureManager
Dim swFeat As Feature
Dim swFeatName As String

Set swFeatMgr = model.FeatureManager
Set swFeat = model.FirstFeature

Do While Not swFeat Is Nothing

swFeatName = swFeat.GetTypeName()
If swFeatName = "FlatPattern" Then swFeat.Select2 False, 0
Set swFeat = swFeat.GetNextFeature
Loop

End Sub

Sub SelectSheetMetal(model As ModelDoc2)

model.ClearSelection2 (True)

Dim swFeatMgr As FeatureManager
Dim swFeat As Feature
Dim swFeatName As String

Set swFeatMgr = model.FeatureManager
Set swFeat = model.FirstFeature

Do While Not swFeat Is Nothing

swFeatName = swFeat.GetTypeName()
If swFeatName = "SheetMetal" Then swFeat.Select2 False, 0
Set swFeat = swFeat.GetNextFeature
Loop

End Sub

Function GetWorkCenterCosts(dblThickness As Double, strMaterial As String, strFileName As String, objPart As ModelDoc2) As Boolean
    Dim objWorksheet As Excel.Worksheet
    Dim objTempBook As Excel.Workbook
    Dim objWorkbook As Excel.Workbook
    'Dim blnTempExcel As Boolean
    Dim objExcel As Excel.Application
    'Dim blnTempWorkBook As Boolean
    Dim strCuttingType As String
   
    'some Excel worksheet constants:
    Dim intRow As Integer
    Dim intPierceColumn As Integer
    Dim intSpeedColumn As Integer
    Dim intThicknessColumn As Integer
    
    'specify rows and colums by machine choice
    intRow = 5
    strCuttingType = objPart.GetCustomInfoValue(gstrConfigName, "CuttingType")
    If strCuttingType <> "" Then
        strCuttingType = Left(strCuttingType, 4)
        strCuttingType = Right(strCuttingType, 3)
    End If
    If strCuttingType = "120" Then
        intPierceColumn = 14
        intSpeedColumn = 13
        intThicknessColumn = 3
    ElseIf strCuttingType = "125" Then
        intPierceColumn = 16
        intSpeedColumn = 15
        intThicknessColumn = 3
    ElseIf strCuttingType = "155" Then
        intPierceColumn = 7
        intSpeedColumn = 6
        intThicknessColumn = 3
    Else '115, 135
        intPierceColumn = 5
        intSpeedColumn = 4
        intThicknessColumn = 3
    End If


    Set objExcel = GlobalExcel
    objExcel.Visible = False
    Set objWorkbook = objExcel.Workbooks.Open(strFileName)
   
    'look through the list see if we can find a match
    Dim strExcelValue As String
    On Error Resume Next
    Err.Clear
    strExcelValue = objWorkbook.Worksheets(strMaterial).Cells(intRow, intThicknessColumn).value
    If Err.Number = 9 Then 'couldn't find the proper tab in Excel
        MsgBox "Could not find " & strMaterial & " tab in Excel tables", vbOKOnly + vbCritical, "Material tab not found"
        GetMaterialConstants = False
        GoTo ExitHere
    End If
    On Error GoTo 0
    Do While strExcelValue <> "" And strExcelValue >= (dblThickness - 0.005)
        'If Fuzz(CDbl(strExcelValue), dblThickness) Then
            dblSpeed = objWorkbook.Worksheets(strMaterial).Cells(intRow, intSpeedColumn).value
            dblPierce = objWorkbook.Worksheets(strMaterial).Cells(intRow, intPierceColumn).value
            GetMaterialConstants = True
            'Exit Do 'found what we are looking for jump out
        'End If
        intRow = intRow + 1
        strExcelValue = objWorkbook.Worksheets(strMaterial).Cells(intRow, intThicknessColumn).value
    Loop

ExitHere:
    objWorkbook.Close False
    Set objWorkbook = Nothing
    Set objExcel = Nothing
End Function

Function TotalCost(objPart As ModelDoc2) As Double
Dim strF210 As String
Dim strF210S As String
Dim strF210R As String

TotalCost = 0
'Order Processing Costs
Dim strOrderSetup As String
Dim strOrderRun As String
Dim strOrderHours As String
Dim strOrderPrice As String
Dim MatPercentage As Double

strOrderSetup = 20  'in dollars
strOrderRun = 3 'in dollars

strOrderPrice = strOrderSetup + strOrderRun * (frmMaterialUpdate.tbQty.value - 1)
TotalCost = TotalCost + strOrderPrice

objPart.DeleteCustomInfo2 gstrConfigName, "NM_Setup"
objPart.DeleteCustomInfo2 gstrConfigName, "NM_Run"
objPart.DeleteCustomInfo2 gstrConfigName, "NM_Hours"
objPart.DeleteCustomInfo2 gstrConfigName, "NM_Price"

objPart.AddCustomInfo3 gstrConfigName, "NM_Setup", swCustomInfoText, Format(strOrderSetup, ".###")
objPart.AddCustomInfo3 gstrConfigName, "NM_Run", swCustomInfoText, Format(strOrderRun, ".###")
objPart.AddCustomInfo3 gstrConfigName, "NM_Hours", swCustomInfoText, Format(strOrderHours, ".###")
objPart.AddCustomInfo3 gstrConfigName, "NM_Price", swCustomInfoText, Format(strOrderPrice, ".###")

'Material Cost
Dim strRawWeight As String
Dim strMatCost As String
Dim strTotalWeight As String
Dim strCostPerLB As String

strCostPerLB = frmMaterialUpdate.tbQty.value
strRawWeight = objPart.GetCustomInfoValue(gstrConfigName, "RawWeight")
strTotalWeight = (strRawWeight * strCostPerLB)
strMatCost = frmMaterialUpdate.tbCostPerLB * strTotalWeight

objPart.DeleteCustomInfo2 gstrConfigName, "Total_Weight"
objPart.DeleteCustomInfo2 gstrConfigName, "MaterailCostPerLB"
objPart.DeleteCustomInfo2 gstrConfigName, "TotalMaterialCost"

objPart.AddCustomInfo3 gstrConfigName, "Total_Weight", swCustomInfoText, Format(strTotalWeight, ".###")
objPart.AddCustomInfo3 gstrConfigName, "MaterailCostPerLB", swCustomInfoText, Format(strCostPerLB, ".###")
objPart.AddCustomInfo3 gstrConfigName, "TotalMaterialCost", swCustomInfoText, Format(strMatCost, ".##")

TotalCost = TotalCost + strMatCost * cdblMaterialMarkup

'F115 Price
Dim strF105 As String
objPart.DeleteCustomInfo2 gstrConfigName, "F115_Price"
objPart.DeleteCustomInfo2 gstrConfigName, "F115_Hours"
strF105 = objPart.GetCustomInfoValue(gstrConfigName, "CuttingType")
If strF105 = "F105" Or strF105 = "F155" Then
    TotalCost = TotalCost + dblF115Price
    objPart.AddCustomInfo3 gstrConfigName, "F115_Price", swCustomInfoText, Format(dblF115Price, ".##")
    objPart.AddCustomInfo3 gstrConfigName, "F115_Hours", swCustomInfoText, Format(dblF115Hours, ".###")
Else
    objPart.DeleteCustomInfo2 gstrConfigName, "F115_S"
    objPart.DeleteCustomInfo2 gstrConfigName, "F115_R"
End If

'F300 Price
Dim strCuttingType As String
Dim strF300S As String
Dim strF300R As String

strCuttingType = objPart.GetCustomInfoValue(gstrConfigName, "CuttingType")
objPart.DeleteCustomInfo2 gstrConfigName, "F300_Price"
objPart.DeleteCustomInfo2 gstrConfigName, "F300_Hours"
If strCuttingType = "F300" Then
    strF300S = objPart.GetCustomInfoValue(gstrConfigName, "F300_S")
    strF300R = objPart.GetCustomInfoValue(gstrConfigName, "F300_R")
    If strF300S = "" Then
        strF300S = 0
    ElseIf strF300R = "" Then
        strF300R = 0
    End If
    
    dblF300Hours = (strF300S + strF300R * frmMaterialUpdate.tbQty.value)
    dblF300Price = dblF300Hours * cdblF300cost
    objPart.AddCustomInfo3 gstrConfigName, "F300_Hours", swCustomInfoText, Format(dblF300Hours, ".###")
    objPart.AddCustomInfo3 gstrConfigName, "F300_Price", swCustomInfoText, Format(dblF300Price, ".##")
    TotalCost = TotalCost + dblF300Price
Else
    objPart.DeleteCustomInfo2 gstrConfigName, "F300_S"
    objPart.DeleteCustomInfo2 gstrConfigName, "F300_R"
End If

'F210 Price
strF210 = objPart.GetCustomInfoValue(gstrConfigName, "F210")
objPart.DeleteCustomInfo2 gstrConfigName, "F210_Price"
objPart.DeleteCustomInfo2 gstrConfigName, "F210_Hours"
If strF210 = 1 Then
    strF210S = objPart.GetCustomInfoValue(gstrConfigName, "F210_S")
    strF210R = objPart.GetCustomInfoValue(gstrConfigName, "F210_R")
    If strF210S = "" Then
        strF210S = 0
    End If
    If strF210R = "" Then
        strF210R = 0
    End If
    
    dblF210Hours = (strF210S + strF210R * frmMaterialUpdate.tbQty.value)
    dblF210Price = dblF210Hours * cdblF210cost
    
    objPart.AddCustomInfo3 gstrConfigName, "F210_Hours", swCustomInfoText, Format(dblF210Hours, ".###")
    objPart.AddCustomInfo3 gstrConfigName, "F210_Price", swCustomInfoText, Format(dblF210Price, ".##")
    TotalCost = TotalCost + dblF210Price
Else
    objPart.DeleteCustomInfo2 gstrConfigName, "F210_S"
    objPart.DeleteCustomInfo2 gstrConfigName, "F210_R"
End If

'F140 Price
Dim strF140 As String
strF140 = objPart.GetCustomInfoValue(gstrConfigName, "PressBrake")
objPart.DeleteCustomInfo2 gstrConfigName, "F140_Price"
objPart.DeleteCustomInfo2 gstrConfigName, "F140_Hours"
If strF140 = "Checked" Then
    TotalCost = TotalCost + dblF140Price
 
    objPart.AddCustomInfo3 gstrConfigName, "F140_Price", swCustomInfoText, Format(dblF140Price, ".##")
    objPart.AddCustomInfo3 gstrConfigName, "F140_Hours", swCustomInfoText, Format(dblF140Hours, ".###")
Else
    objPart.DeleteCustomInfo2 gstrConfigName, "F140_S"
    objPart.DeleteCustomInfo2 gstrConfigName, "F140_R"
End If

'F325 Price
Dim strF325 As String
Dim strF325S As String
Dim strF325R As String

strF325 = objPart.GetCustomInfoValue(gstrConfigName, "F325")
objPart.DeleteCustomInfo2 gstrConfigName, "F325_Price"
objPart.DeleteCustomInfo2 gstrConfigName, "F325_Hours"
If strF325 = "1" Then
    strF325S = objPart.GetCustomInfoValue(gstrConfigName, "F325_S")
    strF325R = objPart.GetCustomInfoValue(gstrConfigName, "F325_R")
    If strF325S = "" Then
        strF325S = 0
    ElseIf strF325R = "" Then
        strF325R = 0
    End If
    dblF325Hours = (strF325S + strF325R * frmMaterialUpdate.tbQty.value)
    dblF325Price = dblF325Hours * cdblF325cost
    TotalCost = TotalCost + dblF325Price
    
    objPart.AddCustomInfo3 gstrConfigName, "F325_Price", swCustomInfoText, Format(dblF325Price, ".##")
    objPart.AddCustomInfo3 gstrConfigName, "F325_Hours", swCustomInfoText, Format(dblF325Hours, ".###")
Else
    objPart.DeleteCustomInfo2 gstrConfigName, "F325_S"
    objPart.DeleteCustomInfo2 gstrConfigName, "F325_R"
End If



'Other WC
Dim strOther As String
Dim strOtherS As String
Dim strOtherR As String
Dim strOtherWC As String
Dim dblOtherCost As String

strOther = objPart.GetCustomInfoValue(gstrConfigName, "OtherWC_CB")
objPart.DeleteCustomInfo2 gstrConfigName, "Other_Price"
objPart.DeleteCustomInfo2 gstrConfigName, "Other_Hours"
If strOther = 1 Then
    strOtherS = objPart.GetCustomInfoValue(gstrConfigName, "Other_S")
    strOtherR = objPart.GetCustomInfoValue(gstrConfigName, "Other_R")
    If strOtherS = "" Then
        strOtherS = 0
    ElseIf strOtherR = "" Then
        strOtherR = 0
    End If
    
    strOtherWC = objPart.GetCustomInfoValue(gstrConfigName, "Other_WC")
    dblOtherCost = 0
    If strOtherWC = "400" Then
        dblOtherCost = cdblF400cost
    ElseIf strOtherWC = "385" Then
         dblOtherCost = cdblF385cost
    ElseIf strOtherWC = "525" Then
         dblOtherCost = cdblF525cost
    ElseIf strOtherWC = "500" Then
         dblOtherCost = cdblF500cost
    ElseIf strOtherWC = "145" Then
         dblOtherCost = cdblF145cost
    End If
    
    dblOtherHours = (strOtherS + strOtherR * frmMaterialUpdate.tbQty.value)
    dblOtherPrice = dblOtherHours * dblOtherCost
    
    TotalCost = TotalCost + dblOtherPrice
    
    objPart.AddCustomInfo3 gstrConfigName, "Other_Price", swCustomInfoText, Format(dblOtherPrice, ".##")
    objPart.AddCustomInfo3 gstrConfigName, "Other_Hours", swCustomInfoText, Format(dblOtherHours, ".##")
Else
    objPart.DeleteCustomInfo2 gstrConfigName, "Other_S"
    objPart.DeleteCustomInfo2 gstrConfigName, "Other_R"
End If

objPart.DeleteCustomInfo2 gstrConfigName, "QuoteQty"
objPart.DeleteCustomInfo2 gstrConfigName, "QuoteMatCost"
objPart.AddCustomInfo3 gstrConfigName, "QuoteQty", swCustomInfoText, frmMaterialUpdate.tbQty.value
objPart.AddCustomInfo3 gstrConfigName, "QuoteMatCost", swCustomInfoText, Format(frmMaterialUpdate.tbCostPerLB.value, ".##")

If frmMaterialUpdate.obTight.value = True Then
    TotalCost = TotalCost * cdblTightPercent
ElseIf frmMaterialUpdate.obLoose.value = True Then
    TotalCost = TotalCost * cdblLoosePercent
Else
    TotalCost = TotalCost * cdblNormalPercent
End If

If TotalCost > 10000 Then
    TotalCost = TotalCost - frmMaterialUpdate.tbQty.value
'    TotalCost = TotalCost - strMatCost * (cdblMaterialMarkup - 1)
ElseIf TotalCost > 1000 Then
    TotalCost = TotalCost - ((TotalCost - 1000) / (10000 - 1000) * (frmMaterialUpdate.tbQty.value))
'    TotalCost = TotalCost - strMatCost * (cdblMaterialMarkup - 1 + 0.05)
End If
MatPercentage = strMatCost / TotalCost
End Function

Sub TappedHoles(swModel As ModelDoc2)

 

    Dim swFeat                                  As SldWorks.Feature

    Dim swSubFeat                               As SldWorks.Feature
    Dim swWizHole                               As SldWorks.WizardHoleFeatureData2

    Dim sFeatType                               As String

    Dim swCosThread                             As SldWorks.CosmeticThreadFeatureData

    

    Dim i                                       As Long

    Dim j                                       As Long

    Dim bRet                                    As Boolean

    Dim strFeatName As String
    Dim intNumHoles As Integer
    Dim intNumSetups As Integer
    Dim dblTapDiameter As Double
    Dim boolIsTappedHole As Boolean
    Dim boolOutsource As Boolean
    Dim boolTappingRequired As Boolean
    Dim dblSetup As Double
    Dim dblDrillHoleDiameter As Double
    Dim strLaserType As String

   ' Debug.Print "File = " & swModel.GetPathName
    intNumSetups = 0
    intNumHoles = 0
    boolOutsource = False
    boolTappingRequired = False
    
    Set swFeat = swModel.FirstFeature

    Do While Not swFeat Is Nothing

        Set swSubFeat = swFeat.GetFirstSubFeature

        If swFeat.IsSuppressed = False Then
        strFeatName = swFeat.Name
        
            boolIsTappedHole = InStr(1, strFeatName, "Tapped", vbBinaryCompare)
            If boolIsTappedHole = True Then
                    'If the material is stainless steel, the hole size needs to change for larger
                    If frmMaterialUpdate.obSS.value = True Then
                        strLaserType = swModel.GetCustomInfoValue(gstrConfigName, "OP20")
                        If strLaserType <> "" Then
                            strLaserType = Left(strLaserType, 4)
                            strLaserType = Right(strLaserType, 3)
                        End If
                        Select Case strLaserType
                            Case Is = "105"
                                Set swWizHole = swFeat.GetDefinition
                                dblDrillHoleDiameter = swWizHole.TapDrillDiameter * 39.36996
                                MsgBox "Confirm drill diamter for " & strFeatName & "?" & vbCrLf & "Current drill size = " & Format(dblDrillHoleDiameter, ".###"), vbOKOnly, "Tapping Stainless Steel"
                                
                            Case Is = "115"
                                Set swWizHole = swFeat.GetDefinition
                                dblDrillHoleDiameter = swWizHole.TapDrillDiameter * 39.36996
                                MsgBox "Confirm drill diamter for " & strFeatName & "?" & vbCrLf & "Current drill size = " & Format(dblDrillHoleDiameter, ".###"), vbOKOnly, "Tapping Stainless Steel"
                                
                            Case Is = "135"
                                Set swWizHole = swFeat.GetDefinition
                                dblDrillHoleDiameter = swWizHole.TapDrillDiameter * 39.36996
                                MsgBox "Confirm drill diamter for " & strFeatName & "?" & vbCrLf & "Current drill size = " & Format(dblDrillHoleDiameter, ".###"), vbOKOnly, "Tapping Stainless Steel"
                                
                            End Select
 
                    End If

                    
                    intNumSetups = intNumSetups + 1
        
        
                    Do While Not swSubFeat Is Nothing

            

                    sFeatType = swSubFeat.GetTypeName

                    If swSubFeat.IsSuppressed = False Then
            

                        Select Case sFeatType

                            Case "CosmeticThread"

                                'Debug.Print "    " & swSubFeat.Name & " [" & sFeatType & "]"

                    

                                Set swCosThread = swSubFeat.GetDefinition

                                dblTapDiameter = swCosThread.Diameter * 39.36996
                    
                                If dblTapDiameter < 1.01 Then
                                    intNumHoles = intNumHoles + 1
                                    boolTappingRequired = True
                                Else
                                    boolOutsource = True
                                End If

                                'Debug.Print "      ApplyThread      = " & swCosThread.ApplyThread

                                'Debug.Print "      BlindDepth       = " & swCosThread.BlindDepth * 1000# & " mm"

                                'Debug.Print "      Diameter         = " & swCosThread.Diameter * 1000# & " mm"

                                'Debug.Print "      DiameterType     = " & swCosThread.DiameterType

                                'Debug.Print "      ThreadCallout    = " & swCosThread.ThreadCallout

                    

                                'Debug.Print ""

                        End Select

            
                    End If
                    Set swSubFeat = swSubFeat.GetNextSubFeature
            
                    Loop
            End If
        End If
        Set swFeat = swFeat.GetNextFeature
        
    Loop
    
    
    If boolOutsource = True Then
        intNumHoles = 0
        intNumSetups = 0
        MsgBox "Holes bigger than 1.0IN diameter.  Outsource to riverside."
    Else
        If boolTappingRequired = True Then
            If swModel.GetCustomInfoValue("", "F220") <> "1" Then
                swModel.DeleteCustomInfo2 "", "F220"
                swModel.AddCustomInfo3 "", "F220", swCustomInfoText, 1
                MsgBox "Tapped Holes < 1.0IN exist, but F220 operation is not selected." & vbCrLf & "F220 selected changed to Checked."
                
            End If
                MsgBox "Number of holes: " & intNumHoles & vbCrLf & "Number of setups: " & intNumSetups
                dblSetup = intNumSetups * 0.015 + 0.085
                If dblSetup < 0.1 Then
                    dblSetup = 0.1
                End If
                swModel.DeleteCustomInfo2 "", "F220_S"
                swModel.DeleteCustomInfo2 "", "F220_R"
                swModel.DeleteCustomInfo2 "", "F220_RN"
                swModel.AddCustomInfo3 "", "F220_S", swCustomInfoText, Format(dblSetup, ".##")
                swModel.AddCustomInfo3 "", "F220_R", swCustomInfoText, Format(intNumHoles * 0.01, ".###")
                swModel.AddCustomInfo3 "", "F220_RN", swCustomInfoText, "TAP HOLES PER CAD"
            
        End If
    End If
    

End Sub

