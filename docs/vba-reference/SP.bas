Attribute VB_Name = "SP"
Public boolENG                 As Boolean
Public boolFindSheetMetal      As Boolean
Public BasePath                As String
Public biggestArea             As Double
Public bRet                    As Boolean
Public bQuote                  As Boolean
Public bQuoteASM               As Boolean
Public BodyCount               As Integer
Public boolstatus              As Boolean
Public bVolumeUP               As Boolean
Public bVolumeDN               As Boolean
Public calcVolume              As Double
Public currentArea             As Double
Public CheckOP20               As String
Public DocName                 As String
Public Errors                  As Integer
Public ExcelFile               As String
Public Feature                 As String
Public Features                As Variant
Public filePath                As String
Public FilePathDXF             As String
Public FilePath2               As String
Public FilePath2DXF            As String
Public FileType                As String
Public FlatPattern             As String
Public FileName                As String
Public gstrConfigName          As String
Public intOverwrite            As Integer
Public KFactor                 As Double
Public longstatus              As Long, longwarnings               As Long, longerrors                  As Long
Public MainAssembly            As String
Public MaterialType            As String
Public ModelName               As String
Public ModelNames()            As String
Public ModelNamesRedo()        As String
Public ModelNameDXF            As String
Public ModelNameRaw            As String
Public ModelNamesRaw()         As String
Public myView                  As View


Public objBendTable            As Excel.Workbook
Public objExcel                As Excel.Application
Public objFeat                 As Object
Public objSheet                As Excel.Worksheet
Public objWorkbook             As Excel.Workbook

' Material Data from Material-2022v4.xlsx
Public MaterialData() As Variant
Public ThickCheckData() As Variant

Public StainlessLoaded As Boolean
Public CarbonLoaded As Boolean
Public AluminumLoaded As Boolean
Public BendLoaded As Boolean
Public ThickCheckLoaded As Boolean

' SolidWorks Objects
Public swApp As SldWorks.SldWorks
Public swAssembly As AssemblyDoc
Public swBody As SldWorks.Body2
Public swComponent As SldWorks.Component2
Public swConfig As SldWorks.Configuration
Public swConfigMgr As SldWorks.ConfigurationManager
Public swDocSpecification As SldWorks.DocumentSpecification
Public swDraw As SldWorks.DrawingDoc
Public swDrawMod As SldWorks.ModelDoc2
Public swEnt As SldWorks.Entity
Public swEqnMgr As SldWorks.EquationMgr
Public swFace As SldWorks.Face
Public swSurface As SldWorks.Surface
Public swFeat As SldWorks.Feature
Public swPart As SldWorks.PartDoc
Public swMass As SldWorks.MassProperty
Public swModel As SldWorks.ModelDoc2
Public swModelDocExt As SldWorks.ModelDocExtension
Public swSelMgr As SldWorks.SelectionMgr
Public swSelData As SldWorks.SelectData
Public swSheetMetal As SldWorks.SheetMetalFeatureData
Public swVolume As Double
Public swVolumeUP As Double
Public swVolumeDN As Double
Public Thickness               As Double
Public vBodies                 As Variant
Public vComponents             As Variant
Public Views                   As Variant
Public vPaletteNames           As Variant
Public FeatureName             As String
Public vFeatureName()          As String
Public swFeatMgr               As SldWorks.FeatureManager
Public vFeats                  As Variant
Public vOutline                As Variant
Public bolTBC                  As Boolean
Public iPartsTotal             As Integer
Public iPartsComplete          As Integer
Public ImportFiles()           As String
Public swDocTypes()            As Integer
Public gImportInt              As Integer
Public WriteDescription        As Boolean
Public GlobalExcel             As Excel.Application
    Public WallThickness As String
    Public CrossSection As String
    Public TubePrefix As String
Public Const cstrRedoFile As String = "I:\SemiRedoDump.txt"
Public Const cstrOutputFile As String = "I:\SemiModelDump.txt"
Public Const cstrArrayFile As String = "I:\SemiArrayDump.txt"
Public Const cstrQuoteASMFile As String = "I:\SemiQuoteASMDump.txt"
Public Const cstrDocTypeFile As String = "I:\SemiDocTypeDump.txt"
Public Const cstrGUIFile As String = "I:\SemiGUIDump.txt"
Public sTime, eTime, loopSTime, loopETime, sTemp, eTemp, totalTime, averageTime, averageCount, averageLast
Public Const DeburRate As Double = 3600 '60 inches per minute, in hours
 
Public Sub TBC()
    Dim recordCounter, X As Integer
    Dim sLineOfText As String
    'read in lines of file
    recordCounter = 0
    Dim fs As New Scripting.FileSystemObject
    With fs
        If .FileExists(cstrQuoteASMFile) Then
            Open cstrQuoteASMFile For Input As #1
            Do While Not EOF(1)
                Line Input #1, sLineOfText
                If Trim(sLineOfText) <> "" Then
                recordCounter = recordCounter + 1
                End If
            Loop
            Close #1
            ReDim ImportFiles(0 To recordCounter - 1) As String
            ReDim swDocTypes(0 To recordCounter - 1) As Integer
        End If
    End With
    
    recordCounter = 0
    Open cstrOutputFile For Input As #1
    Do While Not EOF(1)
        Line Input #1, sLineOfText
        If Trim(sLineOfText) <> "" Then
            recordCounter = recordCounter + 1
        End If
    Loop
    Close #1
    ReDim ModelNames(0 To recordCounter) As String
    
    If recordCounter = 0 Then
        Debug.Print "No more objects"
        Exit Sub
    End If
    
    Open cstrArrayFile For Input As #1
    Do While Not EOF(1)
        Line Input #1, sLineOfText
        If Trim(sLineOfText) <> "" Then
            recordCounter = recordCounter + 1
        End If
    Loop
    Close #1

    'redim arrays
    
    ReDim ModelNamesRaw(0 To recordCounter) As String
    ReDim ModelNamesRedo(0 To recordCounter) As String
    'read in values into arrays
    X = 0
    With fs
      If .FileExists(cstrQuoteASMFile) Then
        Open cstrQuoteASMFile For Input As #1
        Do While Not EOF(1)
            Line Input #1, sLineOfText
            If Trim(sLineOfText) <> "" Then
                ImportFiles(X) = sLineOfText
                X = X + 1
            End If
        Loop
        Close #1
        X = 0
        Open cstrDocTypeFile For Input As #1
        Do While Not EOF(1)
            Line Input #1, sLineOfText
            If Trim(sLineOfText) <> "" Then
                swDocTypes(X) = sLineOfText
                X = X + 1
            End If
        Loop
        Close #1
      End If
    End With
    X = 0
    Open cstrOutputFile For Input As #1
    Do While Not EOF(1)
        Line Input #1, sLineOfText
        If Trim(sLineOfText) <> "" Then
            ModelNames(X) = sLineOfText
            X = X + 1
        End If
    Loop
    Close #1
    X = 0
    Open cstrArrayFile For Input As #1
    Do While Not EOF(1)
        Line Input #1, sLineOfText
        If Trim(sLineOfText) <> "" Then
            ModelNamesRaw(X) = sLineOfText
            X = X + 1
        End If
    Loop
    Close #1
    X = 0
    Open cstrRedoFile For Input As #1
    Do While Not EOF(1)
        Line Input #1, sLineOfText
        If Trim(sLineOfText) <> "" Then
            ModelNamesRedo(X) = sLineOfText
            X = X + 1
        End If
    Loop
    Close #1
    
    
    bolTBC = True
    Call SP.ReadValues
    If Len(Join(ImportFiles)) > 0 Then 'Needs to run Main over all files in ImportFiles array   Array Exists
        For X = 0 To UBound(ImportFiles)
            Call SP.main
            bolTBC = False
            If X <> UBound(ImportFiles) Then  'open next doc if there are more left to be processed
                Set swModel = swApp.OpenDoc6(ImportFiles(X + 1), swDocTypes(X + 1), swOpenDocOptions_Silent, "Default", longerrors, longwarnings)
            End If
        Next X
    Else
        Call SP.main
    End If
End Sub
Public Sub DumpGUI()
    Open cstrGUIFile For Output As #2
    'frMaterial
    If bQuote = True Then ' Set SemiAutoPilot values to QuoteHelper values
        SemiAutoPilot.rb304.value = False
        SemiAutoPilot.rbALNZD.value = False
        SemiAutoPilot.rb6061.value = False
                Set swModel = swApp.OpenDoc6(ImportFiles(X + 1), swDocTypes(X + 1), swOpenDocOptions_Silent, "Default", longerrors, longwarnings)
            End If
        Next X
    Else
        Call SP.main
    End If
End Sub
Public Sub DumpGUI()
    Open cstrGUIFile For Output As #2
            Set swModel = swApp.OpenDoc6(ImportFiles(X + 1), swDocTypes(X + 1), swOpenDocOptions_Silent, "Default", longerrors, longwarnings)
        End If
    Next X
End Sub

Public Sub DumpGUI()
    Open cstrGUIFile For Output As #2
    'frMaterial
    If bQuote = True Then ' Set SemiAutoPilot values to QuoteHelper values
        SemiAutoPilot.rb304.value = False
        SemiAutoPilot.rbALNZD.value = False
        SemiAutoPilot.rb6061.value = False
        SemiAutoPilot.rb316.value = False
        SemiAutoPilot.rbA36.value = False
        SemiAutoPilot.rb5052.value = False
        SemiAutoPilot.rb309.value = False
        SemiAutoPilot.rb2205.value = False
        SemiAutoPilot.rbC22.value = False
        SemiAutoPilot.rbAL6XN.value = False
        SemiAutoPilot.rbOther.value = False
        SemiAutoPilot.rbBendTable.value = False
        SemiAutoPilot.rbKFactor.value = False
        
        If QuoteHelper.rb304.value = True Then
            SemiAutoPilot.rb304.value = True
        ElseIf QuoteHelper.rbALNZD.value = True Then
            SemiAutoPilot.rbALNZD.value = True
        ElseIf QuoteHelper.rb6061.value = True Then
            SemiAutoPilot.rb6061.value = True
        ElseIf QuoteHelper.rb316.value = True Then
            SemiAutoPilot.rb316.value = True
        ElseIf QuoteHelper.rbA36.value = True Then
            SemiAutoPilot.rbA36.value = True
        ElseIf QuoteHelper.rb5052.value = True Then
            SemiAutoPilot.rb5052.value = True
        ElseIf QuoteHelper.rb309.value = True Then
            SemiAutoPilot.rb309.value = True
        ElseIf QuoteHelper.rb2205.value = True Then
            SemiAutoPilot.rb2205.value = True
        ElseIf QuoteHelper.rbC22.value = True Then
            SemiAutoPilot.rbC22.value = True
        ElseIf QuoteHelper.rbAL6XN.value = True Then
            SemiAutoPilot.rbAL6XN.value = True
        ElseIf QuoteHelper.rbOther.value = True Then
            SemiAutoPilot.rbOther.value = True
        End If
    'frBendDeduction
        If QuoteHelper.rbBendTable.value = True Then
            SemiAutoPilot.rbBendTable.value = True
        ElseIf QuoteHelper.rbKFactor.value = True Then
            SemiAutoPilot.rbKFactor.value = True
            SemiAutoPilot.tbKFactor = QuoteHelper.tbKFactor
        End If
    'frDrawings
        If QuoteHelper.cbCreateDXF.value = True Then SemiAutoPilot.cbCreateDXF.value = True
        If QuoteHelper.cbCreateDrawing.value = True Then SemiAutoPilot.cbCreateDrawing.value = True
        If QuoteHelper.obDim.value = True Then SemiAutoPilot.obDim.value = True
        If QuoteHelper.cbVisible.value = True Then SemiAutoPilot.cbVisible.value = True
    End If  ' Dump Values for SemiAutoPilot
    
    If SemiAutoPilot.rb304.value = True Then
        Print #2, "rb304,1"
    ElseIf SemiAutoPilot.rbALNZD.value = True Then
        Print #2, "rbALNZD,1"
    ElseIf SemiAutoPilot.rb6061.value = True Then
        Print #2, "rb6061,1"
    ElseIf SemiAutoPilot.rb316.value = True Then
        Print #2, "rb316,1"
    ElseIf SemiAutoPilot.rbA36.value = True Then
        Print #2, "rbA36,1"
    ElseIf SemiAutoPilot.rb5052.value = True Then
        Print #2, "rb5052,1"
    ElseIf SemiAutoPilot.rb309.value = True Then
        Print #2, "rb309,1"
    ElseIf SemiAutoPilot.rb2205.value = True Then
        Print #2, "rb2205,1"
    ElseIf SemiAutoPilot.rbC22.value = True Then
        Print #2, "rbC22,1"
    ElseIf SemiAutoPilot.rbAL6XN.value = True Then
        Print #2, "rbAL6XN,1"
    ElseIf SemiAutoPilot.rbOther.value = True Then
        Print #2, "rbOther,1"
        If OtherMaterial.rbOtherSS.value = True Then
            Print #2, "rbOtherSS,1"
        ElseIf OtherMaterial.rbOtherCS.value = True Then
            Print #2, "rbOtherCS,1"
        ElseIf OtherMaterial.rbOtherAlum.value = True Then
            Print #2, "rbOtherAlum,1"
        End If
    End If
    'frBendDeduction
    If SemiAutoPilot.rbBendTable.value = True Then
        Print #2, "rbBendTable,1"
    ElseIf SemiAutoPilot.rbKFactor.value = True Then
        Print #2, "rbKFactor,1"
        Print #2, "tbKFactor," & SemiAutoPilot.tbKFactor
    End If
    'frDrawings
    If SemiAutoPilot.cbCreateDXF.value = True Then Print #2, "cbCreateDXF,1"
    If SemiAutoPilot.cbCreateDrawing.value = True Then Print #2, "cbCreateDrawing,1"
    If SemiAutoPilot.obDim.value = True Then Print #2, "obDim,1"
    If SemiAutoPilot.cbVisible.value = True Then Print #2, "cbVisible,1"
    If CustomProps.cbGrain.value = True Then Print #2, "cbGrain,1"
    If CustomProps.cbCommon.value = True Then Print #2, "cbCommon,1"
    If CustomProps.tbCustomer.value <> "" Then Print #2, "tbCustomer," & CustomProps.tbCustomer.value
    If CustomProps.tbPrint.value <> "" Then Print #2, "tbPrint," & CustomProps.tbPrint.value
    If CustomProps.cbPrintFromPart.value = True Then Print #2, "cbPrintFromPart," & CustomProps.cbPrintFromPart.value
    If CustomProps.tbRevision.value <> "" Then Print #2, "tbRevision," & CustomProps.tbRevision.value
    If CustomProps.tbDescription.value <> "" Then Print #2, "tbDescription," & CustomProps.tbDescription.value
    Print #2, "bQuote," & bQuote
    Print #2, "bQuoteASM," & bQuoteASM
    Print #2, "filePath," & filePath
    Close #2
End Sub
Public Function FindValue(ByVal strText As String) As String
    Dim sLineOfText As String
    Open cstrGUIFile For Input As #2
    While Not EOF(2)
        Line Input #2, sLineOfText ' read in data 1 line at a time
        If UCase(Left(sLineOfText, Len(strText))) = UCase(strText) Then
            FindValue = sLineOfText
            Close #2
            Exit Function
        End If
    Wend
    Close #2
End Function
Public Sub ReadValues()
    Dim sLineOfText As String
    Dim guiVar As String
    Open cstrGUIFile For Input As #2
    While Not EOF(2)
        Line Input #2, sLineOfText ' read in data 1 line at a time
        guiVar = Left(sLineOfText, InStr(sLineOfText, ",") - 1)
        Select Case guiVar
            Case "rb304"
                SemiAutoPilot.rb304.value = True
            Case "rbALNZD"
                SemiAutoPilot.rbALNZD.value = True
            Case "rb316"
                SemiAutoPilot.rb316.value = True
            Case "rbA36"
                SemiAutoPilot.rbA36.value = True
            Case "rb6061"
                SemiAutoPilot.rb6061.value = True
            Case "rb5052"
                SemiAutoPilot.rb5052.value = True
            Case "rb309"
                SemiAutoPilot.rb309.value = True
            Case "rb2205"
                SemiAutoPilot.rb2205.value = True
            Case "rbC22"
                SemiAutoPilot.rbC22.value = True
            Case "rbAL6XN"
                SemiAutoPilot.rbAL6XN.value = True
            Case "rbOther"
                SemiAutoPilot.rbOther.value = True
            Case "rbOtherCS"
                OtherMaterial.rbOtherCS.value = True
            Case "rbOtherSS"
                OtherMaterial.rbOtherSS.value = True
                SemiAutoPilot.tbKFactor.value = Mid(sLineOfText, InStrRev(sLineOfText, ",") + 1, Len(sLineOfText) - InStrRev(sLineOfText, ","))
            Case "cbCreateDXF"
                SemiAutoPilot.cbCreateDXF.value = True
            Case "cbCreateDrawing"
                SemiAutoPilot.cbCreateDrawing.value = True
            Case "obDim"
                SemiAutoPilot.obDim.value = True
                SemiAutoPilot.tbKFactor.value = Mid(sLineOfText, InStrRev(sLineOfText, ",") + 1, Len(sLineOfText) - InStrRev(sLineOfText, ","))
            Case "cbCreateDXF"
                SemiAutoPilot.cbCreateDXF.value = True
            Case "cbCreateDrawing"
                SemiAutoPilot.cbCreateDrawing.value = True
            Case "obDim"
                SemiAutoPilot.obDim.value = True
                CustomProps.tbCustomer.value = Mid(sLineOfText, InStrRev(sLineOfText, ",") + 1, Len(sLineOfText) - InStrRev(sLineOfText, ","))
            Case "tbPrint"
                CustomProps.tbPrint.value = Mid(sLineOfText, InStrRev(sLineOfText, ",") + 1, Len(sLineOfText) - InStrRev(sLineOfText, ","))
            Case "cbPrintFromPart"
                CustomProps.cbPrintFromPart.value = Mid(sLineOfText, InStrRev(sLineOfText, ",") + 1, Len(sLineOfText) - InStrRev(sLineOfText, ","))
            Case "tbRevision"
                CustomProps.tbRevision.value = Mid(sLineOfText, InStrRev(sLineOfText, ",") + 1, Len(sLineOfText) - InStrRev(sLineOfText, ","))
            Case "tbDescription"
                CustomProps.tbDescription.value = Mid(sLineOfText, InStrRev(sLineOfText, ",") + 1, Len(sLineOfText) - InStrRev(sLineOfText, ","))
                CustomProps.tbPrint.value = Mid(sLineOfText, InStrRev(sLineOfText, ",") + 1, Len(sLineOfText) - InStrRev(sLineOfText, ","))
            Case "cbPrintFromPart"
                CustomProps.cbPrintFromPart.value = Mid(sLineOfText, InStrRev(sLineOfText, ",") + 1, Len(sLineOfText) - InStrRev(sLineOfText, ","))
            Case "tbRevision"
                CustomProps.tbRevision.value = Mid(sLineOfText, InStrRev(sLineOfText, ",") + 1, Len(sLineOfText) - InStrRev(sLineOfText, ","))
            Case "tbDescription"
                CustomProps.tbDescription.value = Mid(sLineOfText, InStrRev(sLineOfText, ",") + 1, Len(sLineOfText) - InStrRev(sLineOfText, ","))
            Case "bQuote"
                bQuote = Mid(sLineOfText, InStrRev(sLineOfText, ",") + 1, Len(sLineOfText) - InStrRev(sLineOfText, ","))
            Case "bQuoteASM"
                bQuoteASM = Mid(sLineOfText, InStrRev(sLineOfText, ",") + 1, Len(sLineOfText) - InStrRev(sLineOfText, ","))
            Case "filePath"
                filePath = Mid(sLineOfText, InStrRev(sLineOfText, ",") + 1, Len(sLineOfText) - InStrRev(sLineOfText, ","))
            Case Else
                    
        End Select
    Wend
    Close #2
    'Run btnOK_Click to set specific values based off prior GUI settings
    Call SemiAutoPilot.btnOK_Click
End Sub

Sub QuoteStart()

Set fso = CreateObject("Scripting.FileSystemObject")
Dim MyFiles As Scripting.Files
Set MyFiles = fso.GetFolder(filePath).Files
Dim MyFile As Scripting.File
Dim longerrors As Long
Dim longwarnings As Long
Dim swImportData As SldWorks.ImportIgesData
Dim Err As Long
Set swApp = Application.SldWorks
N = 0

''swApp.DocumentVisible False, swDocPART
swApp.UserControl = False
swApp.Visible = False
swApp.Frame.KeepInvisible = True

For Each MyFile In MyFiles
    If UCase(Right(MyFile.Name, 3)) = "IGS" Or UCase(Right(MyFile.Name, 4)) = "STEP" Or UCase(Right(MyFile.Name, 3)) = "SAT" Or UCase(Right(MyFile.Name, 3)) = "STP" Then
        Set swModel = swApp.LoadFile4(MyFile.Path, "r", swImportData, Err)
        modelType = swModel.GetType
        DocName = swModel.GetTitle
        If modelType = swDocASSEMBLY Then
            bRet = swModel.SaveAs4(filePath & DocName & ".sldasm", swSaveAsCurrentVersion, swSaveAsOptions_Silent, longerrors, longwarnings)
        ElseIf modelType = swDocPART Then
            bRet = swModel.SaveAs4(filePath & DocName & ".sldprt", swSaveAsCurrentVersion, swSaveAsOptions_Silent, longerrors, longwarnings)
        End If
        swApp.CloseAllDocuments True
    End If
Next MyFile

Set MyFiles = Nothing
Set MyFiles = fso.GetFolder(filePath).Files

For Each MyFile In MyFiles
    If UCase(Right(MyFile.Name, 6)) = "SLDPRT" Then
        ReDim ModelNames(0 To N) As String
        ReDim ModelNamesRaw(0 To N) As String
        ReDim ModelNamesRedo(0 To N) As String
        N = N + 1
    End If
Next MyFile

N = 0

For Each MyFile In MyFiles
    If UCase(Right(MyFile.Name, 6)) = "SLDPRT" Then
        ModelNames(N) = MyFile.Name
        N = N + 1
    End If
Next MyFile

iPartsTotal = N
iPartsComplete = 0

N = 0

bQuote = True

''swApp.Visible = True
''swApp.FrameHeight = 0
''swApp.FrameLeft = 0
''swApp.FrameTop = 0
''swApp.FrameWidth = 0
''swApp.FrameState = swWindowNormal

swApp.LoadAddIn "C:/Program Files/NorthernMfg/ExtractDataInstaller/SWExtractDataAddin.DLL"
Call SP.main
frmLoading.Hide
'frmLoading.WindowsMediaPlayer1.Close
frmComplete.Show

swApp.FrameHeight = 1000
swApp.FrameLeft = 0
swApp.FrameTop = 0
swApp.FrameWidth = 1800
swApp.FrameState = swWindowNormal

swApp.FrameState = swWindowMaximized

End

End Sub

Sub QuoteStartASM()

Set fso = CreateObject("Scripting.FileSystemObject")
Dim MyFiles As Scripting.Files
Set MyFiles = fso.GetFolder(filePath).Files
Dim MyFile As Scripting.File
Dim longerrors As Long
Dim longwarnings As Long
Dim swImportData As SldWorks.ImportIgesData
Dim Err As Long

Set swApp = Application.SldWorks
N = 0

''swApp.DocumentVisible False, swDocPART
swApp.UserControl = False
swApp.Visible = False
swApp.Frame.KeepInvisible = True

bQuote = False

For Each MyFile In MyFiles
    If UCase(Right(MyFile.Name, 3)) = "IGS" Or UCase(Right(MyFile.Name, 4)) = "STEP" Or UCase(Right(MyFile.Name, 3)) = "SAT" Or UCase(Right(MyFile.Name, 3)) = "STP" Then
        N = N + 1
    End If
Next MyFile
ReDim ImportFiles(0 To N) As String 'leave extra room for FindLastValue
ReDim swDocTypes(0 To N) As Integer

N = 0

For Each MyFile In MyFiles
    If UCase(Right(MyFile.Name, 3)) = "IGS" Or UCase(Right(MyFile.Name, 4)) = "STEP" Or UCase(Right(MyFile.Name, 3)) = "SAT" Or UCase(Right(MyFile.Name, 3)) = "STP" Then
        Set swModel = swApp.LoadFile4(MyFile.Path, "r", swImportData, Err)
        modelType = swModel.GetType
        DocName = swModel.GetTitle
        If modelType = swDocASSEMBLY Then
            bRet = swModel.SaveAs4(filePath & DocName & ".sldasm", swSaveAsCurrentVersion, swSaveAsOptions_Silent, longerrors, longwarnings)
            ImportFiles(N) = filePath & DocName & ".sldasm"
            swDocTypes(N) = swDocASSEMBLY
            N = N + 1
        ElseIf modelType = swDocPART Then
            bRet = swModel.SaveAs4(filePath & DocName & ".sldprt", swSaveAsCurrentVersion, swSaveAsOptions_Silent, longerrors, longwarnings)
            ImportFiles(N) = filePath & DocName & ".sldprt"
            swDocTypes(N) = swDocPART
            N = N + 1
        End If
        swApp.CloseAllDocuments True
    End If
Next MyFile

Set MyFiles = Nothing

bQuoteASM = True

''swApp.Visible = True
''swApp.FrameHeight = 0
''swApp.FrameLeft = 0
''swApp.FrameTop = 0
''swApp.FrameWidth = 0
''swApp.FrameState = swWindowNormal
swApp.LoadAddIn "C:/Program Files/NorthernMfg/ExtractDataInstaller/SWExtractDataAddin.DLL"
For gImportInt = 0 To UBound(ImportFiles) - 1
    Set swModel = swApp.OpenDoc6(ImportFiles(gImportInt), swDocTypes(gImportInt), swOpenDocOptions_Silent, "Default", longerrors, longwarnings)
    Call SP.main
Next gImportInt

N = 0

MsgBox "Complete"

swApp.FrameHeight = 1000
swApp.FrameLeft = 0
swApp.FrameTop = 0
swApp.FrameWidth = 1800
swApp.FrameState = swWindowNormal

swApp.FrameState = swWindowMaximized

End

End Sub
Sub Initialize()
    'This procedure runs for the non quote-helper version of AutoPilot.
    'It initializes the ModelNames array and some other parameters.
    
    Dim MyFiles As Scripting.Files
    Dim MyFile As Scripting.File
    Dim zInt As Integer
    
    If bQuote = False Or bQuoteASM = True Then
        Set swModel = swApp.ActiveDoc
        On Error Resume Next
        saveBol = swModel.Save3(5, longerrors, longwarnings)
        On Error GoTo 0
        If swModel Is Nothing Then
            MsgBox "No model loaded."
            End
        ElseIf saveBol = False Then
            MsgBox "Assembly is read only. Please take ownership before starting."
            swApp.UserControl = True
            swApp.Visible = True
            End
        End If
        
        On Error Resume Next
        Set swAssembly = swModel
        If Not swAssembly Is Nothing Then
            filePath = swModel.GetPathName
            swApp.CloseAllDocuments False
            swApp.DocumentVisible False, swDocASSEMBLY
            swApp.UserControl = False
            swApp.Visible = False
            swApp.Frame.KeepInvisible = True
            Set swModel = swApp.OpenDoc6(filePath, swDocASSEMBLY, swOpenDocOptions_Silent, "", longerrors, longwarnings)
        Else
            filePath = swModel.GetPathName
            swApp.CloseAllDocuments False
            swApp.UserControl = False
            swApp.Visible = False
            swApp.Frame.KeepInvisible = True
            Set swModel = swApp.OpenDoc6(filePath, swDocPART, swOpenDocOptions_Silent, "", longerrors, longwarnings)
        End If

        On Error Resume Next
        Set swAssembly = swModel
        BasePath = swApp.GetCurrentMacroPathFolder
        filePath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\"))
        MainAssembly = swModel.GetTitle
        If Not swAssembly Is Nothing Then
            vComponents = swAssembly.GetComponents(False)
        Else
            ReDim vComponents(0) As Variant
            Set vComponents(0) = swModel
        End If
        iPartsTotal = UBound(vComponents) + 1
        iPartsComplete = 0
        On Error Resume Next
        If Not bolTBC = True Then   'initialize
            If Not bQuoteASM = True Then
                On Error GoTo 0
                SemiAutoPilot.Show
                On Error Resume Next
                If strBendTable <> "-1" Then
                    Set objExcel = GlobalExcel
                    objExcel.Visible = False
                    Set objBendTable = objExcel.Workbooks.Open(strBendTable)
                End If
            End If
            
            ReDim ModelNames(0 To 0) As String
            
            Dim intMN As Integer
            Dim intVC As Integer

            For intVC = 0 To UBound(vComponents)
                If Not swAssembly Is Nothing Then
                    If vComponents(intVC).GetModelDoc2.GetType = swDocPART Then
                        ReDim Preserve ModelNames(0 To intMN)
                        ModelNames(intMN) = vComponents(intVC).GetModelDoc2.GetPathName
                        intMN = intMN + 1
                    End If
                Else
                    ModelNames(intVC) = MainAssembly
                End If
            Next intVC
            
            ReDim ModelNamesRaw(0 To UBound(ModelNames))
            ReDim ModelNamesRedo(0 To UBound(ModelNames))
            
        End If
        Set vComponents = Nothing
        
    ElseIf bolTBC = True Then
        'iPartsTotal and iPartsComplete need to be set here for restart on bQuote = True and bolTBC = True
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set MyFiles = fso.GetFolder(filePath).Files
        zInt = 0
        For Each MyFile In MyFiles
            If UCase(Right(MyFile.Name, 6)) = "SLDPRT" Then
            zInt = zInt + 1
            End If
        Next MyFile
        iPartsTotal = zInt
        'iPartsComplete = iPartsTotal - FindLastValue(ModelNames())
    End If
End Sub
Public Function RestartCheck(ByVal iLoop As Integer, ByVal sCheck As Integer) As Boolean
    If averageCount > 0 Then
        Debug.Print "i :: " & iLoop & " :: AverageTime :: " & Round(averageTime / averageCount, 2)
        If averageTime / averageCount >= sCheck Then
            'output remaining model names to text file
            Open cstrOutputFile For Output As #1
            For j = iLoop To UBound(ModelNames)
                Print #1, ModelNames(j)
            Next j
            Close #1
            'out ModelNamesRaw array so do not repeat duplicates
            Open cstrArrayFile For Output As #1
            For j = 0 To UBound(ModelNamesRaw)
                Print #1, ModelNamesRaw(j)
            Next j
            Close #1
            'output Redo model names to text file to run second loop of parts not completed properly
            Open cstrRedoFile For Output As #1
            For j = 0 To UBound(ModelNamesRedo)
                Print #1, ModelNamesRedo(j)
            Next j
            Close #1
            If Len(Join((ImportFiles))) > 0 Then  'check for ImportFiles to not be empty
                If FindLastValue(ImportFiles()) > 0 Then   'determine if it is running for Quote ASM and dump arrays if needed
                    Open cstrQuoteASMFile For Output As #1
                    For j = gImportInt To UBound(ImportFiles) 'global variables needed to not have to send along with main function
                        Print #1, ImportFiles(j)
                    Next j
                    Close #1
                    Open cstrDocTypeFile For Output As #1
                    For j = gImportInt To UBound(swDocTypes) - 1  'global variables needed to not have to send along with main function
                        Print #1, swDocTypes(j)
                    Next j
                    Close #1
                End If
            End If
            Call DumpGUI
            ProgName = "O:\Engineering Department\Solidworks\Macros\(Semi)Autopilot\Semi\RestartSolidworks.exe"
            If bQuote = False Then
                ProgParam = filePath & MainAssembly  ' pass file name when wanting to reopen a file on SW restart
            Else
                ProgParam = "False"  'pass False when not wanting to open a file on SW restart
            End If
            Debug.Print ProgName & " """ & ProgParam & """"
            RestartCheck = True
            Call Shell(ProgName & " """ & ProgParam & """", vbNormalFocus)
            Exit Function
        End If
    End If
    RestartCheck = False
    If averageCount > 0 Then averageLast = Round((averageTime / averageCount), 2)
    averageTime = 0: averageCount = 0
End Function
Public Function NumberOfBodies() As Integer
    vBodies = swPart.GetBodies2(swAllBodies, True)
    If Not IsEmpty(vBodies) Then
        Set swBody = Nothing
        Set swBody = vBodies(0)
        NumberOfBodies = UBound(vBodies) + 1
    Else
        NumberOfBodies = 0
        Debug.Print "::Error:: Active part had no bodies"
    End If
End Function
Public Function UnsuppressFlatten() As Boolean
    For u = LBound(Features) To UBound(Features)
        Feature = Features(u).GetTypeName
        Set swFeat = Features(u)
        If Feature = "UiBend" Or Feature = "FlatPattern" Then
            swFeat.SetSuppression2 swUnSuppressFeature, swAllConfiguration, Nothing
        End If
    Next u
End Function

Public Function SMInsertBends() As Boolean
    Dim swVolume1 As Double
    Dim swVolume2 As Double
    Dim InsertType As String
    Features = swBody.GetFeatures
    Set swFeat = Nothing
    Set objFeat = Nothing
    For j = LBound(Features) To UBound(Features)
        Feature = Features(j).GetTypeName
        If Feature = "SheetMetal" Then
            Set objFeat = Features(j)
            SMInsertBends = True
            UnsuppressFlatten
            'swApp.DocumentVisible True, swDocPART
            'swApp.ActivateDoc ModelName & ".sldprt"
            Exit Function
        End If
    Next j
'-------------------------------------------------------
'INSERT BENDS - Make Sheetmetal feature
'-------------------------------------------------------
    Call GetLargestFace
    SheetMetal = False
    FaceType = swSurface.Identity
    InsertType = ""
    If FaceType = 4001 Then
        'swApp.ActivateDoc ModelName & ".sldprt"
        Set swMass = swModel.Extension.CreateMassProperty
        swVolume1 = Round(swMass.Volume, 15)
        SheetMetal = False
        On Error Resume Next
        SheetMetal = swPart.InsertBends2(0.001, -1, 0.5, -1, True, 1, True)
        
        If Not SheetMetal = False Then
            InsertType = "InsertBendFace"
        Else
            'SheetMetal = swModel.FeatureManager.InsertConvertToSheetMetal(0.005, False, False, 0.005, 0.001, swSheetMetalReliefTear, 0.5)
            'swModel.ForceRebuild3 True
            'InsertType = "Convert"
        End If
        
        If Not SheetMetal = False Then
            Thickness = 0
            swModel.DeleteCustomInfo "SMThick"
            swModel.AddCustomInfo3 gstrConfigName, "SMThick", 30, """Thickness@$PRP:""SW-File Name"".SLDPRT"""
            strThickness = swModel.GetCustomInfoValue(gstrConfigName, "SMThick")
            badThickness = "SLDPRT" & """"
            strThicknessCheck = UCase(Right(strThickness, 7))
            If strThicknessCheck <> badThickness Then
                Thickness = strThickness * 0.0254
            End If
            swModel.DeleteCustomInfo "SMThick"
        End If
        
        Set swMass = swModel.Extension.CreateMassProperty
        swVolume2 = Round(swMass.Volume, 15)
        swVolumeUP = swVolume2 * 1.005
        swVolumeDN = swVolume2 * 0.995
        bVolumeUP = swVolumeUP > swVolume1
        bVolumeDN = swVolumeDN < swVolume1
        If bVolumeUP = False Or bVolumeDN = False Then
            swModel.EditUndo2 (2)
            SheetMetal = False
            InsertType = ""
        End If
        Debug.Print "SheetMetal = " & SheetMetal
    Else
        Call GetLinearEdge
        Set swMass = swModel.Extension.CreateMassProperty
        swVolume1 = Round(swMass.Volume, 15)
        SheetMetal = False
        On Error Resume Next
        SheetMetal = swPart.InsertBends2(0.001, -1, 0.5, -1, True, 1, True)

        If Not SheetMetal = False Then
            Thickness = 0
            swModel.DeleteCustomInfo "SMThick"
            swModel.AddCustomInfo3 gstrConfigName, "SMThick", 30, """Thickness@$PRP:""SW-File Name"".SLDPRT"""
            strThickness = swModel.GetCustomInfoValue(gstrConfigName, "SMThick")
            badThickness = "SLDPRT" & """"
            strThicknessCheck = UCase(Right(strThickness, 7))
            If strThicknessCheck <> badThickness Then
                Thickness = strThickness * 0.0254
            End If
            swModel.DeleteCustomInfo "SMThick"
            InsertType = "InsertBendEdge"
        End If
        
        Set swMass = swModel.Extension.CreateMassProperty
        swVolume2 = Round(swMass.Volume, 15)
        swVolumeUP = swVolume2 * 1.005
        swVolumeDN = swVolume2 * 0.995
        bVolumeUP = swVolumeUP > swVolume1
        bVolumeDN = swVolumeDN < swVolume1
        If bVolumeUP = False Or bVolumeDN = False Then
            swModel.EditUndo2 (2)
            SheetMetal = False
            InsertType = ""
        End If
        Debug.Print "SheetMetal = " & SheetMetal
    End If
           
    UnsuppressFlatten
    
    If Not SheetMetal = False Then
        Call ValidateFlatPattern
        If CompareMass() = True Then
            swModel.EditUndo2 (2)
            swModel.EditRebuild3
            Call NumberOfBodies
            
            If InsertType = "InsertBendFace" Then
                Call GetLargestFace
                If strBendTable <> -1 Then
                    swApp.FrameHeight = 0
                    swApp.FrameLeft = 0
                    swApp.FrameTop = 0
                    swApp.FrameWidth = 0
                    swApp.FrameState = swWindowMinimized
                    swApp.ActivateDoc ModelName & ".sldprt"
                End If
                
                SheetMetal = swPart.InsertBends2(Thickness, strBendTable, KFactor, -1, True, 1, True)
                
                If strBendTable <> -1 Then
                    swApp.FrameState = swWindowMinimized
                    swApp.FrameHeight = 1000
                    swApp.FrameLeft = 0
                    swApp.FrameTop = 0
                    swApp.FrameWidth = 1800
                    swModel.Save3 5, longerrors, longwarnings
                    swApp.CloseDoc FileName
                    swApp.UserControl = False
                    swApp.Visible = False
                    swApp.Frame.KeepInvisible = True
                    Set swModel = swApp.OpenDoc6(FileName, swDocPART, 1, "", longerrors, longwarnings)
                    Set swPart = swModel
                    Progress.Show (False)
                End If

            ElseIf InsertType = "InsertBendEdge" Then
                Call GetLinearEdge
                If strBendTable <> -1 Then
                    swApp.FrameHeight = 0
                    swApp.FrameLeft = 0
                    swApp.FrameTop = 0
                    swApp.FrameWidth = 0
                    swApp.FrameState = swWindowMinimized
                    swApp.ActivateDoc ModelName & ".sldprt"
                End If
                            
                SheetMetal = swPart.InsertBends2(Thickness, strBendTable, KFactor, -1, True, 1, True)
                
                If strBendTable <> -1 Then
                    swApp.FrameState = swWindowMinimized
                    swApp.FrameHeight = 1000
                    swApp.FrameLeft = 0
                    swApp.FrameTop = 0
                    swApp.FrameWidth = 1800
                    swModel.Save3 5, longerrors, longwarnings
                    swApp.CloseDoc FileName
                    swApp.UserControl = False
                    swApp.Visible = False
                    swApp.Frame.KeepInvisible = True
                    Set swModel = swApp.OpenDoc6(FileName, swDocPART, 1, "", longerrors, longwarnings)
                    Set swPart = swModel
                    Progress.Show (False)
                End If

            ElseIf InsertType = "Convert" Then
                Call GetLargestFace
                If strBendTable <> -1 Then
                    ''swApp.DocumentVisible True, swDocPART
                    swApp.ActivateDoc ModelName & ".sldprt"
                End If
                SheetMetal = swModel.FeatureManager.InsertConvertToSheetMetal2(0.001, False, True, 0.001, 0.001, 3, 0.5, 1, 0.5, False)
                Call BendAllowanceType
            End If
            
            SMInsertBends = True
            UnsuppressFlatten
            Call SaveCurrentModel
            'swApp.CloseDoc ModelName & ".sldprt"
            ''swApp.DocumentVisible False, swDocPART
        Else
            SheetMetal = False
            SMInsertBends = False
            'swModel.EditUndo2 (1)
            Call SaveCurrentModel
        End If
    Else
        SMInsertBends = False
    End If
    
    swPart.ClearSelection2 True
End Function
Sub FindFlatPattern()
    'Find Flat Pattern feature and exit if you find it
    
    boolstatus = swModel.Extension.SelectByID2("Flat-Pattern*", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
    Set swFeat = swModel.FirstFeature
    vFeats = swFeatMgr.GetFeatures(False)
    ReDim vFeatureName(0 To UBound(vFeats)) As String

    For f = LBound(vFeats) To UBound(vFeats)
        Set swFeat = Nothing
        Set swFeat = vFeats(f)
        FeatureName = swFeat.GetTypeName
        vFeatureName(f) = FeatureName
    
        If FeatureName = "FlatPattern" Then Exit Sub
        'Set swFeat = Nothing
        Set swFeat = swFeat.GetNextFeature
        If Not swFeat Is Nothing Then
            FeatureName = swFeat.GetTypeName
        End If
    Next f
End Sub
Public Function CompareMass() As Boolean
    calcVolume = biggestArea * Thickness
    'Set swModel = Nothing
    'Set swModel = swApp.ActiveDoc
    Set swMass = swModel.Extension.CreateMassProperty
    swVolume = swMass.Volume
    swVolumeUP = swVolume * 1.03
    swVolumeDN = swVolume * 0.97
    bVolumeUP = swVolumeUP > calcVolume
    bVolumeDN = swVolumeDN < calcVolume
    If bVolumeUP = False Or bVolumeDN = False Then
        swModel.EditUndo2 (2)
        CompareMass = False
        Exit Function
    End If
    CompareMass = True
End Function
Public Function FindLastValue(ByRef mArray() As String) As Integer
    Dim zInt As Integer
    zInt = 0
    While Trim(mArray(zInt)) <> "" And Not IsNull(Trim(mArray(zInt)))
        zInt = zInt + 1
    Wend
    FindLastValue = zInt
End Function
Sub ProcessModel()
    
    Dim fs As New Scripting.FileSystemObject
    Dim AssemFile As String
    boolFindSheetMetal = False
    WriteDescription = False
    
    'put the model name in the array so it does not get run twice
    ModelNamesRaw(FindLastValue(ModelNamesRaw())) = "[" & ModelNameRaw & "]"   'Finds next blank in ModelNamesRaw array
    If Trim(ModelNameRaw) = "" Or IsNull(Trim(ModelNameRaw)) Then
        Debug.Print "ModelNameRaw being set as blank!!!"
    End If
    ModelName = ModelNameRaw
    FileName = filePath & ModelName & ".sldprt"
    AssemFile = filePath & ModelName & ".sldasm"
    
    'swApp.DocumentVisible True, swDocPART
    'swApp.UserControl = True
    'swApp.Visible = True
    
    'make the current part active
    If bQuote = True Then 'running for all files in a folder
        'swApp.CloseAllDocuments True
        If fs.FileExists(FileName) Then Set swModel = swApp.OpenDoc6(FileName, 1, 1, "Default", longerrors, longwarnings)
    ElseIf bQuote = False Then 'running over the current assembly of parts
        Set swModel = Nothing
        'Set swModel = swApp.ActivateDoc(ModelName & ".sldprt")
        If fs.FileExists(FileName) Then Set swModel = swApp.OpenDoc6(FileName, swDocPART, 0, "", longerrors, longwarnings)
        If swModel Is Nothing Then
            'check if this was sub assembly
            With fs
            If .FileExists(AssemFile) Then
                Debug.Print "Skipping Sub Assembly :: " & AssemFile
                boolFindSheetMetal = True 'Skip redo on this file bc it is a sub assembly, not a part
            Else
                Debug.Print "::Check:: if Lightweight mode OR Large Assembly mode are active"
            End If
            End With
            'Set swModel = swApp.OpenDoc6(FileName, 1, 1, "Default", longerrors, longwarnings)
        Else
            swModel.ShowConfiguration2 ("Default")
        End If
    End If
    
    If Not swModel Is Nothing Then 'found active part and non-standard part
      If swModel.GetCustomInfoValue(gstrConfigName, "rbPartType") = "0" Then  'skip if manually changed
      'And swModel.GetCustomInfoValue(gstrConfigName, "Shape") = ""
'-------------------------------------------------------
'MATERIAL TYPE
'-------------------------------------------------------
        Set swPart = Nothing
        Set swPart = swModel
        swApp.CopyAppearance swPart
        swPart.SetMaterialPropertyName2 "Default", "C:/Program Files/SolidWorks Corp/SolidWorks/lang/english/sldmaterials/SolidWorks Materials.sldmat", MaterialType
        swApp.PasteAppearance swPart, 3
'-------------------------------------------------------
'CHECK FOR PREVIOUS RUN
'-------------------------------------------------------
        CheckOP20 = ""
        gstrConfigName = ""
        CheckOP20 = swModel.GetCustomInfoValue(gstrConfigName, "OP20")
        If CheckOP20 = "" Or CheckOP20 = "N115 - ANY FLAT LASER" Or CheckOP20 = "N120 - 5040" Or CheckOP20 = "N125 - 3060" Then
            BodyCount = NumberOfBodies() 'function to find if there is 1 or more bodies
            If BodyCount = 1 Then
                boolFindSheetMetal = SMInsertBends()
'-------------------------------------------------------
'CHECK FOR VALID FLAT PATTERN
'-------------------------------------------------------
                If boolFindSheetMetal = True Then
                    'Set swModel = swApp.OpenDoc6(FileName, 1, 1, "Default", longerrors, longwarnings)
'------------------------------------------------------------------------------------------
'SET OP20 TO 115
'------------------------------------------------------------------------------------------
                    If CheckOP20 = "" Then
                        OP20 = "N120 - 5040"
                        On Error Resume Next
                        swModel.CustomInfo2(gstrConfigName, "OP20") = OP20
                        If Error <> 0 Then
                            Set swModel = Nothing
                            Set swModel = swApp.OpenDoc6(FileName, 1, 1, "Default", longerrors, longwarnings)
                            swModel.CustomInfo2(gstrConfigName, "OP20") = OP20
                        End If
                    End If
                        On Error Resume Next
'------------------------------------------------------------------------------------------
'MATERIAL
'------------------------------------------------------------------------------------------
                    OptiMaterial = swModel.GetCustomInfoValue(gstrConfigName, "OptiMaterial")
                    If OptiMaterial = "" Then
                        Call SetMaterial
                        swModel.CustomInfo2(gstrConfigName, "OptiMaterial") = OptiMaterial
                    End If
'------------------------------------------------------------------------------------------
'COST
'------------------------------------------------------------------------------------------
                    On Error GoTo 0
                    
                    Debug.Print swModel.Extension.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
                    swModel.Extension.SetUserPreferenceInteger swUserPreferenceIntegerValue_e.swUnitSystem, swUserPreferenceOption_e.swDetailingNoOptionSpecified, swUnitSystem_IPS
                    Debug.Print swModel.Extension.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
                    
                    Call MaterialUpdate
                    Call N325
                    swModel.Save3 1, 0, 0
'------------------------------------------------------------------------------------------
'CUSTOM PROPERTIES
'------------------------------------------------------------------------------------------
                    If swModel.GetCustomInfoValue("", "Description") = "" Then
                        WriteDescription = True
                        If swModel.GetCustomInfoValue(gstrConfigName, "F325") = 1 Then
                            swModel.CustomInfo2("", "Description") = modSetMaterial.strMaterial & " ROLL"
                        ElseIf swModel.GetCustomInfoValue(gstrConfigName, "F325") <> 1 And swModel.GetCustomInfoValue(gstrConfigName, "PressBrake") = "Checked" Then
                            swModel.CustomInfo2("", "Description") = modSetMaterial.strMaterial & " BENT"
                        Else
                            swModel.CustomInfo2("", "Description") = modSetMaterial.strMaterial & " PLATE"
                        End If
                    Else
                        WriteDescription = False
                    End If
                    
                    Call CustomProperties
'------------------------------------------------------------------------------------------
'CREATE DRAWING FROM MODEL
'------------------------------------------------------------------------------------------
                    If SemiAutoPilot.cbCreateDrawing = True Then Call CreateDrawing
                    
                    Call SaveCurrentModel

                Else  'bool Find Sheet Metal = False
                    If swModel.GetCustomInfoValue("", "Description") = "" Then
                        WriteDescription = True
                    Else
                        WriteDescription = False
                    End If
                    
                    ''swApp.DocumentVisible True, swDocPART
                    'swApp.UserControl = True
                    'swApp.Visible = True
                    swApp.ActivateDoc ModelName & ".sldprt"
                    'swApp.FrameState = swWindowMinimized
                    Call ExtractTubeData
                    swApp.CloseDoc ModelName & ".sldprt"
                    
                    ''swApp.DocumentVisible False, swDocPART
                    'swApp.UserControl = False
                    'swApp.Visible = False
                    
                    Call TubeCustomProperties
                    Call CustomProperties
                    If SemiAutoPilot.cbCreateDrawing.value = True Then Call CreateDrawing
                    Call SaveCurrentModel
                    
                End If  'find Sheet Metal
                
                Else
                MsgBox "Body count " & BodyCount & " for " & ModelName
                
            End If  'BodyCount
            
        ElseIf CheckOP20 = "F110 - TUBE LASER" Then
            ''swApp.DocumentVisible True, swDocPART
            'swApp.UserControl = True
            'swApp.Visible = True

            swApp.ActivateDoc ModelName & ".sldprt"
            'swApp.FrameState = swWindowMinimized
            Call ExtractTubeData
            swApp.CloseDoc ModelName & ".sldprt"
            
            ''swApp.DocumentVisible False, swDocPART
            'swApp.UserControl = False
            'swApp.Visible = False
            
            Call TubeCustomProperties
            Call CustomProperties
            If SemiAutoPilot.cbCreateDrawing.value = True Then Call CreateDrawing
            Call SaveCurrentModel

        End If  'Check OP20
      End If ' rbPartType <> 0
    Else
        Debug.Print "::Could Not Activate:: " & ModelName & ".sldprt"
    End If  'Not swModel is Nothing
End Sub
Sub ShowProgress()
   ' Progress.tbBar1.Width = 100
    Progress.tbBar2.Width = Round((iPartsComplete / iPartsTotal), 2) * 300
    Progress.lblProgress = Round((iPartsComplete / iPartsTotal), 4) * 100 & "% Complete"
    If averageCount > 0 Then
        Progress.lblFiles = "Processing file " & iPartsComplete + 1 & " of " & iPartsTotal & " :: Current avg. time is " & Round((averageTime / averageCount), 2) & " sec."
    Else
        Progress.lblFiles = "Processing file " & iPartsComplete + 1 & " of " & iPartsTotal & " :: Current avg. time is " & averageLast & " sec."
    End If
    'With Progress
    '    .StartUpPosition = 0
    '    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    '    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    '    .Show (False)
    'End With
    Progress.Show (False)
End Sub
Sub CheckLA()
    
End Sub
Sub DeleteFiles()
    Dim fs As New Scripting.FileSystemObject
    With fs
        If .FileExists(cstrRedoFile) Then
            .DeleteFile cstrRedoFile
        End If
        If .FileExists(cstrOutputFile) Then
            .DeleteFile cstrOutputFile
        End If
        If .FileExists(cstrArrayFile) Then
            .DeleteFile cstrArrayFile
        End If
        If .FileExists(cstrGUIFile) Then
            .DeleteFile cstrGUIFile
        End If
        If .FileExists(cstrDocTypeFile) Then
            .DeleteFile cstrDocTypeFile
        End If
        If .FileExists(cstrQuoteASMFile) Then
            .DeleteFile cstrQuoteASMFile
        End If
    End With
End Sub
Sub SingleMain()
    
    Set GlobalExcel = CreateObject("Excel.Application")
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    If swModel.GetType = swDocASSEMBLY Then
        Set swAssembly = swModel
        If swModel.SelectionManager.GetSelectedObjectType3(1, -1) = swSelCOMPONENTS Then
            Set swComponent = swModel.SelectionManager.GetSelectedObject6(1, 0)
        Else
            Set swEnt = swModel.SelectionManager.GetSelectedObject6(1, 0)
            If Not swEnt Is Nothing Then
                On Error Resume Next
                Set swComponent = swEnt.GetComponent
            Else
                Exit Sub
            End If
            On Error GoTo 0
        End If
        Set swPartTemp = swComponent.GetModelDoc2
        Set swModel = swPartTemp
    End If
    
    swModel.ShowConfiguration2 "Default"
    swModel.ShowConfiguration2 "Cut"
    
    If swModel.GetCustomInfoValue(gstrConfigName, "OptiMaterial") = "" Then
        SingleCost.Show
        If strBendTable <> "-1" Then
            Set objExcel = GlobalExcel
            objExcel.Visible = False
            Set objBendTable = objExcel.Workbooks.Open(strBendTable)
        End If
    Else
        MatType = Right(Left(swModel.GetCustomInfoValue(gstrConfigName, "OptiMaterial"), 5), 3)
        Select Case MatType
        Case "304"
            SingleCost.rb304 = True
        Case "ALN"
            SingleCost.rbALNZD = True
        Case "HR*"
            SingleCost.rbA36 = True
        Case "CR*"
            SingleCost.rbA36 = True
        Case "P&O"
            SingleCost.rbA36 = True
        Case "606"
            SingleCost.rb6061 = True
        Case "316"
            SingleCost.rb316 = True
        Case "220"
            SingleCost.rb2205 = True
        Case "C22"
            SingleCost.rbC22 = True
        Case "AL6"
            SingleCost.rbAL6XN = True
        Case "409"
            SingleCost.rb409 = True
        Case "A36"
            SingleCost.rbA36 = True
        Case "505"
            SingleCost.rb5052 = True
        End Select
        SingleCost.btnOK_Click
    End If
    
    If Not swModel Is Nothing Then
      If swModel.GetCustomInfoValue(gstrConfigName, "rbPartType") = "0" Then
        Set swPart = Nothing
        Set swPart = swModel
        
        swApp.CopyAppearance swPart
        swPart.SetMaterialPropertyName2 "Default", "C:/Program Files/SolidWorks Corp/SolidWorks/lang/english/sldmaterials/SolidWorks Materials.sldmat", MaterialType
        swApp.PasteAppearance swPart, 3
        
        CheckOP20 = ""
        gstrConfigName = ""
        CheckOP20 = swModel.GetCustomInfoValue(gstrConfigName, "OP20")
        If CheckOP20 = "" Or CheckOP20 = "N115 - ANY FLAT LASER" Or CheckOP20 = "N120 - 5040" Or CheckOP20 = "N125 - 3060" Then
            BodyCount = NumberOfBodies()
            If BodyCount = 1 Then
                boolFindSheetMetal = SMInsertBends()
                If boolFindSheetMetal = True Then
                    If CheckOP20 = "" Then
                        OP20 = "N120 - 5040"
                        swModel.CustomInfo2(gstrConfigName, "OP20") = OP20
                    End If
                    OptiMaterial = swModel.GetCustomInfoValue(gstrConfigName, "OptiMaterial")
                    If OptiMaterial = "" Then
                        Call SetMaterial
                        swModel.CustomInfo2(gstrConfigName, "OptiMaterial") = OptiMaterial
                    End If
                    Call MaterialUpdate
                    Call N325
                    swModel.Save3 1, 0, 0
                    If swModel.GetCustomInfoValue("", "Description") = "" Then
                        WriteDescription = True
                        If swModel.GetCustomInfoValue(gstrConfigName, "F325") = 1 Then
                            swModel.CustomInfo2("", "Description") = modSetMaterial.strMaterial & " ROLL"
                        ElseIf swModel.GetCustomInfoValue(gstrConfigName, "F325") <> 1 And swModel.GetCustomInfoValue(gstrConfigName, "PressBrake") = "Checked" Then
                            swModel.CustomInfo2("", "Description") = modSetMaterial.strMaterial & " BENT"
                        Else
                            swModel.CustomInfo2("", "Description") = modSetMaterial.strMaterial & " PLATE"
                        End If
                    Else
                        WriteDescription = False
                    End If
                    Call CustomProperties
                    If SingleCost.cbCreateDrawing = True Then Call CreateDrawing
                    Call SaveCurrentModel
                Else
                    If swModel.GetCustomInfoValue("", "Description") = "" Then
                        WriteDescription = True
                    Else
                        WriteDescription = False
                    End If
                    
                    Call ExtractTubeData
                    Call TubeCustomProperties
                    Call CustomProperties
                    If SingleCost.cbCreateDrawing = True Then Call CreateDrawing
                    Call SaveCurrentModel
                End If  'find Sheet Metal
            End If  'BodyCount
            
        ElseIf CheckOP20 = "F110 - TUBE LASER" Then
            
            Call ExtractTubeData
            Call TubeCustomProperties
            Call CustomProperties
            If SingleCost.cbCreateDrawing = True Then Call CreateDrawing
            Call SaveCurrentModel

        End If  'Check OP20
      End If ' rbPartType <> 0
    End If  'Not swModel is Nothing
    swModel.ShowConfiguration2 "Default"
    If Not swAssembly Is Nothing Then
        Set swModel = swAssembly
        Set swAssembly = Nothing
        swApp.ActivateDoc swModel.GetPathName
    End If
    swModel.ClearSelection2 True
    swModel.EditRebuild3
    swModel.ForceRebuild3 False

    swApp.DocumentVisible True, swDocPART
    
    GlobalExcel.Quit
    Set GlobalExcel = Nothing
    
End Sub
Sub main()

    Dim boolRestart As Boolean
    Dim NumOfPartsCheck, SecondsCheck As Integer
    Set swApp = Application.SldWorks
    Set GlobalExcel = CreateObject("Excel.Application")

    Set swFSO = CreateObject("Scripting.FileSystemObject")
    averageTime = 0: averageCount = 0
    NumOfPartsCheck = 5  'how many parts to check at a time
    If bQuoteASM = False Then
        SecondsCheck = 9999   'average time acceptable before restarting
    Else
        SecondsCheck = 9999
    End If
    If Not bolTBC = True Then
        Call DeleteFiles
    End If
    'Call CheckLA   ' this would be designed to remove Large Assembly mode automatically in the future
    Call Initialize
    
    For intPart = LBound(ModelNames) To UBound(ModelNames)
        If Not ModelNames(intPart) = "" Then
            boolFindSheetMetal = True
            If Int(intPart / NumOfPartsCheck) = Round(intPart / NumOfPartsCheck, 1) Then
                'If RestartCheck(intPart, SecondsCheck) = True Then  'Runs RestartCheck to kill SW if prog is running too slow
                '    Exit Sub
                'End If
            End If
            Call ShowProgress
            totalTime = 0
            loopSTime = Now()   'start time for loop of one part
            ModelNameRaw = ModelNames(intPart)
    '        If bQuote = False And ModelNames(intPart) <> MainAssembly Then
    '            intPosition = InStrRev(ModelNameRaw, "-")
    '            ModelNameRaw = Left(ModelNameRaw, intPosition - 1)
    '            intPosition = InStrRev(ModelNameRaw, "/")
    '            IntLength = Len(ModelNameRaw)
    '        Else
    '            intPosition = InStrRev(ModelNameRaw, ".")
    '            ModelNameRaw = Left(ModelNameRaw, intPosition - 1)
    '            intPosition = InStrRev(ModelNameRaw, "/")
    '            IntLength = Len(ModelNameRaw)
    '        End If
            
            intPosition = InStrRev(ModelNameRaw, ".")
            ModelNameRaw = Left(ModelNameRaw, intPosition - 1)
            intPosition = InStrRev(ModelNameRaw, "\")
            IntLength = Len(ModelNameRaw)
            
            ModelNameRaw = Right$(ModelNameRaw, IntLength - intPosition)
            Debug.Print "Model Name Raw :: " & ModelNameRaw
            If IsInArray(ModelNameRaw, ModelNamesRaw) = False Then 'if not in array, this part needs to run through macro
                'Do Stuff
                Call ProcessModel
                'Debug.Print boolFindSheetMetal
                If Not boolFindSheetMetal = True Then
                    'Save info for later to edit parts
                    
                    
                    'ModelNamesRedo(FindLastValue(ModelNamesRedo())) = ModelNameRaw
                    
                End If
            End If  'is inArray
        End If
        loopETime = Now()   'stop time for loop
        If DateDiff("s", loopSTime, loopETime) > 1 Then
            averageCount = averageCount + 1
            averageTime = averageTime + DateDiff("s", loopSTime, loopETime)
        End If
        Debug.Print "Loop Time :: " & DateDiff("s", loopSTime, loopETime) & " second(s)"
        'iPartsComplete = iPartsComplete + 1
        iPartsComplete = intPart + 1
    Next intPart
    
    Progress.Hide
    
    If bQuoteASM = False Then   'Temporary loop to skip manual review for Quote Assemblies
    
        If FindLastValue(ModelNamesRedo()) > 0 Then
            For N = 0 To UBound(ModelNamesRedo())
                If Trim(ModelNamesRedo(N)) <> "" Then
                    If bQuote = False Then
                        Set swModel = swApp.ActivateDoc(ModelNamesRedo(N) & ".sldprt")
                        If Not swModel Is Nothing Then
                            swModel.ShowConfiguration2 ("Default")
                        End If
                    Else
                        Set swModel = swApp.OpenDoc6(filePath & ModelNamesRedo(N) & ".sldprt", 1, 1, "Default", longerrors, longwarnings)
                    End If
                    Debug.Print ModelNamesRedo(N)
                    frmPause.Show (False)  'show as non modal to allow user interaction
                    While frmPause.Visible = True   'Loop until user clicks continue
                        DoEvents
                    Wend
                End If
                
                If Not swModel Is Nothing Then
                    swModel.Save3 1, 0, 0
                    swModel.ClearSelection2 True
                    If swModel.GetTitle <> MainAssembly Then
                        swApp.CloseDoc (swModel.GetTitle)
                        Set swModel = Nothing
                    End If
                End If
            Next N
        End If
    End If
    
    Set swModel = swApp.ActiveDoc
    If Not swModel Is Nothing Then
        swModel.ClearSelection2 True
        swApp.CloseDoc (swModel.GetTitle)
        Set swModel = Nothing
    End If
    
    On Error Resume Next
    objBendTable.Close False
    On Error GoTo 0
    
    If SemiAutoPilot.cbReport = True Then
        If bQuote = True Then
            frmLoading.Show
            Call ReportPart
        Else
            frmLoading.Show
            Call Report
        End If
    End If
    
    GlobalExcel.Visible = True
    GlobalExcel.Quit
    Set GlobalExcel = Nothing
    
    swApp.UserControl = True
    swApp.Visible = True
    swApp.DocumentVisible True, swDocPART
    swApp.DocumentVisible True, swDocASSEMBLY
    swApp.DocumentVisible True, swDocDRAWING
    
    Progress.Hide
    frmLoading.Hide
    If Not bQuoteASM = True Then
        swApp.FrameState = swWindowMaximized
        frmComplete.Show
    End If
    
    If boolENG = True Then
        End
    End If
    
End Sub

Function IsInArray(ModelNameRaw As String, ModelNamesRaw As Variant) As Boolean
  IsInArray = UBound(Filter(ModelNamesRaw, "[" & ModelNameRaw & "]")) > -1
  Debug.Print "IsInArray :: " & IsInArray
End Function

Sub CustomProperties()
    If CustomProps.tbCustomer.value <> "" Then
        swModel.CustomInfo2(gstrConfigName, "Customer") = CustomProps.tbCustomer.value
    End If
    
    If CustomProps.cbPrintFromPart.value = True Then
        swModel.CustomInfo2(gstrConfigName, "Print") = ModelNameRaw
        Debug.Print "Setting Print to ModelNameRaw :: " & ModelNameRaw
    ElseIf CustomProps.tbPrint.value <> "" Then
        swModel.CustomInfo2(gstrConfigName, "Print") = CustomProps.tbPrint.value
    End If
    
    If CustomProps.tbRevision.value <> "" Then
        swModel.CustomInfo2(gstrConfigName, "Revision") = CustomProps.tbRevision.value
    End If
    
    If CustomProps.tbDescription.value <> "" And WriteDescription = True Then
        swModel.CustomInfo2(gstrConfigName, "Description") = CustomProps.tbDescription.value & " " & swModel.GetCustomInfoValue(gstrConfigName, "Description")
    End If
    
    If CustomProps.cbGrain.value = True Then
        swModel.CustomInfo2(gstrConfigName, "Grain") = "Y"
    End If
    
    If CustomProps.cbCommon.value = True Then
        swModel.CustomInfo2(gstrConfigName, "ComCut") = ""
        swModel.CustomInfo2(gstrConfigName, "External") = "E"
    End If
End Sub
Sub SingleDrawing()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    Set swFSO = CreateObject("Scripting.FileSystemObject")
    ModelNameDXF = swModel.GetTitle
    FilePathDXF = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\"))
    FilePath2DXF = FilePathDXF & ModelNameDXF
    DocName = swModel.GetTitle
    intPosition = InStrRev(DocName, ".")
    DocName = Left(DocName, intPosition - 1)
    FilePath3 = FilePathDXF & DocName & ".slddrw"
    
    Set swPart = swModel
    NumberOfBodies
    Features = swBody.GetFeatures
    
    If swFSO.FileExists(FilePath3) = True Then
        swApp.OpenDoc6 FilePath3, swDocDRAWING, 0, "", longerrors, longwarnings
    Else
        CreateDrawing
        swApp.DocumentVisible True, swDocPART
        swApp.DocumentVisible True, swDocDRAWING
        swApp.OpenDoc6 FilePath2DXF, swDocPART, 0, "", longerrors, longwarnings
        swApp.OpenDoc6 FilePath3, swDocDRAWING, 0, "", longerrors, longwarnings
    End If
End Sub
Sub CreateDrawing()
    
sTime = Now()
    Dim swConfMgr As SldWorks.ConfigurationManager
    Dim swFlatPatFol As FlatPatternFolder
    Dim swView As SldWorks.View
    swModel.Save3 1, 0, 0
    Set part = swModel
    Set swFSO = CreateObject("Scripting.FileSystemObject")

    ModelNameDXF = swModel.GetTitle
    FilePathDXF = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\"))
    FilePath2DXF = FilePathDXF & ModelNameDXF
    DocName = swModel.GetTitle
    intPosition = InStrRev(DocName, ".")
    DocName = Left(DocName, intPosition - 1)
    FilePath3 = FilePathDXF & DocName & ".slddrw"
                                    
    'check for existing drawing
    bolFileExists = swFSO.FileExists(FilePath3)
    
    OP20 = ""
    OP20 = swModel.GetCustomInfoValue(gstrConfigName, "OP20")
    If bolFileExists = True Then
        'swApp.OpenDoc6 FilePath3, swDocDRAWING, 0, "", longerrors, longwarnings
    ElseIf OP20 <> "" Then
    
        PressBrake = swModel.GetCustomInfoValue(gstrConfigName, "PressBrake")
        PinchRoll = swModel.GetCustomInfoValue(gstrConfigName, "F325")
        'create drawing with flatpattern view
        swApp.DocumentVisible True, swDocDRAWING
        Set swFlatPatFol = swModel.FeatureManager.GetFlatPatternFolder
        Set swConfMgr = swModel.ConfigurationManager
        
        If Left(OP20, 4) = "N115" Or Left(OP20, 4) = "N120" Or Left(OP20, 4) = "N125" Then
            If swFlatPatFol Is Nothing Then
                swModel.SetBendState 2
                swModel.EditRebuild3
                UnsuppressFlatten
                swConfMgr.AddConfiguration "DefaultSM-FLAT-PATTERN", "Flattened state of sheet metal part", "", 0, "Default", "DefaultSM-FLAT-PATTERN"
                swModel.ShowConfiguration2 "Default"
                swModel.SetBendState 3
                swModel.EditRebuild3
            Else
                Features = swFlatPatFol.GetFlatPatterns
                For u = LBound(Features) To UBound(Features)
                    Feature = Features(u).GetTypeName
                    Set swFeat = Features(u)
                    If Feature = "FlatPattern" Then
                        swFeat.SetSuppression2 swFeatureSuppressionAction_e.swUnSuppressFeature, swInConfigurationOpts_e.swThisConfiguration, Nothing
                        Set swSubFeat = swFeat.GetFirstSubFeature
                        While Not swSubFeat Is Nothing
                            Debug.Print swSubFeat.GetTypeName
                            swSubFeat.SetSuppression2 swFeatureSuppressionAction_e.swUnSuppressFeature, swInConfigurationOpts_e.swThisConfiguration, Nothing
                            Set swSubFeat = swSubFeat.GetNextSubFeature
                        Wend
                    End If
                Next u
                swConfMgr.AddConfiguration "DefaultSM-FLAT-PATTERN", "Flattened state of sheet metal part", "", 0, "Default", "DefaultSM-FLAT-PATTERN"
                swModel.ShowConfiguration2 "Default"
                For u = LBound(Features) To UBound(Features)
                    Feature = Features(u).GetTypeName
                    Set swFeat = Features(u)
                    If Feature = "FlatPattern" Then
                        swFeat.SetSuppression2 swFeatureSuppressionAction_e.swSuppressFeature, swInConfigurationOpts_e.swThisConfiguration, Nothing
                    End If
                Next u
            End If
        End If
        
        Set swDraw = Nothing
        Set swDraw = swApp.NewDocument("O:\Engineering Department\Solidworks\Document Templates\Northern-Rev4\A-SIZE.drwdot", 12, 0.2159, 0.2794)
        Views = swDraw.GenerateViewPaletteViews(FilePath2DXF)
        Set swDrawMod = swDraw
        'swApp.DocumentVisible False, swDocDRAWING
        Set myView = Nothing
        Set myView = swDraw.DropDrawingViewFromPalette2("Flat Pattern", 0, 0, 0)
        Set swView = myView
        If myView Is Nothing Then
            Set myView = swDraw.DropDrawingViewFromPalette2("*Right", 0, 0, 0)
        End If
        
        boolstatus = swDraw.ForceRebuild3(False)

'rotate view if needed
        
        vOutline = myView.GetOutline
        MaxX = vOutline(2) - vOutline(0)
        MaxY = vOutline(3) - vOutline(1)
        
        Dim swSheet As SldWorks.Sheet
        Dim SheetProperties As Variant
        Dim SheetScale As Double
        
        Set swSheet = swDraw.GetCurrentSheet
        SheetProperties = swSheet.GetProperties2()
        SheetScale = SheetProperties(2) / SheetProperties(3)
        
        MaxXInch = 1000 * (MaxX / 25.4) / SheetScale
        MaxYInch = 1000 * (MaxY / 25.4) / SheetScale
        
        If MaxX <= MaxY And MaxXInch <= 6 Then 'convert meter to inches
            swModel.CustomInfo2(gstrConfigName, "Grain") = "Y"
            Debug.Print "MaxX = " & Round((MaxX * 39.3701), 2)
        ElseIf (MaxYInch) <= 6 Then
            swModel.CustomInfo2(gstrConfigName, "Grain") = "Y"
            Debug.Print "MaxY = " & Round((MaxY * 39.3701), 2)
        End If
        
        If MaxY > MaxX Then
            swDraw.ActivateView "Drawing View1"
            swDraw.Extension.SelectByID2 "Drawing View1", "DRAWINGVIEW", 0.124472575110103, 0.118161225698324, 0, False, 0, Nothing, 0
            swDraw.Extension.SelectByID2 "Drawing View1", "DRAWINGVIEW", 0.124472575110103, 0.118161225698324, 0, False, 0, Nothing, 0
            swDraw.DrawingViewRotate (1.5707963267949)
        End If
        
        Position = myView.Position
        Position(0) = 0.0445 - myView.GetOutline(0)
        Position(1) = 0.0603 - myView.GetOutline(1)
        myView.Position = Position
        swDraw.EditRebuild

        If myView.GetOutline(2) > 0.268 Then
                Position = myView.Position
                Position(0) = Position(0) - (myView.GetOutline(2) - 0.268)
                myView.Position = Position
                'swDraw.EditRebuild
        End If

        If myView.GetOutline(3) > 0.21 Then
            Position = myView.Position
            Position(1) = Position(1) - (myView.GetOutline(3) - 0.21)
            myView.Position = Position
            'swDraw.EditRebuild
        End If

        'save drawing
        If SemiAutoPilot.cbCreateDXF.value = True Then
            bRet = swDraw.SaveAs4(FilePathDXF & "\" & DocName & ".dxf", _
                swSaveAsCurrentVersion, _
                swSaveAsOptions_Silent, _
                longerrors, _
                longwarnings)
        End If
    
        If SemiAutoPilot.cbCreateDrawing.value = True Then
            bRet = swDraw.SaveAs4(FilePath3, _
                swSaveAsCurrentVersion, _
                swSaveAsOptions_Silent, _
                longerrors, _
                longwarnings)
        End If
        

        swModel.Save3 1, 0, 0
        
        
        swApp.DocumentVisible True, swDocPART
        swApp.ActivateDoc swDraw.GetPathName
        
'add formed view

        If OP20 = "F110 - TUBE LASER" Then
            Position = myView.Position
            Position(0) = 0
            Position(1) = 0
            myView.Position = Position
            Position = myView.Position
            Position(0) = 0.0445 - myView.GetOutline(0)
            Position(1) = 0.1215
            myView.Position = Position
            
            boolstatus = swDraw.ActivateView("Drawing View1")
            boolstatus = swDraw.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0.124472575110103, 0.118161225698324, 0, False, 0, Nothing, 0)
            boolstatus = swDraw.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0.124472575110103, 0.118161225698324, 0, False, 0, Nothing, 0)
            DimensionTube swApp, swDraw, swDraw.ActiveDrawingView
            boolstatus = swDraw.ActivateView("Drawing View1")
            boolstatus = swDraw.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0.124472575110103, 0.118161225698324, 0, False, 0, Nothing, 0)
            boolstatus = swDraw.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0.124472575110103, 0.118161225698324, 0, False, 0, Nothing, 0)
            Set myView = Nothing
            Set myView = swDraw.CreateUnfoldedViewAt3(0.3, swView.Position(1), 0, False)
            swDraw.ClearSelection2 True
            boolstatus = swDraw.ActivateSheet("Sheet1")
            boolstatus = swDraw.ActivateView("Drawing View2")
            boolstatus = swDraw.Extension.SelectByID2("Drawing View2", "DRAWINGVIEW", 0.24995992706541, 0.113486206703911, 0, False, 0, Nothing, 0)
            Set myView = Nothing
            Set myView = swDraw.ActiveDrawingView
            myView.ReferencedConfiguration = "Default"
            swDraw.ForceRebuild
            DimensionTube swApp, swDraw, swDraw.ActiveDrawingView
            swDraw.ClearSelection2 True
            'swDraw.ViewZoomtofit2
            AlignDims.Align swDraw
            If myView.GetOutline(2) > 0.268 Then
                Position = myView.Position
                Position(0) = Position(0) - (myView.GetOutline(2) - 0.268)
                myView.Position = Position
            '    swDraw.EditRebuild
            End If

            If swView.GetOutline(2) > myView.GetOutline(0) - 0.00635 Then
                Position = swView.Position
                ProjPosition = myView.Position
                Position(0) = Position(0) - (swView.GetOutline(2) - (myView.GetOutline(0) - 0.00635))
                swView.Position = Position
                myView.Position = ProjPosition
            '    swDraw.EditRebuild
            End If
            
            Position = swView.Position
            Position(1) = 0.1215
            swView.Position = Position
            TimeOut = 0
            While myView.GetOutline(0) - swView.GetOutline(2) > 0.0254 And TimeOut <> 25
                SheetProperties = swView.Sheet.GetProperties2()
                SheetProperties(2) = SheetProperties(2) * 1.05
                swView.Sheet.SetProperties2 SheetProperties(0), SheetProperties(1), SheetProperties(2), SheetProperties(3), SheetProperties(4), SheetProperties(5), SheetProperties(6), SheetProperties(7)
                
                Position = swView.Position
                Position(0) = 0
                swView.Position = Position
                
                Position = swView.Position
                Position(0) = 0.0445 - swView.GetOutline(0)
                swView.Position = Position
                
                Position = myView.Position
                Position(0) = Position(0) - (myView.GetOutline(2) - 0.268)
                myView.Position = Position
                
                TimeOut = TimeOut + 1
            Wend
            
        Else
            boolstatus = swDraw.ActivateView("Drawing View1")
            boolstatus = swDraw.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0.124472575110103, 0.118161225698324, 0, False, 0, Nothing, 0)
            boolstatus = swDraw.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0.124472575110103, 0.118161225698324, 0, False, 0, Nothing, 0)
            DimensionFlat swApp, swDraw, swView
            AlignDims.Align swDraw
            
        End If
        
        swDraw.EditRebuild
    
'hide planes
        If cbHidePlanes = True Then
    
            Set swFeat = Nothing
            Set swFeat = swModel.FirstFeature
            vFeats = swFeatMgr.GetFeatures(False)
            ReDim vFeatureName(0 To UBound(vFeats)) As String

            For p = LBound(vFeats) To UBound(vFeats)
                Set swFeat = Nothing
                Set swFeat = vFeats(p)
                FeatureName = swFeat.GetTypeName

                If FeatureName = "RefPlane" Then
                    bRet = swFeat.Select(True)
                    swModel.BlankRefGeom
                End If
'                                        Set swFeat = Nothing
                Set swFeat = swFeat.GetNextFeature

                If Not swFeat Is Nothing Then
                    FeatureName = swFeat.GetTypeName
                End If
            Next p
    
            swModel.Save3 1, 0, 0
        End If
        
'make etch marks visible
    Dim EtchArray() As String
    Dim EtchName As String
    
    swModel.ShowConfiguration2 swView.ReferencedConfiguration
    
    Set swFeat = Nothing
    Set swFeat = swModel.FirstFeature
    vFeats = swModel.FeatureManager.GetFeatures(False)
    ReDim EtchArray(0 To UBound(vFeats))

    For p = LBound(vFeats) To UBound(vFeats)
        Set swFeat = Nothing
        Set swFeat = vFeats(p)
        FeatureType = swFeat.GetTypeName
        Debug.Print FeatureType
        If FeatureType = "ProfileFeature" Then
            If swFeat.Visible = 2 Then
                EtchName = swFeat.GetNameForSelection("ProfileFeature")
                If Not IsInArray(EtchName, EtchArray) = True And Not Left(EtchName, 5) = "Bound" Then
                    EtchArray(FindLastValue(EtchArray())) = "[" & EtchName & "]"
                End If
            End If
            
        End If
        Set swFeat = swFeat.GetNextFeature

        If Not swFeat Is Nothing Then
            FeatureName = swFeat.GetTypeName
        End If
    Next p

    For p = LBound(EtchArray) To UBound(EtchArray)
        If Not EtchArray(p) = "" Then
            EtchName = Right(Left(EtchArray(p), Len(EtchArray(p)) - 1), Len(EtchArray(p)) - 2)
            EtchName = EtchName & "@" & swModel.GetTitle & "@" & swView.GetUniqueName
            swDraw.Extension.SelectByID2 EtchName, "SKETCH", 0, 0, 0, False, 0, Nothing, 0
            swDraw.UnblankSketch
            swDraw.ClearSelection2 True
        End If
    Next p
    swModel.ShowConfiguration2 "Default"
    
'final save
    
        If SemiAutoPilot.cbCreateDrawing.value = True Then
            bRet = swDraw.SaveAs4(FilePath3, _
                swSaveAsCurrentVersion, _
                swSaveAsOptions_Silent, _
                longerrors, _
                longwarnings)
        End If
    End If
    
    swModel.Save3 1, 0, 0
    swModel.ClearSelection2 True
    If Not swDrawMod Is Nothing Then
        swApp.CloseDoc (swDrawMod.GetTitle)
        Set swDrawMod = Nothing
    End If
    On Error Resume Next
    Set swDraw = Nothing
    If swModel.GetTitle <> MainAssembly Then
        swApp.CloseDoc (swModel.GetTitle)
        Set swModel = Nothing
    End If
    On Error GoTo 0
    
End Sub

Sub Report()
    
    Dim swFSO As New Scripting.FileSystemObject
    Dim bolFileExists As Boolean
    Dim MyFiles As Scripting.Files
    Dim MyFile As Scripting.File
    Dim swModelDocExt As SldWorks.ModelDocExtension
    Dim swBOMAnnotation As SldWorks.BomTableAnnotation
    Dim swBomFeature As SldWorks.BomFeature
    Dim swTable As SldWorks.TableAnnotation
    Dim TableRows As Integer
    Dim ComponentCount As Variant
    Dim ItemNumber As String
    Dim ModelWeight As Double
    
    On Error GoTo 0
    'If UCase(Right(MainAssembly, 3)) = "ASM" Then
    
    Set swFSO = CreateObject("Scripting.FileSystemObject")
    
    ExcelFile = filePath + Left(MainAssembly, InStrRev(MainAssembly, ".") - 1) + ".xlsm"
    
    Dim longerrors As Long
    Dim longwarnings As Long
    
    bolFileExists = swFSO.FileExists(ExcelFile)
    
    strTemplate = "O:\Engineering Department\Solidworks\Macros\(Semi)Autopilot\SLDPRT-TYSON.xlsm"
    
    Set objExcel = GlobalExcel
    objExcel.Visible = False
    
   
    
    If bolFileExists = True Then
        frmExisting.Label2.Caption = ExcelFile
        frmExisting.Show
        swFSO.DeleteFile (ExcelFile)
    End If
    
    Set objWorkbook = objExcel.Workbooks.Open(strTemplate)
    objWorkbook.SaveAs (ExcelFile)
    objWorkbook.Close True, ExcelFile
    
    Set objWorkbook = objExcel.Workbooks.Open(ExcelFile)
    Set objSheet = objWorkbook.ActiveSheet
    
    If UCase(Right(MainAssembly, 3)) = "ASM" Then
        Set swModel = swApp.OpenDoc6(filePath + MainAssembly, swDocASSEMBLY, swOpenDocOptions_Silent, "Default", longerrors, longwarnings)
        If longerrors > 0 Then
            MsgBox MainAssembly + " has loading errors", vbOKOnly
            
        End If
    Else
        Set swModel = swApp.OpenDoc6(filePath + MainAssembly, swDocPART, swOpenDocOptions_Silent, "Default", longerrors, longwarnings)
    End If
    
    swApp.ActivateDoc filePath + MainAssembly
    
    Set swBOMAnnotation = swModel.Extension.InsertBomTable3("O:\Engineering Department\Solidworks\BOM Templates\bom-standard.sldbomtbt", 0, 1, swBomType_PartsOnly, "Default", False, swIndentedBOMNotSet, False)
    Set swBomFeature = swBOMAnnotation.BomFeature
    swModel.ForceRebuild3 (False)
    Set swTable = swBOMAnnotation
    TableRows = swTable.RowCount
    
    ReDim ComponentCount(0 To TableRows - 2) As String
    ReDim ModelNames(0 To TableRows - 2) As String
    
    For i = 0 To TableRows - 2
        ComponentCount(i) = swBOMAnnotation.GetComponentsCount2(i + 1, "Default", ItemNumber, ModelNames(i))
    Next i
        
    For i = LBound(ModelNames) To UBound(ModelNames)
    
        ModelName = ModelNames(i)
        FileName = filePath & ModelName & ".sldprt"
        'swApp.CloseAllDocuments True
        Set swModel = swApp.OpenDoc6(FileName, swDocPART, swOpenDocOptions_Silent, "Default", longerrors, longwarnings)
        
        objSheet.Hyperlinks.Add objSheet.Cells(i + 3, 1), FileName, , , ModelName
        
        Debug.Print ModelName
        Debug.Print "rbPartType - " & swModel.GetCustomInfoValue("", "rbPartType")
        Debug.Print "rbPartTypeSub - " & swModel.GetCustomInfoValue("", "rbPartTypeSub")
    
        
        If swModel.GetCustomInfoValue("", "rbPartType") = 1 Then
            If swModel.GetCustomInfoValue("", "rbPartTypeSub") = 0 Then
                OP20 = "MP - MACHINED"
                OP20_S = "0"
                OP20_R = "0"
                OptiMat = "MP-" & ModelName
                WeightLength = "1"
            ElseIf swModel.GetCustomInfoValue("", "rbPartTypeSub") = 1 Then
                OP20 = "NPUR - PURCHASED"
                OP20_S = "0"
                OP20_R = "0"
                OptiMat = swModel.GetCustomInfoValue("", "PurchasedPartNumber")
                WeightLength = "1"
            ElseIf swModel.GetCustomInfoValue("", "rbPartTypeSub") = 2 Then
                OP20 = "CUST - SUPPLIED"
                OP20_S = "0"
                OP20_R = "0"
                OptiMat = swModel.GetCustomInfoValue("", "CustPartNumber")
                WeightLength = "1"
            End If
            
            objSheet.Range("C" & i + 3).NumberFormat = "0.00" & """" & "ea" & """"
            objSheet.Range("P" & i + 3).NumberFormat = "0.00" & """" & "ea" & """"
            objSheet.Range("Q" & i + 3).NumberFormat = "$* #,##0.00" & """" & "/ea" & """"
            
            objSheet.Cells(i + 3, 8) = "N"
            objSheet.Cells(i + 3, 9) = "0"
            objSheet.Cells(i + 3, 10) = "0"
            objSheet.Cells(i + 3, 11) = "0"
            objSheet.Cells(i + 3, 12) = "0"
            
        Else
            OP20 = swModel.GetCustomInfoValue("", "OP20")
            OP20_S = swModel.GetCustomInfoValue("", "OP20_S")
            OP20_R = swModel.GetCustomInfoValue("", "OP20_R")
            OptiMat = swModel.GetCustomInfoValue("", "OptiMaterial")
            ModelWeight = swModel.GetCustomInfoValue("", "Weight")
            
            If Not swModel.GetCustomInfoValue("", "RawWeight") = "" Then
                WeightLength = swModel.GetCustomInfoValue("", "RawWeight")
                objSheet.Range("C" & i + 3).NumberFormat = "0.00" & """" & "lb" & """"
                objSheet.Range("P" & i + 3).NumberFormat = "0.00" & """" & "lb" & """"
                objSheet.Range("Q" & i + 3).NumberFormat = "$* #,##0.00" & """" & "/lb" & """"
            Else
                WeightLength = swModel.GetCustomInfoValue("", "F300_Length")
                objSheet.Range("C" & i + 3).NumberFormat = "0.00" & """" & "in" & """"
                objSheet.Range("P" & i + 3).NumberFormat = "0.00" & """" & "in" & """"
                objSheet.Range("Q" & i + 3).NumberFormat = "$* #,##0.00" & """" & "/in" & """"
            End If
            
            If swModel.GetCustomInfoValue("", "Pressbrake") = "Checked" Then
                objSheet.Cells(i + 3, 8) = "Y"
                objSheet.Cells(i + 3, 9) = swModel.GetCustomInfoValue("", "F140_S")
                objSheet.Cells(i + 3, 10) = swModel.GetCustomInfoValue("", "F140_R")
            Else
                objSheet.Cells(i + 3, 8) = "N"
                objSheet.Cells(i + 3, 9) = "0"
                objSheet.Cells(i + 3, 10) = "0"
            End If
            
            objSheet.Cells(i + 3, 11) = swModel.GetCustomInfoValue("", "F220_S") + swModel.GetCustomInfoValue("", "F210_S") + swModel.GetCustomInfoValue("", "F325_S") + _
                    swModel.GetCustomInfoValue("", "Other_S") + swModel.GetCustomInfoValue("", "Other_S2")
            objSheet.Cells(i + 3, 12) = swModel.GetCustomInfoValue("", "F220_R") + swModel.GetCustomInfoValue("", "F210_R") + swModel.GetCustomInfoValue("", "F325_R") + _
                    swModel.GetCustomInfoValue("", "Other_R") + swModel.GetCustomInfoValue("", "Other_R2")
            
        End If
        objSheet.Cells(i + 3, 2) = ModelWeight
        objSheet.Cells(i + 3, 3) = OptiMat
        objSheet.Cells(i + 3, 4) = WeightLength
        objSheet.Cells(i + 3, 5) = OP20
        objSheet.Cells(i + 3, 6) = OP20_S
        objSheet.Cells(i + 3, 7) = OP20_R
        objSheet.Cells(i + 3, 13) = ComponentCount(i)
        'swApp.CloseAllDocuments True
        
    Next i
    

    
 
    objWorkbook.Save
    objWorkbook.Close True, ExcelFile
    Set objSheet = Nothing
    Set objWorkbook = Nothing
    Set objBendTable = Nothing
    Set objExcel = Nothing
    'end if
    swApp.UserControl = True
    swApp.Visible = True
    swApp.CloseAllDocuments True
    Set swModel = swApp.OpenDoc6(filePath + MainAssembly, swDocASSEMBLY, swOpenDocOptions_Silent, "Default", longerrors, longwarnings)
    swApp.ActivateDoc filePath + MainAssembly
    If swModel Is Nothing Then
        swApp.DocumentVisible True, swDocPART
        Set swModel = swApp.OpenDoc6(filePath + MainAssembly, swDocPART, swOpenDocOptions_Silent, "Default", longerrors, longwarnings)
    End If
End Sub

Sub ReportPart()
    Dim swFSO As New Scripting.FileSystemObject
    Dim bolFileExists As Boolean
    Dim MyFiles As Scripting.Files
    Dim MyFile As Scripting.File
    
    Set swFSO = CreateObject("Scripting.FileSystemObject")
    
    ExcelFile = filePath + "Report.xlsm"
    
    Dim longerrors As Long
    Dim longwarnings As Long
    
    bolFileExists = swFSO.FileExists(ExcelFile)
    
    If bolFileExists = True Then
        frmExisting.Label2.Caption = ExcelFile
        frmExisting.Show
        swFSO.DeleteFile (ExcelFile)
    End If
    
    strTemplate = "O:\Engineering Department\Solidworks\Macros\(Semi)Autopilot\SLDPRT-TYSON.xlsm"
    
    Set objExcel = GlobalExcel
    objExcel.Visible = False
    
    Set objWorkbook = objExcel.Workbooks.Open(strTemplate)
    objWorkbook.SaveAs (ExcelFile)
    objWorkbook.Close True, ExcelFile
    
    Set objWorkbook = objExcel.Workbooks.Open(ExcelFile)
    Set objSheet = objWorkbook.ActiveSheet
    
    Set MyFiles = swFSO.GetFolder(filePath).Files
    
    For Each MyFile In MyFiles
        If UCase(Right(MyFile.Name, 6)) = "SLDPRT" Then
            ReDim ModelNames(0 To N) As String
            ReDim ModelNamesRaw(0 To N) As String
            N = N + 1
        End If
    Next MyFile
    
    N = 0
    
    For Each MyFile In MyFiles
        If UCase(Right(MyFile.Name, 6)) = "SLDPRT" Then
            ModelNames(N) = MyFile.Name
            N = N + 1
        End If
    Next MyFile
        
    For i = LBound(ModelNames) To UBound(ModelNames)
    
        ModelName = ModelNames(i)
        intPosition = InStrRev(ModelName, ".")
        ModelName = Left(ModelName, intPosition - 1)
        intPosition = InStrRev(ModelName, "/")
        IntLength = Len(ModelName)
        ModelName = Right$(ModelName, IntLength - intPosition)
        FileName = filePath & ModelName & ".sldprt"
            
        'Set swApp = Application.SldWorks
        Set swModel = swApp.OpenDoc6(FileName, swDocPART, swOpenDocOptions_Silent, "Default", longerrors, longwarnings)
        
        objSheet.Hyperlinks.Add objSheet.Cells(i + 3, 1), FileName, , , ModelName
        
         If swModel.GetCustomInfoValue("", "rbPartType") = 1 Then
            If swModel.GetCustomInfoValue("", "rbPartTypeSub") = 0 Then
                OP20 = "MP - MACHINED"
                OP20_S = "0"
                OP20_R = "0"
                OptiMat = "MP-" & ModelName
                WeightLength = "1"
            ElseIf swModel.GetCustomInfoValue("", "rbPartTypeSub") = 1 Then
                OP20 = "NPUR - PURCHASED"
                OP20_S = "0"
                OP20_R = "0"
                OptiMat = swModel.GetCustomInfoValue("", "PurchasedPartNumber")
                WeightLength = "1"
            ElseIf swModel.GetCustomInfoValue("", "rbPartTypeSub") = 2 Then
                OP20 = "CUST - SUPPLIED"
                OP20_S = "0"
                OP20_R = "0"
                OptiMat = swModel.GetCustomInfoValue("", "CustPartNumber")
                WeightLength = "1"
            End If
            
            objSheet.Range("C" & i + 3).NumberFormat = "0.00" & """" & "ea" & """"
            objSheet.Range("P" & i + 3).NumberFormat = "0.00" & """" & "ea" & """"
            objSheet.Range("Q" & i + 3).NumberFormat = "$* #,##0.00" & """" & "/ea" & """"
            
            objSheet.Cells(i + 3, 7) = "N"
            objSheet.Cells(i + 3, 8) = "0"
            objSheet.Cells(i + 3, 9) = "0"
            objSheet.Cells(i + 3, 10) = "0"
            objSheet.Cells(i + 3, 11) = "0"
            
        Else
            OP20 = swModel.GetCustomInfoValue("", "OP20")
            OP20_S = swModel.GetCustomInfoValue("", "OP20_S")
            OP20_R = swModel.GetCustomInfoValue("", "OP20_R")
            OptiMat = swModel.GetCustomInfoValue("", "OptiMaterial")
            
            If Not swModel.GetCustomInfoValue("", "RawWeight") = "" Then
                WeightLength = swModel.GetCustomInfoValue("", "RawWeight")
                objSheet.Range("C" & i + 3).NumberFormat = "0.00" & """" & "lb" & """"
                objSheet.Range("P" & i + 3).NumberFormat = "0.00" & """" & "lb" & """"
                objSheet.Range("Q" & i + 3).NumberFormat = "$* #,##0.00" & """" & "/lb" & """"
            Else
                WeightLength = swModel.GetCustomInfoValue("", "F300_Length")
                objSheet.Range("C" & i + 3).NumberFormat = "0.00" & """" & "in" & """"
                objSheet.Range("P" & i + 3).NumberFormat = "0.00" & """" & "in" & """"
                objSheet.Range("Q" & i + 3).NumberFormat = "$* #,##0.00" & """" & "/in" & """"
            End If
            
            If swModel.GetCustomInfoValue("", "Pressbrake") = "Checked" Then
                objSheet.Cells(i + 3, 7) = "Y"
                objSheet.Cells(i + 3, 8) = swModel.GetCustomInfoValue("", "F140_S")
                objSheet.Cells(i + 3, 9) = swModel.GetCustomInfoValue("", "F140_R")
            Else
                objSheet.Cells(i + 3, 7) = "N"
                objSheet.Cells(i + 3, 8) = "0"
                objSheet.Cells(i + 3, 9) = "0"
            End If
            
            objSheet.Cells(i + 3, 10) = swModel.GetCustomInfoValue("", "F220_S") + swModel.GetCustomInfoValue("", "F210_S") + swModel.GetCustomInfoValue("", "F325_S") + _
                    swModel.GetCustomInfoValue("", "Other_S") + swModel.GetCustomInfoValue("", "Other_S2")
            objSheet.Cells(i + 3, 11) = swModel.GetCustomInfoValue("", "F220_R") + swModel.GetCustomInfoValue("", "F210_R") + swModel.GetCustomInfoValue("", "F325_R") + _
                    swModel.GetCustomInfoValue("", "Other_R") + swModel.GetCustomInfoValue("", "Other_R2")
            
        End If
        
        objSheet.Cells(i + 3, 2) = OptiMat
        objSheet.Cells(i + 3, 3) = WeightLength
        objSheet.Cells(i + 3, 4) = OP20
        objSheet.Cells(i + 3, 5) = OP20_S
        objSheet.Cells(i + 3, 6) = OP20_R
        
        swApp.CloseAllDocuments True
        
    Next i
    
    objWorkbook.Save
    objWorkbook.Close True, ExcelFile
    Set objWorkbook = Nothing
    Set objBendTable = Nothing
    'Shell "TASKKILL /F /IM Excel.exe", vbHide
    Set objExcel = Nothing
    
    'swApp.UserControl = True
    'swApp.Visible = True

End Sub

Sub ExtractTubeData()
    
    Dim swLib As ExternalStart.iExternalStart
    Set swLib = New ExternalStart.ExternalStart
    swLib.ExternalStart swModel

End Sub

Sub ValidateFlatPattern()
    'Set swModel = swApp.ActivateDoc(FileName)
    'Set swPart = swModel
    Set swSelMgr = swPart.SelectionManager
    Set swSelData = swSelMgr.CreateSelectData
    Set swFeatMgr = swModel.FeatureManager
    
    Call FindFlatPattern    ' Procedure to find flat pattern feature
        
    If Not swFeat Is Nothing Then swFeat.Select (False)   'seems like could just be Set swFeat = Nothing
    swModel.EditUnsuppress2
    biggestArea = 0
    vBodies = swPart.GetBodies2(swAllBodies, True)
    Set swBody = vBodies(0) 'inserted to prevent losing body reference, causing automation error
    Set swFace = swBody.GetFirstFace 'inserted to prevent losing face reference, causing automation error
    swPart.ClearSelection2 True
        
    Do While Not swFace Is Nothing

        currentArea = swFace.GetArea

        If currentArea > biggestArea Then
            biggestArea = currentArea
            swPart.ClearSelection2 True
            Set swEnt = Nothing
            Set swEnt = swFace
            bRet = swEnt.Select4(True, swSelData) ': Debug.Assert bRet
        End If
'                                Set swFace = Nothing
        Set swFace = swFace.GetNextFace
    Loop
    Set swFace = swSelMgr.GetSelectedObject6(1, -1)
    biggestArea = swFace.GetArea

    Thickness = 0
    'Set swModel = swApp.ActiveDoc
    swModel.DeleteCustomInfo "SMThick"
    swModel.AddCustomInfo3 gstrConfigName, "SMThick", 30, """Thickness@$PRP:""SW-File Name"".SLDPRT"""
    strThickness = swModel.GetCustomInfoValue(gstrConfigName, "SMThick")
    badThickness = "SLDPRT" & """"
    strThicknessCheck = UCase(Right(strThickness, 7))
    If strThicknessCheck <> badThickness Then
        Thickness = strThickness * 0.0254
    End If
    
    
End Sub

Sub SaveCurrentModel()

    If Not swModel Is Nothing Then
        swModel.Save3 1, 0, 0
        swModel.ClearSelection2 True
    End If

End Sub

Sub GetLargestFace()
    swPart.ClearSelection2 True
    Set swFace = Nothing
    Set swFace = swBody.GetFirstFace
    swPart.ClearSelection2 True
    biggestArea = 0
    Do While Not swFace Is Nothing
        currentArea = swFace.GetArea
        If currentArea > biggestArea Then
            biggestArea = currentArea
            swPart.ClearSelection2 True
            Set swEnt = Nothing
            Set swEnt = swFace
            bRet = swEnt.Select4(True, swSelData) ': Debug.Assert bRet
        End If
'       Set swFace = Nothing
        Set swFace = swFace.GetNextFace
    Loop
    Set swFace = Nothing
    Set swFace = swEnt
    Set swSurface = Nothing
    Set swSurface = swFace.GetSurface
    
End Sub

Sub TubeCustomProperties()

    Dim dblCrossSection As Double
    Dim dblCrossSectionLeft As Double
    Dim dblCrossSectionRight As Double
    Dim dblThickness As Double
    Dim dblLength As Double
    Dim dblCutLength As Double
    Dim dblHoles As Double

    On Error Resume Next
    Shape = swModel.GetCustomInfoValue("", "Shape")
    If Err <> 0 Then
        Set swModel = Nothing
        Set swModel = swApp.OpenDoc6(FileName, 1, 1, "Default", longerrors, longwarnings)
        Shape = swModel.GetCustomInfoValue("", "Shape")
    End If
    CrossSection = swModel.GetCustomInfoValue("", "CrossSection")
    
    If Shape = "Round" Then
        intPosition = InStrRev(CrossSection, "�")
        IntLength = Len(CrossSection)
        CrossSection = Right(CrossSection, IntLength - intPosition)
        If Not InStrRev(CrossSection, "10.") = 0 Then 'filter for possible 10.xx diameter
            dblCrossSection = CrossSection
        Else
            intPosition = InStrRev(CrossSection, "0.")
            IntLength = Len(CrossSection)
            CrossSection = Right(CrossSection, IntLength - intPosition)
            dblCrossSection = CrossSection
        End If

    Else
        intPosition = InStrRev(CrossSection, " x ")
        IntLength = Len(CrossSection)
        While Not intPosition = 0
            CrossSectionLeft = Left(CrossSection, intPosition - 1)
            dblCrossSectionLeft = CrossSectionLeft
            CrossSectionRight = Right(CrossSection, IntLength - (intPosition + 2))
            dblCrossSectionRight = CrossSectionRight
            intPosition = 0
        Wend
        intPosition = InStrRev(CrossSectionLeft, "0.")
        IntLength = Len(CrossSectionLeft)
        CrossSectionLeft = Right(CrossSectionLeft, IntLength - intPosition)
        intPosition = InStrRev(CrossSectionRight, "0.")
        IntLength = Len(CrossSectionRight)
        CrossSectionRight = Right(CrossSectionRight, IntLength - intPosition)
    End If
    
    WallThickness = swModel.GetCustomInfoValue("", "Wall Thickness")
    dblThickness = WallThickness
    intPosition = InStrRev(WallThickness, "0.")
    IntLength = Len(WallThickness)
    WallThickness = Right(WallThickness, IntLength - intPosition)
    
    If Len(WallThickness) = 3 And intPosition = InStrRev(WallThickness, "5") <> 1 Then
        WallThickness = WallThickness & "0"
    End If
    
    MaterialLength = swModel.GetCustomInfoValue("", "Material Length")
    dblLength = MaterialLength
    
    CutLength = swModel.GetCustomInfoValue("", "Cut Length")
    dblCutLength = CutLength
    
    Holes = swModel.GetCustomInfoValue("", "Number of Holes")
    dblHoles = Holes
    
    swModel.CustomInfo2("", "OP20_S") = "0"
    swModel.CustomInfo2("", "OP20_R") = "0"
    TubeTime = 0
    
    Select Case Shape
    
    Case "Round"
        If dblThickness > dblCrossSection * 0.3 Or dblThickness = 0 Then
            Call RoundBar(dblLength, dblCrossSection, CrossSection)
            Exit Sub
        Else
            Call PipeDiam(dblCrossSection, dblThickness)
            
            Select Case TubePrefix
                Case "P."
                    If modSetMaterial.strMaterial = "A36" Or modSetMaterial.strMaterial = "ALNZD" Then
                        TubeMaterial = "BLK"
                    Else
                        TubeMaterial = modSetMaterial.strMaterial
                    End If
                Case "T."
                    If modSetMaterial.strMaterial = "A36" Or modSetMaterial.strMaterial = "ALNZD" Then
                        TubeMaterial = "HR"
                    Else
                        TubeMaterial = modSetMaterial.strMaterial
                    End If
            End Select
            
            If swModel.GetCustomInfoValue("", "OptiMaterial") = "" Then swModel.CustomInfo2("", "OptiMaterial") = TubePrefix & TubeMaterial & CrossSection & WallThickness
            
            If TubePrefix = "P." Then
                If swModel.GetCustomInfoValue("", "Description") = "" Then
                    WriteDescription = True
                    swModel.CustomInfo2("", "Description") = TubeMaterial & " PIPE"
                Else
                    WriteDescription = False
                End If
                
                swModel.CustomInfo2("", "rbMaterialType") = "1"
                swModel.DeleteCustomInfo2 "", "F300_Length"
                swModel.AddCustomInfo3 "", "F300_Length", 3, dblLength
                '***** ADDED ele if statements for 10in and 6in 2/2/2023
                If dblCrossSection <= 6 Then
                    swModel.CustomInfo2("", "OP20") = "F110 - TUBE LASER"
                    swModel.CustomInfo2("", "OP20_S") = ".15"
                ElseIf dblCrossSection <= 10 Then
                    swModel.CustomInfo2("", "OP20") = "F110 - TUBE LASER"
                    swModel.CustomInfo2("", "OP20_S") = ".5"
                ElseIf dblCrossSection <= 10.75 Then
                    swModel.CustomInfo2("", "OP20") = "F110 - TUBE LASER"
                    swModel.CustomInfo2("", "OP20_S") = "1"
                Else
                    swModel.CustomInfo2("", "OP20") = "N145 - 5-AXIS LASER"
                    swModel.CustomInfo2("", "OP20_S") = ".25"
                End If
                '*************END
            Else
                If swModel.GetCustomInfoValue("", "Description") = "" Then
                    WriteDescription = True
                    swModel.CustomInfo2("", "Description") = TubeMaterial & " TUBE"
                Else
                    WriteDescription = False
                End If
                
                swModel.CustomInfo2("", "rbMaterialType") = "1"
                swModel.DeleteCustomInfo2 "", "F300_Length"
                swModel.AddCustomInfo3 "", "F300_Length", 3, dblLength
                '***** ADDED ele if statements for 10in and 6in 2/2/2023
                If dblCrossSection <= 6 Then
                    swModel.CustomInfo2("", "OP20") = "F110 - TUBE LASER"
                    swModel.CustomInfo2("", "OP20_S") = ".15"
                ElseIf dblCrossSection <= 10 Then
                    swModel.CustomInfo2("", "OP20") = "F110 - TUBE LASER"
                    swModel.CustomInfo2("", "OP20_S") = ".5"
                ElseIf dblCrossSection <= 10.75 Then
                    swModel.CustomInfo2("", "OP20") = "F110 - TUBE LASER"
                    swModel.CustomInfo2("", "OP20_S") = "1"
                Else
                    swModel.CustomInfo2("", "OP20") = "N145 - 5-AXIS LASER"
                    swModel.CustomInfo2("", "OP20_S") = ".25"
                End If
                '*************END
            End If
        End If
        
    Case "Square"
    
        If modSetMaterial.strMaterial = "A36" Or modSetMaterial.strMaterial = "ALNZD" Then
            TubeMaterial = "HR"
        Else
            TubeMaterial = modSetMaterial.strMaterial
        End If
        
        bolA = dblThickness > (dblCrossSectionLeft * 0.3)
        bolB = dblThickness > (dblCrossSectionRight * 0.3)
        If bolA <> True And bolB <> True And dblThickness <> 0 Then
            If swModel.GetCustomInfoValue("", "OptiMaterial") = "" Then swModel.CustomInfo2("", "OptiMaterial") = "T." & TubeMaterial & CrossSectionLeft & """" & "SQX" & WallThickness & """"
            
            If swModel.GetCustomInfoValue("", "Description") = "" Then
                WriteDescription = True
                swModel.CustomInfo2("", "Description") = TubeMaterial & " TUBE"
            Else
                WriteDescription = False
            End If
            
            swModel.CustomInfo2("", "rbMaterialType") = "1"
            swModel.DeleteCustomInfo2 "", "F300_Length"
            swModel.AddCustomInfo3 "", "F300_Length", 3, dblLength
            
             '***** ADDED else if statements for 10in and 6in 2/2/2023
            If dblCrossSectionLeft <= 6 Then
                swModel.CustomInfo2("", "OP20") = "F110 - TUBE LASER"
                swModel.CustomInfo2("", "OP20_S") = ".15"
            ElseIf dblCrossSectionLeft <= 10 Then
                swModel.CustomInfo2("", "OP20") = "F110 - TUBE LASER"
                swModel.CustomInfo2("", "OP20_S") = ".5"
            Else
                swModel.CustomInfo2("", "OP20") = "N145 - 5-AXIS LASER"
                swModel.CustomInfo2("", "OP20_S") = ".25"
            End If
            '****END
        Else
            Exit Sub
        End If
    Case "Rectangle"
    
        If modSetMaterial.strMaterial = "A36" Or modSetMaterial.strMaterial = "ALNZD" Then
            TubeMaterial = "HR"
        Else
            TubeMaterial = modSetMaterial.strMaterial
        End If
        
        bolA = dblThickness > (dblCrossSectionLeft * 0.3)
        bolB = dblThickness > (dblCrossSectionRight * 0.3)
        If bolA <> True And bolB <> True And dblThickness <> 0 Then
            If swModel.GetCustomInfoValue("", "OptiMaterial") = "" Then swModel.CustomInfo2("", "OptiMaterial") = "T." & TubeMaterial & CrossSectionLeft & """" & "X" & CrossSectionRight & """" & "X" & WallThickness & """"

            If swModel.GetCustomInfoValue("", "Description") = "" Then
                WriteDescription = True
                swModel.CustomInfo2("", "Description") = TubeMaterial & " TUBE"
            Else
                WriteDescription = False
            End If
                
            swModel.CustomInfo2("", "rbMaterialType") = "1"
            swModel.DeleteCustomInfo2 "", "F300_Length"
            swModel.AddCustomInfo3 "", "F300_Length", 3, dblLength
             '***** ADDED else if statements for 10in and 6in 2/2/2023
            If dblCrossSectionLeft <= 6 Then
                swModel.CustomInfo2("", "OP20") = "F110 - TUBE LASER"
                swModel.CustomInfo2("", "OP20_S") = ".15"
            ElseIf dblCrossSectionLeft <= 10 Then
                swModel.CustomInfo2("", "OP20") = "F110 - TUBE LASER"
                swModel.CustomInfo2("", "OP20_S") = ".5"
            Else
                swModel.CustomInfo2("", "OP20") = "N145 - 5-AXIS LASER"
                swModel.CustomInfo2("", "OP20_S") = ".25"
            End If
            '****END
        Else
            Exit Sub
        End If
    
    Case "Angle"
    
        If modSetMaterial.strMaterial = "A36" Or modSetMaterial.strMaterial = "ALNZD" Then
            TubeMaterial = "HR"
        Else
            TubeMaterial = modSetMaterial.strMaterial
        End If
        
        bolA = dblThickness > (dblCrossSectionLeft * 0.3)
        bolB = dblThickness > (dblCrossSectionRight * 0.3)
        If bolA <> True And bolB <> True And dblThickness <> 0 Then
            If swModel.GetCustomInfoValue("", "OptiMaterial") = "" Then swModel.CustomInfo2("", "OptiMaterial") = "A." & TubeMaterial & CrossSectionLeft & """" & "X" & CrossSectionRight & """" & "X" & WallThickness & """"
            
            If swModel.GetCustomInfoValue("", "Description") = "" Then
                WriteDescription = True
                swModel.CustomInfo2("", "Description") = TubeMaterial & " ANGLE"
            Else
                WriteDescription = False
            End If
            
            swModel.CustomInfo2("", "rbMaterialType") = "1"
            swModel.DeleteCustomInfo2 "", "F300_Length"
            swModel.AddCustomInfo3 "", "F300_Length", 3, dblLength
            
            If dblCrossSectionLeft <= 6 Then
                swModel.CustomInfo2("", "OP20") = "F110 - TUBE LASER"
                swModel.CustomInfo2("", "OP20_S") = ".25"   '<--- 2/22/23 increased time
            Else
                swModel.CustomInfo2("", "OP20") = "N145 - 5-AXIS LASER"
                swModel.CustomInfo2("", "OP20_S") = ".25"
            End If
        Else
            Exit Sub
        End If
    
    Case "Channel"
    
        If modSetMaterial.strMaterial = "A36" Or modSetMaterial.strMaterial = "ALNZD" Then
            TubeMaterial = "HR"
        Else
            TubeMaterial = modSetMaterial.strMaterial
        End If
        
        bolA = dblThickness > (dblCrossSectionLeft * 0.3)
        bolB = dblThickness > (dblCrossSectionRight * 0.3)
        If bolA <> True And bolB <> True And dblThickness <> 0 Then
            If swModel.GetCustomInfoValue("", "OptiMaterial") = "" Then swModel.CustomInfo2("", "OptiMaterial") = "C." & TubeMaterial & CrossSectionLeft & """" & "X" & CrossSectionRight & """" & "X" & WallThickness & """"
            
            If swModel.GetCustomInfoValue("", "Description") = "" Then
                WriteDescription = True
                swModel.CustomInfo2("", "Description") = TubeMaterial & " ANGLE"
            Else
                WriteDescription = False
            End If
            
            swModel.CustomInfo2("", "rbMaterialType") = "1"
            swModel.DeleteCustomInfo2 "", "F300_Length"
            swModel.AddCustomInfo3 "", "F300_Length", 3, dblLength
            
            If dblCrossSectionLeft <= 6 Then
                swModel.CustomInfo2("", "OP20") = "F110 - TUBE LASER"
                swModel.CustomInfo2("", "OP20_S") = ".25"   '<--- 2/22/23 increased time
            Else
                swModel.CustomInfo2("", "OP20") = "N145 - 5-AXIS LASER"
                swModel.CustomInfo2("", "OP20_S") = ".25"
            End If
        Else
            Exit Sub
        End If
    
    End Select
    
    CutTime = (dblCutLength / TubeFeedRate(dblThickness)) * 60
    CycleTime = ((dblLength / 240) * 45)
    PierceTime = (dblHoles + 2) * (TubePierceTime(dblThickness) + 1.5)
    TraverseTime = (dblLength / 1440) * 60
    TubeTime = Round((CutTime + CycleTime + PierceTime + TraverseTime) / 3600, 4)
    
    
    '2/2/2023 update tube laser times
    '*************************
    Dim dblTubeWeight As Double
    Dim dblTubeSetup As Double
    
    
    dblTubeWeight = swModel.GetCustomInfoValue("", "Weight")

    If Shape = "Angle" Or Shape = "Channel" Then
        dblTubeSetup = swModel.GetCustomInfoValue("", "OP20_S")
        
        dblTubeSetup = dblTubeSetup + 0.25
        TubeTime = TubeTime * 3
        
        swModel.CustomInfo2("", "OP20_S") = dblTubeSetup
    
    End If
    
    If WallThickness > 0.2 Or WallThickness <> "ERROR" Then TubeTime = TubeTime * 2
    If dblTubeWeight > 50 And TubeTime < 0.05 Then TubeTime = 0.05
    
    '****************************
    
    If TubeTime > 0.001 Then swModel.CustomInfo2("", "OP20_R") = TubeTime
        

    
End Sub

Sub RoundBar(dblLength As Double, dblCrossSection As Double, CrossSection As String)
    Set swMass = swModel.Extension.CreateMassProperty
    swVolume = Round(swMass.Volume, 15)
    calcVolume = Round((3.14159265358979 * (((dblCrossSection / 2) * 0.0254) ^ 2) * (dblLength * 0.0254)), 15)
    swVolumeUP = Round(swVolume * 1.03, 15)
    swVolumeDN = Round(swVolume * 0.97, 15)
    If swVolumeUP > calcVolume And swVolumeDN < calcVolume Then
        swModel.CustomInfo2("", "OP20") = "F300 - SAW"
        swModel.CustomInfo2("", "rbMaterialType") = "1"
        swModel.DeleteCustomInfo2 "", "F300_Length"
        swModel.AddCustomInfo3 "", "F300_Length", 3, dblLength
        swModel.CustomInfo2("", "OP20_S") = ".05"
        swModel.CustomInfo2("", "OP20_R") = Round(((dblCrossSection * 90) + 15) / 3600, 4)
        If dblCrossSection = 0.125 Or dblCrossSection = 0.1563 Or dblCrossSection = 0.1875 Or dblCrossSection = 0.25 Or dblCrossSection = 0.3125 Or _
            dblCrossSection = 0.375 Or dblCrossSection = 0.4375 Or dblCrossSection = 0.5 Or dblCrossSection = 0.625 Or dblCrossSection = 0.75 Or _
            dblCrossSection = 0.875 Or dblCrossSection = 0.9375 Or dblCrossSection = 1 Or dblCrossSection = 1.1875 Or dblCrossSection = 1.25 Or _
            dblCrossSection = 1.375 Or dblCrossSection = 1.5 Or dblCrossSection = 1.75 Or dblCrossSection = 2 Or dblCrossSection = 2.5 Or _
            dblCrossSection = 3 Then
            If swModel.GetCustomInfoValue("", "OptiMaterial") = "" Then swModel.CustomInfo2("", "OptiMaterial") = "R." & modSetMaterial.strMaterial & CrossSection & """"
        Else
            If swModel.GetCustomInfoValue("", "OptiMaterial") = "" Then swModel.CustomInfo2("", "OptiMaterial") = "R." & modSetMaterial.strMaterial & "M" & Int(dblCrossSection * 25.4)
        End If
        If swModel.GetCustomInfoValue("", "Description") = "" Then swModel.CustomInfo2("", "Description") = modSetMaterial.strMaterial & " ROUND"
        swModel.ForceRebuild3 True
        Call SaveCurrentModel
    End If
End Sub
Function PipeDiam(dblCrossSection As Double, dblThickness)

Select Case dblCrossSection
    
    Case 0.405
        CrossSection = ".125" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.049
                WallThickness = "10"
            Case 0.068
                WallThickness = "40"
            Case 0.095
                WallThickness = "80"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 0.54
        CrossSection = ".25" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.065
                WallThickness = "10"
            Case 0.088
                WallThickness = "40"
            Case 0.119
                WallThickness = "80"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 0.675
        CrossSection = ".375" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.065
                WallThickness = "10"
            Case 0.091
                WallThickness = "40"
            Case 0.126
                WallThickness = "80"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 0.84
        CrossSection = ".5" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.065
                WallThickness = "5"
            Case 0.083
                WallThickness = "10"
            Case 0.109
                WallThickness = "40"
            Case 0.147
                WallThickness = "80"
            Case 0.187
                WallThickness = "160"
            Case 0.294
                WallThickness = "XX"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 1.05
        CrossSection = ".75" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.065
                WallThickness = "5"
            Case 0.083
                WallThickness = "10"
            Case 0.113
                WallThickness = "40"
            Case 0.154
                WallThickness = "80"
            Case 0.218
                WallThickness = "160"
            Case 0.308
                WallThickness = "XX"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 1.315
        CrossSection = "1" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.065
                WallThickness = "5"
            Case 0.109
                WallThickness = "10"
            Case 0.133
                WallThickness = "40"
            Case 0.179
                WallThickness = "80"
            Case 0.25
                WallThickness = "160"
            Case 0.358
                WallThickness = "XX"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 1.66
        CrossSection = "1.25" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.065
                WallThickness = "5"
            Case 0.109
                WallThickness = "10"
            Case 0.14
                WallThickness = "40"
            Case 0.191
                WallThickness = "80"
            Case 0.25
                WallThickness = "160"
            Case 0.382
                WallThickness = "XX"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 1.9
        CrossSection = "1.5" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.065
                WallThickness = "5"
            Case 0.109
                WallThickness = "10"
            Case 0.145
                WallThickness = "40"
            Case 0.2
                WallThickness = "80"
            Case 0.281
                WallThickness = "160"
            Case 0.4
                WallThickness = "XX"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 2.375
        CrossSection = "2" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.065
                WallThickness = "5"
            Case 0.12
                WallThickness = "10"
            Case 0.154
                WallThickness = "40"
            Case 0.218
                WallThickness = "80"
            Case 0.344
                WallThickness = "160"
            Case 0.436
                WallThickness = "XX"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 2.875
        CrossSection = "2.5" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.083
                WallThickness = "5"
            Case 0.12
                WallThickness = "10"
            Case 0.203
                WallThickness = "40"
            Case 0.276
                WallThickness = "80"
            Case 0.375
                WallThickness = "160"
            Case 0.552
                WallThickness = "XX"
            Case Else
                WallThickness = "ERROR"
        End Select
      
    Case 3.5
        CrossSection = "3" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.083
                WallThickness = "5"
            Case 0.12
                WallThickness = "10"
            Case 0.216
                WallThickness = "40"
            Case 0.3
                WallThickness = "80"
            Case 0.438
                WallThickness = "160"
            Case 0.6
                WallThickness = "XX"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 4
        CrossSection = "3.5" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.083
                WallThickness = "5"
            Case 0.12
                WallThickness = "10"
            Case 0.226
                WallThickness = "40"
            Case 0.318
                WallThickness = "80"
            Case 0.636
                WallThickness = "XX"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 4.5
        CrossSection = "4" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.083
                WallThickness = "5"
            Case 0.12
                WallThickness = "10"
            Case 0.237
                WallThickness = "40"
            Case 0.337
                WallThickness = "80"
            Case 0.438
                WallThickness = "120"
            Case 0.531
                WallThickness = "160"
            Case 0.674
                WallThickness = "XX"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 5
        CrossSection = "4.5" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.247
                WallThickness = "STD"
            Case 0.12
                WallThickness = "XX"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 5.563
        CrossSection = "5" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.109
                WallThickness = "5"
            Case 0.134
                WallThickness = "10"
            Case 0.258
                WallThickness = "40"
            Case 0.375
                WallThickness = "80"
            Case 0.5
                WallThickness = "120"
            Case 0.625
                WallThickness = "160"
            Case 0.75
                WallThickness = "XX"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 6.625
        CrossSection = "6" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.109
                WallThickness = "5"
            Case 0.134
                WallThickness = "10"
            Case 0.28
                WallThickness = "40"
            Case 0.432
                WallThickness = "80"
            Case 0.562
                WallThickness = "120"
            Case 0.718
                WallThickness = "160"
            Case 0.864
                WallThickness = "XX"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 8.625
        CrossSection = "8" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.109
                WallThickness = "5"
            Case 0.148
                WallThickness = "10"
            Case 0.322
                WallThickness = "40"
            Case 0.5
                WallThickness = "80"
            Case 0.718
                WallThickness = "120"
            Case 0.906
                WallThickness = "160"
            Case 0.875
                WallThickness = "XX"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 10.75
        CrossSection = "10" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.134
                WallThickness = "5"
            Case 0.165
                WallThickness = "10"
            Case 0.365
                WallThickness = "40"
            Case 0.5
                WallThickness = "80S"
            Case 0.593
                WallThickness = "80"
            Case 0.843
                WallThickness = "120"
            Case 1.125
                WallThickness = "160"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 12.75
        CrossSection = "12" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.156
                WallThickness = "5"
            Case 0.18
                WallThickness = "10"
            Case 0.375
                WallThickness = "40S"
            Case 0.406
                WallThickness = "40"
            Case 0.5
                WallThickness = "80S"
            Case 0.687
                WallThickness = "80"
            Case 1
                WallThickness = "120"
            Case 1.312
                WallThickness = "160"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 14
        CrossSection = "14" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.156
                WallThickness = "5"
            Case 0.188
                WallThickness = "10S"
            Case 0.25
                WallThickness = "10"
            Case 0.375
                WallThickness = "40S"
            Case 0.437
                WallThickness = "40"
            Case 0.5
                WallThickness = "80S"
            Case 0.75
                WallThickness = "80"
            Case 1.093
                WallThickness = "120"
            Case 1.406
                WallThickness = "160"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 16
        CrossSection = "16" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.156
                WallThickness = "5"
            Case 0.188
                WallThickness = "10S"
            Case 0.25
                WallThickness = "10"
            Case 0.375
                WallThickness = "40S"
            Case 0.5
                If frmMaterialUpdate.obSS.value = True Then
                    WallThickness = "80S"
                Else
                    WallThickness = "40"
                End If
            Case 0.843
                WallThickness = "80"
            Case 1.218
                WallThickness = "120"
            Case 1.437
                WallThickness = "160"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 18
        CrossSection = "18" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.165
                WallThickness = "5"
            Case 0.188
                WallThickness = "10S"
            Case 0.25
                WallThickness = "10"
            Case 0.375
                WallThickness = "40S"
            Case 0.562
                WallThickness = "40"
            Case 0.5
                WallThickness = "80S"
            Case 0.937
                WallThickness = "80"
            Case 1.375
                WallThickness = "120"
            Case 1.781
                WallThickness = "160"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 20
        CrossSection = "20" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.188
                WallThickness = "5"
            Case 0.218
                WallThickness = "10S"
            Case 0.25
                WallThickness = "10"
            Case 0.375
                WallThickness = "40S"
            Case 0.593
                WallThickness = "40"
            Case 0.5
                WallThickness = "80S"
            Case 1.031
                WallThickness = "80"
            Case 1.5
                WallThickness = "120"
            Case 1.968
                WallThickness = "160"
            Case Else
                WallThickness = "ERROR"
        End Select
        
    Case 24
        CrossSection = "24" & """" & "SCH"
        TubePrefix = "P."
        
        Select Case dblThickness
            Case 0.218
                WallThickness = "5"
            Case 0.25
                WallThickness = "10"
            Case 0.375
                WallThickness = "40S"
            Case 0.687
                WallThickness = "40"
            Case 0.5
                WallThickness = "80S"
            Case 1.218
                WallThickness = "80"
            Case 1.812
                WallThickness = "120"
            Case 2.343
                WallThickness = "160"
            Case Else
                WallThickness = "ERROR"
        End Select
    Case Else
        CrossSection = CrossSection & """" & "ODX"
        TubePrefix = "T."
        WallThickness = WallThickness & """"
    End Select
    

End Function

Function TubeFeedRate(dblThickness As Double) As Double

If frmMaterialUpdate.obCS.value = True Then

    Select Case dblThickness
        Case 0 To 0.045
            TubeFeedRate = 295
        Case 0.045 To 0.055
            TubeFeedRate = 271
        Case 0.055 To 0.065
            TubeFeedRate = 251
        Case 0.065 To 0.085
            TubeFeedRate = 196
        Case 0.085 To 0.105
            TubeFeedRate = 161
        Case 0.105 To 0.125
            TubeFeedRate = 149
        Case 0.125 To 0.145
            TubeFeedRate = 137
        Case 0.145 To 0.165
            TubeFeedRate = 129
        Case 0.165 To 0.185
            TubeFeedRate = 122
        Case 0.185 To 0.205
            TubeFeedRate = 118
        Case 0.205 To 0.255
            TubeFeedRate = 106
        Case 0.255 To 0.32
            TubeFeedRate = 87
        Case 0.32 To 0.38
            TubeFeedRate = 47
        Case 0.38 To 0.405
            TubeFeedRate = 36
        Case 0.405 To 0.455
            TubeFeedRate = 19
        Case Is > 0.455
            TubeFeedRate = 7
    End Select
    
ElseIf frmMaterialUpdate.obSS.value = True Then
    
    Select Case dblThickness
        Case 0 To 0.045
            TubeFeedRate = 397
        Case 0.045 To 0.055
            TubeFeedRate = 354
        Case 0.055 To 0.065
            TubeFeedRate = 318
        Case 0.065 To 0.085
            TubeFeedRate = 251
        Case 0.085 To 0.105
            TubeFeedRate = 196
        Case 0.105 To 0.125
            TubeFeedRate = 157
        Case 0.125 To 0.145
            TubeFeedRate = 135
        Case 0.145 To 0.165
            TubeFeedRate = 118
        Case 0.165 To 0.185
            TubeFeedRate = 104
        Case 0.185 To 0.205
            TubeFeedRate = 90
        Case 0.205 To 0.255
            TubeFeedRate = 78
        Case 0.255 To 0.32
            TubeFeedRate = 54
        Case 0.32 To 0.38
            TubeFeedRate = 36
        Case 0.38 To 0.405
            TubeFeedRate = 30
        Case 0.405 To 0.455
            TubeFeedRate = 18
        Case Is > 0.455
            TubeFeedRate = 8
    End Select

ElseIf frmMaterialUpdate.obAL.value = True Then
        TubeFeedRate = 0
        
End If

TubeFeedRate = TubeFeedRate * 0.85

End Function
Function TubePierceTime(dblThickness As Double) As Double

If frmMaterialUpdate.obCS.value = True Then

    Select Case dblThickness
        Case 0 To 0.045
            TubePierceTime = 0.05
        Case 0.045 To 0.055
            TubePierceTime = 0.05
        Case 0.055 To 0.065
            TubePierceTime = 0.05
        Case 0.065 To 0.085
            TubePierceTime = 0.05
        Case 0.085 To 0.105
            TubePierceTime = 0.07
        Case 0.105 To 0.125
            TubePierceTime = 0.1
        Case 0.125 To 0.145
            TubePierceTime = 0.2
        Case 0.145 To 0.165
            TubePierceTime = 0.3
        Case 0.165 To 0.185
            TubePierceTime = 0.4
        Case 0.185 To 0.205
            TubePierceTime = 0.5
        Case 0.205 To 0.255
            TubePierceTime = 0.7
        Case 0.255 To 0.32
            TubePierceTime = 2.8
        Case 0.32 To 0.38
            TubePierceTime = 5
        Case 0.38 To 0.405
            TubePierceTime = 5
        Case 0.405 To 0.455
            TubePierceTime = 5
        Case Is > 0.455
            TubePierceTime = 5
    End Select
    
ElseIf frmMaterialUpdate.obSS.value = True Then
    
    Select Case dblThickness
        Case 0 To 0.045
            TubePierceTime = 0.05
        Case 0.045 To 0.055
            TubePierceTime = 0.05
        Case 0.055 To 0.065
            TubePierceTime = 0.05
        Case 0.065 To 0.085
            TubePierceTime = 0.05
        Case 0.085 To 0.105
            TubePierceTime = 0.07
        Case 0.105 To 0.125
            TubePierceTime = 0.08
        Case 0.125 To 0.145
            TubePierceTime = 0.45
        Case 0.145 To 0.165
            TubePierceTime = 0.6
        Case 0.165 To 0.185
            TubePierceTime = 0.6
        Case 0.185 To 0.205
            TubePierceTime = 0.6
        Case 0.205 To 0.255
            TubePierceTime = 2
        Case 0.255 To 0.32
            TubePierceTime = 3
        Case 0.32 To 0.38
            TubePierceTime = 5
        Case 0.38 To 0.405
            TubePierceTime = 5
        Case 0.405 To 0.455
            TubePierceTime = 5
        Case Is > 0.455
            TubePierceTime = 5
    End Select

ElseIf frmMaterialUpdate.obAL.value = True Then
        TubePierceTime = 0
        
End If

End Function
Sub ExGeo()

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    GeoName = Left(swModel.GetTitle, InStrRev(swModel.GetTitle, ".") - 1)
    Dim TopsWorks As ExternalStart.iExternalStart
    Set TopsWorks = New ExternalStart.ExternalStart
    TopsWorks.ExportGeo GeoName

End Sub
Sub N325()
    
    Dim swBend As BendsFeatureData
    Dim swSubFeat As SldWorks.Feature
    Set swFeat = swModel.FirstFeature
    
    While Not swFeat Is Nothing
        Feature = swFeat.GetTypeName
        
        If swFeat.GetTypeName = "SheetMetal" Then
            Set swSheetMetal = swFeat.GetDefinition
            Thickness = Round(swSheetMetal.Thickness * 39.370078, 4)
        End If
        
        Set swSubFeat = swFeat.GetFirstSubFeature
        
        Do While Not swSubFeat Is Nothing
            SubFeature = swSubFeat.GetTypeName
            'Debug.Print swSubFeat.GetTypeName
            Select Case swSubFeat.GetTypeName
            Case "OneBend"
                Set swOneBend = swSubFeat.GetDefinition
                Radius = Round(swOneBend.BendRadius * 39.370078, 3)
                
                If Radius > maxRadius Then maxRadius = Radius
                'not a bend
            Case "SketchBend"
                Set swOneBend = swSubFeat.GetDefinition
                Radius = Round(swOneBend.BendRadius * 39.370078, 3)
                
                If Radius > maxRadius Then maxRadius = Radius
                'not a bend
            Case Else
            End Select
            
            Set swSubFeat = swSubFeat.GetNextSubFeature()
        Loop
        
        Set swFeat = swFeat.GetNextFeature
    Wend
    
    If maxRadius > 2 Then
        CalcN325 (Thickness)
    End If
End Sub

Function CalcN325(Thickness As String)

'Set swApp = Application.SldWorks
'Set swModel = swApp.ActiveDoc
Set swMass = swModel.Extension.CreateMassProperty

ModelWeight = swMass.Mass * 2.20462

swModel.AddCustomInfo2 "F325_R", 30, ""
swModel.AddCustomInfo2 "F325_S", 30, ""

swModel.CustomInfo2("", "F325") = "1"
swModel.CustomInfo2("", "F325_R") = Round((((swMass.Mass * 2.20462) * (5 / 3600)) + (5 / 60)), 3)

If ModelWeight < 40 Or Thickness < 0.165 Then
    swModel.CustomInfo2("", "F325_S") = ".25"
    swModel.CustomInfo2("", "PressBrake") = "Unchecked"
End If

If ModelWeight < 150 And ModelWeight >= 40 And Thickness < 0.165 Then
    swModel.CustomInfo2("", "F325_S") = ".375"
    swModel.CustomInfo2("", "PressBrake") = "Unchecked"
End If

If ModelWeight < 150 And ModelWeight >= 40 And Thickness >= 0.165 Then
    swModel.CustomInfo2("", "F325_S") = ".375"
    swModel.CustomInfo2("", "PressBrake") = "Checked"
    swModel.CustomInfo2("", "F140_S") = ".2"
    swModel.CustomInfo2("", "F140_R") = ".08"
End If

If ModelWeight >= 150 And Thickness >= 0.165 Then
    swModel.CustomInfo2("", "F325_S") = ".75"
    swModel.CustomInfo2("", "PressBrake") = "Checked"
    swModel.CustomInfo2("", "F140_S") = ".2"
    swModel.CustomInfo2("", "F140_R") = ".25"
End If

swModel.ForceRebuild3 (False)
swModel.ForceRebuild3 (False)

End Function

Sub GetLinearEdge()
    Dim swCurve As SldWorks.Curve
    Dim swEdge As SldWorks.Edge
    Dim vEdge As Variant
    Dim vCurveParam As Variant
    
    On Error GoTo ErrorHandler

    swPart.ClearSelection2 True
    Set swEdge = Nothing
    vEdge = swBody.GetEdges
    longestedge = 0
    swApp.DocumentVisible False, swDocPART
    If UBound(vEdge) = 0 Then Exit Sub 'PART WITH NO EDGES
    For i = LBound(vEdge) To UBound(vEdge)
        Set swEdge = vEdge(i)
        Set swCurve = swEdge.GetCurve
        'Debug.Print swCurve.Identity
        If swCurve.Identity = 3001 Then
            vCurveParam = swEdge.GetCurveParams2
            currentedge = swCurve.GetLength2(vCurveParam(6), vCurveParam(7))
            If currentedge > longestedge Then
                longestedge = currentedge
                swPart.ClearSelection2 True
                Set swEnt = swEdge
                bRet = swEnt.Select4(True, swSelData)
            End If
        End If
    Next i

Exit Sub
ErrorHandler:
Exit Sub

End Sub

Sub BendAllowanceType()

    Dim swSheetMetal            As SldWorks.SheetMetalFeatureData
    Dim bRet                    As Boolean
    Dim vFeats                  As Variant
    Dim swComponent             As SldWorks.Component2
    
    Set swSelMgr = swModel.SelectionManager
    Set swFeatMgr = swModel.FeatureManager
    Set swSelMgr = swModel.SelectionManager
    
    swModel.Extension.SelectByID2 "Sheet-Metal", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0
    Set swFeat = swSelMgr.GetSelectedObject6(1, 0)
    Set swSelMgr = Nothing
    
    If swFeat Is Nothing Then
    
        Set swFeat = swModel.FirstFeature
    
        vFeats = swFeatMgr.GetFeatures(False)
        ReDim vFeatureName(0 To UBound(vFeats)) As String
    
        For f = LBound(vFeats) To UBound(vFeats)
            Set swFeat = Nothing
            Set swFeat = vFeats(f)
            FeatureName = swFeat.GetTypeName
            vFeatureName(f) = FeatureName
            'Debug.Print swFeat.Name
            'Debug.Print FeatureName
            If FeatureName = "SheetMetal" Then Exit For
        Next f
    
    End If
    
    Set swSheetMetal = swFeat.GetDefinition
    
    ' Rollback to change default bend radius
    bRet = swSheetMetal.AccessSelections(swModel, Nothing) ': Debug.Assert bRet

    'Change kFactor
    If strBendTable = "-1" Then
        swSheetMetal.BendAllowanceType = swBendAllowanceKFactor
        swSheetMetal.KFactor = KFactor
    Else
        swSheetMetal.BendAllowanceType = swBendAllowanceBendTable
        swSheetMetal.BendTableFile = strBendTable
    End If
    
    ' Apply changes
    bRet = swFeat.ModifyDefinition(swSheetMetal, swModel, Nothing) ': Debug.Assert bRet
    swModel.ForceRebuild3 True
    swModel.ForceRebuild3 True
    
    swModel.Save
    
End Sub

Function ConvertToSheetMetal() As Boolean

Set swApp = Application.SldWorks
swApp.CloseAllDocuments True
Set swModel = swApp.ActiveDoc
Set swFeatMgr = swModel.FeatureManager

boolSheet = swFeatMgr.InsertConvertToSheetMetal(0.001, False, False, 0.001, 0.001, 2, 0.5)
 
End Function

Sub Recover()
Set swApp = Application.SldWorks
swApp.DocumentVisible True, swDocASSEMBLY
swApp.UserControl = True
swApp.Visible = True
swApp.Frame.KeepInvisible = False
End Sub
Sub N210()
   Set swSelMgr = swModel.SelectionManager
   If swSelMgr.GetSelectedObjectCount2(0) = 0 Then Exit Sub
   If swSelMgr.GetSelectedObjectType2(1) <> swSelFACES Then Exit Sub
   Set swFace = swSelMgr.GetSelectedObject2(1)

   ' Get list of edges and edge count for loop.
   Dim dblEdgeCount As Double
   Dim varEdgeList As Variant
   dblEdgeCount = swFace.GetEdgeCount
   varEdgeList = swFace.GetEdges()

   ' Set perimeter/cut length to zero
   Dim dblLength As Double 'total cut length
   dblLength = 0

   ' Looping through all edges on selected face.
   Dim swCurve As Curve
   Dim swEdge As Edge
   Dim intCount As Integer
   Dim curveParams As Variant
   For intCount = 0 To (dblEdgeCount - 1)
       ' Getting the edge object and curve parameters.
       Set swEdge = varEdgeList(intCount)
       Set swCurve = swEdge.GetCurve
       curveParams = swEdge.GetCurveParams2()
       ' Add length of curve to sum
       dblLength = dblLength + swCurve.GetLength2(curveParams(6), curveParams(7))
   Next intCount
   dblLength = dblLength * 39.36996 'convert to inches
   swModel.CustomInfo2("", "F210_R") = Round(dblLength / DeburRate, 3)
   swModel.CustomInfo2("", "F210_S") = 0.03
   swModel.ForceRebuild3 True
End Sub
